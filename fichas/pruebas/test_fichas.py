"""Pruebas del módulo de fichas maestras, solo con empresas inventadas.

    python -m pytest fichas/pruebas -q

Las pruebas de generación necesitan LibreOffice y pdfplumber.
"""
import datetime
import hashlib
import json
import re
import shutil
import sys
import zipfile
from pathlib import Path
from types import SimpleNamespace

import pytest

AQUI = Path(__file__).resolve().parent
sys.path.insert(0, str(AQUI.parent))

import datos_prueba  # noqa: E402
import fichas  # noqa: E402

ORGS = AQUI / 'organizaciones_prueba_2026-10-13.json'
ENTRADAS = AQUI / 'entradas_prueba.json'
EXPEDIENTES = AQUI / 'expedientes_prueba.json'
COMPLETA = 'biometano-meseta-ficticia'
EPC = 'montajes-ejemplo'
CONSULTORA = 'consultora-imaginaria'
VACIA = 'cooperativa-agraria-inventada'
ALTA = 'depositos-imaginarios-del-levante'   # dada de alta fuera de organizaciones.json
FECHA = '2026-10-14'

SIN_LIBREOFFICE = not any(shutil.which(n) for n in ('soffice', 'libreoffice')) and not any(
    Path(r).exists() for r in fichas.RUTAS_LIBREOFFICE)
con_libreoffice = pytest.mark.skipif(SIN_LIBREOFFICE, reason='LibreOffice no está instalado')


# --- Utilidades ----------------------------------------------------------------

def ejecutar(funcion, *args, **kw):
    """(código de salida, texto del informe)."""
    lineas = []
    codigo = funcion(*args, informe=lineas.append, **kw)
    return codigo, '\n'.join(lineas)


def escribir(ruta, datos):
    ruta.write_text(json.dumps(datos, ensure_ascii=False, indent=2), encoding='utf-8')
    return ruta


def huella_archivo(ruta):
    return hashlib.sha256(Path(ruta).read_bytes()).hexdigest()


def expediente_de_prueba(carpeta):
    ruta = carpeta / 'expedientes.json'
    shutil.copyfile(EXPEDIENTES, ruta)
    return ruta


def documento(ruta_docx):
    with zipfile.ZipFile(ruta_docx) as z:
        return z.read('word/document.xml').decode('utf-8')


def texto(ruta_docx):
    return ''.join(re.findall(r'<w:t[^>]*>([^<]*)</w:t>', documento(ruta_docx)))


def leer(ruta):
    return json.loads(Path(ruta).read_text(encoding='utf-8'))


def generar(exp, salida, panel, tracker, **kw):
    kw.setdefault('fecha', FECHA)
    return ejecutar(fichas.generar, exp, salida, ORGS, panel, tracker, **kw)


@pytest.fixture(scope='module')
def generado(tmp_path_factory):
    """Fichas e índice generados una vez para las pruebas que solo los leen."""
    carpeta = tmp_path_factory.mktemp('generado')
    exp = expediente_de_prueba(carpeta)
    panel, tracker = datos_prueba.crear(carpeta)
    salida = carpeta / 'salida'
    codigo, informe = generar(exp, salida, panel, tracker)
    assert codigo == 0, informe
    return SimpleNamespace(exp=exp, salida=salida, panel=panel, tracker=tracker,
                           fichas=salida / fichas.CARPETA_FICHAS)


# --- Esquema -------------------------------------------------------------------

def test_expediente_de_ejemplo_cumple_el_esquema():
    assert fichas.errores_expedientes(leer(EXPEDIENTES)) == []


def test_validador_coincide_con_jsonschema():
    jsonschema = pytest.importorskip('jsonschema')
    esquema = leer(fichas.ESQUEMA)
    jsonschema.Draft202012Validator.check_schema(esquema)
    validador = jsonschema.Draft202012Validator(esquema, format_checker=jsonschema.FormatChecker())
    bueno = leer(EXPEDIENTES)
    assert not list(validador.iter_errors(bueno))
    malo = json.loads(json.dumps(bueno))
    e = malo['empresas'][COMPLETA]
    e['cronologia'][0]['fecha'] = '2026-13-01'
    e['cronologia'][1]['tipo'] = 'fax'
    e['lo_que_sabemos'][0]['naturaleza'] = 'rumor'
    e['tanques'][0]['volumen_m3'] = -5
    del e['pendientes'][0]['estado']
    e['campo_nuevo'] = 1
    malo['empresas'][ALTA]['relaciones'][0]['papel'] = 'amigo'
    errores = fichas.errores_esquema(malo)
    assert len(errores) == 7, errores
    assert len(list(validador.iter_errors(malo))) == 7


def test_reglas_que_el_esquema_no_expresa():
    exp = leer(EXPEDIENTES)
    e = exp['empresas'][COMPLETA]
    e['pendientes'][0]['estado'] = 'cerrado'
    e['cronologia'][1]['fuente'] = {'tipo': 'gmail'}
    e['lo_que_sabemos'][1]['id'] = 'sab-0001'
    exp['empresas']['otra'] = dict(e, id='distinta')
    del exp['empresas'][ALTA]['contactos']
    exp['empresas'][VACIA]['nombre'] = 'Cooperativa Agraria Inventada'
    errores = '\n'.join(fichas.errores_expedientes(exp))
    assert f'empresas.{ALTA}: falta «contactos» (empresa dada de alta fuera de organizaciones.json)' in errores
    assert f'empresas.{VACIA}: «nombre» solo corresponde a las empresas dadas de alta' in errores
    assert 'un pendiente cerrado necesita «fecha_cierre»' in errores
    assert 'necesita «fuente.ref»' in errores
    assert 'id repetido «sab-0001»' in errores
    assert 'el id «distinta» no coincide con la clave' in errores



def test_ejemplos_de_las_instrucciones_son_validos():
    instrucciones = (AQUI.parent / 'INSTRUCCIONES_FICHAS.md').read_text(encoding='utf-8')
    ejemplos = re.findall(r'```json\n(.*?)```', instrucciones, re.S)
    assert len(ejemplos) >= 2
    for ejemplo in ejemplos:
        assert fichas.errores_esquema(json.loads(ejemplo), {'$ref': '#/$defs/entradas'}, 'entradas') == []
    assert fichas.errores_esquema(leer(ENTRADAS), {'$ref': '#/$defs/entradas'}, 'entradas') == []


# --- inicializar ---------------------------------------------------------------

def test_inicializar_dos_veces_no_duplica(tmp_path):
    ruta = tmp_path / 'expedientes.json'
    codigo, informe = ejecutar(fichas.inicializar, ORGS, ruta)
    assert codigo == 0
    assert '5 organizaciones' in informe and '5 expedientes creados' in informe
    exp = leer(ruta)
    assert list(exp['empresas']) == [o['id'] for o in fichas.cargar_organizaciones(ORGS)]
    assert exp['empresas'][VACIA] == fichas.expediente_vacio(VACIA)
    assert exp['empresas'][VACIA]['relaciones'] == []
    antes = huella_archivo(ruta)
    codigo, informe = ejecutar(fichas.inicializar, ORGS, ruta)
    assert codigo == 0
    assert '0 expedientes creados' in informe and '5 ya existían' in informe
    assert huella_archivo(ruta) == antes
    assert len(leer(ruta)['empresas']) == 5
    assert not (tmp_path / fichas.COPIAS).exists()


def test_inicializar_anade_solo_las_que_faltan_sin_sobrescribir(tmp_path):
    exp = leer(EXPEDIENTES)
    del exp['empresas'][VACIA]
    ruta = escribir(tmp_path / 'expedientes.json', exp)
    codigo, informe = ejecutar(fichas.inicializar, ORGS, ruta)
    assert codigo == 0 and '1 expedientes creados' in informe
    nuevo = leer(ruta)
    assert nuevo['empresas'][COMPLETA] == exp['empresas'][COMPLETA]
    assert nuevo['empresas'][VACIA] == fichas.expediente_vacio(VACIA)
    assert len(list((tmp_path / fichas.COPIAS).glob('expedientes_*.json'))) == 1


# --- añadir --------------------------------------------------------------------

def test_flujo_reproduce_el_expediente_de_ejemplo(tmp_path):
    ruta = tmp_path / 'expedientes.json'
    assert ejecutar(fichas.inicializar, ORGS, ruta)[0] == 0
    codigo, informe = ejecutar(fichas.anadir, ruta, ENTRADAS, ruta_org=ORGS)
    assert codigo == 0, informe
    assert leer(ruta) == leer(EXPEDIENTES)


def test_duplicado_por_referencia_de_gmail_se_omite(tmp_path):
    ruta = expediente_de_prueba(tmp_path)
    antes = huella_archivo(ruta)
    entradas = escribir(tmp_path / 'entradas.json', [{
        'empresa': COMPLETA,
        'cronologia': [{'fecha': '2026-10-03', 'tipo': 'correo_recibido', 'resumen': 'Otro resumen del mismo correo.',
                        'origen': 'automatico', 'fuente': {'tipo': 'gmail', 'ref': '19a1b2c3d4e5f601'}}],
    }])
    codigo, informe = ejecutar(fichas.anadir, ruta, entradas)
    assert codigo == 0
    assert 'duplicada de cro-0002 «02/10/2026 · Correo recibido: Envía las dimensiones' in informe
    assert '(misma referencia de Gmail 19a1b2c3d4e5f601).' in informe
    assert 'Sin cambios' in informe
    assert huella_archivo(ruta) == antes


def test_duplicado_por_fecha_tipo_y_resumen_se_omite_y_lo_nuevo_se_anade(tmp_path):
    ruta = expediente_de_prueba(tmp_path)
    entradas = escribir(tmp_path / 'entradas.json', [{
        'empresa': 'Meseta Bio',
        'cronologia': [
            {'fecha': '2026-10-05', 'tipo': 'llamada',
             'resumen': 'LLAMADA de seguimiento: la planta sur va más retrasada',
             'origen': 'manual', 'fuente': {'tipo': 'conversacion'}},
            {'fecha': '2026-10-13', 'tipo': 'whatsapp', 'resumen': 'Confirma que reciben la oferta.',
             'origen': 'manual', 'fuente': {'tipo': 'conversacion'}},
        ],
    }])
    codigo, informe = ejecutar(fichas.anadir, ruta, entradas)
    assert codigo == 0, informe
    assert 'duplicada de cro-0003 «05/10/2026 · Llamada' in informe and '(misma fecha, tipo y resumen).' in informe
    cronologia = leer(ruta)['empresas'][COMPLETA]['cronologia']
    assert [x['id'] for x in cronologia] == ['cro-0001', 'cro-0002', 'cro-0003', 'cro-0004', 'cro-0005']
    assert cronologia[-1]['tipo'] == 'whatsapp'


def test_empresa_inexistente_es_conflicto_y_no_se_aplica_nada(tmp_path):
    ruta = expediente_de_prueba(tmp_path)
    antes = huella_archivo(ruta)
    entradas = escribir(tmp_path / 'entradas.json', [
        {'empresa': EPC, 'competencia': [{'texto': 'Nueva objeción.', 'fuente': 'Correo', 'fecha': '2026-10-13'}]},
        {'empresa': 'Empresa Que No Existe', 'situacion_actual': {'texto': 'Texto.', 'fecha': '2026-10-13'}},
    ])
    codigo, informe = ejecutar(fichas.anadir, ruta, entradas)
    assert codigo == 2
    assert 'CONFLICTO  bloque 2: «Empresa Que No Existe» no se puede asociar a ninguna empresa' in informe
    assert 'No se ha aplicado nada' in informe
    assert huella_archivo(ruta) == antes
    assert not (tmp_path / fichas.COPIAS).exists()


def test_intento_de_modificar_una_entrada_existente_es_conflicto(tmp_path):
    ruta = expediente_de_prueba(tmp_path)
    antes = huella_archivo(ruta)
    exp = leer(ruta)['empresas'][COMPLETA]
    entradas = escribir(tmp_path / 'entradas.json', [{
        'empresa': COMPLETA,
        'cotizaciones': [dict(exp['cotizaciones'][0], estado='Aceptada')],
        'relaciones': [dict(exp['relaciones'][0], texto='Otro texto.')],
    }])
    codigo, informe = ejecutar(fichas.anadir, ruta, entradas)
    assert codigo == 2
    assert 'intento de modificar la entrada cot-0001 (cambia estado)' in informe
    assert 'intento de modificar la entrada rel-0001 (cambia texto)' in informe
    assert informe.count('CONFLICTO') == 2
    assert huella_archivo(ruta) == antes


DUPLICADOS = {  # lista: (duplicado de la primera entrada, con otros datos; entrada que no es duplicada)
    'lo_que_sabemos': ({'texto': 'LA PLANTA NORTE tratará 60.000 t/año de purines y estiércoles',
                        'naturaleza': 'estimacion', 'fuente': 'Otra', 'fecha': '2026-10-13'},
                       {'texto': 'Otro dato.', 'naturaleza': 'dato', 'fuente': 'Otra', 'fecha': '2026-10-13'}),
    'competencia': ({'texto': 'Comparan con tanques de hormigon de un constructor local, piden argumentos de '
                              'plazo de montaje', 'fuente': 'Otra', 'fecha': '2026-10-13'},
                    {'texto': 'Otra objeción.', 'fuente': 'Otra', 'fecha': '2026-10-13'}),
    'proyectos': ({'nombre': 'planta ficticia norte', 'ubicacion': 'BURGOS', 'situacion': 'En obra.',
                   'fuente': 'Otra', 'fecha': '2026-10-13'},
                  {'nombre': 'Planta Ficticia Norte', 'ubicacion': 'Palencia', 'fuente': 'Otra',
                   'fecha': '2026-10-13'}),
    'relaciones': ({'organizacion': 'Fondo inversor ficticio', 'papel': 'accionista', 'texto': 'Mayoritario.',
                    'fuente': 'Otra', 'fecha': '2026-10-13'},
                   {'organizacion': 'Fondo Inversor Ficticio', 'papel': 'socio', 'fuente': 'Otra',
                    'fecha': '2026-10-13'}),
    'tanques': ({'proyecto': 'Planta Ficticia Norte', 'uso': 'Digestor', 'cantidad': 2, 'diametro_m': 24.36,
                 'altura_m': 12.6, 'volumen_m3': 5900, 'fuente': 'Otra', 'fecha': '2026-10-13'},
                {'proyecto': 'Planta Ficticia Norte', 'uso': 'Digestor', 'cantidad': 2, 'diametro_m': 24.36,
                 'altura_m': 14, 'fuente': 'Otra', 'fecha': '2026-10-13'}),
    'cotizaciones': ({'fecha': '2026-10-08', 'descripcion': 'Dos digestores de 5.870 m³ con techo de doble membrana',
                      'estado': 'Aceptada', 'importe_eur': 400000, 'fuente': 'Otra'},
                     {'fecha': '2026-10-13', 'descripcion': 'Dos digestores de 5.870 m³ con techo de doble membrana',
                      'estado': 'Revisada', 'fuente': 'Otra'}),
    'pendientes': ({'accion': 'enviar referencias de digestores con doble membrana', 'de_quien': 'Stora',
                    'estado': 'abierto'},
                   {'accion': 'Enviar planos de implantación revisados', 'de_quien': 'cliente', 'estado': 'abierto'}),
    'documentos': ({'fecha': '2026-10-13', 'nombre': 'Otro nombre',
                    'ruta': 'DOCUMENTOS\\Biometano Meseta Ficticia\\Oferta_prueba_2026-10-08.pdf'},
                   {'fecha': '2026-10-08', 'nombre': 'Oferta de prueba digestores planta norte',
                    'ruta': 'DOCUMENTOS\\Biometano Meseta Ficticia\\Oferta_prueba_2026-10-08_v2.pdf'}),
}


def test_criterios_de_duplicado(tmp_path):
    ruta = expediente_de_prueba(tmp_path)
    antes = leer(ruta)['empresas'][COMPLETA]
    entradas = escribir(tmp_path / 'entradas.json', [
        {'empresa': COMPLETA, **{lista: list(par) for lista, par in DUPLICADOS.items()}}])
    codigo, informe = ejecutar(fichas.anadir, ruta, entradas)
    assert codigo == 0, informe
    despues = leer(ruta)['empresas'][COMPLETA]
    for lista, (duplicado, nuevo) in DUPLICADOS.items():
        assert despues[lista][:len(antes[lista])] == antes[lista], lista   # lo existente no cambia
        assert despues[lista][len(antes[lista]):] == [dict(nuevo, id=despues[lista][-1]['id'])], lista
    assert informe.count('OMITIDA') == len(DUPLICADOS)
    for motivo in ('mismo texto', 'mismo nombre y ubicación', 'misma organización y papel',
                   'mismo proyecto, uso, cantidad, diámetro y altura', 'misma fecha y descripción',
                   'misma acción, que sigue abierta', 'misma ruta'):
        assert f'({motivo}).' in informe, motivo
    assert 'cot-0001 «Dos digestores de 5.870 m³ con techo de doble membrana» (misma fecha y descripción). ' \
           'Difiere en estado, importe_eur; la entrada existente no se modifica.' in informe


def test_pendiente_igual_a_uno_cerrado_se_admite(tmp_path):
    ruta = expediente_de_prueba(tmp_path)
    accion = {'accion': 'Enviar referencias de digestores con doble membrana', 'de_quien': 'nosotros',
              'estado': 'abierto'}
    entradas = escribir(tmp_path / 'entradas.json', [{'empresa': COMPLETA, 'pendientes': [accion]}])
    codigo, informe = ejecutar(fichas.anadir, ruta, entradas)
    assert codigo == 0 and 'duplicada de pen-0001' in informe and 'Sin cambios' in informe
    escribir(entradas, [{'empresa': COMPLETA, 'cerrar_pendientes': [{'id': 'pen-0001', 'fecha_cierre': '2026-10-13'}],
                         'pendientes': [dict(accion, estado='cerrado', fecha_cierre='2026-10-13')]}])
    codigo, informe = ejecutar(fichas.anadir, ruta, entradas)
    assert codigo == 0 and 'Para cerrar el pendiente abierto, use «cerrar_pendientes».' in informe
    escribir(entradas, [{'empresa': COMPLETA, 'pendientes': [accion]}])
    codigo, informe = ejecutar(fichas.anadir, ruta, entradas)
    assert codigo == 0 and 'pendientes: pen-0004' in informe
    pendientes = leer(ruta)['empresas'][COMPLETA]['pendientes']
    assert [(x['id'], x['estado']) for x in pendientes] == [
        ('pen-0001', 'cerrado'), ('pen-0002', 'abierto'), ('pen-0003', 'abierto'), ('pen-0004', 'abierto')]


def test_id_inventado_y_formato_incorrecto_son_errores(tmp_path):
    ruta = expediente_de_prueba(tmp_path)
    antes = huella_archivo(ruta)
    entradas = escribir(tmp_path / 'entradas.json', [{
        'empresa': COMPLETA,
        'competencia': [{'id': 'com-0099', 'texto': 'Otra.', 'fuente': 'Correo', 'fecha': '2026-10-13'}],
    }])
    codigo, informe = ejecutar(fichas.anadir, ruta, entradas)
    assert codigo == 1 and 'el id com-0099 no existe' in informe
    entradas = escribir(tmp_path / 'entradas.json', [{
        'empresa': COMPLETA,
        'cronologia': [{'fecha': '14/10/2026', 'tipo': 'reunion', 'resumen': 'Visita.', 'origen': 'manual',
                        'fuente': {'tipo': 'conversacion'}}],
    }])
    codigo, informe = ejecutar(fichas.anadir, ruta, entradas)
    assert codigo == 1 and 'entradas[0].cronologia[0].fecha: «14/10/2026» no tiene el formato esperado' in informe
    assert huella_archivo(ruta) == antes


def test_cerrar_pendientes(tmp_path):
    ruta = expediente_de_prueba(tmp_path)
    entradas = escribir(tmp_path / 'entradas.json', [{
        'empresa': COMPLETA,
        'cerrar_pendientes': [{'id': 'pen-0001', 'fecha_cierre': '2026-10-13'},
                              {'accion': 'confirmar plazo de fabricación de la chapa', 'fecha_cierre': '2026-10-14'}],
    }])
    codigo, informe = ejecutar(fichas.anadir, ruta, entradas)
    assert codigo == 0, informe
    pendientes = {p['id']: p for p in leer(ruta)['empresas'][COMPLETA]['pendientes']}
    assert pendientes['pen-0001'] == dict(pendientes['pen-0001'], estado='cerrado', fecha_cierre='2026-10-13')
    assert pendientes['pen-0002']['estado'] == 'cerrado' and pendientes['pen-0002']['fecha_cierre'] == '2026-10-14'
    assert pendientes['pen-0003']['estado'] == 'abierto'
    # Repetir el mismo cierre no cambia nada; cerrarlo con otra fecha es conflicto.
    codigo, informe = ejecutar(fichas.anadir, ruta, entradas)
    assert codigo == 0 and 'Sin cambios' in informe
    escribir(entradas, [{'empresa': COMPLETA, 'cerrar_pendientes': [{'id': 'pen-0001', 'fecha_cierre': '2026-10-20'}]}])
    codigo, informe = ejecutar(fichas.anadir, ruta, entradas)
    assert codigo == 2 and 'pen-0001 ya está cerrado desde el 13/10/2026' in informe


def test_situacion_actual_y_nombres_alternativos(tmp_path):
    ruta = expediente_de_prueba(tmp_path)
    entradas = escribir(tmp_path / 'entradas.json', [{
        'empresa': EPC,
        'nombres_alternativos': ['Montajes Ejemplo Ibérica'],
        'situacion_actual': {'texto': 'Esperan la tarifa orientativa.', 'fecha': '2026-10-13'},
    }])
    assert ejecutar(fichas.anadir, ruta, entradas)[0] == 0
    e = leer(ruta)['empresas'][EPC]
    assert e['nombres_alternativos'] == ['Grupo Montajes Ejemplo', 'Montajes Ejemplo Ibérica']
    assert e['situacion_actual'] == {'texto': 'Esperan la tarifa orientativa.', 'fecha': '2026-10-13'}
    # Una situación más antigua que la registrada y un nombre de otra empresa son conflictos.
    escribir(entradas, [{'empresa': EPC, 'nombres_alternativos': ['Meseta Bio'],
                         'situacion_actual': {'texto': 'Texto anterior.', 'fecha': '2026-10-01'}}])
    codigo, informe = ejecutar(fichas.anadir, ruta, entradas)
    assert codigo == 2
    assert '«Meseta Bio» ya identifica a «biometano-meseta-ficticia»' in informe
    assert 'la situación actual registrada (13/10/2026) es posterior a la propuesta (01/10/2026)' in informe


# --- Empresas fuera de organizaciones.json ---------------------------------------

def alta(nombre, tipo='B', **bloque):
    return {'alta': {'nombre': nombre, 'tipo': tipo}, **bloque}


def test_alta_crea_el_expediente(tmp_path):
    e = leer(EXPEDIENTES)['empresas'][ALTA]
    assert ALTA == fichas.crear_id('Depósitos Imaginarios del Levante')
    assert {k: e[k] for k in ('id', 'nombre', 'tipo', 'fuera_de_organizaciones')} == {
        'id': ALTA, 'nombre': 'Depósitos Imaginarios del Levante', 'tipo': 'B', 'fuera_de_organizaciones': True}
    assert [c['nombre'] for c in e['contactos']] == ['Elena Ficticia Mora', 'Ofertas (Dpto.)']
    assert e['nombres_alternativos'] == ['DIL Tanques'] and len(e['relaciones']) == 2
    # Repetir el alta no es un conflicto: se omite y el resto del bloque se aplica a la empresa.
    ruta = expediente_de_prueba(tmp_path)
    entradas = escribir(tmp_path / 'entradas.json', [alta('Depósitos Imaginarios del Levante', competencia=[
        {'texto': 'Comparan con depósitos de poliéster.', 'fuente': 'Correo', 'fecha': '2026-10-13'}])])
    codigo, informe = ejecutar(fichas.anadir, ruta, entradas, ruta_org=ORGS)
    assert codigo == 0, informe
    assert f'ya estaba dada de alta como «{ALTA}»' in informe
    assert leer(ruta)['empresas'][ALTA]['competencia'][0]['id'] == 'com-0001'


@pytest.mark.parametrize('bloque, mensaje', [
    (alta('Montajes Ejemplo'),
     'el id «montajes-ejemplo» o el nombre ya corresponden a «montajes-ejemplo» (organizaciones.json)'),
    (alta('MONTAJES EJEMPLO, S.L.'), 'ya corresponden a «montajes-ejemplo» (organizaciones.json)'),
    (alta('Ingeniería Modelo Norte'), 'ya corresponden a «ingenieria-modelo-norte» (organizaciones.json)'),
    (alta('Meseta Bio'), 'ya corresponden a «biometano-meseta-ficticia» (organizaciones.json)'),
    (alta('Depositos Imaginarios del Levante', tipo='C'), f'ya corresponden a «{ALTA}» (expediente)'),
    (alta('Tanques Nuevos Inventados', nombres_alternativos=['Grupo Montajes Ejemplo']),
     'el nombre alternativo «Grupo Montajes Ejemplo» ya identifica a «montajes-ejemplo»'),
    (alta('Modelo Norte'), 'el nombre se parece al de «ingenieria-modelo-norte» (organizaciones.json)'),
    (alta('Depósitos Imaginarios del Levante Sur'),
     f'el nombre se parece al de «{ALTA}» (expediente). Si es la misma empresa, use un bloque «empresa»'),
])
def test_alta_que_coincide_con_otra_empresa_es_conflicto(tmp_path, bloque, mensaje):
    ruta = expediente_de_prueba(tmp_path)
    antes = huella_archivo(ruta)
    entradas = escribir(tmp_path / 'entradas.json', [bloque])
    codigo, informe = ejecutar(fichas.anadir, ruta, entradas, ruta_org=ORGS)
    assert codigo == 2, informe
    assert mensaje in informe
    assert huella_archivo(ruta) == antes


def test_alta_con_nombre_parecido_confirmada_por_andres(tmp_path):
    ruta = expediente_de_prueba(tmp_path)
    entradas = escribir(tmp_path / 'entradas.json', [dict(alta('Modelo Norte', tipo='C'), no_es_duplicado=True)])
    codigo, informe = ejecutar(fichas.anadir, ruta, entradas, ruta_org=ORGS)
    assert codigo == 0, informe
    assert 'Llevan «no_es_duplicado» (confirmado por Andrés): bloque 1 (alta «Modelo Norte»).' in informe
    assert leer(ruta)['empresas']['modelo-norte']['fuera_de_organizaciones'] is True


def test_alta_sin_organizaciones_o_con_empresa_es_error(tmp_path):
    ruta = expediente_de_prueba(tmp_path)
    antes = huella_archivo(ruta)
    entradas = escribir(tmp_path / 'entradas.json', [alta('Tanques Nuevos Inventados')])
    codigo, informe = ejecutar(fichas.anadir, ruta, entradas)
    assert codigo == 1 and 'el alta necesita --organizaciones' in informe
    escribir(entradas, [dict(alta('Tanques Nuevos Inventados'), empresa=COMPLETA)])
    codigo, informe = ejecutar(fichas.anadir, ruta, entradas, ruta_org=ORGS)
    assert codigo == 1 and 'debe llevar «empresa» (empresa con expediente) o «alta»' in informe
    assert huella_archivo(ruta) == antes


def test_contactos_de_las_altas(tmp_path):
    ruta = expediente_de_prueba(tmp_path)
    existentes = leer(ruta)['empresas'][ALTA]['contactos']
    entradas = escribir(tmp_path / 'entradas.json', [{'empresa': 'DIL Tanques', 'contactos': [
        {'nombre': 'Elena F. Mora', 'cargo': 'Gerente', 'correo': 'ELENA.FICTICIA@depositos-imaginarios.example'},
        {'nombre': 'elena ficticia mora', 'telefono': '+34 600 000 699'},
        {'nombre': 'Jorge Supuesto Vidal', 'cargo': 'Compras',
         'correo': 'jorge.supuesto@depositos-imaginarios.example'},
    ]}])
    codigo, informe = ejecutar(fichas.anadir, ruta, entradas, ruta_org=ORGS)
    assert codigo == 0, informe
    assert ('«Elena F. Mora» duplicado de «Elena Ficticia Mora» (mismo correo). Difiere en cargo, nombre, '
            'telefono; el contacto existente no se modifica.') in informe
    assert '«elena ficticia mora» duplicado de «Elena Ficticia Mora» (mismo nombre)' in informe
    contactos = leer(ruta)['empresas'][ALTA]['contactos']
    assert contactos[:2] == existentes                       # nunca se modifican ni se borran
    assert contactos[2] == {'nombre': 'Jorge Supuesto Vidal', 'cargo': 'Compras',
                            'correo': 'jorge.supuesto@depositos-imaginarios.example'}
    # Los contactos de las empresas de organizaciones.json no van al expediente.
    escribir(entradas, [{'empresa': COMPLETA, 'contactos': [{'nombre': 'Persona Nueva'}]}])
    codigo, informe = ejecutar(fichas.anadir, ruta, entradas, ruta_org=ORGS)
    assert codigo == 1 and 'los contactos de las empresas de organizaciones.json se mantienen allí' in informe


def test_validar_avisa_si_un_alta_aparece_en_organizaciones(tmp_path):
    exp = expediente_de_prueba(tmp_path)
    datos = leer(ORGS)
    nueva = dict(datos['organizaciones'][-1], id='depositos-levante', nombre='Depósitos Imaginarios del Levante, S.L.')
    datos['organizaciones'].append(nueva)
    orgs = escribir(tmp_path / 'organizaciones.json', datos)
    codigo, informe = ejecutar(fichas.validar, exp, orgs)
    assert codigo == 0
    assert (f'«Depósitos Imaginarios del Levante» ({ALTA}), dada de alta fuera de organizaciones.json, aparece '
            'ahora en organizaciones.json como «Depósitos Imaginarios del Levante, S.L.» (depositos-levante): '
            'hay que migrar su expediente.') in informe
    assert 'depositos-levante» (Depósitos Imaginarios del Levante, S.L.) no tiene expediente' not in informe
    # inicializar no crea un segundo expediente para la misma empresa.
    codigo, informe = ejecutar(fichas.inicializar, orgs, exp)
    assert codigo == 0 and 'hay que migrar su expediente' in informe
    assert '0 expedientes creados · 5 ya existían · 1 pendientes de migrar · 6 en expedientes.json' in informe
    assert 'depositos-levante' not in leer(exp)['empresas']


def test_expediente_anterior_sin_relaciones_se_completa(tmp_path):
    exp = leer(EXPEDIENTES)
    for e in exp['empresas'].values():
        del e['relaciones']
    ruta = escribir(tmp_path / 'expedientes.json', exp)
    cargado, errores = fichas.cargar_expedientes(ruta)
    assert errores == [] and cargado['empresas'][VACIA]['relaciones'] == []


def test_copia_de_seguridad_antes_de_escribir(tmp_path):
    ruta = expediente_de_prueba(tmp_path)
    original = ruta.read_bytes()
    entradas = escribir(tmp_path / 'entradas.json', [{
        'empresa': VACIA,
        'cronologia': [{'fecha': '2026-10-13', 'tipo': 'llamada', 'resumen': 'Primera llamada.', 'origen': 'manual',
                        'fuente': {'tipo': 'conversacion'}}],
    }])
    codigo, informe = ejecutar(fichas.anadir, ruta, entradas, simular=True)
    assert codigo == 0 and 'Simulación' in informe
    assert ruta.read_bytes() == original and not (tmp_path / fichas.COPIAS).exists()
    codigo, informe = ejecutar(fichas.anadir, ruta, entradas)
    assert codigo == 0
    copias = list((tmp_path / fichas.COPIAS).glob('expedientes_*.json'))
    assert len(copias) == 1
    assert re.fullmatch(r'expedientes_\d{4}-\d{2}-\d{2}_\d{6}\.json', copias[0].name)
    assert copias[0].read_bytes() == original
    assert leer(ruta)['empresas'][VACIA]['cronologia'][0]['id'] == 'cro-0001'


# --- Panel, tracker y asociación de nombres ------------------------------------

def test_panel_y_tracker_con_cabecera_localizada_por_su_texto(tmp_path):
    panel, tracker = datos_prueba.crear(tmp_path)
    filas, avisos = fichas.leer_panel(panel)
    assert avisos == []
    assert {f['hoja'] for f in filas} == set(datos_prueba.FILAS_PANEL)
    meseta = next(f for f in filas if f['nombre'] == 'Biometano Meseta Ficticia')
    assert meseta['fila'] == 4 and meseta['estado'] == 'Oferta enviada'
    assert fichas.fecha_es(meseta['fecha']) == '20/10/2026'
    filas, _ = fichas.leer_tracker(tracker)
    assert [f['rating'] for f in filas] == [5, 4, 2, None, 3]


def test_asociacion_de_nombres():
    asociador = fichas.Asociador(fichas.cargar_organizaciones(ORGS), leer(EXPEDIENTES)['empresas'])
    casos = {
        'consultora-imaginaria': CONSULTORA,                          # id
        'Montajes Ejemplo, S.L.': EPC,                                # forma jurídica
        'INGENIERIA MODELO NORTE': 'ingenieria-modelo-norte',         # mayúsculas, tildes y paréntesis
        'Ingeniería Modelo Norte grupo Modelo': 'ingenieria-modelo-norte',
        'Meseta Bio': COMPLETA,                                       # nombre alternativo
        'Grupo Montajes Ejemplo S.A.': EPC,
        'Depositos Imaginarios del Levante S.L.': ALTA,               # nombre de un alta
        'DIL Tanques': ALTA,
    }
    for nombre, oid in casos.items():
        assert asociador.buscar(nombre) == (oid, ''), nombre
    assert asociador.buscar('Empresa Desconocida')[0] is None


def test_valoracion():
    assert [fichas.rango_valoracion(v) for v in (5, 1.0, '3', None, '', '★★★', 'Hot', 'B', 'otra')] == \
        [5, 1, 3, None, None, 3, 3, 2, None]
    assert [fichas.en_escala(v) for v in (1, 5, 3.0, '4', 0, 6, 4.5, '★★★', True)] == \
        [True, True, True, True, False, False, False, False, False]
    assert fichas.texto_valoracion(4.0) == '4' and fichas.texto_valoracion(None) == ''


def test_valoracion_fuera_de_escala_es_aviso(tmp_path):
    exp = leer(EXPEDIENTES)
    filas = [{'nombre': 'Meseta Bio', 'rating': '★★★', 'hoja': 'Lead Overview', 'fila': 3},
             {'nombre': 'Consultora Imaginaria', 'rating': 7, 'hoja': 'Lead Overview', 'fila': 4}]
    datos, _, avisos = fichas.cruzar(exp, fichas.cargar_organizaciones(ORGS), filas_tracker=filas)
    assert 'Tracker, hoja «Lead Overview», fila 3: valoración «★★★» fuera de la escala de 1 a 5.' in avisos
    assert 'Tracker, hoja «Lead Overview», fila 4: valoración «7» fuera de la escala de 1 a 5.' in avisos
    assert {d['id']: d['valoracion'] for d in datos}[COMPLETA] == '★★★'


def test_validar_lista_los_nombres_sin_correspondencia(tmp_path):
    exp = expediente_de_prueba(tmp_path)
    panel, tracker = datos_prueba.crear(tmp_path)
    codigo, informe = ejecutar(fichas.validar, exp, ORGS, panel, tracker)
    assert codigo == 0
    assert (f'AVISO      Panel, hoja «PROMOTORES», fila 5: «{datos_prueba.SIN_ASOCIAR_PANEL}» '
            'no se puede asociar a ninguna empresa') in informe
    assert f'Tracker, hoja «Lead Overview», fila 7: «{datos_prueba.SIN_ASOCIAR_TRACKER}»' in informe
    assert informe.count('no se puede asociar') == 2


def test_validar_id_sin_correspondencia_es_error(tmp_path):
    exp = leer(EXPEDIENTES)
    exp['empresas']['empresa-retirada'] = fichas.expediente_vacio('empresa-retirada')
    ruta = escribir(tmp_path / 'expedientes.json', exp)
    codigo, informe = ejecutar(fichas.validar, ruta, ORGS)
    assert codigo == 1
    assert 'ERROR      expediente «empresa-retirada» sin correspondencia' in informe


# --- generar -------------------------------------------------------------------

@con_libreoffice
def test_genera_fichas_e_indice_con_nombre_fijo(generado):
    oids = list(leer(EXPEDIENTES)['empresas'])
    assert sorted(p.name for p in generado.fichas.iterdir()) == sorted(
        f'Ficha_{oid}.{x}' for oid in oids for x in ('docx', 'pdf'))
    assert (generado.salida / 'Indice_Fichas_Maestras.docx').exists()
    assert (generado.salida / 'Indice_Fichas_Maestras.pdf').exists()
    estado = leer(generado.salida / fichas.ESTADO)
    assert set(estado['fichas']) == set(oids)


@con_libreoffice
def test_solo_se_regeneran_las_que_cambian(tmp_path):
    exp = expediente_de_prueba(tmp_path)
    panel, tracker = datos_prueba.crear(tmp_path)
    salida = tmp_path / 'salida'
    assert generar(exp, salida, panel, tracker)[0] == 0
    carpeta = salida / fichas.CARPETA_FICHAS
    estado = leer(salida / fichas.ESTADO)
    huellas = {p.name: huella_archivo(p) for p in carpeta.iterdir()}

    codigo, informe = generar(exp, salida, panel, tracker)
    assert codigo == 0 and 'Fichas: 0 regenerada(s), 6 sin cambios' in informe
    assert {p.name: huella_archivo(p) for p in carpeta.iterdir()} == huellas

    entradas = escribir(tmp_path / 'entradas.json', [{
        'empresa': EPC,
        'competencia': [{'texto': 'Piden comparar con un depósito de poliéster.', 'fuente': 'Correo',
                         'fecha': '2026-10-13'}],
    }])
    assert ejecutar(fichas.anadir, exp, entradas)[0] == 0
    codigo, informe = generar(exp, salida, panel, tracker)
    assert codigo == 0 and 'Fichas: 1 regenerada(s), 5 sin cambios' in informe and f'Regeneradas: {EPC}.' in informe
    cambiadas = {n for n, h in huellas.items() if huella_archivo(carpeta / n) != h}
    assert cambiadas == {f'Ficha_{EPC}.docx', f'Ficha_{EPC}.pdf'}
    nuevo = leer(salida / fichas.ESTADO)
    assert nuevo['fichas'][EPC]['huella'] != estado['fichas'][EPC]['huella']
    assert {k: v for k, v in nuevo['fichas'].items() if k != EPC} == \
        {k: v for k, v in estado['fichas'].items() if k != EPC}

    (carpeta / f'Ficha_{VACIA}.pdf').unlink()
    codigo, informe = generar(exp, salida, panel, tracker)
    assert f'Regeneradas: {VACIA}.' in informe
    codigo, informe = generar(exp, salida, panel, tracker, todas=True)
    assert codigo == 0 and 'Fichas: 6 regenerada(s), 0 sin cambios' in informe


def test_fecha_vencida_es_anterior_al_dia_de_generacion():
    casos = [(datetime.datetime(2026, 10, 14, 18, 0), True), (datetime.date(2026, 10, 14), True),
             ('2026-10-14', True), (' 2026-10-14 ', True),
             (datetime.datetime(2026, 10, 15), False),             # el mismo día no está vencida
             (datetime.date(2026, 10, 16), False), ('2026-10-20', False),
             (None, False), ('', False), ('la semana próxima', False)]
    for valor, esperado in casos:
        assert fichas.fecha_vencida(valor, '2026-10-15') is esperado, valor
    assert fichas.fecha_vencida(datetime.date(2026, 10, 14), datetime.date(2026, 10, 15)) is True


def test_la_fecha_vencida_se_compara_con_hoy_si_no_se_indica_otro_dia():
    hoy = datetime.date.today()
    filas = [{'nombre': nombre, 'estado': 'Seguimiento', 'fecha': fecha, 'hoja': 'EPC', 'fila': n}
             for n, (nombre, fecha) in enumerate((
                 ('Montajes Ejemplo', hoy - datetime.timedelta(days=1)),
                 ('Consultora Imaginaria', hoy),
                 ('Biometano Meseta Ficticia', hoy + datetime.timedelta(days=1)),
                 ('Cooperativa Agraria Inventada', None)), 4)]
    datos, _, _ = fichas.cruzar(leer(EXPEDIENTES), fichas.cargar_organizaciones(ORGS), filas)
    assert {d['id']: d['vencido'] for d in datos if d['id'] in (EPC, CONSULTORA, COMPLETA, VACIA)} == {
        EPC: True, CONSULTORA: False, COMPLETA: False, VACIA: False}


@con_libreoffice
def test_fecha_vencida_en_la_ficha_y_en_el_indice(generado):
    alta = texto(generado.fichas / f'Ficha_{ALTA}.docx')                    # fecha del panel: 12/10
    assert 'FECHA VENCIDA12/10/2026' in alta and 'PRÓXIMO COMPROMISO' not in alta
    meseta = texto(generado.fichas / f'Ficha_{COMPLETA}.docx')              # fecha del panel: 20/10
    assert 'PRÓXIMO COMPROMISO20/10/2026' in meseta and 'FECHA VENCIDA' not in meseta
    indice = texto(generado.salida / 'Indice_Fichas_Maestras.docx')
    assert 'Vencida: 12/10/2026' in indice and 'Próximo: 20/10/2026' in indice
    assert 'Próximo: 12/10/2026' not in indice and 'Vencida: 20/10/2026' not in indice


@con_libreoffice
def test_la_ficha_se_regenera_el_dia_en_que_la_fecha_vence(tmp_path):
    exp = expediente_de_prueba(tmp_path)
    panel, tracker = datos_prueba.crear(tmp_path)
    salida = tmp_path / 'salida'
    carpeta = salida / fichas.CARPETA_FICHAS
    assert generar(exp, salida, panel, tracker, fecha='2026-10-14')[0] == 0
    assert 'PRÓXIMO COMPROMISO15/10/2026' in texto(carpeta / f'Ficha_{EPC}.docx')
    # El mismo día de la fecha no está vencida: no hay nada que regenerar.
    codigo, informe = generar(exp, salida, panel, tracker, fecha='2026-10-15')
    assert codigo == 0 and 'Fichas: 0 regenerada(s), 6 sin cambios' in informe
    # Al día siguiente vence la fecha del 15/10 y solo cambia esa ficha.
    codigo, informe = generar(exp, salida, panel, tracker, fecha='2026-10-16')
    assert codigo == 0 and 'Fichas: 1 regenerada(s), 5 sin cambios' in informe
    assert f'Regeneradas: {EPC}.' in informe
    t = texto(carpeta / f'Ficha_{EPC}.docx')
    assert 'FECHA VENCIDA15/10/2026' in t and 'PRÓXIMO COMPROMISO' not in t and 'FICHA GENERADA16/10/2026' in t
    # Mientras siga vencida no se vuelve a regenerar.
    codigo, informe = generar(exp, salida, panel, tracker, fecha='2026-10-17')
    assert 'Fichas: 0 regenerada(s), 6 sin cambios' in informe


@con_libreoffice
def test_no_escribe_en_el_panel_ni_en_el_tracker(generado, tmp_path):
    antes = huella_archivo(generado.panel), huella_archivo(generado.tracker)
    assert generar(generado.exp, tmp_path / 'salida', generado.panel, generado.tracker)[0] == 0
    assert (huella_archivo(generado.panel), huella_archivo(generado.tracker)) == antes


SECCIONES = ('SITUACIÓN ACTUAL', 'CONTACTOS', 'PERFIL Y PROYECTOS', 'LO QUE SABEMOS', 'TANQUES', 'COTIZACIONES',
             'PENDIENTES ABIERTOS', 'COMPETENCIA Y OBJECIONES', 'NOTAS DE REUNIÓN')
# La cronología y los documentos se guardan en el expediente, pero la ficha no los imprime.
RETIRADO = ('CRONOLOG', 'DOCUMENTOS', 'REFERENCIAS DE LA', 'AUTOMÁTICA')


def secciones(ruta_docx):
    """Títulos de sección del documento, en orden."""
    xml = documento(ruta_docx)
    return [t for t in re.findall(r'<w:t[^>]*>([^<]*)</w:t>', xml) if t in SECCIONES]


@con_libreoffice
def test_ficha_completa_con_todas_las_secciones_en_orden(generado):
    ruta = generado.fichas / f'Ficha_{COMPLETA}.docx'
    assert secciones(ruta) == list(SECCIONES)
    t = texto(ruta)
    for esperado in ('FICHA MAESTRA', 'Biometano Meseta Ficticia', 'Grupo A · Promotora', 'VALORACIÓN5',
                     'CON QUIÉN TRABAJA', 'Fondo Inversor Ficticio', ' — Accionista', 'Oferta enviada',
                     'PRÓXIMO COMPROMISO20/10/2026', 'FICHA GENERADA14/10/2026',
                     'Copia de organizaciones.json del 13/10/2026', 'Interlocutor', '(estimación)', '(supuesto)',
                     '412.500 €', '5.870', '24,36'):
        assert esperado in t, esperado
    assert 'FECHA VENCIDA' not in t


@con_libreoffice
def test_la_ficha_no_contiene_cronologia_ni_documentos(generado):
    for docx in generado.fichas.glob('*.docx'):
        t = texto(docx)
        for retirado in RETIRADO:
            assert retirado not in t.upper(), (docx.name, retirado)
    # Sus datos siguen en el expediente y no se pierden: ni su texto ni su referencia de Gmail se imprimen.
    t = texto(generado.fichas / f'Ficha_{COMPLETA}.docx')
    for dato in ('Envía las dimensiones de los digestores', '19a1b2c3d4e5f601', 'Oferta de prueba digestores',
                 'Reunión en el stand', 'Foto de la libreta'):
        assert dato not in t, dato
    cronologia = leer(generado.exp)['empresas'][COMPLETA]['cronologia']
    assert len(cronologia) == 4 and leer(generado.exp)['empresas'][COMPLETA]['documentos']
    # El índice sigue mostrando la última interacción.
    assert '08/10/2026Correo enviado' in texto(generado.salida / 'Indice_Fichas_Maestras.docx')


def entradas_de_saber(*pares):
    return {'lo_que_sabemos': [{'id': f'sab-{n:04d}', 'texto': texto_, 'naturaleza': naturaleza,
                                'fuente': 'Fuente que no se imprime', 'fecha': fecha}
                               for n, (texto_, naturaleza, fecha) in enumerate(pares, 1)]}


def test_lo_que_sabemos_es_una_lista_de_vinetas_ordenada():
    e = entradas_de_saber(
        ('Supuesto antiguo.', 'supuesto', '2026-09-01'), ('Dato antiguo.', 'dato', '2026-09-02'),
        ('Estimación reciente.', 'estimacion', '2026-10-09'), ('Dato reciente.\n  con salto', 'dato', '2026-10-08'),
        ('Supuesto reciente.', 'supuesto', '2026-10-10'), ('Estimación antigua.', 'estimacion', '2026-09-03'))
    xml = fichas.seccion_lo_que_sabemos(e)
    assert '<w:tbl>' not in xml
    lineas = [''.join(re.findall(r'<w:t[^>]*>([^<]*)</w:t>', parrafo)) for parrafo in xml.split('</w:p>')[1:-1]]
    assert lineas == ['•Dato reciente. con salto', '•Dato antiguo.',
                      '•Estimación reciente. (estimación)', '•Estimación antigua. (estimación)',
                      '•Supuesto reciente. (supuesto)', '•Supuesto antiguo. (supuesto)']
    assert xml.count('w:color w:val="595959"') == 4          # solo las marcas, en gris
    for retirado in ('DATO', 'ESTIMACIÓN', 'SUPUESTO', 'Fuente que no se imprime', '09/2026', '10/2026'):
        assert retirado not in xml
    assert fichas.seccion_lo_que_sabemos({'lo_que_sabemos': []}) == ''


@con_libreoffice
def test_lo_que_sabemos_de_la_ficha_no_usa_tabla(generado):
    xml = documento(generado.fichas / f'Ficha_{COMPLETA}.docx')
    inicio = xml.rindex('<w:p>', 0, xml.index('LO QUE SABEMOS'))     # desde el párrafo del título
    seccion = xml[inicio:xml.index('TANQUES', inicio)]
    assert '<w:tbl>' not in seccion and '<w:tc>' not in seccion
    t = ''.join(re.findall(r'<w:t[^>]*>([^<]*)</w:t>', seccion))
    assert t == ('LO QUE SABEMOS•La planta norte tratará 60.000 t/año de purines y estiércoles.'
                 '•Presupuesto de tanques en torno a un millón de euros para las dos plantas. (estimación)'
                 '•La decisión final depende del fondo inversor del grupo. (supuesto)')


@con_libreoffice
def test_las_secciones_vacias_se_omiten(generado):
    ruta = generado.fichas / f'Ficha_{CONSULTORA}.docx'
    assert secciones(ruta) == ['SITUACIÓN ACTUAL', 'CONTACTOS', 'PERFIL Y PROYECTOS', 'NOTAS DE REUNIÓN']
    ruta = generado.fichas / f'Ficha_{EPC}.docx'
    assert secciones(ruta) == ['CONTACTOS', 'PERFIL Y PROYECTOS', 'PENDIENTES ABIERTOS', 'NOTAS DE REUNIÓN']
    for docx in generado.fichas.glob('*.docx'):
        t = texto(docx).casefold()
        for marcador in ('sin datos', 'n/d', 'no localizad', 'no encontrad', 'sin verificar', 'none'):
            assert marcador not in t, (docx.name, marcador)


@con_libreoffice
def test_ficha_de_una_empresa_dada_de_alta(generado):
    docx = generado.fichas / f'Ficha_{ALTA}.docx'
    assert secciones(docx) == ['SITUACIÓN ACTUAL', 'CONTACTOS', 'PERFIL Y PROYECTOS', 'TANQUES',
                               'PENDIENTES ABIERTOS', 'NOTAS DE REUNIÓN']
    t = texto(docx)
    for esperado in ('Depósitos Imaginarios del Levante', 'Grupo B · EPC', 'Empresa fuera de organizaciones.json',
                     'Elena Ficticia Mora', 'Directora Técnica', 'elena.ficticia@depositos-imaginarios.example',
                     '+34 600 000 601', 'Ofertas (Dpto.)', 'CON QUIÉN TRABAJA', 'Montajes Ejemplo', ' — Socio',
                     'Primer contacto', 'FECHA VENCIDA12/10/2026', '18/10/2026'):
        assert esperado in t, esperado
    # Sin perfil ni uso de GFS (dependen de organizaciones.json) y sin valoración (lead sin valorar).
    rotulos = re.findall(r'<w:t[^>]*>([^<]*)</w:t>', documento(docx))
    for ausente in ('PERFIL', 'PROYECTOS', 'USO DE GFS', 'VALORACIÓN', 'Copia de organizaciones.json'):
        assert not any(r.startswith(ausente) and r != 'PERFIL Y PROYECTOS' for r in rotulos), ausente
    assert fichas.problemas_paginacion(docx.with_suffix('.pdf'))[1] == []


@con_libreoffice
def test_ficha_sin_interacciones(generado):
    docx = generado.fichas / f'Ficha_{VACIA}.docx'
    assert secciones(docx) == ['CONTACTOS', 'PERFIL Y PROYECTOS', 'NOTAS DE REUNIÓN']
    t = texto(docx)
    for esperado in ('Cooperativa Agraria Inventada', 'Grupo E · Cooperativa', 'Pablo Inventado López', 'Gerente',
                     'pablo.inventado@cooperativa-inventada.example', '+34 600 000 501',
                     'Cooperativa inventada con interés en el tratamiento de purines.', 'Sin contacto'):
        assert esperado in t, esperado
    assert 'VALORACIÓN' not in t and 'PRÓXIMO COMPROMISO' not in t
    total, problemas = fichas.problemas_paginacion(generado.fichas / f'Ficha_{VACIA}.pdf')
    assert (total, problemas) == (2, [])


@con_libreoffice
def test_pagina_de_notas_en_pagina_propia(generado):
    for docx in generado.fichas.glob('*.docx'):
        xml = documento(docx)
        notas = xml.index('NOTAS DE REUNIÓN')
        assert xml.rfind('<w:pageBreakBefore/>', 0, notas) > xml.rfind('</w:tbl>', 0, notas)
        t = texto(docx)
        for titulo in ('FECHA Y HORA', 'LUGAR O MEDIO', 'ASISTENTES DEL CLIENTE', 'ASISTENTES POR NUESTRA PARTE',
                       'NOTAS', 'ACUERDOS', 'PRÓXIMOS PASOS'):
            assert titulo in t
        total, inicios, notas_pdf, _ = fichas.analizar_pdf(docx.with_suffix('.pdf'))
        assert inicios == [0] and notas_pdf == [total - 1] and total >= 2


@con_libreoffice
def test_cada_ficha_empieza_en_pagina_nueva_y_no_se_desborda(generado):
    for pdf in generado.fichas.glob('*.pdf'):
        assert fichas.problemas_paginacion(pdf)[1] == [], pdf.name
    codigo, informe = ejecutar(fichas.validar, generado.exp, ORGS, salida=generado.salida)
    assert codigo == 0 and '6 fichas generadas comprobadas · 0 errores' in informe


@con_libreoffice
def test_una_pagina_de_notas_desbordada_se_detecta(generado, tmp_path, monkeypatch):
    monkeypatch.setattr(fichas, 'RENGLONES_NOTAS', ((('NOTAS',), 40),))
    salida = tmp_path / 'salida'
    codigo, informe = generar(generado.exp, salida, generado.panel, generado.tracker)
    assert codigo == 1
    assert f'Ficha_{COMPLETA}.pdf: la página de notas de la ficha 1 ocupa 2 páginas' in informe
    assert not list((salida / fichas.CARPETA_FICHAS).glob('Ficha_*'))
    assert leer(salida / fichas.ESTADO)['fichas'] == {}


@con_libreoffice
def test_textos_de_los_datos_no_se_confunden_con_cabeceras(tmp_path):
    exp = leer(EXPEDIENTES)
    e = exp['empresas'][COMPLETA]
    e['competencia'] = [{'id': f'com-{n:04d}', 'fecha': '2026-10-01', 'fuente': 'Reunión',
                         'texto': 'NOTAS DE REUNIÓN pasadas a la FICHA MAESTRA.\n' * 3} for n in range(1, 41)]
    datos, _, _ = fichas.cruzar(exp, fichas.cargar_organizaciones(ORGS))
    d = next(x for x in datos if x['id'] == COMPLETA)
    docx = fichas.escribir_ficha(d, tmp_path / f'Ficha_{COMPLETA}.docx', FECHA, '2026-10-13')
    pdf, = fichas.convertir_a_pdf([docx], tmp_path)
    total, inicios, notas, _ = fichas.analizar_pdf(pdf)
    assert total > 4 and inicios == [0] and notas == [total - 1]
    assert fichas.problemas_paginacion(pdf)[1] == []


@con_libreoffice
def test_sin_tablas_anidadas(generado):
    for docx in [*generado.fichas.glob('*.docx'), generado.salida / 'Indice_Fichas_Maestras.docx']:
        profundidad = 0
        for etiqueta in re.findall(r'<(/?)w:(tbl|tc)\b', documento(docx)):
            if etiqueta == ('', 'tbl'):
                assert profundidad == 0, f'{docx.name}: tabla dentro de una celda'
            profundidad += {('', 'tc'): 1, ('/', 'tc'): -1}.get(etiqueta, 0)


@con_libreoffice
def test_los_docx_se_abren_con_python_docx(generado):
    docx = pytest.importorskip('docx')
    d = docx.Document(str(generado.fichas / f'Ficha_{COMPLETA}.docx'))
    assert len(d.tables) > 10
    assert d.sections[0].footer.paragraphs[0].text.startswith(
        'Stora · Ficha maestra · Biometano Meseta Ficticia · Uso interno · Pág.')
    docx.Document(str(generado.salida / 'Indice_Fichas_Maestras.docx'))


@con_libreoffice
def test_indice_ordenado_por_valoracion_y_ultima_interaccion(generado):
    docx = generado.salida / 'Indice_Fichas_Maestras.docx'
    t = texto(docx)
    for titulo in ('EMPRESA', 'TIPO', 'VALORACIÓN', 'ESTADO', 'ÚLTIMA INTERACCIÓN', 'PENDIENTES'):
        assert titulo in t
    # 5, 4 y 2; sin valoración, por la última interacción (el alta, 11/10) y después por nombre.
    nombres = ['Biometano Meseta Ficticia', 'Montajes Ejemplo', 'Consultora Imaginaria',
               'Depósitos Imaginarios del Levante', 'Cooperativa Agraria Inventada',
               'Ingeniería Modelo Norte (grupo Modelo)']
    posiciones = [t.index(n) for n in nombres]
    assert posiciones == sorted(posiciones)
    with zipfile.ZipFile(docx) as z:
        rels = z.read('word/_rels/document.xml.rels').decode('utf-8')
    assert f'Target="fichas/Ficha_{COMPLETA}.pdf" TargetMode="External"' in rels


def test_orden_del_indice_sin_valoracion_por_ultima_interaccion():
    exp = leer(EXPEDIENTES)
    datos, _, _ = fichas.cruzar(exp, fichas.cargar_organizaciones(ORGS))
    assert [d['id'] for d in fichas.orden_indice(datos)] == [
        ALTA, EPC, COMPLETA, CONSULTORA, VACIA, 'ingenieria-modelo-norte']


def test_formatos_de_fecha_y_numero():
    assert fichas.fecha_es('2026-10-14') == '14/10/2026'
    assert fichas.fecha_es('Sin fecha') == 'Sin fecha'
    assert [fichas.numero_es(v) for v in (6000, 5870.0, 24.36, 12.5, 412500, None)] == \
        ['6.000', '5.870', '24,36', '12,5', '412.500', '']

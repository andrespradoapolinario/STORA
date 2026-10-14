"""Pruebas del módulo de fichas maestras, solo con empresas inventadas.

    python -m pytest fichas/pruebas -q

Las pruebas de generación necesitan LibreOffice y pdfplumber.
"""
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
    return ejecutar(fichas.generar, exp, salida, ORGS, panel, tracker, fecha=FECHA, **kw)


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
    errores = fichas.errores_esquema(malo)
    assert len(errores) == 6, errores
    assert len(list(validador.iter_errors(malo))) == 6


def test_reglas_que_el_esquema_no_expresa():
    exp = leer(EXPEDIENTES)
    e = exp['empresas'][COMPLETA]
    e['pendientes'][0]['estado'] = 'cerrado'
    e['cronologia'][1]['fuente'] = {'tipo': 'gmail'}
    e['lo_que_sabemos'][1]['id'] = 'sab-0001'
    exp['empresas']['otra'] = dict(e, id='distinta')
    errores = '\n'.join(fichas.errores_expedientes(exp))
    assert 'un pendiente cerrado necesita «fecha_cierre»' in errores
    assert 'necesita «fuente.ref»' in errores
    assert 'id repetido «sab-0001»' in errores
    assert 'el id «distinta» no coincide con la clave' in errores



def test_ejemplo_de_las_instrucciones_es_valido():
    instrucciones = (AQUI.parent / 'INSTRUCCIONES_FICHAS.md').read_text(encoding='utf-8')
    ejemplo = json.loads(re.search(r'```json\n(.*?)```', instrucciones, re.S).group(1))
    assert fichas.errores_esquema(ejemplo, {'$ref': '#/$defs/entradas'}, 'entradas') == []
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
    assert 'duplicada de cro-0002 (misma referencia de Gmail 19a1b2c3d4e5f601)' in informe
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
    assert 'duplicada de cro-0003 (misma fecha, tipo y resumen)' in informe
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
    cotizacion = dict(exp['cotizaciones'][0], estado='Aceptada')
    sabido = {k: v for k, v in exp['lo_que_sabemos'][1].items() if k != 'id'}
    entradas = escribir(tmp_path / 'entradas.json', [{
        'empresa': COMPLETA,
        'cotizaciones': [cotizacion],
        'lo_que_sabemos': [dict(sabido, naturaleza='dato')],
        'pendientes': [{'accion': 'Enviar planos de implantación', 'de_quien': 'cliente', 'estado': 'cerrado',
                        'fecha_cierre': '2026-10-13'}],
    }])
    codigo, informe = ejecutar(fichas.anadir, ruta, entradas)
    assert codigo == 2
    assert 'intento de modificar la entrada cot-0001 (cambia estado)' in informe
    assert 'intento de modificar la entrada sab-0002' in informe and '(cambia naturaleza)' in informe
    assert ('pen-0003 «Enviar planos de implantación» sigue abierto. '
            'Para cerrarlo, use «cerrar_pendientes»') in informe
    assert informe.count('CONFLICTO') == 3
    assert huella_archivo(ruta) == antes


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
    assert [f['rating'] for f in filas] == ['★★★', '★★', '★', '★★']


def test_asociacion_de_nombres():
    asociador = fichas.Asociador(fichas.cargar_organizaciones(ORGS), leer(EXPEDIENTES)['empresas'])
    casos = {
        'consultora-imaginaria': CONSULTORA,                          # id
        'Montajes Ejemplo, S.L.': EPC,                                # forma jurídica
        'INGENIERIA MODELO NORTE': 'ingenieria-modelo-norte',         # mayúsculas, tildes y paréntesis
        'Ingeniería Modelo Norte grupo Modelo': 'ingenieria-modelo-norte',
        'Meseta Bio': COMPLETA,                                       # nombre alternativo
        'Grupo Montajes Ejemplo S.A.': EPC,
    }
    for nombre, oid in casos.items():
        assert asociador.buscar(nombre) == (oid, ''), nombre
    assert asociador.buscar('Empresa Desconocida')[0] is None


def test_valoracion():
    assert [fichas.rango_valoracion(v) for v in ('★★★', 4, '4,5', 'Hot', 'B', '', None, 'otra')] == \
        [3, 4, 4.5, 3, 2, None, None, None]


def test_validar_lista_los_nombres_sin_correspondencia(tmp_path):
    exp = expediente_de_prueba(tmp_path)
    panel, tracker = datos_prueba.crear(tmp_path)
    codigo, informe = ejecutar(fichas.validar, exp, ORGS, panel, tracker)
    assert codigo == 0
    assert (f'AVISO      Panel, hoja «PROMOTORES», fila 5: «{datos_prueba.SIN_ASOCIAR_PANEL}» '
            'no se puede asociar a ninguna empresa') in informe
    assert f'Tracker, hoja «Lead Overview», fila 6: «{datos_prueba.SIN_ASOCIAR_TRACKER}»' in informe
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
    assert codigo == 0 and 'Fichas: 0 regenerada(s), 5 sin cambios' in informe
    assert {p.name: huella_archivo(p) for p in carpeta.iterdir()} == huellas

    entradas = escribir(tmp_path / 'entradas.json', [{
        'empresa': EPC,
        'cronologia': [{'fecha': '2026-10-13', 'tipo': 'correo_enviado', 'resumen': 'Envío de la tarifa.',
                        'origen': 'automatico', 'fuente': {'tipo': 'gmail', 'ref': '19a1b2c3d4e5f6ff'}}],
    }])
    assert ejecutar(fichas.anadir, exp, entradas)[0] == 0
    codigo, informe = generar(exp, salida, panel, tracker)
    assert codigo == 0 and 'Fichas: 1 regenerada(s), 4 sin cambios' in informe and f'Regeneradas: {EPC}.' in informe
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
    assert codigo == 0 and 'Fichas: 5 regenerada(s), 0 sin cambios' in informe


@con_libreoffice
def test_no_escribe_en_el_panel_ni_en_el_tracker(generado, tmp_path):
    antes = huella_archivo(generado.panel), huella_archivo(generado.tracker)
    assert generar(generado.exp, tmp_path / 'salida', generado.panel, generado.tracker)[0] == 0
    assert (huella_archivo(generado.panel), huella_archivo(generado.tracker)) == antes


SECCIONES = ('SITUACIÓN ACTUAL', 'CONTACTOS', 'PERFIL Y PROYECTOS', 'LO QUE SABEMOS', 'TANQUES', 'COTIZACIONES',
             'PENDIENTES ABIERTOS', 'COMPETENCIA Y OBJECIONES', 'CRONOLOGÍA', 'DOCUMENTOS Y FUENTES',
             'NOTAS DE REUNIÓN')


def secciones(ruta_docx):
    """Títulos de sección del documento, en orden."""
    xml = documento(ruta_docx)
    return [t for t in re.findall(r'<w:t[^>]*>([^<]*)</w:t>', xml) if t in SECCIONES]


@con_libreoffice
def test_ficha_completa_con_todas_las_secciones_en_orden(generado):
    ruta = generado.fichas / f'Ficha_{COMPLETA}.docx'
    assert secciones(ruta) == list(SECCIONES)
    t = texto(ruta)
    for esperado in ('FICHA MAESTRA', 'Biometano Meseta Ficticia', 'Grupo A · Promotora', '★★★',
                     'Oferta enviada', '20/10/2026', '14/10/2026', 'Copia de organizaciones.json del 13/10/2026',
                     'Interlocutor', 'ESTIMACIÓN', 'SUPUESTO', '412.500 €', '5.870', '24,36', 'AUTOMÁTICA', '19a1b2c3d4e5f601'):
        assert esperado in t, esperado
    # Cronología de la más reciente a la más antigua; solo las entradas automáticas, marcadas.
    xml = documento(ruta)
    cronologia = xml[xml.index('CRONOLOGÍA'):xml.index('DOCUMENTOS Y FUENTES')]
    fechas = re.findall(r'\d{2}/\d{2}/2026', cronologia)
    assert fechas == ['08/10/2026', '05/10/2026', '02/10/2026', '30/09/2026']
    assert cronologia.count('AUTOMÁTICA') == 2


@con_libreoffice
def test_las_secciones_vacias_se_omiten(generado):
    ruta = generado.fichas / f'Ficha_{CONSULTORA}.docx'
    assert secciones(ruta) == ['SITUACIÓN ACTUAL', 'CONTACTOS', 'PERFIL Y PROYECTOS', 'NOTAS DE REUNIÓN']
    ruta = generado.fichas / f'Ficha_{EPC}.docx'
    assert secciones(ruta) == ['CONTACTOS', 'PERFIL Y PROYECTOS', 'PENDIENTES ABIERTOS', 'CRONOLOGÍA',
                               'DOCUMENTOS Y FUENTES', 'NOTAS DE REUNIÓN']
    for docx in generado.fichas.glob('*.docx'):
        t = texto(docx).casefold()
        for marcador in ('sin datos', 'n/d', 'no localizad', 'no encontrad', 'sin verificar', 'none'):
            assert marcador not in t, (docx.name, marcador)


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
    assert codigo == 0 and '5 fichas generadas comprobadas · 0 errores' in informe


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
    e['cronologia'] = [{'id': f'cro-{n:04d}', 'fecha': '2026-10-01', 'tipo': 'nota', 'origen': 'manual',
                        'resumen': 'NOTAS DE REUNIÓN pasadas a la FICHA MAESTRA.\n' * 3,
                        'fuente': {'tipo': 'conversacion'}} for n in range(1, 41)]
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
    nombres = ['Biometano Meseta Ficticia', 'Montajes Ejemplo', 'Consultora Imaginaria',
               'Cooperativa Agraria Inventada', 'Ingeniería Modelo Norte (grupo Modelo)']
    posiciones = [t.index(n) for n in nombres]
    assert posiciones == sorted(posiciones)
    with zipfile.ZipFile(docx) as z:
        rels = z.read('word/_rels/document.xml.rels').decode('utf-8')
    assert f'Target="fichas/Ficha_{COMPLETA}.pdf" TargetMode="External"' in rels


def test_orden_del_indice_sin_valoracion_por_ultima_interaccion():
    exp = leer(EXPEDIENTES)
    datos, _, _ = fichas.cruzar(exp, fichas.cargar_organizaciones(ORGS))
    assert [d['id'] for d in fichas.orden_indice(datos)] == [
        EPC, COMPLETA, CONSULTORA, VACIA, 'ingenieria-modelo-norte']


def test_formatos_de_fecha_y_numero():
    assert fichas.fecha_es('2026-10-14') == '14/10/2026'
    assert fichas.fecha_es('Sin fecha') == 'Sin fecha'
    assert [fichas.numero_es(v) for v in (6000, 5870.0, 24.36, 12.5, 412500, None)] == \
        ['6.000', '5.870', '24,36', '12,5', '412.500', '']

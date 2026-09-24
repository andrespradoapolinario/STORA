"""Generador del bloc de notas en Word (.docx).

Lee únicamente los datos que recibe (cargados de datos/organizaciones.json) y
reproduce el formato aprobado del bloc: portada, contenido, índice por
prioridad, una ficha por organización y dos páginas de «Otro contacto».

El número de líneas de NOTAS de cada página lo decide bloc_pdf.py (se rellena
hasta el final de la página); aquí solo se recibe como parámetro.
"""
import re
import zipfile
from xml.sax.saxutils import escape

from comun import GRUPOS, PLANTILLA_BLOC, SLOTS, etiqueta_grupo, num_proyectos

N_OTRO_CONTACTO = 2
TEXTO_NO_ACUDE = '(No acude al congreso)'
DIAS = {'lunes': 'lun.', 'martes': 'mar.', 'miércoles': 'mié.', 'jueves': 'jue.',
        'viernes': 'vie.', 'sábado': 'sáb.', 'domingo': 'dom.'}
SUBTIPOS_ORGANIZACION = ('Asociación', 'Organismo público', 'Centro de investigación')

# Página (twips): A4 con márgenes de 850 arriba y abajo y 1020 a los lados.
PAG_ALTO, PAG_ANCHO, MARGEN_SUP, MARGEN_INF, MARGEN_LAT = 16838, 11906, 850, 850, 1020
ANCHO_UTIL = PAG_ANCHO - 2 * MARGEN_LAT  # 9866
ALTO_FILA_NOTAS = 454

NS = ('xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" '
      'xmlns:o="urn:schemas-microsoft-com:office:office" '
      'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
      'xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math" '
      'xmlns:v="urn:schemas-microsoft-com:vml" '
      'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" '
      'xmlns:w10="urn:schemas-microsoft-com:office:word" '
      'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
      'xmlns:wne="http://schemas.microsoft.com/office/word/2006/wordml" '
      'xmlns:sl="http://schemas.openxmlformats.org/schemaLibrary/2006/main" '
      'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
      'xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture" '
      'xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" '
      'xmlns:lc="http://schemas.openxmlformats.org/drawingml/2006/lockedCanvas" '
      'xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram" '
      'xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" '
      'xmlns:wpg="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup" '
      'xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" '
      'xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml" '
      'xmlns:w16="http://schemas.microsoft.com/office/word/2018/wordml" '
      'xmlns:w16cex="http://schemas.microsoft.com/office/word/2018/wordml/cex" '
      'xmlns:w16cid="http://schemas.microsoft.com/office/word/2016/wordml/cid" '
      'xmlns="http://schemas.microsoft.com/office/tasks/2019/documenttasks" '
      'xmlns:cr="http://schemas.microsoft.com/office/comments/2020/reactions"')

# --- Piezas básicas -----------------------------------------------------------

CAL = '<w:rFonts w:ascii="Calibri" w:cs="Calibri" w:eastAsia="Calibri" w:hAnsi="Calibri"/>'
UNI = ('<w:rFonts w:ascii="Arial Unicode MS" w:cs="Arial Unicode MS" '
       'w:eastAsia="Arial Unicode MS" w:hAnsi="Arial Unicode MS"/>')
PPR_GDOCS = ('<w:pPr><w:keepNext w:val="0"/><w:keepLines w:val="0"/><w:pageBreakBefore w:val="0"/>'
             '<w:widowControl w:val="0"/><w:pBdr><w:top w:space="0" w:sz="0" w:val="nil"/>'
             '<w:left w:space="0" w:sz="0" w:val="nil"/><w:bottom w:space="0" w:sz="0" w:val="nil"/>'
             '<w:right w:space="0" w:sz="0" w:val="nil"/><w:between w:space="0" w:sz="0" w:val="nil"/>'
             '</w:pBdr><w:shd w:fill="auto" w:val="clear"/>'
             '<w:spacing w:after="0" w:before="0" w:line="276" w:lineRule="auto"/>'
             '<w:ind w:left="0" w:right="0" w:firstLine="0"/><w:jc w:val="left"/>{rpr}</w:pPr>')
RPR_INICIAL = ('<w:rPr><w:rFonts w:ascii="Arial" w:cs="Arial" w:eastAsia="Arial" w:hAnsi="Arial"/>'
               '<w:b w:val="0"/><w:bCs w:val="0"/><w:i w:val="0"/><w:iCs w:val="0"/>'
               '<w:smallCaps w:val="0"/><w:strike w:val="0"/><w:color w:val="000000"/>'
               '<w:sz w:val="22"/><w:szCs w:val="22"/><w:u w:val="none"/>'
               '<w:shd w:fill="auto" w:val="clear"/><w:vertAlign w:val="baseline"/></w:rPr>')
P_GDOCS = '<w:p>' + PPR_GDOCS.format(rpr='') + '</w:p>'
P_VACIO = '<w:p><w:pPr/></w:p>'
P_SALTO = ('<w:p><w:pPr><w:pageBreakBefore w:val="1"/>'
           '<w:spacing w:after="0" w:before="0" w:line="20" w:lineRule="auto"/></w:pPr></w:p>')


def p_espacio(after=None, before=None):
    attrs = (f' w:after="{after}"' if after is not None else '') + \
            (f' w:before="{before}"' if before is not None else '')
    return f'<w:p><w:pPr><w:spacing{attrs} w:lineRule="auto"/></w:pPr></w:p>'


def rpr(sz, b=False, i=False, color=None, u=False, fuente=CAL, bcs_primero=False):
    s = fuente
    if b:
        s += '<w:bCs w:val="1"/><w:b w:val="1"/>' if bcs_primero else '<w:b w:val="1"/><w:bCs w:val="1"/>'
    if i:
        s += '<w:i w:val="1"/><w:iCs w:val="1"/>'
    if color:
        s += f'<w:color w:val="{color}"/>'
    if u:
        s += '<w:u w:val="single"/>'
    s += f'<w:sz w:val="{sz}"/><w:szCs w:val="{sz}"/>'
    return f'<w:rPr>{s}</w:rPr>'


def run(texto, props):
    return f'<w:r>{props}<w:t xml:space="preserve">{escape(texto)}</w:t></w:r>'


def parrafo(runs='', ppr=''):
    ppr = f'<w:pPr>{ppr}</w:pPr>' if ppr else '<w:pPr/>'
    return f'<w:p>{ppr}{runs}</w:p>'


CENTRO = '<w:jc w:val="center"/>'
DERECHA = '<w:jc w:val="right"/>'

R_ETIQUETA = rpr(14, b=True, color='000000')
R18 = rpr(18)
R18N = rpr(18, color='000000')
R16GRIS = rpr(16, color='737373')


def borde(lado, spec):
    if spec is None:
        return f'<w:{lado} w:color="000000" w:space="0" w:sz="0" w:val="nil"/>'
    color, sz = spec
    return f'<w:{lado} w:color="{color}" w:space="0" w:sz="{sz}" w:val="single"/>'


def margenes(top, left, bottom, right):
    return ''.join(f'<w:{lado} w:w="{v}.0" w:type="dxa"/>'
                   for lado, v in (('top', top), ('left', left), ('bottom', bottom), ('right', right)))


def celda(contenido, bordes=(None, None, None, None), mar=(50, 100, 50, 100), relleno=None,
          valign=None, span=None):
    tcpr = f'<w:gridSpan w:val="{span}"/>' if span else ''
    tcpr += '<w:tcBorders>' + ''.join(borde(l, b) for l, b in zip(('top', 'left', 'bottom', 'right'), bordes)) \
            + '</w:tcBorders>'
    if relleno:
        tcpr += f'<w:shd w:fill="{relleno}" w:val="clear"/>'
    tcpr += '<w:tcMar>' + margenes(*mar) + '</w:tcMar>'
    if valign:
        tcpr += f'<w:vAlign w:val="{valign}"/>'
    return f'<w:tc><w:tcPr>{tcpr}</w:tcPr>{contenido}</w:tc>'


def fila(celdas, alto=None, cabecera=False):
    trpr = '<w:cantSplit w:val="0"/>'
    if alto:
        trpr += f'<w:trHeight w:val="{alto}" w:hRule="atLeast"/>'
    trpr += f'<w:tblHeader w:val="{1 if cabecera else 0}"/>'
    return f'<w:tr><w:trPr>{trpr}</w:trPr>{"".join(celdas)}</w:tr>'


def tabla(columnas, filas):
    ancho = sum(columnas)
    grid = ''.join(f'<w:gridCol w:w="{c}"/>' for c in columnas)
    bordes = ''.join(f'<w:{l} w:color="000000" w:space="0" w:sz="4" w:val="single"/>'
                     for l in ('top', 'left', 'bottom', 'right', 'insideH', 'insideV'))
    return (f'<w:tbl><w:tblPr><w:tblStyle w:val="@TABLA@"/><w:tblW w:w="{ancho}.0" w:type="dxa"/>'
            f'<w:jc w:val="left"/><w:tblBorders>{bordes}</w:tblBorders><w:tblLayout w:type="fixed"/>'
            f'<w:tblLook w:val="0000"/></w:tblPr><w:tblGrid>{grid}<w:tblGridChange w:id="0">'
            f'<w:tblGrid>{grid}</w:tblGrid></w:tblGridChange></w:tblGrid>{"".join(filas)}</w:tbl>')


GRIS_D8 = ('d8d8d8', 4)
GRIS_CF = ('cfcfcf', 4)
NEGRO_8 = ('000000', 8)
B_INFERIOR = (None, None, GRIS_D8, None)
B_DESTACADO = (NEGRO_8, None, NEGRO_8, None)


# --- Textos derivados de los datos -------------------------------------------

def orden_prioridad(contactos):
    """Interlocutor directo primero, después asistencia confirmada y el resto."""
    return sorted(contactos, key=lambda c: (c.get('slot') != 'directo', not c.get('confirmado_feria')))


def linea_contacto_indice(c):
    t = c['nombre']
    if c.get('puesto'):
        t += f" ({c['puesto']})"
    if c.get('telefono'):
        t += f" · {c['telefono']}"
    return t


def linea_contacto_ficha(c):
    t = c['nombre']
    if c.get('departamento'):
        # Buzón o departamento: en la ficha, solo el nombre y el correo.
        return t + (f" · {c['email']}" if c.get('email') else '')
    if c.get('puesto'):
        t += f" — {c['puesto']}"
    for campo in ('email', 'telefono', 'nota'):
        if c.get(campo):
            t += f' · {c[campo]}'
    return t


def linea_telefono_empresa(org):
    if not org.get('telefono_empresa'):
        return ''
    etiqueta = 'Tel. organización' if org.get('subtipo') in SUBTIPOS_ORGANIZACION else 'Tel. empresa'
    return f"{etiqueta}: {org['telefono_empresa']}"


def cita_indice(cita):
    """«Miércoles 30/09, 13:00 · Stand 294 · …» -> «Cita mié. 30/09, 13:00».

    Solo se resume en el índice una cita con día y hora fijos.
    """
    primero = (cita or '').split(' · ')[0]
    m = re.match(r'^(\w+) (\d{1,2}/\d{1,2}), (\d{1,2}:\d{2}(?:–\d{1,2}:\d{2})?)$', primero)
    if not m or m.group(1).lower() not in DIAS:
        return ''
    return f'Cita {DIAS[m.group(1).lower()]} {m.group(2)}, {m.group(3)}'


def texto_proyectos(n):
    return f"– {n} PROYECTO{'S' if n != 1 else ''}"


# --- Portada, contenido e índice ----------------------------------------------

def portada(evento):
    contenido = (
        parrafo(run('STORA', rpr(22, b=True, color='c5c5c5')), '<w:spacing w:after="200" w:lineRule="auto"/>')
        + parrafo(run('Bloc de notas', rpr(76, b=True, color='ffffff')),
                  '<w:spacing w:after="160" w:lineRule="auto"/>')
        + parrafo(run(evento['nombre'], rpr(30, color='ffffff')))
    )
    t = tabla([ANCHO_UTIL], [fila([celda(contenido, mar=(600, 320, 640, 320), relleno='000000')])])
    return t + p_espacio(before=6600)


def contenido(orgs):
    filas = []
    for letra, nombre in GRUPOS.items():
        n = sum(1 for o in orgs if o['grupo'] == letra)
        if not n:
            continue
        filas.append(fila([
            celda(parrafo(run(f'{letra}. {nombre}', rpr(22))), B_INFERIOR, mar=(90, 40, 90, 40)),
            celda(parrafo(run(str(n), rpr(22, b=True, color='000000')), DERECHA), B_INFERIOR,
                  mar=(90, 40, 90, 80)),
        ]))
    titulo = parrafo(run('CONTENIDO', R_ETIQUETA), '<w:spacing w:after="60" w:lineRule="auto"/>')
    return titulo + tabla([8266, 1600], filas)


def lineas_indice(org):
    """Líneas de la columna PERSONA DE INTERÉS: (texto, props)."""
    lineas = []
    for c in orden_prioridad(org.get('contactos') or []):
        negrita = bool(c.get('confirmado_feria'))
        lineas.append((linea_contacto_indice(c),
                       rpr(16, b=negrita, color='000000' if negrita else '555555',
                           u=c.get('slot') == 'directo')))
    tel = linea_telefono_empresa(org)
    if tel:
        lineas.append((tel, rpr(16, color='000000')))
    if org.get('no_acude'):
        lineas.append((TEXTO_NO_ACUDE, rpr(16, b=True, color='000000')))
    cita = cita_indice(org.get('cita')) if org.get('visita') else ''
    if cita:
        lineas.append((cita, rpr(16, b=True, color='000000')))
    n = num_proyectos(org)
    if n:
        lineas.append((texto_proyectos(n), rpr(16, b=True, color='000000')))
    if not lineas:
        lineas.append(('', rpr(16, color='555555')))
    return lineas


def indice(evento, orgs, marcadores):
    titulo = parrafo(run('Índice por prioridad', rpr(32, b=True, color='000000')),
                     '<w:pageBreakBefore w:val="1"/><w:spacing w:after="60" w:lineRule="auto"/>')
    leyenda = parrafo(run(evento['leyenda_indice'], rpr(16, b=True, color='000000')),
                      '<w:spacing w:after="120" w:lineRule="auto"/>')
    b_cab = (None, None, NEGRO_8, None)
    cab = fila([
        celda(parrafo(run('VISITA', R_ETIQUETA), CENTRO), b_cab),
        celda(parrafo(run('N.º', R_ETIQUETA), CENTRO), b_cab),
        celda(parrafo(run('ORGANIZACIÓN', R_ETIQUETA)), b_cab),
        celda(parrafo(run('STAND', R_ETIQUETA), CENTRO), b_cab),
        celda(parrafo(run('PERSONA DE INTERÉS', R_ETIQUETA)), b_cab),
    ], cabecera=True)
    filas = [cab]
    grupo_actual = None
    for n, org in enumerate(orgs, 1):
        if org['grupo'] != grupo_actual:
            grupo_actual = org['grupo']
            filas.append(fila([celda(
                parrafo(run(f"{grupo_actual}. {GRUPOS[grupo_actual]}", rpr(18, b=True, color='000000'))),
                B_INFERIOR, relleno='f1f1f1', span=5)]))
        if org.get('visita'):
            c_visita = celda(parrafo(run('VISITA', rpr(14, b=True, color='ffffff')), CENTRO), B_INFERIOR,
                             relleno='000000', valign='center')
        else:
            c_visita = celda(parrafo(run('', R18), CENTRO), B_INFERIOR, valign='center')
        nombre = (f'<w:hyperlink w:anchor="{marcadores[n - 1]}">'
                  f'{run(org["nombre"], rpr(18, b=True, color="000000"))}</w:hyperlink>')
        personas = ''.join(parrafo(run(t, p)) for t, p in lineas_indice(org))
        filas.append(fila([
            c_visita,
            celda(parrafo(run(f'{n:02d}', R18), CENTRO), B_INFERIOR, valign='center'),
            celda(parrafo(nombre), B_INFERIOR, valign='center'),
            celda(parrafo(run(org.get('stand') or '—', R18), CENTRO), B_INFERIOR, valign='center'),
            celda(personas, B_INFERIOR, valign='center'),
        ], alto=330))
    return titulo + leyenda + tabla([700, 560, 2800, 800, 5006], filas)


# --- Ficha -------------------------------------------------------------------------

def cabecera_ficha(numero, titulo, subtitulo_runs, derecha, marcador=None, columnas=(900, 5966, 3000)):
    c_num = celda(parrafo(run(numero, rpr(36, b=True, color='ffffff')), CENTRO),
                  mar=(100, 60, 100, 60), relleno='000000', valign='center')
    mc = (f'<w:bookmarkStart w:colFirst="0" w:colLast="0" w:name="{marcador}" w:id="@MARCADOR@"/>'
          f'<w:bookmarkEnd w:id="@MARCADOR@"/>') if marcador else ''
    c_tit = celda(mc + parrafo(run(titulo, rpr(32, b=True, color='000000'))) + parrafo(subtitulo_runs),
                  mar=(100, 160, 100, 100), relleno='f1f1f1', valign='center')
    celdas = [c_num, c_tit]
    if derecha is not None:
        celdas.append(celda(derecha, mar=(80, 100, 80, 160), relleno='f1f1f1', valign='center'))
    return tabla(list(columnas), [fila(celdas)])


def lineas_en_blanco(n=2):
    filas = [fila([celda(P_VACIO, (None, None, GRIS_CF, None), mar=(0, 0, 0, 0))], alto=380) for _ in range(n)]
    return P_GDOCS + tabla([7906], filas) + P_VACIO


def en_ficha(c):
    """Regla 4: en la ficha, solo personas con asistencia confirmada (y los departamentos, que no son personas)."""
    return bool(c.get('confirmado_feria') or c.get('departamento'))


def tabla_contactos(org):
    contactos = [c for c in orden_prioridad(org.get('contactos') or []) if en_ficha(c)]
    filas = []
    for slot, etiqueta in SLOTS.items():
        del_slot = [c for c in contactos if c.get('slot', 'otros') == slot]
        if slot == 'directo' and not del_slot:
            continue
        parrafos = []
        for c in del_slot:
            props = rpr(18, b=True, color='000000', u=True, bcs_primero=True) if slot == 'directo' else R18N
            parrafos.append(parrafo(run(linea_contacto_ficha(c), props)))
        if slot == 'otros' and linea_telefono_empresa(org):
            parrafos.append(parrafo(run(linea_telefono_empresa(org), R18N)))
        if not parrafos:
            parrafos.append(parrafo(run('', R18N)))
        filas.append(fila([
            celda(parrafo(run(etiqueta, R16GRIS)), mar=(20, 0, 20, 60)),
            celda(''.join(parrafos), mar=(20, 0, 20, 0)),
        ]))
    filas.append(fila([
        celda(parrafo(run('', R16GRIS)), mar=(20, 0, 20, 60)),
        celda(parrafo(run('', R18N)), mar=(20, 0, 20, 0)),
    ]))
    return P_GDOCS + tabla([1900, 6006], filas) + P_VACIO


def parrafo_proyecto(p):
    cola = ''
    if p.get('ubicacion'):
        cola += f" — {p['ubicacion']}"
    if p.get('anio'):
        cola += f" · {p['anio']}"
    if p.get('estado'):
        cola += ' · '
    runs = run('• ', R18) + run(p['nombre'], rpr(18, b=True))
    if cola:
        runs += run(cola, R18)
    if p.get('estado'):
        runs += run(p['estado'], rpr(18, i=True, color='000000'))
    return parrafo(runs)


def cuerpo_ficha(org):
    def fila_seccion(etiqueta, contenido, bordes=B_INFERIOR, relleno=None):
        return fila([
            celda(parrafo(run(etiqueta, R_ETIQUETA)), bordes, relleno=relleno),
            celda(contenido, bordes, relleno=relleno),
        ])
    filas = []
    perfil = org.get('perfil')
    filas.append(fila_seccion('PERFIL', parrafo(run(perfil, R18)) if perfil else lineas_en_blanco()))
    filas.append(fila_seccion('CONTACTO', tabla_contactos(org)))
    proyectos = org.get('proyectos') or []
    filas.append(fila_seccion('PROYECTOS', ''.join(parrafo_proyecto(p) for p in proyectos)
                              if proyectos else lineas_en_blanco()))
    if org.get('uso_gfs') in ('Sí', 'No'):
        gfs = parrafo(run(org['uso_gfs'], rpr(18, b=True)))
    else:
        casilla = rpr(20, fuente=UNI)
        gfs = parrafo(run('☐', casilla) + run(' Sí        ', R18) + run('☐', casilla) + run(' No', R18))
    filas.append(fila_seccion('USO DE GFS', gfs))
    for etiqueta, campo in (('CITA', 'cita'), ('AVISO', 'aviso')):
        if org.get(campo):
            filas.append(fila_seccion(etiqueta, parrafo(run(org[campo], rpr(18, b=True))),
                                      B_DESTACADO, relleno='f1f1f1'))
    return tabla([1760, 8106], filas)


def tabla_notas(n_filas):
    filas = [fila([celda(parrafo(run('NOTAS', R_ETIQUETA)), (NEGRO_8, None, GRIS_CF, None), valign='bottom')],
                  alto=ALTO_FILA_NOTAS)]
    for _ in range(max(n_filas, 1) - 1):
        filas.append(fila([celda(P_VACIO, (None, None, GRIS_CF, None))], alto=ALTO_FILA_NOTAS))
    return tabla([ANCHO_UTIL], filas)


def ficha(n, org, marcador, n_notas):
    subtitulo = run(etiqueta_grupo(org), R18N)
    if org.get('no_acude'):
        subtitulo += run('  ' + TEXTO_NO_ACUDE, rpr(18, b=True, color='000000'))
    derecha = (parrafo(run('STAND', rpr(13, b=True, color='000000')), DERECHA)
               + parrafo(run(org.get('stand') or '—', rpr(40, b=True, color='000000')), DERECHA))
    if org.get('nota_stand'):
        derecha += parrafo(run(org['nota_stand'], rpr(15, color='737373')), DERECHA)
    return (P_SALTO
            + cabecera_ficha(f'{n:02d}', org['nombre'], subtitulo, derecha, marcador)
            + p_espacio(after=40) + cuerpo_ficha(org) + p_espacio(after=40) + tabla_notas(n_notas))


def otro_contacto(n, n_notas):
    cab = cabecera_ficha(f'L{n}', 'Otro contacto',
                         run('Organizaciones o personas no incluidas en el bloc', R16GRIS),
                         None, columnas=(900, 8966))
    filas = [fila([
        celda(parrafo(run(etiqueta, R_ETIQUETA)), (None, None, GRIS_CF, None), valign='bottom'),
        celda(P_VACIO, (None, None, GRIS_CF, None)),
    ], alto=620) for etiqueta in ('ORGANIZACIÓN', 'PERSONA Y CARGO', 'CORREO Y TELÉFONO', 'STAND')]
    return (P_SALTO + cab + p_espacio(after=100) + tabla([2300, 7566], filas) + p_espacio(after=60)
            + tabla_notas(n_notas))


# --- Documento completo -------------------------------------------------------

def marcadores_fichas(orgs):
    return [f'ficha_{n:02d}' for n in range(1, len(orgs) + 1)]


def paginas_con_notas(datos):
    """Número de páginas con tabla de NOTAS (fichas y «Otro contacto»), en orden."""
    return len(datos['organizaciones']) + N_OTRO_CONTACTO


def documento(datos, filas_notas=None):
    orgs = datos['organizaciones']
    evento = datos['evento']
    n_pag = paginas_con_notas(datos)
    filas_notas = list(filas_notas) if filas_notas else [1] * n_pag
    assert len(filas_notas) == n_pag
    marcadores = marcadores_fichas(orgs)
    partes = ['<w:p>' + PPR_GDOCS.format(rpr=RPR_INICIAL) + '</w:p>', portada(evento), contenido(orgs),
              indice(evento, orgs, marcadores)]
    for n, org in enumerate(orgs, 1):
        partes.append(ficha(n, org, marcadores[n - 1], filas_notas[n - 1]))
    for k in range(1, N_OTRO_CONTACTO + 1):
        partes.append(otro_contacto(k, filas_notas[len(orgs) + k - 1]))
    partes.append(P_GDOCS)
    seccion = (f'<w:sectPr><w:footerReference r:id="rId6" w:type="default"/>'
               f'<w:pgSz w:h="{PAG_ALTO}" w:w="{PAG_ANCHO}" w:orient="portrait"/>'
               f'<w:pgMar w:bottom="{MARGEN_INF}" w:top="{MARGEN_SUP}" w:left="{MARGEN_LAT}" '
               f'w:right="{MARGEN_LAT}" w:header="708" w:footer="400"/><w:pgNumType w:start="1"/></w:sectPr>')
    cuerpo = ''.join(partes)
    # Numeración de estilos de tabla y marcadores en orden de aparición.
    contador = iter(range(1, 100000))
    cuerpo = re.sub('@TABLA@', lambda m: f'Table{next(contador)}', cuerpo)
    n_tablas = cuerpo.count('<w:tblStyle ')
    ids = iter(range(0, 100000))
    cuerpo = re.sub(r'w:id="@MARCADOR@"/><w:bookmarkEnd w:id="@MARCADOR@"',
                    lambda m: (lambda i: f'w:id="{i}"/><w:bookmarkEnd w:id="{i}"')(next(ids)), cuerpo)
    xml = (f"<?xml version='1.0' encoding='UTF-8' standalone='yes'?>\n"
           f'<w:document {NS}><w:body>{cuerpo}{seccion}</w:body></w:document>')
    return xml, n_tablas


def estilos(n_tablas):
    base = (PLANTILLA_BLOC / 'word' / 'styles.xml').read_text(encoding='utf-8')
    estilo = ('<w:style w:type="table" w:styleId="Table{n}"><w:basedOn w:val="TableNormal"/><w:tblPr>'
              '<w:tblStyleRowBandSize w:val="1"/><w:tblStyleColBandSize w:val="1"/><w:tblCellMar>'
              '<w:top w:w="0.0" w:type="dxa"/><w:left w:w="115.0" w:type="dxa"/>'
              '<w:bottom w:w="0.0" w:type="dxa"/><w:right w:w="115.0" w:type="dxa"/></w:tblCellMar>'
              '</w:tblPr></w:style>')
    tablas = ''.join(estilo.format(n=n) for n in range(1, n_tablas + 1))
    return base.replace('</w:styles>', tablas + '</w:styles>')


def pie(evento):
    base = (PLANTILLA_BLOC / 'word' / 'footer1.xml').read_text(encoding='utf-8')
    return base.replace('{PIE}', escape(evento['pie']))


ESTATICOS = ('[Content_Types].xml', '_rels/.rels', 'word/_rels/document.xml.rels', 'word/settings.xml',
             'word/fontTable.xml', 'word/numbering.xml', 'word/theme/theme1.xml')


def escribir(datos, ruta, filas_notas=None):
    xml, n_tablas = documento(datos, filas_notas)
    with zipfile.ZipFile(ruta, 'w', zipfile.ZIP_DEFLATED) as z:
        z.writestr('[Content_Types].xml', (PLANTILLA_BLOC / '[Content_Types].xml').read_bytes())
        z.writestr('_rels/.rels', (PLANTILLA_BLOC / '_rels' / '.rels').read_bytes())
        z.writestr('word/document.xml', xml.encode('utf-8'))
        for parte in ESTATICOS[2:]:
            z.writestr(parte, (PLANTILLA_BLOC / parte).read_bytes())
        z.writestr('word/styles.xml', estilos(n_tablas).encode('utf-8'))
        z.writestr('word/footer1.xml', pie(datos['evento']).encode('utf-8'))
    return ruta

"""Generador del bloc de notas en PDF y ajuste de las líneas de NOTAS.

El PDF se obtiene convirtiendo el Word con LibreOffice, de modo que ambos
tienen exactamente el mismo contenido. Antes se mide cada página para que la
tabla de NOTAS llegue hasta el final de la página sin desbordarla.
"""
import math
import shutil
import subprocess
import tempfile
from pathlib import Path

import bloc_docx

PASO_NOTAS = bloc_docx.ALTO_FILA_NOTAS / 20            # 22,7 pt por línea de notas
LIMITE_INFERIOR = (bloc_docx.PAG_ALTO - bloc_docx.MARGEN_INF) / 20  # pt desde el borde superior
# Holgura que se deja al final de cada página para que el Word no desborde en
# otros programas (Word, Google Docs), cuyo cálculo de alturas difiere un poco.
MARGEN_SEGURIDAD = 12


class ErrorPaginacion(Exception):
    pass


RUTAS_LIBREOFFICE = (
    '/Applications/LibreOffice.app/Contents/MacOS/soffice',
    r'C:\Program Files\LibreOffice\program\soffice.exe',
    r'C:\Program Files (x86)\LibreOffice\program\soffice.exe',
)


def _soffice():
    for nombre in ('soffice', 'libreoffice'):
        ruta = shutil.which(nombre)
        if ruta:
            return ruta
    for ruta in RUTAS_LIBREOFFICE:
        if Path(ruta).exists():
            return ruta
    raise ErrorPaginacion('No se encuentra LibreOffice (soffice). Es necesario para generar el PDF.')


def convertir_a_pdf(docx, carpeta):
    """Convierte un .docx a PDF con LibreOffice y devuelve la ruta del PDF."""
    docx, carpeta = Path(docx), Path(carpeta)
    perfil = Path(tempfile.mkdtemp(prefix='stora_lo_'))
    try:
        orden = [_soffice(), f'-env:UserInstallation={perfil.as_uri()}', '--headless',
                 '--convert-to', 'pdf', '--outdir', str(carpeta), str(docx)]
        r = subprocess.run(orden, capture_output=True, text=True, timeout=900)
    finally:
        shutil.rmtree(perfil, ignore_errors=True)
    pdf = carpeta / (docx.stem + '.pdf')
    if r.returncode != 0 or not pdf.exists():
        raise ErrorPaginacion(f'LibreOffice no pudo convertir {docx.name}:\n{r.stdout}\n{r.stderr}')
    return pdf


def etiquetas(datos):
    """Número que encabeza cada página con NOTAS: «01»…«88», «L1», «L2»."""
    n = len(datos['organizaciones'])
    return [f'{i:02d}' for i in range(1, n + 1)] + [f'L{k}' for k in range(1, bloc_docx.N_OTRO_CONTACTO + 1)]


def analizar(pdf, datos):
    """Localiza la página de cada ficha y mide dónde termina su contenido.

    Devuelve (total_paginas, paginas, fondo): paginas[i] es la lista de páginas
    (índice 0) que ocupa la ficha i y fondo[i] la coordenada vertical (pt) de la
    última línea horizontal de su primera página.
    """
    import pdfplumber

    esperadas = etiquetas(datos)
    inicio = {}
    fondo_pagina = []
    with pdfplumber.open(str(pdf)) as doc:
        total = len(doc.pages)
        siguiente = 0
        for num_pag, pagina in enumerate(doc.pages):
            palabras = pagina.extract_words(extra_attrs=['size'])
            cabecera = {w['text'] for w in palabras if w['size'] > 15 and w['x0'] < 100 and w['top'] < 140}
            if siguiente < len(esperadas) and esperadas[siguiente] in cabecera:
                inicio[siguiente] = num_pag
                siguiente += 1
            horizontales = [e['top'] for e in pagina.edges
                            if e['orientation'] == 'h' and e['x1'] - e['x0'] > 300]
            fondo_pagina.append(max(horizontales) if horizontales else None)
    if len(inicio) != len(esperadas):
        faltan = [esperadas[i] for i in range(len(esperadas)) if i not in inicio]
        raise ErrorPaginacion(f'No se localizan en el PDF las páginas de: {", ".join(faltan)}')
    paginas, fondo = [], []
    for i in range(len(esperadas)):
        fin = inicio[i + 1] if i + 1 < len(esperadas) else total
        paginas.append(list(range(inicio[i], fin)))
        fondo.append(fondo_pagina[inicio[i]])
    return total, paginas, fondo


def fichas_desbordadas(datos, paginas):
    nombres = [o['nombre'] for o in datos['organizaciones']] + \
              ['Otro contacto'] * bloc_docx.N_OTRO_CONTACTO
    return [(etq, nombres[i], len(p)) for i, (etq, p) in enumerate(zip(etiquetas(datos), paginas)) if len(p) != 1]


def generar(datos, ruta_docx, ruta_pdf, informe=print):
    """Genera el Word y el PDF con las NOTAS ajustadas a una página por ficha."""
    ruta_docx, ruta_pdf = Path(ruta_docx), Path(ruta_pdf)
    n = bloc_docx.paginas_con_notas(datos)
    with tempfile.TemporaryDirectory(prefix='stora_bloc_') as tmp:
        tmp = Path(tmp)
        # 1) Medición: cada ficha con la cabecera de NOTAS solamente.
        medida = bloc_docx.escribir(datos, tmp / 'medida.docx', [1] * n)
        informe('Midiendo el espacio libre de cada página…')
        _, paginas, fondo = analizar(convertir_a_pdf(medida, tmp), datos)
        malas = fichas_desbordadas(datos, paginas)
        if malas:
            raise ErrorPaginacion('Fichas que no caben en una página ni siquiera sin líneas de notas: '
                                  + '; '.join(f'{e} {nom}' for e, nom, _ in malas))
        filas = [1 + max(0, math.floor((LIMITE_INFERIOR - y - MARGEN_SEGURIDAD) / PASO_NOTAS)) for y in fondo]
        # 2) Documento definitivo; si alguna ficha desborda, se quita una línea y se repite.
        for _ in range(4):
            bloc_docx.escribir(datos, tmp / ruta_docx.name, filas)
            informe('Convirtiendo el Word a PDF…')
            pdf = convertir_a_pdf(tmp / ruta_docx.name, tmp)
            total, paginas, _ = analizar(pdf, datos)
            malas = fichas_desbordadas(datos, paginas)
            if not malas:
                break
            for i, p in enumerate(paginas):
                if len(p) != 1:
                    filas[i] = max(1, filas[i] - 1)
        else:
            raise ErrorPaginacion('No se ha conseguido una ficha por página: '
                                  + '; '.join(f'{e} {nom}' for e, nom, _ in malas))
        ruta_docx.parent.mkdir(parents=True, exist_ok=True)
        shutil.copyfile(tmp / ruta_docx.name, ruta_docx)
        shutil.copyfile(pdf, ruta_pdf)
    return total, filas

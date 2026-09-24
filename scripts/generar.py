"""Genera los documentos a partir de datos/organizaciones.json.

Uso:
    python scripts/generar.py bloc [--fecha AAAA-MM-DD]

Ejecuta antes validar.py y se detiene si hay errores. Deja el Word y el PDF
en salidas/ con el nombre Titulo_Descriptivo_AAAA-MM-DD (por defecto, la
fecha de hoy).
"""
import argparse
import datetime
import sys

import bloc_pdf
import validar
from comun import SALIDAS, cargar


def generar_bloc(fecha):
    datos = cargar()
    errores, avisos = validar.validar(datos)
    if avisos:
        print(f'Validación: {len(avisos)} avisos (detalle con «python scripts/validar.py»).')
    if errores:
        for m in errores:
            print('ERROR  ', m)
        print(f'\nValidación fallida: {len(errores)} errores. No se genera ningún documento.')
        return 1
    base = f"{datos['evento']['archivo_bloc']}_{fecha}"
    docx, pdf = SALIDAS / f'{base}.docx', SALIDAS / f'{base}.pdf'
    try:
        bloc_pdf.generar(datos, docx, pdf)
        total, malas = validar.comprobar_paginacion(pdf, datos)
    except bloc_pdf.ErrorPaginacion as exc:
        print('ERROR  ', exc)
        return 1
    if malas:
        for m in malas:
            print('ERROR  ', m)
        return 1
    n = len(datos['organizaciones'])
    print(f'Generado: {docx.relative_to(SALIDAS.parent)}')
    print(f'Generado: {pdf.relative_to(SALIDAS.parent)}')
    print(f'{n} fichas · {total} páginas · una ficha por página, sin desbordes.')
    return 0


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument('documento', choices=['bloc'], help='documento que se genera')
    ap.add_argument('--fecha', default=datetime.date.today().isoformat(),
                    help='fecha del nombre de archivo (AAAA-MM-DD); por defecto, hoy')
    args = ap.parse_args()
    try:
        datetime.date.fromisoformat(args.fecha)
    except ValueError:
        ap.error('la fecha debe tener el formato AAAA-MM-DD')
    return generar_bloc(args.fecha)


if __name__ == '__main__':
    sys.exit(main())

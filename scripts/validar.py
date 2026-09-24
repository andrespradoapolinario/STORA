"""Validación de datos/organizaciones.json (apartado 8 de CLAUDE.md).

Uso:
    python scripts/validar.py                 valida los datos
    python scripts/validar.py --pdf RUTA.pdf  comprueba además la paginación de un bloc generado

Los errores detienen la generación. Los avisos señalan datos que conviene
revisar según las reglas del apartado 4, pero no la detienen.
"""
import argparse
import re
import sys

from comun import (CAMPOS_CONTACTO, CAMPOS_ORGANIZACION, CAMPOS_PROYECTO, DATOS, ESTADOS, GRUPOS,
                   MAX_PROYECTOS, SLOTS, USO_GFS, cargar)

# Apartado 8: términos prohibidos en cualquier texto.
PROHIBIDOS = re.compile(r'no localizad|no encontrad|sin verificar|fuente:', re.IGNORECASE)
# Apartado 4.2 y nombre de la marca.
MARCADORES = re.compile(r'\bN/D\b|Omerastore', re.IGNORECASE)
# Regla 3: correos generales o de departamento. Se mantienen hasta tener uno
# personal de esa organización, así que solo avisan.
GENERALES = re.compile(r'info@|contact@|web@|hablamos@|comercial@|sales@|service@', re.IGNORECASE)
GENERICOS = re.compile(r'^(info|informacion|contacto|contact|ventas|compras|comercial|admin|administracion'
                       r'|oficina|office|hola|hello|atencion|clientes|soporte|support|canalproveedor'
                       r'|proveedores|marketing|prensa|rrhh|general)([._-].*)?$', re.IGNORECASE)
EMAIL = re.compile(r'[\w.+-]+@[\w-]+(\.[\w-]+)+')
ID = re.compile(r'^[a-z0-9]+(-[a-z0-9]+)*$')
ANIO = re.compile(r'^\d{4}(–\d{4})?$')
EMOJI = re.compile('[\U0001F300-\U0001FAFF☀-⛿✀-➿]')
NO_IMPRESOS = {'fuente_interna', 'id'}


def textos(org):
    """(ruta, texto) de todos los campos que se imprimen."""
    for k, v in org.items():
        if k in NO_IMPRESOS:
            continue
        if isinstance(v, str):
            yield k, v
        elif isinstance(v, list):
            for i, elem in enumerate(v):
                for kk, vv in elem.items():
                    if isinstance(vv, str):
                        yield f'{k}[{i + 1}].{kk}', vv


def es_generico(email):
    local, _, dominio = email.partition('@')
    etiqueta = dominio.split('.')[0].lower()
    return bool(GENERALES.search(email) or GENERICOS.match(local)) or local.lower() == etiqueta


def validar(datos):
    errores, avisos = [], []
    if not isinstance(datos, dict) or 'organizaciones' not in datos or 'evento' not in datos:
        return ['El archivo debe tener las claves «evento» y «organizaciones».'], []
    for campo in ('nombre', 'pie', 'leyenda_indice', 'archivo_bloc'):
        if not datos['evento'].get(campo):
            errores.append(f'evento: falta «{campo}».')
    orgs = datos['organizaciones']
    ids, nombres = {}, {}
    for n, org in enumerate(orgs, 1):
        ref = f"{n:02d} {org.get('nombre') or org.get('id') or '(sin nombre)'}"

        def e(m, ref=ref):
            errores.append(f'{ref}: {m}')

        def a(m, ref=ref):
            avisos.append(f'{ref}: {m}')

        for campo in ('id', 'nombre', 'grupo'):
            if not org.get(campo):
                e(f'falta «{campo}».')
        oid = org.get('id', '')
        if oid and not ID.match(oid):
            e(f'id «{oid}» no válido (minúsculas, números y guiones).')
        if oid in ids:
            e(f'id repetido (también en {ids[oid]}).')
        ids.setdefault(oid, ref)
        clave = (org.get('nombre') or '').casefold()
        if clave in nombres:
            e(f'nombre repetido (también en {nombres[clave]}).')
        nombres.setdefault(clave, ref)
        if org.get('grupo') not in GRUPOS:
            e(f"grupo «{org.get('grupo')}» fuera de A–E.")
        for campo in org:
            if campo not in CAMPOS_ORGANIZACION:
                a(f'campo desconocido «{campo}».')
        for campo in ('visita', 'no_acude'):
            if not isinstance(org.get(campo, False), bool):
                e(f'«{campo}» debe ser true o false.')
        if org.get('uso_gfs') not in USO_GFS:
            e(f"uso_gfs «{org.get('uso_gfs')}» no válido (Sí, No o null).")
        if org.get('cita') and not org.get('visita'):
            a('tiene cita pero «visita» es false.')
        if org.get('visita') and not org.get('cita'):
            a('«visita» es true pero no hay cita.')
        # Textos
        generales = []
        for ruta, texto in textos(org):
            for correo in EMAIL.finditer(texto):
                if es_generico(correo.group(0)) and correo.group(0) not in generales:
                    generales.append(correo.group(0))
            m = PROHIBIDOS.search(texto)
            if m:
                e(f'«{m.group(0)}» en {ruta}: «{texto}».')
            m = MARCADORES.search(texto)
            if m:
                e(f'«{m.group(0)}» en {ruta}.')
            if '!' in texto or '¡' in texto:
                a(f'signo de exclamación en {ruta}.')
            if EMOJI.search(texto):
                a(f'emoji en {ruta}.')
            if texto != texto.strip() or '  ' in texto:
                a(f'espacios sobrantes en {ruta}.')
        # Contactos
        vistos = set()
        for i, c in enumerate(org.get('contactos') or [], 1):
            rc = f"contacto {i} ({c.get('nombre', '')})"
            if not c.get('nombre'):
                e(f'contacto {i} sin nombre.')
            if c.get('slot') not in SLOTS:
                pista = ' Para un interlocutor, use «interlocutor»: true.' if c.get('slot') == 'directo' else ''
                e(f"{rc}: slot «{c.get('slot')}» no válido ({', '.join(SLOTS)}).{pista}")
            for campo in ('confirmado_feria', 'interes', 'interlocutor', 'departamento'):
                if not isinstance(c.get(campo, False), bool):
                    e(f'{rc}: «{campo}» debe ser true o false.')
            for campo in c:
                if campo not in CAMPOS_CONTACTO:
                    a(f'{rc}: campo desconocido «{campo}».')
            if c.get('nombre', '').casefold() in vistos:
                a(f'{rc}: contacto repetido.')
            vistos.add(c.get('nombre', '').casefold())
            email = c.get('email', '')
            if email and not EMAIL.fullmatch(email):
                a(f'{rc}: correo con formato no válido «{email}».')
            if '@' in c.get('telefono', ''):
                a(f"{rc}: el teléfono contiene un correo («{c['telefono']}»).")
        for correo in generales:
            quien = next((c['nombre'] for c in org.get('contactos') or []
                          if correo in (c.get('email', ''), c.get('nota', ''))), '')
            a(f"correo general o de departamento{f' de {quien}' if quien else ''}: «{correo}». "
              'Se mantiene hasta disponer de uno personal (regla 3).')
        # Proyectos
        proyectos = org.get('proyectos') or []
        if len(proyectos) > MAX_PROYECTOS:
            e(f'{len(proyectos)} proyectos (máximo {MAX_PROYECTOS}).')
        for i, p in enumerate(proyectos, 1):
            if not p.get('nombre'):
                e(f'proyecto {i} sin nombre.')
            if p.get('estado', '') not in ESTADOS:
                e(f"proyecto {i}: estado «{p.get('estado')}» no admitido.")
            if p.get('anio') and not ANIO.match(p['anio']):
                a(f"proyecto {i}: año «{p['anio']}» con formato inesperado.")
            for campo in p:
                if campo not in CAMPOS_PROYECTO:
                    a(f'proyecto {i}: campo desconocido «{campo}».')
        conocidos = org.get('proyectos_conocidos')
        if conocidos is not None:
            if not isinstance(conocidos, int) or conocidos < 0:
                e('«proyectos_conocidos» debe ser un número entero o null.')
            elif conocidos < len(proyectos):
                e(f'«proyectos_conocidos» ({conocidos}) es menor que los proyectos listados ({len(proyectos)}).')
    return errores, avisos


def comprobar_paginacion(ruta_pdf, datos):
    """Una ficha por página y ninguna desbordada (apartado 8)."""
    import bloc_pdf
    total, paginas, _ = bloc_pdf.analizar(ruta_pdf, datos)
    malas = bloc_pdf.fichas_desbordadas(datos, paginas)
    return total, [f'{e} {nom}: ocupa {n} páginas' for e, nom, n in malas]


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument('--pdf', help='comprueba la paginación de un bloc ya generado')
    ap.add_argument('--datos', default=str(DATOS), help=argparse.SUPPRESS)
    args = ap.parse_args()
    datos = cargar(args.datos)
    errores, avisos = validar(datos)
    for m in avisos:
        print('AVISO  ', m)
    for m in errores:
        print('ERROR  ', m)
    if args.pdf and not errores:
        total, malas = comprobar_paginacion(args.pdf, datos)
        for m in malas:
            print('ERROR  ', m)
        errores += malas
        if not malas:
            print(f'Paginación correcta: {total} páginas, una ficha por página.')
    n = len(datos.get('organizaciones', []))
    print(f'\n{n} organizaciones · {len(errores)} errores · {len(avisos)} avisos')
    return 1 if errores else 0


if __name__ == '__main__':
    sys.exit(main())

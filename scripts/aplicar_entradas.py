"""Aplica los parches pendientes de datos/entrada/ a datos/organizaciones.json.

Uso:
    python scripts/aplicar_entradas.py             aplica y mueve cada parche a datos/entrada/aplicadas/
    python scripts/aplicar_entradas.py --simular   muestra los cambios sin escribir nada

Formato de parche: apartado 7 de CLAUDE.md. Los parches se procesan por orden
de nombre. Cada parche se aplica entero o no se aplica: si alguna operación da
conflicto (un dato ya existente con otro valor) o error, el script se detiene
sin modificar nada de ese parche ni de los siguientes. El informe recoge todos
los conflictos y errores del parche, no solo el primero, e indica siempre qué
operaciones llevan "no_es_duplicado".

Si Andrés confirma que el valor nuevo debe sustituir al existente, basta con
añadir "sustituir": true a esa operación del parche y volver a ejecutar.

Protecciones frente a duplicados y pérdidas de datos:
- Un contacto nuevo cuyo nombre se parece al de uno existente de la misma
  organización (mismo correo, o un nombre contenido en el otro) es conflicto.
- Una organización «nueva» con un contacto cuyo correo ya figura en otra
  organización es conflicto.
  En ambos casos, si Andrés confirma que no es un duplicado, se añade
  "no_es_duplicado": true a la operación.
- Un valor vacío o false del parche nunca borra un dato existente.
- Los cambios que solo afectan a mayúsculas no se aplican.
"""
import argparse
import copy
import json
import re
import shutil
import sys
import unicodedata

import validar
from comun import (APLICADAS, CAMPOS_CONTACTO, CAMPOS_ORGANIZACION, CAMPOS_PROYECTO, CONTACTO_POR_DEFECTO,
                   ENTRADA, ESTADOS, GRUPOS, MAX_PROYECTOS, PROYECTO_POR_DEFECTO, SLOTS, VALORES_POR_DEFECTO,
                   cargar, crear_id, guardar)

ACCIONES = ('actualizar', 'anadir_contacto', 'anadir_proyecto', 'nueva')
NO_ACTUALIZABLES = ('id', 'contactos', 'proyectos')


class Conflicto(Exception):
    pass


class ErrorParche(Exception):
    pass


def vacio(valor):
    return valor is None or valor == '' or valor is False or valor == []


def mostrar(valor):
    if valor is None or valor == '':
        return 'vacío'
    if isinstance(valor, bool):
        return 'sí' if valor else 'no'
    return f'«{valor}»'


def palabras(nombre):
    """«Ángel de Prada» -> ['angel', 'de', 'prada']: minúsculas, sin tildes ni puntuación."""
    base = unicodedata.normalize('NFKD', nombre).encode('ascii', 'ignore').decode().casefold()
    return re.findall(r'[a-z0-9]+', base)


def nombre_contenido(a, b):
    """True si un nombre está contenido en el otro («Juanjo Linares» en «Juanjo Linares Salinas»).

    Una inicial vale por la palabra que empieza por ella («L. Nardelli» en «Lavinia Nardelli»).
    """
    corto, largo = sorted((palabras(a), palabras(b)), key=len)
    if not corto:
        return False
    libres = list(largo)
    for p in corto:
        k = next((i for i, q in enumerate(libres) if q == p or (len(p) == 1 and q.startswith(p))), None)
        if k is None:
            return False
        del libres[k]
    return True


def mismo_correo(a, b):
    return bool(a) and bool(b) and a.casefold() == b.casefold()


def buscar(datos, op):
    orgs = datos['organizaciones']
    por_id = por_nombre = None
    if op.get('id'):
        por_id = next((o for o in orgs if o['id'] == op['id']), None)
        if por_id is None:
            raise ErrorParche(f"no existe ninguna organización con id «{op['id']}».")
    if op.get('nombre'):
        clave = op['nombre'].casefold()
        por_nombre = next((o for o in orgs if o['nombre'].casefold() == clave), None)
        if por_nombre is None and por_id is None:
            raise ErrorParche(f"«{op['nombre']}» no coincide con ninguna organización. "
                              'Indique el id correcto o use la acción «nueva».')
    if por_id and por_nombre and por_id is not por_nombre:
        raise Conflicto(f"el id «{op['id']}» y el nombre «{op['nombre']}» corresponden a organizaciones distintas.")
    org = por_id or por_nombre
    if org is None:
        raise ErrorParche('la operación no indica «id» ni «nombre».')
    return org


def fusionar(destino, nuevos, campos_validos, sustituir, desc, cambios, antes=None):
    """Rellena campos vacíos; un valor distinto sobre uno existente es conflicto."""
    antes = len(cambios) if antes is None else antes
    for campo, valor in nuevos.items():
        if campo not in campos_validos:
            raise ErrorParche(f'{desc}: campo desconocido «{campo}».')
        actual = destino.get(campo)
        if actual == valor:
            continue
        if vacio(valor):
            # Un valor vacío o false del parche nunca borra un dato existente.
            if not vacio(actual):
                cambios.append(f'{desc} · {campo}: se conserva {mostrar(actual)} (el parche lo trae vacío).')
            continue
        if isinstance(actual, str) and isinstance(valor, str) and actual.casefold() == valor.casefold():
            cambios.append(f'{desc} · {campo}: solo cambian mayúsculas; se conserva {mostrar(actual)}.')
            continue
        if not vacio(actual) and not sustituir:
            raise Conflicto(f'{desc}: «{campo}» ya vale {mostrar(actual)} y el parche propone {mostrar(valor)}.')
        destino[campo] = valor
        nota = ' (sustitución confirmada)' if not vacio(actual) else ''
        cambios.append(f'{desc} · {campo}: {mostrar(actual)} → {mostrar(valor)}{nota}')
    if len(cambios) == antes:
        cambios.append(f'{desc}: ya constaba, sin cambios.')


def op_actualizar(datos, op, cambios):
    org = buscar(datos, op)
    campos = op.get('campos')
    if not isinstance(campos, dict) or not campos:
        raise ErrorParche('«actualizar» necesita un objeto «campos».')
    for campo in campos:
        if campo in NO_ACTUALIZABLES:
            raise ErrorParche(f'«{campo}» no se modifica con «actualizar» '
                              '(use anadir_contacto o anadir_proyecto; el id no cambia).')
    fusionar(org, campos, CAMPOS_ORGANIZACION, op.get('sustituir'), org['nombre'], cambios)


def op_anadir_contacto(datos, op, cambios):
    org = buscar(datos, op)
    c = op.get('contacto')
    if not isinstance(c, dict) or not c.get('nombre'):
        raise ErrorParche('«anadir_contacto» necesita un «contacto» con «nombre».')
    if c.get('slot', 'otros') not in SLOTS:
        pista = ' Para un interlocutor, use «interlocutor»: true.' if c.get('slot') == 'directo' else ''
        raise ErrorParche(f"slot «{c.get('slot')}» no válido ({', '.join(SLOTS)}).{pista}")
    # «nombre_anterior» completa el nombre de un contacto existente («L. Nardelli» -> «Lavinia Nardelli»).
    anterior = c.get('nombre_anterior')
    clave = (anterior or c['nombre']).casefold()
    existente = next((x for x in org['contactos'] if x['nombre'].casefold() == clave), None)
    if anterior and existente is None:
        raise ErrorParche(f"{org['nombre']}: no hay ningún contacto «{anterior}» cuyo nombre completar.")
    if anterior and any(x is not existente and x['nombre'].casefold() == c['nombre'].casefold()
                        for x in org['contactos']):
        raise Conflicto(f"{org['nombre']}: ya existe otro contacto «{c['nombre']}»; "
                        f"indique si «{anterior}» es la misma persona.")
    if existente:
        desc = f"{org['nombre']} · contacto {existente['nombre']}"
        antes = len(cambios)
        if existente['nombre'].casefold() != c['nombre'].casefold():
            cambios.append(f"{desc} · nombre: {mostrar(existente['nombre'])} → {mostrar(c['nombre'])}")
            existente['nombre'] = c['nombre']
        resto = {k: v for k, v in c.items() if k not in ('nombre', 'nombre_anterior')}
        fusionar(existente, resto, CAMPOS_CONTACTO, op.get('sustituir'), desc, cambios, antes)
        return
    for campo in c:
        if campo not in CAMPOS_CONTACTO:
            raise ErrorParche(f'contacto {c["nombre"]}: campo desconocido «{campo}».')
    if not op.get('no_es_duplicado'):
        parecidos = []
        for x in org['contactos']:
            if mismo_correo(x.get('email'), c.get('email')):
                parecidos.append(f"«{x['nombre']}» (mismo correo)")
            elif nombre_contenido(x['nombre'], c['nombre']):
                parecidos.append(f"«{x['nombre']}» (nombre parecido)")
        if parecidos:
            raise Conflicto(f"{org['nombre']}: «{c['nombre']}» puede ser {' o '.join(parecidos)}. "
                            'Si es la misma persona, use su nombre exacto o «nombre_anterior»; '
                            'si es otra, añada "no_es_duplicado": true a la operación.')
    nuevo = {'nombre': c['nombre'], **{k: c.get(k, v) for k, v in CONTACTO_POR_DEFECTO.items()}}
    org['contactos'].append(nuevo)
    cambios.append(f"{org['nombre']}: nuevo contacto {nuevo['nombre']} ({SLOTS[nuevo['slot']]}"
                   f"{', asistencia confirmada' if nuevo['confirmado_feria'] else ''}).")


def op_anadir_proyecto(datos, op, cambios):
    org = buscar(datos, op)
    p = op.get('proyecto')
    if not isinstance(p, dict) or not p.get('nombre'):
        raise ErrorParche('«anadir_proyecto» necesita un «proyecto» con «nombre».')
    if p.get('estado', '') not in ESTADOS:
        raise ErrorParche(f"estado «{p.get('estado')}» no admitido ({', '.join(e for e in ESTADOS if e)} o vacío).")
    existente = next((x for x in org['proyectos'] if x['nombre'].casefold() == p['nombre'].casefold()), None)
    if existente:
        resto = {k: v for k, v in p.items() if k != 'nombre'}
        fusionar(existente, resto, CAMPOS_PROYECTO, op.get('sustituir'),
                 f"{org['nombre']} · proyecto {existente['nombre']}", cambios)
        return
    for campo in p:
        if campo not in CAMPOS_PROYECTO:
            raise ErrorParche(f'proyecto {p["nombre"]}: campo desconocido «{campo}».')
    if len(org['proyectos']) >= MAX_PROYECTOS:
        raise Conflicto(f"{org['nombre']} ya tiene {MAX_PROYECTOS} proyectos (máximo por ficha). "
                        f"Indique cuál sustituir para añadir «{p['nombre']}».")
    nuevo = {'nombre': p['nombre'], **{k: p.get(k, v) for k, v in PROYECTO_POR_DEFECTO.items()}}
    org['proyectos'].append(nuevo)
    cambios.append(f"{org['nombre']}: nuevo proyecto {nuevo['nombre']}.")
    conocidos = org.get('proyectos_conocidos')
    if conocidos is not None and conocidos < len(org['proyectos']):
        org['proyectos_conocidos'] = len(org['proyectos'])
        cambios.append(f"{org['nombre']}: proyectos conocidos {conocidos} → {len(org['proyectos'])} "
                       '(no puede ser menor que los proyectos listados).')


def op_nueva(datos, op, cambios):
    o = op.get('organizacion')
    if not isinstance(o, dict) or not o.get('nombre') or not o.get('grupo'):
        raise ErrorParche('«nueva» necesita una «organizacion» con «nombre» y «grupo».')
    if o['grupo'] not in GRUPOS:
        raise ErrorParche(f"grupo «{o['grupo']}» fuera de A–E.")
    for campo in o:
        if campo not in CAMPOS_ORGANIZACION:
            raise ErrorParche(f"{o['nombre']}: campo desconocido «{campo}».")
    orgs = datos['organizaciones']
    oid = o.get('id') or crear_id(o['nombre'])
    if any(x['id'] == oid for x in orgs) or any(x['nombre'].casefold() == o['nombre'].casefold() for x in orgs):
        raise Conflicto(f"ya existe una organización con id «{oid}» o nombre «{o['nombre']}».")
    if not op.get('no_es_duplicado'):
        for c in o.get('contactos') or []:
            for x in orgs:
                y = next((y for y in x['contactos'] if mismo_correo(y.get('email'), c.get('email'))), None)
                if y:
                    raise Conflicto(f"{o['nombre']}: el correo «{c['email']}» de {c.get('nombre', '')} ya figura "
                                    f"en {x['nombre']} (id «{x['id']}», contacto {y['nombre']}). "
                                    'Si es la misma organización, use anadir_contacto con ese id; '
                                    'si no, añada "no_es_duplicado": true a la operación.')
    nueva = {'id': oid, 'nombre': o['nombre'], 'grupo': o['grupo']}
    for campo, defecto in VALORES_POR_DEFECTO.items():
        nueva[campo] = copy.deepcopy(o.get(campo, defecto))
    nueva['contactos'] = [{'nombre': c.get('nombre', ''), **{k: c.get(k, v) for k, v in CONTACTO_POR_DEFECTO.items()}}
                          for c in nueva['contactos']]
    nueva['proyectos'] = [{'nombre': p.get('nombre', ''), **{k: p.get(k, v) for k, v in PROYECTO_POR_DEFECTO.items()}}
                          for p in nueva['proyectos']]
    # Al final de su grupo (el orden dentro del grupo es el de prioridad del índice).
    grupos = list(GRUPOS)
    pos = len(orgs)
    for i, x in enumerate(orgs):
        if grupos.index(x['grupo']) > grupos.index(o['grupo']):
            pos = i
            break
    orgs.insert(pos, nueva)
    cambios.append(f"Nueva organización: {nueva['nombre']} (id {oid}, grupo {nueva['grupo']}, "
                   f"posición {pos + 1} de {len(orgs)}).")


OPERACIONES = {'actualizar': op_actualizar, 'anadir_contacto': op_anadir_contacto,
               'anadir_proyecto': op_anadir_proyecto, 'nueva': op_nueva}


def aplicar_parche(datos, operaciones):
    """Aplica un parche sobre una copia, sin detenerse en la primera operación que falla.

    Devuelve (datos_nuevos, informe). El informe es la lista ordenada de
    (tipo, texto), con tipo «cambio», «CONFLICTO» o «ERROR». Una operación que
    falla no deja ningún cambio, y las siguientes se evalúan sin ella.
    """
    if not isinstance(operaciones, list):
        raise ErrorParche('el parche debe ser una lista de operaciones.')
    nuevos = copy.deepcopy(datos)
    informe = []
    for n, op in enumerate(operaciones, 1):
        if not isinstance(op, dict) or op.get('accion') not in ACCIONES:
            informe.append(('ERROR', f"operación {n}: «accion» debe ser una de {', '.join(ACCIONES)}."))
            continue
        prueba, cambios = copy.deepcopy(nuevos), []
        try:
            OPERACIONES[op['accion']](prueba, op, cambios)
        except Conflicto as exc:
            informe.append(('CONFLICTO', f'operación {n} ({op["accion"]}): {exc}'))
        except ErrorParche as exc:
            informe.append(('ERROR', f'operación {n} ({op["accion"]}): {exc}'))
        else:
            nuevos = prueba
            informe += [('cambio', c) for c in cambios]
    return nuevos, informe


def marcadas_no_duplicado(operaciones):
    """Línea del informe con las operaciones que llevan «no_es_duplicado»."""
    if not isinstance(operaciones, list):
        return None
    marcadas = []
    for n, op in enumerate(operaciones, 1):
        if not isinstance(op, dict) or not op.get('no_es_duplicado'):
            continue
        ref = op.get('id') or op.get('nombre') or ''
        if op.get('accion') == 'nueva':
            ref = f"nueva «{(op.get('organizacion') or {}).get('nombre', '')}»"
        elif isinstance(op.get('contacto'), dict):
            ref = f"{ref} · {op['contacto'].get('nombre', '')}"
        marcadas.append(f'{n} ({ref})')
    if not marcadas:
        return 'Ninguna operación lleva «no_es_duplicado».'
    cuales = f'operación {marcadas[0]}' if len(marcadas) == 1 else f"operaciones {', '.join(marcadas)}"
    return f'Llevan «no_es_duplicado» (confirmado por Andrés): {cuales}.'


def destino_libre(nombre):
    destino = APLICADAS / nombre
    k = 2
    while destino.exists():
        destino = APLICADAS / f'{destino.stem.rsplit("__", 1)[0]}__{k}{destino.suffix}'
        k += 1
    return destino


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument('--simular', action='store_true', help='muestra los cambios sin escribir nada')
    args = ap.parse_args()
    parches = sorted(p for p in ENTRADA.glob('*.json') if p.is_file())
    if not parches:
        print('No hay entradas pendientes en datos/entrada/.')
        return 0
    datos = cargar()
    errores_previos = set(validar.validar(datos)[0])
    aplicados = 0
    for i, ruta in enumerate(parches):
        print(f'\n== {ruta.name}')
        try:
            with open(ruta, encoding='utf-8') as f:
                operaciones = json.load(f)
            marca = marcadas_no_duplicado(operaciones)
            if marca:
                print(f'  {marca}')
            nuevos, informe = aplicar_parche(datos, operaciones)
        except json.JSONDecodeError as exc:
            print(f'ERROR      JSON no válido: {exc}')
            print('Parche no aplicado. Proceso detenido.')
            return 1
        except ErrorParche as exc:
            print(f'ERROR      {exc}')
            print('Parche no aplicado. Proceso detenido.')
            return 1
        for tipo, texto in informe:
            print(f'  · {texto}' if tipo == 'cambio' else f'{tipo:<10} {texto}')
        if not informe:
            print('  · Sin cambios: todos los datos ya constaban.')
        conflictos = sum(tipo == 'CONFLICTO' for tipo, _ in informe)
        errores = sum(tipo == 'ERROR' for tipo, _ in informe)
        for e in validar.validar(nuevos)[0]:
            if e not in errores_previos:
                print(f'ERROR      validación del resultado: {e}')
                errores += 1
        if conflictos or errores:
            partes = [f'{conflictos} conflicto(s)'] if conflictos else []
            partes += [f'{errores} error(es)'] if errores else []
            if conflictos and errores:
                pasos = 'Corrija los errores y confirme con Andrés los conflictos'
            elif conflictos:
                pasos = 'Confirme con Andrés qué valor debe quedar en cada conflicto'
            else:
                pasos = 'Corrija los errores'
            print(f"\nParche no aplicado: {' y '.join(partes)}. {pasos} antes de volver a ejecutar.")
            if len(parches) > i + 1:
                print(f'No se han procesado los parches siguientes ({len(parches) - i - 1}).')
            if args.simular:
                print('Simulación: no se ha modificado ningún archivo.')
            else:
                print('Proceso detenido: no se ha modificado ningún archivo.')
            return 1 if errores else 2
        if args.simular:
            datos = nuevos
            continue
        guardar(nuevos)
        APLICADAS.mkdir(parents=True, exist_ok=True)
        destino = destino_libre(ruta.name)
        shutil.move(str(ruta), destino)
        print(f'  Aplicado y movido a {destino.relative_to(ENTRADA.parent.parent)}')
        datos = nuevos
        aplicados += 1
    if args.simular:
        print('\nSimulación: no se ha modificado ningún archivo.')
    else:
        print(f'\n{aplicados} parche(s) aplicado(s). Siguiente paso: python scripts/generar.py bloc')
    return 0


if __name__ == '__main__':
    sys.exit(main())

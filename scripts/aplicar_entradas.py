"""Aplica los parches pendientes de datos/entrada/ a datos/organizaciones.json.

Uso:
    python scripts/aplicar_entradas.py             aplica y mueve cada parche a datos/entrada/aplicadas/
    python scripts/aplicar_entradas.py --simular   muestra los cambios sin escribir nada

Formato de parche: apartado 7 de CLAUDE.md. Los parches se procesan por orden
de nombre. Cada parche se aplica entero o no se aplica: ante un conflicto
(un dato ya existente con otro valor) o un error, el script se detiene sin
modificar nada de ese parche ni de los siguientes.

Si Andrés confirma que el valor nuevo debe sustituir al existente, basta con
añadir "sustituir": true a esa operación del parche y volver a ejecutar.
"""
import argparse
import copy
import json
import shutil
import sys

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


def fusionar(destino, nuevos, campos_validos, sustituir, desc, cambios):
    """Rellena campos vacíos; un valor distinto sobre uno existente es conflicto."""
    antes = len(cambios)
    for campo, valor in nuevos.items():
        if campo not in campos_validos:
            raise ErrorParche(f'{desc}: campo desconocido «{campo}».')
        actual = destino.get(campo)
        if actual == valor:
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
        raise ErrorParche(f"slot «{c.get('slot')}» no válido ({', '.join(SLOTS)}).")
    existente = next((x for x in org['contactos'] if x['nombre'].casefold() == c['nombre'].casefold()), None)
    if existente:
        resto = {k: v for k, v in c.items() if k != 'nombre'}
        fusionar(existente, resto, CAMPOS_CONTACTO, op.get('sustituir'),
                 f"{org['nombre']} · contacto {existente['nombre']}", cambios)
        return
    for campo in c:
        if campo not in CAMPOS_CONTACTO:
            raise ErrorParche(f'contacto {c["nombre"]}: campo desconocido «{campo}».')
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
    """Aplica un parche sobre una copia. Devuelve (datos_nuevos, cambios)."""
    if not isinstance(operaciones, list):
        raise ErrorParche('el parche debe ser una lista de operaciones.')
    nuevos = copy.deepcopy(datos)
    cambios = []
    for n, op in enumerate(operaciones, 1):
        if not isinstance(op, dict) or op.get('accion') not in ACCIONES:
            raise ErrorParche(f"operación {n}: «accion» debe ser una de {', '.join(ACCIONES)}.")
        try:
            OPERACIONES[op['accion']](nuevos, op, cambios)
        except (Conflicto, ErrorParche) as exc:
            raise type(exc)(f'operación {n} ({op["accion"]}): {exc}') from None
    return nuevos, cambios


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
    for ruta in parches:
        print(f'\n== {ruta.name}')
        try:
            with open(ruta, encoding='utf-8') as f:
                operaciones = json.load(f)
            nuevos, cambios = aplicar_parche(datos, operaciones)
        except json.JSONDecodeError as exc:
            print(f'ERROR      JSON no válido: {exc}')
            print('Parche no aplicado. Proceso detenido.')
            return 1
        except Conflicto as exc:
            print(f'CONFLICTO  {exc}')
            print('Parche no aplicado. Proceso detenido: confirme qué valor debe quedar.')
            return 2
        except ErrorParche as exc:
            print(f'ERROR      {exc}')
            print('Parche no aplicado. Proceso detenido.')
            return 1
        errores_nuevos = [e for e in validar.validar(nuevos)[0] if e not in errores_previos]
        if errores_nuevos:
            for e in errores_nuevos:
                print(f'ERROR      {e}')
            print('Parche no aplicado: el resultado no supera la validación. Proceso detenido.')
            return 1
        for c in cambios or ['Sin cambios: todos los datos ya constaban.']:
            print(f'  · {c}')
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

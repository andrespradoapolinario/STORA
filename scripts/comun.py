"""Rutas, constantes del esquema y utilidades compartidas por los scripts."""
import json
import re
import unicodedata
from pathlib import Path

RAIZ = Path(__file__).resolve().parent.parent
DATOS = RAIZ / 'datos' / 'organizaciones.json'
ENTRADA = RAIZ / 'datos' / 'entrada'
APLICADAS = ENTRADA / 'aplicadas'
SALIDAS = RAIZ / 'salidas'
PLANTILLA_BLOC = Path(__file__).resolve().parent / 'plantilla_bloc'

GRUPOS = {
    'A': 'Promotoras',
    'B': 'EPC',
    'C': 'Ingenierías',
    'D': 'Consultorías',
    'E': 'Cooperativas y asociaciones',
}
ESTADOS = ('Terminado', 'En construcción', 'En desarrollo y planificación', '')
# Filas de CONTACTO de la ficha, en orden de impresión.
SLOTS = {
    'directo': 'Contacto directo',
    'compras': 'Responsable de Compras',
    'tecnico': 'Director Técnico',
    'otros': 'Otros',
}
USO_GFS = ('Sí', 'No', None)
MAX_PROYECTOS = 5

# Orden de los campos al guardar el JSON (esquema del apartado 3 de CLAUDE.md
# más los campos que usa el bloc actual).
CAMPOS_ORGANIZACION = (
    'id', 'nombre', 'grupo', 'subtipo', 'stand', 'nota_stand', 'perfil',
    'contactos', 'telefono_empresa', 'proyectos', 'proyectos_conocidos',
    'uso_gfs', 'visita', 'cita', 'aviso', 'no_acude', 'fuente_interna',
)
CAMPOS_CONTACTO = (
    'nombre', 'puesto', 'telefono', 'email', 'slot', 'confirmado_feria',
    'interes', 'departamento', 'nota',
)
CAMPOS_PROYECTO = ('nombre', 'ubicacion', 'anio', 'estado')

VALORES_POR_DEFECTO = {
    'subtipo': '', 'stand': '', 'nota_stand': '', 'perfil': '', 'contactos': [],
    'telefono_empresa': '', 'proyectos': [], 'proyectos_conocidos': None,
    'uso_gfs': None, 'visita': False, 'cita': '', 'aviso': '', 'no_acude': False,
    'fuente_interna': '',
}
CONTACTO_POR_DEFECTO = {
    'puesto': '', 'telefono': '', 'email': '', 'slot': 'otros',
    'confirmado_feria': False, 'interes': True, 'departamento': False, 'nota': '',
}
PROYECTO_POR_DEFECTO = {'ubicacion': '', 'anio': '', 'estado': ''}


def cargar(ruta=DATOS):
    with open(ruta, encoding='utf-8') as f:
        return json.load(f)


def ordenar_organizacion(org):
    """Devuelve la organización con los campos en el orden canónico."""
    salida = {k: org[k] for k in CAMPOS_ORGANIZACION if k in org}
    salida.update({k: v for k, v in org.items() if k not in salida})
    salida['contactos'] = [
        {**{k: c[k] for k in CAMPOS_CONTACTO if k in c}, **{k: v for k, v in c.items() if k not in CAMPOS_CONTACTO}}
        for c in salida.get('contactos', [])
    ]
    salida['proyectos'] = [
        {**{k: p[k] for k in CAMPOS_PROYECTO if k in p}, **{k: v for k, v in p.items() if k not in CAMPOS_PROYECTO}}
        for p in salida.get('proyectos', [])
    ]
    return salida


def guardar(datos, ruta=DATOS):
    datos = dict(datos)
    datos['organizaciones'] = [ordenar_organizacion(o) for o in datos['organizaciones']]
    tmp = Path(ruta).with_suffix('.json.tmp')
    with open(tmp, 'w', encoding='utf-8') as f:
        json.dump(datos, f, ensure_ascii=False, indent=2)
        f.write('\n')
    tmp.replace(ruta)


def crear_id(nombre):
    """«Biorig (grupo Solarig)» -> «biorig»: minúsculas con guiones, sin paréntesis ni tildes."""
    base = re.sub(r'\s*\([^)]*\)', '', nombre).replace('Ø', '0').replace('ø', '0')
    base = unicodedata.normalize('NFKD', base).encode('ascii', 'ignore').decode()
    return re.sub(r'[^a-z0-9]+', '-', base.lower()).strip('-')


def etiqueta_grupo(org):
    """«Grupo A · Promotora»"""
    return f"Grupo {org['grupo']} · {org['subtipo']}" if org.get('subtipo') else f"Grupo {org['grupo']}"


def num_proyectos(org):
    """Número de proyectos conocidos que se indica en el índice."""
    if org.get('proyectos_conocidos') is not None:
        return org['proyectos_conocidos']
    return len(org.get('proyectos') or [])

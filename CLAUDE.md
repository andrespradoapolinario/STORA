# CLAUDE.md — Sistema comercial Stora (Alliance Albors y Prado)

Este repositorio genera los documentos comerciales de Alliance Albors y Prado S.L. para la venta e instalación de tanques de acero vitrificado de Stora (tecnología GFS) en España: biogás y biometano, agua y centros de datos. Usar siempre el nombre «Stora» (no «Omerastore»).

Responsable: Andrés Prado. Dirigirse a él siempre de usted, en español, con registro formal.

## 1. Reparto de trabajo

- **Claude en Cowork** investiga, lee el correo y el Excel de seguimiento de clientes, decide qué datos entran y deja las actualizaciones en `datos/entrada/`.
- **Claude Code (este repositorio)** mantiene los datos, valida y genera los documentos (PDF y Word). No investiga por su cuenta ni añade datos que no vengan de `datos/entrada/` o de una instrucción expresa de Andrés.
- Si una petición exige información nueva (contactos, proyectos, teléfonos, stands), no la invente: indíquelo y pida que se aporte.

## 2. Estructura

```
datos/organizaciones.json      Fuente única de datos. Todo documento se genera desde aquí.
datos/entrada/                 Actualizaciones pendientes (parches JSON) preparadas por Cowork.
datos/entrada/aplicadas/       Parches ya aplicados (no se borran: son el historial).
scripts/                       Generadores (PDF, Word) y validar.py.
salidas/                       Documentos generados. No se editan a mano.
CLAUDE.md                      Este archivo.
```

Un solo comando regenera todo: `python scripts/generar.py bloc` (o el que se documente en README.md).

## 3. Esquema de una organización (organizaciones.json)

El archivo tiene dos claves: `evento` (datos del evento para la portada, el pie de página y la leyenda del índice) y `organizaciones` (lista ordenada por grupo y, dentro de cada grupo, por prioridad).

```json
{
  "evento": {
    "nombre": "6.º Salón del Gas Renovable y 19.º Congreso Internacional de Bioenergía",   // portada
    "pie": "Salón del Gas Renovable y 19.º Congreso Internacional de Bioenergía · Valladolid, 29 y 30/09/2026",
    "leyenda_indice": "En negrita: asistencia confirmada al Salón o al Congreso. ...",
    "archivo_bloc": "Bloc_Notas_Salon_Gas_Renovable"   // nombre de archivo, sin la fecha
  },
  "organizaciones": [ ... ]
}
```

Cada organización:

```json
{
  "id": "tonello-energy-spain",
  "nombre": "Tonello Energy Spain",
  "grupo": "B",                     // A Promotoras, B EPC, C Ingenierías, D Consultorías, E Cooperativas y asociaciones
  "subtipo": "EPC",
  "stand": "294",
  "nota_stand": "Pabellón 2",
  "perfil": "Texto breve y verificado.",
  "contactos": [
    {"nombre": "", "puesto": "", "telefono": "", "email": "", "slot": "compras|tecnico|otros",
     "confirmado_feria": true, "interes": true, "interlocutor": false, "departamento": false, "nota": ""}
  ],
  "telefono_empresa": "",           // centralita; se imprime como «Tel. empresa:» («Tel. organización:» en asociaciones, organismos públicos y centros de investigación)
  "proyectos": [{"nombre": "", "ubicacion": "", "anio": "", "estado": "Terminado|En construcción|En desarrollo y planificación|"}],
  "proyectos_conocidos": 4,         // el «– N PROYECTOS» del índice; puede superar los 5 listados (null si no se conoce)
  "uso_gfs": "Sí|No|null",
  "visita": true,                   // reunión cerrada en la feria
  "cita": "Miércoles 30/09, 13:00 · Stand 294 · ...",
  "aviso": "",
  "no_acude": false,
  "fuente_interna": "Correo 23/09; programa oficial 19.º CIB"   // nunca se imprime en el bloc
}
```

- `interlocutor: true`: la persona con la que ya hemos hablado. En el índice va primera y subrayada; en la ficha, según la regla 4.
- `nota` del contacto: dato breve que se imprime en la ficha tras el correo y el teléfono (por ejemplo, «Ponente mar. 29/09, 12:40–13:40» o «respondió el 23/09»).
- `departamento: true`: buzón o departamento, no una persona (por ejemplo, «Compras de biometano (Dpto.)»). Va en la ficha aunque no tenga asistencia confirmada, solo con el nombre y el correo; en el índice, en gris y con su teléfono.
- `cita`: si empieza por día, fecha y hora («Miércoles 30/09, 13:00 · …»), el índice muestra además «Cita mié. 30/09, 13:00».

## 4. Reglas de contenido (obligatorias)

1. **No inventar nunca**: datos técnicos, precios, plazos, certificaciones, contactos, teléfonos, proyectos, stands, fechas ni compromisos. Lo que no conste se deja vacío.
2. **Vacío es vacío**: nunca escribir «no localizado», «no encontrado», «sin verificar», «N/D» ni marcadores similares.
3. **Correos**: se prefieren los personales. Los generales y los de departamento se mantienen hasta disponer de un correo personal de esa organización; entonces se sustituyen. **Teléfonos**: se admiten los personales y los de empresa.
4. **Contactos de la ficha**: solo personas confirmadas en la feria (programa oficial, cita cerrada o dato de Andrés). Excepción: los interlocutores (`interlocutor: true`, personas con las que ya hemos hablado) aparecen siempre; con asistencia confirmada, en la fila «Contacto directo»; sin ella, en «Otros» con «Interlocutor:» delante del nombre. Las demás personas de interés van solo al índice, en gris.
5. **Proyectos**: máximo 5 por ficha, formato «Nombre — Ubicación · Año · Estado». Estados admitidos: Terminado, En construcción, En desarrollo y planificación (o vacío).
6. **USO DE GFS**: «No» solo si está verificado que usan otro material; si no, casillas en blanco.
7. **Cifras internas** (costes, márgenes, condiciones con Stora) nunca en documentos para terceros sin instrucción expresa.
8. Distinguir dato verificado, estimación y supuesto cuando el documento sea un informe.

## 5. Reglas de formato

- Ortografía y gramática impecables (tildes, puntuación, mayúsculas). Sin emojis, sin signos de exclamación, sin adornos.
- Sin metacomentarios dentro de los documentos («aquí he quitado…», «esto es lo que pidió…»).
- Tono profesional, técnico y sobrio. Evitar anglicismos evitables y jerga vacía.
- Cada documento se entrega en PDF y, además, en versión Word con el mismo contenido.
- Nombre de archivo: `Titulo_Descriptivo_AAAA-MM-DD.pdf` / `.docx`.

### Bloc de notas de feria (formato aprobado)

- Portada (STORA, «Bloc de notas», nombre del evento), página CONTENIDO (grupo y número de fichas) e «Índice por prioridad».
- Índice: columnas VISITA (estrecha, a la izquierda; «VISITA» si hay reunión cerrada) · N.º · ORGANIZACIÓN · STAND · PERSONA DE INTERÉS. En la última columna: nombre (puesto) · teléfono si existe · «– N PROYECTOS» (solo el número de proyectos conocidos). Negrita = asistencia confirmada; gris = otros contactos; «(No acude al congreso)» en negrita negra cuando proceda (el bloc se imprime en blanco y negro).
- Una ficha por página: PERFIL, STAND arriba a la derecha, CONTACTO (Responsable de Compras, Director Técnico, Otros; líneas en blanco si no hay persona confirmada), PROYECTOS, USO DE GFS, CITA y AVISO destacados, NOTAS con líneas hasta el final de la página.
- Etiqueta de grupo en la ficha: «Grupo A · Promotora», «Grupo B · EPC», etc.

## 6. Flujo de actualización

1. Cowork deja un parche en `datos/entrada/AAAA-MM-DD_cowork.json` (formato en el apartado 7).
2. Andrés pide a Code: «aplica las entradas pendientes y regenera».
3. Code aplica cada parche, ejecuta `validar.py`, regenera los documentos, mueve el parche a `datos/entrada/aplicadas/` y hace commit con un mensaje en español que resuma los cambios.
4. Si un parche choca con un dato existente (por ejemplo, otra fecha de cita), no lo resuelva por su cuenta: muestre el conflicto y pregunte.

## 7. Formato de parche

Cada operación identifica la organización por `id` o, si Cowork no lo conoce, por `nombre` (coincidencia exacta, sin distinguir mayúsculas). Si el nombre no coincide con ninguna organización y la acción no es «nueva», detenerse y preguntar.

```json
[
  {"accion": "actualizar", "id": "id-energy-group",
   "campos": {"visita": true, "cita": "Miércoles 30/09, 12:30–13:00 · ..."}},
  {"accion": "anadir_contacto", "id": "id-energy-group",
   "contacto": {"nombre": "Carlos Ruiz", "puesto": "Senior Mechanical Engineer", "telefono": "+34 682 515 212",
                "slot": "tecnico", "confirmado_feria": true, "interes": true}},
  {"accion": "anadir_proyecto", "id": "biorig", "proyecto": {"nombre": "Caspe", "ubicacion": "Zaragoza", "anio": "", "estado": ""}},
  {"accion": "nueva", "organizacion": { ...esquema completo del apartado 3... }}
]
```

## 8. Validación (validar.py debe pasar antes de generar)

- Ningún texto contiene los términos prohibidos del apartado 4 (expresión regular: `no localizad|no encontrad|sin verificar|fuente:`).
- Los correos generales o de departamento (`info@|contact@|web@|hablamos@|comercial@|sales@|service@` y equivalentes) se señalan como aviso en cada ejecución, sin detener la generación (regla 3).
- Estados de proyecto dentro de la lista admitida; como máximo 5 proyectos por organización.
- `id` y `nombre` únicos; grupo entre A y E.
- Tras generar: una ficha por página (comprobar en el PDF y en el Word convertido) y ninguna ficha desbordada.

## 9. Buenas prácticas en este repositorio

- No borrar organizaciones, contactos ni proyectos sin instrucción expresa.
- Commits pequeños y descriptivos, en español.
- Ante una instrucción ambigua, preguntar antes de asumir.

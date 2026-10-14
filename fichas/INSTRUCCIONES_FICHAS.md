# Fichas maestras: instrucciones para Cowork

El módulo mantiene un expediente por empresa (`expedientes.json`) y genera, para cada una, una ficha maestra en Word y en PDF con toda la información acumulada: situación actual, contactos, perfil y proyectos, con quién trabaja, lo que sabemos, tanques, cotizaciones, pendientes, competencia, cronología completa, documentos y una página final para tomar notas en la reunión.

El expediente solo guarda lo que no está en otra fuente:

| Dato | Fuente | Se guarda en el expediente |
|---|---|---|
| Perfil, contactos, proyectos y uso de GFS | `organizaciones.json` (repositorio) | No |
| Estado y fecha del próximo compromiso | Panel `Seguimiento_Clientes_Stora.xlsx` | No |
| Valoración | Tracker `Stora_Exhibition_Lead_Tracker.xlsx` | No |
| Fases | — | No |
| Nombre, tipo y contactos de las empresas que no están en `organizaciones.json` | Alta en el expediente (apartado 5) | Sí |
| Todo lo demás | Correos, reuniones, llamadas, notas, documentos | Sí |

**Los expedientes reales, las fichas generadas y las copias de seguridad contienen precios y condiciones internas: nunca se suben al repositorio** (el `.gitignore` los excluye). Las fichas son de uso interno y no se envían a terceros.

## 1. Archivos

En el ordenador de Andrés, carpeta `C:\Users\Vicente Prado\Desktop\ANDRES\SEGUIMIENTO DE CLIENTES - PANEL\FICHAS MAESTRAS\`:

```
FICHAS MAESTRAS\
  _motor_fichas\                   Copia del módulo (fichas/ del repositorio):
    fichas.py                        el programa
    esquema_expediente.json          esquema del expediente y de entradas.json
    plantilla\                       estilos del bloc y pie «Uso interno»
    requirements.txt                 dependencias de Python
    INSTRUCCIONES_FICHAS.md          este documento
  organizaciones_AAAA-MM-DD.json   Copia de datos/organizaciones.json del repositorio, con su fecha
  expedientes.json                 Expedientes de todas las empresas
  estado_fichas.json               Huellas de la última generación (lo crea «generar»)
  copias_seguridad\                Copias fechadas de expedientes.json (las crean «añadir» e «inicializar»)
  fichas\Ficha_<id>.docx / .pdf    Una ficha por empresa, con nombre fijo
  Indice_Fichas_Maestras.docx / .pdf
```

El panel y el tracker se leen desde donde Andrés los tenga; el módulo nunca los modifica.

La fecha que se imprime junto a los contactos («Copia de organizaciones.json del …») se toma del nombre del archivo (`organizaciones_2026-10-13.json`); si no la lleva, de su fecha de modificación. También puede indicarse con `--fecha-organizaciones AAAA-MM-DD`. Al copiar `datos/organizaciones.json` del repositorio, póngale la fecha de su último cambio.

## 2. Requisitos del contenedor

- Python 3.10 o posterior y las dependencias de `_motor_fichas/requirements.txt`: `openpyxl` y `pdfplumber`.
- LibreOffice sin interfaz (`soffice`).
- `pdfplumber` comprueba la paginación de cada PDF. Si faltara, el módulo genera igualmente las fichas, pero avisa de que no ha comprobado la paginación.

## 3. Procedimiento

Los comandos se ejecutan desde la carpeta `FICHAS MAESTRAS` copiada al contenedor. En los ejemplos, `ORG` es la copia de `organizaciones.json`, `PANEL` el panel y `TRACKER` el tracker.

1. **Traer los archivos** al contenedor: `_motor_fichas/`, la copia de `organizaciones.json`, `expedientes.json`, `estado_fichas.json` y la carpeta `fichas/` (sin ellos, se regeneran todas las fichas), el panel y el tracker.

2. **Solo la primera vez, o cuando haya organizaciones nuevas** en `organizaciones.json`: crear los expedientes que falten. Nunca sobrescribe los existentes.

   ```
   python _motor_fichas/fichas.py inicializar --organizaciones ORG --salida expedientes.json
   ```

3. **Preparar `entradas.json`** con lo nuevo (formato en los apartados 4 y 5). Use un nombre fechado, por ejemplo `entradas_2026-10-14.json`.

4. **Simular y aplicar las entradas.** Primero con `--simular` para revisar el informe; si no hay conflictos ni errores, sin él. Antes de escribir se guarda una copia de seguridad fechada en `copias_seguridad/`.

   ```
   python _motor_fichas/fichas.py añadir expedientes.json entradas_2026-10-14.json --organizaciones ORG --simular
   python _motor_fichas/fichas.py añadir expedientes.json entradas_2026-10-14.json --organizaciones ORG
   ```

   `anadir` (sin eñe) es equivalente. Pase siempre `--organizaciones`: permite identificar la empresa por su nombre en `organizaciones.json` y es obligatorio para dar de alta empresas.

5. **Generar** las fichas y el índice. Solo se regeneran las empresas cuyo contenido ha cambiado (expediente, datos de `organizaciones.json`, estado y próximo compromiso del panel o valoración del tracker), o cuya ficha falta en `fichas/`. Con `--todas` se regeneran todas. Si cambia el programa o la plantilla, se regeneran todas automáticamente.

   ```
   python _motor_fichas/fichas.py generar expedientes.json . --organizaciones ORG --panel PANEL --tracker TRACKER
   ```

6. **Validar**: esquema, id sin organización, nombres del panel y del tracker que no se asocian a ninguna empresa, valoraciones fuera de la escala de 1 a 5, empresas dadas de alta que ya aparecen en `organizaciones.json` y paginación de las fichas generadas.

   ```
   python _motor_fichas/fichas.py validar expedientes.json --organizaciones ORG --panel PANEL --tracker TRACKER
   ```

7. **Devolver** al ordenador de Andrés: `expedientes.json`, la nueva copia de `copias_seguridad/`, `estado_fichas.json`, las fichas regeneradas de `fichas/` y el índice. El panel y el tracker no se devuelven: no se han modificado.

Código de salida de todos los comandos: 0 si todo es correcto; 1 si hay errores; 2 (solo en `añadir`) si hay conflictos y ningún error.

## 4. Formato de `entradas.json`

Es una lista de bloques, uno por empresa. Cada bloque lleva `empresa` (o `alta` para una empresa nueva, apartado 5) y solo las listas que traen algo nuevo. **Las entradas nuevas no llevan `id`**: lo asigna el script (`cro-0001`, `pen-0003`…) y no cambia nunca. Las fechas van en formato ISO (`AAAA-MM-DD`) y las horas como `HH:MM`. Lo que no conste se omite: no se escriben «sin datos», «N/D» ni marcadores similares.

`empresa` admite el id (preferible), el nombre de `organizaciones.json` o de una empresa dada de alta, o un nombre alternativo ya registrado (por ejemplo, «Renove Group» para `renovebio` o «Sitra» para `sitra`). Se ignoran las mayúsculas, las tildes, «S.L.», «S.A.» y el texto entre paréntesis.

```json
[
  {
    "empresa": "biometano-meseta-ficticia",
    "nombres_alternativos": ["Meseta Bio"],
    "situacion_actual": {
      "texto": "Tres a cinco líneas con la situación de la relación a esta fecha.",
      "fecha": "2026-10-14"
    },
    "lo_que_sabemos": [
      {"texto": "La planta tratará 60.000 t/año.", "naturaleza": "dato",
       "fuente": "Reunión del 30/09 con Lucía Ejemplo", "fecha": "2026-09-30"},
      {"texto": "Presupuesto de tanques en torno a un millón de euros.", "naturaleza": "estimacion",
       "fuente": "Llamada del 05/10", "fecha": "2026-10-05"}
    ],
    "proyectos": [
      {"nombre": "Planta Norte", "ubicacion": "Burgos", "situacion": "Licencia ambiental concedida.",
       "fuente": "Reunión del 30/09", "fecha": "2026-09-30"}
    ],
    "relaciones": [
      {"organizacion": "Fondo Inversor Ficticio", "papel": "accionista", "texto": "Participa en las dos plantas.",
       "fuente": "Notas de Andrés", "fecha": "2026-10-05"}
    ],
    "tanques": [
      {"proyecto": "Planta Norte", "uso": "Digestor", "cantidad": 2, "diametro_m": 24.36, "altura_m": 12.6,
       "volumen_m3": 5870, "techo": "Doble membrana", "aislamiento": "Lana de roca 100 mm", "otros": "",
       "fuente": "Correo del 02/10", "fecha": "2026-10-02"}
    ],
    "cotizaciones": [
      {"fecha": "2026-10-08", "descripcion": "Dos digestores de 5.870 m³", "estado": "Enviada",
       "importe_eur": 412500, "fuente": "Oferta de Stora del 08/10"}
    ],
    "competencia": [
      {"texto": "Comparan con tanques de hormigón; piden argumentos de plazo.",
       "fuente": "Reunión del 30/09", "fecha": "2026-09-30"}
    ],
    "pendientes": [
      {"accion": "Enviar referencias de digestores", "de_quien": "nosotros", "fecha_limite": "2026-10-16",
       "estado": "abierto"}
    ],
    "cerrar_pendientes": [
      {"id": "pen-0002", "fecha_cierre": "2026-10-14"},
      {"accion": "Enviar planos de implantación", "fecha_cierre": "2026-10-14"}
    ],
    "cronologia": [
      {"fecha": "2026-10-02", "hora": "09:14", "tipo": "correo_recibido", "participantes": ["Lucía Ejemplo"],
       "resumen": "Envía las dimensiones de los digestores.", "datos_clave": ["Diámetro 24,36 m"],
       "acuerdos": [], "origen": "automatico", "fuente": {"tipo": "gmail", "ref": "19a1b2c3d4e5f601"}},
      {"fecha": "2026-09-30", "hora": "12:30", "tipo": "reunion", "participantes": ["Lucía Ejemplo", "Andrés Prado"],
       "resumen": "Reunión en el stand.", "acuerdos": ["Enviar oferta en dos semanas"], "origen": "manual",
       "fuente": {"tipo": "foto_nota", "ref": "Foto de la libreta, 30/09"}}
    ],
    "documentos": [
      {"fecha": "2026-10-08", "nombre": "Oferta digestores Planta Norte", "ruta": "DOCUMENTOS\\Meseta\\Oferta_2026-10-08.pdf"}
    ]
  }
]
```

Campos (obligatorios en negrita):

| Lista | Campos | Valores |
|---|---|---|
| `situacion_actual` | **texto**, **fecha** | De 3 a 5 líneas. Sustituye a la anterior si su fecha es igual o posterior. |
| `lo_que_sabemos` | **texto**, **naturaleza**, **fuente**, **fecha** | `naturaleza`: `dato` (verificado), `estimacion` o `supuesto`. |
| `proyectos` | **nombre**, ubicacion, situacion, **fuente**, **fecha** | Proyectos que no están en `organizaciones.json` o su situación actualizada. |
| `relaciones` | **organizacion**, **papel**, texto, **fuente**, **fecha** | Con quién trabaja la empresa. `papel`: `socio`, `cliente`, `EPC`, `ingeniería`, `tecnólogo`, `proveedor`, `accionista` u `otro`. |
| `tanques` | proyecto, uso, cantidad, diametro_m, altura_m, volumen_m3, techo, aislamiento, otros, **fuente**, **fecha** | Al menos un dato técnico. Medidas en metros y m³, como números. |
| `cotizaciones` | **fecha**, **descripcion**, **estado**, importe_eur, **fuente** | `importe_eur` es un número (sin «€»). |
| `competencia` | **texto**, **fuente**, **fecha** | Competidores y objeciones. |
| `pendientes` | **accion**, **de_quien**, fecha_limite, **estado**, fecha_cierre | `de_quien`: `nosotros`, `Stora` o `cliente`; `estado`: `abierto` o `cerrado` (con `fecha_cierre`). |
| `cerrar_pendientes` | id o accion, **fecha_cierre** | Cierra un pendiente abierto, por su id o por el texto de la acción (sin distinguir mayúsculas ni tildes). |
| `cronologia` | **fecha**, hora, **tipo**, participantes, **resumen**, datos_clave, acuerdos, **origen**, **fuente** | `tipo`: `correo_recibido`, `correo_enviado`, `reunion`, `llamada`, `nota`, `whatsapp` o `documento`. `origen`: `automatico` (leído por Cowork) o `manual` (dictado o escrito por Andrés). `fuente`: `{"tipo": "gmail" \| "conversacion" \| "foto_nota" \| "documento", "ref": "…"}`; en los correos, `ref` es el identificador del mensaje de Gmail y es obligatorio. |
| `documentos` | **fecha**, **nombre**, **ruta** | Ruta en el ordenador de Andrés. |
| `nombres_alternativos` | lista de textos | Nombres con que aparece la empresa en el panel, el tracker o los correos. |
| `contactos` | **nombre**, cargo, correo, telefono | Solo en las empresas dadas de alta (apartado 5). |

`fichas/pruebas/entradas_prueba.json` es un ejemplo completo con empresas inventadas.

## 5. Empresas que no están en `organizaciones.json`

Las empresas del mapa maestro y los leads que no están en `organizaciones.json` se dan de alta en el expediente con un bloque `alta` en lugar de `empresa`. El alta crea el expediente con `nombre`, `tipo` (A Promotoras, B EPC, C Ingenierías, D Consultorías, E Cooperativas y asociaciones, como `grupo` en `organizaciones.json`), `contactos` y `fuera_de_organizaciones: true`. El id se forma a partir del nombre con la misma regla del repositorio («Depósitos Imaginarios del Levante» → `depositos-imaginarios-del-levante`). El resto del bloque (nombres alternativos, cronología, etc.) se aplica a la empresa recién creada.

```json
[
  {
    "alta": {
      "nombre": "Depósitos Imaginarios del Levante",
      "tipo": "B",
      "contactos": [
        {"nombre": "Elena Ficticia Mora", "cargo": "Directora Técnica",
         "correo": "elena.ficticia@depositos-imaginarios.example", "telefono": "+34 600 000 601"}
      ]
    },
    "nombres_alternativos": ["DIL Tanques"],
    "cronologia": [
      {"fecha": "2026-10-11", "tipo": "correo_recibido", "resumen": "Piden precio orientativo de dos tanques.",
       "origen": "automatico", "fuente": {"tipo": "gmail", "ref": "19a1b2c3d4e5f604"}}
    ]
  },
  {
    "empresa": "depositos-imaginarios-del-levante",
    "contactos": [
      {"nombre": "Jorge Supuesto Vidal", "cargo": "Compras", "correo": "jorge.supuesto@depositos-imaginarios.example"}
    ]
  }
]
```

- Es un conflicto que el id, el nombre o un nombre alternativo del alta coincidan con una organización de `organizaciones.json` o con otro expediente. Repetir la misma alta (mismo nombre y tipo) no es conflicto: se omite.
- También es conflicto, para no crear duplicados, que el nombre esté contenido en el de otra empresa o la contenga, palabra a palabra («JFA Technologies» y «JFA Technologies / TECON»). Si Andrés confirma que es otra empresa, se añade `"no_es_duplicado": true` al bloque del alta; el informe indica siempre qué bloques llevan esa marca.
- Los contactos posteriores se añaden con un bloque `empresa` y la lista `contactos`. Nunca se borran ni se modifican. Un contacto está duplicado si coincide el correo o el nombre normalizado. Las empresas de `organizaciones.json` no admiten `contactos` en el expediente: los suyos se mantienen allí, por el flujo del bloc.
- La ficha toma los contactos del expediente. El perfil, los proyectos de `organizaciones.json` y el uso de GFS no se imprimen porque no hay datos.
- Si una empresa dada de alta aparece después en `organizaciones.json` (por id, nombre o nombre alternativo), `validar` e `inicializar` lo avisan («hay que migrar su expediente») e `inicializar` no crea un segundo expediente para ella. No lo resuelva por su cuenta: avise a Andrés; la migración requiere su instrucción.

## 6. Duplicados, conflictos y errores

`añadir` solo añade. No modifica ni borra entradas existentes, salvo tres excepciones: cerrar pendientes, reescribir la situación actual y añadir nombres alternativos. **Si hay algún conflicto o error, no se aplica nada** (ni siquiera las entradas correctas) y el informe lista todas las incidencias.

Una entrada duplicada no se añade y se lista en el informe como `OMITIDA`, sin detener el proceso. Si trae otros valores en campos que no son de la comparación, el informe lo indica («Difiere en …; la entrada existente no se modifica»). Se comparan sin distinguir mayúsculas, tildes ni puntuación:

| Lista | Es duplicada si coincide |
|---|---|
| `cronologia` | La referencia de Gmail, o la fecha, el tipo y el resumen. |
| `lo_que_sabemos`, `competencia` | El texto. |
| `proyectos` | El nombre y la ubicación. |
| `relaciones` | La organización y el papel. |
| `tanques` | El proyecto, el uso, la cantidad, el diámetro y la altura. |
| `cotizaciones` | La fecha y la descripción. |
| `pendientes` | La acción, mientras siga abierto otro pendiente igual. Si el anterior está cerrado, el nuevo se admite. |
| `documentos` | La ruta. |
| `contactos` (altas) | El correo o el nombre. |

| Informe | Causa | Qué hacer |
|---|---|---|
| `OMITIDA` | Duplicado (tabla anterior), alta repetida o pendiente que ya estaba cerrado con la misma fecha. | Nada. Si el informe indica diferencias, muéstreselas a Andrés: la entrada existente no se ha modificado. Un pendiente nuevo con `estado: cerrado` igual a uno abierto se omite: para cerrar el abierto, use `cerrar_pendientes`. |
| `CONFLICTO` | La empresa no se encuentra ni por id ni por nombre, o el nombre corresponde a varias. | Compruebe el id. Si es un nombre nuevo de una empresa existente, identifíquela por su id y añada el nombre en `nombres_alternativos`. Si la empresa no está en `organizaciones.json` ni en el expediente, dela de alta (apartado 5) o, si Andrés lo indica, añádala a `organizaciones.json` por el flujo del bloc y ejecute `inicializar`. |
| `CONFLICTO` | Intento de modificar una entrada existente: una entrada con el `id` de otra y contenido distinto. | No lo resuelva por su cuenta: muéstrelo a Andrés. Un cambio se registra como entrada nueva con su fecha (una cotización aceptada es una nueva cotización con la fecha de la aceptación). Corregir una entrada existente requiere instrucción expresa de Andrés: se edita a mano `expedientes.json` (tras una copia de seguridad) y se comprueba con `validar`. |
| `CONFLICTO` | Alta que coincide con una organización o con otro expediente, o con un nombre parecido. | Muéstrelo a Andrés. Si es la misma empresa, use un bloque `empresa` con su id. Si, con un nombre parecido, Andrés confirma que es otra empresa, añada `"no_es_duplicado": true` al bloque del alta. |
| `CONFLICTO` | Cerrar un pendiente ya cerrado con otra fecha, una situación actual más antigua que la registrada o un nombre alternativo que ya identifica a otra empresa. | Muéstrelo a Andrés. |
| `ERROR` | Formato: campo desconocido u obligatorio que falta, fecha no ISO, valor no admitido, `id` inventado, pendiente inexistente, correo de Gmail sin `ref`, tanque sin datos, bloque sin `empresa` ni `alta` (o con los dos), alta sin `--organizaciones`, contactos de una empresa de `organizaciones.json`. | Corrija `entradas.json` y repita. |
| `AVISO` | Situación actual demasiado larga. | Resúmala si procede; no detiene el proceso. |

Tras resolver los conflictos, corrija `entradas.json` (quite las entradas que no proceden) y repita `añadir`. Las entradas ya aplicadas en una ejecución anterior saldrán como `OMITIDA`.

## 7. Panel y tracker

- **Panel** `Seguimiento_Clientes_Stora.xlsx`: hojas PROMOTORES, EPC, INGENIERÍAS Y CONSULTORÍAS y OTROS; columnas Empresa, Estado, Fecha (próximo compromiso), Contacto, Teléfono, Email y Notas. La ficha usa Estado y Fecha.
- **Tracker** `Stora_Exhibition_Lead_Tracker.xlsx`: hoja «Lead Overview», columnas Rating, Company, Exhibition / Event, Contact, Opportunity y Stage. La ficha usa Rating como valoración; si la empresa tiene varias filas, la mejor.
- **Valoración**: número entero del 1 al 5; 5 es la máxima. En los leads nuevos puede estar vacía: entonces la ficha no la imprime y la empresa va al final del índice. Por tolerancia se ordenan también estrellas (`★★★`) y las palabras Hot/Warm/Cold, Alta/Media/Baja y A/B/C; `validar` avisa de cualquier valor fuera de la escala de 1 a 5.
- La fila de cabecera se busca por su texto en las 30 primeras filas de cada hoja (un título encima no molesta). Una columna cuyo título empieza por el nombre esperado también vale («Fecha (próximo compromiso)»).
- Cada nombre se asocia a una empresa primero por id, después por el nombre (de `organizaciones.json` o de una empresa dada de alta) y por último por los nombres alternativos, sin distinguir mayúsculas, tildes, «S.L.», «S.A.» ni el texto entre paréntesis.
- `validar` lista todos los nombres que no se asocian a ninguna empresa (o que corresponden a varias). Para resolverlos, añada el nombre en `nombres_alternativos` del bloque de la empresa correcta, identificada por su id, o dé de alta la empresa si no existe.

## 8. Qué contiene cada ficha

Una ficha por empresa (A4 vertical, estilo del bloc, pie con «Uso interno» y número de página), con este orden: cabecera (nombre, tipo, valoración, estado, próximo compromiso y fecha de generación); situación actual; contactos (de `organizaciones.json`, con la fecha de la copia, o del expediente en las empresas dadas de alta); perfil, proyectos y uso de GFS de `organizaciones.json`, con quién trabaja (lista `relaciones`) y los proyectos del expediente; lo que sabemos, con la naturaleza de cada punto; tanques; cotizaciones; pendientes abiertos; competencia y objeciones; cronología completa, de la más reciente a la más antigua, con las entradas automáticas marcadas; documentos y fuentes; y, en página propia, la hoja de notas de reunión.

Las secciones sin datos no se imprimen.

El índice `Indice_Fichas_Maestras` lista empresa (con enlace a su PDF en `fichas/`), tipo, valoración, estado, última interacción (fecha y tipo) y número de pendientes abiertos, ordenado por valoración (de 5 a 1; sin valoración, al final) y, después, por la fecha de la última interacción.

## 9. Pruebas

Solo con empresas inventadas (`fichas/pruebas/`). Desde la raíz del repositorio:

```
python -m pytest fichas/pruebas -q
```

`python fichas/pruebas/datos_prueba.py CARPETA` crea un panel y un tracker inventados con la estructura real. `fichas/pruebas/ejemplo/` contiene dos fichas de ejemplo de empresas inventadas: una de `organizaciones.json` y otra dada de alta fuera de él.

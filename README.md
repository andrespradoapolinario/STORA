# Sistema comercial Stora

Documentos comerciales de Alliance Albors y Prado S.L. para Stora. Las reglas de contenido y formato están en [CLAUDE.md](CLAUDE.md).

Todo documento se genera desde una única fuente de datos: `datos/organizaciones.json`.

## Comandos

Se ejecutan desde la raíz del repositorio, en este orden:

1. **Aplicar entradas**: aplica los parches que Cowork deja en `datos/entrada/` y los mueve a `datos/entrada/aplicadas/`.

   ```
   python scripts/aplicar_entradas.py            # --simular para ver los cambios sin escribir nada
   ```

   El informe lista todos los conflictos (un dato ya existente con otro valor) y errores del parche, no solo el primero, e indica siempre qué operaciones llevan `"no_es_duplicado"`. Si hay alguno, no se aplica nada de ese parche ni de los siguientes. Si Andrés confirma el valor nuevo, se añade `"sustituir": true` a esa operación del parche, con solo los campos que cambian, y se vuelve a ejecutar. Para completar el nombre de un contacto existente sin duplicarlo, el contacto del parche lleva `"nombre_anterior"` (apartado 7 de CLAUDE.md).

   También son conflicto un contacto nuevo que se parece a uno existente de la misma organización (mismo correo, o un nombre contenido en el otro, como «Juanjo Linares» y «Juanjo Linares Salinas») y una organización «nueva» con un contacto cuyo correo ya figura en otra organización. Si Andrés confirma que no es un duplicado, se añade `"no_es_duplicado": true` a esa operación. Un valor vacío o `false` del parche nunca borra un dato existente, aunque la operación lleve `"sustituir": true`, y los cambios que solo afectan a mayúsculas no se aplican.

2. **Validar**: comprueba los datos con las reglas del apartado 8 de CLAUDE.md. Los errores detienen la generación; los avisos solo señalan datos que conviene revisar.

   ```
   python scripts/validar.py                     # --pdf salidas/ARCHIVO.pdf comprueba además la paginación
   ```

3. **Generar**: valida y genera el Word y el PDF en `salidas/`, con la fecha del día en el nombre (`--fecha AAAA-MM-DD` para otra).

   ```
   python scripts/generar.py bloc
   ```

   El PDF es el Word convertido con LibreOffice, así que ambos tienen el mismo contenido. Las líneas de NOTAS de cada ficha se ajustan para llenar la página sin desbordarla, y al final se comprueba que haya una ficha por página.

## Requisitos

- Python 3.10 o posterior y `pip install -r requirements.txt` (`openpyxl` solo lo usan las fichas maestras). Para las pruebas, además, `pip install pytest`.
- LibreOffice con Writer (para el PDF; en Debian o Ubuntu, paquete `libreoffice-writer`: sin él, la conversión falla con «source file could not be loaded»).
- Fuente Calibri o, en su defecto, Carlito (equivalente métrico libre), para que el PDF reproduzca el Word.

## Estructura

```
datos/organizaciones.json      Fuente única de datos (esquema en el apartado 3 de CLAUDE.md).
datos/entrada/                 Parches pendientes preparados por Cowork.
datos/entrada/aplicadas/       Parches ya aplicados (historial; no se borran).
scripts/aplicar_entradas.py    Aplica los parches.
scripts/validar.py             Validación de datos y de paginación.
scripts/generar.py             Punto de entrada de la generación.
scripts/bloc_docx.py           Generador del bloc en Word.
scripts/bloc_pdf.py            Generador del PDF y ajuste de las NOTAS.
scripts/comun.py               Rutas y constantes del esquema.
scripts/plantilla_bloc/        Partes fijas del Word (estilos, pie, tema).
salidas/                       Documentos generados. No se editan a mano.
fichas/                        Módulo de fichas maestras (expedientes por empresa).
```

## Fichas maestras

`fichas/fichas.py` mantiene un expediente por empresa y genera una ficha maestra en Word y PDF para cada una, más un índice. Los expedientes reales no entran en el repositorio. Uso, formato de entradas y conflictos: [fichas/INSTRUCCIONES_FICHAS.md](fichas/INSTRUCCIONES_FICHAS.md); reglas: apartado 10 de CLAUDE.md.

```
python -m pytest fichas/pruebas -q      # pruebas, solo con empresas inventadas
```

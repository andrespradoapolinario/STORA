# Sistema comercial Stora

Documentos comerciales de Alliance Albors y Prado S.L. para Stora. Las reglas de contenido y formato están en [CLAUDE.md](CLAUDE.md).

Todo documento se genera desde una única fuente de datos: `datos/organizaciones.json`.

## Comandos

Se ejecutan desde la raíz del repositorio, en este orden:

1. **Aplicar entradas**: aplica los parches que Cowork deja en `datos/entrada/` y los mueve a `datos/entrada/aplicadas/`.

   ```
   python scripts/aplicar_entradas.py            # --simular para ver los cambios sin escribir nada
   ```

   Se detiene ante cualquier conflicto (un dato ya existente con otro valor) sin aplicar ese parche. Si Andrés confirma el valor nuevo, se añade `"sustituir": true` a esa operación del parche y se vuelve a ejecutar.

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

- Python 3.10 o posterior y `pip install -r requirements.txt`.
- LibreOffice (para el PDF).
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
```

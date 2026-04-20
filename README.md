# Teamwork Projects — importación Excel

App web en Python/FastAPI que genera un archivo **Excel (.xlsx)** con el mismo esquema que los ejemplos oficiales `Teamwork-Projects-Import-Sample_*.xlsx`, listo para importar el plan en **Teamwork Projects** (etapas y actividades).

## Estructura del proyecto

```
ESTIMADOR/
├── main.py          # Backend FastAPI + generación Excel (openpyxl)
├── index.html       # Frontend (servido por FastAPI)
├── requirements.txt
└── README.md
```

## Instalación

```bash
# 1. Crear entorno virtual (recomendado)
python -m venv venv
source venv/bin/activate  # Windows: venv\Scripts\activate

# 2. Instalar dependencias
pip install -r requirements.txt

# 3. Correr la app
uvicorn main:app --reload
```

## Uso

1. Abrí http://localhost:8000 en tu navegador
2. Completá los datos del proyecto (nombre, cliente, fecha de inicio)
3. Al pasar de **Datos del proyecto** a **Etapas**, se cargan solas las etapas **FASE B — Proyecto**. El panel solo muestra **Desarrollo** para editar. Podés **importar** el Excel de estimación de desarrollo RPA (`Copia de Estimación Desarrollo_*.xlsx`, tabla *Detalle de la estimación*). El modelo v4 se ejecuta **al importar** y otra vez **al ir a Revisión**. Integración bot: `POST /import-desarrollo-xlsx` con `multipart/form-data`, campo `file` (`.xlsx`); respuesta con `subtareas` y `metadata` opcional.
4. La suma **H** en Desarrollo es **solo DEV**, fija en el detalle. Los % BA/LT/QA/PM del panel se aplican sobre **H×factor**; factores/buffers marcan calendario. Las demás etapas se completan en el import y al pasar a Revisión; el detalle de Desarrollo no se redistribuye.
5. Ajustá responsables y estado donde corresponda
6. En **Revisión**, revisá la tabla previa y descargá el **Excel**
7. En Teamwork: **Import** → subí el `.xlsx` según el flujo de importación de proyectos/tareas planificadas

## Formato del Excel

Hoja **Sheet1**, columnas (fila de encabezado):

`TASKLIST`, `TASK`, `DESCRIPTION`, `ASSIGN TO`, `START DATE`, `DUE DATE`, `PRIORITY`, `ESTIMATED TIME`, `TAGS`, `STATUS`

- **TASKLIST** = nombre del **cliente** en todas las filas.
- **TASK**: la primera fila de datos usa el **nombre del proyecto** sin guiones. Debajo, cada **etapa** del formulario es de primer nivel con **un** guión (`-Preventa`, `-DESARROLLO`, …). Las actividades bajo la etapa llevan **dos** guiones (`--…`), y cada nivel anidado suma un guión (`---…`, etc.). Convención alineada al ciclo de vida RPA del portal VORTEX (`rpa-lifecycle.html`).
- Los nombres se guardan **sin** guiones iniciales en el formulario; el Excel los agrega según el nivel.
- **TAGS** = nombre del **proyecto** (referencia).
- **Estimado por** va en la descripción de la fila del **nombre del proyecto**.
- **Estados** del formulario → Teamwork: `DONE` → `Completed`, `Bloqueado` → `Deferred`, el resto → `Active`.
- **Fechas**: a partir de la fecha de inicio del proyecto, cada fila **hoja** (sin hijos) con horas consume días en bloques de 8 h (mínimo 1 día). Las filas contenedor (con subactividades) no avanzan el calendario.

## Campos capturados

### Por proyecto

- Nombre del proyecto
- Cliente
- Fecha de inicio
- Estimado por

### Por etapa / actividad (y subniveles)

- Nombre
- Horas estimadas
- Responsable
- Estado (Pendiente / En curso / DONE / Bloqueado)
- Observaciones

## API (opcional)

- `POST /preview-excel-data` — cuerpo JSON igual que el modelo de proyecto; devuelve filas para la vista previa.
- `POST /generate-xlsx` — mismo cuerpo; devuelve el archivo `.xlsx` para descargar.

# Documentación del Script de Automatización para Registros de Monitorías

## Descripción General

Este script automatiza el proceso de lectura, filtrado y registro de monitorías académicas a partir de un archivo Excel, generando un resumen semanal en un documento de Word y actualizando una bitácora de monitorías en un archivo Excel. La principal funcionalidad es filtrar las monitorías que ocurren dentro de la semana actual y generar un informe detallado para su posterior análisis.

El script consta de los siguientes pasos:
1. Cargar los datos desde un archivo Excel.
2. Filtrar los registros de monitorías para la semana actual.
3. Crear un resumen semanal en un archivo de Word.
4. Actualizar una bitácora en Excel con todas las monitorías procesadas.

## Estructura del Script

### 1. Importaciones de Módulos

```python
import openpyxl
from openpyxl import load_workbook
from datetime import datetime, timedelta
from docx import Document
import os
```

Este bloque importa las bibliotecas necesarias:
- `openpyxl`: Para manejar archivos Excel.
- `datetime`: Para manipular fechas y tiempos.
- `docx`: Para crear y modificar archivos de Word.
- `os`: Para verificar la existencia de archivos y manejar rutas del sistema operativo.

### 2. Función `load_data_from_excel`

```python
def load_data_from_excel(file_path):
    workbook = load_workbook(filename=file_path)
    sheet = workbook.active

    headers = [cell.value.strip() for cell in sheet[1]]
    data = []
    for row in sheet.iter_rows(min_row=2, values_only=True):
        data.append(dict(zip(headers, row)))

    return data
```

#### **Descripción**:
Esta función se encarga de cargar datos desde un archivo Excel (`.xlsx`). Extrae los encabezados (primera fila del archivo) y los utiliza como claves para cada fila de datos que se almacena en un diccionario.

#### **Funcionamiento**:
1. Abre el archivo Excel utilizando `openpyxl`.
2. Lee la primera fila para obtener los encabezados y limpia los espacios en blanco.
3. Lee las filas restantes y las convierte en una lista de diccionarios donde cada clave es el nombre de una columna y cada valor es el dato de la celda correspondiente.

### 3. Función `filter_monitorias_week`

```python
def filter_monitorias_week(data):
    today = datetime.now()
    start_week = today - timedelta(days=today.weekday())
    end_week = start_week + timedelta(days=6)

    monitorias_week = []
    for entry in data:
        if "Fecha de la monitoría" in entry:
            fecha_monitoria = entry["Fecha de la monitoría"]

            if isinstance(fecha_monitoria, str):
                try:
                    fecha_monitoria = datetime.strptime(fecha_monitoria, "%d/%m/%Y")
                except ValueError:
                    print(f"Formato de fecha incorrecto en: {entry}")
                    continue

            if isinstance(fecha_monitoria, datetime) and start_week <= fecha_monitoria <= end_week:
                monitorias_week.append(entry)

    return monitorias_week
```

#### **Descripción**:
Filtra las monitorías que ocurrieron en la semana actual. La semana comienza el lunes y termina el domingo.

#### **Funcionamiento**:
1. Calcula las fechas de inicio y fin de la semana actual.
2. Recorre cada entrada en los datos cargados y convierte la fecha de la monitoría en un objeto `datetime`.
3. Si la fecha está dentro de la semana actual, se añade a la lista `monitorias_week`.

### 4. Función `create_weekly_summary`

```python
def create_weekly_summary(monitorias_week):
    document = Document()
    document.add_heading(f"Resumen Semanal de Monitorías", 0)

    today = datetime.now()
    week_range = f"Semana del {today - timedelta(days=today.weekday()):%d/%m/%Y} al {(today + timedelta(days=6 - today.weekday())):%d/%m/%Y}"
    document.add_paragraph(f"Período: {week_range}")

    if monitorias_week:
        for monitoria in monitorias_week:
            document.add_paragraph("----------")
            document.add_paragraph(f"Estudiante: {monitoria['Nombre completo del estudiante']}")
            document.add_paragraph(f"Código: {monitoria['Código del estudiante']}")
            document.add_paragraph(f"Correo: {monitoria['Correo']}")
            document.add_paragraph(f"Grupo Académico: {monitoria['Grupo académico']}")
            document.add_paragraph(f"Jornada: {monitoria['Jornada de estudios']}")
            document.add_paragraph(f"Tipo de Monitoría: {monitoria['Tipo de monitoría recibida']}")
            document.add_paragraph(f"Fecha de la Monitoría: {monitoria['Fecha de la monitoría']}")
            document.add_paragraph(f"Horario: {monitoria['Horario de la monitoría']}")
            document.add_paragraph(f"Modalidad: {monitoria['Modalidad de la monitoría']}")
            document.add_paragraph(f"Comentarios: {monitoria['Comentarios adicionales']}")
            document.add_paragraph("----------")

    else:
        document.add_paragraph("No hubieron agendaciones esta semana.")
        document.add_paragraph("Se trabajó en la creación de materiales de apoyo para los estudiantes.")
        document.add_paragraph("----------")

    return document
```

#### **Descripción**:
Crea un documento de Word con el resumen de las monitorías realizadas durante la semana actual.

#### **Funcionamiento**:
1. Crea un nuevo documento de Word con un encabezado.
2. Añade información detallada sobre cada monitoría filtrada en la semana actual.
3. Si no hay monitorías, añade un mensaje indicando que no se realizaron agendaciones.

### 5. Función `save_summary`

```python
def save_summary(document, save_path):
    document.save(save_path)
    print(f"Resumen semanal guardado en: {save_path}")
```

#### **Descripción**:
Guarda el resumen semanal en un archivo Word en la ubicación especificada.

### 6. Función `save_log`

```python
def save_log(data, log_file):
    if os.path.exists(log_file):
        workbook = load_workbook(log_file)
        sheet = workbook.active
    else:
        workbook = openpyxl.Workbook()
        sheet = workbook.active
        sheet.append(list(data[0].keys()))

    for entry in data:
        sheet.append(list(entry.values()))

    workbook.save(log_file)
    print(f"Bitácora actualizada en: {log_file}")
```

#### **Descripción**:
Actualiza la bitácora de monitorías guardando todas las entradas, ya sea en un archivo existente o creando uno nuevo.

### 7. Ejecución del Script

```python
excel_file_path = "Datos-excel/Formulario-FPI.xlsx"
n = 0
summary_file_path = f"resumen_semanal{n+1}.docx"
log_file_path = "bitacora_monitorias.xlsx"

data = load_data_from_excel(excel_file_path)
monitorias_week = filter_monitorias_week(data)

document = create_weekly_summary(monitorias_week)
save_summary(document, summary_file_path)

save_log(data, log_file_path)
```

Este bloque principal ejecuta las funciones descritas para cargar, filtrar, generar el resumen y actualizar la bitácora.

## Conclusión

Este script es una solución automatizada para gestionar las monitorías académicas, ofreciendo una forma sencilla de generar informes semanales y mantener un registro histórico. Está diseñado para ser reutilizable y fácilmente adaptable a otros contextos donde sea necesario manejar grandes cantidades de datos a partir de archivos Excel.
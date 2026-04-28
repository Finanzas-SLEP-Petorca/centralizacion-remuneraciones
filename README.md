# Centralización de Remuneraciones — SLEP Petorca

Sistema de centralización contable de remuneraciones para el Servicio Local de Educación Pública de Petorca. Procesa los insumos de **CAS Chile** y genera los asientos contables para **SIGFE**.

## Requisitos

- Python 3.10 o superior
- Dependencias: `openpyxl`, `xlrd`, `pypdf`

```bash
pip install openpyxl xlrd pypdf
```

## Cómo usar

### Interfaz gráfica (recomendado)

Ejecutar haciendo doble clic en `iniciar_app_centralizacion.bat`  
o desde terminal:

```bash
python centralizacion_app.py
```

### Pasos en la interfaz

1. Seleccionar **Área** (Educación o Jardines) y **Período** (mes/año)
2. Cargar los **Maestros de Remuneraciones** (uno o más archivos `.xls` de CAS Chile)
3. Cargar los **Informes de Gasto/Financiamiento** (uno o más archivos `.xls` de CAS Chile)
4. (Opcional) Cargar asientos de referencia o centralizaciones anteriores de CAS Chile
5. Seleccionar el archivo de **Mapeo Maestro**: `mapeo_maestro.xlsx`
6. Indicar la carpeta de salida
7. Hacer clic en **Procesar**

### Archivos de entrada (exportar desde CAS Chile)

| Archivo | Descripción |
|---|---|
| Maestro de Remuneraciones | Contiene hojas `Haberes` y `Descuentos` |
| InformeGastoFinanciamiento | Detalle por fuente de financiamiento |
| InformeCentralizacion (opcional) | Centralización existente en CAS para verificación |
| Asiento Remuneración (opcional) | Asiento de referencia |

### Mapeo de cuentas contables

El archivo `mapeo_maestro.xlsx` contiene el mapeo fijo entre los códigos de CAS Chile y las cuentas contables SIGFE. Este archivo **no se sobreescribe** con cada proceso; solo se agregan filas nuevas para códigos que no existen aún.

Para completar una cuenta nueva:
1. Abrir `mapeo_maestro.xlsx`
2. Ir a la hoja `CompletarAqui`
3. Buscar filas con estado `COMPLETAR`
4. Ingresar la `cuenta_contable` correcta y cambiar el estado a `OK`

## Archivos generados

Los archivos se guardan en la carpeta `salidas/` (no versionada en git):

| Archivo | Descripción |
|---|---|
| `centralizacion_[area]_[periodo].xlsx` | Reporte principal con 10 hojas |
| `mapeo_cuentas_[periodo].csv` | Mapeo en formato CSV |
| `mapeo_cuentas_[periodo].xlsx` | Mapeo editable |

## Estructura del proyecto

```
├── centralizacion_app.py           # Interfaz gráfica (Tkinter)
├── centralizacion_remuneraciones.py # Motor de procesamiento
├── mapeo_maestro.xlsx               # Mapeo fijo de cuentas contables
├── iniciar_app_centralizacion.bat   # Lanzador Windows
└── salidas/                         # Resultados (no versionados)
```

## Flujo de datos

```
CAS Chile (XLS) → Maestros + Gastos
        ↓
  Agregación por código/tipo de cargo
        ↓
  Aplicar mapeo_maestro.xlsx
        ↓
  Excel con asiento propuesto para SIGFE
```

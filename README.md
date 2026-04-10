## Procesamiento ESG de ETFs

Construir una base final con el ESG por ETF usando:
- **Promedio simple** de `ESG Score`
- **Promedio ponderado** por `Company Market Capitalization`

El término ESG bursátil (Environmental, Social, and Governance) se refiere a la integración de criterios ambientales, sociales y
de gobierno corporativo en la evaluación y selección de empresas cotizadas en bolsa.
---

## Estructura

```text
procesar_etf/
  main.py
  README.md
  input/
    ETFS/
      *.xlsx
    data_maria_esg.xlsx
    data_maria_market_cap.xlsx
    Base de données.xlsm
  output/
    ETFS/
      {ticker}_{year}.xlsx (2 hojas: Procesados y Descartados)
    resultado_esg_etf.xlsx
    trazabilidad_procesamiento.xlsx
    etfs_sin_metadatos.xlsx
    Base_de_donnees.xlsx
```

---

## Entradas

### 1) `input/ETFS/*.xlsx`

- Encabezados de tabla en **fila 4** (`header=3` en pandas).
- **Ignorar archivos temporales** que empiecen con `~$`.
- **Metadatos** (búsqueda dinámica en primeras 5 filas, 3 columnas):
  
  **IMPORTANTE**: Los metadatos NO siempre están en las mismas posiciones fijas. Algunos archivos ETF tienen filas adicionales al inicio. El sistema implementa una **búsqueda dinámica** en las primeras 5 filas del archivo y 3 primeras columnas para encontrar:
  
  - `etf_ticker`: 
    - Buscar en la **primera columna, primeras 5 filas**
    - Criterios: 2-10 caracteres, contiene letras mayúsculas
    - Puede incluir punto (ej: `EDEN.K`, `MCHI.O`, `EWI.LTS`)
    - Debe tener al menos 2 letras
    - No debe ser un número puro
    
  - `etf_name`:
    - Buscar en **primeras 5 filas y 3 columnas**
    - Criterios: celda que contenga palabras clave: `"iShares"`, `"MSCI"`, `"ETF"`, `"Fund"`, `"Index"`
    - Debe tener más de 10 caracteres (asegurar nombre completo)
    - Ejemplo: `"iShares MSCI Chile ETF"`
    
  - `date`:
    - Buscar en **primeras 5 filas y 3 columnas**
    - Criterios: primera celda que se pueda parsear como fecha válida
    - Debe estar en rango razonable (2010-2030)
    - Formatos aceptados: `"31-12-2024"`, `"2024-12-31"`, `"31-Dec-2024"`, fechas de Excel, etc.
  
  **Validación**: Si no se encuentran los 3 metadatos, el ETF se marca con warning y se omite del procesamiento. Se genera un reporte `output/etfs_sin_metadatos.xlsx` con el listado de archivos omitidos y el motivo.
  
- Columnas relevantes de la tabla:
  - `RIC`
  - `Name`
  - `Country`
  - `Weight`
  - `No. Shares`

### 2) `input/data_maria_esg.xlsx`

Columnas esperadas:
- `Instrument`
- `Date`
- `ESG Score`

**Nota**: si el archivo no existe, mostrar error y terminar ejecución.
Este archivo en la columna ESG Score tiene valores numéricos con decimales entre 0 y 100 pero algunos pueden estar vacíos o ser NaN, esos casos deben ser manejados para no afectar el cálculo de los promedios ESG. Se descartan valores menores a 0 o NaN.

### 3) `input/data_maria_market_cap.xlsx`

Columnas esperadas:
- `Instrument`
- `Date`
- `Company Market Capitalization`

**Nota**: si el archivo no existe, mostrar error y terminar ejecución.  
Este archivo en la columna Company Market Capitalization tiene valores numéricos grandes (capitalización de mercado en dólares). Algunos pueden estar vacíos o ser NaN, esos casos deben ser manejados para no afectar el cálculo de los promedios ESG. Se descartan valores menores a 0 o NaN.

## Filosofía de operación

**Aprovecha al máximo los datos disponibles y mantiene trazabilidad completa**: 

- El sistema **NO descarta** filas sin RIC del archivo ETF original
- Mantiene todas las filas para tener trazabilidad completa
- Las filas sin RIC no harán match con los masters (tendrán NaN en ESG Score y Market Cap)
- Se generan estadísticas detalladas por cada ETF:
  - **total_instrumentos**: Total de filas en el ETF original
  - **sin_ric**: Instrumentos sin RIC (NaN o vacío)
  - **con_ric**: Instrumentos con RIC válido
  - **match_esg**: Con RIC y match en master ESG
  - **match_market_cap**: Con RIC y match en master Market Cap
  - **utilizados_analisis**: Con RIC y match completo (ESG + Market Cap válidos ≥ 0)
  - **porcentaje_utilizado**: % de instrumentos utilizados sobre el total

- **Requisito**:  El sistema lee TODAS las filas del ETF, **incluyendo las que no tienen RIC**. Esto permite:
  1. Mantener trazabilidad completa de todos los instrumentos del ETF
  2. Identificar cuántos instrumentos no tienen RIC
  3. Distinguir entre "sin RIC" vs "con RIC pero sin match en masters"
  
  El cruce con masters funciona así:
  - Filas **sin RIC**: No hacen match → ESG Score y Market Cap quedan en NaN
  - Filas **con RIC**: Se buscan en masters por RIC + año
    - Si hay match → Se agrega ESG Score y/o Market Cap
    - Si NO hay match → Quedan en NaN
  
  Para el cálculo de los promedios ESG solo se usan instrumentos que cumplan TODOS los criterios:
  - Tienen RIC válido (no NaN, no vacío)
  - Tienen match en master ESG (ESG Score ≥ 0, no NaN)
  - Tienen match en master Market Cap (Market Cap ≥ 0, no NaN)

- usar `pd.isna()` y `.strip() == ""` para detectar RIC vacíos.

## Reglas de normalizacion

### Name

Para mapear `Name -> RIC`:
1. Quitar espacios al inicio y final (`strip()`).
2. Convertir a mayusculas (`upper()`).
3. **Manejar valores NaN/None**: convertir a string vacío antes de normalizar.

Ejemplo:
- `" Enel Americas ord "` -> `"ENEL AMERICAS ORD"`
- `"vapores sa"` -> `"VAPORES SA"`
- `None` -> `""`

### Date

Para poder cruzar ETFs con masters:
1. **Parsear con `pd.to_datetime()`** (manejo automático de formatos múltiples).
2. **Extraer solo el año** de la fecha parseada (2010-2030).
3. El cruce con masters se realiza **solo por año**, no por fecha exacta.
   - Ejemplo: ETF con fecha `2024-12-31` hace match con master `2024-01-15` porque ambos son del año 2024.

Ejemplo de normalización:
- ETF: `31-12-2024` → año `2024`
- Master: `2024-01-15 00:00:00` → año `2024`
- **Match exitoso**: ambos son del año 2024

---

## Etapas del procesamiento

1. Cargar masters `data_maria_esg.xlsx` y `data_maria_market_cap.xlsx`.
2. Normalizar `Date` en ambos masters extrayendo solo el año (2010-2030).
3. Leer cada ETF desde `input/ETFS`:
   - Buscar metadatos en primeras 5 filas, 3 columnas
   - Leer tabla desde fila 4 (header=3)
   - Descartar filas sin RIC
4. Cruzar por:
   - `ETF.RIC` con `Master.Instrument`
   - `ETF.year` con `Master.year` (solo año, no fecha exacta)
5. Calcular dos indicadores ESG por ETF:
   - **Promedio simple de `ESG Score`**: suma los ESG de todas las empresas del fondo y divide por cuántas hay. Trata a todas por igual, sin importar su tamaño. Útil para tener una visión general del fondo.
   - **Promedio ponderado por `Company Market Capitalization`**: le da más peso al ESG de las empresas más grandes del fondo. Si una empresa representa el 30% del mercado y tiene ESG bajo, eso arrastra más el resultado que una empresa pequeña con ESG alto. Este número refleja mejor la realidad del fondo.
6. Exportar por cada ETF un archivo Excel con 2 hojas: Procesados y Descartados.
7. Exportar consolidado final a `output/resultado_esg_etf.xlsx`.
8. Generar reporte de trazabilidad a `output/trazabilidad_procesamiento.xlsx`.
9. Actualizar Base de données con los resultados ESG.

### Promedio simple ESG

`promedio_simple = promedio(ESG Score)`

Se calcula sobre los instrumentos con match valido por `RIC/Instrument` y `Date`.

### Promedio ponderado por Market Cap

`promedio_ponderado = sum(ESG Score * Company Market Capitalization) / sum(Company Market Capitalization)`

Se calcula solo donde existan ambos valores (`ESG Score` y `Company Market Capitalization`).

Sobrescribir completamente en cada ejecución (no acumular histórico).

### Salidas 

`output/resultado_esg_etf.xlsx`

Consolidado final de ETFs con metricas ESG.

Columnas finales:
- `etf_ticker`
- `etf_name`
- `year` (extraido de la fecha del ETF)
- `country`
- `esg_score_simple_avg` (float con 4 decimales)
- `esg_score_weighted_avg` (float con 4 decimales)
- `status` (siempre `"OK"`)
- `instruments_count` (cantidad de instrumentos analizados)

**Manejo de casos sin match**: si un ETF no tiene ningún match en los masters ESG/Market Cap, no aparece en el resultado final (pero se registra en el archivo que lleva la traza del error).

`output/ETFS/{ticker}_{year}.xlsx`

Por cada ETF que se procese, se exporta un Excel con 2 hojas:

**Hoja 1 - "Procesados"**: Instrumentos que tienen match con al menos uno de los masters (ESG o Market Cap):
- `RIC` (puede ser NaN si la fila no tenía RIC originalmente)
- `Name`
- `Country`
- `Weight`
- `No. Shares`
- `ESG Score` (NaN si no hay match en master ESG)
- `Company Market Capitalization` (NaN si no hay match en master Market Cap)

**Hoja 2 - "Descartados"**: Instrumentos sin match en ninguno de los masters:
- `RIC` (puede estar vacío/NaN)
- `Name`
- `Country`
- `Weight`
- `No. Shares`
- `Motivo de descarte`:
  - `"Sin RIC"` - La fila no tiene RIC
  - `"Con RIC pero sin match en masters ESG ni Market Cap"` - Tiene RIC pero no se encontró en ningún master

---

`output/trazabilidad_procesamiento.xlsx`

Reporte de calidad de datos por cada ETF procesado:
- `etf_name`: Nombre del ETF
- `total_instrumentos`: Total de instrumentos en el ETF (todas las filas)
- `sin_ric`: Instrumentos sin RIC (NaN o vacío)
- `con_ric`: Instrumentos con RIC válido
- `match_esg`: Instrumentos con RIC y match en master ESG
- `match_market_cap`: Instrumentos con RIC y match en master Market Cap
- `utilizados_analisis`: Instrumentos con RIC y match ESG Y Market Cap válidos (usados en cálculo)
- `porcentaje_utilizado`: % de instrumentos utilizados sobre el total

---

`output/etfs_sin_metadatos.xlsx`

Solo se genera si hay ETFs omitidos. Lista de archivos que no se pudieron procesar:
- `archivo`: Nombre del archivo ETF
- `motivo`: Razón por la que se omitió (ej: "Metadatos incompletos", "Tabla vacía")

---

`output/Base_de_donnees.xlsx`
`output/Base_de_donnees.xlsx`

Actualización de la Base de données con los resultados ESG mediante **merge bidireccional**:

1. Lee el archivo `input/Base de données.xlsm` (o `.xlsx` si no existe el `.xlsm`)
2. Extrae el ticker sin sufijos del punto: `etf_ticker` "ECH.K" → "ECH" (para match con `Country`)
3. Normaliza columnas `Country` y `Name` en ambos DataFrames (upper + strip)
4. Hace **merge OUTER** por `Country` (ticker sin sufijo), `Name` y `year`:
   - ✅ **Mantiene** todas las filas de Base de données original (incluso sin match con resultado ESG)
   - ✅ **Actualiza** las filas que tienen match con valores ESG
   - ✅ **Agrega** filas nuevas del resultado ESG que no están en Base de données
5. Agrega columnas `ESG moyen` (promedio simple) y `ESG pondéré` (promedio ponderado)
6. Las filas sin match conservan sus datos originales con ESG en NaN

**Resultado**: Base de données completa con:
- Todas las filas originales preservadas
- Nuevas filas del resultado ESG agregadas
- Valores ESG actualizados donde hay match

---

## Instalación

### 1. Crear ambiente virtual en PyCharm

1. **Abrir el proyecto** en PyCharm
2. **Ir a**: `File > Settings > Project: procesamiento_etfs > Python Interpreter`
3. **Hacer clic** en el ícono de engranaje ⚙️ (arriba a la derecha)
4. **Seleccionar**: `Add Interpreter > Add Local Interpreter`
5. **Elegir** la opción `Virtualenv Environment`
6. **Configurar**:
   - Location: dejar la ruta por defecto (será dentro del proyecto en `venv/`)
   - Base interpreter: seleccionar Python 3.8 o superior
7. **Hacer clic** en `OK`
8. **Esperar** a que PyCharm cree el ambiente virtual

Una vez creado el ambiente virtual, instalar las dependencias:

```powershell
pip install -r requirements.txt
```

**Nota**: PyCharm activará automáticamente el ambiente virtual cuando abras una terminal dentro del IDE.

### 2. Crear ejecutable con PyInstaller

#### En Windows:

1. **Instalar PyInstaller** (si no está en requirements.txt):
   ```powershell
   pip install pyinstaller
   ```

2. **Generar el ejecutable**:
   ```powershell
   pyinstaller --onefile --name procesamiento_etfs main.py
   ```

3. **Ubicación del ejecutable**:
   - El archivo `.exe` se generará en la carpeta `dist/`
   - Ruta: `dist/procesamiento_etfs.exe`

4. **Ejecutar**:
   ```powershell
   .\dist\procesamiento_etfs.exe
   ```

#### En Mac:

1. **Instalar PyInstaller** (si no está en requirements.txt):
   ```bash
   pip install pyinstaller
   ```

2. **Generar el ejecutable**:
   ```bash
   pyinstaller --onefile --name procesamiento_etfs main.py
   ```

3. **Ubicación del ejecutable**:
   - El archivo ejecutable se generará en la carpeta `dist/`
   - Ruta: `dist/procesamiento_etfs`

4. **Ejecutar**:
   ```bash
   ./dist/procesamiento_etfs
   ```

### 3. Distribución del ejecutable

El sistema está diseñado para que el ejecutable busque las carpetas `input/` y `output/` de forma **relativa a su ubicación**. Esto permite distribuir fácilmente la solución.

#### Estructura para distribución:

```text
carpeta_distribucion/
  procesamiento_etfs.exe (o procesamiento_etfs en Mac)
  procesamiento_etf.log (se genera automáticamente)
  input/
    ETFS/
      *.xlsx
    data_maria_esg.xlsx
    data_maria_market_cap.xlsx
    Base de donnees.xlsx
  output/
    (los archivos de salida se generarán aquí)
```

#### Pasos para distribuir:

1. **Crear una carpeta** para la distribución (ej: `ETF_Procesador/`)
2. **Copiar el ejecutable** desde `dist/` a esta carpeta
3. **Copiar la carpeta `input/`** completa con todos sus archivos
4. **Crear la carpeta `output/`** vacía (o el programa la creará automáticamente)
5. **Comprimir** la carpeta completa en un archivo .zip
6. **Distribuir** el archivo .zip

**Importante**: 
- El ejecutable siempre buscará las carpetas `input/` y `output/` en el mismo directorio donde esté ubicado
- Los archivos de log también se generarán en la misma ubicación del ejecutable
- Los ejecutables generados en Windows NO funcionarán en Mac y viceversa
- Debes generar el ejecutable en cada plataforma objetivo

---

## Ejecucion

Desde la raiz del proyecto:

```powershell
python main.py
```

Si necesitas instalar dependencias (ejemplo):

```powershell
pip install pandas openpyxl
```

---

## Etapas de desarrollo incremental

Para implementar `main.py` de forma mantenible, se siguieron estas etapas:

### Etapa 1: Estructura base y validaciones
1. ✅ Validaciones iniciales (archivos existen, carpetas de salida)
2. ✅ Función `normalizar_nombre()` - strip + upper
3. ✅ Función `normalizar_fecha_a_año()` - extrae año (2010-2030)

### Etapa 2: Carga de datos
1. ✅ Función `cargar_masters()` - carga ESG y Market Cap, normaliza a año
2. ✅ Función `leer_metadatos_etf()` - busca ticker, nombre, año (5 filas, 3 cols)
3. ✅ Función `leer_tabla_etf()` - lee desde fila 4, descarta sin RIC

### Etapa 3: Procesamiento
1. ✅ Función `cruzar_con_masters()` - merge por RIC + año
2. ✅ Función `calcular_esg()` - promedios simple y ponderado
3. ✅ Función `procesar_etf()` - orquesta procesamiento de 1 ETF

### Etapa 4: Exportación
1. ✅ Función `exportar_etf_individual()` - 2 hojas: Procesados/Descartados
2. ✅ Función `exportar_resultado_final()` - consolidado
3. ✅ Función `generar_reporte_trazabilidad()` - calidad de datos
4. ✅ Función `actualizar_base_donnees()` - merge con Base de données

### Etapa 5: Integración
1. ✅ Función `main()` - orquestación completa del proceso


---

## Criterios de mantenibilidad (sin sobreingenieria)

- Funciones cortas y con una sola responsabilidad.
- Nombres explicitos para columnas y variables.
- `if/for` claros en lugar de expresiones complejas.
- Evitar `lambda` y trucos de Python cuando no aporten claridad.
- Evitar comprensiones o cadenas de metodos complejas cuando afecten legibilidad.
- Mantener toda regla de negocio critica documentada aqui.

---

## Validaciones y manejo de errores

### Al inicio de la ejecución:
- ✅ Verificar que existe `input/ETFS/` y contiene al menos un archivo `.xlsx`
- ✅ Verificar que existen `input/data_maria_esg.xlsx` y `input/data_maria_market_cap.xlsx`
- ✅ Crear carpetas de salida si no existen: `output/`, `output/ETFS/`

### Durante el procesamiento:
- ✅ **Warning** por ETF sin metadatos válidos (no se encontraron ticker, nombre o fecha en primeras 5 filas, 3 columnas)
- ✅ **Warning** por ETF con tabla vacía o sin columnas requeridas
- ✅ **Warning** por ETF sin ningún match en masters ESG/Market Cap
- ✅ **Info** por cada ETF procesado exitosamente con cantidad de instrumentos utilizados

### Archivos de salida:
- ✅ Genera `etfs_sin_metadatos.xlsx` si hay ETFs omitidos
- ✅ Genera `trazabilidad_procesamiento.xlsx` con métricas de calidad
- ✅ Cada ETF exportado con 2 hojas (Procesados y Descartados)


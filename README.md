# procesamiento_etfs

Software que construye una base final con el **ESG por ETF** usando dos métodos de agregación:

- **Promedio simple** del `ESG Score` por ETF
- **Promedio ponderado** del `ESG Score` por `Company Market Capitalization`

El término *ESG bursátil* (Environmental, Social, and Governance) se refiere a la integración de criterios ambientales, sociales y de gobierno corporativo en la evaluación y selección de empresas cotizadas en bolsa.

---

## Estructura del proyecto

```
procesamiento_etfs/
├── data/
│   └── etf_holdings.csv        # Datos de entrada: holdings de cada ETF
├── output/                     # Directorio generado con el resultado
│   └── esg_base.csv
├── src/
│   └── esg_processor.py        # Módulo principal con la lógica de cálculo
├── tests/
│   └── test_esg_processor.py   # Tests unitarios
├── main.py                     # Punto de entrada
└── requirements.txt
```

---

## Formato del archivo de entrada

El CSV de entrada (`data/etf_holdings.csv`) debe contener las siguientes columnas:

| Columna    | Descripción                                      |
|------------|--------------------------------------------------|
| `ETF`      | Identificador del fondo (ej. `SPY`, `QQQ`)       |
| `Company`  | Nombre de la empresa                             |
| `ESG_Score`| Puntuación ESG de la empresa (0–100)             |
| `Market_Cap` | Capitalización bursátil en millones USD        |

---

## Instalación

```bash
pip install -r requirements.txt
```

---

## Uso

```bash
python main.py
# o especificando rutas personalizadas:
python main.py --input data/etf_holdings.csv --output output/esg_base.csv
```

### Salida

El programa genera un CSV con una fila por ETF:

| ETF | ESG_Promedio_Simple | ESG_Promedio_Ponderado |
|-----|---------------------|------------------------|
| IVV | 63.32               | 70.9239                |
| QQQ | 64.62               | 71.1112                |
| SPY | 64.66               | 69.4957                |

---

## Tests

```bash
python -m pytest tests/ -v
```

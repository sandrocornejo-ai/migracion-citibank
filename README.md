# Migración Citibank → Rex+

App Streamlit que transforma el Consolidado de Haberes y Descuentos de Citibank al formato de importación de **liquidaciones en detalle** de Rex+.

## Uso

```bash
pip install -r requirements.txt
streamlit run app.py
```

1. Sube el Consolidado de **una** empresa (CITIGROUP SERVICES CHILE SPA o CITIGROUP CHILE INTERNATIONAL SPA).
2. Indica si se usa fase (y su número). Si no, la columna Fase se elimina.
3. Genera y descarga el archivo de salida y el log de validaciones.

## Archivos

| Archivo | Descripción |
|---|---|
| `data/Equivalencia columnas conceptos.xlsx` | Encabezado de entrada → Id concepto Rex+ y tipo. Se puede reemplazar desde la barra lateral. |

## Reglas principales

- Conceptos desde la columna BK; montos en 0 se omiten salvo `sueldoBase`, `totalesEmpl`, `impuesto`, `cesEmpleado`.
- Columnas que apuntan al mismo concepto se suman (afp, apvi, …). Isapre: Fonasa si > 0; si no, Obligatoria + Adicional.
- Haberes afectos = Haber afecto + Haber afecto Especial + Haber SOLO Tributable. Haberes exentos = Haber exento.
- Afecto cesantía: TOPE IMPONIBLE AFC TRABAJADOR; si está en blanco, FO_6308 / 0,008.
- El log incluye cuadratura contra TOTAL HABERES, TOTAL DESCUENTOS y Líquido.

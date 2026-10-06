# monitor-eventos-fen-peru

Monitor automatizado de eventos hidrometeorológicos en Perú (lluvias intensas, huaicos, deslizamientos, inundaciones) vinculados al **Fenómeno El Niño (FEN)**. Extrae reportes de fuentes oficiales, los clasifica por **nivel de evidencia** y los registra por distrito en un libro de Excel.

> **Principio de diseño:** precisión antes que cobertura. Si una ubicación, fecha o causalidad no se puede justificar con evidencia estructurada, se deja vacía o se envía a revisión en vez de inferirla.

## Qué hace

- **Fuentes oficiales:** INDECI (noticias), COEN-INDECI (reportes de emergencia y sus PDF), Contraloría General de la República y ENFEN.
- **Ubicación:** distrito, provincia, departamento y **ubigeo INEI**. La zona (Niño / Alerta / Control) sale de un maestro territorial a nivel departamento; nunca se infiere del texto de una noticia.
- **Nivel de evidencia FEN** (columna `Nivel_evidencia_FEN`):

  | Nivel | Significado |
  |---|---|
  | 1 - Causa explícita | El reporte atribuye el evento a El Niño / ENFEN |
  | 2 - Declaratoria FEN vigente | El distrito figura en el anexo de un decreto de emergencia por FEN y el evento cae dentro de su vigencia |
  | 3 - Alerta ENFEN vigente | El evento ocurrió durante la alerta de El Niño Costero |
  | 4 - Sin vínculo FEN | Sin evidencia de vínculo |

- **Dos hojas de resultado:** `Eventos_reales` (niveles 1 y 2) y `Eventos_candidatos` (ocurrieron, pero sin vínculo verificado). Las publicaciones informativas, preventivas o de pronóstico solo se registran en el log.
- **Severidad** (`Escala_Magnitud`: Alta / Media / Baja) según los daños de la tabla del reporte, con los criterios de la hoja `Metodologia_y_fuentes`.
- **Modos:** semanal incremental (actualiza el Excel oficial) y *backfill* histórico sobre un rango de fechas (genera un Excel independiente, sin tocar el oficial).

## Instalación

Requiere Python 3.12 o superior.

```bash
pip install -r requirements.txt
```

## Uso

```bash
# Semanal (últimos 14 días). Siempre probar primero con --dry-run
python actualizar_eventos_fen_final.py --dry-run
python actualizar_eventos_fen_final.py

# Backfill de un año o de un rango (escribe un Excel nuevo en la carpeta del proyecto)
python actualizar_eventos_fen_final.py --backfill-year 2026 --dry-run --threads 12
python actualizar_eventos_fen_final.py --from-date 2026-01-01 --to-date 2026-10-06 --threads 12

# Regenerar el índice de decretos FEN a partir de los PDF en data/decretos/
python actualizar_eventos_fen_final.py --build-decretos
```

Opciones útiles: `--workbook` (Excel de entrada), `--output` (Excel de salida del backfill), `--ubigeo`, `--decretos`, `--max-new`, `--no-popup`. Códigos de salida: `0` OK, `1` error de scraping, `2` Excel inexistente, `3` estructura inválida, `4` error al guardar.

Al terminar se generan en `logs/` un log (`actualizacion_fen_*.log`) y un resumen JSON (`resumen_fen_*.json`) con contadores por fuente y por nivel de evidencia.

## Datos de entrada

| Archivo | Contenido |
|---|---|
| `data/Eventos Fenomeno del Niño (2).xlsx` | Libro base con las hojas `Distritos Zonas` (maestro territorial), `Eventos_reales` y `Metodologia_y_fuentes`. No se incluye en el repositorio: si falta, el script crea uno vacío. |
| `data/ubigeo_distrito.csv` | Catálogo de distritos del INEI (ubigeo, departamento, provincia, distrito). |
| `data/decretos/decretos_fen.json` | Distritos y vigencia de los decretos de emergencia por peligro inminente asociados al FEN 2026-2027 (DS 097-2026-PCM y DS 124-2026-PCM). Se genera con `--build-decretos` desde los PDF oficiales. |

## Notas de operación

- **gob.pe limita las ráfagas de peticiones** devolviendo una página casi vacía con código 200. El script limita la concurrencia, reintenta y lo reporta en `gobpe_limitacion` del resumen JSON. Si hay noticias no recuperadas, repetir con menos `--threads`.
- **El Excel de salida no puede estar abierto** mientras corre el backfill; si lo está, se guarda con un nombre alterno con la hora.
- Las peticiones usan `verify=False` por los certificados internos de un entorno corporativo. No es una recomendación general.
- Las fechas de vigencia de los decretos se definen en `DECREE_SOURCES` (el cuerpo de algunos decretos es una imagen ilegible); corregirlas allí y regenerar el índice si cambian.

## Estructura

```
actualizar_eventos_fen_final.py   # script principal (descubrimiento, extracción, clasificación, Excel)
requirements.txt
data/                             # ubigeo, índice de decretos (ver arriba)
README_FEN_AUTOMATIZACION.md      # documento de diseño original (parcialmente desactualizado)
CLAUDE.md                         # guía técnica de arquitectura y reglas para asistentes de código
```

Para la arquitectura y las reglas del proyecto (territorio, fechas, deduplicación, clasificación) ver `CLAUDE.md` y `README_FEN_AUTOMATIZACION.md`.

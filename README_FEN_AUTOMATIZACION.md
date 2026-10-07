# Automatizacion de Eventos del Fenomeno El Nino

Este documento explica el funcionamiento, decisiones de diseño, reglas de negocio y restricciones del script `actualizar_eventos_fen_v16.py`.

Su objetivo principal es que otro desarrollador o un asistente de codigo (por ejemplo Claude) pueda continuar modificando el proyecto **sin romper la metodologia de clasificacion, la deduplicacion, la escritura de Excel ni la inferencia territorial**.

---

## 1. Objetivo del sistema

El script mantiene una tabla de eventos hidrometeorologicos relacionados con el Fenomeno El Nino usando fuentes oficiales de Peru.

Trabaja sobre un Excel sincronizado localmente por OneDrive/SharePoint y tiene dos modos:

- **Weekly**: revision incremental de publicaciones recientes y actualizacion del Excel oficial.
- **Backfill**: reconstruccion historica de un rango de fechas en un workbook independiente.

El sistema distingue entre:

1. `Eventos_reales`: eventos observados con atribucion explicita a El Nino / ENFEN.
2. `Eventos_candidatos`: eventos hidrometeorologicos observados, pero sin atribucion explicita a El Nino.
3. Publicaciones informativas, preventivas, institucionales o pronosticos: solo log; no se insertan.

**Principio central:** precision antes que cobertura. Si una ubicacion, fecha o causalidad no puede justificarse de forma suficientemente robusta, se prefiere dejarla vacia o mandar el caso a revision antes que inferir algo incorrecto.

---

## 2. Archivo principal

Version documentada:

```text
actualizar_eventos_fen_v16.py
```

Dependencias:

```bash
pip install requests beautifulsoup4 openpyxl pypdf
```

Python usado actualmente: Python 3.12 en Windows 11 Enterprise.

El script esta pensado para ejecutarse desde PowerShell, Windows Task Scheduler o manualmente.

---

## 3. Excel de entrada

Ruta por defecto actual:

```text
EXCEL POR AHORA NO LEVANTADO
```

Los logs se guardan junto al script:

```text
<carpeta_del_script>\logs
```

No mover los logs al mismo archivo XLSX ni guardar trazas dentro de SharePoint salvo que exista una razon especifica.

---

## 4. Estructura esperada del workbook

### Hoja `Eventos_reales`

Columnas requeridas:

```text
Evento
Fecha
Magnitud / severidad observada
Escala_Magnitud
Departamento
Provincia
Distrito afectado
Zona
Fuente oficial
Detalle
```

La regla es:

> una fila = un evento x un distrito afectado

Si una publicacion afecta tres distritos, debe poder generar tres filas.

Si el evento cumple los criterios pero no se puede determinar el distrito con suficiente confianza, puede generarse una fila con ubicacion vacia. No inventar el distrito.

### Hoja `Distritos Zonas`

Columnas requeridas:

```text
Distrito (original)
Etiqueta dashboard
Departamento inferido
```

Esta hoja es el maestro territorial del proyecto.

`Zona` y `Departamento` se toman de este maestro cuando el distrito ha sido identificado de manera confiable.

**Nunca inferir una Zona Nino / Alerta / Control desde el texto de la noticia.**

### Hoja `Eventos_candidatos`

Si no existe, el script la crea.

Columnas:

```text
Evento
Fecha
Institucion
Departamento
Distrito
Fuente oficial
Motivo revision
```

---

## 5. Fuentes oficiales activas

Actualmente se consultan estas cuatro fuentes:

### INDECI

```text
https://www.gob.pe/institucion/indeci/noticias
```

### COEN - INDECI

```text
https://portal.indeci.gob.pe/emergencias/
```

### Contraloria General de la Republica

```text
https://www.gob.pe/institucion/contraloria/noticias
```

### ENFEN

```text
https://enfen.imarpe.gob.pe/noticias/
https://enfen.imarpe.gob.pe/comunicados/
```

### Fuente eliminada deliberadamente

**Municipalidad Distrital de Tambogrande NO debe reintroducirse sin redisenar su paginacion.**

Se elimino porque el esquema `/page/N/` generaba multiples 404 y no aportaba suficiente valor frente al costo y ruido del scraping.

---

## 6. Modos de ejecucion

### Weekly normal

```powershell
python actualizar_eventos_fen_v16.py
```

Default:

```text
lookback = 7 dias
threads = 4
```

Actualiza el workbook oficial si hay novedades.

### Weekly dry-run

```powershell
python actualizar_eventos_fen_v16.py --dry-run
```

Procesa, clasifica y genera logs, pero no guarda cambios.

### Backfill de un ano

```powershell
python actualizar_eventos_fen_v16.py --backfill-year 2026
```

Si no se indica `--output`, genera un archivo como:

```text
Eventos Fenomeno del Niño_backfill_20260101_20261231.xlsx
```

**El workbook oficial no se modifica.**

### Backfill dry-run

```powershell
python actualizar_eventos_fen_v16.py --backfill-year 2026 --dry-run
```

Simula una reconstruccion desde cero sin crear/guardar workbook.

### Backfill por rango

```powershell
python actualizar_eventos_fen_v16.py \
  --from-date 2023-01-01 \
  --to-date 2026-10-05 \
  --output "Eventos_FEN_2023_2026.xlsx"
```

### Threads

```powershell
python actualizar_eventos_fen_v16.py --backfill-year 2026 --dry-run --threads 8
```

Defaults:

```text
weekly   = 4 workers
backfill = 8 workers
maximo   = 12 workers
```

No aumentar el maximo sin una razon concreta. El cuello de botella principal es HTTP/PDF, no CPU.

---

## 7. Concurrencia

Se usa `ThreadPoolExecutor` para descargar y parsear HTML/PDF en paralelo.

Arquitectura:

```text
Descubrimiento de URLs
        |
        v
ThreadPoolExecutor
  |- HTML articulo A
  |- HTML articulo B
  |- PDF COEN C
  |- ...
        |
        v
Resultados Article
        |
        v
Clasificacion / deduplicacion / Excel
        |
        v
HILO PRINCIPAL
```

### Regla critica

**Nunca escribir en `openpyxl` desde varios threads.**

Los threads solo hacen I/O y parsing. Toda mutacion del workbook debe permanecer secuencial en el hilo principal.

Cada thread usa su propia `requests.Session` mediante thread-local storage.

---

## 8. Manejo HTTP

El entorno corporativo presenta certificados SSL internos, por lo que actualmente las peticiones se ejecutan con:

```python
verify=False
```

Las advertencias de `urllib3` estan deshabilitadas.

Esto es una concesion del entorno corporativo, no una recomendacion general de seguridad.

### Retries

- 404 y 410: **sin reintento**.
- Otros 4xx permanentes: normalmente sin reintento.
- 408 / 429 y errores de red transitorios: hasta 3 intentos.

No volver a reintentar 404 tres veces: en backfill destruye el rendimiento.

---

## 9. Descubrimiento y paginacion

### Weekly

Usa las paginas iniciales definidas en `SOURCES`.

### Backfill

Expande indices historicos con:

```text
--backfill-max-pages 60
```

Default: 60 paginas por fuente.

La paginacion puede seguir URLs deterministas y enlaces de navegacion detectados.

Si se modifica esta parte, mantener la regla de limite total por fuente para evitar loops o cientos de peticiones inutiles.

---

## 10. Clasificacion de publicaciones

### Senal de El Nino

`DIRECT_FEN_PATTERNS` reconoce expresiones como:

```text
Fenomeno El Nino
El Nino Costero
Nino Costero
ENFEN
region Nino 1+2 / 3.4
```

### Peligros hidrometeorologicos

Entre otros:

```text
lluvias intensas
precipitaciones
inundacion
huaico / huayco
desborde
activacion de quebrada
movimiento en masa
deslizamiento
aluvion
crecida
avenida de rio/quebrada
```

COEN tiene ademas reglas especificas y puede reconocer `vientos fuertes` como tipo hidrometeorologico para decidir si leer un PDF.

### Evidencia de ocurrencia

No basta con hablar de riesgo. Debe haber lenguaje de evento ocurrido, por ejemplo:

```text
se registro
ocurrio
se produjo
provoco
causo
afecto
desbordo
inundo
colapso
interrumpio
```

### No-eventos

Se excluyen:

```text
preparacion
prevencion
coordinacion institucional
capacitaciones
simulacros
pronosticos
escenarios de riesgo
inspecciones
supervision de obras
mantenimiento
servicios publicos
```

---

## 11. Regla para `Eventos_reales`

Una publicacion entra a `Eventos_reales` solo si, en el contexto principal del articulo/reporte:

```text
direct_fen_primary == True
AND hazard_primary == True
AND event_occurred_primary == True
```

Es decir:

1. atribucion explicita a El Nino / ENFEN;
2. peligro hidrometeorologico;
3. evidencia de evento ocurrido.

No asumir causalidad porque una noticia mencione El Nino en un parrafo secundario.

---

## 12. Regla para `Eventos_candidatos`

Un evento puede ir a `Eventos_candidatos` cuando:

```text
hazard_primary == True
event_occurred_primary == True
direct_fen_primary == False
candidate_event_context_ok(...) == True
```

`candidate_event_context_ok()` exige que la evidencia de peligro y ocurrencia esten suficientemente cerca en titulo/lead para evitar falsos positivos.

Motivo actual:

```text
Evento hidrometeorologico observado en fuente oficial, pero sin atribucion explicita al Fenomeno El Nino. Revisar antes de incorporarlo a Eventos_reales.
```

---

## 13. Tratamiento especial de COEN

COEN es la fuente mas importante y tambien la mas delicada.

Las paginas HTML suelen incluir enlaces a muchos documentos no relacionados. Por eso el PDF no se debe elegir por proximidad generica ni descargando todo.

### Seleccion del PDF

1. Extraer el numero de reporte desde URL/titulo.
2. Determinar primero si el tipo de emergencia es hidrometeorologico.
3. Buscar un enlace PDF cuyo URL/archivo contenga **el mismo numero de reporte**.
4. Descargar solo ese PDF.
5. Validar nuevamente el numero dentro del texto extraido.

No volver al comportamiento antiguo de descargar decenas de PDFs y validarlos despues.

### PDFs

Se usa `pypdf`.

No hay OCR.

Limites actuales:

```text
PDF_MAX_BYTES = 25 MB
PDF_MAX_PAGES = 40
PDF_LEAD_CHARS = 5000
```

---

## 14. Inferencia territorial: regla mas importante del proyecto

Los errores territoriales son el principal riesgo del sistema.

**No buscar nombres de distritos sueltos en todo el PDF.**

Los reportes COEN contienen:

- domicilios institucionales;
- firmas;
- fuentes;
- nombres de municipalidades;
- provincias;
- sectores;
- centros poblados;
- referencias a otras ubicaciones.

Un nombre encontrado en el PDF no implica que sea el distrito afectado.

---

## 15. Metodologia territorial de COEN v16

COEN usa `infer_coen_locations()` y no la inferencia generica.

Orden de evidencia:

### 1. Titulo del reporte

Ejemplo:

```text
LLUVIAS INTENSAS EN EL DISTRITO DE SAN MARCOS - ANCASH
```

Esto es evidencia fuerte de:

```text
distrito = SAN MARCOS
departamento = ANCASH
```

### 2. Seccion `1. HECHOS`

Se buscan relaciones linguisticas explicitas:

```text
distrito de X
distrito y provincia de X
dist. X
```

### 3. Seccion `2. UBICACION`

Se aisla esa seccion del PDF antes de analizarla.

Se usa para:

- corroborar distrito;
- identificar departamento;
- desambiguar homonimos.

### 4. Roles geograficos negativos

Un nombre NO debe promoverse a distrito si esta etiquetado como:

```text
centro poblado
sector
provincia
departamento
```

salvo que exista otra evidencia explicita que diga `distrito de X`.

### 5. Restriccion por departamento

Si el reporte indica `ANCASH`, un distrito homonimo del maestro asociado a `ICA` queda descartado.

Esto es obligatorio para evitar errores con nombres repetidos como:

```text
Santa Cruz
San Miguel
Miraflores
San Pedro
San Juan
```

---

## 16. Casos de regresion que NO deben volver a fallar

### Caso A: Chorrillos institucional

PDF:

```text
Centro de Operaciones de Emergencia Nacional
Av. El Sol, Cdra. 4 - Chorrillos, Lima - Peru
...
LLUVIAS INTENSAS EN EL DISTRITO DE YAUYOS - LIMA
...
UBICACION:
LIMA | YAUYOS | YAUYOS | YAUYOS
```

Resultado correcto:

```text
Departamento = Lima
Distrito     = YAUYOS
```

Resultado prohibido:

```text
Distrito = CHORRILLOS
```

`Chorrillos` es domicilio institucional de COEN, no ubicacion del evento.

### Caso B: Santa Cruz de Mosna

Reporte:

```text
LLUVIAS INTENSAS EN EL DISTRITO DE SAN MARCOS - ANCASH
```

HECHOS:

```text
centro poblado de Santa Cruz de Mosna,
distrito de San Marcos,
provincia de Huari
```

UBICACION:

```text
DEPARTAMENTO | PROVINCIA | DISTRITO  | CENTRO POBLADO
ANCASH       | HUARI     | SAN MARCOS | SANTA CRUZ DE MOSNA
```

Resultado correcto:

```text
Departamento = Ancash
Distrito     = SAN MARCOS
```

Resultado prohibido:

```text
Departamento = Ica
Distrito     = SANTA CRUZ
```

Aunque `Santa Cruz` exista como distrito en el maestro, en este reporte tiene rol de **centro poblado**, no distrito.

Estos dos casos deben considerarse tests manuales de regresion cada vez que se modifique la logica territorial.

---

## 17. Limpieza de ruido institucional en PDFs COEN

`clean_coen_pdf_location_text()` elimina lineas como:

```text
Distribucion:
Centro de Operaciones de Emergencia Nacional
Av. El Sol...
Telefonos
COENPeru
www.indeci.gob.pe
pies de pagina
```

No eliminar esta etapa sin reemplazarla por una estrategia mejor.

---

## 18. Fecha del evento

Para COEN se intenta obtener la **fecha real de ocurrencia** desde el reporte.

Prioridad conceptual:

```text
fecha de ocurrencia del PDF/HTML
    > fecha del reporte/titulo
    > fecha de publicacion
```

`effective_event_date()` usa `article.event_date` antes que `article.published`.

No usar automaticamente la fecha de publicacion del portal como fecha del evento si el reporte contiene la fecha de ocurrencia.

---

## 19. Deduplicacion

Hay dos niveles:

### URL

Evita volver a procesar una publicacion ya incorporada en ejecuciones semanales.

### Clave por fila

Se genera con:

```text
titulo normalizado
+ fecha efectiva
+ departamento
+ distrito
```

La URL no puede usarse como unica clave dentro de una publicacion porque una misma noticia puede producir varias filas, una por distrito.

En backfill se parte de:

```python
existing_urls = set()
existing_keys = set()
```

porque la tabla historica se construye desde cero.

---

## 20. Magnitud / severidad

Actualmente es una heuristica simple basada en patrones de texto.

Ejemplos:

```text
Alta: fallecidos, desaparecidos, colapso, destruccion, damnificados, evacuacion...
Media: afectados, danos, interrupciones, inundacion, huaico, desborde...
```

No interpretar esta escala como una escala oficial de INDECI. Es una clasificacion interna para el dataset.

---

## 21. Guardado del Excel

Se usa escritura segura:

```text
workbook -> archivo temporal -> os.replace()
```

Hay retries por bloqueo/OneDrive.

En modo weekly, si no se agregan eventos ni candidatos, el XLSX no se toca para evitar generar versiones inutiles en SharePoint.

En modo backfill se crea un workbook separado y se limpian los datos de `Eventos_reales` y `Eventos_candidatos`, preservando estructura y formato.

---

## 22. Logs y resumen JSON

Cada ejecucion crea:

```text
actualizacion_fen_YYYYMMDD_HHMMSS.log
resumen_fen_YYYYMMDD_HHMMSS.json
```

El resumen incluye, entre otros:

```text
eventos_reales_agregados
articulos_evento_agregados
eventos_candidatos_agregados
fen_sin_evento
omitidos_antiguedad
omitidos_sin_fecha
omitidos_irrelevantes
duplicados
articulos_no_recuperados
pdf_coen_encontrados
pdf_coen_leidos
fuentes
```

Tambien se imprime una linea `PAD_SUMMARY|...` cuando existe consola.

---

## 23. Codigos de salida

```text
0 = OK
1 = error general / scraping no recuperable
2 = workbook inexistente o no se puede abrir
3 = estructura Excel invalida / error preparando backfill
4 = error al guardar/reemplazar Excel
```

---

## 24. Reglas para futuras modificaciones

Claude / desarrollador: antes de cambiar el script, respetar estas reglas.

### NO hacer

- No reintroducir Tambogrande sin resolver correctamente su paginacion.
- No paralelizar escrituras de `openpyxl`.
- No convertir cualquier nombre territorial encontrado en el PDF en distrito.
- No usar domicilio de COEN/INDECI como ubicacion del evento.
- No promover `centro poblado`, `sector` o `provincia` a distrito.
- No resolver homonimos ignorando el departamento.
- No descargar todos los PDFs de una pagina COEN para descubrir cual corresponde.
- No considerar preparacion/pronostico como evento observado.
- No asumir que mencionar El Nino en cualquier parte del cuerpo implica causalidad.
- No usar la fecha de publicacion si existe fecha de ocurrencia confiable.
- No modificar el workbook oficial durante un backfill.

### SI hacer

- Preferir evidencia estructurada del documento.
- Usar titulo + HECHOS + UBICACION para COEN.
- Restringir homonimos por departamento.
- Mantener una fila por distrito.
- Mantener candidatos separados de eventos reales.
- Mantener logs auditables.
- Agregar pruebas de regresion cuando aparezca un nuevo falso positivo territorial.
- Hacer primero `--dry-run` antes de ejecutar cambios de scraping en produccion.

---

## 25. Orden recomendado para depurar un caso incorrecto

Cuando una fila salga mal:

1. Revisar titulo del reporte.
2. Revisar `article.published` y `article.event_date`.
3. Revisar si se encontro el PDF exacto.
4. Revisar `1. HECHOS`.
5. Revisar `2. UBICACION`.
6. Revisar el rol del nombre problematico: distrito / provincia / sector / centro poblado / domicilio institucional.
7. Revisar el maestro `Distritos Zonas` por homonimos.
8. Revisar departamento inferido.
9. Reproducir el caso en `--dry-run`.
10. Convertirlo en test de regresion antes de cambiar la heuristica general.

No solucionar un falso positivo agregando una excepcion especifica del tipo:

```python
if district == "CHORRILLOS": ignore
```

La solucion debe ser metodologica (rol semantico, seccion del documento, departamento, evidencia estructurada), no una blacklist de nombres concretos.

---

## 26. Mejoras futuras recomendadas

Sin cambiar la metodologia actual, las mejoras mas utiles serian:

1. Tests unitarios para `infer_coen_locations()` con PDFs/textos fixture.
2. Parser explicito de la tabla `UBICACION` por columnas:
   `DEPARTAMENTO / PROVINCIA / DISTRITO / SECTOR|CENTRO POBLADO`.
3. Guardar en logs el origen de cada decision territorial (`titulo`, `hechos`, `ubicacion`).
4. Validar jerarquia distrito-provincia-departamento contra un ubigeo oficial.
5. Crear una columna interna o log de `location_confidence` sin alterar el Excel de negocio.
6. Mejorar el parser de fechas COEN para formatos defectuosos extraidos desde PDF.
7. Crear una suite de regresion con todos los falsos positivos conocidos.

---

## 27. Checklist antes de entregar una nueva version

```text
[ ] py_compile pasa
[ ] weekly --dry-run pasa
[ ] backfill 2026 --dry-run arranca y pagina correctamente
[ ] threads no escriben Excel
[ ] 404 no generan 3 retries
[ ] PDF COEN exacto se selecciona por numero de reporte
[ ] Incendios/sismos/accidentes no disparan lectura PDF hidro
[ ] Caso Yauyos no devuelve Chorrillos
[ ] Caso San Marcos no devuelve Santa Cruz/Ica
[ ] una publicacion con N distritos puede producir N filas
[ ] backfill no modifica el Excel oficial
[ ] logs + resumen JSON se generan
[ ] sin novedades weekly no genera una nueva version del XLSX
```

---

## 28. Comando recomendado de validacion actual

```powershell
python actualizar_eventos_fen_v16.py --backfill-year 2026 --dry-run --threads 8
```

Para weekly:

```powershell
python actualizar_eventos_fen_v16.py --dry-run
```

Una vez validado:

```powershell
pythonw.exe actualizar_eventos_fen_v16.py
```

---

## 29. Resumen conceptual para un asistente de codigo

Si solo recuerdas cinco cosas, recuerda estas:

1. **Eventos reales exigen FEN explicito + peligro + ocurrencia.**
2. **Eventos hidro sin FEN explicito son candidatos, no reales.**
3. **COEN debe interpretarse por estructura semantica, no por coincidencias de nombres.**
4. **Titulo/HECHOS/UBICACION y departamento dominan la inferencia territorial.**
5. **Precision > recall: mejor vacio/revision que una ubicacion falsa.**


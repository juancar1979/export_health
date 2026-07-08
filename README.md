# 🩺 Exportador de Apple Salud a Excel

> Convierte la exportación completa de **Apple Salud** (`export.xml`) en un libro **Excel organizado, filtrable y con formato visual**, pensado para revisar históricos largos sin pelearte con el XML original.

![Python](https://img.shields.io/badge/Python-3.10%2B-blue)
![Excel](https://img.shields.io/badge/Salida-.xlsx-green)
![Fuente](https://img.shields.io/badge/Fuente-Apple%20Salud-lightgrey)
![Estado](https://img.shields.io/badge/Uso-local%20sin%20subir%20datos-success)

---

## 📌 Resumen ejecutivo

Apple Salud permite exportar todos los datos de salud y actividad en formato **XML**, pero ese archivo no es cómodo para revisar, filtrar o analizar. Este proyecto automatiza la conversión a Excel y genera un libro con hojas separadas por tipo de dato, rangos de fechas configurables y distintos niveles de detalle.

El objetivo es que puedas pasar de esto:

```text
apple_health_export/export.xml
```

a esto:

```text
salud_202606.xlsx
├─ Resumen
├─ Resumen_Frecuencia cardiaca
├─ Frecuencia cardiaca
├─ Frecuencia cardiaca caminando
├─ Frecuencia cardiaca en reposo
├─ Pasos
├─ Sueño
├─ Peso
├─ Energía activa
├─ Entreno caminar
└─ ...
```

### ✅ Qué aporta

| Necesidad | Cómo lo resuelve |
|---|---|
| Ver datos de Apple Salud sin leer XML | Genera un Excel `.xlsx` estructurado |
| Limitar el análisis a fechas concretas | Parámetros `--inicio` y `--fin` |
| Elegir granularidad | `--detalle dia`, `hora`, `min` o `todos` |
| Revisar frecuencia cardiaca fácilmente | Hoja específica `Resumen_Frecuencia cardiaca` |
| Trabajar con muchos datos | Usa SQLite temporal y divide hojas si superan el límite de Excel |
| Tener un Excel legible | Cabeceras, filtros, tablas, colores de pestañas y anchos predefinidos |
| Ejecutarlo sin tocar comandos complejos | Incluye `.bat` para Windows y `.sh` para macOS/Linux |

---

## 🧭 Índice

1. [Archivos incluidos](#-archivos-incluidos)
2. [Exportar datos desde Apple Salud](#-exportar-datos-desde-apple-salud)
3. [¿Existe algo parecido en Android?](#-existe-algo-parecido-en-android)
4. [Instalación](#-instalación)
5. [Uso rápido](#-uso-rápido)
6. [Ejecución con BAT o SH](#-ejecución-con-bat-o-sh)
7. [Parámetros disponibles](#-parámetros-disponibles)
8. [Ejemplos habituales](#-ejemplos-habituales)
9. [Qué genera el Excel](#-qué-genera-el-excel)
10. [Hoja `Resumen_Frecuencia cardiaca`](#-hoja-resumen_frecuencia-cardiaca)
11. [Formato visual aplicado](#-formato-visual-aplicado)
12. [Limitaciones importantes](#-limitaciones-importantes)
13. [Errores frecuentes](#-errores-frecuentes)
14. [Flujo recomendado](#-flujo-recomendado)

---

## 📦 Archivos incluidos

| Archivo | Para qué sirve |
|---|---|
| `exportar_salud_completo.py` | Script principal de conversión Apple Salud → Excel |
| `ejecutar_exportar_salud.bat` | Lanzador interactivo para Windows |
| `ejecutar_exportar_salud.sh` | Lanzador interactivo para macOS/Linux |
| `README_exportar_salud.md` | Esta guía |

Estructura recomendada:

```text
salud/
├─ exportar_salud_completo.py
├─ ejecutar_exportar_salud.bat
├─ ejecutar_exportar_salud.sh
└─ apple_health_export/
   └─ export.xml
```

---

## 🍎 Exportar datos desde Apple Salud

Apple permite exportar todos los datos de salud y actividad desde la app **Salud** en formato **XML**.

### Pasos en iPhone

1. Abre la app **Salud**.
2. Entra en **Resumen**.
3. Pulsa tu **foto o iniciales** arriba a la derecha.
4. Pulsa **Exportar todos los datos de salud**.
5. Confirma la exportación.
6. Guarda o envía el `.zip` mediante **Archivos**, **iCloud Drive**, **AirDrop** u otro método.
7. Descomprime el `.zip`.
8. Localiza el archivo:

```text
apple_health_export/export.xml
```

> Fuente oficial: [Apple Support — Share your health and fitness data in XML format](https://support.apple.com/guide/iphone/share-your-health-data-iph5ede58c3d/ios)

### Qué archivo necesita el script

El script puede recibir:

```text
apple_health_export/export.xml
```

o directamente la carpeta:

```text
apple_health_export/
```

Si le pasas la carpeta, busca automáticamente `export.xml` dentro de ella y sus subcarpetas.

---

## 🤖 ¿Existe algo parecido en Android?

Sí, pero **no funciona igual que Apple Salud**.

Android puede tener datos en varias plataformas: **Health Connect**, **Google Fit**, apps del fabricante, relojes deportivos, básculas, etc. No hay un único `export.xml` equivalente al de Apple Salud que este script pueda leer directamente.

### Health Connect

Health Connect permite configurar exportaciones/copia de seguridad desde los ajustes de Android.

Ruta orientativa:

```text
Ajustes de Android
→ buscar "Health Connect"
→ Gestionar datos
→ Copia de seguridad y restauración
→ Exportación programada
```

El archivo generado puede llamarse:

```text
Health Connect.zip
```

> Fuente oficial: [Android Help — Back up & restore your Health Connect data](https://support.google.com/android/answer/15323271)

### Google Fit

Google Fit puede exportarse mediante **Google Takeout**.

Proceso orientativo:

1. Entra en **Google Takeout**.
2. Deselecciona todo.
3. Selecciona **Fit**.
4. Crea la exportación.
5. Descarga el archivo generado.

> Fuente oficial: [Google Fit Help — How to download your Google data](https://support.google.com/fit/answer/3024190)

### Limitación real

Este script **solo convierte Apple Salud `export.xml`**.  
Los exports de Android/Google Fit/Health Connect usan otros formatos, así que habría que crear otro parser específico.

---

## 🛠️ Instalación

Necesitas Python y `openpyxl`.

```bat
python -m pip install openpyxl
```

El resto usa librerías estándar de Python:

```text
argparse
sqlite3
csv
json
re
datetime
pathlib
xml.etree.ElementTree
collections
```

### Comprobar Python

```bat
python --version
```

Si Windows no reconoce `python`, prueba:

```bat
py --version
```

---

## ⚡ Uso rápido

Desde la carpeta donde esté el script:

```bat
python exportar_salud_completo.py
```

Valores por defecto:

| Opción | Valor por defecto |
|---|---|
| XML | Busca `export.xml` automáticamente |
| Fechas | Sin límite |
| Detalle | `hora` |
| Salida | Nombre automático tipo `apple_health_hora.xlsx` |

---

## ▶️ Ejecución con BAT o SH

También puedes usar los lanzadores interactivos.

### Windows

Doble clic en:

```text
ejecutar_exportar_salud.bat
```

El BAT pregunta:

| Pregunta | Qué poner |
|---|---|
| Ruta a `export.xml` | Puedes dejarlo vacío si está dentro de la carpeta del proyecto |
| Fecha inicio | `YYYY-MM-DD` o vacío |
| Fecha fin | `YYYY-MM-DD` o vacío |
| Detalle | `dia`, `hora`, `min` o `todos` |
| Nombre Excel salida | Vacío para nombre automático |

El BAT también comprueba si `openpyxl` está instalado e intenta instalarlo si falta.

### macOS / Linux

Primero da permisos:

```sh
chmod +x ejecutar_exportar_salud.sh
```

Luego ejecuta:

```sh
./ejecutar_exportar_salud.sh
```

### Versión esperada

Al ejecutar, debería aparecer una línea como:

```text
Versión script: 2026-07-08_formato_excel_v3_ruta_carpeta
```

Si no aparece, probablemente estás ejecutando una copia antigua de `exportar_salud_completo.py`.

---

## ⚙️ Parámetros disponibles

### `--xml`

Ruta al archivo `export.xml` o a la carpeta `apple_health_export`.

```bat
python exportar_salud_completo.py --xml "C:\Users\juanc\Downloads\salud\apple_health_export\export.xml"
```

También vale indicar la carpeta:

```bat
python exportar_salud_completo.py --xml "C:\Users\juanc\Downloads\salud\apple_health_export"
```

---

### `--inicio`

Fecha inicial incluida.

```bat
python exportar_salud_completo.py --inicio 2024-01-01
```

Formato obligatorio:

```text
YYYY-MM-DD
```

---

### `--fin`

Fecha final incluida.

```bat
python exportar_salud_completo.py --fin 2024-12-31
```

La fecha final incluye el día completo. Por ejemplo:

```bat
--fin 2024-12-31
```

incluye registros hasta:

```text
2024-12-31 23:59:59
```

---

### `--detalle`

Nivel de detalle de la exportación.

| Valor | Qué hace | Uso recomendado |
|---|---|---|
| `dia` | Agrupa por día | Histórico largo y tendencias generales |
| `hora` | Agrupa por hora | Frecuencia cardiaca, actividad, oxígeno, ruido |
| `min` | Agrupa por minuto | Periodos cortos con mucho detalle |
| `todos` | No agrupa; exporta registros originales | Auditoría o revisión exacta |

Alias aceptados:

| Principal | Alias |
|---|---|
| `dia` | `día`, `day` |
| `hora` | `hour` |
| `min` | `minuto`, `minute` |
| `todos` | `todo`, `all`, `raw` |

---

### `--salida`

Nombre personalizado para el Excel generado.

```bat
python exportar_salud_completo.py --salida salud_2024.xlsx
```

---

### `--db`

Nombre del archivo SQLite temporal.

Por defecto:

```text
apple_health_temp.sqlite
```

Normalmente no hace falta cambiarlo.

---

## 🧪 Ejemplos habituales

### Exportar todo el histórico por hora

```bat
python exportar_salud_completo.py --detalle hora
```

### Exportar todo 2024 por día

```bat
python exportar_salud_completo.py --inicio 2024-01-01 --fin 2024-12-31 --detalle dia
```

### Exportar enero de 2024 por hora

```bat
python exportar_salud_completo.py --inicio 2024-01-01 --fin 2024-01-31 --detalle hora
```

### Exportar un día concreto por minuto

```bat
python exportar_salud_completo.py --inicio 2024-05-10 --fin 2024-05-10 --detalle min
```

### Exportar registros originales sin agrupar

```bat
python exportar_salud_completo.py --inicio 2024-01-01 --fin 2024-01-31 --detalle todos
```

### Exportar con ruta exacta y nombre de salida

```bat
python exportar_salud_completo.py --xml "C:\Users\juanc\OneDrive\Personal\salud\apple_health_export" --inicio 2024-06-01 --fin 2026-06-30 --detalle hora --salida salud_202606_v3.xlsx
```

---

## 📊 Qué genera el Excel

El libro generado incluye:

| Hoja | Contenido |
|---|---|
| `Resumen` | Configuración usada, fechas, tipos detectados y número de registros |
| `Resumen_Frecuencia cardiaca` | Panel específico de frecuencia cardiaca, caminando y reposo |
| Una hoja por tipo de dato | Datos agregados o crudos según `--detalle` |
| `Resumen actividad diaria` | Si existen datos suficientes, resumen de actividad diaria |

Ejemplos de hojas posibles:

```text
Resumen
Resumen_Frecuencia cardiaca
Frecuencia cardiaca
Frecuencia cardiaca caminando
Frecuencia cardiaca en reposo
Variabilidad cardiaca HRV
Pasos
Sueño
Peso
Saturación de oxígeno
Energía activa
Energía basal
Distancia andando corriendo
Entreno caminar
Entreno correr
VO2 máx.
Ruido ambiental
```

---

## ❤️ Hoja `Resumen_Frecuencia cardiaca`

El script genera una hoja específica llamada:

```text
Resumen_Frecuencia cardiaca
```

Incluye tres bloques:

| Bloque | Dato Apple usado |
|---|---|
| Frecuencia cardiaca | `HKQuantityTypeIdentifierHeartRate` |
| Frecuencia cardiaca caminando | `HKQuantityTypeIdentifierWalkingHeartRateAverage` |
| Frecuencia cardiaca en reposo | `HKQuantityTypeIdentifierRestingHeartRate` |

Cada bloque resume:

| Sección | Campo | Significado |
|---|---|---|
| Promedio | Muestras | Total de mediciones usadas |
| Promedio | Mínimo | Promedio de los mínimos por intervalo |
| Promedio | Media | Promedio de las medias por intervalo |
| Promedio | Máximo | Promedio de los máximos por intervalo |
| Máximos y mínimos | MIN / Mínimo | Menor mínimo observado |
| Máximos y mínimos | MIN / Media | Menor media observada |
| Máximos y mínimos | MIN / Máximo | Menor máximo observado |
| Máximos y mínimos | MAX / Mínimo | Mayor mínimo observado |
| Máximos y mínimos | MAX / Media | Mayor media observada |
| Máximos y mínimos | MAX / Máximo | Mayor máximo observado |

La hoja usa el mismo nivel de detalle elegido con `--detalle`.

Ejemplo:

| Detalle elegido | Cómo se calcula el resumen |
|---|---|
| `dia` | Sobre intervalos diarios |
| `hora` | Sobre intervalos horarios |
| `min` | Sobre intervalos por minuto |
| `todos` | Sobre registros originales |

> Nota: los valores se escriben ya calculados. No son fórmulas enlazadas a otras hojas. Esto evita errores con hojas muy grandes, hojas partidas o nombres recortados por Excel.

---

## 🎨 Formato visual aplicado

El Excel generado intenta ser cómodo de revisar desde el primer momento.

| Elemento | Resultado |
|---|---|
| Cabeceras | Fondo azul y texto blanco |
| Filtros | Activados en hojas de datos |
| Tablas Excel | Bandas alternas en las hojas de datos |
| Primera fila | Congelada |
| Columnas | Anchos predefinidos según el tipo de dato |
| Cuadrícula | Oculta para una vista más limpia |
| Pestañas | Coloreadas según tipo de hoja |
| Resumen FC | Panel visual específico para frecuencia cardiaca |

Colores de pestañas:

| Color | Uso |
|---|---|
| Azul | Datos numéricos y resúmenes |
| Verde | Categorías, por ejemplo sueño |
| Naranja | Entrenamientos |
| Morado | Resumen de actividad |

Nota honesta: no se aplica formato manual a millones de celdas porque puede hacer que el Excel pese mucho y tarde bastante más en abrirse. Se prioriza el formato que aporta legibilidad sin penalizar demasiado el rendimiento.

---

## 🧾 Columnas generadas

Las columnas dependen del tipo de dato y del nivel de detalle.

### Datos numéricos

En modo `dia`, `hora` o `min`, registros como frecuencia cardiaca, pasos, peso, oxígeno, energía o distancia generan columnas como:

| Columna | Significado |
|---|---|
| `periodo` | Día, hora o minuto agrupado |
| `tipo` | Nombre legible del dato |
| `fuente` | Dispositivo o app que generó el dato |
| `unidad` | Unidad normalizada |
| `mediciones` | Número de registros agrupados |
| `minimo` | Valor mínimo del periodo |
| `media` | Valor medio del periodo |
| `maximo` | Valor máximo del periodo |
| `suma` | Suma de valores del periodo |
| `primera_fecha` | Primer registro del periodo |
| `ultima_fecha` | Último registro del periodo |

### Categorías

Para categorías como sueño:

| Columna | Significado |
|---|---|
| `periodo` | Día, hora o minuto agrupado |
| `tipo` | Tipo de dato |
| `fuente` | Dispositivo o app |
| `estado` | Estado de la categoría |
| `registros` | Número de registros agrupados |
| `minutos_total` | Duración total en minutos |
| `minutos_media_registro` | Duración media por registro |
| `primera_fecha` | Primer registro del periodo |
| `ultima_fecha` | Último registro del periodo |

### Entrenamientos

Para entrenamientos:

| Columna | Significado |
|---|---|
| `periodo` | Día, hora o minuto agrupado |
| `tipo` | Tipo de entrenamiento |
| `fuente` | Dispositivo o app |
| `entrenos` | Número de entrenamientos |
| `minutos_total` | Duración total |
| `minutos_media` | Duración media |
| `distancia_km_total` | Distancia total en km |
| `energia_kcal_total` | Energía total en kcal |
| `primera_fecha` | Primer entrenamiento del periodo |
| `ultima_fecha` | Último entrenamiento del periodo |

### Modo `todos`

Cuando usas:

```bat
--detalle todos
```

las hojas incluyen columnas próximas al dato original:

| Columna | Significado |
|---|---|
| `clase` | Registro, categoría o entrenamiento |
| `tipo` | Nombre legible |
| `tipo_original` | Nombre técnico original de Apple |
| `fuente` | Dispositivo o app |
| `inicio` | Fecha/hora de inicio |
| `fin` | Fecha/hora de fin |
| `duracion_min` | Duración en minutos |
| `valor_num` | Valor numérico normalizado |
| `valor_texto` | Valor textual, si aplica |
| `unidad` | Unidad normalizada |
| `unidad_original` | Unidad original del XML |
| `distancia_km` | Distancia normalizada en km |
| `energia_kcal` | Energía normalizada en kcal |
| `creacion` | Fecha de creación del registro |
| `extra` | Datos adicionales |

---

## 📐 Unidades normalizadas

| Dato | Unidad normalizada |
|---|---|
| Frecuencia cardiaca | `lpm` |
| Frecuencia cardiaca en reposo | `lpm` |
| Frecuencia cardiaca caminando | `lpm` |
| Recuperación cardiaca | `lpm` |
| HRV | `ms` |
| Pasos | `pasos` |
| Peso | `kg` |
| Altura | `cm` |
| Energía activa/basal | `kcal` |
| Distancia andando/corriendo | `km` |
| Distancia ciclismo | `km` |
| Tiempo de ejercicio | `min` |
| Tiempo de pie | `min` |
| Saturación de oxígeno | `%` |
| Frecuencia respiratoria | `resp/min` |
| Temperatura corporal | `°C` |
| VO2 máx. | `ml/kg/min` |

Nota honesta: para saturación de oxígeno, el script asume que si Apple exporta `0.97`, equivale a `97 %`. Esto suele ser correcto, pero conviene verificar algunas filas reales.

---

## ⚠️ Limitaciones importantes

### 1. Límite de Excel

Excel permite un máximo de:

```text
1.048.576 filas por hoja
```

El script deja una fila para cabeceras, así que usa como máximo:

```text
1.048.575 filas de datos por hoja
```

Si una hoja supera ese límite, se divide automáticamente:

```text
Frecuencia cardiaca
Frecuencia cardiaca_2
Frecuencia cardiaca_3
```

---

### 2. Datos con duración

En los modos `dia`, `hora` y `min`, los registros con duración se agrupan por la fecha/hora de inicio.

Ejemplo:

```text
Sueño empieza: 2024-01-01 23:30
Sueño termina: 2024-01-02 07:00
```

Ese registro se asigna al periodo de inicio. No se reparte automáticamente entre las horas o días intermedios.

Para analizar sueño con máxima fidelidad:

```bat
--detalle todos
```

---

### 3. Fuentes duplicadas

Apple Salud puede mezclar datos de:

```text
Apple Watch
iPhone
Apps externas
Báscula inteligente
Otros dispositivos
```

El script agrupa por `fuente`, así que puede haber varias filas para el mismo periodo y tipo si los datos vienen de fuentes distintas.

---

### 4. Android no está soportado en este script

El README explica cómo exportar datos desde Android/Google, pero este script no los convierte. Para Android haría falta un script diferente.

---

## 🧯 Errores frecuentes

### No encuentra `export.xml`

Mensaje posible:

```text
No he encontrado export.xml en esta carpeta ni en subcarpetas.
```

Soluciones:

1. Comprueba que has descomprimido el ZIP de Apple Salud.
2. Ejecuta el script desde la carpeta correcta.
3. Usa `--xml` con la ruta exacta al archivo o a la carpeta.

```bat
python exportar_salud_completo.py --xml "C:\ruta\a\apple_health_export"
```

---

### Has pasado una carpeta y da error de permisos

En versiones antiguas podía aparecer:

```text
PermissionError: [Errno 13] Permission denied
```

Eso pasaba al pasar una carpeta en `--xml`. La versión actual acepta carpeta o archivo.

Comprueba que al ejecutar aparece:

```text
Versión script: 2026-07-08_formato_excel_v3_ruta_carpeta
```

---

### Fecha inválida

Usa siempre:

```text
YYYY-MM-DD
```

Correcto:

```bat
--inicio 2024-01-01
```

Incorrecto:

```bat
--inicio 01/01/2024
```

---

### El Excel sale demasiado grande

Reduce el rango o usa menos detalle.

| Problema | Solución |
|---|---|
| Excel muy pesado | Usa `--detalle dia` o acota fechas |
| Demasiadas filas de pulso | Usa `--detalle hora` en vez de `min` o `todos` |
| Necesitas auditar un día concreto | Usa `--detalle todos` solo para ese día |

---

### El Excel sale sin formato

Comprueba:

1. Que estás ejecutando el script nuevo.
2. Que el `.bat` y `exportar_salud_completo.py` están en la misma carpeta.
3. Que al inicio aparece:

```text
Versión script: 2026-07-08_formato_excel_v3_ruta_carpeta
```

Si no aparece, sustituye el script por la versión actual.

---

## 🧭 Flujo recomendado

### Paso 1 — Histórico general

Empieza con un resumen diario:

```bat
python exportar_salud_completo.py --inicio 2024-01-01 --fin 2024-12-31 --detalle dia --salida salud_2024_dia.xlsx
```

### Paso 2 — Periodo relevante

Si ves algo interesante, baja a detalle horario:

```bat
python exportar_salud_completo.py --inicio 2024-05-01 --fin 2024-05-31 --detalle hora --salida salud_2024_05_hora.xlsx
```

### Paso 3 — Día concreto

Para revisar un día en detalle:

```bat
python exportar_salud_completo.py --inicio 2024-05-10 --fin 2024-05-10 --detalle min --salida salud_2024_05_10_min.xlsx
```

### Paso 4 — Auditoría exacta

Para ver registros originales:

```bat
python exportar_salud_completo.py --inicio 2024-05-10 --fin 2024-05-10 --detalle todos --salida salud_2024_05_10_todos.xlsx
```

---

## 🔐 Privacidad

Este proceso trabaja **en local**. No sube tus datos de salud a ninguna web.

Aun así, trata los archivos generados como información sensible:

```text
export.xml
apple_health_temp.sqlite
*.xlsx
```

Recomendación: guárdalos en una ubicación privada y bórralos si ya no los necesitas.

---

## ✅ Comando recomendado para tu caso

Ejemplo con carpeta `apple_health_export`, rango largo y detalle horario:

```bat
python exportar_salud_completo.py --xml "C:\Users\juanc\OneDrive\Personal\salud\apple_health_export" --inicio 2024-06-01 --fin 2026-06-30 --detalle hora --salida salud_202606_v3.xlsx
```

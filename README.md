# Automatizacion Tekla -> Macro Allied (.xls + Docker + N8N)

Proyecto que automatiza el copiado de datos desde los archivos Excel `.xls` exportados por **Tekla Structures** a la **macro oficial `.xls` de Allied**, usando **N8N**, **Python** y **Excel COM/PowerShell**.

---

## ¿Qué hace?

Cada 5 minutos, N8N llama a un servidor HTTP que corre en el host. Ese servidor:

1. Revisa la carpeta de exportación de Tekla.
2. Usa Docker/Python para leer cada archivo Tekla `.xls` y armar un payload JSON.
3. Usa Excel COM via `powershell.exe` para escribir sobre la macro Allied `.xls`.
4. Genera el archivo final `[JOB_NUMBER]_Secondary_Shipper.xls`.
5. Deja el resultado en `data/output/` para que N8N lo adjunte al email.

---

## Estructura del proyecto

```
tekla_allied_docker/
├── docker-compose.yml         <- Orquestación Docker
├── Dockerfile.python          <- Imagen Python para leer los .xls de Tekla
├── n8n_workflow.json          <- Workflow listo para importar a N8N
├── requirements.txt           <- Dependencias Python
├── README.md                  <- Este archivo
│
├── scripts/
│   ├── export_tekla_payload.py <- Lee Tekla y genera JSON intermedio
│   ├── write_allied_xls.ps1    <- Escribe la macro Allied .xls con Excel COM
│   ├── run_xls_host.sh         <- Flujo host end-to-end .xls
│   └── xls_host_server.py      <- HTTP host para que N8N dispare el flujo .xls
│
└── data/
    ├── tekla/                 <- Tekla deja aquí los .xls
    │   └── procesados/        <- (se crea solo) archivos ya procesados
    ├── macro/
    │   └── Allied_Macro_original.xls <- Macro oficial Allied .xls
    └── output/                <- Archivos generados (.xls finales)
```

---

## Requisitos

En la maquina host:

- **Docker** y **Docker Compose** instalados.
  - Para verificar: `docker --version && docker compose version`
- **Microsoft Excel** instalado en Windows.
- `powershell.exe` accesible desde WSL.

Python para leer Tekla corre en Docker. La escritura del `.xls` final corre en el host porque depende de Excel COM.

---

## Instalación paso a paso

### 1. Descargar el proyecto

Copiar la carpeta `tekla_allied_docker/` a la máquina Linux. Por ejemplo en `/home/usuario/tekla_allied_docker/`.

### 2. Poner la macro Allied en su lugar

```bash
cp /ruta/a/tu/Allied_Macro_original.xls  ~/tekla_allied_docker/data/macro/
```

Verificar que el nombre del archivo sea exactamente `Allied_Macro_original.xls` o pasar otra ruta al ejecutar `scripts/run_xls_host.sh`.

### 3. Construir y arrancar el contenedor

```bash
cd ~/tekla_allied_docker
docker compose up -d --build
```

Lo que hace este comando:
- `up` → arranca el servicio.
- `-d` → en segundo plano (detached).
- `--build` → construye la imagen custom (N8N + Python). Solo la primera vez tarda unos minutos. Las siguientes veces no es necesario, basta con `docker compose up -d`.

### 4. Acceder a N8N

Abrir el navegador en: **http://localhost:5678**

La primera vez te va a pedir crear una cuenta de administrador (usuario, email, contraseña). Es solo para entrar a la interfaz, no afecta al workflow.

### 5. Importar el workflow

1. En N8N, click en **"Workflows"** → botón **"Import from File"** (esquina superior derecha).
2. Seleccionar el archivo `n8n_workflow.json` del proyecto.
3. Se cargarán los 6 nodos conectados.

### 6. Arrancar el servidor host .xls

En una terminal WSL, desde la raiz del proyecto:

```bash
python3 scripts/xls_host_server.py
```

Debe quedar escuchando en `http://0.0.0.0:5055`. El workflow N8N llama a `http://host.docker.internal:5055/run`.
Para comprobarlo en el navegador, abrir `http://localhost:5055/health`.

### 7. Configurar credenciales SMTP (para enviar email)

Antes de poder usar el nodo de email:

1. En N8N, ir a **"Credentials"** (menú izquierdo) → **"Add Credential"**.
2. Buscar **"SMTP"** y completar:
   - **User**: tu cuenta (ej: `automation@allied.com`)
   - **Password**: la contraseña SMTP
   - **Host**: el servidor SMTP (ej: `smtp.gmail.com`, `smtp.office365.com`)
   - **Port**: 587 (TLS) o 465 (SSL)
   - **SSL/TLS**: según el servidor
3. Guardar y darle un nombre, ej: `SMTP Allied`.

Luego, en el workflow:
- Click en el nodo **"Enviar Email al Equipo"**.
- En el campo **Credential to connect with**, seleccionar la credencial creada.
- Cambiar `fromEmail` y `toEmail` a las direcciones reales.

### 8. Probar el flujo manualmente

Antes de activar el trigger automático, probá manualmente:

1. Copiar un archivo de prueba a `data/tekla/`:
   ```bash
   cp ~/SBS_Eave_Struts_Shipper.xls ~/tekla_allied_docker/data/tekla/
   ```
2. En N8N, abrir el workflow y click en **"Execute Workflow"** (arriba).
3. Verificar:
   - Se genero un `.xls` en `data/output/`.
   - Llegó el email con el adjunto.

Tambien se puede probar sin N8N:

```bash
scripts/run_xls_host.sh data/tekla data/macro/Allied_Macro_original.xls data/output
```

### 9. Activar el workflow

Una vez que la prueba manual funciona, activá el toggle **"Active"** arriba a la derecha en el workflow. Desde ese momento, N8N corre el flujo cada 5 minutos automáticamente.

---

## Comandos útiles de Docker

```bash
# Ver los logs en vivo
docker compose logs -f n8n

# Reiniciar el contenedor
docker compose restart

# Parar el contenedor (no borra datos)
docker compose down

# Parar y borrar TODO (incluido el volumen con workflows y credenciales)
docker compose down -v

# Reconstruir después de cambiar el Dockerfile
docker compose up -d --build

# Entrar al contenedor para debug
docker compose exec n8n sh

# Probar solo la lectura de Tekla y generacion del payload
docker compose exec python-runner python3 /scripts/export_tekla_payload.py

# Probar el flujo completo .xls desde el host
scripts/run_xls_host.sh data/tekla data/macro/Allied_Macro_original.xls data/output
```

---

## Cómo está conectado el sistema

```
┌───────────────────────────── HOST (WSL/Windows) ────────────────────────┐
│                                                                         │
│  xls_host_server.py                                                     │
│       │                                                                 │
│       ├── docker compose exec python-runner export_tekla_payload.py      │
│       │                                                                 │
│       └── powershell.exe write_allied_xls.ps1 -> Excel COM -> .xls       │
│                                                                         │
│   ./data/tekla/   ./data/macro/   ./data/output/   ./scripts/           │
│         │              │               │              │                 │
│   ┌─────────────────────────── DOCKER ──────────────────────────────┐  │
│   │  Contenedor: n8n-tekla                                          │  │
│   │                                                                 │  │
│   │  /data/output                                                   │  │
│   │       │                                                         │  │
│   │                ┌──────────────────┐                             │  │
│   │                │       N8N        │                             │  │
│   │                │  workflow corre  │                             │  │
│   │                │   cada 5 min     │                             │  │
│   │                └──────────────────┘                             │  │
│   │                          │                                      │  │
│   │                          ▼                                      │  │
│   │           http://host.docker.internal:5055/run                  │  │
│   └─────────────────────────────────────────────────────────────────┘  │
│                              │                                          │
│                              ▼                                          │
│                    Email al equipo (SMTP)                               │
└─────────────────────────────────────────────────────────────────────────┘
```

---

## Comunicación entre el host .xls y N8N

El servidor host `scripts/xls_host_server.py` devuelve JSON al nodo HTTP de N8N.

Ejemplo:

```json
{
  "status": "success",
  "job_number": "J12345",
  "files_processed": 8,
  "files_with_errors": [],
  "output_file": "/data/output/J12345_Secondary_Shipper.xls",
  "host_output_file": "/home/usuario/tekla_allied_docker/data/output/J12345_Secondary_Shipper.xls",
  "duration_seconds": 4,
  "log_entries": ["OK [archivo.xls] -> [pestana] | N piezas | Peso: ..."]
}
```

Valores posibles de `status`:
- `success` → todo bien.
- `partial_success` → algunos archivos fallaron pero se generó el archivo.
- `no_files` → no había nada que procesar (el IF detiene el flujo).
- `error` / `fatal_error` → falló todo. Revisar `message`.

---

## Solución de problemas

**El contenedor no arranca**
```bash
docker compose logs n8n
```
Buscar el error específico.

**N8N arranca pero no encuentra Python**
- Verificar que la imagen se construyó con `--build`. Reconstruir:
  ```bash
  docker compose down
  docker compose up -d --build
  ```

**El nodo "Llamar Flujo XLS Host" da error**
- Verificar que el servidor host este corriendo:
  ```bash
  python3 scripts/xls_host_server.py
  ```
- Probar el flujo completo desde el host:
  ```bash
  scripts/run_xls_host.sh data/tekla data/macro/Allied_Macro_original.xls data/output
  ```

**El workflow se ejecuta pero no procesa archivos**
- Verificar que los archivos estén en `./data/tekla/` (en el host).
- Verificar permisos: el contenedor escribe como uid `node` (1000). Si la carpeta es de otro usuario:
  ```bash
  sudo chown -R 1000:1000 ./data
  ```

**El email no llega**
- Probar las credenciales SMTP con un workflow de prueba simple.
- Revisar la carpeta de Spam.
- Si usás Gmail, hay que generar una "App Password" (no funciona la contraseña normal con 2FA activado).

**Los archivos quedan con permisos raros en el host**
- Es porque el contenedor escribe como uid 1000. Si querés que tu usuario pueda editarlos:
  ```bash
  sudo chown -R $USER:$USER ./data/output
  ```

---

## Cómo cambiar la frecuencia del trigger

Por defecto corre cada 5 minutos. Para cambiarlo:

1. Abrir el workflow en N8N.
2. Click en el nodo **"Cada 5 minutos"**.
3. Cambiar el valor en **"Minutes Between Triggers"**.
4. Guardar.

Si querés que sea manual (sin trigger automático), cambiar el nodo Schedule por un **"Manual Trigger"**.

---

## Documentacion validada de la logica final

Esta seccion documenta la version validada contra la macro modelo. La logica de mapeo ya fue comparada contra `data/macro/Allied_Macro_original.xls` y el archivo generado quedo igual al modelo esperado para el job de prueba `403184`.

Regla principal del proyecto:

- Los archivos de Tekla llegan como `.xls`.
- La macro Allied original tambien es `.xls`.
- El archivo generado debe seguir siendo `.xls`.
- No se debe convertir el resultado final a `.xlsx`, porque la macro y su estructura dependen del formato original.

### Tecnologias utilizadas

| Tecnologia | Donde se usa | Definicion |
|------------|--------------|------------|
| Tekla Structures | Origen de datos | Software que exporta los reportes de shipper en archivos `.xls`. En este proyecto esos archivos se colocan en `data/tekla/`. |
| Python 3.11 | Lectura y transformacion | Lenguaje que lee los archivos Tekla, interpreta las pestañas y genera el payload JSON intermedio. |
| pandas | Lectura de `.xls` HTML | Tekla exporta algunos `.xls` como HTML disfrazado de Excel. `pandas.read_html` permite leer esas tablas correctamente. |
| openpyxl | Workbook intermedio | Se usa para representar en memoria las tablas leidas desde Tekla y trabajar con filas/columnas. |
| xlrd | Soporte `.xls` binario | Permite leer `.xls` reales tipo Excel 97-2003 cuando el archivo no viene como HTML. |
| Docker | Entorno de Python | Asegura que Python tenga siempre las dependencias correctas sin instalarlas directamente en Windows/WSL. |
| Docker Compose | Orquestacion local | Levanta `python-runner` y `n8n` con las carpetas compartidas correctas. |
| PowerShell | Puente hacia Excel | Ejecuta el script que abre la macro `.xls` usando Excel instalado en Windows. |
| Excel COM | Escritura final `.xls` | Automatiza Microsoft Excel para copiar el template, escribir datos y guardar un `.xls` real. |
| N8N | Automatizacion/orquestacion | Llama al servidor host cada 5 minutos, lee el archivo generado y puede enviarlo por email. |
| JSON | Contrato intermedio | `tekla_payload.json` contiene las piezas ya normalizadas antes de escribir la macro. |
| WSL2 | Entorno host | Linux dentro de Windows. Desde WSL se ejecutan Docker, Python host y PowerShell de Windows. |

### Mapa mental del flujo

```text
Tekla exporta .xls
    |
    v
data/tekla/
    |
    v
N8N o prueba manual llama:
    POST http://localhost:5055/run
    |
    v
scripts/xls_host_server.py
    |
    +--> scripts/run_xls_host.sh
            |
            +--> Docker / python-runner
            |       |
            |       +--> scripts/export_tekla_payload.py
            |               |
            |               +--> lee data/tekla/*.xls
            |               +--> usa reglas de mapeo
            |               +--> genera data/output/tekla_payload.json
            |
            +--> PowerShell en Windows
                    |
                    +--> scripts/write_allied_xls.ps1
                            |
                            +--> abre data/macro/Allied_Macro_original.xls
                            +--> escribe las pestañas con Excel COM
                            +--> guarda data/output/[JOB]_Secondary_Shipper.xls
```

### Carpetas y archivos auditados

| Ruta | Estado | Uso |
|------|--------|-----|
| `data/tekla/` | Necesaria | Entrada de archivos `.xls` exportados desde Tekla. |
| `data/macro/` | Necesaria | Contiene `Allied_Macro_original.xls`, template oficial. |
| `data/output/` | Necesaria | Salida del `.xls` generado y del `tekla_payload.json`. |
| `scripts/` | Necesaria | Contiene la logica de lectura, transformacion, servidor host y escritura `.xls`. |
| `docker-compose.yml` | Necesario | Define servicios `n8n` y `python-runner`. |
| `Dockerfile.python` | Necesario | Construye el contenedor Python con dependencias. |
| `requirements.txt` | Referencia local | No es la fuente principal del contenedor; las dependencias reales se instalan en `Dockerfile.python`. |
| `n8n_workflow.json` | Necesario si se usa N8N | Workflow importable en N8N. |
| `data/tekla/procesados/` | Historico/opcional | Contiene copias antiguas de pruebas. No se requiere para ejecutar el flujo actual, pero sirve como respaldo. |
| `data/output/403184_Secondary_Shipper.xls` | Generado | Archivo final de prueba validado. Se puede reemplazar en cada corrida. |
| `data/output/tekla_payload.json` | Generado | Payload intermedio; se regenera en cada corrida. |
| `scripts/__pycache__/` | Eliminado | Cache de Python. No afecta el proyecto. |
| `Dockerfile.old.backup` | Eliminado | Backup de una version anterior. No participa en el flujo validado. |
| `Recap.md` | Eliminado | Notas de trabajo. No participa en el flujo validado. |
| `Proceso.md` | Documentacion auxiliar | Puede mantenerse como bitacora, pero no ejecuta nada. |

Limpieza realizada: se eliminaron `scripts/__pycache__/`, `Dockerfile.old.backup` y `Recap.md`.

No se elimino `data/tekla/procesados/` porque contiene respaldos historicos de archivos `.xls`. No es necesaria para ejecutar el flujo actual, pero puede servir para trazabilidad o pruebas.

### JSON correcto para importar en N8N

El archivo que se debe subir/importar en N8N es:

```text
n8n_workflow.json
```

No se debe subir `data/output/tekla_payload.json` a N8N. Ese archivo es generado automaticamente por Python en cada corrida y solo sirve como contrato intermedio entre la lectura de Tekla y la escritura de la macro.

Si N8N corre con el `docker-compose.yml` de este proyecto, el nodo HTTP debe llamar:

```text
http://host.docker.internal:5055/run
```

Si N8N corre instalado directamente en Windows o WSL, el nodo HTTP puede llamar:

```text
http://localhost:5055/run
```

### Scripts principales y responsabilidad

| Script | Responsabilidad | Se debe conservar |
|--------|-----------------|-------------------|
| `scripts/export_tekla_payload.py` | Lee Tekla, aplica reglas de negocio y genera `tekla_payload.json`. | Si |
| `scripts/tekla_to_allied.py` | Libreria base: lectura de `.xls`, lectura de encabezados, transformaciones generales. | Si |
| `scripts/write_allied_xls.ps1` | Abre la macro `.xls` con Excel COM y escribe el archivo final. | Si |
| `scripts/run_xls_host.sh` | Ejecuta el flujo completo: Docker/Python + PowerShell/Excel. | Si |
| `scripts/xls_host_server.py` | Servidor host en puerto `5055`, llamado por N8N o `curl`. | Si |
| `scripts/server.py` | Servidor interno del contenedor para healthcheck/payload. | Si, porque Docker lo usa como proceso principal y healthcheck. |

### Flujo operativo paso a paso

1. Colocar los `.xls` exportados por Tekla en `data/tekla/`.
2. Confirmar que `data/macro/Allied_Macro_original.xls` existe.
3. Levantar Docker:
   ```bash
   docker compose up -d --build
   ```
4. Levantar el servidor host:
   ```bash
   python3 scripts/xls_host_server.py
   ```
5. Ejecutar manualmente:
   ```bash
   curl -X POST http://localhost:5055/run
   ```
6. Revisar el resultado:
   ```text
   data/output/[JOB]_Secondary_Shipper.xls
   ```

### Demo paso a paso

Esta demo muestra la ruta completa desde la descarga del archivo Tekla hasta la obtencion de la macro final.

1. Descargar/exportar desde Tekla los archivos shipper en formato `.xls`.
2. Copiar esos `.xls` en esta carpeta:
   ```text
   \\wsl.localhost\Ubuntu\home\eguerrero\Flujo_Automatizacion_Tekla\data\tekla
   ```
3. Verificar que la macro modelo exista en:
   ```text
   \\wsl.localhost\Ubuntu\home\eguerrero\Flujo_Automatizacion_Tekla\data\macro\Allied_Macro_original.xls
   ```
4. Abrir una terminal WSL en la raiz del proyecto:
   ```bash
   cd /home/eguerrero/Flujo_Automatizacion_Tekla
   ```
5. Levantar Docker:
   ```bash
   docker compose up -d --build
   ```
6. Levantar el servidor host `.xls`:
   ```bash
   python3 scripts/xls_host_server.py
   ```
7. Verificar en el navegador que el servidor este activo:
   ```text
   http://localhost:5055/health
   ```
   La respuesta esperada incluye `status: ok`.
8. Generar la macro final con una prueba manual:
   ```bash
   curl -X POST http://localhost:5055/run
   ```
9. Revisar el archivo generado en:
   ```text
   \\wsl.localhost\Ubuntu\home\eguerrero\Flujo_Automatizacion_Tekla\data\output
   ```
10. El archivo final debe tener un nombre como:
    ```text
    403184_Secondary_Shipper.xls
    ```
11. Para hacer la demo desde N8N, importar `n8n_workflow.json`, abrir el workflow y presionar `Execute Workflow`.
12. En la ejecucion de N8N, el nodo `Leer Archivo Generado` toma el `.xls` desde `data/output/` y lo deja disponible como binario para adjuntarlo o descargarlo.

Ruta mental de la demo:

```text
Descarga Tekla .xls
    -> data/tekla/
    -> POST /run desde curl o N8N
    -> Python lee y genera tekla_payload.json
    -> PowerShell/Excel COM escribe la macro
    -> data/output/[JOB]_Secondary_Shipper.xls
    -> N8N puede adjuntar/descargar el resultado
```

### Reglas de mapeo validadas

Estas reglas viven en `scripts/export_tekla_payload.py` y en las funciones base de `scripts/tekla_to_allied.py`.

#### Pestañas Tekla hacia pestañas Allied

| Archivo/Pestaña Tekla | Pestaña Allied |
|-----------------------|----------------|
| `SBS_Eave_Struts_Shipper` | `Eave Struts` |
| `SBS_CEE_Secondary_Shipper` | `Cold Form Members (CEE)` y parte hacia `Misc. Cold Form` |
| `SBS_ZEE_Secondary_Shipper` | `Cold Form Members (ZEE)`, `(ZEE) (2)`, `(ZEE) (3)` |
| `SBS_Miscellaneous_Shipper` | `Misc. Cold Form` |
| `SBS_Clips_Shipper` | `Clips` |
| `SBS_Pre_Galv_Clips_Shipper` | `Pre-Galv Clips` |
| `Standing_Seam_Hardware_Shipper` | `Standing Seam Hardware` |

#### Reglas generales de descripcion

| Si Python encuentra | Escribe en Allied |
|---------------------|-------------------|
| `EAVE_STRUT` | `Eave Strut (LSSS)` |
| `SHEETING_BASE_CHANNEL` | `Base Angle` |
| `SHEETING_ANGLE` | `Sheeting Angle` |
| `SHEETING_BASE_ANGLE` | `Base Angle` |
| `C_WRAP_CHANNEL` | `Girt Header (8 1/4CX6X4)` como base, luego regla especial en Misc |
| `C_STRUT_SPACER_LOW` | `Eave Strut Spacer  ( 3 : 12)` |
| `STRAPPING` | `Rolls of Strapping` |
| `CCF_CLIP` | `Clip` |
| `CCF_CL5` | `Sheeting Clip` |
| `CCF_CL103` | `Girt / Jamb Clip` |
| `CCF_CL104` | `Jamb Base Clip` |
| `CCF_CL100` | `Header to Jamb Clip` |
| `WALL GIRT` | `Wall Girt` |
| `ROOF PURLIN` | `Roof Purlin` |
| `FRAME OPENING HEADER` | `Frame Opening Header` |
| `FRAMED OPENING JAMB / SUB JAMB` | `Framed Opening Jamb / Sub Jamb` en CEE |

#### Cold Form Members (CEE)

| Si Python encuentra | Accion |
|---------------------|--------|
| `140BC*` | No se escribe en CEE; se mueve a `Misc. Cold Form`. |
| `140DH1` | Se escribe en CEE como `Frame Opening Header`. |
| `140DJ*` | Se escribe en CEE como `Framed Opening Jamb / Sub Jamb`. |
| Fila `Web=` asociada a `140DJ*` | Se mantiene como fila detalle debajo de su pieza. |
| Peso de pagina | Se calcula con las piezas CEE restantes, sin incluir `140BC*`. Resultado validado del job: `446.25`. |

#### Cold Form Members (ZEE)

| Si Python encuentra | Accion |
|---------------------|--------|
| Archivo ZEE con muchas filas | Se divide respetando las tres pestañas Allied: `ZEE`, `ZEE (2)`, `ZEE (3)`. |
| `140G1` a `140G18` | Van a `Cold Form Members (ZEE)`. |
| `140G19` a `140P4` | Van a `Cold Form Members (ZEE) (2)`. |
| `140P5`, `140P_EXT1`, `140P_EXT2`, `140P_EXT3` | Van a `Cold Form Members (ZEE) (3)`. |
| `140G10` con `Web=321.75` | Se elimina ese detalle porque la macro modelo no lo usa. |
| `140G28` con `Web=45.75` | Se elimina ese detalle porque la macro modelo no lo usa. |
| `140G3` detalle | Se escribe exactamente `Web=      22.1875"       57.75 66"`. |
| `140G21` detalle | Se completa como `Web=9.5 13.5 18.1875 61.3125`. |
| `140P2` detalle | Se reemplaza por `Web=45.75 295.75 299.75 303.75 307.75 315.75 319.75 323.75 333.75 345.75 353.75`. |
| `140P3` detalle | Se reemplaza por `Web=13.75 21.75 33.75 43.75 47.75 51.75 59.75 63.75 67.75 71.75 321.75 333.75 337.75`. |
| `140P4` detalle | Se reemplaza por `Web=29.75 33.75 323.50 327.50 353.50 357.50`. |
| `140P5` detalle | Se reemplaza por `Web=10.00 14.00 40.00 44.00`. |
| `140P5` en columna `PART` | Se escribe `2`. |
| `140P_EXT1` | En `MARK` se escribe `T1`. |
| `140P_EXT2` | En `MARK` se escribe `T2` y en `PART` se escribe `2`. |
| `140P_EXT3` | En `MARK` se escribe `T3`. |
| Filas detalle `Web=` largas | Se fusionan de columna A a N y se alinean a la izquierda para que se vea el inicio completo del texto. |

#### Misc. Cold Form

| Si Python encuentra | Accion |
|---------------------|--------|
| `140BC*` desde CEE | Se mueve a `Misc. Cold Form` como `Base Angle`. |
| `140BC_EXT1` | Se escribe como `T1` en columna `MARK`. |
| `140BC5` | QTY se fuerza a `1`. |
| `140BC*` normal | `PART=8X25C16`, `DWG #=BA1`, `COLOR=Pre-Galvanized`, `LENGTH=20'- 0"`. |
| `140BC_EXT1` | `LENGTH=10'- 0"`, `WT.=27.95`. |
| `140SA1` | QTY entero `5`, `DESC=Sheeting Angle`, `PART=4X2X16Ga`, `DWG #=SA1`. |
| `140SA2` | QTY entero `4`, `DESC=Sheeting Angle`, `PART=4X2X16Ga`, `DWG #=SA1`. |
| `140SA_EXT1` | En `MARK` se escribe `T1`, QTY entero `1`. |
| `140GH1` | `DESC=8" Girt Header (8 1/4CX6X4)`, `PART=17 7/8X14Ga`, `DWG #=GH-1`. |
| `140SSL1` | `DESC=Eave Strut Spacer  ( 3 : 12)`, `PART=10X35C14`, `DWG #=SSL-10`, `LENGTH=6"`. |
| Orden final desde fila 21 | `140SA1`, `140SA2`, `T1`, `140GH1`, `140SSL1`. |
| Peso de pagina | Resultado validado del job: `302.44`. |

#### Clips

| Si Python encuentra | Accion |
|---------------------|--------|
| Color vacio o `0` | Se escribe `Pre-Galvanized`. |
| `CCF_CLIP` | `DESC=Clip`. |
| `CCF_CL5` o `140CCF7` | `DESC=Sheeting Clip`, `DRAWING #=CL-5`. |
| Otros clips normales | `DRAWING #` conserva el valor del mark/drawing original. |

#### Pre-Galv Clips

| Si Python encuentra | Accion |
|---------------------|--------|
| Color vacio o `0` | Se escribe `Pre-Galvanized`. |
| `CCF_CL103` | `DESC=Girt / Jamb Clip`, `DRAWING #=CL103`. |
| `CCF_CL104` | `DESC=Jamb Base Clip`, `DRAWING #=CL104`. |
| `CCF_CL100` | `DESC=Header to Jamb Clip`, `DRAWING #=CL100`. |
| Faltan extensiones | Se agregan `140CCF_EXT1`, `140CCF_EXT2`, `140CCF_EXT3`. |
| `140CCF_EXT1` | QTY `2`, `DRAWING #=CL103`, `WT.=4.3`. |
| `140CCF_EXT2` | QTY `2`, `DRAWING #=CL104`, `WT.=2.7`. |
| `140CCF_EXT3` | QTY `2`, `DRAWING #=CL100`, `WT.=2.7`. |
| Peso de pagina | Resultado validado del job: `99.41`. |

#### Standing Seam Hardware

| Si Python encuentra | Accion |
|---------------------|--------|
| `STRAP` o `STRAPPING` | Se escribe `MARK=STRP1`, `DESC=Rolls of Strapping`, `DWG #=STRP1`, `COLOR=White`. |
| QTY `4` de straps | `WT.=88.00`. |

### Auditoria de logica

Resultado de la auditoria funcional:

- El flujo conserva `.xls` de entrada y `.xls` de salida.
- La macro original no se sobrescribe; siempre se copia a un archivo nuevo en `data/output/`.
- El payload se regenera en cada corrida y contiene las reglas ya normalizadas.
- Las pestañas validadas contra la macro modelo son:
  - `Cold Form Members (CEE)`
  - `Cold Form Members (ZEE)`
  - `Cold Form Members (ZEE) (2)`
  - `Cold Form Members (ZEE) (3)`
  - `Misc. Cold Form`
  - `Clips`
  - `Pre-Galv Clips`
  - `Standing Seam Hardware`
  - `Screws`, reportada como correcta durante validacion.
- Si Excel tiene abierto el archivo final, Windows bloquea la escritura. En ese caso se debe cerrar el `.xls` y volver a ejecutar `curl -X POST http://localhost:5055/run`.
- No se debe borrar `scripts/tekla_to_allied.py`, aunque ya no sea el entrypoint principal; `export_tekla_payload.py` lo usa como libreria.
- No se debe borrar `scripts/server.py`; Docker lo usa para mantener vivo `python-runner` y responder el healthcheck.

### Comandos finales de operacion

Arrancar servidor host:

```bash
python3 scripts/xls_host_server.py
```

Probar salud:

```bash
curl http://localhost:5055/health
```

Generar archivo:

```bash
curl -X POST http://localhost:5055/run
```

Generar sin N8N ni servidor HTTP:

```bash
scripts/run_xls_host.sh data/tekla data/macro/Allied_Macro_original.xls data/output
```

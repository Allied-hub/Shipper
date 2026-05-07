# Automatización Tekla → Macro Allied (Linux + Docker + N8N)

Proyecto que automatiza el copiado de datos desde los archivos Excel exportados por **Tekla Structures** a la **macro oficial de Allied**, usando **N8N** + **Python** (sin IA), todo corriendo en **Docker sobre Linux**.

---

## ¿Qué hace?

Cada 5 minutos, N8N ejecuta un script Python dentro del contenedor que:

1. Revisa la carpeta de exportación de Tekla.
2. Lee cada archivo `.xlsx` y aplica las transformaciones obligatorias.
3. Escribe los datos en la macro oficial Allied (preservando macros VBA).
4. Genera el archivo final `[JOB_NUMBER]_Secondary_Shipper.xlsm`.
5. Mueve los archivos procesados a `procesados/<timestamp>/` para no reprocesarlos.
6. Envía el resultado por email al equipo (con el `.xlsm` adjunto).

---

## Estructura del proyecto

```
tekla_allied_docker/
├── docker-compose.yml         <- Orquestación Docker
├── Dockerfile                 <- Imagen N8N + Python + openpyxl
├── n8n_workflow.json          <- Workflow listo para importar a N8N
├── requirements.txt           <- Dependencias Python
├── README.md                  <- Este archivo
│
├── scripts/
│   └── tekla_to_allied.py     <- Script de transformación
│
└── data/
    ├── tekla/                 <- Tekla deja aquí los .xlsx
    │   └── procesados/        <- (se crea solo) archivos ya procesados
    ├── macro/
    │   └── Allied_Macro.xlsm  <- Macro oficial Allied (poner aquí)
    └── output/                <- Archivos generados (.xlsm finales)
```

---

## Requisitos

En la máquina Linux:

- **Docker** y **Docker Compose** instalados.
  - Para verificar: `docker --version && docker compose version`
  - Si no los tenés: https://docs.docker.com/engine/install/

Eso es todo. Python, N8N y openpyxl quedan dentro del contenedor.

---

## Instalación paso a paso

### 1. Descargar el proyecto

Copiar la carpeta `tekla_allied_docker/` a la máquina Linux. Por ejemplo en `/home/usuario/tekla_allied_docker/`.

### 2. Poner la macro Allied en su lugar

```bash
cp /ruta/a/tu/Allied_Macro.xlsm  ~/tekla_allied_docker/data/macro/
```

Verificar que el nombre del archivo sea exactamente `Allied_Macro.xlsm` (o cambiar la variable `MACRO_PATH` en `docker-compose.yml` si querés otro nombre).

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

### 6. Configurar credenciales SMTP (para enviar email)

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

### 7. Probar el flujo manualmente

Antes de activar el trigger automático, probá manualmente:

1. Copiar un archivo de prueba a `data/tekla/`:
   ```bash
   cp ~/SBS_Eave_Struts_Shipper.xlsx ~/tekla_allied_docker/data/tekla/
   ```
2. En N8N, abrir el workflow y click en **"Execute Workflow"** (arriba).
3. Verificar:
   - El archivo se movió a `data/tekla/procesados/<timestamp>/`.
   - Se generó un `.xlsm` en `data/output/`.
   - Llegó el email con el adjunto.

### 8. Activar el workflow

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

# Probar el script Python desde dentro del contenedor
docker compose exec n8n python3 /scripts/tekla_to_allied.py
```

---

## Cómo está conectado el sistema

```
┌───────────────────────────── HOST (Linux) ──────────────────────────────┐
│                                                                         │
│   ./data/tekla/   ./data/macro/   ./data/output/   ./scripts/           │
│         │              │               │              │                 │
│         │              │               │              │                 │
│         ▼              ▼               ▼              ▼                 │
│   ┌─────────────────────────── DOCKER ──────────────────────────────┐  │
│   │  Contenedor: n8n-tekla                                          │  │
│   │                                                                 │  │
│   │  /data/tekla    /data/macro   /data/output   /scripts           │  │
│   │       │              │             │              │             │  │
│   │       └──────────────┴─────────────┴──────────────┘             │  │
│   │                          │                                      │  │
│   │                          ▼                                      │  │
│   │                ┌──────────────────┐                             │  │
│   │                │       N8N        │                             │  │
│   │                │  workflow corre  │                             │  │
│   │                │   cada 5 min     │                             │  │
│   │                └──────────────────┘                             │  │
│   │                          │                                      │  │
│   │                          ▼                                      │  │
│   │                  python3 /scripts/                              │  │
│   │                   tekla_to_allied.py                            │  │
│   └─────────────────────────────────────────────────────────────────┘  │
│                              │                                          │
│                              ▼                                          │
│                    Email al equipo (SMTP)                               │
└─────────────────────────────────────────────────────────────────────────┘
```

---

## Comunicación entre el script y N8N

- **stderr** → logs informativos (visibles en el log de ejecución de N8N).
- **stdout** → una sola línea JSON con el resumen, que el nodo "Parsear Resultado" convierte en datos N8N.

Ejemplo de JSON que imprime el script:

```json
{
  "status": "success",
  "job_number": "J12345",
  "files_processed": 8,
  "files_with_errors": [],
  "output_file": "/data/output/J12345_Secondary_Shipper.xlsm",
  "archive_folder": "/data/tekla/procesados/20260507_153022",
  "duration_seconds": 4,
  "log_entries": ["OK [archivo.xlsx] -> [pestana] | N piezas | Peso: ..."]
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

**El nodo "Ejecutar Script Python" da error**
- Probar el script desde dentro del contenedor:
  ```bash
  docker compose exec n8n python3 /scripts/tekla_to_allied.py
  ```
- Si el script tira error ahí, revisar la salida.

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

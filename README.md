# Word Section Transfer

Aplicación web para transferir contenido entre documentos Word (`.docx`) mediante la selección de secciones.

Permite visualizar un documento de origen **A** y un documento de destino **B**, preparar múltiples reemplazos y generar una nueva versión de B con todos los cambios.

El contenido insertado adopta los estilos del documento de destino y los archivos originales no se sobrescriben.

## Funcionalidades

### Carga y visualización
- Selección y carga de documentos `.docx`.
- Previsualización de ambos documentos en paneles independientes.
- Detección de secciones mediante títulos y niveles de esquema de Word.
- Navegación hasta el título seleccionado y resaltado en la previsualización.

### Transferencia de contenido
- Reemplazo individual mediante la API.
- Reemplazo de múltiples secciones desde la interfaz.
- Lista de parejas de origen y destino, con opción para quitar elementos.
- Confirmación antes de aplicar un lote.
- Conservación del título de cada sección de destino.
- Reemplazo del contenido de destino, incluidas sus subsecciones.
- Generación de un único documento con todos los cambios.
- Posibilidad de aplicar nuevos lotes sobre el B actualizado.

### Contenido soportado
- Texto y párrafos.
- Subtítulos con ajuste relativo de nivel.
- Negritas y cursivas aplicadas directamente al texto.
- Listas con viñetas y numeración decimal, incluidos niveles anidados.
- Imágenes en línea.
- Tablas rectangulares.
- Celdas combinadas horizontal y verticalmente.
- Texto, listas e imágenes en línea dentro de las celdas.

### Validaciones y resultados
- Bloqueo de destinos repetidos o superpuestos.
- Protección de secciones llamadas `Contents` y `Table of Contents`.
- Rechazo de elementos todavía no soportados, con mensajes de diagnóstico.
- Previsualización y descarga de B actualizado.
- Resumen por lote y por pareja:
  - Secciones reemplazadas.
  - Párrafos copiados.
  - Tablas copiadas.
  - Imágenes copiadas.
  - Listas creadas.

## Tecnologías

| Componente | Tecnología |
|---|---|
| Backend | Python y FastAPI |
| Servidor de desarrollo | Uvicorn |
| Procesamiento de Word | python-docx y lxml |
| Frontend | React y TypeScript |
| Herramienta de desarrollo frontend | Vite |
| Previsualización DOCX | docx-preview |
| Control de versiones | Git y GitHub |

## Requisitos

- Windows 11 como entorno de desarrollo utilizado.
- Python compatible con las dependencias de `backend/requirements.txt`.
- Node.js compatible con la versión de Vite del proyecto.
- npm, incluido con Node.js.
- Visual Studio Code, recomendado.
- Git o GitHub Desktop para obtener y gestionar el repositorio.
- Microsoft Word, recomendado para validar los documentos generados.

El procesamiento no necesita Microsoft Word instalado, pero la previsualización del navegador puede diferir de la presentación final en Word.

Las versiones de las dependencias están definidas en:

- `backend/requirements.txt`
- `frontend/package.json`
- `frontend/package-lock.json`

## Estructura del proyecto

```text
word-section-transfer/
├── backend/
│   ├── main.py
│   ├── transfer_api.py
│   ├── batch_transfer_api.py
│   ├── list_support.py
│   ├── image_support.py
│   ├── table_support.py
│   ├── transfer_summary.py
│   └── requirements.txt
├── frontend/
│   ├── src/
│   │   ├── App.tsx
│   │   ├── App.css
│   │   ├── DocumentSections.tsx
│   │   ├── DocumentPreview.tsx
│   │   ├── types.ts
│   │   └── ...
│   ├── package.json
│   └── package-lock.json
├── .gitignore
├── .gitattributes
└── README.md
```

## Instalación

Los siguientes comandos están preparados para **PowerShell en Windows**.

### 1. Obtener el proyecto

Con Git:

```powershell
git clone https://github.com/yeimiJuarez/word-section-transfer.git
cd word-section-transfer
```



También puedes usar GitHub Desktop:

1. Selecciona `File → Clone repository`.
2. Elige el repositorio.
3. Selecciona una ubicación local.
4. Abre la carpeta del proyecto en VS Code.

Si el repositorio es privado, necesitas una cuenta con acceso.

### 2. Preparar el backend

Desde la raíz del proyecto:

```powershell
cd backend
py -m venv .venv
.\.venv\Scripts\python.exe -m pip install --upgrade pip
.\.venv\Scripts\python.exe -m pip install -r requirements.txt
```

Los comandos usan directamente el Python del entorno virtual. No es necesario activar el entorno ni modificar la política de ejecución de PowerShell.

En VS Code puedes seleccionar este intérprete mediante:

`Ctrl + Shift + P → Python: Select Interpreter`

Ruta:

```text
backend\.venv\Scripts\python.exe
```

### 3. Preparar el frontend

Desde la raíz del proyecto:

```powershell
cd frontend
npm.cmd ci
```

`npm.cmd ci` instala las dependencias usando `package-lock.json` y requiere que ese archivo esté sincronizado con `package.json`.

En PowerShell se utiliza `npm.cmd` para evitar posibles restricciones de ejecución de `npm.ps1`.

## Ejecución local

Se necesitan **dos terminales abiertas**: una para el backend y otra para el frontend.

### Terminal 1: backend

Desde la carpeta `backend`:

```powershell
.\.venv\Scripts\python.exe -m uvicorn main:app --reload --host 127.0.0.1 --port 8000
```

Direcciones:

- Comprobación del servicio: http://127.0.0.1:8000/health
- Documentación interactiva: http://127.0.0.1:8000/docs

Respuesta esperada de `/health`:

```json
{
  "status": "ok",
  "message": "Backend funcionando"
}
```

La dirección raíz `/` puede mostrar `Not Found`; no se ha definido una página para esa ruta.

### Terminal 2: frontend

Desde la carpeta `frontend`:

```powershell
npm.cmd run dev
```

Abre la dirección que muestre Vite, normalmente:

http://localhost:5173/

Mantén ambas terminales abiertas mientras utilizas la aplicación.

Para detener cualquiera de los servicios, presiona `Ctrl + C` en su terminal.

### Puertos y conexión

El frontend utiliza actualmente:

```text
http://127.0.0.1:8000
```

como dirección de la API.

El backend permite solicitudes desde:

```text
http://localhost:5173
http://127.0.0.1:5173
```

Si cambias el puerto del backend, actualiza la dirección de la API en el frontend.

Si Vite utiliza otro puerto, actualiza los orígenes permitidos en la configuración CORS de `backend/main.py`.

## Uso de la aplicación

### 1. Seleccionar documentos

- En el panel A, selecciona el documento de origen.
- En el panel B, selecciona el documento de destino.

La previsualización se genera localmente en el navegador al seleccionar cada archivo.

### 2. Cargar documentos

Presiona:

- `Cargar documento A`
- `Cargar documento B`

La API recibe los archivos y devuelve sus secciones.

### 3. Preparar reemplazos

1. Selecciona una sección de A.
2. Selecciona una sección de B.
3. Presiona `Agregar a la lista`.
4. Repite para las demás parejas.

Puedes reutilizar una sección de A en distintos destinos.

No puedes repetir un destino ni seleccionar destinos superpuestos, como una sección y una subsección dentro de ella.

### 4. Aplicar el lote

Presiona `Aplicar todos los cambios` y revisa la confirmación.

El backend valida las parejas y trabaja sobre una única copia de B. Si una operación falla, no se guarda un resultado parcialmente actualizado.

### 5. Revisar y descargar

Después de una operación exitosa:

- El panel B muestra el documento actualizado.
- Sus secciones se recalculan.
- La lista pendiente queda vacía.
- Se muestra el resumen del lote.
- Puedes usar `Descargar B actualizado`.

Abre el DOCX en Word para validar el resultado final.

## Reglas de secciones y formato

### Límites de una sección

Una sección comienza en su título y termina antes del siguiente título del mismo nivel o de un nivel superior.

Seleccionar una sección padre incluye sus subsecciones.

### Títulos y texto

- El título principal de origen no se copia.
- El título de destino se conserva.
- Los subtítulos internos ajustan su nivel respecto al destino.
- El cuerpo utiliza el estilo `Normal` de B.
- Las fuentes, tamaños y colores directos de A no se trasladan.
- Se conservan negritas y cursivas directas.

### Listas

- Se crean listas reales de Word.
- Las viñetas usan símbolos predeterminados por nivel.
- La numeración admitida se normaliza a números decimales por nivel.
- No se conserva todavía numeración compuesta como `1.1.1`.
- Los párrafos consecutivos con el mismo identificador de lista de origen forman un grupo.
- Un párrafo normal, un título o una tabla termina ese grupo.
- Cada celda administra sus listas de manera independiente.

### Imágenes

- Se admiten imágenes incrustadas en línea.
- Se conserva su proporción.
- Las imágenes grandes se reducen al ancho disponible.
- Dentro de celdas se consideran el ancho combinado, los márgenes interiores y el espacio de las listas.
- No se reproducen todos los efectos visuales de A.

### Tablas

- Se reconstruyen utilizando el estilo `Table Grid` de B.
- Las columnas de la cuadrícula tienen anchos uniformes.
- Se conservan combinaciones rectangulares horizontales y verticales.
- No se trasladan los bordes, colores, anchos ni alturas manuales de A.

B debe contener los estilos necesarios. Si falta `Table Grid`, aplícalo a una tabla de prueba en B, guarda el documento y vuelve a cargarlo.

### Elementos normalizados

- Los hipervínculos se convierten en texto no clicable.
- Los marcadores de A no se trasladan.
- Las marcas de revisión ortográfica y paginación previa se omiten.
- Los saltos de página explícitos de A se eliminan.
- La paginación y el espaciado pueden diferir respecto a A.

## Límites actuales

- Archivos de entrada de hasta **10 MB**.
- Hasta **10 documentos almacenados por ejecución**, incluidos los resultados.
- Hasta **50 parejas por lote**.
- Almacenamiento temporal en memoria.
- Detección de títulos del cuerpo principal, no de encabezados, pies ni cuadros de texto.
- Sin soporte para tablas anidadas.
- Sin soporte para imágenes flotantes, externas, recortadas o rotadas.
- Algunos campos, revisiones y estructuras especiales se rechazan.
- Algunas definiciones de listas vinculadas a estilos y formatos no decimales no están soportadas.
- Combinaciones horizontales antiguas `w:hMerge` se rechazan.
- La navegación visual relaciona títulos mediante texto y orden; títulos repetidos en índices pueden causar asociaciones incorrectas.
- La tabla de contenido no se actualiza automáticamente después de los reemplazos.

La aplicación sigue siendo un prototipo local. La validación del DOCX generado en Word es parte del flujo de pruebas.

## Almacenamiento y reinicios

Los documentos cargados y generados están en la memoria del backend.

Por lo tanto:

- Reiniciar Uvicorn elimina los documentos y sus identificadores.
- Guardar código Python con `--reload` también puede reiniciar la aplicación.
- Después de un reinicio debes volver a cargar A y B.
- Recargar la página del navegador elimina las selecciones y la lista pendiente.
- Descargar un resultado permite conservarlo fuera de la aplicación.

Para recuperar el trabajo desde un resultado descargado, selecciónalo como B y cárgalo nuevamente.

## API principal

| Método | Ruta | Función |
|---|---|---|
| GET | `/health` | Comprobar el servicio |
| POST | `/documents` | Cargar un DOCX |
| GET | `/documents/{document_id}/sections` | Consultar secciones |
| POST | `/transfers` | Reemplazar una sección |
| POST | `/transfers/batch` | Reemplazar múltiples secciones |
| GET | `/documents/{document_id}/download` | Descargar un documento |

Ejemplo de transferencia por lote:

```json
{
  "source_document_id": "IDENTIFICADOR_DE_A",
  "target_document_id": "IDENTIFICADOR_DE_B",
  "replacements": [
    {
      "source_section_id": "section-10",
      "target_section_id": "section-20"
    },
    {
      "source_section_id": "section-30",
      "target_section_id": "section-40"
    }
  ]
}
```

Usa los identificadores devueltos por la API para los documentos cargados. Los identificadores de sección pueden cambiar después de editar o reemplazar contenido.

## Solución de problemas

### El frontend muestra ERR_CONNECTION_REFUSED

Comprueba que Vite esté ejecutándose y abre la dirección exacta que muestra la terminal.

### PowerShell bloquea npm.ps1

Usa `npm.cmd` en lugar de `npm`.

### La API devuelve 404 para un documento

El identificador puede ser incorrecto o el backend puede haberse reiniciado. Vuelve a cargar los documentos.

### La API devuelve 422

Comprueba el JSON y el endpoint:

- `/transfers` recibe una pareja.
- `/transfers/batch` recibe una lista `replacements`.

### La API devuelve 400

Revisa `detail`. Puede indicar contenido no soportado, destinos superpuestos, una sección protegida o estilos faltantes.

### Se alcanzó el límite de documentos

Descarga los resultados que quieras conservar, reinicia el backend y vuelve a cargar los archivos necesarios.

### La previsualización difiere de Word

Es una limitación del renderizado en navegador. Revisa el DOCX descargado en Microsoft Word.

## Seguridad y control de versiones

- Ejecuta el backend localmente con `127.0.0.1`.
- Usa documentos confiables durante las pruebas.
- No publiques credenciales, tokens, archivos `.env` ni documentos confidenciales.
- No subas `.venv`, `node_modules` ni resultados de pruebas.
- Revisa los archivos antes de cada commit.

El `.gitignore` no elimina del historial los archivos ya registrados por Git.

La aplicación no incluye autenticación, aislamiento multiusuario ni controles suficientes para exposición pública. Antes de desplegarla fuera del entorno local se requiere revisar la seguridad, el almacenamiento y los límites de recursos.

## Mejoras pendientes

- Persistencia temporal con caducidad.
- Pruebas automatizadas con documentos de referencia.
- Mayor fidelidad de espaciado y saltos.
- Formatos adicionales de numeración.
- Tablas anidadas y estructuras especiales.
- Navegación visual más robusta.
- Actualización de campos y tabla de contenido.
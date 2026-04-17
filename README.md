<<<<<<< HEAD
# 📄 Sistema Integral de Cartera y Facturación - JOSÉ A Y GERARDO E ZULUAGA S.A.S.

## 📖 Descripción del Proyecto
Este sistema es una aplicación web local enfocada en automatizar de manera sencilla y eficiente la administración de la facturación en mora, la interacción rápida y la generación masiva de cobros de clientes de Zuluaga Hermanos. 
Con un solo clic, se conecta directamente a la cartera en la base de datos SQL Server, cruza las facturas morosas de un usuario particular, genera un documento PDF profesional respetando su imagen corporativa y los remite de forma automática vía correo electrónico Outlook. 

## 🚀 Funcionalidades Principales
*   **Gestión Visual:** Tablero web web interactivo (FastAPI) donde se lista y filtra la cartera de clientes.
*   **Decisión Dinámica:** El sistema detecta automáticamente la presencia del `correo` del cliente en la base de datos. Si existe, redactará y mandará el correo electrónico anexándole el extracto y el cuerpo del mensaje. Si no tiene, abrirá el archivo PDF en el computador para que lo mandes como prefieras o realizar una impresión manual.
*   **Prevención de Bloqueos (Lock-File):** Lógica que detecta y evade errores por PDFs previamente abiertos en lectores externos como Adobe Acrobat o Chrome creando versiones incrementales protegidas evitando bloquear el sistema.
*   **Apariencia Corporativa (PDF):** Inserción precisa de logotipos, firmas manuscritas, redacción jurídica sobre las Leyes del Habeas Data Colombiano, Datacredito y simulación de hipervínculos azules (`#0033cc`).

---

## 🛠️ Tecnologías Empleadas
*   `Python 3.10+` (Backend Principal)
*   `FastAPI` & `Uvicorn` (Servidor web asíncrono ASGI)
*   `FPDF2` (Generación de reportería y estructura milimétrica en PDF y HTML)
*   `SQLAlchemy` & `PyODBC` / `PyMSSQL` (ORM y capas de conexión al SQL Server local).
*   `Jinja2` (Procesador de plantillas interfaces y vistas dinámicas `cliente.html`).

---

## ⚙️ Pasos de Instalación Rápida y Limpia

### 1. Preparación del Entorno (Librerías Code Base)
Abre una terminal normal (PowerShell o CMD) de Windows y posiciónate en el directorio raíz del proyecto (`interfaz_cartera`):
```bash
# 1. Crear el entorno virtual protegido para aislar descargas del computador base.
python -m venv venv

# 2. Activar el entorno virtual 
# Ejecuta en Windows CMD o PowerShell:
venv\Scripts\activate

# 3. Instalación de librerías y dependencias
pip install -r requirements.txt
```

### 2. Configuración Secreta del Proyecto (`.env`)
El núcleo del proyecto utiliza variables de entorno invisibles y que **NO** están adjuntadas por seguridad. Para configurar conexiones y envíos crea un nuevo documento temporal pero renómbralo estrictamente la extensión y cuerpo a la palabra `.env` y rellénala con lo siguiente:
```env
# Credenciales Servidor Base de Datos SQL SERVER Interfaz Local
DB_USER=desarrollojosea
DB_PASS=Djosea01*
DB_HOST=192.168.1.14
DB_PORT=1433
DB_NAME=cartera_db
DB_ENCRYPT=yes
DB_TRUSTSERVERCERT=yes

# Credenciales cuenta SMTP Corporativa de Envío de Mensajes (Outlook)
SMTP_USER=credito_cartera@josegera.com
SMTP_PASS=Colombia2024.08
SMTP_HOST=smtp-mail.outlook.com
SMTP_PORT=587
```

### 3. Ejecución del Sistema Local (Puesta en marcha)
Existen un par de modalidades muy buenas para arrancar tu app dependiendo de si eres desarrollador probando, o trabajador operando de forma final.
*   **Modo Operador Directo (Recomendado Terminal Windows):** Simplemente haz doble clic sobre el ícono prefabricado del proyecto: `start_server.bat` y mágicamente salpicara el navegador corriendo por completo en el puerto local predeterminado de FastAPI.
*   **Modo Desarrollo En VIVO:** Estando en la terminal y solo si el entorno `venv` virtual está encendido de antemano usa: `uvicorn app:app --reload`. (Esto habilita el detector de salvado automático, de manera de que cualquier letra de Python que remplaces recargara velozmente el documento y te ahorrará abrir la consola una docena de veces).

---

## 🗂️ Control de Estructuras Obligatorias
Para que el creador de cartas no reviente sus flujos con errores, asegúrate de no mover ni renombrar las identidades de estas carpetas:
- **`assets/img/logo.png`**: Representará su imagen logotipo original translúcido sin fondos en la cabecera.
- **`assets/img/firma_sebastian.jpeg`**: Debe incluir el trozo escaneado de la firma oficial de la persona remitente incluyendo el cargo, teléfono y nombre debajo.
- **`cartas_generadas/`**: Carpeta temporal invisible en la que los documentos listos para cobrar se estancan como copias caché o por enviar (Puedes limpiarla mes a mes).
=======
# 📄 Sistema Integral de Cartera y Facturación - JOSÉ A Y GERARDO E ZULUAGA S.A.S.

## 📖 Descripción del Proyecto
Este sistema es una aplicación web local enfocada en automatizar de manera sencilla y eficiente la administración de la facturación en mora, la interacción rápida y la generación masiva de cobros de clientes de Zuluaga Hermanos. 
Con un solo clic, se conecta directamente a la cartera en la base de datos SQL Server, cruza las facturas morosas de un usuario particular, genera un documento PDF profesional respetando su imagen corporativa y los remite de forma automática vía correo electrónico Outlook. 

## 🚀 Funcionalidades Principales
*   **Gestión Visual:** Tablero web web interactivo (FastAPI) donde se lista y filtra la cartera de clientes.
*   **Decisión Dinámica:** El sistema detecta automáticamente la presencia del `correo` del cliente en la base de datos. Si existe, redactará y mandará el correo electrónico anexándole el extracto y el cuerpo del mensaje. Si no tiene, abrirá el archivo PDF en el computador para que lo mandes como prefieras o realizar una impresión manual.
*   **Prevención de Bloqueos (Lock-File):** Lógica que detecta y evade errores por PDFs previamente abiertos en lectores externos como Adobe Acrobat o Chrome creando versiones incrementales protegidas evitando bloquear el sistema.
*   **Apariencia Corporativa (PDF):** Inserción precisa de logotipos, firmas manuscritas, redacción jurídica sobre las Leyes del Habeas Data Colombiano, Datacredito y simulación de hipervínculos azules (`#0033cc`).

---

## 🛠️ Tecnologías Empleadas
*   `Python 3.10+` (Backend Principal)
*   `FastAPI` & `Uvicorn` (Servidor web asíncrono ASGI)
*   `FPDF2` (Generación de reportería y estructura milimétrica en PDF y HTML)
*   `SQLAlchemy` & `PyODBC` / `PyMSSQL` (ORM y capas de conexión al SQL Server local).
*   `Jinja2` (Procesador de plantillas interfaces y vistas dinámicas `cliente.html`).

---

## ⚙️ Pasos de Instalación Rápida y Limpia

### 1. Preparación del Entorno (Librerías Code Base)
Abre una terminal normal (PowerShell o CMD) de Windows y posiciónate en el directorio raíz del proyecto (`interfaz_cartera`):
```bash
# 1. Crear el entorno virtual protegido para aislar descargas del computador base.
python -m venv venv

# 2. Activar el entorno virtual 
# Ejecuta en Windows CMD o PowerShell:
venv\Scripts\activate

# 3. Instalación de librerías y dependencias
pip install -r requirements.txt 
```

### 2. Configuración Secreta del Proyecto (`.env`)
El núcleo del proyecto utiliza variables de entorno invisibles y que **NO** están adjuntadas por seguridad. Para configurar conexiones y envíos crea un nuevo documento temporal pero renómbralo estrictamente la extensión y cuerpo a la palabra `.env` y rellénala con lo siguiente:
```env
# Credenciales Servidor Base de Datos SQL SERVER Interfaz Local
DB_USER=desarrollojosea
DB_PASS=Djosea01*
DB_HOST=192.168.1.14
DB_PORT=1433
DB_NAME=cartera_db
DB_ENCRYPT=yes
DB_TRUSTSERVERCERT=yes

# Credenciales cuenta SMTP Corporativa de Envío de Mensajes (Outlook)
SMTP_USER=credito_cartera@josegera.com
SMTP_PASS=Colombia2024.08
SMTP_HOST=smtp-mail.outlook.com
SMTP_PORT=587
```

### 3. Ejecución del Sistema Local (Puesta en marcha)
Existen un par de modalidades muy buenas para arrancar tu app dependiendo de si eres desarrollador probando, o trabajador operando de forma final.
*   **Modo Operador Directo (Recomendado Terminal Windows):** Simplemente haz doble clic sobre el ícono prefabricado del proyecto: `start_server.bat` y mágicamente salpicara el navegador corriendo por completo en el puerto local predeterminado de FastAPI.
*   **Modo Desarrollo En VIVO:** Estando en la terminal y solo si el entorno `venv` virtual está encendido de antemano usa: `uvicorn app:app --reload`. (Esto habilita el detector de salvado automático, de manera de que cualquier letra de Python que remplaces recargara velozmente el documento y te ahorrará abrir la consola una docena de veces).

---

## 🗂️ Control de Estructuras Obligatorias
Para que el creador de cartas no reviente sus flujos con errores, asegúrate de no mover ni renombrar las identidades de estas carpetas:
- **`assets/img/logo.png`**: Representará su imagen logotipo original translúcido sin fondos en la cabecera.
- **`assets/img/firma_sebastian.jpeg`**: Debe incluir el trozo escaneado de la firma oficial de la persona remitente incluyendo el cargo, teléfono y nombre debajo.
- **`cartas_generadas/`**: Carpeta temporal invisible en la que los documentos listos para cobrar se estancan como copias caché o por enviar (Puedes limpiarla mes a mes).
>>>>>>> 88feaa58f621bd76e66a1e27a61d65605f470c46

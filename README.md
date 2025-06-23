---

# 🧠 Prueba Técnica - Desarrollo RPA con PIX RPA

## 🎯 Objetivo General

Automatizar el análisis diario de productos de una tienda online ficticia mediante un proceso RPA que integre:

* Consumo de APIs REST
* Almacenamiento estructurado en base de datos
* Generación de reportes en Excel
* Automatización web para envío de formularios
* Integración con OneDrive usando Microsoft Graph API

---

## 🧩 Descripción del Proceso

Este robot fue desarrollado utilizando la plantilla universal de **PIX RPA** y ejecuta las siguientes tareas:

1. **Consumo de API Pública**

   * Fuente: [Fake Store API](https://fakestoreapi.com/docs#tag/Products)
   * Endpoint: `https://fakestoreapi.com/products`
   * El robot realiza una solicitud HTTP `GET` para obtener una lista de productos.
   * Guarda el JSON de respuesta en `/Logs/Productos_YYYY-MM-DD.json`.
   * Sube el archivo a OneDrive vía Microsoft Graph API en la ruta `/RPA/Logs/`.

2. **Almacenamiento en Base de Datos**

   * Almacena los datos estructurados en una base de datos (ej: SQLite).
   * Tabla: `Productos`
   * Campos: `id`, `title`, `price`, `category`, `description`, `fecha_insercion`
   * Evita duplicados validando que el `id` no exista previamente.

3. **Generación de Reporte en Excel**

   * Crea el archivo `Reporte_YYYY-MM-DD.xlsx` con:

     * **Hoja 1 - Productos**: lista completa de productos.
     * **Hoja 2 - Resumen**:

       * Total de productos
       * Precio promedio general
       * Precio promedio por categoría
       * Cantidad de productos por categoría
   * Guarda el archivo localmente en `/Reportes/`
   * Sube el reporte a OneDrive en la ruta `/RPA/Reportes/`.

4. **Automatización Web: Envío de Formulario**

   * El robot accede a un formulario web (Google Forms, Jotform o Typeform).
   * Rellena automáticamente los siguientes campos:

     * Nombre del colaborador
     * Fecha de generación del reporte
     * Comentarios (opcional)
     * Adjunta el archivo Excel generado
   * Envía el formulario
   * Captura una evidencia de la confirmación y la guarda como `/Evidencias/formulario_confirmacion.png`

5. **Integración con OneDrive**

   * Utiliza **Microsoft Graph API** para autenticar y subir archivos.
   * Validación de rutas: si la carpeta no existe, se crea automáticamente.
   * Manejo de archivos existentes: sobrescribe o versiona según configuración.
   * Registro de estado: éxito o fallo de cada carga se guarda en el log.

---

## 📦 Estructura del Proyecto

```
📁 Logs/
   └── Productos_YYYY-MM-DD.json
📁 Reportes/
   └── Reporte_YYYY-MM-DD.xlsx
📁 Evidencias/
   └── formulario_confirmacion.png
📁 BaseDatos/
   └── productos.db
📁 Scripts/
   ├── api_handler.py
   ├── db_manager.py
   ├── report_generator.py
   ├── onedrive_uploader.py
   └── web_automation.py
📄 README.md
```

---

## ⚙️ Requisitos Técnicos

* Python 3.10+
* Librerías: `requests`, `pandas`, `openpyxl`, `sqlite3`, `selenium`, `msal` (para Microsoft Graph)
* Driver Web: ChromeDriver o similar
* Cuenta de Microsoft con permisos para OneDrive
* Credenciales registradas en Azure Portal (client\_id, secret, tenant, scopes, etc.)

---

## 📝 Notas Finales

* El proceso es totalmente automatizado.
* La autenticación hacia OneDrive se realiza sin interacción del usuario.
* Se manejan logs y respaldos para trazabilidad completa.
* Ideal para tareas programadas (ej. robots diarios vía orquestador).


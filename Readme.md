# Sheet2docs 📊⇢📄

## Versión en Español (primera sección)

Gracias por usar Sheet2docs  

Este programa permite crear documentos (.docx) masivamente usando una plantilla y datos de hojas de cálculo (.xlsx).  
El archivo comprimido incluye un caso de prueba con documento de texto y hoja de cálculo de ejemplo.  

### TUTORIAL DE USO ###  

1. **Crear plantilla**:  
   - Coloca un archivo .docx que comience con "template" en la carpeta de Sheet2docs.exe  
   (ej: `template_inventario.docx` o `template_contactos.docx`)  
   - Marca las secciones reemplazables con corchetes angulares:  
   (ej: `<<nombre_cliente>>` o `<<número_factura>>`)  

2. **Preparar hoja de cálculo**:  
   - Crea un archivo .xlsx en la misma carpeta  
   - **La primera fila** debe contener palabras clave  
   (ej: `nombre_cliente` o `número_factura`)  
   - Cada fila subsiguiente generará un documento  

3. **Ejecutar el programa**:  
   - Ejecuta Sheet2docs.exe después de preparar los archivos  
   - Los documentos se generarán automáticamente  

### NOMBRES PERSONALIZADOS ###  
Para personalizar nombres de archivo:  
- Incluye una palabra clave entre guiones en el nombre de la plantilla  
(ej: `template_reporte_--id_cliente--.docx`)  
- Agrega una columna coincidente en tu hoja de cálculo  

### SOPORTE DE IMÁGENES (Beta) ###  
Para insertar imágenes:  
1. Agrega rutas de imágenes en celdas de la hoja  
(ej: `imagenes/logo.png`)  
2. **Limitaciones importantes**:  
   - Las imágenes no se redimensionan automáticamente  
   - Imágenes muy grandes/pequeñas pueden afectar el formato  

---

## English Version (following section)

Thank you for using Sheet2docs  

This program bulk-generates .docx documents using templates and spreadsheet data (.xlsx).  
The compressed file includes a test case with sample template and spreadsheet.  

### USAGE TUTORIAL ###  

1. **Create template**:  
   - Place a .docx file starting with "template" in Sheet2docs.exe's folder  
   (e.g., `template_inventory.docx` or `template_contacts.docx`)  
   - Mark replaceable sections with angle brackets:  
   (e.g., `<<client_name>>` or `<<invoice_number>>`)  

2. **Prepare spreadsheet**:  
   - Create a .xlsx file in the same folder  
   - **The first row** must contain template keywords  
   (e.g., `client_name` or `invoice_number`)  
   - Each subsequent row generates one output document  

3. **Run program**:  
   - Execute Sheet2docs.exe after preparing both files  
   - Documents will auto-generate  

### CUSTOM FILENAMES ###  
To personalize output filenames:  
- Include bracketed keyword in template's filename  
(e.g., `template_report_--client_id--.docx`)  
- Add matching column in your spreadsheet  

### IMAGE SUPPORT (Beta) ###  
To embed images:  
1. Add image paths in spreadsheet cells  
(e.g., `assets/logo.png`)  
2. **Key limitations**:  
   - Images don't auto-resize  
   - Oversized/small images may disrupt formatting  
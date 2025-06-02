Gracias por usar Sheet2docs

Este programa te permite crear masivamente documentos (.docx) utilizando una plantilla y datos de una hoja de cálculo (.xlsx)
En el archivo comprimido debería existir un caso de prueba con un documento de texto y una hoja de calculo de ejemplo.

### TUTORIAL DE USO ###

1. En la misma carpeta que se encuentre Sheet2docs.exe crea un documento de texto (.docx), el nombre debe empezar con "template"
	(ejemplo, 'template_algo' o 'templateInventario')
	Dentro de este documento, las secciones que serán reemplazados deben tener una palabra clave encerrada de comillas angulares
	(ejemplo, <<nombre>> o <<direccion_legal>>)

2. En la misma carpeta que se encuentre Sheet2docs.exe crea una plantilla de cálculo (.xlsx)
	La primera columna corresponderán a las palabras clave de la plantilla
	(ejemplo, nombre o direccion_legal)
	Por cada fila aparte del que contenga las palabras clave se creará un documento de texto con la información entregada

3. Una vez el documento de texto con nombre "template (.docx) y la hoja de cálculo (.xlsx) estén listos, 
	haz correr Sheet2docs.exe y se crearán los documentos de texto.

### NOMBRES PERSONALIZADOS ###

Si quieres que el nombre de cada documento sea personalizado
puedes usar una palabra clave de la hoja de cálculo y dejarlo en nombre de la plantilla
(ejemplo, "template example sheet2docs --name--.docx")


### IMAGENES (beta) ###

Si quieres poner imagenes puedes colocar la ruta de la imagen en una celda de la hoja de cálculo
(ejemplo, 'img/imagen ejemplo.jpg')
Pero debes tener cuidado con el tamaño de las imagenes, porque no se ajustan al documento de texto si son muy grandes o muy pequeñas.
versión 1.3
 - Genero el listado ordenado de los documentos.
 - Genero el expediente.
 - Añado los datos estáticos de Origen, Acceso, Idioma, Autor.

versión 1.4
 - Genero autoajuste de columnas.
 - Genero ejecutable.
 - Genero branding.

 versión 1.5
 - Busco en los nombres _F y devuelvo el código de calidad si lo tiene.

 versión 1.6
 - Busco las fechas del archivo.
 - Identifico cuál es el archivo principal.
 - Identifico cuáles son archivos anexos al principal.
 - Pongo las fechas de los archivos principales en las dos columnas de fechas.
 - Para archivos anexos uso la misma del archivo principal justamente anterior.
 - Aún tengo el inconveniente del formato de las columnas de fecha.

 versión 1.7
 - Optimización del código. --- Salió mal.
 - Busco las equivalencias de tipología en la columna "Nombre del archivo".
 - Devuelvo la tipología documental.
 - Busco las equivalencias de nombres en la columna "Nombre del archivo".
 - Devuelvo el nombre del documento.
 ---- ÚLTIMA VERSIÓN ESTABLE.

 versión 1.8
 - Separar el código en varias partes y corregir lo del nombre del documento. --- Salió mal.

 versión 1.9
 - Arreglé el código. Falta arreglar alineación y tipología documental.
 - Añadí las otras hojas.

 versión 1.10
 - Logré reestructurar el código y optimizarlo.

 versión 2.0
 - Cambio la interacción del programa con el usuario. Ahora pregunto si desea continuar con el siguiente expediente.

 versión 2.1
- Vuelvo a la anterior arquitectura. Versión 1.9.
- Genero la página metadatos_tipos_documentales. Falta nombre del documento y tipología documental.
- Genero la página metadatos_expediente. Implementaré un OCR más adelante. Número de orden no se genera correctamente.
- Genero la página Listas, falta la información.
- Estoy autoajustando las columnas.

versión 2.2
- Migré el código a inglés. No está completo.

versión 2.3
VERSIÓN MÁS ESTABLE
- Sigo en español.
- Genero la página metadatos_expediente. Implementaré un OCR más adelante. Número de orden no se genera correctamente.
- Genero la página metadatos_tipos_documentales completamente. Siguen los errores de formato de celda en las fechas.
- Genero la página Listas pero con errores.
- Abro y cierro automáticamente el archivo FCO. 66 para su revisión.
- Todo lo demás se está generando correctamente.

versión 2.4
- Archivo generado con éxito en su totalidad.

versión 2.5
- Pregunto al usuario si desea marcar ese expediente para posterior revisión y guardo el nombre del expediente en un txt.
- Elimino los mensajes de confirmación al cerrar el PDF.
- Mejoro la escritura del código.
- Evito que se analice la carpeta que Autofile genera: carpeta Excel.
- Estoy copiando cada reporte de Excel a cada uno de los expedientes.
- Busco si ese expediente ya fue analizado y tiene un Excel. Si sí, le pregunto al usuario si lo quiere reemplazar.
- Le pregunto al usuario desde qué número quiere empezar el contador. Si no ingresa nada, por defecto se usará 1.
- Corrijo el nombre de la subserie cruzando el código con el archivo equivalencias_subserie.json.
- Mejoro redacció, nombres de funciones y agrego comentarios.
- Cambié los inputs para que sean más directos.
- Compruebo que la carpeta base tenga el nombre correcto.

 ---

## v2.6 (18 de marzo de 2024)

### New features

* 

### Bug fixes

* 

### Improvements

* 

 ----
Tengo este script en Python:


De la hoja metadatos_expediente, las columnas R, S, T, U, V, W, X, Y, Z, AA y AB no se está autoajustando el ancho de las celdas.

Los datos de las columnas:
- De la hoja metadatos_expediente: A, K, O, P, U, X.
- De la hoja metadatos_tipologia_documental: D y E.
Se están almacenando como texto y sale error al abrir el Excel. Necesito que esos datos sean almacenados como número o convertidos a formato número en Excel.

La hoja Listas no se está generando correctamente. Debe ponerse la información estática que está en un archivo en la misma carpeta llamado listas.csv.

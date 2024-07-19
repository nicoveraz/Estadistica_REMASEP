# Estadistica_REMASEP

Este repositorio contiene el código para el sitio para proceso de datos proveniente del reporte excel "Tiempos de Urgencia" de MK web de Clínicas ACHS Salud para la confección del informe REMASEP

Es una versión web, replicando con javascript el proceso de datos del cuaderno python utilizado en Clínica Puerto Montt ACHS Salud

## Seguridad de Datos

**Todo el procesamiento de datos se realiza a nivel local** en el navegador web por lo que ningún dato es transmitido en el proceso

## Metodología

¿Por qué un cuaderno python y luego un sitio web para esto?

>Porque no me gusta utilizar excel

Este sitio web es una derivación del cuaderno python original, utiliza funciones javascript para procesar los datos del reporte excel a nivel local y llenar sobre la plantilla en blanco, permite acceder desde cualquier computardor sin necesidad de configurar el ambiente Python

## Instrucciones de Uso

Seleccionar archivo excel, exportado desde MK web, del mes correspondiente y presionar el botón procesar. El excel sin formato con los datos se descargará automáticamente

## Pendiente

- [ ] Mantener estilo de archivo original
- [ ] Completar todas las secciones
- [ ] Interfaz para revisar casos dudosos manualmente 
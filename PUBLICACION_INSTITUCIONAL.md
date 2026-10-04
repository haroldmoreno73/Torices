# Publicacion institucional privada

Este paquete contiene una interfaz estatica del informe Saber 11 2026 y datos individuales de estudiantes. No debe publicarse como sitio publico.

## Archivos

- `informe_icfes_2026.html`: interfaz del informe.
- `informe_icfes_2026_datos.json`: metricas y resultados individuales, sin identificadores.
- `informe_icfes_2026_ponderado_y_posicion.pdf`: informe descargable.
- `escudo_Torices.jpg`: escudo institucional de la interfaz.

## Requisitos del alojamiento

1. Alojar los cuatro archivos bajo HTTPS en un sitio o carpeta institucional dedicada.
2. Exigir autenticacion institucional antes de servir cualquiera de los cuatro archivos, incluido el PDF y el JSON.
3. Permitir solo solicitudes de lectura (`GET` y `HEAD`) a esos archivos; no habilitar listados de directorio, cargas ni ejecucion de scripts del lado del servidor.
4. Deshabilitar la indexacion por buscadores y el listado de archivos.
5. Crear un enlace corto solo despues de recibir la URL institucional protegida; usar un redireccionador administrado por el colegio.

La interfaz es estatica: no requiere acceso a otros servicios ni a la maquina de trabajo. No se debe exponer esta maquina con port-forwarding o tuneles publicos. La autenticacion se configura en el hosting institucional y debe proteger tambien los archivos de datos.
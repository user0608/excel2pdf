# Plan de mejora de excel2pdf

## Objetivo

Hacer más confiable la conversión de Excel a PDF con LibreOffice y Microsoft Excel, conservando la compatibilidad de `ConvertExcelToPdf` y la prioridad actual de LibreOffice.

## Avance

- **Etapa 1 implementada:** llamadas COM sin variantes `Must`, errores de pánico explícitos, cierre del libro y de Excel en las rutas de error, liberación de los `VARIANT` propietarios y tratamiento correcto de `S_FALSE` al inicializar COM.
- **Verificación de etapa 1:** 12 casos unitarios del ciclo de vida ejecutados en Linux con `go test -race ./...`; `go vet` correcto en Linux y para Windows amd64; pruebas de Windows compiladas mediante compilación cruzada.
- **Pendiente de validar en Windows con Office:** comportamiento real de apertura, exportación, cierre y liberación de recursos. Las pruebas del resultado de inicialización COM están añadidas, pero no se han ejecutado en Windows.
- **Etapa 2 implementada:** rutas de entrada absolutas y archivos regulares accesibles; directorios temporales exclusivos para ambos motores; eliminación de marcas de tiempo y renombrados compartidos; comprobación de que la salida sea un archivo regular no vacío y no un enlace simbólico; limpieza de temporales ante fallos.
- **Verificación de etapa 2:** pruebas de entradas inválidas, rutas relativas, espacios y Unicode, salida ausente o vacía, salida de tipo directorio o enlace, error con PDF parcial, nombres repetidos y liberación del mutex. `go test -race ./...` y `go vet` correctos en Linux; compilación de pruebas y `go vet` correctos para Windows amd64. No se han ejecutado conversiones reales con Office.
- **Contrato de limpieza documentado:** el consumidor elimina el PDF y su directorio exclusivo después de una conversión exitosa; el paquete limpia las conversiones fallidas.
- **Etapa 3 implementada:** perfil temporal exclusivo dentro del directorio de cada conversión, URI `file:` con escape de caracteres y tratamiento de unidades y rutas UNC de Windows, limpieza del perfil y captura de stdout/stderr sin logs duplicados ni escritura directa a la consola del consumidor. Los diagnósticos conservados se limitan a 64 KiB y se incluyen en los errores, también cuando el proceso termina con código cero sin generar PDF.
- **Verificación de etapa 3:** pruebas con procesos auxiliares para éxito, error de salida, ausencia de PDF, exceso de diagnósticos y fallo al iniciar; pruebas de limpieza, escape de URI y límite de salida. `go test -race ./...` y `go vet` correctos en Linux; pruebas compiladas y `go vet` correctos para Windows amd64.
- **Pendiente de integración de etapa 3:** LibreOffice no está instalado en este entorno. Queda por validar la conversión real con otra instancia abierta y desde procesos distintos; los procesos auxiliares solo verifican el contrato de ejecución. Las pruebas específicas de URI con unidad y UNC de Windows están compiladas, no ejecutadas.
- **Etapa 4 implementada con límites explícitos:** nueva API `ConvertExcelToPdfContext`, función original compatible sin timeout predeterminado, propagación del contexto, errores inspeccionables con `errors.Is`, limpieza de resultados cancelados y conservación del rechazo de conversiones simultáneas. LibreOffice usa `exec.CommandContext`, grupo propio en Linux y un límite de dos segundos para esperar sus pipes de salida. Excel comprueba cancelación entre operaciones COM sin omitir la limpieza.
- **Límites de etapa 4:** en Linux solo se alcanza a los descendientes que permanezcan en el grupo. En Windows se cancela el proceso lanzado, no se garantiza terminar sus descendientes; el cierre estricto del árbol requiere un Job Object o un trabajador aislado. No se ha añadido un trabajador para interrumpir COM. Estas garantías y posibles fallos de limpieza por handles abiertos están documentados; no se promete un timeout estricto para Excel.
- **Verificación de etapa 4:** pruebas de contextos nulos, cancelados y vencidos, concurrencia, propagación, limpieza de PDF cancelado, preservación de errores y cancelación entre operaciones de Excel con dobles de prueba. Pruebas de cancelación y vencimiento con procesos auxiliares bloqueados, y comprobación de grupo propio en Linux. `go test -race -count=3 ./...` y `go vet` correctos en Linux; compilación de pruebas y `go vet` correctos para Windows amd64. Pendiente ejecutar las pruebas en Windows y validar con Office real.
- **Siguiente etapa:** detección de motores y plataformas.

## Etapa 1: Corregir errores y liberar recursos de Excel

- Sustituir `MustGetProperty` y `MustCallMethod` por llamadas que devuelvan errores y comprobar sus resultados.
- Evitar que un pánico recuperado termine devolviendo un error nulo.
- Registrar `Excel.Quit` inmediatamente después de obtener la interfaz de la aplicación.
- Registrar el cierre del libro inmediatamente después de abrirlo.
- Liberar las referencias COM y los valores `VARIANT` según su propiedad, evitando liberaciones dobles.
- Revisar el tratamiento de los resultados de inicialización COM y mantener las operaciones en el mismo hilo del sistema operativo.
- Preservar el error principal y comunicar los errores relevantes de limpieza.

**Criterios de aceptación:** un fallo al abrir o exportar devuelve un error explícito; las rutas de error cierran la instancia de Excel creada por el paquete y liberan sus recursos.

## Etapa 2: Aislar archivos temporales y validar resultados

- Crear un directorio temporal único por conversión para ambos motores.
- Dejar de construir nombres únicos mediante marcas de tiempo en segundos.
- Normalizar la ruta de entrada a absoluta y comprobar que corresponde a un archivo regular accesible.
- Comprobar que el PDF esperado existe, es regular y no está vacío antes de devolver éxito.
- Limpiar los archivos auxiliares cuando la conversión falle.
- Definir el contrato de limpieza del PDF devuelto: quién lo elimina y cómo se elimina su directorio temporal.
- No eliminar ni sobrescribir archivos preexistentes ajenos a la conversión.

**Criterios de aceptación:** documentos con el mismo nombre no comparten resultados; una ejecución que no genere un PDF nunca devuelve uno de una conversión anterior.

## Etapa 3: Robustecer la ejecución de LibreOffice

- Crear un perfil temporal exclusivo por conversión.
- Pasarlo mediante `-env:UserInstallation=...`, usando una URI de archivo válida para la plataforma.
- Capturar la salida del proceso e incluir información útil en los errores sin escribir directamente en la salida del consumidor.
- Limpiar el perfil y los archivos auxiliares una vez finalizada la ejecución.
- Probar conversiones con otra instancia de LibreOffice abierta y desde procesos distintos.

**Criterios de aceptación:** el paquete no depende del perfil de LibreOffice del usuario y detecta la ausencia del PDF incluso si el proceso termina con código cero.

## Etapa 4: Añadir cancelación y límites de ejecución

- Añadir una función que acepte `context.Context`, manteniendo la API existente como envoltorio compatible.
- Usar `exec.CommandContext` para LibreOffice y revisar el cierre de procesos descendientes y la espera de sus salidas en cada plataforma.
- Diferenciar errores de conversión, cancelación y vencimiento del plazo mediante errores inspeccionables.
- Documentar que el mutex actual rechaza conversiones simultáneas dentro del mismo proceso; mantener ese comportamiento salvo decisión explícita de cambiarlo.
- Definir por separado el alcance de cancelación de Excel: un contexto no interrumpe automáticamente una llamada COM bloqueada.
- Si se requiere un timeout estricto para Excel, diseñar un proceso trabajador aislado y terminar únicamente los procesos propiedad de esa conversión, nunca otras instancias del usuario.

**Criterios de aceptación:** LibreOffice respeta la cancelación sin dejar bloqueado el paquete; las limitaciones de Excel quedan documentadas y no se promete un timeout que COM no pueda garantizar.

## Etapa 5: Reparar la detección de motores y plataformas

- Sustituir las llamadas a `which` por `exec.LookPath`.
- Intentar los candidatos alternativos cuando no se encuentre el primero, incluyendo el candidato actualmente previsto `libreoffice24.8`.
- Validar `LIBREOFFICE_PATH` y definir el comportamiento cuando esté vacía o apunte a una ubicación inválida.
- Cerrar todas las claves del registro de Windows después de usarlas.
- Revisar las vistas de registro de 32 y 64 bits y las instalaciones por usuario.
- Comprobar que la detección de LibreOffice devuelve un ejecutable utilizable, no solamente un directorio de instalación.
- Diferenciar un motor no instalado de un error al consultar su instalación.
- Mantener y documentar la selección actual: LibreOffice primero y Excel cuando LibreOffice no está instalado. No introducir un fallback tras errores de conversión sin definir antes esa política.
- Decidir si se soportará macOS. Si se incluye, renombrar los archivos `_linux.go`, ajustar las etiquetas y añadir detección específica; en caso contrario, retirar las etiquetas `darwin` engañosas.

**Criterios de aceptación:** la ausencia del primer candidato no impide buscar alternativas; Windows detecta los motores en los escenarios soportados; la matriz de plataformas documentada compila.

## Etapa 6: Reforzar el procesamiento de documentos no confiables

- Configurar explícitamente la seguridad de macros antes de abrir documentos con Excel.
- Desactivar la actualización automática de enlaces externos y evitar diálogos interactivos en la medida permitida por el motor.
- Documentar las limitaciones de la automatización de Office en servicios sin sesión interactiva.
- Recomendar un usuario sin privilegios y un entorno aislado para documentos de terceros.
- Evitar que los diagnósticos expongan contenido del documento o rutas sensibles innecesariamente.

**Criterios de aceptación:** la política de apertura de documentos es explícita y las restricciones de seguridad y despliegue quedan documentadas; no se presenta la librería como un sandbox.

## Etapa 7: Incorporar pruebas, CI y documentación

- Añadir pruebas unitarias para validación de entradas, detección de ejecutables, selección de motor y gestión de errores.
- Probar la ejecución de LibreOffice mediante procesos auxiliares controlados: éxito, código de salida fallido, ausencia de PDF y cancelación.
- Cubrir nombres con espacios, caracteres Unicode, entradas inexistentes, nombres repetidos y `LIBREOFFICE_PATH` vacía o inválida.
- Añadir pruebas de integración opcionales con LibreOffice real en Linux y con Excel y LibreOffice en Windows.
- Verificar en Windows los fallos de apertura y exportación y la liberación de recursos COM.
- Configurar CI con `go test ./...`, `go vet ./...` y compilación cruzada para las plataformas soportadas.
- Ampliar el README con requisitos, configuración, prioridad de motores, cancelación, concurrencia, limpieza de temporales y ejemplos completos.
- Retirar logs duplicados y corregir nombres y mensajes internos sin romper símbolos públicos existentes.

**Criterios de aceptación:** los errores principales tienen pruebas de regresión; las integraciones que requieren Office se ejecutan únicamente en entornos preparados; la API existente continúa compilando y los ejemplos reflejan su contrato real.

## Orden y alcance

Ejecutar las etapas en orden, con cambios pequeños y verificables. Añadir pruebas de regresión junto a cada corrección; la etapa 7 consolida la cobertura y la automatización, no pospone todas las pruebas hasta el final.

Las decisiones sobre soporte de macOS, timeout estricto de Excel y cambios en la política de concurrencia deben resolverse explícitamente antes de ampliar el alcance del paquete.

## Estado inicial de la revisión

- `go test ./...`: correcto, sin archivos de pruebas.
- `go vet ./...`: sin avisos.
- Compilación cruzada para Windows amd64: correcta.
- Compilación cruzada para macOS arm64: falla porque las funciones necesarias están en archivos con sufijo `_linux.go`.
- No se han validado conversiones reales con LibreOffice ni Microsoft Excel durante esta revisión.

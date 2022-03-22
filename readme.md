https://srhuevo.github.io/code-to-post/

##USO
Configurar el correo smtp desde el que se van a enviar los correos.

Configurar el cuerpo del mensaje que se va a enviar. En formato HTML. Cada vez que se modifica se actualizará la previsualización.

Configurar archivos adjuntos en formato HTML, se transformarán a PDF, no funcionarán muchas características de las nuevas de CSS. Si se pone alguna ruta de imagen que no exista fallara y no se previsualizará el archivo. Cada vez que se modifica el template se actualizará la previsualización. Añadir nombre y orientación.

Cargar fichero xslx, cada fila del mismo será un correo que se envíe, la primera fila se ignorará y se utilizará para hacer de clave en los comodines. Saldrá una barra de progreso informando de cuantos emails se van a enviar.

Cuando este todo configurado, pulsar el boton de enviar para que empiece el proceso. La barra de progreso se ira completando en cada envío.

##COMODINES

Los comodines tienen de clave el nombre que la primera fila del excel que se aporte, además existe el comodin "date" para poner una fecha en formato "dd/mm/YYYY"

Tanto en el cuerpo del mensaje como en todos los archivos adjuntos se podrán usar comodines.

En los campos To y Asunto se pueden utilizar comodines.

##RECOMENDACIONES

Si se van a colocar imagenes dentro de imagenes en los PDF se recomienda que estas tengan las mismas dimensiones que los huecos que se dejan, de ser diferentes complicará muchisimo que encajen bien.

En el caso de querer poner una imagen con comodín se recomienda subir una imagen de ejemplo con el comodin dentro de la ruta, ejemplo:
 - students/to-process/certificates inmersive 2022T1/{{Current course}}.png

Antes de enviar los correos a la gente se recomienda poner en el campo TO un email propio para hacer una primera prueba.

No se recomienda usar gmail porque puede bloquearse al enviar demasiados emails.

Actualizar o parar la pagina cuando se están enviando los correos emite una alerta que paraliza el proceso, si no se termina cerrando el proceso se reanuda y si se cierra se finaliza el proceso.
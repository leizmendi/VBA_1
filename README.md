# VBA_1 Access #
## Parámetros, parámetros y más parámetros ##

Se trata de una base de datos Access con una propuesta para el manejo de los parámetros de una aplicación.

Los archivos .def son los objetos (Tabla, formularios y módulo) en texto plano, importables desde Oasis o desde código con la función LoadFromText.

El archivo .rar contiene una copia del .accdb que se puede descomprimir y ejecutar directamente

---

# Lección

Cuando volvemos a abrir un formulario que habíamos utilizado anteriormente, generalmente es práctico recuperar ciertos valores que se teclearon o seleccionaron en lugar de resetear el formulario obligando a repetir dicha selección. Eso lo podemos conseguir guardando esos valores a mantener y al abrir el formulario asignarlos a sus controles correspondientes.

A estos valores les denominé Parámetros, como a los valores que utilizamos para configurar una aplicación.

En esta propuesta se utiliza la tabla cfgParam para guardar los Parámetros, en dicha tabla encontramos el campo NP que es el Nombre del parámetro y es Clave Principal (no puede haber 2 iguales), el campo TipoDato  con los posibles valores de 1=Boolean, 4=Long, 5=Currency, 8=Date, 10=Text(255) y 12=Memo. El campo VP guarda el valor del parámetro cuando TipoTexto = 10 y el resto de campos VPbool, VPlng, VPcur, VPmemo y VPfecha guardan los valores que les corresponden. Este pequeño lío tal vez no sea necesario y se podrían haber guardado todos los valores en VP (salvo los Memo muy grandes) pero salió así y así se quedó y tampoco ha ocasionado mayor problema.

En el módulo basParam de este ejemplo encontramos las funciones SetParam(sNP) y GetParam(sNP) que se encargan respectivamente de grabar y recuperar un parámetro accediendo para ello a la tabla cfgParam. En estas funciones encontramos un segundo argumento opcional bUser del que quiero explicar su utilidad. En una aplicación utilizada por más de un usuari@ los valores a mantener para la próxima ocasión es probable que sea interesante que cada usuari@guarde los suyos y es para esto que si llamamos a una de estas funciones con el valor bUser=True, la función añadirá a NP un sufijo de '_USERNAME' de forma habrá valores diferenciados para cada usuario.

La magia reside en los procedimientos GrabarParam(frm, sPrefijo) y CargarParam(frm, sPrefijo) que recorren todos los controles del formulario y revisan la propiedad Tag, si el tag contiene la cadena "param" se trata de un parámetro si además contiene la cadena paramUS se trata de un parámetro de usuario. Los 3 dígitos de siguen a param o a paramUS indican el tipo de dato: 001 para boolean, 004 para long, 005 currency, 008 date, 010 texto y 012 memo, si no se indican el valor por defecto es 010=Texto.

En resumidas cuentas sólo tenemos que etiquetar (en la propiedad tag) los controles que queremos recordar y llamar a CargarParam (Me, Me.Name & "_") en el Form_Load y a GrabarParam(Me, Me.Name & "_") en el Form_Unload. Lo del prefijo Me.Name & "_" sirve para agrupar todos los parámetros de un formulario precediéndolos del nombre del formulario.

Puedes revisar la tabla cfgParam para observar como se guardan los parámetros

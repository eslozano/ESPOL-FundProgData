# ESPOL - Fundamentos de Programación


## Validation data script
Este script servirá para validar los datos recolectados en Fundamentos de Programación. El script deberá variar para cada semestre según las ponderaciones de los temas de los examenes, proyectos y sustentaciones. 

**Datos a validar(2017-II):**

* Nombre
* Matricula
* Genero
* Paralelo
* cod_carrera
* veces_tomadas
* 1er_proyecto
* 1er_sustent
* 1er_calif\_final
* 1er_exam\_tema1
* 1er_exam\_tema2
* 1er_exam\_tema3
* 1er_exam\_tema4
* 2do_proyecto
* 2do_sustent
* 2do_calif\_final
* 2do_exam\_tema1
* 2do_exam\_tema2
* 2do_exam\_tema3
* calif_final\_practica
* 3er_proyecto
* 3er_calif\_final
* 3er_exam\_tema1
* 3er_exam\_tema2
* 3er_exam\_tema3

**Validaciones implementadas**

| Tipo de error    | Color    | Script fixs it | Comentarios |
| ---------------  |----------| :------:| ----------------|
| Vacio            | Naranja  |         | Se validan los datos básicos |
| Opción no valida | Azul     |         | Debe estar entre las opciones válidas |
| Fuera de rango   | Amarillo |         | Es mayor o menor al valor posible |
| Tipo no numérico | Verde    |         | No es entero ni float |
| No redondeado    | Rosado   |  SI     | Esto aplica para las calificaciones finales |

## Dependencies

Install the next python packages:

* [Openpyxl](https://openpyxl.readthedocs.io/en/stable/)

## Usage

* Complete los datos de su paralelo de fundamentos en el archivo **PARNN_2017II.xlsx**
* Corra el script pero modifique el nombre del archivo en la línea 8. Al finalizar se imprimirá un mensaje con la cantidad de estudiantes en los que se encontraron errores.
* Abra su archivo y verifique si hay celdas con los colores descritos en la tabla descrita anteriormente.
* Corrija los errores y vuelva a correr el script.


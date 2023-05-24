/* 
----------------------------------------------Gestion de Auxiliares-----------------------------------------------------------
*/
 
function LimpiarAuxiliar() {
  //guardamos en la variable la hoja de calculo activa 
  var hojaActiva = SpreadsheetApp.getActiveSpreadsheet();

  //Guardamos en la variable el nombre de la hoja
  var hojaFormulario = hojaActiva.getSheetByName("Auxiliares");

  //celdas ue vamos a eliminar
  var celdasEliminar = ["G4","G6","G8","G10", "G12", "G14"];

  //por medio del for borramos el texto de las casillas
  for (var i=0; i<celdasEliminar.length; i++)
  {
    hojaFormulario.getRange(celdasEliminar[i]).clearContent();
  }
}

function GuardarAuxiliar(){
  //guardamos en la variable la hoja de calculo activa 
  var hojaActiva = SpreadsheetApp.getActiveSpreadsheet();

  //Guardamos en las variable el nombre de la hojas
  var formulario = hojaActiva.getSheetByName("Auxiliares");
  var datos = hojaActiva.getSheetByName("BDauxiliares");

  //guardamos la informacion ingresada por el usuario
  var valores = [[formulario.getRange("G4").getValue() ,
                  formulario.getRange("G6").getValue() ,
                  formulario.getRange("G8").getValue() ,
                  formulario.getRange("G10").getValue(),
                  formulario.getRange("G12").getValue(),
                  formulario.getRange("G14").getValue(),]];

  // añadimos los valores creados a la hoja donde estan ubicados los datos
  datos.getRange(datos.getLastRow()+1,1,1,6).setValues(valores);

  LimpiarAuxiliar();

}

 function ActualizarAuxiliar() {
  //guardamos en la variable la hoja de calculo activa 
  var hojaActiva = SpreadsheetApp.getActiveSpreadsheet();

  //Guardamos en las variable el nombre de la hojas
  var formulario = hojaActiva.getSheetByName("Auxiliares");
  var datos = hojaActiva.getSheetByName("BDauxiliares");

  //celdad de la cual filtraremos
  var valor = formulario.getRange("I8").getValue();
 
  //hoja donde estan los datos
  var valores = hojaActiva.getSheetByName("BDauxiliares").getDataRange().getValues(); 

  for (var i = 0; i < valores.length; i++) {
    var fila = valores[i];
    if(fila[NUM_COLUMNA_BUSQUEDA] == valor) {
      var INT_R = i+1
      
      var valores1 = [[formulario.getRange("G4").getValue(),
                      formulario.getRange("G6").getValue() ,
                      formulario.getRange("G8").getValue() ,
                      formulario.getRange("G10").getValue(),
                      formulario.getRange("G12").getValue(),
                      formulario.getRange("G14").getValue(),]];
      
      datos.getRange(INT_R, 1, 1, 6).setValues(valores1);
      SpreadsheetApp.getUi().alert('Datos actualizados');

      LimpiarAuxiliar(); // Ejecución de función para limpieza de celdas
    }
  }
}


// Buscar
var NUM_COLUMNA_BUSQUEDA = 0;
function BuscarAux() {
  var hojaActiva = SpreadsheetApp.getActiveSpreadsheet();

  // Nombre de hoja del formulario
  var formulario = hojaActiva.getSheetByName("Auxiliares"); 
 
  var valor = formulario.getRange("I8").getValue();
  var valores = hojaActiva.getSheetByName("BDauxiliares").getDataRange().getValues(); // Nombre de hoja donde se almacenan datos
  for (var i = 0; i < valores.length; i++) {
     var fila = valores[i];
    if (fila[NUM_COLUMNA_BUSQUEDA] == valor) {
      
      formulario.getRange("G4").setValue(fila[0]);
      formulario.getRange("G6").setValue(fila[1]);
      formulario.getRange("G8").setValue(fila[2]);
      formulario.getRange("G10").setValue(fila[3]);
      formulario.getRange("G12").setValue(fila[4]);
      formulario.getRange("G14").setValue(fila[5]);
    }
  }
}


function EliminarAux() {
  
  var hojaActiva = SpreadsheetApp.getActiveSpreadsheet();
  // Nombre de hoja del formulario
  var formulario = hojaActiva.getSheetByName("Auxiliares"); 

  // Nombre de hoja donde se almacenan datos
  var datos = hojaActiva.getSheetByName("BDauxiliares"); 
  
  var interface = SpreadsheetApp.getUi();
  //boton
  var respuesta = interface.alert('¿Estas seguro de borrar?',interface.ButtonSet.YES_NO);
  
  // Proceso si el usuario responde
  if (respuesta == interface.Button.YES) {
    
    var valor = formulario.getRange("I8").getValue();

    // Nombre de hoja donde se almacenan datos
    var valores = hojaActiva.getSheetByName("BDauxiliares").getDataRange().getValues(); 

    for (var i = 0; i< valores.length; i++) {
      var fila = valores[i];
      if (fila[NUM_COLUMNA_BUSQUEDA] == valor) {
        var INT_R = i+1
        
        datos.deleteRow(INT_R);
        LimpiarAuxiliar(); 
      }
    }
  }
}


/* 
----------------------------------------------Gestion de Equipos ------------------------------------------------------------
*/

//Equipos
function LimpiarEquipos() {
  //guardamos en la variable la hoja de calculo activa 
  var hojaActiva = SpreadsheetApp.getActiveSpreadsheet();

  //Guardamos en la variable el nombre de la hoja
  var hojaFormulario = hojaActiva.getSheetByName("Equipos");

  //celdas ue vamos a eliminar
  var celdasEliminar = ["G4","G6","G8"];

  //por medio del for borramos el texto de las casillas
  for (var i=0; i<celdasEliminar.length; i++)
  {
    hojaFormulario.getRange(celdasEliminar[i]).clearContent();
  }
}

function GuardarEquipos(){
  //guardamos en la variable la hoja de calculo activa 
  var hojaActiva = SpreadsheetApp.getActiveSpreadsheet();

  //Guardamos en las variable el nombre de la hojas
  var formulario = hojaActiva.getSheetByName("Equipos");
  var datos = hojaActiva.getSheetByName("BDequipos");

  //guardamos la informacion ingresada por el usuario
  var valores = [[formulario.getRange("G4").getValue() ,
                  formulario.getRange("G6").getValue() ,
                  formulario.getRange("G8").getValue(),]];

  // añadimos los valores creados a la hoja donde estan ubicados los datos
  datos.getRange(datos.getLastRow()+1,1,1,3).setValues(valores);

  LimpiarEquipos();

}

 function ActualizarEquipos() {
  //guardamos en la variable la hoja de calculo activa 
  var hojaActiva = SpreadsheetApp.getActiveSpreadsheet();

  //Guardamos en las variable el nombre de la hojas
  var formulario = hojaActiva.getSheetByName("Equipos");
  var datos = hojaActiva.getSheetByName("BDequipos");

  //celdad de la cual filtraremos
  var valor = formulario.getRange("I6").getValue();
 
  //hoja donde estan los datos
  var valores = hojaActiva.getSheetByName("BDequipos").getDataRange().getValues(); 

  for (var i = 0; i < valores.length; i++) {
    var fila = valores[i];
    if(fila[NUM_COLUMNA_BUSQUEDA] == valor) {
      var INT_R = i+1
      
      var valores1 = [[formulario.getRange("G4").getValue(),
                      formulario.getRange("G6").getValue() ,
                      formulario.getRange("G8").getValue(),]];
      
      datos.getRange(INT_R, 1, 1, 3).setValues(valores1);
      SpreadsheetApp.getUi().alert('Datos actualizados');

      LimpiarEquipos(); // Ejecución de función para limpieza de celdas
    }
  }
}


// Buscar
var NUM_COLUMNA_BUSQUEDA = 0;
function BuscarEquipos() {
  var hojaActiva = SpreadsheetApp.getActiveSpreadsheet();

  // Nombre de hoja del formulario
  var formulario = hojaActiva.getSheetByName("Equipos"); 
 
  var valor = formulario.getRange("I6").getValue();
  var valores = hojaActiva.getSheetByName("BDequipos").getDataRange().getValues(); // Nombre de hoja donde se almacenan datos
  for (var i = 0; i < valores.length; i++) {
     var fila = valores[i];
    if (fila[NUM_COLUMNA_BUSQUEDA] == valor) {
      
      formulario.getRange("G4").setValue(fila[0]);
      formulario.getRange("G6").setValue(fila[1]);
      formulario.getRange("G8").setValue(fila[2]);
    }
  }
}


function EliminarEquipos() {
  
  var hojaActiva = SpreadsheetApp.getActiveSpreadsheet();
  // Nombre de hoja del formulario
  var formulario = hojaActiva.getSheetByName("Equipos"); 

  // Nombre de hoja donde se almacenan datos
  var datos = hojaActiva.getSheetByName("BDequipos"); 
  
  var interface = SpreadsheetApp.getUi();
  //boton
  var respuesta = interface.alert('¿Estas seguro de borrar?',interface.ButtonSet.YES_NO);
  
  // Proceso si el usuario responde
  if (respuesta == interface.Button.YES) {
    
    var valor = formulario.getRange("I6").getValue();

    // Nombre de hoja donde se almacenan datos
    var valores = hojaActiva.getSheetByName("BDequipos").getDataRange().getValues(); 
    
    for (var i = 0; i< valores.length; i++) {
      var fila = valores[i];
      if (fila[NUM_COLUMNA_BUSQUEDA] == valor) {
        var INT_R = i+1
        
        datos.deleteRow(INT_R);
        LimpiarEquipos(); 
      }
    }
  }
}


/* 
--------------------------------------------- Gestion de Profesores -----------------------------------------------------------
*/

//Profesores
function LimpiarProfesores() {
  //guardamos en la variable la hoja de calculo activa 
  var hojaActiva = SpreadsheetApp.getActiveSpreadsheet();

  //Guardamos en la variable el nombre de la hoja
  var hojaFormulario = hojaActiva.getSheetByName("Profesores");

  //celdas ue vamos a eliminar
  var celdasEliminar = ["G4","G6","G8"];

  //por medio del for borramos el texto de las casillas
  for (var i=0; i<celdasEliminar.length; i++)
  {
    hojaFormulario.getRange(celdasEliminar[i]).clearContent();
  }
}

function GuardarProfesor(){
  //guardamos en la variable la hoja de calculo activa 
  var hojaActiva = SpreadsheetApp.getActiveSpreadsheet();

  //Guardamos en las variable el nombre de la hojas
  var formulario = hojaActiva.getSheetByName("Profesores");
  var datos = hojaActiva.getSheetByName("BDprofesores");

  //guardamos la informacion ingresada por el usuario
  var valores = [[formulario.getRange("G4").getValue() ,
                  formulario.getRange("G6").getValue() ,
                  formulario.getRange("G8").getValue(),]];

  // añadimos los valores creados a la hoja donde estan ubicados los datos
  datos.getRange(datos.getLastRow()+1,1,1,3).setValues(valores);

  LimpiarProfesores();

}

 function ActualizarProfesor() {
  //guardamos en la variable la hoja de calculo activa 
  var hojaActiva = SpreadsheetApp.getActiveSpreadsheet();

  //Guardamos en las variable el nombre de la hojas
  var formulario = hojaActiva.getSheetByName("Profesores");
  var datos = hojaActiva.getSheetByName("BDprofesores");

  //celdad de la cual filtraremos
  var valor = formulario.getRange("I6").getValue();
 
  //hoja donde estan los datos
  var valores = hojaActiva.getSheetByName("BDprofesores").getDataRange().getValues(); 

  for (var i = 0; i < valores.length; i++) {
    var fila = valores[i];
    if(fila[NUM_COLUMNA_BUSQUEDA] == valor) {
      var INT_R = i+1
      
      var valores1 = [[formulario.getRange("G4").getValue(),
                      formulario.getRange("G6").getValue() ,
                      formulario.getRange("G8").getValue(),]];
      
      datos.getRange(INT_R, 1, 1, 3).setValues(valores1);
      SpreadsheetApp.getUi().alert('Datos actualizados');

      LimpiarProfesores(); // Ejecución de función para limpieza de celdas
    }
  }
}


// Buscar
var NUM_COLUMNA_BUSQUEDA = 0;
function BuscarProfesor() {
  var hojaActiva = SpreadsheetApp.getActiveSpreadsheet();

  // Nombre de hoja del formulario
  var formulario = hojaActiva.getSheetByName("Profesores"); 
 
  var valor = formulario.getRange("I6").getValue();
  var valores = hojaActiva.getSheetByName("BDprofesores").getDataRange().getValues(); // Nombre de hoja donde se almacenan datos
  for (var i = 0; i < valores.length; i++) {
     var fila = valores[i];
    if (fila[NUM_COLUMNA_BUSQUEDA] == valor) {
      
      formulario.getRange("G4").setValue(fila[0]);
      formulario.getRange("G6").setValue(fila[1]);
      formulario.getRange("G8").setValue(fila[2]);
    }
  }
}


function EliminarProfesor() {
  
  var hojaActiva = SpreadsheetApp.getActiveSpreadsheet();
  // Nombre de hoja del formulario
  var formulario = hojaActiva.getSheetByName("Profesores"); 

  // Nombre de hoja donde se almacenan datos
  var datos = hojaActiva.getSheetByName("BDprofesores"); 
  
  var interface = SpreadsheetApp.getUi();
  //boton
  var respuesta = interface.alert('¿Estas seguro de borrar?',interface.ButtonSet.YES_NO);
  
  // Proceso si el usuario responde
  if (respuesta == interface.Button.YES) {
    
    var valor = formulario.getRange("I6").getValue();

    // Nombre de hoja donde se almacenan datos
    var valores = hojaActiva.getSheetByName("BDprofesores").getDataRange().getValues(); 
    
    for (var i = 0; i< valores.length; i++) {
      var fila = valores[i];
      if (fila[NUM_COLUMNA_BUSQUEDA] == valor) {
        var INT_R = i+1
        
        datos.deleteRow(INT_R);
        LimpiarProfesores(); 
      }
    }
  }
}




/* 
--------------------------------------------- Gestion de Prestamos -----------------------------------------------------------
*/

//Prestamos
function LimpiarPrestamos() {
  //guardamos en la variable la hoja de calculo activa 
  var hojaActiva = SpreadsheetApp.getActiveSpreadsheet();

  //Guardamos en la variable el nombre de la hoja
  var hojaFormulario = hojaActiva.getSheetByName("Prestamos");

  //celdas ue vamos a eliminar
  var celdasEliminar = ["G5","G7", "I7"];

  //por medio del for borramos el texto de las casillas
  for (var i=0; i<celdasEliminar.length; i++)
  {
    hojaFormulario.getRange(celdasEliminar[i]).clearContent();
  }
}

function GuardarPrestamo(){
  //guardamos en la variable la hoja de calculo activa 
  var hojaActiva = SpreadsheetApp.getActiveSpreadsheet();

  //Guardamos en las variable el nombre de la hojas
  var formulario = hojaActiva.getSheetByName("Prestamos");
  var datos = hojaActiva.getSheetByName("BDprestamos");

  //guardamos la informacion ingresada por el usuario
  var valores = [[formulario.getRange("G5").getValue() ,
                  formulario.getRange("G7").getValue() ,]];

  // añadimos los valores creados a la hoja donde estan ubicados los datos
  datos.getRange(datos.getLastRow()+1,1,1,2).setValues(valores);

  LimpiarProfesores();

}


// Buscar
var NUM_COLUMNA_BUSQUEDA = 0;
function BuscarPrestamos() {
  var hojaActiva = SpreadsheetApp.getActiveSpreadsheet();

  // Nombre de hoja del formulario
  var formulario = hojaActiva.getSheetByName("Prestamos"); 
 
  var valor = formulario.getRange("I7").getValue();
  var valores = hojaActiva.getSheetByName("BDprestamos").getDataRange().getValues(); // Nombre de hoja donde se almacenan datos
  for (var i = 0; i < valores.length; i++) {
     var fila = valores[i];
    if (fila[NUM_COLUMNA_BUSQUEDA] == valor) {
      
      formulario.getRange("G5").setValue(fila[0]);
      formulario.getRange("G7").setValue(fila[1]);
    }
  }
}


function EliminarPrestamo() {
  
  var hojaActiva = SpreadsheetApp.getActiveSpreadsheet();
  // Nombre de hoja del formulario
  var formulario = hojaActiva.getSheetByName("Prestamos"); 

  // Nombre de hoja donde se almacenan datos
  var datos = hojaActiva.getSheetByName("BDprestamos"); 
  
  var interface = SpreadsheetApp.getUi();
  //boton
  var respuesta = interface.alert('¿Estas seguro de quitar el prestamo?',interface.ButtonSet.YES_NO);
  
  // Proceso si el usuario responde
  if (respuesta == interface.Button.YES) {
    
    var valor = formulario.getRange("I7").getValue();

    // Nombre de hoja donde se almacenan datos
    var valores = hojaActiva.getSheetByName("BDprestamos").getDataRange().getValues(); 
    
    for (var i = 0; i< valores.length; i++) {
      var fila = valores[i];
      if (fila[NUM_COLUMNA_BUSQUEDA] == valor) {
        var INT_R = i+1
        
        datos.deleteRow(INT_R);
        LimpiarPrestamos(); 
      }
    }
  }
}





/* 
-----------------------------------------Conteo de horas----------------------------------------------------------------------
*/


function LimpiarTurno() {
  //guardamos en la variable la hoja de calculo activa 
  var hojaActiva = SpreadsheetApp.getActiveSpreadsheet();

  //Guardamos en la variable el nombre de la hoja
  var hojaFormulario = hojaActiva.getSheetByName("ConteoHoras");

  //celdas ue vamos a eliminar
  var celdasEliminar = ["D9", "D13","D15"];

  //por medio del for borramos el texto de las casillas
  for (var i=0; i<celdasEliminar.length; i++)
  {
    hojaFormulario.getRange(celdasEliminar[i]).clearContent();
  }
}



 function ActualizarHoras() {
  //guardamos en la variable la hoja de calculo activa 
  var hojaActiva = SpreadsheetApp.getActiveSpreadsheet();

  //Guardamos en las variable el nombre de la hojas
  var formulario = hojaActiva.getSheetByName("ConteoHoras");
  var datos = hojaActiva.getSheetByName("BDhoras");

  //celdad de la cual filtraremos
  var valor = formulario.getRange("D4").getValue();

  //hoja donde estan los datos
  var valores = hojaActiva.getSheetByName("BDhoras").getDataRange().getValues(); 

  for (var i = 0; i < valores.length; i++) {
    var fila = valores[i];
    if(fila[NUM_COLUMNA_BUSQUEDA] == valor) {
      var INT_R = i+1

      var numeroFila = i;

      //se obtine las horas acumuladas
      formulario.getRange("D9").setValue("=BDhoras!B"+(numeroFila+1));


    //el "O4" es la suma de las horas acumuladas y las horas realizadas ese dia
      var valores1 = [[formulario.getRange("D4").getValue(),
                      formulario.getRange("L4").getValue(),]];

      datos.getRange(INT_R, 1, 1, 2).setValues(valores1);
      SpreadsheetApp.getUi().alert('Datos actualizados');

    }
  }


}


function ingresarTurno(){
LimpiarTurno();

//creamos una variable para la fecha y guardamos la fecha actual
var fecha = new Date();
Logger.log(fecha);

//obtenemos la hoja activa
var hojaActiva = SpreadsheetApp.getActiveSpreadsheet();
var formulario = hojaActiva.getSheetByName("ConteoHoras");

//ingresamos la fecha actual en la celda
formulario.getRange("D13").setValue(fecha);

}


function terminarTurno(){
var fecha = new Date();
Logger.log(fecha);

var hojaActiva = SpreadsheetApp.getActiveSpreadsheet();
var formulario = hojaActiva.getSheetByName("ConteoHoras");

formulario.getRange("D15").setValue(fecha);

//var fechaFin = formulario.getRange("D15").getValue();


ActualizarHoras();

}



function ConsultarHoras(){
  
//obtenemos la hoja activa
var hojaActiva = SpreadsheetApp.getActiveSpreadsheet();
var formulario = hojaActiva.getSheetByName("ConteoHoras");

var valor = formulario.getRange("J4").getValue();
var valores = hojaActiva.getSheetByName("BDHORAS").getDataRange().getValues(); // Nombre de hoja donde se almacenan datos
  for (var i = 0; i < valores.length; i++) {
     var fila = valores[i];
    if (fila[NUM_COLUMNA_BUSQUEDA] == valor) {
      formulario.getRange("J15").setValue(fila[1]);
      
    }
  }
}


/* 
-------------------------------------------------Horarios-------------------------------------------------------------------
*/

function AgregarAsignacion() {
  
//guardamos en la variable la hoja de calculo activa 
  var hojaActiva = SpreadsheetApp.getActiveSpreadsheet();

  //Guardamos en las variable el nombre de la hojas
  var formulario = hojaActiva.getSheetByName("Asignaciones");
  var datos = hojaActiva.getSheetByName("BDasignaciones");

  //guardamos la informacion ingresada por el usuario
  var valores = [[formulario.getRange("C4" ).getValue(),
                  formulario.getRange("C5" ).getValue(),
                  formulario.getRange("C6" ).getValue(),
                  formulario.getRange("C7" ).getValue(),
                  formulario.getRange("C8" ).getValue(),
                  formulario.getRange("C10").getValue(),
                  formulario.getRange("C11").getValue(),]];

  // añadimos los valores creados a la hoja donde estan ubicados los datos
  datos.getRange(datos.getLastRow()+1,1,1,7).setValues(valores);

  CrearAsignacion();

}


/*
------------------------------------------------- Asignaciones -----------------------------------------------------------
*/

function CrearAsignacion() {

//guardamos en la variable la hoja de calculo activa 
  var hojaActiva = SpreadsheetApp.getActiveSpreadsheet();

  //Guardamos en las variable el nombre de la hojas
  var hojaFuente = hojaActiva.getSheetByName("Asignaciones");
  var hojaDestino = hojaActiva.getSheetByName("Horario");

  var dia = hojaFuente.getRange(4,3).getValue();
  var hora = hojaFuente.getRange(5,3).getValue();
  var sala = hojaFuente.getRange(6,3).getValue();
  var estado = hojaFuente.getRange(7,3).getValue();
  var profesor = hojaFuente.getRange(10,3).getValue();
  var materia = hojaFuente.getRange(11,3).getValue();


if(dia == "LUNES")
{

            if (sala == "SALA A")
            {
                  if(hora == "7:00 am - 9:00 am")
                  {
                        //nombre
                        hojaDestino.getRange(3,3).setValue(profesor);
                        //materia
                        hojaDestino.getRange(4,3).setValue(materia);
                        //estado
                        hojaDestino.getRange(5,3).setValue(estado);

                        Logger.log(profesor);
                        Logger.log(materia);
                        Logger.log(estado);


                  }

                  if(hora == "9:00 am - 11:00 am")
                  {
                        //nombre
                        hojaDestino.getRange(6,3).setValue(profesor);
                        //materia
                        hojaDestino.getRange(7,3).setValue(materia);
                        //estado
                        hojaDestino.getRange(8,3).setValue(estado);

                        Logger.log(profesor);
                        Logger.log(materia);
                        Logger.log(estado);

                  }

                  if(hora == "11:00 am - 1:00 pm")
                  {
                        //nombre
                        hojaDestino.getRange(9,3).setValue(profesor);
                        //materia
                        hojaDestino.getRange(10,3).setValue(materia);
                        //estado
                        hojaDestino.getRange(11,3).setValue(estado);

                        Logger.log(profesor);
                        Logger.log(materia);
                        Logger.log(estado);

                  }

                  if(hora == "2:00 pm - 4:00 pm")
                  {
                        //nombre
                        hojaDestino.getRange(12,3).setValue(profesor);
                        //materia
                        hojaDestino.getRange(13,3).setValue(materia);
                        //estado
                        hojaDestino.getRange(14,3).setValue(estado);

                        Logger.log(profesor);
                        Logger.log(materia);
                        Logger.log(estado);

                  }

                  if(hora == "4:00 pm - 6:00 pm")
                  {
                        //nombre
                        hojaDestino.getRange(15,3).setValue(profesor);
                        //materia
                        hojaDestino.getRange(16,3).setValue(materia);
                        //estado
                        hojaDestino.getRange(17,3).setValue(estado);

                        Logger.log(profesor);
                        Logger.log(materia);
                        Logger.log(estado);

                  }

                  if(hora == "6:00 pm - 8:00 pm")
                  {
                        //nombre
                        hojaDestino.getRange(18,3).setValue(profesor);
                        //materia
                        hojaDestino.getRange(19,3).setValue(materia);
                        //estado
                        hojaDestino.getRange(20,3).setValue(estado);

                        Logger.log(profesor);
                        Logger.log(materia);
                        Logger.log(estado);

                  }

                  if(hora == "8:00 pm - 10:00 pm")
                  {
                        //nombre
                        hojaDestino.getRange(21,3).setValue(profesor);
                        //materia
                        hojaDestino.getRange(22,3).setValue(materia);
                        //estado
                        hojaDestino.getRange(23,3).setValue(estado);

                        Logger.log(profesor);
                        Logger.log(materia);
                        Logger.log(estado);

                  }

                  


            }
            else if(sala == "SALA B")
            {
                  if(hora == "7:00 am - 9:00 am")
                  {
                        //nombre
                        hojaDestino.getRange(3,5).setValue(profesor);
                        //materia
                        hojaDestino.getRange(4,5).setValue(materia);
                        //estado
                        hojaDestino.getRange(5,5).setValue(estado);

                        Logger.log(profesor);
                        Logger.log(materia);
                        Logger.log(estado);


                  }

                  if(hora == "9:00 am - 11:00 am")
                  {
                        //nombre
                        hojaDestino.getRange(6,5).setValue(profesor);
                        //materia
                        hojaDestino.getRange(7,5).setValue(materia);
                        //estado
                        hojaDestino.getRange(8,5).setValue(estado);

                        Logger.log(profesor);
                        Logger.log(materia);
                        Logger.log(estado);

                  }

                  if(hora == "11:00 am - 1:00 pm")
                  {
                        //nombre
                        hojaDestino.getRange(9,5).setValue(profesor);
                        //materia
                        hojaDestino.getRange(10,5).setValue(materia);
                        //estado
                        hojaDestino.getRange(11,5).setValue(estado);

                        Logger.log(profesor);
                        Logger.log(materia);
                        Logger.log(estado);

                  }

                  if(hora == "2:00 pm - 4:00 pm")
                  {
                        //nombre
                        hojaDestino.getRange(12,5).setValue(profesor);
                        //materia
                        hojaDestino.getRange(13,5).setValue(materia);
                        //estado
                        hojaDestino.getRange(14,5).setValue(estado);

                        Logger.log(profesor);
                        Logger.log(materia);
                        Logger.log(estado);

                  }

                  if(hora == "4:00 pm - 6:00 pm")
                  {
                        //nombre
                        hojaDestino.getRange(15,5).setValue(profesor);
                        //materia
                        hojaDestino.getRange(16,5).setValue(materia);
                        //estado
                        hojaDestino.getRange(17,5).setValue(estado);

                        Logger.log(profesor);
                        Logger.log(materia);
                        Logger.log(estado);

                  }

                  if(hora == "6:00 pm - 8:00 pm")
                  {
                        //nombre
                        hojaDestino.getRange(18,5).setValue(profesor);
                        //materia
                        hojaDestino.getRange(19,5).setValue(materia);
                        //estado
                        hojaDestino.getRange(20,5).setValue(estado);

                        Logger.log(profesor);
                        Logger.log(materia);
                        Logger.log(estado);

                  }

                  if(hora == "8:00 pm - 10:00 pm")
                  {
                        //nombre
                        hojaDestino.getRange(21,5).setValue(profesor);
                        //materia
                        hojaDestino.getRange(22,5).setValue(materia);
                        //estado
                        hojaDestino.getRange(23,5).setValue(estado);

                        Logger.log(profesor);
                        Logger.log(materia);
                        Logger.log(estado);

                  }

            }
            else if(sala == "SALA C")
            {
                  if(hora == "7:00 am - 9:00 am")
                  {
                        //nombre
                        hojaDestino.getRange(3,7).setValue(profesor);
                        //materia
                        hojaDestino.getRange(4,7).setValue(materia);
                        //estado
                        hojaDestino.getRange(5,7).setValue(estado);

                        Logger.log(profesor);
                        Logger.log(materia);
                        Logger.log(estado);


                  }

                  if(hora == "9:00 am - 11:00 am")
                  {
                        //nombre
                        hojaDestino.getRange(6,7).setValue(profesor);
                        //materia
                        hojaDestino.getRange(7,7).setValue(materia);
                        //estado
                        hojaDestino.getRange(8,7).setValue(estado);

                        Logger.log(profesor);
                        Logger.log(materia);
                        Logger.log(estado);

                  }

                  if(hora == "11:00 am - 1:00 pm")
                  {
                        //nombre
                        hojaDestino.getRange(9,7).setValue(profesor);
                        //materia
                        hojaDestino.getRange(10,7).setValue(materia);
                        //estado
                        hojaDestino.getRange(11,7).setValue(estado);

                        Logger.log(profesor);
                        Logger.log(materia);
                        Logger.log(estado);

                  }

                  if(hora == "2:00 pm - 4:00 pm")
                  {
                        //nombre
                        hojaDestino.getRange(12,7).setValue(profesor);
                        //materia
                        hojaDestino.getRange(13,7).setValue(materia);
                        //estado
                        hojaDestino.getRange(14,7).setValue(estado);

                        Logger.log(profesor);
                        Logger.log(materia);
                        Logger.log(estado);

                  }

                  if(hora == "4:00 pm - 6:00 pm")
                  {
                        //nombre
                        hojaDestino.getRange(15,7).setValue(profesor);
                        //materia
                        hojaDestino.getRange(16,7).setValue(materia);
                        //estado
                        hojaDestino.getRange(17,7).setValue(estado);

                        Logger.log(profesor);
                        Logger.log(materia);
                        Logger.log(estado);

                  }

                  if(hora == "6:00 pm - 8:00 pm")
                  {
                        //nombre
                        hojaDestino.getRange(18,7).setValue(profesor);
                        //materia
                        hojaDestino.getRange(19,7).setValue(materia);
                        //estado
                        hojaDestino.getRange(20,7).setValue(estado);

                        Logger.log(profesor);
                        Logger.log(materia);
                        Logger.log(estado);

                  }

                  if(hora == "8:00 pm - 10:00 pm")
                  {
                        //nombre
                        hojaDestino.getRange(21,7).setValue(profesor);
                        //materia
                        hojaDestino.getRange(22,7).setValue(materia);
                        //estado
                        hojaDestino.getRange(23,7).setValue(estado);

                        Logger.log(profesor);
                        Logger.log(materia);
                        Logger.log(estado);

                  }

            }
            else if(sala == "SALA D")
            {
                  if(hora == "7:00 am - 9:00 am")
                  {
                        //nombre
                        hojaDestino.getRange(3,9).setValue(profesor);
                        //materia
                        hojaDestino.getRange(4,9).setValue(materia);
                        //estado
                        hojaDestino.getRange(5,9).setValue(estado);

                        Logger.log(profesor);
                        Logger.log(materia);
                        Logger.log(estado);


                  }

                  if(hora == "9:00 am - 11:00 am")
                  {
                        //nombre
                        hojaDestino.getRange(6,9).setValue(profesor);
                        //materia
                        hojaDestino.getRange(7,9).setValue(materia);
                        //estado
                        hojaDestino.getRange(8,9).setValue(estado);

                        Logger.log(profesor);
                        Logger.log(materia);
                        Logger.log(estado);

                  }

                  if(hora == "11:00 am - 1:00 pm")
                  {
                        //nombre
                        hojaDestino.getRange(9,9).setValue(profesor);
                        //materia
                        hojaDestino.getRange(10,9).setValue(materia);
                        //estado
                        hojaDestino.getRange(11,9).setValue(estado);

                        Logger.log(profesor);
                        Logger.log(materia);
                        Logger.log(estado);

                  }

                  if(hora == "2:00 pm - 4:00 pm")
                  {
                        //nombre
                        hojaDestino.getRange(12,9).setValue(profesor);
                        //materia
                        hojaDestino.getRange(13,9).setValue(materia);
                        //estado
                        hojaDestino.getRange(14,9).setValue(estado);

                        Logger.log(profesor);
                        Logger.log(materia);
                        Logger.log(estado);

                  }

                  if(hora == "4:00 pm - 6:00 pm")
                  {
                        //nombre
                        hojaDestino.getRange(15,9).setValue(profesor);
                        //materia
                        hojaDestino.getRange(16,9).setValue(materia);
                        //estado
                        hojaDestino.getRange(17,9).setValue(estado);

                        Logger.log(profesor);
                        Logger.log(materia);
                        Logger.log(estado);

                  }

                  if(hora == "6:00 pm - 8:00 pm")
                  {
                        //nombre
                        hojaDestino.getRange(18,9).setValue(profesor);
                        //materia
                        hojaDestino.getRange(19,9).setValue(materia);
                        //estado
                        hojaDestino.getRange(20,9).setValue(estado);

                        Logger.log(profesor);
                        Logger.log(materia);
                        Logger.log(estado);

                  }

                  if(hora == "8:00 pm - 10:00 pm")
                  {
                        //nombre
                        hojaDestino.getRange(21,9).setValue(profesor);
                        //materia
                        hojaDestino.getRange(22,9).setValue(materia);
                        //estado
                        hojaDestino.getRange(23,9).setValue(estado);

                        Logger.log(profesor);
                        Logger.log(materia);
                        Logger.log(estado);

                  }

            }
                 
}
else if(dia == "MARTES")
{

            if (sala == "SALA A")
            {
                  if(hora == "7:00 am - 9:00 am")
                  {
                        //nombre
                        hojaDestino.getRange(3,11).setValue(profesor);
                        //materia
                        hojaDestino.getRange(4,11).setValue(materia);
                        //estado
                        hojaDestino.getRange(5,11).setValue(estado);

                        Logger.log(profesor);
                        Logger.log(materia);
                        Logger.log(estado);


                  }

                  if(hora == "9:00 am - 11:00 am")
                  {
                        //nombre
                        hojaDestino.getRange(6,11).setValue(profesor);
                        //materia
                        hojaDestino.getRange(7,11).setValue(materia);
                        //estado
                        hojaDestino.getRange(8,11).setValue(estado);

                        Logger.log(profesor);
                        Logger.log(materia);
                        Logger.log(estado);

                  }

                  if(hora == "11:00 am - 1:00 pm")
                  {
                        //nombre
                        hojaDestino.getRange(9,11).setValue(profesor);
                        //materia
                        hojaDestino.getRange(10,11).setValue(materia);
                        //estado
                        hojaDestino.getRange(11,11).setValue(estado);

                        Logger.log(profesor);
                        Logger.log(materia);
                        Logger.log(estado);

                  }

                  if(hora == "2:00 pm - 4:00 pm")
                  {
                        //nombre
                        hojaDestino.getRange(12,11).setValue(profesor);
                        //materia
                        hojaDestino.getRange(13,11).setValue(materia);
                        //estado
                        hojaDestino.getRange(14,11).setValue(estado);

                        Logger.log(profesor);
                        Logger.log(materia);
                        Logger.log(estado);

                  }

                  if(hora == "4:00 pm - 6:00 pm")
                  {
                        //nombre
                        hojaDestino.getRange(15,11).setValue(profesor);
                        //materia
                        hojaDestino.getRange(16,11).setValue(materia);
                        //estado
                        hojaDestino.getRange(17,11).setValue(estado);

                        Logger.log(profesor);
                        Logger.log(materia);
                        Logger.log(estado);

                  }

                  if(hora == "6:00 pm - 8:00 pm")
                  {
                        //nombre
                        hojaDestino.getRange(18,11).setValue(profesor);
                        //materia
                        hojaDestino.getRange(19,11).setValue(materia);
                        //estado
                        hojaDestino.getRange(20,11).setValue(estado);

                        Logger.log(profesor);
                        Logger.log(materia);
                        Logger.log(estado);

                  }

                  if(hora == "8:00 pm - 10:00 pm")
                  {
                        //nombre
                        hojaDestino.getRange(21,11).setValue(profesor);
                        //materia
                        hojaDestino.getRange(22,11).setValue(materia);
                        //estado
                        hojaDestino.getRange(23,11).setValue(estado);

                        Logger.log(profesor);
                        Logger.log(materia);
                        Logger.log(estado);

                  }

                  


            }
            else if(sala == "SALA B")
            {
                  if(hora == "7:00 am - 9:00 am")
                  {
                        //nombre
                        hojaDestino.getRange(3,13).setValue(profesor);
                        //materia
                        hojaDestino.getRange(4,13).setValue(materia);
                        //estado
                        hojaDestino.getRange(5,13).setValue(estado);

                        Logger.log(profesor);
                        Logger.log(materia);
                        Logger.log(estado);


                  }

                  if(hora == "9:00 am - 11:00 am")
                  {
                        //nombre
                        hojaDestino.getRange(6,13).setValue(profesor);
                        //materia
                        hojaDestino.getRange(7,13).setValue(materia);
                        //estado
                        hojaDestino.getRange(8,13).setValue(estado);

                        Logger.log(profesor);
                        Logger.log(materia);
                        Logger.log(estado);

                  }

                  if(hora == "11:00 am - 1:00 pm")
                  {
                        //nombre
                        hojaDestino.getRange(9,13).setValue(profesor);
                        //materia
                        hojaDestino.getRange(10,13).setValue(materia);
                        //estado
                        hojaDestino.getRange(11,13).setValue(estado);

                        Logger.log(profesor);
                        Logger.log(materia);
                        Logger.log(estado);

                  }

                  if(hora == "2:00 pm - 4:00 pm")
                  {
                        //nombre
                        hojaDestino.getRange(12,13).setValue(profesor);
                        //materia
                        hojaDestino.getRange(13,13).setValue(materia);
                        //estado
                        hojaDestino.getRange(14,13).setValue(estado);

                        Logger.log(profesor);
                        Logger.log(materia);
                        Logger.log(estado);

                  }

                  if(hora == "4:00 pm - 6:00 pm")
                  {
                        //nombre
                        hojaDestino.getRange(15,13).setValue(profesor);
                        //materia
                        hojaDestino.getRange(16,13).setValue(materia);
                        //estado
                        hojaDestino.getRange(17,13).setValue(estado);

                        Logger.log(profesor);
                        Logger.log(materia);
                        Logger.log(estado);

                  }

                  if(hora == "6:00 pm - 8:00 pm")
                  {
                        //nombre
                        hojaDestino.getRange(18,13).setValue(profesor);
                        //materia
                        hojaDestino.getRange(19,13).setValue(materia);
                        //estado
                        hojaDestino.getRange(20,13).setValue(estado);

                        Logger.log(profesor);
                        Logger.log(materia);
                        Logger.log(estado);

                  }

                  if(hora == "8:00 pm - 10:00 pm")
                  {
                        //nombre
                        hojaDestino.getRange(21,13).setValue(profesor);
                        //materia
                        hojaDestino.getRange(22,13).setValue(materia);
                        //estado
                        hojaDestino.getRange(23,13).setValue(estado);

                        Logger.log(profesor);
                        Logger.log(materia);
                        Logger.log(estado);

                  }

            }
            else if(sala == "SALA C")
            {
                  if(hora == "7:00 am - 9:00 am")
                  {
                        //nombre
                        hojaDestino.getRange(3,15).setValue(profesor);
                        //materia
                        hojaDestino.getRange(4,15).setValue(materia);
                        //estado
                        hojaDestino.getRange(5,15).setValue(estado);

                        Logger.log(profesor);
                        Logger.log(materia);
                        Logger.log(estado);


                  }

                  if(hora == "9:00 am - 11:00 am")
                  {
                        //nombre
                        hojaDestino.getRange(6,15).setValue(profesor);
                        //materia
                        hojaDestino.getRange(7,15).setValue(materia);
                        //estado
                        hojaDestino.getRange(8,15).setValue(estado);

                        Logger.log(profesor);
                        Logger.log(materia);
                        Logger.log(estado);

                  }

                  if(hora == "11:00 am - 1:00 pm")
                  {
                        //nombre
                        hojaDestino.getRange(9,15).setValue(profesor);
                        //materia
                        hojaDestino.getRange(10,15).setValue(materia);
                        //estado
                        hojaDestino.getRange(11,15).setValue(estado);

                        Logger.log(profesor);
                        Logger.log(materia);
                        Logger.log(estado);

                  }

                  if(hora == "2:00 pm - 4:00 pm")
                  {
                        //nombre
                        hojaDestino.getRange(12,15).setValue(profesor);
                        //materia
                        hojaDestino.getRange(13,15).setValue(materia);
                        //estado
                        hojaDestino.getRange(14,15).setValue(estado);

                        Logger.log(profesor);
                        Logger.log(materia);
                        Logger.log(estado);

                  }

                  if(hora == "4:00 pm - 6:00 pm")
                  {
                        //nombre
                        hojaDestino.getRange(15,15).setValue(profesor);
                        //materia
                        hojaDestino.getRange(16,15).setValue(materia);
                        //estado
                        hojaDestino.getRange(17,15).setValue(estado);

                        Logger.log(profesor);
                        Logger.log(materia);
                        Logger.log(estado);

                  }

                  if(hora == "6:00 pm - 8:00 pm")
                  {
                        //nombre
                        hojaDestino.getRange(18,15).setValue(profesor);
                        //materia
                        hojaDestino.getRange(19,15).setValue(materia);
                        //estado
                        hojaDestino.getRange(20,15).setValue(estado);

                        Logger.log(profesor);
                        Logger.log(materia);
                        Logger.log(estado);

                  }

                  if(hora == "8:00 pm - 10:00 pm")
                  {
                        //nombre
                        hojaDestino.getRange(21,15).setValue(profesor);
                        //materia
                        hojaDestino.getRange(22,15).setValue(materia);
                        //estado
                        hojaDestino.getRange(23,15).setValue(estado);

                        Logger.log(profesor);
                        Logger.log(materia);
                        Logger.log(estado);

                  }

            }
            else if(sala == "SALA D")
            {
                  if(hora == "7:00 am - 9:00 am")
                  {
                        //nombre
                        hojaDestino.getRange(3,17).setValue(profesor);
                        //materia
                        hojaDestino.getRange(4,17).setValue(materia);
                        //estado
                        hojaDestino.getRange(5,17).setValue(estado);

                        Logger.log(profesor);
                        Logger.log(materia);
                        Logger.log(estado);


                  }

                  if(hora == "9:00 am - 11:00 am")
                  {
                        //nombre
                        hojaDestino.getRange(6,17).setValue(profesor);
                        //materia
                        hojaDestino.getRange(7,17).setValue(materia);
                        //estado
                        hojaDestino.getRange(8,17).setValue(estado);

                        Logger.log(profesor);
                        Logger.log(materia);
                        Logger.log(estado);

                  }

                  if(hora == "11:00 am - 1:00 pm")
                  {
                        //nombre
                        hojaDestino.getRange(9,17).setValue(profesor);
                        //materia
                        hojaDestino.getRange(10,17).setValue(materia);
                        //estado
                        hojaDestino.getRange(11,17).setValue(estado);

                        Logger.log(profesor);
                        Logger.log(materia);
                        Logger.log(estado);

                  }

                  if(hora == "2:00 pm - 4:00 pm")
                  {
                        //nombre
                        hojaDestino.getRange(12,17).setValue(profesor);
                        //materia
                        hojaDestino.getRange(13,17).setValue(materia);
                        //estado
                        hojaDestino.getRange(14,17).setValue(estado);

                        Logger.log(profesor);
                        Logger.log(materia);
                        Logger.log(estado);

                  }

                  if(hora == "4:00 pm - 6:00 pm")
                  {
                        //nombre
                        hojaDestino.getRange(15,17).setValue(profesor);
                        //materia
                        hojaDestino.getRange(16,17).setValue(materia);
                        //estado
                        hojaDestino.getRange(17,17).setValue(estado);

                        Logger.log(profesor);
                        Logger.log(materia);
                        Logger.log(estado);

                  }

                  if(hora == "6:00 pm - 8:00 pm")
                  {
                        //nombre
                        hojaDestino.getRange(18,17).setValue(profesor);
                        //materia
                        hojaDestino.getRange(19,17).setValue(materia);
                        //estado
                        hojaDestino.getRange(20,17).setValue(estado);

                        Logger.log(profesor);
                        Logger.log(materia);
                        Logger.log(estado);

                  }

                  if(hora == "8:00 pm - 10:00 pm")
                  {
                        //nombre
                        hojaDestino.getRange(21,17).setValue(profesor);
                        //materia
                        hojaDestino.getRange(22,17).setValue(materia);
                        //estado
                        hojaDestino.getRange(23,17).setValue(estado);

                        Logger.log(profesor);
                        Logger.log(materia);
                        Logger.log(estado);

                  }

            }
                 
}          
else if(dia == "MIERCOLES")
{

            if (sala == "SALA A")
            {
                  if(hora == "7:00 am - 9:00 am")
                  {
                        //nombre
                        hojaDestino.getRange(3,19).setValue(profesor);
                        //materia
                        hojaDestino.getRange(4,19).setValue(materia);
                        //estado
                        hojaDestino.getRange(5,19).setValue(estado);

                        Logger.log(profesor);
                        Logger.log(materia);
                        Logger.log(estado);


                  }

                  if(hora == "9:00 am - 11:00 am")
                  {
                        //nombre
                        hojaDestino.getRange(6,19).setValue(profesor);
                        //materia
                        hojaDestino.getRange(7,19).setValue(materia);
                        //estado
                        hojaDestino.getRange(8,19).setValue(estado);

                        Logger.log(profesor);
                        Logger.log(materia);
                        Logger.log(estado);

                  }

                  if(hora == "11:00 am - 1:00 pm")
                  {
                        //nombre
                        hojaDestino.getRange(9,19).setValue(profesor);
                        //materia
                        hojaDestino.getRange(10,19).setValue(materia);
                        //estado
                        hojaDestino.getRange(11,19).setValue(estado);

                        Logger.log(profesor);
                        Logger.log(materia);
                        Logger.log(estado);

                  }

                  if(hora == "2:00 pm - 4:00 pm")
                  {
                        //nombre
                        hojaDestino.getRange(12,19).setValue(profesor);
                        //materia
                        hojaDestino.getRange(13,19).setValue(materia);
                        //estado
                        hojaDestino.getRange(14,19).setValue(estado);

                        Logger.log(profesor);
                        Logger.log(materia);
                        Logger.log(estado);

                  }

                  if(hora == "4:00 pm - 6:00 pm")
                  {
                        //nombre
                        hojaDestino.getRange(15,19).setValue(profesor);
                        //materia
                        hojaDestino.getRange(16,19).setValue(materia);
                        //estado
                        hojaDestino.getRange(17,19).setValue(estado);

                        Logger.log(profesor);
                        Logger.log(materia);
                        Logger.log(estado);

                  }

                  if(hora == "6:00 pm - 8:00 pm")
                  {
                        //nombre
                        hojaDestino.getRange(18,19).setValue(profesor);
                        //materia
                        hojaDestino.getRange(19,19).setValue(materia);
                        //estado
                        hojaDestino.getRange(20,19).setValue(estado);

                        Logger.log(profesor);
                        Logger.log(materia);
                        Logger.log(estado);

                  }

                  if(hora == "8:00 pm - 10:00 pm")
                  {
                        //nombre
                        hojaDestino.getRange(21,19).setValue(profesor);
                        //materia
                        hojaDestino.getRange(22,19).setValue(materia);
                        //estado
                        hojaDestino.getRange(23,19).setValue(estado);

                        Logger.log(profesor);
                        Logger.log(materia);
                        Logger.log(estado);

                  }

                  


            }
            else if(sala == "SALA B")
            {
                  if(hora == "7:00 am - 9:00 am")
                  {
                        //nombre
                        hojaDestino.getRange(3,21).setValue(profesor);
                        //materia
                        hojaDestino.getRange(4,21).setValue(materia);
                        //estado
                        hojaDestino.getRange(5,21).setValue(estado);

                        Logger.log(profesor);
                        Logger.log(materia);
                        Logger.log(estado);


                  }

                  if(hora == "9:00 am - 11:00 am")
                  {
                        //nombre
                        hojaDestino.getRange(6,21).setValue(profesor);
                        //materia
                        hojaDestino.getRange(7,21).setValue(materia);
                        //estado
                        hojaDestino.getRange(8,21).setValue(estado);

                        Logger.log(profesor);
                        Logger.log(materia);
                        Logger.log(estado);

                  }

                  if(hora == "11:00 am - 1:00 pm")
                  {
                        //nombre
                        hojaDestino.getRange(9,21).setValue(profesor);
                        //materia
                        hojaDestino.getRange(10,21).setValue(materia);
                        //estado
                        hojaDestino.getRange(11,21).setValue(estado);

                        Logger.log(profesor);
                        Logger.log(materia);
                        Logger.log(estado);

                  }

                  if(hora == "2:00 pm - 4:00 pm")
                  {
                        //nombre
                        hojaDestino.getRange(12,21).setValue(profesor);
                        //materia
                        hojaDestino.getRange(13,21).setValue(materia);
                        //estado
                        hojaDestino.getRange(14,21).setValue(estado);

                        Logger.log(profesor);
                        Logger.log(materia);
                        Logger.log(estado);

                  }

                  if(hora == "4:00 pm - 6:00 pm")
                  {
                        //nombre
                        hojaDestino.getRange(15,21).setValue(profesor);
                        //materia
                        hojaDestino.getRange(16,21).setValue(materia);
                        //estado
                        hojaDestino.getRange(17,21).setValue(estado);

                        Logger.log(profesor);
                        Logger.log(materia);
                        Logger.log(estado);

                  }

                  if(hora == "6:00 pm - 8:00 pm")
                  {
                        //nombre
                        hojaDestino.getRange(18,21).setValue(profesor);
                        //materia
                        hojaDestino.getRange(19,21).setValue(materia);
                        //estado
                        hojaDestino.getRange(20,21).setValue(estado);

                        Logger.log(profesor);
                        Logger.log(materia);
                        Logger.log(estado);

                  }

                  if(hora == "8:00 pm - 10:00 pm")
                  {
                        //nombre
                        hojaDestino.getRange(21,21).setValue(profesor);
                        //materia
                        hojaDestino.getRange(22,21).setValue(materia);
                        //estado
                        hojaDestino.getRange(23,21).setValue(estado);

                        Logger.log(profesor);
                        Logger.log(materia);
                        Logger.log(estado);

                  }

            }
            else if(sala == "SALA C")
            {
                  if(hora == "7:00 am - 9:00 am")
                  {
                        //nombre
                        hojaDestino.getRange(3,23).setValue(profesor);
                        //materia
                        hojaDestino.getRange(4,23).setValue(materia);
                        //estado
                        hojaDestino.getRange(5,23).setValue(estado);

                        Logger.log(profesor);
                        Logger.log(materia);
                        Logger.log(estado);


                  }

                  if(hora == "9:00 am - 11:00 am")
                  {
                        //nombre
                        hojaDestino.getRange(6,23).setValue(profesor);
                        //materia
                        hojaDestino.getRange(7,23).setValue(materia);
                        //estado
                        hojaDestino.getRange(8,23).setValue(estado);

                        Logger.log(profesor);
                        Logger.log(materia);
                        Logger.log(estado);

                  }

                  if(hora == "11:00 am - 1:00 pm")
                  {
                        //nombre
                        hojaDestino.getRange(9,23).setValue(profesor);
                        //materia
                        hojaDestino.getRange(10,23).setValue(materia);
                        //estado
                        hojaDestino.getRange(11,23).setValue(estado);

                        Logger.log(profesor);
                        Logger.log(materia);
                        Logger.log(estado);

                  }

                  if(hora == "2:00 pm - 4:00 pm")
                  {
                        //nombre
                        hojaDestino.getRange(12,23).setValue(profesor);
                        //materia
                        hojaDestino.getRange(13,23).setValue(materia);
                        //estado
                        hojaDestino.getRange(14,23).setValue(estado);

                        Logger.log(profesor);
                        Logger.log(materia);
                        Logger.log(estado);

                  }

                  if(hora == "4:00 pm - 6:00 pm")
                  {
                        //nombre
                        hojaDestino.getRange(15,23).setValue(profesor);
                        //materia
                        hojaDestino.getRange(16,23).setValue(materia);
                        //estado
                        hojaDestino.getRange(17,23).setValue(estado);

                        Logger.log(profesor);
                        Logger.log(materia);
                        Logger.log(estado);

                  }

                  if(hora == "6:00 pm - 8:00 pm")
                  {
                        //nombre
                        hojaDestino.getRange(18,23).setValue(profesor);
                        //materia
                        hojaDestino.getRange(19,23).setValue(materia);
                        //estado
                        hojaDestino.getRange(20,23).setValue(estado);

                        Logger.log(profesor);
                        Logger.log(materia);
                        Logger.log(estado);

                  }

                  if(hora == "8:00 pm - 10:00 pm")
                  {
                        //nombre
                        hojaDestino.getRange(21,23).setValue(profesor);
                        //materia
                        hojaDestino.getRange(22,23).setValue(materia);
                        //estado
                        hojaDestino.getRange(23,23).setValue(estado);

                        Logger.log(profesor);
                        Logger.log(materia);
                        Logger.log(estado);

                  }

            }
            else if(sala == "SALA D")
            {
                  if(hora == "7:00 am - 9:00 am")
                  {
                        //nombre
                        hojaDestino.getRange(3,25).setValue(profesor);
                        //materia
                        hojaDestino.getRange(4,25).setValue(materia);
                        //estado
                        hojaDestino.getRange(5,25).setValue(estado);

                        Logger.log(profesor);
                        Logger.log(materia);
                        Logger.log(estado);


                  }

                  if(hora == "9:00 am - 11:00 am")
                  {
                        //nombre
                        hojaDestino.getRange(6,25).setValue(profesor);
                        //materia
                        hojaDestino.getRange(7,25).setValue(materia);
                        //estado
                        hojaDestino.getRange(8,25).setValue(estado);

                        Logger.log(profesor);
                        Logger.log(materia);
                        Logger.log(estado);

                  }

                  if(hora == "11:00 am - 1:00 pm")
                  {
                        //nombre
                        hojaDestino.getRange(9,25).setValue(profesor);
                        //materia
                        hojaDestino.getRange(10,25).setValue(materia);
                        //estado
                        hojaDestino.getRange(11,25).setValue(estado);

                        Logger.log(profesor);
                        Logger.log(materia);
                        Logger.log(estado);

                  }

                  if(hora == "2:00 pm - 4:00 pm")
                  {
                        //nombre
                        hojaDestino.getRange(12,25).setValue(profesor);
                        //materia
                        hojaDestino.getRange(13,25).setValue(materia);
                        //estado
                        hojaDestino.getRange(14,25).setValue(estado);

                        Logger.log(profesor);
                        Logger.log(materia);
                        Logger.log(estado);

                  }

                  if(hora == "4:00 pm - 6:00 pm")
                  {
                        //nombre
                        hojaDestino.getRange(15,25).setValue(profesor);
                        //materia
                        hojaDestino.getRange(16,25).setValue(materia);
                        //estado
                        hojaDestino.getRange(17,25).setValue(estado);

                        Logger.log(profesor);
                        Logger.log(materia);
                        Logger.log(estado);

                  }

                  if(hora == "6:00 pm - 8:00 pm")
                  {
                        //nombre
                        hojaDestino.getRange(18,25).setValue(profesor);
                        //materia
                        hojaDestino.getRange(19,25).setValue(materia);
                        //estado
                        hojaDestino.getRange(20,25).setValue(estado);

                        Logger.log(profesor);
                        Logger.log(materia);
                        Logger.log(estado);

                  }

                  if(hora == "8:00 pm - 10:00 pm")
                  {
                        //nombre
                        hojaDestino.getRange(21,25).setValue(profesor);
                        //materia
                        hojaDestino.getRange(22,25).setValue(materia);
                        //estado
                        hojaDestino.getRange(23,25).setValue(estado);

                        Logger.log(profesor);
                        Logger.log(materia);
                        Logger.log(estado);

                  }

            }
                 
}
else if(dia == "JUEVES")
{

            if (sala == "SALA A")
            {
                  if(hora == "7:00 am - 9:00 am")
                  {
                        //nombre
                        hojaDestino.getRange(3,27).setValue(profesor);
                        //materia
                        hojaDestino.getRange(4,27).setValue(materia);
                        //estado
                        hojaDestino.getRange(5,27).setValue(estado);

                        Logger.log(profesor);
                        Logger.log(materia);
                        Logger.log(estado);


                  }

                  if(hora == "9:00 am - 11:00 am")
                  {
                        //nombre
                        hojaDestino.getRange(6,27).setValue(profesor);
                        //materia
                        hojaDestino.getRange(7,27).setValue(materia);
                        //estado
                        hojaDestino.getRange(8,27).setValue(estado);

                        Logger.log(profesor);
                        Logger.log(materia);
                        Logger.log(estado);

                  }

                  if(hora == "11:00 am - 1:00 pm")
                  {
                        //nombre
                        hojaDestino.getRange(9,27).setValue(profesor);
                        //materia
                        hojaDestino.getRange(10,27).setValue(materia);
                        //estado
                        hojaDestino.getRange(11,27).setValue(estado);

                        Logger.log(profesor);
                        Logger.log(materia);
                        Logger.log(estado);

                  }

                  if(hora == "2:00 pm - 4:00 pm")
                  {
                        //nombre
                        hojaDestino.getRange(12,27).setValue(profesor);
                        //materia
                        hojaDestino.getRange(13,27).setValue(materia);
                        //estado
                        hojaDestino.getRange(14,27).setValue(estado);

                        Logger.log(profesor);
                        Logger.log(materia);
                        Logger.log(estado);

                  }

                  if(hora == "4:00 pm - 6:00 pm")
                  {
                        //nombre
                        hojaDestino.getRange(15,27).setValue(profesor);
                        //materia
                        hojaDestino.getRange(16,27).setValue(materia);
                        //estado
                        hojaDestino.getRange(17,27).setValue(estado);

                        Logger.log(profesor);
                        Logger.log(materia);
                        Logger.log(estado);

                  }

                  if(hora == "6:00 pm - 8:00 pm")
                  {
                        //nombre
                        hojaDestino.getRange(18,27).setValue(profesor);
                        //materia
                        hojaDestino.getRange(19,27).setValue(materia);
                        //estado
                        hojaDestino.getRange(20,27).setValue(estado);

                        Logger.log(profesor);
                        Logger.log(materia);
                        Logger.log(estado);

                  }

                  if(hora == "8:00 pm - 10:00 pm")
                  {
                        //nombre
                        hojaDestino.getRange(21,27).setValue(profesor);
                        //materia
                        hojaDestino.getRange(22,27).setValue(materia);
                        //estado
                        hojaDestino.getRange(23,27).setValue(estado);

                        Logger.log(profesor);
                        Logger.log(materia);
                        Logger.log(estado);

                  }

                  


            }
            else if(sala == "SALA B")
            {
                  if(hora == "7:00 am - 9:00 am")
                  {
                        //nombre
                        hojaDestino.getRange(3,29).setValue(profesor);
                        //materia
                        hojaDestino.getRange(4,29).setValue(materia);
                        //estado
                        hojaDestino.getRange(5,29).setValue(estado);

                        Logger.log(profesor);
                        Logger.log(materia);
                        Logger.log(estado);


                  }

                  if(hora == "9:00 am - 11:00 am")
                  {
                        //nombre
                        hojaDestino.getRange(6,29).setValue(profesor);
                        //materia
                        hojaDestino.getRange(7,29).setValue(materia);
                        //estado
                        hojaDestino.getRange(8,29).setValue(estado);

                        Logger.log(profesor);
                        Logger.log(materia);
                        Logger.log(estado);

                  }

                  if(hora == "11:00 am - 1:00 pm")
                  {
                        //nombre
                        hojaDestino.getRange(9,29).setValue(profesor);
                        //materia
                        hojaDestino.getRange(10,29).setValue(materia);
                        //estado
                        hojaDestino.getRange(11,29).setValue(estado);

                        Logger.log(profesor);
                        Logger.log(materia);
                        Logger.log(estado);

                  }

                  if(hora == "2:00 pm - 4:00 pm")
                  {
                        //nombre
                        hojaDestino.getRange(12,29).setValue(profesor);
                        //materia
                        hojaDestino.getRange(13,29).setValue(materia);
                        //estado
                        hojaDestino.getRange(14,29).setValue(estado);

                        Logger.log(profesor);
                        Logger.log(materia);
                        Logger.log(estado);

                  }

                  if(hora == "4:00 pm - 6:00 pm")
                  {
                        //nombre
                        hojaDestino.getRange(15,29).setValue(profesor);
                        //materia
                        hojaDestino.getRange(16,29).setValue(materia);
                        //estado
                        hojaDestino.getRange(17,29).setValue(estado);

                        Logger.log(profesor);
                        Logger.log(materia);
                        Logger.log(estado);

                  }

                  if(hora == "6:00 pm - 8:00 pm")
                  {
                        //nombre
                        hojaDestino.getRange(18,29).setValue(profesor);
                        //materia
                        hojaDestino.getRange(19,29).setValue(materia);
                        //estado
                        hojaDestino.getRange(20,29).setValue(estado);

                        Logger.log(profesor);
                        Logger.log(materia);
                        Logger.log(estado);

                  }

                  if(hora == "8:00 pm - 10:00 pm")
                  {
                        //nombre
                        hojaDestino.getRange(21,29).setValue(profesor);
                        //materia
                        hojaDestino.getRange(22,29).setValue(materia);
                        //estado
                        hojaDestino.getRange(23,29).setValue(estado);

                        Logger.log(profesor);
                        Logger.log(materia);
                        Logger.log(estado);

                  }

            }
            else if(sala == "SALA C")
            {
                  if(hora == "7:00 am - 9:00 am")
                  {
                        //nombre
                        hojaDestino.getRange(3,31).setValue(profesor);
                        //materia
                        hojaDestino.getRange(4,31).setValue(materia);
                        //estado
                        hojaDestino.getRange(5,31).setValue(estado);

                        Logger.log(profesor);
                        Logger.log(materia);
                        Logger.log(estado);


                  }

                  if(hora == "9:00 am - 11:00 am")
                  {
                        //nombre
                        hojaDestino.getRange(6,31).setValue(profesor);
                        //materia
                        hojaDestino.getRange(7,31).setValue(materia);
                        //estado
                        hojaDestino.getRange(8,31).setValue(estado);

                        Logger.log(profesor);
                        Logger.log(materia);
                        Logger.log(estado);

                  }

                  if(hora == "11:00 am - 1:00 pm")
                  {
                        //nombre
                        hojaDestino.getRange(9,31).setValue(profesor);
                        //materia
                        hojaDestino.getRange(10,31).setValue(materia);
                        //estado
                        hojaDestino.getRange(11,31).setValue(estado);

                        Logger.log(profesor);
                        Logger.log(materia);
                        Logger.log(estado);

                  }

                  if(hora == "2:00 pm - 4:00 pm")
                  {
                        //nombre
                        hojaDestino.getRange(12,31).setValue(profesor);
                        //materia
                        hojaDestino.getRange(13,31).setValue(materia);
                        //estado
                        hojaDestino.getRange(14,31).setValue(estado);

                        Logger.log(profesor);
                        Logger.log(materia);
                        Logger.log(estado);

                  }

                  if(hora == "4:00 pm - 6:00 pm")
                  {
                        //nombre
                        hojaDestino.getRange(15,31).setValue(profesor);
                        //materia
                        hojaDestino.getRange(16,31).setValue(materia);
                        //estado
                        hojaDestino.getRange(17,31).setValue(estado);

                        Logger.log(profesor);
                        Logger.log(materia);
                        Logger.log(estado);

                  }

                  if(hora == "6:00 pm - 8:00 pm")
                  {
                        //nombre
                        hojaDestino.getRange(18,31).setValue(profesor);
                        //materia
                        hojaDestino.getRange(19,31).setValue(materia);
                        //estado
                        hojaDestino.getRange(20,31).setValue(estado);

                        Logger.log(profesor);
                        Logger.log(materia);
                        Logger.log(estado);

                  }

                  if(hora == "8:00 pm - 10:00 pm")
                  {
                        //nombre
                        hojaDestino.getRange(21,31).setValue(profesor);
                        //materia
                        hojaDestino.getRange(22,31).setValue(materia);
                        //estado
                        hojaDestino.getRange(23,31).setValue(estado);

                        Logger.log(profesor);
                        Logger.log(materia);
                        Logger.log(estado);

                  }

            }
            else if(sala == "SALA D")
            {
                  if(hora == "7:00 am - 9:00 am")
                  {
                        //nombre
                        hojaDestino.getRange(3,33).setValue(profesor);
                        //materia
                        hojaDestino.getRange(4,33).setValue(materia);
                        //estado
                        hojaDestino.getRange(5,33).setValue(estado);

                        Logger.log(profesor);
                        Logger.log(materia);
                        Logger.log(estado);


                  }

                  if(hora == "9:00 am - 11:00 am")
                  {
                        //nombre
                        hojaDestino.getRange(6,33).setValue(profesor);
                        //materia
                        hojaDestino.getRange(7,33).setValue(materia);
                        //estado
                        hojaDestino.getRange(8,33).setValue(estado);

                        Logger.log(profesor);
                        Logger.log(materia);
                        Logger.log(estado);

                  }

                  if(hora == "11:00 am - 1:00 pm")
                  {
                        //nombre
                        hojaDestino.getRange(9,33).setValue(profesor);
                        //materia
                        hojaDestino.getRange(10,33).setValue(materia);
                        //estado
                        hojaDestino.getRange(11,33).setValue(estado);

                        Logger.log(profesor);
                        Logger.log(materia);
                        Logger.log(estado);

                  }

                  if(hora == "2:00 pm - 4:00 pm")
                  {
                        //nombre
                        hojaDestino.getRange(12,33).setValue(profesor);
                        //materia
                        hojaDestino.getRange(13,33).setValue(materia);
                        //estado
                        hojaDestino.getRange(14,33).setValue(estado);

                        Logger.log(profesor);
                        Logger.log(materia);
                        Logger.log(estado);

                  }

                  if(hora == "4:00 pm - 6:00 pm")
                  {
                        //nombre
                        hojaDestino.getRange(15,33).setValue(profesor);
                        //materia
                        hojaDestino.getRange(16,33).setValue(materia);
                        //estado
                        hojaDestino.getRange(17,33).setValue(estado);

                        Logger.log(profesor);
                        Logger.log(materia);
                        Logger.log(estado);

                  }

                  if(hora == "6:00 pm - 8:00 pm")
                  {
                        //nombre
                        hojaDestino.getRange(18,33).setValue(profesor);
                        //materia
                        hojaDestino.getRange(19,33).setValue(materia);
                        //estado
                        hojaDestino.getRange(20,33).setValue(estado);

                        Logger.log(profesor);
                        Logger.log(materia);
                        Logger.log(estado);

                  }

                  if(hora == "8:00 pm - 10:00 pm")
                  {
                        //nombre
                        hojaDestino.getRange(21,33).setValue(profesor);
                        //materia
                        hojaDestino.getRange(22,33).setValue(materia);
                        //estado
                        hojaDestino.getRange(23,33).setValue(estado);

                        Logger.log(profesor);
                        Logger.log(materia);
                        Logger.log(estado);

                  }

            }
                 
}
else if(dia == "VIERNES")
{

            if (sala == "SALA A")
            {
                  if(hora == "7:00 am - 9:00 am")
                  {
                        //nombre
                        hojaDestino.getRange(3,35).setValue(profesor);
                        //materia
                        hojaDestino.getRange(4,35).setValue(materia);
                        //estado
                        hojaDestino.getRange(5,35).setValue(estado);

                        Logger.log(profesor);
                        Logger.log(materia);
                        Logger.log(estado);


                  }

                  if(hora == "9:00 am - 11:00 am")
                  {
                        //nombre
                        hojaDestino.getRange(6,35).setValue(profesor);
                        //materia
                        hojaDestino.getRange(7,35).setValue(materia);
                        //estado
                        hojaDestino.getRange(8,35).setValue(estado);

                        Logger.log(profesor);
                        Logger.log(materia);
                        Logger.log(estado);

                  }

                  if(hora == "11:00 am - 1:00 pm")
                  {
                        //nombre
                        hojaDestino.getRange(9,35).setValue(profesor);
                        //materia
                        hojaDestino.getRange(10,35).setValue(materia);
                        //estado
                        hojaDestino.getRange(11,35).setValue(estado);

                        Logger.log(profesor);
                        Logger.log(materia);
                        Logger.log(estado);

                  }

                  if(hora == "2:00 pm - 4:00 pm")
                  {
                        //nombre
                        hojaDestino.getRange(12,35).setValue(profesor);
                        //materia
                        hojaDestino.getRange(13,35).setValue(materia);
                        //estado
                        hojaDestino.getRange(14,35).setValue(estado);

                        Logger.log(profesor);
                        Logger.log(materia);
                        Logger.log(estado);

                  }

                  if(hora == "4:00 pm - 6:00 pm")
                  {
                        //nombre
                        hojaDestino.getRange(15,35).setValue(profesor);
                        //materia
                        hojaDestino.getRange(16,35).setValue(materia);
                        //estado
                        hojaDestino.getRange(17,35).setValue(estado);

                        Logger.log(profesor);
                        Logger.log(materia);
                        Logger.log(estado);

                  }

                  if(hora == "6:00 pm - 8:00 pm")
                  {
                        //nombre
                        hojaDestino.getRange(18,35).setValue(profesor);
                        //materia
                        hojaDestino.getRange(19,35).setValue(materia);
                        //estado
                        hojaDestino.getRange(20,35).setValue(estado);

                        Logger.log(profesor);
                        Logger.log(materia);
                        Logger.log(estado);

                  }

                  if(hora == "8:00 pm - 10:00 pm")
                  {
                        //nombre
                        hojaDestino.getRange(21,35).setValue(profesor);
                        //materia
                        hojaDestino.getRange(22,35).setValue(materia);
                        //estado
                        hojaDestino.getRange(23,35).setValue(estado);

                        Logger.log(profesor);
                        Logger.log(materia);
                        Logger.log(estado);

                  }

                  


            }
            else if(sala == "SALA B")
            {
                  if(hora == "7:00 am - 9:00 am")
                  {
                        //nombre
                        hojaDestino.getRange(3,37).setValue(profesor);
                        //materia
                        hojaDestino.getRange(4,37).setValue(materia);
                        //estado
                        hojaDestino.getRange(5,37).setValue(estado);

                        Logger.log(profesor);
                        Logger.log(materia);
                        Logger.log(estado);


                  }

                  if(hora == "9:00 am - 11:00 am")
                  {
                        //nombre
                        hojaDestino.getRange(6,37).setValue(profesor);
                        //materia
                        hojaDestino.getRange(7,37).setValue(materia);
                        //estado
                        hojaDestino.getRange(8,37).setValue(estado);

                        Logger.log(profesor);
                        Logger.log(materia);
                        Logger.log(estado);

                  }

                  if(hora == "11:00 am - 1:00 pm")
                  {
                        //nombre
                        hojaDestino.getRange(9,37).setValue(profesor);
                        //materia
                        hojaDestino.getRange(10,37).setValue(materia);
                        //estado
                        hojaDestino.getRange(11,37).setValue(estado);

                        Logger.log(profesor);
                        Logger.log(materia);
                        Logger.log(estado);

                  }

                  if(hora == "2:00 pm - 4:00 pm")
                  {
                        //nombre
                        hojaDestino.getRange(12,37).setValue(profesor);
                        //materia
                        hojaDestino.getRange(13,37).setValue(materia);
                        //estado
                        hojaDestino.getRange(14,37).setValue(estado);

                        Logger.log(profesor);
                        Logger.log(materia);
                        Logger.log(estado);

                  }

                  if(hora == "4:00 pm - 6:00 pm")
                  {
                        //nombre
                        hojaDestino.getRange(15,37).setValue(profesor);
                        //materia
                        hojaDestino.getRange(16,37).setValue(materia);
                        //estado
                        hojaDestino.getRange(17,37).setValue(estado);

                        Logger.log(profesor);
                        Logger.log(materia);
                        Logger.log(estado);

                  }

                  if(hora == "6:00 pm - 8:00 pm")
                  {
                        //nombre
                        hojaDestino.getRange(18,37).setValue(profesor);
                        //materia
                        hojaDestino.getRange(19,37).setValue(materia);
                        //estado
                        hojaDestino.getRange(20,37).setValue(estado);

                        Logger.log(profesor);
                        Logger.log(materia);
                        Logger.log(estado);

                  }

                  if(hora == "8:00 pm - 10:00 pm")
                  {
                        //nombre
                        hojaDestino.getRange(21,37).setValue(profesor);
                        //materia
                        hojaDestino.getRange(22,37).setValue(materia);
                        //estado
                        hojaDestino.getRange(23,37).setValue(estado);

                        Logger.log(profesor);
                        Logger.log(materia);
                        Logger.log(estado);

                  }

            }
            else if(sala == "SALA C")
            {
                  if(hora == "7:00 am - 9:00 am")
                  {
                        //nombre
                        hojaDestino.getRange(3,39).setValue(profesor);
                        //materia
                        hojaDestino.getRange(4,39).setValue(materia);
                        //estado
                        hojaDestino.getRange(5,39).setValue(estado);

                        Logger.log(profesor);
                        Logger.log(materia);
                        Logger.log(estado);


                  }

                  if(hora == "9:00 am - 11:00 am")
                  {
                        //nombre
                        hojaDestino.getRange(6,39).setValue(profesor);
                        //materia
                        hojaDestino.getRange(7,39).setValue(materia);
                        //estado
                        hojaDestino.getRange(8,39).setValue(estado);

                        Logger.log(profesor);
                        Logger.log(materia);
                        Logger.log(estado);

                  }

                  if(hora == "11:00 am - 1:00 pm")
                  {
                        //nombre
                        hojaDestino.getRange(9,39).setValue(profesor);
                        //materia
                        hojaDestino.getRange(10,39).setValue(materia);
                        //estado
                        hojaDestino.getRange(11,39).setValue(estado);

                        Logger.log(profesor);
                        Logger.log(materia);
                        Logger.log(estado);

                  }

                  if(hora == "2:00 pm - 4:00 pm")
                  {
                        //nombre
                        hojaDestino.getRange(12,39).setValue(profesor);
                        //materia
                        hojaDestino.getRange(13,39).setValue(materia);
                        //estado
                        hojaDestino.getRange(14,39).setValue(estado);

                        Logger.log(profesor);
                        Logger.log(materia);
                        Logger.log(estado);

                  }

                  if(hora == "4:00 pm - 6:00 pm")
                  {
                        //nombre
                        hojaDestino.getRange(15,39).setValue(profesor);
                        //materia
                        hojaDestino.getRange(16,39).setValue(materia);
                        //estado
                        hojaDestino.getRange(17,39).setValue(estado);

                        Logger.log(profesor);
                        Logger.log(materia);
                        Logger.log(estado);

                  }

                  if(hora == "6:00 pm - 8:00 pm")
                  {
                        //nombre
                        hojaDestino.getRange(18,39).setValue(profesor);
                        //materia
                        hojaDestino.getRange(19,39).setValue(materia);
                        //estado
                        hojaDestino.getRange(20,39).setValue(estado);

                        Logger.log(profesor);
                        Logger.log(materia);
                        Logger.log(estado);

                  }

                  if(hora == "8:00 pm - 10:00 pm")
                  {
                        //nombre
                        hojaDestino.getRange(21,39).setValue(profesor);
                        //materia
                        hojaDestino.getRange(22,39).setValue(materia);
                        //estado
                        hojaDestino.getRange(23,39).setValue(estado);

                        Logger.log(profesor);
                        Logger.log(materia);
                        Logger.log(estado);

                  }

            }
            else if(sala == "SALA D")
            {
                  if(hora == "7:00 am - 9:00 am")
                  {
                        //nombre
                        hojaDestino.getRange(3,41).setValue(profesor);
                        //materia
                        hojaDestino.getRange(4,41).setValue(materia);
                        //estado
                        hojaDestino.getRange(5,41).setValue(estado);

                        Logger.log(profesor);
                        Logger.log(materia);
                        Logger.log(estado);


                  }

                  if(hora == "9:00 am - 11:00 am")
                  {
                        //nombre
                        hojaDestino.getRange(6,41).setValue(profesor);
                        //materia
                        hojaDestino.getRange(7,41).setValue(materia);
                        //estado
                        hojaDestino.getRange(8,41).setValue(estado);

                        Logger.log(profesor);
                        Logger.log(materia);
                        Logger.log(estado);

                  }

                  if(hora == "11:00 am - 1:00 pm")
                  {
                        //nombre
                        hojaDestino.getRange(9,41).setValue(profesor);
                        //materia
                        hojaDestino.getRange(10,41).setValue(materia);
                        //estado
                        hojaDestino.getRange(11,41).setValue(estado);

                        Logger.log(profesor);
                        Logger.log(materia);
                        Logger.log(estado);

                  }

                  if(hora == "2:00 pm - 4:00 pm")
                  {
                        //nombre
                        hojaDestino.getRange(12,41).setValue(profesor);
                        //materia
                        hojaDestino.getRange(13,41).setValue(materia);
                        //estado
                        hojaDestino.getRange(14,41).setValue(estado);

                        Logger.log(profesor);
                        Logger.log(materia);
                        Logger.log(estado);

                  }

                  if(hora == "4:00 pm - 6:00 pm")
                  {
                        //nombre
                        hojaDestino.getRange(15,41).setValue(profesor);
                        //materia
                        hojaDestino.getRange(16,41).setValue(materia);
                        //estado
                        hojaDestino.getRange(17,41).setValue(estado);

                        Logger.log(profesor);
                        Logger.log(materia);
                        Logger.log(estado);

                  }

                  if(hora == "6:00 pm - 8:00 pm")
                  {
                        //nombre
                        hojaDestino.getRange(18,41).setValue(profesor);
                        //materia
                        hojaDestino.getRange(19,41).setValue(materia);
                        //estado
                        hojaDestino.getRange(20,41).setValue(estado);

                        Logger.log(profesor);
                        Logger.log(materia);
                        Logger.log(estado);

                  }

                  if(hora == "8:00 pm - 10:00 pm")
                  {
                        //nombre
                        hojaDestino.getRange(21,41).setValue(profesor);
                        //materia
                        hojaDestino.getRange(22,41).setValue(materia);
                        //estado
                        hojaDestino.getRange(23,41).setValue(estado);

                        Logger.log(profesor);
                        Logger.log(materia);
                        Logger.log(estado);

                  }

            }
                 
}
else if(dia == "SABADO")
{

            if (sala == "SALA A")
            {
                  if(hora == "7:00 am - 9:00 am")
                  {
                        //nombre
                        hojaDestino.getRange(3,43).setValue(profesor);
                        //materia
                        hojaDestino.getRange(4,43).setValue(materia);
                        //estado
                        hojaDestino.getRange(5,43).setValue(estado);

                        Logger.log(profesor);
                        Logger.log(materia);
                        Logger.log(estado);


                  }

                  if(hora == "9:00 am - 11:00 am")
                  {
                        //nombre
                        hojaDestino.getRange(6,43).setValue(profesor);
                        //materia
                        hojaDestino.getRange(7,43).setValue(materia);
                        //estado
                        hojaDestino.getRange(8,43).setValue(estado);

                        Logger.log(profesor);
                        Logger.log(materia);
                        Logger.log(estado);

                  }

                  if(hora == "11:00 am - 1:00 pm")
                  {
                        //nombre
                        hojaDestino.getRange(9,43).setValue(profesor);
                        //materia
                        hojaDestino.getRange(10,43).setValue(materia);
                        //estado
                        hojaDestino.getRange(11,43).setValue(estado);

                        Logger.log(profesor);
                        Logger.log(materia);
                        Logger.log(estado);

                  }

                  if(hora == "2:00 pm - 4:00 pm")
                  {
                        //nombre
                        hojaDestino.getRange(12,43).setValue(profesor);
                        //materia
                        hojaDestino.getRange(13,43).setValue(materia);
                        //estado
                        hojaDestino.getRange(14,43).setValue(estado);

                        Logger.log(profesor);
                        Logger.log(materia);
                        Logger.log(estado);

                  }

                  if(hora == "4:00 pm - 6:00 pm")
                  {
                        //nombre
                        hojaDestino.getRange(15,43).setValue(profesor);
                        //materia
                        hojaDestino.getRange(16,43).setValue(materia);
                        //estado
                        hojaDestino.getRange(17,43).setValue(estado);

                        Logger.log(profesor);
                        Logger.log(materia);
                        Logger.log(estado);

                  }

                  if(hora == "6:00 pm - 8:00 pm")
                  {
                        //nombre
                        hojaDestino.getRange(18,43).setValue(profesor);
                        //materia
                        hojaDestino.getRange(19,43).setValue(materia);
                        //estado
                        hojaDestino.getRange(20,43).setValue(estado);

                        Logger.log(profesor);
                        Logger.log(materia);
                        Logger.log(estado);

                  }

                  if(hora == "8:00 pm - 10:00 pm")
                  {
                        //nombre
                        hojaDestino.getRange(21,43).setValue(profesor);
                        //materia
                        hojaDestino.getRange(22,43).setValue(materia);
                        //estado
                        hojaDestino.getRange(23,43).setValue(estado);

                        Logger.log(profesor);
                        Logger.log(materia);
                        Logger.log(estado);

                  }

                  


            }
            else if(sala == "SALA B")
            {
                  if(hora == "7:00 am - 9:00 am")
                  {
                        //nombre
                        hojaDestino.getRange(3,45).setValue(profesor);
                        //materia
                        hojaDestino.getRange(4,45).setValue(materia);
                        //estado
                        hojaDestino.getRange(5,45).setValue(estado);

                        Logger.log(profesor);
                        Logger.log(materia);
                        Logger.log(estado);


                  }

                  if(hora == "9:00 am - 11:00 am")
                  {
                        //nombre
                        hojaDestino.getRange(6,45).setValue(profesor);
                        //materia
                        hojaDestino.getRange(7,45).setValue(materia);
                        //estado
                        hojaDestino.getRange(8,45).setValue(estado);

                        Logger.log(profesor);
                        Logger.log(materia);
                        Logger.log(estado);

                  }

                  if(hora == "11:00 am - 1:00 pm")
                  {
                        //nombre
                        hojaDestino.getRange(9,45).setValue(profesor);
                        //materia
                        hojaDestino.getRange(10,45).setValue(materia);
                        //estado
                        hojaDestino.getRange(11,45).setValue(estado);

                        Logger.log(profesor);
                        Logger.log(materia);
                        Logger.log(estado);

                  }

                  if(hora == "2:00 pm - 4:00 pm")
                  {
                        //nombre
                        hojaDestino.getRange(12,45).setValue(profesor);
                        //materia
                        hojaDestino.getRange(13,45).setValue(materia);
                        //estado
                        hojaDestino.getRange(14,45).setValue(estado);

                        Logger.log(profesor);
                        Logger.log(materia);
                        Logger.log(estado);

                  }

                  if(hora == "4:00 pm - 6:00 pm")
                  {
                        //nombre
                        hojaDestino.getRange(15,45).setValue(profesor);
                        //materia
                        hojaDestino.getRange(16,45).setValue(materia);
                        //estado
                        hojaDestino.getRange(17,45).setValue(estado);

                        Logger.log(profesor);
                        Logger.log(materia);
                        Logger.log(estado);

                  }

                  if(hora == "6:00 pm - 8:00 pm")
                  {
                        //nombre
                        hojaDestino.getRange(18,45).setValue(profesor);
                        //materia
                        hojaDestino.getRange(19,45).setValue(materia);
                        //estado
                        hojaDestino.getRange(20,45).setValue(estado);

                        Logger.log(profesor);
                        Logger.log(materia);
                        Logger.log(estado);

                  }

                  if(hora == "8:00 pm - 10:00 pm")
                  {
                        //nombre
                        hojaDestino.getRange(21,45).setValue(profesor);
                        //materia
                        hojaDestino.getRange(22,45).setValue(materia);
                        //estado
                        hojaDestino.getRange(23,45).setValue(estado);

                        Logger.log(profesor);
                        Logger.log(materia);
                        Logger.log(estado);

                  }

            }
            else if(sala == "SALA C")
            {
                  if(hora == "7:00 am - 9:00 am")
                  {
                        //nombre
                        hojaDestino.getRange(3,47).setValue(profesor);
                        //materia
                        hojaDestino.getRange(4,47).setValue(materia);
                        //estado
                        hojaDestino.getRange(5,47).setValue(estado);

                        Logger.log(profesor);
                        Logger.log(materia);
                        Logger.log(estado);


                  }

                  if(hora == "9:00 am - 11:00 am")
                  {
                        //nombre
                        hojaDestino.getRange(6,47).setValue(profesor);
                        //materia
                        hojaDestino.getRange(7,47).setValue(materia);
                        //estado
                        hojaDestino.getRange(8,47).setValue(estado);

                        Logger.log(profesor);
                        Logger.log(materia);
                        Logger.log(estado);

                  }

                  if(hora == "11:00 am - 1:00 pm")
                  {
                        //nombre
                        hojaDestino.getRange(9,47).setValue(profesor);
                        //materia
                        hojaDestino.getRange(10,47).setValue(materia);
                        //estado
                        hojaDestino.getRange(11,47).setValue(estado);

                        Logger.log(profesor);
                        Logger.log(materia);
                        Logger.log(estado);

                  }

                  if(hora == "2:00 pm - 4:00 pm")
                  {
                        //nombre
                        hojaDestino.getRange(12,47).setValue(profesor);
                        //materia
                        hojaDestino.getRange(13,47).setValue(materia);
                        //estado
                        hojaDestino.getRange(14,47).setValue(estado);

                        Logger.log(profesor);
                        Logger.log(materia);
                        Logger.log(estado);

                  }

                  if(hora == "4:00 pm - 6:00 pm")
                  {
                        //nombre
                        hojaDestino.getRange(15,47).setValue(profesor);
                        //materia
                        hojaDestino.getRange(16,47).setValue(materia);
                        //estado
                        hojaDestino.getRange(17,47).setValue(estado);

                        Logger.log(profesor);
                        Logger.log(materia);
                        Logger.log(estado);

                  }

                  if(hora == "6:00 pm - 8:00 pm")
                  {
                        //nombre
                        hojaDestino.getRange(18,47).setValue(profesor);
                        //materia
                        hojaDestino.getRange(19,47).setValue(materia);
                        //estado
                        hojaDestino.getRange(20,47).setValue(estado);

                        Logger.log(profesor);
                        Logger.log(materia);
                        Logger.log(estado);

                  }

                  if(hora == "8:00 pm - 10:00 pm")
                  {
                        //nombre
                        hojaDestino.getRange(21,47).setValue(profesor);
                        //materia
                        hojaDestino.getRange(22,47).setValue(materia);
                        //estado
                        hojaDestino.getRange(23,47).setValue(estado);

                        Logger.log(profesor);
                        Logger.log(materia);
                        Logger.log(estado);

                  }

            }
            else if(sala == "SALA D")
            {
                  if(hora == "7:00 am - 9:00 am")
                  {
                        //nombre
                        hojaDestino.getRange(3,49).setValue(profesor);
                        //materia
                        hojaDestino.getRange(4,49).setValue(materia);
                        //estado
                        hojaDestino.getRange(5,49).setValue(estado);

                        Logger.log(profesor);
                        Logger.log(materia);
                        Logger.log(estado);


                  }

                  if(hora == "9:00 am - 11:00 am")
                  {
                        //nombre
                        hojaDestino.getRange(6,49).setValue(profesor);
                        //materia
                        hojaDestino.getRange(7,49).setValue(materia);
                        //estado
                        hojaDestino.getRange(8,49).setValue(estado);

                        Logger.log(profesor);
                        Logger.log(materia);
                        Logger.log(estado);

                  }

                  if(hora == "11:00 am - 1:00 pm")
                  {
                        //nombre
                        hojaDestino.getRange(9,49).setValue(profesor);
                        //materia
                        hojaDestino.getRange(10,49).setValue(materia);
                        //estado
                        hojaDestino.getRange(11,49).setValue(estado);

                        Logger.log(profesor);
                        Logger.log(materia);
                        Logger.log(estado);

                  }

                  if(hora == "2:00 pm - 4:00 pm")
                  {
                        //nombre
                        hojaDestino.getRange(12,49).setValue(profesor);
                        //materia
                        hojaDestino.getRange(13,49).setValue(materia);
                        //estado
                        hojaDestino.getRange(14,49).setValue(estado);

                        Logger.log(profesor);
                        Logger.log(materia);
                        Logger.log(estado);

                  }

                  if(hora == "4:00 pm - 6:00 pm")
                  {
                        //nombre
                        hojaDestino.getRange(15,49).setValue(profesor);
                        //materia
                        hojaDestino.getRange(16,49).setValue(materia);
                        //estado
                        hojaDestino.getRange(17,49).setValue(estado);

                        Logger.log(profesor);
                        Logger.log(materia);
                        Logger.log(estado);

                  }

                  if(hora == "6:00 pm - 8:00 pm")
                  {
                        //nombre
                        hojaDestino.getRange(18,49).setValue(profesor);
                        //materia
                        hojaDestino.getRange(19,49).setValue(materia);
                        //estado
                        hojaDestino.getRange(20,49).setValue(estado);

                        Logger.log(profesor);
                        Logger.log(materia);
                        Logger.log(estado);

                  }

                  if(hora == "8:00 pm - 10:00 pm")
                  {
                        //nombre
                        hojaDestino.getRange(21,49).setValue(profesor);
                        //materia
                        hojaDestino.getRange(22,49).setValue(materia);
                        //estado
                        hojaDestino.getRange(23,49).setValue(estado);

                        Logger.log(profesor);
                        Logger.log(materia);
                        Logger.log(estado);

                  }

            }
                 
}                                                  









}






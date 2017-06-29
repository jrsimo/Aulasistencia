/*
 * Aulasistencia es un script creado para gestionar las faltas de asistencia de
 * de los alumnos en el centro educativo. La versión 1.0 del script está creado
 * adhoc a las necesidades de un centro educativo en concreto (Aula Campus
 * Centro de Estudios, Burjassot, Valencia), pero este script se puede adaptar
 * a las necesidades escolares de cualquier centro. Para cualquier duda se puede
 * poner en contacto con los autores del script a través de correo electrónico.
 *
 * Copyright (C) 2016 Luís Dorado <dorado1984@gmail.com>,
 *                    José Ramón Simó <jramon.simo@gmail.com>
 *
 * This program is free software: you can redistribute it and/or modify it
 * under the terms of the GNU General Public License version 3, as published
 * by the Free Software Foundation.
 *
 * This program is distributed in the hope that it will be useful, but
 * WITHOUT ANY WARRANTY; without even the implied warranties of
 * MERCHANTABILITY, SATISFACTORY QUALITY, or FITNESS FOR A PARTICULAR
 * PURPOSE.  See the GNU General Public License for more details.
 *
 * You should have received a copy of the GNU General Public License along
 * with this program.  If not, see <http://www.gnu.org/licenses/>.
 *
 */

/* Autores: Ramón Simó y Luis Dorado
*  Tareas pendientes Ramón:
*        - COMENTAR TODO EL CÓDIGO.
*        - (v0.6) 06/09 => ubicación a la fecha de hoy en la hoja de seguimiento
*        - (v0.7) 07/09 => se han añadido las 3 funciones para lanzar los triggers por turnos (Ej. informarFaltasTurno1(1))
*                => se ha coloreado las filas de los módulos de los alumnos para mayor visibilidad (en azul)
*                => se ha coloreado en amarillo el día actual (hoy) cuando se ejecuta la función IR A DÍA DE HOY...
*                => se añade funcionalidad en el menú para ir a una fecha concreta
*                => se modifica el menú para que tenga una sección de Administrador (mayor seguridad de no tocar lo que no se debe)
*        - (v0.8) 15/08 => solucionado problema de que no se posicionaba en la fecha correcta en "ir a fecha actual..."
*
*   Tareas pensientes Luis:
*
*  CAMBIOS REALIZADOS LUIS:
*       -Cambios realizados en crearhojaseguimiento;
*       -Cambios realizados en crearNuevaHoja (solo estilo)
*       -enviar_email_porcentajes()
*       -Envío de email con destinatarios ocultos
*       -Creación de triggers por código y menú de control de triggers
*/
//Arrays de string de días y meses
var arrayDias = new Array("Domingo","Lunes", "Martes", "Miércoles", "Jueves", "Viernes", "Sábado");
var arrayMes = new Array("Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre");

//Nombres de las hojas básicas
var nombreHojaDatos = "Datos (Solo Administrador)";
var nombreHojaBievenida = "Bienvenida";
var nombreHojaSeguimiento = "Seguimiento";

//Filas de la hoja Datos donde comienzan los diferentes grupos de datos (Fechas, Módulos, Alumnos)
var filaFechas = 3;
var filaModulos = 8;
var filaAlumnos = 19;

//Porcentajes de aviso
var porcentajeAviso1 = 0.10;
var porcentajeAviso2 = 0.15;

// Fecha en la hoja de seguimiento
var _fecha_seguimiento_hoy = 0;

//Funciones para la ubicación en la columna de la fecha actual
function getFechaSeguimiento()
{
  return _fecha_seguimiento_hoy;
}

function setFechaSeguimiento(fecha)
{
  _fecha_seguimiento_hoy = fecha;
}

function getDireccionTutor (){
  //Obtengo el libro de datos iniciales y sus valores
  var libroDatos = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(nombreHojaDatos);
  var valoresDatos = libroDatos.getDataRange().getValues();
  return valoresDatos[0][4];

}

function getNombreTutor (){
  var libroDatos = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(nombreHojaDatos);
  var valoresDatos = libroDatos.getDataRange().getValues();
  return valoresDatos[0][3];

}

//Obtiene una tabla de módulos y sus datos asociados tal y como están descritos en la hoja de Datos
function getModulos (){
  //Obtengo el libro de datos iniciales y sus valores
  var libroDatos = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(nombreHojaDatos);
  var valoresDatos = libroDatos.getDataRange().getValues();
  //Obtengo un array de módulos con sus datos asociados
  var modulos = []; // Array donde guardaremos todos módulos y sus datos asociados
  var modulo;
  for (var i=filaModulos-1; valoresDatos[i][1] != ""; i++){
    var diasConClase = [];
    //Creamos un array de días que hay clase (de 1 a 5)
    for (var j=5; j<=9;j++){
      if (valoresDatos[i][j] != "") {
        diasConClase.push(valoresDatos[i][j]);
      }
    }
    //Creamos el módulo
    modulo = new Array (valoresDatos[i][1],valoresDatos[i][2],valoresDatos[i][3], valoresDatos[i][4], diasConClase);
    //Lo añadismos al array
    modulos.push(modulo);
  }
  return modulos;
}

function getNumModulos()
{
  return getModulos().length;
}

//Devuelve un vector de nombres de alumnos
function getNombreAlumnos(){
  //Obtengo el libro de datos iniciales y sus valores
  var libroDatos = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(nombreHojaDatos);
  var valoresDatos = libroDatos.getDataRange().getValues();
  //Obtengo los nombres de los alumnos
  return libroDatos.getRange(filaAlumnos, 2, valoresDatos.length -18 ,1).getValues();
}

//Devuelve una tabla de los alumnos con sus mails asociados según el formato de la hoja de datos
function getAlumnos(){
  //Obtengo el libro de datos iniciales y sus valores
  var libroDatos = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(nombreHojaDatos);
  var valoresDatos = libroDatos.getDataRange().getValues();
  //Obtengo los nombres de los alumnos
  return libroDatos.getRange(filaAlumnos, 2, valoresDatos.length -18 ,4).getValues();
}

//Crea todas las hojas de asistencia de cada módulo según la hoja de datos
function LlenarHojasAsistencia() {

  //Obtengo el libro de datos iniciales y sus valores
  var libroDatos = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(nombreHojaDatos);
  var valoresDatos = libroDatos.getDataRange().getValues();

  //Obtengo el nombre del grupo
  var nombreGrupo = valoresDatos [0][1];

  //Obtengo las fechas de inicio y final de cada cuatrimestre (Empiezan en la fila 3)
  var fechasCuatrimestre = libroDatos.getRange(filaFechas, 3, 3, 2).getValues();

  //Obtengo los nombres de los alumnos
  var nombreAlumnos = getNombreAlumnos();

  //Obtengo un array de módulos con sus datos asociados
  var modulos = getModulos(); // Array donde guardaremos todos módulos y sus datos asociados

  ///Para cada módulo llamo a CrearNuevaHoja
  for (var i=0; i < modulos.length;i++){
    CrearNuevaHoja (nombreGrupo,modulos[i], nombreAlumnos,fechasCuatrimestre);
    //POR ALGUNA RAZÓN QUE NO ALCANZO A COMPRENDER LA VARIABLE "FECHASCUATRIMESTRE" SE CORROMPE ELLA SOLITA Y HAY QUE REINICIARLA (INVESTIGAR)
    var fechasCuatrimestre = libroDatos.getRange(3, 3, 3, 2).getValues();
  }

  //Crear hoja de seguimiento de faltas para el tutor
  CrearHojaSeguimiento(nombreGrupo, modulos, nombreAlumnos, fechasCuatrimestre);

}

//Hoja de seguimiento para seguir las faltas y las nostificaciones de cada alumno.
function CrearHojaSeguimiento(nombreGrupo, modulos, nombreAlumnos,fecha)
{
  ss = SpreadsheetApp.getActiveSpreadsheet();
  //Compruebo si existe ya la hoja
  nuevaHoja = ss.getSheetByName(nombreHojaSeguimiento);

  nuevaHoja = ComprobarExisteHoja(nuevaHoja, ss);

  //Creo una nueva hoja
  nuevaHoja = ss.insertSheet(nombreHojaSeguimiento);
  nuevaHoja.getRange(1,1).setValue("Hoja de seguimiento de faltas");
  nuevaHoja.getRange(2,1).setValue(nombreGrupo);
  nuevaHoja.getRange(3,1).setValue("Apellido, Nombre");
  nuevaHoja.getRange(1,2,2,2*modulos.length).setValue("Módulos");
  nuevaHoja.getRange(1,2,2,2*modulos.length).merge();
  nuevaHoja.getRange(1,2,3,2*modulos.length).setHorizontalAlignment("center");
  nuevaHoja.getRange(1,2,3,2*modulos.length).setVerticalAlignment("middle");

  nuevaHoja.setFrozenColumns((2*modulos.length)+1); //Inmovilizo las columnas dedicadas a los porcentajes

  //Preparo las cabeceras con los nombres de los módulos
  for(var i=0;i<modulos.length;i++)
  {
    nuevaHoja.getRange(3,2+(i*2),1,2).setValue(modulos[i][0]);
    nuevaHoja.getRange(3,2+(i*2),1,2).merge();
    nuevaHoja.setColumnWidth(2+(i*2), 45);//Anchura de la columna de porcentajes
    nuevaHoja.setColumnWidth(3+(i*2), 20);//Anchura de la columna de notificaciones
  }

  //nuevaHoja.getRange(i+4, 3).setNumberFormat("0.00%");

  //Contador con la columna donde colocar cada fecha sucesiva
  var contadorColumna=(modulos.length*2)+2;
  //Una iteración por cada trimestre
  for (var i=0; i < modulos[0][1];i++){
    //Voy recorriendo todos los días entre las fechas dadas y escribiendo la fecha siempre que haya clase
    for (var fechaContador=fecha[0][0]; fechaContador<=fecha[i][1]; fechaContador.setDate(fechaContador.getDate() + 1)) {
      var diaSemana = fechaContador.getDay();
        nuevaHoja.getRange(1, contadorColumna).setValue(arrayMes[fechaContador.getMonth()]);
        nuevaHoja.getRange(2, contadorColumna).setValue(arrayDias[diaSemana]);
        nuevaHoja.getRange(3, contadorColumna).setValue(fechaContador);
        contadorColumna++;
    }
  }

  //Incluyo todos los nombres de alumnos
  nuevaHoja.getRange(4, 1, nombreAlumnos.length, 1).setValues(nombreAlumnos);

  //Incluyo todos los porcentajes de las diferentes hojas de módulos mediante fórmulas de referencia
  for (var i=0; i < nombreAlumnos.length;i++){
      for (var j=0; j < modulos.length;j++){
        nuevaHoja.getRange(i+4, 2+(j*2)).setFormulaR1C1("="+modulos[j][0]+"!R[0]C3");
      }
  }

  //Ajusto tamaño de la columna de nombres
  nuevaHoja.autoResizeColumn(1);
}


function ComprobarExisteHoja(nuevaHoja,ss)
{
  //Si existe la hoja
  if (nuevaHoja != null){
    //Si no existe una hoja de backup la renombramos
    if (ss.getSheetByName(nuevaHoja.getName()+"_backup") == null){
       nuevaHoja.setName(nuevaHoja.getName()+"_backup");
    }
    //Si existe preguntamos si se quiere borrar
    else {
      if(AlertaBorrado (nuevaHoja.getName())){
         ss.deleteSheet(ss.getSheetByName(nuevaHoja.getName()+"_backup"));
         nuevaHoja.setName(nuevaHoja.getName()+"_backup");
      }
      else return null;
    }
  }

  return nuevaHoja;
}

function CrearNuevaHoja (nombreGrupo,modulo, nombreAlumnos, fechas){
  ss = SpreadsheetApp.getActiveSpreadsheet();
  //Compruebo si existe ya la hoja
  nuevaHoja = ss.getSheetByName(modulo[0]);

  nuevaHoja = ComprobarExisteHoja(nuevaHoja, ss);

  //Creo una nueva hoja
  nuevaHoja = ss.insertSheet(modulo[0]);
  nuevaHoja.setFrozenColumns(3); //Inmovilizo las tres primeras columnas
  nuevaHoja.getRange(1,1).setValue(nombreGrupo);
  nuevaHoja.getRange(1,2).setValue(modulo[0]);
  nuevaHoja.getRange(2,1).setValue("Nº Horas");
  nuevaHoja.getRange(2,2).setValue(modulo[2]);
  nuevaHoja.getRange(1,3).setValue(modulo[3]);
  nuevaHoja.getRange(3,1).setValue("Apellidos,Nombre");
  nuevaHoja.getRange(3,2).setValue("Nº Faltas");
  nuevaHoja.getRange(3,3).setValue("Porcentaje");

  //Contador con la columna donde colocar cada fecha sucesiva
  var contadorColumna=4;
  //Una iteración por cada trimestre
  for (var i=0; i < modulo[1];i++){
    //Voy recorriendo todos los días entre las fechas dadas y escribiendo la fecha siempre que haya clase
    for (var fechaContador=fechas[i][0]; fechaContador<=fechas[i][1]; fechaContador.setDate(fechaContador.getDate() + 1)) {
      var diaSemana = fechaContador.getDay();
      //Si el día de la semana está en el array de días en los que hay clase escribo la fecha, el día de la semana y el mes
      if(EstaEnArray(diaSemana, modulo[4])){
        nuevaHoja.getRange(1, contadorColumna).setValue(arrayMes[fechaContador.getMonth()]);
        nuevaHoja.getRange(2, contadorColumna).setValue(arrayDias[diaSemana]);
        nuevaHoja.getRange(3, contadorColumna).setValue(fechaContador);
        contadorColumna++;
      }
    }
  }

  //Incluyo todos los nombres de alumnos
  nuevaHoja.getRange(4, 1, nombreAlumnos.length, 1).setValues(nombreAlumnos);

  //Pinta las filas de los alumnos para diferenciarlos mejor
  for (var i=0; i < nombreAlumnos.length;i=i+2){
    nuevaHoja.getRange(i+4,1,1,contadorColumna).setBackground("#88aaff");
  }

  //Incluyo las fórmulas del número de faltas y procentaje de faltas
  var numColumnas = nuevaHoja.getDataRange().getLastColumn(); //Última columna
  for (var i=0; i < nombreAlumnos.length;i++){
    //Fórmula para hallar el total de faltas
    nuevaHoja.getRange(i+4, 2).setFormulaR1C1("=SUM(R[0]C[2]:R[0]C["+(numColumnas-2)+"])");
    nuevaHoja.getRange(i+4, 3).setFormulaR1C1("=R[0]C[-1]/R["+(-i-2)+"]C[-1]");
    nuevaHoja.getRange(i+4, 3).setNumberFormat("0.00%");
  }

  //AÑADIR AÑADIR AÑADIR AÑADIR AÑADIR AÑADIR AÑADIR AÑADIR AÑADIR
  //Ajusto tamaño de la columna de las tres primeras columnas AÑADIR
  for (var i=0;i<3;i++){
    nuevaHoja.autoResizeColumn(i+1);
  }
}

//Compruebo si una fecha en particular tiene docencia del módulo específico
function EstaEnArray (numDiaSemana, arrayDiaSemanaDeClase){
  for (var i=0; i < arrayDiaSemanaDeClase.length;i++){
     if (numDiaSemana == arrayDiaSemanaDeClase[i]) return true;
  }
   return false;
}

function AlertaBorrado(nombreModulo) {
  var ui = SpreadsheetApp.getUi(); // Same variations.

  var result = ui.alert(
     "Confirmación de borrado de copia de seguridad",
     "Se va a borrar hoja de copia de seguridad "+nombreModulo+"_backup.",
      ui.ButtonSet.YES_NO);

  // Process the user's response.
  if (result == ui.Button.YES) {
    // User clicked "Yes".
    return true;
  } else {
    // User clicked "No" or X in the title bar.
    return false;
  }
}

//Al abrir el documento creamos el menú de Aula Campus, activamos la hoja de bienvenida y ocultamos la hoja de datos
function onOpen() {
  // Creamos el menú de acceso
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Menú AulaCampus')
      .addItem('Ir a la fecha actual', 'irADiaHoy')
      .addItem('Ir a fecha concreta...', 'irAFechaConcreta')
      .addSeparator()
      .addSubMenu(ui.createMenu('Administrador')
          .addItem('Ver hoja de datos', 'menuItem2')
          .addItem('Generar hojas de asistencia (PRECAUCIÓN)', 'menuItem3'))
          .addSubMenu(ui.createMenu('Triggers')
          .addItem('Consultar triggers', 'menuItem4')
          .addItem('Crear triggers para grupos de adultos', 'menuItem5')
          .addItem('Crear triggers para grupos de menores', 'menuItem6')
          .addItem('Borrar triggers', 'menuItem7'))
      .addToUi();

  //Activo la hoja de bienvenida
  SpreadsheetApp.getActiveSpreadsheet().getSheetByName(nombreHojaBievenida).showSheet();
  SpreadsheetApp.getActiveSpreadsheet().getSheetByName(nombreHojaBievenida).activate();

  //Ocultamos la hoja de datos
  SpreadsheetApp.getActiveSpreadsheet().getSheetByName(nombreHojaDatos).hideSheet();

  //Obtenemos la posición de la fecha de hoy en la hoja de Seguimiento
  setFechaSeguimiento(getDiaHoy());

}

/*
Función utilizada para posicionar el indicador en la columna correspondiente
a la fecha correspondia al día de hoy.
*/
function irADiaHoy()
{
  var hojaActiva = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var nombreHoja = hojaActiva.getName(); // getName() -> porque le tengo que pasar un String
  var diaEncontrado = getDiaHoy(nombreHoja);

  // Si no existe clase hoy, terminar
  if(diaEncontrado > 0)
     pintarFechaDeHoja(diaEncontrado);

}

function irAFechaConcreta()
{
  var fechaSeleccionada = Browser.inputBox('Elección del día','¿De que día quieres recuperar los registros de faltas? (mm/dd/aaaa)', Browser.Buttons.OK_CANCEL);

  if (fecha!='cancel')
  {
    var fecha = new Date(fechaSeleccionada);
    fecha.setHours(0);
    fecha.setMinutes(0);
    fecha.setSeconds(0);

    var diaEncontrado = buscarFecha(fecha);

    pintarFechaDeHoja(diaEncontrado);
  }
}

function pintarFechaDeHoja(diaEncontrado)
{
  var hojaActiva = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var nombreHoja = hojaActiva.getName(); // getName() -> porque le tengo que pasar un String

  var ubicacionColumna = 0;
  if(nombreHoja == nombreHojaSeguimiento)
    ubicacionColumna = diaEncontrado+(getNumModulos()*2)+2;
  else
    ubicacionColumna = diaEncontrado+4;

  // Situar el indicador de en la columna correspondiente a la fecha actual
  var columnaDiaHoy = hojaActiva.getRange(3,ubicacionColumna,1,1);
  hojaActiva.setActiveRange(columnaDiaHoy);

  // Limpiar el color destacado del día anterior o de la fecha seleccionada
  var rangoFechas = hojaActiva.getDataRange().getValues();
  var fechasMax = rangoFechas[2].length;
  hojaActiva.getRange(3,4,1,fechasMax).setBackground("white");

  // Poner color amarillo al día actual
  columnaDiaHoy.setBackground("yellow");
}

/*
  Devuelve la posición de la columna de la fecha actual (0,1,2...).
  Si no encuentra el día devuelve 0.
*/
function getDiaHoy(nombreHoja)
{
  var diaEncontrado=0;

  var hojaActiva = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(nombreHoja);

  var numColumnasHoja = hojaActiva.getMaxColumns();

  var ubicacionColumna = 0;
  if(nombreHoja == nombreHojaSeguimiento)
    ubicacionColumna = (getNumModulos()*2)+2;
  else
    ubicacionColumna = 4;

  var fechas = hojaActiva.getRange(3,ubicacionColumna,1,numColumnasHoja-3); //Obtengo el Rango donde están los valores de la fechas
  var valoresFechas = fechas.getValues(); // Obtengo los valores de esa fechas
  var maxDias = valoresFechas[0].length;

  // Obtengo la fecha actual
  var fechaHoy = new Date();
  fechaHoy.setDate(fechaHoy.getDate());

  // Pongo la hora a cero para poder comparar
  fechaHoy.setHours(0);
  fechaHoy.setMinutes(0);
  fechaHoy.setSeconds(0);
  fechaHoy.setMilliseconds(0);

  for (var i=0;i<maxDias;i++)
  {
    // Convierto la fecha correspondiente de la hoja a un tipo de objeto Date para poder comparar con la fecha de hoy.
    var fechaHoja = new Date(valoresFechas[0][i]);

    fechaHoja.setHours(0);
    fechaHoja.setMinutes(0);
    fechaHoja.setSeconds(0);
    fechaHoja.setMilliseconds(0);

    // RAMON: Las fechas se comparan con los simbolos ">,<,<=, >=" PERO NO CON "==" porque compararía las referencias de los objetos
    // Por eso no puedo poner el break en el if ya que cuando sea menor corta a la primera
    if (fechaHoja.getTime() == fechaHoy.getTime())
    {
      diaEncontrado=i;
      break; // Obtengo la primera fecha encontrada y termina
    }

  };

  return diaEncontrado;
}

function buscarFecha(fecha)
{
  Logger.log(fecha);
  var diaEncontrado=0;

  var hojaActiva = SpreadsheetApp.getActiveSheet();

  var numColumnasHoja = hojaActiva.getMaxColumns();

  var ubicacionColumna = 0;
  if(hojaActiva.getName() == nombreHojaSeguimiento)
    ubicacionColumna = (getNumModulos()*2)+2;
  else
    ubicacionColumna = 4;

  var fechas = hojaActiva.getRange(3,ubicacionColumna,1,numColumnasHoja-3); //Obtengo el Rango donde están los valores de la fechas
  var valoresFechas = fechas.getValues(); // Obtengo los valores de esa fechas
  var maxDias = valoresFechas[0].length;

  /*
  // Obtengo la fecha actual
  var fechaHoy = new Date();
  fechaHoy.setDate(fechaHoy.getDate());

  // Pongo la hora a cero para poder comparar
  fechaHoy.setHours(0);
  fechaHoy.setMinutes(0);
  fechaHoy.setSeconds(0);
  */

  // Pongo la hora a cero para poder comparar
  fecha.setHours(0);
  fecha.setMinutes(0);
  fecha.setSeconds(0);

  for (var i=0;i<maxDias;i++)
  {
    // Convierto la fecha correspondiente de la hoja a un tipo de objeto Date para poder comparar con la fecha de hoy.
    var fechaHoja = new Date(valoresFechas[0][i]);

    fechaHoja.setHours(0);
    fechaHoja.setMinutes(0);
    fechaHoja.setSeconds(0);

    // RAMON: Las fechas se comparan con los simbolos ">,<,<=, >=" PERO NO CON "==" porque compararía las referencias de los objetos
    // Por eso no puedo poner el break en el if ya que cuando sea menor corta a la primera
    if (fechaHoja <= fecha)
    {
      diaEncontrado=i;
      //break; // Obtengo la primera fecha encontrada y termina
    }
    else
    {
      break;
    }
  };

  return diaEncontrado;
}

//Esta función se encarga de enviar mails a los alumnos que han faltado y marcarlos como notificados en la hoja de seguimiento
function enviar_emails_faltas(idTurno)
{
  // Obtenemos la hoja de seguimiento para registrar las faltas del día en el grupo
  var hoja_seguimiento = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(nombreHojaSeguimiento);
  var rango_seguimiento = hoja_seguimiento.getDataRange();
  var diaHoySeguimiento = getDiaHoy(nombreHojaSeguimiento)+(getNumModulos()*2)+2;

  // Obtenemos los módulos que se imparten en este grupo
  var modulos = getModulos();
  Logger.log("MODULOS: ",modulos[0][0]);

  var horaTurno = 0;
  //idTurno = 1; //ESTO HAY QUE BORRARLO!!!!!
  switch(idTurno)
  {
    case 1:
      horaTurno = "8:30";
      break;
    case 2:
      horaTurno = "10:40";
      break;
    case 3:
      horaTurno = "12:40";
      break;
    default:
      horaTurno = 0;
  }

  // Buscamos las faltas de asistencia en todos los módulos
  for(i=0;i< modulos.length;i++)
  {
    var nombre_hoja = modulos[i][0];

    var hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(nombre_hoja);

    // Buscamos el día de hoy
    var dia_encontrado = getDiaHoy(nombre_hoja);

    // Si no existe clase hoy para ese módulo, continuar el bucle con el siguiente módulo
    if(dia_encontrado == 0)
        continue;

    // Si devuelve 0 es que no ha encontrado y termina la función, en caso contrario devuelve la posición del día en la hoja
    //if(!dia_encontrado)
    //  return;

    var alumnos = getAlumnos();

    var faltas = hoja.getRange(4, dia_encontrado+4, alumnos.length).getValues();

    for(var j=0;j<alumnos.length;j++) {

      if(faltas[j][0] > 0) {

        var nombre_alumno = alumnos[j][0];

        // Obtengo estadísticas para mostrar en el email
        var totalFaltas = hoja.getRange(4+j,2).getValue();
        var porcentaje = (hoja.getRange(4+j,3).getValue()*100).toFixed(2); //el valor del % en la celda se ha de truncar a dos decimales

        // Preprarar el cuerpo del email
        var cuerpo_mensaje ='<p>Estimado miembro de la comunidad educativa,<br><br> Les informamos que el/la estudiante '+ nombre_alumno + ' ha faltado a las ' + horaTurno + ' en el módulo ' + modulos[i][0] + '.</p>';
        cuerpo_mensaje = cuerpo_mensaje + '<p> Total de faltas acumuladas en el módulo ' + modulos[i][0] + ': ' + totalFaltas;
        cuerpo_mensaje = cuerpo_mensaje + '<p> Porcentaje de faltas en el módulo ' + modulos[i][0] + ': ' + porcentaje + "%";
        cuerpo_mensaje = cuerpo_mensaje + '<p>Atentamente,</p><p>Sistema de control de faltas de Aula Campus</p>';
        cuerpo_mensaje = cuerpo_mensaje + '<p>&nbsp;</p><p><b>Nota importante:</b> no respondan a este correo, ya que es un sistema automático y no recibiríamos la respuesta.</p>';
        cuerpo_mensaje = cuerpo_mensaje + '<br><p>Puede ponerse en contacto con nosotros en el teléfono 963642487 ¡Muchas gracias!</p>';

        Logger.log(cuerpo_mensaje);
        // Preparar el asunto
        var asunto = "Control de asistencia de alumnos";

        // Obtener los emails del alumno, padre y madre (o familiar)
        var email_familia1 = alumnos[j][2];
        var email_familia2 = alumnos[j][3];
        var hay_falta = rango_seguimiento.getCell(j+4,diaHoySeguimiento).getValue();
        if(hay_falta != "F"){
            var listaCorreos = email_familia1;
            if(email_familia1 != "" && email_familia2 != ""){
                   listaCorreos = listaCorreos + ","+email_familia2;
            }
            if (listaCorreos != ""){
              //Manda emails a cada uno de los familiares (alumno, padre y madre)
              GmailApp.sendEmail("", asunto, '', {htmlBody: cuerpo_mensaje,bcc:listaCorreos});
            }
       }
            //GmailApp.sendEmail("", asunto, '', {htmlBody: cuerpo_mensaje,bcc:alumnos[j][email_familia1]+","+alumnos[j][email_familia2]})
            // Poner Falta (F) en la hoja general de faltas
            rango_seguimiento.getCell(j+4,diaHoySeguimiento).setValue("F").setBackground("#ff0000").setFontColor("#ffffff").setHorizontalAlignment("center");
        }
      }
    }
  }

//Esta función se encarga de enviar mails a los alumnos que han alcanzado el 10 o 15% y marcarlos como notificados en la hoja de seguimiento
//También se encarga de quitar las marcas de notificación si el porcentaje desciende del 10 o 15%
function enviar_email_porcentajes() {
  // Obtenemos la hoja de seguimiento para registrar las faltas del día en el grupo
  var hoja_seguimiento = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(nombreHojaSeguimiento);
  var rango_seguimiento = hoja_seguimiento.getDataRange();

  //Obtengo la dirección del tutor
  var emailTutor = getDireccionTutor();
  var nombreTutor = getNombreTutor();

  // Obtenemos los módulos que se imparten en este grupo
  var modulos = getModulos();
  // Obtenemos los alumnos
  var alumnos = getAlumnos();

  //Recorremos para cada alumno y cada módulo buscando 10% o 15% para enviar email. Lo marcamos como enviado.
  //Si está notificado pero el porcentaje ha descendido del 10 o 15 según el caso lo desmarcaremos como notificado
  for (var i=0; i< alumnos.length;i++) {
    for (var j=0;j< modulos.length;j++) {
      //Flag para saber si tenemos que enviar email y qué email tenemos que enviar
      var flagEmail = "";
      var porcentajeActual = rango_seguimiento.getCell(i+4, 2+j*2).getValue();
      var notificacionActual = rango_seguimiento.getCell(i+4, 3+j*2).getValue();
      if (porcentajeActual < porcentajeAviso1) {
        if(notificacionActual != "") {
          rango_seguimiento.getCell(i+4, 3+j*2).setValue(""); //Después de justificar faltas quito la marca de notificación
        }
      }
      else if (porcentajeActual >= porcentajeAviso1 && porcentajeActual < porcentajeAviso2){
        if (notificacionActual == "") { //Aviso del 10% y marca de notificación N1
          flagEmail = "10%";
          rango_seguimiento.getCell(i+4, 3+j*2).setValue("N1");
        }
        else if (notificacionActual == "N2") { //Después de justificar faltas quito la marca de notificación N1 y la sustituyo por N2
          rango_seguimiento.getCell(i+4, 3+j*2).setValue("N1");
        }
     }
     else if (porcentajeActual >= porcentajeAviso2){
        if (notificacionActual != "N2") { //Aviso del 15% y marca de notificación N2
          flagEmail = "15%";
          rango_seguimiento.getCell(i+4, 3+j*2).setValue("N2");
        }
     }
     if (flagEmail != ""){

        var nombre_alumno = alumnos[i][0];
        // Preprarar el cuerpo del email
        var cuerpo_mensaje ='<p>Estimado miembro de la comunidad educativa,<br><br> Les informamos que el/la estudiante '+ nombre_alumno + ' ha alcanzado el '+flagEmail+' de faltas en el módulo '+modulos[j][0]+'.</p>';
        cuerpo_mensaje = cuerpo_mensaje + '<p>Le recordamos que al alcanzar el 15% de faltas en un módulo los estudiantes pierden el derecho a evaluación continua.</p>';
        cuerpo_mensaje = cuerpo_mensaje + '<p>Atentamente,</p><p>Sistema de control de faltas de Aula Campus</p>';
        cuerpo_mensaje = cuerpo_mensaje + '<p>&nbsp;</p><p><b>Nota importante:</b> no respondan a este correo, ya que es un sistema automático y no recibiríamos la respuesta.</p>';
       cuerpo_mensaje = cuerpo_mensaje + '<br><p>Puede ponerse en contacto con el tutor '+nombreTutor+' a través del email: '+emailTutor+'</p>';

        // Preparar el asunto
        var asunto = "Control de asistencia de alumnos";

        //Recupero los correos de padre, madre y tutor
        var listaCorreos = emailTutor;

        for(var k=1;k<=3;k++){
          if(alumnos[i][k] != ""){
                listaCorreos = listaCorreos + ","+alumnos [i][k];
          }
        }
       if (listaCorreos != ""){
         //Manda emails a cada uno de los familiares (alumno, padre y madre)
         GmailApp.sendEmail("", asunto, '', {htmlBody: cuerpo_mensaje,bcc:listaCorreos});
       }
     }
    }
  }
}

/*
//Voy al día actual
function menuItem1() {
  //SpreadsheetApp.getActiveSpreadsheet().getSheetByName(nombreHojaDatos).activate();
   //SpreadsheetApp.getUi() // Or DocumentApp or FormApp.
   //  .alert('Ir a día actual');
  irADiaHoy();
}
*/

//Muestro la hoja de datos
function menuItem2() {
  SpreadsheetApp.getActiveSpreadsheet().getSheetByName(nombreHojaDatos).activate();
}

//Genero las hojas de asistencia
function menuItem3() {
  var ui = SpreadsheetApp.getUi();

  var result = ui.alert(
     'Se debe contar con la autorización del administrador de faltas',
     'Esto creará nuevas hojas de asistencia vacías. ¿Desea continuar?',
      ui.ButtonSet.YES_NO);

  // Process the user's response.
  if (result == ui.Button.YES) {
    // User clicked "Yes".
    ui.alert('Las hojas serán creadas transcurridos unos cuantos segundos. Presiona Aceptar para continuar.');
    LlenarHojasAsistencia();
  } else {
    // User clicked "No" or X in the title bar.
    ui.alert('Generación de hojas de faltas cancelada');
  }
}

//Consulto si hay triggers instalados
function menuItem4(){
  var ui = SpreadsheetApp.getUi();
  numTriggers =  ScriptApp.getProjectTriggers().length;
  if (numTriggers == 0){
     ui.alert('No hay disparadores instalados.');
  }
  else {
     ui.alert('Actualmente hay instalado(s) '+numTriggers+' dispador(es).');
  }

}

//Generación de triggers para grupos de adultos
function menuItem5() {
  var ui = SpreadsheetApp.getUi();
  triggers =  ScriptApp.getProjectTriggers();
  var result = ui.alert(
     '¡Solo administradores! LOS DISPARADORES DEBEN CREARSE DESDE LA CUENTA DE NOTIFICACIONES',
     'Esto borrará cualquier trigger existente y creará triggers para grupos de ADULTOS. ¿Desea continuar?',
      ui.ButtonSet.YES_NO);

  // Process the user's response.
  if (result == ui.Button.YES) {
    // User clicked "Yes".
    //Borramos los triggers
    for (var i = 0; i < triggers.length; i++) {
      ScriptApp.deleteTrigger(triggers[i]);
    }
    crearTriggerPorcentajes ()
    ui.alert('Los disparadores para grupos de adultos han sido creados. Presiona Aceptar para continuar.');

  } else {
    // User clicked "No" or X in the title bar.
    ui.alert('Generación de disparadores para grupos de adultos cancelada');
  }
}


//Generación de triggers para grupos de menores
function menuItem6() {
  var ui = SpreadsheetApp.getUi();
  triggers =  ScriptApp.getProjectTriggers();
  var result = ui.alert(
     '¡Solo administradores! LOS DISPARADORES DEBEN CREARSE DESDE LA CUENTA DE NOTIFICACIONES',
     'Esto borrará cualquier trigger existente y creará triggers para grupos de MENORES. ¿Desea continuar?',
      ui.ButtonSet.YES_NO);

  // Process the user's response.
  if (result == ui.Button.YES) {
    // User clicked "Yes".
    //Borramos los triggers
    for (var i = 0; i < triggers.length; i++) {
      ScriptApp.deleteTrigger(triggers[i]);
    }
    crearTriggersFaltas();
    crearTriggerPorcentajes ()
    ui.alert('Los disparadores para grupos de menores han sido creados. Presiona Aceptar para continuar.');

  } else {
    // User clicked "No" or X in the title bar.
    ui.alert('Generación de disparadores para grupos de menores cancelada');
  }
}

//Borrado de triggers
function menuItem7(){
  var ui = SpreadsheetApp.getUi();
  triggers =  ScriptApp.getProjectTriggers();
  if (triggers.length == 0){
     ui.alert('No hay disparadores instalados.');
  }
  else {
    for (var i = 0; i < triggers.length; i++) {
      ScriptApp.deleteTrigger(triggers[i]);
    }
    ui.alert('Todos los disparadores han sido borrados.');
  }
}

function crearTriggersFaltas(){
  // Disparador de faltas de las 9
    ScriptApp.newTrigger('informarFaltasTurno1')
    .timeBased()
    .atHour(9)
    .everyDays(1)
    .create();
    // Disparador de faltas de las 12
    ScriptApp.newTrigger('informarFaltasTurno2')
    .timeBased()
    .atHour(12)
    .everyDays(1)
    .create();
    // Disparador de faltas de las 14
    ScriptApp.newTrigger('informarFaltasTurno3')
    .timeBased()
    .atHour(14)
    .everyDays(1)
    .create();
}

function crearTriggerPorcentajes (){
  // Disparador de porcentajes de las 22
    ScriptApp.newTrigger('enviar_email_porcentajes')
    .timeBased()
    .atHour(22)
    .everyDays(1)
    .create();
}

/*
 Los triggers manuales no aceptan argumentos por eso se crean funciones aparte que ejecutaran
 la función correspondiente con el argumento
*/
function informarFaltasTurno1()
{
   enviar_emails_faltas(1);
}

function informarFaltasTurno2()
{
   enviar_emails_faltas(2);
}

function informarFaltasTurno3()
{
   enviar_emails_faltas(3);
}

# calendar-sheet
Vinculació entre full de càlcul i calendar
// onOpen se ejecuta cada vez que se abre el documento (spreadsheet)
function onOpen() {
  // Se obtiene el user interface
  var ui = SpreadsheetApp.getUi();
  // Se agregan las dos opciones del menú
  ui.createMenu('Menu Calendario')
      .addItem('Cargar Datos', 'setEventsFromCalendar')
      .addSeparator()
      .addItem('Limpiar Semana', 'clearWeek')
      .addToUi();
}

// La función se ejecuta con una de las opciones del menú
function setEventsFromCalendar() {
  
  // Se obtiene la hoja
  var sheet = SpreadsheetApp.getActiveSheet();
  // Se obtiene el id del calendario que vive en la celda B2
  var idCalendar = sheet.getRange(1, 2).getValue();
  // Se obtiene el calendario
  var calendar = CalendarApp.getCalendarById(idCalendar);
  
  // Cuántas filas hay?
  var last = sheet.getLastRow();
  // Obtener el rango de datos, sin cabeceras
  var range = sheet.getRange(3, 1, last-2, 8);
  // Obtener los datos del rango en Object[][]
  var values = range.getValues();
  
  // Recorrer la matriz de datos
  for(var i=0; i<last-2; i++){
    // Concatenar para formar un String de título
    var title = values[i][3] + ' ' + values[i][4];
    // Un String para poder inicializar un Date, de inicio y de fin del evento
    var start = values[i][0] + ' ' + values[i][1] + ' GMT-0500';
    var end = values[i][0] + ' ' + values[i][2] + ' GMT-0500';
    // Un String de la ubicación del evento
    var locationVar = values[i][5];
    // Un String de la descripción del evento
    var descriptionVar = values[i][6];
    // Se crea el evento
    var event = calendar.createEvent(
      title,
      new Date(start),
      new Date(end),
      {location:locationVar,
       description:descriptionVar});
  }
  
}

// Método que se llama con la otra opción del menú
function clearWeek() {
  
  // Se obtiene la hoja
  var sheet = SpreadsheetApp.getActiveSheet();
  // Se obtiene el id del calendario que vive en la celda B2
  var idCalendar = sheet.getRange(1, 2).getValue();
  // Se obtiene el calendario
  var calendar = CalendarApp.getCalendarById(idCalendar);
  
  // Se obtiene el arreglo de eventos durante la semana de la cumbre
  var events = calendar.getEvents(new Date('July 28, 2014 01:00:00 GMT-0500'), new Date('July 31, 2014 23:00:00 GMT-0500'));
  // Se recorren los eventos
  for(var i=0; i<events.length; i++){
    // Se borran los eventos
    events[i].deleteEvent();
  }
}

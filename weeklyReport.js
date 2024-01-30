/**
 * Send weekly charts
 */
function sendWeeklyReport() {
  const MILLIS_PER_DAY = 1000 * 60 * 60 * 24;
  const now = new Date();
  const from = new Date(now.getTime() - 6 * MILLIS_PER_DAY);
  const timeZone = SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone();

  MailApp.sendEmail({
    to: SpreadsheetApp.getActiveSpreadsheet().getOwner().getEmail(),
    subject: "Nest Weekly Charts: "+Utilities.formatDate(from, timeZone, 'MM-dd-yyyy')+" to "+Utilities.formatDate(now, timeZone, 'MM-dd-yyyy'),
    htmlBody: "<img src='cid:pastWeekTemp'><br><img src='cid:pastWeekHumidity'>",
    inlineImages:
      {
        pastWeekTemp: generateChart('Chart Data',1,7,3,5,'Past Week Temp Log','Temperature (F)').getBlob().setContentType('image/png'),
        pastWeekHumidity: generateChart('Chart Data',1,8,4,6,'Past Week Humidity Log','Humidity').getBlob().setContentType('image/png')
      }
  });
}

/**
 * Test rendering of charts and chart options
 */
function testChartRender() {

  var chart = generateChart('Chart Data',1,7,3,5,'Past Week Temp Log','Temperature (F)');

  var htmlOutput = HtmlService.createHtmlOutput().setTitle('My Chart');
  var imageData = Utilities.base64Encode(chart.getAs('image/png').getBytes());
  var imageUrl = "data:image/png;base64," + encodeURI(imageData);
  htmlOutput.append("Render chart server side: <br/>");
  htmlOutput.append("<img border=\"1\" src=\"" + imageUrl + "\">");
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Test Chart Render');
}

/**
 * Creates line chart from sheet for Google Nest Data
 *
 * @param {string} sheetName Name of sheet with data
 * @param {number} dateCol Date column number
 * @param {number} weatherCol Weather column number
 * @param {number} downstairsCol Downstairs column number
 * @param {number} upstairsCol Upstairs column number
 * @param {string} chartTitle Chart title
 * @param {string} yTitle Y axis chart title
 * @return {Object} Chart
 */
function generateChart(sheetName,dateCol,weatherCol,downstairsCol,upstairsCol,chartTitle,yTitle) {
  // Grab all values from sheet
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  var sheetData = sheet.getRange(1,1,sheet.getLastRow(),sheet.getLastColumn()).getValues();

  // Filter last week data
  var filteredData = sheetData.filter(function(row) {
    const today = new Date();
    var d = new Date(row[0]);
    // Calculate milliseconds in a year
    const minute = 1000 * 60;
    const hour = minute * 60;
    const day = hour * 24;
    const week = day * 7;
    return d.getTime() > today.getTime() - week
  });

  // Create data table
  var data = Charts.newDataTable();
  data.addColumn(Charts.ColumnType.DATE, sheetData[0][dateCol-1]);
  data.addColumn(Charts.ColumnType.NUMBER, sheetData[0][weatherCol-1]);
  data.addColumn(Charts.ColumnType.NUMBER, sheetData[0][downstairsCol-1]);
  data.addColumn(Charts.ColumnType.NUMBER, sheetData[0][upstairsCol-1]);
  for (var i = 0;i<filteredData.length;i++) {
    data.addRow([filteredData[i][dateCol-1],filteredData[i][weatherCol-1],filteredData[i][downstairsCol-1],filteredData[i][upstairsCol-1]])
  };

  // Create line chart
  var chart = Charts.newLineChart()
      .setDataTable(data)
      .setTitle(chartTitle)
      .setXAxisTitle('Timestamp')
      .setYAxisTitle(yTitle)
      .setDimensions(640, 480)
      .build();
  
  return chart
}

/**
 * This function processes the imported data.
 * 
 * @param {Event Object} e : The event object from an installed onChange trigger.
 */
function onChange(e)
{
  try
  {
    processImportedData(e)
  }
  catch (error)
  {
    Logger.log(error['stack'])
    Browser.msgBox(error['stack'])
  }
}

/**
 * This function processes the imported data.
 */
function onOpen()
{
  SpreadsheetApp.getUi().createMenu('Update')
    .addItem('Update Quantity or Amount Data', 'collectAllHistoricalData')
    .addToUi();
}

/**
 * This function handles all of the edit events that happen on the spreadsheet, looking out for when the user is trying to use either of the search pages.
 * 
 * @param {Event Object} e : The event object from an installed onChange trigger.
 */
function installedOnEdit(e)
{
  const range = e.range;
  const col = range.columnStart;
  const row = range.rowStart;
  const rowEnd = range.rowEnd;
  const isSingleRow = row == rowEnd;
  const isSingleColumn = col == range.columnEnd;
  const spreadsheet = e.source;
  const sheet = spreadsheet.getActiveSheet();
  const sheetName = sheet.getSheetName();

  if (sheetName === 'Search for Item Quantity or Amount ($)')
  {
    conditional: if (isSingleColumn)
    {
      if (row == 1 && col == 1 && (rowEnd == 16 || rowEnd == 1))
        searchForQuantityOrAmount(spreadsheet, sheet)
      else if (isSingleRow)
      {
        if (row == 2 && col == 5)
          sheet.getRange(2, 9).uncheck()
        else if (row == 2 && col == 9)
          sheet.getRange(2, 5).uncheck()
        else if (row == 3 && col == 5)
        {
          sheet.getRange(3,  9).uncheck()
          sheet.getRange(3, 11).uncheck()
        }
        else if (row == 3 && col == 9)
        {
          sheet.getRange(3,  5).uncheck()
          sheet.getRange(3, 11).uncheck()
        }
        else if (row == 3 && col == 11)
        {
          sheet.getRange(3, 5).uncheck()
          sheet.getRange(3, 9).uncheck()
        }
        else
          break conditional;

        searchForQuantityOrAmount(spreadsheet, sheet)
      }
    }
  }
  else if (sheetName === 'Search for Invoice #s')
    if (row == 1 && col == 1 && (rowEnd == 8 || rowEnd == 1))
      searchForInvoice(spreadsheet, sheet)
}

/**
 * This function opens the LODGE SALES and CHARTER & GUIDE SALES spreadsheets, then loops through each of their yearly data. It keeps track of each active sku and 
 * sums the full sales amount and sold quantity for that item, per each year. It also aggregates both data sets together, to produce a total amount + quantity sold.
 * Each data set is stored on a separate spreadsheet and a user can use the Search for Item Quantity or Amount ($) sheet to choose which data set they want to search 
 * through.
 * 
 * @author Jarren Ralf
 */
function collectAllHistoricalData()
{
  const spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.toast('This may take several minutes...', 'Beginning Data Collection')
  const spreadsheet_lodgeSales = SpreadsheetApp.openById('1Hku1QJFuQXMkpXNW5HA9jSx_TyhsumRvjm1Wyp5QBRk')
  const spreadsheet_charterGuideSales = SpreadsheetApp.openById('1czn_JgE9V3Ie3aIr69AQLuCBpKYhjHIyj4jMrDZK9Cc')
  const currentYear = new Date().getFullYear();
  const numYears = currentYear - 2012 + 1
  const years = new Array(numYears).fill('').map((_, y) => (currentYear - y).toString()).reverse()
  const COL = numYears + 2; // A column index to ensure the correct year is being updated when mapping through each year
  const csvData = Utilities.parseCsv(DriveApp.getFilesByName("inventory.csv").next().getBlob().getDataAsString())
  var quanityData_Lodge = [], quanityData_CharterGuide = [], quanityData_All = [], amountData_Lodge = [], amountData_CharterGuide = [], amountData_All = [];
  var sheet_lodgeSales, sheet_charterGuideSales, index, index_All, item, year_y, year_index;

  // Loop through all of the years
  years.map((year, y) => {
    year_y = COL - y;
    year_index = y + 1; // Reindex starting at 1.

    sheet_lodgeSales = spreadsheet_lodgeSales.getSheetByName(year)
    sheet_lodgeSales.getSheetValues(2, 2, sheet_lodgeSales.getLastRow() - 1, 5).map(lodge => {
      if (isNotBlank(lodge[0])) // Spaces between customers
      {
        item = csvData.find(sku => lodge[1] == sku[6])

        if (item != undefined && item[10] === 'A') // Item is found and Active
        {
          index = quanityData_Lodge.findIndex(d => d[0] === item[1]);
          index_All = quanityData_All.findIndex(d => d[0] === item[1]);

          if (year_index !== numYears) // Not the current year but past years
          {
            if (index !== -1)
            {
              quanityData_Lodge[index][year_y] += Number(lodge[3])
               amountData_Lodge[index][year_y] += Number(lodge[4])
            }
            else
            {
              quanityData_Lodge.push([item[1], 0, 0, ...new Array(numYears).fill(0), '']) // .Push returns the size of the new array. Use it.
               amountData_Lodge.push([item[1], 0, 0, ...new Array(numYears).fill(0), ''])
              quanityData_Lodge[quanityData_Lodge.length - 1][year_y] = Number(lodge[3])
               amountData_Lodge[amountData_Lodge.length  - 1][year_y] = Number(lodge[4])
            }

            if (index_All !== -1)
            {
              quanityData_All[index_All][year_y] += Number(lodge[3])
               amountData_All[index_All][year_y] += Number(lodge[4])
            }
            else
            {
              quanityData_All.push([item[1], 0, 0, ...new Array(numYears).fill(0), ''])
               amountData_All.push([item[1], 0, 0, ...new Array(numYears).fill(0), ''])
              quanityData_All[quanityData_All.length - 1][year_y] = Number(lodge[3])
               amountData_All[amountData_All.length  - 1][year_y] = Number(lodge[4])
            }
          }
          else // This is the current year
          {
            if (index !== -1)
            {
              if (isNotBlank(quanityData_Lodge[index][15]))
              {
                quanityData_Lodge[index][15] += '\n' + lodge[0]
                 amountData_Lodge[index][15] += '\n' + lodge[0]
              }
              else
              {
                quanityData_Lodge[index][15] = lodge[0]
                 amountData_Lodge[index][15] = lodge[0]
              }

              quanityData_Lodge[index][3] += Number(lodge[3])
               amountData_Lodge[index][3] += Number(lodge[4])
            }
            else
            {
              quanityData_Lodge.push([item[1], 0, 0, Number(lodge[3]), ...new Array(numYears - 1).fill(0), lodge[0]])
               amountData_Lodge.push([item[1], 0, 0, Number(lodge[4]), ...new Array(numYears - 1).fill(0), lodge[0]])
            }

            if (index_All !== -1)
            {
              if (isNotBlank(quanityData_All[index_All][15]))
              {
                quanityData_All[index_All][15] += '\n' + lodge[0]
                 amountData_All[index_All][15] += '\n' + lodge[0]
              }
              else
              {
                quanityData_All[index_All][15] = lodge[0]
                 amountData_All[index_All][15] = lodge[0]
              }

              quanityData_All[index_All][3] += Number(lodge[3])
               amountData_All[index_All][3] += Number(lodge[4])
            }
            else
            {
              quanityData_All.push([item[1], 0, 0, Number(lodge[3]), ...new Array(numYears - 1).fill(0), lodge[0]])
               amountData_All.push([item[1], 0, 0, Number(lodge[4]), ...new Array(numYears - 1).fill(0), lodge[0]])
            }
          }
        }
      }
    })

    sheet_charterGuideSales = spreadsheet_charterGuideSales.getSheetByName(year)
    sheet_charterGuideSales.getSheetValues(2, 2, sheet_charterGuideSales.getLastRow() - 1, 5).map(charterGuide => {
      if (isNotBlank(charterGuide[0])) // Spaces between customers
      {
        item = csvData.find(sku => charterGuide[1] == sku[6])

        if (item != undefined && item[10] === 'A') // Item is found and Active
        {
          index = quanityData_CharterGuide.findIndex(d => d[0] === item[1]);
          index_All = quanityData_All.findIndex(d => d[0] === item[1]);

          if (year_index !== numYears)
          {
            if (index !== -1)
            {
              quanityData_CharterGuide[index][year_y] += Number(charterGuide[3])
               amountData_CharterGuide[index][year_y] += Number(charterGuide[4])
            }
            else
            {
              quanityData_CharterGuide.push([item[1], 0, 0, ...new Array(numYears).fill(0), ''])
               amountData_CharterGuide.push([item[1], 0, 0, ...new Array(numYears).fill(0), ''])
              quanityData_CharterGuide[quanityData_CharterGuide.length - 1][year_y] = Number(charterGuide[3])
               amountData_CharterGuide[amountData_CharterGuide.length  - 1][year_y] = Number(charterGuide[4])
            }

            if (index_All !== -1)
            {
              quanityData_All[index_All][year_y] += Number(charterGuide[3])
               amountData_All[index_All][year_y] += Number(charterGuide[4])
            }
            else
            {
              quanityData_All.push([item[1], 0, 0, ...new Array(numYears).fill(0), ''])
               amountData_All.push([item[1], 0, 0, ...new Array(numYears).fill(0), ''])
              quanityData_All[quanityData_All.length - 1][year_y] = Number(charterGuide[3])
               amountData_All[amountData_All.length  - 1][year_y] = Number(charterGuide[4])
            }
          }
          else
          {
            if (index !== -1)
            {
              if (isNotBlank(quanityData_CharterGuide[index][15]))
              {
                quanityData_CharterGuide[index][15] += '\n' + charterGuide[0]
                 amountData_CharterGuide[index][15] += '\n' + charterGuide[0]
              }
              else
              {
                quanityData_CharterGuide[index][15] = charterGuide[0]
                 amountData_CharterGuide[index][15] = charterGuide[0]
              }

              quanityData_CharterGuide[index][3] += Number(charterGuide[3])
               amountData_CharterGuide[index][3] += Number(charterGuide[4])  
            }
            else
            {
              quanityData_CharterGuide.push([item[1], 0, 0, Number(charterGuide[3]), ...new Array(numYears - 1).fill(0), charterGuide[0]])
               amountData_CharterGuide.push([item[1], 0, 0, Number(charterGuide[4]), ...new Array(numYears - 1).fill(0), charterGuide[0]])
            }

            if (index_All !== -1)
            {
              if (isNotBlank(quanityData_All[index_All][15]))
              {
                quanityData_All[index_All][15] += '\n' + charterGuide[0]
                 amountData_All[index_All][15] += '\n' + charterGuide[0]
              }
              else
              {
                quanityData_All[index_All][15] = charterGuide[0]
                 amountData_All[index_All][15] = charterGuide[0]
              }

              quanityData_All[index_All][3] += Number(charterGuide[3])
               amountData_All[index_All][3] += Number(charterGuide[4])
            }
            else
            {
              quanityData_All.push([item[1], 0, 0, Number(charterGuide[3]), ...new Array(numYears - 1).fill(0), charterGuide[0]])
               amountData_All.push([item[1], 0, 0, Number(charterGuide[4]), ...new Array(numYears - 1).fill(0), charterGuide[0]])
            }
          }
        }
      }
    })
  })

  quanityData_Lodge = quanityData_Lodge.map((item, i) => {
      item[1] = Math.round((item[3] + item[4] + item[5] + item[6] + item[7] + item[8])*5/3)/10; // Average
      item[2] = Math.round((item[3] + item[6] + item[7] + item[8])*5/2)/10; // Average - Covid
      item = item.map(qty => (isQtyNotZero(qty)) ? qty : '')
      amountData_Lodge[i][1] = 
        Math.round((amountData_Lodge[i][3] + amountData_Lodge[i][4] + amountData_Lodge[i][5] + amountData_Lodge[i][6] + amountData_Lodge[i][7] + amountData_Lodge[i][8])*50/3)/100; // Average
      amountData_Lodge[i][2] =  Math.round((amountData_Lodge[i][3] + amountData_Lodge[i][6] + amountData_Lodge[i][7] + amountData_Lodge[i][8])*25)/100; // Average - Covid
      amountData_Lodge[i] = amountData_Lodge[i].map(qty => (isQtyNotZero(qty)) ? qty : '')
      return item
    })

  quanityData_CharterGuide = quanityData_CharterGuide.map((item, i) => {
      item[1] = Math.round((item[3] + item[4] + item[5] + item[6] + item[7] + item[8])*5/3)/10; // Average
      item[2] = Math.round((item[3] + item[6] + item[7] + item[8])*5/2)/10; // Average - Covid
      item = item.map(qty => (isQtyNotZero(qty)) ? qty : '')
      amountData_CharterGuide[i][1] = 
        Math.round((amountData_CharterGuide[i][3] + amountData_CharterGuide[i][4] + amountData_CharterGuide[i][5] + amountData_CharterGuide[i][6] + amountData_CharterGuide[i][7] + amountData_CharterGuide[i][8])*50/3)/100; // Average
      amountData_CharterGuide[i][2] =  Math.round((amountData_CharterGuide[i][3] + amountData_CharterGuide[i][6] + amountData_CharterGuide[i][7] + amountData_CharterGuide[i][8])*25)/100; // Average - Covid
      amountData_CharterGuide[i] = amountData_CharterGuide[i].map(qty => (isQtyNotZero(qty)) ? qty : '')
      return item
    })

  quanityData_All = quanityData_All.map((item, i) => {
      item[1] = Math.round((item[3] + item[4] + item[5] + item[6] + item[7] + item[8])*5/3)/10; // Average
      item[2] = Math.round((item[3] + item[6] + item[7] + item[8])*5/2)/10; // Average - Covid
      item = item.map(qty => (isQtyNotZero(qty)) ? qty : '')
      amountData_All[i][1] = 
        Math.round((amountData_All[i][3] + amountData_All[i][4] + amountData_All[i][5] + amountData_All[i][6] + amountData_All[i][7] + amountData_All[i][8])*50/3)/100; // Average
      amountData_All[i][2] =  Math.round((amountData_All[i][3] + amountData_All[i][6] + amountData_All[i][7] + amountData_All[i][8])*25)/100; // Average - Covid
      amountData_All[i] = amountData_All[i].map(qty => (isQtyNotZero(qty)) ? qty : '')
      return item
    })

  const header = ['Descriptions', 'AVG (6 yr)', 'AVG - CoV', ...years.reverse(), 'Customers purchased in 2023'];
  const numRows_lodgeQty = quanityData_Lodge.unshift(header)
  const numRows_lodgeAmt = amountData_Lodge.unshift(header)
  const numRows_charterGuideQty = quanityData_CharterGuide.unshift(header)
  const numRows_charterGuideAmt = amountData_CharterGuide.unshift(header)
  const numRows_AllQty = quanityData_All.unshift(header)
  const numRows_AllAmt = amountData_All.unshift(header)

  spreadsheet.getSheetByName('Lodge Quantity Data').clear().getRange(1, 1, numRows_lodgeQty, quanityData_Lodge[0].length).setValues(quanityData_Lodge)
  spreadsheet.getSheetByName('Lodge Amount Data').clear().getRange(1, 1, numRows_lodgeAmt, amountData_Lodge[0].length).setValues(amountData_Lodge)
  spreadsheet.getSheetByName('Charter & Guide Quantity Data').clear().getRange(1, 1, numRows_charterGuideQty, quanityData_CharterGuide[0].length).setValues(quanityData_CharterGuide)
  spreadsheet.getSheetByName('Charter & Guide Amount Data').clear().getRange(1, 1, numRows_charterGuideAmt, amountData_CharterGuide[0].length).setValues(amountData_CharterGuide)
  spreadsheet.getSheetByName('Quantity Data').clear().getRange(1, 1, numRows_AllQty, quanityData_All[0].length).setValues(quanityData_All)
  spreadsheet.getSheetByName('Amount Data').clear().getRange(1, 1, numRows_AllAmt, amountData_All[0].length).setValues(amountData_All)
  spreadsheet.toast('All Amount / Quantity data for Lodge, Charter, and Guide customers has been updated.', 'COMPLETE', 60)
}

/**
 * This function takes all of the yearly invoice data and concatenates it into one meta set of invoice data. This function can be run on its own or
 * it is Trigger via an import of invoice data.
 * 
 * @author Jarren Ralf
 */
function concatenateAllData()
{
  const spreadsheet = SpreadsheetApp.getActive()
  const currentYear = new Date().getFullYear()
  var sheet, allData = [];

  new Array(currentYear - 2012 + 1).fill('').map((_, y) => (currentYear - y).toString()).map(year => {
    sheet = spreadsheet.getSheetByName(year)
    allData.push(...sheet.getSheetValues(2, 1, sheet.getLastRow() - 1, 8).reverse())
  })

  const lastRow = allData.unshift(['Item Description', 'Customer Name', 'Date', 'Invoice #', 'Location', 'Salesperson', 'Quantity', 'Amount']);
  spreadsheet.getSheetByName('All Data').clearContents().getRange(1, 1, lastRow, 8).setValues(allData)
}

/**
 * This function configures the yearly invoice data into the format that is desired for the spreadsheet to function optimally
 * 
 * @param {Object[][]}    values    : The values of the data that were just imported into the spreadsheet
 * @param {Spreadsheet} spreadsheet : The active spreadsheet
 * @author Jarren Ralf
 */
function configureYearlyInvoiceData(values, spreadsheet)
{
  const currentYear = new Date().getFullYear();
  const customerSheet = spreadsheet.getSheetByName('Customer List');
  const accounts = customerSheet.getSheetValues(2, 1, customerSheet.getLastRow() - 1, 1).map(v => v[0].toString().trim())
  values.shift()
  values.pop() // Remove the final row which contains descriptive stats
  const preData = removeNonImformativeSKUs(values.filter(d => accounts.includes(d[8].toString().trim())).sort(sortByDateThenInvoiveNumber))
  const data = reformatData(preData)
  const year = new Array(currentYear - 2012 + 1).fill('').map((_, y) => (currentYear - y).toString()).reverse().find(p => p == data[0][2].getFullYear()) // The year that the data is representing
  const isSingleYear = data.every(date => date[2].getFullYear() == year);

  if (isSingleYear)
  {
    const numCols = 10;
    const sheets = spreadsheet.getSheets();
    const previousSheet = sheets.find(sheet => sheet.getSheetName() == year)
    var indexAdjustment = 2010

    if (previousSheet != null)
    {
      indexAdjustment--;
      spreadsheet.deleteSheet(previousSheet)
    }
    
    SpreadsheetApp.flush();
    const newSheet = spreadsheet.insertSheet(year, sheets.length - year + indexAdjustment).hideSheet().setColumnWidths(1, 2, 350).setColumnWidths(3, 7, 85).setColumnWidth(10, 150);
    SpreadsheetApp.flush();
    const lastRow = data.unshift(['Item Description', 'Customer Name', 'Date', 'Invoice #', 'Location', 'Salesperson', 'Quantity', 'Amount', 'Customer', 'Item Number']);
    newSheet.deleteColumns(11, 16)
    newSheet.setFrozenRows(1)
    newSheet.getRange(1, 1, 1, numCols).setFontSize(11).setFontWeight('bold').offset(0, 0, lastRow, numCols).setNumberFormat('@').setValues(data)

    ScriptApp.newTrigger('concatenateAllData').timeBased().after(500).create() // Concatenate all of the data
    spreadsheet.getSheetByName('Search for Invoice #s').getRange(1, 1).activate()
  }
  else
    Browser.msgBox('Incorrect Data', 'Data contains more than one year.', Browser.Buttons.OK)
}

/**
 * This function creates the chart for the total sales amount for all lodges, charters, and guides.
 * 
 * @author Jarren Ralf
 */
function createChartForSalesData()
{
  const currentYear = new Date().getFullYear();
  const numYears = currentYear - 2012 + 1
  const spreadsheet = SpreadsheetApp.getActive()
  const salesDataSheet = spreadsheet.getSheetByName('Annual Sales Data');
  const dataRng = salesDataSheet.getRange(3, 1, numYears, 2)
  const grandTotal = salesDataSheet.getSheetValues(numYears + 3, 2, 1, 1)[0][0]

  const chart = salesDataSheet.newChart()
    .asColumnChart()
    .addRange(dataRng)
    .setNumHeaders(0)
    .setXAxisTitle('Year')
    .setYAxisTitle('Sales Total')
    .setTransposeRowsAndColumns(false)
    .setMergeStrategy(Charts.ChartMergeStrategy.MERGE_COLUMNS)
    .setHiddenDimensionStrategy(Charts.ChartHiddenDimensionStrategy.IGNORE_BOTH)
    .setOption('title', 'Annual Sales Data')
    .setOption('subtitle', 'Total: $' + new Intl.NumberFormat().format(twoDecimals(grandTotal)))
    .setOption('isStacked', 'false')
    .setOption('bubble.stroke', '#000000')
    .setOption('textStyle.color', '#000000')
    .setOption('useFirstColumnAsDomain', true)
    .setOption('titleTextStyle.color', '#757575')
    .setOption('legend.textStyle.color', '#1a1a1a')
    .setOption('subtitleTextStyle.color', '#999999')
    .setOption('series', {0: {hasAnnotations: true, dataLabel: 'value'}})
    .setOption('trendlines', {0: {lineWidth: 4, type: 'linear', color: '#6aa84f'}})
    .setOption('hAxis', {textStyle: {color: '#000000'}, titleTextStyle: {color: '#000000'}})
    .setOption('annotations', {domain: {textStyle: {color: '#808080'}}, total: {textStyle : {color: '#808080'}}})
    .setOption('vAxes', {0: {textStyle: {color: '#000000'}, titleTextStyle: {color: '#000000'}, minorGridlines: {count: 2}}})
    .setPosition(1, 1, 0, 0)
    .build();

  salesDataSheet.insertChart(chart);
  spreadsheet.moveChartToObjectSheet(chart).activate().setName('ANNUAL SALES CHART')
}

/**
 * This function checks if a given value is precisely a non-blank string.
 * 
 * @param  {String}  value : A given string.
 * @return {Boolean} Returns a boolean based on whether an inputted string is not-blank or not.
 * @author Jarren Ralf
 */
function isNotBlank(value)
{
  return value !== '';
}

/**
 * This function checks if a given number is precisely a non-zero number.
 * 
 * @param  {Number}  num : A given number.
 * @return {Boolean} Returns a boolean based on whether an inputted number is not-zero or not.
 * @author Jarren Ralf
 */
function isQtyNotZero(num)
{
  return num !== 0;
}

/**
 * This function process the imported data.
 * 
 * @param {Event Object} : The event object on an spreadsheet edit.
 * @author Jarren Ralf
 */
function processImportedData(e)
{
  if (e.changeType === 'INSERT_GRID')
  {
    var spreadsheet = e.source;
    var sheets = spreadsheet.getSheets();
    var info, numRows = 0, numCols = 1, maxRow = 2, maxCol = 3, isYearlyInvoiceData = 4;

    for (var sheet = sheets.length - 1; sheet >= 0; sheet--) // Loop through all of the sheets in this spreadsheet and find the new one
    {
      if (sheets[sheet].getType() == SpreadsheetApp.SheetType.GRID) // Some sheets in this spreadsheet are OBJECT sheets because they contain full charts
      {
        info = [
          sheets[sheet].getLastRow(),
          sheets[sheet].getLastColumn(),
          sheets[sheet].getMaxRows(),
          sheets[sheet].getMaxColumns(),
          sheets[sheet].getRange(1, 7).getValue().toString().includes('Quantity Specif')
        ]
      
        // A new sheet is imported by File -> Import -> Insert new sheet(s) - The left disjunct is for a csv and the right disjunct is for an excel file
        if ((info[maxRow] - info[numRows] === 2 && info[maxCol] - info[numCols] === 2) || 
            (info[maxRow] === 1000 && info[maxCol] === 26 && info[numRows] !== 0 && info[numCols] !== 0) ||
            info[isYearlyInvoiceData]) 
        {
          spreadsheet.toast('Processing imported data...', '', 60)
          const values = sheets[sheet].getSheetValues(1, 1, info[numRows], info[numCols]); 
          var fileName = sheets[sheet].getSheetName()

          if (info[isYearlyInvoiceData])
            configureYearlyInvoiceData(values, spreadsheet)

          if (sheets[sheet].getSheetName().substring(0, 7) !== "Copy Of") // Don't delete the sheets that are duplicates
            spreadsheet.deleteSheet(sheets[sheet]) // Delete the new sheet that was created

          spreadsheet.toast('The data will be updated in less than 5 minutes.', 'Import Complete.')
          break;
        }
      }
    }

    // Try and find the file created and delete it
    var file1 = DriveApp.getFilesByName(fileName + '.xlsx')
    var file2 = DriveApp.getFilesByName("Book1.xlsx")

    if (file1.hasNext())
      file1.next().setTrashed(true)

    if (file2.hasNext())
      file2.next().setTrashed(true)
  }
}

/**
 * This function checks the invoice numbers and reformats the numbers that come from countersales so that they are all displayed in the same format. It also changes
 * the description to the standard Google description so that the items are more easily searched for.
 * 
 * @param {String[][]} preData : The preformatted data.
 * @return {String[][]} The reformatted data
 * @author Jarren Ralf
 */
function reformatData(preData)
{
  const csvData = Utilities.parseCsv(DriveApp.getFilesByName("inventory.csv").next().getBlob().getDataAsString())
  var item;

  return preData.map(itemVals => {
    item = csvData.find(val => val[6] == itemVals[9])

    if (item != null)
      return (itemVals[3].toString().length === 9 && itemVals[3].toString().charAt('I')) ?
        [item[1], itemVals[1], itemVals[2], itemVals[3].substring(1), itemVals[4], itemVals[5], itemVals[6], itemVals[7], itemVals[8], itemVals[9]] :
      (itemVals[3].toString().length === 8 && itemVals[3].toString().charAt('I')) ?
        [item[1], itemVals[1], itemVals[2], '0' + itemVals[3].substring(1), itemVals[4], itemVals[5], itemVals[6], itemVals[7], itemVals[8], itemVals[9]] : 
        [item[1], itemVals[1], itemVals[2], itemVals[3], itemVals[4], itemVals[5], itemVals[6], itemVals[7], itemVals[8], itemVals[9]]
    else
      return (itemVals[3].toString().length === 9 && itemVals[3].toString().charAt('I')) ? 
        [itemVals[9] + ' - ' + itemVals[0] + ' - - -', itemVals[1], itemVals[2], itemVals[3].substring(1), itemVals[4], itemVals[5], itemVals[6], itemVals[7], itemVals[8], itemVals[9]] : 
      (itemVals[3].toString().length === 8 && itemVals[3].toString().charAt('I')) ? 
        [itemVals[9] + ' - ' + itemVals[0] + ' - - -', itemVals[1], itemVals[2], '0' + itemVals[3].substring(1), itemVals[4], itemVals[5], itemVals[6], itemVals[7], itemVals[8], itemVals[9]] : 
        [itemVals[9] + ' - ' + itemVals[0] + ' - - -', itemVals[1], itemVals[2], itemVals[3], itemVals[4], itemVals[5], itemVals[6], itemVals[7], itemVals[8], itemVals[9]]
  })
}

/**
 * This function receives the yearly invoice data and it removes the non-imformative SKU numbers, such as the fishing tackle, freight, and marine sale SKUs.
 * 
 * @param {String[][]} data : The yearly invoice data.
 * @return {String[][]} The yearly invoice data with non-imformative SKUs filtered out.
 * @author Jarren Ralf
 */
function removeNonImformativeSKUs(data)
{
  const fishingTackleSKUs = ["80000129", "80000389", "80000549", "80000349", "80000399", "80000499", "80000799", "80000409", "80000439", "80000599", "80000199", "80000249", "80000459", "80000699", "80000739", "80000999", "80001099", "80001149", "80001249", "80001499", "80001949", "80001999", "80000039", "80000089", "80000829", "80000259", "80000589", "80000899", "80000299", "80001199", "80001599", "80000649", "80000849", "80000025", "80000169", "80000579", "80000939", "80001299", "80000139", "80000329", "80000519", "80000629", "80000769", "80000015", "80000149", "80001549", "80000049", "80000949", "80001899", "80000020", "80000079", "80000179", "80000989", "80000449", "80000429", "80000099", "80001699", "80001649", "80001799", "80001849", "80000029", "80000339", "80000749", "80001399", "80000189", "80000289", "80000689", "80000069", "80000279", "80000159", "80000859", "80000729", "80000979", "80000059", "80000229", "80000119", "80000209", "80000219", "80000319", "80000359", "80000369", "80000419", "80000529", "80000639", "80000889", "80001749", "80000789", "80000609", "80000509", "80001049", "80000539", "80000659", "80001449", "80000109", "80000489", "80000759", "80000669", "80000469", "80000379", "80000869", "80000479", "80000679", "80000239", "80000719", "80000569", "80000709", "80000309", "80000919", "80001349", "80000879", "80000929", "80000269", "80000819", "80000619", "80000839", "80000959", "7000F6000", "7000F10000", "80002999", "7000F4000", "7000F5000", "7000F7000", "7000F3000", "7000F8000", "7000F20000", "7000F30000", "7000F9000", "80000779", "80000559", '7000M10000', '7000M200000', '7000M100000', '7000M125000', '7000M15000', '7000M150000', '7000M20000', '7000M3000', '7000M30000', '7000M4000', '7000M5000', '7000M50000', '7000M6000', '7000M7000', '7000M75000', '7000M8000', '7000M9000', 'FREIGHT', 'MISCITEM', 'MISCWEB']

  return data.filter(v => !fishingTackleSKUs.includes(v[9].toString()))
}

/**
 * This function searches all of the data for the keywords chosen by the user for the purchase of discovering the invoice numbers that contain the keywords.
 * 
 * @param {Spreadsheet}  spreadsheet : The spreadsheet that is being edited
 * @param    {Sheet}        sheet    : The sheet that is being edited
 * @author Jarren Ralf 
 */
function searchForInvoice(spreadsheet, sheet)
{
  const startTime = new Date().getTime();
  const YELLOW = "#ffe599";
  const searchResultsDisplayRange = sheet.getRange(1, 6); // The range that will display the number of items found by the search
  const functionRunTimeRange = sheet.getRange(2, 6);      // The range that will display the runtimes for the search and formatting
  const itemSearchFullRange = sheet.getRange(4, 1, sheet.getMaxRows() - 3, 8); // The entire range of the Item Search page
  //const output = [];
  const invoiceNumberList = [], highlightedRows = []
  const searchesOrNot = sheet.getRange(1, 1, 2).clearFormat()                                       // Clear the formatting of the range of the search box
    .setBorder(true, true, true, true, null, null, 'black', SpreadsheetApp.BorderStyle.SOLID_THICK) // Set the border
    .setFontFamily("Arial").setFontColor("black").setFontWeight("bold").setFontSize(14)             // Set the various font parameters
    .setHorizontalAlignment("center").setVerticalAlignment("middle")                                // Set the alignment
    .setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP)                                              // Set the wrap strategy
    .merge().trimWhitespace()                                                                       // Merge and trim the whitespaces at the end of the string
    .getValue().toString().toUpperCase().split(' NOT ')                                             // Split the search string at the word 'not'

  const searches = searchesOrNot[0].split(' OR ').map(words => words.split(/\s+/)) // Split the search values up by the word 'or' and split the results of that split by whitespace

  if (isNotBlank(searches[0][0])) // If the value in the search box is NOT blank, then compute the search
  {
    spreadsheet.toast('Searching...', '', 30)

    if (searchesOrNot.length === 1) // The word 'not' WASN'T found in the string
    {
      const dataSheet = spreadsheet.getSheetByName('All Data')
      var data = dataSheet.getSheetValues(2, 1, dataSheet.getLastRow() - 1, 8);
      const numSearches = searches.length; // The number searches
      var numSearchWords;

      for (var i = 0; i < data.length; i++) // Loop through all of the descriptions from the search data
      {
        loop: for (var j = 0; j < numSearches; j++) // Loop through the number of searches
        {
          numSearchWords = searches[j].length - 1;

          for (var k = 0; k <= numSearchWords; k++) // Loop through each word in each set of searches
          {
            if (data[i][0].toString().toUpperCase().includes(searches[j][k])) // Does the i-th item description contain the k-th search word in the j-th search
            {
              if (k === numSearchWords) // The last search word was succesfully found in the ith item, and thus, this item is returned in the search
              {
                highlightedRows.push(data[i][0]) // Push description
                if (data[i][3] !== 'I' && !invoiceNumberList.includes(data[i][3])) invoiceNumberList.push(data[i][3]); // Add the invoice number to the list (if it is not already there)
                //output.push(data[i]);
                break loop;
              }
            }
            else
              break; // One of the words in the User's query was NOT contained in the ith item description, therefore move on to the next item
          }
        }
      }
    }
    else // The word 'not' was found in the search string
    {
      var dontIncludeTheseWords = searchesOrNot[1].split(/\s+/);

      const dataSheet = spreadsheet.getSheetByName('All Data')
      var data = dataSheet.getSheetValues(2, 1, dataSheet.getLastRow() - 1, 8);
      const numSearches = searches.length; // The number searches
      var numSearchWords;

      for (var i = 0; i < data.length; i++) // Loop through all of the descriptions from the search data
      {
        loop: for (var j = 0; j < numSearches; j++) // Loop through the number of searches
        {
          numSearchWords = searches[j].length - 1;

          for (var k = 0; k <= numSearchWords; k++) // Loop through each word in each set of searches
          {
            if (data[i][0].toString().toUpperCase().includes(searches[j][k])) // Does the i-th item description contain the k-th search word in the j-th search
            {
              if (k === numSearchWords) // The last search word was succesfully found in the ith item, and thus, this item is returned in the search
              {
                for (var l = 0; l < dontIncludeTheseWords.length; l++)
                {
                  if (!data[i][0].toString().toUpperCase().includes(dontIncludeTheseWords[l]))
                  {
                    if (l === dontIncludeTheseWords.length - 1)
                    {
                      highlightedRows.push(data[i][0]) // Push description
                      if (data[i][3] !== 'I' &&!invoiceNumberList.includes(data[i][3])) invoiceNumberList.push(data[i][3]); // Add the invoice number to the list (if it is not already there)
                      //output.push(data[i]);
                      break loop;
                    }
                  }
                  else
                    break;
                }
              }
            }
            else
              break; // One of the words in the User's query was NOT contained in the ith item description, therefore move on to the next item 
          }
        }
      }
    }

    if (invoiceNumberList.length !== 0)
    {
      var output = data.filter(value => invoiceNumberList.includes(value[3]))
      var numItems = output.length;
      var numFormats = new Array(numItems).fill(['@', '@', 'dd MMM yyyy', '@', '@', '@', '@', '$#,##0.00'])

      var backgrounds = output.map(description => {
        if (highlightedRows.includes(description[0]))
          return [YELLOW, YELLOW, YELLOW, YELLOW, YELLOW, YELLOW, YELLOW, YELLOW]
        else
          return ['White', 'White', 'White', 'White', 'White', 'White', 'White', 'White']
      })
    }
    else
    {
      var output = [];
      var numItems = 0;
    }

    if (numItems === 0) // No items were found
    {
      sheet.getRange('A1').activate(); // Move the user back to the seachbox
      itemSearchFullRange.clearContent().setBackground('white'); // Clear content
      const textStyle = SpreadsheetApp.newTextStyle().setBold(true).setForegroundColor('yellow').build();
      const message = SpreadsheetApp.newRichTextValue().setText("No results found.\nPlease try again.").setTextStyle(0, 16, textStyle).build();
      searchResultsDisplayRange.setRichTextValue(message);
    }
    else
    {
      sheet.getRange('A4').activate(); // Move the user to the top of the search items
      itemSearchFullRange.clearContent().setBackground('white'); // Clear content and reset the text format
      sheet.getRange(4, 1, numItems, 8).setNumberFormats(numFormats).setBackgrounds(backgrounds).setValues(output);
      (numItems !== 1) ? searchResultsDisplayRange.setValue(numItems + " results found.") : searchResultsDisplayRange.setValue("1 result found.");
    }

    spreadsheet.toast('Searching Complete.')
  }
  else
  {
    itemSearchFullRange.clearContent(); // Clear content 
    const textStyle = SpreadsheetApp.newTextStyle().setBold(true).setForegroundColor('yellow').build();
    const message = SpreadsheetApp.newRichTextValue().setText("Invalid search.\nPlease try again.").setTextStyle(0, 14, textStyle).build();
    searchResultsDisplayRange.setRichTextValue(message);
  }

  functionRunTimeRange.setValue((new Date().getTime() - startTime)/1000 + " seconds");
}

/**
 * This function searches for either the amount or quantity of product sold to a particular set of customers, 
 * based on which option the user has selected from the checkboxes on the search sheet.
 * 
 * @param {Spreadsheet}  spreadsheet : The spreadsheet that is being edited
 * @param    {Sheet}        sheet    : The sheet that is being edited
 * @author Jarren Ralf 
 */
function searchForQuantityOrAmount(spreadsheet, sheet)
{
  const startTime = new Date().getTime();
  const searchResultsDisplayRange = sheet.getRange(1, 12); // The range that will display the number of items found by the search
  const functionRunTimeRange = sheet.getRange(2, 12);      // The range that will display the runtimes for the search and formatting
  const itemSearchFullRange = sheet.getRange(6, 1, sheet.getMaxRows() - 5, 16); // The entire range of the Item Search page
  const checkboxes = sheet.getSheetValues(2, 5, 2, 7);
  const output = [];
  const searchesOrNot = sheet.getRange(1, 1, 3).clearFormat()                                       // Clear the formatting of the range of the search box
    .setBorder(true, true, true, true, null, null, 'black', SpreadsheetApp.BorderStyle.SOLID_THICK) // Set the border
    .setFontFamily("Arial").setFontColor("black").setFontWeight("bold").setFontSize(14)             // Set the various font parameters
    .setHorizontalAlignment("center").setVerticalAlignment("middle")                                // Set the alignment
    .setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP)                                              // Set the wrap strategy
    .merge().trimWhitespace()                                                                       // Merge and trim the whitespaces at the end of the string
    .getValue().toString().toUpperCase().split(' NOT ')                                             // Split the search string at the word 'not'

  const searches = searchesOrNot[0].split(' OR ').map(words => words.split(/\s+/)) // Split the search values up by the word 'or' and split the results of that split by whitespace

  if (isNotBlank(searches[0][0])) // If the value in the search box is NOT blank, then compute the search
  {
    spreadsheet.toast('Searching...')

    if (searchesOrNot.length === 1) // The word 'not' WASN'T found in the string
    {
      const dataSheet = selectDataSheet(spreadsheet, checkboxes);
      const data = dataSheet.getSheetValues(2, 1, dataSheet.getLastRow() - 1, 16);
      const numSearches = searches.length; // The number searches
      var numSearchWords;

      for (var i = 0; i < data.length; i++) // Loop through all of the descriptions from the search data
      {
        loop: for (var j = 0; j < numSearches; j++) // Loop through the number of searches
        {
          numSearchWords = searches[j].length - 1;

          for (var k = 0; k <= numSearchWords; k++) // Loop through each word in each set of searches
          {
            if (data[i][0].toString().toUpperCase().includes(searches[j][k])) // Does the i-th item description contain the k-th search word in the j-th search
            {
              if (k === numSearchWords) // The last search word was succesfully found in the ith item, and thus, this item is returned in the search
              {
                output.push(data[i]);
                break loop;
              }
            }
            else
              break; // One of the words in the User's query was NOT contained in the ith item description, therefore move on to the next item
          }
        }
      }
    }
    else // The word 'not' was found in the search string
    {
      var dontIncludeTheseWords = searchesOrNot[1].split(/\s+/);

      const dataSheet = selectDataSheet(spreadsheet, checkboxes);
      const data = dataSheet.getSheetValues(2, 1, dataSheet.getLastRow() - 1, 16);
      const numSearches = searches.length; // The number searches
      var numSearchWords;

      for (var i = 0; i < data.length; i++) // Loop through all of the descriptions from the search data
      {
        loop: for (var j = 0; j < numSearches; j++) // Loop through the number of searches
        {
          numSearchWords = searches[j].length - 1;

          for (var k = 0; k <= numSearchWords; k++) // Loop through each word in each set of searches
          {
            if (data[i][0].toString().toUpperCase().includes(searches[j][k])) // Does the i-th item description contain the k-th search word in the j-th search
            {
              if (k === numSearchWords) // The last search word was succesfully found in the ith item, and thus, this item is returned in the search
              {
                for (var l = 0; l < dontIncludeTheseWords.length; l++)
                {
                  if (!data[i][0].toString().toUpperCase().includes(dontIncludeTheseWords[l]))
                  {
                    if (l === dontIncludeTheseWords.length - 1)
                    {
                      output.push(data[i]);
                      break loop;
                    }
                  }
                  else
                    break;
                }
              }
            }
            else
              break; // One of the words in the User's query was NOT contained in the ith item description, therefore move on to the next item 
          }
        }
      }
    }

    const numItems = output.length;

    if (numItems === 0) // No items were found
    {
      sheet.getRange('A1').activate(); // Move the user back to the seachbox
      itemSearchFullRange.clearContent(); // Clear content
      const textStyle = SpreadsheetApp.newTextStyle().setBold(true).setForegroundColor('yellow').build();
      const message = SpreadsheetApp.newRichTextValue().setText("No results found.\nPlease try again.").setTextStyle(0, 16, textStyle).build();
      searchResultsDisplayRange.setRichTextValue(message);
    }
    else
    {
      var numFormats = (checkboxes[0][0]) ? new Array(numItems).fill(['@', ...new Array(14).fill('$#,##0.00'), '@']) : new Array(numItems).fill([...new Array(16).fill('@')]);
      sheet.getRange('B8').activate(); // Move the user to the top of the search items
      itemSearchFullRange.clearContent().setBackground('white'); // Clear content and reset the text format
      sheet.getRange(6, 1, numItems, output[0].length).setNumberFormats(numFormats).setValues(output);
      (numItems !== 1) ? searchResultsDisplayRange.setValue(numItems + " results found.") : searchResultsDisplayRange.setValue("1 result found.");
    }

    spreadsheet.toast('Searching Complete.')
  }
  else
  {
    itemSearchFullRange.clearContent(); // Clear content 
    const textStyle = SpreadsheetApp.newTextStyle().setBold(true).setForegroundColor('yellow').build();
    const message = SpreadsheetApp.newRichTextValue().setText("Invalid search.\nPlease try again.").setTextStyle(0, 14, textStyle).build();
    searchResultsDisplayRange.setRichTextValue(message);
  }

  functionRunTimeRange.setValue((new Date().getTime() - startTime)/1000 + " seconds");
}

/**
 * This function returns the sheet that contains the data that the user is interested in. The choice of sheet is directly based on the checkboxes selected on the 
 * item search page.
 * 
 * @param {Spreadsheet} spreadsheet : The active spreadsheet.
 * @param {Object[][]}  checkboxes  : The values of the checkboxes
 * @author Jarren Ralf 
 */
function selectDataSheet(spreadsheet, checkboxes)
{
  if (checkboxes[0][0]) // Amount
  {
    if (checkboxes[1][0]) // Lodge
      return spreadsheet.getSheetByName('Lodge Amount Data')
    else if (checkboxes[1][4]) // Charter & Guides
      return spreadsheet.getSheetByName('Charter & Guide Amount Data')
    else if (checkboxes[1][6]) // Both
      return spreadsheet.getSheetByName('Amount Data')
  }
  else if (checkboxes[0][4]) // Quantity
  {
    if (checkboxes[1][0]) // Lodge
      return spreadsheet.getSheetByName('Lodge Quantity Data')
    else if (checkboxes[1][4]) // Charter & Guides
      return spreadsheet.getSheetByName('Charter & Guide Quantity Data')
    else if (checkboxes[1][6]) // Both
      return spreadsheet.getSheetByName('Quantity Data')
  }
}

/**
 * This function sorts the invoice data of a particular year via the date first, then sub-sorts lines with the same date via their invoice number.
 */
function sortByDateThenInvoiveNumber(a, b)
{
  return (a[2] > b[2]) ? 1 : (a[2] < b[2]) ? -1 : (a[3] > b[3]) ? 1 : (a[3] < b[3]) ? -1 : 0;
}

/**
 * This function take a number and rounds it to two decimals to make it suitable as a price.
 * 
 * @param {Number} num : The given number 
 * @return A number rounded to two decimals
 */
function twoDecimals(num)
{
  return Math.round((num + Number.EPSILON) * 100) / 100
}
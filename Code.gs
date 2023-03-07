function removeDupCustomers()
{
  const sheet = SpreadsheetApp.getActiveSheet()
  const range = sheet.getRange(2, 13, sheet.getLastRow() - 1, 1)
  const customers = range.getValues().map(customer => [customer[0].split(',').filter(onlyUnique).sort().join()]);
  range.setValues(customers);
}

function onlyUnique(value, index, self) {
  return self.indexOf(value) === index;
}

function removeAccounts()
{
  const spreadsheet = SpreadsheetApp.getActive()
  const accounts = ["C2729", "C3189", "C3201", "C3203", "C3204", "C3205", "C3234", "C3236", "C3241", "C3242", "C3244", "C3249", "C3259", "C3260", "C3263", "C3266", "C3269", "C3296", "C3299", "C3302", "C3304", "C3319", "C3320", "C3322", "C3350", "C3361", "C3363", "C3373", "C3375", "C3383", "C3387", "C3395", "C3425", "C3515", "C3743", "C3796", "C3812", "C3853", "C3883", "C3926", "C3996", "C4000", "C4001", "C4004", "C4006", "C4037", "C4097", "C4108", "C4115", "C4128", "C4134", "C4201", "C4203", "C4224", "C4229", "C4264", "C4481", "DC1001", "DC1002", "DC1003", "DC1004", "DC1005", "DC1006", "DC1007", "DC1008", "DC1010", "DC1011", "DC1012", "DC1015", "DC1016", "DC1017", "DC1018", "DC1019", "DC1020", "DC1021", "DC1022", "DC1024", "DC1025", "DC1026", "DC1028", "DC1030", "DC1031", "DC1032", "DC1036", "DC1038", "DC1039", "DC1041", "DC1043", "DC1044", "DC1046", "DC1047", "DC1048", "DC1049", "DC1061", "DC1064", "DC1065", "DC1066", "DC1067", "DC1069", "DC1070", "DC1071", "DC1072", "DC1073", "DC1075", "DC1076", "DC1077", "DC1078", "DC1079", "DC1080", "DC1082", "DC1083", "DC1084", "DC1086", "DC1087", "DC1088", "DC1091", "DC1092", "DC1093", "DC1094", "DC1096", "DC1097", "DC1098", "DC1099", "DC2000", "DC2001", "DC2002", "DC2003", "DC2004", "DC2005", "DC2006", "DC2009", "DC2010", "DC2011", "DC2012", "DC2013", "DC2014", "DC2015", "DC2016", "DC2017", "DC2018", "DC2019", "DC2020", "DC2022", "DC2023", "DC2024", "DC2025", "DC2026", "DC2027", "DC2028", "DC2029", "DC2030", "DC2031", "DC2032", "DC2033", "DC2034", "DC2035", "DC2036", "DC2037", "DC2038", "DC2039", "DC2041", "DC2042", "DC2043", "DC2044", "DC2045", "DC2046", "DC2047", "DC2048", "DC2049", "DC2050", "DC2051", "DC2052", "DC2053", "DC2054", "DC2056", "DC2057", "DC2058", "DC2059", "DC2060", "DC2061", "DC2062", "DC2063", "DC2064", "DC2065", "DC2066", "DC2067", "DC2068", "DI1002", "DI1009", "DI1013", "DI1014", "DI1019", "DI1020", "DI1021", "DI1022", "DI1026", "DI1027", "DI1031", "DI1032", "DI1033", "DI1037", "DI1044", "DI1045", "DI1046", "DI1048", "DI1049", "DI1050", "DI1051", "DI1052", "DI1053", "DI1054", "DI1057", "DI1059", "DI1060", "DI1061", "DI1062", "DI1063", "DI1064", "DI1065", "DI1066", "DI1067", "DI1070", "DI1071", "DI1074", "DI1075", "DI1076", "DI1077", "DI1078", "DI1079", "DI1080", "DI1081", "DI1082", "DI1083", "DI1084", "DI1086", "DI1087", "DI1089", "DI1091", "DI1092", "DI1093", "DI1094", "DI1097", "DI1099", "DI1100", "DI1101", "DI1103", "DL1001", "DL1002", "DL1003", "DL1004", "DL1005", "DL1006", "DL1007", "DL1008", "DL1009", "DL1010", "DL1011", "DL1013", "DL1014", "DL1015", "DL1016", "DL1017", "DL1018", "DL1019", "DL1020", "DL1021", "DL1022", "DL1023", "DL1024", "DL1025", "DL1026", "DL1027", "DL1028", "DL1029", "DL1030", "DL1031", "DL1032", "DL1033", "DL1034", "DL1037", "DL1038", "DL1039", "DL1040", "DL1041", "DL1042", "DL1043", "DL1044", "DL1045", "DL1046", "DL1047", "DL1049", "DL1050", "DL1051", "DL1052", "DL1053", "DL1054", "DL1055"]

  const numberFormat = ['@','@', '@', '@', '@', '@', '#', '$#,##0.00', '@', '@'];

  for (var i = 2012; i <= 2022; i++)
  {
    var sheet = spreadsheet.getSheetByName(i)
    var values = sheet.getDataRange().getValues();
    var maxRows = values.length
    var header = values.shift()

    var data = values.filter(v => accounts.includes(v[8].toString().trim()));

    data.unshift(header)

    var numRows = data.length
    var numFormats = new Array(numRows - 1).fill(numberFormat)
    numFormats.unshift(['@','@', '@', '@', '@', '@', '@', '@', '@', '@'])

    sheet.getRange(1, 1, numRows, 10).setNumberFormats(numFormats).setValues(data)

    sheet.deleteRows(numRows + 1, maxRows - numRows)
  }
}

function fixHeaders()
{
  const spreadsheet = SpreadsheetApp.getActive()
  const header = spreadsheet.getSheetByName('2012').getSheetValues(1, 1, 1, 10);
  const sheet = spreadsheet.getActiveSheet();
  sheet.getRange(1, 1, 1, 10).setValues(header).offset(0, 0, sheet.getLastRow(), 10).setBackground('white')
  sheet.autoResizeColumns(1, 10)
}

function concatenateData()
{
  const spreadsheet = SpreadsheetApp.getActive()

  const year21 = spreadsheet.getSheetByName('2021').getDataRange().getValues();
  year21.shift()
  const year20 = spreadsheet.getSheetByName('2020').getDataRange().getValues();
  year20.shift()
  const year19 = spreadsheet.getSheetByName('2019').getDataRange().getValues();
  year19.shift()
  const year18 = spreadsheet.getSheetByName('2018').getDataRange().getValues();
  year18.shift()
  const year17 = spreadsheet.getSheetByName('2017').getDataRange().getValues();
  year17.shift()
  const year16 = spreadsheet.getSheetByName('2016').getDataRange().getValues();
  year16.shift()
  const year15 = spreadsheet.getSheetByName('2015').getDataRange().getValues();
  year15.shift()
  const year14 = spreadsheet.getSheetByName('2014').getDataRange().getValues();
  year14.shift()
  const year13 = spreadsheet.getSheetByName('2013').getDataRange().getValues();
  year13.shift()
  const year12 = spreadsheet.getSheetByName('2012').getDataRange().getValues();
  year12.shift()


  const year = spreadsheet.getSheetByName('2022').getDataRange().getValues().concat(
    year21,
    year20,
    year19, 
    year18,
    year17,
    year16,
    year15,
    year14,
    year13,
    year12)

  const numRows = year.length;

  var numFormats = new Array(numRows - 1).fill(['@','@', '@', '@', '@', '@', '@', '$#,##0.00', '@', '@'])
  numFormats.unshift(['@','@', '@', '@', '@', '@', '@', '@', '@', '@'])

  spreadsheet.getSheetByName('All Data').getRange(1, 1, numRows, year[0].length).setValues(year)
}

function getCustomerName()
{
  const customerSheet = SpreadsheetApp.getActive().getSheetByName('Customer List');
  const accounts = customerSheet.getSheetValues(3, 2, customerSheet.getLastRow() - 2, 2).map(v => [v[0].toString().trim(), v[1].toString().trim()])
  const invoiceSheet = SpreadsheetApp.getActiveSheet()
  const range = invoiceSheet.getRange(2, 8, invoiceSheet.getLastRow() - 1, 3);
  const invoiceData = range.getValues()

  for (var i = 0; i < invoiceData.length; i++)
  {
    for (var j = 0; j < accounts.length; j++)
    {
      if (accounts[j][0] == invoiceData[i][0].toString().trim())
      {
        invoiceData[i][2] = accounts[j][1]
        break;
      }
    }
  }

  range.setValues(invoiceData)
}

function getLodgeData()
{
  const customerSheet = SpreadsheetApp.getActive().getSheetByName('Customer List');
  const accounts = customerSheet.getSheetValues(3, 1, customerSheet.getLastRow() - 2, 1).map(v => v[0].toString().trim())
  const invoiceSheet = SpreadsheetApp.getActiveSheet()
  const maxRows = invoiceSheet.getLastRow()
  const invoiceData = invoiceSheet.getSheetValues(1, 1, maxRows, 10)
  const header = invoiceData.shift()
  const remainingData = invoiceData.filter(d => accounts.includes(d[7].toString().trim()))
  const lastRow = remainingData.unshift(header);

  invoiceSheet.getRange(1, 1, remainingData.length, 10).setValues(remainingData)
  invoiceSheet.deleteRows(lastRow + 1, maxRows - lastRow)
}

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
      if (row == 1 && col == 1 && (rowEnd == 13 || rowEnd == 1))
        searchForQuantityOrAmount(spreadsheet, sheet)
      else if (isSingleRow)
      {
        if (row == 2 && col == 4)
          sheet.getRange(2, 8).uncheck()
        else if (row == 2 && col == 8)
          sheet.getRange(2, 4).uncheck()
        else if (row == 3 && col == 4)
        {
          sheet.getRange(3,  8).uncheck()
          sheet.getRange(3, 10).uncheck()
        }
        else if (row == 3 && col == 8)
        {
          sheet.getRange(3,  4).uncheck()
          sheet.getRange(3, 10).uncheck()
        }
        else if (row == 3 && col == 10)
        {
          sheet.getRange(3, 4).uncheck()
          sheet.getRange(3, 8).uncheck()
        }
        else
          break conditional;

        searchForQuantityOrAmount(spreadsheet, sheet)
      }
    }
  }
  else if (sheetName === 'Search for Invoice #s')
  {
    if (row == 1 && col == 1 && (rowEnd == 8 || rowEnd == 1))
      searchForInvoice(spreadsheet, sheet)

    // conditional: if (isSingleColumn)
    // {
    //   if (row == 1 && col == 1 && (rowEnd == 13 || rowEnd == 1))
    //     search(spreadsheet, sheet)
    //   else if (isSingleRow)
    //   {
    //     if (row == 2 && col == 4)
    //       sheet.getRange(2, 8).uncheck()
    //     else if (row == 2 && col == 8)
    //       sheet.getRange(2, 4).uncheck()
    //     else if (row == 4 && col == 4)
    //     {
    //       sheet.getRange(4,  8).uncheck()
    //       sheet.getRange(4, 10).uncheck()
    //     }
    //     else if (row == 4 && col == 8)
    //     {
    //       sheet.getRange(4,  4).uncheck()
    //       sheet.getRange(4, 10).uncheck()
    //     }
    //     else if (row == 4 && col == 10)
    //     {
    //       sheet.getRange(4, 4).uncheck()
    //       sheet.getRange(4, 8).uncheck()
    //     }
    //     else
    //       break conditional;

    //     search(spreadsheet, sheet)
    //   }
    // }
  }
}

function collectFishingTackleSkus()
{
  const sheet = SpreadsheetApp.getActiveSheet();
  const maxRows = sheet.getLastRow()
  const values = sheet.getSheetValues(1, 1, maxRows, 1)
  const fishingTackleSKUs = values.filter(v => v[0].toString().toUpperCase().includes("FISHING TACKLE") || v[0].toString().toUpperCase().includes("FISHING  TACKLE")).map(u => "\"" + u[0].split(" - ", 1)[0] + "\"")
  Logger.log(fishingTackleSKUs)
}

function removeSomeSKUs()
{
  const sheet = SpreadsheetApp.getActiveSheet();
  const maxRows = sheet.getLastRow()
  const values = sheet.getSheetValues(1, 1, maxRows, 13)
  const header1 = values.shift()
  const header2 = values.shift()
  var sku;

  const fishingTackleSKUs = ["80000129", "80000389", "80000549", "80000349", "80000399", "80000499", "80000799", "80000409", "80000439", "80000599", "80000199", "80000249", "80000459", "80000699", "80000739", "80000999", "80001099", "80001149", "80001249", "80001499", "80001949", "80001999", "80000039", "80000089", "80000829", "80000259", "80000589", "80000899", "80000299", "80001199", "80001599", "80000649", "80000849", "80000025", "80000169", "80000579", "80000939", "80001299", "80000139", "80000329", "80000519", "80000629", "80000769", "80000015", "80000149", "80001549", "80000049", "80000949", "80001899", "80000020", "80000079", "80000179", "80000989", "80000449", "80000099", "80001699", "80001649", "80001799", "80001849", "80000029", "80000339", "80000749", "80001399", "80000189", "80000289", "80000689", "80000069", "80000279", "80000159", "80000859", "80000729", "80000979", "80000059", "80000229", "80000119", "80000209", "80000219", "80000319", "80000359", "80000369", "80000419", "80000529", "80000639", "80000889", "80001749", "80000789", "80000609", "80000509", "80001049", "80000539", "80000659", "80001449", "80000109", "80000489", "80000759", "80000669", "80000469", "80000379", "80000869", "80000479", "80000679", "80000239", "80000719", "80000569", "80000709", "80000309", "80000919", "80001349", "80000879", "80000929", "80000269", "80000819", "80000619", "80000839", "80000959", "7000F6000", "7000F10000", "80002999", "7000F4000", "7000F5000", "7000F7000", "7000F3000", "7000F8000", "7000F20000", "7000F30000", "7000F9000", "80000779", "80000559"]

  const activeSkus = values.filter(v => {
    sku = v[0].split(" - ", 1)[0];
    return !fishingTackleSKUs.includes(sku)
  })

  const numRows = activeSkus.unshift(header1, header2)
  sheet.clearContents().getRange(1, 1, numRows, 13).setValues(activeSkus)
  sheet.deleteRows(numRows + 1, maxRows - numRows);
}

function removeNonActiveItems()
{
  const sheet = SpreadsheetApp.getActiveSheet();
  const maxRows = sheet.getLastRow()
  const values = sheet.getSheetValues(1, 1, maxRows, 15)
  const header1 = values.shift()
  const header2 = values.shift()
  const activeSkus = values.filter(v => v[1] !== '')
  const numRows = activeSkus.unshift(header1, header2)
  sheet.clearContents().getRange(1, 1, numRows, 15).setValues(activeSkus)
  sheet.deleteRows(numRows + 1, maxRows - numRows);
}

function updateDescriptions()
{
  const csvData = Utilities.parseCsv(DriveApp.getFilesByName("inventory.csv").next().getBlob().getDataAsString());
  const sheet = SpreadsheetApp.getActiveSheet();
  const range = sheet.getRange(2, 1, sheet.getLastRow() - 1, 10);
  const values = range.getValues()

  for (var sku = 0; sku < values.length; sku++)
  {
    for (var i = 0; i < csvData.length; i++)
    {
      if (csvData[i][10] === 'A' && values[sku][9] == csvData[i][6]) // Active and Item Number matches
      {
        values[sku][0] = csvData[i][1] // Add the description
        break;
      }
    }

    if (i === csvData.length)
      values[sku][0] = values[sku][9] + ' - ' + values[sku][0] + ' - - -'
  }

  range.setValues(values)
}

function reformatYearlyData()
{
  const spreadsheet = SpreadsheetApp.getActive()
  const ssLodge = SpreadsheetApp.openByUrl("https://docs.google.com/spreadsheets/d/1o8BB1RWkxK1uo81tBjuxGc3VWArvCdhaBctQDssPDJ0/edit")
  const ssCharter = SpreadsheetApp.openByUrl("https://docs.google.com/spreadsheets/d/1kKS6yazOEtCsH-QCLClUI_6NU47wHfRb8CIs-UTZa1U/edit")
  const numYears = 11;
  const firstYear = 2012;
  var lodgeData, charterData, year;

  for (var s = 7; s < numYears; s++)
  {
    year = (firstYear + s).toString()
    lodgeData = ssLodge.getSheetByName((firstYear + s).toString()).getDataRange().getValues().filter(u => u[17] !== '')
    charterData = ssCharter.getSheetByName((firstYear + s).toString()).getDataRange().getValues().filter(v => v[17] !== '')
    spreadsheet.insertSheet(year + ' - Lodge').getRange(2, 1, lodgeData.length, lodgeData[0].length).setNumberFormat('@').setValues(lodgeData)
    spreadsheet.insertSheet(year + ' - Charter & Guide').getRange(2, 1, charterData.length, charterData[0].length).setNumberFormat('@').setValues(charterData)
  }
}

function deleteRows()
{
  const spreadsheet = SpreadsheetApp.getActive();
  const sheets = spreadsheet.getSheets();
  sheets.shift()
  sheets.shift()
  var values;

  for (var s = 0; s < sheets.length; s++)
  {
    values = sheets[s].getDataRange().getValues().filter(v =>
      v[ 2] !== '' || v[ 3] !== '' || v[ 4] !== '' || v[ 5] !== '' ||
      v[ 6] !== '' || v[ 7] !== '' || v[ 8] !== '' || v[ 9] !== '' ||
      v[10] !== '' || v[11] !== '' || v[12] !== '')

    sheets[s].clearContents().getRange(1, 1, values.length, values[0].length).setValues(values)
  }
}

function deleteEmptyColumns()
{
  const sheets = SpreadsheetApp.getActive().getSheets();
  sheets.shift()
  sheets.shift();

  for (var s = 0; s < sheets.length; s++)
  {
    sheets[s].deleteColumn(27)
    sheets[s].deleteColumns(19, 7)
    sheets[s].deleteColumns(13, 5)
    sheets[s].deleteColumns(6, 6)
    sheets[s].deleteColumns(2, 3)
  }
}

function fillInCustomerName()
{
  const sheets = SpreadsheetApp.getActive().getSheets();
  sheets.shift()
  sheets.shift();
  var rng, account, name

  for (var s = 0; s < sheets.length; s++)
  {
    rng = sheets[s].getRange(3, 1, sheets[s].getLastRow() - 2, 2)
    values = rng.getValues();

    for (var i = 0; i < values.length; i++)
    {
      if (values[i][0] !== '')
      {
        account = values[i][0];
        name = values[i][1];
      }
      else
      {
        values[i][0] = account;
        values[i][1] = name;
      }
    }

    rng.setValues(values)
  }
}

function removeCustomers()
{
  const spreadsheet = SpreadsheetApp.getActive()
  const sheets = spreadsheet.getSheets();
  sheets.shift();
  sheets.shift()
  const lodgeAccounts = sheets.shift().getDataRange().getValues().map(v => v[0]);
  const charterAccounts = sheets.shift().getDataRange().getValues().map(v => v[0]);
  var values, data, numRows, maxRows;

  for (var s = 0; s < sheets.length; s++)
  {
    if (sheets[s].getSheetName().split(' - ')[1] === 'Lodge')
    {
      values = sheets[s].getDataRange().getValues()
      maxRows = values.length
      values.shift()
      header = values.shift()
      data = values.filter(a => lodgeAccounts.includes(a[0]))
      numRows = data.unshift(header)
      sheets[s].clearContents().getRange(1, 1, numRows, 6).setValues(data)
      sheets[s].deleteRows(numRows + 1, maxRows - numRows)
    }
    else
    {
      values = sheets[s].getDataRange().getValues()
      maxRows = values.length
      values.shift()
      header = values.shift()
      data = values.filter(a => charterAccounts.includes(a[0]))
      numRows = data.unshift(header)
      sheets[s].clearContents().getRange(1, 1, numRows, 6).setValues(data)
      sheets[s].deleteRows(numRows + 1, maxRows - numRows)
    } 
  }
}

function collectYearlyData()
{
  const spreadsheet = SpreadsheetApp.getActive()
  const sheets = spreadsheet.getSheets();
  sheets.shift();
  sheets.shift()
  var data, index, maxRows, numRows, items, skuList = [];

  for (var s = 0; s < sheets.length; s++)
  {
    data = sheets[s].getDataRange().getValues()
    items = [['Item Number', 'Item Description', 'Lodge', '', 'Charter & Guide', '', 'Customers'], 
             ['', '', 'Quantity', 'Amount', 'Quantity', 'Amount', '']];
    maxRows = data.length

    if (sheets[s].getSheetName().split(' - ')[1] === 'Lodge')
    {
      for (var i = 1; i < data.length; i++)
      {
        index = skuList.indexOf(data[i][2]);

        if (index === -1)
        {
          skuList.push(data[i][2])
          items.push([data[i][2], data[i][3], data[i][4], data[i][5], 0, 0, data[i][1]])
        }
        else
        {
          items[index + 2][2] += data[i][4]
          items[index + 2][3] += data[i][5]
          items[index + 2][6] += ', ' + data[i][1]
        }
      }
    }
    else
    {
      for (var i = 1; i < data.length; i++)
      {
        index = skuList.indexOf(data[i][2]);

        if (index === -1)
        {
          skuList.push(data[i][2])
          items.push([data[i][2], data[i][3], 0, 0, data[i][4], data[i][5], data[i][1]])
        }
        else
        {
          items[index + 2][4] += data[i][4]
          items[index + 2][5] += data[i][5]
          items[index + 2][6] += ', ' + data[i][1]
        }
      }
    }

    numRows = items.length;
    sheets[s].clearContents().getRange(1, 1, numRows, items[0].length).setValues(items)
    sheets[s].deleteRows(numRows + 1, maxRows - numRows)
    skuList.length = 0;
  }
}

function removeSheets()
{
  const spreadsheet = SpreadsheetApp.getActive();
  const sheets = spreadsheet.getSheets();
  const sheetNames = sheets.map(s => s.getSheetName());

  for (var sheet = 0; sheet < sheetNames.length; sheet++)
  {
    if (sheetNames[sheet].split(' - ')[1])
      spreadsheet.deleteSheet(sheets[sheet])
  }
}

function replaceZerosWithBlanks()
{
  const sheet = SpreadsheetApp.getActiveSheet();
  const range = sheet.getDataRange();
  const values = range.getValues();

  for (var i = 2; i < values.length; i++)
  {
    for (var j = 2; j < values[0].length; j++)
    {
      if (values[i][j] === 0)
        values[i][j] = '';
    } 
  }

  range.setValues(values)
}

function fullDataCombine()
{
  const spreadsheet = SpreadsheetApp.getActive();
  const sheets = spreadsheet.getSheets();
  sheets.shift();
  sheets.shift()
  const dataSheet = sheets.shift()
  const data = dataSheet.getDataRange().getValues()
  const skus = data.map(sku => sku[0]);
  var values, index;

  for (var s = 0; s < sheets.length; s++)
  {
    values = sheets[s].getSheetValues(3, 1, sheets[s].getLastRow() - 2, 7)
    
    for (var i = 0; i < values.length; i++)
    {
      index = skus.indexOf(values[i][0])

      if (index === -1)
      {
        skus.push(values[i][0])
        data.push([values[i][0], values[i][1], 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, values[i][6]])
        data[data.length - 1][ 3 + s] = values[i][2]
        data[data.length - 1][14 + s] = values[i][3]
        data[data.length - 1][25 + s] = values[i][4]
        data[data.length - 1][36 + s] = values[i][5]
      }
      else
      {
        data[index][ 3 + s] = values[i][2]
        data[index][14 + s] = values[i][3]
        data[index][25 + s] = values[i][4]
        data[index][36 + s] = values[i][5]
        data[index][46] += ', ' + values[i][6]
      }
    }
  }

  sheets[0].setName('Data').clearContents().getRange(1, 1, data.length, data[0].length).setValues(data)
}

function combineData()
{
  const spreadsheet = SpreadsheetApp.getActive();
  const sheets = spreadsheet.getSheets();
  sheets.shift();
  sheets.shift()
  var values, additionalValues, skus, index;

  for (var s = 0; s < sheets.length; s += 2)
  {
    values = sheets[s].getSheetValues(3, 1, sheets[s].getLastRow() - 2, 7)
    values.unshift(['Item Number', 'Item Description', 'Lodge', '', 'Charter & Guide', '', 'Customers'], 
                   ['', '', 'Quantity', 'Amount', 'Quantity', 'Amount', ''])
    skus = values.map(sku => sku[0]);
    additionalValues = sheets[s + 1].getSheetValues(3, 1, sheets[s + 1].getLastRow() - 2, 7)
    
    for (var i = 0; i < additionalValues.length; i++)
    {
      index = skus.indexOf(additionalValues[i][0])

      if (index === -1)
      {
        skus.push(additionalValues[i][0])
        values.push(additionalValues[i])
      }
      else
      {
        values[index][4] = additionalValues[i][4]
        values[index][5] = additionalValues[i][5]
        values[index][6] += ', ' + additionalValues[i][6]
      }
    }

    sheets[s].setName(sheets[s].getSheetName().split(' - ')[0]).clearContents().getRange(1, 1, values.length, values[0].length).setValues(values)
    skus.length = 0;
  }
}

function createChartForSalesData()
{
  const numYears = 11;
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
 * This function take a number and rounds it to two decimals to make it suitable as a price.
 * 
 * @param {Number} num : The given number 
 * @return A number rounded to two decimals
 */
function twoDecimals(num)
{
  return Math.round((num + Number.EPSILON) * 100) / 100
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
 * This function...
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
 * This function...
 * 
 * @param {Spreadsheet}  spreadsheet : The spreadsheet that is being edited
 * @param    {Sheet}        sheet    : The sheet that is being edited
 * @author Jarren Ralf 
 */
function searchForQuantityOrAmount(spreadsheet, sheet)
{
  const startTime = new Date().getTime();
  const searchResultsDisplayRange = sheet.getRange(1, 11); // The range that will display the number of items found by the search
  const functionRunTimeRange = sheet.getRange(2, 11);      // The range that will display the runtimes for the search and formatting
  const itemSearchFullRange = sheet.getRange(6, 1, sheet.getMaxRows() - 5, 13); // The entire range of the Item Search page
  const checkboxes = sheet.getSheetValues(2, 4, 2, 7);
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
      const data = dataSheet.getSheetValues(2, 1, dataSheet.getLastRow() - 1, 13);
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
      const data = dataSheet.getSheetValues(2, 1, dataSheet.getLastRow() - 1, 13);
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
      var numFormats = (checkboxes[0][0]) ? 
        new Array(numItems).fill(['@', '$#,##0.00', '$#,##0.00', '$#,##0.00', '$#,##0.00', '$#,##0.00', '$#,##0.00', '$#,##0.00', '$#,##0.00', '$#,##0.00', '$#,##0.00', '$#,##0.00', '@']) : 
        new Array(numItems).fill(['@', '@', '@', '@', '@', '@', '@', '@', '@', '@', '@', '@', '@']);
      sheet.getRange('B8').activate(); // Move the user to the top of the search items
      itemSearchFullRange.clearContent().setBackground('white'); // Clear content and reset the text format
      sheet.getRange(6, 1, numItems, 13).setNumberFormats(numFormats).setValues(output);
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
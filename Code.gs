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
      if (row == 1 && col == 1 && (rowEnd == 17 || rowEnd == 1))
        searchForQuantityOrAmount(spreadsheet, sheet)
      else if (isSingleRow)
      {
        if (row == 2 && col == 6) // Amount ($) Data
          sheet.getRange(2, 10).uncheck()
        else if (row == 2 && col == 10) // Quantity Data
          sheet.getRange(2, 6).uncheck()
        else if (row == 3 && col == 6) // Lodge Data
        {
          sheet.getRange(3, 10).uncheck()
          sheet.getRange(3, 12).uncheck()
        }
        else if (row == 3 && col == 10) // Charter & Guide Data
        {
          sheet.getRange(3,  6).uncheck()
          sheet.getRange(3, 12).uncheck()
        }
        else if (row == 3 && col == 12) // Both Data sets
        {
          sheet.getRange(3,  6).uncheck()
          sheet.getRange(3, 10).uncheck()
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
 * This function creates a new drop-down menu and also deletes the triggers that are not in use.
 * 
 * @author Jarren Ralf
 */
function installedOnOpen()
{
  const ui = SpreadsheetApp.getUi()
  var triggerFunction;

  ui.createMenu('PNT Menu')
    .addSubMenu(ui.createMenu('📑 Display Instructions for Updating Data')
      .addItem('📉 Invoice', 'display_Invoice_Instructions') 
      .addItem('📈 Quantity or Amount', 'display_QuantityOrAmount_Instructions'))
    .addSubMenu(ui.createMenu('📊 Add New Customer')
      .addItem('🚣‍♂️ Charter or Guide', 'addNewCharterOrGuideCustomer')
      .addItem('🚢 Lodge', 'addNewLodgeCustomer'))
    .addSubMenu(ui.createMenu('🖱 Manually Update Data')
      .addItem('📉 Invoice', 'concatenateAllData')
      .addItem('📈 Quantity or Amount', 'collectAllHistoricalData'))
    .addToUi();

  // Remove all of the unnecessary triggers. When running one-time triggers, they remain attached to the project (but disabled) and the project has a quota of 20 triggers per script
  ScriptApp.getProjectTriggers().map(trigger => {
    triggerFunction = trigger.getHandlerFunction();
    if (triggerFunction != 'onChange' && triggerFunction != 'installedOnEdit' && triggerFunction != 'installedOnOpen') // Keep all of the event triggers
      ScriptApp.deleteTrigger(trigger)
  })
}

/**
 * This function adds a new customer to the customer list and to the Dashboard. It then creates a template for their data sheet and chart.
 * 
 * @param {Spreadsheet} spreadsheet : The spreadsheet that the user is addeding a customer to.
 * @author Jarren Ralf
 */
function addNewCustomer(spreadsheet)
{
  const ui = SpreadsheetApp.getUi()
  const response1 = ui.prompt('What is the customer number?')

  if (response1.getSelectedButton() === ui.Button.OK)
  {
    const response2 = ui.prompt('What is the customer name?')

    if (response2.getSelectedButton() === ui.Button.OK)
    {
      const response3 = ui.prompt('What is the abbreviated customer name?')

      if (response3.getSelectedButton() === ui.Button.OK)
      {
        const customerNumber = response1.getResponseText().toUpperCase()
          const customerName = response2.getResponseText().toUpperCase()
             const sheetName = response3.getResponseText().toUpperCase() + ' - ' + customerNumber

        if (isNotBlank(customerNumber) && isNotBlank(customerName) && isNotBlank(sheetName))
        {
          const response4 = ui.alert('You entered the following information:\nCustomer #: \t' + customerNumber + '\nCustomer Name: \t' + customerName + '\nSheet Name: \t' + sheetName + '\n\nDoes this look correct?',ui.ButtonSet.YES_NO)

          if (response4 === ui.Button.YES)
          {
            const dashboard = spreadsheet.getSheetByName('Dashboard')
            const customerSheet = spreadsheet.getSheetByName('Customer List')
            var numRows = customerSheet.getLastRow() - 2;
            const customerNumbers = customerSheet.getSheetValues(3, 1, numRows, 1).flat()

            if (customerNumbers.includes(customerNumber))
              ui.alert('Customer is already in the list.')
            else
            {
              const customerSheetList = SpreadsheetApp.getActive().getSheetByName('Customer List'); // The Lodge, Charter, & Guide spreadsheet
              const customerList = customerSheetList.getSheetValues(2, 1, customerSheetList.getLastRow() - 1, 2);
              const numCustomers = customerList.push([customerNumber, customerName])
              customerSheetList.getRange(2, 1, numCustomers, 2).setValues(customerList.sort((a, b) => (a[0] > b[0]) ? 1 : (a[0] < b[0]) ? -1 : 0))  // Sort customer list by CUST #
              
              numRows++; // Increase because the number of customers increased by one
              const range = customerSheet.appendRow([customerNumber, customerName, sheetName]).getRange(3, 1, numRows, 3) // Add customer to either the LODGE or CHARTER & GUIDE SALES ss
              const values = range.getValues().sort((a, b) => (a[1] > b[1]) ? 1 : (a[1] < b[1]) ? -1 : 0) // Sort customer alphabetically
              range.setValues(values)

              // Once list is sorted, find new customer, then identify the customer that proceeds the new one, we will use this insert the customer sheets into the correct location
              const newCustIndex = values.findIndex(custNum => custNum[0] === customerNumber)

              if (newCustIndex !== 0)
              {
                const previousCustomerNum = values[newCustIndex - 1][0]; 
                const sheetNames = spreadsheet.getSheets().map(sht => sht.getSheetName().split(' - '))

                // Figure out what the index should be for the customer data sheet
                for (var i = 4; i < sheetNames.length; i++)
                  if (sheetNames[i][1] === previousCustomerNum)
                    break;
              }
              else
                var i = 2;

              const customerDataSheet = spreadsheet.insertSheet(sheetName, (i + 2), {template: spreadsheet.getSheetByName('Template')}).showSheet()
              const id_chart = createChart_NewCustomer(customerName, sheetName, customerDataSheet, spreadsheet) // Store ID so that we can hyperlink to the chart from Dashboard
              const lastRow = dashboard.getLastRow() + 1;
              const numCols = dashboard.getLastColumn();
              const sheetLinks = dashboard.getRange(4, 1, lastRow - 4, 2).getRichTextValues()
              const formulas_CustomerTotals = dashboard.getRange(4, 4, lastRow - 4).getFormulas()
              dashboard.appendRow(['', '', customerName, ...new Array(numCols - 3).fill('')])
              const dashboardValues = dashboard.getSheetValues(4, 1, lastRow - 3, numCols).sort((a, b) => (a[2] > b[2]) ? 1 : (a[2] < b[2]) ? -1 : 0) // Sort customer alphabetically

              formulas_CustomerTotals.splice(newCustIndex, 0, 
                ['=SUM(E' + (newCustIndex + 4) + ':' + dashboard.getRange(1, numCols).getA1Notation()[0] + (newCustIndex + 4) + ')']
              )

              for (var i = newCustIndex; i < formulas_CustomerTotals.length; i++)
                formulas_CustomerTotals[i][0] = formulas_CustomerTotals[i][0].replaceAll(/\d+/g, i + 4)

              sheetLinks.splice(newCustIndex, 0, 
                [SpreadsheetApp.newRichTextValue().setText(customerNumber).setLinkUrl('#gid=' + customerDataSheet.getSheetId()).build(), 
                 SpreadsheetApp.newRichTextValue().setText(customerNumber).setLinkUrl('#gid=' + id_chart).build()]
              )

              const formulaRange_YearlyTotals = dashboard.getRange(4, 1, lastRow - 3, numCols).setValues(dashboardValues) // Set the customer names and sales values
                .offset(0, 0, lastRow - 3, 2).setRichTextValues(sheetLinks)        // Set the hyperlinked sheet links
                .offset(0, 3, lastRow - 3, 1).setFormulas(formulas_CustomerTotals) // Set the customer totals formulas
                .offset(-1, 0, 1, numCols - 3); // The range of the yearly sales totals; Their formulas need tp be updated because we have an extra row

              const formulas_YearlyTotals = [formulaRange_YearlyTotals.getFormulas()[0].map(formula => formula.toString().substring(0, 9) + lastRow.toString() + ')')]
              formulaRange_YearlyTotals.setFormulas(formulas_YearlyTotals).offset(1, -3, lastRow - 3, numCols).activate() // Set new formulas because we have a new customer and they need to extend one more row
            }
          }
        }
        else
          ui.alert('Atleast one of your typed responses was blank.\n\nPlease redo the process.')
      }
    }
  }
}

/**
 * This function adds a new Charter or Guide customer.
 * 
 * @author Jarren Ralf
 */
function addNewCharterOrGuideCustomer()
{
  addNewCustomer(SpreadsheetApp.openById('1kKS6yazOEtCsH-QCLClUI_6NU47wHfRb8CIs-UTZa1U'))
}

/**
 * This function adds a new Charter or Guide customer.
 * 
 * @author Jarren Ralf
 */
function addNewLodgeCustomer()
{
  addNewCustomer(SpreadsheetApp.openById('1o8BB1RWkxK1uo81tBjuxGc3VWArvCdhaBctQDssPDJ0'))
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
  const spreadsheet_lodgeSales = SpreadsheetApp.openById('1o8BB1RWkxK1uo81tBjuxGc3VWArvCdhaBctQDssPDJ0')
  const spreadsheet_charterGuideSales = SpreadsheetApp.openById('1kKS6yazOEtCsH-QCLClUI_6NU47wHfRb8CIs-UTZa1U')
  const currentYear = new Date().getFullYear();
  const numYears = currentYear - 2012 + 1;
  const years = new Array(numYears).fill('').map((_, y) => (currentYear - y).toString()).reverse(); // Years in ascending order
  const COL = numYears + 2; // A column index to ensure the correct year is being updated when mapping through each year
  const csvData = Utilities.parseCsv(DriveApp.getFilesByName("inventory.csv").next().getBlob().getDataAsString())
  const itemNum = csvData[0].indexOf('Item #');
  const fullDescription = csvData[0].indexOf('Item List')
  var quanityData_Lodge = [], quanityData_CharterGuide = [], quanityData_All = [], amountData_Lodge = [], amountData_CharterGuide = [], amountData_All = [];
  var sheet_lodgeSales, sheet_charterGuideSales, index, index_All, item, year_y, year_index, cust_name_index = numYears + 3;

  // Loop through all of the years
  years.map((year, y) => {
    year_y = COL - y; // The appropriate index for the y-th year
    year_index = y + 2; // Reindex to keep the last 2 years of customer names in the data

    sheet_lodgeSales = spreadsheet_lodgeSales.getSheetByName(year)

    if (sheet_lodgeSales !== null)
    {
      sheet_lodgeSales.getSheetValues(2, 2, sheet_lodgeSales.getLastRow() - 1, 5).map(lodge => { // Loop through all of the lodge data for the y-th year
        if (isNotBlank(lodge[0])) // Spaces between customers
        {
          item = csvData.find(sku => lodge[1] == sku[itemNum]) // Find the current item in the adagio csv data

          if (item != undefined) // Item is found
          {
            index = quanityData_Lodge.findIndex(d => d[0] === item[fullDescription]);   // The index for the current item in the lodge quantity data
            index_All = quanityData_All.findIndex(d => d[0] === item[fullDescription]); // The index for the current item in the combined quantity data

            if (year_index < numYears) // Not last year or the current year either but past years
            {
              if (index !== -1) // Current item is already in lodge data list
              {
                quanityData_Lodge[index][year_y] += Number(lodge[3]) // Increase the quantity
                amountData_Lodge[index][year_y] += Number(lodge[4]) // Increase the amount ($)
              }
              else // The current item is not in the lodge data yet, so add it in
              {
                quanityData_Lodge.push([item[fullDescription], 0, 0, ...new Array(numYears).fill(0), ''])
                amountData_Lodge.push([item[fullDescription], 0, 0, ...new Array(numYears).fill(0), ''])
                quanityData_Lodge[quanityData_Lodge.length - 1][year_y] = Number(lodge[3]) // Add quantity to the appropriate year (column)
                amountData_Lodge[amountData_Lodge.length  - 1][year_y] = Number(lodge[4]) // Add amount ($) to the appropriate year (column)
              }

              if (index_All !== -1) // Current item is already in combined data list
              {
                quanityData_All[index_All][year_y] += Number(lodge[3]) // Increase the quantity
                amountData_All[index_All][year_y] += Number(lodge[4]) // Increase the amount ($)
              }
              else // The current item is not in the combined data yet, so add it in
              {
                quanityData_All.push([item[fullDescription], 0, 0, ...new Array(numYears).fill(0), ''])
                amountData_All.push([item[fullDescription], 0, 0, ...new Array(numYears).fill(0), ''])
                quanityData_All[quanityData_All.length - 1][year_y] = Number(lodge[3]) // Add quantity to the appropriate year (column)
                amountData_All[amountData_All.length  - 1][year_y] = Number(lodge[4]) // Add amount ($) to the appropriate year (column)
              }
            }
            else // This is the the previous year or current year; We want to identify which customers purchased particular items in these years
            {
              if (index !== -1) // Current item is already in lodge data list
              {
                if (isNotBlank(quanityData_Lodge[index][cust_name_index])) // Another lodge customer is added to the list of lodge customers who have purchased this item in the current year
                {
                  quanityData_Lodge[index][cust_name_index] += '\n(' + year + ') ' + lodge[0]
                  amountData_Lodge[index][cust_name_index] += '\n(' + year + ') ' + lodge[0]
                }
                else // This is the first lodge customer to purchase this item in the current year
                {
                  quanityData_Lodge[index][cust_name_index] = '(' + year + ') ' + lodge[0]
                  amountData_Lodge[index][cust_name_index] = '(' + year + ') ' + lodge[0]
                }

                quanityData_Lodge[index][year_y] += Number(lodge[3]) // Increase the quantity
                amountData_Lodge[index][year_y] += Number(lodge[4]) // Increase the quantity 
              }
              else // The current item is not in the lodge data yet, so add it in
              {
                quanityData_Lodge.push([item[fullDescription], 0, 0, ...new Array(numYears).fill(0), '(' + year + ') ' + lodge[0]])
                amountData_Lodge.push([item[fullDescription], 0, 0, ...new Array(numYears).fill(0), '(' + year + ') ' + lodge[0]])
                quanityData_Lodge[quanityData_Lodge.length - 1][year_y] = Number(lodge[3]) // Add quantity to the appropriate year (column)
                amountData_Lodge[amountData_Lodge.length  - 1][year_y] = Number(lodge[4]) // Add amount ($) to the appropriate year (column)
              }

              if (index_All !== -1) // Current item is already in combined data list
              {
                if (isNotBlank(quanityData_All[index_All][cust_name_index])) // Another lodge customer is added to the list of combined customers who have purchased this item in the current year
                {
                  quanityData_All[index_All][cust_name_index] += '\n(' + year + ') ' + lodge[0]
                  amountData_All[index_All][cust_name_index] += '\n(' + year + ') ' + lodge[0]
                }
                else // This is the first lodge customer to purchase this item in the current year
                {
                  quanityData_All[index_All][cust_name_index] = '(' + year + ') ' + lodge[0]
                  amountData_All[index_All][cust_name_index] = '(' + year + ') ' + lodge[0]
                }

                quanityData_All[index_All][year_y] += Number(lodge[3]) // Increase the quantity
                amountData_All[index_All][year_y] += Number(lodge[4]) // Increase the amount ($)
              }
              else // The current item is not in the combined data yet, so add it in
              {
                quanityData_All.push([item[fullDescription], 0, 0, ...new Array(numYears).fill(0), '(' + year + ') ' + lodge[0]])
                amountData_All.push([item[fullDescription], 0, 0, ...new Array(numYears).fill(0), '(' + year + ') ' + lodge[0]])
                quanityData_All[quanityData_All.length - 1][year_y] = Number(lodge[3]) // Add quantity to the appropriate year (column)
                amountData_All[amountData_All.length  - 1][year_y] = Number(lodge[4]) // Add amount ($) to the appropriate year (column)
              }
            }
          }
        }
      })
    }

    sheet_charterGuideSales = spreadsheet_charterGuideSales.getSheetByName(year)

    if (sheet_charterGuideSales !== null)
    {
      sheet_charterGuideSales.getSheetValues(2, 2, sheet_charterGuideSales.getLastRow() - 1, 5).map(charterGuide => { // Loop through all of the charter & guide data for the y-th year
        if (isNotBlank(charterGuide[0])) // Spaces between customers
        {
          item = csvData.find(sku => charterGuide[1] == sku[itemNum]) // Find the current item in the adagio csv data

          if (item != undefined) // Item is found
          {
            index = quanityData_CharterGuide.findIndex(d => d[0] === item[fullDescription]); // The index for the current item in the charter & guide quantity data
            index_All = quanityData_All.findIndex(d => d[0] === item[fullDescription]);      // The index for the current item in the combined quantity data

            if (year_index < numYears) // Not the current year but past years
            {
              if (index !== -1) // Current item is already in charter & guide data list
              {
                quanityData_CharterGuide[index][year_y] += Number(charterGuide[3]) // Increase the quantity
                amountData_CharterGuide[index][year_y] += Number(charterGuide[4]) // Increase the amount ($)
              }
              else // The current item is not in the charter & guide data yet, so add it in
              {
                quanityData_CharterGuide.push([item[fullDescription], 0, 0, ...new Array(numYears).fill(0), ''])
                amountData_CharterGuide.push([item[fullDescription], 0, 0, ...new Array(numYears).fill(0), ''])
                quanityData_CharterGuide[quanityData_CharterGuide.length - 1][year_y] = Number(charterGuide[3]) // Add quantity to the appropriate year (column)
                amountData_CharterGuide[amountData_CharterGuide.length  - 1][year_y] = Number(charterGuide[4]) // Add amount ($) to the appropriate year (column)
              }

              if (index_All !== -1) // Current item is already in combined data list
              {
                quanityData_All[index_All][year_y] += Number(charterGuide[3]) // Increase the quantity
                amountData_All[index_All][year_y] += Number(charterGuide[4]) // Increase the amount ($)
              }
              else // The current item is not in the combined data yet, so add it in
              {
                quanityData_All.push([item[fullDescription], 0, 0, ...new Array(numYears).fill(0), ''])
                amountData_All.push([item[fullDescription], 0, 0, ...new Array(numYears).fill(0), ''])
                quanityData_All[quanityData_All.length - 1][year_y] = Number(charterGuide[3]) // Add quantity to the appropriate year (column)
                amountData_All[amountData_All.length  - 1][year_y] = Number(charterGuide[4]) // Add amount ($) to the appropriate year (column)
              }
            }
            else // This is the the previous year or current year; We want to identify which customers purchased particular items in these years
            {
              if (index !== -1) // Current item is already in charter & guide data list
              { // Another charter & guide customer is added to the list of charter & guide customers who have purchased this item in the current year
                if (isNotBlank(quanityData_CharterGuide[index][cust_name_index])) 
                {
                  quanityData_CharterGuide[index][cust_name_index] += '\n(' + year + ') ' + charterGuide[0]
                  amountData_CharterGuide[index][cust_name_index] += '\n(' + year + ') ' + charterGuide[0]
                }
                else // This is the first charter & guide customer to purchase this item in the current year
                {
                  quanityData_CharterGuide[index][cust_name_index] = '(' + year + ') ' + charterGuide[0]
                  amountData_CharterGuide[index][cust_name_index] = '(' + year + ') ' + charterGuide[0]
                }

                quanityData_CharterGuide[index][year_y] += Number(charterGuide[3]) // Increase the quantity
                amountData_CharterGuide[index][year_y] += Number(charterGuide[4]) // Increase the amount ($)
              }
              else // The current item is not in the charter & guide data yet, so add it in
              {
                quanityData_CharterGuide.push([item[fullDescription], 0, 0, ...new Array(numYears).fill(0), '(' + year + ') ' + charterGuide[0]])
                amountData_CharterGuide.push([item[fullDescription], 0, 0, ...new Array(numYears).fill(0), '(' + year + ') ' + charterGuide[0]])
                quanityData_CharterGuide[quanityData_CharterGuide.length - 1][year_y] = Number(charterGuide[3]) // Add quantity to the appropriate year (column)
                amountData_CharterGuide[amountData_CharterGuide.length  - 1][year_y] = Number(charterGuide[4]) // Add amount ($) to the appropriate year (column)
              }

              if (index_All !== -1) // Current item is already in combined data list
              {
                if (isNotBlank(quanityData_All[index_All][cust_name_index])) // Another charter & guide customer is added to the list of combined customers who have purchased this item in the current year
                {
                  quanityData_All[index_All][cust_name_index] += '\n(' + year + ') ' + charterGuide[0]
                  amountData_All[index_All][cust_name_index] += '\n(' + year + ') ' + charterGuide[0]
                }
                else // This is the first charter & guide customer to purchase this item in the current year
                {
                  quanityData_All[index_All][cust_name_index] = '(' + year + ') ' + charterGuide[0]
                  amountData_All[index_All][cust_name_index] = '(' + year + ') ' + charterGuide[0]
                }

                quanityData_All[index_All][year_y] += Number(charterGuide[3]) // Increase the quantity
                amountData_All[index_All][year_y] += Number(charterGuide[4]) // Increase the amount ($)
              }
              else // The current item is not in the combined data yet, so add it in
              {
                quanityData_All.push([item[fullDescription], 0, 0, ...new Array(numYears).fill(0), '(' + year + ') ' + charterGuide[0]])
                amountData_All.push([item[fullDescription], 0, 0, ...new Array(numYears).fill(0), '(' + year + ') ' + charterGuide[0]])
                quanityData_All[quanityData_All.length - 1][year_y] = Number(charterGuide[3]) // Add quantity to the appropriate year (column)
                amountData_All[amountData_All.length  - 1][year_y] = Number(charterGuide[4]) // Add amount ($) to the appropriate year (column)
              }
            }
          }
        }
      })
    }
  })

  quanityData_Lodge = quanityData_Lodge.map((item, i) => {
      item[1] = Math.round((item[3] + item[4] + item[5] + item[6] + item[7] + item[8])*5/3)/10; // Average
      item[2] = Math.round((item[3] + item[6] + item[7] + item[8])*5/2)/10; // Average - Covid
      item = item.map(qty => (isQtyNotZero(qty)) ? qty : '') // Remove the zeros, '0', and replace them with a blank string (makes the data present cleaner)
      amountData_Lodge[i][1] = 
        Math.round((amountData_Lodge[i][3] + amountData_Lodge[i][4] + amountData_Lodge[i][5] + amountData_Lodge[i][6] + amountData_Lodge[i][7] + amountData_Lodge[i][8])*50/3)/100; // Average
      amountData_Lodge[i][2] =  Math.round((amountData_Lodge[i][3] + amountData_Lodge[i][6] + amountData_Lodge[i][7] + amountData_Lodge[i][8])*25)/100; // Average - Covid
      amountData_Lodge[i] = amountData_Lodge[i].map(qty => (isQtyNotZero(qty)) ? qty : '') // Remove the zeros
      return item
    })

  quanityData_CharterGuide = quanityData_CharterGuide.map((item, i) => {
      item[1] = Math.round((item[3] + item[4] + item[5] + item[6] + item[7] + item[8])*5/3)/10; // Average
      item[2] = Math.round((item[3] + item[6] + item[7] + item[8])*5/2)/10; // Average - Covid
      item = item.map(qty => (isQtyNotZero(qty)) ? qty : '') // Remove the zeros, '0', and replace them with a blank string (makes the data present cleaner)
      amountData_CharterGuide[i][1] = 
        Math.round((amountData_CharterGuide[i][3] + amountData_CharterGuide[i][4] + amountData_CharterGuide[i][5] + amountData_CharterGuide[i][6] + amountData_CharterGuide[i][7] + amountData_CharterGuide[i][8])*50/3)/100; // Average
      amountData_CharterGuide[i][2] =  Math.round((amountData_CharterGuide[i][3] + amountData_CharterGuide[i][6] + amountData_CharterGuide[i][7] + amountData_CharterGuide[i][8])*25)/100; // Average - Covid
      amountData_CharterGuide[i] = amountData_CharterGuide[i].map(qty => (isQtyNotZero(qty)) ? qty : '') // Remove the zeros
      return item
    })

  quanityData_All = quanityData_All.map((item, i) => {
      item[1] = Math.round((item[3] + item[4] + item[5] + item[6] + item[7] + item[8])*5/3)/10; // Average
      item[2] = Math.round((item[3] + item[6] + item[7] + item[8])*5/2)/10; // Average - Covid
      item = item.map(qty => (isQtyNotZero(qty)) ? qty : '') // Remove the zeros, '0', and replace them with a blank string (makes the data present cleaner)
      amountData_All[i][1] = 
        Math.round((amountData_All[i][3] + amountData_All[i][4] + amountData_All[i][5] + amountData_All[i][6] + amountData_All[i][7] + amountData_All[i][8])*50/3)/100; // Average
      amountData_All[i][2] =  Math.round((amountData_All[i][3] + amountData_All[i][6] + amountData_All[i][7] + amountData_All[i][8])*25)/100; // Average - Covid
      amountData_All[i] = amountData_All[i].map(qty => (isQtyNotZero(qty)) ? qty : '') // Remove the zeros
      return item
    })

  const header = ['Descriptions', 'AVG (6 yr)', 'AVG - CoV', ...years.reverse(), 'Customers who purchased these items in ' + (currentYear - 2) + ' and ' + (currentYear - 1)];
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
  spreadsheet.getSheetByName('Search for Item Quantity or Amount ($)').getRange(1, 16, 4)
    .setValues([['Data was last updated on:\n\n' + new Date().toDateString()],[''],[''],
                ['Customers who purchased these items in ' + (currentYear - 1).toString() + ' and ' + currentYear.toString()]])
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

    if (sheet !== null) // Reverse the data so that it is in descending date (as apposed to ascending), so the concatenations between years is seamless i.e. December 2017 is followed by January 2018
      allData.push(...sheet.getSheetValues(2, 1, sheet.getLastRow() - 1, 8).reverse());
  })

  const lastRow = allData.unshift(['Item Description', 'Customer Name', 'Date', 'Invoice #', 'Location', 'Salesperson', 'Quantity', 'Amount']);
  spreadsheet.getSheetByName('All Data').clearContents().getRange(1, 1, lastRow, 8).setValues(allData)
}

/**
 * This function configures the yearly customer item data into the format that is desired for the spreadsheet to function optimally
 * 
 * @param {Object[][]}      values         : The values of the data that were just imported into the spreadsheet
 * @param {String}         fileName        : The name of the new sheet (which will also happen to be the xlxs file name)
 * @param {Boolean} doesPreviousSheetExist : Whether the previous sheet with the same name exists or not
 * @param {Spreadsheet}   spreadsheet      : The active spreadsheet
 * @author Jarren Ralf
 */
function configureYearlyCustomerItemData(values, fileName, doesPreviousSheetExist, spreadsheet)
{
  const currentYear = new Date().getFullYear();
  const customerSheet = spreadsheet.getSheetByName('Customer List');
  const accounts = customerSheet.getSheetValues(3, 1, customerSheet.getLastRow() - 2, 1).map(v => v[0].toString().trim())
  values.shift() // Remove the header
  values.pop()   // Remove the final row which contains descriptive stats
  const preData = values.filter(d => accounts.includes(d[0].toString().trim())); // Remove the account numbers that aren't on the Customer List
  const [data, ranges] = reformatData_YearlyCustomerItemData(preData) // This function spaces out the data and groups it by customer.
  const yearRange = new Array(currentYear - 2012 + 1).fill('').map((_, y) => (currentYear - y).toString()).reverse()
  var year = yearRange.find(p => p == fileName) // The year that the data is representing

  if (year == null) // The tab name in the spreadsheet doesn't not have the current year saved in it, so the user needs to be prompt so that we know the current year
  {
    const ui = SpreadsheetApp.getUi();
    const response = ui.prompt('Enter the year:')

    if (response.getSelectedButton() === ui.Button.OK)
    {
      year = response.getResponseText(); // Text response is assumed to be the year

      if (yearRange.includes(year))
      {
        const numCols = 6;
        const sheets = spreadsheet.getSheets();
        const previousSheet = sheets.find(sheet => sheet.getSheetName() == year)
        var indexAdjustment = 2010

        if (previousSheet != null)
        {
          indexAdjustment--;
          spreadsheet.deleteSheet(previousSheet)
        }

        SpreadsheetApp.flush();
        const newSheet = spreadsheet.insertSheet(year, sheets.length + indexAdjustment - year)
          .setColumnWidth(1, 66).setColumnWidth(2, 300).setColumnWidth(3, 150).setColumnWidth(4, 300).setColumnWidths(5, 2, 75);
        SpreadsheetApp.flush();
        const lastRow = data.unshift(['Customer', 'Customer Name', 'Item Number', 'Item Description', 'Quantity', 'Amount']);
        newSheet.deleteColumns(7, 20)
        newSheet.setTabColor('#a64d79').setFrozenRows(1)
        newSheet.protect()
        newSheet.getRange(1, 1, 1, numCols).setFontSize(11).setFontWeight('bold').setBackground('#c0c0c0')
          .offset(0, 0, lastRow, numCols).setHorizontalAlignments(new Array(lastRow).fill(['left', 'left', 'left', 'left', 'right', 'right'])).setNumberFormat('@').setValues(data)
        newSheet.getRangeList(ranges).setBorder(true, false, true, false, false, false).setBackground('#c0c0c0').setFontWeight('bold')

        const dashboard = spreadsheet.getSheetByName('Dashboard')

        if (currentYear > Number(dashboard.getRange('E2').getValue())) // The current year is not represented on the Dashboard, so add a column for it and make the relevant changes to formulas
        {
          const dashboard_lastRow = dashboard.getLastRow();
          dashboard.insertColumnBefore(5).getRange(2, 5, 2, 1).setValues([[currentYear], ['=SUM(E4:E' + dashboard_lastRow + ')']])
          const grandTotalRange = dashboard.getRange(4, 4, dashboard_lastRow - 3)
          dashboard.getRange(1, 5, 1, dashboard.getLastColumn() - 4).merge();
          grandTotalRange.setFormulas(grandTotalRange.getFormulas().map(formula => [formula[0].replace('F', 'E')]))
        }

        updateAllCustomersSalesData(spreadsheet)
      }
      else
      {
        ui.alert('Invalid Input', 'Please import your data again and enter a 4 digit year in the range of [2012, ' + currentYear + '].',)
        return;
      }
    }
    else
    {
      spreadsheet.toast('Data import proccess has been aborted.', '', 60)
      return;
    }
  }
  else
  {
    const numCols = 6;
    const sheets = spreadsheet.getSheets();
    Logger.log(year)
    const previousSheet = sheets.find(sheet => sheet.getSheetName() == year)
    Logger.log(previousSheet.getSheetName())
    var indexAdjustment = 2011

    if (doesPreviousSheetExist)
    {
      indexAdjustment--;
      spreadsheet.deleteSheet(previousSheet)
    }
    
    SpreadsheetApp.flush();
    const newSheet = spreadsheet.insertSheet(year, sheets.length + indexAdjustment - year)
      .setColumnWidth(1, 66).setColumnWidth(2, 300).setColumnWidth(3, 150).setColumnWidth(4, 300).setColumnWidths(5, 2, 75);
    SpreadsheetApp.flush();
    const lastRow = data.unshift(['Customer', 'Customer Name', 'Item Number', 'Item Description', 'Quantity', 'Amount']);
    newSheet.deleteColumns(7, 20)
    newSheet.setTabColor('#a64d79').setFrozenRows(1)
    newSheet.protect()
    newSheet.getRange(1, 1, 1, numCols).setFontSize(11).setFontWeight('bold').setBackground('#c0c0c0')
      .offset(0, 0, lastRow, numCols).setHorizontalAlignments(new Array(lastRow).fill(['left', 'left', 'left', 'left', 'right', 'right'])).setNumberFormat('@').setValues(data)
    newSheet.getRangeList(ranges).setBorder(true, false, true, false, false, false).setBackground('#c0c0c0').setFontWeight('bold')

    const dashboard = spreadsheet.getSheetByName('Dashboard')

    if (currentYear > Number(dashboard.getRange('E2').getValue())) // The current year is not represented on the Dashboard, so add a column for it and make the relevant changes to formulas
    {
      const dashboard_lastRow = dashboard.getLastRow();
      dashboard.insertColumnBefore(5).getRange(2, 5, 2, 1).setValues([[currentYear], ['=SUM(E4:E' + dashboard_lastRow + ')']])
      const grandTotalRange = dashboard.getRange(4, 4, dashboard_lastRow - 3)
      dashboard.getRange(1, 5, 1, dashboard.getLastColumn() - 4).merge();
      grandTotalRange.setFormulas(grandTotalRange.getFormulas().map(formula => [formula[0].replace('F', 'E')]));
    }

    updateAllCustomersSalesData(spreadsheet)
  }
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
  values.shift() // Remove the header
  values.pop()   // Remove the final row which contains descriptive stats
  const preData = removeNonImformativeSKUs(values.filter(d => accounts.includes(d[8].toString().trim())).sort(sortByDateThenInvoiveNumber))
  const data = reformatData_YearlyInvoiceData(preData)
  const year = new Array(currentYear - 2012 + 1).fill('').map((_, y) => (currentYear - y).toString()).reverse().find(p => p == data[0][2].getFullYear()) // The year that the data is representing
  const isSingleYear = data.every(date => date[2].getFullYear() == year);

  if (isSingleYear) // Does every line of this spreadsheet contain the same year?
  {
    const numCols = 10;
    const sheets = spreadsheet.getSheets();
    const previousSheet = sheets.find(sheet => sheet.getSheetName() == year)
    var indexAdjustment = 2011

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
    newSheet.setTabColor('#a64d79').setFrozenRows(1)
    newSheet.protect()
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
function createChartForAnnualSalesData()
{
  const currentYear = new Date().getFullYear();
  const numYears = currentYear - 2012 + 1
  const spreadsheet = SpreadsheetApp.getActive()
  const salesDataSheet = spreadsheet.getSheetByName('Annual Sales Data');
  const dataRng = salesDataSheet.getRange(4, 1, numYears, 2)
  const grandTotal = salesDataSheet.getSheetValues(2, 2, 1, 1)[0][0]

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
 * This function creates a chart sheet for the new customer that is being created by the user.
 * 
 * @param    {String}  customerName    : The name of the customer
 * @param    {String}    sheetName     : The name of the customer's data sheet
 * @param    {Sheet} customerDataSheet : The sheet containing the customer's data
 * @param {Spreadsheet} spreadsheet    : The active spreadsheet
 * @return {Number} The id of the sheet object that is created for the chart
 * @author Jarren Ralf
 */
function createChart_NewCustomer(customerName, sheetName, customerDataSheet, spreadsheet)
{
  const currentYear = new Date().getFullYear();
  const chartData = new Array(currentYear - 2012 + 1).fill('').map((_, y) => [(currentYear - y).toString(), '']).reverse()
  const numRows = chartData.length;
  const sheetName_Split = sheetName.split(' - ')
  
  const chartDataRng = customerDataSheet.setTabColor('#38761d').getRange(3, 5, numRows, 2).setBackground('white').setBorder(false, false, false, false, false, false)
    .setFontWeight('normal').setHorizontalAlignments(new Array(numRows).fill(['center', 'right'])).setNumberFormats(new Array(numRows).fill(['@', '$#,##0.00'])).setValues(chartData)
  customerDataSheet.setColumnWidth(5, 75).setColumnWidth(6, 100)
    .getRange(1, 1, 1, 5).setValues([[sheetName_Split[1], customerName, 'Total:', '=SUM(' + customerDataSheet.getRange(3, 6, numRows).getA1Notation() + ')', 'Chart Data']])
    .offset(0, 4, 1, 2).merge().setBorder(false, true, true, false, null, null)
    .offset(1, 0, 1, 2).setHorizontalAlignments([['center', 'right']]).setValues([['Year', 'Amount']])

  const chart = customerDataSheet.newChart()
    .asColumnChart()
    .addRange(chartDataRng)
    .setNumHeaders(0)
    .setXAxisTitle('Year')
    .setYAxisTitle('Sales Total')
    .setTransposeRowsAndColumns(false)
    .setMergeStrategy(Charts.ChartMergeStrategy.MERGE_COLUMNS)
    .setHiddenDimensionStrategy(Charts.ChartHiddenDimensionStrategy.IGNORE_BOTH)
    .setOption('title', customerName)
    .setOption('subtitle', 'Total: $' + new Intl.NumberFormat().format(twoDecimals(0)))
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

  customerDataSheet.protect()
  customerDataSheet.insertChart(chart);
  
  return spreadsheet.moveChartToObjectSheet(chart).activate().setName(sheetName_Split[0] + ' CHART - ' + sheetName_Split[1]).setTabColor('#f1c232').getSheetId();
}

/**
 * This function displays a hyperlink to take the user to the Charter & Guide Sales spreadsheet.
 * 
 * @author Jarren Ralf
 */
function display_HyperLinkToCharterAndGuideSalesSpreadsheet()
{
  showSidebar('Hyperlink_CharterAndGuide', 'Update Search for Invoice #s');
}

/**
 * This function displays a hyperlink to take the user to the Lodge Sales spreadsheet.
 * 
 * @author Jarren Ralf
 */
function display_HyperLinkToLodgeSalesSpreadsheet()
{
  showSidebar('Hyperlink_Lodge', 'Update Search for Invoice #s');
}

/**
 * This function displays the instructions for updating the Search for Invoice #s data.
 * 
 * @author Jarren Ralf
 */
function display_Invoice_Instructions()
{
  showSidebar('Instructions_Invoice', 'Update Search for Invoice #s');
}

/**
 * This function displays the instructions for updating the Search for Item Quantity or Amount ($) data.
 * 
 * @author Jarren Ralf
 */
function display_QuantityOrAmount_Instructions()
{
  showSidebar('Instructions_QuantityOrAmount', 'Update Search for Item Quantity or Amount');
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
    var info, numRows = 0, numCols = 1, maxRow = 2, maxCol = 3, isYearlyInvoiceData = 4, isYearlyCustomerItemData = 5;

    for (var sheet = sheets.length - 1; sheet >= 0; sheet--) // Loop through all of the sheets in this spreadsheet and find the new one
    {
      if (sheets[sheet].getType() == SpreadsheetApp.SheetType.GRID) // Some sheets in this spreadsheet are OBJECT sheets because they contain full charts
      {
        info = [
          sheets[sheet].getLastRow(),
          sheets[sheet].getLastColumn(),
          sheets[sheet].getMaxRows(),
          sheets[sheet].getMaxColumns(),
          sheets[sheet].getRange(1, 7).getValue().toString().includes('Quantity Specif'), // A characteristic of the invoice data
          sheets[sheet].getRange(1, 5).getValue().toString().includes('Quantity Specif')  // A characteristic of the customer item data
        ]
      
        // A new sheet is imported by File -> Import -> Insert new sheet(s) - The left disjunct is for a csv and the right disjunct is for an excel file
        if ((info[maxRow] - info[numRows] === 2 && info[maxCol] - info[numCols] === 2) || 
            (info[maxRow] === 1000 && info[maxCol] === 26 && info[numRows] !== 0 && info[numCols] !== 0) ||
            info[isYearlyInvoiceData] || info[isYearlyCustomerItemData]) 
        {
          spreadsheet.toast('Processing imported data...', '', 60)
          const values = sheets[sheet].getSheetValues(1, 1, info[numRows], info[numCols]); 
          const sheetName = sheets[sheet].getSheetName()
          const sheetName_Split = sheetName.split(' ')
          const doesPreviousSheetExist = sheetName_Split[1]
          var fileName = sheetName_Split[0];

          if (sheets[sheet].getSheetName().substring(0, 7) !== "Copy Of") // Don't delete the sheets that are duplicates
            spreadsheet.deleteSheet(sheets[sheet]) // Delete the new sheet that was created

          if (info[isYearlyInvoiceData])
          {
            configureYearlyInvoiceData(values, spreadsheet)
            spreadsheet.toast('The data will be updated in less than 5 minutes.', 'Import Complete.')
          }
          else if (info[isYearlyCustomerItemData])
          {
            configureYearlyCustomerItemData(values, fileName, doesPreviousSheetExist, SpreadsheetApp.openById('1o8BB1RWkxK1uo81tBjuxGc3VWArvCdhaBctQDssPDJ0')) // LODGE SALES spreadsheet
            configureYearlyCustomerItemData(values, fileName, doesPreviousSheetExist, SpreadsheetApp.openById('1kKS6yazOEtCsH-QCLClUI_6NU47wHfRb8CIs-UTZa1U')) // CHARTER & GUIDE SALES spreadsheet
            var triggerDate = new Date(new Date().getTime() + 10000); // Set a trigger for ten seconds from now
            Logger.log('All of the customer item YTD sales data will be begin compiling at:')
            Logger.log(triggerDate)
            ScriptApp.newTrigger('collectAllHistoricalData').timeBased().at(triggerDate).create();
            spreadsheet.getSheetByName('Search for Item Quantity or Amount ($)').getRange(1, 1).activate()
            spreadsheet.toast('The data will be updated in less than 5 minutes.', 'Import Complete.')
          }
          
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
 * This function protects all sheets expect for the search pages on the Lodge, Charter, & Guide data spreadsheet, for those, just the relevant cells in the header are protected.
 * 
 * @author Jarren Ralf
 */
function protectAllSheets()
{
  const users = ['triteswarehouse@gmail.com', 'scottnakashima10@gmail.com', 'scottnakashima@hotmail.com', 'pntparksville@gmail.com', 'derykdawg@gmail.com'];
  var sheetName, chartSheet = SpreadsheetApp.SheetType.OBJECT;

  SpreadsheetApp.getActive().getSheets().map(sheet => {
    if (sheet.getType() !== chartSheet)
    {
      sheetName = sheet.getSheetName();

      if (sheetName !== 'Search for Item Quantity or Amount ($)')
      {
        if (sheetName !==  'Search for Invoice #s')
          sheet.protect().addEditor('jarrencralf@gmail.com').removeEditors(users);
        else
          sheet.protect().addEditor('jarrencralf@gmail.com').removeEditors(users).setUnprotectedRanges([sheet.getRange(1, 1, 2)]);
      }
      else
        sheet.protect().addEditor('jarrencralf@gmail.com').removeEditors(users).setUnprotectedRanges([sheet.getRange(1, 1, 3), sheet.getRange(2, 5, 2), sheet.getRange(2, 9, 2), sheet.getRange(3, 11)]);
      }
  })
}

/**
 * This function spaces out the data and groups it by customer.
 * 
 * @param {String[][]} preData : The preformatted data.
 * @return {String[][], String[]} The reformatted data and a list of ranges to create a RangeList object
 * @author Jarren Ralf
 */
function reformatData_YearlyCustomerItemData(preData)
{
  var qty = 0, amount = 0, row = 0, uniqueCustomerList = [], ranges = [], formattedData = [];

  preData.map((customer, i, previousCustomers) => {
    if (uniqueCustomerList.includes(customer[0])) // Multiple Lines of Same Customer
    {
      qty += customer[4]
      amount += customer[5]
      formattedData.push(customer)
    }
    else if (uniqueCustomerList.length === 0) // First Customer
    {
      qty += customer[4]
      amount += customer[5]
      formattedData.push(customer)
      uniqueCustomerList.push(customer[0])
    }
    else // New Customer
    {
      formattedData.push([previousCustomers[i - 1][0], previousCustomers[ i -1][1], '', '', qty, amount], new Array(6).fill(''), customer)
      row = formattedData.length - 1;
      qty = customer[4];
      amount = customer[5];
      ranges.push('E' + row + ':F' + row)
      uniqueCustomerList.push(customer[0])
    }
  })

  const ii = preData.length - 1;

  // We need to add a row of totals for the final customer
  formattedData.push([preData[ii][0], preData[ii][1], '', '', qty, amount])
  row = formattedData.length + 1;
  ranges.push('E' + row + ':F' + row)

  return [formattedData, ranges]
}

/**
 * This function checks the invoice numbers and reformats the numbers that come from countersales so that they are all displayed in the same format. It also changes
 * the description to the standard Google description so that the items are more easily searched for.
 * 
 * @param {String[][]} preData : The preformatted data.
 * @return {String[][]} The reformatted data
 * @author Jarren Ralf
 */
function reformatData_YearlyInvoiceData(preData)
{
  const csvData = Utilities.parseCsv(DriveApp.getFilesByName("inventory.csv").next().getBlob().getDataAsString())
  const itemNum = csvData[0].indexOf('Item #');
  const fullDescription = csvData[0].indexOf('Item List')
  var item;

  return preData.map(itemVals => {
    item = csvData.find(val => val[itemNum] == itemVals[9])

    if (item != null)
      return (itemVals[3].toString().length === 9 && itemVals[3].toString().charAt('I')) ?
        [item[fullDescription], itemVals[1], itemVals[2], itemVals[3].substring(1), itemVals[4], itemVals[5], itemVals[6], itemVals[7], itemVals[8], itemVals[9]] :
      (itemVals[3].toString().length === 8 && itemVals[3].toString().charAt('I')) ?
        [item[fullDescription], itemVals[1], itemVals[2], '0' + itemVals[3].substring(1), itemVals[4], itemVals[5], itemVals[6], itemVals[7], itemVals[8], itemVals[9]] : 
        [item[fullDescription], itemVals[1], itemVals[2], itemVals[3], itemVals[4], itemVals[5], itemVals[6], itemVals[7], itemVals[8], itemVals[9]]
    else
      return (itemVals[3].toString().length === 9 && itemVals[3].toString().charAt('I')) ? 
        [itemVals[0] + ' - - - - ' + itemVals[9], itemVals[1], itemVals[2], itemVals[3].substring(1), itemVals[4], itemVals[5], itemVals[6], itemVals[7], itemVals[8], itemVals[9]] : 
      (itemVals[3].toString().length === 8 && itemVals[3].toString().charAt('I')) ? 
        [itemVals[0] + ' - - - - ' + itemVals[9], itemVals[1], itemVals[2], '0' + itemVals[3].substring(1), itemVals[4], itemVals[5], itemVals[6], itemVals[7], itemVals[8], itemVals[9]] : 
        [itemVals[0] + ' - - - - ' + itemVals[9], itemVals[1], itemVals[2], itemVals[3], itemVals[4], itemVals[5], itemVals[6], itemVals[7], itemVals[8], itemVals[9]]
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
  const fishingTackleSKUs = ["80000129", "80000389", "80000549", "80000349", "80000399", "80000499", "80000799", "80000409", "80000439", "80000599", "80000199", "80000249", "80000459", "80000699", "80000739", "80000999", "80001099", "80001149", "80001249", "80001499", "80001949", "80001999", "80000039", "80000089", "80000829", "80000259", "80000589", "80000899", "80000299", "80001199", "80001599", "80000649", "80000849", "80000025", "80000169", "80000579", "80000939", "80001299", "80000139", "80000329", "80000519", "80000629", "80000769", "80000015", "80000149", "80001549", "80000049", "80000949", "80001899", "80000020", "80000079", "80000179", "80000989", "80000449", "80000429", "80000099", "80001699", "80001649", "80001799", "80001849", "80000029", "80000339", "80000749", "80001399", "80000189", "80000289", "80000689", "80000069", "80000279", "80000159", "80000859", "80000729", "80000979", "80000059", "80000229", "80000119", "80000209", "80000219", "80000319", "80000359", "80000369", "80000419", "80000529", "80000639", "80000889", "80001749", "80000789", "80000609", "80000509", "80001049", "80000539", "80000659", "80001449", "80000109", "80000489", "80000759", "80000669", "80000469", "80000379", "80000869", "80000479", "80000679", "80000239", "80000719", "80000569", "80000709", "80000309", "80000919", "80001349", "80000879", "80000929", "80000269", "80000819", "80000619", "80000839", "80000959", "7000F6000", "7000F10000", "80002999", "7000F4000", "7000F5000", "7000F7000", "7000F3000", "7000F8000", "7000F20000", "7000F30000", "7000F9000", "80000779", "80000559", '7000M10000', '7000M200000', '7000M100000', '7000M125000', '7000M15000', '7000M150000', '7000M20000', '7000M3000', '7000M30000', '7000M4000', '7000M5000', '7000M50000', '7000M6000', '7000M7000', '7000M75000', '7000M8000', '7000M9000', 'FREIGHT', 'MISCITEM', 'MISCWEB', 'GIFT CERTIFICATE', 'BROKERAGE', 'ROPE SPLICE', '54002800', '54003600', '20110000', '7000C24999', '20120000', '90070020', '90070021', '90070022', '90070014', '7000C11999', '7000C19999', '90070011', '15000900', '90070010', '90070012', '25821000', '90070030']

  return data.filter(v => !fishingTackleSKUs.includes(v[9].toString()))
}

/**
 * This function removes the protections on all sheets.
 * 
 * @author Jarren Ralf
 */
function removeProtectionOnAllSheets()
{
  var chartSheet = SpreadsheetApp.SheetType.OBJECT;

  SpreadsheetApp.getActive().getSheets().map(sheet => {
    if (sheet.getType() !== chartSheet)
      sheet.getProtections(SpreadsheetApp.ProtectionType.SHEET)[0].remove()
  })
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
  const invoiceNumberList = [], highlightedRows = []
  const searchforItems_FilterByCustomer = sheet.getRange(1, 1, 2).clearFormat()                   // Clear the formatting of the range of the search box
    .setBorder(true, true, true, true, null, null, 'black', SpreadsheetApp.BorderStyle.SOLID_THICK) // Set the border
    .setFontFamily("Arial").setFontColor("black").setFontWeight("bold").setFontSize(14)             // Set the various font parameters
    .setHorizontalAlignment("center").setVerticalAlignment("middle")                                // Set the alignment
    .setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP)                                              // Set the wrap strategy
    .merge().trimWhitespace()                                                                       // Merge and trim the whitespaces at the end of the string
    .getValue().toString().toUpperCase().split(' BY ')                                             // Split the search string at the word 'not'

  const searchforItems_FilterByDate = searchforItems_FilterByCustomer[0].split(' IN ')
  const searchesOrNot = searchforItems_FilterByDate[0].split(' NOT ')
  const searches = searchesOrNot[0].split(' OR ').map(words => words.split(/\s+/)) // Split the search values up by the word 'or' and split the results of that split by whitespace

  if (isNotBlank(searches[0][0])) // If the value in the search box is NOT blank, then compute the search
  {
    spreadsheet.toast('Searching...', '', 30)

    if (searchforItems_FilterByCustomer.length === 1) // The word 'by' WASN'T found in the string
    {
      if (searchforItems_FilterByDate.length === 1) // The word 'in' wasn't sound in the string
      {
        if (searchesOrNot.length === 1) // The word 'not' WASN'T found in the string
        {
          const numSearches = searches.length; // The number searches
          const dataSheet = spreadsheet.getSheetByName('All Data')
          var data = dataSheet.getSheetValues(2, 1, dataSheet.getLastRow() - 1, 8);
          var numSearchWords, col; // Which column of data to search ** Default is item description

          switch (searches[0][0].substring(0, 3)) // Based on the search indicator, set the column of data to search in
          {
            case 'CUS':
              col = 1;
              searches[0].shift()
              break;
            case 'DAT':
              col = 2;
              searches[0].shift()
              break;
            case 'INV':
              col = 3;
              searches[0].shift()
              break;
            case 'LOC':
              col = 4;
              searches[0].shift()
              break;
            case 'SAL':
              col = 5;
              searches[0].shift()
              break;
            default:
              col = 0;
          }

          if (col < 2 || col === 3) // Search item descriptions, customer names or invoice numbers
          {
            for (var i = 0; i < data.length; i++) // Loop through all of the descriptions from the search data
            {
              loop: for (var j = 0; j < numSearches; j++) // Loop through the number of searches
              {
                numSearchWords = searches[j].length - 1;

                for (var k = 0; k <= numSearchWords; k++) // Loop through each word in each set of searches
                {
                  if (data[i][col].toString().toUpperCase().includes(searches[j][k])) // Does column 'col' of the i-th row of data contain the k-th search word in the j-th search
                  {
                    if (k === numSearchWords) // The last search word was succesfully found in the ith item, and thus, this item is returned in the search
                    {
                      if (col === 0) highlightedRows.push(data[i][0]) // Push description if we are doing a regular item search
                      if (!invoiceNumberList.includes(data[i][3])) invoiceNumberList.push(data[i][3]); // Add the invoice number to the list (if it is not already there)
                      break loop;
                    }
                  }
                  else
                    break; // One of the words in the User's query was NOT contained in the ith item description, therefore move on to the next item
                }
              }
            }
          }
          else if (col === 2) // Search the date
          {
            // The DATE.getMonth() function returns a numeral instead of the name of a month. Use this object to map to the month name.
            const months = {'JAN': 0, 'FEB': 1, 'MAR': 2, 'APR': 3, 'May': 4, 'JUN': 5, 'JUL': 6, 'AUG': 7, 'SEP': 8, 'OCT': 9, 'NOV': 10, 'DEC': 11} 

            for (var i = 0; i < data.length; i++) // Loop through all of the descriptions from the search data
            {
              loop: for (var j = 0; j < numSearches; j++) // Loop through the number of searches
              {
                numSearchWords = searches[j].length;

                if (numSearchWords === 1) // Assumed to be year only
                {
                  if (searches[j][0].toString().length === 4) // Check that the year is 4 digits
                  {
                    if (data[i][col].toString().toUpperCase().includes(searches[j][0])) // Does column 'col' of the i-th row of data contain the year being searched for
                    {
                      if (!invoiceNumberList.includes(data[i][3])) invoiceNumberList.push(data[i][3]); // Add the invoice number to the list (if it is not already there)
                      break loop;
                    }
                    else
                      break; // One of the words in the User's query was NOT contained in the ith item description, therefore move on to the next item
                  }
                }
                else if (numSearchWords === 2) // Assumed to be month and year
                {
                  if (searches[j][1].toString().length === 4 && searches[j][0].toString().length >= 3) // Check that the year is 4 digits, and the month is atleast 3 characters
                  {     
                    // Does column 'col' of the i-th row of data contain the year and month being searched for
                    if (data[i][col].getFullYear() == searches[j][1] && data[i][col].getMonth() == months[searches[j][0].substring(0, 3)])
                    {
                      if (!invoiceNumberList.includes(data[i][3])) invoiceNumberList.push(data[i][3]); // Add the invoice number to the list (if it is not already there)
                      break loop;
                    }
                    else
                      break; // One of the words in the User's query was NOT contained in the ith item description, therefore move on to the next item
                  }
                }
                else if (numSearchWords === 3) // Assumed to be day, month, and year
                {
                  // Check that the year is 4 digits, the month is atleast 3 characters, and the day is at most 2 characters
                  if (searches[j][2].toString().length === 4 && searches[j][1].toString().length >= 3 && searches[j][0].toString().length <= 2)
                  {
                    // Does column 'col' of the i-th row of data contain the year, month, and day being searched for
                    if (data[i][col].getDate() == searches[j][0] && data[i][col].getFullYear() == searches[j][2] && data[i][col].getMonth() == months[searches[j][1].substring(0, 3)])
                    {
                      if (!invoiceNumberList.includes(data[i][3])) invoiceNumberList.push(data[i][3]); // Add the invoice number to the list (if it is not already there)
                      break loop;
                    }
                    else
                      break; // One of the words in the User's query was NOT contained in the ith item description, therefore move on to the next item
                  }
                }
              }
            }
          }
          else // Search the location or salesperson ** So much data that we will limit the search to 3 years
          {
            const threeYearsAgo = new Date().getFullYear() - 3;

            endSearch: for (var i = 0; i < data.length; i++) // Loop through all of the descriptions from the search data
            {
              loop: for (var j = 0; j < numSearches; j++) // Loop through the number of searches
              {
                numSearchWords = searches[j].length - 1;

                for (var k = 0; k <= numSearchWords; k++) // Loop through each word in each set of searches
                {
                  if (data[i][col].toString().toUpperCase().includes(searches[j][k])) // Does column 'col' of the i-th row of data contain the k-th search word in the j-th search
                  {
                    if (k === numSearchWords) // The last search word was succesfully found in the ith item, and thus, this item is returned in the search
                    {
                      if (!data[i][2].toString().toUpperCase().includes(threeYearsAgo)) // Data is sorted in descending order, so the most recent years will not contains the (currentYear - 3)
                      {
                        if (!invoiceNumberList.includes(data[i][3])) invoiceNumberList.push(data[i][3]); // Add the invoice number to the list (if it is not already there)
                        break loop;
                      }
                      else // This is the first instance of a row of data with a date that contains a date from 3 years ago
                        break endSearch; // Up to three years of data has been found. Stop searching.
                    }
                  }
                  else
                    break; // One of the words in the User's query was NOT contained in the ith item description, therefore move on to the next item
                }
              }
            }
          }
        }
        else // The word 'not' was found in the search string
        {
          const dontIncludeTheseWords = searchesOrNot[1].split(/\s+/);
          const numWordsToNotInclude = dontIncludeTheseWords.length - 1;
          const numSearches = searches.length; // The number searches
          const dataSheet = spreadsheet.getSheetByName('All Data')
          var data = dataSheet.getSheetValues(2, 1, dataSheet.getLastRow() - 1, 8);
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
                    for (var l = 0; l <= numWordsToNotInclude; l++) // Loop through the number of words to not include
                    {
                      if (!data[i][0].toString().toUpperCase().includes(dontIncludeTheseWords[l])) // The i-th description DOES NOT contain the l-th word (of the words that shouldn't be included)
                      {
                        if (l === numWordsToNotInclude) // The i-th description does not include any the words that it is not suppose to
                        {
                          highlightedRows.push(data[i][0]) // Push description
                          if (!invoiceNumberList.includes(data[i][3])) invoiceNumberList.push(data[i][3]); // Add the invoice number to the list (if it is not already there)
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
      }
      else // The word 'in' was found in the string
      {
        const dateSearch = searchforItems_FilterByDate[1].toString().split(" ");

        if (searchesOrNot.length === 1) // The word 'not' WASN'T found in the string
        {




          // // The DATE.getMonth() function returns a numeral instead of the name of a month. Use this object to map to the month name.
          //   const months = {'JAN': 0, 'FEB': 1, 'MAR': 2, 'APR': 3, 'May': 4, 'JUN': 5, 'JUL': 6, 'AUG': 7, 'SEP': 8, 'OCT': 9, 'NOV': 10, 'DEC': 11} 

          //   for (var i = 0; i < data.length; i++) // Loop through all of the descriptions from the search data
          //   {
          //     loop: for (var j = 0; j < numSearches; j++) // Loop through the number of searches
          //     {
          //       numSearchWords = searches[j].length;

          //       if (numSearchWords === 1) // Assumed to be year only
          //       {
          //         if (searches[j][0].toString().length === 4) // Check that the year is 4 digits
          //         {
          //           if (data[i][col].toString().toUpperCase().includes(searches[j][0])) // Does column 'col' of the i-th row of data contain the year being searched for
          //           {
          //             if (!invoiceNumberList.includes(data[i][3])) invoiceNumberList.push(data[i][3]); // Add the invoice number to the list (if it is not already there)
          //             break loop;
          //           }
          //           else
          //             break; // One of the words in the User's query was NOT contained in the ith item description, therefore move on to the next item
          //         }
          //       }
          //       else if (numSearchWords === 2) // Assumed to be month and year
          //       {
          //         if (searches[j][1].toString().length === 4 && searches[j][0].toString().length >= 3) // Check that the year is 4 digits, and the month is atleast 3 characters
          //         {     
          //           // Does column 'col' of the i-th row of data contain the year and month being searched for
          //           if (data[i][col].getFullYear() == searches[j][1] && data[i][col].getMonth() == months[searches[j][0].substring(0, 3)])
          //           {
          //             if (!invoiceNumberList.includes(data[i][3])) invoiceNumberList.push(data[i][3]); // Add the invoice number to the list (if it is not already there)
          //             break loop;
          //           }
          //           else
          //             break; // One of the words in the User's query was NOT contained in the ith item description, therefore move on to the next item
          //         }
          //       }
          //       else if (numSearchWords === 3) // Assumed to be day, month, and year
          //       {
          //         // Check that the year is 4 digits, the month is atleast 3 characters, and the day is at most 2 characters
          //         if (searches[j][2].toString().length === 4 && searches[j][1].toString().length >= 3 && searches[j][0].toString().length <= 2)
          //         {
          //           // Does column 'col' of the i-th row of data contain the year, month, and day being searched for
          //           if (data[i][col].getDate() == searches[j][0] && data[i][col].getFullYear() == searches[j][2] && data[i][col].getMonth() == months[searches[j][1].substring(0, 3)])
          //           {
          //             if (!invoiceNumberList.includes(data[i][3])) invoiceNumberList.push(data[i][3]); // Add the invoice number to the list (if it is not already there)
          //             break loop;
          //           }
          //           else
          //             break; // One of the words in the User's query was NOT contained in the ith item description, therefore move on to the next item
          //         }
          //       }
          //     }
          //   }



          const numSearches = searches.length; // The number searches
          const dataSheet = spreadsheet.getSheetByName('All Data')
          var data = dataSheet.getSheetValues(2, 1, dataSheet.getLastRow() - 1, 8);
          var numSearchWords, col; // Which column of data to search ** Default is item description

          for (var i = 0; i < data.length; i++) // Loop through all of the descriptions from the search data
          {
            loop: for (var j = 0; j < numSearches; j++) // Loop through the number of searches
            {
              numSearchWords = searches[j].length - 1;

              for (var k = 0; k <= numSearchWords; k++) // Loop through each word in each set of searches
              {
                if (data[i][0].toString().toUpperCase().includes(searches[j][k])) // Does column 'col' of the i-th row of data contain the k-th search word in the j-th search
                {
                  if (k === numSearchWords) // The last search word was succesfully found in the ith item, and thus, this item is returned in the search
                  {
                    if (col === 0) highlightedRows.push(data[i][0]) // Push description if we are doing a regular item search
                    if (!invoiceNumberList.includes(data[i][3])) invoiceNumberList.push(data[i][3]); // Add the invoice number to the list (if it is not already there)
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
          const dontIncludeTheseWords = searchesOrNot[1].split(/\s+/);
          const numWordsToNotInclude = dontIncludeTheseWords.length - 1;
          const numSearches = searches.length; // The number searches
          const dataSheet = spreadsheet.getSheetByName('All Data')
          var data = dataSheet.getSheetValues(2, 1, dataSheet.getLastRow() - 1, 8);
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
                    for (var l = 0; l <= numWordsToNotInclude; l++) // Loop through the number of words to not include
                    {
                      if (!data[i][0].toString().toUpperCase().includes(dontIncludeTheseWords[l])) // The i-th description DOES NOT contain the l-th word (of the words that shouldn't be included)
                      {
                        if (l === numWordsToNotInclude) // The i-th description does not include any the words that it is not suppose to
                        {
                          highlightedRows.push(data[i][0]) // Push description
                          if (!invoiceNumberList.includes(data[i][3])) invoiceNumberList.push(data[i][3]); // Add the invoice number to the list (if it is not already there)
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
      }
    }
    else // The word 'by' was found in the string
    {
      const customersSearches = searchforItems_FilterByCustomer[1].split(' OR ').map(words => words.split(/\s+/)); // Multiple customers can be searched for

      if (customersSearches.length === 1) // Search for one customer
      {
        if (searchesOrNot.length === 1) // The word 'not' WASN'T found in the string
        {
          const numSearches = searches.length; // The number searches
          const numCustomerSearchWords = customersSearches[0].length - 1;
          const dataSheet = spreadsheet.getSheetByName('All Data')
          var data = dataSheet.getSheetValues(2, 1, dataSheet.getLastRow() - 1, 8);
          var numSearchWords;

          for (var i = 0; i < data.length; i++) // Loop through all of the customers and descriptions from the search data
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
                    for (var l = 0; l <= numCustomerSearchWords; l++) // Loop through each word in the customer search
                    {
                      if (data[i][1].toString().toUpperCase().includes(customersSearches[0][l])) // Does the i-th customer name contain the l-th search word
                      {
                        if (l === numCustomerSearchWords) // All of the customer search words were located in the customer's name
                        {
                          highlightedRows.push(data[i][0]) // Push description
                          if (!invoiceNumberList.includes(data[i][3])) invoiceNumberList.push(data[i][3]); // Add the invoice number to the list (if it is not already there)
                          break loop;
                        }
                      }
                      else
                        break; // One of the words in the User's query was NOT contained in the ith item description, therefore move on to the next item
                    }
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
          const dontIncludeTheseWords = searchesOrNot[1].split(/\s+/);
          const numWordsToNotInclude = dontIncludeTheseWords.length - 1;
          const numSearches = searches.length; // The number searches
          const numCustomerSearchWords = customersSearches[0].length - 1;
          const dataSheet = spreadsheet.getSheetByName('All Data')
          var data = dataSheet.getSheetValues(2, 1, dataSheet.getLastRow() - 1, 8);
          var numSearchWords;

          for (var i = 0; i < data.length; i++) // Loop through all of the customers and descriptions from the search data
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
                    for (var l = 0; l <= numWordsToNotInclude; l++) // Loop through the number of words to not include
                    {
                      if (!data[i][0].toString().toUpperCase().includes(dontIncludeTheseWords[l])) // The i-th description DOES NOT contain the l-th word (of the words that shouldn't be included)
                      {
                        if (l === numWordsToNotInclude) // The i-th description does not include any the words that it is not suppose to
                        {
                          for (var l = 0; l <= numCustomerSearchWords; l++) // Loop through the number of customer search words
                          {
                            if (data[i][1].toString().toUpperCase().includes(customersSearches[0][l])) // The i-th customer name contains the l-th word of the customer search
                            {
                              if (l === numCustomerSearchWords) // All of the customer search words were located in the customer's name
                              {
                                highlightedRows.push(data[i][0]) // Push description
                                if (!invoiceNumberList.includes(data[i][3])) invoiceNumberList.push(data[i][3]); // Add the invoice number to the list (if it is not already there)
                                break loop;
                              }
                            }
                            else
                              break; // One of the words in the User's query was NOT contained in the ith item description, therefore move on to the next item
                          }
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
      }
      else // Searching for multiple customers
      {
        if (searchesOrNot.length === 1) // The word 'not' WASN'T found in the string
        {
          const numSearches = searches.length; // The number searches
          const numCustomerSearches = customersSearches.length; // The number of customer searches
          const dataSheet = spreadsheet.getSheetByName('All Data')
          var data = dataSheet.getSheetValues(2, 1, dataSheet.getLastRow() - 1, 8);
          var numSearchWords, numCustomerSearchWords;

          for (var i = 0; i < data.length; i++) // Loop through all of the customers and descriptions from the search data
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
                    for (var l = 0; l < numCustomerSearches; l++) // Loop through the number of customer searches
                    {
                      numCustomerSearchWords = customersSearches[l].length - 1;

                      for (var m = 0; m <= numCustomerSearchWords; m++) // Loop through the number of customer search words
                      {
                        if (data[i][1].toString().toUpperCase().includes(customersSearches[l][m])) // Does the i-th customer name contain the m-th search word in the l-th search
                        {
                          if (m === numCustomerSearchWords) // The last customer search word was successfully found in the customer name
                          {
                            highlightedRows.push(data[i][0]) // Push description
                            if (!invoiceNumberList.includes(data[i][3])) invoiceNumberList.push(data[i][3]); // Add the invoice number to the list (if it is not already there)
                            break loop;
                          }
                        }
                        else
                          break;
                      }
                    }
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
          const dontIncludeTheseWords = searchesOrNot[1].split(/\s+/);
          const numWordsToNotInclude = dontIncludeTheseWords.length - 1;
          const numSearches = searches.length; // The number searches
          const numCustomerSearches = customersSearches.length; // The number of customer searches
          const dataSheet = spreadsheet.getSheetByName('All Data')
          var data = dataSheet.getSheetValues(2, 1, dataSheet.getLastRow() - 1, 8);
          var numSearchWords, numCustomerSearchWords;

          for (var i = 0; i < data.length; i++) // Loop through all of the customers and descriptions from the search data
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
                    for (var l = 0; l <= numWordsToNotInclude; l++) // Loop through the number of words to not include
                    {
                      if (!data[i][0].toString().toUpperCase().includes(dontIncludeTheseWords[l])) // The i-th description DOES NOT contain the l-th word (of the words that shouldn't be included)
                      {
                        if (l === numWordsToNotInclude) // The i-th description does not include any the words that it is not suppose to
                        {
                          for (var m = 0; m < numCustomerSearches; m++) // Loop through the number of customer searchs
                          {
                            numCustomerSearchWords = customersSearches[m].length - 1;

                            for (var n = 0; n <= numCustomerSearchWords; n++) // Loop through the number of customer search words
                            {
                              if (data[i][1].toString().toUpperCase().includes(customersSearches[m][n])) // Does the i-th customer name contain the n-th search word in the m-th search
                              {
                                if (n === numCustomerSearchWords) // The last customer search word was successfully found in the customer name
                                {
                                  highlightedRows.push(data[i][0]) // Push description
                                  if (!invoiceNumberList.includes(data[i][3])) invoiceNumberList.push(data[i][3]); // Add the invoice number to the list (if it is not already there)
                                  break loop;
                                }
                              }
                              else
                                break;
                            }
                          }
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
      const numHighlightedRows = highlightedRows.length;

      if (numItems !== 1)
      {
        if (numHighlightedRows > 1)
          searchResultsDisplayRange.setValue(numHighlightedRows + ' results found.\n\n' + numItems + ' total rows.')
        else if (numHighlightedRows === 1)
          searchResultsDisplayRange.setValue('1 result found.\n\n' + numItems + ' total rows.')
        else
          searchResultsDisplayRange.setValue(numItems +  ' results found.')
      }
      else
        searchResultsDisplayRange.setValue('1 result found.')
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
  const searchResultsDisplayRange = sheet.getRange(1, 13); // The range that will display the number of items found by the search
  const functionRunTimeRange = sheet.getRange(2, 13);      // The range that will display the runtimes for the search and formatting
  const itemSearchFullRange = sheet.getRange(6, 1, sheet.getMaxRows() - 5, 17); // The entire range of the Item Search page
  const checkboxes = sheet.getSheetValues(2, 6, 2, 7);
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
      const numCols = dataSheet.getLastColumn()
      const data = dataSheet.getSheetValues(2, 1, dataSheet.getLastRow() - 1, numCols);
      const numSearches = searches.length; // The number searches
      var numSearchWords;

      if (searches[0][0] === 'RECENT') // If the user's search begins with 'RECENT' then they are searching for information in the final column of data, which is the one that contains customer info
      {
        const lastColIndex = numCols - 1;

        if (searches[0].length !== 1) // Also contains a search phrase
        {
          searches[0].shift() // Remove the 'RECENT' keyword
          const searchPhrase = searches[0].join(" ") // Join the rest of the search terms together
          output.push(...data.filter(customer => (isNotBlank(customer[lastColIndex])) ? customer[lastColIndex].includes(searchPhrase) : false)) // If the final column is not blank, find the phrase
        }
        else // Return the last two years of data
          output.push(...data.filter(customer => isNotBlank(customer[lastColIndex])))
      }
      else
      {
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
    }
    else // The word 'not' was found in the search string
    {
      var dontIncludeTheseWords = searchesOrNot[1].split(/\s+/);

      const dataSheet = selectDataSheet(spreadsheet, checkboxes);
      const data = dataSheet.getSheetValues(2, 1, dataSheet.getLastRow() - 1, dataSheet.getLastColumn());
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
      var numCols = output[0].length
      var numFormats = (checkboxes[0][0]) ? new Array(numItems).fill(['@', ...new Array(numCols - 2).fill('$#,##0.00'), '@']) : new Array(numItems).fill([...new Array(numCols).fill('@')]);
      sheet.getRange('A6').activate(); // Move the user to the top of the search items
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
 * This function updates the sheet links on the dashboard.
 * 
 * @param {Spreadsheet} spreadsheet : The spreadsheet of the links that need to be updated.
 * @param   {String}        id      : The sheet ID of the annual sales chart.
 * @author Jarren Ralf
 */
function setSheetLinksOnDashboard(spreadsheet, id)
{
  const sheets = spreadsheet.getSheets();
  const dashboard = sheets.shift()
  const sheetNames = sheets.map(sheet => sheet.getSheetName().split(' - '))
  const numRows = dashboard.getLastRow() - 3

  const sheetLinks = dashboard.getSheetValues(4, 1, numRows, 1).map(custNum => {
    for (var s = 3; s < sheetNames.length; s++)
      if (custNum[0] === sheetNames[s][1])
        return [
          SpreadsheetApp.newRichTextValue().setText(custNum[0]).setLinkUrl('#gid=' + sheets[s    ].getSheetId()).build(), // Link to Customer data
          SpreadsheetApp.newRichTextValue().setText(custNum[0]).setLinkUrl('#gid=' + sheets[s + 1].getSheetId()).build()  // Link to Customer chart
        ]
  })

  dashboard.getRange(4, 1, numRows, 2).setRichTextValues(sheetLinks) // Set the sheet links on the dashboard
    .offset(-3, 4, 1, 1).setRichTextValues([[SpreadsheetApp.newRichTextValue().setText("Sale Totals").setLinkUrl('#gid=' + id).build()]]) // Set the sheet link for the annual sales chart
}

/**
 * This function displays a side bar.
 * 
 * @param {String} htmlFileName : The name of the html file.
 * @param {String}    title     : The title of the sidebar.
 * @author Jarren Ralf
 */
function showSidebar(htmlFileName, title) 
{
  SpreadsheetApp.getUi().showSidebar(HtmlService.createHtmlOutputFromFile(htmlFileName).setTitle(title));
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

/**
 * This function updates the charts on the CHARTER & GUIDE SALES spreadsheet by first deleting them and rebuilding them. This function also contains the feature that if 
 * runtime is going to exceed 6 minutes, the limit for google apps script, then the script creates a trigger that will re-run this function a few minutes later. This 
 * function creates the spreadsheets in a for-loop and if runtime will exceed 6 minutes, it stores the current value of the loop's incrementing variable in Google's 
 * CacheService, which stores string data that will expire after 6 minutes. On rerun, the function can call on the cache and resume within the for-loop where the script
 * was last stopped.
 * 
 * @author Jarren Ralf
 */
function updateAllCharterAndGuideCharts()
{
  var cache = CacheService.getDocumentCache();
  var currentSheet = Number(cache.get('current_sheet_charterAndGuide'));
  var [isIncomplete, sheetIndex, currentTime] = updateAllCharts(currentSheet, SpreadsheetApp.openById('1kKS6yazOEtCsH-QCLClUI_6NU47wHfRb8CIs-UTZa1U'))

  if (isIncomplete)
  {
    const REASONABLE_TIME_TO_WAIT = 60000; // One Minute
    cache.put('current_sheet_charterAndGuide', sheetIndex.toString()); // Store the indexing variable
    var triggerDate = new Date(currentTime + REASONABLE_TIME_TO_WAIT); // Set a trigger for a point in the future
    Logger.log('Next Trigger will run at:')
    Logger.log(triggerDate)

    ScriptApp.newTrigger("updateAllCharterAndGuideCharts").timeBased().at(triggerDate).create();
  }
}

/**
 * This function deletes and rebuilds all of the charts in a particular spreadsheet in order to update the subtitle of the graph, which is the total Sales for a particular customer.
 * 
 * @param   {Number}   currentSheet : The index of the current sheet which is used as the data source for the current chart being created.
 * @param {Spreadsheet} spreadsheet : The active spreadsheet.
 * @returns Whether the for-loop has concluded or not as well as the time the last iterate of the for-loop started plus the value or index of the iterate.
 * @author Jarren Ralf
 */
function updateAllCharts(currentSheet, spreadsheet)
{
  const startTime = new Date(); // The start time of this function
  const MAX_RUNNING_TIME = 300000; // Five minutes
  const sheets = spreadsheet.getSheets();
  const sheetNames = sheets.map(sheet => sheet.getSheetName().split(' - '));
  const numYears = new Date().getFullYear() - 2011;
  const numCustomerSheets = sheetNames.length - numYears - 1;
  const CUST_NAME = 0, SALES_TOTAL = 2;
  var chart, chartTitleInfo, currentTime = 0;

  if (currentSheet === 0) // If the cache was null, set the initial sheet index to 4
    currentSheet = 4;

  // Create the spreadsheets, notice that the index varibale needs to be converted to a number since the Cache stores data as string values
  for (var sheet = currentSheet; sheet < numCustomerSheets; sheet = sheet + 2)
  {
    currentTime = new Date().getTime();
    
    if (currentTime - startTime >= MAX_RUNNING_TIME) // If the function has been running for more than 5 minutes, then set the trigger to run this function again in a few minutes
      return [sheet < numCustomerSheets, sheet, currentTime];
    else
    {
      spreadsheet.deleteSheet(sheets[sheet + 1]); // Delete the chart
      chartTitleInfo = sheets[sheet].getRange(1, 2, 1, 3).getDisplayValues()[0];

      chart = sheets[sheet].newChart()
        .asColumnChart()
        .addRange(sheets[sheet].getRange(3, 5, numYears, 2))
        .setNumHeaders(0)
        .setXAxisTitle('Year')
        .setYAxisTitle('Sales Total')
        .setTransposeRowsAndColumns(false)
        .setMergeStrategy(Charts.ChartMergeStrategy.MERGE_COLUMNS)
        .setHiddenDimensionStrategy(Charts.ChartHiddenDimensionStrategy.IGNORE_BOTH)
        .setOption('title', chartTitleInfo[CUST_NAME])
        .setOption('subtitle', 'Total: ' + chartTitleInfo[SALES_TOTAL])
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

      sheets[sheet].insertChart(chart);
      spreadsheet.moveChartToObjectSheet(chart).setName(sheetNames[sheet][0] + ' CHART - ' + sheetNames[sheet][1]).setTabColor('#f1c232')
    }
  }

  const salesDataSheet = spreadsheet.getSheetByName('Sales Data');
  const spreadsheetName = spreadsheet.getName();
  spreadsheet.deleteSheet(spreadsheet.getSheetByName('ANNUAL ' + spreadsheetName + ' CHART')); // Delete previous sales chart

  const annualSalesChart = salesDataSheet.newChart()
    .asColumnChart()
    .addRange(salesDataSheet.getRange(4, 1, numYears, 2))
    .setNumHeaders(0)
    .setXAxisTitle('Year')
    .setYAxisTitle('Sales Total')
    .setTransposeRowsAndColumns(false)
    .setMergeStrategy(Charts.ChartMergeStrategy.MERGE_COLUMNS)
    .setHiddenDimensionStrategy(Charts.ChartHiddenDimensionStrategy.IGNORE_BOTH)
    .setOption('title', 'ANNUAL ' + spreadsheetName + ' DATA')
    .setOption('subtitle', 'Total: ' + salesDataSheet.getRange(2, 2).getDisplayValue())
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

  salesDataSheet.insertChart(annualSalesChart);
  const annualSalesChartID = spreadsheet.moveChartToObjectSheet(annualSalesChart).activate().setName('ANNUAL ' + spreadsheetName + ' CHART').setTabColor('#f1c232').getSheetId();
  setSheetLinksOnDashboard(spreadsheet, annualSalesChartID)

  const ss = SpreadsheetApp.openById('1xKw4GAtNbAsTEodCDmCMbPCbXUlK9OHv0rt5gYzqx9c') // The Lodge, Charter, & Guide Data spreadsheet
  const annualSalesDataSheet = ss.getSheetByName('Annual Sales Data');
  ss.deleteSheet(ss.getSheetByName('ANNUAL SALES CHART')); // Delete previous sales chart

  const annualSalesChart_BOTH = annualSalesDataSheet.newChart()
    .asColumnChart()
    .addRange(annualSalesDataSheet.getRange(4, 1, numYears, 2))
    .setNumHeaders(0)
    .setXAxisTitle('Year')
    .setYAxisTitle('Sales Total')
    .setTransposeRowsAndColumns(false)
    .setMergeStrategy(Charts.ChartMergeStrategy.MERGE_COLUMNS)
    .setHiddenDimensionStrategy(Charts.ChartHiddenDimensionStrategy.IGNORE_BOTH)
    .setOption('title', 'Annual Sales Data')
    .setOption('subtitle', 'Total: ' + annualSalesDataSheet.getRange(2, 2).getDisplayValue())
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

  annualSalesDataSheet.insertChart(annualSalesChart_BOTH);
  ss.moveChartToObjectSheet(annualSalesChart_BOTH).activate().setName('ANNUAL SALES CHART').setTabColor('#f1c232');

  return [sheet < numCustomerSheets, sheet, currentTime];
}

/**
 * This function updates the charts on the LODGE SALES spreadsheet by first deleting them and rebuilding them. This function also contains the feature that if 
 * runtime is going to exceed 6 minutes, the limit for google apps script, then the script creates a trigger that will re-run this function a few minutes later. This 
 * function creates the spreadsheets in a for-loop and if runtime will exceed 6 minutes, it stores the current value of the loop's incrementing variable in Google's 
 * CacheService, which stores string data that will expire after 6 minutes. On rerun, the function can call on the cache and resume within the for-loop where the script
 * was last stopped.
 * 
 * @author Jarren Ralf
 */
function updateAllLodgeCharts()
{
  var cache = CacheService.getDocumentCache();
  var currentSheet = Number(cache.get('current_sheet_lodge'));
  var [isIncomplete, sheetIndex, currentTime] = updateAllCharts(currentSheet, SpreadsheetApp.openById('1o8BB1RWkxK1uo81tBjuxGc3VWArvCdhaBctQDssPDJ0'))

  if (isIncomplete)
  {
    const REASONABLE_TIME_TO_WAIT = 60000; // Thirty seconds
    cache.put('current_sheet_lodge', sheetIndex.toString()); // Store the indexing variable
    var triggerDate = new Date(currentTime + REASONABLE_TIME_TO_WAIT); // Set a trigger for a point in the future
    Logger.log('Next Trigger will run at:')
    Logger.log(triggerDate)

    ScriptApp.newTrigger("updateAllLodgeCharts").timeBased().at(triggerDate).create();
  }
}

/**
 * This function looks through all of the sheets and updates the sales date for all years since 2012, for all customers. The function
 * finishes by updating the Dashboard with the sales data.
 * 
 * @param {Spreadsheet} spreadsheet : The active spreadsheet.
 * @author Jarren Ralf
 */
function updateAllCustomersSalesData(spreadsheet)
{
  if (arguments.length === 0)
    spreadsheet = SpreadsheetApp.getActive()

  const today = new Date();
  const currentYear = today.getFullYear();
  const currentDate = (today.getDate() < 10) ? '0' + today.getDate() : today.getDate() + ' ' + 
    ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'][today.getMonth()] + ' ' + currentYear; // The current date is shown on the customer's data page
  const numYears = currentYear - 2012 + 1;
  const sheets = spreadsheet.getSheets();
  const dashboard = sheets.shift()
  const sheetNames = sheets.map(sheet => sheet.getSheetName().split(' - '));
  const numCustomerSheets = sheetNames.length - numYears - 1
  const range = dashboard.getRange(4, 5, dashboard.getLastRow() - 3, dashboard.getLastColumn() - 4)
  const salesTotals = range.getValues();
  const hAligns = ['left', 'left', 'right', 'right'], numFormats = ['@', '@', '@', '$#,##0.00']
  const chartDataFormat = new Array(numYears).fill().map(() => ['@', '$#,##0.00']);
  const chartDataH_Alignment = new Array(numYears).fill().map(() => ['center', 'right']);
  var sheet, data, numItems = 0, chartData = [], index = 0, allYearsData, salesData, hAlignments = [], numberFormats = [], 
    yearRange = [], yearRange_RowNum = 3, totalRange = [], totalRange_RowNum = 1;
  
  const years = new Array(numYears).fill('').map((_, y) => (currentYear - y).toString()).map(year_y => {
    chartData.push([year_y, ''])
    sheet = spreadsheet.getSheetByName(year_y)

    return (sheet !== null) ? sheet.getSheetValues(2, 1, sheet.getLastRow() - 1, 6) : sheet;
  }).filter(u => u !== null)

  chartData.reverse();

  for (var s = 3; s < numCustomerSheets; s = s + 2) // Loop through all of the customer sheets
  {
    spreadsheet.toast((index + 1) + ': ' + sheetNames[s][0] + ' - ' + sheetNames[s][1], 'Updating...', 60)
    
    allYearsData = years.map((fullYearData, y) => {
      data = fullYearData.filter(custNum => custNum[0].trim() === sheetNames[s][1])  // Retrieve just the customers data
      numItems = data.length;

      if (numItems !== 0)
      {
        chartData[numYears - y - 1][1] = data[numItems - 1][5]; // Fill in the chart data with the yearly totals
        salesTotals[index][y] = data[numItems - 1][5]; // Fill in the the sales totals for the current customer for year y on the dashboard
        ((currentYear - y) == currentYear) ? 
          data.unshift(['', '', '', '', '01 Jan ' + currentYear, currentDate]) : 
          data.unshift(['', '', '', '', '01 Jan ' + (currentYear - y), '31 Dec ' + (currentYear - y)])
        data.push(['', '', '', '', '', '']);
        totalRange_RowNum += (totalRange_RowNum == 0) ? numItems + 1 : numItems + 2;
        totalRange.push('C' + totalRange_RowNum + ':D' + totalRange_RowNum)
        yearRange.push('C' + yearRange_RowNum + ':D' + yearRange_RowNum)
        yearRange_RowNum += numItems + 2;
      }
      else
      {
        chartData[numYears - y - 1][1] = ''
        salesTotals[index][y] = ''; 
      }

      return data.map(col => [col[2], col[3], col[4], col[5]])
    })

    index++

    salesData = [].concat.apply([], allYearsData);
    salesData.pop()

    hAlignments = new Array(salesData.length).fill().map(() => hAligns)
    numberFormats = new Array(salesData.length).fill().map(() => numFormats)

    sheets[s].getRange(3, 1, sheets[s].getMaxRows() - 2, 6).clearContent().setBackground('white').setBorder(false, false, false, false, false, false)
      .offset(0, 0, salesData.length, 4).setFontWeight('normal').setVerticalAlignment('middle').setHorizontalAlignments(hAlignments).setNumberFormats(numberFormats).setValues(salesData)
      .offset(0, 4, numYears, 2).setNumberFormats(chartDataFormat).setHorizontalAlignments(chartDataH_Alignment).setFontWeight('normal').setValues(chartData)
      .offset(-2, -1, 1, 1).setFormula([['=SUM(F3:F' + (numYears + 2) + ')']])

    sheets[s].getRangeList(yearRange).setFontWeight('bold').setNumberFormat('@') // The year
    sheets[s].getRangeList(totalRange).setBorder(true, false, true, false, false, false).setBackground('#c0c0c0').setFontWeight('bold') // The total quantity and amount

    // Reset the variables
    yearRange.length = 0;
    totalRange.length = 0;
    hAlignments.length = 0;
    numberFormats.length = 0;
    yearRange_RowNum = 3;
    totalRange_RowNum = 1;
  }

  const yearlySales = range.setNumberFormat('$#,##0.00').setValues(salesTotals).activate().offset(-1, 0, 1, numYears).getDisplayValues()[0].reverse();
  const annualSalesData = [];

  // Update the sales data for the annual chart
  if (spreadsheet.getName().split(' ', 1)[0] !== 'CHARTER')
  {
    var charterGuideSalesYearlyData = SpreadsheetApp.openById('1kKS6yazOEtCsH-QCLClUI_6NU47wHfRb8CIs-UTZa1U').getSheetByName('Sales Data').getDataRange().getDisplayValues();
    charterGuideSalesYearlyData.shift()
    charterGuideSalesYearlyData.shift()
    charterGuideSalesYearlyData.shift()

    var annualChartData = yearlySales.map((total, y) => {
      annualSalesData.push([(2012 + y).toString(), '=SUM(C' + (y + 4) + ':D' + (y + 4) + ')', 
        total, (charterGuideSalesYearlyData[y] != null) ? charterGuideSalesYearlyData[y][1] : '$0.00'])
      return [(2012 + y).toString(), total]
    });
  }
  else
  {
    var lodgeSalesYearlyData = SpreadsheetApp.openById('1o8BB1RWkxK1uo81tBjuxGc3VWArvCdhaBctQDssPDJ0').getSheetByName('Sales Data').getDataRange().getDisplayValues();
    lodgeSalesYearlyData.shift()
    lodgeSalesYearlyData.shift()
    lodgeSalesYearlyData.shift()

    var annualChartData = yearlySales.map((total, y) => {
      annualSalesData.push([(2012 + y).toString(), '=SUM(C' + (y + 4) + ':D' + (y + 4) + ')', 
        (lodgeSalesYearlyData[y] != null) ? lodgeSalesYearlyData[y][1] : '$0.00', total])
      return [(2012 + y).toString(), total]
    })
  }

  SpreadsheetApp.openById('1xKw4GAtNbAsTEodCDmCMbPCbXUlK9OHv0rt5gYzqx9c').getSheetByName('Annual Sales Data').getRange(4, 1, numYears, 4).setValues(annualSalesData)
  spreadsheet.getSheetByName('Sales Data').getRange(4, 1, numYears, 2).setNumberFormats(chartDataFormat).setValues(annualChartData)

  var triggerDate = new Date(new Date().getTime() + 60000); // Set a trigger for one minute
  Logger.log('All of the charts will begin updating at:')
  Logger.log(triggerDate)      
  const functionName = (spreadsheet.getName().split(" ", 1)[0] !== "CHARTER") ? "updateAllLodgeCharts" : "updateAllCharterAndGuideCharts";
  ScriptApp.newTrigger(functionName).timeBased().at(triggerDate).create();
  spreadsheet.toast('', 'Full Data Update: COMPLETE', 60)
}
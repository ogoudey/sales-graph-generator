/*

Connect this code to a spreadsheet by opening a spreadsheet and going Extensions > Apps Script. Then paste this code into a .gs file.
When running, Google will ask for generous permissions.

The spreadsheet can be anywhere as long as it references the directory with the data.

For full functionality, the only requirement for that spreadsheet is two have four buttons assigned to 4 scripts

Green button -----> showCustomPopup
Red button -----> chart
Green reset -----> resetSheets
Red reset -----> resetChart

A complete usage would be:
1. Green button (enter some text until input is 'good')
2. Red button to plot it
3. Red reset to clear the chart
4. Green reset to start over

*/

// We use ss throughout to refer to the Graph Generator spreadsheet.
ss = SpreadsheetApp.getActiveSpreadsheet()

// Functions listed in sequential order...

// Assigned to the green button.
// dialogue argument is for LLM interaction.
// Usage note: who know what the LLM will do with the fields.
function showCustomPopup(dialogue="Please enter start date and end date...") {
  html = HtmlService.createHtmlOutput(`
    <div style="font-family: Arial, sans-serif; text-align: center; padding: 20px;">
      <h3>${dialogue}</h3>
      <p>-----</p>
      <label for="field1">First Field:</label>
      <input type="text" id="field1" style="width: 80%; padding: 8px; margin: 10px 0;">
      
      <br>
      
      <label for="field2">Second Field:</label>
      <input type="text" id="field2" style="width: 80%; padding: 8px; margin: 10px 0;">
      
      <br><br>
      
      <button onclick="submitData()" style="padding: 10px 20px; background: #4285F4; color: white; border: none; cursor: pointer;">Submit</button>
      
      <script>
        function submitData() {
          var field1 = document.getElementById("field1").value;
          var field2 = document.getElementById("field2").value;
          
          google.script.run.processFormData(field1, field2);
          google.script.host.close();
        }
      </script>
    </div>
  `).setWidth(400).setHeight(500);
  
  SpreadsheetApp.getUi().showModalDialog(html, "Input Dates");
}

// Called upon submission of the above prompt. Either 
function processFormData(field1, field2) {
  console.log("Processing...")
  data = scrapeSpreadSheets(field1, field2)
  done = data[0] // success/fail
  dialogue = data[1] // LLM feedback
  if (!done){
    showCustomPopup(dialogue)
  }
  else{
    console.log("Done.")
  }
  ss.setActiveSheet(ss.getSheets()[0])
  takeData()
}


// Takes 
function scrapeSpreadSheets(start_date, end_date) {

  folder_name = "graph generator" // CHANGE
  files = getFilesFromFolder(folder_name) // a file iterator object
  
  full_list_of_spreadsheet_names = []
  full_list_of_spreadsheets = []


  // Filter all spreadsheets from list of files
  while (files.hasNext()) {
    file = files.next();
    
    if (file.getMimeType() === MimeType.GOOGLE_SHEETS) {
      spreadsheet_name = file.getName();
      console.log("Found Spreadsheet: " + spreadsheet_name)
      full_list_of_spreadsheet_names.push(spreadsheet_name)
      
      full_list_of_spreadsheets.push(file)
    }
  }

  // Methodology for selecting the requested range of spreadsheets: use an LLM!

  // Prompt formation, semantic regurgitation (TODO: fine tune, test)
  context = "You are a assistive chatbot employed by Daily Table, a company in Cambridge, Massachusetts. It is your job to provide a helpful, productive interface for a tool that employees will use. Using Google Sheets, employees will fill in certain cells on their shift from time to time. Occasionally, Google Apps Scripts (GAS) function call will trigger THIS request. So, you are inside a Google Sheet! Humour that!"
  job_assignment = '\tAnyway, your particular job here is to take their `start_date` and `end_date` inputs (below), and return the subset of the sheets that WITHIN that range (from the full list below). The spreadsheets suggest their data at the end of their name, in year-month-day form. (E.g. spreadsheet2024-06-15 is clearly June 15th 2024.) If their inputs are not specific enough, politely guide the user towards entering text that you can precisely map to the range requested. Please respond with EITHER: 1. ONLY a list of the spreadsheets in the range start_date -- end_date and nothing else, IN PROPER JSON and IN ORDER, or 2. guidance for the user. Here are examples for the two types of responses: 1. ["data2034-09-25", "data2034-10-02"] (notice the doublequotes around the sheet names and the order), or 2. "Please reformat to something I can understand." For giving guidance, be creative! Be friendly and jovial. They are like your coworkers.'

  system_content = context + "\n" + job_assignment + "\n\n" + full_list_of_spreadsheet_names
  messages = [
    { role: "system", content: system_content },
    { role: "user", content: `start_date: ${start_date}, end_date: ${end_date}` }
  ]
  response = callOpenAI(messages, true)
  try {
    sheets_to_find = JSON.parse(response)
  } catch (error) {
    return [false, response]
  }
  console.log("Response is an array (True)? " + Array.isArray(sheets_to_find))

  if (!Array.isArray(sheets_to_find)){
    return [false, response] // redundant?
  }

  sheets_to_find.forEach((sheet_name) => {
    console.log("To include: " + sheet_name)
  })
  
  full_list_of_spreadsheets.forEach((file) => {
    spreadsheet_name = file.getName();

    if (sheets_to_find.includes(spreadsheet_name)) {
      spreadsheet = SpreadsheetApp.openById(file.getId());
      console.log("Opened Spreadsheet: " + spreadsheet_name);
      sheet = spreadsheet.getSheets()[0]; // The important-sheet-at-0 assumption

      new_sheet = ss.insertSheet(spreadsheet_name)
      sourceData = sheet.getDataRange().getValues(); // Get all data
      new_sheet.getRange(1, 1, sourceData.length, sourceData[0].length).setValues(sourceData)

    }
    }
  )
  // returns whether success and (if failure) LLM 'feedback'
  return [true, null]
}

// Self-explanatory helper function
function getFilesFromFolder(folder_name="graph generator") {
  folders = DriveApp.getFoldersByName(folder_name);
  if (!folders.hasNext()) {
    console.log("Folder not found: " + folder_name);
    return;
  }
  folder = folders.next();
  return folder.getFiles();
  
}

// Called to select from all spreadsheets the ones within the range.
function callOpenAI(messages, verbose=false) {
  apiKey = ""; // OpenAI API key
  url = "https://api.openai.com/v1/chat/completions";

  payload = {
    model: "gpt-3.5-turbo", // Specify model
    messages,
    temperature: 0.7
  };

  options = {
    method: "post",
    contentType: "application/json",
    headers: {
      Authorization: "Bearer " + apiKey
    },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  response = UrlFetchApp.fetch(url, options);
  json = JSON.parse(response.getContentText());
  
  content = json.choices[0].message.content; // Logs response in Google Apps Script execution log
  if (verbose) {
    console.log(messages)
    console.log(content)
  }
  return content
}


// Function for relevant data from sheets 1--n to sheet 0.
// N.B. Google Sheets uses indices 1 - n where GAS or any code uses 0-n. E.g. row[0] is Row 1.
function takeData() {
  
  s = ss.getSheets()[0]

  sheets = ss.getSheets().slice(1) // sheets 1--n

  headerSet = new Set() // sheet0 left headers will be the union of all headers of sheets 1--n

  sheets.forEach((sheet) => {
    lastRow = sheet.getLastRow();  
    headers = sheet.getRange(2, 1, lastRow, 1).getValues(); // returns a 2D array.
    headers.forEach(function(row) {
      header = row[0];
      // Check for non-empty headers
      if (header !== "" && header !== null && header !== undefined) {
        headerSet.add(header);
      }
    });
  
    
  })

  unionHeaders = Array.from(headerSet);
  // sort the headers alphabetically (why not)
  unionHeaders.sort();
  output = unionHeaders.map(function(header) {
    return [header];
  });
  
  // Set column A to the union of headers, leaving row 1 empty.
  s.getRange(2, 1, output.length, 1).setValues(output);

  // Now go through the columns. One for each spreadsheet (now sheet)
  for (i = 1; i < ss.getSheets().length; i++){
    sheet = ss.getSheets()[i]
    // Set column header to spreadsheet (now sheet) name
    s.getRange(1, i+1).setValue(sheet.getName());
    for (j = 2; j < s.getLastRow()+1; j++){
      // Here's the painfully slow part I think...
      formula = "=IFERROR(VLOOKUP(A" + j + ", '" + sheet.getName() + "'!$A:$I, 9, FALSE), 0)";
      s.getRange(j, i + 1).setFormula(formula);
    }
  }

  

}

// Part II. Plot! Assigned to the red button.
function chart() {
  s = ss.getSheets()[0]
  dataRange = s.getDataRange()
  chart = s.newChart()
      .setChartType(Charts.ChartType.LINE)
      .addRange(dataRange)
      .setTransposeRowsAndColumns(true) //  Plots items over time
      .setNumHeaders(1) // Makes legend legible
      .setPosition(5, 5, 0, 0)
      .setOption('title', 'Sales by Item per week')
      // Optionally, you can set axis titles:
      .setOption('hAxis', { title: 'Item' })
      .setOption('vAxis', { title: 'Sales' })
      .setOption('legend', { position: 'right' })
      .build();
  
  // Insert the chart into the sheet.
  s.insertChart(chart);
}

// The reset buttons:

//green
function resetSheets() {
  while (ss.getSheets().length > 1) {
    ss.deleteSheet(ss.getSheets()[1])
  }
  range = ss.getSheets()[0].getDataRange().clearContent()
}

//red
function resetChart() {
  charts = ss.getSheets()[0].getCharts();
  for (i = 0; i < charts.length; i++) {
    ss.getSheets()[0].removeChart(charts[i]);
  }
  tempSheet = ss.getSheetByName("TempSheet");
  if (tempSheet) {
    ss.deleteSheet(tempSheet);
  }
  
}

/*

____Graveyard____

*/

/*
  NOT USED. The alternative to text interpretation: parse the input. Why is this a bad idea?
  1. Users don't want to be so strict.
  2. The spreadsheet names might be inconsistent.
*/
function conformDate(date) {
  year = date.getFullYear()
  month = String(date.getMonth() + 1).padStart(2, "0")
  day = String(date.getDate()).padStart(2, "0");
  return `data_spreadsheet${year}-${month}-${day}`
}


// Attempt 1 for plotting. Also NOT USED:
function plotData() {
  s = ss.getSheets()[0]
  range = s.getDataRange()
  dataValues = range.getValues()
  processedData = dataValues.map(function(row) {
    return row.map(function(cell) {
      // Check if the cell's string representation starts with "#".
      if (cell != null && cell.toString().startsWith("#")) {
        return 0;
      } else {
        return cell;
      }
    });
  });
  numRows = processedData.length;
  numCols = processedData[0].length;

  chartData = []

  header = ["X"]


  for (i = 0; i < numRows; i++) {
    console.log(i)
    header.push(processedData[i][0]) // The leftmost column are headers
  }
  chartData.push(header);

  // For each column (data point), create a row in chartData:
  // first column is the X value, then one value from each series (i.e. each original row)
  for (j = 1; j < numCols; j++) {
    rowData = [j + 1]; // X-axis value (for example, data point number)
    for (i = 0; i < numRows; i++) {
      rowData.push(processedData[i][j]);
    }
    chartData.push(rowData);
  }

  chartSheetName = "TempSheet"
  chartSheet = ss.getSheetByName(chartSheetName);
  if (!chartSheet) {
    chartSheet = ss.insertSheet(chartSheetName);
  } else {
    chartSheet.clear(); // Clear previous data if needed.
  }
  chartSheet.getRange(1, 1, chartData.length, chartData[0].length).setValues(chartData);

  dataRangeForChart = chartSheet.getRange(1, 1, chartData.length, chartData[0].length);

  chart = s.newChart()
      .setChartType(Charts.ChartType.LINE)
      .addRange(dataRangeForChart)
      .setPosition(2, s.getLastColumn() + 2, 0, 0)
      .setOption('title', "Line Chart (Rows as Series)")
      .setOption('hAxis', { title: 'Week' })
      .setOption('useFirstColumnAsDomain', true)
      .build()
  // Insert the chart into the sheet.
  s.insertChart(chart);
  ss.setActiveSheet(ss.getSheets()[0])

}

/* Acknowledgements:

The author of this file used ChatGPT o3-mini-high to speed up development, and some lines were generated by it.

*/

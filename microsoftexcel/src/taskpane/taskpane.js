/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global console, document, Excel, Office */

Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("create-table").onclick = () => tryCatch(collectInput);
  }
});

function collectInput() {
  const apiKey = document.getElementById("mp-auth").value;
  const loggedInUser = document.getElementById("mp-username").value;
  const userError = document.getElementById("error-username");
  const keyError = document.getElementById("error-key");
  userError.innerHTML = "";
  keyError.innerHTML = "";

  if (!loggedInUser & !apiKey) {
    keyError.innerHTML = "Please enter your /**/ API key";
    userError.innerHTML = "Please enter your /**/ username";
    return;
  }
  if (!apiKey) {
    keyError.innerHTML = "Please enter your /**/ API key";
    return;
  }
  if (!loggedInUser) {
    userError.innerHTML = "Please enter your /**/ username";
  }
  createTable(apiKey, loggedInUser);
}

async function createTable(apiKey, loggedInUser) {
  await Excel.run(async (context) => {
    const currentWorksheet = context.workbook.worksheets.getActiveWorksheet();
    const log = document.getElementById("log");
    log.innerHTML = "";
    
    const apiUrl = 'https://docs./**/.com/get-v1-data';
    const headers = {
      'Authorization': 'Bearer ' + apiKey,
      'User-Email': loggedInUser
    };
    const response = await fetch(apiUrl, { headers });
    if (response.ok) {
      const responseData = await response.json();
    
      if (responseData.data && responseData.data.length > 0) {
        const data = responseData.data;
    
        const headers = Object.keys(data[0]); 
        currentWorksheet.getUsedRange().clear(Excel.ClearApplyTo.worksheet);

        const transactionsTable= currentWorksheet.tables.add(`A1:${String.fromCharCode(65 + headers.length - 1)}1`, true /*hasHeaders*/);
        transactionsTable.name = "TransactionsTable";
        transactionsTable.getHeaderRowRange().values = [headers.map(header => header.replace(/_/g, ' '))]; // Replace underscores with spaces
        transactionsTable.rows.add(null /*add at the end*/, data.map(item => Object.values(item)));
        
        // Format specific columns if needed
        const amountColumnIndex = headers.indexOf('amount');
        if (amountColumnIndex !== -1) {
          transactionsTable.columns.getItemAt(amountColumnIndex + 1).getRange().numberFormat = [['\u20AC#,##0.00']];
        }
    
        transactionsTable.getRange().format.autofitColumns();
        transactionsTable.getRange().format.autofitRows();
      } else {
        console.error('Data is missing or empty in the response.');
        og.innerHTML ='Data is missing or empty in the response.';

      }
    } else {
      console.error('Failed to fetch data:', response.status, response.statusText);
      log.innerHTML = `Failed to fetch data: ${response.statusText}`;
      
    }
  
    await context.sync();
  });
}


/** Default helper for invoking an action and handling errors. */
async function tryCatch(callback) {
  try {
      await callback();
  } catch (error) {
      // Note: In a production add-in, you'd want to notify the user through your add-in's UI.
      console.error(error);
      log.innerHTML = `An error occurred during data fetch: ${error.message}`
  }
}

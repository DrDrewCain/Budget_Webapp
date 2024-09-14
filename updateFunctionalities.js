function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Budget and Loan Tracker')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function getOrCreateSheet(sheetName) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    switch (sheetName) {
      case 'Expenses':
        sheet.appendRow(['Date', 'Amount', 'Description', 'Category']);
        break;
      case 'Categories':
        sheet.appendRow(['Category']);
        var defaultCategories = ['Food', 'Gifts', 'Health/Medical', 'Home', 'Transportation', 'Personal', 'Pets', 'Utilities', 'Travel', 'Debt', 'Other'];
        defaultCategories.forEach(function (category) {
          sheet.appendRow([category]);
        });
        break;
      case 'Loans':
        sheet.appendRow(['Total Amount', 'APR', 'Term', 'Category', 'Monthly Payment', 'Total Interest', 'Remaining Balance']);
        break;
      case 'Payments':
        sheet.appendRow(['Loan Index', 'Date', 'Amount', 'Principal', 'Interest', 'Remaining Balance']);
        break;
      default:
        // If it's a sheet we haven't explicitly defined, just create it without headers
        Logger.log('Created sheet without predefined headers: ' + sheetName);
    }
  }
  return sheet;
}

function addExpense(date, amount, description, category) {
  Logger.log('Adding expense: ' + JSON.stringify({ date, amount, description, category }));
  try {
    var sheet = getOrCreateSheet('Expenses');
    // Format date as YYYY-MM-DD
    var formattedDate = Utilities.formatDate(new Date(date), Session.getScriptTimeZone(), "yyyy-MM-dd");
    sheet.appendRow([formattedDate, Number(amount), description, category]);
    var expenses = getExpenses();
    Logger.log('Returning expenses after addition: ' + JSON.stringify(expenses));
    return expenses;
  } catch (error) {
    Logger.log('Error in addExpense: ' + error.toString());
    return getExpenses();
  }
}


function getExpenses() {
  Logger.log('Getting expenses');
  try {
    var sheet = getOrCreateSheet('Expenses');
    var data = sheet.getDataRange().getValues();
    var formattedData = data.map(function (row, index) {
      if (index === 0) return row; // Header row
      return [
        Utilities.formatDate(new Date(row[0]), Session.getScriptTimeZone(), "yyyy-MM-dd"),
        Number(row[1]),
        String(row[2]),
        String(row[3])
      ];
    });
    Logger.log('Retrieved expenses: ' + JSON.stringify(formattedData));
    return formattedData.length > 1 ? formattedData : [['Date', 'Amount', 'Description', 'Category']];
  } catch (error) {
    Logger.log('Error in getExpenses: ' + error.toString());
    return [['Date', 'Amount', 'Description', 'Category']];
  }
}




function getCategories() {
  Logger.log("Fetching categories");
  var sheet = getOrCreateSheet('Categories');
  return sheet.getDataRange().getValues().slice(1).map(row => row[0]);  // Exclude header row
}

function addCategory(category) {
  Logger.log("Adding category: " + category);
  var sheet = getOrCreateSheet('Categories');
  sheet.appendRow([category]);
  return getCategories();
}

function removeExpense(index) {
  Logger.log("Removing expense at index: " + index);
  var sheet = getOrCreateSheet('Expenses');
  sheet.deleteRow(index + 2);  // +2 because index is 0-based and we have a header row
  return getExpenses();
}

function addLoan(totalAmount, apr, term, category, monthlyPayment) {
  Logger.log("Adding loan: " + JSON.stringify({ totalAmount, apr, term, category, monthlyPayment }));
  var sheet = getOrCreateSheet('Loans');
  var monthlyRate = apr / 12 / 100;
  var totalInterest = 0;
  var remainingBalance = totalAmount;

  if (term > 0) {
    // Calculate monthly payment if not provided or invalid
    if (!monthlyPayment || monthlyPayment <= 0) {
      monthlyPayment = (totalAmount * monthlyRate * Math.pow(1 + monthlyRate, term)) / (Math.pow(1 + monthlyRate, term) - 1);
    }

    // Calculate total interest
    totalInterest = (monthlyPayment * term) - totalAmount;
  } else {
    // For indefinite loans, we'll estimate total interest for a 30-year term
    term = "Indefinite";
    if (!monthlyPayment || monthlyPayment <= 0) {
      monthlyPayment = totalAmount * monthlyRate;
    }
    totalInterest = (monthlyPayment * 360) - totalAmount; // 360 months = 30 years
  }

  // Ensure total interest is not negative
  totalInterest = Math.max(0, totalInterest);

  sheet.appendRow([totalAmount, apr, term, category, monthlyPayment, totalInterest, remainingBalance]);
  return getLoans();
}


function getLoans() {
  Logger.log("Fetching loans");
  var sheet = getOrCreateSheet('Loans');
  return sheet.getDataRange().getValues();
}


function addPayment(loanIndex, paymentDate, paymentAmount) {
  try {
    Logger.log('Adding payment: ' + JSON.stringify({ loanIndex, paymentDate, paymentAmount }));

    var loansSheet = getOrCreateSheet('Loans');
    var loans = loansSheet.getDataRange().getValues();
    Logger.log('Loans data: ' + JSON.stringify(loans));

    if (loanIndex < 0 || loanIndex >= loans.length - 1) {
      throw new Error('Invalid loan index: ' + loanIndex);
    }

    var loan = loans[loanIndex + 1];  // +1 because of header row
    Logger.log('Selected loan: ' + JSON.stringify(loan));

    var totalAmount = loan[0];
    var apr = loan[1];
    var remainingBalance = loan[6];
    var monthlyRate = apr / 12 / 100;

    var interestPaid = remainingBalance * monthlyRate;
    var principalPaid = paymentAmount - interestPaid;
    var newBalance = remainingBalance - principalPaid;

    // Round the values to two decimal places
    interestPaid = Math.round(interestPaid * 100) / 100;
    principalPaid = Math.round(principalPaid * 100) / 100;
    newBalance = Math.round(newBalance * 100) / 100;

    Logger.log('Calculated payment details: ' + JSON.stringify({
      interestPaid: interestPaid,
      principalPaid: principalPaid,
      newBalance: newBalance
    }));

    var paymentsSheet = getOrCreateSheet('Payments');
    paymentsSheet.appendRow([loanIndex, paymentDate, paymentAmount, principalPaid, interestPaid, newBalance]);

    // Update loan remaining balance
    loansSheet.getRange(loanIndex + 2, 7).setValue(newBalance);  // +2 because of header row and 0-based index

    var result = {
      loans: getLoans(),
      payments: getPayments()
    };
    Logger.log('Returning payment result: ' + JSON.stringify(result));
    return result;
  } catch (error) {
    Logger.log('Error in addPayment: ' + error.toString());
    Logger.log('Error stack: ' + error.stack);
    return {
      error: error.message,
      loans: getLoans(),
      payments: getPayments()
    };
  }
}


function getPayments() {
  try {
    Logger.log('Getting payments');
    var sheet = getOrCreateSheet('Payments');
    var data = sheet.getDataRange().getValues();

    // Ensure we always return at least the header row
    if (data.length === 0) {
      data = [['Loan Index', 'Date', 'Amount', 'Principal', 'Interest', 'Remaining Balance']];
    }

    // Format the date and ensure all numeric values are numbers
    var formattedData = data.map(function (row, index) {
      if (index === 0) return row; // Skip header row
      return [
        Number(row[0]),
        Utilities.formatDate(new Date(row[1]), Session.getScriptTimeZone(), "yyyy-MM-dd"),
        Number(row[2]),
        Number(row[3]),
        Number(row[4]),
        Number(row[5])
      ];
    });

    Logger.log('Retrieved payments: ' + JSON.stringify(formattedData));
    return formattedData;
  } catch (error) {
    Logger.log('Error in getPayments: ' + error.toString());
    Logger.log('Error stack: ' + error.stack);
    // Return at least the header row in case of an error
    return [['Loan Index', 'Date', 'Amount', 'Principal', 'Interest', 'Remaining Balance']];
  }
}

function updateLoan(index, totalAmount, apr, term, category, monthlyPayment) {
  Logger.log("Updating loan: " + JSON.stringify({ index, totalAmount, apr, term, category, monthlyPayment }));
  var sheet = getOrCreateSheet('Loans');
  var monthlyRate = apr / 12 / 100;
  var totalInterest = 0;
  var remainingBalance = totalAmount;

  if (term > 0) {
    // Calculate monthly payment if not provided or invalid
    if (!monthlyPayment || monthlyPayment <= 0) {
      monthlyPayment = (totalAmount * monthlyRate * Math.pow(1 + monthlyRate, term)) / (Math.pow(1 + monthlyRate, term) - 1);
    }

    // Calculate total interest
    totalInterest = (monthlyPayment * term) - totalAmount;
  } else {
    // For indefinite loans, we'll estimate total interest for a 30-year term
    term = "Indefinite";
    if (!monthlyPayment || monthlyPayment <= 0) {
      monthlyPayment = totalAmount * monthlyRate;
    }
    totalInterest = (monthlyPayment * 360) - totalAmount; // 360 months = 30 years
  }

  // Ensure total interest is not negative
  totalInterest = Math.max(0, totalInterest);

  var rowToUpdate = index + 2; // +2 because of header row and 0-based index
  sheet.getRange(rowToUpdate, 1, 1, 7).setValues([[totalAmount, apr, term, category, monthlyPayment, totalInterest, remainingBalance]]);
  return getLoans();
}


function addPayment(loanIndex, paymentDate, paymentAmount) {
  try {
    Logger.log('Adding payment: ' + JSON.stringify({ loanIndex, paymentDate, paymentAmount }));

    var loansSheet = getOrCreateSheet('Loans');
    var loans = loansSheet.getDataRange().getValues();
    Logger.log('Loans data: ' + JSON.stringify(loans));

    if (loanIndex < 0 || loanIndex >= loans.length - 1) {
      throw new Error('Invalid loan index: ' + loanIndex);
    }

    var loan = loans[loanIndex + 1];  // +1 because of header row
    Logger.log('Selected loan: ' + JSON.stringify(loan));

    var totalAmount = loan[0];
    var apr = loan[1];
    var remainingBalance = loan[6];
    var monthlyRate = apr / 12 / 100;

    var interestPaid = remainingBalance * monthlyRate;
    var principalPaid = paymentAmount - interestPaid;
    var newBalance = remainingBalance - principalPaid;

    // Round the values to two decimal places
    interestPaid = Math.round(interestPaid * 100) / 100;
    principalPaid = Math.round(principalPaid * 100) / 100;
    newBalance = Math.round(newBalance * 100) / 100;

    Logger.log('Calculated payment details: ' + JSON.stringify({
      interestPaid: interestPaid,
      principalPaid: principalPaid,
      newBalance: newBalance
    }));

    var paymentsSheet = getOrCreateSheet('Payments');
    paymentsSheet.appendRow([loanIndex, paymentDate, paymentAmount, principalPaid, interestPaid, newBalance]);

    // Update loan remaining balance
    loansSheet.getRange(loanIndex + 2, 7).setValue(newBalance);  // +2 because of header row and 0-based index

    return {
      loans: getLoans(),
      payments: getPayments()
    };
  } catch (error) {
    Logger.log('Error in addPayment: ' + error.toString());
    Logger.log('Error stack: ' + error.stack);
    return {
      error: error.message,
      loans: getLoans(),
      payments: getPayments()
    };
  }
}


function updatePayment(index, loanIndex, paymentDate, paymentAmount) {
  try {
    Logger.log('Updating payment: ' + JSON.stringify({ index, loanIndex, paymentDate, paymentAmount }));

    var paymentsSheet = getOrCreateSheet('Payments');
    var payments = paymentsSheet.getDataRange().getValues();

    if (index < 0 || index >= payments.length - 1) {
      throw new Error('Invalid payment index');
    }

    var loansSheet = getOrCreateSheet('Loans');
    var loans = loansSheet.getDataRange().getValues();

    if (loanIndex < 0 || loanIndex >= loans.length - 1) {
      throw new Error('Invalid loan index');
    }

    var loan = loans[loanIndex + 1]; // +1 because of header row
    var apr = loan[1];
    var monthlyPayment = loan[4];
    var remainingBalance = loan[6];
    var monthlyRate = apr / 12 / 100;

    // Calculate new payment details
    var interestPaid = remainingBalance * monthlyRate;
    var principalPaid = paymentAmount - interestPaid;
    var newBalance = remainingBalance - principalPaid;

    // Calculate payments left
    var paymentsLeft = Math.ceil(newBalance / monthlyPayment);

    // Round the values to two decimal places
    interestPaid = Math.round(interestPaid * 100) / 100;
    principalPaid = Math.round(principalPaid * 100) / 100;
    newBalance = Math.round(newBalance * 100) / 100;

    // Ensure paymentDate is a Date object
    var formattedPaymentDate = Utilities.formatDate(new Date(paymentDate), Session.getScriptTimeZone(), "yyyy-MM-dd");

    // Update the payment
    paymentsSheet.getRange(index + 2, 1, 1, 7).setValues([[loanIndex, formattedPaymentDate, paymentAmount, principalPaid, interestPaid, newBalance, paymentsLeft]]);

    // Update loan remaining balance
    loansSheet.getRange(loanIndex + 2, 7).setValue(newBalance);

    return {
      loans: getLoans(),
      payments: getPayments()
    };
  } catch (error) {
    Logger.log('Error in updatePayment: ' + error.toString());
    Logger.log('Error stack: ' + error.stack);
    return {
      error: error.message,
      loans: getLoans(),
      payments: getPayments()
    };
  }
}
function removeLoan(index) {
  Logger.log("Removing loan at index: " + index);
  var sheet = getOrCreateSheet('Loans');
  sheet.deleteRow(index + 2);  // +2 because index is 0-based and we have a header row
  return getLoans();
}

function removePayment(index) {
  try {
    Logger.log('Removing payment at index: ' + index);
    var paymentsSheet = getOrCreateSheet('Payments');
    var payments = paymentsSheet.getDataRange().getValues();

    if (index < 0 || index >= payments.length - 1) {
      throw new Error('Invalid payment index');
    }

    var paymentToRemove = payments[index + 1]; // +1 because of header row
    var loanIndex = paymentToRemove[0];
    var principalPaid = paymentToRemove[3];

    // Remove the payment
    paymentsSheet.deleteRow(index + 2); // +2 because of header and 0-based index

    // Update loan balance
    var loansSheet = getOrCreateSheet('Loans');
    var loans = loansSheet.getDataRange().getValues();
    var loan = loans[loanIndex + 1]; // +1 because of header row
    var currentBalance = loan[6];
    var newBalance = currentBalance + principalPaid;
    loansSheet.getRange(loanIndex + 2, 7).setValue(newBalance);

    return {
      loans: getLoans(),
      payments: getPayments()
    };
  } catch (error) {
    Logger.log('Error in removePayment: ' + error.toString());
    return {
      error: error.message,
      loans: getLoans(),
      payments: getPayments()
    };
  }
}

function initializeFromExistingSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var originSheet = ss.getSheetByName('Sheet1');

  if (originSheet) {
    var data = originSheet.getDataRange().getValues();
    var loansSheet = getOrCreateSheet('Loans');
    var paymentsSheet = getOrCreateSheet('Payments');

    // Assuming the first row is headers
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      // Adjust these indices based on your actual data structure
      var totalAmount = row[0];
      var apr = row[1];
      var term = row[2];
      var category = row[3];

      addLoan(totalAmount, apr, term, category);
    }

    Logger.log("Initialized loans from existing sheet");
  } else {
    Logger.log("No existing 'Sheet1' found for initialization");
  }
}
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
        sheet.appendRow(['Loan Index', 'Date', 'Amount', 'Principal', 'Interest', 'Remaining Balance', 'Payments Left']);
        break;
      default:
        // If it's a sheet we haven't explicitly defined, just create it without headers
        Logger.log('Created sheet without predefined headers: ' + sheetName);
    }
  }
  return sheet;
}



function getInitialData() {
  var loans = getLoans();
  var payments = getPayments();

  Logger.log('Fetched loans: ' + JSON.stringify(loans));
  Logger.log('Fetched payments: ' + JSON.stringify(payments));

  return {
    loans: loans,
    payments: payments
  };
}

function addExpense(date, amount, description, category) {
  Logger.log('Adding expense: ' + JSON.stringify({ date, amount, description, category }));
  try {
    var sheet = getOrCreateSheet('Expenses');
    var formattedDate = Utilities.formatDate(new Date(date), Session.getScriptTimeZone(), "yyyy-MM-dd");
    sheet.appendRow([formattedDate, Number(amount), description, category]);
    return getExpenseSummary();
  } catch (error) {
    Logger.log('Error in addExpense: ' + error.toString());
    return getExpenseSummary();
  }
}

function getExpenseSummary(searchTerm) {
  var expenses = getExpenses(searchTerm);
  var monthlyExpenses = {};
  var overallExpenses = {};

  expenses.forEach(function (expense, index) {
    if (index === 0) return; // Skip header row
    var date = new Date(expense[0]);
    var month = Utilities.formatDate(date, Session.getScriptTimeZone(), "MMMM yyyy");
    var amount = expense[1];
    var category = expense[3];

    if (!monthlyExpenses[month]) {
      monthlyExpenses[month] = {};
    }
    if (!monthlyExpenses[month][category]) {
      monthlyExpenses[month][category] = 0;
    }
    monthlyExpenses[month][category] += amount;

    if (!overallExpenses[category]) {
      overallExpenses[category] = 0;
    }
    overallExpenses[category] += amount;
  });

  return {
    expenses: expenses,
    monthlyExpenses: monthlyExpenses,
    overallExpenses: overallExpenses
  };
}


function getExpenses(searchTerm) {
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

    // Apply search filter if searchTerm is provided
    if (searchTerm) {
      searchTerm = searchTerm.toLowerCase();
      formattedData = formattedData.filter(function (row, index) {
        if (index === 0) return true; // Keep header row
        return row[0].toLowerCase().includes(searchTerm) ||
          row[2].toLowerCase().includes(searchTerm) ||
          row[3].toLowerCase().includes(searchTerm);
      });
    }

    // Sort by date, oldest first (skip header row)
    var header = formattedData.shift();
    formattedData.sort((a, b) => new Date(a[0]) - new Date(b[0]));
    formattedData.unshift(header);

    Logger.log('Retrieved expenses: ' + JSON.stringify(formattedData));
    return formattedData.length > 1 ? formattedData : [['Date', 'Amount', 'Description', 'Category']];
  } catch (error) {
    Logger.log('Error in getExpenses: ' + error.toString());
    return [['Date', 'Amount', 'Description', 'Category']];
  }
}


function removeExpense(index) {
  Logger.log("Removing expense at index: " + index);
  var sheet = getOrCreateSheet('Expenses');
  sheet.deleteRow(index + 2);  // +2 because index is 0-based and we have a header row
  return getExpenseSummary();
}

function searchExpenses(searchTerm) {
  return getExpenseSummary(searchTerm);
}

function resetExpenseSearch() {
  return getExpenseSummary();
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
  var sheet = getOrCreateSheet('Loans');
  var data = sheet.getDataRange().getValues();
  return data.length > 1 ? data : null;
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
    totalInterest = 0; // 360 months = 30 years
  }

  // Ensure total interest is not negative
  totalInterest = Math.max(0, totalInterest);

  // Round values
  monthlyPayment = Math.round(monthlyPayment * 100) / 100;
  totalInterest = Math.round(totalInterest * 100) / 100;

  var rowToUpdate = index + 2; // +2 because of header row and 0-based index
  sheet.getRange(rowToUpdate, 1, 1, 7).setValues([[totalAmount, apr, term, category, monthlyPayment, totalInterest, remainingBalance]]);
  return getLoans();
}


function getPaymentsAndLoans() {
  try {
    var payments = getPayments();
    var loans = getLoans();
    Logger.log('Payments: ' + JSON.stringify(payments));
    Logger.log('Loans: ' + JSON.stringify(loans));
    return {
      payments: payments,
      loans: loans
    };
  } catch (error) {
    Logger.log('Error in getPaymentsAndLoans: ' + error.toString());
    return {
      error: error.message,
      payments: [],
      loans: []
    };
  }
}


function addPayment(loanIndex, paymentDate, paymentAmount) {
  try {
    Logger.log('Adding payment: ' + JSON.stringify({ loanIndex, paymentDate, paymentAmount }));

    var loansSheet = getOrCreateSheet('Loans');
    var loans = loansSheet.getDataRange().getValues();

    if (loanIndex < 0 || loanIndex >= loans.length - 1) {
      throw new Error('Invalid loan index');
    }

    var loan = loans[loanIndex + 1];  // +1 because of header row
    var totalAmount = loan[0];
    var apr = loan[1];
    var monthlyPayment = loan[4];
    var totalInterest = loan[5];
    var remainingBalance = loan[6];
    var monthlyRate = apr / 12 / 100;

    var interestPaid = remainingBalance * monthlyRate;
    var principalPaid = paymentAmount - interestPaid;
    var newBalance = remainingBalance - principalPaid;

    // Update total interest
    totalInterest += interestPaid;

    // Calculate payments left
    var paymentsLeft = Math.ceil(newBalance / monthlyPayment);
    if (paymentsLeft <= 0 || !isFinite(paymentsLeft)) {
      paymentsLeft = 0;
    }

    // Round the values to two decimal places
    interestPaid = Math.round(interestPaid * 100) / 100;
    principalPaid = Math.round(principalPaid * 100) / 100;
    newBalance = Math.round(newBalance * 100) / 100;
    totalInterest = Math.round(totalInterest * 100) / 100;

    // Ensure paymentDate is a Date object and format it correctly
    var formattedPaymentDate = Utilities.formatDate(new Date(paymentDate), Session.getScriptTimeZone(), "yyyy-MM-dd");

    var paymentsSheet = getOrCreateSheet('Payments');
    paymentsSheet.appendRow([loanIndex, formattedPaymentDate, paymentAmount, principalPaid, interestPaid, newBalance, paymentsLeft]);

    // Update loan remaining balance and total interest
    loansSheet.getRange(loanIndex + 2, 6, 1, 2).setValues([[totalInterest, newBalance]]);

    return getInitialData();
  } catch (error) {
    Logger.log('Error in addPayment: ' + error.toString());
    Logger.log('Error stack: ' + error.stack);
    return {
      error: error.message,
      ...getInitialData()
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
      data = [['Loan Index', 'Date', 'Amount', 'Principal', 'Interest', 'Remaining Balance', 'Payments Left']];
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
        Number(row[5]),
        row[6] !== undefined ? Number(row[6]) : null
      ];
    });

    // Sort by date, oldest first (skip header row)
    var header = formattedData.shift();
    formattedData.sort((a, b) => new Date(a[1]) - new Date(b[1]));
    formattedData.unshift(header);

    Logger.log('Retrieved payments: ' + JSON.stringify(formattedData));
    return formattedData;
  } catch (error) {
    Logger.log('Error in getPayments: ' + error.toString());
    Logger.log('Error stack: ' + error.stack);
    // Return at least the header row in case of an error
    return [['Loan Index', 'Date', 'Amount', 'Principal', 'Interest', 'Remaining Balance', 'Payments Left']];
  }
}


function addPayment(loanIndex, paymentDate, paymentAmount) {
  try {
    Logger.log('Adding payment: ' + JSON.stringify({ loanIndex, paymentDate, paymentAmount }));

    var loansSheet = getOrCreateSheet('Loans');
    var loans = loansSheet.getDataRange().getValues();

    if (loanIndex < 0 || loanIndex >= loans.length - 1) {
      throw new Error('Invalid loan index');
    }

    var loan = loans[loanIndex + 1];  // +1 because of header row
    var totalAmount = loan[0];
    var apr = loan[1];
    var monthlyPayment = loan[4];
    var totalInterest = loan[5];
    var remainingBalance = loan[6];
    var monthlyRate = apr / 12 / 100;

    var interestPaid = remainingBalance * monthlyRate;
    var principalPaid = paymentAmount - interestPaid;
    var newBalance = remainingBalance - principalPaid;

    // Update total interest
    totalInterest += interestPaid;

    // Calculate payments left
    var paymentsLeft = Math.ceil(newBalance / monthlyPayment);
    if (paymentsLeft <= 0 || !isFinite(paymentsLeft)) {
      paymentsLeft = 'N/A';
    }

    // Round the values to two decimal places
    interestPaid = Math.round(interestPaid * 100) / 100;
    principalPaid = Math.round(principalPaid * 100) / 100;
    newBalance = Math.round(newBalance * 100) / 100;
    totalInterest = Math.round(totalInterest * 100) / 100;

    // Ensure paymentDate is a Date object and format it correctly
    var formattedPaymentDate = Utilities.formatDate(new Date(paymentDate), Session.getScriptTimeZone(), "yyyy-MM-dd");

    var paymentsSheet = getOrCreateSheet('Payments');
    paymentsSheet.appendRow([loanIndex, formattedPaymentDate, paymentAmount, principalPaid, interestPaid, newBalance, paymentsLeft]);

    // Update loan remaining balance and total interest
    loansSheet.getRange(loanIndex + 2, 6, 1, 2).setValues([[totalInterest, newBalance]]);

    return getInitialData();
  } catch (error) {
    Logger.log('Error in addPayment: ' + error.toString());
    Logger.log('Error stack: ' + error.stack);
    return {
      error: error.message,
      ...getInitialData()
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

    // Recalculate payment details
    var loansSheet = getOrCreateSheet('Loans');
    var loans = loansSheet.getDataRange().getValues();
    var loan = loans[loanIndex + 1];  // +1 because of header row
    var apr = loan[1];
    var monthlyPayment = loan[4];
    var remainingBalance = loan[6];
    var monthlyRate = apr / 12 / 100;

    var interestPaid = remainingBalance * monthlyRate;
    var principalPaid = paymentAmount - interestPaid;
    var newBalance = remainingBalance - principalPaid;

    // Calculate payments left
    var paymentsLeft = Math.ceil(newBalance / monthlyPayment);
    if (paymentsLeft <= 0 || !isFinite(paymentsLeft)) {
      paymentsLeft = 0;
    }

    // Round the values
    interestPaid = Math.round(interestPaid * 100) / 100;
    principalPaid = Math.round(principalPaid * 100) / 100;
    newBalance = Math.round(newBalance * 100) / 100;

    // Format the date
    var formattedPaymentDate = Utilities.formatDate(new Date(paymentDate), Session.getScriptTimeZone(), "yyyy-MM-dd");

    // Update the payment
    paymentsSheet.getRange(index + 2, 1, 1, 7).setValues([[loanIndex, formattedPaymentDate, paymentAmount, principalPaid, interestPaid, newBalance, paymentsLeft]]);

    // Update loan remaining balance
    loansSheet.getRange(loanIndex + 2, 7).setValue(newBalance);

    return getInitialData();
  } catch (error) {
    Logger.log('Error in updatePayment: ' + error.toString());
    Logger.log('Error stack: ' + error.stack);
    return {
      error: error.message,
      ...getInitialData()
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
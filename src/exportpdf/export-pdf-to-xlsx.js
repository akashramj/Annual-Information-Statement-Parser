// To Convert PDF to XLSX
const PDFServicesSdk = require("@adobe/pdfservices-node-sdk");

try {
  // Initial setup, create credentials instance.
  const credentials =
    PDFServicesSdk.Credentials.serviceAccountCredentialsBuilder()
      .fromFile("pdfservices-api-credentials.json")
      .build();

  //Create an ExecutionContext using credentials and create a new operation instance.
  const executionContext = PDFServicesSdk.ExecutionContext.create(credentials),
    exportPDF = PDFServicesSdk.ExportPDF,
    exportPDFOperation = exportPDF.Operation.createNew(
      //exportPDF.SupportedTargetFormats.DOCX
      exportPDF.SupportedTargetFormats.XLSX
    );

  // Set operation input from a source file
  const input = PDFServicesSdk.FileRef.createFromLocalFile(
    "src/exportpdf/AISsource.pdf"
  );
  exportPDFOperation.setInput(input);

  //Generating a file name
  let outputFilePath = createOutputFilePath();

  // Execute the operation and Save the result to the specified location.
  exportPDFOperation
    .execute(executionContext)
    .then((result) => result.saveAsFile(outputFilePath))
    .catch((err) => {
      if (
        err instanceof PDFServicesSdk.Error.ServiceApiError ||
        err instanceof PDFServicesSdk.Error.ServiceUsageError
      ) {
        console.log("Exception encountered while executing operation", err);
      } else {
        console.log("Exception encountered while executing operation", err);
      }
    });

  //Generates a string containing a directory structure and file name for the output file.
  function createOutputFilePath() {
    let date = new Date();
    let dateString =
      date.getFullYear() +
      "-" +
      ("0" + (date.getMonth() + 1)).slice(-2) +
      "-" +
      ("0" + date.getDate()).slice(-2) +
      "T" +
      ("0" + date.getHours()).slice(-2) +
      "-" +
      ("0" + date.getMinutes()).slice(-2) +
      "-" +
      ("0" + date.getSeconds()).slice(-2);
    return "output/ExportPDFToDOCX/AISForm" + dateString + ".xlsx";
  }
} catch (err) {
  console.log("Exception encountered while executing operation", err);
}

// To Extract data from generated XLSX file
const XLSX = require("xlsx");
const fs = require("fs");

const inputFile = "./excelSource2.xlsx";

// Read the XLSX file
const workbook = XLSX.readFile(inputFile);

// Assuming the first sheet is the target sheet
const sheetName = workbook.SheetNames[0];
const worksheet = workbook.Sheets[sheetName];

// Convert the worksheet data to JSON
const jsonData = XLSX.utils.sheet_to_json(worksheet);

// To return final result;
let finalResult = [];

// To identify the Format
let flag = -1;

// To store Parsed values
let aisData = {
  srNo: null,
  informationCode: null,
  informationDescription: null,
  informationSource: null,
  count: null,
  amount: null,
  category: null,
  tabeData: [],
};

// Iterating through JSON data
for (let i = 1; i < jsonData.length; i++) {
  // Format 1
  let format1 = {
    srNo: null,
    quarter: null,
    dateOfPayment: null,
    amountPaid: null,
    tdsDeducted: null,
    tdsDeposited: null,
    status: null,
  };

  // Format 2
  let format2 = {
    srNo: null,
    reportedOn: null,
    accountNumber: null,
    accountType: null,
    interestAmount: null,
    status: null,
  };

  // Format 3
  let format3 = {
    srNo: null,
    dateOfSale: null,
    securityName: null,
    securityClass: null,
    debitType: null,
    creditType: null,
    assetType: null,
    quantity: null,
    salePricePerUnit: null,
    salesConsideration: null,
    costOfAcquisition: null,
    unitFm: null,
  };

  // Format 4
  let format4 = {
    srNo: null,
    amcName: null,
    dateOfSale: null,
    securityClass: null,
    securityName: null,
    debitType: null,
    creditType: null,
    assetType: null,
    quantity: null,
    salePricePerUnit: null,
    salesConsideration: null,
    st: null,
  };

  // Format 5
  let format5 = {
    srNo: null,
    clientId: null,
    holderFlag: null,
    marketPurchase: null,
    marketSales: null,
    status: null,
  };

  // Format 6
  let format6 = {
    srNo: null,
    clientId: null,
    amcName: null,
    holderFlag: null,
    totalPurchaseAmount: null,
    totalSalesValue: null,
    status: null,
  };

  // Previous JSON length
  const previousJsonLength = Object.keys(jsonData[i - 1]).length;

  // Current JSON length
  const currentJsonLength = Object.keys(jsonData[i]).length;

  // Compute previous keys
  const previousKeys = Object.keys(jsonData[i - 1]);

  // Compute current keys
  const currentKeys = Object.keys(jsonData[i]);

  // Fetching Category and storing it
  if (
    jsonData[i][currentKeys[0]] == "Interest from savings bank" ||
    jsonData[i][currentKeys[0]] ==
      "Sale of securities and units of mutual fund" ||
    jsonData[i][currentKeys[0]] ==
      "Purchase of securities and units of mutual funds" ||
    jsonData[i][currentKeys[0]] == "Refund"
  ) {
    aisData.category = jsonData[i][currentKeys[0]];
  }

  // Storing main Section Table
  if (
    previousJsonLength == 6 &&
    jsonData[i - 1][previousKeys[0]] == "SR. NO." &&
    jsonData[i - 1][previousKeys[1]] == "INFORMATION CODE" &&
    jsonData[i - 1][previousKeys[2]] == "INFORMATION DESCRIPTION" &&
    jsonData[i - 1][previousKeys[3]] == "INFORMATION SOURCE" &&
    jsonData[i - 1][previousKeys[4]] == "COUNT" &&
    jsonData[i - 1][previousKeys[5]] == "AMOUNT"
  ) {
    // Make conditional check for previous JSON data
    // Push object to final result
    finalResult.push(Object.assign({}, aisData));

    // Set default null values
    (aisData.srNo = null),
      (aisData.informationCode = null),
      (aisData.informationDescription = null),
      (aisData.informationSource = null),
      (aisData.count = null),
      (aisData.amount = null),
      (aisData.tabeData = []);
    flag = 0;
  }

  if (flag == 0) {
    // Push current JSON data
    aisData.srNo = jsonData[i][currentKeys[0]];
    aisData.informationCode = jsonData[i][currentKeys[1]];
    aisData.informationDescription = jsonData[i][currentKeys[2]];
    aisData.informationSource = jsonData[i][currentKeys[3]];
    aisData.count = jsonData[i][currentKeys[4]];
    aisData.amount = jsonData[i][currentKeys[5]];

    flag = -1;
  }

  // Format 1
  if (
    previousJsonLength == 7 &&
    jsonData[i - 1][previousKeys[0]] == "SR. NO." &&
    jsonData[i - 1][previousKeys[1]] == "QUARTER" &&
    jsonData[i - 1][previousKeys[2]] == "DATE OF PAYMENT/CREDIT" &&
    jsonData[i - 1][previousKeys[3]] == "AMOUNT PAID/CREDITED" &&
    jsonData[i - 1][previousKeys[4]] == "TDS DEDUCTED" &&
    jsonData[i - 1][previousKeys[5]] == "TDS DEPOSITED" &&
    jsonData[i - 1][previousKeys[6]] == "STATUS"
  ) {
    flag = 1;
  }

  if (flag == 1 && currentJsonLength == 7) {
    format1.srNo = jsonData[i][currentKeys[0]];
    format1.quarter = jsonData[i][currentKeys[1]];
    format1.dateOfPayment = jsonData[i][currentKeys[2]];
    format1.amountPaid = jsonData[i][currentKeys[3]];
    format1.tdsDeducted = jsonData[i][currentKeys[4]];
    format1.tdsDeposited = jsonData[i][currentKeys[5]];
    format1.status = jsonData[i][currentKeys[6]];

    // Store row values
    aisData.tabeData.push(Object.assign({}, format1));
  }

  // Format 2
  if (
    previousJsonLength == 6 &&
    jsonData[i - 1][previousKeys[0]] == "SR. NO." &&
    jsonData[i - 1][previousKeys[1]] == "REPORTED ON" &&
    jsonData[i - 1][previousKeys[2]] == "ACCOUNT NUMBER" &&
    jsonData[i - 1][previousKeys[3]] == "ACCOUNT TYPE" &&
    jsonData[i - 1][previousKeys[4]] == "INTEREST AMOUNT" &&
    jsonData[i - 1][previousKeys[5]] == "STATUS"
  ) {
    flag = 2;
  }

  if (flag == 2 && currentJsonLength == 6) {
    format2.srNo = jsonData[i][currentKeys[0]];
    format2.reportedOn = jsonData[i][currentKeys[1]];
    format2.accountNumber = jsonData[i][currentKeys[2]];
    format2.accountType = jsonData[i][currentKeys[3]];
    format2.interestAmount = jsonData[i][currentKeys[4]];
    format2.status = jsonData[i][currentKeys[5]];

    // Store row values
    aisData.tabeData.push(Object.assign({}, format2));
  }

  // Format 3
  if (
    previousJsonLength == 12 &&
    jsonData[i - 1][previousKeys[0]] == "SR.\r\nNO." &&
    jsonData[i - 1][previousKeys[1]] == "DATE OF SALE/ TRANSFER" &&
    jsonData[i - 1][previousKeys[2]] == "SECURITY NAME (SECURITY CODE)" &&
    jsonData[i - 1][previousKeys[3]] == "SECURITY CLASS" &&
    jsonData[i - 1][previousKeys[4]] == "DEBIT TYPE" &&
    jsonData[i - 1][previousKeys[5]] == "CREDIT TYPE" &&
    jsonData[i - 1][previousKeys[6]] == "ASSET TYPE" &&
    jsonData[i - 1][previousKeys[7]] == "QUANTITY" &&
    jsonData[i - 1][previousKeys[8]] == "SALE PRICE PER\r\nUNIT" &&
    jsonData[i - 1][previousKeys[9]] == "SALES CONSIDER\r\nATION" &&
    jsonData[i - 1][previousKeys[10]] == "COST OF ACQUISITIO\r\nN" &&
    jsonData[i - 1][previousKeys[11]] == "UNIT FM"
  ) {
    flag = 3;
  }

  if (flag == 3 && currentJsonLength == 12) {
    format3.srNo = jsonData[i][currentKeys[0]];
    format3.dateOfSale = jsonData[i][currentKeys[1]];
    format3.securityName = jsonData[i][currentKeys[2]];
    format3.securityClass = jsonData[i][currentKeys[3]];
    format3.debitType = jsonData[i][currentKeys[4]];
    format3.creditType = jsonData[i][currentKeys[5]];
    format3.assetType = jsonData[i][currentKeys[6]];
    format3.quantity = jsonData[i][currentKeys[7]];
    format3.salePricePerUnit = jsonData[i][currentKeys[8]];
    format3.salesConsideration = jsonData[i][currentKeys[9]];
    format3.costOfAcquisition = jsonData[i][currentKeys[10]];
    format3.unitFm = jsonData[i][currentKeys[11]];

    // Store row values
    aisData.tabeData.push(Object.assign({}, format3));
  }

  // Format 4
  if (
    previousJsonLength == 12 &&
    jsonData[i - 1][previousKeys[0]] == "SR.\r\nNO." &&
    jsonData[i - 1][previousKeys[1]] == "AMC NAME (CODE)" &&
    jsonData[i - 1][previousKeys[2]] == "DATE OF SALE/ TRANSFER" &&
    jsonData[i - 1][previousKeys[3]] == "SECURITY CLASS" &&
    jsonData[i - 1][previousKeys[4]] == "SECURITY NAME (SECURITY CODE)" &&
    jsonData[i - 1][previousKeys[5]] == "DEBIT TYPE" &&
    jsonData[i - 1][previousKeys[6]] == "CREDIT TYPE" &&
    jsonData[i - 1][previousKeys[7]] == "ASSET TYPE" &&
    jsonData[i - 1][previousKeys[8]] == "QUANTITY" &&
    jsonData[i - 1][previousKeys[9]] == "SALE PRICE PER\r\nUNIT" &&
    jsonData[i - 1][previousKeys[10]] == "SALES CONSIDERA\r\nTION" &&
    jsonData[i - 1][previousKeys[11]] == "ST"
  ) {
    flag = 4;
  }

  if (flag == 4 && currentJsonLength == 12) {
    format4.srNo = jsonData[i][currentKeys[0]];
    format4.amcName = jsonData[i][currentKeys[1]];
    format4.dateOfSale = jsonData[i][currentKeys[2]];
    format4.securityClass = jsonData[i][currentKeys[3]];
    format4.securityName = jsonData[i][currentKeys[4]];
    format4.debitType = jsonData[i][currentKeys[5]];
    format4.creditType = jsonData[i][currentKeys[6]];
    format4.assetType = jsonData[i][currentKeys[7]];
    format4.quantity = jsonData[i][currentKeys[8]];
    format4.salePricePerUnit = jsonData[i][currentKeys[9]];
    format4.salesConsideration = jsonData[i][currentKeys[10]];
    format4.st = jsonData[i][currentKeys[11]];

    // Store row values
    aisData.tabeData.push(Object.assign({}, format4));
  }

  // Format 5
  if (
    previousJsonLength == 6 &&
    jsonData[i - 1][previousKeys[0]] == "SR. NO." &&
    jsonData[i - 1][previousKeys[1]] == "CLIENT ID" &&
    jsonData[i - 1][previousKeys[2]] == "HOLDER FLAG" &&
    jsonData[i - 1][previousKeys[3]] == "MARKET PURCHASE" &&
    jsonData[i - 1][previousKeys[4]] == "MARKET SALES" &&
    jsonData[i - 1][previousKeys[5]] == "STATUS"
  ) {
    flag = 5;
  }

  if (flag == 5 && currentJsonLength == 6) {
    format5.srNo = jsonData[i][currentKeys[0]];
    format5.clientId = jsonData[i][currentKeys[1]];
    format5.holderFlag = jsonData[i][currentKeys[2]];
    format5.marketPurchase = jsonData[i][currentKeys[3]];
    format5.marketSales = jsonData[i][currentKeys[4]];
    format5.status = jsonData[i][currentKeys[5]];

    // Store row values
    aisData.tabeData.push(Object.assign({}, format5));
  }

  // Format 6
  if (
    previousJsonLength == 7 &&
    jsonData[i - 1][previousKeys[0]] == "SR. NO." &&
    jsonData[i - 1][previousKeys[1]] == "CLIENT ID" &&
    jsonData[i - 1][previousKeys[2]] == "AMC NAME (CODE)" &&
    jsonData[i - 1][previousKeys[3]] == "HOLDER FLAG" &&
    jsonData[i - 1][previousKeys[4]] == "TOTAL PURCHASE AMOUNT" &&
    jsonData[i - 1][previousKeys[5]] == "TOTAL SALES VALUE" &&
    jsonData[i - 1][previousKeys[6]] == "STATUS"
  ) {
    flag = 6;
  }

  if (flag == 6 && currentJsonLength == 7) {
    format6.srNo = jsonData[i][currentKeys[0]];
    format6.clientId = jsonData[i][currentKeys[1]];
    format6.amcName = jsonData[i][currentKeys[2]];
    format6.holderFlag = jsonData[i][currentKeys[3]];
    format6.totalPurchaseAmount = jsonData[i][currentKeys[4]];
    format6.totalSalesValue = jsonData[i][currentKeys[5]];
    format6.status = jsonData[i][currentKeys[6]];

    // Store row values
    aisData.tabeData.push(Object.assign({}, format6));
  }

  if (aisData.tabeData.length > 0) {
    console.log(aisData.tabeData);
  }
}

const outputJSON = "output.json";

// Save the JSON data to a file
fs.writeFileSync(outputJSON, JSON.stringify(jsonData, null, 2));

// Output final result
//console.log(finalResult);

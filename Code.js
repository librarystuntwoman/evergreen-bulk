// date/time function to ensure correct formatting
function safeDate(value) {
  if (!value) return null;
  const d = (value instanceof Date) ? value : new Date(value);
  return isNaN(d.getTime()) ? null : d;
}

// main function to generate .xacts file
function generateXacts() {

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const data = sheet.getDataRange().getValues();

  data.shift(); // remove headers

  const records = [];

  data.forEach(row => {

    const type = (row[0] || "").toString().toLowerCase();

    const patronBarcode = String(row[2] || "");
    const itemBarcode = String(row[3] || "");
    const noncatType = String(row[4] || "");

    const checkoutTime = safeDate(row[5]);
    const dueDate = safeDate(row[6]);
    const checkinTime = safeDate(row[7]);

    // checkouts
    if (type === "checkout") {

      if (!checkoutTime) {
        Logger.log("Skipping checkout row: " + JSON.stringify(row));
        return;
      }

      records.push(JSON.stringify({
        timestamp: Math.floor(checkoutTime.getTime() / 1000),
        type: "checkout",
        delta: 0,
        noncat_type: noncatType,
        patron_barcode: patronBarcode,
        barcode: itemBarcode,
        due_date: dueDate ? dueDate.toISOString() : "",
        checkout_time: checkoutTime.toISOString()
      }));

      return;
    }

    // check-ins
    if (type === "checkin") {

      if (!checkinTime) {
        Logger.log("Skipping checkin row: " + JSON.stringify(row));
        return;
      }

      records.push(JSON.stringify({
        timestamp: Math.floor(checkinTime.getTime() / 1000),
        type: "checkin",
        delta: 0,
        noncat_type: "",
        patron_barcode: patronBarcode,
        barcode: itemBarcode,
        due_date: dueDate ? dueDate.toISOString() : "",
        checkout_time: checkoutTime ? checkoutTime.toISOString() : "",
        checkin_time: checkinTime.toISOString()
      }));

      return;
    }

    // logger to prevent an invalid row from ruining everything
    Logger.log("Unknown row type: " + JSON.stringify(row));
  });

  const content = records.join("\n");

  Logger.log("TOTAL RECORDS: " + records.length);

  // same file name each time, but generated file can be renamed as needed
  const file = DriveApp.createFile(
    "offline_checkouts.xacts",
    content,
    MimeType.PLAIN_TEXT
  );

  SpreadsheetApp.getUi().alert(
    "Created file: " + file.getName() +
    "\n\nView in Google Drive:\n" +
    file.getUrl()
  );
}

// button to generate the .xacts file!
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("Evergreen")
    .addItem("Generate .xacts File", "generateXacts")
    .addToUi();
}
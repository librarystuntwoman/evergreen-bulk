// helper to skip invalid dates that would prevent the script from running
function safeDate(value) {
  if (!value) return null;

  const d = (value instanceof Date) ? value : new Date(value);

  return isNaN(d.getTime()) ? null : d;
}

function generateXacts() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const data = sheet.getDataRange().getValues();

  // Remove header row
  data.shift();

  const records = [];

  data.forEach(row => {
    // Column mapping (to match Google sheet):
    // 0 (A) Type
    // 1 (B) Timestamp
    // 2 (C) Patron Barcode
    // 3 (D) Item Barcode
    // 4 (E) Non-cataloged Type
    // 5 (F) Checkout Date
    // 6 (G) Due Date
    // 7 (H) Checkin Date
    // 8 (I) First Name
    // 9 (J) Last Name

    const type = (row[0] || "").toString().toLowerCase();

    const patronBarcode = String(row[2] || "");
    const itemBarcode = String(row[3] || "");
    const noncatType = String(row[4] || "");

    const checkoutTime = safeDate(row[5]);
    const dueDate = safeDate(row[6]);
    const checkinTime = safeDate(row[7]);

    // Logger.log(JSON.stringify(row));

    // checkout function with safeguard to skip rows with missing or invalid dates; otherwise script will fail if it encounters an error
    if (type === "checkout") {

    if (!checkoutTime) {
      Logger.log("Skipping bad checkout row: " + JSON.stringify(row));
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
    // check in function
    if (type === "checkin") {

      if (!checkinTime) {
        Logger.log("Skipping bad checkin row: " + JSON.stringify(row));
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

    // just in case there's something odd in the type column
    Logger.log("Unknown type row skipped: " + JSON.stringify(row));
    });
  const content = records.join("\n");

  // Logger.log("TOTAL RECORDS: " + records.length);
  const file = DriveApp.createFile(
    "offline_checkouts.xacts", // generic file name, but can be updated as needed
    content,
    MimeType.PLAIN_TEXT
  );

  SpreadsheetApp.getUi().alert(
    "Created file: " + file.getName() +
    "\n\nView in Google Drive:\n" +
    file.getUrl()
  );
}

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("Evergreen") // button in the Google sheet to start .xacts file generation
    .addItem("Generate .xacts File", "generateXacts")
    .addToUi();
}
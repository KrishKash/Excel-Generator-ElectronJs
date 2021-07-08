function createExcel() {
  var ExcelJS = require("exceljs");

  //A new Excel Work Book
  var workbook = new ExcelJS.Workbook();

  // Some information about the Excel Work Book.
  workbook.creator = "Krishna Kashyap";
  workbook.lastModifiedBy = "";
  workbook.created = new Date(2021, 7, 8);
  workbook.modified = new Date();
  workbook.lastPrinted = new Date(2021, 7, 9);
  console.log(workbook);

  // Create a sheet
  var sheet = workbook.addWorksheet("DeviceType");

  //A table header
  sheet.columns = [
    { header: "DeviceTypeName", key: "devicetypename" },
    { header: "Manufacturer", key: "manufacturer" },
    { header: "ModelNumber", key: "modelnumber" },
    { header: "DeviceKind", key: "devicekind" },
  ];

  // Add rows in the above header
  sheet.addRow({
    devicetypename: "PIR-Motion-Sensor",
    manufacturer: "Panasonic",
    modelnumber: "Panasonic EKMC1603111",
    devicekind: "Infrared-Motion-Detector",
  });

  sheet.addRow({
    devicetypename: "Edimax-AI-2003W",
    manufacturer: "Edimax",
    modelnumber: "AI-2003W",
    devicekind: "Multi-Device",
  });

  console.log(sheet);

  // Save Excel on Hard Disk
  workbook.xlsx.writeFile("Spreadsheet.xlsx").then(function () {
    // Success Message
    alert("File Saved");
  });
}

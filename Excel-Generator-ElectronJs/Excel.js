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

  {
    // Create DeviceTypeSheet
    var DeviceTypeSheet = workbook.addWorksheet("DeviceType");

    // table header
    DeviceTypeSheet.columns = [
      { header: "DeviceTypeName", key: "devicetypename", width: 30 },
      { header: "Manufacturer", key: "manufacturer", width: 30 },
      { header: "ModelNumber", key: "modelnumber", width: 30 },
      { header: "DeviceKind", key: "devicekind", width: 30 },
    ];

    // Add rows in the above header
    DeviceTypeSheet.addRow({
      devicetypename: "PIR-Motion-Sensor",
      manufacturer: "Panasonic",
      modelnumber: "Panasonic EKMC1603111",
      devicekind: "Infrared-Motion-Detector",
    });

    DeviceTypeSheet.addRow({
      devicetypename: "Edimax-AI-2003W",
      manufacturer: "Edimax",
      modelnumber: "AI-2003W",
      devicekind: "Multi-Device",
    });

    //background color for column header
    DeviceTypeSheet.eachRow(function (row, rowNumber) {
      row.eachCell((cell) => {
        if (rowNumber == 1) {
          cell.fill = {
            type: "pattern",
            pattern: "solid",
            fgColor: { argb: "5E8CFB" },
          };
        }
      });
    });

    console.log(DeviceTypeSheet);
  }

  {
    // Create DeviceInterfaceSheet
    var DeviceInterfaceSheet = workbook.addWorksheet("DeviceInterface");

    // table header
    DeviceInterfaceSheet.columns = [
      { header: "DeviceName", key: "devicename", width: 30 },
      { header: "DeviceTypeName", key: "devicetypename", width: 30 },
      { header: "LocationPoint", key: "locationpoint", width: 30 },
      { header: "IoTId", key: "iotid", width: 20 },
      { header: "CategoryName", key: "categoryname", width: 20 },
    ];

    // Add rows in the above header
    DeviceInterfaceSheet.addRow({
      devicename: "MS1",
      devicetypename: "PanasPIR-Motion-Sensoronic",
      locationpoint: "{'x': 1.13, 'y': 0.7, 'z': 2.75}",
      iotid: "35670",
      categoryname: "Category-1",
    });

    DeviceInterfaceSheet.addRow({
      devicename: "M2",
      devicetypename: "Edimax-AI-2003W",
      locationpoint: "{'x': 10.13, 'y': 1.7, 'z': 1.75}",
      iotid: "29898",
      categoryname: "Category-2",
    });

    //background color for column header
    DeviceInterfaceSheet.eachRow(function (row, rowNumber) {
      row.eachCell((cell) => {
        if (rowNumber == 1) {
          cell.fill = {
            type: "pattern",
            pattern: "solid",
            fgColor: { argb: "5E8CFB" },
          };
        }
      });
    });
    console.log(DeviceInterfaceSheet);
  }

  {
    // Create TelemetryPointSheet
    var TelemetryPointSheet = workbook.addWorksheet("TelemetryPoint");

    // table header
    TelemetryPointSheet.columns = [
      { header: "ObservationName", key: "observationname", width: 30 },
      { header: "Phenomenon", key: "phenomenon", width: 30 },
      { header: "DeviceName", key: "devicename", width: 30 },
      { header: "ObservedElement", key: "observedelement", width: 40 },
      { header: "IoTId", key: "iotid", width: 20 },
    ];

    // Add rows in the above header
    TelemetryPointSheet.addRow({
      observationname: "MS01",
      phenomenon: "Motion-Detection",
      devicename: "MS1",
      observedelement: '{"ECClassId": "bis.spatialelement", "UserLabel": "Door"}',
      iotid: "35770",
    });

    TelemetryPointSheet.addRow({
      observationname: "AQ01",
      phenomenon: "Air Quality",
      devicename: "M2",
      observedelement: '{"ECClassId": "bis.spatialelement", "ECInstanceId": "0xfa0"}',
      iotid: "29898",
    });

    //background color for column header
    TelemetryPointSheet.eachRow(function (row, rowNumber) {
      row.eachCell((cell) => {
        if (rowNumber == 1) {
          cell.fill = {
            type: "pattern",
            pattern: "solid",
            fgColor: { argb: "5E8CFB" },
          };
        }
      });
    });
    console.log(TelemetryPointSheet);
  }

  {
    // Create LocatorsSheet
    var LocatorsSheet = workbook.addWorksheet("Locators");

    // table header
    LocatorsSheet.columns = [
      { header: "LocatorName", key: "locatorname", width: 20 },
      { header: "ClassName", key: "classname", width: 40 },
      { header: "PropertyName", key: "propertyname", width: 20 },
    ];

    // Add rows in the above header
    LocatorsSheet.addRow({
      locatorname: "ById",
      classname: "bis.spatialelement",
      propertyname: "ECInstanceId",
    });

    //background color for column header
    LocatorsSheet.eachRow(function (row, rowNumber) {
      row.eachCell((cell) => {
        if (rowNumber == 1) {
          cell.fill = {
            type: "pattern",
            pattern: "solid",
            fgColor: { argb: "5E8CFB" },
          };
        }
      });
    });
    console.log(LocatorsSheet);
  }

  // Save Excel on Hard Disk
  workbook.xlsx.writeFile("Spreadsheet.xlsx").then(function () {
    // Success Message
    alert("File Saved");
  });
}

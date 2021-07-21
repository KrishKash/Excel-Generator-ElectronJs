function createExcel() {
  var ExcelJS = require("exceljs");
  const fs = require("fs");

  //A new Excel Work Book
  var workbook = new ExcelJS.Workbook();

  var fileName = document.getElementById("ssname").value;
  if (!fileName) {
    alert("Spreadsheet name can not be empty!");
    return;
  }
  fileName += ".xlsx";
  if (fs.existsSync(fileName)) {
    //disable createExcel button after spreadsheet get created
    document.getElementById("createExcel").disabled = true;
    document.getElementById("createExcel").innerText = "Spreadsheet Created";

    //enable sheet names dropdown
    document.getElementById("sheetnames").style.display = "block";
    document.getElementById("saveData").disabled = false;

    alert(fileName + " already exists!");
    return;
  } else {
    // Some information about the Excel Work Book.
    workbook.creator = "Krishna Kashyap";
    workbook.lastModifiedBy = "";
    workbook.created = new Date(2021, 7, 8);
    workbook.modified = new Date();
    workbook.lastPrinted = new Date(2021, 7, 9);

    // var Sheets= ["DeviceType", "DeviceInterface", "TelemetryPoint", "Locators"];

    // var ColumnHeader = {
    //   DeviceType: ["DeviceTypeName", "Manufacturer", "ModelNumber", "DeviceKind"],
    //   DeviceInterface: ["DeviceName", "DeviceTypeName", "LocationPoint", "IoTId", "CategoryName"],
    //   TelemetryPoint: ["A", "Aa", "Aaa", "Aaaa", "Aaaa"],
    //   Locators: ["B", "Bb", "Bbb", "Bbbb", "Bbbb"]
    // }

    // Sheets.forEach(sheet=>{
    //   var workSheet= workbook.addWorksheet(sheet);
    // });

    // Create DeviceTypeSheet
    var DeviceTypeSheet = workbook.addWorksheet("DeviceType");

    // table header
    DeviceTypeSheet.columns = [
      { header: "DeviceTypeName", key: "devicetypename", width: 30 },
      { header: "Manufacturer", key: "manufacturer", width: 30 },
      { header: "ModelNumber", key: "modelnumber", width: 30 },
      { header: "DeviceKind", key: "devicekind", width: 30 },
    ];

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

    // Create LocatorsSheet
    var LocatorsSheet = workbook.addWorksheet("Locators");

    // table header
    LocatorsSheet.columns = [
      { header: "LocatorName", key: "locatorname", width: 20 },
      { header: "ClassName", key: "classname", width: 40 },
      { header: "PropertyName", key: "propertyname", width: 20 },
    ];

    workbook.eachSheet(function (worksheet) {
      worksheet.eachRow(function (row, rowNumber) {
        row.eachCell((cell) => {
          if (rowNumber == 1) {
            // First set the background of header row
            cell.fill = {
              type: "pattern",
              pattern: "solid",
              fgColor: { argb: "f5b914" },
            };
          }
        });
      });
    });

    // Save Excel on Hard Disk
    workbook.xlsx.writeFile(fileName).then(function () {
      // Success Message
      alert("File Created");
    });
  }
  //disable button after spreadsheet get created
  document.getElementById("createExcel").disabled = true;
  document.getElementById("createExcel").innerText = "Spreadsheet Created";

  //enable sheet names dropdown
  document.getElementById("sheetnames").style.display = "block";
  document.getElementById("saveData").disabled = false;
}

function saveFormData() {
  var ExcelJS = require("exceljs");

  // Excel Work Book
  var workbook = new ExcelJS.Workbook();
  var fileName = document.getElementById("ssname").value;
  fileName += ".xlsx";

  workbook.xlsx.readFile(fileName).then(function () {
    var sheetName = document.getElementById("sname").value;
    var workSheet = workbook.getWorksheet(sheetName);
    var lastRow = workSheet.lastRow.number;

    switch (workSheet.name) {
      case "DeviceType":
        // workSheet.addRow({
        //   DeviceTypeName: "PIR-Motion-Sensor- Device",
        //   Manufacturer: "Panasonic",
        //   ModelNumber: "Panasonic EKMC1603111",
        //   DeviceKind: "Infrared-Motion-Detector",
        // });
        var DeviceTypeName = document.getElementById("devicetypename").value;
        var Manufacturer = document.getElementById("manufacturer").value;
        var ModelNumber = document.getElementById("modelnumber").value;
        var DeviceKind = document.getElementById("devicekind").value;

        workSheet.insertRow(++lastRow, [
          DeviceTypeName,
          Manufacturer,
          ModelNumber,
          DeviceKind,
        ]);

        // workSheet.insertRow(lastRow, ["PIR-Motion-Sensor- Device3", "Panasonic", "Panasonic EKMC1603111", "Infrared-Motion-Detector"]);
        // console.log(JSON.stringify(workSheet.getSheetValues()));
        break;

      case "DeviceInterface":
        var DeviceName = document.getElementById("devicename").value;
        var DeviceTypeName = document.getElementById("devicetypename").value;
        var LocationPoint = document.getElementById("locationpoint").value;
        var IoTId = document.getElementById("iotid").value;
        var CategoryName = document.getElementById("categoryname").value;

        workSheet.insertRow(++lastRow, [
          DeviceName,
          DeviceTypeName,
          LocationPoint,
          IoTId,
          CategoryName,
        ]);
        break;

      case "TelemetryPoint":
        // workSheet.addRow({
        //   observationname: "MS01",
        //   phenomenon: "Motion-Detection",
        //   devicename: "MS1",
        //   observedelement:
        //     '{"ECClassId": "bis.spatialelement", "UserLabel": "Door"}',
        //   iotid: "35770",
        // });
        // break;
        var ObservationName = document.getElementById("observationname").value;
        var Phenomenon = document.getElementById("phenomenon").value;
        var DeviceName = document.getElementById("devicename").value;
        var ObservedElement = document.getElementById("observedelement").value;
        var IoTId = document.getElementById("iotid").value;

        workSheet.insertRow(++lastRow, [
          ObservationName,
          Phenomenon,
          DeviceName,
          ObservedElement,
          IoTId,
        ]);
        break;

      case "Locators":
        // workSheet.addRow({
        //   locatorname: "ById",
        //   classname: "bis.spatialelement",
        //   propertyname: "ECInstanceId",
        // });
        // break;
        var LocatorName = document.getElementById("locatorname").value;
        var ClassName = document.getElementById("classname").value;
        var PropertyName = document.getElementById("propertyname").value;

        workSheet.insertRow(++lastRow, [
          LocatorName,
          ClassName,
          PropertyName,
        ]);
        break;
        
      default:
        alert(workSheet + "doesn't exist in " + fileName);
    }
    //workbook.commit();
    alert("Data Saved!");
    document.getElementById("saveData").disabled = true;
    workbook.xlsx.writeFile(fileName);
  });
}

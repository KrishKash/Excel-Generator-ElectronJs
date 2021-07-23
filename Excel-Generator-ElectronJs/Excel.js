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
    //disable createExcel button if spreadsheet already exists
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
      { header: "IoTId", key: "iotid", width: 20 },
      { header: "PhysicalElement", key: "physicalelement", width: 70 },
    ];

    // Create DeviceInterfaceSheet
    var DevicePhysicalSheet = workbook.addWorksheet("DevicePhysical");

    // table header
    DevicePhysicalSheet.columns = [
      { header: "DeviceName", key: "devicename", width: 30 },
      { header: "LocationPoint", key: "locationpoint", width: 30 },
      { header: "CategoryName", key: "categoryname", width: 30 },
    ];

    // Create TelemetryPointSheet
    var TelemetryPointSheet = workbook.addWorksheet("TelemetryPoint");

    // table header
    TelemetryPointSheet.columns = [
      { header: "ObservationName", key: "observationname", width: 30 },
      { header: "Phenomenon", key: "phenomenon", width: 30 },
      { header: "DeviceName", key: "devicename", width: 30 },
      { header: "ObservedElement", key: "observedelement", width: 70 },
      { header: "IoTId", key: "iotid", width: 20 },
    ];

    // Create LocatorsSheet
    var ActuationPointSheet = workbook.addWorksheet("ActuationPoint");

    // table header
    ActuationPointSheet.columns = [
      { header: "ActuationName", key: "actuationname", width: 30 },
      { header: "DeviceName", key: "devicename", width: 20 },
      { header: "IotId", key: "iotid", width: 20 },
    ];

    workbook.eachSheet(function (worksheet) {
      worksheet.eachRow(function (row, rowNumber) {
        row.eachCell((cell) => {
          if (rowNumber == 1) {
            // First set the background of header row
            cell.fill = {
              type: "pattern",
              pattern: "solid",
              fgColor: { argb: "C5D9F1" },
            };
          }
        });
      });
    });

    // Save Excel on Hard Disk
    workbook.xlsx.writeFile(`./${fileName}`).then(function () {
      // Success Message
      alert(`File '${fileName}' Created`);
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

  workbook.xlsx.readFile(`./${fileName}`).then(function () {
    var sheetName = document.getElementById("sname").value;
    if (!sheetName) {
      alert("Please select a sheet from dropdown first!");
      return;
    }
    var workSheet = workbook.getWorksheet(sheetName);
    var lastRow = workSheet.lastRow.number;

    switch (workSheet.name) {
      case "DeviceType":
        //Sample Data
        //   DeviceTypeName: "PIR-Motion-Sensor",
        //   Manufacturer: "Panasonic",
        //   ModelNumber: "Panasonic EKMC1603111",
        //   DeviceKind: "Infrared-Motion-Detector",
        var DeviceTypeName = document.getElementById("devicetypename").value;
        var Manufacturer = document.getElementById("manufacturer").value;
        var ModelNumber = document.getElementById("modelnumber").value;
        var DeviceKind = document.getElementById("devicekind").value;

        if (!validateFormData("devicetypefrm")) {
          alert("One or more fields cannot be left blank");
          return;
        }
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
        //Sample Data
        //   DeviceName: "MS1",
        //   DeviceTypeName: "PIR-Motion-Sensor",
        //   IoTId: "35670",
        //   PhysicalElement: `{"ECClassId": "Generic.PhysicalObject", "UserLabel": "Window"}`,
        var DeviceName = document.getElementById("devicename").value;
        var DeviceTypeName = document.getElementById("devicetypename").value;
        var IoTId = document.getElementById("iotid").value;
        var PhysicalElement = document.getElementById("physicalelement").value;

        if (!validateFormData("deviceinterfacefrm")) {
          alert("One or more fields cannot be left blank");
          return;
        }
        workSheet.insertRow(++lastRow, [
          DeviceName,
          DeviceTypeName,
          IoTId,
          PhysicalElement,
        ]);
        break;

      case "DevicePhysical":
        //Sample Data
        //   DeviceName: "MS1",
        //   LocationPoint: `{"x": 1.13, "y": 0.7, "z": 2.75}`,
        //   CategoryName: "EnvironmentCategory",
        var DeviceName = document.getElementById("pdevicename").value;
        var LocationPoint = document.getElementById("locationpoint").value;
        var CategoryName = document.getElementById("categoryname").value;

        if (!validateFormData("devicephysicalfrm")) {
          alert("One or more fields cannot be left blank");
          return;
        }
        workSheet.insertRow(++lastRow, [
          DeviceName,
          LocationPoint,
          CategoryName,
        ]);
        break;

      case "TelemetryPoint":
        //   Sample Data
        //   ObservationName: "MS01",
        //   Phenomenon: "Motion-Detection",
        //   DeviceName: "MS1",
        //   ObservedElement:'{"ECClassId": "bis.spatialelement", "UserLabel": "Door"}',
        //   IoTId: "35770",
        var ObservationName = document.getElementById("observationname").value;
        var Phenomenon = document.getElementById("phenomenon").value;
        var DeviceName = document.getElementById("tdevicename").value;
        var ObservedElement = document.getElementById("observedelement").value;
        var IoTId = document.getElementById("tiotid").value;

        if (!validateFormData("telemetrypointfrm")) {
          alert("One or more fields cannot be left blank");
          return;
        }
        workSheet.insertRow(++lastRow, [
          ObservationName,
          Phenomenon,
          DeviceName,
          ObservedElement,
          IoTId,
        ]);
        break;

      case "ActuationPoint":
        //   Sample Data
        //   ActuationName: "Actuator-ENV11",
        //   DeviceName: "MS1",
        //   IoTId: "25791",
        var ActuationName = document.getElementById("actuationname").value;
        var DeviceName = document.getElementById("adevicename").value;
        var IoTId = document.getElementById("aiotid").value;

        if (!validateFormData("actuationpointfrm")) {
          alert("One or more fields cannot be left blank");
          return;
        }
        workSheet.insertRow(++lastRow, [
          ActuationName,
          DeviceName,
          IoTId,
        ]);
        break;

      default:
        alert(`'${workSheet}' doesn't exist in '${fileName}'`);
    }

    alert("Data Saved!");
    document.getElementById("saveData").disabled = true;
    workbook.xlsx.writeFile(`./${fileName}`);
  });
}

function validateFormData(formId) {
  //var elements = document.getElementsByTagName("input");
  var elements = document.getElementById(formId).getElementsByTagName("input");
  isValid = true;
  for (var i = 0; i < elements.length; i++) {
    // console.log(elements[i].value);
    if (elements[i].value == "") isValid = false;
  }
  return isValid;
}

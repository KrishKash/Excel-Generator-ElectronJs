const ExcelJS = require("exceljs");
const fs = require("fs");
function createExcel() {
  //A new Excel Work Book
  var workbook = new ExcelJS.Workbook();

  var fileName = document.getElementById("ssname").value;
  if (!fileName) {
    alert("Spreadsheet name can not be empty!");
    return;
  }
  fileName += ".xlsx";
  var downloadFolder = process.env.USERPROFILE + "/Downloads";
  console.log(`${downloadFolder}`);

  if (fs.existsSync(`${downloadFolder}/${fileName}`)) {
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
            cell.border = {
              top: { style: "thin" },
              left: { style: "thin"},
              bottom: { style: "thin" },
              right: { style: "thin" }
            };
          }
        });
      });
    });

    // Save Excel on Hard Disk
    workbook.xlsx.writeFile(`${downloadFolder}/${fileName}`).then(function () {
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

function saveFormData(filename) {
  // Excel Work Book
  var workbook = new ExcelJS.Workbook();
  var curfile = document.getElementById("ssname").value;
  if (!curfile) {
    alert("Please enter the file name first!");
    return;
  }
  curfile += ".xlsx";
  fileName = filename + ".xlsx";
  var downloadFolder = process.env.USERPROFILE + "/Downloads";

  if (!fs.existsSync(`${downloadFolder}/${fileName}`)) {
    alert(`File mismatched, you are writing in '${curfile}' which does not exist in '${downloadFolder}'`);
    return;
  }

  workbook.xlsx.readFile(`${downloadFolder}/${fileName}`).then(function () {
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

        //reset the form once data is written in the sheet
        document.getElementById("devicetypefrm").reset();
        document.getElementById("saveData").disabled = false;
        break;

      case "DeviceInterface":
        //Sample Data
        //   DeviceName: "MS1",
        //   DeviceTypeName: "PIR-Motion-Sensor",
        //   IoTId: "35670",
        //   PhysicalElement: `{"ECClassId": "Generic.PhysicalObject", "UserLabel": "Window"}`,
        var DeviceName = document.getElementById("devicename").value;
        var DeviceTypeName = document.getElementById("idevicetypename").value;
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

        //reset the form once data is written in the sheet
        document.getElementById("deviceinterfacefrm").reset();
        document.getElementById("saveData").disabled = false;
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

        //reset the form once data is written in the sheet
        document.getElementById("devicephysicalfrm").reset();
        document.getElementById("saveData").disabled = false;
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

        //reset the form once data is written in the sheet
        document.getElementById("telemetrypointfrm").reset();
        document.getElementById("saveData").disabled = false;
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

        //reset the form once data is written in the sheet
        document.getElementById("actuationpointfrm").reset();
        document.getElementById("saveData").disabled = false;
        break;

      default:
        alert(`'${workSheet}' doesn't exist in '${fileName}'`);
    }

    alert("Data Saved!");
    //document.getElementById("saveData").disabled = true;
    workbook.xlsx.writeFile(`${downloadFolder}/${fileName}`);
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

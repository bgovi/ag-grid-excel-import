var gridOptions = {
    columnDefs: [
        { field: "Id", minWidth: 180 },
        { field: "Name", minWidth: 150 },
        { field: "Country"}
    ],

    defaultColDef: {
        resizable: true,
        minWidth: 80,
        flex: 1
    },

    rowData: []
};


// pull out the values we're after, converting it into an array of rowData

function populateGrid(excelRows) {
    // our data is in the first sheet
    //var firstSheetName = workbook.SheetNames[0];
    //var worksheet = workbook.Sheets[firstSheetName];

    // we expect the following columns to be present


    var rowData = [];
    // start at the 2nd row - the first row are the headers
    //var rowIndex = 2;
    //Add the data rows from Excel file.
    for (var i = 0; i < excelRows.length; i++) {
        //Add the data row.
        var row = {}
        row['Id'] = excelRows[i].Id;
        row['Name'] = excelRows[i].Name;
        row['Country'] = excelRows[i].Country;
        rowData.push(row)

    }
    // finally, set the imported rowData into the grid
    gridOptions.api.setRowData(rowData);
}

function importExcel() {
    //Reference the FileUpload element.
    var fileUpload = document.getElementById("myfile");
    console.log(fileUpload.files[0])
    //Validate whether File is valid Excel file.
    var regex = /^([a-zA-Z0-9\s_\\.\-:])+(.xls|.xlsx)$/;
    if (regex.test(fileUpload.value.toLowerCase())) {
        if (typeof (FileReader) != "undefined") {
            var reader = new FileReader();

            //For Browsers other than IE.
            if (reader.readAsBinaryString) {
                reader.onload = function (e) {
                    console.log(e.target.result)
                    ProcessExcel(e.target.result);
                };
                reader.readAsBinaryString(fileUpload.files[0]);
            } else {
                //For IE Browser.
                reader.onload = function (e) {
                    var data = "";
                    var bytes = new Uint8Array(e.target.result);
                    for (var i = 0; i < bytes.byteLength; i++) {
                        data += String.fromCharCode(bytes[i]);
                    }
                    ProcessExcel(data);
                };
                reader.readAsArrayBuffer(fileUpload.files[0]);
            }
        } else {
            alert("This browser does not support HTML5.");
        }
    } else {
        alert("Please upload a valid Excel file.");
    }
};
function ProcessExcel(data) {
    //Read the Excel File data.
    var workbook = XLSX.read(data, {
        type: 'binary'
    });

    //Fetch the name of First Sheet.
    var firstSheet = workbook.SheetNames[0];

    //Read all rows from First Sheet into an JSON array.
    var excelRows = XLSX.utils.sheet_to_row_object_array(workbook.Sheets[firstSheet]);
    //return excelRows
    populateGrid(excelRows)

    //Create a HTML Table element.

};








// wait for the document to be loaded, otherwise
// ag-Grid will not find the div in the document.
document.addEventListener("DOMContentLoaded", function () {

    // lookup the container we want the Grid to use
    var eGridDiv = document.querySelector('#myGrid');

    // create the grid passing in the div to use together with the columns & data we want to use
    new agGrid.Grid(eGridDiv, gridOptions);
});
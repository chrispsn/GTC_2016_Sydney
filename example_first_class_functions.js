'use strict'

// Imports - inspired by http://stackoverflow.com/a/10295725
var FS = new ActiveXObject("Scripting.FileSystemObject");
eval(FS.OpenTextFile("U_GlobalsMgmt.js", 1).ReadAll());

// Set up Excel
var app = new ActiveXObject("Excel.Application");
var directory = WScript.CreateObject("WScript.Shell").CurrentDirectory + '\\'
var wb = app.Workbooks.Open(directory + "excel_test.xlsm",  false)
var lo = wb.Worksheets("Sheet1").ListObjects("table_test")

// Do something
function change_data(lo) {
   var cell = lo.DataBodyRange.Cells(2, 1);
   cell.Value2 = Math.random();
   return cell.Value2
}

var result = U_GlobalsMgmt.temp_disable_screenupdating(
    app, 
    function() { return change_data(lo) }
)

// Tear down Excel
// Need to explicitly quit because will otherwise keep running after script ends: 
// https://technet.microsoft.com/en-us/library/ee198894.aspx
U_GlobalsMgmt.temp_disable_displayalerts(app, function(){ wb.Save(); });
app.Quit(); 

WScript.Echo(result);

<script type="text/javascript">
// run the Excel load handler on page load
// if (window.attachEvent) {
//   window.attachEvent("onload", loadEwaOnPageLoad);
// } else {
//   window.addEventListener("DOMContentLoaded", loadEwaOnPageLoad, false);
// }

function loadEwaOnPageLoad() {
    
//   var fileToken = "SDBBABB911BCD68292!110/-4923638281765748078/t=0&s=0&v=!ALTlXd5D3qSGJKU";
  var fileToken = "2F5C502D2761C73!792/-4923638281765748078/t=0&s=0&v=!ACxhGwfSQOn7RcI";
//   var fileToken = "SD310A16DD64ED7E41!112/3533661997762444865/";
  var props = {
          uiOptions: {
                showGridlines: false,
      selectedCell: "'Sheet1'!C9",
                showRowColumnHeaders: false,
                showParametersTaskPane: false
          },
          interactivityOptions: {
                allowTypingAndFormulaEntry: false,
                allowParameterModification: false,
                allowSorting: false,
                allowFiltering: false,
                allowPivotTableInteractivity: false
          }
  };
  Ewa.EwaControl.loadEwaAsync(fileToken, "myExcelDiv2", props, onEwaLoaded);     
}
function onEwaLoaded() {
    // document.getElementById("loadingdiv").style.display = "none";
}
// This sample gets the value in the highlighted cell. 
// Try clicking on different cells then running the sample.
function execute()
{
// Get unformatted range values (getValuesAsync(1,...) where 1 = Ewa.ValuesFormat.Formatted)
Ewa.getActiveWorkbook().getActiveCell().getValuesAsync(1,getRangeValues,null);
}     

function getRangeValues(asyncResult)
{
// Get the value from asyncResult if the asynchronous operation was successful.
if (asyncResult.getCode() == 0)
{
    // Get the value in active cell (located at row 0, column 0 of the
    // range which consists of a single cell (the "active cell")).
    alert("Result: " + asyncResult.getReturnValue()[0][0]);
}
else 
{
      alert("Operation failed with error message " + asyncResult.getDescription() + ".");
}    
}
</script>
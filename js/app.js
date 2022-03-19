//test table:
//https://onedrive.live.com/edit.aspx?resid=2F5C502D2761C73!792&ithint=file%2cxlsx&authkey=!ACxhGwfSQOn7RcI
var fileToken = "2F5C502D2761C73!792/-4923638281765748078/t=0&s=0&v=!ACxhGwfSQOn7RcI";

//Work table:
//https://onedrive.live.com/view.aspx?resid=7B3CD0791DDF2DD1!7114&ithint=file%2cxlsx&authkey=!AJM7ysmoze_cS3c
var fileToken = "7B3CD0791DDF2DD1!7114/-4923638281765748078/t=0&s=0&v=!AJM7ysmoze_cS3c";

var ewa = null;
var DataArray;

// Run the Excel load handler on page load.
if (window.attachEvent) {
    window.attachEvent("onload", loadEwaOnPageLoad);
} else {
    window.addEventListener("DOMContentLoaded", loadEwaOnPageLoad, false);
}

function loadEwaOnPageLoad() {
    var props = {
        uiOptions: {
            showGridlines: true,
            showRowColumnHeaders: false,
            showParametersTaskPane: false
        },
        interactivityOptions: {
            allowTypingAndFormulaEntry: false,
            allowParameterModification: false,
            allowSorting: false,
            allowFiltering: true,
            allowPivotTableInteractivity: false
        }
    };
    // Embed workbook using loadEwaAsync
    Ewa.EwaControl.loadEwaAsync(fileToken, "myExcelDiv2", props, onEwaLoaded);
    
}

function onEwaLoaded(asyncResult) { 
    if (asyncResult.getSucceeded())
    {
        // Use the AsyncResult.getEwaControl() method to get a reference to the EwaControl object
        // ewa = asyncResult.getEwaControl();
        ewa = Ewa.EwaControl.getInstances().getItem(0);
        console.log("ewa OK");
        LoadData();
    }
    else
    {
        alert("Async operation failed!");
    }
    // ...
}    

function LoadData(){
    var range = ewa.getActiveWorkbook().getActiveSelection();
    console.log(range.getSheet().getName());
    var rowCnt = Number(range.getRow()) + 1;
    console.log(rowCnt);
    var colCnt = Number(range.getColumn());
    // ewa.getActiveWorkbook().getActiveSelection ().getValuesAsync(0,getRangeValues,null);
    var range2 = ewa.getActiveWorkbook().getActiveSheet().getRange(0,0,400,20);
    console.log(range2);
    range2.getValuesAsync(0,getRangeValues,range2);    
    hideTable();
    
}

function getRangeValues(asyncResult){
    // Get the value from asyncResult if the asynchronous operation was successful.

    //clear
    DataArray = [];
        // Get range from user context.
        var range = asyncResult.getUserContext();
        
        // Get the array of range values from asyncResult.
        var values = asyncResult.getReturnValue();
        
        // Display range coordinates in A1 notation and associated values.
        var output = "Values from range" + range.getAddressA1() + "\n";
        output = output + "********\n";
        
        // Loop through the array of range values.
        for (var i = 1; i < values.length; i++)
        {
            output += values[i][0] + "\n";

            var newDataItem = {
                Nr: values[i][0],
                Code: values[i][2],
                Stock: values[i][8],
                Price: values[i][18]
                };
                if(Number(newDataItem.Nr)>0){
                    DataArray.push(newDataItem);  
                }
        // for (var j = 0; j < values[i].length; j++)
        // {
        // 		output= output + values[i][j] + '\n'
        // }
        }
        
        output = output + "********";

        // Display each value in the array returned by getValuesAsync.
        // alert(output);
        GetData();
    
}

function hideTable(){
    var Array = document.getElementsByClassName('wrap');
    Array[0].className = "hidden";
}

//URL: index.html?index=1
function GetData(){
    var index = window.location.search.substring(1).split("=")[1];
    var sheetindex = 0;
    var found = false;
    for(var i = 0; i<DataArray.length;i++){
        if(DataArray[i].Nr == index){
            found = true;
            document.getElementById("Number").innerHTML = DataArray[i].Nr;
            document.getElementById("Code").innerHTML =  DataArray[i].Code;
            document.getElementById("Stock").innerHTML =  DataArray[i].Stock;
            document.getElementById("Price").innerHTML =  DataArray[i].Price;
        }
    }

}

function LoadQRCodes(){
    let inputValue = document.getElementById("ItemNr").value; 
    var ID = "qrcode_" + inputValue;
    var newQRCode = '<div class="QR_Row">  <div class="Description">QR code for item number:' + inputValue + '</div>  <div id="'+ ID + '" style="width:200px; height:200px; margin-top:15px;"></div></div>'
    document.getElementById("QR_Code_wrap").innerHTML += newQRCode;
    var URL = "https://citizenitza.github.io/QR_Code_Onedrive/index.html?index=" + inputValue;
    var qrcode = new QRCode(document.getElementById(ID), {
        width : 200,
        height : 200
    });
    qrcode.makeCode(URL);

}
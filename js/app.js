			/*
			* This code uses the Microsoft Office Excel JavaScript object model to programmatically insert the
			* Excel Web App into a div with id=myExcelDiv. The full API is documented at
			* https://msdn.microsoft.com/library/hh315812.aspx. There you can find out how to programmatically get
			* values from your Excel file and how to use the rest of the object model. 
			*/
		
			// Use this file token to reference Book1.xlsx in the Excel APIs
			// Replace the the placeholder for the  filetoken with your value
            var fileToken = "2F5C502D2761C73!792/-4923638281765748078/t=0&s=0&v=!ACxhGwfSQOn7RcI";
			var ewa = null;
		
			// Run the Excel load handler on page load.
			if (window.attachEvent)
			{
				window.attachEvent("onload", loadEwaOnPageLoad);
			} else
			{
				window.addEventListener("DOMContentLoaded", loadEwaOnPageLoad, false);
			}
		
			function loadEwaOnPageLoad()
			{
				var props = {
					uiOptions: {
						showGridlines: false,
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
				// Embed workbook using loadEwaAsync
				Ewa.EwaControl.loadEwaAsync(fileToken, "myExcelDiv2", props, onEwaLoaded);
			}
		
			function onEwaLoaded(asyncResult)
			{ 
				if (asyncResult.getSucceeded())
				{
					// Use the AsyncResult.getEwaControl() method to get a reference to the EwaControl object
					// ewa = asyncResult.getEwaControl();
					ewa = Ewa.EwaControl.getInstances().getItem(0);
                    console.log("ewa OK");
					
				}
				else
				{
					alert("Async operation failed!");
				}
				// ...
			}    

			function test(){
				var range = ewa.getActiveWorkbook().getActiveSelection();
				console.log(range.getSheet().getName());
				var rowCnt = Number(range.getRow()) + 1;
                console.log(rowCnt);
				var colCnt = Number(range.getColumn());
                // ewa.getActiveWorkbook().getActiveSelection ().getValuesAsync(0,getRangeValues,null);
                var range2 = ewa.getActiveWorkbook().getActiveSheet().getRange(0,0,rowCnt,20);
                console.log(range2);
                range2.getValuesAsync(0,getRangeValues,range2);    
			}

            function getRangeValues(asyncResult){
				// Get the value from asyncResult if the asynchronous operation was successful.

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
					for (var j = 0; j < values[i].length; j++)
					{
							output= output + values[i][j] + '\n'
					}
					}
					
					output = output + "********";

					// Display each value in the array returned by getValuesAsync.
					alert(output);
				

			}
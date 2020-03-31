# ExcelReading
Allows Bizagi to read from an excel spreadsheet in csv, xls and xslx.

Please read the Bizagi Component Library [help page](http://help.bizagi.com/bpm-suite/en/index.html?component_library.htm) to understand how to use component libraries in Bizagi.

## Quick Example
This is an example of a Bizagi Expression that invokes the library
```javascript
/*
*	Rule Name: Evt_ReadFromClaimsFile
*	Creator: Andres Sarmiento
*	Creation Date: 31/03/2020
*	Description: Reads the Excel file using the dll component
*/
Common.Exe_Tracing(Me,"Evt_ReadFromClaimsFile","*** START ***" );

//Verify if file exists
if(<exists(XPATH_TO_FILE)>)
{
	//Get data and file name from the current file
	var data = <XPATH_TO_FILE>.get(0).getXPath("Data");
	var FileName = <XPATH_TO_FILE>.get(0).getXPath("FileName");

	//Initializes the Excel reader
	var ER = ExcelReader.ExcelReader.GetExcelReader(CHelper.ToBase64(data),FileName,"Sheet1",',');
  
	//If trim is set to true the component will trim all the cells that are read
	ER.setTrim(true);
  
	//Current row to read
	var row = 0;
  
	//Set termination flag
	var stop = false;


	//Move the reader to the next avaialable row (row = 0)
	//ER.HasNextRow();
	//While there are more row to read, also the HasNextRow function moves the reader to the next readable line
	while (ER.HasNextRow() && !stop)
	{
		//Updates the row
		row ++;
    
		//Print current row to traces
		Common.MUC_Tracings(Me,"Evt_ReadFromClaimsFile","Reading row " + (row));
    
		// If the file has the first column with data
		if (ER.GetRowField(0) != null) //null is when the row only contains empty cells
		{
			//Read the information for any column that is required for the current row
			var Column0 = ER.GetRowField(0);
			var Column1 = ER.GetRowField(1);
		}
		else //If it is empty the reading will stop by raising the stop flag
			stop = true;
	}//WHILE
}
Common.Exe_Tracing(Me,"Evt_ReadFromClaimsFile","***  END  ***" );
```

## Additional steps

After the first deployment or after a Bizagi upgrade the reference libraries are copied manually into the web application. DO NOT EVERWRITE ANY LIBRARY IF THEY ALREADY EXISTS IN THE DESTINATION FOLDER. These libraries are:
- ICSharpCode.SharpZipLib
- LumenWorks.Framework.IO
- Microsoft.Office.Interop.Excel
- NPOI
- NPOI.OOXML
- NPOI.OpenXml4Net
- NPOI.OpenXmlFormats

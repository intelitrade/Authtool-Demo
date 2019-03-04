
try{
	//Check version of internet explorer running
	var iIEVer = checkIEVer();
	if(iIEVer == 0) 
	{
		alert(sIEInfo);
		window.close();
    }
	
	var sMappingHTML = "";

	function keyDown() {
		//executed when a key has been pressed
		var iKeyCode = event.keyCode
		//Check if enter key was pressed
		if (iKeyCode == 13) 
		{
			okButtonClick();
		}

		//check if ecape key has been pressed
		if (iKeyCode == 27) {
			cancelButtonClick();
		}
	}

	function ApplyButtonClick()
	{
		try{
			//debugger;
			//debugger;
			//Get values to be returned, return in JSON format
			var JSONReturnValue = createReturnValue();
			window.returnValue = JSONReturnValue;
			window.close();
			//oDoc.recalculate(0);	
		}catch(e)
		{
			alert(e.description);
		}
	}

	
	function getRowNumber(oTable,sGiud,sGUIDID)
	{
		try{
			var iRows = oTable.nRows();
			var iRow = 0;
			for(var i=1;i<=iRows;i++)
			{
				var oRow = oTable.getRow(i);
				if(isInputValid(sGUIDID)!=""){
					if(oRow.propGet("TEMPGUID")===sGiud)
					{
						iRow = i;
						break;
					}
				}else{
					if(oRow.propGet("GUID")===sGiud)
					{
						iRow = i;
						break;
					}					
				}
			}
			return iRow;
		}catch(e)
		{
			alert(e.description);
		}finally{
			oTable = null;
		}
	}
	
	function updateMapNo(oElement)
	{
		try{
			
			
			
			/*debugger;
			debugger;
			
			var sStr = 'grp(c1,TOKEN(HEADER,1,"|"),entity(TOKEN(HEADER,2,"|"),TOKEN(rc(-7),1,"|")+TOKEN(COLL3,1,"|")))+grp(c1,TOKEN(HEADER,1,"|"),entity(TOKEN(HEADER,2,"|"),TOKEN(rc(-7),1,"|")+TOKEN(COLL3,2,"|")))+grp(c1,TOKEN(HEADER,1,"|"),entity(TOKEN(HEADER,2,"|"),TOKEN(rc(-7),2,"|")+TOKEN(COLL3,1,"|")))+grp(c1,TOKEN(HEADER,1,"|"),entity(TOKEN(HEADER,2,"|"),TOKEN(rc(-7),2,"|")+TOKEN(COLL3,2,"|")))';
			
			var sRegExp = new RegExp("rc(-7)","gi");
			sStr = sStr.replace(sRegExp,"'1.1.100.100'");
			debugger;
			debugger;*/
			//var sCellToUpdate = oElement.id;
			//var oMappingCell = oDoc.cell(sCellToUpdate)
			//if(oMappingCell)
            //{
                var oSourceElement = document.getElementById(oElement.srcelement);
				if(isInputValid(oSourceElement))
				{
					var sExpression = "grpdesc('M','"+oElement.value+"',3,"+oDoc.cell("LANGUAGE").value+")";
					var sNewDesc = oDoc.interpret(sExpression);	
                    //if (isInputValid(sNewDesc)){
						 document.getElementById("PREVIEWDESC_"+oElement.srcelement).innerText = oSourceElement.description = oSourceElement.innerText = sNewDesc;
						 oSourceElement.mapping = oElement.value;
						//oSourceElement.description = sNewDesc;
						
						//document.getElementById("PREVIEWDESC_"+oElement.srcelement).innerText = sNewDesc;
					//}
				}
				
				/*oMappingCell.value = oElement.value;
				//get the description cell and remove overridden if it is there
				var sDescCell = oElement.getAttribute("desccell");
				var oCell = oDoc.cell(sDescCell);
				if(oCell)
				{
					oCell.overridden = 0;
					//oElement.innerText=oCell.value;

				}
				//MFG.D14
				oDoc.recalculate(0);
				//Update description on UI
				var oSrcElement = document.getElementById(oElement.srcelement);
				if(oSrcElement)
				{
					oSrcElement.innerHTML="["+oDoc.cell(sDescCell).value+"]";
					//Reload preview pane to have the right properties displaying, otherwise some fields will show wrong data
					updatePreviewPane(oSrcElement);
				}*/
			//}
			
		}catch(e)
		{
			alert(e.description);
		}finally{
			oElement = null;
		}
	}

	function updateDescCell(oElement)
	{
		try{
			//debugger;
			//debugger;
			//var sDescCell = oElement.getAttribute("desccell");
			//oDescCell = oDoc.cell(sDescCell);
			//if(oDescCell)
			//{
                var oSourceElement = document.getElementById(oElement.srcelement);
				if(isInputValid(oSourceElement))
				{
                    var sNewDesc = oElement.value;
					//if (isInputValid(sNewDesc)){
						oSourceElement.innerText = sNewDesc;
						oSourceElement.description = sNewDesc;
						document.getElementById("PREVIEWDESC_"+oElement.srcelement).innerText = sNewDesc;
					//}
				}				
				
				/*oDescCell.value = oElement.value;
				oDoc.recalculate(0);
				//Update description on UI
				var oSrcElement = document.getElementById(oElement.srcelement);
				if(oSrcElement)
				{
					oSrcElement.innerHTML="["+oDoc.cell(sDescCell).value+"]";
					//Reload preview pane to have the right properties displaying, otherwise some fields will show wrong data
					updatePreviewPane(oSrcElement);
				}*/
			//}
		}catch(e)
		{
			alert(e.description);
		}finally{
			oElement = null;
		}
	}
	
	function updateColDesc(sCellName,oElement)
	{
		try{
			//debugger;
			//debugger;
			var sSourceElement = oElement.srcelement;
			if(isInputValid(sSourceElement))
			{
				var oSourceElement = document.getElementById(sSourceElement);
				if(oSourceElement)
				{
					oSourceElement.description = oElement.value;
				}
			}
			/*var oCell = oDoc.cell(sCellName);
			if(oCell)
			{
				oCell.calculation = '"'+oElement.value+'"';
			}
			oDoc.recalculate(0);
			*/
		}catch(e)
		{
			alert(e.description);
		}finally{
			oElement = null;
			oCell = null;
		}
	}
	
	function getNoteNumRef()
	{
		try{
			
		}catch(e){
			
		}finally{
		}
	}
		
	function okButtonClick()
	{
		try{
			clearAllHighlighting();
			window.close();
		}catch(e)
		{
			alert(e.description);
		}
	}

	//Event handler (Keyboard)
	document.onkeydown = keyDown
}catch(e)
{
	
}

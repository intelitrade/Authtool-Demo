
try{
	//Check version of internet explorer running
	var iIEVer = checkIEVer();
	if(iIEVer == 0) 
	{
		alert(sIEInfo);
		window.close();
    }
	
	var iExport = 999;
	
	//get the arguments passed into the html
	//var params_list = new VBArray(window.dialogArguments);
	//var params = params_list.toArray();
	//var curDoc = params[0];
	//var sClosestSection = params[1];

	//var sExportType = "pdf";
	
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
	
	function selectAllOptions()
	{
		if(document.getElementById("ALL").checked===true)
		{
			//aSelection=["ALL"];
			var oElements = document.getElementsByName("selection");
			var aSelection = [];
			if(oElements)
			{
				var iCheckBoxOptions = oElements.length;
				for(var i=0;i<iCheckBoxOptions;i++)
				{
					oElements[i].checked=true;
				}
			}
			return;
		}
		return aSelection;	
	}

	function onLoad() 
	{
		window.focus()
	}
	
	function okButtonClick()
	{
		try{
			var aSelection = [];
			
			var aTempSection = selectedoptions();
			if(isInputValid(aTempSection))
				aSelection = aTempSection;
			
			window.returnValue = aSelection;
			window.close();
		}catch(e)
		{
			logError(e);
		}
	}
	
	function selectedoptions()
	{
		var oElements = document.getElementsByName("restoresetting");
		var aSelection = [];
		if(oElements)
		{
			var iCheckBoxOptions = oElements.length;
			for(var i=0;i<iCheckBoxOptions;i++)
			{
				if(oElements[i].checked===true)
				{
					aSelection[aSelection.length] = oElements[i].value;
				}
			}
		}
		
		return aSelection;
	}	
	
	//Event handler (Keyboard)
	document.onkeydown = keyDown
}catch(e){}



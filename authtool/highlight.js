function highlightElement(oElement)
{	
	try{
		//debugger;
		//debugger;
		if(!oElement)	
			oElement = event.srcElement;		
		
		if(oElement)
		{
			//Check if the element already exists and needs highlighting removed as well as if it needs to be removed from the list
			if(aSelectedItems.length>0)
			{
				for(var i=0;i<aSelectedItems.length;i++)
				{
					if(aSelectedItems[i]==oElement)
					{
						oElement.style.borderBottom="1px dotted red";
						aSelectedItems.splice(i, 1);
						return;
					}
				}
			}
		
			var bCtrlKey = false;
			bCtrlKey = event.ctrlKey;
			//Clear array containing items selected if ctrl key not held down
			if(!bCtrlKey)
			{
				aSelectedItems = [];
			}		
			//reset border settings for all other child elements
			//var all = document.getElementsByTagName("*");
			var all = document.getElementsByTagName("LI");
			for (var i=0, max=all.length; i < max; i++) {
				 // Do something with the element here
				 var oListElement = all[i];
				 if(oListElement && !bCtrlKey)
				 {
					//oListElement.style.border="";
					oListElement.style.borderBottom="1px dotted red";
				 }
			}			
			oElement.style.borderBottom="3px solid #800000";
			//oElement.style.fontWeight="bold";
			aSelectedItems[aSelectedItems.length]=oElement;
		}
	}catch(e)
	{
		alert(e.description);
	}finally{
		oElement = null;
		oParentNode = null;
	}
}

function highlightTable(sTable,oTable)
{	
	try{
		var oCirclingList = oDoc.getScriptCircleList();  //get the circling list of the document
		var oCircling = null, oCirclingSection = null;
		var iCount = 0;
		
		if (oCirclingList != null) {
			iCount = oCirclingList.count();
			
			if (iCount > 0) {
				oCirclingList.deleteAllCircles();
			}
			oCirclingList = null;
		}
		//var sSection = sLabel.split("_properties")[0];
		var oCircleList = oDoc.getScriptCircleList();
		if(oCircleList)
		{
			//var oSection = oDoc.sectionByName(sSection)
			//return
			if (oCircleList && oTable)
			{
				var oCircle = oCircleList.addCircle(sTable,1,2,1,0x000000ff,2);
				oCircle.addObject(oTable);
				//draw the circling shape
				oCircleList.drawCircles();
			}
		}

	}catch(e)
	{
		alert(e.description);
	}finally{
		oTable = null;
		oCircleList = null;
		oCirlce = null;
	}
}

function highlightSection(sLabel)
{	
	try{
		var oCirclingList = oDoc.getScriptCircleList();  //get the circling list of the document
		var oCircling = null, oCirclingSection = null;
		var iCount = 0;
		
		if (oCirclingList != null) {
			iCount = oCirclingList.count();
			
			if (iCount > 0) {
				oCirclingList.deleteAllCircles();
			}
			oCirclingList = null;
		}
		var sSection = sLabel.split("_properties")[0];
		var oCircleList = oDoc.getScriptCircleList();
		if(oCircleList)
		{
			var oSection = oDoc.sectionByName(sSection)
			//return
			if (oCircleList && oSection)
			{
				var oCircle = oCircleList.addCircle(sSection,1,2,1,0x000000ff,2);
				oCircle.addObject(oSection);
				//draw the circling shape
				oCircleList.drawCircles();
			}
		}
	}catch(e)
	{
		alert(e.description);
	}finally{
		oCircleList = null;
		oCirlce = null;
	}
}

function highlightRow(sCellName,oRow)
{
	/*
		CaseViewScriptCircle addCircle (BSTR sLabel, int bShow,
				int iLineThickness, int iLineStyle, int iLineColor, int iShape)
				
		Parameters
		BSTR sLabel
		int bShow
		Possible Values
		0 - Circle will NOT be displayed
		1 - Circle will be displayed
		int iLineThickness
		int iLineStyle
		CIRCLINGLINESTYLES
		int iLineColor
		
		How do I specify a color as an integer in CaseView Scripting?
		Detailed Steps
		Assign a value (between 0 - 255) for each Blue | Green | Red value
		Convert each value to Hexadecimal (i.e. between 00 to FF)
		Concatenate [Bv] [Gv] [Rv]
		Convert the resulting Hexadecimal value back to Decimal
		Note: You can use Windows Calculator for ease of conversion.

		CV COLORS
		Examples
		Step 1: Blue = 0; Green = 0; Red = 255
		Step 2: Blue = 00; Green = 00; Red = FF
		Step 3: 0000FF
		Step 4: 255 -> *Integer Value for Red*			

		Constant Name	Description	Value
		BLACK	Black	0x00000000
		LTGRAY	Light Gray	0x00c0c0c0
		MEDGRAY	Medium Gray	0x00a0a0a0
		DKGRAY	Dark Gray	0x00808080
		RED	Red	0x000000ff
		BLUE	Blue	0x00ff3030
		LIME	Lime	0x0000ff00
		MAROON	Maroon	0x00000084
		GREEN	Green	0x00008200
		OLIVE	Olive	0x00008284
		NAVY	Navy	0x00840000
		PURPLE	Purple	0x00840084
		TEAL	Teal	0x00848200
		YELLOW	Yellow	0x0000ffff
		FUSCHIA	Fuschia	0x00ff00ff
		AQUA	Aqua	0x00ffff00
		WHITE	White	0x00ffffff
		
		Remarks
		
		How to specify an integer value for colors?
		To Use Other Colors: 
		Get value (0-255) for each Red/Green/Blue value.
		Convert the integer value to hex (i.e., between 00 to ff)
		Concatenate 0x00 [Bv] [Gv] [Rv]

		Example:

		BLACK = 0x00000000
		WHITE = 0x00ffffff

		How to specify an integer value for colors?
		int iShape
		CIRCLINGSHAPES	

		Constant Name	Description	Value
		CIRCLING_SHAPE_DEFAULT	Script Circle Member	0
		CIRCLING_SHAPE_CIRCLE	Script Circle Member	1
		CIRCLING_SHAPE_RECTANGLE	Script Circle Member	2
		CIRCLING_SHAPE_H_BRACKETS	Script Circle Member	3
		CIRCLING_SHAPE_V_BRACKETS	Script Circle Member	4

		Constant Name	Description	Value
		CIRCLING_LINE_SOLID	Script Circle Member	0
		CIRCLING_LINE_DASH	Script Circle Member	1
		CIRCLING_LINE_DOT	Script Circle Member	2
		CIRCLING_LINE_DASHDOT	Script Circle Member	3
		CIRCLING_LINE_DASHDOTDOT	Script Circle Member	4
	*/
	//debugger;
	//debugger;
	
	try{
		
		if(!isInputValid(sCellName) || !isInputValid(oRow))
			return;
		
		var oCirclingList = oDoc.getScriptCircleList();  //get the circling list of the document
		var oCircling = null, oCirclingSection = null;
		var iCount = 0;
		
		if (oCirclingList != null) {
			iCount = oCirclingList.count();
			
			if (iCount > 0) {
				oCirclingList.deleteAllCircles();
			}
			oCirclingList = null;
		}
		//var sSection = sLabel.split("_properties")[0];
		var oCircleList = oDoc.getScriptCircleList();
		if(oCircleList)
		{
			var oCell = oDoc.cell(sCellName);
			//return
			if (oCircleList && oCell)
			{
				var oCircle = oCircleList.addCircle(sCellName,1,2,1,0x000000ff,2);
				oCircle.addObject(oCell);
				//draw the circling shape
				oCircleList.drawCircles();
			}else if(oCircleList && oRow)
			{
				var oCircle = oCircleList.addCircle("xxxx",1,2,1,0x000000ff,2);
				oCircle.addObject(oRow);
				//draw the circling shape
				oCircleList.drawCircles();					
			}
		}
	}catch(e)
	{
		alert(e.description);
	}finally{
		oRow = null;
		oCircleList = null;
		oCirlce = null;
	}
}	


function removeHighlightOnSection(sLabel)
{
	try{	
		var oCircleList = oDoc.getScriptCircleList();
		if (oCircleList)
		{
			
			var oCircle = oCircleList.getCircleByLabel(sLabel);
			if (oCircle)
			{
				//Get the  number of cirles in the circle list
				var iCirleCount = oCircleList.count();
				for (j=0;j<=iCirleCount;j++)
				{
					if (oCircleList.getCircle(j) && oCircleList.getCircle(j).label===sLabel)
					{
						oCircleList.deleteCircle(j);
						break;
					}
				}
			}
		}	
	}catch(e)
	{
		alert(e.description);
	}finally{
		oCircleList = null;
		oCirlce = null;
	}
}

//Clearss all highlighting in the document
function clearAllHighlighting()
{
	try{
		var oCirclingList = oDoc.getScriptCircleList();  //get the circling list of the document
		var oCircling = null, oCirclingSection = null;

		if (oCirclingList != null) {
			var iCount = oCirclingList.count();
				
			if (iCount > 0) {
				oCirclingList.deleteAllCircles();
			}		
		}
		oDoc.recalculate(0);			
	}catch(e)
	{
		alert(e.description);
	}finally{
		oCirclingList = null;
		oCircling = null;
	}
}
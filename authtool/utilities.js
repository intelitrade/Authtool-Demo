//Handles on mouse up events
function mouseUp()
{
	try{
		debugger;
		debugger;
		if (oElement && isInputValid(sRefElement)) {
			//Check if the element need to be added or note
			if(oElement.style.color=="green")
			{
				var sElementName = oElement.name;
				var oRefElement = document.getElementById(sRefElement);
				var iRowToInsert = 1;//prompt("Please enter number of "+sElementName+" row(s) to insert", 1);
				if(iRowToInsert==null)
				{
					oElement.style.position = 'static';
					oElement.onmousemove = null;		
				}else{
					
					//get rid of the on mouse move event
					oElement.onmousemove = null;
					oElement.style.position = 'static';						

					//Create a progress bar to show number of rows being inserted and show when each row has been inserted
					var oProgBar = oDoc.createProgressBar((iRowToInsert>1?"Inserting rows":"Inserting a row"),iRowToInsert,1)
					//if(!oRefElement)	
					for(var x=1;x<=iRowToInsert;x++)
					{
						//Create an element to add to the html for each new row
						var node=document.createElement("LI");					
						//Give it a label
						var textnode=document.createTextNode(sElementName);
						//Attach the text to the node
						node.appendChild(textnode);
						//set onclick event
						node.setAttribute("onclick","gotoSection(this);highlightElement(this);updatePreviewPane(this)");
						node.onclick = function() { gotoSection();highlightElement();updatePreviewPane() };
						
						//Make the new item cold to make it stand ouot and easier to find as well as get the users attention
						node.style.fontWeight="bold";
						//Get the element to insert after
                        var oRefParent = oRefElement.parentNode;	
						//Get the name of the table to insert the new row in
						var sTableToInsertIn = oRefElement.tablename;
					
						var sRowType = oElement.rowtype;
						var sRowTypeDesc = getRowTypeDesc(sRowType);
						var sTableName = oRefElement.tablename;
						var oTable = oDoc.tableByName(sTableName);
						var sTableType = "";
						var sTableTypeDesc = "";
						if(oTable)
						{
							//Get the table type
							var sTableType = oTable.propGet(CTABLETYPE);
							//get table type description
							var sTableTypeDesc = getTableTypeDesc(sTableType);
						}						
						node.setAttribute("mapcolumn",oRefElement.mapcolumn);
						node.setAttribute("tablename",sTableName);
						node.setAttribute("jumpcode",oRefElement.jumpcode);
						node.setAttribute("objecttype","row");
						node.setAttribute("component","row");
						node.setAttribute("title","Component: Table row\nRow type: "+sRowType+"\nRow type desc: "+sRowTypeDesc+"\nTable name: "+sTableName+"\nTable type: "+sTableType+"\nTable desc: "+sTableTypeDesc);						
						node.setAttribute("add","true");
						node.setAttribute("rowtype",sRowType);
						var ID = createGuid();
						var sTempId = "row_"+ID;
						node.setAttribute("id",sTempId);
						node.setAttribute("guid","");
						node.setAttribute("name",sElementName);
						node.setAttribute("refelementid",oRefElement.getAttribute("guid"));
						if(oRefElement.use=="tablecontainer")
						{
							//insert the new element in the html dialog
							var oNewElement = oRefElement.appendChild(node);
							node.setAttribute("refelementid","");							
						}else{
							//insert the new element in the html dialog
							//var oNewElement = oRefParent.insertBefore(node,oRefElement);
							var oNewElement = insertAfter(node, oRefElement,oRefParent);
							if(oNewElement)
								highlightElement(oNewElement);
						}					
						oProgBar.updateProgress(1);	
						oProgBar.setMessage("Row "+x+" / "+iRowToInsert);						
					}
					//Destroy the progress bar
					oProgBar.destroyProgressBar();
					//reload controls on the UI to keep HTML and CV in synch 
					//onLoad();
				}
			}
		}
	}catch(e)
	{
		alert(e.description);
	}finally{
		if(oElement)
		{
			oElement.style.color="";
			oElement.style.position = 'static';			
			oElement.onmousemove = null;
			oElement.onmouseup = null;
		}
		if(oRefElement)
		{
			oRefElement.style.textDecoration = "none";
		}
		oElement = null;
		oRefElement=null;		
	}
}

function insertAfter(oNewNode, oReferenceNode)
{
	try{
		var oNewElement = oReferenceNode.parentNode.insertBefore(oNewNode, oReferenceNode.nextSibling);
		return oNewElement;
	}catch(e)
	{
		alert(e.description);
	}finally{
		
	}
}

function insertBefore(oNewNode, oReferenceNode,oParentNode)
{
	try{
		var oNewElement = oParentNode.insertBefore(oNewNode, oReferenceNode);
		return oNewElement;
	}catch(e)
	{
		alert(e.description);
	}finally{
		
	}	
}

function updateCurrentSelection(oCurrentElement)
{
	try{
		//debugger;
		//debugger;
		var sId = "";
		//Get the component type
		sObjectType = oCurrentElement.objecttype;
		switch(sObjectType)
		{
			case "section":
				sId = oCurrentElement.jumpcode;
				sCurrentTable = "";	
				break;
			case "row":
				sId = oCurrentElement.guid;
				sCurrentTable = oCurrentElement.tablename;
				break;
			default:
		}
		sCurrentlyElementSelected = sId;
	}catch(e)
	{
		alert(e.description);
	}finally
	{
		oElement = null;
	}
}

//Stolen from the web - http://simonwillison.net/2006/Jan/20/escape/
function escapeSpecialCharacterInString(sText)
{

	try
	{
		//Create method that will handle special characters
		RegExp.escape = function(text){
		if (!arguments.callee.sRE) {
			var specials = [
			'/', '.', '*', '+', '?', '|',
			'(', ')', '[', ']', '{', '}', '\\'
			];
			arguments.callee.sRE = new RegExp(
			'(\\' + specials.join('|\\') + ')', 'g'
			);
		}
		  return text.replace(arguments.callee.sRE, '\\$1');
		}
		var sNewText = RegExp.escape(sText)
		return sNewText
	}
	catch(e)
	{
		//debugger
		alert(e.description);
	}    
}


function deleteItem()
{
	try{
		//debugger;
		//debugger;
		var bResult = false;
		var iSelectedItems = aSelectedItems.length;
		//Check if multiple items need to be deleted
		if(iSelectedItems==0){
			alert("You have not selected any item!");
			return;
		}else if(iSelectedItems>1){
			//Iif deleting multiple items ask
			bResult = confirm("Are you sure you want to delete these "+iSelectedItems+" items?");
		}else if(iSelectedItems==1)
		{
			//if deleting a single item do not ask
			bResult = true;
		}
		
		if(bResult){
			var oProgBar = oDoc.createProgressBar("Deleting...",iSelectedItems,1);
			for(var i=0;i<iSelectedItems;i++)
			{
				var oElementToDelete = aSelectedItems[i];
				if(oElementToDelete && oElementToDelete.getAttribute("isheader")=="true")
				{	
					var sIdOfElementToDelete = oElementToDelete.getAttribute("id").split("_")[1];
					var oParentElementToDelFrom = document.getElementById(sIdOfElementToDelete).parentNode;//.id
					if(isInputValid(oParentElementToDelFrom))
					{
						//oParentElementToDelFrom.removeChild(document.getElementById(sIdOfElementToDelete));
						document.getElementById(sIdOfElementToDelete).style.display="none";
						//Set an attribute to show that it needs to be deleted
						document.getElementById(sIdOfElementToDelete).setAttribute("delete","true");
						
					}
				}
				else if(isInputValid(oElementToDelete) && oElementToDelete.getAttribute("ismainsection")=="true"){
					//if main section destroy the whole structure
					var oContainer = document.getElementById("MainContentContainer");
					//oContainer.innerHTML = "";
					oElementToDelete.parentNode.style.display="none";
					//Set an attribute to show that it needs to be deleted
					oElementToDelete.parentNode.setAttribute("delete","true");
				}
				else if(isInputValid(oElementToDelete))
				{
					var oParentElementToDelFrom = oElementToDelete.parentNode;
					if(isInputValid(oParentElementToDelFrom))
					{
						//oParentElementToDelFrom.removeChild(oElementToDelete);
						oElementToDelete.style.display="none";
						//Set an attribute to show that it needs to be deleted
						oElementToDelete.setAttribute("delete","true");						
					}
				}
				oProgBar.updateProgress(1);	
			}
			oProgBar.destroyProgressBar();
			aSelectedItems = [];			
		}		
	}catch(e)
	{
		alert(e.description);
		//alert(e.description);
	}finally
	{
		
	}	
}

function saveProject()
{
	try{
		
		//debugger;
		//debugger;
		
		//var sContent = document.getElementById("MainContentContainer").innerHTML;
		//Save as html file
		var sPath = oDoc.interpret("clntdir()") + "AuthToolProject.html";
		//var sPath = "C:\\AuthToolProject.html"
		var sContent = document.documentElement.outerHTML;
		createTextFile(sPath,sContent);
		
		alert("Projet saved in the following location\n"+sPath);
	}catch(e)
	{
		alert(e.description);
	}finally{
		
	}
}

function createTextFile(sPath,sContent)
{
	
	var ofso = new ActiveXObject("Scripting.FileSystemObject");
	if(ofso.FileExists(sPath)==true){
		var ofile = ofso.OpenTextFile(sPath, 2, true)
	}else{
		
		ofile = ofso.CreateTextFile(sPath, 2, true)
	}
	
	if(ofile)
	{
		ofile.writeline(sContent);
	}
}

function getElementsByClassName1(node, classname) 
{
    var a = [];
    var re = new RegExp('(^| )'+classname+'( |$)');
    var els = node.getElementsByTagName("*");
    for(var i=0,j=els.length; i<j; i++)
        if(re.test(els[i].className))a.push(els[i]);
    return a;
}

//Sets attributes on html dom elements
function addElementAttribute(sHTMLElement,sAttributeName,sAttributeValue)
{
	try{
		if(isInputValid(sHTMLElement) && isInputValid(sAttribute))
		{
			var oHTMLElement = document.getElementById(sHTMLElement);
			if(oHTMLElement)
			{
				oHTMLElement.setAttribute(sAttributeName,sAttributeValue);
			}
		}
	}catch(e)
	{
		alert(e.description);
	}finally{
		oHTMLElement = null;
	}
}

function createGuid(){
   return 'xxxxxxxx-xxxx-4xxx-yxxx-xxxxxxxxxxxx'.replace(/[xy]/g, function(c) {  
      var r = Math.random()*16|0, v = c === 'x' ? r : (r&0x3|0x8);  
      return v.toString(16);  
})}

function GetColnumber(iCol)
{
	//A=65
	var iStart = 65
	if (iCol<=26)
	{
		var sCol=String.fromCharCode(64 + iCol*1)
	}
	if (iCol==26)
	{
		return sCol;
	}
	var iTableRepeats=iCol/26;
	if (iCol%26==0)
	{
		iTableRepeats -= 1
	}
	iTableRepeats =	Math.floor(iTableRepeats);
	if (iTableRepeats>0)
	{
		//Get the first alpha char in the col name
		var sColInit = String.fromCharCode(65+(iTableRepeats-1))
		//Get the second alpha char in the col name
		//get the amount of cols over 26 to remove from the calc
		var iExcess = (iTableRepeats)*26
		iCol = iCol -iExcess - 1
		sColSecond = String.fromCharCode(65 + (iCol*1))
		//concatenate
		sCol = sColInit + sColSecond
	}
	return sCol;
}

//return the equivalent integer value of a letter e.g. A = 1, B= 2 etc
function getIntFromColLabel(sCol)
{
	if (isInputValid(sCol))
	{
		var iString = sCol.length
		sCol = sCol.toUpperCase()
		if (iString>0)
		{
			if (iString>1)
			{
				var iCharCode = parseInt(sCol.charCodeAt(0)-64)*26
				var iCol = iCharCode+sCol.charCodeAt(1)-64
				return iCol
			}
			else
			{
				var iCol = parseInt(sCol.charCodeAt(0))-64
				return iCol
			}		
		}
	}
}
/*
	Checks if data passed in is valid
	i.e. not null, undefined or empty string
 */
function isInputValid(vInput)
{
	try{
		bTrue = true;
		
		if(vInput===""||typeof(vInput)==="undefined"||vInput===null)
			bTrue = false;
		
		return bTrue;		
	}catch(e)
	{
		alert(e.description);
	}
}

function platformInfo()
{
	var sCWMajorVersion = oCWApplication.majorVersion;
	var sCWMinorVersion = oCWApplication.minorVersion;
    var sCWBuildVersion = oCWApplication.buildVersion;
    var sFullFilePath = oDoc.cwClient.FileName;
	//var sIEMajorVerion = ScriptEngineMajorVersion();
	//var sIEMinorVerion = ScriptEngineMinorVersion();	

	// on IE10 x86 platform preview running in IE7 compatibility mode on Windows 7 64 bit edition
	platform.name; // 'IE'
	platform.version; // '10.0'
	platform.layout; // 'Trident'
	platform.os; // 'Windows Server 2008 R2 / 7 x64'
	platform.description; // 'IE 10.0 x86 (platform preview; running in IE 7 mode) on Windows Server 2008 R2 / 7 x64'
	var sPropertyName = "CWClientName";
	var sClientName = getMetaData(sFullFilePath, sPropertyName);
	var sPropertyName = "CWCustomProperty.Product";
	var sProductName = getMetaData(sFullFilePath, sPropertyName);
	var sPropertyName = "CWUserFriendlyFileVersion";
	var sClientFileVer = getMetaData(sFullFilePath, sPropertyName);
	var sPropertyName = "CWUserFriendlyFileVersion";
	var sClientFileVer = getMetaData(sFullFilePath, sPropertyName);	
	var sPropertyName = "CWCustomProperty.Mapping";
	var sMappingVersion = getMetaData(sFullFilePath, sPropertyName);		
	var sTitle = "About";
    var sContent = "Feature: Authoring tool\nVersion: 2 (Beta)\nDescription: Edit & Build Products\n\nCaseWare\nMajor Version: " + sCWMajorVersion + "\nMinor Version: " + sCWMinorVersion + "\nBuild Version: " + sCWBuildVersion + "\n\nClient File\nClient name: "+sClientName+"\nProduct: "+ sProductName +"\nVersion: " + sClientFileVer+"\nMapping version: "+ sMappingVersion + "\nFile path: "+sFullFilePath+"\n\nPlatform: "+platform.description;
	myAlert(sTitle, 64,sContent);
}

function getMetaData(sFilePath, sPropertyName)
{
    var sMetaValue = "";

	var oCWApp = new ActiveXObject("CaseWare.Application")
	try
	{
		var oMetaData = oCWApp.Clients.GetMetaData(sFilePath)
		if(oMetaData.Exists(sPropertyName))
			var sMetaValue = oMetaData.item(sPropertyName).value;
		else{
			if(oMetaData.Exists("CWCustomProperty."+sPropertyName))
				sMetaValue = oMetaData.item("CWCustomProperty."+sPropertyName).value;
		}
		oCWApp = null;				  
	}
	catch(e)
	{
		alert(e.description);
	}finally{
		oCWApp = null;
		oMetaData = null;
	}

	return sMetaValue;
}
/*
function convertLetterToNumber(str) {
  var out = 0, len = str.length;
  for (pos = 0; pos < len; pos++) {
    out += (str.charCodeAt(pos) - 64) * Math.pow(26, len - pos - 1);
  }
  return out;
}

var testCase = ["A","B","C","Z","AA","AB","BY"];

var converted = testCase.map(function(obj) {
  return {
    letter: obj,
    number: convertLetterToNumber(obj)
  };

});

*/
//Handles on mouse up events
function mouseUp()
{
	//debugger;
	//debugger;
	if(oRowElement)
	{
		var oElement = oRowElement;
	}
	
	if(oSectionElement)
	{
		var oElement = oSectionElement;
	}
	
	if(oTableElement)
	{
		var oElement = oTableElement;
	}

	if(oParaElement)
	{
		var oElement = oParaElement;
	}
	
	try{
		//debugger;
		//debugger;
		if (oElement && isInputValid(sRefElement)) {
			//Check if the element need to be added or note
			if(oElement.style.color=="green")
			{		
				var sObjectType = oElement.getAttribute("objecttype");
				switch(sObjectType)
				{
					case "row":
						InsertNewRow(oElement,sRefElement);
						break;
					case "section":
						InsertNewSection(oElement,sRefElement);
						break;
					case "table":
						InsertNewTable(oElement,sRefElement);
						break;
					case "para":
						InsertNewParagraph(oElement,sRefElement);
						break;
					case "page":
						//insertNewPage(oElement,sRefElement);
						break;						
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
		oSectionElement = null;
		oRowElement = null;
		oTableElement = null;
		oParaElement = null;
		oElement = null;
		oRefElement=null;
		sRefElement=null;
		document.body.style.cursor="default";		
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

function insertAsSub(oNewNode, oReferenceNode)
{
	try{
		var oNewElement = oReferenceNode.appendChild(oNewNode)//, oReferenceNode.nextSibling);
		return oNewElement;
	}catch(e)
	{
		alert(e.description);
	}finally{
		
	}
}

function insertTable(oNewNode, oReferenceNode)
{
	try{
		var oNewElement = oReferenceNode.appendChild(oNewNode)//, oReferenceNode.nextSibling);
		return oNewElement;
	}catch(e)
	{
		alert(e.description);
	}finally{
		
	}
}

function insertPara(oNewNode, oReferenceNode)
{
	try{
		var oNewElement = oReferenceNode.appendChild(oNewNode)//, oReferenceNode.nextSibling);
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
					//var oParentElementToDelFrom = document.getElementById(sIdOfElementToDelete).parentNode;//.id
					//if(isInputValid(oParentElementToDelFrom))
					//{
						//oParentElementToDelFrom.removeChild(document.getElementById(sIdOfElementToDelete));
						//document.getElementById(sIdOfElementToDelete).style.display="none";
						oElementToDelete.style.textDecoration="line-through";//.innerHTML = "<strike style='color:red' title='Will be deleted'>"+oElementToDelete.innerHTML+"</strike>";
						oElementToDelete.style.color="red";//.setAttribute("text-decoration-color","red");
						//Set an attribute to show that it needs to be deleted
						document.getElementById(sIdOfElementToDelete).setAttribute("delete","true");
						document.getElementById(oElementToDelete.getAttribute("id")).setAttribute("delete","true");
						aDeletedItems[aDeletedItems.length] = oElementToDelete;
						var aChildElements = oElementToDelete.parentNode.getElementsByTagName("LI");
						for(var j=0;j<aChildElements.length;j++){
							aChildElements[j].style.textDecoration="line-through";//.innerHTML = "<strike style='color:red' title='Will be deleted'>"+aChildElements[j].innerHTML+"</strike>";
							aChildElements[j].style.color="red";//.setAttribute("text-decoration-color","red");
							aChildElements[j].setAttribute("delete","true");
						}
						//highlightElement(oElementToDelete);	
						//oElementToDelete.disabled=true;
					//}
				}
				else if(isInputValid(oElementToDelete) && oElementToDelete.getAttribute("ismainsection")=="true"){
					//if main section destroy the whole structure
					//var oContainer = document.getElementById("MainContentContainer");
					//oContainer.innerHTML = "";
					//oElementToDelete.parentNode.style.display="none";
					oElementToDelete.parentNode.style.textDecoration="line-through";//.innerHTML = "<strike style='color:red' title='Will be deleted'>"+oElementToDelete.parentNode.innerHTML+"</strike>";
					aDeletedItems[aDeletedItems.length] = oElementToDelete.parentNode;
					oElementToDelete.style.color="red";//.setAttribute("text-decoration-color","red");
					//Set an attribute to show that it needs to be deleted
					oElementToDelete.parentNode.setAttribute("delete","true");
					oElementToDelete.setAttribute("delete","true");
					//highlightElement(oElementToDelete);	
					//oElementToDelete.parentNode.disabled=true;
					var aChildElements = oElementToDelete.parentNode.getElementsByTagName("LI");
					for(var j=0;j<aChildElements.length;j++)
					{
						aChildElements[j].style.textDecoration="line-through";//.innerHTML = "<strike style='color:red' title='Will be deleted'>"+aChildElements[j].innerHTML+"</strike>";	
						aChildElements[j].style.color="red";//.setAttribute("text-decoration-color","red");
						aChildElements[j].setAttribute("delete","true");
					}
				}
				else if(isInputValid(oElementToDelete))
				{
					//var oParentElementToDelFrom = oElementToDelete.parentNode;
					//if(isInputValid(oParentElementToDelFrom))
					//{
						//oParentElementToDelFrom.removeChild(oElementToDelete);
						//oElementToDelete.style.display="none";
						oElementToDelete.style.textDecoration="line-through";//.innerHTML = "<strike style='color:red' title='Will be deleted'>"+oElementToDelete.innerHTML+"</strike>";
						oElementToDelete.style.color="red";//setAttribute("text-decoration-color","red");
						//Set an attribute to show that it needs to be deleted
						oElementToDelete.setAttribute("delete","true");
						aDeletedItems[aDeletedItems.length] = oElementToDelete;						
						//highlightElement(oElementToDelete);
						//oElementToDelete.disabled=true;						
					//}
				}
				oProgBar.updateProgress(1);	
			}
			oProgBar.destroyProgressBar();
			aSelectedItems = [];
			//Clear preview pane
			//document.getElementById("editorPane").innerHTML = "";
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

function copyElement()
{
	try{
		//debugger;
		//debugger;
		//Check if there is any element or elements that need to be copied in the array that keeps selected elements
		var iLength = aSelectedItems.length;
		if(iLength>0)
		{
			for(var i=0;i<iLength;i++)
			{
				var oElementToCopy = aSelectedItems[i];
				if(oElementToCopy.component=="section"){
					var oDuplicatNode = oElementToCopy.parentNode.cloneNode(true);
					oDuplicatNode.id = "UL_"+createGuid();
					//remember to add generic code to change main section, child elements and attributes
					aCopiedItems[aCopiedItems.length] = oDuplicatNode;
				}
				else
				{
					var oDuplicatNode = oElementToCopy.cloneNode(true);
					//change the id of the node
					oDuplicatNode.id = "row_"+createGuid();
					//oDuplicatNode.jumpcode = oDuplicatNode.jumpcode;
					aCopiedItems[aCopiedItems.length] = oDuplicatNode;
				}
			}
		}else{
			alert("You have not selected an item to copy.")
		}		
	}catch(e)
	{
		alert(e.description);
	}finally{
		
	}
}

function pasteElement()
{
	try{
		//debugger;
		//debugger;
		//Check if there is any element or elements that need to be copied in the array that keeps selected elements
		var iLength = aCopiedItems.length;
		if(iLength>0)
		{
			for(var i=0;i<iLength;i++)
			{
				var oElementToPaste = aCopiedItems[i];
				var oItemToInsertAfter = aSelectedItems[0];
				if(oElementToPaste.component=="section")
					oItemToInsertAfter = oItemToInsertAfter.parentNode;

				var oNewElement = insertAfter(oElementToPaste, oItemToInsertAfter);
				if(oNewElement)
					highlightElement(oNewElement);
			}
		}else{
			alert("There are no items to paste");
		}		
	}catch(e)
	{
		alert(e.description);
	}finally{
		
	}	
}

function undoLastActionOnElement()
{
	try{
		//debugger;
		//debugger;
		var iLength = aDeletedItems.length;
		if(iLength)
		{
			for(var i=0;i<iLength;i++)
			{
				aDeletedItems[i].style.textDecoration = "";
				var aChildElements = aDeletedItems[i].getElementsByTagName("LI");
				for(var j=0;j<aChildElements.length;j++)
				{
					aChildElements[j].style.textDecoration="";//.innerHTML = "<strike style='color:red' title='Will be deleted'>"+aChildElements[j].innerHTML+"</strike>";	
					aChildElements[j].style.color="initial";//.setAttribute("text-decoration-color","red");
					aChildElements[j].removeAttributeAttribute("delete");//,"true");
				}				
			}
		}
	}catch(e)
	{
		alert(e.description);
	}finally{
		
	}
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
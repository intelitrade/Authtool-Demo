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
						//oElementToDelete.style.textDecoration="line-through";//.innerHTML = "<strike style='color:red' title='Will be deleted'>"+oElementToDelete.innerHTML+"</strike>";
						//oElementToDelete.style.color="red";//.setAttribute("text-decoration-color","red");
						//Set an attribute to show that it needs to be deleted
						//oElementToDelete.setAttribute("delete","true");
						//document.getElementById(sIdOfElementToDelete).setAttribute("delete","true");
						oElementToDelete.parentNode.setAttribute("delete","true");
						//document.getElementById(oElementToDelete.getAttribute("id")).setAttribute("delete","true");
						aDeletedItems[aDeletedItems.length] = oElementToDelete;
						var aChildElements = oElementToDelete.parentNode.getElementsByTagName("LI");
						for(var j=0;j<aChildElements.length;j++){
							aChildElements[j].style.textDecoration="line-through";//.innerHTML = "<strike style='color:red' title='Will be deleted'>"+aChildElements[j].innerHTML+"</strike>";
							//aChildElements[j].style.color="red";//.setAttribute("text-decoration-color","red");
							aChildElements[j].setAttribute("delete","true");
							aDeletedItems[aDeletedItems.length] = aChildElements[j];
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
					//oElementToDelete.style.color="red";//.setAttribute("text-decoration-color","red");
					//Set an attribute to show that it needs to be deleted
					oElementToDelete.parentNode.setAttribute("delete","true");
					oElementToDelete.setAttribute("delete","true");
					//highlightElement(oElementToDelete);	
					//oElementToDelete.parentNode.disabled=true;
					var aChildElements = oElementToDelete.parentNode.getElementsByTagName("LI");
					for(var j=0;j<aChildElements.length;j++)
					{
						aChildElements[j].style.textDecoration="line-through";//.innerHTML = "<strike style='color:red' title='Will be deleted'>"+aChildElements[j].innerHTML+"</strike>";	
						//aChildElements[j].style.color="red";//.setAttribute("text-decoration-color","red");
						aChildElements[j].setAttribute("delete","true");
						aDeletedItems[aDeletedItems.length] = aChildElements[j];
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
						//oElementToDelete.style.color="red";//setAttribute("text-decoration-color","red");
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
			highlightElement(aSelectedItems[aSelectedItems.length-1]);
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

		var iLength = aSelectedItems.length;
		if(iLength>0)
		{
			for(var i=0;i<iLength;i++)
			{
				if(aSelectedItems[i].getAttribute("delete")=="true")
				{
					aSelectedItems[i].style.textDecoration="";
					aSelectedItems[i].removeAttribute("delete");

					if(aSelectedItems[i].getAttribute("ismainsection")=="true"||aSelectedItems[i].getAttribute("isheader")=="true")
						var aChildElements = aSelectedItems[i].parentNode.getElementsByTagName("LI");
					else
						var aChildElements = aSelectedItems[i].getElementsByTagName("LI");
						
					for(var j=0;j<aChildElements.length;j++)
					{
						aChildElements[j].style.textDecoration="";
						aChildElements[j].removeAttribute("delete");//,"true");	
					}
				}
			}		
		}
	}catch(e)
	{
		alert(e.description);
	}finally{
		
	}
}

function toTitleCase(str) {
	return str.replace(
		/\w\S*/g,
		function(txt) {
			return txt.charAt(0).toUpperCase() + txt.substr(1).toLowerCase();
		}
	);
}

function convertLetterToNumber(str) {
  var out = 0, len = str.length;
  for (pos = 0; pos < len; pos++) {
    out += (str.charCodeAt(pos) - 64) * Math.pow(26, len - pos - 1);
  }
  return out;
}

function sendToPrinter(content)
{
	try{
		//debugger;
		//debugger;
		var content = "editorPane";
		var divName = content;
		var printContents = document.getElementById(divName).innerHTML;
		var originalContents = document.body.innerHTML;

		document.body.innerHTML = printContents;

		window.print();

		document.body.innerHTML = originalContents;		
	}catch(e)
	{
		alert(e.description);
	}finally{
		
	}
}

function sortElements(sSortType)
{
	try{
		/*debugger;
		debugger;
		showMappingDialog();
		return;*/
		//Check if any item has been selected
		if(aSelectedItems.length==1)
		{
			var oElement = aSelectedItems[0];
			if(oElement.isheader=="true")
				oElement = oElement.parentNode;
			
			//Get the parent node
			if(oElement.parentNode && oElement.previousSibling && sSortType=="up"){
				
				if(oElement.previousSibling.isheader=="true")
					return;
				
				insertBefore(oElement, oElement.previousSibling,oElement.parentNode);
			
			}else if(oElement.parentNode && oElement.nextSibling && sSortType=="down"){
			
				insertAfter(oElement, oElement.nextSibling);
			
			}
		}
	}catch(e)
	{
		alert(e.description);
	}finally{
		
	}
}

function getAllMapNoAndDesc()
{
	try{
	
		//Get the client object
		var cwClient = oDoc.cwClient;
		//Get the Mappings collection
		var mappings = cwClient.mappings;
		var oEnumerator = new Enumerator(mappings);
		
		var iLength = mappings.Count;
		var oProgBar = oDoc.createProgressBar("Loading mapping...",iLength,1);	
		//enumerate the mapping database
		var aMappingList = [];
		for (;!oEnumerator.atEnd(); oEnumerator.moveNext())
		{
			//Get an item from the collection
			var mapItem = oEnumerator.item();
			var sMapNoFromMappingDbase = mapItem.ID;
			if (!isInputValid(sMapNoFromMappingDbase))
				continue;
					
			if(isInputValid(mapItem.Name))
			{
				var sMapDesc = mapItem.Name;
				if(sMapDesc.search("(Filtered)")!=-1||sMapDesc.search("N/A")!=-1)
					continue;
			}	
					
			aMappingList[aMappingList.length] = [mapItem.ID, mapItem.Name];
			oProgBar.updateProgress(1);
		}
		oProgBar.destroyProgressBar();
		aMappingList.sort();
		var iLength = aMappingList.length;
		sStr = "";
		var oProgBar = oDoc.createProgressBar("Creating mapping list...",iLength,1);
		for(var i=0;i<iLength;i++)
		{
			sStr = sStr+"<li mapno='"+aMappingList[i][0]+"' description='"+aMappingList[i][1]+"' style='width:100%' onclick='highlightElement(this)'><table style='border:1px solid black;width:100%'><tr><td>"+aMappingList[i][1]+"</td><td style='width:30%;align:left'>"+aMappingList[i][0]+"</td></tr></table></li>";
			oProgBar.updateProgress(1);
			//sStr = sStr+"<li mapno='"+aMapping[i][0]+"' description='"+aMapping[i][1]+"' style='width:100%'>"+aMapping[i][1]+" | "+aMapping[i][0]+"</li>";
		}
		oProgBar.destroyProgressBar();
		return "<ul id='mappinglist'><li style='width:100%'><table style='border:1px solid black;width:100%;font-weight:bold'><tr><td>Description</td><td style='width:30%;align:left'>Map no.</td></tr></table></li>"+sStr+"</ul>";
	}catch(e)
	{
		alert(e.description);
	}finally{
		cwClient = null;
		//Get the Mappings collection
		mappings = null;
		oEnumerator = null;
		//enumerate the mapping database
		aMapping = null;
		oProgBar = null;		
	}
}

function showMappingDialog(sElement)
{
	//debugger;
	//debugger;
	try{
		
		var sHTMLPath = oCWApplication.programPath+"Scripts\\SA IFRS\\HTML\\authtool\\mappingdialog.html";
		if(sMappingString=="")
			sMappingString = getAllMapNoAndDesc();
		//launch custom script interface
		var sResult = oCWApplication.runHTMLDialog2(sHTMLPath, 1, oDoc,oCWApplication,sMappingString);
		//debugger;
		//debugger;		
		if(isInputValid(sResult))
		{	
			var oElement = document.getElementById(sElement);
			oElement.value=sResult;		
			updateMapNo(oElement);
		}
		//debugger;
		//debugger;
	}catch(e)
	{
		alert(e.description);
	}finally{
		
	}
	
	/*var ctxMenu = document.getElementById("ctxMenu");
	ctxMenu.style.display = "block";
	ctxMenu.style.left = event.clientX - 5+"px";//ctxMenu.offsetWidth//"300px"//event.clientX - 10+"px";
	ctxMenu.style.top = event.clientY - 5+"px";//ctxMenu.offsetHeight//"10px"//event.clientY - 10+"px";

	if(!document.getElementById("mappinglist"))
	{
		var sMappingList = getAllMapNoAndDesc();
		if(isInputValid(sMappingList))
			ctxMenu.innerHTML = sMappingList;
	}*/
}

function searchUnorderedList() {
  // Declare variables
  /*var input, filter, ul, li, a, i, txtValue;
  */
  //debugger;
  //debugger;

  try{
	var oInput = document.getElementById('searchbox');
	var sFilter = oInput.value.toUpperCase();
	var oMappingList = document.getElementById("mappinglist");
	var aListItem = oMappingList.getElementsByTagName('li');
	var iLength = aListItem.length;
	var oProgBar = oDoc.createProgressBar("Search...",iLength,1);	
	// Loop through all list items, and hide those who don't match the search query
	for (var i = 0; i < iLength; i++) 
	{
		var el = oItem = aListItem[i];//.getElementsByTagName("a")[0];
		//var txtValue = oItem.mapno || oItem.description;
		if(!isInputValid(oItem.mapno))
			continue;
		
		if (oItem.mapno.toUpperCase().search(sFilter)!=-1||oItem.description.toUpperCase().search(sFilter)!= -1) {
			//aListItem[i].style.display = "";
			//aListItem[i].setAttribute("class")
			removeClass(el, "testit");
		} else {
			//aListItem[i].style.display = "none";
			addClass(el, "testit");
			
		}
		oProgBar.updateProgress(1);
	}
	oProgBar.destroyProgressBar();
  }catch(e)
  {
	alert(e.description);
  }finally{
	  
  }
  
}

function addClass(el, classNameToAdd){
    el.className += ' ' + classNameToAdd;   
}

function removeClass(el, classNameToRemove){
    var elClass = ' ' + el.className + ' ';
    while(elClass.indexOf(' ' + classNameToRemove + ' ') !== -1){
         elClass = elClass.replace(' ' + classNameToRemove + ' ', '');
    }
    el.className = elClass;
}

function sortList(b) {
	//debugger;
	//debugger;
	var list, i, switching, b, shouldSwitch;
	list = document.getElementById("editorPane");
	switching = true;
	/* Make a loop that will continue until
	no switching has been done: */
	while (switching) {
	// Start by saying: no switching is done:
	switching = false;
	//b = list.getElementsByTagName("LI");
	// Loop through all list items:
	for (i = 0; i < (b.length - 1); i++) {
	  // Start by saying there should be no switching:
	  shouldSwitch = false;
	  /* Check if the next item should
	  switch place with the current item: */
	  //if (b[i].innerHTML.toLowerCase() > b[i + 1].innerHTML.toLowerCase()) {
		if (b[i].innerHTML > b[i + 1].innerHTML) {
		/* If next item is alphabetically lower than current item,
		mark as a switch and break the loop: */
		shouldSwitch = true;
		break;
	  }
	}
	if (shouldSwitch) {
	  /* If a switch has been marked, make the switch
	  and mark the switch as done: */
	  b[i].parentNode.insertBefore(b[i + 1], b[i]);
	  switching = true;
	}
	}
}

function exporttoexcel(sTable,sExcelSaveFilePath,oMyDoc)
{
	try{
		debugger;
		debugger;
        var Excel = new ActiveXObject("Excel.Application");
        //Excel.Visible = true;
		var sExcelSaveFilePath = "C:\\Program Files (x86)\\CaseWare\\Scripts\\SA IFRS\\HTML\\authtool\\test.xlsx";
		var sTable = "MFG";
        var oExcelWorkBook = Excel.Workbooks.Add();//Open("C:\\Program Files (x86)\\CaseWare\\Scripts\\SA IFRS\\HTML\\authtool\\test.xlsx");
		if(oExcelWorkBook)
		{
			var oCaseViewTable = oDoc.tableByName(sTable);
			//get the table
			populateExcelWorkSheet(oExcelWorkBook,oCaseViewTable)
			//oExcelWorkBook.ActiveSheet.Cells(1,1).Value = "Test";
			oExcelWorkBook.SaveAs(sExcelSaveFilePath);
			Excel.Visible = true;
			//Excel.Quit();	// Close Excel with the Quit method on the Application object.
		}	
	}catch(e)
	{
		alert(e.description);
	}finally{
	}
}

function populateCVTableExcelWorkSheet(oExcelWorkBook,oCaseViewTable)
{
	try{
			if(oExcelWorkBook && oCaseViewTable)
			{
				var iRows = oCaseViewTable.nRows();
				var iColumns = oCaseViewTable.nColumns();				

				var iRows = oCaseViewTable.nRows();
				var iColumns = oCaseViewTable.nColumns();
				var oProgBar = oDoc.createProgressBar("Exporting data...",iRows,1);
				for(var i = 1; i <= iRows; i++)
				{
					for(var j = 1; j <= iColumns; j++)
					{
						//get the table cell in CaseView Table
						//check if there are any cells in the table
						var iTableCellFirstPara = oCaseViewTable.cellFirstParaIndex(i, j);
						//var iTableCellLastPara = oCVTable.cellLastParaIndex(i, j);						
						var oPara = oDoc.para(iTableCellFirstPara);
						var sText = oPara.getText();
						if(!isInputValid(sText))
							continue;	

						oExcelWorkBook.ActiveSheet.Cells(i,j).Value = sText;
					}
					oProgBar.updateProgress(1);
				}
				oProgBar.destroyProgressBar();
			}			
	}catch(e)
	{
		alert(e.description);
	}finally{
		
	}
}

function exportCaseViewDataToExcel()
{
	try{
		debugger;
		debugger;
		
		var oSection = oDoc.sectionByName(sSectionLabel)//;"NOTES_002");
		var iFirstPara = oSection.firstParaIndex;
		var iLastPara = oSection.lastParaIndex;
		var sStr = "";
		var aTables = [];
        var Excel = new ActiveXObject("Excel.Application");
        //Excel.Visible = true;
		var sExcelSaveFilePath = "C:\\Program Files (x86)\\CaseWare\\Scripts\\SA IFRS\\HTML\\authtool\\test.xlsx";
		//var sTable = "MFG";
        var oExcelWorkBook = Excel.Workbooks.Add();
		var iExcelPointer = 1;
		if(oExcelWorkBook)// && oCaseViewTable)
		{		
			for(var i=iFirstPara;i<=iLastPara;i++)
			{			
				var oProgBar = oDoc.createProgressBar("Exporting data...",iLastPara,1);
				
				var oCaseViewTable = oDoc.tableByParaIndex(i);
				if(oCaseViewTable)
				{
					var iRows = oCaseViewTable.nRows();
					var iColumns = oCaseViewTable.nColumns();					
					for(var k = 1; k <= iRows; k++)
					{
						var bDataExported = false;					
						for(var j = 1; j <= iColumns; j++)
						{
							//get the table cell in CaseView Table
							//check if there are any cells in the table
							var iTableCellFirstPara = oCaseViewTable.cellFirstParaIndex(k, j);
							//var iTableCellLastPara = oCVTable.cellLastParaIndex(i, j);						
							var oPara = oDoc.para(iTableCellFirstPara);
							var sText = oPara.getText();
							if(!isInputValid(sText))
								continue;	
							
							oExcelWorkBook.ActiveSheet.Cells(iExcelPointer,j).Value = sText;
							bDataExported = true;
						}
						if(bDataExported)
							iExcelPointer++;
					}
					i=oCaseViewTable.lastParaIndex;
				}else{
					
					oExcelWorkBook.ActiveSheet.Cells(iExcelPointer,1).Value = oDoc.para(i).getText();
					iExcelPointer++;
				}	
				oProgBar.updateProgress(1);
			}
			oProgBar.destroyProgressBar();
			oExcelWorkBook.SaveAs(sExcelSaveFilePath);
			Excel.Visible = true;
		}
	}catch(e)
	{
		alert(e.description);
	}finally{
	}
}

function exportCaseViewDataToWord()
{
	try{
		debugger;
		debugger;
		
		var oSection = oDoc.sectionByName(sSectionLabel)//;"NOTES_002");
		var iFirstPara = oSection.firstParaIndex;
		var iLastPara = oSection.lastParaIndex;
		var sStr = "";
		var aTables = [];
        var Word = new ActiveXObject("Word.Application");
        //Excel.Visible = true;
		var sWordSaveFilePath = "C:\\Program Files (x86)\\CaseWare\\Scripts\\SA IFRS\\HTML\\authtool\\test11.docx";
		//var sTable = "MFG";
		var oWordDocument = Word.Documents.Add();//Open(sWordSaveFilePath);
		//oWordDocument.Content.insertParagraph("This is the first text with a custom style.","End");
		//Word.Selection.TypeText("Hello world!\n787");
		//oWordDocument.SaveAs(sWordSaveFilePath);
		//Word.Visible = true;
		//debugger;
		//debugger;
		//var iLength = 1000000;
		 //for(var il = 1;il<=iLength;il++)
			// v=il;
		//debugger;
		//debugger
		//return;
		var sText = "";
        //var oWord = Excel.Workbooks.Add();
		var iExcelPointer = 1;
		if(oWordDocument)// && oCaseViewTable)
		{		
			var iFirstPara = 1;
			var iLastPara = oDoc.paraCount();
			var oProgBar = oDoc.createProgressBar("Exporting data...",iLastPara,1);
			for(var k=iFirstPara;k<=iLastPara;k++)
			{							
				var oCaseViewTable = oDoc.tableByParaIndex(i);
				if(oCaseViewTable)
				{
					var oCaseViewTable = oDoc.tableByName("MFG");
					var iRows = oCaseViewTable.nRows();
					var iColumns = oCaseViewTable.nColumns();
					
					var oTable = Word.Selection.Tables.Add(Word.Selection.Range,iRows,iColumns);
					for(var i=1;i<=iRows;i++)
					{
						for(var j=1;j<=iColumns;j++)
						{
							//oDoc.para(oCaseViewTable.cellFirstParaIndex(j, i)).getText();
							oTable.Cell(i, j).Range.InsertAfter(oDoc.para(oCaseViewTable.cellFirstParaIndex(i, j)).getText());
						}
					}					
					/*var iRows = oCaseViewTable.nRows();
					var iColumns = oCaseViewTable.nColumns();					
					for(var k = 1; k <= iRows; k++)
					{
						var bDataExported = false;					
						for(var j = 1; j <= iColumns; j++)
						{
							//get the table cell in CaseView Table
							//check if there are any cells in the table
							var iTableCellFirstPara = oCaseViewTable.cellFirstParaIndex(k, j);
							//var iTableCellLastPara = oCVTable.cellLastParaIndex(i, j);						
							var oPara = oDoc.para(iTableCellFirstPara);
							var sText = oPara.getText();
							if(!isInputValid(sText))
								continue;	
							
							oExcelWorkBook.ActiveSheet.Cells(iExcelPointer,j).Value = sText;
							bDataExported = true;
						}
						if(bDataExported)
							iExcelPointer++;
					}*/
					k=oCaseViewTable.lastParaIndex;
				}else{
					
					
					//Word.Selection.TypeText("\n"+oDoc.para(i).getText());
					//oExcelWorkBook.ActiveSheet.Cells(iExcelPointer,1).Value = oDoc.para(i).getText();
					if(sText=="")
						sText = oDoc.para(i).getText();
					else
						sText = sText + "\n" + oDoc.para(i).getText();
						
					//iExcelPointer++;
				}	
				
				
				oProgBar.updateProgress(1);
			}
			debugger;
			debugger;

           // var oCell = oTable.Cell(1, 1);
            //oCell.Range.InsertAfter("Chim");
			//Word.Selection.TypeText(sText);
			oProgBar.destroyProgressBar();
			oWordDocument.SaveAs(sWordSaveFilePath);
			Word.Visible = true;
		}
	}catch(e)
	{
		alert(e.description);
	}finally{
	}
}

function toggleboard(sElementId)
{
	try{
		//debugger;
		//debugger;
		var oElement = document.getElementById(sElementId);
		if(oElement){
			if(oElement.style.display!="none")
				oElement.style.display="none";
			else
				oElement.style.display="block";
		}
	}catch(e)
	{
		alert(e.description);
	}finally{
	}
}

function sanitize(szInput)  
{ 
   // var szInput = myDiv.innerHTML; 
    var szStaticHTML = toStaticHTML(szInput); 
   /* ResultComment = "\ntoStaticHTML sanitized the HTML fragment as follows:\n" 
        + "Original Content:\n" + szInput + "\n" 
        + "Static Content:\n" + szStaticHTML + "\n"; 
    myDiv.innerText = ResultComment; */
	return szStaticHTML;
}

function copyStringToClipboard (str) {
   // Create new element
   var el = document.createElement('textarea');
   // Set value (string to be copied)
   el.value = str;
   // Set non-editable to avoid focus and move outside of view
   el.setAttribute('readonly', '');
   el.style = {position: 'absolute', left: '-9999px'};
   document.body.appendChild(el);
   // Select text inside element
   el.select();
   // Copy text to clipboard
   document.execCommand('copy');
   // Remove temporary element
   document.body.removeChild(el);
}
/*
var testCase = ["A","B","C","Z","AA","AB","BY"];

var converted = testCase.map(function(obj) {
  return {
    letter: obj,
    number: convertLetterToNumber(obj)
  };

});

*/
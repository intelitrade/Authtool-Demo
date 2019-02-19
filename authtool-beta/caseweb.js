//Build a treeview based on section passed in
function buildTreeViewEx(sSectionLabel,oDoc)//,oProgBar)
{
	//debugger;
	//debugger;
	try{
		var CTABLETYPE = "CTABLETYPE";
		var sColor = "";
		if(isInputValid(sSectionLabel))
		{
			//var oContainer = document.getElementById("MainContentContainer");
			var oContainer = document.getElementById("MainContentContainer");
			var oSection = oDoc.sectionByName(sSectionLabel); 
			
			if(oSection)
			{
				if(oContainer.children.length==0)
				{				
					var sSectionTitle = getSectionName(oDoc,sSectionLabel);
					var sMainSectionId = "S"+oSection.index;
					var oParentNode = document.createElement("UL");
					var oParentListItemNode = document.createElement("LI");
					//var oTextnode = document.createTextNode(sSectionLabel);
					var oTextnode = document.createTextNode(sSectionTitle);
					oParentListItemNode.innerHTML="<SPAN>"+sSectionTitle+"</SPAN>"//.appendChild(oTextnode);
					oParentNode.setAttribute("id",sMainSectionId);
					//oParentListItemNode.setAttribute("style","background-color:#F5DEB3");
					oParentListItemNode.style.backgroundColor = "#F5DEB3";
					oParentListItemNode.style.fontWeight = "bold"
					oParentListItemNode.style.padding="10px";
					
					oParentListItemNode.setAttribute("ismainsection","true");
					
					oParentListItemNode.setAttribute("component","section");
					oParentListItemNode.setAttribute("objecttype","section"); 
					
					oParentNode.appendChild(oParentListItemNode);				
					oContainer.appendChild(oParentNode);
					oParentNode.setAttribute("id",sMainSectionId);
					oParentNode.setAttribute("use","heading");
					oParentListItemNode.setAttribute("jumpcode",sSectionLabel);
					oParentListItemNode.setAttribute("onclick","gotoSection(this);highlightElement(this)");
					oParentListItemNode.onclick = function() { gotoSection();highlightElement();  };					
					oParentNode.setAttribute("title","Component: Section\nSection name: "+sSectionTitle+"\nSection label: "+sSectionLabel+"\nSection index: "+oSection.index()+"\nVersion: "+oSection.propGet("CVERSION")+"\nSection type: "+oSection.propGet("CTYPE"));
					
					oParentNode.setAttribute("typeofsection",oSection.propGet("CTYPE"));
					oParentListItemNode.setAttribute("typeofsection",oSection.propGet("CTYPE")); 
					
					if(document.getElementById("docmapheading"))
						document.getElementById("docmapheading").innerHTML = "Edit options<br>"+sSectionTitle;

					var iSectionSkipCond = oSection.evaluateSkip();
					var iSectionHideCond = oSection.evaluateHide();
					
					if(iSectionSkipCond===1 || iSectionHideCond==1)
						sColor = '#0000FF';
					else
						sColor = '#000000';	
						
					oParentListItemNode.style.color = sColor;

				}
				var aSubSection = getSubsections(sSectionLabel,oDoc)//,iSectionIndex);
				if(isInputValid(aSubSection))
				{
					//oProgBar.setMessage(sSectionTitle);
					for(var i=0;i<aSubSection.length;i++)
					{							
						var oSubSection = aSubSection[i];
						if(isInputValid(oSubSection))
						{
							var iSubSectionId = "S"+oSubSection.index;
							if(document.getElementById(iSubSectionId))
								continue;
							
							var sSubSection = oSubSection.label;
							var aOuterSections = getOuterSections(sSubSection,oDoc);
							var sCtypeVal = oSubSection.propGet("CTYPE");
							if(isInputValid(aOuterSections))
							{
								for(var j=aOuterSections.length;j>=0;j--)
								{
									var sOuterSection = aOuterSections[j];
									if(!isInputValid(sOuterSection))
										continue;
									
									var oOuterSection = oDoc.sectionByName(sOuterSection);
									var sHTMLParentElementId = "S"+oOuterSection.index;
									var oHTMLParentElement = document.getElementById(sHTMLParentElementId);
									if(oOuterSection && oHTMLParentElement)
									{
										var sSectionTitle = getSectionName(oDoc,sSubSection);
										var oChildNode = document.createElement("UL");
										var oChildListItemNode = document.createElement("LI");
										var oChildTextnode = document.createTextNode(sSectionTitle);
										oChildNode.setAttribute("id",iSubSectionId);
										oChildListItemNode.innerHTML="<SPAN>"+sSectionTitle+"</SPAN>"//.appendChild(oChildTextnode);
										oChildNode.appendChild(oChildListItemNode);
										oHTMLParentElement.appendChild(oChildNode);
										oChildNode.setAttribute("typeofsection",oSubSection.propGet("CTYPE"));
										oChildListItemNode.setAttribute("typeofsection",oSubSection.propGet("CTYPE"));
										
										var iSectionSkipCond = oSubSection.evaluateSkip();
										var iSectionHideCond = oSubSection.evaluateHide();
										
										if(iSectionSkipCond===1 || iSectionHideCond==1)
											sColor = '#0000FF';
										else
											sColor = '#000000';
										
										
										if(sCtypeVal=="NOTECONTROL"||sCtypeVal=="SUBNOTECONTROL"||sCtypeVal=="EXPANDCOLLAPSE"||sCtypeVal=="SUBNOTECONTROLGROUP"||sCtypeVal=="EXPANDCOLLAPSECOMPANY3RDYEAR")
										{
											oChildListItemNode.setAttribute("id","HEADER_"+iSubSectionId);
											oChildListItemNode.setAttribute("isheader","true");
											oChildListItemNode.setAttribute("jumpcode",sSubSection);
											oChildListItemNode.setAttribute("onclick","gotoSection(this);highlightElement(this)");
											oChildListItemNode.onclick = function() { gotoSection();highlightElement();  };
	
											oChildNode.style.color=sColor;
											oChildNode.setAttribute("title","Component: Section\nSection name: "+sSectionTitle+"\nSection label: "+sSubSection+"\nSection index: "+oSubSection.index()+"\nVersion: "+oSubSection.propGet("CVERSION")+"\nSection type: "+oSubSection.propGet("CTYPE"));
											oChildListItemNode.style.backgroundColor = "#F5DEB3";
											oChildListItemNode.style.fontWeight = "bold"
											oChildListItemNode.style.padding="10px";

											oChildListItemNode.setAttribute("component","section");
											oChildListItemNode.setAttribute("objecttype","section");
											oChildNode.setAttribute("use","sectioncontainer");
											
										}
										
										if(sCtypeVal=="SUMMARYTABLE"||sCtypeVal=="STD6COLUMNTABLE"||sCtypeVal=="RECONTABLE"||sCtypeVal=="SUMMARYTABLEGROUP"||sCtypeVal=="HIDDENTABLEWHICHPULLSTOBS"||sCtypeVal=="RADIOBUTTONTABLE"||sCtypeVal=="GUIDANCETABLE"||sCtypeVal=="DEPMETHODSANDUSEFULLLIFETYPE1")
										{
											//debugger;
											//debugger;
											
											//Get table in section
											var sStr = getTableData(oDoc,oSubSection);
											if(sStr=="")
												continue;
											var sTableTypeDesc = getTableTypeDesc(sCtypeVal);
											oChildNode.innerHTML = sStr;
											oChildNode.setAttribute("title","Component: Section\nSection name: "+sSectionTitle+"\nSection label: "+sSubSection+"\nSection GUID: "+oSubSection.propGet("GUID")+"\nVersion: "+oSubSection.propGet("CVERSION")+"\nTable Type: "+sCtypeVal+"\nTable type desc: "+sTableTypeDesc+"\nUse: Table container"+"\nSection type: "+oSubSection.propGet("CTYPE"));
											oChildNode.setAttribute("use","tablecontainer");
											var aTableName = getTableinSection("",oDoc,oSubSection.index);
											oChildNode.setAttribute("tablename",aTableName[0]);
											//oChildNode.setAttribute("typeofsection",oSubSection.propGet("CTYPE"));
											//var olastChild = oChildNode.lastChild;
											//olastChild.style.borderBottom="1px dotted";
										}
										
										if(sCtypeVal=="PARAGRAPH")
										{
											var sStr = getParaData(oDoc,oSubSection);	
											oChildNode.innerHTML = sStr;
											oChildNode.setAttribute("title","Component: Section\nSection name: "+sSectionTitle+"\nSection label: "+"\nSection index: "+oSubSection.index()+"\nVersion: "+oSubSection.propGet("CVERSION")+"\nSection type: "+oSubSection.propGet("CTYPE"));
											oChildNode.setAttribute("component","para");
											oChildNode.setAttribute("objecttype","para");	
											oChildNode.setAttribute("typeofsection",oSubSection.propGet("CTYPE"));
										}		
										break;
									}
								}
								buildTreeViewEx(sSubSection,oDoc);//,oProgBar);
							}							
						}
					}
				}
			}
		}
	}catch(e)
	{
		alert(e.description);
	}finally{
		//Clear all objects
		oChildNode = null;
		oChildListItemNode = null;
		oHTMLParentElement = null;
		oContainer = null;
		oSection = null;
		oTextnode =  null;
		oParentListItemNode =  null;
		oParentNode =  null;
		oDoc = null;
	}	
}

function getSkipCondIndicator(oSection)
{
	var iSectionSkipCond = oSection.evaluateSkip();
	if(iSectionSkipCond===1)
		var sColor = '#0000FF';
	else
		var sColor = '#000000';	

	return sColor;
}				

function createReturnValue()
{
	try{
		//debugger;
		//debugger;
		var JSONReturnValue = {
			"deleteItem":{"row":[],"section":[],"paragraph":[]},
			"addItem":{"row":[],"section":{},"paragraph":[]},
			"modifyItem":{"row":[],"section":[],"paragraph":[]}
		}
	
		var aAllElements = document.getElementsByTagName("*");
		for(var i=0;i<aAllElements.length;i++)
		{
			var oHTMLElement = aAllElements[i];
			if(isInputValid(oHTMLElement) && (oHTMLElement.tagName=="LI"||oHTMLElement.tagName=="UL"))
			{
				//get the type of element
				var sComponent = oHTMLElement.getAttribute("component");
				var sDelete = oHTMLElement.getAttribute("delete");
				var sAdd = oHTMLElement.getAttribute("add");
				var sUpdate = oHTMLElement.getAttribute("update");
				var sRefElementId = oHTMLElement.getAttribute("refelementid");
				var sID = oHTMLElement.id;
				var sPosition = "";
				if(sComponent=="row" && isInputValid(sID))
				{	
					//Get the data needed for the appropriate action on a row
					var sTableName = oHTMLElement.getAttribute("tablename");
					var sRowType = oHTMLElement.getAttribute("rowtype");
					
					//Check if element needs to be deleted, if it does add it to the item and do not check if any other action needs to be performed on it
					//its an item thats being deleted nothing more can be done to it, if it has add equal to true do not add it to this list
					if(sDelete=="true" && sAdd!="true"){
						JSONReturnValue.deleteItem.row[JSONReturnValue.deleteItem.row.length] = {"id":sID,"tablename":sTableName};
						continue;
					}
					
					if(isInputValid(sRefElementId))
					{
						sPosition = "after";
					}
					else
					{
						sRefElementId = "";
						//if no refrence then add to bottom of the table
						sPosition = "bottom";
					}
					
					//Check if an item needs to be added
					if(sAdd=="true")
						JSONReturnValue.addItem.row[JSONReturnValue.addItem.row.length]={"position":sPosition,"refelementid":sRefElementId,"tablename":sTableName,"rowtype":sRowType,"mappingno":"","description":""};
					
					//Check if an item needs to be modified, if the item does not have an id or is an item with the flag add it means its a new item and there is no need to be modified
					if(sUpdate=="true" && sAdd!="true")
						JSONReturnValue.modifyItem.row[JSONReturnValue.updateItem.row.length] = {"id":sID,"tablename":sTableName};
				}
				
				if(sComponent=="section")// && isInputValid(sID))
				{
					//Check if element needs to be deleted, if it does add it to the item and do not check if any other action needs to be performed on it
					//its an item thats being deleted nothing more can be done to it, if it has add equal to true do not add it to this list
					if(sDelete=="true" && sAdd!="true"){
						JSONReturnValue.deleteItem.section[JSONReturnValue.deleteItem.section.length] = {"id":oHTMLElement.jumpcode};
						continue;
					}
					//Check if an item needs to be added
					if(sAdd=="true")
						JSONReturnValue.addItem.section[JSONReturnValue.addItem.section.length]={"sectiontype":"NOTECONTROL","position":"before","refelementid":sRefElementId};
					
					//Check if an item needs to be modified, if the item does not have an id or is an item with the flag add it means its a new item and there is no need to be modified
					if(sUpdate=="true" && sAdd!="true")
						JSONReturnValue.updateItem.section[JSONReturnValue.updateItem.section.length] = {"id":oHTMLElement.jumpcode};
				}

				if(sComponent=="para")// && isInputValid(sID))
				{
					//Check if element needs to be deleted, if it does add it to the item and do not check if any other action needs to be performed on it
					//its an item thats being deleted nothing more can be done to it, if it has add equal to true do not add it to this list
					if(sDelete=="true" && sAdd!="true"){
						JSONReturnValue.deleteItem.paragraph[JSONReturnValue.deleteItem.paragraph.length] = {"id":sID,"sectionname":"NOTES_002_01_03"};
						continue;
					}
					//Check if an item needs to be added
					if(sAdd=="true")
						JSONReturnValue.addItem.paragraph[JSONReturnValue.addItem.paragraph.length]={"id":sID,"sectionname":"NOTES_002_01_03","refelementid":sRefElementId}
					
					//Check if an item needs to be modified, if the item does not have an id or is an item with the flag add it means its a new item and there is no need to be modified
					if(sUpdate=="true" && sAdd!="true")
						JSONReturnValue.updateItem.paragraph[JSONReturnValue.updateItem.paragraph.length] = {"id":sID,"new content":""};
				}				
			}
		}
		//debugger;
		//debugger;
		JSONReturnValue = JSON.stringify(JSONReturnValue);
		return JSONReturnValue;
	}catch(e)
	{
		alert(e.description);
	}finally{
	}
}

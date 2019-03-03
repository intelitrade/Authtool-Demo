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
					//oParentListItemNode.innerHTML="<SPAN>"+sSectionTitle+"</SPAN>"//.appendChild(oTextnode);
					oParentListItemNode.innerHTML=sSectionTitle;
					oParentNode.setAttribute("id",sMainSectionId);
					oParentNode.setAttribute("component","section");
					oParentNode.setAttribute("objecttype","section");
					//oParentListItemNode.setAttribute("style","background-color:#F5DEB3");
					oParentListItemNode.style.backgroundColor = "#F5DEB3";
					oParentListItemNode.style.fontWeight = "bold"
					oParentListItemNode.style.padding="10px";
					
					oParentListItemNode.setAttribute("id","HEADER_"+sMainSectionId);
					oParentListItemNode.setAttribute("isheader","true");					
					
					oParentListItemNode.setAttribute("ismainsection","true");
					
					oParentListItemNode.setAttribute("component","section");
					oParentListItemNode.setAttribute("objecttype","section"); 
					oParentNode.setAttribute("use","sectioncontainer");
					
					oParentNode.appendChild(oParentListItemNode);				
					oContainer.appendChild(oParentNode);
					oParentNode.setAttribute("id",sMainSectionId);
					oParentNode.setAttribute("use","heading");
					oParentListItemNode.setAttribute("jumpcode",sSectionLabel);
					oParentNode.setAttribute("jumpcode",sSectionLabel);
					oParentListItemNode.setAttribute("onclick","gotoSection(this);highlightElement(this)","updatePreviewPane(this)");
					oParentListItemNode.onclick = function() { gotoSection();highlightElement();updatePreviewPane();};					
					oParentNode.setAttribute("title","Component: Section\nSection name: "+sSectionTitle+"\nSection label: "+sSectionLabel+"\nSection index: "+oSection.index()+"\nVersion: "+oSection.propGet("CVERSION")+"\nSection type: "+oSection.propGet("CTYPE"));
					
					oParentNode.setAttribute("typeofsection",oSection.propGet("CTYPE"));
					oParentListItemNode.setAttribute("typeofsection",oSection.propGet("CTYPE"));
					oParentListItemNode.setAttribute("use","sectioncontainer");	

					//oParentNode.setAttribute("name",sElementName);	
					oParentNode.setAttribute("name",sSectionLabel);	
					oParentNode.setAttribute("guid",oSection.propGet("GUID"));	

					oParentListItemNode.setAttribute("name",sSectionLabel);	
					oParentListItemNode.setAttribute("guid",oSection.propGet("GUID"));	

					//Add cursor style
					oParentNode.style.cursor = sTableLineItemCursor;
					oParentListItemNode.style.cursor = sTableLineItemCursor;
					
					if(document.getElementById("docmapheading"))
						document.getElementById("docmapheading").innerHTML = "Edit options<br>"+sSectionTitle;

					var iSectionSkipCond = oSection.evaluateSkip();
					var iSectionHideCond = oSection.evaluateHide();
					
					if(iSectionSkipCond===1 || iSectionHideCond==1)
						sColor = '#0000FF';
					else
						sColor = '#000000';	
						
					oParentListItemNode.style.color = sColor;
					
					var aOuterSections = getOuterSections(sSectionLabel,oDoc);
					if(isInputValid(aOuterSections))
					{
						for(var j=aOuterSections.length;j>=0;j--)
						{
							var sOuterSection = aOuterSections[j];
							if(!isInputValid(sOuterSection))
								continue;

							oParentNode.setAttribute("parentsection",sOuterSection);
							oParentListItemNode.setAttribute("parentsection",sOuterSection);

							oParentNode.setAttribute("parentguid",oDoc.sectionByName(sOuterSection).propGet("GUID"));
							oParentListItemNode.setAttribute("parentguid",oDoc.sectionByName(sOuterSection).propGet("GUID"));

							break;
						}
					}
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
										//oChildListItemNode.innerHTML="<SPAN>"+sSectionTitle+"</SPAN>"//.appendChild(oChildTextnode);
										oChildListItemNode.innerHTML=sSectionTitle
										oChildNode.appendChild(oChildListItemNode);
										oHTMLParentElement.appendChild(oChildNode);
										oChildNode.setAttribute("typeofsection",oSubSection.propGet("CTYPE"));
										oChildListItemNode.setAttribute("typeofsection",oSubSection.propGet("CTYPE"));
										
										oChildNode.setAttribute("parentsection",sOuterSection);
										oChildListItemNode.setAttribute("parentsection",sOuterSection);
										
										oChildNode.setAttribute("parentguid",oOuterSection.propGet("GUID"));
										oChildListItemNode.setAttribute("parentguid",oOuterSection.propGet("GUID"));

										oChildNode.setAttribute("name",sSubSection);	
										oChildNode.setAttribute("guid",oSubSection.propGet("GUID"));	

										oChildListItemNode.setAttribute("name",sSubSection);	
										oChildListItemNode.setAttribute("guid",oSubSection.propGet("GUID"));	

										var iSectionSkipCond = oSubSection.evaluateSkip();
										var iSectionHideCond = oSubSection.evaluateHide();
										
										if(iSectionSkipCond===1 || iSectionHideCond==1)
											sColor = '#0000FF';
										else
											sColor = '#000000';
										
										//Set mouse pointer
										oChildNode.style.cursor = sTableLineItemCursor;
										oChildListItemNode.style.cursor = sTableLineItemCursor;
										
										if(sCtypeVal=="NOTECONTROL"||sCtypeVal=="SUBNOTECONTROL"||sCtypeVal=="EXPANDCOLLAPSE"||sCtypeVal=="SUBNOTECONTROLGROUP"||sCtypeVal=="EXPANDCOLLAPSECOMPANY3RDYEAR")
										{
											oChildListItemNode.setAttribute("id","HEADER_"+iSubSectionId);
											oChildListItemNode.setAttribute("isheader","true");
											oChildListItemNode.setAttribute("jumpcode",sSubSection);
											oChildNode.setAttribute("jumpcode",sSubSection);
											
											oChildListItemNode.setAttribute("onclick","gotoSection(this);highlightElement(this)","updatePreviewPane(this)");
											oChildListItemNode.onclick = function() { gotoSection();highlightElement();updatePreviewPane();};
											
											oChildNode.style.color=sColor;
											oChildNode.setAttribute("title","Component: Section\nSection name: "+sSectionTitle+"\nSection label: "+sSubSection+"\nSection index: "+oSubSection.index()+"\nVersion: "+oSubSection.propGet("CVERSION")+"\nSection type: "+oSubSection.propGet("CTYPE"));
											oChildListItemNode.style.backgroundColor = "#F5DEB3";
											oChildListItemNode.style.fontWeight = "bold"
											oChildListItemNode.style.padding="10px";

											oChildListItemNode.setAttribute("component","section");
											oChildListItemNode.setAttribute("objecttype","section");
											oChildNode.setAttribute("use","sectioncontainer");
											oChildListItemNode.setAttribute("use","sectioncontainer");										
										}
										
										if(sCtypeVal=="SUMMARYTABLE"||sCtypeVal=="STD6COLUMNTABLE"||sCtypeVal=="RECONTABLE"||sCtypeVal=="SUMMARYTABLEGROUP"||sCtypeVal=="HIDDENTABLEWHICHPULLSTOBS"||sCtypeVal=="RADIOBUTTONTABLE"||sCtypeVal=="GUIDANCETABLE"||sCtypeVal=="DEPMETHODSANDUSEFULLLIFETYPE1")
										{
											//Get table in section
											var sStr = getTableData(oDoc,oSubSection);
											if(sStr=="")
												continue;
											
											var sTableTypeDesc = getTableTypeDesc(sCtypeVal);
											
											/*
											//Create a list item that will act as the header of the section or keep all data about the new section being added
											var oNewSectionListItemNode = document.createElement("LI");
											//Set the text the users will see on screen
											oNewSectionListItemNode.innerHTML=sElementName+"<sup style='color:red;'>New*</sup>";
											//Set background color to Wheat
											oNewSectionListItemNode.style.backgroundColor = "#F5DEB3";
											//Make the text bold
											oNewSectionListItemNode.style.fontWeight = "bold";
											//Set padding on all sides i.e. left, right, top and bottom to 10px i.e. space between the text and the borders surrounding it
											oNewSectionListItemNode.style.padding="10px";	
											//set the id for the LI item element
											oNewSectionListItemNode.setAttribute("id","HEADER_"+sSubSectionId);
											//Set the attribute isHeader to indicate that this li item is a header for the section and not one of the items that need to be inserted
											//this store the name of the section and other details
											oNewSectionListItemNode.setAttribute("isheader","true");											
											var sTableHeader = "<li style='font-weight:bold;background-color:#F5DEB3;'></li>"
											node.setAttribute("onclick","gotoSection(this);highlightElement(this);updatePreviewPane(this);");
											node.onclick = function() { gotoSection();highlightElement();updatePreviewPane() };
											*/
											oChildNode.innerHTML = "<li id='HEADER_"+iSubSectionId+"' style='font-weight:bold;background-color:#BDB76B;padding:10px' onclick='gotoSection(this);highlightElement(this);updatePreviewPane(this);' isheader='true' jumpcode='"+sSubSection+"'>Table ("+toTitleCase(sCtypeVal)+")</li>"+sStr;
											oChildNode.setAttribute("title","Component: Section\nSection name: "+sSectionTitle+"\nSection label: "+sSubSection+"\nSection GUID: "+oSubSection.propGet("GUID")+"\nVersion: "+oSubSection.propGet("CVERSION")+"\nTable Type: "+sCtypeVal+"\nTable type desc: "+sTableTypeDesc+"\nUse: Table container"+"\nSection type: "+oSubSection.propGet("CTYPE"));
											oChildNode.setAttribute("use","tablecontainer");
											var aTableName = getTableinSection("",oDoc,oSubSection.index);
											oChildNode.setAttribute("tablename",aTableName[0]);
										}
										
										//The following block is for input text wrapped around a section and not in a table row, that is why the component and object type is set to section
										if(sCtypeVal=="PARAGRAPH")
										{
											var sStr = getParaData(oDoc,oSubSection);	
											oChildNode.innerHTML = sStr;
											oChildNode.setAttribute("title","Component: Section\nSection name: "+sSectionTitle+"\nSection label: "+"\nSection index: "+oSubSection.index()+"\nVersion: "+oSubSection.propGet("CVERSION")+"\nSection type: "+oSubSection.propGet("CTYPE"));
											oChildNode.setAttribute("component","section");
											oChildNode.setAttribute("objecttype","section");	
											oChildNode.setAttribute("typeofsection",oSubSection.propGet("CTYPE"));
											oChildNode.setAttribute("sectiontype",oSubSection.propGet("CTYPE"));
											oChildListItemNode.setAttribute("jumpcode",sSubSection);
											oChildNode.setAttribute("jumpcode",sSubSection);																					
										}		
										break;
									}
								}
								buildTreeViewEx(sSubSection,oDoc);
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
		/*
			Element reference ids should be obtained during this process and not prior as this will cause issues if other elements are added in between or elements are deleted
		*/
		var JSONReturnValue = {
			"deleteItem":{"row":[],"section":[],"paragraph":[]},
			"addItem":{"row":[],"section":[],"paragraph":[]},
			"modifyItem":{"row":[],"section":[],"paragraph":[]}
		}
		
		var sRefElementId = "";
		
		//Speed improvement
		//var aAllElements = document.getElementsByTagName("*");
		var aAllElements = document.getElementById('MainContentContainer').getElementsByTagName('*');
		//return;
		//debugger;
		//debugger;
		//for(var i=document.getElementById("MainContentContainer").sourceIndex;i<aAllElements.length;i++)
		var iElements = aAllElements.length;
		for(var i=0;i<iElements;i++)
		{
			var oHTMLElement = aAllElements[i];
			if(isInputValid(oHTMLElement) && (oHTMLElement.tagName=="LI"||oHTMLElement.tagName=="UL"))
			{
				//get the type of element
				var sComponent = oHTMLElement.getAttribute("component");
				var sDelete = oHTMLElement.getAttribute("delete");
				var sAdd = oHTMLElement.getAttribute("add");
				var sUpdate = oHTMLElement.getAttribute("update");

				/*
					Element reference ids should be obtained during this process and not prior as this will cause issues if other elements are added in between or elements are deleted
				*/				
				
				//var sRefElementId = oHTMLElement.getAttribute("refelementid");
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
					else highlightElement(this.parentNode)
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
				
				/*if(sComponent=="section"||sComponent=="inputtextsection")// && isInputValid(sID))
				{

					var sSectionType = oHTMLElement.typeofsection;
					if(oHTMLElement.parentNode.typeofsection=="PARAGRAPH"||oHTMLElement.typeofsection=="PARAGRAPH");
					{
						//debugger;
						//debugger;
					}
					
					if(oHTMLElement.parentNode.typeofsection=="PARAGRAPH")
					{
						sSectionType="PARAGRAPH";
						oHTMLElement.parentNode.name
						var sSectionGUID =  oHTMLElement.parentNode.guid;
						var sSectonName = oHTMLElement.parentNode.name;
						var sParentSection =  oHTMLElement.parentNode.parentsection;
						var sParentGUID =  oHTMLElement.parentNode.parentguid;
						//var sDelete = oHTMLElement.parentNode.getAttribute("delete");						
					}else{
						var sSectionGUID = oHTMLElement.guid;
						var sSectonName = oHTMLElement.name;
						var sParentSection = oHTMLElement.parentsection;
						var sParentGUID = oHTMLElement.parentguid;
					}
					//Check if element needs to be deleted, if it does add it to the item and do not check if any other action needs to be performed on it
					//its an item thats being deleted nothing more can be done to it, if it has add equal to true do not add it to this list
					if(sDelete=="true"){// && sAdd!="true"){
						JSONReturnValue.deleteItem.section[JSONReturnValue.deleteItem.section.length] = {"sectiontype":sSectionType,"guid":sSectionGUID,"name":sSectonName,"refelementid":sRefId,"parentsection":sParentSection,"parentguid":sParentGUID};
						continue;
					}
					//Check if an item needs to be added
					if(sAdd=="true"){
						
						var sRefId = "";
						//Get the 
						var oPreviousSibling = oHTMLElement.parentNode.previousSibling;
						if(isInputValid(oPreviousSibling))
							sRefId = oPreviousSibling.jumpcode;
						
						JSONReturnValue.addItem.section[JSONReturnValue.addItem.section.length]={"sectiontype":sSectionType,"guid":sSectionGUID,"name":sSectonName,"refelementid":sRefId,"parentsection":sParentSection,"parentguid":sParentGUID};
					}
					//Check if an item needs to be modified, if the item does not have an id or is an item with the flag add it means its a new item and there is no need to be modified
					if(sUpdate=="true" && sAdd!="true")
						JSONReturnValue.updateItem.section[JSONReturnValue.updateItem.section.length] = {"sectiontype":sSectionType,"guid":sSectionGUID,"name":sSectonName,"refelementid":sRefId,"parentsection":sParentSection,"parentguid":sParentGUID};
				}

				if(sComponent=="para")// && isInputValid(sID))
				{
					//debugger;
					//debugger
					//Check if element needs to be deleted, if it does add it to the item and do not check if any other action needs to be performed on it
					//its an item thats being deleted nothing more can be done to it, if it has add equal to true do not add it to this list
					if(sDelete=="true"){// && sAdd!="true"){
						JSONReturnValue.deleteItem.paragraph[JSONReturnValue.deleteItem.paragraph.length] = {"guid":oHTMLElement.guid,"sectionname":oHTMLElement.jumpcode};
						continue;
					}
					//Check if an item needs to be added
					if(sAdd=="true")
						JSONReturnValue.addItem.paragraph[JSONReturnValue.addItem.paragraph.length]={"guid":oHTMLElement.guid,"sectionname":oHTMLElement.jumpcode,"refelementid":sRefElementId}
					
					//Check if an item needs to be modified, if the item does not have an id or is an item with the flag add it means its a new item and there is no need to be modified
					if(sUpdate=="true")
						JSONReturnValue.updateItem.paragraph[JSONReturnValue.updateItem.paragraph.length] = {"guid":oHTMLElement.guid,"sectionname":oHTMLElement.jumpcode,"new content":""};
				}*/		
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

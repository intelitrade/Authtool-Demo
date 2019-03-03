function getSectionTitle(sSection)
{
	try{
		var aResult = [];
		var oSection = oDoc.sectionByName(sSection);
		var sSectionName = "";
		if(oSection && oSection.propExists(CSECTIONCONTROL))
		{
			sSectionName = getSectionName(oDoc, sSection);
			if (isInputValid((sSectionName) && sSectionName != "Section name")) {
				//return sSectionName;
			}else{
				//get the family tree of the section and start going up the tree to get the name
				//this might not be the case with content that has just been added
				var aOuterSections = getSectionFamilyTreeLib(sSection,oDoc);
				if(isInputValid(aOuterSections))
				{
					for(var i=aOuterSections.length;i>=0;i--)
					{
						sSection = aOuterSections[i];
						sSectionName = getSectionTitle(sSection);
						if (isInputValid(sSectionName) && sSectionName != "Section name") {
							break;//return sSectionName;
						}
					}
				}
			}
		}
		var sHeaderCell = getHeaderCellInSection(sSection,sSectionName);
		//Get the cell with the section title
		return aResult[aResult.length] = [sSectionName,sHeaderCell];
	}catch(e)
	{
		alert(e.description);
	}finally{
		oSection = null;
		oCircleList = null;
		oCirlce = null;
		aOuterSections = null;
	}
}

function getHeaderCellInSection(sSection,sSectionName)
{
	try{
		var aTable = getTableinSection(sSection,oDoc);
		var sCellName = "";
		if(isInputValid(aTable))
		{
			for(i=0;i<aTable.length;i++)
			{
				var sTable = aTable[i];
				var sCellTempName = sTable+".HEADER";
				var oCell = oDoc.cell(sCellTempName);
				if(oCell && oCell.value===sSectionName)
				{
					sCellName = sCellTempName;
					break;
				}
			}
		}
		return sCellName;
	}catch(e)
	{
		alert(e.description);
	}finally{
		oCell = null;
	}
}

function getClosestSectionWithCtype(sSection)
{
	try{
		//get the family tree of the section and start going up the tree to get the name
		//this might not be the case with content that has just been added
		var sSectionName = getSectionName(oDoc, sSection);
		if (isInputValid(sSectionName) && sSectionName != "Section name") {
			//return sSectionName;
		}else{
			var aOuterSections = getSectionFamilyTreeLib(sSection,oDoc);
			if(isInputValid(aOuterSections))
			{
				for(var i=aOuterSections.length;i>=0;i--)
				{
					sSection = aOuterSections[i];
					sSectionName = getSectionTitle(sSection);
					if (isInputValid(sSectionName) && sSectionName != "Section name")
					{
						break;
					}
				}
			}
		}
		return sSectionName; 			
	}catch(e)
	{
		alert(e.description);
	}finally{
		aOuterSections = null;
	}
}

function getMainSectionHeaderName()
{
	try{
		var sSectionName = "";
		//Check if the section is a notecontrol table
		var oSection = oDoc.sectionByName(sClosestSection);
		if(oSection.propExists("CTYPE")==1 && oSection.propGet("CTYPE")!="NOTECONTROL")
		{
			var aOuterSections = getSectionFamilyTreeLib(sClosestSection,oDoc);
			for(var i=0;i<aOuterSections.length;i++)
			{
				var sSection = aOuterSections[i];
				var oNoteControlSection = oDoc.sectionByName(sSection);
				if(oNoteControlSection.propExists("CTYPE")==1 && oNoteControlSection.propGet("CTYPE")=="NOTECONTROL")
				{
					sSectionName = getSectionName(oDoc,sSection);
					break;
				}
			}
		}else{
			sSectionName = getSectionName(oDoc,sClosestSection);
		}
		return sSectionName;
	}catch(e)
	{
		alert(e.description);
	}finally{
		
	}
}

function getSubSectionName(sSection)
{
	try{
		//debugger;
		//debugger;
		var sSectionLabel="";
		if(isInputValid(sSection))
		{
			var oSection = oDoc.sectionByName(sSection);
			sSectionLabel = sSection;
		}
		else
		{
			var oSection = oDoc.sectionByName(sClosestSection);
			sSectionLabel = sClosestSection;
		}
		var sSectionName = "";
		if(oSection)
		{
			var sCtypeVal = oSection.propGet("CTYPE");
			if(sCtypeVal=="NOTECONTROL"||sCtypeVal=="SUBNOTECONTROL"||sCtypeVal=="EXPANDCOLLAPSE"||sCtypeVal=="SUBNOTECONTROLGROUP"||sCtypeVal=="EXPANDCOLLAPSECOMPANY3RDYEAR")
			{					
				sSectionName = getSectionName(oDoc,sSectionLabel);
			}
			else if(sCtypeVal=="SUMMARYTABLE"||sCtypeVal=="STD6COLUMNTABLE"||sCtypeVal=="RECONTABLE"||sCtypeVal=="SUMMARYTABLEGROUP"||sCtypeVal=="HIDDENTABLEWHICHPULLSTOBS"||sCtypeVal=="RADIOBUTTONTABLE"||sCtypeVal=="GUIDANCETABLE"||sCtypeVal=="DEPMETHODSANDUSEFULLLIFETYPE1")
			{
				//Get the sections sort group
				var sSortGroup = oSection.sortGroup;
				var sParentSection = getParentSection(oDoc,"",sSectionLabel);
				sSectionName = getSubSectionName(sParentSection);
			}				
		}
		return sSectionName;
	}catch(e)
	{
		alert(e.description);
	}finally{
		
	}
}

function updateSectionHeader(oElement)
{
	try{
		//debugger;
		//debugger;
		var sSection = oElement.getAttribute("sectionlabel");
		var aTable = getTableinSection(sSection,oDoc);
		var sCellName = "";
		if(isInputValid(aTable))
		{
			for(i=0;i<aTable.length;i++)
			{
				var sTable = aTable[i];
				var sCellTempName = sTable+".HEADER";
				var oCell = oDoc.cell(sCellTempName);
				if(oCell)// && oCell.value===sSectionName)
				{
					oCell.value = oElement.value;
					oDoc.recalculate(0);
					break;
				}
			}
		}			
	}catch(e)
	{
		alert(e.description);
	}finally{
		
	}
}

function updateTextSection(oElement)
{
	try{
		if(oElement.tagName=="SPAN" && oElement.getAttribute("paragraphno")!="")
		{
			var iParaNo = oElement.getAttribute("paragraphno");
			var sChildText = oElement.innerText;
			var oPara = oDoc.para(iParaNo);
			if(oPara){
				oPara.replaceText(0, (oPara.textLength()-1), sChildText);
			}
		}
		oDoc.recalculate(0);
	}catch(e)
	{
		alert(e.description);
	}finally{
		oElement = null;
	}
}

function getSectionName(oDoc,sSection)
{
	try{
		
		var CONTROL_TABLE = "PC";
		var SECTIONHEAD_TABLE = "H1";
		var SECTIONSUBHEAD_TABLE = "H2";
		var CTABLETYPE = "CTABLETYPE";
		var sSectionName = "";
		//Get the table in the section
		var aTable = getTableinSection(sSection,oDoc,"");	  
		for (var i=0;i<aTable.length;i++)
		{
			var sTable = aTable[i];
			var oTable = oDoc.tableByName(sTable);
			var sTableType = oTable.propGet(CTABLETYPE);
			//Get control table
			//if(sTableType==CONTROL_TABLE | sTableType== SECTIONSUBHEAD_TABLE){
			if(sTableType == CONTROL_TABLE || sTableType == SECTIONSUBHEAD_TABLE || sTableType == SECTIONSUBHEAD_TABLE|| sTableType == "6H2" || sTableType == "TH"){
				//var sTable = aTable[i];
				//Get the section name
				var oCell = oDoc.cell(sTable+".HEADER");
				sSectionName = oDoc.cell(sTable+".HEADER").value;
				break;
			}else{
				//If a control table is not found then return the default name section name
				sSectionName = "Section name";
			}
		} 
		if(sSectionName == "")
			sSectionName = "Section name";
		
		return sSectionName;
	}catch(e)
	{
		alert(e.description);
	}
}

//Gets all section srrounding the section passed in
function getOuterSectionsold(sSection,oDoc)
{
	try
	{	
		if(isInputValid(sSection))
		{
			var aOuterSections = new Array();
			//iParaIndex = oDoc.curParaIndex();
			if (oDoc.sectionByName(sSection))
			{
				var iParaIndex = oDoc.sectionByName(sSection).firstParaIndex;
				var iLastParaIndex = oDoc.sectionByName(sSection).lastParaIndex;
				
				for (var i = 1; i <= oDoc.sectionCount; i++)
				{
					var oSection = oDoc.section (i);

					if (oSection.firstParaIndex <= iParaIndex && oSection.lastParaIndex >= iLastParaIndex && oSection.label!="" && oSection.label!=sSection)
					{
						aOuterSections[aOuterSections.length] = oSection.Label;
					}
				}
			}
		}
		return aOuterSections;
	}catch(e)
	{
		alert(e);
	}
}

//Gets all section srrounding the section passed in
function getOuterSections(sSection,oDoc)
{
	try
	{
		if(isInputValid(sSection))
		{
			var aOuterSections = new Array();
			//iParaIndex = oDoc.curParaIndex();
			var oSection = oDoc.sectionByName(sSection);
			if (oSection)
			{
				var iFirstParaIndex = oDoc.sectionByName(sSection).firstParaIndex;
				var iLastParaIndex = oDoc.sectionByName(sSection).lastParaIndex;
				var iSectionIndex = oSection.index();
				for(var i = iSectionIndex;i>=1;i--)
				{
					var oOuterSection = oDoc.section(i);
					if (oOuterSection.index<=iSectionIndex) // Check if the section exists in the document and whether it is not the section we are looping in
					{
						iOuterFirstPara = oOuterSection.firstParaIndex // assign the fisrt paragraph index to iSubFirstPara
						iOuterLastPara = oOuterSection.lastParaIndex// assign the last paragraph index to iSubLastPara
						
						//check if these fall outside the first sections para range
						if (iOuterFirstPara <= iFirstParaIndex & iOuterLastPara >= iLastParaIndex && oOuterSection.label!="" && oOuterSection.label!=sSection)
						{
							aOuterSections.unshift(oOuterSection.label);
						}							
						
						//Check if main page has been reached. If main page break out of the loop this is the outter most section
						if(oOuterSection.propExists("CFORMAT")==1)
							break;												
						
					}
				}
			}
		}
		return aOuterSections;
	}catch(e)
	{
		alert(e);
	}
}

//Get all subsections belonging to a section
function getSubsections(sSectionLabel,oDoc,iSectionIndex)
{
	try{
		
		if (!isInputValid(sSectionLabel)&& iSectionIndex>0)
		{
			var oSection = oDoc.Section(iSectionIndex);//Get section object
		}else
		{
			var oSection = oDoc.SectionByName(sSectionLabel);//Get section object
		}
		
		if (oSection) // Check if section exists
		{
			var aSecParaIndex = getSectionParaIndexLib(oSection); //Get the first and last Paragraph of a section
			var iFirstPara  = aSecParaIndex[0]; // assign the fisrt paragraph index to iFirstPara
			var iLastPara = aSecParaIndex[1]; // assign the last paragraph index to iLastPara
			var iSections = oDoc.sectionCount(); //Get the number of sections in a document
			var aSection = new Array(); // Create an array object to hold all first level sections within a specific section
			var iSectionIndex = oSection.index();
			//for (var j=1;j<=iSections;j++) // loop through sections in the document
			for (var j=iSectionIndex;j<=iSections;j++)
			{
				oSubSection = oDoc.section(j) // assign sections in the document to oSubSection

				if (oSubSection.index>=iSectionIndex) //section must be equal or greater that the section passed in this function
				{
					iSubFirstPara = oSubSection.firstParaIndex // assign the fisrt paragraph index to iSubFirstPara
					iSubLastPara = oSubSection.lastParaIndex// assign the last paragraph index to iSubLastPara
					
					if(oSubSection.firstParaIndex>oSection.lastParaIndex)
						break;
					
					
					//check if these fall within the first sections para range
					if (iSubFirstPara >= iFirstPara & iSubLastPara <= iLastPara)
					{
						aSection[aSection.length]=oSubSection // if a section is found within the specified section add to array aSections
					}
				}
			}
			return aSection; // return an array with sections found within that section i.e. the Section Object is returned
		}
	}catch(e)
	{
		alert(e.description);
	}finally{
		
	}
}

//gets the first and last paragraph index of a section and returns it as an array
function getSectionParaIndexLib(oSection)
{
	try{
		var iFirstParaIndex = oSection.firstParaIndex; // assign the first paragraph index of a section to iFirstParaIndex
		var iLastParaIndex = oSection.lastParaIndex; // assign the last paragraph index of a section to iLastParaIndex
		var aParaIndex = new Array(iFirstParaIndex,iLastParaIndex);		
		return aParaIndex;
	}catch(e)
	{
		alert(e.description);
	}finally{
		
	}
}

//Jump to different place in CaseView document
function gotoSection(oElement)
{
	try{
		//debugger;
		//debugger;
		var sLabel = "";
		if(isInputValid(oDoc))
			oCaseViewDocument = oDoc;
			
		if(!oElement)	
			oElement = event.srcElement;

		if(oElement && isInputValid(oElement.getAttribute("jumpcode")))
			oDoc.gotoSection(oElement.getAttribute("jumpcode"));
		
		//oCaseViewDocument.gotoSection(oElement.getAttribute("jumpcode"));
		//oElement.scrollIntoView(true);
	}catch(e)
	{
		alert(e.description);
	}finally{
		oElement = null;
	}
}	

function InsertNewSection(oElement,sRefElement)
{
	try{
		//debugger;
		//debugger;
		if (oElement && isInputValid(sRefElement)) {
			//Check if the element need to be added or note
			if(oElement.style.color=="green")
			{
				var sElementName = oElement.name;
				var oRefElement = document.getElementById(sRefElement);
				var iSectionToInsert = 1;//prompt("Please enter number of "+sElementName+" row(s) to insert", 1);
				if(iSectionToInsert==null)
				{
					oElement.style.position = 'static';
					oElement.onmousemove = null;		
				}else{
					
					//get rid of the on mouse move event
					oElement.onmousemove = null;
					oElement.style.position = 'static';						

					//Create a progress bar to show number of rows being inserted and show when each row has been inserted
					var oProgBar = oDoc.createProgressBar((iSectionToInsert>1?"Inserting sections":"Inserting a section"),iSectionToInsert,1)
					//if(!oRefElement)	
					for(var x=1;x<=iSectionToInsert;x++)
					{
						var ID = createGuid();
                        var sSubSectionId = "UL_" + ID;//createGuid();//oSubSection.index;
						//create an Unordered list container i.e. UL that will hold all components with data about the new section that needs to be created
						var oNewSectionNode = document.createElement("UL");
						//Set the ID
						oNewSectionNode.setAttribute("id",sSubSectionId);
						//set a variabe that indicates what this element is used for
						oNewSectionNode.setAttribute("use","sectioncontainer");
						//oNewSectionNode.setAttribute("guid",ID);
						
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
						//Set the jumpcode - NEW Section not needed, same with got function not needed to be added to this element, the section does not exist in the CV document
						oNewSectionListItemNode.setAttribute("jumpcode","SECTION_"+sSubSectionId);
						oNewSectionNode.setAttribute("jumpcode","SECTION_"+sSubSectionId);
						//Attach events to the list item, event for jump to code 
						oNewSectionListItemNode.setAttribute("onclick","highlightElement(this);updatePreviewPane(this);");
						oNewSectionListItemNode.onclick = function() {highlightElement();updatePreviewPane(); };

						//Set the type section being inserted - currently 2 propertues are being used one of them will be dropped
						oNewSectionNode.setAttribute("typeofsection",oElement.sectiontype);
						oNewSectionListItemNode.setAttribute("typeofsection",oElement.sectiontype);

						//Add a guid to for the new row
						oNewSectionListItemNode.setAttribute("guid",ID);
						//Give the new row a name which user can identify with, name it after the element it is inserted from
						oNewSectionListItemNode.setAttribute("name",sElementName);
						
						//oChildNode.style.color=sColor;
						oNewSectionNode.setAttribute("title","Component: Section\nSection name: "+sElementName+"\nSection label: "+sSubSectionId+"\nSection index: \nVersion: \nSection type: "+oElement.sectiontype);

						oNewSectionListItemNode.setAttribute("component","section");
						oNewSectionListItemNode.setAttribute("objecttype","section");
						oNewSectionListItemNode.setAttribute("use","sectioncontainer");	
						oNewSectionListItemNode.setAttribute("add","true");	

						oNewSectionNode.setAttribute("parentsection",oRefElement.parentsection);
						oNewSectionNode.setAttribute("parentguid",oRefElement.parentguid);
						
						oNewSectionListItemNode.setAttribute("parentsection",oRefElement.parentsection);
						oNewSectionListItemNode.setAttribute("parentguid",oRefElement.parentguid);
						
						//Append to the parent element oRefElement.parentNode.parentNode.id
						oNewSectionNode.appendChild(oNewSectionListItemNode);
						if(oRefElement.typeofsection=="NOTECONTROL" && oSectionElement.sectiontype!="NOTECONTROL")
							var oNewElement = insertAsSub(oNewSectionNode, oRefElement.parentNode);
						else
							var oNewElement = insertAfter(oNewSectionNode, oRefElement.parentNode);
						
						if(oNewElement){
							//oNewElement.setAttribute("add","true");
							highlightElement(oNewElement.children[0]);
							oNewElement.children[0].scrollIntoView();
						}
						oProgBar.updateProgress(1);	
						oProgBar.setMessage("Section "+x+" / "+iSectionToInsert);						
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
		
	}
}

function onMouseDownSections(sId)
{
	try{
		event.cancelBubble = true;
		oSectionElement = document.getElementById(sId);
		//var oPosition = getElementsXYPosition(oSectionElement)
		var oParentNode = oSectionElement.parentNode;
		iPos = getTheElementsPosition(oParentNode,oSectionElement);
		//debugger;
		//debugger;
		//Get original settigs of the element before moving it around
		var iOriginalZIndex = oSectionElement.style.zIndex;
		var iElementLeft = oSectionElement.offsetLeft;//oSectionElement.style.left;
		var iElementTop = oSectionElement.offsetTop;//oSectionElement.style.top;
		
		// (4) onmouseup, remove unneeded handlers
		// (1) start the process
		// (2) prepare to moving: make absolute and on top by z-index
		oSectionElement.style.position = 'absolute';
		oSectionElement.style.zIndex = 1000;
		// ...and put that absolutely positioned element under thesElementUse=="tablecontainer" cursor
		oSectionElement.onmousemove = function ()
		{
			//moveAt(event.pageX, event.pageY);
			moveAtSections(event.clientX, event.clientY);
		}			
		
		// centers the oSectionElement at (pageX, pageY) coordinates
		function moveAtSections(pageX, pageY)
		{
			if(oSectionElement)
			{	
				oSectionElement.style.left = pageX - oSectionElement.offsetWidth + 'px';
				oSectionElement.style.top = pageY - oSectionElement.offsetHeight + 'px';
				oRefElement = document.elementFromPoint(event.clientX, event.clientY);
				if(oRefElement)
				{
					var sElementId = oRefElement.id;
					var sElementName = oSectionElement.name;
					var sElementUse = oRefElement.use;
					var sRefElementSectionType = oRefElement.typeofsection;
					
					//document.getElementById("insertoptions").innerHTML = sElementId+" "+oRefElement.typeofsection+" "+oSectionElement.sectiontype;
					
					//document.getElementById("results").innerHTML = oSectionElement.style.left;	
					//if(isInputValid(sElementId) && sElementId.search("HEADER_")!=-1 && (oRefElement.typeofsection==oSectionElement.sectiontype||oRefElement.typeofsection=="NOTECONTROL"&&oSectionElement.sectiontype!="NOTECONTROL")||((oRefElement.component=="section"||oRefElement.component=="row")&&oSectionElement.sectiontype=="SUBNOTECONTROL"))// && sElementId.search("row_")!=-1||sElementUse=="tablecontainer")
					if(isInputValid(sElementId) &&  oRefElement.ismenu!="true")
					{
						//debugger;
						//debugger;
						//oSectionElement.setAttribute("refelement", sElementId);
						sRefElement = sElementId;
						oSectionElement.style.color="green";
						var oRefParent = oRefElement.parentNode;
						var iParentNode = oRefParent.children.length;
						
						for(var i=0;i<iParentNode;i++)
						{
							//oParentNode.children[i].style.border="";
							try{
								var sChildId = oRefParent.children[i].id;
								if(sChildId.search("row_")!=-1)
								{
									oRefParent.children[i].style.textDecoration = "none";	
								}								
							}catch(e)
							{
							}
						}
						//debugger;
						//debugger;
						//oRefElement.style.textDecoration = "underline";
					}else{
						oSectionElement.style.color="red";
					}
				}
			}
		}	
		
		function onMouseMoveSections(event) {
			moveAtSections(event.clientX, event.clientY);
		}

		if(document.addEventListener)
		{ 
			document.addEventListener('mousemove', onMouseMoveSections);
		}else
		{
			document.attachEvent("onmousemove", onMouseMoveSections);		
		}	
	}catch(e)
	{
		alert(e.description);
	}finally{
		oSectionElement.style.left = "auto";
		oSectionElement.style.top = "auto";
		oSectionElement.style.color="";		
	}
}	
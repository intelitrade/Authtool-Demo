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

function testIt()
{
	debugger;
	debugger;
	var sSection = "NOTES_002_01";
	var aMySecV2 = getOuterSectionsV2(sSection,oDoc);
	var aSecTy = getOuterSections(sSection,oDoc);
	
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
					var oOuterSection = oDoc.section (i);
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

				if (oSubSection.index>=iSectionIndex) // Check if the section exists in the document and whether it is not the section we are looping in
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
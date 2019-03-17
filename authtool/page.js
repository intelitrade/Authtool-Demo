function loadAllPages()
{
	try{
		//debugger;
		//debugger;
		var sPageList = "";
		var iPages = oDoc.sectionCount;
		var iPageCount = 0;
		for (var i = 1; i <= oDoc.sectionCount; i++)
		{
			var oSection = oDoc.section(i);
			
			if(oSection.propExists("CFORMAT")==1)
			{
				var sPageLabel = oSection.label;
				var sPageName = getSectionName(oDoc,sPageLabel)
				iPageCount++;
				var sBgColor = "#FFFFFF";
				var sColor = "#000000";
				if(oSection.evaluateHide()==1 || oSection.evaluateSkip()==1)
					sColor = "blue";
				
				var sSectionState = "Not Skipped or Hidden";
				var sVersion = oSection.propGet("CVERSION");
				var sGUID = oSection.propGet("GUID");
				
				//Check if section skip or hidden or both
				if(oSection.evaluateHide()==1 && oSection.evaluateSkip()==1)
					sSectionState = "Skipped & Hidden";
				else if(oSection.evaluateSkip()==1)
					sSectionState = "Skipped";
				else if(oSection.evaluateHide()==1)
					sSectionState = "Hidden";
				
				//if(iPageCount%2==0)
					//sColor = "#FFF8DC";
				var sTitleText = "title='Page label: "+sPageLabel+"\nState: "+sSectionState+"\nVersion: "+sVersion+"\nGUID: "+sGUID+"'"
				if(sPageList=="")
					sPageList = "<li id='"+sPageLabel+"' component='page' objecttype='page' jumpcode='"+sPageLabel+"' style='margin:0px;padding:5px;font-weight:bold;color:"+sColor+"'"+sTitleText+" onclick='highlightElement(this);updatePreviewPane(this);'>"+sPageName+"</li>";
				else
					sPageList = sPageList + "<li id='"+sPageLabel+"' component='page' objecttype='page' jumpcode='"+sPageLabel+"' style='background-color:"+sBgColor+";margin:0px;padding:5px;font-weight:bold;color:"+sColor+"'"+sTitleText+" onclick='highlightElement(this);updatePreviewPane(this);'>"+sPageName+"</li>";
				
				//i = oSection.index;
			}
		}
		
		if(sPageList!=""){
			document.getElementById("docmapheading").innerHTML="<b>Edit Board</b><br>All Pages";
			document.getElementById("MainContentContainer").innerHTML="<ul id='pages'>"+sPageList+"</ul>";
		}
		
	}catch(e)
	{
		alert(e.description);
	}finally{
		
	}
}

//Get all pages
function getPagesInDocument(oDocument)
{
	try{
		if(oDocument)
			var oDoc = oDocument;
		
		var iSectionCount = oDoc.sectionCount;
		var aPageList = [];
		for (var i = 1; i <= iSectionCount; i++)
		{
			var oSection = oDoc.section(i);
			if(oSection.propExists("CFORMAT")==1)
			{
				aPageList[aPageList.length] = oSection;
			}
		}
		return aPageList;
	}catch(e)
	{
		alert(e.description)
	}finally{
		
	}
}

function getSubSectionBasedOnLevel(sSection)
{
	try{
		//debugger;
		//debugger;
			
		if(!isInputValid(sSection))
			var sSectionLabel = oDoc.interpret("cvdocpos('CFORMAT')");
		else	
			var sSectionLabel = sSection;//oDoc.interpret("cvdocpos('CFORMAT')");
		
		var oSection = oDoc.SectionByName(sSectionLabel);//Get section object

		if (oSection) // Check if section exists
		{
			//var aSecParaIndex = getSectionParaIndexLib(oSection); //Get the first and last Paragraph of a section
			var iFirstPara  = oSection.firstParaIndex;//aSecParaIndex[0]; // assign the fisrt paragraph index to iFirstPara
			var iLastPara = oSection.lastParaIndex;//aSecParaIndex[1]; // assign the last paragraph index to iLastPara
			var iSections = oDoc.sectionCount(); //Get the number of sections in a document
			var aSection = new Array(); // Create an array object to hold all first level sections within a specific section
			var iSectionIndex = oSection.index();
			//for (var j=1;j<=iSections;j++) // loop through sections in the document
			for (var j=iSectionIndex;j<=iSections;j++)
			{
				var oSubSection = oDoc.section(j) // assign sections in the document to oSubSection

				if (oSubSection.index>=iSectionIndex && oSubSection.label!="") //section must be equal or greater that the section passed in this function
				{
					iSubFirstPara = oSubSection.firstParaIndex // assign the fisrt paragraph index to iSubFirstPara
					iSubLastPara = oSubSection.lastParaIndex// assign the last paragraph index to iSubLastPara
					
					if(oSubSection.firstParaIndex>oSection.lastParaIndex)
						break;
					
					
					//check if these fall within the first sections para range
					if (iSubFirstPara >= iFirstPara & iSubLastPara <= iLastPara)
					{
						//get the outer section and check if it is the main section passed
						var aOuterSections = getOuterSections(sSubSectionLabel,oDoc);
						if(isInputValid(aOuterSections) && aOuterSections[aOuterSections.length-1]==sSectionLabel && isInputValid(aOuterSections[aOuterSections-1]))
						{
							var sSectionName = getSectionName(oDoc,oSubSection.label);							
							aSection[aSection.length]=oSubSection; // if a section is found within the specified section add to array aSections
						}
					}
				}
			}
			//debugger;
			//debugger;
			return aSection; // return an array with sections found within that section i.e. the Section Object is returned
		}	
	}catch(e)
	{
		alert(e.description);
	}finally{
		aSection = null;
		oSubSection = null;
		oSection = null;
	}
}

function loadLevel1SubSections(sSection)
{
	try{
		//debugger;
		//debugger;
		//get the current page
		if(!isInputValid(sSection))
			var sMainSection = oDoc.interpret("cvdocpos('CFORMAT')");
		else	
		 var sMainSection = sSection;//oDoc.interpret("cvdocpos('CFORMAT')");
		
		var aSubSection = getSubsections(sMainSection,oDoc);
		var aSiblingSections = [];
		var oMainSection = oDoc.sectionByName(sMainSection);
		var iMainSectionIndex = oMainSection.index;
		var iArrayEnd = aSubSection[aSubSection.length-1].index;
		var sPageList = "";
		var oProgBar = oDoc.createProgressBar("Loading Section List...",(iArrayEnd-iMainSectionIndex),1);
		for(var i=iMainSectionIndex;i<iArrayEnd;i++)
		{
			var oSubSection = aSubSection[i];
			if(!isInputValid(oSubSection))
				continue;
			
			var sSubSectionLabel = oSubSection.label;
			
			//if(sSubSectionLabel==sMainSection)
				//continue;
			
			//get the outer section and check if it is the main section passed
			var aOuterSections = getOuterSections(sSubSectionLabel,oDoc);
			if(isInputValid(aOuterSections) && aOuterSections[aOuterSections.length-1]==sMainSection)
			{
				var sSectionName = getSectionName(oDoc,sSubSectionLabel)
				
				if(sPageList=="")
					sPageList = "<li id='"+sSubSectionLabel+"' component='section' objecttype='section' jumpcode='"+sSubSectionLabel+"' style='margin:0px;padding:5px;font-weight:bold;onclick='highlightElement(this);updatePreviewPane(this);'>"+sSectionName+"</li>";
				else
					sPageList = sPageList + "<li id='"+sSubSectionLabel+"' component='section' objecttype='section' jumpcode='"+sSubSectionLabel+"' style='margin:0px;padding:5px;font-weight:bold;onclick='highlightElement(this);updatePreviewPane(this);'>"+sSectionName+"</li>";
			}
			oProgBar.updateProgress(1);	
			//aSiblingSections[aSiblingSections.length] = sSubSectionLabel;
			//for(var j=0;j<aOuterSections.lenght)
		}
	
		if(sPageList!=""){
			document.getElementById("docmapheading").innerHTML="<b>Edit Board</b><br>All Subsections";
			document.getElementById("MainContentContainer").innerHTML="<ul id='subsections'>"+sPageList+"</ul>";
		}	
		oProgBar.destroyProgressBar();
	
	}catch(e)
	{
		alert(e.description);
	}finally{
		
	}
}

function getImmediateSubSectionList(sSection)
{
	try{
		//debugger;
		//debugger;
			
		if(!isInputValid(sSection))
			var sSectionLabel = oDoc.interpret("cvdocpos('CFORMAT')");
		else	
			var sSectionLabel = sSection;//oDoc.interpret("cvdocpos('CFORMAT')");
		
		var oSection = oDoc.SectionByName(sSectionLabel);//Get section object

		if (oSection) // Check if section exists
		{
			//var aSecParaIndex = getSectionParaIndexLib(oSection); //Get the first and last Paragraph of a section
			var iFirstPara  = oSection.firstParaIndex;//aSecParaIndex[0]; // assign the fisrt paragraph index to iFirstPara
			var iLastPara = oSection.lastParaIndex;//aSecParaIndex[1]; // assign the last paragraph index to iLastPara
			var iSections = oDoc.sectionCount(); //Get the number of sections in a document
			var aSection = new Array(); // Create an array object to hold all first level sections within a specific section
			var iSectionIndex = oSection.index();
			//for (var j=1;j<=iSections;j++) // loop through sections in the document
			for (var j=iSectionIndex;j<=iSections;j++)
			{
				var oSubSection = oDoc.section(j) // assign sections in the document to oSubSection
				var sSubSectionLabel = oSubSection.label;
				if (oSubSection.index>=iSectionIndex && sSubSectionLabel!="") //section must be equal or greater that the section passed in this function
				{
					iSubFirstPara = oSubSection.firstParaIndex // assign the fisrt paragraph index to iSubFirstPara
					iSubLastPara = oSubSection.lastParaIndex// assign the last paragraph index to iSubLastPara
					
					if(oSubSection.firstParaIndex>oSection.lastParaIndex)
						break;
					
					//check if these fall within the first sections para range
					if (iSubFirstPara >= iFirstPara & iSubLastPara <= iLastPara)
					{
						//get the outer section and check if it is the main section passed
						var aOuterSections = getImmediateOuterSection(sSubSectionLabel,oDoc);
						if(isInputValid(aOuterSections) && aOuterSections[aOuterSections.length-1]==sSectionLabel && isInputValid(aOuterSections[aOuterSections.length-1]))
						{
							//var sSectionName = getSectionName(oDoc,sSubSectionLabel);							
							aSection[aSection.length]=oSubSection;//{'section':oSubSection,'sectionlabel':sSubSectionLabel,'sectionname':getSectionName(oDoc,sSubSectionLabel)};//[oSubSection,sSubSectionLabel,getSectionName(oDoc,sSubSectionLabel)]; // if a section is found within the specified section add to array aSections
						}
					}
				}
			}
			//debugger;
			//debugger;
			return aSection; // return an array with sections found within that section i.e. the Section Object is returned
		}	
	}catch(e)
	{
		alert(e.description);
	}finally{
		aSection = null;
		oSubSection = null;
		oSection = null;
		aOuterSections = null;
	}
}

function loadSubSectionsHTML(sSection)
{
	try{
		//debugger;
		//debugger;
			
		if(!isInputValid(sSection))
			var sSectionLabel = oDoc.interpret("cvdocpos('CFORMAT')");
		else	
			var sSectionLabel = sSection;//oDoc.interpret("cvdocpos('CFORMAT')");
		
		var oSection = oDoc.SectionByName(sSectionLabel);//Get section object

		if (oSection) // Check if section exists
		{
			//var aSecParaIndex = getSectionParaIndexLib(oSection); //Get the first and last Paragraph of a section
			var iFirstPara  = oSection.firstParaIndex;//aSecParaIndex[0]; // assign the fisrt paragraph index to iFirstPara
			var iLastPara = oSection.lastParaIndex;//aSecParaIndex[1]; // assign the last paragraph index to iLastPara
			var iSections = oDoc.sectionCount(); //Get the number of sections in a document
			var aSection = new Array(); // Create an array object to hold all first level sections within a specific section
			var iSectionIndex = oSection.index();
			//for (var j=1;j<=iSections;j++) // loop through sections in the document
			for (var j=iSectionIndex;j<=iSections;j++)
			{
				var oSubSection = oDoc.section(j) // assign sections in the document to oSubSection
				var sSubSectionLabel = oSubSection.label;
				if (oSubSection.index>=iSectionIndex && sSubSectionLabel!="") //section must be equal or greater that the section passed in this function
				{
					iSubFirstPara = oSubSection.firstParaIndex // assign the fisrt paragraph index to iSubFirstPara
					iSubLastPara = oSubSection.lastParaIndex// assign the last paragraph index to iSubLastPara
					
					if(oSubSection.firstParaIndex>oSection.lastParaIndex)
						break;
					
					//check if these fall within the first sections para range
					if (iSubFirstPara >= iFirstPara & iSubLastPara <= iLastPara)
					{
						//get the outer section and check if it is the main section passed
						var aOuterSections = getImmediateOuterSection(sSubSectionLabel,oDoc);
						if(isInputValid(aOuterSections) && aOuterSections[aOuterSections.length-1]==sSectionLabel && isInputValid(aOuterSections[aOuterSections.length-1]))
						{
							//var sSectionName = getSectionName(oDoc,sSubSectionLabel);							
							aSection[aSection.length]=oSubSection;//{'section':oSubSection,'sectionlabel':sSubSectionLabel,'sectionname':getSectionName(oDoc,sSubSectionLabel)};//[oSubSection,sSubSectionLabel,getSectionName(oDoc,sSubSectionLabel)]; // if a section is found within the specified section add to array aSections
						}
					}
				}
			}
			//debugger;
			//debugger;
			return aSection; // return an array with sections found within that section i.e. the Section Object is returned
		}	
	}catch(e)
	{
		alert(e.description);
	}finally{
		aSection = null;
		oSubSection = null;
		oSection = null;
		aOuterSections = null;
	}
}

//Gets all section srrounding the section passed in
function getImmediateOuterSection(sSection,oDoc)
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
							break;
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

function loadAllTables()
{
	try{
		
	}catch(e)
	{
		alert(e.description);
	}finally{
		
	}
}

function loadAllPara()
{
	try{
		
	}catch(e)
	{
		alert(e.description);
	}finally{
		
	}
}

function turnOffCurrentPage()
{
	try{
		
	}catch(e)
	{
		alert(e.description);
	}finally{
		
	}
}

function turnOffAllPages()
{
	try{
		
	}catch(e)
	{
		alert(e.description);
	}finally{
		
	}
}

function printCurrentPages()
{
	try{
		
	}catch(e)
	{
		alert(e.description);
	}finally{
		
	}
}

function printAllPages()
{
	try{
		
	}catch(e)
	{
		alert(e.description);
	}finally{
		
	}
}
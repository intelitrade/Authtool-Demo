function onLoad() 
{
	var sOriginalContent = document.documentElement.innerHTML;
	//load controls
	//loadDocMapDataEx();
	//window.focus();
}

function loadDocMapDataEx()
{
	return;
	/*
	if no page has been passed the sort order for the whole document will be exporter else it will only do for the page
	specified
	*/
	//debugger;
	//debugger;
	
	try{
		//if(!oDoc)
		//	var oDoc = document;
	
		//get the closest outer csectioncontrol custom property
		//get all inner sections from that section
		//loop through all sections
		//read the section custom property looking for "CTYPE"
		//take the value from ctype and add it to the document map and add the friendly name(Property, plant and equipment)
		//then we need to look at the value of CTYPE
			//if it is "NOTECONTROL" or "SUBNOTECONTROL" or "EXPANDCOLLAPSE" or "SUBNOTECONTROLGROUP" then dont worry about the table rows
			//if it is "SUMMARYTABLE" then loop through the table rows
			//if it reads a CTYPE that is not in this list then it must throw an error message showin the secction label and ctype that is missing in the code
		
		if(sClosestSection==="")
		{
			oDoc.messagebox("Error","CSECTIONCONTROL not received",0)
			return;
		}
		
		var sMainSectionName = getMainSectionHeaderName();
		if(isInputValid(sMainSectionName))
		{
			document.getElementById("docmapheading").innerHTML = "<b>Edit Board</b><br>("+sMainSectionName+")";
		}				

		var sStr = "";
	
		var aSection = [];
		var oSection = oDoc.sectionByName(sClosestSection);

		var aReturnSubSection = getSectionInSectionLib(sClosestSection,oDoc);
		//Remoe all sections without a label
		if(isInputValid(aReturnSubSection))
		{
			for(var i=0;i<aReturnSubSection.length;i++)
			{
				var oReturnedSection = aReturnSubSection[i];
				var sReturnedSection = oReturnedSection.label;
				
				if(isInputValid(sReturnedSection)||oReturnedSection.propExists("CTYPE")==1)
					aSection[aSection.length] = oReturnedSection; 
			}
		}

		//Check if a section has any section has been added to array aSection
		if(aSection.length==0)
		{
			//get the family tree of the section and start going up the tree to get the name
			//this might not be the case with content that has just been added
			var aOuterSections = getSectionFamilyTreeLib(sClosestSection,oDoc);
			if(isInputValid(aOuterSections))
			{
				for(var i=aOuterSections.length;i>=0;i--)
				{
					var sOuterSection = aOuterSections[i];
					var oOuterSection = oDoc.sectionByName(sOuterSection);
					if(oOuterSection)
					{
						var sOuterSectionCtype = oOuterSection.propGet("CTYPE");
						if(sOuterSectionCtype=="NOTECONTROL"||sOuterSectionCtype=="SUBNOTECONTROL"||sOuterSectionCtype=="EXPANDCOLLAPSE"||sOuterSectionCtype=="SUBNOTECONTROLGROUP"||sOuterSectionCtype=="EXPANDCOLLAPSECOMPANY3RDYEAR")
						{
							aSection[aSection.length]=oOuterSection;
							//Add the other section within oOuterSection
							var aReturnSubSection = getSectionInSectionLib(oOuterSection.label,oDoc);
							//Remoe all sections without a label
							if(isInputValid(aReturnSubSection))
							{
								for(var i=0;i<aReturnSubSection.length;i++)
								{
									var oReturnedSection = aReturnSubSection[i];
									var sReturnedSection = oReturnedSection.label;
									if(isInputValid(sReturnedSection))
										aSection[aSection.length] = oReturnedSection; 
								}
							}
							break;
						}
					}
				}
			}				
		}
		
		//Make sure that a least the first section in the list is a control table
		//If not get the first elements control table and make it the first one
		var oFirstSubSection = aSection[0];
		var sCType = oFirstSubSection.propGet("CTYPE")
		if(sCType!="NOTECONTROL"&&sCType!="SUBNOTECONTROL"&&sCType!="EXPANDCOLLAPSE"&&sCType!="SUBNOTECONTROLGROUP"&&sCType!="EXPANDCOLLAPSECOMPANY3RDYEAR")
		{
			//get the family tree of the section and start going up the tree to get the name
			//this might not be the case with content that has just been added
			var aOuterSections = getSectionFamilyTreeLib(oFirstSubSection.label,oDoc);
			if(isInputValid(aOuterSections))
			{
				for(var i=aOuterSections.length;i>=0;i--)
				{
					var sOuterSection = aOuterSections[i];
					var oOuterSection = oDoc.sectionByName(sOuterSection);
					if(oOuterSection)
					{
						var sOuterSectionCtype = oOuterSection.propGet("CTYPE");
						if(sOuterSectionCtype=="NOTECONTROL"||sOuterSectionCtype=="SUBNOTECONTROL"||sOuterSectionCtype=="EXPANDCOLLAPSE"||sOuterSectionCtype=="SUBNOTECONTROLGROUP"||sOuterSectionCtype=="EXPANDCOLLAPSECOMPANY3RDYEAR")
						{
							aSection.unshift(oOuterSection);
							break;
						}
					}
				}
			}				
		}
		
		var sTitle = "";
		sCtypeVal = oSection.propGet("CTYPE");
		//debugger;
		//debugger;
		var iHeaderCount = 0;
		var iSectionCount = 0;
		var sParentLabel = "";
		//debugger;
		//debugger;
		for (var i =0; i< aSection.length ;i++ )
		{
			var oInnerSection = aSection[i]
			
			var sLabel = oInnerSection.label;
			if(oInnerSection.propExists("CTYPE")==1)
			{

				var sCtypeVal = oInnerSection.propGet("CTYPE");
				var aHeaderProperties  = getSectionTitle(sLabel);
				if(isInputValid(aHeaderProperties))
					sTitle = aHeaderProperties[0];
				
				var iSectionSkipCond = oInnerSection.evaluateSkip();
				if(iSectionSkipCond===1)
					var sColor = '#0000FF';
				else
					var sColor = '#000000';							
				
				var sGuid = oInnerSection.propGet("GUID");
				if(sCtypeVal=="NOTECONTROL"||sCtypeVal=="SUBNOTECONTROL"||sCtypeVal=="EXPANDCOLLAPSE"||sCtypeVal=="SUBNOTECONTROLGROUP"||sCtypeVal=="EXPANDCOLLAPSECOMPANY3RDYEAR")
				{	
					sStr = sStr+"<ul style='margin-bottom:5px;margin-top:5px;background-color:#F5DEB3' id='"+sLabel+"'><li jumpcode='"+sLabel+"' id='"+aHeaderProperties[1]+"' onclick='gotoSection(this);highlightElement(this);highlightRow(this.id);updatePreviewPane(this);updateCurrentSelection(this)' style='margin-left:0px;padding:10px;color:"+sColor+"' objecttype='section' component='sectionheader' ctype='"+sCtypeVal+"' guid='"+sGuid+"'><b>"+sTitle+"</b></li></ul>";
					sParentLabel = sLabel;
				}
				else if(sCtypeVal=="SUMMARYTABLE"||sCtypeVal=="STD6COLUMNTABLE"||sCtypeVal=="RECONTABLE"||sCtypeVal=="SUMMARYTABLEGROUP"||sCtypeVal=="HIDDENTABLEWHICHPULLSTOBS"||sCtypeVal=="RADIOBUTTONTABLE"||sCtypeVal=="GUIDANCETABLE"||sCtypeVal=="DEPMETHODSANDUSEFULLLIFETYPE1")
				{
					//Get table in section
					var aTable = getTableinSection(sLabel,oDoc);
					if(isInputValid(aTable))
					{
						var sTable = aTable[0];
						var oTable = oDoc.tableByName(sTable);
						if(oTable)
						{
							var sTableType = oTable.propGet(CTABLETYPE);
							var sTableStr =  getTableLineItems(oTable,sTableType,sLabel,iSectionSkipCond);
							if(sTableStr=="")
							{
								continue;
							}
							
							sStr = sStr + "<ul style='margin-bottom:5px;margin-top:5px;' parentlabel='"+sParentLabel+"' id='"+sLabel+"'><li jumpcode='"+sLabel+"' onclick='gotoSection(this);updatePreviewPane(this);highlightElement(this)' component='column' tablename='"+sTable+"' style='color:"+sColor+";padding-left:25px;margin-top:5px'> Column header row</li><li style='color:"+sColor+";padding-left:25px;margin-top:5px' jumpcode='"+sLabel+"' onclick='gotoSection(this);updatePreviewPane(this);highlightElement(this)' component='mapcolumn' tablename='"+sTable+"'>Column mapping row</li>"+sTableStr+"</ul>";	
						}
					}
				}
				else if(sCtypeVal=="PARAGRAPH")
				{
					var iFirstParaIndex = oInnerSection.firstParaIndex;
					var oPara = oDoc.para(iFirstParaIndex);
					var sParaText = oPara.getText();
					var iSkipCond = oPara.evaluateSkip();
					var sParaGuid = oPara.propGet("GUID");
					if(iSkipCond===1 || iSectionSkipCond===1)
						var sColor = '#0000FF';
					else
						var sColor = '#000000';
					
					if(!isInputValid(sParaText))
						sParaText = "Paragraph text";
					else{
						var iTextLength = sParaText.length;
						if(iTextLength>50)
							sParaText = sParaText.substr(0,50)+"...";
					}
					sStr = sStr+"<ul style='margin-bottom:5px;margin-top:5px;'><li component='inputtextsection' id='"+sLabel+"' objecttype='section' jumpcode='"+sLabel+"' onclick='gotoSection(this);highlightElement(this);highlightSection(this.jumpcode);updatePreviewPane(this);updateCurrentSelection(this);' guid='"+sParaGuid+"' style='background-color:#CCFFCC;color:"+sColor+";margin-left:25px;margin-top:5px'>"+sParaText+"</li></ul>";
				}
				else
				{
					alert("Nothing" +sCtypeVal);
				}
			}
		}

		document.getElementById("MainContentContainer").innerHTML = sStr;			
	}catch(e)
	{
		alert(e.description);
	}finally{
		aHeaderProperties = null;
		oPara = null;
		oTable = null;
		oInnerSection = null;
		aSection = null;
		oSectionLevel = null;
	}		
}	


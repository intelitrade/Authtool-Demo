//Returns a row number based a a cell number
//Need to be made smarter by using regular expressions
function getRowNumberFromCell(sCellNumber)
{	//If no GUID found make the script backward compatible
	var aCellNumber = sCellNumber.split(".");
	var sCellSuffix = aCellNumber[1];
	var iCellSuffixLength = sCellSuffix.length;
	var sRowNum = "";
	for(var i=0;i<=iCellSuffixLength;i++)
	{
		if(sCellSuffix.charCodeAt(i)>=48&&sCellSuffix.charCodeAt(i)<=57)
		{
			sRowNum = sRowNum+sCellSuffix.charAt(i);
		}
	}					
	var iRowNumber = parseInt(sRowNum);	
	return iRowNumber;
}

function addGUIDToRow(oElement)
{
	//debugger;
	//debugger;
	if(oElement)
	{
		//Editing code not to touch CaseView Code should only be responsible for sending results back to CV not changing CV
		//Must act like a REST server
		//Create new GUID ID
		var ID = createGuid();
		//get the row with the GUID
		var oGuidElement = document.getElementById("guidrow");
		oGuidElement.value = ID;
		oElement.disabled=true;
		/*
		//}
		var sTableName = oElement.getAttribute("tablename");
		var sGuid = oElement.getAttribute("tempguid");
		var oTable = oDoc.tableByName(sTableName);
		if(oTable)
		{
			var iRowNumber = getRowNumber(oTable,sGuid,"TEMPGUID");				
			var oRow = oTable.getRow(iRowNumber);
			if(oRow)
			{
				//Create new GUID ID
				var ID = createGuid();
				//Add new GUID
				oRow.propSet("GUID",ID)
				//remove the temporary GUID;
				oRow.propDelete("TEMPGUID");
				//get the row with the GUID
				var oGuidElement = document.getElementById("guidrow");
				oGuidElement.value = ID;
				oElement.disabled=true;
			}
		}*/
	}
}


function getTableTypeDesc(sTableType)
{
	try{
		var sDescription = "";

		switch(sTableType)
		{
			case GENERIC_TABLE:
				sDescription = "Generic table";
				break;
			case GENERIC_TABLE_SCHED:
				sDescription = "Generic schedules table";
				break;
			case CONTROL_TABLE:
				sDescription = "Controls table";
				break;
			case SECTIONHEAD_TABLE:
				sDescription = "section Header table";
				break;			
			case SECTIONSUBHEAD_TABLE:
				sDescription = "section Header table";
				break;			
			case BS_TABLE:
				sDescription = "Balance sheet table";
				break;			
			case IS_TABLE:
				sDescription = "Income statemment table";
				break;			
			case DETIS_TABLE:
				sDescription = "Detailed Income Statement table";
				break;
			case DETIS_SUBTABLE	:
				sDescription = "Detailed income statement subtable";
				break;
			case CF_TABLE:
				sDescription = "Cash flow statement table";
				break;
			case EQUITY_TABLE:
				sDescription = "Statement of changes in equity";
				break;
			case FARM_TABLE:
				sDescription = "Farming table";
				break;
			case NOTECONTROL_TABLE:
				sDescription = "Note number / control table";
				break;
			case NOTE1_TABLE:
				sDescription = "Note type 1 table - Standard note table";
				break;
			case "6N2":
			case NOTE2_TABLE:
				sDescription = "Note type 2 table - Note Total table";
				break;
			case NOTE3_TABLE:
				sDescription = "Note type 3 table - balance check";
				break;
			case INVEST_NOTE_TABLE:
				sDescription = "Note Investment type table - used in investment to subs, JV's & Associates";
				break;
			case "6N9":
			case NOTE9_TABLE:
				sDescription = "Note type 9 table - carrying value table";
				break;
			case HORIZONTAL_NOTE_TABLE:
				sDescription = "Horizontal Note table - used in horizontal notes";
				break;
			case NOTE_HIDE_TABLE:
				sDescription = "Hidden note table used for linking";
				break;
			case NOTE_TOTAL3_TABLE:
				sDescription = "Note total";
				break;
			case CDIRTABLE:
				sDescription = "Custom property that indicates if directors table to be sorted";
				break;
			//table uses
			case MODTABLE:
				sDescription = "Table is modifiable and as such has the various controls available, e.g sort rows";
				break;
			case STATEMENTTABLE:
				sDescription = "Statement table e.g. Balance sheet";
				break;
			case NOTETABLE:
				sDescription = "Note Table";
				break;
			case HEADERCTRL:
				sDescription = "Header control table - The following table hold controls associated with the header";
				break;
			case HEADING_TABLE:
				sDescription = "Heading table";
				break;
			case FOOTER_TABLE:
				sDescription = "Footer table";
				break;
			case INDEX_TABLE:
				sDescription = "Index table";
				break;
		}
		return sDescription;
	}catch(e)
	{
		alert(e.description);
	}finally{
		
	}
}

function getRowTypeDesc(sRowType)
{
	try{
		var sDescription = "";

		switch(sRowType)
		{
			case CONTROL_ROW:
				sDescription = "Control row";
				break;
			case HEADING_ROW:
				sDescription = "Generic schedules table";
				break;
			case SUBHEADING_ROW:
				sDescription = "Subheading row";
				break;
			case CALC1_ROW:
				sDescription = "Calculation row  [grp]	->	grp(c1,ayg,entity(groupid,rc(-4)))";
				break;			
			case CALC2_ROW:
				sDescription = "Calculation row  [-grp]	->	-grp(c1,ayg,entity(groupid,rc(-4)))";
				break;			
			case CALC3_ROW:
				sDescription = "Calculation row [grpent,1]	->	grpent(c1,ayg,entity(groupid,rc(-4)),1)";
				break;			
			case CALC4_ROW:
				sDescription = "Calculation row [grpent,-1]	->	grpent(c1,ayg,entity(groupid,rc(-4)),-1)";
				break;			
			case CALC5_ROW:
				sDescription = "Cell reference calculation row / Multiple calculation row";
				break;
			case CALC6_ROW	:
				sDescription = "[grp] with Description based on map col less last 4 characters used in PPEV recon note";
				break;
			case CALC7_ROW:
				sDescription = "[grp] with additional terms and conditions paragraph below";
				break;
			case CALC8_ROW:
				sDescription = "pos of grp calcultion";
				break;
			case CALC9_ROW:
				sDescription = "pos of -grp calculation";
				break;
			case CALC10_ROW:
				sDescription = "Movement calculation - PY - CY";
				break;
			case CALCB_ROW:
				sDescription = "PY value pulls through to the CY col. PY col is input";
				break;
			case CALCC_ROW:
				sDescription = "Movement calculation - PY - CY -- Negative";
				break;			
			case INPUT_ROW:
				sDescription = "Input row";
				break;
			case TEXTONLY_ROW:
				sDescription = "Text only row - skips if row above is skipped";
				break;
			case TEXTCALC1_ROW:
				sDescription = "Input row with opening balance being last years total";
				break;
			case INPUTPERCENT_ROW:
				sDescription = "Input row with percentage columns";
				break;			
			case INPUTROLLFORWARD_ROW:
				sDescription = "Input row roll forward to PY";
				break;
			case INPUTDESCNOT_ROW:
				sDescription = "Input row, description columns not input";
				break;
			case INPUTBULLET_ROW:
				sDescription = "Input row with bullet, description pulls through from Map number";
				break;
			case INPUTTOTAL_ROW:
				sDescription = "Input row with last column being a total of other 3 columns";
				break;
			case INPUTANDTEXT_ROW:
				sDescription = "Input row with additional paragraph input text";
				break;
			//table uses
			case ACTCALC1_ROW:
				sDescription = "Calculation row pulling through from account numbers";
				break;
			case ACTCALC2_ROW:
				sDescription = "Negative Calculation row pulling through from account numbers";
				break;
			case SUBTOTAL_ROW:
				sDescription = "Subtotal row";
				break;
			case THINLINE_ROW:
				sDescription = "Thin line row";
				break;
			case THINKLINE_ROW:
				sDescription = "Thick line row";
				break;
			case TOTAL_ROW:
				sDescription = "Total row";
				break;
			case LINKTOTAL_ROW:
				sDescription = "Total row with linkage";
				break;
			case LINKSUBTOTAL_ROW:
				sDescription = "Total row based on row position in other tables";
				break;				
			case LINKTOTALSKIP_ROW:
				sDescription = "Total row with linkage --Always skipped and hidden";
				break;
			case YEARHEADER_ROW:
				sDescription = "Year heading row";
				break;
			case NOTE_ROW:
				sDescription = "Note linkage row";
				break;
			case PAGEBREAK_ROW:
				sDescription = "Page break row";
				break;
			case BALCHK_ROW:
				sDescription = "Balance check row";
				break;
			case CHEADDEFAULTPRINT:
				sDescription = "";
				break;
			case CHEADSKIPDEFAULT:
				sDescription = "";
				break;
			case CACCPOLLINK:
				sDescription = "sets which accouting policies are associated with this note";
				break;
			case HEADING_ROW:
				sDescription = "Heading row";
				break;
		}
		return sDescription;
	}catch(e)
	{
		alert(e.description);
	}finally{
		
	}
}

function getColumnTypeDesc(sColType)
{
	try{
		var sDescription = "";

		switch(sColType)
		{
			case CALC_COL:
				sDescription = "Calculation column";
				break;
			case SPACE_COL:
				sDescription = "Space column";
				break;
			case DESC_COL:
				sDescription = "Description column";
				break;
			case VARIANCE_COL:
				sDescription = "Variance column";
				break;			
			case TOTAL_COL:
				sDescription = "Total column";
				break;			
			case HIDDEN_COL:
				sDescription = "Hidden (controls) column";
				break;			
			case MAP_COL:
				sDescription = "Map Number column -- also the link to the Note";
				break;			
			case NOTEREF_COL:
				sDescription = "Note reference column";
				break;
			case PERCENT_COL	:
				sDescription = "Percentage Column";
				break;
			case LS_COL:
				sDescription = "Lead sheet column";
				break;
			case PRINT_CONT_COL:
				sDescription = "Print control column";
				break;
			case ST_REF_COL:
				sDescription = "Statement reference column";
				break;
			case CTRLCELLCOL:
				sDescription = "Column holding control cells";
				break;
			case ANNOTATION_COL:
				sDescription = "Annotations column";
				break;
			case PAGENO_COL:
				sDescription = "Page no. column";
				break;			
			case CACCBASIS_COL:
				sDescription = "Basis of accounting column";
				break;
			case PYC_COL:
				sDescription = "Company prior year column";
				break;
			case AYC_COL:
				sDescription = "Company current year column";
				break;
			case PYG_COL:
				sDescription = "Group prior year column";
				break;			
			case AYG_COL:
				sDescription = "Group current year column";
				break;
			case PYG2_COL:
				sDescription = "Group prior prior year column";
				break;
			case PYC2_COL:
				sDescription = "Company prior prior year column";
				break;	
			case COL_MAPPING_ROW:
				sDescription = "Columng mapping row"; //Columng mapping row
				break;
			}
		return sDescription;
	}catch(e)
	{
		alert(e.description);
	}finally{
		
	}
}

//returns all the tables in a section
function getTableinSection(sSection,oDoc,iSection)
{
	try{
		var aTables = new Array();
		var sTableName;
		var sTableCheck ="";
		if (oDoc)
		{
			if ((iSection>0 && typeof(iSection)=="integer") || sSection=="")
			{
				var oSection = oDoc.section(iSection);	
			}else{
				var oSection = oDoc.sectionByName(sSection);
			}
		}
		else{
			if ((iSection>0 && typeof(iSection)=="integer") || sSection=="")
			{
				var oSection = document.section(sSection);
			}
			else
			{
				var oSection = document.sectionByName(sSection);
			}
		}
		//check if section exits
		if (oSection)
		{
			var iFirstSectIndex = oSection.firstparaIndex;
			var iLastSectIndex = oSection.lastparaIndex;
			for (var j=iFirstSectIndex;j<iLastSectIndex;j++)
			{
				if (oDoc)
				{
					var oTable = oDoc.tableByParaIndex(j);
				}
				else
				{
					var oTable = document.tableByParaIndex(j);
				}
				if (oTable)
				{
					sTableName = oTable.getLabel;
					if (sTableCheck != sTableName)
					{
						aTables[aTables.length] = sTableName;
						sTableCheck = sTableName;
						j = oTable.lastParaIndex - 1;
					}
				}
			}
		}
		return aTables;
	}catch(e){
		alert(e.description);
	}finally{
		oDoc = null;
		oSection = null;
		oTable = null;
		aTables = null;
	}
}

function getTableData(oDoc,oSection)
{
	try{
		var sStr = "";
		//Show whether the section is being skipped or not
		var iSectionSkipCond = oSection.evaluateSkip();
		var sSection = oSection.label;
		var sColor = getSkipCondIndicator(oSection);
		var aTable = getTableinSection(sSection,oDoc);
		if(isInputValid(aTable))
		{
			var sTable = aTable[0];
			var oTable = oDoc.tableByName(sTable);
			if(oTable)
			{
				var sTableType = oTable.propGet(CTABLETYPE);
				
				//get table type description
				var sTableTypeDesc = getTableTypeDesc(sTableType);
				
				var sTableStr = getTableLineItems(oTable,sTableType,sSection,iSectionSkipCond,oDoc);
				if(isInputValid(sTableStr)){
					sStr = sTableStr;//"<li jumpcode='"+sSection+"' onclick='gotoSection(this);highlightElement(this);updatePreviewPane(this);' component='column' title='Component: Table row\nRow type: HEADING_ROW\nRow desc: Column heading row\nTable name: "+sTable+"\nTable type: "+sTableType+"\nTable desc: "+sTableTypeDesc+"' tablename='"+sTable+"' style='background-color:#ddd;'>[Column header row]</li><li style='color:"+sColor+";' jumpcode='"+sSection+"' onclick='gotoSection(this);highlightElement(this);updatePreviewPane(this);' component='mapcolumn' title='Component: Table row\nRow type: HEADING_ROW\nRow desc: Column mapping row\nTable name: "+sTable+"\nTable type: "+sTableType+"\nTable desc: "+sTableTypeDesc+"' tablename='"+sTable+"' style='background-color:#ddd;'>[Column mapping row]</li>"+sTableStr;

				}					
			}
		}
		return sStr;
	}catch(e)
	{
		alert(e.description);
	}finally{
		oSection = null;
		aTable = null;
		oTable = null;
	}
}

function getTableLineItems(oTable,sTableType,sSectionLabel,iSectionSkipCond,oDoc)
{
	try{		
		var sString = "";
		var ssTable = "";
		//debugger;
		//debugger;
		//Check table type
		switch (sTableType)
		{
			case "6N8":	
			case "N3":
			case "N2X":
			case "S1":
			case "S3":
			case "N11": // Hidden note table used for linking//NOTE_HIDE_TABLE:
			case "6G2":
			case "G2":
			case "J1":
			case "6N9":		
			case "6N2":
				var iRows = oTable.nRows();
				var sTableName = oTable.getLabel();
				var oStartCell = oDoc.cell(sTableName+".START");
				var oEndCell = oDoc.cell(sTableName+".END");
				if(!oStartCell||!oEndCell)
				{
					var iStartRow = 1;
					var iEndRow = oTable.nRows();
				}
				else{	
					//Get Start & End Row
					var iStartRow = parseInt(oStartCell.value)+1;
					var iEndRow = parseInt(oEndCell.value);
				}

				//Get the description column
				var aDescCol = getColWithProp(CCOLTYPE,DESC_COL,oTable);
				var aMapCol = getColWithProp(CCOLTYPE,MAP_COL,oTable);
				var aCalcCol = getColWithProp(CCOLTYPE,CALC_COL);
				var sDescCol = aDescCol[0];	
				var sMapCol = aMapCol[0];				
				var iColumn = getIntFromColLabel(sDescCol);
				var sTableTypeDesc = getTableTypeDesc(sTableType);
				//for(var iCurrentRowNumber=iStartRow;iCurrentRowNumber<=iEndRow;iCurrentRowNumber++)
				for(var iCurrentRowNumber=1;iCurrentRowNumber<=iEndRow;iCurrentRowNumber++)
				{				
					var oRow = oTable.getRow(iCurrentRowNumber);
					if(oRow.propGet(CROWTYPE)===CONTROL_ROW && !isInputValid(oRow.propGet("ROWTYPE")))
						continue;

					var iSkipCond = oRow.evaluateSkip();
					
					if(iSkipCond===1 || iSectionSkipCond===1)
						var sColor = '#0000FF';
					else
						var sColor = '#000000';
										
					//Get the row guid
					var sGuid = oRow.propGet("GUID");
					if(oRow && oRow.propExists("GUID")!=1){
						//generate temporary guid to the row to make it easy to access it
						var sGUIDID = createGuid();
						oRow.propSet("TEMPGUID") = sGUIDID;
					}
					var sRowType = "";
					//HEADING_ROW,COL_MAPPING_ROW
					if(isInputValid(oRow.propGet("ROWTYPE")))
					{
						var sRowType = oRow.propGet("ROWTYPE");
						var sRowTypeDesc = getRowTypeDesc(sRowType);
					}else{
						var sRowType = oRow.propGet("CROWTYPE");
						var sRowTypeDesc = getRowTypeDesc(sRowType);						
					}
					
					if(sRowType===TEXTONLY_ROW)
					{
						//debugger;
						//debugger;
						var iTableCellFirstPara = oTable.cellFirstParaIndex(iCurrentRowNumber, iColumn);
						var iTableCellLastPara = oTable.cellLastParaIndex(iCurrentRowNumber, iColumn);
						var oPara = oDoc.para(iTableCellFirstPara);
						var sDesc = oPara.getText();
						
						if(!isInputValid(sDesc))
						{
							sDesc = "Paragraph text";
						}
						else{
							var iTextLength = sDesc.length;
							if(iTextLength>50)
								sDesc = sDesc.substr(0,50)+"...";						
						}
						
						if(isInputValid(sDesc))
						{
							if(sGUIDID!="")
							{
								sString = sString + "<li style='background-color:#CCFFCC;color:"+sColor+";' component='para' paraindex='"+oPara.index+"' jumpcode='"+sSectionLabel+"' onclick='gotoSection(this);highlightElement(this);updatePreviewPane(this);' guid='"+sGuid+"'  id='row_"+sGuid+"' tempguid='"+sGUIDID+"' tablename='"+sTableName+"' objecttype='row' title='Component: Table row\nRow type: "+sRowType+"\nRow type desc: "+sRowTypeDesc+"\nTable name: "+sTableName+"\nTable type: "+sTableType+"\nTable desc: "+sTableTypeDesc+"'><span>["+sDesc+"]</span></li>";								
							
							}else{
								
								sString = sString + "<li style='display:table-cell;background-color:#CCFFCC;color:"+sColor+";' component='para' paraindex='"+oPara.index+"' jumpcode='"+sSectionLabel+"' onclick='gotoSection(this);highlightElement(this);updatePreviewPane(this);' guid='"+sGuid+"'  id='row_"+sGuid+"' tablename='"+sTableName+"' objecttype='row' title='Component: Table row\nRow type: "+sRowType+"\nRow type desc: "+sRowTypeDesc+"\nTable name: "+sTableName+"\nTable type: "+sTableType+"\nTable desc: "+sTableTypeDesc+"'><span>["+sDesc+"]</span></li>";
							}
						}
					}else if(sRowType==="HEADING_ROW"){
						
						sString = sString + "<li style='color:"+sColor+";' tablename='"+sTableName+"' jumpcode='"+sSectionLabel+"' onclick='gotoSection(this);highlightElement(this);updatePreviewPane(this);' guid='"+sGuid+"' id='row_"+sGuid+"' component='column' title='Component: Table row\nRow type: "+sRowType+"\nRow type desc: "+sRowTypeDesc+"\nTable name: "+sTableName+"\nTable type: "+sTableType+"\nTable desc: "+sTableTypeDesc+"' style='background-color:#ddd;background-color:#FFFFC6;'><span>[Column header row]</span></li>";
						
					}else if(sRowType==="COL_MAPPING_ROW"){

						sString = sString + "<li style='color:"+sColor+";' tablename='"+sTableName+"' jumpcode='"+sSectionLabel+"' onclick='gotoSection(this);highlightElement(this);updatePreviewPane(this);' guid='"+sGuid+"' id='row_"+sGuid+"' component='mapcolumn' title='Component: Table row\nRow type: "+sRowType+"\nRow type desc: "+sRowTypeDesc+"\nTable name: "+sTableName+"\nTable type: "+sTableType+"\nTable desc: "+sTableTypeDesc+"' style='background-color:#ddd;'>[Column mapping row]</li>"								
					}
					else
					{
						var sDescCell = getCellFromTableCell(sTableName,iColumn,iCurrentRowNumber,oDoc)[0];
						var oDescCell = oDoc.cell(sDescCell);
						if(oDescCell)
						{
							var sDesc = oDescCell.value;
							if(sDesc==="BR:YR1:Y1:SIGN:RND:SUB"||sDesc==="BR:YR0:Y1:SIGN:RND:SUB|0"||!isInputValid(sDesc))
								continue;
							
							if(sGUIDID!="")
							{
								//Check if the cell is input or not							
								if (oDescCell.input===1)
								{
									sString = sString + "<li style='background-color:#FFFFC6;color:"+sColor+"' component='row' cellnumber='"+oDescCell.number+"' guid='"+sGuid+"' tempguid='"+sGUIDID+"' mapcolumn='"+sMapCol+"' tablename='"+sTableName+"' jumpcode='"+sSectionLabel+"' onclick='gotoSection(this);highlightElement(this);updatePreviewPane(this);' id='row_"+sGuid+"' objecttype='row' title='Component: Table row\nRow type: "+sRowType+"\nRow type desc: "+sRowTypeDesc+"\nTable name: "+sTableName+"\nTable type: "+sTableType+"\nTable desc: "+sTableTypeDesc+"'><span>["+sDesc+"]</span></li>";								
								}else{
									sString = sString + "<li style='color:"+sColor+";' component='row' cellnumber='"+oDescCell.number+"' guid='"+sGuid+"' mapcolumn='"+sMapCol+"' tempguid='"+sGUIDID+"' tablename='"+sTableName+"' jumpcode='"+sSectionLabel+"' onclick='gotoSection(this);highlightElement(this);updatePreviewPane(this);' id='row_"+sGuid+"' objecttype='row' title='Component: Table row\nRow type: "+sRowType+"\nRow type desc: "+sRowTypeDesc+"\nTable name: "+sTableName+"\nTable type: "+sTableType+"\nTable desc: "+sTableTypeDesc+"'><span>["+sDesc+"]</span></li>";	
								}
							}else{
								//Check if the cell is input or not							
								if (oDescCell.input===1)
								{
									sString = sString + "<li style='background-color:#FFFFC6;color:"+sColor+";' component='row' cellnumber='"+oDescCell.number+"' guid='"+sGuid+"' mapcolumn='"+sMapCol+"' tablename='"+sTableName+"' jumpcode='"+sSectionLabel+"' onclick='gotoSection(this);highlightElement(this);updatePreviewPane(this);' id='row_"+sGuid+"' objecttype='row' title='Component: Table row\nRow type: "+sRowType+"\nRow type desc: "+sRowTypeDesc+"\nTable name: "+sTableName+"\nTable type: "+sTableType+"\nTable desc: "+sTableTypeDesc+"'><span>["+sDesc+"]</span></li>";								
								}else{
									sString = sString + "<li style='display:table-cell;color:"+sColor+";' component='row' cellnumber='"+oDescCell.number+"' guid='"+sGuid+"' mapcolumn='"+sMapCol+"' tablename='"+sTableName+"' jumpcode='"+sSectionLabel+"' onclick='gotoSection(this);highlightElement(this);updatePreviewPane(this);' id='row_"+sGuid+"' objecttype='row' title='Component: Table row\nRow type: "+sRowType+"\nRow type desc: "+sRowTypeDesc+"\nTable name: "+sTableName+"\nTable type: "+sTableType+"\nTable desc: "+sTableTypeDesc+"'><span>["+sDesc+"]</span></li>";	
								}								
							}
						}						
					}
				}
				if(iSectionSkipCond===1)
					var sColor = '#0000FF';
				else
					var sColor = '#000000';				
				
				sString = ssTable + sString;	
				break;
			default:
		}

		return sString;
	}catch(e)
	{
		alert(e.description);
	}finally{
		oDescCell = null;
		oPara = null;
		oCell = null;
		oStartCell = null;
		oEndCell = null;
		oTable = null;
	}
}	

//method to get a row number based on its type
function getRowWithProp(sProp,sValue,oTable)
{
	try
	{
		var aRows = new Array();
		var iRow = oTable.nRows();
		for (row=1;row<=iRow;row++)
		{
			var oRow = oTable.getRow(row);
			var sRowCType = oRow.propGet(sProp);//getProperty.getProp(sProp,oRow)
			if (sValue==sRowCType)
			{
				aRows[aRows.length] = row-1;
			}		
		}
		return aRows;
	}
	catch(e)
	{
	}	
}

//method to get a column based on its type
function getColWithProp(sProp,sValue,oTable)
{	
	try
	{
		var iCol = oTable.nColumns();
		var aCols = new Array()
		for ( var col=1;col<=iCol;col++)
		{
			var oCol = oTable.getColumn(col);
			var sColCType = oCol.propGet(sProp);//getProperty.getProp(sProp,oCol)
			if (sValue==sColCType)
			{
				aCols[aCols.length] = GetColnumber(col);
			}			
		}
		return aCols;
	}
	catch(e)
	{
	}
}
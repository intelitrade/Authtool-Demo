//Returns an array of cell names in a table cell.
function getCellFromTableCell(sTable,iColumn,iRow,oDoc)
{
	/*if (!oDoc)
	{
		var oDoc = document
	}*/

	var oTable = oDoc.tableByName(sTable)
	if (oTable)
	{
		var oTableCell = oTable.getCell(iColumn,iRow)
		if (oTableCell)
		{
			//create an array to store the cells in the table cell
			var aCells = new Array()
			//get the first and last paragraph of the table cell
			var iTableCellFirstParaIndex = oTable.cellFirstParaIndex(iRow,iColumn) 

			var iTableCellLastParaIndex = oTable.cellLastParaIndex(iRow,iColumn)

			for (var i=iTableCellFirstParaIndex;i<=iTableCellLastParaIndex;i++)
			{
				//get each paragraph and check if there is a cell on that paragraph
				var oPara = oDoc.para(i)
				if (oPara)
				{
					var iTextLength = oPara.textLength()
					var sTempCellName = ""
					for (var j=1;j<iTextLength;j++)
					{
						var oObj = oPara.getObjectAt(j)
						if(oObj && oObj.typeid==1)
						{
							var oCell = oObj.cell
							if (oCell && oCell.number != sTempCellName)
							{
								//add the cell to an array
								aCells[aCells.length] = oCell.number
								sTempCellName = oCell.number
							}
						}
					}
				}
			}
		}
	}
	return aCells
}

function onMouseDownSections(sId)
{
	try{
		event.cancelBubble = true;
		oElement = document.getElementById(sId);
		var oParentNode = oElement.parentNode;
		iPos = getTheElementsPosition(oParentNode,oElement);
		//Get original settigs of the element before moving it around
		var iOriginalZIndex = oElement.style.zIndex;
		var iElementLeft = oElement.style.left;
		var iElementTop = oElement.style.top;	
		// (1) start the process
		// (2) prepare to moving: make absolute and on top by z-index
		oElement.style.position = 'absolute';
		oElement.style.zIndex = 1000;
		// ...and put that absolutely positioned element under the cursor
		oElement.onmousemove = function ()
		{
			//moveAt(event.pageX, event.pageY);
			moveAt(event.clientX, event.clientY);
		}			
		
		// centers the oElement at (pageX, pageY) coordinates
		function moveAt(pageX, pageY)
		{
			if(oElement)
			{	
				oElement.style.left = pageX - oElement.offsetWidth + 'px';
				oElement.style.top = pageY - oElement.offsetHeight + 'px';
				oRefElement = document.elementFromPoint(event.clientX, event.clientY);
				if(oRefElement)
				{
					var sElementId = oRefElement.id;
					var sElementName = oElement.name;
					var sElementUse = oRefElement.use;
					var sRefElementTypeOfSection = oRefElement.typeofsection;
					var sElementTypeOfSection = oElement.sectiontype;	
					if( (sRefElementTypeOfSection==sElementTypeOfSection))//"NOTECONTROL"||sTypeOfSection=="SUBNOTECONTROL"||sTypeOfSection=="EXPANDCOLLAPSE"||sTypeOfSection=="SUBNOTECONTROLGROUP"||sTypeOfSection=="EXPANDCOLLAPSECOMPANY3RDYEAR"))sElementId!="" && typeof(sElementId)!="undefined" && sElementId!=null &&
					{
						sRefElement = sElementId;
						oElement.style.color="green";
						var oRefParent = oRefElement.parentNode;
						var iParentNode = oRefParent.children.length;
						
						for(var i=0;i<iParentNode;i++)
						{
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
					}else{
						oElement.style.color="red";
					}
				}
			}
		}	
		
		function onMouseMove(event) {
			moveAt(event.clientX, event.clientY);
		}

		if(document.addEventListener)
		{ 
			document.addEventListener('mousemove', onMouseMove);
		}else
		{
			document.attachEvent("onmousemove", onMouseMove);		
		}	
	}catch(e)
	{
		alert(e.description);
	}
}	
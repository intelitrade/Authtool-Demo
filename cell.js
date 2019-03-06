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

// Helper function to get an element's exact position
function getElementsXYPosition(el)
{
	var xPos = 0;
	var yPos = 0;

	while (el){
		if (el.tagName == "BODY")
		{
			// deal with browser quirks with body/window/document and page scroll
			var xScroll = el.scrollLeft || document.documentElement.scrollLeft;
			var yScroll = el.scrollTop || document.documentElement.scrollTop;

			xPos += (el.offsetLeft - xScroll + el.clientLeft);
			yPos += (el.offsetTop - yScroll + el.clientTop);
		} else
		{
			// for all other non-BODY elements
			xPos += (el.offsetLeft - el.scrollLeft + el.clientLeft);
			yPos += (el.offsetTop - el.scrollTop + el.clientTop);
		}
		el = el.offsetParent;
	}
	return{
		x: xPos,
		y: yPos
	};
}

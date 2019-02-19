function getTheElementsPosition(oParentNode,oElement)
{
	try{
		var iParentNode = oParentNode.children.length;
		for(var i=0;i<iParentNode;i++)
		{
			if(oElement===oParentNode.children[i])
			{
				iPos = i;
				break;
			}
		}
		return iPos;
	}catch(e)
	{
		alert(e.description);
	}
}

function onMouseDownRows(sId)
{
	event.cancelBubble = true;
	oElement = document.getElementById(sId);
	var oParentNode = oElement.parentNode;
	iPos = getTheElementsPosition(oParentNode,oElement);
	//debugger;
	//debugger;
	//Get original settigs of the element before moving it around
	var iOriginalZIndex = oElement.style.zIndex;
	var iElementLeft = oElement.style.left;
	var iElementTop = oElement.style.top;
	
	// (4) onmouseup, remove unneeded handlers
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
				
				//document.getElementById("results").innerHTML = oElement.style.left;	
				if(sElementId!="" && typeof(sElementId)!="undefined" && sElementId!=null && sElementId.search("row_")!=-1||sElementUse=="tablecontainer")
				{
					//oElement.setAttribute("refelement", sElementId);
					sRefElement = sElementId;
					oElement.style.color="green";
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
					//oRefElement.style.textDecoration = "underline";
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
}	
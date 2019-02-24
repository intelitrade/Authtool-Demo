function getParaData(oDoc,oSubSection)
{
	try{
		var sStr = "";
		var iFirstParaIndex = oSubSection.firstParaIndex;
		var sSubSection = oSubSection.label;
		var sCtypeVal = oSubSection.propGet("CTYPE");
		var oPara = oDoc.para(iFirstParaIndex);
		var sParaText = oPara.getText();
		var iSkipCond = oPara.evaluateSkip();
		var iHideCond = oPara.evaluateHide();
		var sParaGuid = oPara.propGet("GUID");
		if(iSkipCond===1 || iHideCond===1)
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
		//var sStr = "<ul><li component='inputtextsection' id='"+sSubSection+"' objecttype='section' jumpcode='"+sSubSection+"' onclick='gotoSection(this);highlightElement(this);highlightSection(this.jumpcode);updatePreviewPane(this);updateCurrentSelection(this);' guid='"+sParaGuid+"' style='background-color:#CCFFCC;color:"+sColor+";'>"+sParaText+"</li></ul>";
		sStr = "<li component='inputtextsection' id='"+sSubSection+"' objecttype='section' jumpcode='"+sSubSection+"' onclick='gotoSection(this);highlightElement(this);updatePreviewPane(this);updateCurrentSelection(this);' guid='"+sParaGuid+"' style='background-color:#CCFFCC;color:"+sColor+";' ctype='"+sCtypeVal+"' title='Component: Paragraph\nParagraph index: "+oPara.index+"'>"+sParaText+"</li>";	
		
		return sStr;
	}catch(e)
	{
		alert(e.description);
	}finally{
		
	}
}

function InsertNewParagraph(oElement,sRefElement)
{
	try{
		debugger;
		debugger;

		if (oElement && isInputValid(sRefElement)) {
			//Check if the element need to be added or note
			if(oElement.style.color=="green")
			{
				var sElementName = oElement.name;
				var oRefElement = document.getElementById(sRefElement);
				var iParaToInsert = 1;//prompt("Please enter number of "+sElementName+" row(s) to insert", 1);
				if(iParaToInsert==null)
				{
					oElement.style.position = 'static';
					oElement.onmousemove = null;		
				}else{
					
					//get rid of the on mouse move event
					oElement.onmousemove = null;
					oElement.style.position = 'static';						

					//Create a progress bar to show number of rows being inserted and show when each row has been inserted
					var oProgBar = oDoc.createProgressBar((iParaToInsert>1?"Inserting paragraph":"Inserting a paragraph"),iParaToInsert,1)
					//if(!oRefElement)	
					for(var x=1;x<=iParaToInsert;x++)
					{
						var ID = createGuid();
						var sParaSectionId = "UL_"+ID;
						//var sParaSectionId = "section_"+ID;//oSubSection.index;
						var oNewParaNode = document.createElement("UL");
						oNewParaNode.setAttribute("id",sParaSectionId);
						oNewParaNode.setAttribute("use","paracontainer");
						oNewParaNode.setAttribute("paratype",oElement.paratype);
						//oNewParaNode.setAttribute("typeofsection",oElement.tabletype);
						//oNewParaNode.setAttribute("title","Component: Table\nTable name: "+sElementName+"\nTable label: "+sSubSectionId+"\nTable index: \nVersion: \nTable type: "+oElement.sectiontype);
						oNewParaNode.setAttribute("component","inputtextsection");
						oNewParaNode.setAttribute("objecttype","section");
						oNewParaNode.style.backgroundColor="#CCFFCC";
						//style='background-color:#CCFFCC;color:"						
						
						//use="tablecontainer" typeofsection="SUMMARYTABLE" tablename="MFF"
						var oNewParaListItemNode = document.createElement("LI");
						//var sParaId = "para_"+ID;
						//oNewParaListItemNode.setAttribute("id",sParaId);
						oNewParaListItemNode.setAttribute("id","HEADER_"+sParaSectionId);
						oNewParaListItemNode.setAttribute("isheader","true");
						oNewParaListItemNode.setAttribute("jumpcode","SECTION_"+sParaSectionId);
						oNewParaNode.setAttribute("jumpcode","SECTION_"+sParaSectionId);



						
						//Attach events to the list item, event for jump to code 
						oNewParaListItemNode.setAttribute("onclick","highlightElement(this);updatePreviewPane(this);");
						oNewParaListItemNode.onclick = function() {highlightElement();updatePreviewPane(); };

						//Set the type section being inserted - currently 2 propertues are being used one of them will be dropped
						oNewParaNode.setAttribute("typeofsection",oElement.sectiontype);
						oNewParaListItemNode.setAttribute("typeofsection",oElement.sectiontype);

						//Add a guid to for the new row
						oNewParaListItemNode.setAttribute("guid",ID);
						//Give the new row a name which user can identify with, name it after the element it is inserted from
						oNewParaListItemNode.setAttribute("name",sElementName);
						
						//oChildNode.style.color=sColor;
						oNewParaNode.setAttribute("title","Component: Paragraph\nSection name: "+sElementName+"\nSection label: "+sParaSectionId+"\nVersion: \nParatype type: "+oElement.paratype);

						oNewParaListItemNode.setAttribute("component","inputtextsection");
						oNewParaListItemNode.setAttribute("objecttype","section");
						oNewParaListItemNode.setAttribute("use","sectioncontainer");	
						oNewParaListItemNode.setAttribute("add","true");	

						oNewParaNode.setAttribute("parentsection",oRefElement.parentsection);
						oNewParaNode.setAttribute("parentguid",oRefElement.parentguid);
						
						oNewParaListItemNode.setAttribute("parentsection",oRefElement.parentsection);
						oNewParaListItemNode.setAttribute("parentguid",oRefElement.parentguid);						
						
						
						
						

						
						oNewParaListItemNode.innerHTML=sElementName+"<sup style='color:red;font-size: smaller;'>New*</sup>";
						//oNewParaListItemNode.style.backgroundColor = "#F5DEB3";
						//oNewParaListItemNode.style.fontWeight = "bold";
						oNewParaListItemNode.style.padding="2px";
						//Append to the parent element oRefElement.parentNode.parentNode.id
						oNewParaNode.appendChild(oNewParaListItemNode);

						var oNewElement = insertPara(oNewParaNode, oRefElement.parentNode);
						
						if(oNewElement)
						{
							oNewElement.setAttribute("add","true");
							highlightElement(oNewElement.children[0]);
							oNewElement.scrollIntoView();
						}
					
						oProgBar.updateProgress(1);	
						oProgBar.setMessage("Paragraph "+x+" / "+iParaToInsert);						
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

/*
function InsertNewParagraph(oElement,sRefElement)
{
	try{
		debugger;
		debugger;
		if (oElement && isInputValid(sRefElement)) {
			//Check if the element need to be added or note
			if(oElement.style.color=="green")
			{
				var sElementName = oElement.name;
				var oRefElement = document.getElementById(sRefElement);
				var iParaToInsert = 1;//prompt("Please enter number of "+sElementName+" row(s) to insert", 1);
				if(iParaToInsert==null)
				{
					oElement.style.position = 'static';
					oElement.onmousemove = null;		
				}else{
					
					//get rid of the on mouse move event
					oElement.onmousemove = null;
					oElement.style.position = 'static';						

					//Create a progress bar to show number of rows being inserted and show when each row has been inserted
					var oProgBar = oDoc.createProgressBar((iParaToInsert>1?"Inserting paragraph":"Inserting a paragraph"),iParaToInsert,1)
					//if(!oRefElement)	
					for(var x=1;x<=iParaToInsert;x++)
					{
						var ID = createGuid();
						var sParaSectionId = "section_"+ID;//oSubSection.index;
						var oNewParaNode = document.createElement("UL");
						oNewParaNode.setAttribute("id",sParaSectionId);
						oNewParaNode.setAttribute("use","paracontainer");
						oNewParaNode.setAttribute("paratype",oElement.paratype);
						//oNewParaNode.setAttribute("typeofsection",oElement.tabletype);
						//oNewParaNode.setAttribute("title","Component: Table\nTable name: "+sElementName+"\nTable label: "+sSubSectionId+"\nTable index: \nVersion: \nTable type: "+oElement.sectiontype);
						oNewParaNode.setAttribute("component","para");
						oNewParaNode.setAttribute("objecttype","para");
						oNewParaNode.style.backgroundColor="#CCFFCC";
						//style='background-color:#CCFFCC;color:"						
						
						//use="tablecontainer" typeofsection="SUMMARYTABLE" tablename="MFF"
						var oNewParaListItemNode = document.createElement("LI");
						var sParaId = "para_"+ID;
						oNewParaListItemNode.setAttribute("id",sParaId);

						oNewParaListItemNode.innerHTML=sElementName+"<font style='color:red;weight:bold'> (New paragraph*)</font>";
						//oNewParaListItemNode.style.backgroundColor = "#F5DEB3";
						//oNewParaListItemNode.style.fontWeight = "bold";
						//oNewParaListItemNode.style.padding="10px";
						//Append to the parent element oRefElement.parentNode.parentNode.id
						oNewParaNode.appendChild(oNewParaListItemNode);

						var oNewElement = insertPara(oNewParaNode, oRefElement.parentNode);
						
						if(oNewElement)
						{
							oNewElement.setAttribute("add","true");
							highlightElement(oNewElement.children[0]);
							oNewElement.scrollIntoView();
						}
					
						oProgBar.updateProgress(1);	
						oProgBar.setMessage("Paragraph "+x+" / "+iParaToInsert);						
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
}*/

function onMouseDownPara(sId)
{
	try{
		event.cancelBubble = true;
		oParaElement = document.getElementById(sId);
		//var oPosition = getElementsXYPosition(oParaElement)
		var oParentNode = oParaElement.parentNode;
		iPos = getTheElementsPosition(oParentNode,oParaElement);
		//debugger;
		//debugger;
		//Get original settigs of the element before moving it around
		var iOriginalZIndex = oParaElement.style.zIndex;
		var iElementLeft = oParaElement.offsetLeft;//oParaElement.style.left;
		var iElementTop = oParaElement.offsetTop;//oParaElement.style.top;
		
		// (4) onmouseup, remove unneeded handlers
		// (1) start the process
		// (2) prepare to moving: make absolute and on top by z-index
		oParaElement.style.position = 'absolute';
		oParaElement.style.zIndex = 1000;
		// ...and put that absolutely positioned element under thesElementUse=="tablecontainer" cursor
		oParaElement.onmousemove = function ()
		{
			//moveAt(event.pageX, event.pageY);
			moveAtPara(event.clientX, event.clientY);
		}			
		
		// centers the oParaElement at (pageX, pageY) coordinates
		function moveAtPara(pageX, pageY)
		{
			if(oParaElement)
			{	
				oParaElement.style.left = pageX - oParaElement.offsetWidth + 'px';
				oParaElement.style.top = pageY - oParaElement.offsetHeight + 'px';
				oRefElement = document.elementFromPoint(event.clientX, event.clientY);
				if(oRefElement)
				{
					var sElementId = oRefElement.id;
					var sElementName = oParaElement.name;
					var sElementUse = oRefElement.use;
					var sRefElementSectionType = oRefElement.typeofsection;
					//document.getElementById("insertoptions").innerHTML = sElementId+" "+oRefElement.typeofsection+" "+oParaElement.sectiontype+" - "+event.type;
					//document.getElementById("results").innerHTML = oParaElement.style.left;	
					if(isInputValid(sElementId) && sElementId.search("HEADER_")!=-1 && sElementUse!="table" && sElementUse!="para")
					{
						sRefElement = sElementId;
						oParaElement.style.color="green";
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
						oParaElement.style.color="red";
					}
				}
			}
		}	
		
		function onMouseMovePara(event) {
			moveAtPara(event.clientX, event.clientY);
		}

		if(document.addEventListener)
		{ 
			document.addEventListener('mousemove', onMouseMovePara);
		}else
		{
			document.attachEvent("onmousemove", onMouseMovePara);		
		}	
	}catch(e)
	{
		alert(e.description);
	}finally{
		oParaElement.style.left = "auto";
		oParaElement.style.top = "auto";
		oParaElement.style.color="";		
	}
}
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
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

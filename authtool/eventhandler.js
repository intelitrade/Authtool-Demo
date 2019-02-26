/*
The following file contains functions that handle different types of events
*/
//Cancel button event handler
function cancelButtonClick()
{
	try
	{
		//close the dialog
		//No value needs to be returned;
		//window.returnValue;
		window.close();
	}
	catch(e)
	{
		alert(e.description);
	}
}
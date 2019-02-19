/*
The following file contains function that handle different types of events
*/
//Cancel button event handler
function cancelButtonClick()
{
	try
	{
		//close the dialog
		window.returnValue = 0
		window.close()
	}
	catch(e)
	{
		alert(e.description);
	}
}
//Save to disk
function SaveToDisk(sPath,sContent)
{
	try{
	
		var ofso = new ActiveXObject("Scripting.FileSystemObject");
		var ofileDest = ofso.CreateTextFile(sPath, true);
		if (ofileDest)
		{
		   ofileDest.Write(sContent);//document.documentElement.outerHTML);
		   ofileDest.close();
		}
		else
		{
		   alert("Unable to create file " + sPath);
		}
	}catch(e)
	{
		
	}finally{
		
	}
} 

/*

<HTML>
<HEAD>
<script LANGUAGE="JavaScript">
    function SaveToDisk(sUrl, sPath)
    {
        var x = new ActiveXObject("Msxml2.XMLHTTP.4.0");
        x.open("GET", sUrl, false);
        x.send();
        if (x.status == 200)
        {
          var fso = new ActiveXObject("Scripting.FileSystemObject");
          var fileDest = fso.CreateTextFile(sPath, true, true);
          if (fileDest)
          {
             fileDest.Write(x.responseText);
             fileDest.close();
          }
          else
          {
             alert("unable to create file " + sPath);
          }
        }
        else
        {
          alert("HTTP error status " + x.status + " from " + sUrl);
        }

    }        
</script>
</HEAD>

<BODY onload="SaveToDisk('http://www.p2p.wrox.com', 'c:\\temp\\123.htm');">
<P>The rest of the page is here...</P>
</BODY>
</HTML>

*/
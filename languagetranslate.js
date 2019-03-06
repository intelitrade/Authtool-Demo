function translateAPI(textAPI,langAPI)
{
	try{
		debugger;
		debugger;
		
		var url = "https://translate.yandex.net/api/v1.5/tr.json/translate",
			keyAPI = "trnsl.1.1.20130922T110455Z.4a9208e68c61a760.f819c1db302ba637c2bea1befa4db9f784e9fbb8";//"trnsl.1.1.20190304T101653Z.d5dd800eccf8d0ef.14cee982f1f984b73bd5c413db9f9fac57f35744"; //

		//document.querySelector('#translate').addEventListener('click', function() {
			var xhr = new XMLHttpRequest(),
				textAPI = "hello"//document.querySelector('#source').value,
				//langAPI = document.querySelector('#lang').value
				data = "key="+keyAPI+"&text="+textAPI+"&lang="+langAPI;
			xhr.open("POST",url,true);
			xhr.setRequestHeader("Content-type","application/x-www-form-urlencoded");
			xhr.send(data);
			//xhr.onreadystatechange = function() {
				//if (this.readyState==4 && this.status==200) {
					debugger;
					debugger;
					var res = this.responseText;
					//document.querySelector('#json').innerHTML = res;
					var json = JSON.parse(res);
					if(json.code == 200) {
						document.querySelector('#output').innerHTML = json.text[0];
					}
					else {
						document.querySelector('#output').innerHTML = "Error Code: " + json.code;
					}
				//}
			//}
		//}, false);	
	}catch(e)
	{
	
	}finally{
	
	}
}

String.prototype.translate = function (){
	
	if(this.toString()=="Note control"){
		return "Nota beheer";
	}else if(this.toString()=="Property, plant and equipment"){
		return "Eiendom, aanleg en toerusting";
	}
}

Array.prototype.translate = function (){
	
	for(var i=0;i<this.length;i++)
	{
		if(this[i]=="Note control"){
			this[i] = "Nota beheer";
		}else if(this.toString()=="Property, plant and equipment"){
			this[i] = "Eiendom, aanleg en toerusting";
		}
	}
	
	return this;
}
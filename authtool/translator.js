/* Translator-js v1.0.0 | @zalomea | MIT/GPL2 Licensed */
var Translator = function(options){

	var t = {
		lang : null,
		dict : [],
		htmlfriend : false
    }

	if(typeof options == 'undefined'){
		options = {};
	}

	if(typeof options.language != 'undefined'){
		t.lang = options.language;
	}
	if(typeof options.dictionary != 'undefined'){
		t.dict = options.dictionary;
	}
	if(typeof options.htmlfriendly != 'undefined' && options.htmlfriendly){
		t.htmlfriend = true;
	}

	var translate = function(index){
		if(t.lang == null){
			return index;
		}else if(t.dict[t.lang] != null){
			if(t.dict[t.lang][index] != null){
				return t.dict[t.lang][index];
			}
		}
		return index;
	};

	t.setLanguage = function(language){
		t.lang = language;
	};
	t.setLanguageRun = function(language){
		t.setLanguage(language);
		t.runTranslation();
	};

	t.addTranslation = function(language, index, translation){
		if(typeof translation == 'undefined'){
			t.dict[language] = index;
		}else{
			if(typeof t.dict[language] == 'undefined'){
				t.dict[language] = [];
			}
			t.dict[language][index] = translation;
		}
	};

	t.runTranslation = function() {
		var elements = [];
		if(t.htmlfriend){
			elements = document.documentElement.getElementsByTagName('*');
			//debugger;
			//debugger;
			var iLength = elements.length;
			var oProgBar = oDoc.createProgressBar("Translating...",iLength,1);	
			for (var i = 0; i < iLength; i++) {
				//if()
				if(elements.item(i) && elements.item(i).getAttribute('data-translate') != null){
					if(elements.item(i).getAttribute('data-translate')=="Addn - Owned")
					{
						//debugger;
						//debugger;
					}
					elements.item(i).innerHTML = translate(elements.item(i).getAttribute('data-translate'));
				}
				oProgBar.updateProgress(1);
			}
			oProgBar.destroyProgressBar();
		}else{
			elements = document.getElementsByTagName('translatag');
			for (var i = 0; i < elements.length; i++) {
				elements.item(i).innerHTML = translate(elements.item(i).getAttribute('data-translate'));
			}
		}

	};

	if(typeof options.autostart != 'undefined' && options.autostart){
		if(window.addEventListener)
			window.addEventListener('load', t.runTranslation, false);
		else if(window.attachEvent)
			window.attachEvent('onload', t.runTranslation);
	}

	return t;
};
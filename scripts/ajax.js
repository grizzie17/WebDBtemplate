

var loadedobjects = ""
var rootdomain = "http://" + window.location.host

function ajaxPage(url, containerid) {
	var page_request = false;
	if (window.XMLHttpRequest) // if Mozilla, Safari etc
	{
		page_request = new XMLHttpRequest();
	}
	else if (window.ActiveXObject) // if IE
	{
		try {
			page_request = new ActiveXObject("Msxml2.XMLHTTP");
		}
		catch (e) {
			try {
				page_request = new ActiveXObject("Microsoft.XMLHTTP");
			}
			catch (e) {
				return false;
			}
		}
	}
	else {
		return false;
	}

	var sURL = url;
	var sUID = "zzz=" + escape((new Date()).toUTCString());
	sURL += (sURL.indexOf("?") + 1) ? "&" : "?";
	sURL += sUID;

	page_request.onreadystatechange = function () {
		ajaxHandleStateChange(page_request, containerid);
	}

	page_request.sTargetURL = url;
	window.status = "Request to Load = " + url;
	page_request.open('GET', sURL, true);
	page_request.send(null);
}

function ajaxHandleStateChange(page_request, containerid) {
	if (page_request.readyState == 4) {
		if (page_request.status == 200
				|| window.location.href.indexOf("http") == -1) {
			document.getElementById(containerid).innerHTML = page_request.responseText;
			window.status = "loaded external url";
			setTimeout("ajaxClearStatus()", 3 * 1000);
		}
		else {
			window.status = "retry loading external url";
			page_request.open("GET", page_request.sTargetURL, true);
			page_request.send(null);
		}
	}
}

function ajaxClearStatus() {
	window.status = "";
}

function loadobjs() {
	if (!document.getElementById)
		return;
	for (i = 0; i < arguments.length; i++) {
		var file = arguments[i];
		var fileref = "";
		if (loadedobjects.indexOf(file) == -1) //Check to see if this object has not already been added to page before proceeding
		{
			if (file.indexOf(".js") != -1) //If object is a js file
			{
				fileref = document.createElement('script');
				fileref.setAttribute("type", "text/javascript");
				fileref.setAttribute("src", file);
			}
			else if (file.indexOf(".css") != -1) //If object is a css file
			{
				fileref = document.createElement("link");
				fileref.setAttribute("rel", "stylesheet");
				fileref.setAttribute("type", "text/css");
				fileref.setAttribute("href", file);
			}
		}
		if (fileref != "") {
			document.getElementsByTagName("head").item(0).appendChild(fileref);
			loadedobjects += file + " "; //Remember this object as being already added to page
		}
	}
}





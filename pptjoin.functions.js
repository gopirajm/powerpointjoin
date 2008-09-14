function $(id) {
	return document.getElementById(id);
}

function init() {
	window.resizeTo(500, 300);
	$('step1').innerHTML = STRINGS.STEP1;
	$('step2').innerHTML = STRINGS.STEP2;
	$('total-slides').innerHTML = STRINGS.TOTAL_SLIDES;
	document.title = STRINGS.APP_TITLE;
}
function validateFile(e) {
	var filename = e.value;
	debug("validating " + filename);
	if(filename.substr(filename.length - 3) == "txt") {
		debug("Selected Text file = " + filename);
	} else if(filename.length == 0) {
		debug("user hit cancel");
	} else {
		alert(sprintf(STRINGS.INVALID_FILE_CHOICE, filename));
	}
}

function processFiles() {
	var filename = $('filename').value;
	debug("Reading from file " + filename);
	var fso = new ActiveXObject("Scripting.FileSystemObject");
	var objPpt;
	var pptMain;
	try {
		objPpt = new ActiveXObject("Powerpoint.Application");
		objPpt.Presentations.Add();
		objPpt.Visible = true;
		debug("PowerPoint.Application");
		var iMainCharts = 0;
		debug("NewPresentation");
		if(fso.FileExists(filename)) {
			var txt = fso.OpenTextFile(filename, 1);
			debug("OpenTextFile");
			var directory = fso.GetParentFolderName(filename);
			var pptMainFilename = directory + "\\" + STRINGS.COMBINED_CHARTS_FILENAME;
			debug("pptMainFilename = " + pptMainFilename);
			fso.CopyFile(txt.ReadLine(), pptMainFilename);
			pptMain = objPpt.Presentations.Open(pptMainFilename);
			debug("Open");

			//pptMain.SaveAs(pptMainFilename);
			if(! fso.FileExists(pptMainFilename) ) {
				debug("Did not save correctly");
				return;
			} else {
				debug("Saved");
			}
			while (! txt.AtEndOfStream) {
				try {
					var line = txt.ReadLine();
					debug("Reading line " + line);
					var ppt = line;
					if(! fso.FileExists(ppt)) {
						debug("it ain't " + ppt);
						if(! fso.FileExists(directory + "\\" + ppt) ) {
							alert(sprintf(STRINGS.CHARTS_NOT_FOUND, ppt));
							break;
						} else {
							ppt = directory + "\\" + ppt;
							debug("must be " + ppt);
						}
					}
					debug("  Found ppt '" + ppt + "'");
					iMainCharts = pptMain.Slides.Count;
					updateTotal(iMainCharts);
					debug("iMainCharts = " + iMainCharts);
					debug("Inserting " + ppt);
					pptMain.Slides.InsertFromFile(directory + "\\" + ppt, iMainCharts);
					debug("InsertFromFile");
					$('step3').innerHTML = sprintf(STRINGS.STEP3, STRINGS.COMBINED_CHARTS_FILENAME);
				} catch (e) {
					//alert("Oh No!\n" + e.description);
					break;
				}
			}
			updateTotal(pptMain.Slides.Count);
			debug("Finished");
			txt = null;
		} else {
			debug("Can't find this file:  " + filename);
		}
	} catch (e) {
		alert("OH NO! \n" + e.description + "\n" + sprintf(STRINGS.CLOSE_COMBINED, STRINGS.COMBINED_CHARTS_FILENAME));
	} 
	pptMain.Save();
	pptMain.Close();
	debug("closed");
	pptMain = null;
	objPpt.Quit();
	objPpt = null;
	fso = null;
}

function openCharts() {
	var txtfile = $('filename').value;
	var pptcharts = STRINGS.COMBINED_CHARTS_FILENAME;
	var fso = new ActiveXObject("Scripting.FileSystemObject");
	var directory = fso.GetParentFolderName(txtfile);
	if(! fso.FileExists(pptcharts)) {
		pptcharts = directory + "\\" + pptcharts;
	}
	fso = null;
	var sh = new ActiveXObject("WScript.Shell");
	debug("opening " + pptcharts);
	sh.Run(pptcharts);
}

function openFolder() {
	var txtfile = $('filename').value;
	var fso = new ActiveXObject("Scripting.FileSystemObject");
	var directory = fso.GetParentFolderName(txtfile);
	fso = null;
	var sh = new ActiveXObject("WScript.Shell");
	var cmd = STRINGS.PATH_TO_EXPLORER + " " + directory;
	debug("Running " + cmd);
	sh.Run(cmd);
	sh = null;
}

function debug(msg) {
	$('debug').innerHTML = $('debug').innerHTML + "<BR>" + msg;
}

function showHelp(msg) {
	$('help-message').innerHTML = msg;
	$('help').style.display = "";
}

function hideHelp() {
	$('help').style.display = 'none';
}

function updateTotal(n) {
	$('total-charts').innerHTML = n;
}


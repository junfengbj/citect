// Define execute result
var CONST_CHARS_IN_LINE 		= 78;
var CONST_SEPARATE_LINE_1		= "  ----------------------------------------------------------------------------";
var CONST_SEPARATE_LINE_2		= "    --------------------------------------------------------------------------";
var CONST_SEPARATE_LINE_3		= "      ------------------------------------------------------------------------";
var CONST_SEPARATE_LINE_4		= "        ----------------------------------------------------------------------";
var CONST_LIST_LEVEL_1			= "  ";
var CONST_LIST_LEVEL_2			= "    ";
var CONST_LIST_LEVEL_3			= "      ";
var CONST_LIST_LEVEL_4			= "        ";

var CONST_RESULT_OK 			= " [  OK  ]";
var CONST_RESULT_ERROR 			= " [ ERROR]";
var CONST_RESULT_FAILED 		= " [ FAIL ]";

// To include the common module
var objFSO = WScript.CreateObject("Scripting.FileSystemObject");

var objFolder = objFSO.GetFolder(".");

var strName, strDate;

WScript.Echo(objFolder.Path);
WScript.Echo(CONST_SEPARATE_LINE_1);

var colFolder = new Enumerator(objFolder.SubFolders);

for (; !colFolder.atEnd(); colFolder.moveNext()) {
	strName = colFolder.item().Name;
	strDate = colFolder.item().DateCreated;
	strDate = FormatDate(strDate);

	if (strName.indexOf(strDate) > -1) {
		continue;
	}

	try {
		objFSO.moveFolder(strName, strDate + "." + strName);

		DisplayResult(CONST_LIST_LEVEL_1 + strDate + "." + strName.substring(0, 40) + " ", CONST_RESULT_OK, false);
	} catch (e) {
		DisplayResult(CONST_LIST_LEVEL_1 + strDate + "." + strName.substring(0, 40) + " ", CONST_RESULT_FAILED, false);
	}
}

WScript.Echo(CONST_SEPARATE_LINE_1);




//********************************************************************
//* Function: DisplayResult
//*
//* Purpose: Display a result message in entire row.
//*
//* Input:
//*  [in]    strMessage		The message which will be displayed.
//*  [in]    strResult		The result which will be displayed.
//*  [in]    blnStdOut		True = WScript.Std.Out, False = WScript.Echo
//*  [in]    strSperator	The character of sperator.
//* Output:
//*  [out]	 none.
//*
//********************************************************************
function DisplayResult(strMessage, strResult, blnStdOut, strSperator) {
	var j = 0;
	var k = CONST_CHARS_IN_LINE - strResult.length;
	var m = 0;

	var strSperator = arguments[3] ? arguments[3] : ".";

	strMessage = StringPadding(strMessage, CONST_CHARS_IN_LINE, strSperator);

	m = strMessage.length

	for (var i = 0; i < m; i++) {
		j++;

		if (strMessage.substr(i, 1).charCodeAt() > 256) {
			j++;
		}

		if (j > k) {
			strMessage = strMessage.substr(0, i);
			i = m;
		}
	}

	strMessage = strMessage + strResult;

	if (blnStdOut == true) {
		WScript.StdOut.Write(strMessage + "\r");
	} else {
		WScript.Echo(strMessage);
	}
}

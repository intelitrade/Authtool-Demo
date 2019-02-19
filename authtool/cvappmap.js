//XJW: CaseView application scripting object property/method mapping

//XJW: ICaseViewAppOld "application" object property mapping:
//     - "this" in this scope is the global object
//     - for readability, declare "var global = (function(){return this}());", then "global" can be used to to reference the global object
//     - well, make sure that your script does not have the global variables with the names below (the names below are reserved)

Object.defineProperty(this, 'baseUnit', { get: function() {return application.baseUnit;}, set: function(x) { application.baseUnit = x; }});
Object.defineProperty(this, 'programPath', { get: function() {return application.programPath;}, set: function(x) { application.programPath = x; }});
Object.defineProperty(this, 'majorVersion', { get: function() {return application.majorVersion;}, set: function(x) { application.majorVersion = x; }});
Object.defineProperty(this, 'minorVersion', { get: function() {return application.minorVersion;}, set: function(x) { application.minorVersion = x; }});
Object.defineProperty(this, 'buildVersion', { get: function() {return application.buildVersion;}, set: function(x) { application.buildVersion = x; }});
Object.defineProperty(this, 'releaseNumber', { get: function() {return application.releaseNumber;}, set: function(x) { application.releaseNumber = x; }});
Object.defineProperty(this, 'scriptBackwardCompatible', { get: function() {return application.scriptBackwardCompatible;}, set: function(x) { application.scriptBackwardCompatible = x; }});
Object.defineProperty(this, 'appBaseUnitType', { get: function() {return application.appBaseUnitType;}, set: function(x) { application.appBaseUnitType = x; }});
Object.defineProperty(this, 'platformName', { get: function() {return application.platformName;}, set: function(x) { application.platformName = x; }});
Object.defineProperty(this, 'dbChangeFlags', { get: function() {return application.dbChangeFlags;}, set: function(x) { application.dbChangeFlags = x; }});
Object.defineProperty(this, 'dbChangeFlagsTimer', { get: function() {return application.dbChangeFlagsTimer;}, set: function(x) { application.dbChangeFlagsTimer = x; }});
Object.defineProperty(this, 'showRecalError', { get: function() {return application.showRecalError;}, set: function(x) { application.showRecalError = x; }});
Object.defineProperty(this, 'Document', { get: function() {return application.Document();}});  //XJW: special handling to work with existing erroneous syntax "Document.something" in our existing Active Script engine where "poperty_get" is used to retrieve the document object
Object.defineProperty(this, 'document', { get: function() {return application.Document();}});  //XJW: same as above but with lower case 'd' at the begin
Object.defineProperty(this, 'Application', { get: function () { return application; } });  //XJW: map 'Application' to 'application'

//XJW: ICaseViewAppOld "application" object method mapping:
//     - well, make sure that your script does not have the global functions with the names below (the names below are reserved as well)

//XJW: get the document by function; has to be renamed as "Document" is mapped to a property; see my comments above
function DocumentF()
{
	return application.Document();
}

function OpenDocument(File, Flags)
{
	return application.OpenDocument(File, Flags);
}

function renameFile(curName, newName)
{
	return application.renameFile(curName, newName);
}

function runProgram(Program, Parameters, DefaultPath)
{
	application.runProgram(Program, Parameters, DefaultPath);
}

function exitScript()
{
	application.exitScript();
}

function OpenBgDocument(File, Flags, Timeout)
{
	return application.OpenBgDocument(File, Flags, Timeout);
}

function CloseBgDocument(File, ignoreRefCount)
{
	application.CloseBgDocument(File, ignoreRefCount);
}

function GetBgDocument(File)
{
	return application.GetBgDocument(File);
}

function getDocumentVersion(sCvFileName)
{
	return application.getDocumentVersion(sCvFileName);
}

function runHTMLDialog(sHTMLFileName)
{
	return application.runHTMLDialog(sHTMLFileName);
}

function runHTMLDialog2(sHTMLFileName, Flags)
{
	return application.runHTMLDialog2(sHTMLFileName, Flags);
}

//CP: CV-16477 - Added missing entry for runHTMLDialogCef.
function runHTMLDialogCef(sHTMLFileName, Flags, isModal, dialogParam)
{
    return application.runHTMLDialogCef(sHTMLFileName, Flags, isModal, dialogParam);
}

function getFontList()
{
	return application.getFontList();
}

function getKliSettings(sKliDbFullName, sUserName)
{
	return application.getKliSettings(sKliDbFullName, sUserName);
}

function repackageKli(sSrcKliFile, sSrcKliFilter, sDstKliFile, sDstKliLabel, sDstNewAcctMngPswd, bDstCopyRefUserOnly, bDstDelFldContent)
{
	return application.repackageKli(sSrcKliFile, sSrcKliFilter, sDstKliFile, sDstKliLabel, sDstNewAcctMngPswd, bDstCopyRefUserOnly, bDstDelFldContent);
}

function newKliDatabase(sFileName, sLabel, sAcctMngPswd, sUserName, sUserPswd)
{
	return application.newKliDatabase(sFileName, sLabel, sAcctMngPswd, sUserName, sUserPswd);
}

function fileExists(name)
{
	return application.fileExists(name);
}

function folderExists(name)
{
	return application.folderExists(name);
}

function copyFile(src, dest, overwrite)
{
	application.copyFile(src, dest, overwrite);
}

function moveFile(src, dest)
{
	application.moveFile(src, dest);
}

function deleteFile(name)
{
	application.deleteFile(name);
}

function getSpecialFolder(folder)
{
	application.getSpecialFolder(folder);
}

function createFolder(name)
{
	application.createFolder(name);
}

function deleteFolder(name)
{
	application.deleteFolder(name);
}

function copyFolder(source, dest, overwrite)
{
	application.copyFolder(source, dest, overwrite);
}

function moveFolder(source, dest, verwrite)
{
	application.moveFolder(source, dest, verwrite);
}

function getTempName(len)
{
	return application.getTempName(len);
}

function getFileName(path)
{
	return application.getFileName(path);
}

function getFolder(path)
{
	return application.getFolder(path);
}

function getFile(path)
{
	return application.getFile(path);
}

function CreateTextFile(Path, Flags)
{
	return application.CreateTextFile(Path, Flags);
}

function OpenTextFile(Path, iomode, Flags)
{
	return application.OpenTextFile(Path, iomode, Flags);
}

function GetDocumentXMLStorageCount(File)
{
	return application.GetDocumentXMLStorageCount(File);
}

function GetDocumentXMLStorageReadOnly(File, index)
{
	return application.GetDocumentXMLStorageReadOnly(File, index);
}

function getMSXMLVersion(bMajorVerOnly)
{
	return application.getMSXMLVersion(bMajorVerOnly);
}

function abbreviatePathName(sPathName, iMaxLength, bNameAtLeast)
{
	return application.abbreviatePathName(sPathName, iMaxLength, bNameAtLeast);
}

//XJW: deprecated method
function addActiveXObj(sObjectProgId)
{
	return application.addActiveXObj(sObjectProgId);
}

//XJW: deprecated method
function delActiveXObj(sObjectProgId)
{
	application.delActiveXObj(sObjectProgId);
}

function isDocumentOpen(sFullFileName)
{
	return application.isDocumentOpen(sFullFileName);
}

function reverseEscape(sString)
{
	return application.reverseEscape(sString);
}

function getTaxonomyFolder()
{
	return application.getTaxonomyFolder();
}

function setTaxonomyFolder(taxFld)
{
	return application.setTaxonomyFolder(taxFld);
}

function getXBRLDefLang()
{
	return application.getXBRLDefLang();
}

function setXBRLDefLang(languageCode)
{
	application.setXBRLDefLang(languageCode);
}

function getStringFromID(iStringId)
{
	return application.getStringFromID(iStringId);
}

function zipFiles(sArchiveName, vSourceFileList)
{
	return application.zipFiles(sArchiveName, vSourceFileList);
}

function unzipFiles(sArchiveName, sTargetFolder, bDirectories)
{
	return application.unzipFiles(sArchiveName, sTargetFolder, bDirectories);
}

function sendEmail(sName, sAddress, sSubject, sText, sAttachment)
{
	return application.sendEmail(sName, sAddress, sSubject, sText, sAttachment);
}

function RegQueryValue(sRoot, sKey, sValue)
{
	return application.RegQueryValue(sRoot, sKey, sValue);
}

function GetNewGUID()
{
	return application.GetNewGUID();
}

function getColorValRGB(Red, Green, Blue)
{
	return application.getColorValRGB(Red, Green, Blue);
}

function getRColorVal(colorVal)
{
	return application.getRColorVal(colorVal);
}

function getGColorVal(colorVal)
{
	return application.getGColorVal(colorVal);
}

function getBColorVal(colorVal)
{
	return application.getBColorVal(colorVal);
}

function clearRepositoryCache(ClientPath)
{
	application.clearRepositoryCache(ClientPath);
}


var importNames = JavaImporter();
importNames.importPackage(Packages.com.ms.com);
importNames.importPackage(Packages.com.ms.activeX);
importClass(Packages.com.jacob.activeX.ActiveXComponent);
importClass(Packages.com.jacob.com.Dispatch);
importClass(Packages.com.jacob.com.Variant);

function NovaDatovaZprava() {
    var ax = new ActiveXComponent("Fischl.ELO.ISDS.Client.ActiveX");
    var ao = ax.getObject();

    var s = Dispatch.call(ao, "DoJavaCreateMessage");
}

function OdpovedetNaDatovouZpravu() {
    var ax = new ActiveXComponent("Fischl.ELO.ISDS.Client.ActiveX");
    var ao = ax.getObject();

    var activeView = workspace.getActiveView();
    var selected = activeView.getFirstSelected();

    var sExt = selected.getSord().getDocVersion().getExt().toLowerCase();
    var sId = selected.getId();

    if (sExt == "zfo") {
        Dispatch.call(ao, "DoJavaReplyMessage", sId);
    }
    else
        workspace.showInfoBox("Chyba", "Chyba, toto není ZFO dokument.");
}

function OveritDatovouZpravu() {
    var ax = new ActiveXComponent("Fischl.ELO.ISDS.Client.ActiveX");
    var ao = ax.getObject();

    var activeView = workspace.getActiveView();
    var selected = activeView.getFirstSelected();

    var sExt = selected.getSord().getDocVersion().getExt().toLowerCase();
    var sFile = selected.getFile().getPath();

    if (sExt == "zfo")
        Dispatch.call(ao, "DoJavaAuthenticateMessage", sFile);
    else
        workspace.showInfoBox("Chyba", "Chyba, toto není ZFO dokument.");
}

function eloScriptButton101Start() {
    NovaDatovaZprava();
}

function eloScriptButton102Start() {
    OdpovedetNaDatovouZpravu();
}

function eloScriptButton103Start() {
    OveritDatovouZpravu();
}

function eloScriptButton104Start() {
    NovaDatovaZpravaAPI();
}

function getScriptButtonPositions() {
    return "101,home,Datové schránky;102,home,Datové schránky;103,home,Datové schránky;104,home,Datové schránky";
}

function getScriptButton101Name() {
    return "Nová datová zpráva";
}

function getScriptButton102Name() {
    return "Odpovědět na datovou zprávu";
}

function getScriptButton103Name() {
    return "Ověřit datovou zprávu";
}

function getScriptButton104Name() {
    return "Nová datová zpráva (API)";
}

function getExtraBands() {
    return "home,16,Datové schránky";
}

function eloComIsdsClearClipboardItems(param) {
    var objIds = param.split("|");
    var objIdsCount = objIds.length;
    for (var i = 0; i < objIdsCount; i++) {
        clipboard.removeId(objIds[i]);
    }
}

function eloComIsdsGetClipboardItems(param) {
    var enumElements = clipboard.getElements();
    var s = "<schranka>";
    while (enumElements.hasMoreElements()) {
        var schrankaDoc = enumElements.nextElement();
        var line = "<line sID=\"$1\" sGUID=\"$2\" sName=\"$3\" sMask=\"$4\"/>";
        line = line.replace("$1", schrankaDoc.getId());
        line = line.replace("$2", schrankaDoc.getSord().getGuid());
        line = line.replace("$3", escape(schrankaDoc.getName()));
        line = line.replace("$4", escape(schrankaDoc.getDocMaskName()));

        s += line + "\n";
    }
    s += "</schranka>";
    return s;
}

function eloComIsdsGetUsersGroups(param) {
    var userGroups = workspace.getUserGroups();
    var resultString = userGroups.join("|");
    return resultString;
}

function eloComIsdsGetUsersGroupID(param) {
    return archive.lookupUserId(param);
}

function eloComIsdsGetMetadataForReply(param) {
    var activeView = workspace.getActiveView();
    var selected = activeView.getFirstSelected();

    var sID = selected.getId();
    var dbIDSender = selected.getObjKeyValue("DBIDSENDER");
    var dmAnnotation = selected.getName();
    var dmLegalTitleLaw = selected.getObjKeyValue("DMLEGALTITLELAW");
    var dmLegalTitlePar = selected.getObjKeyValue("DMLEGALTITLEPAR");
    var dmLegalTitlePoint = selected.getObjKeyValue("DMLEGALTITLEPOINT");
    var dmLegalTitleSect = selected.getObjKeyValue("DMLEGALTITLESECT");
    var dmLegalTitleYear = selected.getObjKeyValue("DMLEGALTITLEYEAR");
    var dmSenderIdent = selected.getObjKeyValue("DMSENDERIDENT");
    var dmSenderRefNumber = selected.getObjKeyValue("DMSENDERREFNUMBER");
    var dmSenderType = selected.getObjKeyValue("DMSENDERTYPE");

    var xml = "<reply>";
    xml += "<dbIDSender>" + escape(dbIDSender) + "</dbIDSender>";
    xml += "<dmAnnotation>" + escape(dmAnnotation) + "</dmAnnotation>";
    xml += "<dmLegalTitleLaw>" + escape(dmLegalTitleLaw) + "</dmLegalTitleLaw>";
    xml += "<dmLegalTitlePar>" + escape(dmLegalTitlePar) + "</dmLegalTitlePar>";
    xml += "<dmLegalTitlePoint>" + escape(dmLegalTitlePoint) + "</dmLegalTitlePoint>";
    xml += "<dmLegalTitleSect>" + escape(dmLegalTitleSect) + "</dmLegalTitleSect>";
    xml += "<dmLegalTitleYear>" + escape(dmLegalTitleYear) + "</dmLegalTitleYear>";
    xml += "<dmSenderIdent>" + escape(dmSenderIdent) + "</dmSenderIdent>";
    xml += "<dmSenderRefNumber>" + escape(dmSenderRefNumber) + "</dmSenderRefNumber>";
    xml += "<dmSenderType>" + escape(dmSenderType) + "</dmSenderType>";
    xml += "</reply>";

    return xml;
}

function eloComIsdsGetBaseAddress(param) {
    return "http://localhost:8001/Fischl.ELO.ISDS.Service/DSEngineService/";
}

function eloComIsdsGetAddressBook(param) {

    var xml = "<addressBook>";

 	{
        var line = "<line sID=\"$1\" sName=\"$2\" sAddress=\"$3\"/>";
        line = line.replace("$1", "j73aet6");
        line = line.replace("$2", "MPO");
        line = line.replace("$3", "Na Františku 1039/32, 11015 Praha 1, CZ");
        xml += line + "\n";
    }

    xml += "</addressBook>";

    return xml;
}

function NovaDatovaZpravaAPI() {
    var ax = new ActiveXComponent("Fischl.ELO.ISDS.Client.ActiveX");
    var ao = ax.getObject();

    var xml = "<message>";

    xml += "<envelope>";
    xml += "<dbIDSender>" + "Ing. Karel Fischl (OVM)" + "</dbIDSender>";
    xml += "<dbIDRecipient>" + "uzcaety" + "</dbIDRecipient>";
    xml += "<dmRecipientType>" + "20" + "</dmRecipientType>";
    xml += "<dmAnnotation>" + "1003" + "</dmAnnotation>";
    xml += "<dmLegalTitleLaw>" + "1004" + "</dmLegalTitleLaw>";
    xml += "<dmLegalTitlePar>" + "1005" + "</dmLegalTitlePar>";
    xml += "<dmLegalTitlePoint>" + "1006" + "</dmLegalTitlePoint>";
    xml += "<dmLegalTitleSect>" + "1007" + "</dmLegalTitleSect>";
    xml += "<dmLegalTitleYear>" + "1008" + "</dmLegalTitleYear>";
    xml += "<dmSenderIdent>" + "1009" + "</dmSenderIdent>";
    xml += "<dmSenderRefNumber>" + "1010" + "</dmSenderRefNumber>";
    xml += "<dmRecipientIdent>" + "1011" + "</dmRecipientIdent>";
    xml += "<dmRecipientRefNumber>" + "1012" + "</dmRecipientRefNumber>";
    xml += "<dmSenderType>" + "OVM" + "</dmSenderType>";
    xml += "</envelope>";
    xml += "<attachments>";
    xml += "<line sID=\"2437\" sGUID=\"(FDBC90C3-2D67-47D6-A29E-BDB0DA1920C6)\" sName=\"130\" sMask=\"140\"/>";
    xml += "<line sID=\"2435\" sGUID=\"(79D17237-86D3-4315-A0F2-7D6FD5FB2158)\" sName=\"230\" sMask=\"240\"/>";
    xml += "<line sID=\"2430\" sGUID=\"(BEE6B537-100B-471E-B717-1852ADA18806)\" sName=\"330\" sMask=\"340\"/>";
    xml += "</attachments>";

    xml += "</message>";

    var s = Dispatch.call(ao, "DoJavaCreateMessageApi", xml);

// JavaScript source code
function FormLoad(executionContext) {

    try {
    debugger;
    var formContext = executionContext.getFormContext();

    formContext.getAttribute("name").setValue(formContext.getAttribute("ordernumber").getValue());
    //קישור להזמנה במערכת ERP
    //prmElement
    HideFieldsOrder(formContext);
    //formContext.data.entity.save();
    //createDocFolder(); //Remove later
    onLoadDocFrame(formContext);//Remove later
    addOnChangefunc(formContext);
    //getUsersLanguage(formContext);

    } catch (err) {
        alert("FormLoad : \n" + err.message);
    }

}
function getUsersLanguage(formContext) {
    //returns the LCID value that represents the Microsoft Dynamics CRM Language Pack that is the user selected as their preferred language
    var language = Xrm.Utility.getGlobalContext().userSettings.languageId
}

function addOnChangefunc(formContext) {
    formContext.getAttribute("bnt_ssp").addOnChange(saveSalesOrderForm);
}

function saveSalesOrderForm(executionContext) {
    var formContext = executionContext.getFormContext();
    formContext.data.entity.save();
}

function HideFieldsOrder(formContext) {
    // Xrm.Page.ui.tabs.get("FirstTab").sections.get("SectionOne").setVisible(false);
    formContext.getControl("name").setVisible(false);
    formContext.getControl("transactioncurrencyid").setVisible(false);
    //Xrm.Page.getControl("customerid").setVisible(false);
    formContext.getControl("pricelevelid").setVisible(false);
    // Xrm.Page.getControl("bnt_docurl").setVisible(false);
    // Xrm.Page.ui.tabs.get("summary_tab").sections.get("summary_tab_section_8").setVisible(false);
    //Xrm.Page.ui.tabs.get("summary_tab").sections.get("summary_tab_section_5").setVisible(false);
    formContext.ui.tabs.get("tab_3").setVisible(false);
    //Xrm.Page.ui.tabs.get("tab_4").setVisible(false);
}


function onLoadDocFrame(formContext) {
    debugger;
    if (formContext == null) {
        formContext = Xrm.Page;
    }
    var recordId = formContext.data.entity.getId().replace("{", "").replace("}", "");//Entity record id
    var oTypeCode = "1088";//formContext.context.getQueryStringParameters().etc;//Entity Type code. for lead it is 4
    var IFrame = "IFRAME_NewDocuments";//Iframe Name
    //IFRAME_orderDoc
    //var IFrame = "IFRAME_orderDoc";
    var relName = "areaSPDocuments"; //Entity Relation name. Like for Lead it is Lead_SharepointDocument
    SetIframeContent(IFrame, oTypeCode, relName, recordId, formContext);
}
async function SetIframeContent(iframeObjId, objectType, areaName, recordId, formContext) {
    var iframeObject = formContext.getControl(iframeObjId);//Get iFrame control
    if (iframeObject != null) {
        //Set iframe URL
        var globalContext = Xrm.Utility.getGlobalContext();
        //var app = await globalContext.getCurrentAppProperties();
        //var formId = formContext.ui.formSelector.getCurrentItem().getId().replace('{', '').replace('}', '');
        iframeObject.setSrc("/userdefined/areas.aspx?inlineEdit=1&oId=%7b" + recordId + "%7d&oType=" + objectType + "&pagemode=iframe&rof=true&security=852023&tabSet=" + areaName + "&theme=Outlook15White");
        //iframeObject.setSrc("/userdefined/areas.aspx?appid=" + app.appId + "&formid=" + formId + "&inlineEdit=1&navItemName=Documents&oId=%7b" + recordId + "%7d&oType=" + objectType + "&pagemode=iframe&rof=true&security=852023&tabSet=" + areaName + "&theme=Outlook15White");

    }
}
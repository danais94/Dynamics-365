// JavaScript source code
function startWorkflowForm(formContext) {
    try {
        formContext.data.entity.save();
        var sspVal = formContext.data.entity.attributes.get("bnt_ssp").getValue();
        if (sspVal == null) {
            alert("אנא הזן SSP");
            return;
        }
        if (window.confirm('? האם את/ה בטוח/ה שברצונך לשגר הזמנה')) {
            ExecuteWF(formContext);
        }
    } catch (err) {
        alert("startWorkflowForm : \n" + err.message);
    }
}

function ExecuteWF(formContext) {
    try {

        formContext.getAttribute("bnt_sendorder").setValue(true);

        if (formContext.getAttribute('bnt_docurl').getValue() == null ||
            formContext.getAttribute('bnt_docurl').getValue().length < 20) {
            formContext.getAttribute("bnt_docurl").setValue("...");
        }
         
        formContext.data.entity.save();

        var entityId = formContext.data.entity.getId();
        var options = {};
        options["entityName"] = "salesorder";
        options["entityId"] = entityId;
        Xrm.Navigation.openForm(options).then(
            function () {
                debugger;
                var Opp = formContext.getAttribute("opportunityid").getValue();
                if (Opp != null) {
                    var options = {};
                    options["entityName"] = "opportunity";
                    options["entityId"] = Opp[0].id;
                    Xrm.Navigation.openForm(options);
                }
            }, function (error) { alert(error); });

    } catch (err) {
        alert("ExecuteWF : \n" +err.message);
    }

}

function startWorkflow(formContext, guid, showMailMessage) {
    var url = Xrm.Utility.getGlobalContext().getClientUrl();

    var workflowId = getWorkflowId(formContext, "Order-SendMailToSSP");
    var OrgServicePath = "/XRMServices/2011/Organization.svc/web";
    url = url + OrgServicePath;
    var request;
    request = "<s:Envelope xmlns:s=\"http://schemas.xmlsoap.org/soap/envelope/\">" +
        "<s:Body>" +
        "<Execute xmlns=\"http://schemas.microsoft.com/xrm/2011/Contracts/Services\" xmlns:i=\"http://www.w3.org/2001/XMLSchema-instance\">" +
        "<request i:type=\"b:ExecuteWorkflowRequest\" xmlns:a=\"http://schemas.microsoft.com/xrm/2011/Contracts\" xmlns:b=\"http://schemas.microsoft.com/crm/2011/Contracts\">" +
        "<a:Parameters xmlns:c=\"http://schemas.datacontract.org/2004/07/System.Collections.Generic\">" +
        "<a:KeyValuePairOfstringanyType>" +
        "<c:key>EntityId</c:key>" +
        "<c:value i:type=\"d:guid\" xmlns:d=\"http://schemas.microsoft.com/2003/10/Serialization/\">" + guid + "</c:value>" +
        "</a:KeyValuePairOfstringanyType>" +
        "<a:KeyValuePairOfstringanyType>" +
        "<c:key>WorkflowId</c:key>" +
        "<c:value i:type=\"d:guid\" xmlns:d=\"http://schemas.microsoft.com/2003/10/Serialization/\">" + workflowId + "</c:value>" +
        "</a:KeyValuePairOfstringanyType>" +
        "</a:Parameters>" +
        "<a:RequestId i:nil=\"true\" />" +
        "<a:RequestName>ExecuteWorkflow</a:RequestName>" +
        "</request>" +
        "</Execute>" +
        "</s:Body>" +
        "</s:Envelope>";

    var req = new XMLHttpRequest();
    req.open("POST", url, true)
    // Responses will return XML. It isn't possible to return JSON.
    req.setRequestHeader("Accept", "application/xml, text/xml, */*");
    req.setRequestHeader("Content-Type", "text/xml; charset=utf-8");
    req.setRequestHeader("SOAPAction", "http://schemas.microsoft.com/xrm/2011/Contracts/Services/IOrganizationService/Execute");
    if (showMailMessage == true) {
        req.onreadystatechange = function () { assignResponse(formContext, req); };
    }
    req.send(request);
}

function assignResponse(formContext, req) {
    try {
        if (req.readyState == 4) {
            if (req.status == 200) {
                //alert('הדואר תשלח בהקדם');
                var subgrid = formContext.ui.controls.get("Orders");
                if (subgrid != null) {
                    subgrid.refresh();
                }
            }
            else if (req.status == 500) {
                alert('ארעה שגיאה בעת יצירת ההזמנה, אנא פנה למנהל המערכת');
            }
        }

    } catch (err) {
        alert("assignResponse : \n" + err.message);
    }
}

function getWorkflowId(formContext, workflowName) {
    var odataSelect = Xrm.Utility.getGlobalContext().getClientUrl() +
        '/XRMServices/2011/OrganizationData.svc/WorkflowSet?$select=WorkflowId&$filter=StateCode/Value eq 1 and ParentWorkflowId/Id eq null and Name eq \''
        + workflowName + '\'';
    var xmlHttp = new XMLHttpRequest();
    xmlHttp.open("GET", odataSelect, false);
    xmlHttp.send();
    if (xmlHttp.status == 200) {
        var result = xmlHttp.responseText;
        var xmlDoc = parseXml(result);
        return xmlDoc.getElementsByTagName("d:WorkflowId")[0].childNodes[0].nodeValue;
    }
}

function parseXml(xmlStr) {
    if (window.DOMParser) {
        return (new window.DOMParser()).parseFromString(xmlStr, "text/xml");
    }
    else if (typeof window.ActiveXObject != "undefined" && new window.ActiveXObject("Microsoft.XMLDOM")) {
        debugger;
        var xmlDoc = new window.ActiveXObject("Microsoft.XMLDOM");
        xmlDoc.async = "false";
        xmlDoc.loadXML(xmlStr);
        return xmlDoc;
    }
    else {
        return null;
    }
    return null;
}

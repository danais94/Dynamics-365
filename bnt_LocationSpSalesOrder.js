// JavaScript source code
var firstTime = true;

function onLoadMisgeret(executionContext) {
    try {

        var formContext = executionContext.getFormContext();
         

 debugger;


        if (formContext.getAttribute('bnt_docurl').getValue() == null ||
            formContext.getAttribute('bnt_docurl').getValue().length < 20) {

            formContext.getAttribute("bnt_docurl").setValue("...");
            formContext.data.save();
            //createDocFolder(formContext);
        }

       // setTimeout(function () {
             

       

        //getLocation(formContext);
        //formContext.data.save();
        //onLoadDocFrame(formContext);
        /*
        }, 3000);
        */
    } catch (err) {
        alert("onLoadMisgeret : \n" + err.message);
    }
}

//Function to be called on onload of the page.
var myFormType = null;
function createDocFolder(formContext) {
    try {
        if (formContext == null) {
            formContext = Xrm.Page;
        }

        var globalContext = Xrm.Utility.getGlobalContext();

        debugger;

        if (getDocsLocation(formContext) == 0) { // calls this method only when new record is saved.
            var clientUrl = globalContext.getClientUrl();
            var url = clientUrl + "/XRMServices/2011/Organization.svc/web";
            var xmlhttp = new XMLHttpRequest();
            xmlhttp.open("POST", url);
            xmlhttp.setRequestHeader("Accept", "application/xml, text/xml, */*");
            xmlhttp.setRequestHeader("Content-Type", "text/xml; charset=utf-8");
            xmlhttp.setRequestHeader("SOAPAction", "http://schemas.microsoft.com/xrm/2011/Contracts/Services/IOrganizationService/Execute");
            xmlhttp.onreadystatechange = function () {

              //  alert("xmlhttp.status = " + xmlhttp.status);

                if (xmlhttp.readyState == 4 && xmlhttp.status != 200) {

                    //alert("ישנו עומס זמני נא לסגור הזמנה ולפתוח שוב");
                    formContext.getAttribute("bnt_docurl").setValue("...");
                    formContext.data.save();
                    if (firstTime) {
                        firstTime = false;
                        setTimeout(function () {
                            if (formContext.getAttribute('bnt_docurl').getValue() == null ||
                                formContext.getAttribute('bnt_docurl').getValue().length < 20) {

                                createDocFolder(formContext);

                                setTimeout(function () {
                                    getLocation(formContext);
                                    formContext.data.save();
                                    onLoadDocFrame(formContext);
                                }, 3000);
                            }
                        }, 3000);
                    }
                }
                else if (xmlhttp.readyState == 4 && xmlhttp.status == 200) {

                    debugger;
                    ///SaveForm
                    ///UPDATE iFRAME
                    formContext.getAttribute("bnt_docurl").setValue("...");
                    formContext.data.save();
                    //onLoadDocFrame();
                }
            };
            xmlhttp.send(buildCreateReq(formContext));
        }
        else {
            onLoadDocFrame(formContext);
        }
    } catch (err) {
        alert("createDocFolder : \n" + err.message);
    }
}

//Function to build request
function buildCreateReq(formContext) {
    try {
        var OppDocLocation = null;
        var Opp = formContext.getAttribute("opportunityid").getValue();
        var OppId = null;
        //if have value
        if (Opp != null) {
            OppId = Opp[0].id;//get the guid of the record

            RetreiveMultipleRecords("sharepointdocumentlocations", "_regardingobjectid_value eq " + OppId, null,
                function (result) {
                    result = JSON.parse(result)
                    if (result != null && result.value != null) {
                        if (result.value.length == 1) {
                            var record = result.value[0];
                            OppDocLocation = record["sharepointdocumentlocationid"];
                        }
                    }
                }, function (error) { OppDocLocation = null; }, false);
        }

        var recordId = formContext.data.entity.getId().replace("{", "").replace("}", "");
        if (OppDocLocation != null) {
            OppDocLocation = OppDocLocation.replace("{", "").replace("}", "");
        }
        //Moran
        var SoapRequest = "<s:Envelope xmlns:s=\"http://schemas.xmlsoap.org/soap/envelope/\">" +
            " <s:Header>" +
            " <SdkClientVersion xmlns=\"http://schemas.microsoft.com/xrm/2011/Contracts\"></SdkClientVersion>" +
            " </s:Header>" +
            " <s:Body>" +
            "<Execute xmlns=\"http://schemas.microsoft.com/xrm/2011/Contracts/Services\" xmlns:i=\"http://www.w3.org/2001/XMLSchema-instance\">" +
            " <request i:type=\"b:CreateFolderRequest\" xmlns:a=\"http://schemas.microsoft.com/xrm/2011/Contracts\" xmlns:b=\"http://schemas.microsoft.com/crm/2011/Contracts\">" +
            " <a:Parameters xmlns:b=\"http://schemas.datacontract.org/2004/07/System.Collections.Generic\">" +
            " <a:KeyValuePairOfstringanyType>" +
            " <b:key>FolderName</b:key>" +
            " <b:value i:type=\"c:string\" xmlns:c=\"http://www.w3.org/2001/XMLSchema\">" + formContext.data.entity.getPrimaryAttributeValue().trim() + "</b:value>" +
            " </a:KeyValuePairOfstringanyType>" +
            " <a:KeyValuePairOfstringanyType>" +
            " <b:key>RegardingObjectType</b:key>" +
            " <b:value i:type=\"c:int\" xmlns:c=\"http://www.w3.org/2001/XMLSchema\">salesorder</b:value>" +
            " </a:KeyValuePairOfstringanyType>" +
            " <a:KeyValuePairOfstringanyType>" +
            " <b:key>RegardingObjectId</b:key>" +
            " <b:value i:type=\"c:string\" xmlns:c=\"http://www.w3.org/2001/XMLSchema\">" + recordId + "</b:value>" +
            " </a:KeyValuePairOfstringanyType>" +
            " <a:KeyValuePairOfstringanyType>" +
            " <b:key>DocumentType</b:key>" +
            " <b:value i:nil=\"true\" />" +
            " </a:KeyValuePairOfstringanyType>" +
            " <a:KeyValuePairOfstringanyType>" +
            " <b:key>ParentLocationId</b:key>";
        if (OppDocLocation != null) {
            SoapRequest = SoapRequest + " <b:value i:type=\"c:string\" xmlns:c=\"http://www.w3.org/2001/XMLSchema\">" + OppDocLocation + "</b:value>";
        }
        else {
            SoapRequest = SoapRequest + " <b:value i:nil=\"true\" />";
        }

        SoapRequest = SoapRequest +
            " </a:KeyValuePairOfstringanyType>" +
            " </a:Parameters>" +
            " <a:RequestId i:nil=\"true\" />" +
            " <a:RequestName>CreateFolder</a:RequestName>" +
            " </request>" +
            "</Execute>" +
            "</s:Body>" +
            "</s:Envelope>";

        return SoapRequest;

    } catch (err) {
        alert("buildCreateReq : \n" + err.message);
    }
}

function getLocation(formContext) {
    try {
        if (getDocsLocation(formContext) == 1) {

            var recordId = formContext.data.entity.getId();
            var url = getLocationOrUrl(recordId, formContext);
        }

    } catch (err) {
        alert("getLocation : \n" + err.message);
    }
}

function getDocsLocation(formContext) {
    try {
        if (formContext == null) {
            formContext = Xrm.Page;
        }
        var id = formContext.data.entity.getId();
        var docArrayExists = 0;
        if (formContext.ui.getFormType() == 1) {
            return -1;
        }

        var OrderSPLocation = GetValueFromSystemSettings("OrderSPLocation");
        //alert(OrderSPLocation);


        RetreiveMultipleRecords("sharepointdocumentlocations", "_regardingobjectid_value eq " + id, null,
            function (results) {
                results = JSON.parse(results)
                if (results != null && results.value != null) {
                    if (results.value.length > 1) {
                        for (var i = 0; i < results.length; i++) {
                            if (results.value[i].ParentSiteOrLocation != null) {
                                var ParentSite = results.value[i]["_parentsiteorlocation_value"];
                                ParentSite = ParentSite.replace("{", "").replace("}", "").toLowerCase();
                                var ParentLocationRecord = null;

                                RetreiveRecord("sharepointdocumentlocations", ParentSite, null,
                                    function (result) {

                                        if (result != null) {
                                            result = JSON.parse(result);
                                            ParentLocationRecord = result;
                                        }
                                    }, function (error) { }, false);

                                if (ParentLocationRecord["relativeurl"] != OrderSPLocation) {
                                    DeleteSPRecord(results.value[i]["sharepointdocumentlocationid"]);
                                }
                                else {
                                    //if have location record under order
                                    docArrayExists = 1;//mark as exist
                                }
                            }
                        }
                    }
                }
            }, function (error) { docArrayExists = 0; }, false);

        return docArrayExists;

    } catch (err) {
        alert("getDocsLocation : \n" + err.message);
    }
}

function DeleteSPRecord(recordId) {
    try {
        var globalContext = Xrm.Utility.getGlobalContext();

        var req = new XMLHttpRequest();
        req.open("DELETE", globalContext.getClientUrl() + "/api/data/v9.2/sharepointdocumentlocations(" + recordId + ")", false);
        req.setRequestHeader("Accept", "application/json");
        req.setRequestHeader("Content-Type", "application/json; charset=utf-8");
        req.setRequestHeader("OData-MaxVersion", "4.0");
        req.setRequestHeader("OData-Version", "4.0");
        req.onreadystatechange = function () {
            if (this.readyState === 4) {
                req.onreadystatechange = null;
                if (this.status === 204 || this.status === 1223) {
                    //Success - No Return Data - Do Something
                } else {
                    //Xrm.Utility.alertDialog(this.statusText);
                }
            }
        };
        req.send();

    } catch (err) {
        alert("DeleteSPRecord : \n" + err.message);
    }
}

function GetValueFromSystemSettings(value) {
    try {
        var ConfigValue = null;

        //    RetreiveMultipleRecords("bnt_systemsettingses", "$filter=bnt_name eq '" + value + "'", null,
        RetreiveMultipleRecords("bnt_systemsettingses", "bnt_name eq '" + value + "'", null,
            function (result) {
                result = JSON.parse(result)
                if (result.value.length > 0) {
                    var bnt_value = result.value[0]["bnt_value"];
                    ConfigValue = bnt_value;
                }
            }, function (error) { ConfigValue = null; }, false);
        return ConfigValue;

    } catch (err) {
        alert("GetValueFromSystemSettings : \n" + err.message);
    }
}

function getLocationOrUrl(formContext) {
    try {
        if (formContext == null) {
            formContext = Xrm.Page;
        }
        var globalContext = Xrm.Utility.getGlobalContext();
        var recordId = formContext.data.entity.getId();
        if (recordId == null) {
            return;
        }

        var req = new XMLHttpRequest();
        req.open("GET", globalContext.getClientUrl() + "/api/data/v9.2/sharepointdocumentlocations?$select=absoluteurl,description,name,_parentsiteorlocation_value,relativeurl,sharepointdocumentlocationid,sitecollectionid&$filter=_regardingobjectid_value eq " + recordId, false);
        req.setRequestHeader("OData-MaxVersion", "4.0");
        req.setRequestHeader("OData-Version", "4.0");
        req.setRequestHeader("Accept", "application/json");
        req.setRequestHeader("Content-Type", "application/json; charset=utf-8");
        req.setRequestHeader("Prefer", "odata.include-annotations=\"*\"");
        req.onreadystatechange = function () {
            if (this.readyState === 4) {
                req.onreadystatechange = null;
                if (this.status === 200) {
                    var results = JSON.parse(this.response);
                    for (var i = 0; i < results.value.length; i++) {

                        debugger;

                        var absoluteurl = results.value[i]["absoluteurl"];
                        var description = results.value[i]["description"];
                        var name = results.value[i]["name"];
                        var _parentsiteorlocation_value = results.value[i]["_parentsiteorlocation_value"];
                        var _parentsiteorlocation_value_formatted = results.value[i]["_parentsiteorlocation_value@OData.Community.Display.V1.FormattedValue"];
                        var _parentsiteorlocation_value_lookuplogicalname = results.value[i]["_parentsiteorlocation_value@Microsoft.Dynamics.CRM.lookuplogicalname"];
                        var relativeurl = results.value[i]["relativeurl"];
                        var sharepointdocumentlocationid = results.value[i]["sharepointdocumentlocationid"];
                        var sitecollectionid = results.value[i]["sitecollectionid"];
                    }
                } else {
                    var isOppConnected = formContext.getAttribute("bnt_opportunityconnected").getValue();
                    alert(isOppConnected);
                    if (isOppConnected != null) {

                        alert("ישנו עומס זמני נא לסגור הזמנה ולפתוח שוב");
                    }
                }
            };
            req.send();
        }

    } catch (err) {
        alert("getLocationOrUrl : \n" + err.message);
    }
}

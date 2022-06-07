/* Outlook specific API library */
/* Version: 15.0.4927.1000 */
/* Update: 3 */
/*
	Copyright (c) Microsoft Corporation.  All rights reserved.
*/

/*
	Your use of this file is governed by the Microsoft Services Agreement http://go.microsoft.com/fwlink/?LinkId=266419.
*/

Type.registerNamespace("Microsoft.Office.WebExtension.MailboxEnums");
Microsoft.Office.WebExtension.MailboxEnums.EntityType={
	MeetingSuggestion: "meetingSuggestion",
	TaskSuggestion: "taskSuggestion",
	Address: "address",
	EmailAddress: "emailAddress",
	Url: "url",
	PhoneNumber: "phoneNumber",
	Contact: "contact",
	FlightReservations: "flightReservations",
	ParcelDeliveries: "parcelDeliveries"
};
Microsoft.Office.WebExtension.MailboxEnums.ItemType={
	Message: "message",
	Appointment: "appointment"
};
Microsoft.Office.WebExtension.MailboxEnums.ResponseType={
	None: "none",
	Organizer: "organizer",
	Tentative: "tentative",
	Accepted: "accepted",
	Declined: "declined"
};
Microsoft.Office.WebExtension.MailboxEnums.RecipientType={
	Other: "other",
	DistributionList: "distributionList",
	User: "user",
	ExternalUser: "externalUser"
};
Microsoft.Office.WebExtension.MailboxEnums.AttachmentType={
	File: "file",
	Item: "item",
	Cloud: "cloud"
};
Microsoft.Office.WebExtension.MailboxEnums.BodyType={
	Text: "text",
	Html: "html"
};
Microsoft.Office.WebExtension.MailboxEnums.ItemNotificationMessageType={
	ProgressIndicator: "progressIndicator",
	InformationalMessage: "informationalMessage",
	ErrorMessage: "errorMessage"
};
Microsoft.Office.WebExtension.CoercionType={
	Text: "text",
	Html: "html"
};
Microsoft.Office.WebExtension.MailboxEnums.UserProfileType={
	Office365: "office365",
	OutlookCom: "outlookCom",
	Enterprise: "enterprise"
};
Microsoft.Office.WebExtension.MailboxEnums.RestVersion={
	v1_0: "v1.0",
	v2_0: "v2.0",
	Beta: "beta"
};
Type.registerNamespace("OSF.DDA");
var OSF=window.OSF || {};
OSF.DDA=OSF.DDA || {};
window["OSF"]["DDA"]["OutlookAppOm"]=OSF.DDA.OutlookAppOm=function(officeAppContext, targetWindow, appReadyCallback)
{
	this.$$d__callAppReadyCallback$p$0=Function.createDelegate(this,this._callAppReadyCallback$p$0);
	this.$$d_displayContactCardAsync=Function.createDelegate(this,this.displayContactCardAsync);
	this.$$d_displayNewMessageFormApi=Function.createDelegate(this,this.displayNewMessageFormApi);
	this.$$d__displayNewAppointmentFormApi$p$0=Function.createDelegate(this,this._displayNewAppointmentFormApi$p$0);
	this.$$d_windowOpenOverrideHandler=Function.createDelegate(this,this.windowOpenOverrideHandler);
	this.$$d__getEwsUrl$p$0=Function.createDelegate(this,this._getEwsUrl$p$0);
	this.$$d__getDiagnostics$p$0=Function.createDelegate(this,this._getDiagnostics$p$0);
	this.$$d__getUserProfile$p$0=Function.createDelegate(this,this._getUserProfile$p$0);
	this.$$d__getItem$p$0=Function.createDelegate(this,this._getItem$p$0);
	this.$$d__getInitialDataResponseHandler$p$0=Function.createDelegate(this,this._getInitialDataResponseHandler$p$0);
	window["OSF"]["DDA"]["OutlookAppOm"]._instance$p=this;
	this._officeAppContext$p$0=officeAppContext;
	this._appReadyCallback$p$0=appReadyCallback;
	var $$t_4=this;
	var stringLoadedCallback=function()
		{
			if(appReadyCallback)
				if(!officeAppContext["get_isDialog"]())
					$$t_4.invokeHostMethod(1,null,$$t_4.$$d__getInitialDataResponseHandler$p$0);
				else
					window.setTimeout($$t_4.$$d__callAppReadyCallback$p$0,0)	
		};
	if(this._areStringsLoaded$p$0())
		stringLoadedCallback();
	else
		this._loadLocalizedScript$p$0(stringLoadedCallback)
};
window["OSF"]["DDA"]["OutlookAppOm"]._throwOnPropertyAccessForRestrictedPermission$i=function(currentPermissionLevel)
{
	if(!currentPermissionLevel)
		throw Error.create(window["_u"]["ExtensibilityStrings"]["l_ElevatedPermissionNeeded_Text"]);
};
window["OSF"]["DDA"]["OutlookAppOm"]._throwOnOutOfRange$i=function(value, minValue, maxValue, argumentName)
{
	if(value < minValue || value > maxValue)
		throw Error.argumentOutOfRange(argumentName);
};
window["OSF"]["DDA"]["OutlookAppOm"]._getHtmlBody$p=function(data)
{
	var htmlBody="";
	if("htmlBody" in data)
	{
		window["OSF"]["DDA"]["OutlookAppOm"]._throwOnInvalidHtmlBody$p(data["htmlBody"]);
		htmlBody=data["htmlBody"]
	}
	return htmlBody
};
window["OSF"]["DDA"]["OutlookAppOm"]._getAttachments$p=function(data)
{
	var attachments=[];
	if("attachments" in data)
	{
		attachments=data["attachments"];
		window["OSF"]["DDA"]["OutlookAppOm"]._throwOnInvalidAttachmentsArray$p(attachments)
	}
	return attachments
};
window["OSF"]["DDA"]["OutlookAppOm"]._getOptionsAndCallback$p=function(data)
{
	var args=[];
	if("options" in data)
		args[0]=data["options"];
	if("callback" in data)
		args[args.length]=data["callback"];
	return args
};
window["OSF"]["DDA"]["OutlookAppOm"]._createAttachmentsDataForHost$p=function(attachments)
{
	var attachmentsData=new Array(0);
	if(Array.isInstanceOfType(attachments))
		for(var i=0; i < attachments["length"]; i++)
			if(Object.isInstanceOfType(attachments[i]))
			{
				var attachment=attachments[i];
				window["OSF"]["DDA"]["OutlookAppOm"]._throwOnInvalidAttachment$p(attachment);
				attachmentsData[i]=window["OSF"]["DDA"]["OutlookAppOm"]._createAttachmentData$p(attachment)
			}
			else
				throw Error.argument("attachments");
	return attachmentsData
};
window["OSF"]["DDA"]["OutlookAppOm"]._throwOnInvalidHtmlBody$p=function(htmlBody)
{
	if(!String.isInstanceOfType(htmlBody))
		throw Error.argument("htmlBody");
	if($h.ScriptHelpers.isNullOrUndefined(htmlBody))
		throw Error.argument("htmlBody");
	window["OSF"]["DDA"]["OutlookAppOm"]._throwOnOutOfRange$i(htmlBody["length"],0,window["OSF"]["DDA"]["OutlookAppOm"].maxBodyLength,"htmlBody")
};
window["OSF"]["DDA"]["OutlookAppOm"]._throwOnInvalidAttachmentsArray$p=function(attachments)
{
	if(!Array.isInstanceOfType(attachments))
		throw Error.argument("attachments");
};
window["OSF"]["DDA"]["OutlookAppOm"]._throwOnInvalidAttachment$p=function(attachment)
{
	if(!Object.isInstanceOfType(attachment))
		throw Error.argument("attachments");
	if(!("type" in attachment) || !("name" in attachment))
		throw Error.argument("attachments");
	if(!("url" in attachment || "itemId" in attachment))
		throw Error.argument("attachments");
};
window["OSF"]["DDA"]["OutlookAppOm"]._createAttachmentData$p=function(attachment)
{
	var attachmentData=null;
	if(attachment["type"]==="file")
	{
		var url=attachment["url"];
		var name=attachment["name"];
		window["OSF"]["DDA"]["OutlookAppOm"]._throwOnInvalidAttachmentUrlOrName$p(url,name);
		attachmentData=window["OSF"]["DDA"]["OutlookAppOm"]._createFileAttachmentData$p(url,name)
	}
	else if(attachment["type"]==="item")
	{
		var itemId=window["OSF"]["DDA"]["OutlookAppOm"].getItemIdBasedOnHost(attachment["itemId"]);
		var name=attachment["name"];
		window["OSF"]["DDA"]["OutlookAppOm"]._throwOnInvalidAttachmentItemIdOrName$p(itemId,name);
		attachmentData=window["OSF"]["DDA"]["OutlookAppOm"]._createItemAttachmentData$p(itemId,name)
	}
	else
		throw Error.argument("attachments");
	return attachmentData
};
window["OSF"]["DDA"]["OutlookAppOm"]._createFileAttachmentData$p=function(url, name)
{
	return["file",name,url]
};
window["OSF"]["DDA"]["OutlookAppOm"]._createItemAttachmentData$p=function(itemId, name)
{
	return["item",name,itemId]
};
window["OSF"]["DDA"]["OutlookAppOm"]._throwOnInvalidAttachmentUrlOrName$p=function(url, name)
{
	if(!String.isInstanceOfType(url) || !String.isInstanceOfType(name))
		throw Error.argument("attachments");
	if(url["length"] > 2048)
		throw Error.argumentOutOfRange("attachments",url["length"],window["_u"]["ExtensibilityStrings"]["l_AttachmentUrlTooLong_Text"]);
	window["OSF"]["DDA"]["OutlookAppOm"]._throwOnInvalidAttachmentName$p(name)
};
window["OSF"]["DDA"]["OutlookAppOm"]._throwOnInvalidAttachmentItemIdOrName$p=function(itemId, name)
{
	if(!String.isInstanceOfType(itemId) || !String.isInstanceOfType(name))
		throw Error.argument("attachments");
	if(itemId["length"] > 200)
		throw Error.argumentOutOfRange("attachments",itemId["length"],window["_u"]["ExtensibilityStrings"]["l_AttachmentItemIdTooLong_Text"]);
	window["OSF"]["DDA"]["OutlookAppOm"]._throwOnInvalidAttachmentName$p(name)
};
window["OSF"]["DDA"]["OutlookAppOm"]._throwOnInvalidAttachmentName$p=function(name)
{
	if(name["length"] > 255)
		throw Error.argumentOutOfRange("attachments",name["length"],window["_u"]["ExtensibilityStrings"]["l_AttachmentNameTooLong_Text"]);
};
window["OSF"]["DDA"]["OutlookAppOm"]._throwOnInvalidRestVersion$p=function(restVersion)
{
	if(!restVersion)
		throw Error.argumentNull("restVersion");
	if(restVersion !==window["Microsoft"]["Office"]["WebExtension"]["MailboxEnums"]["RestVersion"]["v1_0"] && restVersion !==window["Microsoft"]["Office"]["WebExtension"]["MailboxEnums"]["RestVersion"]["v2_0"] && restVersion !==window["Microsoft"]["Office"]["WebExtension"]["MailboxEnums"]["RestVersion"]["Beta"])
		throw Error.argument("restVersion");
};
window["OSF"]["DDA"]["OutlookAppOm"].getItemIdBasedOnHost=function(itemId)
{
	if(window["OSF"]["DDA"]["OutlookAppOm"]._instance$p._initialData$p$0 && window["OSF"]["DDA"]["OutlookAppOm"]._instance$p._initialData$p$0.get__isRestIdSupported$i$0())
		return window["OSF"]["DDA"]["OutlookAppOm"]._instance$p["convertToRestId"](itemId,window["Microsoft"]["Office"]["WebExtension"]["MailboxEnums"]["RestVersion"]["v1_0"]);
	return window["OSF"]["DDA"]["OutlookAppOm"]._instance$p["convertToEwsId"](itemId,window["Microsoft"]["Office"]["WebExtension"]["MailboxEnums"]["RestVersion"]["v1_0"])
};
window["OSF"]["DDA"]["OutlookAppOm"]._throwOnArgumentType$p=function(value, expectedType, argumentName)
{
	if(Object["getType"](value) !==expectedType)
		throw Error.argumentType(argumentName);
};
window["OSF"]["DDA"]["OutlookAppOm"]._validateOptionalStringParameter$p=function(value, minLength, maxLength, name)
{
	if($h.ScriptHelpers.isNullOrUndefined(value))
		return;
	window["OSF"]["DDA"]["OutlookAppOm"]._throwOnArgumentType$p(value,String,name);
	var stringValue=value;
	window["OSF"]["DDA"]["OutlookAppOm"]._throwOnOutOfRange$i(stringValue["length"],minLength,maxLength,name)
};
window["OSF"]["DDA"]["OutlookAppOm"]._convertToOutlookParameters$p=function(dispid, data)
{
	var executeParameters=null;
	switch(dispid)
	{
		case 1:
		case 2:
		case 12:
		case 3:
		case 14:
		case 18:
		case 26:
		case 32:
		case 41:
		case 34:
			break;
		case 4:
			var jsonProperty=window["JSON"]["stringify"](data["customProperties"]);
			executeParameters=[jsonProperty];
			break;
		case 5:
			executeParameters=[data["body"]];
			break;
		case 8:
		case 9:
			executeParameters=[data["itemId"]];
			break;
		case 7:
			executeParameters=[window["OSF"]["DDA"]["OutlookAppOm"]._convertRecipientArrayParameterForOutlookForDisplayApi$p(data["requiredAttendees"]),window["OSF"]["DDA"]["OutlookAppOm"]._convertRecipientArrayParameterForOutlookForDisplayApi$p(data["optionalAttendees"]),data["start"],data["end"],data["location"],window["OSF"]["DDA"]["OutlookAppOm"]._convertRecipientArrayParameterForOutlookForDisplayApi$p(data["resources"]),data["subject"],data["body"]];
			break;
		case 44:
			executeParameters=[window["OSF"]["DDA"]["OutlookAppOm"]._convertRecipientArrayParameterForOutlookForDisplayApi$p(data["toRecipients"]),window["OSF"]["DDA"]["OutlookAppOm"]._convertRecipientArrayParameterForOutlookForDisplayApi$p(data["ccRecipients"]),window["OSF"]["DDA"]["OutlookAppOm"]._convertRecipientArrayParameterForOutlookForDisplayApi$p(data["bccRecipients"]),data["subject"],data["htmlBody"],data["attachments"]];
			break;
		case 43:
			executeParameters=[data["ewsIdOrEmail"]];
			break;
		case 40:
			executeParameters=[data["extensionId"],data["consentState"]];
			break;
		case 11:
		case 10:
			executeParameters=[data["htmlBody"]];
			break;
		case 31:
		case 30:
			executeParameters=[data["htmlBody"],data["attachments"]];
			break;
		case 23:
		case 13:
		case 38:
		case 29:
			executeParameters=[data["data"],data["coercionType"]];
			break;
		case 37:
		case 28:
			executeParameters=[data["coercionType"]];
			break;
		case 17:
			executeParameters=[data["subject"]];
			break;
		case 15:
			executeParameters=[data["recipientField"]];
			break;
		case 22:
		case 21:
			executeParameters=[data["recipientField"],window["OSF"]["DDA"]["OutlookAppOm"]._convertComposeEmailDictionaryParameterForSetApi$p(data["recipientArray"])];
			break;
		case 19:
			executeParameters=[data["itemId"],data["name"]];
			break;
		case 16:
			executeParameters=[data["uri"],data["name"]];
			break;
		case 20:
			executeParameters=[data["attachmentIndex"]];
			break;
		case 25:
			executeParameters=[data["TimeProperty"],data["time"]];
			break;
		case 24:
			executeParameters=[data["TimeProperty"]];
			break;
		case 27:
			executeParameters=[data["location"]];
			break;
		case 33:
		case 35:
			executeParameters=[data["key"],data["type"],data["persistent"],data["message"],data["icon"]];
			break;
		case 36:
			executeParameters=[data["key"]];
			break;
		default:
			Sys.Debug.fail("Unexpected method dispid");
			break
	}
	return executeParameters
};
window["OSF"]["DDA"]["OutlookAppOm"]._convertRecipientArrayParameterForOutlookForDisplayApi$p=function(array)
{
	return array ? array["join"](";") : null
};
window["OSF"]["DDA"]["OutlookAppOm"]._convertComposeEmailDictionaryParameterForSetApi$p=function(recipients)
{
	if(!recipients)
		return null;
	var results=new Array(recipients.length);
	for(var i=0; i < recipients.length; i++)
		results[i]=[recipients[i]["address"],recipients[i]["name"]];
	return results
};
window["OSF"]["DDA"]["OutlookAppOm"]._validateAndNormalizeRecipientEmails$p=function(emailset, name)
{
	if($h.ScriptHelpers.isNullOrUndefined(emailset))
		return null;
	window["OSF"]["DDA"]["OutlookAppOm"]._throwOnArgumentType$p(emailset,Array,name);
	var originalAttendees=emailset;
	var updatedAttendees=null;
	var normalizationNeeded=false;
	window["OSF"]["DDA"]["OutlookAppOm"]._throwOnOutOfRange$i(originalAttendees["length"],0,window["OSF"]["DDA"]["OutlookAppOm"]._maxRecipients$p,String.format("{0}.length",name));
	for(var i=0; i < originalAttendees["length"]; i++)
		if($h.EmailAddressDetails.isInstanceOfType(originalAttendees[i]))
		{
			normalizationNeeded=true;
			break
		}
	if(normalizationNeeded)
		updatedAttendees=[];
	for(var i=0; i < originalAttendees["length"]; i++)
		if(normalizationNeeded)
		{
			updatedAttendees[i]=$h.EmailAddressDetails.isInstanceOfType(originalAttendees[i]) ? originalAttendees[i]["emailAddress"] : originalAttendees[i];
			window["OSF"]["DDA"]["OutlookAppOm"]._throwOnArgumentType$p(updatedAttendees[i],String,String.format("{0}[{1}]",name,i))
		}
		else
			window["OSF"]["DDA"]["OutlookAppOm"]._throwOnArgumentType$p(originalAttendees[i],String,String.format("{0}[{1}]",name,i));
	return updatedAttendees
};
OSF.DDA.OutlookAppOm.prototype={
	_initialData$p$0: null,
	_item$p$0: null,
	_userProfile$p$0: null,
	_diagnostics$p$0: null,
	_officeAppContext$p$0: null,
	_appReadyCallback$p$0: null,
	_clientEndPoint$p$0: null,
	get_clientEndPoint: function()
	{
		if(!this._clientEndPoint$p$0)
			this._clientEndPoint$p$0=window["OSF"]["_OfficeAppFactory"]["getClientEndPoint"]();
		return this._clientEndPoint$p$0
	},
	set_clientEndPoint: function(value)
	{
		this._clientEndPoint$p$0=value;
		return value
	},
	get_initialData: function()
	{
		return this._initialData$p$0
	},
	get__appName$i$0: function()
	{
		return this._officeAppContext$p$0["get_appName"]()
	},
	initialize: function(initialData)
	{
		var ItemTypeKey="itemType";
		this._initialData$p$0=new $h.InitialData(initialData);
		if(1===initialData[ItemTypeKey])
			this._item$p$0=new $h.Mess
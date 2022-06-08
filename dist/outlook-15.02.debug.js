/* Outlook specific API library */
/* Version: 15.0.4726.1000 */
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
	Contact: "contact"
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
	Item: "item"
};
Microsoft.Office.WebExtension.MailboxEnums.BodyType={
	Text: "text",
	Html: "html"
};
Microsoft.Office.WebExtension.CoercionType={
	Text: "text",
	Html: "html"
};
Type.registerNamespace("OSF.DDA");
OSF.DDA.OutlookAppOm=function(officeAppContext, targetWindow, appReadyCallback)
{
	this.$$d__callAppReadyCallback$p$0=Function.createDelegate(this,this._callAppReadyCallback$p$0);
	this.$$d__displayNewAppointmentFormApi$p$0=Function.createDelegate(this,this._displayNewAppointmentFormApi$p$0);
	this.$$d_windowOpenOverrideHandler=Function.createDelegate(this,this.windowOpenOverrideHandler);
	this.$$d__getEwsUrl$p$0=Function.createDelegate(this,this._getEwsUrl$p$0);
	this.$$d__getDiagnostics$p$0=Function.createDelegate(this,this._getDiagnostics$p$0);
	this.$$d__getUserProfile$p$0=Function.createDelegate(this,this._getUserProfile$p$0);
	this.$$d__getItem$p$0=Function.createDelegate(this,this._getItem$p$0);
	this.$$d__getInitialDataResponseHandler$p$0=Function.createDelegate(this,this._getInitialDataResponseHandler$p$0);
	OSF.DDA.OutlookAppOm._instance$p=this;
	this._officeAppContext$p$0=officeAppContext;
	this._appReadyCallback$p$0=appReadyCallback;
	var $$t_4=this;
	var stringLoadedCallback=function()
		{
			if(appReadyCallback)
				$$t_4._invokeHostMethod$i$0(1,"GetInitialData",null,$$t_4.$$d__getInitialDataResponseHandler$p$0)
		};
	if(this._areStringsLoaded$p$0())
		stringLoadedCallback();
	else
		this._loadLocalizedScript$p$0(stringLoadedCallback)
};
OSF.DDA.OutlookAppOm._throwOnPropertyAccessForRestrictedPermission$i=function(currentPermissionLevel)
{
	if(!currentPermissionLevel)
		throw Error.create(_u.ExtensibilityStrings.l_ElevatedPermissionNeeded_Text);
};
OSF.DDA.OutlookAppOm._throwOnOutOfRange$i=function(value, minValue, maxValue, argumentName)
{
	if(value < minValue || value > maxValue)
		throw Error.argumentOutOfRange(argumentName);
};
OSF.DDA.OutlookAppOm._getHtmlBody$p=function(data)
{
	var htmlBody="";
	if("htmlBody" in data)
	{
		OSF.DDA.OutlookAppOm._throwOnInvalidHtmlBody$p(data["htmlBody"]);
		htmlBody=data["htmlBody"]
	}
	return htmlBody
};
OSF.DDA.OutlookAppOm._getAttachments$p=function(data)
{
	var attachments=[];
	if("attachments" in data)
	{
		attachments=data["attachments"];
		OSF.DDA.OutlookAppOm._throwOnInvalidAttachmentsArray$p(attachments)
	}
	return attachments
};
OSF.DDA.OutlookAppOm._getOptionsAndCallback$p=function(data)
{
	var args=[];
	if("options" in data)
		args[0]=data["options"];
	if("callback" in data)
		args[args.length]=data["callback"];
	return args
};
OSF.DDA.OutlookAppOm._createAttachmentsDataForHost$p=function(attachments)
{
	var attachmentsData=new Array(0);
	if(Array.isInstanceOfType(attachments))
		for(var i=0; i < attachments.length; i++)
			if(Object.isInstanceOfType(attachments[i]))
			{
				var attachment=attachments[i];
				OSF.DDA.OutlookAppOm._throwOnInvalidAttachment$p(attachment);
				attachmentsData[i]=OSF.DDA.OutlookAppOm._createAttachmentData$p(attachment)
			}
			else
				throw Error.argument("attachments");
	return attachmentsData
};
OSF.DDA.OutlookAppOm._throwOnInvalidHtmlBody$p=function(htmlBody)
{
	if(!String.isInstanceOfType(htmlBody))
		throw Error.argument("htmlBody");
	if($h.ScriptHelpers.isNullOrUndefined(htmlBody))
		throw Error.argument("htmlBody");
	OSF.DDA.OutlookAppOm._throwOnOutOfRange$i(htmlBody.length,0,OSF.DDA.OutlookAppOm.maxBodyLength,"htmlBody")
};
OSF.DDA.OutlookAppOm._throwOnInvalidAttachmentsArray$p=function(attachments)
{
	if(!Array.isInstanceOfType(attachments))
		throw Error.argument("attachments");
};
OSF.DDA.OutlookAppOm._throwOnInvalidAttachment$p=function(attachment)
{
	if(!Object.isInstanceOfType(attachment))
		throw Error.argument("attachments");
	if(!("type" in attachment) || !("name" in attachment))
		throw Error.argument("attachments");
	if(!("url" in attachment || "itemId" in attachment))
		throw Error.argument("attachments");
};
OSF.DDA.OutlookAppOm._createAttachmentData$p=function(attachment)
{
	var attachmentData=null;
	if(attachment["type"]==="file")
	{
		var url=attachment["url"];
		var name=attachment["name"];
		OSF.DDA.OutlookAppOm._throwOnInvalidAttachmentUrlOrName$p(url,name);
		attachmentData=OSF.DDA.OutlookAppOm._createFileAttachmentData$p(url,name)
	}
	else if(attachment["type"]==="item")
	{
		var itemId=attachment["itemId"];
		var name=attachment["name"];
		OSF.DDA.OutlookAppOm._throwOnInvalidAttachmentItemIdOrName$p(itemId,name);
		attachmentData=OSF.DDA.OutlookAppOm._createItemAttachmentData$p(itemId,name)
	}
	else
		throw Error.argument("attachments");
	return attachmentData
};
OSF.DDA.OutlookAppOm._createFileAttachmentData$p=function(url, name)
{
	return["file",name,url]
};
OSF.DDA.OutlookAppOm._createItemAttachmentData$p=function(itemId, name)
{
	return["item",name,itemId]
};
OSF.DDA.OutlookAppOm._throwOnInvalidAttachmentUrlOrName$p=function(url, name)
{
	if(!String.isInstanceOfType(url) || !String.isInstanceOfType(name))
		throw Error.argument("attachments");
	if(url.length > 2048)
		throw Error.argumentOutOfRange("attachments",url.length,_u.ExtensibilityStrings.l_AttachmentUrlTooLong_Text);
	OSF.DDA.OutlookAppOm._throwOnInvalidAttachmentName$p(name)
};
OSF.DDA.OutlookAppOm._throwOnInvalidAttachmentItemIdOrName$p=function(itemId, name)
{
	if(!String.isInstanceOfType(itemId) || !String.isInstanceOfType(name))
		throw Error.argument("attachments");
	if(itemId.length > 200)
		throw Error.argumentOutOfRange("attachments",itemId.length,_u.ExtensibilityStrings.l_AttachmentItemIdTooLong_Text);
	OSF.DDA.OutlookAppOm._throwOnInvalidAttachmentName$p(name)
};
OSF.DDA.OutlookAppOm._throwOnInvalidAttachmentName$p=function(name)
{
	if(name.length > 255)
		throw Error.argumentOutOfRange("attachments",name.length,_u.ExtensibilityStrings.l_AttachmentNameTooLong_Text);
};
OSF.DDA.OutlookAppOm._throwOnArgumentType$p=function(value, expectedType, argumentName)
{
	if(Object.getType(value) !==expectedType)
		throw Error.argumentType(argumentName);
};
OSF.DDA.OutlookAppOm._validateOptionalStringParameter$p=function(value, minLength, maxLength, name)
{
	if($h.ScriptHelpers.isNullOrUndefined(value))
		return;
	OSF.DDA.OutlookAppOm._throwOnArgumentType$p(value,String,name);
	var stringValue=value;
	OSF.DDA.OutlookAppOm._throwOnOutOfRange$i(stringValue.length,minLength,maxLength,name)
};
OSF.DDA.OutlookAppOm._convertToOutlookParameters$p=function(dispid, data)
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
			break;
		case 4:
			var jsonProperty=JSON.stringify(data["customProperties"]);
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
			executeParameters=[OSF.DDA.OutlookAppOm._convertRecipientArrayParameterForOutlookForDisplayApi$p(data["requiredAttendees"]),OSF.DDA.OutlookAppOm._convertRecipientArrayParameterForOutlookForDisplayApi$p(data["optionalAttendees"]),data["start"],data["end"],data["location"],OSF.DDA.OutlookAppOm._convertRecipientArrayParameterForOutlookForDisplayApi$p(data["resources"]),data["subject"],data["body"]];
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
		case 29:
			executeParameters=[data["data"],data["coercionType"]];
			break;
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
			executeParameters=[data["recipientField"],OSF.DDA.OutlookAppOm._convertComposeEmailDictionaryParameterForSetApi$p(data["recipientArray"])];
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
		default:
			Sys.Debug.fail("Unexpected method dispid");
			break
	}
	return executeParameters
};
OSF.DDA.OutlookAppOm._convertRecipientArrayParameterForOutlookForDisplayApi$p=function(array)
{
	return array ? array.join(";") : null
};
OSF.DDA.OutlookAppOm._convertComposeEmailDictionaryParameterForSetApi$p=function(recipients)
{
	if(!recipients)
		return null;
	var results=new Array(recipients.length);
	for(var i=0; i < recipients.length; i++)
		results[i]=[recipients[i]["address"],recipients[i]["name"]];
	return results
};
OSF.DDA.OutlookAppOm._validateAndNormalizeRecipientEmails$p=function(emailset, name)
{
	if($h.ScriptHelpers.isNullOrUndefined(emailset))
		return null;
	OSF.DDA.OutlookAppOm._throwOnArgumentType$p(emailset,Array,name);
	var originalAttendees=emailset;
	var updatedAttendees=null;
	var normalizationNeeded=false;
	OSF.DDA.OutlookAppOm._throwOnOutOfRange$i(originalAttendees.length,0,OSF.DDA.OutlookAppOm._maxRecipients$p,String.format("{0}.length",name));
	for(var i=0; i < originalAttendees.length; i++)
		if($h.EmailAddressDetails.isInstanceOfType(originalAttendees[i]))
		{
			normalizationNeeded=true;
			break
		}
	if(normalizationNeeded)
		updatedAttendees=[];
	for(var i=0; i < originalAttendees.length; i++)
		if(normalizationNeeded)
		{
			updatedAttendees[i]=$h.EmailAddressDetails.isInstanceOfType(originalAttendees[i]) ? originalAttendees[i].emailAddress : originalAttendees[i];
			OSF.DDA.OutlookAppOm._throwOnArgumentType$p(updatedAttendees[i],String,String.format("{0}[{1}]",name,i))
		}
		else
			OSF.DDA.OutlookAppOm._throwOnArgumentType$p(originalAttendees[i],String,String.format("{0}[{1}]",name,i));
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
			this._clientEndPoint$p$0=OSF._OfficeAppFactory.getClientEndPoint();
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
		return this._officeAppContext$p$0.get_appName()
	},
	initialize: function(initialData)
	{
		var ItemTypeKey="itemType";
		this._initialData$p$0=new $h.InitialData(initialData);
		if(1===initialData[ItemTypeKey])
			this._item$p$0=new $h.Message(this._initialData$p$0);
		else if(3===initialData[ItemTypeKey])
			this._item$p$0=new $h.MeetingRequest(this._initialData$p$0);
		else if(2===initialData[ItemTypeKey])
			this._item$p$0=new $h.Appointment(this._initialData$p$0);
		else if(4===initialData[ItemTypeKey])
			this._item$p$0=new $h.MessageCompose(this._initialData$p$0);
		else if(5===initialData[ItemTypeKey])
			this._item$p$0=new $h.AppointmentCompose(this._initialData$p$0);
		else
			Sys.Debug.trace("Unexpected item type was received from the host.");
		this._userProfile$p$0=new $h.UserProfile(this._initialData$p$0);
		this._diagnostics$p$0=new $h.Diagnostics(this._initialData$p$0,this._officeAppContext$p$0.get_appName());
		this._initializeMethods$p$0();
		$h.InitialData._defineReadOnlyProperty$i(this,"item",this.$$d__getItem$p$0);
		$h.InitialData._defineReadOnlyProperty$i(this,"userProfile",this.$$d__getUserProfile$p$0);
		$h.InitialData._defineReadOnlyProperty$i(this,"diagnostics",this.$$d__getDiagnostics$p$0);
		$h.InitialData._defineReadOnlyProperty$i(this,"ewsUrl",this.$$d__getEwsUrl$p$0);
		if(OSF.DDA.OutlookAppOm._instance$p.get__appName$i$0()===64)
			if(this._initialData$p$0.get__overrideWindowOpen$i$0())
				window.open=this.$$d_windowOpenOverrideHandler
	},
	windowOpenOverrideHandler: function(url, targetName, features, replace)
	{
		this._invokeHostMethod$i$0(0,"LaunchPalUrl",{launchUrl: url},null)
	},
	makeEwsRequestAsync: function(data)
	{
		var args=[];
		for(var $$pai_5=1; $$pai_5 < arguments.length;++$$pai_5)
			args[$$pai_5 - 1]=arguments[$$pai_5];
		if($h.ScriptHelpers.isNullOrUndefined(data))
			throw Error.argumentNull("data");
		if(data.length > OSF.DDA.OutlookAppOm._maxEwsRequestSize$p)
			throw Error.argument("data",_u.ExtensibilityStrings.l_EwsRequestOversized_Text);
		this._throwOnMethodCallForInsufficientPermission$i$0(3,"makeEwsRequestAsync");
		var parameters=$h.CommonParameters.parse(args,true,true);
		var ewsRequest=new $h.EwsRequest(parameters._asyncContext$p$0);
		var $$t_4=this;
		ewsRequest.onreadystatechange=function()
		{
			if(4===ewsRequest.get__requestState$i$1())
				parameters._callback$p$0(ewsRequest._asyncResult$p$0)
		};
		ewsRequest.send(data)
	},
	recordDataPoint: function(data)
	{
		if($h.ScriptHelpers.isNullOrUndefined(data))
			throw Error.argumentNull("data");
		this._invokeHostMethod$i$0(0,"RecordDataPoint",data,null)
	},
	recordTrace: function(data)
	{
		if($h.ScriptHelpers.isNullOrUndefined(data))
			throw Error.argumentNull("data");
		this._invokeHostMethod$i$0(0,"RecordTrace",data,null)
	},
	trackCtq: function(data)
	{
		if($h.ScriptHelpers.isNullOrUndefined(data))
			throw Error.argumentNull("data");
		this._invokeHostMethod$i$0(0,"TrackCtq",data,null)
	},
	convertToLocalClientTime: function(timeValue)
	{
		var date=new Date(timeValue.getTime());
		var offset=date.getTimezoneOffset() * -1;
		if(this._initialData$p$0 && this._initialData$p$0.get__timeZoneOffsets$i$0())
		{
			date.setUTCMinutes(date.getUTCMinutes() - offset);
			offset=this._findOffset$p$0(date);
			date.setUTCMinutes(date.getUTCMinutes()+offset)
		}
		var retValue=this._dateToDictionary$i$0(date);
		retValue["timezoneOffset"]=offset;
		return retValue
	},
	convertToUtcClientTime: function(input)
	{
		var retValue=this._dictionaryToDate$i$0(input);
		if(this._initialData$p$0 && this._initialData$p$0.get__timeZoneOffsets$i$0())
		{
			var offset=this._findOffset$p$0(retValue);
			retValue.setUTCMinutes(retValue.getUTCMinutes() - offset);
			offset=!input["timezoneOffset"] ? retValue.getTimezoneOffset() * -1 : input["timezoneOffset"];
			retValue.setUTCMinutes(retValue.getUTCMinutes()+offset)
		}
		return retValue
	},
	getUserIdentityTokenAsync: function()
	{
		var args=[];
		for(var $$pai
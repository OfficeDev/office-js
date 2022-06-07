/* Outlook specific API library */
/* Version: 15.0.4615.1000 */
/*
	Copyright (c) Microsoft Corporation.  All rights reserved.
*/

/*
	Your use of this file is governed by the Microsoft Services Agreement http://go.microsoft.com/fwlink/?LinkId=266419.
*/

;
Type.registerNamespace('Microsoft.Office.WebExtension.MailboxEnums');
Microsoft.Office.WebExtension.MailboxEnums.EntityType={
	MeetingSuggestion: "meetingSuggestion", TaskSuggestion: "taskSuggestion", Address: "address", EmailAddress: "emailAddress", Url: "url", PhoneNumber: "phoneNumber", Contact: "contact"
};
Microsoft.Office.WebExtension.MailboxEnums.ItemType={
	Message: 'message', Appointment: 'appointment'
};
Microsoft.Office.WebExtension.MailboxEnums.ResponseType={
	None: "none", Organizer: "organizer", Tentative: "tentative", Accepted: "accepted", Declined: "declined"
};
Microsoft.Office.WebExtension.MailboxEnums.RecipientType={
	Other: "other", DistributionList: "distributionList", User: "user", ExternalUser: "externalUser"
};
Microsoft.Office.WebExtension.MailboxEnums.AttachmentType={
	File: "file", Item: "item"
};
Microsoft.Office.WebExtension.MailboxEnums.BodyType={
	Text: "text", Html: "html"
};
Microsoft.Office.WebExtension.CoercionType={
	Text: "text", Html: "html"
};
;
Type.registerNamespace('OSF.DDA');
OSF.DDA.OutlookAppOm=function OSF_DDA_OutlookAppOm(officeAppContext, targetWindow, appReadyCallback)
{
	this.$$d__callAppReadyCallback$p$0=Function.createDelegate(this, this._callAppReadyCallback$p$0);
	this.$$d__displayNewAppointmentFormApi$p$0=Function.createDelegate(this, this._displayNewAppointmentFormApi$p$0);
	this.$$d_windowOpenOverrideHandler=Function.createDelegate(this, this.windowOpenOverrideHandler);
	this.$$d__getEwsUrl$p$0=Function.createDelegate(this, this._getEwsUrl$p$0);
	this.$$d__getDiagnostics$p$0=Function.createDelegate(this, this._getDiagnostics$p$0);
	this.$$d__getUserProfile$p$0=Function.createDelegate(this, this._getUserProfile$p$0);
	this.$$d__getItem$p$0=Function.createDelegate(this, this._getItem$p$0);
	this.$$d__getInitialDataResponseHandler$p$0=Function.createDelegate(this, this._getInitialDataResponseHandler$p$0);
	OSF.DDA.OutlookAppOm._instance$p=this;
	this._officeAppContext$p$0=officeAppContext;
	this._appReadyCallback$p$0=appReadyCallback;
	var $$t_4=this;
	var stringLoadedCallback=function()
		{
			if (appReadyCallback)
			{
				$$t_4._invokeHostMethod$i$0(1, 'GetInitialData', null, $$t_4.$$d__getInitialDataResponseHandler$p$0)
			}
		};
	if (this._areStringsLoaded$p$0())
	{
		stringLoadedCallback()
	}
	else
	{
		this._loadLocalizedScript$p$0(stringLoadedCallback)
	}
};
OSF.DDA.OutlookAppOm._throwOnPropertyAccessForRestrictedPermission$i=function OSF_DDA_OutlookAppOm$_throwOnPropertyAccessForRestrictedPermission$i(currentPermissionLevel)
{
	if (!currentPermissionLevel)
	{
		throw Error.create(_u.ExtensibilityStrings.l_ElevatedPermissionNeeded_Text);
	}
};
OSF.DDA.OutlookAppOm._throwOnOutOfRange$i=function OSF_DDA_OutlookAppOm$_throwOnOutOfRange$i(value, minValue, maxValue, argumentName)
{
	if (value < minValue || value > maxValue)
	{
		throw Error.argumentOutOfRange(argumentName);
	}
};
OSF.DDA.OutlookAppOm._throwOnArgumentType$p=function OSF_DDA_OutlookAppOm$_throwOnArgumentType$p(value, expectedType, argumentName)
{
	if (Object.getType(value) !==expectedType)
	{
		throw Error.argumentType(argumentName);
	}
};
OSF.DDA.OutlookAppOm._validateOptionalStringParameter$p=function OSF_DDA_OutlookAppOm$_validateOptionalStringParameter$p(value, minLength, maxLength, name)
{
	if ($h.ScriptHelpers.isNullOrUndefined(value))
	{
		return
	}
	OSF.DDA.OutlookAppOm._throwOnArgumentType$p(value, String, name);
	var stringValue=value;
	OSF.DDA.OutlookAppOm._throwOnOutOfRange$i(stringValue.length, minLength, maxLength, name)
};
OSF.DDA.OutlookAppOm._convertToOutlookParameters$p=function OSF_DDA_OutlookAppOm$_convertToOutlookParameters$p(dispid, data)
{
	var executeParameters=null;
	switch (dispid)
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
			var jsonProperty=JSON.stringify(data['customProperties']);
			executeParameters=[jsonProperty];
			break;
		case 5:
			executeParameters=[data['body']];
			break;
		case 8:
		case 9:
			executeParameters=[data['itemId']];
			break;
		case 7:
			executeParameters=[OSF.DDA.OutlookAppOm._convertRecipientArrayParameterForOutlookForDisplayApi$p(data['requiredAttendees']), OSF.DDA.OutlookAppOm._convertRecipientArrayParameterForOutlookForDisplayApi$p(data['optionalAttendees']), data['start'], data['end'], data['location'], OSF.DDA.OutlookAppOm._convertRecipientArrayParameterForOutlookForDisplayApi$p(data['resources']), data['subject'], data['body']];
			break;
		case 11:
		case 10:
			executeParameters=[data['htmlBody']];
			break;
		case 23:
		case 13:
			executeParameters=[data['data'], data['coercionType'] || null];
			break;
		case 17:
			executeParameters=[data['subject']];
			break;
		case 15:
			executeParameters=[data['recipientField']];
			break;
		case 22:
		case 21:
			executeParameters=[data['recipientField'], OSF.DDA.OutlookAppOm._convertComposeEmailDictionaryParameterForSetApi$p(data['recipientArray'])];
			break;
		case 19:
			executeParameters=[data['itemId'], data['name']];
			break;
		case 16:
			executeParameters=[data['uri'], data['name']];
			break;
		case 20:
			executeParameters=[data['attachmentIndex']];
			break;
		case 25:
			executeParameters=[data['TimeProperty'], data['time']];
			break;
		case 24:
			executeParameters=[data['TimeProperty']];
			break;
		case 27:
			executeParameters=[data['location']];
			break;
		default:
			Sys.Debug.fail('Unexpected method dispid');
			break
	}
	return executeParameters
};
OSF.DDA.OutlookAppOm._convertRecipientArrayParameterForOutlookForDisplayApi$p=function OSF_DDA_OutlookAppOm$_convertRecipientArrayParameterForOutlookForDisplayApi$p(array)
{
	return (array) ? array.join(';') : null
};
OSF.DDA.OutlookAppOm._convertComposeEmailDictionaryParameterForSetApi$p=function OSF_DDA_OutlookAppOm$_convertComposeEmailDictionaryParameterForSetApi$p(recipients)
{
	if (!recipients)
	{
		return null
	}
	var results=new Array(recipients.length);
	for (var i=0; i < recipients.length; i++)
	{
		results[i]=[recipients[i]['address'], recipients[i]['name']]
	}
	return results
};
OSF.DDA.OutlookAppOm._validateAndNormalizeRecipientEmails$p=function OSF_DDA_OutlookAppOm$_validateAndNormalizeRecipientEmails$p(emailset, name)
{
	if ($h.ScriptHelpers.isNullOrUndefined(emailset))
	{
		return null
	}
	OSF.DDA.OutlookAppOm._throwOnArgumentType$p(emailset, Array, name);
	var originalAttendees=emailset;
	var updatedAttendees=null;
	var normalizationNeeded=false;
	OSF.DDA.OutlookAppOm._throwOnOutOfRange$i(originalAttendees.length, 0, OSF.DDA.OutlookAppOm._maxRecipients$p, String.format('{0}.length', name));
	for (var i=0; i < originalAttendees.length; i++)
	{
		if ($h.EmailAddressDetails.isInstanceOfType(originalAttendees[i]))
		{
			normalizationNeeded=true;
			break
		}
	}
	if (normalizationNeeded)
	{
		updatedAttendees=[]
	}
	for (var i=0; i < originalAttendees.length; i++)
	{
		if (normalizationNeeded)
		{
			updatedAttendees[i]=($h.EmailAddressDetails.isInstanceOfType(originalAttendees[i])) ? (originalAttendees[i]).emailAddress : originalAttendees[i];
			OSF.DDA.OutlookAppOm._throwOnArgumentType$p(updatedAttendees[i], String, String.format('{0}[{1}]', name, i))
		}
		else
		{
			OSF.DDA.OutlookAppOm._throwOnArgumentType$p(originalAttendees[i], String, String.format('{0}[{1}]', name, i))
		}
	}
	return updatedAttendees
};
OSF.DDA.OutlookAppOm.prototype={
	_initialData$p$0: null, _item$p$0: null, _userProfile$p$0: null, _diagnostics$p$0: null, _officeAppContext$p$0: null, _appReadyCallback$p$0: null, _clientEndPoint$p$0: null, get_clientEndPoint: function OSF_DDA_OutlookAppOm$get_clientEndPoint()
		{
			if (!this._clientEndPoint$p$0)
			{
				this._clientEndPoint$p$0=OSF._OfficeAppFactory.getClientEndPoint()
			}
			return this._clientEndPoint$p$0
		}, set_clientEndPoint: function OSF_DDA_OutlookAppOm$set_clientEndPoint(value)
		{
			this._clientEndPoint$p$0=value;
			return value
		}, get_initialData: function OSF_DDA_OutlookAppOm$get_initialData()
		{
			return this._initialData$p$0
		}, get__appName$i$0: function OSF_DDA_OutlookAppOm$get__appName$i$0()
		{
			return this._officeAppContext$p$0.get_appName()
		}, initialize: function OSF_DDA_OutlookAppOm$initialize(initialData)
		{
			var ItemTypeKey='itemType';
			this._initialData$p$0=new $h.InitialData(initialData);
			if (1===initialData[ItemTypeKey])
			{
				this._item$p$0=new $h.Message(this._initialData$p$0)
			}
			else if (3===initialData[ItemTypeKey])
			{
				this._item$p$0=new $h.MeetingRequest(this._initialData$p$0)
			}
			else if (2===initialData[ItemTypeKey])
			{
				this._item$p$0=new $h.Appointment(this._initialData$p$0)
			}
			else if (4===initialData[ItemTypeKey])
			{
				this._item$p$0=new $h.MessageCompose(this._initialData$p$0)
			}
			else if (5===initialData[ItemTypeKey])
			{
				this._item$p$0=new $h.AppointmentCompose(this._initialData$p$0)
			}
			else
			{
				Sys.Debug.trace('Unexpected item type was received from the host.')
			}
			this._userProfile$p$0=new $h.UserProfile(this._initialData$p$0);
			this._diagnostics$p$0=new $h.Diagnostics(this._initialData$p$0, this._officeAppContext$p$0.get_appName());
			this._initializeMethods$p$0();
			$h.InitialData._defineReadOnlyProperty$i(this, 'item', this.$$d__getItem$p$0);
			$h.InitialData._defineReadOnlyProperty$i(this, 'userProfile', this.$$d__getUserProfile$p$0);
			$h.InitialData._defineReadOnlyProperty$i(this, 'diagnostics', this.$$d__getDiagnostics$p$0);
			$h.InitialData._defineReadOnlyProperty$i(this, 'ewsUrl', this.$$d__getEwsUrl$p$0);
			if (OSF.DDA.OutlookAppOm._instance$p.get__appName$i$0()===64)
			{
				if (this._initialData$p$0.get__overrideWindowOpen$i$0())
				{
					window.open=this.$$d_windowOpenOverrideHandler
				}
			}
		}, windowOpenOverrideHandler: function OSF_DDA_OutlookAppOm$windowOpenOverrideHandler(url, targetName, features, replace)
		{
			this._invokeHostMethod$i$0(0, 'LaunchPalUrl', {launchUrl: url}, null)
		}, makeEwsRequestAsync: function OSF_DDA_OutlookAppOm$makeEwsRequestAsync(data)
		{
			var args=[];
			for (var $$pai_5=1; $$pai_5 < arguments.length;++$$pai_5)
			{
				args[$$pai_5 - 1]=arguments[$$pai_5]
			}
			if ($h.ScriptHelpers.isNullOrUndefined(data))
			{
				throw Error.argumentNull('data');
			}
			if (data.length > OSF.DDA.OutlookAppOm._maxEwsRequestSize$p)
			{
				throw Error.argument('data', _u.ExtensibilityStrings.l_EwsRequestOversized_Text);
			}
			this._throwOnMethodCallForInsufficientPermission$i$0(3, 'makeEwsRequestAsync');
			var parameters=$h.CommonParameters.parse(args, true, true);
			var ewsRequest=new $h.EwsRequest(parameters._asyncContext$p$0);
			var $$t_4=this;
			ewsRequest.onreadystatechange=function()
			{
				if (4===ewsRequest.get__requestState$i$1())
				{
					parameters.get_callback()(ewsRequest._asyncResult$p$0)
				}
			};
			ewsRequest.send(data)
		}, recordDataPoint: function OSF_DDA_OutlookAppOm$recordDataPoint(data)
		{
			if ($h.ScriptHelpers.isNullOrUndefined(data))
			{
				throw Error.argumentNull('data');
			}
			this._invokeHostMethod$i$0(0, 'RecordDataPoint', data, null)
		}, recordTrace: function OSF_DDA_OutlookAppOm$recordTrace(data)
		{
			if ($h.ScriptHelpers.isNullOrUndefined(data))
			{
				throw Error.argumentNull('data');
			}
			this._invokeHostMethod$i$0(0, 'RecordTrace', data, null)
		}, trackCtq: function OSF_DDA_OutlookAppOm$trackCtq(data)
		{
			if ($h.ScriptHelpers.isNullOrUndefined(data))
			{
				throw Error.argumentNull('data');
			}
			this._invokeHostMethod$i$0(0, 'TrackCtq', data, null)
		}, convertToLocalClientTime: function OSF_DDA_OutlookAppOm$convertToLocalClientTime(timeValue)
		{
			var date=new Date(timeValue.getTime());
			var offset=date.getTimezoneOffset() * -1;
			if (this._initialData$p$0 && this._initialData$p$0.get__timeZoneOffsets$i$0())
			{
				date.setUTCMinutes(date.getUTCMinutes() - offset);
				offset=this._findOffset$p$0(date);
				date.setUTCMinutes(date.getUTCMinutes()+offset)
			}
			var retValue=this._dateToDictionary$i$0(date);
			retValue['timezoneOffset']=offset;
			return retValue
		}, convertToUtcClientTime: function OSF_DDA_OutlookAppOm$convertToUtcClientTime(input)
		{
			var retValue=this._dictionaryToDate$i$0(input);
			if (this._initialData$p$0 && this._initialData$p$0.get__timeZoneOffsets$i$0())
			{
				var offset=this._findOffset$p$0(retValue);
				retValue.setUTCMinutes(retValue.getUTCMinutes() - offset);
				offset=(!input['timezoneOffset']) ? retValue.getTimezoneOffset() * -1 : input['timezoneOffset'];
				retValue.setUTCMinutes(retValue.getUTCMinutes()+offset)
			}
			return retValue
		}, getUserIdentityTokenAsync: function OSF_DDA_OutlookAppOm$getUserIdentityTokenAsync()
		{
			var args=[];
			for (var $$pai_2=0; $$pai_2 < arguments.length;++$$pai_2)
			{
				args[$$pai_2]=arguments[$$pai_2]
			}
			this._throwOnMethodCallForInsufficientPermission$i$0(1, 'getUserIdentityTokenAsync');
			var parameters=$h.CommonParameters.parse(args, true, true);
			this._invokeGetTokenMethodAsync$p$0(2, 'GetUserIdentityToken', parameters.get_callback(), parameters._asyncContext$p$0)
		}, getCallbackTokenAsync: function OSF_DDA_OutlookAppOm$getCallbackTokenAsync()
		{
			var args=[];
			for (var $$pai_2=0; $$pai_2 < arguments.length;++$$pai_2)
			{
				args[$$pai_2]=arguments[$$pai_2]
			}
			this._throwOnMethodCallForInsufficientPermission$i$0(1, 'getCallbackTokenAsync');
			var parameters=$h.CommonParameters.parse(args, true, true);
			this._invokeGetTokenMethodAsync$p$0(12, 'GetCallbackToken', parameters.get_callback(), parameters._asyncContext$p$0)
		}, displayMessageForm: function OSF_DDA_OutlookAppOm$displayMessageForm(itemId)
		{
			if ($h.ScriptHelpers.isNullOrUndefined(itemId))
			{
				throw Error.argumentNull('itemId');
			}
			this._invokeHostMethod$i$0(8, 'DisplayExistingMessageForm', {itemId: itemId}, null)
		}, displayAppointmentForm: function OSF_DDA_OutlookAppOm$displayAppointmentForm(itemId)
		{
			if ($h.ScriptHelpers.isNullOrUndefined(itemId))
			{
				throw Error.argumentNull('itemId');
			}
			this._invokeHostMethod$i$0(9, 'DisplayExistingAppointmentForm', {itemId: itemId}, null)
		}, createAsyncResult: function OSF_DDA_OutlookAppOm$createAsyncResult(value, errorCode, errorDescription, userContext)
		{
			var initArgs={};
			var errorArgs=null;
			initArgs[OSF.DDA.AsyncResultEnum.Properties.Value]=value;
			initArgs[OSF.DDA.AsyncResultEnum.Properties.Context]=userContext;
			if (0 !==errorCode)
			{
				errorArgs={};
				errorArgs[OSF.DDA.AsyncResultEnum.ErrorProperties.Name]=errorCode;
				errorArgs[OSF.DDA.AsyncResultEnum.ErrorProperties.Message]=errorDescription
			}
			return new OSF.DDA.AsyncResult(initArgs, errorArgs)
		}, standardCreateAsyncResult: function OSF_DDA_OutlookAppOm$standardCreateAsyncResult(value, errorCode, detailedErrorCode, userContext)
		{
			var initArgs={};
			var errorArgs=null;
			initArgs[OSF.DDA.AsyncResultEnum.Properties.Value]=value;
			initArgs[OSF.DDA.AsyncResultEnum.Properties.Context]=userContext;
			if (0 !==errorCode)
			{
				errorArgs={};
				var errorProperties=$h.OutlookErrorManager.getErrorArgs(detailedErrorCode);
				errorArgs[OSF.DDA.AsyncResultEnum.ErrorProperties.Name]=errorProperties['name'];
				errorArgs[OSF.DDA.AsyncResultEnum.ErrorProperties.Message]=errorProperties['message'];
				errorArgs[OSF.DDA.AsyncResultEnum.ErrorProperties.Code]=detailedErrorCode
			}
			return new O
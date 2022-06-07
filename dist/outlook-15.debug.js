/* Outlook specific API library */
/* Version: 15.0.4420.1017 Build Time: 03/31/2014 */
/*
	Copyright (c) Microsoft Corporation.  All rights reserved.
*/

/*
	Your use of this file is governed by the Microsoft Services Agreement http://go.microsoft.com/fwlink/?LinkId=266419.
*/

Type.registerNamespace('Microsoft.Office.WebExtension.MailboxEnums');
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
	Message: 'message',
	Appointment: 'appointment'
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
Type.registerNamespace('OSF.DDA');
OSF.DDA.OutlookAppOm=function OSF_DDA_OutlookAppOm(officeAppContext, targetWindow, appReadyCallback) {
	this.$$d__callAppReadyCallback$p$0=Function.createDelegate(this, this._callAppReadyCallback$p$0);
	this.$$d__getEwsUrl$p$0=Function.createDelegate(this, this._getEwsUrl$p$0);
	this.$$d__getDiagnostics$p$0=Function.createDelegate(this, this._getDiagnostics$p$0);
	this.$$d__getUserProfile$p$0=Function.createDelegate(this, this._getUserProfile$p$0);
	this.$$d__getItem$p$0=Function.createDelegate(this, this._getItem$p$0);
	this.$$d__getInitialDataResponseHandler$p$0=Function.createDelegate(this, this._getInitialDataResponseHandler$p$0);
	OSF.DDA.OutlookAppOm._instance$p=this;
	this._officeAppContext$p$0=officeAppContext;
	this._appReadyCallback$p$0=appReadyCallback;
	var $$t_4=this;
	var stringLoadedCallback=function() {
		if (appReadyCallback) {
			$$t_4._invokeHostMethod$i$0(1, 'GetInitialData', null, $$t_4.$$d__getInitialDataResponseHandler$p$0);
		}
	};
	if (this._areStringsLoaded$p$0()) {
		stringLoadedCallback();
	}
	else {
		this._loadLocalizedScript$p$0(stringLoadedCallback);
	}
}
OSF.DDA.OutlookAppOm._createAsyncResult$i=function OSF_DDA_OutlookAppOm$_createAsyncResult$i(value, errorCode, errorDescription, userContext) {
	var initArgs={};
	initArgs[OSF.DDA.AsyncResultEnum.Properties.Value]=value;
	initArgs[OSF.DDA.AsyncResultEnum.Properties.Context]=userContext;
	var errorArgs=null;
	if (0 !==errorCode) {
		errorArgs={};
		errorArgs[OSF.DDA.AsyncResultEnum.ErrorProperties.Name]=errorCode;
		errorArgs[OSF.DDA.AsyncResultEnum.ErrorProperties.Message]=errorDescription;
	}
	return new OSF.DDA.AsyncResult(initArgs, errorArgs);
}
OSF.DDA.OutlookAppOm._throwOnPropertyAccessForRestrictedPermission$i=function OSF_DDA_OutlookAppOm$_throwOnPropertyAccessForRestrictedPermission$i(currentPermissionLevel) {
	if (!currentPermissionLevel) {
		throw Error.create(_u.ExtensibilityStrings.l_ElevatedPermissionNeeded_Text);
	}
}
OSF.DDA.OutlookAppOm._throwOnMethodCallForInsufficientPermission$i=function OSF_DDA_OutlookAppOm$_throwOnMethodCallForInsufficientPermission$i(currentPermissionLevel, requiredPermissionLevel, methodName) {
	if (currentPermissionLevel < requiredPermissionLevel) {
		throw Error.create(String.format(_u.ExtensibilityStrings.l_ElevatedPermissionNeededForMethod_Text, methodName));
	}
}
OSF.DDA.OutlookAppOm._throwOnArgumentType$p=function OSF_DDA_OutlookAppOm$_throwOnArgumentType$p(value, expectedType, argumentName) {
	if (Object.getType(value) !==expectedType) {
		throw Error.argumentType(argumentName);
	}
}
OSF.DDA.OutlookAppOm._throwOnOutOfRange$p=function OSF_DDA_OutlookAppOm$_throwOnOutOfRange$p(value, minValue, maxValue, argumentName) {
	if (value < minValue || value > maxValue) {
		throw Error.argumentOutOfRange(argumentName);
	}
}
OSF.DDA.OutlookAppOm._validateOptionalStringParameter$p=function OSF_DDA_OutlookAppOm$_validateOptionalStringParameter$p(value, minLength, maxLength, name) {
	if ($h.ScriptHelpers.isNullOrUndefined(value)) {
		return;
	}
	OSF.DDA.OutlookAppOm._throwOnArgumentType$p(value, String, name);
	var stringValue=value;
	OSF.DDA.OutlookAppOm._throwOnOutOfRange$p(stringValue.length, minLength, maxLength, name);
}
OSF.DDA.OutlookAppOm._convertToOutlookParameters$p=function OSF_DDA_OutlookAppOm$_convertToOutlookParameters$p(dispid, data) {
	var executeParameters=null;
	switch (dispid) {
		case 1:
		case 2:
		case 12:
		case 3:
			break;
		case 4:
			var jsonProperty=JSON.stringify(data['customProperties']);
			executeParameters=[ jsonProperty ];
			break;
		case 5:
			executeParameters=[ data['body'] ];
			break;
		case 8:
		case 9:
			executeParameters=[ data['itemId'] ];
			break;
		case 7:
			executeParameters=[ OSF.DDA.OutlookAppOm._convertRecipientArrayParameterForOutlook$p(data['requiredAttendees']), OSF.DDA.OutlookAppOm._convertRecipientArrayParameterForOutlook$p(data['optionalAttendees']), data['start'], data['end'], data['location'], OSF.DDA.OutlookAppOm._convertRecipientArrayParameterForOutlook$p(data['resources']), data['subject'], data['body'] ];
			break;
		case 11:
		case 10:
			executeParameters=[ data['htmlBody'] ];
			break;
		default:
			Sys.Debug.fail('Unexpected method dispid');
			break;
	}
	return executeParameters;
}
OSF.DDA.OutlookAppOm._convertRecipientArrayParameterForOutlook$p=function OSF_DDA_OutlookAppOm$_convertRecipientArrayParameterForOutlook$p(array) {
	return (array) ? array.join(';') : null;
}
OSF.DDA.OutlookAppOm._validateAndNormalizeRecipientEmails$p=function OSF_DDA_OutlookAppOm$_validateAndNormalizeRecipientEmails$p(emailset, name) {
	if ($h.ScriptHelpers.isNullOrUndefined(emailset)) {
		return null;
	}
	OSF.DDA.OutlookAppOm._throwOnArgumentType$p(emailset, Array, name);
	var originalAttendees=emailset;
	var updatedAttendees=null;
	var normalizationNeeded=false;
	OSF.DDA.OutlookAppOm._throwOnOutOfRange$p(originalAttendees.length, 0, OSF.DDA.OutlookAppOm._maxRecipients$p, String.format('{0}.length', name));
	for (var i=0; i < originalAttendees.length; i++) {
		if ($h.EmailAddressDetails.isInstanceOfType(originalAttendees[i])) {
			normalizationNeeded=true;
			break;
		}
	}
	if (normalizationNeeded) {
		updatedAttendees=[];
	}
	for (var i=0; i < originalAttendees.length; i++) {
		if (normalizationNeeded) {
			updatedAttendees[i]=($h.EmailAddressDetails.isInstanceOfType(originalAttendees[i])) ? (originalAttendees[i]).emailAddress : originalAttendees[i];
			OSF.DDA.OutlookAppOm._throwOnArgumentType$p(updatedAttendees[i], String, String.format('{0}[{1}]', name, i));
		}
		else {
			OSF.DDA.OutlookAppOm._throwOnArgumentType$p(originalAttendees[i], String, String.format('{0}[{1}]', name, i));
		}
	}
	return updatedAttendees;
}
OSF.DDA.OutlookAppOm.prototype={
	_initialData$p$0: null,
	_item$p$0: null,
	_userProfile$p$0: null,
	_diagnostics$p$0: null,
	_officeAppContext$p$0: null,
	_appReadyCallback$p$0: null,
	get__appName$i$0: function OSF_DDA_OutlookAppOm$get__appName$i$0() {
		return this._officeAppContext$p$0.get_appName();
	},
	initialize: function OSF_DDA_OutlookAppOm$initialize(initialData) {
		var ItemTypeKey='itemType';
		this._initialData$p$0=new $h.InitialData(initialData);
		if (1===initialData[ItemTypeKey]) {
			this._item$p$0=new $h.Message(this._initialData$p$0);
		}
		else if (3===initialData[ItemTypeKey]) {
			this._item$p$0=new $h.MeetingRequest(this._initialData$p$0);
		}
		else if (2===initialData[ItemTypeKey]) {
			this._item$p$0=new $h.Appointment(this._initialData$p$0);
		}
		else {
			Sys.Debug.trace('Unexpected item type was received from the host.');
		}
		this._userProfile$p$0=new $h.UserProfile(this._initialData$p$0);
		this._diagnostics$p$0=new $h.Diagnostics(this._initialData$p$0, this._officeAppContext$p$0.get_appName());
		$h.InitialData._defineReadOnlyProperty$i(this, 'item', this.$$d__getItem$p$0);
		$h.InitialData._defineReadOnlyProperty$i(this, 'userProfile', this.$$d__getUserProfile$p$0);
		$h.InitialData._defineReadOnlyProperty$i(this, 'diagnostics', this.$$d__getDiagnostics$p$0);
		if (OSF.DDA.OutlookAppOm._instance$p.get__appName$i$0()===64) {
			$h.InitialData._defineReadOnlyProperty$i(this, 'ewsUrl', this.$$d__getEwsUrl$p$0);
		}
	},
	makeEwsRequestAsync: function OSF_DDA_OutlookAppOm$makeEwsRequestAsync(data, callback, userContext) {
		if ($h.ScriptHelpers.isNullOrUndefined(data)) {
			throw Error.argumentNull('data');
		}
		if (data.length > OSF.DDA.OutlookAppOm._maxEwsRequestSize$p) {
			throw Error.argument('data', _u.ExtensibilityStrings.l_EwsRequestOversized_Text);
		}
		OSF.DDA.OutlookAppOm._throwOnMethodCallForInsufficientPermission$i(this._initialData$p$0.get__permissionLevel$i$0(), 2, 'makeEwsRequestAsync');
		var ewsRequest=new $h.EwsRequest(userContext);
		var $$t_4=this;
		ewsRequest.onreadystatechange=function() {
			if (4===ewsRequest.get__requestState$i$1()) {
				callback(ewsRequest._asyncResult$p$0);
			}
		};
		ewsRequest.send(data);
	},
	recordDataPoint: function OSF_DDA_OutlookAppOm$recordDataPoint(data) {
		if ($h.ScriptHelpers.isNullOrUndefined(data)) {
			throw Error.argumentNull('data');
		}
		this._invokeHostMethod$i$0(0, 'RecordDataPoint', data, null);
	},
	recordTrace: function OSF_DDA_OutlookAppOm$recordTrace(data) {
		if ($h.ScriptHelpers.isNullOrUndefined(data)) {
			throw Error.argumentNull('data');
		}
		this._invokeHostMethod$i$0(0, 'RecordTrace', data, null);
	},
	trackCtq: function OSF_DDA_OutlookAppOm$trackCtq(data) {
		if ($h.ScriptHelpers.isNullOrUndefined(data)) {
			throw Error.argumentNull('data');
		}
		this._invokeHostMethod$i$0(0, 'TrackCtq', data, null);
	},
	convertToLocalClientTime: function OSF_DDA_OutlookAppOm$convertToLocalClientTime(timeValue) {
		var date=new Date(timeValue.getTime());
		var offset=date.getTimezoneOffset() * -1;
		if (this._initialData$p$0 && this._initialData$p$0.get__timeZoneOffsets$i$0()) {
			date.setUTCMinutes(date.getUTCMinutes() - offset);
			offset=this._findOffset$p$0(date);
			date.setUTCMinutes(date.getUTCMinutes()+offset);
		}
		var retValue=this._dateToDictionary$i$0(date);
		retValue['timezoneOffset']=offset;
		return retValue;
	},
	convertToUtcClientTime: function OSF_DDA_OutlookAppOm$convertToUtcClientTime(input) {
		var retValue=this._dictionaryToDate$i$0(input);
		if (this._initialData$p$0 && this._initialData$p$0.get__timeZoneOffsets$i$0()) {
			var offset=this._findOffset$p$0(retValue);
			retValue.setUTCMinutes(retValue.getUTCMinutes() - offset);
			offset=(!input['timezoneOffset']) ? retValue.getTimezoneOffset() * -1 : input['timezoneOffset'];
			retValue.setUTCMinutes(retValue.getUTCMinutes()+offset);
		}
		return retValue;
	},
	getUserIdentityTokenAsync: function OSF_DDA_OutlookAppOm$getUserIdentityTokenAsync(callback, userContext) {
		OSF.DDA.OutlookAppOm._throwOnMethodCallForInsufficientPermission$i(this._initialData$p$0.get__permissionLevel$i$0(), 1, 'getUserIdentityTokenAsync');
		this._invokeGetTokenMethodAsync$p$0(2, 'GetUserIdentityToken', callback, userContext);
	},
	getCallbackTokenAsync: function OSF_DDA_OutlookAppOm$getCallbackTokenAsync(callback, userContext) {
		OSF.DDA.OutlookAppOm._throwOnMethodCallForInsufficientPermission$i(this._initialData$p$0.get__permissionLevel$i$0(), 1, 'getCallbackTokenAsync');
		if (64 !==this._officeAppContext$p$0.get_appName()) {
			throw Error.notImplemented('The getCallbackTokenAsync is not supported by outlook for now.');
		}
		this._invokeGetTokenMethodAsync$p$0(12, 'GetCallbackToken', callback, userContext);
	},
	displayMessageForm: function OSF_DDA_OutlookAppOm$displayMessageForm(itemId) {
		if ($h.ScriptHelpers.isNullOrUndefined(itemId)) {
			throw Error.argumentNull('itemId');
		}
		this._invokeHostMethod$i$0(8, 'DisplayExistingMessageForm', { itemId: itemId }, null);
	},
	displayAppointmentForm: function OSF_DDA_OutlookAppOm$displayAppointmentForm(itemId) {
		if ($h.ScriptHelpers.isNullOrUndefined(itemId)) {
			throw Error.argumentNull('itemId');
		}
		this._invokeHostMethod$i$0(9, 'DisplayExistingAppointmentForm', { itemId: itemId }, null);
	},
	displayNewAppointmentForm: function OSF_DDA_OutlookAppOm$displayNewAppointmentForm(parameters) {
		var normalizedRequiredAttendees=OSF.DDA.OutlookAppOm._validateAndNormalizeRecipientEmails$p(parameters['requiredAttendees'], 'requiredAttendees');
		var normalizedOptionalAttendees=OSF.DDA.OutlookAppOm._validateAndNormalizeRecipientEmails$p(parameters['optionalAttendees'], 'optionalAttendees');
		OSF.DDA.OutlookAppOm._validateOptionalStringParameter$p(parameters['location'], 0, OSF.DDA.OutlookAppOm._maxLocationLength$p, 'location');
		OSF.DDA.OutlookAppOm._validateOptionalStringParameter$p(parameters['body'], 0, OSF.DDA.OutlookAppOm._maxBodyLength$p, 'body');
		OSF.DDA.OutlookAppOm._validateOptionalStringParameter$p(parameters['subject'], 0, OSF.DDA.OutlookAppOm._maxSubjectLength$p, 'subject');
		if (!$h.ScriptHelpers.isNullOrUndefined(parameters['start'])) {
			OSF.DDA.OutlookAppOm._throwOnArgumentType$p(parameters['start'], Date, 'start');
			var startDateTime=parameters['start'];
			parameters['start']=startDateTime.getTime();
			if (!$h.ScriptHelpers.isNullOrUndefined(parameters['end'])) {
				OSF.DDA.OutlookAppOm._throwOnArgumentType$p(parameters['end'], Date, 'end');
				var endDateTime=parameters['end'];
				if (endDateTime < startDateTime) {
					throw Error.argumentOutOfRange('end', endDateTime, _u.ExtensibilityStrings.l_InvalidEventDates_Text);
				}
				parameters['end']=endDateTime.getTime();
			}
		}
		var updatedParameters=null;
		if (normalizedRequiredAttendees || normalizedOptionalAttendees) {
			updatedParameters={};
			var $$dict_6=parameters;
			for (var $$key_7 in $$dict_6) {
				var entry={ key: $$key_7, value: $$dict_6[$$key_7] };
				updatedParameters[entry.key]=entry.value;
			}
			if (normalizedRequiredAttendees) {
				updatedParameters['requiredAttendees']=normalizedRequiredAttendees;
			}
			if (normalizedOptionalAttendees) {
				updatedParameters['optionalAttendees']=normalizedOptionalAttendees;
			}
		}
		this._invokeHostMethod$i$0(7, 'DisplayNewAppointmentForm', updatedParameters || parameters, null);
	},
	_displayReplyForm$i$0: function OSF_DDA_OutlookAppOm$_displayReplyForm$i$0(htmlBody) {
		if (!$h.ScriptHelpers.isNullOrUndefined(htmlBody)) {
			OSF.DDA.OutlookAppOm._throwOnOutOfRange$p(htmlBody.length, 0, OSF.DDA.OutlookAppOm._maxBodyLength$p, 'htmlBody');
		}
		this._invokeHostMethod$i$0(10, 'DisplayReplyForm', { htmlBody: htmlBody }, null);
	},
	_displayReplyAllForm$i$0: function OSF_DDA_OutlookAppOm$_displayReplyAllForm$i$0(htmlBody) {
		if (!$h.ScriptHelpers.isNullOrUndefined(htmlBody)) {
			OSF.DDA.OutlookAppOm._throwOnOutOfRange$p(htmlBody.length, 0, OSF.DDA.OutlookAppOm._maxBodyLength$p, 'htmlBody');
		}
		this._invokeHostMethod$i$0(11, 'DisplayReplyAllForm', { htmlBody: htmlBody }, null);
	},
	_invokeHostMethod$i$0: function OSF_DDA_OutlookAppOm$_invokeHostMethod$i$0(dispid, name, data, responseCallback) {
		if (64===this._officeAppContext$p$0.get_appName()) {
			OSF._OfficeAppFactory.getClientEndPoint().invoke(name, responseCallback, data);
		}
		else if (dispid) {
			var executeParameters=OSF.DDA.OutlookAppOm._convertToOutlookParameters$p(dispid, data);
			var $$t_9=this;
			window.external.Execute(dispid, executeParameters, function(nativeData, resultCode) {
				if (responseCallback) {
					var serializedData=nativeData.getItem(0);
					var deserializedData=JSON.parse(serializedData);
					responseCallback(resultCode, deserializedData);
				}
			});
		}
		else if (responseCallback) {
			responseCallback(-2, null);

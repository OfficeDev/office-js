/* Excel Desktop-specific API library */
/* Version: 16.0.8613.3000 */

/* Office.js Version: 16.0.8616.1000 */ 
/*
	Copyright (c) Microsoft Corporation.  All rights reserved.
*/

/*
	Your use of this file is governed by the Microsoft Services Agreement http://go.microsoft.com/fwlink/?LinkId=266419.
*/

/*
* @overview es6-promise - a tiny implementation of Promises/A+.
* @copyright Copyright (c) 2014 Yehuda Katz, Tom Dale, Stefan Penner and contributors (Conversion to ES6 API by Jake Archibald)
* @license   Licensed under MIT license
*            See https://raw.githubusercontent.com/jakearchibald/es6-promise/master/LICENSE
* @version   2.3.0
*/

var __extends=(this && this.__extends) || function (d, b) {
	for (var p in b) if (b.hasOwnProperty(p)) d[p]=b[p];
	function __() { this.constructor=d; }
	d.prototype=b===null ? Object.create(b) : (__.prototype=b.prototype, new __());
};
var OfficeExt;
(function (OfficeExt) {
	var MicrosoftAjaxFactory=(function () {
		function MicrosoftAjaxFactory() {
		}
		MicrosoftAjaxFactory.prototype.isMsAjaxLoaded=function () {
			if (typeof (Sys) !=='undefined' && typeof (Type) !=='undefined' &&
				Sys.StringBuilder && typeof (Sys.StringBuilder)==="function" &&
				Type.registerNamespace && typeof (Type.registerNamespace)==="function" &&
				Type.registerClass && typeof (Type.registerClass)==="function" &&
				typeof (Function._validateParams)==="function" &&
				Sys.Serialization && Sys.Serialization.JavaScriptSerializer && typeof (Sys.Serialization.JavaScriptSerializer.serialize)==="function") {
				return true;
			}
			else {
				return false;
			}
		};
		MicrosoftAjaxFactory.prototype.loadMsAjaxFull=function (callback) {
			var msAjaxCDNPath=(window.location.protocol.toLowerCase()==='https:' ? 'https:' : 'http:')+'//ajax.aspnetcdn.com/ajax/3.5/MicrosoftAjax.js';
			OSF.OUtil.loadScript(msAjaxCDNPath, callback);
		};
		Object.defineProperty(MicrosoftAjaxFactory.prototype, "msAjaxError", {
			get: function () {
				if (this._msAjaxError==null && this.isMsAjaxLoaded()) {
					this._msAjaxError=Error;
				}
				return this._msAjaxError;
			},
			set: function (errorClass) {
				this._msAjaxError=errorClass;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(MicrosoftAjaxFactory.prototype, "msAjaxString", {
			get: function () {
				if (this._msAjaxString==null && this.isMsAjaxLoaded()) {
					this._msAjaxString=String;
				}
				return this._msAjaxString;
			},
			set: function (stringClass) {
				this._msAjaxString=stringClass;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(MicrosoftAjaxFactory.prototype, "msAjaxDebug", {
			get: function () {
				if (this._msAjaxDebug==null && this.isMsAjaxLoaded()) {
					this._msAjaxDebug=Sys.Debug;
				}
				return this._msAjaxDebug;
			},
			set: function (debugClass) {
				this._msAjaxDebug=debugClass;
			},
			enumerable: true,
			configurable: true
		});
		return MicrosoftAjaxFactory;
	})();
	OfficeExt.MicrosoftAjaxFactory=MicrosoftAjaxFactory;
})(OfficeExt || (OfficeExt={}));
var OsfMsAjaxFactory=new OfficeExt.MicrosoftAjaxFactory();
var OSF=OSF || {};
var OfficeExt;
(function (OfficeExt) {
	var SafeStorage=(function () {
		function SafeStorage(_internalStorage) {
			this._internalStorage=_internalStorage;
		}
		SafeStorage.prototype.getItem=function (key) {
			try {
				return this._internalStorage && this._internalStorage.getItem(key);
			}
			catch (e) {
				return null;
			}
		};
		SafeStorage.prototype.setItem=function (key, data) {
			try {
				this._internalStorage && this._internalStorage.setItem(key, data);
			}
			catch (e) {
			}
		};
		SafeStorage.prototype.clear=function () {
			try {
				this._internalStorage && this._internalStorage.clear();
			}
			catch (e) {
			}
		};
		SafeStorage.prototype.removeItem=function (key) {
			try {
				this._internalStorage && this._internalStorage.removeItem(key);
			}
			catch (e) {
			}
		};
		SafeStorage.prototype.getKeysWithPrefix=function (keyPrefix) {
			var keyList=[];
			try {
				var len=this._internalStorage && this._internalStorage.length || 0;
				for (var i=0; i < len; i++) {
					var key=this._internalStorage.key(i);
					if (key.indexOf(keyPrefix)===0) {
						keyList.push(key);
					}
				}
			}
			catch (e) {
			}
			return keyList;
		};
		return SafeStorage;
	})();
	OfficeExt.SafeStorage=SafeStorage;
})(OfficeExt || (OfficeExt={}));
OSF.XdmFieldName={
	ConversationUrl: "ConversationUrl",
	AppId: "AppId"
};
OSF.WindowNameItemKeys={
	BaseFrameName: "baseFrameName",
	HostInfo: "hostInfo",
	XdmInfo: "xdmInfo",
	SerializerVersion: "serializerVersion",
	AppContext: "appContext"
};
OSF.OUtil=(function () {
	var _uniqueId=-1;
	var _xdmInfoKey='&_xdm_Info=';
	var _serializerVersionKey='&_serializer_version=';
	var _xdmSessionKeyPrefix='_xdm_';
	var _serializerVersionKeyPrefix='_serializer_version=';
	var _fragmentSeparator='#';
	var _fragmentInfoDelimiter='&';
	var _classN="class";
	var _loadedScripts={};
	var _defaultScriptLoadingTimeout=30000;
	var _safeSessionStorage=null;
	var _safeLocalStorage=null;
	var _rndentropy=new Date().getTime();
	function _random() {
		var nextrand=0x7fffffff * (Math.random());
		nextrand ^=_rndentropy ^ ((new Date().getMilliseconds()) << Math.floor(Math.random() * (31 - 10)));
		return nextrand.toString(16);
	}
	;
	function _getSessionStorage() {
		if (!_safeSessionStorage) {
			try {
				var sessionStorage=window.sessionStorage;
			}
			catch (ex) {
				sessionStorage=null;
			}
			_safeSessionStorage=new OfficeExt.SafeStorage(sessionStorage);
		}
		return _safeSessionStorage;
	}
	;
	function _reOrderTabbableElements(elements) {
		var bucket0=[];
		var bucketPositive=[];
		var i;
		var len=elements.length;
		var ele;
		for (i=0; i < len; i++) {
			ele=elements[i];
			if (ele.tabIndex) {
				if (ele.tabIndex > 0) {
					bucketPositive.push(ele);
				}
				else if (ele.tabIndex===0) {
					bucket0.push(ele);
				}
			}
			else {
				bucket0.push(ele);
			}
		}
		bucketPositive=bucketPositive.sort(function (left, right) {
			var diff=left.tabIndex - right.tabIndex;
			if (diff===0) {
				diff=bucketPositive.indexOf(left) - bucketPositive.indexOf(right);
			}
			return diff;
		});
		return [].concat(bucketPositive, bucket0);
	}
	;
	return {
		set_entropy: function OSF_OUtil$set_entropy(entropy) {
			if (typeof entropy=="string") {
				for (var i=0; i < entropy.length; i+=4) {
					var temp=0;
					for (var j=0; j < 4 && i+j < entropy.length; j++) {
						temp=(temp << 8)+entropy.charCodeAt(i+j);
					}
					_rndentropy ^=temp;
				}
			}
			else if (typeof entropy=="number") {
				_rndentropy ^=entropy;
			}
			else {
				_rndentropy ^=0x7fffffff * Math.random();
			}
			_rndentropy &=0x7fffffff;
		},
		extend: function OSF_OUtil$extend(child, parent) {
			var F=function () { };
			F.prototype=parent.prototype;
			child.prototype=new F();
			child.prototype.constructor=child;
			child.uber=parent.prototype;
			if (parent.prototype.constructor===Object.prototype.constructor) {
				parent.prototype.constructor=parent;
			}
		},
		setNamespace: function OSF_OUtil$setNamespace(name, parent) {
			if (parent && name && !parent[name]) {
				parent[name]={};
			}
		},
		unsetNamespace: function OSF_OUtil$unsetNamespace(name, parent) {
			if (parent && name && parent[name]) {
				delete parent[name];
			}
		},
		loadScript: function OSF_OUtil$loadScript(url, callback, timeoutInMs) {
			if (url && callback) {
				var doc=window.document;
				var _loadedScriptEntry=_loadedScripts[url];
				if (!_loadedScriptEntry) {
					var script=doc.createElement("script");
					script.type="text/javascript";
					_loadedScriptEntry={ loaded: false, pendingCallbacks: [callback], timer: null };
					_loadedScripts[url]=_loadedScriptEntry;
					var onLoadCallback=function OSF_OUtil_loadScript$onLoadCallback() {
						if (_loadedScriptEntry.timer !=null) {
							clearTimeout(_loadedScriptEntry.timer);
							delete _loadedScriptEntry.timer;
						}
						_loadedScriptEntry.loaded=true;
						var pendingCallbackCount=_loadedScriptEntry.pendingCallbacks.length;
						for (var i=0; i < pendingCallbackCount; i++) {
							var currentCallback=_loadedScriptEntry.pendingCallbacks.shift();
							currentCallback();
						}
					};
					var onLoadError=function OSF_OUtil_loadScript$onLoadError() {
						delete _loadedScripts[url];
						if (_loadedScriptEntry.timer !=null) {
							clearTimeout(_loadedScriptEntry.timer);
							delete _loadedScriptEntry.timer;
						}
						var pendingCallbackCount=_loadedScriptEntry.pendingCallbacks.length;
						for (var i=0; i < pendingCallbackCount; i++) {
							var currentCallback=_loadedScriptEntry.pendingCallbacks.shift();
							currentCallback();
						}
					};
					if (script.readyState) {
						script.onreadystatechange=function () {
							if (script.readyState=="loaded" || script.readyState=="complete") {
								script.onreadystatechange=null;
								onLoadCallback();
							}
						};
					}
					else {
						script.onload=onLoadCallback;
					}
					script.onerror=onLoadError;
					timeoutInMs=timeoutInMs || _defaultScriptLoadingTimeout;
					_loadedScriptEntry.timer=setTimeout(onLoadError, timeoutInMs);
					script.setAttribute("crossOrigin", "anonymous");
					script.src=url;
					doc.getElementsByTagName("head")[0].appendChild(script);
				}
				else if (_loadedScriptEntry.loaded) {
					callback();
				}
				else {
					_loadedScriptEntry.pendingCallbacks.push(callback);
				}
			}
		},
		loadCSS: function OSF_OUtil$loadCSS(url) {
			if (url) {
				var doc=window.document;
				var link=doc.createElement("link");
				link.type="text/css";
				link.rel="stylesheet";
				link.href=url;
				doc.getElementsByTagName("head")[0].appendChild(link);
			}
		},
		parseEnum: function OSF_OUtil$parseEnum(str, enumObject) {
			var parsed=enumObject[str.trim()];
			if (typeof (parsed)=='undefined') {
				OsfMsAjaxFactory.msAjaxDebug.trace("invalid enumeration string:"+str);
				throw OsfMsAjaxFactory.msAjaxError.argument("str");
			}
			return parsed;
		},
		delayExecutionAndCache: function OSF_OUtil$delayExecutionAndCache() {
			var obj={ calc: arguments[0] };
			return function () {
				if (obj.calc) {
					obj.val=obj.calc.apply(this, arguments);
					delete obj.calc;
				}
				return obj.val;
			};
		},
		getUniqueId: function OSF_OUtil$getUniqueId() {
			_uniqueId=_uniqueId+1;
			return _uniqueId.toString();
		},
		formatString: function OSF_OUtil$formatString() {
			var args=arguments;
			var source=args[0];
			return source.replace(/{(\d+)}/gm, function (match, number) {
				var index=parseInt(number, 10)+1;
				return args[index]===undefined ? '{'+number+'}' : args[index];
			});
		},
		generateConversationId: function OSF_OUtil$generateConversationId() {
			return [_random(), _random(), (new Date()).getTime().toString()].join('_');
		},
		getFrameName: function OSF_OUtil$getFrameName(cacheKey) {
			return _xdmSessionKeyPrefix+cacheKey+this.generateConversationId();
		},
		addXdmInfoAsHash: function OSF_OUtil$addXdmInfoAsHash(url, xdmInfoValue) {
			return OSF.OUtil.addInfoAsHash(url, _xdmInfoKey, xdmInfoValue, false);
		},
		addSerializerVersionAsHash: function OSF_OUtil$addSerializerVersionAsHash(url, serializerVersion) {
			return OSF.OUtil.addInfoAsHash(url, _serializerVersionKey, serializerVersion, true);
		},
		addInfoAsHash: function OSF_OUtil$addInfoAsHash(url, keyName, infoValue, encodeInfo) {
			url=url.trim() || '';
			var urlParts=url.split(_fragmentSeparator);
			var urlWithoutFragment=urlParts.shift();
			var fragment=urlParts.join(_fragmentSeparator);
			var newFragment;
			if (encodeInfo) {
				newFragment=[keyName, encodeURIComponent(infoValue), fragment].join('');
			}
			else {
				newFragment=[fragment, keyName, infoValue].join('');
			}
			return [urlWithoutFragment, _fragmentSeparator, newFragment].join('');
		},
		parseHostInfoFromWindowName: function OSF_OUtil$parseHostInfoFromWindowName(skipSessionStorage, windowName) {
			return OSF.OUtil.parseInfoFromWindowName(skipSessionStorage, windowName, OSF.WindowNameItemKeys.HostInfo);
		},
		parseXdmInfo: function OSF_OUtil$parseXdmInfo(skipSessionStorage) {
			var xdmInfoValue=OSF.OUtil.parseXdmInfoWithGivenFragment(skipSessionStorage, window.location.hash);
			if (!xdmInfoValue) {
				xdmInfoValue=OSF.OUtil.parseXdmInfoFromWindowName(skipSessionStorage, window.name);
			}
			return xdmInfoValue;
		},
		parseXdmInfoFromWindowName: function OSF_OUtil$parseXdmInfoFromWindowName(skipSessionStorage, windowName) {
			return OSF.OUtil.parseInfoFromWindowName(skipSessionStorage, windowName, OSF.WindowNameItemKeys.XdmInfo);
		},
		parseXdmInfoWithGivenFragment: function OSF_OUtil$parseXdmInfoWithGivenFragment(skipSessionStorage, fragment) {
			return OSF.OUtil.parseInfoWithGivenFragment(_xdmInfoKey, _xdmSessionKeyPrefix, false, skipSessionStorage, fragment);
		},
		parseSerializerVersion: function OSF_OUtil$parseSerializerVersion(skipSessionStorage) {
			var serializerVersion=OSF.OUtil.parseSerializerVersionWithGivenFragment(skipSessionStorage, window.location.hash);
			if (isNaN(serializerVersion)) {
				serializerVersion=OSF.OUtil.parseSerializerVersionFromWindowName(skipSessionStorage, window.name);
			}
			return serializerVersion;
		},
		parseSerializerVersionFromWindowName: function OSF_OUtil$parseSerializerVersionFromWindowName(skipSessionStorage, windowName) {
			return parseInt(OSF.OUtil.parseInfoFromWindowName(skipSessionStorage, windowName, OSF.WindowNameItemKeys.SerializerVersion));
		},
		parseSerializerVersionWithGivenFragment: function OSF_OUtil$parseSerializerVersionWithGivenFragment(skipSessionStorage, fragment) {
			return parseInt(OSF.OUtil.parseInfoWithGivenFragment(_serializerVersionKey, _serializerVersionKeyPrefix, true, skipSessionStorage, fragment));
		},
		parseInfoFromWindowName: function OSF_OUtil$parseInfoFromWindowName(skipSessionStorage, windowName, infoKey) {
			try {
				var windowNameObj=JSON.parse(windowName);
				var infoValue=windowNameObj !=null ? windowNameObj[infoKey] : null;
				var osfSessionStorage=_getSessionStorage();
				if (!skipSessionStorage && osfSessionStorage && windowNameObj !=null) {
					var sessionKey=windowNameObj[OSF.WindowNameItemKeys.BaseFrameName]+infoKey;
					if (infoValue) {
						osfSessionStorage.setItem(sessionKey, infoValue);
					}
					else {
						infoValue=osfSessionStorage.getItem(sessionKey);
					}
				}
				return infoValue;
			}
			catch (Exception) {
				return null;
			}
		},
		parseInfoWithGivenFragment: function OSF_OUtil$parseInfoWithGivenFragment(infoKey, infoKeyPrefix, decodeInfo, skipSessionStorage, fragment) {
			var fragmentParts=fragment.split(infoKey);
			var infoValue=fragmentParts.length > 1 ? fragmentParts[fragmentParts.length - 1] : null;
			if (decodeInfo && infoValue !=null) {
				if (infoValue.indexOf(_fragmentInfoDelimiter) >=0) {
					infoValue=infoValue.split(_fragmentInfoDelimiter)[0];
				}
				infoValue=decodeURIComponent(infoValue);
			}
			var osfSessionStorage=_getSessionStorage();
			if (!skipSessionStorage && osfSessionStorage) {
				var sessionKeyStart=window.name.indexOf(infoKeyPrefix);
				if (sessionKeyStart > -1) {
					var sessionKeyEnd=window.name.indexOf(";", sessionKeyStart);
					if (sessionKeyEnd==-1) {
						sessionKeyEnd=window.name.length;
					}
					var sessionKey=window.name.substring(sessionKeyStart, sessionKeyEnd);
					if (infoValue) {
						osfSessionStorage.setItem(sessionKey, infoValue);
					}
					else {
						infoValue=osfSessionStorage.getItem(sessionKey);
					}
				}
			}
			return infoValue;
		},
		getConversationId: function OSF_OUtil$getConversationId() {
			var searchString=window.location.search;
			var conversationId=null;
			if (searchString) {
				var index=searchString.indexOf("&");
				conversationId=index > 0 ? searchString.substring(1, index) : searchString.substr(1);
				if (conversationId && conversationId.charAt(conversationId.length - 1)==='=') {
					conversationId=conversationId.substring(0, conversationId.length - 1);
					if (conversationId) {
						conversationId=decodeURIComponent(conversationId);
					}
				}
			}
			return conversationId;
		},
		getInfoItems: function OSF_OUtil$getInfoItems(strInfo) {
			var items=strInfo.split("$");
			if (typeof items[1]=="undefined") {
				items=strInfo.split("|");
			}
			if (typeof items[1]=="undefined") {
				items=strInfo.split("%7C");
			}
			return items;
		},
		getXdmFieldValue: function OSF_OUtil$getXdmFieldValue(xdmFieldName, skipSessionStorage) {
			var fieldValue='';
			var xdmInfoValue=OSF.OUtil.parseXdmInfo(skipSessionStorage);
			if (xdmInfoValue) {
				var items=OSF.OUtil.getInfoItems(xdmInfoValue);
				if (items !=undefined && items.length >=3) {
					switch (xdmFieldName) {
						case OSF.XdmFieldName.ConversationUrl:
							fieldValue=items[2];
							break;
						case OSF.XdmFieldName.AppId:
							fieldValue=items[1];
							break;
					}
				}
			}
			return fieldValue;
		},
		validateParamObject: function OSF_OUtil$validateParamObject(params, expectedProperties, callback) {
			var e=Function._validateParams(arguments, [{ name: "params", type: Object, mayBeNull: false },
				{ name: "expectedProperties", type: Object, mayBeNull: false },
				{ name: "callback", type: Function, mayBeNull: true }
			]);
			if (e)
				throw e;
			for (var p in expectedProperties) {
				e=Function._validateParameter(params[p], expectedProperties[p], p);
				if (e)
					throw e;
			}
		},
		writeProfilerMark: function OSF_OUtil$writeProfilerMark(text) {
			if (window.msWriteProfilerMark) {
				window.msWriteProfilerMark(text);
				OsfMsAjaxFactory.msAjaxDebug.trace(text);
			}
		},
		outputDebug: function OSF_OUtil$outputDebug(text) {
			if (typeof (OsfMsAjaxFactory) !=='undefined' && OsfMsAjaxFactory.msAjaxDebug && OsfMsAjaxFactory.msAjaxDebug.trace) {
				OsfMsAjaxFactory.msAjaxDebug.trace(text);
			}
		},
		defineNondefaultProperty: function OSF_OUtil$defineNondefaultProperty(obj, prop, descriptor, attributes) {
			descriptor=descriptor || {};
			for (var nd in attributes) {
				var attribute=attributes[nd];
				if (descriptor[attribute]==undefined) {
					descriptor[attribute]=true;
				}
			}
			Object.defineProperty(obj, prop, descriptor);
			return obj;
		},
		defineNondefaultProperties: function OSF_OUtil$defineNondefaultProperties(obj, descriptors, attributes) {
			descriptors=descriptors || {};
			for (var prop in descriptors) {
				OSF.OUtil.defineNondefaultProperty(obj, prop, descriptors[prop], attributes);
			}
			return obj;
		},
		defineEnumerableProperty: function OSF_OUtil$defineEnumerableProperty(obj, prop, descriptor) {
			return OSF.OUtil.defineNondefaultProperty(obj, prop, descriptor, ["enumerable"]);
		},
		defineEnumerableProperties: function OSF_OUtil$defineEnumerableProperties(obj, descriptors) {
			return OSF.OUtil.defineNondefaultProperties(obj, descriptors, ["enumerable"]);
		},
		defineMutableProperty: function OSF_OUtil$defineMutableProperty(obj, prop, descriptor) {
			return OSF.OUtil.defineNondefaultProperty(obj, prop, descriptor, ["writable", "enumerable", "configurable"]);
		},
		defineMutableProperties: function OSF_OUtil$defineMutableProperties(obj, descriptors) {
			return OSF.OUtil.defineNondefaultProperties(obj, descriptors, ["writable", "enumerable", "configurable"]);
		},
		finalizeProperties: function OSF_OUtil$finalizeProperties(obj, descriptor) {
			descriptor=descriptor || {};
			var props=Object.getOwnPropertyNames(obj);
			var propsLength=props.length;
			for (var i=0; i < propsLength; i++) {
				var prop=props[i];
				var desc=Object.getOwnPropertyDescriptor(obj, prop);
				if (!desc.get && !desc.set) {
					desc.writable=descriptor.writable || false;
				}
				desc.configurable=descriptor.configurable || false;
				desc.enumerable=descriptor.enumerable || true;
				Object.defineProperty(obj, prop, desc);
			}
			return obj;
		},
		mapList: function OSF_OUtil$MapList(list, mapFunction) {
			var ret=[];
			if (list) {
				for (var item in list) {
					ret.push(mapFunction(list[item]));
				}
			}
			return ret;
		},
		listContainsKey: function OSF_OUtil$listContainsKey(list, key) {
			for (var item in list) {
				if (key==item) {
					return true;
				}
			}
			return false;
		},
		listContainsValue: function OSF_OUtil$listContainsElement(list, value) {
			for (var item in list) {
				if (value==list[item]) {
					return true;
				}
			}
			return false;
		},
		augmentList: function OSF_OUtil$augmentList(list, addenda) {
			var add=list.push ? function (key, value) { list.push(value); } : function (key, value) { list[key]=value; };
			for (var key in addenda) {
				add(key, addenda[key]);
			}
		},
		redefineList: function OSF_Outil$redefineList(oldList, newList) {
			for (var key1 in oldList) {
				delete oldList[key1];
			}
			for (var key2 in newList) {
				oldList[key2]=newList[key2];
			}
		},
		isArray: function OSF_OUtil$isArray(obj) {
			return Object.prototype.toString.apply(obj)==="[object Array]";
		},
		isFunction: function OSF_OUtil$isFunction(obj) {
			return Object.prototype.toString.apply(obj)==="[object Function]";
		},
		isDate: function OSF_OUtil$isDate(obj) {
			return Object.prototype.toString.apply(obj)==="[object Date]";
		},
		addEventListener: function OSF_OUtil$addEventListener(element, eventName, listener) {
			if (element.addEventListener) {
				element.addEventListener(eventName, listener, false);
			}
			else if ((Sys.Browser.agent===Sys.Browser.InternetExplorer) && element.attachEvent) {
				element.attachEvent("on"+eventName, listener);
			}
			else {
				element["on"+eventName]=listener;
			}
		},
		removeEventListener: function OSF_OUtil$removeEventListener(element, eventName, listener) {
			if (element.removeEventListener) {
				element.removeEventListener(eventName, listener, false);
			}
			else if ((Sys.Browser.agent===Sys.Browser.InternetExplorer) && element.detachEvent) {
				element.detachEvent("on"+eventName, listener);
			}
			else {
				element["on"+eventName]=null;
			}
		},
		getCookieValue: function OSF_OUtil$getCookieValue(cookieName) {
			var tmpCookieString=RegExp(cookieName+"[^;]+").exec(document.cookie);
			return tmpCookieString.toString().replace(/^[^=]+./, "");
		},
		xhrGet: function OSF_OUtil$xhrGet(url, onSuccess, onError) {
			var xmlhttp;
			try {
				xmlhttp=new XMLHttpRequest();
				xmlhttp.onreadystatechange=function () {
					if (xmlhttp.readyState==4) {
						if (xmlhttp.status==200) {
							onSuccess(xmlhttp.responseText);
						}
						else {
							onError(xmlhttp.status);
						}
					}
				};
				xmlhttp.open("GET", url, true);
				xmlhttp.send();
			}
			catch (ex) {
				onError(ex);
			}
		},
		xhrGetFull: function OSF_OUtil$xhrGetFull(url, oneDriveFileName, onSuccess, onError) {
			var xmlhttp;
			var requestedFileName=oneDriveFileName;
			try {
				xmlhttp=new XMLHttpRequest();
				xmlhttp.onreadystatechange=function () {
					if (xmlhttp.readyState==4) {
						if (xmlhttp.status==200) {
							onSuccess(xmlhttp, requestedFileName);
						}
						else {
							onError(xmlhttp.status);
						}
					}
				};
				xmlhttp.open("GET", url, true);
				xmlhttp.send();
			}
			catch (ex) {
				onError(ex);
			}
		},
		encodeBase64: function OSF_Outil$encodeBase64(input) {
			if (!input)
				return input;
			var codex="ABCDEFGHIJKLMNOP"+"QRSTUVWXYZabcdef"+"ghijklmnopqrstuv"+"wxyz0123456789+/=";
			var output=[];
			var temp=[];
			var index=0;
			var c1, c2, c3, a, b, c;
			var i;
			var length=input.length;
			do {
				c1=input.charCodeAt(index++);
				c2=input.charCodeAt(index++);
				c3=input.charCodeAt(index++);
				i=0;
				a=c1 & 255;
				b=c1 >> 8;
				c=c2 & 255;
				temp[i++]=a >> 2;
				temp[i++]=((a & 3) << 4) | (b >> 4);
				temp[i++]=((b & 15) << 2) | (c >> 6);
				temp[i++]=c & 63;
				if (!isNaN(c2)) {
					a=c2 >> 8;
					b=c3 & 255;
					c=c3 >> 8;
					temp[i++]=a >> 2;
					temp[i++]=((a & 3) << 4) | (b >> 4);
					temp[i++]=((b & 15) << 2) | (c >> 6);
					temp[i++]=c & 63;
				}
				if (isNaN(c2)) {
					temp[i - 1]=64;
				}
				else if (isNaN(c3)) {
					temp[i - 2]=64;
					temp[i - 1]=64;
				}
				for (var t=0; t < i; t++) {
					output.push(codex.charAt(temp[t]));
				}
			} while (index < length);
			return output.join("");
		},
		getSessionStorage: function OSF_Outil$getSessionStorage() {
			return _getSessionStorage();
		},
		getLocalStorage: function OSF_Outil$getLocalStorage() {
			if (!_safeLocalStorage) {
				try {
					var localStorage=window.localStorage;
				}
				catch (ex) {
					localStorage=null;
				}
				_safeLocalStorage=new OfficeExt.SafeStorage(localStorage);
			}
			return _safeLocalStorage;
		},
		convertIntToCssHexColor: function OSF_Outil$convertIntToCssHexColor(val) {
			var hex="#"+(Number(val)+0x1000000).toString(16).slice(-6);
			return hex;
		},
		attachClickHandler: function OSF_Outil$attachClickHandler(element, handler) {
			element.onclick=function (e) {
				handler();
			};
			element.ontouchend=function (e) {
				handler();
				e.preventDefault();
			};
		},
		getQueryStringParamValue: function OSF_Outil$getQueryStringParamValue(queryString, paramName) {
			var e=Function._validateParams(arguments, [{ name: "queryString", type: String, mayBeNull: false },
				{ name: "paramName", type: String, mayBeNull: false }
			]);
			if (e) {
				OsfMsAjaxFactory.msAjaxDebug.trace("OSF_Outil_getQueryStringParamValue: Parameters cannot be null.");
				return "";
			}
			var queryExp=new RegExp("[\\?&]"+paramName+"=([^&#]*)", "i");
			if (!queryExp.test(queryString)) {
				OsfMsAjaxFactory.msAjaxDebug.trace("OSF_Outil_getQueryStringParamValue: The parameter is not found.");
				return "";
			}
			return queryExp.exec(queryString)[1];
		},
		isiOS: function OSF_Outil$isiOS() {
			return (window.navigator.userAgent.match(/(iPad|iPhone|iPod)/g) ? true : false);
		},
		isChrome: function OSF_Outil$isChrome() {
			return (window.navigator.userAgent.indexOf("Chrome") > 0) && !OSF.OUtil.isEdge();
		},
		isEdge: function OSF_Outil$isEdge() {
			return window.navigator.userAgent.indexOf("Edge") > 0;
		},
		isIE: function OSF_Outil$isIE() {
			return window.navigator.userAgent.indexOf("Trident") > 0;
		},
		isFirefox: function OSF_Outil$isFirefox() {
			return window.navigator.userAgent.indexOf("Firefox") > 0;
		},
		shallowCopy: function OSF_Outil$shallowCopy(sourceObj) {
			if (sourceObj==null) {
				return null;
			}
			else if (!(sourceObj instanceof Object)) {
				return sourceObj;
			}
			else if (Array.isArray(sourceObj)) {
				var copyArr=[];
				for (var i=0; i < sourceObj.length; i++) {
					copyArr.push(sourceObj[i]);
				}
				return copyArr;
			}
			else {
				var copyObj=sourceObj.constructor();
				for (var property in sourceObj) {
					if (sourceObj.hasOwnProperty(property)) {
						copyObj[property]=sourceObj[property];
					}
				}
				return copyObj;
			}
		},
		createObject: function OSF_Outil$createObject(properties) {
			var obj=null;
			if (properties) {
				obj={};
				var len=properties.length;
				for (var i=0; i < len; i++) {
					obj[properties[i].name]=properties[i].value;
				}
			}
			return obj;
		},
		addClass: function OSF_OUtil$addClass(elmt, val) {
			if (!OSF.OUtil.hasClass(elmt, val)) {
				var className=elmt.getAttribute(_classN);
				if (className) {
					elmt.setAttribute(_classN, className+" "+val);
				}
				else {
					elmt.setAttribute(_classN, val);
				}
			}
		},
		removeClass: function OSF_OUtil$removeClass(elmt, val) {
			if (OSF.OUtil.hasClass(elmt, val)) {
				var className=elmt.getAttribute(_classN);
				var reg=new RegExp('(\\s|^)'+val+'(\\s|$)');
				className=className.replace(reg, '');
				elmt.setAttribute(_classN, className);
			}
		},
		hasClass: function OSF_OUtil$hasClass(elmt, clsName) {
			var className=elmt.getAttribute(_classN);
			return className && className.match(new RegExp('(\\s|^)'+clsName+'(\\s|$)'));
		},
		focusToFirstTabbable: function OSF_OUtil$focusToFirstTabbable(all, backward) {
			var next;
			var focused=false;
			var candidate;
			var setFlag=function (e) {
				focused=true;
			};
			var findNextPos=function (allLen, currPos, backward) {
				if (currPos < 0 || currPos > allLen) {
					return -1;
				}
				else if (currPos===0 && backward) {
					return -1;
				}
				else if (currPos===allLen - 1 && !backward) {
					return -1;
				}
				if (backward) {
					return currPos - 1;
				}
				else {
					return currPos+1;
				}
			};
			all=_reOrderTabbableElements(all);
			next=backward ? all.length - 1 : 0;
			if (all.length===0) {
				return null;
			}
			while (!focused && next >=0 && next < all.length) {
				candidate=all[next];
				window.focus();
				candidate.addEventListener('focus', setFlag);
				candidate.focus();
				candidate.removeEventListener('focus', setFlag);
				next=findNextPos(all.length, next, backward);
				if (!focused && candidate===document.activeElement) {
					focused=true;
				}
			}
			if (focused) {
				return candidate;
			}
			else {
				return null;
			}
		},
		focusToNextTabbable: function OSF_OUtil$focusToNextTabbable(all, curr, shift) {
			var currPos;
			var next;
			var focused=false;
			var candidate;
			var setFlag=function (e) {
				focused=true;
			};
			var findCurrPos=function (all, curr) {
				var i=0;
				for (; i < all.length; i++) {
					if (all[i]===curr) {
						return i;
					}
				}
				return -1;
			};
			var findNextPos=function (allLen, currPos, shift) {
				if (currPos < 0 || currPos > allLen) {
					return -1;
				}
				else if (currPos===0 && shift) {
					return -1;
				}
				else if (currPos===allLen - 1 && !shift) {
					return -1;
				}
				if (shift) {
					return currPos - 1;
				}
				else {
					return currPos+1;
				}
			};
			all=_reOrderTabbableElements(all);
			currPos=findCurrPos(all, curr);
			next=findNextPos(all.length, currPos, shift);
			if (next < 0) {
				return null;
			}
			while (!focused && next >=0 && next < all.length) {
				candidate=all[next];
				candidate.addEventListener('focus', setFlag);
				candidate.focus();
				candidate.removeEventListener('focus', setFlag);
				next=findNextPos(all.length, next, shift);
				if (!focused && candidate===document.activeElement) {
					focused=true;
				}
			}
			if (focused) {
				return candidate;
			}
			else {
				return null;
			}
		}
	};
})();
OSF.OUtil.Guid=(function () {
	var hexCode=["0", "1", "2", "3", "4", "5", "6", "7", "8", "9", "a", "b", "c", "d", "e", "f"];
	return {
		generateNewGuid: function OSF_Outil_Guid$generateNewGuid() {
			var result="";
			var tick=(new Date()).getTime();
			var index=0;
			for (; index < 32 && tick > 0; index++) {
				if (index==8 || index==12 || index==16 || index==20) {
					result+="-";
				}
				result+=hexCode[tick % 16];
				tick=Math.floor(tick / 16);
			}
			for (; index < 32; index++) {
				if (index==8 || index==12 || index==16 || index==20) {
					result+="-";
				}
				result+=hexCode[Math.floor(Math.random() * 16)];
			}
			return result;
		}
	};
})();
window.OSF=OSF;
OSF.OUtil.setNamespace("OSF", window);
OSF.AppName={
	Unsupported: 0,
	Excel: 1,
	Word: 2,
	PowerPoint: 4,
	Outlook: 8,
	ExcelWebApp: 16,
	WordWebApp: 32,
	OutlookWebApp: 64,
	Project: 128,
	AccessWebApp: 256,
	PowerpointWebApp: 512,
	ExcelIOS: 1024,
	Sway: 2048,
	WordIOS: 4096,
	PowerPointIOS: 8192,
	Access: 16384,
	Lync: 32768,
	OutlookIOS: 65536,
	OneNoteWebApp: 131072,
	OneNote: 262144,
	ExcelWinRT: 524288,
	WordWinRT: 1048576,
	PowerpointWinRT: 2097152,
	OutlookAndroid: 4194304,
	OneNoteWinRT: 8388608,
	ExcelAndroid: 8388609,
	VisioWebApp: 8388610,
	OneNoteIOS: 8388611
};
OSF.InternalPerfMarker={
	DataCoercionBegin: "Agave.HostCall.CoerceDataStart",
	DataCoercionEnd: "Agave.HostCall.CoerceDataEnd"
};
OSF.HostCallPerfMarker={
	IssueCall: "Agave.HostCall.IssueCall",
	ReceiveResponse: "Agave.HostCall.ReceiveResponse",
	RuntimeExceptionRaised: "Agave.HostCall.RuntimeExecptionRaised"
};
OSF.AgaveHostAction={
	"Select": 0,
	"UnSelect": 1,
	"CancelDialog": 2,
	"InsertAgave": 3,
	"CtrlF6In": 4,
	"CtrlF6Exit": 5,
	"CtrlF6ExitShift": 6,
	"SelectWithError": 7,
	"NotifyHostError": 8,
	"RefreshAddinCommands": 9,
	"PageIsReady": 10,
	"TabIn": 11,
	"TabInShift": 12,
	"TabExit": 13,
	"TabExitShift": 14,
	"EscExit": 15,
	"F2Exit": 16,
	"ExitNoFocusable": 17,
	"ExitNoFocusableShift": 18,
	"MouseEnter": 19,
	"MouseLeave": 20
};
OSF.SharedConstants={
	"NotificationConversationIdSuffix": '_ntf'
};
OSF.DialogMessageType={
	DialogMessageReceived: 0,
	DialogParentMessageReceived: 1,
	DialogClosed: 12006
};
OSF.OfficeAppContext=function OSF_OfficeAppContext(id, appName, appVersion, appUILocale, dataLocale, docUrl, clientMode, settings, reason, osfControlType, eToken, correlationId, appInstanceId, touchEnabled, commerceAllowed, appMinorVersion, requirementMatrix, hostCustomMessage, hostFullVersion, clientWindowHeight, clientWindowWidth, addinName, appDomains, dialogRequirementMatrix) {
	this._id=id;
	this._appName=appName;
	this._appVersion=appVersion;
	this._appUILocale=appUILocale;
	this._dataLocale=dataLocale;
	this._docUrl=docUrl;
	this._clientMode=clientMode;
	this._settings=settings;
	this._reason=reason;
	this._osfControlType=osfControlType;
	this._eToken=eToken;
	this._correlationId=correlationId;
	this._appInstanceId=appInstanceId;
	this._touchEnabled=touchEnabled;
	this._commerceAllowed=commerceAllowed;
	this._appMinorVersion=appMinorVersion;
	this._requirementMatrix=requirementMatrix;
	this._hostCustomMessage=hostCustomMessage;
	this._hostFullVersion=hostFullVersion;
	this._isDialog=false;
	this._clientWindowHeight=clientWindowHeight;
	this._clientWindowWidth=clientWindowWidth;
	this._addinName=addinName;
	this._appDomains=appDomains;
	this._dialogRequirementMatrix=dialogRequirementMatrix;
	this.get_id=function get_id() { return this._id; };
	this.get_appName=function get_appName() { return this._appName; };
	this.get_appVersion=function get_appVersion() { return this._appVersion; };
	this.get_appUILocale=function get_appUILocale() { return this._appUILocale; };
	this.get_dataLocale=function get_dataLocale() { return this._dataLocale; };
	this.get_docUrl=function get_docUrl() { return this._docUrl; };
	this.get_clientMode=function get_clientMode() { return this._clientMode; };
	this.get_bindings=function get_bindings() { return this._bindings; };
	this.get_settings=function get_settings() { return this._settings; };
	this.get_reason=function get_reason() { return this._reason; };
	this.get_osfControlType=function get_osfControlType() { return this._osfControlType; };
	this.get_eToken=function get_eToken() { return this._eToken; };
	this.get_correlationId=function get_correlationId() { return this._correlationId; };
	this.get_appInstanceId=function get_appInstanceId() { return this._appInstanceId; };
	this.get_touchEnabled=function get_touchEnabled() { return this._touchEnabled; };
	this.get_commerceAllowed=function get_commerceAllowed() { return this._commerceAllowed; };
	this.get_appMinorVersion=function get_appMinorVersion() { return this._appMinorVersion; };
	this.get_requirementMatrix=function get_requirementMatrix() { return this._requirementMatrix; };
	this.get_dialogRequirementMatrix=function get_dialogRequirementMatrix() { return this._dialogRequirementMatrix; };
	this.get_hostCustomMessage=function get_hostCustomMessage() { return this._hostCustomMessage; };
	this.get_hostFullVersion=function get_hostFullVersion() { return this._hostFullVersion; };
	this.get_isDialog=function get_isDialog() { return this._isDialog; };
	this.get_clientWindowHeight=function get_clientWindowHeight() { return this._clientWindowHeight; };
	this.get_clientWindowWidth=function get_clientWindowWidth() { return this._clientWindowWidth; };
	this.get_addinName=function get_addinName() { return this._addinName; };
	this.get_appDomains=function get_appDomains() { return this._appDomains; };
};
OSF.OsfControlType={
	DocumentLevel: 0,
	ContainerLevel: 1
};
OSF.ClientMode={
	ReadOnly: 0,
	ReadWrite: 1
};
OSF.OUtil.setNamespace("Microsoft", window);
OSF.OUtil.setNamespace("Office", Microsoft);
OSF.OUtil.setNamespace("Client", Microsoft.Office);
OSF.OUtil.setNamespace("WebExtension", Microsoft.Office);
Microsoft.Office.WebExtension.InitializationReason={
	Inserted: "inserted",
	DocumentOpened: "documentOpened"
};
Microsoft.Office.WebExtension.ValueFormat={
	Unformatted: "unformatted",
	Formatted: "formatted"
};
Microsoft.Office.WebExtension.FilterType={
	All: "all"
};
Microsoft.Office.WebExtension.PlatformType={
	PC: "PC",
	OfficeOnline: "OfficeOnline",
	Mac: "Mac",
	iOS: "iOS",
	Android: "Android",
	Universal: "Universal"
};
Microsoft.Office.WebExtension.HostType={
	Word: "Word",
	Excel: "Excel",
	PowerPoint: "PowerPoint",
	Outlook: "Outlook",
	OneNote: "OneNote",
	Project: "Project",
	Access: "Access"
};
Microsoft.Office.WebExtension.Parameters={
	BindingType: "bindingType",
	CoercionType: "coercionType",
	ValueFormat: "valueFormat",
	FilterType: "filterType",
	Columns: "columns",
	SampleData: "sampleData",
	GoToType: "goToType",
	SelectionMode: "selectionMode",
	Id: "id",
	PromptText: "promptText",
	ItemName: "itemName",
	FailOnCollision: "failOnCollision",
	StartRow: "startRow",
	StartColumn: "startColumn",
	RowCount: "rowCount",
	ColumnCount: "columnCount",
	Callback: "callback",
	AsyncContext: "asyncContext",
	Data: "data",
	Rows: "rows",
	OverwriteIfStale: "overwriteIfStale",
	FileType: "fileType",
	EventType: "eventType",
	Handler: "handler",
	SliceSize: "sliceSize",
	SliceIndex: "sliceIndex",
	ActiveView: "activeView",
	Status: "status",
	PlatformType: "platformType",
	HostType: "hostType",
	ForceConsent: "forceConsent",
	ForceAddAccount: "forceAddAccount",
	AuthChallenge: "authChallenge",
	Xml: "xml",
	Namespace: "namespace",
	Prefix: "prefix",
	XPath: "xPath",
	Text: "text",
	ImageLeft: "imageLeft",
	ImageTop: "imageTop",
	ImageWidth: "imageWidth",
	ImageHeight: "imageHeight",
	TaskId: "taskId",
	FieldId: "fieldId",
	FieldValue: "fieldValue",
	ServerUrl: "serverUrl",
	ListName: "listName",
	ResourceId: "resourceId",
	ViewType: "viewType",
	ViewName: "viewName",
	GetRawValue: "getRawValue",
	CellFormat: "cellFormat",
	TableOptions: "tableOptions",
	TaskIndex: "taskIndex",
	ResourceIndex: "resourceIndex",
	CustomFieldId: "customFieldId",
	Url: "url",
	MessageHandler: "messageHandler",
	Width: "width",
	Height: "height",
	RequireHTTPs: "requireHTTPS",
	MessageToParent: "messageToParent",
	DisplayInIframe: "displayInIframe",
	MessageContent: "messageContent",
	HideTitle: "hideTitle",
	AppCommandInvocationCompletedData: "appCommandInvocationCompletedData"
};
OSF.OUtil.setNamespace("DDA", OSF);
OSF.DDA.DocumentMode={
	ReadOnly: 1,
	ReadWrite: 0
};
OSF.DDA.PropertyDescriptors={
	AsyncResultStatus: "AsyncResultStatus"
};
OSF.DDA.EventDescriptors={};
OSF.DDA.ListDescriptors={};
OSF.DDA.UI={};
OSF.DDA.getXdmEventName=function OSF_DDA$GetXdmEventName(id, eventType) {
	if (eventType==Microsoft.Office.WebExtension.EventType.BindingSelectionChanged ||
		eventType==Microsoft.Office.WebExtension.EventType.BindingDataChanged ||
		eventType==Microsoft.Office.WebExtension.EventType.DataNodeDeleted ||
		eventType==Microsoft.Office.WebExtension.EventType.DataNodeInserted ||
		eventType==Microsoft.Office.WebExtension.EventType.DataNodeReplaced) {
		return id+"_"+eventType;
	}
	else {
		return eventType;
	}
};
OSF.DDA.MethodDispId={
	dispidMethodMin: 64,
	dispidGetSelectedDataMethod: 64,
	dispidSetSelectedDataMethod: 65,
	dispidAddBindingFromSelectionMethod: 66,
	dispidAddBindingFromPromptMethod: 67,
	dispidGetBindingMethod: 68,
	dispidReleaseBindingMethod: 69,
	dispidGetBindingDataMethod: 70,
	dispidSetBindingDataMethod: 71,
	dispidAddRowsMethod: 72,
	dispidClearAllRowsMethod: 73,
	dispidGetAllBindingsMethod: 74,
	dispidLoadSettingsMethod: 75,
	dispidSaveSettingsMethod: 76,
	dispidGetDocumentCopyMethod: 77,
	dispidAddBindingFromNamedItemMethod: 78,
	dispidAddColumnsMethod: 79,
	dispidGetDocumentCopyChunkMethod: 80,
	dispidReleaseDocumentCopyMethod: 81,
	dispidNavigateToMethod: 82,
	dispidGetActiveViewMethod: 83,
	dispidGetDocumentThemeMethod: 84,
	dispidGetOfficeThemeMethod: 85,
	dispidGetFilePropertiesMethod: 86,
	dispidClearFormatsMethod: 87,
	dispidSetTableOptionsMethod: 88,
	dispidSetFormatsMethod: 89,
	dispidExecuteRichApiRequestMethod: 93,
	dispidAppCommandInvocationCompletedMethod: 94,
	dispidCloseContainerMethod: 97,
	dispidGetAccessTokenMethod: 98,
	dispidGetSelectedTaskMethod: 110,
	dispidGetSelectedResourceMethod: 111,
	dispidGetTaskMethod: 112,
	dispidGetResourceFieldMethod: 113,
	dispidGetWSSUrlMethod: 114,
	dispidGetTaskFieldMethod: 115,
	dispidGetProjectFieldMethod: 116,
	dispidGetSelectedViewMethod: 117,
	dispidGetTaskByIndexMethod: 118,
	dispidGetResourceByIndexMethod: 119,
	dispidSetTaskFieldMethod: 120,
	dispidSetResourceFieldMethod: 121,
	dispidGetMaxTaskIndexMethod: 122,
	dispidGetMaxResourceIndexMethod: 123,
	dispidCreateTaskMethod: 124,
	dispidAddDataPartMethod: 128,
	dispidGetDataPartByIdMethod: 129,
	dispidGetDataPartsByNamespaceMethod: 130,
	dispidGetDataPartXmlMethod: 131,
	dispidGetDataPartNodesMethod: 132,
	dispidDeleteDataPartMethod: 133,
	dispidGetDataNodeValueMethod: 134,
	dispidGetDataNodeXmlMethod: 135,
	dispidGetDataNodesMethod: 136,
	dispidSetDataNodeValueMethod: 137,
	dispidSetDataNodeXmlMethod: 138,
	dispidAddDataNamespaceMethod: 139,
	dispidGetDataUriByPrefixMethod: 140,
	dispidGetDataPrefixByUriMethod: 141,
	dispidGetDataNodeTextMethod: 142,
	dispidSetDataNodeTextMethod: 143,
	dispidMessageParentMethod: 144,
	dispidSendMessageMethod: 145,
	dispidMethodMax: 145
};
OSF.DDA.EventDispId={
	dispidEventMin: 0,
	dispidInitializeEvent: 0,
	dispidSettingsChangedEvent: 1,
	dispidDocumentSelectionChangedEvent: 2,
	dispidBindingSelectionChangedEvent: 3,
	dispidBindingDataChangedEvent: 4,
	dispidDocumentOpenEvent: 5,
	dispidDocumentCloseEvent: 6,
	dispidActiveViewChangedEvent: 7,
	dispidDocumentThemeChangedEvent: 8,
	dispidOfficeThemeChangedEvent: 9,
	dispidDialogMessageReceivedEvent: 10,
	dispidDialogNotificationShownInAddinEvent: 11,
	dispidDialogParentMessageReceivedEvent: 12,
	dispidObjectDeletedEvent: 13,
	dispidObjectSelectionChangedEvent: 14,
	dispidObjectDataChangedEvent: 15,
	dispidContentControlAddedEvent: 16,
	dispidActivationStatusChangedEvent: 32,
	dispidRichApiMessageEvent: 33,
	dispidAppCommandInvokedEvent: 39,
	dispidOlkItemSelectedChangedEvent: 46,
	dispidOlkRecipientsChangedEvent: 47,
	dispidOlkAppointmentTimeChangedEvent: 48,
	dispidTaskSelectionChangedEvent: 56,
	dispidResourceSelectionChangedEvent: 57,
	dispidViewSelectionChangedEvent: 58,
	dispidDataNodeAddedEvent: 60,
	dispidDataNodeReplacedEvent: 61,
	dispidDataNodeDeletedEvent: 62,
	dispidEventMax: 63
};
OSF.DDA.ErrorCodeManager=(function () {
	var _errorMappings={};
	return {
		getErrorArgs: function OSF_DDA_ErrorCodeManager$getErrorArgs(errorCode) {
			var errorArgs=_errorMappings[errorCode];
			if (!errorArgs) {
				errorArgs=_errorMappings[this.errorCodes.ooeInternalError];
			}
			else {
				if (!errorArgs.name) {
					errorArgs.name=_errorMappings[this.errorCodes.ooeInternalError].name;
				}
				if (!errorArgs.message) {
					errorArgs.message=_errorMappings[this.errorCodes.ooeInternalError].message;
				}
			}
			return errorArgs;
		},
		addErrorMessage: function OSF_DDA_ErrorCodeManager$addErrorMessage(errorCode, errorNameMessage) {
			_errorMappings[errorCode]=errorNameMessage;
		},
		errorCodes: {
			ooeSuccess: 0,
			ooeChunkResult: 1,
			ooeCoercionTypeNotSupported: 1000,
			ooeGetSelectionNotMatchDataType: 1001,
			ooeCoercionTypeNotMatchBinding: 1002,
			ooeInvalidGetRowColumnCounts: 1003,
			ooeSelectionNotSupportCoercionType: 1004,
			ooeInvalidGetStartRowColumn: 1005,
			ooeNonUniformPartialGetNotSupported: 1006,
			ooeGetDataIsTooLarge: 1008,
			ooeFileTypeNotSupported: 1009,
			ooeGetDataParametersConflict: 1010,
			ooeInvalidGetColumns: 1011,
			ooeInvalidGetRows: 1012,
			ooeInvalidReadForBlankRow: 1013,
			ooeUnsupportedDataObject: 2000,
			ooeCannotWriteToSelection: 2001,
			ooeDataNotMatchSelection: 2002,
			ooeOverwriteWorksheetData: 2003,
			ooeDataNotMatchBindingSize: 2004,
			ooeInvalidSetStartRowColumn: 2005,
			ooeInvalidDataFormat: 2006,
			ooeDataNotMatchCoercionType: 2007,
			ooeDataNotMatchBindingType: 2008,
			ooeSetDataIsTooLarge: 2009,
			ooeNonUniformPartialSetNotSupported: 2010,
			ooeInvalidSetColumns: 2011,
			ooeInvalidSetRows: 2012,
			ooeSetDataParametersConflict: 2013,
			ooeCellDataAmountBeyondLimits: 2014,
			ooeSelectionCannotBound: 3000,
			ooeBindingNotExist: 3002,
			ooeBindingToMultipleSelection: 3003,
			ooeInvalidSelectionForBindingType: 3004,
			ooeOperationNotSupportedOnThisBindingType: 3005,
			ooeNamedItemNotFound: 3006,
			ooeMultipleNamedItemFound: 3007,
			ooeInvalidNamedItemForBindingType: 3008,
			ooeUnknownBindingType: 3009,
			ooeOperationNotSupportedOnMatrixData: 3010,
			ooeInvalidColumnsForBinding: 3011,
			ooeSettingNameNotExist: 4000,
			ooeSettingsCannotSave: 4001,
			ooeSettingsAreStale: 4002,
			ooeOperationNotSupported: 5000,
			ooeInternalError: 5001,
			ooeDocumentReadOnly: 5002,
			ooeEventHandlerNotExist: 5003,
			ooeInvalidApiCallInContext: 5004,
			ooeShuttingDown: 5005,
			ooeUnsupportedEnumeration: 5007,
			ooeIndexOutOfRange: 5008,
			ooeBrowserAPINotSupported: 5009,
			ooeInvalidParam: 5010,
			ooeRequestTimeout: 5011,
			ooeInvalidOrTimedOutSession: 5012,
			ooeInvalidApiArguments: 5013,
			ooeTooManyIncompleteRequests: 5100,
			ooeRequestTokenUnavailable: 5101,
			ooeActivityLimitReached: 5102,
			ooeCustomXmlNodeNotFound: 6000,
			ooeCustomXmlError: 6100,
			ooeCustomXmlExceedQuota: 6101,
			ooeCustomXmlOutOfDate: 6102,
			ooeNoCapability: 7000,
			ooeCannotNavTo: 7001,
			ooeSpecifiedIdNotExist: 7002,
			ooeNavOutOfBound: 7004,
			ooeElementMissing: 8000,
			ooeProtectedError: 8001,
			ooeInvalidCellsValue: 8010,
			ooeInvalidTableOptionValue: 8011,
			ooeInvalidFormatValue: 8012,
			ooeRowIndexOutOfRange: 8020,
			ooeColIndexOutOfRange: 8021,
			ooeFormatValueOutOfRange: 8022,
			ooeCellFormatAmountBeyondLimits: 8023,
			ooeMemoryFileLimit: 11000,
			ooeNetworkProblemRetrieveFile: 11001,
			ooeInvalidSliceSize: 11002,
			ooeInvalidCallback: 11101,
			ooeInvalidWidth: 12000,
			ooeInvalidHeight: 12001,
			ooeNavigationError: 12002,
			ooeInvalidScheme: 12003,
			ooeAppDomains: 12004,
			ooeRequireHTTPS: 12005,
			ooeWebDialogClosed: 12006,
			ooeDialogAlreadyOpened: 12007,
			ooeEndUserAllow: 12008,
			ooeEndUserIgnore: 12009,
			ooeNotUILessDialog: 12010,
			ooeCrossZone: 12011,
			ooeNotSSOAgave: 13000,
			ooeSSOUserNotSignedIn: 13001,
			ooeSSOUserAborted: 13002,
			ooeSSOUnsupportedUserIdentity: 13003,
			ooeSSOInvalidResourceUrl: 13004,
			ooeSSOInvalidGrant: 13005,
			ooeSSOClientError: 13006,
			ooeSSOServerError: 13007,
			ooeAddinIsAlreadyRequestingToken: 13008,
			ooeSSOUserConsentNotSupportedByCurrentAddinCategory: 13009
		},
		initializeErrorMessages: function OSF_DDA_ErrorCodeManager$initializeErrorMessages(stringNS) {
			_errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeCoercionTypeNotSupported]={ name: stringNS.L_InvalidCoercion, message: stringNS.L_CoercionTypeNotSupported };
			_errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeGetSelectionNotMatchDataType]={ name: stringNS.L_DataReadError, message: stringNS.L_GetSelectionNotSupported };
			_errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeCoercionTypeNotMatchBinding]={ name: stringNS.L_InvalidCoercion, message: stringNS.L_CoercionTypeNotMatchBinding };
			_errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeInvalidGetRowColumnCounts]={ name: stringNS.L_DataReadError, message: stringNS.L_InvalidGetRowColumnCounts };
			_errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeSelectionNotSupportCoercionType]={ name: stringNS.L_DataReadError, message: stringNS.L_SelectionNotSupportCoercionType };
			_errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeInvalidGetStartRowColumn]={ name: stringNS.L_DataReadError, message: stringNS.L_InvalidGetStartRowColumn };
			_errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeNonUniformPartialGetNotSupported]={ name: stringNS.L_DataReadError, message: stringNS.L_NonUniformPartialGetNotSupported };
			_errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeGetDataIsTooLarge]={ name: stringNS.L_DataReadError, message: stringNS.L_GetDataIsTooLarge };
			_errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeFileTypeNotSupported]={ name: stringNS.L_DataReadError, message: stringNS.L_FileTypeNotSupported };
			_errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeGetDataParametersConflict]={ name: stringNS.L_DataReadError, message: stringNS.L_GetDataParametersConflict };
			_errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeInvalidGetColumns]={ name: stringNS.L_DataReadError, message: stringNS.L_InvalidGetColumns };
			_errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeInvalidGetRows]={ name: stringNS.L_DataReadError, message: stringNS.L_InvalidGetRows };
			_errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeInvalidReadForBlankRow]={ name: stringNS.L_DataReadError, message: stringNS.L_InvalidReadForBlankRow };
			_errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeUnsupportedDataObject]={ name: stringNS.L_DataWriteError, message: stringNS.L_UnsupportedDataObject };
			_errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeCannotWriteToSelection]={ name: stringNS.L_DataWriteError, message: stringNS.L_CannotWriteToSelection };
			_errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeDataNotMatchSelection]={ name: stringNS.L_DataWriteError, message: stringNS.L_DataNotMatchSelection };
			_errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeOverwriteWorksheetData]={ name: stringNS.L_DataWriteError, message: stringNS.L_OverwriteWorksheetData };
			_errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeDataNotMatchBindingSize]={ name: stringNS.L_DataWriteError, message: stringNS.L_DataNotMatchBindingSize };
			_errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeInvalidSetStartRowColumn]={ name: stringNS.L_DataWriteError, message: stringNS.L_InvalidSetStartRowColumn };
			_errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeInvalidDataFormat]={ name: stringNS.L_InvalidFormat, message: stringNS.L_InvalidDataFormat };
			_errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeDataNotMatchCoercionType]={ name: stringNS.L_InvalidDataObject, message: stringNS.L_DataNotMatchCoercionType };
			_errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeDataNotMatchBindingType]={ name: stringNS.L_InvalidDataObject, message: stringNS.L_DataNotMatchBindingType };
			_errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeSetDataIsTooLarge]={ name: stringNS.L_DataWriteError, message: stringNS.L_SetDataIsTooLarge };
			_errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeNonUniformPartialSetNotSupported]={ name: stringNS.L_DataWriteError, message: stringNS.L_NonUniformPartialSetNotSupported };
			_errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeInvalidSetColumns]={ name: stringNS.L_DataWriteError, message: stringNS.L_InvalidSetColumns };
			_errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeInvalidSetRows]={ name: stringNS.L_DataWriteError, message: stringNS.L_InvalidSetRows };
			_errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeSetDataParametersConflict]={ name: stringNS.L_DataWriteError, message: stringNS.L_SetDataParametersConflict };
			_errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeSelectionCannotBound]={ name: stringNS.L_BindingCreationError, message: stringNS.L_SelectionCannotBound };
			_errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeBindingNotExist]={ name: stringNS.L_InvalidBindingError, message: stringNS.L_BindingNotExist };
			_errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeBindingToMultipleSelection]={ name: stringNS.L_BindingCreationError, message: stringNS.L_BindingToMultipleSelection };
			_errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeInvalidSelectionForBindingType]={ name: stringNS.L_BindingCreationError, message: stringNS.L_InvalidSelectionForBindingType };
			_errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeOperationNotSupportedOnThisBindingType]={ name: stringNS.L_InvalidBindingOperation, message: stringNS.L_OperationNotSupportedOnThisBindingType };
			_errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeNamedItemNotFound]={ name: stringNS.L_BindingCreationError, message: stringNS.L_NamedItemNotFound };
			_errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeMultipleNamedItemFound]={ name: stringNS.L_BindingCreationError, message: stringNS.L_MultipleNamedItemFound };
			_errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeInvalidNamedItemForBindingType]={ name: stringNS.L_BindingCreationError, message: stringNS.L_InvalidNamedItemForBindingType };
			_errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeUnknownBindingType]={ name: stringNS.L_InvalidBinding, message: stringNS.L_UnknownBindingType };
			_errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeOperationNotSupportedOnMatrixData]={ name: stringNS.L_InvalidBindingOperation, message: stringNS.L_OperationNotSupportedOnMatrixData };
			_errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeInvalidColumnsForBinding]={ name: stringNS.L_InvalidBinding, message: stringNS.L_InvalidColumnsForBinding };
			_errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeSettingNameNotExist]={ name: stringNS.L_ReadSettingsError, message: stringNS.L_SettingNameNotExist };
			_errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeSettingsCannotSave]={ name: stringNS.L_SaveSettingsError, message: stringNS.L_SettingsCannotSave };
			_errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeSettingsAreStale]={ name: stringNS.L_SettingsStaleError, message: stringNS.L_SettingsAreStale };
			_errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeOperationNotSupported]={ name: stringNS.L_HostError, message: stringNS.L_OperationNotSupported };
			_errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeInternalError]={ name: stringNS.L_InternalError, message: stringNS.L_InternalErrorDescription };
			_errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeDocumentReadOnly]={ name: stringNS.L_PermissionDenied, message: stringNS.L_DocumentReadOnly };
			_errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeEventHandlerNotExist]={ name: stringNS.L_EventRegistrationError, message: stringNS.L_EventHandlerNotExist };
			_errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeInvalidApiCallInContext]={ name: stringNS.L_InvalidAPICall, message: stringNS.L_InvalidApiCallInContext };
			_errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeShuttingDown]={ name: stringNS.L_ShuttingDown, message: stringNS.L_ShuttingDown };
			_errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeUnsupportedEnumeration]={ name: stringNS.L_UnsupportedEnumeration, message: stringNS.L_UnsupportedEnumerationMessage };
			_errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeIndexOutOfRange]={ name: stringNS.L_IndexOutOfRange, message: stringNS.L_IndexOutOfRange };
			_errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeBrowserAPINotSupported]={ name: stringNS.L_APINotSupported, message: stringNS.L_BrowserAPINotSupported };
			_errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeRequestTimeout]={ name: stringNS.L_APICallFailed, message: stringNS.L_RequestTimeout };
			_errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeInvalidOrTimedOutSession]={ name: stringNS.L_InvalidOrTimedOutSession, message: stringNS.L_InvalidOrTimedOutSessionMessage };
			_errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeTooManyIncompleteRequests]={ name: stringNS.L_APICallFailed, message: stringNS.L_TooManyIncompleteRequests };
			_errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeRequestTokenUnavailable]={ name: stringNS.L_APICallFailed, message: stringNS.L_RequestTokenUnavailable };
			_errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeActivityLimitReached]={ name: stringNS.L_APICallFailed, message: stringNS.L_ActivityLimitReached };
			_errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeInvalidApiArguments]={ name: stringNS.L_APICallFailed, message: stringNS.L_InvalidApiArgumentsMessage };
			_errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeCustomXmlNodeNotFound]={ name: stringNS.L_InvalidNode, message: stringNS.L_CustomXmlNodeNotFound };
			_errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeCustomXmlError]={ name: stringNS.L_CustomXmlError, message: stringNS.L_CustomXmlError };
			_errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeCustomXmlExceedQuota]={ name: stringNS.L_CustomXmlExceedQuotaName, message: stringNS.L_CustomXmlExceedQuotaMessage };
			_errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeCustomXmlOutOfDate]={ name: stringNS.L_CustomXmlOutOfDateName, message: stringNS.L_CustomXmlOutOfDateMessage };
			_errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeNoCapability]={ name: stringNS.L_PermissionDenied, message: stringNS.L_NoCapability };
			_errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeCannotNavTo]={ name: stringNS.L_CannotNavigateTo, message: stringNS.L_CannotNavigateTo };
			_errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeSpecifiedIdNotExist]={ name: stringNS.L_SpecifiedIdNotExist, message: stringNS.L_SpecifiedIdNotExist };
			_errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeNavOutOfBound]={ name: stringNS.L_NavOutOfBound, message: stringNS.L_NavOutOfBound };
			_errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeCellDataAmountBeyondLimits]={ name: stringNS.L_DataWriteReminder, message: stringNS.L_CellDataAmountBeyondLimits };
			_errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeElementMissing]={ name: stringNS.L_MissingParameter, message: stringNS.L_ElementMissing };
			_errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeProtectedError]={ name: stringNS.L_PermissionDenied, message: stringNS.L_NoCapability };
			_errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeInvalidCellsValue]={ name: stringNS.L_InvalidValue, message: stringNS.L_InvalidCellsValue };
			_errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeInvalidTableOptionValue]={ name: stringNS.L_InvalidValue, message: stringNS.L_InvalidTableOptionValue };
			_errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeInvalidFormatValue]={ name: stringNS.L_InvalidValue, message: stringNS.L_InvalidFormatValue };
			_errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeRowIndexOutOfRange]={ name: stringNS.L_OutOfRange, message: stringNS.L_RowIndexOutOfRange };
			_errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeColIndexOutOfRange]={ name: stringNS.L_OutOfRange, message: stringNS.L_ColIndexOutOfRange };
			_errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeFormatValueOutOfRange]={ name: stringNS.L_OutOfRange, message: stringNS.L_FormatValueOutOfRange };
			_errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeCellFormatAmountBeyondLimits]={ name: stringNS.L_FormattingReminder, message: stringNS.L_CellFormatAmountBeyondLimits };
			_errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeMemoryFileLimit]={ name: stringNS.L_MemoryLimit, message: stringNS.L_CloseFileBeforeRetrieve };
			_errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeNetworkProblemRetrieveFile]={ name: stringNS.L_NetworkProblem, message: stringNS.L_NetworkProblemRetrieveFile };
			_errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeInvalidSliceSize]={ name: stringNS.L_InvalidValue, message: stringNS.L_SliceSizeNotSupported };
			_errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeDialogAlreadyOpened]={ name: stringNS.L_DisplayDialogError, message: stringNS.L_DialogAlreadyOpened };
			_errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeInvalidWidth]={ name: stringNS.L_IndexOutOfRange, message: stringNS.L_IndexOutOfRange };
			_errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeInvalidHeight]={ name: stringNS.L_IndexOutOfRange, message: stringNS.L_IndexOutOfRange };
			_errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeNavigationError]={ name: stringNS.L_DisplayDialogError, message: stringNS.L_NetworkProblem };
			_errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeInvalidScheme]={ name: stringNS.L_DialogNavigateError, message: stringNS.L_DialogInvalidScheme };
			_errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeAppDomains]={ name: stringNS.L_DisplayDialogError, message: stringNS.L_DialogAddressNotTrusted };
			_errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeRequireHTTPS]={ name: stringNS.L_DisplayDialogError, message: stringNS.L_DialogRequireHTTPS };
			_errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeEndUserIgnore]={ name: stringNS.L_DisplayDialogError, message: stringNS.L_UserClickIgnore };
			_errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeCrossZone]={ name: stringNS.L_DisplayDialogError, message: stringNS.L_NewWindowCrossZoneErrorString };
			_errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeNotSSOAgave]={ name: stringNS.L_APINotSupported, message: stringNS.L_InvalidSSOAddinMessage };
			_errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeSSOUserNotSignedIn]={ name: stringNS.L_UserNotSignedIn, message: stringNS.L_UserNotSignedIn };
			_errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeSSOUserAborted]={ name: stringNS.L_UserAborted, message: stringNS.L_UserAbortedMessage };
			_errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeSSOUnsupportedUserIdentity]={ name: stringNS.L_UnsupportedUserIdentity, message: stringNS.L_UnsupportedUserIdentityMessage };
			_errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeSSOInvalidResourceUrl]={ name: stringNS.L_InvalidResourceUrl, message: stringNS.L_InvalidResourceUrlMessage };
			_errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeSSOInvalidGrant]={ name: stringNS.L_InvalidGrant, message: stringNS.L_InvalidGrantMessage };
			_errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeSSOClientError]={ name: stringNS.L_SSOClientError, message: stringNS.L_SSOClientErrorMessage };
			_errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeSSOServerError]={ name: stringNS.L_SSOServerError, message: stringNS.L_SSOServerErrorMessage };
			_errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeAddinIsAlreadyRequestingToken]={ name: stringNS.L_AddinIsAlreadyRequestingToken, message: stringNS.L_AddinIsAlreadyRequestingTokenMessage };
			_errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeSSOUserConsentNotSupportedByCurrentAddinCategory]={ name: stringNS.L_SSOUserConsentNotSupportedByCurrentAddinCategory, message: stringNS.L_SSOUserConsentNotSupportedByCurrentAddinCategoryMessage };
		}
	};
})();
var OfficeExt;
(function (OfficeExt) {
	var Requirement;
	(function (Requirement) {
		var RequirementVersion=(function () {
			function RequirementVersion() {
			}
			return RequirementVersion;
		})();
		Requirement.RequirementVersion=RequirementVersion;
		var RequirementMatrix=(function () {
			function RequirementMatrix(_setMap) {
				this.isSetSupported=function _isSetSupported(name, minVersion) {
					if (name==undefined) {
						return false;
					}
					if (minVersion==undefined) {
						minVersion=0;
					}
					var setSupportArray=this._setMap;
					var sets=setSupportArray._sets;
					if (sets.hasOwnProperty(name.toLowerCase())) {
						var setMaxVersion=sets[name.toLowerCase()];
						try {
							var setMaxVersionNum=this._getVersion(setMaxVersion);
							minVersion=minVersion+"";
							var minVersionNum=this._getVersion(minVersion);
							if (setMaxVersionNum.major > 0 && setMaxVersionNum.major > minVersionNum.major) {
								return true;
							}
							if (setMaxVersionNum.minor > 0 &&
								setMaxVersionNum.minor > 0 &&
								setMaxVersionNum.major==minVersionNum.major &&
								setMaxVersionNum.minor >=minVersionNum.minor) {
								return true;
							}
						}
						catch (e) {
							return false;
						}
					}
					return false;
				};
				this._getVersion=function (version) {
					var temp=version.split(".");
					var major=0;
					var minor=0;
					if (temp.length < 2 && isNaN(Number(version))) {
						throw "version format incorrect";
					}
					else {
						major=Number(temp[0]);
						if (temp.length >=2) {
							minor=Number(temp[1]);
						}
						if (isNaN(major) || isNaN(minor)) {
							throw "version format incorrect";
						}
					}
					var result={ "minor": minor, "major": major };
					return result;
				};
				this._setMap=_setMap;
				this.isSetSupported=this.isSetSupported.bind(this);
			}
			return RequirementMatrix;
		})();
		Requirement.RequirementMatrix=RequirementMatrix;
		var DefaultSetRequirement=(function () {
			function DefaultSetRequirement(setMap) {
				this._addSetMap=function DefaultSetRequirement_addSetMap(addedSet) {
					for (var name in addedSet) {
						this._sets[name]=addedSet[name];
					}
				};
				this._sets=setMap;
			}
			return DefaultSetRequirement;
		})();
		Requirement.DefaultSetRequirement=DefaultSetRequirement;
		var DefaultDialogSetRequirement=(function (_super) {
			__extends(DefaultDialogSetRequirement, _super);
			function DefaultDialogSetRequirement() {
				_super.call(this, {
					"dialogapi": 1.1
				});
			}
			return DefaultDialogSetRequirement;
		})(DefaultSetRequirement);
		Requirement.DefaultDialogSetRequirement=DefaultDialogSetRequirement;
		var ExcelClientDefaultSetRequirement=(function (_super) {
			__extends(ExcelClientDefaultSetRequirement, _super);
			function ExcelClientDefaultSetRequirement() {
				_super.call(this, {
					"bindingevents": 1.1,
					"documentevents": 1.1,
					"excelapi": 1.1,
					"matrixbindings": 1.1,
					"matrixcoercion": 1.1,
					"selection": 1.1,
					"settings": 1.1,
					"tablebindings": 1.1,
					"tablecoercion": 1.1,
					"textbindings": 1.1,
					"textcoercion": 1.1
				});
			}
			return ExcelClientDefaultSetRequirement;
		})(DefaultSetRequirement);
		Requirement.ExcelClientDefaultSetRequirement=ExcelClientDefaultSetRequirement;
		var ExcelClientV1DefaultSetRequirement=(function (_super) {
			__extends(ExcelClientV1DefaultSetRequirement, _super);
			function ExcelClientV1DefaultSetRequirement() {
				_super.call(this);
				this._addSetMap({
					"imagecoercion": 1.1
				});
			}
			return ExcelClientV1DefaultSetRequirement;
		})(ExcelClientDefaultSetRequirement);
		Requirement.ExcelClientV1DefaultSetRequirement=ExcelClientV1DefaultSetRequirement;
		var OutlookClientDefaultSetRequirement=(function (_super) {
			__extends(OutlookClientDefaultSetRequirement, _super);
			function OutlookClientDefaultSetRequirement() {
				_super.call(this, {
					"mailbox": 1.3
				});
			}
			return OutlookClientDefaultSetRequirement;
		})(DefaultSetRequirement);
		Requirement.OutlookClientDefaultSetRequirement=OutlookClientDefaultSetRequirement;
		var WordClientDefaultSetRequirement=(function (_super) {
			__extends(WordClientDefaultSetRequirement, _super);
			function WordClientDefaultSetRequirement() {
				_super.call(this, {
					"bindingevents": 1.1,
					"compressedfile": 1.1,
					"customxmlparts": 1.1,
					"documentevents": 1.1,
					"file": 1.1,
					"htmlcoercion": 1.1,
					"matrixbindings": 1.1,
					"matrixcoercion": 1.1,
					"ooxmlcoercion": 1.1,
					"pdffile": 1.1,
					"selection": 1.1,
					"settings": 1.1,
					"tablebindings": 1.1,
					"tablecoercion": 1.1,
					"textbindings": 1.1,
					"textcoercion": 1.1,
					"textfile": 1.1,
					"wordapi": 1.1
				});
			}
			return WordClientDefaultSetRequirement;
		})(DefaultSetRequirement);
		Requirement.WordClientDefaultSetRequirement=WordClientDefaultSetRequirement;
		var WordClientV1DefaultSetRequirement=(function (_super) {
			__extends(WordClientV1DefaultSetRequirement, _super);
			function WordClientV1DefaultSetRequirement() {
				_super.call(this);
				this._addSetMap({
					"customxmlparts": 1.2,
					"wordapi": 1.2,
					"imagecoercion": 1.1
				});
			}
			return WordClientV1DefaultSetRequirement;
		})(WordClientDefaultSetRequirement);
		Requirement.WordClientV1DefaultSetRequirement=WordClientV1DefaultSetRequirement;
		var PowerpointClientDefaultSetRequirement=(function (_super) {
			__extends(PowerpointClientDefaultSetRequirement, _super);
			function PowerpointClientDefaultSetRequirement() {
				_super.call(this, {
					"activeview": 1.1,
					"compressedfile": 1.1,
					"documentevents": 1.1,
					"file": 1.1,
					"pdffile": 1.1,
					"selection": 1.1,
					"settings": 1.1,
					"textcoercion": 1.1
				});
			}
			return PowerpointClientDefaultSetRequirement;
		})(DefaultSetRequirement);
		Requirement.PowerpointClientDefaultSetRequirement=PowerpointClientDefaultSetRequirement;
		var PowerpointClientV1DefaultSetRequirement=(function (_super) {
			__extends(PowerpointClientV1DefaultSetRequirement, _super);
			function PowerpointClientV1DefaultSetRequirement() {
				_super.call(this);
				this._addSetMap({
					"imagecoercion": 1.1
				});
			}
			return PowerpointClientV1DefaultSetRequirement;
		})(PowerpointClientDefaultSetRequirement);
		Requirement.PowerpointClientV1DefaultSetRequirement=PowerpointClientV1DefaultSetRequirement;
		var ProjectClientDefaultSetRequirement=(function (_super) {
			__extends(ProjectClientDefaultSetRequirement, _super);
			function ProjectClientDefaultSetRequirement() {
				_super.call(this, {
					"selection": 1.1,
					"textcoercion": 1.1
				});
			}
			return ProjectClientDefaultSetRequirement;
		})(DefaultSetRequirement);
		Requirement.ProjectClientDefaultSetRequirement=ProjectClientDefaultSetRequirement;
		var ExcelWebDefaultSetRequirement=(function (_super) {
			__extends(ExcelWebDefaultSetRequirement, _super);
			function ExcelWebDefaultSetRequirement() {
				_super.call(this, {
					"bindingevents": 1.1,
					"documentevents": 1.1,
					"matrixbindings": 1.1,
					"matrixcoercion": 1.1,
					"selection": 1.1,
					"settings": 1.1,
					"tablebindings": 1.1,
					"tablecoercion": 1.1,
					"textbindings": 1.1,
					"textcoercion": 1.1,
					"file": 1.1
				});
			}
			return ExcelWebDefaultSetRequirement;
		})(DefaultSetRequirement);
		Requirement.ExcelWebDefaultSetRequirement=ExcelWebDefaultSetRequirement;
		var WordWebDefaultSetRequirement=(function (_super) {
			__extends(WordWebDefaultSetRequirement, _super);
			function WordWebDefaultSetRequirement() {
				_super.call(this, {
					"compressedfile": 1.1,
					"documentevents": 1.1,
					"file": 1.1,
					"imagecoercion": 1.1,
					"matrixcoercion": 1.1,
					"ooxmlcoercion": 1.1,
					"pdffile": 1.1,
					"selection": 1.1,
					"settings": 1.1,
					"tablecoercion": 1.1,
					"textcoercion": 1.1,
					"textfile": 1.1
				});
			}
			return WordWebDefaultSetRequirement;
		})(DefaultSetRequirement);
		Requirement.WordWebDefaultSetRequirement=WordWebDefaultSetRequirement;
		var PowerpointWebDefaultSetRequirement=(function (_super) {
			__extends(PowerpointWebDefaultSetRequirement, _super);
			function PowerpointWebDefaultSetRequirement() {
				_super.call(this, {
					"activeview": 1.1,
					"settings": 1.1
				});
			}
			return PowerpointWebDefaultSetRequirement;
		})(DefaultSetRequirement);
		Requirement.PowerpointWebDefaultSetRequirement=PowerpointWebDefaultSetRequirement;
		var OutlookWebDefaultSetRequirement=(function (_super) {
			__extends(OutlookWebDefaultSetRequirement, _super);
			function OutlookWebDefaultSetRequirement() {
				_super.call(this, {
					"mailbox": 1.3
				});
			}
			return OutlookWebDefaultSetRequirement;
		})(DefaultSetRequirement);
		Requirement.OutlookWebDefaultSetRequirement=OutlookWebDefaultSetRequirement;
		var SwayWebDefaultSetRequirement=(function (_super) {
			__extends(SwayWebDefaultSetRequirement, _super);
			function SwayWebDefaultSetRequirement() {
				_super.call(this, {
					"activeview": 1.1,
					"documentevents": 1.1,
					"selection": 1.1,
					"settings": 1.1,
					"textcoercion": 1.1
				});
			}
			return SwayWebDefaultSetRequirement;
		})(DefaultSetRequirement);
		Requirement.SwayWebDefaultSetRequirement=SwayWebDefaultSetRequirement;
		var AccessWebDefaultSetRequirement=(function (_super) {
			__extends(AccessWebDefaultSetRequirement, _super);
			function AccessWebDefaultSetRequirement() {
				_super.call(this, {
					"bindingevents": 1.1,
					"partialtablebindings": 1.1,
					"settings": 1.1,
					"tablebindings": 1.1,
					"tablecoercion": 1.1
				});
			}
			return AccessWebDefaultSetRequirement;
		})(DefaultSetRequirement);
		Requirement.AccessWebDefaultSetRequirement=AccessWebDefaultSetRequirement;
		var ExcelIOSDefaultSetRequirement=(function (_super) {
			__extends(ExcelIOSDefaultSetRequirement, _super);
			function ExcelIOSDefaultSetRequirement() {
				_super.call(this, {
					"bindingevents": 1.1,
					"documentevents": 1.1,
					"matrixbindings": 1.1,
					"matrixcoercion": 1.1,
					"selection": 1.1,
					"settings": 1.1,
					"tablebindings": 1.1,
					"tablecoercion": 1.1,
					"textbindings": 1.1,
					"textcoercion": 1.1
				});
			}
			return ExcelIOSDefaultSetRequirement;
		})(DefaultSetRequirement);
		Requirement.ExcelIOSDefaultSetRequirement=ExcelIOSDefaultSetRequirement;
		var WordIOSDefaultSetRequirement=(function (_super) {
			__extends(WordIOSDefaultSetRequirement, _super);
			function WordIOSDefaultSetRequirement() {
				_super.call(this, {
					"bindingevents": 1.1,
					"compressedfile": 1.1,
					"customxmlparts": 1.1,
					"documentevents": 1.1,
					"file": 1.1,
					"htmlcoercion": 1.1,
					"matrixbindings": 1.1,
					"matrixcoercion": 1.1,
					"ooxmlcoercion": 1.1,
					"pdffile": 1.1,
					"selection": 1.1,
					"settings": 1.1,
					"tablebindings": 1.1,
					"tablecoercion": 1.1,
					"textbindings": 1.1,
					"textcoercion": 1.1,
					"textfile": 1.1
				});
			}
			return WordIOSDefaultSetRequirement;
		})(DefaultSetRequirement);
		Requirement.WordIOSDefaultSetRequirement=WordIOSDefaultSetRequirement;
		var WordIOSV1DefaultSetRequirement=(function (_super) {
			__extends(WordIOSV1DefaultSetRequirement, _super);
			function WordIOSV1DefaultSetRequirement() {
				_super.call(this);
				this._addSetMap({
					"customxmlparts": 1.2,
					"wordapi": 1.2
				});
			}
			return WordIOSV1DefaultSetRequirement;
		})(WordIOSDefaultSetRequirement);
		Requirement.WordIOSV1DefaultSetRequirement=WordIOSV1DefaultSetRequirement;
		var PowerpointIOSDefaultSetRequirement=(function (_super) {
			__extends(PowerpointIOSDefaultSetRequirement, _super);
			function PowerpointIOSDefaultSetRequirement() {
				_super.call(this, {
					"activeview": 1.1,
					"compressedfile": 1.1,
					"documentevents": 1.1,
					"file": 1.1,
					"pdffile": 1.1,
					"selection": 1.1,
					"settings": 1.1,
					"textcoercion": 1.1
				});
			}
			return PowerpointIOSDefaultSetRequirement;
		})(DefaultSetRequirement);
		Requirement.PowerpointIOSDefaultSetRequirement=PowerpointIOSDefaultSetRequirement;
		var OutlookIOSDefaultSetRequirement=(function (_super) {
			__extends(OutlookIOSDefaultSetRequirement, _super);
			function OutlookIOSDefaultSetRequirement() {
				_super.call(this, {
					"mailbox": 1.1
				});
			}
			return OutlookIOSDefaultSetRequirement;
		})(DefaultSetRequirement);
		Requirement.OutlookIOSDefaultSetRequirement=OutlookIOSDefaultSetRequirement;
		var RequirementsMatrixFactory=(function () {
			function RequirementsMatrixFactory() {
			}
			RequirementsMatrixFactory.initializeOsfDda=function () {
				OSF.OUtil.setNamespace("Requirement", OSF.DDA);
			};
			RequirementsMatrixFactory.getDefaultRequirementMatrix=function (appContext) {
				this.initializeDefaultSetMatrix();
				var defaultRequirementMatrix=undefined;
				var clientRequirement=appContext.get_requirementMatrix();
				if (clientRequirement !=undefined && clientRequirement.length > 0 && typeof (JSON) !=="undefined") {
					var matrixItem=JSON.parse(appContext.get_requirementMatrix().toLowerCase());
					defaultRequirementMatrix=new RequirementMatrix(new DefaultSetRequirement(matrixItem));
				}
				else {
					var appLocator=RequirementsMatrixFactory.getClientFullVersionString(appContext);
					if (RequirementsMatrixFactory.DefaultSetArrayMatrix !=undefined && RequirementsMatrixFactory.DefaultSetArrayMatrix[appLocator] !=undefined) {
						defaultRequirementMatrix=new RequirementMatrix(RequirementsMatrixFactory.DefaultSetArrayMatrix[appLocator]);
					}
					else {
						defaultRequirementMatrix=new RequirementMatrix(new DefaultSetRequirement({}));
					}
				}
				return defaultRequirementMatrix;
			};
			RequirementsMatrixFactory.getDefaultDialogRequirementMatrix=function (appContext) {
				var defaultRequirementMatrix=undefined;
				var clientRequirement=appContext.get_dialogRequirementMatrix();
				if (clientRequirement !=undefined && clientRequirement.length > 0 && typeof (JSON) !=="undefined") {
					var matrixItem=JSON.parse(appContext.get_requirementMatrix().toLowerCase());
					defaultRequirementMatrix=new RequirementMatrix(new DefaultSetRequirement(matrixItem));
				}
				else {
					defaultRequirementMatrix=new RequirementMatrix(new DefaultDialogSetRequirement());
				}
				return defaultRequirementMatrix;
			};
			RequirementsMatrixFactory.getClientFullVersionString=function (appContext) {
				var appMinorVersion=appContext.get_appMinorVersion();
				var appMinorVersionString="";
				var appFullVersion="";
				var appName=appContext.get_appName();
				var isIOSClient=appName==1024 ||
					appName==4096 ||
					appName==8192 ||
					appName==65536;
				if (isIOSClient && appContext.get_appVersion()==1) {
					if (appName==4096 && appMinorVersion >=15) {
						appFullVersion="16.00.01";
					}
					else {
						appFullVersion="16.00";
					}
				}
				else if (appContext.get_appName()==64) {
					appFullVersion=appContext.get_appVersion();
				}
				else {
					if (appMinorVersion < 10) {
						appMinorVersionString="0"+appMinorVersion;
					}
					else {
						appMinorVersionString=""+appMinorVersion;
					}
					appFullVersion=appContext.get_appVersion()+"."+appMinorVersionString;
				}
				return appContext.get_appName()+"-"+appFullVersion;
			};
			RequirementsMatrixFactory.initializeDefaultSetMatrix=function () {
				RequirementsMatrixFactory.DefaultSetArrayMatrix[RequirementsMatrixFactory.Excel_RCLIENT_1600]=new ExcelClientDefaultSetRequirement();
				RequirementsMatrixFactory.DefaultSetArrayMatrix[RequirementsMatrixFactory.Word_RCLIENT_1600]=new WordClientDefaultSetRequirement();
				RequirementsMatrixFactory.DefaultSetArrayMatrix[RequirementsMatrixFactory.PowerPoint_RCLIENT_1600]=new PowerpointClientDefaultSetRequirement();
				RequirementsMatrixFactory.DefaultSetArrayMatrix[RequirementsMatrixFactory.Excel_RCLIENT_1601]=new ExcelClientV1DefaultSetRequirement();
				RequirementsMatrixFactory.DefaultSetArrayMatrix[RequirementsMatrixFactory.Word_RCLIENT_1601]=new WordClientV1DefaultSetRequirement();
				RequirementsMatrixFactory.DefaultSetArrayMatrix[RequirementsMatrixFactory.PowerPoint_RCLIENT_1601]=new PowerpointClientV1DefaultSetRequirement();
				RequirementsMatrixFactory.DefaultSetArrayMatrix[RequirementsMatrixFactory.Outlook_RCLIENT_1600]=new OutlookClientDefaultSetRequirement();
				RequirementsMatrixFactory.DefaultSetArrayMatrix[RequirementsMatrixFactory.Excel_WAC_1600]=new ExcelWebDefaultSetRequirement();
				RequirementsMatrixFactory.DefaultSetArrayMatrix[RequirementsMatrixFactory.Word_WAC_1600]=new WordWebDefaultSetRequirement();
				RequirementsMatrixFactory.DefaultSetArrayMatrix[RequirementsMatrixFactory.Outlook_WAC_1600]=new OutlookWebDefaultSetRequirement();
				RequirementsMatrixFactory.DefaultSetArrayMatrix[RequirementsMatrixFactory.Outlook_WAC_1601]=new OutlookWebDefaultSetRequirement();
				RequirementsMatrixFactory.DefaultSetArrayMatrix[RequirementsMatrixFactory.Project_RCLIENT_1600]=new ProjectClientDefaultSetRequirement();
				RequirementsMatrixFactory.DefaultSetArrayMatrix[RequirementsMatrixFactory.Access_WAC_1600]=new AccessWebDefaultSetRequirement();
				RequirementsMatrixFactory.DefaultSetArrayMatrix[RequirementsMatrixFactory.PowerPoint_WAC_1600]=new PowerpointWebDefaultSetRequirement();
				RequirementsMatrixFactory.DefaultSetArrayMatrix[RequirementsMatrixFactory.Excel_IOS_1600]=new ExcelIOSDefaultSetRequirement();
				RequirementsMatrixFactory.DefaultSetArrayMatrix[RequirementsMatrixFactory.SWAY_WAC_1600]=new SwayWebDefaultSetRequirement();
				RequirementsMatrixFactory.DefaultSetArrayMatrix[RequirementsMatrixFactory.Word_IOS_1600]=new WordIOSDefaultSetRequirement();
				RequirementsMatrixFactory.DefaultSetArrayMatrix[RequirementsMatrixFactory.Word_IOS_16001]=new WordIOSV1DefaultSetRequirement();
				RequirementsMatrixFactory.DefaultSetArrayMatrix[RequirementsMatrixFactory.PowerPoint_IOS_1600]=new PowerpointIOSDefaultSetRequirement();
				RequirementsMatrixFactory.DefaultSetArrayMatrix[RequirementsMatrixFactory.Outlook_IOS_1600]=new OutlookIOSDefaultSetRequirement();
			};
			RequirementsMatrixFactory.Excel_RCLIENT_1600="1-16.00";
			RequirementsMatrixFactory.Excel_RCLIENT_1601="1-16.01";
			RequirementsMatrixFactory.Word_RCLIENT_1600="2-16.00";
			RequirementsMatrixFactory.Word_RCLIENT_1601="2-16.01";
			RequirementsMatrixFactory.PowerPoint_RCLIENT_1600="4-16.00";
			RequirementsMatrixFactory.PowerPoint_RCLIENT_1601="4-16.01";
			RequirementsMatrixFactory.Outlook_RCLIENT_1600="8-16.00";
			RequirementsMatrixFactory.Excel_WAC_1600="16-16.00";
			RequirementsMatrixFactory.Word_WAC_1600="32-16.00";
			RequirementsMatrixFactory.Outlook_WAC_1600="64-16.00";
			RequirementsMatrixFactory.Outlook_WAC_1601="64-16.01";
			RequirementsMatrixFactory.Project_RCLIENT_1600="128-16.00";
			RequirementsMatrixFactory.Access_WAC_1600="256-16.00";
			RequirementsMatrixFactory.PowerPoint_WAC_1600="512-16.00";
			RequirementsMatrixFactory.Excel_IOS_1600="1024-16.00";
			RequirementsMatrixFactory.SWAY_WAC_1600="2048-16.00";
			RequirementsMatrixFactory.Word_IOS_1600="4096-16.00";
			RequirementsMatrixFactory.Word_IOS_16001="4096-16.00.01";
			RequirementsMatrixFactory.PowerPoint_IOS_1600="8192-16.00";
			RequirementsMatrixFactory.Outlook_IOS_1600="65536-16.00";
			RequirementsMatrixFactory.DefaultSetArrayMatrix={};
			return RequirementsMatrixFactory;
		})();
		Requirement.RequirementsMatrixFactory=RequirementsMatrixFactory;
	})(Requirement=OfficeExt.Requirement || (OfficeExt.Requirement={}));
})(OfficeExt || (OfficeExt={}));
OfficeExt.Requirement.RequirementsMatrixFactory.initializeOsfDda();
var OfficeExt;
(function (OfficeExt) {
	var HostName;
	(function (HostName) {
		var Host=(function () {
			function Host() {
				this.getDiagnostics=function _getDiagnostics(version) {
					var diagnostics={
						host: this.getHost(),
						version: (version || this.getDefaultVersion()),
						platform: this.getPlatform()
					};
					return diagnostics;
				};
				this.platformRemappings={
					web: Microsoft.Office.WebExtension.PlatformType.OfficeOnline,
					winrt: Microsoft.Office.WebExtension.PlatformType.Universal,
					win32: Microsoft.Office.WebExtension.PlatformType.PC,
					mac: Microsoft.Office.WebExtension.PlatformType.Mac,
					ios: Microsoft.Office.WebExtension.PlatformType.iOS,
					android: Microsoft.Office.WebExtension.PlatformType.Android
				};
				this.camelCaseMappings={
					powerpoint: Microsoft.Office.WebExtension.HostType.PowerPoint,
					onenote: Microsoft.Office.WebExtension.HostType.OneNote
				};
				this.hostInfo=OSF._OfficeAppFactory.getHostInfo();
				this.getHost=this.getHost.bind(this);
				this.getPlatform=this.getPlatform.bind(this);
				this.getDiagnostics=this.getDiagnostics.bind(this);
			}
			Host.prototype.capitalizeFirstLetter=function (input) {
				if (input) {
					return (input[0].toUpperCase()+input.slice(1).toLowerCase());
				}
				return input;
			};
			Host.getInstance=function () {
				if (Host.hostObj===undefined) {
					Host.hostObj=new Host();
				}
				return Host.hostObj;
			};
			Host.prototype.getPlatform=function () {
				if (this.hostInfo.hostPlatform) {
					var hostPlatform=this.hostInfo.hostPlatform.toLowerCase();
					if (this.platformRemappings[hostPlatform]) {
						return this.platformRemappings[hostPlatform];
					}
				}
				return null;
			};
			Host.prototype.getHost=function () {
				if (this.hostInfo.hostType) {
					var hostType=this.hostInfo.hostType.toLowerCase();
					if (this.camelCaseMappings[hostType]) {
						return this.camelCaseMappings[hostType];
					}
					hostType=this.capitalizeFirstLetter(this.hostInfo.hostType);
					if (Microsoft.Office.WebExtension.HostType[hostType]) {
						return Microsoft.Office.WebExtension.HostType[hostType];
					}
				}
				return null;
			};
			Host.prototype.getDefaultVersion=function () {
				if (this.getHost()) {
					return "16.0.0000.0000";
				}
				return null;
			};
			return Host;
		})();
		HostName.Host=Host;
	})(HostName=OfficeExt.HostName || (OfficeExt.HostName={}));
})(OfficeExt || (OfficeExt={}));
Microsoft.Office.WebExtension.ApplicationMode={
	WebEditor: "webEditor",
	WebViewer: "webViewer",
	Client: "client"
};
Microsoft.Office.WebExtension.DocumentMode={
	ReadOnly: "readOnly",
	ReadWrite: "readWrite"
};
OSF.NamespaceManager=(function OSF_NamespaceManager() {
	var _userOffice;
	var _useShortcut=false;
	return {
		enableShortcut: function OSF_NamespaceManager$enableShortcut() {
			if (!_useShortcut) {
				if (window.Office) {
					_userOffice=window.Office;
				}
				else {
					OSF.OUtil.setNamespace("Office", window);
				}
				window.Office=Microsoft.Office.WebExtension;
				_useShortcut=true;
			}
		},
		disableShortcut: function OSF_NamespaceManager$disableShortcut() {
			if (_useShortcut) {
				if (_userOffice) {
					window.Office=_userOffice;
				}
				else {
					OSF.OUtil.unsetNamespace("Office", window);
				}
				_useShortcut=false;
			}
		}
	};
})();
OSF.NamespaceManager.enableShortcut();
Microsoft.Office.WebExtension.useShortNamespace=function Microsoft_Office_WebExtension_useShortNamespace(useShortcut) {
	if (useShortcut) {
		OSF.NamespaceManager.enableShortcut();
	}
	else {
		OSF.NamespaceManager.disableShortcut();
	}
};
Microsoft.Office.WebExtension.select=function Microsoft_Office_WebExtension_select(str, errorCallback) {
	var promise;
	if (str && typeof str=="string") {
		var index=str.indexOf("#");
		if (index !=-1) {
			var op=str.substring(0, index);
			var target=str.substring(index+1);
			switch (op) {
				case "binding":
				case "bindings":
					if (target) {
						promise=new OSF.DDA.BindingPromise(target);
					}
					break;
			}
		}
	}
	if (!promise) {
		if (errorCallback) {
			var callbackType=typeof errorCallback;
			if (callbackType=="function") {
				var callArgs={};
				callArgs[Microsoft.Office.WebExtension.Parameters.Callback]=errorCallback;
				OSF.DDA.issueAsyncResult(callArgs, OSF.DDA.ErrorCodeManager.errorCodes.ooeInvalidApiCallInContext, OSF.DDA.ErrorCodeManager.getErrorArgs(OSF.DDA.ErrorCodeManager.errorCodes.ooeInvalidApiCallInContext));
			}
			else {
				throw OSF.OUtil.formatString(Strings.OfficeOM.L_CallbackNotAFunction, callbackType);
			}
		}
	}
	else {
		promise.onFail=errorCallback;
		return promise;
	}
};
OSF.DDA.Context=function OSF_DDA_Context(officeAppContext, document, license, appOM, getOfficeTheme) {
	OSF.OUtil.defineEnumerableProperties(this, {
		"contentLanguage": {
			value: officeAppContext.get_dataLocale()
		},
		"displayLanguage": {
			value: officeAppContext.get_appUILocale()
		},
		"touchEnabled": {
			value: officeAppContext.get_touchEnabled()
		},
		"commerceAllowed": {
			value: officeAppContext.get_commerceAllowed()
		},
		"host": {
			value: OfficeExt.HostName.Host.getInstance().getHost()
		},
		"platform": {
			value: OfficeExt.HostName.Host.getInstance().getPlatform()
		},
		"diagnostics": {
			value: OfficeExt.HostName.Host.getInstance().getDiagnostics(officeAppContext.get_hostFullVersion())
		}
	});
	if (license) {
		OSF.OUtil.defineEnumerableProperty(this, "license", {
			value: license
		});
	}
	if (officeAppContext.ui) {
		OSF.OUtil.defineEnumerableProperty(this, "ui", {
			value: officeAppContext.ui
		});
	}
	if (officeAppContext.auth) {
		OSF.OUtil.defineEnumerableProperty(this, "auth", {
			value: officeAppContext.auth
		});
	}
	if (officeAppContext.get_isDialog()) {
		var requirements=OfficeExt.Requirement.RequirementsMatrixFactory.getDefaultDialogRequirementMatrix(officeAppContext);
		OSF.OUtil.defineEnumerableProperty(this, "requirements", {
			value: requirements
		});
	}
	else {
		if (document) {
			OSF.OUtil.defineEnumerableProperty(this, "document", {
				value: document
			});
		}
		if (appOM) {
			var displayName=appOM.displayName || "appOM";
			delete appOM.displayName;
			OSF.OUtil.defineEnumerableProperty(this, displayName, {
				value: appOM
			});
		}
		if (getOfficeTheme) {
			OSF.OUtil.defineEnumerableProperty(this, "officeTheme", {
				get: function () {
					return getOfficeTheme();
				}
			});
		}
		var requirements=OfficeExt.Requirement.RequirementsMatrixFactory.getDefaultRequirementMatrix(officeAppContext);
		OSF.OUtil.defineEnumerableProperty(this, "requirements", {
			value: requirements
		});
	}
};
OSF.DDA.OutlookContext=function OSF_DDA_OutlookContext(appContext, settings, license, appOM, getOfficeTheme) {
	OSF.DDA.OutlookContext.uber.constructor.call(this, appContext, null, license, appOM, getOfficeTheme);
	if (settings) {
		OSF.OUtil.defineEnumerableProperty(this, "roamingSettings", {
			value: settings
		});
	}
};
OSF.OUtil.extend(OSF.DDA.OutlookContext, OSF.DDA.Context);
OSF.DDA.OutlookAppOm=function OSF_DDA_OutlookAppOm(appContext, window, appReady) { };
OSF.DDA.Document=function OSF_DDA_Document(officeAppContext, settings) {
	var mode;
	switch (officeAppContext.get_clientMode()) {
		case OSF.ClientMode.ReadOnly:
			mode=Microsoft.Office.WebExtension.DocumentMode.ReadOnly;
			break;
		case OSF.ClientMode.ReadWrite:
			mode=Microsoft.Office.WebExtension.DocumentMode.ReadWrite;
			break;
	}
	;
	if (settings) {
		OSF.OUtil.defineEnumerableProperty(this, "settings", {
			value: settings
		});
	}
	;
	OSF.OUtil.defineMutableProperties(this, {
		"mode": {
			value: mode
		},
		"url": {
			value: officeAppContext.get_docUrl()
		}
	});
};
OSF.DDA.JsomDocument=function OSF_DDA_JsomDocument(officeAppContext, bindingFacade, settings) {
	OSF.DDA.JsomDocument.uber.constructor.call(this, officeAppContext, settings);
	if (bindingFacade) {
		OSF.OUtil.defineEnumerableProperty(this, "bindings", {
			get: function OSF_DDA_Document$GetBindings() { return bindingFacade; }
		});
	}
	var am=OSF.DDA.AsyncMethodNames;
	OSF.DDA.DispIdHost.addAsyncMethods(this, [
		am.GetSelectedDataAsync,
		am.SetSelectedDataAsync
	]);
	OSF.DDA.DispIdHost.addEventSupport(this, new OSF.EventDispatch([Microsoft.Office.WebExtension.EventType.DocumentSelectionChanged]));
};
OSF.OUtil.extend(OSF.DDA.JsomDocument, OSF.DDA.Document);
OSF.OUtil.defineEnumerableProperty(Microsoft.Office.WebExtension, "context", {
	get: function Microsoft_Office_WebExtension$GetContext() {
		var context;
		if (OSF && OSF._OfficeAppFactory) {
			context=OSF._OfficeAppFactory.getContext();
		}
		return context;
	}
});
OSF.DDA.License=function OSF_DDA_License(eToken) {
	OSF.OUtil.defineEnumerableProperty(this, "value", {
		value: eToken
	});
};
OSF.DDA.ApiMethodCall=function OSF_DDA_ApiMethodCall(requiredParameters, supportedOptions, privateStateCallbacks, checkCallArgs, displayName) {
	var requiredCount=requiredParameters.length;
	var getInvalidParameterString=OSF.OUtil.delayExecutionAndCache(function () {
		return OSF.OUtil.formatString(Strings.OfficeOM.L_InvalidParameters, displayName);
	});
	this.verifyArguments=function OSF_DDA_ApiMethodCall$VerifyArguments(params, args) {
		for (var name in params) {
			var param=params[name];
			var arg=args[name];
			if (param["enum"]) {
				switch (typeof arg) {
					case "string":
						if (OSF.OUtil.listContainsValue(param["enum"], arg)) {
							break;
						}
					case "undefined":
						throw OSF.DDA.ErrorCodeManager.errorCodes.ooeUnsupportedEnumeration;
					default:
						throw getInvalidParameterString();
				}
			}
			if (param["types"]) {
				if (!OSF.OUtil.listContainsValue(param["types"], typeof arg)) {
					throw getInvalidParameterString();
				}
			}
		}
	};
	this.extractRequiredArguments=function OSF_DDA_ApiMethodCall$ExtractRequiredArguments(userArgs, caller, stateInfo) {
		if (userArgs.length < requiredCount) {
			throw OsfMsAjaxFactory.msAjaxError.parameterCount(Strings.OfficeOM.L_MissingRequiredArguments);
		}
		var requiredArgs=[];
		var index;
		for (index=0; index < requiredCount; index++) {
			requiredArgs.push(userArgs[index]);
		}
		this.verifyArguments(requiredParameters, requiredArgs);
		var ret={};
		for (index=0; index < requiredCount; index++) {
			var param=requiredParameters[index];
			var arg=requiredArgs[index];
			if (param.verify) {
				var isValid=param.verify(arg, caller, stateInfo);
				if (!isValid) {
					throw getInvalidParameterString();
				}
			}
			ret[param.name]=arg;
		}
		return ret;
	},
		this.fillOptions=function OSF_DDA_ApiMethodCall$FillOptions(options, requiredArgs, caller, stateInfo) {
			options=options || {};
			for (var optionName in supportedOptions) {
				if (!OSF.OUtil.listContainsKey(options, optionName)) {
					var value=undefined;
					var option=supportedOptions[optionName];
					if (option.calculate && requiredArgs) {
						value=option.calculate(requiredArgs, caller, stateInfo);
					}
					if (!value && option.defaultValue !==undefined) {
						value=option.defaultValue;
					}
					options[optionName]=value;
				}
			}
			return options;
		};
	this.constructCallArgs=function OSF_DAA_ApiMethodCall$ConstructCallArgs(required, options, caller, stateInfo) {
		var callArgs={};
		for (var r in required) {
			callArgs[r]=required[r];
		}
		for (var o in options) {
			callArgs[o]=options[o];
		}
		for (var s in privateStateCallbacks) {
			callArgs[s]=privateStateCallbacks[s](caller, stateInfo);
		}
		if (checkCallArgs) {
			callArgs=checkCallArgs(callArgs, caller, stateInfo);
		}
		return callArgs;
	};
};
OSF.OUtil.setNamespace("AsyncResultEnum", OSF.DDA);
OSF.DDA.AsyncResultEnum.Properties={
	Context: "Context",
	Value: "Value",
	Status: "Status",
	Error: "Error"
};
Microsoft.Office.WebExtension.AsyncResultStatus={
	Succeeded: "succeeded",
	Failed: "failed"
};
OSF.DDA.AsyncResultEnum.ErrorCode={
	Success: 0,
	Failed: 1
};
OSF.DDA.AsyncResultEnum.ErrorProperties={
	Name: "Name",
	Message: "Message",
	Code: "Code"
};
OSF.DDA.AsyncMethodNames={};
OSF.DDA.AsyncMethodNames.addNames=function (methodNames) {
	for (var entry in methodNames) {
		var am={};
		OSF.OUtil.defineEnumerableProperties(am, {
			"id": {
				value: entry
			},
			"displayName": {
				value: methodNames[entry]
			}
		});
		OSF.DDA.AsyncMethodNames[entry]=am;
	}
};
OSF.DDA.AsyncMethodCall=function OSF_DDA_AsyncMethodCall(requiredParameters, supportedOptions, privateStateCallbacks, onSucceeded, onFailed, checkCallArgs, displayName) {
	var requiredCount=requiredParameters.length;
	var apiMethods=new OSF.DDA.ApiMethodCall(requiredParameters, supportedOptions, privateStateCallbacks, checkCallArgs, displayName);
	function OSF_DAA_AsyncMethodCall$ExtractOptions(userArgs, requiredArgs, caller, stateInfo) {
		if (userArgs.length > requiredCount+2) {
			throw OsfMsAjaxFactory.msAjaxError.parameterCount(Strings.OfficeOM.L_TooManyArguments);
		}
		var options, parameterCallback;
		for (var i=userArgs.length - 1; i >=requiredCount; i--) {
			var argument=userArgs[i];
			switch (typeof argument) {
				case "object":
					if (options) {
						throw OsfMsAjaxFactory.msAjaxError.parameterCount(Strings.OfficeOM.L_TooManyOptionalObjects);
					}
					else {
						options=argument;
					}
					break;
				case "function":
					if (parameterCallback) {
						throw OsfMsAjaxFactory.msAjaxError.parameterCount(Strings.OfficeOM.L_TooManyOptionalFunction);
					}
					else {
						parameterCallback=argument;
					}
					break;
				default:
					throw OsfMsAjaxFactory.msAjaxError.argument(Strings.OfficeOM.L_InValidOptionalArgument);
					break;
			}
		}
		options=apiMethods.fillOptions(options, requiredArgs, caller, stateInfo);
		if (parameterCallback) {
			if (options[Microsoft.Office.WebExtension.Parameters.Callback]) {
				throw Strings.OfficeOM.L_RedundantCallbackSpecification;
			}
			else {
				options[Microsoft.Office.WebExtension.Parameters.Callback]=parameterCallback;
			}
		}
		apiMethods.verifyArguments(supportedOptions, options);
		return options;
	}
	;
	this.verifyAndExtractCall=function OSF_DAA_AsyncMethodCall$VerifyAndExtractCall(userArgs, caller, stateInfo) {
		var required=apiMethods.extractRequiredArguments(userArgs, caller, stateInfo);
		var options=OSF_DAA_AsyncMethodCall$ExtractOptions(userArgs, required, caller, stateInfo);
		var callArgs=apiMethods.constructCallArgs(required, options, caller, stateInfo);
		return callArgs;
	};
	this.processResponse=function OSF_DAA_AsyncMethodCall$ProcessResponse(status, response, caller, callArgs) {
		var payload;
		if (status==OSF.DDA.ErrorCodeManager.errorCodes.ooeSuccess) {
			if (onSucceeded) {
				payload=onSucceeded(response, caller, callArgs);
			}
			else {
				payload=response;
			}
		}
		else {
			if (onFailed) {
				payload=onFailed(status, response);
			}
			else {
				payload=OSF.DDA.ErrorCodeManager.getErrorArgs(status);
			}
		}
		return payload;
	};
	this.getCallArgs=function (suppliedArgs) {
		var options, parameterCallback;
		for (var i=suppliedArgs.length - 1; i >=requiredCount; i--) {
			var argument=suppliedArgs[i];
			switch (typeof argument) {
				case "object":
					options=argument;
					break;
				case "function":
					parameterCallback=argument;
					break;
			}
		}
		options=options || {};
		if (parameterCallback) {
			options[Microsoft.Office.WebExtension.Parameters.Callback]=parameterCallback;
		}
		return options;
	};
};
OSF.DDA.AsyncMethodCallFactory=(function () {
	return {
		manufacture: function (params) {
			var supportedOptions=params.supportedOptions ? OSF.OUtil.createObject(params.supportedOptions) : [];
			var privateStateCallbacks=params.privateStateCallbacks ? OSF.OUtil.createObject(params.privateStateCallbacks) : [];
			return new OSF.DDA.AsyncMethodCall(params.requiredArguments || [], supportedOptions, privateStateCallbacks, params.onSucceeded, params.onFailed, params.checkCallArgs, params.method.displayName);
		}
	};
})();
OSF.DDA.AsyncMethodCalls={};
OSF.DDA.AsyncMethodCalls.define=function (callDefinition) {
	OSF.DDA.AsyncMethodCalls[callDefinition.method.id]=OSF.DDA.AsyncMethodCallFactory.manufacture(callDefinition);
};
OSF.DDA.Error=function OSF_DDA_Error(name, message, code) {
	OSF.OUtil.defineEnumerableProperties(this, {
		"name": {
			value: name
		},
		"message": {
			value: message
		},
		"code": {
			value: code
		}
	});
};
OSF.DDA.AsyncResult=function OSF_DDA_AsyncResult(initArgs, errorArgs) {
	OSF.OUtil.defineEnumerableProperties(this, {
		"value": {
			value: initArgs[OSF.DDA.AsyncResultEnum.Properties.Value]
		},
		"status": {
			value: errorArgs ? Microsoft.Office.WebExtension.AsyncResultStatus.Failed : Microsoft.Office.WebExtension.AsyncResultStatus.Succeeded
		}
	});
	if (initArgs[OSF.DDA.AsyncResultEnum.Properties.Context]) {
		OSF.OUtil.defineEnumerableProperty(this, "asyncContext", {
			value: initArgs[OSF.DDA.AsyncResultEnum.Properties.Context]
		});
	}
	if (errorArgs) {
		OSF.OUtil.defineEnumerableProperty(this, "error", {
			value: new OSF.DDA.Error(errorArgs[OSF.DDA.AsyncResultEnum.ErrorProperties.Name], errorArgs[OSF.DDA.AsyncResultEnum.ErrorProperties.Message], errorArgs[OSF.DDA.AsyncResultEnum.ErrorProperties.Code])
		});
	}
};
OSF.DDA.issueAsyncResult=function OSF_DDA$IssueAsyncResult(callArgs, status, payload) {
	var callback=callArgs[Microsoft.Office.WebExtension.Parameters.Callback];
	if (callback) {
		var asyncInitArgs={};
		asyncInitArgs[OSF.DDA.AsyncResultEnum.Properties.Context]=callArgs[Microsoft.Office.WebExtension.Parameters.AsyncContext];
		var errorArgs;
		if (status==OSF.DDA.ErrorCodeManager.errorCodes.ooeSuccess) {
			asyncInitArgs[OSF.DDA.AsyncResultEnum.Properties.Value]=payload;
		}
		else {
			errorArgs={};
			payload=payload || OSF.DDA.ErrorCodeManager.getErrorArgs(OSF.DDA.ErrorCodeManager.errorCodes.ooeInternalError);
			errorArgs[OSF.DDA.AsyncResultEnum.ErrorProperties.Code]=status || OSF.DDA.ErrorCodeManager.errorCodes.ooeInternalError;
			errorArgs[OSF.DDA.AsyncResultEnum.ErrorProperties.Name]=payload.name || payload;
			errorArgs[OSF.DDA.AsyncResultEnum.ErrorProperties.Message]=payload.message || payload;
		}
		callback(new OSF.DDA.AsyncResult(asyncInitArgs, errorArgs));
	}
};
OSF.DDA.SyncMethodNames={};
OSF.DDA.SyncMethodNames.addNames=function (methodNames) {
	for (var entry in methodNames) {
		var am={};
		OSF.OUtil.defineEnumerableProperties(am, {
			"id": {
				value: entry
			},
			"displayName": {
				value: methodNames[entry]
			}
		});
		OSF.DDA.SyncMethodNames[entry]=am;
	}
};
OSF.DDA.SyncMethodCall=function OSF_DDA_SyncMethodCall(requiredParameters, supportedOptions, privateStateCallbacks, checkCallArgs, displayName) {
	var requiredCount=requiredParameters.length;
	var apiMethods=new OSF.DDA.ApiMethodCall(requiredParameters, supportedOptions, privateStateCallbacks, checkCallArgs, displayName);
	function OSF_DAA_SyncMethodCall$ExtractOptions(userArgs, requiredArgs, caller, stateInfo) {
		if (userArgs.length > requiredCount+1) {
			throw OsfMsAjaxFactory.msAjaxError.parameterCount(Strings.OfficeOM.L_TooManyArguments);
		}
		var options, parameterCallback;
		for (var i=userArgs.length - 1; i >=requiredCount; i--) {
			var argument=userArgs[i];
			switch (typeof argument) {
				case "object":
					if (options) {
						throw OsfMsAjaxFactory.msAjaxError.parameterCount(Strings.OfficeOM.L_TooManyOptionalObjects);
					}
					else {
						options=argument;
					}
					break;
				default:
					throw OsfMsAjaxFactory.msAjaxError.argument(Strings.OfficeOM.L_InValidOptionalArgument);
					break;
			}
		}
		options=apiMethods.fillOptions(options, requiredArgs, caller, stateInfo);
		apiMethods.verifyArguments(supportedOptions, options);
		return options;
	}
	;
	this.verifyAndExtractCall=function OSF_DAA_AsyncMethodCall$VerifyAndExtractCall(userArgs, caller, stateInfo) {
		var required=apiMethods.extractRequiredArguments(userArgs, caller, stateInfo);
		var options=OSF_DAA_SyncMethodCall$ExtractOptions(userArgs, required, caller, stateInfo);
		var callArgs=apiMethods.constructCallArgs(required, options, caller, stateInfo);
		return callArgs;
	};
};
OSF.DDA.SyncMethodCallFactory=(function () {
	return {
		manufacture: function (params) {
			var supportedOptions=params.supportedOptions ? OSF.OUtil.createObject(params.supportedOptions) : [];
			return new OSF.DDA.SyncMethodCall(params.requiredArguments || [], supportedOptions, params.privateStateCallbacks, params.checkCallArgs, params.method.displayName);
		}
	};
})();
OSF.DDA.SyncMethodCalls={};
OSF.DDA.SyncMethodCalls.define=function (callDefinition) {
	OSF.DDA.SyncMethodCalls[callDefinition.method.id]=OSF.DDA.SyncMethodCallFactory.manufacture(callDefinition);
};
OSF.DDA.ListType=(function () {
	var listTypes={};
	return {
		setListType: function OSF_DDA_ListType$AddListType(t, prop) { listTypes[t]=prop; },
		isListType: function OSF_DDA_ListType$IsListType(t) { return OSF.OUtil.listContainsKey(listTypes, t); },
		getDescriptor: function OSF_DDA_ListType$getDescriptor(t) { return listTypes[t]; }
	};
})();
OSF.DDA.HostParameterMap=function (specialProcessor, mappings) {
	var toHostMap="toHost";
	var fromHostMap="fromHost";
	var sourceData="sourceData";
	var self="self";
	var dynamicTypes={};
	dynamicTypes[Microsoft.Office.WebExtension.Parameters.Data]={
		toHost: function (data) {
			if (data !=null && data.rows !==undefined) {
				var tableData={};
				tableData[OSF.DDA.TableDataProperties.TableRows]=data.rows;
				tableData[OSF.DDA.TableDataProperties.TableHeaders]=data.headers;
				data=tableData;
			}
			return data;
		},
		fromHost: function (args) {
			return args;
		}
	};
	dynamicTypes[Microsoft.Office.WebExtension.Parameters.SampleData]=dynamicTypes[Microsoft.Office.WebExtension.Parameters.Data];
	function mapValues(preimageSet, mapping) {
		var ret=preimageSet ? {} : undefined;
		for (var entry in preimageSet) {
			var preimage=preimageSet[entry];
			var image;
			if (OSF.DDA.ListType.isListType(entry)) {
				image=[];
				for (var subEntry in preimage) {
					image.push(mapValues(preimage[subEntry], mapping));
				}
			}
			else if (OSF.OUtil.listContainsKey(dynamicTypes, entry)) {
				image=dynamicTypes[entry][mapping](preimage);
			}
			else if (mapping==fromHostMap && specialProcessor.preserveNesting(entry)) {
				image=mapValues(preimage, mapping);
			}
			else {
				var maps=mappings[entry];
				if (maps) {
					var map=maps[mapping];
					if (map) {
						image=map[preimage];
						if (image===undefined) {
							image=preimage;
						}
					}
				}
				else {
					image=preimage;
				}
			}
			ret[entry]=image;
		}
		return ret;
	}
	;
	function generateArguments(imageSet, parameters) {
		var ret;
		for (var param in parameters) {
			var arg;
			if (specialProcessor.isComplexType(param)) {
				arg=generateArguments(imageSet, mappings[param][toHostMap]);
			}
			else {
				arg=imageSet[param];
			}
			if (arg !=undefined) {
				if (!ret) {
					ret={};
				}
				var index=parameters[param];
				if (index==self) {
					index=param;
				}
				ret[index]=specialProcessor.pack(param, arg);
			}
		}
		return ret;
	}
	;
	function extractArguments(source, parameters, extracted) {
		if (!extracted) {
			extracted={};
		}
		for (var param in parameters) {
			var index=parameters[param];
			var value;
			if (index==self) {
				value=source;
			}
			else if (index==sourceData) {
				extracted[param]=source.toArray();
				continue;
			}
			else {
				value=source[index];
			}
			if (value===null || value===undefined) {
				extracted[param]=undefined;
			}
			else {
				value=specialProcessor.unpack(param, value);
				var map;
				if (specialProcessor.isComplexType(param)) {
					map=mappings[param][fromHostMap];
					if (specialProcessor.preserveNesting(param)) {
						extracted[param]=extractArguments(value, map);
					}
					else {
						extractArguments(value, map, extracted);
					}
				}
				else {
					if (OSF.DDA.ListType.isListType(param)) {
						map={};
						var entryDescriptor=OSF.DDA.ListType.getDescriptor(param);
						map[entryDescriptor]=self;
						var extractedValues=new Array(value.length);
						for (var item in value) {
							extractedValues[item]=extractArguments(value[item], map);
						}
						extracted[param]=extractedValues;
					}
					else {
						extracted[param]=value;
					}
				}
			}
		}
		return extracted;
	}
	;
	function applyMap(mapName, preimage, mapping) {
		var parameters=mappings[mapName][mapping];
		var image;
		if (mapping=="toHost") {
			var imageSet=mapValues(preimage, mapping);
			image=generateArguments(imageSet, parameters);
		}
		else if (mapping=="fromHost") {
			var argumentSet=extractArguments(preimage, parameters);
			image=mapValues(argumentSet, mapping);
		}
		return image;
	}
	;
	if (!mappings) {
		mappings={};
	}
	this.addMapping=function (mapName, description) {
		var toHost, fromHost;
		if (description.map) {
			toHost=description.map;
			fromHost={};
			for (var preimage in toHost) {
				var image=toHost[preimage];
				if (image==self) {
					image=preimage;
				}
				fromHost[image]=preimage;
			}
		}
		else {
			toHost=description.toHost;
			fromHost=description.fromHost;
		}
		var pair=mappings[mapName];
		if (pair) {
			var currMap=pair[toHostMap];
			for (var th in currMap)
				toHost[th]=currMap[th];
			currMap=pair[fromHostMap];
			for (var fh in currMap)
				fromHost[fh]=currMap[fh];
		}
		else {
			pair=mappings[mapName]={};
		}
		pair[toHostMap]=toHost;
		pair[fromHostMap]=fromHost;
	};
	this.toHost=function (mapName, preimage) { return applyMap(mapName, preimage, toHostMap); };
	this.fromHost=function (mapName, image) { return applyMap(mapName, image, fromHostMap); };
	this.self=self;
	this.sourceData=sourceData;
	this.addComplexType=function (ct) { specialProcessor.addComplexType(ct); };
	this.getDynamicType=function (dt) { return specialProcessor.getDynamicType(dt); };
	this.setDynamicType=function (dt, handler) { specialProcessor.setDynamicType(dt, handler); };
	this.dynamicTypes=dynamicTypes;
	this.doMapValues=function (preimageSet, mapping) { return mapValues(preimageSet, mapping); };
};
OSF.DDA.SpecialProcessor=function (complexTypes, dynamicTypes) {
	this.addComplexType=function OSF_DDA_SpecialProcessor$addComplexType(ct) {
		complexTypes.push(ct);
	};
	this.getDynamicType=function OSF_DDA_SpecialProcessor$getDynamicType(dt) {
		return dynamicTypes[dt];
	};
	this.setDynamicType=function OSF_DDA_SpecialProcessor$setDynamicType(dt, handler) {
		dynamicTypes[dt]=handler;
	};
	this.isComplexType=function OSF_DDA_SpecialProcessor$isComplexType(t) {
		return OSF.OUtil.listContainsValue(complexTypes, t);
	};
	this.isDynamicType=function OSF_DDA_SpecialProcessor$isDynamicType(p) {
		return OSF.OUtil.listContainsKey(dynamicTypes, p);
	};
	this.preserveNesting=function OSF_DDA_SpecialProcessor$preserveNesting(p) {
		var pn=[];
		if (OSF.DDA.PropertyDescriptors)
			pn.push(OSF.DDA.PropertyDescriptors.Subset);
		if (OSF.DDA.DataNodeEventProperties) {
			pn=pn.concat([
				OSF.DDA.DataNodeEventProperties.OldNode,
				OSF.DDA.DataNodeEventProperties.NewNode,
				OSF.DDA.DataNodeEventProperties.NextSiblingNode
			]);
		}
		return OSF.OUtil.listContainsValue(pn, p);
	};
	this.pack=function OSF_DDA_SpecialProcessor$pack(param, arg) {
		var value;
		if (this.isDynamicType(param)) {
			value=dynamicTypes[param].toHost(arg);
		}
		else {
			value=arg;
		}
		return value;
	};
	this.unpack=function OSF_DDA_SpecialProcessor$unpack(param, arg) {
		var value;
		if (this.isDynamicType(param)) {
			value=dynamicTypes[param].fromHost(arg);
		}
		else {
			value=arg;
		}
		return value;
	};
};
OSF.DDA.getDecoratedParameterMap=function (specialProcessor, initialDefs) {
	var parameterMap=new OSF.DDA.HostParameterMap(specialProcessor);
	var self=parameterMap.self;
	function createObject(properties) {
		var obj=null;
		if (properties) {
			obj={};
			var len=properties.length;
			for (var i=0; i < len; i++) {
				obj[properties[i].name]=properties[i].value;
			}
		}
		return obj;
	}
	parameterMap.define=function define(definition) {
		var args={};
		var toHost=createObject(definition.toHost);
		if (definition.invertible) {
			args.map=toHost;
		}
		else if (definition.canonical) {
			args.toHost=args.fromHost=toHost;
		}
		else {
			args.toHost=toHost;
			args.fromHost=createObject(definition.fromHost);
		}
		parameterMap.addMapping(definition.type, args);
		if (definition.isComplexType)
			parameterMap.addComplexType(definition.type);
	};
	for (var id in initialDefs)
		parameterMap.define(initialDefs[id]);
	return parameterMap;
};
OSF.OUtil.setNamespace("DispIdHost", OSF.DDA);
OSF.DDA.DispIdHost.Methods={
	InvokeMethod: "invokeMethod",
	AddEventHandler: "addEventHandler",
	RemoveEventHandler: "removeEventHandler",
	OpenDialog: "openDialog",
	CloseDialog: "closeDialog",
	MessageParent: "messageParent",
	SendMessage: "sendMessage"
};
OSF.DDA.DispIdHost.Delegates={
	ExecuteAsync: "executeAsync",
	RegisterEventAsync: "registerEventAsync",
	UnregisterEventAsync: "unregisterEventAsync",
	ParameterMap: "parameterMap",
	OpenDialog: "openDialog",
	CloseDialog: "closeDialog",
	MessageParent: "messageParent",
	SendMessage: "sendMessage"
};
OSF.DDA.DispIdHost.Facade=function OSF_DDA_DispIdHost_Facade(getDelegateMethods, parameterMap) {
	var dispIdMap={};
	var jsom=OSF.DDA.AsyncMethodNames;
	var did=OSF.DDA.MethodDispId;
	var methodMap={
		"GoToByIdAsync": did.dispidNavigateToMethod,
		"GetSelectedDataAsync": did.dispidGetSelectedDataMethod,
		"SetSelectedDataAsync": did.dispidSetSelectedDataMethod,
		"GetDocumentCopyChunkAsync": did.dispidGetDocumentCopyChunkMethod,
		"ReleaseDocumentCopyAsync": did.dispidReleaseDocumentCopyMethod,
		"GetDocumentCopyAsync": did.dispidGetDocumentCopyMethod,
		"AddFromSelectionAsync": did.dispidAddBindingFromSelectionMethod,
		"AddFromPromptAsync": did.dispidAddBindingFromPromptMethod,
		"AddFromNamedItemAsync": did.dispidAddBindingFromNamedItemMethod,
		"GetAllAsync": did.dispidGetAllBindingsMethod,
		"GetByIdAsync": did.dispidGetBindingMethod,
		"ReleaseByIdAsync": did.dispidReleaseBindingMethod,
		"GetDataAsync": did.dispidGetBindingDataMethod,
		"SetDataAsync": did.dispidSetBindingDataMethod,
		"AddRowsAsync": did.dispidAddRowsMethod,
		"AddColumnsAsync": did.dispidAddColumnsMethod,
		"DeleteAllDataValuesAsync": did.dispidClearAllRowsMethod,
		"RefreshAsync": did.dispidLoadSettingsMethod,
		"SaveAsync": did.dispidSaveSettingsMethod,
		"GetActiveViewAsync": did.dispidGetActiveViewMethod,
		"GetFilePropertiesAsync": did.dispidGetFilePropertiesMethod,
		"GetOfficeThemeAsync": did.dispidGetOfficeThemeMethod,
		"GetDocumentThemeAsync": did.dispidGetDocumentThemeMethod,
		"ClearFormatsAsync": did.dispidClearFormatsMethod,
		"SetTableOptionsAsync": did.dispidSetTableOptionsMethod,
		"SetFormatsAsync": did.dispidSetFormatsMethod,
		"GetAccessTokenAsync": did.dispidGetAccessTokenMethod,
		"ExecuteRichApiRequestAsync": did.dispidExecuteRichApiRequestMethod,
		"AppCommandInvocationCompletedAsync": did.dispidAppCommandInvocationCompletedMethod,
		"CloseContainerAsync": did.dispidCloseContainerMethod,
		"AddDataPartAsync": did.dispidAddDataPartMethod,
		"GetDataPartByIdAsync": did.dispidGetDataPartByIdMethod,
		"GetDataPartsByNameSpaceAsync": did.dispidGetDataPartsByNamespaceMethod,
		"GetPartXmlAsync": did.dispidGetDataPartXmlMethod,
		"GetPartNodesAsync": did.dispidGetDataPartNodesMethod,
		"DeleteDataPartAsync": did.dispidDeleteDataPartMethod,
		"GetNodeValueAsync": did.dispidGetDataNodeValueMethod,
		"GetNodeXmlAsync": did.dispidGetDataNodeXmlMethod,
		"GetRelativeNodesAsync": did.dispidGetDataNodesMethod,
		"SetNodeValueAsync": did.dispidSetDataNodeValueMethod,
		"SetNodeXmlAsync": did.dispidSetDataNodeXmlMethod,
		"AddDataPartNamespaceAsync": did.dispidAddDataNamespaceMethod,
		"GetDataPartNamespaceAsync": did.dispidGetDataUriByPrefixMethod,
		"GetDataPartPrefixAsync": did.dispidGetDataPrefixByUriMethod,
		"GetNodeTextAsync": did.dispidGetDataNodeTextMethod,
		"SetNodeTextAsync": did.dispidSetDataNodeTextMethod,
		"GetSelectedTask": did.dispidGetSelectedTaskMethod,
		"GetTask": did.dispidGetTaskMethod,
		"GetWSSUrl": did.dispidGetWSSUrlMethod,
		"GetTaskField": did.dispidGetTaskFieldMethod,
		"GetSelectedResource": did.dispidGetSelectedResourceMethod,
		"GetResourceField": did.dispidGetResourceFieldMethod,
		"GetProjectField": did.dispidGetProjectFieldMethod,
		"GetSelectedView": did.dispidGetSelectedViewMethod,
		"GetTaskByIndex": did.dispidGetTaskByIndexMethod,
		"GetResourceByIndex": did.dispidGetResourceByIndexMethod,
		"SetTaskField": did.dispidSetTaskFieldMethod,
		"SetResourceField": did.dispidSetResourceFieldMethod,
		"GetMaxTaskIndex": did.dispidGetMaxTaskIndexMethod,
		"GetMaxResourceIndex": did.dispidGetMaxResourceIndexMethod,
		"CreateTask": did.dispidCreateTaskMethod
	};
	for (var method in methodMap) {
		if (jsom[method]) {
			dispIdMap[jsom[method].id]=methodMap[method];
		}
	}
	jsom=OSF.DDA.SyncMethodNames;
	did=OSF.DDA.MethodDispId;
	var asyncMethodMap={
		"MessageParent": did.dispidMessageParentMethod,
		"SendMessage": did.dispidSendMessageMethod
	};
	for (var method in asyncMethodMap) {
		if (jsom[method]) {
			dispIdMap[jsom[method].id]=asyncMethodMap[method];
		}
	}
	jsom=Microsoft.Office.WebExtension.EventType;
	did=OSF.DDA.EventDispId;
	var eventMap={
		"SettingsChanged": did.dispidSettingsChangedEvent,
		"DocumentSelectionChanged": did.dispidDocumentSelectionChangedEvent,
		"BindingSelectionChanged": did.dispidBindingSelectionChangedEvent,
		"BindingDataChanged": did.dispidBindingDataChangedEvent,
		"ActiveViewChanged": did.dispidActiveViewChangedEvent,
		"OfficeThemeChanged": did.dispidOfficeThemeChangedEvent,
		"DocumentThemeChanged": did.dispidDocumentThemeChangedEvent,
		"AppCommandInvoked": did.dispidAppCommandInvokedEvent,
		"DialogMessageReceived": did.dispidDialogMessageReceivedEvent,
		"DialogParentMessageReceived": did.dispidDialogParentMessageReceivedEvent,
		"ObjectDeleted": did.dispidObjectDeletedEvent,
		"ObjectSelectionChanged": did.dispidObjectSelectionChangedEvent,
		"ObjectDataChanged": did.dispidObjectDataChangedEvent,
		"ContentControlAdded": did.dispidContentControlAddedEvent,
		"RichApiMessage": did.dispidRichApiMessageEvent,
		"ItemChanged": did.dispidOlkItemSelectedChangedEvent,
		"RecipientsChanged": did.dispidOlkRecipientsChangedEvent,
		"AppointmentTimeChanged": did.dispidOlkAppointmentTimeChangedEvent,
		"TaskSelectionChanged": did.dispidTaskSelectionChangedEvent,
		"ResourceSelectionChanged": did.dispidResourceSelectionChangedEvent,
		"ViewSelectionChanged": did.dispidViewSelectionChangedEvent,
		"DataNodeInserted": did.dispidDataNodeAddedEvent,
		"DataNodeReplaced": did.dispidDataNodeReplacedEvent,
		"DataNodeDeleted": did.dispidDataNodeDeletedEvent
	};
	for (var event in eventMap) {
		if (jsom[event]) {
			dispIdMap[jsom[event]]=eventMap[event];
		}
	}
	function IsObjectEvent(dispId) {
		return (dispId==OSF.DDA.EventDispId.dispidObjectDeletedEvent ||
			dispId==OSF.DDA.EventDispId.dispidObjectSelectionChangedEvent ||
			dispId==OSF.DDA.EventDispId.dispidObjectDataChangedEvent ||
			dispId==OSF.DDA.EventDispId.dispidContentControlAddedEvent);
	}
	function onException(ex, asyncMethodCall, suppliedArgs, callArgs) {
		if (typeof ex=="number") {
			if (!callArgs) {
				callArgs=asyncMethodCall.getCallArgs(suppliedArgs);
			}
			OSF.DDA.issueAsyncResult(callArgs, ex, OSF.DDA.ErrorCodeManager.getErrorArgs(ex));
		}
		else {
			throw ex;
		}
	}
	;
	this[OSF.DDA.DispIdHost.Methods.InvokeMethod]=function OSF_DDA_DispIdHost_Facade$InvokeMethod(method, suppliedArguments, caller, privateState) {
		var callArgs;
		try {
			var methodName=method.id;
			var asyncMethodCall=OSF.DDA.AsyncMethodCalls[methodName];
			callArgs=asyncMethodCall.verifyAndExtractCall(suppliedArguments, caller, privateState);
			var dispId=dispIdMap[methodName];
			var delegate=getDelegateMethods(methodName);
			var richApiInExcelMethodSubstitution=null;
			if (window.Excel && window.Office.context.requirements.isSetSupported("RedirectV1Api")) {
				window.Excel._RedirectV1APIs=true;
			}
			if (window.Excel && window.Excel._RedirectV1APIs && (richApiInExcelMethodSubstitution=window.Excel._V1APIMap[methodName])) {
				var preprocessedCallArgs=OSF.OUtil.shallowCopy(callArgs);
				delete preprocessedCallArgs[Microsoft.Office.WebExtension.Parameters.AsyncContext];
				if (richApiInExcelMethodSubstitution.preprocess) {
					preprocessedCallArgs=richApiInExcelMethodSubstitution.preprocess(preprocessedCallArgs);
				}
				var ctx=new window.Excel.RequestContext();
				var result=richApiInExcelMethodSubstitution.call(ctx, preprocessedCallArgs);
				ctx.sync()
					.then(function () {
					var response=result.value;
					var status=response.status;
					delete response["status"];
					delete response["@odata.type"];
					if (richApiInExcelMethodSubstitution.postprocess) {
						response=richApiInExcelMethodSubstitution.postprocess(response, preprocessedCallArgs);
					}
					if (status !=0) {
						response=OSF.DDA.ErrorCodeManager.getErrorArgs(status);
					}
					OSF.DDA.issueAsyncResult(callArgs, status, response);
				})["catch"](function (error) {
					OSF.DDA.issueAsyncResult(callArgs, OSF.DDA.ErrorCodeManager.errorCodes.ooeFailure, null);
				});
			}
			else {
				var hostCallArgs;
				if (parameterMap.toHost) {
					hostCallArgs=parameterMap.toHost(dispId, callArgs);
				}
				else {
					hostCallArgs=callArgs;
				}
				var startTime=(new Date()).getTime();
				delegate[OSF.DDA.DispIdHost.Delegates.ExecuteAsync]({
					"dispId": dispId,
					"hostCallArgs": hostCallArgs,
					"onCalling": function OSF_DDA_DispIdFacade$Execute_onCalling() { },
					"onReceiving": function OSF_DDA_DispIdFacade$Execute_onReceiving() { },
					"onComplete": function (status, hostResponseArgs) {
						var responseArgs;
						if (status==OSF.DDA.ErrorCodeManager.errorCodes.ooeSuccess) {
							if (parameterMap.fromHost) {
								responseArgs=parameterMap.fromHost(dispId, hostResponseArgs);
							}
							else {
								responseArgs=hostResponseArgs;
							}
						}
						else {
							responseArgs=hostResponseArgs;
						}
						var payload=asyncMethodCall.processResponse(status, responseArgs, caller, callArgs);
						OSF.DDA.issueAsyncResult(callArgs, status, payload);
						if (OSF.AppTelemetry) {
							OSF.AppTelemetry.onMethodDone(dispId, hostCallArgs, Math.abs((new Date()).getTime() - startTime), status);
						}
					}
				});
			}
		}
		catch (ex) {
			onException(ex, asyncMethodCall, suppliedArguments, callArgs);
		}
	};
	this[OSF.DDA.DispIdHost.Methods.AddEventHandler]=function OSF_DDA_DispIdHost_Facade$AddEventHandler(suppliedArguments, eventDispatch, caller, isPopupWindow) {
		var callArgs;
		var eventType, handler;
		var isObjectEvent=false;
		function onEnsureRegistration(status) {
			if (status==OSF.DDA.ErrorCodeManager.errorCodes.ooeSuccess) {
				var added=!isObjectEvent ? eventDispatch.addEventHandler(eventType, handler) :
					eventDispatch.addObjectEventHandler(eventType, callArgs[Microsoft.Office.WebExtension.Parameters.Id], handler);
				if (!added) {
					status=OSF.DDA.ErrorCodeManager.errorCodes.ooeEventHandlerAdditionFailed;
				}
			}
			var error;
			if (status !=OSF.DDA.ErrorCodeManager.errorCodes.ooeSuccess) {
				error=OSF.DDA.ErrorCodeManager.getErrorArgs(status);
			}
			OSF.DDA.issueAsyncResult(callArgs, status, error);
		}
		try {
			var asyncMethodCall=OSF.DDA.AsyncMethodCalls[OSF.DDA.AsyncMethodNames.AddHandlerAsync.id];
			callArgs=asyncMethodCall.verifyAndExtractCall(suppliedArguments, caller, eventDispatch);
			eventType=callArgs[Microsoft.Office.WebExtension.Parameters.EventType];
			handler=callArgs[Microsoft.Office.WebExtension.Parameters.Handler];
			if (isPopupWindow) {
				onEnsureRegistration(OSF.DDA.ErrorCodeManager.errorCodes.ooeSuccess);
				return;
			}
			var dispId=dispIdMap[eventType];
			isObjectEvent=IsObjectEvent(dispId);
			var targetId=(isObjectEvent ? callArgs[Microsoft.Office.WebExtension.Parameters.Id] : (caller.id || ""));
			var count=isObjectEvent ? eventDispatch.getObjectEventHandlerCount(eventType, targetId) : eventDispatch.getEventHandlerCount(eventType);
			if (count==0) {
				var invoker=getDelegateMethods(eventType)[OSF.DDA.DispIdHost.Delegates.RegisterEventAsync];
				invoker({
					"eventType": eventType,
					"dispId": dispId,
					"targetId": targetId,
					"onCalling": function OSF_DDA_DispIdFacade$Execute_onCalling() { OSF.OUtil.writeProfilerMark(OSF.HostCallPerfMarker.IssueCall); },
					"onReceiving": function OSF_DDA_DispIdFacade$Execute_onReceiving() { OSF.OUtil.writeProfilerMark(OSF.HostCallPerfMarker.ReceiveResponse); },
					"onComplete": onEnsureRegistration,
					"onEvent": function handleEvent(hostArgs) {
						var args=parameterMap.fromHost(dispId, hostArgs);
						if (!isObjectEvent)
							eventDispatch.fireEvent(OSF.DDA.OMFactory.manufactureEventArgs(eventType, caller, args));
						else
							eventDispatch.fireObjectEvent(targetId, OSF.DDA.OMFactory.manufactureEventArgs(eventType, targetId, args));
					}
				});
			}
			else {
				onEnsureRegistration(OSF.DDA.ErrorCodeManager.errorCodes.ooeSuccess);
			}
		}
		catch (ex) {
			onException(ex, asyncMethodCall, suppliedArguments, callArgs);
		}
	};
	this[OSF.DDA.DispIdHost.Methods.RemoveEventHandler]=function OSF_DDA_DispIdHost_Facade$RemoveEventHandler(suppliedArguments, eventDispatch, caller) {
		var callArgs;
		var eventType, handler;
		var isObjectEvent=false;
		function onEnsureRegistration(status) {
			var error;
			if (status !=OSF.DDA.ErrorCodeManager.errorCodes.ooeSuccess) {
				error=OSF.DDA.ErrorCodeManager.getErrorArgs(status);
			}
			OSF.DDA.issueAsyncResult(callArgs, status, error);
		}
		try {
			var asyncMethodCall=OSF.DDA.AsyncMethodCalls[OSF.DDA.AsyncMethodNames.RemoveHandlerAsync.id];
			callArgs=asyncMethodCall.verifyAndExtractCall(suppliedArguments, caller, eventDispatch);
			eventType=callArgs[Microsoft.Office.WebExtension.Parameters.EventType];
			handler=callArgs[Microsoft.Office.WebExtension.Parameters.Handler];
			var dispId=dispIdMap[eventType];
			isObjectEvent=IsObjectEvent(dispId);
			var targetId=(isObjectEvent ? callArgs[Microsoft.Office.WebExtension.Parameters.Id] : (caller.id || ""));
			var status, removeSuccess;
			if (handler===null) {
				removeSuccess=isObjectEvent ? eventDispatch.clearObjectEventHandlers(eventType, targetId) : eventDispatch.clearEventHandlers(eventType);
				status=OSF.DDA.ErrorCodeManager.errorCodes.ooeSuccess;
			}
			else {
				removeSuccess=isObjectEvent ? eventDispatch.removeObjectEventHandler(eventType, targetId, handler) : eventDispatch.removeEventHandler(eventType, handler);
				status=removeSuccess ? OSF.DDA.ErrorCodeManager.errorCodes.ooeSuccess : OSF.DDA.ErrorCodeManager.errorCodes.ooeEventHandlerNotExist;
			}
			var count=isObjectEvent ? eventDispatch.getObjectEventHandlerCount(eventType, targetId) : eventDispatch.getEventHandlerCount(eventType);
			if (removeSuccess && count==0) {
				var invoker=getDelegateMethods(eventType)[OSF.DDA.DispIdHost.Delegates.UnregisterEventAsync];
				invoker({
					"eventType": eventType,
					"dispId": dispId,
					"targetId": targetId,
					"onCalling": function OSF_DDA_DispIdFacade$Execute_onCalling() { OSF.OUtil.writeProfilerMark(OSF.HostCallPerfMarker.IssueCall); },
					"onReceiving": function OSF_DDA_DispIdFacade$Execute_onReceiving() { OSF.OUtil.writeProfilerMark(OSF.HostCallPerfMarker.ReceiveResponse); },
					"onComplete": onEnsureRegistration
				});
			}
			else {
				onEnsureRegistration(status);
			}
		}
		catch (ex) {
			onException(ex, asyncMethodCall, suppliedArguments, callArgs);
		}
	};
	this[OSF.DDA.DispIdHost.Methods.OpenDialog]=function OSF_DDA_DispIdHost_Facade$OpenDialog(suppliedArguments, eventDispatch, caller) {
		var callArgs;
		var targetId;
		var dialogMessageEvent=Microsoft.Office.WebExtension.EventType.DialogMessageReceived;
		var dialogOtherEvent=Microsoft.Office.WebExtension.EventType.DialogEventReceived;
		function onEnsureRegistration(status) {
			var payload;
			if (status !=OSF.DDA.ErrorCodeManager.errorCodes.ooeSuccess) {
				payload=OSF.DDA.ErrorCodeManager.getErrorArgs(status);
			}
			else {
				var onSucceedArgs={};
				onSucceedArgs[Microsoft.Office.WebExtension.Parameters.Id]=targetId;
				onSucceedArgs[Microsoft.Office.WebExtension.Parameters.Data]=eventDispatch;
				var payload=asyncMethodCall.processResponse(status, onSucceedArgs, caller, callArgs);
				OSF.DialogShownStatus.hasDialogShown=true;
				eventDispatch.clearEventHandlers(dialogMessageEvent);
				eventDispatch.clearEventHandlers(dialogOtherEvent);
			}
			OSF.DDA.issueAsyncResult(callArgs, status, payload);
		}
		try {
			if (dialogMessageEvent==undefined || dialogOtherEvent==undefined) {
				onEnsureRegistration(OSF.DDA.ErrorCodeManager.ooeOperationNotSupported);
			}
			if (OSF.DDA.AsyncMethodNames.DisplayDialogAsync==null) {
				onEnsureRegistration(OSF.DDA.ErrorCodeManager.errorCodes.ooeInternalError);
				return;
			}
			var asyncMethodCall=OSF.DDA.AsyncMethodCalls[OSF.DDA.AsyncMethodNames.DisplayDialogAsync.id];
			callArgs=asyncMethodCall.verifyAndExtractCall(suppliedArguments, caller, eventDispatch);
			var dispId=dispIdMap[dialogMessageEvent];
			var delegateMethods=getDelegateMethods(dialogMessageEvent);
			var invoker=delegateMethods[OSF.DDA.DispIdHost.Delegates.OpenDialog] !=undefined ?
				delegateMethods[OSF.DDA.DispIdHost.Delegates.OpenDialog] :
				delegateMethods[OSF.DDA.DispIdHost.Delegates.RegisterEventAsync];
			targetId=JSON.stringify(callArgs);
			if (!OSF.DialogShownStatus.hasDialogShown) {
				eventDispatch.clearQueuedEvent(dialogMessageEvent);
				eventDispatch.clearQueuedEvent(dialogOtherEvent);
				eventDispatch.clearQueuedEvent(Microsoft.Office.WebExtension.EventType.DialogParentMessageReceived);
			}
			invoker({
				"eventType": dialogMessageEvent,
				"dispId": dispId,
				"targetId": targetId,
				"onCalling": function OSF_DDA_DispIdFacade$Execute_onCalling() { OSF.OUtil.writeProfilerMark(OSF.HostCallPerfMarker.IssueCall); },
				"onReceiving": function OSF_DDA_DispIdFacade$Execute_onReceiving() { OSF.OUtil.writeProfilerMark(OSF.HostCallPerfMarker.ReceiveResponse); },
				"onComplete": onEnsureRegistration,
				"onEvent": function handleEvent(hostArgs) {
					var args=parameterMap.fromHost(dispId, hostArgs);
					var event=OSF.DDA.OMFactory.manufactureEventArgs(dialogMessageEvent, caller, args);
					if (event.type==dialogOtherEvent) {
						var payload=OSF.DDA.ErrorCodeManager.getErrorArgs(event.error);
						var errorArgs={};
						errorArgs[OSF.DDA.AsyncResultEnum.ErrorProperties.Code]=status || OSF.DDA.ErrorCodeManager.errorCodes.ooeInternalError;
						errorArgs[OSF.DDA.AsyncResultEnum.ErrorProperties.Name]=payload.name || payload;
						errorArgs[OSF.DDA.AsyncResultEnum.ErrorProperties.Message]=payload.message || payload;
						event.error=new OSF.DDA.Error(errorArgs[OSF.DDA.AsyncResultEnum.ErrorProperties.Name], errorArgs[OSF.DDA.AsyncResultEnum.ErrorProperties.Message], errorArgs[OSF.DDA.AsyncResultEnum.ErrorProperties.Code]);
					}
					eventDispatch.fireOrQueueEvent(event);
					if (args[OSF.DDA.PropertyDescriptors.MessageType]==OSF.DialogMessageType.DialogClosed) {
						eventDispatch.clearEventHandlers(dialogMessageEvent);
						eventDispatch.clearEventHandlers(dialogOtherEvent);
						eventDispatch.clearEventHandlers(Microsoft.Office.WebExtension.EventType.DialogParentMessageReceived);
						OSF.DialogShownStatus.hasDialogShown=false;
					}
				}
			});
		}
		catch (ex) {
			onException(ex, asyncMethodCall, suppliedArguments, callArgs);
		}
	};
	this[OSF.DDA.DispIdHost.Methods.CloseDialog]=function OSF_DDA_DispIdHost_Facade$CloseDialog(suppliedArguments, targetId, eventDispatch, caller) {
		var callArgs;
		var dialogMessageEvent, dialogOtherEvent;
		var closeStatus=OSF.DDA.ErrorCodeManager.errorCodes.ooeSuccess;
		function closeCallback(status) {
			closeStatus=status;
			OSF.DialogShownStatus.hasDialogShown=false;
		}
		try {
			var asyncMethodCall=OSF.DDA.AsyncMethodCalls[OSF.DDA.AsyncMethodNames.CloseAsync.id];
			callArgs=asyncMethodCall.verifyAndExtractCall(suppliedArguments, caller, eventDispatch);
			dialogMessageEvent=Microsoft.Office.WebExtension.EventType.DialogMessageReceived;
			dialogOtherEvent=Microsoft.Office.WebExtension.EventType.DialogEventReceived;
			eventDispatch.clearEventHandlers(dialogMessageEvent);
			eventDispatch.clearEventHandlers(dialogOtherEvent);
			var dispId=dispIdMap[dialogMessageEvent];
			var delegateMethods=getDelegateMethods(dialogMessageEvent);
			var invoker=delegateMethods[OSF.DDA.DispIdHost.Delegates.CloseDialog] !=undefined ?
				delegateMethods[OSF.DDA.DispIdHost.Delegates.CloseDialog] :
				delegateMethods[OSF.DDA.DispIdHost.Delegates.UnregisterEventAsync];
			invoker({
				"eventType": dialogMessageEvent,
				"dispId": dispId,
				"targetId": targetId,
				"onCalling": function OSF_DDA_DispIdFacade$Execute_onCalling() { OSF.OUtil.writeProfilerMark(OSF.HostCallPerfMarker.IssueCall); },
				"onReceiving": function OSF_DDA_DispIdFacade$Execute_onReceiving() { OSF.OUtil.writeProfilerMark(OSF.HostCallPerfMarker.ReceiveResponse); },
				"onComplete": closeCallback
			});
		}
		catch (ex) {
			onException(ex, asyncMethodCall, suppliedArguments, callArgs);
		}
		if (closeStatus !=OSF.DDA.ErrorCodeManager.errorCodes.ooeSuccess) {
			throw OSF.OUtil.formatString(Strings.OfficeOM.L_FunctionCallFailed, OSF.DDA.AsyncMethodNames.CloseAsync.displayName, closeStatus);
		}
	};
	this[OSF.DDA.DispIdHost.Methods.MessageParent]=function OSF_DDA_DispIdHost_Facade$MessageParent(suppliedArguments, caller) {
		var stateInfo={};
		var syncMethodCall=OSF.DDA.SyncMethodCalls[OSF.DDA.SyncMethodNames.MessageParent.id];
		var callArgs=syncMethodCall.verifyAndExtractCall(suppliedArguments, caller, stateInfo);
		var delegate=getDelegateMethods(OSF.DDA.SyncMethodNames.MessageParent.id);
		var invoker=delegate[OSF.DDA.DispIdHost.Delegates.MessageParent];
		var dispId=dispIdMap[OSF.DDA.SyncMethodNames.MessageParent.id];
		return invoker({
			"dispId": dispId,
			"hostCallArgs": callArgs,
			"onCalling": function OSF_DDA_DispIdFacade$Execute_onCalling() { OSF.OUtil.writeProfilerMark(OSF.HostCallPerfMarker.IssueCall); },
			"onReceiving": function OSF_DDA_DispIdFacade$Execute_onReceiving() { OSF.OUtil.writeProfilerMark(OSF.HostCallPerfMarker.ReceiveResponse); }
		});
	};
	this[OSF.DDA.DispIdHost.Methods.SendMessage]=function OSF_DDA_DispIdHost_Facade$SendMessage(suppliedArguments, eventDispatch, caller) {
		var stateInfo={};
		var syncMethodCall=OSF.DDA.SyncMethodCalls[OSF.DDA.SyncMethodNames.SendMessage.id];
		var callArgs=syncMethodCall.verifyAndExtractCall(suppliedArguments, caller, stateInfo);
		var delegate=getDelegateMethods(OSF.DDA.SyncMethodNames.SendMessage.id);
		var invoker=delegate[OSF.DDA.DispIdHost.Delegates.SendMessage];
		var dispId=dispIdMap[OSF.DDA.SyncMethodNames.SendMessage.id];
		return invoker({
			"dispId": dispId,
			"hostCallArgs": callArgs,
			"onCalling": function OSF_DDA_DispIdFacade$Execute_onCalling() { OSF.OUtil.writeProfilerMark(OSF.HostCallPerfMarker.IssueCall); },
			"onReceiving": function OSF_DDA_DispIdFacade$Execute_onReceiving() { OSF.OUtil.writeProfilerMark(OSF.HostCallPerfMarker.ReceiveResponse); }
		});
	};
};
OSF.DDA.DispIdHost.addAsyncMethods=function OSF_DDA_DispIdHost$AddAsyncMethods(target, asyncMethodNames, privateState) {
	for (var entry in asyncMethodNames) {
		var method=asyncMethodNames[entry];
		var name=method.displayName;
		if (!target[name]) {
			OSF.OUtil.defineEnumerableProperty(target, name, {
				value: (function (asyncMethod) {
					return function () {
						var invokeMethod=OSF._OfficeAppFactory.getHostFacade()[OSF.DDA.DispIdHost.Methods.InvokeMethod];
						invokeMethod(asyncMethod, arguments, target, privateState);
					};
				})(method)
			});
		}
	}
};
OSF.DDA.DispIdHost.addEventSupport=function OSF_DDA_DispIdHost$AddEventSupport(target, eventDispatch, isPopupWindow) {
	var add=OSF.DDA.AsyncMethodNames.AddHandlerAsync.displayName;
	var remove=OSF.DDA.AsyncMethodNames.RemoveHandlerAsync.displayName;
	if (!target[add]) {
		OSF.OUtil.defineEnumerableProperty(target, add, {
			value: function () {
				var addEventHandler=OSF._OfficeAppFactory.getHostFacade()[OSF.DDA.DispIdHost.Methods.AddEventHandler];
				addEventHandler(arguments, eventDispatch, target, isPopupWindow);
			}
		});
	}
	if (!target[remove]) {
		OSF.OUtil.defineEnumerableProperty(target, remove, {
			value: function () {
				var removeEventHandler=OSF._OfficeAppFactory.getHostFacade()[OSF.DDA.DispIdHost.Methods.RemoveEventHandler];
				removeEventHandler(arguments, eventDispatch, target);
			}
		});
	}
};
var OfficeExt;
(function (OfficeExt) {
	var MsAjaxTypeHelper=(function () {
		function MsAjaxTypeHelper() {
		}
		MsAjaxTypeHelper.isInstanceOfType=function (type, instance) {
			if (typeof (instance)==="undefined" || instance===null)
				return false;
			if (instance instanceof type)
				return true;
			var instanceType=instance.constructor;
			if (!instanceType || (typeof (instanceType) !=="function") || !instanceType.__typeName || instanceType.__typeName==='Object') {
				instanceType=Object;
			}
			return !!(instanceType===type) ||
				(instanceType.__typeName && type.__typeName && instanceType.__typeName===type.__typeName);
		};
		return MsAjaxTypeHelper;
	})();
	OfficeExt.MsAjaxTypeHelper=MsAjaxTypeHelper;
	var MsAjaxError=(function () {
		function MsAjaxError() {
		}
		MsAjaxError.create=function (message, errorInfo) {
			var err=new Error(message);
			err.message=message;
			if (errorInfo) {
				for (var v in errorInfo) {
					err[v]=errorInfo[v];
				}
			}
			err.popStackFrame();
			return err;
		};
		MsAjaxError.parameterCount=function (message) {
			var displayMessage="Sys.ParameterCountException: "+(message ? message : "Parameter count mismatch.");
			var err=MsAjaxError.create(displayMessage, { name: 'Sys.ParameterCountException' });
			err.popStackFrame();
			return err;
		};
		MsAjaxError.argument=function (paramName, message) {
			var displayMessage="Sys.ArgumentException: "+(message ? message : "Value does not fall within the expected range.");
			if (paramName) {
				displayMessage+="\n"+MsAjaxString.format("Parameter name: {0}", paramName);
			}
			var err=MsAjaxError.create(displayMessage, { name: "Sys.ArgumentException", paramName: paramName });
			err.popStackFrame();
			return err;
		};
		MsAjaxError.argumentNull=function (paramName, message) {
			var displayMessage="Sys.ArgumentNullException: "+(message ? message : "Value cannot be null.");
			if (paramName) {
				displayMessage+="\n"+MsAjaxString.format("Parameter name: {0}", paramName);
			}
			var err=MsAjaxError.create(displayMessage, { name: "Sys.ArgumentNullException", paramName: paramName });
			err.popStackFrame();
			return err;
		};
		MsAjaxError.argumentOutOfRange=function (paramName, actualValue, message) {
			var displayMessage="Sys.ArgumentOutOfRangeException: "+(message ? message : "Specified argument was out of the range of valid values.");
			if (paramName) {
				displayMessage+="\n"+MsAjaxString.format("Parameter name: {0}", paramName);
			}
			if (typeof (actualValue) !=="undefined" && actualValue !==null) {
				displayMessage+="\n"+MsAjaxString.format("Actual value was {0}.", actualValue);
			}
			var err=MsAjaxError.create(displayMessage, {
				name: "Sys.ArgumentOutOfRangeException",
				paramName: paramName,
				actualValue: actualValue
			});
			err.popStackFrame();
			return err;
		};
		MsAjaxError.argumentType=function (paramName, actualType, expectedType, message) {
			var displayMessage="Sys.ArgumentTypeException: ";
			if (message) {
				displayMessage+=message;
			}
			else if (actualType && expectedType) {
				displayMessage+=MsAjaxString.format("Object of type '{0}' cannot be converted to type '{1}'.", actualType.getName ? actualType.getName() : actualType, expectedType.getName ? expectedType.getName() : expectedType);
			}
			else {
				displayMessage+="Object cannot be converted to the required type.";
			}
			if (paramName) {
				displayMessage+="\n"+MsAjaxString.format("Parameter name: {0}", paramName);
			}
			var err=MsAjaxError.create(displayMessage, {
				name: "Sys.ArgumentTypeException",
				paramName: paramName,
				actualType: actualType,
				expectedType: expectedType
			});
			err.popStackFrame();
			return err;
		};
		MsAjaxError.argumentUndefined=function (paramName, message) {
			var displayMessage="Sys.ArgumentUndefinedException: "+(message ? message : "Value cannot be undefined.");
			if (paramName) {
				displayMessage+="\n"+MsAjaxString.format("Parameter name: {0}", paramName);
			}
			var err=MsAjaxError.create(displayMessage, { name: "Sys.ArgumentUndefinedException", paramName: paramName });
			err.popStackFrame();
			return err;
		};
		MsAjaxError.invalidOperation=function (message) {
			var displayMessage="Sys.InvalidOperationException: "+(message ? message : "Operation is not valid due to the current state of the object.");
			var err=MsAjaxError.create(displayMessage, { name: 'Sys.InvalidOperationException' });
			err.popStackFrame();
			return err;
		};
		return MsAjaxError;
	})();
	OfficeExt.MsAjaxError=MsAjaxError;
	var MsAjaxString=(function () {
		function MsAjaxString() {
		}
		MsAjaxString.format=function (format) {
			var args=[];
			for (var _i=1; _i < arguments.length; _i++) {
				args[_i - 1]=arguments[_i];
			}
			var source=format;
			return source.replace(/{(\d+)}/gm, function (match, number) {
				var index=parseInt(number, 10);
				return args[index]===undefined ? '{'+number+'}' : args[index];
			});
		};
		MsAjaxString.startsWith=function (str, prefix) {
			return (str.substr(0, prefix.length)===prefix);
		};
		return MsAjaxString;
	})();
	OfficeExt.MsAjaxString=MsAjaxString;
	var MsAjaxDebug=(function () {
		function MsAjaxDebug() {
		}
		MsAjaxDebug.trace=function (text) {
			if (typeof Debug !=="undefined" && Debug.writeln)
				Debug.writeln(text);
			if (window.console && window.console.log)
				window.console.log(text);
			if (window.opera && window.opera.postError)
				window.opera.postError(text);
			if (window.debugService && window.debugService.trace)
				window.debugService.trace(text);
			var a=document.getElementById("TraceConsole");
			if (a && a.tagName.toUpperCase()==="TEXTAREA") {
				a.innerHTML+=text+"\n";
			}
		};
		return MsAjaxDebug;
	})();
	OfficeExt.MsAjaxDebug=MsAjaxDebug;
	if (!OsfMsAjaxFactory.isMsAjaxLoaded()) {
		var registerTypeInternal=function registerTypeInternal(type, name, isClass) {
			if (type.__typeName===undefined) {
				type.__typeName=name;
			}
			if (type.__class===undefined) {
				type.__class=isClass;
			}
		};
		registerTypeInternal(Function, "Function", true);
		registerTypeInternal(Error, "Error", true);
		registerTypeInternal(Object, "Object", true);
		registerTypeInternal(String, "String", true);
		registerTypeInternal(Boolean, "Boolean", true);
		registerTypeInternal(Date, "Date", true);
		registerTypeInternal(Number, "Number", true);
		registerTypeInternal(RegExp, "RegExp", true);
		registerTypeInternal(Array, "Array", true);
		if (!Function.createCallback) {
			Function.createCallback=function Function$createCallback(method, context) {
				var e=Function._validateParams(arguments, [
					{ name: "method", type: Function },
					{ name: "context", mayBeNull: true }
				]);
				if (e)
					throw e;
				return function () {
					var l=arguments.length;
					if (l > 0) {
						var args=[];
						for (var i=0; i < l; i++) {
							args[i]=arguments[i];
						}
						args[l]=context;
						return method.apply(this, args);
					}
					return method.call(this, context);
				};
			};
		}
		if (!Function.createDelegate) {
			Function.createDelegate=function Function$createDelegate(instance, method) {
				var e=Function._validateParams(arguments, [
					{ name: "instance", mayBeNull: true },
					{ name: "method", type: Function }
				]);
				if (e)
					throw e;
				return function () {
					return method.apply(instance, arguments);
				};
			};
		}
		if (!Function._validateParams) {
			Function._validateParams=function (params, expectedParams, validateParameterCount) {
				var e, expectedLength=expectedParams.length;
				validateParameterCount=validateParameterCount || (typeof (validateParameterCount)==="undefined");
				e=Function._validateParameterCount(params, expectedParams, validateParameterCount);
				if (e) {
					e.popStackFrame();
					return e;
				}
				for (var i=0, l=params.length; i < l; i++) {
					var expectedParam=expectedParams[Math.min(i, expectedLength - 1)], paramName=expectedParam.name;
					if (expectedParam.parameterArray) {
						paramName+="["+(i - expectedLength+1)+"]";
					}
					else if (!validateParameterCount && (i >=expectedLength)) {
						break;
					}
					e=Function._validateParameter(params[i], expectedParam, paramName);
					if (e) {
						e.popStackFrame();
						return e;
					}
				}
				return null;
			};
		}
		if (!Function._validateParameterCount) {
			Function._validateParameterCount=function (params, expectedParams, validateParameterCount) {
				var i, error, expectedLen=expectedParams.length, actualLen=params.length;
				if (actualLen < expectedLen) {
					var minParams=expectedLen;
					for (i=0; i < expectedLen; i++) {
						var param=expectedParams[i];
						if (param.optional || param.parameterArray) {
							minParams--;
						}
					}
					if (actualLen < minParams) {
						error=true;
					}
				}
				else if (validateParameterCount && (actualLen > expectedLen)) {
					error=true;
					for (i=0; i < expectedLen; i++) {
						if (expectedParams[i].parameterArray) {
							error=false;
							break;
						}
					}
				}
				if (error) {
					var e=MsAjaxError.parameterCount();
					e.popStackFrame();
					return e;
				}
				return null;
			};
		}
		if (!Function._validateParameter) {
			Function._validateParameter=function (param, expectedParam, paramName) {
				var e, expectedType=expectedParam.type, expectedInteger=!!expectedParam.integer, expectedDomElement=!!expectedParam.domElement, mayBeNull=!!expectedParam.mayBeNull;
				e=Function._validateParameterType(param, expectedType, expectedInteger, expectedDomElement, mayBeNull, paramName);
				if (e) {
					e.popStackFrame();
					return e;
				}
				var expectedElementType=expectedParam.elementType, elementMayBeNull=!!expectedParam.elementMayBeNull;
				if (expectedType===Array && typeof (param) !=="undefined" && param !==null &&
					(expectedElementType || !elementMayBeNull)) {
					var expectedElementInteger=!!expectedParam.elementInteger, expectedElementDomElement=!!expectedParam.elementDomElement;
					for (var i=0; i < param.length; i++) {
						var elem=param[i];
						e=Function._validateParameterType(elem, expectedElementType, expectedElementInteger, expectedElementDomElement, elementMayBeNull, paramName+"["+i+"]");
						if (e) {
							e.popStackFrame();
							return e;
						}
					}
				}
				return null;
			};
		}
		if (!Function._validateParameterType) {
			Function._validateParameterType=function (param, expectedType, expectedInteger, expectedDomElement, mayBeNull, paramName) {
				var e, i;
				if (typeof (param)==="undefined") {
					if (mayBeNull) {
						return null;
					}
					else {
						e=OfficeExt.MsAjaxError.argumentUndefined(paramName);
						e.popStackFrame();
						return e;
					}
				}
				if (param===null) {
					if (mayBeNull) {
						return null;
					}
					else {
						e=OfficeExt.MsAjaxError.argumentNull(paramName);
						e.popStackFrame();
						return e;
					}
				}
				if (expectedType && !OfficeExt.MsAjaxTypeHelper.isInstanceOfType(expectedType, param)) {
					e=OfficeExt.MsAjaxError.argumentType(paramName, typeof (param), expectedType);
					e.popStackFrame();
					return e;
				}
				return null;
			};
		}
		if (!window.Type) {
			window.Type=Function;
		}
		if (!Type.registerNamespace) {
			Type.registerNamespace=function (ns) {
				var namespaceParts=ns.split('.');
				var currentNamespace=window;
				for (var i=0; i < namespaceParts.length; i++) {
					currentNamespace[namespaceParts[i]]=currentNamespace[namespaceParts[i]] || {};
					currentNamespace=currentNamespace[namespaceParts[i]];
				}
			};
		}
		if (!Type.prototype.registerClass) {
			Type.prototype.registerClass=function (cls) { cls={}; };
		}
		if (typeof (Sys)==="undefined") {
			Type.registerNamespace('Sys');
		}
		if (!Error.prototype.popStackFrame) {
			Error.prototype.popStackFrame=function () {
				if (arguments.length !==0)
					throw MsAjaxError.parameterCount();
				if (typeof (this.stack)==="undefined" || this.stack===null ||
					typeof (this.fileName)==="undefined" || this.fileName===null ||
					typeof (this.lineNumber)==="undefined" || this.lineNumber===null) {
					return;
				}
				var stackFrames=this.stack.split("\n");
				var currentFrame=stackFrames[0];
				var pattern=this.fileName+":"+this.lineNumber;
				while (typeof (currentFrame) !=="undefined" &&
					currentFrame !==null &&
					currentFrame.indexOf(pattern)===-1) {
					stackFrames.shift();
					currentFrame=stackFrames[0];
				}
				var nextFrame=stackFrames[1];
				if (typeof (nextFrame)==="undefined" || nextFrame===null) {
					return;
				}
				var nextFrameParts=nextFrame.match(/@(.*):(\d+)$/);
				if (typeof (nextFrameParts)==="undefined" || nextFrameParts===null) {
					return;
				}
				this.fileName=nextFrameParts[1];
				this.lineNumber=parseInt(nextFrameParts[2]);
				stackFrames.shift();
				this.stack=stackFrames.join("\n");
			};
		}
		OsfMsAjaxFactory.msAjaxError=MsAjaxError;
		OsfMsAjaxFactory.msAjaxString=MsAjaxString;
		OsfMsAjaxFactory.msAjaxDebug=MsAjaxDebug;
	}
})(OfficeExt || (OfficeExt={}));
OSF.OUtil.setNamespace("SafeArray", OSF.DDA);
OSF.DDA.SafeArray.Response={
	Status: 0,
	Payload: 1
};
OSF.DDA.SafeArray.UniqueArguments={
	Offset: "offset",
	Run: "run",
	BindingSpecificData: "bindingSpecificData",
	MergedCellGuid: "{66e7831f-81b2-42e2-823c-89e872d541b3}"
};
OSF.OUtil.setNamespace("Delegate", OSF.DDA.SafeArray);
OSF.DDA.SafeArray.Delegate._onException=function OSF_DDA_SafeArray_Delegate$OnException(ex, args) {
	var status;
	var statusNumber=ex.number;
	if (statusNumber) {
		switch (statusNumber) {
			case -2146828218:
				status=OSF.DDA.ErrorCodeManager.errorCodes.ooeNoCapability;
				break;
			case -2147467259:
				status=OSF.DDA.ErrorCodeManager.errorCodes.ooeDialogAlreadyOpened;
				break;
			case -2146828283:
				status=OSF.DDA.ErrorCodeManager.errorCodes.ooeInvalidParam;
				break;
			case -2147209089:
				status=OSF.DDA.ErrorCodeManager.errorCodes.ooeInvalidParam;
				break;
			case -2146827850:
			default:
				status=OSF.DDA.ErrorCodeManager.errorCodes.ooeInternalError;
				break;
		}
	}
	if (args.onComplete) {
		args.onComplete(status || OSF.DDA.ErrorCodeManager.errorCodes.ooeInternalError);
	}
};
OSF.DDA.SafeArray.Delegate._onExceptionSyncMethod=function OSF_DDA_SafeArray_Delegate$OnExceptionSyncMethod(ex, args) {
	var status;
	var number=ex.number;
	if (number) {
		switch (number) {
			case -2146828218:
				status=OSF.DDA.ErrorCodeManager.errorCodes.ooeNoCapability;
				break;
			case -2146827850:
			default:
				status=OSF.DDA.ErrorCodeManager.errorCodes.ooeInternalError;
				break;
		}
	}
	return status || OSF.DDA.ErrorCodeManager.errorCodes.ooeInternalError;
};
OSF.DDA.SafeArray.Delegate.SpecialProcessor=function OSF_DDA_SafeArray_Delegate_SpecialProcessor() {
	function _2DVBArrayToJaggedArray(vbArr) {
		var ret;
		try {
			var rows=vbArr.ubound(1);
			var cols=vbArr.ubound(2);
			vbArr=vbArr.toArray();
			if (rows==1 && cols==1) {
				ret=[vbArr];
			}
			else {
				ret=[];
				for (var row=0; row < rows; row++) {
					var rowArr=[];
					for (var col=0; col < cols; col++) {
						var datum=vbArr[row * cols+col];
						if (datum !=OSF.DDA.SafeArray.UniqueArguments.MergedCellGuid) {
							rowArr.push(datum);
						}
					}
					if (rowArr.length > 0) {
						ret.push(rowArr);
					}
				}
			}
		}
		catch (ex) {
		}
		return ret;
	}
	var complexTypes=[];
	var dynamicTypes={};
	dynamicTypes[Microsoft.Office.WebExtension.Parameters.Data]=(function () {
		var tableRows=0;
		var tableHeaders=1;
		return {
			toHost: function OSF_DDA_SafeArray_Delegate_SpecialProcessor_Data$toHost(data) {
				if (OSF.DDA.TableDataProperties && typeof data !="string" && data[OSF.DDA.TableDataProperties.TableRows] !==undefined) {
					var tableData=[];
					tableData[tableRows]=data[OSF.DDA.TableDataProperties.TableRows];
					tableData[tableHeaders]=data[OSF.DDA.TableDataProperties.TableHeaders];
					data=tableData;
				}
				return data;
			},
			fromHost: function OSF_DDA_SafeArray_Delegate_SpecialProcessor_Data$fromHost(hostArgs) {
				var ret;
				if (hostArgs.toArray) {
					var dimensions=hostArgs.dimensions();
					if (dimensions===2) {
						ret=_2DVBArrayToJaggedArray(hostArgs);
					}
					else {
						var array=hostArgs.toArray();
						if (array.length===2 && ((array[0] !=null && array[0].toArray) || (array[1] !=null && array[1].toArray))) {
							ret={};
							ret[OSF.DDA.TableDataProperties.TableRows]=_2DVBArrayToJaggedArray(array[tableRows]);
							ret[OSF.DDA.TableDataProperties.TableHeaders]=_2DVBArrayToJaggedArray(array[tableHeaders]);
						}
						else {
							ret=array;
						}
					}
				}
				else {
					ret=hostArgs;
				}
				return ret;
			}
		};
	})();
	OSF.DDA.SafeArray.Delegate.SpecialProcessor.uber.constructor.call(this, complexTypes, dynamicTypes);
	this.unpack=function OSF_DDA_SafeArray_Delegate_SpecialProcessor$unpack(param, arg) {
		var value;
		if (this.isComplexType(param) || OSF.DDA.ListType.isListType(param)) {
			var toArraySupported=(arg || typeof arg==="unknown") && arg.toArray;
			value=toArraySupported ? arg.toArray() : arg || {};
		}
		else if (this.isDynamicType(param)) {
			value=dynamicTypes[param].fromHost(arg);
		}
		else {
			value=arg;
		}
		return value;
	};
};
OSF.OUtil.extend(OSF.DDA.SafeArray.Delegate.SpecialProcessor, OSF.DDA.SpecialProcessor);
OSF.DDA.SafeArray.Delegate.ParameterMap=OSF.DDA.getDecoratedParameterMap(new OSF.DDA.SafeArray.Delegate.SpecialProcessor(), [
	{
		type: Microsoft.Office.WebExtension.Parameters.ValueFormat,
		toHost: [
			{ name: Microsoft.Office.WebExtension.ValueFormat.Unformatted, value: 0 },
			{ name: Microsoft.Office.WebExtension.ValueFormat.Formatted, value: 1 }
		]
	},
	{
		type: Microsoft.Office.WebExtension.Parameters.FilterType,
		toHost: [
			{ name: Microsoft.Office.WebExtension.FilterType.All, value: 0 }
		]
	}
]);
OSF.DDA.SafeArray.Delegate.ParameterMap.define({
	type: OSF.DDA.PropertyDescriptors.AsyncResultStatus,
	fromHost: [
		{ name: Microsoft.Office.WebExtension.AsyncResultStatus.Succeeded, value: 0 },
		{ name: Microsoft.Office.WebExtension.AsyncResultStatus.Failed, value: 1 }
	]
});
OSF.DDA.SafeArray.Delegate.executeAsync=function OSF_DDA_SafeArray_Delegate$ExecuteAsync(args) {
	function toArray(args) {
		var arrArgs=args;
		if (OSF.OUtil.isArray(args)) {
			var len=arrArgs.length;
			for (var i=0; i < len; i++) {
				arrArgs[i]=toArray(arrArgs[i]);
			}
		}
		else if (OSF.OUtil.isDate(args)) {
			arrArgs=args.getVarDate();
		}
		else if (typeof args==="object" && !OSF.OUtil.isArray(args)) {
			arrArgs=[];
			for (var index in args) {
				if (!OSF.OUtil.isFunction(args[index])) {
					arrArgs[index]=toArray(args[index]);
				}
			}
		}
		return arrArgs;
	}
	function fromSafeArray(value) {
		var ret=value;
		if (value !=null && value.toArray) {
			var arrayResult=value.toArray();
			ret=new Array(arrayResult.length);
			for (var i=0; i < arrayResult.length; i++) {
				ret[i]=fromSafeArray(arrayResult[i]);
			}
		}
		return ret;
	}
	try {
		if (args.onCalling) {
			args.onCalling();
		}
		OSF.ClientHostController.execute(args.dispId, toArray(args.hostCallArgs), function OSF_DDA_SafeArrayFacade$Execute_OnResponse(hostResponseArgs, resultCode) {
			var result=hostResponseArgs.toArray();
			var status=result[OSF.DDA.SafeArray.Response.Status];
			if (status==OSF.DDA.ErrorCodeManager.errorCodes.ooeChunkResult) {
				var payload=result[OSF.DDA.SafeArray.Response.Payload];
				payload=fromSafeArray(payload);
				if (payload !=null) {
					if (!args._chunkResultData) {
						args._chunkResultData=new Array();
					}
					args._chunkResultData[payload[0]]=payload[1];
				}
				return false;
			}
			if (args.onReceiving) {
				args.onReceiving();
			}
			if (args.onComplete) {
				var payload;
				if (status==OSF.DDA.ErrorCodeManager.errorCodes.ooeSuccess) {
					if (result.length > 2) {
						payload=[];
						for (var i=1; i < result.length; i++)
							payload[i - 1]=result[i];
					}
					else {
						payload=result[OSF.DDA.SafeArray.Response.Payload];
					}
					if (args._chunkResultData) {
						payload=fromSafeArray(payload);
						if (payload !=null) {
							var expectedChunkCount=payload[payload.length - 1];
							if (args._chunkResultData.length==expectedChunkCount) {
								payload[payload.length - 1]=args._chunkResultData;
							}
							else {
								status=OSF.DDA.ErrorCodeManager.errorCodes.ooeInternalError;
							}
						}
					}
				}
				else {
					payload=result[OSF.DDA.SafeArray.Response.Payload];
				}
				args.onComplete(status, payload);
			}
			return true;
		});
	}
	catch (ex) {
		OSF.DDA.SafeArray.Delegate._onException(ex, args);
	}
};
OSF.DDA.SafeArray.Delegate._getOnAfterRegisterEvent=function OSF_DDA_SafeArrayDelegate$GetOnAfterRegisterEvent(register, args) {
	var startTime=(new Date()).getTime();
	return function OSF_DDA_SafeArrayDelegate$OnAfterRegisterEvent(hostResponseArgs) {
		if (args.onReceiving) {
			args.onReceiving();
		}
		var status=hostResponseArgs.toArray ? hostResponseArgs.toArray()[OSF.DDA.SafeArray.Response.Status] : hostResponseArgs;
		if (args.onComplete) {
			args.onComplete(status);
		}
		if (OSF.AppTelemetry) {
			OSF.AppTelemetry.onRegisterDone(register, args.dispId, Math.abs((new Date()).getTime() - startTime), status);
		}
	};
};
OSF.DDA.SafeArray.Delegate.registerEventAsync=function OSF_DDA_SafeArray_Delegate$RegisterEventAsync(args) {
	if (args.onCalling) {
		args.onCalling();
	}
	var callback=OSF.DDA.SafeArray.Delegate._getOnAfterRegisterEvent(true, args);
	try {
		OSF.ClientHostController.registerEvent(args.dispId, args.targetId, function OSF_DDA_SafeArrayDelegate$RegisterEventAsync_OnEvent(eventDispId, payload) {
			if (args.onEvent) {
				args.onEvent(payload);
			}
			if (OSF.AppTelemetry) {
				OSF.AppTelemetry.onEventDone(args.dispId);
			}
		}, callback);
	}
	catch (ex) {
		OSF.DDA.SafeArray.Delegate._onException(ex, args);
	}
};
OSF.DDA.SafeArray.Delegate.unregisterEventAsync=function OSF_DDA_SafeArray_Delegate$UnregisterEventAsync(args) {
	if (args.onCalling) {
		args.onCalling();
	}
	var callback=OSF.DDA.SafeArray.Delegate._getOnAfterRegisterEvent(false, args);
	try {
		OSF.ClientHostController.unregisterEvent(args.dispId, args.targetId, callback);
	}
	catch (ex) {
		OSF.DDA.SafeArray.Delegate._onException(ex, args);
	}
};
OSF.ClientMode={
	ReadWrite: 0,
	ReadOnly: 1
};
OSF.DDA.RichInitializationReason={
	1: Microsoft.Office.WebExtension.InitializationReason.Inserted,
	2: Microsoft.Office.WebExtension.InitializationReason.DocumentOpened
};
OSF.InitializationHelper=function OSF_InitializationHelper(hostInfo, webAppState, context, settings, hostFacade) {
	this._hostInfo=hostInfo;
	this._webAppState=webAppState;
	this._context=context;
	this._settings=settings;
	this._hostFacade=hostFacade;
	this._initializeSettings=this.initializeSettings;
};
OSF.InitializationHelper.prototype.deserializeSettings=function OSF_InitializationHelper$deserializeSettings(serializedSettings, refreshSupported) {
	var settings;
	var osfSessionStorage=OSF.OUtil.getSessionStorage();
	if (osfSessionStorage) {
		var storageSettings=osfSessionStorage.getItem(OSF._OfficeAppFactory.getCachedSessionSettingsKey());
		if (storageSettings) {
			serializedSettings=JSON.parse(storageSettings);
		}
		else {
			storageSettings=JSON.stringify(serializedSettings);
			osfSessionStorage.setItem(OSF._OfficeAppFactory.getCachedSessionSettingsKey(), storageSettings);
		}
	}
	var deserializedSettings=OSF.DDA.SettingsManager.deserializeSettings(serializedSettings);
	if (refreshSupported) {
		settings=new OSF.DDA.RefreshableSettings(deserializedSettings);
	}
	else {
		settings=new OSF.DDA.Settings(deserializedSettings);
	}
	return settings;
};
OSF.InitializationHelper.prototype.saveAndSetDialogInfo=function OSF_InitializationHelper$saveAndSetDialogInfo(hostInfoValue) {
};
OSF.InitializationHelper.prototype.setAgaveHostCommunication=function OSF_InitializationHelper$setAgaveHostCommunication() {
};
OSF.InitializationHelper.prototype.prepareRightBeforeWebExtensionInitialize=function OSF_InitializationHelper$prepareRightBeforeWebExtensionInitialize(appContext) {
	this.prepareApiSurface(appContext);
	Microsoft.Office.WebExtension.initialize(this.getInitializationReason(appContext));
};
OSF.InitializationHelper.prototype.prepareApiSurface=function OSF_InitializationHelper$prepareApiSurfaceAndInitialize(appContext) {
	var license=new OSF.DDA.License(appContext.get_eToken());
	var getOfficeThemeHandler=(OSF.DDA.OfficeTheme && OSF.DDA.OfficeTheme.getOfficeTheme) ? OSF.DDA.OfficeTheme.getOfficeTheme : null;
	if (appContext.get_isDialog()) {
		if (OSF.DDA.UI.ChildUI) {
			appContext.ui=new OSF.DDA.UI.ChildUI();
		}
	}
	else {
		if (OSF.DDA.UI.ParentUI) {
			appContext.ui=new OSF.DDA.UI.ParentUI();
			if (OfficeExt.Container) {
				OSF.DDA.DispIdHost.addAsyncMethods(appContext.ui, [OSF.DDA.AsyncMethodNames.CloseContainerAsync]);
			}
		}
	}
	if (OSF.DDA.Auth) {
		appContext.auth=new OSF.DDA.Auth();
		OSF.DDA.DispIdHost.addAsyncMethods(appContext.auth, [OSF.DDA.AsyncMethodNames.GetAccessTokenAsync]);
	}
	OSF._OfficeAppFactory.setContext(new OSF.DDA.Context(appContext, appContext.doc, license, null, getOfficeThemeHandler));
	var getDelegateMethods, parameterMap;
	getDelegateMethods=OSF.DDA.DispIdHost.getClientDelegateMethods;
	parameterMap=OSF.DDA.SafeArray.Delegate.ParameterMap;
	OSF._OfficeAppFactory.setHostFacade(new OSF.DDA.DispIdHost.Facade(getDelegateMethods, parameterMap));
};
OSF.InitializationHelper.prototype.getInitializationReason=function (appContext) { return OSF.DDA.RichInitializationReason[appContext.get_reason()]; };
OSF.DDA.DispIdHost.getClientDelegateMethods=function (actionId) {
	var delegateMethods={};
	delegateMethods[OSF.DDA.DispIdHost.Delegates.ExecuteAsync]=OSF.DDA.SafeArray.Delegate.executeAsync;
	delegateMethods[OSF.DDA.DispIdHost.Delegates.RegisterEventAsync]=OSF.DDA.SafeArray.Delegate.registerEventAsync;
	delegateMethods[OSF.DDA.DispIdHost.Delegates.UnregisterEventAsync]=OSF.DDA.SafeArray.Delegate.unregisterEventAsync;
	delegateMethods[OSF.DDA.DispIdHost.Delegates.OpenDialog]=OSF.DDA.SafeArray.Delegate.openDialog;
	delegateMethods[OSF.DDA.DispIdHost.Delegates.CloseDialog]=OSF.DDA.SafeArray.Delegate.closeDialog;
	delegateMethods[OSF.DDA.DispIdHost.Delegates.MessageParent]=OSF.DDA.SafeArray.Delegate.messageParent;
	delegateMethods[OSF.DDA.DispIdHost.Delegates.SendMessage]=OSF.DDA.SafeArray.Delegate.sendMessage;
	if (OSF.DDA.AsyncMethodNames.RefreshAsync && actionId==OSF.DDA.AsyncMethodNames.RefreshAsync.id) {
		var readSerializedSettings=function (hostCallArgs, onCalling, onReceiving) {
			return OSF.DDA.ClientSettingsManager.read(onCalling, onReceiving);
		};
		delegateMethods[OSF.DDA.DispIdHost.Delegates.ExecuteAsync]=OSF.DDA.ClientSettingsManager.getSettingsExecuteMethod(readSerializedSettings);
	}
	if (OSF.DDA.AsyncMethodNames.SaveAsync && actionId==OSF.DDA.AsyncMethodNames.SaveAsync.id) {
		var writeSerializedSettings=function (hostCallArgs, onCalling, onReceiving) {
			return OSF.DDA.ClientSettingsManager.write(hostCallArgs[OSF.DDA.SettingsManager.SerializedSettings], hostCallArgs[Microsoft.Office.WebExtension.Parameters.OverwriteIfStale], onCalling, onReceiving);
		};
		delegateMethods[OSF.DDA.DispIdHost.Delegates.ExecuteAsync]=OSF.DDA.ClientSettingsManager.getSettingsExecuteMethod(writeSerializedSettings);
	}
	return delegateMethods;
};
var OfficeExt;
(function (OfficeExt) {
	var RichClientHostController=(function () {
		function RichClientHostController() {
		}
		RichClientHostController.prototype.execute=function (id, params, callback) {
			window.external.Execute(id, params, callback);
		};
		RichClientHostController.prototype.registerEvent=function (id, targetId, handler, callback) {
			window.external.RegisterEvent(id, targetId, handler, callback);
		};
		RichClientHostController.prototype.unregisterEvent=function (id, targetId, callback) {
			window.external.UnregisterEvent(id, targetId, callback);
		};
		return RichClientHostController;
	})();
	OfficeExt.RichClientHostController=RichClientHostController;
})(OfficeExt || (OfficeExt={}));
var OfficeExt;
(function (OfficeExt) {
	var Win32RichClientHostController=(function (_super) {
		__extends(Win32RichClientHostController, _super);
		function Win32RichClientHostController() {
			_super.apply(this, arguments);
		}
		Win32RichClientHostController.prototype.messageParent=function (params) {
			var message=params[Microsoft.Office.WebExtension.Parameters.MessageToParent];
			window.external.MessageParent(message);
		};
		Win32RichClientHostController.prototype.openDialog=function (id, targetId, handler, callback) {
			this.registerEvent(id, targetId, handler, callback);
		};
		Win32RichClientHostController.prototype.closeDialog=function (id, targetId, callback) {
			this.unregisterEvent(id, targetId, callback);
		};
		Win32RichClientHostController.prototype.sendMessage=function (params) {
		};
		return Win32RichClientHostController;
	})(OfficeExt.RichClientHostController);
	OfficeExt.Win32RichClientHostController=Win32RichClientHostController;
})(OfficeExt || (OfficeExt={}));
OSF.ClientHostController=new OfficeExt.Win32RichClientHostController();
var OfficeExt;
(function (OfficeExt) {
	var OfficeTheme;
	(function (OfficeTheme) {
		var OfficeThemeManager=(function () {
			function OfficeThemeManager() {
				this._osfOfficeTheme=null;
				this._osfOfficeThemeTimeStamp=null;
			}
			OfficeThemeManager.prototype.getOfficeTheme=function () {
				if (OSF.DDA._OsfControlContext) {
					if (this._osfOfficeTheme && this._osfOfficeThemeTimeStamp && ((new Date()).getTime() - this._osfOfficeThemeTimeStamp < OfficeThemeManager._osfOfficeThemeCacheValidPeriod)) {
						if (OSF.AppTelemetry) {
							OSF.AppTelemetry.onPropertyDone("GetOfficeThemeInfo", 0);
						}
					}
					else {
						var startTime=(new Date()).getTime();
						var osfOfficeTheme=OSF.DDA._OsfControlContext.GetOfficeThemeInfo();
						var endTime=(new Date()).getTime();
						if (OSF.AppTelemetry) {
							OSF.AppTelemetry.onPropertyDone("GetOfficeThemeInfo", Math.abs(endTime - startTime));
						}
						this._osfOfficeTheme=JSON.parse(osfOfficeTheme);
						for (var color in this._osfOfficeTheme) {
							this._osfOfficeTheme[color]=OSF.OUtil.convertIntToCssHexColor(this._osfOfficeTheme[color]);
						}
						this._osfOfficeThemeTimeStamp=endTime;
					}
					return this._osfOfficeTheme;
				}
			};
			OfficeThemeManager.instance=function () {
				if (OfficeThemeManager._instance==null) {
					OfficeThemeManager._instance=new OfficeThemeManager();
				}
				return OfficeThemeManager._instance;
			};
			OfficeThemeManager._osfOfficeThemeCacheValidPeriod=5000;
			OfficeThemeManager._instance=null;
			return OfficeThemeManager;
		})();
		OfficeTheme.OfficeThemeManager=OfficeThemeManager;
		OSF.OUtil.setNamespace("OfficeTheme", OSF.DDA);
		OSF.DDA.OfficeTheme.getOfficeTheme=OfficeExt.OfficeTheme.OfficeThemeManager.instance().getOfficeTheme;
	})(OfficeTheme=OfficeExt.OfficeTheme || (OfficeExt.OfficeTheme={}));
})(OfficeExt || (OfficeExt={}));
OSF.DDA.ClientSettingsManager={
	getSettingsExecuteMethod: function OSF_DDA_ClientSettingsManager$getSettingsExecuteMethod(hostDelegateMethod) {
		return function (args) {
			var status, response;
			try {
				response=hostDelegateMethod(args.hostCallArgs, args.onCalling, args.onReceiving);
				status=OSF.DDA.ErrorCodeManager.errorCodes.ooeSuccess;
			}
			catch (ex) {
				status=OSF.DDA.ErrorCodeManager.errorCodes.ooeInternalError;
				response={ name: Strings.OfficeOM.L_InternalError, message: ex };
			}
			if (args.onComplete) {
				args.onComplete(status, response);
			}
		};
	},
	read: function OSF_DDA_ClientSettingsManager$read(onCalling, onReceiving) {
		var keys=[];
		var values=[];
		if (onCalling) {
			onCalling();
		}
		OSF.DDA._OsfControlContext.GetSettings().Read(keys, values);
		if (onReceiving) {
			onReceiving();
		}
		var serializedSettings={};
		for (var index=0; index < keys.length; index++) {
			serializedSettings[keys[index]]=values[index];
		}
		return serializedSettings;
	},
	write: function OSF_DDA_ClientSettingsManager$write(serializedSettings, overwriteIfStale, onCalling, onReceiving) {
		var keys=[];
		var values=[];
		for (var key in serializedSettings) {
			keys.push(key);
			values.push(serializedSettings[key]);
		}
		if (onCalling) {
			onCalling();
		}
		OSF.DDA._OsfControlContext.GetSettings().Write(keys, values);
		if (onReceiving) {
			onReceiving();
		}
	}
};
OSF.InitializationHelper.prototype.initializeSettings=function OSF_InitializationHelper$initializeSettings(refreshSupported) {
	var serializedSettings=OSF.DDA.ClientSettingsManager.read();
	var settings=this.deserializeSettings(serializedSettings, refreshSupported);
	return settings;
};
OSF.InitializationHelper.prototype.getAppContext=function OSF_InitializationHelper$getAppContext(wnd, gotAppContext) {
	var returnedContext;
	var context;
	var warningText="Warning: Office.js is loaded outside of Office client";
	try {
		if (window.external && typeof window.external.GetContext !=='undefined') {
			context=OSF.DDA._OsfControlContext=window.external.GetContext();
		}
		else {
			OsfMsAjaxFactory.msAjaxDebug.trace(warningText);
			return;
		}
	}
	catch (e) {
		OsfMsAjaxFactory.msAjaxDebug.trace(warningText);
		return;
	}
	var appType=context.GetAppType();
	var id=context.GetSolutionRef();
	var version=context.GetAppVersionMajor();
	var minorVersion=context.GetAppVersionMinor();
	var UILocale=context.GetAppUILocale();
	var dataLocale=context.GetAppDataLocale();
	var docUrl=context.GetDocUrl();
	var clientMode=context.GetAppCapabilities();
	var reason=context.GetActivationMode();
	var osfControlType=context.GetControlIntegrationLevel();
	var settings=[];
	var eToken;
	try {
		eToken=context.GetSolutionToken();
	}
	catch (ex) {
	}
	var correlationId;
	if (typeof context.GetCorrelationId !=="undefined") {
		correlationId=context.GetCorrelationId();
	}
	var appInstanceId;
	if (typeof context.GetInstanceId !=="undefined") {
		appInstanceId=context.GetInstanceId();
	}
	var touchEnabled;
	if (typeof context.GetTouchEnabled !=="undefined") {
		touchEnabled=context.GetTouchEnabled();
	}
	var commerceAllowed;
	if (typeof context.GetCommerceAllowed !=="undefined") {
		commerceAllowed=context.GetCommerceAllowed();
	}
	var requirementMatrix;
	if (typeof context.GetSupportedMatrix !=="undefined") {
		requirementMatrix=context.GetSupportedMatrix();
	}
	var hostCustomMessage;
	if (typeof context.GetHostCustomMessage !=="undefined") {
		hostCustomMessage=context.GetHostCustomMessage();
	}
	var hostFullVersion;
	if (typeof context.GetHostFullVersion !=="undefined") {
		hostFullVersion=context.GetHostFullVersion();
	}
	var dialogRequirementMatrix;
	if (typeof context.GetDialogRequirementMatrix !="undefined") {
		dialogRequirementMatrix=context.GetDialogRequirementMatrix();
	}
	eToken=eToken ? eToken.toString() : "";
	returnedContext=new OSF.OfficeAppContext(id, appType, version, UILocale, dataLocale, docUrl, clientMode, settings, reason, osfControlType, eToken, correlationId, appInstanceId, touchEnabled, commerceAllowed, minorVersion, requirementMatrix, hostCustomMessage, hostFullVersion, undefined, undefined, undefined, dialogRequirementMatrix);
	if (OSF.AppTelemetry) {
		OSF.AppTelemetry.initialize(returnedContext);
	}
	gotAppContext(returnedContext);
};
var OSFLog;
(function (OSFLog) {
	var BaseUsageData=(function () {
		function BaseUsageData(table) {
			this._table=table;
			this._fields={};
		}
		Object.defineProperty(BaseUsageData.prototype, "Fields", {
			get: function () {
				return this._fields;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(BaseUsageData.prototype, "Table", {
			get: function () {
				return this._table;
			},
			enumerable: true,
			configurable: true
		});
		BaseUsageData.prototype.SerializeFields=function () {
		};
		BaseUsageData.prototype.SetSerializedField=function (key, value) {
			if (typeof (value) !=="undefined" && value !==null) {
				this._serializedFields[key]=value.toString();
			}
		};
		BaseUsageData.prototype.SerializeRow=function () {
			this._serializedFields={};
			this.SetSerializedField("Table", this._table);
			this.SerializeFields();
			return JSON.stringify(this._serializedFields);
		};
		return BaseUsageData;
	})();
	OSFLog.BaseUsageData=BaseUsageData;
	var AppActivatedUsageData=(function (_super) {
		__extends(AppActivatedUsageData, _super);
		function AppActivatedUsageData() {
			_super.call(this, "AppActivated");
		}
		Object.defineProperty(AppActivatedUsageData.prototype, "CorrelationId", {
			get: function () { return this.Fields["CorrelationId"]; },
			set: function (value) { this.Fields["CorrelationId"]=value; },
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(AppActivatedUsageData.prototype, "SessionId", {
			get: function () { return this.Fields["SessionId"]; },
			set: function (value) { this.Fields["SessionId"]=value; },
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(AppActivatedUsageData.prototype, "AppId", {
			get: function () { return this.Fields["AppId"]; },
			set: function (value) { this.Fields["AppId"]=value; },
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(AppActivatedUsageData.prototype, "AppInstanceId", {
			get: function () { return this.Fields["AppInstanceId"]; },
			set: function (value) { this.Fields["AppInstanceId"]=value; },
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(AppActivatedUsageData.prototype, "AppURL", {
			get: function () { return this.Fields["AppURL"]; },
			set: function (value) { this.Fields["AppURL"]=value; },
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(AppActivatedUsageData.prototype, "AssetId", {
			get: function () { return this.Fields["AssetId"]; },
			set: function (value) { this.Fields["AssetId"]=value; },
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(AppActivatedUsageData.prototype, "Browser", {
			get: function () { return this.Fields["Browser"]; },
			set: function (value) { this.Fields["Browser"]=value; },
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(AppActivatedUsageData.prototype, "UserId", {
			get: function () { return this.Fields["UserId"]; },
			set: function (value) { this.Fields["UserId"]=value; },
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(AppActivatedUsageData.prototype, "Host", {
			get: function () { return this.Fields["Host"]; },
			set: function (value) { this.Fields["Host"]=value; },
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(AppActivatedUsageData.prototype, "HostVersion", {
			get: function () { return this.Fields["HostVersion"]; },
			set: function (value) { this.Fields["HostVersion"]=value; },
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(AppActivatedUsageData.prototype, "ClientId", {
			get: function () { return this.Fields["ClientId"]; },
			set: function (value) { this.Fields["ClientId"]=value; },
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(AppActivatedUsageData.prototype, "AppSizeWidth", {
			get: function () { return this.Fields["AppSizeWidth"]; },
			set: function (value) { this.Fields["AppSizeWidth"]=value; },
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(AppActivatedUsageData.prototype, "AppSizeHeight", {
			get: function () { return this.Fields["AppSizeHeight"]; },
			set: function (value) { this.Fields["AppSizeHeight"]=value; },
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(AppActivatedUsageData.prototype, "Message", {
			get: function () { return this.Fields["Message"]; },
			set: function (value) { this.Fields["Message"]=value; },
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(AppActivatedUsageData.prototype, "DocUrl", {
			get: function () { return this.Fields["DocUrl"]; },
			set: function (value) { this.Fields["DocUrl"]=value; },
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(AppActivatedUsageData.prototype, "OfficeJSVersion", {
			get: function () { return this.Fields["OfficeJSVersion"]; },
			set: function (value) { this.Fields["OfficeJSVersion"]=value; },
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(AppActivatedUsageData.prototype, "HostJSVersion", {
			get: function () { return this.Fields["HostJSVersion"]; },
			set: function (value) { this.Fields["HostJSVersion"]=value; },
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(AppActivatedUsageData.prototype, "WacHostEnvironment", {
			get: function () { return this.Fields["WacHostEnvironment"]; },
			set: function (value) { this.Fields["WacHostEnvironment"]=value; },
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(AppActivatedUsageData.prototype, "IsFromWacAutomation", {
			get: function () { return this.Fields["IsFromWacAutomation"]; },
			set: function (value) { this.Fields["IsFromWacAutomation"]=value; },
			enumerable: true,
			configurable: true
		});
		AppActivatedUsageData.prototype.SerializeFields=function () {
			this.SetSerializedField("CorrelationId", this.CorrelationId);
			this.SetSerializedField("SessionId", this.SessionId);
			this.SetSerializedField("AppId", this.AppId);
			this.SetSerializedField("AppInstanceId", this.AppInstanceId);
			this.SetSerializedField("AppURL", this.AppURL);
			this.SetSerializedField("AssetId", this.AssetId);
			this.SetSerializedField("Browser", this.Browser);
			this.SetSerializedField("UserId", this.UserId);
			this.SetSerializedField("Host", this.Host);
			this.SetSerializedField("HostVersion", this.HostVersion);
			this.SetSerializedField("ClientId", this.ClientId);
			this.SetSerializedField("AppSizeWidth", this.AppSizeWidth);
			this.SetSerializedField("AppSizeHeight", this.AppSizeHeight);
			this.SetSerializedField("Message", this.Message);
			this.SetSerializedField("DocUrl", this.DocUrl);
			this.SetSerializedField("OfficeJSVersion", this.OfficeJSVersion);
			this.SetSerializedField("HostJSVersion", this.HostJSVersion);
			this.SetSerializedField("WacHostEnvironment", this.WacHostEnvironment);
			this.SetSerializedField("IsFromWacAutomation", this.IsFromWacAutomation);
		};
		return AppActivatedUsageData;
	})(BaseUsageData);
	OSFLog.AppActivatedUsageData=AppActivatedUsageData;
	var ScriptLoadUsageData=(function (_super) {
		__extends(ScriptLoadUsageData, _super);
		function ScriptLoadUsageData() {
			_super.call(this, "ScriptLoad");
		}
		Object.defineProperty(ScriptLoadUsageData.prototype, "CorrelationId", {
			get: function () { return this.Fields["CorrelationId"]; },
			set: function (value) { this.Fields["CorrelationId"]=value; },
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ScriptLoadUsageData.prototype, "SessionId", {
			get: function () { return this.Fields["SessionId"]; },
			set: function (value) { this.Fields["SessionId"]=value; },
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ScriptLoadUsageData.prototype, "ScriptId", {
			get: function () { return this.Fields["ScriptId"]; },
			set: function (value) { this.Fields["ScriptId"]=value; },
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ScriptLoadUsageData.prototype, "StartTime", {
			get: function () { return this.Fields["StartTime"]; },
			set: function (value) { this.Fields["StartTime"]=value; },
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ScriptLoadUsageData.prototype, "ResponseTime", {
			get: function () { return this.Fields["ResponseTime"]; },
			set: function (value) { this.Fields["ResponseTime"]=value; },
			enumerable: true,
			configurable: true
		});
		ScriptLoadUsageData.prototype.SerializeFields=function () {
			this.SetSerializedField("CorrelationId", this.CorrelationId);
			this.SetSerializedField("SessionId", this.SessionId);
			this.SetSerializedField("ScriptId", this.ScriptId);
			this.SetSerializedField("StartTime", this.StartTime);
			this.SetSerializedField("ResponseTime", this.ResponseTime);
		};
		return ScriptLoadUsageData;
	})(BaseUsageData);
	OSFLog.ScriptLoadUsageData=ScriptLoadUsageData;
	var AppClosedUsageData=(function (_super) {
		__extends(AppClosedUsageData, _super);
		function AppClosedUsageData() {
			_super.call(this, "AppClosed");
		}
		Object.defineProperty(AppClosedUsageData.prototype, "CorrelationId", {
			get: function () { return this.Fields["CorrelationId"]; },
			set: function (value) { this.Fields["CorrelationId"]=value; },
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(AppClosedUsageData.prototype, "SessionId", {
			get: function () { return this.Fields["SessionId"]; },
			set: function (value) { this.Fields["SessionId"]=value; },
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(AppClosedUsageData.prototype, "FocusTime", {
			get: function () { return this.Fields["FocusTime"]; },
			set: function (value) { this.Fields["FocusTime"]=value; },
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(AppClosedUsageData.prototype, "AppSizeFinalWidth", {
			get: function () { return this.Fields["AppSizeFinalWidth"]; },
			set: function (value) { this.Fields["AppSizeFinalWidth"]=value; },
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(AppClosedUsageData.prototype, "AppSizeFinalHeight", {
			get: function () { return this.Fields["AppSizeFinalHeight"]; },
			set: function (value) { this.Fields["AppSizeFinalHeight"]=value; },
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(AppClosedUsageData.prototype, "OpenTime", {
			get: function () { return this.Fields["OpenTime"]; },
			set: function (value) { this.Fields["OpenTime"]=value; },
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(AppClosedUsageData.prototype, "CloseMethod", {
			get: function () { return this.Fields["CloseMethod"]; },
			set: function (value) { this.Fields["CloseMethod"]=value; },
			enumerable: true,
			configurable: true
		});
		AppClosedUsageData.prototype.SerializeFields=function () {
			this.SetSerializedField("CorrelationId", this.CorrelationId);
			this.SetSerializedField("SessionId", this.SessionId);
			this.SetSerializedField("FocusTime", this.FocusTime);
			this.SetSerializedField("AppSizeFinalWidth", this.AppSizeFinalWidth);
			this.SetSerializedField("AppSizeFinalHeight", this.AppSizeFinalHeight);
			this.SetSerializedField("OpenTime", this.OpenTime);
			this.SetSerializedField("CloseMethod", this.CloseMethod);
		};
		return AppClosedUsageData;
	})(BaseUsageData);
	OSFLog.AppClosedUsageData=AppClosedUsageData;
	var APIUsageUsageData=(function (_super) {
		__extends(APIUsageUsageData, _super);
		function APIUsageUsageData() {
			_super.call(this, "APIUsage");
		}
		Object.defineProperty(APIUsageUsageData.prototype, "CorrelationId", {
			get: function () { return this.Fields["CorrelationId"]; },
			set: function (value) { this.Fields["CorrelationId"]=value; },
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(APIUsageUsageData.prototype, "SessionId", {
			get: function () { return this.Fields["SessionId"]; },
			set: function (value) { this.Fields["SessionId"]=value; },
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(APIUsageUsageData.prototype, "APIType", {
			get: function () { return this.Fields["APIType"]; },
			set: function (value) { this.Fields["APIType"]=value; },
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(APIUsageUsageData.prototype, "APIID", {
			get: function () { return this.Fields["APIID"]; },
			set: function (value) { this.Fields["APIID"]=value; },
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(APIUsageUsageData.prototype, "Parameters", {
			get: function () { return this.Fields["Parameters"]; },
			set: function (value) { this.Fields["Parameters"]=value; },
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(APIUsageUsageData.prototype, "ResponseTime", {
			get: function () { return this.Fields["ResponseTime"]; },
			set: function (value) { this.Fields["ResponseTime"]=value; },
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(APIUsageUsageData.prototype, "ErrorType", {
			get: function () { return this.Fields["ErrorType"]; },
			set: function (value) { this.Fields["ErrorType"]=value; },
			enumerable: true,
			configurable: true
		});
		APIUsageUsageData.prototype.SerializeFields=function () {
			this.SetSerializedField("CorrelationId", this.CorrelationId);
			this.SetSerializedField("SessionId", this.SessionId);
			this.SetSerializedField("APIType", this.APIType);
			this.SetSerializedField("APIID", this.APIID);
			this.SetSerializedField("Parameters", this.Parameters);
			this.SetSerializedField("ResponseTime", this.ResponseTime);
			this.SetSerializedField("ErrorType", this.ErrorType);
		};
		return APIUsageUsageData;
	})(BaseUsageData);
	OSFLog.APIUsageUsageData=APIUsageUsageData;
	var AppInitializationUsageData=(function (_super) {
		__extends(AppInitializationUsageData, _super);
		function AppInitializationUsageData() {
			_super.call(this, "AppInitialization");
		}
		Object.defineProperty(AppInitializationUsageData.prototype, "CorrelationId", {
			get: function () { return this.Fields["CorrelationId"]; },
			set: function (value) { this.Fields["CorrelationId"]=value; },
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(AppInitializationUsageData.prototype, "SessionId", {
			get: function () { return this.Fields["SessionId"]; },
			set: function (value) { this.Fields["SessionId"]=value; },
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(AppInitializationUsageData.prototype, "SuccessCode", {
			get: function () { return this.Fields["SuccessCode"]; },
			set: function (value) { this.Fields["SuccessCode"]=value; },
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(AppInitializationUsageData.prototype, "Message", {
			get: function () { return this.Fields["Message"]; },
			set: function (value) { this.Fields["Message"]=value; },
			enumerable: true,
			configurable: true
		});
		AppInitializationUsageData.prototype.SerializeFields=function () {
			this.SetSerializedField("CorrelationId", this.CorrelationId);
			this.SetSerializedField("SessionId", this.SessionId);
			this.SetSerializedField("SuccessCode", this.SuccessCode);
			this.SetSerializedField("Message", this.Message);
		};
		return AppInitializationUsageData;
	})(BaseUsageData);
	OSFLog.AppInitializationUsageData=AppInitializationUsageData;
})(OSFLog || (OSFLog={}));
var Logger;
(function (Logger) {
	"use strict";
	(function (TraceLevel) {
		TraceLevel[TraceLevel["info"]=0]="info";
		TraceLevel[TraceLevel["warning"]=1]="warning";
		TraceLevel[TraceLevel["error"]=2]="error";
	})(Logger.TraceLevel || (Logger.TraceLevel={}));
	var TraceLevel=Logger.TraceLevel;
	(function (SendFlag) {
		SendFlag[SendFlag["none"]=0]="none";
		SendFlag[SendFlag["flush"]=1]="flush";
	})(Logger.SendFlag || (Logger.SendFlag={}));
	var SendFlag=Logger.SendFlag;
	function allowUploadingData() {
		if (OSF.Logger && OSF.Logger.ulsEndpoint) {
			OSF.Logger.ulsEndpoint.loadProxyFrame();
		}
	}
	Logger.allowUploadingData=allowUploadingData;
	function sendLog(traceLevel, message, flag) {
		if (OSF.Logger && OSF.Logger.ulsEndpoint) {
			var jsonObj={ traceLevel: traceLevel, message: message, flag: flag, internalLog: true };
			var logs=JSON.stringify(jsonObj);
			OSF.Logger.ulsEndpoint.writeLog(logs);
		}
	}
	Logger.sendLog=sendLog;
	function creatULSEndpoint() {
		try {
			return new ULSEndpointProxy();
		}
		catch (e) {
			return null;
		}
	}
	var ULSEndpointProxy=(function () {
		function ULSEndpointProxy() {
			var _this=this;
			this.proxyFrame=null;
			this.telemetryEndPoint="https://telemetryservice.firstpartyapps.oaspapps.com/telemetryservice/telemetryproxy.html";
			this.buffer=[];
			this.proxyFrameReady=false;
			OSF.OUtil.addEventListener(window, "message", function (e) { return _this.tellProxyFrameReady(e); });
			setTimeout(function () {
				_this.loadProxyFrame();
			}, 3000);
		}
		ULSEndpointProxy.prototype.writeLog=function (log) {
			if (this.proxyFrameReady===true) {
				this.proxyFrame.contentWindow.postMessage(log, ULSEndpointProxy.telemetryOrigin);
			}
			else {
				if (this.buffer.length < 128) {
					this.buffer.push(log);
				}
			}
		};
		ULSEndpointProxy.prototype.loadProxyFrame=function () {
			if (this.proxyFrame==null) {
				this.proxyFrame=document.createElement("iframe");
				this.proxyFrame.setAttribute("style", "display:none");
				this.proxyFrame.setAttribute("src", this.telemetryEndPoint);
				document.head.appendChild(this.proxyFrame);
			}
		};
		ULSEndpointProxy.prototype.tellProxyFrameReady=function (e) {
			var _this=this;
			if (e.data==="ProxyFrameReadyToLog") {
				this.proxyFrameReady=true;
				for (var i=0; i < this.buffer.length; i++) {
					this.writeLog(this.buffer[i]);
				}
				this.buffer.length=0;
				OSF.OUtil.removeEventListener(window, "message", function (e) { return _this.tellProxyFrameReady(e); });
			}
			else if (e.data==="ProxyFrameReadyToInit") {
				var initJson={ appName: "Office APPs", sessionId: OSF.OUtil.Guid.generateNewGuid() };
				var initStr=JSON.stringify(initJson);
				this.proxyFrame.contentWindow.postMessage(initStr, ULSEndpointProxy.telemetryOrigin);
			}
		};
		ULSEndpointProxy.telemetryOrigin="https://telemetryservice.firstpartyapps.oaspapps.com";
		return ULSEndpointProxy;
	})();
	if (!OSF.Logger) {
		OSF.Logger=Logger;
	}
	Logger.ulsEndpoint=creatULSEndpoint();
})(Logger || (Logger={}));
var OSFAriaLogger;
(function (OSFAriaLogger) {
	var AriaLogger=(function () {
		function AriaLogger() {
		}
		AriaLogger.prototype.getAriaCDNLocation=function () {
			return (OSF._OfficeAppFactory.getLoadScriptHelper().getOfficeJsBasePath()+"/ariatelemetry/aria-web-telemetry.js");
		};
		AriaLogger.getInstance=function () {
			if (AriaLogger.AriaLoggerObj===undefined) {
				AriaLogger.AriaLoggerObj=new AriaLogger();
			}
			return AriaLogger.AriaLoggerObj;
		};
		AriaLogger.prototype.isIUsageData=function (arg) {
			return arg["Fields"] !==undefined;
		};
		AriaLogger.prototype.loadAriaScriptAndLog=function (tableName, telemetryData) {
			var startAfterMs=1000;
			OSF.OUtil.loadScript(this.getAriaCDNLocation(), function () {
				try {
					if (!this.ALogger) {
						var OfficeExtensibilityTenantID="db334b301e7b474db5e0f02f07c51a47-a1b5bc36-1bbe-482f-a64a-c2d9cb606706-7439";
						this.ALogger=AWTLogManager.initialize(OfficeExtensibilityTenantID);
					}
					var eventProperties=new AWTEventProperties();
					eventProperties.setName("Office.Extensibility.OfficeJS."+tableName);
					for (var key in telemetryData) {
						if (key.toLowerCase() !=="table") {
							eventProperties.setProperty(key, telemetryData[key]);
						}
					}
					var today=new Date();
					eventProperties.setProperty("Date", today.toISOString());
					this.ALogger.logEvent(eventProperties);
				}
				catch (e) {
				}
			}, startAfterMs);
		};
		AriaLogger.prototype.logData=function (data) {
			if (this.isIUsageData(data)) {
				this.loadAriaScriptAndLog(data["Table"], data["Fields"]);
			}
			else {
				this.loadAriaScriptAndLog(data["Table"], data);
			}
		};
		return AriaLogger;
	})();
	OSFAriaLogger.AriaLogger=AriaLogger;
})(OSFAriaLogger || (OSFAriaLogger={}));
var OSFAppTelemetry;
(function (OSFAppTelemetry) {
	"use strict";
	var appInfo;
	var sessionId=OSF.OUtil.Guid.generateNewGuid();
	var osfControlAppCorrelationId="";
	var omexDomainRegex=new RegExp("^https?://store\\.office(ppe|-int)?\\.com/", "i");
	OSFAppTelemetry.enableTelemetry=true;
	;
	var AppInfo=(function () {
		function AppInfo() {
		}
		return AppInfo;
	})();
	var Event=(function () {
		function Event(name, handler) {
			this.name=name;
			this.handler=handler;
		}
		return Event;
	})();
	var AppStorage=(function () {
		function AppStorage() {
			this.clientIDKey="Office API client";
			this.logIdSetKey="Office App Log Id Set";
		}
		AppStorage.prototype.getClientId=function () {
			var clientId=this.getValue(this.clientIDKey);
			if (!clientId || clientId.length <=0 || clientId.length > 40) {
				clientId=OSF.OUtil.Guid.generateNewGuid();
				this.setValue(this.clientIDKey, clientId);
			}
			return clientId;
		};
		AppStorage.prototype.saveLog=function (logId, log) {
			var logIdSet=this.getValue(this.logIdSetKey);
			logIdSet=((logIdSet && logIdSet.length > 0) ? (logIdSet+";") : "")+logId;
			this.setValue(this.logIdSetKey, logIdSet);
			this.setValue(logId, log);
		};
		AppStorage.prototype.enumerateLog=function (callback, clean) {
			var logIdSet=this.getValue(this.logIdSetKey);
			if (logIdSet) {
				var ids=logIdSet.split(";");
				for (var id in ids) {
					var logId=ids[id];
					var log=this.getValue(logId);
					if (log) {
						if (callback) {
							callback(logId, log);
						}
						if (clean) {
							this.remove(logId);
						}
					}
				}
				if (clean) {
					this.remove(this.logIdSetKey);
				}
			}
		};
		AppStorage.prototype.getValue=function (key) {
			var osfLocalStorage=OSF.OUtil.getLocalStorage();
			var value="";
			if (osfLocalStorage) {
				value=osfLocalStorage.getItem(key);
			}
			return value;
		};
		AppStorage.prototype.setValue=function (key, value) {
			var osfLocalStorage=OSF.OUtil.getLocalStorage();
			if (osfLocalStorage) {
				osfLocalStorage.setItem(key, value);
			}
		};
		AppStorage.prototype.remove=function (key) {
			var osfLocalStorage=OSF.OUtil.getLocalStorage();
			if (osfLocalStorage) {
				try {
					osfLocalStorage.removeItem(key);
				}
				catch (ex) {
				}
			}
		};
		return AppStorage;
	})();
	var AppLogger=(function () {
		function AppLogger() {
		}
		AppLogger.prototype.LogData=function (data) {
			if (!OSF.Logger || !OSFAppTelemetry.enableTelemetry) {
				return;
			}
			try {
				OSFAriaLogger.AriaLogger.getInstance().logData(data);
			}
			catch (e) {
			}
		};
		AppLogger.prototype.LogRawData=function (log) {
			if (!OSF.Logger || !OSFAppTelemetry.enableTelemetry) {
				return;
			}
			try {
				OSFAriaLogger.AriaLogger.getInstance().logData(JSON.parse(log));
			}
			catch (e) {
			}
		};
		return AppLogger;
	})();
	function trimStringToLowerCase(input) {
		if (input) {
			input=input.replace(/[{}]/g, "").toLowerCase();
		}
		return (input || "");
	}
	function initialize(context) {
		if (!OSF.Logger) {
			return;
		}
		if (appInfo) {
			return;
		}
		appInfo=new AppInfo();
		if (context.get_hostFullVersion()) {
			appInfo.hostVersion=context.get_hostFullVersion();
		}
		else {
			appInfo.hostVersion=context.get_appVersion();
		}
		appInfo.appId=context.get_id();
		appInfo.host=context.get_appName();
		appInfo.browser=window.navigator.userAgent;
		appInfo.correlationId=trimStringToLowerCase(context.get_correlationId());
		appInfo.clientId=(new AppStorage()).getClientId();
		appInfo.appInstanceId=context.get_appInstanceId();
		if (appInfo.appInstanceId) {
			appInfo.appInstanceId=appInfo.appInstanceId.replace(/[{}]/g, "").toLowerCase();
		}
		appInfo.message=context.get_hostCustomMessage();
		appInfo.officeJSVersion=OSF.ConstantNames.FileVersion;
		appInfo.hostJSVersion="16.0.8616.1000";
		if (context._wacHostEnvironment) {
			appInfo.wacHostEnvironment=context._wacHostEnvironment;
		}
		if (context._isFromWacAutomation !==undefined && context._isFromWacAutomation !==null) {
			appInfo.isFromWacAutomation=context._isFromWacAutomation.toString().toLowerCase();
		}
		var docUrl=context.get_docUrl();
		appInfo.docUrl=omexDomainRegex.test(docUrl) ? docUrl : "";
		var url=location.href;
		if (url) {
			url=url.split("?")[0].split("#")[0];
		}
		appInfo.appURL=url;
		(function getUserIdAndAssetIdFromToken(token, appInfo) {
			var xmlContent;
			var parser;
			var xmlDoc;
			appInfo.assetId="";
			appInfo.userId="";
			try {
				xmlContent=decodeURIComponent(token);
				parser=new DOMParser();
				xmlDoc=parser.parseFromString(xmlContent, "text/xml");
				var cidNode=xmlDoc.getElementsByTagName("t")[0].attributes.getNamedItem("cid");
				var oidNode=xmlDoc.getElementsByTagName("t")[0].attributes.getNamedItem("oid");
				if (cidNode && cidNode.nodeValue) {
					appInfo.userId=cidNode.nodeValue;
				}
				else if (oidNode && oidNode.nodeValue) {
					appInfo.userId=oidNode.nodeValue;
				}
				appInfo.assetId=xmlDoc.getElementsByTagName("t")[0].attributes.getNamedItem("aid").nodeValue;
			}
			catch (e) {
			}
			finally {
				xmlContent=null;
				xmlDoc=null;
				parser=null;
			}
		})(context.get_eToken(), appInfo);
		(function handleLifecycle() {
			var startTime=new Date();
			var lastFocus=null;
			var focusTime=0;
			var finished=false;
			var adjustFocusTime=function () {
				if (document.hasFocus()) {
					if (lastFocus==null) {
						lastFocus=new Date();
					}
				}
				else if (lastFocus) {
					focusTime+=Math.abs((new Date()).getTime() - lastFocus.getTime());
					lastFocus=null;
				}
			};
			var eventList=[];
			eventList.push(new Event("focus", adjustFocusTime));
			eventList.push(new Event("blur", adjustFocusTime));
			eventList.push(new Event("focusout", adjustFocusTime));
			eventList.push(new Event("focusin", adjustFocusTime));
			var exitFunction=function () {
				for (var i=0; i < eventList.length; i++) {
					OSF.OUtil.removeEventListener(window, eventList[i].name, eventList[i].handler);
				}
				eventList.length=0;
				if (!finished) {
					if (document.hasFocus() && lastFocus) {
						focusTime+=Math.abs((new Date()).getTime() - lastFocus.getTime());
						lastFocus=null;
					}
					OSFAppTelemetry.onAppClosed(Math.abs((new Date()).getTime() - startTime.getTime()), focusTime);
					finished=true;
				}
			};
			eventList.push(new Event("beforeunload", exitFunction));
			eventList.push(new Event("unload", exitFunction));
			for (var i=0; i < eventList.length; i++) {
				OSF.OUtil.addEventListener(window, eventList[i].name, eventList[i].handler);
			}
			adjustFocusTime();
		})();
		OSFAppTelemetry.onAppActivated();
	}
	OSFAppTelemetry.initialize=initialize;
	function onAppActivated() {
		if (!appInfo) {
			return;
		}
		(new AppStorage()).enumerateLog(function (id, log) { return (new AppLogger()).LogRawData(log); }, true);
		var data=new OSFLog.AppActivatedUsageData();
		data.SessionId=sessionId;
		data.AppId=appInfo.appId;
		data.AssetId=appInfo.assetId;
		data.AppURL=appInfo.appURL;
		data.UserId=appInfo.userId;
		data.ClientId=appInfo.clientId;
		data.Browser=appInfo.browser;
		data.Host=appInfo.host;
		data.HostVersion=appInfo.hostVersion;
		data.CorrelationId=trimStringToLowerCase(appInfo.correlationId);
		data.AppSizeWidth=window.innerWidth;
		data.AppSizeHeight=window.innerHeight;
		data.AppInstanceId=appInfo.appInstanceId;
		data.Message=appInfo.message;
		data.DocUrl=appInfo.docUrl;
		data.OfficeJSVersion=appInfo.officeJSVersion;
		data.HostJSVersion=appInfo.hostJSVersion;
		if (appInfo.wacHostEnvironment) {
			data.WacHostEnvironment=appInfo.wacHostEnvironment;
		}
		if (appInfo.isFromWacAutomation !==undefined && appInfo.isFromWacAutomation !==null) {
			data.IsFromWacAutomation=appInfo.isFromWacAutomation;
		}
		(new AppLogger()).LogData(data);
		setTimeout(function () {
			if (!OSF.Logger) {
				return;
			}
			OSF.Logger.allowUploadingData();
		}, 100);
	}
	OSFAppTelemetry.onAppActivated=onAppActivated;
	function onScriptDone(scriptId, msStartTime, msResponseTime, appCorrelationId) {
		var data=new OSFLog.ScriptLoadUsageData();
		data.CorrelationId=trimStringToLowerCase(appCorrelationId);
		data.SessionId=sessionId;
		data.ScriptId=scriptId;
		data.StartTime=msStartTime;
		data.ResponseTime=msResponseTime;
		(new AppLogger()).LogData(data);
	}
	OSFAppTelemetry.onScriptDone=onScriptDone;
	function onCallDone(apiType, id, parameters, msResponseTime, errorType) {
		if (!appInfo) {
			return;
		}
		var data=new OSFLog.APIUsageUsageData();
		data.CorrelationId=trimStringToLowerCase(osfControlAppCorrelationId);
		data.SessionId=sessionId;
		data.APIType=apiType;
		data.APIID=id;
		data.Parameters=parameters;
		data.ResponseTime=msResponseTime;
		data.ErrorType=errorType;
		(new AppLogger()).LogData(data);
	}
	OSFAppTelemetry.onCallDone=onCallDone;
	;
	function onMethodDone(id, args, msResponseTime, errorType) {
		var parameters=null;
		if (args) {
			if (typeof args=="number") {
				parameters=String(args);
			}
			else if (typeof args==="object") {
				for (var index in args) {
					if (parameters !==null) {
						parameters+=",";
					}
					else {
						parameters="";
					}
					if (typeof args[index]=="number") {
						parameters+=String(args[index]);
					}
				}
			}
			else {
				parameters="";
			}
		}
		OSF.AppTelemetry.onCallDone("method", id, parameters, msResponseTime, errorType);
	}
	OSFAppTelemetry.onMethodDone=onMethodDone;
	function onPropertyDone(propertyName, msResponseTime) {
		OSF.AppTelemetry.onCallDone("property", -1, propertyName, msResponseTime);
	}
	OSFAppTelemetry.onPropertyDone=onPropertyDone;
	function onEventDone(id, errorType) {
		OSF.AppTelemetry.onCallDone("event", id, null, 0, errorType);
	}
	OSFAppTelemetry.onEventDone=onEventDone;
	function onRegisterDone(register, id, msResponseTime, errorType) {
		OSF.AppTelemetry.onCallDone(register ? "registerevent" : "unregisterevent", id, null, msResponseTime, errorType);
	}
	OSFAppTelemetry.onRegisterDone=onRegisterDone;
	function onAppClosed(openTime, focusTime) {
		if (!appInfo) {
			return;
		}
		var data=new OSFLog.AppClosedUsageData();
		data.CorrelationId=trimStringToLowerCase(osfControlAppCorrelationId);
		data.SessionId=sessionId;
		data.FocusTime=focusTime;
		data.OpenTime=openTime;
		data.AppSizeFinalWidth=window.innerWidth;
		data.AppSizeFinalHeight=window.innerHeight;
		(new AppStorage()).saveLog(sessionId, data.SerializeRow());
	}
	OSFAppTelemetry.onAppClosed=onAppClosed;
	function setOsfControlAppCorrelationId(correlationId) {
		osfControlAppCorrelationId=trimStringToLowerCase(correlationId);
	}
	OSFAppTelemetry.setOsfControlAppCorrelationId=setOsfControlAppCorrelationId;
	function doAppInitializationLogging(isException, message) {
		var data=new OSFLog.AppInitializationUsageData();
		data.CorrelationId=trimStringToLowerCase(osfControlAppCorrelationId);
		data.SessionId=sessionId;
		data.SuccessCode=isException ? 1 : 0;
		data.Message=message;
		(new AppLogger()).LogData(data);
	}
	OSFAppTelemetry.doAppInitializationLogging=doAppInitializationLogging;
	function logAppCommonMessage(message) {
		doAppInitializationLogging(false, message);
	}
	OSFAppTelemetry.logAppCommonMessage=logAppCommonMessage;
	function logAppException(errorMessage) {
		doAppInitializationLogging(true, errorMessage);
	}
	OSFAppTelemetry.logAppException=logAppException;
	OSF.AppTelemetry=OSFAppTelemetry;
})(OSFAppTelemetry || (OSFAppTelemetry={}));
Microsoft.Office.WebExtension.EventType={};
OSF.EventDispatch=function OSF_EventDispatch(eventTypes) {
	this._eventHandlers={};
	this._objectEventHandlers={};
	this._queuedEventsArgs={};
	for (var entry in eventTypes) {
		var eventType=eventTypes[entry];
		var isObjectEvent=(eventType=="objectDeleted" || eventType=="objectSelectionChanged" || eventType=="objectDataChanged" || eventType=="contentControlAdded");
		if (!isObjectEvent)
			this._eventHandlers[eventType]=[];
		else
			this._objectEventHandlers[eventType]={};
		this._queuedEventsArgs[eventType]=[];
	}
};
OSF.EventDispatch.prototype={
	getSupportedEvents: function OSF_EventDispatch$getSupportedEvents() {
		var events=[];
		for (var eventName in this._eventHandlers)
			events.push(eventName);
		for (var eventName in this._objectEventHandlers)
			events.push(eventName);
		return events;
	},
	supportsEvent: function OSF_EventDispatch$supportsEvent(event) {
		for (var eventName in this._eventHandlers) {
			if (event==eventName)
				return true;
		}
		for (var eventName in this._objectEventHandlers) {
			if (event==eventName)
				return true;
		}
		return false;
	},
	hasEventHandler: function OSF_EventDispatch$hasEventHandler(eventType, handler) {
		var handlers=this._eventHandlers[eventType];
		if (handlers && handlers.length > 0) {
			for (var h in handlers) {
				if (handlers[h]===handler)
					return true;
			}
		}
		return false;
	},
	hasObjectEventHandler: function OSF_EventDispatch$hasObjectEventHandler(eventType, objectId, handler) {
		var handlers=this._objectEventHandlers[eventType];
		if (handlers !=null) {
			var _handlers=handlers[objectId];
			for (var i=0; _handlers !=null && i < _handlers.length; i++) {
				if (_handlers[i]===handler)
					return true;
			}
		}
		return false;
	},
	addEventHandler: function OSF_EventDispatch$addEventHandler(eventType, handler) {
		if (typeof handler !="function") {
			return false;
		}
		var handlers=this._eventHandlers[eventType];
		if (handlers && !this.hasEventHandler(eventType, handler)) {
			handlers.push(handler);
			return true;
		}
		else {
			return false;
		}
	},
	addObjectEventHandler: function OSF_EventDispatch$addObjectEventHandler(eventType, objectId, handler) {
		if (typeof handler !="function") {
			return false;
		}
		var handlers=this._objectEventHandlers[eventType];
		if (handlers && !this.hasObjectEventHandler(eventType, objectId, handler)) {
			if (handlers[objectId]==null)
				handlers[objectId]=[];
			handlers[objectId].push(handler);
			return true;
		}
		return false;
	},
	addEventHandlerAndFireQueuedEvent: function OSF_EventDispatch$addEventHandlerAndFireQueuedEvent(eventType, handler) {
		var handlers=this._eventHandlers[eventType];
		var isFirstHandler=handlers.length==0;
		var succeed=this.addEventHandler(eventType, handler);
		if (isFirstHandler && succeed) {
			this.fireQueuedEvent(eventType);
		}
		return succeed;
	},
	removeEventHandler: function OSF_EventDispatch$removeEventHandler(eventType, handler) {
		var handlers=this._eventHandlers[eventType];
		if (handlers && handlers.length > 0) {
			for (var index=0; index < handlers.length; index++) {
				if (handlers[index]===handler) {
					handlers.splice(index, 1);
					return true;
				}
			}
		}
		return false;
	},
	removeObjectEventHandler: function OSF_EventDispatch$removeObjectEventHandler(eventType, objectId, handler) {
		var handlers=this._objectEventHandlers[eventType];
		if (handlers !=null) {
			var _handlers=handlers[objectId];
			for (var i=0; _handlers !=null && i < _handlers.length; i++) {
				if (_handlers[i]===handler) {
					_handlers.splice(i, 1);
					return true;
				}
			}
		}
		return false;
	},
	clearEventHandlers: function OSF_EventDispatch$clearEventHandlers(eventType) {
		if (typeof this._eventHandlers[eventType] !="undefined" && this._eventHandlers[eventType].length > 0) {
			this._eventHandlers[eventType]=[];
			return true;
		}
		return false;
	},
	clearObjectEventHandlers: function OSF_EventDispatch$clearObjectEventHandlers(eventType, objectId) {
		if (this._objectEventHandlers[eventType] !=null && this._objectEventHandlers[eventType][objectId] !=null) {
			this._objectEventHandlers[eventType][objectId]=[];
			return true;
		}
		return false;
	},
	getEventHandlerCount: function OSF_EventDispatch$getEventHandlerCount(eventType) {
		return this._eventHandlers[eventType] !=undefined ? this._eventHandlers[eventType].length : -1;
	},
	getObjectEventHandlerCount: function OSF_EventDispatch$getObjectEventHandlerCount(eventType, objectId) {
		if (this._objectEventHandlers[eventType]==null || this._objectEventHandlers[eventType][objectId]==null)
			return 0;
		return this._objectEventHandlers[eventType][objectId].length;
	},
	fireEvent: function OSF_EventDispatch$fireEvent(eventArgs) {
		if (eventArgs.type==undefined)
			return false;
		var eventType=eventArgs.type;
		if (eventType && this._eventHandlers[eventType]) {
			var eventHandlers=this._eventHandlers[eventType];
			for (var handler in eventHandlers)
				eventHandlers[handler](eventArgs);
			return true;
		}
		else {
			return false;
		}
	},
	fireObjectEvent: function OSF_EventDispatch$fireObjectEvent(objectId, eventArgs) {
		if (eventArgs.type==undefined)
			return false;
		var eventType=eventArgs.type;
		if (eventType && this._objectEventHandlers[eventType]) {
			var eventHandlers=this._objectEventHandlers[eventType];
			var _handlers=eventHandlers[objectId];
			if (_handlers !=null) {
				for (var i=0; i < _handlers.length; i++)
					_handlers[i](eventArgs);
				return true;
			}
		}
		return false;
	},
	fireOrQueueEvent: function OSF_EventDispatch$fireOrQueueEvent(eventArgs) {
		var eventType=eventArgs.type;
		if (eventType && this._eventHandlers[eventType]) {
			var eventHandlers=this._eventHandlers[eventType];
			var queuedEvents=this._queuedEventsArgs[eventType];
			if (eventHandlers.length==0) {
				queuedEvents.push(eventArgs);
			}
			else {
				this.fireEvent(eventArgs);
			}
			return true;
		}
		else {
			return false;
		}
	},
	fireQueuedEvent: function OSF_EventDispatch$queueEvent(eventType) {
		if (eventType && this._eventHandlers[eventType]) {
			var eventHandlers=this._eventHandlers[eventType];
			var queuedEvents=this._queuedEventsArgs[eventType];
			if (eventHandlers.length > 0) {
				var eventHandler=eventHandlers[0];
				while (queuedEvents.length > 0) {
					var eventArgs=queuedEvents.shift();
					eventHandler(eventArgs);
				}
				return true;
			}
		}
		return false;
	},
	clearQueuedEvent: function OSF_EventDispatch$clearQueuedEvent(eventType) {
		if (eventType && this._eventHandlers[eventType]) {
			var queuedEvents=this._queuedEventsArgs[eventType];
			if (queuedEvents) {
				this._queuedEventsArgs[eventType]=[];
			}
		}
	}
};
OSF.DDA.OMFactory=OSF.DDA.OMFactory || {};
OSF.DDA.OMFactory.manufactureEventArgs=function OSF_DDA_OMFactory$manufactureEventArgs(eventType, target, eventProperties) {
	var args;
	switch (eventType) {
		case Microsoft.Office.WebExtension.EventType.DocumentSelectionChanged:
			args=new OSF.DDA.DocumentSelectionChangedEventArgs(target);
			break;
		case Microsoft.Office.WebExtension.EventType.BindingSelectionChanged:
			args=new OSF.DDA.BindingSelectionChangedEventArgs(this.manufactureBinding(eventProperties, target.document), eventProperties[OSF.DDA.PropertyDescriptors.Subset]);
			break;
		case Microsoft.Office.WebExtension.EventType.BindingDataChanged:
			args=new OSF.DDA.BindingDataChangedEventArgs(this.manufactureBinding(eventProperties, target.document));
			break;
		case Microsoft.Office.WebExtension.EventType.SettingsChanged:
			args=new OSF.DDA.SettingsChangedEventArgs(target);
			break;
		case Microsoft.Office.WebExtension.EventType.ActiveViewChanged:
			args=new OSF.DDA.ActiveViewChangedEventArgs(eventProperties);
			break;
		case Microsoft.Office.WebExtension.EventType.OfficeThemeChanged:
			args=new OSF.DDA.Theming.OfficeThemeChangedEventArgs(eventProperties);
			break;
		case Microsoft.Office.WebExtension.EventType.DocumentThemeChanged:
			args=new OSF.DDA.Theming.DocumentThemeChangedEventArgs(eventProperties);
			break;
		case Microsoft.Office.WebExtension.EventType.AppCommandInvoked:
			args=OSF.DDA.AppCommand.AppCommandInvokedEventArgs.create(eventProperties);
			break;
		case Microsoft.Office.WebExtension.EventType.ObjectDeleted:
		case Microsoft.Office.WebExtension.EventType.ObjectSelectionChanged:
		case Microsoft.Office.WebExtension.EventType.ObjectDataChanged:
		case Microsoft.Office.WebExtension.EventType.ContentControlAdded:
			args=new OSF.DDA.ObjectEventArgs(eventType, eventProperties[Microsoft.Office.WebExtension.Parameters.Id]);
			break;
		case Microsoft.Office.WebExtension.EventType.RichApiMessage:
			args=new OSF.DDA.RichApiMessageEventArgs(eventType, eventProperties);
			break;
		case Microsoft.Office.WebExtension.EventType.DataNodeInserted:
			args=new OSF.DDA.NodeInsertedEventArgs(this.manufactureDataNode(eventProperties[OSF.DDA.DataNodeEventProperties.NewNode]), eventProperties[OSF.DDA.DataNodeEventProperties.InUndoRedo]);
			break;
		case Microsoft.Office.WebExtension.EventType.DataNodeReplaced:
			args=new OSF.DDA.NodeReplacedEventArgs(this.manufactureDataNode(eventProperties[OSF.DDA.DataNodeEventProperties.OldNode]), this.manufactureDataNode(eventProperties[OSF.DDA.DataNodeEventProperties.NewNode]), eventProperties[OSF.DDA.DataNodeEventProperties.InUndoRedo]);
			break;
		case Microsoft.Office.WebExtension.EventType.DataNodeDeleted:
			args=new OSF.DDA.NodeDeletedEventArgs(this.manufactureDataNode(eventProperties[OSF.DDA.DataNodeEventProperties.OldNode]), this.manufactureDataNode(eventProperties[OSF.DDA.DataNodeEventProperties.NextSiblingNode]), eventProperties[OSF.DDA.DataNodeEventProperties.InUndoRedo]);
			break;
		case Microsoft.Office.WebExtension.EventType.TaskSelectionChanged:
			args=new OSF.DDA.TaskSelectionChangedEventArgs(target);
			break;
		case Microsoft.Office.WebExtension.EventType.ResourceSelectionChanged:
			args=new OSF.DDA.ResourceSelectionChangedEventArgs(target);
			break;
		case Microsoft.Office.WebExtension.EventType.ViewSelectionChanged:
			args=new OSF.DDA.ViewSelectionChangedEventArgs(target);
			break;
		case Microsoft.Office.WebExtension.EventType.DialogMessageReceived:
			args=new OSF.DDA.DialogEventArgs(eventProperties);
			break;
		case Microsoft.Office.WebExtension.EventType.DialogParentMessageReceived:
			args=new OSF.DDA.DialogParentEventArgs(eventProperties);
			break;
		case Microsoft.Office.WebExtension.EventType.ItemChanged:
			if (OSF._OfficeAppFactory.getHostInfo()["hostType"]=="outlook" || OSF._OfficeAppFactory.getHostInfo()["hostType"]=="outlookwebapp") {
				args=new OSF.DDA.OlkItemSelectedChangedEventArgs(eventProperties);
				target.initialize(args["initialData"]);
				target.setCurrentItemNumber(args["itemNumber"].itemNumber);
			}
			else {
				throw OsfMsAjaxFactory.msAjaxError.argument(Microsoft.Office.WebExtension.Parameters.EventType, OSF.OUtil.formatString(Strings.OfficeOM.L_NotSupportedEventType, eventType));
			}
			break;
		case Microsoft.Office.WebExtension.EventType.RecipientsChanged:
			if (OSF._OfficeAppFactory.getHostInfo()["hostType"]=="outlook" || OSF._OfficeAppFactory.getHostInfo()["hostType"]=="outlookwebapp") {
				args=new OSF.DDA.OlkRecipientsChangedEventArgs(eventProperties);
			}
			else {
				throw OsfMsAjaxFactory.msAjaxError.argument(Microsoft.Office.WebExtension.Parameters.EventType, OSF.OUtil.formatString(Strings.OfficeOM.L_NotSupportedEventType, eventType));
			}
			break;
		case Microsoft.Office.WebExtension.EventType.AppointmentTimeChanged:
			if (OSF._OfficeAppFactory.getHostInfo()["hostType"]=="outlook" || OSF._OfficeAppFactory.getHostInfo()["hostType"]=="outlookwebapp") {
				args=new OSF.DDA.OlkAppointmentTimeChangedEventArgs(eventProperties);
			}
			else {
				throw OsfMsAjaxFactory.msAjaxError.argument(Microsoft.Office.WebExtension.Parameters.EventType, OSF.OUtil.formatString(Strings.OfficeOM.L_NotSupportedEventType, eventType));
			}
			break;
		default:
			throw OsfMsAjaxFactory.msAjaxError.argument(Microsoft.Office.WebExtension.Parameters.EventType, OSF.OUtil.formatString(Strings.OfficeOM.L_NotSupportedEventType, eventType));
	}
	return args;
};
OSF.DDA.AsyncMethodNames.addNames({
	AddHandlerAsync: "addHandlerAsync",
	RemoveHandlerAsync: "removeHandlerAsync"
});
OSF.DDA.AsyncMethodCalls.define({
	method: OSF.DDA.AsyncMethodNames.AddHandlerAsync,
	requiredArguments: [{
			"name": Microsoft.Office.WebExtension.Parameters.EventType,
			"enum": Microsoft.Office.WebExtension.EventType,
			"verify": function (eventType, caller, eventDispatch) { return eventDispatch.supportsEvent(eventType); }
		},
		{
			"name": Microsoft.Office.WebExtension.Parameters.Handler,
			"types": ["function"]
		}
	],
	supportedOptions: [],
	privateStateCallbacks: []
});
OSF.DDA.AsyncMethodCalls.define({
	method: OSF.DDA.AsyncMethodNames.RemoveHandlerAsync,
	requiredArguments: [
		{
			"name": Microsoft.Office.WebExtension.Parameters.EventType,
			"enum": Microsoft.Office.WebExtension.EventType,
			"verify": function (eventType, caller, eventDispatch) { return eventDispatch.supportsEvent(eventType); }
		}
	],
	supportedOptions: [
		{
			name: Microsoft.Office.WebExtension.Parameters.Handler,
			value: {
				"types": ["function", "object"],
				"defaultValue": null
			}
		}
	],
	privateStateCallbacks: []
});
OSF.DialogShownStatus={ hasDialogShown: false, isWindowDialog: false };
OSF.OUtil.augmentList(OSF.DDA.EventDescriptors, {
	DialogMessageReceivedEvent: "DialogMessageReceivedEvent"
});
OSF.OUtil.augmentList(Microsoft.Office.WebExtension.EventType, {
	DialogMessageReceived: "dialogMessageReceived",
	DialogEventReceived: "dialogEventReceived"
});
OSF.OUtil.augmentList(OSF.DDA.PropertyDescriptors, {
	MessageType: "messageType",
	MessageContent: "messageContent"
});
OSF.DDA.DialogEventType={};
OSF.OUtil.augmentList(OSF.DDA.DialogEventType, {
	DialogClosed: "dialogClosed",
	NavigationFailed: "naviationFailed"
});
OSF.DDA.AsyncMethodNames.addNames({
	DisplayDialogAsync: "displayDialogAsync",
	CloseAsync: "close"
});
OSF.DDA.SyncMethodNames.addNames({
	MessageParent: "messageParent",
	AddMessageHandler: "addEventHandler",
	SendMessage: "sendMessage"
});
OSF.DDA.UI.ParentUI=function OSF_DDA_ParentUI() {
	var eventDispatch=new OSF.EventDispatch([
		Microsoft.Office.WebExtension.EventType.DialogMessageReceived,
		Microsoft.Office.WebExtension.EventType.DialogEventReceived,
		Microsoft.Office.WebExtension.EventType.DialogParentMessageReceived
	]);
	var openDialogName=OSF.DDA.AsyncMethodNames.DisplayDialogAsync.displayName;
	var target=this;
	if (!target[openDialogName]) {
		OSF.OUtil.defineEnumerableProperty(target, openDialogName, {
			value: function () {
				var openDialog=OSF._OfficeAppFactory.getHostFacade()[OSF.DDA.DispIdHost.Methods.OpenDialog];
				openDialog(arguments, eventDispatch, target);
			}
		});
	}
	OSF.OUtil.finalizeProperties(this);
};
OSF.DDA.UI.ChildUI=function OSF_DDA_ChildUI(isPopupWindow) {
	var messageParentName=OSF.DDA.SyncMethodNames.MessageParent.displayName;
	var target=this;
	if (!target[messageParentName]) {
		OSF.OUtil.defineEnumerableProperty(target, messageParentName, {
			value: function () {
				var messageParent=OSF._OfficeAppFactory.getHostFacade()[OSF.DDA.DispIdHost.Methods.MessageParent];
				return messageParent(arguments, target);
			}
		});
	}
	var addEventHandler=OSF.DDA.SyncMethodNames.AddMessageHandler.displayName;
	if (!target[addEventHandler] && typeof OSF.DialogParentMessageEventDispatch !="undefined") {
		OSF.DDA.DispIdHost.addEventSupport(target, OSF.DialogParentMessageEventDispatch, isPopupWindow);
	}
	OSF.OUtil.finalizeProperties(this);
};
OSF.DialogHandler=function OSF_DialogHandler() { };
OSF.DDA.DialogEventArgs=function OSF_DDA_DialogEventArgs(message) {
	if (message[OSF.DDA.PropertyDescriptors.MessageType]==OSF.DialogMessageType.DialogMessageReceived) {
		OSF.OUtil.defineEnumerableProperties(this, {
			"type": {
				value: Microsoft.Office.WebExtension.EventType.DialogMessageReceived
			},
			"message": {
				value: message[OSF.DDA.PropertyDescriptors.MessageContent]
			}
		});
	}
	else {
		OSF.OUtil.defineEnumerableProperties(this, {
			"type": {
				value: Microsoft.Office.WebExtension.EventType.DialogEventReceived
			},
			"error": {
				value: message[OSF.DDA.PropertyDescriptors.MessageType]
			}
		});
	}
};
OSF.DDA.DialogParentEventArgs=function OSF_DDA_DialogParentEventArgs(message) {
	OSF.OUtil.defineEnumerableProperties(this, {
		"type": {
			value: Microsoft.Office.WebExtension.EventType.DialogParentMessageReceived
		},
		"message": {
			value: message[OSF.DDA.PropertyDescriptors.MessageContent]
		}
	});
};
OSF.DDA.AsyncMethodCalls.define({
	method: OSF.DDA.AsyncMethodNames.DisplayDialogAsync,
	requiredArguments: [
		{
			"name": Microsoft.Office.WebExtension.Parameters.Url,
			"types": ["string"]
		}
	],
	supportedOptions: [
		{
			name: Microsoft.Office.WebExtension.Parameters.Width,
			value: {
				"types": ["number"],
				"defaultValue": 99
			}
		},
		{
			name: Microsoft.Office.WebExtension.Parameters.Height,
			value: {
				"types": ["number"],
				"defaultValue": 99
			}
		},
		{
			name: Microsoft.Office.WebExtension.Parameters.RequireHTTPs,
			value: {
				"types": ["boolean"],
				"defaultValue": true
			}
		},
		{
			name: Microsoft.Office.WebExtension.Parameters.DisplayInIframe,
			value: {
				"types": ["boolean"],
				"defaultValue": false
			}
		},
		{
			name: Microsoft.Office.WebExtension.Parameters.HideTitle,
			value: {
				"types": ["boolean"],
				"defaultValue": false
			}
		}
	],
	privateStateCallbacks: [],
	onSucceeded: function (args, caller, callArgs) {
		var targetId=args[Microsoft.Office.WebExtension.Parameters.Id];
		var eventDispatch=args[Microsoft.Office.WebExtension.Parameters.Data];
		var dialog=new OSF.DialogHandler();
		var closeDialog=OSF.DDA.AsyncMethodNames.CloseAsync.displayName;
		OSF.OUtil.defineEnumerableProperty(dialog, closeDialog, {
			value: function () {
				var closeDialogfunction=OSF._OfficeAppFactory.getHostFacade()[OSF.DDA.DispIdHost.Methods.CloseDialog];
				closeDialogfunction(arguments, targetId, eventDispatch, dialog);
			}
		});
		var addHandler=OSF.DDA.SyncMethodNames.AddMessageHandler.displayName;
		OSF.OUtil.defineEnumerableProperty(dialog, addHandler, {
			value: function () {
				var syncMethodCall=OSF.DDA.SyncMethodCalls[OSF.DDA.SyncMethodNames.AddMessageHandler.id];
				var callArgs=syncMethodCall.verifyAndExtractCall(arguments, dialog, eventDispatch);
				var eventType=callArgs[Microsoft.Office.WebExtension.Parameters.EventType];
				var handler=callArgs[Microsoft.Office.WebExtension.Parameters.Handler];
				return eventDispatch.addEventHandlerAndFireQueuedEvent(eventType, handler);
			}
		});
		var sendMessage=OSF.DDA.SyncMethodNames.SendMessage.displayName;
		OSF.OUtil.defineEnumerableProperty(dialog, sendMessage, {
			value: function () {
				var execute=OSF._OfficeAppFactory.getHostFacade()[OSF.DDA.DispIdHost.Methods.SendMessage];
				return execute(arguments, eventDispatch, dialog);
			}
		});
		return dialog;
	},
	checkCallArgs: function (callArgs, caller, stateInfo) {
		if (callArgs[Microsoft.Office.WebExtension.Parameters.Width] <=0) {
			callArgs[Microsoft.Office.WebExtension.Parameters.Width]=1;
		}
		if (callArgs[Microsoft.Office.WebExtension.Parameters.Width] > 100) {
			callArgs[Microsoft.Office.WebExtension.Parameters.Width]=99;
		}
		if (callArgs[Microsoft.Office.WebExtension.Parameters.Height] <=0) {
			callArgs[Microsoft.Office.WebExtension.Parameters.Height]=1;
		}
		if (callArgs[Microsoft.Office.WebExtension.Parameters.Height] > 100) {
			callArgs[Microsoft.Office.WebExtension.Parameters.Height]=99;
		}
		if (!callArgs[Microsoft.Office.WebExtension.Parameters.RequireHTTPs]) {
			callArgs[Microsoft.Office.WebExtension.Parameters.RequireHTTPs]=true;
		}
		return callArgs;
	}
});
OSF.DDA.AsyncMethodCalls.define({
	method: OSF.DDA.AsyncMethodNames.CloseAsync,
	requiredArguments: [],
	supportedOptions: [],
	privateStateCallbacks: []
});
OSF.DDA.SyncMethodCalls.define({
	method: OSF.DDA.SyncMethodNames.MessageParent,
	requiredArguments: [
		{
			"name": Microsoft.Office.WebExtension.Parameters.MessageToParent,
			"types": ["string", "number", "boolean"]
		}
	],
	supportedOptions: []
});
OSF.DDA.SyncMethodCalls.define({
	method: OSF.DDA.SyncMethodNames.AddMessageHandler,
	requiredArguments: [
		{
			"name": Microsoft.Office.WebExtension.Parameters.EventType,
			"enum": Microsoft.Office.WebExtension.EventType,
			"verify": function (eventType, caller, eventDispatch) { return eventDispatch.supportsEvent(eventType); }
		},
		{
			"name": Microsoft.Office.WebExtension.Parameters.Handler,
			"types": ["function"]
		}
	],
	supportedOptions: []
});
OSF.DDA.SyncMethodCalls.define({
	method: OSF.DDA.SyncMethodNames.SendMessage,
	requiredArguments: [
		{
			"name": Microsoft.Office.WebExtension.Parameters.MessageContent,
			"types": ["string"]
		}
	],
	supportedOptions: [],
	privateStateCallbacks: []
});
OSF.DDA.SafeArray.Delegate.openDialog=function OSF_DDA_SafeArray_Delegate$OpenDialog(args) {
	try {
		if (args.onCalling) {
			args.onCalling();
		}
		var callback=OSF.DDA.SafeArray.Delegate._getOnAfterRegisterEvent(true, args);
		OSF.ClientHostController.openDialog(args.dispId, args.targetId, function OSF_DDA_SafeArrayDelegate$RegisterEventAsync_OnEvent(eventDispId, payload) {
			if (args.onEvent) {
				args.onEvent(payload);
			}
			if (OSF.AppTelemetry) {
				OSF.AppTelemetry.onEventDone(args.dispId);
			}
		}, callback);
	}
	catch (ex) {
		OSF.DDA.SafeArray.Delegate._onException(ex, args);
	}
};
OSF.DDA.SafeArray.Delegate.closeDialog=function OSF_DDA_SafeArray_Delegate$CloseDialog(args) {
	if (args.onCalling) {
		args.onCalling();
	}
	var callback=OSF.DDA.SafeArray.Delegate._getOnAfterRegisterEvent(false, args);
	try {
		OSF.ClientHostController.closeDialog(args.dispId, args.targetId, callback);
	}
	catch (ex) {
		OSF.DDA.SafeArray.Delegate._onException(ex, args);
	}
};
OSF.DDA.SafeArray.Delegate.messageParent=function OSF_DDA_SafeArray_Delegate$MessageParent(args) {
	try {
		if (args.onCalling) {
			args.onCalling();
		}
		var startTime=(new Date()).getTime();
		var result=OSF.ClientHostController.messageParent(args.hostCallArgs);
		if (args.onReceiving) {
			args.onReceiving();
		}
		if (OSF.AppTelemetry) {
			OSF.AppTelemetry.onMethodDone(args.dispId, args.hostCallArgs, Math.abs((new Date()).getTime() - startTime), result);
		}
		return result;
	}
	catch (ex) {
		return OSF.DDA.SafeArray.Delegate._onExceptionSyncMethod(ex);
	}
};
OSF.DDA.SafeArray.Delegate.ParameterMap.define({
	type: OSF.DDA.EventDispId.dispidDialogMessageReceivedEvent,
	fromHost: [
		{ name: OSF.DDA.EventDescriptors.DialogMessageReceivedEvent, value: OSF.DDA.SafeArray.Delegate.ParameterMap.self }
	],
	isComplexType: true
});
OSF.DDA.SafeArray.Delegate.ParameterMap.define({
	type: OSF.DDA.EventDescriptors.DialogMessageReceivedEvent,
	fromHost: [
		{ name: OSF.DDA.PropertyDescriptors.MessageType, value: 0 },
		{ name: OSF.DDA.PropertyDescriptors.MessageContent, value: 1 }
	],
	isComplexType: true
});
OSF.DDA.SafeArray.Delegate.sendMessage=function OSF_DDA_SafeArray_Delegate$SendMessage(args) {
	try {
		if (args.onCalling) {
			args.onCalling();
		}
		var startTime=(new Date()).getTime();
		var result=OSF.ClientHostController.sendMessage(args.hostCallArgs);
		if (args.onReceiving) {
			args.onReceiving();
		}
		return result;
	}
	catch (ex) {
		return OSF.DDA.SafeArray.Delegate._onExceptionSyncMethod(ex);
	}
};
Microsoft.Office.WebExtension.TableData=function Microsoft_Office_WebExtension_TableData(rows, headers) {
	function fixData(data) {
		if (data==null || data==undefined) {
			return null;
		}
		try {
			for (var dim=OSF.DDA.DataCoercion.findArrayDimensionality(data, 2); dim < 2; dim++) {
				data=[data];
			}
			return data;
		}
		catch (ex) {
		}
	}
	;
	OSF.OUtil.defineEnumerableProperties(this, {
		"headers": {
			get: function () { return headers; },
			set: function (value) {
				headers=fixData(value);
			}
		},
		"rows": {
			get: function () { return rows; },
			set: function (value) {
				rows=(value==null || (OSF.OUtil.isArray(value) && (value.length==0))) ?
					[] :
					fixData(value);
			}
		}
	});
	this.headers=headers;
	this.rows=rows;
};
OSF.DDA.OMFactory=OSF.DDA.OMFactory || {};
OSF.DDA.OMFactory.manufactureTableData=function OSF_DDA_OMFactory$manufactureTableData(tableDataProperties) {
	return new Microsoft.Office.WebExtension.TableData(tableDataProperties[OSF.DDA.TableDataProperties.TableRows], tableDataProperties[OSF.DDA.TableDataProperties.TableHeaders]);
};
Microsoft.Office.WebExtension.CoercionType={
	Text: "text",
	Matrix: "matrix",
	Table: "table"
};
OSF.DDA.DataCoercion=(function OSF_DDA_DataCoercion() {
	return {
		findArrayDimensionality: function OSF_DDA_DataCoercion$findArrayDimensionality(obj) {
			if (OSF.OUtil.isArray(obj)) {
				var dim=0;
				for (var index=0; index < obj.length; index++) {
					dim=Math.max(dim, OSF.DDA.DataCoercion.findArrayDimensionality(obj[index]));
				}
				return dim+1;
			}
			else {
				return 0;
			}
		},
		getCoercionDefaultForBinding: function OSF_DDA_DataCoercion$getCoercionDefaultForBinding(bindingType) {
			switch (bindingType) {
				case Microsoft.Office.WebExtension.BindingType.Matrix: return Microsoft.Office.WebExtension.CoercionType.Matrix;
				case Microsoft.Office.WebExtension.BindingType.Table: return Microsoft.Office.WebExtension.CoercionType.Table;
				case Microsoft.Office.WebExtension.BindingType.Text:
				default:
					return Microsoft.Office.WebExtension.CoercionType.Text;
			}
		},
		getBindingDefaultForCoercion: function OSF_DDA_DataCoercion$getBindingDefaultForCoercion(coercionType) {
			switch (coercionType) {
				case Microsoft.Office.WebExtension.CoercionType.Matrix: return Microsoft.Office.WebExtension.BindingType.Matrix;
				case Microsoft.Office.WebExtension.CoercionType.Table: return Microsoft.Office.WebExtension.BindingType.Table;
				case Microsoft.Office.WebExtension.CoercionType.Text:
				case Microsoft.Office.WebExtension.CoercionType.Html:
				case Microsoft.Office.WebExtension.CoercionType.Ooxml:
				default:
					return Microsoft.Office.WebExtension.BindingType.Text;
			}
		},
		determineCoercionType: function OSF_DDA_DataCoercion$determineCoercionType(data) {
			if (data==null || data==undefined)
				return null;
			var sourceType=null;
			var runtimeType=typeof data;
			if (data.rows !==undefined) {
				sourceType=Microsoft.Office.WebExtension.CoercionType.Table;
			}
			else if (OSF.OUtil.isArray(data)) {
				sourceType=Microsoft.Office.WebExtension.CoercionType.Matrix;
			}
			else if (runtimeType=="string" || runtimeType=="number" || runtimeType=="boolean" || OSF.OUtil.isDate(data)) {
				sourceType=Microsoft.Office.WebExtension.CoercionType.Text;
			}
			else {
				throw OSF.DDA.ErrorCodeManager.errorCodes.ooeUnsupportedDataObject;
			}
			return sourceType;
		},
		coerceData: function OSF_DDA_DataCoercion$coerceData(data, destinationType, sourceType) {
			sourceType=sourceType || OSF.DDA.DataCoercion.determineCoercionType(data);
			if (sourceType && sourceType !=destinationType) {
				OSF.OUtil.writeProfilerMark(OSF.InternalPerfMarker.DataCoercionBegin);
				data=OSF.DDA.DataCoercion._coerceDataFromTable(destinationType, OSF.DDA.DataCoercion._coerceDataToTable(data, sourceType));
				OSF.OUtil.writeProfilerMark(OSF.InternalPerfMarker.DataCoercionEnd);
			}
			return data;
		},
		_matrixToText: function OSF_DDA_DataCoercion$_matrixToText(matrix) {
			if (matrix.length==1 && matrix[0].length==1)
				return ""+matrix[0][0];
			var val="";
			for (var i=0; i < matrix.length; i++) {
				val+=matrix[i].join("\t")+"\n";
			}
			return val.substring(0, val.length - 1);
		},
		_textToMatrix: function OSF_DDA_DataCoercion$_textToMatrix(text) {
			var ret=text.split("\n");
			for (var i=0; i < ret.length; i++)
				ret[i]=ret[i].split("\t");
			return ret;
		},
		_tableToText: function OSF_DDA_DataCoercion$_tableToText(table) {
			var headers="";
			if (table.headers !=null) {
				headers=OSF.DDA.DataCoercion._matrixToText([table.headers])+"\n";
			}
			var rows=OSF.DDA.DataCoercion._matrixToText(table.rows);
			if (rows=="") {
				headers=headers.substring(0, headers.length - 1);
			}
			return headers+rows;
		},
		_tableToMatrix: function OSF_DDA_DataCoercion$_tableToMatrix(table) {
			var matrix=table.rows;
			if (table.headers !=null) {
				matrix.unshift(table.headers);
			}
			return matrix;
		},
		_coerceDataFromTable: function OSF_DDA_DataCoercion$_coerceDataFromTable(coercionType, table) {
			var value;
			switch (coercionType) {
				case Microsoft.Office.WebExtension.CoercionType.Table:
					value=table;
					break;
				case Microsoft.Office.WebExtension.CoercionType.Matrix:
					value=OSF.DDA.DataCoercion._tableToMatrix(table);
					break;
				case Microsoft.Office.WebExtension.CoercionType.SlideRange:
					value=null;
					if (OSF.DDA.OMFactory.manufactureSlideRange) {
						value=OSF.DDA.OMFactory.manufactureSlideRange(OSF.DDA.DataCoercion._tableToText(table));
					}
					if (value==null) {
						value=OSF.DDA.DataCoercion._tableToText(table);
					}
					break;
				case Microsoft.Office.WebExtension.CoercionType.Text:
				case Microsoft.Office.WebExtension.CoercionType.Html:
				case Microsoft.Office.WebExtension.CoercionType.Ooxml:
				default:
					value=OSF.DDA.DataCoercion._tableToText(table);
					break;
			}
			return value;
		},
		_coerceDataToTable: function OSF_DDA_DataCoercion$_coerceDataToTable(data, sourceType) {
			if (sourceType==undefined) {
				sourceType=OSF.DDA.DataCoercion.determineCoercionType(data);
			}
			var value;
			switch (sourceType) {
				case Microsoft.Office.WebExtension.CoercionType.Table:
					value=data;
					break;
				case Microsoft.Office.WebExtension.CoercionType.Matrix:
					value=new Microsoft.Office.WebExtension.TableData(data);
					break;
				case Microsoft.Office.WebExtension.CoercionType.Text:
				case Microsoft.Office.WebExtension.CoercionType.Html:
				case Microsoft.Office.WebExtension.CoercionType.Ooxml:
				default:
					value=new Microsoft.Office.WebExtension.TableData(OSF.DDA.DataCoercion._textToMatrix(data));
					break;
			}
			return value;
		}
	};
})();
OSF.DDA.SafeArray.Delegate.ParameterMap.define({
	type: Microsoft.Office.WebExtension.Parameters.CoercionType,
	toHost: [
		{ name: Microsoft.Office.WebExtension.CoercionType.Text, value: 0 },
		{ name: Microsoft.Office.WebExtension.CoercionType.Matrix, value: 1 },
		{ name: Microsoft.Office.WebExtension.CoercionType.Table, value: 2 }
	]
});
OSF.DDA.AsyncMethodNames.addNames({
	GetSelectedDataAsync: "getSelectedDataAsync",
	SetSelectedDataAsync: "setSelectedDataAsync"
});
(function () {
	function processData(dataDescriptor, caller, callArgs) {
		var data=dataDescriptor[Microsoft.Office.WebExtension.Parameters.Data];
		if (OSF.DDA.TableDataProperties && data && (data[OSF.DDA.TableDataProperties.TableRows] !=undefined || data[OSF.DDA.TableDataProperties.TableHeaders] !=undefined)) {
			data=OSF.DDA.OMFactory.manufactureTableData(data);
		}
		data=OSF.DDA.DataCoercion.coerceData(data, callArgs[Microsoft.Office.WebExtension.Parameters.CoercionType]);
		return data==undefined ? null : data;
	}
	OSF.DDA.AsyncMethodCalls.define({
		method: OSF.DDA.AsyncMethodNames.GetSelectedDataAsync,
		requiredArguments: [
			{
				"name": Microsoft.Office.WebExtension.Parameters.CoercionType,
				"enum": Microsoft.Office.WebExtension.CoercionType
			}
		],
		supportedOptions: [
			{
				name: Microsoft.Office.WebExtension.Parameters.ValueFormat,
				value: {
					"enum": Microsoft.Office.WebExtension.ValueFormat,
					"defaultValue": Microsoft.Office.WebExtension.ValueFormat.Unformatted
				}
			},
			{
				name: Microsoft.Office.WebExtension.Parameters.FilterType,
				value: {
					"enum": Microsoft.Office.WebExtension.FilterType,
					"defaultValue": Microsoft.Office.WebExtension.FilterType.All
				}
			}
		],
		privateStateCallbacks: [],
		onSucceeded: processData
	});
	OSF.DDA.AsyncMethodCalls.define({
		method: OSF.DDA.AsyncMethodNames.SetSelectedDataAsync,
		requiredArguments: [
			{
				"name": Microsoft.Office.WebExtension.Parameters.Data,
				"types": ["string", "object", "number", "boolean"]
			}
		],
		supportedOptions: [
			{
				name: Microsoft.Office.WebExtension.Parameters.CoercionType,
				value: {
					"enum": Microsoft.Office.WebExtension.CoercionType,
					"calculate": function (requiredArgs) {
						return OSF.DDA.DataCoercion.determineCoercionType(requiredArgs[Microsoft.Office.WebExtension.Parameters.Data]);
					}
				}
			},
			{
				name: Microsoft.Office.WebExtension.Parameters.ImageLeft,
				value: {
					"types": ["number", "boolean"],
					"defaultValue": false
				}
			},
			{
				name: Microsoft.Office.WebExtension.Parameters.ImageTop,
				value: {
					"types": ["number", "boolean"],
					"defaultValue": false
				}
			},
			{
				name: Microsoft.Office.WebExtension.Parameters.ImageWidth,
				value: {
					"types": ["number", "boolean"],
					"defaultValue": false
				}
			},
			{
				name: Microsoft.Office.WebExtension.Parameters.ImageHeight,
				value: {
					"types": ["number", "boolean"],
					"defaultValue": false
				}
			}
		],
		privateStateCallbacks: []
	});
})();
OSF.DDA.SafeArray.Delegate.ParameterMap.define({
	type: OSF.DDA.MethodDispId.dispidGetSelectedDataMethod,
	fromHost: [
		{ name: Microsoft.Office.WebExtension.Parameters.Data, value: OSF.DDA.SafeArray.Delegate.ParameterMap.self }
	],
	toHost: [
		{ name: Microsoft.Office.WebExtension.Parameters.CoercionType, value: 0 },
		{ name: Microsoft.Office.WebExtension.Parameters.ValueFormat, value: 1 },
		{ name: Microsoft.Office.WebExtension.Parameters.FilterType, value: 2 }
	]
});
OSF.DDA.SafeArray.Delegate.ParameterMap.define({
	type: OSF.DDA.MethodDispId.dispidSetSelectedDataMethod,
	toHost: [
		{ name: Microsoft.Office.WebExtension.Parameters.CoercionType, value: 0 },
		{ name: Microsoft.Office.WebExtension.Parameters.Data, value: 1 },
		{ name: Microsoft.Office.WebExtension.Parameters.ImageLeft, value: 2 },
		{ name: Microsoft.Office.WebExtension.Parameters.ImageTop, value: 3 },
		{ name: Microsoft.Office.WebExtension.Parameters.ImageWidth, value: 4 },
		{ name: Microsoft.Office.WebExtension.Parameters.ImageHeight, value: 5 },
	]
});
OSF.DDA.SettingsManager={
	SerializedSettings: "serializedSettings",
	RefreshingSettings: "refreshingSettings",
	DateJSONPrefix: "Date(",
	DataJSONSuffix: ")",
	serializeSettings: function OSF_DDA_SettingsManager$serializeSettings(settingsCollection) {
		var ret={};
		for (var key in settingsCollection) {
			var value=settingsCollection[key];
			try {
				if (JSON) {
					value=JSON.stringify(value, function dateReplacer(k, v) {
						return OSF.OUtil.isDate(this[k]) ? OSF.DDA.SettingsManager.DateJSONPrefix+this[k].getTime()+OSF.DDA.SettingsManager.DataJSONSuffix : v;
					});
				}
				else {
					value=Sys.Serialization.JavaScriptSerializer.serialize(value);
				}
				ret[key]=value;
			}
			catch (ex) {
			}
		}
		return ret;
	},
	deserializeSettings: function OSF_DDA_SettingsManager$deserializeSettings(serializedSettings) {
		var ret={};
		serializedSettings=serializedSettings || {};
		for (var key in serializedSettings) {
			var value=serializedSettings[key];
			try {
				if (JSON) {
					value=JSON.parse(value, function dateReviver(k, v) {
						var d;
						if (typeof v==='string' && v && v.length > 6 && v.slice(0, 5)===OSF.DDA.SettingsManager.DateJSONPrefix && v.slice(-1)===OSF.DDA.SettingsManager.DataJSONSuffix) {
							d=new Date(parseInt(v.slice(5, -1)));
							if (d) {
								return d;
							}
						}
						return v;
					});
				}
				else {
					value=Sys.Serialization.JavaScriptSerializer.deserialize(value, true);
				}
				ret[key]=value;
			}
			catch (ex) {
			}
		}
		return ret;
	}
};
OSF.DDA.Settings=function OSF_DDA_Settings(settings) {
	settings=settings || {};
	var cacheSessionSettings=function (settings) {
		var osfSessionStorage=OSF.OUtil.getSessionStorage();
		if (osfSessionStorage) {
			var serializedSettings=OSF.DDA.SettingsManager.serializeSettings(settings);
			var storageSettings=JSON ? JSON.stringify(serializedSettings) : Sys.Serialization.JavaScriptSerializer.serialize(serializedSettings);
			osfSessionStorage.setItem(OSF._OfficeAppFactory.getCachedSessionSettingsKey(), storageSettings);
		}
	};
	OSF.OUtil.defineEnumerableProperties(this, {
		"get": {
			value: function OSF_DDA_Settings$get(name) {
				var e=Function._validateParams(arguments, [
					{ name: "name", type: String, mayBeNull: false }
				]);
				if (e)
					throw e;
				var setting=settings[name];
				return typeof (setting)==='undefined' ? null : setting;
			}
		},
		"set": {
			value: function OSF_DDA_Settings$set(name, value) {
				var e=Function._validateParams(arguments, [
					{ name: "name", type: String, mayBeNull: false },
					{ name: "value", mayBeNull: true }
				]);
				if (e)
					throw e;
				settings[name]=value;
				cacheSessionSettings(settings);
			}
		},
		"remove": {
			value: function OSF_DDA_Settings$remove(name) {
				var e=Function._validateParams(arguments, [
					{ name: "name", type: String, mayBeNull: false }
				]);
				if (e)
					throw e;
				delete settings[name];
				cacheSessionSettings(settings);
			}
		}
	});
	OSF.DDA.DispIdHost.addAsyncMethods(this, [OSF.DDA.AsyncMethodNames.SaveAsync], settings);
};
OSF.DDA.RefreshableSettings=function OSF_DDA_RefreshableSettings(settings) {
	OSF.DDA.RefreshableSettings.uber.constructor.call(this, settings);
	OSF.DDA.DispIdHost.addAsyncMethods(this, [OSF.DDA.AsyncMethodNames.RefreshAsync], settings);
	OSF.DDA.DispIdHost.addEventSupport(this, new OSF.EventDispatch([Microsoft.Office.WebExtension.EventType.SettingsChanged]));
};
OSF.OUtil.extend(OSF.DDA.RefreshableSettings, OSF.DDA.Settings);
OSF.OUtil.augmentList(Microsoft.Office.WebExtension.EventType, {
	SettingsChanged: "settingsChanged"
});
OSF.DDA.SettingsChangedEventArgs=function OSF_DDA_SettingsChangedEventArgs(settingsInstance) {
	OSF.OUtil.defineEnumerableProperties(this, {
		"type": {
			value: Microsoft.Office.WebExtension.EventType.SettingsChanged
		},
		"settings": {
			value: settingsInstance
		}
	});
};
OSF.DDA.AsyncMethodNames.addNames({
	RefreshAsync: "refreshAsync",
	SaveAsync: "saveAsync"
});
OSF.DDA.AsyncMethodCalls.define({
	method: OSF.DDA.AsyncMethodNames.RefreshAsync,
	requiredArguments: [],
	supportedOptions: [],
	privateStateCallbacks: [
		{
			name: OSF.DDA.SettingsManager.RefreshingSettings,
			value: function getRefreshingSettings(settingsInstance, settingsCollection) {
				return settingsCollection;
			}
		}
	],
	onSucceeded: function deserializeSettings(serializedSettingsDescriptor, refreshingSettings, refreshingSettingsArgs) {
		var serializedSettings=serializedSettingsDescriptor[OSF.DDA.SettingsManager.SerializedSettings];
		var newSettings=OSF.DDA.SettingsManager.deserializeSettings(serializedSettings);
		var oldSettings=refreshingSettingsArgs[OSF.DDA.SettingsManager.RefreshingSettings];
		for (var setting in oldSettings) {
			refreshingSettings.remove(setting);
		}
		for (var setting in newSettings) {
			refreshingSettings.set(setting, newSettings[setting]);
		}
		return refreshingSettings;
	}
});
OSF.DDA.AsyncMethodCalls.define({
	method: OSF.DDA.AsyncMethodNames.SaveAsync,
	requiredArguments: [],
	supportedOptions: [
		{
			name: Microsoft.Office.WebExtension.Parameters.OverwriteIfStale,
			value: {
				"types": ["boolean"],
				"defaultValue": true
			}
		}
	],
	privateStateCallbacks: [
		{
			name: OSF.DDA.SettingsManager.SerializedSettings,
			value: function serializeSettings(settingsInstance, settingsCollection) {
				return OSF.DDA.SettingsManager.serializeSettings(settingsCollection);
			}
		}
	]
});
OSF.DDA.SafeArray.Delegate.ParameterMap.define({
	type: OSF.DDA.MethodDispId.dispidLoadSettingsMethod,
	fromHost: [
		{ name: OSF.DDA.SettingsManager.SerializedSettings, value: OSF.DDA.SafeArray.Delegate.ParameterMap.self }
	]
});
OSF.DDA.SafeArray.Delegate.ParameterMap.define({
	type: OSF.DDA.MethodDispId.dispidSaveSettingsMethod,
	toHost: [
		{ name: OSF.DDA.SettingsManager.SerializedSettings, value: OSF.DDA.SettingsManager.SerializedSettings },
		{ name: Microsoft.Office.WebExtension.Parameters.OverwriteIfStale, value: Microsoft.Office.WebExtension.Parameters.OverwriteIfStale }
	]
});
OSF.DDA.SafeArray.Delegate.ParameterMap.define({ type: OSF.DDA.EventDispId.dispidSettingsChangedEvent });
Microsoft.Office.WebExtension.BindingType={
	Table: "table",
	Text: "text",
	Matrix: "matrix"
};
OSF.DDA.BindingProperties={
	Id: "BindingId",
	Type: Microsoft.Office.WebExtension.Parameters.BindingType
};
OSF.OUtil.augmentList(OSF.DDA.ListDescriptors, { BindingList: "BindingList" });
OSF.OUtil.augmentList(OSF.DDA.PropertyDescriptors, {
	Subset: "subset",
	BindingProperties: "BindingProperties"
});
OSF.DDA.ListType.setListType(OSF.DDA.ListDescriptors.BindingList, OSF.DDA.PropertyDescriptors.BindingProperties);
OSF.DDA.BindingPromise=function OSF_DDA_BindingPromise(bindingId, errorCallback) {
	this._id=bindingId;
	OSF.OUtil.defineEnumerableProperty(this, "onFail", {
		get: function () {
			return errorCallback;
		},
		set: function (onError) {
			var t=typeof onError;
			if (t !="undefined" && t !="function") {
				throw OSF.OUtil.formatString(Strings.OfficeOM.L_CallbackNotAFunction, t);
			}
			errorCallback=onError;
		}
	});
};
OSF.DDA.BindingPromise.prototype={
	_fetch: function OSF_DDA_BindingPromise$_fetch(onComplete) {
		if (this.binding) {
			if (onComplete)
				onComplete(this.binding);
		}
		else {
			if (!this._binding) {
				var me=this;
				Microsoft.Office.WebExtension.context.document.bindings.getByIdAsync(this._id, function (asyncResult) {
					if (asyncResult.status==Microsoft.Office.WebExtension.AsyncResultStatus.Succeeded) {
						OSF.OUtil.defineEnumerableProperty(me, "binding", {
							value: asyncResult.value
						});
						if (onComplete)
							onComplete(me.binding);
					}
					else {
						if (me.onFail)
							me.onFail(asyncResult);
					}
				});
			}
		}
		return this;
	},
	getDataAsync: function OSF_DDA_BindingPromise$getDataAsync() {
		var args=arguments;
		this._fetch(function onComplete(binding) { binding.getDataAsync.apply(binding, args); });
		return this;
	},
	setDataAsync: function OSF_DDA_BindingPromise$setDataAsync() {
		var args=arguments;
		this._fetch(function onComplete(binding) { binding.setDataAsync.apply(binding, args); });
		return this;
	},
	addHandlerAsync: function OSF_DDA_BindingPromise$addHandlerAsync() {
		var args=arguments;
		this._fetch(function onComplete(binding) { binding.addHandlerAsync.apply(binding, args); });
		return this;
	},
	removeHandlerAsync: function OSF_DDA_BindingPromise$removeHandlerAsync() {
		var args=arguments;
		this._fetch(function onComplete(binding) { binding.removeHandlerAsync.apply(binding, args); });
		return this;
	}
};
OSF.DDA.BindingFacade=function OSF_DDA_BindingFacade(docInstance) {
	this._eventDispatches=[];
	OSF.OUtil.defineEnumerableProperty(this, "document", {
		value: docInstance
	});
	var am=OSF.DDA.AsyncMethodNames;
	OSF.DDA.DispIdHost.addAsyncMethods(this, [
		am.AddFromSelectionAsync,
		am.AddFromNamedItemAsync,
		am.GetAllAsync,
		am.GetByIdAsync,
		am.ReleaseByIdAsync
	]);
};
OSF.DDA.UnknownBinding=function OSF_DDA_UknonwnBinding(id, docInstance) {
	OSF.OUtil.defineEnumerableProperties(this, {
		"document": { value: docInstance },
		"id": { value: id }
	});
};
OSF.DDA.Binding=function OSF_DDA_Binding(id, docInstance) {
	OSF.OUtil.defineEnumerableProperties(this, {
		"document": {
			value: docInstance
		},
		"id": {
			value: id
		}
	});
	var am=OSF.DDA.AsyncMethodNames;
	OSF.DDA.DispIdHost.addAsyncMethods(this, [
		am.GetDataAsync,
		am.SetDataAsync
	]);
	var et=Microsoft.Office.WebExtension.EventType;
	var bindingEventDispatches=docInstance.bindings._eventDispatches;
	if (!bindingEventDispatches[id]) {
		bindingEventDispatches[id]=new OSF.EventDispatch([
			et.BindingSelectionChanged,
			et.BindingDataChanged
		]);
	}
	var eventDispatch=bindingEventDispatches[id];
	OSF.DDA.DispIdHost.addEventSupport(this, eventDispatch);
};
OSF.DDA.generateBindingId=function OSF_DDA$GenerateBindingId() {
	return "UnnamedBinding_"+OSF.OUtil.getUniqueId()+"_"+new Date().getTime();
};
OSF.DDA.OMFactory=OSF.DDA.OMFactory || {};
OSF.DDA.OMFactory.manufactureBinding=function OSF_DDA_OMFactory$manufactureBinding(bindingProperties, containingDocument) {
	var id=bindingProperties[OSF.DDA.BindingProperties.Id];
	var rows=bindingProperties[OSF.DDA.BindingProperties.RowCount];
	var cols=bindingProperties[OSF.DDA.BindingProperties.ColumnCount];
	var hasHeaders=bindingProperties[OSF.DDA.BindingProperties.HasHeaders];
	var binding;
	switch (bindingProperties[OSF.DDA.BindingProperties.Type]) {
		case Microsoft.Office.WebExtension.BindingType.Text:
			binding=new OSF.DDA.TextBinding(id, containingDocument);
			break;
		case Microsoft.Office.WebExtension.BindingType.Matrix:
			binding=new OSF.DDA.MatrixBinding(id, containingDocument, rows, cols);
			break;
		case Microsoft.Office.WebExtension.BindingType.Table:
			var isExcelApp=function () {
				return (OSF.DDA.ExcelDocument)
					&& (Microsoft.Office.WebExtension.context.document)
					&& (Microsoft.Office.WebExtension.context.document instanceof OSF.DDA.ExcelDocument);
			};
			var tableBindingObject;
			if (isExcelApp() && OSF.DDA.ExcelTableBinding) {
				tableBindingObject=OSF.DDA.ExcelTableBinding;
			}
			else {
				tableBindingObject=OSF.DDA.TableBinding;
			}
			binding=new tableBindingObject(id, containingDocument, rows, cols, hasHeaders);
			break;
		default:
			binding=new OSF.DDA.UnknownBinding(id, containingDocument);
	}
	return binding;
};
OSF.DDA.AsyncMethodNames.addNames({
	AddFromSelectionAsync: "addFromSelectionAsync",
	AddFromNamedItemAsync: "addFromNamedItemAsync",
	GetAllAsync: "getAllAsync",
	GetByIdAsync: "getByIdAsync",
	ReleaseByIdAsync: "releaseByIdAsync",
	GetDataAsync: "getDataAsync",
	SetDataAsync: "setDataAsync"
});
(function () {
	function processBinding(bindingDescriptor) {
		return OSF.DDA.OMFactory.manufactureBinding(bindingDescriptor, Microsoft.Office.WebExtension.context.document);
	}
	function getObjectId(obj) { return obj.id; }
	function processData(dataDescriptor, caller, callArgs) {
		var data=dataDescriptor[Microsoft.Office.WebExtension.Parameters.Data];
		if (OSF.DDA.TableDataProperties && data && (data[OSF.DDA.TableDataProperties.TableRows] !=undefined || data[OSF.DDA.TableDataProperties.TableHeaders] !=undefined)) {
			data=OSF.DDA.OMFactory.manufactureTableData(data);
		}
		data=OSF.DDA.DataCoercion.coerceData(data, callArgs[Microsoft.Office.WebExtension.Parameters.CoercionType]);
		return data==undefined ? null : data;
	}
	OSF.DDA.AsyncMethodCalls.define({
		method: OSF.DDA.AsyncMethodNames.AddFromSelectionAsync,
		requiredArguments: [
			{
				"name": Microsoft.Office.WebExtension.Parameters.BindingType,
				"enum": Microsoft.Office.WebExtension.BindingType
			}
		],
		supportedOptions: [{
				name: Microsoft.Office.WebExtension.Parameters.Id,
				value: {
					"types": ["string"],
					"calculate": OSF.DDA.generateBindingId
				}
			},
			{
				name: Microsoft.Office.WebExtension.Parameters.Columns,
				value: {
					"types": ["object"],
					"defaultValue": null
				}
			}
		],
		privateStateCallbacks: [],
		onSucceeded: processBinding
	});
	OSF.DDA.AsyncMethodCalls.define({
		method: OSF.DDA.AsyncMethodNames.AddFromNamedItemAsync,
		requiredArguments: [{
				"name": Microsoft.Office.WebExtension.Parameters.ItemName,
				"types": ["string"]
			},
			{
				"name": Microsoft.Office.WebExtension.Parameters.BindingType,
				"enum": Microsoft.Office.WebExtension.BindingType
			}
		],
		supportedOptions: [{
				name: Microsoft.Office.WebExtension.Parameters.Id,
				value: {
					"types": ["string"],
					"calculate": OSF.DDA.generateBindingId
				}
			},
			{
				name: Microsoft.Office.WebExtension.Parameters.Columns,
				value: {
					"types": ["object"],
					"defaultValue": null
				}
			}
		],
		privateStateCallbacks: [
			{
				name: Microsoft.Office.WebExtension.Parameters.FailOnCollision,
				value: function () { return true; }
			}
		],
		onSucceeded: processBinding
	});
	OSF.DDA.AsyncMethodCalls.define({
		method: OSF.DDA.AsyncMethodNames.GetAllAsync,
		requiredArguments: [],
		supportedOptions: [],
		privateStateCallbacks: [],
		onSucceeded: function (response) { return OSF.OUtil.mapList(response[OSF.DDA.ListDescriptors.BindingList], processBinding); }
	});
	OSF.DDA.AsyncMethodCalls.define({
		method: OSF.DDA.AsyncMethodNames.GetByIdAsync,
		requiredArguments: [
			{
				"name": Microsoft.Office.WebExtension.Parameters.Id,
				"types": ["string"]
			}
		],
		supportedOptions: [],
		privateStateCallbacks: [],
		onSucceeded: processBinding
	});
	OSF.DDA.AsyncMethodCalls.define({
		method: OSF.DDA.AsyncMethodNames.ReleaseByIdAsync,
		requiredArguments: [
			{
				"name": Microsoft.Office.WebExtension.Parameters.Id,
				"types": ["string"]
			}
		],
		supportedOptions: [],
		privateStateCallbacks: [],
		onSucceeded: function (response, caller, callArgs) {
			var id=callArgs[Microsoft.Office.WebExtension.Parameters.Id];
			delete caller._eventDispatches[id];
		}
	});
	OSF.DDA.AsyncMethodCalls.define({
		method: OSF.DDA.AsyncMethodNames.GetDataAsync,
		requiredArguments: [],
		supportedOptions: [{
				name: Microsoft.Office.WebExtension.Parameters.CoercionType,
				value: {
					"enum": Microsoft.Office.WebExtension.CoercionType,
					"calculate": function (requiredArgs, binding) { return OSF.DDA.DataCoercion.getCoercionDefaultForBinding(binding.type); }
				}
			},
			{
				name: Microsoft.Office.WebExtension.Parameters.ValueFormat,
				value: {
					"enum": Microsoft.Office.WebExtension.ValueFormat,
					"defaultValue": Microsoft.Office.WebExtension.ValueFormat.Unformatted
				}
			},
			{
				name: Microsoft.Office.WebExtension.Parameters.FilterType,
				value: {
					"enum": Microsoft.Office.WebExtension.FilterType,
					"defaultValue": Microsoft.Office.WebExtension.FilterType.All
				}
			},
			{
				name: Microsoft.Office.WebExtension.Parameters.Rows,
				value: {
					"types": ["object", "string"],
					"defaultValue": null
				}
			},
			{
				name: Microsoft.Office.WebExtension.Parameters.Columns,
				value: {
					"types": ["object"],
					"defaultValue": null
				}
			},
			{
				name: Microsoft.Office.WebExtension.Parameters.StartRow,
				value: {
					"types": ["number"],
					"defaultValue": 0
				}
			},
			{
				name: Microsoft.Office.WebExtension.Parameters.StartColumn,
				value: {
					"types": ["number"],
					"defaultValue": 0
				}
			},
			{
				name: Microsoft.Office.WebExtension.Parameters.RowCount,
				value: {
					"types": ["number"],
					"defaultValue": 0
				}
			},
			{
				name: Microsoft.Office.WebExtension.Parameters.ColumnCount,
				value: {
					"types": ["number"],
					"defaultValue": 0
				}
			}
		],
		checkCallArgs: function (callArgs, caller, stateInfo) {
			if (callArgs[Microsoft.Office.WebExtension.Parameters.StartRow]==0 &&
				callArgs[Microsoft.Office.WebExtension.Parameters.StartColumn]==0 &&
				callArgs[Microsoft.Office.WebExtension.Parameters.RowCount]==0 &&
				callArgs[Microsoft.Office.WebExtension.Parameters.ColumnCount]==0) {
				delete callArgs[Microsoft.Office.WebExtension.Parameters.StartRow];
				delete callArgs[Microsoft.Office.WebExtension.Parameters.StartColumn];
				delete callArgs[Microsoft.Office.WebExtension.Parameters.RowCount];
				delete callArgs[Microsoft.Office.WebExtension.Parameters.ColumnCount];
			}
			if (callArgs[Microsoft.Office.WebExtension.Parameters.CoercionType] !=OSF.DDA.DataCoercion.getCoercionDefaultForBinding(caller.type) &&
				(callArgs[Microsoft.Office.WebExtension.Parameters.StartRow] ||
					callArgs[Microsoft.Office.WebExtension.Parameters.StartColumn] ||
					callArgs[Microsoft.Office.WebExtension.Parameters.RowCount] ||
					callArgs[Microsoft.Office.WebExtension.Parameters.ColumnCount])) {
				throw OSF.DDA.ErrorCodeManager.errorCodes.ooeCoercionTypeNotMatchBinding;
			}
			return callArgs;
		},
		privateStateCallbacks: [
			{
				name: Microsoft.Office.WebExtension.Parameters.Id,
				value: getObjectId
			}
		],
		onSucceeded: processData
	});
	OSF.DDA.AsyncMethodCalls.define({
		method: OSF.DDA.AsyncMethodNames.SetDataAsync,
		requiredArguments: [
			{
				"name": Microsoft.Office.WebExtension.Parameters.Data,
				"types": ["string", "object", "number", "boolean"]
			}
		],
		supportedOptions: [{
				name: Microsoft.Office.WebExtension.Parameters.CoercionType,
				value: {
					"enum": Microsoft.Office.WebExtension.CoercionType,
					"calculate": function (requiredArgs) { return OSF.DDA.DataCoercion.determineCoercionType(requiredArgs[Microsoft.Office.WebExtension.Parameters.Data]); }
				}
			},
			{
				name: Microsoft.Office.WebExtension.Parameters.Rows,
				value: {
					"types": ["object", "string"],
					"defaultValue": null
				}
			},
			{
				name: Microsoft.Office.WebExtension.Parameters.Columns,
				value: {
					"types": ["object"],
					"defaultValue": null
				}
			},
			{
				name: Microsoft.Office.WebExtension.Parameters.StartRow,
				value: {
					"types": ["number"],
					"defaultValue": 0
				}
			},
			{
				name: Microsoft.Office.WebExtension.Parameters.StartColumn,
				value: {
					"types": ["number"],
					"defaultValue": 0
				}
			}
		],
		checkCallArgs: function (callArgs, caller, stateInfo) {
			if (callArgs[Microsoft.Office.WebExtension.Parameters.StartRow]==0 &&
				callArgs[Microsoft.Office.WebExtension.Parameters.StartColumn]==0) {
				delete callArgs[Microsoft.Office.WebExtension.Parameters.StartRow];
				delete callArgs[Microsoft.Office.WebExtension.Parameters.StartColumn];
			}
			if (callArgs[Microsoft.Office.WebExtension.Parameters.CoercionType] !=OSF.DDA.DataCoercion.getCoercionDefaultForBinding(caller.type) &&
				(callArgs[Microsoft.Office.WebExtension.Parameters.StartRow] ||
					callArgs[Microsoft.Office.WebExtension.Parameters.StartColumn])) {
				throw OSF.DDA.ErrorCodeManager.errorCodes.ooeCoercionTypeNotMatchBinding;
			}
			return callArgs;
		},
		privateStateCallbacks: [
			{
				name: Microsoft.Office.WebExtension.Parameters.Id,
				value: getObjectId
			}
		]
	});
})();
OSF.OUtil.augmentList(OSF.DDA.BindingProperties, {
	RowCount: "BindingRowCount",
	ColumnCount: "BindingColumnCount",
	HasHeaders: "HasHeaders"
});
OSF.DDA.MatrixBinding=function OSF_DDA_MatrixBinding(id, docInstance, rows, cols) {
	OSF.DDA.MatrixBinding.uber.constructor.call(this, id, docInstance);
	OSF.OUtil.defineEnumerableProperties(this, {
		"type": {
			value: Microsoft.Office.WebExtension.BindingType.Matrix
		},
		"rowCount": {
			value: rows ? rows : 0
		},
		"columnCount": {
			value: cols ? cols : 0
		}
	});
};
OSF.OUtil.extend(OSF.DDA.MatrixBinding, OSF.DDA.Binding);
OSF.DDA.SafeArray.Delegate.ParameterMap.define({
	type: OSF.DDA.PropertyDescriptors.BindingProperties,
	fromHost: [
		{ name: OSF.DDA.BindingProperties.Id, value: 0 },
		{ name: OSF.DDA.BindingProperties.Type, value: 1 },
		{ name: OSF.DDA.SafeArray.UniqueArguments.BindingSpecificData, value: 2 }
	],
	isComplexType: true
});
OSF.DDA.SafeArray.Delegate.ParameterMap.define({
	type: Microsoft.Office.WebExtension.Parameters.BindingType,
	toHost: [
		{ name: Microsoft.Office.WebExtension.BindingType.Text, value: 0 },
		{ name: Microsoft.Office.WebExtension.BindingType.Matrix, value: 1 },
		{ name: Microsoft.Office.WebExtension.BindingType.Table, value: 2 }
	],
	invertible: true
});
OSF.DDA.SafeArray.Delegate.ParameterMap.define({
	type: OSF.DDA.MethodDispId.dispidAddBindingFromSelectionMethod,
	fromHost: [
		{ name: OSF.DDA.PropertyDescriptors.BindingProperties, value: OSF.DDA.SafeArray.Delegate.ParameterMap.self }
	],
	toHost: [
		{ name: Microsoft.Office.WebExtension.Parameters.Id, value: 0 },
		{ name: Microsoft.Office.WebExtension.Parameters.BindingType, value: 1 }
	]
});
OSF.DDA.SafeArray.Delegate.ParameterMap.define({
	type: OSF.DDA.MethodDispId.dispidAddBindingFromNamedItemMethod,
	fromHost: [
		{ name: OSF.DDA.PropertyDescriptors.BindingProperties, value: OSF.DDA.SafeArray.Delegate.ParameterMap.self }
	],
	toHost: [
		{ name: Microsoft.Office.WebExtension.Parameters.ItemName, value: 0 },
		{ name: Microsoft.Office.WebExtension.Parameters.Id, value: 1 },
		{ name: Microsoft.Office.WebExtension.Parameters.BindingType, value: 2 },
		{ name: Microsoft.Office.WebExtension.Parameters.FailOnCollision, value: 3 }
	]
});
OSF.DDA.SafeArray.Delegate.ParameterMap.define({
	type: OSF.DDA.MethodDispId.dispidReleaseBindingMethod,
	toHost: [
		{ name: Microsoft.Office.WebExtension.Parameters.Id, value: 0 }
	]
});
OSF.DDA.SafeArray.Delegate.ParameterMap.define({
	type: OSF.DDA.MethodDispId.dispidGetBindingMethod,
	fromHost: [
		{ name: OSF.DDA.PropertyDescriptors.BindingProperties, value: OSF.DDA.SafeArray.Delegate.ParameterMap.self }
	],
	toHost: [
		{ name: Microsoft.Office.WebExtension.Parameters.Id, value: 0 }
	]
});
OSF.DDA.SafeArray.Delegate.ParameterMap.define({
	type: OSF.DDA.MethodDispId.dispidGetAllBindingsMethod,
	fromHost: [
		{ name: OSF.DDA.ListDescriptors.BindingList, value: OSF.DDA.SafeArray.Delegate.ParameterMap.self }
	]
});
OSF.DDA.SafeArray.Delegate.ParameterMap.define({
	type: OSF.DDA.MethodDispId.dispidGetBindingDataMethod,
	fromHost: [
		{ name: Microsoft.Office.WebExtension.Parameters.Data, value: OSF.DDA.SafeArray.Delegate.ParameterMap.self }
	],
	toHost: [
		{ name: Microsoft.Office.WebExtension.Parameters.Id, value: 0 },
		{ name: Microsoft.Office.WebExtension.Parameters.CoercionType, value: 1 },
		{ name: Microsoft.Office.WebExtension.Parameters.ValueFormat, value: 2 },
		{ name: Microsoft.Office.WebExtension.Parameters.FilterType, value: 3 },
		{ name: OSF.DDA.PropertyDescriptors.Subset, value: 4 }
	]
});
OSF.DDA.SafeArray.Delegate.ParameterMap.define({
	type: OSF.DDA.MethodDispId.dispidSetBindingDataMethod,
	toHost: [
		{ name: Microsoft.Office.WebExtension.Parameters.Id, value: 0 },
		{ name: Microsoft.Office.WebExtension.Parameters.CoercionType, value: 1 },
		{ name: Microsoft.Office.WebExtension.Parameters.Data, value: 2 },
		{ name: OSF.DDA.SafeArray.UniqueArguments.Offset, value: 3 }
	]
});
OSF.DDA.SafeArray.Delegate.ParameterMap.define({
	type: OSF.DDA.SafeArray.UniqueArguments.BindingSpecificData,
	fromHost: [
		{ name: OSF.DDA.BindingProperties.RowCount, value: 0 },
		{ name: OSF.DDA.BindingProperties.ColumnCount, value: 1 },
		{ name: OSF.DDA.BindingProperties.HasHeaders, value: 2 }
	],
	isComplexType: true
});
OSF.DDA.SafeArray.Delegate.ParameterMap.define({
	type: OSF.DDA.PropertyDescriptors.Subset,
	toHost: [
		{ name: OSF.DDA.SafeArray.UniqueArguments.Offset, value: 0 },
		{ name: OSF.DDA.SafeArray.UniqueArguments.Run, value: 1 }
	],
	canonical: true,
	isComplexType: true
});
OSF.DDA.SafeArray.Delegate.ParameterMap.define({
	type: OSF.DDA.SafeArray.UniqueArguments.Offset,
	toHost: [
		{ name: Microsoft.Office.WebExtension.Parameters.StartRow, value: 0 },
		{ name: Microsoft.Office.WebExtension.Parameters.StartColumn, value: 1 }
	],
	canonical: true,
	isComplexType: true
});
OSF.DDA.SafeArray.Delegate.ParameterMap.define({
	type: OSF.DDA.SafeArray.UniqueArguments.Run,
	toHost: [
		{ name: Microsoft.Office.WebExtension.Parameters.RowCount, value: 0 },
		{ name: Microsoft.Office.WebExtension.Parameters.ColumnCount, value: 1 }
	],
	canonical: true,
	isComplexType: true
});
OSF.DDA.SafeArray.Delegate.ParameterMap.define({
	type: OSF.DDA.MethodDispId.dispidAddRowsMethod,
	toHost: [
		{ name: Microsoft.Office.WebExtension.Parameters.Id, value: 0 },
		{ name: Microsoft.Office.WebExtension.Parameters.Data, value: 1 }
	]
});
OSF.DDA.SafeArray.Delegate.ParameterMap.define({
	type: OSF.DDA.MethodDispId.dispidAddColumnsMethod,
	toHost: [
		{ name: Microsoft.Office.WebExtension.Parameters.Id, value: 0 },
		{ name: Microsoft.Office.WebExtension.Parameters.Data, value: 1 }
	]
});
OSF.DDA.SafeArray.Delegate.ParameterMap.define({
	type: OSF.DDA.MethodDispId.dispidClearAllRowsMethod,
	toHost: [
		{ name: Microsoft.Office.WebExtension.Parameters.Id, value: 0 }
	]
});
OSF.OUtil.augmentList(OSF.DDA.PropertyDescriptors, { TableDataProperties: "TableDataProperties" });
OSF.OUtil.augmentList(OSF.DDA.BindingProperties, {
	RowCount: "BindingRowCount",
	ColumnCount: "BindingColumnCount",
	HasHeaders: "HasHeaders"
});
OSF.DDA.TableDataProperties={
	TableRows: "TableRows",
	TableHeaders: "TableHeaders"
};
OSF.DDA.TableBinding=function OSF_DDA_TableBinding(id, docInstance, rows, cols, hasHeaders) {
	OSF.DDA.TableBinding.uber.constructor.call(this, id, docInstance);
	OSF.OUtil.defineEnumerableProperties(this, {
		"type": {
			value: Microsoft.Office.WebExtension.BindingType.Table
		},
		"rowCount": {
			value: rows ? rows : 0
		},
		"columnCount": {
			value: cols ? cols : 0
		},
		"hasHeaders": {
			value: hasHeaders ? hasHeaders : false
		}
	});
	var am=OSF.DDA.AsyncMethodNames;
	OSF.DDA.DispIdHost.addAsyncMethods(this, [
		am.AddRowsAsync,
		am.AddColumnsAsync,
		am.DeleteAllDataValuesAsync
	]);
};
OSF.OUtil.extend(OSF.DDA.TableBinding, OSF.DDA.Binding);
OSF.DDA.AsyncMethodNames.addNames({
	AddRowsAsync: "addRowsAsync",
	AddColumnsAsync: "addColumnsAsync",
	DeleteAllDataValuesAsync: "deleteAllDataValuesAsync"
});
(function () {
	function getObjectId(obj) { return obj.id; }
	OSF.DDA.AsyncMethodCalls.define({
		method: OSF.DDA.AsyncMethodNames.AddRowsAsync,
		requiredArguments: [
			{
				"name": Microsoft.Office.WebExtension.Parameters.Data,
				"types": ["object"]
			}
		],
		supportedOptions: [],
		privateStateCallbacks: [
			{
				name: Microsoft.Office.WebExtension.Parameters.Id,
				value: getObjectId
			}
		]
	});
	OSF.DDA.AsyncMethodCalls.define({
		method: OSF.DDA.AsyncMethodNames.AddColumnsAsync,
		requiredArguments: [
			{
				"name": Microsoft.Office.WebExtension.Parameters.Data,
				"types": ["object"]
			}
		],
		supportedOptions: [],
		privateStateCallbacks: [
			{
				name: Microsoft.Office.WebExtension.Parameters.Id,
				value: getObjectId
			}
		]
	});
	OSF.DDA.AsyncMethodCalls.define({
		method: OSF.DDA.AsyncMethodNames.DeleteAllDataValuesAsync,
		requiredArguments: [],
		supportedOptions: [],
		privateStateCallbacks: [
			{
				name: Microsoft.Office.WebExtension.Parameters.Id,
				value: getObjectId
			}
		]
	});
})();
OSF.DDA.TextBinding=function OSF_DDA_TextBinding(id, docInstance) {
	OSF.DDA.TextBinding.uber.constructor.call(this, id, docInstance);
	OSF.OUtil.defineEnumerableProperty(this, "type", {
		value: Microsoft.Office.WebExtension.BindingType.Text
	});
};
OSF.OUtil.extend(OSF.DDA.TextBinding, OSF.DDA.Binding);
OSF.DDA.AsyncMethodNames.addNames({ AddFromPromptAsync: "addFromPromptAsync" });
OSF.DDA.AsyncMethodCalls.define({
	method: OSF.DDA.AsyncMethodNames.AddFromPromptAsync,
	requiredArguments: [
		{
			"name": Microsoft.Office.WebExtension.Parameters.BindingType,
			"enum": Microsoft.Office.WebExtension.BindingType
		}
	],
	supportedOptions: [{
			name: Microsoft.Office.WebExtension.Parameters.Id,
			value: {
				"types": ["string"],
				"calculate": OSF.DDA.generateBindingId
			}
		},
		{
			name: Microsoft.Office.WebExtension.Parameters.PromptText,
			value: {
				"types": ["string"],
				"calculate": function () { return Strings.OfficeOM.L_AddBindingFromPromptDefaultText; }
			}
		},
		{
			name: Microsoft.Office.WebExtension.Parameters.SampleData,
			value: {
				"types": ["object"],
				"defaultValue": null
			}
		}
	],
	privateStateCallbacks: [],
	onSucceeded: function (bindingDescriptor) { return OSF.DDA.OMFactory.manufactureBinding(bindingDescriptor, Microsoft.Office.WebExtension.context.document); }
});
OSF.DDA.SafeArray.Delegate.ParameterMap.define({
	type: OSF.DDA.MethodDispId.dispidAddBindingFromPromptMethod,
	fromHost: [
		{ name: OSF.DDA.PropertyDescriptors.BindingProperties, value: OSF.DDA.SafeArray.Delegate.ParameterMap.self }
	],
	toHost: [
		{ name: Microsoft.Office.WebExtension.Parameters.Id, value: 0 },
		{ name: Microsoft.Office.WebExtension.Parameters.BindingType, value: 1 },
		{ name: Microsoft.Office.WebExtension.Parameters.PromptText, value: 2 }
	]
});
OSF.OUtil.augmentList(Microsoft.Office.WebExtension.EventType, { DocumentSelectionChanged: "documentSelectionChanged" });
OSF.DDA.DocumentSelectionChangedEventArgs=function OSF_DDA_DocumentSelectionChangedEventArgs(docInstance) {
	OSF.OUtil.defineEnumerableProperties(this, {
		"type": {
			value: Microsoft.Office.WebExtension.EventType.DocumentSelectionChanged
		},
		"document": {
			value: docInstance
		}
	});
};
OSF.OUtil.augmentList(Microsoft.Office.WebExtension.EventType, { ObjectDeleted: "objectDeleted" });
OSF.OUtil.augmentList(Microsoft.Office.WebExtension.EventType, { ObjectSelectionChanged: "objectSelectionChanged" });
OSF.OUtil.augmentList(Microsoft.Office.WebExtension.EventType, { ObjectDataChanged: "objectDataChanged" });
OSF.OUtil.augmentList(Microsoft.Office.WebExtension.EventType, { ContentControlAdded: "contentControlAdded" });
OSF.DDA.ObjectEventArgs=function OSF_DDA_ObjectEventArgs(eventType, object) {
	OSF.OUtil.defineEnumerableProperties(this, {
		"type": { value: eventType },
		"object": { value: object }
	});
};
OSF.DDA.SafeArray.Delegate.ParameterMap.define({ type: OSF.DDA.EventDispId.dispidDocumentSelectionChangedEvent });
OSF.DDA.SafeArray.Delegate.ParameterMap.define({
	type: OSF.DDA.EventDispId.dispidObjectDeletedEvent,
	toHost: [
		{ name: Microsoft.Office.WebExtension.Parameters.Id, value: 0 }
	],
	fromHost: [
		{ name: Microsoft.Office.WebExtension.Parameters.Id, value: OSF.DDA.SafeArray.Delegate.ParameterMap.sourceData }
	]
});
OSF.DDA.SafeArray.Delegate.ParameterMap.define({
	type: OSF.DDA.EventDispId.dispidObjectSelectionChangedEvent,
	toHost: [
		{ name: Microsoft.Office.WebExtension.Parameters.Id, value: 0 }
	],
	fromHost: [
		{ name: Microsoft.Office.WebExtension.Parameters.Id, value: OSF.DDA.SafeArray.Delegate.ParameterMap.sourceData }
	]
});
OSF.DDA.SafeArray.Delegate.ParameterMap.define({
	type: OSF.DDA.EventDispId.dispidObjectDataChangedEvent,
	toHost: [
		{ name: Microsoft.Office.WebExtension.Parameters.Id, value: 0 }
	],
	fromHost: [
		{ name: Microsoft.Office.WebExtension.Parameters.Id, value: OSF.DDA.SafeArray.Delegate.ParameterMap.sourceData }
	]
});
OSF.DDA.SafeArray.Delegate.ParameterMap.define({
	type: OSF.DDA.EventDispId.dispidContentControlAddedEvent,
	toHost: [
		{ name: Microsoft.Office.WebExtension.Parameters.Id, value: 0 }
	],
	fromHost: [
		{ name: Microsoft.Office.WebExtension.Parameters.Id, value: OSF.DDA.SafeArray.Delegate.ParameterMap.sourceData }
	]
});
OSF.OUtil.augmentList(Microsoft.Office.WebExtension.EventType, {
	BindingSelectionChanged: "bindingSelectionChanged",
	BindingDataChanged: "bindingDataChanged"
});
OSF.OUtil.augmentList(OSF.DDA.EventDescriptors, { BindingSelectionChangedEvent: "BindingSelectionChangedEvent" });
OSF.DDA.BindingSelectionChangedEventArgs=function OSF_DDA_BindingSelectionChangedEventArgs(bindingInstance, subset) {
	OSF.OUtil.defineEnumerableProperties(this, {
		"type": {
			value: Microsoft.Office.WebExtension.EventType.BindingSelectionChanged
		},
		"binding": {
			value: bindingInstance
		}
	});
	for (var prop in subset) {
		OSF.OUtil.defineEnumerableProperty(this, prop, {
			value: subset[prop]
		});
	}
};
OSF.DDA.BindingDataChangedEventArgs=function OSF_DDA_BindingDataChangedEventArgs(bindingInstance) {
	OSF.OUtil.defineEnumerableProperties(this, {
		"type": {
			value: Microsoft.Office.WebExtension.EventType.BindingDataChanged
		},
		"binding": {
			value: bindingInstance
		}
	});
};
OSF.DDA.SafeArray.Delegate.ParameterMap.define({
	type: OSF.DDA.EventDescriptors.BindingSelectionChangedEvent,
	fromHost: [
		{ name: OSF.DDA.PropertyDescriptors.BindingProperties, value: 0 },
		{ name: OSF.DDA.PropertyDescriptors.Subset, value: 1 }
	],
	isComplexType: true
});
OSF.DDA.SafeArray.Delegate.ParameterMap.define({
	type: OSF.DDA.EventDispId.dispidBindingSelectionChangedEvent,
	fromHost: [
		{ name: OSF.DDA.EventDescriptors.BindingSelectionChangedEvent, value: OSF.DDA.SafeArray.Delegate.ParameterMap.self }
	],
	isComplexType: true
});
OSF.DDA.SafeArray.Delegate.ParameterMap.define({
	type: OSF.DDA.EventDispId.dispidBindingDataChangedEvent,
	fromHost: [{ name: OSF.DDA.PropertyDescriptors.BindingProperties, value: OSF.DDA.SafeArray.Delegate.ParameterMap.self }]
});
OSF.OUtil.augmentList(Microsoft.Office.WebExtension.FilterType, { OnlyVisible: "onlyVisible" });
OSF.DDA.SafeArray.Delegate.ParameterMap.define({
	type: Microsoft.Office.WebExtension.Parameters.FilterType,
	toHost: [{ name: Microsoft.Office.WebExtension.FilterType.OnlyVisible, value: 1 }]
});
Microsoft.Office.WebExtension.GoToType={
	Binding: "binding",
	NamedItem: "namedItem",
	Slide: "slide",
	Index: "index"
};
Microsoft.Office.WebExtension.SelectionMode={
	Default: "default",
	Selected: "selected",
	None: "none"
};
Microsoft.Office.WebExtension.Index={
	First: "first",
	Last: "last",
	Next: "next",
	Previous: "previous"
};
OSF.DDA.AsyncMethodNames.addNames({ GoToByIdAsync: "goToByIdAsync" });
OSF.DDA.AsyncMethodCalls.define({
	method: OSF.DDA.AsyncMethodNames.GoToByIdAsync,
	requiredArguments: [{
			"name": Microsoft.Office.WebExtension.Parameters.Id,
			"types": ["string", "number"]
		},
		{
			"name": Microsoft.Office.WebExtension.Parameters.GoToType,
			"enum": Microsoft.Office.WebExtension.GoToType
		}
	],
	supportedOptions: [
		{
			name: Microsoft.Office.WebExtension.Parameters.SelectionMode,
			value: {
				"enum": Microsoft.Office.WebExtension.SelectionMode,
				"defaultValue": Microsoft.Office.WebExtension.SelectionMode.Default
			}
		}
	]
});
OSF.DDA.SafeArray.Delegate.ParameterMap.define({
	type: Microsoft.Office.WebExtension.Parameters.GoToType,
	toHost: [
		{ name: Microsoft.Office.WebExtension.GoToType.Binding, value: 0 },
		{ name: Microsoft.Office.WebExtension.GoToType.NamedItem, value: 1 },
		{ name: Microsoft.Office.WebExtension.GoToType.Slide, value: 2 },
		{ name: Microsoft.Office.WebExtension.GoToType.Index, value: 3 }
	]
});
OSF.DDA.SafeArray.Delegate.ParameterMap.define({
	type: Microsoft.Office.WebExtension.Parameters.SelectionMode,
	toHost: [
		{ name: Microsoft.Office.WebExtension.SelectionMode.Default, value: 0 },
		{ name: Microsoft.Office.WebExtension.SelectionMode.Selected, value: 1 },
		{ name: Microsoft.Office.WebExtension.SelectionMode.None, value: 2 }
	]
});
OSF.DDA.SafeArray.Delegate.ParameterMap.define({
	type: OSF.DDA.MethodDispId.dispidNavigateToMethod,
	toHost: [
		{ name: Microsoft.Office.WebExtension.Parameters.Id, value: 0 },
		{ name: Microsoft.Office.WebExtension.Parameters.GoToType, value: 1 },
		{ name: Microsoft.Office.WebExtension.Parameters.SelectionMode, value: 2 }
	]
});
OSF.DDA.AsyncMethodNames.addNames({
	ExecuteRichApiRequestAsync: "executeRichApiRequestAsync"
});
OSF.DDA.AsyncMethodCalls.define({
	method: OSF.DDA.AsyncMethodNames.ExecuteRichApiRequestAsync,
	requiredArguments: [
		{
			name: Microsoft.Office.WebExtension.Parameters.Data,
			types: ["object"]
		}
	],
	supportedOptions: []
});
OSF.OUtil.setNamespace("RichApi", OSF.DDA);
OSF.DDA.SafeArray.Delegate.ParameterMap.define({
	type: OSF.DDA.MethodDispId.dispidExecuteRichApiRequestMethod,
	toHost: [
		{ name: Microsoft.Office.WebExtension.Parameters.Data, value: 0 }
	],
	fromHost: [
		{ name: Microsoft.Office.WebExtension.Parameters.Data, value: OSF.DDA.SafeArray.Delegate.ParameterMap.self }
	]
});
OSF.OUtil.augmentList(Microsoft.Office.WebExtension.EventType, { RichApiMessage: "richApiMessage" });
OSF.DDA.RichApiMessageEventArgs=function OSF_DDA_RichApiMessageEventArgs(eventType, eventProperties) {
	var entryArray=eventProperties[Microsoft.Office.WebExtension.Parameters.Data];
	var entries=[];
	if (entryArray) {
		for (var i=0; i < entryArray.length; i++) {
			var elem=entryArray[i];
			if (elem.toArray) {
				elem=elem.toArray();
			}
			entries.push({
				messageCategory: elem[0],
				messageType: elem[1],
				targetId: elem[2],
				message: elem[3],
				id: elem[4],
				isRemoteOverride: elem[5]
			});
		}
	}
	OSF.OUtil.defineEnumerableProperties(this, {
		"type": { value: Microsoft.Office.WebExtension.EventType.RichApiMessage },
		"entries": { value: entries }
	});
};
var OfficeExt;
(function (OfficeExt) {
	var RichApiMessageManager=(function () {
		function RichApiMessageManager() {
			this._eventDispatch=null;
			this._eventDispatch=new OSF.EventDispatch([
				Microsoft.Office.WebExtension.EventType.RichApiMessage,
			]);
			OSF.DDA.DispIdHost.addEventSupport(this, this._eventDispatch);
		}
		return RichApiMessageManager;
	})();
	OfficeExt.RichApiMessageManager=RichApiMessageManager;
})(OfficeExt || (OfficeExt={}));
OSF.DDA.SafeArray.Delegate.ParameterMap.define({
	type: OSF.DDA.EventDispId.dispidRichApiMessageEvent,
	toHost: [
		{ name: Microsoft.Office.WebExtension.Parameters.Data, value: 0 }
	],
	fromHost: [
		{ name: Microsoft.Office.WebExtension.Parameters.Data, value: OSF.DDA.SafeArray.Delegate.ParameterMap.sourceData }
	]
});
Microsoft.Office.WebExtension.FileType={
	Text: "text",
	Compressed: "compressed",
	Pdf: "pdf"
};
OSF.OUtil.augmentList(OSF.DDA.PropertyDescriptors, {
	FileProperties: "FileProperties",
	FileSliceProperties: "FileSliceProperties"
});
OSF.DDA.FileProperties={
	Handle: "FileHandle",
	FileSize: "FileSize",
	SliceSize: Microsoft.Office.WebExtension.Parameters.SliceSize
};
OSF.DDA.File=function OSF_DDA_File(handle, fileSize, sliceSize) {
	OSF.OUtil.defineEnumerableProperties(this, {
		"size": {
			value: fileSize
		},
		"sliceCount": {
			value: Math.ceil(fileSize / sliceSize)
		}
	});
	var privateState={};
	privateState[OSF.DDA.FileProperties.Handle]=handle;
	privateState[OSF.DDA.FileProperties.SliceSize]=sliceSize;
	var am=OSF.DDA.AsyncMethodNames;
	OSF.DDA.DispIdHost.addAsyncMethods(this, [
		am.GetDocumentCopyChunkAsync,
		am.ReleaseDocumentCopyAsync
	], privateState);
};
OSF.DDA.FileSliceOffset="fileSliceoffset";
OSF.DDA.AsyncMethodNames.addNames({
	GetDocumentCopyAsync: "getFileAsync",
	GetDocumentCopyChunkAsync: "getSliceAsync",
	ReleaseDocumentCopyAsync: "closeAsync"
});
OSF.DDA.AsyncMethodCalls.define({
	method: OSF.DDA.AsyncMethodNames.GetDocumentCopyAsync,
	requiredArguments: [
		{
			"name": Microsoft.Office.WebExtension.Parameters.FileType,
			"enum": Microsoft.Office.WebExtension.FileType
		}
	],
	supportedOptions: [
		{
			name: Microsoft.Office.WebExtension.Parameters.SliceSize,
			value: {
				"types": ["number"],
				"defaultValue": 4 * 1024 * 1024
			}
		}
	],
	checkCallArgs: function (callArgs, caller, stateInfo) {
		var sliceSize=callArgs[Microsoft.Office.WebExtension.Parameters.SliceSize];
		if (sliceSize <=0 || sliceSize > (4 * 1024 * 1024)) {
			throw OSF.DDA.ErrorCodeManager.errorCodes.ooeInvalidSliceSize;
		}
		return callArgs;
	},
	onSucceeded: function (fileDescriptor, caller, callArgs) {
		return new OSF.DDA.File(fileDescriptor[OSF.DDA.FileProperties.Handle], fileDescriptor[OSF.DDA.FileProperties.FileSize], callArgs[Microsoft.Office.WebExtension.Parameters.SliceSize]);
	}
});
OSF.DDA.AsyncMethodCalls.define({
	method: OSF.DDA.AsyncMethodNames.GetDocumentCopyChunkAsync,
	requiredArguments: [
		{
			"name": Microsoft.Office.WebExtension.Parameters.SliceIndex,
			"types": ["number"]
		}
	],
	privateStateCallbacks: [
		{
			name: OSF.DDA.FileProperties.Handle,
			value: function (caller, stateInfo) { return stateInfo[OSF.DDA.FileProperties.Handle]; }
		},
		{
			name: OSF.DDA.FileProperties.SliceSize,
			value: function (caller, stateInfo) { return stateInfo[OSF.DDA.FileProperties.SliceSize]; }
		}
	],
	checkCallArgs: function (callArgs, caller, stateInfo) {
		var index=callArgs[Microsoft.Office.WebExtension.Parameters.SliceIndex];
		if (index < 0 || index >=caller.sliceCount) {
			throw OSF.DDA.ErrorCodeManager.errorCodes.ooeIndexOutOfRange;
		}
		callArgs[OSF.DDA.FileSliceOffset]=parseInt((index * stateInfo[OSF.DDA.FileProperties.SliceSize]).toString());
		return callArgs;
	},
	onSucceeded: function (sliceDescriptor, caller, callArgs) {
		var slice={};
		OSF.OUtil.defineEnumerableProperties(slice, {
			"data": {
				value: sliceDescriptor[Microsoft.Office.WebExtension.Parameters.Data]
			},
			"index": {
				value: callArgs[Microsoft.Office.WebExtension.Parameters.SliceIndex]
			},
			"size": {
				value: sliceDescriptor[OSF.DDA.FileProperties.SliceSize]
			}
		});
		return slice;
	}
});
OSF.DDA.AsyncMethodCalls.define({
	method: OSF.DDA.AsyncMethodNames.ReleaseDocumentCopyAsync,
	privateStateCallbacks: [
		{
			name: OSF.DDA.FileProperties.Handle,
			value: function (caller, stateInfo) { return stateInfo[OSF.DDA.FileProperties.Handle]; }
		}
	]
});
OSF.DDA.SafeArray.Delegate.ParameterMap.define({
	type: OSF.DDA.PropertyDescriptors.FileProperties,
	fromHost: [
		{ name: OSF.DDA.FileProperties.Handle, value: 0 },
		{ name: OSF.DDA.FileProperties.FileSize, value: 1 }
	],
	isComplexType: true
});
OSF.DDA.SafeArray.Delegate.ParameterMap.define({
	type: OSF.DDA.PropertyDescriptors.FileSliceProperties,
	fromHost: [
		{ name: Microsoft.Office.WebExtension.Parameters.Data, value: 0 },
		{ name: OSF.DDA.FileProperties.SliceSize, value: 1 }
	],
	isComplexType: true
});
OSF.DDA.SafeArray.Delegate.ParameterMap.define({
	type: Microsoft.Office.WebExtension.Parameters.FileType,
	toHost: [
		{ name: Microsoft.Office.WebExtension.FileType.Text, value: 0 },
		{ name: Microsoft.Office.WebExtension.FileType.Compressed, value: 5 },
		{ name: Microsoft.Office.WebExtension.FileType.Pdf, value: 6 }
	]
});
OSF.DDA.SafeArray.Delegate.ParameterMap.define({
	type: OSF.DDA.MethodDispId.dispidGetDocumentCopyMethod,
	toHost: [{ name: Microsoft.Office.WebExtension.Parameters.FileType, value: 0 }],
	fromHost: [
		{ name: OSF.DDA.PropertyDescriptors.FileProperties, value: OSF.DDA.SafeArray.Delegate.ParameterMap.self }
	]
});
OSF.DDA.SafeArray.Delegate.ParameterMap.define({
	type: OSF.DDA.MethodDispId.dispidGetDocumentCopyChunkMethod,
	toHost: [
		{ name: OSF.DDA.FileProperties.Handle, value: 0 },
		{ name: OSF.DDA.FileSliceOffset, value: 1 },
		{ name: OSF.DDA.FileProperties.SliceSize, value: 2 }
	],
	fromHost: [
		{ name: OSF.DDA.PropertyDescriptors.FileSliceProperties, value: OSF.DDA.SafeArray.Delegate.ParameterMap.self }
	]
});
OSF.DDA.SafeArray.Delegate.ParameterMap.define({
	type: OSF.DDA.MethodDispId.dispidReleaseDocumentCopyMethod,
	toHost: [{ name: OSF.DDA.FileProperties.Handle, value: 0 }]
});
OSF.DDA.FilePropertiesDescriptor={
	Url: "Url"
};
OSF.OUtil.augmentList(OSF.DDA.PropertyDescriptors, {
	FilePropertiesDescriptor: "FilePropertiesDescriptor"
});
Microsoft.Office.WebExtension.FileProperties=function Microsoft_Office_WebExtension_FileProperties(filePropertiesDescriptor) {
	OSF.OUtil.defineEnumerableProperties(this, {
		"url": {
			value: filePropertiesDescriptor[OSF.DDA.FilePropertiesDescriptor.Url]
		}
	});
};
OSF.DDA.AsyncMethodNames.addNames({ GetFilePropertiesAsync: "getFilePropertiesAsync" });
OSF.DDA.AsyncMethodCalls.define({
	method: OSF.DDA.AsyncMethodNames.GetFilePropertiesAsync,
	fromHost: [
		{ name: OSF.DDA.PropertyDescriptors.FilePropertiesDescriptor, value: 0 }
	],
	requiredArguments: [],
	supportedOptions: [],
	onSucceeded: function (filePropertiesDescriptor, caller, callArgs) {
		return new Microsoft.Office.WebExtension.FileProperties(filePropertiesDescriptor);
	}
});
OSF.DDA.SafeArray.Delegate.ParameterMap.define({
	type: OSF.DDA.PropertyDescriptors.FilePropertiesDescriptor,
	fromHost: [
		{ name: OSF.DDA.FilePropertiesDescriptor.Url, value: 0 }
	],
	isComplexType: true
});
OSF.DDA.SafeArray.Delegate.ParameterMap.define({
	type: OSF.DDA.MethodDispId.dispidGetFilePropertiesMethod,
	fromHost: [
		{ name: OSF.DDA.PropertyDescriptors.FilePropertiesDescriptor, value: OSF.DDA.SafeArray.Delegate.ParameterMap.self }
	]
});
OSF.DDA.ExcelTableBinding=function OSF_DDA_ExcelTableBinding(id, docInstance, rows, cols, hasHeaders) {
	var am=OSF.DDA.AsyncMethodNames;
	OSF.DDA.DispIdHost.addAsyncMethods(this, [
		am.ClearFormatsAsync,
		am.SetTableOptionsAsync,
		am.SetFormatsAsync
	]);
	OSF.DDA.ExcelTableBinding.uber.constructor.call(this, id, docInstance, rows, cols, hasHeaders);
	OSF.OUtil.finalizeProperties(this);
};
OSF.OUtil.extend(OSF.DDA.ExcelTableBinding, OSF.DDA.TableBinding);
(function () {
	OSF.DDA.AsyncMethodCalls.define({
		method: OSF.DDA.AsyncMethodNames.SetSelectedDataAsync,
		requiredArguments: [
			{
				"name": Microsoft.Office.WebExtension.Parameters.Data,
				"types": ["string", "object", "number", "boolean"]
			}
		],
		supportedOptions: [{
				name: Microsoft.Office.WebExtension.Parameters.CoercionType,
				value: {
					"enum": Microsoft.Office.WebExtension.CoercionType,
					"calculate": function (requiredArgs) { return OSF.DDA.DataCoercion.determineCoercionType(requiredArgs[Microsoft.Office.WebExtension.Parameters.Data]); }
				}
			},
			{
				name: Microsoft.Office.WebExtension.Parameters.CellFormat,
				value: {
					"types": ["object"],
					"defaultValue": []
				}
			},
			{
				name: Microsoft.Office.WebExtension.Parameters.TableOptions,
				value: {
					"types": ["object"],
					"defaultValue": []
				}
			}
		],
		privateStateCallbacks: []
	});
	OSF.DDA.AsyncMethodCalls.define({
		method: OSF.DDA.AsyncMethodNames.SetDataAsync,
		requiredArguments: [
			{
				"name": Microsoft.Office.WebExtension.Parameters.Data,
				"types": ["string", "object", "number", "boolean"]
			}
		],
		supportedOptions: [{
				name: Microsoft.Office.WebExtension.Parameters.CoercionType,
				value: {
					"enum": Microsoft.Office.WebExtension.CoercionType,
					"calculate": function (requiredArgs) { return OSF.DDA.DataCoercion.determineCoercionType(requiredArgs[Microsoft.Office.WebExtension.Parameters.Data]); }
				}
			},
			{
				name: Microsoft.Office.WebExtension.Parameters.Rows,
				value: {
					"types": ["object", "string"],
					"defaultValue": null
				}
			},
			{
				name: Microsoft.Office.WebExtension.Parameters.Columns,
				value: {
					"types": ["object"],
					"defaultValue": null
				}
			},
			{
				name: Microsoft.Office.WebExtension.Parameters.StartRow,
				value: {
					"types": ["number"],
					"defaultValue": 0
				}
			},
			{
				name: Microsoft.Office.WebExtension.Parameters.StartColumn,
				value: {
					"types": ["number"],
					"defaultValue": 0
				}
			},
			{
				name: Microsoft.Office.WebExtension.Parameters.CellFormat,
				value: {
					"types": ["object"],
					"defaultValue": []
				}
			},
			{
				name: Microsoft.Office.WebExtension.Parameters.TableOptions,
				value: {
					"types": ["object"],
					"defaultValue": []
				}
			}
		],
		checkCallArgs: function (callArgs, caller, stateInfo) {
			var Parameters=Microsoft.Office.WebExtension.Parameters;
			if (callArgs[Parameters.StartRow]==0 &&
				callArgs[Parameters.StartColumn]==0 &&
				OSF.OUtil.isArray(callArgs[Parameters.CellFormat]) && callArgs[Parameters.CellFormat].length===0 &&
				OSF.OUtil.isArray(callArgs[Parameters.TableOptions]) && callArgs[Parameters.TableOptions].length===0) {
				delete callArgs[Parameters.StartRow];
				delete callArgs[Parameters.StartColumn];
				delete callArgs[Parameters.CellFormat];
				delete callArgs[Parameters.TableOptions];
			}
			if (callArgs[Parameters.CoercionType] !=OSF.DDA.DataCoercion.getCoercionDefaultForBinding(caller.type) &&
				((callArgs[Parameters.StartRow] && callArgs[Parameters.StartRow] !=0) ||
					(callArgs[Parameters.StartColumn] && callArgs[Parameters.StartColumn] !=0) ||
					callArgs[Parameters.CellFormat] ||
					callArgs[Parameters.TableOptions])) {
				throw OSF.DDA.ErrorCodeManager.errorCodes.ooeCoercionTypeNotMatchBinding;
			}
			return callArgs;
		},
		privateStateCallbacks: [
			{
				name: Microsoft.Office.WebExtension.Parameters.Id,
				value: function (obj) { return obj.id; }
			}
		]
	});
	OSF.DDA.BindingPromise.prototype.setTableOptionsAsync=function OSF_DDA_BindingPromise$setTableOptionsAsync() {
		var args=arguments;
		this._fetch(function onComplete(binding) { binding.setTableOptionsAsync.apply(binding, args); });
		return this;
	},
		OSF.DDA.BindingPromise.prototype.setFormatsAsync=function OSF_DDA_BindingPromise$setFormatsAsync() {
			var args=arguments;
			this._fetch(function onComplete(binding) { binding.setFormatsAsync.apply(binding, args); });
			return this;
		},
		OSF.DDA.BindingPromise.prototype.clearFormatsAsync=function OSF_DDA_BindingPromise$clearFormatsAsync() {
			var args=arguments;
			this._fetch(function onComplete(binding) { binding.clearFormatsAsync.apply(binding, args); });
			return this;
		};
})();
(function () {
	function getObjectId(obj) { return obj.id; }
	OSF.DDA.AsyncMethodNames.addNames({
		ClearFormatsAsync: "clearFormatsAsync",
		SetTableOptionsAsync: "setTableOptionsAsync",
		SetFormatsAsync: "setFormatsAsync"
	});
	OSF.DDA.AsyncMethodCalls.define({
		method: OSF.DDA.AsyncMethodNames.ClearFormatsAsync,
		requiredArguments: [],
		supportedOptions: [],
		privateStateCallbacks: [
			{
				name: Microsoft.Office.WebExtension.Parameters.Id,
				value: getObjectId
			}
		]
	});
	OSF.DDA.AsyncMethodCalls.define({
		method: OSF.DDA.AsyncMethodNames.SetTableOptionsAsync,
		requiredArguments: [
			{
				"name": Microsoft.Office.WebExtension.Parameters.TableOptions,
				"defaultValue": []
			}
		],
		privateStateCallbacks: [
			{
				name: Microsoft.Office.WebExtension.Parameters.Id,
				value: getObjectId
			}
		]
	});
	OSF.DDA.AsyncMethodCalls.define({
		method: OSF.DDA.AsyncMethodNames.SetFormatsAsync,
		requiredArguments: [
			{
				"name": Microsoft.Office.WebExtension.Parameters.CellFormat,
				"defaultValue": []
			}
		],
		privateStateCallbacks: [
			{
				name: Microsoft.Office.WebExtension.Parameters.Id,
				value: getObjectId
			}
		]
	});
})();
Microsoft.Office.WebExtension.Table={
	All: 0,
	Data: 1,
	Headers: 2
};
(function () {
	OSF.DDA.SafeArray.Delegate.ParameterMap.define({
		type: OSF.DDA.MethodDispId.dispidClearFormatsMethod,
		toHost: [
			{ name: Microsoft.Office.WebExtension.Parameters.Id, value: 0 }
		]
	});
	OSF.DDA.SafeArray.Delegate.ParameterMap.define({
		type: OSF.DDA.MethodDispId.dispidSetTableOptionsMethod,
		toHost: [
			{ name: Microsoft.Office.WebExtension.Parameters.Id, value: 0 },
			{ name: Microsoft.Office.WebExtension.Parameters.TableOptions, value: 1 },
		]
	});
	OSF.DDA.SafeArray.Delegate.ParameterMap.define({
		type: OSF.DDA.MethodDispId.dispidSetFormatsMethod,
		toHost: [
			{ name: Microsoft.Office.WebExtension.Parameters.Id, value: 0 },
			{ name: Microsoft.Office.WebExtension.Parameters.CellFormat, value: 1 },
		]
	});
	OSF.DDA.SafeArray.Delegate.ParameterMap.define({
		type: OSF.DDA.MethodDispId.dispidSetSelectedDataMethod,
		toHost: [
			{ name: Microsoft.Office.WebExtension.Parameters.CoercionType, value: 0 },
			{ name: Microsoft.Office.WebExtension.Parameters.Data, value: 1 },
			{ name: Microsoft.Office.WebExtension.Parameters.CellFormat, value: 2 },
			{ name: Microsoft.Office.WebExtension.Parameters.TableOptions, value: 3 }
		]
	});
	OSF.DDA.SafeArray.Delegate.ParameterMap.define({
		type: OSF.DDA.MethodDispId.dispidSetBindingDataMethod,
		toHost: [
			{ name: Microsoft.Office.WebExtension.Parameters.Id, value: 0 },
			{ name: Microsoft.Office.WebExtension.Parameters.CoercionType, value: 1 },
			{ name: Microsoft.Office.WebExtension.Parameters.Data, value: 2 },
			{ name: OSF.DDA.SafeArray.UniqueArguments.Offset, value: 3 },
			{ name: Microsoft.Office.WebExtension.Parameters.CellFormat, value: 4 },
			{ name: Microsoft.Office.WebExtension.Parameters.TableOptions, value: 5 }
		]
	});
	var tableOptionProperties={
		headerRow: 0,
		bandedRows: 1,
		firstColumn: 2,
		lastColumn: 3,
		bandedColumns: 4,
		filterButton: 5,
		style: 6,
		totalRow: 7
	};
	var cellProperties={
		row: 0,
		column: 1
	};
	var formatProperties={
		alignHorizontal: { text: "alignHorizontal", type: 1 },
		alignVertical: { text: "alignVertical", type: 2 },
		backgroundColor: { text: "backgroundColor", type: 101 },
		borderStyle: { text: "borderStyle", type: 201 },
		borderColor: { text: "borderColor", type: 202 },
		borderTopStyle: { text: "borderTopStyle", type: 203 },
		borderTopColor: { text: "borderTopColor", type: 204 },
		borderBottomStyle: { text: "borderBottomStyle", type: 205 },
		borderBottomColor: { text: "borderBottomColor", type: 206 },
		borderLeftStyle: { text: "borderLeftStyle", type: 207 },
		borderLeftColor: { text: "borderLeftColor", type: 208 },
		borderRightStyle: { text: "borderRightStyle", type: 209 },
		borderRightColor: { text: "borderRightColor", type: 210 },
		borderOutlineStyle: { text: "borderOutlineStyle", type: 211 },
		borderOutlineColor: { text: "borderOutlineColor", type: 212 },
		borderInlineStyle: { text: "borderInlineStyle", type: 213 },
		borderInlineColor: { text: "borderInlineColor", type: 214 },
		fontFamily: { text: "fontFamily", type: 301 },
		fontStyle: { text: "fontStyle", type: 302 },
		fontSize: { text: "fontSize", type: 303 },
		fontUnderlineStyle: { text: "fontUnderlineStyle", type: 304 },
		fontColor: { text: "fontColor", type: 305 },
		fontDirection: { text: "fontDirection", type: 306 },
		fontStrikethrough: { text: "fontStrikethrough", type: 307 },
		fontSuperscript: { text: "fontSuperscript", type: 308 },
		fontSubscript: { text: "fontSubscript", type: 309 },
		fontNormal: { text: "fontNormal", type: 310 },
		indentLeft: { text: "indentLeft", type: 401 },
		indentRight: { text: "indentRight", type: 402 },
		numberFormat: { text: "numberFormat", type: 501 },
		width: { text: "width", type: 701 },
		height: { text: "height", type: 702 },
		wrapping: { text: "wrapping", type: 703 }
	};
	var borderStyleSet=[
		{ name: "none", value: 0 },
		{ name: "thin", value: 1 },
		{ name: "medium", value: 2 },
		{ name: "dashed", value: 3 },
		{ name: "dotted", value: 4 },
		{ name: "thick", value: 5 },
		{ name: "double", value: 6 },
		{ name: "hair", value: 7 },
		{ name: "medium dashed", value: 8 },
		{ name: "dash dot", value: 9 },
		{ name: "medium dash dot", value: 10 },
		{ name: "dash dot dot", value: 11 },
		{ name: "medium dash dot dot", value: 12 },
		{ name: "slant dash dot", value: 13 },
	];
	var colorSet=[
		{ name: "none", value: 0 },
		{ name: "black", value: 1 },
		{ name: "blue", value: 2 },
		{ name: "gray", value: 3 },
		{ name: "green", value: 4 },
		{ name: "orange", value: 5 },
		{ name: "pink", value: 6 },
		{ name: "purple", value: 7 },
		{ name: "red", value: 8 },
		{ name: "teal", value: 9 },
		{ name: "turquoise", value: 10 },
		{ name: "violet", value: 11 },
		{ name: "white", value: 12 },
		{ name: "yellow", value: 13 },
		{ name: "automatic", value: 14 },
	];
	var ns=OSF.DDA.SafeArray.Delegate.ParameterMap;
	ns.define({
		type: formatProperties.alignHorizontal.text,
		toHost: [
			{ name: "general", value: 0 },
			{ name: "left", value: 1 },
			{ name: "center", value: 2 },
			{ name: "right", value: 3 },
			{ name: "fill", value: 4 },
			{ name: "justify", value: 5 },
			{ name: "center across selection", value: 6 },
			{ name: "distributed", value: 7 },
		] });
	ns.define({
		type: formatProperties.alignVertical.text,
		toHost: [
			{ name: "top", value: 0 },
			{ name: "center", value: 1 },
			{ name: "bottom", value: 2 },
			{ name: "justify", value: 3 },
			{ name: "distributed", value: 4 },
		] });
	ns.define({
		type: formatProperties.backgroundColor.text,
		toHost: colorSet
	});
	ns.define({
		type: formatProperties.borderStyle.text,
		toHost: borderStyleSet
	});
	ns.define({
		type: formatProperties.borderColor.text,
		toHost: colorSet
	});
	ns.define({
		type: formatProperties.borderTopStyle.text,
		toHost: borderStyleSet
	});
	ns.define({
		type: formatProperties.borderTopColor.text,
		toHost: colorSet
	});
	ns.define({
		type: formatProperties.borderBottomStyle.text,
		toHost: borderStyleSet
	});
	ns.define({
		type: formatProperties.borderBottomColor.text,
		toHost: colorSet
	});
	ns.define({
		type: formatProperties.borderLeftStyle.text,
		toHost: borderStyleSet
	});
	ns.define({
		type: formatProperties.borderLeftColor.text,
		toHost: colorSet
	});
	ns.define({
		type: formatProperties.borderRightStyle.text,
		toHost: borderStyleSet
	});
	ns.define({
		type: formatProperties.borderRightColor.text,
		toHost: colorSet
	});
	ns.define({
		type: formatProperties.borderOutlineStyle.text,
		toHost: borderStyleSet
	});
	ns.define({
		type: formatProperties.borderOutlineColor.text,
		toHost: colorSet
	});
	ns.define({
		type: formatProperties.borderInlineStyle.text,
		toHost: borderStyleSet
	});
	ns.define({
		type: formatProperties.borderInlineColor.text,
		toHost: colorSet
	});
	ns.define({
		type: formatProperties.fontStyle.text,
		toHost: [
			{ name: "regular", value: 0 },
			{ name: "italic", value: 1 },
			{ name: "bold", value: 2 },
			{ name: "bold italic", value: 3 },
		] });
	ns.define({
		type: formatProperties.fontUnderlineStyle.text,
		toHost: [
			{ name: "none", value: 0 },
			{ name: "single", value: 1 },
			{ name: "double", value: 2 },
			{ name: "single accounting", value: 3 },
			{ name: "double accounting", value: 4 },
		] });
	ns.define({
		type: formatProperties.fontColor.text,
		toHost: colorSet
	});
	ns.define({
		type: formatProperties.fontDirection.text,
		toHost: [
			{ name: "context", value: 0 },
			{ name: "left-to-right", value: 1 },
			{ name: "right-to-left", value: 2 },
		] });
	ns.define({
		type: formatProperties.width.text,
		toHost: [
			{ name: "auto fit", value: -1 },
		] });
	ns.define({
		type: formatProperties.height.text,
		toHost: [
			{ name: "auto fit", value: -1 },
		] });
	ns.define({
		type: Microsoft.Office.WebExtension.Parameters.TableOptions,
		toHost: [
			{ name: "headerRow", value: 0 },
			{ name: "bandedRows", value: 1 },
			{ name: "firstColumn", value: 2 },
			{ name: "lastColumn", value: 3 },
			{ name: "bandedColumns", value: 4 },
			{ name: "filterButton", value: 5 },
			{ name: "style", value: 6 },
			{ name: "totalRow", value: 7 }
		] });
	ns.dynamicTypes[Microsoft.Office.WebExtension.Parameters.CellFormat]={
		toHost: function (data) {
			for (var entry in data) {
				if (data[entry].format) {
					data[entry].format=ns.doMapValues(data[entry].format, "toHost");
				}
			}
			return data;
		},
		fromHost: function (args) {
			return args;
		}
	};
	ns.setDynamicType(Microsoft.Office.WebExtension.Parameters.CellFormat, {
		toHost: function OSF_DDA_SafeArray_Delegate_SpecialProcessor_CellFormat$toHost(cellFormats) {
			var textCells="cells";
			var textFormat="format";
			var posCells=0;
			var posFormat=1;
			var ret=[];
			for (var index in cellFormats) {
				var cfOld=cellFormats[index];
				var cfNew=[];
				if (typeof (cfOld[textCells]) !=='undefined') {
					var cellsOld=cfOld[textCells];
					var cellsNew;
					if (typeof cfOld[textCells]==="object") {
						cellsNew=[];
						for (var entry in cellsOld) {
							if (typeof (cellProperties[entry]) !=='undefined') {
								cellsNew[cellProperties[entry]]=cellsOld[entry];
							}
						}
					}
					else {
						cellsNew=cellsOld;
					}
					cfNew[posCells]=cellsNew;
				}
				if (cfOld[textFormat]) {
					var formatOld=cfOld[textFormat];
					var formatNew=[];
					for (var entry2 in formatOld) {
						if (typeof (formatProperties[entry2]) !=='undefined') {
							formatNew.push([
								formatProperties[entry2].type,
								formatOld[entry2]
							]);
						}
					}
					cfNew[posFormat]=formatNew;
				}
				ret[index]=cfNew;
			}
			return ret;
		},
		fromHost: function OSF_DDA_SafeArray_Delegate_SpecialProcessor_CellFormat$fromHost(hostArgs) {
			return hostArgs;
		}
	});
	ns.setDynamicType(Microsoft.Office.WebExtension.Parameters.TableOptions, {
		toHost: function OSF_DDA_SafeArray_Delegate_SpecialProcessor_TableOptions$toHost(tableOptions) {
			var ret=[];
			for (var entry in tableOptions) {
				if (typeof (tableOptionProperties[entry]) !=='undefined') {
					ret[tableOptionProperties[entry]]=tableOptions[entry];
				}
			}
			return ret;
		},
		fromHost: function OSF_DDA_SafeArray_Delegate_SpecialProcessor_TableOptions$fromHost(hostArgs) {
			return hostArgs;
		}
	});
})();
var OfficeExt;
(function (OfficeExt) {
	var AppCommand;
	(function (AppCommand) {
		var AppCommandManager=(function () {
			function AppCommandManager() {
				var _this=this;
				this._pseudoDocument=null;
				this._eventDispatch=null;
				this._processAppCommandInvocation=function (args) {
					var verifyResult=_this._verifyManifestCallback(args.callbackName);
					if (verifyResult.errorCode !=OSF.DDA.ErrorCodeManager.errorCodes.ooeSuccess) {
						_this._invokeAppCommandCompletedMethod(args.appCommandId, verifyResult.errorCode, "");
						return;
					}
					var eventObj=_this._constructEventObjectForCallback(args);
					if (eventObj) {
						window.setTimeout(function () { verifyResult.callback(eventObj); }, 0);
					}
					else {
						_this._invokeAppCommandCompletedMethod(args.appCommandId, OSF.DDA.ErrorCodeManager.errorCodes.ooeInternalError, "");
					}
				};
			}
			AppCommandManager.initializeOsfDda=function () {
				OSF.DDA.AsyncMethodNames.addNames({
					AppCommandInvocationCompletedAsync: "appCommandInvocationCompletedAsync"
				});
				OSF.DDA.AsyncMethodCalls.define({
					method: OSF.DDA.AsyncMethodNames.AppCommandInvocationCompletedAsync,
					requiredArguments: [{
							"name": Microsoft.Office.WebExtension.Parameters.Id,
							"types": ["string"]
						},
						{
							"name": Microsoft.Office.WebExtension.Parameters.Status,
							"types": ["number"]
						},
						{
							"name": Microsoft.Office.WebExtension.Parameters.AppCommandInvocationCompletedData,
							"types": ["string"]
						}
					]
				});
				OSF.OUtil.augmentList(OSF.DDA.EventDescriptors, {
					AppCommandInvokedEvent: "AppCommandInvokedEvent"
				});
				OSF.OUtil.augmentList(Microsoft.Office.WebExtension.EventType, {
					AppCommandInvoked: "appCommandInvoked"
				});
				OSF.OUtil.setNamespace("AppCommand", OSF.DDA);
				OSF.DDA.AppCommand.AppCommandInvokedEventArgs=OfficeExt.AppCommand.AppCommandInvokedEventArgs;
			};
			AppCommandManager.prototype.initializeAndChangeOnce=function (callback) {
				AppCommand.registerDdaFacade();
				this._pseudoDocument={};
				OSF.DDA.DispIdHost.addAsyncMethods(this._pseudoDocument, [
					OSF.DDA.AsyncMethodNames.AppCommandInvocationCompletedAsync,
				]);
				this._eventDispatch=new OSF.EventDispatch([
					Microsoft.Office.WebExtension.EventType.AppCommandInvoked,
				]);
				var onRegisterCompleted=function (result) {
					if (callback) {
						if (result.status=="succeeded") {
							callback(OSF.DDA.ErrorCodeManager.errorCodes.ooeSuccess);
						}
						else {
							callback(OSF.DDA.ErrorCodeManager.errorCodes.ooeInternalError);
						}
					}
				};
				OSF.DDA.DispIdHost.addEventSupport(this._pseudoDocument, this._eventDispatch);
				this._pseudoDocument.addHandlerAsync(Microsoft.Office.WebExtension.EventType.AppCommandInvoked, this._processAppCommandInvocation, onRegisterCompleted);
			};
			AppCommandManager.prototype._verifyManifestCallback=function (callbackName) {
				var defaultResult={ callback: null, errorCode: OSF.DDA.ErrorCodeManager.errorCodes.ooeInvalidCallback };
				callbackName=callbackName.trim();
				try {
					var callList=callbackName.split(".");
					var parentObject=window;
					for (var i=0; i < callList.length - 1; i++) {
						if (parentObject[callList[i]] && (typeof parentObject[callList[i]]=="object" || typeof parentObject[callList[i]]=="function")) {
							parentObject=parentObject[callList[i]];
						}
						else {
							return defaultResult;
						}
					}
					var callbackFunc=parentObject[callList[callList.length - 1]];
					if (typeof callbackFunc !="function") {
						return defaultResult;
					}
				}
				catch (e) {
					return defaultResult;
				}
				return { callback: callbackFunc, errorCode: OSF.DDA.ErrorCodeManager.errorCodes.ooeSuccess };
			};
			AppCommandManager.prototype._invokeAppCommandCompletedMethod=function (appCommandId, resultCode, data) {
				this._pseudoDocument.appCommandInvocationCompletedAsync(appCommandId, resultCode, data);
			};
			AppCommandManager.prototype._constructEventObjectForCallback=function (args) {
				var _this=this;
				var eventObj=new AppCommandCallbackEventArgs();
				try {
					var jsonData=JSON.parse(args.eventObjStr);
					this._translateEventObjectInternal(jsonData, eventObj);
					Object.defineProperty(eventObj, 'completed', {
						value: function (completedContext) {
							eventObj.completedContext=completedContext;
							var jsonString=JSON.stringify(eventObj);
							_this._invokeAppCommandCompletedMethod(args.appCommandId, OSF.DDA.ErrorCodeManager.errorCodes.ooeSuccess, jsonString);
						},
						enumerable: true
					});
				}
				catch (e) {
					eventObj=null;
				}
				return eventObj;
			};
			AppCommandManager.prototype._translateEventObjectInternal=function (input, output) {
				for (var key in input) {
					if (!input.hasOwnProperty(key))
						continue;
					var inputChild=input[key];
					if (typeof inputChild=="object" && inputChild !=null) {
						OSF.OUtil.defineEnumerableProperty(output, key, {
							value: {}
						});
						this._translateEventObjectInternal(inputChild, output[key]);
					}
					else {
						Object.defineProperty(output, key, {
							value: inputChild,
							enumerable: true,
							writable: true
						});
					}
				}
			};
			AppCommandManager.prototype._constructObjectByTemplate=function (template, input) {
				var output={};
				if (!template || !input)
					return output;
				for (var key in template) {
					if (template.hasOwnProperty(key)) {
						output[key]=null;
						if (input[key] !=null) {
							var templateChild=template[key];
							var inputChild=input[key];
							var inputChildType=typeof inputChild;
							if (typeof templateChild=="object" && templateChild !=null) {
								output[key]=this._constructObjectByTemplate(templateChild, inputChild);
							}
							else if (inputChildType=="number" || inputChildType=="string" || inputChildType=="boolean") {
								output[key]=inputChild;
							}
						}
					}
				}
				return output;
			};
			AppCommandManager.instance=function () {
				if (AppCommandManager._instance==null) {
					AppCommandManager._instance=new AppCommandManager();
				}
				return AppCommandManager._instance;
			};
			AppCommandManager._instance=null;
			return AppCommandManager;
		})();
		AppCommand.AppCommandManager=AppCommandManager;
		var AppCommandInvokedEventArgs=(function () {
			function AppCommandInvokedEventArgs(appCommandId, callbackName, eventObjStr) {
				this.type=Microsoft.Office.WebExtension.EventType.AppCommandInvoked;
				this.appCommandId=appCommandId;
				this.callbackName=callbackName;
				this.eventObjStr=eventObjStr;
			}
			AppCommandInvokedEventArgs.create=function (eventProperties) {
				return new AppCommandInvokedEventArgs(eventProperties[AppCommand.AppCommandInvokedEventEnums.AppCommandId], eventProperties[AppCommand.AppCommandInvokedEventEnums.CallbackName], eventProperties[AppCommand.AppCommandInvokedEventEnums.EventObjStr]);
			};
			return AppCommandInvokedEventArgs;
		})();
		AppCommand.AppCommandInvokedEventArgs=AppCommandInvokedEventArgs;
		var AppCommandCallbackEventArgs=(function () {
			function AppCommandCallbackEventArgs() {
			}
			return AppCommandCallbackEventArgs;
		})();
		AppCommand.AppCommandCallbackEventArgs=AppCommandCallbackEventArgs;
		AppCommand.AppCommandInvokedEventEnums={
			AppCommandId: "appCommandId",
			CallbackName: "callbackName",
			EventObjStr: "eventObjStr"
		};
	})(AppCommand=OfficeExt.AppCommand || (OfficeExt.AppCommand={}));
})(OfficeExt || (OfficeExt={}));
OfficeExt.AppCommand.AppCommandManager.initializeOsfDda();
var OfficeExt;
(function (OfficeExt) {
	var AppCommand;
	(function (AppCommand) {
		function registerDdaFacade() {
			if (OSF.DDA.SafeArray) {
				var parameterMap=OSF.DDA.SafeArray.Delegate.ParameterMap;
				parameterMap.define({
					type: OSF.DDA.MethodDispId.dispidAppCommandInvocationCompletedMethod,
					toHost: [
						{ name: Microsoft.Office.WebExtension.Parameters.Id, value: 0 },
						{ name: Microsoft.Office.WebExtension.Parameters.Status, value: 1 },
						{ name: Microsoft.Office.WebExtension.Parameters.AppCommandInvocationCompletedData, value: 2 }
					]
				});
				parameterMap.define({
					type: OSF.DDA.EventDispId.dispidAppCommandInvokedEvent,
					fromHost: [
						{ name: OSF.DDA.EventDescriptors.AppCommandInvokedEvent, value: parameterMap.self }
					],
					isComplexType: true
				});
				parameterMap.define({
					type: OSF.DDA.EventDescriptors.AppCommandInvokedEvent,
					fromHost: [
						{ name: OfficeExt.AppCommand.AppCommandInvokedEventEnums.AppCommandId, value: 0 },
						{ name: OfficeExt.AppCommand.AppCommandInvokedEventEnums.CallbackName, value: 1 },
						{ name: OfficeExt.AppCommand.AppCommandInvokedEventEnums.EventObjStr, value: 2 },
					],
					isComplexType: true
				});
			}
		}
		AppCommand.registerDdaFacade=registerDdaFacade;
	})(AppCommand=OfficeExt.AppCommand || (OfficeExt.AppCommand={}));
})(OfficeExt || (OfficeExt={}));
OSF.OUtil.augmentList(Microsoft.Office.WebExtension.CoercionType, { Image: "image" });
OSF.DDA.SafeArray.Delegate.ParameterMap.define({
	type: Microsoft.Office.WebExtension.Parameters.CoercionType,
	toHost: [
		{ name: Microsoft.Office.WebExtension.CoercionType.Image, value: 8 }
	]
});
OSF.DDA.AsyncMethodNames.addNames({ GetAccessTokenAsync: "getAccessTokenAsync" });
OSF.DDA.Auth=function OSF_DDA_Auth() {
};
OSF.DDA.AsyncMethodCalls.define({
	method: OSF.DDA.AsyncMethodNames.GetAccessTokenAsync,
	requiredArguments: [],
	supportedOptions: [
		{
			name: Microsoft.Office.WebExtension.Parameters.ForceConsent,
			value: {
				"types": ["boolean"],
				"defaultValue": false
			}
		},
		{
			name: Microsoft.Office.WebExtension.Parameters.ForceAddAccount,
			value: {
				"types": ["boolean"],
				"defaultValue": false
			}
		},
		{
			name: Microsoft.Office.WebExtension.Parameters.AuthChallenge,
			value: {
				"types": ["string"],
				"defaultValue": ""
			}
		}
	],
	onSucceeded: function (dataDescriptor, caller, callArgs) {
		var data=dataDescriptor[Microsoft.Office.WebExtension.Parameters.Data];
		return data;
	}
});
OSF.DDA.SafeArray.Delegate.ParameterMap.define({
	type: OSF.DDA.MethodDispId.dispidGetAccessTokenMethod,
	toHost: [
		{ name: Microsoft.Office.WebExtension.Parameters.ForceConsent, value: 0 },
		{ name: Microsoft.Office.WebExtension.Parameters.ForceAddAccount, value: 1 },
		{ name: Microsoft.Office.WebExtension.Parameters.AuthChallenge, value: 2 }
	],
	fromHost: [
		{ name: Microsoft.Office.WebExtension.Parameters.Data, value: OSF.DDA.SafeArray.Delegate.ParameterMap.self }
	]
});
OSF.DDA.ExcelDocument=function OSF_DDA_ExcelDocument(officeAppContext, settings) {
	var bf=new OSF.DDA.BindingFacade(this);
	OSF.DDA.DispIdHost.addAsyncMethods(bf, [OSF.DDA.AsyncMethodNames.AddFromPromptAsync]);
	OSF.DDA.DispIdHost.addAsyncMethods(this, [OSF.DDA.AsyncMethodNames.GoToByIdAsync]);
	OSF.DDA.DispIdHost.addAsyncMethods(this, [OSF.DDA.AsyncMethodNames.GetDocumentCopyAsync]);
	OSF.DDA.DispIdHost.addAsyncMethods(this, [OSF.DDA.AsyncMethodNames.GetFilePropertiesAsync]);
	OSF.DDA.ExcelDocument.uber.constructor.call(this, officeAppContext, bf, settings);
	OSF.OUtil.finalizeProperties(this);
};
OSF.OUtil.extend(OSF.DDA.ExcelDocument, OSF.DDA.JsomDocument);
OSF.InitializationHelper.prototype.prepareRightAfterWebExtensionInitialize=function OSF_InitializationHelper$prepareRightAfterWebExtensionInitialize() {
	var appCommandHandler=OfficeExt.AppCommand.AppCommandManager.instance();
	appCommandHandler.initializeAndChangeOnce();
};
OSF.InitializationHelper.prototype.loadAppSpecificScriptAndCreateOM=function OSF_InitializationHelper$loadAppSpecificScriptAndCreateOM(appContext, appReady, basePath) {
	OSF.DDA.ErrorCodeManager.initializeErrorMessages(Strings.OfficeOM);
	appContext.doc=new OSF.DDA.ExcelDocument(appContext, this._initializeSettings(true));
	OSF.DDA.DispIdHost.addAsyncMethods(OSF.DDA.RichApi, [OSF.DDA.AsyncMethodNames.ExecuteRichApiRequestAsync]);
	OSF.DDA.RichApi.richApiMessageManager=new OfficeExt.RichApiMessageManager();
	appReady();
};
(function () {
	OSF.DDA.AsyncMethodCalls.define({
		method: OSF.DDA.AsyncMethodNames.SetSelectedDataAsync,
		requiredArguments: [
			{
				"name": Microsoft.Office.WebExtension.Parameters.Data,
				"types": ["string", "object", "number", "boolean"]
			}
		],
		supportedOptions: [{
				name: Microsoft.Office.WebExtension.Parameters.CoercionType,
				value: {
					"enum": Microsoft.Office.WebExtension.CoercionType,
					"calculate": function (requiredArgs) {
						return OSF.DDA.DataCoercion.determineCoercionType(requiredArgs[Microsoft.Office.WebExtension.Parameters.Data]);
					}
				}
			},
			{
				name: Microsoft.Office.WebExtension.Parameters.CellFormat,
				value: {
					"types": ["number", "object"],
					"defaultValue": []
				}
			},
			{
				name: Microsoft.Office.WebExtension.Parameters.TableOptions,
				value: {
					"types": ["number", "object"],
					"defaultValue": []
				}
			},
			{
				name: Microsoft.Office.WebExtension.Parameters.ImageWidth,
				value: {
					"types": ["number", "boolean"],
					"defaultValue": false
				}
			},
			{
				name: Microsoft.Office.WebExtension.Parameters.ImageHeight,
				value: {
					"types": ["number", "boolean"],
					"defaultValue": false
				}
			}
		],
		privateStateCallbacks: []
	});
})();

var __extends=(this && this.__extends) || function (d, b) {
	for (var p in b) if (b.hasOwnProperty(p)) d[p]=b[p];
	function __() { this.constructor=d; }
	d.prototype=b===null ? Object.create(b) : (__.prototype=b.prototype, new __());
};
var OfficeExtension;
(function (OfficeExtension) {
	var Action=(function () {
		function Action(actionInfo, isWriteOperation) {
			this.m_actionInfo=actionInfo;
			this.m_isWriteOperation=isWriteOperation;
		}
		Object.defineProperty(Action.prototype, "actionInfo", {
			get: function () {
				return this.m_actionInfo;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Action.prototype, "isWriteOperation", {
			get: function () {
				return this.m_isWriteOperation;
			},
			enumerable: true,
			configurable: true
		});
		return Action;
	}());
	OfficeExtension.Action=Action;
})(OfficeExtension || (OfficeExtension={}));
var OfficeExtension;
(function (OfficeExtension) {
	var TraceMarkerActionResultHandler=(function () {
		function TraceMarkerActionResultHandler(callback) {
			this.m_callback=callback;
		}
		TraceMarkerActionResultHandler.prototype._handleResult=function (value) {
			if (this.m_callback) {
				this.m_callback();
			}
		};
		return TraceMarkerActionResultHandler;
	}());
	var ActionFactory=(function () {
		function ActionFactory() {
		}
		ActionFactory.createSetPropertyAction=function (context, parent, propertyName, value) {
			OfficeExtension.Utility.validateObjectPath(parent);
			var actionInfo={
				Id: context._nextId(),
				ActionType: 4,
				Name: propertyName,
				ObjectPathId: parent._objectPath.objectPathInfo.Id,
				ArgumentInfo: {}
			};
			var args=[value];
			var referencedArgumentObjectPaths=OfficeExtension.Utility.setMethodArguments(context, actionInfo.ArgumentInfo, args);
			OfficeExtension.Utility.validateReferencedObjectPaths(referencedArgumentObjectPaths);
			var ret=new OfficeExtension.Action(actionInfo, true);
			context._pendingRequest.addAction(ret);
			context._pendingRequest.addReferencedObjectPath(parent._objectPath);
			context._pendingRequest.addReferencedObjectPaths(referencedArgumentObjectPaths);
			return ret;
		};
		ActionFactory.createMethodAction=function (context, parent, methodName, operationType, args) {
			OfficeExtension.Utility.validateObjectPath(parent);
			var actionInfo={
				Id: context._nextId(),
				ActionType: 3,
				Name: methodName,
				ObjectPathId: parent._objectPath.objectPathInfo.Id,
				ArgumentInfo: {}
			};
			var referencedArgumentObjectPaths=OfficeExtension.Utility.setMethodArguments(context, actionInfo.ArgumentInfo, args);
			OfficeExtension.Utility.validateReferencedObjectPaths(referencedArgumentObjectPaths);
			var isWriteOperation=operationType !=1;
			var ret=new OfficeExtension.Action(actionInfo, isWriteOperation);
			context._pendingRequest.addAction(ret);
			context._pendingRequest.addReferencedObjectPath(parent._objectPath);
			context._pendingRequest.addReferencedObjectPaths(referencedArgumentObjectPaths);
			return ret;
		};
		ActionFactory.createQueryAction=function (context, parent, queryOption) {
			OfficeExtension.Utility.validateObjectPath(parent);
			var actionInfo={
				Id: context._nextId(),
				ActionType: 2,
				Name: "",
				ObjectPathId: parent._objectPath.objectPathInfo.Id,
			};
			actionInfo.QueryInfo=queryOption;
			var ret=new OfficeExtension.Action(actionInfo, false);
			context._pendingRequest.addAction(ret);
			context._pendingRequest.addReferencedObjectPath(parent._objectPath);
			return ret;
		};
		ActionFactory.createRecursiveQueryAction=function (context, parent, query) {
			OfficeExtension.Utility.validateObjectPath(parent);
			var actionInfo={
				Id: context._nextId(),
				ActionType: 6,
				Name: "",
				ObjectPathId: parent._objectPath.objectPathInfo.Id,
				RecursiveQueryInfo: query
			};
			var ret=new OfficeExtension.Action(actionInfo, false);
			context._pendingRequest.addAction(ret);
			context._pendingRequest.addReferencedObjectPath(parent._objectPath);
			return ret;
		};
		ActionFactory.createInstantiateAction=function (context, obj) {
			OfficeExtension.Utility.validateObjectPath(obj);
			var actionInfo={
				Id: context._nextId(),
				ActionType: 1,
				Name: "",
				ObjectPathId: obj._objectPath.objectPathInfo.Id
			};
			var ret=new OfficeExtension.Action(actionInfo, false);
			context._pendingRequest.addAction(ret);
			context._pendingRequest.addReferencedObjectPath(obj._objectPath);
			context._pendingRequest.addActionResultHandler(ret, new OfficeExtension.InstantiateActionResultHandler(obj));
			return ret;
		};
		ActionFactory.createTraceAction=function (context, message, addTraceMessage) {
			var actionInfo={
				Id: context._nextId(),
				ActionType: 5,
				Name: "Trace",
				ObjectPathId: 0
			};
			var ret=new OfficeExtension.Action(actionInfo, false);
			context._pendingRequest.addAction(ret);
			if (addTraceMessage) {
				context._pendingRequest.addTrace(actionInfo.Id, message);
			}
			return ret;
		};
		ActionFactory.createTraceMarkerForCallback=function (context, callback) {
			var action=ActionFactory.createTraceAction(context, null, false);
			context._pendingRequest.addActionResultHandler(action, new TraceMarkerActionResultHandler(callback));
		};
		return ActionFactory;
	}());
	OfficeExtension.ActionFactory=ActionFactory;
})(OfficeExtension || (OfficeExtension={}));
var OfficeExtension;
(function (OfficeExtension) {
	var ClientObject=(function () {
		function ClientObject(context, objectPath) {
			OfficeExtension.Utility.checkArgumentNull(context, "context");
			this.m_context=context;
			this.m_objectPath=objectPath;
			if (this.m_objectPath) {
				if (!context._processingResult) {
					OfficeExtension.ActionFactory.createInstantiateAction(context, this);
					if ((context._autoCleanup) && (this._KeepReference)) {
						context.trackedObjects._autoAdd(this);
					}
				}
			}
		}
		Object.defineProperty(ClientObject.prototype, "context", {
			get: function () {
				return this.m_context;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ClientObject.prototype, "_objectPath", {
			get: function () {
				return this.m_objectPath;
			},
			set: function (value) {
				this.m_objectPath=value;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ClientObject.prototype, "isNull", {
			get: function () {
				OfficeExtension.Utility.throwIfNotLoaded("isNull", this._isNull, null, this._isNull);
				return this._isNull;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ClientObject.prototype, "isNullObject", {
			get: function () {
				OfficeExtension.Utility.throwIfNotLoaded("isNullObject", this._isNull, null, this._isNull);
				return this._isNull;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ClientObject.prototype, "_isNull", {
			get: function () {
				return this.m_isNull;
			},
			set: function (value) {
				this.m_isNull=value;
				if (value && this.m_objectPath) {
					this.m_objectPath._updateAsNullObject();
				}
			},
			enumerable: true,
			configurable: true
		});
		ClientObject.prototype._handleResult=function (value) {
			this._isNull=OfficeExtension.Utility.isNullOrUndefined(value);
			this.context.trackedObjects._autoTrackIfNecessaryWhenHandleObjectResultValue(this, value);
		};
		ClientObject.prototype._handleIdResult=function (value) {
			this._isNull=OfficeExtension.Utility.isNullOrUndefined(value);
			OfficeExtension.Utility.fixObjectPathIfNecessary(this, value);
			this.context.trackedObjects._autoTrackIfNecessaryWhenHandleObjectResultValue(this, value);
		};
		ClientObject.prototype._recursivelySet=function (input, options, scalarWriteablePropertyNames, objectPropertyNames, notAllowedToBeSetPropertyNames) {
			var isClientObject=(input instanceof ClientObject);
			if (isClientObject) {
				if (Object.getPrototypeOf(this)===Object.getPrototypeOf(input)) {
					input=JSON.parse(JSON.stringify(input));
				}
				else {
					throw OfficeExtension._Internal.RuntimeError._createInvalidArgError({
						argumentName: 'properties',
						errorLocation: this._className+".set"
					});
				}
			}
			try {
				var prop;
				for (var i=0; i < scalarWriteablePropertyNames.length; i++) {
					prop=scalarWriteablePropertyNames[i];
					if (input.hasOwnProperty(prop)) {
						this[prop]=input[prop];
					}
				}
				for (var i=0; i < objectPropertyNames.length; i++) {
					prop=objectPropertyNames[i];
					if (input.hasOwnProperty(prop)) {
						this[prop].set(input[prop], options);
					}
				}
				for (var i=0; i < notAllowedToBeSetPropertyNames.length; i++) {
					prop=notAllowedToBeSetPropertyNames[i];
					if (input.hasOwnProperty(prop)) {
						throw new OfficeExtension._Internal.RuntimeError({
							code: OfficeExtension.ErrorCodes.invalidArgument,
							message: OfficeExtension.Utility._getResourceString(OfficeExtension.ResourceStrings.cannotApplyPropertyThroughSetMethod, prop),
							debugInfo: {
								errorLocation: prop
							}
						});
					}
				}
				var throwOnReadOnly=!isClientObject;
				if (options && !OfficeExtension.Utility.isNullOrUndefined(throwOnReadOnly)) {
					throwOnReadOnly=options.throwOnReadOnly;
				}
				for (prop in input) {
					if (scalarWriteablePropertyNames.indexOf(prop) < 0 && objectPropertyNames.indexOf(prop) < 0) {
						var propertyDescriptor=Object.getOwnPropertyDescriptor(Object.getPrototypeOf(this), prop);
						if (!propertyDescriptor) {
							throw new OfficeExtension._Internal.RuntimeError({
								code: OfficeExtension.ErrorCodes.invalidArgument,
								message: OfficeExtension.Utility._getResourceString(OfficeExtension.ResourceStrings.propertyDoesNotExist, prop),
								debugInfo: {
									errorLocation: prop
								}
							});
						}
						if (throwOnReadOnly && !propertyDescriptor.set) {
							throw new OfficeExtension._Internal.RuntimeError({
								code: OfficeExtension.ErrorCodes.invalidArgument,
								message: OfficeExtension.Utility._getResourceString(OfficeExtension.ResourceStrings.attemptingToSetReadOnlyProperty, prop),
								debugInfo: {
									errorLocation: prop
								}
							});
						}
					}
				}
			}
			catch (innerError) {
				throw new OfficeExtension._Internal.RuntimeError({
					code: OfficeExtension.ErrorCodes.invalidArgument,
					message: OfficeExtension.Utility._getResourceString(OfficeExtension.ResourceStrings.invalidArgument, 'properties'),
					debugInfo: {
						errorLocation: this._className+".set"
					},
					innerError: innerError
				});
			}
		};
		return ClientObject;
	}());
	OfficeExtension.ClientObject=ClientObject;
})(OfficeExtension || (OfficeExtension={}));
var OfficeExtension;
(function (OfficeExtension) {
	var ClientRequest=(function () {
		function ClientRequest(context) {
			this.m_context=context;
			this.m_actions=[];
			this.m_actionResultHandler={};
			this.m_referencedObjectPaths={};
			this.m_flags=0;
			this.m_traceInfos={};
			this.m_pendingProcessEventHandlers=[];
			this.m_pendingEventHandlerActions={};
			this.m_responseTraceIds={};
			this.m_responseTraceMessages=[];
			this.m_preSyncPromises=[];
		}
		Object.defineProperty(ClientRequest.prototype, "flags", {
			get: function () {
				return this.m_flags;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ClientRequest.prototype, "traceInfos", {
			get: function () {
				return this.m_traceInfos;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ClientRequest.prototype, "_responseTraceMessages", {
			get: function () {
				return this.m_responseTraceMessages;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ClientRequest.prototype, "_responseTraceIds", {
			get: function () {
				return this.m_responseTraceIds;
			},
			enumerable: true,
			configurable: true
		});
		ClientRequest.prototype._setResponseTraceIds=function (value) {
			if (value) {
				for (var i=0; i < value.length; i++) {
					var traceId=value[i];
					this.m_responseTraceIds[traceId]=traceId;
					var message=this.m_traceInfos[traceId];
					if (!OfficeExtension.Utility.isNullOrUndefined(message)) {
						this.m_responseTraceMessages.push(message);
					}
				}
			}
		};
		ClientRequest.prototype.addAction=function (action) {
			if (action.isWriteOperation) {
				this.m_flags=this.m_flags | 1;
			}
			this.m_actions.push(action);
		};
		Object.defineProperty(ClientRequest.prototype, "hasActions", {
			get: function () {
				return this.m_actions.length > 0;
			},
			enumerable: true,
			configurable: true
		});
		ClientRequest.prototype.addTrace=function (actionId, message) {
			this.m_traceInfos[actionId]=message;
		};
		ClientRequest.prototype.addReferencedObjectPath=function (objectPath) {
			if (this.m_referencedObjectPaths[objectPath.objectPathInfo.Id]) {
				return;
			}
			if (!objectPath.isValid) {
				throw new OfficeExtension._Internal.RuntimeError({
					code: OfficeExtension.ErrorCodes.invalidObjectPath,
					message: OfficeExtension.Utility._getResourceString(OfficeExtension.ResourceStrings.invalidObjectPath, OfficeExtension.Utility.getObjectPathExpression(objectPath)),
					debugInfo: {
						errorLocation: OfficeExtension.Utility.getObjectPathExpression(objectPath)
					}
				});
			}
			while (objectPath) {
				if (objectPath.isWriteOperation) {
					this.m_flags=this.m_flags | 1;
				}
				this.m_referencedObjectPaths[objectPath.objectPathInfo.Id]=objectPath;
				if (objectPath.objectPathInfo.ObjectPathType==3) {
					this.addReferencedObjectPaths(objectPath.argumentObjectPaths);
				}
				objectPath=objectPath.parentObjectPath;
			}
		};
		ClientRequest.prototype.addReferencedObjectPaths=function (objectPaths) {
			if (objectPaths) {
				for (var i=0; i < objectPaths.length; i++) {
					this.addReferencedObjectPath(objectPaths[i]);
				}
			}
		};
		ClientRequest.prototype.addActionResultHandler=function (action, resultHandler) {
			this.m_actionResultHandler[action.actionInfo.Id]=resultHandler;
		};
		ClientRequest.prototype.buildRequestMessageBody=function () {
			var objectPaths={};
			for (var i in this.m_referencedObjectPaths) {
				objectPaths[i]=this.m_referencedObjectPaths[i].objectPathInfo;
			}
			var actions=[];
			for (var index=0; index < this.m_actions.length; index++) {
				actions.push(this.m_actions[index].actionInfo);
			}
			var ret={
				AutoKeepReference: this.m_context._autoCleanup,
				Actions: actions,
				ObjectPaths: objectPaths
			};
			return ret;
		};
		ClientRequest.prototype.processResponse=function (actionResults) {
			if (actionResults) {
				for (var i=0; i < actionResults.length; i++) {
					var actionResult=actionResults[i];
					var handler=this.m_actionResultHandler[actionResult.ActionId];
					if (handler) {
						handler._handleResult(actionResult.Value);
					}
				}
			}
		};
		ClientRequest.prototype.invalidatePendingInvalidObjectPaths=function () {
			for (var i in this.m_referencedObjectPaths) {
				if (this.m_referencedObjectPaths[i].isInvalidAfterRequest) {
					this.m_referencedObjectPaths[i].isValid=false;
				}
			}
		};
		ClientRequest.prototype._addPendingEventHandlerAction=function (eventHandlers, action) {
			if (!this.m_pendingEventHandlerActions[eventHandlers._id]) {
				this.m_pendingEventHandlerActions[eventHandlers._id]=[];
				this.m_pendingProcessEventHandlers.push(eventHandlers);
			}
			this.m_pendingEventHandlerActions[eventHandlers._id].push(action);
		};
		Object.defineProperty(ClientRequest.prototype, "_pendingProcessEventHandlers", {
			get: function () {
				return this.m_pendingProcessEventHandlers;
			},
			enumerable: true,
			configurable: true
		});
		ClientRequest.prototype._getPendingEventHandlerActions=function (eventHandlers) {
			return this.m_pendingEventHandlerActions[eventHandlers._id];
		};
		ClientRequest.prototype._addPreSyncPromise=function (value) {
			this.m_preSyncPromises.push(value);
		};
		Object.defineProperty(ClientRequest.prototype, "_preSyncPromises", {
			get: function () {
				return this.m_preSyncPromises;
			},
			enumerable: true,
			configurable: true
		});
		return ClientRequest;
	}());
	OfficeExtension.ClientRequest=ClientRequest;
})(OfficeExtension || (OfficeExtension={}));
var OfficeExtension;
(function (OfficeExtension) {
	var SessionBase=(function () {
		function SessionBase() {
		}
		SessionBase.prototype._resolveRequestUrlAndHeaderInfo=function () {
			return OfficeExtension.Utility._createPromiseFromResult(null);
		};
		SessionBase.prototype._createRequestExecutorOrNull=function () {
			return null;
		};
		Object.defineProperty(SessionBase.prototype, "eventRegistration", {
			get: function () {
				return OfficeExtension._Internal.officeJsEventRegistration;
			},
			enumerable: true,
			configurable: true
		});
		return SessionBase;
	}());
	OfficeExtension.SessionBase=SessionBase;
	var ClientRequestContext=(function () {
		function ClientRequestContext(url) {
			this.m_customRequestHeaders={};
			this._onRunFinishedNotifiers=[];
			this.m_nextId=0;
			if (ClientRequestContext._overrideSession) {
				this.m_requestUrlAndHeaderInfoResolver=ClientRequestContext._overrideSession;
			}
			else {
				if (OfficeExtension.Utility.isNullOrUndefined(url) || typeof (url)==="string" && url.length===0) {
					url=ClientRequestContext.defaultRequestUrlAndHeaders;
					if (!url) {
						url={ url: OfficeExtension.Constants.localDocument, headers: {} };
					}
				}
				if (typeof (url)==="string") {
					this.m_requestUrlAndHeaderInfo={ url: url, headers: {} };
				}
				else if (ClientRequestContext.isRequestUrlAndHeaderInfoResolver(url)) {
					this.m_requestUrlAndHeaderInfoResolver=url;
				}
				else if (ClientRequestContext.isRequestUrlAndHeaderInfo(url)) {
					var requestInfo=url;
					this.m_requestUrlAndHeaderInfo={ url: requestInfo.url, headers: {} };
					OfficeExtension.Utility._copyHeaders(requestInfo.headers, this.m_requestUrlAndHeaderInfo.headers);
				}
				else {
					throw OfficeExtension._Internal.RuntimeError._createInvalidArgError("url");
				}
			}
			if (this.m_requestUrlAndHeaderInfoResolver instanceof SessionBase) {
				this.m_session=this.m_requestUrlAndHeaderInfoResolver;
			}
			this._processingResult=false;
			this._customData=OfficeExtension.Constants.iterativeExecutor;
			this.sync=this.sync.bind(this);
		}
		Object.defineProperty(ClientRequestContext.prototype, "session", {
			get: function () {
				return this.m_session;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ClientRequestContext.prototype, "eventRegistration", {
			get: function () {
				if (this.m_session) {
					return this.m_session.eventRegistration;
				}
				return OfficeExtension._Internal.officeJsEventRegistration;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ClientRequestContext.prototype, "_url", {
			get: function () {
				if (this.m_requestUrlAndHeaderInfo) {
					return this.m_requestUrlAndHeaderInfo.url;
				}
				return null;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ClientRequestContext.prototype, "_pendingRequest", {
			get: function () {
				if (this.m_pendingRequest==null) {
					this.m_pendingRequest=new OfficeExtension.ClientRequest(this);
				}
				return this.m_pendingRequest;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ClientRequestContext.prototype, "trackedObjects", {
			get: function () {
				if (!this.m_trackedObjects) {
					this.m_trackedObjects=new OfficeExtension.TrackedObjects(this);
				}
				return this.m_trackedObjects;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ClientRequestContext.prototype, "requestHeaders", {
			get: function () {
				return this.m_customRequestHeaders;
			},
			enumerable: true,
			configurable: true
		});
		ClientRequestContext.prototype.load=function (clientObj, option) {
			OfficeExtension.Utility.validateContext(this, clientObj);
			var queryOption=ClientRequestContext.parseQueryOption(option);
			var action=OfficeExtension.ActionFactory.createQueryAction(this, clientObj, queryOption);
			this._pendingRequest.addActionResultHandler(action, clientObj);
		};
		ClientRequestContext.parseQueryOption=function (option) {
			var queryOption={};
			if (typeof (option)=="string") {
				var select=option;
				queryOption.Select=OfficeExtension.Utility._parseSelectExpand(select);
			}
			else if (Array.isArray(option)) {
				queryOption.Select=option;
			}
			else if (typeof (option)=="object") {
				var loadOption=option;
				if (typeof (loadOption.select)=="string") {
					queryOption.Select=OfficeExtension.Utility._parseSelectExpand(loadOption.select);
				}
				else if (Array.isArray(loadOption.select)) {
					queryOption.Select=loadOption.select;
				}
				else if (!OfficeExtension.Utility.isNullOrUndefined(loadOption.select)) {
					OfficeExtension.Utility.throwError(OfficeExtension.ResourceStrings.invalidArgument, "option.select");
				}
				if (typeof (loadOption.expand)=="string") {
					queryOption.Expand=OfficeExtension.Utility._parseSelectExpand(loadOption.expand);
				}
				else if (Array.isArray(loadOption.expand)) {
					queryOption.Expand=loadOption.expand;
				}
				else if (!OfficeExtension.Utility.isNullOrUndefined(loadOption.expand)) {
					OfficeExtension.Utility.throwError(OfficeExtension.ResourceStrings.invalidArgument, "option.expand");
				}
				if (typeof (loadOption.top)=="number") {
					queryOption.Top=loadOption.top;
				}
				else if (!OfficeExtension.Utility.isNullOrUndefined(loadOption.top)) {
					OfficeExtension.Utility.throwError(OfficeExtension.ResourceStrings.invalidArgument, "option.top");
				}
				if (typeof (loadOption.skip)=="number") {
					queryOption.Skip=loadOption.skip;
				}
				else if (!OfficeExtension.Utility.isNullOrUndefined(loadOption.skip)) {
					OfficeExtension.Utility.throwError(OfficeExtension.ResourceStrings.invalidArgument, "option.skip");
				}
			}
			else if (!OfficeExtension.Utility.isNullOrUndefined(option)) {
				OfficeExtension.Utility.throwError(OfficeExtension.ResourceStrings.invalidArgument, "option");
			}
			return queryOption;
		};
		ClientRequestContext.prototype.loadRecursive=function (clientObj, options, maxDepth) {
			if (!OfficeExtension.Utility.isPlainJsonObject(options)) {
				throw OfficeExtension._Internal.RuntimeError._createInvalidArgError("options");
			}
			var quries={};
			for (var key in options) {
				quries[key]=ClientRequestContext.parseQueryOption(options[key]);
			}
			var action=OfficeExtension.ActionFactory.createRecursiveQueryAction(this, clientObj, { Queries: quries, MaxDepth: maxDepth });
			this._pendingRequest.addActionResultHandler(action, clientObj);
		};
		ClientRequestContext.prototype.trace=function (message) {
			OfficeExtension.ActionFactory.createTraceAction(this, message, true);
		};
		ClientRequestContext.prototype._processOfficeJsErrorResponse=function (officeJsErrorCode, response) {
		};
		ClientRequestContext.prototype.syncPrivateMain=function () {
			var _this=this;
			return OfficeExtension.Utility._createPromiseFromResult(null)
				.then(function () {
				if (!_this.m_requestUrlAndHeaderInfo) {
					return _this.m_requestUrlAndHeaderInfoResolver._resolveRequestUrlAndHeaderInfo()
						.then(function (value) {
						_this.m_requestUrlAndHeaderInfo=value;
						if (!_this.m_requestUrlAndHeaderInfo) {
							_this.m_requestUrlAndHeaderInfo={ url: OfficeExtension.Constants.localDocument, headers: {} };
						}
						if (OfficeExtension.Utility.isNullOrEmptyString(_this.m_requestUrlAndHeaderInfo.url)) {
							_this.m_requestUrlAndHeaderInfo.url=OfficeExtension.Constants.localDocument;
						}
						if (!_this.m_requestUrlAndHeaderInfo.headers) {
							_this.m_requestUrlAndHeaderInfo.headers={};
						}
						if (typeof (_this.m_requestUrlAndHeaderInfoResolver._createRequestExecutorOrNull)==="function") {
							var executor=_this.m_requestUrlAndHeaderInfoResolver._createRequestExecutorOrNull();
							if (executor) {
								_this._requestExecutor=executor;
							}
						}
					});
				}
			})
				.then(function () {
				var req=_this._pendingRequest;
				_this.m_pendingRequest=null;
				return _this.processPreSyncPromises(req)
					.then(function () { return _this.syncPrivate(req); });
			});
		};
		ClientRequestContext.prototype.syncPrivate=function (req) {
			var _this=this;
			if (!req.hasActions) {
				return this.processPendingEventHandlers(req);
			}
			var msgBody=req.buildRequestMessageBody();
			var requestFlags=req.flags;
			if (!this._requestExecutor) {
				if (OfficeExtension.Utility._isLocalDocumentUrl(this.m_requestUrlAndHeaderInfo.url)) {
					this._requestExecutor=new OfficeExtension.OfficeJsRequestExecutor(this);
				}
				else {
					this._requestExecutor=new OfficeExtension.HttpRequestExecutor();
				}
			}
			var requestExecutor=this._requestExecutor;
			var headers={};
			OfficeExtension.Utility._copyHeaders(this.m_requestUrlAndHeaderInfo.headers, headers);
			OfficeExtension.Utility._copyHeaders(this.m_customRequestHeaders, headers);
			var requestExecutorRequestMessage={
				Url: this.m_requestUrlAndHeaderInfo.url,
				Headers: headers,
				Body: msgBody
			};
			req.invalidatePendingInvalidObjectPaths();
			var errorFromResponse=null;
			var errorFromProcessEventHandlers=null;
			return requestExecutor.executeAsync(this._customData, requestFlags, requestExecutorRequestMessage)
				.then(function (response) {
				errorFromResponse=_this.processRequestExecutorResponseMessage(req, response);
				return _this.processPendingEventHandlers(req)
					.catch(function (ex) {
					OfficeExtension.Utility.log("Error in processPendingEventHandlers");
					OfficeExtension.Utility.log(JSON.stringify(ex));
					errorFromProcessEventHandlers=ex;
				});
			})
				.then(function () {
				if (errorFromResponse) {
					OfficeExtension.Utility.log("Throw error from response: "+JSON.stringify(errorFromResponse));
					throw errorFromResponse;
				}
				if (errorFromProcessEventHandlers) {
					OfficeExtension.Utility.log("Throw error from ProcessEventHandler: "+JSON.stringify(errorFromProcessEventHandlers));
					var transformedError=null;
					if (errorFromProcessEventHandlers instanceof OfficeExtension._Internal.RuntimeError) {
						transformedError=errorFromProcessEventHandlers;
						transformedError.traceMessages=req._responseTraceMessages;
					}
					else {
						var message=null;
						if (typeof (errorFromProcessEventHandlers)==="string") {
							message=errorFromProcessEventHandlers;
						}
						else {
							message=errorFromProcessEventHandlers.message;
						}
						if (OfficeExtension.Utility.isNullOrEmptyString(message)) {
							message=OfficeExtension.Utility._getResourceString(OfficeExtension.ResourceStrings.cannotRegisterEvent);
						}
						transformedError=new OfficeExtension._Internal.RuntimeError({
							code: OfficeExtension.ErrorCodes.cannotRegisterEvent,
							message: message,
							traceMessages: req._responseTraceMessages
						});
					}
					throw transformedError;
				}
			});
		};
		ClientRequestContext.prototype.processRequestExecutorResponseMessage=function (req, response) {
			if (response.Body && response.Body.TraceIds) {
				req._setResponseTraceIds(response.Body.TraceIds);
			}
			var traceMessages=req._responseTraceMessages;
			if (response.Body) {
				var actionResults=null;
				if (response.Body.Results) {
					actionResults=response.Body.Results;
				}
				else if (response.Body.ProcessedResults && response.Body.ProcessedResults.Results) {
					actionResults=response.Body.ProcessedResults.Results;
				}
				if (actionResults) {
					this._processingResult=true;
					try {
						req.processResponse(actionResults);
					}
					finally {
						this._processingResult=false;
					}
				}
			}
			if (!OfficeExtension.Utility.isNullOrEmptyString(response.ErrorCode)) {
				return new OfficeExtension._Internal.RuntimeError({
					code: response.ErrorCode,
					message: response.ErrorMessage,
					traceMessages: traceMessages
				});
			}
			else if (response.Body && response.Body.Error) {
				return new OfficeExtension._Internal.RuntimeError({
					code: response.Body.Error.Code,
					message: response.Body.Error.Message,
					traceMessages: traceMessages,
					debugInfo: {
						errorLocation: response.Body.Error.Location
					}
				});
			}
			return null;
		};
		ClientRequestContext.prototype.processPendingEventHandlers=function (req) {
			var ret=OfficeExtension.Utility._createPromiseFromResult(null);
			for (var i=0; i < req._pendingProcessEventHandlers.length; i++) {
				var eventHandlers=req._pendingProcessEventHandlers[i];
				ret=ret.then(this.createProcessOneEventHandlersFunc(eventHandlers, req));
			}
			return ret;
		};
		ClientRequestContext.prototype.createProcessOneEventHandlersFunc=function (eventHandlers, req) {
			return function () { return eventHandlers._processRegistration(req); };
		};
		ClientRequestContext.prototype.processPreSyncPromises=function (req) {
			var ret=OfficeExtension.Utility._createPromiseFromResult(null);
			for (var i=0; i < req._preSyncPromises.length; i++) {
				var p=req._preSyncPromises[i];
				ret=ret.then(this.createProcessOneProSyncFunc(p));
			}
			return ret;
		};
		ClientRequestContext.prototype.createProcessOneProSyncFunc=function (p) {
			return function () { return p; };
		};
		ClientRequestContext.prototype.sync=function (passThroughValue) {
			return this.syncPrivateMain().then(function () { return passThroughValue; });
		};
		ClientRequestContext._run=function (ctxInitializer, batch, numCleanupAttempts, retryDelay, onCleanupSuccess, onCleanupFailure) {
			if (numCleanupAttempts===void 0) { numCleanupAttempts=3; }
			if (retryDelay===void 0) { retryDelay=5000; }
			return ClientRequestContext._runCommon("run", null, ctxInitializer, batch, numCleanupAttempts, retryDelay, onCleanupSuccess, onCleanupFailure);
		};
		ClientRequestContext.isRequestUrlAndHeaderInfo=function (value) {
			return (typeof (value)==="object" &&
				value !==null &&
				Object.getPrototypeOf(value)===Object.getPrototypeOf({}) &&
				!OfficeExtension.Utility.isNullOrUndefined(value.url));
		};
		ClientRequestContext.isRequestUrlAndHeaderInfoResolver=function (value) {
			return (typeof (value)==="object" &&
				value !==null &&
				typeof (value._resolveRequestUrlAndHeaderInfo)==="function");
		};
		ClientRequestContext._runBatch=function (functionName, receivedRunArgs, ctxInitializer, numCleanupAttempts, retryDelay, onCleanupSuccess, onCleanupFailure) {
			if (numCleanupAttempts===void 0) { numCleanupAttempts=3; }
			if (retryDelay===void 0) { retryDelay=5000; }
			var ctxRetriever;
			var batch;
			var requestInfo=null;
			var argOffset=0;
			if (receivedRunArgs.length > 0 &&
				(typeof (receivedRunArgs[0])==="string" ||
					ClientRequestContext.isRequestUrlAndHeaderInfo(receivedRunArgs[0]) ||
					ClientRequestContext.isRequestUrlAndHeaderInfoResolver(receivedRunArgs[0]))) {
				requestInfo=receivedRunArgs[0];
				argOffset=1;
			}
			if (receivedRunArgs.length==argOffset+1) {
				ctxRetriever=ctxInitializer;
				batch=receivedRunArgs[argOffset+0];
			}
			else if (receivedRunArgs.length==argOffset+2) {
				if (receivedRunArgs[argOffset+0] instanceof OfficeExtension.ClientObject) {
					ctxRetriever=function () { return receivedRunArgs[argOffset+0].context; };
				}
				else if (Array.isArray(receivedRunArgs[argOffset+0])) {
					var array=receivedRunArgs[argOffset+0];
					if (array.length==0) {
						return ClientRequestContext.createErrorPromise(functionName);
					}
					for (var i=0; i < array.length; i++) {
						if (!(array[i] instanceof OfficeExtension.ClientObject)) {
							return ClientRequestContext.createErrorPromise(functionName);
						}
						if (array[i].context !=array[0].context) {
							return ClientRequestContext.createErrorPromise(functionName, OfficeExtension.ResourceStrings.invalidRequestContext);
						}
					}
					ctxRetriever=function () { return array[0].context; };
				}
				else {
					return ClientRequestContext.createErrorPromise(functionName);
				}
				batch=receivedRunArgs[argOffset+1];
			}
			else {
				return ClientRequestContext.createErrorPromise(functionName);
			}
			return ClientRequestContext._runCommon(functionName, requestInfo, ctxRetriever, batch, numCleanupAttempts, retryDelay, onCleanupSuccess, onCleanupFailure);
		};
		ClientRequestContext.createErrorPromise=function (functionName, code) {
			if (code===void 0) { code=OfficeExtension.ResourceStrings.invalidArgument; }
			return OfficeExtension._Internal.OfficePromise.reject(OfficeExtension.Utility.createRuntimeError(code, OfficeExtension.Utility._getResourceString(code), functionName));
		};
		ClientRequestContext._runCommon=function (functionName, requestInfo, ctxRetriever, batch, numCleanupAttempts, retryDelay, onCleanupSuccess, onCleanupFailure) {
			if (ClientRequestContext._overrideSession) {
				requestInfo=ClientRequestContext._overrideSession;
			}
			var starterPromise=new OfficeExtension._Internal.OfficePromise(function (resolve, reject) { resolve(); });
			var ctx;
			var succeeded=false;
			var resultOrError;
			return starterPromise
				.then(function () {
				ctx=ctxRetriever(requestInfo);
				if (ctx._autoCleanup) {
					return new OfficeExtension._Internal.OfficePromise(function (resolve, reject) {
						ctx._onRunFinishedNotifiers.push(function () {
							ctx._autoCleanup=true;
							resolve();
						});
					});
				}
				else {
					ctx._autoCleanup=true;
				}
			})
				.then(function () {
				if (typeof batch !=='function') {
					return ClientRequestContext.createErrorPromise(functionName);
				}
				var batchResult=batch(ctx);
				if (OfficeExtension.Utility.isNullOrUndefined(batchResult) || (typeof batchResult.then !=='function')) {
					OfficeExtension.Utility.throwError(OfficeExtension.ResourceStrings.runMustReturnPromise);
				}
				return batchResult;
			})
				.then(function (batchResult) {
				return ctx.sync(batchResult);
			})
				.then(function (result) {
				succeeded=true;
				resultOrError=result;
			})
				.catch(function (error) {
				resultOrError=error;
			})
				.then(function () {
				var itemsToRemove=ctx.trackedObjects._retrieveAndClearAutoCleanupList();
				ctx._autoCleanup=false;
				for (var key in itemsToRemove) {
					itemsToRemove[key]._objectPath.isValid=false;
				}
				var cleanupCounter=0;
				if (OfficeExtension.Utility._synchronousCleanup || ClientRequestContext.isRequestUrlAndHeaderInfoResolver(requestInfo)) {
					return attemptCleanup();
				}
				else {
					attemptCleanup();
				}
				function attemptCleanup() {
					cleanupCounter++;
					for (var key in itemsToRemove) {
						ctx.trackedObjects.remove(itemsToRemove[key]);
					}
					return ctx.sync()
						.then(function () {
						if (onCleanupSuccess) {
							onCleanupSuccess(cleanupCounter);
						}
					})
						.catch(function () {
						if (onCleanupFailure) {
							onCleanupFailure(cleanupCounter);
						}
						if (cleanupCounter < numCleanupAttempts) {
							setTimeout(function () {
								attemptCleanup();
							}, retryDelay);
						}
					});
				}
			})
				.then(function () {
				if (ctx._onRunFinishedNotifiers && ctx._onRunFinishedNotifiers.length > 0) {
					var func=ctx._onRunFinishedNotifiers.shift();
					func();
				}
				if (succeeded) {
					return resultOrError;
				}
				else {
					throw resultOrError;
				}
			});
		};
		ClientRequestContext.prototype._nextId=function () {
			return++this.m_nextId;
		};
		return ClientRequestContext;
	}());
	OfficeExtension.ClientRequestContext=ClientRequestContext;
})(OfficeExtension || (OfficeExtension={}));
var OfficeExtension;
(function (OfficeExtension) {
	var ClientResult=(function () {
		function ClientResult(type) {
			this.m_type=type;
		}
		Object.defineProperty(ClientResult.prototype, "value", {
			get: function () {
				if (!this.m_isLoaded) {
					throw new OfficeExtension._Internal.RuntimeError({
						code: OfficeExtension.ErrorCodes.valueNotLoaded,
						message: OfficeExtension.Utility._getResourceString(OfficeExtension.ResourceStrings.valueNotLoaded),
						debugInfo: {
							errorLocation: "clientResult.value"
						}
					});
				}
				return this.m_value;
			},
			enumerable: true,
			configurable: true
		});
		ClientResult.prototype._handleResult=function (value) {
			this.m_isLoaded=true;
			if (typeof (value)==="object" && value && value._IsNull) {
				return;
			}
			if (this.m_type===1) {
				this.m_value=OfficeExtension.Utility.adjustToDateTime(value);
			}
			else {
				this.m_value=value;
			}
		};
		return ClientResult;
	}());
	OfficeExtension.ClientResult=ClientResult;
})(OfficeExtension || (OfficeExtension={}));
var OfficeExtension;
(function (OfficeExtension) {
	var Constants=(function () {
		function Constants() {
		}
		Constants.flags="flags";
		Constants.getItemAt="GetItemAt";
		Constants.id="Id";
		Constants.idPrivate="_Id";
		Constants.index="_Index";
		Constants.items="_Items";
		Constants.iterativeExecutor="IterativeExecutor";
		Constants.localDocument="http://document.localhost/";
		Constants.localDocumentApiPrefix="http://document.localhost/_api/";
		Constants.processQuery="ProcessQuery";
		Constants.referenceId="_ReferenceId";
		Constants.isTracked="_IsTracked";
		Constants.sourceLibHeader="SdkVersion";
		Constants.embeddingPageOrigin="EmbeddingPageOrigin";
		Constants.embeddingPageSessionInfo="EmbeddingPageSessionInfo";
		Constants.eventMessageCategory=65536;
		return Constants;
	}());
	OfficeExtension.Constants=Constants;
})(OfficeExtension || (OfficeExtension={}));
var OfficeExtension;
(function (OfficeExtension) {
	var versionToken=1;
	var internalConfiguration={
		invokeRequestModifier: function (request) {
			request.DdaMethod.Version=versionToken;
			return request;
		},
		invokeResponseModifier: function (args) {
			versionToken=args.Version;
			if (args.Error) {
				args.error={};
				args.error.Code=args.Error;
			}
			return args;
		}
	};
	var EmbeddedApiStatus;
	(function (EmbeddedApiStatus) {
		EmbeddedApiStatus[EmbeddedApiStatus["Success"]=0]="Success";
		EmbeddedApiStatus[EmbeddedApiStatus["Timeout"]=1]="Timeout";
		EmbeddedApiStatus[EmbeddedApiStatus["InternalError"]=5001]="InternalError";
	})(EmbeddedApiStatus || (EmbeddedApiStatus={}));
	var CommunicationConstants;
	(function (CommunicationConstants) {
		CommunicationConstants.SendingId="sId";
		CommunicationConstants.RespondingId="rId";
		CommunicationConstants.CommandKey="command";
		CommunicationConstants.SessionInfoKey="sessionInfo";
		CommunicationConstants.ParamsKey="params";
		CommunicationConstants.ApiReadyCommand="apiready";
		CommunicationConstants.ExecuteMethodCommand="executeMethod";
		CommunicationConstants.GetAppContextCommand="getAppContext";
		CommunicationConstants.RegisterEventCommand="registerEvent";
		CommunicationConstants.UnregisterEventCommand="unregisterEvent";
		CommunicationConstants.FireEventCommand="fireEvent";
	})(CommunicationConstants || (CommunicationConstants={}));
	var EmbeddedSession=(function (_super) {
		__extends(EmbeddedSession, _super);
		function EmbeddedSession(url, options) {
			_super.call(this);
			this.m_chosenWindow=null;
			this.m_chosenOrigin=null;
			this.m_enabled=true;
			this.m_onMessageHandler=this._onMessage.bind(this);
			this.m_callbackList={};
			this.m_id=0;
			this.m_timeoutId=-1;
			this.m_appContext=null;
			this.m_url=url;
			this.m_options=options;
			if (!this.m_options) {
				this.m_options={ sessionKey: Math.random().toString() };
			}
			if (!this.m_options.sessionKey) {
				this.m_options.sessionKey=Math.random().toString();
			}
			if (!this.m_options.container) {
				this.m_options.container=document.body;
			}
			if (!this.m_options.timeoutInMilliseconds) {
				this.m_options.timeoutInMilliseconds=60000;
			}
			if (!this.m_options.height) {
				this.m_options.height="400px";
			}
			if (!this.m_options.width) {
				this.m_options.width="100%";
			}
		}
		EmbeddedSession.prototype._getIFrameSrc=function () {
			var origin=window.location.protocol+"//"+window.location.host;
			var toAppend=OfficeExtension.Constants.embeddingPageOrigin+"="+encodeURIComponent(origin)+"&"+OfficeExtension.Constants.embeddingPageSessionInfo+"="+encodeURIComponent(this.m_options.sessionKey);
			var useHash=false;
			if (this.m_url.toLowerCase().indexOf("/_layouts/preauth.aspx") > 0) {
				useHash=true;
			}
			var a=document.createElement("a");
			a.href=this.m_url;
			if (useHash) {
				if (a.hash.length===0 || a.hash==="#") {
					a.hash="#"+toAppend;
				}
				else {
					a.hash=a.hash+"&"+toAppend;
				}
			}
			else {
				if (a.search.length===0 || a.search==="?") {
					a.search="?"+toAppend;
				}
				else {
					a.search=a.search+"&"+toAppend;
				}
			}
			var iframeSrc=a.href;
			return iframeSrc;
		};
		EmbeddedSession.prototype.init=function () {
			var _this=this;
			window.addEventListener("message", this.m_onMessageHandler);
			var iframeSrc=this._getIFrameSrc();
			return new OfficeExtension._Internal.OfficePromise(function (resolve, reject) {
				var iframeElement=document.createElement("iframe");
				if (_this.m_options.id) {
					iframeElement.id=_this.m_options.id;
				}
				iframeElement.style.height=_this.m_options.height;
				iframeElement.style.width=_this.m_options.width;
				iframeElement.src=iframeSrc;
				_this.m_options.container.appendChild(iframeElement);
				_this.m_timeoutId=setTimeout(function () {
					_this.close();
					var err=OfficeExtension.Utility.createRuntimeError(OfficeExtension.ErrorCodes.timeout, OfficeExtension.Utility._getResourceString(OfficeExtension.ResourceStrings.timeout), "EmbeddedSession.init");
					reject(err);
				}, _this.m_options.timeoutInMilliseconds);
				_this.m_promiseResolver=resolve;
			});
		};
		EmbeddedSession.prototype._invoke=function (method, callback, params) {
			if (!this.m_enabled) {
				callback(EmbeddedApiStatus.InternalError, null);
				return;
			}
			if (internalConfiguration.invokeRequestModifier) {
				params=internalConfiguration.invokeRequestModifier(params);
			}
			this._sendMessageWithCallback(this.m_id++, method, params, function (args) {
				if (internalConfiguration.invokeResponseModifier) {
					args=internalConfiguration.invokeResponseModifier(args);
				}
				var errorCode=args["Error"];
				delete args["Error"];
				callback(errorCode || EmbeddedApiStatus.Success, args);
			});
		};
		EmbeddedSession.prototype.close=function () {
			window.removeEventListener("message", this.m_onMessageHandler);
			window.clearTimeout(this.m_timeoutId);
			this.m_enabled=false;
		};
		Object.defineProperty(EmbeddedSession.prototype, "eventRegistration", {
			get: function () {
				if (!this.m_sessionEventManager) {
					this.m_sessionEventManager=new OfficeExtension.EventRegistration(this._registerEventImpl.bind(this), this._unregisterEventImpl.bind(this));
				}
				return this.m_sessionEventManager;
			},
			enumerable: true,
			configurable: true
		});
		EmbeddedSession.prototype._createRequestExecutorOrNull=function () {
			return new EmbeddedRequestExecutor(this);
		};
		EmbeddedSession.prototype._resolveRequestUrlAndHeaderInfo=function () {
			return OfficeExtension.Utility._createPromiseFromResult(null);
		};
		EmbeddedSession.prototype._registerEventImpl=function (eventId, targetId) {
			var _this=this;
			return new OfficeExtension._Internal.OfficePromise(function (resolve, reject) {
				_this._sendMessageWithCallback(_this.m_id++, CommunicationConstants.RegisterEventCommand, { EventId: eventId, TargetId: targetId }, function () {
					resolve(null);
				});
			});
		};
		EmbeddedSession.prototype._unregisterEventImpl=function (eventId, targetId) {
			var _this=this;
			return new OfficeExtension._Internal.OfficePromise(function (resolve, reject) {
				_this._sendMessageWithCallback(_this.m_id++, CommunicationConstants.UnregisterEventCommand, { EventId: eventId, TargetId: targetId }, function () {
					resolve();
				});
			});
		};
		EmbeddedSession.prototype._onMessage=function (event) {
			var _this=this;
			if (!this.m_enabled) {
				return;
			}
			if (this.m_chosenWindow
				&& (this.m_chosenWindow !==event.source || this.m_chosenOrigin !==event.origin)) {
				return;
			}
			var eventData=event.data;
			if (eventData && eventData[CommunicationConstants.CommandKey]===CommunicationConstants.ApiReadyCommand) {
				if (!this.m_chosenWindow
					&& this._isValidDescendant(event.source)
					&& eventData[CommunicationConstants.SessionInfoKey]===this.m_options.sessionKey) {
					this.m_chosenWindow=event.source;
					this.m_chosenOrigin=event.origin;
					this._sendMessageWithCallback(this.m_id++, CommunicationConstants.GetAppContextCommand, null, function (appContext) {
						_this._setupContext(appContext);
						window.clearTimeout(_this.m_timeoutId);
						_this.m_promiseResolver();
					});
				}
				return;
			}
			if (eventData && eventData[CommunicationConstants.CommandKey]===CommunicationConstants.FireEventCommand) {
				var msg=eventData[CommunicationConstants.ParamsKey];
				var eventId=msg["EventId"];
				var targetId=msg["TargetId"];
				var data=msg["Data"];
				if (this.m_sessionEventManager) {
					var handlers=this.m_sessionEventManager.getHandlers(eventId, targetId);
					for (var i=0; i < handlers.length; i++) {
						handlers[i](data);
					}
				}
				return;
			}
			if (eventData && eventData.hasOwnProperty(CommunicationConstants.RespondingId)) {
				var rId=eventData[CommunicationConstants.RespondingId];
				var callback=this.m_callbackList[rId];
				if (typeof callback==="function") {
					callback(eventData[CommunicationConstants.ParamsKey]);
				}
				delete this.m_callbackList[rId];
			}
		};
		EmbeddedSession.prototype._sendMessageWithCallback=function (id, command, data, callback) {
			this.m_callbackList[id]=callback;
			var message={};
			message[CommunicationConstants.SendingId]=id;
			message[CommunicationConstants.CommandKey]=command;
			message[CommunicationConstants.ParamsKey]=data;
			this.m_chosenWindow.postMessage(JSON.stringify(message), this.m_chosenOrigin);
		};
		EmbeddedSession.prototype._isValidDescendant=function (wnd) {
			var container=this.m_options.container || document.body;
			function doesFrameWindow(containerWindow) {
				if (containerWindow===wnd) {
					return true;
				}
				for (var i=0, len=containerWindow.frames.length; i < len; i++) {
					if (doesFrameWindow(containerWindow.frames[i])) {
						return true;
					}
				}
				return false;
			}
			var iframes=container.getElementsByTagName("iframe");
			for (var i=0, len=iframes.length; i < len; i++) {
				if (doesFrameWindow(iframes[i].contentWindow)) {
					return true;
				}
			}
			return false;
		};
		EmbeddedSession.prototype._setupContext=function (appContext) {
			if (!(this.m_appContext=appContext)) {
				return;
			}
		};
		return EmbeddedSession;
	}(OfficeExtension.SessionBase));
	OfficeExtension.EmbeddedSession=EmbeddedSession;
	var EmbeddedRequestExecutor=(function () {
		function EmbeddedRequestExecutor(session) {
			this.m_session=session;
		}
		EmbeddedRequestExecutor.prototype.executeAsync=function (customData, requestFlags, requestMessage) {
			var _this=this;
			var messageSafearray=OfficeExtension.RichApiMessageUtility.buildMessageArrayForIRequestExecutor(customData, requestFlags, requestMessage, EmbeddedRequestExecutor.SourceLibHeaderValue);
			return new OfficeExtension._Internal.OfficePromise(function (resolve, reject) {
				_this.m_session._invoke(CommunicationConstants.ExecuteMethodCommand, function (status, result) {
					OfficeExtension.Utility.log("Response:");
					OfficeExtension.Utility.log(JSON.stringify(result));
					var response;
					if (status==EmbeddedApiStatus.Success) {
						response=OfficeExtension.RichApiMessageUtility.buildResponseOnSuccess(OfficeExtension.RichApiMessageUtility.getResponseBodyFromSafeArray(result.Data), OfficeExtension.RichApiMessageUtility.getResponseHeadersFromSafeArray(result.Data));
					}
					else {
						response=OfficeExtension.RichApiMessageUtility.buildResponseOnError(result.error.Code, result.error.Message);
					}
					resolve(response);
				}, EmbeddedRequestExecutor._transformMessageArrayIntoParams(messageSafearray));
			});
		};
		EmbeddedRequestExecutor._transformMessageArrayIntoParams=function (msgArray) {
			return {
				ArrayData: msgArray,
				DdaMethod: {
					DispatchId: EmbeddedRequestExecutor.DispidExecuteRichApiRequestMethod
				}
			};
		};
		EmbeddedRequestExecutor.DispidExecuteRichApiRequestMethod=93;
		EmbeddedRequestExecutor.SourceLibHeaderValue="Embedded";
		return EmbeddedRequestExecutor;
	}());
})(OfficeExtension || (OfficeExtension={}));
var OfficeExtension;
(function (OfficeExtension) {
	var _Internal;
	(function (_Internal) {
		var RuntimeError=(function (_super) {
			__extends(RuntimeError, _super);
			function RuntimeError(error) {
				_super.call(this, (typeof error==="string") ? error : error.message);
				this.name="OfficeExtension.Error";
				if (typeof error==="string") {
					this.message=error;
				}
				else {
					this.code=error.code;
					this.message=error.message;
					this.traceMessages=error.traceMessages || [];
					this.innerError=error.innerError || null;
					this.debugInfo=this._createDebugInfo(error.debugInfo || {});
				}
			}
			RuntimeError.prototype.toString=function () {
				return this.code+': '+this.message;
			};
			RuntimeError.prototype._createDebugInfo=function (partialDebugInfo) {
				var debugInfo={
					code: this.code,
					message: this.message,
					toString: function () {
						return JSON.stringify(this);
					}
				};
				for (var key in partialDebugInfo) {
					debugInfo[key]=partialDebugInfo[key];
				}
				if (this.innerError) {
					if (this.innerError instanceof OfficeExtension.Error) {
						debugInfo.innerError=this.innerError.debugInfo;
					}
					else {
						debugInfo.innerError=this.innerError;
					}
				}
				return debugInfo;
			};
			RuntimeError._createInvalidArgError=function (error) {
				return new _Internal.RuntimeError({
					code: OfficeExtension.ErrorCodes.invalidArgument,
					message: (OfficeExtension.Utility.isNullOrEmptyString(error.argumentName) ?
						OfficeExtension.Utility._getResourceString(OfficeExtension.ResourceStrings.invalidArgumentGeneric) :
						OfficeExtension.Utility._getResourceString(OfficeExtension.ResourceStrings.invalidArgument, error.argumentName)),
					debugInfo: error.errorLocation ? { errorLocation: error.errorLocation } : {},
					innerError: error.innerError
				});
			};
			return RuntimeError;
		}(Error));
		_Internal.RuntimeError=RuntimeError;
	})(_Internal=OfficeExtension._Internal || (OfficeExtension._Internal={}));
	OfficeExtension.Error=_Internal.RuntimeError;
})(OfficeExtension || (OfficeExtension={}));
var OfficeExtension;
(function (OfficeExtension) {
	var ErrorCodes=(function () {
		function ErrorCodes() {
		}
		ErrorCodes.accessDenied="AccessDenied";
		ErrorCodes.generalException="GeneralException";
		ErrorCodes.activityLimitReached="ActivityLimitReached";
		ErrorCodes.invalidObjectPath="InvalidObjectPath";
		ErrorCodes.propertyNotLoaded="PropertyNotLoaded";
		ErrorCodes.valueNotLoaded="ValueNotLoaded";
		ErrorCodes.invalidRequestContext="InvalidRequestContext";
		ErrorCodes.invalidArgument="InvalidArgument";
		ErrorCodes.runMustReturnPromise="RunMustReturnPromise";
		ErrorCodes.cannotRegisterEvent="CannotRegisterEvent";
		ErrorCodes.apiNotFound="ApiNotFound";
		ErrorCodes.connectionFailure="ConnectionFailure";
		ErrorCodes.timeout="Timeout";
		ErrorCodes.invalidOrTimedOutSession="InvalidOrTimedOutSession";
		return ErrorCodes;
	}());
	OfficeExtension.ErrorCodes=ErrorCodes;
})(OfficeExtension || (OfficeExtension={}));
var OfficeExtension;
(function (OfficeExtension) {
	var EventHandlers=(function () {
		function EventHandlers(context, parentObject, name, eventInfo) {
			var _this=this;
			this.m_id=context._nextId();
			this.m_context=context;
			this.m_name=name;
			this.m_handlers=[];
			this.m_registered=false;
			this.m_eventInfo=eventInfo;
			this.m_callback=function (args) {
				_this.m_eventInfo.eventArgsTransformFunc(args)
					.then(function (newArgs) { return _this.fireEvent(newArgs); });
			};
		}
		Object.defineProperty(EventHandlers.prototype, "_registered", {
			get: function () {
				return this.m_registered;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(EventHandlers.prototype, "_id", {
			get: function () {
				return this.m_id;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(EventHandlers.prototype, "_handlers", {
			get: function () {
				return this.m_handlers;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(EventHandlers.prototype, "_context", {
			get: function () {
				return this.m_context;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(EventHandlers.prototype, "_callback", {
			get: function () {
				return this.m_callback;
			},
			enumerable: true,
			configurable: true
		});
		EventHandlers.prototype.add=function (handler) {
			var action=OfficeExtension.ActionFactory.createTraceAction(this.m_context, null, false);
			this.m_context._pendingRequest._addPendingEventHandlerAction(this, { id: action.actionInfo.Id, handler: handler, operation: 0 });
			return new OfficeExtension.EventHandlerResult(this.m_context, this, handler);
		};
		EventHandlers.prototype.remove=function (handler) {
			var action=OfficeExtension.ActionFactory.createTraceAction(this.m_context, null, false);
			this.m_context._pendingRequest._addPendingEventHandlerAction(this, { id: action.actionInfo.Id, handler: handler, operation: 1 });
		};
		EventHandlers.prototype.removeAll=function () {
			var action=OfficeExtension.ActionFactory.createTraceAction(this.m_context, null, false);
			this.m_context._pendingRequest._addPendingEventHandlerAction(this, { id: action.actionInfo.Id, handler: null, operation: 2 });
		};
		EventHandlers.prototype._processRegistration=function (req) {
			var _this=this;
			var ret=OfficeExtension.Utility._createPromiseFromResult(null);
			var actions=req._getPendingEventHandlerActions(this);
			if (!actions) {
				return ret;
			}
			var handlersResult=[];
			for (var i=0; i < this.m_handlers.length; i++) {
				handlersResult.push(this.m_handlers[i]);
			}
			var hasChange=false;
			for (var i=0; i < actions.length; i++) {
				if (req._responseTraceIds[actions[i].id]) {
					hasChange=true;
					switch (actions[i].operation) {
						case 0:
							handlersResult.push(actions[i].handler);
							break;
						case 1:
							for (var index=handlersResult.length - 1; index >=0; index--) {
								if (handlersResult[index]===actions[i].handler) {
									handlersResult.splice(index, 1);
									break;
								}
							}
							break;
						case 2:
							handlersResult=[];
							break;
					}
				}
			}
			if (hasChange) {
				if (!this.m_registered && handlersResult.length > 0) {
					ret=ret
						.then(function () { return _this.m_eventInfo.registerFunc(_this.m_callback); })
						.then(function () { return (_this.m_registered=true); });
				}
				else if (this.m_registered && handlersResult.length==0) {
					ret=ret
						.then(function () { return _this.m_eventInfo.unregisterFunc(_this.m_callback); })
						.catch(function (ex) {
						OfficeExtension.Utility.log("Error when unregister event: "+JSON.stringify(ex));
					})
						.then(function () { return (_this.m_registered=false); });
				}
				ret=ret
					.then(function () { return (_this.m_handlers=handlersResult); });
			}
			return ret;
		};
		EventHandlers.prototype.fireEvent=function (args) {
			var promises=[];
			for (var i=0; i < this.m_handlers.length; i++) {
				var handler=this.m_handlers[i];
				var p=OfficeExtension.Utility._createPromiseFromResult(null)
					.then(this.createFireOneEventHandlerFunc(handler, args))
					.catch(function (ex) {
					OfficeExtension.Utility.log("Error when invoke handler: "+JSON.stringify(ex));
				});
				promises.push(p);
			}
			OfficeExtension._Internal.OfficePromise.all(promises);
		};
		EventHandlers.prototype.createFireOneEventHandlerFunc=function (handler, args) {
			return function () { return handler(args); };
		};
		return EventHandlers;
	}());
	OfficeExtension.EventHandlers=EventHandlers;
})(OfficeExtension || (OfficeExtension={}));
var OfficeExtension;
(function (OfficeExtension) {
	var EventHandlerResult=(function () {
		function EventHandlerResult(context, handlers, handler) {
			this.m_context=context;
			this.m_allHandlers=handlers;
			this.m_handler=handler;
		}
		EventHandlerResult.prototype.remove=function () {
			if (this.m_allHandlers && this.m_handler) {
				this.m_allHandlers.remove(this.m_handler);
				this.m_allHandlers=null;
				this.m_handler=null;
			}
		};
		return EventHandlerResult;
	}());
	OfficeExtension.EventHandlerResult=EventHandlerResult;
})(OfficeExtension || (OfficeExtension={}));
var OfficeExtension;
(function (OfficeExtension) {
	var _Internal;
	(function (_Internal) {
		var OfficeJsEventRegistration=(function () {
			function OfficeJsEventRegistration() {
			}
			OfficeJsEventRegistration.prototype.register=function (eventId, targetId, handler) {
				switch (eventId) {
					case 4:
						return OfficeExtension.Utility.promisify(function (callback) { return Office.context.document.bindings.getByIdAsync(targetId, callback); })
							.then(function (officeBinding) {
							return OfficeExtension.Utility.promisify(function (callback) { return officeBinding.addHandlerAsync(Office.EventType.BindingDataChanged, handler, callback); });
						});
					case 3:
						return OfficeExtension.Utility.promisify(function (callback) { return Office.context.document.bindings.getByIdAsync(targetId, callback); })
							.then(function (officeBinding) {
							return OfficeExtension.Utility.promisify(function (callback) { return officeBinding.addHandlerAsync(Office.EventType.BindingSelectionChanged, handler, callback); });
						});
					case 2:
						return OfficeExtension.Utility.promisify(function (callback) { return Office.context.document.addHandlerAsync(Office.EventType.DocumentSelectionChanged, handler, callback); });
					case 1:
						return OfficeExtension.Utility.promisify(function (callback) { return Office.context.document.settings.addHandlerAsync(Office.EventType.SettingsChanged, handler, callback); });
					case 5:
						return OfficeExtension.Utility.promisify(function (callback) { return OSF.DDA.RichApi.richApiMessageManager.addHandlerAsync("richApiMessage", handler, callback); });
					case 13:
						return OfficeExtension.Utility.promisify(function (callback) { return Office.context.document.addHandlerAsync(Office.EventType.ObjectDeleted, handler, { id: targetId }, callback); });
					case 14:
						return OfficeExtension.Utility.promisify(function (callback) { return Office.context.document.addHandlerAsync(Office.EventType.ObjectSelectionChanged, handler, { id: targetId }, callback); });
					case 15:
						return OfficeExtension.Utility.promisify(function (callback) { return Office.context.document.addHandlerAsync(Office.EventType.ObjectDataChanged, handler, { id: targetId }, callback); });
					case 16:
						return OfficeExtension.Utility.promisify(function (callback) { return Office.context.document.addHandlerAsync(Office.EventType.ContentControlAdded, handler, { id: targetId }, callback); });
					default:
						throw _Internal.RuntimeError._createInvalidArgError("eventId");
				}
			};
			OfficeJsEventRegistration.prototype.unregister=function (eventId, targetId, handler) {
				switch (eventId) {
					case 4:
						return OfficeExtension.Utility.promisify(function (callback) { return Office.context.document.bindings.getByIdAsync(targetId, callback); })
							.then(function (officeBinding) {
							return OfficeExtension.Utility.promisify(function (callback) { return officeBinding.removeHandlerAsync(Office.EventType.BindingDataChanged, { handler: handler }, callback); });
						});
					case 3:
						return OfficeExtension.Utility.promisify(function (callback) { return Office.context.document.bindings.getByIdAsync(targetId, callback); })
							.then(function (officeBinding) {
							return OfficeExtension.Utility.promisify(function (callback) { return officeBinding.removeHandlerAsync(Office.EventType.BindingSelectionChanged, { handler: handler }, callback); });
						});
					case 2:
						return OfficeExtension.Utility.promisify(function (callback) { return Office.context.document.removeHandlerAsync(Office.EventType.DocumentSelectionChanged, { handler: handler }, callback); });
					case 1:
						return OfficeExtension.Utility.promisify(function (callback) { return Office.context.document.settings.removeHandlerAsync(Office.EventType.SettingsChanged, { handler: handler }, callback); });
					case 5:
						return OfficeExtension.Utility.promisify(function (callback) { return OSF.DDA.RichApi.richApiMessageManager.removeHandlerAsync("richApiMessage", { handler: handler }, callback); });
					case 13:
						return OfficeExtension.Utility.promisify(function (callback) { return Office.context.document.removeHandlerAsync(Office.EventType.ObjectDeleted, { id: targetId, handler: handler }, callback); });
					case 14:
						return OfficeExtension.Utility.promisify(function (callback) { return Office.context.document.removeHandlerAsync(Office.EventType.ObjectSelectionChanged, { id: targetId, handler: handler }, callback); });
					case 15:
						return OfficeExtension.Utility.promisify(function (callback) { return Office.context.document.removeHandlerAsync(Office.EventType.ObjectDataChanged, { id: targetId, handler: handler }, callback); });
					case 16:
						return OfficeExtension.Utility.promisify(function (callback) { return Office.context.document.removeHandlerAsync(Office.EventType.ContentControlAdded, { id: targetId, handler: handler }, callback); });
					default:
						throw _Internal.RuntimeError._createInvalidArgError("eventId");
				}
			};
			return OfficeJsEventRegistration;
		}());
		_Internal.officeJsEventRegistration=new OfficeJsEventRegistration();
	})(_Internal=OfficeExtension._Internal || (OfficeExtension._Internal={}));
	var EventRegistration=(function () {
		function EventRegistration(registerEventImpl, unregisterEventImpl) {
			this.m_handlersByEventByTarget={};
			this.m_registerEventImpl=registerEventImpl;
			this.m_unregisterEventImpl=unregisterEventImpl;
		}
		EventRegistration.prototype.getHandlers=function (eventId, targetId) {
			if (OfficeExtension.Utility.isNullOrUndefined(targetId)) {
				targetId="";
			}
			var handlersById=this.m_handlersByEventByTarget[eventId];
			if (!handlersById) {
				handlersById={};
				this.m_handlersByEventByTarget[eventId]=handlersById;
			}
			var handlers=handlersById[targetId];
			if (!handlers) {
				handlers=[];
				handlersById[targetId]=handlers;
			}
			return handlers;
		};
		EventRegistration.prototype.register=function (eventId, targetId, handler) {
			if (!handler) {
				throw _Internal.RuntimeError._createInvalidArgError("handler");
			}
			var handlers=this.getHandlers(eventId, targetId);
			handlers.push(handler);
			if (handlers.length===1) {
				return this.m_registerEventImpl(eventId, targetId);
			}
			return OfficeExtension.Utility._createPromiseFromResult(null);
		};
		EventRegistration.prototype.unregister=function (eventId, targetId, handler) {
			if (!handler) {
				throw _Internal.RuntimeError._createInvalidArgError("handler");
			}
			var handlers=this.getHandlers(eventId, targetId);
			for (var index=handlers.length - 1; index >=0; index--) {
				if (handlers[index]===handler) {
					handlers.splice(index, 1);
					break;
				}
			}
			if (handlers.length===0) {
				return this.m_unregisterEventImpl(eventId, targetId);
			}
			return OfficeExtension.Utility._createPromiseFromResult(null);
		};
		return EventRegistration;
	}());
	OfficeExtension.EventRegistration=EventRegistration;
})(OfficeExtension || (OfficeExtension={}));
var OfficeExtension;
(function (OfficeExtension) {
	var GenericEventRegistration=(function () {
		function GenericEventRegistration() {
			this.m_eventRegistration=new OfficeExtension.EventRegistration(this._registerEventImpl.bind(this), this._unregisterEventImpl.bind(this));
			this.m_richApiMessageHandler=this._handleRichApiMessage.bind(this);
		}
		GenericEventRegistration.prototype.ready=function () {
			var _this=this;
			if (!this.m_ready) {
				if (GenericEventRegistration._testReadyImpl) {
					this.m_ready=GenericEventRegistration._testReadyImpl()
						.then(function () {
						_this.m_isReady=true;
					});
				}
				else {
					this.m_ready=OfficeExtension._Internal.officeJsEventRegistration.register(5, "", this.m_richApiMessageHandler)
						.then(function () {
						_this.m_isReady=true;
					});
				}
			}
			return this.m_ready;
		};
		Object.defineProperty(GenericEventRegistration.prototype, "isReady", {
			get: function () {
				return this.m_isReady;
			},
			enumerable: true,
			configurable: true
		});
		GenericEventRegistration.prototype.register=function (eventId, targetId, handler) {
			var _this=this;
			return this.ready()
				.then(function () { return _this.m_eventRegistration.register(eventId, targetId, handler); });
		};
		GenericEventRegistration.prototype.unregister=function (eventId, targetId, handler) {
			var _this=this;
			return this.ready()
				.then(function () { return _this.m_eventRegistration.unregister(eventId, targetId, handler); });
		};
		GenericEventRegistration.prototype._registerEventImpl=function (eventId, targetId) {
			return OfficeExtension.Utility._createPromiseFromResult(null);
		};
		GenericEventRegistration.prototype._unregisterEventImpl=function (eventId, targetId) {
			return OfficeExtension.Utility._createPromiseFromResult(null);
		};
		GenericEventRegistration.prototype._handleRichApiMessage=function (msg) {
			if (msg && msg.entries) {
				for (var entryIndex=0; entryIndex < msg.entries.length; entryIndex++) {
					var entry=msg.entries[entryIndex];
					if (entry.messageCategory==OfficeExtension.Constants.eventMessageCategory) {
						if (OfficeExtension.Utility._logEnabled) {
							OfficeExtension.Utility.log(JSON.stringify(entry));
						}
						var funcs=this.m_eventRegistration.getHandlers(entry.messageType, entry.targetId);
						if (funcs.length > 0) {
							var arg=JSON.parse(entry.message);
							for (var i=0; i < funcs.length; i++) {
								funcs[i](arg);
							}
						}
					}
				}
			}
		};
		GenericEventRegistration.getGenericEventRegistration=function () {
			if (!GenericEventRegistration.s_genericEventRegistration) {
				GenericEventRegistration.s_genericEventRegistration=new GenericEventRegistration();
			}
			return GenericEventRegistration.s_genericEventRegistration;
		};
		GenericEventRegistration.richApiMessageEventCategory=65536;
		return GenericEventRegistration;
	}());
	function _testSetRichApiMessageReadyImpl(impl) {
		GenericEventRegistration._testReadyImpl=impl;
	}
	OfficeExtension._testSetRichApiMessageReadyImpl=_testSetRichApiMessageReadyImpl;
	function _testTriggerRichApiMessageEvent(msg) {
		GenericEventRegistration.getGenericEventRegistration()._handleRichApiMessage(msg);
	}
	OfficeExtension._testTriggerRichApiMessageEvent=_testTriggerRichApiMessageEvent;
	var GenericEventHandlers=(function (_super) {
		__extends(GenericEventHandlers, _super);
		function GenericEventHandlers(context, parentObject, name, eventInfo) {
			_super.call(this, context, parentObject, name, eventInfo);
			this.m_genericEventInfo=eventInfo;
		}
		GenericEventHandlers.prototype.add=function (handler) {
			var _this=this;
			if (this.m_genericEventInfo.registerFunc) {
				this.m_genericEventInfo.registerFunc();
			}
			if (!GenericEventRegistration.getGenericEventRegistration().isReady) {
				this._context._pendingRequest._addPreSyncPromise(GenericEventRegistration.getGenericEventRegistration().ready());
			}
			OfficeExtension.ActionFactory.createTraceMarkerForCallback(this._context, function () {
				_this._handlers.push(handler);
				if (_this._handlers.length==1) {
					GenericEventRegistration.getGenericEventRegistration().register(_this.m_genericEventInfo.eventType, _this.m_genericEventInfo.getTargetIdFunc(), _this._callback);
				}
			});
			return new OfficeExtension.EventHandlerResult(this._context, this, handler);
		};
		GenericEventHandlers.prototype.remove=function (handler) {
			var _this=this;
			if (this.m_genericEventInfo.unregisterFunc) {
				this.m_genericEventInfo.unregisterFunc();
			}
			OfficeExtension.ActionFactory.createTraceMarkerForCallback(this._context, function () {
				var handlers=_this._handlers;
				for (var index=handlers.length - 1; index >=0; index--) {
					if (handlers[index]===handler) {
						handlers.splice(index, 1);
						break;
					}
				}
				if (handlers.length==0) {
					GenericEventRegistration.getGenericEventRegistration().unregister(_this.m_genericEventInfo.eventType, _this.m_genericEventInfo.getTargetIdFunc(), _this._callback);
				}
			});
		};
		GenericEventHandlers.prototype.removeAll=function () {
		};
		return GenericEventHandlers;
	}(OfficeExtension.EventHandlers));
	OfficeExtension.GenericEventHandlers=GenericEventHandlers;
})(OfficeExtension || (OfficeExtension={}));
var OfficeExtension;
(function (OfficeExtension) {
	var HttpRequestExecutor=(function () {
		function HttpRequestExecutor() {
		}
		HttpRequestExecutor.prototype.executeAsync=function (customData, requestFlags, requestMessage) {
			var requestMessageText=JSON.stringify(requestMessage.Body);
			var url=requestMessage.Url;
			if (url.charAt(url.length - 1) !="/") {
				url=url+"/";
			}
			url=url+OfficeExtension.Constants.processQuery;
			url=url+"?"+OfficeExtension.Constants.flags+"="+requestFlags.toString();
			var requestInfo={
				method: "POST",
				url: url,
				headers: {},
				body: requestMessageText
			};
			requestInfo.headers[OfficeExtension.Constants.sourceLibHeader]=HttpRequestExecutor.SourceLibHeaderValue;
			requestInfo.headers["CONTENT-TYPE"]="application/json";
			if (requestMessage.Headers) {
				for (var key in requestMessage.Headers) {
					requestInfo.headers[key]=requestMessage.Headers[key];
				}
			}
			return OfficeExtension.HttpUtility.sendRequest(requestInfo)
				.then(function (responseInfo) {
				var response;
				if (responseInfo.statusCode===200) {
					response={ ErrorCode: null, ErrorMessage: null, Headers: responseInfo.headers, Body: JSON.parse(responseInfo.body) };
				}
				else {
					OfficeExtension.Utility.log("Error Response:"+responseInfo.body);
					var error=OfficeExtension.Utility._parseErrorResponse(responseInfo);
					response={
						ErrorCode: error.errorCode,
						ErrorMessage: error.errorMessage,
						Headers: responseInfo.headers,
						Body: null
					};
				}
				return response;
			});
		};
		HttpRequestExecutor.SourceLibHeaderValue="officejs-rest";
		return HttpRequestExecutor;
	}());
	OfficeExtension.HttpRequestExecutor=HttpRequestExecutor;
})(OfficeExtension || (OfficeExtension={}));
var OfficeExtension;
(function (OfficeExtension) {
	var HttpUtility=(function () {
		function HttpUtility() {
		}
		HttpUtility.setCustomSendRequestFunc=function (func) {
			HttpUtility.s_customSendRequestFunc=func;
		};
		HttpUtility.xhrSendRequestFunc=function (request) {
			return new OfficeExtension._Internal.OfficePromise(function (resolve, reject) {
				var xhr=new XMLHttpRequest();
				xhr.open(request.method, request.url);
				xhr.onload=function () {
					var resp={
						statusCode: xhr.status,
						headers: OfficeExtension.Utility._parseHttpResponseHeaders(xhr.getAllResponseHeaders()),
						body: xhr.responseText
					};
					resolve(resp);
				};
				xhr.onerror=function () {
					reject(new OfficeExtension._Internal.RuntimeError({
						code: OfficeExtension.ErrorCodes.connectionFailure,
						message: OfficeExtension.Utility._getResourceString(OfficeExtension.ResourceStrings.connectionFailureWithStatus, xhr.statusText)
					}));
				};
				if (request.headers) {
					for (var key in request.headers) {
						xhr.setRequestHeader(key, request.headers[key]);
					}
				}
				xhr.send(request.body);
			});
		};
		HttpUtility.sendRequest=function (request) {
			HttpUtility.validateAndNormalizeRequest(request);
			var func=HttpUtility.s_customSendRequestFunc;
			if (!func) {
				func=HttpUtility.xhrSendRequestFunc;
			}
			return func(request);
		};
		HttpUtility.setCustomSendLocalDocumentRequestFunc=function (func) {
			HttpUtility.s_customSendLocalDocumentRequestFunc=func;
		};
		HttpUtility.sendLocalDocumentRequest=function (request) {
			HttpUtility.validateAndNormalizeRequest(request);
			var func;
			func=HttpUtility.s_customSendLocalDocumentRequestFunc || HttpUtility.officeJsSendLocalDocumentRequestFunc;
			return func(request);
		};
		HttpUtility.officeJsSendLocalDocumentRequestFunc=function (request) {
			request=OfficeExtension.Utility._validateLocalDocumentRequest(request);
			var requestSafeArray=OfficeExtension.Utility._buildRequestMessageSafeArray(request);
			return new OfficeExtension._Internal.OfficePromise(function (resolve, reject) {
				OSF.DDA.RichApi.executeRichApiRequestAsync(requestSafeArray, function (asyncResult) {
					var response;
					if (asyncResult.status=="succeeded") {
						response=							{
								statusCode: OfficeExtension.RichApiMessageUtility.getResponseStatusCode(asyncResult),
								headers: OfficeExtension.RichApiMessageUtility.getResponseHeaders(asyncResult),
								body: OfficeExtension.RichApiMessageUtility.getResponseBody(asyncResult)
							};
					}
					else {
						response=OfficeExtension.RichApiMessageUtility.buildHttpResponseFromOfficeJsError(asyncResult.error.code, asyncResult.error.message);
					}
					OfficeExtension.Utility.log(JSON.stringify(response));
					resolve(response);
				});
			});
		};
		HttpUtility.validateAndNormalizeRequest=function (request) {
			if (OfficeExtension.Utility.isNullOrUndefined(request)) {
				throw OfficeExtension._Internal.RuntimeError._createInvalidArgError({
					argumentName: "request"
				});
			}
			if (OfficeExtension.Utility.isNullOrEmptyString(request.method)) {
				request.method="GET";
			}
			request.method=request.method.toUpperCase();
		};
		HttpUtility.logRequest=function (request) {
			if (OfficeExtension.Utility._logEnabled) {
				OfficeExtension.Utility.log("---HTTP Request---");
				OfficeExtension.Utility.log(request.method+" "+request.url);
				if (request.headers) {
					for (var key in request.headers) {
						OfficeExtension.Utility.log(key+": "+request.headers[key]);
					}
				}
				if (HttpUtility._logBody) {
					OfficeExtension.Utility.log(request.body);
				}
			}
		};
		HttpUtility.logResponse=function (response) {
			if (OfficeExtension.Utility._logEnabled) {
				OfficeExtension.Utility.log("---HTTP Response---");
				OfficeExtension.Utility.log(""+response.statusCode);
				if (response.headers) {
					for (var key in response.headers) {
						OfficeExtension.Utility.log(key+": "+response.headers[key]);
					}
				}
				if (HttpUtility._logBody) {
					OfficeExtension.Utility.log(response.body);
				}
			}
		};
		HttpUtility._logBody=false;
		return HttpUtility;
	}());
	OfficeExtension.HttpUtility=HttpUtility;
})(OfficeExtension || (OfficeExtension={}));
var OfficeExtension;
(function (OfficeExtension) {
	var InstantiateActionResultHandler=(function () {
		function InstantiateActionResultHandler(clientObject) {
			this.m_clientObject=clientObject;
		}
		InstantiateActionResultHandler.prototype._handleResult=function (value) {
			this.m_clientObject._handleIdResult(value);
		};
		return InstantiateActionResultHandler;
	}());
	OfficeExtension.InstantiateActionResultHandler=InstantiateActionResultHandler;
})(OfficeExtension || (OfficeExtension={}));
var OfficeExtension;
(function (OfficeExtension) {
	;
})(OfficeExtension || (OfficeExtension={}));
var OfficeExtension;
(function (OfficeExtension) {
	var ObjectPath=(function () {
		function ObjectPath(objectPathInfo, parentObjectPath, isCollection, isInvalidAfterRequest) {
			this.m_objectPathInfo=objectPathInfo;
			this.m_parentObjectPath=parentObjectPath;
			this.m_isWriteOperation=false;
			this.m_isCollection=isCollection;
			this.m_isInvalidAfterRequest=isInvalidAfterRequest;
			this.m_isValid=true;
		}
		Object.defineProperty(ObjectPath.prototype, "objectPathInfo", {
			get: function () {
				return this.m_objectPathInfo;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ObjectPath.prototype, "isWriteOperation", {
			get: function () {
				return this.m_isWriteOperation;
			},
			set: function (value) {
				this.m_isWriteOperation=value;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ObjectPath.prototype, "isCollection", {
			get: function () {
				return this.m_isCollection;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ObjectPath.prototype, "isInvalidAfterRequest", {
			get: function () {
				return this.m_isInvalidAfterRequest;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ObjectPath.prototype, "parentObjectPath", {
			get: function () {
				return this.m_parentObjectPath;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ObjectPath.prototype, "argumentObjectPaths", {
			get: function () {
				return this.m_argumentObjectPaths;
			},
			set: function (value) {
				this.m_argumentObjectPaths=value;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ObjectPath.prototype, "isValid", {
			get: function () {
				return this.m_isValid;
			},
			set: function (value) {
				this.m_isValid=value;
				if (!value &&
					this.m_objectPathInfo.ObjectPathType===6 &&
					this.m_savedObjectPathInfo) {
					ObjectPath.copyObjectPathInfo(this.m_savedObjectPathInfo.pathInfo, this.m_objectPathInfo);
					this.m_parentObjectPath=this.m_savedObjectPathInfo.parent;
					this.m_isValid=true;
					this.m_savedObjectPathInfo=null;
				}
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ObjectPath.prototype, "getByIdMethodName", {
			get: function () {
				return this.m_getByIdMethodName;
			},
			set: function (value) {
				this.m_getByIdMethodName=value;
			},
			enumerable: true,
			configurable: true
		});
		ObjectPath.prototype._updateAsNullObject=function () {
			this.m_isInvalidAfterRequest=false;
			this.m_isValid=true;
			this.m_objectPathInfo.ObjectPathType=7;
			this.m_objectPathInfo.Name="";
			this.m_objectPathInfo.ArgumentInfo={};
			this.m_parentObjectPath=null;
			this.m_argumentObjectPaths=null;
		};
		ObjectPath.prototype.updateUsingObjectData=function (value) {
			var referenceId=value[OfficeExtension.Constants.referenceId];
			if (!OfficeExtension.Utility.isNullOrEmptyString(referenceId)) {
				if (!this.m_savedObjectPathInfo &&
					!this.isInvalidAfterRequest &&
					ObjectPath.isRestorableObjectPath(this.m_objectPathInfo.ObjectPathType)) {
					var pathInfo={};
					ObjectPath.copyObjectPathInfo(this.m_objectPathInfo, pathInfo);
					this.m_savedObjectPathInfo={
						pathInfo: pathInfo,
						parent: this.m_parentObjectPath
					};
				}
				this.m_isInvalidAfterRequest=false;
				this.m_isValid=true;
				this.m_objectPathInfo.ObjectPathType=6;
				this.m_objectPathInfo.Name=referenceId;
				this.m_objectPathInfo.ArgumentInfo={};
				this.m_parentObjectPath=null;
				this.m_argumentObjectPaths=null;
				return;
			}
			var parentIsCollection=this.parentObjectPath && this.parentObjectPath.isCollection;
			var getByIdMethodName=this.getByIdMethodName;
			if (parentIsCollection || !OfficeExtension.Utility.isNullOrEmptyString(getByIdMethodName)) {
				var id=value[OfficeExtension.Constants.id];
				if (OfficeExtension.Utility.isNullOrUndefined(id)) {
					id=value[OfficeExtension.Constants.idPrivate];
				}
				if (!OfficeExtension.Utility.isNullOrUndefined(id)) {
					this.m_isInvalidAfterRequest=false;
					this.m_isValid=true;
					if (!OfficeExtension.Utility.isNullOrEmptyString(getByIdMethodName)) {
						this.m_objectPathInfo.ObjectPathType=3;
						this.m_objectPathInfo.Name=getByIdMethodName;
						this.m_getByIdMethodName=null;
					}
					else {
						this.m_objectPathInfo.ObjectPathType=5;
						this.m_objectPathInfo.Name="";
					}
					this.isWriteOperation=false;
					this.m_objectPathInfo.ArgumentInfo={};
					this.m_objectPathInfo.ArgumentInfo.Arguments=[id];
					this.m_argumentObjectPaths=null;
					return;
				}
			}
		};
		ObjectPath.isRestorableObjectPath=function (objectPathType) {
			return (objectPathType===1 ||
				objectPathType===5 ||
				objectPathType===3 ||
				objectPathType===4);
		};
		ObjectPath.copyObjectPathInfo=function (src, dest) {
			dest.Id=src.Id;
			dest.ArgumentInfo=src.ArgumentInfo;
			dest.Name=src.Name;
			dest.ObjectPathType=src.ObjectPathType;
			dest.ParentObjectPathId=src.ParentObjectPathId;
		};
		return ObjectPath;
	}());
	OfficeExtension.ObjectPath=ObjectPath;
})(OfficeExtension || (OfficeExtension={}));
var OfficeExtension;
(function (OfficeExtension) {
	var ObjectPathFactory=(function () {
		function ObjectPathFactory() {
		}
		ObjectPathFactory.createGlobalObjectObjectPath=function (context) {
			var objectPathInfo={ Id: context._nextId(), ObjectPathType: 1, Name: "" };
			return new OfficeExtension.ObjectPath(objectPathInfo, null, false, false);
		};
		ObjectPathFactory.createNewObjectObjectPath=function (context, typeName, isCollection) {
			var objectPathInfo={ Id: context._nextId(), ObjectPathType: 2, Name: typeName };
			return new OfficeExtension.ObjectPath(objectPathInfo, null, isCollection, false);
		};
		ObjectPathFactory.createPropertyObjectPath=function (context, parent, propertyName, isCollection, isInvalidAfterRequest) {
			var objectPathInfo={
				Id: context._nextId(),
				ObjectPathType: 4,
				Name: propertyName,
				ParentObjectPathId: parent._objectPath.objectPathInfo.Id,
			};
			return new OfficeExtension.ObjectPath(objectPathInfo, parent._objectPath, isCollection, isInvalidAfterRequest);
		};
		ObjectPathFactory.createIndexerObjectPath=function (context, parent, args) {
			var objectPathInfo={
				Id: context._nextId(),
				ObjectPathType: 5,
				Name: "",
				ParentObjectPathId: parent._objectPath.objectPathInfo.Id,
				ArgumentInfo: {}
			};
			objectPathInfo.ArgumentInfo.Arguments=args;
			return new OfficeExtension.ObjectPath(objectPathInfo, parent._objectPath, false, false);
		};
		ObjectPathFactory.createIndexerObjectPathUsingParentPath=function (context, parentObjectPath, args) {
			var objectPathInfo={
				Id: context._nextId(),
				ObjectPathType: 5,
				Name: "",
				ParentObjectPathId: parentObjectPath.objectPathInfo.Id,
				ArgumentInfo: {}
			};
			objectPathInfo.ArgumentInfo.Arguments=args;
			return new OfficeExtension.ObjectPath(objectPathInfo, parentObjectPath, false, false);
		};
		ObjectPathFactory.createMethodObjectPath=function (context, parent, methodName, operationType, args, isCollection, isInvalidAfterRequest, getByIdMethodName) {
			var objectPathInfo={
				Id: context._nextId(),
				ObjectPathType: 3,
				Name: methodName,
				ParentObjectPathId: parent._objectPath.objectPathInfo.Id,
				ArgumentInfo: {}
			};
			var argumentObjectPaths=OfficeExtension.Utility.setMethodArguments(context, objectPathInfo.ArgumentInfo, args);
			var ret=new OfficeExtension.ObjectPath(objectPathInfo, parent._objectPath, isCollection, isInvalidAfterRequest);
			ret.argumentObjectPaths=argumentObjectPaths;
			ret.isWriteOperation=(operationType !=1);
			ret.getByIdMethodName=getByIdMethodName;
			return ret;
		};
		ObjectPathFactory.createReferenceIdObjectPath=function (context, referenceId) {
			var objectPathInfo={
				Id: context._nextId(),
				ObjectPathType: 6,
				Name: referenceId,
				ArgumentInfo: {}
			};
			var ret=new OfficeExtension.ObjectPath(objectPathInfo, null, false, false);
			return ret;
		};
		ObjectPathFactory.createChildItemObjectPathUsingIndexerOrGetItemAt=function (hasIndexerMethod, context, parent, childItem, index) {
			var id=childItem[OfficeExtension.Constants.id];
			if (OfficeExtension.Utility.isNullOrUndefined(id)) {
				id=childItem[OfficeExtension.Constants.idPrivate];
			}
			if (hasIndexerMethod && !OfficeExtension.Utility.isNullOrUndefined(id)) {
				return ObjectPathFactory.createChildItemObjectPathUsingIndexer(context, parent, childItem);
			}
			else {
				return ObjectPathFactory.createChildItemObjectPathUsingGetItemAt(context, parent, childItem, index);
			}
		};
		ObjectPathFactory.createChildItemObjectPathUsingIndexer=function (context, parent, childItem) {
			var id=childItem[OfficeExtension.Constants.id];
			if (OfficeExtension.Utility.isNullOrUndefined(id)) {
				id=childItem[OfficeExtension.Constants.idPrivate];
			}
			var objectPathInfo=objectPathInfo=				{
					Id: context._nextId(),
					ObjectPathType: 5,
					Name: "",
					ParentObjectPathId: parent._objectPath.objectPathInfo.Id,
					ArgumentInfo: {}
				};
			objectPathInfo.ArgumentInfo.Arguments=[id];
			return new OfficeExtension.ObjectPath(objectPathInfo, parent._objectPath, false, false);
		};
		ObjectPathFactory.createChildItemObjectPathUsingGetItemAt=function (context, parent, childItem, index) {
			var indexFromServer=childItem[OfficeExtension.Constants.index];
			if (indexFromServer) {
				index=indexFromServer;
			}
			var objectPathInfo={
				Id: context._nextId(),
				ObjectPathType: 3,
				Name: OfficeExtension.Constants.getItemAt,
				ParentObjectPathId: parent._objectPath.objectPathInfo.Id,
				ArgumentInfo: {}
			};
			objectPathInfo.ArgumentInfo.Arguments=[index];
			return new OfficeExtension.ObjectPath(objectPathInfo, parent._objectPath, false, false);
		};
		return ObjectPathFactory;
	}());
	OfficeExtension.ObjectPathFactory=ObjectPathFactory;
})(OfficeExtension || (OfficeExtension={}));
var OfficeExtension;
(function (OfficeExtension) {
	var OfficeJsRequestExecutor=(function () {
		function OfficeJsRequestExecutor(context) {
			this.m_context=context;
		}
		OfficeJsRequestExecutor.prototype.executeAsync=function (customData, requestFlags, requestMessage) {
			var _this=this;
			var messageSafearray=OfficeExtension.RichApiMessageUtility.buildMessageArrayForIRequestExecutor(customData, requestFlags, requestMessage, OfficeJsRequestExecutor.SourceLibHeaderValue);
			return new OfficeExtension._Internal.OfficePromise(function (resolve, reject) {
				OSF.DDA.RichApi.executeRichApiRequestAsync(messageSafearray, function (result) {
					OfficeExtension.Utility.log("Response:");
					OfficeExtension.Utility.log(JSON.stringify(result));
					var response;
					if (result.status=="succeeded") {
						response=OfficeExtension.RichApiMessageUtility.buildResponseOnSuccess(OfficeExtension.RichApiMessageUtility.getResponseBody(result), OfficeExtension.RichApiMessageUtility.getResponseHeaders(result));
					}
					else {
						response=OfficeExtension.RichApiMessageUtility.buildResponseOnError(result.error.code, result.error.message);
						_this.m_context._processOfficeJsErrorResponse(result.error.code, response);
					}
					resolve(response);
				});
			});
		};
		OfficeJsRequestExecutor.SourceLibHeaderValue="officejs";
		return OfficeJsRequestExecutor;
	}());
	OfficeExtension.OfficeJsRequestExecutor=OfficeJsRequestExecutor;
})(OfficeExtension || (OfficeExtension={}));
var OfficeExtension;
(function (OfficeExtension) {
	var _Internal;
	(function (_Internal) {
		var PromiseImpl;
		(function (PromiseImpl) {
			function Init() {
				return (function () {
					"use strict";
					function lib$es6$promise$utils$$objectOrFunction(x) {
						return typeof x==='function' || (typeof x==='object' && x !==null);
					}
					function lib$es6$promise$utils$$isFunction(x) {
						return typeof x==='function';
					}
					function lib$es6$promise$utils$$isMaybeThenable(x) {
						return typeof x==='object' && x !==null;
					}
					var lib$es6$promise$utils$$_isArray;
					if (!Array.isArray) {
						lib$es6$promise$utils$$_isArray=function (x) {
							return Object.prototype.toString.call(x)==='[object Array]';
						};
					}
					else {
						lib$es6$promise$utils$$_isArray=Array.isArray;
					}
					var lib$es6$promise$utils$$isArray=lib$es6$promise$utils$$_isArray;
					var lib$es6$promise$asap$$len=0;
					var lib$es6$promise$asap$$toString={}.toString;
					var lib$es6$promise$asap$$vertxNext;
					var lib$es6$promise$asap$$customSchedulerFn;
					var lib$es6$promise$asap$$asap=function asap(callback, arg) {
						lib$es6$promise$asap$$queue[lib$es6$promise$asap$$len]=callback;
						lib$es6$promise$asap$$queue[lib$es6$promise$asap$$len+1]=arg;
						lib$es6$promise$asap$$len+=2;
						if (lib$es6$promise$asap$$len===2) {
							if (lib$es6$promise$asap$$customSchedulerFn) {
								lib$es6$promise$asap$$customSchedulerFn(lib$es6$promise$asap$$flush);
							}
							else {
								lib$es6$promise$asap$$scheduleFlush();
							}
						}
					};
					function lib$es6$promise$asap$$setScheduler(scheduleFn) {
						lib$es6$promise$asap$$customSchedulerFn=scheduleFn;
					}
					function lib$es6$promise$asap$$setAsap(asapFn) {
						lib$es6$promise$asap$$asap=asapFn;
					}
					var lib$es6$promise$asap$$browserWindow=(typeof window !=='undefined') ? window : undefined;
					var lib$es6$promise$asap$$browserGlobal=lib$es6$promise$asap$$browserWindow || {};
					var lib$es6$promise$asap$$BrowserMutationObserver=lib$es6$promise$asap$$browserGlobal.MutationObserver || lib$es6$promise$asap$$browserGlobal.WebKitMutationObserver;
					var lib$es6$promise$asap$$isNode=typeof process !=='undefined' && {}.toString.call(process)==='[object process]';
					var lib$es6$promise$asap$$isWorker=typeof Uint8ClampedArray !=='undefined' &&
						typeof importScripts !=='undefined' &&
						typeof MessageChannel !=='undefined';
					function lib$es6$promise$asap$$useNextTick() {
						var nextTick=process.nextTick;
						var version=process.versions.node.match(/^(?:(\d+)\.)?(?:(\d+)\.)?(\*|\d+)$/);
						if (Array.isArray(version) && version[1]==='0' && version[2]==='10') {
							nextTick=setImmediate;
						}
						return function () {
							nextTick(lib$es6$promise$asap$$flush);
						};
					}
					function lib$es6$promise$asap$$useVertxTimer() {
						return function () {
							lib$es6$promise$asap$$vertxNext(lib$es6$promise$asap$$flush);
						};
					}
					function lib$es6$promise$asap$$useMutationObserver() {
						var iterations=0;
						var observer=new lib$es6$promise$asap$$BrowserMutationObserver(lib$es6$promise$asap$$flush);
						var node=document.createTextNode('');
						observer.observe(node, { characterData: true });
						return function () {
							node.data=(iterations=++iterations % 2);
						};
					}
					function lib$es6$promise$asap$$useMessageChannel() {
						var channel=new MessageChannel();
						channel.port1.onmessage=lib$es6$promise$asap$$flush;
						return function () {
							channel.port2.postMessage(0);
						};
					}
					function lib$es6$promise$asap$$useSetTimeout() {
						return function () {
							setTimeout(lib$es6$promise$asap$$flush, 1);
						};
					}
					var lib$es6$promise$asap$$queue=new Array(1000);
					function lib$es6$promise$asap$$flush() {
						for (var i=0; i < lib$es6$promise$asap$$len; i+=2) {
							var callback=lib$es6$promise$asap$$queue[i];
							var arg=lib$es6$promise$asap$$queue[i+1];
							callback(arg);
							lib$es6$promise$asap$$queue[i]=undefined;
							lib$es6$promise$asap$$queue[i+1]=undefined;
						}
						lib$es6$promise$asap$$len=0;
					}
					var lib$es6$promise$asap$$scheduleFlush;
					if (lib$es6$promise$asap$$isNode) {
						lib$es6$promise$asap$$scheduleFlush=lib$es6$promise$asap$$useNextTick();
					}
					else if (lib$es6$promise$asap$$BrowserMutationObserver) {
						lib$es6$promise$asap$$scheduleFlush=lib$es6$promise$asap$$useMutationObserver();
					}
					else if (lib$es6$promise$asap$$isWorker) {
						lib$es6$promise$asap$$scheduleFlush=lib$es6$promise$asap$$useMessageChannel();
					}
					else {
						lib$es6$promise$asap$$scheduleFlush=lib$es6$promise$asap$$useSetTimeout();
					}
					function lib$es6$promise$$internal$$noop() { }
					var lib$es6$promise$$internal$$PENDING=void 0;
					var lib$es6$promise$$internal$$FULFILLED=1;
					var lib$es6$promise$$internal$$REJECTED=2;
					var lib$es6$promise$$internal$$GET_THEN_ERROR=new lib$es6$promise$$internal$$ErrorObject();
					function lib$es6$promise$$internal$$selfFullfillment() {
						return new TypeError("You cannot resolve a promise with itself");
					}
					function lib$es6$promise$$internal$$cannotReturnOwn() {
						return new TypeError('A promises callback cannot return that same promise.');
					}
					function lib$es6$promise$$internal$$getThen(promise) {
						try {
							return promise.then;
						}
						catch (error) {
							lib$es6$promise$$internal$$GET_THEN_ERROR.error=error;
							return lib$es6$promise$$internal$$GET_THEN_ERROR;
						}
					}
					function lib$es6$promise$$internal$$tryThen(then, value, fulfillmentHandler, rejectionHandler) {
						try {
							then.call(value, fulfillmentHandler, rejectionHandler);
						}
						catch (e) {
							return e;
						}
					}
					function lib$es6$promise$$internal$$handleForeignThenable(promise, thenable, then) {
						lib$es6$promise$asap$$asap(function (promise) {
							var sealed=false;
							var error=lib$es6$promise$$internal$$tryThen(then, thenable, function (value) {
								if (sealed) {
									return;
								}
								sealed=true;
								if (thenable !==value) {
									lib$es6$promise$$internal$$resolve(promise, value);
								}
								else {
									lib$es6$promise$$internal$$fulfill(promise, value);
								}
							}, function (reason) {
								if (sealed) {
									return;
								}
								sealed=true;
								lib$es6$promise$$internal$$reject(promise, reason);
							}, 'Settle: '+(promise._label || ' unknown promise'));
							if (!sealed && error) {
								sealed=true;
								lib$es6$promise$$internal$$reject(promise, error);
							}
						}, promise);
					}
					function lib$es6$promise$$internal$$handleOwnThenable(promise, thenable) {
						if (thenable._state===lib$es6$promise$$internal$$FULFILLED) {
							lib$es6$promise$$internal$$fulfill(promise, thenable._result);
						}
						else if (thenable._state===lib$es6$promise$$internal$$REJECTED) {
							lib$es6$promise$$internal$$reject(promise, thenable._result);
						}
						else {
							lib$es6$promise$$internal$$subscribe(thenable, undefined, function (value) {
								lib$es6$promise$$internal$$resolve(promise, value);
							}, function (reason) {
								lib$es6$promise$$internal$$reject(promise, reason);
							});
						}
					}
					function lib$es6$promise$$internal$$handleMaybeThenable(promise, maybeThenable) {
						if (maybeThenable.constructor===promise.constructor) {
							lib$es6$promise$$internal$$handleOwnThenable(promise, maybeThenable);
						}
						else {
							var then=lib$es6$promise$$internal$$getThen(maybeThenable);
							if (then===lib$es6$promise$$internal$$GET_THEN_ERROR) {
								lib$es6$promise$$internal$$reject(promise, lib$es6$promise$$internal$$GET_THEN_ERROR.error);
							}
							else if (then===undefined) {
								lib$es6$promise$$internal$$fulfill(promise, maybeThenable);
							}
							else if (lib$es6$promise$utils$$isFunction(then)) {
								lib$es6$promise$$internal$$handleForeignThenable(promise, maybeThenable, then);
							}
							else {
								lib$es6$promise$$internal$$fulfill(promise, maybeThenable);
							}
						}
					}
					function lib$es6$promise$$internal$$resolve(promise, value) {
						if (promise===value) {
							lib$es6$promise$$internal$$reject(promise, lib$es6$promise$$internal$$selfFullfillment());
						}
						else if (lib$es6$promise$utils$$objectOrFunction(value)) {
							lib$es6$promise$$internal$$handleMaybeThenable(promise, value);
						}
						else {
							lib$es6$promise$$internal$$fulfill(promise, value);
						}
					}
					function lib$es6$promise$$internal$$publishRejection(promise) {
						if (promise._onerror) {
							promise._onerror(promise._result);
						}
						lib$es6$promise$$internal$$publish(promise);
					}
					function lib$es6$promise$$internal$$fulfill(promise, value) {
						if (promise._state !==lib$es6$promise$$internal$$PENDING) {
							return;
						}
						promise._result=value;
						promise._state=lib$es6$promise$$internal$$FULFILLED;
						if (promise._subscribers.length !==0) {
							lib$es6$promise$asap$$asap(lib$es6$promise$$internal$$publish, promise);
						}
					}
					function lib$es6$promise$$internal$$reject(promise, reason) {
						if (promise._state !==lib$es6$promise$$internal$$PENDING) {
							return;
						}
						promise._state=lib$es6$promise$$internal$$REJECTED;
						promise._result=reason;
						lib$es6$promise$asap$$asap(lib$es6$promise$$internal$$publishRejection, promise);
					}
					function lib$es6$promise$$internal$$subscribe(parent, child, onFulfillment, onRejection) {
						var subscribers=parent._subscribers;
						var length=subscribers.length;
						parent._onerror=null;
						subscribers[length]=child;
						subscribers[length+lib$es6$promise$$internal$$FULFILLED]=onFulfillment;
						subscribers[length+lib$es6$promise$$internal$$REJECTED]=onRejection;
						if (length===0 && parent._state) {
							lib$es6$promise$asap$$asap(lib$es6$promise$$internal$$publish, parent);
						}
					}
					function lib$es6$promise$$internal$$publish(promise) {
						var subscribers=promise._subscribers;
						var settled=promise._state;
						if (subscribers.length===0) {
							return;
						}
						var child, callback, detail=promise._result;
						for (var i=0; i < subscribers.length; i+=3) {
							child=subscribers[i];
							callback=subscribers[i+settled];
							if (child) {
								lib$es6$promise$$internal$$invokeCallback(settled, child, callback, detail);
							}
							else {
								callback(detail);
							}
						}
						promise._subscribers.length=0;
					}
					function lib$es6$promise$$internal$$ErrorObject() {
						this.error=null;
					}
					var lib$es6$promise$$internal$$TRY_CATCH_ERROR=new lib$es6$promise$$internal$$ErrorObject();
					function lib$es6$promise$$internal$$tryCatch(callback, detail) {
						try {
							return callback(detail);
						}
						catch (e) {
							lib$es6$promise$$internal$$TRY_CATCH_ERROR.error=e;
							return lib$es6$promise$$internal$$TRY_CATCH_ERROR;
						}
					}
					function lib$es6$promise$$internal$$invokeCallback(settled, promise, callback, detail) {
						var hasCallback=lib$es6$promise$utils$$isFunction(callback), value, error, succeeded, failed;
						if (hasCallback) {
							value=lib$es6$promise$$internal$$tryCatch(callback, detail);
							if (value===lib$es6$promise$$internal$$TRY_CATCH_ERROR) {
								failed=true;
								error=value.error;
								value=null;
							}
							else {
								succeeded=true;
							}
							if (promise===value) {
								lib$es6$promise$$internal$$reject(promise, lib$es6$promise$$internal$$cannotReturnOwn());
								return;
							}
						}
						else {
							value=detail;
							succeeded=true;
						}
						if (promise._state !==lib$es6$promise$$internal$$PENDING) {
						}
						else if (hasCallback && succeeded) {
							lib$es6$promise$$internal$$resolve(promise, value);
						}
						else if (failed) {
							lib$es6$promise$$internal$$reject(promise, error);
						}
						else if (settled===lib$es6$promise$$internal$$FULFILLED) {
							lib$es6$promise$$internal$$fulfill(promise, value);
						}
						else if (settled===lib$es6$promise$$internal$$REJECTED) {
							lib$es6$promise$$internal$$reject(promise, value);
						}
					}
					function lib$es6$promise$$internal$$initializePromise(promise, resolver) {
						try {
							resolver(function resolvePromise(value) {
								lib$es6$promise$$internal$$resolve(promise, value);
							}, function rejectPromise(reason) {
								lib$es6$promise$$internal$$reject(promise, reason);
							});
						}
						catch (e) {
							lib$es6$promise$$internal$$reject(promise, e);
						}
					}
					function lib$es6$promise$enumerator$$Enumerator(Constructor, input) {
						var enumerator=this;
						enumerator._instanceConstructor=Constructor;
						enumerator.promise=new Constructor(lib$es6$promise$$internal$$noop);
						if (enumerator._validateInput(input)) {
							enumerator._input=input;
							enumerator.length=input.length;
							enumerator._remaining=input.length;
							enumerator._init();
							if (enumerator.length===0) {
								lib$es6$promise$$internal$$fulfill(enumerator.promise, enumerator._result);
							}
							else {
								enumerator.length=enumerator.length || 0;
								enumerator._enumerate();
								if (enumerator._remaining===0) {
									lib$es6$promise$$internal$$fulfill(enumerator.promise, enumerator._result);
								}
							}
						}
						else {
							lib$es6$promise$$internal$$reject(enumerator.promise, enumerator._validationError());
						}
					}
					lib$es6$promise$enumerator$$Enumerator.prototype._validateInput=function (input) {
						return lib$es6$promise$utils$$isArray(input);
					};
					lib$es6$promise$enumerator$$Enumerator.prototype._validationError=function () {
						return new _Internal.Error('Array Methods must be provided an Array');
					};
					lib$es6$promise$enumerator$$Enumerator.prototype._init=function () {
						this._result=new Array(this.length);
					};
					var lib$es6$promise$enumerator$$default=lib$es6$promise$enumerator$$Enumerator;
					lib$es6$promise$enumerator$$Enumerator.prototype._enumerate=function () {
						var enumerator=this;
						var length=enumerator.length;
						var promise=enumerator.promise;
						var input=enumerator._input;
						for (var i=0; promise._state===lib$es6$promise$$internal$$PENDING && i < length; i++) {
							enumerator._eachEntry(input[i], i);
						}
					};
					lib$es6$promise$enumerator$$Enumerator.prototype._eachEntry=function (entry, i) {
						var enumerator=this;
						var c=enumerator._instanceConstructor;
						if (lib$es6$promise$utils$$isMaybeThenable(entry)) {
							if (entry.constructor===c && entry._state !==lib$es6$promise$$internal$$PENDING) {
								entry._onerror=null;
								enumerator._settledAt(entry._state, i, entry._result);
							}
							else {
								enumerator._willSettleAt(c.resolve(entry), i);
							}
						}
						else {
							enumerator._remaining--;
							enumerator._result[i]=entry;
						}
					};
					lib$es6$promise$enumerator$$Enumerator.prototype._settledAt=function (state, i, value) {
						var enumerator=this;
						var promise=enumerator.promise;
						if (promise._state===lib$es6$promise$$internal$$PENDING) {
							enumerator._remaining--;
							if (state===lib$es6$promise$$internal$$REJECTED) {
								lib$es6$promise$$internal$$reject(promise, value);
							}
							else {
								enumerator._result[i]=value;
							}
						}
						if (enumerator._remaining===0) {
							lib$es6$promise$$internal$$fulfill(promise, enumerator._result);
						}
					};
					lib$es6$promise$enumerator$$Enumerator.prototype._willSettleAt=function (promise, i) {
						var enumerator=this;
						lib$es6$promise$$internal$$subscribe(promise, undefined, function (value) {
							enumerator._settledAt(lib$es6$promise$$internal$$FULFILLED, i, value);
						}, function (reason) {
							enumerator._settledAt(lib$es6$promise$$internal$$REJECTED, i, reason);
						});
					};
					function lib$es6$promise$promise$all$$all(entries) {
						return new lib$es6$promise$enumerator$$default(this, entries).promise;
					}
					var lib$es6$promise$promise$all$$default=lib$es6$promise$promise$all$$all;
					function lib$es6$promise$promise$race$$race(entries) {
						var Constructor=this;
						var promise=new Constructor(lib$es6$promise$$internal$$noop);
						if (!lib$es6$promise$utils$$isArray(entries)) {
							lib$es6$promise$$internal$$reject(promise, new TypeError('You must pass an array to race.'));
							return promise;
						}
						var length=entries.length;
						function onFulfillment(value) {
							lib$es6$promise$$internal$$resolve(promise, value);
						}
						function onRejection(reason) {
							lib$es6$promise$$internal$$reject(promise, reason);
						}
						for (var i=0; promise._state===lib$es6$promise$$internal$$PENDING && i < length; i++) {
							lib$es6$promise$$internal$$subscribe(Constructor.resolve(entries[i]), undefined, onFulfillment, onRejection);
						}
						return promise;
					}
					var lib$es6$promise$promise$race$$default=lib$es6$promise$promise$race$$race;
					function lib$es6$promise$promise$resolve$$resolve(object) {
						var Constructor=this;
						if (object && typeof object==='object' && object.constructor===Constructor) {
							return object;
						}
						var promise=new Constructor(lib$es6$promise$$internal$$noop);
						lib$es6$promise$$internal$$resolve(promise, object);
						return promise;
					}
					var lib$es6$promise$promise$resolve$$default=lib$es6$promise$promise$resolve$$resolve;
					function lib$es6$promise$promise$reject$$reject(reason) {
						var Constructor=this;
						var promise=new Constructor(lib$es6$promise$$internal$$noop);
						lib$es6$promise$$internal$$reject(promise, reason);
						return promise;
					}
					var lib$es6$promise$promise$reject$$default=lib$es6$promise$promise$reject$$reject;
					var lib$es6$promise$promise$$counter=0;
					function lib$es6$promise$promise$$needsResolver() {
						throw new TypeError('You must pass a resolver function as the first argument to the promise constructor');
					}
					function lib$es6$promise$promise$$needsNew() {
						throw new TypeError("Failed to construct 'Promise': Please use the 'new' operator, this object constructor cannot be called as a function.");
					}
					var lib$es6$promise$promise$$default=lib$es6$promise$promise$$Promise;
					function lib$es6$promise$promise$$Promise(resolver) {
						this._id=lib$es6$promise$promise$$counter++;
						this._state=undefined;
						this._result=undefined;
						this._subscribers=[];
						if (lib$es6$promise$$internal$$noop !==resolver) {
							if (!lib$es6$promise$utils$$isFunction(resolver)) {
								lib$es6$promise$promise$$needsResolver();
							}
							if (!(this instanceof lib$es6$promise$promise$$Promise)) {
								lib$es6$promise$promise$$needsNew();
							}
							lib$es6$promise$$internal$$initializePromise(this, resolver);
						}
					}
					lib$es6$promise$promise$$Promise.all=lib$es6$promise$promise$all$$default;
					lib$es6$promise$promise$$Promise.race=lib$es6$promise$promise$race$$default;
					lib$es6$promise$promise$$Promise.resolve=lib$es6$promise$promise$resolve$$default;
					lib$es6$promise$promise$$Promise.reject=lib$es6$promise$promise$reject$$default;
					lib$es6$promise$promise$$Promise._setScheduler=lib$es6$promise$asap$$setScheduler;
					lib$es6$promise$promise$$Promise._setAsap=lib$es6$promise$asap$$setAsap;
					lib$es6$promise$promise$$Promise._asap=lib$es6$promise$asap$$asap;
					lib$es6$promise$promise$$Promise.prototype={
						constructor: lib$es6$promise$promise$$Promise,
						then: function (onFulfillment, onRejection) {
							var parent=this;
							var state=parent._state;
							if (state===lib$es6$promise$$internal$$FULFILLED && !onFulfillment || state===lib$es6$promise$$internal$$REJECTED && !onRejection) {
								return this;
							}
							var child=new this.constructor(lib$es6$promise$$internal$$noop);
							var result=parent._result;
							if (state) {
								var callback=arguments[state - 1];
								lib$es6$promise$asap$$asap(function () {
									lib$es6$promise$$internal$$invokeCallback(state, child, callback, result);
								});
							}
							else {
								lib$es6$promise$$internal$$subscribe(parent, child, onFulfillment, onRejection);
							}
							return child;
						},
						'catch': function (onRejection) {
							return this.then(null, onRejection);
						}
					};
					return lib$es6$promise$promise$$default;
				}).call(this);
			}
			PromiseImpl.Init=Init;
		})(PromiseImpl=_Internal.PromiseImpl || (_Internal.PromiseImpl={}));
	})(_Internal=OfficeExtension._Internal || (OfficeExtension._Internal={}));
	var _Internal;
	(function (_Internal) {
		function isEdgeLessThan14() {
			var userAgent=window.navigator.userAgent;
			var versionIdx=userAgent.indexOf("Edge/");
			if (versionIdx >=0) {
				userAgent=userAgent.substring(versionIdx+5, userAgent.length);
				if (userAgent < "14.14393")
					return true;
				else
					return false;
			}
			return false;
		}
		function determinePromise() {
			if (typeof (window)==="undefined" && typeof (Promise)==="function") {
				return Promise;
			}
			if (typeof (window) !=="undefined" && window.Promise) {
				if (isEdgeLessThan14()) {
					return _Internal.PromiseImpl.Init();
				}
				else {
					return window.Promise;
				}
			}
			else {
				return _Internal.PromiseImpl.Init();
			}
		}
		_Internal.OfficePromise=determinePromise();
	})(_Internal=OfficeExtension._Internal || (OfficeExtension._Internal={}));
	var OfficePromise=_Internal.OfficePromise;
	OfficeExtension.Promise=OfficePromise;
})(OfficeExtension || (OfficeExtension={}));
var OfficeExtension;
(function (OfficeExtension) {
	var TrackedObjects=(function () {
		function TrackedObjects(context) {
			this._autoCleanupList={};
			this.m_context=context;
		}
		TrackedObjects.prototype.add=function (param) {
			var _this=this;
			if (Array.isArray(param)) {
				param.forEach(function (item) { return _this._addCommon(item, true); });
			}
			else {
				this._addCommon(param, true);
			}
		};
		TrackedObjects.prototype._autoAdd=function (object) {
			this._addCommon(object, false);
			this._autoCleanupList[object._objectPath.objectPathInfo.Id]=object;
		};
		TrackedObjects.prototype._autoTrackIfNecessaryWhenHandleObjectResultValue=function (object, resultValue) {
			var shouldAutoTrack=(this.m_context._autoCleanup &&
				!object[OfficeExtension.Constants.isTracked] &&
				object !==this.m_context._rootObject &&
				resultValue &&
				!OfficeExtension.Utility.isNullOrEmptyString(resultValue[OfficeExtension.Constants.referenceId]));
			if (shouldAutoTrack) {
				this._autoCleanupList[object._objectPath.objectPathInfo.Id]=object;
				object[OfficeExtension.Constants.isTracked]=true;
			}
		};
		TrackedObjects.prototype._addCommon=function (object, isExplicitlyAdded) {
			if (object[OfficeExtension.Constants.isTracked]) {
				if (isExplicitlyAdded && this.m_context._autoCleanup) {
					delete this._autoCleanupList[object._objectPath.objectPathInfo.Id];
				}
				return;
			}
			var referenceId=object[OfficeExtension.Constants.referenceId];
			if (OfficeExtension.Utility.isNullOrEmptyString(referenceId) && object._KeepReference) {
				object._KeepReference();
				OfficeExtension.ActionFactory.createInstantiateAction(this.m_context, object);
				if (isExplicitlyAdded && this.m_context._autoCleanup) {
					delete this._autoCleanupList[object._objectPath.objectPathInfo.Id];
				}
				object[OfficeExtension.Constants.isTracked]=true;
			}
		};
		TrackedObjects.prototype.remove=function (param) {
			var _this=this;
			if (Array.isArray(param)) {
				param.forEach(function (item) { return _this._removeCommon(item); });
			}
			else {
				this._removeCommon(param);
			}
		};
		TrackedObjects.prototype._removeCommon=function (object) {
			var referenceId=object[OfficeExtension.Constants.referenceId];
			if (!OfficeExtension.Utility.isNullOrEmptyString(referenceId)) {
				var rootObject=this.m_context._rootObject;
				if (rootObject._RemoveReference) {
					rootObject._RemoveReference(referenceId);
				}
				delete object[OfficeExtension.Constants.isTracked];
			}
		};
		TrackedObjects.prototype._retrieveAndClearAutoCleanupList=function () {
			var list=this._autoCleanupList;
			this._autoCleanupList={};
			return list;
		};
		return TrackedObjects;
	}());
	OfficeExtension.TrackedObjects=TrackedObjects;
})(OfficeExtension || (OfficeExtension={}));
var OfficeExtension;
(function (OfficeExtension) {
	var ResourceStrings=(function () {
		function ResourceStrings() {
		}
		ResourceStrings.cannotRegisterEvent="CannotRegisterEvent";
		ResourceStrings.connectionFailureWithStatus="ConnectionFailureWithStatus";
		ResourceStrings.connectionFailureWithDetails="ConnectionFailureWithDetails";
		ResourceStrings.invalidObjectPath="InvalidObjectPath";
		ResourceStrings.invalidRequestContext="InvalidRequestContext";
		ResourceStrings.invalidArgument="InvalidArgument";
		ResourceStrings.invalidArgumentGeneric="InvalidArgumentGeneric";
		ResourceStrings.propertyNotLoaded="PropertyNotLoaded";
		ResourceStrings.runMustReturnPromise="RunMustReturnPromise";
		ResourceStrings.timeout="Timeout";
		ResourceStrings.propertyDoesNotExist="PropertyDoesNotExist";
		ResourceStrings.attemptingToSetReadOnlyProperty="AttemptingToSetReadOnlyProperty";
		ResourceStrings.moreInfoInnerError="MoreInfoInnerError";
		ResourceStrings.cannotApplyPropertyThroughSetMethod="CannotApplyPropertyThroughSetMethod";
		ResourceStrings.valueNotLoaded="ValueNotLoaded";
		ResourceStrings.invalidOrTimedOutSessionMessage="InvalidOrTimedOutSessionMessage";
		ResourceStrings.invalidOperationInCellEditMode="InvalidOperationInCellEditMode";
		ResourceStrings.customFunctionDefintionMissing="CustomFunctionDefintionMissing";
		ResourceStrings.customFunctionImplementationMissing="CustomFunctionImplementationMissing";
		ResourceStrings.apiNotFoundDetails="ApiNotFoundDetails";
		return ResourceStrings;
	}());
	OfficeExtension.ResourceStrings=ResourceStrings;
})(OfficeExtension || (OfficeExtension={}));
var OfficeExtension;
(function (OfficeExtension) {
	var ResourceStringValues=(function () {
		function ResourceStringValues() {
		}
		ResourceStringValues.CannotRegisterEvent="The event handler cannot be registered.";
		ResourceStringValues.ConnectionFailureWithStatus="The request failed with status code of {0}.";
		ResourceStringValues.ConnectionFailureWithDetails="The request failed with status code of {0}, error code {1} and the following error message: {2}";
		ResourceStringValues.InvalidArgument="The argument '{0}' doesn't work for this situation, is missing, or isn't in the right format.";
		ResourceStringValues.InvalidObjectPath="The object path '{0}' isn't working for what you're trying to do. If you're using the object across multiple \"context.sync\" calls and outside the sequential execution of a \".run\" batch, please use the \"context.trackedObjects.add()\" and \"context.trackedObjects.remove()\" methods to manage the object's lifetime.";
		ResourceStringValues.InvalidRequestContext="Cannot use the object across different request contexts.";
		ResourceStringValues.PropertyNotLoaded="The property '{0}' is not available. Before reading the property's value, call the load method on the containing object and call \"context.sync()\" on the associated request context.";
		ResourceStringValues.RunMustReturnPromise="The batch function passed to the \".run\" method didn't return a promise. The function must return a promise, so that any automatically-tracked objects can be released at the completion of the batch operation. Typically, you return a promise by returning the response from \"context.sync()\".";
		ResourceStringValues.Timeout="The operation has timed out.";
		ResourceStringValues.ValueNotLoaded="The value of the result object has not been loaded yet. Before reading the value property, call \"context.sync()\" on the associated request context.";
		ResourceStringValues.InvalidOrTimedOutSessionMessage="Your Office Online session has expired or is invalid. To continue, refresh the page.";
		ResourceStringValues.InvalidOperationInCellEditMode="Excel is in cell-editing mode. Please exit the edit mode by pressing ENTER or TAB or selecting another cell, and then try again.";
		ResourceStringValues.CustomFunctionDefintionMissing="A property with this name that represents the function's definition must exist on Excel.CustomFunctions";
		ResourceStringValues.CustomFunctionImplementationMissing="The property with this name on Excel.CustomFunctions that represents the function's definition must contain a 'call' property that implements the function.";
		ResourceStringValues.ApiNotFoundDetails="The method or property {0} is part of the {1} requirement set, which is not available in your version of {2}.";
		return ResourceStringValues;
	}());
	OfficeExtension.ResourceStringValues=ResourceStringValues;
})(OfficeExtension || (OfficeExtension={}));
var OfficeExtension;
(function (OfficeExtension) {
	var RichApiMessageUtility=(function () {
		function RichApiMessageUtility() {
		}
		RichApiMessageUtility.buildMessageArrayForIRequestExecutor=function (customData, requestFlags, requestMessage, sourceLibHeaderValue) {
			var requestMessageText=JSON.stringify(requestMessage.Body);
			OfficeExtension.Utility.log("Request:");
			OfficeExtension.Utility.log(requestMessageText);
			var headers={};
			headers[OfficeExtension.Constants.sourceLibHeader]=sourceLibHeaderValue;
			var messageSafearray=RichApiMessageUtility.buildRequestMessageSafeArray(customData, requestFlags, "POST", "ProcessQuery", headers, requestMessageText);
			return messageSafearray;
		};
		RichApiMessageUtility.buildResponseOnSuccess=function (responseBody, responseHeaders) {
			var response={ ErrorCode: '', ErrorMessage: '', Headers: null, Body: null };
			response.Body=JSON.parse(responseBody);
			response.Headers=responseHeaders;
			return response;
		};
		RichApiMessageUtility.buildResponseOnError=function (errorCode, message) {
			var response={ ErrorCode: '', ErrorMessage: '', Headers: null, Body: null };
			response.ErrorCode=OfficeExtension.ErrorCodes.generalException;
			response.ErrorMessage=message;
			if (errorCode==RichApiMessageUtility.OfficeJsErrorCode_ooeNoCapability) {
				response.ErrorCode=OfficeExtension.ErrorCodes.accessDenied;
			}
			else if (errorCode==RichApiMessageUtility.OfficeJsErrorCode_ooeActivityLimitReached) {
				response.ErrorCode=OfficeExtension.ErrorCodes.activityLimitReached;
			}
			else if (errorCode==RichApiMessageUtility.OfficeJsErrorCode_ooeInvalidOrTimedOutSession) {
				response.ErrorCode=OfficeExtension.ErrorCodes.invalidOrTimedOutSession;
				response.ErrorMessage=OfficeExtension.Utility._getResourceString(OfficeExtension.ResourceStrings.invalidOrTimedOutSessionMessage);
			}
			return response;
		};
		RichApiMessageUtility.buildHttpResponseFromOfficeJsError=function (errorCode, message) {
			var statusCode=500;
			var errorBody={};
			errorBody["error"]={};
			errorBody["error"]["code"]=OfficeExtension.ErrorCodes.generalException;
			errorBody["error"]["message"]=message;
			if (errorCode===RichApiMessageUtility.OfficeJsErrorCode_ooeNoCapability) {
				statusCode=403;
				errorBody["error"]["code"]=OfficeExtension.ErrorCodes.accessDenied;
			}
			else if (errorCode===RichApiMessageUtility.OfficeJsErrorCode_ooeActivityLimitReached) {
				statusCode=429;
				errorBody["error"]["code"]=OfficeExtension.ErrorCodes.activityLimitReached;
			}
			return { statusCode: statusCode, headers: {}, body: JSON.stringify(errorBody) };
		};
		RichApiMessageUtility.buildRequestMessageSafeArray=function (customData, requestFlags, method, path, headers, body) {
			var headerArray=[];
			if (headers) {
				for (var headerName in headers) {
					headerArray.push(headerName);
					headerArray.push(headers[headerName]);
				}
			}
			var appPermission=0;
			var solutionId="";
			var instanceId="";
			var marketplaceType="";
			return [
				customData,
				method,
				path,
				headerArray,
				body,
				appPermission,
				requestFlags,
				solutionId,
				instanceId,
				marketplaceType
			];
		};
		RichApiMessageUtility.getResponseBody=function (result) {
			return RichApiMessageUtility.getResponseBodyFromSafeArray(result.value.data);
		};
		RichApiMessageUtility.getResponseHeaders=function (result) {
			return RichApiMessageUtility.getResponseHeadersFromSafeArray(result.value.data);
		};
		RichApiMessageUtility.getResponseBodyFromSafeArray=function (data) {
			var ret=data[2];
			if (typeof (ret)==="string") {
				return ret;
			}
			var arr=ret;
			return arr.join("");
		};
		RichApiMessageUtility.getResponseHeadersFromSafeArray=function (data) {
			var arrayHeader=data[1];
			if (!arrayHeader) {
				return null;
			}
			var headers={};
			for (var i=0; i < arrayHeader.length - 1; i+=2) {
				headers[arrayHeader[i]]=arrayHeader[i+1];
			}
			return headers;
		};
		RichApiMessageUtility.getResponseStatusCode=function (result) {
			return RichApiMessageUtility.getResponseStatusCodeFromSafeArray(result.value.data);
		};
		RichApiMessageUtility.getResponseStatusCodeFromSafeArray=function (data) {
			return data[0];
		};
		RichApiMessageUtility.OfficeJsErrorCode_ooeInvalidOrTimedOutSession=5012;
		RichApiMessageUtility.OfficeJsErrorCode_ooeActivityLimitReached=5102;
		RichApiMessageUtility.OfficeJsErrorCode_ooeNoCapability=7000;
		return RichApiMessageUtility;
	}());
	OfficeExtension.RichApiMessageUtility=RichApiMessageUtility;
})(OfficeExtension || (OfficeExtension={}));
var OfficeExtension;
(function (OfficeExtension) {
	var Utility=(function () {
		function Utility() {
		}
		Utility.checkArgumentNull=function (value, name) {
			if (Utility.isNullOrUndefined(value)) {
				throw OfficeExtension._Internal.RuntimeError._createInvalidArgError(name);
			}
		};
		Utility.isNullOrUndefined=function (value) {
			if (value===null) {
				return true;
			}
			if (typeof (value)==="undefined") {
				return true;
			}
			return false;
		};
		Utility.isUndefined=function (value) {
			if (typeof (value)==="undefined") {
				return true;
			}
			return false;
		};
		Utility.isNullOrEmptyString=function (value) {
			if (value===null) {
				return true;
			}
			if (typeof (value)==="undefined") {
				return true;
			}
			if (value.length==0) {
				return true;
			}
			return false;
		};
		Utility.isPlainJsonObject=function (value) {
			if (Utility.isNullOrUndefined(value)) {
				return false;
			}
			if (typeof (value) !=="object") {
				return false;
			}
			return Object.getPrototypeOf(value)===Object.getPrototypeOf({});
		};
		Utility.trim=function (str) {
			return str.replace(new RegExp("^\\s+|\\s+$", "g"), "");
		};
		Utility.caseInsensitiveCompareString=function (str1, str2) {
			if (Utility.isNullOrUndefined(str1)) {
				return Utility.isNullOrUndefined(str2);
			}
			else {
				if (Utility.isNullOrUndefined(str2)) {
					return false;
				}
				else {
					return str1.toUpperCase()==str2.toUpperCase();
				}
			}
		};
		Utility.adjustToDateTime=function (value) {
			if (Utility.isNullOrUndefined(value)) {
				return null;
			}
			if (typeof (value)==="string") {
				return new Date(value);
			}
			if (Array.isArray(value)) {
				var arr=value;
				for (var i=0; i < arr.length; i++) {
					arr[i]=Utility.adjustToDateTime(arr[i]);
				}
				return arr;
			}
			throw OfficeExtension._Internal.RuntimeError._createInvalidArgError("date");
		};
		Utility.isReadonlyRestRequest=function (method) {
			return Utility.caseInsensitiveCompareString(method, "GET");
		};
		Utility.setMethodArguments=function (context, argumentInfo, args) {
			if (Utility.isNullOrUndefined(args)) {
				return null;
			}
			var referencedObjectPaths=new Array();
			var referencedObjectPathIds=new Array();
			var hasOne=Utility.collectObjectPathInfos(context, args, referencedObjectPaths, referencedObjectPathIds);
			argumentInfo.Arguments=args;
			if (hasOne) {
				argumentInfo.ReferencedObjectPathIds=referencedObjectPathIds;
				return referencedObjectPaths;
			}
			return null;
		};
		Utility.collectObjectPathInfos=function (context, args, referencedObjectPaths, referencedObjectPathIds) {
			var hasOne=false;
			for (var i=0; i < args.length; i++) {
				if (args[i] instanceof OfficeExtension.ClientObject) {
					var clientObject=args[i];
					Utility.validateContext(context, clientObject);
					args[i]=clientObject._objectPath.objectPathInfo.Id;
					referencedObjectPathIds.push(clientObject._objectPath.objectPathInfo.Id);
					referencedObjectPaths.push(clientObject._objectPath);
					hasOne=true;
				}
				else if (Array.isArray(args[i])) {
					var childArrayObjectPathIds=new Array();
					var childArrayHasOne=Utility.collectObjectPathInfos(context, args[i], referencedObjectPaths, childArrayObjectPathIds);
					if (childArrayHasOne) {
						referencedObjectPathIds.push(childArrayObjectPathIds);
						hasOne=true;
					}
					else {
						referencedObjectPathIds.push(0);
					}
				}
				else {
					referencedObjectPathIds.push(0);
				}
			}
			return hasOne;
		};
		Utility.fixObjectPathIfNecessary=function (clientObject, value) {
			if (clientObject && clientObject._objectPath && value) {
				clientObject._objectPath.updateUsingObjectData(value);
			}
		};
		Utility.validateObjectPath=function (clientObject) {
			var objectPath=clientObject._objectPath;
			while (objectPath) {
				if (!objectPath.isValid) {
					throw new OfficeExtension._Internal.RuntimeError({
						code: OfficeExtension.ErrorCodes.invalidObjectPath,
						message: Utility._getResourceString(OfficeExtension.ResourceStrings.invalidObjectPath, Utility.getObjectPathExpression(objectPath)),
						debugInfo: {
							errorLocation: Utility.getObjectPathExpression(objectPath)
						}
					});
				}
				objectPath=objectPath.parentObjectPath;
			}
		};
		Utility.validateReferencedObjectPaths=function (objectPaths) {
			if (objectPaths) {
				for (var i=0; i < objectPaths.length; i++) {
					var objectPath=objectPaths[i];
					while (objectPath) {
						if (!objectPath.isValid) {
							throw new OfficeExtension._Internal.RuntimeError({
								code: OfficeExtension.ErrorCodes.invalidObjectPath,
								message: Utility._getResourceString(OfficeExtension.ResourceStrings.invalidObjectPath, Utility.getObjectPathExpression(objectPath))
							});
						}
						objectPath=objectPath.parentObjectPath;
					}
				}
			}
		};
		Utility.validateContext=function (context, obj) {
			if (obj && obj.context !==context) {
				throw new OfficeExtension._Internal.RuntimeError({
					code: OfficeExtension.ErrorCodes.invalidRequestContext,
					message: Utility._getResourceString(OfficeExtension.ResourceStrings.invalidRequestContext)
				});
			}
		};
		Utility.log=function (message) {
			if (Utility._logEnabled && typeof (console) !=="undefined" && console.log) {
				console.log(message);
			}
		};
		Utility.load=function (clientObj, option) {
			clientObj.context.load(clientObj, option);
		};
		Utility._parseSelectExpand=function (select) {
			var args=[];
			if (!Utility.isNullOrEmptyString(select)) {
				var propertyNames=select.split(",");
				for (var i=0; i < propertyNames.length; i++) {
					var propertyName=propertyNames[i];
					propertyName=sanitizeForAnyItemsSlash(propertyName.trim());
					if (propertyName.length > 0) {
						args.push(propertyName);
					}
				}
			}
			return args;
			function sanitizeForAnyItemsSlash(propertyName) {
				var propertyNameLower=propertyName.toLowerCase();
				if (propertyNameLower==="items" || propertyNameLower==="items/") {
					return '*';
				}
				var itemsSlashLength=6;
				if (propertyNameLower.substr(0, itemsSlashLength)==="items/") {
					propertyName=propertyName.substr(itemsSlashLength);
				}
				return propertyName.replace(new RegExp("\/items\/", "gi"), "/");
			}
		};
		Utility.throwError=function (resourceId, arg, errorLocation) {
			throw new OfficeExtension._Internal.RuntimeError({
				code: resourceId,
				message: Utility._getResourceString(resourceId, arg),
				debugInfo: errorLocation ? { errorLocation: errorLocation } : undefined
			});
		};
		Utility.createRuntimeError=function (code, message, location) {
			return (new OfficeExtension._Internal.RuntimeError({
				code: code,
				message: message,
				debugInfo: { errorLocation: location }
			}));
		};
		Utility._getResourceString=function (resourceId, arg) {
			var ret;
			if (typeof (window) !=="undefined" && window.Strings && window.Strings.OfficeOM) {
				var stringName="L_"+resourceId;
				var stringValue=window.Strings.OfficeOM[stringName];
				if (stringValue) {
					ret=stringValue;
				}
			}
			if (!ret) {
				ret=OfficeExtension.ResourceStringValues[resourceId];
			}
			if (!ret) {
				ret=resourceId;
			}
			if (!Utility.isNullOrUndefined(arg)) {
				if (Array.isArray(arg)) {
					var arrArg=arg;
					ret=Utility._formatString(ret, arrArg);
				}
				else {
					ret=ret.replace("{0}", arg);
				}
			}
			return ret;
		};
		Utility._formatString=function (format, arrArg) {
			return format.replace(/\{\d\}/g, function (v) {
				var position=parseInt(v.substr(1, v.length - 2));
				if (position < arrArg.length) {
					return arrArg[position];
				}
				else {
					throw OfficeExtension._Internal.RuntimeError._createInvalidArgError("format");
				}
			});
		};
		Utility.throwIfNotLoaded=function (propertyName, fieldValue, entityName, isNull) {
			if (!isNull && Utility.isUndefined(fieldValue) && propertyName.charCodeAt(0) !=Utility.s_underscoreCharCode) {
				throw new OfficeExtension._Internal.RuntimeError({
					code: OfficeExtension.ErrorCodes.propertyNotLoaded,
					message: Utility._getResourceString(OfficeExtension.ResourceStrings.propertyNotLoaded, propertyName),
					debugInfo: entityName ? { errorLocation: entityName+"."+propertyName } : undefined
				});
			}
		};
		Utility.throwIfApiNotSupported=function (apiFullName, apiSetName, apiSetVersion, hostName) {
			if (!Utility._doApiNotSupportedCheck) {
				return;
			}
			if (typeof (window) !=="undefined" && window.Office && window.Office.context) {
				if (!window.Office.context.requirements.isSetSupported(apiSetName, apiSetVersion)) {
					var message=Utility._getResourceString(OfficeExtension.ResourceStrings.apiNotFoundDetails, [apiFullName, apiSetName+" "+apiSetVersion, hostName]);
					throw new OfficeExtension._Internal.RuntimeError({
						code: OfficeExtension.ErrorCodes.apiNotFound,
						message: message,
						debugInfo: { errorLocation: apiFullName }
					});
				}
			}
		};
		Utility.getObjectPathExpression=function (objectPath) {
			var ret="";
			while (objectPath) {
				switch (objectPath.objectPathInfo.ObjectPathType) {
					case 1:
						ret=ret;
						break;
					case 2:
						ret="new()"+(ret.length > 0 ? "." : "")+ret;
						break;
					case 3:
						ret=Utility.normalizeName(objectPath.objectPathInfo.Name)+"()"+(ret.length > 0 ? "." : "")+ret;
						break;
					case 4:
						ret=Utility.normalizeName(objectPath.objectPathInfo.Name)+(ret.length > 0 ? "." : "")+ret;
						break;
					case 5:
						ret="getItem()"+(ret.length > 0 ? "." : "")+ret;
						break;
					case 6:
						ret="_reference()"+(ret.length > 0 ? "." : "")+ret;
						break;
				}
				objectPath=objectPath.parentObjectPath;
			}
			return ret;
		};
		Utility._createPromiseFromResult=function (value) {
			return new OfficeExtension._Internal.OfficePromise(function (resolve, reject) {
				resolve(value);
			});
		};
		Utility._createTimeoutPromise=function (timeout) {
			return new OfficeExtension._Internal.OfficePromise(function (resolve, reject) {
				setTimeout(function () {
					resolve(null);
				}, timeout);
			});
		};
		Utility.promisify=function (action) {
			return new OfficeExtension._Internal.OfficePromise(function (resolve, reject) {
				var callback=function (result) {
					if (result.status=="failed") {
						reject(result.error);
					}
					else {
						resolve(result.value);
					}
				};
				action(callback);
			});
		};
		Utility._addActionResultHandler=function (clientObj, action, resultHandler) {
			clientObj.context._pendingRequest.addActionResultHandler(action, resultHandler);
		};
		Utility._handleNavigationPropertyResults=function (clientObj, objectValue, propertyNames) {
			for (var i=0; i < propertyNames.length - 1; i+=2) {
				if (!Utility.isUndefined(objectValue[propertyNames[i+1]])) {
					clientObj[propertyNames[i]]._handleResult(objectValue[propertyNames[i+1]]);
				}
			}
		};
		Utility.normalizeName=function (name) {
			return name.substr(0, 1).toLowerCase()+name.substr(1);
		};
		Utility._isLocalDocumentUrl=function (url) {
			return Utility._getLocalDocumentUrlPrefixLength(url) > 0;
		};
		Utility._getLocalDocumentUrlPrefixLength=function (url) {
			var localDocumentPrefixes=["http://document.localhost", "https://document.localhost", "//document.localhost"];
			var urlLower=url.toLowerCase().trim();
			for (var i=0; i < localDocumentPrefixes.length; i++) {
				if (urlLower===localDocumentPrefixes[i]) {
					return localDocumentPrefixes[i].length;
				}
				else if (urlLower.substr(0, localDocumentPrefixes[i].length+1)===localDocumentPrefixes[i]+"/") {
					return localDocumentPrefixes[i].length+1;
				}
			}
			return 0;
		};
		Utility._validateLocalDocumentRequest=function (request) {
			var index=Utility._getLocalDocumentUrlPrefixLength(request.url);
			if (index <=0) {
				throw OfficeExtension._Internal.RuntimeError._createInvalidArgError({
					argumentName: "request"
				});
			}
			var path=request.url.substr(index);
			var pathLower=path.toLowerCase();
			if (pathLower==="_api") {
				path="";
			}
			else if (pathLower.substr(0, "_api/".length)==="_api/") {
				path=path.substr("_api/".length);
			}
			return {
				method: request.method,
				url: path,
				headers: request.headers,
				body: request.body
			};
		};
		Utility._buildRequestMessageSafeArray=function (request) {
			var requestFlags=0;
			if (!Utility.isReadonlyRestRequest(request.method)) {
				requestFlags=1;
			}
			if (request.url.substr(0, OfficeExtension.Constants.processQuery.length).toLowerCase()===OfficeExtension.Constants.processQuery.toLowerCase()) {
				var index=request.url.indexOf("?");
				if (index > 0) {
					var queryString=request.url.substr(index+1);
					var parts=queryString.split("&");
					for (var i=0; i < parts.length; i++) {
						var keyvalue=parts[i].split("=");
						if (keyvalue[0].toLowerCase()===OfficeExtension.Constants.flags) {
							var flags=parseInt(keyvalue[1]);
							requestFlags=flags;
							requestFlags=requestFlags & 1;
							break;
						}
					}
				}
			}
			return OfficeExtension.RichApiMessageUtility.buildRequestMessageSafeArray("", requestFlags, request.method, request.url, request.headers, request.body);
		};
		Utility._parseHttpResponseHeaders=function (allResponseHeaders) {
			var responseHeaders={};
			if (!Utility.isNullOrEmptyString(allResponseHeaders)) {
				var regex=new RegExp("\r?\n");
				var entries=allResponseHeaders.split(regex);
				for (var i=0; i < entries.length; i++) {
					var entry=entries[i];
					if (entry !=null) {
						var index=entry.indexOf(':');
						if (index > 0) {
							var key=entry.substr(0, index);
							var value=entry.substr(index+1);
							key=Utility.trim(key);
							value=Utility.trim(value);
							responseHeaders[key.toUpperCase()]=value;
						}
					}
				}
			}
			return responseHeaders;
		};
		Utility._parseErrorResponse=function (responseInfo) {
			var errorObj=null;
			if (!Utility.isNullOrEmptyString(responseInfo.body)) {
				var errorResponseBody=Utility.trim(responseInfo.body);
				try {
					errorObj=JSON.parse(errorResponseBody);
				}
				catch (e) {
					Utility.log("Error when parse "+errorResponseBody);
				}
			}
			var errorMessage;
			var errorCode;
			if (!Utility.isNullOrUndefined(errorObj) && typeof (errorObj)==="object" && errorObj.error) {
				errorCode=errorObj.error.code;
				errorMessage=Utility._getResourceString(OfficeExtension.ResourceStrings.connectionFailureWithDetails, [responseInfo.statusCode.toString(), errorObj.error.code, errorObj.error.message]);
			}
			else {
				errorMessage=Utility._getResourceString(OfficeExtension.ResourceStrings.connectionFailureWithStatus, responseInfo.statusCode.toString());
			}
			if (Utility.isNullOrEmptyString(errorCode)) {
				errorCode=OfficeExtension.ErrorCodes.connectionFailure;
			}
			return { errorCode: errorCode, errorMessage: errorMessage };
		};
		Utility._copyHeaders=function (src, dest) {
			if (src && dest) {
				for (var key in src) {
					dest[key]=src[key];
				}
			}
		};
		Utility._logEnabled=false;
		Utility._synchronousCleanup=false;
		Utility._doApiNotSupportedCheck=false;
		Utility.s_underscoreCharCode="_".charCodeAt(0);
		return Utility;
	}());
	OfficeExtension.Utility=Utility;
})(OfficeExtension || (OfficeExtension={}));

var __extends=(this && this.__extends) || function (d, b) {
	for (var p in b) if (b.hasOwnProperty(p)) d[p]=b[p];
	function __() { this.constructor=d; }
	d.prototype=b===null ? Object.create(b) : (__.prototype=b.prototype, new __());
};
var OfficeCore;
(function (OfficeCore) {
	var _hostName="OfficeCore";
	var _defaultApiSetName="ExperimentApi";
	var _createPropertyObjectPath=OfficeExtension.ObjectPathFactory.createPropertyObjectPath;
	var _createMethodObjectPath=OfficeExtension.ObjectPathFactory.createMethodObjectPath;
	var _createIndexerObjectPath=OfficeExtension.ObjectPathFactory.createIndexerObjectPath;
	var _createNewObjectObjectPath=OfficeExtension.ObjectPathFactory.createNewObjectObjectPath;
	var _createChildItemObjectPathUsingIndexer=OfficeExtension.ObjectPathFactory.createChildItemObjectPathUsingIndexer;
	var _createChildItemObjectPathUsingGetItemAt=OfficeExtension.ObjectPathFactory.createChildItemObjectPathUsingGetItemAt;
	var _createChildItemObjectPathUsingIndexerOrGetItemAt=OfficeExtension.ObjectPathFactory.createChildItemObjectPathUsingIndexerOrGetItemAt;
	var _createMethodAction=OfficeExtension.ActionFactory.createMethodAction;
	var _createSetPropertyAction=OfficeExtension.ActionFactory.createSetPropertyAction;
	var _isNullOrUndefined=OfficeExtension.Utility.isNullOrUndefined;
	var _isUndefined=OfficeExtension.Utility.isUndefined;
	var _throwIfNotLoaded=OfficeExtension.Utility.throwIfNotLoaded;
	var _throwIfApiNotSupported=OfficeExtension.Utility.throwIfApiNotSupported;
	var _load=OfficeExtension.Utility.load;
	var _fixObjectPathIfNecessary=OfficeExtension.Utility.fixObjectPathIfNecessary;
	var _addActionResultHandler=OfficeExtension.Utility._addActionResultHandler;
	var _handleNavigationPropertyResults=OfficeExtension.Utility._handleNavigationPropertyResults;
	var _adjustToDateTime=OfficeExtension.Utility.adjustToDateTime;
	var FlightingService=(function (_super) {
		__extends(FlightingService, _super);
		function FlightingService() {
			_super.apply(this, arguments);
		}
		Object.defineProperty(FlightingService.prototype, "_className", {
			get: function () {
				return "FlightingService";
			},
			enumerable: true,
			configurable: true
		});
		FlightingService.prototype.getClientSessionId=function () {
			var action=_createMethodAction(this.context, this, "GetClientSessionId", 0, []);
			var ret=new OfficeExtension.ClientResult();
			_addActionResultHandler(this, action, ret);
			return ret;
		};
		FlightingService.prototype.getDeferredFlights=function () {
			var action=_createMethodAction(this.context, this, "GetDeferredFlights", 0, []);
			var ret=new OfficeExtension.ClientResult();
			_addActionResultHandler(this, action, ret);
			return ret;
		};
		FlightingService.prototype.getFeature=function (featureName, type, defaultValue, possibleValues) {
			return new OfficeCore.ABType(this.context, _createMethodObjectPath(this.context, this, "GetFeature", 0, [featureName, type, defaultValue, possibleValues], false, false, null));
		};
		FlightingService.prototype.getFeatureGate=function (featureName, scope) {
			return new OfficeCore.ABType(this.context, _createMethodObjectPath(this.context, this, "GetFeatureGate", 0, [featureName, scope], false, false, null));
		};
		FlightingService.prototype.resetOverride=function (featureName) {
			_createMethodAction(this.context, this, "ResetOverride", 0, [featureName]);
		};
		FlightingService.prototype.setOverride=function (featureName, type, value) {
			_createMethodAction(this.context, this, "SetOverride", 0, [featureName, type, value]);
		};
		FlightingService.prototype._handleResult=function (value) {
			_super.prototype._handleResult.call(this, value);
			if (_isNullOrUndefined(value))
				return;
			var obj=value;
			_fixObjectPathIfNecessary(this, obj);
		};
		FlightingService.newObject=function (context) {
			var ret=new OfficeCore.FlightingService(context, _createNewObjectObjectPath(context, "Microsoft.Experiment.FlightingService", false));
			return ret;
		};
		FlightingService.prototype.toJSON=function () {
			return {};
		};
		return FlightingService;
	}(OfficeExtension.ClientObject));
	OfficeCore.FlightingService=FlightingService;
	var ABType=(function (_super) {
		__extends(ABType, _super);
		function ABType() {
			_super.apply(this, arguments);
		}
		Object.defineProperty(ABType.prototype, "_className", {
			get: function () {
				return "ABType";
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ABType.prototype, "value", {
			get: function () {
				_throwIfNotLoaded("value", this.m_value, "ABType", this._isNull);
				return this.m_value;
			},
			enumerable: true,
			configurable: true
		});
		ABType.prototype._handleResult=function (value) {
			_super.prototype._handleResult.call(this, value);
			if (_isNullOrUndefined(value))
				return;
			var obj=value;
			_fixObjectPathIfNecessary(this, obj);
			if (!_isUndefined(obj["Value"])) {
				this.m_value=obj["Value"];
			}
		};
		ABType.prototype.load=function (option) {
			_load(this, option);
			return this;
		};
		ABType.prototype.toJSON=function () {
			return {
				"value": this.m_value
			};
		};
		return ABType;
	}(OfficeExtension.ClientObject));
	OfficeCore.ABType=ABType;
	var FeatureType;
	(function (FeatureType) {
		FeatureType.boolean="Boolean";
		FeatureType.integer="Integer";
		FeatureType.string="String";
	})(FeatureType=OfficeCore.FeatureType || (OfficeCore.FeatureType={}));
	var ExperimentErrorCodes;
	(function (ExperimentErrorCodes) {
		ExperimentErrorCodes.generalException="GeneralException";
	})(ExperimentErrorCodes=OfficeCore.ExperimentErrorCodes || (OfficeCore.ExperimentErrorCodes={}));
})(OfficeCore || (OfficeCore={}));
var OfficeCore;
(function (OfficeCore) {
	var RequestContext=(function (_super) {
		__extends(RequestContext, _super);
		function RequestContext(url) {
			_super.call(this, url);
		}
		Object.defineProperty(RequestContext.prototype, "flightingService", {
			get: function () {
				if (!this.m_flightingService) {
					this.m_flightingService=OfficeCore.FlightingService.newObject(this);
				}
				return this.m_flightingService;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(RequestContext.prototype, "flighting", {
			get: function () {
				return this.flightingService;
			},
			enumerable: true,
			configurable: true
		});
		return RequestContext;
	}(OfficeExtension.ClientRequestContext));
	OfficeCore.RequestContext=RequestContext;
})(OfficeCore || (OfficeCore={}));

var __extends=(this && this.__extends) || (function () {
	var extendStatics=Object.setPrototypeOf ||
		({ __proto__: [] } instanceof Array && function (d, b) { d.__proto__=b; }) ||
		function (d, b) { for (var p in b) if (b.hasOwnProperty(p)) d[p]=b[p]; };
	return function (d, b) {
		extendStatics(d, b);
		function __() { this.constructor=d; }
		d.prototype=b===null ? Object.create(b) : (__.prototype=b.prototype, new __());
	};
})();
var Excel;
(function (Excel) {
	function lowerCaseFirst(str) {
		return str[0].toLowerCase()+str.slice(1);
	}
	var iconSets=["ThreeArrows",
		"ThreeArrowsGray",
		"ThreeFlags",
		"ThreeTrafficLights1",
		"ThreeTrafficLights2",
		"ThreeSigns",
		"ThreeSymbols",
		"ThreeSymbols2",
		"FourArrows",
		"FourArrowsGray",
		"FourRedToBlack",
		"FourRating",
		"FourTrafficLights",
		"FiveArrows",
		"FiveArrowsGray",
		"FiveRating",
		"FiveQuarters",
		"ThreeStars",
		"ThreeTriangles",
		"FiveBoxes"];
	var iconNames=[["RedDownArrow", "YellowSideArrow", "GreenUpArrow"],
		["GrayDownArrow", "GraySideArrow", "GrayUpArrow"],
		["RedFlag", "YellowFlag", "GreenFlag"],
		["RedCircleWithBorder", "YellowCircle", "GreenCircle"],
		["RedTrafficLight", "YellowTrafficLight", "GreenTrafficLight"],
		["RedDiamond", "YellowTriangle", "GreenCircle"],
		["RedCrossSymbol", "YellowExclamationSymbol", "GreenCheckSymbol"],
		["RedCross", "YellowExclamation", "GreenCheck"],
		["RedDownArrow", "YellowDownInclineArrow", "YellowUpInclineArrow", "GreenUpArrow"],
		["GrayDownArrow", "GrayDownInclineArrow", "GrayUpInclineArrow", "GrayUpArrow"],
		["BlackCircle", "GrayCircle", "PinkCircle", "RedCircle"],
		["OneBar", "TwoBars", "ThreeBars", "FourBars"],
		["BlackCircleWithBorder", "RedCircleWithBorder", "YellowCircle", "GreenCircle"],
		["RedDownArrow", "YellowDownInclineArrow", "YellowSideArrow", "YellowUpInclineArrow", "GreenUpArrow"],
		["GrayDownArrow", "GrayDownInclineArrow", "GraySideArrow", "GrayUpInclineArrow", "GrayUpArrow"],
		["NoBars", "OneBar", "TwoBars", "ThreeBars", "FourBars"],
		["WhiteCircleAllWhiteQuarters", "CircleWithThreeWhiteQuarters", "CircleWithTwoWhiteQuarters", "CircleWithOneWhiteQuarter", "BlackCircle"],
		["SilverStar", "HalfGoldStar", "GoldStar"],
		["RedDownTriangle", "YellowDash", "GreenUpTriangle"],
		["NoFilledBoxes", "OneFilledBox", "TwoFilledBoxes", "ThreeFilledBoxes", "FourFilledBoxes"],];
	Excel.icons={};
	iconSets.map(function (title, i) {
		var camelTitle=lowerCaseFirst(title);
		Excel.icons[camelTitle]=[];
		iconNames[i].map(function (iconName, j) {
			iconName=lowerCaseFirst(iconName);
			var obj={ set: title, index: j };
			Excel.icons[camelTitle].push(obj);
			Excel.icons[camelTitle][iconName]=obj;
		});
	});
	function setRangePropertiesInBulk(range, propertyName, values) {
		var maxCellCount=1500;
		if (Array.isArray(values) && values.length > 0 && Array.isArray(values[0]) && (values.length * values[0].length > maxCellCount) && isExcel1_3OrAbove()) {
			var maxRowCount=Math.max(1, Math.round(maxCellCount / values[0].length));
			range._ValidateArraySize(values.length, values[0].length);
			for (var startRowIndex=0; startRowIndex < values.length; startRowIndex+=maxRowCount) {
				var rowCount=maxRowCount;
				if (startRowIndex+rowCount > values.length) {
					rowCount=values.length - startRowIndex;
				}
				var chunk=range.getRow(startRowIndex).getBoundingRect(range.getRow(startRowIndex+rowCount - 1));
				var valueSlice=values.slice(startRowIndex, startRowIndex+rowCount);
				_createSetPropertyAction(chunk.context, chunk, propertyName, valueSlice);
			}
			return true;
		}
		return false;
	}
	function isExcel1_3OrAbove() {
		if (typeof (window) !=="undefined" && window.Office && window.Office.context && window.Office.context.requirements) {
			return window.Office.context.requirements.isSetSupported("ExcelApi", 1.3);
		}
		else {
			return true;
		}
	}
	var Session=(function () {
		function Session(workbookUrl, requestHeaders, persisted) {
			this.m_workbookUrl=workbookUrl;
			this.m_requestHeaders=requestHeaders;
			if (!this.m_requestHeaders) {
				this.m_requestHeaders={};
			}
			if (OfficeExtension.Utility.isNullOrUndefined(persisted)) {
				persisted=true;
			}
			this.m_persisted=persisted;
		}
		Session.prototype.close=function () {
			var _this=this;
			if (this.m_requestUrlAndHeaderInfo &&
				!OfficeExtension.Utility._isLocalDocumentUrl(this.m_requestUrlAndHeaderInfo.url)) {
				var url=this.m_requestUrlAndHeaderInfo.url;
				if (url.charAt(url.length - 1) !="/") {
					url=url+"/";
				}
				url=url+"closeSession";
				var headers=this.m_requestUrlAndHeaderInfo;
				var req={ method: "POST", url: url, headers: this.m_requestUrlAndHeaderInfo.headers, body: "" };
				this.m_requestUrlAndHeaderInfo=null;
				return OfficeExtension.HttpUtility.sendRequest(req)
					.then(function (resp) {
					if (resp.statusCode !=204) {
						var err=OfficeExtension.Utility._parseErrorResponse(resp);
						throw OfficeExtension.Utility.createRuntimeError(err.errorCode, err.errorMessage, "Session.close");
					}
					_this.m_requestUrlAndHeaderInfo=null;
					var foundSessionKey=null;
					for (var key in _this.m_requestHeaders) {
						if (key.toLowerCase()==Session.WorkbookSessionIdHeaderNameLower) {
							foundSessionKey=key;
							break;
						}
					}
					if (foundSessionKey) {
						delete _this.m_requestHeaders[foundSessionKey];
					}
				});
			}
			else {
				return OfficeExtension.Utility._createPromiseFromResult(null);
			}
		};
		Session.prototype._resolveRequestUrlAndHeaderInfo=function () {
			var _this=this;
			if (this.m_requestUrlAndHeaderInfo) {
				return OfficeExtension.Utility._createPromiseFromResult(this.m_requestUrlAndHeaderInfo);
			}
			if (OfficeExtension.Utility.isNullOrEmptyString(this.m_workbookUrl) ||
				OfficeExtension.Utility._isLocalDocumentUrl(this.m_workbookUrl)) {
				this.m_requestUrlAndHeaderInfo={ url: this.m_workbookUrl, headers: this.m_requestHeaders };
				return OfficeExtension.Utility._createPromiseFromResult(this.m_requestUrlAndHeaderInfo);
			}
			var foundSessionId=false;
			for (var key in this.m_requestHeaders) {
				if (key.toLowerCase()==Session.WorkbookSessionIdHeaderNameLower) {
					foundSessionId=true;
					break;
				}
			}
			if (foundSessionId) {
				this.m_requestUrlAndHeaderInfo={ url: this.m_workbookUrl, headers: this.m_requestHeaders };
				return OfficeExtension.Utility._createPromiseFromResult(this.m_requestUrlAndHeaderInfo);
			}
			var url=this.m_workbookUrl;
			if (url.charAt(url.length - 1) !="/") {
				url=url+"/";
			}
			url=url+"createSession";
			var headers={};
			OfficeExtension.Utility._copyHeaders(this.m_requestHeaders, headers);
			headers["Content-Type"]="application/json";
			var body={};
			body.persistChanges=this.m_persisted;
			var req={ method: "POST", url: url, headers: headers, body: JSON.stringify(body) };
			return OfficeExtension.HttpUtility.sendRequest(req)
				.then(function (resp) {
				if (resp.statusCode !==201) {
					var err=OfficeExtension.Utility._parseErrorResponse(resp);
					throw OfficeExtension.Utility.createRuntimeError(err.errorCode, err.errorMessage, "Session.resolveRequestUrlAndHeaderInfo");
				}
				var session=JSON.parse(resp.body);
				var sessionId=session.id;
				headers={};
				OfficeExtension.Utility._copyHeaders(_this.m_requestHeaders, headers);
				headers[Session.WorkbookSessionIdHeaderName]=sessionId;
				_this.m_requestUrlAndHeaderInfo={ url: _this.m_workbookUrl, headers: headers };
				return _this.m_requestUrlAndHeaderInfo;
			});
		};
		return Session;
	}());
	Session.WorkbookSessionIdHeaderName="Workbook-Session-Id";
	Session.WorkbookSessionIdHeaderNameLower="workbook-session-id";
	Excel.Session=Session;
	var RequestContext=(function (_super) {
		__extends(RequestContext, _super);
		function RequestContext(url) {
			var _this=_super.call(this, url) || this;
			_this.m_workbook=new Workbook(_this, OfficeExtension.ObjectPathFactory.createGlobalObjectObjectPath(_this));
			_this._rootObject=_this.m_workbook;
			return _this;
		}
		RequestContext.prototype._processOfficeJsErrorResponse=function (officeJsErrorCode, response) {
			var ooeInvalidApiCallInContext=5004;
			if (officeJsErrorCode==ooeInvalidApiCallInContext) {
				response.ErrorCode=ErrorCodes.invalidOperationInCellEditMode;
				response.ErrorMessage=OfficeExtension.Utility._getResourceString(OfficeExtension.ResourceStrings.invalidOperationInCellEditMode);
			}
		};
		Object.defineProperty(RequestContext.prototype, "workbook", {
			get: function () {
				return this.m_workbook;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(RequestContext.prototype, "application", {
			get: function () {
				return this.workbook.application;
			},
			enumerable: true,
			configurable: true
		});
		return RequestContext;
	}(OfficeCore.RequestContext));
	Excel.RequestContext=RequestContext;
	function run(arg1, arg2, arg3) {
		return OfficeExtension.ClientRequestContext._runBatch("Excel.run", arguments, function (requestInfo) {
			var ret=new Excel.RequestContext(requestInfo);
			return ret;
		});
	}
	Excel.run=run;
	Excel._RedirectV1APIs=false;
	Excel._V1APIMap={
		"GetDataAsync": {
			call: function (ctx, callArgs) { return ctx.workbook._V1Api.bindingGetData(callArgs); },
			postprocess: getDataCommonPostprocess
		},
		"GetSelectedDataAsync": {
			call: function (ctx, callArgs) { return ctx.workbook._V1Api.getSelectedData(callArgs); },
			postprocess: getDataCommonPostprocess
		},
		"GoToByIdAsync": {
			call: function (ctx, callArgs) { return ctx.workbook._V1Api.gotoById(callArgs); }
		},
		"AddColumnsAsync": {
			call: function (ctx, callArgs) { return ctx.workbook._V1Api.bindingAddColumns(callArgs); }
		},
		"AddFromSelectionAsync": {
			call: function (ctx, callArgs) { return ctx.workbook._V1Api.bindingAddFromSelection(callArgs); },
			postprocess: postprocessBindingDescriptor
		},
		"AddFromNamedItemAsync": {
			call: function (ctx, callArgs) { return ctx.workbook._V1Api.bindingAddFromNamedItem(callArgs); },
			postprocess: postprocessBindingDescriptor
		},
		"AddFromPromptAsync": {
			call: function (ctx, callArgs) { return ctx.workbook._V1Api.bindingAddFromPrompt(callArgs); },
			postprocess: postprocessBindingDescriptor
		},
		"AddRowsAsync": {
			call: function (ctx, callArgs) { return ctx.workbook._V1Api.bindingAddRows(callArgs); }
		},
		"GetByIdAsync": {
			call: function (ctx, callArgs) { return ctx.workbook._V1Api.bindingGetById(callArgs); },
			postprocess: postprocessBindingDescriptor
		},
		"ReleaseByIdAsync": {
			call: function (ctx, callArgs) { return ctx.workbook._V1Api.bindingReleaseById(callArgs); }
		},
		"GetAllAsync": {
			call: function (ctx) { return ctx.workbook._V1Api.bindingGetAll(); },
			postprocess: function (response) {
				return response.bindings.map(function (descriptor) { return postprocessBindingDescriptor(descriptor); });
			}
		},
		"DeleteAllDataValuesAsync": {
			call: function (ctx, callArgs) { return ctx.workbook._V1Api.bindingDeleteAllDataValues(callArgs); }
		},
		"SetSelectedDataAsync": {
			preprocess: function (callArgs) {
				var preimage=callArgs["cellFormat"];
				if (typeof (window) !=="undefined" && window.OSF.DDA.SafeArray) {
					if (window.OSF.OUtil.listContainsKey(window.OSF.DDA.SafeArray.Delegate.ParameterMap.dynamicTypes, "cellFormat")) {
						callArgs["cellFormat"]=window.OSF.DDA.SafeArray.Delegate.ParameterMap.dynamicTypes["cellFormat"]["toHost"](preimage);
					}
				}
				else if (typeof (window) !=="undefined" && window.OSF.DDA.WAC) {
					if (window.OSF.OUtil.listContainsKey(window.OSF.DDA.WAC.Delegate.ParameterMap.dynamicTypes, "cellFormat")) {
						callArgs["cellFormat"]=window.OSF.DDA.WAC.Delegate.ParameterMap.dynamicTypes["cellFormat"]["toHost"](preimage);
					}
				}
				return callArgs;
			},
			call: function (ctx, callArgs) { return ctx.workbook._V1Api.setSelectedData(callArgs); }
		},
		"SetDataAsync": {
			preprocess: function (callArgs) {
				var preimage=callArgs["cellFormat"];
				if (typeof (window) !=="undefined" && window.OSF.DDA.SafeArray) {
					if (window.OSF.OUtil.listContainsKey(window.OSF.DDA.SafeArray.Delegate.ParameterMap.dynamicTypes, "cellFormat")) {
						callArgs["cellFormat"]=window.OSF.DDA.SafeArray.Delegate.ParameterMap.dynamicTypes["cellFormat"]["toHost"](preimage);
					}
				}
				else if (typeof (window) !=="undefined" && window.OSF.DDA.WAC) {
					if (window.OSF.OUtil.listContainsKey(window.OSF.DDA.WAC.Delegate.ParameterMap.dynamicTypes, "cellFormat")) {
						callArgs["cellFormat"]=window.OSF.DDA.WAC.Delegate.ParameterMap.dynamicTypes["cellFormat"]["toHost"](preimage);
					}
				}
				return callArgs;
			},
			call: function (ctx, callArgs) { return ctx.workbook._V1Api.bindingSetData(callArgs); }
		},
		"SetFormatsAsync": {
			preprocess: function (callArgs) {
				var preimage=callArgs["cellFormat"];
				if (typeof (window) !=="undefined" && window.OSF.DDA.SafeArray) {
					if (window.OSF.OUtil.listContainsKey(window.OSF.DDA.SafeArray.Delegate.ParameterMap.dynamicTypes, "cellFormat")) {
						callArgs["cellFormat"]=window.OSF.DDA.SafeArray.Delegate.ParameterMap.dynamicTypes["cellFormat"]["toHost"](preimage);
					}
				}
				else if (typeof (window) !=="undefined" && window.OSF.DDA.WAC) {
					if (window.OSF.OUtil.listContainsKey(window.OSF.DDA.WAC.Delegate.ParameterMap.dynamicTypes, "cellFormat")) {
						callArgs["cellFormat"]=window.OSF.DDA.WAC.Delegate.ParameterMap.dynamicTypes["cellFormat"]["toHost"](preimage);
					}
				}
				return callArgs;
			},
			call: function (ctx, callArgs) { return ctx.workbook._V1Api.bindingSetFormats(callArgs); }
		},
		"SetTableOptionsAsync": {
			call: function (ctx, callArgs) { return ctx.workbook._V1Api.bindingSetTableOptions(callArgs); }
		},
		"ClearFormatsAsync": {
			call: function (ctx, callArgs) { return ctx.workbook._V1Api.bindingClearFormats(callArgs); }
		},
		"GetFilePropertiesAsync": {
			call: function (ctx, callArgs) { return ctx.workbook._V1Api.getFilePropertiesAsync(callArgs); }
		},
	};
	function postprocessBindingDescriptor(response) {
		var bindingDescriptor={
			BindingColumnCount: response.bindingColumnCount,
			BindingId: response.bindingId,
			BindingRowCount: response.bindingRowCount,
			bindingType: response.bindingType,
			HasHeaders: response.hasHeaders
		};
		return window.OSF.DDA.OMFactory.manufactureBinding(bindingDescriptor, window.Microsoft.Office.WebExtension.context.document);
	}
	function getDataCommonPostprocess(response, callArgs) {
		var isPlainData=response.headers==null;
		var data;
		if (isPlainData) {
			data=response.rows;
		}
		else {
			data=response;
		}
		data=window.OSF.DDA.DataCoercion.coerceData(data, callArgs[window.Microsoft.Office.WebExtension.Parameters.CoercionType]);
		return data==undefined ? null : data;
	}
	function versionNumberIsEarlierThan(desiredMajor, desiredMinor) {
		var hasOfficeVersion = typeof (window) !== "undefined" && window.Office && window.Office.context && window.Office.context.diagnostics || window.Office.context.diagnostics.version);
		if (!hasOfficeVersion) {
			return false;
		}
		var version = window.Office.context.diagnostics.version;
		var versionExtractor = /^(\d+)\.\d+\.(\d+)\.\d+$/;
		var result = versionExtractor.exec(version);
		if (result) {
			var major = Number.parseInt(result[1]);
			var minor = Number.parseInt(result[2]);
			if (major < desiredMajor) {
				return true;
			}
			if (major == desiredMajor && minor < desiredMinor) {
				return true;
			}
		}
		return false;
	}
	var _hostName="Excel";
	var _defaultApiSetName="ExcelApi";
	var _createPropertyObjectPath=OfficeExtension.ObjectPathFactory.createPropertyObjectPath;
	var _createMethodObjectPath=OfficeExtension.ObjectPathFactory.createMethodObjectPath;
	var _createIndexerObjectPath=OfficeExtension.ObjectPathFactory.createIndexerObjectPath;
	var _createNewObjectObjectPath=OfficeExtension.ObjectPathFactory.createNewObjectObjectPath;
	var _createChildItemObjectPathUsingIndexer=OfficeExtension.ObjectPathFactory.createChildItemObjectPathUsingIndexer;
	var _createChildItemObjectPathUsingGetItemAt=OfficeExtension.ObjectPathFactory.createChildItemObjectPathUsingGetItemAt;
	var _createChildItemObjectPathUsingIndexerOrGetItemAt=OfficeExtension.ObjectPathFactory.createChildItemObjectPathUsingIndexerOrGetItemAt;
	var _createMethodAction=OfficeExtension.ActionFactory.createMethodAction;
	var _createSetPropertyAction=OfficeExtension.ActionFactory.createSetPropertyAction;
	var _isNullOrUndefined=OfficeExtension.Utility.isNullOrUndefined;
	var _isUndefined=OfficeExtension.Utility.isUndefined;
	var _throwIfNotLoaded=OfficeExtension.Utility.throwIfNotLoaded;
	var _throwIfApiNotSupported=OfficeExtension.Utility.throwIfApiNotSupported;
	var _load=OfficeExtension.Utility.load;
	var _fixObjectPathIfNecessary=OfficeExtension.Utility.fixObjectPathIfNecessary;
	var _addActionResultHandler=OfficeExtension.Utility._addActionResultHandler;
	var _handleNavigationPropertyResults=OfficeExtension.Utility._handleNavigationPropertyResults;
	var _adjustToDateTime=OfficeExtension.Utility.adjustToDateTime;
	var _typeApplication="Application";
	var Application=(function (_super) {
		__extends(Application, _super);
		function Application() {
			return _super !==null && _super.apply(this, arguments) || this;
		}
		Object.defineProperty(Application.prototype, "_className", {
			get: function () {
				return "Application";
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Application.prototype, "calculationMode", {
			get: function () {
				_throwIfNotLoaded("calculationMode", this._C, _typeApplication, this._isNull);
				return this._C;
			},
			enumerable: true,
			configurable: true
		});
		Application.prototype.calculate=function (calculationType) {
			_createMethodAction(this.context, this, "Calculate", 0, [calculationType]);
		};
		Application.prototype.suspendApiCalculationUntilNextSync=function () {
			_throwIfApiNotSupported("Application.suspendApiCalculationUntilNextSync", _defaultApiSetName, "1.6", _hostName);
			_createMethodAction(this.context, this, "SuspendApiCalculationUntilNextSync", 0, []);
		};
		Application.prototype._handleResult=function (value) {
			_super.prototype._handleResult.call(this, value);
			if (_isNullOrUndefined(value))
				return;
			var obj=value;
			_fixObjectPathIfNecessary(this, obj);
			if (!_isUndefined(obj["CalculationMode"])) {
				this._C=obj["CalculationMode"];
			}
		};
		Application.prototype.load=function (option) {
			_load(this, option);
			return this;
		};
		Application.prototype.toJSON=function () {
			return {
				"calculationMode": this._C
			};
		};
		return Application;
	}(OfficeExtension.ClientObject));
	Excel.Application=Application;
	var _typeWorkbook="Workbook";
	var Workbook=(function (_super) {
		__extends(Workbook, _super);
		function Workbook() {
			return _super !==null && _super.apply(this, arguments) || this;
		}
		Object.defineProperty(Workbook.prototype, "_className", {
			get: function () {
				return "Workbook";
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Workbook.prototype, "application", {
			get: function () {
				if (!this._A) {
					this._A=new Excel.Application(this.context, _createPropertyObjectPath(this.context, this, "Application", false, false));
				}
				return this._A;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Workbook.prototype, "bindings", {
			get: function () {
				if (!this._B) {
					this._B=new Excel.BindingCollection(this.context, _createPropertyObjectPath(this.context, this, "Bindings", true, false));
				}
				return this._B;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Workbook.prototype, "customXmlParts", {
			get: function () {
				_throwIfApiNotSupported("Workbook.customXmlParts", _defaultApiSetName, "1.5", _hostName);
				if (!this._C) {
					this._C=new Excel.CustomXmlPartCollection(this.context, _createPropertyObjectPath(this.context, this, "CustomXmlParts", true, false));
				}
				return this._C;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Workbook.prototype, "functions", {
			get: function () {
				_throwIfApiNotSupported("Workbook.functions", _defaultApiSetName, "1.2", _hostName);
				if (!this._F) {
					this._F=new Excel.Functions(this.context, _createPropertyObjectPath(this.context, this, "Functions", false, false));
				}
				return this._F;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Workbook.prototype, "internalTest", {
			get: function () {
				_throwIfApiNotSupported("Workbook.internalTest", _defaultApiSetName, "1.6", _hostName);
				if (!this._I) {
					this._I=new Excel.InternalTest(this.context, _createPropertyObjectPath(this.context, this, "InternalTest", false, false));
				}
				return this._I;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Workbook.prototype, "names", {
			get: function () {
				if (!this._N) {
					this._N=new Excel.NamedItemCollection(this.context, _createPropertyObjectPath(this.context, this, "Names", true, false));
				}
				return this._N;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Workbook.prototype, "pivotTables", {
			get: function () {
				_throwIfApiNotSupported("Workbook.pivotTables", _defaultApiSetName, "1.3", _hostName);
				if (!this._P) {
					this._P=new Excel.PivotTableCollection(this.context, _createPropertyObjectPath(this.context, this, "PivotTables", true, false));
				}
				return this._P;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Workbook.prototype, "settings", {
			get: function () {
				_throwIfApiNotSupported("Workbook.settings", _defaultApiSetName, "1.4", _hostName);
				if (!this._S) {
					this._S=new Excel.SettingCollection(this.context, _createPropertyObjectPath(this.context, this, "Settings", true, false));
				}
				return this._S;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Workbook.prototype, "tables", {
			get: function () {
				if (!this._T) {
					this._T=new Excel.TableCollection(this.context, _createPropertyObjectPath(this.context, this, "Tables", true, false));
				}
				return this._T;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Workbook.prototype, "worksheets", {
			get: function () {
				if (!this._W) {
					this._W=new Excel.WorksheetCollection(this.context, _createPropertyObjectPath(this.context, this, "Worksheets", true, false));
				}
				return this._W;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Workbook.prototype, "_V1Api", {
			get: function () {
				_throwIfApiNotSupported("Workbook._V1Api", _defaultApiSetName, "1.3", _hostName);
				if (!this.__V) {
					this.__V=new Excel._V1Api(this.context, _createPropertyObjectPath(this.context, this, "_V1Api", false, false));
				}
				return this.__V;
			},
			enumerable: true,
			configurable: true
		});
		Workbook.prototype.getSelectedRange=function () {
			return new Excel.Range(this.context, _createMethodObjectPath(this.context, this, "GetSelectedRange", 1, [], false, true, null));
		};
		Workbook.prototype._GetObjectByReferenceId=function (bstrReferenceId) {
			var action=_createMethodAction(this.context, this, "_GetObjectByReferenceId", 1, [bstrReferenceId]);
			var ret=new OfficeExtension.ClientResult();
			_addActionResultHandler(this, action, ret);
			return ret;
		};
		Workbook.prototype._GetObjectTypeNameByReferenceId=function (bstrReferenceId) {
			var action=_createMethodAction(this.context, this, "_GetObjectTypeNameByReferenceId", 1, [bstrReferenceId]);
			var ret=new OfficeExtension.ClientResult();
			_addActionResultHandler(this, action, ret);
			return ret;
		};
		Workbook.prototype._GetReferenceCount=function () {
			var action=_createMethodAction(this.context, this, "_GetReferenceCount", 1, []);
			var ret=new OfficeExtension.ClientResult();
			_addActionResultHandler(this, action, ret);
			return ret;
		};
		Workbook.prototype._RemoveAllReferences=function () {
			_createMethodAction(this.context, this, "_RemoveAllReferences", 1, []);
		};
		Workbook.prototype._RemoveReference=function (bstrReferenceId) {
			_createMethodAction(this.context, this, "_RemoveReference", 1, [bstrReferenceId]);
		};
		Workbook.prototype._handleResult=function (value) {
			_super.prototype._handleResult.call(this, value);
			if (_isNullOrUndefined(value))
				return;
			var obj=value;
			_fixObjectPathIfNecessary(this, obj);
			_handleNavigationPropertyResults(this, obj, ["application", "Application", "bindings", "Bindings", "customXmlParts", "CustomXmlParts", "functions", "Functions", "internalTest", "InternalTest", "names", "Names", "pivotTables", "PivotTables", "settings", "Settings", "tables", "Tables", "worksheets", "Worksheets", "_V1Api", "_V1Api"]);
		};
		Workbook.prototype.load=function (option) {
			_load(this, option);
			return this;
		};
		Object.defineProperty(Workbook.prototype, "onSelectionChanged", {
			get: function () {
				var _this=this;
				_throwIfApiNotSupported("Workbook.onSelectionChanged", _defaultApiSetName, "1.3", _hostName);
				if (!this.m_selectionChanged) {
					this.m_selectionChanged=new OfficeExtension.EventHandlers(this.context, this, "SelectionChanged", {
						registerFunc: function (handlerCallback) {
							return _this.context.eventRegistration.register(2, "", handlerCallback);
						},
						unregisterFunc: function (handlerCallback) {
							return _this.context.eventRegistration.unregister(2, "", handlerCallback);
						},
						eventArgsTransformFunc: function (args) {
							return OfficeExtension.Utility._createPromiseFromResult({ workbook: _this });
						}
					});
				}
				return this.m_selectionChanged;
			},
			enumerable: true,
			configurable: true
		});
		Workbook.prototype.toJSON=function () {
			return {};
		};
		return Workbook;
	}(OfficeExtension.ClientObject));
	Excel.Workbook=Workbook;
	var _typeWorksheet="Worksheet";
	var Worksheet=(function (_super) {
		__extends(Worksheet, _super);
		function Worksheet() {
			return _super !==null && _super.apply(this, arguments) || this;
		}
		Object.defineProperty(Worksheet.prototype, "_className", {
			get: function () {
				return "Worksheet";
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Worksheet.prototype, "charts", {
			get: function () {
				if (!this._C) {
					this._C=new Excel.ChartCollection(this.context, _createPropertyObjectPath(this.context, this, "Charts", true, false));
				}
				return this._C;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Worksheet.prototype, "names", {
			get: function () {
				_throwIfApiNotSupported("Worksheet.names", _defaultApiSetName, "1.4", _hostName);
				if (!this._Na) {
					this._Na=new Excel.NamedItemCollection(this.context, _createPropertyObjectPath(this.context, this, "Names", true, false));
				}
				return this._Na;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Worksheet.prototype, "pivotTables", {
			get: function () {
				_throwIfApiNotSupported("Worksheet.pivotTables", _defaultApiSetName, "1.3", _hostName);
				if (!this._P) {
					this._P=new Excel.PivotTableCollection(this.context, _createPropertyObjectPath(this.context, this, "PivotTables", true, false));
				}
				return this._P;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Worksheet.prototype, "protection", {
			get: function () {
				_throwIfApiNotSupported("Worksheet.protection", _defaultApiSetName, "1.2", _hostName);
				if (!this._Pr) {
					this._Pr=new Excel.WorksheetProtection(this.context, _createPropertyObjectPath(this.context, this, "Protection", false, false));
				}
				return this._Pr;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Worksheet.prototype, "tables", {
			get: function () {
				if (!this.m_tables) {
					this.m_tables=new Excel.TableCollection(this.context, _createPropertyObjectPath(this.context, this, "Tables", true, false));
				}
				this.m_tables._ParentObject=this;
				return this.m_tables;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Worksheet.prototype, "id", {
			get: function () {
				_throwIfNotLoaded("id", this._I, _typeWorksheet, this._isNull);
				return this._I;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Worksheet.prototype, "name", {
			get: function () {
				_throwIfNotLoaded("name", this._N, _typeWorksheet, this._isNull);
				return this._N;
			},
			set: function (value) {
				this._N=value;
				_createSetPropertyAction(this.context, this, "Name", value);
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Worksheet.prototype, "position", {
			get: function () {
				_throwIfNotLoaded("position", this._Po, _typeWorksheet, this._isNull);
				return this._Po;
			},
			set: function (value) {
				this._Po=value;
				_createSetPropertyAction(this.context, this, "Position", value);
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Worksheet.prototype, "visibility", {
			get: function () {
				_throwIfNotLoaded("visibility", this._V, _typeWorksheet, this._isNull);
				return this._V;
			},
			set: function (value) {
				this._V=value;
				_createSetPropertyAction(this.context, this, "Visibility", value);
			},
			enumerable: true,
			configurable: true
		});
		Worksheet.prototype.set=function (properties, options) {
			this._recursivelySet(properties, options, ["name", "position", "visibility"], [], [
				"charts",
				"names",
				"pivotTables",
				"tables",
				"charts",
				"names",
				"pivotTables",
				"protection",
				"tables"
			]);
		};
		Worksheet.prototype.activate=function () {
			_createMethodAction(this.context, this, "Activate", 1, []);
		};
		Worksheet.prototype.calculate=function (markAllDirty) {
			_throwIfApiNotSupported("Worksheet.calculate", _defaultApiSetName, "1.6", _hostName);
			_createMethodAction(this.context, this, "Calculate", 0, [markAllDirty]);
		};
		Worksheet.prototype.delete=function () {
			_createMethodAction(this.context, this, "Delete", 0, []);
		};
		Worksheet.prototype.getCell=function (row, column) {
			return new Excel.Range(this.context, _createMethodObjectPath(this.context, this, "GetCell", 1, [row, column], false, true, null));
		};
		Worksheet.prototype.getNext=function (visibleOnly) {
			_throwIfApiNotSupported("Worksheet.getNext", _defaultApiSetName, "1.5", _hostName);
			return new Excel.Worksheet(this.context, _createMethodObjectPath(this.context, this, "GetNext", 1, [visibleOnly], false, true, null));
		};
		Worksheet.prototype.getNextOrNullObject=function (visibleOnly) {
			_throwIfApiNotSupported("Worksheet.getNextOrNullObject", _defaultApiSetName, "1.5", _hostName);
			return new Excel.Worksheet(this.context, _createMethodObjectPath(this.context, this, "GetNextOrNullObject", 1, [visibleOnly], false, true, null));
		};
		Worksheet.prototype.getPrevious=function (visibleOnly) {
			_throwIfApiNotSupported("Worksheet.getPrevious", _defaultApiSetName, "1.5", _hostName);
			return new Excel.Worksheet(this.context, _createMethodObjectPath(this.context, this, "GetPrevious", 1, [visibleOnly], false, true, null));
		};
		Worksheet.prototype.getPreviousOrNullObject=function (visibleOnly) {
			_throwIfApiNotSupported("Worksheet.getPreviousOrNullObject", _defaultApiSetName, "1.5", _hostName);
			return new Excel.Worksheet(this.context, _createMethodObjectPath(this.context, this, "GetPreviousOrNullObject", 1, [visibleOnly], false, true, null));
		};
		Worksheet.prototype.getRange=function (address) {
			return new Excel.Range(this.context, _createMethodObjectPath(this.context, this, "GetRange", 1, [address], false, true, null));
		};
		Worksheet.prototype.getUsedRange=function (valuesOnly) {
			return new Excel.Range(this.context, _createMethodObjectPath(this.context, this, "GetUsedRange", 1, [valuesOnly], false, true, null));
		};
		Worksheet.prototype.getUsedRangeOrNullObject=function (valuesOnly) {
			_throwIfApiNotSupported("Worksheet.getUsedRangeOrNullObject", _defaultApiSetName, "1.4", _hostName);
			return new Excel.Range(this.context, _createMethodObjectPath(this.context, this, "GetUsedRangeOrNullObject", 1, [valuesOnly], false, true, null));
		};
		Worksheet.prototype._handleResult=function (value) {
			_super.prototype._handleResult.call(this, value);
			if (_isNullOrUndefined(value))
				return;
			var obj=value;
			_fixObjectPathIfNecessary(this, obj);
			if (!_isUndefined(obj["Id"])) {
				this._I=obj["Id"];
			}
			if (!_isUndefined(obj["Name"])) {
				this._N=obj["Name"];
			}
			if (!_isUndefined(obj["Position"])) {
				this._Po=obj["Position"];
			}
			if (!_isUndefined(obj["Visibility"])) {
				this._V=obj["Visibility"];
			}
			_handleNavigationPropertyResults(this, obj, ["charts", "Charts", "names", "Names", "pivotTables", "PivotTables", "protection", "Protection", "tables", "Tables"]);
		};
		Worksheet.prototype.load=function (option) {
			_load(this, option);
			return this;
		};
		Worksheet.prototype._handleIdResult=function (value) {
			_super.prototype._handleIdResult.call(this, value);
			if (_isNullOrUndefined(value)) {
				return;
			}
			if (!_isUndefined(value["Id"])) {
				this._I=value["Id"];
			}
		};
		Worksheet.prototype.toJSON=function () {
			return {
				"id": this._I,
				"name": this._N,
				"position": this._Po,
				"protection": this._Pr,
				"visibility": this._V
			};
		};
		return Worksheet;
	}(OfficeExtension.ClientObject));
	Excel.Worksheet=Worksheet;
	var _typeWorksheetCollection="WorksheetCollection";
	var WorksheetCollection=(function (_super) {
		__extends(WorksheetCollection, _super);
		function WorksheetCollection() {
			return _super !==null && _super.apply(this, arguments) || this;
		}
		Object.defineProperty(WorksheetCollection.prototype, "_className", {
			get: function () {
				return "WorksheetCollection";
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(WorksheetCollection.prototype, "items", {
			get: function () {
				_throwIfNotLoaded("items", this.m__items, _typeWorksheetCollection, this._isNull);
				return this.m__items;
			},
			enumerable: true,
			configurable: true
		});
		WorksheetCollection.prototype.add=function (name) {
			return new Excel.Worksheet(this.context, _createMethodObjectPath(this.context, this, "Add", 0, [name], false, true, null));
		};
		WorksheetCollection.prototype.getActiveWorksheet=function () {
			return new Excel.Worksheet(this.context, _createMethodObjectPath(this.context, this, "GetActiveWorksheet", 1, [], false, false, null));
		};
		WorksheetCollection.prototype.getCount=function (visibleOnly) {
			_throwIfApiNotSupported("WorksheetCollection.getCount", _defaultApiSetName, "1.4", _hostName);
			var action=_createMethodAction(this.context, this, "GetCount", 1, [visibleOnly]);
			var ret=new OfficeExtension.ClientResult();
			_addActionResultHandler(this, action, ret);
			return ret;
		};
		WorksheetCollection.prototype.getFirst=function (visibleOnly) {
			_throwIfApiNotSupported("WorksheetCollection.getFirst", _defaultApiSetName, "1.5", _hostName);
			return new Excel.Worksheet(this.context, _createMethodObjectPath(this.context, this, "GetFirst", 1, [visibleOnly], false, true, null));
		};
		WorksheetCollection.prototype.getItem=function (key) {
			return new Excel.Worksheet(this.context, _createIndexerObjectPath(this.context, this, [key]));
		};
		WorksheetCollection.prototype.getItemOrNullObject=function (key) {
			_throwIfApiNotSupported("WorksheetCollection.getItemOrNullObject", _defaultApiSetName, "1.4", _hostName);
			return new Excel.Worksheet(this.context, _createMethodObjectPath(this.context, this, "GetItemOrNullObject", 1, [key], false, false, null));
		};
		WorksheetCollection.prototype.getLast=function (visibleOnly) {
			_throwIfApiNotSupported("WorksheetCollection.getLast", _defaultApiSetName, "1.5", _hostName);
			return new Excel.Worksheet(this.context, _createMethodObjectPath(this.context, this, "GetLast", 1, [visibleOnly], false, true, null));
		};
		WorksheetCollection.prototype._handleResult=function (value) {
			_super.prototype._handleResult.call(this, value);
			if (_isNullOrUndefined(value))
				return;
			var obj=value;
			_fixObjectPathIfNecessary(this, obj);
			if (!_isNullOrUndefined(obj[OfficeExtension.Constants.items])) {
				this.m__items=[];
				var _data=obj[OfficeExtension.Constants.items];
				for (var i=0; i < _data.length; i++) {
					var _item=new Excel.Worksheet(this.context, _createChildItemObjectPathUsingIndexerOrGetItemAt(true, this.context, this, _data[i], i));
					_item._handleResult(_data[i]);
					this.m__items.push(_item);
				}
			}
		};
		WorksheetCollection.prototype.load=function (option) {
			_load(this, option);
			return this;
		};
		WorksheetCollection.prototype.toJSON=function () {
			return {};
		};
		return WorksheetCollection;
	}(OfficeExtension.ClientObject));
	Excel.WorksheetCollection=WorksheetCollection;
	var _typeWorksheetProtection="WorksheetProtection";
	var WorksheetProtection=(function (_super) {
		__extends(WorksheetProtection, _super);
		function WorksheetProtection() {
			return _super !==null && _super.apply(this, arguments) || this;
		}
		Object.defineProperty(WorksheetProtection.prototype, "_className", {
			get: function () {
				return "WorksheetProtection";
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(WorksheetProtection.prototype, "options", {
			get: function () {
				_throwIfNotLoaded("options", this._O, _typeWorksheetProtection, this._isNull);
				return this._O;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(WorksheetProtection.prototype, "protected", {
			get: function () {
				_throwIfNotLoaded("protected", this._P, _typeWorksheetProtection, this._isNull);
				return this._P;
			},
			enumerable: true,
			configurable: true
		});
		WorksheetProtection.prototype.protect=function (options, password) {
			if (versionNumberIsEarlierThan(16, 8716)) {
				_createMethodAction(this.context, this, "Protect", 0, [options]);
				return;
			}
			_createMethodAction(this.context, this, "Protect", 0, [options, password]);
		};
		WorksheetProtection.prototype.unprotect=function () {
			_createMethodAction(this.context, this, "Unprotect", 0, []);
		};
		WorksheetProtection.prototype._handleResult=function (value) {
			_super.prototype._handleResult.call(this, value);
			if (_isNullOrUndefined(value))
				return;
			var obj=value;
			_fixObjectPathIfNecessary(this, obj);
			if (!_isUndefined(obj["Options"])) {
				this._O=obj["Options"];
			}
			if (!_isUndefined(obj["Protected"])) {
				this._P=obj["Protected"];
			}
		};
		WorksheetProtection.prototype.load=function (option) {
			_load(this, option);
			return this;
		};
		WorksheetProtection.prototype.toJSON=function () {
			return {
				"options": this._O,
				"protected": this._P
			};
		};
		return WorksheetProtection;
	}(OfficeExtension.ClientObject));
	Excel.WorksheetProtection=WorksheetProtection;
	var _typeRange="Range";
	var Range=(function (_super) {
		__extends(Range, _super);
		function Range() {
			return _super !==null && _super.apply(this, arguments) || this;
		}
		Object.defineProperty(Range.prototype, "_className", {
			get: function () {
				return "Range";
			},
			enumerable: true,
			configurable: true
		});
		Range.prototype._ensureInteger=function (num, methodName) {
			if (!(typeof num==="number" && isFinite(num) && Math.floor(num)===num)) {
				throw new OfficeExtension.Utility.throwError(Excel.ErrorCodes.invalidArgument, num, methodName);
			}
		};
		Range.prototype._getAdjacentRange=function (functionName, count, referenceRange, rowDirection, columnDirection) {
			if (count==null) {
				count=1;
			}
			this._ensureInteger(count, functionName);
			var startRange;
			var rowOffset=0;
			var columnOffset=0;
			if (count > 0) {
				startRange=referenceRange.getOffsetRange(rowDirection, columnDirection);
			}
			else {
				startRange=referenceRange;
				rowOffset=rowDirection;
				columnOffset=columnDirection;
			}
			if (Math.abs(count)==1) {
				return startRange;
			}
			return startRange.getBoundingRect(referenceRange.getOffsetRange(rowDirection * count+rowOffset, columnDirection * count+columnOffset));
		};
		Object.defineProperty(Range.prototype, "conditionalFormats", {
			get: function () {
				_throwIfApiNotSupported("Range.conditionalFormats", _defaultApiSetName, "1.6", _hostName);
				if (!this._Con) {
					this._Con=new Excel.ConditionalFormatCollection(this.context, _createPropertyObjectPath(this.context, this, "ConditionalFormats", true, false));
				}
				return this._Con;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Range.prototype, "format", {
			get: function () {
				if (!this._F) {
					this._F=new Excel.RangeFormat(this.context, _createPropertyObjectPath(this.context, this, "Format", false, false));
				}
				return this._F;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Range.prototype, "sort", {
			get: function () {
				_throwIfApiNotSupported("Range.sort", _defaultApiSetName, "1.2", _hostName);
				if (!this._S) {
					this._S=new Excel.RangeSort(this.context, _createPropertyObjectPath(this.context, this, "Sort", false, false));
				}
				return this._S;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Range.prototype, "worksheet", {
			get: function () {
				if (!this._W) {
					this._W=new Excel.Worksheet(this.context, _createPropertyObjectPath(this.context, this, "Worksheet", false, false));
				}
				return this._W;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Range.prototype, "address", {
			get: function () {
				_throwIfNotLoaded("address", this._A, _typeRange, this._isNull);
				return this._A;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Range.prototype, "addressLocal", {
			get: function () {
				_throwIfNotLoaded("addressLocal", this._Ad, _typeRange, this._isNull);
				return this._Ad;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Range.prototype, "cellCount", {
			get: function () {
				_throwIfNotLoaded("cellCount", this._C, _typeRange, this._isNull);
				return this._C;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Range.prototype, "columnCount", {
			get: function () {
				_throwIfNotLoaded("columnCount", this._Co, _typeRange, this._isNull);
				return this._Co;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Range.prototype, "columnHidden", {
			get: function () {
				_throwIfNotLoaded("columnHidden", this._Col, _typeRange, this._isNull);
				_throwIfApiNotSupported("Range.columnHidden", _defaultApiSetName, "1.2", _hostName);
				return this._Col;
			},
			set: function (value) {
				this._Col=value;
				_createSetPropertyAction(this.context, this, "ColumnHidden", value);
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Range.prototype, "columnIndex", {
			get: function () {
				_throwIfNotLoaded("columnIndex", this._Colu, _typeRange, this._isNull);
				return this._Colu;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Range.prototype, "formulas", {
			get: function () {
				_throwIfNotLoaded("formulas", this.m_formulas, _typeRange, this._isNull);
				return this.m_formulas;
			},
			set: function (value) {
				this.m_formulas=value;
				if (setRangePropertiesInBulk(this, "Formulas", value)) {
					return;
				}
				this.m_formulas=value;
				_createSetPropertyAction(this.context, this, "Formulas", value);
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Range.prototype, "formulasLocal", {
			get: function () {
				_throwIfNotLoaded("formulasLocal", this.m_formulasLocal, _typeRange, this._isNull);
				return this.m_formulasLocal;
			},
			set: function (value) {
				this.m_formulasLocal=value;
				if (setRangePropertiesInBulk(this, "FormulasLocal", value)) {
					return;
				}
				this.m_formulasLocal=value;
				_createSetPropertyAction(this.context, this, "FormulasLocal", value);
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Range.prototype, "formulasR1C1", {
			get: function () {
				_throwIfNotLoaded("formulasR1C1", this.m_formulasR1C1, _typeRange, this._isNull);
				_throwIfApiNotSupported("Range.formulasR1C1", _defaultApiSetName, "1.2", _hostName);
				return this.m_formulasR1C1;
			},
			set: function (value) {
				this.m_formulasR1C1=value;
				if (setRangePropertiesInBulk(this, "FormulasR1C1", value)) {
					return;
				}
				this.m_formulasR1C1=value;
				_createSetPropertyAction(this.context, this, "FormulasR1C1", value);
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Range.prototype, "hidden", {
			get: function () {
				_throwIfNotLoaded("hidden", this._H, _typeRange, this._isNull);
				_throwIfApiNotSupported("Range.hidden", _defaultApiSetName, "1.2", _hostName);
				return this._H;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Range.prototype, "numberFormat", {
			get: function () {
				_throwIfNotLoaded("numberFormat", this.m_numberFormat, _typeRange, this._isNull);
				return this.m_numberFormat;
			},
			set: function (value) {
				this.m_numberFormat=value;
				if (setRangePropertiesInBulk(this, "NumberFormat", value)) {
					return;
				}
				this.m_numberFormat=value;
				_createSetPropertyAction(this.context, this, "NumberFormat", value);
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Range.prototype, "rowCount", {
			get: function () {
				_throwIfNotLoaded("rowCount", this._R, _typeRange, this._isNull);
				return this._R;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Range.prototype, "rowHidden", {
			get: function () {
				_throwIfNotLoaded("rowHidden", this._Ro, _typeRange, this._isNull);
				_throwIfApiNotSupported("Range.rowHidden", _defaultApiSetName, "1.2", _hostName);
				return this._Ro;
			},
			set: function (value) {
				this._Ro=value;
				_createSetPropertyAction(this.context, this, "RowHidden", value);
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Range.prototype, "rowIndex", {
			get: function () {
				_throwIfNotLoaded("rowIndex", this._Row, _typeRange, this._isNull);
				return this._Row;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Range.prototype, "text", {
			get: function () {
				_throwIfNotLoaded("text", this._T, _typeRange, this._isNull);
				return this._T;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Range.prototype, "valueTypes", {
			get: function () {
				_throwIfNotLoaded("valueTypes", this._V, _typeRange, this._isNull);
				return this._V;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Range.prototype, "values", {
			get: function () {
				_throwIfNotLoaded("values", this.m_values, _typeRange, this._isNull);
				return this.m_values;
			},
			set: function (value) {
				this.m_values=value;
				if (setRangePropertiesInBulk(this, "Values", value)) {
					return;
				}
				this.m_values=value;
				_createSetPropertyAction(this.context, this, "Values", value);
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Range.prototype, "_ReferenceId", {
			get: function () {
				_throwIfNotLoaded("_ReferenceId", this.__R, _typeRange, this._isNull);
				return this.__R;
			},
			enumerable: true,
			configurable: true
		});
		Range.prototype.set=function (properties, options) {
			this._recursivelySet(properties, options, ["numberFormat", "values", "formulas", "formulasLocal", "formulasR1C1", "rowHidden", "columnHidden"], ["format"], [
				"conditionalFormats",
				"sort",
				"worksheet",
				"conditionalFormats",
				"sort",
				"worksheet"
			]);
		};
		Range.prototype.calculate=function () {
			_throwIfApiNotSupported("Range.calculate", _defaultApiSetName, "1.6", _hostName);
			_createMethodAction(this.context, this, "Calculate", 0, []);
		};
		Range.prototype.clear=function (applyTo) {
			_createMethodAction(this.context, this, "Clear", 0, [applyTo]);
		};
		Range.prototype.delete=function (shift) {
			_createMethodAction(this.context, this, "Delete", 0, [shift]);
		};
		Range.prototype.getBoundingRect=function (anotherRange) {
			return new Excel.Range(this.context, _createMethodObjectPath(this.context, this, "GetBoundingRect", 1, [anotherRange], false, true, null));
		};
		Range.prototype.getCell=function (row, column) {
			return new Excel.Range(this.context, _createMethodObjectPath(this.context, this, "GetCell", 1, [row, column], false, true, null));
		};
		Range.prototype.getColumn=function (column) {
			return new Excel.Range(this.context, _createMethodObjectPath(this.context, this, "GetColumn", 1, [column], false, true, null));
		};
		Range.prototype.getColumnsAfter=function (count) {
			if (!isExcel1_3OrAbove()) {
				if (count==null) {
					count=1;
				}
				this._ensureInteger(count, "RowsAbove");
				if (count==0) {
					throw new OfficeExtension.Utility.throwError(Excel.ErrorCodes.invalidArgument, "count", "RowsAbove");
				}
				return this._getAdjacentRange("getColumnsAfter", count, this.getLastColumn(), 0, 1);
			}
			_throwIfApiNotSupported("Range.getColumnsAfter", _defaultApiSetName, "1.3", _hostName);
			return new Excel.Range(this.context, _createMethodObjectPath(this.context, this, "GetColumnsAfter", 1, [count], false, true, null));
		};
		Range.prototype.getColumnsBefore=function (count) {
			if (!isExcel1_3OrAbove()) {
				if (count==null) {
					count=1;
				}
				this._ensureInteger(count, "RowsAbove");
				if (count==0) {
					throw new OfficeExtension.Utility.throwError(Excel.ErrorCodes.invalidArgument, "count", "RowsAbove");
				}
				return this._getAdjacentRange("getColumnsBefore", count, this.getColumn(0), 0, -1);
			}
			_throwIfApiNotSupported("Range.getColumnsBefore", _defaultApiSetName, "1.3", _hostName);
			return new Excel.Range(this.context, _createMethodObjectPath(this.context, this, "GetColumnsBefore", 1, [count], false, true, null));
		};
		Range.prototype.getEntireColumn=function () {
			return new Excel.Range(this.context, _createMethodObjectPath(this.context, this, "GetEntireColumn", 1, [], false, true, null));
		};
		Range.prototype.getEntireRow=function () {
			return new Excel.Range(this.context, _createMethodObjectPath(this.context, this, "GetEntireRow", 1, [], false, true, null));
		};
		Range.prototype.getIntersection=function (anotherRange) {
			return new Excel.Range(this.context, _createMethodObjectPath(this.context, this, "GetIntersection", 1, [anotherRange], false, true, null));
		};
		Range.prototype.getIntersectionOrNullObject=function (anotherRange) {
			_throwIfApiNotSupported("Range.getIntersectionOrNullObject", _defaultApiSetName, "1.4", _hostName);
			return new Excel.Range(this.context, _createMethodObjectPath(this.context, this, "GetIntersectionOrNullObject", 1, [anotherRange], false, true, null));
		};
		Range.prototype.getLastCell=function () {
			return new Excel.Range(this.context, _createMethodObjectPath(this.context, this, "GetLastCell", 1, [], false, true, null));
		};
		Range.prototype.getLastColumn=function () {
			return new Excel.Range(this.context, _createMethodObjectPath(this.context, this, "GetLastColumn", 1, [], false, true, null));
		};
		Range.prototype.getLastRow=function () {
			return new Excel.Range(this.context, _createMethodObjectPath(this.context, this, "GetLastRow", 1, [], false, true, null));
		};
		Range.prototype.getOffsetRange=function (rowOffset, columnOffset) {
			return new Excel.Range(this.context, _createMethodObjectPath(this.context, this, "GetOffsetRange", 1, [rowOffset, columnOffset], false, true, null));
		};
		Range.prototype.getResizedRange=function (deltaRows, deltaColumns) {
			if (!isExcel1_3OrAbove()) {
				this._ensureInteger(deltaRows, "getResizedRange");
				this._ensureInteger(deltaColumns, "getResizedRange");
				var referenceRange=(deltaRows >=0 && deltaColumns >=0) ? this : this.getCell(0, 0);
				return referenceRange.getBoundingRect(this.getLastCell().getOffsetRange(deltaRows, deltaColumns));
			}
			_throwIfApiNotSupported("Range.getResizedRange", _defaultApiSetName, "1.3", _hostName);
			return new Excel.Range(this.context, _createMethodObjectPath(this.context, this, "GetResizedRange", 1, [deltaRows, deltaColumns], false, true, null));
		};
		Range.prototype.getRow=function (row) {
			return new Excel.Range(this.context, _createMethodObjectPath(this.context, this, "GetRow", 1, [row], false, true, null));
		};
		Range.prototype.getRowsAbove=function (count) {
			if (!isExcel1_3OrAbove()) {
				if (count==null) {
					count=1;
				}
				this._ensureInteger(count, "RowsAbove");
				if (count==0) {
					throw new OfficeExtension.Utility.throwError(Excel.ErrorCodes.invalidArgument, "count", "RowsAbove");
				}
				return this._getAdjacentRange("getRowsAbove", count, this.getRow(0), -1, 0);
			}
			_throwIfApiNotSupported("Range.getRowsAbove", _defaultApiSetName, "1.3", _hostName);
			return new Excel.Range(this.context, _createMethodObjectPath(this.context, this, "GetRowsAbove", 1, [count], false, true, null));
		};
		Range.prototype.getRowsBelow=function (count) {
			if (!isExcel1_3OrAbove()) {
				if (count==null) {
					count=1;
				}
				this._ensureInteger(count, "RowsAbove");
				if (count==0) {
					throw new OfficeExtension.Utility.throwError(Excel.ErrorCodes.invalidArgument, "count", "RowsAbove");
				}
				return this._getAdjacentRange("getRowsBelow", count, this.getLastRow(), 1, 0);
			}
			_throwIfApiNotSupported("Range.getRowsBelow", _defaultApiSetName, "1.3", _hostName);
			return new Excel.Range(this.context, _createMethodObjectPath(this.context, this, "GetRowsBelow", 1, [count], false, true, null));
		};
		Range.prototype.getUsedRange=function (valuesOnly) {
			return new Excel.Range(this.context, _createMethodObjectPath(this.context, this, "GetUsedRange", 1, [valuesOnly], false, true, null));
		};
		Range.prototype.getUsedRangeOrNullObject=function (valuesOnly) {
			_throwIfApiNotSupported("Range.getUsedRangeOrNullObject", _defaultApiSetName, "1.4", _hostName);
			return new Excel.Range(this.context, _createMethodObjectPath(this.context, this, "GetUsedRangeOrNullObject", 1, [valuesOnly], false, true, null));
		};
		Range.prototype.getVisibleView=function () {
			_throwIfApiNotSupported("Range.getVisibleView", _defaultApiSetName, "1.3", _hostName);
			return new Excel.RangeView(this.context, _createMethodObjectPath(this.context, this, "GetVisibleView", 1, [], false, false, null));
		};
		Range.prototype.insert=function (shift) {
			return new Excel.Range(this.context, _createMethodObjectPath(this.context, this, "Insert", 0, [shift], false, true, null));
		};
		Range.prototype.merge=function (across) {
			_throwIfApiNotSupported("Range.merge", _defaultApiSetName, "1.2", _hostName);
			_createMethodAction(this.context, this, "Merge", 0, [across]);
		};
		Range.prototype.select=function () {
			_createMethodAction(this.context, this, "Select", 1, []);
		};
		Range.prototype.unmerge=function () {
			_throwIfApiNotSupported("Range.unmerge", _defaultApiSetName, "1.2", _hostName);
			_createMethodAction(this.context, this, "Unmerge", 0, []);
		};
		Range.prototype._KeepReference=function () {
			_createMethodAction(this.context, this, "_KeepReference", 1, []);
		};
		Range.prototype._ValidateArraySize=function (rows, columns) {
			_throwIfApiNotSupported("Range._ValidateArraySize", _defaultApiSetName, "1.3", _hostName);
			_createMethodAction(this.context, this, "_ValidateArraySize", 1, [rows, columns]);
		};
		Range.prototype._handleResult=function (value) {
			_super.prototype._handleResult.call(this, value);
			if (_isNullOrUndefined(value))
				return;
			var obj=value;
			_fixObjectPathIfNecessary(this, obj);
			if (!_isUndefined(obj["Address"])) {
				this._A=obj["Address"];
			}
			if (!_isUndefined(obj["AddressLocal"])) {
				this._Ad=obj["AddressLocal"];
			}
			if (!_isUndefined(obj["CellCount"])) {
				this._C=obj["CellCount"];
			}
			if (!_isUndefined(obj["ColumnCount"])) {
				this._Co=obj["ColumnCount"];
			}
			if (!_isUndefined(obj["ColumnHidden"])) {
				this._Col=obj["ColumnHidden"];
			}
			if (!_isUndefined(obj["ColumnIndex"])) {
				this._Colu=obj["ColumnIndex"];
			}
			if (!_isUndefined(obj["Formulas"])) {
				this.m_formulas=obj["Formulas"];
			}
			if (!_isUndefined(obj["FormulasLocal"])) {
				this.m_formulasLocal=obj["FormulasLocal"];
			}
			if (!_isUndefined(obj["FormulasR1C1"])) {
				this.m_formulasR1C1=obj["FormulasR1C1"];
			}
			if (!_isUndefined(obj["Hidden"])) {
				this._H=obj["Hidden"];
			}
			if (!_isUndefined(obj["NumberFormat"])) {
				this.m_numberFormat=obj["NumberFormat"];
			}
			if (!_isUndefined(obj["RowCount"])) {
				this._R=obj["RowCount"];
			}
			if (!_isUndefined(obj["RowHidden"])) {
				this._Ro=obj["RowHidden"];
			}
			if (!_isUndefined(obj["RowIndex"])) {
				this._Row=obj["RowIndex"];
			}
			if (!_isUndefined(obj["Text"])) {
				this._T=obj["Text"];
			}
			if (!_isUndefined(obj["ValueTypes"])) {
				this._V=obj["ValueTypes"];
			}
			if (!_isUndefined(obj["Values"])) {
				this.m_values=obj["Values"];
			}
			if (!_isUndefined(obj["_ReferenceId"])) {
				this.__R=obj["_ReferenceId"];
			}
			_handleNavigationPropertyResults(this, obj, ["conditionalFormats", "ConditionalFormats", "format", "Format", "sort", "Sort", "worksheet", "Worksheet"]);
		};
		Range.prototype.load=function (option) {
			_load(this, option);
			return this;
		};
		Range.prototype._handleIdResult=function (value) {
			_super.prototype._handleIdResult.call(this, value);
			if (_isNullOrUndefined(value)) {
				return;
			}
			if (!_isUndefined(value["_ReferenceId"])) {
				this.__R=value["_ReferenceId"];
			}
		};
		Range.prototype.track=function () {
			this.context.trackedObjects.add(this);
			return this;
		};
		Range.prototype.untrack=function () {
			this.context.trackedObjects.remove(this);
			return this;
		};
		Range.prototype.toJSON=function () {
			return {
				"address": this._A,
				"addressLocal": this._Ad,
				"cellCount": this._C,
				"columnCount": this._Co,
				"columnHidden": this._Col,
				"columnIndex": this._Colu,
				"format": this._F,
				"formulas": this.m_formulas,
				"formulasLocal": this.m_formulasLocal,
				"formulasR1C1": this.m_formulasR1C1,
				"hidden": this._H,
				"numberFormat": this.m_numberFormat,
				"rowCount": this._R,
				"rowHidden": this._Ro,
				"rowIndex": this._Row,
				"text": this._T,
				"values": this.m_values,
				"valueTypes": this._V
			};
		};
		return Range;
	}(OfficeExtension.ClientObject));
	Excel.Range=Range;
	var _typeRangeView="RangeView";
	var RangeView=(function (_super) {
		__extends(RangeView, _super);
		function RangeView() {
			return _super !==null && _super.apply(this, arguments) || this;
		}
		Object.defineProperty(RangeView.prototype, "_className", {
			get: function () {
				return "RangeView";
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(RangeView.prototype, "rows", {
			get: function () {
				if (!this._Ro) {
					this._Ro=new Excel.RangeViewCollection(this.context, _createPropertyObjectPath(this.context, this, "Rows", true, false));
				}
				return this._Ro;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(RangeView.prototype, "cellAddresses", {
			get: function () {
				_throwIfNotLoaded("cellAddresses", this._C, _typeRangeView, this._isNull);
				return this._C;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(RangeView.prototype, "columnCount", {
			get: function () {
				_throwIfNotLoaded("columnCount", this._Co, _typeRangeView, this._isNull);
				return this._Co;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(RangeView.prototype, "formulas", {
			get: function () {
				_throwIfNotLoaded("formulas", this._F, _typeRangeView, this._isNull);
				return this._F;
			},
			set: function (value) {
				this._F=value;
				_createSetPropertyAction(this.context, this, "Formulas", value);
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(RangeView.prototype, "formulasLocal", {
			get: function () {
				_throwIfNotLoaded("formulasLocal", this._Fo, _typeRangeView, this._isNull);
				return this._Fo;
			},
			set: function (value) {
				this._Fo=value;
				_createSetPropertyAction(this.context, this, "FormulasLocal", value);
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(RangeView.prototype, "formulasR1C1", {
			get: function () {
				_throwIfNotLoaded("formulasR1C1", this._For, _typeRangeView, this._isNull);
				return this._For;
			},
			set: function (value) {
				this._For=value;
				_createSetPropertyAction(this.context, this, "FormulasR1C1", value);
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(RangeView.prototype, "index", {
			get: function () {
				_throwIfNotLoaded("index", this._I, _typeRangeView, this._isNull);
				return this._I;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(RangeView.prototype, "numberFormat", {
			get: function () {
				_throwIfNotLoaded("numberFormat", this._N, _typeRangeView, this._isNull);
				return this._N;
			},
			set: function (value) {
				this._N=value;
				_createSetPropertyAction(this.context, this, "NumberFormat", value);
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(RangeView.prototype, "rowCount", {
			get: function () {
				_throwIfNotLoaded("rowCount", this._R, _typeRangeView, this._isNull);
				return this._R;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(RangeView.prototype, "text", {
			get: function () {
				_throwIfNotLoaded("text", this._T, _typeRangeView, this._isNull);
				return this._T;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(RangeView.prototype, "valueTypes", {
			get: function () {
				_throwIfNotLoaded("valueTypes", this._Va, _typeRangeView, this._isNull);
				return this._Va;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(RangeView.prototype, "values", {
			get: function () {
				_throwIfNotLoaded("values", this._V, _typeRangeView, this._isNull);
				return this._V;
			},
			set: function (value) {
				this._V=value;
				_createSetPropertyAction(this.context, this, "Values", value);
			},
			enumerable: true,
			configurable: true
		});
		RangeView.prototype.set=function (properties, options) {
			this._recursivelySet(properties, options, ["numberFormat", "values", "formulas", "formulasLocal", "formulasR1C1"], [], [
				"rows",
				"rows"
			]);
		};
		RangeView.prototype.getRange=function () {
			return new Excel.Range(this.context, _createMethodObjectPath(this.context, this, "GetRange", 1, [], false, true, null));
		};
		RangeView.prototype._handleResult=function (value) {
			_super.prototype._handleResult.call(this, value);
			if (_isNullOrUndefined(value))
				return;
			var obj=value;
			_fixObjectPathIfNecessary(this, obj);
			if (!_isUndefined(obj["CellAddresses"])) {
				this._C=obj["CellAddresses"];
			}
			if (!_isUndefined(obj["ColumnCount"])) {
				this._Co=obj["ColumnCount"];
			}
			if (!_isUndefined(obj["Formulas"])) {
				this._F=obj["Formulas"];
			}
			if (!_isUndefined(obj["FormulasLocal"])) {
				this._Fo=obj["FormulasLocal"];
			}
			if (!_isUndefined(obj["FormulasR1C1"])) {
				this._For=obj["FormulasR1C1"];
			}
			if (!_isUndefined(obj["Index"])) {
				this._I=obj["Index"];
			}
			if (!_isUndefined(obj["NumberFormat"])) {
				this._N=obj["NumberFormat"];
			}
			if (!_isUndefined(obj["RowCount"])) {
				this._R=obj["RowCount"];
			}
			if (!_isUndefined(obj["Text"])) {
				this._T=obj["Text"];
			}
			if (!_isUndefined(obj["ValueTypes"])) {
				this._Va=obj["ValueTypes"];
			}
			if (!_isUndefined(obj["Values"])) {
				this._V=obj["Values"];
			}
			_handleNavigationPropertyResults(this, obj, ["rows", "Rows"]);
		};
		RangeView.prototype.load=function (option) {
			_load(this, option);
			return this;
		};
		RangeView.prototype.toJSON=function () {
			return {
				"cellAddresses": this._C,
				"columnCount": this._Co,
				"formulas": this._F,
				"formulasLocal": this._Fo,
				"formulasR1C1": this._For,
				"index": this._I,
				"numberFormat": this._N,
				"rowCount": this._R,
				"text": this._T,
				"values": this._V,
				"valueTypes": this._Va
			};
		};
		return RangeView;
	}(OfficeExtension.ClientObject));
	Excel.RangeView=RangeView;
	var _typeRangeViewCollection="RangeViewCollection";
	var RangeViewCollection=(function (_super) {
		__extends(RangeViewCollection, _super);
		function RangeViewCollection() {
			return _super !==null && _super.apply(this, arguments) || this;
		}
		Object.defineProperty(RangeViewCollection.prototype, "_className", {
			get: function () {
				return "RangeViewCollection";
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(RangeViewCollection.prototype, "items", {
			get: function () {
				_throwIfNotLoaded("items", this.m__items, _typeRangeViewCollection, this._isNull);
				return this.m__items;
			},
			enumerable: true,
			configurable: true
		});
		RangeViewCollection.prototype.getCount=function () {
			_throwIfApiNotSupported("RangeViewCollection.getCount", _defaultApiSetName, "1.4", _hostName);
			var action=_createMethodAction(this.context, this, "GetCount", 1, []);
			var ret=new OfficeExtension.ClientResult();
			_addActionResultHandler(this, action, ret);
			return ret;
		};
		RangeViewCollection.prototype.getItemAt=function (index) {
			return new Excel.RangeView(this.context, _createMethodObjectPath(this.context, this, "GetItemAt", 1, [index], false, false, null));
		};
		RangeViewCollection.prototype._handleResult=function (value) {
			_super.prototype._handleResult.call(this, value);
			if (_isNullOrUndefined(value))
				return;
			var obj=value;
			_fixObjectPathIfNecessary(this, obj);
			if (!_isNullOrUndefined(obj[OfficeExtension.Constants.items])) {
				this.m__items=[];
				var _data=obj[OfficeExtension.Constants.items];
				for (var i=0; i < _data.length; i++) {
					var _item=new Excel.RangeView(this.context, _createChildItemObjectPathUsingIndexerOrGetItemAt(false, this.context, this, _data[i], i));
					_item._handleResult(_data[i]);
					this.m__items.push(_item);
				}
			}
		};
		RangeViewCollection.prototype.load=function (option) {
			_load(this, option);
			return this;
		};
		RangeViewCollection.prototype.toJSON=function () {
			return {};
		};
		return RangeViewCollection;
	}(OfficeExtension.ClientObject));
	Excel.RangeViewCollection=RangeViewCollection;
	var _typeSettingCollection="SettingCollection";
	var SettingCollection=(function (_super) {
		__extends(SettingCollection, _super);
		function SettingCollection() {
			return _super !==null && _super.apply(this, arguments) || this;
		}
		Object.defineProperty(SettingCollection.prototype, "_className", {
			get: function () {
				return "SettingCollection";
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(SettingCollection.prototype, "items", {
			get: function () {
				_throwIfNotLoaded("items", this.m__items, _typeSettingCollection, this._isNull);
				return this.m__items;
			},
			enumerable: true,
			configurable: true
		});
		SettingCollection.prototype.add=function (key, value) {
			value=Setting._replaceDateWithStringDate(value);
			return new Excel.Setting(this.context, _createMethodObjectPath(this.context, this, "Add", 0, [key, value], false, true, null));
		};
		SettingCollection.prototype.getCount=function () {
			var action=_createMethodAction(this.context, this, "GetCount", 1, []);
			var ret=new OfficeExtension.ClientResult();
			_addActionResultHandler(this, action, ret);
			return ret;
		};
		SettingCollection.prototype.getItem=function (key) {
			return new Excel.Setting(this.context, _createIndexerObjectPath(this.context, this, [key]));
		};
		SettingCollection.prototype.getItemOrNullObject=function (key) {
			return new Excel.Setting(this.context, _createMethodObjectPath(this.context, this, "GetItemOrNullObject", 1, [key], false, false, null));
		};
		SettingCollection.prototype._handleResult=function (value) {
			_super.prototype._handleResult.call(this, value);
			if (_isNullOrUndefined(value))
				return;
			var obj=value;
			_fixObjectPathIfNecessary(this, obj);
			if (!_isNullOrUndefined(obj[OfficeExtension.Constants.items])) {
				this.m__items=[];
				var _data=obj[OfficeExtension.Constants.items];
				for (var i=0; i < _data.length; i++) {
					var _item=new Excel.Setting(this.context, _createChildItemObjectPathUsingIndexerOrGetItemAt(true, this.context, this, _data[i], i));
					_item._handleResult(_data[i]);
					this.m__items.push(_item);
				}
			}
		};
		SettingCollection.prototype.load=function (option) {
			_load(this, option);
			return this;
		};
		Object.defineProperty(SettingCollection.prototype, "onSettingsChanged", {
			get: function () {
				var _this=this;
				if (!this.m_settingsChanged) {
					this.m_settingsChanged=new OfficeExtension.EventHandlers(this.context, this, "SettingsChanged", {
						registerFunc: function (handlerCallback) {
							return _this.context.eventRegistration.register(1, "", handlerCallback);
						},
						unregisterFunc: function (handlerCallback) {
							return _this.context.eventRegistration.unregister(1, "", handlerCallback);
						},
						eventArgsTransformFunc: function (args) {
							return OfficeExtension.Utility._createPromiseFromResult({ settings: _this });
						}
					});
				}
				return this.m_settingsChanged;
			},
			enumerable: true,
			configurable: true
		});
		SettingCollection.prototype.toJSON=function () {
			return {};
		};
		return SettingCollection;
	}(OfficeExtension.ClientObject));
	Excel.SettingCollection=SettingCollection;
	var _typeSetting="Setting";
	var Setting=(function (_super) {
		__extends(Setting, _super);
		function Setting() {
			return _super !==null && _super.apply(this, arguments) || this;
		}
		Object.defineProperty(Setting.prototype, "_className", {
			get: function () {
				return "Setting";
			},
			enumerable: true,
			configurable: true
		});
		Setting.replaceStringDateWithDate=function (value) {
			var strValue=JSON.stringify(value);
			value=JSON.parse(strValue, function dateReviver(k, v) {
				var d;
				if (typeof v==='string' && v && v.length > 6 && v.slice(0, 5)===Setting.DateJSONPrefix && v.slice(-1)===Setting.DateJSONSuffix) {
					d=new Date(parseInt(v.slice(5, -1)));
					if (d) {
						return d;
					}
				}
				return v;
			});
			return value;
		};
		Setting._replaceDateWithStringDate=function (value) {
			var strValue=JSON.stringify(value, function dateReplacer(k, v) {
				return (this[k] instanceof Date) ? (Setting.DateJSONPrefix+this[k].getTime()+Setting.DateJSONSuffix) : v;
			});
			value=JSON.parse(strValue);
			return value;
		};
		Object.defineProperty(Setting.prototype, "key", {
			get: function () {
				_throwIfNotLoaded("key", this._K, _typeSetting, this._isNull);
				return this._K;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Setting.prototype, "value", {
			get: function () {
				_throwIfNotLoaded("value", this.m_value, _typeSetting, this._isNull);
				return this.m_value;
			},
			set: function (value) {
				if (!_isNullOrUndefined(value)) {
					this.m_value=value;
					var newValue=Setting._replaceDateWithStringDate(value);
					_createSetPropertyAction(this.context, this, "Value", newValue);
					return;
				}
				this.m_value=value;
				_createSetPropertyAction(this.context, this, "Value", value);
			},
			enumerable: true,
			configurable: true
		});
		Setting.prototype.set=function (properties, options) {
			this._recursivelySet(properties, options, ["value"], [], []);
		};
		Setting.prototype.delete=function () {
			_createMethodAction(this.context, this, "Delete", 0, []);
		};
		Setting.prototype._handleResult=function (value) {
			_super.prototype._handleResult.call(this, value);
			if (_isNullOrUndefined(value))
				return;
			var obj=value;
			_fixObjectPathIfNecessary(this, obj);
			if (!_isUndefined(obj["Key"])) {
				this._K=obj["Key"];
			}
			if (!_isUndefined(obj["Value"])) {
				this.m_value=obj["Value"];
				this.m_value=Setting.replaceStringDateWithDate(this.m_value);
			}
		};
		Setting.prototype.load=function (option) {
			_load(this, option);
			return this;
		};
		Setting.prototype.toJSON=function () {
			return {
				"key": this._K,
				"value": this.m_value
			};
		};
		return Setting;
	}(OfficeExtension.ClientObject));
	Setting.DateJSONPrefix="Date(";
	Setting.DateJSONSuffix=")";
	Excel.Setting=Setting;
	var _typeNamedItemCollection="NamedItemCollection";
	var NamedItemCollection=(function (_super) {
		__extends(NamedItemCollection, _super);
		function NamedItemCollection() {
			return _super !==null && _super.apply(this, arguments) || this;
		}
		Object.defineProperty(NamedItemCollection.prototype, "_className", {
			get: function () {
				return "NamedItemCollection";
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(NamedItemCollection.prototype, "items", {
			get: function () {
				_throwIfNotLoaded("items", this.m__items, _typeNamedItemCollection, this._isNull);
				return this.m__items;
			},
			enumerable: true,
			configurable: true
		});
		NamedItemCollection.prototype.add=function (name, reference, comment) {
			_throwIfApiNotSupported("NamedItemCollection.add", _defaultApiSetName, "1.4", _hostName);
			return new Excel.NamedItem(this.context, _createMethodObjectPath(this.context, this, "Add", 0, [name, reference, comment], false, true, null));
		};
		NamedItemCollection.prototype.addFormulaLocal=function (name, formula, comment) {
			_throwIfApiNotSupported("NamedItemCollection.addFormulaLocal", _defaultApiSetName, "1.4", _hostName);
			return new Excel.NamedItem(this.context, _createMethodObjectPath(this.context, this, "AddFormulaLocal", 0, [name, formula, comment], false, false, null));
		};
		NamedItemCollection.prototype.getCount=function () {
			_throwIfApiNotSupported("NamedItemCollection.getCount", _defaultApiSetName, "1.4", _hostName);
			var action=_createMethodAction(this.context, this, "GetCount", 1, []);
			var ret=new OfficeExtension.ClientResult();
			_addActionResultHandler(this, action, ret);
			return ret;
		};
		NamedItemCollection.prototype.getItem=function (name) {
			return new Excel.NamedItem(this.context, _createIndexerObjectPath(this.context, this, [name]));
		};
		NamedItemCollection.prototype.getItemOrNullObject=function (name) {
			_throwIfApiNotSupported("NamedItemCollection.getItemOrNullObject", _defaultApiSetName, "1.4", _hostName);
			return new Excel.NamedItem(this.context, _createMethodObjectPath(this.context, this, "GetItemOrNullObject", 1, [name], false, false, null));
		};
		NamedItemCollection.prototype._handleResult=function (value) {
			_super.prototype._handleResult.call(this, value);
			if (_isNullOrUndefined(value))
				return;
			var obj=value;
			_fixObjectPathIfNecessary(this, obj);
			if (!_isNullOrUndefined(obj[OfficeExtension.Constants.items])) {
				this.m__items=[];
				var _data=obj[OfficeExtension.Constants.items];
				for (var i=0; i < _data.length; i++) {
					var _item=new Excel.NamedItem(this.context, _createChildItemObjectPathUsingIndexerOrGetItemAt(true, this.context, this, _data[i], i));
					_item._handleResult(_data[i]);
					this.m__items.push(_item);
				}
			}
		};
		NamedItemCollection.prototype.load=function (option) {
			_load(this, option);
			return this;
		};
		NamedItemCollection.prototype.toJSON=function () {
			return {};
		};
		return NamedItemCollection;
	}(OfficeExtension.ClientObject));
	Excel.NamedItemCollection=NamedItemCollection;
	var _typeNamedItem="NamedItem";
	var NamedItem=(function (_super) {
		__extends(NamedItem, _super);
		function NamedItem() {
			return _super !==null && _super.apply(this, arguments) || this;
		}
		Object.defineProperty(NamedItem.prototype, "_className", {
			get: function () {
				return "NamedItem";
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(NamedItem.prototype, "worksheet", {
			get: function () {
				_throwIfApiNotSupported("NamedItem.worksheet", _defaultApiSetName, "1.4", _hostName);
				if (!this._W) {
					this._W=new Excel.Worksheet(this.context, _createPropertyObjectPath(this.context, this, "Worksheet", false, false));
				}
				return this._W;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(NamedItem.prototype, "worksheetOrNullObject", {
			get: function () {
				_throwIfApiNotSupported("NamedItem.worksheetOrNullObject", _defaultApiSetName, "1.4", _hostName);
				if (!this._Wo) {
					this._Wo=new Excel.Worksheet(this.context, _createPropertyObjectPath(this.context, this, "WorksheetOrNullObject", false, false));
				}
				return this._Wo;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(NamedItem.prototype, "comment", {
			get: function () {
				_throwIfNotLoaded("comment", this._C, _typeNamedItem, this._isNull);
				_throwIfApiNotSupported("NamedItem.comment", _defaultApiSetName, "1.4", _hostName);
				return this._C;
			},
			set: function (value) {
				this._C=value;
				_createSetPropertyAction(this.context, this, "Comment", value);
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(NamedItem.prototype, "name", {
			get: function () {
				_throwIfNotLoaded("name", this._N, _typeNamedItem, this._isNull);
				return this._N;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(NamedItem.prototype, "scope", {
			get: function () {
				_throwIfNotLoaded("scope", this._S, _typeNamedItem, this._isNull);
				_throwIfApiNotSupported("NamedItem.scope", _defaultApiSetName, "1.4", _hostName);
				return this._S;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(NamedItem.prototype, "type", {
			get: function () {
				_throwIfNotLoaded("type", this._T, _typeNamedItem, this._isNull);
				return this._T;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(NamedItem.prototype, "value", {
			get: function () {
				_throwIfNotLoaded("value", this._V, _typeNamedItem, this._isNull);
				return this._V;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(NamedItem.prototype, "visible", {
			get: function () {
				_throwIfNotLoaded("visible", this._Vi, _typeNamedItem, this._isNull);
				return this._Vi;
			},
			set: function (value) {
				this._Vi=value;
				_createSetPropertyAction(this.context, this, "Visible", value);
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(NamedItem.prototype, "_Id", {
			get: function () {
				_throwIfNotLoaded("_Id", this.__I, _typeNamedItem, this._isNull);
				return this.__I;
			},
			enumerable: true,
			configurable: true
		});
		NamedItem.prototype.set=function (properties, options) {
			this._recursivelySet(properties, options, ["visible", "comment"], [], [
				"worksheet",
				"worksheetOrNullObject",
				"worksheet",
				"worksheetOrNullObject"
			]);
		};
		NamedItem.prototype.delete=function () {
			_throwIfApiNotSupported("NamedItem.delete", _defaultApiSetName, "1.4", _hostName);
			_createMethodAction(this.context, this, "Delete", 0, []);
		};
		NamedItem.prototype.getRange=function () {
			return new Excel.Range(this.context, _createMethodObjectPath(this.context, this, "GetRange", 1, [], false, true, null));
		};
		NamedItem.prototype.getRangeOrNullObject=function () {
			_throwIfApiNotSupported("NamedItem.getRangeOrNullObject", _defaultApiSetName, "1.4", _hostName);
			return new Excel.Range(this.context, _createMethodObjectPath(this.context, this, "GetRangeOrNullObject", 1, [], false, true, null));
		};
		NamedItem.prototype._handleResult=function (value) {
			_super.prototype._handleResult.call(this, value);
			if (_isNullOrUndefined(value))
				return;
			var obj=value;
			_fixObjectPathIfNecessary(this, obj);
			if (!_isUndefined(obj["Comment"])) {
				this._C=obj["Comment"];
			}
			if (!_isUndefined(obj["Name"])) {
				this._N=obj["Name"];
			}
			if (!_isUndefined(obj["Scope"])) {
				this._S=obj["Scope"];
			}
			if (!_isUndefined(obj["Type"])) {
				this._T=obj["Type"];
			}
			if (!_isUndefined(obj["Value"])) {
				this._V=obj["Value"];
			}
			if (!_isUndefined(obj["Visible"])) {
				this._Vi=obj["Visible"];
			}
			if (!_isUndefined(obj["_Id"])) {
				this.__I=obj["_Id"];
			}
			_handleNavigationPropertyResults(this, obj, ["worksheet", "Worksheet", "worksheetOrNullObject", "WorksheetOrNullObject"]);
		};
		NamedItem.prototype.load=function (option) {
			_load(this, option);
			return this;
		};
		NamedItem.prototype._handleIdResult=function (value) {
			_super.prototype._handleIdResult.call(this, value);
			if (_isNullOrUndefined(value)) {
				return;
			}
			if (!_isUndefined(value["_Id"])) {
				this.__I=value["_Id"];
			}
		};
		NamedItem.prototype.toJSON=function () {
			return {
				"comment": this._C,
				"name": this._N,
				"scope": this._S,
				"type": this._T,
				"value": this._V,
				"visible": this._Vi
			};
		};
		return NamedItem;
	}(OfficeExtension.ClientObject));
	Excel.NamedItem=NamedItem;
	var _typeBinding="Binding";
	var Binding=(function (_super) {
		__extends(Binding, _super);
		function Binding() {
			return _super !==null && _super.apply(this, arguments) || this;
		}
		Object.defineProperty(Binding.prototype, "_className", {
			get: function () {
				return "Binding";
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Binding.prototype, "id", {
			get: function () {
				_throwIfNotLoaded("id", this._I, _typeBinding, this._isNull);
				return this._I;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Binding.prototype, "type", {
			get: function () {
				_throwIfNotLoaded("type", this._T, _typeBinding, this._isNull);
				return this._T;
			},
			enumerable: true,
			configurable: true
		});
		Binding.prototype.delete=function () {
			_throwIfApiNotSupported("Binding.delete", _defaultApiSetName, "1.3", _hostName);
			_createMethodAction(this.context, this, "Delete", 0, []);
		};
		Binding.prototype.getRange=function () {
			return new Excel.Range(this.context, _createMethodObjectPath(this.context, this, "GetRange", 1, [], false, false, null));
		};
		Binding.prototype.getTable=function () {
			return new Excel.Table(this.context, _createMethodObjectPath(this.context, this, "GetTable", 1, [], false, false, null));
		};
		Binding.prototype.getText=function () {
			var action=_createMethodAction(this.context, this, "GetText", 1, []);
			var ret=new OfficeExtension.ClientResult();
			_addActionResultHandler(this, action, ret);
			return ret;
		};
		Binding.prototype._handleResult=function (value) {
			_super.prototype._handleResult.call(this, value);
			if (_isNullOrUndefined(value))
				return;
			var obj=value;
			_fixObjectPathIfNecessary(this, obj);
			if (!_isUndefined(obj["Id"])) {
				this._I=obj["Id"];
			}
			if (!_isUndefined(obj["Type"])) {
				this._T=obj["Type"];
			}
		};
		Binding.prototype.load=function (option) {
			_load(this, option);
			return this;
		};
		Binding.prototype._handleIdResult=function (value) {
			_super.prototype._handleIdResult.call(this, value);
			if (_isNullOrUndefined(value)) {
				return;
			}
			if (!_isUndefined(value["Id"])) {
				this._I=value["Id"];
			}
		};
		Object.defineProperty(Binding.prototype, "onDataChanged", {
			get: function () {
				var _this=this;
				_throwIfApiNotSupported("Binding.onDataChanged", _defaultApiSetName, "1.3", _hostName);
				if (!this.m_dataChanged) {
					this.m_dataChanged=new OfficeExtension.EventHandlers(this.context, this, "DataChanged", {
						registerFunc: function (handlerCallback) {
							return _this.context.eventRegistration.register(4, _this.id, handlerCallback);
						},
						unregisterFunc: function (handlerCallback) {
							return _this.context.eventRegistration.unregister(4, _this.id, handlerCallback);
						},
						eventArgsTransformFunc: function (args) {
							var evt={
								binding: _this
							};
							return OfficeExtension.Utility._createPromiseFromResult(evt);
						}
					});
				}
				return this.m_dataChanged;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Binding.prototype, "onSelectionChanged", {
			get: function () {
				var _this=this;
				_throwIfApiNotSupported("Binding.onSelectionChanged", _defaultApiSetName, "1.3", _hostName);
				if (!this.m_selectionChanged) {
					this.m_selectionChanged=new OfficeExtension.EventHandlers(this.context, this, "SelectionChanged", {
						registerFunc: function (handlerCallback) {
							return _this.context.eventRegistration.register(3, _this.id, handlerCallback);
						},
						unregisterFunc: function (handlerCallback) {
							return _this.context.eventRegistration.unregister(3, _this.id, handlerCallback);
						},
						eventArgsTransformFunc: function (args) {
							var evt={
								binding: _this,
								columnCount: args.columnCount,
								rowCount: args.rowCount,
								startColumn: args.startColumn,
								startRow: args.startRow
							};
							return OfficeExtension.Utility._createPromiseFromResult(evt);
						}
					});
				}
				return this.m_selectionChanged;
			},
			enumerable: true,
			configurable: true
		});
		Binding.prototype.toJSON=function () {
			return {
				"id": this._I,
				"type": this._T
			};
		};
		return Binding;
	}(OfficeExtension.ClientObject));
	Excel.Binding=Binding;
	var _typeBindingCollection="BindingCollection";
	var BindingCollection=(function (_super) {
		__extends(BindingCollection, _super);
		function BindingCollection() {
			return _super !==null && _super.apply(this, arguments) || this;
		}
		Object.defineProperty(BindingCollection.prototype, "_className", {
			get: function () {
				return "BindingCollection";
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(BindingCollection.prototype, "items", {
			get: function () {
				_throwIfNotLoaded("items", this.m__items, _typeBindingCollection, this._isNull);
				return this.m__items;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(BindingCollection.prototype, "count", {
			get: function () {
				_throwIfNotLoaded("count", this._C, _typeBindingCollection, this._isNull);
				return this._C;
			},
			enumerable: true,
			configurable: true
		});
		BindingCollection.prototype.add=function (range, bindingType, id) {
			_throwIfApiNotSupported("BindingCollection.add", _defaultApiSetName, "1.3", _hostName);
			return new Excel.Binding(this.context, _createMethodObjectPath(this.context, this, "Add", 0, [range, bindingType, id], false, true, null));
		};
		BindingCollection.prototype.addFromNamedItem=function (name, bindingType, id) {
			_throwIfApiNotSupported("BindingCollection.addFromNamedItem", _defaultApiSetName, "1.3", _hostName);
			return new Excel.Binding(this.context, _createMethodObjectPath(this.context, this, "AddFromNamedItem", 0, [name, bindingType, id], false, false, null));
		};
		BindingCollection.prototype.addFromSelection=function (bindingType, id) {
			_throwIfApiNotSupported("BindingCollection.addFromSelection", _defaultApiSetName, "1.3", _hostName);
			return new Excel.Binding(this.context, _createMethodObjectPath(this.context, this, "AddFromSelection", 0, [bindingType, id], false, false, null));
		};
		BindingCollection.prototype.getCount=function () {
			_throwIfApiNotSupported("BindingCollection.getCount", _defaultApiSetName, "1.4", _hostName);
			var action=_createMethodAction(this.context, this, "GetCount", 1, []);
			var ret=new OfficeExtension.ClientResult();
			_addActionResultHandler(this, action, ret);
			return ret;
		};
		BindingCollection.prototype.getItem=function (id) {
			return new Excel.Binding(this.context, _createIndexerObjectPath(this.context, this, [id]));
		};
		BindingCollection.prototype.getItemAt=function (index) {
			return new Excel.Binding(this.context, _createMethodObjectPath(this.context, this, "GetItemAt", 1, [index], false, false, null));
		};
		BindingCollection.prototype.getItemOrNullObject=function (id) {
			_throwIfApiNotSupported("BindingCollection.getItemOrNullObject", _defaultApiSetName, "1.4", _hostName);
			return new Excel.Binding(this.context, _createMethodObjectPath(this.context, this, "GetItemOrNullObject", 1, [id], false, false, null));
		};
		BindingCollection.prototype._handleResult=function (value) {
			_super.prototype._handleResult.call(this, value);
			if (_isNullOrUndefined(value))
				return;
			var obj=value;
			_fixObjectPathIfNecessary(this, obj);
			if (!_isUndefined(obj["Count"])) {
				this._C=obj["Count"];
			}
			if (!_isNullOrUndefined(obj[OfficeExtension.Constants.items])) {
				this.m__items=[];
				var _data=obj[OfficeExtension.Constants.items];
				for (var i=0; i < _data.length; i++) {
					var _item=new Excel.Binding(this.context, _createChildItemObjectPathUsingIndexerOrGetItemAt(true, this.context, this, _data[i], i));
					_item._handleResult(_data[i]);
					this.m__items.push(_item);
				}
			}
		};
		BindingCollection.prototype.load=function (option) {
			_load(this, option);
			return this;
		};
		BindingCollection.prototype.toJSON=function () {
			return {
				"count": this._C
			};
		};
		return BindingCollection;
	}(OfficeExtension.ClientObject));
	Excel.BindingCollection=BindingCollection;
	var _typeTableCollection="TableCollection";
	var TableCollection=(function (_super) {
		__extends(TableCollection, _super);
		function TableCollection() {
			return _super !==null && _super.apply(this, arguments) || this;
		}
		Object.defineProperty(TableCollection.prototype, "_className", {
			get: function () {
				return "TableCollection";
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(TableCollection.prototype, "_ParentObject", {
			get: function () {
				return this.m__ParentObject;
			},
			set: function (value) {
				this.m__ParentObject=value;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(TableCollection.prototype, "items", {
			get: function () {
				_throwIfNotLoaded("items", this.m__items, _typeTableCollection, this._isNull);
				return this.m__items;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(TableCollection.prototype, "count", {
			get: function () {
				_throwIfNotLoaded("count", this._C, _typeTableCollection, this._isNull);
				return this._C;
			},
			enumerable: true,
			configurable: true
		});
		TableCollection.prototype.add=function (address, hasHeaders) {
			return new Excel.Table(this.context, _createMethodObjectPath(this.context, this, "Add", 0, [address, hasHeaders], false, true, null));
		};
		TableCollection.prototype.getCount=function () {
			_throwIfApiNotSupported("TableCollection.getCount", _defaultApiSetName, "1.4", _hostName);
			var action=_createMethodAction(this.context, this, "GetCount", 1, []);
			var ret=new OfficeExtension.ClientResult();
			_addActionResultHandler(this, action, ret);
			return ret;
		};
		TableCollection.prototype.getItem=function (key) {
			return new Excel.Table(this.context, _createIndexerObjectPath(this.context, this, [key]));
		};
		TableCollection.prototype.getItemAt=function (index) {
			return new Excel.Table(this.context, _createMethodObjectPath(this.context, this, "GetItemAt", 1, [index], false, false, null));
		};
		TableCollection.prototype.getItemOrNullObject=function (key) {
			_throwIfApiNotSupported("TableCollection.getItemOrNullObject", _defaultApiSetName, "1.4", _hostName);
			return new Excel.Table(this.context, _createMethodObjectPath(this.context, this, "GetItemOrNullObject", 1, [key], false, false, null));
		};
		TableCollection.prototype._handleResult=function (value) {
			_super.prototype._handleResult.call(this, value);
			if (_isNullOrUndefined(value))
				return;
			var obj=value;
			_fixObjectPathIfNecessary(this, obj);
			if (!_isUndefined(obj["Count"])) {
				this._C=obj["Count"];
			}
			if (!_isNullOrUndefined(obj[OfficeExtension.Constants.items])) {
				this.m__items=[];
				var _data=obj[OfficeExtension.Constants.items];
				for (var i=0; i < _data.length; i++) {
					var _item=new Excel.Table(this.context, _createChildItemObjectPathUsingIndexerOrGetItemAt(true, this.context, this, _data[i], i));
					_item._handleResult(_data[i]);
					this.m__items.push(_item);
				}
			}
		};
		TableCollection.prototype.load=function (option) {
			_load(this, option);
			return this;
		};
		TableCollection.prototype.toJSON=function () {
			return {
				"count": this._C
			};
		};
		return TableCollection;
	}(OfficeExtension.ClientObject));
	Excel.TableCollection=TableCollection;
	var _typeTable="Table";
	var Table=(function (_super) {
		__extends(Table, _super);
		function Table() {
			return _super !==null && _super.apply(this, arguments) || this;
		}
		Object.defineProperty(Table.prototype, "_className", {
			get: function () {
				return "Table";
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Table.prototype, "columns", {
			get: function () {
				if (!this._C) {
					this._C=new Excel.TableColumnCollection(this.context, _createPropertyObjectPath(this.context, this, "Columns", true, false));
				}
				return this._C;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Table.prototype, "rows", {
			get: function () {
				if (!this._R) {
					this._R=new Excel.TableRowCollection(this.context, _createPropertyObjectPath(this.context, this, "Rows", true, false));
				}
				return this._R;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Table.prototype, "sort", {
			get: function () {
				_throwIfApiNotSupported("Table.sort", _defaultApiSetName, "1.2", _hostName);
				if (!this._So) {
					this._So=new Excel.TableSort(this.context, _createPropertyObjectPath(this.context, this, "Sort", false, false));
				}
				return this._So;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Table.prototype, "worksheet", {
			get: function () {
				_throwIfApiNotSupported("Table.worksheet", _defaultApiSetName, "1.2", _hostName);
				if (!this._W) {
					this._W=new Excel.Worksheet(this.context, _createPropertyObjectPath(this.context, this, "Worksheet", false, false));
				}
				return this._W;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Table.prototype, "highlightFirstColumn", {
			get: function () {
				_throwIfNotLoaded("highlightFirstColumn", this._H, _typeTable, this._isNull);
				_throwIfApiNotSupported("Table.highlightFirstColumn", _defaultApiSetName, "1.3", _hostName);
				return this._H;
			},
			set: function (value) {
				this._H=value;
				_createSetPropertyAction(this.context, this, "HighlightFirstColumn", value);
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Table.prototype, "highlightLastColumn", {
			get: function () {
				_throwIfNotLoaded("highlightLastColumn", this._Hi, _typeTable, this._isNull);
				_throwIfApiNotSupported("Table.highlightLastColumn", _defaultApiSetName, "1.3", _hostName);
				return this._Hi;
			},
			set: function (value) {
				this._Hi=value;
				_createSetPropertyAction(this.context, this, "HighlightLastColumn", value);
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Table.prototype, "id", {
			get: function () {
				_throwIfNotLoaded("id", this._I, _typeTable, this._isNull);
				return this._I;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Table.prototype, "name", {
			get: function () {
				_throwIfNotLoaded("name", this._N, _typeTable, this._isNull);
				return this._N;
			},
			set: function (value) {
				this._N=value;
				_createSetPropertyAction(this.context, this, "Name", value);
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Table.prototype, "showBandedColumns", {
			get: function () {
				_throwIfNotLoaded("showBandedColumns", this._S, _typeTable, this._isNull);
				_throwIfApiNotSupported("Table.showBandedColumns", _defaultApiSetName, "1.3", _hostName);
				return this._S;
			},
			set: function (value) {
				this._S=value;
				_createSetPropertyAction(this.context, this, "ShowBandedColumns", value);
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Table.prototype, "showBandedRows", {
			get: function () {
				_throwIfNotLoaded("showBandedRows", this._Sh, _typeTable, this._isNull);
				_throwIfApiNotSupported("Table.showBandedRows", _defaultApiSetName, "1.3", _hostName);
				return this._Sh;
			},
			set: function (value) {
				this._Sh=value;
				_createSetPropertyAction(this.context, this, "ShowBandedRows", value);
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Table.prototype, "showFilterButton", {
			get: function () {
				_throwIfNotLoaded("showFilterButton", this._Sho, _typeTable, this._isNull);
				_throwIfApiNotSupported("Table.showFilterButton", _defaultApiSetName, "1.3", _hostName);
				return this._Sho;
			},
			set: function (value) {
				this._Sho=value;
				_createSetPropertyAction(this.context, this, "ShowFilterButton", value);
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Table.prototype, "showHeaders", {
			get: function () {
				_throwIfNotLoaded("showHeaders", this._Show, _typeTable, this._isNull);
				return this._Show;
			},
			set: function (value) {
				this._Show=value;
				_createSetPropertyAction(this.context, this, "ShowHeaders", value);
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Table.prototype, "showTotals", {
			get: function () {
				_throwIfNotLoaded("showTotals", this._ShowT, _typeTable, this._isNull);
				return this._ShowT;
			},
			set: function (value) {
				this._ShowT=value;
				_createSetPropertyAction(this.context, this, "ShowTotals", value);
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Table.prototype, "style", {
			get: function () {
				_throwIfNotLoaded("style", this._St, _typeTable, this._isNull);
				return this._St;
			},
			set: function (value) {
				this._St=value;
				_createSetPropertyAction(this.context, this, "Style", value);
			},
			enumerable: true,
			configurable: true
		});
		Table.prototype.set=function (properties, options) {
			this._recursivelySet(properties, options, ["name", "showHeaders", "showTotals", "style", "highlightFirstColumn", "highlightLastColumn", "showBandedRows", "showBandedColumns", "showFilterButton"], [], [
				"columns",
				"rows",
				"sort",
				"worksheet",
				"columns",
				"rows",
				"sort",
				"worksheet"
			]);
		};
		Table.prototype.clearFilters=function () {
			_throwIfApiNotSupported("Table.clearFilters", _defaultApiSetName, "1.2", _hostName);
			_createMethodAction(this.context, this, "ClearFilters", 0, []);
		};
		Table.prototype.convertToRange=function () {
			_throwIfApiNotSupported("Table.convertToRange", _defaultApiSetName, "1.2", _hostName);
			return new Excel.Range(this.context, _createMethodObjectPath(this.context, this, "ConvertToRange", 0, [], false, true, null));
		};
		Table.prototype.delete=function () {
			_createMethodAction(this.context, this, "Delete", 0, []);
		};
		Table.prototype.getDataBodyRange=function () {
			return new Excel.Range(this.context, _createMethodObjectPath(this.context, this, "GetDataBodyRange", 1, [], false, true, null));
		};
		Table.prototype.getHeaderRowRange=function () {
			return new Excel.Range(this.context, _createMethodObjectPath(this.context, this, "GetHeaderRowRange", 1, [], false, true, null));
		};
		Table.prototype.getRange=function () {
			return new Excel.Range(this.context, _createMethodObjectPath(this.context, this, "GetRange", 1, [], false, true, null));
		};
		Table.prototype.getTotalRowRange=function () {
			return new Excel.Range(this.context, _createMethodObjectPath(this.context, this, "GetTotalRowRange", 1, [], false, true, null));
		};
		Table.prototype.reapplyFilters=function () {
			_throwIfApiNotSupported("Table.reapplyFilters", _defaultApiSetName, "1.2", _hostName);
			_createMethodAction(this.context, this, "ReapplyFilters", 0, []);
		};
		Table.prototype._handleResult=function (value) {
			_super.prototype._handleResult.call(this, value);
			if (_isNullOrUndefined(value))
				return;
			var obj=value;
			_fixObjectPathIfNecessary(this, obj);
			if (!_isUndefined(obj["HighlightFirstColumn"])) {
				this._H=obj["HighlightFirstColumn"];
			}
			if (!_isUndefined(obj["HighlightLastColumn"])) {
				this._Hi=obj["HighlightLastColumn"];
			}
			if (!_isUndefined(obj["Id"])) {
				this._I=obj["Id"];
			}
			if (!_isUndefined(obj["Name"])) {
				this._N=obj["Name"];
			}
			if (!_isUndefined(obj["ShowBandedColumns"])) {
				this._S=obj["ShowBandedColumns"];
			}
			if (!_isUndefined(obj["ShowBandedRows"])) {
				this._Sh=obj["ShowBandedRows"];
			}
			if (!_isUndefined(obj["ShowFilterButton"])) {
				this._Sho=obj["ShowFilterButton"];
			}
			if (!_isUndefined(obj["ShowHeaders"])) {
				this._Show=obj["ShowHeaders"];
			}
			if (!_isUndefined(obj["ShowTotals"])) {
				this._ShowT=obj["ShowTotals"];
			}
			if (!_isUndefined(obj["Style"])) {
				this._St=obj["Style"];
			}
			_handleNavigationPropertyResults(this, obj, ["columns", "Columns", "rows", "Rows", "sort", "Sort", "worksheet", "Worksheet"]);
		};
		Table.prototype.load=function (option) {
			_load(this, option);
			return this;
		};
		Table.prototype._handleIdResult=function (value) {
			_super.prototype._handleIdResult.call(this, value);
			if (_isNullOrUndefined(value)) {
				return;
			}
			if (!_isUndefined(value["Id"])) {
				this._I=value["Id"];
			}
		};
		Table.prototype.toJSON=function () {
			return {
				"highlightFirstColumn": this._H,
				"highlightLastColumn": this._Hi,
				"id": this._I,
				"name": this._N,
				"showBandedColumns": this._S,
				"showBandedRows": this._Sh,
				"showFilterButton": this._Sho,
				"showHeaders": this._Show,
				"showTotals": this._ShowT,
				"style": this._St
			};
		};
		return Table;
	}(OfficeExtension.ClientObject));
	Excel.Table=Table;
	var _typeTableColumnCollection="TableColumnCollection";
	var TableColumnCollection=(function (_super) {
		__extends(TableColumnCollection, _super);
		function TableColumnCollection() {
			return _super !==null && _super.apply(this, arguments) || this;
		}
		Object.defineProperty(TableColumnCollection.prototype, "_className", {
			get: function () {
				return "TableColumnCollection";
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(TableColumnCollection.prototype, "items", {
			get: function () {
				_throwIfNotLoaded("items", this.m__items, _typeTableColumnCollection, this._isNull);
				return this.m__items;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(TableColumnCollection.prototype, "count", {
			get: function () {
				_throwIfNotLoaded("count", this._C, _typeTableColumnCollection, this._isNull);
				return this._C;
			},
			enumerable: true,
			configurable: true
		});
		TableColumnCollection.prototype.add=function (index, values, name) {
			return new Excel.TableColumn(this.context, _createMethodObjectPath(this.context, this, "Add", 0, [index, values, name], false, true, null));
		};
		TableColumnCollection.prototype.getCount=function () {
			_throwIfApiNotSupported("TableColumnCollection.getCount", _defaultApiSetName, "1.4", _hostName);
			var action=_createMethodAction(this.context, this, "GetCount", 1, []);
			var ret=new OfficeExtension.ClientResult();
			_addActionResultHandler(this, action, ret);
			return ret;
		};
		TableColumnCollection.prototype.getItem=function (key) {
			return new Excel.TableColumn(this.context, _createIndexerObjectPath(this.context, this, [key]));
		};
		TableColumnCollection.prototype.getItemAt=function (index) {
			return new Excel.TableColumn(this.context, _createMethodObjectPath(this.context, this, "GetItemAt", 1, [index], false, false, null));
		};
		TableColumnCollection.prototype.getItemOrNullObject=function (key) {
			_throwIfApiNotSupported("TableColumnCollection.getItemOrNullObject", _defaultApiSetName, "1.4", _hostName);
			return new Excel.TableColumn(this.context, _createMethodObjectPath(this.context, this, "GetItemOrNullObject", 1, [key], false, false, null));
		};
		TableColumnCollection.prototype._handleResult=function (value) {
			_super.prototype._handleResult.call(this, value);
			if (_isNullOrUndefined(value))
				return;
			var obj=value;
			_fixObjectPathIfNecessary(this, obj);
			if (!_isUndefined(obj["Count"])) {
				this._C=obj["Count"];
			}
			if (!_isNullOrUndefined(obj[OfficeExtension.Constants.items])) {
				this.m__items=[];
				var _data=obj[OfficeExtension.Constants.items];
				for (var i=0; i < _data.length; i++) {
					var _item=new Excel.TableColumn(this.context, _createChildItemObjectPathUsingIndexerOrGetItemAt(true, this.context, this, _data[i], i));
					_item._handleResult(_data[i]);
					this.m__items.push(_item);
				}
			}
		};
		TableColumnCollection.prototype.load=function (option) {
			_load(this, option);
			return this;
		};
		TableColumnCollection.prototype.toJSON=function () {
			return {
				"count": this._C
			};
		};
		return TableColumnCollection;
	}(OfficeExtension.ClientObject));
	Excel.TableColumnCollection=TableColumnCollection;
	var _typeTableColumn="TableColumn";
	var TableColumn=(function (_super) {
		__extends(TableColumn, _super);
		function TableColumn() {
			return _super !==null && _super.apply(this, arguments) || this;
		}
		Object.defineProperty(TableColumn.prototype, "_className", {
			get: function () {
				return "TableColumn";
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(TableColumn.prototype, "filter", {
			get: function () {
				_throwIfApiNotSupported("TableColumn.filter", _defaultApiSetName, "1.2", _hostName);
				if (!this._F) {
					this._F=new Excel.Filter(this.context, _createPropertyObjectPath(this.context, this, "Filter", false, false));
				}
				return this._F;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(TableColumn.prototype, "id", {
			get: function () {
				_throwIfNotLoaded("id", this._I, _typeTableColumn, this._isNull);
				return this._I;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(TableColumn.prototype, "index", {
			get: function () {
				_throwIfNotLoaded("index", this._In, _typeTableColumn, this._isNull);
				return this._In;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(TableColumn.prototype, "name", {
			get: function () {
				_throwIfNotLoaded("name", this._N, _typeTableColumn, this._isNull);
				return this._N;
			},
			set: function (value) {
				this._N=value;
				_createSetPropertyAction(this.context, this, "Name", value);
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(TableColumn.prototype, "values", {
			get: function () {
				_throwIfNotLoaded("values", this._V, _typeTableColumn, this._isNull);
				return this._V;
			},
			set: function (value) {
				this._V=value;
				_createSetPropertyAction(this.context, this, "Values", value);
			},
			enumerable: true,
			configurable: true
		});
		TableColumn.prototype.set=function (properties, options) {
			this._recursivelySet(properties, options, ["values", "name"], [], [
				"filter",
				"filter"
			]);
		};
		TableColumn.prototype.delete=function () {
			_createMethodAction(this.context, this, "Delete", 0, []);
		};
		TableColumn.prototype.getDataBodyRange=function () {
			return new Excel.Range(this.context, _createMethodObjectPath(this.context, this, "GetDataBodyRange", 1, [], false, true, null));
		};
		TableColumn.prototype.getHeaderRowRange=function () {
			return new Excel.Range(this.context, _createMethodObjectPath(this.context, this, "GetHeaderRowRange", 1, [], false, true, null));
		};
		TableColumn.prototype.getRange=function () {
			return new Excel.Range(this.context, _createMethodObjectPath(this.context, this, "GetRange", 1, [], false, true, null));
		};
		TableColumn.prototype.getTotalRowRange=function () {
			return new Excel.Range(this.context, _createMethodObjectPath(this.context, this, "GetTotalRowRange", 1, [], false, true, null));
		};
		TableColumn.prototype._handleResult=function (value) {
			_super.prototype._handleResult.call(this, value);
			if (_isNullOrUndefined(value))
				return;
			var obj=value;
			_fixObjectPathIfNecessary(this, obj);
			if (!_isUndefined(obj["Id"])) {
				this._I=obj["Id"];
			}
			if (!_isUndefined(obj["Index"])) {
				this._In=obj["Index"];
			}
			if (!_isUndefined(obj["Name"])) {
				this._N=obj["Name"];
			}
			if (!_isUndefined(obj["Values"])) {
				this._V=obj["Values"];
			}
			_handleNavigationPropertyResults(this, obj, ["filter", "Filter"]);
		};
		TableColumn.prototype.load=function (option) {
			_load(this, option);
			return this;
		};
		TableColumn.prototype._handleIdResult=function (value) {
			_super.prototype._handleIdResult.call(this, value);
			if (_isNullOrUndefined(value)) {
				return;
			}
			if (!_isUndefined(value["Id"])) {
				this._I=value["Id"];
			}
		};
		TableColumn.prototype.toJSON=function () {
			return {
				"id": this._I,
				"index": this._In,
				"name": this._N,
				"values": this._V
			};
		};
		return TableColumn;
	}(OfficeExtension.ClientObject));
	Excel.TableColumn=TableColumn;
	var _typeTableRowCollection="TableRowCollection";
	var TableRowCollection=(function (_super) {
		__extends(TableRowCollection, _super);
		function TableRowCollection() {
			return _super !==null && _super.apply(this, arguments) || this;
		}
		Object.defineProperty(TableRowCollection.prototype, "_className", {
			get: function () {
				return "TableRowCollection";
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(TableRowCollection.prototype, "items", {
			get: function () {
				_throwIfNotLoaded("items", this.m__items, _typeTableRowCollection, this._isNull);
				return this.m__items;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(TableRowCollection.prototype, "count", {
			get: function () {
				_throwIfNotLoaded("count", this._C, _typeTableRowCollection, this._isNull);
				return this._C;
			},
			enumerable: true,
			configurable: true
		});
		TableRowCollection.prototype.add=function (index, values) {
			return new Excel.TableRow(this.context, _createMethodObjectPath(this.context, this, "Add", 0, [index, values], false, true, null));
		};
		TableRowCollection.prototype.getCount=function () {
			_throwIfApiNotSupported("TableRowCollection.getCount", _defaultApiSetName, "1.4", _hostName);
			var action=_createMethodAction(this.context, this, "GetCount", 1, []);
			var ret=new OfficeExtension.ClientResult();
			_addActionResultHandler(this, action, ret);
			return ret;
		};
		TableRowCollection.prototype.getItemAt=function (index) {
			return new Excel.TableRow(this.context, _createMethodObjectPath(this.context, this, "GetItemAt", 1, [index], false, false, null));
		};
		TableRowCollection.prototype._handleResult=function (value) {
			_super.prototype._handleResult.call(this, value);
			if (_isNullOrUndefined(value))
				return;
			var obj=value;
			_fixObjectPathIfNecessary(this, obj);
			if (!_isUndefined(obj["Count"])) {
				this._C=obj["Count"];
			}
			if (!_isNullOrUndefined(obj[OfficeExtension.Constants.items])) {
				this.m__items=[];
				var _data=obj[OfficeExtension.Constants.items];
				for (var i=0; i < _data.length; i++) {
					var _item=new Excel.TableRow(this.context, _createChildItemObjectPathUsingIndexerOrGetItemAt(false, this.context, this, _data[i], i));
					_item._handleResult(_data[i]);
					this.m__items.push(_item);
				}
			}
		};
		TableRowCollection.prototype.load=function (option) {
			_load(this, option);
			return this;
		};
		TableRowCollection.prototype.toJSON=function () {
			return {
				"count": this._C
			};
		};
		return TableRowCollection;
	}(OfficeExtension.ClientObject));
	Excel.TableRowCollection=TableRowCollection;
	var _typeTableRow="TableRow";
	var TableRow=(function (_super) {
		__extends(TableRow, _super);
		function TableRow() {
			return _super !==null && _super.apply(this, arguments) || this;
		}
		Object.defineProperty(TableRow.prototype, "_className", {
			get: function () {
				return "TableRow";
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(TableRow.prototype, "index", {
			get: function () {
				_throwIfNotLoaded("index", this._I, _typeTableRow, this._isNull);
				return this._I;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(TableRow.prototype, "values", {
			get: function () {
				_throwIfNotLoaded("values", this._V, _typeTableRow, this._isNull);
				return this._V;
			},
			set: function (value) {
				this._V=value;
				_createSetPropertyAction(this.context, this, "Values", value);
			},
			enumerable: true,
			configurable: true
		});
		TableRow.prototype.set=function (properties, options) {
			this._recursivelySet(properties, options, ["values"], [], []);
		};
		TableRow.prototype.delete=function () {
			_createMethodAction(this.context, this, "Delete", 0, []);
		};
		TableRow.prototype.getRange=function () {
			return new Excel.Range(this.context, _createMethodObjectPath(this.context, this, "GetRange", 1, [], false, true, null));
		};
		TableRow.prototype._handleResult=function (value) {
			_super.prototype._handleResult.call(this, value);
			if (_isNullOrUndefined(value))
				return;
			var obj=value;
			_fixObjectPathIfNecessary(this, obj);
			if (!_isUndefined(obj["Index"])) {
				this._I=obj["Index"];
			}
			if (!_isUndefined(obj["Values"])) {
				this._V=obj["Values"];
			}
		};
		TableRow.prototype.load=function (option) {
			_load(this, option);
			return this;
		};
		TableRow.prototype.toJSON=function () {
			return {
				"index": this._I,
				"values": this._V
			};
		};
		return TableRow;
	}(OfficeExtension.ClientObject));
	Excel.TableRow=TableRow;
	var _typeRangeFormat="RangeFormat";
	var RangeFormat=(function (_super) {
		__extends(RangeFormat, _super);
		function RangeFormat() {
			return _super !==null && _super.apply(this, arguments) || this;
		}
		Object.defineProperty(RangeFormat.prototype, "_className", {
			get: function () {
				return "RangeFormat";
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(RangeFormat.prototype, "borders", {
			get: function () {
				if (!this._B) {
					this._B=new Excel.RangeBorderCollection(this.context, _createPropertyObjectPath(this.context, this, "Borders", true, false));
				}
				return this._B;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(RangeFormat.prototype, "fill", {
			get: function () {
				if (!this._F) {
					this._F=new Excel.RangeFill(this.context, _createPropertyObjectPath(this.context, this, "Fill", false, false));
				}
				return this._F;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(RangeFormat.prototype, "font", {
			get: function () {
				if (!this._Fo) {
					this._Fo=new Excel.RangeFont(this.context, _createPropertyObjectPath(this.context, this, "Font", false, false));
				}
				return this._Fo;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(RangeFormat.prototype, "protection", {
			get: function () {
				_throwIfApiNotSupported("RangeFormat.protection", _defaultApiSetName, "1.2", _hostName);
				if (!this._P) {
					this._P=new Excel.FormatProtection(this.context, _createPropertyObjectPath(this.context, this, "Protection", false, false));
				}
				return this._P;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(RangeFormat.prototype, "columnWidth", {
			get: function () {
				_throwIfNotLoaded("columnWidth", this._C, _typeRangeFormat, this._isNull);
				_throwIfApiNotSupported("RangeFormat.columnWidth", _defaultApiSetName, "1.2", _hostName);
				return this._C;
			},
			set: function (value) {
				this._C=value;
				_createSetPropertyAction(this.context, this, "ColumnWidth", value);
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(RangeFormat.prototype, "horizontalAlignment", {
			get: function () {
				_throwIfNotLoaded("horizontalAlignment", this._H, _typeRangeFormat, this._isNull);
				return this._H;
			},
			set: function (value) {
				this._H=value;
				_createSetPropertyAction(this.context, this, "HorizontalAlignment", value);
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(RangeFormat.prototype, "rowHeight", {
			get: function () {
				_throwIfNotLoaded("rowHeight", this._R, _typeRangeFormat, this._isNull);
				_throwIfApiNotSupported("RangeFormat.rowHeight", _defaultApiSetName, "1.2", _hostName);
				return this._R;
			},
			set: function (value) {
				this._R=value;
				_createSetPropertyAction(this.context, this, "RowHeight", value);
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(RangeFormat.prototype, "verticalAlignment", {
			get: function () {
				_throwIfNotLoaded("verticalAlignment", this._V, _typeRangeFormat, this._isNull);
				return this._V;
			},
			set: function (value) {
				this._V=value;
				_createSetPropertyAction(this.context, this, "VerticalAlignment", value);
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(RangeFormat.prototype, "wrapText", {
			get: function () {
				_throwIfNotLoaded("wrapText", this._W, _typeRangeFormat, this._isNull);
				return this._W;
			},
			set: function (value) {
				this._W=value;
				_createSetPropertyAction(this.context, this, "WrapText", value);
			},
			enumerable: true,
			configurable: true
		});
		RangeFormat.prototype.set=function (properties, options) {
			this._recursivelySet(properties, options, ["wrapText", "horizontalAlignment", "verticalAlignment", "columnWidth", "rowHeight"], ["fill", "font", "protection"], [
				"borders",
				"borders"
			]);
		};
		RangeFormat.prototype.autofitColumns=function () {
			_throwIfApiNotSupported("RangeFormat.autofitColumns", _defaultApiSetName, "1.2", _hostName);
			_createMethodAction(this.context, this, "AutofitColumns", 0, []);
		};
		RangeFormat.prototype.autofitRows=function () {
			_throwIfApiNotSupported("RangeFormat.autofitRows", _defaultApiSetName, "1.2", _hostName);
			_createMethodAction(this.context, this, "AutofitRows", 0, []);
		};
		RangeFormat.prototype._handleResult=function (value) {
			_super.prototype._handleResult.call(this, value);
			if (_isNullOrUndefined(value))
				return;
			var obj=value;
			_fixObjectPathIfNecessary(this, obj);
			if (!_isUndefined(obj["ColumnWidth"])) {
				this._C=obj["ColumnWidth"];
			}
			if (!_isUndefined(obj["HorizontalAlignment"])) {
				this._H=obj["HorizontalAlignment"];
			}
			if (!_isUndefined(obj["RowHeight"])) {
				this._R=obj["RowHeight"];
			}
			if (!_isUndefined(obj["VerticalAlignment"])) {
				this._V=obj["VerticalAlignment"];
			}
			if (!_isUndefined(obj["WrapText"])) {
				this._W=obj["WrapText"];
			}
			_handleNavigationPropertyResults(this, obj, ["borders", "Borders", "fill", "Fill", "font", "Font", "protection", "Protection"]);
		};
		RangeFormat.prototype.load=function (option) {
			_load(this, option);
			return this;
		};
		RangeFormat.prototype.toJSON=function () {
			return {
				"columnWidth": this._C,
				"fill": this._F,
				"font": this._Fo,
				"horizontalAlignment": this._H,
				"protection": this._P,
				"rowHeight": this._R,
				"verticalAlignment": this._V,
				"wrapText": this._W
			};
		};
		return RangeFormat;
	}(OfficeExtension.ClientObject));
	Excel.RangeFormat=RangeFormat;
	var _typeFormatProtection="FormatProtection";
	var FormatProtection=(function (_super) {
		__extends(FormatProtection, _super);
		function FormatProtection() {
			return _super !==null && _super.apply(this, arguments) || this;
		}
		Object.defineProperty(FormatProtection.prototype, "_className", {
			get: function () {
				return "FormatProtection";
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(FormatProtection.prototype, "formulaHidden", {
			get: function () {
				_throwIfNotLoaded("formulaHidden", this._F, _typeFormatProtection, this._isNull);
				return this._F;
			},
			set: function (value) {
				this._F=value;
				_createSetPropertyAction(this.context, this, "FormulaHidden", value);
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(FormatProtection.prototype, "locked", {
			get: function () {
				_throwIfNotLoaded("locked", this._L, _typeFormatProtection, this._isNull);
				return this._L;
			},
			set: function (value) {
				this._L=value;
				_createSetPropertyAction(this.context, this, "Locked", value);
			},
			enumerable: true,
			configurable: true
		});
		FormatProtection.prototype.set=function (properties, options) {
			this._recursivelySet(properties, options, ["locked", "formulaHidden"], [], []);
		};
		FormatProtection.prototype._handleResult=function (value) {
			_super.prototype._handleResult.call(this, value);
			if (_isNullOrUndefined(value))
				return;
			var obj=value;
			_fixObjectPathIfNecessary(this, obj);
			if (!_isUndefined(obj["FormulaHidden"])) {
				this._F=obj["FormulaHidden"];
			}
			if (!_isUndefined(obj["Locked"])) {
				this._L=obj["Locked"];
			}
		};
		FormatProtection.prototype.load=function (option) {
			_load(this, option);
			return this;
		};
		FormatProtection.prototype.toJSON=function () {
			return {
				"formulaHidden": this._F,
				"locked": this._L
			};
		};
		return FormatProtection;
	}(OfficeExtension.ClientObject));
	Excel.FormatProtection=FormatProtection;
	var _typeRangeFill="RangeFill";
	var RangeFill=(function (_super) {
		__extends(RangeFill, _super);
		function RangeFill() {
			return _super !==null && _super.apply(this, arguments) || this;
		}
		Object.defineProperty(RangeFill.prototype, "_className", {
			get: function () {
				return "RangeFill";
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(RangeFill.prototype, "color", {
			get: function () {
				_throwIfNotLoaded("color", this._C, _typeRangeFill, this._isNull);
				return this._C;
			},
			set: function (value) {
				this._C=value;
				_createSetPropertyAction(this.context, this, "Color", value);
			},
			enumerable: true,
			configurable: true
		});
		RangeFill.prototype.set=function (properties, options) {
			this._recursivelySet(properties, options, ["color"], [], []);
		};
		RangeFill.prototype.clear=function () {
			_createMethodAction(this.context, this, "Clear", 0, []);
		};
		RangeFill.prototype._handleResult=function (value) {
			_super.prototype._handleResult.call(this, value);
			if (_isNullOrUndefined(value))
				return;
			var obj=value;
			_fixObjectPathIfNecessary(this, obj);
			if (!_isUndefined(obj["Color"])) {
				this._C=obj["Color"];
			}
		};
		RangeFill.prototype.load=function (option) {
			_load(this, option);
			return this;
		};
		RangeFill.prototype.toJSON=function () {
			return {
				"color": this._C
			};
		};
		return RangeFill;
	}(OfficeExtension.ClientObject));
	Excel.RangeFill=RangeFill;
	var _typeRangeBorder="RangeBorder";
	var RangeBorder=(function (_super) {
		__extends(RangeBorder, _super);
		function RangeBorder() {
			return _super !==null && _super.apply(this, arguments) || this;
		}
		Object.defineProperty(RangeBorder.prototype, "_className", {
			get: function () {
				return "RangeBorder";
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(RangeBorder.prototype, "color", {
			get: function () {
				_throwIfNotLoaded("color", this._C, _typeRangeBorder, this._isNull);
				return this._C;
			},
			set: function (value) {
				this._C=value;
				_createSetPropertyAction(this.context, this, "Color", value);
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(RangeBorder.prototype, "sideIndex", {
			get: function () {
				_throwIfNotLoaded("sideIndex", this._S, _typeRangeBorder, this._isNull);
				return this._S;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(RangeBorder.prototype, "style", {
			get: function () {
				_throwIfNotLoaded("style", this._St, _typeRangeBorder, this._isNull);
				return this._St;
			},
			set: function (value) {
				this._St=value;
				_createSetPropertyAction(this.context, this, "Style", value);
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(RangeBorder.prototype, "weight", {
			get: function () {
				_throwIfNotLoaded("weight", this._W, _typeRangeBorder, this._isNull);
				return this._W;
			},
			set: function (value) {
				this._W=value;
				_createSetPropertyAction(this.context, this, "Weight", value);
			},
			enumerable: true,
			configurable: true
		});
		RangeBorder.prototype.set=function (properties, options) {
			this._recursivelySet(properties, options, ["style", "weight", "color"], [], []);
		};
		RangeBorder.prototype._handleResult=function (value) {
			_super.prototype._handleResult.call(this, value);
			if (_isNullOrUndefined(value))
				return;
			var obj=value;
			_fixObjectPathIfNecessary(this, obj);
			if (!_isUndefined(obj["Color"])) {
				this._C=obj["Color"];
			}
			if (!_isUndefined(obj["SideIndex"])) {
				this._S=obj["SideIndex"];
			}
			if (!_isUndefined(obj["Style"])) {
				this._St=obj["Style"];
			}
			if (!_isUndefined(obj["Weight"])) {
				this._W=obj["Weight"];
			}
		};
		RangeBorder.prototype.load=function (option) {
			_load(this, option);
			return this;
		};
		RangeBorder.prototype.toJSON=function () {
			return {
				"color": this._C,
				"sideIndex": this._S,
				"style": this._St,
				"weight": this._W
			};
		};
		return RangeBorder;
	}(OfficeExtension.ClientObject));
	Excel.RangeBorder=RangeBorder;
	var _typeRangeBorderCollection="RangeBorderCollection";
	var RangeBorderCollection=(function (_super) {
		__extends(RangeBorderCollection, _super);
		function RangeBorderCollection() {
			return _super !==null && _super.apply(this, arguments) || this;
		}
		Object.defineProperty(RangeBorderCollection.prototype, "_className", {
			get: function () {
				return "RangeBorderCollection";
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(RangeBorderCollection.prototype, "items", {
			get: function () {
				_throwIfNotLoaded("items", this.m__items, _typeRangeBorderCollection, this._isNull);
				return this.m__items;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(RangeBorderCollection.prototype, "count", {
			get: function () {
				_throwIfNotLoaded("count", this._C, _typeRangeBorderCollection, this._isNull);
				return this._C;
			},
			enumerable: true,
			configurable: true
		});
		RangeBorderCollection.prototype.getItem=function (index) {
			return new Excel.RangeBorder(this.context, _createIndexerObjectPath(this.context, this, [index]));
		};
		RangeBorderCollection.prototype.getItemAt=function (index) {
			return new Excel.RangeBorder(this.context, _createMethodObjectPath(this.context, this, "GetItemAt", 1, [index], false, false, null));
		};
		RangeBorderCollection.prototype._handleResult=function (value) {
			_super.prototype._handleResult.call(this, value);
			if (_isNullOrUndefined(value))
				return;
			var obj=value;
			_fixObjectPathIfNecessary(this, obj);
			if (!_isUndefined(obj["Count"])) {
				this._C=obj["Count"];
			}
			if (!_isNullOrUndefined(obj[OfficeExtension.Constants.items])) {
				this.m__items=[];
				var _data=obj[OfficeExtension.Constants.items];
				for (var i=0; i < _data.length; i++) {
					var _item=new Excel.RangeBorder(this.context, _createChildItemObjectPathUsingIndexerOrGetItemAt(true, this.context, this, _data[i], i));
					_item._handleResult(_data[i]);
					this.m__items.push(_item);
				}
			}
		};
		RangeBorderCollection.prototype.load=function (option) {
			_load(this, option);
			return this;
		};
		RangeBorderCollection.prototype.toJSON=function () {
			return {
				"count": this._C
			};
		};
		return RangeBorderCollection;
	}(OfficeExtension.ClientObject));
	Excel.RangeBorderCollection=RangeBorderCollection;
	var _typeRangeFont="RangeFont";
	var RangeFont=(function (_super) {
		__extends(RangeFont, _super);
		function RangeFont() {
			return _super !==null && _super.apply(this, arguments) || this;
		}
		Object.defineProperty(RangeFont.prototype, "_className", {
			get: function () {
				return "RangeFont";
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(RangeFont.prototype, "bold", {
			get: function () {
				_throwIfNotLoaded("bold", this._B, _typeRangeFont, this._isNull);
				return this._B;
			},
			set: function (value) {
				this._B=value;
				_createSetPropertyAction(this.context, this, "Bold", value);
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(RangeFont.prototype, "color", {
			get: function () {
				_throwIfNotLoaded("color", this._C, _typeRangeFont, this._isNull);
				return this._C;
			},
			set: function (value) {
				this._C=value;
				_createSetPropertyAction(this.context, this, "Color", value);
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(RangeFont.prototype, "italic", {
			get: function () {
				_throwIfNotLoaded("italic", this._I, _typeRangeFont, this._isNull);
				return this._I;
			},
			set: function (value) {
				this._I=value;
				_createSetPropertyAction(this.context, this, "Italic", value);
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(RangeFont.prototype, "name", {
			get: function () {
				_throwIfNotLoaded("name", this._N, _typeRangeFont, this._isNull);
				return this._N;
			},
			set: function (value) {
				this._N=value;
				_createSetPropertyAction(this.context, this, "Name", value);
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(RangeFont.prototype, "size", {
			get: function () {
				_throwIfNotLoaded("size", this._S, _typeRangeFont, this._isNull);
				return this._S;
			},
			set: function (value) {
				this._S=value;
				_createSetPropertyAction(this.context, this, "Size", value);
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(RangeFont.prototype, "underline", {
			get: function () {
				_throwIfNotLoaded("underline", this._U, _typeRangeFont, this._isNull);
				return this._U;
			},
			set: function (value) {
				this._U=value;
				_createSetPropertyAction(this.context, this, "Underline", value);
			},
			enumerable: true,
			configurable: true
		});
		RangeFont.prototype.set=function (properties, options) {
			this._recursivelySet(properties, options, ["name", "size", "color", "italic", "bold", "underline"], [], []);
		};
		RangeFont.prototype._handleResult=function (value) {
			_super.prototype._handleResult.call(this, value);
			if (_isNullOrUndefined(value))
				return;
			var obj=value;
			_fixObjectPathIfNecessary(this, obj);
			if (!_isUndefined(obj["Bold"])) {
				this._B=obj["Bold"];
			}
			if (!_isUndefined(obj["Color"])) {
				this._C=obj["Color"];
			}
			if (!_isUndefined(obj["Italic"])) {
				this._I=obj["Italic"];
			}
			if (!_isUndefined(obj["Name"])) {
				this._N=obj["Name"];
			}
			if (!_isUndefined(obj["Size"])) {
				this._S=obj["Size"];
			}
			if (!_isUndefined(obj["Underline"])) {
				this._U=obj["Underline"];
			}
		};
		RangeFont.prototype.load=function (option) {
			_load(this, option);
			return this;
		};
		RangeFont.prototype.toJSON=function () {
			return {
				"bold": this._B,
				"color": this._C,
				"italic": this._I,
				"name": this._N,
				"size": this._S,
				"underline": this._U
			};
		};
		return RangeFont;
	}(OfficeExtension.ClientObject));
	Excel.RangeFont=RangeFont;
	var _typeChartCollection="ChartCollection";
	var ChartCollection=(function (_super) {
		__extends(ChartCollection, _super);
		function ChartCollection() {
			return _super !==null && _super.apply(this, arguments) || this;
		}
		Object.defineProperty(ChartCollection.prototype, "_className", {
			get: function () {
				return "ChartCollection";
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ChartCollection.prototype, "items", {
			get: function () {
				_throwIfNotLoaded("items", this.m__items, _typeChartCollection, this._isNull);
				return this.m__items;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ChartCollection.prototype, "count", {
			get: function () {
				_throwIfNotLoaded("count", this._C, _typeChartCollection, this._isNull);
				return this._C;
			},
			enumerable: true,
			configurable: true
		});
		ChartCollection.prototype.add=function (type, sourceData, seriesBy) {
			if (!(sourceData instanceof Range)) {
				throw OfficeExtension.Utility.createRuntimeError(OfficeExtension.ResourceStrings.invalidArgument, "sourceData", "Charts.Add");
			}
			return new Excel.Chart(this.context, _createMethodObjectPath(this.context, this, "Add", 0, [type, sourceData, seriesBy], false, true, null));
		};
		ChartCollection.prototype.getCount=function () {
			_throwIfApiNotSupported("ChartCollection.getCount", _defaultApiSetName, "1.4", _hostName);
			var action=_createMethodAction(this.context, this, "GetCount", 1, []);
			var ret=new OfficeExtension.ClientResult();
			_addActionResultHandler(this, action, ret);
			return ret;
		};
		ChartCollection.prototype.getItem=function (name) {
			return new Excel.Chart(this.context, _createMethodObjectPath(this.context, this, "GetItem", 1, [name], false, false, null));
		};
		ChartCollection.prototype.getItemAt=function (index) {
			return new Excel.Chart(this.context, _createMethodObjectPath(this.context, this, "GetItemAt", 1, [index], false, false, null));
		};
		ChartCollection.prototype.getItemOrNullObject=function (name) {
			_throwIfApiNotSupported("ChartCollection.getItemOrNullObject", _defaultApiSetName, "1.4", _hostName);
			return new Excel.Chart(this.context, _createMethodObjectPath(this.context, this, "GetItemOrNullObject", 1, [name], false, false, null));
		};
		ChartCollection.prototype._GetItem=function (key) {
			return new Excel.Chart(this.context, _createIndexerObjectPath(this.context, this, [key]));
		};
		ChartCollection.prototype._handleResult=function (value) {
			_super.prototype._handleResult.call(this, value);
			if (_isNullOrUndefined(value))
				return;
			var obj=value;
			_fixObjectPathIfNecessary(this, obj);
			if (!_isUndefined(obj["Count"])) {
				this._C=obj["Count"];
			}
			if (!_isNullOrUndefined(obj[OfficeExtension.Constants.items])) {
				this.m__items=[];
				var _data=obj[OfficeExtension.Constants.items];
				for (var i=0; i < _data.length; i++) {
					var _item=new Excel.Chart(this.context, _createChildItemObjectPathUsingIndexerOrGetItemAt(true, this.context, this, _data[i], i));
					_item._handleResult(_data[i]);
					this.m__items.push(_item);
				}
			}
		};
		ChartCollection.prototype.load=function (option) {
			_load(this, option);
			return this;
		};
		ChartCollection.prototype.toJSON=function () {
			return {
				"count": this._C
			};
		};
		return ChartCollection;
	}(OfficeExtension.ClientObject));
	Excel.ChartCollection=ChartCollection;
	var _typeChart="Chart";
	var Chart=(function (_super) {
		__extends(Chart, _super);
		function Chart() {
			return _super !==null && _super.apply(this, arguments) || this;
		}
		Object.defineProperty(Chart.prototype, "_className", {
			get: function () {
				return "Chart";
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Chart.prototype, "axes", {
			get: function () {
				if (!this._A) {
					this._A=new Excel.ChartAxes(this.context, _createPropertyObjectPath(this.context, this, "Axes", false, false));
				}
				return this._A;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Chart.prototype, "dataLabels", {
			get: function () {
				if (!this._D) {
					this._D=new Excel.ChartDataLabels(this.context, _createPropertyObjectPath(this.context, this, "DataLabels", false, false));
				}
				return this._D;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Chart.prototype, "format", {
			get: function () {
				if (!this._F) {
					this._F=new Excel.ChartAreaFormat(this.context, _createPropertyObjectPath(this.context, this, "Format", false, false));
				}
				return this._F;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Chart.prototype, "legend", {
			get: function () {
				if (!this._Le) {
					this._Le=new Excel.ChartLegend(this.context, _createPropertyObjectPath(this.context, this, "Legend", false, false));
				}
				return this._Le;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Chart.prototype, "series", {
			get: function () {
				if (!this._S) {
					this._S=new Excel.ChartSeriesCollection(this.context, _createPropertyObjectPath(this.context, this, "Series", true, false));
				}
				return this._S;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Chart.prototype, "title", {
			get: function () {
				if (!this._T) {
					this._T=new Excel.ChartTitle(this.context, _createPropertyObjectPath(this.context, this, "Title", false, false));
				}
				return this._T;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Chart.prototype, "worksheet", {
			get: function () {
				_throwIfApiNotSupported("Chart.worksheet", _defaultApiSetName, "1.2", _hostName);
				if (!this._Wo) {
					this._Wo=new Excel.Worksheet(this.context, _createPropertyObjectPath(this.context, this, "Worksheet", false, false));
				}
				return this._Wo;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Chart.prototype, "height", {
			get: function () {
				_throwIfNotLoaded("height", this._H, _typeChart, this._isNull);
				return this._H;
			},
			set: function (value) {
				this._H=value;
				_createSetPropertyAction(this.context, this, "Height", value);
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Chart.prototype, "left", {
			get: function () {
				_throwIfNotLoaded("left", this._L, _typeChart, this._isNull);
				return this._L;
			},
			set: function (value) {
				this._L=value;
				_createSetPropertyAction(this.context, this, "Left", value);
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Chart.prototype, "name", {
			get: function () {
				_throwIfNotLoaded("name", this._N, _typeChart, this._isNull);
				return this._N;
			},
			set: function (value) {
				this._N=value;
				_createSetPropertyAction(this.context, this, "Name", value);
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Chart.prototype, "top", {
			get: function () {
				_throwIfNotLoaded("top", this._To, _typeChart, this._isNull);
				return this._To;
			},
			set: function (value) {
				this._To=value;
				_createSetPropertyAction(this.context, this, "Top", value);
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Chart.prototype, "width", {
			get: function () {
				_throwIfNotLoaded("width", this._W, _typeChart, this._isNull);
				return this._W;
			},
			set: function (value) {
				this._W=value;
				_createSetPropertyAction(this.context, this, "Width", value);
			},
			enumerable: true,
			configurable: true
		});
		Chart.prototype.set=function (properties, options) {
			this._recursivelySet(properties, options, ["name", "top", "left", "width", "height"], ["title", "dataLabels", "legend", "axes", "format"], [
				"series",
				"worksheet",
				"series",
				"worksheet"
			]);
		};
		Chart.prototype.delete=function () {
			_createMethodAction(this.context, this, "Delete", 0, []);
		};
		Chart.prototype.getImage=function (width, height, fittingMode) {
			_throwIfApiNotSupported("Chart.getImage", _defaultApiSetName, "1.2", _hostName);
			var action=_createMethodAction(this.context, this, "GetImage", 1, [width, height, fittingMode]);
			var ret=new OfficeExtension.ClientResult();
			_addActionResultHandler(this, action, ret);
			return ret;
		};
		Chart.prototype.setData=function (sourceData, seriesBy) {
			if (!(sourceData instanceof Range)) {
				throw OfficeExtension.Utility.createRuntimeError(OfficeExtension.ResourceStrings.invalidArgument, "sourceData", "Chart.setData");
			}
			_createMethodAction(this.context, this, "SetData", 0, [sourceData, seriesBy]);
		};
		Chart.prototype.setPosition=function (startCell, endCell) {
			_createMethodAction(this.context, this, "SetPosition", 0, [startCell, endCell]);
		};
		Chart.prototype._handleResult=function (value) {
			_super.prototype._handleResult.call(this, value);
			if (_isNullOrUndefined(value))
				return;
			var obj=value;
			_fixObjectPathIfNecessary(this, obj);
			if (!_isUndefined(obj["Height"])) {
				this._H=obj["Height"];
			}
			if (!_isUndefined(obj["Left"])) {
				this._L=obj["Left"];
			}
			if (!_isUndefined(obj["Name"])) {
				this._N=obj["Name"];
			}
			if (!_isUndefined(obj["Top"])) {
				this._To=obj["Top"];
			}
			if (!_isUndefined(obj["Width"])) {
				this._W=obj["Width"];
			}
			_handleNavigationPropertyResults(this, obj, ["axes", "Axes", "dataLabels", "DataLabels", "format", "Format", "legend", "Legend", "series", "Series", "title", "Title", "worksheet", "Worksheet"]);
		};
		Chart.prototype.load=function (option) {
			_load(this, option);
			return this;
		};
		Chart.prototype.toJSON=function () {
			return {
				"axes": this._A,
				"dataLabels": this._D,
				"format": this._F,
				"height": this._H,
				"left": this._L,
				"legend": this._Le,
				"name": this._N,
				"title": this._T,
				"top": this._To,
				"width": this._W
			};
		};
		return Chart;
	}(OfficeExtension.ClientObject));
	Excel.Chart=Chart;
	var _typeChartAreaFormat="ChartAreaFormat";
	var ChartAreaFormat=(function (_super) {
		__extends(ChartAreaFormat, _super);
		function ChartAreaFormat() {
			return _super !==null && _super.apply(this, arguments) || this;
		}
		Object.defineProperty(ChartAreaFormat.prototype, "_className", {
			get: function () {
				return "ChartAreaFormat";
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ChartAreaFormat.prototype, "fill", {
			get: function () {
				if (!this._F) {
					this._F=new Excel.ChartFill(this.context, _createPropertyObjectPath(this.context, this, "Fill", false, false));
				}
				return this._F;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ChartAreaFormat.prototype, "font", {
			get: function () {
				if (!this._Fo) {
					this._Fo=new Excel.ChartFont(this.context, _createPropertyObjectPath(this.context, this, "Font", false, false));
				}
				return this._Fo;
			},
			enumerable: true,
			configurable: true
		});
		ChartAreaFormat.prototype.set=function (properties, options) {
			this._recursivelySet(properties, options, [], ["font"], [
				"fill"
			]);
		};
		ChartAreaFormat.prototype._handleResult=function (value) {
			_super.prototype._handleResult.call(this, value);
			if (_isNullOrUndefined(value))
				return;
			var obj=value;
			_fixObjectPathIfNecessary(this, obj);
			_handleNavigationPropertyResults(this, obj, ["fill", "Fill", "font", "Font"]);
		};
		ChartAreaFormat.prototype.load=function (option) {
			_load(this, option);
			return this;
		};
		ChartAreaFormat.prototype.toJSON=function () {
			return {
				"fill": this._F,
				"font": this._Fo
			};
		};
		return ChartAreaFormat;
	}(OfficeExtension.ClientObject));
	Excel.ChartAreaFormat=ChartAreaFormat;
	var _typeChartSeriesCollection="ChartSeriesCollection";
	var ChartSeriesCollection=(function (_super) {
		__extends(ChartSeriesCollection, _super);
		function ChartSeriesCollection() {
			return _super !==null && _super.apply(this, arguments) || this;
		}
		Object.defineProperty(ChartSeriesCollection.prototype, "_className", {
			get: function () {
				return "ChartSeriesCollection";
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ChartSeriesCollection.prototype, "items", {
			get: function () {
				_throwIfNotLoaded("items", this.m__items, _typeChartSeriesCollection, this._isNull);
				return this.m__items;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ChartSeriesCollection.prototype, "count", {
			get: function () {
				_throwIfNotLoaded("count", this._C, _typeChartSeriesCollection, this._isNull);
				return this._C;
			},
			enumerable: true,
			configurable: true
		});
		ChartSeriesCollection.prototype.getCount=function () {
			_throwIfApiNotSupported("ChartSeriesCollection.getCount", _defaultApiSetName, "1.4", _hostName);
			var action=_createMethodAction(this.context, this, "GetCount", 1, []);
			var ret=new OfficeExtension.ClientResult();
			_addActionResultHandler(this, action, ret);
			return ret;
		};
		ChartSeriesCollection.prototype.getItemAt=function (index) {
			return new Excel.ChartSeries(this.context, _createMethodObjectPath(this.context, this, "GetItemAt", 1, [index], false, false, null));
		};
		ChartSeriesCollection.prototype._handleResult=function (value) {
			_super.prototype._handleResult.call(this, value);
			if (_isNullOrUndefined(value))
				return;
			var obj=value;
			_fixObjectPathIfNecessary(this, obj);
			if (!_isUndefined(obj["Count"])) {
				this._C=obj["Count"];
			}
			if (!_isNullOrUndefined(obj[OfficeExtension.Constants.items])) {
				this.m__items=[];
				var _data=obj[OfficeExtension.Constants.items];
				for (var i=0; i < _data.length; i++) {
					var _item=new Excel.ChartSeries(this.context, _createChildItemObjectPathUsingIndexerOrGetItemAt(false, this.context, this, _data[i], i));
					_item._handleResult(_data[i]);
					this.m__items.push(_item);
				}
			}
		};
		ChartSeriesCollection.prototype.load=function (option) {
			_load(this, option);
			return this;
		};
		ChartSeriesCollection.prototype.toJSON=function () {
			return {
				"count": this._C
			};
		};
		return ChartSeriesCollection;
	}(OfficeExtension.ClientObject));
	Excel.ChartSeriesCollection=ChartSeriesCollection;
	var _typeChartSeries="ChartSeries";
	var ChartSeries=(function (_super) {
		__extends(ChartSeries, _super);
		function ChartSeries() {
			return _super !==null && _super.apply(this, arguments) || this;
		}
		Object.defineProperty(ChartSeries.prototype, "_className", {
			get: function () {
				return "ChartSeries";
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ChartSeries.prototype, "format", {
			get: function () {
				if (!this._F) {
					this._F=new Excel.ChartSeriesFormat(this.context, _createPropertyObjectPath(this.context, this, "Format", false, false));
				}
				return this._F;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ChartSeries.prototype, "points", {
			get: function () {
				if (!this._P) {
					this._P=new Excel.ChartPointsCollection(this.context, _createPropertyObjectPath(this.context, this, "Points", true, false));
				}
				return this._P;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ChartSeries.prototype, "name", {
			get: function () {
				_throwIfNotLoaded("name", this._N, _typeChartSeries, this._isNull);
				return this._N;
			},
			set: function (value) {
				this._N=value;
				_createSetPropertyAction(this.context, this, "Name", value);
			},
			enumerable: true,
			configurable: true
		});
		ChartSeries.prototype.set=function (properties, options) {
			this._recursivelySet(properties, options, ["name"], ["format"], [
				"points",
				"points"
			]);
		};
		ChartSeries.prototype._handleResult=function (value) {
			_super.prototype._handleResult.call(this, value);
			if (_isNullOrUndefined(value))
				return;
			var obj=value;
			_fixObjectPathIfNecessary(this, obj);
			if (!_isUndefined(obj["Name"])) {
				this._N=obj["Name"];
			}
			_handleNavigationPropertyResults(this, obj, ["format", "Format", "points", "Points"]);
		};
		ChartSeries.prototype.load=function (option) {
			_load(this, option);
			return this;
		};
		ChartSeries.prototype.toJSON=function () {
			return {
				"format": this._F,
				"name": this._N
			};
		};
		return ChartSeries;
	}(OfficeExtension.ClientObject));
	Excel.ChartSeries=ChartSeries;
	var _typeChartSeriesFormat="ChartSeriesFormat";
	var ChartSeriesFormat=(function (_super) {
		__extends(ChartSeriesFormat, _super);
		function ChartSeriesFormat() {
			return _super !==null && _super.apply(this, arguments) || this;
		}
		Object.defineProperty(ChartSeriesFormat.prototype, "_className", {
			get: function () {
				return "ChartSeriesFormat";
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ChartSeriesFormat.prototype, "fill", {
			get: function () {
				if (!this._F) {
					this._F=new Excel.ChartFill(this.context, _createPropertyObjectPath(this.context, this, "Fill", false, false));
				}
				return this._F;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ChartSeriesFormat.prototype, "line", {
			get: function () {
				if (!this._L) {
					this._L=new Excel.ChartLineFormat(this.context, _createPropertyObjectPath(this.context, this, "Line", false, false));
				}
				return this._L;
			},
			enumerable: true,
			configurable: true
		});
		ChartSeriesFormat.prototype.set=function (properties, options) {
			this._recursivelySet(properties, options, [], ["line"], [
				"fill"
			]);
		};
		ChartSeriesFormat.prototype._handleResult=function (value) {
			_super.prototype._handleResult.call(this, value);
			if (_isNullOrUndefined(value))
				return;
			var obj=value;
			_fixObjectPathIfNecessary(this, obj);
			_handleNavigationPropertyResults(this, obj, ["fill", "Fill", "line", "Line"]);
		};
		ChartSeriesFormat.prototype.load=function (option) {
			_load(this, option);
			return this;
		};
		ChartSeriesFormat.prototype.toJSON=function () {
			return {
				"fill": this._F,
				"line": this._L
			};
		};
		return ChartSeriesFormat;
	}(OfficeExtension.ClientObject));
	Excel.ChartSeriesFormat=ChartSeriesFormat;
	var _typeChartPointsCollection="ChartPointsCollection";
	var ChartPointsCollection=(function (_super) {
		__extends(ChartPointsCollection, _super);
		function ChartPointsCollection() {
			return _super !==null && _super.apply(this, arguments) || this;
		}
		Object.defineProperty(ChartPointsCollection.prototype, "_className", {
			get: function () {
				return "ChartPointsCollection";
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ChartPointsCollection.prototype, "items", {
			get: function () {
				_throwIfNotLoaded("items", this.m__items, _typeChartPointsCollection, this._isNull);
				return this.m__items;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ChartPointsCollection.prototype, "count", {
			get: function () {
				_throwIfNotLoaded("count", this._C, _typeChartPointsCollection, this._isNull);
				return this._C;
			},
			enumerable: true,
			configurable: true
		});
		ChartPointsCollection.prototype.getCount=function () {
			_throwIfApiNotSupported("ChartPointsCollection.getCount", _defaultApiSetName, "1.4", _hostName);
			var action=_createMethodAction(this.context, this, "GetCount", 1, []);
			var ret=new OfficeExtension.ClientResult();
			_addActionResultHandler(this, action, ret);
			return ret;
		};
		ChartPointsCollection.prototype.getItemAt=function (index) {
			return new Excel.ChartPoint(this.context, _createMethodObjectPath(this.context, this, "GetItemAt", 1, [index], false, false, null));
		};
		ChartPointsCollection.prototype._handleResult=function (value) {
			_super.prototype._handleResult.call(this, value);
			if (_isNullOrUndefined(value))
				return;
			var obj=value;
			_fixObjectPathIfNecessary(this, obj);
			if (!_isUndefined(obj["Count"])) {
				this._C=obj["Count"];
			}
			if (!_isNullOrUndefined(obj[OfficeExtension.Constants.items])) {
				this.m__items=[];
				var _data=obj[OfficeExtension.Constants.items];
				for (var i=0; i < _data.length; i++) {
					var _item=new Excel.ChartPoint(this.context, _createChildItemObjectPathUsingIndexerOrGetItemAt(false, this.context, this, _data[i], i));
					_item._handleResult(_data[i]);
					this.m__items.push(_item);
				}
			}
		};
		ChartPointsCollection.prototype.load=function (option) {
			_load(this, option);
			return this;
		};
		ChartPointsCollection.prototype.toJSON=function () {
			return {
				"count": this._C
			};
		};
		return ChartPointsCollection;
	}(OfficeExtension.ClientObject));
	Excel.ChartPointsCollection=ChartPointsCollection;
	var _typeChartPoint="ChartPoint";
	var ChartPoint=(function (_super) {
		__extends(ChartPoint, _super);
		function ChartPoint() {
			return _super !==null && _super.apply(this, arguments) || this;
		}
		Object.defineProperty(ChartPoint.prototype, "_className", {
			get: function () {
				return "ChartPoint";
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ChartPoint.prototype, "format", {
			get: function () {
				if (!this._F) {
					this._F=new Excel.ChartPointFormat(this.context, _createPropertyObjectPath(this.context, this, "Format", false, false));
				}
				return this._F;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ChartPoint.prototype, "value", {
			get: function () {
				_throwIfNotLoaded("value", this._V, _typeChartPoint, this._isNull);
				return this._V;
			},
			enumerable: true,
			configurable: true
		});
		ChartPoint.prototype._handleResult=function (value) {
			_super.prototype._handleResult.call(this, value);
			if (_isNullOrUndefined(value))
				return;
			var obj=value;
			_fixObjectPathIfNecessary(this, obj);
			if (!_isUndefined(obj["Value"])) {
				this._V=obj["Value"];
			}
			_handleNavigationPropertyResults(this, obj, ["format", "Format"]);
		};
		ChartPoint.prototype.load=function (option) {
			_load(this, option);
			return this;
		};
		ChartPoint.prototype.toJSON=function () {
			return {
				"format": this._F,
				"value": this._V
			};
		};
		return ChartPoint;
	}(OfficeExtension.ClientObject));
	Excel.ChartPoint=ChartPoint;
	var _typeChartPointFormat="ChartPointFormat";
	var ChartPointFormat=(function (_super) {
		__extends(ChartPointFormat, _super);
		function ChartPointFormat() {
			return _super !==null && _super.apply(this, arguments) || this;
		}
		Object.defineProperty(ChartPointFormat.prototype, "_className", {
			get: function () {
				return "ChartPointFormat";
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ChartPointFormat.prototype, "fill", {
			get: function () {
				if (!this._F) {
					this._F=new Excel.ChartFill(this.context, _createPropertyObjectPath(this.context, this, "Fill", false, false));
				}
				return this._F;
			},
			enumerable: true,
			configurable: true
		});
		ChartPointFormat.prototype._handleResult=function (value) {
			_super.prototype._handleResult.call(this, value);
			if (_isNullOrUndefined(value))
				return;
			var obj=value;
			_fixObjectPathIfNecessary(this, obj);
			_handleNavigationPropertyResults(this, obj, ["fill", "Fill"]);
		};
		ChartPointFormat.prototype.load=function (option) {
			_load(this, option);
			return this;
		};
		ChartPointFormat.prototype.toJSON=function () {
			return {
				"fill": this._F
			};
		};
		return ChartPointFormat;
	}(OfficeExtension.ClientObject));
	Excel.ChartPointFormat=ChartPointFormat;
	var _typeChartAxes="ChartAxes";
	var ChartAxes=(function (_super) {
		__extends(ChartAxes, _super);
		function ChartAxes() {
			return _super !==null && _super.apply(this, arguments) || this;
		}
		Object.defineProperty(ChartAxes.prototype, "_className", {
			get: function () {
				return "ChartAxes";
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ChartAxes.prototype, "categoryAxis", {
			get: function () {
				if (!this._C) {
					this._C=new Excel.ChartAxis(this.context, _createPropertyObjectPath(this.context, this, "CategoryAxis", false, false));
				}
				return this._C;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ChartAxes.prototype, "seriesAxis", {
			get: function () {
				if (!this._S) {
					this._S=new Excel.ChartAxis(this.context, _createPropertyObjectPath(this.context, this, "SeriesAxis", false, false));
				}
				return this._S;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ChartAxes.prototype, "valueAxis", {
			get: function () {
				if (!this._V) {
					this._V=new Excel.ChartAxis(this.context, _createPropertyObjectPath(this.context, this, "ValueAxis", false, false));
				}
				return this._V;
			},
			enumerable: true,
			configurable: true
		});
		ChartAxes.prototype.set=function (properties, options) {
			this._recursivelySet(properties, options, [], ["categoryAxis", "seriesAxis", "valueAxis"], []);
		};
		ChartAxes.prototype._handleResult=function (value) {
			_super.prototype._handleResult.call(this, value);
			if (_isNullOrUndefined(value))
				return;
			var obj=value;
			_fixObjectPathIfNecessary(this, obj);
			_handleNavigationPropertyResults(this, obj, ["categoryAxis", "CategoryAxis", "seriesAxis", "SeriesAxis", "valueAxis", "ValueAxis"]);
		};
		ChartAxes.prototype.load=function (option) {
			_load(this, option);
			return this;
		};
		ChartAxes.prototype.toJSON=function () {
			return {
				"categoryAxis": this._C,
				"seriesAxis": this._S,
				"valueAxis": this._V
			};
		};
		return ChartAxes;
	}(OfficeExtension.ClientObject));
	Excel.ChartAxes=ChartAxes;
	var _typeChartAxis="ChartAxis";
	var ChartAxis=(function (_super) {
		__extends(ChartAxis, _super);
		function ChartAxis() {
			return _super !==null && _super.apply(this, arguments) || this;
		}
		Object.defineProperty(ChartAxis.prototype, "_className", {
			get: function () {
				return "ChartAxis";
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ChartAxis.prototype, "format", {
			get: function () {
				if (!this._F) {
					this._F=new Excel.ChartAxisFormat(this.context, _createPropertyObjectPath(this.context, this, "Format", false, false));
				}
				return this._F;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ChartAxis.prototype, "majorGridlines", {
			get: function () {
				if (!this._M) {
					this._M=new Excel.ChartGridlines(this.context, _createPropertyObjectPath(this.context, this, "MajorGridlines", false, false));
				}
				return this._M;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ChartAxis.prototype, "minorGridlines", {
			get: function () {
				if (!this._Min) {
					this._Min=new Excel.ChartGridlines(this.context, _createPropertyObjectPath(this.context, this, "MinorGridlines", false, false));
				}
				return this._Min;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ChartAxis.prototype, "title", {
			get: function () {
				if (!this._T) {
					this._T=new Excel.ChartAxisTitle(this.context, _createPropertyObjectPath(this.context, this, "Title", false, false));
				}
				return this._T;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ChartAxis.prototype, "majorUnit", {
			get: function () {
				_throwIfNotLoaded("majorUnit", this._Ma, _typeChartAxis, this._isNull);
				return this._Ma;
			},
			set: function (value) {
				this._Ma=value;
				_createSetPropertyAction(this.context, this, "MajorUnit", value);
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ChartAxis.prototype, "maximum", {
			get: function () {
				_throwIfNotLoaded("maximum", this._Max, _typeChartAxis, this._isNull);
				return this._Max;
			},
			set: function (value) {
				this._Max=value;
				_createSetPropertyAction(this.context, this, "Maximum", value);
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ChartAxis.prototype, "minimum", {
			get: function () {
				_throwIfNotLoaded("minimum", this._Mi, _typeChartAxis, this._isNull);
				return this._Mi;
			},
			set: function (value) {
				this._Mi=value;
				_createSetPropertyAction(this.context, this, "Minimum", value);
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ChartAxis.prototype, "minorUnit", {
			get: function () {
				_throwIfNotLoaded("minorUnit", this._Mino, _typeChartAxis, this._isNull);
				return this._Mino;
			},
			set: function (value) {
				this._Mino=value;
				_createSetPropertyAction(this.context, this, "MinorUnit", value);
			},
			enumerable: true,
			configurable: true
		});
		ChartAxis.prototype.set=function (properties, options) {
			this._recursivelySet(properties, options, ["majorUnit", "maximum", "minimum", "minorUnit"], ["majorGridlines", "minorGridlines", "title", "format"], []);
		};
		ChartAxis.prototype._handleResult=function (value) {
			_super.prototype._handleResult.call(this, value);
			if (_isNullOrUndefined(value))
				return;
			var obj=value;
			_fixObjectPathIfNecessary(this, obj);
			if (!_isUndefined(obj["MajorUnit"])) {
				this._Ma=obj["MajorUnit"];
			}
			if (!_isUndefined(obj["Maximum"])) {
				this._Max=obj["Maximum"];
			}
			if (!_isUndefined(obj["Minimum"])) {
				this._Mi=obj["Minimum"];
			}
			if (!_isUndefined(obj["MinorUnit"])) {
				this._Mino=obj["MinorUnit"];
			}
			_handleNavigationPropertyResults(this, obj, ["format", "Format", "majorGridlines", "MajorGridlines", "minorGridlines", "MinorGridlines", "title", "Title"]);
		};
		ChartAxis.prototype.load=function (option) {
			_load(this, option);
			return this;
		};
		ChartAxis.prototype.toJSON=function () {
			return {
				"format": this._F,
				"majorGridlines": this._M,
				"majorUnit": this._Ma,
				"maximum": this._Max,
				"minimum": this._Mi,
				"minorGridlines": this._Min,
				"minorUnit": this._Mino,
				"title": this._T
			};
		};
		return ChartAxis;
	}(OfficeExtension.ClientObject));
	Excel.ChartAxis=ChartAxis;
	var _typeChartAxisFormat="ChartAxisFormat";
	var ChartAxisFormat=(function (_super) {
		__extends(ChartAxisFormat, _super);
		function ChartAxisFormat() {
			return _super !==null && _super.apply(this, arguments) || this;
		}
		Object.defineProperty(ChartAxisFormat.prototype, "_className", {
			get: function () {
				return "ChartAxisFormat";
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ChartAxisFormat.prototype, "font", {
			get: function () {
				if (!this._F) {
					this._F=new Excel.ChartFont(this.context, _createPropertyObjectPath(this.context, this, "Font", false, false));
				}
				return this._F;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ChartAxisFormat.prototype, "line", {
			get: function () {
				if (!this._L) {
					this._L=new Excel.ChartLineFormat(this.context, _createPropertyObjectPath(this.context, this, "Line", false, false));
				}
				return this._L;
			},
			enumerable: true,
			configurable: true
		});
		ChartAxisFormat.prototype.set=function (properties, options) {
			this._recursivelySet(properties, options, [], ["font", "line"], []);
		};
		ChartAxisFormat.prototype._handleResult=function (value) {
			_super.prototype._handleResult.call(this, value);
			if (_isNullOrUndefined(value))
				return;
			var obj=value;
			_fixObjectPathIfNecessary(this, obj);
			_handleNavigationPropertyResults(this, obj, ["font", "Font", "line", "Line"]);
		};
		ChartAxisFormat.prototype.load=function (option) {
			_load(this, option);
			return this;
		};
		ChartAxisFormat.prototype.toJSON=function () {
			return {
				"font": this._F,
				"line": this._L
			};
		};
		return ChartAxisFormat;
	}(OfficeExtension.ClientObject));
	Excel.ChartAxisFormat=ChartAxisFormat;
	var _typeChartAxisTitle="ChartAxisTitle";
	var ChartAxisTitle=(function (_super) {
		__extends(ChartAxisTitle, _super);
		function ChartAxisTitle() {
			return _super !==null && _super.apply(this, arguments) || this;
		}
		Object.defineProperty(ChartAxisTitle.prototype, "_className", {
			get: function () {
				return "ChartAxisTitle";
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ChartAxisTitle.prototype, "format", {
			get: function () {
				if (!this._F) {
					this._F=new Excel.ChartAxisTitleFormat(this.context, _createPropertyObjectPath(this.context, this, "Format", false, false));
				}
				return this._F;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ChartAxisTitle.prototype, "text", {
			get: function () {
				_throwIfNotLoaded("text", this._T, _typeChartAxisTitle, this._isNull);
				return this._T;
			},
			set: function (value) {
				this._T=value;
				_createSetPropertyAction(this.context, this, "Text", value);
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ChartAxisTitle.prototype, "visible", {
			get: function () {
				_throwIfNotLoaded("visible", this._V, _typeChartAxisTitle, this._isNull);
				return this._V;
			},
			set: function (value) {
				this._V=value;
				_createSetPropertyAction(this.context, this, "Visible", value);
			},
			enumerable: true,
			configurable: true
		});
		ChartAxisTitle.prototype.set=function (properties, options) {
			this._recursivelySet(properties, options, ["text", "visible"], ["format"], []);
		};
		ChartAxisTitle.prototype._handleResult=function (value) {
			_super.prototype._handleResult.call(this, value);
			if (_isNullOrUndefined(value))
				return;
			var obj=value;
			_fixObjectPathIfNecessary(this, obj);
			if (!_isUndefined(obj["Text"])) {
				this._T=obj["Text"];
			}
			if (!_isUndefined(obj["Visible"])) {
				this._V=obj["Visible"];
			}
			_handleNavigationPropertyResults(this, obj, ["format", "Format"]);
		};
		ChartAxisTitle.prototype.load=function (option) {
			_load(this, option);
			return this;
		};
		ChartAxisTitle.prototype.toJSON=function () {
			return {
				"format": this._F,
				"text": this._T,
				"visible": this._V
			};
		};
		return ChartAxisTitle;
	}(OfficeExtension.ClientObject));
	Excel.ChartAxisTitle=ChartAxisTitle;
	var _typeChartAxisTitleFormat="ChartAxisTitleFormat";
	var ChartAxisTitleFormat=(function (_super) {
		__extends(ChartAxisTitleFormat, _super);
		function ChartAxisTitleFormat() {
			return _super !==null && _super.apply(this, arguments) || this;
		}
		Object.defineProperty(ChartAxisTitleFormat.prototype, "_className", {
			get: function () {
				return "ChartAxisTitleFormat";
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ChartAxisTitleFormat.prototype, "font", {
			get: function () {
				if (!this._F) {
					this._F=new Excel.ChartFont(this.context, _createPropertyObjectPath(this.context, this, "Font", false, false));
				}
				return this._F;
			},
			enumerable: true,
			configurable: true
		});
		ChartAxisTitleFormat.prototype.set=function (properties, options) {
			this._recursivelySet(properties, options, [], ["font"], []);
		};
		ChartAxisTitleFormat.prototype._handleResult=function (value) {
			_super.prototype._handleResult.call(this, value);
			if (_isNullOrUndefined(value))
				return;
			var obj=value;
			_fixObjectPathIfNecessary(this, obj);
			_handleNavigationPropertyResults(this, obj, ["font", "Font"]);
		};
		ChartAxisTitleFormat.prototype.load=function (option) {
			_load(this, option);
			return this;
		};
		ChartAxisTitleFormat.prototype.toJSON=function () {
			return {
				"font": this._F
			};
		};
		return ChartAxisTitleFormat;
	}(OfficeExtension.ClientObject));
	Excel.ChartAxisTitleFormat=ChartAxisTitleFormat;
	var _typeChartDataLabels="ChartDataLabels";
	var ChartDataLabels=(function (_super) {
		__extends(ChartDataLabels, _super);
		function ChartDataLabels() {
			return _super !==null && _super.apply(this, arguments) || this;
		}
		Object.defineProperty(ChartDataLabels.prototype, "_className", {
			get: function () {
				return "ChartDataLabels";
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ChartDataLabels.prototype, "format", {
			get: function () {
				if (!this._F) {
					this._F=new Excel.ChartDataLabelFormat(this.context, _createPropertyObjectPath(this.context, this, "Format", false, false));
				}
				return this._F;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ChartDataLabels.prototype, "position", {
			get: function () {
				_throwIfNotLoaded("position", this._P, _typeChartDataLabels, this._isNull);
				return this._P;
			},
			set: function (value) {
				this._P=value;
				_createSetPropertyAction(this.context, this, "Position", value);
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ChartDataLabels.prototype, "separator", {
			get: function () {
				_throwIfNotLoaded("separator", this._S, _typeChartDataLabels, this._isNull);
				return this._S;
			},
			set: function (value) {
				this._S=value;
				_createSetPropertyAction(this.context, this, "Separator", value);
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ChartDataLabels.prototype, "showBubbleSize", {
			get: function () {
				_throwIfNotLoaded("showBubbleSize", this._Sh, _typeChartDataLabels, this._isNull);
				return this._Sh;
			},
			set: function (value) {
				this._Sh=value;
				_createSetPropertyAction(this.context, this, "ShowBubbleSize", value);
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ChartDataLabels.prototype, "showCategoryName", {
			get: function () {
				_throwIfNotLoaded("showCategoryName", this._Sho, _typeChartDataLabels, this._isNull);
				return this._Sho;
			},
			set: function (value) {
				this._Sho=value;
				_createSetPropertyAction(this.context, this, "ShowCategoryName", value);
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ChartDataLabels.prototype, "showLegendKey", {
			get: function () {
				_throwIfNotLoaded("showLegendKey", this._Show, _typeChartDataLabels, this._isNull);
				return this._Show;
			},
			set: function (value) {
				this._Show=value;
				_createSetPropertyAction(this.context, this, "ShowLegendKey", value);
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ChartDataLabels.prototype, "showPercentage", {
			get: function () {
				_throwIfNotLoaded("showPercentage", this._ShowP, _typeChartDataLabels, this._isNull);
				return this._ShowP;
			},
			set: function (value) {
				this._ShowP=value;
				_createSetPropertyAction(this.context, this, "ShowPercentage", value);
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ChartDataLabels.prototype, "showSeriesName", {
			get: function () {
				_throwIfNotLoaded("showSeriesName", this._ShowS, _typeChartDataLabels, this._isNull);
				return this._ShowS;
			},
			set: function (value) {
				this._ShowS=value;
				_createSetPropertyAction(this.context, this, "ShowSeriesName", value);
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ChartDataLabels.prototype, "showValue", {
			get: function () {
				_throwIfNotLoaded("showValue", this._ShowV, _typeChartDataLabels, this._isNull);
				return this._ShowV;
			},
			set: function (value) {
				this._ShowV=value;
				_createSetPropertyAction(this.context, this, "ShowValue", value);
			},
			enumerable: true,
			configurable: true
		});
		ChartDataLabels.prototype.set=function (properties, options) {
			this._recursivelySet(properties, options, ["position", "showValue", "showSeriesName", "showCategoryName", "showLegendKey", "showPercentage", "showBubbleSize", "separator"], ["format"], []);
		};
		ChartDataLabels.prototype._handleResult=function (value) {
			_super.prototype._handleResult.call(this, value);
			if (_isNullOrUndefined(value))
				return;
			var obj=value;
			_fixObjectPathIfNecessary(this, obj);
			if (!_isUndefined(obj["Position"])) {
				this._P=obj["Position"];
			}
			if (!_isUndefined(obj["Separator"])) {
				this._S=obj["Separator"];
			}
			if (!_isUndefined(obj["ShowBubbleSize"])) {
				this._Sh=obj["ShowBubbleSize"];
			}
			if (!_isUndefined(obj["ShowCategoryName"])) {
				this._Sho=obj["ShowCategoryName"];
			}
			if (!_isUndefined(obj["ShowLegendKey"])) {
				this._Show=obj["ShowLegendKey"];
			}
			if (!_isUndefined(obj["ShowPercentage"])) {
				this._ShowP=obj["ShowPercentage"];
			}
			if (!_isUndefined(obj["ShowSeriesName"])) {
				this._ShowS=obj["ShowSeriesName"];
			}
			if (!_isUndefined(obj["ShowValue"])) {
				this._ShowV=obj["ShowValue"];
			}
			_handleNavigationPropertyResults(this, obj, ["format", "Format"]);
		};
		ChartDataLabels.prototype.load=function (option) {
			_load(this, option);
			return this;
		};
		ChartDataLabels.prototype.toJSON=function () {
			return {
				"format": this._F,
				"position": this._P,
				"separator": this._S,
				"showBubbleSize": this._Sh,
				"showCategoryName": this._Sho,
				"showLegendKey": this._Show,
				"showPercentage": this._ShowP,
				"showSeriesName": this._ShowS,
				"showValue": this._ShowV
			};
		};
		return ChartDataLabels;
	}(OfficeExtension.ClientObject));
	Excel.ChartDataLabels=ChartDataLabels;
	var _typeChartDataLabelFormat="ChartDataLabelFormat";
	var ChartDataLabelFormat=(function (_super) {
		__extends(ChartDataLabelFormat, _super);
		function ChartDataLabelFormat() {
			return _super !==null && _super.apply(this, arguments) || this;
		}
		Object.defineProperty(ChartDataLabelFormat.prototype, "_className", {
			get: function () {
				return "ChartDataLabelFormat";
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ChartDataLabelFormat.prototype, "fill", {
			get: function () {
				if (!this._F) {
					this._F=new Excel.ChartFill(this.context, _createPropertyObjectPath(this.context, this, "Fill", false, false));
				}
				return this._F;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ChartDataLabelFormat.prototype, "font", {
			get: function () {
				if (!this._Fo) {
					this._Fo=new Excel.ChartFont(this.context, _createPropertyObjectPath(this.context, this, "Font", false, false));
				}
				return this._Fo;
			},
			enumerable: true,
			configurable: true
		});
		ChartDataLabelFormat.prototype.set=function (properties, options) {
			this._recursivelySet(properties, options, [], ["font"], [
				"fill"
			]);
		};
		ChartDataLabelFormat.prototype._handleResult=function (value) {
			_super.prototype._handleResult.call(this, value);
			if (_isNullOrUndefined(value))
				return;
			var obj=value;
			_fixObjectPathIfNecessary(this, obj);
			_handleNavigationPropertyResults(this, obj, ["fill", "Fill", "font", "Font"]);
		};
		ChartDataLabelFormat.prototype.load=function (option) {
			_load(this, option);
			return this;
		};
		ChartDataLabelFormat.prototype.toJSON=function () {
			return {
				"fill": this._F,
				"font": this._Fo
			};
		};
		return ChartDataLabelFormat;
	}(OfficeExtension.ClientObject));
	Excel.ChartDataLabelFormat=ChartDataLabelFormat;
	var _typeChartGridlines="ChartGridlines";
	var ChartGridlines=(function (_super) {
		__extends(ChartGridlines, _super);
		function ChartGridlines() {
			return _super !==null && _super.apply(this, arguments) || this;
		}
		Object.defineProperty(ChartGridlines.prototype, "_className", {
			get: function () {
				return "ChartGridlines";
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ChartGridlines.prototype, "format", {
			get: function () {
				if (!this._F) {
					this._F=new Excel.ChartGridlinesFormat(this.context, _createPropertyObjectPath(this.context, this, "Format", false, false));
				}
				return this._F;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ChartGridlines.prototype, "visible", {
			get: function () {
				_throwIfNotLoaded("visible", this._V, _typeChartGridlines, this._isNull);
				return this._V;
			},
			set: function (value) {
				this._V=value;
				_createSetPropertyAction(this.context, this, "Visible", value);
			},
			enumerable: true,
			configurable: true
		});
		ChartGridlines.prototype.set=function (properties, options) {
			this._recursivelySet(properties, options, ["visible"], ["format"], []);
		};
		ChartGridlines.prototype._handleResult=function (value) {
			_super.prototype._handleResult.call(this, value);
			if (_isNullOrUndefined(value))
				return;
			var obj=value;
			_fixObjectPathIfNecessary(this, obj);
			if (!_isUndefined(obj["Visible"])) {
				this._V=obj["Visible"];
			}
			_handleNavigationPropertyResults(this, obj, ["format", "Format"]);
		};
		ChartGridlines.prototype.load=function (option) {
			_load(this, option);
			return this;
		};
		ChartGridlines.prototype.toJSON=function () {
			return {
				"format": this._F,
				"visible": this._V
			};
		};
		return ChartGridlines;
	}(OfficeExtension.ClientObject));
	Excel.ChartGridlines=ChartGridlines;
	var _typeChartGridlinesFormat="ChartGridlinesFormat";
	var ChartGridlinesFormat=(function (_super) {
		__extends(ChartGridlinesFormat, _super);
		function ChartGridlinesFormat() {
			return _super !==null && _super.apply(this, arguments) || this;
		}
		Object.defineProperty(ChartGridlinesFormat.prototype, "_className", {
			get: function () {
				return "ChartGridlinesFormat";
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ChartGridlinesFormat.prototype, "line", {
			get: function () {
				if (!this._L) {
					this._L=new Excel.ChartLineFormat(this.context, _createPropertyObjectPath(this.context, this, "Line", false, false));
				}
				return this._L;
			},
			enumerable: true,
			configurable: true
		});
		ChartGridlinesFormat.prototype.set=function (properties, options) {
			this._recursivelySet(properties, options, [], ["line"], []);
		};
		ChartGridlinesFormat.prototype._handleResult=function (value) {
			_super.prototype._handleResult.call(this, value);
			if (_isNullOrUndefined(value))
				return;
			var obj=value;
			_fixObjectPathIfNecessary(this, obj);
			_handleNavigationPropertyResults(this, obj, ["line", "Line"]);
		};
		ChartGridlinesFormat.prototype.load=function (option) {
			_load(this, option);
			return this;
		};
		ChartGridlinesFormat.prototype.toJSON=function () {
			return {
				"line": this._L
			};
		};
		return ChartGridlinesFormat;
	}(OfficeExtension.ClientObject));
	Excel.ChartGridlinesFormat=ChartGridlinesFormat;
	var _typeChartLegend="ChartLegend";
	var ChartLegend=(function (_super) {
		__extends(ChartLegend, _super);
		function ChartLegend() {
			return _super !==null && _super.apply(this, arguments) || this;
		}
		Object.defineProperty(ChartLegend.prototype, "_className", {
			get: function () {
				return "ChartLegend";
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ChartLegend.prototype, "format", {
			get: function () {
				if (!this._F) {
					this._F=new Excel.ChartLegendFormat(this.context, _createPropertyObjectPath(this.context, this, "Format", false, false));
				}
				return this._F;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ChartLegend.prototype, "overlay", {
			get: function () {
				_throwIfNotLoaded("overlay", this._O, _typeChartLegend, this._isNull);
				return this._O;
			},
			set: function (value) {
				this._O=value;
				_createSetPropertyAction(this.context, this, "Overlay", value);
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ChartLegend.prototype, "position", {
			get: function () {
				_throwIfNotLoaded("position", this._P, _typeChartLegend, this._isNull);
				return this._P;
			},
			set: function (value) {
				this._P=value;
				_createSetPropertyAction(this.context, this, "Position", value);
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ChartLegend.prototype, "visible", {
			get: function () {
				_throwIfNotLoaded("visible", this._V, _typeChartLegend, this._isNull);
				return this._V;
			},
			set: function (value) {
				this._V=value;
				_createSetPropertyAction(this.context, this, "Visible", value);
			},
			enumerable: true,
			configurable: true
		});
		ChartLegend.prototype.set=function (properties, options) {
			this._recursivelySet(properties, options, ["visible", "position", "overlay"], ["format"], []);
		};
		ChartLegend.prototype._handleResult=function (value) {
			_super.prototype._handleResult.call(this, value);
			if (_isNullOrUndefined(value))
				return;
			var obj=value;
			_fixObjectPathIfNecessary(this, obj);
			if (!_isUndefined(obj["Overlay"])) {
				this._O=obj["Overlay"];
			}
			if (!_isUndefined(obj["Position"])) {
				this._P=obj["Position"];
			}
			if (!_isUndefined(obj["Visible"])) {
				this._V=obj["Visible"];
			}
			_handleNavigationPropertyResults(this, obj, ["format", "Format"]);
		};
		ChartLegend.prototype.load=function (option) {
			_load(this, option);
			return this;
		};
		ChartLegend.prototype.toJSON=function () {
			return {
				"format": this._F,
				"overlay": this._O,
				"position": this._P,
				"visible": this._V
			};
		};
		return ChartLegend;
	}(OfficeExtension.ClientObject));
	Excel.ChartLegend=ChartLegend;
	var _typeChartLegendFormat="ChartLegendFormat";
	var ChartLegendFormat=(function (_super) {
		__extends(ChartLegendFormat, _super);
		function ChartLegendFormat() {
			return _super !==null && _super.apply(this, arguments) || this;
		}
		Object.defineProperty(ChartLegendFormat.prototype, "_className", {
			get: function () {
				return "ChartLegendFormat";
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ChartLegendFormat.prototype, "fill", {
			get: function () {
				if (!this._F) {
					this._F=new Excel.ChartFill(this.context, _createPropertyObjectPath(this.context, this, "Fill", false, false));
				}
				return this._F;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ChartLegendFormat.prototype, "font", {
			get: function () {
				if (!this._Fo) {
					this._Fo=new Excel.ChartFont(this.context, _createPropertyObjectPath(this.context, this, "Font", false, false));
				}
				return this._Fo;
			},
			enumerable: true,
			configurable: true
		});
		ChartLegendFormat.prototype.set=function (properties, options) {
			this._recursivelySet(properties, options, [], ["font"], [
				"fill"
			]);
		};
		ChartLegendFormat.prototype._handleResult=function (value) {
			_super.prototype._handleResult.call(this, value);
			if (_isNullOrUndefined(value))
				return;
			var obj=value;
			_fixObjectPathIfNecessary(this, obj);
			_handleNavigationPropertyResults(this, obj, ["fill", "Fill", "font", "Font"]);
		};
		ChartLegendFormat.prototype.load=function (option) {
			_load(this, option);
			return this;
		};
		ChartLegendFormat.prototype.toJSON=function () {
			return {
				"fill": this._F,
				"font": this._Fo
			};
		};
		return ChartLegendFormat;
	}(OfficeExtension.ClientObject));
	Excel.ChartLegendFormat=ChartLegendFormat;
	var _typeChartTitle="ChartTitle";
	var ChartTitle=(function (_super) {
		__extends(ChartTitle, _super);
		function ChartTitle() {
			return _super !==null && _super.apply(this, arguments) || this;
		}
		Object.defineProperty(ChartTitle.prototype, "_className", {
			get: function () {
				return "ChartTitle";
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ChartTitle.prototype, "format", {
			get: function () {
				if (!this._F) {
					this._F=new Excel.ChartTitleFormat(this.context, _createPropertyObjectPath(this.context, this, "Format", false, false));
				}
				return this._F;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ChartTitle.prototype, "overlay", {
			get: function () {
				_throwIfNotLoaded("overlay", this._O, _typeChartTitle, this._isNull);
				return this._O;
			},
			set: function (value) {
				this._O=value;
				_createSetPropertyAction(this.context, this, "Overlay", value);
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ChartTitle.prototype, "text", {
			get: function () {
				_throwIfNotLoaded("text", this._T, _typeChartTitle, this._isNull);
				return this._T;
			},
			set: function (value) {
				this._T=value;
				_createSetPropertyAction(this.context, this, "Text", value);
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ChartTitle.prototype, "visible", {
			get: function () {
				_throwIfNotLoaded("visible", this._V, _typeChartTitle, this._isNull);
				return this._V;
			},
			set: function (value) {
				this._V=value;
				_createSetPropertyAction(this.context, this, "Visible", value);
			},
			enumerable: true,
			configurable: true
		});
		ChartTitle.prototype.set=function (properties, options) {
			this._recursivelySet(properties, options, ["visible", "text", "overlay"], ["format"], []);
		};
		ChartTitle.prototype._handleResult=function (value) {
			_super.prototype._handleResult.call(this, value);
			if (_isNullOrUndefined(value))
				return;
			var obj=value;
			_fixObjectPathIfNecessary(this, obj);
			if (!_isUndefined(obj["Overlay"])) {
				this._O=obj["Overlay"];
			}
			if (!_isUndefined(obj["Text"])) {
				this._T=obj["Text"];
			}
			if (!_isUndefined(obj["Visible"])) {
				this._V=obj["Visible"];
			}
			_handleNavigationPropertyResults(this, obj, ["format", "Format"]);
		};
		ChartTitle.prototype.load=function (option) {
			_load(this, option);
			return this;
		};
		ChartTitle.prototype.toJSON=function () {
			return {
				"format": this._F,
				"overlay": this._O,
				"text": this._T,
				"visible": this._V
			};
		};
		return ChartTitle;
	}(OfficeExtension.ClientObject));
	Excel.ChartTitle=ChartTitle;
	var _typeChartTitleFormat="ChartTitleFormat";
	var ChartTitleFormat=(function (_super) {
		__extends(ChartTitleFormat, _super);
		function ChartTitleFormat() {
			return _super !==null && _super.apply(this, arguments) || this;
		}
		Object.defineProperty(ChartTitleFormat.prototype, "_className", {
			get: function () {
				return "ChartTitleFormat";
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ChartTitleFormat.prototype, "fill", {
			get: function () {
				if (!this._F) {
					this._F=new Excel.ChartFill(this.context, _createPropertyObjectPath(this.context, this, "Fill", false, false));
				}
				return this._F;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ChartTitleFormat.prototype, "font", {
			get: function () {
				if (!this._Fo) {
					this._Fo=new Excel.ChartFont(this.context, _createPropertyObjectPath(this.context, this, "Font", false, false));
				}
				return this._Fo;
			},
			enumerable: true,
			configurable: true
		});
		ChartTitleFormat.prototype.set=function (properties, options) {
			this._recursivelySet(properties, options, [], ["font"], [
				"fill"
			]);
		};
		ChartTitleFormat.prototype._handleResult=function (value) {
			_super.prototype._handleResult.call(this, value);
			if (_isNullOrUndefined(value))
				return;
			var obj=value;
			_fixObjectPathIfNecessary(this, obj);
			_handleNavigationPropertyResults(this, obj, ["fill", "Fill", "font", "Font"]);
		};
		ChartTitleFormat.prototype.load=function (option) {
			_load(this, option);
			return this;
		};
		ChartTitleFormat.prototype.toJSON=function () {
			return {
				"fill": this._F,
				"font": this._Fo
			};
		};
		return ChartTitleFormat;
	}(OfficeExtension.ClientObject));
	Excel.ChartTitleFormat=ChartTitleFormat;
	var _typeChartFill="ChartFill";
	var ChartFill=(function (_super) {
		__extends(ChartFill, _super);
		function ChartFill() {
			return _super !==null && _super.apply(this, arguments) || this;
		}
		Object.defineProperty(ChartFill.prototype, "_className", {
			get: function () {
				return "ChartFill";
			},
			enumerable: true,
			configurable: true
		});
		ChartFill.prototype.load=function (option) {
			_load(this, option);
			return this;
		};
		ChartFill.prototype.clear=function () {
			_createMethodAction(this.context, this, "Clear", 0, []);
		};
		ChartFill.prototype.setSolidColor=function (color) {
			_createMethodAction(this.context, this, "SetSolidColor", 0, [color]);
		};
		ChartFill.prototype._handleResult=function (value) {
			_super.prototype._handleResult.call(this, value);
			if (_isNullOrUndefined(value))
				return;
			var obj=value;
			_fixObjectPathIfNecessary(this, obj);
		};
		ChartFill.prototype.toJSON=function () {
			return {};
		};
		return ChartFill;
	}(OfficeExtension.ClientObject));
	Excel.ChartFill=ChartFill;
	var _typeChartLineFormat="ChartLineFormat";
	var ChartLineFormat=(function (_super) {
		__extends(ChartLineFormat, _super);
		function ChartLineFormat() {
			return _super !==null && _super.apply(this, arguments) || this;
		}
		Object.defineProperty(ChartLineFormat.prototype, "_className", {
			get: function () {
				return "ChartLineFormat";
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ChartLineFormat.prototype, "color", {
			get: function () {
				_throwIfNotLoaded("color", this._C, _typeChartLineFormat, this._isNull);
				return this._C;
			},
			set: function (value) {
				this._C=value;
				_createSetPropertyAction(this.context, this, "Color", value);
			},
			enumerable: true,
			configurable: true
		});
		ChartLineFormat.prototype.set=function (properties, options) {
			this._recursivelySet(properties, options, ["color"], [], []);
		};
		ChartLineFormat.prototype.clear=function () {
			_createMethodAction(this.context, this, "Clear", 0, []);
		};
		ChartLineFormat.prototype._handleResult=function (value) {
			_super.prototype._handleResult.call(this, value);
			if (_isNullOrUndefined(value))
				return;
			var obj=value;
			_fixObjectPathIfNecessary(this, obj);
			if (!_isUndefined(obj["Color"])) {
				this._C=obj["Color"];
			}
		};
		ChartLineFormat.prototype.load=function (option) {
			_load(this, option);
			return this;
		};
		ChartLineFormat.prototype.toJSON=function () {
			return {
				"color": this._C
			};
		};
		return ChartLineFormat;
	}(OfficeExtension.ClientObject));
	Excel.ChartLineFormat=ChartLineFormat;
	var _typeChartFont="ChartFont";
	var ChartFont=(function (_super) {
		__extends(ChartFont, _super);
		function ChartFont() {
			return _super !==null && _super.apply(this, arguments) || this;
		}
		Object.defineProperty(ChartFont.prototype, "_className", {
			get: function () {
				return "ChartFont";
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ChartFont.prototype, "bold", {
			get: function () {
				_throwIfNotLoaded("bold", this._B, _typeChartFont, this._isNull);
				return this._B;
			},
			set: function (value) {
				this._B=value;
				_createSetPropertyAction(this.context, this, "Bold", value);
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ChartFont.prototype, "color", {
			get: function () {
				_throwIfNotLoaded("color", this._C, _typeChartFont, this._isNull);
				return this._C;
			},
			set: function (value) {
				this._C=value;
				_createSetPropertyAction(this.context, this, "Color", value);
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ChartFont.prototype, "italic", {
			get: function () {
				_throwIfNotLoaded("italic", this._I, _typeChartFont, this._isNull);
				return this._I;
			},
			set: function (value) {
				this._I=value;
				_createSetPropertyAction(this.context, this, "Italic", value);
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ChartFont.prototype, "name", {
			get: function () {
				_throwIfNotLoaded("name", this._N, _typeChartFont, this._isNull);
				return this._N;
			},
			set: function (value) {
				this._N=value;
				_createSetPropertyAction(this.context, this, "Name", value);
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ChartFont.prototype, "size", {
			get: function () {
				_throwIfNotLoaded("size", this._S, _typeChartFont, this._isNull);
				return this._S;
			},
			set: function (value) {
				this._S=value;
				_createSetPropertyAction(this.context, this, "Size", value);
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ChartFont.prototype, "underline", {
			get: function () {
				_throwIfNotLoaded("underline", this._U, _typeChartFont, this._isNull);
				return this._U;
			},
			set: function (value) {
				this._U=value;
				_createSetPropertyAction(this.context, this, "Underline", value);
			},
			enumerable: true,
			configurable: true
		});
		ChartFont.prototype.set=function (properties, options) {
			this._recursivelySet(properties, options, ["bold", "color", "italic", "name", "size", "underline"], [], []);
		};
		ChartFont.prototype._handleResult=function (value) {
			_super.prototype._handleResult.call(this, value);
			if (_isNullOrUndefined(value))
				return;
			var obj=value;
			_fixObjectPathIfNecessary(this, obj);
			if (!_isUndefined(obj["Bold"])) {
				this._B=obj["Bold"];
			}
			if (!_isUndefined(obj["Color"])) {
				this._C=obj["Color"];
			}
			if (!_isUndefined(obj["Italic"])) {
				this._I=obj["Italic"];
			}
			if (!_isUndefined(obj["Name"])) {
				this._N=obj["Name"];
			}
			if (!_isUndefined(obj["Size"])) {
				this._S=obj["Size"];
			}
			if (!_isUndefined(obj["Underline"])) {
				this._U=obj["Underline"];
			}
		};
		ChartFont.prototype.load=function (option) {
			_load(this, option);
			return this;
		};
		ChartFont.prototype.toJSON=function () {
			return {
				"bold": this._B,
				"color": this._C,
				"italic": this._I,
				"name": this._N,
				"size": this._S,
				"underline": this._U
			};
		};
		return ChartFont;
	}(OfficeExtension.ClientObject));
	Excel.ChartFont=ChartFont;
	var _typeRangeSort="RangeSort";
	var RangeSort=(function (_super) {
		__extends(RangeSort, _super);
		function RangeSort() {
			return _super !==null && _super.apply(this, arguments) || this;
		}
		Object.defineProperty(RangeSort.prototype, "_className", {
			get: function () {
				return "RangeSort";
			},
			enumerable: true,
			configurable: true
		});
		RangeSort.prototype.apply=function (fields, matchCase, hasHeaders, orientation, method) {
			_createMethodAction(this.context, this, "Apply", 0, [fields, matchCase, hasHeaders, orientation, method]);
		};
		RangeSort.prototype._handleResult=function (value) {
			_super.prototype._handleResult.call(this, value);
			if (_isNullOrUndefined(value))
				return;
			var obj=value;
			_fixObjectPathIfNecessary(this, obj);
		};
		RangeSort.prototype.toJSON=function () {
			return {};
		};
		return RangeSort;
	}(OfficeExtension.ClientObject));
	Excel.RangeSort=RangeSort;
	var _typeTableSort="TableSort";
	var TableSort=(function (_super) {
		__extends(TableSort, _super);
		function TableSort() {
			return _super !==null && _super.apply(this, arguments) || this;
		}
		Object.defineProperty(TableSort.prototype, "_className", {
			get: function () {
				return "TableSort";
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(TableSort.prototype, "fields", {
			get: function () {
				_throwIfNotLoaded("fields", this._F, _typeTableSort, this._isNull);
				return this._F;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(TableSort.prototype, "matchCase", {
			get: function () {
				_throwIfNotLoaded("matchCase", this._M, _typeTableSort, this._isNull);
				return this._M;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(TableSort.prototype, "method", {
			get: function () {
				_throwIfNotLoaded("method", this._Me, _typeTableSort, this._isNull);
				return this._Me;
			},
			enumerable: true,
			configurable: true
		});
		TableSort.prototype.apply=function (fields, matchCase, method) {
			_createMethodAction(this.context, this, "Apply", 0, [fields, matchCase, method]);
		};
		TableSort.prototype.clear=function () {
			_createMethodAction(this.context, this, "Clear", 0, []);
		};
		TableSort.prototype.reapply=function () {
			_createMethodAction(this.context, this, "Reapply", 0, []);
		};
		TableSort.prototype._handleResult=function (value) {
			_super.prototype._handleResult.call(this, value);
			if (_isNullOrUndefined(value))
				return;
			var obj=value;
			_fixObjectPathIfNecessary(this, obj);
			if (!_isUndefined(obj["Fields"])) {
				this._F=obj["Fields"];
			}
			if (!_isUndefined(obj["MatchCase"])) {
				this._M=obj["MatchCase"];
			}
			if (!_isUndefined(obj["Method"])) {
				this._Me=obj["Method"];
			}
		};
		TableSort.prototype.load=function (option) {
			_load(this, option);
			return this;
		};
		TableSort.prototype.toJSON=function () {
			return {
				"fields": this._F,
				"matchCase": this._M,
				"method": this._Me
			};
		};
		return TableSort;
	}(OfficeExtension.ClientObject));
	Excel.TableSort=TableSort;
	var _typeFilter="Filter";
	var Filter=(function (_super) {
		__extends(Filter, _super);
		function Filter() {
			return _super !==null && _super.apply(this, arguments) || this;
		}
		Object.defineProperty(Filter.prototype, "_className", {
			get: function () {
				return "Filter";
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Filter.prototype, "criteria", {
			get: function () {
				_throwIfNotLoaded("criteria", this._C, _typeFilter, this._isNull);
				return this._C;
			},
			enumerable: true,
			configurable: true
		});
		Filter.prototype.apply=function (criteria) {
			_createMethodAction(this.context, this, "Apply", 0, [criteria]);
		};
		Filter.prototype.applyBottomItemsFilter=function (count) {
			_createMethodAction(this.context, this, "ApplyBottomItemsFilter", 0, [count]);
		};
		Filter.prototype.applyBottomPercentFilter=function (percent) {
			_createMethodAction(this.context, this, "ApplyBottomPercentFilter", 0, [percent]);
		};
		Filter.prototype.applyCellColorFilter=function (color) {
			_createMethodAction(this.context, this, "ApplyCellColorFilter", 0, [color]);
		};
		Filter.prototype.applyCustomFilter=function (criteria1, criteria2, oper) {
			_createMethodAction(this.context, this, "ApplyCustomFilter", 0, [criteria1, criteria2, oper]);
		};
		Filter.prototype.applyDynamicFilter=function (criteria) {
			_createMethodAction(this.context, this, "ApplyDynamicFilter", 0, [criteria]);
		};
		Filter.prototype.applyFontColorFilter=function (color) {
			_createMethodAction(this.context, this, "ApplyFontColorFilter", 0, [color]);
		};
		Filter.prototype.applyIconFilter=function (icon) {
			_createMethodAction(this.context, this, "ApplyIconFilter", 0, [icon]);
		};
		Filter.prototype.applyTopItemsFilter=function (count) {
			_createMethodAction(this.context, this, "ApplyTopItemsFilter", 0, [count]);
		};
		Filter.prototype.applyTopPercentFilter=function (percent) {
			_createMethodAction(this.context, this, "ApplyTopPercentFilter", 0, [percent]);
		};
		Filter.prototype.applyValuesFilter=function (values) {
			_createMethodAction(this.context, this, "ApplyValuesFilter", 0, [values]);
		};
		Filter.prototype.clear=function () {
			_createMethodAction(this.context, this, "Clear", 0, []);
		};
		Filter.prototype._handleResult=function (value) {
			_super.prototype._handleResult.call(this, value);
			if (_isNullOrUndefined(value))
				return;
			var obj=value;
			_fixObjectPathIfNecessary(this, obj);
			if (!_isUndefined(obj["Criteria"])) {
				this._C=obj["Criteria"];
			}
		};
		Filter.prototype.load=function (option) {
			_load(this, option);
			return this;
		};
		Filter.prototype.toJSON=function () {
			return {
				"criteria": this._C
			};
		};
		return Filter;
	}(OfficeExtension.ClientObject));
	Excel.Filter=Filter;
	var _typeCustomXmlPartScopedCollection="CustomXmlPartScopedCollection";
	var CustomXmlPartScopedCollection=(function (_super) {
		__extends(CustomXmlPartScopedCollection, _super);
		function CustomXmlPartScopedCollection() {
			return _super !==null && _super.apply(this, arguments) || this;
		}
		Object.defineProperty(CustomXmlPartScopedCollection.prototype, "_className", {
			get: function () {
				return "CustomXmlPartScopedCollection";
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(CustomXmlPartScopedCollection.prototype, "items", {
			get: function () {
				_throwIfNotLoaded("items", this.m__items, _typeCustomXmlPartScopedCollection, this._isNull);
				return this.m__items;
			},
			enumerable: true,
			configurable: true
		});
		CustomXmlPartScopedCollection.prototype.getCount=function () {
			var action=_createMethodAction(this.context, this, "GetCount", 1, []);
			var ret=new OfficeExtension.ClientResult();
			_addActionResultHandler(this, action, ret);
			return ret;
		};
		CustomXmlPartScopedCollection.prototype.getItem=function (id) {
			return new Excel.CustomXmlPart(this.context, _createIndexerObjectPath(this.context, this, [id]));
		};
		CustomXmlPartScopedCollection.prototype.getItemOrNullObject=function (id) {
			return new Excel.CustomXmlPart(this.context, _createMethodObjectPath(this.context, this, "GetItemOrNullObject", 1, [id], false, false, null));
		};
		CustomXmlPartScopedCollection.prototype.getOnlyItem=function () {
			return new Excel.CustomXmlPart(this.context, _createMethodObjectPath(this.context, this, "GetOnlyItem", 1, [], false, false, null));
		};
		CustomXmlPartScopedCollection.prototype.getOnlyItemOrNullObject=function () {
			return new Excel.CustomXmlPart(this.context, _createMethodObjectPath(this.context, this, "GetOnlyItemOrNullObject", 1, [], false, false, null));
		};
		CustomXmlPartScopedCollection.prototype._handleResult=function (value) {
			_super.prototype._handleResult.call(this, value);
			if (_isNullOrUndefined(value))
				return;
			var obj=value;
			_fixObjectPathIfNecessary(this, obj);
			if (!_isNullOrUndefined(obj[OfficeExtension.Constants.items])) {
				this.m__items=[];
				var _data=obj[OfficeExtension.Constants.items];
				for (var i=0; i < _data.length; i++) {
					var _item=new Excel.CustomXmlPart(this.context, _createChildItemObjectPathUsingIndexerOrGetItemAt(true, this.context, this, _data[i], i));
					_item._handleResult(_data[i]);
					this.m__items.push(_item);
				}
			}
		};
		CustomXmlPartScopedCollection.prototype.load=function (option) {
			_load(this, option);
			return this;
		};
		CustomXmlPartScopedCollection.prototype.toJSON=function () {
			return {};
		};
		return CustomXmlPartScopedCollection;
	}(OfficeExtension.ClientObject));
	Excel.CustomXmlPartScopedCollection=CustomXmlPartScopedCollection;
	var _typeCustomXmlPartCollection="CustomXmlPartCollection";
	var CustomXmlPartCollection=(function (_super) {
		__extends(CustomXmlPartCollection, _super);
		function CustomXmlPartCollection() {
			return _super !==null && _super.apply(this, arguments) || this;
		}
		Object.defineProperty(CustomXmlPartCollection.prototype, "_className", {
			get: function () {
				return "CustomXmlPartCollection";
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(CustomXmlPartCollection.prototype, "items", {
			get: function () {
				_throwIfNotLoaded("items", this.m__items, _typeCustomXmlPartCollection, this._isNull);
				return this.m__items;
			},
			enumerable: true,
			configurable: true
		});
		CustomXmlPartCollection.prototype.add=function (xml) {
			return new Excel.CustomXmlPart(this.context, _createMethodObjectPath(this.context, this, "Add", 0, [xml], false, true, null));
		};
		CustomXmlPartCollection.prototype.getByNamespace=function (namespaceUri) {
			return new Excel.CustomXmlPartScopedCollection(this.context, _createMethodObjectPath(this.context, this, "GetByNamespace", 1, [namespaceUri], true, false, null));
		};
		CustomXmlPartCollection.prototype.getCount=function () {
			var action=_createMethodAction(this.context, this, "GetCount", 1, []);
			var ret=new OfficeExtension.ClientResult();
			_addActionResultHandler(this, action, ret);
			return ret;
		};
		CustomXmlPartCollection.prototype.getItem=function (id) {
			return new Excel.CustomXmlPart(this.context, _createIndexerObjectPath(this.context, this, [id]));
		};
		CustomXmlPartCollection.prototype.getItemOrNullObject=function (id) {
			return new Excel.CustomXmlPart(this.context, _createMethodObjectPath(this.context, this, "GetItemOrNullObject", 1, [id], false, false, null));
		};
		CustomXmlPartCollection.prototype._handleResult=function (value) {
			_super.prototype._handleResult.call(this, value);
			if (_isNullOrUndefined(value))
				return;
			var obj=value;
			_fixObjectPathIfNecessary(this, obj);
			if (!_isNullOrUndefined(obj[OfficeExtension.Constants.items])) {
				this.m__items=[];
				var _data=obj[OfficeExtension.Constants.items];
				for (var i=0; i < _data.length; i++) {
					var _item=new Excel.CustomXmlPart(this.context, _createChildItemObjectPathUsingIndexerOrGetItemAt(true, this.context, this, _data[i], i));
					_item._handleResult(_data[i]);
					this.m__items.push(_item);
				}
			}
		};
		CustomXmlPartCollection.prototype.load=function (option) {
			_load(this, option);
			return this;
		};
		CustomXmlPartCollection.prototype.toJSON=function () {
			return {};
		};
		return CustomXmlPartCollection;
	}(OfficeExtension.ClientObject));
	Excel.CustomXmlPartCollection=CustomXmlPartCollection;
	var _typeCustomXmlPart="CustomXmlPart";
	var CustomXmlPart=(function (_super) {
		__extends(CustomXmlPart, _super);
		function CustomXmlPart() {
			return _super !==null && _super.apply(this, arguments) || this;
		}
		Object.defineProperty(CustomXmlPart.prototype, "_className", {
			get: function () {
				return "CustomXmlPart";
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(CustomXmlPart.prototype, "id", {
			get: function () {
				_throwIfNotLoaded("id", this._I, _typeCustomXmlPart, this._isNull);
				return this._I;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(CustomXmlPart.prototype, "namespaceUri", {
			get: function () {
				_throwIfNotLoaded("namespaceUri", this._N, _typeCustomXmlPart, this._isNull);
				return this._N;
			},
			enumerable: true,
			configurable: true
		});
		CustomXmlPart.prototype.delete=function () {
			_createMethodAction(this.context, this, "Delete", 0, []);
		};
		CustomXmlPart.prototype.getXml=function () {
			var action=_createMethodAction(this.context, this, "GetXml", 1, []);
			var ret=new OfficeExtension.ClientResult();
			_addActionResultHandler(this, action, ret);
			return ret;
		};
		CustomXmlPart.prototype.setXml=function (xml) {
			_createMethodAction(this.context, this, "SetXml", 0, [xml]);
		};
		CustomXmlPart.prototype._handleResult=function (value) {
			_super.prototype._handleResult.call(this, value);
			if (_isNullOrUndefined(value))
				return;
			var obj=value;
			_fixObjectPathIfNecessary(this, obj);
			if (!_isUndefined(obj["Id"])) {
				this._I=obj["Id"];
			}
			if (!_isUndefined(obj["NamespaceUri"])) {
				this._N=obj["NamespaceUri"];
			}
		};
		CustomXmlPart.prototype.load=function (option) {
			_load(this, option);
			return this;
		};
		CustomXmlPart.prototype._handleIdResult=function (value) {
			_super.prototype._handleIdResult.call(this, value);
			if (_isNullOrUndefined(value)) {
				return;
			}
			if (!_isUndefined(value["Id"])) {
				this._I=value["Id"];
			}
		};
		CustomXmlPart.prototype.toJSON=function () {
			return {
				"id": this._I,
				"namespaceUri": this._N
			};
		};
		return CustomXmlPart;
	}(OfficeExtension.ClientObject));
	Excel.CustomXmlPart=CustomXmlPart;
	var _type_V1Api="_V1Api";
	var _V1Api=(function (_super) {
		__extends(_V1Api, _super);
		function _V1Api() {
			return _super !==null && _super.apply(this, arguments) || this;
		}
		Object.defineProperty(_V1Api.prototype, "_className", {
			get: function () {
				return "_V1Api";
			},
			enumerable: true,
			configurable: true
		});
		_V1Api.prototype.bindingAddColumns=function (input) {
			var action=_createMethodAction(this.context, this, "BindingAddColumns", 0, [input]);
			var ret=new OfficeExtension.ClientResult();
			_addActionResultHandler(this, action, ret);
			return ret;
		};
		_V1Api.prototype.bindingAddFromNamedItem=function (input) {
			var action=_createMethodAction(this.context, this, "BindingAddFromNamedItem", 1, [input]);
			var ret=new OfficeExtension.ClientResult();
			_addActionResultHandler(this, action, ret);
			return ret;
		};
		_V1Api.prototype.bindingAddFromPrompt=function (input) {
			var action=_createMethodAction(this.context, this, "BindingAddFromPrompt", 1, [input]);
			var ret=new OfficeExtension.ClientResult();
			_addActionResultHandler(this, action, ret);
			return ret;
		};
		_V1Api.prototype.bindingAddFromSelection=function (input) {
			var action=_createMethodAction(this.context, this, "BindingAddFromSelection", 1, [input]);
			var ret=new OfficeExtension.ClientResult();
			_addActionResultHandler(this, action, ret);
			return ret;
		};
		_V1Api.prototype.bindingAddRows=function (input) {
			var action=_createMethodAction(this.context, this, "BindingAddRows", 0, [input]);
			var ret=new OfficeExtension.ClientResult();
			_addActionResultHandler(this, action, ret);
			return ret;
		};
		_V1Api.prototype.bindingClearFormats=function (input) {
			var action=_createMethodAction(this.context, this, "BindingClearFormats", 0, [input]);
			var ret=new OfficeExtension.ClientResult();
			_addActionResultHandler(this, action, ret);
			return ret;
		};
		_V1Api.prototype.bindingDeleteAllDataValues=function (input) {
			var action=_createMethodAction(this.context, this, "BindingDeleteAllDataValues", 0, [input]);
			var ret=new OfficeExtension.ClientResult();
			_addActionResultHandler(this, action, ret);
			return ret;
		};
		_V1Api.prototype.bindingGetAll=function () {
			var action=_createMethodAction(this.context, this, "BindingGetAll", 1, []);
			var ret=new OfficeExtension.ClientResult();
			_addActionResultHandler(this, action, ret);
			return ret;
		};
		_V1Api.prototype.bindingGetById=function (input) {
			var action=_createMethodAction(this.context, this, "BindingGetById", 1, [input]);
			var ret=new OfficeExtension.ClientResult();
			_addActionResultHandler(this, action, ret);
			return ret;
		};
		_V1Api.prototype.bindingGetData=function (input) {
			var action=_createMethodAction(this.context, this, "BindingGetData", 1, [input]);
			var ret=new OfficeExtension.ClientResult();
			_addActionResultHandler(this, action, ret);
			return ret;
		};
		_V1Api.prototype.bindingReleaseById=function (input) {
			var action=_createMethodAction(this.context, this, "BindingReleaseById", 1, [input]);
			var ret=new OfficeExtension.ClientResult();
			_addActionResultHandler(this, action, ret);
			return ret;
		};
		_V1Api.prototype.bindingSetData=function (input) {
			var action=_createMethodAction(this.context, this, "BindingSetData", 0, [input]);
			var ret=new OfficeExtension.ClientResult();
			_addActionResultHandler(this, action, ret);
			return ret;
		};
		_V1Api.prototype.bindingSetFormats=function (input) {
			var action=_createMethodAction(this.context, this, "BindingSetFormats", 0, [input]);
			var ret=new OfficeExtension.ClientResult();
			_addActionResultHandler(this, action, ret);
			return ret;
		};
		_V1Api.prototype.bindingSetTableOptions=function (input) {
			var action=_createMethodAction(this.context, this, "BindingSetTableOptions", 0, [input]);
			var ret=new OfficeExtension.ClientResult();
			_addActionResultHandler(this, action, ret);
			return ret;
		};
		_V1Api.prototype.getFilePropertiesAsync=function () {
			_throwIfApiNotSupported("_V1Api.getFilePropertiesAsync", _defaultApiSetName, "1.6", _hostName);
			var action=_createMethodAction(this.context, this, "GetFilePropertiesAsync", 1, []);
			var ret=new OfficeExtension.ClientResult();
			_addActionResultHandler(this, action, ret);
			return ret;
		};
		_V1Api.prototype.getSelectedData=function (input) {
			var action=_createMethodAction(this.context, this, "GetSelectedData", 1, [input]);
			var ret=new OfficeExtension.ClientResult();
			_addActionResultHandler(this, action, ret);
			return ret;
		};
		_V1Api.prototype.gotoById=function (input) {
			var action=_createMethodAction(this.context, this, "GotoById", 1, [input]);
			var ret=new OfficeExtension.ClientResult();
			_addActionResultHandler(this, action, ret);
			return ret;
		};
		_V1Api.prototype.setSelectedData=function (input) {
			var action=_createMethodAction(this.context, this, "SetSelectedData", 0, [input]);
			var ret=new OfficeExtension.ClientResult();
			_addActionResultHandler(this, action, ret);
			return ret;
		};
		_V1Api.prototype._handleResult=function (value) {
			_super.prototype._handleResult.call(this, value);
			if (_isNullOrUndefined(value))
				return;
			var obj=value;
			_fixObjectPathIfNecessary(this, obj);
		};
		_V1Api.prototype.toJSON=function () {
			return {};
		};
		return _V1Api;
	}(OfficeExtension.ClientObject));
	Excel._V1Api=_V1Api;
	var _typePivotTableCollection="PivotTableCollection";
	var PivotTableCollection=(function (_super) {
		__extends(PivotTableCollection, _super);
		function PivotTableCollection() {
			return _super !==null && _super.apply(this, arguments) || this;
		}
		Object.defineProperty(PivotTableCollection.prototype, "_className", {
			get: function () {
				return "PivotTableCollection";
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(PivotTableCollection.prototype, "items", {
			get: function () {
				_throwIfNotLoaded("items", this.m__items, _typePivotTableCollection, this._isNull);
				return this.m__items;
			},
			enumerable: true,
			configurable: true
		});
		PivotTableCollection.prototype.getCount=function () {
			_throwIfApiNotSupported("PivotTableCollection.getCount", _defaultApiSetName, "1.4", _hostName);
			var action=_createMethodAction(this.context, this, "GetCount", 1, []);
			var ret=new OfficeExtension.ClientResult();
			_addActionResultHandler(this, action, ret);
			return ret;
		};
		PivotTableCollection.prototype.getItem=function (name) {
			return new Excel.PivotTable(this.context, _createIndexerObjectPath(this.context, this, [name]));
		};
		PivotTableCollection.prototype.getItemOrNullObject=function (name) {
			_throwIfApiNotSupported("PivotTableCollection.getItemOrNullObject", _defaultApiSetName, "1.4", _hostName);
			return new Excel.PivotTable(this.context, _createMethodObjectPath(this.context, this, "GetItemOrNullObject", 1, [name], false, false, null));
		};
		PivotTableCollection.prototype.refreshAll=function () {
			_createMethodAction(this.context, this, "RefreshAll", 0, []);
		};
		PivotTableCollection.prototype._handleResult=function (value) {
			_super.prototype._handleResult.call(this, value);
			if (_isNullOrUndefined(value))
				return;
			var obj=value;
			_fixObjectPathIfNecessary(this, obj);
			if (!_isNullOrUndefined(obj[OfficeExtension.Constants.items])) {
				this.m__items=[];
				var _data=obj[OfficeExtension.Constants.items];
				for (var i=0; i < _data.length; i++) {
					var _item=new Excel.PivotTable(this.context, _createChildItemObjectPathUsingIndexerOrGetItemAt(true, this.context, this, _data[i], i));
					_item._handleResult(_data[i]);
					this.m__items.push(_item);
				}
			}
		};
		PivotTableCollection.prototype.load=function (option) {
			_load(this, option);
			return this;
		};
		PivotTableCollection.prototype.toJSON=function () {
			return {};
		};
		return PivotTableCollection;
	}(OfficeExtension.ClientObject));
	Excel.PivotTableCollection=PivotTableCollection;
	var _typePivotTable="PivotTable";
	var PivotTable=(function (_super) {
		__extends(PivotTable, _super);
		function PivotTable() {
			return _super !==null && _super.apply(this, arguments) || this;
		}
		Object.defineProperty(PivotTable.prototype, "_className", {
			get: function () {
				return "PivotTable";
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(PivotTable.prototype, "worksheet", {
			get: function () {
				if (!this._W) {
					this._W=new Excel.Worksheet(this.context, _createPropertyObjectPath(this.context, this, "Worksheet", false, false));
				}
				return this._W;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(PivotTable.prototype, "id", {
			get: function () {
				_throwIfNotLoaded("id", this._I, _typePivotTable, this._isNull);
				_throwIfApiNotSupported("PivotTable.id", _defaultApiSetName, "1.5", _hostName);
				return this._I;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(PivotTable.prototype, "name", {
			get: function () {
				_throwIfNotLoaded("name", this._N, _typePivotTable, this._isNull);
				return this._N;
			},
			set: function (value) {
				this._N=value;
				_createSetPropertyAction(this.context, this, "Name", value);
			},
			enumerable: true,
			configurable: true
		});
		PivotTable.prototype.set=function (properties, options) {
			this._recursivelySet(properties, options, ["name"], [], [
				"worksheet",
				"worksheet"
			]);
		};
		PivotTable.prototype.refresh=function () {
			_createMethodAction(this.context, this, "Refresh", 0, []);
		};
		PivotTable.prototype._handleResult=function (value) {
			_super.prototype._handleResult.call(this, value);
			if (_isNullOrUndefined(value))
				return;
			var obj=value;
			_fixObjectPathIfNecessary(this, obj);
			if (!_isUndefined(obj["Id"])) {
				this._I=obj["Id"];
			}
			if (!_isUndefined(obj["Name"])) {
				this._N=obj["Name"];
			}
			_handleNavigationPropertyResults(this, obj, ["worksheet", "Worksheet"]);
		};
		PivotTable.prototype.load=function (option) {
			_load(this, option);
			return this;
		};
		PivotTable.prototype._handleIdResult=function (value) {
			_super.prototype._handleIdResult.call(this, value);
			if (_isNullOrUndefined(value)) {
				return;
			}
			if (!_isUndefined(value["Id"])) {
				this._I=value["Id"];
			}
		};
		PivotTable.prototype.toJSON=function () {
			return {
				"id": this._I,
				"name": this._N
			};
		};
		return PivotTable;
	}(OfficeExtension.ClientObject));
	Excel.PivotTable=PivotTable;
	var _typeConditionalFormatCollection="ConditionalFormatCollection";
	var ConditionalFormatCollection=(function (_super) {
		__extends(ConditionalFormatCollection, _super);
		function ConditionalFormatCollection() {
			return _super !==null && _super.apply(this, arguments) || this;
		}
		Object.defineProperty(ConditionalFormatCollection.prototype, "_className", {
			get: function () {
				return "ConditionalFormatCollection";
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ConditionalFormatCollection.prototype, "items", {
			get: function () {
				_throwIfNotLoaded("items", this.m__items, _typeConditionalFormatCollection, this._isNull);
				return this.m__items;
			},
			enumerable: true,
			configurable: true
		});
		ConditionalFormatCollection.prototype.add=function (type) {
			return new Excel.ConditionalFormat(this.context, _createMethodObjectPath(this.context, this, "Add", 0, [type], false, true, null));
		};
		ConditionalFormatCollection.prototype.clearAll=function () {
			_createMethodAction(this.context, this, "ClearAll", 0, []);
		};
		ConditionalFormatCollection.prototype.getCount=function () {
			var action=_createMethodAction(this.context, this, "GetCount", 1, []);
			var ret=new OfficeExtension.ClientResult();
			_addActionResultHandler(this, action, ret);
			return ret;
		};
		ConditionalFormatCollection.prototype.getItem=function (id) {
			return new Excel.ConditionalFormat(this.context, _createIndexerObjectPath(this.context, this, [id]));
		};
		ConditionalFormatCollection.prototype.getItemAt=function (index) {
			return new Excel.ConditionalFormat(this.context, _createMethodObjectPath(this.context, this, "GetItemAt", 1, [index], false, false, null));
		};
		ConditionalFormatCollection.prototype._handleResult=function (value) {
			_super.prototype._handleResult.call(this, value);
			if (_isNullOrUndefined(value))
				return;
			var obj=value;
			_fixObjectPathIfNecessary(this, obj);
			if (!_isNullOrUndefined(obj[OfficeExtension.Constants.items])) {
				this.m__items=[];
				var _data=obj[OfficeExtension.Constants.items];
				for (var i=0; i < _data.length; i++) {
					var _item=new Excel.ConditionalFormat(this.context, _createChildItemObjectPathUsingIndexerOrGetItemAt(true, this.context, this, _data[i], i));
					_item._handleResult(_data[i]);
					this.m__items.push(_item);
				}
			}
		};
		ConditionalFormatCollection.prototype.load=function (option) {
			_load(this, option);
			return this;
		};
		ConditionalFormatCollection.prototype.toJSON=function () {
			return {};
		};
		return ConditionalFormatCollection;
	}(OfficeExtension.ClientObject));
	Excel.ConditionalFormatCollection=ConditionalFormatCollection;
	var _typeConditionalFormat="ConditionalFormat";
	var ConditionalFormat=(function (_super) {
		__extends(ConditionalFormat, _super);
		function ConditionalFormat() {
			return _super !==null && _super.apply(this, arguments) || this;
		}
		Object.defineProperty(ConditionalFormat.prototype, "_className", {
			get: function () {
				return "ConditionalFormat";
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ConditionalFormat.prototype, "cellValue", {
			get: function () {
				if (!this._C) {
					this._C=new Excel.CellValueConditionalFormat(this.context, _createPropertyObjectPath(this.context, this, "CellValue", false, false));
				}
				return this._C;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ConditionalFormat.prototype, "cellValueOrNullObject", {
			get: function () {
				if (!this._Ce) {
					this._Ce=new Excel.CellValueConditionalFormat(this.context, _createPropertyObjectPath(this.context, this, "CellValueOrNullObject", false, false));
				}
				return this._Ce;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ConditionalFormat.prototype, "colorScale", {
			get: function () {
				if (!this._Co) {
					this._Co=new Excel.ColorScaleConditionalFormat(this.context, _createPropertyObjectPath(this.context, this, "ColorScale", false, false));
				}
				return this._Co;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ConditionalFormat.prototype, "colorScaleOrNullObject", {
			get: function () {
				if (!this._Col) {
					this._Col=new Excel.ColorScaleConditionalFormat(this.context, _createPropertyObjectPath(this.context, this, "ColorScaleOrNullObject", false, false));
				}
				return this._Col;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ConditionalFormat.prototype, "custom", {
			get: function () {
				if (!this._Cu) {
					this._Cu=new Excel.CustomConditionalFormat(this.context, _createPropertyObjectPath(this.context, this, "Custom", false, false));
				}
				return this._Cu;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ConditionalFormat.prototype, "customOrNullObject", {
			get: function () {
				if (!this._Cus) {
					this._Cus=new Excel.CustomConditionalFormat(this.context, _createPropertyObjectPath(this.context, this, "CustomOrNullObject", false, false));
				}
				return this._Cus;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ConditionalFormat.prototype, "dataBar", {
			get: function () {
				if (!this._D) {
					this._D=new Excel.DataBarConditionalFormat(this.context, _createPropertyObjectPath(this.context, this, "DataBar", false, false));
				}
				return this._D;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ConditionalFormat.prototype, "dataBarOrNullObject", {
			get: function () {
				if (!this._Da) {
					this._Da=new Excel.DataBarConditionalFormat(this.context, _createPropertyObjectPath(this.context, this, "DataBarOrNullObject", false, false));
				}
				return this._Da;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ConditionalFormat.prototype, "iconSet", {
			get: function () {
				if (!this._I) {
					this._I=new Excel.IconSetConditionalFormat(this.context, _createPropertyObjectPath(this.context, this, "IconSet", false, false));
				}
				return this._I;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ConditionalFormat.prototype, "iconSetOrNullObject", {
			get: function () {
				if (!this._Ic) {
					this._Ic=new Excel.IconSetConditionalFormat(this.context, _createPropertyObjectPath(this.context, this, "IconSetOrNullObject", false, false));
				}
				return this._Ic;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ConditionalFormat.prototype, "preset", {
			get: function () {
				if (!this._P) {
					this._P=new Excel.PresetCriteriaConditionalFormat(this.context, _createPropertyObjectPath(this.context, this, "Preset", false, false));
				}
				return this._P;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ConditionalFormat.prototype, "presetOrNullObject", {
			get: function () {
				if (!this._Pr) {
					this._Pr=new Excel.PresetCriteriaConditionalFormat(this.context, _createPropertyObjectPath(this.context, this, "PresetOrNullObject", false, false));
				}
				return this._Pr;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ConditionalFormat.prototype, "textComparison", {
			get: function () {
				if (!this._T) {
					this._T=new Excel.TextConditionalFormat(this.context, _createPropertyObjectPath(this.context, this, "TextComparison", false, false));
				}
				return this._T;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ConditionalFormat.prototype, "textComparisonOrNullObject", {
			get: function () {
				if (!this._Te) {
					this._Te=new Excel.TextConditionalFormat(this.context, _createPropertyObjectPath(this.context, this, "TextComparisonOrNullObject", false, false));
				}
				return this._Te;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ConditionalFormat.prototype, "topBottom", {
			get: function () {
				if (!this._To) {
					this._To=new Excel.TopBottomConditionalFormat(this.context, _createPropertyObjectPath(this.context, this, "TopBottom", false, false));
				}
				return this._To;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ConditionalFormat.prototype, "topBottomOrNullObject", {
			get: function () {
				if (!this._Top) {
					this._Top=new Excel.TopBottomConditionalFormat(this.context, _createPropertyObjectPath(this.context, this, "TopBottomOrNullObject", false, false));
				}
				return this._Top;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ConditionalFormat.prototype, "id", {
			get: function () {
				_throwIfNotLoaded("id", this._Id0, _typeConditionalFormat, this._isNull);
				return this._Id0;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ConditionalFormat.prototype, "priority", {
			get: function () {
				_throwIfNotLoaded("priority", this._Pri, _typeConditionalFormat, this._isNull);
				return this._Pri;
			},
			set: function (value) {
				this._Pri=value;
				_createSetPropertyAction(this.context, this, "Priority", value);
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ConditionalFormat.prototype, "stopIfTrue", {
			get: function () {
				_throwIfNotLoaded("stopIfTrue", this._S, _typeConditionalFormat, this._isNull);
				return this._S;
			},
			set: function (value) {
				this._S=value;
				_createSetPropertyAction(this.context, this, "StopIfTrue", value);
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ConditionalFormat.prototype, "type", {
			get: function () {
				_throwIfNotLoaded("type", this._Ty, _typeConditionalFormat, this._isNull);
				return this._Ty;
			},
			enumerable: true,
			configurable: true
		});
		ConditionalFormat.prototype.set=function (properties, options) {
			this._recursivelySet(properties, options, ["stopIfTrue", "priority"], ["dataBarOrNullObject", "dataBar", "customOrNullObject", "custom", "iconSet", "iconSetOrNullObject", "colorScale", "colorScaleOrNullObject", "topBottom", "topBottomOrNullObject", "preset", "presetOrNullObject", "textComparison", "textComparisonOrNullObject", "cellValue", "cellValueOrNullObject"], []);
		};
		ConditionalFormat.prototype.delete=function () {
			_createMethodAction(this.context, this, "Delete", 0, []);
		};
		ConditionalFormat.prototype.getRange=function () {
			return new Excel.Range(this.context, _createMethodObjectPath(this.context, this, "GetRange", 1, [], false, true, null));
		};
		ConditionalFormat.prototype.getRangeOrNullObject=function () {
			return new Excel.Range(this.context, _createMethodObjectPath(this.context, this, "GetRangeOrNullObject", 1, [], false, true, null));
		};
		ConditionalFormat.prototype._handleResult=function (value) {
			_super.prototype._handleResult.call(this, value);
			if (_isNullOrUndefined(value))
				return;
			var obj=value;
			_fixObjectPathIfNecessary(this, obj);
			if (!_isUndefined(obj["Id"])) {
				this._Id0=obj["Id"];
			}
			if (!_isUndefined(obj["Priority"])) {
				this._Pri=obj["Priority"];
			}
			if (!_isUndefined(obj["StopIfTrue"])) {
				this._S=obj["StopIfTrue"];
			}
			if (!_isUndefined(obj["Type"])) {
				this._Ty=obj["Type"];
			}
			_handleNavigationPropertyResults(this, obj, ["cellValue", "CellValue", "cellValueOrNullObject", "CellValueOrNullObject", "colorScale", "ColorScale", "colorScaleOrNullObject", "ColorScaleOrNullObject", "custom", "Custom", "customOrNullObject", "CustomOrNullObject", "dataBar", "DataBar", "dataBarOrNullObject", "DataBarOrNullObject", "iconSet", "IconSet", "iconSetOrNullObject", "IconSetOrNullObject", "preset", "Preset", "presetOrNullObject", "PresetOrNullObject", "textComparison", "TextComparison", "textComparisonOrNullObject", "TextComparisonOrNullObject", "topBottom", "TopBottom", "topBottomOrNullObject", "TopBottomOrNullObject"]);
		};
		ConditionalFormat.prototype.load=function (option) {
			_load(this, option);
			return this;
		};
		ConditionalFormat.prototype._handleIdResult=function (value) {
			_super.prototype._handleIdResult.call(this, value);
			if (_isNullOrUndefined(value)) {
				return;
			}
			if (!_isUndefined(value["Id"])) {
				this._Id0=value["Id"];
			}
		};
		ConditionalFormat.prototype.toJSON=function () {
			return {
				"cellValue": this._C,
				"cellValueOrNullObject": this._Ce,
				"colorScale": this._Co,
				"colorScaleOrNullObject": this._Col,
				"custom": this._Cu,
				"customOrNullObject": this._Cus,
				"dataBar": this._D,
				"dataBarOrNullObject": this._Da,
				"iconSet": this._I,
				"iconSetOrNullObject": this._Ic,
				"id": this._Id0,
				"preset": this._P,
				"presetOrNullObject": this._Pr,
				"priority": this._Pri,
				"stopIfTrue": this._S,
				"textComparison": this._T,
				"textComparisonOrNullObject": this._Te,
				"topBottom": this._To,
				"topBottomOrNullObject": this._Top,
				"type": this._Ty
			};
		};
		return ConditionalFormat;
	}(OfficeExtension.ClientObject));
	Excel.ConditionalFormat=ConditionalFormat;
	var _typeDataBarConditionalFormat="DataBarConditionalFormat";
	var DataBarConditionalFormat=(function (_super) {
		__extends(DataBarConditionalFormat, _super);
		function DataBarConditionalFormat() {
			return _super !==null && _super.apply(this, arguments) || this;
		}
		Object.defineProperty(DataBarConditionalFormat.prototype, "_className", {
			get: function () {
				return "DataBarConditionalFormat";
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(DataBarConditionalFormat.prototype, "negativeFormat", {
			get: function () {
				if (!this._N) {
					this._N=new Excel.ConditionalDataBarNegativeFormat(this.context, _createPropertyObjectPath(this.context, this, "NegativeFormat", false, false));
				}
				return this._N;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(DataBarConditionalFormat.prototype, "positiveFormat", {
			get: function () {
				if (!this._P) {
					this._P=new Excel.ConditionalDataBarPositiveFormat(this.context, _createPropertyObjectPath(this.context, this, "PositiveFormat", false, false));
				}
				return this._P;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(DataBarConditionalFormat.prototype, "axisColor", {
			get: function () {
				_throwIfNotLoaded("axisColor", this._A, _typeDataBarConditionalFormat, this._isNull);
				return this._A;
			},
			set: function (value) {
				this._A=value;
				_createSetPropertyAction(this.context, this, "AxisColor", value);
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(DataBarConditionalFormat.prototype, "axisFormat", {
			get: function () {
				_throwIfNotLoaded("axisFormat", this._Ax, _typeDataBarConditionalFormat, this._isNull);
				return this._Ax;
			},
			set: function (value) {
				this._Ax=value;
				_createSetPropertyAction(this.context, this, "AxisFormat", value);
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(DataBarConditionalFormat.prototype, "barDirection", {
			get: function () {
				_throwIfNotLoaded("barDirection", this._B, _typeDataBarConditionalFormat, this._isNull);
				return this._B;
			},
			set: function (value) {
				this._B=value;
				_createSetPropertyAction(this.context, this, "BarDirection", value);
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(DataBarConditionalFormat.prototype, "lowerBoundRule", {
			get: function () {
				_throwIfNotLoaded("lowerBoundRule", this._L, _typeDataBarConditionalFormat, this._isNull);
				return this._L;
			},
			set: function (value) {
				this._L=value;
				_createSetPropertyAction(this.context, this, "LowerBoundRule", value);
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(DataBarConditionalFormat.prototype, "showDataBarOnly", {
			get: function () {
				_throwIfNotLoaded("showDataBarOnly", this._S, _typeDataBarConditionalFormat, this._isNull);
				return this._S;
			},
			set: function (value) {
				this._S=value;
				_createSetPropertyAction(this.context, this, "ShowDataBarOnly", value);
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(DataBarConditionalFormat.prototype, "upperBoundRule", {
			get: function () {
				_throwIfNotLoaded("upperBoundRule", this._U, _typeDataBarConditionalFormat, this._isNull);
				return this._U;
			},
			set: function (value) {
				this._U=value;
				_createSetPropertyAction(this.context, this, "UpperBoundRule", value);
			},
			enumerable: true,
			configurable: true
		});
		DataBarConditionalFormat.prototype.set=function (properties, options) {
			this._recursivelySet(properties, options, ["showDataBarOnly", "barDirection", "axisFormat", "axisColor", "lowerBoundRule", "upperBoundRule"], ["positiveFormat", "negativeFormat"], []);
		};
		DataBarConditionalFormat.prototype._handleResult=function (value) {
			_super.prototype._handleResult.call(this, value);
			if (_isNullOrUndefined(value))
				return;
			var obj=value;
			_fixObjectPathIfNecessary(this, obj);
			if (!_isUndefined(obj["AxisColor"])) {
				this._A=obj["AxisColor"];
			}
			if (!_isUndefined(obj["AxisFormat"])) {
				this._Ax=obj["AxisFormat"];
			}
			if (!_isUndefined(obj["BarDirection"])) {
				this._B=obj["BarDirection"];
			}
			if (!_isUndefined(obj["LowerBoundRule"])) {
				this._L=obj["LowerBoundRule"];
			}
			if (!_isUndefined(obj["ShowDataBarOnly"])) {
				this._S=obj["ShowDataBarOnly"];
			}
			if (!_isUndefined(obj["UpperBoundRule"])) {
				this._U=obj["UpperBoundRule"];
			}
			_handleNavigationPropertyResults(this, obj, ["negativeFormat", "NegativeFormat", "positiveFormat", "PositiveFormat"]);
		};
		DataBarConditionalFormat.prototype.load=function (option) {
			_load(this, option);
			return this;
		};
		DataBarConditionalFormat.prototype.toJSON=function () {
			return {
				"axisColor": this._A,
				"axisFormat": this._Ax,
				"barDirection": this._B,
				"lowerBoundRule": this._L,
				"negativeFormat": this._N,
				"positiveFormat": this._P,
				"showDataBarOnly": this._S,
				"upperBoundRule": this._U
			};
		};
		return DataBarConditionalFormat;
	}(OfficeExtension.ClientObject));
	Excel.DataBarConditionalFormat=DataBarConditionalFormat;
	var _typeConditionalDataBarPositiveFormat="ConditionalDataBarPositiveFormat";
	var ConditionalDataBarPositiveFormat=(function (_super) {
		__extends(ConditionalDataBarPositiveFormat, _super);
		function ConditionalDataBarPositiveFormat() {
			return _super !==null && _super.apply(this, arguments) || this;
		}
		Object.defineProperty(ConditionalDataBarPositiveFormat.prototype, "_className", {
			get: function () {
				return "ConditionalDataBarPositiveFormat";
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ConditionalDataBarPositiveFormat.prototype, "borderColor", {
			get: function () {
				_throwIfNotLoaded("borderColor", this._B, _typeConditionalDataBarPositiveFormat, this._isNull);
				return this._B;
			},
			set: function (value) {
				this._B=value;
				_createSetPropertyAction(this.context, this, "BorderColor", value);
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ConditionalDataBarPositiveFormat.prototype, "fillColor", {
			get: function () {
				_throwIfNotLoaded("fillColor", this._F, _typeConditionalDataBarPositiveFormat, this._isNull);
				return this._F;
			},
			set: function (value) {
				this._F=value;
				_createSetPropertyAction(this.context, this, "FillColor", value);
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ConditionalDataBarPositiveFormat.prototype, "gradientFill", {
			get: function () {
				_throwIfNotLoaded("gradientFill", this._G, _typeConditionalDataBarPositiveFormat, this._isNull);
				return this._G;
			},
			set: function (value) {
				this._G=value;
				_createSetPropertyAction(this.context, this, "GradientFill", value);
			},
			enumerable: true,
			configurable: true
		});
		ConditionalDataBarPositiveFormat.prototype.set=function (properties, options) {
			this._recursivelySet(properties, options, ["fillColor", "gradientFill", "borderColor"], [], []);
		};
		ConditionalDataBarPositiveFormat.prototype._handleResult=function (value) {
			_super.prototype._handleResult.call(this, value);
			if (_isNullOrUndefined(value))
				return;
			var obj=value;
			_fixObjectPathIfNecessary(this, obj);
			if (!_isUndefined(obj["BorderColor"])) {
				this._B=obj["BorderColor"];
			}
			if (!_isUndefined(obj["FillColor"])) {
				this._F=obj["FillColor"];
			}
			if (!_isUndefined(obj["GradientFill"])) {
				this._G=obj["GradientFill"];
			}
		};
		ConditionalDataBarPositiveFormat.prototype.load=function (option) {
			_load(this, option);
			return this;
		};
		ConditionalDataBarPositiveFormat.prototype.toJSON=function () {
			return {
				"borderColor": this._B,
				"fillColor": this._F,
				"gradientFill": this._G
			};
		};
		return ConditionalDataBarPositiveFormat;
	}(OfficeExtension.ClientObject));
	Excel.ConditionalDataBarPositiveFormat=ConditionalDataBarPositiveFormat;
	var _typeConditionalDataBarNegativeFormat="ConditionalDataBarNegativeFormat";
	var ConditionalDataBarNegativeFormat=(function (_super) {
		__extends(ConditionalDataBarNegativeFormat, _super);
		function ConditionalDataBarNegativeFormat() {
			return _super !==null && _super.apply(this, arguments) || this;
		}
		Object.defineProperty(ConditionalDataBarNegativeFormat.prototype, "_className", {
			get: function () {
				return "ConditionalDataBarNegativeFormat";
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ConditionalDataBarNegativeFormat.prototype, "borderColor", {
			get: function () {
				_throwIfNotLoaded("borderColor", this._B, _typeConditionalDataBarNegativeFormat, this._isNull);
				return this._B;
			},
			set: function (value) {
				this._B=value;
				_createSetPropertyAction(this.context, this, "BorderColor", value);
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ConditionalDataBarNegativeFormat.prototype, "fillColor", {
			get: function () {
				_throwIfNotLoaded("fillColor", this._F, _typeConditionalDataBarNegativeFormat, this._isNull);
				return this._F;
			},
			set: function (value) {
				this._F=value;
				_createSetPropertyAction(this.context, this, "FillColor", value);
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ConditionalDataBarNegativeFormat.prototype, "matchPositiveBorderColor", {
			get: function () {
				_throwIfNotLoaded("matchPositiveBorderColor", this._M, _typeConditionalDataBarNegativeFormat, this._isNull);
				return this._M;
			},
			set: function (value) {
				this._M=value;
				_createSetPropertyAction(this.context, this, "MatchPositiveBorderColor", value);
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ConditionalDataBarNegativeFormat.prototype, "matchPositiveFillColor", {
			get: function () {
				_throwIfNotLoaded("matchPositiveFillColor", this._Ma, _typeConditionalDataBarNegativeFormat, this._isNull);
				return this._Ma;
			},
			set: function (value) {
				this._Ma=value;
				_createSetPropertyAction(this.context, this, "MatchPositiveFillColor", value);
			},
			enumerable: true,
			configurable: true
		});
		ConditionalDataBarNegativeFormat.prototype.set=function (properties, options) {
			this._recursivelySet(properties, options, ["fillColor", "matchPositiveFillColor", "borderColor", "matchPositiveBorderColor"], [], []);
		};
		ConditionalDataBarNegativeFormat.prototype._handleResult=function (value) {
			_super.prototype._handleResult.call(this, value);
			if (_isNullOrUndefined(value))
				return;
			var obj=value;
			_fixObjectPathIfNecessary(this, obj);
			if (!_isUndefined(obj["BorderColor"])) {
				this._B=obj["BorderColor"];
			}
			if (!_isUndefined(obj["FillColor"])) {
				this._F=obj["FillColor"];
			}
			if (!_isUndefined(obj["MatchPositiveBorderColor"])) {
				this._M=obj["MatchPositiveBorderColor"];
			}
			if (!_isUndefined(obj["MatchPositiveFillColor"])) {
				this._Ma=obj["MatchPositiveFillColor"];
			}
		};
		ConditionalDataBarNegativeFormat.prototype.load=function (option) {
			_load(this, option);
			return this;
		};
		ConditionalDataBarNegativeFormat.prototype.toJSON=function () {
			return {
				"borderColor": this._B,
				"fillColor": this._F,
				"matchPositiveBorderColor": this._M,
				"matchPositiveFillColor": this._Ma
			};
		};
		return ConditionalDataBarNegativeFormat;
	}(OfficeExtension.ClientObject));
	Excel.ConditionalDataBarNegativeFormat=ConditionalDataBarNegativeFormat;
	var _typeCustomConditionalFormat="CustomConditionalFormat";
	var CustomConditionalFormat=(function (_super) {
		__extends(CustomConditionalFormat, _super);
		function CustomConditionalFormat() {
			return _super !==null && _super.apply(this, arguments) || this;
		}
		Object.defineProperty(CustomConditionalFormat.prototype, "_className", {
			get: function () {
				return "CustomConditionalFormat";
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(CustomConditionalFormat.prototype, "format", {
			get: function () {
				if (!this._F) {
					this._F=new Excel.ConditionalRangeFormat(this.context, _createPropertyObjectPath(this.context, this, "Format", false, false));
				}
				return this._F;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(CustomConditionalFormat.prototype, "rule", {
			get: function () {
				if (!this._R) {
					this._R=new Excel.ConditionalFormatRule(this.context, _createPropertyObjectPath(this.context, this, "Rule", false, false));
				}
				return this._R;
			},
			enumerable: true,
			configurable: true
		});
		CustomConditionalFormat.prototype.set=function (properties, options) {
			this._recursivelySet(properties, options, [], ["rule", "format"], []);
		};
		CustomConditionalFormat.prototype._handleResult=function (value) {
			_super.prototype._handleResult.call(this, value);
			if (_isNullOrUndefined(value))
				return;
			var obj=value;
			_fixObjectPathIfNecessary(this, obj);
			_handleNavigationPropertyResults(this, obj, ["format", "Format", "rule", "Rule"]);
		};
		CustomConditionalFormat.prototype.load=function (option) {
			_load(this, option);
			return this;
		};
		CustomConditionalFormat.prototype.toJSON=function () {
			return {
				"format": this._F,
				"rule": this._R
			};
		};
		return CustomConditionalFormat;
	}(OfficeExtension.ClientObject));
	Excel.CustomConditionalFormat=CustomConditionalFormat;
	var _typeConditionalFormatRule="ConditionalFormatRule";
	var ConditionalFormatRule=(function (_super) {
		__extends(ConditionalFormatRule, _super);
		function ConditionalFormatRule() {
			return _super !==null && _super.apply(this, arguments) || this;
		}
		Object.defineProperty(ConditionalFormatRule.prototype, "_className", {
			get: function () {
				return "ConditionalFormatRule";
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ConditionalFormatRule.prototype, "formula", {
			get: function () {
				_throwIfNotLoaded("formula", this._F, _typeConditionalFormatRule, this._isNull);
				return this._F;
			},
			set: function (value) {
				this._F=value;
				_createSetPropertyAction(this.context, this, "Formula", value);
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ConditionalFormatRule.prototype, "formulaLocal", {
			get: function () {
				_throwIfNotLoaded("formulaLocal", this._Fo, _typeConditionalFormatRule, this._isNull);
				return this._Fo;
			},
			set: function (value) {
				this._Fo=value;
				_createSetPropertyAction(this.context, this, "FormulaLocal", value);
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ConditionalFormatRule.prototype, "formulaR1C1", {
			get: function () {
				_throwIfNotLoaded("formulaR1C1", this._For, _typeConditionalFormatRule, this._isNull);
				return this._For;
			},
			set: function (value) {
				this._For=value;
				_createSetPropertyAction(this.context, this, "FormulaR1C1", value);
			},
			enumerable: true,
			configurable: true
		});
		ConditionalFormatRule.prototype.set=function (properties, options) {
			this._recursivelySet(properties, options, ["formula", "formulaLocal", "formulaR1C1"], [], []);
		};
		ConditionalFormatRule.prototype._handleResult=function (value) {
			_super.prototype._handleResult.call(this, value);
			if (_isNullOrUndefined(value))
				return;
			var obj=value;
			_fixObjectPathIfNecessary(this, obj);
			if (!_isUndefined(obj["Formula"])) {
				this._F=obj["Formula"];
			}
			if (!_isUndefined(obj["FormulaLocal"])) {
				this._Fo=obj["FormulaLocal"];
			}
			if (!_isUndefined(obj["FormulaR1C1"])) {
				this._For=obj["FormulaR1C1"];
			}
		};
		ConditionalFormatRule.prototype.load=function (option) {
			_load(this, option);
			return this;
		};
		ConditionalFormatRule.prototype.toJSON=function () {
			return {
				"formula": this._F,
				"formulaLocal": this._Fo,
				"formulaR1C1": this._For
			};
		};
		return ConditionalFormatRule;
	}(OfficeExtension.ClientObject));
	Excel.ConditionalFormatRule=ConditionalFormatRule;
	var _typeIconSetConditionalFormat="IconSetConditionalFormat";
	var IconSetConditionalFormat=(function (_super) {
		__extends(IconSetConditionalFormat, _super);
		function IconSetConditionalFormat() {
			return _super !==null && _super.apply(this, arguments) || this;
		}
		Object.defineProperty(IconSetConditionalFormat.prototype, "_className", {
			get: function () {
				return "IconSetConditionalFormat";
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(IconSetConditionalFormat.prototype, "criteria", {
			get: function () {
				_throwIfNotLoaded("criteria", this._C, _typeIconSetConditionalFormat, this._isNull);
				return this._C;
			},
			set: function (value) {
				this._C=value;
				_createSetPropertyAction(this.context, this, "Criteria", value);
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(IconSetConditionalFormat.prototype, "reverseIconOrder", {
			get: function () {
				_throwIfNotLoaded("reverseIconOrder", this._R, _typeIconSetConditionalFormat, this._isNull);
				return this._R;
			},
			set: function (value) {
				this._R=value;
				_createSetPropertyAction(this.context, this, "ReverseIconOrder", value);
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(IconSetConditionalFormat.prototype, "showIconOnly", {
			get: function () {
				_throwIfNotLoaded("showIconOnly", this._S, _typeIconSetConditionalFormat, this._isNull);
				return this._S;
			},
			set: function (value) {
				this._S=value;
				_createSetPropertyAction(this.context, this, "ShowIconOnly", value);
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(IconSetConditionalFormat.prototype, "style", {
			get: function () {
				_throwIfNotLoaded("style", this._St, _typeIconSetConditionalFormat, this._isNull);
				return this._St;
			},
			set: function (value) {
				this._St=value;
				_createSetPropertyAction(this.context, this, "Style", value);
			},
			enumerable: true,
			configurable: true
		});
		IconSetConditionalFormat.prototype.set=function (properties, options) {
			this._recursivelySet(properties, options, ["reverseIconOrder", "showIconOnly", "style", "criteria"], [], []);
		};
		IconSetConditionalFormat.prototype._handleResult=function (value) {
			_super.prototype._handleResult.call(this, value);
			if (_isNullOrUndefined(value))
				return;
			var obj=value;
			_fixObjectPathIfNecessary(this, obj);
			if (!_isUndefined(obj["Criteria"])) {
				this._C=obj["Criteria"];
			}
			if (!_isUndefined(obj["ReverseIconOrder"])) {
				this._R=obj["ReverseIconOrder"];
			}
			if (!_isUndefined(obj["ShowIconOnly"])) {
				this._S=obj["ShowIconOnly"];
			}
			if (!_isUndefined(obj["Style"])) {
				this._St=obj["Style"];
			}
		};
		IconSetConditionalFormat.prototype.load=function (option) {
			_load(this, option);
			return this;
		};
		IconSetConditionalFormat.prototype.toJSON=function () {
			return {
				"criteria": this._C,
				"reverseIconOrder": this._R,
				"showIconOnly": this._S,
				"style": this._St
			};
		};
		return IconSetConditionalFormat;
	}(OfficeExtension.ClientObject));
	Excel.IconSetConditionalFormat=IconSetConditionalFormat;
	var _typeColorScaleConditionalFormat="ColorScaleConditionalFormat";
	var ColorScaleConditionalFormat=(function (_super) {
		__extends(ColorScaleConditionalFormat, _super);
		function ColorScaleConditionalFormat() {
			return _super !==null && _super.apply(this, arguments) || this;
		}
		Object.defineProperty(ColorScaleConditionalFormat.prototype, "_className", {
			get: function () {
				return "ColorScaleConditionalFormat";
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ColorScaleConditionalFormat.prototype, "criteria", {
			get: function () {
				_throwIfNotLoaded("criteria", this._C, _typeColorScaleConditionalFormat, this._isNull);
				return this._C;
			},
			set: function (value) {
				this._C=value;
				_createSetPropertyAction(this.context, this, "Criteria", value);
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ColorScaleConditionalFormat.prototype, "threeColorScale", {
			get: function () {
				_throwIfNotLoaded("threeColorScale", this._T, _typeColorScaleConditionalFormat, this._isNull);
				return this._T;
			},
			enumerable: true,
			configurable: true
		});
		ColorScaleConditionalFormat.prototype.set=function (properties, options) {
			this._recursivelySet(properties, options, ["criteria"], [], []);
		};
		ColorScaleConditionalFormat.prototype._handleResult=function (value) {
			_super.prototype._handleResult.call(this, value);
			if (_isNullOrUndefined(value))
				return;
			var obj=value;
			_fixObjectPathIfNecessary(this, obj);
			if (!_isUndefined(obj["Criteria"])) {
				this._C=obj["Criteria"];
			}
			if (!_isUndefined(obj["ThreeColorScale"])) {
				this._T=obj["ThreeColorScale"];
			}
		};
		ColorScaleConditionalFormat.prototype.load=function (option) {
			_load(this, option);
			return this;
		};
		ColorScaleConditionalFormat.prototype.toJSON=function () {
			return {
				"criteria": this._C,
				"threeColorScale": this._T
			};
		};
		return ColorScaleConditionalFormat;
	}(OfficeExtension.ClientObject));
	Excel.ColorScaleConditionalFormat=ColorScaleConditionalFormat;
	var _typeTopBottomConditionalFormat="TopBottomConditionalFormat";
	var TopBottomConditionalFormat=(function (_super) {
		__extends(TopBottomConditionalFormat, _super);
		function TopBottomConditionalFormat() {
			return _super !==null && _super.apply(this, arguments) || this;
		}
		Object.defineProperty(TopBottomConditionalFormat.prototype, "_className", {
			get: function () {
				return "TopBottomConditionalFormat";
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(TopBottomConditionalFormat.prototype, "format", {
			get: function () {
				if (!this._F) {
					this._F=new Excel.ConditionalRangeFormat(this.context, _createPropertyObjectPath(this.context, this, "Format", false, false));
				}
				return this._F;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(TopBottomConditionalFormat.prototype, "rule", {
			get: function () {
				_throwIfNotLoaded("rule", this._R, _typeTopBottomConditionalFormat, this._isNull);
				return this._R;
			},
			set: function (value) {
				this._R=value;
				_createSetPropertyAction(this.context, this, "Rule", value);
			},
			enumerable: true,
			configurable: true
		});
		TopBottomConditionalFormat.prototype.set=function (properties, options) {
			this._recursivelySet(properties, options, ["rule"], ["format"], []);
		};
		TopBottomConditionalFormat.prototype._handleResult=function (value) {
			_super.prototype._handleResult.call(this, value);
			if (_isNullOrUndefined(value))
				return;
			var obj=value;
			_fixObjectPathIfNecessary(this, obj);
			if (!_isUndefined(obj["Rule"])) {
				this._R=obj["Rule"];
			}
			_handleNavigationPropertyResults(this, obj, ["format", "Format"]);
		};
		TopBottomConditionalFormat.prototype.load=function (option) {
			_load(this, option);
			return this;
		};
		TopBottomConditionalFormat.prototype.toJSON=function () {
			return {
				"format": this._F,
				"rule": this._R
			};
		};
		return TopBottomConditionalFormat;
	}(OfficeExtension.ClientObject));
	Excel.TopBottomConditionalFormat=TopBottomConditionalFormat;
	var _typePresetCriteriaConditionalFormat="PresetCriteriaConditionalFormat";
	var PresetCriteriaConditionalFormat=(function (_super) {
		__extends(PresetCriteriaConditionalFormat, _super);
		function PresetCriteriaConditionalFormat() {
			return _super !==null && _super.apply(this, arguments) || this;
		}
		Object.defineProperty(PresetCriteriaConditionalFormat.prototype, "_className", {
			get: function () {
				return "PresetCriteriaConditionalFormat";
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(PresetCriteriaConditionalFormat.prototype, "format", {
			get: function () {
				if (!this._F) {
					this._F=new Excel.ConditionalRangeFormat(this.context, _createPropertyObjectPath(this.context, this, "Format", false, false));
				}
				return this._F;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(PresetCriteriaConditionalFormat.prototype, "rule", {
			get: function () {
				_throwIfNotLoaded("rule", this._R, _typePresetCriteriaConditionalFormat, this._isNull);
				return this._R;
			},
			set: function (value) {
				this._R=value;
				_createSetPropertyAction(this.context, this, "Rule", value);
			},
			enumerable: true,
			configurable: true
		});
		PresetCriteriaConditionalFormat.prototype.set=function (properties, options) {
			this._recursivelySet(properties, options, ["rule"], ["format"], []);
		};
		PresetCriteriaConditionalFormat.prototype._handleResult=function (value) {
			_super.prototype._handleResult.call(this, value);
			if (_isNullOrUndefined(value))
				return;
			var obj=value;
			_fixObjectPathIfNecessary(this, obj);
			if (!_isUndefined(obj["Rule"])) {
				this._R=obj["Rule"];
			}
			_handleNavigationPropertyResults(this, obj, ["format", "Format"]);
		};
		PresetCriteriaConditionalFormat.prototype.load=function (option) {
			_load(this, option);
			return this;
		};
		PresetCriteriaConditionalFormat.prototype.toJSON=function () {
			return {
				"format": this._F,
				"rule": this._R
			};
		};
		return PresetCriteriaConditionalFormat;
	}(OfficeExtension.ClientObject));
	Excel.PresetCriteriaConditionalFormat=PresetCriteriaConditionalFormat;
	var _typeTextConditionalFormat="TextConditionalFormat";
	var TextConditionalFormat=(function (_super) {
		__extends(TextConditionalFormat, _super);
		function TextConditionalFormat() {
			return _super !==null && _super.apply(this, arguments) || this;
		}
		Object.defineProperty(TextConditionalFormat.prototype, "_className", {
			get: function () {
				return "TextConditionalFormat";
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(TextConditionalFormat.prototype, "format", {
			get: function () {
				if (!this._F) {
					this._F=new Excel.ConditionalRangeFormat(this.context, _createPropertyObjectPath(this.context, this, "Format", false, false));
				}
				return this._F;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(TextConditionalFormat.prototype, "rule", {
			get: function () {
				_throwIfNotLoaded("rule", this._R, _typeTextConditionalFormat, this._isNull);
				return this._R;
			},
			set: function (value) {
				this._R=value;
				_createSetPropertyAction(this.context, this, "Rule", value);
			},
			enumerable: true,
			configurable: true
		});
		TextConditionalFormat.prototype.set=function (properties, options) {
			this._recursivelySet(properties, options, ["rule"], ["format"], []);
		};
		TextConditionalFormat.prototype._handleResult=function (value) {
			_super.prototype._handleResult.call(this, value);
			if (_isNullOrUndefined(value))
				return;
			var obj=value;
			_fixObjectPathIfNecessary(this, obj);
			if (!_isUndefined(obj["Rule"])) {
				this._R=obj["Rule"];
			}
			_handleNavigationPropertyResults(this, obj, ["format", "Format"]);
		};
		TextConditionalFormat.prototype.load=function (option) {
			_load(this, option);
			return this;
		};
		TextConditionalFormat.prototype.toJSON=function () {
			return {
				"format": this._F,
				"rule": this._R
			};
		};
		return TextConditionalFormat;
	}(OfficeExtension.ClientObject));
	Excel.TextConditionalFormat=TextConditionalFormat;
	var _typeCellValueConditionalFormat="CellValueConditionalFormat";
	var CellValueConditionalFormat=(function (_super) {
		__extends(CellValueConditionalFormat, _super);
		function CellValueConditionalFormat() {
			return _super !==null && _super.apply(this, arguments) || this;
		}
		Object.defineProperty(CellValueConditionalFormat.prototype, "_className", {
			get: function () {
				return "CellValueConditionalFormat";
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(CellValueConditionalFormat.prototype, "format", {
			get: function () {
				if (!this._F) {
					this._F=new Excel.ConditionalRangeFormat(this.context, _createPropertyObjectPath(this.context, this, "Format", false, false));
				}
				return this._F;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(CellValueConditionalFormat.prototype, "rule", {
			get: function () {
				_throwIfNotLoaded("rule", this._R, _typeCellValueConditionalFormat, this._isNull);
				return this._R;
			},
			set: function (value) {
				this._R=value;
				_createSetPropertyAction(this.context, this, "Rule", value);
			},
			enumerable: true,
			configurable: true
		});
		CellValueConditionalFormat.prototype.set=function (properties, options) {
			this._recursivelySet(properties, options, ["rule"], ["format"], []);
		};
		CellValueConditionalFormat.prototype._handleResult=function (value) {
			_super.prototype._handleResult.call(this, value);
			if (_isNullOrUndefined(value))
				return;
			var obj=value;
			_fixObjectPathIfNecessary(this, obj);
			if (!_isUndefined(obj["Rule"])) {
				this._R=obj["Rule"];
			}
			_handleNavigationPropertyResults(this, obj, ["format", "Format"]);
		};
		CellValueConditionalFormat.prototype.load=function (option) {
			_load(this, option);
			return this;
		};
		CellValueConditionalFormat.prototype.toJSON=function () {
			return {
				"format": this._F,
				"rule": this._R
			};
		};
		return CellValueConditionalFormat;
	}(OfficeExtension.ClientObject));
	Excel.CellValueConditionalFormat=CellValueConditionalFormat;
	var _typeConditionalRangeFormat="ConditionalRangeFormat";
	var ConditionalRangeFormat=(function (_super) {
		__extends(ConditionalRangeFormat, _super);
		function ConditionalRangeFormat() {
			return _super !==null && _super.apply(this, arguments) || this;
		}
		Object.defineProperty(ConditionalRangeFormat.prototype, "_className", {
			get: function () {
				return "ConditionalRangeFormat";
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ConditionalRangeFormat.prototype, "borders", {
			get: function () {
				if (!this._B) {
					this._B=new Excel.ConditionalRangeBorderCollection(this.context, _createPropertyObjectPath(this.context, this, "Borders", true, false));
				}
				return this._B;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ConditionalRangeFormat.prototype, "fill", {
			get: function () {
				if (!this._F) {
					this._F=new Excel.ConditionalRangeFill(this.context, _createPropertyObjectPath(this.context, this, "Fill", false, false));
				}
				return this._F;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ConditionalRangeFormat.prototype, "font", {
			get: function () {
				if (!this._Fo) {
					this._Fo=new Excel.ConditionalRangeFont(this.context, _createPropertyObjectPath(this.context, this, "Font", false, false));
				}
				return this._Fo;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ConditionalRangeFormat.prototype, "numberFormat", {
			get: function () {
				_throwIfNotLoaded("numberFormat", this._N, _typeConditionalRangeFormat, this._isNull);
				return this._N;
			},
			set: function (value) {
				this._N=value;
				_createSetPropertyAction(this.context, this, "NumberFormat", value);
			},
			enumerable: true,
			configurable: true
		});
		ConditionalRangeFormat.prototype.set=function (properties, options) {
			this._recursivelySet(properties, options, ["numberFormat"], [], [
				"borders",
				"fill",
				"font",
				"borders",
				"fill",
				"font"
			]);
		};
		ConditionalRangeFormat.prototype._handleResult=function (value) {
			_super.prototype._handleResult.call(this, value);
			if (_isNullOrUndefined(value))
				return;
			var obj=value;
			_fixObjectPathIfNecessary(this, obj);
			if (!_isUndefined(obj["NumberFormat"])) {
				this._N=obj["NumberFormat"];
			}
			_handleNavigationPropertyResults(this, obj, ["borders", "Borders", "fill", "Fill", "font", "Font"]);
		};
		ConditionalRangeFormat.prototype.load=function (option) {
			_load(this, option);
			return this;
		};
		ConditionalRangeFormat.prototype.toJSON=function () {
			return {
				"numberFormat": this._N
			};
		};
		return ConditionalRangeFormat;
	}(OfficeExtension.ClientObject));
	Excel.ConditionalRangeFormat=ConditionalRangeFormat;
	var _typeConditionalRangeFont="ConditionalRangeFont";
	var ConditionalRangeFont=(function (_super) {
		__extends(ConditionalRangeFont, _super);
		function ConditionalRangeFont() {
			return _super !==null && _super.apply(this, arguments) || this;
		}
		Object.defineProperty(ConditionalRangeFont.prototype, "_className", {
			get: function () {
				return "ConditionalRangeFont";
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ConditionalRangeFont.prototype, "bold", {
			get: function () {
				_throwIfNotLoaded("bold", this._B, _typeConditionalRangeFont, this._isNull);
				return this._B;
			},
			set: function (value) {
				this._B=value;
				_createSetPropertyAction(this.context, this, "Bold", value);
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ConditionalRangeFont.prototype, "color", {
			get: function () {
				_throwIfNotLoaded("color", this._C, _typeConditionalRangeFont, this._isNull);
				return this._C;
			},
			set: function (value) {
				this._C=value;
				_createSetPropertyAction(this.context, this, "Color", value);
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ConditionalRangeFont.prototype, "italic", {
			get: function () {
				_throwIfNotLoaded("italic", this._I, _typeConditionalRangeFont, this._isNull);
				return this._I;
			},
			set: function (value) {
				this._I=value;
				_createSetPropertyAction(this.context, this, "Italic", value);
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ConditionalRangeFont.prototype, "strikethrough", {
			get: function () {
				_throwIfNotLoaded("strikethrough", this._S, _typeConditionalRangeFont, this._isNull);
				return this._S;
			},
			set: function (value) {
				this._S=value;
				_createSetPropertyAction(this.context, this, "Strikethrough", value);
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ConditionalRangeFont.prototype, "underline", {
			get: function () {
				_throwIfNotLoaded("underline", this._U, _typeConditionalRangeFont, this._isNull);
				return this._U;
			},
			set: function (value) {
				this._U=value;
				_createSetPropertyAction(this.context, this, "Underline", value);
			},
			enumerable: true,
			configurable: true
		});
		ConditionalRangeFont.prototype.set=function (properties, options) {
			this._recursivelySet(properties, options, ["color", "italic", "bold", "underline", "strikethrough"], [], []);
		};
		ConditionalRangeFont.prototype.clear=function () {
			_createMethodAction(this.context, this, "Clear", 0, []);
		};
		ConditionalRangeFont.prototype._handleResult=function (value) {
			_super.prototype._handleResult.call(this, value);
			if (_isNullOrUndefined(value))
				return;
			var obj=value;
			_fixObjectPathIfNecessary(this, obj);
			if (!_isUndefined(obj["Bold"])) {
				this._B=obj["Bold"];
			}
			if (!_isUndefined(obj["Color"])) {
				this._C=obj["Color"];
			}
			if (!_isUndefined(obj["Italic"])) {
				this._I=obj["Italic"];
			}
			if (!_isUndefined(obj["Strikethrough"])) {
				this._S=obj["Strikethrough"];
			}
			if (!_isUndefined(obj["Underline"])) {
				this._U=obj["Underline"];
			}
		};
		ConditionalRangeFont.prototype.load=function (option) {
			_load(this, option);
			return this;
		};
		ConditionalRangeFont.prototype.toJSON=function () {
			return {
				"bold": this._B,
				"color": this._C,
				"italic": this._I,
				"strikethrough": this._S,
				"underline": this._U
			};
		};
		return ConditionalRangeFont;
	}(OfficeExtension.ClientObject));
	Excel.ConditionalRangeFont=ConditionalRangeFont;
	var _typeConditionalRangeFill="ConditionalRangeFill";
	var ConditionalRangeFill=(function (_super) {
		__extends(ConditionalRangeFill, _super);
		function ConditionalRangeFill() {
			return _super !==null && _super.apply(this, arguments) || this;
		}
		Object.defineProperty(ConditionalRangeFill.prototype, "_className", {
			get: function () {
				return "ConditionalRangeFill";
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ConditionalRangeFill.prototype, "color", {
			get: function () {
				_throwIfNotLoaded("color", this._C, _typeConditionalRangeFill, this._isNull);
				return this._C;
			},
			set: function (value) {
				this._C=value;
				_createSetPropertyAction(this.context, this, "Color", value);
			},
			enumerable: true,
			configurable: true
		});
		ConditionalRangeFill.prototype.set=function (properties, options) {
			this._recursivelySet(properties, options, ["color"], [], []);
		};
		ConditionalRangeFill.prototype.clear=function () {
			_createMethodAction(this.context, this, "Clear", 0, []);
		};
		ConditionalRangeFill.prototype._handleResult=function (value) {
			_super.prototype._handleResult.call(this, value);
			if (_isNullOrUndefined(value))
				return;
			var obj=value;
			_fixObjectPathIfNecessary(this, obj);
			if (!_isUndefined(obj["Color"])) {
				this._C=obj["Color"];
			}
		};
		ConditionalRangeFill.prototype.load=function (option) {
			_load(this, option);
			return this;
		};
		ConditionalRangeFill.prototype.toJSON=function () {
			return {
				"color": this._C
			};
		};
		return ConditionalRangeFill;
	}(OfficeExtension.ClientObject));
	Excel.ConditionalRangeFill=ConditionalRangeFill;
	var _typeConditionalRangeBorder="ConditionalRangeBorder";
	var ConditionalRangeBorder=(function (_super) {
		__extends(ConditionalRangeBorder, _super);
		function ConditionalRangeBorder() {
			return _super !==null && _super.apply(this, arguments) || this;
		}
		Object.defineProperty(ConditionalRangeBorder.prototype, "_className", {
			get: function () {
				return "ConditionalRangeBorder";
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ConditionalRangeBorder.prototype, "color", {
			get: function () {
				_throwIfNotLoaded("color", this._C, _typeConditionalRangeBorder, this._isNull);
				return this._C;
			},
			set: function (value) {
				this._C=value;
				_createSetPropertyAction(this.context, this, "Color", value);
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ConditionalRangeBorder.prototype, "sideIndex", {
			get: function () {
				_throwIfNotLoaded("sideIndex", this._S, _typeConditionalRangeBorder, this._isNull);
				return this._S;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ConditionalRangeBorder.prototype, "style", {
			get: function () {
				_throwIfNotLoaded("style", this._St, _typeConditionalRangeBorder, this._isNull);
				return this._St;
			},
			set: function (value) {
				this._St=value;
				_createSetPropertyAction(this.context, this, "Style", value);
			},
			enumerable: true,
			configurable: true
		});
		ConditionalRangeBorder.prototype.set=function (properties, options) {
			this._recursivelySet(properties, options, ["style", "color"], [], []);
		};
		ConditionalRangeBorder.prototype._handleResult=function (value) {
			_super.prototype._handleResult.call(this, value);
			if (_isNullOrUndefined(value))
				return;
			var obj=value;
			_fixObjectPathIfNecessary(this, obj);
			if (!_isUndefined(obj["Color"])) {
				this._C=obj["Color"];
			}
			if (!_isUndefined(obj["SideIndex"])) {
				this._S=obj["SideIndex"];
			}
			if (!_isUndefined(obj["Style"])) {
				this._St=obj["Style"];
			}
		};
		ConditionalRangeBorder.prototype.load=function (option) {
			_load(this, option);
			return this;
		};
		ConditionalRangeBorder.prototype.toJSON=function () {
			return {
				"color": this._C,
				"sideIndex": this._S,
				"style": this._St
			};
		};
		return ConditionalRangeBorder;
	}(OfficeExtension.ClientObject));
	Excel.ConditionalRangeBorder=ConditionalRangeBorder;
	var _typeConditionalRangeBorderCollection="ConditionalRangeBorderCollection";
	var ConditionalRangeBorderCollection=(function (_super) {
		__extends(ConditionalRangeBorderCollection, _super);
		function ConditionalRangeBorderCollection() {
			return _super !==null && _super.apply(this, arguments) || this;
		}
		Object.defineProperty(ConditionalRangeBorderCollection.prototype, "_className", {
			get: function () {
				return "ConditionalRangeBorderCollection";
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ConditionalRangeBorderCollection.prototype, "bottom", {
			get: function () {
				if (!this._B) {
					this._B=new Excel.ConditionalRangeBorder(this.context, _createPropertyObjectPath(this.context, this, "Bottom", false, false));
				}
				return this._B;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ConditionalRangeBorderCollection.prototype, "left", {
			get: function () {
				if (!this._L) {
					this._L=new Excel.ConditionalRangeBorder(this.context, _createPropertyObjectPath(this.context, this, "Left", false, false));
				}
				return this._L;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ConditionalRangeBorderCollection.prototype, "right", {
			get: function () {
				if (!this._R) {
					this._R=new Excel.ConditionalRangeBorder(this.context, _createPropertyObjectPath(this.context, this, "Right", false, false));
				}
				return this._R;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ConditionalRangeBorderCollection.prototype, "top", {
			get: function () {
				if (!this._T) {
					this._T=new Excel.ConditionalRangeBorder(this.context, _createPropertyObjectPath(this.context, this, "Top", false, false));
				}
				return this._T;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ConditionalRangeBorderCollection.prototype, "items", {
			get: function () {
				_throwIfNotLoaded("items", this.m__items, _typeConditionalRangeBorderCollection, this._isNull);
				return this.m__items;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ConditionalRangeBorderCollection.prototype, "count", {
			get: function () {
				_throwIfNotLoaded("count", this._C, _typeConditionalRangeBorderCollection, this._isNull);
				return this._C;
			},
			enumerable: true,
			configurable: true
		});
		ConditionalRangeBorderCollection.prototype.getItem=function (index) {
			return new Excel.ConditionalRangeBorder(this.context, _createIndexerObjectPath(this.context, this, [index]));
		};
		ConditionalRangeBorderCollection.prototype.getItemAt=function (index) {
			return new Excel.ConditionalRangeBorder(this.context, _createMethodObjectPath(this.context, this, "GetItemAt", 1, [index], false, false, null));
		};
		ConditionalRangeBorderCollection.prototype._handleResult=function (value) {
			_super.prototype._handleResult.call(this, value);
			if (_isNullOrUndefined(value))
				return;
			var obj=value;
			_fixObjectPathIfNecessary(this, obj);
			if (!_isUndefined(obj["Count"])) {
				this._C=obj["Count"];
			}
			_handleNavigationPropertyResults(this, obj, ["bottom", "Bottom", "left", "Left", "right", "Right", "top", "Top"]);
			if (!_isNullOrUndefined(obj[OfficeExtension.Constants.items])) {
				this.m__items=[];
				var _data=obj[OfficeExtension.Constants.items];
				for (var i=0; i < _data.length; i++) {
					var _item=new Excel.ConditionalRangeBorder(this.context, _createChildItemObjectPathUsingIndexerOrGetItemAt(true, this.context, this, _data[i], i));
					_item._handleResult(_data[i]);
					this.m__items.push(_item);
				}
			}
		};
		ConditionalRangeBorderCollection.prototype.load=function (option) {
			_load(this, option);
			return this;
		};
		ConditionalRangeBorderCollection.prototype.toJSON=function () {
			return {
				"count": this._C
			};
		};
		return ConditionalRangeBorderCollection;
	}(OfficeExtension.ClientObject));
	Excel.ConditionalRangeBorderCollection=ConditionalRangeBorderCollection;
	var _typeInternalTest="InternalTest";
	var InternalTest=(function (_super) {
		__extends(InternalTest, _super);
		function InternalTest() {
			return _super !==null && _super.apply(this, arguments) || this;
		}
		Object.defineProperty(InternalTest.prototype, "_className", {
			get: function () {
				return "InternalTest";
			},
			enumerable: true,
			configurable: true
		});
		InternalTest.prototype.delay=function (seconds) {
			var action=_createMethodAction(this.context, this, "Delay", 0, [seconds]);
			var ret=new OfficeExtension.ClientResult();
			_addActionResultHandler(this, action, ret);
			return ret;
		};
		InternalTest.prototype._handleResult=function (value) {
			_super.prototype._handleResult.call(this, value);
			if (_isNullOrUndefined(value))
				return;
			var obj=value;
			_fixObjectPathIfNecessary(this, obj);
		};
		InternalTest.prototype.toJSON=function () {
			return {};
		};
		return InternalTest;
	}(OfficeExtension.ClientObject));
	Excel.InternalTest=InternalTest;
	var BindingType;
	(function (BindingType) {
		BindingType.range="Range";
		BindingType.table="Table";
		BindingType.text="Text";
	})(BindingType=Excel.BindingType || (Excel.BindingType={}));
	var BorderIndex;
	(function (BorderIndex) {
		BorderIndex.edgeTop="EdgeTop";
		BorderIndex.edgeBottom="EdgeBottom";
		BorderIndex.edgeLeft="EdgeLeft";
		BorderIndex.edgeRight="EdgeRight";
		BorderIndex.insideVertical="InsideVertical";
		BorderIndex.insideHorizontal="InsideHorizontal";
		BorderIndex.diagonalDown="DiagonalDown";
		BorderIndex.diagonalUp="DiagonalUp";
	})(BorderIndex=Excel.BorderIndex || (Excel.BorderIndex={}));
	var BorderLineStyle;
	(function (BorderLineStyle) {
		BorderLineStyle.none="None";
		BorderLineStyle.continuous="Continuous";
		BorderLineStyle.dash="Dash";
		BorderLineStyle.dashDot="DashDot";
		BorderLineStyle.dashDotDot="DashDotDot";
		BorderLineStyle.dot="Dot";
		BorderLineStyle.double="Double";
		BorderLineStyle.slantDashDot="SlantDashDot";
	})(BorderLineStyle=Excel.BorderLineStyle || (Excel.BorderLineStyle={}));
	var BorderWeight;
	(function (BorderWeight) {
		BorderWeight.hairline="Hairline";
		BorderWeight.thin="Thin";
		BorderWeight.medium="Medium";
		BorderWeight.thick="Thick";
	})(BorderWeight=Excel.BorderWeight || (Excel.BorderWeight={}));
	var CalculationMode;
	(function (CalculationMode) {
		CalculationMode.automatic="Automatic";
		CalculationMode.automaticExceptTables="AutomaticExceptTables";
		CalculationMode.manual="Manual";
	})(CalculationMode=Excel.CalculationMode || (Excel.CalculationMode={}));
	var CalculationType;
	(function (CalculationType) {
		CalculationType.recalculate="Recalculate";
		CalculationType.full="Full";
		CalculationType.fullRebuild="FullRebuild";
	})(CalculationType=Excel.CalculationType || (Excel.CalculationType={}));
	var ClearApplyTo;
	(function (ClearApplyTo) {
		ClearApplyTo.all="All";
		ClearApplyTo.formats="Formats";
		ClearApplyTo.contents="Contents";
		ClearApplyTo.hyperlinks="Hyperlinks";
		ClearApplyTo.removeHyperlinks="RemoveHyperlinks";
	})(ClearApplyTo=Excel.ClearApplyTo || (Excel.ClearApplyTo={}));
	var ChartDataLabelPosition;
	(function (ChartDataLabelPosition) {
		ChartDataLabelPosition.invalid="Invalid";
		ChartDataLabelPosition.none="None";
		ChartDataLabelPosition.center="Center";
		ChartDataLabelPosition.insideEnd="InsideEnd";
		ChartDataLabelPosition.insideBase="InsideBase";
		ChartDataLabelPosition.outsideEnd="OutsideEnd";
		ChartDataLabelPosition.left="Left";
		ChartDataLabelPosition.right="Right";
		ChartDataLabelPosition.top="Top";
		ChartDataLabelPosition.bottom="Bottom";
		ChartDataLabelPosition.bestFit="BestFit";
		ChartDataLabelPosition.callout="Callout";
	})(ChartDataLabelPosition=Excel.ChartDataLabelPosition || (Excel.ChartDataLabelPosition={}));
	var ChartLegendPosition;
	(function (ChartLegendPosition) {
		ChartLegendPosition.invalid="Invalid";
		ChartLegendPosition.top="Top";
		ChartLegendPosition.bottom="Bottom";
		ChartLegendPosition.left="Left";
		ChartLegendPosition.right="Right";
		ChartLegendPosition.corner="Corner";
		ChartLegendPosition.custom="Custom";
	})(ChartLegendPosition=Excel.ChartLegendPosition || (Excel.ChartLegendPosition={}));
	var ChartSeriesBy;
	(function (ChartSeriesBy) {
		ChartSeriesBy.auto="Auto";
		ChartSeriesBy.columns="Columns";
		ChartSeriesBy.rows="Rows";
	})(ChartSeriesBy=Excel.ChartSeriesBy || (Excel.ChartSeriesBy={}));
	var ChartType;
	(function (ChartType) {
		ChartType.invalid="Invalid";
		ChartType.columnClustered="ColumnClustered";
		ChartType.columnStacked="ColumnStacked";
		ChartType.columnStacked100="ColumnStacked100";
		ChartType._3DColumnClustered="3DColumnClustered";
		ChartType._3DColumnStacked="3DColumnStacked";
		ChartType._3DColumnStacked100="3DColumnStacked100";
		ChartType.barClustered="BarClustered";
		ChartType.barStacked="BarStacked";
		ChartType.barStacked100="BarStacked100";
		ChartType._3DBarClustered="3DBarClustered";
		ChartType._3DBarStacked="3DBarStacked";
		ChartType._3DBarStacked100="3DBarStacked100";
		ChartType.lineStacked="LineStacked";
		ChartType.lineStacked100="LineStacked100";
		ChartType.lineMarkers="LineMarkers";
		ChartType.lineMarkersStacked="LineMarkersStacked";
		ChartType.lineMarkersStacked100="LineMarkersStacked100";
		ChartType.pieOfPie="PieOfPie";
		ChartType.pieExploded="PieExploded";
		ChartType._3DPieExploded="3DPieExploded";
		ChartType.barOfPie="BarOfPie";
		ChartType.xyscatterSmooth="XYScatterSmooth";
		ChartType.xyscatterSmoothNoMarkers="XYScatterSmoothNoMarkers";
		ChartType.xyscatterLines="XYScatterLines";
		ChartType.xyscatterLinesNoMarkers="XYScatterLinesNoMarkers";
		ChartType.areaStacked="AreaStacked";
		ChartType.areaStacked100="AreaStacked100";
		ChartType._3DAreaStacked="3DAreaStacked";
		ChartType._3DAreaStacked100="3DAreaStacked100";
		ChartType.doughnutExploded="DoughnutExploded";
		ChartType.radarMarkers="RadarMarkers";
		ChartType.radarFilled="RadarFilled";
		ChartType.surface="Surface";
		ChartType.surfaceWireframe="SurfaceWireframe";
		ChartType.surfaceTopView="SurfaceTopView";
		ChartType.surfaceTopViewWireframe="SurfaceTopViewWireframe";
		ChartType.bubble="Bubble";
		ChartType.bubble3DEffect="Bubble3DEffect";
		ChartType.stockHLC="StockHLC";
		ChartType.stockOHLC="StockOHLC";
		ChartType.stockVHLC="StockVHLC";
		ChartType.stockVOHLC="StockVOHLC";
		ChartType.cylinderColClustered="CylinderColClustered";
		ChartType.cylinderColStacked="CylinderColStacked";
		ChartType.cylinderColStacked100="CylinderColStacked100";
		ChartType.cylinderBarClustered="CylinderBarClustered";
		ChartType.cylinderBarStacked="CylinderBarStacked";
		ChartType.cylinderBarStacked100="CylinderBarStacked100";
		ChartType.cylinderCol="CylinderCol";
		ChartType.coneColClustered="ConeColClustered";
		ChartType.coneColStacked="ConeColStacked";
		ChartType.coneColStacked100="ConeColStacked100";
		ChartType.coneBarClustered="ConeBarClustered";
		ChartType.coneBarStacked="ConeBarStacked";
		ChartType.coneBarStacked100="ConeBarStacked100";
		ChartType.coneCol="ConeCol";
		ChartType.pyramidColClustered="PyramidColClustered";
		ChartType.pyramidColStacked="PyramidColStacked";
		ChartType.pyramidColStacked100="PyramidColStacked100";
		ChartType.pyramidBarClustered="PyramidBarClustered";
		ChartType.pyramidBarStacked="PyramidBarStacked";
		ChartType.pyramidBarStacked100="PyramidBarStacked100";
		ChartType.pyramidCol="PyramidCol";
		ChartType._3DColumn="3DColumn";
		ChartType.line="Line";
		ChartType._3DLine="3DLine";
		ChartType._3DPie="3DPie";
		ChartType.pie="Pie";
		ChartType.xyscatter="XYScatter";
		ChartType._3DArea="3DArea";
		ChartType.area="Area";
		ChartType.doughnut="Doughnut";
		ChartType.radar="Radar";
	})(ChartType=Excel.ChartType || (Excel.ChartType={}));
	var ChartUnderlineStyle;
	(function (ChartUnderlineStyle) {
		ChartUnderlineStyle.none="None";
		ChartUnderlineStyle.single="Single";
	})(ChartUnderlineStyle=Excel.ChartUnderlineStyle || (Excel.ChartUnderlineStyle={}));
	var ConditionalDataBarAxisFormat;
	(function (ConditionalDataBarAxisFormat) {
		ConditionalDataBarAxisFormat.automatic="Automatic";
		ConditionalDataBarAxisFormat.none="None";
		ConditionalDataBarAxisFormat.cellMidPoint="CellMidPoint";
	})(ConditionalDataBarAxisFormat=Excel.ConditionalDataBarAxisFormat || (Excel.ConditionalDataBarAxisFormat={}));
	var ConditionalDataBarDirection;
	(function (ConditionalDataBarDirection) {
		ConditionalDataBarDirection.context="Context";
		ConditionalDataBarDirection.leftToRight="LeftToRight";
		ConditionalDataBarDirection.rightToLeft="RightToLeft";
	})(ConditionalDataBarDirection=Excel.ConditionalDataBarDirection || (Excel.ConditionalDataBarDirection={}));
	var ConditionalFormatDirection;
	(function (ConditionalFormatDirection) {
		ConditionalFormatDirection.top="Top";
		ConditionalFormatDirection.bottom="Bottom";
	})(ConditionalFormatDirection=Excel.ConditionalFormatDirection || (Excel.ConditionalFormatDirection={}));
	var ConditionalFormatType;
	(function (ConditionalFormatType) {
		ConditionalFormatType.custom="Custom";
		ConditionalFormatType.dataBar="DataBar";
		ConditionalFormatType.colorScale="ColorScale";
		ConditionalFormatType.iconSet="IconSet";
		ConditionalFormatType.topBottom="TopBottom";
		ConditionalFormatType.presetCriteria="PresetCriteria";
		ConditionalFormatType.containsText="ContainsText";
		ConditionalFormatType.cellValue="CellValue";
	})(ConditionalFormatType=Excel.ConditionalFormatType || (Excel.ConditionalFormatType={}));
	var ConditionalFormatRuleType;
	(function (ConditionalFormatRuleType) {
		ConditionalFormatRuleType.invalid="Invalid";
		ConditionalFormatRuleType.automatic="Automatic";
		ConditionalFormatRuleType.lowestValue="LowestValue";
		ConditionalFormatRuleType.highestValue="HighestValue";
		ConditionalFormatRuleType.number="Number";
		ConditionalFormatRuleType.percent="Percent";
		ConditionalFormatRuleType.formula="Formula";
		ConditionalFormatRuleType.percentile="Percentile";
	})(ConditionalFormatRuleType=Excel.ConditionalFormatRuleType || (Excel.ConditionalFormatRuleType={}));
	var ConditionalFormatIconRuleType;
	(function (ConditionalFormatIconRuleType) {
		ConditionalFormatIconRuleType.invalid="Invalid";
		ConditionalFormatIconRuleType.number="Number";
		ConditionalFormatIconRuleType.percent="Percent";
		ConditionalFormatIconRuleType.formula="Formula";
		ConditionalFormatIconRuleType.percentile="Percentile";
	})(ConditionalFormatIconRuleType=Excel.ConditionalFormatIconRuleType || (Excel.ConditionalFormatIconRuleType={}));
	var ConditionalFormatColorCriterionType;
	(function (ConditionalFormatColorCriterionType) {
		ConditionalFormatColorCriterionType.invalid="Invalid";
		ConditionalFormatColorCriterionType.lowestValue="LowestValue";
		ConditionalFormatColorCriterionType.highestValue="HighestValue";
		ConditionalFormatColorCriterionType.number="Number";
		ConditionalFormatColorCriterionType.percent="Percent";
		ConditionalFormatColorCriterionType.formula="Formula";
		ConditionalFormatColorCriterionType.percentile="Percentile";
	})(ConditionalFormatColorCriterionType=Excel.ConditionalFormatColorCriterionType || (Excel.ConditionalFormatColorCriterionType={}));
	var ConditionalTopBottomCriterionType;
	(function (ConditionalTopBottomCriterionType) {
		ConditionalTopBottomCriterionType.invalid="Invalid";
		ConditionalTopBottomCriterionType.topItems="TopItems";
		ConditionalTopBottomCriterionType.topPercent="TopPercent";
		ConditionalTopBottomCriterionType.bottomItems="BottomItems";
		ConditionalTopBottomCriterionType.bottomPercent="BottomPercent";
	})(ConditionalTopBottomCriterionType=Excel.ConditionalTopBottomCriterionType || (Excel.ConditionalTopBottomCriterionType={}));
	var ConditionalFormatPresetCriterion;
	(function (ConditionalFormatPresetCriterion) {
		ConditionalFormatPresetCriterion.invalid="Invalid";
		ConditionalFormatPresetCriterion.blanks="Blanks";
		ConditionalFormatPresetCriterion.nonBlanks="NonBlanks";
		ConditionalFormatPresetCriterion.errors="Errors";
		ConditionalFormatPresetCriterion.nonErrors="NonErrors";
		ConditionalFormatPresetCriterion.yesterday="Yesterday";
		ConditionalFormatPresetCriterion.today="Today";
		ConditionalFormatPresetCriterion.tomorrow="Tomorrow";
		ConditionalFormatPresetCriterion.lastSevenDays="LastSevenDays";
		ConditionalFormatPresetCriterion.lastWeek="LastWeek";
		ConditionalFormatPresetCriterion.thisWeek="ThisWeek";
		ConditionalFormatPresetCriterion.nextWeek="NextWeek";
		ConditionalFormatPresetCriterion.lastMonth="LastMonth";
		ConditionalFormatPresetCriterion.thisMonth="ThisMonth";
		ConditionalFormatPresetCriterion.nextMonth="NextMonth";
		ConditionalFormatPresetCriterion.aboveAverage="AboveAverage";
		ConditionalFormatPresetCriterion.belowAverage="BelowAverage";
		ConditionalFormatPresetCriterion.equalOrAboveAverage="EqualOrAboveAverage";
		ConditionalFormatPresetCriterion.equalOrBelowAverage="EqualOrBelowAverage";
		ConditionalFormatPresetCriterion.oneStdDevAboveAverage="OneStdDevAboveAverage";
		ConditionalFormatPresetCriterion.oneStdDevBelowAverage="OneStdDevBelowAverage";
		ConditionalFormatPresetCriterion.twoStdDevAboveAverage="TwoStdDevAboveAverage";
		ConditionalFormatPresetCriterion.twoStdDevBelowAverage="TwoStdDevBelowAverage";
		ConditionalFormatPresetCriterion.threeStdDevAboveAverage="ThreeStdDevAboveAverage";
		ConditionalFormatPresetCriterion.threeStdDevBelowAverage="ThreeStdDevBelowAverage";
		ConditionalFormatPresetCriterion.uniqueValues="UniqueValues";
		ConditionalFormatPresetCriterion.duplicateValues="DuplicateValues";
	})(ConditionalFormatPresetCriterion=Excel.ConditionalFormatPresetCriterion || (Excel.ConditionalFormatPresetCriterion={}));
	var ConditionalTextOperator;
	(function (ConditionalTextOperator) {
		ConditionalTextOperator.invalid="Invalid";
		ConditionalTextOperator.contains="Contains";
		ConditionalTextOperator.notContains="NotContains";
		ConditionalTextOperator.beginsWith="BeginsWith";
		ConditionalTextOperator.endsWith="EndsWith";
	})(ConditionalTextOperator=Excel.ConditionalTextOperator || (Excel.ConditionalTextOperator={}));
	var ConditionalCellValueOperator;
	(function (ConditionalCellValueOperator) {
		ConditionalCellValueOperator.invalid="Invalid";
		ConditionalCellValueOperator.between="Between";
		ConditionalCellValueOperator.notBetween="NotBetween";
		ConditionalCellValueOperator.equalTo="EqualTo";
		ConditionalCellValueOperator.notEqualTo="NotEqualTo";
		ConditionalCellValueOperator.greaterThan="GreaterThan";
		ConditionalCellValueOperator.lessThan="LessThan";
		ConditionalCellValueOperator.greaterThanOrEqual="GreaterThanOrEqual";
		ConditionalCellValueOperator.lessThanOrEqual="LessThanOrEqual";
	})(ConditionalCellValueOperator=Excel.ConditionalCellValueOperator || (Excel.ConditionalCellValueOperator={}));
	var ConditionalIconCriterionOperator;
	(function (ConditionalIconCriterionOperator) {
		ConditionalIconCriterionOperator.invalid="Invalid";
		ConditionalIconCriterionOperator.greaterThan="GreaterThan";
		ConditionalIconCriterionOperator.greaterThanOrEqual="GreaterThanOrEqual";
	})(ConditionalIconCriterionOperator=Excel.ConditionalIconCriterionOperator || (Excel.ConditionalIconCriterionOperator={}));
	var ConditionalRangeBorderIndex;
	(function (ConditionalRangeBorderIndex) {
		ConditionalRangeBorderIndex.edgeTop="EdgeTop";
		ConditionalRangeBorderIndex.edgeBottom="EdgeBottom";
		ConditionalRangeBorderIndex.edgeLeft="EdgeLeft";
		ConditionalRangeBorderIndex.edgeRight="EdgeRight";
	})(ConditionalRangeBorderIndex=Excel.ConditionalRangeBorderIndex || (Excel.ConditionalRangeBorderIndex={}));
	var ConditionalRangeBorderLineStyle;
	(function (ConditionalRangeBorderLineStyle) {
		ConditionalRangeBorderLineStyle.none="None";
		ConditionalRangeBorderLineStyle.continuous="Continuous";
		ConditionalRangeBorderLineStyle.dash="Dash";
		ConditionalRangeBorderLineStyle.dashDot="DashDot";
		ConditionalRangeBorderLineStyle.dashDotDot="DashDotDot";
		ConditionalRangeBorderLineStyle.dot="Dot";
	})(ConditionalRangeBorderLineStyle=Excel.ConditionalRangeBorderLineStyle || (Excel.ConditionalRangeBorderLineStyle={}));
	var ConditionalRangeFontUnderlineStyle;
	(function (ConditionalRangeFontUnderlineStyle) {
		ConditionalRangeFontUnderlineStyle.none="None";
		ConditionalRangeFontUnderlineStyle.single="Single";
		ConditionalRangeFontUnderlineStyle.double="Double";
	})(ConditionalRangeFontUnderlineStyle=Excel.ConditionalRangeFontUnderlineStyle || (Excel.ConditionalRangeFontUnderlineStyle={}));
	var DeleteShiftDirection;
	(function (DeleteShiftDirection) {
		DeleteShiftDirection.up="Up";
		DeleteShiftDirection.left="Left";
	})(DeleteShiftDirection=Excel.DeleteShiftDirection || (Excel.DeleteShiftDirection={}));
	var DynamicFilterCriteria;
	(function (DynamicFilterCriteria) {
		DynamicFilterCriteria.unknown="Unknown";
		DynamicFilterCriteria.aboveAverage="AboveAverage";
		DynamicFilterCriteria.allDatesInPeriodApril="AllDatesInPeriodApril";
		DynamicFilterCriteria.allDatesInPeriodAugust="AllDatesInPeriodAugust";
		DynamicFilterCriteria.allDatesInPeriodDecember="AllDatesInPeriodDecember";
		DynamicFilterCriteria.allDatesInPeriodFebruray="AllDatesInPeriodFebruray";
		DynamicFilterCriteria.allDatesInPeriodJanuary="AllDatesInPeriodJanuary";
		DynamicFilterCriteria.allDatesInPeriodJuly="AllDatesInPeriodJuly";
		DynamicFilterCriteria.allDatesInPeriodJune="AllDatesInPeriodJune";
		DynamicFilterCriteria.allDatesInPeriodMarch="AllDatesInPeriodMarch";
		DynamicFilterCriteria.allDatesInPeriodMay="AllDatesInPeriodMay";
		DynamicFilterCriteria.allDatesInPeriodNovember="AllDatesInPeriodNovember";
		DynamicFilterCriteria.allDatesInPeriodOctober="AllDatesInPeriodOctober";
		DynamicFilterCriteria.allDatesInPeriodQuarter1="AllDatesInPeriodQuarter1";
		DynamicFilterCriteria.allDatesInPeriodQuarter2="AllDatesInPeriodQuarter2";
		DynamicFilterCriteria.allDatesInPeriodQuarter3="AllDatesInPeriodQuarter3";
		DynamicFilterCriteria.allDatesInPeriodQuarter4="AllDatesInPeriodQuarter4";
		DynamicFilterCriteria.allDatesInPeriodSeptember="AllDatesInPeriodSeptember";
		DynamicFilterCriteria.belowAverage="BelowAverage";
		DynamicFilterCriteria.lastMonth="LastMonth";
		DynamicFilterCriteria.lastQuarter="LastQuarter";
		DynamicFilterCriteria.lastWeek="LastWeek";
		DynamicFilterCriteria.lastYear="LastYear";
		DynamicFilterCriteria.nextMonth="NextMonth";
		DynamicFilterCriteria.nextQuarter="NextQuarter";
		DynamicFilterCriteria.nextWeek="NextWeek";
		DynamicFilterCriteria.nextYear="NextYear";
		DynamicFilterCriteria.thisMonth="ThisMonth";
		DynamicFilterCriteria.thisQuarter="ThisQuarter";
		DynamicFilterCriteria.thisWeek="ThisWeek";
		DynamicFilterCriteria.thisYear="ThisYear";
		DynamicFilterCriteria.today="Today";
		DynamicFilterCriteria.tomorrow="Tomorrow";
		DynamicFilterCriteria.yearToDate="YearToDate";
		DynamicFilterCriteria.yesterday="Yesterday";
	})(DynamicFilterCriteria=Excel.DynamicFilterCriteria || (Excel.DynamicFilterCriteria={}));
	var FilterDatetimeSpecificity;
	(function (FilterDatetimeSpecificity) {
		FilterDatetimeSpecificity.year="Year";
		FilterDatetimeSpecificity.month="Month";
		FilterDatetimeSpecificity.day="Day";
		FilterDatetimeSpecificity.hour="Hour";
		FilterDatetimeSpecificity.minute="Minute";
		FilterDatetimeSpecificity.second="Second";
	})(FilterDatetimeSpecificity=Excel.FilterDatetimeSpecificity || (Excel.FilterDatetimeSpecificity={}));
	var FilterOn;
	(function (FilterOn) {
		FilterOn.bottomItems="BottomItems";
		FilterOn.bottomPercent="BottomPercent";
		FilterOn.cellColor="CellColor";
		FilterOn.dynamic="Dynamic";
		FilterOn.fontColor="FontColor";
		FilterOn.values="Values";
		FilterOn.topItems="TopItems";
		FilterOn.topPercent="TopPercent";
		FilterOn.icon="Icon";
		FilterOn.custom="Custom";
	})(FilterOn=Excel.FilterOn || (Excel.FilterOn={}));
	var FilterOperator;
	(function (FilterOperator) {
		FilterOperator.and="And";
		FilterOperator.or="Or";
	})(FilterOperator=Excel.FilterOperator || (Excel.FilterOperator={}));
	var HorizontalAlignment;
	(function (HorizontalAlignment) {
		HorizontalAlignment.general="General";
		HorizontalAlignment.left="Left";
		HorizontalAlignment.center="Center";
		HorizontalAlignment.right="Right";
		HorizontalAlignment.fill="Fill";
		HorizontalAlignment.justify="Justify";
		HorizontalAlignment.centerAcrossSelection="CenterAcrossSelection";
		HorizontalAlignment.distributed="Distributed";
	})(HorizontalAlignment=Excel.HorizontalAlignment || (Excel.HorizontalAlignment={}));
	var IconSet;
	(function (IconSet) {
		IconSet.invalid="Invalid";
		IconSet.threeArrows="ThreeArrows";
		IconSet.threeArrowsGray="ThreeArrowsGray";
		IconSet.threeFlags="ThreeFlags";
		IconSet.threeTrafficLights1="ThreeTrafficLights1";
		IconSet.threeTrafficLights2="ThreeTrafficLights2";
		IconSet.threeSigns="ThreeSigns";
		IconSet.threeSymbols="ThreeSymbols";
		IconSet.threeSymbols2="ThreeSymbols2";
		IconSet.fourArrows="FourArrows";
		IconSet.fourArrowsGray="FourArrowsGray";
		IconSet.fourRedToBlack="FourRedToBlack";
		IconSet.fourRating="FourRating";
		IconSet.fourTrafficLights="FourTrafficLights";
		IconSet.fiveArrows="FiveArrows";
		IconSet.fiveArrowsGray="FiveArrowsGray";
		IconSet.fiveRating="FiveRating";
		IconSet.fiveQuarters="FiveQuarters";
		IconSet.threeStars="ThreeStars";
		IconSet.threeTriangles="ThreeTriangles";
		IconSet.fiveBoxes="FiveBoxes";
	})(IconSet=Excel.IconSet || (Excel.IconSet={}));
	var ImageFittingMode;
	(function (ImageFittingMode) {
		ImageFittingMode.fit="Fit";
		ImageFittingMode.fitAndCenter="FitAndCenter";
		ImageFittingMode.fill="Fill";
	})(ImageFittingMode=Excel.ImageFittingMode || (Excel.ImageFittingMode={}));
	var InsertShiftDirection;
	(function (InsertShiftDirection) {
		InsertShiftDirection.down="Down";
		InsertShiftDirection.right="Right";
	})(InsertShiftDirection=Excel.InsertShiftDirection || (Excel.InsertShiftDirection={}));
	var NamedItemScope;
	(function (NamedItemScope) {
		NamedItemScope.worksheet="Worksheet";
		NamedItemScope.workbook="Workbook";
	})(NamedItemScope=Excel.NamedItemScope || (Excel.NamedItemScope={}));
	var NamedItemType;
	(function (NamedItemType) {
		NamedItemType.string="String";
		NamedItemType.integer="Integer";
		NamedItemType.double="Double";
		NamedItemType.boolean="Boolean";
		NamedItemType.range="Range";
		NamedItemType.error="Error";
		NamedItemType.array="Array";
	})(NamedItemType=Excel.NamedItemType || (Excel.NamedItemType={}));
	var RangeUnderlineStyle;
	(function (RangeUnderlineStyle) {
		RangeUnderlineStyle.none="None";
		RangeUnderlineStyle.single="Single";
		RangeUnderlineStyle.double="Double";
		RangeUnderlineStyle.singleAccountant="SingleAccountant";
		RangeUnderlineStyle.doubleAccountant="DoubleAccountant";
	})(RangeUnderlineStyle=Excel.RangeUnderlineStyle || (Excel.RangeUnderlineStyle={}));
	var SheetVisibility;
	(function (SheetVisibility) {
		SheetVisibility.visible="Visible";
		SheetVisibility.hidden="Hidden";
		SheetVisibility.veryHidden="VeryHidden";
	})(SheetVisibility=Excel.SheetVisibility || (Excel.SheetVisibility={}));
	var RangeValueType;
	(function (RangeValueType) {
		RangeValueType.unknown="Unknown";
		RangeValueType.empty="Empty";
		RangeValueType.string="String";
		RangeValueType.integer="Integer";
		RangeValueType.double="Double";
		RangeValueType.boolean="Boolean";
		RangeValueType.error="Error";
	})(RangeValueType=Excel.RangeValueType || (Excel.RangeValueType={}));
	var SortOrientation;
	(function (SortOrientation) {
		SortOrientation.rows="Rows";
		SortOrientation.columns="Columns";
	})(SortOrientation=Excel.SortOrientation || (Excel.SortOrientation={}));
	var SortOn;
	(function (SortOn) {
		SortOn.value="Value";
		SortOn.cellColor="CellColor";
		SortOn.fontColor="FontColor";
		SortOn.icon="Icon";
	})(SortOn=Excel.SortOn || (Excel.SortOn={}));
	var SortDataOption;
	(function (SortDataOption) {
		SortDataOption.normal="Normal";
		SortDataOption.textAsNumber="TextAsNumber";
	})(SortDataOption=Excel.SortDataOption || (Excel.SortDataOption={}));
	var SortMethod;
	(function (SortMethod) {
		SortMethod.pinYin="PinYin";
		SortMethod.strokeCount="StrokeCount";
	})(SortMethod=Excel.SortMethod || (Excel.SortMethod={}));
	var VerticalAlignment;
	(function (VerticalAlignment) {
		VerticalAlignment.top="Top";
		VerticalAlignment.center="Center";
		VerticalAlignment.bottom="Bottom";
		VerticalAlignment.justify="Justify";
		VerticalAlignment.distributed="Distributed";
	})(VerticalAlignment=Excel.VerticalAlignment || (Excel.VerticalAlignment={}));
	var _typeFunctionResult="FunctionResult";
	var FunctionResult=(function (_super) {
		__extends(FunctionResult, _super);
		function FunctionResult() {
			return _super !==null && _super.apply(this, arguments) || this;
		}
		Object.defineProperty(FunctionResult.prototype, "_className", {
			get: function () {
				return "FunctionResult<T>";
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(FunctionResult.prototype, "error", {
			get: function () {
				_throwIfNotLoaded("error", this._E, _typeFunctionResult, this._isNull);
				return this._E;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(FunctionResult.prototype, "value", {
			get: function () {
				_throwIfNotLoaded("value", this._V, _typeFunctionResult, this._isNull);
				return this._V;
			},
			enumerable: true,
			configurable: true
		});
		FunctionResult.prototype._handleResult=function (value) {
			_super.prototype._handleResult.call(this, value);
			if (_isNullOrUndefined(value))
				return;
			var obj=value;
			_fixObjectPathIfNecessary(this, obj);
			if (!_isUndefined(obj["Error"])) {
				this._E=obj["Error"];
			}
			if (!_isUndefined(obj["Value"])) {
				this._V=obj["Value"];
			}
		};
		FunctionResult.prototype.load=function (option) {
			_load(this, option);
			return this;
		};
		FunctionResult.prototype.toJSON=function () {
			return {
				"error": this._E,
				"value": this._V
			};
		};
		return FunctionResult;
	}(OfficeExtension.ClientObject));
	Excel.FunctionResult=FunctionResult;
	var _typeFunctions="Functions";
	var Functions=(function (_super) {
		__extends(Functions, _super);
		function Functions() {
			return _super !==null && _super.apply(this, arguments) || this;
		}
		Object.defineProperty(Functions.prototype, "_className", {
			get: function () {
				return "Functions";
			},
			enumerable: true,
			configurable: true
		});
		Functions.prototype.abs=function (number) {
			return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Abs", 0, [number], false, true, null));
		};
		Functions.prototype.accrInt=function (issue, firstInterest, settlement, rate, par, frequency, basis, calcMethod) {
			return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "AccrInt", 0, [issue, firstInterest, settlement, rate, par, frequency, basis, calcMethod], false, true, null));
		};
		Functions.prototype.accrIntM=function (issue, settlement, rate, par, basis) {
			return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "AccrIntM", 0, [issue, settlement, rate, par, basis], false, true, null));
		};
		Functions.prototype.acos=function (number) {
			return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Acos", 0, [number], false, true, null));
		};
		Functions.prototype.acosh=function (number) {
			return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Acosh", 0, [number], false, true, null));
		};
		Functions.prototype.acot=function (number) {
			return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Acot", 0, [number], false, true, null));
		};
		Functions.prototype.acoth=function (number) {
			return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Acoth", 0, [number], false, true, null));
		};
		Functions.prototype.amorDegrc=function (cost, datePurchased, firstPeriod, salvage, period, rate, basis) {
			return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "AmorDegrc", 0, [cost, datePurchased, firstPeriod, salvage, period, rate, basis], false, true, null));
		};
		Functions.prototype.amorLinc=function (cost, datePurchased, firstPeriod, salvage, period, rate, basis) {
			return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "AmorLinc", 0, [cost, datePurchased, firstPeriod, salvage, period, rate, basis], false, true, null));
		};
		Functions.prototype.and=function () {
			var values=[];
			for (var _i=0; _i < arguments.length; _i++) {
				values[_i]=arguments[_i];
			}
			return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "And", 0, [values], false, true, null));
		};
		Functions.prototype.arabic=function (text) {
			return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Arabic", 0, [text], false, true, null));
		};
		Functions.prototype.areas=function (reference) {
			return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Areas", 0, [reference], false, true, null));
		};
		Functions.prototype.asc=function (text) {
			return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Asc", 0, [text], false, true, null));
		};
		Functions.prototype.asin=function (number) {
			return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Asin", 0, [number], false, true, null));
		};
		Functions.prototype.asinh=function (number) {
			return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Asinh", 0, [number], false, true, null));
		};
		Functions.prototype.atan=function (number) {
			return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Atan", 0, [number], false, true, null));
		};
		Functions.prototype.atan2=function (xNum, yNum) {
			return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Atan2", 0, [xNum, yNum], false, true, null));
		};
		Functions.prototype.atanh=function (number) {
			return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Atanh", 0, [number], false, true, null));
		};
		Functions.prototype.aveDev=function () {
			var values=[];
			for (var _i=0; _i < arguments.length; _i++) {
				values[_i]=arguments[_i];
			}
			return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "AveDev", 0, [values], false, true, null));
		};
		Functions.prototype.average=function () {
			var values=[];
			for (var _i=0; _i < arguments.length; _i++) {
				values[_i]=arguments[_i];
			}
			return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Average", 0, [values], false, true, null));
		};
		Functions.prototype.averageA=function () {
			var values=[];
			for (var _i=0; _i < arguments.length; _i++) {
				values[_i]=arguments[_i];
			}
			return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "AverageA", 0, [values], false, true, null));
		};
		Functions.prototype.averageIf=function (range, criteria, averageRange) {
			return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "AverageIf", 0, [range, criteria, averageRange], false, true, null));
		};
		Functions.prototype.averageIfs=function (averageRange) {
			var values=[];
			for (var _i=1; _i < arguments.length; _i++) {
				values[_i - 1]=arguments[_i];
			}
			return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "AverageIfs", 0, [averageRange, values], false, true, null));
		};
		Functions.prototype.bahtText=function (number) {
			return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "BahtText", 0, [number], false, true, null));
		};
		Functions.prototype.base=function (number, radix, minLength) {
			return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Base", 0, [number, radix, minLength], false, true, null));
		};
		Functions.prototype.besselI=function (x, n) {
			return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "BesselI", 0, [x, n], false, true, null));
		};
		Functions.prototype.besselJ=function (x, n) {
			return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "BesselJ", 0, [x, n], false, true, null));
		};
		Functions.prototype.besselK=function (x, n) {
			return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "BesselK", 0, [x, n], false, true, null));
		};
		Functions.prototype.besselY=function (x, n) {
			return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "BesselY", 0, [x, n], false, true, null));
		};
		Functions.prototype.beta_Dist=function (x, alpha, beta, cumulative, A, B) {
			return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Beta_Dist", 0, [x, alpha, beta, cumulative, A, B], false, true, null));
		};
		Functions.prototype.beta_Inv=function (probability, alpha, beta, A, B) {
			return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Beta_Inv", 0, [probability, alpha, beta, A, B], false, true, null));
		};
		Functions.prototype.bin2Dec=function (number) {
			return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Bin2Dec", 0, [number], false, true, null));
		};
		Functions.prototype.bin2Hex=function (number, places) {
			return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Bin2Hex", 0, [number, places], false, true, null));
		};
		Functions.prototype.bin2Oct=function (number, places) {
			return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Bin2Oct", 0, [number, places], false, true, null));
		};
		Functions.prototype.binom_Dist=function (numberS, trials, probabilityS, cumulative) {
			return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Binom_Dist", 0, [numberS, trials, probabilityS, cumulative], false, true, null));
		};
		Functions.prototype.binom_Dist_Range=function (trials, probabilityS, numberS, numberS2) {
			return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Binom_Dist_Range", 0, [trials, probabilityS, numberS, numberS2], false, true, null));
		};
		Functions.prototype.binom_Inv=function (trials, probabilityS, alpha) {
			return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Binom_Inv", 0, [trials, probabilityS, alpha], false, true, null));
		};
		Functions.prototype.bitand=function (number1, number2) {
			return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Bitand", 0, [number1, number2], false, true, null));
		};
		Functions.prototype.bitlshift=function (number, shiftAmount) {
			return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Bitlshift", 0, [number, shiftAmount], false, true, null));
		};
		Functions.prototype.bitor=function (number1, number2) {
			return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Bitor", 0, [number1, number2], false, true, null));
		};
		Functions.prototype.bitrshift=function (number, shiftAmount) {
			return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Bitrshift", 0, [number, shiftAmount], false, true, null));
		};
		Functions.prototype.bitxor=function (number1, number2) {
			return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Bitxor", 0, [number1, number2], false, true, null));
		};
		Functions.prototype.ceiling_Math=function (number, significance, mode) {
			return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Ceiling_Math", 0, [number, significance, mode], false, true, null));
		};
		Functions.prototype.ceiling_Precise=function (number, significance) {
			return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Ceiling_Precise", 0, [number, significance], false, true, null));
		};
		Functions.prototype.char=function (number) {
			return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Char", 0, [number], false, true, null));
		};
		Functions.prototype.chiSq_Dist=function (x, degFreedom, cumulative) {
			return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "ChiSq_Dist", 0, [x, degFreedom, cumulative], false, true, null));
		};
		Functions.prototype.chiSq_Dist_RT=function (x, degFreedom) {
			return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "ChiSq_Dist_RT", 0, [x, degFreedom], false, true, null));
		};
		Functions.prototype.chiSq_Inv=function (probability, degFreedom) {
			return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "ChiSq_Inv", 0, [probability, degFreedom], false, true, null));
		};
		Functions.prototype.chiSq_Inv_RT=function (probability, degFreedom) {
			return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "ChiSq_Inv_RT", 0, [probability, degFreedom], false, true, null));
		};
		Functions.prototype.choose=function (indexNum) {
			var values=[];
			for (var _i=1; _i < arguments.length; _i++) {
				values[_i - 1]=arguments[_i];
			}
			return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Choose", 0, [indexNum, values], false, true, null));
		};
		Functions.prototype.clean=function (text) {
			return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Clean", 0, [text], false, true, null));
		};
		Functions.prototype.code=function (text) {
			return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Code", 0, [text], false, true, null));
		};
		Functions.prototype.columns=function (array) {
			return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Columns", 0, [array], false, true, null));
		};
		Functions.prototype.combin=function (number, numberChosen) {
			return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Combin", 0, [number, numberChosen], false, true, null));
		};
		Functions.prototype.combina=function (number, numberChosen) {
			return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Combina", 0, [number, numberChosen], false, true, null));
		};
		Functions.prototype.complex=function (realNum, iNum, suffix) {
			return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Complex", 0, [realNum, iNum, suffix], false, true, null));
		};
		Functions.prototype.concatenate=function () {
			var values=[];
			for (var _i=0; _i < arguments.length; _i++) {
				values[_i]=arguments[_i];
			}
			return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Concatenate", 0, [values], false, true, null));
		};
		Functions.prototype.confidence_Norm=function (alpha, standardDev, size) {
			return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Confidence_Norm", 0, [alpha, standardDev, size], false, true, null));
		};
		Functions.prototype.confidence_T=function (alpha, standardDev, size) {
			return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Confidence_T", 0, [alpha, standardDev, size], false, true, null));
		};
		Functions.prototype.convert=function (number, fromUnit, toUnit) {
			return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Convert", 0, [number, fromUnit, toUnit], false, true, null));
		};
		Functions.prototype.cos=function (number) {
			return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Cos", 0, [number], false, true, null));
		};
		Functions.prototype.cosh=function (number) {
			return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Cosh", 0, [number], false, true, null));
		};
		Functions.prototype.cot=function (number) {
			return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Cot", 0, [number], false, true, null));
		};
		Functions.prototype.coth=function (number) {
			return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Coth", 0, [number], false, true, null));
		};
		Functions.prototype.count=function () {
			var values=[];
			for (var _i=0; _i < arguments.length; _i++) {
				values[_i]=arguments[_i];
			}
			return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Count", 0, [values], false, true, null));
		};
		Functions.prototype.countA=function () {
			var values=[];
			for (var _i=0; _i < arguments.length; _i++) {
				values[_i]=arguments[_i];
			}
			return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "CountA", 0, [values], false, true, null));
		};
		Functions.prototype.countBlank=function (range) {
			return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "CountBlank", 0, [range], false, true, null));
		};
		Functions.prototype.countIf=function (range, criteria) {
			return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "CountIf", 0, [range, criteria], false, true, null));
		};
		Functions.prototype.countIfs=function () {
			var values=[];
			for (var _i=0; _i < arguments.length; _i++) {
				values[_i]=arguments[_i];
			}
			return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "CountIfs", 0, [values], false, true, null));
		};
		Functions.prototype.coupDayBs=function (settlement, maturity, frequency, basis) {
			return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "CoupDayBs", 0, [settlement, maturity, frequency, basis], false, true, null));
		};
		Functions.prototype.coupDays=function (settlement, maturity, frequency, basis) {
			return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "CoupDays", 0, [settlement, maturity, frequency, basis], false, true, null));
		};
		Functions.prototype.coupDaysNc=function (settlement, maturity, frequency, basis) {
			return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "CoupDaysNc", 0, [settlement, maturity, frequency, basis], false, true, null));
		};
		Functions.prototype.coupNcd=function (settlement, maturity, frequency, basis) {
			return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "CoupNcd", 0, [settlement, maturity, frequency, basis], false, true, null));
		};
		Functions.prototype.coupNum=function (settlement, maturity, frequency, basis) {
			return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "CoupNum", 0, [settlement, maturity, frequency, basis], false, true, null));
		};
		Functions.prototype.coupPcd=function (settlement, maturity, frequency, basis) {
			return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "CoupPcd", 0, [settlement, maturity, frequency, basis], false, true, null));
		};
		Functions.prototype.csc=function (number) {
			return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Csc", 0, [number], false, true, null));
		};
		Functions.prototype.csch=function (number) {
			return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Csch", 0, [number], false, true, null));
		};
		Functions.prototype.cumIPmt=function (rate, nper, pv, startPeriod, endPeriod, type) {
			return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "CumIPmt", 0, [rate, nper, pv, startPeriod, endPeriod, type], false, true, null));
		};
		Functions.prototype.cumPrinc=function (rate, nper, pv, startPeriod, endPeriod, type) {
			return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "CumPrinc", 0, [rate, nper, pv, startPeriod, endPeriod, type], false, true, null));
		};
		Functions.prototype.daverage=function (database, field, criteria) {
			return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "DAverage", 0, [database, field, criteria], false, true, null));
		};
		Functions.prototype.dcount=function (database, field, criteria) {
			return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "DCount", 0, [database, field, criteria], false, true, null));
		};
		Functions.prototype.dcountA=function (database, field, criteria) {
			return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "DCountA", 0, [database, field, criteria], false, true, null));
		};
		Functions.prototype.dget=function (database, field, criteria) {
			return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "DGet", 0, [database, field, criteria], false, true, null));
		};
		Functions.prototype.dmax=function (database, field, criteria) {
			return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "DMax", 0, [database, field, criteria], false, true, null));
		};
		Functions.prototype.dmin=function (database, field, criteria) {
			return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "DMin", 0, [database, field, criteria], false, true, null));
		};
		Functions.prototype.dproduct=function (database, field, criteria) {
			return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "DProduct", 0, [database, field, criteria], false, true, null));
		};
		Functions.prototype.dstDev=function (database, field, criteria) {
			return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "DStDev", 0, [database, field, criteria], false, true, null));
		};
		Functions.prototype.dstDevP=function (database, field, criteria) {
			return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "DStDevP", 0, [database, field, criteria], false, true, null));
		};
		Functions.prototype.dsum=function (database, field, criteria) {
			return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "DSum", 0, [database, field, criteria], false, true, null));
		};
		Functions.prototype.dvar=function (database, field, criteria) {
			return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "DVar", 0, [database, field, criteria], false, true, null));
		};
		Functions.prototype.dvarP=function (database, field, criteria) {
			return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "DVarP", 0, [database, field, criteria], false, true, null));
		};
		Functions.prototype.date=function (year, month, day) {
			return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Date", 0, [year, month, day], false, true, null));
		};
		Functions.prototype.datevalue=function (dateText) {
			return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Datevalue", 0, [dateText], false, true, null));
		};
		Functions.prototype.day=function (serialNumber) {
			return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Day", 0, [serialNumber], false, true, null));
		};
		Functions.prototype.days=function (endDate, startDate) {
			return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Days", 0, [endDate, startDate], false, true, null));
		};
		Functions.prototype.days360=function (startDate, endDate, method) {
			return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Days360", 0, [startDate, endDate, method], false, true, null));
		};
		Functions.prototype.db=function (cost, salvage, life, period, month) {
			return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Db", 0, [cost, salvage, life, period, month], false, true, null));
		};
		Functions.prototype.dbcs=function (text) {
			return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Dbcs", 0, [text], false, true, null));
		};
		Functions.prototype.ddb=function (cost, salvage, life, period, factor) {
			return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Ddb", 0, [cost, salvage, life, period, factor], false, true, null));
		};
		Functions.prototype.dec2Bin=function (number, places) {
			return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Dec2Bin", 0, [number, places], false, true, null));
		};
		Functions.prototype.dec2Hex=function (number, places) {
			return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Dec2Hex", 0, [number, places], false, true, null));
		};
		Functions.prototype.dec2Oct=function (number, places) {
			return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Dec2Oct", 0, [number, places], false, true, null));
		};
		Functions.prototype.decimal=function (number, radix) {
			return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Decimal", 0, [number, radix], false, true, null));
		};
		Functions.prototype.degrees=function (angle) {
			return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Degrees", 0, [angle], false, true, null));
		};
		Functions.prototype.delta=function (number1, number2) {
			return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Delta", 0, [number1, number2], false, true, null));
		};
		Functions.prototype.devSq=function () {
			var values=[];
			for (var _i=0; _i < arguments.length; _i++) {
				values[_i]=arguments[_i];
			}
			return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "DevSq", 0, [values], false, true, null));
		};
		Functions.prototype.disc=function (settlement, maturity, pr, redemption, basis) {
			return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Disc", 0, [settlement, maturity, pr, redemption, basis], false, true, null));
		};
		Functions.prototype.dollar=function (number, decimals) {
			return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Dollar", 0, [number, decimals], false, true, null));
		};
		Functions.prototype.dollarDe=function (fractionalDollar, fraction) {
			return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "DollarDe", 0, [fractionalDollar, fraction], false, true, null));
		};
		Functions.prototype.dollarFr=function (decimalDollar, fraction) {
			return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "DollarFr", 0, [decimalDollar, fraction], false, true, null));
		};
		Functions.prototype.duration=function (settlement, maturity, coupon, yld, frequency, basis) {
			return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Duration", 0, [settlement, maturity, coupon, yld, frequency, basis], false, true, null));
		};
		Functions.prototype.ecma_Ceiling=function (number, significance) {
			return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "ECMA_Ceiling", 0, [number, significance], false, true, null));
		};
		Functions.prototype.edate=function (startDate, months) {
			return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "EDate", 0, [startDate, months], false, true, null));
		};
		Functions.prototype.effect=function (nominalRate, npery) {
			return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Effect", 0, [nominalRate, npery], false, true, null));
		};
		Functions.prototype.eoMonth=function (startDate, months) {
			return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "EoMonth", 0, [startDate, months], false, true, null));
		};
		Functions.prototype.erf=function (lowerLimit, upperLimit) {
			return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Erf", 0, [lowerLimit, upperLimit], false, true, null));
		};
		Functions.prototype.erfC=function (x) {
			return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "ErfC", 0, [x], false, true, null));
		};
		Functions.prototype.erfC_Precise=function (X) {
			return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "ErfC_Precise", 0, [X], false, true, null));
		};
		Functions.prototype.erf_Precise=function (X) {
			return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Erf_Precise", 0, [X], false, true, null));
		};
		Functions.prototype.error_Type=function (errorVal) {
			return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Error_Type", 0, [errorVal], false, true, null));
		};
		Functions.prototype.even=function (number) {
			return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Even", 0, [number], false, true, null));
		};
		Functions.prototype.exact=function (text1, text2) {
			return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Exact", 0, [text1, text2], false, true, null));
		};
		Functions.prototype.exp=function (number) {
			return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Exp", 0, [number], false, true, null));
		};
		Functions.prototype.expon_Dist=function (x, lambda, cumulative) {
			return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Expon_Dist", 0, [x, lambda, cumulative], false, true, null));
		};
		Functions.prototype.fvschedule=function (principal, schedule) {
			return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "FVSchedule", 0, [principal, schedule], false, true, null));
		};
		Functions.prototype.f_Dist=function (x, degFreedom1, degFreedom2, cumulative) {
			return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "F_Dist", 0, [x, degFreedom1, degFreedom2, cumulative], false, true, null));
		};
		Functions.prototype.f_Dist_RT=function (x, degFreedom1, degFreedom2) {
			return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "F_Dist_RT", 0, [x, degFreedom1, degFreedom2], false, true, null));
		};
		Functions.prototype.f_Inv=function (probability, degFreedom1, degFreedom2) {
			return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "F_Inv", 0, [probability, degFreedom1, degFreedom2], false, true, null));
		};
		Functions.prototype.f_Inv_RT=function (probability, degFreedom1, degFreedom2) {
			return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "F_Inv_RT", 0, [probability, degFreedom1, degFreedom2], false, true, null));
		};
		Functions.prototype.fact=function (number) {
			return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Fact", 0, [number], false, true, null));
		};
		Functions.prototype.factDouble=function (number) {
			return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "FactDouble", 0, [number], false, true, null));
		};
		Functions.prototype.false=function () {
			return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "False", 0, [], false, true, null));
		};
		Functions.prototype.find=function (findText, withinText, startNum) {
			return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Find", 0, [findText, withinText, startNum], false, true, null));
		};
		Functions.prototype.findB=function (findText, withinText, startNum) {
			return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "FindB", 0, [findText, withinText, startNum], false, true, null));
		};
		Functions.prototype.fisher=function (x) {
			return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Fisher", 0, [x], false, true, null));
		};
		Functions.prototype.fisherInv=function (y) {
			return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "FisherInv", 0, [y], false, true, null));
		};
		Functions.prototype.fixed=function (number, decimals, noCommas) {
			return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Fixed", 0, [number, decimals, noCommas], false, true, null));
		};
		Functions.prototype.floor_Math=function (number, significance, mode) {
			return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Floor_Math", 0, [number, significance, mode], false, true, null));
		};
		Functions.prototype.floor_Precise=function (number, significance) {
			return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Floor_Precise", 0, [number, significance], false, true, null));
		};
		Functions.prototype.fv=function (rate, nper, pmt, pv, type) {
			return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Fv", 0, [rate, nper, pmt, pv, type], false, true, null));
		};
		Functions.prototype.gamma=function (x) {
			return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Gamma", 0, [x], false, true, null));
		};
		Functions.prototype.gammaLn=function (x) {
			return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "GammaLn", 0, [x], false, true, null));
		};
		Functions.prototype.gammaLn_Precise=function (x) {
			return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "GammaLn_Precise", 0, [x], false, true, null));
		};
		Functions.prototype.gamma_Dist=function (x, alpha, beta, cumulative) {
			return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Gamma_Dist", 0, [x, alpha, beta, cumulative], false, true, null));
		};
		Functions.prototype.gamma_Inv=function (probability, alpha, beta) {
			return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Gamma_Inv", 0, [probability, alpha, beta], false, true, null));
		};
		Functions.prototype.gauss=function (x) {
			return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Gauss", 0, [x], false, true, null));
		};
		Functions.prototype.gcd=function () {
			var values=[];
			for (var _i=0; _i < arguments.length; _i++) {
				values[_i]=arguments[_i];
			}
			return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Gcd", 0, [values], false, true, null));
		};
		Functions.prototype.geStep=function (number, step) {
			return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "GeStep", 0, [number, step], false, true, null));
		};
		Functions.prototype.geoMean=function () {
			var values=[];
			for (var _i=0; _i < arguments.length; _i++) {
				values[_i]=arguments[_i];
			}
			return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "GeoMean", 0, [values], false, true, null));
		};
		Functions.prototype.hlookup=function (lookupValue, tableArray, rowIndexNum, rangeLookup) {
			return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "HLookup", 0, [lookupValue, tableArray, rowIndexNum, rangeLookup], false, true, null));
		};
		Functions.prototype.harMean=function () {
			var values=[];
			for (var _i=0; _i < arguments.length; _i++) {
				values[_i]=arguments[_i];
			}
			return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "HarMean", 0, [values], false, true, null));
		};
		Functions.prototype.hex2Bin=function (number, places) {
			return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Hex2Bin", 0, [number, places], false, true, null));
		};
		Functions.prototype.hex2Dec=function (number) {
			return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Hex2Dec", 0, [number], false, true, null));
		};
		Functions.prototype.hex2Oct=function (number, places) {
			return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Hex2Oct", 0, [number, places], false, true, null));
		};
		Functions.prototype.hour=function (serialNumber) {
			return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Hour", 0, [serialNumber], false, true, null));
		};
		Functions.prototype.hypGeom_Dist=function (sampleS, numberSample, populationS, numberPop, cumulative) {
			return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "HypGeom_Dist", 0, [sampleS, numberSample, populationS, numberPop, cumulative], false, true, null));
		};
		Functions.prototype.hyperlink=function (linkLocation, friendlyName) {
			return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Hyperlink", 0, [linkLocation, friendlyName], false, true, null));
		};
		Functions.prototype.iso_Ceiling=function (number, significance) {
			return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "ISO_Ceiling", 0, [number, significance], false, true, null));
		};
		Functions.prototype.if=function (logicalTest, valueIfTrue, valueIfFalse) {
			return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "If", 0, [logicalTest, valueIfTrue, valueIfFalse], false, true, null));
		};
		Functions.prototype.imAbs=function (inumber) {
			return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "ImAbs", 0, [inumber], false, true, null));
		};
		Functions.prototype.imArgument=function (inumber) {
			return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "ImArgument", 0, [inumber], false, true, null));
		};
		Functions.prototype.imConjugate=function (inumber) {
			return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "ImConjugate", 0, [inumber], false, true, null));
		};
		Functions.prototype.imCos=function (inumber) {
			return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "ImCos", 0, [inumber], false, true, null));
		};
		Functions.prototype.imCosh=function (inumber) {
			return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "ImCosh", 0, [inumber], false, true, null));
		};
		Functions.prototype.imCot=function (inumber) {
			return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "ImCot", 0, [inumber], false, true, null));
		};
		Functions.prototype.imCsc=function (inumber) {
			return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "ImCsc", 0, [inumber], false, true, null));
		};
		Functions.prototype.imCsch=function (inumber) {
			return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "ImCsch", 0, [inumber], false, true, null));
		};
		Functions.prototype.imDiv=function (inumber1, inumber2) {
			return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "ImDiv", 0, [inumber1, inumber2], false, true, null));
		};
		Functions.prototype.imExp=function (inumber) {
			return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "ImExp", 0, [inumber], false, true, null));
		};
		Functions.prototype.imLn=function (inumber) {
			return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "ImLn", 0, [inumber], false, true, null));
		};
		Functions.prototype.imLog10=function (inumber) {
			return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "ImLog10", 0, [inumber], false, true, null));
		};
		Functions.prototype.imLog2=function (inumber) {
			return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "ImLog2", 0, [inumber], false, true, null));
		};
		Functions.prototype.imPower=function (inumber, number) {
			return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "ImPower", 0, [inumber, number], false, true, null));
		};
		Functions.prototype.imProduct=function () {
			var values=[];
			for (var _i=0; _i < arguments.length; _i++) {
				values[_i]=arguments[_i];
			}
			return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "ImProduct", 0, [values], false, true, null));
		};
		Functions.prototype.imReal=function (inumber) {
			return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "ImReal", 0, [inumber], false, true, null));
		};
		Functions.prototype.imSec=function (inumber) {
			return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "ImSec", 0, [inumber], false, true, null));
		};
		Functions.prototype.imSech=function (inumber) {
			return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "ImSech", 0, [inumber], false, true, null));
		};
		Functions.prototype.imSin=function (inumber) {
			return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "ImSin", 0, [inumber], false, true, null));
		};
		Functions.prototype.imSinh=function (inumber) {
			return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "ImSinh", 0, [inumber], false, true, null));
		};
		Functions.prototype.imSqrt=function (inumber) {
			return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "ImSqrt", 0, [inumber], false, true, null));
		};
		Functions.prototype.imSub=function (inumber1, inumber2) {
			return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "ImSub", 0, [inumber1, inumber2], false, true, null));
		};
		Functions.prototype.imSum=function () {
			var values=[];
			for (var _i=0; _i < arguments.length; _i++) {
				values[_i]=arguments[_i];
			}
			return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "ImSum", 0, [values], false, true, null));
		};
		Functions.prototype.imTan=function (inumber) {
			return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "ImTan", 0, [inumber], false, true, null));
		};
		Functions.prototype.imaginary=function (inumber) {
			return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Imaginary", 0, [inumber], false, true, null));
		};
		Functions.prototype.int=function (number) {
			return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Int", 0, [number], false, true, null));
		};
		Functions.prototype.intRate=function (settlement, maturity, investment, redemption, basis) {
			return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "IntRate", 0, [settlement, maturity, investment, redemption, basis], false, true, null));
		};
		Functions.prototype.ipmt=function (rate, per, nper, pv, fv, type) {
			return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Ipmt", 0, [rate, per, nper, pv, fv, type], false, true, null));
		};
		Functions.prototype.irr=function (values, guess) {
			return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Irr", 0, [values, guess], false, true, null));
		};
		Functions.prototype.isErr=function (value) {
			return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "IsErr", 0, [value], false, true, null));
		};
		Functions.prototype.isError=function (value) {
			return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "IsError", 0, [value], false, true, null));
		};
		Functions.prototype.isEven=function (number) {
			return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "IsEven", 0, [number], false, true, null));
		};
		Functions.prototype.isFormula=function (reference) {
			return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "IsFormula", 0, [reference], false, true, null));
		};
		Functions.prototype.isLogical=function (value) {
			return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "IsLogical", 0, [value], false, true, null));
		};
		Functions.prototype.isNA=function (value) {
			return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "IsNA", 0, [value], false, true, null));
		};
		Functions.prototype.isNonText=function (value) {
			return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "IsNonText", 0, [value], false, true, null));
		};
		Functions.prototype.isNumber=function (value) {
			return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "IsNumber", 0, [value], false, true, null));
		};
		Functions.prototype.isOdd=function (number) {
			return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "IsOdd", 0, [number], false, true, null));
		};
		Functions.prototype.isText=function (value) {
			return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "IsText", 0, [value], false, true, null));
		};
		Functions.prototype.isoWeekNum=function (date) {
			return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "IsoWeekNum", 0, [date], false, true, null));
		};
		Functions.prototype.ispmt=function (rate, per, nper, pv) {
			return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Ispmt", 0, [rate, per, nper, pv], false, true, null));
		};
		Functions.prototype.isref=function (value) {
			return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Isref", 0, [value], false, true, null));
		};
		Functions.prototype.kurt=function () {
			var values=[];
			for (var _i=0; _i < arguments.length; _i++) {
				values[_i]=arguments[_i];
			}
			return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Kurt", 0, [values], false, true, null));
		};
		Functions.prototype.large=function (array, k) {
			return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Large", 0, [array, k], false, true, null));
		};
		Functions.prototype.lcm=function () {
			var values=[];
			for (var _i=0; _i < arguments.length; _i++) {
				values[_i]=arguments[_i];
			}
			return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Lcm", 0, [values], false, true, null));
		};
		Functions.prototype.left=function (text, numChars) {
			return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Left", 0, [text, numChars], false, true, null));
		};
		Functions.prototype.leftb=function (text, numBytes) {
			return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Leftb", 0, [text, numBytes], false, true, null));
		};
		Functions.prototype.len=function (text) {
			return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Len", 0, [text], false, true, null));
		};
		Functions.prototype.lenb=function (text) {
			return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Lenb", 0, [text], false, true, null));
		};
		Functions.prototype.ln=function (number) {
			return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Ln", 0, [number], false, true, null));
		};
		Functions.prototype.log=function (number, base) {
			return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Log", 0, [number, base], false, true, null));
		};
		Functions.prototype.log10=function (number) {
			return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Log10", 0, [number], false, true, null));
		};
		Functions.prototype.logNorm_Dist=function (x, mean, standardDev, cumulative) {
			return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "LogNorm_Dist", 0, [x, mean, standardDev, cumulative], false, true, null));
		};
		Functions.prototype.logNorm_Inv=function (probability, mean, standardDev) {
			return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "LogNorm_Inv", 0, [probability, mean, standardDev], false, true, null));
		};
		Functions.prototype.lookup=function (lookupValue, lookupVector, resultVector) {
			return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Lookup", 0, [lookupValue, lookupVector, resultVector], false, true, null));
		};
		Functions.prototype.lower=function (text) {
			return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Lower", 0, [text], false, true, null));
		};
		Functions.prototype.mduration=function (settlement, maturity, coupon, yld, frequency, basis) {
			return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "MDuration", 0, [settlement, maturity, coupon, yld, frequency, basis], false, true, null));
		};
		Functions.prototype.mirr=function (values, financeRate, reinvestRate) {
			return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "MIrr", 0, [values, financeRate, reinvestRate], false, true, null));
		};
		Functions.prototype.mround=function (number, multiple) {
			return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "MRound", 0, [number, multiple], false, true, null));
		};
		Functions.prototype.match=function (lookupValue, lookupArray, matchType) {
			return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Match", 0, [lookupValue, lookupArray, matchType], false, true, null));
		};
		Functions.prototype.max=function () {
			var values=[];
			for (var _i=0; _i < arguments.length; _i++) {
				values[_i]=arguments[_i];
			}
			return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Max", 0, [values], false, true, null));
		};
		Functions.prototype.maxA=function () {
			var values=[];
			for (var _i=0; _i < arguments.length; _i++) {
				values[_i]=arguments[_i];
			}
			return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "MaxA", 0, [values], false, true, null));
		};
		Functions.prototype.median=function () {
			var values=[];
			for (var _i=0; _i < arguments.length; _i++) {
				values[_i]=arguments[_i];
			}
			return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Median", 0, [values], false, true, null));
		};
		Functions.prototype.mid=function (text, startNum, numChars) {
			return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Mid", 0, [text, startNum, numChars], false, true, null));
		};
		Functions.prototype.midb=function (text, startNum, numBytes) {
			return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Midb", 0, [text, startNum, numBytes], false, true, null));
		};
		Functions.prototype.min=function () {
			var values=[];
			for (var _i=0; _i < arguments.length; _i++) {
				values[_i]=arguments[_i];
			}
			return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Min", 0, [values], false, true, null));
		};
		Functions.prototype.minA=function () {
			var values=[];
			for (var _i=0; _i < arguments.length; _i++) {
				values[_i]=arguments[_i];
			}
			return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "MinA", 0, [values], false, true, null));
		};
		Functions.prototype.minute=function (serialNumber) {
			return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Minute", 0, [serialNumber], false, true, null));
		};
		Functions.prototype.mod=function (number, divisor) {
			return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Mod", 0, [number, divisor], false, true, null));
		};
		Functions.prototype.month=function (serialNumber) {
			return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Month", 0, [serialNumber], false, true, null));
		};
		Functions.prototype.multiNomial=function () {
			var values=[];
			for (var _i=0; _i < arguments.length; _i++) {
				values[_i]=arguments[_i];
			}
			return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "MultiNomial", 0, [values], false, true, null));
		};
		Functions.prototype.n=function (value) {
			return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "N", 0, [value], false, true, null));
		};
		Functions.prototype.nper=function (rate, pmt, pv, fv, type) {
			return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "NPer", 0, [rate, pmt, pv, fv, type], false, true, null));
		};
		Functions.prototype.na=function () {
			return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Na", 0, [], false, true, null));
		};
		Functions.prototype.negBinom_Dist=function (numberF, numberS, probabilityS, cumulative) {
			return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "NegBinom_Dist", 0, [numberF, numberS, probabilityS, cumulative], false, true, null));
		};
		Functions.prototype.networkDays=function (startDate, endDate, holidays) {
			return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "NetworkDays", 0, [startDate, endDate, holidays], false, true, null));
		};
		Functions.prototype.networkDays_Intl=function (startDate, endDate, weekend, holidays) {
			return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "NetworkDays_Intl", 0, [startDate, endDate, weekend, holidays], false, true, null));
		};
		Functions.prototype.nominal=function (effectRate, npery) {
			return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Nominal", 0, [effectRate, npery], false, true, null));
		};
		Functions.prototype.norm_Dist=function (x, mean, standardDev, cumulative) {
			return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Norm_Dist", 0, [x, mean, standardDev, cumulative], false, true, null));
		};
		Functions.prototype.norm_Inv=function (probability, mean, standardDev) {
			return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Norm_Inv", 0, [probability, mean, standardDev], false, true, null));
		};
		Functions.prototype.norm_S_Dist=function (z, cumulative) {
			return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Norm_S_Dist", 0, [z, cumulative], false, true, null));
		};
		Functions.prototype.norm_S_Inv=function (probability) {
			return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Norm_S_Inv", 0, [probability], false, true, null));
		};
		Functions.prototype.not=function (logical) {
			return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Not", 0, [logical], false, true, null));
		};
		Functions.prototype.now=function () {
			return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Now", 0, [], false, true, null));
		};
		Functions.prototype.npv=function (rate) {
			var values=[];
			for (var _i=1; _i < arguments.length; _i++) {
				values[_i - 1]=arguments[_i];
			}
			return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Npv", 0, [rate, values], false, true, null));
		};
		Functions.prototype.numberValue=function (text, decimalSeparator, groupSeparator) {
			return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "NumberValue", 0, [text, decimalSeparator, groupSeparator], false, true, null));
		};
		Functions.prototype.oct2Bin=function (number, places) {
			return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Oct2Bin", 0, [number, places], false, true, null));
		};
		Functions.prototype.oct2Dec=function (number) {
			return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Oct2Dec", 0, [number], false, true, null));
		};
		Functions.prototype.oct2Hex=function (number, places) {
			return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Oct2Hex", 0, [number, places], false, true, null));
		};
		Functions.prototype.odd=function (number) {
			return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Odd", 0, [number], false, true, null));
		};
		Functions.prototype.oddFPrice=function (settlement, maturity, issue, firstCoupon, rate, yld, redemption, frequency, basis) {
			return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "OddFPrice", 0, [settlement, maturity, issue, firstCoupon, rate, yld, redemption, frequency, basis], false, true, null));
		};
		Functions.prototype.oddFYield=function (settlement, maturity, issue, firstCoupon, rate, pr, redemption, frequency, basis) {
			return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "OddFYield", 0, [settlement, maturity, issue, firstCoupon, rate, pr, redemption, frequency, basis], false, true, null));
		};
		Functions.prototype.oddLPrice=function (settlement, maturity, lastInterest, rate, yld, redemption, frequency, basis) {
			return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "OddLPrice", 0, [settlement, maturity, lastInterest, rate, yld, redemption, frequency, basis], false, true, null));
		};
		Functions.prototype.oddLYield=function (settlement, maturity, lastInterest, rate, pr, redemption, frequency, basis) {
			return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "OddLYield", 0, [settlement, maturity, lastInterest, rate, pr, redemption, frequency, basis], false, true, null));
		};
		Functions.prototype.or=function () {
			var values=[];
			for (var _i=0; _i < arguments.length; _i++) {
				values[_i]=arguments[_i];
			}
			return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Or", 0, [values], false, true, null));
		};
		Functions.prototype.pduration=function (rate, pv, fv) {
			return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "PDuration", 0, [rate, pv, fv], false, true, null));
		};
		Functions.prototype.percentRank_Exc=function (array, x, significance) {
			return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "PercentRank_Exc", 0, [array, x, significance], false, true, null));
		};
		Functions.prototype.percentRank_Inc=function (array, x, significance) {
			return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "PercentRank_Inc", 0, [array, x, significance], false, true, null));
		};
		Functions.prototype.percentile_Exc=function (array, k) {
			return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Percentile_Exc", 0, [array, k], false, true, null));
		};
		Functions.prototype.percentile_Inc=function (array, k) {
			return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Percentile_Inc", 0, [array, k], false, true, null));
		};
		Functions.prototype.permut=function (number, numberChosen) {
			return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Permut", 0, [number, numberChosen], false, true, null));
		};
		Functions.prototype.permutationa=function (number, numberChosen) {
			return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Permutationa", 0, [number, numberChosen], false, true, null));
		};
		Functions.prototype.phi=function (x) {
			return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Phi", 0, [x], false, true, null));
		};
		Functions.prototype.pi=function () {
			return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Pi", 0, [], false, true, null));
		};
		Functions.prototype.pmt=function (rate, nper, pv, fv, type) {
			return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Pmt", 0, [rate, nper, pv, fv, type], false, true, null));
		};
		Functions.prototype.poisson_Dist=function (x, mean, cumulative) {
			return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Poisson_Dist", 0, [x, mean, cumulative], false, true, null));
		};
		Functions.prototype.power=function (number, power) {
			return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Power", 0, [number, power], false, true, null));
		};
		Functions.prototype.ppmt=function (rate, per, nper, pv, fv, type) {
			return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Ppmt", 0, [rate, per, nper, pv, fv, type], false, true, null));
		};
		Functions.prototype.price=function (settlement, maturity, rate, yld, redemption, frequency, basis) {
			return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Price", 0, [settlement, maturity, rate, yld, redemption, frequency, basis], false, true, null));
		};
		Functions.prototype.priceDisc=function (settlement, maturity, discount, redemption, basis) {
			return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "PriceDisc", 0, [settlement, maturity, discount, redemption, basis], false, true, null));
		};
		Functions.prototype.priceMat=function (settlement, maturity, issue, rate, yld, basis) {
			return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "PriceMat", 0, [settlement, maturity, issue, rate, yld, basis], false, true, null));
		};
		Functions.prototype.product=function () {
			var values=[];
			for (var _i=0; _i < arguments.length; _i++) {
				values[_i]=arguments[_i];
			}
			return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Product", 0, [values], false, true, null));
		};
		Functions.prototype.proper=function (text) {
			return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Proper", 0, [text], false, true, null));
		};
		Functions.prototype.pv=function (rate, nper, pmt, fv, type) {
			return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Pv", 0, [rate, nper, pmt, fv, type], false, true, null));
		};
		Functions.prototype.quartile_Exc=function (array, quart) {
			return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Quartile_Exc", 0, [array, quart], false, true, null));
		};
		Functions.prototype.quartile_Inc=function (array, quart) {
			return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Quartile_Inc", 0, [array, quart], false, true, null));
		};
		Functions.prototype.quotient=function (numerator, denominator) {
			return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Quotient", 0, [numerator, denominator], false, true, null));
		};
		Functions.prototype.radians=function (angle) {
			return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Radians", 0, [angle], false, true, null));
		};
		Functions.prototype.rand=function () {
			return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Rand", 0, [], false, true, null));
		};
		Functions.prototype.randBetween=function (bottom, top) {
			return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "RandBetween", 0, [bottom, top], false, true, null));
		};
		Functions.prototype.rank_Avg=function (number, ref, order) {
			return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Rank_Avg", 0, [number, ref, order], false, true, null));
		};
		Functions.prototype.rank_Eq=function (number, ref, order) {
			return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Rank_Eq", 0, [number, ref, order], false, true, null));
		};
		Functions.prototype.rate=function (nper, pmt, pv, fv, type, guess) {
			return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Rate", 0, [nper, pmt, pv, fv, type, guess], false, true, null));
		};
		Functions.prototype.received=function (settlement, maturity, investment, discount, basis) {
			return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Received", 0, [settlement, maturity, investment, discount, basis], false, true, null));
		};
		Functions.prototype.replace=function (oldText, startNum, numChars, newText) {
			return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Replace", 0, [oldText, startNum, numChars, newText], false, true, null));
		};
		Functions.prototype.replaceB=function (oldText, startNum, numBytes, newText) {
			return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "ReplaceB", 0, [oldText, startNum, numBytes, newText], false, true, null));
		};
		Functions.prototype.rept=function (text, numberTimes) {
			return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Rept", 0, [text, numberTimes], false, true, null));
		};
		Functions.prototype.right=function (text, numChars) {
			return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Right", 0, [text, numChars], false, true, null));
		};
		Functions.prototype.rightb=function (text, numBytes) {
			return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Rightb", 0, [text, numBytes], false, true, null));
		};
		Functions.prototype.roman=function (number, form) {
			return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Roman", 0, [number, form], false, true, null));
		};
		Functions.prototype.round=function (number, numDigits) {
			return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Round", 0, [number, numDigits], false, true, null));
		};
		Functions.prototype.roundDown=function (number, numDigits) {
			return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "RoundDown", 0, [number, numDigits], false, true, null));
		};
		Functions.prototype.roundUp=function (number, numDigits) {
			return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "RoundUp", 0, [number, numDigits], false, true, null));
		};
		Functions.prototype.rows=function (array) {
			return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Rows", 0, [array], false, true, null));
		};
		Functions.prototype.rri=function (nper, pv, fv) {
			return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Rri", 0, [nper, pv, fv], false, true, null));
		};
		Functions.prototype.sec=function (number) {
			return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Sec", 0, [number], false, true, null));
		};
		Functions.prototype.sech=function (number) {
			return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Sech", 0, [number], false, true, null));
		};
		Functions.prototype.second=function (serialNumber) {
			return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Second", 0, [serialNumber], false, true, null));
		};
		Functions.prototype.seriesSum=function (x, n, m, coefficients) {
			return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "SeriesSum", 0, [x, n, m, coefficients], false, true, null));
		};
		Functions.prototype.sheet=function (value) {
			return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Sheet", 0, [value], false, true, null));
		};
		Functions.prototype.sheets=function (reference) {
			return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Sheets", 0, [reference], false, true, null));
		};
		Functions.prototype.sign=function (number) {
			return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Sign", 0, [number], false, true, null));
		};
		Functions.prototype.sin=function (number) {
			return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Sin", 0, [number], false, true, null));
		};
		Functions.prototype.sinh=function (number) {
			return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Sinh", 0, [number], false, true, null));
		};
		Functions.prototype.skew=function () {
			var values=[];
			for (var _i=0; _i < arguments.length; _i++) {
				values[_i]=arguments[_i];
			}
			return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Skew", 0, [values], false, true, null));
		};
		Functions.prototype.skew_p=function () {
			var values=[];
			for (var _i=0; _i < arguments.length; _i++) {
				values[_i]=arguments[_i];
			}
			return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Skew_p", 0, [values], false, true, null));
		};
		Functions.prototype.sln=function (cost, salvage, life) {
			return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Sln", 0, [cost, salvage, life], false, true, null));
		};
		Functions.prototype.small=function (array, k) {
			return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Small", 0, [array, k], false, true, null));
		};
		Functions.prototype.sqrt=function (number) {
			return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Sqrt", 0, [number], false, true, null));
		};
		Functions.prototype.sqrtPi=function (number) {
			return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "SqrtPi", 0, [number], false, true, null));
		};
		Functions.prototype.stDevA=function () {
			var values=[];
			for (var _i=0; _i < arguments.length; _i++) {
				values[_i]=arguments[_i];
			}
			return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "StDevA", 0, [values], false, true, null));
		};
		Functions.prototype.stDevPA=function () {
			var values=[];
			for (var _i=0; _i < arguments.length; _i++) {
				values[_i]=arguments[_i];
			}
			return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "StDevPA", 0, [values], false, true, null));
		};
		Functions.prototype.stDev_P=function () {
			var values=[];
			for (var _i=0; _i < arguments.length; _i++) {
				values[_i]=arguments[_i];
			}
			return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "StDev_P", 0, [values], false, true, null));
		};
		Functions.prototype.stDev_S=function () {
			var values=[];
			for (var _i=0; _i < arguments.length; _i++) {
				values[_i]=arguments[_i];
			}
			return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "StDev_S", 0, [values], false, true, null));
		};
		Functions.prototype.standardize=function (x, mean, standardDev) {
			return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Standardize", 0, [x, mean, standardDev], false, true, null));
		};
		Functions.prototype.substitute=function (text, oldText, newText, instanceNum) {
			return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Substitute", 0, [text, oldText, newText, instanceNum], false, true, null));
		};
		Functions.prototype.subtotal=function (functionNum) {
			var values=[];
			for (var _i=1; _i < arguments.length; _i++) {
				values[_i - 1]=arguments[_i];
			}
			return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Subtotal", 0, [functionNum, values], false, true, null));
		};
		Functions.prototype.sum=function () {
			var values=[];
			for (var _i=0; _i < arguments.length; _i++) {
				values[_i]=arguments[_i];
			}
			return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Sum", 0, [values], false, true, null));
		};
		Functions.prototype.sumIf=function (range, criteria, sumRange) {
			return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "SumIf", 0, [range, criteria, sumRange], false, true, null));
		};
		Functions.prototype.sumIfs=function (sumRange) {
			var values=[];
			for (var _i=1; _i < arguments.length; _i++) {
				values[_i - 1]=arguments[_i];
			}
			return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "SumIfs", 0, [sumRange, values], false, true, null));
		};
		Functions.prototype.sumSq=function () {
			var values=[];
			for (var _i=0; _i < arguments.length; _i++) {
				values[_i]=arguments[_i];
			}
			return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "SumSq", 0, [values], false, true, null));
		};
		Functions.prototype.syd=function (cost, salvage, life, per) {
			return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Syd", 0, [cost, salvage, life, per], false, true, null));
		};
		Functions.prototype.t=function (value) {
			return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "T", 0, [value], false, true, null));
		};
		Functions.prototype.tbillEq=function (settlement, maturity, discount) {
			return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "TBillEq", 0, [settlement, maturity, discount], false, true, null));
		};
		Functions.prototype.tbillPrice=function (settlement, maturity, discount) {
			return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "TBillPrice", 0, [settlement, maturity, discount], false, true, null));
		};
		Functions.prototype.tbillYield=function (settlement, maturity, pr) {
			return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "TBillYield", 0, [settlement, maturity, pr], false, true, null));
		};
		Functions.prototype.t_Dist=function (x, degFreedom, cumulative) {
			return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "T_Dist", 0, [x, degFreedom, cumulative], false, true, null));
		};
		Functions.prototype.t_Dist_2T=function (x, degFreedom) {
			return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "T_Dist_2T", 0, [x, degFreedom], false, true, null));
		};
		Functions.prototype.t_Dist_RT=function (x, degFreedom) {
			return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "T_Dist_RT", 0, [x, degFreedom], false, true, null));
		};
		Functions.prototype.t_Inv=function (probability, degFreedom) {
			return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "T_Inv", 0, [probability, degFreedom], false, true, null));
		};
		Functions.prototype.t_Inv_2T=function (probability, degFreedom) {
			return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "T_Inv_2T", 0, [probability, degFreedom], false, true, null));
		};
		Functions.prototype.tan=function (number) {
			return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Tan", 0, [number], false, true, null));
		};
		Functions.prototype.tanh=function (number) {
			return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Tanh", 0, [number], false, true, null));
		};
		Functions.prototype.text=function (value, formatText) {
			return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Text", 0, [value, formatText], false, true, null));
		};
		Functions.prototype.time=function (hour, minute, second) {
			return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Time", 0, [hour, minute, second], false, true, null));
		};
		Functions.prototype.timevalue=function (timeText) {
			return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Timevalue", 0, [timeText], false, true, null));
		};
		Functions.prototype.today=function () {
			return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Today", 0, [], false, true, null));
		};
		Functions.prototype.trim=function (text) {
			return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Trim", 0, [text], false, true, null));
		};
		Functions.prototype.trimMean=function (array, percent) {
			return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "TrimMean", 0, [array, percent], false, true, null));
		};
		Functions.prototype.true=function () {
			return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "True", 0, [], false, true, null));
		};
		Functions.prototype.trunc=function (number, numDigits) {
			return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Trunc", 0, [number, numDigits], false, true, null));
		};
		Functions.prototype.type=function (value) {
			return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Type", 0, [value], false, true, null));
		};
		Functions.prototype.usdollar=function (number, decimals) {
			return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "USDollar", 0, [number, decimals], false, true, null));
		};
		Functions.prototype.unichar=function (number) {
			return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Unichar", 0, [number], false, true, null));
		};
		Functions.prototype.unicode=function (text) {
			return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Unicode", 0, [text], false, true, null));
		};
		Functions.prototype.upper=function (text) {
			return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Upper", 0, [text], false, true, null));
		};
		Functions.prototype.vlookup=function (lookupValue, tableArray, colIndexNum, rangeLookup) {
			return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "VLookup", 0, [lookupValue, tableArray, colIndexNum, rangeLookup], false, true, null));
		};
		Functions.prototype.value=function (text) {
			return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Value", 0, [text], false, true, null));
		};
		Functions.prototype.varA=function () {
			var values=[];
			for (var _i=0; _i < arguments.length; _i++) {
				values[_i]=arguments[_i];
			}
			return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "VarA", 0, [values], false, true, null));
		};
		Functions.prototype.varPA=function () {
			var values=[];
			for (var _i=0; _i < arguments.length; _i++) {
				values[_i]=arguments[_i];
			}
			return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "VarPA", 0, [values], false, true, null));
		};
		Functions.prototype.var_P=function () {
			var values=[];
			for (var _i=0; _i < arguments.length; _i++) {
				values[_i]=arguments[_i];
			}
			return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Var_P", 0, [values], false, true, null));
		};
		Functions.prototype.var_S=function () {
			var values=[];
			for (var _i=0; _i < arguments.length; _i++) {
				values[_i]=arguments[_i];
			}
			return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Var_S", 0, [values], false, true, null));
		};
		Functions.prototype.vdb=function (cost, salvage, life, startPeriod, endPeriod, factor, noSwitch) {
			return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Vdb", 0, [cost, salvage, life, startPeriod, endPeriod, factor, noSwitch], false, true, null));
		};
		Functions.prototype.weekNum=function (serialNumber, returnType) {
			return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "WeekNum", 0, [serialNumber, returnType], false, true, null));
		};
		Functions.prototype.weekday=function (serialNumber, returnType) {
			return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Weekday", 0, [serialNumber, returnType], false, true, null));
		};
		Functions.prototype.weibull_Dist=function (x, alpha, beta, cumulative) {
			return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Weibull_Dist", 0, [x, alpha, beta, cumulative], false, true, null));
		};
		Functions.prototype.workDay=function (startDate, days, holidays) {
			return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "WorkDay", 0, [startDate, days, holidays], false, true, null));
		};
		Functions.prototype.workDay_Intl=function (startDate, days, weekend, holidays) {
			return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "WorkDay_Intl", 0, [startDate, days, weekend, holidays], false, true, null));
		};
		Functions.prototype.xirr=function (values, dates, guess) {
			return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Xirr", 0, [values, dates, guess], false, true, null));
		};
		Functions.prototype.xnpv=function (rate, values, dates) {
			return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Xnpv", 0, [rate, values, dates], false, true, null));
		};
		Functions.prototype.xor=function () {
			var values=[];
			for (var _i=0; _i < arguments.length; _i++) {
				values[_i]=arguments[_i];
			}
			return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Xor", 0, [values], false, true, null));
		};
		Functions.prototype.year=function (serialNumber) {
			return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Year", 0, [serialNumber], false, true, null));
		};
		Functions.prototype.yearFrac=function (startDate, endDate, basis) {
			return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "YearFrac", 0, [startDate, endDate, basis], false, true, null));
		};
		Functions.prototype.yield=function (settlement, maturity, rate, pr, redemption, frequency, basis) {
			return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Yield", 0, [settlement, maturity, rate, pr, redemption, frequency, basis], false, true, null));
		};
		Functions.prototype.yieldDisc=function (settlement, maturity, pr, redemption, basis) {
			return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "YieldDisc", 0, [settlement, maturity, pr, redemption, basis], false, true, null));
		};
		Functions.prototype.yieldMat=function (settlement, maturity, issue, rate, pr, basis) {
			return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "YieldMat", 0, [settlement, maturity, issue, rate, pr, basis], false, true, null));
		};
		Functions.prototype.z_Test=function (array, x, sigma) {
			return new FunctionResult(this.context, _createMethodObjectPath(this.context, this, "Z_Test", 0, [array, x, sigma], false, true, null));
		};
		Functions.prototype._handleResult=function (value) {
			_super.prototype._handleResult.call(this, value);
			if (_isNullOrUndefined(value))
				return;
			var obj=value;
			_fixObjectPathIfNecessary(this, obj);
		};
		Functions.prototype.toJSON=function () {
			return {};
		};
		return Functions;
	}(OfficeExtension.ClientObject));
	Excel.Functions=Functions;
	var ErrorCodes;
	(function (ErrorCodes) {
		ErrorCodes.accessDenied="AccessDenied";
		ErrorCodes.apiNotFound="ApiNotFound";
		ErrorCodes.generalException="GeneralException";
		ErrorCodes.insertDeleteConflict="InsertDeleteConflict";
		ErrorCodes.invalidArgument="InvalidArgument";
		ErrorCodes.invalidBinding="InvalidBinding";
		ErrorCodes.invalidOperation="InvalidOperation";
		ErrorCodes.invalidReference="InvalidReference";
		ErrorCodes.invalidSelection="InvalidSelection";
		ErrorCodes.itemAlreadyExists="ItemAlreadyExists";
		ErrorCodes.itemNotFound="ItemNotFound";
		ErrorCodes.notImplemented="NotImplemented";
		ErrorCodes.unsupportedOperation="UnsupportedOperation";
		ErrorCodes.invalidOperationInCellEditMode="InvalidOperationInCellEditMode";
	})(ErrorCodes=Excel.ErrorCodes || (Excel.ErrorCodes={}));
})(Excel || (Excel={}));

OfficeExtension.Utility._doApiNotSupportedCheck=true;



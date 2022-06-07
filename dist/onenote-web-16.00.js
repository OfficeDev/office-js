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


// Sources:
// osfweb: 16.0\15202.10000
// runtime: 16.0\15202.10000
// core: 16.0\15202.10000
// host: 16.0\15202.10000



var OfficeExt,__extends=this&&this.__extends||function(){var e=function(t,n){return(e=Object.setPrototypeOf||{__proto__:[]}instanceof Array&&function(e,t){e.__proto__=t}||function(e,t){for(var n in t)t.hasOwnProperty(n)&&(e[n]=t[n])})(t,n)};return function(t,n){function o(){this.constructor=t}e(t,n),t.prototype=null===n?Object.create(n):(o.prototype=n.prototype,new o)}}();!function(e){var t=function(){function e(){}return e.prototype.isMsAjaxLoaded=function(){return!!("undefined"!==typeof Sys&&"undefined"!==typeof Type&&Sys.StringBuilder&&"function"===typeof Sys.StringBuilder&&Type.registerNamespace&&"function"===typeof Type.registerNamespace&&Type.registerClass&&"function"===typeof Type.registerClass&&"function"===typeof Function._validateParams&&Sys.Serialization&&Sys.Serialization.JavaScriptSerializer&&"function"===typeof Sys.Serialization.JavaScriptSerializer.serialize)},e.prototype.loadMsAjaxFull=function(e){var t=("https:"===window.location.protocol.toLowerCase()?"https:":"http:")+"//ajax.aspnetcdn.com/ajax/3.5/MicrosoftAjax.js";OSF.OUtil.loadScript(t,e)},Object.defineProperty(e.prototype,"msAjaxError",{get:function(){return null==this._msAjaxError&&this.isMsAjaxLoaded()&&(this._msAjaxError=Error),this._msAjaxError},set:function(e){this._msAjaxError=e},enumerable:!0,configurable:!0}),Object.defineProperty(e.prototype,"msAjaxString",{get:function(){return null==this._msAjaxString&&this.isMsAjaxLoaded()&&(this._msAjaxString=String),this._msAjaxString},set:function(e){this._msAjaxString=e},enumerable:!0,configurable:!0}),Object.defineProperty(e.prototype,"msAjaxDebug",{get:function(){return null==this._msAjaxDebug&&this.isMsAjaxLoaded()&&(this._msAjaxDebug=Sys.Debug),this._msAjaxDebug},set:function(e){this._msAjaxDebug=e},enumerable:!0,configurable:!0}),e}();e.MicrosoftAjaxFactory=t}(OfficeExt||(OfficeExt={}));var OSFLog,Logger,OSFAriaLogger,OSFAppTelemetry,OSFPerfUtil,OSFWebAuth,OSF_DDA_Marshaling_FilePropertiesKeys,OSF_DDA_Marshaling_File_FilePropertiesKeys,OSF_DDA_Marshaling_File_SlicePropertiesKeys,OSF_DDA_Marshaling_File_FileType,OSF_DDA_Marshaling_File_ParameterKeys,OSF_DDA_Marshaling_AppCommand_AppCommandInvokedEventKeys,OSF_DDA_Marshaling_AppCommand_AppCommandCompletedMethodParameterKeys,OsfMsAjaxFactory=new OfficeExt.MicrosoftAjaxFactory,OSF=OSF||{};!function(e){var t=function(){function e(e){this._internalStorage=e}return e.prototype.getItem=function(e){try{return this._internalStorage&&this._internalStorage.getItem(e)}catch(e){return null}},e.prototype.setItem=function(e,t){try{this._internalStorage&&this._internalStorage.setItem(e,t)}catch(e){}},e.prototype.clear=function(){try{this._internalStorage&&this._internalStorage.clear()}catch(e){}},e.prototype.removeItem=function(e){try{this._internalStorage&&this._internalStorage.removeItem(e)}catch(e){}},e.prototype.getKeysWithPrefix=function(e){var t=[];try{for(var n=this._internalStorage&&this._internalStorage.length||0,o=0;o<n;o++){var r=this._internalStorage.key(o);0===r.indexOf(e)&&t.push(r)}}catch(e){}return t},e.prototype.isLocalStorageAvailable=function(){return null!=this._internalStorage},e}();e.SafeStorage=t}(OfficeExt||(OfficeExt={})),OSF.XdmFieldName={ConversationUrl:"ConversationUrl",AppId:"AppId"},OSF.TestFlightStart=1e3,OSF.TestFlightEnd=1009,OSF.FlightNames={UseOriginNotUrl:0,AddinEnforceHttps:2,FirstPartyAnonymousProxyReadyCheckTimeout:6,AddinRibbonIdAllowUnknown:9,ManifestParserDevConsoleLog:15,AddinActionDefinitionHybridMode:18,UseActionIdForUILessCommand:20,RequirementSetRibbonApiOnePointTwo:21,SetFocusToTaskpaneIsEnabled:22,ShortcutInfoArrayInUserPreferenceData:23,OSFTestFlight1000:OSF.TestFlightStart,OSFTestFlight1001:OSF.TestFlightStart+1,OSFTestFlight1002:OSF.TestFlightStart+2,OSFTestFlight1003:OSF.TestFlightStart+3,OSFTestFlight1004:OSF.TestFlightStart+4,OSFTestFlight1005:OSF.TestFlightStart+5,OSFTestFlight1006:OSF.TestFlightStart+6,OSFTestFlight1007:OSF.TestFlightStart+7,OSFTestFlight1008:OSF.TestFlightStart+8,OSFTestFlight1009:OSF.TestFlightEnd},OSF.TrustUXFlightValues={TrustUXControlA:0,TrustUXExperimentB:1,TrustUXExperimentC:2},OSF.FlightTreatmentNames={AllowStorageAccessByUserActivationOnIFrameCheck:"Microsoft.Office.SharedOnline.AllowStorageAccessByUserActivationOnIFrameCheck",WopiPreinstalledAddInsEnabled:"Microsoft.Office.SharedOnline.WopiPreinstalledAddInsEnabled",WopiUseNewActivate:"Microsoft.Office.SharedOnline.WopiUseNewActivate",CheckProxyIsReadyRetry:"Microsoft.Office.SharedOnline.OEP.CheckProxyIsReadyRetry",InsertionDialogFixesEnabled:"Microsoft.Office.SharedOnline.InsertionDialogFixesEnabled",BlockAutoOpenAddInIfStoreDisabled:"Microsoft.Office.SharedOnline.BlockAutoOpenAddInIfStoreDisabled",AddinTrustUXImprovement:"Microsoft.Office.SharedOnline.AddinTrustUXImprovement",TeachingUIForPrivateCatelogEnabled:"Microsoft.Office.SharedOnline.TeachingUIForPrivateCatelogEnabled"},OSF.Flights=[],OSF.IntFlights={},OSF.Settings={},OSF.WindowNameItemKeys={BaseFrameName:"baseFrameName",HostInfo:"hostInfo",XdmInfo:"xdmInfo",SerializerVersion:"serializerVersion",AppContext:"appContext",Flights:"flights"},OSF.OUtil=function(){var e=-1,t="#",n={},o=null,r=null,i=(new Date).getTime();function a(){var e=2147483647*Math.random();return(e^=i^(new Date).getMilliseconds()<<Math.floor(21*Math.random())).toString(16)}function s(){if(!o){try{var e=window.sessionStorage}catch(t){e=null}o=new OfficeExt.SafeStorage(e)}return o}function c(e){var t,n,o=[],r=[],i=e.length;for(t=0;t<i;t++)(n=e[t]).tabIndex?n.tabIndex>0?r.push(n):0===n.tabIndex&&o.push(n):o.push(n);return r=r.sort((function(e,t){var n=e.tabIndex-t.tabIndex;return 0===n&&(n=r.indexOf(e)-r.indexOf(t)),n})),[].concat(r,o)}return{set_entropy:function(e){if("string"==typeof e)for(var t=0;t<e.length;t+=4){for(var n=0,o=0;o<4&&t+o<e.length;o++)n=(n<<8)+e.charCodeAt(t+o);i^=n}else i^="number"==typeof e?e:2147483647*Math.random();i&=2147483647},extend:function(e,t){var n=function(){};n.prototype=t.prototype,e.prototype=new n,e.prototype.constructor=e,e.uber=t.prototype,t.prototype.constructor===Object.prototype.constructor&&(t.prototype.constructor=t)},setNamespace:function(e,t){t&&e&&!t[e]&&(t[e]={})},unsetNamespace:function(e,t){t&&e&&t[e]&&delete t[e]},serializeSettings:function(e){var t={};for(var n in e){var o=e[n];try{o=JSON?JSON.stringify(o,(function(e,t){return OSF.OUtil.isDate(this[e])?OSF.DDA.SettingsManager.DateJSONPrefix+this[e].getTime()+OSF.DDA.SettingsManager.DataJSONSuffix:t})):Sys.Serialization.JavaScriptSerializer.serialize(o),t[n]=o}catch(e){}}return t},deserializeSettings:function(e){var t={};for(var n in e=e||{}){var o=e[n];try{o=JSON?JSON.parse(o,(function(e,t){var n;return"string"===typeof t&&t&&t.length>6&&t.slice(0,5)===OSF.DDA.SettingsManager.DateJSONPrefix&&t.slice(-1)===OSF.DDA.SettingsManager.DataJSONSuffix&&(n=new Date(parseInt(t.slice(5,-1))))?n:t})):Sys.Serialization.JavaScriptSerializer.deserialize(o,!0),t[n]=o}catch(e){}}return t},loadScript:function(e,t,o){if(e&&t){var r=window.document,i=n[e];if(i)i.loaded?t():i.pendingCallbacks.push(t);else{var a=r.createElement("script");a.type="text/javascript",i={loaded:!1,pendingCallbacks:[t],timer:null},n[e]=i;var s=function(){null!=i.timer&&(clearTimeout(i.timer),delete i.timer),i.loaded=!0;for(var e=i.pendingCallbacks.length,t=0;t<e;t++){i.pendingCallbacks.shift()()}},c=function(t){delete n[e],null!=i.timer&&(clearTimeout(i.timer),delete i.timer);for(var o=i.pendingCallbacks.length,r=0;r<o;r++){i.pendingCallbacks.shift()()}};a.readyState?a.onreadystatechange=function(){"loaded"!=a.readyState&&"complete"!=a.readyState||(a.onreadystatechange=null,s())}:a.onload=s,a.onerror=c,o=o||3e4,i.timer=setTimeout((function(){window.navigator.userAgent.indexOf("Trident")>0?c(null):c(new Event("Script load timed out"))}),o),a.setAttribute("crossOrigin","anonymous"),a.src=e,r.getElementsByTagName("head")[0].appendChild(a)}}},loadCSS:function(e){if(e){var t=window.document,n=t.createElement("link");n.type="text/css",n.rel="stylesheet",n.href=e,t.getElementsByTagName("head")[0].appendChild(n)}},parseEnum:function(e,t){var n=t[e.trim()];if("undefined"==typeof n)throw OsfMsAjaxFactory.msAjaxDebug.trace("invalid enumeration string:"+e),OsfMsAjaxFactory.msAjaxError.argument("str");return n},delayExecutionAndCache:function(){var e={calc:arguments[0]};return function(){return e.calc&&(e.val=e.calc.apply(this,arguments),delete e.calc),e.val}},getUniqueId:function(){return(e+=1).toString()},formatString:function(){var e=arguments,t=e[0];return t.replace(/{(\d+)}/gm,(function(t,n){var o=parseInt(n,10)+1;return void 0===e[o]?"{"+n+"}":e[o]}))},generateConversationId:function(){return[a(),a(),(new Date).getTime().toString()].join("_")},getFrameName:function(e){return"_xdm_"+e+this.generateConversationId()},addXdmInfoAsHash:function(e,t){return OSF.OUtil.addInfoAsHash(e,"&_xdm_Info=",t,!1)},addSerializerVersionAsHash:function(e,t){return OSF.OUtil.addInfoAsHash(e,"&_serializer_version=",t,!0)},addFlightsAsHash:function(e,t){return OSF.OUtil.addInfoAsHash(e,"&_flights=",t,!0)},addInfoAsHash:function(e,n,o,r){var i,a=(e=e.trim()||"").split(t),s=a.shift(),c=a.join(t);return i=r?[n,encodeURIComponent(o),c].join(""):[c,n,o].join(""),[s,t,i].join("")},parseHostInfoFromWindowName:function(e,t){return OSF.OUtil.parseInfoFromWindowName(e,t,OSF.WindowNameItemKeys.HostInfo)},parseXdmInfo:function(e){var t=OSF.OUtil.parseXdmInfoWithGivenFragment(e,window.location.hash);return t||(t=OSF.OUtil.parseXdmInfoFromWindowName(e,window.name)),t},parseXdmInfoFromWindowName:function(e,t){return OSF.OUtil.parseInfoFromWindowName(e,t,OSF.WindowNameItemKeys.XdmInfo)},parseXdmInfoWithGivenFragment:function(e,t){return OSF.OUtil.parseInfoWithGivenFragment("&_xdm_Info=","_xdm_",!1,e,t)},parseSerializerVersion:function(e){var t=OSF.OUtil.parseSerializerVersionWithGivenFragment(e,window.location.hash);return isNaN(t)&&(t=OSF.OUtil.parseSerializerVersionFromWindowName(e,window.name)),t},parseSerializerVersionFromWindowName:function(e,t){return parseInt(OSF.OUtil.parseInfoFromWindowName(e,t,OSF.WindowNameItemKeys.SerializerVersion))},parseSerializerVersionWithGivenFragment:function(e,t){return parseInt(OSF.OUtil.parseInfoWithGivenFragment("&_serializer_version=","_serializer_version=",!0,e,t))},parseFlights:function(e){var t=OSF.OUtil.parseFlightsWithGivenFragment(e,window.location.hash);return 0==t.length&&(t=OSF.OUtil.parseFlightsFromWindowName(e,window.name)),t},checkFlight:function(e){return OSF.Flights&&OSF.Flights.indexOf(e)>=0},pushFlight:function(e){return OSF.Flights.indexOf(e)<0&&(OSF.Flights.push(e),!0)},getBooleanSetting:function(e){return OSF.OUtil.getBooleanFromDictionary(OSF.Settings,e)},getBooleanFromDictionary:function(e,t){var n=e&&t&&void 0!==e[t]&&e[t]&&("string"===typeof e[t]&&"TRUE"===e[t].toUpperCase()||"boolean"===typeof e[t]&&e[t]);return void 0!==n&&n},getIntFromDictionary:function(e,t){return e&&t&&void 0!==e[t]&&"string"===typeof e[t]?parseInt(e[t]):NaN},pushIntFlight:function(e,t){return!(e in OSF.IntFlights)&&(OSF.IntFlights[e]=t,!0)},getIntFlight:function(e){return OSF.IntFlights&&e in OSF.IntFlights?OSF.IntFlights[e]:NaN},parseFlightsFromWindowName:function(e,t){return OSF.OUtil.parseArrayWithDefault(OSF.OUtil.parseInfoFromWindowName(e,t,OSF.WindowNameItemKeys.Flights))},parseFlightsWithGivenFragment:function(e,t){return OSF.OUtil.parseArrayWithDefault(OSF.OUtil.parseInfoWithGivenFragment("&_flights=","_flights=",!0,e,t))},parseArrayWithDefault:function(e){var t=[];try{t=JSON.parse(e)}catch(e){}return Array.isArray(t)||(t=[]),t},parseInfoFromWindowName:function(e,t,n){try{var o=JSON.parse(t),r=null!=o?o[n]:null,i=s();if(!e&&i&&null!=o){var a=o[OSF.WindowNameItemKeys.BaseFrameName]+n;r?i.setItem(a,r):r=i.getItem(a)}return r}catch(e){return null}},parseInfoWithGivenFragment:function(e,t,n,o,r){var i=r.split(e),a=i.length>1?i[i.length-1]:null;n&&null!=a&&(a.indexOf("&")>=0&&(a=a.split("&")[0]),a=decodeURIComponent(a));var c=s();if(!o&&c){var l=window.name.indexOf(t);if(l>-1){var u=window.name.indexOf(";",l);-1==u&&(u=window.name.length);var p=window.name.substring(l,u);a?c.setItem(p,a):a=c.getItem(p)}}return a},getConversationId:function(){var e=window.location.search,t=null;if(e){var n=e.indexOf("&");(t=n>0?e.substring(1,n):e.substr(1))&&"="===t.charAt(t.length-1)&&(t=t.substring(0,t.length-1))&&(t=decodeURIComponent(t))}return t},getInfoItems:function(e){var t=e.split("$");return"undefined"==typeof t[1]&&(t=e.split("|")),"undefined"==typeof t[1]&&(t=e.split("%7C")),t},getXdmFieldValue:function(e,t){var n="",o=OSF.OUtil.parseXdmInfo(t);if(o){var r=OSF.OUtil.getInfoItems(o);if(void 0!=r&&r.length>=3)switch(e){case OSF.XdmFieldName.ConversationUrl:n=r[2];break;case OSF.XdmFieldName.AppId:n=r[1]}}return n},validateParamObject:function(e,t,n){var o=Function._validateParams(arguments,[{name:"params",type:Object,mayBeNull:!1},{name:"expectedProperties",type:Object,mayBeNull:!1},{name:"callback",type:Function,mayBeNull:!0}]);if(o)throw o;for(var r in t)if(o=Function._validateParameter(e[r],t[r],r))throw o},writeProfilerMark:function(e){window.msWriteProfilerMark&&(window.msWriteProfilerMark(e),OsfMsAjaxFactory.msAjaxDebug.trace(e))},outputDebug:function(e){"undefined"!==typeof OsfMsAjaxFactory&&OsfMsAjaxFactory.msAjaxDebug&&OsfMsAjaxFactory.msAjaxDebug.trace&&OsfMsAjaxFactory.msAjaxDebug.trace(e)},defineNondefaultProperty:function(e,t,n,o){for(var r in n=n||{},o){var i=o[r];void 0==n[i]&&(n[i]=!0)}return Object.defineProperty(e,t,n),e},defineNondefaultProperties:function(e,t,n){for(var o in t=t||{})OSF.OUtil.defineNondefaultProperty(e,o,t[o],n);return e},defineEnumerableProperty:function(e,t,n){return OSF.OUtil.defineNondefaultProperty(e,t,n,["enumerable"])},defineEnumerableProperties:function(e,t){return OSF.OUtil.defineNondefaultProperties(e,t,["enumerable"])},defineMutableProperty:function(e,t,n){return OSF.OUtil.defineNondefaultProperty(e,t,n,["writable","enumerable","configurable"])},defineMutableProperties:function(e,t){return OSF.OUtil.defineNondefaultProperties(e,t,["writable","enumerable","configurable"])},finalizeProperties:function(e,t){t=t||{};for(var n=Object.getOwnPropertyNames(e),o=n.length,r=0;r<o;r++){var i=n[r],a=Object.getOwnPropertyDescriptor(e,i);a.get||a.set||(a.writable=t.writable||!1),a.configurable=t.configurable||!1,a.enumerable=t.enumerable||!0,Object.defineProperty(e,i,a)}return e},mapList:function(e,t){var n=[];if(e)for(var o in e)n.push(t(e[o]));return n},listContainsKey:function(e,t){for(var n in e)if(t==n)return!0;return!1},listContainsValue:function(e,t){for(var n in e)if(t==e[n])return!0;return!1},augmentList:function(e,t){var n=e.push?function(t,n){e.push(n)}:function(t,n){e[t]=n};for(var o in t)n(o,t[o])},redefineList:function(e,t){for(var n in e)delete e[n];for(var o in t)e[o]=t[o]},isArray:function(e){return"[object Array]"===Object.prototype.toString.apply(e)},isFunction:function(e){return"[object Function]"===Object.prototype.toString.apply(e)},isDate:function(e){return"[object Date]"===Object.prototype.toString.apply(e)},addEventListener:function(e,t,n){e.addEventListener?e.addEventListener(t,n,!1):Sys.Browser.agent===Sys.Browser.InternetExplorer&&e.attachEvent?e.attachEvent("on"+t,n):e["on"+t]=n},removeEventListener:function(e,t,n){e.removeEventListener?e.removeEventListener(t,n,!1):Sys.Browser.agent===Sys.Browser.InternetExplorer&&e.detachEve
/* Outlook Mac specific API library */
/* osfweb version: 16.0.15303.10000 */
/* office-js-api version: 20220505.3 */
/*
	Copyright (c) Microsoft Corporation.  All rights reserved.
*/
/*
    Your use of this file is governed by the Microsoft Services Agreement http://go.microsoft.com/fwlink/?LinkId=266419.

    This file also contains the following Promise implementation (with a few small modifications):
        * @overview es6-promise - a tiny implementation of Promises/A+.
        * @copyright Copyright (c) 2014 Yehuda Katz, Tom Dale, Stefan Penner and contributors (Conversion to ES6 API by Jake Archibald)
        * @license   Licensed under MIT license
        *            See https://raw.githubusercontent.com/jakearchibald/es6-promise/master/LICENSE
        * @version   2.3.0
*/
"undefined"!==typeof OSFPerformance&&(OSFPerformance.hostInitializationStart=OSFPerformance.now());
/* Outlook Mac client specific API library */
/* Version: 16.0.15303.10000 */
var __extends=this&&this.__extends||function(){var a=function(c,b){a=Object.setPrototypeOf||{__proto__:[]}instanceof Array&&function(b,a){b.__proto__=a}||function(c,a){for(var b in a)if(a.hasOwnProperty(b))c[b]=a[b]};return a(c,b)};return function(c,b){a(c,b);function d(){this.constructor=c}c.prototype=b===null?Object.create(b):(d.prototype=b.prototype,new d)}}(),OfficeExt;(function(b){var a=function(){var a=true;function b(){}b.prototype.isMsAjaxLoaded=function(){var b="function",c="undefined";if(typeof Sys!==c&&typeof Type!==c&&Sys.StringBuilder&&typeof Sys.StringBuilder===b&&Type.registerNamespace&&typeof Type.registerNamespace===b&&Type.registerClass&&typeof Type.registerClass===b&&typeof Function._validateParams===b&&Sys.Serialization&&Sys.Serialization.JavaScriptSerializer&&typeof Sys.Serialization.JavaScriptSerializer.serialize===b)return a;else return false};b.prototype.loadMsAjaxFull=function(b){var a=(window.location.protocol.toLowerCase()==="https:"?"https:":"http:")+"//ajax.aspnetcdn.com/ajax/3.5/MicrosoftAjax.js";OSF.OUtil.loadScript(a,b)};Object.defineProperty(b.prototype,"msAjaxError",{"get":function(){var a=this;if(a._msAjaxError==null&&a.isMsAjaxLoaded())a._msAjaxError=Error;return a._msAjaxError},"set":function(a){this._msAjaxError=a},enumerable:a,configurable:a});Object.defineProperty(b.prototype,"msAjaxString",{"get":function(){var a=this;if(a._msAjaxString==null&&a.isMsAjaxLoaded())a._msAjaxString=String;return a._msAjaxString},"set":function(a){this._msAjaxString=a},enumerable:a,configurable:a});Object.defineProperty(b.prototype,"msAjaxDebug",{"get":function(){var a=this;if(a._msAjaxDebug==null&&a.isMsAjaxLoaded())a._msAjaxDebug=Sys.Debug;return a._msAjaxDebug},"set":function(a){this._msAjaxDebug=a},enumerable:a,configurable:a});return b}();b.MicrosoftAjaxFactory=a})(OfficeExt||(OfficeExt={}));var OsfMsAjaxFactory=new OfficeExt.MicrosoftAjaxFactory,OSF=OSF||{};(function(b){var a=function(){function a(a){this._internalStorage=a}a.prototype.getItem=function(a){try{return this._internalStorage&&this._internalStorage.getItem(a)}catch(b){return null}};a.prototype.setItem=function(b,a){try{this._internalStorage&&this._internalStorage.setItem(b,a)}catch(c){}};a.prototype.clear=function(){try{this._internalStorage&&this._internalStorage.clear()}catch(a){}};a.prototype.removeItem=function(a){try{this._internalStorage&&this._internalStorage.removeItem(a)}catch(b){}};a.prototype.getKeysWithPrefix=function(d){var b=[];try{for(var e=this._internalStorage&&this._internalStorage.length||0,a=0;a<e;a++){var c=this._internalStorage.key(a);c.indexOf(d)===0&&b.push(c)}}catch(f){}return b};a.prototype.isLocalStorageAvailable=function(){return this._internalStorage!=null};return a}();b.SafeStorage=a})(OfficeExt||(OfficeExt={}));OSF.XdmFieldName={ConversationUrl:"ConversationUrl",AppId:"AppId"};OSF.TestFlightStart=1e3;OSF.TestFlightEnd=1009;OSF.FlightNames={UseOriginNotUrl:0,AddinEnforceHttps:2,FirstPartyAnonymousProxyReadyCheckTimeout:6,AddinRibbonIdAllowUnknown:9,ManifestParserDevConsoleLog:15,AddinActionDefinitionHybridMode:18,UseActionIdForUILessCommand:20,RequirementSetRibbonApiOnePointTwo:21,SetFocusToTaskpaneIsEnabled:22,ShortcutInfoArrayInUserPreferenceData:23,OSFTestFlight1000:OSF.TestFlightStart,OSFTestFlight1001:OSF.TestFlightStart+1,OSFTestFlight1002:OSF.TestFlightStart+2,OSFTestFlight1003:OSF.TestFlightStart+3,OSFTestFlight1004:OSF.TestFlightStart+4,OSFTestFlight1005:OSF.TestFlightStart+5,OSFTestFlight1006:OSF.TestFlightStart+6,OSFTestFlight1007:OSF.TestFlightStart+7,OSFTestFlight1008:OSF.TestFlightStart+8,OSFTestFlight1009:OSF.TestFlightEnd};OSF.TrustUXFlightValues={TrustUXControlA:0,TrustUXExperimentB:1,TrustUXExperimentC:2};OSF.FlightTreatmentNames={AddinDialogIFrameContentWindowKillSwitch:"Microsoft.Office.SharedOnline.AddinDialogIFrameContentWindowKillSwitch",AddinTrustUXImprovement:"Microsoft.Office.SharedOnline.AddinTrustUXImprovement",AllowStorageAccessByUserActivationOnIFrameCheck:"Microsoft.Office.SharedOnline.AllowStorageAccessByUserActivationOnIFrameCheck",BlockAutoOpenAddInIfStoreDisabled:"Microsoft.Office.SharedOnline.BlockAutoOpenAddInIfStoreDisabled",CheckProxyIsReadyRetry:"Microsoft.Office.SharedOnline.OEP.CheckProxyIsReadyRetry",GetOmexPrefetchDeprecation:"Microsoft.Office.SharedOnline.GetOmexPrefetchDeprecation",InsertionDialogFixesEnabled:"Microsoft.Office.SharedOnline.InsertionDialogFixesEnabled",TeachingUIForPrivateCatelogEnabled:"Microsoft.Office.SharedOnline.TeachingUIForPrivateCatelogEnabled",WopiPreinstalledAddInsEnabled:"Microsoft.Office.SharedOnline.WopiPreinstalledAddInsEnabled",WopiUseNewActivate:"Microsoft.Office.SharedOnline.WopiUseNewActivate",MosManifestEnabled:"Microsoft.Office.SharedOnline.OEP.MosManifest"};OSF.Flights=[];OSF.IntFlights={};OSF.Settings={};OSF.WindowNameItemKeys={BaseFrameName:"baseFrameName",HostInfo:"hostInfo",XdmInfo:"xdmInfo",SerializerVersion:"serializerVersion",AppContext:"appContext",Flights:"flights"};OSF.OUtil=function(){var l="focus",k="https:",j="on",q="configurable",p="writable",i="enumerable",e="",f="undefined",c=false,b=true,h="string",m=2147483647,a=null,g="#",d=-1,w=d,C="&_xdm_Info=",z="&_serializer_version=",B="&_flights=",A="_xdm_",F="_serializer_version=",G="_flights=",s=g,y="&",n="class",v={},E=3e4,r=a,u=a,o=(new Date).getTime();function D(){var a=m*Math.random();a^=o^(new Date).getMilliseconds()<<Math.floor(Math.random()*(31-10));return a.toString(16)}function t(){if(!r){try{var b=window.sessionStorage}catch(c){b=a}r=new OfficeExt.SafeStorage(b)}return r}function x(e){for(var c=[],b=[],f=e.length,a,d=0;d<f;d++){a=e[d];if(a.tabIndex)if(a.tabIndex>0)b.push(a);else a.tabIndex===0&&c.push(a);else c.push(a)}b=b.sort(function(d,c){var a=d.tabIndex-c.tabIndex;if(a===0)a=b.indexOf(d)-b.indexOf(c);return a});return [].concat(b,c)}return {set_entropy:function(a){if(typeof a==h)for(var b=0;b<a.length;b+=4){for(var d=0,c=0;c<4&&b+c<a.length;c++)d=(d<<8)+a.charCodeAt(b+c);o^=d}else if(typeof a=="number")o^=a;else o^=m*Math.random();o&=m},extend:function(b,a){var c=function(){};c.prototype=a.prototype;b.prototype=new c;b.prototype.constructor=b;b.uber=a.prototype;if(a.prototype.constructor===Object.prototype.constructor)a.prototype.constructor=a},setNamespace:function(b,a){if(a&&b&&!a[b])a[b]={}},unsetNamespace:function(b,a){if(a&&b&&a[b])delete a[b]},serializeSettings:function(b){var d={};for(var c in b){var a=b[c];try{if(JSON)a=JSON.stringify(a,function(a,b){return OSF.OUtil.isDate(this[a])?OSF.DDA.SettingsManager.DateJSONPrefix+this[a].getTime()+OSF.DDA.SettingsManager.DataJSONSuffix:b});else a=Sys.Serialization.JavaScriptSerializer.serialize(a);d[c]=a}catch(e){}}return d},deserializeSettings:function(c){var f={};c=c||{};for(var e in c){var a=c[e];try{if(JSON)a=JSON.parse(a,function(c,a){var b;if(typeof a===h&&a&&a.length>6&&a.slice(0,5)===OSF.DDA.SettingsManager.DateJSONPrefix&&a.slice(d)===OSF.DDA.SettingsManager.DataJSONSuffix){b=new Date(parseInt(a.slice(5,d)));if(b)return b}return a});else a=Sys.Serialization.JavaScriptSerializer.deserialize(a,b);f[e]=a}catch(g){}}return f},loadScript:function(f,g,i){if(f&&g){var k=window.document,d=v[f];if(!d){var e=k.createElement("script");e.type="text/javascript";d={loaded:c,pendingCallbacks:[g],timer:a};v[f]=d;var j=function(){if(d.timer!=a){clearTimeout(d.timer);delete d.timer}d.loaded=b;for(var e=d.pendingCallbacks.length,c=0;c<e;c++){var f=d.pendingCallbacks.shift();f()}},l=function(){if(window.navigator.userAgent.indexOf("Trident")>0)h(a);else h(new Event("Script load timed out"))},h=function(){delete v[f];if(d.timer!=a){clearTimeout(d.timer);delete d.timer}for(var c=d.pendingCallbacks.length,b=0;b<c;b++){var e=d.pendingCallbacks.shift();e()}};if(e.readyState)e.onreadystatechange=function(){if(e.readyState=="loaded"||e.readyState=="complete"){e.onreadystatechange=a;j()}};else e.onload=j;e.onerror=h;i=i||E;d.timer=setTimeout(l,i);e.setAttribute("crossOrigin","anonymous");e.src=f;k.getElementsByTagName("head")[0].appendChild(e)}else if(d.loaded)g();else d.pendingCallbacks.push(g)}},loadCSS:function(c){if(c){var b=window.document,a=b.createElement("link");a.type="text/css";a.rel="stylesheet";a.href=c;b.getElementsByTagName("head")[0].appendChild(a)}},parseEnum:function(b,c){var a=c[b.trim()];if(typeof a==f){OsfMsAjaxFactory.msAjaxDebug.trace("invalid enumeration string:"+b);throw OsfMsAjaxFactory.msAjaxError.argument("str")}return a},delayExecutionAndCache:function(){var a={calc:arguments[0]};return function(){if(a.calc){a.val=a.calc.apply(this,arguments);delete a.calc}return a.val}},getUniqueId:function(){w=w+1;return w.toString()},formatString:function(){var a=arguments,b=a[0];return b.replace(/{(\d+)}/gm,function(d,b){var c=parseInt(b,10)+1;return a[c]===undefined?"{"+b+"}":a[c]})},generateConversationId:function(){return [D(),D(),(new Date).getTime().toString()].join("_")},getFrameName:function(a){return A+a+this.generateConversationId()},addXdmInfoAsHash:function(b,a){return OSF.OUtil.addInfoAsHash(b,C,a,c)},addSerializerVersionAsHash:function(c,a){return OSF.OUtil.addInfoAsHash(c,z,a,b)},addFlightsAsHash:function(c,a){return OSF.OUtil.addInfoAsHash(c,B,a,b)},addInfoAsHash:function(b,g,c,i){b=b.trim()||e;var f=b.split(s),h=f.shift(),d=f.join(s),a;if(i)a=[g,encodeURIComponent(c),d].join(e);else a=[d,g,c].join(e);return [h,s,a].join(e)},parseHostInfoFromWindowName:function(a,b){return OSF.OUtil.parseInfoFromWindowName(a,b,OSF.WindowNameItemKeys.HostInfo)},parseXdmInfo:function(b){var a=OSF.OUtil.parseXdmInfoWithGivenFragment(b,window.location.hash);if(!a)a=OSF.OUtil.parseXdmInfoFromWindowName(b,window.name);return a},parseXdmInfoFromWindowName:function(a,b){return OSF.OUtil.parseInfoFromWindowName(a,b,OSF.WindowNameItemKeys.XdmInfo)},parseXdmInfoWithGivenFragment:function(a,b){return OSF.OUtil.parseInfoWithGivenFragment(C,A,c,a,b)},parseSerializerVersion:function(b){var a=OSF.OUtil.parseSerializerVersionWithGivenFragment(b,window.location.hash);if(isNaN(a))a=OSF.OUtil.parseSerializerVersionFromWindowName(b,window.name);return a},parseSerializerVersionFromWindowName:function(a,b){return parseInt(OSF.OUtil.parseInfoFromWindowName(a,b,OSF.WindowNameItemKeys.SerializerVersion))},parseSerializerVersionWithGivenFragment:function(a,c){return parseInt(OSF.OUtil.parseInfoWithGivenFragment(z,F,b,a,c))},parseFlights:function(b){var a=OSF.OUtil.parseFlightsWithGivenFragment(b,window.location.hash);if(a.length==0)a=OSF.OUtil.parseFlightsFromWindowName(b,window.name);return a},checkFlight:function(a){return OSF.Flights&&OSF.Flights.indexOf(a)>=0},pushFlight:function(a){if(OSF.Flights.indexOf(a)<0){OSF.Flights.push(a);return b}return c},getBooleanSetting:function(a){return OSF.OUtil.getBooleanFromDictionary(OSF.Settings,a)},getBooleanFromDictionary:function(b,a){var d=b&&a&&b[a]!==undefined&&b[a]&&(typeof b[a]===h&&b[a].toUpperCase()==="TRUE"||typeof b[a]==="boolean"&&b[a]);return d!==undefined?d:c},getIntFromDictionary:function(b,a){if(b&&a&&b[a]!==undefined&&typeof b[a]===h)return parseInt(b[a]);else return NaN},pushIntFlight:function(a,d){if(!(a in OSF.IntFlights)){OSF.IntFlights[a]=d;return b}return c},getIntFlight:function(a){if(OSF.IntFlights&&a in OSF.IntFlights)return OSF.IntFlights[a];else return NaN},parseFlightsFromWindowName:function(a,b){return OSF.OUtil.parseArrayWithDefault(OSF.OUtil.parseInfoFromWindowName(a,b,OSF.WindowNameItemKeys.Flights))},parseFlightsWithGivenFragment:function(a,c){return OSF.OUtil.parseArrayWithDefault(OSF.OUtil.parseInfoWithGivenFragment(B,G,b,a,c))},parseArrayWithDefault:function(b){var a=[];try{a=JSON.parse(b)}catch(c){}if(!Array.isArray(a))a=[];return a},parseInfoFromWindowName:function(g,h,f){try{var b=JSON.parse(h),c=b!=a?b[f]:a,d=t();if(!g&&d&&b!=a){var e=b[OSF.WindowNameItemKeys.BaseFrameName]+f;if(c)d.setItem(e,c);else c=d.getItem(e)}return c}catch(i){return a}},parseInfoWithGivenFragment:function(m,j,k,i,l){var f=l.split(m),b=f.length>1?f[f.length-1]:a;if(k&&b!=a){if(b.indexOf(y)>=0)b=b.split(y)[0];b=decodeURIComponent(b)}var c=t();if(!i&&c){var e=window.name.indexOf(j);if(e>d){var g=window.name.indexOf(";",e);if(g==d)g=window.name.length;var h=window.name.substring(e,g);if(b)c.setItem(h,b);else b=c.getItem(h)}}return b},getConversationId:function(){var c=window.location.search,b=a;if(c){var d=c.indexOf("&");b=d>0?c.substring(1,d):c.substr(1);if(b&&b.charAt(b.length-1)==="="){b=b.substring(0,b.length-1);if(b)b=decodeURIComponent(b)}}return b},getInfoItems:function(b){var a=b.split("$");if(typeof a[1]==f)a=b.split("|");if(typeof a[1]==f)a=b.split("%7C");return a},getXdmFieldValue:function(f,d){var b=e,c=OSF.OUtil.parseXdmInfo(d);if(c){var a=OSF.OUtil.getInfoItems(c);if(a!=undefined&&a.length>=3)switch(f){case OSF.XdmFieldName.ConversationUrl:b=a[2];break;case OSF.XdmFieldName.AppId:b=a[1]}}return b},validateParamObject:function(f,e){var a=Function._validateParams(arguments,[{name:"params",type:Object,mayBeNull:c},{name:"expectedProperties",type:Object,mayBeNull:c},{name:"callback",type:Function,mayBeNull:b}]);if(a)throw a;for(var d in e){a=Function._validateParameter(f[d],e[d],d);if(a)throw a}},writeProfilerMark:function(a){if(window.msWriteProfilerMark){window.msWriteProfilerMark(a);OsfMsAjaxFactory.msAjaxDebug.trace(a)}},outputDebug:function(a){typeof OsfMsAjaxFactory!==f&&OsfMsAjaxFactory.msAjaxDebug&&OsfMsAjaxFactory.msAjaxDebug.trace&&OsfMsAjaxFactory.msAjaxDebug.trace(a)},defineNondefaultProperty:function(e,f,a,c){a=a||{};for(var g in c){var d=c[g];if(a[d]==undefined)a[d]=b}Object.defineProperty(e,f,a);return e},defineNondefaultProperties:function(c,a,d){a=a||{};for(var b in a)OSF.OUtil.defineNondefaultProperty(c,b,a[b],d);return c},defineEnumerableProperty:function(c,b,a){return OSF.OUtil.defineNondefaultProperty(c,b,a,[i])},defineEnumerableProperties:function(b,a){return OSF.OUtil.defineNondefaultProperties(b,a,[i])},defineMutableProperty:function(c,b,a){return OSF.OUtil.defineNondefaultProperty(c,b,a,[p,i,q])},defineMutableProperties:function(b,a){return OSF.OUtil.defineNondefaultProperties(b,a,[p,i,q])},finalizeProperties:function(e,d){d=d||{};for(var g=Object.getOwnPropertyNames(e),i=g.length,f=0;f<i;f++){var h=g[f],a=Object.getOwnPropertyDescriptor(e,h);if(!a.get&&!a.set)a.writable=d.writable||c;a.configurable=d.configurable||c;a.enumerable=d.enumerable||b;Object.defineProperty(e,h,a)}return e},mapList:function(a,c){var b=[];if(a)for(var d in a)b.push(c(a[d]));return b},listContainsKey:function(d,e){for(var a in d)if(e==a)return b;return c},listContainsValue:function(a,d){for(var e in a)if(d==a[e])return b;return c},augmentList:function(a,b){var d=a.push?function(c,b){a.push(b)}:function(c,b){a[c]=b};for(var c in b)d(c,b[c])},redefineList:function(a,b){for(var d in a)delete a[d];for(var c in b)a[c]=b[c]},isArray:function(a){return Object.prototype.toString.apply(a)==="[object Array]"},isFunction:function(a){return Object.prototype.toString.apply(a)==="[object Function]"},isDate:function(a){return Object.prototype.toString.apply(a)==="[object Date]"},addEventListener:function(a,b,d){if(a.a
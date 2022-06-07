/*! Outlook specific API library */
/*! Version: 16.0.6807.1000 */
/*! Update: 8 */
/*!
	Copyright (c) Microsoft Corporation.  All rights reserved.
*/

/*!
	Your use of this file is governed by the Microsoft Services Agreement http://go.microsoft.com/fwlink/?LinkId=266419.
*/
var __extends=this&&this.__extends||function(b,a){for(var c in a)if(a.hasOwnProperty(c))b[c]=a[c];function d(){this.constructor=b}b.prototype=a===null?Object.create(a):(d.prototype=a.prototype,new d)},OfficeExt;(function(b){var a=function(){function a(){}a.prototype.isMsAjaxLoaded=function(){return typeof Sys!=="undefined"&&typeof Type!=="undefined"&&Sys.StringBuilder&&typeof Sys.StringBuilder==="function"&&Type.registerNamespace&&typeof Type.registerNamespace==="function"&&Type.registerClass&&typeof Type.registerClass==="function"&&typeof Function._validateParams==="function"?true:false};a.prototype.loadMsAjaxFull=function(b){var a=(window.location.protocol.toLowerCase()==="https:"?"https:":"http:")+"//ajax.aspnetcdn.com/ajax/3.5/MicrosoftAjax.js";OSF.OUtil.loadScript(a,b)};Object.defineProperty(a.prototype,"msAjaxError",{"get":function(){if(this._msAjaxError==null&&this.isMsAjaxLoaded())this._msAjaxError=Error;return this._msAjaxError},"set":function(a){this._msAjaxError=a},enumerable:true,configurable:true});Object.defineProperty(a.prototype,"msAjaxSerializer",{"get":function(){if(this._msAjaxSerializer==null&&this.isMsAjaxLoaded())this._msAjaxSerializer=Sys.Serialization.JavaScriptSerializer;return this._msAjaxSerializer},"set":function(a){this._msAjaxSerializer=a},enumerable:true,configurable:true});Object.defineProperty(a.prototype,"msAjaxString",{"get":function(){if(this._msAjaxString==null&&this.isMsAjaxLoaded())this._msAjaxSerializer=String;return this._msAjaxString},"set":function(a){this._msAjaxString=a},enumerable:true,configurable:true});Object.defineProperty(a.prototype,"msAjaxDebug",{"get":function(){if(this._msAjaxDebug==null&&this.isMsAjaxLoaded())this._msAjaxDebug=Sys.Debug;return this._msAjaxDebug},"set":function(a){this._msAjaxDebug=a},enumerable:true,configurable:true});return a}();b.MicrosoftAjaxFactory=a})(OfficeExt||(OfficeExt={}));var OsfMsAjaxFactory=new OfficeExt.MicrosoftAjaxFactory,OSF=OSF||{},OfficeExt;(function(b){var a=function(){function a(a){this._internalStorage=a}a.prototype.getItem=function(a){try{return this._internalStorage&&this._internalStorage.getItem(a)}catch(b){return null}};a.prototype.setItem=function(b,a){try{this._internalStorage&&this._internalStorage.setItem(b,a)}catch(c){}};a.prototype.clear=function(){try{this._internalStorage&&this._internalStorage.clear()}catch(a){}};a.prototype.removeItem=function(a){try{this._internalStorage&&this._internalStorage.removeItem(a)}catch(b){}};a.prototype.getKeysWithPrefix=function(d){var b=[];try{for(var e=this._internalStorage&&this._internalStorage.length||0,a=0;a<e;a++){var c=this._internalStorage.key(a);c.indexOf(d)===0&&b.push(c)}}catch(f){}return b};return a}();b.SafeStorage=a})(OfficeExt||(OfficeExt={}));OSF.OUtil=function(){var f=-1,j="&_xdm_Info=",g="&_serializer_version=",h="_xdm_",m="_serializer_version=",c="#",e={},l=3e4,b=null,d=null,a=+new Date;function k(){var b=2147483647*Math.random();b^=a^(new Date).getMilliseconds()<<Math.floor(Math.random()*(31-10));return b.toString(16)}function i(){if(!b){try{var a=window.sessionStorage}catch(c){a=null}b=new OfficeExt.SafeStorage(a)}return b}return{set_entropy:function(b){if(typeof b=="string")for(var c=0;c<b.length;c+=4){for(var e=0,d=0;d<4&&c+d<b.length;d++)e=(e<<8)+b.charCodeAt(c+d);a^=e}else if(typeof b=="number")a^=b;else a^=2147483647*Math.random();a&=2147483647},extend:function(b,a){var c=function(){};c.prototype=a.prototype;b.prototype=new c;b.prototype.constructor=b;b.uber=a.prototype;if(a.prototype.constructor===Object.prototype.constructor)a.prototype.constructor=a},setNamespace:function(b,a){if(a&&b&&!a[b])a[b]={}},unsetNamespace:function(b,a){if(a&&b&&a[b])delete a[b]},loadScript:function(c,d,f){if(c&&d){var i=window.document,a=e[c];if(!a){var b=i.createElement("script");b.type="text/javascript";a={loaded:false,pendingCallbacks:[d],timer:null};e[c]=a;var g=function(){if(a.timer!=null){clearTimeout(a.timer);delete a.timer}a.loaded=true;for(var c=a.pendingCallbacks.length,b=0;b<c;b++){var d=a.pendingCallbacks.shift();d()}},h=function(){delete e[c];if(a.timer!=null){clearTimeout(a.timer);delete a.timer}for(var d=a.pendingCallbacks.length,b=0;b<d;b++){var f=a.pendingCallbacks.shift();f()}};if(b.readyState)b.onreadystatechange=function(){if(b.readyState=="loaded"||b.readyState=="complete"){b.onreadystatechange=null;g()}};else b.onload=g;b.onerror=h;f=f||l;a.timer=setTimeout(h,f);b.src=c;i.getElementsByTagName("head")[0].appendChild(b)}else if(a.loaded)d();else a.pendingCallbacks.push(d)}},loadCSS:function(c){if(c){var b=window.document,a=b.createElement("link");a.type="text/css";a.rel="stylesheet";a.href=c;b.getElementsByTagName("head")[0].appendChild(a)}},parseEnum:function(b,c){var a=c[b.trim()];if(typeof a=="undefined"){OsfMsAjaxFactory.msAjaxDebug.trace("invalid enumeration string:"+b);throw OsfMsAjaxFactory.msAjaxError.argument("str");}return a},delayExecutionAndCache:function(){var a={calc:arguments[0]};return function(){if(a.calc){a.val=a.calc.apply(this,arguments);delete a.calc}return a.val}},getUniqueId:function(){f=f+1;return f.toString()},formatString:function(){var a=arguments,b=a[0];return b.replace(/{(\d+)}/gm,function(d,b){var c=parseInt(b,10)+1;return a[c]===undefined?"{"+b+"}":a[c]})},generateConversationId:function(){return[k(),k(),(+new Date).toString()].join("_")},getFrameNameAndConversationId:function(b,c){var a=h+b+this.generateConversationId();c.setAttribute("name",a);return this.generateConversationId()},addXdmInfoAsHash:function(b,a){return OSF.OUtil.addInfoAsHash(b,j,a)},addSerializerVersionAsHash:function(b,a){return OSF.OUtil.addInfoAsHash(b,g,a)},addInfoAsHash:function(a,g,e){a=a.trim()||"";var b=a.split(c),d=b.shift(),f=b.join(c);return[d,c,f,g,e].join("")},parseXdmInfo:function(a){return OSF.OUtil.parseXdmInfoWithGivenFragment(a,window.location.hash)},parseXdmInfoWithGivenFragment:function(a,b){return OSF.OUtil.parseInfoWithGivenFragment(j,h,a,b)},parseSerializerVersion:function(a){return OSF.OUtil.parseSerializerVersionWithGivenFragment(a,window.location.hash)},parseSerializerVersionWithGivenFragment:function(a,b){return parseInt(OSF.OUtil.parseInfoWithGivenFragment(g,m,a,b))},parseInfoWithGivenFragment:function(k,h,g,j){var d=j.split(k),a=d.length>1?d[d.length-1]:null,b=i();if(!g&&b){var c=window.name.indexOf(h);if(c>-1){var e=window.name.indexOf(";",c);if(e==-1)e=window.name.length;var f=window.name.substring(c,e);if(a)b.setItem(f,a);else a=b.getItem(f)}}return a},getConversationId:function(){var b=window.location.search,a=null;if(b){var c=b.indexOf("&");a=c>0?b.substring(1,c):b.substr(1);if(a&&a.charAt(a.length-1)==="="){a=a.substring(0,a.length-1);if(a)a=decodeURIComponent(a)}}return a},getInfoItems:function(b){var a=b.split("$");if(typeof a[1]=="undefined")a=b.split("|");return a},getConversationUrl:function(){var b="",c=OSF.OUtil.parseXdmInfo(true);if(c){var a=OSF.OUtil.getInfoItems(c);if(a!=undefined&&a.length>=3)b=a[2]}return b},validateParamObject:function(d,c){var a=Function._validateParams(arguments,[{name:"params",type:Object,mayBeNull:false},{name:"expectedProperties",type:Object,mayBeNull:false},{name:"callback",type:Function,mayBeNull:true}]);if(a)throw a;for(var b in c){a=Function._validateParameter(d[b],c[b],b);if(a)throw a;}},writeProfilerMark:function(a){if(window.msWriteProfilerMark){window.msWriteProfilerMark(a);OsfMsAjaxFactory.msAjaxDebug.trace(a)}},outputDebug:function(a){typeof Sys!=="undefined"&&Sys&&Sys.Debug&&OsfMsAjaxFactory.msAjaxDebug.trace(a)},defineNondefaultProperty:function(d,e,a,b){a=a||{};for(var f in b){var c=b[f];if(a[c]==undefined)a[c]=true}Object.defineProperty(d,e,a);return d},defineNondefaultProperties:function(c,a,d){a=a||{};for(var b in a)OSF.OUtil.defineNondefaultProperty(c,b,a[b],d);return c},defineEnumerableProperty:function(c,b,a){return OSF.OUtil.defineNondefaultProperty(c,b,a,["enumerable"])},defineEnumerableProperties:function(b,a){return OSF.OUtil.defineNondefaultProperties(b,a,["enumerable"])},defineMutableProperty:function(c,b,a){return OSF.OUtil.defineNondefaultProperty(c,b,a,["writable","enumerable","configurable"])},defineMutableProperties:function(b,a){return OSF.OUtil.defineNondefaultProperties(b,a,["writable","enumerable","configurable"])},finalizeProperties:function(c,b){b=b||{};for(var e=Object.getOwnPropertyNames(c),g=e.length,d=0;d<g;d++){var f=e[d],a=Object.getOwnPropertyDescriptor(c,f);if(!a.get&&!a.set)a.writable=b.writable||false;a.configurable=b.configurable||false;a.enumerable=b.enumerable||true;Object.defineProperty(c,f,a)}return c},mapList:function(a,c){var b=[];if(a)for(var d in a)b.push(c(a[d]));return b},listContainsKey:function(b,c){for(var a in b)if(c==a)return true;return false},listContainsValue:function(a,b){for(var c in a)if(b==a[c])return true;return false},augmentList:function(a,b){var d=a.push?function(c,b){a.push(b)}:function(c,b){a[c]=b};for(var c in b)d(c,b[c])},redefineList:function(a,b){for(var d in a)delete a[d];for(var c in b)a[c]=b[c]},isArray:function(a){return Object.prototype.toString.apply(a)==="[object Array]"},isFunction:function(a){return Object.prototype.toString.apply(a)==="[object Function]"},isDate:function(a){return Object.prototype.toString.apply(a)==="[object Date]"},addEventListener:function(a,b,c){if(a.addEventListener)a.addEventListener(b,c,false);else if(Sys.Browser.agent===Sys.Browser.InternetExplorer&&a.attachEvent)a.attachEvent("on"+b,c);else a["on"+b]=c},removeEventListener:function(a,b,c){if(a.removeEventListener)a.removeEventListener(b,c,false);else if(Sys.Browser.agent===Sys.Browser.InternetExplorer&&a.detachEvent)a.detachEvent("on"+b,c);else a["on"+b]=null},getCookieValue:function(b){var a=RegExp(b+"[^;]+").exec(document.cookie);return a.toString().replace(/^[^=]+./,"")},xhrGet:function(d,c,b){var a;try{a=new XMLHttpRequest;a.onreadystatechange=function(){if(a.readyState==4)if(a.status==200)c(a.responseText);else b(a.status)};a.open("GET",d,true);a.send()}catch(e){b(e)}},xhrGetFull:function(f,d,e,b){var a,c=d;try{a=new XMLHttpRequest;a.onreadystatechange=function(){if(a.readyState==4)if(a.status==200)e(a,c);else b(a.status)};a.open("GET",f,true);a.send()}catch(g){b(g)}},encodeBase64:function(c){if(!c)return c;var n="ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/=",l=[],b=[],h=0,j,g,i,d,e,f,a,m=c.length;do{j=c.charCodeAt(h++);g=c.charCodeAt(h++);i=c.charCodeAt(h++);a=0;d=j&255;e=j>>8;f=g&255;b[a++]=d>>2;b[a++]=(d&3)<<4|e>>4;b[a++]=(e&15)<<2|f>>6;b[a++]=f&63;if(!isNaN(g)){d=g>>8;e=i&255;f=i>>8;b[a++]=d>>2;b[a++]=(d&3)<<4|e>>4;b[a++]=(e&15)<<2|f>>6;b[a++]=f&63}if(isNaN(g))b[a-1]=64;else if(isNaN(i)){b[a-2]=64;b[a-1]=64}for(var k=0;k<a;k++)l.push(n.charAt(b[k]))}while(h<m);return l.join("")},getSessionStorage:function(){return i()},getLocalStorage:function(){if(!d){try{var a=window.localStorage}catch(b){a=null}d=new OfficeExt.SafeStorage(a)}return d},convertIntToCssHexColor:function(b){return"#"+(Number(b)+16777216).toString(16).slice(-6)},attachClickHandler:function(a,b){a.onclick=function(){b()};a.ontouchend=function(a){b();a.preventDefault()}},getQueryStringParamValue:function(a,c){var d=Function._validateParams(arguments,[{name:"queryString",type:String,mayBeNull:false},{name:"paramName",type:String,mayBeNull:false}]);if(d){OsfMsAjaxFactory.msAjaxDebug.trace("OSF_Outil_getQueryStringParamValue: Parameters cannot be null.");return""}var b=new RegExp("[\\?&]"+c+"=([^&#]*)","i");if(!b.test(a)){OsfMsAjaxFactory.msAjaxDebug.trace("OSF_Outil_getQueryStringParamValue: The parameter is not found.");return""}return b.exec(a)[1]},isiOS:function(){return window.navigator.userAgent.match(/(iPad|iPhone|iPod)/g)?true:false},shallowCopy:function(a){var c=a.constructor();for(var b in a)if(a.hasOwnProperty(b))c[b]=a[b];return c},serializeOMEXResponseErrorMessage:function(a){if(typeof JSON!=="undefined")try{return JSON.stringify(a)}catch(b){}return""},createObject:function(a){var c=null;if(a){c={};for(var d=a.length,b=0;b<d;b++)c[a[b].name]=a[b].value}return c}}}();OSF.OUtil.Guid=function(){var a=["0","1","2","3","4","5","6","7","8","9","a","b","c","d","e","f"];return{generateNewGuid:function(){for(var c="",d=+new Date,b=0;b<32&&d>0;b++){if(b==8||b==12||b==16||b==20)c+="-";c+=a[d%16];d=Math.floor(d/16)}for(;b<32;b++){if(b==8||b==12||b==16||b==20)c+="-";c+=a[Math.floor(Math.random()*16)]}return c}}}();window.OSF=OSF;OSF.OUtil.setNamespace("OSF",window);OSF.AppName={Unsupported:0,Excel:1,Word:2,PowerPoint:4,Outlook:8,ExcelWebApp:16,WordWebApp:32,OutlookWebApp:64,Project:128,AccessWebApp:256,PowerpointWebApp:512,ExcelIOS:1024,Sway:2048,WordIOS:4096,PowerPointIOS:8192,Access:16384,Lync:32768,OutlookIOS:65536,OneNoteWebApp:131072};OSF.InternalPerfMarker={DataCoercionBegin:"Agave.HostCall.CoerceDataStart",DataCoercionEnd:"Agave.HostCall.CoerceDataEnd"};OSF.HostCallPerfMarker={IssueCall:"Agave.HostCall.IssueCall",ReceiveResponse:"Agave.HostCall.ReceiveResponse",RuntimeExceptionRaised:"Agave.HostCall.RuntimeExecptionRaised"};OSF.AgaveHostAction={Select:0,UnSelect:1,CancelDialog:2,InsertAgave:3,CtrlF6In:4,CtrlF6Exit:5,CtrlF6ExitShift:6,SelectWithError:7,NotifyHostError:8,RefreshAddinCommands:9};OSF.SharedConstants={NotificationConversationIdSuffix:"_ntf"};OSF.DialogMessageType={DialogMessageReceived:0,DialogClosed:1,NavigationFailed:2,InvalidSchema:3};OSF.OfficeAppContext=function(q,m,i,h,k,n,j,l,p,d,o,f,e,g,c,b,a){this._id=q;this._appName=m;this._appVersion=i;this._appUILocale=h;this._dataLocale=k;this._docUrl=n;this._clientMode=j;this._settings=l;this._reason=p;this._osfControlType=d;this._eToken=o;this._correlationId=f;this._appInstanceId=e;this._touchEnabled=g;this._commerceAllowed=c;this._appMinorVersion=b;this._requirementMatrix=a;this._isDialog=false;this.get_id=function(){return this._id};this.get_appName=function(){return this._appName};this.get_appVersion=function(){return"16.0.4909.1000"};this.get_appUILocale=function(){return this._appUILocale};this.get_dataLocale=function(){return this._dataLocale};this.get_docUrl=function(){return this._docUrl};this.get_clientMode=function(){return this._clientMode};this.get_bindings=function(){return this._bindings};this.get_settings=function(){return this._settings};this.get_reason=function(){return this._reason};this.get_osfControlType=function(){return this._osfControlType};this.get_eToken=function(){return this._eToken};this.get_correlationId=function(){return this._correlationId};this.get_appInstanceId=function(){return this._appInstanceId};this.get_touchEnabled=function(){return this._touchEnabled};this.get_commerceAllowed=function(){return this._commerceAllowed};this.get_appMinorVersion=function(){return this._appMinorVersion};this.get_requirementMatrix=function(){return this._requirementMatrix};this.get_isDialog=function(){return this._isDialog}};OSF.OsfControlType={DocumentLevel:0,ContainerLevel:1};OSF.ClientMode={ReadOnly:0,ReadWrite:1};OSF.OUtil.setNamespace("Microsoft",window);OSF.OUtil.setNamespace("Office",Microsoft);OSF.OUtil.setNamespace("Client",Microsoft.Office);OSF.OUtil.setNamespace("WebExtension",Microsoft.Office);Microsoft.Office.WebExtension.InitializationReason={Inserted:"inserted",DocumentOpened:"documentOpened"};Microsoft.Office.WebExtension.ValueFormat={Unformatted:"unformatted",Formatted:"formatted"};Microsoft.Office.WebExtension.FilterType={All:"all"};Microsoft.Office.WebExtension.Parameters={BindingType:"bindingType",CoercionType:"coercionType",ValueFormat:"valueFormat",FilterType:"filterType",Columns:"columns",SampleData:"sampleData",GoToType:"goToType",SelectionMode:"selectionMode",Id:"id",PromptText:"promptText",ItemName:"itemName",FailOnCollision:"failOnCollision",StartRow:"startRow",StartColumn:"startColumn",RowCount:"rowCount",ColumnCount:"columnCount",Callback:"callback"
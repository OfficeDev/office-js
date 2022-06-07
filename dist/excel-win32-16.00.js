/* Excel specific API library */
/* Version: 16.0.6216.3006 */
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

var __extends=this.__extends||function(n,t){function r(){this.constructor=n}for(var i in t)t.hasOwnProperty(i)&&(n[i]=t[i]);r.prototype=t.prototype;n.prototype=new r},OsfMsAjaxFactory,OSF,msAjaxCDNPath,OSFRichclient,OfficeExt,OSFLog,Logger,OSFAppTelemetry,OfficeExtension,Excel;(function(n){var t=function(){function t(){}var i=null,n=!0;return t.prototype.isMsAjaxLoaded=function(){var t="function",i="undefined";return typeof Sys!==i&&typeof Type!==i&&Sys.StringBuilder&&typeof Sys.StringBuilder===t&&Type.registerNamespace&&typeof Type.registerNamespace===t&&Type.registerClass&&typeof Type.registerClass===t&&typeof Function._validateParams===t?n:!1},t.prototype.loadMsAjaxFull=function(n){var t=(window.location.protocol.toLowerCase()==="https:"?"https:":"http:")+"//ajax.aspnetcdn.com/ajax/3.5/MicrosoftAjax.js";OSF.OUtil.loadScript(t,n)},Object.defineProperty(t.prototype,"msAjaxError",{get:function(){var n=this;return n._msAjaxError==i&&n.isMsAjaxLoaded()&&(n._msAjaxError=Error),n._msAjaxError},set:function(n){this._msAjaxError=n},enumerable:n,configurable:n}),Object.defineProperty(t.prototype,"msAjaxSerializer",{get:function(){var n=this;return n._msAjaxSerializer==i&&n.isMsAjaxLoaded()&&(n._msAjaxSerializer=Sys.Serialization.JavaScriptSerializer),n._msAjaxSerializer},set:function(n){this._msAjaxSerializer=n},enumerable:n,configurable:n}),Object.defineProperty(t.prototype,"msAjaxString",{get:function(){var n=this;return n._msAjaxString==i&&n.isMsAjaxLoaded()&&(n._msAjaxSerializer=String),n._msAjaxString},set:function(n){this._msAjaxString=n},enumerable:n,configurable:n}),Object.defineProperty(t.prototype,"msAjaxDebug",{get:function(){var n=this;return n._msAjaxDebug==i&&n.isMsAjaxLoaded()&&(n._msAjaxDebug=Sys.Debug),n._msAjaxDebug},set:function(n){this._msAjaxDebug=n},enumerable:n,configurable:n}),t}();n.MicrosoftAjaxFactory=t})(OfficeExt||(OfficeExt={}));OsfMsAjaxFactory=new OfficeExt.MicrosoftAjaxFactory;OSF=OSF||{},function(n){var t=function(){function n(n){this._internalStorage=n}return n.prototype.getItem=function(n){try{return this._internalStorage&&this._internalStorage.getItem(n)}catch(t){return null}},n.prototype.setItem=function(n,t){try{this._internalStorage&&this._internalStorage.setItem(n,t)}catch(i){}},n.prototype.clear=function(){try{this._internalStorage&&this._internalStorage.clear()}catch(n){}},n.prototype.removeItem=function(n){try{this._internalStorage&&this._internalStorage.removeItem(n)}catch(t){}},n.prototype.getKeysWithPrefix=function(n){var r=[],u,t,i;try{for(u=this._internalStorage&&this._internalStorage.length||0,t=0;t<u;t++)i=this._internalStorage.key(t),i.indexOf(n)===0&&r.push(i)}catch(f){}return r},n}();n.SafeStorage=t}(OfficeExt||(OfficeExt={}));OSF.OUtil=function(){function k(){var n=o*Math.random();return n^=r^(new Date).getMilliseconds()<<Math.floor(Math.random()*21),n.toString(16)}function d(){if(!l){try{var n=window.sessionStorage}catch(i){n=t}l=new OfficeExt.SafeStorage(n)}return l}var u="on",v="configurable",y="writable",f="enumerable",e="undefined",i=!0,n=!1,o=2147483647,t=null,s=-1,p="&_xdm_Info=",w="&_serializer_version=",b="_xdm_",g="_serializer_version=",h="#",c={},nt=3e4,l=t,a=t,r=(new Date).getTime();return{set_entropy:function(n){var t,u,i;if(typeof n=="string")for(t=0;t<n.length;t+=4){for(u=0,i=0;i<4&&t+i<n.length;i++)u=(u<<8)+n.charCodeAt(t+i);r^=u}else r^=typeof n=="number"?n:o*Math.random();r&=o},extend:function(n,t){var i=function(){};i.prototype=t.prototype;n.prototype=new i;n.prototype.constructor=n;n.uber=t.prototype;t.prototype.constructor===Object.prototype.constructor&&(t.prototype.constructor=t)},setNamespace:function(n,t){t&&n&&!t[n]&&(t[n]={})},unsetNamespace:function(n,t){t&&n&&t[n]&&delete t[n]},loadScript:function(r,u,f){var s,e,o,h,l;r&&u&&(s=window.document,e=c[r],e?e.loaded?u():e.pendingCallbacks.push(u):(o=s.createElement("script"),o.type="text/javascript",e={loaded:n,pendingCallbacks:[u],timer:t},c[r]=e,h=function(){var r,n,u;for(e.timer!=t&&(clearTimeout(e.timer),delete e.timer),e.loaded=i,r=e.pendingCallbacks.length,n=0;n<r;n++)u=e.pendingCallbacks.shift(),u()},l=function(){var i,n,u;for(delete c[r],e.timer!=t&&(clearTimeout(e.timer),delete e.timer),i=e.pendingCallbacks.length,n=0;n<i;n++)u=e.pendingCallbacks.shift(),u()},o.readyState?o.onreadystatechange=function(){(o.readyState=="loaded"||o.readyState=="complete")&&(o.onreadystatechange=t,h())}:o.onload=h,o.onerror=l,f=f||nt,e.timer=setTimeout(l,f),o.src=r,s.getElementsByTagName("head")[0].appendChild(o)))},loadCSS:function(n){if(n){var i=window.document,t=i.createElement("link");t.type="text/css";t.rel="stylesheet";t.href=n;i.getElementsByTagName("head")[0].appendChild(t)}},parseEnum:function(n,t){var i=t[n.trim()];if(typeof i==e){OsfMsAjaxFactory.msAjaxDebug.trace("invalid enumeration string:"+n);throw OsfMsAjaxFactory.msAjaxError.argument("str");}return i},delayExecutionAndCache:function(){var n={calc:arguments[0]};return function(){return n.calc&&(n.val=n.calc.apply(this,arguments),delete n.calc),n.val}},getUniqueId:function(){return s=s+1,s.toString()},formatString:function(){var n=arguments,t=n[0];return t.replace(/{(\d+)}/gm,function(t,i){var r=parseInt(i,10)+1;return n[r]===undefined?"{"+i+"}":n[r]})},generateConversationId:function(){return[k(),k(),(new Date).getTime().toString()].join("_")},getFrameNameAndConversationId:function(n,t){var i=b+n+this.generateConversationId();return t.setAttribute("name",i),this.generateConversationId()},addXdmInfoAsHash:function(n,t){return OSF.OUtil.addInfoAsHash(n,p,t)},addSerializerVersionAsHash:function(n,t){return OSF.OUtil.addInfoAsHash(n,w,t)},addInfoAsHash:function(n,t,i){n=n.trim()||"";var r=n.split(h),u=r.shift(),f=r.join(h);return[u,h,f,t,i].join("")},parseXdmInfo:function(n){return OSF.OUtil.parseXdmInfoWithGivenFragment(n,window.location.hash)},parseXdmInfoWithGivenFragment:function(n,t){return OSF.OUtil.parseInfoWithGivenFragment(p,b,n,t)},parseSerializerVersion:function(n){return OSF.OUtil.parseSerializerVersionWithGivenFragment(n,window.location.hash)},parseSerializerVersionWithGivenFragment:function(n,t){return parseInt(OSF.OUtil.parseInfoWithGivenFragment(w,g,n,t))},parseInfoWithGivenFragment:function(n,i,r,u){var s=u.split(n),f=s.length>1?s[s.length-1]:t,h=d(),e,o,c;return!r&&h&&(e=window.name.indexOf(i),e>-1&&(o=window.name.indexOf(";",e),o==-1&&(o=window.name.length),c=window.name.substring(e,o),f?h.setItem(c,f):f=h.getItem(c))),f},getConversationId:function(){var i=window.location.search,n=t,r;return i&&(r=i.indexOf("&"),n=r>0?i.substring(1,r):i.substr(1),n&&n.charAt(n.length-1)==="="&&(n=n.substring(0,n.length-1),n&&(n=decodeURIComponent(n)))),n},getInfoItems:function(n){var t=n.split("$");return typeof t[1]==e&&(t=n.split("|")),t},getConversationUrl:function(){var t="",r=OSF.OUtil.parseXdmInfo(i),n;return r&&(n=OSF.OUtil.getInfoItems(r),n!=undefined&&n.length>=3&&(t=n[2])),t},validateParamObject:function(t,r){var u=Function._validateParams(arguments,[{name:"params",type:Object,mayBeNull:n},{name:"expectedProperties",type:Object,mayBeNull:n},{name:"callback",type:Function,mayBeNull:i}]),f;if(u)throw u;for(f in r)if(u=Function._validateParameter(t[f],r[f],f),u)throw u;},writeProfilerMark:function(n){window.msWriteProfilerMark&&(window.msWriteProfilerMark(n),OsfMsAjaxFactory.msAjaxDebug.trace(n))},outputDebug:function(n){typeof Sys!==e&&Sys&&Sys.Debug&&OsfMsAjaxFactory.msAjaxDebug.trace(n)},defineNondefaultProperty:function(n,t,r,u){var e,f;r=r||{};for(e in u)f=u[e],r[f]==undefined&&(r[f]=i);return Object.defineProperty(n,t,r),n},defineNondefaultProperties:function(n,t,i){t=t||{};for(var r in t)OSF.OUtil.defineNondefaultProperty(n,r,t[r],i);return n},defineEnumerableProperty:function(n,t,i){return OSF.OUtil.defineNondefaultProperty(n,t,i,[f])},defineEnumerableProperties:function(n,t){return OSF.OUtil.defineNondefaultProperties(n,t,[f])},defineMutableProperty:function(n,t,i){return OSF.OUtil.defineNondefaultProperty(n,t,i,[y,f,v])},defineMutableProperties:function(n,t){return OSF.OUtil.defineNondefaultProperties(n,t,[y,f,v])},finalizeProperties:function(t,r){var e,u;r=r||{};for(var o=Object.getOwnPropertyNames(t),s=o.length,f=0;f<s;f++)e=o[f],u=Object.getOwnPropertyDescriptor(t,e),u.get||u.set||(u.writable=r.writable||n),u.configurable=r.configurable||n,u.enumerable=r.enumerable||i,Object.defineProperty(t,e,u);return t},mapList:function(n,t){var i=[],r;if(n)for(r in n)i.push(t(n[r]));return i},listContainsKey:function(t,r){for(var u in t)if(r==u)return i;return n},listContainsValue:function(t,r){for(var u in t)if(r==t[u])return i;return n},augmentList:function(n,t){var r=n.push?function(t,i){n.push(i)}:function(t,i){n[t]=i};for(var i in t)r(i,t[i])},redefineList:function(n,t){var r,i;for(r in n)delete n[r];for(i in t)n[i]=t[i]},isArray:function(n){return Object.prototype.toString.apply(n)==="[object Array]"},isFunction:function(n){return Object.prototype.toString.apply(n)==="[object Function]"},isDate:function(n){return Object.prototype.toString.apply(n)==="[object Date]"},addEventListener:function(t,i,r){t.addEventListener?t.addEventListener(i,r,n):Sys.Browser.agent===Sys.Browser.InternetExplorer&&t.attachEvent?t.attachEvent(u+i,r):t[u+i]=r},removeEventListener:function(i,r,f){i.removeEventListener?i.removeEventListener(r,f,n):Sys.Browser.agent===Sys.Browser.InternetExplorer&&i.detachEvent?i.detachEvent(u+r,f):i[u+r]=t},xhrGet:function(n,t,r){var u;try{u=new XMLHttpRequest;u.onreadystatechange=function(){u.readyState==4&&(u.status==200?t(u.responseText):r(u.status))};u.open("GET",n,i);u.send()}catch(f){r(f)}},encodeBase64:function(n){var h;if(!n)return n;var a="ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/=",l=[],i=[],o=0,c,e,s,r,u,f,t,v=n.length;do for(c=n.charCodeAt(o++),e=n.charCodeAt(o++),s=n.charCodeAt(o++),t=0,r=c&255,u=c>>8,f=e&255,i[t++]=r>>2,i[t++]=(r&3)<<4|u>>4,i[t++]=(u&15)<<2|f>>6,i[t++]=f&63,isNaN(e)||(r=e>>8,u=s&255,f=s>>8,i[t++]=r>>2,i[t++]=(r&3)<<4|u>>4,i[t++]=(u&15)<<2|f>>6,i[t++]=f&63),isNaN(e)?i[t-1]=64:isNaN(s)&&(i[t-2]=64,i[t-1]=64),h=0;h<t;h++)l.push(a.charAt(i[h]));while(o<v);return l.join("")},getSessionStorage:function(){return d()},getLocalStorage:function(){if(!a){try{var n=window.localStorage}catch(i){n=t}a=new OfficeExt.SafeStorage(n)}return a},convertIntToCssHexColor:function(n){return"#"+(Number(n)+16777216).toString(16).slice(-6)},attachClickHandler:function(n,t){n.onclick=function(){t()};n.ontouchend=function(n){t();n.preventDefault()}},getQueryStringParamValue:function(t,i){var u=Function._validateParams(arguments,[{name:"queryString",type:String,mayBeNull:n},{name:"paramName",type:String,mayBeNull:n}]),r;return u?(OsfMsAjaxFactory.msAjaxDebug.trace("OSF_Outil_getQueryStringParamValue: Parameters cannot be null."),""):(r=new RegExp("[\\?&]"+i+"=([^&#]*)","i"),!r.test(t))?(OsfMsAjaxFactory.msAjaxDebug.trace("OSF_Outil_getQueryStringParamValue: The parameter is not found."),""):r.exec(t)[1]},isiOS:function(){return window.navigator.userAgent.match(/(iPad|iPhone|iPod)/g)?i:n},shallowCopy:function(n){var i=n.constructor();for(var t in n)n.hasOwnProperty(t)&&(i[t]=n[t]);return i}}}();OSF.OUtil.Guid=function(){var n=["0","1","2","3","4","5","6","7","8","9","a","b","c","d","e","f"];return{generateNewGuid:function(){for(var i="",r=(new Date).getTime(),t=0;t<32&&r>0;t++)(t==8||t==12||t==16||t==20)&&(i+="-"),i+=n[r%16],r=Math.floor(r/16);for(;t<32;t++)(t==8||t==12||t==16||t==20)&&(i+="-"),i+=n[Math.floor(Math.random()*16)];return i}}}();window.OSF=OSF;OSF.OUtil.setNamespace("OSF",window);OSF.AppName={Unsupported:0,Excel:1,Word:2,PowerPoint:4,Outlook:8,ExcelWebApp:16,WordWebApp:32,OutlookWebApp:64,Project:128,AccessWebApp:256,PowerpointWebApp:512,ExcelIOS:1024,Sway:2048,WordIOS:4096,PowerPointIOS:8192,Access:16384,Lync:32768,OutlookIOS:65536,OneNoteWebApp:131072};OSF.InternalPerfMarker={DataCoercionBegin:"Agave.HostCall.CoerceDataStart",DataCoercionEnd:"Agave.HostCall.CoerceDataEnd"};OSF.HostCallPerfMarker={IssueCall:"Agave.HostCall.IssueCall",ReceiveResponse:"Agave.HostCall.ReceiveResponse",RuntimeExceptionRaised:"Agave.HostCall.RuntimeExecptionRaised"};OSF.AgaveHostAction={Select:0,UnSelect:1,CancelDialog:2,InsertAgave:3,CtrlF6In:4,CtrlF6Exit:5,CtrlF6ExitShift:6,SelectWithError:7,NotifyHostError:8};OSF.SharedConstants={NotificationConversationIdSuffix:"_ntf"};OSF.OfficeAppContext=function(n,t,i,r,u,f,e,o,s,h,c,l,a,v,y,p,w){var b=this;b._id=n;b._appName=t;b._appVersion=i;b._appUILocale=r;b._dataLocale=u;b._docUrl=f;b._clientMode=e;b._settings=o;b._reason=s;b._osfControlType=h;b._eToken=c;b._correlationId=l;b._appInstanceId=a;b._touchEnabled=v;b._commerceAllowed=y;b._appMinorVersion=p;b._requirementMatrix=w;b.get_id=function(){return this._id};b.get_appName=function(){return this._appName};b.get_appVersion=function(){return this._appVersion};b.get_appUILocale=function(){return this._appUILocale};b.get_dataLocale=function(){return this._dataLocale};b.get_docUrl=function(){return this._docUrl};b.get_clientMode=function(){return this._clientMode};b.get_bindings=function(){return this._bindings};b.get_settings=function(){return this._settings};b.get_reason=function(){return this._reason};b.get_osfControlType=function(){return this._osfControlType};b.get_eToken=function(){return this._eToken};b.get_correlationId=function(){return this._correlationId};b.get_appInstanceId=function(){return this._appInstanceId};b.get_touchEnabled=function(){return this._touchEnabled};b.get_commerceAllowed=function(){return this._commerceAllowed};b.get_appMinorVersion=function(){return this._appMinorVersion};b.get_requirementMatrix=function(){return this._requirementMatrix}};OSF.OsfControlType={DocumentLevel:0,ContainerLevel:1};OSF.ClientMode={ReadOnly:0,ReadWrite:1};OSF.OUtil.setNamespace("Microsoft",window);OSF.OUtil.setNamespace("Office",Microsoft);OSF.OUtil.setNamespace("Client",Microsoft.Office);OSF.OUtil.setNamespace("WebExtension",Microsoft.Office);Microsoft.Office.WebExtension.InitializationReason={Inserted:"inserted",DocumentOpened:"documentOpened"};Microsoft.Office.WebExtension.ValueFormat={Unformatted:"unformatted",Formatted:"formatted"};Microsoft.Office.WebExtension.FilterType={All:"all"};Microsoft.Office.WebExtension.Parameters={BindingType:"bindingType",CoercionType:"coercionType",ValueFormat:"valueFormat",FilterType:"filterType",Columns:"columns",SampleData:"sampleData",GoToType:"goToType",SelectionMode:"selectionMode",Id:"id",PromptText:"promptText",ItemName:"itemName",FailOnCollision:"failOnCollision",StartRow:"startRow",StartColumn:"startColumn",RowCount:"rowCount",ColumnCount:"columnCount",Callback:"callback",AsyncContext:"asyncContext",Data:"data",Rows:"rows",OverwriteIfStale:"overwriteIfStale",FileType:"fileType",EventType:"eventType",Handler:"handler",SliceSize:"sliceSize",SliceIndex:"sliceIndex",ActiveView:"activeView",Status:"status",Xml:"xml",Namespace:"namespace",Prefix:"prefix",XPath:"xPath",ImageLeft:"imageLeft",ImageTop:"imageTop",ImageWidth:"imageWidth",ImageHeight:"imageHeight",TaskId:"taskId",FieldId:"fieldId",FieldValue:"fieldValue",ServerUrl:"serverUrl",ListName:"listName",ResourceId:"resourceId",ViewType:"viewType",ViewName:"viewName",GetRawValue:"getRawValue",CellFormat:"cellFormat",TableOptions:"tableOptions",TaskIndex:"taskIndex",ResourceIndex:"resourceIndex"};OSF.OUtil.setNamespace("DDA",OSF);OSF.DDA.DocumentMode={ReadOnly:1,ReadWrite:0};OSF.DDA.PropertyDescriptors={AsyncResultSt
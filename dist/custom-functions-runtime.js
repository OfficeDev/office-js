/* Office JavaScript API library - Custom Functions */

/*
	Copyright (c) Microsoft Corporation.  All rights reserved.
*/

/*
	This file incorporates the "whatwg-fetch" implementation, version 2.0.3, licensed under MIT with the following licensing notice:
	(See github.com/github/fetch/blob/master/LICENSE)

		Copyright (c) 2014-2016 GitHub, Inc.

		Permission is hereby granted, free of charge, to any person obtaining
		a copy of this software and associated documentation files (the
		"Software"), to deal in the Software without restriction, including
		without limitation the rights to use, copy, modify, merge, publish,
		distribute, sublicense, and/or sell copies of the Software, and to
		permit persons to whom the Software is furnished to do so, subject to
		the following conditions:

		The above copyright notice and this permission notice shall be
		included in all copies or substantial portions of the Software.

		THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND,
		EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF
		MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND
		NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE
		LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION
		OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION
		WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.

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
var OSF=OSF||{};OSF.ConstantNames={FileVersion:"0.0.0.0",OfficeJS:"custom-functions-runtime.js",OfficeDebugJS:"custom-functions-runtime.debug.js",HostFileScriptSuffix:"core",IsCustomFunctionsRuntime:true};var OSF=OSF||{};OSF.HostSpecificFileVersionDefault="16.00";OSF.HostSpecificFileVersionMap={access:{web:"16.00"},agavito:{winrt:"16.00"},excel:{ios:"16.00",mac:"16.00",web:"16.00",win32:"16.01",winrt:"16.00"},onenote:{android:"16.00",web:"16.00",win32:"16.00",winrt:"16.00"},outlook:{ios:"16.00",mac:"16.00",web:"16.01",win32:"16.02"},powerpoint:{ios:"16.00",mac:"16.00",web:"16.00",win32:"16.01",winrt:"16.00"},project:{win32:"16.00"},sway:{web:"16.00"},word:{ios:"16.00",mac:"16.00",web:"16.00",win32:"16.01",winrt:"16.00"},visio:{web:"16.00",win32:"16.00"}};OSF.SupportedLocales={"ar-sa":true,"bg-bg":true,"bn-in":true,"ca-es":true,"cs-cz":true,"da-dk":true,"de-de":true,"el-gr":true,"en-us":true,"es-es":true,"et-ee":true,"eu-es":true,"fa-ir":true,"fi-fi":true,"fr-fr":true,"gl-es":true,"he-il":true,"hi-in":true,"hr-hr":true,"hu-hu":true,"id-id":true,"it-it":true,"ja-jp":true,"kk-kz":true,"ko-kr":true,"lo-la":true,"lt-lt":true,"lv-lv":true,"ms-my":true,"nb-no":true,"nl-nl":true,"nn-no":true,"pl-pl":true,"pt-br":true,"pt-pt":true,"ro-ro":true,"ru-ru":true,"sk-sk":true,"sl-si":true,"sr-cyrl-cs":true,"sr-cyrl-rs":true,"sr-latn-cs":true,"sr-latn-rs":true,"sv-se":true,"th-th":true,"tr-tr":true,"uk-ua":true,"ur-pk":true,"vi-vn":true,"zh-cn":true,"zh-tw":true};OSF.AssociatedLocales={ar:"ar-sa",bg:"bg-bg",bn:"bn-in",ca:"ca-es",cs:"cs-cz",da:"da-dk",de:"de-de",el:"el-gr",en:"en-us",es:"es-es",et:"et-ee",eu:"eu-es",fa:"fa-ir",fi:"fi-fi",fr:"fr-fr",gl:"gl-es",he:"he-il",hi:"hi-in",hr:"hr-hr",hu:"hu-hu",id:"id-id",it:"it-it",ja:"ja-jp",kk:"kk-kz",ko:"ko-kr",lo:"lo-la",lt:"lt-lt",lv:"lv-lv",ms:"ms-my",nb:"nb-no",nl:"nl-nl",nn:"nn-no",pl:"pl-pl",pt:"pt-br",ro:"ro-ro",ru:"ru-ru",sk:"sk-sk",sl:"sl-si",sr:"sr-cyrl-cs",sv:"sv-se",th:"th-th",tr:"tr-tr",uk:"uk-ua",ur:"ur-pk",vi:"vi-vn",zh:"zh-cn"};OSF.getSupportedLocale=function(a,c){if(c===void 0)c="en-us";if(!a)return c;var b;a=a.toLowerCase();if(a in OSF.SupportedLocales)b=a;else{var d=a.split("-",1);if(d&&d.length>0)b=OSF.AssociatedLocales[d[0]]}if(!b)b=c;return b};var ScriptLoading;(function(e){var a=false,b=function(){function b(g,e,d,f,c){var b=this;b.url=g;b.isReady=e;b.hasStarted=d;b.timer=f;b.hasError=a;b.pendingCallbacks=[];b.pendingCallbacks.push(c)}return b}(),d=function(){function a(c,b,a){this.scriptId=c;this.startTime=b;this.msResponseTime=a}return a}(),c=function(){var e=true,c=null;function f(b){var a=this;if(b===void 0)b={OfficeJS:"office.js",OfficeDebugJS:"office.debug.js"};a.constantNames=b;a.defaultScriptLoadingTimeout=1e4;a.loadedScriptByIds={};a.scriptTelemetryBuffer=[];a.osfControlAppCorrelationId="";a.basePath=c;a.getUseAssociatedActionsOnly=c}f.prototype.isScriptLoading=function(a){return !!(this.loadedScriptByIds[a]&&this.loadedScriptByIds[a].hasStarted)};f.prototype.getOfficeJsBasePath=function(){var b=this;if(b.basePath)return b.basePath;else{var h=function(b,c){var d,a,e;e=b.toLowerCase();a=e.indexOf(c);if(a>=0&&a===b.length-c.length&&(a===0||b.charAt(a-1)==="/"||b.charAt(a-1)==="\\"))d=b.substring(0,a);else if(a>=0&&a<b.length-c.length&&b.charAt(a+c.length)==="?"&&(a===0||b.charAt(a-1)==="/"||b.charAt(a-1)==="\\"))d=b.substring(0,a);return d},d=document.getElementsByTagName("script"),i=d.length,f=[b.constantNames.OfficeJS,b.constantNames.OfficeDebugJS],g=f.length;b.getUseAssociatedActionsOnly=a;for(var e,c=0;!b.basePath&&c<i;c++)if(d[c].src)for(e=0;!b.basePath&&e<g;e++){b.basePath=h(d[c].src,f[e]);if(b.basePath)try{var j=d[c].getAttribute("data-use-associated-actions-only");b.getUseAssociatedActionsOnly=j==="1"}catch(k){}}return b.basePath}};f.prototype.getUseAssociatedActionsOnlyDefined=function(){this.getOfficeJsBasePath();return this.getUseAssociatedActionsOnly};f.prototype.loadScript=function(e,d,c,a,b){this.loadScriptInternal(e,d,c,a,b)};f.prototype.loadScriptParallel=function(e,d,b){this.loadScriptInternal(e,d,c,a,b)};f.prototype.waitForFunction=function(g,d,h,i){var b=h,f,c=function(){b--;if(g()){d(e);return}else if(b>0){f=window.setTimeout(c,i);b--}else{window.clearTimeout(f);d(a)}};c()};f.prototype.waitForScripts=function(b,e){var f=this;if(this.invokeCallbackIfScriptsReady(b,e)==a)for(var c=0;c<b.length;c++){var g=b[c],d=this.loadedScriptByIds[g];d&&d.pendingCallbacks.push(function(){f.invokeCallbackIfScriptsReady(b,e)})}};f.prototype.logScriptLoading=function(c,a,b){a=Math.floor(a);if(OSF.AppTelemetry&&OSF.AppTelemetry.onScriptDone)if(OSF.AppTelemetry.onScriptDone.length==3)OSF.AppTelemetry.onScriptDone(c,a,b);else OSF.AppTelemetry.onScriptDone(c,a,b,this.osfControlAppCorrelationId);else{var e=new d(c,a,b);this.scriptTelemetryBuffer.push(e)}};f.prototype.setAppCorrelationId=function(a){this.osfControlAppCorrelationId=a};f.prototype.invokeCallbackIfScriptsReady=function(h,j){for(var g=a,f=0;f<h.length;f++){var i=h[f],d=this.loadedScriptByIds[i];if(!d){d=new b("",a,a,c,c);this.loadedScriptByIds[i]=d}if(d.isReady==a)return a;else if(d.hasError)g=e}j(!g);return e};f.prototype.getScriptEntryByUrl=function(d){for(var b in this.loadedScriptByIds){var a=this.loadedScriptByIds[b];if(this.loadedScriptByIds.hasOwnProperty(b)&&a.url===d)return a}return c};f.prototype.loadScriptInternal=function(h,g,i,n,k){var j=this;if(h){var q=j,r=window.document,d=g&&j.loadedScriptByIds[g]?j.loadedScriptByIds[g]:j.getScriptEntryByUrl(h);if(!d||d.hasError||d.url.toLowerCase()!=h.toLowerCase()){var f=r.createElement("script");f.type="text/javascript";if(g)f.id=g;if(!d){d=new b(h,a,a,c,c);j.loadedScriptByIds[g?g:h]=d}else{d.url=h;d.hasError=a;d.isReady=a}if(i)if(n)d.pendingCallbacks.unshift(i);else d.pendingCallbacks.push(i);var l=-1;if(window.performance&&window.performance.now)l=window.performance.now();var s=(new Date).getTime(),o=function(b){if(g){var a=(new Date).getTime()-s;if(!b)a=-a;q.logScriptLoading(g,l,a)}q.flushTelemetryBuffer()},m=function(){if(!OSF._OfficeAppFactory.getLoggingAllowed()&&typeof OSF.AppTelemetry!=="undefined")OSF.AppTelemetry.enableTelemetry=a;o(e);d.isReady=e;if(d.timer!=c){clearTimeout(d.timer);delete d.timer}for(var g=d.pendingCallbacks.length,f=0;f<g;f++){var b=d.pendingCallbacks.shift();if(b){var h=b(e);if(h===a)break}}},p=function(){o(a);d.hasError=e;d.isReady=e;if(d.timer!=c){clearTimeout(d.timer);delete d.timer}for(var g=d.pendingCallbacks.length,f=0;f<g;f++){var b=d.pendingCallbacks.shift();if(b){var h=b(a);if(h===a)break}}};if(f.readyState)f.onreadystatechange=function(){if(f.readyState=="loaded"||f.readyState=="complete"){f.onreadystatechange=c;m()}};else f.onload=m;f.onerror=p;k=k||j.defaultScriptLoadingTimeout;d.timer=setTimeout(p,k);d.hasStarted=e;f.setAttribute("crossOrigin","anonymous");f.src=h;r.getElementsByTagName("head")[0].appendChild(f)}else if(d.isReady)i(e);else if(n)d.pendingCallbacks.unshift(i);else d.pendingCallbacks.push(i)}};f.prototype.flushTelemetryBuffer=function(){var b=this;if(OSF.AppTelemetry&&OSF.AppTelemetry.onScriptDone){for(var c=0;c<b.scriptTelemetryBuffer.length;c++){var a=b.scriptTelemetryBuffer[c];if(OSF.AppTelemetry.onScriptDone.length==3)OSF.AppTelemetry.onScriptDone(a.scriptId,a.startTime,a.msResponseTime);else OSF.AppTelemetry.onScriptDone(a.scriptId,a.startTime,a.msResponseTime,b.osfControlAppCorrelationId)}b.scriptTelemetryBuffer=[]}};return f}();e.LoadScriptHelper=c})(ScriptLoading||(ScriptLoading={}));var OfficeExt;(function(a){var b;(function(a){var b=function(){function a(){var a=this;a.getDiagnostics=function(b){var a={host:this.getHost(),version:b||this.getDefaultVersion(),platform:this.getPlatform()};return a};a.platformRemappings={web:Microsoft.Office.WebExtension.PlatformType.OfficeOnline,winrt:Microsoft.Office.WebExtension.PlatformType.Universal,win32:Microsoft.Office.WebExtension.PlatformType.PC,mac:Microsoft.Office.WebExtension.PlatformType.Mac,ios:Microsoft.Office.WebExtension.PlatformType.iOS,android:Microsoft.Office.WebExtension.PlatformType.Android};a.camelCaseMappings={powerpoint:Microsoft.Office.WebExtension.HostType.PowerPoint,onenote:Microsoft.Office.WebExtension.HostType.OneNote};a.hostInfo=OSF._OfficeAppFactory.getHostInfo();a.getHost=a.getHost.bind(a);a.getPlatform=a.getPlatform.bind(a);a.getDiagnostics=a.getDiagnostics.bind(a)}a.prototype.capitalizeFirstLetter=function(a){if(a)return a[0].toUpperCase()+a.slice(1).toLowerCase();return a};a.getInstance=function(){if(a.hostObj===undefined)a.hostObj=new a;return a.hostObj};a.prototype.getPlatform=function(){var a=this;if(a.hostInfo.hostPlatform){var b=a.hostInfo.hostPlatform.toLowerCase();if(a.platformRemappings[b])return a.platformRemappings[b]}return null};a.prototype.getHost=function(){var a=this;if(a.hostInfo.hostType){var b=a.hostInfo.hostType.toLowerCase();if(a.camelCaseMappings[b])return a.camelCaseMappings[b];b=a.capitalizeFirstLetter(a.hostInfo.hostType);if(Microsoft.Office.WebExtension.HostType[b])return Microsoft.Office.WebExtension.HostType[b]}return null};a.prototype.getDefaultVersion=function(){if(this.getHost())return "16.0.0000.0000";return null};return a}();a.Host=b})(b=a.HostName||(a.HostName={}))})(OfficeExt||(OfficeExt={}));var Office;(function(d){var a=true,b="undefined",c="function",e;(function(d){var e;(function(d){function e(){return function(){var d=null,e="object";"use strict";function Q(a){return typeof a===c||typeof a===e&&a!==d}function y(a){return typeof a===c}function T(a){return typeof a===e&&a!==d}var z;if(!Array.isArray)z=function(a){return Object.prototype.toString.call(a)==="[object Array]"};else z=Array.isArray;var D=z,r=0,pb={}.toString,jb,w,l=function(a,b){q[r]=a;q[r+1]=b;r+=2;if(r===2)if(w)w(p);else u()};function cb(a){w=a}function lb(a){l=a}var X=typeof window!==b?window:undefined,C=X||{},G=C.MutationObserver||C.WebKitMutationObserver,mb=typeof process!==b&&{}.toString.call(process)==="[object process]",kb=typeof Uint8ClampedArray!==b&&typeof importScripts!==b&&typeof MessageChannel!==b;function eb(){var b=process.nextTick,a=process.versions.node.match(/^(?:(\d+)\.)?(?:(\d+)\.)?(\*|\d+)$/);if(Array.isArray(a)&&a[1]==="0"&&a[2]==="10")b=setImmediate;return function(){b(p)}}function O(){var a=new MessageChannel;a.port1.onmessage=p;return function(){a.port2.postMessage(0)}}function Y(){return function(){setTimeout(p,1)}}var q=new Array(1e3);function p(){for(var a=0;a<r;a+=2){var b=q[a],c=q[a+1];b(c);q[a]=undefined;q[a+1]=undefined}r=0}var u;if(mb)u=eb();else if(kb)u=O();else u=Y();function o(){}var k=void 0,m=1,j=2,s=new B;function K(){return new TypeError("You cannot resolve a promise with itself")}function L(){return new TypeError("A promises callback cannot return that same promise.")}function ab(b){try{return b.then}catch(a){s.error=a;return s}}function bb(e,d,b,c){try{e.call(d,b,c)}catch(a){return a}}function E(c,b,d){l(function(e){var c=false,g=bb(d,b,function(d){if(c)return;c=a;if(b!==d)n(e,d);else h(e,d)},function(b){if(c)return;c=a;f(e,b)},"Settle: "+(e._label||" unknown promise"));if(!c&&g){c=a;f(e,g)}},c)}function H(b,a){if(a._state===m)h(b,a._result);else if(a._state===j)f(b,a._result);else t(a,undefined,function(a){n(b,a)},function(a){f(b,a)})}function F(b,a){if(a.constructor===b.constructor)H(b,a);else{var c=ab(a);if(c===s)f(b,s.error);else if(c===undefined)h(b,a);else if(y(c))E(b,a,c);else h(b,a)}}function n(a,b){if(a===b)f(a,K());else if(Q(b))F(a,b);else h(a,b)}function J(a){a._onerror&&a._onerror(a._result);x(a)}function h(a,b){if(a._state!==k)return;a._result=b;a._state=m;a._subscribers.length!==0&&l(x,a)}function f(a,b){if(a._state!==k)return;a._state=j;a._result=b;l(J,a)}function t(c,g,e,f){var a=c._subscribers,b=a.length;c._onerror=d;a[b]=g;a[b+m]=e;a[b+j]=f;b===0&&c._state&&l(x,c)}function x(b){var a=b._subscribers,f=b._state;if(a.length===0)return;for(var e,d,g=b._result,c=0;c<a.length;c+=3){e=a[c];d=a[c+f];if(e)A(f,e,d,g);else d(g)}b._subscribers.length=0}function B(){this.error=d}var v=new B;function W(b,c){try{return b(c)}catch(a){v.error=a;return v}}function A(l,c,i,o){var g=y(i),b,q,e,p;if(g){b=W(i,o);if(b===v){p=a;q=b.error;b=d}else e=a;if(c===b){f(c,L());return}}else{b=o;e=a}if(c._state===k)if(g&&e)n(c,b);else if(p)f(c,q);else if(l===m)h(c,b);else l===j&&f(c,b)}function I(a,c){try{c(function(b){n(a,b)},function(b){f(a,b)})}catch(b){f(a,b)}}function i(c,b){var a=this;a._instanceConstructor=c;a.promise=new c(o);if(a._validateInput(b)){a._input=b;a.length=b.length;a._remaining=b.length;a._init();if(a.length===0)h(a.promise,a._result);else{a.length=a.length||0;a._enumerate();a._remaining===0&&h(a.promise,a._result)}}else f(a.promise,a._validationError())}i.prototype._validateInput=function(a){return D(a)};i.prototype._validationError=function(){return new Error("Array Methods must be provided an Array")};i.prototype._init=function(){this._result=new Array(this.length)};var Z=i;i.prototype._enumerate=function(){for(var a=this,d=a.length,c=a.promise,e=a._input,b=0;c._state===k&&b<d;b++)a._eachEntry(e[b],b)};i.prototype._eachEntry=function(a,c){var b=this,e=b._instanceConstructor;if(T(a))if(a.constructor===e&&a._state!==k){a._onerror=d;b._settledAt(a._state,c,a._result)}else b._willSettleAt(e.resolve(a),c);else{b._remaining--;b._result[c]=a}};i.prototype._settledAt=function(d,e,c){var a=this,b=a.promise;if(b._state===k){a._remaining--;if(d===j)f(b,c);else a._result[e]=c}a._remaining===0&&h(b,a._result)};i.prototype._willSettleAt=function(c,b){var a=this;t(c,undefined,function(c){a._settledAt(m,b,c)},function(c){a._settledAt(j,b,c)})};function ib(a){return (new Z(this,a)).promise}var V=ib;function db(b){var d=this,a=new d(o);if(!D(b)){f(a,new TypeError("You must pass an array to race."));return a}var h=b.length;function e(b){n(a,b)}function g(b){f(a,b)}for(var c=0;a._state===k&&c<h;c++)t(d.resolve(b[c]),undefined,e,g);return a}var U=db;function N(a){var b=this;if(a&&typeof a===e&&a.constructor===b)return a;var c=new b(o);n(c,a);return c}var M=N;function S(c){var
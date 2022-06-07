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
var OSF = OSF || {};
OSF.ConstantNames = {
    FileVersion: "0.0.0.0",
    OfficeJS: "custom-functions-runtime.js",
    OfficeDebugJS: "custom-functions-runtime.debug.js",
    HostFileScriptSuffix: "core",
    IsCustomFunctionsRuntime: true
};
var OSF = OSF || {};
OSF.HostSpecificFileVersionDefault = "16.00";
OSF.HostSpecificFileVersionMap = {
    "access": {
        "web": "16.00"
    },
    "agavito": {
        "winrt": "16.00"
    },
    "excel": {
        "ios": "16.00",
        "mac": "16.00",
        "web": "16.00",
        "win32": "16.01",
        "winrt": "16.00"
    },
    "onenote": {
        "android": "16.00",
        "web": "16.00",
        "win32": "16.00",
        "winrt": "16.00"
    },
    "outlook": {
        "ios": "16.00",
        "mac": "16.00",
        "web": "16.01",
        "win32": "16.02"
    },
    "powerpoint": {
        "ios": "16.00",
        "mac": "16.00",
        "web": "16.00",
        "win32": "16.01",
        "winrt": "16.00"
    },
    "project": {
        "win32": "16.00"
    },
    "sway": {
        "web": "16.00"
    },
    "word": {
        "ios": "16.00",
        "mac": "16.00",
        "web": "16.00",
        "win32": "16.01",
        "winrt": "16.00"
    },
    "visio": {
        "web": "16.00",
        "win32": "16.00"
    }
};
OSF.SupportedLocales = {
    "ar-sa": true,
    "bg-bg": true,
    "bn-in": true,
    "ca-es": true,
    "cs-cz": true,
    "da-dk": true,
    "de-de": true,
    "el-gr": true,
    "en-us": true,
    "es-es": true,
    "et-ee": true,
    "eu-es": true,
    "fa-ir": true,
    "fi-fi": true,
    "fr-fr": true,
    "gl-es": true,
    "he-il": true,
    "hi-in": true,
    "hr-hr": true,
    "hu-hu": true,
    "id-id": true,
    "it-it": true,
    "ja-jp": true,
    "kk-kz": true,
    "ko-kr": true,
    "lo-la": true,
    "lt-lt": true,
    "lv-lv": true,
    "ms-my": true,
    "nb-no": true,
    "nl-nl": true,
    "nn-no": true,
    "pl-pl": true,
    "pt-br": true,
    "pt-pt": true,
    "ro-ro": true,
    "ru-ru": true,
    "sk-sk": true,
    "sl-si": true,
    "sr-cyrl-cs": true,
    "sr-cyrl-rs": true,
    "sr-latn-cs": true,
    "sr-latn-rs": true,
    "sv-se": true,
    "th-th": true,
    "tr-tr": true,
    "uk-ua": true,
    "ur-pk": true,
    "vi-vn": true,
    "zh-cn": true,
    "zh-tw": true
};
OSF.AssociatedLocales = {
    ar: "ar-sa",
    bg: "bg-bg",
    bn: "bn-in",
    ca: "ca-es",
    cs: "cs-cz",
    da: "da-dk",
    de: "de-de",
    el: "el-gr",
    en: "en-us",
    es: "es-es",
    et: "et-ee",
    eu: "eu-es",
    fa: "fa-ir",
    fi: "fi-fi",
    fr: "fr-fr",
    gl: "gl-es",
    he: "he-il",
    hi: "hi-in",
    hr: "hr-hr",
    hu: "hu-hu",
    id: "id-id",
    it: "it-it",
    ja: "ja-jp",
    kk: "kk-kz",
    ko: "ko-kr",
    lo: "lo-la",
    lt: "lt-lt",
    lv: "lv-lv",
    ms: "ms-my",
    nb: "nb-no",
    nl: "nl-nl",
    nn: "nn-no",
    pl: "pl-pl",
    pt: "pt-br",
    ro: "ro-ro",
    ru: "ru-ru",
    sk: "sk-sk",
    sl: "sl-si",
    sr: "sr-cyrl-cs",
    sv: "sv-se",
    th: "th-th",
    tr: "tr-tr",
    uk: "uk-ua",
    ur: "ur-pk",
    vi: "vi-vn",
    zh: "zh-cn"
};
OSF.getSupportedLocale = function OSF$getSupportedLocale(locale, defaultLocale) {
    if (defaultLocale === void 0) { defaultLocale = "en-us"; }
    if (!locale) {
        return defaultLocale;
    }
    var supportedLocale;
    locale = locale.toLowerCase();
    if (locale in OSF.SupportedLocales) {
        supportedLocale = locale;
    }
    else {
        var localeParts = locale.split('-', 1);
        if (localeParts && localeParts.length > 0) {
            supportedLocale = OSF.AssociatedLocales[localeParts[0]];
        }
    }
    if (!supportedLocale) {
        supportedLocale = defaultLocale;
    }
    return supportedLocale;
};
var ScriptLoading;
(function (ScriptLoading) {
    var ScriptInfo = (function () {
        function ScriptInfo(url, isReady, hasStarted, timer, pendingCallback) {
            this.url = url;
            this.isReady = isReady;
            this.hasStarted = hasStarted;
            this.timer = timer;
            this.hasError = false;
            this.pendingCallbacks = [];
            this.pendingCallbacks.push(pendingCallback);
        }
        return ScriptInfo;
    }());
    var ScriptTelemetry = (function () {
        function ScriptTelemetry(scriptId, startTime, msResponseTime) {
            this.scriptId = scriptId;
            this.startTime = startTime;
            this.msResponseTime = msResponseTime;
        }
        return ScriptTelemetry;
    }());
    var LoadScriptHelper = (function () {
        function LoadScriptHelper(constantNames) {
            if (constantNames === void 0) { constantNames = {
                OfficeJS: "office.js",
                OfficeDebugJS: "office.debug.js"
            }; }
            this.constantNames = constantNames;
            this.defaultScriptLoadingTimeout = 10000;
            this.loadedScriptByIds = {};
            this.scriptTelemetryBuffer = [];
            this.osfControlAppCorrelationId = "";
            this.basePath = null;
            this.getUseAssociatedActionsOnly = null;
        }
        LoadScriptHelper.prototype.isScriptLoading = function (id) {
            return !!(this.loadedScriptByIds[id] && this.loadedScriptByIds[id].hasStarted);
        };
        LoadScriptHelper.prototype.getOfficeJsBasePath = function () {
            if (this.basePath) {
                return this.basePath;
            }
            else {
                var getScriptBase = function (scriptSrc, scriptNameToCheck) {
                    var scriptBase, indexOfJS, scriptSrcLowerCase;
                    scriptSrcLowerCase = scriptSrc.toLowerCase();
                    indexOfJS = scriptSrcLowerCase.indexOf(scriptNameToCheck);
                    if (indexOfJS >= 0 && indexOfJS === (scriptSrc.length - scriptNameToCheck.length) && (indexOfJS === 0 || scriptSrc.charAt(indexOfJS - 1) === '/' || scriptSrc.charAt(indexOfJS - 1) === '\\')) {
                        scriptBase = scriptSrc.substring(0, indexOfJS);
                    }
                    else if (indexOfJS >= 0
                        && indexOfJS < (scriptSrc.length - scriptNameToCheck.length)
                        && scriptSrc.charAt(indexOfJS + scriptNameToCheck.length) === '?'
                        && (indexOfJS === 0 || scriptSrc.charAt(indexOfJS - 1) === '/' || scriptSrc.charAt(indexOfJS - 1) === '\\')) {
                        scriptBase = scriptSrc.substring(0, indexOfJS);
                    }
                    return scriptBase;
                };
                var scripts = document.getElementsByTagName("script");
                var scriptsCount = scripts.length;
                var officeScripts = [this.constantNames.OfficeJS, this.constantNames.OfficeDebugJS];
                var officeScriptsCount = officeScripts.length;
                this.getUseAssociatedActionsOnly = false;
                var i, j;
                for (i = 0; !this.basePath && i < scriptsCount; i++) {
                    if (scripts[i].src) {
                        for (j = 0; !this.basePath && j < officeScriptsCount; j++) {
                            this.basePath = getScriptBase(scripts[i].src, officeScripts[j]);
                            if (this.basePath) {
                                try {
                                    var attr = scripts[i].getAttribute("data-use-associated-actions-only");
                                    this.getUseAssociatedActionsOnly = attr === "1";
                                }
                                catch (ex) {
                                }
                            }
                        }
                    }
                }
                return this.basePath;
            }
        };
        LoadScriptHelper.prototype.getUseAssociatedActionsOnlyDefined = function () {
            this.getOfficeJsBasePath();
            return this.getUseAssociatedActionsOnly;
        };
        LoadScriptHelper.prototype.loadScript = function (url, scriptId, callback, highPriority, timeoutInMs) {
            this.loadScriptInternal(url, scriptId, callback, highPriority, timeoutInMs);
        };
        LoadScriptHelper.prototype.loadScriptParallel = function (url, scriptId, timeoutInMs) {
            this.loadScriptInternal(url, scriptId, null, false, timeoutInMs);
        };
        LoadScriptHelper.prototype.waitForFunction = function (scriptLoadTest, callback, numberOfTries, delay) {
            var attemptsRemaining = numberOfTries;
            var timerId;
            var validateFunction = function () {
                attemptsRemaining--;
                if (scriptLoadTest()) {
                    callback(true);
                    return;
                }
                else if (attemptsRemaining > 0) {
                    timerId = window.setTimeout(validateFunction, delay);
                    attemptsRemaining--;
                }
                else {
                    window.clearTimeout(timerId);
                    callback(false);
                }
            };
            validateFunction();
        };
        LoadScriptHelper.prototype.waitForScripts = function (ids, callback) {
            var _this = this;
            if (this.invokeCallbackIfScriptsReady(ids, callback) == false) {
                for (var i = 0; i < ids.length; i++) {
                    var id = ids[i];
                    var loadedScriptEntry = this.loadedScriptByIds[id];
                    if (loadedScriptEntry) {
                        loadedScriptEntry.pendingCallbacks.push(function () {
                            _this.invokeCallbackIfScriptsReady(ids, callback);
                        });
                    }
                }
            }
        };
        LoadScriptHelper.prototype.logScriptLoading = function (scriptId, startTime, msResponseTime) {
            startTime = Math.floor(startTime);
            if (OSF.AppTelemetry && OSF.AppTelemetry.onScriptDone) {
                if (OSF.AppTelemetry.onScriptDone.length == 3) {
                    OSF.AppTelemetry.onScriptDone(scriptId, startTime, msResponseTime);
                }
                else {
                    OSF.AppTelemetry.onScriptDone(scriptId, startTime, msResponseTime, this.osfControlAppCorrelationId);
                }
            }
            else {
                var scriptTelemetry = new ScriptTelemetry(scriptId, startTime, msResponseTime);
                this.scriptTelemetryBuffer.push(scriptTelemetry);
            }
        };
        LoadScriptHelper.prototype.setAppCorrelationId = function (appCorrelationId) {
            this.osfControlAppCorrelationId = appCorrelationId;
        };
        LoadScriptHelper.prototype.invokeCallbackIfScriptsReady = function (ids, callback) {
            var hasError = false;
            for (var i = 0; i < ids.length; i++) {
                var id = ids[i];
                var loadedScriptEntry = this.loadedScriptByIds[id];
                if (!loadedScriptEntry) {
                    loadedScriptEntry = new ScriptInfo("", false, false, null, null);
                    this.loadedScriptByIds[id] = loadedScriptEntry;
                }
                if (loadedScriptEntry.isReady == false) {
                    return false;
                }
                else if (loadedScriptEntry.hasError) {
                    hasError = true;
                }
            }
            callback(!hasError);
            return true;
        };
        LoadScriptHelper.prototype.getScriptEntryByUrl = function (url) {
            for (var key in this.loadedScriptByIds) {
                var scriptEntry = this.loadedScriptByIds[key];
                if (this.loadedScriptByIds.hasOwnProperty(key) && scriptEntry.url === url) {
                    return scriptEntry;
                }
            }
            return null;
        };
        LoadScriptHelper.prototype.loadScriptInternal = function (url, scriptId, callback, highPriority, timeoutInMs) {
            if (url) {
                var self = this;
                var doc = window.document;
                var loadedScriptEntry = (scriptId && this.loadedScriptByIds[scriptId]) ? this.loadedScriptByIds[scriptId] : this.getScriptEntryByUrl(url);
                if (!loadedScriptEntry || loadedScriptEntry.hasError || loadedScriptEntry.url.toLowerCase() != url.toLowerCase()) {
                    var script = doc.createElement("script");
                    script.type = "text/javascript";
                    if (scriptId) {
                        script.id = scriptId;
                    }
                    if (!loadedScriptEntry) {
                        loadedScriptEntry = new ScriptInfo(url, false, false, null, null);
                        this.loadedScriptByIds[(scriptId ? scriptId : url)] = loadedScriptEntry;
                    }
                    else {
                        loadedScriptEntry.url = url;
                        loadedScriptEntry.hasError = false;
                        loadedScriptEntry.isReady = false;
                    }
                    if (callback) {
                        if (highPriority) {
                            loadedScriptEntry.pendingCallbacks.unshift(callback);
                        }
                        else {
                            loadedScriptEntry.pendingCallbacks.push(callback);
                        }
                    }
                    var timeFromPageInit = -1;
                    if (window.performance && window.performance.now) {
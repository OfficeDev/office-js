var OSFPerformance;
(function (OSFPerformance) {
    OSFPerformance.officeExecuteStart = 0;
    OSFPerformance.officeExecuteEnd = 0;
    OSFPerformance.hostInitializationStart = 0;
    OSFPerformance.hostInitializationEnd = 0;
    OSFPerformance.createOMEnd = 0;
    OSFPerformance.hostSpecificFileName = "";
    function now() {
        if (performance && performance.now) {
            return performance.now();
        }
        else {
            return 0;
        }
    }
    OSFPerformance.now = now;
})(OSFPerformance || (OSFPerformance = {}));
;
OSFPerformance.officeExecuteStart = OSFPerformance.now();



/* Office JavaScript API library */

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
    })();
    var ScriptTelemetry = (function () {
        function ScriptTelemetry(scriptId, startTime, msResponseTime) {
            this.scriptId = scriptId;
            this.startTime = startTime;
            this.msResponseTime = msResponseTime;
        }
        return ScriptTelemetry;
    })();
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
                var i, j;
                for (i = 0; !this.basePath && i < scriptsCount; i++) {
                    if (scripts[i].src) {
                        for (j = 0; !this.basePath && j < officeScriptsCount; j++) {
                            this.basePath = getScriptBase(scripts[i].src, officeScripts[j]);
                        }
                    }
                }
                return this.basePath;
            }
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
                        timeFromPageInit = window.performance.now();
                    }
                    var startTime = (new Date()).getTime();
                    var logTelemetry = function (succeeded) {
                        if (scriptId) {
                            var totalTime = (new Date()).getTime() - startTime;
                            if (!succeeded) {
                                totalTime = -totalTime;
                            }
                            self.logScriptLoading(scriptId, timeFromPageInit, totalTime);
                        }
                        self.flushTelemetryBuffer();
                    };
                    var onLoadCallback = function OSF_OUtil_loadScript$onLoadCallback() {
                        if (OSF._OfficeAppFactory.getHostInfo().hostType == "onenote" && (typeof OSF.AppTelemetry !== 'undefined') && (typeof OSF.AppTelemetry.enableTelemetry !== 'undefined')) {
                            OSF.AppTelemetry.enableTelemetry = false;
                        }
                        if (!OSF._OfficeAppFactory.getLoggingAllowed() && (typeof OSF.AppTelemetry !== 'undefined')) {
                            OSF.AppTelemetry.enableTelemetry = false;
                        }
                        logTelemetry(true);
                        loadedScriptEntry.isReady = true;
                        if (loadedScriptEntry.timer != null) {
                            clearTimeout(loadedScriptEntry.timer);
                            delete loadedScriptEntry.timer;
                        }
                        var pendingCallbackCount = loadedScriptEntry.pendingCallbacks.length;
                        for (var i = 0; i < pendingCallbackCount; i++) {
                            var currentCallback = loadedScriptEntry.pendingCallbacks.shift();
                            if (currentCallback) {
                                var result = currentCallback(true);
                                if (result === false) {
                                    break;
                                }
                            }
                        }
                    };
                    var onLoadError = function () {
                        logTelemetry(false);
                        loadedScriptEntry.hasError = true;
                        loadedScriptEntry.isReady = true;
                        if (loadedScriptEntry.timer != null) {
                            clearTimeout(loadedScriptEntry.timer);
                            delete loadedScriptEntry.timer;
                        }
                        var pendingCallbackCount = loadedScriptEntry.pendingCallbacks.length;
                        for (var i = 0; i < pendingCallbackCount; i++) {
                            var currentCallback = loadedScriptEntry.pendingCallbacks.shift();
                            if (currentCallback) {
                                var result = currentCallback(false);
                                if (result === false) {
                                    break;
                                }
                            }
                        }
                    };
                    if (script.readyState) {
                        script.onreadystatechange = function () {
                            if (script.readyState == "loaded" || script.readyState == "complete") {
                                script.onreadystatechange = null;
                                onLoadCallback();
                            }
                        };
                    }
                    else {
                        script.onload = onLoadCallback;
                    }
                    script.onerror = onLoadError;
                    timeoutInMs = timeoutInMs || this.defaultScriptLoadingTimeout;
                    loadedScriptEntry.timer = setTimeout(onLoadError, timeoutInMs);
                    loadedScriptEntry.hasStarted = true;
                    script.setAttribute("crossOrigin", "anonymous");
                    script.src = url;
                    doc.getElementsByTagName("head")[0].appendChild(script);
                }
                else if (loadedScriptEntry.isReady) {
                    callback(true);
                }
                else {
                    if (highPriority) {
                        loadedScriptEntry.pendingCallbacks.unshift(callback);
                    }
                    else {
                        loadedScriptEntry.pendingCallbacks.push(callback);
                    }
                }
            }
        };
        LoadScriptHelper.prototype.flushTelemetryBuffer = function () {
            if (OSF.AppTelemetry && OSF.AppTelemetry.onScriptDone) {
                for (var i = 0; i < this.scriptTelemetryBuffer.length; i++) {
                    var scriptTelemetry = this.scriptTelemetryBuffer[i];
                    if (OSF.AppTelemetry.onScriptDone.length == 3) {
                        OSF.AppTelemetry.onScriptDone(scriptTelemetry.scriptId, scriptTelemetry.startTime, scriptTelemetry.msResponseTime);
                    }
                    else {
                        OSF.AppTelemetry.onScriptDone(scriptTelemetry.scriptId, scriptTelemetry.startTime, scriptTelemetry.msResponseTime, this.osfControlAppCorrelationId);
                    }
                }
                this.scriptTelemetryBuffer = [];
            }
        };
        return LoadScriptHelper;
    })();
    ScriptLoading.LoadScriptHelper = LoadScriptHelper;
})(ScriptLoading || (ScriptLoading = {}));
var OfficeExt;
(function (OfficeExt) {
    var HostName;
    (function (HostName) {
        var Host = (function () {
            function Host() {
                this.getDiagnostics = function _getDiagnostics(version) {
                    var diagnostics = {
                        host: this.getHost(),
                        version: (version || this.getDefaultVersion()),
                        platform: this.getPlatform()
                    };
                    return diagnostics;
                };
                this.platformRemappings = {
                    web: Microsoft.Office.WebExtension.PlatformType.OfficeOnline,
                    winrt: Microsoft.Office.WebExtension.PlatformType.Universal,
                    win32: Microsoft.Office.WebExtension.PlatformType.PC,
                    mac: Microsoft.Office.WebExtension.PlatformType.Mac,
                    ios: Microsoft.Office.WebExtension.PlatformType.iOS,
                    android: Microsoft.Office.WebExtension.PlatformType.Android
                };
                this.camelCaseMappings = {
                    powerpoint: Microsoft.Office.WebExtension.HostType.PowerPoint,
                    onenote: Microsoft.Office.WebExtension.HostType.OneNote
                };
                this.hostInfo = OSF._OfficeAppFactory.getHostInfo();
                this.getHost = this.getHost.bind(this);
                this.getPlatform = this.getPlatform.bind(this);
                this.getDiagnostics = this.getDiagnostics.bind(this);
            }
            Host.prototype.capitalizeFirstLetter = function (input) {
                if (input) {
                    return (input[0].toUpperCase() + input.slice(1).toLowerCase());
                }
                return input;
            };
            Host.getInstance = function () {
                if (Host.hostObj === undefined) {
                    Host.hostObj = new Host();
                }
                return Host.hostObj;
            };
            Host.prototype.getPlatform = function (appNumber) {
                if (this.hostInfo.hostPlatform) {
                    var hostPlatform = this.hostInfo.hostPlatform.toLowerCase();
                    if (this.platformRemappings[hostPlatform]) {
                        return this.platformRemappings[hostPlatform];
                    }
                }
                return null;
            };
            Host.prototype.getHost = function (appNumber) {
                if (this.hostInfo.hostType) {
                    var hostType = this.hostInfo.hostType.toLowerCase();
                    if (this.camelCaseMappings[hostType]) {
                        return this.camelCaseMappings[hostType];
                    }
                    hostType = this.capitalizeFirstLetter(this.hostInfo.hostType);
                    if (Microsoft.Office.WebExtension.HostType[hostType]) {
                        return Microsoft.Office.WebExtension.HostType[hostType];
                    }
                }
                return null;
            };
            Host.prototype.getDefaultVersion = function () {
                if (this.getHost()) {
                    return "16.0.0000.0000";
                }
                return null;
            };
            return Host;
        })();
        HostName.Host = Host;
    })(HostName = OfficeExt.HostName || (OfficeExt.HostName = {}));
})(OfficeExt || (OfficeExt = {}));
var Office;
(function (Office) {
    var _Internal;
    (function (_Internal) {
        var PromiseImpl;
        (function (PromiseImpl) {
            function Init() {
                return (function () {
                    "use strict";
                    function lib$es6$promise$utils$$objectOrFunction(x) {
                        return typeof x === 'function' || (typeof x === 'object' && x !== null);
                    }
                    function lib$es6$promise$utils$$isFunction(x) {
                        return typeof x === 'function';
                    }
                    function lib$es6$promise$utils$$isMaybeThenable(x) {
                        return typeof x === 'object' && x !== null;
                    }
                    var lib$es6$promise$utils$$_isArray;
                    if (!Array.isArray) {
                        lib$es6$promise$utils$$_isArray = function (x) {
                            return Object.prototype.toString.call(x) === '[object Array]';
                        };
                    }
                    else {
                        lib$es6$promise$utils$$_isArray = Array.isArray;
                    }
                    var lib$es6$promise$utils$$isArray = lib$es6$promise$utils$$_isArray;
                    var lib$es6$promise$asap$$len = 0;
                    var lib$es6$promise$asap$$toString = {}.toString;
                    var lib$es6$promise$asap$$vertxNext;
                    var lib$es6$promise$asap$$customSchedulerFn;
                    var lib$es6$promise$asap$$asap = function asap(callback, arg) {
                        lib$es6$promise$asap$$queue[lib$es6$promise$asap$$len] = callback;
                        lib$es6$promise$asap$$queue[lib$es6$promise$asap$$len + 1] = arg;
                        lib$es6$promise$asap$$len += 2;
                        if (lib$es6$promise$asap$$len === 2) {
                            if (lib$es6$promise$asap$$customSchedulerFn) {
                                lib$es6$promise$asap$$customSchedulerFn(lib$es6$promise$asap$$flush);
                            }
                            else {
                                lib$es6$promise$asap$$scheduleFlush();
                            }
                        }
                    };
                    function lib$es6$promise$asap$$setScheduler(scheduleFn) {
                        lib$es6$promise$asap$$customSchedulerFn = scheduleFn;
                    }
                    function lib$es6$promise$asap$$setAsap(asapFn) {
                        lib$es6$promise$asap$$asap = asapFn;
                    }
                    var lib$es6$promise$asap$$browserWindow = (typeof window !== 'undefined') ? window : undefined;
                    var lib$es6$promise$asap$$browserGlobal = lib$es6$promise$asap$$browserWindow || {};
                    var lib$es6$promise$asap$$BrowserMutationObserver = lib$es6$promise$asap$$browserGlobal.MutationObserver || lib$es6$promise$asap$$browserGlobal.WebKitMutationObserver;
                    var lib$es6$promise$asap$$isNode = typeof process !== 'undefined' && {}.toString.call(process) === '[object process]';
                    var lib$es6$promise$asap$$isWorker = typeof Uint8ClampedArray !== 'undefined' &&
                        typeof importScripts !== 'undefined' &&
                        typeof MessageChannel !== 'undefined';
                    function lib$es6$promise$asap$$useNextTick() {
                        var nextTick = process.nextTick;
                        var version = process.versions.node.match(/^(?:(\d+)\.)?(?:(\d+)\.)?(\*|\d+)$/);
                        if (Array.isArray(version) && version[1] === '0' && version[2] === '10') {
                            nextTick = setImmediate;
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
                        var iterations = 0;
                        var observer = new lib$es6$promise$asap$$BrowserMutationObserver(lib$es6$promise$asap$$flush);
                        var node = document.createTextNode('');
                        observer.observe(node, { characterData: true });
                        return function () {
                            node.data = (iterations = ++iterations % 2);
                        };
                    }
                    function lib$es6$promise$asap$$useMessageChannel() {
                        var channel = new MessageChannel();
                        channel.port1.onmessage = lib$es6$promise$asap$$flush;
                        return function () {
                            channel.port2.postMessage(0);
                        };
                    }
                    function lib$es6$promise$asap$$useSetTimeout() {
                        return function () {
                            setTimeout(lib$es6$promise$asap$$flush, 1);
                        };
                    }
                    var lib$es6$promise$asap$$queue = new Array(1000);
                    function lib$es6$promise$asap$$flush() {
                        for (var i = 0; i < lib$es6$promise$asap$$len; i += 2) {
                            var callback = lib$es6$promise$asap$$queue[i];
                            var arg = lib$es6$promise$asap$$queue[i + 1];
                            callback(arg);
                            lib$es6$promise$asap$$queue[i] = undefined;
                            lib$es6$promise$asap$$queue[i + 1] = undefined;
                        }
                        lib$es6$promise$asap$$len = 0;
                    }
                    var lib$es6$promise$asap$$scheduleFlush;
                    if (lib$es6$promise$asap$$isNode) {
                        lib$es6$promise$asap$$scheduleFlush = lib$es6$promise$asap$$useNextTick();
                    }
                    else if (lib$es6$promise$asap$$isWorker) {
                        lib$es6$promise$asap$$scheduleFlush = lib$es6$promise$asap$$useMessageChannel();
                    }
                    else {
                        lib$es6$promise$asap$$scheduleFlush = lib$es6$promise$asap$$useSetTimeout();
                    }
                    function lib$es6$promise$$internal$$noop() { }
                    var lib$es6$promise$$internal$$PENDING = void 0;
                    var lib$es6$promise$$internal$$FULFILLED = 1;
                    var lib$es6$promise$$internal$$REJECTED = 2;
                    var lib$es6$promise$$internal$$GET_THEN_ERROR = new lib$es6$promise$$internal$$ErrorObject();
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
                            lib$es6$promise$$internal$$GET_THEN_ERROR.error = error;
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
                            var sealed = false;
                            var error = lib$es6$promise$$internal$$tryThen(then, thenable, function (value) {
                                if (sealed) {
                                    return;
                                }
                                sealed = true;
                                if (thenable !== value) {
                                    lib$es6$promise$$internal$$resolve(promise, value);
                                }
                                else {
                                    lib$es6$promise$$internal$$fulfill(promise, value);
                                }
                            }, function (reason) {
                                if (sealed) {
                                    return;
                                }
                                sealed = true;
                                lib$es6$promise$$internal$$reject(promise, reason);
                            }, 'Settle: ' + (promise._label || ' unknown promise'));
                            if (!sealed && error) {
                                sealed = true;
                                lib$es6$promise$$internal$$reject(promise, error);
                            }
                        }, promise);
                    }
                    function lib$es6$promise$$internal$$handleOwnThenable(promise, thenable) {
                        if (thenable._state === lib$es6$promise$$internal$$FULFILLED) {
                            lib$es6$promise$$internal$$fulfill(promise, thenable._result);
                        }
                        else if (thenable._state === lib$es6$promise$$internal$$REJECTED) {
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
                        if (maybeThenable.constructor === promise.constructor) {
                            lib$es6$promise$$internal$$handleOwnThenable(promise, maybeThenable);
                        }
                        else {
                            var then = lib$es6$promise$$internal$$getThen(maybeThenable);
                            if (then === lib$es6$promise$$internal$$GET_THEN_ERROR) {
                                lib$es6$promise$$internal$$reject(promise, lib$es6$promise$$internal$$GET_THEN_ERROR.error);
                            }
                            else if (then === undefined) {
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
                        if (promise === value) {
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
                        if (promise._state !== lib$es6$promise$$internal$$PENDING) {
                            return;
                        }
                        promise._result = value;
                        promise._state = lib$es6$promise$$internal$$FULFILLED;
                        if (promise._subscribers.length !== 0) {
                            lib$es6$promise$asap$$asap(lib$es6$promise$$internal$$publish, promise);
                        }
                    }
                    function lib$es6$promise$$internal$$reject(promise, reason) {
                        if (promise._state !== lib$es6$promise$$internal$$PENDING) {
                            return;
                        }
                        promise._state = lib$es6$promise$$internal$$REJECTED;
                        promise._result = reason;
                        lib$es6$promise$asap$$asap(lib$es6$promise$$internal$$publishRejection, promise);
                    }
                    function lib$es6$promise$$internal$$subscribe(parent, child, onFulfillment, onRejection) {
                        var subscribers = parent._subscribers;
                        var length = subscribers.length;
                        parent._onerror = null;
                        subscribers[length] = child;
                        subscribers[length + lib$es6$promise$$internal$$FULFILLED] = onFulfillment;
                        subscribers[length + lib$es6$promise$$internal$$REJECTED] = onRejection;
                        if (length === 0 && parent._state) {
                            lib$es6$promise$asap$$asap(lib$es6$promise$$internal$$publish, parent);
                        }
                    }
                    function lib$es6$promise$$internal$$publish(promise) {
                        var subscribers = promise._subscribers;
                        var settled = promise._state;
                        if (subscribers.length === 0) {
                            return;
                        }
                        var child, callback, detail = promise._result;
                        for (var i = 0; i < subscribers.length; i += 3) {
                            child = subscribers[i];
                            callback = subscribers[i + settled];
                            if (child) {
                                lib$es6$promise$$internal$$invokeCallback(settled, child, callback, detail);
                            }
                            else {
                                callback(detail);
                            }
                        }
                        promise._subscribers.length = 0;
                    }
                    function lib$es6$promise$$internal$$ErrorObject() {
                        this.error = null;
                    }
                    var lib$es6$promise$$internal$$TRY_CATCH_ERROR = new lib$es6$promise$$internal$$ErrorObject();
                    function lib$es6$promise$$internal$$tryCatch(callback, detail) {
                        try {
                            return callback(detail);
                        }
                        catch (e) {
                            lib$es6$promise$$internal$$TRY_CATCH_ERROR.error = e;
                            return lib$es6$promise$$internal$$TRY_CATCH_ERROR;
                        }
                    }
                    function lib$es6$promise$$internal$$invokeCallback(settled, promise, callback, detail) {
                        var hasCallback = lib$es6$promise$utils$$isFunction(callback), value, error, succeeded, failed;
                        if (hasCallback) {
                            value = lib$es6$promise$$internal$$tryCatch(callback, detail);
                            if (value === lib$es6$promise$$internal$$TRY_CATCH_ERROR) {
                                failed = true;
                                error = value.error;
                                value = null;
                            }
                            else {
                                succeeded = true;
                            }
                            if (promise === value) {
                                lib$es6$promise$$internal$$reject(promise, lib$es6$promise$$internal$$cannotReturnOwn());
                                return;
                            }
                        }
                        else {
                            value = detail;
                            succeeded = true;
                        }
                        if (promise._state !== lib$es6$promise$$internal$$PENDING) {
                        }
                        else if (hasCallback && succeeded) {
                            lib$es6$promise$$internal$$resolve(promise, value);
                        }
                        else if (failed) {
                            lib$es6$promise$$internal$$reject(promise, error);
                        }
                        else if (settled === lib$es6$promise$$internal$$FULFILLED) {
                            lib$es6$promise$$internal$$fulfill(promise, value);
                        }
                        else if (settled === lib$es6$promise$$internal$$REJECTED) {
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
                        var enumerator = this;
                        enumerator._instanceConstructor = Constructor;
                        enumerator.promise = new Constructor(lib$es6$promise$$internal$$noop);
                        if (enumerator._validateInput(input)) {
                            enumerator._input = input;
                            enumerator.length = input.length;
                            enumerator._remaining = input.length;
                            enumerator._init();
                            if (enumerator.length === 0) {
                                lib$es6$promise$$internal$$fulfill(enumerator.promise, enumerator._result);
                            }
                            else {
                                enumerator.length = enumerator.length || 0;
                                enumerator._enumerate();
                                if (enumerator._remaining === 0) {
                                    lib$es6$promise$$internal$$fulfill(enumerator.promise, enumerator._result);
                                }
                            }
                        }
                        else {
                            lib$es6$promise$$internal$$reject(enumerator.promise, enumerator._validationError());
                        }
                    }
                    lib$es6$promise$enumerator$$Enumerator.prototype._validateInput = function (input) {
                        return lib$es6$promise$utils$$isArray(input);
                    };
                    lib$es6$promise$enumerator$$Enumerator.prototype._validationError = function () {
                        return new Error('Array Methods must be provided an Array');
                    };
                    lib$es6$promise$enumerator$$Enumerator.prototype._init = function () {
                        this._result = new Array(this.length);
                    };
                    var lib$es6$promise$enumerator$$default = lib$es6$promise$enumerator$$Enumerator;
                    lib$es6$promise$enumerator$$Enumerator.prototype._enumerate = function () {
                        var enumerator = this;
                        var length = enumerator.length;
                        var promise = enumerator.promise;
                        var input = enumerator._input;
                        for (var i = 0; promise._state === lib$es6$promise$$internal$$PENDING && i < length; i++) {
                            enumerator._eachEntry(input[i], i);
                        }
                    };
                    lib$es6$promise$enumerator$$Enumerator.prototype._eachEntry = function (entry, i) {
                        var enumerator = this;
                        var c = enumerator._instanceConstructor;
                        if (lib$es6$promise$utils$$isMaybeThenable(entry)) {
                            if (entry.constructor === c && entry._state !== lib$es6$promise$$internal$$PENDING) {
                                entry._onerror = null;
                                enumerator._settledAt(entry._state, i, entry._result);
                            }
                            else {
                                enumerator._willSettleAt(c.resolve(entry), i);
                            }
                        }
                        else {
                            enumerator._remaining--;
                            enumerator._result[i] = entry;
                        }
                    };
                    lib$es6$promise$enumerator$$Enumerator.prototype._settledAt = function (state, i, value) {
                        var enumerator = this;
                        var promise = enumerator.promise;
                        if (promise._state === lib$es6$promise$$internal$$PENDING) {
                            enumerator._remaining--;
                            if (state === lib$es6$promise$$internal$$REJECTED) {
                                lib$es6$promise$$internal$$reject(promise, value);
                            }
                            else {
                                enumerator._result[i] = value;
                            }
                        }
                        if (enumerator._remaining === 0) {
                            lib$es6$promise$$internal$$fulfill(promise, enumerator._result);
                        }
                    };
                    lib$es6$promise$enumerator$$Enumerator.prototype._willSettleAt = function (promise, i) {
                        var enumerator = this;
                        lib$es6$promise$$internal$$subscribe(promise, undefined, function (value) {
                            enumerator._settledAt(lib$es6$promise$$internal$$FULFILLED, i, value);
                        }, function (reason) {
                            enumerator._settledAt(lib$es6$promise$$internal$$REJECTED, i, reason);
                        });
                    };
                    function lib$es6$promise$promise$all$$all(entries) {
                        return new lib$es6$promise$enumerator$$default(this, entries).promise;
                    }
                    var lib$es6$promise$promise$all$$default = lib$es6$promise$promise$all$$all;
                    function lib$es6$promise$promise$race$$race(entries) {
                        var Constructor = this;
                        var promise = new Constructor(lib$es6$promise$$internal$$noop);
                        if (!lib$es6$promise$utils$$isArray(entries)) {
                            lib$es6$promise$$internal$$reject(promise, new TypeError('You must pass an array to race.'));
                            return promise;
                        }
                        var length = entries.length;
                        function onFulfillment(value) {
                            lib$es6$promise$$internal$$resolve(promise, value);
                        }
                        function onRejection(reason) {
                            lib$es6$promise$$internal$$reject(promise, reason);
                        }
                        for (var i = 0; promise._state === lib$es6$promise$$internal$$PENDING && i < length; i++) {
                            lib$es6$promise$$internal$$subscribe(Constructor.resolve(entries[i]), undefined, onFulfillment, onRejection);
                        }
                        return promise;
                    }
                    var lib$es6$promise$promise$race$$default = lib$es6$promise$promise$race$$race;
                    function lib$es6$promise$promise$resolve$$resolve(object) {
                        var Constructor = this;
                        if (object && typeof object === 'object' && object.constructor === Constructor) {
                            return object;
                        }
                        var promise = new Constructor(lib$es6$promise$$internal$$noop);
                        lib$es6$promise$$internal$$resolve(promise, object);
                        return promise;
                    }
                    var lib$es6$promise$promise$resolve$$default = lib$es6$promise$promise$resolve$$resolve;
                    function lib$es6$promise$promise$reject$$reject(reason) {
                        var Constructor = this;
                        var promise = new Constructor(lib$es6$promise$$internal$$noop);
                        lib$es6$promise$$internal$$reject(promise, reason);
                        return promise;
                    }
                    var lib$es6$promise$promise$reject$$default = lib$es6$promise$promise$reject$$reject;
                    var lib$es6$promise$promise$$counter = 0;
                    function lib$es6$promise$promise$$needsResolver() {
                        throw new TypeError('You must pass a resolver function as the first argument to the promise constructor');
                    }
                    function lib$es6$promise$promise$$needsNew() {
                        throw new TypeError("Failed to construct 'Promise': Please use the 'new' operator, this object constructor cannot be called as a function.");
                    }
                    var lib$es6$promise$promise$$default = lib$es6$promise$promise$$Promise;
                    function lib$es6$promise$promise$$Promise(resolver) {
                        this._id = lib$es6$promise$promise$$counter++;
                        this._state = undefined;
                        this._result = undefined;
                        this._subscribers = [];
                        if (lib$es6$promise$$internal$$noop !== resolver) {
                            if (!lib$es6$promise$utils$$isFunction(resolver)) {
                                lib$es6$promise$promise$$needsResolver();
                            }
                            if (!(this instanceof lib$es6$promise$promise$$Promise)) {
                                lib$es6$promise$promise$$needsNew();
                            }
                            lib$es6$promise$$internal$$initializePromise(this, resolver);
                        }
                    }
                    lib$es6$promise$promise$$Promise.all = lib$es6$promise$promise$all$$default;
                    lib$es6$promise$promise$$Promise.race = lib$es6$promise$promise$race$$default;
                    lib$es6$promise$promise$$Promise.resolve = lib$es6$promise$promise$resolve$$default;
                    lib$es6$promise$promise$$Promise.reject = lib$es6$promise$promise$reject$$default;
                    lib$es6$promise$promise$$Promise._setScheduler = lib$es6$promise$asap$$setScheduler;
                    lib$es6$promise$promise$$Promise._setAsap = lib$es6$promise$asap$$setAsap;
                    lib$es6$promise$promise$$Promise._asap = lib$es6$promise$asap$$asap;
                    lib$es6$promise$promise$$Promise.prototype = {
                        constructor: lib$es6$promise$promise$$Promise,
                        then: function (onFulfillment, onRejection) {
                            var parent = this;
                            var state = parent._state;
                            if (state === lib$es6$promise$$internal$$FULFILLED && !onFulfillment || state === lib$es6$promise$$internal$$REJECTED && !onRejection) {
                                return this;
                            }
                            var child = new this.constructor(lib$es6$promise$$internal$$noop);
                            var result = parent._result;
                            if (state) {
                                var callback = arguments[state - 1];
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
            PromiseImpl.Init = Init;
        })(PromiseImpl = _Internal.PromiseImpl || (_Internal.PromiseImpl = {}));
    })(_Internal = Office._Internal || (Office._Internal = {}));
    var _Internal;
    (function (_Internal) {
        function isEdgeLessThan14() {
            var userAgent = window.navigator.userAgent;
            var versionIdx = userAgent.indexOf("Edge/");
            if (versionIdx >= 0) {
                userAgent = userAgent.substring(versionIdx + 5, userAgent.length);
                if (userAgent < "14.14393")
                    return true;
                else
                    return false;
            }
            return false;
        }
        function determinePromise() {
            if (typeof (window) === "undefined" && typeof (Promise) === "function") {
                return Promise;
            }
            if (typeof (window) !== "undefined" && window.Promise) {
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
        _Internal.OfficePromise = determinePromise();
    })(_Internal = Office._Internal || (Office._Internal = {}));
    var OfficePromise = _Internal.OfficePromise;
    Office.Promise = OfficePromise;
})(Office || (Office = {}));
var OTel;
(function (OTel) {
    var CDN_PATH_OTELJS_AGAVE = 'telemetry/oteljs_agave.js';
    var OTelLogger = (function () {
        function OTelLogger() {
        }
        OTelLogger.loaded = function () {
            return !(OTelLogger.logger === undefined);
        };
        OTelLogger.getOtelSinkCDNLocation = function () {
            return (OSF._OfficeAppFactory.getLoadScriptHelper().getOfficeJsBasePath() + CDN_PATH_OTELJS_AGAVE);
        };
        OTelLogger.getMapName = function (map, name) {
            if (name !== undefined && map.hasOwnProperty(name)) {
                return map[name];
            }
            return name;
        };
        OTelLogger.getHost = function () {
            var host = OSF._OfficeAppFactory.getHostInfo()["hostType"];
            var map = {
                "excel": "Excel",
                "onenote": "OneNote",
                "outlook": "Outlook",
                "powerpoint": "PowerPoint",
                "project": "Project",
                "visio": "Visio",
                "word": "Word"
            };
            var mappedName = OTelLogger.getMapName(map, host);
            return mappedName;
        };
        OTelLogger.getFlavor = function () {
            var flavor = OSF._OfficeAppFactory.getHostInfo()["hostPlatform"];
            var map = {
                "android": "Android",
                "ios": "iOS",
                "mac": "Mac",
                "universal": "Universal",
                "web": "Web",
                "win32": "Win32"
            };
            var mappedName = OTelLogger.getMapName(map, flavor);
            return mappedName;
        };
        OTelLogger.ensureValue = function (value, alternative) {
            if (!value) {
                return alternative;
            }
            return value;
        };
        OTelLogger.create = function (info) {
            var contract = {
                id: info.appId,
                assetId: info.assetId,
                officeJsVersion: info.officeJSVersion,
                hostJsVersion: info.hostJSVersion,
                browserToken: info.clientId,
                instanceId: info.appInstanceId,
                name: info.name,
                sessionId: info.sessionId
            };
            var fields = oteljs.Contracts.Office.System.SDX.getFields("SDX", contract);
            var host = OTelLogger.getHost();
            var flavor = OTelLogger.getFlavor();
            var version = (flavor === "Web" && info.hostVersion.slice(0, 2) === "0.") ? "16.0.0.0" : info.hostVersion;
            var context = {
                'App.Name': host,
                'App.Platform': flavor,
                'App.Version': version,
                'Session.Id': OTelLogger.ensureValue(info.correlationId, "00000000-0000-0000-0000-000000000000")
            };
            var sink = oteljs_agave.AgaveSink.createInstance(context);
            var namespace = "Office.Extensibility.OfficeJs";
            var ariaTenantToken = 'db334b301e7b474db5e0f02f07c51a47-a1b5bc36-1bbe-482f-a64a-c2d9cb606706-7439';
            var nexusTenantToken = 1755;
            var logger = new oteljs.TelemetryLogger(undefined, fields);
            logger.addSink(sink);
            logger.setTenantToken(namespace, ariaTenantToken, nexusTenantToken);
            return logger;
        };
        OTelLogger.initialize = function (info) {
            if (!OTelLogger.Enabled) {
                OTelLogger.promises = [];
                return;
            }
            var timeoutAfterOneSecond = 1000;
            var afterOnReady = function () {
                if ((typeof oteljs === "undefined") || (typeof oteljs_agave === "undefined")) {
                    return;
                }
                if (!OTelLogger.loaded()) {
                    OTelLogger.logger = OTelLogger.create(info);
                }
                if (OTelLogger.loaded()) {
                    OTelLogger.promises.forEach(function (resolve) {
                        resolve();
                    });
                }
            };
            var afterLoadOtelSink = function () {
                Microsoft.Office.WebExtension.onReadyInternal().then(function () { return afterOnReady(); });
            };
            OSF.OUtil.loadScript(OTelLogger.getOtelSinkCDNLocation(), afterLoadOtelSink, timeoutAfterOneSecond);
        };
        OTelLogger.sendTelemetryEvent = function (telemetryEvent) {
            OTelLogger.onTelemetryLoaded(function () {
                try {
                    OTelLogger.logger.sendTelemetryEvent(telemetryEvent);
                }
                catch (e) {
                }
            });
        };
        OTelLogger.onTelemetryLoaded = function (resolve) {
            if (!OTelLogger.Enabled) {
                return;
            }
            if (OTelLogger.loaded()) {
                resolve();
            }
            else {
                OTelLogger.promises.push(resolve);
            }
        };
        OTelLogger.promises = [];
        OTelLogger.Enabled = true;
        return OTelLogger;
    })();
    OTel.OTelLogger = OTelLogger;
})(OTel || (OTel = {}));
var OfficeExt;
(function (OfficeExt) {
    var Association = (function () {
        function Association() {
            this.m_mappings = {};
            this.m_onchangeHandlers = [];
        }
        Association.prototype.associate = function (arg1, arg2) {
            function consoleWarn(message) {
                if (typeof console !== 'undefined' && console.warn) {
                    console.warn(message);
                }
            }
            if (arguments.length == 1 && typeof arguments[0] === 'object' && arguments[0]) {
                var mappings = arguments[0];
                for (var key in mappings) {
                    this.associate(key, mappings[key]);
                }
            }
            else if (arguments.length == 2) {
                var name_1 = arguments[0];
                var func = arguments[1];
                if (typeof name_1 !== 'string') {
                    consoleWarn('[InvalidArg] Function=associate');
                    return;
                }
                if (typeof func !== 'function') {
                    consoleWarn('[InvalidArg] Function=associate');
                    return;
                }
                var nameUpperCase = name_1.toUpperCase();
                if (this.m_mappings[nameUpperCase]) {
                    consoleWarn('[DuplicatedName] Function=' + name_1);
                }
                this.m_mappings[nameUpperCase] = func;
                for (var i = 0; i < this.m_onchangeHandlers.length; i++) {
                    this.m_onchangeHandlers[i]();
                }
            }
            else {
                consoleWarn('[InvalidArg] Function=associate');
            }
        };
        Association.prototype.onchange = function (handler) {
            if (handler) {
                this.m_onchangeHandlers.push(handler);
            }
        };
        Object.defineProperty(Association.prototype, "mappings", {
            get: function () {
                return this.m_mappings;
            },
            enumerable: true,
            configurable: true
        });
        return Association;
    })();
    OfficeExt.Association = Association;
})(OfficeExt || (OfficeExt = {}));
var CustomFunctionMappings = window.CustomFunctionMappings || {};
var CustomFunctions;
(function (CustomFunctions) {
    function delayInitialization() {
        CustomFunctionMappings['__delay__'] = true;
    }
    CustomFunctions.delayInitialization = delayInitialization;
    ;
    CustomFunctions._association = new OfficeExt.Association();
    function associate() {
        CustomFunctions._association.associate.apply(CustomFunctions._association, arguments);
        delete CustomFunctionMappings['__delay__'];
    }
    CustomFunctions.associate = associate;
    ;
})(CustomFunctions || (CustomFunctions = {}));
(function () {
    var previousConstantNames = OSF.ConstantNames || {};
    OSF.ConstantNames = {
        FileVersion: "0.0.0.0",
        OfficeJS: "office.js",
        OfficeDebugJS: "office.debug.js",
        DefaultLocale: "en-us",
        LocaleStringLoadingTimeout: 5000,
        MicrosoftAjaxId: "MSAJAX",
        OfficeStringsId: "OFFICESTRINGS",
        OfficeJsId: "OFFICEJS",
        HostFileId: "HOST",
        O15MappingId: "O15Mapping",
        OfficeStringJS: "office_strings.debug.js",
        O15InitHelper: "o15apptofilemappingtable.debug.js",
        SupportedLocales: OSF.SupportedLocales,
        AssociatedLocales: OSF.AssociatedLocales
    };
    for (var key in previousConstantNames) {
        OSF.ConstantNames[key] = previousConstantNames[key];
    }
})();
OSF.InitializationHelper = function OSF_InitializationHelper(hostInfo, webAppState, context, settings, hostFacade) {
    this._hostInfo = hostInfo;
    this._webAppState = webAppState;
    this._context = context;
    this._settings = settings;
    this._hostFacade = hostFacade;
};
OSF.InitializationHelper.prototype.saveAndSetDialogInfo = function OSF_InitializationHelper$saveAndSetDialogInfo(hostInfoValue) {
};
OSF.InitializationHelper.prototype.getAppContext = function OSF_InitializationHelper$getAppContext(wnd, gotAppContext) {
};
OSF.InitializationHelper.prototype.setAgaveHostCommunication = function OSF_InitializationHelper$setAgaveHostCommunication() {
};
OSF.InitializationHelper.prototype.prepareRightBeforeWebExtensionInitialize = function OSF_InitializationHelper$prepareRightBeforeWebExtensionInitialize(appContext) {
};
OSF.InitializationHelper.prototype.loadAppSpecificScriptAndCreateOM = function OSF_InitializationHelper$loadAppSpecificScriptAndCreateOM(appContext, appReady, basePath) {
};
OSF.InitializationHelper.prototype.prepareRightAfterWebExtensionInitialize = function OSF_InitializationHelper$prepareRightAfterWebExtensionInitialize() {
};
OSF.HostInfoFlags = {
    SharedApp: 1,
    CustomFunction: 2
};
OSF._OfficeAppFactory = (function OSF__OfficeAppFactory() {
    var _setNamespace = function OSF_OUtil$_setNamespace(name, parent) {
        if (parent && name && !parent[name]) {
            parent[name] = {};
        }
    };
    _setNamespace("Office", window);
    _setNamespace("Microsoft", window);
    _setNamespace("Office", Microsoft);
    _setNamespace("WebExtension", Microsoft.Office);
    if (typeof (window.Office) === 'object') {
        for (var p in window.Office) {
            if (window.Office.hasOwnProperty(p)) {
                Microsoft.Office.WebExtension[p] = window.Office[p];
            }
        }
    }
    window.Office = Microsoft.Office.WebExtension;
    Microsoft.Office.WebExtension.PlatformType = {
        PC: "PC",
        OfficeOnline: "OfficeOnline",
        Mac: "Mac",
        iOS: "iOS",
        Android: "Android",
        Universal: "Universal"
    };
    Microsoft.Office.WebExtension.HostType = {
        Word: "Word",
        Excel: "Excel",
        PowerPoint: "PowerPoint",
        Outlook: "Outlook",
        OneNote: "OneNote",
        Project: "Project",
        Access: "Access",
        Visio: "Visio"
    };
    var _context = {};
    var _settings = {};
    var _hostFacade = {};
    var _WebAppState = { id: null, webAppUrl: null, conversationID: null, clientEndPoint: null, wnd: window.parent, focused: false };
    var _hostInfo = { isO15: true, isRichClient: true, hostType: "", hostPlatform: "", hostSpecificFileVersion: "", hostLocale: "", osfControlAppCorrelationId: "", isDialog: false, disableLogging: false, flags: 0 };
    var _isLoggingAllowed = true;
    var _initializationHelper = {};
    var _appInstanceId = null;
    var _isOfficeJsLoaded = false;
    var _officeOnReadyPendingResolves = [];
    var _isOfficeOnReadyCalled = false;
    var _officeOnReadyHostAndPlatformInfo = { host: null, platform: null };
    var _loadScriptHelper = new ScriptLoading.LoadScriptHelper({
        OfficeJS: OSF.ConstantNames.OfficeJS,
        OfficeDebugJS: OSF.ConstantNames.OfficeDebugJS
    });
    if (window.performance && window.performance.now) {
        _loadScriptHelper.logScriptLoading(OSF.ConstantNames.OfficeJsId, -1, window.performance.now());
    }
    var _windowLocationHash = window.location.hash;
    var _windowLocationSearch = window.location.search;
    var _windowName = window.name;
    var setOfficeJsAsLoadedAndDispatchPendingOnReadyCallbacks = function (_a) {
        var host = _a.host, platform = _a.platform;
        _isOfficeJsLoaded = true;
        _officeOnReadyHostAndPlatformInfo = { host: host, platform: platform };
        while (_officeOnReadyPendingResolves.length > 0) {
            _officeOnReadyPendingResolves.shift()(_officeOnReadyHostAndPlatformInfo);
        }
    };
    Microsoft.Office.WebExtension.FeatureGates = {};
    Microsoft.Office.WebExtension.sendTelemetryEvent = function Microsoft_Office_WebExtension_sendTelemetryEvent(telemetryEvent) {
        OTel.OTelLogger.sendTelemetryEvent(telemetryEvent);
    };
    Microsoft.Office.WebExtension.onReadyInternal = function Microsoft_Office_WebExtension_onReadyInternal(callback) {
        if (_isOfficeJsLoaded) {
            var host = _officeOnReadyHostAndPlatformInfo.host, platform = _officeOnReadyHostAndPlatformInfo.platform;
            if (callback) {
                var result = callback({ host: host, platform: platform });
                if (result && result.then && typeof result.then === "function") {
                    return result.then(function () { return Office.Promise.resolve({ host: host, platform: platform }); });
                }
            }
            return Office.Promise.resolve({ host: host, platform: platform });
        }
        if (callback) {
            return new Office.Promise(function (resolve) {
                _officeOnReadyPendingResolves.push(function (receivedHostAndPlatform) {
                    var result = callback(receivedHostAndPlatform);
                    if (result && result.then && typeof result.then === "function") {
                        return result.then(function () { return resolve(receivedHostAndPlatform); });
                    }
                    resolve(receivedHostAndPlatform);
                });
            });
        }
        return new Office.Promise(function (resolve) {
            _officeOnReadyPendingResolves.push(resolve);
        });
    };
    Microsoft.Office.WebExtension.onReady = function Microsoft_Office_WebExtension_onReady(callback) {
        _isOfficeOnReadyCalled = true;
        return Microsoft.Office.WebExtension.onReadyInternal(callback);
    };
    var getQueryStringValue = function OSF__OfficeAppFactory$getQueryStringValue(paramName) {
        var hostInfoValue;
        var searchString = window.location.search;
        if (searchString) {
            var hostInfoParts = searchString.split(paramName + "=");
            if (hostInfoParts.length > 1) {
                var hostInfoValueRestString = hostInfoParts[1];
                var separatorRegex = new RegExp("[&#]", "g");
                var hostInfoValueParts = hostInfoValueRestString.split(separatorRegex);
                if (hostInfoValueParts.length > 0) {
                    hostInfoValue = hostInfoValueParts[0];
                }
            }
        }
        return hostInfoValue;
    };
    var compareVersions = function _compareVersions(version1, version2) {
        var splitVersion1 = version1.split(".");
        var splitVersion2 = version2.split(".");
        var iter;
        for (iter in splitVersion1) {
            if (parseInt(splitVersion1[iter]) < parseInt(splitVersion2[iter])) {
                return false;
            }
            else if (parseInt(splitVersion1[iter]) > parseInt(splitVersion2[iter])) {
                return true;
            }
        }
        return false;
    };
    var shouldLoadOldMacJs = function _shouldLoadOldMacJs() {
        var versionToUseNewJS = "15.30.1128.0";
        var currentHostVersion = window.external.GetContext().GetHostFullVersion();
        return !!compareVersions(versionToUseNewJS, currentHostVersion);
    };
    var _retrieveLoggingAllowed = function OSF__OfficeAppFactory$_retrieveLoggingAllowed() {
        _isLoggingAllowed = true;
        try {
            if (_hostInfo.disableLogging) {
                _isLoggingAllowed = false;
                return;
            }
            window.external = window.external || {};
            if (typeof window.external.GetLoggingAllowed === 'undefined') {
                _isLoggingAllowed = true;
            }
            else {
                _isLoggingAllowed = window.external.GetLoggingAllowed();
            }
        }
        catch (Exception) {
        }
    };
    var _retrieveHostInfo = function OSF__OfficeAppFactory$_retrieveHostInfo() {
        var hostInfoParaName = "_host_Info";
        var hostInfoValue = getQueryStringValue(hostInfoParaName);
        if (!hostInfoValue) {
            try {
                var windowNameObj = JSON.parse(_windowName);
                hostInfoValue = windowNameObj ? windowNameObj["hostInfo"] : null;
            }
            catch (Exception) {
            }
        }
        if (!hostInfoValue) {
            try {
                window.external = window.external || {};
                if (typeof agaveHost !== "undefined" && agaveHost.GetHostInfo) {
                    window.external.GetHostInfo = function () {
                        return agaveHost.GetHostInfo();
                    };
                }
                var fallbackHostInfo = window.external.GetHostInfo();
                if (fallbackHostInfo == "isDialog") {
                    _hostInfo.isO15 = true;
                    _hostInfo.isDialog = true;
                }
                else if (fallbackHostInfo.toLowerCase().indexOf("mac") !== -1 && fallbackHostInfo.toLowerCase().indexOf("outlook") !== -1 && shouldLoadOldMacJs()) {
                    _hostInfo.isO15 = true;
                }
                else {
                    var hostInfoParts = fallbackHostInfo.split(hostInfoParaName + "=");
                    if (hostInfoParts.length > 1) {
                        hostInfoValue = hostInfoParts[1];
                    }
                    else {
                        hostInfoValue = fallbackHostInfo;
                    }
                }
            }
            catch (Exception) {
            }
        }
        var getSessionStorage = function OSF__OfficeAppFactory$_retrieveHostInfo$getSessionStorage() {
            var osfSessionStorage = null;
            try {
                if (window.sessionStorage) {
                    osfSessionStorage = window.sessionStorage;
                }
            }
            catch (ex) {
            }
            return osfSessionStorage;
        };
        var osfSessionStorage = getSessionStorage();
        if (!hostInfoValue && osfSessionStorage && osfSessionStorage.getItem("hostInfoValue")) {
            hostInfoValue = osfSessionStorage.getItem("hostInfoValue");
        }
        if (hostInfoValue) {
            hostInfoValue = decodeURIComponent(hostInfoValue);
            _hostInfo.isO15 = false;
            var items = hostInfoValue.split("$");
            if (typeof items[2] == "undefined") {
                items = hostInfoValue.split("|");
            }
            _hostInfo.hostType = (typeof items[0] == "undefined") ? "" : items[0].toLowerCase();
            _hostInfo.hostPlatform = (typeof items[1] == "undefined") ? "" : items[1].toLowerCase();
            ;
            _hostInfo.hostSpecificFileVersion = (typeof items[2] == "undefined") ? "" : items[2].toLowerCase();
            _hostInfo.hostLocale = (typeof items[3] == "undefined") ? "" : items[3].toLowerCase();
            _hostInfo.osfControlAppCorrelationId = (typeof items[4] == "undefined") ? "" : items[4];
            if (_hostInfo.osfControlAppCorrelationId == "telemetry") {
                _hostInfo.osfControlAppCorrelationId = "";
            }
            _hostInfo.isDialog = (((typeof items[5]) != "undefined") && items[5] == "isDialog") ? true : false;
            _hostInfo.disableLogging = (((typeof items[6]) != "undefined") && items[6] == "disableLogging") ? true : false;
            _hostInfo.flags = (((typeof items[7]) === "string") && items[7].length > 0) ? parseInt(items[7]) : 0;
            var hostSpecificFileVersionValue = parseFloat(_hostInfo.hostSpecificFileVersion);
            var fallbackVersion = OSF.HostSpecificFileVersionDefault;
            if (OSF.HostSpecificFileVersionMap[_hostInfo.hostType] && OSF.HostSpecificFileVersionMap[_hostInfo.hostType][_hostInfo.hostPlatform]) {
                fallbackVersion = OSF.HostSpecificFileVersionMap[_hostInfo.hostType][_hostInfo.hostPlatform];
            }
            if (hostSpecificFileVersionValue > parseFloat(fallbackVersion)) {
                _hostInfo.hostSpecificFileVersion = fallbackVersion;
            }
            if (osfSessionStorage) {
                try {
                    osfSessionStorage.setItem("hostInfoValue", hostInfoValue);
                }
                catch (e) {
                }
            }
        }
        else {
            _hostInfo.isO15 = true;
            _hostInfo.hostLocale = getQueryStringValue("locale");
        }
    };
    var getAppContextAsync = function OSF__OfficeAppFactory$getAppContextAsync(wnd, gotAppContext) {
        if (OSF.AppTelemetry && OSF.AppTelemetry.logAppCommonMessage) {
            OSF.AppTelemetry.logAppCommonMessage("getAppContextAsync starts");
        }
        _initializationHelper.getAppContext(wnd, gotAppContext);
    };
    var initialize = function OSF__OfficeAppFactory$initialize() {
        _retrieveHostInfo();
        _retrieveLoggingAllowed();
        if (_hostInfo.hostPlatform == "web" && _hostInfo.isDialog && window == window.top && window.opener == null) {
            window.open('', '_self', '');
            window.close();
        }
        if ((_hostInfo.flags & (OSF.HostInfoFlags.SharedApp | OSF.HostInfoFlags.CustomFunction)) !== 0) {
            if (typeof (window.Promise) === 'undefined') {
                window.Promise = window.Office.Promise;
            }
        }
        _loadScriptHelper.setAppCorrelationId(_hostInfo.osfControlAppCorrelationId);
        var basePath = _loadScriptHelper.getOfficeJsBasePath();
        var requiresMsAjax = false;
        if (!basePath)
            throw "Office Web Extension script library file name should be " + OSF.ConstantNames.OfficeJS + " or " + OSF.ConstantNames.OfficeDebugJS + ".";
        var isMicrosftAjaxLoaded = function OSF$isMicrosftAjaxLoaded() {
            if ((typeof (Sys) !== 'undefined' && typeof (Type) !== 'undefined' &&
                Sys.StringBuilder && typeof (Sys.StringBuilder) === "function" &&
                Type.registerNamespace && typeof (Type.registerNamespace) === "function" &&
                Type.registerClass && typeof (Type.registerClass) === "function") ||
                (typeof (OfficeExt) !== "undefined" && OfficeExt.MsAjaxError)) {
                return true;
            }
            else {
                return false;
            }
        };
        var officeStrings = null;
        var loadLocaleStrings = function OSF__OfficeAppFactory_initialize$loadLocaleStrings(appLocale) {
            var fallbackLocaleTried = false;
            var loadLocaleStringCallback = function OSF__OfficeAppFactory_initialize$loadLocaleStringCallback() {
                if (typeof Strings == 'undefined' || typeof Strings.OfficeOM == 'undefined') {
                    if (!fallbackLocaleTried) {
                        fallbackLocaleTried = true;
                        var fallbackLocaleStringFile = basePath + OSF.ConstantNames.DefaultLocale + "/" + OSF.ConstantNames.OfficeStringJS;
                        _loadScriptHelper.loadScript(fallbackLocaleStringFile, OSF.ConstantNames.OfficeStringsId, loadLocaleStringCallback, true, OSF.ConstantNames.LocaleStringLoadingTimeout);
                        return false;
                    }
                    else {
                        throw "Neither the locale, " + appLocale.toLowerCase() + ", provided by the host app nor the fallback locale " + OSF.ConstantNames.DefaultLocale + " are supported.";
                    }
                }
                else {
                    fallbackLocaleTried = false;
                    officeStrings = Strings.OfficeOM;
                }
            };
            if (!isMicrosftAjaxLoaded()) {
                window.Type = Function;
                Type.registerNamespace = function (ns) {
                    window[ns] = window[ns] || {};
                };
                Type.prototype.registerClass = function (cls) {
                    cls = {};
                };
            }
            var localeStringFile = basePath + OSF.getSupportedLocale(appLocale, OSF.ConstantNames.DefaultLocale) + "/" + OSF.ConstantNames.OfficeStringJS;
            _loadScriptHelper.loadScript(localeStringFile, OSF.ConstantNames.OfficeStringsId, loadLocaleStringCallback, true, OSF.ConstantNames.LocaleStringLoadingTimeout);
        };
        var onAppCodeAndMSAjaxReady = function OSF__OfficeAppFactory_initialize$onAppCodeAndMSAjaxReady(loadSuccess) {
            if (loadSuccess) {
                _initializationHelper = new OSF.InitializationHelper(_hostInfo, _WebAppState, _context, _settings, _hostFacade);
                if (_hostInfo.hostPlatform == "web" && _initializationHelper.saveAndSetDialogInfo) {
                    _initializationHelper.saveAndSetDialogInfo(getQueryStringValue("_host_Info"));
                }
                _initializationHelper.setAgaveHostCommunication();
                getAppContextAsync(_WebAppState.wnd, function (appContext) {
                    if (OSF.AppTelemetry && OSF.AppTelemetry.logAppCommonMessage) {
                        OSF.AppTelemetry.logAppCommonMessage("getAppContextAsync callback start");
                    }
                    _appInstanceId = appContext._appInstanceId;
                    if (appContext.get_featureGates) {
                        var featureGates = appContext.get_featureGates();
                        if (featureGates) {
                            Microsoft.Office.WebExtension.FeatureGates = featureGates;
                        }
                    }
                    var updateVersionInfo = function updateVersionInfo() {
                        var hostVersionItems = _hostInfo.hostSpecificFileVersion.split(".");
                        if (appContext.get_appMinorVersion) {
                            var isIOS = _hostInfo.hostPlatform == "ios";
                            if (!isIOS) {
                                if (isNaN(appContext.get_appMinorVersion())) {
                                    appContext._appMinorVersion = parseInt(hostVersionItems[1]);
                                }
                                else if (hostVersionItems.length > 1 && !isNaN(Number(hostVersionItems[1]))) {
                                    appContext._appMinorVersion = parseInt(hostVersionItems[1]);
                                }
                            }
                        }
                        if (_hostInfo.isDialog) {
                            appContext._isDialog = _hostInfo.isDialog;
                        }
                    };
                    updateVersionInfo();
                    var appReady = function appReady() {
                        _initializationHelper.prepareApiSurface && _initializationHelper.prepareApiSurface(appContext);
                        _loadScriptHelper.waitForFunction(function () { return (Microsoft.Office.WebExtension.initialize != undefined || _isOfficeOnReadyCalled); }, function (initializedDeclaredOrOfficeOnReadyCalled) {
                            if (initializedDeclaredOrOfficeOnReadyCalled) {
                                if (_initializationHelper.prepareApiSurface) {
                                    if (Microsoft.Office.WebExtension.initialize) {
                                        Microsoft.Office.WebExtension.initialize(_initializationHelper.getInitializationReason(appContext));
                                    }
                                }
                                else {
                                    if (!Microsoft.Office.WebExtension.initialize) {
                                        Microsoft.Office.WebExtension.initialize = function () { };
                                    }
                                    _initializationHelper.prepareRightBeforeWebExtensionInitialize(appContext);
                                }
                                _initializationHelper.prepareRightAfterWebExtensionInitialize && _initializationHelper.prepareRightAfterWebExtensionInitialize();
                                var appNumber = appContext.get_appName();
                                setOfficeJsAsLoadedAndDispatchPendingOnReadyCallbacks({
                                    host: OfficeExt.HostName.Host.getInstance().getHost(appNumber),
                                    platform: OfficeExt.HostName.Host.getInstance().getPlatform(appNumber)
                                });
                            }
                            else {
                                throw new Error("Office.js has not fully loaded. Your app must call \"Office.onReady()\" as part of it's loading sequence (or set the \"Office.initialize\" function). If your app has this functionality, try reloading this page.");
                            }
                        }, 400, 50);
                    };
                    if (!_loadScriptHelper.isScriptLoading(OSF.ConstantNames.OfficeStringsId)) {
                        loadLocaleStrings(appContext.get_appUILocale());
                    }
                    _loadScriptHelper.waitForScripts([OSF.ConstantNames.OfficeStringsId], function () {
                        if (officeStrings && !Strings.OfficeOM) {
                            Strings.OfficeOM = officeStrings;
                        }
                        _initializationHelper.loadAppSpecificScriptAndCreateOM(appContext, appReady, basePath);
                    });
                });
                if (_hostInfo.isO15) {
                    var wacXdmInfoIsMissing = (OSF.OUtil.parseXdmInfo() == null);
                    if (wacXdmInfoIsMissing) {
                        var isPlainBrowser = true;
                        if (window.external && typeof window.external.GetContext !== 'undefined') {
                            try {
                                window.external.GetContext();
                                isPlainBrowser = false;
                            }
                            catch (e) {
                            }
                        }
                        if (isPlainBrowser) {
                            setOfficeJsAsLoadedAndDispatchPendingOnReadyCallbacks({
                                host: null,
                                platform: null
                            });
                        }
                    }
                }
            }
            else {
                var errorMsg = "MicrosoftAjax.js is not loaded successfully.";
                if (OSF.AppTelemetry && OSF.AppTelemetry.logAppException) {
                    OSF.AppTelemetry.logAppException(errorMsg);
                }
                throw errorMsg;
            }
        };
        var onAppCodeReady = function OSF__OfficeAppFactory_initialize$onAppCodeReady() {
            if (OSF.AppTelemetry && OSF.AppTelemetry.setOsfControlAppCorrelationId) {
                OSF.AppTelemetry.setOsfControlAppCorrelationId(_hostInfo.osfControlAppCorrelationId);
            }
            if (_loadScriptHelper.isScriptLoading(OSF.ConstantNames.MicrosoftAjaxId)) {
                _loadScriptHelper.waitForScripts([OSF.ConstantNames.MicrosoftAjaxId], onAppCodeAndMSAjaxReady);
            }
            else {
                _loadScriptHelper.waitForFunction(isMicrosftAjaxLoaded, onAppCodeAndMSAjaxReady, 500, 100);
            }
        };
        if (_hostInfo.isO15) {
            _loadScriptHelper.loadScript(basePath + OSF.ConstantNames.O15InitHelper, OSF.ConstantNames.O15MappingId, onAppCodeReady);
        }
        else {
            var hostSpecificFileName = ([
                _hostInfo.hostType,
                _hostInfo.hostPlatform,
                _hostInfo.hostSpecificFileVersion,
                OSF.ConstantNames.HostFileScriptSuffix || null,
            ]
                .filter(function (part) { return part != null; })
                .join("-"))
                +
                    ".debug.js";
            _loadScriptHelper.loadScript(basePath + hostSpecificFileName.toLowerCase(), OSF.ConstantNames.HostFileId, onAppCodeReady);
            if (typeof OSFPerformance !== "undefined") {
                OSFPerformance.hostSpecificFileName = hostSpecificFileName;
            }
        }
        if (_hostInfo.hostLocale) {
            loadLocaleStrings(_hostInfo.hostLocale);
        }
        if (requiresMsAjax && !isMicrosftAjaxLoaded()) {
            var msAjaxCDNPath = (window.location.protocol.toLowerCase() === 'https:' ? 'https:' : 'http:') + '//ajax.aspnetcdn.com/ajax/3.5/MicrosoftAjax.js';
            _loadScriptHelper.loadScriptParallel(msAjaxCDNPath, OSF.ConstantNames.MicrosoftAjaxId);
        }
        window.confirm = function OSF__OfficeAppFactory_initialize$confirm(message) {
            throw new Error('Function window.confirm is not supported.');
        };
        window.alert = function OSF__OfficeAppFactory_initialize$alert(message) {
            throw new Error('Function window.alert is not supported.');
        };
        window.prompt = function OSF__OfficeAppFactory_initialize$prompt(message, defaultvalue) {
            throw new Error('Function window.prompt is not supported.');
        };
        var isOutlookAndroid = _hostInfo.hostType == "outlook" && _hostInfo.hostPlatform == "android";
        if (!isOutlookAndroid) {
            window.history.replaceState = null;
            window.history.pushState = null;
        }
    };
    initialize();
    return {
        getId: function OSF__OfficeAppFactory$getId() { return _WebAppState.id; },
        getClientEndPoint: function OSF__OfficeAppFactory$getClientEndPoint() { return _WebAppState.clientEndPoint; },
        getContext: function OSF__OfficeAppFactory$getContext() { return _context; },
        setContext: function OSF__OfficeAppFactory$setContext(context) { _context = context; },
        getHostInfo: function OSF_OfficeAppFactory$getHostInfo() { return _hostInfo; },
        getLoggingAllowed: function OSF_OfficeAppFactory$getLoggingAllowed() { return _isLoggingAllowed; },
        getHostFacade: function OSF__OfficeAppFactory$getHostFacade() { return _hostFacade; },
        setHostFacade: function setHostFacade(hostFacade) { _hostFacade = hostFacade; },
        getInitializationHelper: function OSF__OfficeAppFactory$getInitializationHelper() { return _initializationHelper; },
        getCachedSessionSettingsKey: function OSF__OfficeAppFactory$getCachedSessionSettingsKey() {
            return (_WebAppState.conversationID != null ? _WebAppState.conversationID : _appInstanceId) + "CachedSessionSettings";
        },
        getWebAppState: function OSF__OfficeAppFactory$getWebAppState() { return _WebAppState; },
        getWindowLocationHash: function OSF__OfficeAppFactory$getHash() { return _windowLocationHash; },
        getWindowLocationSearch: function OSF__OfficeAppFactory$getSearch() { return _windowLocationSearch; },
        getLoadScriptHelper: function OSF__OfficeAppFactory$getLoadScriptHelper() { return _loadScriptHelper; },
        getWindowName: function OSF__OfficeAppFactory$getWindowName() { return _windowName; }
    };
})();



var oteljs = function(modules) {
    var installedModules = {};
    function __webpack_require__(moduleId) {
        if (installedModules[moduleId]) {
            return installedModules[moduleId].exports;
        }
        var module = installedModules[moduleId] = {
            i: moduleId,
            l: false,
            exports: {}
        };
        modules[moduleId].call(module.exports, module, module.exports, __webpack_require__);
        module.l = true;
        return module.exports;
    }
    __webpack_require__.m = modules;
    __webpack_require__.c = installedModules;
    __webpack_require__.d = function(exports, name, getter) {
        if (!__webpack_require__.o(exports, name)) {
            Object.defineProperty(exports, name, {
                enumerable: true,
                get: getter
            });
        }
    };
    __webpack_require__.r = function(exports) {
        if (typeof Symbol !== "undefined" && Symbol.toStringTag) {
            Object.defineProperty(exports, Symbol.toStringTag, {
                value: "Module"
            });
        }
        Object.defineProperty(exports, "__esModule", {
            value: true
        });
    };
    __webpack_require__.t = function(value, mode) {
        if (mode & 1) value = __webpack_require__(value);
        if (mode & 8) return value;
        if (mode & 4 && typeof value === "object" && value && value.__esModule) return value;
        var ns = Object.create(null);
        __webpack_require__.r(ns);
        Object.defineProperty(ns, "default", {
            enumerable: true,
            value: value
        });
        if (mode & 2 && typeof value != "string") for (var key in value) __webpack_require__.d(ns, key, function(key) {
            return value[key];
        }.bind(null, key));
        return ns;
    };
    __webpack_require__.n = function(module) {
        var getter = module && module.__esModule ? function getDefault() {
            return module["default"];
        } : function getModuleExports() {
            return module;
        };
        __webpack_require__.d(getter, "a", getter);
        return getter;
    };
    __webpack_require__.o = function(object, property) {
        return Object.prototype.hasOwnProperty.call(object, property);
    };
    __webpack_require__.p = "";
    return __webpack_require__(__webpack_require__.s = 0);
}([ function(module, exports, __webpack_require__) {
    module.exports = __webpack_require__(1);
}, function(module, __webpack_exports__, __webpack_require__) {
    "use strict";
    __webpack_require__.r(__webpack_exports__);
    var DataFieldType;
    (function(DataFieldType) {
        DataFieldType[DataFieldType["String"] = 0] = "String";
        DataFieldType[DataFieldType["Boolean"] = 1] = "Boolean";
        DataFieldType[DataFieldType["Int64"] = 2] = "Int64";
        DataFieldType[DataFieldType["Double"] = 3] = "Double";
    })(DataFieldType || (DataFieldType = {}));
    function makeBooleanDataField(name, value) {
        return {
            name: name,
            dataType: DataFieldType.Boolean,
            value: value
        };
    }
    function makeInt64DataField(name, value) {
        return {
            name: name,
            dataType: DataFieldType.Int64,
            value: value
        };
    }
    function makeDoubleDataField(name, value) {
        return {
            name: name,
            dataType: DataFieldType.Double,
            value: value
        };
    }
    function makeStringDataField(name, value) {
        return {
            name: name,
            dataType: DataFieldType.String,
            value: value
        };
    }
    function getFieldsForContract(instanceName, contractName, contractFields) {
        var dataFields = contractFields.map(function(contractField) {
            return {
                name: instanceName + "." + contractField.name,
                value: contractField.value,
                dataType: contractField.dataType
            };
        });
        addContractField(dataFields, instanceName, contractName);
        return dataFields;
    }
    function addContractField(dataFields, instanceName, contractName) {
        dataFields.push(makeStringDataField("zC." + instanceName, contractName));
    }
    var officeeventschema_tml_Result;
    (function(Result) {
        var contractName = "Office.System.Result";
        function getFields(instanceName, contract) {
            var dataFields = [];
            dataFields.push(makeInt64DataField(instanceName + ".Code", contract.code));
            if (contract.type !== undefined) {
                dataFields.push(makeStringDataField(instanceName + ".Type", contract.type));
            }
            if (contract.tag !== undefined) {
                dataFields.push(makeInt64DataField(instanceName + ".Tag", contract.tag));
            }
            if (contract.isExpected !== undefined) {
                dataFields.push(makeBooleanDataField(instanceName + ".IsExpected", contract.isExpected));
            }
            addContractField(dataFields, instanceName, contractName);
            return dataFields;
        }
        Result.getFields = getFields;
    })(officeeventschema_tml_Result || (officeeventschema_tml_Result = {}));
    var officeeventschema_tml_Activity;
    (function(Activity) {
        Activity.contractName = "Office.System.Activity";
        function getFields(contract) {
            var instanceName = "Activity";
            var dataFields = [];
            if (contract.cV !== undefined) {
                dataFields.push(makeStringDataField(instanceName + ".CV", contract.cV));
            }
            dataFields.push(makeInt64DataField(instanceName + ".Duration", contract.duration));
            dataFields.push(makeInt64DataField(instanceName + ".Count", contract.count));
            dataFields.push(makeInt64DataField(instanceName + ".AggMode", contract.aggMode));
            if (contract.success !== undefined) {
                dataFields.push(makeBooleanDataField(instanceName + ".Success", contract.success));
            }
            if (contract.result !== undefined) {
                dataFields.push.apply(dataFields, officeeventschema_tml_Result.getFields(instanceName + ".Result", contract.result));
            }
            addContractField(dataFields, instanceName, Activity.contractName);
            return dataFields;
        }
        Activity.getFields = getFields;
    })(officeeventschema_tml_Activity || (officeeventschema_tml_Activity = {}));
    var officeeventschema_tml_Host;
    (function(Host) {
        var contractName = "Office.System.Host";
        function getFields(instanceName, contract) {
            var dataFields = [];
            if (contract.id !== undefined) {
                dataFields.push(makeStringDataField(instanceName + ".Id", contract.id));
            }
            if (contract.version !== undefined) {
                dataFields.push(makeStringDataField(instanceName + ".Version", contract.version));
            }
            if (contract.sessionId !== undefined) {
                dataFields.push(makeStringDataField(instanceName + ".SessionId", contract.sessionId));
            }
            addContractField(dataFields, instanceName, contractName);
            return dataFields;
        }
        Host.getFields = getFields;
    })(officeeventschema_tml_Host || (officeeventschema_tml_Host = {}));
    var officeeventschema_tml_User;
    (function(User) {
        var contractName = "Office.System.User";
        function getFields(instanceName, contract) {
            var dataFields = [];
            if (contract.alias !== undefined) {
                dataFields.push(makeStringDataField(instanceName + ".Alias", contract.alias));
            }
            if (contract.primaryIdentityHash !== undefined) {
                dataFields.push(makeStringDataField(instanceName + ".PrimaryIdentityHash", contract.primaryIdentityHash));
            }
            if (contract.primaryIdentitySpace !== undefined) {
                dataFields.push(makeStringDataField(instanceName + ".PrimaryIdentitySpace", contract.primaryIdentitySpace));
            }
            if (contract.tenantId !== undefined) {
                dataFields.push(makeStringDataField(instanceName + ".TenantId", contract.tenantId));
            }
            if (contract.tenantGroup !== undefined) {
                dataFields.push(makeStringDataField(instanceName + ".TenantGroup", contract.tenantGroup));
            }
            if (contract.isAnonymous !== undefined) {
                dataFields.push(makeBooleanDataField(instanceName + ".IsAnonymous", contract.isAnonymous));
            }
            addContractField(dataFields, instanceName, contractName);
            return dataFields;
        }
        User.getFields = getFields;
    })(officeeventschema_tml_User || (officeeventschema_tml_User = {}));
    var officeeventschema_tml_SDX;
    (function(SDX) {
        var contractName = "Office.System.SDX";
        function getFields(instanceName, contract) {
            var dataFields = [];
            if (contract.id !== undefined) {
                dataFields.push(makeStringDataField(instanceName + ".Id", contract.id));
            }
            if (contract.version !== undefined) {
                dataFields.push(makeStringDataField(instanceName + ".Version", contract.version));
            }
            if (contract.instanceId !== undefined) {
                dataFields.push(makeStringDataField(instanceName + ".InstanceId", contract.instanceId));
            }
            if (contract.name !== undefined) {
                dataFields.push(makeStringDataField(instanceName + ".Name", contract.name));
            }
            if (contract.marketplaceType !== undefined) {
                dataFields.push(makeStringDataField(instanceName + ".MarketplaceType", contract.marketplaceType));
            }
            if (contract.sessionId !== undefined) {
                dataFields.push(makeStringDataField(instanceName + ".SessionId", contract.sessionId));
            }
            if (contract.browserToken !== undefined) {
                dataFields.push(makeStringDataField(instanceName + ".BrowserToken", contract.browserToken));
            }
            if (contract.osfRuntimeVersion !== undefined) {
                dataFields.push(makeStringDataField(instanceName + ".OsfRuntimeVersion", contract.osfRuntimeVersion));
            }
            if (contract.officeJsVersion !== undefined) {
                dataFields.push(makeStringDataField(instanceName + ".OfficeJsVersion", contract.officeJsVersion));
            }
            if (contract.hostJsVersion !== undefined) {
                dataFields.push(makeStringDataField(instanceName + ".HostJsVersion", contract.hostJsVersion));
            }
            if (contract.assetId !== undefined) {
                dataFields.push(makeStringDataField(instanceName + ".AssetId", contract.assetId));
            }
            if (contract.providerName !== undefined) {
                dataFields.push(makeStringDataField(instanceName + ".ProviderName", contract.providerName));
            }
            if (contract.type !== undefined) {
                dataFields.push(makeStringDataField(instanceName + ".Type", contract.type));
            }
            addContractField(dataFields, instanceName, contractName);
            return dataFields;
        }
        SDX.getFields = getFields;
    })(officeeventschema_tml_SDX || (officeeventschema_tml_SDX = {}));
    var officeeventschema_tml_Funnel;
    (function(Funnel) {
        var contractName = "Office.System.Funnel";
        function getFields(instanceName, contract) {
            var dataFields = [];
            if (contract.name !== undefined) {
                dataFields.push(makeStringDataField(instanceName + ".Name", contract.name));
            }
            if (contract.state !== undefined) {
                dataFields.push(makeStringDataField(instanceName + ".State", contract.state));
            }
            addContractField(dataFields, instanceName, contractName);
            return dataFields;
        }
        Funnel.getFields = getFields;
    })(officeeventschema_tml_Funnel || (officeeventschema_tml_Funnel = {}));
    var officeeventschema_tml_UserAction;
    (function(UserAction) {
        var contractName = "Office.System.UserAction";
        function getFields(instanceName, contract) {
            var dataFields = [];
            if (contract.id !== undefined) {
                dataFields.push(makeInt64DataField(instanceName + ".Id", contract.id));
            }
            if (contract.name !== undefined) {
                dataFields.push(makeStringDataField(instanceName + ".Name", contract.name));
            }
            if (contract.commandSurface !== undefined) {
                dataFields.push(makeStringDataField(instanceName + ".CommandSurface", contract.commandSurface));
            }
            if (contract.parentName !== undefined) {
                dataFields.push(makeStringDataField(instanceName + ".ParentName", contract.parentName));
            }
            if (contract.triggerMethod !== undefined) {
                dataFields.push(makeStringDataField(instanceName + ".TriggerMethod", contract.triggerMethod));
            }
            if (contract.timeOffsetMs !== undefined) {
                dataFields.push(makeInt64DataField(instanceName + ".TimeOffsetMs", contract.timeOffsetMs));
            }
            addContractField(dataFields, instanceName, contractName);
            return dataFields;
        }
        UserAction.getFields = getFields;
    })(officeeventschema_tml_UserAction || (officeeventschema_tml_UserAction = {}));
    var Office_System_Error_Error;
    (function(Error) {
        var contractName = "Office.System.Error";
        function getFields(instanceName, contract) {
            var dataFields = [];
            dataFields.push(makeStringDataField(instanceName + ".ErrorGroup", contract.errorGroup));
            dataFields.push(makeInt64DataField(instanceName + ".Tag", contract.tag));
            if (contract.code !== undefined) {
                dataFields.push(makeInt64DataField(instanceName + ".Code", contract.code));
            }
            if (contract.id !== undefined) {
                dataFields.push(makeInt64DataField(instanceName + ".Id", contract.id));
            }
            if (contract.count !== undefined) {
                dataFields.push(makeInt64DataField(instanceName + ".Count", contract.count));
            }
            addContractField(dataFields, instanceName, contractName);
            return dataFields;
        }
        Error.getFields = getFields;
    })(Office_System_Error_Error || (Office_System_Error_Error = {}));
    var _Activity = officeeventschema_tml_Activity;
    var _Result = officeeventschema_tml_Result;
    var _Error = Office_System_Error_Error;
    var _Funnel = officeeventschema_tml_Funnel;
    var _Host = officeeventschema_tml_Host;
    var _SDX = officeeventschema_tml_SDX;
    var _UserAction = officeeventschema_tml_UserAction;
    var _User = officeeventschema_tml_User;
    var Contracts;
    (function(Contracts) {
        var Office;
        (function(Office) {
            var System;
            (function(System) {
                System.Activity = _Activity;
                System.Result = _Result;
                System.Error = _Error;
                System.Funnel = _Funnel;
                System.Host = _Host;
                System.SDX = _SDX;
                System.User = _User;
                System.UserAction = _UserAction;
            })(System = Office.System || (Office.System = {}));
        })(Office = Contracts.Office || (Contracts.Office = {}));
    })(Contracts || (Contracts = {}));
    var CorrelationVector;
    (function(CorrelationVector) {
        var baseHash;
        var baseId = 0;
        function getNext() {
            if (baseHash === undefined) {
                baseHash = generatePseudoHash();
            }
            return new CV(baseHash, ++baseId);
        }
        CorrelationVector.getNext = getNext;
        function getNextChild(parent) {
            return new CV(parent.getString(), ++parent.nextChild);
        }
        CorrelationVector.getNextChild = getNextChild;
        var CV = function() {
            function CV(base, id) {
                this.base = base;
                this.id = id;
                this.nextChild = 0;
            }
            CV.prototype.getString = function() {
                return this.base + "." + this.id;
            };
            return CV;
        }();
        CorrelationVector.CV = CV;
        function generatePseudoHash() {
            var characters = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/";
            var hashLength = 22;
            var result = [];
            for (var i = 0; i < hashLength; i++) {
                result.push(characters.charAt(Math.floor(Math.random() * characters.length)));
            }
            return result.join("");
        }
    })(CorrelationVector || (CorrelationVector = {}));
    var Event = function() {
        function Event() {
            this._listeners = [];
        }
        Event.prototype.fireEvent = function(args) {
            this._listeners.forEach(function(listener) {
                return listener(args);
            });
        };
        Event.prototype.addListener = function(listener) {
            if (listener) {
                this._listeners.push(listener);
            }
        };
        Event.prototype.removeListener = function(listener) {
            this._listeners = this._listeners.filter(function(h) {
                return h !== listener;
            });
        };
        Event.prototype.getListenerCount = function() {
            return this._listeners.length;
        };
        return Event;
    }();
    var onNotificationEvent = new Event();
    var LogLevel;
    (function(LogLevel) {
        LogLevel[LogLevel["Error"] = 0] = "Error";
        LogLevel[LogLevel["Warning"] = 1] = "Warning";
        LogLevel[LogLevel["Info"] = 2] = "Info";
        LogLevel[LogLevel["Verbose"] = 3] = "Verbose";
    })(LogLevel || (LogLevel = {}));
    var Category;
    (function(Category) {
        Category[Category["Core"] = 0] = "Core";
        Category[Category["Sink"] = 1] = "Sink";
        Category[Category["Transport"] = 2] = "Transport";
    })(Category || (Category = {}));
    function onNotification() {
        return onNotificationEvent;
    }
    function logNotification(level, category, message) {
        onNotificationEvent.fireEvent({
            level: level,
            category: category,
            message: message
        });
    }
    function logError(category, message, error) {
        logNotification(LogLevel.Error, category, function() {
            var errorMessage = error instanceof Error ? error.message : "";
            return message + ": " + errorMessage;
        });
    }
    var __awaiter = undefined && undefined.__awaiter || function(thisArg, _arguments, P, generator) {
        function adopt(value) {
            return value instanceof P ? value : new P(function(resolve) {
                resolve(value);
            });
        }
        return new (P || (P = Promise))(function(resolve, reject) {
            function fulfilled(value) {
                try {
                    step(generator.next(value));
                } catch (e) {
                    reject(e);
                }
            }
            function rejected(value) {
                try {
                    step(generator["throw"](value));
                } catch (e) {
                    reject(e);
                }
            }
            function step(result) {
                result.done ? resolve(result.value) : adopt(result.value).then(fulfilled, rejected);
            }
            step((generator = generator.apply(thisArg, _arguments || [])).next());
        });
    };
    var __generator = undefined && undefined.__generator || function(thisArg, body) {
        var _ = {
            label: 0,
            sent: function() {
                if (t[0] & 1) throw t[1];
                return t[1];
            },
            trys: [],
            ops: []
        }, f, y, t, g;
        return g = {
            next: verb(0),
            throw: verb(1),
            return: verb(2)
        }, typeof Symbol === "function" && (g[Symbol.iterator] = function() {
            return this;
        }), g;
        function verb(n) {
            return function(v) {
                return step([ n, v ]);
            };
        }
        function step(op) {
            if (f) throw new TypeError("Generator is already executing.");
            while (_) try {
                if (f = 1, y && (t = op[0] & 2 ? y["return"] : op[0] ? y["throw"] || ((t = y["return"]) && t.call(y), 
                0) : y.next) && !(t = t.call(y, op[1])).done) return t;
                if (y = 0, t) op = [ op[0] & 2, t.value ];
                switch (op[0]) {
                  case 0:
                  case 1:
                    t = op;
                    break;

                  case 4:
                    _.label++;
                    return {
                        value: op[1],
                        done: false
                    };

                  case 5:
                    _.label++;
                    y = op[1];
                    op = [ 0 ];
                    continue;

                  case 7:
                    op = _.ops.pop();
                    _.trys.pop();
                    continue;

                  default:
                    if (!(t = _.trys, t = t.length > 0 && t[t.length - 1]) && (op[0] === 6 || op[0] === 2)) {
                        _ = 0;
                        continue;
                    }
                    if (op[0] === 3 && (!t || op[1] > t[0] && op[1] < t[3])) {
                        _.label = op[1];
                        break;
                    }
                    if (op[0] === 6 && _.label < t[1]) {
                        _.label = t[1];
                        t = op;
                        break;
                    }
                    if (t && _.label < t[2]) {
                        _.label = t[2];
                        _.ops.push(op);
                        break;
                    }
                    if (t[2]) _.ops.pop();
                    _.trys.pop();
                    continue;
                }
                op = body.call(thisArg, _);
            } catch (e) {
                op = [ 6, e ];
                y = 0;
            } finally {
                f = t = 0;
            }
            if (op[0] & 5) throw op[1];
            return {
                value: op[0] ? op[1] : void 0,
                done: true
            };
        }
    };
    var ACTIVITY_COUNT = 1;
    var ACTIVITY_AGGMODE = 0;
    var getCurrentMicroseconds = function() {
        return Date.now() * 1e3;
    };
    if (typeof window.performance === "object" && "now" in window.performance) {
        getCurrentMicroseconds = function() {
            return Math.floor(window.performance.now()) * 1e3;
        };
    }
    var Activity_ActivityScope = function() {
        function ActivityScope(telemetryLogger, activityName, parent) {
            this._optionalEventFlags = {};
            this._ended = false;
            this._telemetryLogger = telemetryLogger;
            this._activityName = activityName;
            if (parent) {
                this._cv = CorrelationVector.getNextChild(parent._cv);
            } else {
                this._cv = CorrelationVector.getNext();
            }
            this._dataFields = [];
            this._success = undefined;
            this._startTime = getCurrentMicroseconds();
        }
        ActivityScope.createNew = function(telemetryLogger, activityName) {
            return new ActivityScope(telemetryLogger, activityName);
        };
        ActivityScope.prototype.createChildActivity = function(activityName) {
            var childActivity = new ActivityScope(this._telemetryLogger, activityName, this);
            return childActivity;
        };
        ActivityScope.prototype.setEventFlags = function(eventFlags) {
            this._optionalEventFlags = eventFlags;
        };
        ActivityScope.prototype.addDataField = function(dataField) {
            this._dataFields.push(dataField);
        };
        ActivityScope.prototype.addDataFields = function(dataFields) {
            var _a;
            (_a = this._dataFields).push.apply(_a, dataFields);
        };
        ActivityScope.prototype.setSuccess = function(success) {
            this._success = success;
        };
        ActivityScope.prototype.setResult = function(code, type, tag) {
            this._result = {
                code: code,
                type: type,
                tag: tag
            };
        };
        ActivityScope.prototype.endNow = function() {
            if (this._ended) {
                logNotification(LogLevel.Error, Category.Core, function() {
                    return "Activity has already ended";
                });
                return;
            }
            if (this._success === undefined && this._result === undefined) {
                logNotification(LogLevel.Warning, Category.Core, function() {
                    return "Activity does not have success or result set";
                });
            }
            var endTime = getCurrentMicroseconds();
            var duration = endTime - this._startTime;
            this._ended = true;
            var activity = {
                duration: duration,
                count: ACTIVITY_COUNT,
                aggMode: ACTIVITY_AGGMODE,
                cV: this._cv.getString(),
                success: this._success,
                result: this._result
            };
            return this._telemetryLogger.sendActivity(this._activityName, activity, this._dataFields, this._optionalEventFlags);
        };
        ActivityScope.prototype.executeAsync = function(activityBody) {
            return __awaiter(this, void 0, void 0, function() {
                var _this = this;
                return __generator(this, function(_a) {
                    return [ 2, activityBody(this).then(function(result) {
                        _this.endNow();
                        return result;
                    }).catch(function(e) {
                        _this.endNow();
                        throw e;
                    }) ];
                });
            });
        };
        ActivityScope.prototype.executeSync = function(activityBody) {
            try {
                var ret = activityBody(this);
                this.endNow();
                return ret;
            } catch (e) {
                this.endNow();
                throw e;
            }
        };
        ActivityScope.prototype.executeChildActivityAsync = function(activityName, activityBody) {
            return __awaiter(this, void 0, void 0, function() {
                return __generator(this, function(_a) {
                    return [ 2, this.createChildActivity(activityName).executeAsync(activityBody) ];
                });
            });
        };
        ActivityScope.prototype.executeChildActivitySync = function(activityName, activityBody) {
            return this.createChildActivity(activityName).executeSync(activityBody);
        };
        return ActivityScope;
    }();
    var DataClassification;
    (function(DataClassification) {
        DataClassification[DataClassification["EssentialServiceMetadata"] = 1] = "EssentialServiceMetadata";
        DataClassification[DataClassification["AccountData"] = 2] = "AccountData";
        DataClassification[DataClassification["SystemMetadata"] = 4] = "SystemMetadata";
        DataClassification[DataClassification["OrganizationIdentifiableInformation"] = 8] = "OrganizationIdentifiableInformation";
        DataClassification[DataClassification["EndUserIdentifiableInformation"] = 16] = "EndUserIdentifiableInformation";
        DataClassification[DataClassification["CustomerContent"] = 32] = "CustomerContent";
        DataClassification[DataClassification["AccessControl"] = 64] = "AccessControl";
    })(DataClassification || (DataClassification = {}));
    var SamplingPolicy;
    (function(SamplingPolicy) {
        SamplingPolicy[SamplingPolicy["NotSet"] = 0] = "NotSet";
        SamplingPolicy[SamplingPolicy["Measure"] = 1] = "Measure";
        SamplingPolicy[SamplingPolicy["Diagnostics"] = 2] = "Diagnostics";
        SamplingPolicy[SamplingPolicy["CriticalBusinessImpact"] = 191] = "CriticalBusinessImpact";
        SamplingPolicy[SamplingPolicy["CriticalCensus"] = 192] = "CriticalCensus";
        SamplingPolicy[SamplingPolicy["CriticalExperimentation"] = 193] = "CriticalExperimentation";
        SamplingPolicy[SamplingPolicy["CriticalUsage"] = 194] = "CriticalUsage";
    })(SamplingPolicy || (SamplingPolicy = {}));
    var PersistencePriority;
    (function(PersistencePriority) {
        PersistencePriority[PersistencePriority["NotSet"] = 0] = "NotSet";
        PersistencePriority[PersistencePriority["Normal"] = 1] = "Normal";
        PersistencePriority[PersistencePriority["High"] = 2] = "High";
    })(PersistencePriority || (PersistencePriority = {}));
    var CostPriority;
    (function(CostPriority) {
        CostPriority[CostPriority["NotSet"] = 0] = "NotSet";
        CostPriority[CostPriority["Normal"] = 1] = "Normal";
        CostPriority[CostPriority["High"] = 2] = "High";
    })(CostPriority || (CostPriority = {}));
    var DataCategories;
    (function(DataCategories) {
        DataCategories[DataCategories["NotSet"] = 0] = "NotSet";
        DataCategories[DataCategories["SoftwareSetup"] = 1] = "SoftwareSetup";
        DataCategories[DataCategories["ProductServiceUsage"] = 2] = "ProductServiceUsage";
        DataCategories[DataCategories["ProductServicePerformance"] = 4] = "ProductServicePerformance";
        DataCategories[DataCategories["DeviceConfiguration"] = 8] = "DeviceConfiguration";
        DataCategories[DataCategories["InkingTypingSpeech"] = 16] = "InkingTypingSpeech";
    })(DataCategories || (DataCategories = {}));
    var DiagnosticLevel;
    (function(DiagnosticLevel) {
        DiagnosticLevel[DiagnosticLevel["ReservedDoNotUse"] = 0] = "ReservedDoNotUse";
        DiagnosticLevel[DiagnosticLevel["BasicEvent"] = 10] = "BasicEvent";
        DiagnosticLevel[DiagnosticLevel["FullEvent"] = 100] = "FullEvent";
        DiagnosticLevel[DiagnosticLevel["NecessaryServiceDataEvent"] = 110] = "NecessaryServiceDataEvent";
        DiagnosticLevel[DiagnosticLevel["AlwaysOnNecessaryServiceDataEvent"] = 120] = "AlwaysOnNecessaryServiceDataEvent";
    })(DiagnosticLevel || (DiagnosticLevel = {}));
    function getEffectiveEventFlags(telemetryEvent) {
        var eventFlags = {
            costPriority: CostPriority.Normal,
            samplingPolicy: SamplingPolicy.Measure,
            persistencePriority: PersistencePriority.Normal,
            dataCategories: DataCategories.NotSet,
            diagnosticLevel: DiagnosticLevel.FullEvent
        };
        if (!telemetryEvent.eventFlags || !telemetryEvent.eventFlags.dataCategories) {
            logNotification(LogLevel.Error, Category.Core, function() {
                return "Event is missing DataCategories event flag";
            });
        }
        if (!telemetryEvent.eventFlags) {
            return eventFlags;
        }
        if (telemetryEvent.eventFlags.costPriority) {
            eventFlags.costPriority = telemetryEvent.eventFlags.costPriority;
        }
        if (telemetryEvent.eventFlags.samplingPolicy) {
            eventFlags.samplingPolicy = telemetryEvent.eventFlags.samplingPolicy;
        }
        if (telemetryEvent.eventFlags.persistencePriority) {
            eventFlags.persistencePriority = telemetryEvent.eventFlags.persistencePriority;
        }
        if (telemetryEvent.eventFlags.dataCategories) {
            eventFlags.dataCategories = telemetryEvent.eventFlags.dataCategories;
        }
        if (telemetryEvent.eventFlags.diagnosticLevel) {
            eventFlags.diagnosticLevel = telemetryEvent.eventFlags.diagnosticLevel;
        }
        return eventFlags;
    }
    var TokenType;
    (function(TokenType) {
        TokenType[TokenType["Aria"] = 0] = "Aria";
        TokenType[TokenType["Nexus"] = 1] = "Nexus";
    })(TokenType || (TokenType = {}));
    var TenantTokenManager_TenantTokenManager;
    (function(TenantTokenManager) {
        var ariaTokenMap = {};
        var nexusTokenMap = {};
        var tenantTokens = {};
        function setTenantToken(namespace, ariaTenantToken, nexusTenantToken) {
            var parts = namespace.split(".");
            if (parts.length < 2 || parts[0] !== "Office") {
                logNotification(LogLevel.Error, Category.Core, function() {
                    return "Invalid namespace: " + namespace;
                });
                return;
            }
            var leaf = Object.create(Object.prototype);
            if (ariaTenantToken) {
                leaf["ariaTenantToken"] = ariaTenantToken;
            }
            if (nexusTenantToken) {
                leaf["nexusTenantToken"] = nexusTenantToken;
            }
            var node = leaf;
            var index;
            for (index = parts.length - 1; index >= 0; --index) {
                var parentNode = Object.create(Object.prototype);
                parentNode[parts[index]] = node;
                node = parentNode;
            }
            setTenantTokens(node);
        }
        TenantTokenManager.setTenantToken = setTenantToken;
        function setTenantTokens(tokenTree) {
            if (typeof tokenTree !== "object") {
                throw new Error("tokenTree must be an object");
            }
            tenantTokens = mergeTenantTokens(tenantTokens, tokenTree);
        }
        TenantTokenManager.setTenantTokens = setTenantTokens;
        function getTenantTokens(eventName) {
            var ariaTenantToken = getAriaTenantToken(eventName);
            var nexusTenantToken = getNexusTenantToken(eventName);
            if (!nexusTenantToken || !ariaTenantToken) {
                throw new Error("Could not find tenant token for " + eventName);
            }
            return {
                ariaTenantToken: ariaTenantToken,
                nexusTenantToken: nexusTenantToken
            };
        }
        TenantTokenManager.getTenantTokens = getTenantTokens;
        function getAriaTenantToken(eventName) {
            if (ariaTokenMap[eventName]) {
                return ariaTokenMap[eventName];
            }
            var ariaToken = getTenantToken(eventName, TokenType.Aria);
            if (typeof ariaToken === "string") {
                ariaTokenMap[eventName] = ariaToken;
                return ariaToken;
            }
            return undefined;
        }
        TenantTokenManager.getAriaTenantToken = getAriaTenantToken;
        function getNexusTenantToken(eventName) {
            if (nexusTokenMap[eventName]) {
                return nexusTokenMap[eventName];
            }
            var nexusToken = getTenantToken(eventName, TokenType.Nexus);
            if (typeof nexusToken === "number") {
                nexusTokenMap[eventName] = nexusToken;
                return nexusToken;
            }
            return undefined;
        }
        TenantTokenManager.getNexusTenantToken = getNexusTenantToken;
        function getTenantToken(eventName, tokenType) {
            var pieces = eventName.split(".");
            var node = tenantTokens;
            var token = undefined;
            if (!node) {
                return undefined;
            }
            for (var i = 0; i < pieces.length - 1; i++) {
                if (node[pieces[i]]) {
                    node = node[pieces[i]];
                    if (tokenType === TokenType.Aria && typeof node.ariaTenantToken === "string") {
                        token = node.ariaTenantToken;
                    } else if (tokenType === TokenType.Nexus && typeof node.nexusTenantToken === "number") {
                        token = node.nexusTenantToken;
                    }
                }
            }
            return token;
        }
        function mergeTenantTokens(existingTokenTree, newTokenTree) {
            if (typeof newTokenTree !== "object") {
                return newTokenTree;
            }
            for (var _i = 0, _a = Object.keys(newTokenTree); _i < _a.length; _i++) {
                var key = _a[_i];
                if (key in existingTokenTree && typeof (existingTokenTree[key] === "object")) {
                    existingTokenTree[key] = mergeTenantTokens(existingTokenTree[key], newTokenTree[key]);
                } else {
                    existingTokenTree[key] = newTokenTree[key];
                }
            }
            return existingTokenTree;
        }
        function clear() {
            ariaTokenMap = {};
            nexusTokenMap = {};
            tenantTokens = {};
        }
        TenantTokenManager.clear = clear;
    })(TenantTokenManager_TenantTokenManager || (TenantTokenManager_TenantTokenManager = {}));
    var TelemetryEventValidator_TelemetryEventValidator;
    (function(TelemetryEventValidator) {
        var INT64_MIN = -9007199254740991;
        var INT64_MAX = 9007199254740991;
        var StartsWithCapitalRegex = /^[A-Z][a-zA-Z0-9]*$/;
        var AlphanumericRegex = /^[a-zA-Z0-9_\.]*$/;
        function validateTelemetryEvent(event) {
            if (!isEventNameValid(event.eventName)) {
                throw new Error("Invalid eventName");
            }
            if (event.eventContract && !isEventContractValid(event.eventContract)) {
                throw new Error("Invalid eventContract");
            }
            if (event.dataFields != null) {
                for (var i = 0; i < event.dataFields.length; i++) {
                    validateDataField(event.dataFields[i]);
                }
            }
        }
        TelemetryEventValidator.validateTelemetryEvent = validateTelemetryEvent;
        function isNamespaceValid(eventNamePieces) {
            return !!eventNamePieces && eventNamePieces.length >= 3 && eventNamePieces[0] === "Office";
        }
        function isEventNodeValid(eventNode) {
            return eventNode !== undefined && StartsWithCapitalRegex.test(eventNode);
        }
        function isEventNameValid(eventName) {
            var maxEventNameLength = 98;
            if (!eventName || eventName.length > maxEventNameLength) {
                return false;
            }
            var eventNamePieces = eventName.split(".");
            var eventNodeName = eventNamePieces[eventNamePieces.length - 1];
            return isNamespaceValid(eventNamePieces) && isEventNodeValid(eventNodeName);
        }
        function isEventContractValid(eventContract) {
            return isNameValid(eventContract.name);
        }
        function isDataFieldNameValid(dataFieldName) {
            var maxDataFieldNameLength = 100;
            var dataFieldPrefixLength = 5;
            return !!dataFieldName && isNameValid(dataFieldName) && dataFieldName.length + dataFieldPrefixLength < maxDataFieldNameLength;
        }
        function isNameValid(name) {
            return name !== undefined && AlphanumericRegex.test(name);
        }
        function validateDataField(dataField) {
            if (!isDataFieldNameValid(dataField.name)) {
                throw new Error("Invalid dataField name");
            }
            if (dataField.dataType === DataFieldType.Int64) {
                validateInt(dataField.value);
            }
        }
        function validateInt(value) {
            if (typeof value !== "number" || !isFinite(value) || Math.floor(value) !== value || value < INT64_MIN || value > INT64_MAX) {
                throw new Error("Invalid integer " + JSON.stringify(value));
            }
        }
        TelemetryEventValidator.validateInt = validateInt;
    })(TelemetryEventValidator_TelemetryEventValidator || (TelemetryEventValidator_TelemetryEventValidator = {}));
    var oteljsVersion = "3.1.28";
    var __assign = undefined && undefined.__assign || function() {
        __assign = Object.assign || function(t) {
            for (var s, i = 1, n = arguments.length; i < n; i++) {
                s = arguments[i];
                for (var p in s) if (Object.prototype.hasOwnProperty.call(s, p)) t[p] = s[p];
            }
            return t;
        };
        return __assign.apply(this, arguments);
    };
    var SuppressNexus = -1;
    var SimpleTelemetryLogger_SimpleTelemetryLogger = function() {
        function SimpleTelemetryLogger(parent, persistentDataFields, config) {
            var _a, _b;
            this.onSendEvent = new Event();
            this.persistentDataFields = [];
            this.config = config || {};
            if (parent) {
                this.onSendEvent = parent.onSendEvent;
                (_a = this.persistentDataFields).push.apply(_a, parent.persistentDataFields);
                this.config = __assign(__assign({}, parent.getConfig()), this.config);
            } else {
                this.persistentDataFields.push(makeStringDataField("OTelJS.Version", oteljsVersion));
            }
            if (persistentDataFields) {
                (_b = this.persistentDataFields).push.apply(_b, persistentDataFields);
            }
        }
        SimpleTelemetryLogger.prototype.sendTelemetryEvent = function(event) {
            var localEvent;
            try {
                if (this.onSendEvent.getListenerCount() === 0) {
                    logNotification(LogLevel.Warning, Category.Core, function() {
                        return "No telemetry sinks are attached.";
                    });
                    return;
                }
                localEvent = this.cloneEvent(event);
                this.processTelemetryEvent(localEvent);
            } catch (error) {
                logError(Category.Core, "SendTelemetryEvent", error);
                return;
            }
            try {
                this.onSendEvent.fireEvent(localEvent);
            } catch (_e) {}
        };
        SimpleTelemetryLogger.prototype.processTelemetryEvent = function(event) {
            var _a;
            if (!event.telemetryProperties) {
                event.telemetryProperties = TenantTokenManager_TenantTokenManager.getTenantTokens(event.eventName);
            }
            (_a = event.dataFields).push.apply(_a, this.persistentDataFields);
            if (!this.config.disableValidation) {
                TelemetryEventValidator_TelemetryEventValidator.validateTelemetryEvent(event);
            }
        };
        SimpleTelemetryLogger.prototype.addSink = function(sink) {
            this.onSendEvent.addListener(function(event) {
                return sink.sendTelemetryEvent(event);
            });
        };
        SimpleTelemetryLogger.prototype.setTenantToken = function(namespace, ariaTenantToken, nexusTenantToken) {
            TenantTokenManager_TenantTokenManager.setTenantToken(namespace, ariaTenantToken, nexusTenantToken);
        };
        SimpleTelemetryLogger.prototype.setTenantTokens = function(tokenTree) {
            TenantTokenManager_TenantTokenManager.setTenantTokens(tokenTree);
        };
        SimpleTelemetryLogger.prototype.cloneEvent = function(event) {
            var localEvent = {
                eventName: event.eventName,
                eventFlags: event.eventFlags
            };
            if (!!event.telemetryProperties) {
                localEvent.telemetryProperties = {
                    ariaTenantToken: event.telemetryProperties.ariaTenantToken,
                    nexusTenantToken: event.telemetryProperties.nexusTenantToken
                };
            }
            if (!!event.eventContract) {
                localEvent.eventContract = {
                    name: event.eventContract.name,
                    dataFields: event.eventContract.dataFields.slice()
                };
            }
            localEvent.dataFields = !!event.dataFields ? event.dataFields.slice() : [];
            return localEvent;
        };
        SimpleTelemetryLogger.prototype.getConfig = function() {
            return this.config;
        };
        return SimpleTelemetryLogger;
    }();
    var __extends = undefined && undefined.__extends || function() {
        var extendStatics = function(d, b) {
            extendStatics = Object.setPrototypeOf || {
                __proto__: []
            } instanceof Array && function(d, b) {
                d.__proto__ = b;
            } || function(d, b) {
                for (var p in b) if (b.hasOwnProperty(p)) d[p] = b[p];
            };
            return extendStatics(d, b);
        };
        return function(d, b) {
            extendStatics(d, b);
            function __() {
                this.constructor = d;
            }
            d.prototype = b === null ? Object.create(b) : (__.prototype = b.prototype, new __());
        };
    }();
    var TelemetryLogger_awaiter = undefined && undefined.__awaiter || function(thisArg, _arguments, P, generator) {
        function adopt(value) {
            return value instanceof P ? value : new P(function(resolve) {
                resolve(value);
            });
        }
        return new (P || (P = Promise))(function(resolve, reject) {
            function fulfilled(value) {
                try {
                    step(generator.next(value));
                } catch (e) {
                    reject(e);
                }
            }
            function rejected(value) {
                try {
                    step(generator["throw"](value));
                } catch (e) {
                    reject(e);
                }
            }
            function step(result) {
                result.done ? resolve(result.value) : adopt(result.value).then(fulfilled, rejected);
            }
            step((generator = generator.apply(thisArg, _arguments || [])).next());
        });
    };
    var TelemetryLogger_generator = undefined && undefined.__generator || function(thisArg, body) {
        var _ = {
            label: 0,
            sent: function() {
                if (t[0] & 1) throw t[1];
                return t[1];
            },
            trys: [],
            ops: []
        }, f, y, t, g;
        return g = {
            next: verb(0),
            throw: verb(1),
            return: verb(2)
        }, typeof Symbol === "function" && (g[Symbol.iterator] = function() {
            return this;
        }), g;
        function verb(n) {
            return function(v) {
                return step([ n, v ]);
            };
        }
        function step(op) {
            if (f) throw new TypeError("Generator is already executing.");
            while (_) try {
                if (f = 1, y && (t = op[0] & 2 ? y["return"] : op[0] ? y["throw"] || ((t = y["return"]) && t.call(y), 
                0) : y.next) && !(t = t.call(y, op[1])).done) return t;
                if (y = 0, t) op = [ op[0] & 2, t.value ];
                switch (op[0]) {
                  case 0:
                  case 1:
                    t = op;
                    break;

                  case 4:
                    _.label++;
                    return {
                        value: op[1],
                        done: false
                    };

                  case 5:
                    _.label++;
                    y = op[1];
                    op = [ 0 ];
                    continue;

                  case 7:
                    op = _.ops.pop();
                    _.trys.pop();
                    continue;

                  default:
                    if (!(t = _.trys, t = t.length > 0 && t[t.length - 1]) && (op[0] === 6 || op[0] === 2)) {
                        _ = 0;
                        continue;
                    }
                    if (op[0] === 3 && (!t || op[1] > t[0] && op[1] < t[3])) {
                        _.label = op[1];
                        break;
                    }
                    if (op[0] === 6 && _.label < t[1]) {
                        _.label = t[1];
                        t = op;
                        break;
                    }
                    if (t && _.label < t[2]) {
                        _.label = t[2];
                        _.ops.push(op);
                        break;
                    }
                    if (t[2]) _.ops.pop();
                    _.trys.pop();
                    continue;
                }
                op = body.call(thisArg, _);
            } catch (e) {
                op = [ 6, e ];
                y = 0;
            } finally {
                f = t = 0;
            }
            if (op[0] & 5) throw op[1];
            return {
                value: op[0] ? op[1] : void 0,
                done: true
            };
        }
    };
    var TelemetryLogger_TelemetryLogger = function(_super) {
        __extends(TelemetryLogger, _super);
        function TelemetryLogger() {
            return _super !== null && _super.apply(this, arguments) || this;
        }
        TelemetryLogger.prototype.executeActivityAsync = function(activityName, activityBody) {
            return TelemetryLogger_awaiter(this, void 0, void 0, function() {
                return TelemetryLogger_generator(this, function(_a) {
                    return [ 2, this.createNewActivity(activityName).executeAsync(activityBody) ];
                });
            });
        };
        TelemetryLogger.prototype.executeActivitySync = function(activityName, activityBody) {
            return this.createNewActivity(activityName).executeSync(activityBody);
        };
        TelemetryLogger.prototype.createNewActivity = function(activityName) {
            return Activity_ActivityScope.createNew(this, activityName);
        };
        TelemetryLogger.prototype.sendActivity = function(activityName, activity, dataFields, optionalEventFlags) {
            return this.sendTelemetryEvent({
                eventName: activityName,
                eventContract: {
                    name: Contracts.Office.System.Activity.contractName,
                    dataFields: Contracts.Office.System.Activity.getFields(activity)
                },
                dataFields: dataFields,
                eventFlags: optionalEventFlags
            });
        };
        TelemetryLogger.prototype.sendError = function(error) {
            var dataFields = Office_System_Error_Error.getFields("Error", error.error);
            if (error.dataFields != null) {
                dataFields.push.apply(dataFields, error.dataFields);
            }
            return this.sendTelemetryEvent({
                eventName: error.eventName,
                dataFields: dataFields,
                eventFlags: error.eventFlags
            });
        };
        return TelemetryLogger;
    }(SimpleTelemetryLogger_SimpleTelemetryLogger);
    __webpack_require__.d(__webpack_exports__, "Contracts", function() {
        return Contracts;
    });
    __webpack_require__.d(__webpack_exports__, "ActivityScope", function() {
        return Activity_ActivityScope;
    });
    __webpack_require__.d(__webpack_exports__, "getFieldsForContract", function() {
        return getFieldsForContract;
    });
    __webpack_require__.d(__webpack_exports__, "addContractField", function() {
        return addContractField;
    });
    __webpack_require__.d(__webpack_exports__, "DataClassification", function() {
        return DataClassification;
    });
    __webpack_require__.d(__webpack_exports__, "makeBooleanDataField", function() {
        return makeBooleanDataField;
    });
    __webpack_require__.d(__webpack_exports__, "makeInt64DataField", function() {
        return makeInt64DataField;
    });
    __webpack_require__.d(__webpack_exports__, "makeDoubleDataField", function() {
        return makeDoubleDataField;
    });
    __webpack_require__.d(__webpack_exports__, "makeStringDataField", function() {
        return makeStringDataField;
    });
    __webpack_require__.d(__webpack_exports__, "DataFieldType", function() {
        return DataFieldType;
    });
    __webpack_require__.d(__webpack_exports__, "getEffectiveEventFlags", function() {
        return getEffectiveEventFlags;
    });
    __webpack_require__.d(__webpack_exports__, "SamplingPolicy", function() {
        return SamplingPolicy;
    });
    __webpack_require__.d(__webpack_exports__, "PersistencePriority", function() {
        return PersistencePriority;
    });
    __webpack_require__.d(__webpack_exports__, "CostPriority", function() {
        return CostPriority;
    });
    __webpack_require__.d(__webpack_exports__, "DataCategories", function() {
        return DataCategories;
    });
    __webpack_require__.d(__webpack_exports__, "DiagnosticLevel", function() {
        return DiagnosticLevel;
    });
    __webpack_require__.d(__webpack_exports__, "LogLevel", function() {
        return LogLevel;
    });
    __webpack_require__.d(__webpack_exports__, "Category", function() {
        return Category;
    });
    __webpack_require__.d(__webpack_exports__, "onNotification", function() {
        return onNotification;
    });
    __webpack_require__.d(__webpack_exports__, "logNotification", function() {
        return logNotification;
    });
    __webpack_require__.d(__webpack_exports__, "logError", function() {
        return logError;
    });
    __webpack_require__.d(__webpack_exports__, "SuppressNexus", function() {
        return SuppressNexus;
    });
    __webpack_require__.d(__webpack_exports__, "SimpleTelemetryLogger", function() {
        return SimpleTelemetryLogger_SimpleTelemetryLogger;
    });
    __webpack_require__.d(__webpack_exports__, "TelemetryLogger", function() {
        return TelemetryLogger_TelemetryLogger;
    });
} ]);



OSFPerformance.officeExecuteEnd = OSFPerformance.now();
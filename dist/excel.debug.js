/*
 * Office JavaScript API library
 *
 * Copyright (c) Microsoft Corporation.  All rights reserved.
 *
 * Your use of this file is governed by the Microsoft Services Agreement http://go.microsoft.com/fwlink/?LinkId=266419.
 *
 * This file also contains the following Promise implementation (with a few small modifications):
 *      * @overview es6-promise - a tiny implementation of Promises/A+.
 *      * @copyright Copyright (c) 2014 Yehuda Katz, Tom Dale, Stefan Penner and contributors (Conversion to ES6 API by Jake Archibald)
 *      * @license   Licensed under MIT license
 *      *            See https://raw.githubusercontent.com/jakearchibald/es6-promise/master/LICENSE
 *      * @version   2.3.0
 */
var __assign = (this && this.__assign) || function () {
    __assign = Object.assign || function(t) {
        for (var s, i = 1, n = arguments.length; i < n; i++) {
            s = arguments[i];
            for (var p in s) if (Object.prototype.hasOwnProperty.call(s, p))
                t[p] = s[p];
        }
        return t;
    };
    return __assign.apply(this, arguments);
};
var __extends = (this && this.__extends) || (function () {
    var extendStatics = function (d, b) {
        extendStatics = Object.setPrototypeOf ||
            ({ __proto__: [] } instanceof Array && function (d, b) { d.__proto__ = b; }) ||
            function (d, b) { for (var p in b) if (b.hasOwnProperty(p)) d[p] = b[p]; };
        return extendStatics(d, b);
    };
    return function (d, b) {
        extendStatics(d, b);
        function __() { this.constructor = d; }
        d.prototype = b === null ? Object.create(b) : (__.prototype = b.prototype, new __());
    };
})();
var OSFPerformance;
(function (OSFPerformance) {
    function now() {
        if (performance && performance.now) {
            return performance.now();
        }
        else {
            return 0;
        }
    }
    OSFPerformance.now = now;
    OSFPerformance.officeExecuteStartDate = Date.now();
    OSFPerformance.officeExecuteStart = now();
    OSFPerformance.officeExecuteEnd = 0;
    OSFPerformance.hostInitializationStart = 0;
    OSFPerformance.hostInitializationEnd = 0;
    OSFPerformance.createOMEnd = 0;
    OSFPerformance.hostSpecificFileName = "";
    OSFPerformance.getAppContextStart = 0;
    OSFPerformance.getAppContextEnd = 0;
    OSFPerformance.getAppContextXdmStart = 0;
    OSFPerformance.getAppContextXdmEnd = 0;
    OSFPerformance.officeOnReady = 0;
})(OSFPerformance || (OSFPerformance = {}));
var OSF;
(function (OSF) {
    function definePropertyOnNamespace(o, name, getFunction) {
        Object.defineProperty(o, name, {
            get: function () {
                return getFunction();
            },
            configurable: true,
            enumerable: true
        });
    }
    OSF.definePropertyOnNamespace = definePropertyOnNamespace;
})(OSF || (OSF = {}));
var OSF;
(function (OSF) {
    var AgaveHostAction;
    (function (AgaveHostAction) {
        AgaveHostAction[AgaveHostAction["Select"] = 0] = "Select";
        AgaveHostAction[AgaveHostAction["UnSelect"] = 1] = "UnSelect";
        AgaveHostAction[AgaveHostAction["CancelDialog"] = 2] = "CancelDialog";
        AgaveHostAction[AgaveHostAction["InsertAgave"] = 3] = "InsertAgave";
        AgaveHostAction[AgaveHostAction["CtrlF6In"] = 4] = "CtrlF6In";
        AgaveHostAction[AgaveHostAction["CtrlF6Exit"] = 5] = "CtrlF6Exit";
        AgaveHostAction[AgaveHostAction["CtrlF6ExitShift"] = 6] = "CtrlF6ExitShift";
        AgaveHostAction[AgaveHostAction["SelectWithError"] = 7] = "SelectWithError";
        AgaveHostAction[AgaveHostAction["NotifyHostError"] = 8] = "NotifyHostError";
        AgaveHostAction[AgaveHostAction["RefreshAddinCommands"] = 9] = "RefreshAddinCommands";
        AgaveHostAction[AgaveHostAction["PageIsReady"] = 10] = "PageIsReady";
        AgaveHostAction[AgaveHostAction["TabIn"] = 11] = "TabIn";
        AgaveHostAction[AgaveHostAction["TabInShift"] = 12] = "TabInShift";
        AgaveHostAction[AgaveHostAction["TabExit"] = 13] = "TabExit";
        AgaveHostAction[AgaveHostAction["TabExitShift"] = 14] = "TabExitShift";
        AgaveHostAction[AgaveHostAction["EscExit"] = 15] = "EscExit";
        AgaveHostAction[AgaveHostAction["F2Exit"] = 16] = "F2Exit";
        AgaveHostAction[AgaveHostAction["ExitNoFocusable"] = 17] = "ExitNoFocusable";
        AgaveHostAction[AgaveHostAction["ExitNoFocusableShift"] = 18] = "ExitNoFocusableShift";
        AgaveHostAction[AgaveHostAction["MouseEnter"] = 19] = "MouseEnter";
        AgaveHostAction[AgaveHostAction["MouseLeave"] = 20] = "MouseLeave";
        AgaveHostAction[AgaveHostAction["UpdateTargetUrl"] = 21] = "UpdateTargetUrl";
        AgaveHostAction[AgaveHostAction["InstallCustomFunctions"] = 22] = "InstallCustomFunctions";
        AgaveHostAction[AgaveHostAction["SendTelemetryEvent"] = 23] = "SendTelemetryEvent";
        AgaveHostAction[AgaveHostAction["UninstallCustomFunctions"] = 24] = "UninstallCustomFunctions";
        AgaveHostAction[AgaveHostAction["SendMessage"] = 25] = "SendMessage";
        AgaveHostAction[AgaveHostAction["LaunchExtensionComponent"] = 26] = "LaunchExtensionComponent";
        AgaveHostAction[AgaveHostAction["StopExtensionComponent"] = 27] = "StopExtensionComponent";
        AgaveHostAction[AgaveHostAction["RestartExtensionComponent"] = 28] = "RestartExtensionComponent";
        AgaveHostAction[AgaveHostAction["EnableTaskPaneHeaderButton"] = 29] = "EnableTaskPaneHeaderButton";
        AgaveHostAction[AgaveHostAction["DisableTaskPaneHeaderButton"] = 30] = "DisableTaskPaneHeaderButton";
        AgaveHostAction[AgaveHostAction["TaskPaneHeaderButtonClicked"] = 31] = "TaskPaneHeaderButtonClicked";
        AgaveHostAction[AgaveHostAction["RemoveAppCommandsAddin"] = 32] = "RemoveAppCommandsAddin";
        AgaveHostAction[AgaveHostAction["RefreshRibbonGallery"] = 33] = "RefreshRibbonGallery";
    })(AgaveHostAction = OSF.AgaveHostAction || (OSF.AgaveHostAction = {}));
})(OSF || (OSF = {}));
var OSF;
(function (OSF) {
    var AppCommandManager = (function () {
        function AppCommandManager() {
            var _this = this;
            this._processAppCommandInvocation = function (args) {
                var verifyResult = _this._verifyManifestCallback(args.callbackName);
                if (verifyResult.errorCode != 0) {
                    _this._invokeAppCommandCompletedMethod(args.appCommandId, verifyResult.errorCode, "");
                    return;
                }
                var eventObj = _this._constructEventObjectForCallback(args);
                if (eventObj) {
                    window.setTimeout(function () { verifyResult.callback(eventObj); }, 0);
                }
                else {
                    _this._invokeAppCommandCompletedMethod(args.appCommandId, 5001, "");
                }
            };
            this._eventDispatch = new OSF.EventDispatch([
                {
                    type: OSF.EventType.AppCommandInvoked,
                    id: 39,
                    getTargetId: function () { return ''; },
                    fromSafeArrayHost: function (payload) {
                        return {
                            type: OSF.EventType.AppCommandInvoked,
                            appCommandId: payload[0],
                            callbackName: payload[1],
                            eventObjStr: payload[2]
                        };
                    },
                    fromWebHost: function (payload) {
                        return {
                            type: OSF.EventType.AppCommandInvoked,
                            appCommandId: payload[0],
                            callbackName: payload[1],
                            eventObjStr: payload[2]
                        };
                    }
                }
            ]);
        }
        AppCommandManager.prototype.initializeEventHandler = function (callback) {
            var _this = this;
            this.addHandlerAsync(OSF.EventType.AppCommandInvoked, function (args) {
                _this._processAppCommandInvocation(args);
            }, callback);
        };
        AppCommandManager.prototype.appCommandInvocationCompletedAsync = function (id, status, completedData, callback) {
            var asyncMethodExecutor = OSF._OfficeAppFactory.getAsyncMethodExecutor();
            asyncMethodExecutor.executeAsync(94, {
                fromSafeArrayHost: function (payload) {
                    return payload;
                },
                fromWebHost: function (payload) {
                    return payload;
                },
                toSafeArrayHost: function () {
                    return [id, status, completedData];
                },
                toWebHost: function () {
                    var obj = {};
                    obj[0] = id;
                    obj[1] = status;
                    obj[2] = completedData;
                    return obj;
                }
            }, callback);
        };
        AppCommandManager.prototype.addHandlerAsync = function (eventType, handler, callback) {
            OSF.EventHelper.addEventHandler(eventType, handler, callback, this._eventDispatch, false);
        };
        AppCommandManager.prototype._verifyManifestCallback = function (callbackName) {
            var defaultResult = { callback: null, errorCode: 11101 };
            callbackName = callbackName.trim();
            try {
                var callList = callbackName.split(".");
                var parentObject = window;
                for (var i = 0; i < callList.length - 1; i++) {
                    if (parentObject[callList[i]] && (typeof parentObject[callList[i]] == "object" || typeof parentObject[callList[i]] == "function")) {
                        parentObject = parentObject[callList[i]];
                    }
                    else {
                        return defaultResult;
                    }
                }
                var callbackFunc = parentObject[callList[callList.length - 1]];
                if (typeof callbackFunc != "function") {
                    return defaultResult;
                }
            }
            catch (e) {
                return defaultResult;
            }
            return { callback: callbackFunc, errorCode: 0 };
        };
        AppCommandManager.prototype._invokeAppCommandCompletedMethod = function (appCommandId, resultCode, data) {
            this.appCommandInvocationCompletedAsync(appCommandId, resultCode, data, function (result) {
                if (result.status !== Office.AsyncResultStatus.succeeded) {
                    console.error("Failed to notify the host thta app command is completed");
                }
            });
        };
        AppCommandManager.prototype._constructEventObjectForCallback = function (args) {
            var _this = this;
            var eventObj;
            var eventObjCopy;
            try {
                eventObj = JSON.parse(args.eventObjStr);
                eventObjCopy = JSON.parse(args.eventObjStr);
            }
            catch (ex) {
            }
            if (!eventObj) {
                eventObj = {};
            }
            if (!eventObjCopy) {
                eventObjCopy = {};
            }
            eventObj.completed = function (completedContext) {
                eventObjCopy.completedContext = completedContext;
                var jsonString = JSON.stringify(eventObjCopy);
                _this._invokeAppCommandCompletedMethod(args.appCommandId, 0, jsonString);
            };
            return eventObj;
        };
        AppCommandManager.initialize = function () {
            if (AppCommandManager._instance == null) {
                AppCommandManager._instance = new AppCommandManager();
                AppCommandManager._instance.initializeEventHandler(function (result) {
                    if (result.status !== Office.AsyncResultStatus.succeeded) {
                        console.error('Cannot initialize app command: ' + JSON.stringify(result));
                    }
                });
            }
        };
        AppCommandManager._instance = null;
        return AppCommandManager;
    }());
    OSF.AppCommandManager = AppCommandManager;
})(OSF || (OSF = {}));
var OSF;
(function (OSF) {
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
    }());
    OSF.Association = Association;
})(OSF || (OSF = {}));
var OSF;
(function (OSF) {
    var AsyncMethodExecutor = (function () {
        function AsyncMethodExecutor() {
        }
        AsyncMethodExecutor.prototype.invokeCallback = function (dispId, callback, status, value) {
            if (status == 0) {
                var successResult = {
                    status: Office.AsyncResultStatus.succeeded,
                    value: value
                };
                callback(successResult);
            }
            else {
                var errorResult = {
                    status: Office.AsyncResultStatus.failed,
                    error: {
                        code: status
                    }
                };
                callback(errorResult);
            }
        };
        return AsyncMethodExecutor;
    }());
    OSF.AsyncMethodExecutor = AsyncMethodExecutor;
})(OSF || (OSF = {}));
var OSF;
(function (OSF) {
    var ConstantNames;
    (function (ConstantNames) {
        ConstantNames["OfficeJS"] = "office.js";
        ConstantNames["OfficeDebugJS"] = "office.debug.js";
        ConstantNames["OfficeStringsId"] = "OFFICESTRINGS";
        ConstantNames["OfficeJsId"] = "OFFICEJS";
        ConstantNames["HostFileId"] = "HOST";
        ConstantNames["OfficeStringJS"] = "office_strings.js";
        ConstantNames["OfficeStringDebugJS"] = "office_strings.debug.js";
    })(ConstantNames = OSF.ConstantNames || (OSF.ConstantNames = {}));
})(OSF || (OSF = {}));
var OSF;
(function (OSF) {
    var Constants = (function () {
        function Constants() {
        }
        Constants.NotificationConversationIdSuffix = '_ntf';
        return Constants;
    }());
    OSF.Constants = Constants;
})(OSF || (OSF = {}));
var CustomFunctionMappings = window.CustomFunctionMappings || {};
var CustomFunctions;
(function (CustomFunctions) {
    function delayInitialization() {
        CustomFunctionMappings['__delay__'] = true;
    }
    CustomFunctions.delayInitialization = delayInitialization;
    ;
    CustomFunctions._association = new OSF.Association();
    function associate() {
        CustomFunctions._association.associate.apply(CustomFunctions._association, arguments);
        delete CustomFunctionMappings['__delay__'];
    }
    CustomFunctions.associate = associate;
    ;
})(CustomFunctions || (CustomFunctions = {}));
var OSF;
(function (OSF) {
    var ErrorCodeManager = (function () {
        function ErrorCodeManager() {
        }
        ErrorCodeManager.getAsyncResult = function (code) {
            if (code == 0) {
                return {
                    status: Office.AsyncResultStatus.succeeded
                };
            }
            else {
                return {
                    status: Office.AsyncResultStatus.failed,
                    error: {
                        code: code
                    }
                };
            }
        };
        return ErrorCodeManager;
    }());
    OSF.ErrorCodeManager = ErrorCodeManager;
})(OSF || (OSF = {}));
var OSF;
(function (OSF) {
    var EventDispatch = (function () {
        function EventDispatch(eventInfos) {
            this._eventInfos = {};
            this._queuedEventsArgs = {};
            this._eventHandlers = {};
            this._queuedEventsArgs = {};
            if (eventInfos != null) {
                for (var i = 0; i < eventInfos.length; i++) {
                    var eventType = eventInfos[i].type;
                    this._eventInfos[eventType] = eventInfos[i];
                    this._eventHandlers[eventType] = [];
                    this._queuedEventsArgs[eventType] = [];
                }
            }
        }
        EventDispatch.prototype.getSupportedEvents = function () {
            var events = [];
            for (var eventName in this._eventHandlers)
                events.push(eventName);
            return events;
        };
        EventDispatch.prototype.supportsEvent = function (event) {
            for (var eventName in this._eventHandlers) {
                if (event == eventName)
                    return true;
            }
            return false;
        };
        EventDispatch.prototype.hasEventHandler = function (eventType, handler) {
            var handlers = this._eventHandlers[eventType];
            if (handlers && handlers.length > 0) {
                for (var i = 0; i < handlers.length; i++) {
                    if (handlers[i] === handler)
                        return true;
                }
            }
            return false;
        };
        EventDispatch.prototype.addEventHandler = function (eventType, handler) {
            if (typeof handler != "function") {
                return false;
            }
            var handlers = this._eventHandlers[eventType];
            if (handlers && !this.hasEventHandler(eventType, handler)) {
                handlers.push(handler);
                return true;
            }
            else {
                return false;
            }
        };
        EventDispatch.prototype.addEventHandlerAndFireQueuedEvent = function (eventType, handler) {
            var handlers = this._eventHandlers[eventType];
            var isFirstHandler = (!handlers || handlers.length == 0);
            var succeed = this.addEventHandler(eventType, handler);
            if (isFirstHandler && succeed) {
                this.fireQueuedEvent(eventType);
            }
            return succeed;
        };
        EventDispatch.prototype.removeEventHandler = function (eventType, handler) {
            var handlers = this._eventHandlers[eventType];
            if (handlers && handlers.length > 0) {
                for (var index = 0; index < handlers.length; index++) {
                    if (handlers[index] === handler) {
                        handlers.splice(index, 1);
                        return true;
                    }
                }
            }
            return false;
        };
        EventDispatch.prototype.clearEventHandlers = function (eventType) {
            if (typeof this._eventHandlers[eventType] != "undefined" && this._eventHandlers[eventType].length > 0) {
                this._eventHandlers[eventType] = [];
                return true;
            }
            return false;
        };
        EventDispatch.prototype.getEventHandlerCount = function (eventType) {
            return this._eventHandlers[eventType] != undefined ? this._eventHandlers[eventType].length : -1;
        };
        EventDispatch.prototype.getEventInfo = function (eventType) {
            return this._eventInfos[eventType];
        };
        EventDispatch.prototype.fireEvent = function (eventArgs) {
            if (eventArgs.type == undefined)
                return false;
            var eventType = eventArgs.type;
            if (eventType && this._eventHandlers[eventType]) {
                var eventHandlers = this._eventHandlers[eventType];
                for (var i = 0; i < eventHandlers.length; i++) {
                    eventHandlers[i](eventArgs);
                }
                return true;
            }
            else {
                return false;
            }
        };
        EventDispatch.prototype.fireQueuedEvent = function (eventType) {
            if (eventType && this._eventHandlers[eventType]) {
                var eventHandlers = this._eventHandlers[eventType];
                var queuedEvents = this._queuedEventsArgs[eventType];
                if (eventHandlers.length > 0) {
                    var eventHandler = eventHandlers[0];
                    while (queuedEvents.length > 0) {
                        var eventArgs = queuedEvents.shift();
                        eventHandler(eventArgs);
                    }
                    return true;
                }
            }
            return false;
        };
        EventDispatch.prototype.clearQueuedEvent = function (eventType) {
            if (eventType && this._eventHandlers[eventType]) {
                var queuedEvents = this._queuedEventsArgs[eventType];
                if (queuedEvents) {
                    this._queuedEventsArgs[eventType] = [];
                }
            }
        };
        return EventDispatch;
    }());
    OSF.EventDispatch = EventDispatch;
})(OSF || (OSF = {}));
var OSF;
(function (OSF) {
    var EventHelper = (function () {
        function EventHelper() {
        }
        EventHelper.addEventHandler = function (eventType, handler, callback, eventDispatch, isPopupWindow) {
            var dispId = 0;
            function onEnsureRegistration(status) {
                if (status == 0) {
                    if (!eventDispatch.hasEventHandler(eventType, handler)) {
                        var added = eventDispatch.addEventHandler(eventType, handler);
                        if (!added) {
                            status = 5010;
                        }
                    }
                }
                var asyncMethodExecutor = OSF._OfficeAppFactory.getAsyncMethodExecutor();
                asyncMethodExecutor.invokeCallback(dispId, callback, status, null);
            }
            var eventInfo = eventDispatch.getEventInfo(eventType);
            if (!eventInfo) {
                onEnsureRegistration(5010);
                return;
            }
            try {
                if (isPopupWindow) {
                    onEnsureRegistration(0);
                    return;
                }
                dispId = eventInfo.id;
                var targetId = eventInfo.getTargetId();
                var count = eventDispatch.getEventHandlerCount(eventType);
                if (count == 0) {
                    var asyncMethodExecutor = OSF._OfficeAppFactory.getAsyncMethodExecutor();
                    asyncMethodExecutor.registerEventAsync(dispId, eventInfo.type, targetId, function (eventArgs) {
                        eventDispatch.fireEvent(eventArgs);
                    }, eventInfo, function (result) {
                        onEnsureRegistration(OSF.Utility.getErrorCodeFromAsyncResult(result));
                    });
                }
                else {
                    onEnsureRegistration(0);
                }
            }
            catch (ex) {
                EventHelper.onException(dispId, ex, callback);
            }
        };
        EventHelper.removeEventHandler = function (eventType, handler, callback, eventDispatch, isPopupWindow) {
            var dispId = 0;
            function onEnsureRegistration(status) {
                var asyncMethodExecutor = OSF._OfficeAppFactory.getAsyncMethodExecutor();
                asyncMethodExecutor.invokeCallback(dispId, callback, status, null);
            }
            var eventInfo = eventDispatch.getEventInfo(eventType);
            if (!eventInfo) {
                onEnsureRegistration(5010);
                return;
            }
            try {
                dispId = eventInfo.id;
                var targetId = eventInfo.getTargetId();
                var status_1 = 0;
                var removeSuccess = true;
                if (handler === null) {
                    removeSuccess = eventDispatch.clearEventHandlers(eventType);
                    status_1 = 0;
                }
                else {
                    removeSuccess = eventDispatch.removeEventHandler(eventType, handler);
                    status_1 = removeSuccess ? 0 : 5003;
                }
                var count = eventDispatch.getEventHandlerCount(eventType);
                if (removeSuccess && count == 0) {
                    var asyncMethodExecutor = OSF._OfficeAppFactory.getAsyncMethodExecutor();
                    asyncMethodExecutor.unregisterEventAsync(dispId, eventInfo.type, targetId, function (result) {
                        onEnsureRegistration(OSF.Utility.getErrorCodeFromAsyncResult(result));
                    });
                }
                else {
                    onEnsureRegistration(status_1);
                }
            }
            catch (ex) {
                EventHelper.onException(dispId, ex, callback);
            }
        };
        EventHelper.onException = function (dispId, ex, callback) {
            if (typeof ex == "number") {
                var asyncMethodExecutor = OSF._OfficeAppFactory.getAsyncMethodExecutor();
                asyncMethodExecutor.invokeCallback(dispId, callback, ex, null);
            }
            else {
                throw ex;
            }
        };
        return EventHelper;
    }());
    OSF.EventHelper = EventHelper;
})(OSF || (OSF = {}));
var OSF;
(function (OSF) {
    var EventType;
    (function (EventType) {
        EventType["AppCommandInvoked"] = "appCommandInvoked";
        EventType["RichApiMessage"] = "richApiMessage";
        EventType["BindingSelectionChanged"] = "bindingSelectionChanged";
        EventType["BindingDataChanged"] = "bindingDataChanged";
        EventType["DataNodeDeleted"] = "nodeDeleted";
        EventType["DataNodeInserted"] = "nodeInserted";
        EventType["DataNodeReplaced"] = "nodeReplaced";
        EventType["SettingsChanged"] = "settingsChanged";
    })(EventType = OSF.EventType || (OSF.EventType = {}));
})(OSF || (OSF = {}));
var OSF;
(function (OSF) {
    var HostName;
    (function (HostName) {
        var Host = (function () {
            function Host() {
                this.platformRemappings = {
                    web: Office.PlatformType.OfficeOnline,
                    winrt: Office.PlatformType.Universal,
                    win32: Office.PlatformType.PC,
                    mac: Office.PlatformType.Mac,
                    ios: Office.PlatformType.iOS,
                    android: Office.PlatformType.Android
                };
                this.camelCaseMappings = {
                    powerpoint: Office.HostType.PowerPoint,
                    onenote: Office.HostType.OneNote
                };
                this.hostInfo = OSF._OfficeAppFactory.getHostInfo();
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
                    if (Office.HostType[hostType]) {
                        return Office.HostType[hostType];
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
            Host.prototype.getDiagnostics = function (version) {
                var diagnostics = {
                    host: this.getHost(),
                    version: (version || this.getDefaultVersion()),
                    platform: this.getPlatform()
                };
                return diagnostics;
            };
            return Host;
        }());
        HostName.Host = Host;
    })(HostName = OSF.HostName || (OSF.HostName = {}));
})(OSF || (OSF = {}));
var OSF;
(function (OSF) {
    var HostInfoHostType;
    (function (HostInfoHostType) {
        HostInfoHostType["excel"] = "excel";
        HostInfoHostType["word"] = "word";
    })(HostInfoHostType = OSF.HostInfoHostType || (OSF.HostInfoHostType = {}));
})(OSF || (OSF = {}));
var OSF;
(function (OSF) {
    var HostInfoPlatform;
    (function (HostInfoPlatform) {
        HostInfoPlatform["web"] = "web";
        HostInfoPlatform["winrt"] = "winrt";
        HostInfoPlatform["win32"] = "win32";
        HostInfoPlatform["mac"] = "mac";
        HostInfoPlatform["ios"] = "ios";
        HostInfoPlatform["android"] = "android";
    })(HostInfoPlatform = OSF.HostInfoPlatform || (OSF.HostInfoPlatform = {}));
})(OSF || (OSF = {}));
var OSF;
(function (OSF) {
})(OSF || (OSF = {}));
var OSF;
(function (OSF) {
    var InitializationHelper = (function () {
        function InitializationHelper(hostInfo, webAppState, context, hostFacade) {
            this._hostInfo = hostInfo;
            this._webAppState = webAppState;
            this._context = context;
            this._hostFacade = hostFacade;
        }
        ;
        InitializationHelper.prototype.saveAndSetDialogInfo = function (hostInfoValue) {
        };
        InitializationHelper.prototype.setAgaveHostCommunication = function () {
        };
        InitializationHelper.prototype.createClientHostController = function () {
            return null;
        };
        InitializationHelper.prototype.createAsyncMethodExecutor = function () {
            return null;
        };
        InitializationHelper.prototype.createClientSettingsManager = function () {
            return null;
        };
        InitializationHelper.prototype.createSettings = function (serializedSettings) {
            var osfSessionStorage = OSF.OUtil.getSessionStorage();
            if (osfSessionStorage) {
                var storageSettings = osfSessionStorage.getItem(OSF._OfficeAppFactory.getCachedSessionSettingsKey());
                if (storageSettings) {
                    serializedSettings = JSON.parse(storageSettings);
                }
                else {
                    storageSettings = JSON.stringify(serializedSettings);
                    osfSessionStorage.setItem(OSF._OfficeAppFactory.getCachedSessionSettingsKey(), storageSettings);
                }
            }
            var deserializedSettings = OSF.OUtil.deserializeSettings(serializedSettings);
            var clientSettingsManager = this.createClientSettingsManager();
            var settings = new Office.Settings(deserializedSettings, clientSettingsManager);
            return settings;
        };
        InitializationHelper.prototype.prepareApiSurface = function (officeAppContext) {
            var featureGates = officeAppContext.get_featureGates();
            if (featureGates) {
                Microsoft.Office.WebExtension.FeatureGates = featureGates;
            }
            OSF.AppCommandManager.initialize();
            OSFPerformance.createOMEnd = OSFPerformance.now();
        };
        return InitializationHelper;
    }());
    OSF.InitializationHelper = InitializationHelper;
})(OSF || (OSF = {}));
var OSF;
(function (OSF) {
    var InitializationReason;
    (function (InitializationReason) {
        InitializationReason["Inserted"] = "inserted";
        InitializationReason["DocumentOpened"] = "documentOpened";
    })(InitializationReason = OSF.InitializationReason || (OSF.InitializationReason = {}));
})(OSF || (OSF = {}));
var OSF;
(function (OSF) {
    var LoadScriptHelper;
    (function (LoadScriptHelper) {
        var _scriptInfo;
        var _officeScriptBase = ['excel', 'word', 'powerpoint'];
        var _officeScriptSuffix = ['.js', '.debug.js'];
        function getHostBundleJsBasePath() {
            ensureScriptInfo();
            return _scriptInfo.basePath;
        }
        LoadScriptHelper.getHostBundleJsBasePath = getHostBundleJsBasePath;
        function getHostBundleJsName() {
            ensureScriptInfo();
            return _scriptInfo.name;
        }
        LoadScriptHelper.getHostBundleJsName = getHostBundleJsName;
        function ensureScriptInfo() {
            if (_scriptInfo) {
                return;
            }
            var getScriptBase = function (scriptSrc, scriptNameToCheck) {
                var scriptSrcLowerCase = scriptSrc.toLowerCase();
                var indexOfJS = scriptSrcLowerCase.indexOf(scriptNameToCheck);
                if (indexOfJS >= 0 &&
                    (indexOfJS === 0 || scriptSrc.charAt(indexOfJS - 1) === '/' || scriptSrc.charAt(indexOfJS - 1) === '\\') &&
                    (indexOfJS + scriptNameToCheck.length === scriptSrc.length || scriptSrc.charAt(indexOfJS + scriptNameToCheck.length) === '?')) {
                    var scriptBase = scriptSrc.substring(0, indexOfJS);
                    return { basePath: scriptBase, name: scriptNameToCheck };
                }
                return null;
            };
            var scripts = document.getElementsByTagName("script");
            var scriptsCount = scripts.length;
            for (var i = 0; i < scriptsCount; i++) {
                if (scripts[i].src) {
                    for (var j = 0; j < _officeScriptBase.length; j++) {
                        for (var k = 0; k < _officeScriptSuffix.length; k++) {
                            _scriptInfo = getScriptBase(scripts[i].src, _officeScriptBase[j] + _officeScriptSuffix[k]);
                            if (_scriptInfo) {
                                return;
                            }
                        }
                    }
                }
            }
            _scriptInfo = {
                basePath: "",
                name: ""
            };
        }
    })(LoadScriptHelper = OSF.LoadScriptHelper || (OSF.LoadScriptHelper = {}));
})(OSF || (OSF = {}));
var Microsoft;
(function (Microsoft) {
    var Office;
    (function (Office) {
        var WebExtension;
        (function (WebExtension) {
            WebExtension.FeatureGates = {};
            function sendTelemetryEvent(telemetryEvent) {
                OTel.OTelLogger.sendTelemetryEvent(telemetryEvent);
            }
            WebExtension.sendTelemetryEvent = sendTelemetryEvent;
        })(WebExtension = Office.WebExtension || (Office.WebExtension = {}));
    })(Office = Microsoft.Office || (Microsoft.Office = {}));
})(Microsoft || (Microsoft = {}));
var OSFPerfUtil;
(function (OSFPerfUtil) {
    function prepareDataFieldsForOtel(resource, name) {
        name = name + "_Resource";
        if (oteljs !== undefined) {
            return [
                oteljs.makeStringDataField(name + "_name", resource.name),
                oteljs.makeDoubleDataField(name + "_responseEnd", resource.responseEnd),
                oteljs.makeDoubleDataField(name + "_responseStart", resource.responseStart),
                oteljs.makeDoubleDataField(name + "_startTime", resource.startTime),
                oteljs.makeDoubleDataField(name + "_transferSize", resource.transferSize)
            ];
        }
        return [];
    }
    function sendPerformanceTelemetry() {
        if (OSF.AppTelemetry.enableTelemetry) {
            var hostPerfResource_1;
            var officePerfResource_1;
            var hostSpecificFileName_1 = OSF.LoadScriptHelper.getHostBundleJsName();
            var resources = performance.getEntriesByType("resource");
            resources.forEach(function (resource) {
                if (OSF.Utility.stringEndsWith(resource.name, hostSpecificFileName_1)) {
                    hostPerfResource_1 = resource;
                }
                else if (OSF.Utility.stringEndsWith(resource.name, OSF.ConstantNames.OfficeDebugJS) ||
                    OSF.Utility.stringEndsWith(resource.name, OSF.ConstantNames.OfficeJS)) {
                    officePerfResource_1 = resource;
                }
            });
            OTel.OTelLogger.onTelemetryLoaded(function () {
                var dataFields = [];
                if (hostPerfResource_1) {
                    dataFields = dataFields.concat(prepareDataFieldsForOtel(hostPerfResource_1, "HostJs"));
                }
                if (officePerfResource_1) {
                    dataFields = dataFields.concat(prepareDataFieldsForOtel(officePerfResource_1, "OfficeJs"));
                }
                dataFields = dataFields.concat([
                    oteljs.makeDoubleDataField("officeExecuteStartDate", OSFPerformance.officeExecuteStartDate),
                    oteljs.makeDoubleDataField("officeExecuteStart", OSFPerformance.officeExecuteStart),
                    oteljs.makeDoubleDataField("officeExecuteEnd", OSFPerformance.officeExecuteEnd),
                    oteljs.makeDoubleDataField("hostInitializationStart", OSFPerformance.hostInitializationStart),
                    oteljs.makeDoubleDataField("hostInitializationEnd", OSFPerformance.hostInitializationEnd),
                    oteljs.makeDoubleDataField("getAppContextStart", OSFPerformance.getAppContextStart),
                    oteljs.makeDoubleDataField("getAppContextEnd", OSFPerformance.getAppContextEnd),
                    oteljs.makeDoubleDataField("getAppContextXdmStart", OSFPerformance.getAppContextXdmStart),
                    oteljs.makeDoubleDataField("getAppContextXdmEnd", OSFPerformance.getAppContextXdmEnd),
                    oteljs.makeDoubleDataField("createOMEnd", OSFPerformance.createOMEnd),
                    oteljs.makeDoubleDataField("officeOnReady", OSFPerformance.officeOnReady)
                ]);
                Microsoft.Office.WebExtension.sendTelemetryEvent({
                    eventName: "Office.Extensibility.OfficeJs.JSPerformanceTelemetryV06",
                    dataFields: dataFields,
                    eventFlags: {
                        dataCategories: oteljs.DataCategories.ProductServiceUsage,
                        diagnosticLevel: oteljs.DiagnosticLevel.FullEvent
                    }
                });
            });
        }
    }
    OSFPerfUtil.sendPerformanceTelemetry = sendPerformanceTelemetry;
})(OSFPerfUtil || (OSFPerfUtil = {}));
var OSF;
(function (OSF) {
    var OUtil;
    (function (OUtil) {
        var _uniqueId = -1;
        var _xdmInfoKey = '&_xdm_Info=';
        var _serializerVersionKey = '&_serializer_version=';
        var _xdmSessionKeyPrefix = '_xdm_';
        var _serializerVersionKeyPrefix = '_serializer_version=';
        var _fragmentSeparator = '#';
        var _fragmentInfoDelimiter = '&';
        var _loadedScripts = {};
        var _defaultScriptLoadingTimeout = 30000;
        var _safeSessionStorage;
        var _safeLocalStorage;
        var Guid;
        (function (Guid) {
            var hexCode = ["0", "1", "2", "3", "4", "5", "6", "7", "8", "9", "a", "b", "c", "d", "e", "f"];
            function generateNewGuid() {
                var result = "";
                var tick = (new Date()).getTime();
                var index = 0;
                for (; index < 32 && tick > 0; index++) {
                    if (index == 8 || index == 12 || index == 16 || index == 20) {
                        result += "-";
                    }
                    result += hexCode[tick % 16];
                    tick = Math.floor(tick / 16);
                }
                for (; index < 32; index++) {
                    if (index == 8 || index == 12 || index == 16 || index == 20) {
                        result += "-";
                    }
                    result += hexCode[Math.floor(Math.random() * 16)];
                }
                return result;
            }
            Guid.generateNewGuid = generateNewGuid;
        })(Guid = OUtil.Guid || (OUtil.Guid = {}));
        function isArray(obj) {
            return Object.prototype.toString.apply(obj) === "[object Array]";
        }
        OUtil.isArray = isArray;
        function isFunction(obj) {
            return Object.prototype.toString.apply(obj) === "[object Function]";
        }
        OUtil.isFunction = isFunction;
        function isDate(obj) {
            return Object.prototype.toString.apply(obj) === "[object Date]";
        }
        OUtil.isDate = isDate;
        function addEventListener(element, eventName, listener) {
            if (element.addEventListener) {
                element.addEventListener(eventName, listener, false);
            }
            else if (element.attachEvent) {
                element.attachEvent("on" + eventName, listener);
            }
            else {
                throw new Error("Cannot attach event");
            }
        }
        OUtil.addEventListener = addEventListener;
        function removeEventListener(element, eventName, listener) {
            if (element.removeEventListener) {
                element.removeEventListener(eventName, listener, false);
            }
            else if (element.detachEvent) {
                element.detachEvent("on" + eventName, listener);
            }
            else {
                throw new Error("Cannot remove event");
            }
        }
        OUtil.removeEventListener = removeEventListener;
        var DateJSONPrefix = "Date(";
        var DataJSONSuffix = ")";
        function serializeSettings(settingsCollection) {
            var ret = {};
            for (var key in settingsCollection) {
                var value = settingsCollection[key];
                try {
                    value = JSON.stringify(value, function dateReplacer(k, v) {
                        return OSF.OUtil.isDate(this[k]) ? DateJSONPrefix + this[k].getTime() + DataJSONSuffix : v;
                    });
                    ret[key] = value;
                }
                catch (ex) {
                }
            }
            return ret;
        }
        OUtil.serializeSettings = serializeSettings;
        function deserializeSettings(serializedSettings) {
            var ret = {};
            serializedSettings = serializedSettings || {};
            for (var key in serializedSettings) {
                var value = serializedSettings[key];
                try {
                    value = JSON.parse(value, function dateReviver(k, v) {
                        var d;
                        if (typeof v === 'string' && v && v.length > 6 && v.slice(0, 5) === DateJSONPrefix && v.slice(-1) === DataJSONSuffix) {
                            d = new Date(parseInt(v.slice(5, -1)));
                            if (d) {
                                return d;
                            }
                        }
                        return v;
                    });
                    ret[key] = value;
                }
                catch (ex) {
                }
            }
            return ret;
        }
        OUtil.deserializeSettings = deserializeSettings;
        function loadScript(url, callback, timeoutInMs) {
            if (url && callback) {
                var doc = window.document;
                var _loadedScriptEntry = _loadedScripts[url];
                if (!_loadedScriptEntry) {
                    var script = doc.createElement("script");
                    script.type = "text/javascript";
                    _loadedScriptEntry = { loaded: false, pendingCallbacks: [callback], timer: null };
                    _loadedScripts[url] = _loadedScriptEntry;
                    var onLoadCallback = function OSF_OUtil_loadScript$onLoadCallback() {
                        if (_loadedScriptEntry.timer != null) {
                            clearTimeout(_loadedScriptEntry.timer);
                            delete _loadedScriptEntry.timer;
                        }
                        _loadedScriptEntry.loaded = true;
                        var pendingCallbackCount = _loadedScriptEntry.pendingCallbacks.length;
                        for (var i = 0; i < pendingCallbackCount; i++) {
                            var currentCallback = _loadedScriptEntry.pendingCallbacks.shift();
                            currentCallback(true);
                        }
                    };
                    var onLoadError = function OSF_OUtil_loadScript$onLoadError() {
                        delete _loadedScripts[url];
                        if (_loadedScriptEntry.timer != null) {
                            clearTimeout(_loadedScriptEntry.timer);
                            delete _loadedScriptEntry.timer;
                        }
                        var pendingCallbackCount = _loadedScriptEntry.pendingCallbacks.length;
                        for (var i = 0; i < pendingCallbackCount; i++) {
                            var currentCallback = _loadedScriptEntry.pendingCallbacks.shift();
                            currentCallback(false);
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
                    timeoutInMs = timeoutInMs || _defaultScriptLoadingTimeout;
                    _loadedScriptEntry.timer = setTimeout(onLoadError, timeoutInMs);
                    script.setAttribute("crossOrigin", "anonymous");
                    script.src = url;
                    doc.getElementsByTagName("head")[0].appendChild(script);
                }
                else if (_loadedScriptEntry.loaded) {
                    callback(true);
                }
                else {
                    _loadedScriptEntry.pendingCallbacks.push(callback);
                }
            }
        }
        OUtil.loadScript = loadScript;
        function getSessionStorage() {
            if (!_safeSessionStorage) {
                try {
                    var sessionStorage = window.sessionStorage;
                }
                catch (ex) {
                    sessionStorage = null;
                }
                _safeSessionStorage = new OSF.SafeStorage(sessionStorage);
            }
            return _safeSessionStorage;
        }
        OUtil.getSessionStorage = getSessionStorage;
        function getLocalStorage() {
            if (!_safeLocalStorage) {
                try {
                    var localStorage = window.localStorage;
                }
                catch (ex) {
                    localStorage = null;
                }
                _safeLocalStorage = new OSF.SafeStorage(localStorage);
            }
            return _safeLocalStorage;
        }
        OUtil.getLocalStorage = getLocalStorage;
        function convertIntToCssHexColor(val) {
            var hex = "#" + (Number(val) + 0x1000000).toString(16).slice(-6);
            return hex;
        }
        OUtil.convertIntToCssHexColor = convertIntToCssHexColor;
        function parseAppContextFromWindowName(skipSessionStorage, windowName) {
            return OUtil.parseInfoFromWindowName(skipSessionStorage, windowName, "appContext");
        }
        OUtil.parseAppContextFromWindowName = parseAppContextFromWindowName;
        function parseHostInfoFromWindowName(skipSessionStorage, windowName) {
            return OSF.OUtil.parseInfoFromWindowName(skipSessionStorage, windowName, "hostInfo");
        }
        OUtil.parseHostInfoFromWindowName = parseHostInfoFromWindowName;
        function parseXdmInfo(skipSessionStorage) {
            var xdmInfoValue = OUtil.parseXdmInfoWithGivenFragment(skipSessionStorage, window.location.hash);
            if (!xdmInfoValue) {
                xdmInfoValue = OUtil.parseXdmInfoFromWindowName(skipSessionStorage, window.name);
            }
            return xdmInfoValue;
        }
        OUtil.parseXdmInfo = parseXdmInfo;
        function parseXdmInfoFromWindowName(skipSessionStorage, windowName) {
            return OSF.OUtil.parseInfoFromWindowName(skipSessionStorage, windowName, "xdmInfo");
        }
        OUtil.parseXdmInfoFromWindowName = parseXdmInfoFromWindowName;
        function parseXdmInfoWithGivenFragment(skipSessionStorage, fragment) {
            return OSF.OUtil.parseInfoWithGivenFragment(_xdmInfoKey, _xdmSessionKeyPrefix, false, skipSessionStorage, fragment);
        }
        OUtil.parseXdmInfoWithGivenFragment = parseXdmInfoWithGivenFragment;
        function parseSerializerVersion(skipSessionStorage) {
            var serializerVersion = OSF.OUtil.parseSerializerVersionWithGivenFragment(skipSessionStorage, window.location.hash);
            if (isNaN(serializerVersion)) {
                serializerVersion = OSF.OUtil.parseSerializerVersionFromWindowName(skipSessionStorage, window.name);
            }
            return serializerVersion;
        }
        OUtil.parseSerializerVersion = parseSerializerVersion;
        function parseSerializerVersionFromWindowName(skipSessionStorage, windowName) {
            return parseInt(OSF.OUtil.parseInfoFromWindowName(skipSessionStorage, windowName, "serializerVersion"));
        }
        OUtil.parseSerializerVersionFromWindowName = parseSerializerVersionFromWindowName;
        function parseSerializerVersionWithGivenFragment(skipSessionStorage, fragment) {
            return parseInt(OSF.OUtil.parseInfoWithGivenFragment(_serializerVersionKey, _serializerVersionKeyPrefix, true, skipSessionStorage, fragment));
        }
        OUtil.parseSerializerVersionWithGivenFragment = parseSerializerVersionWithGivenFragment;
        function parseInfoFromWindowName(skipSessionStorage, windowName, infoKey) {
            try {
                var windowNameObj = JSON.parse(windowName);
                var infoValue = windowNameObj != null ? windowNameObj[infoKey] : null;
                var osfSessionStorage = OUtil.getSessionStorage();
                if (!skipSessionStorage && osfSessionStorage && windowNameObj != null) {
                    var sessionKey = windowNameObj["baseFrameName"] + infoKey;
                    if (infoValue) {
                        osfSessionStorage.setItem(sessionKey, infoValue);
                    }
                    else {
                        infoValue = osfSessionStorage.getItem(sessionKey);
                    }
                }
                return infoValue;
            }
            catch (Exception) {
                return null;
            }
        }
        OUtil.parseInfoFromWindowName = parseInfoFromWindowName;
        function parseInfoWithGivenFragment(infoKey, infoKeyPrefix, decodeInfo, skipSessionStorage, fragment) {
            var fragmentParts = fragment.split(infoKey);
            var infoValue = fragmentParts.length > 1 ? fragmentParts[fragmentParts.length - 1] : null;
            if (decodeInfo && infoValue != null) {
                if (infoValue.indexOf(_fragmentInfoDelimiter) >= 0) {
                    infoValue = infoValue.split(_fragmentInfoDelimiter)[0];
                }
                infoValue = decodeURIComponent(infoValue);
            }
            var osfSessionStorage = OUtil.getSessionStorage();
            if (!skipSessionStorage && osfSessionStorage) {
                var sessionKeyStart = window.name.indexOf(infoKeyPrefix);
                if (sessionKeyStart > -1) {
                    var sessionKeyEnd = window.name.indexOf(";", sessionKeyStart);
                    if (sessionKeyEnd == -1) {
                        sessionKeyEnd = window.name.length;
                    }
                    var sessionKey = window.name.substring(sessionKeyStart, sessionKeyEnd);
                    if (infoValue) {
                        osfSessionStorage.setItem(sessionKey, infoValue);
                    }
                    else {
                        infoValue = osfSessionStorage.getItem(sessionKey);
                    }
                }
            }
            return infoValue;
        }
        OUtil.parseInfoWithGivenFragment = parseInfoWithGivenFragment;
        function getConversationId() {
            var searchString = window.location.search;
            var conversationId = null;
            if (searchString) {
                var index = searchString.indexOf("&");
                conversationId = index > 0 ? searchString.substring(1, index) : searchString.substr(1);
                if (conversationId && conversationId.charAt(conversationId.length - 1) === '=') {
                    conversationId = conversationId.substring(0, conversationId.length - 1);
                    if (conversationId) {
                        conversationId = decodeURIComponent(conversationId);
                    }
                }
            }
            return conversationId;
        }
        OUtil.getConversationId = getConversationId;
        function getInfoItems(strInfo) {
            var items = strInfo.split('$');
            if (typeof items[1] == "undefined") {
                items = strInfo.split("|");
            }
            if (typeof items[1] == "undefined") {
                items = strInfo.split("%7C");
            }
            return items;
        }
        OUtil.getInfoItems = getInfoItems;
        function getXdmFieldValue(xdmFieldName, skipSessionStorage) {
            var fieldValue = '';
            var xdmInfoValue = OSF.OUtil.parseXdmInfo(skipSessionStorage);
            if (xdmInfoValue) {
                var items = OSF.OUtil.getInfoItems(xdmInfoValue);
                if (items != undefined && items.length >= 3) {
                    switch (xdmFieldName) {
                        case "ConversationUrl":
                            fieldValue = items[2];
                            break;
                        case "AppId":
                            fieldValue = items[1];
                            break;
                    }
                }
            }
            return fieldValue;
        }
        OUtil.getXdmFieldValue = getXdmFieldValue;
        function _reOrderTabbableElements(elements) {
            var bucket0 = [];
            var bucketPositive = [];
            var i;
            var len = elements.length;
            var ele;
            for (i = 0; i < len; i++) {
                ele = elements[i];
                if (ele.tabIndex) {
                    if (ele.tabIndex > 0) {
                        bucketPositive.push(ele);
                    }
                    else if (ele.tabIndex === 0) {
                        bucket0.push(ele);
                    }
                }
                else {
                    bucket0.push(ele);
                }
            }
            bucketPositive = bucketPositive.sort(function (left, right) {
                var diff = left.tabIndex - right.tabIndex;
                if (diff === 0) {
                    diff = bucketPositive.indexOf(left) - bucketPositive.indexOf(right);
                }
                return diff;
            });
            return [].concat(bucketPositive, bucket0);
        }
        ;
        function focusToFirstTabbable(all, backward) {
            var next;
            var focused = false;
            var candidate;
            var setFlag = function (e) {
                focused = true;
            };
            var findNextPos = function (allLen, currPos, backward) {
                if (currPos < 0 || currPos > allLen) {
                    return -1;
                }
                else if (currPos === 0 && backward) {
                    return -1;
                }
                else if (currPos === allLen - 1 && !backward) {
                    return -1;
                }
                if (backward) {
                    return currPos - 1;
                }
                else {
                    return currPos + 1;
                }
            };
            all = _reOrderTabbableElements(all);
            next = backward ? all.length - 1 : 0;
            if (all.length === 0) {
                return null;
            }
            while (!focused && next >= 0 && next < all.length) {
                candidate = all[next];
                window.focus();
                candidate.addEventListener('focus', setFlag);
                candidate.focus();
                candidate.removeEventListener('focus', setFlag);
                next = findNextPos(all.length, next, backward);
                if (!focused && candidate === document.activeElement) {
                    focused = true;
                }
            }
            if (focused) {
                return candidate;
            }
            else {
                return null;
            }
        }
        OUtil.focusToFirstTabbable = focusToFirstTabbable;
        function focusToNextTabbable(all, curr, shift) {
            var currPos;
            var next;
            var focused = false;
            var candidate;
            var setFlag = function (e) {
                focused = true;
            };
            var findCurrPos = function (all, curr) {
                var i = 0;
                for (; i < all.length; i++) {
                    if (all[i] === curr) {
                        return i;
                    }
                }
                return -1;
            };
            var findNextPos = function (allLen, currPos, shift) {
                if (currPos < 0 || currPos > allLen) {
                    return -1;
                }
                else if (currPos === 0 && shift) {
                    return -1;
                }
                else if (currPos === allLen - 1 && !shift) {
                    return -1;
                }
                if (shift) {
                    return currPos - 1;
                }
                else {
                    return currPos + 1;
                }
            };
            all = _reOrderTabbableElements(all);
            currPos = findCurrPos(all, curr);
            next = findNextPos(all.length, currPos, shift);
            if (next < 0) {
                return null;
            }
            while (!focused && next >= 0 && next < all.length) {
                candidate = all[next];
                candidate.addEventListener('focus', setFlag);
                candidate.focus();
                candidate.removeEventListener('focus', setFlag);
                next = findNextPos(all.length, next, shift);
                if (!focused && candidate === document.activeElement) {
                    focused = true;
                }
            }
            if (focused) {
                return candidate;
            }
            else {
                return null;
            }
        }
        OUtil.focusToNextTabbable = focusToNextTabbable;
    })(OUtil = OSF.OUtil || (OSF.OUtil = {}));
})(OSF || (OSF = {}));
var Office;
(function (Office) {
    var context;
    (function (context) {
        var document;
        (function (document) {
            function get_url() {
                return OSF._OfficeAppFactory.getOfficeAppContext().get_docUrl();
            }
            OSF.definePropertyOnNamespace(document, 'url', get_url);
            function get_mode() {
                var clientMode = OSF._OfficeAppFactory.getOfficeAppContext().get_clientMode();
                if (clientMode == 0) {
                    return Office.DocumentMode.ReadOnly;
                }
                return Office.DocumentMode.ReadWrite;
            }
            OSF.definePropertyOnNamespace(document, 'mode', get_mode);
            var _settings;
            function get_settings() {
                if (!_settings) {
                    var settingsFunc = OSF._OfficeAppFactory.getOfficeAppContext().get_settingsFunc();
                    var serializedSettings = settingsFunc();
                    _settings = OSF._OfficeAppFactory.getInitializationHelper().createSettings(serializedSettings);
                }
                return _settings;
            }
            OSF.definePropertyOnNamespace(document, 'settings', get_settings);
        })(document = context.document || (context.document = {}));
    })(context = Office.context || (Office.context = {}));
})(Office || (Office = {}));
var Office;
(function (Office) {
    var context;
    (function (context) {
        var messaging;
        (function (messaging) {
            function sendMessage(message) {
                var hostInfo = OSF._OfficeAppFactory.getHostInfo();
                if (hostInfo.hostPlatform == OSF.HostInfoPlatform.web) {
                    var webAppState = OSF._OfficeAppFactory.getWebAppState();
                    webAppState.clientEndPoint.invoke("ContextActivationManager_notifyHost", null, [webAppState.id, OSF.AgaveHostAction.SendMessage, message]);
                }
                else {
                    throw OSF.Utility.createNotImplementedException();
                }
            }
            messaging.sendMessage = sendMessage;
        })(messaging = context.messaging || (context.messaging = {}));
    })(context = Office.context || (Office.context = {}));
})(Office || (Office = {}));
var Office;
(function (Office) {
    var context;
    (function (context) {
        function get_contentLanguage() {
            return OSF._OfficeAppFactory.getOfficeAppContext().get_dataLocale();
        }
        OSF.definePropertyOnNamespace(context, 'contentLanguage', get_contentLanguage);
        function get_displayLanguage() {
            return OSF._OfficeAppFactory.getOfficeAppContext().get_appUILocale();
        }
        OSF.definePropertyOnNamespace(context, 'displayLanguage', get_displayLanguage);
        function get_isDialog() {
            return OSF._OfficeAppFactory.getHostInfo().isDialog;
        }
        OSF.definePropertyOnNamespace(context, 'isDialog', get_isDialog);
        function get_touchEnabled() {
            return OSF._OfficeAppFactory.getOfficeAppContext().get_touchEnabled();
        }
        OSF.definePropertyOnNamespace(context, 'touchEnabled', get_touchEnabled);
        function get_commerceAllowed() {
            return OSF._OfficeAppFactory.getOfficeAppContext().get_commerceAllowed();
        }
        OSF.definePropertyOnNamespace(context, 'commerceAllowed', get_commerceAllowed);
        function get_host() {
            return OSF.HostName.Host.getInstance().getHost();
        }
        OSF.definePropertyOnNamespace(context, 'host', get_host);
        function get_platform() {
            return OSF.HostName.Host.getInstance().getPlatform();
        }
        OSF.definePropertyOnNamespace(context, 'platform', get_platform);
        function get_diagnostics() {
            return OSF.HostName.Host.getInstance().getDiagnostics(OSF._OfficeAppFactory.getOfficeAppContext().get_hostFullVersion());
        }
        OSF.definePropertyOnNamespace(context, 'diagnostics', get_diagnostics);
        var _requirements;
        function get_requirements() {
            if (!_requirements) {
                var appContext = OSF._OfficeAppFactory.getOfficeAppContext();
                if (appContext.get_isDialog()) {
                    _requirements = OSF.Requirement.RequirementsMatrixFactory.getDefaultDialogRequirementMatrix(appContext);
                }
                else {
                    _requirements = OSF.Requirement.RequirementsMatrixFactory.getDefaultRequirementMatrix(appContext);
                }
            }
            return _requirements;
        }
        OSF.definePropertyOnNamespace(context, 'requirements', get_requirements);
        var _officeTheme;
        function get_officeTheme() {
            if (!_officeTheme) {
                var func = OSF._OfficeAppFactory.getOfficeAppContext().get_officeThemeFunc();
                if (func) {
                    _officeTheme = func();
                }
                else {
                    return undefined;
                }
            }
            return _officeTheme;
        }
        OSF.definePropertyOnNamespace(context, 'officeTheme', get_officeTheme);
    })(context = Office.context || (Office.context = {}));
})(Office || (Office = {}));
var Office;
(function (Office) {
    var context;
    (function (context) {
        var ui;
        (function (ui) {
            var taskPaneAction;
            (function (taskPaneAction) {
            })(taskPaneAction = ui.taskPaneAction || (ui.taskPaneAction = {}));
        })(ui = context.ui || (context.ui = {}));
    })(context = Office.context || (Office.context = {}));
})(Office || (Office = {}));
var Office;
(function (Office) {
    var _isOfficeOnReadyCalled = false;
    var _officeOnReadyPromise = null;
    var _officeOnReadyPromiseResolve = null;
    var _officeOnReadyCallbacks = [];
    var _officeOnReadyHostAndPlatformInfo;
    var _officeOnReadyFired;
    function ensureOfficeOnReadyPromise() {
        if (!_officeOnReadyPromise) {
            _officeOnReadyPromise = new Office.Promise(function (resolve, reject) {
                _officeOnReadyPromiseResolve = resolve;
            });
        }
    }
    function onReadyInternal(callback) {
        ensureOfficeOnReadyPromise();
        if (callback) {
            if (_officeOnReadyFired) {
                callback(_officeOnReadyHostAndPlatformInfo);
            }
            else {
                _officeOnReadyCallbacks.push(callback);
            }
        }
        return _officeOnReadyPromise;
    }
    Office.onReadyInternal = onReadyInternal;
    function onReady(callback) {
        _isOfficeOnReadyCalled = true;
        return onReadyInternal(callback);
    }
    Office.onReady = onReady;
    function fireOnReady(hostAndPlatformInfo) {
        ensureOfficeOnReadyPromise();
        _officeOnReadyHostAndPlatformInfo = __assign({}, hostAndPlatformInfo);
        _officeOnReadyFired = true;
        OSFPerformance.officeOnReady = OSFPerformance.now();
        while (_officeOnReadyCallbacks.length > 0) {
            _officeOnReadyCallbacks.shift()(_officeOnReadyHostAndPlatformInfo);
        }
        _officeOnReadyPromiseResolve(_officeOnReadyHostAndPlatformInfo);
    }
    Office.fireOnReady = fireOnReady;
    function sendTelemetryEvent(telemetryEvent) {
        Microsoft.Office.WebExtension.sendTelemetryEvent(telemetryEvent);
    }
    Office.sendTelemetryEvent = sendTelemetryEvent;
})(Office || (Office = {}));
var OSF;
(function (OSF) {
    var OfficeAppContext = (function () {
        function OfficeAppContext(id, appName, appVersion, appUILocale, dataLocale, docUrl, clientMode, settingsFunc, reason, osfControlType, eToken, correlationId, appInstanceId, touchEnabled, commerceAllowed, appMinorVersion, requirementMatrix, hostCustomMessage, hostFullVersion, clientWindowHeight, clientWindowWidth, addinName, appDomains, dialogRequirementMatrix, featureGates, officeThemeFunc, initialDisplayMode) {
            this._isDialog = false;
            this._id = id;
            this._appName = appName;
            this._appVersion = appVersion;
            this._appUILocale = appUILocale;
            this._dataLocale = dataLocale;
            this._docUrl = docUrl;
            this._clientMode = clientMode;
            this._settingsFunc = settingsFunc;
            this._reason = reason;
            this._osfControlType = osfControlType;
            this._eToken = eToken;
            this._correlationId = correlationId;
            this._appInstanceId = appInstanceId;
            this._touchEnabled = touchEnabled;
            this._commerceAllowed = commerceAllowed;
            this._appMinorVersion = appMinorVersion;
            this._requirementMatrix = requirementMatrix;
            this._hostCustomMessage = hostCustomMessage;
            this._hostFullVersion = hostFullVersion;
            this._isDialog = false;
            this._clientWindowHeight = clientWindowHeight;
            this._clientWindowWidth = clientWindowWidth;
            this._addinName = addinName;
            this._appDomains = appDomains;
            this._dialogRequirementMatrix = dialogRequirementMatrix;
            this._featureGates = featureGates;
            this._officeThemeFunc = officeThemeFunc;
            this._initialDisplayMode = initialDisplayMode;
        }
        OfficeAppContext.prototype.get_id = function () {
            return this._id;
        };
        OfficeAppContext.prototype.get_appName = function () {
            return this._appName;
        };
        OfficeAppContext.prototype.get_appVersion = function () {
            return this._appVersion;
        };
        OfficeAppContext.prototype.get_appUILocale = function () {
            return this._appUILocale;
        };
        OfficeAppContext.prototype.get_dataLocale = function () { return this._dataLocale; };
        OfficeAppContext.prototype.get_docUrl = function () { return this._docUrl; };
        OfficeAppContext.prototype.get_clientMode = function () { return this._clientMode; };
        OfficeAppContext.prototype.get_settingsFunc = function () { return this._settingsFunc; };
        OfficeAppContext.prototype.get_reason = function () { return this._reason; };
        OfficeAppContext.prototype.get_osfControlType = function () { return this._osfControlType; };
        OfficeAppContext.prototype.get_eToken = function () { return this._eToken; };
        OfficeAppContext.prototype.get_correlationId = function () { return this._correlationId; };
        OfficeAppContext.prototype.get_appInstanceId = function () { return this._appInstanceId; };
        OfficeAppContext.prototype.get_touchEnabled = function () { return this._touchEnabled; };
        OfficeAppContext.prototype.get_commerceAllowed = function () { return this._commerceAllowed; };
        OfficeAppContext.prototype.get_appMinorVersion = function () { return this._appMinorVersion; };
        OfficeAppContext.prototype.get_requirementMatrix = function () { return this._requirementMatrix; };
        OfficeAppContext.prototype.get_dialogRequirementMatrix = function () { return this._dialogRequirementMatrix; };
        OfficeAppContext.prototype.get_hostCustomMessage = function () { return this._hostCustomMessage; };
        OfficeAppContext.prototype.get_hostFullVersion = function () { return this._hostFullVersion; };
        OfficeAppContext.prototype.get_isDialog = function () { return this._isDialog; };
        OfficeAppContext.prototype.get_clientWindowHeight = function () { return this._clientWindowHeight; };
        OfficeAppContext.prototype.get_clientWindowWidth = function () { return this._clientWindowWidth; };
        OfficeAppContext.prototype.get_addinName = function () { return this._addinName; };
        OfficeAppContext.prototype.get_appDomains = function () { return this._appDomains; };
        OfficeAppContext.prototype.get_featureGates = function () { return this._featureGates; };
        OfficeAppContext.prototype.get_officeThemeFunc = function () { return this._officeThemeFunc; };
        OfficeAppContext.prototype.get_initialDisplayMode = function () { return this._initialDisplayMode ? this._initialDisplayMode : 0; };
        return OfficeAppContext;
    }());
    OSF.OfficeAppContext = OfficeAppContext;
})(OSF || (OSF = {}));
var OSF;
(function (OSF) {
    var _OfficeAppFactory;
    (function (_OfficeAppFactory) {
        var _windowLocationHash = window.location.hash;
        var _windowLocationSearch = window.location.search;
        var _windowName = window.name;
        var _hostInfo;
        var _webAppState;
        var _isLoggingAllowed;
        var _initializationHelper;
        var _asyncMethodExecutor;
        var _officeAppContext;
        var _initialDisplayModeMappings = {
            0: "Unknown",
            1: "Hidden",
            2: "Taskpane",
            3: "Dialog"
        };
        function bootstrap(onSuccess, onError) {
            _webAppState = {
                id: null,
                webAppUrl: null,
                conversationID: null,
                clientEndPoint: null,
                wnd: window.parent,
                focused: false,
                serviceEndPoint: null,
                serializerVersion: 1
            };
            retrieveHostInfo();
            retrieveLoggingAllowed();
            createInitializationHelper();
            if (_hostInfo.hostPlatform === OSF.HostInfoPlatform.web) {
                _initializationHelper.saveAndSetDialogInfo(OSF.Utility.getQueryStringValue("_host_Info"));
            }
            _initializationHelper.setAgaveHostCommunication();
            OSFPerformance.getAppContextStart = OSFPerformance.now();
            var onGetAppContextSuccess = function (officeAppContext) {
                OSFPerformance.getAppContextEnd = OSFPerformance.now();
                OSF.AppTelemetry.initialize(officeAppContext);
                _officeAppContext = officeAppContext;
                _initializationHelper.createClientHostController();
                _asyncMethodExecutor = _initializationHelper.createAsyncMethodExecutor();
                _initializationHelper.prepareApiSurface(officeAppContext);
                var appNameNumber = officeAppContext.get_appName();
                var addinInfo = null;
                if ((_hostInfo.flags & 1) !== 0) {
                    addinInfo = {
                        visibilityMode: _initialDisplayModeMappings[officeAppContext.get_initialDisplayMode()]
                    };
                }
                Office.fireOnReady({
                    host: OSF.HostName.Host.getInstance().getHost(appNameNumber),
                    platform: OSF.HostName.Host.getInstance().getPlatform(appNameNumber),
                    addin: addinInfo
                });
                onSuccess(officeAppContext);
            };
            var onGetAppContextError = function (e) {
                onError(e);
            };
            _initializationHelper.getAppContext(window, onGetAppContextSuccess, onGetAppContextError);
        }
        _OfficeAppFactory.bootstrap = bootstrap;
        function retrieveHostInfo() {
            _hostInfo = {
                isO15: true,
                isRichClient: true,
                hostType: "",
                hostPlatform: "",
                hostSpecificFileVersion: "",
                hostLocale: "",
                osfControlAppCorrelationId: "",
                isDialog: false,
                disableLogging: false,
                flags: 0
            };
            var hostInfoParaName = "_host_Info";
            var hostInfoValue = OSF.Utility.getQueryStringValue(hostInfoParaName);
            if (!hostInfoValue) {
                try {
                    var windowNameObj = JSON.parse(window.name);
                    hostInfoValue = windowNameObj ? windowNameObj["hostInfo"] : null;
                }
                catch (ex) {
                    OSF.Utility.log(JSON.stringify(ex));
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
                    if (window.external.GetHostInfo) {
                        var fallbackHostInfo = window.external.GetHostInfo();
                        if (fallbackHostInfo == "isDialog") {
                            _hostInfo.isO15 = true;
                            _hostInfo.isDialog = true;
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
                }
                catch (ex) {
                    OSF.Utility.log(JSON.stringify(ex));
                }
            }
            var osfSessionStorage = OSF.OUtil.getSessionStorage();
            if (!hostInfoValue && osfSessionStorage.getItem("hostInfoValue")) {
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
                osfSessionStorage.setItem("hostInfoValue", hostInfoValue);
            }
            else {
                _hostInfo.isO15 = true;
                _hostInfo.hostLocale = OSF.Utility.getQueryStringValue("locale");
            }
        }
        function retrieveLoggingAllowed() {
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
            catch (ex) {
            }
        }
        function createInitializationHelper() {
            if (_hostInfo.hostPlatform === OSF.HostInfoPlatform.web) {
                _initializationHelper = new OSF.WebInitializationHelper(_hostInfo, _webAppState, null, null);
            }
            else if (_hostInfo.hostPlatform === OSF.HostInfoPlatform.win32) {
                _initializationHelper = new OSF.RichClientInitializationHelper(_hostInfo, _webAppState, null, null);
            }
            else if (_hostInfo.hostPlatform === OSF.HostInfoPlatform.ios || _hostInfo.hostPlatform === OSF.HostInfoPlatform.mac) {
                if (isWebkit2Sandbox()) {
                    _initializationHelper = new OSF.WebkitInitializationHelper(_hostInfo, _webAppState, null, null);
                }
                else {
                    throw OSF.Utility.createNotImplementedException();
                }
            }
        }
        function isWebkit2Sandbox() {
            return window.webkit && window.webkit.messageHandlers && window.webkit.messageHandlers.Agave;
        }
        function getWindowName() {
            return _windowName;
        }
        _OfficeAppFactory.getWindowName = getWindowName;
        function getWindowLocationHash() {
            return _windowLocationHash;
        }
        _OfficeAppFactory.getWindowLocationHash = getWindowLocationHash;
        function getWindowLocationSearch() {
            return _windowLocationSearch;
        }
        _OfficeAppFactory.getWindowLocationSearch = getWindowLocationSearch;
        function getAsyncMethodExecutor() {
            return _asyncMethodExecutor;
        }
        _OfficeAppFactory.getAsyncMethodExecutor = getAsyncMethodExecutor;
        function getOfficeAppContext() {
            return _officeAppContext;
        }
        _OfficeAppFactory.getOfficeAppContext = getOfficeAppContext;
        function getHostInfo() {
            return _hostInfo;
        }
        _OfficeAppFactory.getHostInfo = getHostInfo;
        function getCachedSessionSettingsKey() {
            return (_webAppState.conversationID != null ? _webAppState.conversationID : _officeAppContext.get_appInstanceId()) + "CachedSessionSettings";
        }
        _OfficeAppFactory.getCachedSessionSettingsKey = getCachedSessionSettingsKey;
        function getWebAppState() {
            return _webAppState;
        }
        _OfficeAppFactory.getWebAppState = getWebAppState;
        function getId() {
            return _webAppState.id;
        }
        _OfficeAppFactory.getId = getId;
        function getInitializationHelper() {
            return _initializationHelper;
        }
        _OfficeAppFactory.getInitializationHelper = getInitializationHelper;
    })(_OfficeAppFactory = OSF._OfficeAppFactory || (OSF._OfficeAppFactory = {}));
    function getClientEndPoint() {
        return _OfficeAppFactory.getWebAppState().clientEndPoint;
    }
    OSF.getClientEndPoint = getClientEndPoint;
})(OSF || (OSF = {}));
var Office;
(function (Office) {
    var VisibilityMode;
    (function (VisibilityMode) {
        VisibilityMode["hidden"] = "Hidden";
        VisibilityMode["taskpane"] = "Taskpane";
    })(VisibilityMode = Office.VisibilityMode || (Office.VisibilityMode = {}));
    var AsyncResultStatus;
    (function (AsyncResultStatus) {
        AsyncResultStatus["succeeded"] = "succeeded";
        AsyncResultStatus["failed"] = "failed";
    })(AsyncResultStatus = Office.AsyncResultStatus || (Office.AsyncResultStatus = {}));
    var DocumentMode;
    (function (DocumentMode) {
        DocumentMode["ReadOnly"] = "readOnly";
        DocumentMode["ReadWrite"] = "readWrite";
    })(DocumentMode = Office.DocumentMode || (Office.DocumentMode = {}));
    var HostType;
    (function (HostType) {
        HostType["Word"] = "Word";
        HostType["Excel"] = "Excel";
        HostType["PowerPoint"] = "PowerPoint";
        HostType["Outlook"] = "Outlook";
        HostType["OneNote"] = "OneNote";
        HostType["Project"] = "Project";
        HostType["Access"] = "Access";
        HostType["Visio"] = "Visio";
    })(HostType = Office.HostType || (Office.HostType = {}));
    var InitializationReason;
    (function (InitializationReason) {
        InitializationReason["Inserted"] = "inserted";
        InitializationReason["DocumentOpened"] = "documentOpened";
    })(InitializationReason = Office.InitializationReason || (Office.InitializationReason = {}));
    var PlatformType;
    (function (PlatformType) {
        PlatformType["PC"] = "PC";
        PlatformType["OfficeOnline"] = "OfficeOnline";
        PlatformType["Mac"] = "Mac";
        PlatformType["iOS"] = "iOS";
        PlatformType["Android"] = "Android";
        PlatformType["Universal"] = "Universal";
    })(PlatformType = Office.PlatformType || (Office.PlatformType = {}));
})(Office || (Office = {}));
var OfficeExt;
(function (OfficeExt) {
    function appSpecificCheckOriginFunction(allowed_domains, eventObj, origin, checkOriginFunction) {
        return false;
    }
    ;
    OfficeExt.appSpecificCheckOrigin = appSpecificCheckOriginFunction;
})(OfficeExt || (OfficeExt = {}));
var OSF;
(function (OSF) {
    ;
    var XdmMethodObject = (function () {
        function XdmMethodObject(method, invokeType, blockingOthers) {
            this._method = method;
            this._invokeType = invokeType;
            this._blockingOthers = blockingOthers;
        }
        XdmMethodObject.prototype.getMethod = function () {
            return this._method;
        };
        XdmMethodObject.prototype.getInvokeType = function () {
            return this._invokeType;
        };
        XdmMethodObject.prototype.getBlockingFlag = function () {
            return this._blockingOthers;
        };
        return XdmMethodObject;
    }());
    OSF.XdmMethodObject = XdmMethodObject;
    var XdmEventMethodObject = (function () {
        function XdmEventMethodObject(registerMethodObject, unregisterMethodObject) {
            this._registerMethodObject = registerMethodObject;
            this._unregisterMethodObject = unregisterMethodObject;
        }
        XdmEventMethodObject.prototype.getRegisterMethodObject = function () {
            return this._registerMethodObject;
        };
        XdmEventMethodObject.prototype.getUnregisterMethodObject = function () {
            return this._unregisterMethodObject;
        };
        return XdmEventMethodObject;
    }());
    OSF.XdmEventMethodObject = XdmEventMethodObject;
    ;
    var XdmServiceEndPoint = (function () {
        function XdmServiceEndPoint(serviceEndPointId) {
            this._methodObjectList = {};
            this._eventHandlerProxyList = {};
            this._conversations = {};
            this._policyManager = null;
            this._appDomains = {};
            this._onHandleRequestError = null;
            this._methodObjectList = {};
            this._eventHandlerProxyList = {};
            this._Id = serviceEndPointId;
            this._conversations = {};
            this._policyManager = null;
            this._appDomains = {};
            this._onHandleRequestError = null;
        }
        XdmServiceEndPoint.prototype.registerMethod = function (methodName, method, invokeType, blockingOthers) {
            if (invokeType !== 0
                && invokeType !== 1) {
                throw OSF.Utility.createArgumentException("invokeType");
            }
            var methodObject = new XdmMethodObject(method, invokeType, blockingOthers);
            this._methodObjectList[methodName] = methodObject;
        };
        XdmServiceEndPoint.prototype.unregisterMethod = function (methodName) {
            delete this._methodObjectList[methodName];
        };
        XdmServiceEndPoint.prototype.registerEvent = function (eventName, registerMethod, unregisterMethod) {
            var methodObject = new XdmEventMethodObject(new XdmMethodObject(registerMethod, 4, false), new XdmMethodObject(unregisterMethod, 5, false));
            this._methodObjectList[eventName] = methodObject;
        };
        XdmServiceEndPoint.prototype.registerEventEx = function (eventName, registerMethod, registerMethodInvokeType, unregisterMethod, unregisterMethodInvokeType) {
            var methodObject = new XdmEventMethodObject(new XdmMethodObject(registerMethod, registerMethodInvokeType, false), new XdmMethodObject(unregisterMethod, unregisterMethodInvokeType, false));
            this._methodObjectList[eventName] = methodObject;
        };
        XdmServiceEndPoint.prototype.unregisterEvent = function (eventName) {
            this.unregisterMethod(eventName);
        };
        XdmServiceEndPoint.prototype.registerConversation = function (conversationId, conversationUrl, appDomains, serializerVersion) {
            OSF.Utility.log("registerConversation: cId=" + conversationId + " Url=" + conversationUrl);
            if (appDomains) {
                if (!Array.isArray(appDomains)) {
                    throw OSF.Utility.createArgumentException("appDomains");
                }
                this._appDomains[conversationId] = appDomains;
            }
            this._conversations[conversationId] = { url: conversationUrl, serializerVersion: serializerVersion };
        };
        XdmServiceEndPoint.prototype.unregisterConversation = function (conversationId) {
            delete this._conversations[conversationId];
        };
        XdmServiceEndPoint.prototype.setPolicyManager = function (policyManager) {
            if (!policyManager.checkPermission) {
                throw OSF.Utility.createArgumentException("policyManager");
            }
            this._policyManager = policyManager;
        };
        XdmServiceEndPoint.prototype.getPolicyManager = function () {
            return this._policyManager;
        };
        XdmServiceEndPoint.prototype.dispose = function () {
            this._methodObjectList = null;
            this._eventHandlerProxyList = null;
            this._Id = null;
            this._conversations = null;
            this._policyManager = null;
            this._appDomains = null;
            this._onHandleRequestError = null;
        };
        return XdmServiceEndPoint;
    }());
    OSF.XdmServiceEndPoint = XdmServiceEndPoint;
    var XdmClientEndPoint = (function () {
        function XdmClientEndPoint(conversationId, targetWindow, targetUrl, serializerVersion) {
            this._callbackList = {};
            this._eventHandlerList = {};
            this._conversationId = conversationId;
            this._targetWindow = targetWindow;
            this._targetUrl = targetUrl;
            this._callingIndex = 0;
            this._callbackList = {};
            this._eventHandlerList = {};
            if (serializerVersion != null) {
                this._serializerVersion = serializerVersion;
            }
            else {
                this._serializerVersion = 1;
            }
            this._onSenderOriginNotTrusted = null;
        }
        ;
        XdmClientEndPoint.prototype.invoke = function (targetMethodName, callback, param) {
            var correlationId = this._callingIndex++;
            var now = new Date();
            var callbackEntry = { "callback": callback, "createdOn": now.getTime() };
            if (param && typeof param === "object" && typeof param.__timeout__ === "number") {
                callbackEntry.timeout = param.__timeout__;
                delete param.__timeout__;
            }
            this._callbackList[correlationId] = callbackEntry;
            try {
                var callRequest = new XdmRequest(targetMethodName, 0, this._conversationId, correlationId, param);
                var msg = XdmMessagePackager.envelope(callRequest, this._serializerVersion);
                this._targetWindow.postMessage(msg, this._targetUrl);
                XdmCommunicationManager._startMethodTimeoutTimer();
            }
            catch (ex) {
                try {
                    if (callback !== null)
                        callback(-1, ex);
                }
                finally {
                    delete this._callbackList[correlationId];
                }
            }
        };
        XdmClientEndPoint.prototype.registerForEvent = function (targetEventName, eventHandler, callback, data) {
            var correlationId = this._callingIndex++;
            var now = new Date();
            this._callbackList[correlationId] = { "callback": callback, "createdOn": now.getTime() };
            try {
                var callRequest = new XdmRequest(targetEventName, 1, this._conversationId, correlationId, data);
                var msg = XdmMessagePackager.envelope(callRequest, this._serializerVersion);
                this._targetWindow.postMessage(msg, this._targetUrl);
                XdmCommunicationManager._startMethodTimeoutTimer();
                this._eventHandlerList[targetEventName] = eventHandler;
            }
            catch (ex) {
                try {
                    if (callback !== null) {
                        callback(-1, ex);
                    }
                }
                finally {
                    delete this._callbackList[correlationId];
                }
            }
        };
        XdmClientEndPoint.prototype.unregisterForEvent = function (targetEventName, callback, data) {
            var correlationId = this._callingIndex++;
            var now = new Date();
            this._callbackList[correlationId] = { "callback": callback, "createdOn": now.getTime() };
            try {
                var callRequest = new XdmRequest(targetEventName, 2, this._conversationId, correlationId, data);
                var msg = XdmMessagePackager.envelope(callRequest, this._serializerVersion);
                this._targetWindow.postMessage(msg, this._targetUrl);
                XdmCommunicationManager._startMethodTimeoutTimer();
            }
            catch (ex) {
                try {
                    if (callback !== null) {
                        callback(-1, ex);
                    }
                }
                finally {
                    delete this._callbackList[correlationId];
                }
            }
            finally {
                delete this._eventHandlerList[targetEventName];
            }
        };
        return XdmClientEndPoint;
    }());
    OSF.XdmClientEndPoint = XdmClientEndPoint;
    ;
    var XdmCommunicationManager;
    (function (XdmCommunicationManager) {
        var _invokerQueue = [];
        var _lastMessageProcessTime = null;
        var _messageProcessingTimer = null;
        var _processInterval = 10;
        var _blockingFlag = false;
        var _methodTimeoutTimer = null;
        var _methodTimeoutProcessInterval = 2000;
        var _methodTimeoutDefault = 65000;
        var _methodTimeout = _methodTimeoutDefault;
        var _serviceEndPoints = {};
        var _clientEndPoints = {};
        var _initialized = false;
        function _lookupServiceEndPoint(conversationId) {
            for (var id in _serviceEndPoints) {
                if (_serviceEndPoints[id]._conversations[conversationId]) {
                    return _serviceEndPoints[id];
                }
            }
            throw OSF.Utility.createArgumentException("conversationId");
        }
        ;
        function _lookupClientEndPoint(conversationId) {
            var clientEndPoint = _clientEndPoints[conversationId];
            if (!clientEndPoint) {
                OSF.Utility.log("Unknown conversation Id.");
            }
            return clientEndPoint;
        }
        ;
        function _lookupMethodObject(serviceEndPoint, messageObject) {
            var methodOrEventMethodObject = serviceEndPoint._methodObjectList[messageObject._actionName];
            if (!methodOrEventMethodObject) {
                OSF.Utility.log("The specified method is not registered on service endpoint:" + messageObject._actionName);
                throw OSF.Utility.createArgumentException("messageObject");
            }
            var methodObject = null;
            if (messageObject._actionType === 0) {
                methodObject = methodOrEventMethodObject;
            }
            else if (messageObject._actionType === 1) {
                methodObject = methodOrEventMethodObject.getRegisterMethodObject();
            }
            else {
                methodObject = methodOrEventMethodObject.getUnregisterMethodObject();
            }
            return methodObject;
        }
        ;
        function _enqueInvoker(invoker) {
            _invokerQueue.push(invoker);
        }
        ;
        function _dequeInvoker() {
            if (_messageProcessingTimer !== null) {
                if (!_blockingFlag) {
                    if (_invokerQueue.length > 0) {
                        var invoker = _invokerQueue.shift();
                        _executeCommand(invoker);
                    }
                    else {
                        clearInterval(_messageProcessingTimer);
                        _messageProcessingTimer = null;
                    }
                }
            }
            else {
                OSF.Utility.log("channel is not ready.");
            }
        }
        ;
        function _executeCommand(invoker) {
            _blockingFlag = invoker.getInvokeBlockingFlag();
            invoker.invoke();
            _lastMessageProcessTime = (new Date()).getTime();
        }
        ;
        function _checkMethodTimeout() {
            if (_methodTimeoutTimer) {
                var clientEndPoint;
                var methodCallsNotTimedout = 0;
                var now = new Date();
                var timeoutValue;
                for (var conversationId in _clientEndPoints) {
                    clientEndPoint = _clientEndPoints[conversationId];
                    for (var correlationId in clientEndPoint._callbackList) {
                        var callbackEntry = clientEndPoint._callbackList[correlationId];
                        timeoutValue = callbackEntry.timeout ? callbackEntry.timeout : _methodTimeout;
                        if (timeoutValue >= 0 && Math.abs(now.getTime() - callbackEntry.createdOn) >= timeoutValue) {
                            try {
                                if (callbackEntry.callback) {
                                    callbackEntry.callback(-6, null);
                                }
                            }
                            finally {
                                delete clientEndPoint._callbackList[correlationId];
                            }
                        }
                        else {
                            methodCallsNotTimedout++;
                        }
                        ;
                    }
                }
                if (methodCallsNotTimedout === 0) {
                    clearInterval(_methodTimeoutTimer);
                    _methodTimeoutTimer = null;
                }
            }
            else {
                OSF.Utility.log("channel is not ready.");
            }
        }
        ;
        function _postCallbackHandler() {
            _blockingFlag = false;
        }
        ;
        function _registerListener(listener) {
            if (window.addEventListener) {
                window.addEventListener("message", listener, false);
            }
            else if ((navigator.userAgent.indexOf("MSIE") > -1) && window.attachEvent) {
                window.attachEvent("onmessage", listener);
            }
            else {
                OSF.Utility.log("Browser doesn't support the required API.");
                throw OSF.Utility.createArgumentException("Browser");
            }
        }
        ;
        function _checkOrigin(url, origin) {
            var res = false;
            if (!url || !origin || url === "null" || origin === "null" || !url.length || !origin.length) {
                return res;
            }
            var url_parser, org_parser;
            url_parser = document.createElement('a');
            org_parser = document.createElement('a');
            url_parser.href = url;
            org_parser.href = origin;
            res = _urlCompare(url_parser, org_parser);
            return res;
        }
        function _checkOriginWithAppDomains(allowed_domains, origin) {
            var res = false;
            if (!origin || origin === "null" || !origin.length || !(allowed_domains) || !(allowed_domains instanceof Array) || !allowed_domains.length) {
                return res;
            }
            var org_parser = document.createElement('a');
            var app_domain_parser = document.createElement('a');
            org_parser.href = origin;
            for (var i = 0; i < allowed_domains.length && !res; i++) {
                if (allowed_domains[i].indexOf("://") !== -1) {
                    app_domain_parser.href = allowed_domains[i];
                    res = _urlCompare(org_parser, app_domain_parser);
                }
            }
            return res;
        }
        function _isHostNameValidWacDomain(hostName) {
            if (!hostName || hostName === "null") {
                return false;
            }
            var regexHostNameStringArray = new Array("^office-int\\.com$", "^officeapps\\.live-int\\.com$", "^.*\\.dod\\.online\\.office365\\.us$", "^.*\\.gov\\.online\\.office365\\.us$", "^.*\\.officeapps\\.live\\.com$", "^.*\\.officeapps\\.live-int\\.com$", "^.*\\.officeapps-df\\.live\\.com$", "^.*\\.online\\.office\\.de$", "^.*\\.partner\\.officewebapps\\.cn$", "^" + document.domain.replace(new RegExp("\\.", "g"), "\\.") + "$");
            var regexHostName = new RegExp(regexHostNameStringArray.join("|"));
            return regexHostName.test(hostName);
        }
        function _isTargetSubdomainOfSourceLocation(sourceLocation, messageOrigin) {
            if (!sourceLocation || !messageOrigin || sourceLocation === "null" || messageOrigin === "null") {
                return false;
            }
            var sourceLocationParser = document.createElement('a');
            sourceLocationParser.href = sourceLocation;
            var messageOriginParser = document.createElement('a');
            messageOriginParser.href = messageOrigin;
            var isSameProtocol = sourceLocationParser.protocol === messageOriginParser.protocol;
            var isSamePort = sourceLocationParser.port === messageOriginParser.port;
            var originHostName = messageOriginParser.hostname;
            var sourceLocationHostName = sourceLocationParser.hostname;
            var isSameDomain = originHostName === sourceLocationHostName;
            var isSubDomain = false;
            if (!isSameDomain && originHostName.length > sourceLocationHostName.length + 1) {
                isSubDomain = originHostName.slice(-(sourceLocationHostName.length + 1)) === '.' + sourceLocationHostName;
            }
            var isSameDomainOrSubdomain = isSameDomain || isSubDomain;
            return isSamePort && isSameProtocol && isSameDomainOrSubdomain;
        }
        function _urlCompare(url_parser1, url_parser2) {
            return ((url_parser1.hostname == url_parser2.hostname) &&
                (url_parser1.protocol == url_parser2.protocol) &&
                (url_parser1.port == url_parser2.port));
        }
        function _receive(e) {
            if (e.data != '') {
                var messageObject;
                var serializerVersion = 1;
                var serializedMessage = e.data;
                try {
                    messageObject = XdmMessagePackager.unenvelope(serializedMessage, 1);
                    serializerVersion = messageObject._serializerVersion != null ? messageObject._serializerVersion : serializerVersion;
                }
                catch (ex) {
                    return;
                }
                OSF.Utility.debugLog(serializedMessage);
                if (messageObject._messageType === 0) {
                    var requesterUrl = (e.origin == null || e.origin === "null") ? messageObject._origin : e.origin;
                    try {
                        var serviceEndPoint = _lookupServiceEndPoint(messageObject._conversationId);
                        OSF.Utility.log("_receive: request, origin=" + requesterUrl + " sourceURL:" + serviceEndPoint._conversations[messageObject._conversationId]);
                        var conversation = serviceEndPoint._conversations[messageObject._conversationId];
                        serializerVersion = conversation.serializerVersion != null ? conversation.serializerVersion : serializerVersion;
                        OSF.Utility.log("_receive: request, origin=" + requesterUrl + " sourceURL:" + conversation.url);
                        var allowedDomains = [conversation.url].concat(serviceEndPoint._appDomains[messageObject._conversationId]);
                        if (!_checkOriginWithAppDomains(allowedDomains, e.origin)) {
                            if (!OfficeExt.appSpecificCheckOrigin(allowedDomains, e, messageObject._origin, _checkOriginWithAppDomains)) {
                                var isOriginSubdomain = _isTargetSubdomainOfSourceLocation(conversation.url, e.origin);
                                if (!isOriginSubdomain) {
                                    throw "Failed origin check";
                                }
                            }
                        }
                        var policyManager = serviceEndPoint.getPolicyManager();
                        if (policyManager && !policyManager.checkPermission(messageObject._conversationId, messageObject._actionName, messageObject._data)) {
                            throw "Access Denied";
                        }
                        var methodObject = _lookupMethodObject(serviceEndPoint, messageObject);
                        var invokeCompleteCallback = new XdmInvokeCompleteCallback(e.source, requesterUrl, messageObject._actionName, messageObject._conversationId, messageObject._correlationId, _postCallbackHandler, serializerVersion);
                        var invoker = new XdmInvoker(methodObject, messageObject._data, invokeCompleteCallback, serviceEndPoint._eventHandlerProxyList, messageObject._conversationId, messageObject._actionName, serializerVersion);
                        var shouldEnque = true;
                        if (_messageProcessingTimer == null) {
                            if ((_lastMessageProcessTime == null || (new Date()).getTime() - _lastMessageProcessTime > _processInterval) && !_blockingFlag) {
                                _executeCommand(invoker);
                                shouldEnque = false;
                            }
                            else {
                                _messageProcessingTimer = setInterval(_dequeInvoker, _processInterval);
                            }
                        }
                        if (shouldEnque) {
                            _enqueInvoker(invoker);
                        }
                    }
                    catch (ex) {
                        if (serviceEndPoint && serviceEndPoint._onHandleRequestError) {
                            serviceEndPoint._onHandleRequestError(messageObject, ex);
                        }
                        var errorCode = -2;
                        if (ex == "Access Denied") {
                            errorCode = -5;
                        }
                        var callResponse = new XdmResponse(messageObject._actionName, messageObject._conversationId, messageObject._correlationId, errorCode, 0, ex);
                        var envelopedResult = XdmMessagePackager.envelope(callResponse, serializerVersion);
                        var canPostMessage = false;
                        try {
                            canPostMessage = !!(e.source && e.source.postMessage);
                        }
                        catch (ex) {
                        }
                        var isOriginValid = false;
                        if (window.location.href && e.origin && e.origin !== "null" && _isTargetSubdomainOfSourceLocation(window.location.href, e.origin)) {
                            isOriginValid = true;
                        }
                        else {
                            if (e.origin && e.origin !== "null") {
                                var parser = document.createElement("a");
                                parser.href = e.origin;
                                isOriginValid = _isHostNameValidWacDomain(parser.hostname);
                            }
                        }
                        if (canPostMessage && isOriginValid) {
                            e.source.postMessage(envelopedResult, requesterUrl);
                        }
                    }
                }
                else if (messageObject._messageType === 1) {
                    var clientEndPoint = _lookupClientEndPoint(messageObject._conversationId);
                    if (messageObject._actionName == "ContextActivationManager_getAppContextAsync") {
                        try {
                            var wacorigin = e.origin;
                            var parser = document.createElement("a");
                            parser.href = wacorigin;
                            var isOriginValid = _isHostNameValidWacDomain(parser.hostname);
                            var isWacKnownHost = isOriginValid ? 1 : 0;
                            if (!isWacKnownHost) {
                                if (clientEndPoint && clientEndPoint._onSenderOriginNotTrusted) {
                                    clientEndPoint._onSenderOriginNotTrusted();
                                }
                            }
                        }
                        catch (ex) {
                        }
                    }
                    if (!clientEndPoint) {
                        return;
                    }
                    clientEndPoint._serializerVersion = serializerVersion;
                    OSF.Utility.log("_receive: response, origin=" + e.origin + " targetURL:" + clientEndPoint._targetUrl);
                    if (!_checkOrigin(clientEndPoint._targetUrl, e.origin)) {
                        throw "Failed orgin check";
                    }
                    if (messageObject._responseType === 0) {
                        var callbackEntry = clientEndPoint._callbackList[messageObject._correlationId];
                        if (callbackEntry) {
                            try {
                                if (callbackEntry.callback)
                                    callbackEntry.callback(messageObject._errorCode, messageObject._data);
                            }
                            finally {
                                delete clientEndPoint._callbackList[messageObject._correlationId];
                            }
                        }
                    }
                    else {
                        var eventhandler = clientEndPoint._eventHandlerList[messageObject._actionName];
                        if (eventhandler !== undefined && eventhandler !== null) {
                            eventhandler(messageObject._data);
                        }
                    }
                }
                else {
                    return;
                }
            }
        }
        ;
        function _initialize() {
            if (!_initialized) {
                _registerListener(_receive);
                _initialized = true;
            }
        }
        ;
        function connect(conversationId, targetWindow, targetUrl, serializerVersion) {
            var clientEndPoint = _clientEndPoints[conversationId];
            if (!clientEndPoint) {
                _initialize();
                clientEndPoint = new XdmClientEndPoint(conversationId, targetWindow, targetUrl, serializerVersion);
                _clientEndPoints[conversationId] = clientEndPoint;
            }
            return clientEndPoint;
        }
        XdmCommunicationManager.connect = connect;
        function getClientEndPoint(conversationId) {
            return _clientEndPoints[conversationId];
        }
        XdmCommunicationManager.getClientEndPoint = getClientEndPoint;
        function createServiceEndPoint(serviceEndPointId) {
            _initialize();
            var serviceEndPoint = new XdmServiceEndPoint(serviceEndPointId);
            _serviceEndPoints[serviceEndPointId] = serviceEndPoint;
            return serviceEndPoint;
        }
        XdmCommunicationManager.createServiceEndPoint = createServiceEndPoint;
        function getServiceEndPoint(serviceEndPointId) {
            return _serviceEndPoints[serviceEndPointId];
        }
        XdmCommunicationManager.getServiceEndPoint = getServiceEndPoint;
        function deleteClientEndPoint(conversationId) {
            delete _clientEndPoints[conversationId];
        }
        XdmCommunicationManager.deleteClientEndPoint = deleteClientEndPoint;
        function deleteServiceEndPoint(serviceEndPointId) {
            delete _serviceEndPoints[serviceEndPointId];
        }
        XdmCommunicationManager.deleteServiceEndPoint = deleteServiceEndPoint;
        function checkUrlWithAppDomains(appDomains, origin) {
            return _checkOriginWithAppDomains(appDomains, origin);
        }
        XdmCommunicationManager.checkUrlWithAppDomains = checkUrlWithAppDomains;
        ;
        function isTargetSubdomainOfSourceLocation(sourceLocation, messageOrigin) {
            return _isTargetSubdomainOfSourceLocation(sourceLocation, messageOrigin);
        }
        XdmCommunicationManager.isTargetSubdomainOfSourceLocation = isTargetSubdomainOfSourceLocation;
        function _setMethodTimeout(methodTimeout) {
            _methodTimeout = (methodTimeout <= 0) ? _methodTimeoutDefault : methodTimeout;
        }
        XdmCommunicationManager._setMethodTimeout = _setMethodTimeout;
        function _startMethodTimeoutTimer() {
        }
        XdmCommunicationManager._startMethodTimeoutTimer = _startMethodTimeoutTimer;
    })(XdmCommunicationManager = OSF.XdmCommunicationManager || (OSF.XdmCommunicationManager = {}));
    var XdmMessage = (function () {
        function XdmMessage(messageType, actionName, conversationId, correlationId, data) {
            this._messageType = messageType;
            this._actionName = actionName;
            this._conversationId = conversationId;
            this._correlationId = correlationId;
            this._origin = window.location.href;
            if (typeof data === "undefined") {
                this._data = null;
            }
            else {
                this._data = data;
            }
        }
        XdmMessage.prototype.getActionName = function () {
            return this._actionName;
        };
        XdmMessage.prototype.getConversationId = function () {
            return this._conversationId;
        };
        XdmMessage.prototype.getCorrelationId = function () {
            return this._correlationId;
        };
        XdmMessage.prototype.getOrigin = function () {
            return this._origin;
        };
        XdmMessage.prototype.getData = function () {
            return this._data;
        };
        XdmMessage.prototype.getMessageType = function () {
            return this._messageType;
        };
        return XdmMessage;
    }());
    var XdmRequest = (function (_super) {
        __extends(XdmRequest, _super);
        function XdmRequest(actionName, actionType, conversationId, correlationId, data) {
            var _this = _super.call(this, 0, actionName, conversationId, correlationId, data) || this;
            _this._actionType = actionType;
            return _this;
        }
        ;
        XdmRequest.prototype.getActionType = function () {
            return this._actionType;
        };
        return XdmRequest;
    }(XdmMessage));
    var XdmResponse = (function (_super) {
        __extends(XdmResponse, _super);
        function XdmResponse(actionName, conversationId, correlationId, errorCode, responseType, data) {
            var _this = _super.call(this, 1, actionName, conversationId, correlationId, data) || this;
            _this._errorCode = errorCode;
            _this._responseType = responseType;
            return _this;
        }
        XdmResponse.prototype.getErrorCode = function () {
            return this._errorCode;
        };
        XdmResponse.prototype.getResponseType = function () {
            return this._responseType;
        };
        return XdmResponse;
    }(XdmMessage));
    var XdmMessagePackager = (function () {
        function XdmMessagePackager() {
        }
        XdmMessagePackager.envelope = function (messageObject, serializerVersion) {
            if (typeof (messageObject) === "object") {
                messageObject._serializerVersion = 1;
            }
            return JSON.stringify(messageObject);
        };
        XdmMessagePackager.unenvelope = function (messageObject, serializerVersion) {
            return JSON.parse(messageObject);
        };
        return XdmMessagePackager;
    }());
    var XdmResponseSender = (function () {
        function XdmResponseSender(requesterWindow, requesterUrl, actionName, conversationId, correlationId, responseType, serializerVersion) {
            var _this = this;
            this._invokeResultCode = 0;
            this._requesterWindow = requesterWindow;
            this._requesterUrl = requesterUrl;
            this._actionName = actionName;
            this._conversationId = conversationId;
            this._correlationId = correlationId;
            this._invokeResultCode = 0;
            this._responseType = responseType;
            this._serializerVersion = serializerVersion;
            this._send = function (result) {
                try {
                    var response = new XdmResponse(_this._actionName, _this._conversationId, _this._correlationId, _this._invokeResultCode, _this._responseType, result);
                    var envelopedResult = XdmMessagePackager.envelope(response, _this._serializerVersion);
                    _this._requesterWindow.postMessage(envelopedResult, _this._requesterUrl);
                    OSF.Utility.log("_send: requestUrl=" + _this._requesterUrl + " _actionName:" + _this._actionName);
                }
                catch (ex) {
                    OSF.Utility.log("ResponseSender._send error:" + ex.message);
                }
            };
        }
        XdmResponseSender.prototype.getRequesterWindow = function () {
            return this._requesterWindow;
        };
        XdmResponseSender.prototype.getRequesterUrl = function () {
            return this._requesterUrl;
        };
        XdmResponseSender.prototype.getActionName = function () {
            return this._actionName;
        };
        XdmResponseSender.prototype.getConversationId = function () {
            return this._conversationId;
        };
        XdmResponseSender.prototype.getCorrelationId = function () {
            return this._correlationId;
        };
        XdmResponseSender.prototype.getSend = function () {
            return this._send;
        };
        XdmResponseSender.prototype.setResultCode = function (resultCode) {
            this._invokeResultCode = resultCode;
        };
        return XdmResponseSender;
    }());
    var XdmInvokeCompleteCallback = (function (_super) {
        __extends(XdmInvokeCompleteCallback, _super);
        function XdmInvokeCompleteCallback(requesterWindow, requesterUrl, actionName, conversationId, correlationId, postCallbackHandler, serializerVersion) {
            var _this = _super.call(this, requesterWindow, requesterUrl, actionName, conversationId, correlationId, 0, serializerVersion) || this;
            _this._postCallbackHandler = postCallbackHandler;
            _this._send = function (result, responseCode) {
                if (responseCode != undefined) {
                    _this._invokeResultCode = responseCode;
                }
                try {
                    var response = new XdmResponse(_this._actionName, _this._conversationId, _this._correlationId, _this._invokeResultCode, _this._responseType, result);
                    var envelopedResult = XdmMessagePackager.envelope(response, _this._serializerVersion);
                    _this._requesterWindow.postMessage(envelopedResult, _this._requesterUrl);
                    _this._postCallbackHandler();
                }
                catch (ex) {
                    OSF.Utility.log("InvokeCompleteCallback._send error:" + ex.message);
                }
            };
            return _this;
        }
        return XdmInvokeCompleteCallback;
    }(XdmResponseSender));
    var XdmInvoker = (function () {
        function XdmInvoker(methodObject, paramValue, invokeCompleteCallback, eventHandlerProxyList, conversationId, eventName, serializerVersion) {
            this._callerId = '';
            this._methodObject = methodObject;
            this._param = paramValue;
            this._invokeCompleteCallback = invokeCompleteCallback;
            this._eventHandlerProxyList = eventHandlerProxyList;
            this._conversationId = conversationId;
            this._eventName = eventName;
            this._serializerVersion = serializerVersion;
        }
        XdmInvoker.prototype.invoke = function () {
            try {
                var result;
                switch (this._methodObject.getInvokeType()) {
                    case 0:
                        this._methodObject.getMethod()(this._param, this._invokeCompleteCallback.getSend());
                        break;
                    case 1:
                        result = this._methodObject.getMethod()(this._param);
                        this._invokeCompleteCallback.getSend()(result);
                        break;
                    case 4:
                        var eventHandlerProxy = this._createEventHandlerProxyObject(this._invokeCompleteCallback);
                        result = this._methodObject.getMethod()(eventHandlerProxy.getSend(), this._param);
                        this._eventHandlerProxyList[this._conversationId + this._eventName] = eventHandlerProxy.getSend();
                        this._invokeCompleteCallback.getSend()(result);
                        break;
                    case 5:
                        var eventHandler = this._eventHandlerProxyList[this._conversationId + this._eventName];
                        result = this._methodObject.getMethod()(eventHandler, this._param);
                        delete this._eventHandlerProxyList[this._conversationId + this._eventName];
                        this._invokeCompleteCallback.getSend()(result);
                        break;
                    case 2:
                        var eventHandlerProxyAsync = this._createEventHandlerProxyObject(this._invokeCompleteCallback);
                        this._methodObject.getMethod()(eventHandlerProxyAsync.getSend(), this._invokeCompleteCallback.getSend(), this._param);
                        this._eventHandlerProxyList[this._callerId + this._eventName] = eventHandlerProxyAsync.getSend();
                        break;
                    case 3:
                        var eventHandlerAsync = this._eventHandlerProxyList[this._callerId + this._eventName];
                        this._methodObject.getMethod()(eventHandlerAsync, this._invokeCompleteCallback.getSend(), this._param);
                        delete this._eventHandlerProxyList[this._callerId + this._eventName];
                        break;
                    default:
                        break;
                }
            }
            catch (ex) {
                this._invokeCompleteCallback.setResultCode(-3);
                this._invokeCompleteCallback.getSend()(ex);
            }
        };
        XdmInvoker.prototype.getInvokeBlockingFlag = function () {
            return this._methodObject.getBlockingFlag();
        };
        XdmInvoker.prototype._createEventHandlerProxyObject = function (invokeCompleteObject) {
            return new XdmResponseSender(invokeCompleteObject.getRequesterWindow(), invokeCompleteObject.getRequesterUrl(), invokeCompleteObject.getActionName(), invokeCompleteObject.getConversationId(), invokeCompleteObject.getCorrelationId(), 1, this._serializerVersion);
        };
        return XdmInvoker;
    }());
})(OSF || (OSF = {}));
var OSF;
(function (OSF) {
    var ParameterNames;
    (function (ParameterNames) {
        ParameterNames["Callback"] = "callback";
        ParameterNames["AsyncContext"] = "asyncContext";
        ParameterNames["Data"] = "data";
        ParameterNames["MessageToParent"] = "messageToParent";
        ParameterNames["MessageContent"] = "messageContent";
        ParameterNames["AppCommandInvocationCompletedData"] = "appCommandInvocationCompletedData";
    })(ParameterNames = OSF.ParameterNames || (OSF.ParameterNames = {}));
})(OSF || (OSF = {}));
var OSF;
(function (OSF) {
    var Requirement;
    (function (Requirement) {
        var RequirementVersion = (function () {
            function RequirementVersion() {
            }
            return RequirementVersion;
        }());
        Requirement.RequirementVersion = RequirementVersion;
        var RequirementMatrix = (function () {
            function RequirementMatrix(_setMap) {
                this._setMap = _setMap;
            }
            RequirementMatrix.prototype.isSetSupported = function (name, minVersion) {
                if (name == undefined) {
                    return false;
                }
                if (minVersion == undefined) {
                    minVersion = 0;
                }
                var setSupportArray = this._setMap;
                var sets = setSupportArray._sets;
                if (sets.hasOwnProperty(name.toLowerCase())) {
                    var setMaxVersion = sets[name.toLowerCase()];
                    try {
                        var setMaxVersionNum = this._getVersion(setMaxVersion + "");
                        minVersion = minVersion + "";
                        var minVersionNum = this._getVersion(minVersion);
                        if (setMaxVersionNum.major > 0 && setMaxVersionNum.major > minVersionNum.major) {
                            return true;
                        }
                        if (setMaxVersionNum.major > 0 &&
                            setMaxVersionNum.minor >= 0 &&
                            setMaxVersionNum.major == minVersionNum.major &&
                            setMaxVersionNum.minor >= minVersionNum.minor) {
                            return true;
                        }
                    }
                    catch (e) {
                        return false;
                    }
                }
                return false;
            };
            RequirementMatrix.prototype._getVersion = function (version) {
                version = version + "";
                var temp = version.split(".");
                var major = 0;
                var minor = 0;
                if (temp.length < 2 && isNaN(Number(version))) {
                    throw "version format incorrect";
                }
                else {
                    major = Number(temp[0]);
                    if (temp.length >= 2) {
                        minor = Number(temp[1]);
                    }
                    if (isNaN(major) || isNaN(minor)) {
                        throw "version format incorrect";
                    }
                }
                var result = { "minor": minor, "major": major };
                return result;
            };
            return RequirementMatrix;
        }());
        Requirement.RequirementMatrix = RequirementMatrix;
        var DefaultSetRequirement = (function () {
            function DefaultSetRequirement(setMap) {
                this._sets = setMap;
            }
            DefaultSetRequirement.prototype._addSetMap = function (addedSet) {
                for (var name in addedSet) {
                    this._sets[name] = addedSet[name];
                }
            };
            return DefaultSetRequirement;
        }());
        Requirement.DefaultSetRequirement = DefaultSetRequirement;
        var DefaultDialogSetRequirement = (function (_super) {
            __extends(DefaultDialogSetRequirement, _super);
            function DefaultDialogSetRequirement() {
                return _super.call(this, {
                    "dialogapi": 1.1
                }) || this;
            }
            return DefaultDialogSetRequirement;
        }(DefaultSetRequirement));
        Requirement.DefaultDialogSetRequirement = DefaultDialogSetRequirement;
        var RequirementsMatrixFactory = (function () {
            function RequirementsMatrixFactory() {
            }
            RequirementsMatrixFactory.getDefaultRequirementMatrix = function (appContext) {
                var defaultRequirementMatrix = undefined;
                var clientRequirement = appContext.get_requirementMatrix();
                if (clientRequirement != undefined && clientRequirement.length > 0) {
                    var matrixItem = JSON.parse(appContext.get_requirementMatrix().toLowerCase());
                    defaultRequirementMatrix = new RequirementMatrix(new DefaultSetRequirement(matrixItem));
                }
                else {
                    defaultRequirementMatrix = new RequirementMatrix(new DefaultSetRequirement({}));
                }
                return defaultRequirementMatrix;
            };
            RequirementsMatrixFactory.getDefaultDialogRequirementMatrix = function (appContext) {
                var defaultRequirementMatrix = undefined;
                var clientRequirement = appContext.get_dialogRequirementMatrix();
                if (clientRequirement != undefined && clientRequirement.length > 0) {
                    var matrixItem = JSON.parse(appContext.get_requirementMatrix().toLowerCase());
                    defaultRequirementMatrix = new RequirementMatrix(new DefaultSetRequirement(matrixItem));
                }
                else {
                    defaultRequirementMatrix = new RequirementMatrix(new DefaultDialogSetRequirement());
                }
                return defaultRequirementMatrix;
            };
            return RequirementsMatrixFactory;
        }());
        Requirement.RequirementsMatrixFactory = RequirementsMatrixFactory;
    })(Requirement = OSF.Requirement || (OSF.Requirement = {}));
})(OSF || (OSF = {}));
var OSF;
(function (OSF) {
    var DDA;
    (function (DDA) {
        var RichApi;
        (function (RichApi) {
            function executeRichApiRequestAsync(messageSafearray, callback) {
                var asyncMethodExecutor = OSF._OfficeAppFactory.getAsyncMethodExecutor();
                var dataTransform = {
                    toSafeArrayHost: function () {
                        return [messageSafearray];
                    },
                    fromSafeArrayHost: function (payload) {
                        return {
                            data: payload
                        };
                    },
                    toWebHost: function () {
                        return {
                            ArrayData: messageSafearray
                        };
                    },
                    fromWebHost: function (payload) {
                        return {
                            data: payload.Data
                        };
                    }
                };
                asyncMethodExecutor.executeAsync(93, dataTransform, callback);
            }
            RichApi.executeRichApiRequestAsync = executeRichApiRequestAsync;
            var _richApiMessageManager;
            Object.defineProperty(RichApi, 'richApiMessageManager', {
                get: function () {
                    if (!_richApiMessageManager) {
                        _richApiMessageManager = new OSF.RichApiMessageManager();
                    }
                    return _richApiMessageManager;
                }
            });
        })(RichApi = DDA.RichApi || (DDA.RichApi = {}));
    })(DDA = OSF.DDA || (OSF.DDA = {}));
})(OSF || (OSF = {}));
var OSF;
(function (OSF) {
    var RichApiMessageManager = (function () {
        function RichApiMessageManager() {
            this._eventDispatch = new OSF.EventDispatch([
                {
                    type: OSF.EventType.RichApiMessage,
                    id: 33,
                    getTargetId: function () { return ''; },
                    fromSafeArrayHost: function (payload) {
                        var entryArray = payload;
                        return RichApiMessageManager.transferEventArgument(entryArray);
                    },
                    fromWebHost: function (payload) {
                        var entryArray = payload.ArrayData;
                        return RichApiMessageManager.transferEventArgument(entryArray);
                    }
                }
            ]);
        }
        RichApiMessageManager.transferEventArgument = function (entryArray) {
            var entries = [];
            if (entryArray) {
                for (var i = 0; i < entryArray.length; i++) {
                    var elem = entryArray[i];
                    if (elem.toArray) {
                        elem = elem.toArray();
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
            return {
                type: OSF.EventType.RichApiMessage,
                entries: entries
            };
        };
        RichApiMessageManager.prototype.addHandlerAsync = function (eventType, handler, callback) {
            OSF.EventHelper.addEventHandler(eventType, handler, callback, this._eventDispatch);
        };
        RichApiMessageManager.prototype.removeHandlerAsync = function (eventType, handler, callback) {
            OSF.EventHelper.removeEventHandler(eventType, handler, callback, this._eventDispatch);
        };
        return RichApiMessageManager;
    }());
    OSF.RichApiMessageManager = RichApiMessageManager;
})(OSF || (OSF = {}));
var OSF;
(function (OSF) {
    var RichClientHostController = (function () {
        function RichClientHostController() {
        }
        RichClientHostController.prototype.execute = function (id, params, callback) {
            if (typeof OsfOMToken != 'undefined' && OsfOMToken) {
                window.external.Execute(id, params, callback, OsfOMToken);
            }
            else {
                window.external.Execute(id, params, callback);
            }
        };
        RichClientHostController.prototype.registerEvent = function (id, eventType, targetId, handler, callback) {
            if (typeof OsfOMToken != 'undefined' && OsfOMToken) {
                window.external.RegisterEvent(id, targetId, handler, callback, OsfOMToken);
            }
            else {
                window.external.RegisterEvent(id, targetId, handler, callback);
            }
        };
        RichClientHostController.prototype.unregisterEvent = function (id, eventType, targetId, callback) {
            if (typeof OsfOMToken != 'undefined' && OsfOMToken) {
                window.external.UnregisterEvent(id, targetId, callback, OsfOMToken);
            }
            else {
                window.external.UnregisterEvent(id, targetId, callback);
            }
        };
        return RichClientHostController;
    }());
    OSF.RichClientHostController = RichClientHostController;
})(OSF || (OSF = {}));
var OSF;
(function (OSF) {
    var RichClientInitializationHelper = (function (_super) {
        __extends(RichClientInitializationHelper, _super);
        function RichClientInitializationHelper() {
            return _super !== null && _super.apply(this, arguments) || this;
        }
        RichClientInitializationHelper.prototype.getOsfControlContext = function () {
            if (!this._osfControlContext) {
                var warningText = "Warning: Office.js is loaded outside of Office client";
                try {
                    if (window.external) {
                        this._osfControlContext = window.external.GetContext();
                    }
                    else {
                        console.error("There is no window.external.");
                        OSF.Utility.trace(warningText);
                        return null;
                    }
                }
                catch (e) {
                    console.error("Error when call window.external.GetContext() :" + JSON.stringify(e));
                    OSF.Utility.trace(warningText);
                    return null;
                }
            }
            return this._osfControlContext;
        };
        RichClientInitializationHelper.prototype.getAppContext = function (wnd, onSuccess, onError) {
            var _this = this;
            var context = this.getOsfControlContext();
            if (!context) {
                onError(new Error("The Office.js is loaded outside of Office client"));
                return;
            }
            var appType = context.GetAppType();
            var id = context.GetSolutionRef();
            var version = context.GetAppVersionMajor();
            var minorVersion = context.GetAppVersionMinor();
            var UILocale = context.GetAppUILocale();
            var dataLocale = context.GetAppDataLocale();
            var docUrl = context.GetDocUrl();
            var clientMode = context.GetAppCapabilities();
            var activationMode = context.GetActivationMode();
            var reason = (activationMode === 2) ? Office.InitializationReason.DocumentOpened : Office.InitializationReason.Inserted;
            var osfControlType = context.GetControlIntegrationLevel();
            var eToken;
            try {
                eToken = context.GetSolutionToken();
            }
            catch (ex) {
            }
            var correlationId;
            var externalNativeFunctionExists = OSF.Utility.externalNativeFunctionExists;
            if (externalNativeFunctionExists(typeof context.GetCorrelationId)) {
                correlationId = context.GetCorrelationId();
            }
            var appInstanceId;
            if (externalNativeFunctionExists(typeof context.GetInstanceId)) {
                appInstanceId = context.GetInstanceId();
            }
            var touchEnabled;
            if (externalNativeFunctionExists(typeof context.GetTouchEnabled)) {
                touchEnabled = context.GetTouchEnabled();
            }
            var commerceAllowed;
            if (externalNativeFunctionExists(typeof context.GetCommerceAllowed)) {
                commerceAllowed = context.GetCommerceAllowed();
            }
            var requirementMatrix;
            if (externalNativeFunctionExists(typeof context.GetSupportedMatrix)) {
                requirementMatrix = context.GetSupportedMatrix();
            }
            var hostCustomMessage;
            if (externalNativeFunctionExists(typeof context.GetHostCustomMessage)) {
                hostCustomMessage = context.GetHostCustomMessage();
            }
            var hostFullVersion;
            if (externalNativeFunctionExists(typeof context.GetHostFullVersion)) {
                hostFullVersion = context.GetHostFullVersion();
            }
            var dialogRequirementMatrix;
            if (externalNativeFunctionExists(typeof context.GetDialogRequirementMatrix)) {
                dialogRequirementMatrix = context.GetDialogRequirementMatrix();
            }
            var sdxFeatureGates;
            if (externalNativeFunctionExists(typeof context.GetFeaturesForSolution)) {
                try {
                    var sdxFeatureGatesJson = context.GetFeaturesForSolution();
                    if (sdxFeatureGatesJson) {
                        sdxFeatureGates = JSON.parse(sdxFeatureGatesJson);
                    }
                }
                catch (ex) {
                    OSF.Utility.trace("Exception while creating the SDX FeatureGates object. Details: " + ex);
                }
            }
            var initialDisplayMode = 0;
            if (externalNativeFunctionExists(typeof context.GetInitialDisplayMode)) {
                initialDisplayMode = context.GetInitialDisplayMode();
            }
            eToken = eToken ? eToken.toString() : "";
            var settingsFunc = function () { return _this.getSerializedSettings(); };
            var officeThemeFunc = function () { return _this.getOfficeTheme(); };
            var returnedContext = new OSF.OfficeAppContext(id, appType, version, UILocale, dataLocale, docUrl, clientMode, settingsFunc, reason, osfControlType, eToken, correlationId, appInstanceId, touchEnabled, commerceAllowed, minorVersion, requirementMatrix, hostCustomMessage, hostFullVersion, undefined, undefined, undefined, undefined, dialogRequirementMatrix, sdxFeatureGates, officeThemeFunc, initialDisplayMode);
            onSuccess(returnedContext);
            return;
        };
        RichClientInitializationHelper.prototype.createClientHostController = function () {
            if (!this._clientHostController) {
                if (this._hostInfo.hostPlatform === OSF.HostInfoPlatform.win32) {
                    this._clientHostController = new OSF.Win32RichClientHostController();
                }
                else {
                    throw OSF.Utility.createNotImplementedException();
                }
            }
            return this._clientHostController;
        };
        RichClientInitializationHelper.prototype.createAsyncMethodExecutor = function () {
            return new OSF.SafeArrayAsyncMethodExecutor(this._clientHostController);
        };
        RichClientInitializationHelper.prototype.createClientSettingsManager = function () {
            var manager = new OSF.RichClientSettingsManager(this.getOsfControlContext());
            return manager;
        };
        RichClientInitializationHelper.prototype.getSerializedSettings = function () {
            var osfControlContext = this.getOsfControlContext();
            var keys = [];
            var values = [];
            if (typeof OsfOMToken != 'undefined' && OsfOMToken) {
                osfControlContext.GetSettings(OsfOMToken).Read(keys, values);
            }
            else {
                osfControlContext.GetSettings().Read(keys, values);
            }
            var serializedSettings = {};
            for (var index = 0; index < keys.length; index++) {
                serializedSettings[keys[index]] = values[index];
            }
            return serializedSettings;
        };
        RichClientInitializationHelper.prototype.initializeSettings = function () {
            var osfControlContext = this.getOsfControlContext();
            var keys = [];
            var values = [];
            if (typeof OsfOMToken != 'undefined' && OsfOMToken) {
                osfControlContext.GetSettings(OsfOMToken).Read(keys, values);
            }
            else {
                osfControlContext.GetSettings().Read(keys, values);
            }
            var serializedSettings = {};
            for (var index = 0; index < keys.length; index++) {
                serializedSettings[keys[index]] = values[index];
            }
            return this.createSettings(serializedSettings);
        };
        RichClientInitializationHelper.prototype.getOfficeTheme = function () {
            var osfControlContext = this.getOsfControlContext();
            var osfOfficeThemeInfoString = osfControlContext.GetOfficeThemeInfo();
            var osfOfficeTheme = JSON.parse(osfOfficeThemeInfoString);
            for (var color in osfOfficeTheme) {
                osfOfficeTheme[color] = OSF.OUtil.convertIntToCssHexColor(osfOfficeTheme[color]);
            }
            return osfOfficeTheme;
        };
        return RichClientInitializationHelper;
    }(OSF.InitializationHelper));
    OSF.RichClientInitializationHelper = RichClientInitializationHelper;
})(OSF || (OSF = {}));
var OSF;
(function (OSF) {
    var RichClientSettingsManager = (function () {
        function RichClientSettingsManager(_osfClientContext) {
            this._osfClientContext = _osfClientContext;
        }
        RichClientSettingsManager.prototype.read = function (onComplete) {
            var keys = [];
            var values = [];
            var osfControlContext = this._osfClientContext;
            if (typeof OsfOMToken != 'undefined' && OsfOMToken) {
                osfControlContext.GetSettings(OsfOMToken).Read(keys, values);
            }
            else {
                osfControlContext.GetSettings().Read(keys, values);
            }
            var serializedSettings = {};
            for (var index = 0; index < keys.length; index++) {
                serializedSettings[keys[index]] = values[index];
            }
            if (onComplete) {
                onComplete(0, serializedSettings);
            }
        };
        RichClientSettingsManager.prototype.write = function (serializedSettings, onComplete) {
            var keys = [];
            var values = [];
            for (var key in serializedSettings) {
                keys.push(key);
                values.push(serializedSettings[key]);
            }
            var osfControlContext = this._osfClientContext;
            var settingObj;
            if (typeof OsfOMToken != 'undefined' && OsfOMToken) {
                settingObj = osfControlContext.GetSettings(OsfOMToken);
            }
            else {
                settingObj = osfControlContext.GetSettings();
            }
            if (typeof settingObj.WriteAsync != 'undefined') {
                settingObj.WriteAsync(keys, values, onComplete);
            }
            else {
                settingObj.Write(keys, values);
                onComplete(0);
            }
        };
        return RichClientSettingsManager;
    }());
    OSF.RichClientSettingsManager = RichClientSettingsManager;
})(OSF || (OSF = {}));
var OSF;
(function (OSF) {
    var SafeArrayAsyncMethodExecutor = (function (_super) {
        __extends(SafeArrayAsyncMethodExecutor, _super);
        function SafeArrayAsyncMethodExecutor(_clientHostController) {
            var _this = _super.call(this) || this;
            _this._clientHostController = _clientHostController;
            return _this;
        }
        SafeArrayAsyncMethodExecutor.prototype.executeAsync = function (id, dataTransform, callback) {
            var _this = this;
            try {
                var chunkResultData;
                this._clientHostController.execute(id, dataTransform.toSafeArrayHost(), function (hostResponseArgsNative, resultCode) {
                    var result;
                    var status;
                    var hostResponseArgs = OSF.Utility.fromSafeArray(hostResponseArgsNative);
                    if (typeof hostResponseArgs === "number") {
                        result = [];
                        status = hostResponseArgs;
                    }
                    else {
                        result = hostResponseArgs;
                        status = result[0];
                    }
                    if (status == 1) {
                        var payload = result[1];
                        if (payload != null) {
                            if (!chunkResultData) {
                                chunkResultData = new Array();
                            }
                            chunkResultData[payload[0]] = payload[1];
                        }
                        return false;
                    }
                    if (callback) {
                        var payload;
                        if (status == 0) {
                            if (result.length > 2) {
                                payload = [];
                                for (var i = 1; i < result.length; i++)
                                    payload[i - 1] = result[i];
                            }
                            else {
                                payload = result[1];
                            }
                            if (chunkResultData) {
                                if (payload != null) {
                                    var expectedChunkCount = payload[payload.length - 1];
                                    if (chunkResultData.length == expectedChunkCount) {
                                        payload[payload.length - 1] = chunkResultData;
                                    }
                                    else {
                                        status = 5001;
                                    }
                                }
                            }
                        }
                        else {
                            payload = result[1];
                        }
                        var value = null;
                        if (status == 0) {
                            value = dataTransform.fromSafeArrayHost(payload);
                        }
                        _this.invokeCallback(id, callback, status, value);
                    }
                    return true;
                });
            }
            catch (ex) {
                this.onException(ex, id, callback);
            }
        };
        SafeArrayAsyncMethodExecutor.prototype.registerEventAsync = function (id, eventType, targetId, handler, dataTransform, callback) {
            var _this = this;
            try {
                this._clientHostController.registerEvent(id, eventType, targetId, function (eventDispId, payload) {
                    var eventPayload = OSF.Utility.fromSafeArray(payload);
                    var eventArgs = dataTransform.fromSafeArrayHost(eventPayload);
                    handler(eventArgs);
                }, function (hostResponseArgsNative) {
                    var result;
                    var status;
                    var hostResponseArgs = OSF.Utility.fromSafeArray(hostResponseArgsNative);
                    if (typeof hostResponseArgs === "number") {
                        result = [];
                        status = hostResponseArgs;
                    }
                    else {
                        result = hostResponseArgs;
                        status = result[0];
                    }
                    _this.invokeCallback(id, callback, status, null);
                    return true;
                });
            }
            catch (ex) {
                this.onException(ex, id, callback);
            }
        };
        SafeArrayAsyncMethodExecutor.prototype.unregisterEventAsync = function (id, eventType, targetId, callback) {
            var _this = this;
            try {
                this._clientHostController.unregisterEvent(id, eventType, targetId, function (hostResponseArgsNative, resultCode) {
                    var result;
                    var status;
                    var hostResponseArgs = OSF.Utility.fromSafeArray(hostResponseArgsNative);
                    if (typeof hostResponseArgs === "number") {
                        result = [];
                        status = hostResponseArgs;
                    }
                    else {
                        result = hostResponseArgs;
                        status = result[0];
                    }
                    _this.invokeCallback(id, callback, status, null);
                    return true;
                });
            }
            catch (ex) {
                this.onException(ex, id, callback);
            }
        };
        SafeArrayAsyncMethodExecutor.prototype.onException = function (ex, dispId, callback) {
            var status;
            var statusNumber = ex.number;
            if (statusNumber) {
                switch (statusNumber) {
                    case -2146828218:
                        status = 7000;
                        break;
                    case -2147467259:
                        if (dispId == 10) {
                            status = 12007;
                        }
                        else {
                            status = 5001;
                        }
                        break;
                    case -2146828283:
                        status = 5010;
                        break;
                    case -2147209089:
                        status = 5010;
                        break;
                    case -2147208704:
                        status = 5100;
                        break;
                    case -2146827850:
                    default:
                        status = 5001;
                        break;
                }
            }
            if (callback) {
                this.invokeCallback(dispId, callback, status || 5001, null);
            }
        };
        return SafeArrayAsyncMethodExecutor;
    }(OSF.AsyncMethodExecutor));
    OSF.SafeArrayAsyncMethodExecutor = SafeArrayAsyncMethodExecutor;
})(OSF || (OSF = {}));
var OSF;
(function (OSF) {
    var SafeStorage = (function () {
        function SafeStorage(_internalStorage) {
            this._internalStorage = _internalStorage;
        }
        SafeStorage.prototype.getItem = function (key) {
            try {
                return this._internalStorage && this._internalStorage.getItem(key);
            }
            catch (e) {
                return null;
            }
        };
        SafeStorage.prototype.setItem = function (key, data) {
            try {
                this._internalStorage && this._internalStorage.setItem(key, data);
            }
            catch (e) {
            }
        };
        SafeStorage.prototype.clear = function () {
            try {
                this._internalStorage && this._internalStorage.clear();
            }
            catch (e) {
            }
        };
        SafeStorage.prototype.removeItem = function (key) {
            try {
                this._internalStorage && this._internalStorage.removeItem(key);
            }
            catch (e) {
            }
        };
        SafeStorage.prototype.getKeysWithPrefix = function (keyPrefix) {
            var keyList = [];
            try {
                var len = this._internalStorage && this._internalStorage.length || 0;
                for (var i = 0; i < len; i++) {
                    var key = this._internalStorage.key(i);
                    if (key.indexOf(keyPrefix) === 0) {
                        keyList.push(key);
                    }
                }
            }
            catch (e) {
            }
            return keyList;
        };
        return SafeStorage;
    }());
    OSF.SafeStorage = SafeStorage;
})(OSF || (OSF = {}));
var Office;
(function (Office) {
    var Settings = (function () {
        function Settings(settings, _clientSettingsManager) {
            var _this = this;
            this._clientSettingsManager = _clientSettingsManager;
            settings = settings || {};
            this._settings = settings;
            this._eventDispatch = new OSF.EventDispatch([
                {
                    id: 1,
                    type: OSF.EventType.SettingsChanged,
                    getTargetId: function () { return ''; },
                    fromSafeArrayHost: function (payload) {
                        return {
                            type: OSF.EventType.SettingsChanged,
                            settings: _this
                        };
                    },
                    fromWebHost: function (payload) {
                        return {
                            type: OSF.EventType.SettingsChanged,
                            settings: _this
                        };
                    }
                }
            ]);
        }
        Settings.prototype.cacheSessionSettings = function (settings) {
            var osfSessionStorage = OSF.OUtil.getSessionStorage();
            if (osfSessionStorage) {
                var serializedSettings = OSF.OUtil.serializeSettings(settings);
                var storageSettings = JSON.stringify(serializedSettings);
                osfSessionStorage.setItem(OSF._OfficeAppFactory.getCachedSessionSettingsKey(), storageSettings);
            }
        };
        Settings.prototype.get = function (name) {
            var setting = this._settings[name];
            return typeof (setting) === 'undefined' ? null : setting;
        };
        Settings.prototype.set = function (name, value) {
            this._settings[name] = value;
            this.cacheSessionSettings(this._settings);
        };
        Settings.prototype.remove = function (name) {
            delete this._settings[name];
            this.cacheSessionSettings(this._settings);
        };
        Settings.prototype.saveAsync = function (callback) {
            var settingsManager = this._clientSettingsManager;
            var serializedSettings = OSF.OUtil.serializeSettings(this._settings);
            settingsManager.write(serializedSettings, function (errorCode) {
                var result = OSF.Utility.asyncResultFromErrorCode(errorCode);
                if (callback) {
                    callback(result);
                }
            });
        };
        Settings.prototype.refreshAsync = function (callback) {
            var _this = this;
            var settingsManager = this._clientSettingsManager;
            settingsManager.read(function (errorCode, serializedSettings) {
                var result = OSF.Utility.asyncResultFromErrorCode(errorCode);
                if (result.status === Office.AsyncResultStatus.succeeded) {
                    _this._settings = OSF.OUtil.deserializeSettings(serializedSettings);
                    result.value = _this;
                }
                if (callback) {
                    callback(result);
                }
            });
        };
        Settings.prototype.addHandlerAsync = function (eventType, handler, callback) {
            OSF.EventHelper.addEventHandler(eventType, handler, callback, this._eventDispatch);
        };
        Settings.prototype.removeHandlerAsync = function (eventType, handler, callback) {
            OSF.EventHelper.removeEventHandler(eventType, handler, callback, this._eventDispatch);
        };
        Settings.prototype.toJSON = function () {
            return this._settings;
        };
        return Settings;
    }());
    Office.Settings = Settings;
})(Office || (Office = {}));
var OSF;
(function (OSF) {
    var Utility;
    (function (Utility) {
        function createArgumentException(name) {
            return new Error("Invalid argument " + name);
        }
        Utility.createArgumentException = createArgumentException;
        function createNotImplementedException() {
            return new Error("Not implemented yet");
        }
        Utility.createNotImplementedException = createNotImplementedException;
        function log(message) {
            console.log(message);
        }
        Utility.log = log;
        function trace(message) {
            console.log(message);
        }
        Utility.trace = trace;
        function debugLog(message) {
            console.log(message);
        }
        Utility.debugLog = debugLog;
        function createPromiseFromResult(result) {
            return Promise.resolve(result);
        }
        Utility.createPromiseFromResult = createPromiseFromResult;
        function createPromise(executor) {
            var ret = new Promise(executor);
            return ret;
        }
        Utility.createPromise = createPromise;
        function compareVersions(version1, version2) {
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
        }
        Utility.compareVersions = compareVersions;
        function getQueryStringValue(paramName) {
            if (window.location.search) {
                var regex = new RegExp('[?&]' + paramName + '=([^&]*)');
                var match = regex.exec(window.location.search);
                if (match) {
                    var ret = match[1];
                    return ret;
                }
            }
            return null;
        }
        Utility.getQueryStringValue = getQueryStringValue;
        function getErrorCodeFromAsyncResult(asyncResult) {
            if (asyncResult.status === Office.AsyncResultStatus.succeeded) {
                return 0;
            }
            if (asyncResult.error && asyncResult.error.code) {
                return asyncResult.error.code;
            }
            return 5001;
        }
        Utility.getErrorCodeFromAsyncResult = getErrorCodeFromAsyncResult;
        function isNullOrUndefined(value) {
            if (typeof (value) === "undefined") {
                return true;
            }
            if (value === null) {
                return true;
            }
            return false;
        }
        Utility.isNullOrUndefined = isNullOrUndefined;
        function isNullOrEmpty(value) {
            if (isNullOrUndefined(value)) {
                return true;
            }
            return (value.length === 0);
        }
        Utility.isNullOrEmpty = isNullOrEmpty;
        function externalNativeFunctionExists(type) {
            return type === 'unknown' || type !== 'undefined';
        }
        Utility.externalNativeFunctionExists = externalNativeFunctionExists;
        function stringEndsWith(value, subString) {
            if (isNullOrUndefined(value)) {
                throw createArgumentException("value");
            }
            if (isNullOrUndefined(subString)) {
                throw createArgumentException("subString");
            }
            if (subString.length > value.length) {
                return false;
            }
            if (value.substr(value.length - subString.length) === subString) {
                return true;
            }
            return false;
        }
        Utility.stringEndsWith = stringEndsWith;
        function fromSafeArray(value) {
            var ret = value;
            if (value != null && value.toArray) {
                var arrayResult = value.toArray();
                ret = new Array(arrayResult.length);
                for (var i = 0; i < arrayResult.length; i++) {
                    ret[i] = fromSafeArray(arrayResult[i]);
                }
            }
            return ret;
        }
        Utility.fromSafeArray = fromSafeArray;
        function asyncResultFromErrorCode(errorCode) {
            if (errorCode === 0) {
                return {
                    status: Office.AsyncResultStatus.succeeded
                };
            }
            else {
                return {
                    status: Office.AsyncResultStatus.failed,
                    error: {
                        code: errorCode
                    }
                };
            }
        }
        Utility.asyncResultFromErrorCode = asyncResultFromErrorCode;
        Utility._DebugXdm = false;
    })(Utility = OSF.Utility || (OSF.Utility = {}));
})(OSF || (OSF = {}));
var OSF;
(function (OSF) {
    var WebAsyncMethodExecutor = (function (_super) {
        __extends(WebAsyncMethodExecutor, _super);
        function WebAsyncMethodExecutor(_clientHostController) {
            var _this = _super.call(this) || this;
            _this._clientHostController = _clientHostController;
            return _this;
        }
        WebAsyncMethodExecutor.prototype.executeAsync = function (id, dataTransform, callback) {
            var _this = this;
            this._clientHostController.execute(id, dataTransform.toWebHost(), function (resultCode, payload) {
                if (callback) {
                    var value = null;
                    if (resultCode == 0) {
                        value = dataTransform.fromWebHost(payload);
                    }
                    _this.invokeCallback(id, callback, resultCode, value);
                }
                return true;
            });
        };
        WebAsyncMethodExecutor.prototype.registerEventAsync = function (id, eventType, targetId, handler, dataTransform, callback) {
            var _this = this;
            this._clientHostController.registerEvent(id, eventType, targetId, function (payload) {
                var eventPayload = payload;
                var eventArgs = dataTransform.fromWebHost(eventPayload);
                handler(eventArgs);
            }, function (resultCode, payload) {
                if (callback) {
                    _this.invokeCallback(id, callback, resultCode, null);
                }
                return true;
            });
        };
        WebAsyncMethodExecutor.prototype.unregisterEventAsync = function (id, eventType, targetId, callback) {
            var _this = this;
            this._clientHostController.unregisterEvent(id, eventType, targetId, function (resultCode, payload) {
                if (callback) {
                    _this.invokeCallback(id, callback, resultCode, null);
                }
                return true;
            });
        };
        return WebAsyncMethodExecutor;
    }(OSF.AsyncMethodExecutor));
    OSF.WebAsyncMethodExecutor = WebAsyncMethodExecutor;
})(OSF || (OSF = {}));
var OSF;
(function (OSF) {
    var WebClientHostController = (function () {
        function WebClientHostController(webAppState) {
            this._delegateVersion = 1;
            this._webAppState = webAppState;
        }
        WebClientHostController.prototype.execute = function (id, params, callback) {
            var _this = this;
            var hostCallArgs = params;
            if (!hostCallArgs) {
                hostCallArgs = {};
            }
            hostCallArgs.DdaMethod = {
                ControlId: this._webAppState.id,
                DispatchId: id,
                Version: this._delegateVersion
            };
            hostCallArgs.__timeout__ = -1;
            this._webAppState.clientEndPoint.invoke('executeMethod', function (xdmStatus, payload) {
                var error = 0;
                if (xdmStatus == 0) {
                    _this._delegateVersion = payload["Version"];
                    error = payload["Error"];
                }
                else {
                    switch (xdmStatus) {
                        case -5:
                            error = 7000;
                            break;
                        default:
                            error = 5001;
                            break;
                    }
                }
                if (callback) {
                    callback(error, payload);
                }
            }, hostCallArgs);
        };
        WebClientHostController.prototype.registerEvent = function (id, eventType, targetId, handler, callback) {
            this._webAppState.clientEndPoint.registerForEvent(this.getXdmEventName(targetId, eventType), function (payload) {
                if (handler) {
                    handler(payload);
                }
            }, this._getOnAfterRegisterEvent(true, id, callback), {
                controlId: this._webAppState.id,
                eventDispId: id,
                targetId: targetId,
                __timeout__: -1
            });
        };
        WebClientHostController.prototype.unregisterEvent = function (id, eventType, targetId, callback) {
            this._webAppState.clientEndPoint.unregisterForEvent(this.getXdmEventName(targetId, eventType), this._getOnAfterRegisterEvent(false, id, callback), {
                controlId: this._webAppState.id,
                eventDispId: id,
                targetId: targetId,
                __timeout__: -1
            });
        };
        WebClientHostController.prototype.messageParent = function (params) {
            throw OSF.Utility.createNotImplementedException();
        };
        WebClientHostController.prototype.openDialog = function (id, eventType, targetId, handler, callback) {
            throw OSF.Utility.createNotImplementedException();
        };
        WebClientHostController.prototype.closeDialog = function (id, eventType, targetId, callback) {
            throw OSF.Utility.createNotImplementedException();
        };
        WebClientHostController.prototype.sendMessage = function (params) {
            throw OSF.Utility.createNotImplementedException();
        };
        WebClientHostController.prototype.getXdmEventName = function (targetId, eventType) {
            if (eventType == OSF.EventType.BindingSelectionChanged ||
                eventType == OSF.EventType.BindingDataChanged ||
                eventType == OSF.EventType.DataNodeDeleted ||
                eventType == OSF.EventType.DataNodeInserted ||
                eventType == OSF.EventType.DataNodeReplaced) {
                return targetId + "_" + eventType;
            }
            else {
                return eventType;
            }
        };
        WebClientHostController.prototype._getOnAfterRegisterEvent = function (register, id, callback) {
            var startTime = (new Date()).getTime();
            return function (xdmStatus, payload) {
                var status;
                if (xdmStatus != 0) {
                    switch (xdmStatus) {
                        case -5:
                            status = 7000;
                            break;
                        default:
                            status = 5001;
                            break;
                    }
                }
                else {
                    if (payload) {
                        if (payload["Error"]) {
                            status = payload["Error"];
                        }
                        else {
                            status = 0;
                        }
                    }
                    else {
                        status = 5001;
                    }
                }
                if (callback) {
                    callback(status);
                }
                if (OSF.AppTelemetry) {
                    OSF.AppTelemetry.onRegisterDone(register, id, Math.abs((new Date()).getTime() - startTime), status);
                }
            };
        };
        return WebClientHostController;
    }());
    OSF.WebClientHostController = WebClientHostController;
})(OSF || (OSF = {}));
var OSF;
(function (OSF) {
    var WebClientSettingsManager = (function () {
        function WebClientSettingsManager() {
        }
        WebClientSettingsManager.prototype.read = function (onComplete) {
            var asyncMethodExecutor = OSF._OfficeAppFactory.getAsyncMethodExecutor();
            var dataTransform = {
                toSafeArrayHost: function () {
                    return [];
                },
                fromSafeArrayHost: function (payload) {
                    return null;
                },
                toWebHost: function () {
                    return {};
                },
                fromWebHost: function (payload) {
                    return payload.Properties.Settings;
                }
            };
            var callback = function (result) {
                if (result.status === Office.AsyncResultStatus.succeeded) {
                    var serializedSettings = {};
                    for (var i = 0; i < result.value.length; i++) {
                        var entry = result.value[i];
                        if (Array.isArray(entry)) {
                            serializedSettings[entry[0]] = entry[1];
                        }
                        else {
                            serializedSettings[entry.Name] = entry.Value;
                        }
                    }
                    onComplete(0, serializedSettings);
                }
                else {
                    var errorCode = result.error.code;
                    onComplete(errorCode, {});
                }
            };
            asyncMethodExecutor.executeAsync(75, dataTransform, callback);
        };
        WebClientSettingsManager.prototype.write = function (serializedSettings, onComplete) {
            var asyncMethodExecutor = OSF._OfficeAppFactory.getAsyncMethodExecutor();
            var properties = [];
            for (var key in serializedSettings) {
                var entry = [];
                entry.push(key);
                entry.push(serializedSettings[key]);
                properties.push(entry);
            }
            var dataTransform = {
                toSafeArrayHost: function () {
                    return null;
                },
                fromSafeArrayHost: function (payload) {
                    return null;
                },
                toWebHost: function () {
                    return {
                        DdaSettingsMethod: {
                            OverwriteIfStale: true,
                            Properties: properties
                        }
                    };
                },
                fromWebHost: function (payload) {
                    return null;
                }
            };
            var callback = function (result) {
                if (result.status === Office.AsyncResultStatus.succeeded) {
                    onComplete(0);
                }
                else {
                    var errorCode = result.error.code;
                    onComplete(errorCode);
                }
            };
            asyncMethodExecutor.executeAsync(76, dataTransform, callback);
        };
        return WebClientSettingsManager;
    }());
    OSF.WebClientSettingsManager = WebClientSettingsManager;
})(OSF || (OSF = {}));
var OSF;
(function (OSF) {
    var WebInitializationHelper = (function (_super) {
        __extends(WebInitializationHelper, _super);
        function WebInitializationHelper(hostInfo, webAppState, context, hostFacade) {
            var _this = _super.call(this, hostInfo, webAppState, context, hostFacade) || this;
            _this._appContext = {};
            _this._tabbableElements = "a[href]:not([tabindex='-1'])," +
                "area[href]:not([tabindex='-1'])," +
                "button:not([disabled]):not([tabindex='-1'])," +
                "input:not([disabled]):not([tabindex='-1'])," +
                "select:not([disabled]):not([tabindex='-1'])," +
                "textarea:not([disabled]):not([tabindex='-1'])," +
                "*[tabindex]:not([tabindex='-1'])," +
                "*[contenteditable]:not([disabled]):not([tabindex='-1'])";
            return _this;
        }
        WebInitializationHelper.prototype.saveAndSetDialogInfo = function (hostInfoValue) {
            function getAppIdFromWindowLocation() {
                var xdmInfoValue = OSF.OUtil.parseXdmInfo(true);
                if (xdmInfoValue) {
                    var items = xdmInfoValue.split("|");
                    return items[1];
                }
                return null;
            }
            ;
            var osfSessionStorage = OSF.OUtil.getSessionStorage();
            if (osfSessionStorage) {
                if (!hostInfoValue) {
                    hostInfoValue = OSF.OUtil.parseHostInfoFromWindowName(true, OSF._OfficeAppFactory.getWindowName());
                }
                if (hostInfoValue && hostInfoValue.indexOf("isDialog") > -1) {
                    var appId = getAppIdFromWindowLocation();
                    if (appId != null) {
                        osfSessionStorage.setItem(appId + "IsDialog", "true");
                    }
                    this._hostInfo.isDialog = true;
                    return;
                }
                this._hostInfo.isDialog = osfSessionStorage.getItem(OSF.OUtil.getXdmFieldValue("AppId", false) + "IsDialog") != null ? true : false;
            }
        };
        WebInitializationHelper.prototype.setAgaveHostCommunication = function () {
            try {
                var me = this;
                var xdmInfoValue = OSF.OUtil.parseXdmInfoWithGivenFragment(false, OSF._OfficeAppFactory.getWindowLocationHash());
                if (!xdmInfoValue) {
                    xdmInfoValue = OSF.OUtil.parseXdmInfoFromWindowName(false, OSF._OfficeAppFactory.getWindowName());
                }
                if (xdmInfoValue) {
                    var xdmItems = OSF.OUtil.getInfoItems(xdmInfoValue);
                    if (xdmItems != undefined && xdmItems.length >= 3) {
                        me._webAppState.conversationID = xdmItems[0];
                        me._webAppState.id = xdmItems[1];
                        me._webAppState.webAppUrl = xdmItems[2].indexOf(":") >= 0 ? xdmItems[2] : decodeURIComponent(xdmItems[2]);
                    }
                }
                me._webAppState.wnd = window.opener != null ? window.opener : window.parent;
                var serializerVersion = OSF.OUtil.parseSerializerVersionWithGivenFragment(false, OSF._OfficeAppFactory.getWindowLocationHash());
                if (isNaN(serializerVersion)) {
                    serializerVersion = OSF.OUtil.parseSerializerVersionFromWindowName(false, OSF._OfficeAppFactory.getWindowName());
                }
                me._webAppState.serializerVersion = serializerVersion;
                if (this._hostInfo.isDialog && window.opener != null) {
                    return;
                }
                me._webAppState.clientEndPoint = OSF.XdmCommunicationManager.connect(me._webAppState.conversationID, me._webAppState.wnd, me._webAppState.webAppUrl, me._webAppState.serializerVersion);
                me._webAppState.serviceEndPoint = OSF.XdmCommunicationManager.createServiceEndPoint(me._webAppState.id);
                var notificationConversationId = me._webAppState.conversationID + OSF.Constants.NotificationConversationIdSuffix;
                me._webAppState.serviceEndPoint.registerConversation(notificationConversationId, me._webAppState.webAppUrl);
                var notifyAgave = function (params) {
                    var actionId;
                    if (typeof params == "string") {
                        actionId = params;
                    }
                    else {
                        actionId = params[0];
                    }
                    switch (actionId) {
                        case OSF.AgaveHostAction.Select:
                            me._webAppState.focused = true;
                            break;
                        case OSF.AgaveHostAction.UnSelect:
                            me._webAppState.focused = false;
                            break;
                        case OSF.AgaveHostAction.TabIn:
                        case OSF.AgaveHostAction.CtrlF6In:
                            window.focus();
                            var list = document.querySelectorAll(me._tabbableElements);
                            var focused = OSF.OUtil.focusToFirstTabbable(list, false);
                            if (!focused) {
                                window.blur();
                                me._webAppState.focused = false;
                                me._webAppState.clientEndPoint.invoke("ContextActivationManager_notifyHost", null, [me._webAppState.id, OSF.AgaveHostAction.ExitNoFocusable]);
                            }
                            break;
                        case OSF.AgaveHostAction.TabInShift:
                            window.focus();
                            var list = document.querySelectorAll(me._tabbableElements);
                            var focused = OSF.OUtil.focusToFirstTabbable(list, true);
                            if (!focused) {
                                window.blur();
                                me._webAppState.focused = false;
                                me._webAppState.clientEndPoint.invoke("ContextActivationManager_notifyHost", null, [me._webAppState.id, OSF.AgaveHostAction.ExitNoFocusableShift]);
                            }
                            break;
                        case OSF.AgaveHostAction.SendMessage:
                            if (Office.context.messaging.onMessage) {
                                var message = params[1];
                                Office.context.messaging.onMessage(message);
                            }
                            break;
                        case OSF.AgaveHostAction.TaskPaneHeaderButtonClicked:
                            if (Office.context.ui.taskPaneAction.onHeaderButtonClick) {
                                Office.context.ui.taskPaneAction.onHeaderButtonClick();
                            }
                            break;
                        default:
                            OSF.Utility.trace("actionId " + actionId + " notifyAgave is wrong.");
                            break;
                    }
                };
                me._webAppState.serviceEndPoint.registerMethod("Office_notifyAgave", notifyAgave, 0, false);
                me.addOrRemoveEventListenersForWindow(true);
            }
            catch (ex) {
                if (OSF.AppTelemetry) {
                    OSF.AppTelemetry.logAppException("Exception thrown in setAgaveHostCommunication. Exception:[" + ex + "]");
                }
                throw ex;
            }
        };
        WebInitializationHelper.prototype.getAppContext = function (wnd, onSuccess, onError) {
            var _this = this;
            var me = this;
            var getInvocationCallbackWebApp = function (errorCode, appContext) {
                OSFPerformance.getAppContextXdmEnd = OSFPerformance.now();
                if (appContext._appName === 16) {
                    var serializedSettingsFromHost = appContext._settings;
                    _this._serializedSettings = {};
                    for (var index in serializedSettingsFromHost) {
                        var setting = serializedSettingsFromHost[index];
                        _this._serializedSettings[setting[0]] = setting[1];
                    }
                }
                else {
                    _this._serializedSettings = appContext._settings;
                }
                if (!me._hostInfo.isDialog || window.opener == null) {
                    var pageUrl = window.location.href;
                    me._webAppState.clientEndPoint.invoke("ContextActivationManager_notifyHost", null, [me._webAppState.id, OSF.AgaveHostAction.UpdateTargetUrl, pageUrl]);
                }
                if (errorCode === 0 && appContext._id != undefined && appContext._appName != undefined && appContext._appVersion != undefined && appContext._appUILocale != undefined && appContext._dataLocale != undefined &&
                    appContext._docUrl != undefined && appContext._clientMode != undefined && appContext._settings != undefined && appContext._reason != undefined) {
                    me._appContext = appContext;
                    var appInstanceId = (appContext._appInstanceId ? appContext._appInstanceId : appContext._id);
                    var touchEnabled = false;
                    var commerceAllowed = true;
                    var minorVersion = 0;
                    if (appContext._appMinorVersion != undefined) {
                        minorVersion = appContext._appMinorVersion;
                    }
                    var requirementMatrix = undefined;
                    if (appContext._requirementMatrix != undefined) {
                        requirementMatrix = appContext._requirementMatrix;
                    }
                    appContext.eToken = appContext.eToken ? appContext.eToken : "";
                    var settingsFunc = function () {
                        return _this._serializedSettings;
                    };
                    var returnedContext = new OSF.OfficeAppContext(appContext._id, appContext._appName, appContext._appVersion, appContext._appUILocale, appContext._dataLocale, appContext._docUrl, appContext._clientMode, settingsFunc, appContext._reason, appContext._osfControlType, appContext._eToken, appContext._correlationId, appInstanceId, touchEnabled, commerceAllowed, minorVersion, requirementMatrix, appContext._hostCustomMessage, appContext._hostFullVersion, appContext._clientWindowHeight, appContext._clientWindowWidth, appContext._addinName, appContext._appDomains, appContext._dialogRequirementMatrix, appContext._featureGates, undefined, appContext._initialDisplayMode);
                    onSuccess(returnedContext);
                }
                else {
                    var errorMsg = "Function ContextActivationManager_getAppContextAsync call failed. ErrorCode is " + errorCode + ", exception: " + appContext;
                    if (OSF.AppTelemetry) {
                        OSF.AppTelemetry.logAppException(errorMsg);
                    }
                    onError(errorMsg);
                }
            };
            try {
                var skipSessionStorage = true;
                if (this._hostInfo.isDialog && window.opener != null) {
                    skipSessionStorage = false;
                }
                var appContext = OSF.OUtil.parseAppContextFromWindowName(skipSessionStorage, OSF._OfficeAppFactory.getWindowName());
                if (appContext) {
                    getInvocationCallbackWebApp(0, appContext);
                }
                else {
                    OSFPerformance.getAppContextXdmStart = OSFPerformance.now();
                    this._webAppState.clientEndPoint.invoke("ContextActivationManager_getAppContextAsync", getInvocationCallbackWebApp, this._webAppState.id);
                }
            }
            catch (ex) {
                if (OSF.AppTelemetry) {
                    OSF.AppTelemetry.logAppException("Exception thrown when trying to invoke getAppContextAsync. Exception:[" + ex + "]");
                }
                onError(ex);
            }
        };
        WebInitializationHelper.prototype.createClientHostController = function () {
            if (!this._clientHostController) {
                this._clientHostController = new OSF.WebClientHostController(this._webAppState);
            }
            return this._clientHostController;
        };
        WebInitializationHelper.prototype.createAsyncMethodExecutor = function () {
            return new OSF.WebAsyncMethodExecutor(this._clientHostController);
        };
        WebInitializationHelper.prototype.createClientSettingsManager = function () {
            return new OSF.WebClientSettingsManager();
        };
        WebInitializationHelper.prototype.addOrRemoveEventListenersForWindow = function (isAdd) {
            var me = this;
            var onWindowFocus = function () {
                if (!me._webAppState.focused) {
                    me._webAppState.focused = true;
                }
                me._webAppState.clientEndPoint.invoke("ContextActivationManager_notifyHost", null, [me._webAppState.id, OSF.AgaveHostAction.Select]);
            };
            var onWindowBlur = function () {
                if (!OSF) {
                    return;
                }
                if (me._webAppState.focused) {
                    me._webAppState.focused = false;
                }
                me._webAppState.clientEndPoint.invoke("ContextActivationManager_notifyHost", null, [me._webAppState.id, OSF.AgaveHostAction.UnSelect]);
            };
            var onWindowKeydown = function (e) {
                e.preventDefault = e.preventDefault || function () {
                    e.returnValue = false;
                };
                if (e.keyCode == 117 && (e.ctrlKey || e.metaKey)) {
                    e.preventDefault();
                    var actionId = OSF.AgaveHostAction.CtrlF6Exit;
                    if (e.shiftKey) {
                        actionId = OSF.AgaveHostAction.CtrlF6ExitShift;
                    }
                    me._webAppState.clientEndPoint.invoke("ContextActivationManager_notifyHost", null, [me._webAppState.id, actionId]);
                }
                else if (e.keyCode == 9) {
                    e.preventDefault();
                    var allTabbableElements = document.querySelectorAll(me._tabbableElements);
                    var focused = OSF.OUtil.focusToNextTabbable(allTabbableElements, e.target || e.srcElement, e.shiftKey);
                    if (!focused) {
                        if (me._hostInfo.isDialog) {
                            OSF.OUtil.focusToFirstTabbable(allTabbableElements, e.shiftKey);
                        }
                        else {
                            if (e.shiftKey) {
                                me._webAppState.clientEndPoint.invoke("ContextActivationManager_notifyHost", null, [me._webAppState.id, OSF.AgaveHostAction.TabExitShift]);
                            }
                            else {
                                me._webAppState.clientEndPoint.invoke("ContextActivationManager_notifyHost", null, [me._webAppState.id, OSF.AgaveHostAction.TabExit]);
                            }
                        }
                    }
                }
                else if (e.keyCode == 27) {
                    e.preventDefault();
                    me._webAppState.clientEndPoint.invoke("ContextActivationManager_notifyHost", null, [me._webAppState.id, OSF.AgaveHostAction.EscExit]);
                }
                else if (e.keyCode == 113) {
                    e.preventDefault();
                    me._webAppState.clientEndPoint.invoke("ContextActivationManager_notifyHost", null, [me._webAppState.id, OSF.AgaveHostAction.F2Exit]);
                }
            };
            var onWindowKeypress = function (e) {
                if (e.keyCode == 117 && e.ctrlKey) {
                    if (e.preventDefault) {
                        e.preventDefault();
                    }
                    else {
                        e.returnValue = false;
                    }
                }
            };
            if (!OSF.Utility._DebugXdm) {
                if (isAdd) {
                    OSF.OUtil.addEventListener(window, "focus", onWindowFocus);
                    OSF.OUtil.addEventListener(window, "blur", onWindowBlur);
                    OSF.OUtil.addEventListener(window, "keydown", onWindowKeydown);
                    OSF.OUtil.addEventListener(window, "keypress", onWindowKeypress);
                }
                else {
                    OSF.OUtil.removeEventListener(window, "focus", onWindowFocus);
                    OSF.OUtil.removeEventListener(window, "blur", onWindowBlur);
                    OSF.OUtil.removeEventListener(window, "keydown", onWindowKeydown);
                    OSF.OUtil.removeEventListener(window, "keypress", onWindowKeypress);
                }
            }
        };
        return WebInitializationHelper;
    }(OSF.InitializationHelper));
    OSF.WebInitializationHelper = WebInitializationHelper;
})(OSF || (OSF = {}));
var OSF;
(function (OSF) {
    var WebkitClientSettingsManager = (function () {
        function WebkitClientSettingsManager(_initializationHelper, _scriptMessager) {
            this._initializationHelper = _initializationHelper;
            this._scriptMessager = _scriptMessager;
        }
        WebkitClientSettingsManager.prototype.read = function (onComplete) {
            var keys = [];
            var values = [];
            var initializationHelper = this._initializationHelper;
            var onGetAppContextSuccess = function (appContext) {
                if (onComplete) {
                    var serializedSettings = appContext.get_settingsFunc()();
                    onComplete(0, serializedSettings);
                }
            };
            var onGetAppContextError = function (e) {
                if (onComplete) {
                    onComplete(5001, {});
                }
            };
            initializationHelper.getAppContext(null, onGetAppContextSuccess, onGetAppContextError);
        };
        WebkitClientSettingsManager.prototype.write = function (serializedSettings, onComplete) {
            var hostParams = {};
            var keys = [];
            var values = [];
            for (var key in serializedSettings) {
                keys.push(key);
                values.push(serializedSettings[key]);
            }
            hostParams["keys"] = keys;
            hostParams["values"] = values;
            var onWriteCompleted = function onWriteCompleted(status) {
                if (onComplete) {
                    onComplete(status[0]);
                }
            };
            this._scriptMessager.invokeMethod(OSF.Webkit.MessageHandlerName, OSF.Webkit.MethodId.WriteSettings, hostParams, onWriteCompleted);
        };
        return WebkitClientSettingsManager;
    }());
    OSF.WebkitClientSettingsManager = WebkitClientSettingsManager;
})(OSF || (OSF = {}));
var OSF;
(function (OSF) {
    var Webkit;
    (function (Webkit) {
        Webkit.MessageHandlerName = "Agave";
        Webkit.PopupMessageHandlerName = "WefPopupHandler";
        var AppContextProperties;
        (function (AppContextProperties) {
            AppContextProperties[AppContextProperties["Settings"] = 0] = "Settings";
            AppContextProperties[AppContextProperties["SolutionReferenceId"] = 1] = "SolutionReferenceId";
            AppContextProperties[AppContextProperties["AppType"] = 2] = "AppType";
            AppContextProperties[AppContextProperties["MajorVersion"] = 3] = "MajorVersion";
            AppContextProperties[AppContextProperties["MinorVersion"] = 4] = "MinorVersion";
            AppContextProperties[AppContextProperties["RevisionVersion"] = 5] = "RevisionVersion";
            AppContextProperties[AppContextProperties["APIVersionSequence"] = 6] = "APIVersionSequence";
            AppContextProperties[AppContextProperties["AppCapabilities"] = 7] = "AppCapabilities";
            AppContextProperties[AppContextProperties["APPUILocale"] = 8] = "APPUILocale";
            AppContextProperties[AppContextProperties["AppDataLocale"] = 9] = "AppDataLocale";
            AppContextProperties[AppContextProperties["BindingCount"] = 10] = "BindingCount";
            AppContextProperties[AppContextProperties["DocumentUrl"] = 11] = "DocumentUrl";
            AppContextProperties[AppContextProperties["ActivationMode"] = 12] = "ActivationMode";
            AppContextProperties[AppContextProperties["ControlIntegrationLevel"] = 13] = "ControlIntegrationLevel";
            AppContextProperties[AppContextProperties["SolutionToken"] = 14] = "SolutionToken";
            AppContextProperties[AppContextProperties["APISetVersion"] = 15] = "APISetVersion";
            AppContextProperties[AppContextProperties["CorrelationId"] = 16] = "CorrelationId";
            AppContextProperties[AppContextProperties["InstanceId"] = 17] = "InstanceId";
            AppContextProperties[AppContextProperties["TouchEnabled"] = 18] = "TouchEnabled";
            AppContextProperties[AppContextProperties["CommerceAllowed"] = 19] = "CommerceAllowed";
            AppContextProperties[AppContextProperties["RequirementMatrix"] = 20] = "RequirementMatrix";
            AppContextProperties[AppContextProperties["HostCustomMessage"] = 21] = "HostCustomMessage";
            AppContextProperties[AppContextProperties["HostFullVersion"] = 22] = "HostFullVersion";
            AppContextProperties[AppContextProperties["InitialDisplayMode"] = 23] = "InitialDisplayMode";
        })(AppContextProperties = Webkit.AppContextProperties || (Webkit.AppContextProperties = {}));
        var MethodId;
        (function (MethodId) {
            MethodId[MethodId["Execute"] = 1] = "Execute";
            MethodId[MethodId["RegisterEvent"] = 2] = "RegisterEvent";
            MethodId[MethodId["UnregisterEvent"] = 3] = "UnregisterEvent";
            MethodId[MethodId["WriteSettings"] = 4] = "WriteSettings";
            MethodId[MethodId["GetContext"] = 5] = "GetContext";
            MethodId[MethodId["SendMessage"] = 6] = "SendMessage";
            MethodId[MethodId["MessageParent"] = 7] = "MessageParent";
        })(MethodId = Webkit.MethodId || (Webkit.MethodId = {}));
        var WebkitHostController = (function () {
            function WebkitHostController(hostScriptProxy) {
                this.hostScriptProxy = hostScriptProxy;
                this.useFullDialogAPI = !!window._enableFullDialogAPI;
            }
            WebkitHostController.prototype.execute = function (id, params, callback) {
                var hostParams = {
                    id: id,
                    apiArgs: params
                };
                var agaveResponseCallback = function (payload) {
                    if (callback) {
                        var invokeArguments = [];
                        if (OSF.OUtil.isArray(payload)) {
                            for (var i = 0; i < payload.length; i++) {
                                var element = payload[i];
                                if (OSF.OUtil.isArray(element)) {
                                    element = new OSF.WebkitSafeArray(element);
                                }
                                invokeArguments.unshift(element);
                            }
                        }
                        return callback.apply(null, invokeArguments);
                    }
                };
                this.hostScriptProxy.invokeMethod(OSF.Webkit.MessageHandlerName, OSF.Webkit.MethodId.Execute, hostParams, agaveResponseCallback);
            };
            WebkitHostController.prototype.registerEvent = function (id, eventType, targetId, handler, callback) {
                var agaveEventHandlerCallback = function (payload) {
                    var safeArraySource = payload;
                    var eventId = 0;
                    if (OSF.OUtil.isArray(payload) && payload.length >= 2) {
                        safeArraySource = payload[0];
                        eventId = payload[1];
                    }
                    if (handler) {
                        handler(eventId, new OSF.WebkitSafeArray(safeArraySource));
                    }
                };
                var agaveResponseCallback = function (payload) {
                    if (callback) {
                        return callback(new OSF.WebkitSafeArray(payload));
                    }
                };
                this.hostScriptProxy.registerEvent(OSF.Webkit.MessageHandlerName, OSF.Webkit.MethodId.RegisterEvent, id, targetId, agaveEventHandlerCallback, agaveResponseCallback);
            };
            WebkitHostController.prototype.unregisterEvent = function (id, eventType, targetId, callback) {
                var agaveResponseCallback = function (response) {
                    return callback(new OSF.WebkitSafeArray(response));
                };
                this.hostScriptProxy.unregisterEvent(OSF.Webkit.MessageHandlerName, OSF.Webkit.MethodId.UnregisterEvent, id, targetId, agaveResponseCallback);
            };
            WebkitHostController.prototype.messageParent = function (params) {
                var message = params[OSF.ParameterNames.MessageToParent];
                if (this.useFullDialogAPI) {
                    this.hostScriptProxy.invokeMethod(OSF.Webkit.MessageHandlerName, OSF.Webkit.MethodId.MessageParent, message, null);
                }
                else {
                    var messageObj = { dialogMessage: { messageType: 0, messageContent: message } };
                    window.opener.postMessage(JSON.stringify(messageObj), window.location.origin);
                }
            };
            WebkitHostController.prototype.openDialog = function (id, eventType, targetId, handler, callback) {
                if (this.useFullDialogAPI) {
                    this.registerEvent(id, eventType, targetId, handler, callback);
                    return;
                }
                if (WebkitHostController.popup && !WebkitHostController.popup.closed) {
                    callback(12007);
                    return;
                }
                var magicWord = "action=displayDialog";
                WebkitHostController.OpenDialogCallback = undefined;
                var fragmentSeparator = '#';
                var callArgs = JSON.parse(targetId);
                var callUrl = callArgs.url;
                if (!callUrl) {
                    callback(12003);
                    return;
                }
                var urlParts = callUrl.split(fragmentSeparator);
                var seperator = "?";
                if (urlParts[0].indexOf("?") > -1) {
                    seperator = "&";
                }
                var width = screen.width * callArgs.width / 100;
                var height = screen.height * callArgs.height / 100;
                var params = "width=" + width + ", height=" + height;
                urlParts[0] = urlParts[0].concat(seperator).concat(magicWord);
                var openUrl = urlParts.join(fragmentSeparator);
                WebkitHostController.popup = window.open(openUrl, "", params);
                function receiveMessage(event) {
                    if (event.origin == window.location.origin) {
                        try {
                            var messageObj = JSON.parse(event.data);
                            if (messageObj.dialogMessage) {
                                handler(id, [0, messageObj.dialogMessage.messageContent]);
                            }
                        }
                        catch (e) {
                            OSF.Utility.trace("messages received cannot be handlered. Message:" + event.data);
                        }
                    }
                }
                WebkitHostController.DialogEventListener = receiveMessage;
                function checkWindowClose() {
                    try {
                        if (WebkitHostController.popup == null || WebkitHostController.popup.closed) {
                            window.clearInterval(WebkitHostController.interval);
                            window.removeEventListener("message", WebkitHostController.DialogEventListener);
                            WebkitHostController.NotifyError = null;
                            WebkitHostController.popup = null;
                            handler(id, [12006]);
                        }
                    }
                    catch (e) {
                        OSF.Utility.trace("Error happened when popup window closed.");
                    }
                }
                WebkitHostController.OpenDialogCallback = function (code) {
                    if (code == 0) {
                        window.addEventListener("message", WebkitHostController.DialogEventListener);
                        WebkitHostController.interval = window.setInterval(checkWindowClose, 1000);
                        WebkitHostController.NotifyError = function (errorCode) {
                            handler(id, [errorCode]);
                        };
                    }
                    callback(code);
                };
            };
            WebkitHostController.prototype.closeDialog = function (id, eventType, targetId, callback) {
                if (this.useFullDialogAPI) {
                    this.unregisterEvent(id, eventType, targetId, callback);
                }
                else {
                    if (WebkitHostController.popup) {
                        if (WebkitHostController.interval) {
                            window.clearInterval(WebkitHostController.interval);
                        }
                        WebkitHostController.popup.close();
                        WebkitHostController.popup = null;
                        window.removeEventListener("message", WebkitHostController.DialogEventListener);
                        WebkitHostController.NotifyError = null;
                        callback(0);
                    }
                    else {
                        callback(5001);
                    }
                }
            };
            WebkitHostController.prototype.sendMessage = function (params) {
                var message = params[OSF.ParameterNames.MessageContent];
                if (!isNaN(parseFloat(message)) && isFinite(message)) {
                    message = message.toString();
                }
                this.hostScriptProxy.invokeMethod(OSF.Webkit.MessageHandlerName, OSF.Webkit.MethodId.SendMessage, message, null);
            };
            return WebkitHostController;
        }());
        Webkit.WebkitHostController = WebkitHostController;
    })(Webkit = OSF.Webkit || (OSF.Webkit = {}));
})(OSF || (OSF = {}));
var OSF;
(function (OSF) {
    var WebkitInitializationHelper = (function (_super) {
        __extends(WebkitInitializationHelper, _super);
        function WebkitInitializationHelper(hostInfo, webAppState, context, hostFacade) {
            var _this = _super.call(this, hostInfo, webAppState, context, hostFacade) || this;
            _this.initializeWebkitMessaging();
            return _this;
        }
        WebkitInitializationHelper.prototype.initializeWebkitMessaging = function () {
            OSF.ScriptMessaging = OSFWebkit.ScriptMessaging;
        };
        WebkitInitializationHelper.prototype.getAppContext = function (wnd, onSuccess, onError) {
            var _this = this;
            var getInvocationCallback = function (appContext) {
                var returnedContext;
                var appContextProperties = OSF.Webkit.AppContextProperties;
                var appType = appContext[appContextProperties.AppType];
                var hostSettings = appContext[appContextProperties.Settings];
                _this._serializedSettings = {};
                var keys = hostSettings[0];
                var values = hostSettings[1];
                for (var index = 0; index < keys.length; index++) {
                    _this._serializedSettings[keys[index]] = values[index];
                }
                var id = appContext[appContextProperties.SolutionReferenceId];
                var version = appContext[appContextProperties.MajorVersion];
                var minorVersion = appContext[appContextProperties.MinorVersion];
                var clientMode = appContext[appContextProperties.AppCapabilities];
                var UILocale = appContext[appContextProperties.APPUILocale];
                var dataLocale = appContext[appContextProperties.AppDataLocale];
                var docUrl = appContext[appContextProperties.DocumentUrl];
                var reason = appContext[appContextProperties.ActivationMode];
                var osfControlType = appContext[appContextProperties.ControlIntegrationLevel];
                var eToken = appContext[appContextProperties.SolutionToken];
                eToken = eToken ? eToken.toString() : "";
                var correlationId = appContext[appContextProperties.CorrelationId];
                var appInstanceId = appContext[appContextProperties.InstanceId];
                var touchEnabled = appContext[appContextProperties.TouchEnabled];
                var commerceAllowed = appContext[appContextProperties.CommerceAllowed];
                var requirementMatrix = appContext[appContextProperties.RequirementMatrix];
                var hostCustomMessage = appContext[appContextProperties.HostCustomMessage];
                var hostFullVersion = appContext[appContextProperties.HostFullVersion];
                var initialDisplayMode = appContext[appContextProperties.InitialDisplayMode];
                var settingsFunc = function () {
                    return _this._serializedSettings;
                };
                returnedContext = new OSF.OfficeAppContext(id, appType, version, UILocale, dataLocale, docUrl, clientMode, settingsFunc, reason, osfControlType, eToken, correlationId, appInstanceId, touchEnabled, commerceAllowed, minorVersion, requirementMatrix, hostCustomMessage, hostFullVersion, undefined, undefined, undefined, undefined, undefined, undefined, undefined, initialDisplayMode);
                onSuccess(returnedContext);
            };
            var handler;
            if (this._hostInfo.isDialog && window.webkit.messageHandlers[OSF.Webkit.PopupMessageHandlerName]) {
                handler = OSF.Webkit.PopupMessageHandlerName;
            }
            else {
                handler = OSF.Webkit.MessageHandlerName;
            }
            OSF.ScriptMessaging.GetScriptMessenger().invokeMethod(handler, OSF.Webkit.MethodId.GetContext, [], getInvocationCallback);
        };
        WebkitInitializationHelper.prototype.createClientHostController = function () {
            if (!this._clientHostController) {
                this._clientHostController = new OSF.Webkit.WebkitHostController(OSF.ScriptMessaging.GetScriptMessenger());
            }
            return this._clientHostController;
        };
        WebkitInitializationHelper.prototype.createAsyncMethodExecutor = function () {
            return new OSF.SafeArrayAsyncMethodExecutor(this.createClientHostController());
        };
        WebkitInitializationHelper.prototype.createClientSettingsManager = function () {
            return new OSF.WebkitClientSettingsManager(this, OSF.ScriptMessaging.GetScriptMessenger());
        };
        return WebkitInitializationHelper;
    }(OSF.InitializationHelper));
    OSF.WebkitInitializationHelper = WebkitInitializationHelper;
})(OSF || (OSF = {}));
var OSF;
(function (OSF) {
    var WebkitSafeArray = (function () {
        function WebkitSafeArray(data) {
            this.data = data;
            this.safeArrayFlag = this.isSafeArray(data);
        }
        WebkitSafeArray.prototype.dimensions = function () {
            var dimensions = 0;
            if (this.safeArrayFlag) {
                dimensions = this.data[0][0];
            }
            else if (this.isArray()) {
                dimensions = 2;
            }
            return dimensions;
        };
        WebkitSafeArray.prototype.getItem = function () {
            var array = [];
            var element = null;
            if (this.safeArrayFlag) {
                array = this.toArray();
            }
            else {
                array = this.data;
            }
            element = array;
            for (var i = 0; i < arguments.length; i++) {
                element = element[arguments[i]];
            }
            return element;
        };
        WebkitSafeArray.prototype.lbound = function (dimension) {
            return 0;
        };
        WebkitSafeArray.prototype.ubound = function (dimension) {
            var ubound = 0;
            if (this.safeArrayFlag) {
                ubound = this.data[0][dimension];
            }
            else if (this.isArray()) {
                if (dimension == 1) {
                    return this.data.length;
                }
                else if (dimension == 2) {
                    if (OSF.OUtil.isArray(this.data[0])) {
                        return this.data[0].length;
                    }
                    else if (this.data[0] != null) {
                        return 1;
                    }
                }
            }
            return ubound;
        };
        WebkitSafeArray.prototype.toArray = function () {
            if (this.isArray() == false) {
                return this.data;
            }
            var arr = [];
            var startingIndex = this.safeArrayFlag ? 1 : 0;
            for (var i = startingIndex; i < this.data.length; i++) {
                var element = this.data[i];
                if (this.isSafeArray(element)) {
                    arr.push(new WebkitSafeArray(element));
                }
                else {
                    arr.push(element);
                }
            }
            return arr;
        };
        WebkitSafeArray.prototype.isArray = function () {
            return OSF.OUtil.isArray(this.data);
        };
        WebkitSafeArray.prototype.isSafeArray = function (obj) {
            var isSafeArray = false;
            if (OSF.OUtil.isArray(obj) && OSF.OUtil.isArray(obj[0])) {
                var bounds = obj[0];
                var dimensions = bounds[0];
                if (bounds.length != dimensions + 1) {
                    return false;
                }
                var expectedArraySize = 1;
                for (var i = 1; i < bounds.length; i++) {
                    var dimension = bounds[i];
                    if (isFinite(dimension) == false) {
                        return false;
                    }
                    expectedArraySize = expectedArraySize * dimension;
                }
                expectedArraySize++;
                isSafeArray = (expectedArraySize == obj.length);
            }
            return isSafeArray;
        };
        return WebkitSafeArray;
    }());
    OSF.WebkitSafeArray = WebkitSafeArray;
})(OSF || (OSF = {}));
var OSFWebkit;
(function (OSFWebkit) {
    var ScriptMessaging;
    (function (ScriptMessaging) {
        var scriptMessenger = null;
        function agaveHostCallback(callbackId, params) {
            scriptMessenger.agaveHostCallback(callbackId, params);
        }
        ScriptMessaging.agaveHostCallback = agaveHostCallback;
        function agaveHostEventCallback(callbackId, params) {
            scriptMessenger.agaveHostEventCallback(callbackId, params);
        }
        ScriptMessaging.agaveHostEventCallback = agaveHostEventCallback;
        function GetScriptMessenger() {
            if (scriptMessenger == null) {
                scriptMessenger = new WebkitScriptMessaging("OSF.ScriptMessaging.agaveHostCallback", "OSF.ScriptMessaging.agaveHostEventCallback");
            }
            return scriptMessenger;
        }
        ScriptMessaging.GetScriptMessenger = GetScriptMessenger;
        var EventHandlerCallback = (function () {
            function EventHandlerCallback(id, targetId, handler) {
                this.id = id;
                this.targetId = targetId;
                this.handler = handler;
            }
            return EventHandlerCallback;
        }());
        var WebkitScriptMessaging = (function () {
            function WebkitScriptMessaging(methodCallbackName, eventCallbackName) {
                this.callingIndex = 0;
                this.callbackList = {};
                this.eventHandlerList = {};
                this.asyncMethodCallbackFunctionName = methodCallbackName;
                this.eventCallbackFunctionName = eventCallbackName;
                this.conversationId = WebkitScriptMessaging.getCurrentTimeMS().toString();
            }
            WebkitScriptMessaging.prototype.invokeMethod = function (handlerName, methodId, params, callback) {
                var messagingArgs = {};
                this.postWebkitMessage(messagingArgs, handlerName, methodId, params, callback);
            };
            WebkitScriptMessaging.prototype.registerEvent = function (handlerName, methodId, dispId, targetId, handler, callback) {
                var messagingArgs = {
                    eventCallbackFunction: this.eventCallbackFunctionName
                };
                var hostArgs = {
                    id: dispId,
                    targetId: targetId
                };
                var correlationId = this.postWebkitMessage(messagingArgs, handlerName, methodId, hostArgs, callback);
                this.eventHandlerList[correlationId] = new EventHandlerCallback(dispId, targetId, handler);
            };
            WebkitScriptMessaging.prototype.unregisterEvent = function (handlerName, methodId, dispId, targetId, callback) {
                var hostArgs = {
                    id: dispId,
                    targetId: targetId
                };
                for (var key in this.eventHandlerList) {
                    if (this.eventHandlerList.hasOwnProperty(key)) {
                        var eventCallback = this.eventHandlerList[key];
                        if (eventCallback.id == dispId && eventCallback.targetId == targetId) {
                            delete this.eventHandlerList[key];
                        }
                    }
                }
                this.invokeMethod(handlerName, methodId, hostArgs, callback);
            };
            WebkitScriptMessaging.prototype.agaveHostCallback = function (callbackId, params) {
                var callbackFunction = this.callbackList[callbackId];
                if (callbackFunction) {
                    var callbacksDone = callbackFunction(params);
                    if (callbacksDone === undefined || callbacksDone === true) {
                        delete this.callbackList[callbackId];
                    }
                }
            };
            WebkitScriptMessaging.prototype.agaveHostEventCallback = function (callbackId, params) {
                var eventCallback = this.eventHandlerList[callbackId];
                if (eventCallback) {
                    eventCallback.handler(params);
                }
            };
            WebkitScriptMessaging.prototype.postWebkitMessage = function (messagingArgs, handlerName, methodId, params, callback) {
                messagingArgs.methodId = methodId;
                messagingArgs.params = params;
                var correlationId = "";
                if (callback) {
                    correlationId = this.generateCorrelationId();
                    this.callbackList[correlationId] = callback;
                    messagingArgs.callbackId = correlationId;
                    messagingArgs.callbackFunction = this.asyncMethodCallbackFunctionName;
                }
                var invokePostMessage = function () {
                    window.webkit.messageHandlers[handlerName].postMessage(JSON.stringify(messagingArgs));
                };
                var currentTimestamp = WebkitScriptMessaging.getCurrentTimeMS();
                if (this.lastMessageTimestamp == null || (currentTimestamp - this.lastMessageTimestamp >= WebkitScriptMessaging.MESSAGE_TIME_DELTA)) {
                    invokePostMessage();
                    this.lastMessageTimestamp = currentTimestamp;
                }
                else {
                    this.lastMessageTimestamp += WebkitScriptMessaging.MESSAGE_TIME_DELTA;
                    setTimeout(function () {
                        invokePostMessage();
                    }, this.lastMessageTimestamp - currentTimestamp);
                }
                return correlationId;
            };
            WebkitScriptMessaging.prototype.generateCorrelationId = function () {
                ++this.callingIndex;
                return this.conversationId + this.callingIndex;
            };
            WebkitScriptMessaging.getCurrentTimeMS = function () {
                return (new Date).getTime();
            };
            WebkitScriptMessaging.MESSAGE_TIME_DELTA = 10;
            return WebkitScriptMessaging;
        }());
    })(ScriptMessaging = OSFWebkit.ScriptMessaging || (OSFWebkit.ScriptMessaging = {}));
})(OSFWebkit || (OSFWebkit = {}));
var OSF;
(function (OSF) {
    var Win32RichClientHostController = (function (_super) {
        __extends(Win32RichClientHostController, _super);
        function Win32RichClientHostController() {
            return _super !== null && _super.apply(this, arguments) || this;
        }
        Win32RichClientHostController.prototype.messageParent = function (params) {
            var message = params[OSF.ParameterNames.MessageToParent];
            window.external.MessageParent(message);
        };
        Win32RichClientHostController.prototype.openDialog = function (id, eventType, targetId, handler, callback) {
            this.registerEvent(id, eventType, targetId, handler, callback);
        };
        Win32RichClientHostController.prototype.closeDialog = function (id, eventType, targetId, callback) {
            this.unregisterEvent(id, eventType, targetId, callback);
        };
        Win32RichClientHostController.prototype.sendMessage = function (params) {
            var message = params[OSF.ParameterNames.MessageContent];
            window.external.MessageChild(message);
        };
        return Win32RichClientHostController;
    }(OSF.RichClientHostController));
    OSF.Win32RichClientHostController = Win32RichClientHostController;
})(OSF || (OSF = {}));
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
            if (typeof (window) !== "undefined") {
                if (window.Promise) {
                    return window.Promise;
                }
                else {
                    var ret = _Internal.PromiseImpl.Init();
                    window.Promise = ret;
                    return ret;
                }
            }
        }
        _Internal.OfficePromise = determinePromise();
    })(_Internal = Office._Internal || (Office._Internal = {}));
    Office.OfficePromise = _Internal.OfficePromise;
    Office.Promise = Office.OfficePromise;
})(Office || (Office = {}));
var OSF;
(function (OSF) {
    var AppTelemetry;
    (function (AppTelemetry) {
        var appInfo;
        var sessionId = OSF.OUtil.Guid.generateNewGuid();
        var osfControlAppCorrelationId = "";
        var omexDomainRegex = new RegExp("^https?://store\\.office(ppe|-int)?\\.com/", "i");
        AppTelemetry.enableTelemetry = true;
        var AppInfo = (function () {
            function AppInfo() {
            }
            return AppInfo;
        }());
        AppTelemetry.AppInfo = AppInfo;
        var AppStorage = (function () {
            function AppStorage() {
                this.clientIDKey = "Office API client";
                this.logIdSetKey = "Office App Log Id Set";
            }
            AppStorage.prototype.getClientId = function () {
                var clientId = this.getValue(this.clientIDKey);
                if (!clientId || clientId.length <= 0 || clientId.length > 40) {
                    clientId = OSF.OUtil.Guid.generateNewGuid();
                    this.setValue(this.clientIDKey, clientId);
                }
                return clientId;
            };
            AppStorage.prototype.saveLog = function (logId, log) {
                var logIdSet = this.getValue(this.logIdSetKey);
                logIdSet = ((logIdSet && logIdSet.length > 0) ? (logIdSet + ";") : "") + logId;
                this.setValue(this.logIdSetKey, logIdSet);
                this.setValue(logId, log);
            };
            AppStorage.prototype.enumerateLog = function (callback, clean) {
                var logIdSet = this.getValue(this.logIdSetKey);
                if (logIdSet) {
                    var ids = logIdSet.split(";");
                    for (var id in ids) {
                        var logId = ids[id];
                        var log = this.getValue(logId);
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
            AppStorage.prototype.getValue = function (key) {
                var osfLocalStorage = OSF.OUtil.getLocalStorage();
                var value = "";
                if (osfLocalStorage) {
                    value = osfLocalStorage.getItem(key);
                }
                return value;
            };
            AppStorage.prototype.setValue = function (key, value) {
                var osfLocalStorage = OSF.OUtil.getLocalStorage();
                if (osfLocalStorage) {
                    osfLocalStorage.setItem(key, value);
                }
            };
            AppStorage.prototype.remove = function (key) {
                var osfLocalStorage = OSF.OUtil.getLocalStorage();
                if (osfLocalStorage) {
                    try {
                        osfLocalStorage.removeItem(key);
                    }
                    catch (ex) {
                    }
                }
            };
            return AppStorage;
        }());
        function trimStringToLowerCase(input) {
            if (input) {
                input = input.replace(/[{}]/g, "").toLowerCase();
            }
            return (input || "");
        }
        function initialize(context) {
            if (!AppTelemetry.enableTelemetry) {
                return;
            }
            if (appInfo) {
                return;
            }
            appInfo = new AppInfo();
            if (context.get_hostFullVersion()) {
                appInfo.hostVersion = context.get_hostFullVersion();
            }
            else {
                appInfo.hostVersion = context.get_appVersion();
            }
            appInfo.appId = context.get_id();
            appInfo.host = "" + context.get_appName();
            appInfo.browser = window.navigator.userAgent;
            appInfo.correlationId = trimStringToLowerCase(context.get_correlationId());
            appInfo.clientId = (new AppStorage()).getClientId();
            appInfo.appInstanceId = context.get_appInstanceId();
            if (appInfo.appInstanceId) {
                appInfo.appInstanceId = appInfo.appInstanceId.replace(/[{}]/g, "").toLowerCase();
            }
            appInfo.message = context.get_hostCustomMessage();
            appInfo.officeJSVersion = "16.0";
            appInfo.hostJSVersion = "16.0";
            if (context._wacHostEnvironment) {
                appInfo.wacHostEnvironment = context._wacHostEnvironment;
            }
            if (context._isFromWacAutomation !== undefined && context._isFromWacAutomation !== null) {
                appInfo.isFromWacAutomation = context._isFromWacAutomation.toString().toLowerCase();
            }
            var docUrl = context.get_docUrl();
            appInfo.docUrl = omexDomainRegex.test(docUrl) ? docUrl : "";
            var url = location.href;
            if (url) {
                url = url.split("?")[0].split("#")[0];
            }
            appInfo.appURL = AppTelemetry.UrlFilter.filter(url);
            (function getUserIdAndAssetIdFromToken(token, appInfo) {
                appInfo.assetId = "";
                appInfo.userId = "";
                try {
                    if (!OSF.Utility.isNullOrEmpty(token)) {
                        var xmlContent = decodeURIComponent(token);
                        var parser = new DOMParser();
                        var xmlDoc = parser.parseFromString(xmlContent, "text/xml");
                        var cidNode = xmlDoc.getElementsByTagName("t")[0].attributes.getNamedItem("cid");
                        var oidNode = xmlDoc.getElementsByTagName("t")[0].attributes.getNamedItem("oid");
                        if (cidNode && cidNode.nodeValue) {
                            appInfo.userId = cidNode.nodeValue;
                        }
                        else if (oidNode && oidNode.nodeValue) {
                            appInfo.userId = oidNode.nodeValue;
                        }
                        appInfo.assetId = xmlDoc.getElementsByTagName("t")[0].attributes.getNamedItem("aid").nodeValue;
                    }
                }
                catch (e) {
                }
            })(context.get_eToken(), appInfo);
            appInfo.sessionId = sessionId;
            appInfo.name = context.get_addinName();
            setTimeout(function () {
                OTel.OTelLogger.initialize(appInfo);
            }, 10 * 1000);
            AppTelemetry.onAppActivated();
        }
        AppTelemetry.initialize = initialize;
        function onAppActivated() {
            if (!appInfo) {
                return;
            }
        }
        AppTelemetry.onAppActivated = onAppActivated;
        function onScriptDone(scriptId, msStartTime, msResponseTime, appCorrelationId) {
        }
        AppTelemetry.onScriptDone = onScriptDone;
        function onCallDone(apiType, id, parameters, msResponseTime, errorType) {
            if (!appInfo) {
                return;
            }
        }
        AppTelemetry.onCallDone = onCallDone;
        ;
        function onMethodDone(id, args, msResponseTime, errorType) {
        }
        AppTelemetry.onMethodDone = onMethodDone;
        function onPropertyDone(propertyName, msResponseTime) {
            OSF.AppTelemetry.onCallDone("property", -1, propertyName, msResponseTime, 0);
        }
        AppTelemetry.onPropertyDone = onPropertyDone;
        function onCheckWACHost(isWacKnownHost, solutionId, hostType, hostPlatform, correlationId, wacDomain) {
        }
        AppTelemetry.onCheckWACHost = onCheckWACHost;
        function onEventDone(id, errorType) {
            OSF.AppTelemetry.onCallDone("event", id, null, 0, errorType);
        }
        AppTelemetry.onEventDone = onEventDone;
        function onRegisterDone(register, id, msResponseTime, errorType) {
            OSF.AppTelemetry.onCallDone(register ? "registerevent" : "unregisterevent", id, null, msResponseTime, errorType);
        }
        AppTelemetry.onRegisterDone = onRegisterDone;
        function onAppClosed(openTime, focusTime) {
            if (!appInfo) {
                return;
            }
        }
        AppTelemetry.onAppClosed = onAppClosed;
        function setOsfControlAppCorrelationId(correlationId) {
            osfControlAppCorrelationId = trimStringToLowerCase(correlationId);
        }
        AppTelemetry.setOsfControlAppCorrelationId = setOsfControlAppCorrelationId;
        function doAppInitializationLogging(isException, message) {
        }
        AppTelemetry.doAppInitializationLogging = doAppInitializationLogging;
        function logAppCommonMessage(message) {
            doAppInitializationLogging(false, message);
        }
        AppTelemetry.logAppCommonMessage = logAppCommonMessage;
        function logAppException(errorMessage) {
            doAppInitializationLogging(true, errorMessage);
        }
        AppTelemetry.logAppException = logAppException;
    })(AppTelemetry = OSF.AppTelemetry || (OSF.AppTelemetry = {}));
})(OSF || (OSF = {}));
var OTel;
(function (OTel) {
    var CDN_PATH_OTELJS_AGAVE = 'telemetry/oteljs_agave.js';
    var CDN_PATH_OTELJS = 'telemetry/oteljs.js';
    var OTelLogger = (function () {
        function OTelLogger() {
        }
        OTelLogger.loaded = function () {
            return !(OTelLogger.logger === undefined);
        };
        OTelLogger.getOtelCDNLocation = function () {
            return (OSF.LoadScriptHelper.getHostBundleJsBasePath() + CDN_PATH_OTELJS);
        };
        OTelLogger.getOtelSinkCDNLocation = function () {
            return (OSF.LoadScriptHelper.getHostBundleJsBasePath() + CDN_PATH_OTELJS_AGAVE);
        };
        OTelLogger.getMapName = function (map, name) {
            if (name !== undefined && map.hasOwnProperty(name)) {
                return map[name];
            }
            return name;
        };
        OTelLogger.getHost = function () {
            var host = OSF._OfficeAppFactory.getHostInfo().hostType;
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
            var flavor = OSF._OfficeAppFactory.getHostInfo().hostPlatform;
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
            var timeoutScriptLoadMilliseconds = 15000;
            var afterOnReady = function () {
                if ((typeof oteljs === "undefined") || (typeof oteljs_agave === "undefined")) {
                    console.error("oteljs.js or oteljs_agave.js is not loaded");
                    return;
                }
                if (!OTelLogger.loaded()) {
                    OSF.Utility.debugLog("Creating OTelLogger");
                    OTelLogger.logger = OTelLogger.create(info);
                }
                if (OTelLogger.loaded()) {
                    OTelLogger.promises.forEach(function (resolve) {
                        resolve();
                    });
                }
            };
            var afterLoadOtelSink = function (success) {
                if (success) {
                    Office.onReadyInternal().then(function () {
                        setTimeout(afterOnReady, 0);
                    });
                }
                else {
                    console.error("Cannot load " + OTelLogger.getOtelSinkCDNLocation());
                }
            };
            if (typeof (window.oteljs) !== 'undefined') {
                OSF.OUtil.loadScript(OTelLogger.getOtelSinkCDNLocation(), afterLoadOtelSink, timeoutScriptLoadMilliseconds);
            }
            else {
                OSF.OUtil.loadScript(OTelLogger.getOtelCDNLocation(), function (success) {
                    if (success) {
                        OSF.OUtil.loadScript(OTelLogger.getOtelSinkCDNLocation(), afterLoadOtelSink, timeoutScriptLoadMilliseconds);
                    }
                    else {
                        console.error("Cannot load " + OTelLogger.getOtelCDNLocation());
                    }
                }, timeoutScriptLoadMilliseconds);
            }
        };
        OTelLogger.sendTelemetryEvent = function (telemetryEvent) {
            OTelLogger.onTelemetryLoaded(function () {
                try {
                    OTelLogger.logger.sendTelemetryEvent(telemetryEvent);
                    OSF.Utility.debugLog("Sent telemetry");
                }
                catch (e) {
                    console.error("Cannot send telemetry event: " + JSON.stringify(e));
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
    }());
    OTel.OTelLogger = OTelLogger;
})(OTel || (OTel = {}));
var OSF;
(function (OSF) {
    var AppTelemetry;
    (function (AppTelemetry) {
        var UrlFilter = (function () {
            function UrlFilter() {
            }
            UrlFilter.hashString = function (s) {
                var hash = 0;
                if (s.length === 0) {
                    return hash;
                }
                for (var i = 0; i < s.length; i++) {
                    var c = s.charCodeAt(i);
                    hash = ((hash << 5) - hash) + c;
                    hash |= 0;
                }
                return hash;
            };
            ;
            UrlFilter.stringToHash = function (s) {
                var hash = UrlFilter.hashString(s);
                var stringHash = hash.toString();
                if (hash < 0) {
                    stringHash = "1" + stringHash.substring(1);
                }
                else {
                    stringHash = "0" + stringHash;
                }
                return stringHash;
            };
            UrlFilter.startsWith = function (s, prefix) {
                return s.indexOf(prefix) == -0;
            };
            UrlFilter.isFileUrl = function (url) {
                return UrlFilter.startsWith(url.toLowerCase(), "file:");
            };
            UrlFilter.removeHttpPrefix = function (url) {
                var prefix = "";
                if (UrlFilter.startsWith(url.toLowerCase(), UrlFilter.httpsPrefix)) {
                    prefix = UrlFilter.httpsPrefix;
                }
                else if (UrlFilter.startsWith(url.toLowerCase(), UrlFilter.httpPrefix)) {
                    prefix = UrlFilter.httpPrefix;
                }
                var clean = url.slice(prefix.length);
                return clean;
            };
            UrlFilter.getUrlDomain = function (url) {
                var domain = UrlFilter.removeHttpPrefix(url);
                domain = domain.split("/")[0];
                domain = domain.split(":")[0];
                return domain;
            };
            UrlFilter.isIp4Address = function (domain) {
                var ipv4Regex = /^(25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)\.(25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)\.(25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)\.(25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)$/;
                return ipv4Regex.test(domain);
            };
            UrlFilter.filter = function (url) {
                if (UrlFilter.isFileUrl(url)) {
                    var hash = UrlFilter.stringToHash(url);
                    return "file://" + hash;
                }
                var domain = UrlFilter.getUrlDomain(url);
                if (UrlFilter.isIp4Address(domain)) {
                    var hash = UrlFilter.stringToHash(url);
                    if (UrlFilter.startsWith(domain, "10.")) {
                        return "IP10Range_" + hash;
                    }
                    else if (UrlFilter.startsWith(domain, "192.")) {
                        return "IP192Range_" + hash;
                    }
                    else if (UrlFilter.startsWith(domain, "127.")) {
                        return "IP127Range_" + hash;
                    }
                    return "IPOther_" + hash;
                }
                return domain;
            };
            UrlFilter.httpPrefix = "http://";
            UrlFilter.httpsPrefix = "https://";
            return UrlFilter;
        }());
        AppTelemetry.UrlFilter = UrlFilter;
    })(AppTelemetry = OSF.AppTelemetry || (OSF.AppTelemetry = {}));
})(OSF || (OSF = {}));
var OSF;
(function (OSF) {
    var isBrowser = new Function("try { return typeof(window) !== 'undefined' && this===window; } catch(e) { return false; }");
    if (isBrowser()) {
        OSF._OfficeAppFactory.bootstrap(function () { }, function (e) {
            console.error(JSON.stringify(e));
        });
        window.addEventListener('DOMContentLoaded', function (event) {
            OSFPerformance.hostSpecificFileName = OSF.LoadScriptHelper.getHostBundleJsName();
            Office.onReadyInternal(function () {
                OSFPerfUtil.sendPerformanceTelemetry();
            });
        });
    }
})(OSF || (OSF = {}));
var OfficeRuntime;
(function (OfficeRuntime) {
    var experimentation;
    (function (experimentation) {
        function getBooleanFeatureGate(featureName, defaultValue) {
            try {
                var featureGates = Microsoft.Office.WebExtension.FeatureGates;
                var featureGateValue = featureGates[featureName];
                return void 0 === featureGateValue || null === featureGateValue ? defaultValue : "true" === featureGateValue.toString().toLowerCase();
            }
            catch (error) {
                return defaultValue;
            }
        }
        experimentation.getBooleanFeatureGate = getBooleanFeatureGate;
        function getIntFeatureGate(featureName, defaultValue) {
            try {
                var featureGates = Microsoft.Office.WebExtension.FeatureGates;
                var featureGateValue = parseInt(featureGates[featureName]);
                return isNaN(featureGateValue) ? defaultValue : featureGateValue;
            }
            catch (error) {
                return defaultValue;
            }
        }
        experimentation.getIntFeatureGate = getIntFeatureGate;
        function getStringFeatureGate(featureName, defaultValue) {
            try {
                var featureGates = Microsoft.Office.WebExtension.FeatureGates;
                var featureGateValue = featureGates[featureName];
                return void 0 === featureGateValue || null === featureGateValue ? defaultValue : featureGateValue;
            }
            catch (error) {
                return defaultValue;
            }
        }
        experimentation.getStringFeatureGate = getStringFeatureGate;
        function getBooleanFeatureGateAsync(featureName, defaultValue) {
            return Promise.resolve(getBooleanFeatureGate(featureName, defaultValue));
        }
        experimentation.getBooleanFeatureGateAsync = getBooleanFeatureGateAsync;
        function getIntFeatureGateAsync(featureName, defaultValue) {
            return Promise.resolve(getIntFeatureGate(featureName, defaultValue));
        }
        experimentation.getIntFeatureGateAsync = getIntFeatureGateAsync;
        function getStringFeatureGateAsync(featureName, defaultValue) {
            return Promise.resolve(getStringFeatureGate(featureName, defaultValue));
        }
        experimentation.getStringFeatureGateAsync = getStringFeatureGateAsync;
    })(experimentation = OfficeRuntime.experimentation || (OfficeRuntime.experimentation = {}));
})(OfficeRuntime || (OfficeRuntime = {}));
var __extends = (this && this.__extends) || (function () {
    var extendStatics = function (d, b) {
        extendStatics = Object.setPrototypeOf ||
            ({ __proto__: [] } instanceof Array && function (d, b) { d.__proto__ = b; }) ||
            function (d, b) { for (var p in b) if (b.hasOwnProperty(p)) d[p] = b[p]; };
        return extendStatics(d, b);
    };
    return function (d, b) {
        extendStatics(d, b);
        function __() { this.constructor = d; }
        d.prototype = b === null ? Object.create(b) : (__.prototype = b.prototype, new __());
    };
})();
var OfficeExtension;
(function (OfficeExtension_1) {
    var SessionBase = (function () {
        function SessionBase() {
        }
        SessionBase.prototype._resolveRequestUrlAndHeaderInfo = function () {
            return CoreUtility._createPromiseFromResult(null);
        };
        SessionBase.prototype._createRequestExecutorOrNull = function () {
            return null;
        };
        Object.defineProperty(SessionBase.prototype, "eventRegistration", {
            get: function () {
                return null;
            },
            enumerable: true,
            configurable: true
        });
        return SessionBase;
    }());
    OfficeExtension_1.SessionBase = SessionBase;
    var HttpUtility = (function () {
        function HttpUtility() {
        }
        HttpUtility.setCustomSendRequestFunc = function (func) {
            HttpUtility.s_customSendRequestFunc = func;
        };
        HttpUtility.xhrSendRequestFunc = function (request) {
            return CoreUtility.createPromise(function (resolve, reject) {
                var xhr = new XMLHttpRequest();
                xhr.open(request.method, request.url);
                xhr.onload = function () {
                    var resp = {
                        statusCode: xhr.status,
                        headers: CoreUtility._parseHttpResponseHeaders(xhr.getAllResponseHeaders()),
                        body: xhr.responseText
                    };
                    resolve(resp);
                };
                xhr.onerror = function () {
                    reject(new _Internal.RuntimeError({
                        code: CoreErrorCodes.connectionFailure,
                        message: CoreUtility._getResourceString(CoreResourceStrings.connectionFailureWithStatus, xhr.statusText)
                    }));
                };
                if (request.headers) {
                    for (var key in request.headers) {
                        xhr.setRequestHeader(key, request.headers[key]);
                    }
                }
                xhr.send(CoreUtility._getRequestBodyText(request));
            });
        };
        HttpUtility.sendRequest = function (request) {
            HttpUtility.validateAndNormalizeRequest(request);
            var func = HttpUtility.s_customSendRequestFunc;
            if (!func) {
                func = HttpUtility.xhrSendRequestFunc;
            }
            return func(request);
        };
        HttpUtility.setCustomSendLocalDocumentRequestFunc = function (func) {
            HttpUtility.s_customSendLocalDocumentRequestFunc = func;
        };
        HttpUtility.sendLocalDocumentRequest = function (request) {
            HttpUtility.validateAndNormalizeRequest(request);
            var func;
            func = HttpUtility.s_customSendLocalDocumentRequestFunc || HttpUtility.officeJsSendLocalDocumentRequestFunc;
            return func(request);
        };
        HttpUtility.officeJsSendLocalDocumentRequestFunc = function (request) {
            request = CoreUtility._validateLocalDocumentRequest(request);
            var requestSafeArray = CoreUtility._buildRequestMessageSafeArray(request);
            return CoreUtility.createPromise(function (resolve, reject) {
                OSF.DDA.RichApi.executeRichApiRequestAsync(requestSafeArray, function (asyncResult) {
                    var response;
                    if (asyncResult.status == 'succeeded') {
                        response = {
                            statusCode: RichApiMessageUtility.getResponseStatusCode(asyncResult),
                            headers: RichApiMessageUtility.getResponseHeaders(asyncResult),
                            body: RichApiMessageUtility.getResponseBody(asyncResult)
                        };
                    }
                    else {
                        response = RichApiMessageUtility.buildHttpResponseFromOfficeJsError(asyncResult.error.code, asyncResult.error.message);
                    }
                    CoreUtility.log('Response:');
                    CoreUtility.log(JSON.stringify(response));
                    resolve(response);
                });
            });
        };
        HttpUtility.validateAndNormalizeRequest = function (request) {
            if (CoreUtility.isNullOrUndefined(request)) {
                throw _Internal.RuntimeError._createInvalidArgError({
                    argumentName: 'request'
                });
            }
            if (CoreUtility.isNullOrEmptyString(request.method)) {
                request.method = 'GET';
            }
            request.method = request.method.toUpperCase();
        };
        HttpUtility.logRequest = function (request) {
            if (CoreUtility._logEnabled) {
                CoreUtility.log('---HTTP Request---');
                CoreUtility.log(request.method + ' ' + request.url);
                if (request.headers) {
                    for (var key in request.headers) {
                        CoreUtility.log(key + ': ' + request.headers[key]);
                    }
                }
                if (HttpUtility._logBodyEnabled) {
                    CoreUtility.log(CoreUtility._getRequestBodyText(request));
                }
            }
        };
        HttpUtility.logResponse = function (response) {
            if (CoreUtility._logEnabled) {
                CoreUtility.log('---HTTP Response---');
                CoreUtility.log('' + response.statusCode);
                if (response.headers) {
                    for (var key in response.headers) {
                        CoreUtility.log(key + ': ' + response.headers[key]);
                    }
                }
                if (HttpUtility._logBodyEnabled) {
                    CoreUtility.log(response.body);
                }
            }
        };
        HttpUtility._logBodyEnabled = false;
        return HttpUtility;
    }());
    OfficeExtension_1.HttpUtility = HttpUtility;
    var HostBridge = (function () {
        function HostBridge(m_bridge) {
            var _this = this;
            this.m_bridge = m_bridge;
            this.m_promiseResolver = {};
            this.m_handlers = [];
            this.m_bridge.onMessageFromHost = function (messageText) {
                var message = JSON.parse(messageText);
                if (message.type == 3) {
                    var genericMessageBody = message.message;
                    if (genericMessageBody && genericMessageBody.entries) {
                        for (var i = 0; i < genericMessageBody.entries.length; i++) {
                            var entryObjectOrArray = genericMessageBody.entries[i];
                            if (Array.isArray(entryObjectOrArray)) {
                                var entry = {
                                    messageCategory: entryObjectOrArray[0],
                                    messageType: entryObjectOrArray[1],
                                    targetId: entryObjectOrArray[2],
                                    message: entryObjectOrArray[3],
                                    id: entryObjectOrArray[4]
                                };
                                genericMessageBody.entries[i] = entry;
                            }
                        }
                    }
                }
                _this.dispatchMessage(message);
            };
        }
        HostBridge.init = function (bridge) {
            if (typeof bridge !== 'object' || !bridge) {
                return;
            }
            var instance = new HostBridge(bridge);
            HostBridge.s_instance = instance;
            HttpUtility.setCustomSendLocalDocumentRequestFunc(function (request) {
                request = CoreUtility._validateLocalDocumentRequest(request);
                var requestFlags = 0;
                if (!CoreUtility.isReadonlyRestRequest(request.method)) {
                    requestFlags = 1;
                }
                var index = request.url.indexOf('?');
                if (index >= 0) {
                    var query = request.url.substr(index + 1);
                    var flagsAndCustomData = CoreUtility._parseRequestFlagsAndCustomDataFromQueryStringIfAny(query);
                    if (flagsAndCustomData.flags >= 0) {
                        requestFlags = flagsAndCustomData.flags;
                    }
                }
                var bridgeMessage = {
                    id: HostBridge.nextId(),
                    type: 1,
                    flags: requestFlags,
                    message: request
                };
                return instance.sendMessageToHostAndExpectResponse(bridgeMessage).then(function (bridgeResponse) {
                    var responseInfo = bridgeResponse.message;
                    return responseInfo;
                });
            });
            for (var i = 0; i < HostBridge.s_onInitedHandlers.length; i++) {
                HostBridge.s_onInitedHandlers[i](instance);
            }
        };
        Object.defineProperty(HostBridge, "instance", {
            get: function () {
                return HostBridge.s_instance;
            },
            enumerable: true,
            configurable: true
        });
        HostBridge.prototype.sendMessageToHost = function (message) {
            this.m_bridge.sendMessageToHost(JSON.stringify(message));
        };
        HostBridge.prototype.sendMessageToHostAndExpectResponse = function (message) {
            var _this = this;
            var ret = CoreUtility.createPromise(function (resolve, reject) {
                _this.m_promiseResolver[message.id] = resolve;
            });
            this.m_bridge.sendMessageToHost(JSON.stringify(message));
            return ret;
        };
        HostBridge.prototype.addHostMessageHandler = function (handler) {
            this.m_handlers.push(handler);
        };
        HostBridge.prototype.removeHostMessageHandler = function (handler) {
            var index = this.m_handlers.indexOf(handler);
            if (index >= 0) {
                this.m_handlers.splice(index, 1);
            }
        };
        HostBridge.onInited = function (handler) {
            HostBridge.s_onInitedHandlers.push(handler);
            if (HostBridge.s_instance) {
                handler(HostBridge.s_instance);
            }
        };
        HostBridge.prototype.dispatchMessage = function (message) {
            if (typeof message.id === 'number') {
                var resolve = this.m_promiseResolver[message.id];
                if (resolve) {
                    resolve(message);
                    delete this.m_promiseResolver[message.id];
                    return;
                }
            }
            for (var i = 0; i < this.m_handlers.length; i++) {
                this.m_handlers[i](message);
            }
        };
        HostBridge.nextId = function () {
            return HostBridge.s_nextId++;
        };
        HostBridge.s_onInitedHandlers = [];
        HostBridge.s_nextId = 1;
        return HostBridge;
    }());
    OfficeExtension_1.HostBridge = HostBridge;
    if (typeof _richApiNativeBridge === 'object' && _richApiNativeBridge) {
        HostBridge.init(_richApiNativeBridge);
    }
    var _Internal;
    (function (_Internal) {
        var RuntimeError = (function (_super) {
            __extends(RuntimeError, _super);
            function RuntimeError(error) {
                var _this = _super.call(this, typeof error === 'string' ? error : error.message) || this;
                Object.setPrototypeOf(_this, RuntimeError.prototype);
                _this.name = 'RichApi.Error';
                if (typeof error === 'string') {
                    _this.message = error;
                }
                else {
                    _this.code = error.code;
                    _this.message = error.message;
                    _this.traceMessages = error.traceMessages || [];
                    _this.innerError = error.innerError || null;
                    _this.debugInfo = _this._createDebugInfo(error.debugInfo || {});
                }
                return _this;
            }
            RuntimeError.prototype.toString = function () {
                return this.code + ': ' + this.message;
            };
            RuntimeError.prototype._createDebugInfo = function (partialDebugInfo) {
                var debugInfo = {
                    code: this.code,
                    message: this.message
                };
                debugInfo.toString = function () {
                    return JSON.stringify(this);
                };
                for (var key in partialDebugInfo) {
                    debugInfo[key] = partialDebugInfo[key];
                }
                if (this.innerError) {
                    if (this.innerError instanceof _Internal.RuntimeError) {
                        debugInfo.innerError = this.innerError.debugInfo;
                    }
                    else {
                        debugInfo.innerError = this.innerError;
                    }
                }
                return debugInfo;
            };
            RuntimeError._createInvalidArgError = function (error) {
                return new _Internal.RuntimeError({
                    code: CoreErrorCodes.invalidArgument,
                    message: CoreUtility.isNullOrEmptyString(error.argumentName)
                        ? CoreUtility._getResourceString(CoreResourceStrings.invalidArgumentGeneric)
                        : CoreUtility._getResourceString(CoreResourceStrings.invalidArgument, error.argumentName),
                    debugInfo: error.errorLocation ? { errorLocation: error.errorLocation } : {},
                    innerError: error.innerError
                });
            };
            return RuntimeError;
        }(Error));
        _Internal.RuntimeError = RuntimeError;
    })(_Internal = OfficeExtension_1._Internal || (OfficeExtension_1._Internal = {}));
    OfficeExtension_1.Error = _Internal.RuntimeError;
    var CoreErrorCodes = (function () {
        function CoreErrorCodes() {
        }
        CoreErrorCodes.apiNotFound = 'ApiNotFound';
        CoreErrorCodes.accessDenied = 'AccessDenied';
        CoreErrorCodes.generalException = 'GeneralException';
        CoreErrorCodes.activityLimitReached = 'ActivityLimitReached';
        CoreErrorCodes.invalidArgument = 'InvalidArgument';
        CoreErrorCodes.connectionFailure = 'ConnectionFailure';
        CoreErrorCodes.timeout = 'Timeout';
        CoreErrorCodes.invalidOrTimedOutSession = 'InvalidOrTimedOutSession';
        CoreErrorCodes.invalidObjectPath = 'InvalidObjectPath';
        CoreErrorCodes.invalidRequestContext = 'InvalidRequestContext';
        CoreErrorCodes.valueNotLoaded = 'ValueNotLoaded';
        return CoreErrorCodes;
    }());
    OfficeExtension_1.CoreErrorCodes = CoreErrorCodes;
    var CoreResourceStrings = (function () {
        function CoreResourceStrings() {
        }
        CoreResourceStrings.apiNotFoundDetails = 'ApiNotFoundDetails';
        CoreResourceStrings.connectionFailureWithStatus = 'ConnectionFailureWithStatus';
        CoreResourceStrings.connectionFailureWithDetails = 'ConnectionFailureWithDetails';
        CoreResourceStrings.invalidArgument = 'InvalidArgument';
        CoreResourceStrings.invalidArgumentGeneric = 'InvalidArgumentGeneric';
        CoreResourceStrings.timeout = 'Timeout';
        CoreResourceStrings.invalidOrTimedOutSessionMessage = 'InvalidOrTimedOutSessionMessage';
        CoreResourceStrings.invalidObjectPath = 'InvalidObjectPath';
        CoreResourceStrings.invalidRequestContext = 'InvalidRequestContext';
        CoreResourceStrings.valueNotLoaded = 'ValueNotLoaded';
        return CoreResourceStrings;
    }());
    OfficeExtension_1.CoreResourceStrings = CoreResourceStrings;
    var CoreConstants = (function () {
        function CoreConstants() {
        }
        CoreConstants.flags = 'flags';
        CoreConstants.sourceLibHeader = 'SdkVersion';
        CoreConstants.processQuery = 'ProcessQuery';
        CoreConstants.localDocument = 'http://document.localhost/';
        CoreConstants.localDocumentApiPrefix = 'http://document.localhost/_api/';
        CoreConstants.customData = 'customdata';
        return CoreConstants;
    }());
    OfficeExtension_1.CoreConstants = CoreConstants;
    var RichApiMessageUtility = (function () {
        function RichApiMessageUtility() {
        }
        RichApiMessageUtility.buildMessageArrayForIRequestExecutor = function (customData, requestFlags, requestMessage, sourceLibHeaderValue) {
            var requestMessageText = JSON.stringify(requestMessage.Body);
            CoreUtility.log('Request:');
            CoreUtility.log(requestMessageText);
            var headers = {};
            CoreUtility._copyHeaders(requestMessage.Headers, headers);
            headers[CoreConstants.sourceLibHeader] = sourceLibHeaderValue;
            var messageSafearray = RichApiMessageUtility.buildRequestMessageSafeArray(customData, requestFlags, 'POST', CoreConstants.processQuery, headers, requestMessageText);
            return messageSafearray;
        };
        RichApiMessageUtility.buildResponseOnSuccess = function (responseBody, responseHeaders) {
            var response = { ErrorCode: '', ErrorMessage: '', Headers: null, Body: null };
            response.Body = JSON.parse(responseBody);
            response.Headers = responseHeaders;
            return response;
        };
        RichApiMessageUtility.buildResponseOnError = function (errorCode, message) {
            var response = { ErrorCode: '', ErrorMessage: '', Headers: null, Body: null };
            response.ErrorCode = CoreErrorCodes.generalException;
            response.ErrorMessage = message;
            if (errorCode == RichApiMessageUtility.OfficeJsErrorCode_ooeNoCapability) {
                response.ErrorCode = CoreErrorCodes.accessDenied;
            }
            else if (errorCode == RichApiMessageUtility.OfficeJsErrorCode_ooeActivityLimitReached) {
                response.ErrorCode = CoreErrorCodes.activityLimitReached;
            }
            else if (errorCode == RichApiMessageUtility.OfficeJsErrorCode_ooeInvalidOrTimedOutSession) {
                response.ErrorCode = CoreErrorCodes.invalidOrTimedOutSession;
                response.ErrorMessage = CoreUtility._getResourceString(CoreResourceStrings.invalidOrTimedOutSessionMessage);
            }
            return response;
        };
        RichApiMessageUtility.buildHttpResponseFromOfficeJsError = function (errorCode, message) {
            var statusCode = 500;
            var errorBody = {};
            errorBody['error'] = {};
            errorBody['error']['code'] = CoreErrorCodes.generalException;
            errorBody['error']['message'] = message;
            if (errorCode === RichApiMessageUtility.OfficeJsErrorCode_ooeNoCapability) {
                statusCode = 403;
                errorBody['error']['code'] = CoreErrorCodes.accessDenied;
            }
            else if (errorCode === RichApiMessageUtility.OfficeJsErrorCode_ooeActivityLimitReached) {
                statusCode = 429;
                errorBody['error']['code'] = CoreErrorCodes.activityLimitReached;
            }
            return { statusCode: statusCode, headers: {}, body: JSON.stringify(errorBody) };
        };
        RichApiMessageUtility.buildRequestMessageSafeArray = function (customData, requestFlags, method, path, headers, body) {
            var headerArray = [];
            if (headers) {
                for (var headerName in headers) {
                    headerArray.push(headerName);
                    headerArray.push(headers[headerName]);
                }
            }
            var appPermission = 0;
            var solutionId = '';
            var instanceId = '';
            var marketplaceType = '';
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
        RichApiMessageUtility.getResponseBody = function (result) {
            return RichApiMessageUtility.getResponseBodyFromSafeArray(result.value.data);
        };
        RichApiMessageUtility.getResponseHeaders = function (result) {
            return RichApiMessageUtility.getResponseHeadersFromSafeArray(result.value.data);
        };
        RichApiMessageUtility.getResponseBodyFromSafeArray = function (data) {
            var ret = data[2];
            if (typeof ret === 'string') {
                return ret;
            }
            var arr = ret;
            return arr.join('');
        };
        RichApiMessageUtility.getResponseHeadersFromSafeArray = function (data) {
            var arrayHeader = data[1];
            if (!arrayHeader) {
                return null;
            }
            var headers = {};
            for (var i = 0; i < arrayHeader.length - 1; i += 2) {
                headers[arrayHeader[i]] = arrayHeader[i + 1];
            }
            return headers;
        };
        RichApiMessageUtility.getResponseStatusCode = function (result) {
            return RichApiMessageUtility.getResponseStatusCodeFromSafeArray(result.value.data);
        };
        RichApiMessageUtility.getResponseStatusCodeFromSafeArray = function (data) {
            return data[0];
        };
        RichApiMessageUtility.OfficeJsErrorCode_ooeInvalidOrTimedOutSession = 5012;
        RichApiMessageUtility.OfficeJsErrorCode_ooeActivityLimitReached = 5102;
        RichApiMessageUtility.OfficeJsErrorCode_ooeNoCapability = 7000;
        return RichApiMessageUtility;
    }());
    OfficeExtension_1.RichApiMessageUtility = RichApiMessageUtility;
    (function (_Internal) {
        function getPromiseType() {
            if (typeof Promise !== 'undefined') {
                return Promise;
            }
            if (typeof Office !== 'undefined') {
                if (Office.Promise) {
                    return Office.Promise;
                }
            }
            if (typeof OfficeExtension !== 'undefined') {
                if (OfficeExtension.Promise) {
                    return OfficeExtension.Promise;
                }
            }
            throw new _Internal.Error('No Promise implementation found');
        }
        _Internal.getPromiseType = getPromiseType;
    })(_Internal = OfficeExtension_1._Internal || (OfficeExtension_1._Internal = {}));
    var CoreUtility = (function () {
        function CoreUtility() {
        }
        CoreUtility.log = function (message) {
            if (CoreUtility._logEnabled && typeof console !== 'undefined' && console.log) {
                console.log(message);
            }
        };
        CoreUtility.checkArgumentNull = function (value, name) {
            if (CoreUtility.isNullOrUndefined(value)) {
                throw _Internal.RuntimeError._createInvalidArgError({ argumentName: name });
            }
        };
        CoreUtility.isNullOrUndefined = function (value) {
            if (value === null) {
                return true;
            }
            if (typeof value === 'undefined') {
                return true;
            }
            return false;
        };
        CoreUtility.isUndefined = function (value) {
            if (typeof value === 'undefined') {
                return true;
            }
            return false;
        };
        CoreUtility.isNullOrEmptyString = function (value) {
            if (value === null) {
                return true;
            }
            if (typeof value === 'undefined') {
                return true;
            }
            if (value.length == 0) {
                return true;
            }
            return false;
        };
        CoreUtility.isPlainJsonObject = function (value) {
            if (CoreUtility.isNullOrUndefined(value)) {
                return false;
            }
            if (typeof value !== 'object') {
                return false;
            }
            if (Object.prototype.toString.apply(value) !== '[object Object]') {
                return false;
            }
            if (value.constructor &&
                !Object.prototype.hasOwnProperty.call(value, 'constructor') &&
                !Object.prototype.hasOwnProperty.call(value.constructor.prototype, 'hasOwnProperty')) {
                return false;
            }
            for (var key in value) {
                if (!Object.prototype.hasOwnProperty.call(value, key)) {
                    return false;
                }
            }
            return true;
        };
        CoreUtility.trim = function (str) {
            return str.replace(new RegExp('^\\s+|\\s+$', 'g'), '');
        };
        CoreUtility.caseInsensitiveCompareString = function (str1, str2) {
            if (CoreUtility.isNullOrUndefined(str1)) {
                return CoreUtility.isNullOrUndefined(str2);
            }
            else {
                if (CoreUtility.isNullOrUndefined(str2)) {
                    return false;
                }
                else {
                    return str1.toUpperCase() == str2.toUpperCase();
                }
            }
        };
        CoreUtility.isReadonlyRestRequest = function (method) {
            return CoreUtility.caseInsensitiveCompareString(method, 'GET');
        };
        CoreUtility._getResourceString = function (resourceId, arg) {
            var ret;
            if (typeof window !== 'undefined' && window.Strings && window.Strings.OfficeOM) {
                var stringName = 'L_' + resourceId;
                var stringValue = window.Strings.OfficeOM[stringName];
                if (stringValue) {
                    ret = stringValue;
                }
            }
            if (!ret) {
                ret = CoreUtility.s_resourceStringValues[resourceId];
            }
            if (!ret) {
                ret = resourceId;
            }
            if (!CoreUtility.isNullOrUndefined(arg)) {
                if (Array.isArray(arg)) {
                    var arrArg = arg;
                    ret = CoreUtility._formatString(ret, arrArg);
                }
                else {
                    ret = ret.replace('{0}', arg);
                }
            }
            return ret;
        };
        CoreUtility._formatString = function (format, arrArg) {
            return format.replace(/\{\d\}/g, function (v) {
                var position = parseInt(v.substr(1, v.length - 2));
                if (position < arrArg.length) {
                    return arrArg[position];
                }
                else {
                    throw _Internal.RuntimeError._createInvalidArgError({ argumentName: 'format' });
                }
            });
        };
        Object.defineProperty(CoreUtility, "Promise", {
            get: function () {
                return _Internal.getPromiseType();
            },
            enumerable: true,
            configurable: true
        });
        CoreUtility.createPromise = function (executor) {
            var ret = new CoreUtility.Promise(executor);
            return ret;
        };
        CoreUtility._createPromiseFromResult = function (value) {
            return CoreUtility.createPromise(function (resolve, reject) {
                resolve(value);
            });
        };
        CoreUtility._createPromiseFromException = function (reason) {
            return CoreUtility.createPromise(function (resolve, reject) {
                reject(reason);
            });
        };
        CoreUtility._createTimeoutPromise = function (timeout) {
            return CoreUtility.createPromise(function (resolve, reject) {
                setTimeout(function () {
                    resolve(null);
                }, timeout);
            });
        };
        CoreUtility._createInvalidArgError = function (error) {
            return _Internal.RuntimeError._createInvalidArgError(error);
        };
        CoreUtility._isLocalDocumentUrl = function (url) {
            return CoreUtility._getLocalDocumentUrlPrefixLength(url) > 0;
        };
        CoreUtility._getLocalDocumentUrlPrefixLength = function (url) {
            var localDocumentPrefixes = [
                'http://document.localhost',
                'https://document.localhost',
                '//document.localhost'
            ];
            var urlLower = url.toLowerCase().trim();
            for (var i = 0; i < localDocumentPrefixes.length; i++) {
                if (urlLower === localDocumentPrefixes[i]) {
                    return localDocumentPrefixes[i].length;
                }
                else if (urlLower.substr(0, localDocumentPrefixes[i].length + 1) === localDocumentPrefixes[i] + '/') {
                    return localDocumentPrefixes[i].length + 1;
                }
            }
            return 0;
        };
        CoreUtility._validateLocalDocumentRequest = function (request) {
            var index = CoreUtility._getLocalDocumentUrlPrefixLength(request.url);
            if (index <= 0) {
                throw _Internal.RuntimeError._createInvalidArgError({
                    argumentName: 'request'
                });
            }
            var path = request.url.substr(index);
            var pathLower = path.toLowerCase();
            if (pathLower === '_api') {
                path = '';
            }
            else if (pathLower.substr(0, '_api/'.length) === '_api/') {
                path = path.substr('_api/'.length);
            }
            return {
                method: request.method,
                url: path,
                headers: request.headers,
                body: request.body
            };
        };
        CoreUtility._parseRequestFlagsAndCustomDataFromQueryStringIfAny = function (queryString) {
            var ret = { flags: -1, customData: '' };
            var parts = queryString.split('&');
            for (var i = 0; i < parts.length; i++) {
                var keyvalue = parts[i].split('=');
                if (keyvalue[0].toLowerCase() === CoreConstants.flags) {
                    var flags = parseInt(keyvalue[1]);
                    flags = flags & 4095;
                    ret.flags = flags;
                }
                else if (keyvalue[0].toLowerCase() === CoreConstants.customData) {
                    ret.customData = decodeURIComponent(keyvalue[1]);
                }
            }
            return ret;
        };
        CoreUtility._getRequestBodyText = function (request) {
            var body = '';
            if (typeof request.body === 'string') {
                body = request.body;
            }
            else if (request.body && typeof request.body === 'object') {
                body = JSON.stringify(request.body);
            }
            return body;
        };
        CoreUtility._parseResponseBody = function (response) {
            if (typeof response.body === 'string') {
                var bodyText = CoreUtility.trim(response.body);
                return JSON.parse(bodyText);
            }
            else {
                return response.body;
            }
        };
        CoreUtility._buildRequestMessageSafeArray = function (request) {
            var requestFlags = 0;
            if (!CoreUtility.isReadonlyRestRequest(request.method)) {
                requestFlags = 1;
            }
            var customData = '';
            if (request.url.substr(0, CoreConstants.processQuery.length).toLowerCase() ===
                CoreConstants.processQuery.toLowerCase()) {
                var index = request.url.indexOf('?');
                if (index > 0) {
                    var queryString = request.url.substr(index + 1);
                    var flagsAndCustomData = CoreUtility._parseRequestFlagsAndCustomDataFromQueryStringIfAny(queryString);
                    if (flagsAndCustomData.flags >= 0) {
                        requestFlags = flagsAndCustomData.flags;
                    }
                    customData = flagsAndCustomData.customData;
                }
            }
            return RichApiMessageUtility.buildRequestMessageSafeArray(customData, requestFlags, request.method, request.url, request.headers, CoreUtility._getRequestBodyText(request));
        };
        CoreUtility._parseHttpResponseHeaders = function (allResponseHeaders) {
            var responseHeaders = {};
            if (!CoreUtility.isNullOrEmptyString(allResponseHeaders)) {
                var regex = new RegExp('\r?\n');
                var entries = allResponseHeaders.split(regex);
                for (var i = 0; i < entries.length; i++) {
                    var entry = entries[i];
                    if (entry != null) {
                        var index = entry.indexOf(':');
                        if (index > 0) {
                            var key = entry.substr(0, index);
                            var value = entry.substr(index + 1);
                            key = CoreUtility.trim(key);
                            value = CoreUtility.trim(value);
                            responseHeaders[key.toUpperCase()] = value;
                        }
                    }
                }
            }
            return responseHeaders;
        };
        CoreUtility._parseErrorResponse = function (responseInfo) {
            var errorObj = null;
            if (CoreUtility.isPlainJsonObject(responseInfo.body)) {
                errorObj = responseInfo.body;
            }
            else if (!CoreUtility.isNullOrEmptyString(responseInfo.body)) {
                var errorResponseBody = CoreUtility.trim(responseInfo.body);
                try {
                    errorObj = JSON.parse(errorResponseBody);
                }
                catch (e) {
                    CoreUtility.log('Error when parse ' + errorResponseBody);
                }
            }
            var errorMessage;
            var errorCode;
            if (!CoreUtility.isNullOrUndefined(errorObj) && typeof errorObj === 'object' && errorObj.error) {
                errorCode = errorObj.error.code;
                errorMessage = CoreUtility._getResourceString(CoreResourceStrings.connectionFailureWithDetails, [
                    responseInfo.statusCode.toString(),
                    errorObj.error.code,
                    errorObj.error.message
                ]);
            }
            else {
                errorMessage = CoreUtility._getResourceString(CoreResourceStrings.connectionFailureWithStatus, responseInfo.statusCode.toString());
            }
            if (CoreUtility.isNullOrEmptyString(errorCode)) {
                errorCode = CoreErrorCodes.connectionFailure;
            }
            return { errorCode: errorCode, errorMessage: errorMessage };
        };
        CoreUtility._copyHeaders = function (src, dest) {
            if (src && dest) {
                for (var key in src) {
                    dest[key] = src[key];
                }
            }
        };
        CoreUtility.addResourceStringValues = function (values) {
            for (var key in values) {
                CoreUtility.s_resourceStringValues[key] = values[key];
            }
        };
        CoreUtility._logEnabled = false;
        CoreUtility.s_resourceStringValues = {
            ApiNotFoundDetails: 'The method or property {0} is part of the {1} requirement set, which is not available in your version of {2}.',
            ConnectionFailureWithStatus: 'The request failed with status code of {0}.',
            ConnectionFailureWithDetails: 'The request failed with status code of {0}, error code {1} and the following error message: {2}',
            InvalidArgument: "The argument '{0}' doesn't work for this situation, is missing, or isn't in the right format.",
            InvalidObjectPath: 'The object path \'{0}\' isn\'t working for what you\'re trying to do. If you\'re using the object across multiple "context.sync" calls and outside the sequential execution of a ".run" batch, please use the "context.trackedObjects.add()" and "context.trackedObjects.remove()" methods to manage the object\'s lifetime.',
            InvalidRequestContext: 'Cannot use the object across different request contexts.',
            Timeout: 'The operation has timed out.',
            ValueNotLoaded: 'The value of the result object has not been loaded yet. Before reading the value property, call "context.sync()" on the associated request context.'
        };
        return CoreUtility;
    }());
    OfficeExtension_1.CoreUtility = CoreUtility;
    var TestUtility = (function () {
        function TestUtility() {
        }
        TestUtility.setMock = function (value) {
            TestUtility.s_isMock = value;
        };
        TestUtility.isMock = function () {
            return TestUtility.s_isMock;
        };
        return TestUtility;
    }());
    OfficeExtension_1.TestUtility = TestUtility;
    OfficeExtension_1._internalConfig = {
        showDisposeInfoInDebugInfo: false,
        showInternalApiInDebugInfo: false,
        enableEarlyDispose: true,
        alwaysPolyfillClientObjectUpdateMethod: false,
        alwaysPolyfillClientObjectRetrieveMethod: false,
        enableConcurrentFlag: true,
        enableUndoableFlag: true
    };
    OfficeExtension_1.config = {
        extendedErrorLogging: false
    };
    var CommonActionFactory = (function () {
        function CommonActionFactory() {
        }
        CommonActionFactory.createSetPropertyAction = function (context, parent, propertyName, value, flags) {
            CommonUtility.validateObjectPath(parent);
            var actionInfo = {
                Id: context._nextId(),
                ActionType: 4,
                Name: propertyName,
                ObjectPathId: parent._objectPath.objectPathInfo.Id,
                ArgumentInfo: {}
            };
            var args = [value];
            var referencedArgumentObjectPaths = CommonUtility.setMethodArguments(context, actionInfo.ArgumentInfo, args);
            CommonUtility.validateReferencedObjectPaths(referencedArgumentObjectPaths);
            var action = new Action(actionInfo, 0, flags);
            action.referencedObjectPath = parent._objectPath;
            action.referencedArgumentObjectPaths = referencedArgumentObjectPaths;
            return parent._addAction(action);
        };
        CommonActionFactory.createQueryAction = function (context, parent, queryOption, resultHandler) {
            CommonUtility.validateObjectPath(parent);
            var actionInfo = {
                Id: context._nextId(),
                ActionType: 2,
                Name: '',
                ObjectPathId: parent._objectPath.objectPathInfo.Id,
                QueryInfo: queryOption
            };
            var action = new Action(actionInfo, 1, 4);
            action.referencedObjectPath = parent._objectPath;
            return parent._addAction(action, resultHandler);
        };
        CommonActionFactory.createQueryAsJsonAction = function (context, parent, queryOption, resultHandler) {
            CommonUtility.validateObjectPath(parent);
            var actionInfo = {
                Id: context._nextId(),
                ActionType: 7,
                Name: '',
                ObjectPathId: parent._objectPath.objectPathInfo.Id,
                QueryInfo: queryOption
            };
            var action = new Action(actionInfo, 1, 4);
            action.referencedObjectPath = parent._objectPath;
            return parent._addAction(action, resultHandler);
        };
        CommonActionFactory.createUpdateAction = function (context, parent, objectState) {
            CommonUtility.validateObjectPath(parent);
            var actionInfo = {
                Id: context._nextId(),
                ActionType: 9,
                Name: '',
                ObjectPathId: parent._objectPath.objectPathInfo.Id,
                ObjectState: objectState
            };
            var action = new Action(actionInfo, 0, 0);
            action.referencedObjectPath = parent._objectPath;
            return parent._addAction(action);
        };
        return CommonActionFactory;
    }());
    OfficeExtension_1.CommonActionFactory = CommonActionFactory;
    var ClientObjectBase = (function () {
        function ClientObjectBase(contextBase, objectPath) {
            this.m_contextBase = contextBase;
            this.m_objectPath = objectPath;
        }
        Object.defineProperty(ClientObjectBase.prototype, "_objectPath", {
            get: function () {
                return this.m_objectPath;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(ClientObjectBase.prototype, "_context", {
            get: function () {
                return this.m_contextBase;
            },
            enumerable: true,
            configurable: true
        });
        ClientObjectBase.prototype._addAction = function (action, resultHandler) {
            var _this = this;
            if (resultHandler === void 0) { resultHandler = null; }
            return CoreUtility.createPromise(function (resolve, reject) {
                _this._context._addServiceApiAction(action, resultHandler, resolve, reject);
            });
        };
        ClientObjectBase.prototype._retrieve = function (option, resultHandler) {
            var shouldPolyfill = OfficeExtension_1._internalConfig.alwaysPolyfillClientObjectRetrieveMethod;
            if (!shouldPolyfill) {
                shouldPolyfill = !CommonUtility.isSetSupported('RichApiRuntime', '1.1');
            }
            var queryOption = ClientRequestContextBase._parseQueryOption(option);
            if (shouldPolyfill) {
                return CommonActionFactory.createQueryAction(this._context, this, queryOption, resultHandler);
            }
            return CommonActionFactory.createQueryAsJsonAction(this._context, this, queryOption, resultHandler);
        };
        ClientObjectBase.prototype._recursivelyUpdate = function (properties) {
            var shouldPolyfill = OfficeExtension_1._internalConfig.alwaysPolyfillClientObjectUpdateMethod;
            if (!shouldPolyfill) {
                shouldPolyfill = !CommonUtility.isSetSupported('RichApiRuntime', '1.2');
            }
            try {
                var scalarPropNames = this[CommonConstants.scalarPropertyNames];
                if (!scalarPropNames) {
                    scalarPropNames = [];
                }
                var scalarPropUpdatable = this[CommonConstants.scalarPropertyUpdateable];
                if (!scalarPropUpdatable) {
                    scalarPropUpdatable = [];
                    for (var i = 0; i < scalarPropNames.length; i++) {
                        scalarPropUpdatable.push(false);
                    }
                }
                var navigationPropNames = this[CommonConstants.navigationPropertyNames];
                if (!navigationPropNames) {
                    navigationPropNames = [];
                }
                var scalarProps = {};
                var navigationProps = {};
                var scalarPropCount = 0;
                for (var propName in properties) {
                    var index = scalarPropNames.indexOf(propName);
                    if (index >= 0) {
                        if (!scalarPropUpdatable[index]) {
                            throw new _Internal.RuntimeError({
                                code: CoreErrorCodes.invalidArgument,
                                message: CoreUtility._getResourceString(CommonResourceStrings.attemptingToSetReadOnlyProperty, propName),
                                debugInfo: {
                                    errorLocation: propName
                                }
                            });
                        }
                        scalarProps[propName] = properties[propName];
                        ++scalarPropCount;
                    }
                    else if (navigationPropNames.indexOf(propName) >= 0) {
                        navigationProps[propName] = properties[propName];
                    }
                    else {
                        throw new _Internal.RuntimeError({
                            code: CoreErrorCodes.invalidArgument,
                            message: CoreUtility._getResourceString(CommonResourceStrings.propertyDoesNotExist, propName),
                            debugInfo: {
                                errorLocation: propName
                            }
                        });
                    }
                }
                if (scalarPropCount > 0) {
                    if (shouldPolyfill) {
                        for (var i = 0; i < scalarPropNames.length; i++) {
                            var propName = scalarPropNames[i];
                            var propValue = scalarProps[propName];
                            if (!CommonUtility.isUndefined(propValue)) {
                                CommonActionFactory.createSetPropertyAction(this._context, this, propName, propValue);
                            }
                        }
                    }
                    else {
                        CommonActionFactory.createUpdateAction(this._context, this, scalarProps);
                    }
                }
                for (var propName in navigationProps) {
                    var navigationPropProxy = this[propName];
                    var navigationPropValue = navigationProps[propName];
                    navigationPropProxy._recursivelyUpdate(navigationPropValue);
                }
            }
            catch (innerError) {
                throw new _Internal.RuntimeError({
                    code: CoreErrorCodes.invalidArgument,
                    message: CoreUtility._getResourceString(CoreResourceStrings.invalidArgument, 'properties'),
                    debugInfo: {
                        errorLocation: this._className + '.update'
                    },
                    innerError: innerError
                });
            }
        };
        return ClientObjectBase;
    }());
    OfficeExtension_1.ClientObjectBase = ClientObjectBase;
    var Action = (function () {
        function Action(actionInfo, operationType, flags) {
            this.m_actionInfo = actionInfo;
            this.m_operationType = operationType;
            this.m_flags = flags;
        }
        Object.defineProperty(Action.prototype, "actionInfo", {
            get: function () {
                return this.m_actionInfo;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(Action.prototype, "operationType", {
            get: function () {
                return this.m_operationType;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(Action.prototype, "flags", {
            get: function () {
                return this.m_flags;
            },
            enumerable: true,
            configurable: true
        });
        return Action;
    }());
    OfficeExtension_1.Action = Action;
    var ObjectPath = (function () {
        function ObjectPath(objectPathInfo, parentObjectPath, isCollection, isInvalidAfterRequest, operationType, flags) {
            this.m_objectPathInfo = objectPathInfo;
            this.m_parentObjectPath = parentObjectPath;
            this.m_isCollection = isCollection;
            this.m_isInvalidAfterRequest = isInvalidAfterRequest;
            this.m_isValid = true;
            this.m_operationType = operationType;
            this.m_flags = flags;
        }
        Object.defineProperty(ObjectPath.prototype, "id", {
            get: function () {
                var argumentInfo = this.m_objectPathInfo.ArgumentInfo;
                if (!argumentInfo) {
                    return undefined;
                }
                var argument = argumentInfo.Arguments;
                if (!argument) {
                    return undefined;
                }
                return argument[0];
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(ObjectPath.prototype, "parent", {
            get: function () {
                var parent = this.m_parentObjectPath;
                if (!parent) {
                    return undefined;
                }
                return parent;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(ObjectPath.prototype, "parentId", {
            get: function () {
                return this.parent ? this.parent.id : undefined;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(ObjectPath.prototype, "objectPathInfo", {
            get: function () {
                return this.m_objectPathInfo;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(ObjectPath.prototype, "operationType", {
            get: function () {
                return this.m_operationType;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(ObjectPath.prototype, "flags", {
            get: function () {
                return this.m_flags;
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
                this.m_argumentObjectPaths = value;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(ObjectPath.prototype, "isValid", {
            get: function () {
                return this.m_isValid;
            },
            set: function (value) {
                this.m_isValid = value;
                if (!value &&
                    this.m_objectPathInfo.ObjectPathType === 6 &&
                    this.m_savedObjectPathInfo) {
                    ObjectPath.copyObjectPathInfo(this.m_savedObjectPathInfo.pathInfo, this.m_objectPathInfo);
                    this.m_parentObjectPath = this.m_savedObjectPathInfo.parent;
                    this.m_isValid = true;
                    this.m_savedObjectPathInfo = null;
                }
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(ObjectPath.prototype, "originalObjectPathInfo", {
            get: function () {
                return this.m_originalObjectPathInfo;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(ObjectPath.prototype, "getByIdMethodName", {
            get: function () {
                return this.m_getByIdMethodName;
            },
            set: function (value) {
                this.m_getByIdMethodName = value;
            },
            enumerable: true,
            configurable: true
        });
        ObjectPath.prototype._updateAsNullObject = function () {
            this.resetForUpdateUsingObjectData();
            this.m_objectPathInfo.ObjectPathType = 7;
            this.m_objectPathInfo.Name = '';
            this.m_parentObjectPath = null;
        };
        ObjectPath.prototype.saveOriginalObjectPathInfo = function () {
            if (OfficeExtension_1.config.extendedErrorLogging && !this.m_originalObjectPathInfo) {
                this.m_originalObjectPathInfo = {};
                ObjectPath.copyObjectPathInfo(this.m_objectPathInfo, this.m_originalObjectPathInfo);
            }
        };
        ObjectPath.prototype.updateUsingObjectData = function (value, clientObject) {
            var referenceId = value[CommonConstants.referenceId];
            if (!CoreUtility.isNullOrEmptyString(referenceId)) {
                if (!this.m_savedObjectPathInfo &&
                    !this.isInvalidAfterRequest &&
                    ObjectPath.isRestorableObjectPath(this.m_objectPathInfo.ObjectPathType)) {
                    var pathInfo = {};
                    ObjectPath.copyObjectPathInfo(this.m_objectPathInfo, pathInfo);
                    this.m_savedObjectPathInfo = {
                        pathInfo: pathInfo,
                        parent: this.m_parentObjectPath
                    };
                }
                this.saveOriginalObjectPathInfo();
                this.resetForUpdateUsingObjectData();
                this.m_objectPathInfo.ObjectPathType = 6;
                this.m_objectPathInfo.Name = referenceId;
                delete this.m_objectPathInfo.ParentObjectPathId;
                this.m_parentObjectPath = null;
                return;
            }
            if (clientObject) {
                var collectionPropertyPath = clientObject[CommonConstants.collectionPropertyPath];
                if (!CoreUtility.isNullOrEmptyString(collectionPropertyPath) && clientObject.context) {
                    var id = CommonUtility.tryGetObjectIdFromLoadOrRetrieveResult(value);
                    if (!CoreUtility.isNullOrUndefined(id)) {
                        var propNames = collectionPropertyPath.split('.');
                        var parent_1 = clientObject.context[propNames[0]];
                        for (var i = 1; i < propNames.length; i++) {
                            parent_1 = parent_1[propNames[i]];
                        }
                        this.saveOriginalObjectPathInfo();
                        this.resetForUpdateUsingObjectData();
                        this.m_parentObjectPath = parent_1._objectPath;
                        this.m_objectPathInfo.ParentObjectPathId = this.m_parentObjectPath.objectPathInfo.Id;
                        this.m_objectPathInfo.ObjectPathType = 5;
                        this.m_objectPathInfo.Name = '';
                        this.m_objectPathInfo.ArgumentInfo.Arguments = [id];
                        return;
                    }
                }
            }
            var parentIsCollection = this.parentObjectPath && this.parentObjectPath.isCollection;
            var getByIdMethodName = this.getByIdMethodName;
            if (parentIsCollection || !CoreUtility.isNullOrEmptyString(getByIdMethodName)) {
                var id = CommonUtility.tryGetObjectIdFromLoadOrRetrieveResult(value);
                if (!CoreUtility.isNullOrUndefined(id)) {
                    this.saveOriginalObjectPathInfo();
                    this.resetForUpdateUsingObjectData();
                    if (!CoreUtility.isNullOrEmptyString(getByIdMethodName)) {
                        this.m_objectPathInfo.ObjectPathType = 3;
                        this.m_objectPathInfo.Name = getByIdMethodName;
                    }
                    else {
                        this.m_objectPathInfo.ObjectPathType = 5;
                        this.m_objectPathInfo.Name = '';
                    }
                    this.m_objectPathInfo.ArgumentInfo.Arguments = [id];
                    return;
                }
            }
        };
        ObjectPath.prototype.resetForUpdateUsingObjectData = function () {
            this.m_isInvalidAfterRequest = false;
            this.m_isValid = true;
            this.m_operationType = 1;
            this.m_flags = 4;
            this.m_objectPathInfo.ArgumentInfo = {};
            this.m_argumentObjectPaths = null;
            this.m_getByIdMethodName = null;
        };
        ObjectPath.isRestorableObjectPath = function (objectPathType) {
            return (objectPathType === 1 ||
                objectPathType === 5 ||
                objectPathType === 3 ||
                objectPathType === 4);
        };
        ObjectPath.copyObjectPathInfo = function (src, dest) {
            dest.Id = src.Id;
            dest.ArgumentInfo = src.ArgumentInfo;
            dest.Name = src.Name;
            dest.ObjectPathType = src.ObjectPathType;
            dest.ParentObjectPathId = src.ParentObjectPathId;
        };
        return ObjectPath;
    }());
    OfficeExtension_1.ObjectPath = ObjectPath;
    var ClientRequestContextBase = (function () {
        function ClientRequestContextBase() {
            this.m_nextId = 0;
        }
        ClientRequestContextBase.prototype._nextId = function () {
            return ++this.m_nextId;
        };
        ClientRequestContextBase.prototype._addServiceApiAction = function (action, resultHandler, resolve, reject) {
            if (!this.m_serviceApiQueue) {
                this.m_serviceApiQueue = new ServiceApiQueue(this);
            }
            this.m_serviceApiQueue.add(action, resultHandler, resolve, reject);
        };
        ClientRequestContextBase._parseQueryOption = function (option) {
            var queryOption = {};
            if (typeof option === 'string') {
                var select = option;
                queryOption.Select = CommonUtility._parseSelectExpand(select);
            }
            else if (Array.isArray(option)) {
                queryOption.Select = option;
            }
            else if (typeof option === 'object') {
                var loadOption = option;
                if (ClientRequestContextBase.isLoadOption(loadOption)) {
                    if (typeof loadOption.select === 'string') {
                        queryOption.Select = CommonUtility._parseSelectExpand(loadOption.select);
                    }
                    else if (Array.isArray(loadOption.select)) {
                        queryOption.Select = loadOption.select;
                    }
                    else if (!CommonUtility.isNullOrUndefined(loadOption.select)) {
                        throw _Internal.RuntimeError._createInvalidArgError({ argumentName: 'option.select' });
                    }
                    if (typeof loadOption.expand === 'string') {
                        queryOption.Expand = CommonUtility._parseSelectExpand(loadOption.expand);
                    }
                    else if (Array.isArray(loadOption.expand)) {
                        queryOption.Expand = loadOption.expand;
                    }
                    else if (!CommonUtility.isNullOrUndefined(loadOption.expand)) {
                        throw _Internal.RuntimeError._createInvalidArgError({ argumentName: 'option.expand' });
                    }
                    if (typeof loadOption.top === 'number') {
                        queryOption.Top = loadOption.top;
                    }
                    else if (!CommonUtility.isNullOrUndefined(loadOption.top)) {
                        throw _Internal.RuntimeError._createInvalidArgError({ argumentName: 'option.top' });
                    }
                    if (typeof loadOption.skip === 'number') {
                        queryOption.Skip = loadOption.skip;
                    }
                    else if (!CommonUtility.isNullOrUndefined(loadOption.skip)) {
                        throw _Internal.RuntimeError._createInvalidArgError({ argumentName: 'option.skip' });
                    }
                }
                else {
                    queryOption = ClientRequestContextBase.parseStrictLoadOption(option);
                }
            }
            else if (!CommonUtility.isNullOrUndefined(option)) {
                throw _Internal.RuntimeError._createInvalidArgError({ argumentName: 'option' });
            }
            return queryOption;
        };
        ClientRequestContextBase.isLoadOption = function (loadOption) {
            if (!CommonUtility.isUndefined(loadOption.select) &&
                (typeof loadOption.select === 'string' || Array.isArray(loadOption.select)))
                return true;
            if (!CommonUtility.isUndefined(loadOption.expand) &&
                (typeof loadOption.expand === 'string' || Array.isArray(loadOption.expand)))
                return true;
            if (!CommonUtility.isUndefined(loadOption.top) && typeof loadOption.top === 'number')
                return true;
            if (!CommonUtility.isUndefined(loadOption.skip) && typeof loadOption.skip === 'number')
                return true;
            for (var i in loadOption) {
                return false;
            }
            return true;
        };
        ClientRequestContextBase.parseStrictLoadOption = function (option) {
            var ret = { Select: [] };
            ClientRequestContextBase.parseStrictLoadOptionHelper(ret, '', 'option', option);
            return ret;
        };
        ClientRequestContextBase.combineQueryPath = function (pathPrefix, key, separator) {
            if (pathPrefix.length === 0) {
                return key;
            }
            else {
                return pathPrefix + separator + key;
            }
        };
        ClientRequestContextBase.parseStrictLoadOptionHelper = function (queryInfo, pathPrefix, argPrefix, option) {
            for (var key in option) {
                var value = option[key];
                if (key === '$all') {
                    if (typeof value !== 'boolean') {
                        throw _Internal.RuntimeError._createInvalidArgError({
                            argumentName: ClientRequestContextBase.combineQueryPath(argPrefix, key, '.')
                        });
                    }
                    if (value) {
                        queryInfo.Select.push(ClientRequestContextBase.combineQueryPath(pathPrefix, '*', '/'));
                    }
                }
                else if (key === '$top') {
                    if (typeof value !== 'number' || pathPrefix.length > 0) {
                        throw _Internal.RuntimeError._createInvalidArgError({
                            argumentName: ClientRequestContextBase.combineQueryPath(argPrefix, key, '.')
                        });
                    }
                    queryInfo.Top = value;
                }
                else if (key === '$skip') {
                    if (typeof value !== 'number' || pathPrefix.length > 0) {
                        throw _Internal.RuntimeError._createInvalidArgError({
                            argumentName: ClientRequestContextBase.combineQueryPath(argPrefix, key, '.')
                        });
                    }
                    queryInfo.Skip = value;
                }
                else {
                    if (typeof value === 'boolean') {
                        if (value) {
                            queryInfo.Select.push(ClientRequestContextBase.combineQueryPath(pathPrefix, key, '/'));
                        }
                    }
                    else if (typeof value === 'object') {
                        ClientRequestContextBase.parseStrictLoadOptionHelper(queryInfo, ClientRequestContextBase.combineQueryPath(pathPrefix, key, '/'), ClientRequestContextBase.combineQueryPath(argPrefix, key, '.'), value);
                    }
                    else {
                        throw _Internal.RuntimeError._createInvalidArgError({
                            argumentName: ClientRequestContextBase.combineQueryPath(argPrefix, key, '.')
                        });
                    }
                }
            }
        };
        return ClientRequestContextBase;
    }());
    OfficeExtension_1.ClientRequestContextBase = ClientRequestContextBase;
    var InstantiateActionUpdateObjectPathHandler = (function () {
        function InstantiateActionUpdateObjectPathHandler(m_objectPath) {
            this.m_objectPath = m_objectPath;
        }
        InstantiateActionUpdateObjectPathHandler.prototype._handleResult = function (value) {
            if (CoreUtility.isNullOrUndefined(value)) {
                this.m_objectPath._updateAsNullObject();
            }
            else {
                this.m_objectPath.updateUsingObjectData(value, null);
            }
        };
        return InstantiateActionUpdateObjectPathHandler;
    }());
    var ClientRequestBase = (function () {
        function ClientRequestBase(context) {
            this.m_contextBase = context;
            this.m_actions = [];
            this.m_actionResultHandler = {};
            this.m_referencedObjectPaths = {};
            this.m_instantiatedObjectPaths = {};
            this.m_preSyncPromises = [];
        }
        ClientRequestBase.prototype.addAction = function (action) {
            this.m_actions.push(action);
            if (action.actionInfo.ActionType == 1) {
                this.m_instantiatedObjectPaths[action.actionInfo.ObjectPathId] = action;
            }
        };
        Object.defineProperty(ClientRequestBase.prototype, "hasActions", {
            get: function () {
                return this.m_actions.length > 0;
            },
            enumerable: true,
            configurable: true
        });
        ClientRequestBase.prototype._getLastAction = function () {
            return this.m_actions[this.m_actions.length - 1];
        };
        ClientRequestBase.prototype.ensureInstantiateObjectPath = function (objectPath) {
            if (objectPath) {
                if (this.m_instantiatedObjectPaths[objectPath.objectPathInfo.Id]) {
                    return;
                }
                this.ensureInstantiateObjectPath(objectPath.parentObjectPath);
                this.ensureInstantiateObjectPaths(objectPath.argumentObjectPaths);
                if (!this.m_instantiatedObjectPaths[objectPath.objectPathInfo.Id]) {
                    var actionInfo = {
                        Id: this.m_contextBase._nextId(),
                        ActionType: 1,
                        Name: '',
                        ObjectPathId: objectPath.objectPathInfo.Id
                    };
                    var instantiateAction = new Action(actionInfo, 1, 4);
                    instantiateAction.referencedObjectPath = objectPath;
                    this.addReferencedObjectPath(objectPath);
                    this.addAction(instantiateAction);
                    var resultHandler = new InstantiateActionUpdateObjectPathHandler(objectPath);
                    this.addActionResultHandler(instantiateAction, resultHandler);
                }
            }
        };
        ClientRequestBase.prototype.ensureInstantiateObjectPaths = function (objectPaths) {
            if (objectPaths) {
                for (var i = 0; i < objectPaths.length; i++) {
                    this.ensureInstantiateObjectPath(objectPaths[i]);
                }
            }
        };
        ClientRequestBase.prototype.addReferencedObjectPath = function (objectPath) {
            if (!objectPath || this.m_referencedObjectPaths[objectPath.objectPathInfo.Id]) {
                return;
            }
            if (!objectPath.isValid) {
                throw new _Internal.RuntimeError({
                    code: CoreErrorCodes.invalidObjectPath,
                    message: CoreUtility._getResourceString(CoreResourceStrings.invalidObjectPath, CommonUtility.getObjectPathExpression(objectPath)),
                    debugInfo: {
                        errorLocation: CommonUtility.getObjectPathExpression(objectPath)
                    }
                });
            }
            while (objectPath) {
                this.m_referencedObjectPaths[objectPath.objectPathInfo.Id] = objectPath;
                if (objectPath.objectPathInfo.ObjectPathType == 3) {
                    this.addReferencedObjectPaths(objectPath.argumentObjectPaths);
                }
                objectPath = objectPath.parentObjectPath;
            }
        };
        ClientRequestBase.prototype.addReferencedObjectPaths = function (objectPaths) {
            if (objectPaths) {
                for (var i = 0; i < objectPaths.length; i++) {
                    this.addReferencedObjectPath(objectPaths[i]);
                }
            }
        };
        ClientRequestBase.prototype.addActionResultHandler = function (action, resultHandler) {
            this.m_actionResultHandler[action.actionInfo.Id] = resultHandler;
        };
        ClientRequestBase.prototype.aggregrateRequestFlags = function (requestFlags, operationType, flags) {
            if (operationType === 0) {
                requestFlags = requestFlags | 1;
                if ((flags & 2) === 0) {
                    requestFlags = requestFlags & ~16;
                }
                if ((flags & 8) === 0) {
                    requestFlags = requestFlags & ~256;
                }
                requestFlags = requestFlags & ~4;
            }
            if (flags & 1) {
                requestFlags = requestFlags | 2;
            }
            if ((flags & 4) === 0) {
                requestFlags = requestFlags & ~4;
            }
            return requestFlags;
        };
        ClientRequestBase.prototype.finallyNormalizeFlags = function (requestFlags) {
            if ((requestFlags & 1) === 0) {
                requestFlags = requestFlags & ~16;
                requestFlags = requestFlags & ~256;
            }
            if (!OfficeExtension_1._internalConfig.enableConcurrentFlag) {
                requestFlags = requestFlags & ~4;
            }
            if (!OfficeExtension_1._internalConfig.enableUndoableFlag) {
                requestFlags = requestFlags & ~16;
            }
            if (!CommonUtility.isSetSupported('RichApiRuntimeFlag', '1.1')) {
                requestFlags = requestFlags & ~4;
                requestFlags = requestFlags & ~16;
            }
            if (!CommonUtility.isSetSupported('RichApiRuntimeFlag', '1.2')) {
                requestFlags = requestFlags & ~256;
            }
            if (typeof this.m_flagsForTesting === 'number') {
                requestFlags = this.m_flagsForTesting;
            }
            return requestFlags;
        };
        ClientRequestBase.prototype.buildRequestMessageBodyAndRequestFlags = function () {
            if (OfficeExtension_1._internalConfig.enableEarlyDispose) {
                ClientRequestBase._calculateLastUsedObjectPathIds(this.m_actions);
            }
            var requestFlags = 4 |
                16 |
                256;
            var objectPaths = {};
            for (var i in this.m_referencedObjectPaths) {
                requestFlags = this.aggregrateRequestFlags(requestFlags, this.m_referencedObjectPaths[i].operationType, this.m_referencedObjectPaths[i].flags);
                objectPaths[i] = this.m_referencedObjectPaths[i].objectPathInfo;
            }
            var actions = [];
            var hasKeepReference = false;
            for (var index = 0; index < this.m_actions.length; index++) {
                var action = this.m_actions[index];
                if (action.actionInfo.ActionType === 3 &&
                    action.actionInfo.Name === CommonConstants.keepReference) {
                    hasKeepReference = true;
                }
                requestFlags = this.aggregrateRequestFlags(requestFlags, action.operationType, action.flags);
                actions.push(action.actionInfo);
            }
            requestFlags = this.finallyNormalizeFlags(requestFlags);
            var body = {
                AutoKeepReference: this.m_contextBase._autoCleanup && hasKeepReference,
                Actions: actions,
                ObjectPaths: objectPaths
            };
            return {
                body: body,
                flags: requestFlags
            };
        };
        ClientRequestBase.prototype.processResponse = function (actionResults) {
            if (actionResults) {
                for (var i = 0; i < actionResults.length; i++) {
                    var actionResult = actionResults[i];
                    var handler = this.m_actionResultHandler[actionResult.ActionId];
                    if (handler) {
                        handler._handleResult(actionResult.Value);
                    }
                }
            }
        };
        ClientRequestBase.prototype.invalidatePendingInvalidObjectPaths = function () {
            for (var i in this.m_referencedObjectPaths) {
                if (this.m_referencedObjectPaths[i].isInvalidAfterRequest) {
                    this.m_referencedObjectPaths[i].isValid = false;
                }
            }
        };
        ClientRequestBase.prototype._addPreSyncPromise = function (value) {
            this.m_preSyncPromises.push(value);
        };
        Object.defineProperty(ClientRequestBase.prototype, "_preSyncPromises", {
            get: function () {
                return this.m_preSyncPromises;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(ClientRequestBase.prototype, "_actions", {
            get: function () {
                return this.m_actions;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(ClientRequestBase.prototype, "_objectPaths", {
            get: function () {
                return this.m_referencedObjectPaths;
            },
            enumerable: true,
            configurable: true
        });
        ClientRequestBase.prototype._removeKeepReferenceAction = function (objectPathId) {
            for (var i = this.m_actions.length - 1; i >= 0; i--) {
                var actionInfo = this.m_actions[i].actionInfo;
                if (actionInfo.ObjectPathId === objectPathId &&
                    actionInfo.ActionType === 3 &&
                    actionInfo.Name === CommonConstants.keepReference) {
                    this.m_actions.splice(i, 1);
                    break;
                }
            }
        };
        ClientRequestBase._updateLastUsedActionIdOfObjectPathId = function (lastUsedActionIdOfObjectPathId, objectPath, actionId) {
            while (objectPath) {
                if (lastUsedActionIdOfObjectPathId[objectPath.objectPathInfo.Id]) {
                    return;
                }
                lastUsedActionIdOfObjectPathId[objectPath.objectPathInfo.Id] = actionId;
                var argumentObjectPaths = objectPath.argumentObjectPaths;
                if (argumentObjectPaths) {
                    var argumentObjectPathsLength = argumentObjectPaths.length;
                    for (var i = 0; i < argumentObjectPathsLength; i++) {
                        ClientRequestBase._updateLastUsedActionIdOfObjectPathId(lastUsedActionIdOfObjectPathId, argumentObjectPaths[i], actionId);
                    }
                }
                objectPath = objectPath.parentObjectPath;
            }
        };
        ClientRequestBase._calculateLastUsedObjectPathIds = function (actions) {
            var lastUsedActionIdOfObjectPathId = {};
            var actionsLength = actions.length;
            for (var index = actionsLength - 1; index >= 0; --index) {
                var action = actions[index];
                var actionId = action.actionInfo.Id;
                if (action.referencedObjectPath) {
                    ClientRequestBase._updateLastUsedActionIdOfObjectPathId(lastUsedActionIdOfObjectPathId, action.referencedObjectPath, actionId);
                }
                var referencedObjectPaths = action.referencedArgumentObjectPaths;
                if (referencedObjectPaths) {
                    var referencedObjectPathsLength = referencedObjectPaths.length;
                    for (var refIndex = 0; refIndex < referencedObjectPathsLength; refIndex++) {
                        ClientRequestBase._updateLastUsedActionIdOfObjectPathId(lastUsedActionIdOfObjectPathId, referencedObjectPaths[refIndex], actionId);
                    }
                }
            }
            var lastUsedObjectPathIdsOfAction = {};
            for (var key in lastUsedActionIdOfObjectPathId) {
                var actionId = lastUsedActionIdOfObjectPathId[key];
                var objectPathIds = lastUsedObjectPathIdsOfAction[actionId];
                if (!objectPathIds) {
                    objectPathIds = [];
                    lastUsedObjectPathIdsOfAction[actionId] = objectPathIds;
                }
                objectPathIds.push(parseInt(key));
            }
            for (var index = 0; index < actionsLength; index++) {
                var action = actions[index];
                var lastUsedObjectPathIds = lastUsedObjectPathIdsOfAction[action.actionInfo.Id];
                if (lastUsedObjectPathIds && lastUsedObjectPathIds.length > 0) {
                    action.actionInfo.L = lastUsedObjectPathIds;
                }
                else if (action.actionInfo.L) {
                    delete action.actionInfo.L;
                }
            }
        };
        return ClientRequestBase;
    }());
    OfficeExtension_1.ClientRequestBase = ClientRequestBase;
    var ClientResult = (function () {
        function ClientResult(m_type) {
            this.m_type = m_type;
        }
        Object.defineProperty(ClientResult.prototype, "value", {
            get: function () {
                if (!this.m_isLoaded) {
                    throw new _Internal.RuntimeError({
                        code: CoreErrorCodes.valueNotLoaded,
                        message: CoreUtility._getResourceString(CoreResourceStrings.valueNotLoaded),
                        debugInfo: {
                            errorLocation: 'clientResult.value'
                        }
                    });
                }
                return this.m_value;
            },
            enumerable: true,
            configurable: true
        });
        ClientResult.prototype._handleResult = function (value) {
            this.m_isLoaded = true;
            if (typeof value === 'object' && value && value._IsNull) {
                return;
            }
            if (this.m_type === 1) {
                this.m_value = CommonUtility.adjustToDateTime(value);
            }
            else {
                this.m_value = value;
            }
        };
        return ClientResult;
    }());
    OfficeExtension_1.ClientResult = ClientResult;
    var ServiceApiQueue = (function () {
        function ServiceApiQueue(m_context) {
            this.m_context = m_context;
            this.m_actions = [];
        }
        ServiceApiQueue.prototype.add = function (action, resultHandler, resolve, reject) {
            var _this = this;
            this.m_actions.push({ action: action, resultHandler: resultHandler, resolve: resolve, reject: reject });
            if (this.m_actions.length === 1) {
                setTimeout(function () { return _this.processActions(); }, 0);
            }
        };
        ServiceApiQueue.prototype.processActions = function () {
            var _this = this;
            if (this.m_actions.length === 0) {
                return;
            }
            var actions = this.m_actions;
            this.m_actions = [];
            var request = new ClientRequestBase(this.m_context);
            for (var i = 0; i < actions.length; i++) {
                var action = actions[i];
                request.ensureInstantiateObjectPath(action.action.referencedObjectPath);
                request.ensureInstantiateObjectPaths(action.action.referencedArgumentObjectPaths);
                request.addAction(action.action);
                request.addReferencedObjectPath(action.action.referencedObjectPath);
                request.addReferencedObjectPaths(action.action.referencedArgumentObjectPaths);
            }
            var _a = request.buildRequestMessageBodyAndRequestFlags(), body = _a.body, flags = _a.flags;
            var requestMessage = {
                Url: CoreConstants.localDocumentApiPrefix,
                Headers: null,
                Body: body
            };
            CoreUtility.log('Request:');
            CoreUtility.log(JSON.stringify(body));
            var executor = new HttpRequestExecutor();
            executor
                .executeAsync(this.m_context._customData, flags, requestMessage)
                .then(function (response) {
                _this.processResponse(request, actions, response);
            })
                .catch(function (ex) {
                for (var i = 0; i < actions.length; i++) {
                    var action = actions[i];
                    action.reject(ex);
                }
            });
        };
        ServiceApiQueue.prototype.processResponse = function (request, actions, response) {
            var error = this.getErrorFromResponse(response);
            var actionResults = null;
            if (response.Body.Results) {
                actionResults = response.Body.Results;
            }
            else if (response.Body.ProcessedResults && response.Body.ProcessedResults.Results) {
                actionResults = response.Body.ProcessedResults.Results;
            }
            if (!actionResults) {
                actionResults = [];
            }
            this.processActionResults(request, actions, actionResults, error);
        };
        ServiceApiQueue.prototype.getErrorFromResponse = function (response) {
            if (!CoreUtility.isNullOrEmptyString(response.ErrorCode)) {
                return new _Internal.RuntimeError({
                    code: response.ErrorCode,
                    message: response.ErrorMessage
                });
            }
            if (response.Body && response.Body.Error) {
                return new _Internal.RuntimeError({
                    code: response.Body.Error.Code,
                    message: response.Body.Error.Message
                });
            }
            return null;
        };
        ServiceApiQueue.prototype.processActionResults = function (request, actions, actionResults, err) {
            request.processResponse(actionResults);
            for (var i = 0; i < actions.length; i++) {
                var action = actions[i];
                var actionId = action.action.actionInfo.Id;
                var hasResult = false;
                for (var j = 0; j < actionResults.length; j++) {
                    if (actionId == actionResults[j].ActionId) {
                        var resultValue = actionResults[j].Value;
                        if (action.resultHandler) {
                            action.resultHandler._handleResult(resultValue);
                            resultValue = action.resultHandler.value;
                        }
                        if (action.resolve) {
                            action.resolve(resultValue);
                        }
                        hasResult = true;
                        break;
                    }
                }
                if (!hasResult && action.reject) {
                    if (err) {
                        action.reject(err);
                    }
                    else {
                        action.reject('No response for the action.');
                    }
                }
            }
        };
        return ServiceApiQueue;
    }());
    var HttpRequestExecutor = (function () {
        function HttpRequestExecutor() {
        }
        HttpRequestExecutor.prototype.getRequestUrl = function (baseUrl, requestFlags) {
            if (baseUrl.charAt(baseUrl.length - 1) != '/') {
                baseUrl = baseUrl + '/';
            }
            baseUrl = baseUrl + CoreConstants.processQuery;
            baseUrl = baseUrl + '?' + CoreConstants.flags + '=' + requestFlags.toString();
            return baseUrl;
        };
        HttpRequestExecutor.prototype.executeAsync = function (customData, requestFlags, requestMessage) {
            var url = this.getRequestUrl(requestMessage.Url, requestFlags);
            var requestInfo = {
                method: 'POST',
                url: url,
                headers: {},
                body: requestMessage.Body
            };
            requestInfo.headers[CoreConstants.sourceLibHeader] = HttpRequestExecutor.SourceLibHeaderValue;
            requestInfo.headers['CONTENT-TYPE'] = 'application/json';
            if (requestMessage.Headers) {
                for (var key in requestMessage.Headers) {
                    requestInfo.headers[key] = requestMessage.Headers[key];
                }
            }
            var sendRequestFunc = CoreUtility._isLocalDocumentUrl(requestInfo.url)
                ? HttpUtility.sendLocalDocumentRequest
                : HttpUtility.sendRequest;
            return sendRequestFunc(requestInfo).then(function (responseInfo) {
                var response;
                if (responseInfo.statusCode === 200) {
                    response = {
                        ErrorCode: null,
                        ErrorMessage: null,
                        Headers: responseInfo.headers,
                        Body: CoreUtility._parseResponseBody(responseInfo)
                    };
                }
                else {
                    CoreUtility.log('Error Response:' + responseInfo.body);
                    var error = CoreUtility._parseErrorResponse(responseInfo);
                    response = {
                        ErrorCode: error.errorCode,
                        ErrorMessage: error.errorMessage,
                        Headers: responseInfo.headers,
                        Body: null
                    };
                }
                return response;
            });
        };
        HttpRequestExecutor.SourceLibHeaderValue = 'officejs-rest';
        return HttpRequestExecutor;
    }());
    OfficeExtension_1.HttpRequestExecutor = HttpRequestExecutor;
    var CommonConstants = (function (_super) {
        __extends(CommonConstants, _super);
        function CommonConstants() {
            return _super !== null && _super.apply(this, arguments) || this;
        }
        CommonConstants.collectionPropertyPath = '_collectionPropertyPath';
        CommonConstants.id = 'Id';
        CommonConstants.idLowerCase = 'id';
        CommonConstants.idPrivate = '_Id';
        CommonConstants.keepReference = '_KeepReference';
        CommonConstants.objectPathIdPrivate = '_ObjectPathId';
        CommonConstants.referenceId = '_ReferenceId';
        CommonConstants.items = '_Items';
        CommonConstants.itemsLowerCase = 'items';
        CommonConstants.scalarPropertyNames = '_scalarPropertyNames';
        CommonConstants.scalarPropertyOriginalNames = '_scalarPropertyOriginalNames';
        CommonConstants.navigationPropertyNames = '_navigationPropertyNames';
        CommonConstants.scalarPropertyUpdateable = '_scalarPropertyUpdateable';
        return CommonConstants;
    }(CoreConstants));
    OfficeExtension_1.CommonConstants = CommonConstants;
    var CommonUtility = (function (_super) {
        __extends(CommonUtility, _super);
        function CommonUtility() {
            return _super !== null && _super.apply(this, arguments) || this;
        }
        CommonUtility.validateObjectPath = function (clientObject) {
            var objectPath = clientObject._objectPath;
            while (objectPath) {
                if (!objectPath.isValid) {
                    throw new _Internal.RuntimeError({
                        code: CoreErrorCodes.invalidObjectPath,
                        message: CoreUtility._getResourceString(CoreResourceStrings.invalidObjectPath, CommonUtility.getObjectPathExpression(objectPath)),
                        debugInfo: {
                            errorLocation: CommonUtility.getObjectPathExpression(objectPath)
                        }
                    });
                }
                objectPath = objectPath.parentObjectPath;
            }
        };
        CommonUtility.validateReferencedObjectPaths = function (objectPaths) {
            if (objectPaths) {
                for (var i = 0; i < objectPaths.length; i++) {
                    var objectPath = objectPaths[i];
                    while (objectPath) {
                        if (!objectPath.isValid) {
                            throw new _Internal.RuntimeError({
                                code: CoreErrorCodes.invalidObjectPath,
                                message: CoreUtility._getResourceString(CoreResourceStrings.invalidObjectPath, CommonUtility.getObjectPathExpression(objectPath))
                            });
                        }
                        objectPath = objectPath.parentObjectPath;
                    }
                }
            }
        };
        CommonUtility._toCamelLowerCase = function (name) {
            if (CoreUtility.isNullOrEmptyString(name)) {
                return name;
            }
            var index = 0;
            while (index < name.length && name.charCodeAt(index) >= 65 && name.charCodeAt(index) <= 90) {
                index++;
            }
            if (index < name.length) {
                return name.substr(0, index).toLowerCase() + name.substr(index);
            }
            else {
                return name.toLowerCase();
            }
        };
        CommonUtility.adjustToDateTime = function (value) {
            if (CoreUtility.isNullOrUndefined(value)) {
                return null;
            }
            if (typeof value === 'string') {
                return new Date(value);
            }
            if (Array.isArray(value)) {
                var arr = value;
                for (var i = 0; i < arr.length; i++) {
                    arr[i] = CommonUtility.adjustToDateTime(arr[i]);
                }
                return arr;
            }
            throw CoreUtility._createInvalidArgError({ argumentName: 'date' });
        };
        CommonUtility.tryGetObjectIdFromLoadOrRetrieveResult = function (value) {
            var id = value[CommonConstants.id];
            if (CoreUtility.isNullOrUndefined(id)) {
                id = value[CommonConstants.idLowerCase];
            }
            if (CoreUtility.isNullOrUndefined(id)) {
                id = value[CommonConstants.idPrivate];
            }
            return id;
        };
        CommonUtility.getObjectPathExpression = function (objectPath) {
            var ret = '';
            while (objectPath) {
                switch (objectPath.objectPathInfo.ObjectPathType) {
                    case 1:
                        ret = ret;
                        break;
                    case 2:
                        ret = 'new()' + (ret.length > 0 ? '.' : '') + ret;
                        break;
                    case 3:
                        ret = CommonUtility.normalizeName(objectPath.objectPathInfo.Name) + '()' + (ret.length > 0 ? '.' : '') + ret;
                        break;
                    case 4:
                        ret = CommonUtility.normalizeName(objectPath.objectPathInfo.Name) + (ret.length > 0 ? '.' : '') + ret;
                        break;
                    case 5:
                        ret = 'getItem()' + (ret.length > 0 ? '.' : '') + ret;
                        break;
                    case 6:
                        ret = '_reference()' + (ret.length > 0 ? '.' : '') + ret;
                        break;
                }
                objectPath = objectPath.parentObjectPath;
            }
            return ret;
        };
        CommonUtility.setMethodArguments = function (context, argumentInfo, args) {
            if (CoreUtility.isNullOrUndefined(args)) {
                return null;
            }
            var referencedObjectPaths = new Array();
            var referencedObjectPathIds = new Array();
            var hasOne = CommonUtility.collectObjectPathInfos(context, args, referencedObjectPaths, referencedObjectPathIds);
            argumentInfo.Arguments = args;
            if (hasOne) {
                argumentInfo.ReferencedObjectPathIds = referencedObjectPathIds;
            }
            return referencedObjectPaths;
        };
        CommonUtility.validateContext = function (context, obj) {
            if (context && obj && obj._context !== context) {
                throw new _Internal.RuntimeError({
                    code: CoreErrorCodes.invalidRequestContext,
                    message: CoreUtility._getResourceString(CoreResourceStrings.invalidRequestContext)
                });
            }
        };
        CommonUtility.isSetSupported = function (apiSetName, apiSetVersion) {
            if (typeof window !== 'undefined' &&
                window.Office &&
                window.Office.context &&
                window.Office.context.requirements) {
                return window.Office.context.requirements.isSetSupported(apiSetName, apiSetVersion);
            }
            return true;
        };
        CommonUtility.throwIfApiNotSupported = function (apiFullName, apiSetName, apiSetVersion, hostName) {
            if (!CommonUtility._doApiNotSupportedCheck) {
                return;
            }
            if (!CommonUtility.isSetSupported(apiSetName, apiSetVersion)) {
                var message = CoreUtility._getResourceString(CoreResourceStrings.apiNotFoundDetails, [
                    apiFullName,
                    apiSetName + ' ' + apiSetVersion,
                    hostName
                ]);
                throw new _Internal.RuntimeError({
                    code: CoreErrorCodes.apiNotFound,
                    message: message,
                    debugInfo: { errorLocation: apiFullName }
                });
            }
        };
        CommonUtility._parseSelectExpand = function (select) {
            var args = [];
            if (!CoreUtility.isNullOrEmptyString(select)) {
                var propertyNames = select.split(',');
                for (var i = 0; i < propertyNames.length; i++) {
                    var propertyName = propertyNames[i];
                    propertyName = sanitizeForAnyItemsSlash(propertyName.trim());
                    if (propertyName.length > 0) {
                        args.push(propertyName);
                    }
                }
            }
            return args;
            function sanitizeForAnyItemsSlash(propertyName) {
                var propertyNameLower = propertyName.toLowerCase();
                if (propertyNameLower === 'items' || propertyNameLower === 'items/') {
                    return '*';
                }
                var itemsSlashLength = 6;
                var isItemsSlashOrItemsDot = propertyNameLower.substr(0, itemsSlashLength) === 'items/' ||
                    propertyNameLower.substr(0, itemsSlashLength) === 'items.';
                if (isItemsSlashOrItemsDot) {
                    propertyName = propertyName.substr(itemsSlashLength);
                }
                return propertyName.replace(new RegExp('[/.]items[/.]', 'gi'), '/');
            }
        };
        CommonUtility.changePropertyNameToCamelLowerCase = function (value) {
            var charCodeUnderscore = 95;
            if (Array.isArray(value)) {
                var ret = [];
                for (var i = 0; i < value.length; i++) {
                    ret.push(this.changePropertyNameToCamelLowerCase(value[i]));
                }
                return ret;
            }
            else if (typeof value === 'object' && value !== null) {
                var ret = {};
                for (var key in value) {
                    var propValue = value[key];
                    if (key === CommonConstants.items) {
                        ret = {};
                        ret[CommonConstants.itemsLowerCase] = this.changePropertyNameToCamelLowerCase(propValue);
                        break;
                    }
                    else {
                        var propName = CommonUtility._toCamelLowerCase(key);
                        ret[propName] = this.changePropertyNameToCamelLowerCase(propValue);
                    }
                }
                return ret;
            }
            else {
                return value;
            }
        };
        CommonUtility.purifyJson = function (value) {
            var charCodeUnderscore = 95;
            if (Array.isArray(value)) {
                var ret = [];
                for (var i = 0; i < value.length; i++) {
                    ret.push(this.purifyJson(value[i]));
                }
                return ret;
            }
            else if (typeof value === 'object' && value !== null) {
                var ret = {};
                for (var key in value) {
                    if (key.charCodeAt(0) !== charCodeUnderscore) {
                        var propValue = value[key];
                        if (typeof propValue === 'object' && propValue !== null && Array.isArray(propValue['items'])) {
                            propValue = propValue['items'];
                        }
                        ret[key] = this.purifyJson(propValue);
                    }
                }
                return ret;
            }
            else {
                return value;
            }
        };
        CommonUtility.collectObjectPathInfos = function (context, args, referencedObjectPaths, referencedObjectPathIds) {
            var hasOne = false;
            for (var i = 0; i < args.length; i++) {
                if (args[i] instanceof ClientObjectBase) {
                    var clientObject = args[i];
                    CommonUtility.validateContext(context, clientObject);
                    args[i] = clientObject._objectPath.objectPathInfo.Id;
                    referencedObjectPathIds.push(clientObject._objectPath.objectPathInfo.Id);
                    referencedObjectPaths.push(clientObject._objectPath);
                    hasOne = true;
                }
                else if (Array.isArray(args[i])) {
                    var childArrayObjectPathIds = new Array();
                    var childArrayHasOne = CommonUtility.collectObjectPathInfos(context, args[i], referencedObjectPaths, childArrayObjectPathIds);
                    if (childArrayHasOne) {
                        referencedObjectPathIds.push(childArrayObjectPathIds);
                        hasOne = true;
                    }
                    else {
                        referencedObjectPathIds.push(0);
                    }
                }
                else if (CoreUtility.isPlainJsonObject(args[i])) {
                    referencedObjectPathIds.push(0);
                    CommonUtility.replaceClientObjectPropertiesWithObjectPathIds(args[i], referencedObjectPaths);
                }
                else {
                    referencedObjectPathIds.push(0);
                }
            }
            return hasOne;
        };
        CommonUtility.replaceClientObjectPropertiesWithObjectPathIds = function (value, referencedObjectPaths) {
            var _a, _b;
            for (var key in value) {
                var propValue = value[key];
                if (propValue instanceof ClientObjectBase) {
                    referencedObjectPaths.push(propValue._objectPath);
                    value[key] = (_a = {}, _a[CommonConstants.objectPathIdPrivate] = propValue._objectPath.objectPathInfo.Id, _a);
                }
                else if (Array.isArray(propValue)) {
                    for (var i = 0; i < propValue.length; i++) {
                        if (propValue[i] instanceof ClientObjectBase) {
                            var elem = propValue[i];
                            referencedObjectPaths.push(elem._objectPath);
                            propValue[i] = (_b = {}, _b[CommonConstants.objectPathIdPrivate] = elem._objectPath.objectPathInfo.Id, _b);
                        }
                        else if (CoreUtility.isPlainJsonObject(propValue[i])) {
                            CommonUtility.replaceClientObjectPropertiesWithObjectPathIds(propValue[i], referencedObjectPaths);
                        }
                    }
                }
                else if (CoreUtility.isPlainJsonObject(propValue)) {
                    CommonUtility.replaceClientObjectPropertiesWithObjectPathIds(propValue, referencedObjectPaths);
                }
                else {
                }
            }
        };
        CommonUtility.normalizeName = function (name) {
            return name.substr(0, 1).toLowerCase() + name.substr(1);
        };
        CommonUtility._doApiNotSupportedCheck = false;
        return CommonUtility;
    }(CoreUtility));
    OfficeExtension_1.CommonUtility = CommonUtility;
    var CommonResourceStrings = (function (_super) {
        __extends(CommonResourceStrings, _super);
        function CommonResourceStrings() {
            return _super !== null && _super.apply(this, arguments) || this;
        }
        CommonResourceStrings.propertyDoesNotExist = 'PropertyDoesNotExist';
        CommonResourceStrings.attemptingToSetReadOnlyProperty = 'AttemptingToSetReadOnlyProperty';
        return CommonResourceStrings;
    }(CoreResourceStrings));
    OfficeExtension_1.CommonResourceStrings = CommonResourceStrings;
    var ErrorCodes = (function (_super) {
        __extends(ErrorCodes, _super);
        function ErrorCodes() {
            return _super !== null && _super.apply(this, arguments) || this;
        }
        ErrorCodes.propertyNotLoaded = 'PropertyNotLoaded';
        ErrorCodes.runMustReturnPromise = 'RunMustReturnPromise';
        ErrorCodes.cannotRegisterEvent = 'CannotRegisterEvent';
        ErrorCodes.invalidOrTimedOutSession = 'InvalidOrTimedOutSession';
        ErrorCodes.cannotUpdateReadOnlyProperty = 'CannotUpdateReadOnlyProperty';
        return ErrorCodes;
    }(CoreErrorCodes));
    OfficeExtension_1.ErrorCodes = ErrorCodes;
    var TraceMarkerActionResultHandler = (function () {
        function TraceMarkerActionResultHandler(callback) {
            this.m_callback = callback;
        }
        TraceMarkerActionResultHandler.prototype._handleResult = function (value) {
            if (this.m_callback) {
                this.m_callback();
            }
        };
        return TraceMarkerActionResultHandler;
    }());
    var ActionFactory = (function (_super) {
        __extends(ActionFactory, _super);
        function ActionFactory() {
            return _super !== null && _super.apply(this, arguments) || this;
        }
        ActionFactory.createMethodAction = function (context, parent, methodName, operationType, args, flags) {
            Utility.validateObjectPath(parent);
            var actionInfo = {
                Id: context._nextId(),
                ActionType: 3,
                Name: methodName,
                ObjectPathId: parent._objectPath.objectPathInfo.Id,
                ArgumentInfo: {}
            };
            var referencedArgumentObjectPaths = Utility.setMethodArguments(context, actionInfo.ArgumentInfo, args);
            Utility.validateReferencedObjectPaths(referencedArgumentObjectPaths);
            var action = new Action(actionInfo, operationType, Utility._fixupApiFlags(flags));
            action.referencedObjectPath = parent._objectPath;
            action.referencedArgumentObjectPaths = referencedArgumentObjectPaths;
            parent._addAction(action);
            return action;
        };
        ActionFactory.createRecursiveQueryAction = function (context, parent, query) {
            Utility.validateObjectPath(parent);
            var actionInfo = {
                Id: context._nextId(),
                ActionType: 6,
                Name: '',
                ObjectPathId: parent._objectPath.objectPathInfo.Id,
                RecursiveQueryInfo: query
            };
            var action = new Action(actionInfo, 1, 4);
            action.referencedObjectPath = parent._objectPath;
            parent._addAction(action);
            return action;
        };
        ActionFactory.createEnsureUnchangedAction = function (context, parent, objectState) {
            Utility.validateObjectPath(parent);
            var actionInfo = {
                Id: context._nextId(),
                ActionType: 8,
                Name: '',
                ObjectPathId: parent._objectPath.objectPathInfo.Id,
                ObjectState: objectState
            };
            var action = new Action(actionInfo, 1, 4);
            action.referencedObjectPath = parent._objectPath;
            parent._addAction(action);
            return action;
        };
        ActionFactory.createInstantiateAction = function (context, obj) {
            Utility.validateObjectPath(obj);
            context._pendingRequest.ensureInstantiateObjectPath(obj._objectPath.parentObjectPath);
            context._pendingRequest.ensureInstantiateObjectPaths(obj._objectPath.argumentObjectPaths);
            var actionInfo = {
                Id: context._nextId(),
                ActionType: 1,
                Name: '',
                ObjectPathId: obj._objectPath.objectPathInfo.Id
            };
            var action = new Action(actionInfo, 1, 4);
            action.referencedObjectPath = obj._objectPath;
            obj._addAction(action, new InstantiateActionResultHandler(obj), true);
            return action;
        };
        ActionFactory.createTraceAction = function (context, message, addTraceMessage) {
            var actionInfo = {
                Id: context._nextId(),
                ActionType: 5,
                Name: 'Trace',
                ObjectPathId: 0
            };
            var ret = new Action(actionInfo, 1, 4);
            context._pendingRequest.addAction(ret);
            if (addTraceMessage) {
                context._pendingRequest.addTrace(actionInfo.Id, message);
            }
            return ret;
        };
        ActionFactory.createTraceMarkerForCallback = function (context, callback) {
            var action = ActionFactory.createTraceAction(context, null, false);
            context._pendingRequest.addActionResultHandler(action, new TraceMarkerActionResultHandler(callback));
        };
        return ActionFactory;
    }(CommonActionFactory));
    OfficeExtension_1.ActionFactory = ActionFactory;
    var ClientObject = (function (_super) {
        __extends(ClientObject, _super);
        function ClientObject(context, objectPath) {
            var _this = _super.call(this, context, objectPath) || this;
            Utility.checkArgumentNull(context, 'context');
            _this.m_context = context;
            if (_this._objectPath) {
                if (!context._processingResult && context._pendingRequest) {
                    ActionFactory.createInstantiateAction(context, _this);
                    if (context._autoCleanup && _this._KeepReference) {
                        context.trackedObjects._autoAdd(_this);
                    }
                }
            }
            return _this;
        }
        Object.defineProperty(ClientObject.prototype, "context", {
            get: function () {
                return this.m_context;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(ClientObject.prototype, "isNull", {
            get: function () {
                if (typeof (this.m_isNull) === 'undefined' && TestUtility.isMock()) {
                    return false;
                }
                Utility.throwIfNotLoaded('isNull', this._isNull, null, this._isNull);
                return this._isNull;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(ClientObject.prototype, "isNullObject", {
            get: function () {
                if (typeof (this.m_isNull) === 'undefined' && TestUtility.isMock()) {
                    return false;
                }
                Utility.throwIfNotLoaded('isNullObject', this._isNull, null, this._isNull);
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
                this.m_isNull = value;
                if (value && this._objectPath) {
                    this._objectPath._updateAsNullObject();
                }
            },
            enumerable: true,
            configurable: true
        });
        ClientObject.prototype._addAction = function (action, resultHandler, isInstantiationEnsured) {
            if (resultHandler === void 0) { resultHandler = null; }
            if (!isInstantiationEnsured) {
                this.context._pendingRequest.ensureInstantiateObjectPath(this._objectPath);
                this.context._pendingRequest.ensureInstantiateObjectPaths(action.referencedArgumentObjectPaths);
            }
            this.context._pendingRequest.addAction(action);
            this.context._pendingRequest.addReferencedObjectPath(this._objectPath);
            this.context._pendingRequest.addReferencedObjectPaths(action.referencedArgumentObjectPaths);
            this.context._pendingRequest.addActionResultHandler(action, resultHandler);
            return CoreUtility._createPromiseFromResult(null);
        };
        ClientObject.prototype._handleResult = function (value) {
            this._isNull = Utility.isNullOrUndefined(value);
            this.context.trackedObjects._autoTrackIfNecessaryWhenHandleObjectResultValue(this, value);
        };
        ClientObject.prototype._handleIdResult = function (value) {
            this._isNull = Utility.isNullOrUndefined(value);
            Utility.fixObjectPathIfNecessary(this, value);
            this.context.trackedObjects._autoTrackIfNecessaryWhenHandleObjectResultValue(this, value);
        };
        ClientObject.prototype._handleRetrieveResult = function (value, result) {
            this._handleIdResult(value);
        };
        ClientObject.prototype._recursivelySet = function (input, options, scalarWriteablePropertyNames, objectPropertyNames, notAllowedToBeSetPropertyNames) {
            var isClientObject = input instanceof ClientObject;
            var originalInput = input;
            if (isClientObject) {
                if (Object.getPrototypeOf(this) === Object.getPrototypeOf(input)) {
                    input = JSON.parse(JSON.stringify(input));
                }
                else {
                    throw _Internal.RuntimeError._createInvalidArgError({
                        argumentName: 'properties',
                        errorLocation: this._className + '.set'
                    });
                }
            }
            try {
                var prop;
                for (var i = 0; i < scalarWriteablePropertyNames.length; i++) {
                    prop = scalarWriteablePropertyNames[i];
                    if (input.hasOwnProperty(prop)) {
                        if (typeof input[prop] !== 'undefined') {
                            this[prop] = input[prop];
                        }
                    }
                }
                for (var i = 0; i < objectPropertyNames.length; i++) {
                    prop = objectPropertyNames[i];
                    if (input.hasOwnProperty(prop)) {
                        if (typeof input[prop] !== 'undefined') {
                            var dataToPassToSet = isClientObject ? originalInput[prop] : input[prop];
                            this[prop].set(dataToPassToSet, options);
                        }
                    }
                }
                var throwOnReadOnly = !isClientObject;
                if (options && !Utility.isNullOrUndefined(throwOnReadOnly)) {
                    throwOnReadOnly = options.throwOnReadOnly;
                }
                for (var i = 0; i < notAllowedToBeSetPropertyNames.length; i++) {
                    prop = notAllowedToBeSetPropertyNames[i];
                    if (input.hasOwnProperty(prop)) {
                        if (typeof input[prop] !== 'undefined' && throwOnReadOnly) {
                            throw new _Internal.RuntimeError({
                                code: CoreErrorCodes.invalidArgument,
                                message: CoreUtility._getResourceString(ResourceStrings.cannotApplyPropertyThroughSetMethod, prop),
                                debugInfo: {
                                    errorLocation: prop
                                }
                            });
                        }
                    }
                }
                for (prop in input) {
                    if (scalarWriteablePropertyNames.indexOf(prop) < 0 && objectPropertyNames.indexOf(prop) < 0) {
                        var propertyDescriptor = Object.getOwnPropertyDescriptor(Object.getPrototypeOf(this), prop);
                        if (!propertyDescriptor) {
                            throw new _Internal.RuntimeError({
                                code: CoreErrorCodes.invalidArgument,
                                message: CoreUtility._getResourceString(CommonResourceStrings.propertyDoesNotExist, prop),
                                debugInfo: {
                                    errorLocation: prop
                                }
                            });
                        }
                        if (throwOnReadOnly && !propertyDescriptor.set) {
                            throw new _Internal.RuntimeError({
                                code: CoreErrorCodes.invalidArgument,
                                message: CoreUtility._getResourceString(CommonResourceStrings.attemptingToSetReadOnlyProperty, prop),
                                debugInfo: {
                                    errorLocation: prop
                                }
                            });
                        }
                    }
                }
            }
            catch (innerError) {
                throw new _Internal.RuntimeError({
                    code: CoreErrorCodes.invalidArgument,
                    message: CoreUtility._getResourceString(CoreResourceStrings.invalidArgument, 'properties'),
                    debugInfo: {
                        errorLocation: this._className + '.set'
                    },
                    innerError: innerError
                });
            }
        };
        return ClientObject;
    }(ClientObjectBase));
    OfficeExtension_1.ClientObject = ClientObject;
    var HostBridgeRequestExecutor = (function () {
        function HostBridgeRequestExecutor(session) {
            this.m_session = session;
        }
        HostBridgeRequestExecutor.prototype.executeAsync = function (customData, requestFlags, requestMessage) {
            var httpRequestInfo = {
                url: CoreConstants.processQuery,
                method: 'POST',
                headers: requestMessage.Headers,
                body: requestMessage.Body
            };
            var message = {
                id: HostBridge.nextId(),
                type: 1,
                flags: requestFlags,
                message: httpRequestInfo
            };
            CoreUtility.log(JSON.stringify(message));
            return this.m_session.sendMessageToHost(message).then(function (nativeBridgeResponse) {
                CoreUtility.log('Received response: ' + JSON.stringify(nativeBridgeResponse));
                var responseInfo = nativeBridgeResponse.message;
                var response;
                if (responseInfo.statusCode === 200) {
                    response = {
                        ErrorCode: null,
                        ErrorMessage: null,
                        Headers: responseInfo.headers,
                        Body: CoreUtility._parseResponseBody(responseInfo)
                    };
                }
                else {
                    CoreUtility.log('Error Response:' + responseInfo.body);
                    var error = CoreUtility._parseErrorResponse(responseInfo);
                    response = {
                        ErrorCode: error.errorCode,
                        ErrorMessage: error.errorMessage,
                        Headers: responseInfo.headers,
                        Body: null
                    };
                }
                return response;
            });
        };
        return HostBridgeRequestExecutor;
    }());
    var HostBridgeSession = (function (_super) {
        __extends(HostBridgeSession, _super);
        function HostBridgeSession(m_bridge) {
            var _this = _super.call(this) || this;
            _this.m_bridge = m_bridge;
            _this.m_bridge.addHostMessageHandler(function (message) {
                if (message.type === 3) {
                    GenericEventRegistration.getGenericEventRegistration()._handleRichApiMessage(message.message);
                }
            });
            return _this;
        }
        HostBridgeSession.getInstanceIfHostBridgeInited = function () {
            if (HostBridge.instance) {
                if (CoreUtility.isNullOrUndefined(HostBridgeSession.s_instance) ||
                    HostBridgeSession.s_instance.m_bridge !== HostBridge.instance) {
                    HostBridgeSession.s_instance = new HostBridgeSession(HostBridge.instance);
                }
                return HostBridgeSession.s_instance;
            }
            return null;
        };
        HostBridgeSession.prototype._resolveRequestUrlAndHeaderInfo = function () {
            return CoreUtility._createPromiseFromResult(null);
        };
        HostBridgeSession.prototype._createRequestExecutorOrNull = function () {
            CoreUtility.log('NativeBridgeSession::CreateRequestExecutor');
            return new HostBridgeRequestExecutor(this);
        };
        Object.defineProperty(HostBridgeSession.prototype, "eventRegistration", {
            get: function () {
                return GenericEventRegistration.getGenericEventRegistration();
            },
            enumerable: true,
            configurable: true
        });
        HostBridgeSession.prototype.sendMessageToHost = function (message) {
            return this.m_bridge.sendMessageToHostAndExpectResponse(message);
        };
        return HostBridgeSession;
    }(SessionBase));
    OfficeExtension_1.HostBridgeSession = HostBridgeSession;
    var ClientRequestContext = (function (_super) {
        __extends(ClientRequestContext, _super);
        function ClientRequestContext(url) {
            var _this = _super.call(this) || this;
            _this.m_customRequestHeaders = {};
            _this.m_batchMode = 0;
            _this._onRunFinishedNotifiers = [];
            if (SessionBase._overrideSession) {
                _this.m_requestUrlAndHeaderInfoResolver = SessionBase._overrideSession;
            }
            else {
                if (Utility.isNullOrUndefined(url) || (typeof url === 'string' && url.length === 0)) {
                    url = ClientRequestContext.defaultRequestUrlAndHeaders;
                    if (!url) {
                        url = { url: CoreConstants.localDocument, headers: {} };
                    }
                }
                if (typeof url === 'string') {
                    _this.m_requestUrlAndHeaderInfo = { url: url, headers: {} };
                }
                else if (ClientRequestContext.isRequestUrlAndHeaderInfoResolver(url)) {
                    _this.m_requestUrlAndHeaderInfoResolver = url;
                }
                else if (ClientRequestContext.isRequestUrlAndHeaderInfo(url)) {
                    var requestInfo = url;
                    _this.m_requestUrlAndHeaderInfo = { url: requestInfo.url, headers: {} };
                    CoreUtility._copyHeaders(requestInfo.headers, _this.m_requestUrlAndHeaderInfo.headers);
                }
                else {
                    throw _Internal.RuntimeError._createInvalidArgError({ argumentName: 'url' });
                }
            }
            if (!_this.m_requestUrlAndHeaderInfoResolver &&
                _this.m_requestUrlAndHeaderInfo &&
                CoreUtility._isLocalDocumentUrl(_this.m_requestUrlAndHeaderInfo.url) &&
                HostBridgeSession.getInstanceIfHostBridgeInited()) {
                _this.m_requestUrlAndHeaderInfo = null;
                _this.m_requestUrlAndHeaderInfoResolver = HostBridgeSession.getInstanceIfHostBridgeInited();
            }
            if (_this.m_requestUrlAndHeaderInfoResolver instanceof SessionBase) {
                _this.m_session = _this.m_requestUrlAndHeaderInfoResolver;
            }
            _this._processingResult = false;
            _this._customData = Constants.iterativeExecutor;
            _this.sync = _this.sync.bind(_this);
            return _this;
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
                return _Internal.officeJsEventRegistration;
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
                if (this.m_pendingRequest == null) {
                    this.m_pendingRequest = new ClientRequest(this);
                }
                return this.m_pendingRequest;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(ClientRequestContext.prototype, "debugInfo", {
            get: function () {
                var prettyPrinter = new RequestPrettyPrinter(this._rootObjectPropertyName, this._pendingRequest._objectPaths, this._pendingRequest._actions, OfficeExtension_1._internalConfig.showDisposeInfoInDebugInfo);
                var statements = prettyPrinter.process();
                return { pendingStatements: statements };
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(ClientRequestContext.prototype, "trackedObjects", {
            get: function () {
                if (!this.m_trackedObjects) {
                    this.m_trackedObjects = new TrackedObjects(this);
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
        Object.defineProperty(ClientRequestContext.prototype, "batchMode", {
            get: function () {
                return this.m_batchMode;
            },
            enumerable: true,
            configurable: true
        });
        ClientRequestContext.prototype.ensureInProgressBatchIfBatchMode = function () {
            if (this.m_batchMode === 1 && !this.m_explicitBatchInProgress) {
                throw Utility.createRuntimeError(CoreErrorCodes.generalException, CoreUtility._getResourceString(ResourceStrings.notInsideBatch), null);
            }
        };
        ClientRequestContext.prototype.load = function (clientObj, option) {
            Utility.validateContext(this, clientObj);
            var queryOption = ClientRequestContext._parseQueryOption(option);
            CommonActionFactory.createQueryAction(this, clientObj, queryOption, clientObj);
        };
        ClientRequestContext.prototype.loadRecursive = function (clientObj, options, maxDepth) {
            if (!Utility.isPlainJsonObject(options)) {
                throw _Internal.RuntimeError._createInvalidArgError({ argumentName: 'options' });
            }
            var quries = {};
            for (var key in options) {
                quries[key] = ClientRequestContext._parseQueryOption(options[key]);
            }
            var action = ActionFactory.createRecursiveQueryAction(this, clientObj, { Queries: quries, MaxDepth: maxDepth });
            this._pendingRequest.addActionResultHandler(action, clientObj);
        };
        ClientRequestContext.prototype.trace = function (message) {
            ActionFactory.createTraceAction(this, message, true);
        };
        ClientRequestContext.prototype._processOfficeJsErrorResponse = function (officeJsErrorCode, response) { };
        ClientRequestContext.prototype.ensureRequestUrlAndHeaderInfo = function () {
            var _this = this;
            return Utility._createPromiseFromResult(null).then(function () {
                if (!_this.m_requestUrlAndHeaderInfo) {
                    return _this.m_requestUrlAndHeaderInfoResolver._resolveRequestUrlAndHeaderInfo().then(function (value) {
                        _this.m_requestUrlAndHeaderInfo = value;
                        if (!_this.m_requestUrlAndHeaderInfo) {
                            _this.m_requestUrlAndHeaderInfo = { url: CoreConstants.localDocument, headers: {} };
                        }
                        if (Utility.isNullOrEmptyString(_this.m_requestUrlAndHeaderInfo.url)) {
                            _this.m_requestUrlAndHeaderInfo.url = CoreConstants.localDocument;
                        }
                        if (!_this.m_requestUrlAndHeaderInfo.headers) {
                            _this.m_requestUrlAndHeaderInfo.headers = {};
                        }
                        if (typeof _this.m_requestUrlAndHeaderInfoResolver._createRequestExecutorOrNull === 'function') {
                            var executor = _this.m_requestUrlAndHeaderInfoResolver._createRequestExecutorOrNull();
                            if (executor) {
                                _this._requestExecutor = executor;
                            }
                        }
                    });
                }
            });
        };
        ClientRequestContext.prototype.syncPrivateMain = function () {
            var _this = this;
            return this.ensureRequestUrlAndHeaderInfo().then(function () {
                var req = _this._pendingRequest;
                _this.m_pendingRequest = null;
                return _this.processPreSyncPromises(req).then(function () { return _this.syncPrivate(req); });
            });
        };
        ClientRequestContext.prototype.syncPrivate = function (req) {
            var _this = this;
            if (TestUtility.isMock()) {
                return CoreUtility._createPromiseFromResult(null);
            }
            if (!req.hasActions) {
                return this.processPendingEventHandlers(req);
            }
            var _a = req.buildRequestMessageBodyAndRequestFlags(), msgBody = _a.body, requestFlags = _a.flags;
            if (this._requestFlagModifier) {
                requestFlags |= this._requestFlagModifier;
            }
            if (!this._requestExecutor) {
                if (CoreUtility._isLocalDocumentUrl(this.m_requestUrlAndHeaderInfo.url)) {
                    this._requestExecutor = new OfficeJsRequestExecutor(this);
                }
                else {
                    this._requestExecutor = new HttpRequestExecutor();
                }
            }
            var requestExecutor = this._requestExecutor;
            var headers = {};
            CoreUtility._copyHeaders(this.m_requestUrlAndHeaderInfo.headers, headers);
            CoreUtility._copyHeaders(this.m_customRequestHeaders, headers);
            delete this.m_customRequestHeaders[Constants.officeScriptEventId];
            var requestExecutorRequestMessage = {
                Url: this.m_requestUrlAndHeaderInfo.url,
                Headers: headers,
                Body: msgBody
            };
            req.invalidatePendingInvalidObjectPaths();
            var errorFromResponse = null;
            var errorFromProcessEventHandlers = null;
            this._lastSyncStart = typeof performance === 'undefined' ? 0 : performance.now();
            this._lastRequestFlags = requestFlags;
            return requestExecutor
                .executeAsync(this._customData, requestFlags, requestExecutorRequestMessage)
                .then(function (response) {
                _this._lastSyncEnd = typeof performance === 'undefined' ? 0 : performance.now();
                errorFromResponse = _this.processRequestExecutorResponseMessage(req, response);
                return _this.processPendingEventHandlers(req).catch(function (ex) {
                    CoreUtility.log('Error in processPendingEventHandlers');
                    CoreUtility.log(JSON.stringify(ex));
                    errorFromProcessEventHandlers = ex;
                });
            })
                .then(function () {
                if (errorFromResponse) {
                    CoreUtility.log('Throw error from response: ' + JSON.stringify(errorFromResponse));
                    throw errorFromResponse;
                }
                if (errorFromProcessEventHandlers) {
                    CoreUtility.log('Throw error from ProcessEventHandler: ' + JSON.stringify(errorFromProcessEventHandlers));
                    var transformedError = null;
                    if (errorFromProcessEventHandlers instanceof _Internal.RuntimeError) {
                        transformedError = errorFromProcessEventHandlers;
                        transformedError.traceMessages = req._responseTraceMessages;
                    }
                    else {
                        var message = null;
                        if (typeof errorFromProcessEventHandlers === 'string') {
                            message = errorFromProcessEventHandlers;
                        }
                        else {
                            message = errorFromProcessEventHandlers.message;
                        }
                        if (Utility.isNullOrEmptyString(message)) {
                            message = CoreUtility._getResourceString(ResourceStrings.cannotRegisterEvent);
                        }
                        transformedError = new _Internal.RuntimeError({
                            code: ErrorCodes.cannotRegisterEvent,
                            message: message,
                            traceMessages: req._responseTraceMessages
                        });
                    }
                    throw transformedError;
                }
            });
        };
        ClientRequestContext.prototype.processRequestExecutorResponseMessage = function (req, response) {
            if (response.Body && response.Body.TraceIds) {
                req._setResponseTraceIds(response.Body.TraceIds);
            }
            var traceMessages = req._responseTraceMessages;
            var errorStatementInfo = null;
            if (response.Body) {
                if (response.Body.Error && response.Body.Error.ActionIndex >= 0) {
                    var prettyPrinter = new RequestPrettyPrinter(this._rootObjectPropertyName, req._objectPaths, req._actions, false, true);
                    var debugInfoStatementInfo = prettyPrinter.processForDebugStatementInfo(response.Body.Error.ActionIndex);
                    errorStatementInfo = {
                        statement: debugInfoStatementInfo.statement,
                        surroundingStatements: debugInfoStatementInfo.surroundingStatements,
                        fullStatements: ['Please enable config.extendedErrorLogging to see full statements.']
                    };
                    if (OfficeExtension_1.config.extendedErrorLogging) {
                        prettyPrinter = new RequestPrettyPrinter(this._rootObjectPropertyName, req._objectPaths, req._actions, false, false);
                        errorStatementInfo.fullStatements = prettyPrinter.process();
                    }
                }
                var actionResults = null;
                if (response.Body.Results) {
                    actionResults = response.Body.Results;
                }
                else if (response.Body.ProcessedResults && response.Body.ProcessedResults.Results) {
                    actionResults = response.Body.ProcessedResults.Results;
                }
                if (actionResults) {
                    this._processingResult = true;
                    try {
                        req.processResponse(actionResults);
                    }
                    finally {
                        this._processingResult = false;
                    }
                }
            }
            if (!Utility.isNullOrEmptyString(response.ErrorCode)) {
                return new _Internal.RuntimeError({
                    code: response.ErrorCode,
                    message: response.ErrorMessage,
                    traceMessages: traceMessages
                });
            }
            else if (response.Body && response.Body.Error) {
                var debugInfo = {
                    errorLocation: response.Body.Error.Location
                };
                if (errorStatementInfo) {
                    debugInfo.statement = errorStatementInfo.statement;
                    debugInfo.surroundingStatements = errorStatementInfo.surroundingStatements;
                    debugInfo.fullStatements = errorStatementInfo.fullStatements;
                }
                return new _Internal.RuntimeError({
                    code: response.Body.Error.Code,
                    message: response.Body.Error.Message,
                    traceMessages: traceMessages,
                    debugInfo: debugInfo
                });
            }
            return null;
        };
        ClientRequestContext.prototype.processPendingEventHandlers = function (req) {
            var ret = Utility._createPromiseFromResult(null);
            for (var i = 0; i < req._pendingProcessEventHandlers.length; i++) {
                var eventHandlers = req._pendingProcessEventHandlers[i];
                ret = ret.then(this.createProcessOneEventHandlersFunc(eventHandlers, req));
            }
            return ret;
        };
        ClientRequestContext.prototype.createProcessOneEventHandlersFunc = function (eventHandlers, req) {
            return function () { return eventHandlers._processRegistration(req); };
        };
        ClientRequestContext.prototype.processPreSyncPromises = function (req) {
            var ret = Utility._createPromiseFromResult(null);
            for (var i = 0; i < req._preSyncPromises.length; i++) {
                var p = req._preSyncPromises[i];
                ret = ret.then(this.createProcessOneProSyncFunc(p));
            }
            return ret;
        };
        ClientRequestContext.prototype.createProcessOneProSyncFunc = function (p) {
            return function () { return p; };
        };
        ClientRequestContext.prototype.sync = function (passThroughValue) {
            if (TestUtility.isMock()) {
                return CoreUtility._createPromiseFromResult(passThroughValue);
            }
            return this.syncPrivateMain().then(function () { return passThroughValue; });
        };
        ClientRequestContext.prototype.batch = function (batchBody) {
            var _this = this;
            if (this.m_batchMode !== 1) {
                return CoreUtility._createPromiseFromException(Utility.createRuntimeError(CoreErrorCodes.generalException, null, null));
            }
            if (this.m_explicitBatchInProgress) {
                return CoreUtility._createPromiseFromException(Utility.createRuntimeError(CoreErrorCodes.generalException, CoreUtility._getResourceString(ResourceStrings.pendingBatchInProgress), null));
            }
            if (Utility.isNullOrUndefined(batchBody)) {
                return Utility._createPromiseFromResult(null);
            }
            this.m_explicitBatchInProgress = true;
            var previousRequest = this.m_pendingRequest;
            this.m_pendingRequest = new ClientRequest(this);
            var batchBodyResult;
            try {
                batchBodyResult = batchBody(this._rootObject, this);
            }
            catch (ex) {
                this.m_explicitBatchInProgress = false;
                this.m_pendingRequest = previousRequest;
                return CoreUtility._createPromiseFromException(ex);
            }
            var request;
            var batchBodyResultPromise;
            if (typeof batchBodyResult === 'object' && batchBodyResult && typeof batchBodyResult.then === 'function') {
                batchBodyResultPromise = Utility._createPromiseFromResult(null)
                    .then(function () {
                    return batchBodyResult;
                })
                    .then(function (result) {
                    _this.m_explicitBatchInProgress = false;
                    request = _this.m_pendingRequest;
                    _this.m_pendingRequest = previousRequest;
                    return result;
                })
                    .catch(function (ex) {
                    _this.m_explicitBatchInProgress = false;
                    request = _this.m_pendingRequest;
                    _this.m_pendingRequest = previousRequest;
                    return CoreUtility._createPromiseFromException(ex);
                });
            }
            else {
                this.m_explicitBatchInProgress = false;
                request = this.m_pendingRequest;
                this.m_pendingRequest = previousRequest;
                batchBodyResultPromise = Utility._createPromiseFromResult(batchBodyResult);
            }
            return batchBodyResultPromise.then(function (result) {
                return _this.ensureRequestUrlAndHeaderInfo()
                    .then(function () {
                    return _this.syncPrivate(request);
                })
                    .then(function () {
                    return result;
                });
            });
        };
        ClientRequestContext._run = function (ctxInitializer, runBody, numCleanupAttempts, retryDelay, onCleanupSuccess, onCleanupFailure) {
            if (numCleanupAttempts === void 0) { numCleanupAttempts = 3; }
            if (retryDelay === void 0) { retryDelay = 5000; }
            return ClientRequestContext._runCommon('run', null, ctxInitializer, 0, runBody, numCleanupAttempts, retryDelay, null, onCleanupSuccess, onCleanupFailure);
        };
        ClientRequestContext.isValidRequestInfo = function (value) {
            return (typeof value === 'string' ||
                ClientRequestContext.isRequestUrlAndHeaderInfo(value) ||
                ClientRequestContext.isRequestUrlAndHeaderInfoResolver(value));
        };
        ClientRequestContext.isRequestUrlAndHeaderInfo = function (value) {
            return (typeof value === 'object' &&
                value !== null &&
                Object.getPrototypeOf(value) === Object.getPrototypeOf({}) &&
                !Utility.isNullOrUndefined(value.url));
        };
        ClientRequestContext.isRequestUrlAndHeaderInfoResolver = function (value) {
            return typeof value === 'object' && value !== null && typeof value._resolveRequestUrlAndHeaderInfo === 'function';
        };
        ClientRequestContext._runBatch = function (functionName, receivedRunArgs, ctxInitializer, onBeforeRun, numCleanupAttempts, retryDelay, onCleanupSuccess, onCleanupFailure) {
            if (numCleanupAttempts === void 0) { numCleanupAttempts = 3; }
            if (retryDelay === void 0) { retryDelay = 5000; }
            return ClientRequestContext._runBatchCommon(0, functionName, receivedRunArgs, ctxInitializer, numCleanupAttempts, retryDelay, onBeforeRun, onCleanupSuccess, onCleanupFailure);
        };
        ClientRequestContext._runExplicitBatch = function (functionName, receivedRunArgs, ctxInitializer, onBeforeRun, numCleanupAttempts, retryDelay, onCleanupSuccess, onCleanupFailure) {
            if (numCleanupAttempts === void 0) { numCleanupAttempts = 3; }
            if (retryDelay === void 0) { retryDelay = 5000; }
            return ClientRequestContext._runBatchCommon(1, functionName, receivedRunArgs, ctxInitializer, numCleanupAttempts, retryDelay, onBeforeRun, onCleanupSuccess, onCleanupFailure);
        };
        ClientRequestContext._runBatchCommon = function (batchMode, functionName, receivedRunArgs, ctxInitializer, numCleanupAttempts, retryDelay, onBeforeRun, onCleanupSuccess, onCleanupFailure) {
            if (numCleanupAttempts === void 0) { numCleanupAttempts = 3; }
            if (retryDelay === void 0) { retryDelay = 5000; }
            var ctxRetriever;
            var batch;
            var requestInfo = null;
            var previousObjects = null;
            var argOffset = 0;
            var options = null;
            if (receivedRunArgs.length > 0) {
                if (ClientRequestContext.isValidRequestInfo(receivedRunArgs[0])) {
                    requestInfo = receivedRunArgs[0];
                    argOffset = 1;
                }
                else if (Utility.isPlainJsonObject(receivedRunArgs[0])) {
                    options = receivedRunArgs[0];
                    requestInfo = options.session;
                    if (requestInfo != null && !ClientRequestContext.isValidRequestInfo(requestInfo)) {
                        return ClientRequestContext.createErrorPromise(functionName);
                    }
                    previousObjects = options.previousObjects;
                    argOffset = 1;
                }
            }
            if (receivedRunArgs.length == argOffset + 1) {
                batch = receivedRunArgs[argOffset + 0];
            }
            else if (options == null && receivedRunArgs.length == argOffset + 2) {
                previousObjects = receivedRunArgs[argOffset + 0];
                batch = receivedRunArgs[argOffset + 1];
            }
            else {
                return ClientRequestContext.createErrorPromise(functionName);
            }
            if (previousObjects != null) {
                if (previousObjects instanceof ClientObject) {
                    ctxRetriever = function () { return previousObjects.context; };
                }
                else if (previousObjects instanceof ClientRequestContext) {
                    ctxRetriever = function () { return previousObjects; };
                }
                else if (Array.isArray(previousObjects)) {
                    var array = previousObjects;
                    if (array.length == 0) {
                        return ClientRequestContext.createErrorPromise(functionName);
                    }
                    for (var i = 0; i < array.length; i++) {
                        if (!(array[i] instanceof ClientObject)) {
                            return ClientRequestContext.createErrorPromise(functionName);
                        }
                        if (array[i].context != array[0].context) {
                            return ClientRequestContext.createErrorPromise(functionName, ResourceStrings.invalidRequestContext);
                        }
                    }
                    ctxRetriever = function () { return array[0].context; };
                }
                else {
                    return ClientRequestContext.createErrorPromise(functionName);
                }
            }
            else {
                ctxRetriever = ctxInitializer;
            }
            var onBeforeRunWithOptions = null;
            if (onBeforeRun) {
                onBeforeRunWithOptions = function (context) { return onBeforeRun(options || {}, context); };
            }
            return ClientRequestContext._runCommon(functionName, requestInfo, ctxRetriever, batchMode, batch, numCleanupAttempts, retryDelay, onBeforeRunWithOptions, onCleanupSuccess, onCleanupFailure);
        };
        ClientRequestContext.createErrorPromise = function (functionName, code) {
            if (code === void 0) { code = CoreResourceStrings.invalidArgument; }
            return CoreUtility._createPromiseFromException(Utility.createRuntimeError(code, CoreUtility._getResourceString(code), functionName));
        };
        ClientRequestContext._runCommon = function (functionName, requestInfo, ctxRetriever, batchMode, runBody, numCleanupAttempts, retryDelay, onBeforeRun, onCleanupSuccess, onCleanupFailure) {
            if (SessionBase._overrideSession) {
                requestInfo = SessionBase._overrideSession;
            }
            var starterPromise = CoreUtility.createPromise(function (resolve, reject) {
                resolve();
            });
            var ctx;
            var succeeded = false;
            var resultOrError;
            var previousBatchMode;
            return starterPromise
                .then(function () {
                ctx = ctxRetriever(requestInfo);
                if (ctx._autoCleanup) {
                    return new OfficeExtension_1.Promise(function (resolve, reject) {
                        ctx._onRunFinishedNotifiers.push(function () {
                            ctx._autoCleanup = true;
                            resolve();
                        });
                    });
                }
                else {
                    ctx._autoCleanup = true;
                }
            })
                .then(function () {
                if (typeof runBody !== 'function') {
                    return ClientRequestContext.createErrorPromise(functionName);
                }
                previousBatchMode = ctx.m_batchMode;
                ctx.m_batchMode = batchMode;
                if (onBeforeRun) {
                    onBeforeRun(ctx);
                }
                var runBodyResult;
                if (batchMode == 1) {
                    runBodyResult = runBody(ctx.batch.bind(ctx));
                }
                else {
                    runBodyResult = runBody(ctx);
                }
                if (Utility.isNullOrUndefined(runBodyResult) || typeof runBodyResult.then !== 'function') {
                    Utility.throwError(ResourceStrings.runMustReturnPromise);
                }
                return runBodyResult;
            })
                .then(function (runBodyResult) {
                if (batchMode === 1) {
                    return runBodyResult;
                }
                else {
                    return ctx.sync(runBodyResult);
                }
            })
                .then(function (result) {
                succeeded = true;
                resultOrError = result;
            })
                .catch(function (error) {
                resultOrError = error;
            })
                .then(function () {
                var itemsToRemove = ctx.trackedObjects._retrieveAndClearAutoCleanupList();
                ctx._autoCleanup = false;
                ctx.m_batchMode = previousBatchMode;
                for (var key in itemsToRemove) {
                    itemsToRemove[key]._objectPath.isValid = false;
                }
                var cleanupCounter = 0;
                if (Utility._synchronousCleanup || ClientRequestContext.isRequestUrlAndHeaderInfoResolver(requestInfo)) {
                    return attemptCleanup();
                }
                else {
                    attemptCleanup();
                }
                function attemptCleanup() {
                    cleanupCounter++;
                    var savedPendingRequest = ctx.m_pendingRequest;
                    var savedBatchMode = ctx.m_batchMode;
                    var request = new ClientRequest(ctx);
                    ctx.m_pendingRequest = request;
                    ctx.m_batchMode = 0;
                    try {
                        for (var key in itemsToRemove) {
                            ctx.trackedObjects.remove(itemsToRemove[key]);
                        }
                    }
                    finally {
                        ctx.m_batchMode = savedBatchMode;
                        ctx.m_pendingRequest = savedPendingRequest;
                    }
                    return ctx
                        .syncPrivate(request)
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
                    var func = ctx._onRunFinishedNotifiers.shift();
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
        return ClientRequestContext;
    }(ClientRequestContextBase));
    OfficeExtension_1.ClientRequestContext = ClientRequestContext;
    var RetrieveResultImpl = (function () {
        function RetrieveResultImpl(m_proxy, m_shouldPolyfill) {
            this.m_proxy = m_proxy;
            this.m_shouldPolyfill = m_shouldPolyfill;
            var scalarPropertyNames = m_proxy[Constants.scalarPropertyNames];
            var navigationPropertyNames = m_proxy[Constants.navigationPropertyNames];
            var typeName = m_proxy[Constants.className];
            var isCollection = m_proxy[Constants.isCollection];
            if (scalarPropertyNames) {
                for (var i = 0; i < scalarPropertyNames.length; i++) {
                    Utility.definePropertyThrowUnloadedException(this, typeName, scalarPropertyNames[i]);
                }
            }
            if (navigationPropertyNames) {
                for (var i = 0; i < navigationPropertyNames.length; i++) {
                    Utility.definePropertyThrowUnloadedException(this, typeName, navigationPropertyNames[i]);
                }
            }
            if (isCollection) {
                Utility.definePropertyThrowUnloadedException(this, typeName, Constants.itemsLowerCase);
            }
        }
        Object.defineProperty(RetrieveResultImpl.prototype, "$proxy", {
            get: function () {
                return this.m_proxy;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(RetrieveResultImpl.prototype, "$isNullObject", {
            get: function () {
                if (!this.m_isLoaded) {
                    throw new _Internal.RuntimeError({
                        code: ErrorCodes.valueNotLoaded,
                        message: CoreUtility._getResourceString(ResourceStrings.valueNotLoaded),
                        debugInfo: {
                            errorLocation: 'retrieveResult.$isNullObject'
                        }
                    });
                }
                return this.m_isNullObject;
            },
            enumerable: true,
            configurable: true
        });
        RetrieveResultImpl.prototype.toJSON = function () {
            if (!this.m_isLoaded) {
                return undefined;
            }
            if (this.m_isNullObject) {
                return null;
            }
            if (Utility.isUndefined(this.m_json)) {
                this.m_json = Utility.purifyJson(this.m_value);
            }
            return this.m_json;
        };
        RetrieveResultImpl.prototype.toString = function () {
            return JSON.stringify(this.toJSON());
        };
        RetrieveResultImpl.prototype._handleResult = function (value) {
            this.m_isLoaded = true;
            if (value === null || (typeof value === 'object' && value && value._IsNull)) {
                this.m_isNullObject = true;
                value = null;
            }
            else {
                this.m_isNullObject = false;
            }
            if (this.m_shouldPolyfill) {
                value = Utility.changePropertyNameToCamelLowerCase(value);
            }
            this.m_value = value;
            this.m_proxy._handleRetrieveResult(value, this);
        };
        return RetrieveResultImpl;
    }());
    var Constants = (function (_super) {
        __extends(Constants, _super);
        function Constants() {
            return _super !== null && _super.apply(this, arguments) || this;
        }
        Constants.getItemAt = 'GetItemAt';
        Constants.index = '_Index';
        Constants.iterativeExecutor = 'IterativeExecutor';
        Constants.isTracked = '_IsTracked';
        Constants.eventMessageCategory = 65536;
        Constants.eventWorkbookId = 'Workbook';
        Constants.eventSourceRemote = 'Remote';
        Constants.proxy = '$proxy';
        Constants.className = '_className';
        Constants.isCollection = '_isCollection';
        Constants.collectionPropertyPath = '_collectionPropertyPath';
        Constants.objectPathInfoDoNotKeepReferenceFieldName = 'D';
        Constants.officeScriptEventId = 'X-OfficeScriptEventId';
        return Constants;
    }(CommonConstants));
    OfficeExtension_1.Constants = Constants;
    var ClientRequest = (function (_super) {
        __extends(ClientRequest, _super);
        function ClientRequest(context) {
            var _this = _super.call(this, context) || this;
            _this.m_context = context;
            _this.m_pendingProcessEventHandlers = [];
            _this.m_pendingEventHandlerActions = {};
            _this.m_traceInfos = {};
            _this.m_responseTraceIds = {};
            _this.m_responseTraceMessages = [];
            return _this;
        }
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
        ClientRequest.prototype._setResponseTraceIds = function (value) {
            if (value) {
                for (var i = 0; i < value.length; i++) {
                    var traceId = value[i];
                    this.m_responseTraceIds[traceId] = traceId;
                    var message = this.m_traceInfos[traceId];
                    if (!CoreUtility.isNullOrUndefined(message)) {
                        this.m_responseTraceMessages.push(message);
                    }
                }
            }
        };
        ClientRequest.prototype.addTrace = function (actionId, message) {
            this.m_traceInfos[actionId] = message;
        };
        ClientRequest.prototype._addPendingEventHandlerAction = function (eventHandlers, action) {
            if (!this.m_pendingEventHandlerActions[eventHandlers._id]) {
                this.m_pendingEventHandlerActions[eventHandlers._id] = [];
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
        ClientRequest.prototype._getPendingEventHandlerActions = function (eventHandlers) {
            return this.m_pendingEventHandlerActions[eventHandlers._id];
        };
        return ClientRequest;
    }(ClientRequestBase));
    OfficeExtension_1.ClientRequest = ClientRequest;
    var EventHandlers = (function () {
        function EventHandlers(context, parentObject, name, eventInfo) {
            var _this = this;
            this.m_id = context._nextId();
            this.m_context = context;
            this.m_name = name;
            this.m_handlers = [];
            this.m_registered = false;
            this.m_eventInfo = eventInfo;
            this.m_callback = function (args) {
                _this.m_eventInfo.eventArgsTransformFunc(args).then(function (newArgs) { return _this.fireEvent(newArgs); });
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
        EventHandlers.prototype.add = function (handler) {
            var action = ActionFactory.createTraceAction(this.m_context, null, false);
            this.m_context._pendingRequest._addPendingEventHandlerAction(this, {
                id: action.actionInfo.Id,
                handler: handler,
                operation: 0
            });
            return new EventHandlerResult(this.m_context, this, handler);
        };
        EventHandlers.prototype.remove = function (handler) {
            var action = ActionFactory.createTraceAction(this.m_context, null, false);
            this.m_context._pendingRequest._addPendingEventHandlerAction(this, {
                id: action.actionInfo.Id,
                handler: handler,
                operation: 1
            });
        };
        EventHandlers.prototype.removeAll = function () {
            var action = ActionFactory.createTraceAction(this.m_context, null, false);
            this.m_context._pendingRequest._addPendingEventHandlerAction(this, {
                id: action.actionInfo.Id,
                handler: null,
                operation: 2
            });
        };
        EventHandlers.prototype._processRegistration = function (req) {
            var _this = this;
            var ret = CoreUtility._createPromiseFromResult(null);
            var actions = req._getPendingEventHandlerActions(this);
            if (!actions) {
                return ret;
            }
            var handlersResult = [];
            for (var i = 0; i < this.m_handlers.length; i++) {
                handlersResult.push(this.m_handlers[i]);
            }
            var hasChange = false;
            for (var i = 0; i < actions.length; i++) {
                if (req._responseTraceIds[actions[i].id]) {
                    hasChange = true;
                    switch (actions[i].operation) {
                        case 0:
                            handlersResult.push(actions[i].handler);
                            break;
                        case 1:
                            for (var index = handlersResult.length - 1; index >= 0; index--) {
                                if (handlersResult[index] === actions[i].handler) {
                                    handlersResult.splice(index, 1);
                                    break;
                                }
                            }
                            break;
                        case 2:
                            handlersResult = [];
                            break;
                    }
                }
            }
            if (hasChange) {
                if (!this.m_registered && handlersResult.length > 0) {
                    ret = ret.then(function () { return _this.m_eventInfo.registerFunc(_this.m_callback); }).then(function () { return (_this.m_registered = true); });
                }
                else if (this.m_registered && handlersResult.length == 0) {
                    ret = ret
                        .then(function () { return _this.m_eventInfo.unregisterFunc(_this.m_callback); })
                        .catch(function (ex) {
                        CoreUtility.log('Error when unregister event: ' + JSON.stringify(ex));
                    })
                        .then(function () { return (_this.m_registered = false); });
                }
                ret = ret.then(function () { return (_this.m_handlers = handlersResult); });
            }
            return ret;
        };
        EventHandlers.prototype.fireEvent = function (args) {
            var promises = [];
            for (var i = 0; i < this.m_handlers.length; i++) {
                var handler = this.m_handlers[i];
                var p = CoreUtility._createPromiseFromResult(null)
                    .then(this.createFireOneEventHandlerFunc(handler, args))
                    .catch(function (ex) {
                    CoreUtility.log('Error when invoke handler: ' + JSON.stringify(ex));
                });
                promises.push(p);
            }
            CoreUtility.Promise.all(promises);
        };
        EventHandlers.prototype.createFireOneEventHandlerFunc = function (handler, args) {
            return function () { return handler(args); };
        };
        return EventHandlers;
    }());
    OfficeExtension_1.EventHandlers = EventHandlers;
    var EventHandlerResult = (function () {
        function EventHandlerResult(context, handlers, handler) {
            this.m_context = context;
            this.m_allHandlers = handlers;
            this.m_handler = handler;
        }
        Object.defineProperty(EventHandlerResult.prototype, "context", {
            get: function () {
                return this.m_context;
            },
            enumerable: true,
            configurable: true
        });
        EventHandlerResult.prototype.remove = function () {
            if (this.m_allHandlers && this.m_handler) {
                this.m_allHandlers.remove(this.m_handler);
                this.m_allHandlers = null;
                this.m_handler = null;
            }
        };
        return EventHandlerResult;
    }());
    OfficeExtension_1.EventHandlerResult = EventHandlerResult;
    (function (_Internal) {
        var OfficeJsEventRegistration = (function () {
            function OfficeJsEventRegistration() {
            }
            OfficeJsEventRegistration.prototype.register = function (eventId, targetId, handler) {
                switch (eventId) {
                    case 4:
                        return Utility.promisify(function (callback) { return Office.context.document.bindings.getByIdAsync(targetId, callback); }).then(function (officeBinding) {
                            return Utility.promisify(function (callback) {
                                return officeBinding.addHandlerAsync(Office.EventType.BindingDataChanged, handler, callback);
                            });
                        });
                    case 3:
                        return Utility.promisify(function (callback) { return Office.context.document.bindings.getByIdAsync(targetId, callback); }).then(function (officeBinding) {
                            return Utility.promisify(function (callback) {
                                return officeBinding.addHandlerAsync(Office.EventType.BindingSelectionChanged, handler, callback);
                            });
                        });
                    case 2:
                        return Utility.promisify(function (callback) {
                            return Office.context.document.addHandlerAsync(Office.EventType.DocumentSelectionChanged, handler, callback);
                        });
                    case 1:
                        return Utility.promisify(function (callback) {
                            return Office.context.document.settings.addHandlerAsync(Office.EventType.SettingsChanged, handler, callback);
                        });
                    case 5:
                        return Utility.promisify(function (callback) {
                            return OSF.DDA.RichApi.richApiMessageManager.addHandlerAsync('richApiMessage', handler, callback);
                        });
                    case 13:
                        return Utility.promisify(function (callback) {
                            return Office.context.document.addHandlerAsync(Office.EventType.ObjectDeleted, handler, { id: targetId }, callback);
                        });
                    case 14:
                        return Utility.promisify(function (callback) {
                            return Office.context.document.addHandlerAsync(Office.EventType.ObjectSelectionChanged, handler, { id: targetId }, callback);
                        });
                    case 15:
                        return Utility.promisify(function (callback) {
                            return Office.context.document.addHandlerAsync(Office.EventType.ObjectDataChanged, handler, { id: targetId }, callback);
                        });
                    case 16:
                        return Utility.promisify(function (callback) {
                            return Office.context.document.addHandlerAsync(Office.EventType.ContentControlAdded, handler, { id: targetId }, callback);
                        });
                    default:
                        throw _Internal.RuntimeError._createInvalidArgError({ argumentName: 'eventId' });
                }
            };
            OfficeJsEventRegistration.prototype.unregister = function (eventId, targetId, handler) {
                switch (eventId) {
                    case 4:
                        return Utility.promisify(function (callback) { return Office.context.document.bindings.getByIdAsync(targetId, callback); }).then(function (officeBinding) {
                            return Utility.promisify(function (callback) {
                                return officeBinding.removeHandlerAsync(Office.EventType.BindingDataChanged, { handler: handler }, callback);
                            });
                        });
                    case 3:
                        return Utility.promisify(function (callback) { return Office.context.document.bindings.getByIdAsync(targetId, callback); }).then(function (officeBinding) {
                            return Utility.promisify(function (callback) {
                                return officeBinding.removeHandlerAsync(Office.EventType.BindingSelectionChanged, { handler: handler }, callback);
                            });
                        });
                    case 2:
                        return Utility.promisify(function (callback) {
                            return Office.context.document.removeHandlerAsync(Office.EventType.DocumentSelectionChanged, { handler: handler }, callback);
                        });
                    case 1:
                        return Utility.promisify(function (callback) {
                            return Office.context.document.settings.removeHandlerAsync(Office.EventType.SettingsChanged, { handler: handler }, callback);
                        });
                    case 5:
                        return Utility.promisify(function (callback) {
                            return OSF.DDA.RichApi.richApiMessageManager.removeHandlerAsync('richApiMessage', { handler: handler }, callback);
                        });
                    case 13:
                        return Utility.promisify(function (callback) {
                            return Office.context.document.removeHandlerAsync(Office.EventType.ObjectDeleted, { id: targetId, handler: handler }, callback);
                        });
                    case 14:
                        return Utility.promisify(function (callback) {
                            return Office.context.document.removeHandlerAsync(Office.EventType.ObjectSelectionChanged, { id: targetId, handler: handler }, callback);
                        });
                    case 15:
                        return Utility.promisify(function (callback) {
                            return Office.context.document.removeHandlerAsync(Office.EventType.ObjectDataChanged, { id: targetId, handler: handler }, callback);
                        });
                    case 16:
                        return Utility.promisify(function (callback) {
                            return Office.context.document.removeHandlerAsync(Office.EventType.ContentControlAdded, { id: targetId, handler: handler }, callback);
                        });
                    default:
                        throw _Internal.RuntimeError._createInvalidArgError({ argumentName: 'eventId' });
                }
            };
            return OfficeJsEventRegistration;
        }());
        _Internal.officeJsEventRegistration = new OfficeJsEventRegistration();
    })(_Internal = OfficeExtension_1._Internal || (OfficeExtension_1._Internal = {}));
    var EventRegistration = (function () {
        function EventRegistration(registerEventImpl, unregisterEventImpl) {
            this.m_handlersByEventByTarget = {};
            this.m_registerEventImpl = registerEventImpl;
            this.m_unregisterEventImpl = unregisterEventImpl;
        }
        EventRegistration.getTargetIdOrDefault = function (targetId) {
            if (Utility.isNullOrUndefined(targetId)) {
                return '';
            }
            return targetId;
        };
        EventRegistration.prototype.getHandlers = function (eventId, targetId) {
            targetId = EventRegistration.getTargetIdOrDefault(targetId);
            var handlersById = this.m_handlersByEventByTarget[eventId];
            if (!handlersById) {
                handlersById = {};
                this.m_handlersByEventByTarget[eventId] = handlersById;
            }
            var handlers = handlersById[targetId];
            if (!handlers) {
                handlers = [];
                handlersById[targetId] = handlers;
            }
            return handlers;
        };
        EventRegistration.prototype.callHandlers = function (eventId, targetId, argument) {
            var funcs = this.getHandlers(eventId, targetId);
            for (var i = 0; i < funcs.length; i++) {
                funcs[i](argument);
            }
        };
        EventRegistration.prototype.hasHandlers = function (eventId, targetId) {
            return this.getHandlers(eventId, targetId).length > 0;
        };
        EventRegistration.prototype.register = function (eventId, targetId, handler) {
            if (!handler) {
                throw _Internal.RuntimeError._createInvalidArgError({ argumentName: 'handler' });
            }
            var handlers = this.getHandlers(eventId, targetId);
            handlers.push(handler);
            if (handlers.length === 1) {
                return this.m_registerEventImpl(eventId, targetId);
            }
            return Utility._createPromiseFromResult(null);
        };
        EventRegistration.prototype.unregister = function (eventId, targetId, handler) {
            if (!handler) {
                throw _Internal.RuntimeError._createInvalidArgError({ argumentName: 'handler' });
            }
            var handlers = this.getHandlers(eventId, targetId);
            for (var index = handlers.length - 1; index >= 0; index--) {
                if (handlers[index] === handler) {
                    handlers.splice(index, 1);
                    break;
                }
            }
            if (handlers.length === 0) {
                return this.m_unregisterEventImpl(eventId, targetId);
            }
            return Utility._createPromiseFromResult(null);
        };
        return EventRegistration;
    }());
    OfficeExtension_1.EventRegistration = EventRegistration;
    var GenericEventRegistration = (function () {
        function GenericEventRegistration() {
            this.m_eventRegistration = new EventRegistration(this._registerEventImpl.bind(this), this._unregisterEventImpl.bind(this));
            this.m_richApiMessageHandler = this._handleRichApiMessage.bind(this);
        }
        GenericEventRegistration.prototype.ready = function () {
            var _this = this;
            if (!this.m_ready) {
                if (GenericEventRegistration._testReadyImpl) {
                    this.m_ready = GenericEventRegistration._testReadyImpl().then(function () {
                        _this.m_isReady = true;
                    });
                }
                else if (HostBridge.instance) {
                    this.m_ready = Utility._createPromiseFromResult(null).then(function () {
                        _this.m_isReady = true;
                    });
                }
                else {
                    this.m_ready = _Internal.officeJsEventRegistration
                        .register(5, '', this.m_richApiMessageHandler)
                        .then(function () {
                        _this.m_isReady = true;
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
        GenericEventRegistration.prototype.register = function (eventId, targetId, handler) {
            var _this = this;
            return this.ready().then(function () { return _this.m_eventRegistration.register(eventId, targetId, handler); });
        };
        GenericEventRegistration.prototype.unregister = function (eventId, targetId, handler) {
            var _this = this;
            return this.ready().then(function () { return _this.m_eventRegistration.unregister(eventId, targetId, handler); });
        };
        GenericEventRegistration.prototype._registerEventImpl = function (eventId, targetId) {
            return Utility._createPromiseFromResult(null);
        };
        GenericEventRegistration.prototype._unregisterEventImpl = function (eventId, targetId) {
            return Utility._createPromiseFromResult(null);
        };
        GenericEventRegistration.prototype._handleRichApiMessage = function (msg) {
            if (msg && msg.entries) {
                for (var entryIndex = 0; entryIndex < msg.entries.length; entryIndex++) {
                    var entry = msg.entries[entryIndex];
                    if (entry.messageCategory == Constants.eventMessageCategory) {
                        if (CoreUtility._logEnabled) {
                            CoreUtility.log(JSON.stringify(entry));
                        }
                        var eventId = entry.messageType;
                        var targetId = entry.targetId;
                        var hasHandlers = this.m_eventRegistration.hasHandlers(eventId, targetId);
                        if (hasHandlers) {
                            var arg = JSON.parse(entry.message);
                            if (entry.isRemoteOverride) {
                                arg.source = Constants.eventSourceRemote;
                            }
                            this.m_eventRegistration.callHandlers(eventId, targetId, arg);
                        }
                    }
                }
            }
        };
        GenericEventRegistration.getGenericEventRegistration = function () {
            if (!GenericEventRegistration.s_genericEventRegistration) {
                GenericEventRegistration.s_genericEventRegistration = new GenericEventRegistration();
            }
            return GenericEventRegistration.s_genericEventRegistration;
        };
        GenericEventRegistration.richApiMessageEventCategory = 65536;
        return GenericEventRegistration;
    }());
    OfficeExtension_1.GenericEventRegistration = GenericEventRegistration;
    function _testSetRichApiMessageReadyImpl(impl) {
        GenericEventRegistration._testReadyImpl = impl;
    }
    OfficeExtension_1._testSetRichApiMessageReadyImpl = _testSetRichApiMessageReadyImpl;
    function _testTriggerRichApiMessageEvent(msg) {
        GenericEventRegistration.getGenericEventRegistration()._handleRichApiMessage(msg);
    }
    OfficeExtension_1._testTriggerRichApiMessageEvent = _testTriggerRichApiMessageEvent;
    var GenericEventHandlers = (function (_super) {
        __extends(GenericEventHandlers, _super);
        function GenericEventHandlers(context, parentObject, name, eventInfo) {
            var _this = _super.call(this, context, parentObject, name, eventInfo) || this;
            _this.m_genericEventInfo = eventInfo;
            return _this;
        }
        GenericEventHandlers.prototype.add = function (handler) {
            var _this = this;
            if (this._handlers.length == 0 && this.m_genericEventInfo.registerFunc) {
                this.m_genericEventInfo.registerFunc();
            }
            if (!GenericEventRegistration.getGenericEventRegistration().isReady) {
                this._context._pendingRequest._addPreSyncPromise(GenericEventRegistration.getGenericEventRegistration().ready());
            }
            ActionFactory.createTraceMarkerForCallback(this._context, function () {
                _this._handlers.push(handler);
                if (_this._handlers.length == 1) {
                    GenericEventRegistration.getGenericEventRegistration().register(_this.m_genericEventInfo.eventType, _this.m_genericEventInfo.getTargetIdFunc(), _this._callback);
                }
            });
            return new EventHandlerResult(this._context, this, handler);
        };
        GenericEventHandlers.prototype.remove = function (handler) {
            var _this = this;
            if (this._handlers.length == 1 && this.m_genericEventInfo.unregisterFunc) {
                this.m_genericEventInfo.unregisterFunc();
            }
            ActionFactory.createTraceMarkerForCallback(this._context, function () {
                var handlers = _this._handlers;
                for (var index = handlers.length - 1; index >= 0; index--) {
                    if (handlers[index] === handler) {
                        handlers.splice(index, 1);
                        break;
                    }
                }
                if (handlers.length == 0) {
                    GenericEventRegistration.getGenericEventRegistration().unregister(_this.m_genericEventInfo.eventType, _this.m_genericEventInfo.getTargetIdFunc(), _this._callback);
                }
            });
        };
        GenericEventHandlers.prototype.removeAll = function () { };
        return GenericEventHandlers;
    }(EventHandlers));
    OfficeExtension_1.GenericEventHandlers = GenericEventHandlers;
    var InstantiateActionResultHandler = (function () {
        function InstantiateActionResultHandler(clientObject) {
            this.m_clientObject = clientObject;
        }
        InstantiateActionResultHandler.prototype._handleResult = function (value) {
            this.m_clientObject._handleIdResult(value);
        };
        return InstantiateActionResultHandler;
    }());
    var ObjectPathFactory = (function () {
        function ObjectPathFactory() {
        }
        ObjectPathFactory.createGlobalObjectObjectPath = function (context) {
            var objectPathInfo = {
                Id: context._nextId(),
                ObjectPathType: 1,
                Name: ''
            };
            return new ObjectPath(objectPathInfo, null, false, false, 1, 4);
        };
        ObjectPathFactory.createNewObjectObjectPath = function (context, typeName, isCollection, flags) {
            var objectPathInfo = {
                Id: context._nextId(),
                ObjectPathType: 2,
                Name: typeName
            };
            var ret = new ObjectPath(objectPathInfo, null, isCollection, false, 1, Utility._fixupApiFlags(flags));
            return ret;
        };
        ObjectPathFactory.createPropertyObjectPath = function (context, parent, propertyName, isCollection, isInvalidAfterRequest, flags) {
            var objectPathInfo = {
                Id: context._nextId(),
                ObjectPathType: 4,
                Name: propertyName,
                ParentObjectPathId: parent._objectPath.objectPathInfo.Id
            };
            var ret = new ObjectPath(objectPathInfo, parent._objectPath, isCollection, isInvalidAfterRequest, 1, Utility._fixupApiFlags(flags));
            return ret;
        };
        ObjectPathFactory.createIndexerObjectPath = function (context, parent, args) {
            var objectPathInfo = {
                Id: context._nextId(),
                ObjectPathType: 5,
                Name: '',
                ParentObjectPathId: parent._objectPath.objectPathInfo.Id,
                ArgumentInfo: {}
            };
            objectPathInfo.ArgumentInfo.Arguments = args;
            return new ObjectPath(objectPathInfo, parent._objectPath, false, false, 1, 4);
        };
        ObjectPathFactory.createIndexerObjectPathUsingParentPath = function (context, parentObjectPath, args) {
            var objectPathInfo = {
                Id: context._nextId(),
                ObjectPathType: 5,
                Name: '',
                ParentObjectPathId: parentObjectPath.objectPathInfo.Id,
                ArgumentInfo: {}
            };
            objectPathInfo.ArgumentInfo.Arguments = args;
            return new ObjectPath(objectPathInfo, parentObjectPath, false, false, 1, 4);
        };
        ObjectPathFactory.createMethodObjectPath = function (context, parent, methodName, operationType, args, isCollection, isInvalidAfterRequest, getByIdMethodName, flags) {
            var objectPathInfo = {
                Id: context._nextId(),
                ObjectPathType: 3,
                Name: methodName,
                ParentObjectPathId: parent._objectPath.objectPathInfo.Id,
                ArgumentInfo: {}
            };
            var argumentObjectPaths = Utility.setMethodArguments(context, objectPathInfo.ArgumentInfo, args);
            var ret = new ObjectPath(objectPathInfo, parent._objectPath, isCollection, isInvalidAfterRequest, operationType, Utility._fixupApiFlags(flags));
            ret.argumentObjectPaths = argumentObjectPaths;
            ret.getByIdMethodName = getByIdMethodName;
            return ret;
        };
        ObjectPathFactory.createReferenceIdObjectPath = function (context, referenceId) {
            var objectPathInfo = {
                Id: context._nextId(),
                ObjectPathType: 6,
                Name: referenceId,
                ArgumentInfo: {}
            };
            var ret = new ObjectPath(objectPathInfo, null, false, false, 1, 4);
            return ret;
        };
        ObjectPathFactory.createChildItemObjectPathUsingIndexerOrGetItemAt = function (hasIndexerMethod, context, parent, childItem, index) {
            var id = Utility.tryGetObjectIdFromLoadOrRetrieveResult(childItem);
            if (hasIndexerMethod && !Utility.isNullOrUndefined(id)) {
                return ObjectPathFactory.createChildItemObjectPathUsingIndexer(context, parent, childItem);
            }
            else {
                return ObjectPathFactory.createChildItemObjectPathUsingGetItemAt(context, parent, childItem, index);
            }
        };
        ObjectPathFactory.createChildItemObjectPathUsingIndexer = function (context, parent, childItem) {
            var id = Utility.tryGetObjectIdFromLoadOrRetrieveResult(childItem);
            var objectPathInfo = (objectPathInfo = {
                Id: context._nextId(),
                ObjectPathType: 5,
                Name: '',
                ParentObjectPathId: parent._objectPath.objectPathInfo.Id,
                ArgumentInfo: {}
            });
            objectPathInfo.ArgumentInfo.Arguments = [id];
            return new ObjectPath(objectPathInfo, parent._objectPath, false, false, 1, 4);
        };
        ObjectPathFactory.createChildItemObjectPathUsingGetItemAt = function (context, parent, childItem, index) {
            var indexFromServer = childItem[Constants.index];
            if (indexFromServer) {
                index = indexFromServer;
            }
            var objectPathInfo = {
                Id: context._nextId(),
                ObjectPathType: 3,
                Name: Constants.getItemAt,
                ParentObjectPathId: parent._objectPath.objectPathInfo.Id,
                ArgumentInfo: {}
            };
            objectPathInfo.ArgumentInfo.Arguments = [index];
            return new ObjectPath(objectPathInfo, parent._objectPath, false, false, 1, 4);
        };
        return ObjectPathFactory;
    }());
    OfficeExtension_1.ObjectPathFactory = ObjectPathFactory;
    var OfficeJsRequestExecutor = (function () {
        function OfficeJsRequestExecutor(context) {
            this.m_context = context;
        }
        OfficeJsRequestExecutor.prototype.executeAsync = function (customData, requestFlags, requestMessage) {
            var _this = this;
            var messageSafearray = RichApiMessageUtility.buildMessageArrayForIRequestExecutor(customData, requestFlags, requestMessage, OfficeJsRequestExecutor.SourceLibHeaderValue);
            return new OfficeExtension_1.Promise(function (resolve, reject) {
                OSF.DDA.RichApi.executeRichApiRequestAsync(messageSafearray, function (result) {
                    CoreUtility.log('Response:');
                    CoreUtility.log(JSON.stringify(result));
                    var response;
                    if (result.status == 'succeeded') {
                        response = RichApiMessageUtility.buildResponseOnSuccess(RichApiMessageUtility.getResponseBody(result), RichApiMessageUtility.getResponseHeaders(result));
                    }
                    else {
                        response = RichApiMessageUtility.buildResponseOnError(result.error.code, result.error.message);
                        _this.m_context._processOfficeJsErrorResponse(result.error.code, response);
                    }
                    resolve(response);
                });
            });
        };
        OfficeJsRequestExecutor.SourceLibHeaderValue = 'officejs';
        return OfficeJsRequestExecutor;
    }());
    var TrackedObjects = (function () {
        function TrackedObjects(context) {
            this._autoCleanupList = {};
            this.m_context = context;
        }
        TrackedObjects.prototype.add = function (param) {
            var _this = this;
            if (Array.isArray(param)) {
                param.forEach(function (item) { return _this._addCommon(item, true); });
            }
            else {
                this._addCommon(param, true);
            }
        };
        TrackedObjects.prototype._autoAdd = function (object) {
            this._addCommon(object, false);
            this._autoCleanupList[object._objectPath.objectPathInfo.Id] = object;
        };
        TrackedObjects.prototype._autoTrackIfNecessaryWhenHandleObjectResultValue = function (object, resultValue) {
            var shouldAutoTrack = this.m_context._autoCleanup &&
                !object[Constants.isTracked] &&
                object !== this.m_context._rootObject &&
                resultValue &&
                !Utility.isNullOrEmptyString(resultValue[Constants.referenceId]);
            if (shouldAutoTrack) {
                this._autoCleanupList[object._objectPath.objectPathInfo.Id] = object;
                object[Constants.isTracked] = true;
            }
        };
        TrackedObjects.prototype._addCommon = function (object, isExplicitlyAdded) {
            if (object[Constants.isTracked]) {
                if (isExplicitlyAdded && this.m_context._autoCleanup) {
                    delete this._autoCleanupList[object._objectPath.objectPathInfo.Id];
                }
                return;
            }
            var referenceId = object[Constants.referenceId];
            var donotKeepReference = object._objectPath.objectPathInfo[Constants.objectPathInfoDoNotKeepReferenceFieldName];
            if (donotKeepReference) {
                throw Utility.createRuntimeError(CoreErrorCodes.generalException, CoreUtility._getResourceString(ResourceStrings.objectIsUntracked), null);
            }
            if (Utility.isNullOrEmptyString(referenceId) && object._KeepReference) {
                object._KeepReference();
                ActionFactory.createInstantiateAction(this.m_context, object);
                if (isExplicitlyAdded && this.m_context._autoCleanup) {
                    delete this._autoCleanupList[object._objectPath.objectPathInfo.Id];
                }
                object[Constants.isTracked] = true;
            }
        };
        TrackedObjects.prototype.remove = function (param) {
            var _this = this;
            if (Array.isArray(param)) {
                param.forEach(function (item) { return _this._removeCommon(item); });
            }
            else {
                this._removeCommon(param);
            }
        };
        TrackedObjects.prototype._removeCommon = function (object) {
            object._objectPath.objectPathInfo[Constants.objectPathInfoDoNotKeepReferenceFieldName] = true;
            object.context._pendingRequest._removeKeepReferenceAction(object._objectPath.objectPathInfo.Id);
            var referenceId = object[Constants.referenceId];
            if (!Utility.isNullOrEmptyString(referenceId)) {
                var rootObject = this.m_context._rootObject;
                if (rootObject._RemoveReference) {
                    rootObject._RemoveReference(referenceId);
                }
            }
            delete object[Constants.isTracked];
        };
        TrackedObjects.prototype._retrieveAndClearAutoCleanupList = function () {
            var list = this._autoCleanupList;
            this._autoCleanupList = {};
            return list;
        };
        return TrackedObjects;
    }());
    OfficeExtension_1.TrackedObjects = TrackedObjects;
    var RequestPrettyPrinter = (function () {
        function RequestPrettyPrinter(globalObjName, referencedObjectPaths, actions, showDispose, removePII) {
            if (!globalObjName) {
                globalObjName = 'root';
            }
            this.m_globalObjName = globalObjName;
            this.m_referencedObjectPaths = referencedObjectPaths;
            this.m_actions = actions;
            this.m_statements = [];
            this.m_variableNameForObjectPathMap = {};
            this.m_variableNameToObjectPathMap = {};
            this.m_declaredObjectPathMap = {};
            this.m_showDispose = showDispose;
            this.m_removePII = removePII;
        }
        RequestPrettyPrinter.prototype.process = function () {
            if (this.m_showDispose) {
                ClientRequest._calculateLastUsedObjectPathIds(this.m_actions);
            }
            for (var i = 0; i < this.m_actions.length; i++) {
                this.processOneAction(this.m_actions[i]);
            }
            return this.m_statements;
        };
        RequestPrettyPrinter.prototype.processForDebugStatementInfo = function (actionIndex) {
            if (this.m_showDispose) {
                ClientRequest._calculateLastUsedObjectPathIds(this.m_actions);
            }
            var surroundingCount = 5;
            this.m_statements = [];
            var oneStatement = '';
            var statementIndex = -1;
            for (var i = 0; i < this.m_actions.length; i++) {
                this.processOneAction(this.m_actions[i]);
                if (actionIndex == i) {
                    statementIndex = this.m_statements.length - 1;
                }
                if (statementIndex >= 0 && this.m_statements.length > statementIndex + surroundingCount + 1) {
                    break;
                }
            }
            if (statementIndex < 0) {
                return null;
            }
            var startIndex = statementIndex - surroundingCount;
            if (startIndex < 0) {
                startIndex = 0;
            }
            var endIndex = statementIndex + 1 + surroundingCount;
            if (endIndex > this.m_statements.length) {
                endIndex = this.m_statements.length;
            }
            var surroundingStatements = [];
            if (startIndex != 0) {
                surroundingStatements.push('...');
            }
            for (var i_1 = startIndex; i_1 < statementIndex; i_1++) {
                surroundingStatements.push(this.m_statements[i_1]);
            }
            surroundingStatements.push('// >>>>>');
            surroundingStatements.push(this.m_statements[statementIndex]);
            surroundingStatements.push('// <<<<<');
            for (var i_2 = statementIndex + 1; i_2 < endIndex; i_2++) {
                surroundingStatements.push(this.m_statements[i_2]);
            }
            if (endIndex < this.m_statements.length) {
                surroundingStatements.push('...');
            }
            return {
                statement: this.m_statements[statementIndex],
                surroundingStatements: surroundingStatements
            };
        };
        RequestPrettyPrinter.prototype.processOneAction = function (action) {
            var actionInfo = action.actionInfo;
            switch (actionInfo.ActionType) {
                case 1:
                    this.processInstantiateAction(action);
                    break;
                case 3:
                    this.processMethodAction(action);
                    break;
                case 2:
                    this.processQueryAction(action);
                    break;
                case 7:
                    this.processQueryAsJsonAction(action);
                    break;
                case 6:
                    this.processRecursiveQueryAction(action);
                    break;
                case 4:
                    this.processSetPropertyAction(action);
                    break;
                case 5:
                    this.processTraceAction(action);
                    break;
                case 8:
                    this.processEnsureUnchangedAction(action);
                    break;
                case 9:
                    this.processUpdateAction(action);
                    break;
            }
        };
        RequestPrettyPrinter.prototype.processInstantiateAction = function (action) {
            var objId = action.actionInfo.ObjectPathId;
            var objPath = this.m_referencedObjectPaths[objId];
            var varName = this.getObjVarName(objId);
            if (!this.m_declaredObjectPathMap[objId]) {
                var statement = 'var ' + varName + ' = ' + this.buildObjectPathExpressionWithParent(objPath) + ';';
                statement = this.appendDisposeCommentIfRelevant(statement, action);
                this.m_statements.push(statement);
                this.m_declaredObjectPathMap[objId] = varName;
            }
            else {
                var statement = '// Instantiate {' + varName + '}';
                statement = this.appendDisposeCommentIfRelevant(statement, action);
                this.m_statements.push(statement);
            }
        };
        RequestPrettyPrinter.prototype.processMethodAction = function (action) {
            var methodName = action.actionInfo.Name;
            if (methodName === '_KeepReference') {
                if (!OfficeExtension_1._internalConfig.showInternalApiInDebugInfo) {
                    return;
                }
                methodName = 'track';
            }
            var statement = this.getObjVarName(action.actionInfo.ObjectPathId) +
                '.' +
                Utility._toCamelLowerCase(methodName) +
                '(' +
                this.buildArgumentsExpression(action.actionInfo.ArgumentInfo) +
                ');';
            statement = this.appendDisposeCommentIfRelevant(statement, action);
            this.m_statements.push(statement);
        };
        RequestPrettyPrinter.prototype.processQueryAction = function (action) {
            var queryExp = this.buildQueryExpression(action);
            var statement = this.getObjVarName(action.actionInfo.ObjectPathId) + '.load(' + queryExp + ');';
            statement = this.appendDisposeCommentIfRelevant(statement, action);
            this.m_statements.push(statement);
        };
        RequestPrettyPrinter.prototype.processQueryAsJsonAction = function (action) {
            var queryExp = this.buildQueryExpression(action);
            var statement = this.getObjVarName(action.actionInfo.ObjectPathId) + '.retrieve(' + queryExp + ');';
            statement = this.appendDisposeCommentIfRelevant(statement, action);
            this.m_statements.push(statement);
        };
        RequestPrettyPrinter.prototype.processRecursiveQueryAction = function (action) {
            var queryExp = '';
            if (action.actionInfo.RecursiveQueryInfo) {
                queryExp = JSON.stringify(action.actionInfo.RecursiveQueryInfo);
            }
            var statement = this.getObjVarName(action.actionInfo.ObjectPathId) + '.loadRecursive(' + queryExp + ');';
            statement = this.appendDisposeCommentIfRelevant(statement, action);
            this.m_statements.push(statement);
        };
        RequestPrettyPrinter.prototype.processSetPropertyAction = function (action) {
            var statement = this.getObjVarName(action.actionInfo.ObjectPathId) +
                '.' +
                Utility._toCamelLowerCase(action.actionInfo.Name) +
                ' = ' +
                this.buildArgumentsExpression(action.actionInfo.ArgumentInfo) +
                ';';
            statement = this.appendDisposeCommentIfRelevant(statement, action);
            this.m_statements.push(statement);
        };
        RequestPrettyPrinter.prototype.processTraceAction = function (action) {
            var statement = 'context.trace();';
            statement = this.appendDisposeCommentIfRelevant(statement, action);
            this.m_statements.push(statement);
        };
        RequestPrettyPrinter.prototype.processEnsureUnchangedAction = function (action) {
            var statement = this.getObjVarName(action.actionInfo.ObjectPathId) +
                '.ensureUnchanged(' +
                JSON.stringify(action.actionInfo.ObjectState) +
                ');';
            statement = this.appendDisposeCommentIfRelevant(statement, action);
            this.m_statements.push(statement);
        };
        RequestPrettyPrinter.prototype.processUpdateAction = function (action) {
            var statement = this.getObjVarName(action.actionInfo.ObjectPathId) +
                '.update(' +
                JSON.stringify(action.actionInfo.ObjectState) +
                ');';
            statement = this.appendDisposeCommentIfRelevant(statement, action);
            this.m_statements.push(statement);
        };
        RequestPrettyPrinter.prototype.appendDisposeCommentIfRelevant = function (statement, action) {
            var _this = this;
            if (this.m_showDispose) {
                var lastUsedObjectPathIds = action.actionInfo.L;
                if (lastUsedObjectPathIds && lastUsedObjectPathIds.length > 0) {
                    var objectNamesToDispose = lastUsedObjectPathIds.map(function (item) { return _this.getObjVarName(item); }).join(', ');
                    return statement + ' // And then dispose {' + objectNamesToDispose + '}';
                }
            }
            return statement;
        };
        RequestPrettyPrinter.prototype.buildQueryExpression = function (action) {
            if (action.actionInfo.QueryInfo) {
                var option = {};
                option.select = action.actionInfo.QueryInfo.Select;
                option.expand = action.actionInfo.QueryInfo.Expand;
                option.skip = action.actionInfo.QueryInfo.Skip;
                option.top = action.actionInfo.QueryInfo.Top;
                if (typeof option.top === 'undefined' &&
                    typeof option.skip === 'undefined' &&
                    typeof option.expand === 'undefined') {
                    if (typeof option.select === 'undefined') {
                        return '';
                    }
                    else {
                        return JSON.stringify(option.select);
                    }
                }
                else {
                    return JSON.stringify(option);
                }
            }
            return '';
        };
        RequestPrettyPrinter.prototype.buildObjectPathExpressionWithParent = function (objPath) {
            var hasParent = objPath.objectPathInfo.ObjectPathType == 5 ||
                objPath.objectPathInfo.ObjectPathType == 3 ||
                objPath.objectPathInfo.ObjectPathType == 4;
            if (hasParent && objPath.objectPathInfo.ParentObjectPathId) {
                return (this.getObjVarName(objPath.objectPathInfo.ParentObjectPathId) + '.' + this.buildObjectPathExpression(objPath));
            }
            return this.buildObjectPathExpression(objPath);
        };
        RequestPrettyPrinter.prototype.buildObjectPathExpression = function (objPath) {
            var expr = this.buildObjectPathInfoExpression(objPath.objectPathInfo);
            var originalObjectPathInfo = objPath.originalObjectPathInfo;
            if (originalObjectPathInfo) {
                expr = expr + ' /* originally ' + this.buildObjectPathInfoExpression(originalObjectPathInfo) + ' */';
            }
            return expr;
        };
        RequestPrettyPrinter.prototype.buildObjectPathInfoExpression = function (objectPathInfo) {
            switch (objectPathInfo.ObjectPathType) {
                case 1:
                    return 'context.' + this.m_globalObjName;
                case 5:
                    return 'getItem(' + this.buildArgumentsExpression(objectPathInfo.ArgumentInfo) + ')';
                case 3:
                    return (Utility._toCamelLowerCase(objectPathInfo.Name) +
                        '(' +
                        this.buildArgumentsExpression(objectPathInfo.ArgumentInfo) +
                        ')');
                case 2:
                    return objectPathInfo.Name + '.newObject()';
                case 7:
                    return 'null';
                case 4:
                    return Utility._toCamelLowerCase(objectPathInfo.Name);
                case 6:
                    return ('context.' + this.m_globalObjName + '._getObjectByReferenceId(' + JSON.stringify(objectPathInfo.Name) + ')');
            }
        };
        RequestPrettyPrinter.prototype.buildArgumentsExpression = function (args) {
            var ret = '';
            if (!args.Arguments || args.Arguments.length === 0) {
                return ret;
            }
            if (this.m_removePII) {
                if (typeof args.Arguments[0] === 'undefined') {
                    return ret;
                }
                return '...';
            }
            for (var i = 0; i < args.Arguments.length; i++) {
                if (i > 0) {
                    ret = ret + ', ';
                }
                ret =
                    ret +
                        this.buildArgumentLiteral(args.Arguments[i], args.ReferencedObjectPathIds ? args.ReferencedObjectPathIds[i] : null);
            }
            if (ret === 'undefined') {
                ret = '';
            }
            return ret;
        };
        RequestPrettyPrinter.prototype.buildArgumentLiteral = function (value, objectPathId) {
            if (typeof value == 'number' && value === objectPathId) {
                return this.getObjVarName(objectPathId);
            }
            else {
                return JSON.stringify(value);
            }
        };
        RequestPrettyPrinter.prototype.getObjVarNameBase = function (objectPathId) {
            var ret = 'v';
            var objPath = this.m_referencedObjectPaths[objectPathId];
            if (objPath) {
                switch (objPath.objectPathInfo.ObjectPathType) {
                    case 1:
                        ret = this.m_globalObjName;
                        break;
                    case 4:
                        ret = Utility._toCamelLowerCase(objPath.objectPathInfo.Name);
                        break;
                    case 3:
                        var methodName = objPath.objectPathInfo.Name;
                        if (methodName.length > 3 && methodName.substr(0, 3) === 'Get') {
                            methodName = methodName.substr(3);
                        }
                        ret = Utility._toCamelLowerCase(methodName);
                        break;
                    case 5:
                        var parentName = this.getObjVarNameBase(objPath.objectPathInfo.ParentObjectPathId);
                        if (parentName.charAt(parentName.length - 1) === 's') {
                            ret = parentName.substr(0, parentName.length - 1);
                        }
                        else {
                            ret = parentName + 'Item';
                        }
                        break;
                }
            }
            return ret;
        };
        RequestPrettyPrinter.prototype.getObjVarName = function (objectPathId) {
            if (this.m_variableNameForObjectPathMap[objectPathId]) {
                return this.m_variableNameForObjectPathMap[objectPathId];
            }
            var ret = this.getObjVarNameBase(objectPathId);
            if (!this.m_variableNameToObjectPathMap[ret]) {
                this.m_variableNameForObjectPathMap[objectPathId] = ret;
                this.m_variableNameToObjectPathMap[ret] = objectPathId;
                return ret;
            }
            var i = 1;
            while (this.m_variableNameToObjectPathMap[ret + i.toString()]) {
                i++;
            }
            ret = ret + i.toString();
            this.m_variableNameForObjectPathMap[objectPathId] = ret;
            this.m_variableNameToObjectPathMap[ret] = objectPathId;
            return ret;
        };
        return RequestPrettyPrinter;
    }());
    var ResourceStrings = (function (_super) {
        __extends(ResourceStrings, _super);
        function ResourceStrings() {
            return _super !== null && _super.apply(this, arguments) || this;
        }
        ResourceStrings.cannotRegisterEvent = 'CannotRegisterEvent';
        ResourceStrings.connectionFailureWithStatus = 'ConnectionFailureWithStatus';
        ResourceStrings.connectionFailureWithDetails = 'ConnectionFailureWithDetails';
        ResourceStrings.propertyNotLoaded = 'PropertyNotLoaded';
        ResourceStrings.runMustReturnPromise = 'RunMustReturnPromise';
        ResourceStrings.moreInfoInnerError = 'MoreInfoInnerError';
        ResourceStrings.cannotApplyPropertyThroughSetMethod = 'CannotApplyPropertyThroughSetMethod';
        ResourceStrings.invalidOperationInCellEditMode = 'InvalidOperationInCellEditMode';
        ResourceStrings.objectIsUntracked = 'ObjectIsUntracked';
        ResourceStrings.customFunctionDefintionMissing = 'CustomFunctionDefintionMissing';
        ResourceStrings.customFunctionImplementationMissing = 'CustomFunctionImplementationMissing';
        ResourceStrings.customFunctionNameContainsBadChars = 'CustomFunctionNameContainsBadChars';
        ResourceStrings.customFunctionNameCannotSplit = 'CustomFunctionNameCannotSplit';
        ResourceStrings.customFunctionUnexpectedNumberOfEntriesInResultBatch = 'CustomFunctionUnexpectedNumberOfEntriesInResultBatch';
        ResourceStrings.customFunctionCancellationHandlerMissing = 'CustomFunctionCancellationHandlerMissing';
        ResourceStrings.customFunctionInvalidFunction = 'CustomFunctionInvalidFunction';
        ResourceStrings.customFunctionInvalidFunctionMapping = 'CustomFunctionInvalidFunctionMapping';
        ResourceStrings.customFunctionWindowMissing = 'CustomFunctionWindowMissing';
        ResourceStrings.customFunctionDefintionMissingOnWindow = 'CustomFunctionDefintionMissingOnWindow';
        ResourceStrings.pendingBatchInProgress = 'PendingBatchInProgress';
        ResourceStrings.notInsideBatch = 'NotInsideBatch';
        ResourceStrings.cannotUpdateReadOnlyProperty = 'CannotUpdateReadOnlyProperty';
        return ResourceStrings;
    }(CommonResourceStrings));
    OfficeExtension_1.ResourceStrings = ResourceStrings;
    CoreUtility.addResourceStringValues({
        CannotRegisterEvent: 'The event handler cannot be registered.',
        PropertyNotLoaded: "The property '{0}' is not available. Before reading the property's value, call the load method on the containing object and call \"context.sync()\" on the associated request context.",
        RunMustReturnPromise: 'The batch function passed to the ".run" method didn\'t return a promise. The function must return a promise, so that any automatically-tracked objects can be released at the completion of the batch operation. Typically, you return a promise by returning the response from "context.sync()".',
        InvalidOrTimedOutSessionMessage: 'Your Office Online session has expired or is invalid. To continue, refresh the page.',
        InvalidOperationInCellEditMode: 'Excel is in cell-editing mode. Please exit the edit mode by pressing ENTER or TAB or selecting another cell, and then try again.',
        CustomFunctionDefintionMissing: "A property with the name '{0}' that represents the function's definition must exist on Excel.Script.CustomFunctions.",
        CustomFunctionDefintionMissingOnWindow: "A property with the name '{0}' that represents the function's definition must exist on the window object.",
        CustomFunctionImplementationMissing: "The property with the name '{0}' on Excel.Script.CustomFunctions that represents the function's definition must contain a 'call' property that implements the function.",
        CustomFunctionNameContainsBadChars: 'The function name may only contain letters, digits, underscores, and periods.',
        CustomFunctionNameCannotSplit: 'The function name must contain a non-empty namespace and a non-empty short name.',
        CustomFunctionUnexpectedNumberOfEntriesInResultBatch: "The batching function returned a number of results that doesn't match the number of parameter value sets that were passed into it.",
        CustomFunctionCancellationHandlerMissing: 'The cancellation handler onCanceled is missing in the function. The handler must be present as the function is defined as cancelable.',
        CustomFunctionInvalidFunction: "The property with the name '{0}' that represents the function's definition is not a valid function.",
        CustomFunctionInvalidFunctionMapping: "The property with the name '{0}' on CustomFunctionMappings that represents the function's definition is not a valid function.",
        CustomFunctionWindowMissing: 'The window object was not found.',
        PendingBatchInProgress: 'There is a pending batch in progress. The batch method may not be called inside another batch, or simultaneously with another batch.',
        NotInsideBatch: 'Operations may not be invoked outside of a batch method.',
        CannotUpdateReadOnlyProperty: "The property '{0}' is read-only and it cannot be updated.",
        ObjectIsUntracked: 'The object is untracked.'
    });
    var Utility = (function (_super) {
        __extends(Utility, _super);
        function Utility() {
            return _super !== null && _super.apply(this, arguments) || this;
        }
        Utility.fixObjectPathIfNecessary = function (clientObject, value) {
            if (clientObject && clientObject._objectPath && value) {
                clientObject._objectPath.updateUsingObjectData(value, clientObject);
            }
        };
        Utility.load = function (clientObj, option) {
            clientObj.context.load(clientObj, option);
            return clientObj;
        };
        Utility.loadAndSync = function (clientObj, option) {
            clientObj.context.load(clientObj, option);
            return clientObj.context.sync().then(function () { return clientObj; });
        };
        Utility.retrieve = function (clientObj, option) {
            var shouldPolyfill = OfficeExtension_1._internalConfig.alwaysPolyfillClientObjectRetrieveMethod;
            if (!shouldPolyfill) {
                shouldPolyfill = !Utility.isSetSupported('RichApiRuntime', '1.1');
            }
            var result = new RetrieveResultImpl(clientObj, shouldPolyfill);
            clientObj._retrieve(option, result);
            return result;
        };
        Utility.retrieveAndSync = function (clientObj, option) {
            var result = Utility.retrieve(clientObj, option);
            return clientObj.context.sync().then(function () { return result; });
        };
        Utility.toJson = function (clientObj, scalarProperties, navigationProperties, collectionItemsIfAny) {
            var result = {};
            for (var prop in scalarProperties) {
                var value = scalarProperties[prop];
                if (typeof value !== 'undefined') {
                    result[prop] = value;
                }
            }
            for (var prop in navigationProperties) {
                var value = navigationProperties[prop];
                if (typeof value !== 'undefined') {
                    if (value[Utility.fieldName_isCollection] && typeof value[Utility.fieldName_m__items] !== 'undefined') {
                        result[prop] = value.toJSON()['items'];
                    }
                    else {
                        result[prop] = value.toJSON();
                    }
                }
            }
            if (collectionItemsIfAny) {
                result['items'] = collectionItemsIfAny.map(function (item) { return item.toJSON(); });
            }
            return result;
        };
        Utility.throwError = function (resourceId, arg, errorLocation) {
            throw new _Internal.RuntimeError({
                code: resourceId,
                message: CoreUtility._getResourceString(resourceId, arg),
                debugInfo: errorLocation ? { errorLocation: errorLocation } : undefined
            });
        };
        Utility.createRuntimeError = function (code, message, location) {
            return new _Internal.RuntimeError({
                code: code,
                message: message,
                debugInfo: { errorLocation: location }
            });
        };
        Utility.throwIfNotLoaded = function (propertyName, fieldValue, entityName, isNull) {
            if (!isNull &&
                CoreUtility.isUndefined(fieldValue) &&
                propertyName.charCodeAt(0) != Utility.s_underscoreCharCode) {
                throw Utility.createPropertyNotLoadedException(entityName, propertyName);
            }
        };
        Utility.createPropertyNotLoadedException = function (entityName, propertyName) {
            return new _Internal.RuntimeError({
                code: ErrorCodes.propertyNotLoaded,
                message: CoreUtility._getResourceString(ResourceStrings.propertyNotLoaded, propertyName),
                debugInfo: entityName ? { errorLocation: entityName + '.' + propertyName } : undefined
            });
        };
        Utility.createCannotUpdateReadOnlyPropertyException = function (entityName, propertyName) {
            return new _Internal.RuntimeError({
                code: ErrorCodes.cannotUpdateReadOnlyProperty,
                message: CoreUtility._getResourceString(ResourceStrings.cannotUpdateReadOnlyProperty, propertyName),
                debugInfo: entityName ? { errorLocation: entityName + '.' + propertyName } : undefined
            });
        };
        Utility.promisify = function (action) {
            return new OfficeExtension_1.Promise(function (resolve, reject) {
                var callback = function (result) {
                    if (result.status == 'failed') {
                        reject(result.error);
                    }
                    else {
                        resolve(result.value);
                    }
                };
                action(callback);
            });
        };
        Utility._addActionResultHandler = function (clientObj, action, resultHandler) {
            clientObj.context._pendingRequest.addActionResultHandler(action, resultHandler);
        };
        Utility._handleNavigationPropertyResults = function (clientObj, objectValue, propertyNames) {
            for (var i = 0; i < propertyNames.length - 1; i += 2) {
                if (!CoreUtility.isUndefined(objectValue[propertyNames[i + 1]])) {
                    clientObj[propertyNames[i]]._handleResult(objectValue[propertyNames[i + 1]]);
                }
            }
        };
        Utility._fixupApiFlags = function (flags) {
            if (typeof flags === 'boolean') {
                if (flags) {
                    flags = 1;
                }
                else {
                    flags = 0;
                }
            }
            return flags;
        };
        Utility.definePropertyThrowUnloadedException = function (obj, typeName, propertyName) {
            Object.defineProperty(obj, propertyName, {
                configurable: true,
                enumerable: true,
                get: function () {
                    throw Utility.createPropertyNotLoadedException(typeName, propertyName);
                },
                set: function () {
                    throw Utility.createCannotUpdateReadOnlyPropertyException(typeName, propertyName);
                }
            });
        };
        Utility.defineReadOnlyPropertyWithValue = function (obj, propertyName, value) {
            Object.defineProperty(obj, propertyName, {
                configurable: true,
                enumerable: true,
                get: function () {
                    return value;
                },
                set: function () {
                    throw Utility.createCannotUpdateReadOnlyPropertyException(null, propertyName);
                }
            });
        };
        Utility.processRetrieveResult = function (proxy, value, result, childItemCreateFunc) {
            if (CoreUtility.isNullOrUndefined(value)) {
                return;
            }
            if (childItemCreateFunc) {
                var data = value[Constants.itemsLowerCase];
                if (Array.isArray(data)) {
                    var itemsResult = [];
                    for (var i = 0; i < data.length; i++) {
                        var itemProxy = childItemCreateFunc(data[i], i);
                        var itemResult = {};
                        itemResult[Constants.proxy] = itemProxy;
                        itemProxy._handleRetrieveResult(data[i], itemResult);
                        itemsResult.push(itemResult);
                    }
                    Utility.defineReadOnlyPropertyWithValue(result, Constants.itemsLowerCase, itemsResult);
                }
            }
            else {
                var scalarPropertyNames = proxy[Constants.scalarPropertyNames];
                var navigationPropertyNames = proxy[Constants.navigationPropertyNames];
                var typeName = proxy[Constants.className];
                if (scalarPropertyNames) {
                    for (var i = 0; i < scalarPropertyNames.length; i++) {
                        var propName = scalarPropertyNames[i];
                        var propValue = value[propName];
                        if (CoreUtility.isUndefined(propValue)) {
                            Utility.definePropertyThrowUnloadedException(result, typeName, propName);
                        }
                        else {
                            Utility.defineReadOnlyPropertyWithValue(result, propName, propValue);
                        }
                    }
                }
                if (navigationPropertyNames) {
                    for (var i = 0; i < navigationPropertyNames.length; i++) {
                        var propName = navigationPropertyNames[i];
                        var propValue = value[propName];
                        if (CoreUtility.isUndefined(propValue)) {
                            Utility.definePropertyThrowUnloadedException(result, typeName, propName);
                        }
                        else {
                            var propProxy = proxy[propName];
                            var propResult = {};
                            propProxy._handleRetrieveResult(propValue, propResult);
                            propResult[Constants.proxy] = propProxy;
                            if (Array.isArray(propResult[Constants.itemsLowerCase])) {
                                propResult = propResult[Constants.itemsLowerCase];
                            }
                            Utility.defineReadOnlyPropertyWithValue(result, propName, propResult);
                        }
                    }
                }
            }
        };
        Utility.setMockData = function (clientObj, value, childItemCreateFunc, setItemsFunc) {
            if (CoreUtility.isNullOrUndefined(value)) {
                clientObj._handleResult(value);
                return;
            }
            if (clientObj[Constants.scalarPropertyOriginalNames]) {
                var result = {};
                var scalarPropertyOriginalNames = clientObj[Constants.scalarPropertyOriginalNames];
                var scalarPropertyNames = clientObj[Constants.scalarPropertyNames];
                for (var i = 0; i < scalarPropertyNames.length; i++) {
                    if (typeof (value[scalarPropertyNames[i]]) !== 'undefined') {
                        result[scalarPropertyOriginalNames[i]] = value[scalarPropertyNames[i]];
                    }
                }
                clientObj._handleResult(result);
            }
            if (clientObj[Constants.navigationPropertyNames]) {
                var navigationPropertyNames = clientObj[Constants.navigationPropertyNames];
                for (var i = 0; i < navigationPropertyNames.length; i++) {
                    if (typeof (value[navigationPropertyNames[i]]) !== 'undefined') {
                        var navigationPropValue = clientObj[navigationPropertyNames[i]];
                        if (navigationPropValue.setMockData) {
                            navigationPropValue.setMockData(value[navigationPropertyNames[i]]);
                        }
                    }
                }
            }
            if (clientObj[Constants.isCollection] && childItemCreateFunc) {
                var itemsData = Array.isArray(value) ? value : value[Constants.itemsLowerCase];
                if (Array.isArray(itemsData)) {
                    var items = [];
                    for (var i = 0; i < itemsData.length; i++) {
                        var item = childItemCreateFunc(itemsData, i);
                        Utility.setMockData(item, itemsData[i]);
                        items.push(item);
                    }
                    setItemsFunc(items);
                }
            }
        };
        Utility.applyMixin = function (derived, base) {
            Object.getOwnPropertyNames(base.prototype).forEach(function (name) {
                if (name !== 'constructor') {
                    Object.defineProperty(derived.prototype, name, Object.getOwnPropertyDescriptor(base.prototype, name));
                }
            });
        };
        Utility.fieldName_m__items = 'm__items';
        Utility.fieldName_isCollection = '_isCollection';
        Utility._synchronousCleanup = false;
        Utility.s_underscoreCharCode = '_'.charCodeAt(0);
        return Utility;
    }(CommonUtility));
    OfficeExtension_1.Utility = Utility;
    var BatchApiHelper = (function () {
        function BatchApiHelper() {
        }
        BatchApiHelper.invokeMethod = function (obj, methodName, operationType, args, flags, resultProcessType) {
            var action = ActionFactory.createMethodAction(obj.context, obj, methodName, operationType, args, flags);
            var result = new ClientResult(resultProcessType);
            Utility._addActionResultHandler(obj, action, result);
            return result;
        };
        BatchApiHelper.invokeEnsureUnchanged = function (obj, objectState) {
            ActionFactory.createEnsureUnchangedAction(obj.context, obj, objectState);
        };
        BatchApiHelper.invokeSetProperty = function (obj, propName, propValue, flags) {
            ActionFactory.createSetPropertyAction(obj.context, obj, propName, propValue, flags);
        };
        BatchApiHelper.createRootServiceObject = function (type, context) {
            var objectPath = ObjectPathFactory.createGlobalObjectObjectPath(context);
            return new type(context, objectPath);
        };
        BatchApiHelper.createObjectFromReferenceId = function (type, context, referenceId) {
            var objectPath = ObjectPathFactory.createReferenceIdObjectPath(context, referenceId);
            return new type(context, objectPath);
        };
        BatchApiHelper.createTopLevelServiceObject = function (type, context, typeName, isCollection, flags) {
            var objectPath = ObjectPathFactory.createNewObjectObjectPath(context, typeName, isCollection, flags);
            return new type(context, objectPath);
        };
        BatchApiHelper.createPropertyObject = function (type, parent, propertyName, isCollection, flags) {
            var objectPath = ObjectPathFactory.createPropertyObjectPath(parent.context, parent, propertyName, isCollection, false, flags);
            return new type(parent.context, objectPath);
        };
        BatchApiHelper.createIndexerObject = function (type, parent, args) {
            var objectPath = ObjectPathFactory.createIndexerObjectPath(parent.context, parent, args);
            return new type(parent.context, objectPath);
        };
        BatchApiHelper.createMethodObject = function (type, parent, methodName, operationType, args, isCollection, isInvalidAfterRequest, getByIdMethodName, flags) {
            var objectPath = ObjectPathFactory.createMethodObjectPath(parent.context, parent, methodName, operationType, args, isCollection, isInvalidAfterRequest, getByIdMethodName, flags);
            return new type(parent.context, objectPath);
        };
        BatchApiHelper.createChildItemObject = function (type, hasIndexerMethod, parent, chileItem, index) {
            var objectPath = ObjectPathFactory.createChildItemObjectPathUsingIndexerOrGetItemAt(hasIndexerMethod, parent.context, parent, chileItem, index);
            return new type(parent.context, objectPath);
        };
        return BatchApiHelper;
    }());
    OfficeExtension_1.BatchApiHelper = BatchApiHelper;
    var LibraryBuilder = (function () {
        function LibraryBuilder(options) {
            this.m_namespaceMap = {};
            this.m_namespace = options.metadata.name;
            this.m_targetNamespaceObject = options.targetNamespaceObject;
            this.m_namespaceMap[this.m_namespace] = options.targetNamespaceObject;
            if (options.namespaceMap) {
                for (var ns in options.namespaceMap) {
                    this.m_namespaceMap[ns] = options.namespaceMap[ns];
                }
            }
            this.m_defaultApiSetName = options.metadata.defaultApiSetName;
            this.m_hostName = options.metadata.hostName;
            var metadata = options.metadata;
            if (metadata.enumTypes) {
                for (var i = 0; i < metadata.enumTypes.length; i++) {
                    this.buildEnumType(metadata.enumTypes[i]);
                }
            }
            if (metadata.apiSets) {
                for (var i = 0; i < metadata.apiSets.length; i++) {
                    var elem = metadata.apiSets[i];
                    if (Array.isArray(elem)) {
                        metadata.apiSets[i] = {
                            version: elem[0],
                            name: elem[1] || this.m_defaultApiSetName
                        };
                    }
                }
                this.m_apiSets = metadata.apiSets;
            }
            this.m_strings = metadata.strings;
            if (metadata.clientObjectTypes) {
                for (var i = 0; i < metadata.clientObjectTypes.length; i++) {
                    var elem = metadata.clientObjectTypes[i];
                    if (Array.isArray(elem)) {
                        this.ensureArraySize(elem, 11);
                        metadata.clientObjectTypes[i] = {
                            name: this.getString(elem[0]),
                            behaviorFlags: elem[1],
                            collectionPropertyPath: this.getString(elem[6]),
                            newObjectServerTypeFullName: this.getString(elem[9]),
                            newObjectApiFlags: elem[10],
                            childItemTypeFullName: this.getString(elem[7]),
                            scalarProperties: elem[2],
                            navigationProperties: elem[3],
                            scalarMethods: elem[4],
                            navigationMethods: elem[5],
                            events: elem[8],
                        };
                    }
                    this.buildClientObjectType(metadata.clientObjectTypes[i], options.fullyInitialize);
                }
            }
        }
        LibraryBuilder.prototype.ensureArraySize = function (value, size) {
            var count = size - value.length;
            while (count > 0) {
                value.push(0);
                count--;
            }
        };
        LibraryBuilder.prototype.getString = function (ordinalOrValue) {
            if (typeof (ordinalOrValue) === "number") {
                if (ordinalOrValue > 0) {
                    return this.m_strings[ordinalOrValue - 1];
                }
                return null;
            }
            return ordinalOrValue;
        };
        LibraryBuilder.prototype.buildEnumType = function (elem) {
            var enumType;
            if (Array.isArray(elem)) {
                enumType = {
                    name: elem[0],
                    fields: elem[2]
                };
                if (!enumType.fields) {
                    enumType.fields = {};
                }
                var fieldsWithCamelUpperCaseValue = elem[1];
                if (Array.isArray(fieldsWithCamelUpperCaseValue)) {
                    for (var index = 0; index < fieldsWithCamelUpperCaseValue.length; index++) {
                        enumType.fields[fieldsWithCamelUpperCaseValue[index]] = this.toSimpleCamelUpperCase(fieldsWithCamelUpperCaseValue[index]);
                    }
                }
            }
            else {
                enumType = elem;
            }
            this.m_targetNamespaceObject[enumType.name] = enumType.fields;
        };
        LibraryBuilder.prototype.buildClientObjectType = function (typeInfo, fullyInitialize) {
            var thisBuilder = this;
            var type = function (context, objectPath) {
                ClientObject.apply(this, arguments);
                if (!thisBuilder.m_targetNamespaceObject[typeInfo.name]._typeInited) {
                    thisBuilder.buildPrototype(thisBuilder.m_targetNamespaceObject[typeInfo.name], typeInfo);
                    thisBuilder.m_targetNamespaceObject[typeInfo.name]._typeInited = true;
                }
            };
            this.m_targetNamespaceObject[typeInfo.name] = type;
            this.extendsType(type, ClientObject);
            this.buildNewObject(type, typeInfo);
            if ((typeInfo.behaviorFlags & 2) !== 0) {
                type.prototype._KeepReference = function () {
                    BatchApiHelper.invokeMethod(this, "_KeepReference", 1, [], 0, 0);
                };
            }
            if ((typeInfo.behaviorFlags & 32) !== 0) {
                var func = this.getFunction(LibraryBuilder.CustomizationCodeNamespace + "." + typeInfo.name + "_StaticCustomize");
                func.call(null, type);
            }
            if (fullyInitialize) {
                this.buildPrototype(type, typeInfo);
                type._typeInited = true;
            }
        };
        LibraryBuilder.prototype.extendsType = function (d, b) {
            function __() { this.constructor = d; }
            d.prototype = b === null ? Object.create(b) : (__.prototype = b.prototype, new __());
        };
        LibraryBuilder.prototype.findObjectUnderPath = function (top, paths, pathStartIndex) {
            var obj = top;
            for (var i = pathStartIndex; i < paths.length; i++) {
                if (typeof (obj) !== 'object') {
                    throw new OfficeExtension_1.Error("Cannot find " + paths.join("."));
                }
                obj = obj[paths[i]];
            }
            return obj;
        };
        LibraryBuilder.prototype.getFunction = function (fullName) {
            var ret = this.resolveObjectByFullName(fullName);
            if (typeof (ret) !== 'function') {
                throw new OfficeExtension_1.Error("Cannot find function or type: " + fullName);
            }
            return ret;
        };
        LibraryBuilder.prototype.resolveObjectByFullName = function (fullName) {
            var parts = fullName.split('.');
            if (parts.length === 1) {
                return this.m_targetNamespaceObject[parts[0]];
            }
            var rootName = parts[0];
            if (rootName === this.m_namespace) {
                return this.findObjectUnderPath(this.m_targetNamespaceObject, parts, 1);
            }
            if (this.m_namespaceMap[rootName]) {
                return this.findObjectUnderPath(this.m_namespaceMap[rootName], parts, 1);
            }
            return this.findObjectUnderPath(this.m_targetNamespaceObject, parts, 0);
        };
        LibraryBuilder.prototype.evaluateExpression = function (expression, thisObj) {
            if (Utility.isNullOrUndefined(expression)) {
                return null;
            }
            var result = new Function('return ' + expression).call(thisObj);
            return result;
        };
        LibraryBuilder.prototype.evaluateEventTargetId = function (targetIdExpression, thisObj) {
            if (Utility.isNullOrEmptyString(targetIdExpression)) {
                return '';
            }
            return this.evaluateExpression(targetIdExpression, thisObj);
        };
        LibraryBuilder.prototype.isAllDigits = function (expression) {
            var charZero = '0'.charCodeAt(0);
            var charNine = '9'.charCodeAt(0);
            for (var i = 0; i < expression.length; i++) {
                if (expression.charCodeAt(i) < charZero ||
                    expression.charCodeAt(i) > charNine) {
                    return false;
                }
            }
            return true;
        };
        LibraryBuilder.prototype.evaluateEventType = function (eventTypeExpression) {
            if (Utility.isNullOrEmptyString(eventTypeExpression)) {
                return 0;
            }
            if (this.isAllDigits(eventTypeExpression)) {
                return parseInt(eventTypeExpression);
            }
            var ret = this.resolveObjectByFullName(eventTypeExpression);
            if (typeof (ret) !== 'number') {
                throw new OfficeExtension_1.Error("Invalid event type: " + eventTypeExpression);
            }
            return ret;
        };
        LibraryBuilder.prototype.buildPrototype = function (type, typeInfo) {
            this.buildScalarProperties(type, typeInfo);
            this.buildNavigationProperties(type, typeInfo);
            this.buildScalarMethods(type, typeInfo);
            this.buildNavigationMethods(type, typeInfo);
            this.buildEvents(type, typeInfo);
            this.buildHandleResult(type, typeInfo);
            this.buildHandleIdResult(type, typeInfo);
            this.buildHandleRetrieveResult(type, typeInfo);
            this.buildLoad(type, typeInfo);
            this.buildRetrieve(type, typeInfo);
            this.buildSetMockData(type, typeInfo);
            this.buildEnsureUnchanged(type, typeInfo);
            this.buildUpdate(type, typeInfo);
            this.buildSet(type, typeInfo);
            this.buildToJSON(type, typeInfo);
            this.buildItems(type, typeInfo);
            this.buildTypeMetadataInfo(type, typeInfo);
            this.buildTrackUntrack(type, typeInfo);
            this.buildMixin(type, typeInfo);
        };
        LibraryBuilder.prototype.toSimpleCamelUpperCase = function (name) {
            return name.substr(0, 1).toUpperCase() + name.substr(1);
        };
        LibraryBuilder.prototype.ensureOriginalName = function (member) {
            if (member.originalName === null) {
                member.originalName = this.toSimpleCamelUpperCase(member.name);
            }
        };
        LibraryBuilder.prototype.getFieldName = function (member) {
            return "m_" + member.name;
        };
        LibraryBuilder.prototype.throwIfApiNotSupported = function (typeInfo, member) {
            if (this.m_apiSets && member.apiSetInfoOrdinal > 0) {
                var apiSetInfo = this.m_apiSets[member.apiSetInfoOrdinal - 1];
                if (apiSetInfo) {
                    Utility.throwIfApiNotSupported(typeInfo.name + "." + member.name, apiSetInfo.name, apiSetInfo.version, this.m_hostName);
                }
            }
        };
        LibraryBuilder.prototype.buildScalarProperties = function (type, typeInfo) {
            if (Array.isArray(typeInfo.scalarProperties)) {
                for (var i = 0; i < typeInfo.scalarProperties.length; i++) {
                    var elem = typeInfo.scalarProperties[i];
                    if (Array.isArray(elem)) {
                        this.ensureArraySize(elem, 5);
                        typeInfo.scalarProperties[i] = {
                            name: this.getString(elem[0]),
                            behaviorFlags: elem[1],
                            apiSetInfoOrdinal: elem[2],
                            originalName: this.getString(elem[3]),
                            setMethodApiFlags: elem[4]
                        };
                    }
                    this.buildScalarProperty(type, typeInfo, typeInfo.scalarProperties[i]);
                }
            }
        };
        LibraryBuilder.prototype.buildScalarProperty = function (type, typeInfo, propInfo) {
            this.ensureOriginalName(propInfo);
            var thisBuilder = this;
            var fieldName = this.getFieldName(propInfo);
            var descriptor = {
                get: function () {
                    Utility.throwIfNotLoaded(propInfo.name, this[fieldName], typeInfo.name, this._isNull);
                    thisBuilder.throwIfApiNotSupported(typeInfo, propInfo);
                    return this[fieldName];
                },
                enumerable: true,
                configurable: true
            };
            if ((propInfo.behaviorFlags & 2) === 0) {
                descriptor.set = function (value) {
                    if (propInfo.behaviorFlags & 4) {
                        var customizationFunc = thisBuilder.getFunction(LibraryBuilder.CustomizationCodeNamespace + "." + typeInfo.name + "_" + propInfo.originalName + "_Set");
                        var handled = customizationFunc.call(this, this, value).handled;
                        if (handled) {
                            return;
                        }
                    }
                    this[fieldName] = value;
                    BatchApiHelper.invokeSetProperty(this, propInfo.originalName, value, propInfo.setMethodApiFlags);
                };
            }
            Object.defineProperty(type.prototype, propInfo.name, descriptor);
        };
        LibraryBuilder.prototype.buildNavigationProperties = function (type, typeInfo) {
            if (Array.isArray(typeInfo.navigationProperties)) {
                for (var i = 0; i < typeInfo.navigationProperties.length; i++) {
                    var elem = typeInfo.navigationProperties[i];
                    if (Array.isArray(elem)) {
                        this.ensureArraySize(elem, 7);
                        typeInfo.navigationProperties[i] = {
                            name: this.getString(elem[0]),
                            behaviorFlags: elem[2],
                            apiSetInfoOrdinal: elem[3],
                            originalName: this.getString(elem[4]),
                            getMethodApiFlags: elem[5],
                            setMethodApiFlags: elem[6],
                            propertyTypeFullName: this.getString(elem[1])
                        };
                    }
                    this.buildNavigationProperty(type, typeInfo, typeInfo.navigationProperties[i]);
                }
            }
        };
        LibraryBuilder.prototype.buildNavigationProperty = function (type, typeInfo, propInfo) {
            this.ensureOriginalName(propInfo);
            var thisBuilder = this;
            var fieldName = this.getFieldName(propInfo);
            var descriptor = {
                get: function () {
                    if (!this[thisBuilder.getFieldName(propInfo)]) {
                        thisBuilder.throwIfApiNotSupported(typeInfo, propInfo);
                        this[fieldName] = BatchApiHelper.createPropertyObject(thisBuilder.getFunction(propInfo.propertyTypeFullName), this, propInfo.originalName, (propInfo.behaviorFlags & 16) !== 0, propInfo.getMethodApiFlags);
                    }
                    if (propInfo.behaviorFlags & 64) {
                        var customizationFunc = thisBuilder.getFunction(LibraryBuilder.CustomizationCodeNamespace + "." + typeInfo.name + "_" + propInfo.originalName + "_Get");
                        customizationFunc.call(this, this, this[fieldName]);
                    }
                    return this[fieldName];
                },
                enumerable: true,
                configurable: true
            };
            if ((propInfo.behaviorFlags & 2) === 0) {
                descriptor.set = function (value) {
                    if (propInfo.behaviorFlags & 4) {
                        var customizationFunc = thisBuilder.getFunction(LibraryBuilder.CustomizationCodeNamespace + "." + typeInfo.name + "_" + propInfo.originalName + "_Set");
                        var handled = customizationFunc.call(this, this, value).handled;
                        if (handled) {
                            return;
                        }
                    }
                    this[fieldName] = value;
                    BatchApiHelper.invokeSetProperty(this, propInfo.originalName, value, propInfo.setMethodApiFlags);
                };
            }
            Object.defineProperty(type.prototype, propInfo.name, descriptor);
        };
        LibraryBuilder.prototype.buildScalarMethods = function (type, typeInfo) {
            if (Array.isArray(typeInfo.scalarMethods)) {
                for (var i = 0; i < typeInfo.scalarMethods.length; i++) {
                    var elem = typeInfo.scalarMethods[i];
                    if (Array.isArray(elem)) {
                        this.ensureArraySize(elem, 6);
                        typeInfo.scalarMethods[i] = {
                            name: this.getString(elem[0]),
                            behaviorFlags: elem[2],
                            apiSetInfoOrdinal: elem[3],
                            originalName: this.getString(elem[5]),
                            apiFlags: elem[4],
                            parameterCount: elem[1]
                        };
                    }
                    this.buildScalarMethod(type, typeInfo, typeInfo.scalarMethods[i]);
                }
            }
        };
        LibraryBuilder.prototype.buildScalarMethod = function (type, typeInfo, methodInfo) {
            this.ensureOriginalName(methodInfo);
            var thisBuilder = this;
            type.prototype[methodInfo.name] = function () {
                var args = [];
                if ((methodInfo.behaviorFlags & 64) && methodInfo.parameterCount > 0) {
                    for (var i = 0; i < methodInfo.parameterCount - 1; i++) {
                        args.push(arguments[i]);
                    }
                    var rest = [];
                    for (var i = methodInfo.parameterCount - 1; i < arguments.length; i++) {
                        rest.push(arguments[i]);
                    }
                    args.push(rest);
                }
                else {
                    for (var i = 0; i < arguments.length; i++) {
                        args.push(arguments[i]);
                    }
                }
                if (methodInfo.behaviorFlags & 1) {
                    var customizationFunc = thisBuilder.getFunction(LibraryBuilder.CustomizationCodeNamespace + "." + typeInfo.name + "_" + methodInfo.originalName);
                    var applyArgs = [this];
                    for (var i = 0; i < args.length; i++) {
                        applyArgs.push(args[i]);
                    }
                    var _a = customizationFunc.apply(this, applyArgs), handled = _a.handled, result = _a.result;
                    if (handled) {
                        return result;
                    }
                }
                thisBuilder.throwIfApiNotSupported(typeInfo, methodInfo);
                var resultProcessType = 0;
                if (methodInfo.behaviorFlags & 32) {
                    resultProcessType = 1;
                }
                var operationType = 0;
                if (methodInfo.behaviorFlags & 2) {
                    operationType = 1;
                }
                return BatchApiHelper.invokeMethod(this, methodInfo.originalName, operationType, args, methodInfo.apiFlags, resultProcessType);
            };
        };
        LibraryBuilder.prototype.buildNavigationMethods = function (type, typeInfo) {
            if (Array.isArray(typeInfo.navigationMethods)) {
                for (var i = 0; i < typeInfo.navigationMethods.length; i++) {
                    var elem = typeInfo.navigationMethods[i];
                    if (Array.isArray(elem)) {
                        this.ensureArraySize(elem, 8);
                        typeInfo.navigationMethods[i] = {
                            name: this.getString(elem[0]),
                            behaviorFlags: elem[3],
                            apiSetInfoOrdinal: elem[4],
                            originalName: this.getString(elem[6]),
                            apiFlags: elem[5],
                            parameterCount: elem[2],
                            returnTypeFullName: this.getString(elem[1]),
                            returnObjectGetByIdMethodName: this.getString(elem[7])
                        };
                    }
                    this.buildNavigationMethod(type, typeInfo, typeInfo.navigationMethods[i]);
                }
            }
        };
        LibraryBuilder.prototype.buildNavigationMethod = function (type, typeInfo, methodInfo) {
            this.ensureOriginalName(methodInfo);
            var thisBuilder = this;
            type.prototype[methodInfo.name] = function () {
                var args = [];
                if ((methodInfo.behaviorFlags & 64) && methodInfo.parameterCount > 0) {
                    for (var i = 0; i < methodInfo.parameterCount - 1; i++) {
                        args.push(arguments[i]);
                    }
                    var rest = [];
                    for (var i = methodInfo.parameterCount - 1; i < arguments.length; i++) {
                        rest.push(arguments[i]);
                    }
                    args.push(rest);
                }
                else {
                    for (var i = 0; i < arguments.length; i++) {
                        args.push(arguments[i]);
                    }
                }
                if (methodInfo.behaviorFlags & 1) {
                    var customizationFunc = thisBuilder.getFunction(LibraryBuilder.CustomizationCodeNamespace + "." + typeInfo.name + "_" + methodInfo.originalName);
                    var applyArgs = [this];
                    for (var i = 0; i < args.length; i++) {
                        applyArgs.push(args[i]);
                    }
                    var _a = customizationFunc.apply(this, applyArgs), handled = _a.handled, result = _a.result;
                    if (handled) {
                        return result;
                    }
                }
                thisBuilder.throwIfApiNotSupported(typeInfo, methodInfo);
                if ((methodInfo.behaviorFlags & 16) !== 0) {
                    return BatchApiHelper.createIndexerObject(thisBuilder.getFunction(methodInfo.returnTypeFullName), this, args);
                }
                else {
                    var operationType = 0;
                    if (methodInfo.behaviorFlags & 2) {
                        operationType = 1;
                    }
                    return BatchApiHelper.createMethodObject(thisBuilder.getFunction(methodInfo.returnTypeFullName), this, methodInfo.originalName, operationType, args, (methodInfo.behaviorFlags & 4) !== 0, (methodInfo.behaviorFlags & 8) !== 0, methodInfo.returnObjectGetByIdMethodName, methodInfo.apiFlags);
                }
            };
        };
        LibraryBuilder.prototype.buildHandleResult = function (type, typeInfo) {
            var thisBuilder = this;
            type.prototype._handleResult = function (value) {
                ClientObject.prototype._handleResult.call(this, value);
                if (Utility.isNullOrUndefined(value)) {
                    return;
                }
                Utility.fixObjectPathIfNecessary(this, value);
                if (typeInfo.behaviorFlags & 8) {
                    var customizationFunc = thisBuilder.getFunction(LibraryBuilder.CustomizationCodeNamespace + "." + typeInfo.name + "_HandleResult");
                    customizationFunc.call(this, this, value);
                }
                if (typeInfo.scalarProperties) {
                    for (var i_3 = 0; i_3 < typeInfo.scalarProperties.length; i_3++) {
                        if (!Utility.isUndefined(value[typeInfo.scalarProperties[i_3].originalName])) {
                            if ((typeInfo.scalarProperties[i_3].behaviorFlags & 8) !== 0) {
                                this[thisBuilder.getFieldName(typeInfo.scalarProperties[i_3])] = Utility.adjustToDateTime(value[typeInfo.scalarProperties[i_3].originalName]);
                            }
                            else {
                                this[thisBuilder.getFieldName(typeInfo.scalarProperties[i_3])] = value[typeInfo.scalarProperties[i_3].originalName];
                            }
                        }
                    }
                }
                if (typeInfo.navigationProperties) {
                    var propNames = [];
                    for (var i_4 = 0; i_4 < typeInfo.navigationProperties.length; i_4++) {
                        propNames.push(typeInfo.navigationProperties[i_4].name);
                        propNames.push(typeInfo.navigationProperties[i_4].originalName);
                    }
                    Utility._handleNavigationPropertyResults(this, value, propNames);
                }
                if ((typeInfo.behaviorFlags & 1) !== 0) {
                    var hasIndexerMethod = thisBuilder.hasIndexMethod(typeInfo);
                    if (!Utility.isNullOrUndefined(value[Constants.items])) {
                        this.m__items = [];
                        var _data = value[Constants.items];
                        var childItemType = thisBuilder.getFunction(typeInfo.childItemTypeFullName);
                        for (var i = 0; i < _data.length; i++) {
                            var _item = BatchApiHelper.createChildItemObject(childItemType, hasIndexerMethod, this, _data[i], i);
                            _item._handleResult(_data[i]);
                            this.m__items.push(_item);
                        }
                    }
                }
            };
        };
        LibraryBuilder.prototype.buildHandleRetrieveResult = function (type, typeInfo) {
            var thisBuilder = this;
            type.prototype._handleRetrieveResult = function (value, result) {
                ClientObject.prototype._handleRetrieveResult.call(this, value, result);
                if (Utility.isNullOrUndefined(value)) {
                    return;
                }
                if (typeInfo.scalarProperties) {
                    for (var i = 0; i < typeInfo.scalarProperties.length; i++) {
                        if (typeInfo.scalarProperties[i].behaviorFlags & 8) {
                            if (!Utility.isNullOrUndefined(value[typeInfo.scalarProperties[i].name])) {
                                value[typeInfo.scalarProperties[i].name] = Utility.adjustToDateTime(value[typeInfo.scalarProperties[i].name]);
                            }
                        }
                    }
                }
                if (typeInfo.behaviorFlags & 1) {
                    var hasIndexerMethod_1 = thisBuilder.hasIndexMethod(typeInfo);
                    var childItemType_1 = thisBuilder.getFunction(typeInfo.childItemTypeFullName);
                    var thisObj_1 = this;
                    Utility.processRetrieveResult(thisObj_1, value, result, function (childItemData, index) { return BatchApiHelper.createChildItemObject(childItemType_1, hasIndexerMethod_1, thisObj_1, childItemData, index); });
                }
                else {
                    Utility.processRetrieveResult(this, value, result);
                }
            };
        };
        LibraryBuilder.prototype.buildHandleIdResult = function (type, typeInfo) {
            var thisBuilder = this;
            type.prototype._handleIdResult = function (value) {
                ClientObject.prototype._handleIdResult.call(this, value);
                if (Utility.isNullOrUndefined(value)) {
                    return;
                }
                if (typeInfo.behaviorFlags & 16) {
                    var customizationFunc = thisBuilder.getFunction(LibraryBuilder.CustomizationCodeNamespace + "." + typeInfo.name + "_HandleIdResult");
                    customizationFunc.call(this, this, value);
                }
                if (typeInfo.scalarProperties) {
                    for (var i = 0; i < typeInfo.scalarProperties.length; i++) {
                        var propName = typeInfo.scalarProperties[i].originalName;
                        if (propName === "Id" || propName === "_Id" || propName === "_ReferenceId") {
                            if (!Utility.isNullOrUndefined(value[typeInfo.scalarProperties[i].originalName])) {
                                this[thisBuilder.getFieldName(typeInfo.scalarProperties[i])] = value[typeInfo.scalarProperties[i].originalName];
                            }
                        }
                    }
                }
            };
        };
        LibraryBuilder.prototype.buildLoad = function (type, typeInfo) {
            type.prototype.load = function (options) {
                return Utility.load(this, options);
            };
        };
        LibraryBuilder.prototype.buildRetrieve = function (type, typeInfo) {
            type.prototype.retrieve = function (options) {
                return Utility.retrieve(this, options);
            };
        };
        LibraryBuilder.prototype.buildNewObject = function (type, typeInfo) {
            if (!Utility.isNullOrEmptyString(typeInfo.newObjectServerTypeFullName)) {
                type.newObject = function (context) {
                    return BatchApiHelper.createTopLevelServiceObject(type, context, typeInfo.newObjectServerTypeFullName, (typeInfo.behaviorFlags & 1) !== 0, typeInfo.newObjectApiFlags);
                };
            }
        };
        LibraryBuilder.prototype.buildSetMockData = function (type, typeInfo) {
            var thisBuilder = this;
            if (typeInfo.behaviorFlags & 1) {
                var hasIndexMethod_1 = thisBuilder.hasIndexMethod(typeInfo);
                type.prototype.setMockData = function (data) {
                    var thisObj = this;
                    Utility.setMockData(thisObj, data, function (childItemData, index) {
                        return BatchApiHelper.createChildItemObject(thisBuilder.getFunction(typeInfo.childItemTypeFullName), hasIndexMethod_1, thisObj, childItemData, index);
                    }, function (items) {
                        thisObj.m__items = items;
                    });
                };
            }
            else {
                type.prototype.setMockData = function (data) {
                    Utility.setMockData(this, data);
                };
            }
        };
        LibraryBuilder.prototype.buildEnsureUnchanged = function (type, typeInfo) {
            type.prototype.ensureUnchanged = function (data) {
                BatchApiHelper.invokeEnsureUnchanged(this, data);
            };
        };
        LibraryBuilder.prototype.buildUpdate = function (type, typeInfo) {
            type.prototype.update = function (properties) {
                this._recursivelyUpdate(properties);
            };
        };
        LibraryBuilder.prototype.buildSet = function (type, typeInfo) {
            if ((typeInfo.behaviorFlags & 1) !== 0) {
                return;
            }
            var notAllowedToBeSetPropertyNames = [];
            var allowedScalarPropertyNames = [];
            if (typeInfo.scalarProperties) {
                for (var i = 0; i < typeInfo.scalarProperties.length; i++) {
                    if ((typeInfo.scalarProperties[i].behaviorFlags & 2) === 0 &&
                        (typeInfo.scalarProperties[i].behaviorFlags & 1) !== 0) {
                        allowedScalarPropertyNames.push(typeInfo.scalarProperties[i].name);
                    }
                    else {
                        notAllowedToBeSetPropertyNames.push(typeInfo.scalarProperties[i].name);
                    }
                }
            }
            var allowedNavigationPropertyNames = [];
            if (typeInfo.navigationProperties) {
                for (var i = 0; i < typeInfo.navigationProperties.length; i++) {
                    if ((typeInfo.navigationProperties[i].behaviorFlags & 16) !== 0) {
                        notAllowedToBeSetPropertyNames.push(typeInfo.navigationProperties[i].name);
                    }
                    else if ((typeInfo.navigationProperties[i].behaviorFlags & 1) === 0) {
                        notAllowedToBeSetPropertyNames.push(typeInfo.navigationProperties[i].name);
                    }
                    else if ((typeInfo.navigationProperties[i].behaviorFlags & 32) === 0) {
                        notAllowedToBeSetPropertyNames.push(typeInfo.navigationProperties[i].name);
                    }
                    else {
                        allowedNavigationPropertyNames.push(typeInfo.navigationProperties[i].name);
                    }
                }
            }
            if (allowedNavigationPropertyNames.length === 0 && allowedScalarPropertyNames.length === 0) {
                return;
            }
            type.prototype.set = function (properties, options) {
                this._recursivelySet(properties, options, allowedScalarPropertyNames, allowedNavigationPropertyNames, notAllowedToBeSetPropertyNames);
            };
        };
        LibraryBuilder.prototype.buildItems = function (type, typeInfo) {
            if ((typeInfo.behaviorFlags & 1) === 0) {
                return;
            }
            Object.defineProperty(type.prototype, "items", {
                get: function () {
                    Utility.throwIfNotLoaded("items", this.m__items, typeInfo.name, this._isNull);
                    return this.m__items;
                },
                enumerable: true,
                configurable: true
            });
        };
        LibraryBuilder.prototype.buildToJSON = function (type, typeInfo) {
            var thisBuilder = this;
            if ((typeInfo.behaviorFlags & 1) !== 0) {
                type.prototype.toJSON = function () {
                    return Utility.toJson(this, {}, {}, this.m__items);
                };
                return;
            }
            else {
                type.prototype.toJSON = function () {
                    var scalarProperties = {};
                    if (typeInfo.scalarProperties) {
                        for (var i = 0; i < typeInfo.scalarProperties.length; i++) {
                            if ((typeInfo.scalarProperties[i].behaviorFlags & 1) !== 0) {
                                scalarProperties[typeInfo.scalarProperties[i].name] = this[thisBuilder.getFieldName(typeInfo.scalarProperties[i])];
                            }
                        }
                    }
                    var navProperties = {};
                    if (typeInfo.navigationProperties) {
                        for (var i = 0; i < typeInfo.navigationProperties.length; i++) {
                            if ((typeInfo.navigationProperties[i].behaviorFlags & 1) !== 0) {
                                navProperties[typeInfo.navigationProperties[i].name] = this[thisBuilder.getFieldName(typeInfo.navigationProperties[i])];
                            }
                        }
                    }
                    return Utility.toJson(this, scalarProperties, navProperties);
                };
            }
        };
        LibraryBuilder.prototype.buildTypeMetadataInfo = function (type, typeInfo) {
            Object.defineProperty(type.prototype, "_className", {
                get: function () {
                    return typeInfo.name;
                },
                enumerable: true,
                configurable: true
            });
            Object.defineProperty(type.prototype, "_isCollection", {
                get: function () {
                    return (typeInfo.behaviorFlags & 1) !== 0;
                },
                enumerable: true,
                configurable: true
            });
            if (!Utility.isNullOrEmptyString(typeInfo.collectionPropertyPath)) {
                Object.defineProperty(type.prototype, "_collectionPropertyPath", {
                    get: function () {
                        return typeInfo.collectionPropertyPath;
                    },
                    enumerable: true,
                    configurable: true
                });
            }
            if (typeInfo.scalarProperties && typeInfo.scalarProperties.length > 0) {
                Object.defineProperty(type.prototype, "_scalarPropertyNames", {
                    get: function () {
                        if (!this.m__scalarPropertyNames) {
                            this.m__scalarPropertyNames = typeInfo.scalarProperties.map(function (p) { return p.name; });
                        }
                        return this.m__scalarPropertyNames;
                    },
                    enumerable: true,
                    configurable: true
                });
                Object.defineProperty(type.prototype, "_scalarPropertyOriginalNames", {
                    get: function () {
                        if (!this.m__scalarPropertyOriginalNames) {
                            this.m__scalarPropertyOriginalNames = typeInfo.scalarProperties.map(function (p) { return p.originalName; });
                        }
                        return this.m__scalarPropertyOriginalNames;
                    },
                    enumerable: true,
                    configurable: true
                });
                Object.defineProperty(type.prototype, "_scalarPropertyUpdateable", {
                    get: function () {
                        if (!this.m__scalarPropertyUpdateable) {
                            this.m__scalarPropertyUpdateable = typeInfo.scalarProperties.map(function (p) { return (p.behaviorFlags & 2) === 0; });
                        }
                        return this.m__scalarPropertyUpdateable;
                    },
                    enumerable: true,
                    configurable: true
                });
            }
            if (typeInfo.navigationProperties && typeInfo.navigationProperties.length > 0) {
                Object.defineProperty(type.prototype, "_navigationPropertyNames", {
                    get: function () {
                        if (!this.m__navigationPropertyNames) {
                            this.m__navigationPropertyNames = typeInfo.navigationProperties.map(function (p) { return p.name; });
                        }
                        return this.m__navigationPropertyNames;
                    },
                    enumerable: true,
                    configurable: true
                });
            }
        };
        LibraryBuilder.prototype.buildTrackUntrack = function (type, typeInfo) {
            if (typeInfo.behaviorFlags & 2) {
                type.prototype.track = function () {
                    this.context.trackedObjects.add(this);
                    return this;
                };
                type.prototype.untrack = function () {
                    this.context.trackedObjects.remove(this);
                    return this;
                };
            }
        };
        LibraryBuilder.prototype.buildMixin = function (type, typeInfo) {
            if (typeInfo.behaviorFlags & 4) {
                var mixinType = this.getFunction(typeInfo.name + 'Custom');
                Utility.applyMixin(type, mixinType);
            }
        };
        LibraryBuilder.prototype.getOnEventName = function (name) {
            if (name[0] === '_') {
                return '_on' + name.substr(1);
            }
            return 'on' + name;
        };
        LibraryBuilder.prototype.buildEvents = function (type, typeInfo) {
            if (typeInfo.events) {
                for (var i = 0; i < typeInfo.events.length; i++) {
                    var elem = typeInfo.events[i];
                    if (Array.isArray(elem)) {
                        this.ensureArraySize(elem, 7);
                        typeInfo.events[i] = {
                            name: this.getString(elem[0]),
                            behaviorFlags: elem[1],
                            apiSetInfoOrdinal: elem[2],
                            typeExpression: this.getString(elem[3]),
                            targetIdExpression: this.getString(elem[4]),
                            register: this.getString(elem[5]),
                            unregister: this.getString(elem[6])
                        };
                    }
                    this.buildEvent(type, typeInfo, typeInfo.events[i]);
                }
            }
        };
        LibraryBuilder.prototype.buildEvent = function (type, typeInfo, evt) {
            if (evt.behaviorFlags & 1) {
                this.buildV0Event(type, typeInfo, evt);
            }
            else {
                this.buildV2Event(type, typeInfo, evt);
            }
        };
        LibraryBuilder.prototype.buildV2Event = function (type, typeInfo, evt) {
            var thisBuilder = this;
            var eventName = this.getOnEventName(evt.name);
            var fieldName = this.getFieldName(evt);
            Object.defineProperty(type.prototype, eventName, {
                get: function () {
                    if (!this[fieldName]) {
                        thisBuilder.throwIfApiNotSupported(typeInfo, evt);
                        var thisObj = this;
                        var registerFunc = null;
                        if (evt.register !== 'null') {
                            registerFunc = this[evt.register].bind(this);
                        }
                        var unregisterFunc = null;
                        if (evt.unregister !== 'null') {
                            unregisterFunc = this[evt.unregister].bind(this);
                        }
                        var getTargetIdFunc = function () {
                            return thisBuilder.evaluateEventTargetId(evt.targetIdExpression, thisObj);
                        };
                        var func = null;
                        if (evt.behaviorFlags & 2) {
                            func = thisBuilder.getFunction(LibraryBuilder.CustomizationCodeNamespace + "." + typeInfo.name + "_" + evt.name + "_EventArgsTransform");
                        }
                        var eventArgsTransformFunc = function (value) {
                            if (func) {
                                value = func.call(thisObj, thisObj, value);
                            }
                            return Utility._createPromiseFromResult(value);
                        };
                        var eventType = thisBuilder.evaluateEventType(evt.typeExpression);
                        this[fieldName] = new GenericEventHandlers(this.context, this, evt.name, {
                            eventType: eventType,
                            getTargetIdFunc: getTargetIdFunc,
                            registerFunc: registerFunc,
                            unregisterFunc: unregisterFunc,
                            eventArgsTransformFunc: eventArgsTransformFunc
                        });
                    }
                    return this[fieldName];
                },
                enumerable: true,
                configurable: true
            });
        };
        LibraryBuilder.prototype.buildV0Event = function (type, typeInfo, evt) {
            var thisBuilder = this;
            var eventName = this.getOnEventName(evt.name);
            var fieldName = this.getFieldName(evt);
            Object.defineProperty(type.prototype, eventName, {
                get: function () {
                    if (!this[fieldName]) {
                        thisBuilder.throwIfApiNotSupported(typeInfo, evt);
                        var thisObj = this;
                        var registerFunc = null;
                        if (Utility.isNullOrEmptyString(evt.register)) {
                            var eventType_1 = thisBuilder.evaluateEventType(evt.typeExpression);
                            registerFunc =
                                function (handlerCallback) {
                                    var targetId = thisBuilder.evaluateEventTargetId(evt.targetIdExpression, thisObj);
                                    return thisObj.context.eventRegistration.register(eventType_1, targetId, handlerCallback);
                                };
                        }
                        else if (evt.register !== 'null') {
                            var func_1 = thisBuilder.getFunction(evt.register);
                            registerFunc =
                                function (handlerCallback) {
                                    return func_1.call(thisObj, thisObj, handlerCallback);
                                };
                        }
                        var unregisterFunc = null;
                        if (Utility.isNullOrEmptyString(evt.unregister)) {
                            var eventType_2 = thisBuilder.evaluateEventType(evt.typeExpression);
                            unregisterFunc =
                                function (handlerCallback) {
                                    var targetId = thisBuilder.evaluateEventTargetId(evt.targetIdExpression, thisObj);
                                    return thisObj.context.eventRegistration.unregister(eventType_2, targetId, handlerCallback);
                                };
                        }
                        else if (evt.unregister !== 'null') {
                            var func_2 = thisBuilder.getFunction(evt.unregister);
                            unregisterFunc =
                                function (handlerCallback) {
                                    return func_2.call(thisObj, thisObj, handlerCallback);
                                };
                        }
                        var func = null;
                        if (evt.behaviorFlags & 2) {
                            func = thisBuilder.getFunction(LibraryBuilder.CustomizationCodeNamespace + "." + typeInfo.name + "_" + evt.name + "_EventArgsTransform");
                        }
                        var eventArgsTransformFunc = function (value) {
                            if (func) {
                                value = func.call(thisObj, thisObj, value);
                            }
                            return Utility._createPromiseFromResult(value);
                        };
                        this[fieldName] = new EventHandlers(this.context, this, evt.name, {
                            registerFunc: registerFunc,
                            unregisterFunc: unregisterFunc,
                            eventArgsTransformFunc: eventArgsTransformFunc
                        });
                    }
                    return this[fieldName];
                },
                enumerable: true,
                configurable: true
            });
        };
        LibraryBuilder.prototype.hasIndexMethod = function (typeInfo) {
            var ret = false;
            if (typeInfo.navigationMethods) {
                for (var i = 0; i < typeInfo.navigationMethods.length; i++) {
                    if ((typeInfo.navigationMethods[i].behaviorFlags & 16) !== 0) {
                        ret = true;
                        break;
                    }
                }
            }
            return ret;
        };
        LibraryBuilder.CustomizationCodeNamespace = "_CC";
        return LibraryBuilder;
    }());
    OfficeExtension_1.LibraryBuilder = LibraryBuilder;
})(OfficeExtension || (OfficeExtension = {}));
var OfficeExtension;
(function (OfficeExtension) {
    var _InternalPromise;
    (function (_InternalPromise) {
        function getPromiseType() {
            if (typeof Promise !== 'undefined') {
                return Promise;
            }
            if (typeof Office !== 'undefined') {
                if (Office.Promise) {
                    return Office.Promise;
                }
            }
            throw new OfficeExtension.Error('No Promise implementation found');
        }
        _InternalPromise.getPromiseType = getPromiseType;
    })(_InternalPromise || (_InternalPromise = {}));
    Object.defineProperty(OfficeExtension, "Promise", {
        get: function () {
            return _InternalPromise.getPromiseType();
        },
        enumerable: true,
        configurable: true
    });
})(OfficeExtension || (OfficeExtension = {}));
var __extends = (this && this.__extends) || (function () {
    var extendStatics = function (d, b) {
        extendStatics = Object.setPrototypeOf ||
            ({ __proto__: [] } instanceof Array && function (d, b) { d.__proto__ = b; }) ||
            function (d, b) { for (var p in b) if (b.hasOwnProperty(p)) d[p] = b[p]; };
        return extendStatics(d, b);
    };
    return function (d, b) {
        extendStatics(d, b);
        function __() { this.constructor = d; }
        d.prototype = b === null ? Object.create(b) : (__.prototype = b.prototype, new __());
    };
})();
var __awaiter = (this && this.__awaiter) || function (thisArg, _arguments, P, generator) {
    return new (P || (P = Promise))(function (resolve, reject) {
        function fulfilled(value) { try { step(generator.next(value)); } catch (e) { reject(e); } }
        function rejected(value) { try { step(generator["throw"](value)); } catch (e) { reject(e); } }
        function step(result) { result.done ? resolve(result.value) : new P(function (resolve) { resolve(result.value); }).then(fulfilled, rejected); }
        step((generator = generator.apply(thisArg, _arguments || [])).next());
    });
};
var __generator = (this && this.__generator) || function (thisArg, body) {
    var _ = { label: 0, sent: function() { if (t[0] & 1) throw t[1]; return t[1]; }, trys: [], ops: [] }, f, y, t, g;
    return g = { next: verb(0), "throw": verb(1), "return": verb(2) }, typeof Symbol === "function" && (g[Symbol.iterator] = function() { return this; }), g;
    function verb(n) { return function (v) { return step([n, v]); }; }
    function step(op) {
        if (f) throw new TypeError("Generator is already executing.");
        while (_) try {
            if (f = 1, y && (t = op[0] & 2 ? y["return"] : op[0] ? y["throw"] || ((t = y["return"]) && t.call(y), 0) : y.next) && !(t = t.call(y, op[1])).done) return t;
            if (y = 0, t) op = [op[0] & 2, t.value];
            switch (op[0]) {
                case 0: case 1: t = op; break;
                case 4: _.label++; return { value: op[1], done: false };
                case 5: _.label++; y = op[1]; op = [0]; continue;
                case 7: op = _.ops.pop(); _.trys.pop(); continue;
                default:
                    if (!(t = _.trys, t = t.length > 0 && t[t.length - 1]) && (op[0] === 6 || op[0] === 2)) { _ = 0; continue; }
                    if (op[0] === 3 && (!t || (op[1] > t[0] && op[1] < t[3]))) { _.label = op[1]; break; }
                    if (op[0] === 6 && _.label < t[1]) { _.label = t[1]; t = op; break; }
                    if (t && _.label < t[2]) { _.label = t[2]; _.ops.push(op); break; }
                    if (t[2]) _.ops.pop();
                    _.trys.pop(); continue;
            }
            op = body.call(thisArg, _);
        } catch (e) { op = [6, e]; y = 0; } finally { f = t = 0; }
        if (op[0] & 5) throw op[1]; return { value: op[0] ? op[1] : void 0, done: true };
    }
};
var OfficeCore;
(function (OfficeCore) {
    var _hostName = "OfficeCore";
    var _defaultApiSetName = "AgaveVisualApi";
    var _createPropertyObject = OfficeExtension.BatchApiHelper.createPropertyObject;
    var _createMethodObject = OfficeExtension.BatchApiHelper.createMethodObject;
    var _createIndexerObject = OfficeExtension.BatchApiHelper.createIndexerObject;
    var _createRootServiceObject = OfficeExtension.BatchApiHelper.createRootServiceObject;
    var _createTopLevelServiceObject = OfficeExtension.BatchApiHelper.createTopLevelServiceObject;
    var _createChildItemObject = OfficeExtension.BatchApiHelper.createChildItemObject;
    var _invokeMethod = OfficeExtension.BatchApiHelper.invokeMethod;
    var _invokeEnsureUnchanged = OfficeExtension.BatchApiHelper.invokeEnsureUnchanged;
    var _invokeSetProperty = OfficeExtension.BatchApiHelper.invokeSetProperty;
    var _isNullOrUndefined = OfficeExtension.Utility.isNullOrUndefined;
    var _isUndefined = OfficeExtension.Utility.isUndefined;
    var _throwIfNotLoaded = OfficeExtension.Utility.throwIfNotLoaded;
    var _throwIfApiNotSupported = OfficeExtension.Utility.throwIfApiNotSupported;
    var _load = OfficeExtension.Utility.load;
    var _retrieve = OfficeExtension.Utility.retrieve;
    var _toJson = OfficeExtension.Utility.toJson;
    var _fixObjectPathIfNecessary = OfficeExtension.Utility.fixObjectPathIfNecessary;
    var _handleNavigationPropertyResults = OfficeExtension.Utility._handleNavigationPropertyResults;
    var _adjustToDateTime = OfficeExtension.Utility.adjustToDateTime;
    var _processRetrieveResult = OfficeExtension.Utility.processRetrieveResult;
    var _typeBiShim = "BiShim";
    var BiShim = (function (_super) {
        __extends(BiShim, _super);
        function BiShim() {
            return _super !== null && _super.apply(this, arguments) || this;
        }
        Object.defineProperty(BiShim.prototype, "_className", {
            get: function () {
                return "BiShim";
            },
            enumerable: true,
            configurable: true
        });
        BiShim.prototype.initialize = function (capabilities) {
            _invokeMethod(this, "Initialize", 0, [capabilities], 0, 0);
        };
        BiShim.prototype.getData = function () {
            return _invokeMethod(this, "getData", 1, [], 4, 0);
        };
        BiShim.prototype.setVisualObjects = function (visualObjects) {
            _invokeMethod(this, "setVisualObjects", 0, [visualObjects], 2, 0);
        };
        BiShim.prototype.setVisualObjectsToPersist = function (visualObjectsToPersist) {
            _invokeMethod(this, "setVisualObjectsToPersist", 0, [visualObjectsToPersist], 2, 0);
        };
        BiShim.prototype._handleResult = function (value) {
            _super.prototype._handleResult.call(this, value);
            if (_isNullOrUndefined(value))
                return;
            var obj = value;
            _fixObjectPathIfNecessary(this, obj);
        };
        BiShim.prototype._handleRetrieveResult = function (value, result) {
            _super.prototype._handleRetrieveResult.call(this, value, result);
            _processRetrieveResult(this, value, result);
        };
        BiShim.newObject = function (context) {
            return _createTopLevelServiceObject(OfficeCore.BiShim, context, "Microsoft.AgaveVisual.BiShim", false, 4);
        };
        BiShim.prototype.toJSON = function () {
            return _toJson(this, {}, {});
        };
        return BiShim;
    }(OfficeExtension.ClientObject));
    OfficeCore.BiShim = BiShim;
    var AgaveVisualErrorCodes;
    (function (AgaveVisualErrorCodes) {
        AgaveVisualErrorCodes["generalException1"] = "GeneralException";
    })(AgaveVisualErrorCodes = OfficeCore.AgaveVisualErrorCodes || (OfficeCore.AgaveVisualErrorCodes = {}));
})(OfficeCore || (OfficeCore = {}));
var OfficeCore;
(function (OfficeCore) {
    var _hostName = "OfficeCore";
    var _defaultApiSetName = "ExperimentApi";
    var _createPropertyObject = OfficeExtension.BatchApiHelper.createPropertyObject;
    var _createMethodObject = OfficeExtension.BatchApiHelper.createMethodObject;
    var _createIndexerObject = OfficeExtension.BatchApiHelper.createIndexerObject;
    var _createRootServiceObject = OfficeExtension.BatchApiHelper.createRootServiceObject;
    var _createTopLevelServiceObject = OfficeExtension.BatchApiHelper.createTopLevelServiceObject;
    var _createChildItemObject = OfficeExtension.BatchApiHelper.createChildItemObject;
    var _invokeMethod = OfficeExtension.BatchApiHelper.invokeMethod;
    var _invokeEnsureUnchanged = OfficeExtension.BatchApiHelper.invokeEnsureUnchanged;
    var _invokeSetProperty = OfficeExtension.BatchApiHelper.invokeSetProperty;
    var _isNullOrUndefined = OfficeExtension.Utility.isNullOrUndefined;
    var _isUndefined = OfficeExtension.Utility.isUndefined;
    var _throwIfNotLoaded = OfficeExtension.Utility.throwIfNotLoaded;
    var _throwIfApiNotSupported = OfficeExtension.Utility.throwIfApiNotSupported;
    var _load = OfficeExtension.Utility.load;
    var _retrieve = OfficeExtension.Utility.retrieve;
    var _toJson = OfficeExtension.Utility.toJson;
    var _fixObjectPathIfNecessary = OfficeExtension.Utility.fixObjectPathIfNecessary;
    var _handleNavigationPropertyResults = OfficeExtension.Utility._handleNavigationPropertyResults;
    var _adjustToDateTime = OfficeExtension.Utility.adjustToDateTime;
    var _processRetrieveResult = OfficeExtension.Utility.processRetrieveResult;
    var _typeFlightingService = "FlightingService";
    var FlightingService = (function (_super) {
        __extends(FlightingService, _super);
        function FlightingService() {
            return _super !== null && _super.apply(this, arguments) || this;
        }
        Object.defineProperty(FlightingService.prototype, "_className", {
            get: function () {
                return "FlightingService";
            },
            enumerable: true,
            configurable: true
        });
        FlightingService.prototype.getClientSessionId = function () {
            return _invokeMethod(this, "GetClientSessionId", 1, [], 4, 0);
        };
        FlightingService.prototype.getDeferredFlights = function () {
            return _invokeMethod(this, "GetDeferredFlights", 1, [], 4, 0);
        };
        FlightingService.prototype.getFeature = function (featureName, type, defaultValue, possibleValues) {
            return _createMethodObject(OfficeCore.ABType, this, "GetFeature", 1, [featureName, type, defaultValue, possibleValues], false, false, null, 4);
        };
        FlightingService.prototype.getFeatureGate = function (featureName, scope) {
            return _createMethodObject(OfficeCore.ABType, this, "GetFeatureGate", 1, [featureName, scope], false, false, null, 4);
        };
        FlightingService.prototype.resetOverride = function (featureName) {
            _invokeMethod(this, "ResetOverride", 0, [featureName], 0, 0);
        };
        FlightingService.prototype.setOverride = function (featureName, type, value) {
            _invokeMethod(this, "SetOverride", 0, [featureName, type, value], 0, 0);
        };
        FlightingService.prototype._handleResult = function (value) {
            _super.prototype._handleResult.call(this, value);
            if (_isNullOrUndefined(value))
                return;
            var obj = value;
            _fixObjectPathIfNecessary(this, obj);
        };
        FlightingService.prototype._handleRetrieveResult = function (value, result) {
            _super.prototype._handleRetrieveResult.call(this, value, result);
            _processRetrieveResult(this, value, result);
        };
        FlightingService.newObject = function (context) {
            return _createTopLevelServiceObject(OfficeCore.FlightingService, context, "Microsoft.Experiment.FlightingService", false, 4);
        };
        FlightingService.prototype.toJSON = function () {
            return _toJson(this, {}, {});
        };
        return FlightingService;
    }(OfficeExtension.ClientObject));
    OfficeCore.FlightingService = FlightingService;
    var _typeABType = "ABType";
    var ABType = (function (_super) {
        __extends(ABType, _super);
        function ABType() {
            return _super !== null && _super.apply(this, arguments) || this;
        }
        Object.defineProperty(ABType.prototype, "_className", {
            get: function () {
                return "ABType";
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(ABType.prototype, "_scalarPropertyNames", {
            get: function () {
                return ["value"];
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(ABType.prototype, "value", {
            get: function () {
                _throwIfNotLoaded("value", this._V, _typeABType, this._isNull);
                return this._V;
            },
            enumerable: true,
            configurable: true
        });
        ABType.prototype._handleResult = function (value) {
            _super.prototype._handleResult.call(this, value);
            if (_isNullOrUndefined(value))
                return;
            var obj = value;
            _fixObjectPathIfNecessary(this, obj);
            if (!_isUndefined(obj["Value"])) {
                this._V = obj["Value"];
            }
        };
        ABType.prototype.load = function (option) {
            return _load(this, option);
        };
        ABType.prototype.retrieve = function (option) {
            return _retrieve(this, option);
        };
        ABType.prototype._handleRetrieveResult = function (value, result) {
            _super.prototype._handleRetrieveResult.call(this, value, result);
            _processRetrieveResult(this, value, result);
        };
        ABType.prototype.toJSON = function () {
            return _toJson(this, {
                "value": this._V,
            }, {});
        };
        ABType.prototype.ensureUnchanged = function (data) {
            _invokeEnsureUnchanged(this, data);
            return;
        };
        return ABType;
    }(OfficeExtension.ClientObject));
    OfficeCore.ABType = ABType;
    var FeatureType;
    (function (FeatureType) {
        FeatureType["boolean"] = "Boolean";
        FeatureType["integer"] = "Integer";
        FeatureType["string"] = "String";
    })(FeatureType = OfficeCore.FeatureType || (OfficeCore.FeatureType = {}));
    var ExperimentErrorCodes;
    (function (ExperimentErrorCodes) {
        ExperimentErrorCodes["generalException"] = "GeneralException";
    })(ExperimentErrorCodes = OfficeCore.ExperimentErrorCodes || (OfficeCore.ExperimentErrorCodes = {}));
})(OfficeCore || (OfficeCore = {}));
var OfficeCore;
(function (OfficeCore) {
    OfficeCore.OfficeOnlineDomainList = [
        "*.dod.online.office365.us",
        "*.gov.online.office365.us",
        "*.officeapps-df.live.com",
        "*.officeapps.live.com",
        "*.online.office.de",
        "*.partner.officewebapps.cn"
    ];
    function isHostOriginTrusted() {
        if (typeof window.external === 'undefined' ||
            typeof window.external.GetContext === 'undefined') {
            var hostUrl = OSF.getClientEndPoint()._targetUrl;
            var hostname_1 = getHostNameFromUrl(hostUrl);
            if (hostUrl.indexOf("https:") != 0) {
                return false;
            }
            OfficeCore.OfficeOnlineDomainList.forEach(function (domain) {
                if (domain.indexOf("*.") == 0) {
                    domain = domain.substring(2);
                }
                if (hostname_1.indexOf(domain) == hostname_1.length - domain.length) {
                    return true;
                }
            });
            return false;
        }
        return true;
    }
    OfficeCore.isHostOriginTrusted = isHostOriginTrusted;
    function getHostNameFromUrl(url) {
        var hostName = "";
        hostName = url.split("/")[2];
        hostName = hostName.split(":")[0];
        hostName = hostName.split("?")[0];
        return hostName;
    }
})(OfficeCore || (OfficeCore = {}));
var OfficeCore;
(function (OfficeCore) {
    var FirstPartyApis = (function () {
        function FirstPartyApis(context) {
            this.context = context;
        }
        Object.defineProperty(FirstPartyApis.prototype, "roamingSettings", {
            get: function () {
                if (!this.m_roamingSettings) {
                    this.m_roamingSettings = OfficeCore.AuthenticationService.newObject(this.context).roamingSettings;
                }
                return this.m_roamingSettings;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(FirstPartyApis.prototype, "tap", {
            get: function () {
                if (!this.m_tap) {
                    this.m_tap = OfficeCore.Tap.newObject(this.context);
                }
                return this.m_tap;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(FirstPartyApis.prototype, "skill", {
            get: function () {
                if (!this.m_skill) {
                    this.m_skill = OfficeCore.Skill.newObject(this.context);
                }
                return this.m_skill;
            },
            enumerable: true,
            configurable: true
        });
        return FirstPartyApis;
    }());
    OfficeCore.FirstPartyApis = FirstPartyApis;
    var RequestContext = (function (_super) {
        __extends(RequestContext, _super);
        function RequestContext(url) {
            return _super.call(this, url) || this;
        }
        Object.defineProperty(RequestContext.prototype, "firstParty", {
            get: function () {
                if (!this.m_firstPartyApis) {
                    this.m_firstPartyApis = new FirstPartyApis(this);
                }
                return this.m_firstPartyApis;
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
        Object.defineProperty(RequestContext.prototype, "telemetry", {
            get: function () {
                if (!this.m_telemetry) {
                    this.m_telemetry = OfficeCore.TelemetryService.newObject(this);
                }
                return this.m_telemetry;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(RequestContext.prototype, "bi", {
            get: function () {
                if (!this.m_biShim) {
                    this.m_biShim = OfficeCore.BiShim.newObject(this);
                }
                return this.m_biShim;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(RequestContext.prototype, "flightingService", {
            get: function () {
                if (!this.m_flightingService) {
                    this.m_flightingService = OfficeCore.FlightingService.newObject(this);
                }
                return this.m_flightingService;
            },
            enumerable: true,
            configurable: true
        });
        return RequestContext;
    }(OfficeExtension.ClientRequestContext));
    OfficeCore.RequestContext = RequestContext;
    function run(arg1, arg2) {
        return OfficeExtension.ClientRequestContext._runBatch("OfficeCore.run", arguments, function (requestInfo) { return new OfficeCore.RequestContext(requestInfo); });
    }
    OfficeCore.run = run;
})(OfficeCore || (OfficeCore = {}));
var OfficeCore;
(function (OfficeCore) {
    var _hostName = "Office";
    var _defaultApiSetName = "OfficeSharedApi";
    var _createPropertyObject = OfficeExtension.BatchApiHelper.createPropertyObject;
    var _createMethodObject = OfficeExtension.BatchApiHelper.createMethodObject;
    var _createIndexerObject = OfficeExtension.BatchApiHelper.createIndexerObject;
    var _createRootServiceObject = OfficeExtension.BatchApiHelper.createRootServiceObject;
    var _createTopLevelServiceObject = OfficeExtension.BatchApiHelper.createTopLevelServiceObject;
    var _createChildItemObject = OfficeExtension.BatchApiHelper.createChildItemObject;
    var _invokeMethod = OfficeExtension.BatchApiHelper.invokeMethod;
    var _invokeEnsureUnchanged = OfficeExtension.BatchApiHelper.invokeEnsureUnchanged;
    var _invokeSetProperty = OfficeExtension.BatchApiHelper.invokeSetProperty;
    var _isNullOrUndefined = OfficeExtension.Utility.isNullOrUndefined;
    var _isUndefined = OfficeExtension.Utility.isUndefined;
    var _throwIfNotLoaded = OfficeExtension.Utility.throwIfNotLoaded;
    var _throwIfApiNotSupported = OfficeExtension.Utility.throwIfApiNotSupported;
    var _load = OfficeExtension.Utility.load;
    var _retrieve = OfficeExtension.Utility.retrieve;
    var _toJson = OfficeExtension.Utility.toJson;
    var _fixObjectPathIfNecessary = OfficeExtension.Utility.fixObjectPathIfNecessary;
    var _handleNavigationPropertyResults = OfficeExtension.Utility._handleNavigationPropertyResults;
    var _adjustToDateTime = OfficeExtension.Utility.adjustToDateTime;
    var _processRetrieveResult = OfficeExtension.Utility.processRetrieveResult;
    var _setMockData = OfficeExtension.Utility.setMockData;
    var _CC;
    (function (_CC) {
        function Skill_HostSkillEvent_EventArgsTransform(thisObj, args) {
            var transformedArgs = {
                type: args.type,
                data: args.data
            };
            return transformedArgs;
        }
        _CC.Skill_HostSkillEvent_EventArgsTransform = Skill_HostSkillEvent_EventArgsTransform;
    })(_CC = OfficeCore._CC || (OfficeCore._CC = {}));
    var SkillErrorCodes;
    (function (SkillErrorCodes) {
        SkillErrorCodes["generalException"] = "GeneralException";
    })(SkillErrorCodes = OfficeCore.SkillErrorCodes || (OfficeCore.SkillErrorCodes = {}));
    var _libraryMetadataSkillApi = { "version": "1.0.0",
        "name": "OfficeCore",
        "defaultApiSetName": "OfficeSharedApi",
        "hostName": "Office",
        "apiSets": [],
        "strings": ["Skill", "registerHostSkillEvent", "unregisterHostSkillEvent"],
        "enumTypes": [],
        "clientObjectTypes": [[1, 0, 0, 0, [["executeAction", 3, 2, 0, 5], ["notifyPaneEvent", 2, 2, 0, 5], [2, 0, 0, 0, 1], [3, 0, 0, 0, 1], ["testFireEvent", 0, 0, 0, 1]], 0, 0, 0, [["HostSkillEvent", 2, 0, "RichApiMessageEventType.skillRichApiMessageEvent", "", 2, 3]], "Microsoft.SkillApi.Skill", 4]] };
    var _builder = new OfficeExtension.LibraryBuilder({ metadata: _libraryMetadataSkillApi, targetNamespaceObject: OfficeCore });
})(OfficeCore || (OfficeCore = {}));
var OfficeCore;
(function (OfficeCore) {
    var _hostName = "OfficeCore";
    var _defaultApiSetName = "TelemetryApi";
    var _createPropertyObject = OfficeExtension.BatchApiHelper.createPropertyObject;
    var _createMethodObject = OfficeExtension.BatchApiHelper.createMethodObject;
    var _createIndexerObject = OfficeExtension.BatchApiHelper.createIndexerObject;
    var _createRootServiceObject = OfficeExtension.BatchApiHelper.createRootServiceObject;
    var _createTopLevelServiceObject = OfficeExtension.BatchApiHelper.createTopLevelServiceObject;
    var _createChildItemObject = OfficeExtension.BatchApiHelper.createChildItemObject;
    var _invokeMethod = OfficeExtension.BatchApiHelper.invokeMethod;
    var _invokeEnsureUnchanged = OfficeExtension.BatchApiHelper.invokeEnsureUnchanged;
    var _invokeSetProperty = OfficeExtension.BatchApiHelper.invokeSetProperty;
    var _isNullOrUndefined = OfficeExtension.Utility.isNullOrUndefined;
    var _isUndefined = OfficeExtension.Utility.isUndefined;
    var _throwIfNotLoaded = OfficeExtension.Utility.throwIfNotLoaded;
    var _throwIfApiNotSupported = OfficeExtension.Utility.throwIfApiNotSupported;
    var _load = OfficeExtension.Utility.load;
    var _retrieve = OfficeExtension.Utility.retrieve;
    var _toJson = OfficeExtension.Utility.toJson;
    var _fixObjectPathIfNecessary = OfficeExtension.Utility.fixObjectPathIfNecessary;
    var _handleNavigationPropertyResults = OfficeExtension.Utility._handleNavigationPropertyResults;
    var _adjustToDateTime = OfficeExtension.Utility.adjustToDateTime;
    var _processRetrieveResult = OfficeExtension.Utility.processRetrieveResult;
    var _typeTelemetryService = "TelemetryService";
    var TelemetryService = (function (_super) {
        __extends(TelemetryService, _super);
        function TelemetryService() {
            return _super !== null && _super.apply(this, arguments) || this;
        }
        Object.defineProperty(TelemetryService.prototype, "_className", {
            get: function () {
                return "TelemetryService";
            },
            enumerable: true,
            configurable: true
        });
        TelemetryService.prototype.sendTelemetryEvent = function (telemetryProperties, eventName, eventContract, eventFlags, value) {
            _invokeMethod(this, "SendTelemetryEvent", 1, [telemetryProperties, eventName, eventContract, eventFlags, value], 4, 0);
        };
        TelemetryService.prototype._handleResult = function (value) {
            _super.prototype._handleResult.call(this, value);
            if (_isNullOrUndefined(value))
                return;
            var obj = value;
            _fixObjectPathIfNecessary(this, obj);
        };
        TelemetryService.prototype._handleRetrieveResult = function (value, result) {
            _super.prototype._handleRetrieveResult.call(this, value, result);
            _processRetrieveResult(this, value, result);
        };
        TelemetryService.newObject = function (context) {
            return _createTopLevelServiceObject(OfficeCore.TelemetryService, context, "Microsoft.Telemetry.TelemetryService", false, 4);
        };
        TelemetryService.prototype.toJSON = function () {
            return _toJson(this, {}, {});
        };
        return TelemetryService;
    }(OfficeExtension.ClientObject));
    OfficeCore.TelemetryService = TelemetryService;
    var DataFieldType;
    (function (DataFieldType) {
        DataFieldType["unset"] = "Unset";
        DataFieldType["string"] = "String";
        DataFieldType["boolean"] = "Boolean";
        DataFieldType["int64"] = "Int64";
        DataFieldType["double"] = "Double";
    })(DataFieldType = OfficeCore.DataFieldType || (OfficeCore.DataFieldType = {}));
    var TelemetryErrorCodes;
    (function (TelemetryErrorCodes) {
        TelemetryErrorCodes["generalException"] = "GeneralException";
    })(TelemetryErrorCodes = OfficeCore.TelemetryErrorCodes || (OfficeCore.TelemetryErrorCodes = {}));
})(OfficeCore || (OfficeCore = {}));
var OfficeFirstPartyAuth;
(function (OfficeFirstPartyAuth) {
    var ErrorCode = (function () {
        function ErrorCode() {
        }
        ErrorCode.GetAuthContextAsyncMissing = "GetAuthContextAsyncMissing";
        ErrorCode.CannotGetAuthContext = "CannotGetAuthContext";
        ErrorCode.PackageNotLoaded = "PackageNotLoaded";
        ErrorCode.FailedToLoad = "FailedToLoad";
        return ErrorCode;
    }());
    var WebAuthReplyUrlsStorageKey = "officeWebAuthReplyUrls";
    var retrievedAuthContext = false;
    var errorMessage;
    OfficeFirstPartyAuth.debugging = false;
    function load(replyUrl) {
        return new OfficeExtension.CoreUtility.Promise(function (resolve, reject) {
            if (OSF.WebAuth && OSF._OfficeAppFactory.getHostInfo().hostPlatform == "web") {
                try {
                    if (!Office || !Office.context || !Office.context.webAuth) {
                        reject({ code: ErrorCode.GetAuthContextAsyncMissing, message: (Strings && Strings.OfficeOM.L_ImplicitGetAuthContextMissing) ? Strings.OfficeOM.L_ImplicitGetAuthContextMissing : "" });
                    }
                    Office.context.webAuth.getAuthContextAsync(function (result) {
                        if (result.status === "succeeded") {
                            retrievedAuthContext = true;
                            var authContext = result.value;
                            if (!authContext || authContext.isAnonymous) {
                                return false;
                            }
                            var isMsa = authContext.authorityType.toLowerCase() === 'msa';
                            OSF.WebAuth.config = {
                                idp: authContext.authorityType.toLowerCase(),
                                appIds: [isMsa ? (authContext.msaAppId) ? authContext.msaAppId : authContext.appId : authContext.appId],
                                authority: (OfficeFirstPartyAuth.authorityOverride && OfficeFirstPartyAuth.debugging) ? OfficeFirstPartyAuth.authorityOverride : authContext.authority,
                                redirectUri: (replyUrl) ? replyUrl : null,
                                upn: authContext.upn,
                                enableConsoleLogging: OfficeFirstPartyAuth.debugging,
                                telemetryInstance: 'otel',
                                telemetry: { HashedUserId: authContext.userId }
                            };
                            var succeeded = false;
                            var loadResult = OSF.WebAuth.load(function (loaded) {
                                if (loaded) {
                                    succeeded = true;
                                    resolve();
                                }
                                reject({ code: ErrorCode.PackageNotLoaded, message: (Strings && Strings.OfficeOM.L_ImplicitNotLoaded) ? Strings.OfficeOM.L_ImplicitNotLoaded : "" });
                            });
                            logLoadEvent(loadResult, succeeded);
                            var finalReplyUrl = (replyUrl) ? replyUrl : window.location.href.split("?")[0];
                            var replyUrls = sessionStorage.getItem(WebAuthReplyUrlsStorageKey);
                            if (replyUrls || replyUrls === "") {
                                replyUrls = finalReplyUrl;
                            }
                            else {
                                replyUrls += ", " + finalReplyUrl;
                            }
                            sessionStorage.setItem(WebAuthReplyUrlsStorageKey, replyUrls);
                        }
                        else {
                            retrievedAuthContext = false;
                            OSF.WebAuth.config = null;
                            errorMessage = JSON.stringify(result);
                            reject({ code: ErrorCode.FailedToLoad, message: errorMessage });
                        }
                    });
                }
                catch (e) {
                    retrievedAuthContext = false;
                    OSF.WebAuth.config = null;
                    errorMessage = e;
                    OSF.WebAuth.load(function (loaded) {
                        if (loaded) {
                            resolve();
                        }
                        reject({ code: ErrorCode.FailedToLoad, message: errorMessage });
                    });
                }
            }
            else {
                resolve();
            }
        });
    }
    OfficeFirstPartyAuth.load = load;
    function getAccessToken(options, behaviorOption) {
        if (OSF.WebAuth && OSF.WebAuth.loaded && OSF._OfficeAppFactory.getHostInfo().hostPlatform == "web") {
            return new OfficeExtension.CoreUtility.Promise(function (resolve, reject) {
                if (behaviorOption && behaviorOption.forceRefresh) {
                    OSF.WebAuth.clearCache();
                }
                var identityType = (OSF.WebAuth.config.idp.toLowerCase() == "msa")
                    ? OfficeCore.IdentityType.microsoftAccount
                    : OfficeCore.IdentityType.organizationAccount;
                if (OSF.WebAuth.config.appIds[0]) {
                    OSF.WebAuth.getToken(options.resource, OSF.WebAuth.config.appIds[0], OSF._OfficeAppFactory.getHostInfo().osfControlAppCorrelationId, (behaviorOption && behaviorOption.popup) ? behaviorOption.popup : null).then(function (result) {
                        logAcquireEvent(result, true, options.resource, (behaviorOption && behaviorOption.popup) ? behaviorOption.popup : null);
                        resolve({ accessToken: result.Token, tokenIdenityType: identityType });
                    }).catch(function (result) {
                        logAcquireEvent(result, false, options.resource, (behaviorOption && behaviorOption.popup) ? behaviorOption.popup : null, result.ErrorCode);
                        reject({ code: result.ErrorCode, message: result.ErrorMessage });
                    });
                }
            });
        }
        var context = new OfficeCore.RequestContext();
        var auth = OfficeCore.AuthenticationService.newObject(context);
        context._customData = "WacPartition";
        if (OSF._OfficeAppFactory.getHostInfo().hostPlatform == "web" && OSF._OfficeAppFactory.getHostInfo().hostType == "word") {
            var result_1 = auth.getAccessToken(options, null);
            return context.sync().then(function () { return result_1.value; });
        }
        else {
            return new OfficeExtension.CoreUtility.Promise(function (resolve, reject) {
                var handler = auth.onTokenReceived.add(function (arg) {
                    if (!OfficeExtension.CoreUtility.isNullOrUndefined(arg)) {
                        handler.remove();
                        context.sync().catch(function () {
                        });
                        if (arg.code == 0) {
                            resolve(arg.tokenValue);
                        }
                        else {
                            if (OfficeExtension.CoreUtility.isNullOrUndefined(arg.errorInfo)) {
                                reject({ code: arg.code });
                            }
                            else {
                                try {
                                    reject(JSON.parse(arg.errorInfo));
                                }
                                catch (e) {
                                    reject({ code: arg.code, message: arg.errorInfo });
                                }
                            }
                        }
                    }
                    return null;
                });
                context.sync()
                    .then(function () {
                    var apiResult = auth.getAccessToken(options, auth._targetId);
                    return context.sync()
                        .then(function () {
                        if (OfficeExtension.CoreUtility.isNullOrUndefined(apiResult.value)) {
                            return null;
                        }
                        var tokenValue = apiResult.value.accessToken;
                        if (!OfficeExtension.CoreUtility.isNullOrUndefined(tokenValue)) {
                            resolve(apiResult.value);
                        }
                    });
                })
                    .catch(function (e) {
                    reject(e);
                });
            });
        }
    }
    OfficeFirstPartyAuth.getAccessToken = getAccessToken;
    function getPrimaryIdentityInfo() {
        var context = new OfficeCore.RequestContext();
        var auth = OfficeCore.AuthenticationService.newObject(context);
        context._customData = "WacPartition";
        var result = auth.getPrimaryIdentityInfo();
        return context.sync().then(function () { return result.value; });
    }
    OfficeFirstPartyAuth.getPrimaryIdentityInfo = getPrimaryIdentityInfo;
    function getIdentities() {
        var context = new OfficeCore.RequestContext();
        var auth_service = OfficeCore.AuthenticationService.newObject(context);
        var result = auth_service.getIdentities();
        return context.sync().then(function () { return result.value; });
    }
    OfficeFirstPartyAuth.getIdentities = getIdentities;
    function logLoadEvent(result, succeeded) {
        if (OfficeFirstPartyAuth.debugging) {
            console.log("Logging Implicit load event");
        }
        if (typeof OTel !== "undefined") {
            OTel.OTelLogger.onTelemetryLoaded(function () {
                var telemetryData = [
                    oteljs.makeStringDataField('IdentityProvider', OSF.WebAuth.config.idp),
                    oteljs.makeStringDataField('AppId', OSF.WebAuth.config.appIds[0]),
                    oteljs.makeBooleanDataField('Js', typeof Implicit !== "undefined" ? true : false),
                    oteljs.makeBooleanDataField('Result', succeeded)
                ];
                if (OSF.WebAuth.config.telemetry) {
                    for (var key in OSF.WebAuth.config.telemetry) {
                        telemetryData.push(oteljs.makeStringDataField(key, OSF.WebAuth.config.telemetry[key]));
                    }
                }
                if (result && result.Telemetry) {
                    for (var key in result.Telemetry) {
                        if (!result.Telemetry[key]) {
                            continue;
                        }
                        switch (key) {
                            case 'succeeded':
                                telemetryData.push(oteljs.makeBooleanDataField(key, result.Telemetry[key]));
                                break;
                            case 'loadedApplicationCount':
                                telemetryData.push(oteljs.makeInt64DataField(key, result.Telemetry[key]));
                                break;
                            case 'timeToLoad':
                                telemetryData.push(oteljs.makeInt64DataField(key, result.Telemetry[key]));
                                break;
                            default:
                                telemetryData.push(oteljs.makeStringDataField(key, result.Telemetry[key]));
                        }
                    }
                }
                OTel.OTelLogger.sendTelemetryEvent({
                    eventName: "Office.Extensibility.OfficeJs.OfficeFirstPartyAuth.Load",
                    dataFields: telemetryData,
                    eventFlags: {
                        dataCategories: oteljs.DataCategories.ProductServiceUsage
                    }
                });
            });
        }
    }
    function logAcquireEvent(result, succeeded, target, popup, message) {
        if (OfficeFirstPartyAuth.debugging) {
            console.log("Logging Implicit acquire event");
        }
        if (typeof OTel !== "undefined") {
            OTel.OTelLogger.onTelemetryLoaded(function () {
                var telemetryData = [
                    oteljs.makeStringDataField('IdentityProvider', OSF.WebAuth.config.idp),
                    oteljs.makeStringDataField('AppId', OSF.WebAuth.config.appIds[0]),
                    oteljs.makeStringDataField('Target', target),
                    oteljs.makeBooleanDataField('Popup', (typeof popup === "boolean") ? popup : false),
                    oteljs.makeBooleanDataField('Result', succeeded),
                    oteljs.makeStringDataField('Error', message)
                ];
                if (OSF.WebAuth.config.telemetry) {
                    for (var key in OSF.WebAuth.config.telemetry) {
                        telemetryData.push(oteljs.makeStringDataField(key, OSF.WebAuth.config.telemetry[key]));
                    }
                }
                if (result && result.Telemetry) {
                    for (var key in result.Telemetry) {
                        if (!result.Telemetry[key]) {
                            continue;
                        }
                        switch (key) {
                            case 'succeeded':
                                telemetryData.push(oteljs.makeBooleanDataField(key, result.Telemetry[key]));
                                break;
                            case 'timeToGetToken':
                                telemetryData.push(oteljs.makeInt64DataField(key, result.Telemetry[key]));
                                break;
                            default:
                                telemetryData.push(oteljs.makeStringDataField(key, result.Telemetry[key]));
                        }
                    }
                }
                OTel.OTelLogger.sendTelemetryEvent({
                    eventName: "Office.Extensibility.OfficeJs.OfficeFirstPartyAuth.GetAccessToken",
                    dataFields: telemetryData,
                    eventFlags: {
                        dataCategories: oteljs.DataCategories.ProductServiceUsage
                    }
                });
            });
        }
    }
    function loadWebAuthForReplyPage() {
        if (!window.sessionStorage) {
            return;
        }
        var webAuthRedirectUrls = sessionStorage.getItem(WebAuthReplyUrlsStorageKey);
        if (webAuthRedirectUrls !== null && webAuthRedirectUrls.indexOf(window.location.origin + window.location.pathname) !== -1) {
            load();
        }
    }
    if (window.OSF) {
        loadWebAuthForReplyPage();
    }
})(OfficeFirstPartyAuth || (OfficeFirstPartyAuth = {}));
var OfficeCore;
(function (OfficeCore) {
    var _hostName = "Office";
    var _defaultApiSetName = "OfficeSharedApi";
    var _createPropertyObject = OfficeExtension.BatchApiHelper.createPropertyObject;
    var _createMethodObject = OfficeExtension.BatchApiHelper.createMethodObject;
    var _createIndexerObject = OfficeExtension.BatchApiHelper.createIndexerObject;
    var _createRootServiceObject = OfficeExtension.BatchApiHelper.createRootServiceObject;
    var _createTopLevelServiceObject = OfficeExtension.BatchApiHelper.createTopLevelServiceObject;
    var _createChildItemObject = OfficeExtension.BatchApiHelper.createChildItemObject;
    var _invokeMethod = OfficeExtension.BatchApiHelper.invokeMethod;
    var _invokeEnsureUnchanged = OfficeExtension.BatchApiHelper.invokeEnsureUnchanged;
    var _invokeSetProperty = OfficeExtension.BatchApiHelper.invokeSetProperty;
    var _isNullOrUndefined = OfficeExtension.Utility.isNullOrUndefined;
    var _isUndefined = OfficeExtension.Utility.isUndefined;
    var _throwIfNotLoaded = OfficeExtension.Utility.throwIfNotLoaded;
    var _throwIfApiNotSupported = OfficeExtension.Utility.throwIfApiNotSupported;
    var _load = OfficeExtension.Utility.load;
    var _retrieve = OfficeExtension.Utility.retrieve;
    var _toJson = OfficeExtension.Utility.toJson;
    var _fixObjectPathIfNecessary = OfficeExtension.Utility.fixObjectPathIfNecessary;
    var _handleNavigationPropertyResults = OfficeExtension.Utility._handleNavigationPropertyResults;
    var _adjustToDateTime = OfficeExtension.Utility.adjustToDateTime;
    var _processRetrieveResult = OfficeExtension.Utility.processRetrieveResult;
    var _setMockData = OfficeExtension.Utility.setMockData;
    var AuthenticationServiceCustom = (function () {
        function AuthenticationServiceCustom() {
        }
        Object.defineProperty(AuthenticationServiceCustom.prototype, "_targetId", {
            get: function () {
                if (this.m_targetId == undefined) {
                    if (typeof (OSF) !== 'undefined' && OSF.OUtil) {
                        this.m_targetId = OSF.OUtil.Guid.generateNewGuid();
                    }
                    else {
                        this.m_targetId = "" + this.context._nextId();
                    }
                }
                return this.m_targetId;
            },
            enumerable: true,
            configurable: true
        });
        return AuthenticationServiceCustom;
    }());
    OfficeCore.AuthenticationServiceCustom = AuthenticationServiceCustom;
    var _CC;
    (function (_CC) {
        function AuthenticationService_TokenReceived_EventArgsTransform(thisObj, args) {
            var value = args;
            var newArgs = {
                tokenValue: value.tokenValue,
                code: value.code,
                errorInfo: value.errorInfo
            };
            return newArgs;
        }
        _CC.AuthenticationService_TokenReceived_EventArgsTransform = AuthenticationService_TokenReceived_EventArgsTransform;
    })(_CC = OfficeCore._CC || (OfficeCore._CC = {}));
    var PersonaPromiseType;
    (function (PersonaPromiseType) {
        PersonaPromiseType[PersonaPromiseType["immediate"] = 0] = "immediate";
        PersonaPromiseType[PersonaPromiseType["load"] = 3] = "load";
    })(PersonaPromiseType = OfficeCore.PersonaPromiseType || (OfficeCore.PersonaPromiseType = {}));
    var PersonaInfoAndSource = (function () {
        function PersonaInfoAndSource() {
        }
        return PersonaInfoAndSource;
    }());
    OfficeCore.PersonaInfoAndSource = PersonaInfoAndSource;
    ;
    var PersonaCustom = (function () {
        function PersonaCustom() {
        }
        PersonaCustom.prototype.performAsyncOperation = function (type, waitFor, action, check) {
            var _this = this;
            if (type == PersonaPromiseType.immediate) {
                action();
                return;
            }
            check().then(function (isWarmedUp) {
                if (isWarmedUp) {
                    action();
                }
                else {
                    var persona = _this;
                    persona.load("hostId");
                    persona.context.sync().then(function () {
                        var hostId = persona.hostId;
                        _this.getPersonaLifetime().then(function (personaLifetime) {
                            var eventHandler = function (args) {
                                return new OfficeExtension.CoreUtility.Promise(function (resolve, reject) {
                                    if (args.sendingPersonaHostId == hostId) {
                                        for (var index = 0; index < args.dataUpdated.length; ++index) {
                                            var updated = args.dataUpdated[index];
                                            if (waitFor == updated) {
                                                check().then(function (isWarmedUp) {
                                                    if (isWarmedUp) {
                                                        action();
                                                        personaLifetime.onPersonaUpdated.remove(eventHandler);
                                                        persona.context.sync();
                                                    }
                                                    resolve(isWarmedUp);
                                                });
                                                return;
                                            }
                                        }
                                    }
                                    resolve(false);
                                });
                            };
                            personaLifetime.onPersonaUpdated.add(eventHandler);
                            persona.context.sync();
                        });
                    });
                }
            });
        };
        PersonaCustom.prototype.getOrganizationAsync = function (type) {
            var _this = this;
            return new OfficeExtension.CoreUtility.Promise(function (resolve, reject) {
                var persona = _this;
                var action = function () {
                    var organization = persona.organization;
                    organization.load("*");
                    persona.context.sync().then(function () {
                        resolve(organization);
                    });
                };
                var check = function () {
                    return new OfficeExtension.CoreUtility.Promise(function (isWarmedUpResolve, isWarmedUpReject) {
                        var organization = persona.organization;
                        organization.load("isWarmedUp");
                        persona.context.sync().then(function () {
                            isWarmedUpResolve(organization.isWarmedUp);
                        });
                    });
                };
                _this.performAsyncOperation(type, PersonaDataUpdated.organization, action, check);
            });
        };
        PersonaCustom.prototype.getIsPersonaInfoResolvedCheck = function () {
            var persona = this;
            return new OfficeExtension.CoreUtility.Promise(function (resolve, reject) {
                var info = persona.personaInfo;
                info.load("isPersonResolved");
                persona.context.sync().then(function () {
                    resolve(info.isPersonResolved);
                });
            });
        };
        PersonaCustom.prototype.getPersonaInfoAsync = function (type) {
            var _this = this;
            return new OfficeExtension.CoreUtility.Promise(function (resolve, reject) {
                var persona = _this;
                var action = function () {
                    var info = persona.personaInfo;
                    info.load();
                    persona.context.sync().then(function () {
                        resolve(info);
                    });
                };
                var check = function () {
                    return _this.getIsPersonaInfoResolvedCheck();
                };
                _this.performAsyncOperation(type, PersonaDataUpdated.personaInfo, action, check);
            });
        };
        PersonaCustom.prototype.getPersonaInfoWithSourceAsync = function (type) {
            var _this = this;
            return new OfficeExtension.CoreUtility.Promise(function (resolve, reject) {
                var persona = _this;
                var action = function () {
                    var result = new PersonaInfoAndSource();
                    result.info = persona.personaInfo;
                    result.info.load();
                    result.source = persona.personaInfo.sources;
                    result.source.load();
                    persona.context.sync().then(function () {
                        resolve(result);
                    });
                };
                var check = function () {
                    return _this.getIsPersonaInfoResolvedCheck();
                };
                _this.performAsyncOperation(type, PersonaDataUpdated.personaInfo, action, check);
            });
        };
        PersonaCustom.prototype.getUnifiedCommunicationInfo = function (type) {
            var _this = this;
            return new OfficeExtension.CoreUtility.Promise(function (resolve, reject) {
                var persona = _this;
                var action = function () {
                    var ucInfo = persona.unifiedCommunicationInfo;
                    ucInfo.load("*");
                    persona.context.sync().then(function () {
                        resolve(ucInfo);
                    });
                };
                var check = function () {
                    return _this.getIsPersonaInfoResolvedCheck();
                };
                _this.performAsyncOperation(type, PersonaDataUpdated.personaInfo, action, check);
            });
        };
        PersonaCustom.prototype.getUnifiedGroupInfoAsync = function (type) {
            var _this = this;
            return new OfficeExtension.CoreUtility.Promise(function (resolve, reject) {
                var persona = _this;
                var action = function () {
                    var group = persona.unifiedGroupInfo;
                    group.load("*");
                    persona.context.sync().then(function () {
                        resolve(group);
                    });
                };
                var check = function () {
                    return _this.getIsPersonaInfoResolvedCheck();
                };
                _this.performAsyncOperation(type, PersonaDataUpdated.personaInfo, action, check);
            });
        };
        PersonaCustom.prototype.getTypeAsync = function (type) {
            var _this = this;
            return new OfficeExtension.CoreUtility.Promise(function (resolve, reject) {
                var persona = _this;
                var action = function () {
                    persona.load("type");
                    persona.context.sync().then(function () {
                        resolve(OfficeCore.PersonaType[persona.type.valueOf()]);
                    });
                };
                var check = function () {
                    return _this.getIsPersonaInfoResolvedCheck();
                };
                _this.performAsyncOperation(type, PersonaDataUpdated.personaInfo, action, check);
            });
        };
        PersonaCustom.prototype.getCustomizationsAsync = function (type) {
            var _this = this;
            return new OfficeExtension.CoreUtility.Promise(function (resolve, reject) {
                var persona = _this;
                var action = function () {
                    var customizations = persona.getCustomizations();
                    persona.context.sync().then(function () {
                        resolve(customizations.value);
                    });
                };
                var check = function () {
                    return _this.getIsPersonaInfoResolvedCheck();
                };
                _this.performAsyncOperation(type, PersonaDataUpdated.personaInfo, action, check);
            });
        };
        PersonaCustom.prototype.getMembersAsync = function (type) {
            var _this = this;
            return new OfficeExtension.CoreUtility.Promise(function (resolve, rejcet) {
                var persona = _this;
                var action = function () {
                    var members = persona.getMembers();
                    members.load("isWarmedUp");
                    persona.context.sync().then(function () {
                        resolve(members);
                    });
                };
                var check = function () {
                    return new OfficeExtension.CoreUtility.Promise(function (isWarmedUpResolve, isWarmedUpReject) {
                        var members = persona.getMembers();
                        members.load("isWarmedUp");
                        persona.context.sync().then(function () {
                            isWarmedUpResolve(members.isWarmedUp);
                        });
                    });
                };
                _this.performAsyncOperation(type, PersonaDataUpdated.members, action, check);
            });
        };
        PersonaCustom.prototype.getMembershipAsync = function (type) {
            var _this = this;
            return new OfficeExtension.CoreUtility.Promise(function (resolve, reject) {
                var persona = _this;
                var action = function () {
                    var membership = persona.getMembership();
                    membership.load("*");
                    persona.context.sync().then(function () {
                        resolve(membership);
                    });
                };
                var check = function () {
                    return new OfficeExtension.CoreUtility.Promise(function (isWarmedUpResolve) {
                        var membership = persona.getMembership();
                        membership.load("isWarmedUp");
                        persona.context.sync().then(function () {
                            isWarmedUpResolve(membership.isWarmedUp);
                        });
                    });
                };
                _this.performAsyncOperation(type, PersonaDataUpdated.membership, action, check);
            });
        };
        PersonaCustom.prototype.getPersonaLifetime = function () {
            var _this = this;
            return new OfficeExtension.CoreUtility.Promise(function (resolve, reject) {
                var persona = _this;
                persona.load("instanceId");
                persona.context.sync().then(function () {
                    var peopleApi = new PeopleApiContext(persona.context, persona.instanceId);
                    peopleApi.getPersonaLifetime().then(function (lifetime) {
                        resolve(lifetime);
                    });
                });
            });
        };
        return PersonaCustom;
    }());
    OfficeCore.PersonaCustom = PersonaCustom;
    var PeopleApiContext = (function () {
        function PeopleApiContext(context, instanceId) {
            this.context = context;
            this.instanceId = instanceId;
        }
        Object.defineProperty(PeopleApiContext.prototype, "serviceContext", {
            get: function () {
                if (!this.m_serviceConext) {
                    this.m_serviceConext = OfficeCore.ServiceContext.newObject(this.context);
                }
                return this.m_serviceConext;
            },
            enumerable: true,
            configurable: true
        });
        PeopleApiContext.prototype.getPersonaLifetime = function () {
            var _this = this;
            return new OfficeExtension.CoreUtility.Promise(function (resolve, reject) {
                var lifetime = _this.serviceContext.getPersonaLifetime(_this.instanceId);
                _this.context.sync().then(function () {
                    lifetime.load("instanceId");
                    _this.context.sync().then(function () {
                        resolve(lifetime);
                    });
                });
            });
        };
        PeopleApiContext.prototype.getInitialPersona = function () {
            var _this = this;
            return new OfficeExtension.CoreUtility.Promise(function (resolve, reject) {
                var persona = _this.serviceContext.getInitialPersona(_this.instanceId);
                _this.context.sync().then(function () {
                    resolve(persona);
                });
            });
        };
        PeopleApiContext.prototype.getLokiTokenProvider = function () {
            var _this = this;
            return new OfficeExtension.CoreUtility.Promise(function (resolve, reject) {
                var provider = _this.serviceContext.getLokiTokenProvider(_this.instanceId);
                _this.context.sync().then(function () {
                    provider.load("instanceId");
                    _this.context.sync().then(function () {
                        resolve(provider);
                    });
                });
            });
        };
        return PeopleApiContext;
    }());
    OfficeCore.PeopleApiContext = PeopleApiContext;
    (function (_CC) {
        function AppRuntimeService_VisibilityChanged_EventArgsTransform(thisObj, args) {
            var value = args;
            if (value && value.message) {
                value.visibility = value.message;
            }
            return value;
        }
        _CC.AppRuntimeService_VisibilityChanged_EventArgsTransform = AppRuntimeService_VisibilityChanged_EventArgsTransform;
    })(_CC = OfficeCore._CC || (OfficeCore._CC = {}));
    var ErrorCodes;
    (function (ErrorCodes) {
        ErrorCodes["apiNotAvailable"] = "ApiNotAvailable";
        ErrorCodes["clientError"] = "ClientError";
        ErrorCodes["controlIdNotFound"] = "ControlIdNotFound";
        ErrorCodes["entryIdRequired"] = "EntryIdRequired";
        ErrorCodes["generalException"] = "GeneralException";
        ErrorCodes["instanceNotFound"] = "InstanceNotFound";
        ErrorCodes["interactiveFlowAborted"] = "InteractiveFlowAborted";
        ErrorCodes["invalidArgument"] = "InvalidArgument";
        ErrorCodes["invalidGrant"] = "InvalidGrant";
        ErrorCodes["invalidResourceUrl"] = "InvalidResourceUrl";
        ErrorCodes["noPhoto"] = "NoPhoto";
        ErrorCodes["objectNotFound"] = "ObjectNotFound";
        ErrorCodes["resourceNotSupported"] = "ResourceNotSupported";
        ErrorCodes["serverError"] = "ServerError";
        ErrorCodes["ticketInvalidParams"] = "TicketInvalidParams";
        ErrorCodes["ticketNetworkError"] = "TicketNetworkError";
        ErrorCodes["ticketUnauthorized"] = "TicketUnauthorized";
        ErrorCodes["ticketUninitialized"] = "TicketUninitialized";
        ErrorCodes["ticketUnknownError"] = "TicketUnknownError";
        ErrorCodes["unexpectedError"] = "UnexpectedError";
        ErrorCodes["unsupportedUserIdentity"] = "UnsupportedUserIdentity";
        ErrorCodes["userNotSignedIn"] = "UserNotSignedIn";
    })(ErrorCodes = OfficeCore.ErrorCodes || (OfficeCore.ErrorCodes = {}));
    var Interfaces;
    (function (Interfaces) {
    })(Interfaces = OfficeCore.Interfaces || (OfficeCore.Interfaces = {}));
    var _libraryMetadataOfficeSharedApi = { "version": "1.0.0",
        "name": "OfficeCore",
        "defaultApiSetName": "OfficeSharedApi",
        "hostName": "Office",
        "apiSets": [["1.2", "FirstPartyAuthentication"], ["1.3", "FirstPartyAuthentication"]],
        "strings": ["AuthenticationService", "RoamingSetting", "RoamingSettingCollection", "LinkedIn", "NetworkUsage", "DynamicRibbon", "RibbonTab", "RibbonButton", "RibbonButtonCollection", "OfficeServicesManagerApi", "Comment", "CommentCollection", "MemberInfoList", "PersonaActions", "PersonaInfoSource", "PersonaInfo", "PersonaUnifiedCommunicationInfo", "PersonaPhotoInfo", "PersonaCollection", "PersonaOrganizationInfo", "UnifiedGroupInfo", "Persona", "PersonaLifetime", "LokiTokenProvider", "LokiTokenProviderFactory", "ServiceContext", "Tap", "AppRuntimePersistenceService", "AppRuntimeService", "null", "id", "getItem", "getCount", "isWarmedUp", "isWarmingUp", "displayName", "email", "emailAddresses", "sipAddresses", "birthday", "birthdays", "title", "jobInfoDepartment", "companyName", "office", "linkedTitles", "linkedDepartments", "linkedCompanyNames", "linkedOffices", "webSites", "notes", "getImageUri", "instanceId", "dispose", "_RegisterPersonaUpdatedEvent", "_UnregisterPersonaUpdatedEvent", "this.instanceId", "_RegisterLokiTokenAvailableEvent", "_UnregisterLokiTokenAvailableEvent", "_RegisterIdentityUniqueIdAvailableEvent", "_UnregisterIdentityUniqueIdAvailableEvent", "getLokiTokenProvider"],
        "enumTypes": [["IdentityType", ["organizationAccount", "microsoftAccount", "unsupported"]],
            ["CommentTextFormat", ["plain", "markdown", "delta"]],
            ["PersonaCardPerfPoint", ["placeHolderRendered", "initialCardRendered"]],
            ["MessageType", [], { "personaLifetimePersonaUpdatedEvent": 3502, "lokiTokenProviderLokiTokenAvailableEvent": 3503, "lokiTokenProviderIdentityUniqueIdAvailableEvent": 3504 }],
            ["UnifiedCommunicationAvailability", ["notSet", "free", "idle", "busy", "idleBusy", "doNotDisturb", "unalertable", "unavailable"]],
            ["UnifiedCommunicationStatus", ["online", "notOnline", "away", "busy", "beRightBack", "onThePhone", "outToLunch", "inAMeeting", "outOfOffice", "doNotDisturb", "inAConference", "getting", "notABuddy", "disconnected", "notInstalled", "urgentInterruptionsOnly", "mayBeAvailable", "idle", "inPresentation"]],
            ["FreeBusyCalendarState", ["unknown", "free", "busy", "elsewhere", "tentative", "outOfOffice"]],
            ["PersonaType", ["unknown", "enterprise", "contact", "bot", "phoneOnly", "oneOff", "distributionList", "personalDistributionList", "anonymous", "unifiedGroup"]],
            ["PhoneType", ["workPhone", "homePhone", "mobilePhone", "businessFax", "otherPhone"]],
            ["AddressType", ["workAddress", "homeAddress", "otherAddress"]],
            ["MemberType", ["unknown", "individual", "group"]],
            ["PersonaDataUpdated", ["hostId", "type", "photo", "personaInfo", "unifiedCommunicationInfo", "organization", "unifiedGroupInfo", "members", "membership", "capabilities", "customizations", "viewableSources"]],
            ["CustomizedData", ["email", "workPhone", "workPhone2", "workFax", "mobilePhone", "homePhone", "homePhone2", "otherPhone", "sipAddress", "profile", "office", "company", "workAddress", "homeAddress", "otherAddress", "birthday"]],
            ["ObjectType", ["unknown", "chart", "smartArt", "table", "image", "slide", "text"], { "ole": "OLE" }],
            ["AppRuntimeState", ["inactive", "background", "visible"]],
            ["Visibility", ["hidden", "visible"]]],
        "clientObjectTypes": [[1, 4, 0, [["roamingSettings", 3, 2, 0, 0, 4]], [["getAccessToken", 2, 2, 0, 5], ["getPrimaryIdentityInfo", 0, 2, 1, 5], ["getIdentities", 0, 2, 2, 5]], 0, 0, 0, [["TokenReceived", 2, 1, 3001, "this._targetId", 30, 30]], "Microsoft.Authentication.AuthenticationService", 4],
            [2, 0, [[31, 3], ["value", 1]]],
            [3, 0, 0, 0, 0, [[32, 2, 1, 2, 0, 4], ["getItemOrNullObject", 2, 1, 2, 0, 4]]],
            [4, 0, 0, 0, [["isEnabledForOffice"], ["recordLinkedInSettingsCompliance", 2]], 0, 0, 0, 0, "Microsoft.DesktopCompliance.LinkedIn", 4],
            [5, 0, 0, 0, [["isInOnlineMode"]], 0, 0, 0, 0, "Microsoft.DesktopCompliance.NetworkUsage", 4],
            [6, 0, 0, [["buttons", 9, 19, 0, 0, 4]], 0, [["getButton", 8, 1], ["getTab", 7, 1], ["executeRequestUpdate", 8, 1]], 0, 0, 0, "Microsoft.DynamicRibbon.DynamicRibbon", 4],
            [7, 0, [[31, 3]], 0, [["setVisibility", 1]]],
            [8, 0, [[31, 3], ["enabled", 1], ["label", 3]], 0, [["setEnabled", 1]]],
            [9, 1, 0, 0, [[33, 0, 2, 0, 4]], [[32, 8, 1, 18, 0, 4]], 0, 8],
            [10, 0, 0, 0, [["bindServiceToProfile", 2]], 0, 0, 0, 0, "Microsoft.OfficeServicesManager.OfficeServicesManagerApi", 4],
            [11, 0, [[31, 3], ["text", 1], ["created", 11], ["level", 3], ["resolved", 1], ["author", 3], ["mentions", 3]], [["parent", 11, 2, 0, 0, 4], ["parentOrNullObject", 11, 2, 0, 0, 4], ["replies", 12, 19, 0, 0, 4]], [["getRichText", 1, 2, 0, 4], ["setRichText", 2], ["delete"]], [["getParentOrSelf", 11, 0, 2, 0, 4], ["reply", 11, 2]]],
            [12, 1, 0, 0, [[33, 0, 2, 0, 4]], [[32, 11, 1, 18, 0, 4]], 0, 11],
            [13, 0, [[34, 3], [35, 3]], 0, [["items"]], [["getPersonaForMember", 22, 1]]],
            [14, 0, 0, 0, [["addContact"], ["editContact"], ["composeEmail", 1], ["composeInstantMessage", 1], ["callPhoneNumber", 1], ["pinPersonaToQuickContacts"], ["toggleTagForAlerts"], ["scheduleMeeting"], ["openLinkContactUx"], ["editContactByIdentifier", 1], ["showHoverCardForPersona", 6], ["hideHoverCardForPersona"], ["showContextMenu", 6], ["showContactCard", 6], ["showExpandedCard", 6]]],
            [15, 0, [[36, 3], [37, 3], [38, 3], [39, 3], [40, 3], [41, 3], [42, 3], [43, 3], [44, 3], [45, 3], [46, 3], [47, 3], [48, 3], [49, 3], ["phones", 3], ["addresses", 3], [50, 3], [51, 3]]],
            [16, 0, [[36, 3], [37, 3], [38, 3], [39, 3], [40, 11], [41, 11], [42, 3], [43, 3], [44, 3], [45, 3], [46, 3], [47, 3], [48, 3], [49, 3], [50, 3], [51, 3], ["isPersonResolved", 3]], [["sources", 15, 3, 0, 0, 4]], [["getPhones", 0, 2, 0, 4], ["getAddresses", 0, 2, 0, 4]]],
            [17, 0, [["availability", 3], ["status", 3], ["isSelf", 3], ["isTagged", 3], ["customStatusString", 3], ["isBlocked", 3], ["presenceTooltip", 3], ["isOutOfOffice", 3], ["outOfOfficeNote", 3], ["timezone", 3], ["meetingLocation", 3], ["meetingSubject", 3], ["timezoneBias", 3], ["idleStartTime", 11], ["overallCapability", 3], ["isOnBuddyList", 3], ["presenceNote", 3], ["voiceMailUri", 3], ["availabilityText", 3], ["availabilityTooltip", 3], ["isDurationInAvailabilityText", 3], ["freeBusyStatus", 3], ["calendarState", 3]]],
            [18, 0, 0, 0, [[52, 1, 0, 0, 0, 52]]],
            [19, 1, 0, 0, [[33, 0, 2, 0, 4]], [[32, 22, 1, 18, 0, 4]], 0, 22],
            [20, 0, [[34, 3], [35, 3]], [["hierarchy", 19, 18, 0, 0, 4], ["manager", 22, 2, 0, 0, 4], ["directReports", 19, 18, 0, 0, 4]]],
            [21, 0, [["description", 1], ["oneDrive", 1], ["oneNote", 1], ["isPublic", 1], ["amIOwner", 1], ["amIMember", 1], ["amISubscribed", 1], ["memberCount", 1], ["ownerCount", 1], ["hasGuests", 1], ["site", 1], ["planner", 1], ["classification", 1], ["subscriptionEnabled", 1]]],
            [22, 4, [["hostId", 3], ["type", 3], ["capabilities", 3], ["diagnosticId", 3], [53, 3]], [["photo", 18, 3, 0, 0, 4], ["personaInfo", 16, 3, 0, 0, 4], ["unifiedCommunicationInfo", 17, 3, 0, 0, 4], ["organization", 20, 3, 0, 0, 4], ["unifiedGroupInfo", 21, 35, 0, 0, 4], ["actions", 14, 2, 0, 0, 4]], [["getCustomizations"], ["warmup", 1], [54], ["getViewableSources"], ["reportTimeForRender", 2]], [["getMembers", 13], ["getMembership", 13]]],
            [23, 0, [[53, 3]], 0, [["getPolicies"], [55], [56]], [["getPersona", 22, 1], ["getPersonaForOrgEntry", 22, 4], ["getPersonaForOrgByEntryId", 22, 4]], 0, 0, [["PersonaUpdated", 0, 0, "MessageType.personaLifetimePersonaUpdatedEvent", 57, 55, 56]]],
            [24, 0, [["emailOrUpn", 3], [53, 3]], 0, [["requestToken"], [58], [59], ["requestIdentityUniqueId"], [60], [61]], 0, 0, 0, [["IdentityUniqueIdAvailable", 0, 0, "MessageType.lokiTokenProviderIdentityUniqueIdAvailableEvent", 57, 60, 61], ["LokiTokenAvailable", 0, 0, "MessageType.lokiTokenProviderLokiTokenAvailableEvent", 57, 58, 59]]],
            [25, 0, 0, 0, 0, [[62, 24, 1]], 0, 0, 0, "Microsoft.People.LokiTokenProviderFactory", 4],
            [26, 0, 0, 0, [[54, 1], ["accountEmailOrUpn", 1], ["getPersonaPolicies"]], [[62, 24, 1], ["getPersonaLifetime", 23, 1], ["getInitialPersona", 22, 1]], 0, 0, 0, "Microsoft.People.ServiceContext", 4],
            [27, 0, 0, 0, [["getEnterpriseUserInfo", 0, 2, 0, 5], ["getMruFriendlyPath", 1, 2, 0, 5], ["launchFileUrlInOfficeApp", 2, 2, 0, 5], ["performLocalSearch", 4, 2, 0, 5], ["readSearchCache", 3, 2, 0, 5], ["writeSearchCache", 3, 2, 0, 5]], 0, 0, 0, 0, "Microsoft.TapRichApi.Tap", 4],
            [28, 0, 0, 0, [["setAppRuntimeStartState", 1], ["getAppRuntimeStartState", 0, 2, 0, 4]], 0, 0, 0, 0, "Microsoft.AppRuntime.AppRuntimePersistenceService", 4],
            [29, 0, 0, 0, [["setAppRuntimeState", 1], ["getAppRuntimeState", 0, 2, 0, 4]], 0, 0, 0, [["VisibilityChanged", 2, 0, "65539", "", 30, 30]], "Microsoft.AppRuntime.AppRuntimeService", 4]] };
    var _builder = new OfficeExtension.LibraryBuilder({ metadata: _libraryMetadataOfficeSharedApi, targetNamespaceObject: OfficeCore });
})(OfficeCore || (OfficeCore = {}));
var Office;
(function (Office) {
    Office.Visibility = OfficeCore.Visibility;
    var StartupBehavior;
    (function (StartupBehavior) {
        StartupBehavior["none"] = "None";
        StartupBehavior["load"] = "Load";
    })(StartupBehavior = Office.StartupBehavior || (Office.StartupBehavior = {}));
    var addin;
    (function (addin) {
        var _this = this;
        function _createRequestContext(wacPartition) {
            var context = new OfficeCore.RequestContext();
            context._requestFlagModifier |= 64;
            if (wacPartition) {
                context._customData = 'WacPartition';
            }
            return context;
        }
        function setStartupBehavior(behavior) {
            return __awaiter(this, void 0, void 0, function () {
                var state, context, appRuntimePersistenceService;
                return __generator(this, function (_a) {
                    switch (_a.label) {
                        case 0:
                            if (behavior !== StartupBehavior.load && behavior !== StartupBehavior.none) {
                                throw OfficeExtension.Utility.createRuntimeError(OfficeExtension.ErrorCodes.invalidArgument, null, null);
                            }
                            state = (behavior == StartupBehavior.load ? OfficeCore.AppRuntimeState.background : OfficeCore.AppRuntimeState.inactive);
                            context = _createRequestContext(false);
                            appRuntimePersistenceService = OfficeCore.AppRuntimePersistenceService.newObject(context);
                            appRuntimePersistenceService.setAppRuntimeStartState(state);
                            return [4, context.sync()];
                        case 1:
                            _a.sent();
                            return [2];
                    }
                });
            });
        }
        addin.setStartupBehavior = setStartupBehavior;
        function getStartupBehavior() {
            return __awaiter(this, void 0, void 0, function () {
                var context, appRuntimePersistenceService, stateResult, state, ret;
                return __generator(this, function (_a) {
                    switch (_a.label) {
                        case 0:
                            context = _createRequestContext(false);
                            appRuntimePersistenceService = OfficeCore.AppRuntimePersistenceService.newObject(context);
                            stateResult = appRuntimePersistenceService.getAppRuntimeStartState();
                            return [4, context.sync()];
                        case 1:
                            _a.sent();
                            state = stateResult.value;
                            ret = (state == OfficeCore.AppRuntimeState.inactive ? StartupBehavior.none : StartupBehavior.load);
                            return [2, ret];
                    }
                });
            });
        }
        addin.getStartupBehavior = getStartupBehavior;
        function _setState(state) {
            return __awaiter(this, void 0, void 0, function () {
                var context, appRuntimeService;
                return __generator(this, function (_a) {
                    switch (_a.label) {
                        case 0:
                            context = _createRequestContext(true);
                            appRuntimeService = OfficeCore.AppRuntimeService.newObject(context);
                            appRuntimeService.setAppRuntimeState(state);
                            return [4, context.sync()];
                        case 1:
                            _a.sent();
                            return [2];
                    }
                });
            });
        }
        function _getState() {
            return __awaiter(this, void 0, void 0, function () {
                var context, appRuntimeService, stateResult;
                return __generator(this, function (_a) {
                    switch (_a.label) {
                        case 0:
                            context = _createRequestContext(true);
                            appRuntimeService = OfficeCore.AppRuntimeService.newObject(context);
                            stateResult = appRuntimeService.getAppRuntimeState();
                            return [4, context.sync()];
                        case 1:
                            _a.sent();
                            return [2, stateResult.value];
                    }
                });
            });
        }
        addin._getState = _getState;
        function showAsTaskpane() {
            return _setState(OfficeCore.AppRuntimeState.visible);
        }
        addin.showAsTaskpane = showAsTaskpane;
        function hide() {
            return _setState(OfficeCore.AppRuntimeState.background);
        }
        addin.hide = hide;
        var _appRuntimeEvent;
        function _getAppRuntimeEventService() {
            if (!_appRuntimeEvent) {
                var context = _createRequestContext(true);
                _appRuntimeEvent = OfficeCore.AppRuntimeService.newObject(context);
            }
            return _appRuntimeEvent;
        }
        addin.onVisibilityChanged = {
            add: function (handler) { return __awaiter(_this, void 0, void 0, function () {
                var eventService, registrationToken;
                var _this = this;
                return __generator(this, function (_a) {
                    switch (_a.label) {
                        case 0:
                            eventService = _getAppRuntimeEventService();
                            registrationToken = eventService.onVisibilityChanged.add(handler);
                            return [4, eventService.context.sync()];
                        case 1:
                            _a.sent();
                            return [2, {
                                    remove: function () { return __awaiter(_this, void 0, void 0, function () {
                                        return __generator(this, function (_a) {
                                            switch (_a.label) {
                                                case 0:
                                                    registrationToken.remove();
                                                    return [4, eventService.context.sync()];
                                                case 1:
                                                    _a.sent();
                                                    return [2];
                                            }
                                        });
                                    }); }
                                }];
                    }
                });
            }); }
        };
    })(addin = Office.addin || (Office.addin = {}));
})(Office || (Office = {}));
var __extends = (this && this.__extends) || (function () {
    var extendStatics = function (d, b) {
        extendStatics = Object.setPrototypeOf ||
            ({ __proto__: [] } instanceof Array && function (d, b) { d.__proto__ = b; }) ||
            function (d, b) { for (var p in b) if (b.hasOwnProperty(p)) d[p] = b[p]; };
        return extendStatics(d, b);
    };
    return function (d, b) {
        extendStatics(d, b);
        function __() { this.constructor = d; }
        d.prototype = b === null ? Object.create(b) : (__.prototype = b.prototype, new __());
    };
})();
var _BeginExcel = "_BeginExcel";
var Excel;
(function (Excel) {
    function _runOnRegularOrWacContext(options, callback) {
        var context = isOfficePlatform("OfficeOnline")
            ? new WacSpecificRequestContext()
            : new RequestContext();
        onBeforeExcelRun(options, context);
        return OfficeExtension.CoreUtility.Promise.resolve()
            .then(function () { return callback(context); })
            .then(context.sync);
    }
    var WacSpecificRequestContext = (function (_super) {
        __extends(WacSpecificRequestContext, _super);
        function WacSpecificRequestContext(url) {
            var _this = _super.call(this, url) || this;
            _this._customData = "WacPartition";
            _this.m_wacWorkbook = _createRootServiceObject(WacWorkbook, _this);
            _this._rootObject = _this.m_wacWorkbook;
            _this._rootObjectPropertyName = "wacWorkbook";
            return _this;
        }
        Object.defineProperty(WacSpecificRequestContext.prototype, "wacWorkbook", {
            get: function () {
                return this.m_wacWorkbook;
            },
            enumerable: true,
            configurable: true
        });
        return WacSpecificRequestContext;
    }(OfficeCore.RequestContext));
    var WacWorkbook = (function (_super) {
        __extends(WacWorkbook, _super);
        function WacWorkbook() {
            return _super !== null && _super.apply(this, arguments) || this;
        }
        WacWorkbook.prototype.createAndOpenWorkbook = function (base64) {
            _invokeMethod(this, "CreateAndOpenWorkbook", 0, [base64], 0, 0);
        };
        return WacWorkbook;
    }(OfficeExtension.ClientObject));
    function createWorkbook(base64) {
        _throwIfApiNotSupported("Excel.createWorkbook", _defaultApiSetName, "1.8", _hostName);
        return _runOnRegularOrWacContext({ delayForCellEdit: true }, function (context) {
            if (context instanceof WacSpecificRequestContext) {
                context.wacWorkbook.createAndOpenWorkbook(base64);
            }
            else {
                context.workbook.application.createWorkbook(base64).open();
            }
        });
    }
    Excel.createWorkbook = createWorkbook;
    function lowerCaseFirst(str) {
        return str[0].toLowerCase() + str.slice(1);
    }
    var iconSets = ["ThreeArrows",
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
    var iconNames = [["RedDownArrow", "YellowSideArrow", "GreenUpArrow"],
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
    Excel.icons = {};
    iconSets.map(function (title, i) {
        var camelTitle = lowerCaseFirst(title);
        Excel.icons[camelTitle] = [];
        iconNames[i].map(function (iconName, j) {
            iconName = lowerCaseFirst(iconName);
            var obj = { set: title, index: j };
            Excel.icons[camelTitle].push(obj);
            Excel.icons[camelTitle][iconName] = obj;
        });
    });
    function setRangePropertiesInBulk(range, propertyName, values) {
        var maxCellCount = 1500;
        if (Array.isArray(values) && values.length > 0 && Array.isArray(values[0]) && (values.length * values[0].length > maxCellCount) && isExcel1_3OrAbove()) {
            var maxRowCount = Math.max(1, Math.round(maxCellCount / values[0].length));
            range._ValidateArraySize(values.length, values[0].length);
            for (var startRowIndex = 0; startRowIndex < values.length; startRowIndex += maxRowCount) {
                var rowCount = maxRowCount;
                if (startRowIndex + rowCount > values.length) {
                    rowCount = values.length - startRowIndex;
                }
                var chunk = range.getRow(startRowIndex).untrack().getBoundingRect(range.getRow(startRowIndex + rowCount - 1).untrack()).untrack();
                var valueSlice = values.slice(startRowIndex, startRowIndex + rowCount);
                _invokeSetProperty(chunk, propertyName, valueSlice, 0);
            }
            return true;
        }
        return false;
    }
    function isExcelApiSetSupported(version) {
        if (typeof (window) !== "undefined" && window.Office && window.Office.context && window.Office.context.requirements) {
            return window.Office.context.requirements.isSetSupported("ExcelApi", version);
        }
        else {
            return true;
        }
    }
    function isExcel1_3OrAbove() {
        return isExcelApiSetSupported(1.3);
    }
    function isOfficePlatform(platform) {
        if (typeof (window) !== "undefined" && window.Office && window.Office.context) {
            return window.Office.context.platform === platform;
        }
        return false;
    }
    var Session = (function () {
        function Session(workbookUrl, requestHeaders, persisted) {
            this.m_workbookUrl = workbookUrl;
            this.m_requestHeaders = requestHeaders;
            if (!this.m_requestHeaders) {
                this.m_requestHeaders = {};
            }
            if (OfficeExtension.Utility.isNullOrUndefined(persisted)) {
                persisted = true;
            }
            this.m_persisted = persisted;
        }
        Session.prototype.close = function () {
            var _this = this;
            if (this.m_requestUrlAndHeaderInfo &&
                !OfficeExtension.Utility._isLocalDocumentUrl(this.m_requestUrlAndHeaderInfo.url)) {
                var url = this.m_requestUrlAndHeaderInfo.url;
                if (url.charAt(url.length - 1) !== "/") {
                    url = url + "/";
                }
                url = url + "closeSession";
                var headers = this.m_requestUrlAndHeaderInfo;
                var req = { method: "POST", url: url, headers: this.m_requestUrlAndHeaderInfo.headers, body: "" };
                this.m_requestUrlAndHeaderInfo = null;
                return OfficeExtension.HttpUtility.sendRequest(req)
                    .then(function (resp) {
                    if (resp.statusCode !== 204) {
                        var err = OfficeExtension.Utility._parseErrorResponse(resp);
                        throw OfficeExtension.Utility.createRuntimeError(err.errorCode, err.errorMessage, "Session.close");
                    }
                    _this.m_requestUrlAndHeaderInfo = null;
                    var foundSessionKey = null;
                    for (var key in _this.m_requestHeaders) {
                        if (key.toLowerCase() === Session.WorkbookSessionIdHeaderNameLower) {
                            foundSessionKey = key;
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
        Session.prototype._resolveRequestUrlAndHeaderInfo = function () {
            var _this = this;
            if (this.m_requestUrlAndHeaderInfo) {
                return OfficeExtension.Utility._createPromiseFromResult(this.m_requestUrlAndHeaderInfo);
            }
            if (OfficeExtension.Utility.isNullOrEmptyString(this.m_workbookUrl) ||
                OfficeExtension.Utility._isLocalDocumentUrl(this.m_workbookUrl)) {
                this.m_requestUrlAndHeaderInfo = { url: this.m_workbookUrl, headers: this.m_requestHeaders };
                return OfficeExtension.Utility._createPromiseFromResult(this.m_requestUrlAndHeaderInfo);
            }
            var foundSessionId = false;
            for (var key in this.m_requestHeaders) {
                if (key.toLowerCase() === Session.WorkbookSessionIdHeaderNameLower) {
                    foundSessionId = true;
                    break;
                }
            }
            if (foundSessionId) {
                this.m_requestUrlAndHeaderInfo = { url: this.m_workbookUrl, headers: this.m_requestHeaders };
                return OfficeExtension.Utility._createPromiseFromResult(this.m_requestUrlAndHeaderInfo);
            }
            var url = this.m_workbookUrl;
            if (url.charAt(url.length - 1) !== "/") {
                url = url + "/";
            }
            url = url + "createSession";
            var headers = {};
            OfficeExtension.Utility._copyHeaders(this.m_requestHeaders, headers);
            headers["Content-Type"] = "application/json";
            var body = {};
            body.persistChanges = this.m_persisted;
            var req = { method: "POST", url: url, headers: headers, body: JSON.stringify(body) };
            return OfficeExtension.HttpUtility.sendRequest(req)
                .then(function (resp) {
                if (resp.statusCode !== 201) {
                    var err = OfficeExtension.Utility._parseErrorResponse(resp);
                    throw OfficeExtension.Utility.createRuntimeError(err.errorCode, err.errorMessage, "Session.resolveRequestUrlAndHeaderInfo");
                }
                var session = JSON.parse(resp.body);
                var sessionId = session.id;
                headers = {};
                OfficeExtension.Utility._copyHeaders(_this.m_requestHeaders, headers);
                headers[Session.WorkbookSessionIdHeaderName] = sessionId;
                _this.m_requestUrlAndHeaderInfo = { url: _this.m_workbookUrl, headers: headers };
                return _this.m_requestUrlAndHeaderInfo;
            });
        };
        Session.WorkbookSessionIdHeaderName = "Workbook-Session-Id";
        Session.WorkbookSessionIdHeaderNameLower = "workbook-session-id";
        return Session;
    }());
    Excel.Session = Session;
    var RequestContext = (function (_super) {
        __extends(RequestContext, _super);
        function RequestContext(url) {
            var _this = _super.call(this, url) || this;
            _this.m_workbook = _createRootServiceObject(Excel.Workbook, _this);
            _this._rootObject = _this.m_workbook;
            _this._rootObjectPropertyName = "workbook";
            return _this;
        }
        RequestContext.prototype._processOfficeJsErrorResponse = function (officeJsErrorCode, response) {
            var ooeInvalidApiCallInContext = 5004;
            if (officeJsErrorCode === ooeInvalidApiCallInContext) {
                response.ErrorCode = ErrorCodes.invalidOperationInCellEditMode;
                response.ErrorMessage = OfficeExtension.Utility._getResourceString(OfficeExtension.ResourceStrings.invalidOperationInCellEditMode);
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
        Object.defineProperty(RequestContext.prototype, "runtime", {
            get: function () {
                return this.workbook._Runtime;
            },
            enumerable: true,
            configurable: true
        });
        return RequestContext;
    }(OfficeCore.RequestContext));
    Excel.RequestContext = RequestContext;
    function onBeforeExcelRun(options, context) {
        var excelOptions = options;
        if (excelOptions.delayForCellEdit && OfficeExtension.CommonUtility.isSetSupported("DelayForCellEdit")) {
            context._requestFlagModifier |= 64;
        }
        else {
            context._requestFlagModifier &= ~64;
        }
    }
    function run(arg1, arg2) {
        return OfficeExtension.ClientRequestContext._runBatch("Excel.run", arguments, function (requestInfo) {
            var ret = new Excel.RequestContext(requestInfo);
            return ret;
        }, onBeforeExcelRun);
    }
    Excel.run = run;
    function runBatch(arg1, arg2) {
        return OfficeExtension.ClientRequestContext._runExplicitBatch("Excel.runBatch", arguments, function (requestInfo) {
            var ret = new Excel.RequestContext(requestInfo);
            return ret;
        }, onBeforeExcelRun);
    }
    Excel.runBatch = runBatch;
    Excel._RedirectV1APIs = false;
    Excel._V1APIMap = {
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
            call: function (ctx, callArgs) {
                if (versionNumberIsEarlierThan({ ios: { desiredMajor: 2, desiredMinor: 20, desiredBuild: 0 } }) && OfficeExtension.CommonUtility.isSetSupported("DelayForCellEdit")) {
                    ctx._requestFlagModifier |= 64;
                }
                return ctx.workbook._V1Api.bindingAddFromPrompt(callArgs);
            },
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
                var preimage = callArgs["cellFormat"];
                if (typeof (window) !== "undefined" && window.OSF.DDA.SafeArray) {
                    if (window.OSF.OUtil.listContainsKey(window.OSF.DDA.SafeArray.Delegate.ParameterMap.dynamicTypes, "cellFormat")) {
                        callArgs["cellFormat"] = window.OSF.DDA.SafeArray.Delegate.ParameterMap.dynamicTypes["cellFormat"]["toHost"](preimage);
                    }
                }
                else if (typeof (window) !== "undefined" && window.OSF.DDA.WAC) {
                    if (window.OSF.OUtil.listContainsKey(window.OSF.DDA.WAC.Delegate.ParameterMap.dynamicTypes, "cellFormat")) {
                        callArgs["cellFormat"] = window.OSF.DDA.WAC.Delegate.ParameterMap.dynamicTypes["cellFormat"]["toHost"](preimage);
                    }
                }
                return callArgs;
            },
            call: function (ctx, callArgs) { return ctx.workbook._V1Api.setSelectedData(callArgs); }
        },
        "SetDataAsync": {
            preprocess: function (callArgs) {
                var preimage = callArgs["cellFormat"];
                if (typeof (window) !== "undefined" && window.OSF.DDA.SafeArray) {
                    if (window.OSF.OUtil.listContainsKey(window.OSF.DDA.SafeArray.Delegate.ParameterMap.dynamicTypes, "cellFormat")) {
                        callArgs["cellFormat"] = window.OSF.DDA.SafeArray.Delegate.ParameterMap.dynamicTypes["cellFormat"]["toHost"](preimage);
                    }
                }
                else if (typeof (window) !== "undefined" && window.OSF.DDA.WAC) {
                    if (window.OSF.OUtil.listContainsKey(window.OSF.DDA.WAC.Delegate.ParameterMap.dynamicTypes, "cellFormat")) {
                        callArgs["cellFormat"] = window.OSF.DDA.WAC.Delegate.ParameterMap.dynamicTypes["cellFormat"]["toHost"](preimage);
                    }
                }
                return callArgs;
            },
            call: function (ctx, callArgs) { return ctx.workbook._V1Api.bindingSetData(callArgs); }
        },
        "SetFormatsAsync": {
            preprocess: function (callArgs) {
                var preimage = callArgs["cellFormat"];
                if (typeof (window) !== "undefined" && window.OSF.DDA.SafeArray) {
                    if (window.OSF.OUtil.listContainsKey(window.OSF.DDA.SafeArray.Delegate.ParameterMap.dynamicTypes, "cellFormat")) {
                        callArgs["cellFormat"] = window.OSF.DDA.SafeArray.Delegate.ParameterMap.dynamicTypes["cellFormat"]["toHost"](preimage);
                    }
                }
                else if (typeof (window) !== "undefined" && window.OSF.DDA.WAC) {
                    if (window.OSF.OUtil.listContainsKey(window.OSF.DDA.WAC.Delegate.ParameterMap.dynamicTypes, "cellFormat")) {
                        callArgs["cellFormat"] = window.OSF.DDA.WAC.Delegate.ParameterMap.dynamicTypes["cellFormat"]["toHost"](preimage);
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
            call: function (ctx) { return ctx.workbook._V1Api.getFilePropertiesAsync(); }
        },
    };
    function postprocessBindingDescriptor(response) {
        var bindingDescriptor = {
            BindingColumnCount: response.bindingColumnCount,
            BindingId: response.bindingId,
            BindingRowCount: response.bindingRowCount,
            bindingType: response.bindingType,
            HasHeaders: response.hasHeaders
        };
        return window.OSF.DDA.OMFactory.manufactureBinding(bindingDescriptor, window.Microsoft.Office.WebExtension.context.document);
    }
    function getDataCommonPostprocess(response, callArgs) {
        var isPlainData = response.headers == null;
        var data;
        if (isPlainData) {
            data = response.rows;
        }
        else {
            data = response;
        }
        data = window.OSF.DDA.DataCoercion.coerceData(data, callArgs[window.Microsoft.Office.WebExtension.Parameters.CoercionType]);
        return data === undefined ? null : data;
    }
    function versionNumberIsEarlierThan(versionsForPlatformMap) {
        var hasOfficeVersion = typeof (window) !== "undefined" &&
            window.Office &&
            window.Office.context &&
            window.Office.context.diagnostics &&
            window.Office.context.diagnostics.version;
        if (!hasOfficeVersion) {
            return false;
        }
        var platform = window.Office.context.diagnostics.platform.toLowerCase();
        if (platform === 'officeonline') {
            return false;
        }
        var versionForCurrentPlatform = versionsForPlatformMap[platform];
        if (versionForCurrentPlatform == null) {
            versionForCurrentPlatform = versionsForPlatformMap.general;
        }
        var version = window.Office.context.diagnostics.version;
        var versionExtractor = /^(\d+)\.(\d+)\.(\d+)\.(\d+)$/;
        var result = versionExtractor.exec(version);
        if (result) {
            var major = parseInt(result[1]);
            var minor = parseInt(result[2]);
            var build = parseInt(result[3]);
            if (major < versionForCurrentPlatform.desiredMajor) {
                return true;
            }
            if (major === versionForCurrentPlatform.desiredMajor && minor < versionForCurrentPlatform.desiredMinor) {
                return true;
            }
            if (major === versionForCurrentPlatform.desiredMajor && minor === versionForCurrentPlatform.desiredMinor && build < versionForCurrentPlatform.desiredBuild) {
                var revisionString = result[4];
                var devBuildValidation = /^3\d\d\d+$/;
                var isDevBuild = devBuildValidation.exec(revisionString);
                if (isDevBuild) {
                    return true;
                }
            }
        }
        return false;
    }
    var ALWAYS_TRUE_PLACEHOLDER_OVERRIDE = true;
    var _CC;
    (function (_CC) {
        _CC.office10EventIdBindingSelectionChangedEvent = 3;
        _CC.office10EventIdBindingDataChangedEvent = 4;
        _CC.office10EventIdDocumentSelectionChangedEvent = 2;
        _CC.office10EventIdRichApiMessageEvent = 5;
        _CC.office10EventIdSettingsChangedEvent = 1;
    })(_CC = Excel._CC || (Excel._CC = {}));
    var _hostName = "Excel";
    var _defaultApiSetName = "ExcelApi";
    var _createPropertyObject = OfficeExtension.BatchApiHelper.createPropertyObject;
    var _createMethodObject = OfficeExtension.BatchApiHelper.createMethodObject;
    var _createIndexerObject = OfficeExtension.BatchApiHelper.createIndexerObject;
    var _createRootServiceObject = OfficeExtension.BatchApiHelper.createRootServiceObject;
    var _createTopLevelServiceObject = OfficeExtension.BatchApiHelper.createTopLevelServiceObject;
    var _createChildItemObject = OfficeExtension.BatchApiHelper.createChildItemObject;
    var _invokeMethod = OfficeExtension.BatchApiHelper.invokeMethod;
    var _invokeEnsureUnchanged = OfficeExtension.BatchApiHelper.invokeEnsureUnchanged;
    var _invokeSetProperty = OfficeExtension.BatchApiHelper.invokeSetProperty;
    var _isNullOrUndefined = OfficeExtension.Utility.isNullOrUndefined;
    var _isUndefined = OfficeExtension.Utility.isUndefined;
    var _throwIfNotLoaded = OfficeExtension.Utility.throwIfNotLoaded;
    var _throwIfApiNotSupported = OfficeExtension.Utility.throwIfApiNotSupported;
    var _load = OfficeExtension.Utility.load;
    var _retrieve = OfficeExtension.Utility.retrieve;
    var _toJson = OfficeExtension.Utility.toJson;
    var _fixObjectPathIfNecessary = OfficeExtension.Utility.fixObjectPathIfNecessary;
    var _handleNavigationPropertyResults = OfficeExtension.Utility._handleNavigationPropertyResults;
    var _adjustToDateTime = OfficeExtension.Utility.adjustToDateTime;
    var _processRetrieveResult = OfficeExtension.Utility.processRetrieveResult;
    var _setMockData = OfficeExtension.Utility.setMockData;
    (function (_CC) {
        function Application_SuspendApiCalculationUntilNextSync(thisObj) {
            if (isOfficePlatform("Mac") && isExcelApiSetSupported(1.6) && !isExcelApiSetSupported(1.7)) {
                return { handled: true };
            }
            return { handled: false };
        }
        _CC.Application_SuspendApiCalculationUntilNextSync = Application_SuspendApiCalculationUntilNextSync;
    })(_CC = Excel._CC || (Excel._CC = {}));
    (function (_CC) {
        function Workbook_SelectionChanged_EventArgsTransform(thisObj, args) {
            return { workbook: thisObj };
        }
        _CC.Workbook_SelectionChanged_EventArgsTransform = Workbook_SelectionChanged_EventArgsTransform;
        function Workbook_WACOperationEvent_EventArgsTransform(thisObj, args) {
            var event = {
                eventId: args.eventId,
                payload: args.payload,
                type: Excel.EventType.wacoperationEvent,
            };
            return event;
        }
        _CC.Workbook_WACOperationEvent_EventArgsTransform = Workbook_WACOperationEvent_EventArgsTransform;
        function Workbook__Message_EventArgsTransform(thisObj, args) {
            return {
                entries: args.entries,
                workbook: thisObj
            };
        }
        _CC.Workbook__Message_EventArgsTransform = Workbook__Message_EventArgsTransform;
    })(_CC = Excel._CC || (Excel._CC = {}));
    (function (_CC) {
        function Worksheet_Charts_Get(thisObj, ret) {
            ret._ParentObject = thisObj;
        }
        _CC.Worksheet_Charts_Get = Worksheet_Charts_Get;
        function Worksheet_ShowGridlines_Set(thisObj, value) {
            if (ALWAYS_TRUE_PLACEHOLDER_OVERRIDE) {
                thisObj.m_showGridlines = value;
                _invokeSetProperty(thisObj, "Gridlines", value, 0);
                return { handled: true };
            }
            return { handled: false };
        }
        _CC.Worksheet_ShowGridlines_Set = Worksheet_ShowGridlines_Set;
        function Worksheet_ShowHeadings_Set(thisObj, value) {
            if (ALWAYS_TRUE_PLACEHOLDER_OVERRIDE) {
                thisObj.m_showHeadings = value;
                _invokeSetProperty(thisObj, "Headings", value, 0);
                return { handled: true };
            }
            return { handled: false };
        }
        _CC.Worksheet_ShowHeadings_Set = Worksheet_ShowHeadings_Set;
        function Worksheet_Tables_Get(thisObj, ret) {
            ret._ParentObject = thisObj;
        }
        _CC.Worksheet_Tables_Get = Worksheet_Tables_Get;
        function Worksheet_Visuals_Get(thisObj, ret) {
            ret._ParentObject = thisObj;
        }
        _CC.Worksheet_Visuals_Get = Worksheet_Visuals_Get;
        function Worksheet_Changed_EventArgsTransform(thisObj, args) {
            var value = args;
            var details;
            if (value.valueBefore != null || value.valueAfter != null) {
                details = {
                    valueBefore: value.valueBefore,
                    valueAfter: value.valueAfter,
                    valueTypeBefore: value.valueTypeBefore,
                    valueTypeAfter: value.valueTypeAfter
                };
            }
            var newArgs = {
                type: Excel.EventType.worksheetChanged,
                changeType: value.changeType,
                source: value.source,
                worksheetId: thisObj.id,
                address: value.address,
                getRange: function (ctx) {
                    _throwIfApiNotSupported("WorksheetChangedEventArgs.getRange", _defaultApiSetName, "1.8", _hostName);
                    return ctx.workbook._GetRangeForEventByReferenceId(value.referenceId);
                },
                getRangeOrNullObject: function (ctx) {
                    _throwIfApiNotSupported("WorksheetChangedEventArgs.getRangeOrNullObject", _defaultApiSetName, "1.8", _hostName);
                    return ctx.workbook._GetRangeOrNullObjectForEventByReferenceId(value.referenceId);
                },
                details: details
            };
            return newArgs;
        }
        _CC.Worksheet_Changed_EventArgsTransform = Worksheet_Changed_EventArgsTransform;
        function Worksheet_FormatChanged_EventArgsTransform(thisObj, args) {
            var value = args;
            var newArgs = {
                type: Excel.EventType.worksheetFormatChanged,
                source: value.source,
                worksheetId: thisObj.id,
                address: value.address,
                getRange: function (ctx) {
                    _throwIfApiNotSupported("WorksheetFormatChangedEventArgs.getRange", _defaultApiSetName, "1.9", _hostName);
                    return ctx.workbook._GetRangeForEventByReferenceId(value.referenceId);
                },
                getRangeOrNullObject: function (ctx) {
                    _throwIfApiNotSupported("WorksheetFormatChangedEventArgs.getRangeOrNullObject", _defaultApiSetName, "1.9", _hostName);
                    return ctx.workbook._GetRangeOrNullObjectForEventByReferenceId(value.referenceId);
                }
            };
            return newArgs;
        }
        _CC.Worksheet_FormatChanged_EventArgsTransform = Worksheet_FormatChanged_EventArgsTransform;
    })(_CC = Excel._CC || (Excel._CC = {}));
    (function (_CC) {
        function _overrideWorksheetCollectionEventMethod(thisObj, methodName, apiVersion) {
            if ((!isOfficePlatform("OfficeOnline") && !isExcelApiSetSupported(1.12)) ||
                (isOfficePlatform("OfficeOnline") && !isExcelApiSetSupported(1.11))) {
                _throwIfApiNotSupported("WorksheetCollection." + methodName, _defaultApiSetName, apiVersion, _hostName);
                _invokeMethod(thisObj, methodName, 0, [], 0, 0);
                return { handled: true };
            }
            return { handled: false };
        }
        function WorksheetCollection__RegisterActivatedEvent(thisObj) {
            if (!isExcelApiSetSupported(1.9)) {
                _throwIfApiNotSupported("WorksheetCollection._RegisterActivatedEvent", _defaultApiSetName, "1.7", _hostName);
                _invokeMethod(thisObj, "_RegisterActivatedEvent", 0, [], 0, 0);
                return { handled: true };
            }
            return { handled: false };
        }
        _CC.WorksheetCollection__RegisterActivatedEvent = WorksheetCollection__RegisterActivatedEvent;
        function WorksheetCollection__RegisterAddedEvent(thisObj) {
            return _overrideWorksheetCollectionEventMethod(thisObj, "_RegisterAddedEvent", "1.7");
        }
        _CC.WorksheetCollection__RegisterAddedEvent = WorksheetCollection__RegisterAddedEvent;
        function WorksheetCollection__RegisterColumnSortedEvent(thisObj) {
            return _overrideWorksheetCollectionEventMethod(thisObj, "_RegisterColumnSortedEvent", "1.10");
        }
        _CC.WorksheetCollection__RegisterColumnSortedEvent = WorksheetCollection__RegisterColumnSortedEvent;
        function WorksheetCollection__RegisterDataChangedEvent(thisObj) {
            return _overrideWorksheetCollectionEventMethod(thisObj, "_RegisterDataChangedEvent", "1.9");
        }
        _CC.WorksheetCollection__RegisterDataChangedEvent = WorksheetCollection__RegisterDataChangedEvent;
        function WorksheetCollection__RegisterDeletedEvent(thisObj) {
            return _overrideWorksheetCollectionEventMethod(thisObj, "_RegisterDeletedEvent", "1.7");
        }
        _CC.WorksheetCollection__RegisterDeletedEvent = WorksheetCollection__RegisterDeletedEvent;
        function WorksheetCollection__RegisterRowSortedEvent(thisObj) {
            return _overrideWorksheetCollectionEventMethod(thisObj, "_RegisterRowSortedEvent", "1.10");
        }
        _CC.WorksheetCollection__RegisterRowSortedEvent = WorksheetCollection__RegisterRowSortedEvent;
        function WorksheetCollection__UnregisterAddedEvent(thisObj) {
            return _overrideWorksheetCollectionEventMethod(thisObj, "_UnregisterAddedEvent", "1.7");
        }
        _CC.WorksheetCollection__UnregisterAddedEvent = WorksheetCollection__UnregisterAddedEvent;
        function WorksheetCollection__UnregisterColumnSortedEvent(thisObj) {
            return _overrideWorksheetCollectionEventMethod(thisObj, "_UnregisterColumnSortedEvent", "1.10");
        }
        _CC.WorksheetCollection__UnregisterColumnSortedEvent = WorksheetCollection__UnregisterColumnSortedEvent;
        function WorksheetCollection__UnregisterDataChangedEvent(thisObj) {
            return _overrideWorksheetCollectionEventMethod(thisObj, "_UnregisterDataChangedEvent", "1.9");
        }
        _CC.WorksheetCollection__UnregisterDataChangedEvent = WorksheetCollection__UnregisterDataChangedEvent;
        function WorksheetCollection__UnregisterDeletedEvent(thisObj) {
            return _overrideWorksheetCollectionEventMethod(thisObj, "_UnregisterDeletedEvent", "1.7");
        }
        _CC.WorksheetCollection__UnregisterDeletedEvent = WorksheetCollection__UnregisterDeletedEvent;
        function WorksheetCollection__UnregisterRowSortedEvent(thisObj) {
            return _overrideWorksheetCollectionEventMethod(thisObj, "_UnregisterRowSortedEvent", "1.10");
        }
        _CC.WorksheetCollection__UnregisterRowSortedEvent = WorksheetCollection__UnregisterRowSortedEvent;
        function WorksheetCollection_Changed_EventArgsTransform(thisObj, args) {
            var value = args;
            var details;
            if (value.valueBefore != null || value.valueAfter != null) {
                details = {
                    valueBefore: value.valueBefore,
                    valueAfter: value.valueAfter,
                    valueTypeBefore: value.valueTypeBefore,
                    valueTypeAfter: value.valueTypeAfter
                };
            }
            var newArgs = {
                type: Excel.EventType.worksheetChanged,
                changeType: value.changeType,
                source: value.source,
                worksheetId: value.worksheetId,
                address: value.address,
                getRange: function (ctx) {
                    _throwIfApiNotSupported("WorksheetChangedEventArgs.getRange", _defaultApiSetName, "1.9", _hostName);
                    return ctx.workbook._GetRangeForEventByReferenceId(value.referenceId);
                },
                getRangeOrNullObject: function (ctx) {
                    _throwIfApiNotSupported("WorksheetChangedEventArgs.getRangeOrNullObject", _defaultApiSetName, "1.9", _hostName);
                    return ctx.workbook._GetRangeOrNullObjectForEventByReferenceId(value.referenceId);
                },
                details: details
            };
            return newArgs;
        }
        _CC.WorksheetCollection_Changed_EventArgsTransform = WorksheetCollection_Changed_EventArgsTransform;
        function WorksheetCollection_FormatChanged_EventArgsTransform(thisObj, args) {
            var value = args;
            var newArgs = {
                type: Excel.EventType.worksheetFormatChanged,
                source: value.source,
                worksheetId: value.worksheetId,
                address: value.address,
                getRange: function (ctx) {
                    _throwIfApiNotSupported("WorksheetFormatChangedEventArgs.getRange", _defaultApiSetName, "1.9", _hostName);
                    return ctx.workbook._GetRangeForEventByReferenceId(value.referenceId);
                },
                getRangeOrNullObject: function (ctx) {
                    _throwIfApiNotSupported("WorksheetFormatChangedEventArgs.getRangeOrNullObject", _defaultApiSetName, "1.9", _hostName);
                    return ctx.workbook._GetRangeOrNullObjectForEventByReferenceId(value.referenceId);
                }
            };
            return newArgs;
        }
        _CC.WorksheetCollection_FormatChanged_EventArgsTransform = WorksheetCollection_FormatChanged_EventArgsTransform;
    })(_CC = Excel._CC || (Excel._CC = {}));
    (function (_CC) {
        function WorksheetProtection_Protect(thisObj, options, password) {
            if (versionNumberIsEarlierThan({
                mac: { desiredMajor: 16, desiredMinor: 7, desiredBuild: 17101600 },
                ios: { desiredMajor: 2, desiredMinor: 7, desiredBuild: 1016 },
                general: { desiredMajor: 16, desiredMinor: 0, desiredBuild: 8716 }
            })) {
                _invokeMethod(thisObj, "Protect", 0, [options], 0, 0);
                return { handled: true };
            }
            return { handled: false };
        }
        _CC.WorksheetProtection_Protect = WorksheetProtection_Protect;
    })(_CC = Excel._CC || (Excel._CC = {}));
    var RangeCustom = (function () {
        function RangeCustom() {
        }
        RangeCustom.prototype._ensureInteger = function (num, methodName) {
            if (!(typeof num === "number" && isFinite(num) && Math.floor(num) === num)) {
                OfficeExtension.Utility.throwError(Excel.ErrorCodes.invalidArgument, num, methodName);
            }
        };
        RangeCustom.prototype._getAdjacentRange = function (functionName, count, referenceRange, rowDirection, columnDirection) {
            if (count == null) {
                count = 1;
            }
            this._ensureInteger(count, functionName);
            var startRange;
            var rowOffset = 0;
            var columnOffset = 0;
            if (count > 0) {
                startRange = referenceRange.getOffsetRange(rowDirection, columnDirection);
            }
            else {
                startRange = referenceRange;
                rowOffset = rowDirection;
                columnOffset = columnDirection;
            }
            if (Math.abs(count) === 1) {
                return startRange;
            }
            return startRange.getBoundingRect(referenceRange.getOffsetRange(rowDirection * count + rowOffset, columnDirection * count + columnOffset));
        };
        return RangeCustom;
    }());
    Excel.RangeCustom = RangeCustom;
    (function (_CC) {
        function Range_HandleResult(thisObj, value) {
            if (!_isUndefined(value["isEntireColumn"])) {
                thisObj.m_isEntireColumn = value["isEntireColumn"];
            }
            if (!_isUndefined(value["isEntireRow"])) {
                thisObj.m_isEntireRow = value["isEntireRow"];
            }
        }
        _CC.Range_HandleResult = Range_HandleResult;
        function Range_GetColumnsAfter(thisObj, count) {
            if (!isExcel1_3OrAbove()) {
                if (count == null) {
                    count = 1;
                }
                thisObj._ensureInteger(count, "RowsAbove");
                if (count === 0) {
                    OfficeExtension.Utility.throwError(Excel.ErrorCodes.invalidArgument, "count", "RowsAbove");
                }
                return {
                    handled: true,
                    result: thisObj._getAdjacentRange("getColumnsAfter", count, thisObj.getLastColumn(), 0, 1)
                };
            }
            return { handled: false, result: null };
        }
        _CC.Range_GetColumnsAfter = Range_GetColumnsAfter;
        function Range_GetColumnsBefore(thisObj, count) {
            if (!isExcel1_3OrAbove()) {
                if (count == null) {
                    count = 1;
                }
                thisObj._ensureInteger(count, "RowsAbove");
                if (count === 0) {
                    OfficeExtension.Utility.throwError(Excel.ErrorCodes.invalidArgument, "count", "RowsAbove");
                }
                return {
                    handled: true,
                    result: thisObj._getAdjacentRange("getColumnsBefore", count, thisObj.getColumn(0), 0, -1)
                };
            }
            return { handled: false, result: null };
        }
        _CC.Range_GetColumnsBefore = Range_GetColumnsBefore;
        function Range_GetResizedRange(thisObj, deltaRows, deltaColumns) {
            if (!isExcel1_3OrAbove()) {
                thisObj._ensureInteger(deltaRows, "getResizedRange");
                thisObj._ensureInteger(deltaColumns, "getResizedRange");
                var referenceRange = (deltaRows >= 0 && deltaColumns >= 0) ? thisObj : thisObj.getCell(0, 0);
                var result = referenceRange.getBoundingRect(thisObj.getLastCell().getOffsetRange(deltaRows, deltaColumns));
                return { handled: true, result: result };
            }
            return { handled: false, result: null };
        }
        _CC.Range_GetResizedRange = Range_GetResizedRange;
        function Range_GetRowsAbove(thisObj, count) {
            if (!isExcel1_3OrAbove()) {
                if (count == null) {
                    count = 1;
                }
                thisObj._ensureInteger(count, "RowsAbove");
                if (count === 0) {
                    OfficeExtension.Utility.throwError(Excel.ErrorCodes.invalidArgument, "count", "RowsAbove");
                }
                var result = thisObj._getAdjacentRange("getRowsAbove", count, thisObj.getRow(0), -1, 0);
                return { handled: true, result: result };
            }
            return { handled: false, result: null };
        }
        _CC.Range_GetRowsAbove = Range_GetRowsAbove;
        function Range_GetRowsBelow(thisObj, count) {
            if (!isExcel1_3OrAbove()) {
                if (count == null) {
                    count = 1;
                }
                thisObj._ensureInteger(count, "RowsAbove");
                if (count === 0) {
                    OfficeExtension.Utility.throwError(Excel.ErrorCodes.invalidArgument, "count", "RowsAbove");
                }
                var result = this._getAdjacentRange("getRowsBelow", count, thisObj.getLastRow(), 1, 0);
                return { handled: true, result: result };
            }
            return { handled: false, result: null };
        }
        _CC.Range_GetRowsBelow = Range_GetRowsBelow;
        function Range_Formulas_Set(thisObj, value) {
            thisObj.m_formulas = value;
            if (setRangePropertiesInBulk(thisObj, "Formulas", value)) {
                return { handled: true };
            }
            return { handled: false };
        }
        _CC.Range_Formulas_Set = Range_Formulas_Set;
        function Range_FormulasLocal_Set(thisObj, value) {
            thisObj.m_formulasLocal = value;
            if (setRangePropertiesInBulk(thisObj, "FormulasLocal", value)) {
                return { handled: true };
            }
            return { handled: false };
        }
        _CC.Range_FormulasLocal_Set = Range_FormulasLocal_Set;
        function Range_FormulasR1C1_Set(thisObj, value) {
            thisObj.m_formulasR1C1 = value;
            if (setRangePropertiesInBulk(thisObj, "FormulasR1C1", value)) {
                return { handled: true };
            }
            return { handled: false };
        }
        _CC.Range_FormulasR1C1_Set = Range_FormulasR1C1_Set;
        function Range_NumberFormat_Set(thisObj, value) {
            thisObj.m_numberFormat = value;
            if (setRangePropertiesInBulk(thisObj, "NumberFormat", value)) {
                return { handled: true };
            }
            return { handled: false };
        }
        _CC.Range_NumberFormat_Set = Range_NumberFormat_Set;
        function Range_Values_Set(thisObj, value) {
            thisObj.m_values = value;
            if (setRangePropertiesInBulk(thisObj, "Values", value)) {
                return { handled: true };
            }
            return { handled: false };
        }
        _CC.Range_Values_Set = Range_Values_Set;
    })(_CC = Excel._CC || (Excel._CC = {}));
    (function (_CC) {
        function SettingCollection_Add(thisObj, key, value) {
            value = _CC._replaceDateWithStringDate(value);
            var result = _createMethodObject(Excel.Setting, thisObj, "Add", 0, [key, value], false, true, null, 0);
            return { handled: true, result: result };
        }
        _CC.SettingCollection_Add = SettingCollection_Add;
        function SettingCollection_SettingsChanged_EventArgsTransform(thisObj, args) {
            return {
                settings: thisObj
            };
        }
        _CC.SettingCollection_SettingsChanged_EventArgsTransform = SettingCollection_SettingsChanged_EventArgsTransform;
    })(_CC = Excel._CC || (Excel._CC = {}));
    (function (_CC) {
        var DateJSONPrefix = "Date(";
        var DateJSONSuffix = ")";
        function _replaceStringDateWithDate(value) {
            var strValue = JSON.stringify(value);
            value = JSON.parse(strValue, function dateReviver(k, v) {
                var d;
                if (typeof v === 'string' && v && v.length > 6 && v.slice(0, 5) === DateJSONPrefix && v.slice(-1) === DateJSONSuffix) {
                    d = new Date(parseInt(v.slice(5, -1)));
                    if (d) {
                        return d;
                    }
                }
                return v;
            });
            return value;
        }
        function _replaceDateWithStringDate(value) {
            var strValue = JSON.stringify(value, function dateReplacer(k, v) {
                return (this[k] instanceof Date) ? (DateJSONPrefix + this[k].getTime() + DateJSONSuffix) : v;
            });
            value = JSON.parse(strValue);
            return value;
        }
        _CC._replaceDateWithStringDate = _replaceDateWithStringDate;
        function Setting_HandleResult(thisObj, value) {
            if (!_isUndefined(value["Value"])) {
                value["Value"] = _replaceStringDateWithDate(value["Value"]);
            }
            ;
        }
        _CC.Setting_HandleResult = Setting_HandleResult;
        function Setting_Value_Set(thisObj, value) {
            if (!_isNullOrUndefined(value)) {
                thisObj.m_value = value;
                var newValue = _replaceDateWithStringDate(value);
                _invokeSetProperty(thisObj, "Value", newValue, 0);
                return { handled: true };
            }
            return { handled: false };
        }
        _CC.Setting_Value_Set = Setting_Value_Set;
    })(_CC = Excel._CC || (Excel._CC = {}));
    (function (_CC) {
        function Binding_DataChanged_EventArgsTransform(thisObj, args) {
            var evt = {
                binding: thisObj
            };
            return evt;
        }
        _CC.Binding_DataChanged_EventArgsTransform = Binding_DataChanged_EventArgsTransform;
        function Binding_SelectionChanged_EventArgsTransform(thisObj, args) {
            var evt = {
                binding: thisObj,
                columnCount: args.columnCount,
                rowCount: args.rowCount,
                startColumn: args.startColumn,
                startRow: args.startRow
            };
            return evt;
        }
        _CC.Binding_SelectionChanged_EventArgsTransform = Binding_SelectionChanged_EventArgsTransform;
    })(_CC = Excel._CC || (Excel._CC = {}));
    var TableCollectionCustom = (function () {
        function TableCollectionCustom() {
        }
        Object.defineProperty(TableCollectionCustom.prototype, "_ParentObject", {
            get: function () {
                return this.m__ParentObject;
            },
            set: function (value) {
                this.m__ParentObject = value;
            },
            enumerable: true,
            configurable: true
        });
        return TableCollectionCustom;
    }());
    Excel.TableCollectionCustom = TableCollectionCustom;
    (function (_CC) {
        function TableCollection_Changed_EventArgsTransform(thisObj, args) {
            var value = args;
            var details;
            if (value.valueBefore != null || value.valueAfter != null) {
                details = {
                    valueBefore: value.valueBefore,
                    valueAfter: value.valueAfter,
                    valueTypeBefore: value.valueTypeBefore,
                    valueTypeAfter: value.valueTypeAfter
                };
            }
            var newArgs = {
                type: Excel.EventType.tableChanged,
                changeType: value.changeType,
                source: value.source,
                worksheetId: value.worksheetId,
                tableId: value.tableId,
                address: value.address,
                getRange: function (ctx) {
                    _throwIfApiNotSupported("TableChangedEventArgs.getRange", _defaultApiSetName, "1.8", _hostName);
                    return ctx.workbook._GetRangeForEventByReferenceId(value.referenceId);
                },
                getRangeOrNullObject: function (ctx) {
                    _throwIfApiNotSupported("TableChangedEventArgs.getRangeOrNullObject", _defaultApiSetName, "1.8", _hostName);
                    return ctx.workbook._GetRangeOrNullObjectForEventByReferenceId(value.referenceId);
                },
                details: details
            };
            return newArgs;
        }
        _CC.TableCollection_Changed_EventArgsTransform = TableCollection_Changed_EventArgsTransform;
    })(_CC = Excel._CC || (Excel._CC = {}));
    (function (_CC) {
        function Table_HandleIdResult(thisObj, value) {
            if (!_isUndefined(value["Id"])) {
                value["Id"] = value["Id"].toString();
            }
        }
        _CC.Table_HandleIdResult = Table_HandleIdResult;
        function Table_HandleResult(thisObj, value) {
            if (!_isUndefined(value["Id"])) {
                value["Id"] = value["Id"].toString();
            }
        }
        _CC.Table_HandleResult = Table_HandleResult;
        function Table_Changed_EventArgsTransform(thisObj, args) {
            var value = args;
            var details;
            if (value.valueBefore != null || value.valueAfter != null) {
                details = {
                    valueBefore: value.valueBefore,
                    valueAfter: value.valueAfter,
                    valueTypeBefore: value.valueTypeBefore,
                    valueTypeAfter: value.valueTypeAfter
                };
            }
            var newArgs = {
                type: Excel.EventType.tableChanged,
                changeType: value.changeType,
                source: value.source,
                worksheetId: value.worksheetId,
                tableId: value.tableId,
                address: value.address,
                getRange: function (ctx) {
                    _throwIfApiNotSupported("TableChangedEventArgs.getRange", _defaultApiSetName, "1.8", _hostName);
                    return ctx.workbook._GetRangeForEventByReferenceId(value.referenceId);
                },
                getRangeOrNullObject: function (ctx) {
                    _throwIfApiNotSupported("TableChangedEventArgs.getRangeOrNullObject", _defaultApiSetName, "1.8", _hostName);
                    return ctx.workbook._GetRangeOrNullObjectForEventByReferenceId(value.referenceId);
                },
                details: details
            };
            return newArgs;
        }
        _CC.Table_Changed_EventArgsTransform = Table_Changed_EventArgsTransform;
        function Table_SelectionChanged_EventArgsTransform(thisObj, args) {
            var value = args;
            var isAddressNullOrEmpty = (!value.address || value.address.length === 0);
            var newArgs = {
                type: Excel.EventType.tableSelectionChanged,
                isInsideTable: !isAddressNullOrEmpty,
                worksheetId: value.worksheetId,
                tableId: thisObj.id,
                address: value.address
            };
            return newArgs;
        }
        _CC.Table_SelectionChanged_EventArgsTransform = Table_SelectionChanged_EventArgsTransform;
    })(_CC = Excel._CC || (Excel._CC = {}));
    var ChartCollectionCustom = (function () {
        function ChartCollectionCustom() {
        }
        Object.defineProperty(ChartCollectionCustom.prototype, "_ParentObject", {
            get: function () {
                return this.m__ParentObject;
            },
            set: function (value) {
                this.m__ParentObject = value;
            },
            enumerable: true,
            configurable: true
        });
        return ChartCollectionCustom;
    }());
    Excel.ChartCollectionCustom = ChartCollectionCustom;
    (function (_CC) {
        function ChartCollection_Add(thisObj, type, sourceData, seriesBy) {
            if (!(sourceData instanceof Excel.Range)) {
                throw OfficeExtension.Utility.createRuntimeError(OfficeExtension.ResourceStrings.invalidArgument, "sourceData", "Charts.Add");
            }
            return { handled: false, result: null };
        }
        _CC.ChartCollection_Add = ChartCollection_Add;
    })(_CC = Excel._CC || (Excel._CC = {}));
    (function (_CC) {
        function Chart_SetData(thisObj, sourceData, seriesBy) {
            if (!(sourceData instanceof Excel.Range)) {
                throw OfficeExtension.Utility.createRuntimeError(OfficeExtension.ResourceStrings.invalidArgument, "sourceData", "Chart.setData");
            }
            return { handled: false };
        }
        _CC.Chart_SetData = Chart_SetData;
    })(_CC = Excel._CC || (Excel._CC = {}));
    var ChartFillCustom = (function () {
        function ChartFillCustom() {
        }
        ChartFillCustom.prototype.load = function (option) {
            _load(this, option);
            return this;
        };
        return ChartFillCustom;
    }());
    Excel.ChartFillCustom = ChartFillCustom;
    var VisualCollectionCustom = (function () {
        function VisualCollectionCustom() {
        }
        Object.defineProperty(VisualCollectionCustom.prototype, "_ParentObject", {
            get: function () {
                return this.m__ParentObject;
            },
            set: function (value) {
                this.m__ParentObject = value;
            },
            enumerable: true,
            configurable: true
        });
        return VisualCollectionCustom;
    }());
    Excel.VisualCollectionCustom = VisualCollectionCustom;
    (function (_CC) {
        function VisualCollection_AgaveVisualUpdate_EventArgsTransform(thisObj, args) {
            var value = args;
            var newArgs = {
                type: value.type,
                payload: value.payload
            };
            return newArgs;
        }
        _CC.VisualCollection_AgaveVisualUpdate_EventArgsTransform = VisualCollection_AgaveVisualUpdate_EventArgsTransform;
    })(_CC = Excel._CC || (Excel._CC = {}));
    (function (_CC) {
        function Visual_ChangeNotification_EventArgsTransform(thisObj, args) {
            var value = args;
            var newArgs = {
                targetId: value.targetId,
                changeType: value.changeType,
                payload: value.payload,
                type: Excel.EventType.visualChange,
            };
            return newArgs;
        }
        _CC.Visual_ChangeNotification_EventArgsTransform = Visual_ChangeNotification_EventArgsTransform;
    })(_CC = Excel._CC || (Excel._CC = {}));
    (function (_CC) {
        function PivotField_SortByLabels(thisObj, sortBy) {
            if (typeof sortBy === "string") {
                sortBy = (sortBy.toLowerCase() === "ascending");
            }
            _invokeMethod(thisObj, "SortByLabels", 0, [sortBy], 0, 0);
            return { handled: true };
        }
        _CC.PivotField_SortByLabels = PivotField_SortByLabels;
    })(_CC = Excel._CC || (Excel._CC = {}));
    var CustomFunctionManagerCustom = (function () {
        function CustomFunctionManagerCustom() {
        }
        CustomFunctionManagerCustom.register = function (metadata, javascript) {
            _throwIfApiNotSupported("CustomFunctionManager.register", "CustomFunctions", "1.3", _hostName);
            return _runOnRegularOrWacContext({ delayForCellEdit: true }, function (context) { return Excel.CustomFunctionManager.newObject(context).register(metadata, javascript); });
        };
        CustomFunctionManagerCustom.getStatus = function () {
            _throwIfApiNotSupported("CustomFunctionManager.register", "CustomFunctions", "1.3", _hostName);
            return _runOnRegularOrWacContext({ delayForCellEdit: true }, function (context) {
                var manager = Excel.CustomFunctionManager.newObject(context).load("status");
                return context.sync().then(function () { return manager.status; });
            });
        };
        return CustomFunctionManagerCustom;
    }());
    Excel.CustomFunctionManagerCustom = CustomFunctionManagerCustom;
    (function (_CC) {
        function CustomFunctionManager_StaticCustomize(type) {
            type.register = CustomFunctionManagerCustom.register;
            type.getStatus = CustomFunctionManagerCustom.getStatus;
        }
        _CC.CustomFunctionManager_StaticCustomize = CustomFunctionManager_StaticCustomize;
    })(_CC = Excel._CC || (Excel._CC = {}));
    (function (_CC) {
        function Style_TextOrientation_Set(thisObj, value) {
            if (ALWAYS_TRUE_PLACEHOLDER_OVERRIDE) {
                thisObj.m_textOrientation = value;
                _invokeSetProperty(thisObj, "Orientation", value, 0);
                return { handled: true };
            }
            return { handled: false };
        }
        _CC.Style_TextOrientation_Set = Style_TextOrientation_Set;
    })(_CC = Excel._CC || (Excel._CC = {}));
    (function (_CC) {
        function InternalTest_Test1Event_EventArgsTransform(thisObj, args) {
            var value = args;
            var newArgs = {
                prop1: value.prop1,
                worksheet: thisObj.context.workbook.worksheets.getItem(value.worksheetId)
            };
            return newArgs;
        }
        _CC.InternalTest_Test1Event_EventArgsTransform = InternalTest_Test1Event_EventArgsTransform;
        function InternalTest_TestEvent_EventArgsTransform(thisObj, args) {
            var value = args;
            var newArgs = {
                prop1: value.prop1,
                worksheet: thisObj.context.workbook.worksheets.getItem(value.worksheetId)
            };
            return newArgs;
        }
        _CC.InternalTest_TestEvent_EventArgsTransform = InternalTest_TestEvent_EventArgsTransform;
    })(_CC = Excel._CC || (Excel._CC = {}));
    var ErrorCodes;
    (function (ErrorCodes) {
        ErrorCodes["accessDenied"] = "AccessDenied";
        ErrorCodes["apiNotFound"] = "ApiNotFound";
        ErrorCodes["conflict"] = "Conflict";
        ErrorCodes["generalException"] = "GeneralException";
        ErrorCodes["insertDeleteConflict"] = "InsertDeleteConflict";
        ErrorCodes["invalidArgument"] = "InvalidArgument";
        ErrorCodes["invalidBinding"] = "InvalidBinding";
        ErrorCodes["invalidOperation"] = "InvalidOperation";
        ErrorCodes["invalidReference"] = "InvalidReference";
        ErrorCodes["invalidSelection"] = "InvalidSelection";
        ErrorCodes["itemAlreadyExists"] = "ItemAlreadyExists";
        ErrorCodes["itemNotFound"] = "ItemNotFound";
        ErrorCodes["nonBlankCellOffSheet"] = "NonBlankCellOffSheet";
        ErrorCodes["notImplemented"] = "NotImplemented";
        ErrorCodes["rangeExceedsLimit"] = "RangeExceedsLimit";
        ErrorCodes["requestAborted"] = "RequestAborted";
        ErrorCodes["unsupportedOperation"] = "UnsupportedOperation";
        ErrorCodes["invalidOperationInCellEditMode"] = "InvalidOperationInCellEditMode";
    })(ErrorCodes = Excel.ErrorCodes || (Excel.ErrorCodes = {}));
    var Interfaces;
    (function (Interfaces) {
    })(Interfaces = Excel.Interfaces || (Excel.Interfaces = {}));
    var _libraryMetadataXlapi = { "version": "1.0.0",
        "name": "Excel",
        "defaultApiSetName": "ExcelApi",
        "hostName": "Excel",
        "apiSets": [["1.8"], ["1.9"], ["1.11"], ["1.6"], ["1.7"], ["1.12"], ["1.2"], ["1.3"], ["1.4"], ["1.5"], ["1.10"], ["1.1", "CustomFunctions"], ["1.1", "WACOperationEvents"], ["99.9"], ["1.1", "ExcelApiOnline"]],
        "strings": ["DataConnection", "Runtime", "Application", "IterativeCalculation", "Workbook", "WorkbookProtection", "WorkbookCreated", "Worksheet", "WorksheetCollection", "WorksheetProtection", "WorksheetFreezePanes", "Range", "RangeAreas", "RangeView", "RangeViewCollection", "SettingCollection", "Setting", "NamedItemCollection", "NamedItem", "NamedItemArrayValues", "Binding", "BindingCollection", "TableCollection", "TableScopedCollection", "Table", "TableColumnCollection", "TableColumn", "TableRowCollection", "TableRow", "DataValidation", "RemoveDuplicatesResult", "RangeFormat", "FormatProtection", "RangeFill", "RangeBorder", "RangeBorderCollection", "RangeFont", "ChartCollection", "Chart", "ChartPivotOptions", "ChartAreaFormat", "ChartSeriesCollection", "ChartSeries", "ChartSeriesFormat", "ChartPointsCollection", "ChartPoint", "ChartPointFormat", "ChartAxes", "ChartAxis", "ChartAxisFormat", "ChartAxisTitle", "ChartAxisTitleFormat", "ChartDataLabels", "ChartDataLabel", "ChartDataLabelFormat", "ChartErrorBars", "ChartErrorBarsFormat", "ChartGridlines", "ChartGridlinesFormat", "ChartLegend", "ChartLegendEntry", "ChartLegendEntryCollection", "ChartLegendFormat", "ChartMapOptions", "ChartTitle", "ChartFormatString", "ChartTitleFormat", "ChartFill", "ChartBorder", "ChartBinOptions", "ChartBoxwhiskerOptions", "ChartLineFormat", "ChartFont", "ChartTrendline", "ChartTrendlineCollection", "ChartTrendlineFormat", "ChartTrendlineLabel", "ChartTrendlineLabelFormat", "ChartPlotArea", "ChartPlotAreaFormat", "VisualCollection", "Visual", "VisualProperty", "VisualPropertyCollection", "DataControllerClient", "RangeSort", "TableSort", "Filter", "AutoFilter", "CultureInfo", "NumberFormatInfo", "DatetimeFormatInfo", "CustomXmlPartScopedCollection", "CustomXmlPartCollection", "CustomXmlPart", "_V1Api", "PivotTableScopedCollection", "PivotTableCollection", "PivotTable", "PivotLayout", "PivotHierarchyCollection", "PivotHierarchy", "RowColumnPivotHierarchyCollection", "RowColumnPivotHierarchy", "FilterPivotHierarchyCollection", "FilterPivotHierarchy", "DataPivotHierarchyCollection", "DataPivotHierarchy", "PivotFieldCollection", "PivotField", "PivotItemCollection", "PivotItem", "WorksheetCustomProperty", "WorksheetCustomPropertyCollection", "DocumentProperties", "CustomProperty", "CustomPropertyCollection", "ConditionalFormatCollection", "ConditionalFormat", "DataBarConditionalFormat", "ConditionalDataBarPositiveFormat", "ConditionalDataBarNegativeFormat", "CustomConditionalFormat", "ConditionalFormatRule", "IconSetConditionalFormat", "ColorScaleConditionalFormat", "TopBottomConditionalFormat", "PresetCriteriaConditionalFormat", "TextConditionalFormat", "CellValueConditionalFormat", "ConditionalRangeFormat", "ConditionalRangeFont", "ConditionalRangeFill", "ConditionalRangeBorder", "ConditionalRangeBorderCollection", "NumberFormattingService", "NumberFormatter", "CustomFunctionManager", "Style", "StyleCollection", "TableStyleCollection", "TableStyle", "PivotTableStyleCollection", "PivotTableStyle", "SlicerStyleCollection", "SlicerStyle", "TimelineStyleCollection", "TimelineStyle", "InternalTest", "PageLayout", "HeaderFooter", "HeaderFooterGroup", "PageBreak", "PageBreakCollection", "DataConnectionCollection", "RangeCollection", "CommentCollection", "Comment", "CommentReplyCollection", "CommentReply", "ShapeCollection", "Shape", "GeometricShape", "Image", "ShapeGroup", "GroupShapeCollection", "Line", "ShapeFill", "ShapeLineFormat", "TextFrame", "TextRange", "ShapeFont", "Slicer", "SlicerCollection", "SlicerItem", "SlicerItemCollection", "Ribbon", "SheetView", "SheetViewCollection", "FunctionResult", "Functions", "name", "calculationEngineVersion", "calculate", "_GetWorkbookCreatedById", "enabled", "readOnly", "names", "tables", "pivotTables", "properties", "protection", "comments", "slicers", "_RegisterAutoSaveSettingChangedEvent", "_UnregisterAutoSaveSettingChangedEvent", "_RegisterWACOperationEvent", "_UnregisterWACOperationEvent", "OfficeExtension.Constants.eventWorkbookId", "SelectionChanged", "", "protected", "protect", "unprotect", "id", "position", "shapes", "autoFilter", "delete", "activate", "_RegisterDataChangedEvent", "_UnregisterDataChangedEvent", "_RegisterActivatedEvent", "_UnregisterActivatedEvent", "_RegisterDeactivatedEvent", "_UnregisterDeactivatedEvent", "_RegisterSelectionChangedEvent", "_UnregisterSelectionChangedEvent", "_RegisterCalculatedEvent", "_UnregisterCalculatedEvent", "replaceAll", "_RegisterFilteredEvent", "_UnregisterFilteredEvent", "_RegisterFormatChangedEvent", "_UnregisterFormatChangedEvent", "_RegisterRowSortedEvent", "_UnregisterRowSortedEvent", "_RegisterColumnSortedEvent", "_UnregisterColumnSortedEvent", "_RegisterSingleClickedEvent", "_UnregisterSingleClickedEvent", "_RegisterRowHiddenChangedEvent", "_UnregisterRowHiddenChangedEvent", "getRange", "getUsedRange", "getCell", "getUsedRangeOrNullObject", "_GetSheetById", "_GetAnotherWorksheetById", "getRanges", "Activated", "MessageType.worksheetActivatedEvent", "this.id", "Calculated", "MessageType.worksheetCalculatedEvent", "Changed", "MessageType.worksheetDataChangedEvent", "ColumnSorted", "MessageType.worksheetColumnSortedEvent", "Deactivated", "MessageType.worksheetDeactivatedEvent", "Filtered", "MessageType.worksheetFilteredEvent", "FormatChanged", "MessageType.worksheetFormatChangedEvent", "RowHiddenChanged", "MessageType.worksheetRowHiddenChangedEvent", "RowSorted", "MessageType.worksheetRowSortedEvent", "MessageType.worksheetSelectionChangedEvent", "SingleClicked", "MessageType.worksheetSingleClickedEvent", "getCount", "_RegisterAddedEvent", "_UnregisterAddedEvent", "_RegisterDeletedEvent", "_UnregisterDeletedEvent", "getItem", "add", "getItemOrNullObject", "getFirst", "Added", "Deleted", "options", "getLocation", "numberFormat", "numberFormatLocal", "values", "text", "formulas", "formulasLocal", "rowIndex", "columnIndex", "rowCount", "columnCount", "address", "addressLocal", "cellCount", "_ReferenceId", "valueTypes", "formulasR1C1", "isEntireColumn", "isEntireRow", "hyperlink", "style", "top", "left", "height", "width", "format", "worksheet", "sort", "conditionalFormats", "dataValidation", "clear", "select", "_KeepReference", "getImage", "copyFrom", "convertToLinkedDataType", "convertDataTypeToText", "setDirty", "showTeachingCallout", "group", "ungroup", "getEntireColumn", "getEntireRow", "getIntersection", "getIntersectionOrNullObject", "getTables", "find", "getSpecialCells", "getSpecialCellsOrNullObject", "areas", "index", "rows", "getItemAt", "key", "value", "_Id", "type", "visible", "formula", "getRangeOrNullObject", "count", "this._ParentObject ? this._ParentObject.id : OfficeExtension.Constants.eventWorkbookId", "MessageType.tableDataChangedEvent", "MessageType.tableFilteredEvent", "columns", "clearFilters", "setStyle", "getHeaderRowRange", "getDataBodyRange", "getTotalRowRange", "rule", "wrapText", "horizontalAlignment", "verticalAlignment", "rowHeight", "textOrientation", "readingOrder", "shrinkToFit", "indentLevel", "autoIndent", "fill", "font", "borders", "locked", "formulaHidden", "color", "tintAndShade", "sideIndex", "weight", "size", "italic", "bold", "underline", "strikethrough", "_GetItem", "MessageType.chartActivatedEvent", "this._ParentObject.id", "MessageType.chartDeactivatedEvent", "chartType", "title", "dataLabels", "border", "markerSize", "markerStyle", "showShadow", "markerBackgroundColor", "markerForegroundColor", "axisGroup", "line", "linkNumberFormat", "setFormula", "showValue", "showSeriesName", "showCategoryName", "showLegendKey", "showPercentage", "showBubbleSize", "separator", "autoText", "overlay", "level", "getSubstring", "setSolidColor", "lineStyle", "null", "_RegisterChangeEvent", "_UnregisterChangeEvent", "min", "max", "apply", "fields", "reapply", "criteria", "remove", "refreshAll", "setToDefault", "creationDate", "custom", "fillColor", "borderColor", "right", "isText", "setDefault", "getDefault", "duplicate", "_RegisterTestEvent", "_UnregisterTestEvent", "_RegisterTest1Event", "_UnregisterTest1Event", "_RegisterCustomFunctionExecutionBeginEvent", "_UnregisterCustomFunctionExecutionBeginEvent", "_RegisterCustomFunctionExecutionEndEvent", "_UnregisterCustomFunctionExecutionEndEvent", "orientation", "leftMargin", "rightMargin", "topMargin", "bottomMargin", "content", "authorName", "authorEmail", "resolved", "richContent", "mentions", "contentType", "updateMentions", "_GetShapeById", "shape", "connectorType", "transparency", "_RegisterCommandExecutedEvent", "_UnregisterCommandExecutedEvent"],
        "enumTypes": [["DataSourceType", ["unknown", "cube"]],
            ["DateFilterCondition", ["unknown", "equals", "before", "beforeOrEqualTo", "after", "afterOrEqualTo", "between", "tomorrow", "today", "yesterday", "nextWeek", "thisWeek", "lastWeek", "nextMonth", "thisMonth", "lastMonth", "nextQuarter", "thisQuarter", "lastQuarter", "nextYear", "thisYear", "lastYear", "yearToDate", "allDatesInPeriodQuarter1", "allDatesInPeriodQuarter2", "allDatesInPeriodQuarter3", "allDatesInPeriodQuarter4", "allDatesInPeriodJanuary", "allDatesInPeriodFebruary", "allDatesInPeriodMarch", "allDatesInPeriodApril", "allDatesInPeriodMay", "allDatesInPeriodJune", "allDatesInPeriodJuly", "allDatesInPeriodAugust", "allDatesInPeriodSeptember", "allDatesInPeriodOctober", "allDatesInPeriodNovember", "allDatesInPeriodDecember"]],
            ["LabelFilterCondition", ["unknown", "equals", "beginsWith", "endsWith", "contains", "greaterThan", "greaterThanOrEqualTo", "lessThan", "lessThanOrEqualTo", "between"]],
            ["PivotFilterType", ["unknown", "value", "manual", "label", "date"]],
            ["TopBottomSelectionType", ["items", "percent", "sum"]],
            ["ValueFilterCondition", ["unknown", "equals", "greaterThan", "greaterThanOrEqualTo", "lessThan", "lessThanOrEqualTo", "between", "topN", "bottomN"]],
            ["BuiltInPivotTableStyle", ["none", "light1", "light2", "light3", "light4", "light5", "light6", "light7", "light8", "light9", "light10", "light11", "light12", "light13", "light14", "light15", "light16", "light17", "light18", "light19", "light20", "light21", "light22", "light23", "light24", "light25", "light26", "light27", "light28", "medium1", "medium2", "medium3", "medium4", "medium5", "medium6", "medium7", "medium8", "medium9", "medium10", "medium11", "medium12", "medium13", "medium14", "medium15", "medium16", "medium17", "medium18", "medium19", "medium20", "medium21", "medium22", "medium23", "medium24", "medium25", "medium26", "medium27", "medium28", "dark1", "dark2", "dark3", "dark4", "dark5", "dark6", "dark7", "dark8", "dark9", "dark10", "dark11", "dark12", "dark13", "dark14", "dark15", "dark16", "dark17", "dark18", "dark19", "dark20", "dark21", "dark22", "dark23", "dark24", "dark25", "dark26", "dark27", "dark28"]],
            ["BuiltInSlicerStyle", ["light1", "light2", "light3", "light4", "light5", "light6", "other1", "other2", "dark1", "dark2", "dark3", "dark4", "dark5", "dark6"]],
            ["BuiltInTableStyle", ["light1", "light2", "light3", "light4", "light5", "light6", "light7", "light8", "light9", "light10", "light11", "light12", "light13", "light14", "light15", "light16", "light17", "light18", "light19", "light20", "light21", "medium1", "medium2", "medium3", "medium4", "medium5", "medium6", "medium7", "medium8", "medium9", "medium10", "medium11", "medium12", "medium13", "medium14", "medium15", "medium16", "medium17", "medium18", "medium19", "medium20", "medium21", "medium22", "medium23", "medium24", "medium25", "medium26", "medium27", "medium28", "dark1", "dark2", "dark3", "dark4", "dark5", "dark6", "dark7", "dark8", "dark9", "dark10", "dark11"]],
            ["ChartSeriesDimension", ["categories", "values"], { "xvalues": "XValues", "yvalues": "YValues" }],
            ["PivotFilterTopBottomCriterion", ["invalid", "topItems", "topPercent", "topSum", "bottomItems", "bottomPercent", "bottomSum"]],
            ["SortBy", ["ascending", "descending"]],
            ["AggregationFunction", ["unknown", "automatic", "sum", "count", "average", "max", "min", "product", "countNumbers", "standardDeviation", "standardDeviationP", "variance", "varianceP"]],
            ["ShowAsCalculation", ["unknown", "none", "percentOfGrandTotal", "percentOfRowTotal", "percentOfColumnTotal", "percentOfParentRowTotal", "percentOfParentColumnTotal", "percentOfParentTotal", "percentOf", "runningTotal", "percentRunningTotal", "differenceFrom", "percentDifferenceFrom", "rankAscending", "rankDecending", "index"]],
            ["PivotAxis", ["unknown", "row", "column", "data", "filter"]],
            ["AxisType", ["invalid", "category", "value", "series"]],
            ["AxisGroup", ["primary", "secondary"]],
            ["AxisScaleType", ["linear", "logarithmic"]],
            ["AxisCrosses", ["automatic", "maximum", "minimum", "custom"]],
            ["AxisTickMark", ["none", "cross", "inside", "outside"]],
            ["AxisTickLabelPosition", ["nextToAxis", "high", "low", "none"]],
            ["TrendlineType", ["linear", "exponential", "logarithmic", "movingAverage", "polynomial", "power"]],
            ["ChartAxisType", ["invalid", "category", "value", "series"]],
            ["ChartAxisGroup", ["primary", "secondary"]],
            ["ChartAxisScaleType", ["linear", "logarithmic"]],
            ["ChartAxisPosition", ["automatic", "maximum", "minimum", "custom"]],
            ["ChartAxisTickMark", ["none", "cross", "inside", "outside"]],
            ["CalculationState", ["done", "calculating", "pending"]],
            ["ChartAxisTickLabelPosition", ["nextToAxis", "high", "low", "none"]],
            ["ChartAxisDisplayUnit", ["none", "hundreds", "thousands", "tenThousands", "hundredThousands", "millions", "tenMillions", "hundredMillions", "billions", "trillions", "custom"]],
            ["ChartAxisTimeUnit", ["days", "months", "years"]],
            ["ChartBoxQuartileCalculation", ["inclusive", "exclusive"]],
            ["ChartAxisCategoryType", ["automatic", "textAxis", "dateAxis"]],
            ["ChartBinType", ["category", "auto", "binWidth", "binCount"]],
            ["ChartLineStyle", ["none", "continuous", "dash", "dashDot", "dashDotDot", "dot", "grey25", "grey50", "grey75", "automatic", "roundDot"]],
            ["ChartDataLabelPosition", ["invalid", "none", "center", "insideEnd", "insideBase", "outsideEnd", "left", "right", "top", "bottom", "bestFit", "callout"]],
            ["ChartErrorBarsInclude", ["both", "minusValues", "plusValues"]],
            ["ChartErrorBarsType", ["fixedValue", "percent", "stDev", "stError", "custom"]],
            ["ChartMapAreaLevel", ["automatic", "dataOnly", "city", "county", "state", "country", "continent", "world"]],
            ["ChartGradientStyle", ["twoPhaseColor", "threePhaseColor"]],
            ["ChartGradientStyleType", ["extremeValue", "number", "percent"]],
            ["ChartTitlePosition", ["automatic", "top", "bottom", "left", "right"]],
            ["ChartLegendPosition", ["invalid", "top", "bottom", "left", "right", "corner", "custom"]],
            ["ChartMarkerStyle", ["invalid", "automatic", "none", "square", "diamond", "triangle", "star", "dot", "dash", "circle", "plus", "picture"], { "x": "X" }],
            ["ChartPlotAreaPosition", ["automatic", "custom"]],
            ["ChartMapLabelStrategy", ["none", "bestFit", "showAll"]],
            ["ChartMapProjectionType", ["automatic", "mercator", "miller", "robinson", "albers"]],
            ["ChartParentLabelStrategy", ["none", "banner", "overlapping"]],
            ["ChartSeriesBy", ["auto", "columns", "rows"]],
            ["ChartTextHorizontalAlignment", ["center", "left", "right", "justify", "distributed"]],
            ["ChartTextVerticalAlignment", ["center", "bottom", "top", "justify", "distributed"]],
            ["ChartTickLabelAlignment", ["center", "left", "right"]],
            ["ChartType", ["invalid", "columnClustered", "columnStacked", "columnStacked100", "barClustered", "barStacked", "barStacked100", "lineStacked", "lineStacked100", "lineMarkers", "lineMarkersStacked", "lineMarkersStacked100", "pieOfPie", "pieExploded", "barOfPie", "areaStacked", "areaStacked100", "doughnutExploded", "radarMarkers", "radarFilled", "surface", "surfaceWireframe", "surfaceTopView", "surfaceTopViewWireframe", "bubble", "bubble3DEffect", "stockHLC", "stockOHLC", "stockVHLC", "stockVOHLC", "cylinderColClustered", "cylinderColStacked", "cylinderColStacked100", "cylinderBarClustered", "cylinderBarStacked", "cylinderBarStacked100", "cylinderCol", "coneColClustered", "coneColStacked", "coneColStacked100", "coneBarClustered", "coneBarStacked", "coneBarStacked100", "coneCol", "pyramidColClustered", "pyramidColStacked", "pyramidColStacked100", "pyramidBarClustered", "pyramidBarStacked", "pyramidBarStacked100", "pyramidCol", "line", "pie", "area", "doughnut", "radar", "histogram", "boxwhisker", "pareto", "regionMap", "treemap", "waterfall", "sunburst", "funnel"], { "_3DColumnClustered": "3DColumnClustered", "_3DColumnStacked": "3DColumnStacked", "_3DColumnStacked100": "3DColumnStacked100", "_3DBarClustered": "3DBarClustered", "_3DBarStacked": "3DBarStacked", "_3DBarStacked100": "3DBarStacked100", "_3DPieExploded": "3DPieExploded", "xyscatterSmooth": "XYScatterSmooth", "xyscatterSmoothNoMarkers": "XYScatterSmoothNoMarkers", "xyscatterLines": "XYScatterLines", "xyscatterLinesNoMarkers": "XYScatterLinesNoMarkers", "_3DAreaStacked": "3DAreaStacked", "_3DAreaStacked100": "3DAreaStacked100", "_3DColumn": "3DColumn", "_3DLine": "3DLine", "_3DPie": "3DPie", "xyscatter": "XYScatter", "_3DArea": "3DArea" }],
            ["ChartUnderlineStyle", ["none", "single"]],
            ["ChartDisplayBlanksAs", ["notPlotted", "zero", "interplotted"]],
            ["ChartPlotBy", ["rows", "columns"]],
            ["ChartSplitType", ["splitByPosition", "splitByValue", "splitByPercentValue", "splitByCustomSplit"]],
            ["ChartColorScheme", ["colorfulPalette1", "colorfulPalette2", "colorfulPalette3", "colorfulPalette4", "monochromaticPalette1", "monochromaticPalette2", "monochromaticPalette3", "monochromaticPalette4", "monochromaticPalette5", "monochromaticPalette6", "monochromaticPalette7", "monochromaticPalette8", "monochromaticPalette9", "monochromaticPalette10", "monochromaticPalette11", "monochromaticPalette12", "monochromaticPalette13"]],
            ["ChartTrendlineType", ["linear", "exponential", "logarithmic", "movingAverage", "polynomial", "power"]],
            ["ShapeZOrder", ["bringToFront", "bringForward", "sendToBack", "sendBackward"]],
            ["ShapeType", ["unsupported", "image", "geometricShape", "group", "line"]],
            ["ShapeScaleType", ["currentSize", "originalSize"]],
            ["ShapeScaleFrom", ["scaleFromTopLeft", "scaleFromMiddle", "scaleFromBottomRight"]],
            ["ShapeFillType", ["noFill", "solid", "gradient", "pattern", "pictureAndTexture", "mixed"]],
            ["ShapeFontUnderlineStyle", ["none", "single", "double", "heavy", "dotted", "dottedHeavy", "dash", "dashHeavy", "dashLong", "dashLongHeavy", "dotDash", "dotDashHeavy", "dotDotDash", "dotDotDashHeavy", "wavy", "wavyHeavy", "wavyDouble"]],
            ["PictureFormat", [], { "unknown": "UNKNOWN", "bmp": "BMP", "jpeg": "JPEG", "gif": "GIF", "png": "PNG", "svg": "SVG" }],
            ["ShapeLineStyle", ["single", "thickBetweenThin", "thickThin", "thinThick", "thinThin"]],
            ["ShapeLineDashStyle", ["dash", "dashDot", "dashDotDot", "longDash", "longDashDot", "roundDot", "solid", "squareDot", "longDashDotDot", "systemDash", "systemDot", "systemDashDot"]],
            ["ArrowheadLength", ["short", "medium", "long"]],
            ["ArrowheadStyle", ["none", "triangle", "stealth", "diamond", "oval", "open"]],
            ["ArrowheadWidth", ["narrow", "medium", "wide"]],
            ["BindingType", ["range", "table", "text"]],
            ["BorderIndex", ["edgeTop", "edgeBottom", "edgeLeft", "edgeRight", "insideVertical", "insideHorizontal", "diagonalDown", "diagonalUp"]],
            ["BorderLineStyle", ["none", "continuous", "dash", "dashDot", "dashDotDot", "dot", "double", "slantDashDot"]],
            ["BorderWeight", ["hairline", "thin", "medium", "thick"]],
            ["CalculationMode", ["automatic", "automaticExceptTables", "manual"]],
            ["RuntimeCalculationMode", [], { "undefined": 0, "automatic": 1, "automaticExceptTables": 2, "manual": 3 }],
            ["CalculationType", ["recalculate", "full", "fullRebuild"]],
            ["ClearApplyTo", ["all", "formats", "contents", "hyperlinks", "removeHyperlinks"]],
            ["VisualCategory", ["column", "bar", "line", "area", "pie", "donut", "scatter", "bubble", "statistical", "stock", "combo", "hierarchy", "surface", "map", "funnel", "radar", "waterfall", "threeD", "other"]],
            ["VisualPropertyType", ["object", "collection", "string", "double", "int", "bool", "enum", "color"]],
            ["VisualChangeType", ["dataChange", "propertyChange", "genericChange", "selectionChange"]],
            ["BoolMetaPropertyType", ["writeOnly", "readOnly", "hideEntireSubtreeUI", "hideMeButShowChildrenUI", "expandableUI", "nextPropOnSameLine", "hideLabel", "showResetUI", "hasOwnExpandableSection", "nextPropOnSameLineFOTP", "showResetUIFOTP", "untransferable"]],
            ["ConditionalDataBarAxisFormat", ["automatic", "none", "cellMidPoint"]],
            ["ConditionalDataBarDirection", ["context", "leftToRight", "rightToLeft"]],
            ["ConditionalFormatDirection", ["top", "bottom"]],
            ["ConditionalFormatType", ["custom", "dataBar", "colorScale", "iconSet", "topBottom", "presetCriteria", "containsText", "cellValue"]],
            ["ConditionalFormatRuleType", ["invalid", "automatic", "lowestValue", "highestValue", "number", "percent", "formula", "percentile"]],
            ["ConditionalFormatIconRuleType", ["invalid", "number", "percent", "formula", "percentile"]],
            ["ConditionalFormatColorCriterionType", ["invalid", "lowestValue", "highestValue", "number", "percent", "formula", "percentile"]],
            ["ConditionalTopBottomCriterionType", ["invalid", "topItems", "topPercent", "bottomItems", "bottomPercent"]],
            ["ConditionalFormatPresetCriterion", ["invalid", "blanks", "nonBlanks", "errors", "nonErrors", "yesterday", "today", "tomorrow", "lastSevenDays", "lastWeek", "thisWeek", "nextWeek", "lastMonth", "thisMonth", "nextMonth", "aboveAverage", "belowAverage", "equalOrAboveAverage", "equalOrBelowAverage", "oneStdDevAboveAverage", "oneStdDevBelowAverage", "twoStdDevAboveAverage", "twoStdDevBelowAverage", "threeStdDevAboveAverage", "threeStdDevBelowAverage", "uniqueValues", "duplicateValues"]],
            ["ConditionalTextOperator", ["invalid", "contains", "notContains", "beginsWith", "endsWith"]],
            ["ConditionalCellValueOperator", ["invalid", "between", "notBetween", "equalTo", "notEqualTo", "greaterThan", "lessThan", "greaterThanOrEqual", "lessThanOrEqual"]],
            ["ConditionalIconCriterionOperator", ["invalid", "greaterThan", "greaterThanOrEqual"]],
            ["ConditionalRangeBorderIndex", ["edgeTop", "edgeBottom", "edgeLeft", "edgeRight"]],
            ["ConditionalRangeBorderLineStyle", ["none", "continuous", "dash", "dashDot", "dashDotDot", "dot"]],
            ["ConditionalRangeFontUnderlineStyle", ["none", "single", "double"]],
            ["CustomFunctionType", ["invalid", "script", "webService"]],
            ["CustomFunctionMetadataFormat", ["invalid", "openApi"]],
            ["DataValidationType", ["none", "wholeNumber", "decimal", "list", "date", "time", "textLength", "custom", "inconsistent", "mixedCriteria"]],
            ["DataValidationOperator", ["between", "notBetween", "equalTo", "notEqualTo", "greaterThan", "lessThan", "greaterThanOrEqualTo", "lessThanOrEqualTo"]],
            ["DataValidationAlertStyle", ["stop", "warning", "information"]],
            ["DeleteShiftDirection", ["up", "left"]],
            ["DynamicFilterCriteria", ["unknown", "aboveAverage", "allDatesInPeriodApril", "allDatesInPeriodAugust", "allDatesInPeriodDecember", "allDatesInPeriodFebruray", "allDatesInPeriodJanuary", "allDatesInPeriodJuly", "allDatesInPeriodJune", "allDatesInPeriodMarch", "allDatesInPeriodMay", "allDatesInPeriodNovember", "allDatesInPeriodOctober", "allDatesInPeriodQuarter1", "allDatesInPeriodQuarter2", "allDatesInPeriodQuarter3", "allDatesInPeriodQuarter4", "allDatesInPeriodSeptember", "belowAverage", "lastMonth", "lastQuarter", "lastWeek", "lastYear", "nextMonth", "nextQuarter", "nextWeek", "nextYear", "thisMonth", "thisQuarter", "thisWeek", "thisYear", "today", "tomorrow", "yearToDate", "yesterday"]],
            ["FilterDatetimeSpecificity", ["year", "month", "day", "hour", "minute", "second"]],
            ["FilterOn", ["bottomItems", "bottomPercent", "cellColor", "dynamic", "fontColor", "values", "topItems", "topPercent", "icon", "custom"]],
            ["FilterOperator", ["and", "or"]],
            ["HorizontalAlignment", ["general", "left", "center", "right", "fill", "justify", "centerAcrossSelection", "distributed"]],
            ["IconSet", ["invalid", "threeArrows", "threeArrowsGray", "threeFlags", "threeTrafficLights1", "threeTrafficLights2", "threeSigns", "threeSymbols", "threeSymbols2", "fourArrows", "fourArrowsGray", "fourRedToBlack", "fourRating", "fourTrafficLights", "fiveArrows", "fiveArrowsGray", "fiveRating", "fiveQuarters", "threeStars", "threeTriangles", "fiveBoxes"]],
            ["ImageFittingMode", ["fit", "fitAndCenter", "fill"]],
            ["InsertShiftDirection", ["down", "right"]],
            ["NamedItemScope", ["worksheet", "workbook"]],
            ["NamedItemType", ["string", "integer", "double", "boolean", "range", "error", "array"]],
            ["RangeUnderlineStyle", ["none", "single", "double", "singleAccountant", "doubleAccountant"]],
            ["SheetVisibility", ["visible", "hidden", "veryHidden"]],
            ["RangeValueType", ["unknown", "empty", "string", "integer", "double", "boolean", "error", "richValue"]],
            ["SearchDirection", ["forward", "backwards"]],
            ["SortOrientation", ["rows", "columns"]],
            ["SortOn", ["value", "cellColor", "fontColor", "icon"]],
            ["SortDataOption", ["normal", "textAsNumber"]],
            ["SortMethod", ["pinYin", "strokeCount"]],
            ["VerticalAlignment", ["top", "center", "bottom", "justify", "distributed"]],
            ["MessageCategory", [], { "none": 0, "customFunction": 1, "action": 2, "event": 65536 }],
            ["MessageType", [], { "none": 0, "testEvent": 1, "test1Event": 2, "worksheetDataChangedEvent": 10, "worksheetActivatedEvent": 11, "worksheetDeactivatedEvent": 12, "worksheetAddedEvent": 13, "worksheetSelectionChangedEvent": 14, "worksheetDeletedEvent": 15, "worksheetCalculatedEvent": 16, "worksheetFilteredEvent": 17, "worksheetFormatChangedEvent": 18, "worksheetRowSortedEvent": 19, "worksheetColumnSortedEvent": 20, "worksheetSingleClickedEvent": 21, "worksheetRowHiddenChangedEvent": 22, "chartAddedEvent": 50, "chartActivatedEvent": 51, "chartDeactivatedEvent": 52, "chartDeletedEvent": 53, "tableSelectionChangedEvent": 100, "tableDataChangedEvent": 101, "tableAddedEvent": 102, "tableDeletedEvent": 103, "tableFilteredEvent": 104, "agaveVisualUpdateEvent": 150, "customFunctionExecutionBeginEvent": 200, "customFunctionExecutionEndEvent": 201, "invocationMessage": 1000, "cancellationMessage": 1001, "metadataMessage": 1002, "visualSelectionChangedEvent": 2000, "visualChangeEvent": 2001, "shapeSelectionChangedEvent": 2100, "shapeActivatedEvent": 2101, "shapeDeactivatedEvent": 2102, "workbookAutoSaveSettingChangedEvent": 2200, "wacoperationEvent": 2201, "ribbonCommandExecutedEvent": 2300, "appRuntimeVisibilityChangedEvent": 65539 }],
            ["DocumentPropertyType", ["number", "boolean", "date", "string", "float"]],
            ["EventSource", ["local", "remote"]],
            ["DataChangeType", ["unknown", "rangeEdited", "rowInserted", "rowDeleted", "columnInserted", "columnDeleted", "cellInserted", "cellDeleted"]],
            ["RowHiddenChangeType", ["unhidden", "hidden"]],
            ["EventType", ["worksheetChanged", "worksheetSelectionChanged", "worksheetAdded", "worksheetActivated", "worksheetDeactivated", "tableChanged", "tableSelectionChanged", "worksheetDeleted", "chartAdded", "chartActivated", "chartDeactivated", "chartDeleted", "worksheetCalculated", "visualSelectionChanged", "agaveVisualUpdate", "tableAdded", "tableDeleted", "tableFiltered", "worksheetFiltered", "shapeActivated", "shapeDeactivated", "visualChange", "workbookAutoSaveSettingChanged", "worksheetFormatChanged", "ribbonCommandExecuted", "worksheetRowSorted", "worksheetColumnSorted", "worksheetSingleClicked", "worksheetRowHiddenChanged"], { "wacoperationEvent": "WACOperationEvent" }],
            ["DocumentPropertyItem", ["title", "subject", "author", "keywords", "comments", "template", "lastAuth", "revision", "appName", "lastPrint", "creation", "lastSave", "category", "format", "manager", "company"]],
            ["SubtotalLocationType", ["atTop", "atBottom", "off"]],
            ["PivotLayoutType", ["compact", "tabular", "outline"]],
            ["ProtectionSelectionMode", ["normal", "unlocked", "none"]],
            ["PageOrientation", ["portrait", "landscape"]],
            ["PaperType", ["letter", "letterSmall", "tabloid", "ledger", "legal", "statement", "executive", "a3", "a4", "a4Small", "a5", "b4", "b5", "folio", "quatro", "paper10x14", "paper11x17", "note", "envelope9", "envelope10", "envelope11", "envelope12", "envelope14", "csheet", "dsheet", "esheet", "envelopeDL", "envelopeC5", "envelopeC3", "envelopeC4", "envelopeC6", "envelopeC65", "envelopeB4", "envelopeB5", "envelopeB6", "envelopeItaly", "envelopeMonarch", "envelopePersonal", "fanfoldUS", "fanfoldStdGerman", "fanfoldLegalGerman"]],
            ["ReadingOrder", ["context", "leftToRight", "rightToLeft"]],
            ["BuiltInStyle", ["normal", "comma", "currency", "percent", "wholeComma", "wholeDollar", "hlink", "hlinkTrav", "note", "warningText", "emphasis1", "emphasis2", "emphasis3", "sheetTitle", "heading1", "heading2", "heading3", "heading4", "input", "output", "calculation", "checkCell", "linkedCell", "total", "good", "bad", "neutral", "accent1", "accent1_20", "accent1_40", "accent1_60", "accent2", "accent2_20", "accent2_40", "accent2_60", "accent3", "accent3_20", "accent3_40", "accent3_60", "accent4", "accent4_20", "accent4_40", "accent4_60", "accent5", "accent5_20", "accent5_40", "accent5_60", "accent6", "accent6_20", "accent6_40", "accent6_60", "explanatoryText"]],
            ["PrintErrorType", ["asDisplayed", "blank", "dash", "notAvailable"]],
            ["WorksheetPositionType", ["none", "before", "after", "beginning", "end"]],
            ["PrintComments", ["noComments", "endSheet", "inPlace"]],
            ["PrintOrder", ["downThenOver", "overThenDown"]],
            ["PrintMarginUnit", ["points", "inches", "centimeters"]],
            ["HeaderFooterState", ["default", "firstAndDefault", "oddAndEven", "firstOddAndEven"]],
            ["AutoFillType", ["fillDefault", "fillCopy", "fillSeries", "fillFormats", "fillValues", "fillDays", "fillWeekdays", "fillMonths", "fillYears", "linearTrend", "growthTrend", "flashFill"]],
            ["GroupOption", ["byRows", "byColumns"]],
            ["RangeCopyType", ["all", "formulas", "values", "formats"]],
            ["LinkedDataTypeState", ["none", "validLinkedData", "disambiguationNeeded", "brokenLinkedData", "fetchingData"]],
            ["GeometricShapeType", ["lineInverse", "triangle", "rightTriangle", "rectangle", "diamond", "parallelogram", "trapezoid", "nonIsoscelesTrapezoid", "pentagon", "hexagon", "heptagon", "octagon", "decagon", "dodecagon", "star4", "star5", "star6", "star7", "star8", "star10", "star12", "star16", "star24", "star32", "roundRectangle", "round1Rectangle", "round2SameRectangle", "round2DiagonalRectangle", "snipRoundRectangle", "snip1Rectangle", "snip2SameRectangle", "snip2DiagonalRectangle", "plaque", "ellipse", "teardrop", "homePlate", "chevron", "pieWedge", "pie", "blockArc", "donut", "noSmoking", "rightArrow", "leftArrow", "upArrow", "downArrow", "stripedRightArrow", "notchedRightArrow", "bentUpArrow", "leftRightArrow", "upDownArrow", "leftUpArrow", "leftRightUpArrow", "quadArrow", "leftArrowCallout", "rightArrowCallout", "upArrowCallout", "downArrowCallout", "leftRightArrowCallout", "upDownArrowCallout", "quadArrowCallout", "bentArrow", "uturnArrow", "circularArrow", "leftCircularArrow", "leftRightCircularArrow", "curvedRightArrow", "curvedLeftArrow", "curvedUpArrow", "curvedDownArrow", "swooshArrow", "cube", "can", "lightningBolt", "heart", "sun", "moon", "smileyFace", "irregularSeal1", "irregularSeal2", "foldedCorner", "bevel", "frame", "halfFrame", "corner", "diagonalStripe", "chord", "arc", "leftBracket", "rightBracket", "leftBrace", "rightBrace", "bracketPair", "bracePair", "callout1", "callout2", "callout3", "accentCallout1", "accentCallout2", "accentCallout3", "borderCallout1", "borderCallout2", "borderCallout3", "accentBorderCallout1", "accentBorderCallout2", "accentBorderCallout3", "wedgeRectCallout", "wedgeRRectCallout", "wedgeEllipseCallout", "cloudCallout", "cloud", "ribbon", "ribbon2", "ellipseRibbon", "ellipseRibbon2", "leftRightRibbon", "verticalScroll", "horizontalScroll", "wave", "doubleWave", "plus", "flowChartProcess", "flowChartDecision", "flowChartInputOutput", "flowChartPredefinedProcess", "flowChartInternalStorage", "flowChartDocument", "flowChartMultidocument", "flowChartTerminator", "flowChartPreparation", "flowChartManualInput", "flowChartManualOperation", "flowChartConnector", "flowChartPunchedCard", "flowChartPunchedTape", "flowChartSummingJunction", "flowChartOr", "flowChartCollate", "flowChartSort", "flowChartExtract", "flowChartMerge", "flowChartOfflineStorage", "flowChartOnlineStorage", "flowChartMagneticTape", "flowChartMagneticDisk", "flowChartMagneticDrum", "flowChartDisplay", "flowChartDelay", "flowChartAlternateProcess", "flowChartOffpageConnector", "actionButtonBlank", "actionButtonHome", "actionButtonHelp", "actionButtonInformation", "actionButtonForwardNext", "actionButtonBackPrevious", "actionButtonEnd", "actionButtonBeginning", "actionButtonReturn", "actionButtonDocument", "actionButtonSound", "actionButtonMovie", "gear6", "gear9", "funnel", "mathPlus", "mathMinus", "mathMultiply", "mathDivide", "mathEqual", "mathNotEqual", "cornerTabs", "squareTabs", "plaqueTabs", "chartX", "chartStar", "chartPlus"]],
            ["ConnectorType", ["straight", "elbow", "curve"]],
            ["ContentType", ["plain", "mention"]],
            ["SpecialCellType", ["conditionalFormats", "dataValidations", "blanks", "constants", "formulas", "sameConditionalFormat", "sameDataValidation", "visible"]],
            ["SpecialCellValueType", ["all", "errors", "errorsLogical", "errorsNumbers", "errorsText", "errorsLogicalNumber", "errorsLogicalText", "errorsNumberText", "logical", "logicalNumbers", "logicalText", "logicalNumbersText", "numbers", "numbersText", "text"]],
            ["Placement", ["twoCell", "oneCell", "absolute"]],
            ["FillPattern", ["none", "solid", "gray50", "gray75", "gray25", "horizontal", "vertical", "down", "up", "checker", "semiGray75", "lightHorizontal", "lightVertical", "lightDown", "lightUp", "grid", "crissCross", "gray16", "gray8", "linearGradient", "rectangularGradient"]],
            ["ShapeTextHorizontalAlignment", ["left", "center", "right", "justify", "justifyLow", "distributed", "thaiDistributed"]],
            ["ShapeTextVerticalAlignment", ["top", "middle", "bottom", "justified", "distributed"]],
            ["ShapeTextVerticalOverflow", ["overflow", "ellipsis", "clip"]],
            ["ShapeTextHorizontalOverflow", ["overflow", "clip"]],
            ["ShapeTextReadingOrder", ["leftToRight", "rightToLeft"]],
            ["ShapeTextOrientation", ["horizontal", "vertical", "vertical270", "wordArtVertical", "eastAsianVertical", "mongolianVertical", "wordArtVerticalRTL"]],
            ["ShapeAutoSize", ["autoSizeNone", "autoSizeTextToFitShape", "autoSizeShapeToFitText", "autoSizeMixed"]],
            ["CloseBehavior", ["save", "skipSave"]],
            ["SaveBehavior", ["save", "prompt"]],
            ["SlicerSortType", ["dataSourceOrder", "ascending", "descending"]],
            ["RibbonTab", ["others", "home", "insert", "draw", "pageLayout", "formulas", "data", "review", "view", "developer", "addIns", "help"]],
            ["NumberFormatCategory", ["general", "number", "currency", "accounting", "date", "time", "percentage", "fraction", "scientific", "text", "special", "custom"]]],
        "clientObjectTypes": [[1, 0, [["connectionString", 2], [182, 2], ["commandText", 2], ["dataSourceType", 2]]],
            [2, 0, [["enableEvents", 1, 1]]],
            [3, 0, [["calculationMode", 1], [183, 3, 2], ["calculationState", 3, 2], ["decimalSeparator", 3, 3], ["thousandsSeparator", 3, 3], ["useSystemSeparators", 3, 3]], [["iterativeCalculation", 4, 35, 2, 0, 4], ["ribbon", 177, 35, 2, 0, 4], ["cultureInfo", 90, 3, 3, 0, 4]], [[184, 1], ["suspendApiCalculationUntilNextSync", 0, 1, 4], ["suspendScreenUpdatingUntilNextSync", 0, 0, 2]], [["createWorkbook", 7, 1, 10, 1, 0, 0, 185], [185, 7, 1, 2, 1, 4]]],
            [4, 0, [[186, 1], ["maxIteration", 1], ["maxChange", 1]]],
            [5, 0, [[182, 3, 5], [187, 3, 1], ["isDirty", 1, 2], ["use1904DateSystem", 1, 6], ["chartDataPointTrack", 1, 2], ["usePrecisionAsDisplayed", 1, 2], [183, 3, 2], ["autoSave", 3, 2], ["previouslySaved", 3, 2]], [["worksheets", 9, 19, 0, 0, 4], [188, 18, 19, 0, 0, 4], [189, 23, 19, 0, 0, 4], ["application", 3, 2, 0, 0, 4], ["bindings", 22, 19, 0, 0, 4], ["functions", 181, 2, 7, 0, 4], [96, 96, 2, 8, 0, 4], [190, 98, 19, 8, 0, 4], ["settings", 16, 19, 9, 0, 4], ["customXmlParts", 94, 19, 10, 0, 4], ["internalTest", 149, 2, 4, 0, 4], [191, 115, 35, 5, 0, 4], ["styles", 140, 19, 5, 0, 4], [192, 6, 3, 5, 0, 4], ["dataConnections", 155, 3, 5, 0, 4], ["_Runtime", 2, 2, 10, 0, 4], [193, 157, 19, 11, 0, 4], [194, 174, 19, 11, 0, 4], ["tableStyles", 141, 19, 11, 0, 4], ["pivotTableStyles", 143, 19, 11, 0, 4], ["slicerStyles", 145, 19, 11, 0, 4], ["timelineStyles", 147, 19, 11, 0, 4]], [["_RemoveReference", 1, 2], ["_GetObjectByReferenceId", 1, 2, 0, 4], ["_GetObjectTypeNameByReferenceId", 1, 2, 0, 4], ["_RemoveAllReferences", 0, 2], ["_GetReferenceCount", 0, 2, 0, 4], ["getIsActiveCollabSession", 0, 0, 2], ["registerCustomFunctions", 7, 0, 12, 8], ["_SetOsfControlContainerReadyForCustomFunctions", 0, 0, 12], ["close", 1, 0, 3], ["save", 1, 0, 3], [195, 0, 0, 2], [196, 0, 0, 2], [197, 0, 0, 13], [198, 0, 0, 13]], [["getSelectedRange", 12, 0, 10, 0, 4], ["getActiveCell", 12, 0, 10, 5, 4], ["_GetRangeForEventByReferenceId", 12, 1, 2, 0, 4], ["_GetRangeOrNullObjectForEventByReferenceId", 12, 1, 2, 0, 4], ["getActiveChart", 39, 0, 2, 2, 4], ["getActiveChartOrNullObject", 39, 0, 2, 2, 4], ["getSelectedRanges", 13, 0, 10, 2, 4], ["_GetRangesForEventByReferenceId", 13, 1, 2, 2, 4], ["_GetRangesOrNullObjectForEventByReferenceId", 13, 1, 2, 2, 4], ["getActiveSlicer", 173, 0, 2, 11, 4], ["getActiveSlicerOrNullObject", 173, 0, 2, 11, 4]], 0, 0, [["AutoSaveSettingChanged", 0, 2, "MessageType.workbookAutoSaveSettingChangedEvent", 199, 195, 196], [200, 3, 8, "_CC.office10EventIdDocumentSelectionChangedEvent", 201, 201, 201], ["WACOperationEvent", 2, 13, "MessageType.wacoperationEvent", 199, 197, 198], ["_Message", 3, 5, "_CC.office10EventIdRichApiMessageEvent", 201, 201, 201]]],
            [6, 0, [[202, 3]], 0, [[203, 1], [204, 1]]],
            [7, 0, [[205, 3]], 0, [["open", 0, 2, 0, 4]]],
            [8, 0, [[182, 1], [205, 3], [206, 1], ["visibility", 1], ["tabColor", 1, 5], ["standardWidth", 1, 5], ["standardHeight", 3, 5], ["showGridlines", 5, 1], ["showHeadings", 5, 1], ["enableCalculation", 1, 2]], [["charts", 38, 83, 0, 0, 4], [189, 23, 83, 0, 0, 4], [192, 10, 3, 7, 0, 4], [190, 98, 19, 8, 0, 4], [188, 18, 19, 9, 0, 4], ["freezePanes", 11, 3, 5, 0, 4], ["pageLayout", 150, 35, 2, 0, 4], ["visuals", 81, 83, 11, 0, 4], [207, 161, 19, 2, 0, 4], ["horizontalPageBreaks", 154, 19, 2, 0, 4], ["verticalPageBreaks", 154, 19, 2, 0, 4], [208, 89, 3, 2, 0, 4], [194, 174, 19, 11, 0, 4], [193, 157, 19, 11, 0, 4], ["customProperties", 114, 19, 6, 0, 4], ["sheetViews", 179, 18, 14, 0, 4]], [[209], [210, 0, 2], [184, 1, 0, 4], [211, 0, 0, 5], [212, 0, 0, 5], [213, 0, 0, 5], [214, 0, 0, 5], [215, 0, 0, 5], [216, 0, 0, 5], [217, 0, 0, 5], [218, 0, 0, 5], [219, 0, 0, 1], [220, 0, 0, 1], [221, 3, 0, 2], [222, 0, 0, 6], [223, 0, 0, 6], [224, 0, 0, 2], [225, 0, 0, 2], [226, 0, 0, 11], [227, 0, 0, 11], [228, 0, 0, 11], [229, 0, 0, 11], [230, 0, 0, 11], [231, 0, 0, 11], [232, 0, 0, 3], [233, 0, 0, 3], ["showOutlineLevels", 2, 0, 11]], [[234, 12, 1, 10, 0, 4], [235, 12, 1, 10, 0, 4], [236, 12, 2, 10, 0, 4], [237, 12, 1, 10, 9, 4], ["getRangeByIndexes", 12, 4, 10, 5, 4], ["getPrevious", 8, 1, 10, 10, 4, 0, 238], ["getPreviousOrNullObject", 8, 1, 10, 10, 4, 0, 238], ["getNext", 8, 1, 10, 10, 4, 0, 238], ["getNextOrNullObject", 8, 1, 10, 10, 4, 0, 238], [238, 8, 1, 2, 5, 4], ["copy", 8, 2, 0, 5, 0, 0, 239], [239, 8, 1, 0, 5], ["findAll", 13, 2, 10, 2, 4], ["findAllOrNullObject", 13, 2, 10, 2, 4], [240, 13, 1, 10, 2, 4]], "workbook.worksheets", 0, [[241, 0, 5, 242, 243, 213, 214], [244, 0, 1, 245, 243, 219, 220], [246, 2, 5, 247, 243, 211, 212], [248, 0, 11, 249, 243, 228, 229], [250, 0, 5, 251, 243, 215, 216], [252, 0, 6, 253, 243, 222, 223], [254, 2, 2, 255, 243, 224, 225], [256, 0, 3, 257, 243, 232, 233], [258, 0, 11, 259, 243, 226, 227], [200, 0, 5, 260, 243, 217, 218], [261, 0, 11, 262, 243, 230, 231]]],
            [9, 1, 0, 0, [[263, 1, 2, 9, 4], [264, 0, 3, 5], [265, 0, 3, 5], [213, 0, 3, 5], [214, 0, 0, 5], [215, 0, 0, 5], [216, 0, 0, 5], [266, 0, 3, 5], [267, 0, 3, 5], [219, 0, 0, 1], [220, 0, 0, 1], [211, 0, 3, 2], [212, 0, 3, 2], [217, 0, 0, 2], [218, 0, 0, 2], ["addFromBase64", 4, 0, 6], [222, 0, 0, 6], [223, 0, 0, 6], [224, 0, 0, 2], [225, 0, 0, 2], [226, 0, 3, 11], [227, 0, 3, 11], [228, 0, 3, 11], [229, 0, 3, 11], [230, 0, 0, 11], [231, 0, 0, 11], [232, 0, 0, 3], [233, 0, 0, 3]], [[268, 8, 1, 18, 0, 4], [269, 8, 1, 8], ["getActiveWorksheet", 8, 0, 2, 0, 4], [270, 8, 1, 2, 9, 4], [271, 8, 1, 10, 10, 4], ["getLast", 8, 1, 10, 10, 4]], 0, 8, [[241, 0, 5, 242, 199, 213, 214], [272, 0, 5, "MessageType.worksheetAddedEvent", 199, 264, 265], [244, 0, 1, 245, 199, 219, 220], [246, 2, 2, 247, 199, 211, 212], [248, 0, 11, 249, 199, 228, 229], [250, 0, 5, 251, 199, 215, 216], [273, 0, 5, "MessageType.worksheetDeletedEvent", 199, 266, 267], [252, 0, 6, 253, 199, 222, 223], [254, 2, 2, 255, 199, 224, 225], [256, 0, 3, 257, 199, 232, 233], [258, 0, 11, 259, 199, 226, 227], [200, 0, 2, 260, 199, 217, 218], [261, 0, 11, 262, 199, 230, 231]]],
            [10, 0, [[202, 3], [274, 3]], 0, [[203, 2, 1], [204, 1]]],
            [11, 0, 0, 0, [["unfreeze"], ["freezeAt", 1], ["freezeRows", 1], ["freezeColumns", 1]], [[275, 12, 0, 10, 0, 4], ["getLocationOrNullObject", 12, 0, 10, 0, 4]]],
            [12, 14, [[276, 5], [277, 1, 5], [278, 5], [279, 3], [280, 5], [281, 5], [282, 3], [283, 3], [284, 3], [285, 3], [286, 3], [287, 3], [288, 3], [289, 2], [290, 3], [291, 5, 7], ["hidden", 3, 7], ["rowHidden", 1, 7], ["columnHidden", 1, 7], [292, 3, 5], [293, 3, 5], [294, 1, 5], [295, 1, 5], ["linkedDataTypeState", 3, 2], ["hasSpill", 3, 6], [296, 3, 11], [297, 3, 11], [298, 3, 11], [299, 3, 11], ["savedAsArray", 3, 6], ["numberFormatCategories", 3, 6]], [[300, 32, 35, 0, 0, 4], [301, 8, 2, 0, 0, 4], [302, 86, 3, 7, 0, 4], [303, 118, 19, 4, 0, 4], [304, 30, 35, 1, 0, 4]], [[305, 1], [209, 1], [306, 0, 2], [307, 0, 2], ["merge", 1, 0, 7], ["unmerge", 0, 0, 7], ["_ValidateArraySize", 2, 2, 8, 4], [184, 0, 0, 4], ["showCard", 0, 0, 5], [308, 0, 2, 5, 4], [221, 3, 0, 2], [309, 4, 0, 2], [310, 2, 0, 2], [311, 0, 0, 2], [312, 0, 0, 2], ["getCellProperties", 1, 0, 2], ["getRowProperties", 1, 0, 2], ["getColumnProperties", 1, 0, 2], ["setCellProperties", 1, 0, 2], ["setRowProperties", 1, 0, 2], ["setColumnProperties", 1, 0, 2], ["autoFill", 2, 0, 2], [313, 2, 0, 2], ["flashFill", 0, 0, 2], [314, 1, 0, 11], [315, 1, 0, 11], ["showGroupDetails", 1, 0, 11], ["hideGroupDetails", 1, 0, 11], ["moveTo", 1, 0, 15]], [[236, 12, 2, 10, 0, 4], [235, 12, 1, 10, 0, 4], ["insert", 12, 1, 8], [316, 12, 0, 10, 0, 4], [317, 12, 0, 10, 0, 4], ["getOffsetRange", 12, 2, 10, 0, 4], ["getRow", 12, 1, 10, 0, 4], ["getColumn", 12, 1, 10, 0, 4], [318, 12, 1, 10, 0, 4], ["getBoundingRect", 12, 1, 10, 0, 4], ["getLastCell", 12, 0, 10, 0, 4], ["getLastColumn", 12, 0, 10, 0, 4], ["getLastRow", 12, 0, 10, 0, 4], [319, 12, 1, 10, 9, 4], ["getRowsAbove", 12, 1, 11, 8, 4], ["getRowsBelow", 12, 1, 11, 8, 4], ["getColumnsBefore", 12, 1, 11, 8, 4], ["getColumnsAfter", 12, 1, 11, 8, 4], ["getResizedRange", 12, 2, 11, 8, 4], ["getVisibleView", 14, 0, 2, 8, 4], [237, 12, 1, 10, 9, 4], ["getSurroundingRegion", 12, 0, 10, 5, 4], ["getAbsoluteResizedRange", 12, 2, 10, 5, 4], [320, 24, 1, 6, 2, 4], [321, 12, 2, 10, 2, 4], ["findOrNullObject", 12, 2, 10, 2, 4], ["removeDuplicates", 31, 2, 8, 2], [322, 13, 2, 10, 2, 4], [323, 13, 2, 10, 2, 4], ["getSpillingToRange", 12, 0, 10, 6, 4], ["getSpillParent", 12, 0, 10, 6, 4], ["getSpillingToRangeOrNullObject", 12, 0, 10, 6, 4], ["getSpillParentOrNullObject", 12, 0, 10, 6, 4], ["getMergedAreas", 13, 0, 10, 6, 4], ["getPivotTables", 97, 1, 6, 6, 4]]],
            [13, 2, [[289, 2], [286, 3], [287, 3], ["areaCount", 3], [288, 3], [292, 3], [293, 3], [295, 1]], [[324, 156, 19, 0, 0, 4], [303, 118, 19, 0, 0, 4], [300, 32, 35, 0, 0, 4], [304, 30, 35, 0, 0, 4], [301, 8, 2, 0, 0, 4]], [[307, 0, 2], [184], [305, 1], [312], [309, 4], [310, 2], [311], [306, 0, 2, 6]], [[316, 13, 0, 10, 0, 4], [317, 13, 0, 10, 0, 4], [318, 13, 1, 10, 0, 4], [319, 13, 1, 10, 0, 4], ["getOffsetRangeAreas", 13, 2, 10, 0, 4], ["getUsedRangeAreas", 13, 1, 10, 0, 4], ["getUsedRangeAreasOrNullObject", 13, 1, 10, 0, 4], [320, 24, 1, 6, 0, 4], [322, 13, 2, 10, 0, 4], [323, 13, 2, 10, 0, 4]]],
            [14, 0, [[276, 1], [278, 1], [279, 3], [280, 1], [281, 1], [291, 1], [290, 3], [284, 3], [285, 3], ["cellAddresses", 3], [325, 3]], [[326, 15, 19, 0, 0, 4]], 0, [[234, 12, 0, 10, 0, 4]]],
            [15, 1, 0, 0, [[263, 0, 2, 9, 4]], [[327, 14, 1, 2, 0, 4]], 0, 14],
            [16, 1, 0, 0, [[263, 0, 2, 0, 4]], [[268, 17, 1, 18, 0, 4], [269, 17, 2, 9], [270, 17, 1, 2, 0, 4]], 0, 17, [["SettingsChanged", 3, 0, "_CC.office10EventIdSettingsChangedEvent", 201, 201, 201]]],
            [17, 8, [[328, 3], [329, 5], [330, 2]], 0, [[209]]],
            [18, 1, 0, 0, [[263, 0, 2, 9, 4]], [[268, 19, 1, 18, 0, 4], [270, 19, 1, 2, 9, 4], [269, 19, 3, 8, 9], ["addFormulaLocal", 19, 3, 0, 9]], 0, 19],
            [19, 0, [[182, 3], [331, 3], [329, 3], [332, 1], [330, 2], ["comment", 1, 9], ["scope", 3, 9], [333, 1, 5]], [[301, 8, 2, 9, 0, 4], ["worksheetOrNullObject", 8, 2, 9, 0, 4], ["arrayValues", 20, 3, 5, 0, 4]], [[209, 0, 0, 9]], [[234, 12, 0, 10, 0, 4], [334, 12, 0, 10, 9, 4]]],
            [20, 0, [[278, 3], ["types", 3]]],
            [21, 0, [[205, 3], [331, 3]], 0, [["getText", 0, 2, 0, 4], [209, 0, 0, 8]], [["getTable", 25, 0, 2, 0, 4], [234, 12, 0, 2, 0, 4]], 0, 0, [["DataChanged", 3, 8, "_CC.office10EventIdBindingDataChangedEvent", 243, 201, 201], [200, 3, 8, "_CC.office10EventIdBindingSelectionChangedEvent", 243, 201, 201]]],
            [22, 1, [[335, 3]], 0, [[263, 0, 2, 9, 4]], [[268, 21, 1, 18, 0, 4], [327, 21, 1, 2, 0, 4], [269, 21, 3, 8, 8], ["addFromNamedItem", 21, 3, 0, 8], ["addFromSelection", 21, 2, 0, 8], [270, 21, 1, 2, 9, 4]], 0, 21],
            [23, 5, [[335, 3]], 0, [[263, 0, 2, 9, 4], [211, 0, 0, 5], [212, 0, 0, 5], [264, 0, 0, 2], [265, 0, 0, 2], [266, 0, 0, 2], [267, 0, 0, 2], [222, 0, 0, 6], [223, 0, 0, 6]], [[268, 25, 1, 18, 0, 4], [327, 25, 1, 2, 0, 4], [269, 25, 2, 8], [270, 25, 1, 2, 9, 4]], 0, 25, [[272, 0, 2, "MessageType.tableAddedEvent", 336, 264, 265], [246, 2, 5, 337, 336, 211, 212], [273, 0, 2, "MessageType.tableDeletedEvent", 336, 266, 267], [252, 0, 6, 338, 336, 222, 223]]],
            [24, 1, 0, 0, [[263, 0, 2, 0, 4]], [[271, 25, 0, 10, 0, 4], [268, 25, 1, 18, 0, 4]], 0, 25],
            [25, 24, [[205, 3], [182, 1], ["showHeaders", 1], ["showTotals", 1], [295, 1], ["highlightFirstColumn", 1, 8], ["highlightLastColumn", 1, 8], ["showBandedRows", 1, 8], ["showBandedColumns", 1, 8], ["showFilterButton", 1, 8], ["legacyId", 3, 1]], [[339, 26, 19, 0, 0, 4], [326, 28, 19, 0, 0, 4], [302, 87, 3, 7, 0, 4], [301, 8, 2, 7, 0, 4], [208, 89, 3, 2, 0, 4], ["tableStyle", 142, 35, 6, 0, 4]], [[209], [340, 0, 0, 7], ["reapplyFilters", 0, 0, 7], [217, 0, 0, 5], [218, 0, 0, 5], [211, 0, 0, 5], [212, 0, 0, 5], ["clearStyle", 0, 0, 6], [222, 0, 0, 6], [223, 0, 0, 6], [341, 1, 0, 6]], [[234, 12, 0, 10, 0, 4], [342, 12, 0, 10, 0, 4], [343, 12, 0, 10, 0, 4], [344, 12, 0, 10, 0, 4], ["convertToRange", 12, 0, 8, 7]], "workbook.tables", 0, [[246, 2, 5, 337, 243, 211, 212], [252, 0, 6, 338, 243, 222, 223], [200, 2, 5, "MessageType.tableSelectionChangedEvent", 243, 217, 218]]],
            [26, 1, [[335, 3]], 0, [[263, 0, 2, 9, 4]], [[268, 27, 1, 18, 0, 4], [327, 27, 1, 2, 0, 4], [269, 27, 3, 8], [270, 27, 1, 2, 9, 4]], 0, 27],
            [27, 0, [[205, 3], [325, 3], [278, 1], [182, 1]], [["filter", 88, 3, 7, 0, 4]], [[209]], [[234, 12, 0, 10, 0, 4], [342, 12, 0, 10, 0, 4], [343, 12, 0, 10, 0, 4], [344, 12, 0, 10, 0, 4]]],
            [28, 1, [[335, 3]], 0, [[263, 0, 2, 9, 4]], [[327, 29, 1, 2, 0, 4], [269, 29, 2, 8]], 0, 29],
            [29, 0, [[325, 3], [278, 1]], 0, [[209]], [[234, 12, 0, 10, 0, 4]]],
            [30, 0, [[331, 3], [345, 1], ["prompt", 1], ["errorAlert", 1], ["ignoreBlanks", 1], ["valid", 3]], 0, [[305]], [["getInvalidCells", 13, 0, 10, 2, 4], ["getInvalidCellsOrNullObject", 13, 0, 10, 2, 4]]],
            [31, 0, [["removed", 3], ["uniqueRemaining", 3]]],
            [32, 0, [[346, 1], [347, 1], [348, 1], ["columnWidth", 1, 7], [349, 1, 7], [350, 1, 5], ["useStandardHeight", 1, 5], ["useStandardWidth", 1, 5], [351, 1, 2], [352, 1, 2], [353, 1, 2], [354, 1, 2]], [[355, 34, 35, 0, 0, 4], [356, 37, 35, 0, 0, 4], [357, 36, 51, 0, 0, 4], [192, 33, 35, 7, 0, 4]], [["autofitColumns", 0, 0, 7], ["autofitRows", 0, 0, 7], ["adjustIndent", 1, 0, 15]]],
            [33, 0, [[358, 1], [359, 1]]],
            [34, 0, [[360, 1], [361, 1, 2], ["patternTintAndShade", 1, 2], ["pattern", 1, 2], ["patternColor", 1, 2]], 0, [[305]]],
            [35, 0, [[362, 3], [295, 1], [363, 1], [360, 1], [361, 1, 2]]],
            [36, 1, [[335, 3], [361, 1, 2]], 0, 0, [[268, 35, 1, 18, 0, 4], [327, 35, 1, 2, 0, 4]], 0, 35],
            [37, 0, [[182, 1], [364, 1], [360, 1], [365, 1], [366, 1], [367, 1], [368, 1, 2], ["subscript", 1, 2], ["superscript", 1, 2], [361, 1, 2]]],
            [38, 5, [[335, 3]], 0, [[263, 0, 2, 9, 4], [264, 0, 0, 1], [265, 0, 0, 1], [213, 0, 0, 1], [214, 0, 0, 1], [215, 0, 0, 1], [216, 0, 0, 1], [266, 0, 0, 1], [267, 0, 0, 1]], [[269, 39, 3, 9], [327, 39, 1, 2, 0, 4], [369, 39, 1, 18, 0, 4], [268, 39, 1, 2, 0, 4], [270, 39, 1, 2, 9, 4]], 0, 39, [[241, 0, 1, 370, 371, 213, 214], [272, 0, 1, "MessageType.chartAddedEvent", 371, 264, 265], [250, 0, 1, 372, 371, 215, 216], [273, 0, 1, "MessageType.chartDeletedEvent", 371, 266, 267]]],
            [39, 0, [[182, 1], [296, 1], [297, 1], [299, 1], [298, 1], [205, 3, 5], ["showAllFieldButtons", 1, 5], [373, 1, 5], ["showDataLabelsOverMaximum", 1, 1], ["categoryLabelLevel", 1, 1], [295, 1, 1], ["displayBlanksAs", 1, 1], ["plotBy", 1, 1], ["plotVisibleOnly", 1, 1], ["seriesNameLevel", 1, 1]], [[374, 65, 35, 0, 0, 4], [375, 53, 35, 0, 0, 4], ["legend", 60, 35, 0, 0, 4], ["series", 42, 19, 0, 0, 4], ["axes", 48, 35, 0, 0, 4], [300, 41, 35, 0, 0, 4], [301, 8, 2, 7, 0, 4], ["plotArea", 79, 35, 1, 0, 4], ["pivotOptions", 40, 35, 2, 0, 4]], [["setData", 2, 1], [209], ["setPosition", 2], [308, 3, 2, 7, 4], [213, 0, 0, 1], [214, 0, 0, 1], [215, 0, 0, 1], [216, 0, 0, 1], [210, 0, 2, 2]], 0, 0, 0, [[241, 0, 1, 370, 243, 213, 214], [250, 0, 1, 372, 243, 215, 216]]],
            [40, 0, [["showAxisFieldButtons", 1], ["showLegendFieldButtons", 1], ["showReportFilterFieldButtons", 1], ["showValueFieldButtons", 1]]],
            [41, 0, [["roundedCorners", 1, 2], ["colorScheme", 1, 2]], [[355, 68, 3, 0, 0, 4], [356, 73, 35, 0, 0, 4], [376, 69, 35, 5, 0, 4]]],
            [42, 1, [[335, 3]], 0, [[263, 0, 2, 9, 4]], [[327, 43, 1, 2, 0, 4], [269, 43, 2, 8, 5]], 0, 43],
            [43, 0, [[182, 1], [373, 1, 5], ["hasDataLabels", 1, 5], ["filtered", 1, 5], [377, 1, 5], [378, 1, 5], [379, 1, 5], [380, 1, 5], [381, 1, 5], ["smooth", 1, 5], ["plotOrder", 1, 5], ["gapWidth", 1, 5], ["doughnutHoleSize", 1, 5], [382, 1, 1], ["explosion", 1, 1], ["firstSliceAngle", 1, 1], ["invertIfNegative", 1, 1], ["bubbleScale", 1, 2], ["secondPlotSize", 1, 1], ["splitType", 1, 1], ["splitValue", 1, 2], ["varyByCategories", 1, 1], ["showLeaderLines", 1, 2], ["overlap", 1, 1], ["gradientStyle", 1, 2], ["gradientMinimumType", 1, 2], ["gradientMidpointType", 1, 2], ["gradientMaximumType", 1, 2], ["gradientMinimumValue", 1, 2], ["gradientMidpointValue", 1, 2], ["gradientMaximumValue", 1, 2], ["gradientMinimumColor", 1, 2], ["gradientMidpointColor", 1, 2], ["gradientMaximumColor", 1, 2], ["parentLabelStrategy", 1, 2], ["showConnectorLines", 1, 2], ["invertColor", 1, 2]], [["points", 45, 19, 0, 0, 4], [300, 44, 35, 0, 0, 4], ["trendlines", 75, 19, 5, 0, 4], ["xErrorBars", 56, 35, 2, 0, 4], ["yErrorBars", 56, 35, 2, 0, 4], [375, 53, 35, 1, 0, 4], ["binOptions", 70, 35, 2, 0, 4], ["mapOptions", 64, 35, 2, 0, 4], ["boxwhiskerOptions", 71, 35, 2, 0, 4]], [[209, 0, 0, 5], ["setXAxisValues", 1, 0, 5], ["setValues", 1, 0, 5], ["setBubbleSizes", 1, 0, 5], ["getDimensionValues", 1, 0, 6]]],
            [44, 0, 0, [[355, 68, 3, 0, 0, 4], [383, 72, 35, 0, 0, 4]]],
            [45, 1, [[335, 3]], 0, [[263, 0, 2, 9, 4]], [[327, 46, 1, 2, 0, 4]], 0, 46],
            [46, 0, [[329, 3], ["hasDataLabel", 1, 5], [378, 1, 5], [377, 1, 5], [380, 1, 5], [381, 1, 5]], [[300, 47, 35, 0, 0, 4], ["dataLabel", 54, 35, 5, 0, 4]]],
            [47, 0, 0, [[355, 68, 3, 0, 0, 4], [376, 69, 35, 5, 0, 4]]],
            [48, 0, 0, [["categoryAxis", 49, 35, 0, 0, 4], ["seriesAxis", 49, 35, 0, 0, 4], ["valueAxis", 49, 35, 0, 0, 4]], 0, [[268, 49, 2, 2, 5, 4]]],
            [49, 0, [["majorUnit", 1], ["maximum", 1], ["minimum", 1], ["minorUnit", 1], ["displayUnit", 1, 5], ["showDisplayUnitLabel", 1, 5], ["customDisplayUnit", 3, 5], [331, 3, 5], ["minorTimeUnitScale", 1, 5], ["majorTimeUnitScale", 1, 5], ["baseTimeUnit", 1, 5], ["categoryType", 1, 5], [382, 3, 5], ["scaleType", 1, 5], ["logBase", 1, 5], [297, 3, 5], [296, 3, 5], [298, 3, 5], [299, 3, 5], ["reversePlotOrder", 1, 5], ["crosses", 1, 5], ["crossesAt", 3, 5], [332, 1, 5], ["isBetweenCategories", 1, 1], ["majorTickMark", 1, 5], ["minorTickMark", 1, 5], ["tickMarkSpacing", 1, 5], ["tickLabelPosition", 1, 5], ["tickLabelSpacing", 1, 5], ["alignment", 1, 1], ["multiLevel", 1, 1], [276, 1, 1], [384, 1, 2], ["offset", 1, 1], [350, 1, 1], [206, 1, 1], ["positionAt", 3, 1]], [["majorGridlines", 58, 35, 0, 0, 4], ["minorGridlines", 58, 35, 0, 0, 4], [374, 51, 35, 0, 0, 4], [300, 50, 35, 0, 0, 4]], [["setCategoryNames", 1, 0, 5], ["setCustomDisplayUnit", 1, 0, 5], ["setCrossesAt", 1, 0, 5], ["setPositionAt", 1, 0, 1]]],
            [50, 0, 0, [[356, 73, 35, 0, 0, 4], [383, 72, 35, 0, 0, 4], [355, 68, 3, 1, 0, 4]]],
            [51, 0, [[279, 1], [332, 1], [350, 1, 6]], [[300, 52, 35, 0, 0, 4]], [[385, 1, 0, 1]]],
            [52, 0, 0, [[356, 73, 35, 0, 0, 4], [355, 68, 3, 1, 0, 4], [376, 69, 35, 1, 0, 4]]],
            [53, 0, [[206, 1], [386, 1], [387, 1], [388, 1], [389, 1], [390, 1], [391, 1], [392, 1], [276, 1, 1], [384, 1, 2], [350, 1, 1], [393, 1, 1], [347, 1, 1], [348, 1, 1]], [[300, 55, 35, 0, 0, 4]]],
            [54, 0, [[206, 1], [386, 1], [387, 1], [388, 1], [389, 1], [390, 1], [391, 1], [392, 1], [296, 1, 1], [297, 1, 1], [299, 3, 1], [298, 3, 1], [333, 1, 1], [350, 1, 1], [347, 1, 1], [348, 1, 1], [279, 1, 1], [393, 1, 1], [276, 1, 1], [384, 1, 2]], [[300, 55, 35, 1, 0, 4]]],
            [55, 0, 0, [[356, 73, 35, 0, 0, 4], [355, 68, 3, 0, 0, 4], [376, 69, 35, 1, 0, 4]]],
            [56, 0, [["endStyleCap", 1], ["include", 1], [331, 1], [332, 1]], [[300, 57, 35, 0, 0, 4]]],
            [57, 0, 0, [[383, 72, 35, 0, 0, 4]]],
            [58, 0, [[332, 1]], [[300, 59, 35, 0, 0, 4]]],
            [59, 0, 0, [[383, 72, 35, 0, 0, 4]]],
            [60, 0, [[332, 1], [206, 1], [394, 1], [297, 1, 5], [296, 1, 5], [299, 1, 5], [298, 1, 5], [379, 1, 5]], [[300, 63, 35, 0, 0, 4], ["legendEntries", 62, 19, 5, 0, 4]]],
            [61, 0, [[332, 1], [297, 3, 1], [296, 3, 1], [299, 3, 1], [298, 3, 1], [325, 3, 1]]],
            [62, 1, 0, 0, [[263, 0, 2, 0, 4]], [[327, 61, 1, 2, 0, 4]], 0, 61],
            [63, 0, 0, [[356, 73, 35, 0, 0, 4], [355, 68, 3, 0, 0, 4], [376, 69, 35, 1, 0, 4]]],
            [64, 0, [[395, 1], ["labelStrategy", 1], ["projectionType", 1]]],
            [65, 0, [[332, 1], [279, 1], [394, 1], [347, 1, 5], [296, 1, 5], [297, 1, 5], [299, 3, 5], [298, 3, 5], [348, 1, 5], [350, 1, 5], [206, 1, 5], [379, 1, 5]], [[300, 67, 35, 0, 0, 4]], [[385, 1, 0, 5]], [[396, 66, 2, 2, 5, 4]]],
            [66, 0, 0, [[356, 73, 35, 0, 0, 4]]],
            [67, 0, 0, [[356, 73, 35, 0, 0, 4], [355, 68, 3, 0, 0, 4], [376, 69, 35, 5, 0, 4]]],
            [68, 4, 0, 0, [[397, 1], [305]]],
            [69, 0, [[360, 1], [398, 1], [363, 1]], 0, [[305, 0, 0, 1]]],
            [70, 0, [[331, 1], [299, 1], [335, 1], ["allowOverflow", 1], ["allowUnderflow", 1], ["overflowValue", 1], ["underflowValue", 1]]],
            [71, 0, [["showInnerPoints", 1], ["showOutlierPoints", 1], ["showMeanMarker", 1], ["showMeanLine", 1], ["quartileCalculation", 1]]],
            [72, 0, [[360, 1], [398, 1, 5], [363, 1, 5]], 0, [[305]]],
            [73, 0, [[366, 1], [360, 1], [365, 1], [182, 1], [364, 1], [367, 1]]],
            [74, 0, [[331, 1], ["polynomialOrder", 1], ["movingAveragePeriod", 1], [330, 2], ["showEquation", 1, 1], ["showRSquared", 1, 1], ["forwardPeriod", 1, 1], ["backwardPeriod", 1, 1], [182, 1], ["intercept", 1]], [[300, 76, 35, 0, 0, 4], ["label", 77, 35, 1, 0, 4]], [[209]]],
            [75, 1, 0, 0, [[263, 0, 2, 0, 4]], [[269, 74, 1, 8], [268, 74, 1, 18, 0, 4]], 0, 74],
            [76, 0, 0, [[383, 72, 35, 0, 0, 4]]],
            [77, 0, [[296, 1], [297, 1], [299, 3], [298, 3], [333, 1], [350, 1], [347, 1], [348, 1], [279, 1], [393, 1], [276, 1], [384, 1, 2]], [[300, 78, 35, 0, 0, 4]]],
            [78, 0, 0, [[355, 68, 3, 0, 0, 4], [376, 69, 35, 0, 0, 4], [356, 73, 35, 0, 0, 4]]],
            [79, 0, [[297, 1], [296, 1], [299, 1], [298, 1], ["insideLeft", 1], ["insideTop", 1], ["insideWidth", 1], ["insideHeight", 1], [206, 1]], [[300, 80, 35, 0, 0, 4]]],
            [80, 0, 0, [[376, 69, 35, 0, 0, 4], [355, 68, 3, 0, 0, 4]]],
            [81, 5, 0, 0, [["getDefinitions", 0, 2, 0, 4], ["getPreview", 4, 2, 0, 4], ["bootstrapAgaveVisual", 0, 0, 0, 2], [263, 0, 2, 0, 4], [217, 0, 2], [218, 0, 2]], [[269, 82, 3, 8, 0, 2], [369, 82, 1, 18, 0, 4], ["getSelectedOrNullObject", 82, 0, 2, 0, 4]], 0, 82, [["AgaveVisualUpdate", 2, 0, "MessageType.agaveVisualUpdateEvent", 201, 399, 399], [200, 0, 0, "MessageType.visualSelectionChangedEvent", 371, 217, 218]]],
            [82, 0, [[205, 3], ["isSupportedInVisualTaskpane", 3]], [[191, 84, 18, 0, 0, 4]], [[209, 0, 0, 0, 2], ["getProperty", 1, 2, 0, 4], ["setProperty", 2, 0, 0, 2], ["changeDataSource", 2, 0, 0, 2], ["getDataSource", 0, 2, 0, 4], ["setPropertyToDefault", 1, 0, 0, 2], [400, 0, 2], [401, 0, 2], ["serializeProperties", 0, 2, 0, 4], ["deserializeProperties", 1, 0, 0, 2]], [["getChildProperties", 84, 2, 6, 0, 4], ["getDataControllerClient", 85, 0, 2, 0, 4], ["getElementChildProperties", 84, 3, 6, 0, 4]], 0, 0, [["ChangeNotification", 2, 0, "MessageType.visualChangeEvent", 243, 400, 401]]],
            [83, 0, [[331, 3], [329, 3], [205, 3], ["localizedName", 3], [274, 3], ["localizedOptions", 3], ["hasDefault", 3], ["isDefault", 3], [402, 3], [403, 3], ["stepSize", 3], ["hideMeButShowChildrenUI", 3], ["expandableUI", 3], ["nextPropOnSameLine", 3], ["showResetUI", 3]], 0, [["getBoolMetaProperty", 1, 2, 0, 4]]],
            [84, 1, 0, 0, [[263, 0, 2, 0, 4]], [[268, 83, 1, 18, 0, 4], [327, 83, 1, 2, 0, 4]], 0, 83],
            [85, 0, 0, 0, [["getWells", 0, 2, 0, 4], ["getAssociatedFields", 1, 2, 0, 4], ["getAvailableFields", 1, 2, 0, 4], ["addField", 3, 0, 0, 2], ["removeField", 2, 0, 0, 2], ["moveField", 3, 0, 0, 2]]],
            [86, 0, 0, 0, [[404, 5]]],
            [87, 0, [["matchCase", 3], ["method", 3], [405, 3]], 0, [[404, 3], [305], [406]]],
            [88, 0, [[407, 3]], 0, [[404, 1], [305], ["applyBottomItemsFilter", 1], ["applyBottomPercentFilter", 1], ["applyCellColorFilter", 1], ["applyDynamicFilter", 1], ["applyFontColorFilter", 1], ["applyValuesFilter", 1], ["applyTopItemsFilter", 1], ["applyTopPercentFilter", 1], ["applyIconFilter", 1], ["applyCustomFilter", 3]]],
            [89, 0, [[186, 3], ["isDataFiltered", 3], [407, 3]], 0, [[404, 3], [406], [408], ["clearCriteria"]], [[234, 12, 0, 10, 0, 4], [334, 12, 0, 10, 0, 4]]],
            [90, 0, [[182, 3]], [[276, 91, 3, 0, 0, 4], ["datetimeFormat", 92, 3, 6, 0, 4]]],
            [91, 0, [["numberDecimalSeparator", 3], ["numberGroupSeparator", 3]]],
            [92, 0, [["dateSeparator", 3], ["longDatePattern", 3], ["shortDatePattern", 3], ["timeSeparator", 3], ["longTimePattern", 3]]],
            [93, 1, 0, 0, [[263, 0, 2, 0, 4]], [[268, 95, 1, 18, 0, 4], [270, 95, 1, 2, 0, 4], ["getOnlyItem", 95, 0, 2, 0, 4], ["getOnlyItemOrNullObject", 95, 0, 2, 0, 4]], 0, 95],
            [94, 1, 0, 0, [[263, 0, 2, 0, 4]], [[268, 95, 1, 18, 0, 4], [269, 95, 1, 8], ["getByNamespace", 93, 1, 6, 0, 4], [270, 95, 1, 2, 0, 4]], 0, 95],
            [95, 0, [[205, 3], ["namespaceUri", 3]], 0, [[209], ["getXml", 0, 2, 0, 4], ["setXml", 1]]],
            [96, 0, 0, 0, [["bindingGetData", 1, 2, 0, 4], ["getSelectedData", 1, 2, 0, 4], ["gotoById", 1, 2, 0, 4], ["bindingAddFromSelection", 1, 2], ["bindingGetById", 1, 2, 0, 4], ["bindingReleaseById", 1, 2], ["bindingGetAll", 0, 2, 0, 4], ["bindingAddFromNamedItem", 1, 2], ["bindingAddFromPrompt", 1, 2], ["bindingDeleteAllDataValues", 1], ["setSelectedData", 1], ["bindingClearFormats", 1], ["bindingSetData", 1], ["bindingSetFormats", 1], ["bindingSetTableOptions", 1], ["bindingAddRows", 1], ["bindingAddColumns", 1], ["getFilePropertiesAsync", 0, 2, 4, 4]]],
            [97, 1, 0, 0, [[263, 0, 2, 0, 4]], [[268, 99, 1, 18, 0, 4], [271, 99, 0, 10, 0, 4], [270, 99, 1, 2, 0, 4]], 0, 99],
            [98, 1, 0, 0, [[409], [263, 0, 2, 9, 4]], [[268, 99, 1, 18, 0, 4], [270, 99, 1, 2, 9, 4], [269, 99, 3, 8, 1]], 0, 99],
            [99, 0, [[182, 1], [205, 3, 10], ["useCustomSortLists", 1, 2], ["enableDataValueEditing", 1, 2], ["allowMultipleFiltersPerField", 0, 3]], [[301, 8, 2, 0, 0, 4], ["hierarchies", 101, 19, 1, 0, 4], ["rowHierarchies", 103, 19, 1, 0, 4], ["columnHierarchies", 103, 19, 1, 0, 4], ["dataHierarchies", 107, 19, 1, 0, 4], ["filterHierarchies", 105, 19, 1, 0, 4], ["layout", 100, 2, 1, 0, 4]], [["refresh"], [209, 0, 0, 1]]],
            [100, 0, [["showColumnGrandTotals", 1], ["showRowGrandTotals", 1], ["enableFieldList", 1, 11], ["subtotalLocation", 1], ["layoutType", 1], ["autoFormat", 1, 2], ["preserveFormatting", 1, 2]], [["pivotStyle", 144, 35, 6, 0, 4]], [["getPivotItems", 2, 0, 2], ["setAutoSortOnCell", 2, 0, 2], [341, 1, 0, 6]], [[234, 12], ["getRowLabelRange", 12], ["getColumnLabelRange", 12], ["getFilterAxisRange", 12], [343, 12], [236, 12, 3, 0, 6], ["getDataHierarchy", 108, 1, 0, 2]]],
            [101, 1, 0, 0, [[263, 0, 2, 0, 4]], [[268, 102, 1, 18, 0, 4], [270, 102, 1, 2, 0, 4]], 0, 102],
            [102, 0, [[205, 3], [182, 1]], [[405, 109, 19, 0, 0, 4]]],
            [103, 1, 0, 0, [[263, 0, 2, 0, 4], [408, 1]], [[268, 104, 1, 18, 0, 4], [270, 104, 1, 2, 0, 4], [269, 104, 1, 8]], 0, 104],
            [104, 0, [[205, 3], [182, 1], [206, 1]], [[405, 109, 19, 0, 0, 4]], [[410]]],
            [105, 1, 0, 0, [[263, 0, 2, 0, 4], [408, 1]], [[268, 106, 1, 18, 0, 4], [270, 106, 1, 2, 0, 4], [269, 106, 1, 8]], 0, 106],
            [106, 0, [[205, 3], [182, 1], [206, 1], ["enableMultipleFilterItems", 1]], [[405, 109, 19, 0, 0, 4]], [[410]]],
            [107, 1, 0, 0, [[263, 0, 2, 0, 4], [408, 1]], [[268, 108, 1, 18, 0, 4], [270, 108, 1, 2, 0, 4], [269, 108, 1, 8]], 0, 108],
            [108, 0, [[205, 3], [182, 1], [206, 1], [276, 1], ["summarizeBy", 1], ["showAs", 1]], [["field", 110, 35, 0, 0, 4]], [[410]]],
            [109, 1, 0, 0, [[263, 0, 2, 0, 4]], [[268, 110, 1, 18, 0, 4], [270, 110, 1, 2, 0, 4]], 0, 110],
            [110, 0, [[205, 3], [182, 1], ["subtotals", 1], ["showAllItems", 1]], [["items", 111, 19, 0, 0, 4]], [["sortByLabels", 1, 1], ["sortByValues", 3, 0, 2], ["applyFilter", 1, 0, 3], ["clearAllFilters", 0, 0, 3], ["clearFilter", 1, 0, 3], ["getFilters", 0, 0, 3], ["isFiltered", 1, 0, 3]]],
            [111, 1, 0, 0, [[263, 0, 2, 0, 4]], [[268, 112, 1, 18, 0, 4], [270, 112, 1, 2, 0, 4]], 0, 112],
            [112, 0, [[205, 3], [182, 1], ["isExpanded", 1], [332, 1]]],
            [113, 0, [[328, 3], [329, 3]]],
            [114, 1, 0, 0, [[263, 0, 2, 0, 4]], [[268, 113, 1, 18, 0, 4], [270, 113, 1]], 0, 113],
            [115, 0, [[374, 1], ["subject", 1], ["author", 1], ["keywords", 1], [193, 1], ["lastAuthor", 3], ["revisionNumber", 1], [411, 11], ["category", 1], ["manager", 1], ["company", 1]], [[412, 117, 19, 0, 0, 4]]],
            [116, 0, [[328, 3], [329, 1], [331, 3]], 0, [[209]]],
            [117, 1, 0, 0, [[263, 0, 2, 0, 4], ["deleteAll"]], [[270, 116, 1, 2, 0, 4], [269, 116, 2, 8], [268, 116, 1, 18, 0, 4]], 0, 116],
            [118, 1, 0, 0, [[263, 0, 2, 0, 4], ["clearAll"]], [[327, 119, 1, 2, 0, 4], [269, 119, 1, 8], [268, 119, 1, 18, 0, 4]], 0, 119],
            [119, 0, [["stopIfTrue", 1], ["priority", 1], [331, 3], [205, 3]], [["dataBarOrNullObject", 120, 35, 0, 0, 4], ["dataBar", 120, 35, 0, 0, 4], ["customOrNullObject", 123, 35, 0, 0, 4], [412, 123, 35, 0, 0, 4], ["iconSet", 125, 35, 0, 0, 4], ["iconSetOrNullObject", 125, 35, 0, 0, 4], ["colorScale", 126, 35, 0, 0, 4], ["colorScaleOrNullObject", 126, 35, 0, 0, 4], ["topBottom", 127, 35, 0, 0, 4], ["topBottomOrNullObject", 127, 35, 0, 0, 4], ["preset", 128, 35, 0, 0, 4], ["presetOrNullObject", 128, 35, 0, 0, 4], ["textComparison", 129, 35, 0, 0, 4], ["textComparisonOrNullObject", 129, 35, 0, 0, 4], ["cellValue", 130, 35, 0, 0, 4], ["cellValueOrNullObject", 130, 35, 0, 0, 4]], [[209]], [[234, 12, 0, 10, 0, 4], [334, 12, 0, 10, 0, 4], [240, 13, 0, 10, 2, 4]]],
            [120, 0, [["showDataBarOnly", 1], ["barDirection", 1], ["axisFormat", 1], ["axisColor", 1], ["lowerBoundRule", 1], ["upperBoundRule", 1]], [["positiveFormat", 121, 35, 0, 0, 4], ["negativeFormat", 122, 35, 0, 0, 4]]],
            [121, 0, [[413, 1], ["gradientFill", 1], [414, 1]]],
            [122, 0, [[413, 1], ["matchPositiveFillColor", 1], [414, 1], ["matchPositiveBorderColor", 1]]],
            [123, 0, 0, [[345, 124, 35, 0, 0, 4], [300, 131, 35, 0, 0, 4]]],
            [124, 0, [[333, 1], ["formulaLocal", 1], ["formulaR1C1", 1]]],
            [125, 0, [["reverseIconOrder", 1], ["showIconOnly", 1], [295, 1], [407, 1]]],
            [126, 0, [["threeColorScale", 3], [407, 1]]],
            [127, 0, [[345, 1]], [[300, 131, 35, 0, 0, 4]]],
            [128, 0, [[345, 1]], [[300, 131, 35, 0, 0, 4]]],
            [129, 0, [[345, 1]], [[300, 131, 35, 0, 0, 4]]],
            [130, 0, [[345, 1]], [[300, 131, 35, 0, 0, 4]]],
            [131, 0, [[276, 1]], [[355, 133, 35, 0, 0, 4], [356, 132, 35, 0, 0, 4], [357, 135, 51, 0, 0, 4]]],
            [132, 0, [[360, 1], [365, 1], [366, 1], [367, 1], [368, 1]], 0, [[305]]],
            [133, 0, [[360, 1]], 0, [[305]]],
            [134, 0, [[362, 3], [295, 1], [360, 1]]],
            [135, 1, [[335, 3]], [[296, 134, 35, 0, 0, 4], ["bottom", 134, 35, 0, 0, 4], [297, 134, 35, 0, 0, 4], [415, 134, 35, 0, 0, 4]], 0, [[268, 134, 1, 18, 0, 4], [327, 134, 1, 2, 0, 4]], 0, 134],
            [136, 0, 0, 0, 0, [["getFormatter", 137, 1]], 0, 0, 0, "Microsoft.ExcelServices.NumberFormattingService", 4],
            [137, 0, [["isDateTime", 3], ["isPercent", 3], ["isCurrency", 3], ["isNumeric", 3], [416, 3], ["hasYear", 3], ["hasMonth", 3], ["hasDayOfWeek", 3]], 0, [[300, 1]]],
            [138, 36, [["status", 2]], 0, [["register", 2]], 0, 0, 0, 0, "Microsoft.ExcelServices.CustomFunctionManager", 4],
            [139, 0, [["builtIn", 3], [359, 1], [347, 1], ["includeAlignment", 1], ["includeBorder", 1], ["includeFont", 1], ["includeNumber", 1], ["includePatterns", 1], ["includeProtection", 1], [353, 1], [358, 1], [182, 3], [276, 1], [277, 1], [351, 1], [352, 1], [348, 1], [346, 1], [350, 5, 1], [354, 1, 1]], [[357, 36, 51, 0, 0, 4], [356, 37, 35, 0, 0, 4], [355, 34, 35, 0, 0, 4]], [[209]]],
            [140, 1, 0, 0, [[269, 1], [263, 0, 2, 2, 4]], [[268, 139, 1, 18, 0, 4], [327, 139, 1, 2, 2, 4]], 0, 139],
            [141, 1, 0, 0, [[263, 0, 2, 0, 4], [417, 1]], [[268, 142, 1, 18, 0, 4], [270, 142, 1, 2, 0, 4], [269, 142, 2, 8], [418, 142]], 0, 142],
            [142, 0, [[182, 1], [187, 3], [330, 2]], 0, [[209]], [[419, 142]]],
            [143, 1, 0, 0, [[263, 0, 2, 0, 4], [417, 1]], [[268, 144, 1, 18, 0, 4], [270, 144, 1, 2, 0, 4], [269, 144, 2, 8], [418, 144]], 0, 144],
            [144, 0, [[182, 1], [187, 3], [330, 2]], 0, [[209]], [[419, 144]]],
            [145, 1, 0, 0, [[263, 0, 2, 0, 4], [417, 1]], [[268, 146, 1, 18, 0, 4], [270, 146, 1, 2, 0, 4], [269, 146, 2, 8], [418, 146]], 0, 146],
            [146, 0, [[182, 1], [187, 3], [330, 2]], 0, [[209]], [[419, 146]]],
            [147, 1, 0, 0, [[263, 0, 2, 0, 4], [417, 1]], [[268, 148, 1, 18, 0, 4], [270, 148, 1, 2, 0, 4], [269, 148, 2, 8], [418, 148]], 0, 148],
            [148, 0, [[182, 1], [187, 3], [330, 2]], 0, [[209]], [[419, 148]]],
            [149, 0, 0, 0, [["delay", 1], ["triggerMessage", 4, 0, 5], [420, 0, 0, 5], [421, 0, 0, 5], ["triggerTestEvent", 2, 0, 5], ["triggerPostProcess", 0, 0, 5], [422, 0, 0, 5], [423, 0, 0, 5], ["triggerTestEventWithFilter", 3, 0, 5], ["firstPartyMethod", 0, 2, 5, 5], [424, 0, 0, 12], [425, 0, 0, 12], [426, 0, 0, 12], [427, 0, 0, 12], ["unregisterAllCustomFunctionExecutionEvents", 0, 0, 12], ["saveWorkbookToTempFile", 1, 2, 14], ["compareTempFilesAreIdentical", 2, 2, 14], ["updateRangeValueOnCurrentSheet", 2, 0, 14, 2], ["triggerUserUndo", 0, 2, 14], ["triggerUserRedo", 0, 2, 14], ["enterCellEdit", 1, 0, 2], ["installCustomFunctionsFromCache", 0, 0, 2], ["recalc", 2, 0, 2], ["recalcBySolutionId", 1, 0, 2], ["safeForCellEditModeMethod", 1, 0, 2], ["exitCellEdit", 0, 0, 2], ["noPermissionMethod", 1, 0, 2], ["verifyCustomFunctionListExist", 0, 0, 2]], 0, 0, 0, [["CustomFunctionExecutionBeginEvent", 0, 12, "MessageType.customFunctionExecutionBeginEvent", 201, 424, 425], ["CustomFunctionExecutionEndEvent", 0, 12, "MessageType.customFunctionExecutionEndEvent", 201, 426, 427], ["Test1Event", 2, 5, "MessageType.test1Event", 201, 422, 423], ["TestEvent", 2, 5, "MessageType.testEvent", 201, 420, 421]], "Microsoft.ExcelServices.InternalTest", 4],
            [150, 0, [[428, 1], ["paperSize", 1], ["blackAndWhite", 1], ["printErrors", 1], ["zoom", 1], ["centerHorizontally", 1], ["centerVertically", 1], ["printHeadings", 1], ["printGridlines", 1], [429, 1], [430, 1], [431, 1], [432, 1], ["headerMargin", 1], ["footerMargin", 1], ["printComments", 1], ["draftMode", 1], ["firstPageNumber", 1], ["printOrder", 1]], [["headersFooters", 152, 35, 0, 0, 4]], [["setPrintMargins", 2], ["setPrintArea", 1], ["setPrintTitleRows", 1], ["setPrintTitleColumns", 1]], [["getPrintArea", 13, 0, 10, 0, 4], ["getPrintAreaOrNullObject", 13, 0, 10, 0, 4], ["getPrintTitleRows", 12, 0, 10, 0, 4], ["getPrintTitleRowsOrNullObject", 12, 0, 10, 0, 4], ["getPrintTitleColumns", 12, 0, 10, 0, 4], ["getPrintTitleColumnsOrNullObject", 12, 0, 10, 0, 4]]],
            [151, 0, [["leftHeader", 1], ["centerHeader", 1], ["rightHeader", 1], ["leftFooter", 1], ["centerFooter", 1], ["rightFooter", 1]]],
            [152, 0, [["state", 1], ["useSheetMargins", 1], ["useSheetScale", 1]], [["defaultForAllPages", 151, 35, 0, 0, 4], ["firstPage", 151, 35, 0, 0, 4], ["evenPages", 151, 35, 0, 0, 4], ["oddPages", 151, 35, 0, 0, 4]]],
            [153, 0, [[330, 2], [283, 3], [282, 3]], 0, [[209]], [["getCellAfterBreak", 12, 0, 10, 0, 4]]],
            [154, 1, 0, 0, [["removePageBreaks"], [263, 0, 2, 0, 4]], [[268, 153, 1, 18, 0, 4], [269, 153, 1, 8]], 0, 153],
            [155, 0, 0, 0, [[409]], [[269, 1, 4, 8, 6]]],
            [156, 1, 0, 0, [[263, 0, 2, 0, 4]], [[327, 12, 1, 2, 0, 4]], 0, 12],
            [157, 1, 0, 0, [[263, 0, 2, 0, 4]], [[268, 158, 1, 18, 0, 4], [327, 158, 1, 2, 0, 4], [269, 158, 3, 8], ["getItemByReplyId", 158, 1, 2, 0, 4], ["getItemByCell", 158, 1, 2, 0, 4]], 0, 158],
            [158, 0, [[205, 3], [433, 1], [434, 3], [435, 3], [411, 11], [436, 1, 3], [437, 3, 15], [438, 3, 15], [439, 3, 6]], [["replies", 159, 19, 0, 0, 4]], [[209], [440, 1, 0, 15]], [[275, 12, 0, 10, 0, 4]]],
            [159, 1, 0, 0, [[263, 0, 2, 0, 4]], [[269, 160, 2, 8], [268, 160, 1, 18, 0, 4], [327, 160, 1, 2, 0, 4]], 0, 160],
            [160, 0, [[205, 3], [433, 1], [434, 3], [435, 3], [411, 11], [436, 3, 3], [437, 3, 15], [438, 3, 15], [439, 3, 6]], 0, [[209], [440, 1, 0, 15]], [[275, 12, 0, 10, 0, 4], ["getParentComment", 158]]],
            [161, 1, 0, 0, [[263, 0, 2, 0, 4]], [["addImage", 162, 1], [369, 162, 1, 18, 0, 4], ["addGeometricShape", 162, 1], ["addTextBox", 162, 1], ["addSvg", 162, 1, 0, 6], ["addGroup", 162, 1], [327, 162, 1, 2, 0, 4], [268, 162, 1, 2, 0, 4], ["addLine", 162, 5]], 0, 162],
            [162, 0, [[205, 3], [182, 1], [297, 1], [296, 1], [299, 1], [298, 1], ["rotation", 1], ["zOrderPosition", 3], ["altTextTitle", 1], ["altTextDescription", 1], [331, 3], ["lockAspectRatio", 1], ["placement", 1, 11], ["geometricShapeType", 1], [332, 1], [395, 3], ["connectionSiteCount", 3]], [["geometricShape", 163, 2, 0, 0, 4], ["image", 164, 2, 0, 0, 4], ["textFrame", 170, 2, 0, 0, 4], [355, 168, 35, 0, 0, 4], [314, 165, 2, 0, 0, 4], ["parentGroup", 162, 2, 0, 0, 4], [383, 167, 2, 0, 0, 4], ["lineFormat", 169, 35, 0, 0, 4]], [["setZOrder", 1], ["incrementLeft", 1], ["incrementTop", 1], ["incrementRotation", 1], ["scaleHeight", 3], ["scaleWidth", 3], [209], ["getAsImage", 1], [213], [214], [215], [216]], [["copyTo", 162, 1, 8, 11, 0, 0, 441], [441, 162, 1, 2, 0, 4]], 0, 0, [[241, 0, 0, "MessageType.shapeActivatedEvent", 243, 213, 214], [250, 0, 0, "MessageType.shapeDeactivatedEvent", 243, 215, 216]]],
            [163, 0, [[205, 3]], [[442, 162, 2, 0, 0, 4]]],
            [164, 0, [[205, 3], [300, 3, 0, 300]], [[442, 162, 2, 0, 0, 4]]],
            [165, 0, [[205, 3]], [[207, 166, 19, 0, 0, 4], [442, 162, 2, 0, 0, 4]], [[315]]],
            [166, 1, 0, 0, [[263, 0, 2, 0, 4]], [[369, 162, 1, 18, 0, 4], [327, 162, 1, 2, 0, 4], [268, 162, 1, 2, 0, 4]], 0, 162],
            [167, 0, [[205, 3], [443, 1, 0, 443], ["beginArrowheadLength", 1], ["beginArrowheadStyle", 1], ["beginArrowheadWidth", 1], ["endArrowheadLength", 1], ["endArrowheadStyle", 1], ["endArrowheadWidth", 1], ["isBeginConnected", 3], ["beginConnectedSite", 3], ["isEndConnected", 3], ["endConnectedSite", 3]], [[442, 162, 2, 0, 0, 4], ["beginConnectedShape", 162, 2, 0, 0, 4], ["endConnectedShape", 162, 2, 0, 0, 4]], [["connectBeginShape", 2], ["disconnectBeginShape"], ["connectEndShape", 2], ["disconnectEndShape"]]],
            [168, 0, [["foregroundColor", 1], [331, 3], [444, 1]], 0, [[305], [397, 1]]],
            [169, 0, [[332, 1], [360, 1], [295, 1], [363, 1], ["dashStyle", 1], [444, 1]]],
            [170, 0, [[429, 1], [430, 1], [431, 1], [432, 1], [347, 1], ["horizontalOverflow", 1], [348, 1], ["verticalOverflow", 1], [428, 1], [351, 1], ["hasText", 3], ["autoSizeSetting", 1]], [["textRange", 171, 2, 0, 0, 4]], [["deleteText"]]],
            [171, 0, [[279, 1]], [[356, 172, 35, 0, 0, 4]], 0, [[396, 171, 2]]],
            [172, 0, [[364, 1], [182, 1], [360, 1], [366, 1], [365, 1], [367, 1]]],
            [173, 0, [[205, 3], [182, 1], ["caption", 1], [297, 1], [296, 1], [299, 1], [298, 1], ["nameInFormula", 1, 3], ["isFilterCleared", 3], [295, 1], ["sortBy", 1], ["sortUsingCustomLists", 1, 14], [285, 1, 14], ["disableMoveResizeUI", 1, 14], ["displayHeader", 1, 14], [349, 1, 14]], [["slicerItems", 176, 19, 0, 0, 4], [301, 8, 35, 0, 0, 4], ["slicerStyle", 146, 35, 6, 0, 4]], [[209], [340], ["getSelectedItems"], ["selectItems", 1], [210, 0, 2, 14], [341, 1, 0, 6]]],
            [174, 1, 0, 0, [[263, 0, 2, 0, 4]], [[269, 173, 3, 8], [327, 173, 1, 2, 0, 4], [268, 173, 1, 18, 0, 4], [270, 173, 1, 2, 0, 4]], 0, 173],
            [175, 0, [[328, 3], [182, 3], ["isSelected", 1], ["hasData", 3]]],
            [176, 1, 0, 0, [[263, 0, 2, 0, 4]], [[327, 175, 1, 2, 0, 4], [268, 175, 1, 18, 0, 4], [270, 175, 1, 2, 0, 4]], 0, 175],
            [177, 0, [["activeTab", 1]], 0, [["executeCommand", 2], [313, 3], [445], [446]], 0, 0, 0, [["CommandExecuted", 0, 0, "MessageType.ribbonCommandExecutedEvent", 199, 445, 446]]],
            [178, 0, [[182]], 0, [[210], [209]], [[419, 178, 1]]],
            [179, 1, 0, 0, [[263, 0, 2, 0, 4], ["exit"]], [[269, 178, 1, 8], ["enterTemporary", 178], ["getActive", 178], [268, 178, 1, 18, 0, 4], [327, 178, 1, 2, 0, 4]], 0, 178],
            [180, 0, [["error", 3], [329, 3]]],
            [181, 0, 0, 0, 0, [[335, 180, 1, 72], ["if", 180, 3, 8], ["isNA", 180, 1, 8], ["isError", 180, 1, 8], ["sum", 180, 1, 72], ["average", 180, 1, 72], [402, 180, 1, 72], [403, 180, 1, 72], ["na", 180, 0, 8], ["npv", 180, 2, 72], ["dollar", 180, 2, 8], ["fixed", 180, 3, 8], ["sin", 180, 1, 8], ["cos", 180, 1, 8], ["tan", 180, 1, 8], ["atan", 180, 1, 8], ["pi", 180, 0, 8], ["sqrt", 180, 1, 8], ["exp", 180, 1, 8], ["ln", 180, 1, 8], ["log10", 180, 1, 8], ["abs", 180, 1, 8], ["int", 180, 1, 8], ["sign", 180, 1, 8], ["round", 180, 2, 8], ["lookup", 180, 3, 8], ["rept", 180, 2, 8], ["mid", 180, 3, 8], ["len", 180, 1, 8], [329, 180, 1, 8], ["true", 180, 0, 8], ["false", 180, 0, 8], ["and", 180, 1, 72], ["or", 180, 1, 72], ["not", 180, 1, 8], ["mod", 180, 2, 8], ["dcount", 180, 3, 8, 0, 0, "DCount"], ["dsum", 180, 3, 8, 0, 0, "DSum"], ["daverage", 180, 3, 8, 0, 0, "DAverage"], ["dmin", 180, 3, 8, 0, 0, "DMin"], ["dmax", 180, 3, 8, 0, 0, "DMax"], ["dstDev", 180, 3, 8, 0, 0, "DStDev"], ["dvar", 180, 3, 8, 0, 0, "DVar"], [279, 180, 2, 8], ["pv", 180, 5, 8], ["fv", 180, 5, 8], ["nper", 180, 5, 8, 0, 0, "NPer"], ["pmt", 180, 5, 8], ["rate", 180, 6, 8], ["mirr", 180, 3, 8, 0, 0, "MIrr"], ["irr", 180, 2, 8], ["rand", 180, 0, 8], ["match", 180, 3, 8], ["date", 180, 3, 8], ["time", 180, 3, 8], ["day", 180, 1, 8], ["month", 180, 1, 8], ["year", 180, 1, 8], ["weekday", 180, 2, 8], ["hour", 180, 1, 8], ["minute", 180, 1, 8], ["second", 180, 1, 8], ["now", 180, 0, 8], [324, 180, 1, 8], [326, 180, 1, 8], [339, 180, 1, 8], [331, 180, 1, 8], ["atan2", 180, 2, 8], ["asin", 180, 1, 8], ["acos", 180, 1, 8], ["choose", 180, 2, 72], ["hlookup", 180, 4, 8, 0, 0, "HLookup"], ["vlookup", 180, 4, 8, 0, 0, "VLookup"], ["isref", 180, 1, 8], ["log", 180, 2, 8], ["char", 180, 1, 8], ["lower", 180, 1, 8], ["upper", 180, 1, 8], ["proper", 180, 1, 8], [297, 180, 2, 8], [415, 180, 2, 8], ["exact", 180, 2, 8], ["trim", 180, 1, 8], ["replace", 180, 4, 8], ["substitute", 180, 4, 8], ["code", 180, 1, 8], [321, 180, 3, 8], ["isErr", 180, 1, 8], [416, 180, 1, 8], ["isNumber", 180, 1, 8], ["t", 180, 1, 8, 0, 0, "T"], ["n", 180, 1, 8, 0, 0, "N"], ["datevalue", 180, 1, 8], ["timevalue", 180, 1, 8], ["sln", 180, 3, 8], ["syd", 180, 4, 8], ["ddb", 180, 5, 8], ["clean", 180, 1, 8], ["ipmt", 180, 6, 8], ["ppmt", 180, 6, 8], ["countA", 180, 1, 72], ["product", 180, 1, 72], ["fact", 180, 1, 8], ["dproduct", 180, 3, 8, 0, 0, "DProduct"], ["isNonText", 180, 1, 8], ["dstDevP", 180, 3, 8, 0, 0, "DStDevP"], ["dvarP", 180, 3, 8, 0, 0, "DVarP"], ["trunc", 180, 2, 8], ["isLogical", 180, 1, 8], ["dcountA", 180, 3, 8, 0, 0, "DCountA"], ["usdollar", 180, 2, 8, 0, 0, "USDollar"], ["findB", 180, 3, 8], ["replaceB", 180, 4, 8], ["leftb", 180, 2, 8], ["rightb", 180, 2, 8], ["midb", 180, 3, 8], ["lenb", 180, 1, 8], ["roundUp", 180, 2, 8], ["roundDown", 180, 2, 8], ["asc", 180, 1, 8], ["dbcs", 180, 1, 8], ["days360", 180, 3, 8], ["today", 180, 0, 8], ["vdb", 180, 7, 8], ["median", 180, 1, 72], ["sinh", 180, 1, 8], ["cosh", 180, 1, 8], ["tanh", 180, 1, 8], ["asinh", 180, 1, 8], ["acosh", 180, 1, 8], ["atanh", 180, 1, 8], ["dget", 180, 3, 8, 0, 0, "DGet"], ["db", 180, 5, 8], ["error_Type", 180, 1, 8], ["aveDev", 180, 1, 72], ["gammaLn", 180, 1, 8], ["combin", 180, 2, 8], ["even", 180, 1, 8], ["fisher", 180, 1, 8], ["fisherInv", 180, 1, 8], ["standardize", 180, 3, 8], ["odd", 180, 1, 8], ["permut", 180, 2, 8], ["devSq", 180, 1, 72], ["geoMean", 180, 1, 72], ["harMean", 180, 1, 72], ["sumSq", 180, 1, 72], ["kurt", 180, 1, 72], ["skew", 180, 1, 72], ["large", 180, 2, 8], ["small", 180, 2, 8], ["trimMean", 180, 2, 8], ["concatenate", 180, 1, 72], ["power", 180, 2, 8], ["radians", 180, 1, 8], ["degrees", 180, 1, 8], ["subtotal", 180, 2, 72], ["sumIf", 180, 3, 8], ["countIf", 180, 2, 8], ["countBlank", 180, 1, 8], ["ispmt", 180, 4, 8], ["roman", 180, 2, 8], [294, 180, 2, 8], ["averageA", 180, 1, 72], ["maxA", 180, 1, 72], ["minA", 180, 1, 72], ["stDevPA", 180, 1, 72], ["varPA", 180, 1, 72], ["stDevA", 180, 1, 72], ["varA", 180, 1, 72], ["bahtText", 180, 1, 8], ["hex2Bin", 180, 2, 8], ["hex2Dec", 180, 1, 8], ["hex2Oct", 180, 2, 8], ["dec2Bin", 180, 2, 8], ["dec2Hex", 180, 2, 8], ["dec2Oct", 180, 2, 8], ["oct2Bin", 180, 2, 8], ["oct2Hex", 180, 2, 8], ["oct2Dec", 180, 1, 8], ["bin2Dec", 180, 1, 8], ["bin2Oct", 180, 2, 8], ["bin2Hex", 180, 2, 8], ["imSub", 180, 2, 8], ["imDiv", 180, 2, 8], ["imPower", 180, 2, 8], ["imAbs", 180, 1, 8], ["imSqrt", 180, 1, 8], ["imLn", 180, 1, 8], ["imLog2", 180, 1, 8], ["imLog10", 180, 1, 8], ["imSin", 180, 1, 8], ["imCos", 180, 1, 8], ["imExp", 180, 1, 8], ["imArgument", 180, 1, 8], ["imConjugate", 180, 1, 8], ["imaginary", 180, 1, 8], ["imReal", 180, 1, 8], ["complex", 180, 3, 8], ["imSum", 180, 1, 72], ["imProduct", 180, 1, 72], ["seriesSum", 180, 4, 8], ["factDouble", 180, 1, 8], ["sqrtPi", 180, 1, 8], ["quotient", 180, 2, 8], ["delta", 180, 2, 8], ["geStep", 180, 2, 8], ["isEven", 180, 1, 8], ["isOdd", 180, 1, 8], ["mround", 180, 2, 8, 0, 0, "MRound"], ["erf", 180, 2, 8], ["erfC", 180, 1, 8], ["besselJ", 180, 2, 8], ["besselK", 180, 2, 8], ["besselY", 180, 2, 8], ["besselI", 180, 2, 8], ["xirr", 180, 3, 8], ["xnpv", 180, 3, 8], ["priceMat", 180, 6, 8], ["yieldMat", 180, 6, 8], ["intRate", 180, 5, 8], ["received", 180, 5, 8], ["disc", 180, 5, 8], ["priceDisc", 180, 5, 8], ["yieldDisc", 180, 5, 8], ["tbillEq", 180, 3, 8, 0, 0, "TBillEq"], ["tbillPrice", 180, 3, 8, 0, 0, "TBillPrice"], ["tbillYield", 180, 3, 8, 0, 0, "TBillYield"], ["price", 180, 7, 8], ["yield", 180, 7, 8], ["dollarDe", 180, 2, 8], ["dollarFr", 180, 2, 8], ["nominal", 180, 2, 8], ["effect", 180, 2, 8], ["cumPrinc", 180, 6, 8], ["cumIPmt", 180, 6, 8], ["edate", 180, 2, 8, 0, 0, "EDate"], ["eoMonth", 180, 2, 8], ["yearFrac", 180, 3, 8], ["coupDayBs", 180, 4, 8], ["coupDays", 180, 4, 8], ["coupDaysNc", 180, 4, 8], ["coupNcd", 180, 4, 8], ["coupNum", 180, 4, 8], ["coupPcd", 180, 4, 8], ["duration", 180, 6, 8], ["mduration", 180, 6, 8, 0, 0, "MDuration"], ["oddLPrice", 180, 8, 8], ["oddLYield", 180, 8, 8], ["oddFPrice", 180, 9, 8], ["oddFYield", 180, 9, 8], ["randBetween", 180, 2, 8], ["weekNum", 180, 2, 8], ["amorDegrc", 180, 7, 8], ["amorLinc", 180, 7, 8], ["convert", 180, 3, 8], ["accrInt", 180, 8, 8], ["accrIntM", 180, 5, 8], ["workDay", 180, 3, 8], ["networkDays", 180, 3, 8], ["gcd", 180, 1, 72], ["multiNomial", 180, 1, 72], ["lcm", 180, 1, 72], ["fvschedule", 180, 2, 8, 0, 0, "FVSchedule"], ["countIfs", 180, 1, 72], ["sumIfs", 180, 2, 72], ["averageIf", 180, 3, 8], ["averageIfs", 180, 2, 72], ["binom_Dist", 180, 4, 8], ["binom_Inv", 180, 3, 8], ["confidence_Norm", 180, 3, 8], ["confidence_T", 180, 3, 8], ["expon_Dist", 180, 3, 8], ["gamma_Dist", 180, 4, 8], ["gamma_Inv", 180, 3, 8], ["norm_Dist", 180, 4, 8], ["norm_Inv", 180, 3, 8], ["percentile_Exc", 180, 2, 8], ["percentile_Inc", 180, 2, 8], ["percentRank_Exc", 180, 3, 8], ["percentRank_Inc", 180, 3, 8], ["poisson_Dist", 180, 3, 8], ["quartile_Exc", 180, 2, 8], ["quartile_Inc", 180, 2, 8], ["rank_Avg", 180, 3, 8], ["rank_Eq", 180, 3, 8], ["stDev_S", 180, 1, 72], ["stDev_P", 180, 1, 72], ["t_Dist", 180, 3, 8], ["t_Dist_2T", 180, 2, 8], ["t_Dist_RT", 180, 2, 8], ["t_Inv", 180, 2, 8], ["t_Inv_2T", 180, 2, 8], ["var_S", 180, 1, 72], ["var_P", 180, 1, 72], ["weibull_Dist", 180, 4, 8], ["networkDays_Intl", 180, 4, 8], ["workDay_Intl", 180, 4, 8], ["ecma_Ceiling", 180, 2, 8, 0, 0, "ECMA_Ceiling"], ["iso_Ceiling", 180, 2, 8, 0, 0, "ISO_Ceiling"], ["beta_Dist", 180, 6, 8], ["beta_Inv", 180, 5, 8], ["chiSq_Dist", 180, 3, 8], ["chiSq_Dist_RT", 180, 2, 8], ["chiSq_Inv", 180, 2, 8], ["chiSq_Inv_RT", 180, 2, 8], ["f_Dist", 180, 4, 8], ["f_Dist_RT", 180, 3, 8], ["f_Inv", 180, 3, 8], ["f_Inv_RT", 180, 3, 8], ["hypGeom_Dist", 180, 5, 8], ["logNorm_Dist", 180, 4, 8], ["logNorm_Inv", 180, 3, 8], ["negBinom_Dist", 180, 4, 8], ["norm_S_Dist", 180, 2, 8], ["norm_S_Inv", 180, 1, 8], ["z_Test", 180, 3, 8], ["erf_Precise", 180, 1, 8], ["erfC_Precise", 180, 1, 8], ["gammaLn_Precise", 180, 1, 8], ["ceiling_Precise", 180, 2, 8], ["floor_Precise", 180, 2, 8], ["acot", 180, 1, 8], ["acoth", 180, 1, 8], ["cot", 180, 1, 8], ["coth", 180, 1, 8], ["csc", 180, 1, 8], ["csch", 180, 1, 8], ["sec", 180, 1, 8], ["sech", 180, 1, 8], ["imTan", 180, 1, 8], ["imCot", 180, 1, 8], ["imCsc", 180, 1, 8], ["imCsch", 180, 1, 8], ["imSec", 180, 1, 8], ["imSech", 180, 1, 8], ["bitand", 180, 2, 8], ["bitor", 180, 2, 8], ["bitxor", 180, 2, 8], ["bitlshift", 180, 2, 8], ["bitrshift", 180, 2, 8], ["permutationa", 180, 2, 8], ["combina", 180, 2, 8], ["xor", 180, 1, 72], ["pduration", 180, 3, 8, 0, 0, "PDuration"], ["base", 180, 3, 8], ["decimal", 180, 2, 8], ["days", 180, 2, 8], ["binom_Dist_Range", 180, 4, 8], ["gamma", 180, 1, 8], ["skew_p", 180, 1, 72], ["gauss", 180, 1, 8], ["phi", 180, 1, 8], ["rri", 180, 3, 8], ["unichar", 180, 1, 8], ["unicode", 180, 1, 8], ["arabic", 180, 1, 8], ["isoWeekNum", 180, 1, 8], ["numberValue", 180, 3, 8], ["sheet", 180, 1, 8], ["sheets", 180, 1, 8], ["isFormula", 180, 1, 8], ["ceiling_Math", 180, 3, 8], ["floor_Math", 180, 3, 8], ["imSinh", 180, 1, 8], ["imCosh", 180, 1, 8]]]] };
    var _builder = new OfficeExtension.LibraryBuilder({ metadata: _libraryMetadataXlapi, targetNamespaceObject: Excel });
})(Excel || (Excel = {}));
var _EndExcel = "_EndExcel";
OSFPerformance.officeExecuteEnd = OSFPerformance.now();


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
var OSFPerformance;
(function (OSFPerformance) {
    function now() {
        if (typeof (performance) !== 'undefined' && performance.now) {
            return performance.now();
        }
        else {
            return 0;
        }
    }
    OSFPerformance.now = now;
    OSFPerformance.officeExecuteStartDate = Date.now();
    OSFPerformance.officeExecuteStart = 0;
    OSFPerformance.officeExecuteEnd = 0;
    OSFPerformance.hostInitializationStart = now();
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
var Office;
(function (Office) {
    var actions;
    (function (actions) {
        var m_association;
        function get_association() {
            if (!m_association) {
                m_association = new OSF.Association();
            }
            return m_association;
        }
        function associate() {
            get_association().associate.apply(get_association(), arguments);
        }
        actions.associate = associate;
        ;
        OSF.definePropertyOnNamespace(actions, '_association', get_association);
    })(actions = Office.actions || (Office.actions = {}));
})(Office || (Office = {}));
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
        AgaveHostAction[AgaveHostAction["GetOriginalControlId"] = 34] = "GetOriginalControlId";
        AgaveHostAction[AgaveHostAction["OfficeJsReady"] = 35] = "OfficeJsReady";
        AgaveHostAction[AgaveHostAction["InsertDevManifest"] = 36] = "InsertDevManifest";
        AgaveHostAction[AgaveHostAction["InsertDevManifestError"] = 37] = "InsertDevManifestError";
        AgaveHostAction[AgaveHostAction["SendCustomerContent"] = 38] = "SendCustomerContent";
        AgaveHostAction[AgaveHostAction["KeyboardShortcuts"] = 39] = "KeyboardShortcuts";
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
                    id: OSF.EventDispId.dispidAppCommandInvokedEvent,
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
            OSF.EventHelper.addEventHandler(eventType, handler, callback, this._eventDispatch, undefined, OSF.OUtil.isPopupWindow());
        };
        AppCommandManager.prototype._verifyManifestCallback = function (callbackName) {
            var defaultResult = { callback: null, errorCode: 11101 };
            callbackName = callbackName.trim();
            try {
                var callbackFunc = this._getCallbackFunc(callbackName);
                if (typeof callbackFunc != "function") {
                    return defaultResult;
                }
            }
            catch (e) {
                return defaultResult;
            }
            return { callback: callbackFunc, errorCode: 0 };
        };
        AppCommandManager.prototype._getCallbackFuncFromWindow = function (callbackName) {
            var callList = callbackName.split(".");
            var parentObject = window;
            for (var i = 0; i < callList.length - 1; i++) {
                if (parentObject[callList[i]] && (typeof parentObject[callList[i]] == "object" || typeof parentObject[callList[i]] == "function")) {
                    parentObject = parentObject[callList[i]];
                }
                else {
                    return null;
                }
            }
            var callbackFunc = parentObject[callList[callList.length - 1]];
            return callbackFunc;
        };
        AppCommandManager.prototype._getCallbackFuncFromActionAssociateTable = function (callbackName) {
            var nameUpperCase = callbackName.toUpperCase();
            return Office.actions._association.mappings[nameUpperCase];
        };
        AppCommandManager.prototype._getCallbackFunc = function (callbackName) {
            var callbackFunc = this._getCallbackFuncFromWindow(callbackName);
            if (!callbackFunc) {
                callbackFunc = this._getCallbackFuncFromActionAssociateTable(callbackName);
            }
            return callbackFunc;
        };
        AppCommandManager.prototype._invokeAppCommandCompletedMethod = function (appCommandId, resultCode, data) {
            this.appCommandInvocationCompletedAsync(appCommandId, resultCode, data, function (result) {
                if (result.status !== Office.AsyncResultStatus.Succeeded) {
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
                    if (result.status !== Office.AsyncResultStatus.Succeeded) {
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
    var AppName;
    (function (AppName) {
        AppName[AppName["Unsupported"] = 0] = "Unsupported";
        AppName[AppName["Excel"] = 1] = "Excel";
        AppName[AppName["Word"] = 2] = "Word";
        AppName[AppName["PowerPoint"] = 4] = "PowerPoint";
        AppName[AppName["Outlook"] = 8] = "Outlook";
        AppName[AppName["ExcelWebApp"] = 16] = "ExcelWebApp";
        AppName[AppName["WordWebApp"] = 32] = "WordWebApp";
        AppName[AppName["OutlookWebApp"] = 64] = "OutlookWebApp";
        AppName[AppName["Project"] = 128] = "Project";
        AppName[AppName["AccessWebApp"] = 256] = "AccessWebApp";
        AppName[AppName["PowerpointWebApp"] = 512] = "PowerpointWebApp";
        AppName[AppName["ExcelIOS"] = 1024] = "ExcelIOS";
        AppName[AppName["Sway"] = 2048] = "Sway";
        AppName[AppName["WordIOS"] = 4096] = "WordIOS";
        AppName[AppName["PowerPointIOS"] = 8192] = "PowerPointIOS";
        AppName[AppName["Access"] = 16384] = "Access";
        AppName[AppName["Lync"] = 32768] = "Lync";
        AppName[AppName["OutlookIOS"] = 65536] = "OutlookIOS";
        AppName[AppName["OneNoteWebApp"] = 131072] = "OneNoteWebApp";
        AppName[AppName["OneNote"] = 262144] = "OneNote";
        AppName[AppName["ExcelWinRT"] = 524288] = "ExcelWinRT";
        AppName[AppName["WordWinRT"] = 1048576] = "WordWinRT";
        AppName[AppName["PowerpointWinRT"] = 2097152] = "PowerpointWinRT";
        AppName[AppName["OutlookAndroid"] = 4194304] = "OutlookAndroid";
        AppName[AppName["OneNoteWinRT"] = 8388608] = "OneNoteWinRT";
        AppName[AppName["ExcelAndroid"] = 8388609] = "ExcelAndroid";
        AppName[AppName["VisioWebApp"] = 8388610] = "VisioWebApp";
        AppName[AppName["OneNoteIOS"] = 8388611] = "OneNoteIOS";
        AppName[AppName["WordAndroid"] = 8388613] = "WordAndroid";
        AppName[AppName["PowerpointAndroid"] = 8388614] = "PowerpointAndroid";
        AppName[AppName["Visio"] = 8388615] = "Visio";
        AppName[AppName["OneNoteAndroid"] = 4194305] = "OneNoteAndroid";
    })(AppName = OSF.AppName || (OSF.AppName = {}));
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
        AsyncMethodExecutor.prototype.invokeCallback = function (dispId, callback, status, value, asyncContext) {
            if (status == 0) {
                var successResult = {
                    status: Office.AsyncResultStatus.Succeeded,
                    value: value,
                    asyncContext: asyncContext
                };
                if (typeof callback == "function") {
                    callback(successResult);
                }
            }
            else {
                var errorResult = {
                    status: Office.AsyncResultStatus.Failed,
                    error: {
                        code: status
                    },
                    asyncContext: asyncContext
                };
                if (typeof callback == "function") {
                    callback(errorResult);
                }
            }
        };
        return AsyncMethodExecutor;
    }());
    OSF.AsyncMethodExecutor = AsyncMethodExecutor;
})(OSF || (OSF = {}));
var OSF;
(function (OSF) {
    var AsyncMethodExecutorHelper = (function () {
        function AsyncMethodExecutorHelper(asyncMethodExecutor) {
            this._asyncMethodExecutor = asyncMethodExecutor;
        }
        AsyncMethodExecutorHelper.prototype.handleSafeArrayHostResponse = function (hostResponseArgs, resultCode, chunkResultData, callback, dataTransform, id, asyncContext) {
            var result;
            var status;
            var hostResponseArgs = OSF.Utility.fromSafeArray(hostResponseArgs);
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
                    if (chunkResultData.length > 0) {
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
                this._asyncMethodExecutor.invokeCallback(id, callback, status, value, asyncContext);
            }
            return true;
        };
        AsyncMethodExecutorHelper.prototype.handleWebHostResponse = function (hostResponseArgs, resultCode, callback, dataTransform, id, asyncContext) {
            var value = null;
            if (resultCode == 0) {
                value = dataTransform.fromWebHost(hostResponseArgs);
            }
            this._asyncMethodExecutor.invokeCallback(id, callback, resultCode, value, asyncContext);
        };
        return AsyncMethodExecutorHelper;
    }());
    OSF.AsyncMethodExecutorHelper = AsyncMethodExecutorHelper;
})(OSF || (OSF = {}));
var OSF;
(function (OSF) {
    var DDA;
    (function (DDA) {
        var AsyncResultEnum;
        (function (AsyncResultEnum) {
            var Properties;
            (function (Properties) {
                Properties["Context"] = "Context";
                Properties["Value"] = "Value";
                Properties["Status"] = "Status";
                Properties["Error"] = "Error";
            })(Properties = AsyncResultEnum.Properties || (AsyncResultEnum.Properties = {}));
            ;
            var ErrorCode;
            (function (ErrorCode) {
                ErrorCode[ErrorCode["Success"] = 0] = "Success";
                ErrorCode[ErrorCode["Failed"] = 1] = "Failed";
            })(ErrorCode = AsyncResultEnum.ErrorCode || (AsyncResultEnum.ErrorCode = {}));
            ;
            var ErrorProperties;
            (function (ErrorProperties) {
                ErrorProperties["Name"] = "Name";
                ErrorProperties["Message"] = "Message";
                ErrorProperties["Code"] = "Code";
            })(ErrorProperties = AsyncResultEnum.ErrorProperties || (AsyncResultEnum.ErrorProperties = {}));
            ;
        })(AsyncResultEnum = DDA.AsyncResultEnum || (DDA.AsyncResultEnum = {}));
        var AsyncResult = (function () {
            function AsyncResult(initArgs, errorArgs) {
                this.value = initArgs.Value;
                this.status = (errorArgs ? Office.AsyncResultStatus.Failed : Office.AsyncResultStatus.Succeeded);
                if (initArgs.Context) {
                    this.asyncContext = initArgs.Context;
                }
                if (errorArgs) {
                    this.error = new Error(errorArgs.Name, errorArgs.Message, errorArgs.Code);
                }
            }
            return AsyncResult;
        }());
        DDA.AsyncResult = AsyncResult;
        var Error = (function () {
            function Error(name, message, code) {
                this.name = name;
                this.message = message;
                this.code = code;
            }
            return Error;
        }());
        DDA.Error = Error;
    })(DDA = OSF.DDA || (OSF.DDA = {}));
})(OSF || (OSF = {}));
var OSF;
(function (OSF) {
    var BootStrapExtension;
    (function (BootStrapExtension) {
        BootStrapExtension.createWebClientHostControllerHelper = function (webAppState, delegateVersion) {
            return new OSF.WebClientHostControllerHelper(webAppState, delegateVersion);
        };
        BootStrapExtension.createAsyncMethodExecutorHelper = function (asyncMethodExecutor) {
            return new OSF.AsyncMethodExecutorHelper(asyncMethodExecutor);
        };
    })(BootStrapExtension = OSF.BootStrapExtension || (OSF.BootStrapExtension = {}));
})(OSF || (OSF = {}));
var OSF;
(function (OSF) {
    var ConstantNames;
    (function (ConstantNames) {
        ConstantNames["DefaultLocale"] = "en-us";
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
var CustomFunctionMappings;
(function (CustomFunctionMappings) {
})(CustomFunctionMappings || (CustomFunctionMappings = {}));
var CustomFunctions;
(function (CustomFunctions) {
    function delayInitialization() {
        CustomFunctionMappings.__delay__ = true;
    }
    CustomFunctions.delayInitialization = delayInitialization;
    ;
    CustomFunctions._association = new OSF.Association();
    function associate() {
        CustomFunctions._association.associate.apply(CustomFunctions._association, arguments);
        delete CustomFunctionMappings.__delay__;
    }
    CustomFunctions.associate = associate;
    ;
})(CustomFunctions || (CustomFunctions = {}));
var OSF;
(function (OSF) {
    var DDA;
    (function (DDA) {
        var ErrorCodeManager;
        (function (ErrorCodeManager) {
            var _errorMappings = {};
            var _isErrorMessagesInitializedFromOfficeString = false;
            function getErrorArgs(errorCode) {
                if (!_isErrorMessagesInitializedFromOfficeString) {
                    initializeErrorMessages(Strings.OfficeOM);
                }
                var errorArgs = _errorMappings[errorCode];
                if (!errorArgs) {
                    errorArgs = _errorMappings[5001];
                }
                else {
                    if (!errorArgs.name) {
                        errorArgs.name = _errorMappings[5001].name;
                    }
                    if (!errorArgs.message) {
                        errorArgs.message = _errorMappings[5001].message;
                    }
                }
                return errorArgs;
            }
            ErrorCodeManager.getErrorArgs = getErrorArgs;
            function addErrorMessage(errorCode, errorNameMessage) {
                _errorMappings[errorCode] = errorNameMessage;
            }
            ErrorCodeManager.addErrorMessage = addErrorMessage;
            function initializeErrorMessages(stringNS) {
                _errorMappings[1000] = { name: stringNS.L_InvalidCoercion, message: stringNS.L_CoercionTypeNotSupported };
                _errorMappings[1001] = { name: stringNS.L_DataReadError, message: stringNS.L_GetSelectionNotSupported };
                _errorMappings[1002] = { name: stringNS.L_InvalidCoercion, message: stringNS.L_CoercionTypeNotMatchBinding };
                _errorMappings[1003] = { name: stringNS.L_DataReadError, message: stringNS.L_InvalidGetRowColumnCounts };
                _errorMappings[1004] = { name: stringNS.L_DataReadError, message: stringNS.L_SelectionNotSupportCoercionType };
                _errorMappings[1005] = { name: stringNS.L_DataReadError, message: stringNS.L_InvalidGetStartRowColumn };
                _errorMappings[1006] = { name: stringNS.L_DataReadError, message: stringNS.L_NonUniformPartialGetNotSupported };
                _errorMappings[1008] = { name: stringNS.L_DataReadError, message: stringNS.L_GetDataIsTooLarge };
                _errorMappings[1009] = { name: stringNS.L_DataReadError, message: stringNS.L_FileTypeNotSupported };
                _errorMappings[1010] = { name: stringNS.L_DataReadError, message: stringNS.L_GetDataParametersConflict };
                _errorMappings[1011] = { name: stringNS.L_DataReadError, message: stringNS.L_InvalidGetColumns };
                _errorMappings[1012] = { name: stringNS.L_DataReadError, message: stringNS.L_InvalidGetRows };
                _errorMappings[1013] = { name: stringNS.L_DataReadError, message: stringNS.L_InvalidReadForBlankRow };
                _errorMappings[2000] = { name: stringNS.L_DataWriteError, message: stringNS.L_UnsupportedDataObject };
                _errorMappings[2001] = { name: stringNS.L_DataWriteError, message: stringNS.L_CannotWriteToSelection };
                _errorMappings[2002] = { name: stringNS.L_DataWriteError, message: stringNS.L_DataNotMatchSelection };
                _errorMappings[2003] = { name: stringNS.L_DataWriteError, message: stringNS.L_OverwriteWorksheetData };
                _errorMappings[2004] = { name: stringNS.L_DataWriteError, message: stringNS.L_DataNotMatchBindingSize };
                _errorMappings[2005] = { name: stringNS.L_DataWriteError, message: stringNS.L_InvalidSetStartRowColumn };
                _errorMappings[2006] = { name: stringNS.L_InvalidFormat, message: stringNS.L_InvalidDataFormat };
                _errorMappings[2007] = { name: stringNS.L_InvalidDataObject, message: stringNS.L_DataNotMatchCoercionType };
                _errorMappings[2008] = { name: stringNS.L_InvalidDataObject, message: stringNS.L_DataNotMatchBindingType };
                _errorMappings[2009] = { name: stringNS.L_DataWriteError, message: stringNS.L_SetDataIsTooLarge };
                _errorMappings[2010] = { name: stringNS.L_DataWriteError, message: stringNS.L_NonUniformPartialSetNotSupported };
                _errorMappings[2011] = { name: stringNS.L_DataWriteError, message: stringNS.L_InvalidSetColumns };
                _errorMappings[2012] = { name: stringNS.L_DataWriteError, message: stringNS.L_InvalidSetRows };
                _errorMappings[2013] = { name: stringNS.L_DataWriteError, message: stringNS.L_SetDataParametersConflict };
                _errorMappings[3000] = { name: stringNS.L_BindingCreationError, message: stringNS.L_SelectionCannotBound };
                _errorMappings[3002] = { name: stringNS.L_InvalidBindingError, message: stringNS.L_BindingNotExist };
                _errorMappings[3003] = { name: stringNS.L_BindingCreationError, message: stringNS.L_BindingToMultipleSelection };
                _errorMappings[3004] = { name: stringNS.L_BindingCreationError, message: stringNS.L_InvalidSelectionForBindingType };
                _errorMappings[3005] = { name: stringNS.L_InvalidBindingOperation, message: stringNS.L_OperationNotSupportedOnThisBindingType };
                _errorMappings[3006] = { name: stringNS.L_BindingCreationError, message: stringNS.L_NamedItemNotFound };
                _errorMappings[3007] = { name: stringNS.L_BindingCreationError, message: stringNS.L_MultipleNamedItemFound };
                _errorMappings[3008] = { name: stringNS.L_BindingCreationError, message: stringNS.L_InvalidNamedItemForBindingType };
                _errorMappings[3009] = { name: stringNS.L_InvalidBinding, message: stringNS.L_UnknownBindingType };
                _errorMappings[3010] = { name: stringNS.L_InvalidBindingOperation, message: stringNS.L_OperationNotSupportedOnMatrixData };
                _errorMappings[3011] = { name: stringNS.L_InvalidBinding, message: stringNS.L_InvalidColumnsForBinding };
                _errorMappings[4000] = { name: stringNS.L_ReadSettingsError, message: stringNS.L_SettingNameNotExist };
                _errorMappings[4001] = { name: stringNS.L_SaveSettingsError, message: stringNS.L_SettingsCannotSave };
                _errorMappings[4002] = { name: stringNS.L_SettingsStaleError, message: stringNS.L_SettingsAreStale };
                _errorMappings[5000] = { name: stringNS.L_HostError, message: stringNS.L_OperationNotSupported };
                _errorMappings[5001] = { name: stringNS.L_InternalError, message: stringNS.L_InternalErrorDescription };
                _errorMappings[5002] = { name: stringNS.L_PermissionDenied, message: stringNS.L_DocumentReadOnly };
                _errorMappings[5003] = { name: stringNS.L_EventRegistrationError, message: stringNS.L_EventHandlerNotExist };
                _errorMappings[5004] = { name: stringNS.L_InvalidAPICall, message: stringNS.L_InvalidApiCallInContext };
                _errorMappings[5005] = { name: stringNS.L_ShuttingDown, message: stringNS.L_ShuttingDown };
                _errorMappings[5007] = { name: stringNS.L_UnsupportedEnumeration, message: stringNS.L_UnsupportedEnumerationMessage };
                _errorMappings[5008] = { name: stringNS.L_IndexOutOfRange, message: stringNS.L_IndexOutOfRange };
                _errorMappings[5009] = { name: stringNS.L_APINotSupported, message: stringNS.L_BrowserAPINotSupported };
                _errorMappings[5011] = { name: stringNS.L_APICallFailed, message: stringNS.L_RequestTimeout };
                _errorMappings[5012] = { name: stringNS.L_InvalidOrTimedOutSession, message: stringNS.L_InvalidOrTimedOutSessionMessage };
                _errorMappings[5013] = { name: stringNS.L_APICallFailed, message: stringNS.L_InvalidApiArgumentsMessage };
                _errorMappings[5015] = { name: stringNS.L_APICallFailed, message: stringNS.L_WorkbookHiddenMessage };
                _errorMappings[5016] = { name: stringNS.L_APICallFailed, message: stringNS.L_WriteNotSupportedWhenModalDialogOpen };
                _errorMappings[5100] = { name: stringNS.L_APICallFailed, message: stringNS.L_TooManyIncompleteRequests };
                _errorMappings[5101] = { name: stringNS.L_APICallFailed, message: stringNS.L_RequestTokenUnavailable };
                _errorMappings[5102] = { name: stringNS.L_APICallFailed, message: stringNS.L_ActivityLimitReached };
                _errorMappings[5103] = { name: stringNS.L_APICallFailed, message: stringNS.L_RequestPayloadSizeLimitExceededMessage };
                _errorMappings[5104] = { name: stringNS.L_APICallFailed, message: stringNS.L_ResponsePayloadSizeLimitExceededMessage };
                _errorMappings[6000] = { name: stringNS.L_InvalidNode, message: stringNS.L_CustomXmlNodeNotFound };
                _errorMappings[6100] = { name: stringNS.L_CustomXmlError, message: stringNS.L_CustomXmlError };
                _errorMappings[6101] = { name: stringNS.L_CustomXmlExceedQuotaName, message: stringNS.L_CustomXmlExceedQuotaMessage };
                _errorMappings[6102] = { name: stringNS.L_CustomXmlOutOfDateName, message: stringNS.L_CustomXmlOutOfDateMessage };
                _errorMappings[7000] = { name: stringNS.L_PermissionDenied, message: stringNS.L_NoCapability };
                _errorMappings[7001] = { name: stringNS.L_CannotNavigateTo, message: stringNS.L_CannotNavigateTo };
                _errorMappings[7002] = { name: stringNS.L_SpecifiedIdNotExist, message: stringNS.L_SpecifiedIdNotExist };
                _errorMappings[7004] = { name: stringNS.L_NavOutOfBound, message: stringNS.L_NavOutOfBound };
                _errorMappings[2014] = { name: stringNS.L_DataWriteReminder, message: stringNS.L_CellDataAmountBeyondLimits };
                _errorMappings[8000] = { name: stringNS.L_MissingParameter, message: stringNS.L_ElementMissing };
                _errorMappings[8001] = { name: stringNS.L_PermissionDenied, message: stringNS.L_NoCapability };
                _errorMappings[8010] = { name: stringNS.L_InvalidValue, message: stringNS.L_InvalidCellsValue };
                _errorMappings[8011] = { name: stringNS.L_InvalidValue, message: stringNS.L_InvalidTableOptionValue };
                _errorMappings[8012] = { name: stringNS.L_InvalidValue, message: stringNS.L_InvalidFormatValue };
                _errorMappings[8020] = { name: stringNS.L_OutOfRange, message: stringNS.L_RowIndexOutOfRange };
                _errorMappings[8021] = { name: stringNS.L_OutOfRange, message: stringNS.L_ColIndexOutOfRange };
                _errorMappings[8022] = { name: stringNS.L_OutOfRange, message: stringNS.L_FormatValueOutOfRange };
                _errorMappings[8023] = { name: stringNS.L_FormattingReminder, message: stringNS.L_CellFormatAmountBeyondLimits };
                _errorMappings[10000] = { name: stringNS.L_UserNotSignedIn, message: stringNS.L_UserNotSignedIn };
                _errorMappings[11000] = { name: stringNS.L_MemoryLimit, message: stringNS.L_CloseFileBeforeRetrieve };
                _errorMappings[11001] = { name: stringNS.L_NetworkProblem, message: stringNS.L_NetworkProblemRetrieveFile };
                _errorMappings[11002] = { name: stringNS.L_InvalidValue, message: stringNS.L_SliceSizeNotSupported };
                _errorMappings[12007] = { name: stringNS.L_DisplayDialogError, message: stringNS.L_DialogAlreadyOpened };
                _errorMappings[12000] = { name: stringNS.L_IndexOutOfRange, message: stringNS.L_IndexOutOfRange };
                _errorMappings[12001] = { name: stringNS.L_IndexOutOfRange, message: stringNS.L_IndexOutOfRange };
                _errorMappings[12002] = { name: stringNS.L_DisplayDialogError, message: stringNS.L_NetworkProblem };
                _errorMappings[12003] = { name: stringNS.L_DialogNavigateError, message: stringNS.L_DialogInvalidScheme };
                _errorMappings[12004] = { name: stringNS.L_DisplayDialogError, message: stringNS.L_DialogAddressNotTrusted };
                _errorMappings[12005] = { name: stringNS.L_DisplayDialogError, message: stringNS.L_DialogRequireHTTPS };
                _errorMappings[12009] = { name: stringNS.L_DisplayDialogError, message: stringNS.L_UserClickIgnore };
                _errorMappings[12011] = { name: stringNS.L_DisplayDialogError, message: stringNS.L_NewWindowCrossZoneErrorString };
                _errorMappings[13000] = { name: stringNS.L_APINotSupported, message: stringNS.L_InvalidSSOAddinMessage };
                _errorMappings[13001] = { name: stringNS.L_UserNotSignedIn, message: stringNS.L_UserNotSignedIn };
                _errorMappings[13002] = { name: stringNS.L_UserAborted, message: stringNS.L_UserAbortedMessage };
                _errorMappings[13003] = { name: stringNS.L_UnsupportedUserIdentity, message: stringNS.L_UnsupportedUserIdentityMessage };
                _errorMappings[13004] = { name: stringNS.L_InvalidResourceUrl, message: stringNS.L_InvalidResourceUrlMessage };
                _errorMappings[13005] = { name: stringNS.L_InvalidGrant, message: stringNS.L_InvalidGrantMessage };
                _errorMappings[13006] = { name: stringNS.L_SSOClientError, message: stringNS.L_SSOClientErrorMessage };
                _errorMappings[13007] = { name: stringNS.L_SSOServerError, message: stringNS.L_SSOServerErrorMessage };
                _errorMappings[13008] = { name: stringNS.L_AddinIsAlreadyRequestingToken, message: stringNS.L_AddinIsAlreadyRequestingTokenMessage };
                _errorMappings[13009] = { name: stringNS.L_SSOUserConsentNotSupportedByCurrentAddinCategory, message: stringNS.L_SSOUserConsentNotSupportedByCurrentAddinCategoryMessage };
                _errorMappings[13010] = { name: stringNS.L_SSOConnectionLostError, message: stringNS.L_SSOConnectionLostErrorMessage };
                _errorMappings[13012] = { name: stringNS.L_APINotSupported, message: stringNS.L_SSOUnsupportedPlatform };
                _errorMappings[13013] = { name: stringNS.L_APICallFailed, message: stringNS.L_RequestTokenUnavailable };
                _errorMappings[5014] = { name: stringNS.L_OperationCancelledError, message: stringNS.L_OperationCancelledErrorMessage };
                _isErrorMessagesInitializedFromOfficeString = true;
            }
            function getAsyncResult(code) {
                if (code == 0) {
                    return {
                        status: Office.AsyncResultStatus.Succeeded
                    };
                }
                else {
                    return {
                        status: Office.AsyncResultStatus.Failed,
                        error: {
                            code: code
                        }
                    };
                }
            }
        })(ErrorCodeManager = DDA.ErrorCodeManager || (DDA.ErrorCodeManager = {}));
    })(DDA = OSF.DDA || (OSF.DDA = {}));
})(OSF || (OSF = {}));
var OSF;
(function (OSF) {
    var EventDispatch = (function () {
        function EventDispatch(args) {
            this._eventInfos = {};
            this._queuedEventsArgs = {};
            this._eventHandlers = {};
            this._queuedEventsArgs = {};
            if (args != null) {
                for (var i = 0; i < args.length; i++) {
                    if (typeof args[i] === "string") {
                        var eventType = args[i];
                        this._eventHandlers[eventType] = [];
                        this._queuedEventsArgs[eventType] = [];
                    }
                    else {
                        var eventType = args[i].type;
                        this._eventInfos[eventType] = args[i];
                        this._eventHandlers[eventType] = [];
                        this._queuedEventsArgs[eventType] = [];
                    }
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
        EventDispatch.prototype.fireOrQueueEvent = function (eventArgs) {
            if (eventArgs.type == undefined)
                return false;
            var eventType = eventArgs.type;
            if (eventType && this._eventHandlers[eventType]) {
                var eventHandlers = this._eventHandlers[eventType];
                var queuedEvents = this._queuedEventsArgs[eventType];
                if (eventHandlers.length == 0) {
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
    var EventDispId;
    (function (EventDispId) {
        EventDispId.dispidEventMin = 0;
        EventDispId.dispidInitializeEvent = 0;
        EventDispId.dispidSettingsChangedEvent = 1;
        EventDispId.dispidDocumentSelectionChangedEvent = 2;
        EventDispId.dispidBindingSelectionChangedEvent = 3;
        EventDispId.dispidBindingDataChangedEvent = 4;
        EventDispId.dispidDocumentOpenEvent = 5;
        EventDispId.dispidDocumentCloseEvent = 6;
        EventDispId.dispidActiveViewChangedEvent = 7;
        EventDispId.dispidDocumentThemeChangedEvent = 8;
        EventDispId.dispidOfficeThemeChangedEvent = 9;
        EventDispId.dispidDialogMessageReceivedEvent = 10;
        EventDispId.dispidDialogNotificationShownInAddinEvent = 11;
        EventDispId.dispidDialogParentMessageReceivedEvent = 12;
        EventDispId.dispidObjectDeletedEvent = 13;
        EventDispId.dispidObjectSelectionChangedEvent = 14;
        EventDispId.dispidObjectDataChangedEvent = 15;
        EventDispId.dispidContentControlAddedEvent = 16;
        EventDispId.dispidActivationStatusChangedEvent = 32;
        EventDispId.dispidRichApiMessageEvent = 33;
        EventDispId.dispidAppCommandInvokedEvent = 39;
        EventDispId.dispidDataNodeAddedEvent = 60;
        EventDispId.dispidDataNodeReplacedEvent = 61;
        EventDispId.dispidDataNodeDeletedEvent = 62;
    })(EventDispId = OSF.EventDispId || (OSF.EventDispId = {}));
})(OSF || (OSF = {}));
var OSF;
(function (OSF) {
    var EventHelper = (function () {
        function EventHelper() {
        }
        EventHelper.addEventHandler = function (eventType, handler, callback, eventDispatch, asyncContext, isPopupWindow) {
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
                asyncMethodExecutor.invokeCallback(dispId, callback, status, null, asyncContext);
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
                    }, asyncContext);
                }
                else {
                    onEnsureRegistration(0);
                }
            }
            catch (ex) {
                EventHelper.onException(dispId, ex, callback);
            }
        };
        EventHelper.removeEventHandler = function (eventType, handler, callback, eventDispatch, asyncContext, isPopupWindow) {
            var dispId = 0;
            function onEnsureRegistration(status) {
                var asyncMethodExecutor = OSF._OfficeAppFactory.getAsyncMethodExecutor();
                asyncMethodExecutor.invokeCallback(dispId, callback, status, null, asyncContext);
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
                    }, asyncContext);
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
        EventType.AppCommandInvoked = "appCommandInvoked";
        EventType.RichApiMessage = "richApiMessage";
        EventType.BindingSelectionChanged = "bindingSelectionChanged";
        EventType.BindingDataChanged = "bindingDataChanged";
        EventType.DataNodeDeleted = "nodeDeleted";
        EventType.DataNodeInserted = "nodeInserted";
        EventType.DataNodeReplaced = "nodeReplaced";
        EventType.SettingsChanged = "settingsChanged";
        EventType.DialogMessageReceived = "dialogMessageReceived";
        EventType.DialogParentMessageReceived = "dialogParentMessageReceived";
        EventType.DialogParentEventReceived = "dialogParentEventReceived";
        EventType.DialogEventReceived = "dialogEventReceived";
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
            if (OSF.BootStrapExtension.prepareApiSurface) {
                OSF.BootStrapExtension.prepareApiSurface();
            }
            OSFPerformance.createOMEnd = OSFPerformance.now();
        };
        InitializationHelper.prototype.getTabbableElements = function () {
            return null;
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
})(OSF || (OSF = {}));
var OSF;
(function (OSF) {
    var LoadScriptHelper;
    (function (LoadScriptHelper) {
        var _scriptInfo;
        var _officeScriptBase = ['excel', 'word', 'powerpoint', 'outlook', 'office-common', 'office.common'];
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
        function getOfficeStringJsName() {
            ensureScriptInfo();
            return _scriptInfo.isDebugJs ? OSF.ConstantNames.OfficeStringDebugJS : OSF.ConstantNames.OfficeStringJS;
        }
        LoadScriptHelper.getOfficeStringJsName = getOfficeStringJsName;
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
                    var isDebugJs = scriptSrcLowerCase.indexOf(".debug.js", indexOfJS) > 0;
                    return { basePath: scriptBase, name: scriptNameToCheck, isDebugJs: isDebugJs };
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
                name: "",
                isDebugJs: false
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
var OSF;
(function (OSF) {
    var SupportedLocales = {
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
    var AssociatedLocales = {
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
    function getSupportedLocale(locale, defaultLocale) {
        if (defaultLocale === void 0) { defaultLocale = "en-us"; }
        if (!locale) {
            return defaultLocale;
        }
        var supportedLocale;
        locale = locale.toLowerCase();
        if (locale in SupportedLocales) {
            supportedLocale = locale;
        }
        else {
            var localeParts = locale.split('-', 1);
            if (localeParts && localeParts.length > 0) {
                supportedLocale = AssociatedLocales[localeParts[0]];
            }
        }
        if (!supportedLocale) {
            supportedLocale = defaultLocale;
        }
        return supportedLocale;
    }
    OSF.getSupportedLocale = getSupportedLocale;
})(OSF || (OSF = {}));
var Strings;
(function (Strings) {
    var OfficeOM;
    (function (OfficeOM) {
    })(OfficeOM = Strings.OfficeOM || (Strings.OfficeOM = {}));
})(Strings || (Strings = {}));
var OSF;
(function (OSF) {
    var OUtil;
    (function (OUtil) {
        var officeStringsJsLoadPromise;
        function ensureOfficeStringsJs() {
            if (!officeStringsJsLoadPromise) {
                officeStringsJsLoadPromise = new Office.Promise(function (resolve, reject) {
                    if (!OSF._OfficeAppFactory.getHostInfo().hostLocale) {
                        reject(new Error("No host locale"));
                        return;
                    }
                    var url = OSF.LoadScriptHelper.getHostBundleJsBasePath() + OSF._OfficeAppFactory.getHostInfo().hostLocale + "/" + OSF.LoadScriptHelper.getOfficeStringJsName();
                    OSF.OUtil.loadScript(url, function (success) {
                        if (success) {
                            resolve();
                        }
                        else {
                            var fallbackUrl = OSF.LoadScriptHelper.getHostBundleJsBasePath() + OSF.ConstantNames.DefaultLocale + "/" + OSF.LoadScriptHelper.getOfficeStringJsName();
                            OUtil.loadScript(fallbackUrl, function (fallbackUrlSuccess) {
                                if (fallbackUrlSuccess) {
                                    resolve();
                                }
                                else {
                                    reject(new Error("Cannot load " + OSF.ConstantNames.OfficeStringJS));
                                }
                            });
                        }
                    });
                });
            }
            return officeStringsJsLoadPromise;
        }
        OUtil.ensureOfficeStringsJs = ensureOfficeStringsJs;
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
            function goToByIdAsync(id, goToType, arg1, arg2) {
                var goToTypeMap = {};
                goToTypeMap[Office.GoToType.Binding] = 0;
                goToTypeMap[Office.GoToType.NamedItem] = 1;
                goToTypeMap[Office.GoToType.Slide] = 2;
                goToTypeMap[Office.GoToType.Index] = 3;
                var selectionModeMap = {};
                selectionModeMap[Office.SelectionMode.Default] = 0;
                selectionModeMap[Office.SelectionMode.Selected] = 1;
                selectionModeMap[Office.SelectionMode.None] = 2;
                var goToTypeHostValue = goToTypeMap[goToType];
                var selectionModeHostValue = 0;
                var callback = arg2;
                if (typeof arg1 === "function") {
                    callback = arg1;
                }
                else if (typeof arg1 !== "undefined") {
                    selectionModeHostValue = selectionModeMap[arg1];
                }
                var asyncMethodExecutor = OSF._OfficeAppFactory.getAsyncMethodExecutor();
                var dataTransform = {
                    toSafeArrayHost: function () {
                        return [id, goToTypeHostValue, selectionModeHostValue];
                    },
                    fromSafeArrayHost: function (payload) {
                        return payload;
                    },
                    toWebHost: function () {
                        var navigationRequestParam = {
                            Id: id,
                            GoToType: goToTypeHostValue,
                            SelectionMode: selectionModeHostValue
                        };
                        var obj = {
                            DdaGoToByIdMethod: navigationRequestParam
                        };
                        return obj;
                    },
                    fromWebHost: function (payload) {
                        return payload;
                    }
                };
                asyncMethodExecutor.executeAsync(82, dataTransform, callback);
            }
            document.goToByIdAsync = goToByIdAsync;
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
                if (appContext) {
                    if (appContext.get_isDialog()) {
                        _requirements = OSF.Requirement.RequirementsMatrixFactory.getDefaultDialogRequirementMatrix(appContext);
                    }
                    else {
                        _requirements = OSF.Requirement.RequirementsMatrixFactory.getDefaultRequirementMatrix(appContext);
                    }
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
        OSF.definePropertyOnNamespace(context, 'webAuth', function () {
            if (OSF.DDA.WebAuth) {
                return OSF.DDA.WebAuth;
            }
            return undefined;
        });
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
        if (!_isOfficeOnReadyCalled) {
            OSF.OUtil.waitForFunction(function () { return (typeof (Office.initialize) === 'function'); }, function (initializedDeclared) {
                if (initializedDeclared) {
                    Office.initialize(OSF._OfficeAppFactory.getOfficeAppContext().get_reason());
                }
            }, 400, 50);
        }
    }
    Office.fireOnReady = fireOnReady;
    function sendTelemetryEvent(telemetryEvent) {
        Microsoft.Office.WebExtension.sendTelemetryEvent(telemetryEvent);
    }
    Office.sendTelemetryEvent = sendTelemetryEvent;
    Microsoft.Office.WebExtension.onReadyInternal = Office.onReadyInternal;
})(Office || (Office = {}));
var OSF;
(function (OSF) {
    var OfficeAppContext = (function () {
        function OfficeAppContext(id, appName, appVersion, appUILocale, dataLocale, docUrl, clientMode, settingsFunc, reason, osfControlType, eToken, correlationId, appInstanceId, touchEnabled, commerceAllowed, appMinorVersion, requirementMatrix, hostCustomMessage, hostFullVersion, clientWindowHeight, clientWindowWidth, addinName, appDomains, dialogRequirementMatrix, featureGates, officeThemeFunc, initialDisplayMode, isFromWacAutomation, wopiHostOriginForSingleSignOn) {
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
            this._isDialog = OSF.OUtil.isDialog();
            this._clientWindowHeight = clientWindowHeight;
            this._clientWindowWidth = clientWindowWidth;
            this._addinName = addinName;
            this._appDomains = appDomains;
            this._dialogRequirementMatrix = dialogRequirementMatrix;
            this._featureGates = featureGates;
            this._officeThemeFunc = officeThemeFunc;
            this._initialDisplayMode = initialDisplayMode;
            this._isFromWacAutomation = isFromWacAutomation;
            this._wopiHostOriginForSingleSignOn = wopiHostOriginForSingleSignOn;
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
        OfficeAppContext.prototype.get_isFromWacAutomation = function () { return this._isFromWacAutomation; };
        OfficeAppContext.prototype.get_wopiHostOriginForSingleSignOn = function () { return this._wopiHostOriginForSingleSignOn; };
        return OfficeAppContext;
    }());
    OSF.OfficeAppContext = OfficeAppContext;
})(OSF || (OSF = {}));
var OSF;
(function (OSF) {
    var _OfficeAppFactory;
    (function (_OfficeAppFactory) {
        var _windowLocationHash;
        var _windowLocationSearch;
        var _windowName;
        if (typeof (window) !== 'undefined') {
            if (window.location) {
                _windowLocationHash = window.location.hash;
                _windowLocationSearch = window.location.search;
            }
            _windowName = window.name;
        }
        var _hostInfo;
        var _webAppState;
        var _isLoggingAllowed;
        var _initializationHelper;
        var _asyncMethodExecutor;
        var _officeAppContext;
        var _clientHostController;
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
            if (!_initializationHelper) {
                onError(new Error("Office.js cannot be initialized."));
                return;
            }
            if (_hostInfo.hostPlatform === OSF.HostInfoPlatform.web) {
                _initializationHelper.saveAndSetDialogInfo(OSF.Utility.getQueryStringValue("_host_Info"));
            }
            _initializationHelper.setAgaveHostCommunication();
            OSFPerformance.getAppContextStart = OSFPerformance.now();
            var onGetAppContextSuccess = function (officeAppContext) {
                OSFPerformance.getAppContextEnd = OSFPerformance.now();
                OSF.AppTelemetry.initialize(officeAppContext);
                _officeAppContext = officeAppContext;
                _clientHostController = _initializationHelper.createClientHostController();
                _asyncMethodExecutor = _initializationHelper.createAsyncMethodExecutor();
                _initializationHelper.prepareApiSurface(officeAppContext);
                if (OSF.BootStrapExtension.onGetAppContext) {
                    OSF.BootStrapExtension.onGetAppContext(officeAppContext, _webAppState.wnd)
                        .then(function () {
                        fireOfficeOnReady(officeAppContext, onSuccess);
                    });
                }
                else {
                    fireOfficeOnReady(officeAppContext, onSuccess);
                }
            };
            var onGetAppContextError = function (e) {
                onError(e);
            };
            _initializationHelper.getAppContext(window, onGetAppContextSuccess, onGetAppContextError);
        }
        _OfficeAppFactory.bootstrap = bootstrap;
        function fireOfficeOnReady(officeAppContext, onSuccess) {
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
            notifyHostOfficeReady();
            onSuccess(officeAppContext);
        }
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
                    var windowName = window.name;
                    if (windowName) {
                        var windowNameObj = JSON.parse(windowName);
                        hostInfoValue = windowNameObj ? windowNameObj["hostInfo"] : null;
                    }
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
            else if (_hostInfo.hostPlatform === OSF.HostInfoPlatform.android || _hostInfo.hostPlatform === OSF.HostInfoPlatform.winrt) {
                _initializationHelper = new OSF.WebViewInitializationHelper(_hostInfo, _webAppState, null, null);
            }
            else {
                console.warn("Office.js is loaded inside in unknown host or platform " + _hostInfo.hostPlatform);
            }
        }
        function isWebkit2Sandbox() {
            return window.webkit && window.webkit.messageHandlers && window.webkit.messageHandlers.Agave;
        }
        function notifyHostOfficeReady() {
            if (_hostInfo.hostPlatform == OSF.HostInfoPlatform.web) {
                if (_webAppState.clientEndPoint != null) {
                    _webAppState.clientEndPoint.invoke("ContextActivationManager_notifyHost", null, [_webAppState.id, OSF.AgaveHostAction.OfficeJsReady, Date.now()]);
                }
            }
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
        function getClientHostController() {
            return _clientHostController;
        }
        _OfficeAppFactory.getClientHostController = getClientHostController;
    })(_OfficeAppFactory = OSF._OfficeAppFactory || (OSF._OfficeAppFactory = {}));
    function getClientEndPoint() {
        return _OfficeAppFactory.getWebAppState().clientEndPoint;
    }
    OSF.getClientEndPoint = getClientEndPoint;
})(OSF || (OSF = {}));
var Office;
(function (Office) {
    var AsyncResultStatus;
    (function (AsyncResultStatus) {
        AsyncResultStatus["Succeeded"] = "succeeded";
        AsyncResultStatus["Failed"] = "failed";
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
    var GoToType;
    (function (GoToType) {
        GoToType["Binding"] = "binding";
        GoToType["NamedItem"] = "namedItem";
        GoToType["Slide"] = "slide";
        GoToType["Index"] = "index";
    })(GoToType = Office.GoToType || (Office.GoToType = {}));
    var SelectionMode;
    (function (SelectionMode) {
        SelectionMode["Default"] = "default";
        SelectionMode["Selected"] = "selected";
        SelectionMode["None"] = "none";
    })(SelectionMode = Office.SelectionMode || (Office.SelectionMode = {}));
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
            OSF.Utility.xdmDebugLog("registerConversation: cId=" + conversationId + " Url=" + conversationUrl);
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
            this._hostTrustCheckStatus = 0;
            this._checkStatusLogged = false;
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
            this._checkReceiverOriginAndRun = null;
            ;
        }
        ;
        Object.defineProperty(XdmClientEndPoint.prototype, "targetUrl", {
            get: function () {
                return this._targetUrl;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(XdmClientEndPoint.prototype, "hostTrustCheckStatus", {
            get: function () {
                return this._hostTrustCheckStatus;
            },
            set: function (value) {
                this._hostTrustCheckStatus = value;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(XdmClientEndPoint.prototype, "checkStatusLogged", {
            get: function () {
                return this._checkStatusLogged;
            },
            set: function (value) {
                this._checkStatusLogged = value;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(XdmClientEndPoint.prototype, "checkReceiverOriginAndRun", {
            get: function () {
                return this._checkReceiverOriginAndRun;
            },
            set: function (value) {
                this._checkReceiverOriginAndRun = value;
            },
            enumerable: true,
            configurable: true
        });
        XdmClientEndPoint.prototype.invoke = function (targetMethodName, callback, param) {
            var _this = this;
            var funcToRun = function () {
                var correlationId = _this._callingIndex++;
                var now = new Date();
                var callbackEntry = { "callback": callback, "createdOn": now.getTime() };
                if (param && typeof param === "object" && typeof param.__timeout__ === "number") {
                    callbackEntry.timeout = param.__timeout__;
                    delete param.__timeout__;
                }
                _this._callbackList[correlationId] = callbackEntry;
                try {
                    if (_this._hostTrustCheckStatus !== 3) {
                        if (targetMethodName !== "ContextActivationManager_getAppContextAsync") {
                            throw "Access Denied";
                        }
                    }
                    var callRequest = new XdmRequest(targetMethodName, 0, _this._conversationId, correlationId, param);
                    var msg = XdmMessagePackager.envelope(callRequest, _this._serializerVersion);
                    _this._targetWindow.postMessage(msg, _this._targetUrl);
                    XdmCommunicationManager._startMethodTimeoutTimer();
                }
                catch (ex) {
                    try {
                        if (callback !== null)
                            callback(-1, ex);
                    }
                    finally {
                        delete _this._callbackList[correlationId];
                    }
                }
            };
            if (this._checkReceiverOriginAndRun) {
                this._checkReceiverOriginAndRun(funcToRun);
            }
            else {
                this._hostTrustCheckStatus = 3;
                funcToRun();
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
                OSF.Utility.xdmDebugLog("Unknown conversation Id.");
            }
            return clientEndPoint;
        }
        ;
        function _lookupMethodObject(serviceEndPoint, messageObject) {
            var methodOrEventMethodObject = serviceEndPoint._methodObjectList[messageObject._actionName];
            if (!methodOrEventMethodObject) {
                OSF.Utility.xdmDebugLog("The specified method is not registered on service endpoint:" + messageObject._actionName);
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
                OSF.Utility.xdmDebugLog("channel is not ready.");
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
                OSF.Utility.xdmDebugLog("channel is not ready.");
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
                OSF.Utility.xdmDebugLog("Browser doesn't support the required API.");
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
            var regexHostNameStringArray = new Array("^outlook\\.office\\.com$", "^outlook-sdf\\.office\\.com$", "^outlook\\.office\\.com$", "^outlook-sdf\\.office\\.com$", "^outlook\\.live\\.com$", "^outlook-sdf\\.live\\.com$", "^consumer\\.live-int\\.com$", "^outlook-tdf\\.live\\.com$", "^sdfpilot\\.live\\.com$", "^outlook\\.office\\.de$", "^outlook\\.office365\\.us$", "^outlook\\.office365\\.com$", "^partner\\.outlook\\.cn$", "^exchangelabs\\.live-int\\.com$", "^office-int\\.com$", "^officeapps\\.live-int\\.com$", "^.*\\.dod\\.online\\.office365\\.us$", "^.*\\.gov\\.online\\.office365\\.us$", "^.*\\.officeapps\\.live\\.com$", "^.*\\.officeapps\\.live-int\\.com$", "^.*\\.officeapps-df\\.live\\.com$", "^.*\\.online\\.office\\.de$", "^.*\\.partner\\.officewebapps\\.cn$", "^.*\\.office\\.net$", "^" + document.domain.replace(new RegExp("\\.", "g"), "\\.") + "$");
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
                OSF.Utility.xdmDebugLog(serializedMessage);
                if (messageObject._messageType === 0) {
                    var requesterUrl = (e.origin == null || e.origin === "null") ? messageObject._origin : e.origin;
                    try {
                        var serviceEndPoint = _lookupServiceEndPoint(messageObject._conversationId);
                        OSF.Utility.xdmDebugLog("_receive: request, origin=" + requesterUrl + " sourceURL:" + serviceEndPoint._conversations[messageObject._conversationId]);
                        var conversation = serviceEndPoint._conversations[messageObject._conversationId];
                        serializerVersion = conversation.serializerVersion != null ? conversation.serializerVersion : serializerVersion;
                        OSF.Utility.xdmDebugLog("_receive: request, origin=" + requesterUrl + " sourceURL:" + conversation.url);
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
                    if (!clientEndPoint) {
                        return;
                    }
                    clientEndPoint._serializerVersion = serializerVersion;
                    OSF.Utility.xdmDebugLog("_receive: response, origin=" + e.origin + " targetURL:" + clientEndPoint._targetUrl);
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
        function isHostNameValidWacDomain(hostName) {
            return _isHostNameValidWacDomain(hostName);
        }
        XdmCommunicationManager.isHostNameValidWacDomain = isHostNameValidWacDomain;
    })(XdmCommunicationManager = OSF.XdmCommunicationManager || (OSF.XdmCommunicationManager = {}));
    var XdmMessage = (function () {
        function XdmMessage(messageType, actionName, conversationId, correlationId, data) {
            this._messageType = messageType;
            this._actionName = actionName;
            this._conversationId = conversationId;
            this._correlationId = correlationId;
            this._origin = window.location.origin;
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
                    OSF.Utility.xdmDebugLog("_send: requestUrl=" + _this._requesterUrl + " _actionName:" + _this._actionName);
                }
                catch (ex) {
                    OSF.Utility.xdmDebugLog("ResponseSender._send error:" + ex.message);
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
                    OSF.Utility.xdmDebugLog("InvokeCompleteCallback._send error:" + ex.message);
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
                        diagnosticLevel: oteljs.DiagnosticLevel.NecessaryServiceDataEvent
                    }
                });
            });
        }
    }
    OSFPerfUtil.sendPerformanceTelemetry = sendPerformanceTelemetry;
})(OSFPerfUtil || (OSFPerfUtil = {}));
var OSF;
(function (OSF) {
    OSF.Flights = [];
    var TestFlightStart = 1000;
    var TestFlightEnd = 1009;
    var OUtil;
    (function (OUtil) {
        var _uniqueId = -1;
        var _xdmInfoKey = '&_xdm_Info=';
        var _serializerVersionKey = '&_serializer_version=';
        var _flightsKey = '&_flights=';
        var _xdmSessionKeyPrefix = '_xdm_';
        var _serializerVersionKeyPrefix = '_serializer_version=';
        var _flightsKeyPrefix = '_flights=';
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
        function isEdge() {
            return typeof (window) !== "undefined" && typeof (window.navigator) !== "undefined" && window.navigator.userAgent.indexOf("Edge") > 0;
        }
        function isIE() {
            return typeof (window) !== "undefined" && typeof (window.navigator) !== "undefined" && window.navigator.userAgent.indexOf("Trident") > 0;
        }
        function startsWith(originalString, patternToCheck) {
            return originalString.substr(0, patternToCheck.length) === patternToCheck;
        }
        function containsPort(url, protocol, hostname, portNumber) {
            return startsWith(url, protocol + "//" + hostname + ":" + portNumber) || startsWith(url, hostname + ":" + portNumber);
        }
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
        function parseFlights(skipSessionStorage) {
            var flights = parseFlightsWithGivenFragment(skipSessionStorage, window.location.hash);
            if (flights.length == 0) {
                flights = parseFlightsFromWindowName(skipSessionStorage, window.name);
            }
            return flights;
        }
        OUtil.parseFlights = parseFlights;
        function checkFlight(flight) {
            return OSF.Flights && OSF.Flights.indexOf(flight) >= 0;
        }
        OUtil.checkFlight = checkFlight;
        function parseFlightsFromWindowName(skipSessionStorage, windowName) {
            return parseArrayWithDefault(parseInfoFromWindowName(skipSessionStorage, windowName, "flights"));
        }
        function parseFlightsWithGivenFragment(skipSessionStorage, fragment) {
            return parseArrayWithDefault(parseInfoWithGivenFragment(_flightsKey, _flightsKeyPrefix, true, skipSessionStorage, fragment));
        }
        function parseArrayWithDefault(jsonString) {
            var array = [];
            try {
                array = JSON.parse(jsonString);
            }
            catch (ex) { }
            if (!Array.isArray(array)) {
                array = [];
            }
            return array;
        }
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
        function parseUrl(url, enforceHttps) {
            if (enforceHttps === void 0) { enforceHttps = false; }
            if (typeof url === "undefined" || !url) {
                return undefined;
            }
            var notHttpsErrorMessage = "NotHttps";
            var invalidUrlErrorMessage = "InvalidUrl";
            var isIEBoolean = isIE();
            var isEdgeBoolean = isEdge();
            var parsedUrlObj = {
                protocol: undefined,
                hostname: undefined,
                host: undefined,
                port: undefined,
                pathname: undefined,
                search: undefined,
                hash: undefined,
                isPortPartOfUrl: undefined
            };
            try {
                if (isIEBoolean) {
                    var parser = document.createElement("a");
                    parser.href = url;
                    if (!parser || !parser.protocol || !parser.host || !parser.hostname || !parser.href
                        || this.cleanUrl(parser.href) !== this.cleanUrl(url)) {
                        throw invalidUrlErrorMessage;
                    }
                    if (OSF.OUtil.checkFlight(2)) {
                        if (enforceHttps && parser.protocol != "https:")
                            throw new Error(notHttpsErrorMessage);
                    }
                    var redundandPortString = this.getRedundandPortString(url, parser);
                    parsedUrlObj.protocol = parser.protocol;
                    parsedUrlObj.hostname = parser.hostname;
                    parsedUrlObj.port = (redundandPortString == "") ? parser.port : "";
                    parsedUrlObj.host = (redundandPortString != "") ? parser.hostname : parser.host;
                    parsedUrlObj.pathname = (isIEBoolean ? "/" : "") + parser.pathname;
                    parsedUrlObj.search = parser.search;
                    parsedUrlObj.hash = parser.hash;
                    parsedUrlObj.isPortPartOfUrl = this.containsPort(url, parser.protocol, parser.hostname, parser.port);
                }
                else {
                    var urlObj = new URL(url);
                    if (urlObj && urlObj.protocol && urlObj.host && urlObj.hostname) {
                        if (OSF.OUtil.checkFlight(2)) {
                            if (enforceHttps && urlObj.protocol != "https:")
                                throw new Error(notHttpsErrorMessage);
                        }
                        parsedUrlObj.protocol = urlObj.protocol;
                        parsedUrlObj.hostname = urlObj.hostname;
                        parsedUrlObj.port = urlObj.port;
                        parsedUrlObj.host = urlObj.host;
                        parsedUrlObj.pathname = urlObj.pathname;
                        parsedUrlObj.search = urlObj.search;
                        parsedUrlObj.hash = urlObj.hash;
                        parsedUrlObj.isPortPartOfUrl = urlObj.host.lastIndexOf(":" + urlObj.port) == (urlObj.host.length - urlObj.port.length - 1);
                    }
                }
            }
            catch (err) {
                if (err.message === notHttpsErrorMessage)
                    throw err;
            }
            return parsedUrlObj;
        }
        OUtil.parseUrl = parseUrl;
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
        function waitForFunction(test, callback, numberOfTries, delay) {
            var attemptsRemaining = numberOfTries;
            var timerId;
            var validateFunction = function () {
                attemptsRemaining--;
                if (test()) {
                    callback(true);
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
        }
        OUtil.waitForFunction = waitForFunction;
        function defineMethodOnNamespace(o, name, method) {
            o[name] = method;
        }
        OUtil.defineMethodOnNamespace = defineMethodOnNamespace;
        function isDialog() {
            return OSF._OfficeAppFactory.getHostInfo().isDialog;
        }
        OUtil.isDialog = isDialog;
        function isPopupWindow() {
            return OSF.OUtil.isDialog()
                && OSF._OfficeAppFactory.getHostInfo().hostPlatform == OSF.HostInfoPlatform.web
                && window.opener != null;
        }
        OUtil.isPopupWindow = isPopupWindow;
        function getHostPlatform() {
            var hostInfo = OSF._OfficeAppFactory.getHostInfo();
            return hostInfo.hostPlatform;
        }
        OUtil.getHostPlatform = getHostPlatform;
    })(OUtil = OSF.OUtil || (OSF.OUtil = {}));
})(OSF || (OSF = {}));
var OSF;
(function (OSF) {
    var ParameterNames;
    (function (ParameterNames) {
        ParameterNames.Callback = "callback";
        ParameterNames.AsyncContext = "asyncContext";
        ParameterNames.Data = "data";
        ParameterNames.MessageToParent = "messageToParent";
        ParameterNames.MessageContent = "messageContent";
        ParameterNames.AppCommandInvocationCompletedData = "appCommandInvocationCompletedData";
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
            this._registerHandlers = [];
            this._eventDispatch = new OSF.EventDispatch([
                {
                    type: OSF.EventType.RichApiMessage,
                    id: OSF.EventDispId.dispidRichApiMessageEvent,
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
        RichApiMessageManager.prototype.register = function (handler) {
            var _this = this;
            if (!this._registerPromise) {
                this._registerPromise = new Office.Promise(function (resolve, reject) {
                    _this.addHandlerAsync(OSF.EventType.RichApiMessage, function (args) {
                        _this._registerHandlers.forEach(function (value) {
                            if (value) {
                                value(args);
                            }
                        });
                    }, function (asyncResult) {
                        if (asyncResult.status == 'failed') {
                            reject(asyncResult.error);
                        }
                        else {
                            resolve();
                        }
                    });
                });
            }
            return this._registerPromise.then(function () {
                _this._registerHandlers.push(handler);
            });
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
            var appType;
            var id;
            var version;
            var minorVersion;
            var UILocale;
            var dataLocale;
            var docUrl;
            var clientMode;
            var activationMode;
            var reason;
            var osfControlType;
            var eToken;
            var correlationId;
            var appInstanceId;
            var touchEnabled;
            var commerceAllowed;
            var requirementMatrix;
            var hostCustomMessage;
            var hostFullVersion;
            var dialogRequirementMatrix;
            var sdxFeatureGates;
            var initialDisplayMode = 0;
            var settingsFunc;
            var officeThemeFunc;
            var fallback = false;
            var externalNativeFunctionExists = OSF.Utility.externalNativeFunctionExists;
            if (!externalNativeFunctionExists(typeof context.GetContextDataInJson)) {
                fallback = true;
            }
            else {
                var contextJsonString;
                if (typeof OsfOMToken !== 'undefined' && OsfOMToken) {
                    contextJsonString = context.GetContextDataInJson(OsfOMToken);
                    var contextJson;
                    if (contextJsonString) {
                        contextJson = JSON.parse(contextJsonString);
                    }
                    if (!contextJson) {
                        fallback = true;
                    }
                    else {
                        appType = contextJson.appType;
                        id = contextJson.solutionRef;
                        version = contextJson.versionMajor;
                        minorVersion = contextJson.versionMinor;
                        UILocale = contextJson.uiLocale;
                        dataLocale = contextJson.dataLocale;
                        docUrl = contextJson.docUrl;
                        clientMode = contextJson.clientMode;
                        activationMode = contextJson.activationMode;
                        osfControlType = contextJson.controlType;
                        eToken = contextJson.eToken;
                        correlationId = contextJson.correlationId;
                        appInstanceId = contextJson.appInstanceId;
                        touchEnabled = contextJson.touchEnabled;
                        commerceAllowed = context.commerceAllowed;
                        requirementMatrix = contextJson.requirementMatrix;
                        hostFullVersion = contextJson.hostFullVersion;
                        dialogRequirementMatrix = contextJson.requirementMatrix;
                        var sdxFeatureGatesJson = contextJson.featureGates;
                        if (sdxFeatureGatesJson) {
                            sdxFeatureGates = JSON.parse(sdxFeatureGatesJson);
                        }
                        initialDisplayMode = contextJson.initialDisplayMode;
                        settingsFunc = function () {
                            var settingsString = contextJson.settings;
                            var settings;
                            if (settingsString) {
                                settings = JSON.parse(settingsString);
                            }
                            var serializedSettings = {};
                            if (settings) {
                                var names = settings.names;
                                var values = settings.values;
                                for (var index = 0; index < names.length; index++) {
                                    serializedSettings[names[index]] = values[index];
                                }
                            }
                            return serializedSettings;
                        };
                        officeThemeFunc = function () {
                            var osfOfficeThemeInfoString = contextJson.themeInfo;
                            return _this.getOfficeThemeFromInfoString(osfOfficeThemeInfoString);
                        };
                    }
                }
                else {
                    fallback = true;
                }
            }
            if (fallback) {
                appType = context.GetAppType();
                id = context.GetSolutionRef();
                version = context.GetAppVersionMajor();
                minorVersion = context.GetAppVersionMinor();
                UILocale = context.GetAppUILocale();
                dataLocale = context.GetAppDataLocale();
                docUrl = context.GetDocUrl();
                clientMode = context.GetAppCapabilities();
                activationMode = context.GetActivationMode();
                osfControlType = context.GetControlIntegrationLevel();
                try {
                    eToken = context.GetSolutionToken();
                }
                catch (ex) {
                }
                var externalNativeFunctionExists = OSF.Utility.externalNativeFunctionExists;
                if (externalNativeFunctionExists(typeof context.GetCorrelationId)) {
                    correlationId = context.GetCorrelationId();
                }
                if (externalNativeFunctionExists(typeof context.GetInstanceId)) {
                    appInstanceId = context.GetInstanceId();
                }
                if (externalNativeFunctionExists(typeof context.GetTouchEnabled)) {
                    touchEnabled = context.GetTouchEnabled();
                }
                if (externalNativeFunctionExists(typeof context.GetCommerceAllowed)) {
                    commerceAllowed = context.GetCommerceAllowed();
                }
                if (externalNativeFunctionExists(typeof context.GetSupportedMatrix)) {
                    requirementMatrix = context.GetSupportedMatrix();
                }
                if (externalNativeFunctionExists(typeof context.GetHostCustomMessage)) {
                    hostCustomMessage = context.GetHostCustomMessage();
                }
                if (externalNativeFunctionExists(typeof context.GetHostFullVersion)) {
                    hostFullVersion = context.GetHostFullVersion();
                }
                if (externalNativeFunctionExists(typeof context.GetDialogRequirementMatrix)) {
                    dialogRequirementMatrix = context.GetDialogRequirementMatrix();
                }
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
                if (externalNativeFunctionExists(typeof context.GetInitialDisplayMode)) {
                    initialDisplayMode = context.GetInitialDisplayMode();
                }
                settingsFunc = function () { return _this.getSerializedSettings(); };
                officeThemeFunc = function () { return _this.getOfficeTheme(); };
            }
            reason = (activationMode === 2) ? Office.InitializationReason.DocumentOpened : Office.InitializationReason.Inserted;
            eToken = eToken ? eToken.toString() : "";
            var returnedContext = new OSF.OfficeAppContext(id, appType, version, UILocale, dataLocale, docUrl, clientMode, settingsFunc, reason, osfControlType, eToken, correlationId, appInstanceId, touchEnabled, commerceAllowed, minorVersion, requirementMatrix, hostCustomMessage, hostFullVersion, undefined, undefined, undefined, undefined, dialogRequirementMatrix, sdxFeatureGates, officeThemeFunc, initialDisplayMode, undefined, undefined);
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
            return this.getOfficeThemeFromInfoString(osfOfficeThemeInfoString);
        };
        RichClientInitializationHelper.prototype.getOfficeThemeFromInfoString = function (osfOfficeThemeInfoString) {
            var osfOfficeTheme;
            if (osfOfficeThemeInfoString) {
                osfOfficeTheme = JSON.parse(osfOfficeThemeInfoString);
                for (var color in osfOfficeTheme) {
                    osfOfficeTheme[color] = OSF.OUtil.convertIntToCssHexColor(osfOfficeTheme[color]);
                }
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
            _this._asyncMethodExecutorHelper = OSF.BootStrapExtension.createAsyncMethodExecutorHelper(_this);
            return _this;
        }
        SafeArrayAsyncMethodExecutor.prototype.executeAsync = function (id, dataTransform, callback, asyncContext) {
            var _this = this;
            try {
                var chunkResultData = new Array();
                this._clientHostController.execute(id, dataTransform.toSafeArrayHost(), function (hostResponseArgsNative, resultCode) {
                    var hostResponseArgs = OSF.Utility.fromSafeArray(hostResponseArgsNative);
                    return _this._asyncMethodExecutorHelper.handleSafeArrayHostResponse(hostResponseArgs, resultCode, chunkResultData, callback, dataTransform, id, asyncContext);
                });
            }
            catch (ex) {
                this.onException(ex, id, callback);
            }
        };
        SafeArrayAsyncMethodExecutor.prototype.registerEventAsync = function (id, eventType, targetId, handler, dataTransform, callback, asyncContext) {
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
                    _this.invokeCallback(id, callback, status, null, asyncContext);
                    return true;
                });
            }
            catch (ex) {
                this.onException(ex, id, callback);
            }
        };
        SafeArrayAsyncMethodExecutor.prototype.unregisterEventAsync = function (id, eventType, targetId, callback, asyncContext) {
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
                    _this.invokeCallback(id, callback, status, null, asyncContext);
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
                        if (dispId == OSF.EventDispId.dispidDialogMessageReceivedEvent) {
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
                    id: OSF.EventDispId.dispidSettingsChangedEvent,
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
                if (result.status === Office.AsyncResultStatus.Succeeded) {
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
    function isOfficeReactNative() {
        try {
            return (typeof OfficePlatformGlobal !== 'undefined'
                && typeof OfficePlatformGlobal.ReactNativeReka !== 'undefined');
        }
        catch (e) {
            return false;
        }
    }
    OSF.isOfficeReactNative = isOfficeReactNative;
    var Utility;
    (function (Utility) {
        function createParameterException(message) {
            return new Error("Parameter count mismatch: " + message);
        }
        Utility.createParameterException = createParameterException;
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
        function xdmDebugLog(message) {
            if (Utility._DebugXdm) {
                console.log(message);
            }
        }
        Utility.xdmDebugLog = xdmDebugLog;
        function enableDebugXdm() {
            Utility._DebugXdm = true;
        }
        Utility.enableDebugXdm = enableDebugXdm;
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
            if (typeof (window) !== 'undefined' && window.location && window.location.search) {
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
            if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
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
                    status: Office.AsyncResultStatus.Succeeded
                };
            }
            else {
                return {
                    status: Office.AsyncResultStatus.Failed,
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
            _this._asyncMethodExecutorHelper = OSF.BootStrapExtension.createAsyncMethodExecutorHelper(_this);
            return _this;
        }
        WebAsyncMethodExecutor.prototype.executeAsync = function (id, dataTransform, callback, asyncContext) {
            var _this = this;
            this._clientHostController.execute(id, dataTransform.toWebHost(), function (resultCode, payload) {
                if (callback) {
                    _this._asyncMethodExecutorHelper.handleWebHostResponse(payload, resultCode, callback, dataTransform, id, asyncContext);
                }
                return true;
            });
        };
        WebAsyncMethodExecutor.prototype.registerEventAsync = function (id, eventType, targetId, handler, dataTransform, callback, asyncContext) {
            var _this = this;
            this._clientHostController.registerEvent(id, eventType, targetId, function (payload) {
                var eventPayload = payload;
                var eventArgs = dataTransform.fromWebHost(eventPayload);
                handler(eventArgs);
            }, function (resultCode, payload) {
                if (callback) {
                    _this.invokeCallback(id, callback, resultCode, null, asyncContext);
                }
                return true;
            });
        };
        WebAsyncMethodExecutor.prototype.unregisterEventAsync = function (id, eventType, targetId, callback, asyncContext) {
            var _this = this;
            this._clientHostController.unregisterEvent(id, eventType, targetId, function (resultCode, payload) {
                if (callback) {
                    _this.invokeCallback(id, callback, resultCode, null, asyncContext);
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
    var DDA;
    (function (DDA) {
        var WebAuth;
        (function (WebAuth) {
            function getAuthContextAsync(callback) {
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
                        return payload.authContext;
                    }
                };
                asyncMethodExecutor.executeAsync(99, dataTransform, callback);
            }
            WebAuth.getAuthContextAsync = getAuthContextAsync;
        })(WebAuth = DDA.WebAuth || (DDA.WebAuth = {}));
    })(DDA = OSF.DDA || (OSF.DDA = {}));
    var WebAuth;
    (function (WebAuth) {
        var AuthFlow;
        (function (AuthFlow) {
            AuthFlow["Implicit"] = "implicit";
            AuthFlow["AuthCode"] = "authcode";
            AuthFlow["Broker"] = "broker";
        })(AuthFlow = WebAuth.AuthFlow || (WebAuth.AuthFlow = {}));
        WebAuth.loadAttempts = 0;
        function load() {
            if (WebAuth.config && WebAuth.config.authFlow === AuthFlow.AuthCode) {
                return loadForAuthCode();
            }
            else {
                return loadForImplicit();
            }
        }
        WebAuth.load = load;
        function getToken(target, applicationId, correlationId, popup) {
            var authLib;
            if (WebAuth.config && WebAuth.config.authFlow === AuthFlow.AuthCode) {
                authLib = BrowserAuth;
            }
            else {
                authLib = Implicit;
            }
            return authLib.GetToken(target, applicationId, correlationId, !!popup);
        }
        WebAuth.getToken = getToken;
        function loadForImplicit() {
            WebAuth.loadAttempts++;
            var IMPLICIT_DEBUG = 'webauth/webauth.implicit.debug.js';
            var IMPLICIT_SHIP = 'webauth/webauth.implicit.js';
            var Implicit_Cdn_Path = OSF.LoadScriptHelper.getHostBundleJsBasePath() + ((WebAuth.config && WebAuth.config.debugging) ? IMPLICIT_DEBUG : IMPLICIT_SHIP);
            return new Promise(function (resolve, reject) {
                OSF.OUtil.loadScript(Implicit_Cdn_Path, function () {
                    if (WebAuth.config) {
                        resolve(Implicit.Load(WebAuth.config, OSF._OfficeAppFactory.getHostInfo().osfControlAppCorrelationId));
                    }
                    else {
                        Implicit.GetAuthConfig().then(function (configParent) {
                            WebAuth.config = configParent;
                            resolve(Implicit.Load(WebAuth.config, OSF._OfficeAppFactory.getHostInfo().osfControlAppCorrelationId));
                        }, function () { reject(null); });
                    }
                });
            });
        }
        function loadForAuthCode() {
            WebAuth.loadAttempts++;
            var BROWSERAUTH_PATH = 'webauth/';
            var BROWSERAUTH_JS_DEBUG = 'webauth.browserauth.debug.js';
            var BROWSERAUTH_JS_SHIP = 'webauth.browserauth.js';
            var BrowserAuth_Js = (WebAuth.config && WebAuth.config.debugging) ? BROWSERAUTH_JS_DEBUG : BROWSERAUTH_JS_SHIP;
            var BrowserAuth_Cdn_Path = (WebAuth.config && WebAuth.config.authVersion)
                ? OSF.LoadScriptHelper.getHostBundleJsBasePath() + BROWSERAUTH_PATH + WebAuth.config.authVersion + "/" + BrowserAuth_Js
                : OSF.LoadScriptHelper.getHostBundleJsBasePath() + BROWSERAUTH_PATH + BrowserAuth_Js;
            return new Promise(function (resolve, reject) {
                OSF.OUtil.loadScript(BrowserAuth_Cdn_Path, function () {
                    if (WebAuth.config) {
                        BrowserAuth.Load(WebAuth.config, OSF._OfficeAppFactory.getHostInfo().osfControlAppCorrelationId).then(function (result) { resolve(result); }, function (result) { reject(result); });
                    }
                    else {
                        BrowserAuth.GetAuthConfig().then(function (configParent) {
                            WebAuth.config = configParent;
                            BrowserAuth.Load(WebAuth.config, OSF._OfficeAppFactory.getHostInfo().osfControlAppCorrelationId).then(function (result) { resolve(result); }, function (result) { reject(result); });
                        }, function () { reject(null); });
                    }
                });
            });
        }
    })(WebAuth = OSF.WebAuth || (OSF.WebAuth = {}));
})(OSF || (OSF = {}));
var OSF;
(function (OSF) {
    var WebClientHostController = (function () {
        function WebClientHostController(webAppState) {
            this._delegateVersion = 1;
            this._webAppState = webAppState;
            this._webClientHostControllerHelper = OSF.BootStrapExtension.createWebClientHostControllerHelper(this._webAppState, this._delegateVersion);
        }
        WebClientHostController.prototype.execute = function (id, params, callback) {
            var _this = this;
            var hostCallArgs = this._webClientHostControllerHelper.getHostCallArgs(id, params);
            var targetMethodName = this._webClientHostControllerHelper.getTargetMethodName(id);
            this._webAppState.clientEndPoint.invoke(targetMethodName, function (xdmStatus, payload) {
                var error = 0;
                if (xdmStatus == 0) {
                    _this._delegateVersion = payload["Version"];
                    error = _this._webClientHostControllerHelper.parseErrorFromPayload(id, payload);
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
    var WebClientHostControllerHelper = (function () {
        function WebClientHostControllerHelper(webAppState, delegateVersion) {
            this._webAppState = webAppState;
            this._delegateVersion = delegateVersion;
        }
        WebClientHostControllerHelper.prototype.getHostCallArgs = function (id, params) {
            var hostCallArgs = params;
            if (!hostCallArgs) {
                hostCallArgs = {};
            }
            hostCallArgs.DdaMethod = {
                ControlId: this.getControlId(),
                DispatchId: id,
                Version: this._delegateVersion
            };
            hostCallArgs.__timeout__ = -1;
            return hostCallArgs;
        };
        WebClientHostControllerHelper.prototype.getTargetMethodName = function (id) {
            return 'executeMethod';
        };
        WebClientHostControllerHelper.prototype.parseErrorFromPayload = function (id, payload) {
            return payload["Error"];
        };
        WebClientHostControllerHelper.prototype.getControlId = function () {
            return this._webAppState.id;
        };
        return WebClientHostControllerHelper;
    }());
    OSF.WebClientHostControllerHelper = WebClientHostControllerHelper;
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
                if (result.status === Office.AsyncResultStatus.Succeeded) {
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
                if (result.status === Office.AsyncResultStatus.Succeeded) {
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
        Object.defineProperty(WebInitializationHelper.prototype, "isHostOriginTrusted", {
            get: function () {
                return this._isHostOriginTrustedFunc;
            },
            set: function (value) {
                this._isHostOriginTrustedFunc = value;
            },
            enumerable: true,
            configurable: true
        });
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
        WebInitializationHelper.prototype.checkReceiverOriginAndRun = function (funcToRun) {
            var me = this;
            var parsedHostname = OSF.OUtil.parseUrl(me._webAppState.clientEndPoint.targetUrl, false);
            var isHttps = parsedHostname.protocol == "https:";
            var notHttpsErrorMessage = "NotHttps";
            if (me._webAppState.clientEndPoint.hostTrustCheckStatus === 0) {
                if (!isHttps)
                    me._webAppState.clientEndPoint.hostTrustCheckStatus = 2;
                if (me._webAppState.clientEndPoint.hostTrustCheckStatus != 2) {
                    var isOriginValid = OSF.XdmCommunicationManager.isHostNameValidWacDomain(parsedHostname.hostname);
                    if (me.isHostOriginTrusted) {
                        isOriginValid = isOriginValid || me.isHostOriginTrusted(parsedHostname.hostname);
                    }
                    if (isOriginValid)
                        me._webAppState.clientEndPoint.hostTrustCheckStatus = 3;
                }
            }
            if (!me._webAppState.clientEndPoint.checkStatusLogged && me._hostInfo != null && me._hostInfo !== undefined) {
                OSF.AppTelemetry.onCheckWACHost(me._webAppState.clientEndPoint.hostTrustCheckStatus, me._webAppState.id, me._hostInfo.hostType, me._hostInfo.hostPlatform, me._webAppState.clientEndPoint.targetUrl);
                me._webAppState.clientEndPoint.checkStatusLogged = true;
            }
            if (me._webAppState.clientEndPoint.hostTrustCheckStatus != 3) {
                var loadAgaveErrorUX = function () {
                    var officejsCDNHost = OSF.LoadScriptHelper.getHostBundleJsBasePath().match(/^https?:\/\/[^:/?#]*(?::([0-9]+))?/);
                    if (officejsCDNHost && officejsCDNHost[0]) {
                        var agaveErrorUXPath = OSF.LoadScriptHelper.getHostBundleJsBasePath() + 'AgaveErrorUX/index.html#';
                        var hashObj = {
                            error: "NotTrustedWAC",
                            locale: OSF.getSupportedLocale(me._hostInfo.hostLocale, OSF.ConstantNames.DefaultLocale),
                            hostname: parsedHostname.hostname,
                            noHttps: !isHttps,
                            validate: false
                        };
                        var hostUserTrustIframe = document.createElement("iframe");
                        hostUserTrustIframe.style.visibility = "hidden";
                        hostUserTrustIframe.style.height = "0";
                        hostUserTrustIframe.style.width = "0";
                        var hostUserTrustCallback = function (event) {
                            if ((event.source == hostUserTrustIframe.contentWindow) &&
                                (event.origin == officejsCDNHost[0])) {
                                try {
                                    var receivedObj = JSON.parse(event.data);
                                    if (receivedObj.hostUserTrusted === true) {
                                        me._webAppState.clientEndPoint.hostTrustCheckStatus = 3;
                                        OSF.OUtil.removeEventListener(window, "message", hostUserTrustCallback);
                                        document.body.removeChild(hostUserTrustIframe);
                                    }
                                    else {
                                        hashObj.validate = false;
                                        window.location.replace(agaveErrorUXPath + encodeURIComponent(JSON.stringify(hashObj)));
                                    }
                                    funcToRun();
                                }
                                catch (e) {
                                    OSF.OUtil.ensureOfficeStringsJs().then(function () {
                                        document.body.innerHTML = Strings.OfficeOM.L_NotTrustedWAC;
                                    });
                                }
                            }
                        };
                        OSF.OUtil.addEventListener(window, "message", hostUserTrustCallback);
                        hashObj.validate = true;
                        hostUserTrustIframe.setAttribute('src', agaveErrorUXPath + encodeURIComponent(JSON.stringify(hashObj)));
                        hostUserTrustIframe.onload = function () {
                            var postingObj = {
                                hostname: parsedHostname.hostname,
                                noHttps: !isHttps
                            };
                            hostUserTrustIframe.contentWindow.postMessage(JSON.stringify(postingObj), officejsCDNHost[0]);
                        };
                        document.body.appendChild(hostUserTrustIframe);
                    }
                    else {
                        OSF.OUtil.ensureOfficeStringsJs().then(function () {
                            document.body.innerHTML = Strings.OfficeOM.L_NotTrustedWAC;
                        });
                    }
                    if (OSF.OUtil.checkFlight(2)) {
                        if (!isHttps)
                            throw new Error(notHttpsErrorMessage);
                    }
                };
                if (document.body) {
                    loadAgaveErrorUX();
                }
                else {
                    var checkDone = false;
                    document.addEventListener('DOMContentLoaded', function () {
                        if (!checkDone) {
                            checkDone = true;
                            loadAgaveErrorUX();
                        }
                    });
                }
            }
            else {
                funcToRun();
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
                me._webAppState.clientEndPoint.checkReceiverOriginAndRun = function (funcToRun) {
                    me.checkReceiverOriginAndRun(funcToRun);
                };
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
                if (appContext._appName === OSF.AppName.ExcelWebApp) {
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
                    appContext._docUrl != undefined && appContext._clientMode != undefined && appContext._reason != undefined) {
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
                    var returnedContext = new OSF.OfficeAppContext(appContext._id, appContext._appName, appContext._appVersion, appContext._appUILocale, appContext._dataLocale, appContext._docUrl, appContext._clientMode, settingsFunc, appContext._reason, appContext._osfControlType, appContext._eToken, appContext._correlationId, appInstanceId, touchEnabled, commerceAllowed, minorVersion, requirementMatrix, appContext._hostCustomMessage, appContext._hostFullVersion, appContext._clientWindowHeight, appContext._clientWindowWidth, appContext._addinName, appContext._appDomains, appContext._dialogRequirementMatrix, appContext._featureGates, undefined, appContext._initialDisplayMode, appContext._isFromWacAutomation, appContext._wopiHostOriginForSingleSignOn);
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
                else if ((e.ctrlKey || e.metaKey || e.shiftKey || e.altKey) && e.keyCode >= 1 && e.keyCode <= 255) {
                    var params = {
                        "keyCode": e.keyCode,
                        "shiftKey": e.shiftKey,
                        "altKey": e.altKey,
                        "ctrlKey": e.ctrlKey,
                        "metaKey": e.metaKey
                    };
                    me._webAppState.clientEndPoint.invoke("ContextActivationManager_notifyHost", null, [me._webAppState.id, OSF.AgaveHostAction.KeyboardShortcuts, params]);
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
        WebInitializationHelper.prototype.getTabbableElements = function () {
            return this._tabbableElements;
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
                returnedContext = new OSF.OfficeAppContext(id, appType, version, UILocale, dataLocale, docUrl, clientMode, settingsFunc, reason, osfControlType, eToken, correlationId, appInstanceId, touchEnabled, commerceAllowed, minorVersion, requirementMatrix, hostCustomMessage, hostFullVersion, undefined, undefined, undefined, undefined, undefined, undefined, undefined, initialDisplayMode, undefined, undefined);
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
                if (this._hostInfo.hostPlatform === OSF.HostInfoPlatform.mac) {
                    this._clientHostController = new OSF.MacRichClientHostController(OSF.ScriptMessaging.GetScriptMessenger());
                }
                else {
                    this._clientHostController = new OSF.Webkit.WebkitHostController(OSF.ScriptMessaging.GetScriptMessenger());
                }
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
    var WebViewClientSettingsManager = (function () {
        function WebViewClientSettingsManager(_initializationHelper, _scriptMessager) {
            this._initializationHelper = _initializationHelper;
            this._scriptMessager = _scriptMessager;
        }
        WebViewClientSettingsManager.prototype.read = function (onComplete) {
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
        WebViewClientSettingsManager.prototype.write = function (serializedSettings, onComplete) {
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
            this._scriptMessager.invokeMethod(OSF.WebView.MessageHandlerName, OSF.WebView.MethodId.WriteSettings, hostParams, onWriteCompleted);
        };
        return WebViewClientSettingsManager;
    }());
    OSF.WebViewClientSettingsManager = WebViewClientSettingsManager;
})(OSF || (OSF = {}));
var OSF;
(function (OSF) {
    var WebView;
    (function (WebView) {
        WebView.MessageHandlerName = "Agave";
        WebView.PopupMessageHandlerName = "WefPopupHandler";
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
            AppContextProperties[AppContextProperties["OfficeThemeInfo"] = 21] = "OfficeThemeInfo";
        })(AppContextProperties = WebView.AppContextProperties || (WebView.AppContextProperties = {}));
        var MethodId;
        (function (MethodId) {
            MethodId[MethodId["Execute"] = 1] = "Execute";
            MethodId[MethodId["RegisterEvent"] = 2] = "RegisterEvent";
            MethodId[MethodId["UnregisterEvent"] = 3] = "UnregisterEvent";
            MethodId[MethodId["WriteSettings"] = 4] = "WriteSettings";
            MethodId[MethodId["GetContext"] = 5] = "GetContext";
            MethodId[MethodId["OnKeydown"] = 6] = "OnKeydown";
            MethodId[MethodId["AddinInitialized"] = 7] = "AddinInitialized";
            MethodId[MethodId["OpenWindow"] = 8] = "OpenWindow";
            MethodId[MethodId["MessageParent"] = 9] = "MessageParent";
            MethodId[MethodId["SendMessage"] = 10] = "SendMessage";
        })(MethodId = WebView.MethodId || (WebView.MethodId = {}));
        var WebViewHostController = (function () {
            function WebViewHostController(hostScriptProxy) {
                this.hostScriptProxy = hostScriptProxy;
            }
            WebViewHostController.prototype.execute = function (id, params, callback) {
                var args = params;
                if (args == null) {
                    args = [];
                }
                var hostParams = {
                    id: id,
                    apiArgs: args
                };
                var agaveResponseCallback = function (payload) {
                    var safeArraySource = payload;
                    if (OSF.OUtil.isArray(payload) && payload.length >= 2) {
                        var hrStatus = payload[0];
                        safeArraySource = payload[1];
                    }
                    if (callback) {
                        return callback(new OSF.WebkitSafeArray(safeArraySource));
                    }
                };
                this.hostScriptProxy.invokeMethod(OSF.WebView.MessageHandlerName, OSF.WebView.MethodId.Execute, hostParams, agaveResponseCallback);
            };
            WebViewHostController.prototype.registerEvent = function (id, eventType, targetId, handler, callback) {
                var agaveEventHandlerCallback = function (payload) {
                    var safeArraySource = payload;
                    var eventId = 0;
                    if (OSF.OUtil.isArray(payload) && payload.length >= 2) {
                        eventId = payload[0];
                        safeArraySource = payload[1];
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
                this.hostScriptProxy.registerEvent(OSF.WebView.MessageHandlerName, OSF.WebView.MethodId.RegisterEvent, id, targetId, agaveEventHandlerCallback, agaveResponseCallback);
            };
            WebViewHostController.prototype.unregisterEvent = function (id, eventType, targetId, callback) {
                var agaveResponseCallback = function (response) {
                    return callback(new OSF.WebkitSafeArray(response));
                };
                this.hostScriptProxy.unregisterEvent(OSF.WebView.MessageHandlerName, OSF.WebView.MethodId.UnregisterEvent, id, targetId, agaveResponseCallback);
            };
            WebViewHostController.prototype.messageParent = function (params) {
                var message = params[OSF.ParameterNames.MessageToParent];
                var messageObj = { dialogMessage: { messageType: 0, messageContent: message } };
                window.opener.postMessage(JSON.stringify(messageObj), window.location.origin);
            };
            WebViewHostController.prototype.openDialog = function (id, eventType, targetId, handler, callback) {
                var magicWord = "action=displayDialog";
                var callArgs = JSON.parse(targetId);
                var callUrl = callArgs.url;
                if (!callUrl) {
                    return;
                }
                var fragmentSeparator = '#';
                var urlParts = callUrl.split(fragmentSeparator);
                var seperator = "?";
                if (callUrl.indexOf("?") > -1) {
                    seperator = "&";
                }
                var width = screen.width * callArgs.width / 100;
                var height = screen.height * callArgs.height / 100;
                var params = "width=" + width + ", height=" + height;
                urlParts[0] = urlParts[0].concat(seperator).concat(magicWord);
                var openUrl = urlParts.join(fragmentSeparator);
                WebViewHostController.popup = window.open(openUrl, "", params);
                function receiveMessage(event) {
                    if (event.source == WebViewHostController.popup) {
                        try {
                            var messageObj = JSON.parse(event.data);
                            if (messageObj.dialogMessage) {
                                handler(id, [0, messageObj.dialogMessage.messageContent]);
                            }
                        }
                        catch (e) {
                            OSF.Utility.trace("messages received cannot be handled. Message:" + event.data);
                        }
                    }
                }
                window.addEventListener("message", receiveMessage);
                var interval;
                function checkWindowClose() {
                    try {
                        if (WebViewHostController.popup == null || WebViewHostController.popup.closed) {
                            window.clearInterval(interval);
                            handler(id, [12006]);
                        }
                    }
                    catch (e) {
                        OSF.Utility.trace("Error happened when popup window closed.");
                    }
                }
                interval = window.setInterval(checkWindowClose, 1000);
                callback(0);
            };
            WebViewHostController.prototype.closeDialog = function (id, eventType, targetId, callback) {
                if (WebViewHostController.popup) {
                    WebViewHostController.popup.close();
                    WebViewHostController.popup = null;
                    callback(0);
                }
                else {
                    callback(5001);
                }
            };
            WebViewHostController.prototype.sendMessage = function (params) {
                var message = params[OSF.ParameterNames.MessageContent];
                if (!isNaN(parseFloat(message)) && isFinite(message)) {
                    message = message.toString();
                }
                this.hostScriptProxy.invokeMethod(OSF.WebView.MessageHandlerName, OSF.WebView.MethodId.SendMessage, message, null);
            };
            return WebViewHostController;
        }());
        WebView.WebViewHostController = WebViewHostController;
    })(WebView = OSF.WebView || (OSF.WebView = {}));
})(OSF || (OSF = {}));
var OSF;
(function (OSF) {
    var WebViewInitializationHelper = (function (_super) {
        __extends(WebViewInitializationHelper, _super);
        function WebViewInitializationHelper(hostInfo, webAppState, context, hostFacade) {
            var _this = _super.call(this, hostInfo, webAppState, context, hostFacade) || this;
            _this.initializeWebViewMessaging();
            return _this;
        }
        WebViewInitializationHelper.prototype.initializeWebViewMessaging = function () {
            OSF.ScriptMessaging = OSF.WebView.ScriptMessaging;
        };
        WebViewInitializationHelper.prototype.getAppContext = function (wnd, onSuccess, onError) {
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
                returnedContext = new OSF.OfficeAppContext(id, appType, version, UILocale, dataLocale, docUrl, clientMode, settingsFunc, reason, osfControlType, eToken, correlationId, appInstanceId, touchEnabled, commerceAllowed, minorVersion, requirementMatrix, hostCustomMessage, hostFullVersion, undefined, undefined, undefined, undefined, undefined, undefined, undefined, initialDisplayMode, undefined, undefined);
                onSuccess(returnedContext);
            };
            var handler;
            if (this._hostInfo.isDialog) {
                handler = OSF.WebView.PopupMessageHandlerName;
            }
            else {
                handler = OSF.WebView.MessageHandlerName;
            }
            OSF.ScriptMessaging.GetScriptMessenger().invokeMethod(handler, OSF.WebView.MethodId.GetContext, [], getInvocationCallback);
        };
        WebViewInitializationHelper.prototype.createClientHostController = function () {
            if (!this._clientHostController) {
                this._clientHostController = new OSF.WebView.WebViewHostController(OSF.ScriptMessaging.GetScriptMessenger());
            }
            return this._clientHostController;
        };
        WebViewInitializationHelper.prototype.createAsyncMethodExecutor = function () {
            return new OSF.SafeArrayAsyncMethodExecutor(this.createClientHostController());
        };
        WebViewInitializationHelper.prototype.createClientSettingsManager = function () {
            return new OSF.WebViewClientSettingsManager(this, OSF.ScriptMessaging.GetScriptMessenger());
        };
        return WebViewInitializationHelper;
    }(OSF.InitializationHelper));
    OSF.WebViewInitializationHelper = WebViewInitializationHelper;
})(OSF || (OSF = {}));
var OSF;
(function (OSF) {
    var CrossIFrameCommon;
    (function (CrossIFrameCommon) {
        var CallbackType;
        (function (CallbackType) {
            CallbackType[CallbackType["MethodCallback"] = 0] = "MethodCallback";
            CallbackType[CallbackType["EventCallback"] = 1] = "EventCallback";
        })(CallbackType = CrossIFrameCommon.CallbackType || (CrossIFrameCommon.CallbackType = {}));
        var CallbackData = (function () {
            function CallbackData(callbackType, callbackId, params) {
                this.callbackType = callbackType;
                this.callbackId = callbackId;
                this.params = params;
            }
            return CallbackData;
        }());
        CrossIFrameCommon.CallbackData = CallbackData;
    })(CrossIFrameCommon || (CrossIFrameCommon = {}));
    var Android;
    (function (Android) {
        var Poster = (function () {
            function Poster() {
            }
            Poster.getInstance = function () {
                if (Poster.uniqueInstance == null) {
                    Poster.uniqueInstance = new Poster();
                }
                return Poster.uniqueInstance;
            };
            Poster.prototype.postMessage = function (handlerName, message) {
                agaveHost.postMessage(message);
            };
            Poster.prototype.ReceiveMessage = function (cbData) {
                switch (cbData.callbackType) {
                    case CrossIFrameCommon.CallbackType.MethodCallback:
                        OSF.WebView.ScriptMessaging.agaveHostCallback(cbData.callbackId, cbData.params);
                        break;
                    case CrossIFrameCommon.CallbackType.EventCallback:
                        OSF.WebView.ScriptMessaging.agaveHostEventCallback(cbData.callbackId, cbData.params);
                        break;
                    default:
                        break;
                }
            };
            return Poster;
        }());
        Android.Poster = Poster;
    })(Android = OSF.Android || (OSF.Android = {}));
    var WinRT;
    (function (WinRT) {
        var Poster = (function () {
            function Poster() {
                window.addEventListener("message", this.OnReceiveMessage.bind(this));
            }
            Poster.prototype.postMessage = function (handlerName, message) {
                window.parent.postMessage(message, "*");
            };
            Poster.prototype.OnReceiveMessage = function (event) {
                if (event.source != window.parent || window.parent != window.top || !event.origin.startsWith("ms-appx-web://")) {
                    return;
                }
                var cbData;
                try {
                    cbData = JSON.parse(event.data);
                }
                catch (ex) {
                    return;
                }
                switch (cbData.callbackType) {
                    case CrossIFrameCommon.CallbackType.MethodCallback:
                        OSF.WebView.ScriptMessaging.agaveHostCallback(cbData.callbackId, JSON.parse(cbData.params));
                        break;
                    case CrossIFrameCommon.CallbackType.EventCallback:
                        OSF.WebView.ScriptMessaging.agaveHostEventCallback(cbData.callbackId, JSON.parse(cbData.params));
                        break;
                    default:
                        break;
                }
            };
            return Poster;
        }());
        WinRT.Poster = Poster;
    })(WinRT = OSF.WinRT || (OSF.WinRT = {}));
})(OSF || (OSF = {}));
(function (OSF) {
    var WebView;
    (function (WebView) {
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
                    var hostPlatform = OSF._OfficeAppFactory.getHostInfo().hostPlatform;
                    if (hostPlatform === OSF.HostInfoPlatform.android) {
                        scriptMessenger = new WebViewScriptMessaging("OSF.ScriptMessaging.agaveHostCallback", "OSF.ScriptMessaging.agaveHostEventCallback", OSF.Android.Poster.getInstance());
                    }
                    else if (hostPlatform === OSF.HostInfoPlatform.winrt) {
                        scriptMessenger = new WebViewScriptMessaging("agaveHostCallback", "agaveHostEventCallback", new OSF.WinRT.Poster());
                    }
                    else {
                        throw OSF.Utility.createNotImplementedException();
                    }
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
            var WebViewScriptMessaging = (function () {
                function WebViewScriptMessaging(methodCallbackName, eventCallbackName, messagePoster) {
                    this.callingIndex = 0;
                    this.callbackList = {};
                    this.eventHandlerList = {};
                    this.asyncMethodCallbackFunctionName = methodCallbackName;
                    this.eventCallbackFunctionName = eventCallbackName;
                    this.poster = messagePoster;
                    this.conversationId = WebViewScriptMessaging.getCurrentTimeMS().toString();
                }
                WebViewScriptMessaging.prototype.invokeMethod = function (handlerName, methodId, params, callback) {
                    var messagingArgs = {};
                    this.postMessage(messagingArgs, handlerName, methodId, params, callback);
                };
                WebViewScriptMessaging.prototype.registerEvent = function (handlerName, methodId, dispId, targetId, handler, callback) {
                    var messagingArgs = {
                        eventCallbackFunction: this.eventCallbackFunctionName
                    };
                    var hostArgs = {
                        id: dispId,
                        targetId: targetId
                    };
                    var correlationId = this.postMessage(messagingArgs, handlerName, methodId, hostArgs, callback);
                    this.eventHandlerList[correlationId] = new EventHandlerCallback(dispId, targetId, handler);
                };
                WebViewScriptMessaging.prototype.unregisterEvent = function (handlerName, methodId, dispId, targetId, callback) {
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
                WebViewScriptMessaging.prototype.agaveHostCallback = function (callbackId, params) {
                    var callbackFunction = this.callbackList[callbackId];
                    if (callbackFunction) {
                        var callbacksDone = callbackFunction(params);
                        if (callbacksDone === undefined || callbacksDone === true) {
                            delete this.callbackList[callbackId];
                        }
                    }
                };
                WebViewScriptMessaging.prototype.agaveHostEventCallback = function (callbackId, params) {
                    var eventCallback = this.eventHandlerList[callbackId];
                    if (eventCallback) {
                        eventCallback.handler(params);
                    }
                };
                WebViewScriptMessaging.prototype.postMessage = function (messagingArgs, handlerName, methodId, params, callback) {
                    var correlationId = this.generateCorrelationId();
                    this.callbackList[correlationId] = callback;
                    messagingArgs.methodId = methodId;
                    messagingArgs.params = params;
                    messagingArgs.callbackId = correlationId;
                    messagingArgs.callbackFunction = this.asyncMethodCallbackFunctionName;
                    this.poster.postMessage(handlerName, JSON.stringify(messagingArgs));
                    return correlationId;
                };
                WebViewScriptMessaging.prototype.generateCorrelationId = function () {
                    ++this.callingIndex;
                    return this.conversationId + this.callingIndex;
                };
                WebViewScriptMessaging.getCurrentTimeMS = function () {
                    return (new Date).getTime();
                };
                WebViewScriptMessaging.MESSAGE_TIME_DELTA = 10;
                return WebViewScriptMessaging;
            }());
        })(ScriptMessaging = WebView.ScriptMessaging || (WebView.ScriptMessaging = {}));
    })(WebView = OSF.WebView || (OSF.WebView = {}));
})(OSF || (OSF = {}));
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
var OSF;
(function (OSF) {
    var MacRichClientHostController = (function (_super) {
        __extends(MacRichClientHostController, _super);
        function MacRichClientHostController() {
            return _super !== null && _super.apply(this, arguments) || this;
        }
        MacRichClientHostController.prototype.openDialog = function (id, eventType, targetId, handler, callback) {
            if (MacRichClientHostController.popup && !MacRichClientHostController.popup.closed) {
                callback(12007);
                return;
            }
            var magicWord = "action=displayDialog";
            window.dialogAPIErrorCode = undefined;
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
            MacRichClientHostController.popup = window.open(openUrl, "", params);
            function receiveMessage(event) {
                if (event.source == MacRichClientHostController.popup) {
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
            MacRichClientHostController.DialogEventListener = receiveMessage;
            function checkWindowCloseNotifyError(errorCode) {
                handler(id, [errorCode]);
            }
            function checkWindowClose() {
                try {
                    if (MacRichClientHostController.popup == null || MacRichClientHostController.popup.closed) {
                        window.clearInterval(MacRichClientHostController.interval);
                        window.removeEventListener("message", MacRichClientHostController.DialogEventListener);
                        MacRichClientHostController.NotifyError = null;
                        handler(id, [12006]);
                    }
                }
                catch (e) {
                    OSF.Utility.trace("Error happened when popup window closed.");
                }
            }
            if (MacRichClientHostController.popup != undefined && window.dialogAPIErrorCode == undefined) {
                window.addEventListener("message", MacRichClientHostController.DialogEventListener);
                MacRichClientHostController.interval = window.setInterval(checkWindowClose, 500);
                MacRichClientHostController.NotifyError = checkWindowCloseNotifyError;
                callback(0);
            }
            else {
                var error = 5001;
                if (window.dialogAPIErrorCode) {
                    error = window.dialogAPIErrorCode;
                }
                callback(error);
            }
        };
        MacRichClientHostController.prototype.messageParent = function (params) {
            var message = params[OSF.ParameterNames.MessageToParent];
            var messageObj = { dialogMessage: { messageType: 0, messageContent: message } };
            window.opener.postMessage(JSON.stringify(messageObj), window.location.origin);
        };
        MacRichClientHostController.prototype.closeDialog = function (id, eventType, targetId, callback) {
            if (MacRichClientHostController.popup) {
                if (MacRichClientHostController.interval) {
                    window.clearInterval(MacRichClientHostController.interval);
                }
                MacRichClientHostController.popup.close();
                MacRichClientHostController.popup = null;
                window.removeEventListener("message", MacRichClientHostController.DialogEventListener);
                MacRichClientHostController.NotifyError = null;
                callback(0);
            }
            else {
                callback(5001);
            }
        };
        return MacRichClientHostController;
    }(OSF.Webkit.WebkitHostController));
    OSF.MacRichClientHostController = MacRichClientHostController;
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
        var isAppActivatedPending = false;
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
            var isFromWacAutomation = context.get_isFromWacAutomation();
            if (isFromWacAutomation !== undefined && isFromWacAutomation !== null) {
                appInfo.isFromWacAutomation = isFromWacAutomation.toString().toLowerCase();
            }
            var docUrl = context.get_docUrl();
            appInfo.docUrl = omexDomainRegex.test(docUrl) ? docUrl : "";
            var url = location.href;
            if (url) {
                appInfo.isPreload = url.indexOf('preload=1') !== -1 ? true : false;
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
            OTel.OTelLogger.initialize(appInfo);
            if (appInfo.isPreload) {
                isAppActivatedPending = true;
            }
            else {
                AppTelemetry.onAppActivated();
            }
        }
        AppTelemetry.initialize = initialize;
        function onAppActivated() {
            if (!appInfo) {
                return;
            }
            if (isAppActivatedPending) {
                return;
            }
            OTel.OTelLogger.onTelemetryLoaded(function () {
                var dataFields = [];
                dataFields = dataFields.concat([
                    oteljs.makeStringDataField("Browser", appInfo.browser),
                    oteljs.makeStringDataField("AppURL", appInfo.appURL),
                    oteljs.makeInt64DataField("AppSizeWidth", window.innerWidth),
                    oteljs.makeInt64DataField("AppSizeHeight", window.innerHeight)
                ]);
                Microsoft.Office.WebExtension.sendTelemetryEvent({
                    eventName: "Office.Extensibility.OfficeJs.AppActivatedX",
                    dataFields: dataFields,
                    eventFlags: {
                        dataCategories: oteljs.DataCategories.ProductServiceUsage,
                        diagnosticLevel: oteljs.DiagnosticLevel.NecessaryServiceDataEvent,
                        samplingPolicy: oteljs.SamplingPolicy.CriticalBusinessImpact
                    }
                });
            });
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
        function onCheckWACHost(isWacKnownHost, instanceId, hostType, hostPlatform, wacDomain) {
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
        function CallOnAppActivatedIfPending() {
            if (isAppActivatedPending) {
                isAppActivatedPending = false;
                onAppActivated();
            }
        }
        AppTelemetry.CallOnAppActivatedIfPending = CallOnAppActivatedIfPending;
        function canSendAddinId() {
            var isPublic = (OSF._OfficeAppFactory.getHostInfo().flags & 16) != 0;
            if (isPublic) {
                return isPublic;
            }
            if (!appInfo) {
                return false;
            }
            var hostPlatform = OSF._OfficeAppFactory.getHostInfo().hostPlatform;
            var hostVersion = appInfo.hostVersion;
            return _isComplianceExceptedHost(hostPlatform, hostVersion);
        }
        AppTelemetry.canSendAddinId = canSendAddinId;
        function _isComplianceExceptedHost(hostPlatform, hostVersion) {
            var excepted = false;
            var versionExtractor = /^(\d+)\.(\d+)\.(\d+)\.(\d+)$/;
            var result = versionExtractor.exec(hostVersion);
            if (result) {
                var major = parseInt(result[1]);
                var minor = parseInt(result[2]);
                var build = parseInt(result[3]);
                if (hostPlatform == OSF.HostInfoPlatform.win32) {
                    if (major < 16 || major == 16 && build < 14225) {
                        excepted = true;
                    }
                }
                else if (hostPlatform == OSF.HostInfoPlatform.mac) {
                    if (major < 16 || major == 16 && build < 21062700) {
                        excepted = true;
                    }
                }
                else if (hostPlatform == OSF.HostInfoPlatform.ios) {
                    if (minor < 2122) {
                        excepted = true;
                    }
                }
                else if (hostPlatform == OSF.HostInfoPlatform.android) {
                    if (minor < 2120) {
                        excepted = true;
                    }
                }
            }
            return excepted;
        }
        AppTelemetry._isComplianceExceptedHost = _isComplianceExceptedHost;
    })(AppTelemetry = OSF.AppTelemetry || (OSF.AppTelemetry = {}));
})(OSF || (OSF = {}));
var OTel;
(function (OTel) {
    var OTelLogger = (function () {
        function OTelLogger() {
        }
        OTelLogger.loaded = function () {
            return !(OTelLogger.logger === undefined);
        };
        OTelLogger.create = function (info) {
            var contract = {
                id: info.appId,
                assetId: info.assetId,
                officeJsVersion: info.officeJSVersion,
                hostJsVersion: info.hostJSVersion,
                browserToken: info.clientId,
                instanceId: info.appInstanceId,
                sessionId: info.sessionId
            };
            var fields = oteljs.Contracts.Office.System.SDX.getFields("SDX", contract);
            var flavor = OSF._OfficeAppFactory.getHostInfo().hostPlatform;
            var sink;
            if (flavor === "web") {
                sink = new OTel.SdxWacSink();
            }
            else if (Office.context.requirements.isSetSupported('Telemetry', '1.2')) {
                sink = new OTel.RichApiSink();
            }
            else {
                console.error('Cannot create telemetry sink successfully');
                return null;
            }
            var namespace = "Office.Extensibility.OfficeJs";
            var ariaTenantToken = 'db334b301e7b474db5e0f02f07c51a47-a1b5bc36-1bbe-482f-a64a-c2d9cb606706-7439';
            var nexusTenantToken = 1755;
            var logger = new oteljs.SimpleTelemetryLogger(undefined, fields);
            logger.addSink(sink);
            logger.setTenantToken(namespace, ariaTenantToken, nexusTenantToken);
            oteljs.onNotification().addListener(function (notification) {
                OSF.Utility.debugLog(notification.message());
            });
            return logger;
        };
        OTelLogger.checkAndResolvePromises = function () {
            if (OTelLogger.loaded()) {
                OTelLogger.promises.forEach(function (resolve) {
                    resolve();
                });
                OTelLogger.promises = [];
            }
        };
        OTelLogger.initialize = function (info) {
            if (!OTelLogger.Enabled) {
                OTelLogger.promises = [];
                return;
            }
            Office.onReadyInternal().then(function () {
                if (!OTelLogger.loaded()) {
                    OSF.Utility.debugLog("Creating OTelLogger");
                    OTelLogger.logger = OTelLogger.create(info);
                    OTelLogger.checkAndResolvePromises();
                }
            });
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
var OTel;
(function (OTel) {
    var DEFAULT_MINIMUM_MILLISECONDS_BETWEEN_CALLS = 1000;
    var _minimumMillisecondsBeforeFirstCall = DEFAULT_MINIMUM_MILLISECONDS_BETWEEN_CALLS;
    var _minimumMillisecondsBetweenCalls = DEFAULT_MINIMUM_MILLISECONDS_BETWEEN_CALLS;
    var RichApiSink = (function () {
        function RichApiSink() {
            var _this = this;
            this._requestIsPending = true;
            this._telemetryQueue = [];
            this.pause(_minimumMillisecondsBeforeFirstCall).then(function () {
                var currentWork = _this._telemetryQueue;
                _this._telemetryQueue = [];
                _this._requestIsPending = false;
                _this.processTelemetryEvents(currentWork);
            });
        }
        RichApiSink.prototype.sendTelemetryEvent = function (telemetryEvent) {
            this._telemetryQueue.push(telemetryEvent);
            if (this._requestIsPending) {
                return;
            }
            this.processWorkBacklog();
        };
        RichApiSink.prototype.processWorkBacklog = function () {
            var _this = this;
            this._requestIsPending = true;
            var currentWork = this._telemetryQueue;
            this._telemetryQueue = [];
            this.processTelemetryEvents(currentWork)
                .then(this.waitAndProcessMore)
                .catch(function (error) {
                oteljs.logError(oteljs.Category.Sink, "RichApiSink Error", error);
                _this.waitAndProcessMore();
            });
        };
        RichApiSink.prototype.waitAndProcessMore = function () {
            var _this = this;
            this.pause(_minimumMillisecondsBetweenCalls)
                .then(function () {
                if (_this._telemetryQueue.length > 0) {
                    setTimeout(function () { return _this.processWorkBacklog(); }, 0);
                }
                _this._requestIsPending = false;
            })
                .catch(function () { });
        };
        RichApiSink.prototype.processTelemetryEvents = function (telemetryEvents) {
            var _this = this;
            var ctx = new OfficeCore.RequestContext();
            telemetryEvents.forEach(function (telemetryEvent) {
                if (!telemetryEvent.telemetryProperties) {
                    return;
                }
                var dataFields = [];
                _this.addDataFields(dataFields, telemetryEvent.dataFields);
                var contractName = !!telemetryEvent.eventContract ? telemetryEvent.eventContract.name : '';
                if (!!telemetryEvent.eventContract) {
                    _this.addDataFields(dataFields, telemetryEvent.eventContract.dataFields);
                }
                ctx.telemetry.sendTelemetryEvent(telemetryEvent.telemetryProperties, telemetryEvent.eventName, contractName, oteljs.getEffectiveEventFlags(telemetryEvent), dataFields);
            });
            return ctx.sync().catch(function () {
                oteljs.logNotification(oteljs.LogLevel.Info, oteljs.Category.Sink, function () { return 'RichApi telemetry call failed.'; });
            });
        };
        RichApiSink.prototype.addDataFields = function (richApiDataFields, dataFields) {
            if (dataFields) {
                dataFields.forEach(function (dataField) {
                    richApiDataFields.push({
                        name: dataField.name,
                        value: dataField.value,
                        classification: dataField.classification ? dataField.classification : oteljs.DataClassification.SystemMetadata,
                        type: dataField.dataType
                    });
                });
            }
        };
        RichApiSink.prototype.pause = function (ms) {
            return new Office.Promise(function (resolve) { return setTimeout(resolve, ms); });
        };
        return RichApiSink;
    }());
    OTel.RichApiSink = RichApiSink;
})(OTel || (OTel = {}));
var OTel;
(function (OTel) {
    var SdxWacSink = (function () {
        function SdxWacSink() {
        }
        SdxWacSink.prototype.sendTelemetryEvent = function (event, _timestamp) {
            try {
                if (event.dataFields &&
                    event.dataFields.filter(function (dataField) {
                        return dataField.classification && dataField.classification !== oteljs.DataClassification.SystemMetadata;
                    }).length > 0) {
                    return;
                }
                var id = OSF._OfficeAppFactory.getId();
                var SendTelemetryEventId = OSF.AgaveHostAction.SendTelemetryEvent;
                OSF.getClientEndPoint().invoke('ContextActivationManager_notifyHost', null, [id, SendTelemetryEventId, event]);
            }
            catch (error) {
                oteljs.logError(oteljs.Category.Sink, "AgaveWacSink", error);
            }
        };
        return SdxWacSink;
    }());
    OTel.SdxWacSink = SdxWacSink;
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
                return domain.split(".").slice(-2).join(".");
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
    var EventDispId;
    (function (EventDispId) {
        EventDispId.dispidOlkItemSelectedChangedEvent = 46;
        EventDispId.dispidOlkRecipientsChangedEvent = 47;
        EventDispId.dispidOlkAppointmentTimeChangedEvent = 48;
        EventDispId.dispidOlkRecurrenceChangedEvent = 49;
        EventDispId.dispidOlkAttachmentsChangedEvent = 50;
        EventDispId.dispidOlkEnhancedLocationsChangedEvent = 51;
        EventDispId.dispidOlkInfobarClickedEvent = 52;
    })(EventDispId = OSF.EventDispId || (OSF.EventDispId = {}));
})(OSF || (OSF = {}));
var OSF;
(function (OSF) {
    var EventType;
    (function (EventType) {
        EventType.ItemChanged = "olkItemSelectedChanged";
        EventType.RecipientsChanged = "olkRecipientsChanged";
        EventType.AppointmentTimeChanged = "olkAppointmentTimeChanged";
        EventType.AttachmentsChanged = "olkAttachmentsChanged";
        EventType.EnhancedLocationsChanged = "olkEnhancedLocationsChanged";
        EventType.InfobarClicked = "olkInfobarClicked";
        EventType.RecurrenceChanged = "olkRecurrenceChanged";
        EventType.OfficeThemeChanged = "officeThemeChanged";
    })(EventType = OSF.EventType || (OSF.EventType = {}));
})(OSF || (OSF = {}));
var Microsoft;
(function (Microsoft) {
    var Office;
    (function (Office) {
        var WebExtension;
        (function (WebExtension) {
            var OutlookBase;
            (function (OutlookBase) {
            })(OutlookBase = WebExtension.OutlookBase || (WebExtension.OutlookBase = {}));
        })(WebExtension = Office.WebExtension || (Office.WebExtension = {}));
    })(Office = Microsoft.Office || (Microsoft.Office = {}));
})(Microsoft || (Microsoft = {}));
var Office;
(function (Office) {
    var context;
    (function (context) {
        var _mailbox;
        OSF.definePropertyOnNamespace(context, 'mailbox', function () {
            if (!_mailbox) {
                _mailbox = OSF.OutlookInitializeManager.getMailboxObject();
            }
            return _mailbox;
        });
        var _roamingSettings;
        function get_roamingSettings() {
            if (!_roamingSettings) {
                var serializedSettings = OSF._OfficeAppFactory.getOfficeAppContext().get_settingsFunc()();
                var deserializedSettings = OSF.OUtil.deserializeSettings(serializedSettings);
                _roamingSettings = new OSF.DDA.Settings(deserializedSettings);
                OSF.DDA.ClientSettingsManager = OSF._OfficeAppFactory.getInitializationHelper().createClientSettingsManager();
            }
            return _roamingSettings;
        }
        OSF.definePropertyOnNamespace(context, 'roamingSettings', get_roamingSettings);
    })(context = Office.context || (Office.context = {}));
})(Office || (Office = {}));
var Office;
(function (Office) {
    var CoercionType;
    (function (CoercionType) {
        CoercionType["Html"] = "html";
        CoercionType["Text"] = "text";
    })(CoercionType = Office.CoercionType || (Office.CoercionType = {}));
    var MailboxEnums;
    (function (MailboxEnums) {
        var EntityType;
        (function (EntityType) {
            EntityType["MeetingSuggestion"] = "meetingSuggestion";
            EntityType["TaskSuggestion"] = "taskSuggestion";
            EntityType["Address"] = "address";
            EntityType["EmailAddress"] = "emailAddress";
            EntityType["Url"] = "url";
            EntityType["PhoneNumber"] = "phoneNumber";
            EntityType["Contact"] = "contact";
            EntityType["FlightReservations"] = "flightReservations";
            EntityType["ParcelDeliveries"] = "parcelDeliveries";
        })(EntityType = MailboxEnums.EntityType || (MailboxEnums.EntityType = {}));
        var ItemType;
        (function (ItemType) {
            ItemType["Message"] = "message";
            ItemType["Appointment"] = "appointment";
        })(ItemType = MailboxEnums.ItemType || (MailboxEnums.ItemType = {}));
        var ResponseType;
        (function (ResponseType) {
            ResponseType["None"] = "none";
            ResponseType["Organizer"] = "organizer";
            ResponseType["Tentative"] = "tentative";
            ResponseType["Accepted"] = "accepted";
            ResponseType["Declined"] = "declined";
        })(ResponseType = MailboxEnums.ResponseType || (MailboxEnums.ResponseType = {}));
        var RecipientType;
        (function (RecipientType) {
            RecipientType["Other"] = "other";
            RecipientType["DistributionList"] = "distributionList";
            RecipientType["User"] = "user";
            RecipientType["ExternalUser"] = "externalUser";
        })(RecipientType = MailboxEnums.RecipientType || (MailboxEnums.RecipientType = {}));
        var AttachmentType;
        (function (AttachmentType) {
            AttachmentType["File"] = "file";
            AttachmentType["Item"] = "item";
            AttachmentType["Cloud"] = "cloud";
        })(AttachmentType = MailboxEnums.AttachmentType || (MailboxEnums.AttachmentType = {}));
        var BodyType;
        (function (BodyType) {
            BodyType["Text"] = "text";
            BodyType["Html"] = "html";
        })(BodyType = MailboxEnums.BodyType || (MailboxEnums.BodyType = {}));
        var ItemNotificationMessageType;
        (function (ItemNotificationMessageType) {
            ItemNotificationMessageType["ProgressIndicator"] = "progressIndicator";
            ItemNotificationMessageType["InformationalMessage"] = "informationalMessage";
            ItemNotificationMessageType["ErrorMessage"] = "errorMessage";
            ItemNotificationMessageType["InsightMessage"] = "insightMessage";
        })(ItemNotificationMessageType = MailboxEnums.ItemNotificationMessageType || (MailboxEnums.ItemNotificationMessageType = {}));
        var Folder;
        (function (Folder) {
            Folder["Inbox"] = "inbox";
            Folder["Junk"] = "junk";
            Folder["DeletedItems"] = "deletedItems";
        })(Folder = MailboxEnums.Folder || (MailboxEnums.Folder = {}));
        var UserProfileType;
        (function (UserProfileType) {
            UserProfileType["Office365"] = "office365";
            UserProfileType["OutlookCom"] = "outlookCom";
            UserProfileType["Enterprise"] = "enterprise";
        })(UserProfileType = MailboxEnums.UserProfileType || (MailboxEnums.UserProfileType = {}));
        var RestVersion;
        (function (RestVersion) {
            RestVersion["v1_0"] = "v1.0";
            RestVersion["v2_0"] = "v2.0";
            RestVersion["Beta"] = "beta";
        })(RestVersion = MailboxEnums.RestVersion || (MailboxEnums.RestVersion = {}));
        var ModuleType;
        (function (ModuleType) {
            ModuleType["Addins"] = "addins";
        })(ModuleType = MailboxEnums.ModuleType || (MailboxEnums.ModuleType = {}));
        var ActionType;
        (function (ActionType) {
            ActionType["ShowTaskPane"] = "showTaskPane";
        })(ActionType = MailboxEnums.ActionType || (MailboxEnums.ActionType = {}));
    })(MailboxEnums = Office.MailboxEnums || (Office.MailboxEnums = {}));
    var EventType;
    (function (EventType) {
        EventType["ItemChanged"] = "olkItemSelectedChanged";
        EventType["RecipientsChanged"] = "olkRecipientsChanged";
        EventType["AppointmentTimeChanged"] = "olkAppointmentTimeChanged";
        EventType["AttachmentsChanged"] = "olkAttachmentsChanged";
        EventType["EnhancedLocationsChanged"] = "olkEnhancedLocationsChanged";
        EventType["InfobarClicked"] = "olkInfobarClicked";
        EventType["RecurrenceChanged"] = "olkRecurrenceChanged";
        EventType["OfficeThemeChanged"] = "officeThemeChanged";
    })(EventType = Office.EventType || (Office.EventType = {}));
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
var OSF;
(function (OSF) {
    var OutlookAsyncMethodExecutorHelper = (function (_super) {
        __extends(OutlookAsyncMethodExecutorHelper, _super);
        function OutlookAsyncMethodExecutorHelper() {
            return _super !== null && _super.apply(this, arguments) || this;
        }
        OutlookAsyncMethodExecutorHelper.prototype.handleSafeArrayHostResponse = function (hostResponseArgs, resultCode, chunkResultData, callback, dataTransform, id, asyncContext) {
            if (typeof callback == "function") {
                return callback(resultCode, hostResponseArgs);
            }
            return true;
        };
        OutlookAsyncMethodExecutorHelper.prototype.handleWebHostResponse = function (hostResponseArgs, resultCode, callback, dataTransform, id, asyncContext) {
            if (typeof callback == "function") {
                return callback(resultCode, hostResponseArgs);
            }
        };
        return OutlookAsyncMethodExecutorHelper;
    }(OSF.AsyncMethodExecutorHelper));
    OSF.OutlookAsyncMethodExecutorHelper = OutlookAsyncMethodExecutorHelper;
})(OSF || (OSF = {}));
var OSF;
(function (OSF) {
    var OutlookInitializationHelper;
    (function (OutlookInitializationHelper) {
        function addEventDispatchToTarget(target, eventDispatch) {
            if (!target.addHandlerAsync) {
                OSF.OUtil.defineMethodOnNamespace(target, 'addHandlerAsync', function (eventType, handler, asyncContext, callback) {
                    OSF.EventHelper.addEventHandler(eventType, handler, callback, eventDispatch, asyncContext.asyncContext);
                });
            }
            if (!target.removeHandlerAsync) {
                OSF.OUtil.defineMethodOnNamespace(target, 'removeHandlerAsync', function (eventType, asyncContext, callback) {
                    OSF.EventHelper.removeEventHandler(eventType, null, callback, eventDispatch, asyncContext.asyncContext);
                });
            }
        }
        OutlookInitializationHelper.addEventDispatchToTarget = addEventDispatchToTarget;
        function getMailboxItemEventDispatch() {
            return new OSF.EventDispatch([
                {
                    type: OSF.EventType.RecipientsChanged,
                    id: OSF.EventDispId.dispidOlkRecipientsChangedEvent,
                    getTargetId: function () { return ''; },
                    fromSafeArrayHost: function (payload) {
                        return buildRecipientsChangedEventArgs(payload).safeArrayHost;
                    },
                    fromWebHost: function (payload) {
                        return buildRecipientsChangedEventArgs(payload).webHost;
                    }
                },
                {
                    type: OSF.EventType.AppointmentTimeChanged,
                    id: OSF.EventDispId.dispidOlkAppointmentTimeChangedEvent,
                    getTargetId: function () { return ''; },
                    fromSafeArrayHost: function (payload) {
                        return buildAppointmentTimeChangedEventArgs(payload).safeArrayHost;
                    },
                    fromWebHost: function (payload) {
                        return buildAppointmentTimeChangedEventArgs(payload).webHost;
                    }
                },
                {
                    type: OSF.EventType.AttachmentsChanged,
                    id: OSF.EventDispId.dispidOlkAttachmentsChangedEvent,
                    getTargetId: function () { return ''; },
                    fromSafeArrayHost: function (payload) {
                        return buildAttachmentsChangedEventArgs(payload).safeArrayHost;
                    },
                    fromWebHost: function (payload) {
                        return buildAttachmentsChangedEventArgs(payload).webHost;
                    }
                },
                {
                    type: OSF.EventType.EnhancedLocationsChanged,
                    id: OSF.EventDispId.dispidOlkEnhancedLocationsChangedEvent,
                    getTargetId: function () { return ''; },
                    fromSafeArrayHost: function (payload) {
                        return buildEnhancedLocationsChangedEventArgs(payload).safeArrayHost;
                    },
                    fromWebHost: function (payload) {
                        return buildEnhancedLocationsChangedEventArgs(payload).webHost;
                    }
                },
                {
                    type: OSF.EventType.InfobarClicked,
                    id: OSF.EventDispId.dispidOlkInfobarClickedEvent,
                    getTargetId: function () { return ''; },
                    fromSafeArrayHost: function (payload) {
                        return {
                            type: OSF.EventType.InfobarClicked,
                            infobarDetails: payload[0]
                        };
                    },
                    fromWebHost: function (payload) {
                        return {
                            type: OSF.EventType.InfobarClicked,
                            infobarDetails: payload[0]
                        };
                    }
                },
                {
                    type: OSF.EventType.RecurrenceChanged,
                    id: OSF.EventDispId.dispidOlkRecurrenceChangedEvent,
                    getTargetId: function () { return ''; },
                    fromSafeArrayHost: function (payload) {
                        return buildRecurrenceChangedEventArgs(payload).safeArrayHost;
                    },
                    fromWebHost: function (payload) {
                        return buildRecurrenceChangedEventArgs(payload).webHost;
                    }
                }
            ]);
        }
        OutlookInitializationHelper.getMailboxItemEventDispatch = getMailboxItemEventDispatch;
        function getMailboxEventDispatch() {
            return new OSF.EventDispatch([
                {
                    type: OSF.EventType.ItemChanged,
                    id: OSF.EventDispId.dispidOlkItemSelectedChangedEvent,
                    getTargetId: function () { return ''; },
                    fromSafeArrayHost: function (payload) {
                        return buildItemChangedEventArgs(payload).safeArrayHost;
                    },
                    fromWebHost: function (payload) {
                        return buildItemChangedEventArgs(payload).webHost;
                    }
                },
                {
                    type: OSF.EventType.OfficeThemeChanged,
                    id: OSF.EventDispId.dispidOfficeThemeChangedEvent,
                    getTargetId: function () { return ''; },
                    fromSafeArrayHost: function (payload) {
                        return buildOfficeThemeChangedEventArgs(payload).safeArrayHost;
                    },
                    fromWebHost: function (payload) {
                        return buildOfficeThemeChangedEventArgs(payload).webHost;
                    }
                }
            ]);
        }
        OutlookInitializationHelper.getMailboxEventDispatch = getMailboxEventDispatch;
        function buildItemChangedEventArgs(payload) {
            var initialData = null, itemNumber = null;
            try {
                initialData = JSON.parse(payload[0]);
                itemNumber = JSON.parse(payload[1]);
            }
            catch (e) { }
            return {
                safeArrayHost: {
                    type: OSF.EventType.ItemChanged,
                    initialData: initialData,
                    itemNumber: itemNumber
                },
                webHost: {
                    type: OSF.EventType.ItemChanged,
                    initialData: payload[0]
                }
            };
        }
        function buildOfficeThemeChangedEventArgs(payload) {
            var themeData = null, themeDataHex = {};
            try {
                themeData = JSON.parse(payload[0]);
                for (var color in themeData) {
                    themeDataHex[color] = OSF.OUtil.convertIntToCssHexColor(themeData[color]);
                }
            }
            catch (e) { }
            return {
                safeArrayHost: {
                    type: OSF.EventType.OfficeThemeChanged,
                    officeTheme: themeDataHex
                },
                webHost: {
                    type: OSF.EventType.OfficeThemeChanged,
                    officeTheme: themeDataHex
                }
            };
        }
        function buildRecipientsChangedEventArgs(payload) {
            var changedRecipientFields = null;
            try {
                changedRecipientFields = JSON.parse(payload[0]);
            }
            catch (e) { }
            return {
                safeArrayHost: {
                    type: OSF.EventType.RecipientsChanged,
                    changedRecipientFields: changedRecipientFields
                },
                webHost: {
                    type: OSF.EventType.RecipientsChanged,
                    changedRecipientFields: changedRecipientFields
                }
            };
        }
        function buildAppointmentTimeChangedEventArgs(payload) {
            var start = null, end = null;
            try {
                var appointmentTime = JSON.parse(payload[0]);
                start = new Date(appointmentTime.start).toISOString();
                end = new Date(appointmentTime.end).toISOString();
            }
            catch (e) { }
            return {
                safeArrayHost: {
                    type: OSF.EventType.AppointmentTimeChanged,
                    start: start,
                    end: end
                },
                webHost: {
                    type: OSF.EventType.AppointmentTimeChanged,
                    start: start,
                    end: end
                }
            };
        }
        function buildAttachmentsChangedEventArgs(payload) {
            var attachmentStatus = null, attachmentDetails = null;
            try {
                var attachmentChangedObject = JSON.parse(payload[0]);
                attachmentStatus = attachmentChangedObject.attachmentStatus;
                attachmentDetails = Microsoft.Office.WebExtension.OutlookBase.CreateAttachmentDetails(attachmentChangedObject.attachmentDetails);
            }
            catch (e) { }
            return {
                safeArrayHost: {
                    type: OSF.EventType.AttachmentsChanged,
                    attachmentStatus: attachmentStatus,
                    attachmentDetails: attachmentDetails
                },
                webHost: {
                    type: OSF.EventType.AttachmentsChanged,
                    attachmentStatus: attachmentStatus,
                    attachmentDetails: attachmentDetails
                }
            };
        }
        function buildEnhancedLocationsChangedEventArgs(payload) {
            var enhancedLocations = null;
            try {
                var enhancedLocationsChangedObject = JSON.parse(payload[0]);
                enhancedLocations = enhancedLocationsChangedObject.enhancedLocations;
            }
            catch (e) { }
            return {
                safeArrayHost: {
                    type: OSF.EventType.EnhancedLocationsChanged,
                    enhancedLocations: enhancedLocations
                },
                webHost: {
                    type: OSF.EventType.EnhancedLocationsChanged,
                    enhancedLocations: enhancedLocations
                }
            };
        }
        function buildRecurrenceChangedEventArgs(payload) {
            var recurrenceObject = null;
            try {
                var dataObject = JSON.parse(payload[0]);
                if (dataObject.recurrence != null) {
                    recurrenceObject = JSON.parse(dataObject.recurrence);
                    recurrenceObject = Microsoft.Office.WebExtension.OutlookBase.SeriesTimeJsonConverter(recurrenceObject);
                }
            }
            catch (e) { }
            return {
                safeArrayHost: {
                    type: OSF.EventType.RecurrenceChanged,
                    recurrence: recurrenceObject
                },
                webHost: {
                    type: OSF.EventType.RecurrenceChanged,
                    recurrence: recurrenceObject
                }
            };
        }
    })(OutlookInitializationHelper = OSF.OutlookInitializationHelper || (OSF.OutlookInitializationHelper = {}));
})(OSF || (OSF = {}));
var OSF;
(function (OSF) {
    var DDA;
    (function (DDA) {
    })(DDA = OSF.DDA || (OSF.DDA = {}));
    var OutlookInitializeManager;
    (function (OutlookInitializeManager) {
        var _mailbox;
        function initializeMailboxObject(officeAppContext, wnd, appReady) {
            if (!_mailbox) {
                _mailbox = new OSF.DDA.OutlookAppOm(officeAppContext, wnd, appReady);
                OSF.OutlookInitializationHelper.addEventDispatchToTarget(_mailbox, OSF.OutlookInitializationHelper.getMailboxEventDispatch());
            }
        }
        OutlookInitializeManager.initializeMailboxObject = initializeMailboxObject;
        function getMailboxObject() {
            return _mailbox;
        }
        OutlookInitializeManager.getMailboxObject = getMailboxObject;
    })(OutlookInitializeManager = OSF.OutlookInitializeManager || (OSF.OutlookInitializeManager = {}));
})(OSF || (OSF = {}));
var OSF;
(function (OSF) {
    var OutlookWebClientHostControllerHelper = (function (_super) {
        __extends(OutlookWebClientHostControllerHelper, _super);
        function OutlookWebClientHostControllerHelper() {
            return _super !== null && _super.apply(this, arguments) || this;
        }
        OutlookWebClientHostControllerHelper.prototype.getHostCallArgs = function (id, params) {
            var hostCallArgs;
            if (id == 1) {
                hostCallArgs = {
                    MethodData: {
                        ControlId: this.getControlId(),
                        DispatchId: id
                    }
                };
            }
            else {
                hostCallArgs = {
                    ApiParams: params,
                    MethodData: {
                        ControlId: this.getControlId(),
                        DispatchId: id
                    }
                };
            }
            return hostCallArgs;
        };
        OutlookWebClientHostControllerHelper.prototype.getTargetMethodName = function (id) {
            var targetMethodName;
            if (id == 1) {
                targetMethodName = 'GetInitialData';
            }
            else {
                targetMethodName = 'ExecuteMethod';
            }
            return targetMethodName;
        };
        OutlookWebClientHostControllerHelper.prototype.parseErrorFromPayload = function (id, payload) {
            var error = 0;
            if (id != 1) {
                error = payload["error"];
            }
            return error;
        };
        return OutlookWebClientHostControllerHelper;
    }(OSF.WebClientHostControllerHelper));
    OSF.OutlookWebClientHostControllerHelper = OutlookWebClientHostControllerHelper;
})(OSF || (OSF = {}));
var OSF;
(function (OSF) {
    var DDA;
    (function (DDA) {
        var SettingsManager;
        (function (SettingsManager) {
            function serializeSettings(settingsCollection) {
                return OSF.OUtil.serializeSettings(settingsCollection);
            }
            SettingsManager.serializeSettings = serializeSettings;
            function deserializeSettings(serializedSettings) {
                return OSF.OUtil.deserializeSettings(serializedSettings);
            }
            SettingsManager.deserializeSettings = deserializeSettings;
        })(SettingsManager = DDA.SettingsManager || (DDA.SettingsManager = {}));
    })(DDA = OSF.DDA || (OSF.DDA = {}));
})(OSF || (OSF = {}));
var OSF;
(function (OSF) {
    OSF.BootStrapExtension.onGetAppContext = function (officeAppContext, wnd) {
        return OSF.OUtil.ensureOfficeStringsJs().then(function () {
            return new Promise(function (resolve, reject) {
                OSF.OutlookInitializeManager.initializeMailboxObject(officeAppContext, wnd, function () {
                    resolve();
                });
            });
        });
    };
    OSF.BootStrapExtension.createWebClientHostControllerHelper = function (webAppState, delegateVersion) {
        return new OSF.OutlookWebClientHostControllerHelper(webAppState, delegateVersion);
    };
    OSF.BootStrapExtension.createAsyncMethodExecutorHelper = function (asyncMethodExecutor) {
        return new OSF.OutlookAsyncMethodExecutorHelper(asyncMethodExecutor);
    };
})(OSF || (OSF = {}));
var OSF;
(function (OSF) {
    var ApiMethodCall = (function () {
        function ApiMethodCall(requiredParameters, supportedOptions, privateStateCallbacks, checkCallArgs, displayName) {
            this._requiredParameters = requiredParameters;
            this._supportedOptions = supportedOptions;
            this._privateStateCallbacks = privateStateCallbacks;
            this._checkCallArgs = checkCallArgs;
            this._displayName = displayName;
            this._requiredCount = requiredParameters.length;
        }
        ApiMethodCall.prototype.verifyArguments = function (params, args) {
            for (var name in params) {
                var param = params[name];
                var arg = args[name];
                if (param["enum"]) {
                    switch (typeof arg) {
                        case "string":
                            if (OSF.OUtil.listContainsValue(param["enum"], arg)) {
                                break;
                            }
                        case "undefined":
                            throw 5007;
                        default:
                            throw this.getInvalidParameterString();
                    }
                }
                if (param["types"]) {
                    if (!OSF.OUtil.listContainsValue(param["types"], typeof arg)) {
                        throw this.getInvalidParameterString();
                    }
                }
            }
        };
        ApiMethodCall.prototype.extractRequiredArguments = function (userArgs, caller, stateInfo) {
            if (userArgs.length < this._requiredCount) {
                throw OSF.Utility.createParameterException(Strings.OfficeOM.L_MissingRequiredArguments);
            }
            var requiredArgs = [];
            var index;
            for (index = 0; index < this._requiredCount; index++) {
                requiredArgs.push(userArgs[index]);
            }
            this.verifyArguments(this._requiredParameters, requiredArgs);
            var ret = {};
            for (index = 0; index < this._requiredCount; index++) {
                var param = this._requiredParameters[index];
                var arg = requiredArgs[index];
                if (param.verify) {
                    var isValid = param.verify(arg, caller, stateInfo);
                    if (!isValid) {
                        throw this.getInvalidParameterString();
                    }
                }
                ret[param.name] = arg;
            }
            return ret;
        };
        ApiMethodCall.prototype.fillOptions = function (options, requiredArgs, caller, stateInfo) {
            options = options || {};
            for (var optionName in this._supportedOptions) {
                if (!OSF.OUtil.listContainsKey(options, optionName)) {
                    var value = undefined;
                    var option = this._supportedOptions[optionName];
                    if (option.calculate && requiredArgs) {
                        value = option.calculate(requiredArgs, caller, stateInfo);
                    }
                    if (!value && option.defaultValue !== undefined) {
                        value = option.defaultValue;
                    }
                    options[optionName] = value;
                }
            }
            return options;
        };
        ApiMethodCall.prototype.constructCallArgs = function (required, options, caller, stateInfo) {
            var callArgs = {};
            for (var r in required) {
                callArgs[r] = required[r];
            }
            for (var o in options) {
                callArgs[o] = options[o];
            }
            for (var s in this._privateStateCallbacks) {
                callArgs[s] = this._privateStateCallbacks[s](caller, stateInfo);
            }
            if (this._checkCallArgs) {
                callArgs = this._checkCallArgs(callArgs, caller, stateInfo);
            }
            return callArgs;
        };
        ;
        ApiMethodCall.prototype.getInvalidParameterString = function () {
            var _this = this;
            OSF.OUtil.delayExecutionAndCache(function () {
                return OSF.OUtil.formatString(Strings.OfficeOM.L_InvalidParameters, _this._displayName);
            });
        };
        return ApiMethodCall;
    }());
    OSF.ApiMethodCall = ApiMethodCall;
})(OSF || (OSF = {}));
var OSF;
(function (OSF) {
    var AsyncMethodCall = (function () {
        function AsyncMethodCall(requiredParameters, supportedOptions, privateStateCallbacks, onSucceeded, onFailed, checkCallArgs, displayName) {
            this._requiredParameters = requiredParameters;
            this._supportedOptions = supportedOptions;
            this._privateStateCallbacks = privateStateCallbacks;
            this._onSucceeded = onSucceeded;
            this._onFailed = onFailed;
            this._displayName = displayName;
            this._checkCallArgs = checkCallArgs;
            this._requiredCount = requiredParameters.length;
            this._apiMethods = new OSF.ApiMethodCall(requiredParameters, supportedOptions, privateStateCallbacks, checkCallArgs, displayName);
        }
        AsyncMethodCall.prototype.verifyAndExtractCall = function (userArgs, caller, stateInfo) {
            var required = this._apiMethods.extractRequiredArguments(userArgs, caller, stateInfo);
            var options = this.extractOptions(userArgs, required, caller, stateInfo);
            var callArgs = this._apiMethods.constructCallArgs(required, options, caller, stateInfo);
            return callArgs;
        };
        AsyncMethodCall.prototype.processResponse = function (status, response, caller, callArgs) {
            var payload;
            if (status == 0) {
                if (this._onSucceeded) {
                    payload = this._onSucceeded(response, caller, callArgs);
                }
                else {
                    payload = response;
                }
            }
            else {
                if (this._onFailed) {
                    payload = this._onFailed(status, response);
                }
                else {
                    payload = OSF.DDA.ErrorCodeManager.getErrorArgs(status);
                }
            }
            return payload;
        };
        AsyncMethodCall.prototype.getCallArgs = function (suppliedArgs) {
            var options, parameterCallback;
            for (var i = suppliedArgs.length - 1; i >= this._requiredCount; i--) {
                var argument = suppliedArgs[i];
                switch (typeof argument) {
                    case "object":
                        options = argument;
                        break;
                    case "function":
                        parameterCallback = argument;
                        break;
                }
            }
            options = options || {};
            if (parameterCallback) {
                options[OSF.ParameterNames.Callback] = parameterCallback;
            }
            return options;
        };
        AsyncMethodCall.prototype.extractOptions = function (userArgs, requiredArgs, caller, stateInfo) {
            if (userArgs.length > this._requiredCount + 2) {
                throw OSF.Utility.createParameterException(Strings.OfficeOM.L_TooManyArguments);
            }
            var options, parameterCallback;
            for (var i = userArgs.length - 1; i >= this._requiredCount; i--) {
                var argument = userArgs[i];
                switch (typeof argument) {
                    case "object":
                        if (options) {
                            throw OSF.Utility.createParameterException(Strings.OfficeOM.L_TooManyOptionalObjects);
                        }
                        else {
                            options = argument;
                        }
                        break;
                    case "function":
                        if (parameterCallback) {
                            throw OSF.Utility.createParameterException(Strings.OfficeOM.L_TooManyOptionalFunction);
                        }
                        else {
                            parameterCallback = argument;
                        }
                        break;
                    default:
                        throw OSF.Utility.createArgumentException(Strings.OfficeOM.L_InValidOptionalArgument);
                        break;
                }
            }
            options = this._apiMethods.fillOptions(options, requiredArgs, caller, stateInfo);
            if (parameterCallback) {
                if (options[OSF.ParameterNames.Callback]) {
                    throw Strings.OfficeOM.L_RedundantCallbackSpecification;
                }
                else {
                    options[OSF.ParameterNames.Callback] = parameterCallback;
                }
            }
            this._apiMethods.verifyArguments(this._supportedOptions, options);
            return options;
        };
        return AsyncMethodCall;
    }());
    OSF.AsyncMethodCall = AsyncMethodCall;
})(OSF || (OSF = {}));
var OSF;
(function (OSF) {
    var AsyncMethodCalls;
    (function (AsyncMethodCalls) {
        var mappings = {};
        function define(callDefinition) {
            mappings[callDefinition.method] = manufacture(callDefinition);
        }
        AsyncMethodCalls.define = define;
        function get(method) {
            return mappings[method];
        }
        AsyncMethodCalls.get = get;
        function manufacture(params) {
            var supportedOptions = params.supportedOptions ? OSF.OUtil.createObject(params.supportedOptions) : [];
            var privateStateCallbacks = params.privateStateCallbacks ? OSF.OUtil.createObject(params.privateStateCallbacks) : [];
            return new OSF.AsyncMethodCall(params.requiredArguments || [], supportedOptions, privateStateCallbacks, params.onSucceeded, params.onFailed, params.checkCallArgs, params.method);
        }
    })(AsyncMethodCalls = OSF.AsyncMethodCalls || (OSF.AsyncMethodCalls = {}));
})(OSF || (OSF = {}));
var OSF;
(function (OSF) {
    OSF.AsyncMethods = {
        AddHandlerAsync: "addHandlerAsync",
        CloseAsync: "close",
        CloseContainerAsync: "closeContainer",
        DisplayDialogAsync: "displayDialogAsync",
        GetAccessTokenAsync: "getAccessTokenAsync",
        OpenBrowserWindow: "openBrowserWindow",
        RemoveHandlerAsync: "removeHandlerAsync",
    };
})(OSF || (OSF = {}));
var OSF;
(function (OSF) {
    var DialogEventArgs = (function () {
        function DialogEventArgs(message) {
            if (message[OSF.PropertyDescriptors.MessageType] == 0) {
                OSF.OUtil.defineEnumerableProperties(this, {
                    "type": {
                        value: OSF.EventType.DialogMessageReceived
                    },
                    "message": {
                        value: message[OSF.PropertyDescriptors.MessageContent]
                    }
                });
            }
            else {
                OSF.OUtil.defineEnumerableProperties(this, {
                    "type": {
                        value: OSF.EventType.DialogEventReceived
                    },
                    "error": {
                        value: message[OSF.PropertyDescriptors.MessageType]
                    }
                });
            }
        }
        return DialogEventArgs;
    }());
    OSF.DialogEventArgs = DialogEventArgs;
    var PropertyDescriptors;
    (function (PropertyDescriptors) {
        PropertyDescriptors.MessageType = "messageType";
        PropertyDescriptors.MessageContent = "messageContent";
    })(PropertyDescriptors = OSF.PropertyDescriptors || (OSF.PropertyDescriptors = {}));
})(OSF || (OSF = {}));
var OSF;
(function (OSF) {
    var DialogParentEventArgs = (function () {
        function DialogParentEventArgs(message) {
            OSF.OUtil.defineEnumerableProperties(this, {
                "type": {
                    value: OSF.EventType.DialogParentMessageReceived
                },
                "message": {
                    value: message[OSF.PropertyDescriptors.MessageContent]
                }
            });
        }
        return DialogParentEventArgs;
    }());
    OSF.DialogParentEventArgs = DialogParentEventArgs;
})(OSF || (OSF = {}));
var OSF;
(function (OSF) {
    ;
    OSF.DialogShownStatus = {
        hasDialogShown: false,
        isWindowDialog: false
    };
    var DispIdHost;
    (function (DispIdHost) {
        var dispIdMap;
        function InvokeMethod(methodName, suppliedArguments, caller, privateState) {
            var callArgs;
            try {
                var asyncMethodCall = OSF.AsyncMethodCalls.get(methodName);
                callArgs = asyncMethodCall.verifyAndExtractCall(suppliedArguments, caller, privateState);
                var dispId = getDispIdMap()[methodName];
                var invoker = getHostDelegates("executeAsync");
                var richApiInExcelMethodSubstitution = null;
                if (window.Excel && window.Office.context.requirements.isSetSupported("RedirectV1Api")) {
                    window.Excel._RedirectV1APIs = true;
                }
                if (window.Excel && window.Excel._RedirectV1APIs && (richApiInExcelMethodSubstitution = window.Excel._V1APIMap[methodName])) {
                    var preprocessedCallArgs = OSF.OUtil.shallowCopy(callArgs);
                    delete preprocessedCallArgs[OSF.ParameterNames.AsyncContext];
                    if (richApiInExcelMethodSubstitution.preprocess) {
                        preprocessedCallArgs = richApiInExcelMethodSubstitution.preprocess(preprocessedCallArgs);
                    }
                    var ctx = new window.Excel.RequestContext();
                    var result = richApiInExcelMethodSubstitution.call(ctx, preprocessedCallArgs);
                    ctx.sync()
                        .then(function () {
                        var response = result.value;
                        var status = response.status;
                        delete response["status"];
                        delete response["@odata.type"];
                        if (richApiInExcelMethodSubstitution.postprocess) {
                            response = richApiInExcelMethodSubstitution.postprocess(response, preprocessedCallArgs);
                        }
                        if (status != 0) {
                            response = OSF.DDA.ErrorCodeManager.getErrorArgs(status);
                        }
                        issueAsyncResult(callArgs, status, response);
                    })["catch"](function (error) {
                        issueAsyncResult(callArgs, 13991, null);
                    });
                }
                else {
                    var hostCallArgs;
                    hostCallArgs = OSF.HostParameterMap.toHost(dispId, callArgs);
                    var startTime = (new Date()).getTime();
                    invoker({
                        "dispId": dispId,
                        "hostCallArgs": hostCallArgs,
                        "onCalling": function OSF_DDA_DispIdFacade$Execute_onCalling() { },
                        "onReceiving": function OSF_DDA_DispIdFacade$Execute_onReceiving() { },
                        "onComplete": function (status, hostResponseArgs) {
                            var responseArgs;
                            if (status == 0) {
                                responseArgs = OSF.HostParameterMap.fromHost(dispId, hostResponseArgs);
                            }
                            else {
                                responseArgs = hostResponseArgs;
                            }
                            var payload = asyncMethodCall.processResponse(status, responseArgs, caller, callArgs);
                            issueAsyncResult(callArgs, status, payload);
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
        }
        DispIdHost.InvokeMethod = InvokeMethod;
        function AddEventHandler(suppliedArguments, eventDispatch, caller, isPopupWindow) {
            var callArgs;
            var eventType, handler;
            var isObjectEvent = false;
            function onEnsureRegistration(status) {
                if (status == 0) {
                    var added = !isObjectEvent ? eventDispatch.addEventHandler(eventType, handler) :
                        eventDispatch.addObjectEventHandler(eventType, callArgs[OSF.ParameterNames.Id], handler);
                    if (!added) {
                        status = 13991;
                    }
                }
                var error;
                if (status != 0) {
                    error = OSF.DDA.ErrorCodeManager.getErrorArgs(status);
                }
                issueAsyncResult(callArgs, status, error);
            }
            try {
                var asyncMethodCall = OSF.AsyncMethodCalls.get(OSF.AsyncMethods.AddHandlerAsync);
                callArgs = asyncMethodCall.verifyAndExtractCall(suppliedArguments, caller, eventDispatch);
                eventType = callArgs[OSF.ParameterNames.EventType];
                handler = callArgs[OSF.ParameterNames.Handler];
                if (isPopupWindow) {
                    onEnsureRegistration(0);
                    return;
                }
                var dispId_1 = getDispIdMap()[eventType];
                isObjectEvent = IsObjectEvent(dispId_1);
                var targetId_1 = (isObjectEvent ? callArgs[OSF.ParameterNames.Id] : (caller.id || ""));
                var count = isObjectEvent ? eventDispatch.getObjectEventHandlerCount(eventType, targetId_1) : eventDispatch.getEventHandlerCount(eventType);
                if (count == 0) {
                    var invoker = getHostDelegates("registerEventAsync");
                    invoker({
                        "eventType": eventType,
                        "dispId": dispId_1,
                        "targetId": targetId_1,
                        "onCalling": function OSF_DDA_DispIdFacade$Execute_onCalling() { },
                        "onReceiving": function OSF_DDA_DispIdFacade$Execute_onReceiving() { },
                        "onComplete": onEnsureRegistration,
                        "onEvent": function handleEvent(hostArgs) {
                            var args = OSF.HostParameterMap.fromHost(dispId_1, hostArgs);
                            if (!isObjectEvent)
                                eventDispatch.fireEvent(OSF.manufactureEventArgs(eventType, caller, args));
                            else
                                eventDispatch.fireObjectEvent(targetId_1, OSF.manufactureEventArgs(eventType, targetId_1, args));
                        }
                    });
                }
                else {
                    onEnsureRegistration(0);
                }
            }
            catch (ex) {
                onException(ex, asyncMethodCall, suppliedArguments, callArgs);
            }
        }
        DispIdHost.AddEventHandler = AddEventHandler;
        function RemoveEventHandler(suppliedArguments, eventDispatch, caller) {
            var callArgs;
            var eventType, handler;
            var isObjectEvent = false;
            function onEnsureRegistration(status) {
                var error;
                if (status != 0) {
                    error = OSF.DDA.ErrorCodeManager.getErrorArgs(status);
                }
                issueAsyncResult(callArgs, status, error);
            }
            try {
                var asyncMethodCall = OSF.AsyncMethodCalls.get(OSF.AsyncMethods.RemoveHandlerAsync);
                callArgs = asyncMethodCall.verifyAndExtractCall(suppliedArguments, caller, eventDispatch);
                eventType = callArgs[OSF.ParameterNames.EventType];
                handler = callArgs[OSF.ParameterNames.Handler];
                var dispId = getDispIdMap()[eventType];
                isObjectEvent = IsObjectEvent(dispId);
                var targetId = (isObjectEvent ? callArgs[OSF.ParameterNames.Id] : (caller.id || ""));
                var status, removeSuccess;
                if (handler === null) {
                    removeSuccess = isObjectEvent ? eventDispatch.clearObjectEventHandlers(eventType, targetId) : eventDispatch.clearEventHandlers(eventType);
                    status = 0;
                }
                else {
                    removeSuccess = isObjectEvent ? eventDispatch.removeObjectEventHandler(eventType, targetId, handler) : eventDispatch.removeEventHandler(eventType, handler);
                    status = removeSuccess ? 0 : 5003;
                }
                var count = isObjectEvent ? eventDispatch.getObjectEventHandlerCount(eventType, targetId) : eventDispatch.getEventHandlerCount(eventType);
                if (removeSuccess && count == 0) {
                    var invoker = getHostDelegates("unregisterEventAsync");
                    invoker({
                        "eventType": eventType,
                        "dispId": dispId,
                        "targetId": targetId,
                        "onCalling": function OSF_DDA_DispIdFacade$Execute_onCalling() { },
                        "onReceiving": function OSF_DDA_DispIdFacade$Execute_onReceiving() { },
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
        }
        DispIdHost.RemoveEventHandler = RemoveEventHandler;
        function OpenDialog(suppliedArguments, eventDispatch, caller) {
            var callArgs;
            var targetId;
            var dialogMessageEvent = OSF.EventType.DialogMessageReceived;
            var dialogOtherEvent = OSF.EventType.DialogEventReceived;
            function onEnsureRegistration(status) {
                var payload;
                if (status != 0) {
                    payload = OSF.DDA.ErrorCodeManager.getErrorArgs(status);
                }
                else {
                    var onSucceedArgs = {};
                    onSucceedArgs["id"] = targetId;
                    onSucceedArgs["data"] = eventDispatch;
                    var payload = asyncMethodCall.processResponse(status, onSucceedArgs, caller, callArgs);
                    OSF.DialogShownStatus.hasDialogShown = true;
                    eventDispatch.clearEventHandlers(dialogMessageEvent);
                    eventDispatch.clearEventHandlers(dialogOtherEvent);
                }
                issueAsyncResult(callArgs, status, payload);
            }
            try {
                if (dialogMessageEvent == undefined || dialogOtherEvent == undefined) {
                    onEnsureRegistration(5000);
                }
                if (OSF.AsyncMethods.DisplayDialogAsync == null) {
                    onEnsureRegistration(5001);
                    return;
                }
                var asyncMethodCall = OSF.AsyncMethodCalls.get(OSF.AsyncMethods.DisplayDialogAsync);
                callArgs = asyncMethodCall.verifyAndExtractCall(suppliedArguments, caller, eventDispatch);
                var dispId = getDispIdMap()[dialogMessageEvent];
                var invoker = getHostDelegates("openDialog");
                targetId = JSON.stringify(callArgs);
                if (!OSF.DialogShownStatus.hasDialogShown) {
                    eventDispatch.clearQueuedEvent(dialogMessageEvent);
                    eventDispatch.clearQueuedEvent(dialogOtherEvent);
                    eventDispatch.clearQueuedEvent(OSF.EventType.DialogParentMessageReceived);
                }
                invoker({
                    "eventType": dialogMessageEvent,
                    "dispId": dispId,
                    "targetId": targetId,
                    "onCalling": function OSF_DDA_DispIdFacade$Execute_onCalling() {
                    },
                    "onReceiving": function OSF_DDA_DispIdFacade$Execute_onReceiving() {
                    },
                    "onComplete": onEnsureRegistration,
                    "onEvent": function handleEvent(hostArgs) {
                        var args = OSF.HostParameterMap.fromHost(dispId, hostArgs);
                        var event = OSF.manufactureEventArgs(dialogMessageEvent, caller, args);
                        if (event.type == dialogOtherEvent) {
                            var payload = OSF.DDA.ErrorCodeManager.getErrorArgs(event.error);
                            var errorArgs = {};
                            errorArgs["code"] = status || 5001;
                            event.error = new OSF.DDA.Error(errorArgs[OSF.DDA.AsyncResultEnum.ErrorProperties.Name], errorArgs[OSF.DDA.AsyncResultEnum.ErrorProperties.Message], errorArgs[OSF.DDA.AsyncResultEnum.ErrorProperties.Code]);
                        }
                        eventDispatch.fireOrQueueEvent(event);
                        if (args["messageType"] == 12006) {
                            eventDispatch.clearEventHandlers(dialogMessageEvent);
                            eventDispatch.clearEventHandlers(dialogOtherEvent);
                            eventDispatch.clearEventHandlers(OSF.EventType.DialogParentMessageReceived);
                            OSF.DialogShownStatus.hasDialogShown = false;
                        }
                    }
                });
            }
            catch (ex) {
                onException(ex, asyncMethodCall, suppliedArguments, callArgs);
            }
        }
        DispIdHost.OpenDialog = OpenDialog;
        function CloseDialog(suppliedArguments, targetId, eventDispatch, caller) {
            var callArgs;
            var dialogMessageEvent, dialogOtherEvent;
            var closeStatus = 0;
            function closeCallback(status) {
                closeStatus = status;
                OSF.DialogShownStatus.hasDialogShown = false;
            }
            try {
                var asyncMethodCall = OSF.AsyncMethodCalls.get(OSF.AsyncMethods.CloseAsync);
                callArgs = asyncMethodCall.verifyAndExtractCall(suppliedArguments, caller, eventDispatch);
                dialogMessageEvent = OSF.EventType.DialogMessageReceived;
                dialogOtherEvent = OSF.EventType.DialogEventReceived;
                eventDispatch.clearEventHandlers(dialogMessageEvent);
                eventDispatch.clearEventHandlers(dialogOtherEvent);
                var dispId = getDispIdMap()[dialogMessageEvent];
                var invoker = getHostDelegates("closeDialog");
                invoker({
                    "eventType": dialogMessageEvent,
                    "dispId": dispId,
                    "targetId": targetId,
                    "onCalling": function OSF_DDA_DispIdFacade$Execute_onCalling() { },
                    "onReceiving": function OSF_DDA_DispIdFacade$Execute_onReceiving() { },
                    "onComplete": closeCallback
                });
            }
            catch (ex) {
                onException(ex, asyncMethodCall, suppliedArguments, callArgs);
            }
            if (closeStatus != 0) {
            }
        }
        DispIdHost.CloseDialog = CloseDialog;
        function MessageParent(suppliedArguments, caller) {
            var stateInfo = {};
            var syncMethodCall = OSF.SyncMethodCalls.get(OSF.SyncMethods.MessageParent);
            var callArgs = syncMethodCall.verifyAndExtractCall(suppliedArguments, caller, stateInfo);
            var invoker = getHostDelegates("messageParent");
            var dispId = getDispIdMap()[OSF.SyncMethods.MessageParent];
            return invoker({
                "dispId": dispId,
                "hostCallArgs": callArgs,
                "onCalling": function OSF_DDA_DispIdFacade$Execute_onCalling() { },
                "onReceiving": function OSF_DDA_DispIdFacade$Execute_onReceiving() { },
            });
        }
        DispIdHost.MessageParent = MessageParent;
        function SendMessage(suppliedArguments, eventDispatch, caller) {
            var stateInfo = {};
            var syncMethodCall = OSF.SyncMethodCalls.get(OSF.SyncMethods.SendMessage);
            var callArgs = syncMethodCall.verifyAndExtractCall(suppliedArguments, caller, stateInfo);
            var invoker = getHostDelegates("sendMessage");
            var dispId = getDispIdMap()[OSF.SyncMethods.SendMessage];
            return invoker({
                "dispId": dispId,
                "hostCallArgs": callArgs,
                "onCalling": function OSF_DDA_DispIdFacade$Execute_onCalling() { },
                "onReceiving": function OSF_DDA_DispIdFacade$Execute_onReceiving() { },
            });
        }
        DispIdHost.SendMessage = SendMessage;
        function addAsyncMethods(target, asyncMethodNames, privateState) {
            for (var entry in asyncMethodNames) {
                var methodName = asyncMethodNames[entry];
                if (!target[methodName]) {
                    OSF.OUtil.defineEnumerableProperty(target, methodName, {
                        value: (function (asyncMethod) {
                            return function () {
                                var invokeMethod = OSF.DispIdHost.InvokeMethod;
                                invokeMethod(asyncMethod, arguments, target, privateState);
                            };
                        })(methodName)
                    });
                }
            }
        }
        DispIdHost.addAsyncMethods = addAsyncMethods;
        function addEventSupport(target, eventDispatch, isPopupWindow) {
            var add = OSF.AsyncMethods.AddHandlerAsync;
            var remove = OSF.AsyncMethods.RemoveHandlerAsync;
            if (!target[add]) {
                OSF.OUtil.defineEnumerableProperty(target, add, {
                    value: function () {
                        var addEventHandler = OSF.DispIdHost.AddEventHandler;
                        addEventHandler(arguments, eventDispatch, target, isPopupWindow);
                    }
                });
            }
            if (!target[remove]) {
                OSF.OUtil.defineEnumerableProperty(target, remove, {
                    value: function () {
                        var removeEventHandler = OSF.DispIdHost.RemoveEventHandler;
                        removeEventHandler(arguments, eventDispatch, target);
                    }
                });
            }
        }
        DispIdHost.addEventSupport = addEventSupport;
        function IsObjectEvent(dispId) {
            return (dispId == OSF.EventDispId.dispidObjectDeletedEvent ||
                dispId == OSF.EventDispId.dispidObjectSelectionChangedEvent ||
                dispId == OSF.EventDispId.dispidObjectDataChangedEvent ||
                dispId == OSF.EventDispId.dispidContentControlAddedEvent);
        }
        function onException(ex, asyncMethodCall, suppliedArgs, callArgs) {
            if (typeof ex == "number") {
                if (!callArgs) {
                    callArgs = asyncMethodCall.getCallArgs(suppliedArgs);
                }
                issueAsyncResult(callArgs, ex, OSF.DDA.ErrorCodeManager.getErrorArgs(ex));
            }
            else {
                throw ex;
            }
        }
        function getHostDelegates(method) {
            var namespace;
            var hostInfo = OSF._OfficeAppFactory.getHostInfo();
            if (hostInfo.hostPlatform == OSF.HostInfoPlatform.web) {
                namespace = OSF.WACDelegate;
            }
            else {
                namespace = OSF.SafeArrayDelegate;
            }
            return namespace[method];
        }
        function initializeDispIdHostFacade() {
            dispIdMap = {};
            var methodMap = {
                "GoToByIdAsync": 82,
                "GetSelectedDataAsync": 64,
                "SetSelectedDataAsync": 65,
                "GetDocumentCopyChunkAsync": 80,
                "ReleaseDocumentCopyAsync": 81,
                "GetDocumentCopyAsync": 77,
                "AddFromSelectionAsync": 66,
                "AddFromPromptAsync": 67,
                "AddFromNamedItemAsync": 78,
                "GetAllAsync": 74,
                "GetByIdAsync": 68,
                "ReleaseByIdAsync": 69,
                "GetDataAsync": 70,
                "SetDataAsync": 71,
                "AddRowsAsync": 72,
                "AddColumnsAsync": 79,
                "DeleteAllDataValuesAsync": 73,
                "RefreshAsync": 75,
                "SaveAsync": 76,
                "GetActiveViewAsync": 83,
                "GetFilePropertiesAsync": 86,
                "GetOfficeThemeAsync": 85,
                "GetDocumentThemeAsync": 84,
                "ClearFormatsAsync": 87,
                "SetTableOptionsAsync": 88,
                "SetFormatsAsync": 89,
                "GetUserIdentityInfoAsync": 92,
                "GetAccessTokenAsync": 98,
                "GetAuthContextAsync": 99,
                "ExecuteRichApiRequestAsync": 93,
                "AppCommandInvocationCompletedAsync": 94,
                "CloseContainerAsync": 97,
                "OpenBrowserWindow": 102,
                "CreateDocumentAsync": 105,
                "InsertFormAsync": 106,
                "ExecuteFeature": 146,
                "QueryFeature": 147,
                "AddDataPartAsync": 128,
                "GetDataPartByIdAsync": 129,
                "GetDataPartsByNameSpaceAsync": 130,
                "GetPartXmlAsync": 131,
                "GetPartNodesAsync": 132,
                "DeleteDataPartAsync": 133,
                "GetNodeValueAsync": 134,
                "GetNodeXmlAsync": 135,
                "GetRelativeNodesAsync": 136,
                "SetNodeValueAsync": 137,
                "SetNodeXmlAsync": 138,
                "AddDataPartNamespaceAsync": 139,
                "GetDataPartNamespaceAsync": 140,
                "GetDataPartPrefixAsync": 141,
                "GetNodeTextAsync": 142,
                "SetNodeTextAsync": 143,
                "GetSelectedTask": 110,
                "GetTask": 112,
                "GetWSSUrl": 114,
                "GetTaskField": 115,
                "GetSelectedResource": 111,
                "GetResourceField": 113,
                "GetProjectField": 116,
                "GetSelectedView": 117,
                "GetTaskByIndex": 118,
                "GetResourceByIndex": 119,
                "SetTaskField": 120,
                "SetResourceField": 121,
                "GetMaxTaskIndex": 122,
                "GetMaxResourceIndex": 123,
                "CreateTask": 124
            };
            for (var method in methodMap) {
                if (OSF.AsyncMethods[method]) {
                    dispIdMap[OSF.AsyncMethods[method]] = methodMap[method];
                }
            }
            var syncMethodMap = {
                "MessageParent": 144,
                "SendMessage": 145
            };
            for (var method in syncMethodMap) {
                if (OSF.SyncMethods[method]) {
                    dispIdMap[OSF.SyncMethods[method]] = syncMethodMap[method];
                }
            }
            var eventMap = {
                "SettingsChanged": OSF.EventDispId.dispidSettingsChangedEvent,
                "DocumentSelectionChanged": OSF.EventDispId.dispidDocumentSelectionChangedEvent,
                "BindingSelectionChanged": OSF.EventDispId.dispidBindingSelectionChangedEvent,
                "BindingDataChanged": OSF.EventDispId.dispidBindingDataChangedEvent,
                "ActiveViewChanged": OSF.EventDispId.dispidActiveViewChangedEvent,
                "OfficeThemeChanged": OSF.EventDispId.dispidOfficeThemeChangedEvent,
                "DocumentThemeChanged": OSF.EventDispId.dispidDocumentThemeChangedEvent,
                "AppCommandInvoked": OSF.EventDispId.dispidAppCommandInvokedEvent,
                "DialogMessageReceived": OSF.EventDispId.dispidDialogMessageReceivedEvent,
                "DialogParentMessageReceived": OSF.EventDispId.dispidDialogParentMessageReceivedEvent,
                "ObjectDeleted": OSF.EventDispId.dispidObjectDeletedEvent,
                "ObjectSelectionChanged": OSF.EventDispId.dispidObjectSelectionChangedEvent,
                "ObjectDataChanged": OSF.EventDispId.dispidObjectDataChangedEvent,
                "ContentControlAdded": OSF.EventDispId.dispidContentControlAddedEvent,
                "RichApiMessage": OSF.EventDispId.dispidRichApiMessageEvent,
                "DataNodeInserted": OSF.EventDispId.dispidDataNodeAddedEvent,
                "DataNodeReplaced": OSF.EventDispId.dispidDataNodeReplacedEvent,
                "DataNodeDeleted": OSF.EventDispId.dispidDataNodeDeletedEvent
            };
            for (var event in eventMap) {
                if (OSF.EventType[event]) {
                    dispIdMap[OSF.EventType[event]] = eventMap[event];
                }
            }
        }
        function getDispIdMap() {
            if (!dispIdMap) {
                initializeDispIdHostFacade();
            }
            return dispIdMap;
        }
        function issueAsyncResult(callArgs, status, payload) {
            var callback = callArgs[OSF.ParameterNames.Callback];
            if (callback) {
                var asyncInitArgs = {};
                asyncInitArgs[OSF.DDA.AsyncResultEnum.Properties.Context] = callArgs[OSF.ParameterNames.AsyncContext];
                var errorArgs;
                if (status == 0) {
                    asyncInitArgs[OSF.DDA.AsyncResultEnum.Properties.Value] = payload;
                }
                else {
                    errorArgs = {};
                    payload = payload || OSF.DDA.ErrorCodeManager.getErrorArgs(5001);
                    errorArgs[OSF.DDA.AsyncResultEnum.ErrorProperties.Code] = status || 5001;
                    errorArgs[OSF.DDA.AsyncResultEnum.ErrorProperties.Name] = payload.name || payload;
                    errorArgs[OSF.DDA.AsyncResultEnum.ErrorProperties.Message] = payload.message || payload;
                }
                callback(new OSF.DDA.AsyncResult(asyncInitArgs, errorArgs));
            }
        }
    })(DispIdHost = OSF.DispIdHost || (OSF.DispIdHost = {}));
})(OSF || (OSF = {}));
var OSF;
(function (OSF) {
    function manufactureEventArgs(eventType, target, eventProperties) {
        var args;
        switch (eventType) {
            case OSF.EventType.DialogMessageReceived:
                args = new OSF.DialogEventArgs(eventProperties);
                break;
            case OSF.EventType.DialogParentMessageReceived:
                args = new OSF.DialogParentEventArgs(eventProperties);
                break;
        }
        return args;
    }
    OSF.manufactureEventArgs = manufactureEventArgs;
})(OSF || (OSF = {}));
var OSF;
(function (OSF) {
    var toHostMap = "toHost";
    var fromHostMap = "fromHost";
    var sourceData = "sourceData";
    var HostParameterMap;
    (function (HostParameterMap) {
        HostParameterMap.self = "self";
        var dynamicTypes = {
            data: {
                toHost: function (data) {
                    if (data != null && data.rows !== undefined) {
                        var tableData = {};
                        tableData["tableRows"] = data.rows;
                        tableData["tableHeaders"] = data.headers;
                        data = tableData;
                    }
                    return data;
                },
                fromHost: function (args) {
                    return args;
                }
            }
        };
        dynamicTypes["sampleData"] = dynamicTypes["data"];
        var specialProcessor;
        var mappings = {};
        function define(definition) {
            var args = {};
            var toHost = createObject(definition.toHost);
            if (definition.invertible) {
                args.map = toHost;
            }
            else if (definition.canonical) {
                args.toHost = args.fromHost = toHost;
            }
            else {
                args.toHost = toHost;
                args.fromHost = createObject(definition.fromHost);
            }
            addMapping(definition.type, args);
            if (definition.isComplexType)
                addComplexType(definition.type);
        }
        HostParameterMap.define = define;
        function toHost(mapName, preimage) {
            return applyMap(mapName, preimage, toHostMap);
        }
        HostParameterMap.toHost = toHost;
        function fromHost(mapName, image) {
            return applyMap(mapName, image, fromHostMap);
        }
        HostParameterMap.fromHost = fromHost;
        function addMapping(mapName, description) {
            var toHost, fromHost;
            if (description.map) {
                toHost = description.map;
                fromHost = {};
                for (var preimage in toHost) {
                    var image = toHost[preimage];
                    if (image == HostParameterMap.self) {
                        image = preimage;
                    }
                    fromHost[image] = preimage;
                }
            }
            else {
                toHost = description.toHost;
                fromHost = description.fromHost;
            }
            var pair = mappings[mapName];
            if (pair) {
                var currMap = pair[toHostMap];
                for (var th in currMap)
                    toHost[th] = currMap[th];
                currMap = pair[fromHostMap];
                for (var fh in currMap)
                    fromHost[fh] = currMap[fh];
            }
            else {
                pair = mappings[mapName] = {};
            }
            pair[toHostMap] = toHost;
            pair[fromHostMap] = fromHost;
        }
        HostParameterMap.addMapping = addMapping;
        function addComplexType(ct) {
            getSpecialProcessor().addComplexType(ct);
        }
        HostParameterMap.addComplexType = addComplexType;
        function getDynamicType(dt) {
            return getSpecialProcessor().getDynamicType(dt);
        }
        HostParameterMap.getDynamicType = getDynamicType;
        function setDynamicType(dt, handler) {
            getSpecialProcessor().setDynamicType(dt, handler);
        }
        HostParameterMap.setDynamicType = setDynamicType;
        function doMapValues(preimageSet, mapping) {
            return mapValues(preimageSet, mapping);
        }
        HostParameterMap.doMapValues = doMapValues;
        function mapValues(preimageSet, mapping) {
            var ret = preimageSet ? {} : undefined;
            for (var entry in preimageSet) {
                var preimage = preimageSet[entry];
                var image;
                if (OSF.ListType.isListType(entry)) {
                    image = [];
                    for (var subEntry in preimage) {
                        image.push(mapValues(preimage[subEntry], mapping));
                    }
                }
                else if (OSF.OUtil.listContainsKey(dynamicTypes, entry)) {
                    image = dynamicTypes[entry][mapping](preimage);
                }
                else if (mapping == fromHostMap && getSpecialProcessor().preserveNesting(entry)) {
                    image = mapValues(preimage, mapping);
                }
                else {
                    var maps = mappings[entry];
                    if (maps) {
                        var map = maps[mapping];
                        if (map) {
                            image = map[preimage];
                            if (image === undefined) {
                                image = preimage;
                            }
                        }
                    }
                    else {
                        image = preimage;
                    }
                }
                ret[entry] = image;
            }
            return ret;
        }
        function generateArguments(imageSet, parameters) {
            var ret;
            for (var param in parameters) {
                var arg;
                if (getSpecialProcessor().isComplexType(param)) {
                    arg = generateArguments(imageSet, mappings[param][toHostMap]);
                }
                else {
                    arg = imageSet[param];
                }
                if (arg != undefined) {
                    if (!ret) {
                        ret = {};
                    }
                    var index = parameters[param];
                    if (index == HostParameterMap.self) {
                        index = param;
                    }
                    ret[index] = getSpecialProcessor().pack(param, arg);
                }
            }
            return ret;
        }
        function extractArguments(source, parameters, extracted) {
            if (!extracted) {
                extracted = {};
            }
            for (var param in parameters) {
                var index = parameters[param];
                var value;
                if (index == HostParameterMap.self) {
                    value = source;
                }
                else if (index == sourceData) {
                    extracted[param] = source.toArray();
                    continue;
                }
                else {
                    value = source[index];
                }
                if (value === null || value === undefined) {
                    extracted[param] = undefined;
                }
                else {
                    value = getSpecialProcessor().unpack(param, value);
                    var map;
                    if (getSpecialProcessor().isComplexType(param)) {
                        map = mappings[param][fromHostMap];
                        if (getSpecialProcessor().preserveNesting(param)) {
                            extracted[param] = extractArguments(value, map);
                        }
                        else {
                            extractArguments(value, map, extracted);
                        }
                    }
                    else {
                        if (OSF.ListType.isListType(param)) {
                            map = {};
                            var entryDescriptor = OSF.ListType.getDescriptor(param);
                            map[entryDescriptor] = HostParameterMap.self;
                            var extractedValues = new Array(value.length);
                            for (var item in value) {
                                extractedValues[item] = extractArguments(value[item], map);
                            }
                            extracted[param] = extractedValues;
                        }
                        else {
                            extracted[param] = value;
                        }
                    }
                }
            }
            return extracted;
        }
        function applyMap(mapName, preimage, mapping) {
            var parameters = mappings[mapName][mapping];
            var image;
            if (mapping == "toHost") {
                var imageSet = mapValues(preimage, mapping);
                image = generateArguments(imageSet, parameters);
            }
            else if (mapping == "fromHost") {
                var argumentSet = extractArguments(preimage, parameters);
                image = mapValues(argumentSet, mapping);
            }
            return image;
        }
        function getSpecialProcessor() {
            if (!specialProcessor) {
                var hostInfo = OSF._OfficeAppFactory.getHostInfo();
                if (hostInfo.hostPlatform == OSF.HostInfoPlatform.web) {
                    specialProcessor = new OSF.WebSpecialProcessor();
                }
                else {
                    specialProcessor = new OSF.SafeArraySpecialProcessor();
                }
            }
            return specialProcessor;
        }
        function createObject(properties) {
            var obj = null;
            if (properties) {
                obj = {};
                var len = properties.length;
                for (var i = 0; i < len; i++) {
                    obj[properties[i].name] = properties[i].value;
                }
            }
            return obj;
        }
    })(HostParameterMap = OSF.HostParameterMap || (OSF.HostParameterMap = {}));
})(OSF || (OSF = {}));
var OSF;
(function (OSF) {
    var ListType;
    (function (ListType) {
        var listTypes;
        function setListType(t, prop) {
            listTypes[t] = prop;
        }
        ListType.setListType = setListType;
        function isListType(t) {
            return OSF.OUtil.listContainsKey(listTypes, t);
        }
        ListType.isListType = isListType;
        function getDescriptor(t) {
            return listTypes[t];
        }
        ListType.getDescriptor = getDescriptor;
    })(ListType = OSF.ListType || (OSF.ListType = {}));
})(OSF || (OSF = {}));
var Office;
(function (Office) {
    var EventType;
    (function (EventType) {
        EventType.DialogMessageReceived = "dialogMessageReceived";
        EventType.DialogParentMessageReceived = "dialogParentMessageReceived";
        EventType.DialogParentEventReceived = "dialogParentEventReceived";
        EventType.DialogEventReceived = "dialogEventReceived";
    })(EventType = Office.EventType || (Office.EventType = {}));
})(Office || (Office = {}));
var OSF;
(function (OSF) {
    var OUtil;
    (function (OUtil) {
        var HostThemeButtonStyleKeys;
        (function (HostThemeButtonStyleKeys) {
            HostThemeButtonStyleKeys["ButtonBorderColor"] = "buttonBorderColor";
            HostThemeButtonStyleKeys["ButtonBackgroundColor"] = "buttonBackgroundColor";
        })(HostThemeButtonStyleKeys = OUtil.HostThemeButtonStyleKeys || (OUtil.HostThemeButtonStyleKeys = {}));
        var ExcelCommonUI;
        (function (ExcelCommonUI) {
            ExcelCommonUI["HostButtonBorderColor"] = "#86bfa0";
            ExcelCommonUI["HostButtonBackgroundColor"] = "#d3f0e0";
        })(ExcelCommonUI || (ExcelCommonUI = {}));
        ;
        var WordCommonUI;
        (function (WordCommonUI) {
            WordCommonUI["HostButtonBorderColor"] = "#a3bde3";
            WordCommonUI["HostButtonBackgroundColor"] = "#d5e1f2";
        })(WordCommonUI || (WordCommonUI = {}));
        ;
        var PowerPointCommonUI;
        (function (PowerPointCommonUI) {
            PowerPointCommonUI["HostButtonBorderColor"] = "#f5ba9d";
            PowerPointCommonUI["HostButtonBackgroundColor"] = "#fcf0ed";
        })(PowerPointCommonUI || (PowerPointCommonUI = {}));
        ;
        function finalizeProperties(obj, descriptor) {
            descriptor = descriptor || {};
            var props = Object.getOwnPropertyNames(obj);
            var propsLength = props.length;
            for (var i = 0; i < propsLength; i++) {
                var prop = props[i];
                var desc = Object.getOwnPropertyDescriptor(obj, prop);
                if (!desc.get && !desc.set) {
                    desc.writable = descriptor.writable || false;
                }
                desc.configurable = descriptor.configurable || false;
                desc.enumerable = descriptor.enumerable || true;
                Object.defineProperty(obj, prop, desc);
            }
            return obj;
        }
        OUtil.finalizeProperties = finalizeProperties;
        function defineEnumerableProperties(obj, descriptors) {
            return defineNondefaultProperties(obj, descriptors, ["enumerable"]);
        }
        OUtil.defineEnumerableProperties = defineEnumerableProperties;
        function defineEnumerableProperty(obj, prop, descriptor) {
            return defineNondefaultProperty(obj, prop, descriptor, ["enumerable"]);
        }
        OUtil.defineEnumerableProperty = defineEnumerableProperty;
        function listContainsKey(list, key) {
            for (var item in list) {
                if (key == item) {
                    return true;
                }
            }
            return false;
        }
        OUtil.listContainsKey = listContainsKey;
        function createObject(properties) {
            var obj = null;
            if (properties) {
                obj = {};
                var len = properties.length;
                for (var i = 0; i < len; i++) {
                    obj[properties[i].name] = properties[i].value;
                }
            }
            return obj;
        }
        OUtil.createObject = createObject;
        function listContainsValue(list, value) {
            for (var item in list) {
                if (value == list[item]) {
                    return true;
                }
            }
            return false;
        }
        OUtil.listContainsValue = listContainsValue;
        function shouldUseLocalStorageToPassMessage() {
            try {
                var osList = [
                    "Windows NT 6.1",
                    "Windows NT 6.2",
                    "Windows NT 6.3",
                    "Windows NT 10.0"
                ];
                var userAgent = window.navigator.userAgent;
                for (var i = 0, len = osList.length; i < len; i++) {
                    if (userAgent.indexOf(osList[i]) > -1) {
                        return isInternetExplorer();
                    }
                }
                return false;
            }
            catch (e) {
                logExceptionToBrowserConsole("Error happens in shouldUseLocalStorageToPassMessage.", e);
                return false;
            }
        }
        OUtil.shouldUseLocalStorageToPassMessage = shouldUseLocalStorageToPassMessage;
        function isInternetExplorer() {
            try {
                var userAgent = window.navigator.userAgent;
                return userAgent.indexOf("MSIE ") > -1 || userAgent.indexOf("Trident/") > -1 || userAgent.indexOf("Edge/") > -1;
            }
            catch (e) {
                logExceptionToBrowserConsole("Error happens in isInternetExplorer.", e);
                return false;
            }
        }
        OUtil.isInternetExplorer = isInternetExplorer;
        function serializeObjectToString(obj) {
            if (typeof (JSON) !== "undefined") {
                try {
                    return JSON.stringify(obj);
                }
                catch (ex) {
                }
            }
            return "";
        }
        OUtil.serializeObjectToString = serializeObjectToString;
        function formatString() {
            var arglist = [];
            for (var _i = 0; _i < arguments.length; _i++) {
                arglist[_i] = arguments[_i];
            }
            var args = arglist;
            var source = args[0];
            return source.replace(/{(\d+)}/gm, function (match, number) {
                var index = parseInt(number, 10) + 1;
                return args[index] === undefined ? '{' + number + '}' : args[index];
            });
        }
        OUtil.formatString = formatString;
        function addHostInfoAsQueryParam(url, hostInfoValue) {
            if (!url) {
                return null;
            }
            url = url.trim() || '';
            var questionMark = "?";
            var hostInfo = "_host_Info=";
            var ampHostInfo = "&_host_Info=";
            var fragmentSeparator = "#";
            var urlParts = url.split(fragmentSeparator);
            var urlWithoutFragment = urlParts.shift();
            var fragment = urlParts.join(fragmentSeparator);
            var querySplits = urlWithoutFragment.split(questionMark);
            var urlWithoutFragmentWithHostInfo;
            if (querySplits.length > 1) {
                urlWithoutFragmentWithHostInfo = urlWithoutFragment + ampHostInfo + hostInfoValue;
            }
            else if (querySplits.length > 0) {
                urlWithoutFragmentWithHostInfo = urlWithoutFragment + questionMark + hostInfo + hostInfoValue;
            }
            if (fragment) {
                return [urlWithoutFragmentWithHostInfo, fragmentSeparator, fragment].join('');
            }
            else {
                return urlWithoutFragmentWithHostInfo;
            }
        }
        OUtil.addHostInfoAsQueryParam = addHostInfoAsQueryParam;
        function getHostnamePortionForLogging(hostname) {
            var hostnameSubstrings = hostname.split('.');
            var len = hostnameSubstrings.length;
            if (len >= 2) {
                return hostnameSubstrings[len - 2] + "." + hostnameSubstrings[len - 1];
            }
            else if (len == 1) {
                return hostnameSubstrings[0];
            }
        }
        OUtil.getHostnamePortionForLogging = getHostnamePortionForLogging;
        function shallowCopy(sourceObj) {
            if (sourceObj == null) {
                return null;
            }
            else if (!(sourceObj instanceof Object)) {
                return sourceObj;
            }
            else if (Array.isArray(sourceObj)) {
                var copyArr = [];
                for (var i = 0; i < sourceObj.length; i++) {
                    copyArr.push(sourceObj[i]);
                }
                return copyArr;
            }
            else {
                var copyObj = sourceObj.constructor();
                for (var property in sourceObj) {
                    if (sourceObj.hasOwnProperty(property)) {
                        copyObj[property] = sourceObj[property];
                    }
                }
                return copyObj;
            }
        }
        OUtil.shallowCopy = shallowCopy;
        function getXdmEventName(targetId, eventType) {
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
        }
        OUtil.getXdmEventName = getXdmEventName;
        function getCommonUI() {
            var hostType = Office.context.host;
            switch (hostType) {
                case Office.HostType.Excel:
                    return ExcelCommonUI;
                case Office.HostType.Word:
                    return WordCommonUI;
                case Office.HostType.PowerPoint:
                    return PowerPointCommonUI;
            }
            return null;
        }
        OUtil.getCommonUI = getCommonUI;
        function getDomainForUrl(url) {
            if (!url) {
                return null;
            }
            var url_parser = document.createElement('a');
            url_parser.href = url;
            return url_parser.protocol + "//" + url_parser.host;
        }
        OUtil.getDomainForUrl = getDomainForUrl;
        function delayExecutionAndCache() {
            var args = [];
            for (var _i = 0; _i < arguments.length; _i++) {
                args[_i] = arguments[_i];
            }
            var obj = { calc: args[0] };
            if (obj.calc) {
                obj.val = obj.calc.apply(this, args);
                delete obj.calc;
            }
            return obj.val;
        }
        OUtil.delayExecutionAndCache = delayExecutionAndCache;
        function defineNondefaultProperties(obj, descriptors, attributes) {
            descriptors = descriptors || {};
            for (var prop in descriptors) {
                defineNondefaultProperty(obj, prop, descriptors[prop], attributes);
            }
            return obj;
        }
        function defineNondefaultProperty(obj, prop, descriptor, attributes) {
            descriptor = descriptor || {};
            for (var nd in attributes) {
                var attribute = attributes[nd];
                if (descriptor[attribute] == undefined) {
                    descriptor[attribute] = true;
                }
            }
            Object.defineProperty(obj, prop, descriptor);
            return obj;
        }
        function logExceptionToBrowserConsole(message, exception) {
            OSF.Utility.trace(message + " Exception details: " + serializeObjectToString(exception));
        }
    })(OUtil = OSF.OUtil || (OSF.OUtil = {}));
})(OSF || (OSF = {}));
var OSF;
(function (OSF) {
    var ParameterNames;
    (function (ParameterNames) {
        ParameterNames.BindingType = "bindingType";
        ParameterNames.CoercionType = "coercionType";
        ParameterNames.ValueFormat = "valueFormat";
        ParameterNames.FilterType = "filterType";
        ParameterNames.Columns = "columns";
        ParameterNames.SampleData = "sampleData";
        ParameterNames.GoToType = "goToType";
        ParameterNames.SelectionMode = "selectionMode";
        ParameterNames.Id = "id";
        ParameterNames.PromptText = "promptText";
        ParameterNames.ItemName = "itemName";
        ParameterNames.FailOnCollision = "failOnCollision";
        ParameterNames.StartRow = "startRow";
        ParameterNames.StartColumn = "startColumn";
        ParameterNames.RowCount = "rowCount";
        ParameterNames.ColumnCount = "columnCount";
        ParameterNames.Rows = "rows";
        ParameterNames.OverwriteIfStale = "overwriteIfStale";
        ParameterNames.FileType = "fileType";
        ParameterNames.EventType = "eventType";
        ParameterNames.Handler = "handler";
        ParameterNames.SliceSize = "sliceSize";
        ParameterNames.SliceIndex = "sliceIndex";
        ParameterNames.ActiveView = "activeView";
        ParameterNames.Status = "status";
        ParameterNames.PlatformType = "platformType";
        ParameterNames.HostType = "hostType";
        ParameterNames.Email = "email";
        ParameterNames.ForceConsent = "forceConsent";
        ParameterNames.ForceAddAccount = "forceAddAccount";
        ParameterNames.AuthChallenge = "authChallenge";
        ParameterNames.AllowConsentPrompt = "allowConsentPrompt";
        ParameterNames.ForMSGraphAccess = "forMSGraphAccess";
        ParameterNames.AllowSignInPrompt = "allowSignInPrompt";
        ParameterNames.JsonPayload = "jsonPayload";
        ParameterNames.EnableNewHosts = "enableNewHosts";
        ParameterNames.AccountTypeFilter = "accountTypeFilter";
        ParameterNames.AddinTrustId = "addinTrustId";
        ParameterNames.Reserved = "reserved";
        ParameterNames.Tcid = "tcid";
        ParameterNames.Xml = "xml";
        ParameterNames.Namespace = "namespace";
        ParameterNames.Prefix = "prefix";
        ParameterNames.XPath = "xPath";
        ParameterNames.Text = "text";
        ParameterNames.ImageLeft = "imageLeft";
        ParameterNames.ImageTop = "imageTop";
        ParameterNames.ImageWidth = "imageWidth";
        ParameterNames.ImageHeight = "imageHeight";
        ParameterNames.TaskId = "taskId";
        ParameterNames.FieldId = "fieldId";
        ParameterNames.FieldValue = "fieldValue";
        ParameterNames.ServerUrl = "serverUrl";
        ParameterNames.ListName = "listName";
        ParameterNames.ResourceId = "resourceId";
        ParameterNames.ViewType = "viewType";
        ParameterNames.ViewName = "viewName";
        ParameterNames.GetRawValue = "getRawValue";
        ParameterNames.CellFormat = "cellFormat";
        ParameterNames.TableOptions = "tableOptions";
        ParameterNames.TaskIndex = "taskIndex";
        ParameterNames.ResourceIndex = "resourceIndex";
        ParameterNames.CustomFieldId = "customFieldId";
        ParameterNames.Url = "url";
        ParameterNames.MessageHandler = "messageHandler";
        ParameterNames.Width = "width";
        ParameterNames.Height = "height";
        ParameterNames.RequireHTTPs = "requireHTTPS";
        ParameterNames.DisplayInIframe = "displayInIframe";
        ParameterNames.HideTitle = "hideTitle";
        ParameterNames.UseDeviceIndependentPixels = "useDeviceIndependentPixels";
        ParameterNames.PromptBeforeOpen = "promptBeforeOpen";
        ParameterNames.EnforceAppDomain = "enforceAppDomain";
        ParameterNames.UrlNoHostInfo = "urlNoHostInfo";
        ParameterNames.Base64 = "base64";
        ParameterNames.FormId = "formId";
    })(ParameterNames = OSF.ParameterNames || (OSF.ParameterNames = {}));
})(OSF || (OSF = {}));
var OSF;
(function (OSF) {
    var SafeArrayDelegate;
    (function (SafeArrayDelegate) {
        ;
        ;
        function executeAsync(args) {
            try {
                if (args.onCalling) {
                    args.onCalling();
                }
                OSF._OfficeAppFactory.getClientHostController().execute(args.dispId, toArray(args.hostCallArgs), function OSF_DDA_SafeArrayFacade$Execute_OnResponse(hostResponseArgs, resultCode) {
                    var result;
                    var status;
                    if (typeof hostResponseArgs === "number") {
                        result = [];
                        status = hostResponseArgs;
                    }
                    else {
                        result = hostResponseArgs.toArray();
                        status = result[0];
                    }
                    if (status == 1) {
                        var payload = result[1];
                        payload = fromSafeArray(payload);
                        if (payload != null) {
                            if (!args._chunkResultData) {
                                args._chunkResultData = new Array();
                            }
                            args._chunkResultData[payload[0]] = payload[1];
                        }
                        return false;
                    }
                    if (args.onReceiving) {
                        args.onReceiving();
                    }
                    if (args.onComplete) {
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
                            if (args._chunkResultData) {
                                payload = fromSafeArray(payload);
                                if (payload != null) {
                                    var expectedChunkCount = payload[payload.length - 1];
                                    if (args._chunkResultData.length == expectedChunkCount) {
                                        payload[payload.length - 1] = args._chunkResultData;
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
                        args.onComplete(status, payload);
                    }
                    return true;
                });
            }
            catch (ex) {
                OSF.SafeArrayDelegate.onException(ex, args);
            }
        }
        SafeArrayDelegate.executeAsync = executeAsync;
        function registerEventAsync(args) {
            if (args.onCalling) {
                args.onCalling();
            }
            var callback = getOnAfterRegisterEvent(true, args);
            try {
                OSF._OfficeAppFactory.getClientHostController().registerEvent(args.dispId, undefined, args.targetId, function OSF_DDA_SafeArrayDelegate$RegisterEventAsync_OnEvent(eventDispId, payload) {
                    if (args.onEvent) {
                        args.onEvent(payload);
                    }
                    if (OSF.AppTelemetry) {
                        OSF.AppTelemetry.onEventDone(args.dispId);
                    }
                }, callback);
            }
            catch (ex) {
                OSF.SafeArrayDelegate.onException(ex, args);
            }
        }
        SafeArrayDelegate.registerEventAsync = registerEventAsync;
        function unregisterEventAsync(args) {
            if (args.onCalling) {
                args.onCalling();
            }
            var callback = getOnAfterRegisterEvent(false, args);
            try {
                OSF._OfficeAppFactory.getClientHostController().unregisterEvent(args.dispId, undefined, args.targetId, callback);
            }
            catch (ex) {
                OSF.SafeArrayDelegate.onException(ex, args);
            }
        }
        SafeArrayDelegate.unregisterEventAsync = unregisterEventAsync;
        function onException(ex, args) {
            var status;
            var statusNumber = ex.number;
            if (statusNumber) {
                switch (statusNumber) {
                    case -2146828218:
                        status = 7000;
                        break;
                    case -2147467259:
                        if (args.dispId == OSF.EventDispId.dispidDialogMessageReceivedEvent) {
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
            if (args.onComplete) {
                args.onComplete(status || 5001);
            }
        }
        SafeArrayDelegate.onException = onException;
        function onExceptionSyncMethod(ex, args) {
            var status;
            var number = ex.number;
            if (number) {
                switch (number) {
                    case -2146828218:
                        status = 7000;
                        break;
                    case -2146827850:
                    default:
                        status = 5001;
                        break;
                }
            }
            return status || 5001;
        }
        SafeArrayDelegate.onExceptionSyncMethod = onExceptionSyncMethod;
        function getOnAfterRegisterEvent(register, args) {
            var startTime = (new Date()).getTime();
            return function (hostResponseArgs) {
                if (args.onReceiving) {
                    args.onReceiving();
                }
                var status = hostResponseArgs.toArray ? hostResponseArgs.toArray()[0] : hostResponseArgs;
                if (args.onComplete) {
                    args.onComplete(status);
                }
                if (OSF.AppTelemetry) {
                    OSF.AppTelemetry.onRegisterDone(register, args.dispId, Math.abs((new Date()).getTime() - startTime), status);
                }
                return true;
            };
        }
        SafeArrayDelegate.getOnAfterRegisterEvent = getOnAfterRegisterEvent;
        function toArray(args) {
            var arrArgs = args;
            if (OSF.OUtil.isArray(args)) {
                var len = arrArgs.length;
                for (var i = 0; i < len; i++) {
                    arrArgs[i] = toArray(arrArgs[i]);
                }
            }
            else if (OSF.OUtil.isDate(args)) {
                arrArgs = args.getVarDate();
            }
            else if (typeof args === "object" && !OSF.OUtil.isArray(args)) {
                arrArgs = [];
                for (var index in args) {
                    if (!OSF.OUtil.isFunction(args[index])) {
                        arrArgs[index] = toArray(args[index]);
                    }
                }
            }
            return arrArgs;
        }
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
    })(SafeArrayDelegate = OSF.SafeArrayDelegate || (OSF.SafeArrayDelegate = {}));
})(OSF || (OSF = {}));
var OSF;
(function (OSF) {
    var SpecialProcessor = (function () {
        function SpecialProcessor(complexTypes, dynamicTypes) {
            this._complexTypes = complexTypes;
            this.dynamicTypes = dynamicTypes;
        }
        SpecialProcessor.prototype.addComplexType = function (complexType) {
            this._complexTypes.push(complexType);
        };
        SpecialProcessor.prototype.getDynamicType = function (dynamicType) {
            return this.dynamicTypes[dynamicType];
        };
        SpecialProcessor.prototype.setDynamicType = function (dynamicType, handler) {
            this.dynamicTypes[dynamicType] = handler;
        };
        SpecialProcessor.prototype.isComplexType = function (type) {
            return OSF.OUtil.listContainsValue(this._complexTypes, type);
        };
        SpecialProcessor.prototype.isDynamicType = function (type) {
            return OSF.OUtil.listContainsKey(this.dynamicTypes, type);
        };
        SpecialProcessor.prototype.preserveNesting = function (p) {
            return false;
        };
        SpecialProcessor.prototype.pack = function (type, arg) {
            var value;
            if (this.isDynamicType(type)) {
                value = this.dynamicTypes[type].toHost(arg);
            }
            else {
                value = arg;
            }
            return value;
        };
        SpecialProcessor.prototype.unpack = function (type, arg) {
            var value;
            if (this.isDynamicType(type)) {
                value = this.dynamicTypes[type].fromHost(arg);
            }
            else {
                value = arg;
            }
            return value;
        };
        return SpecialProcessor;
    }());
    OSF.SpecialProcessor = SpecialProcessor;
})(OSF || (OSF = {}));
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
var OSF;
(function (OSF) {
    var SafeArraySpecialProcessor = (function (_super) {
        __extends(SafeArraySpecialProcessor, _super);
        function SafeArraySpecialProcessor() {
            var _this = this;
            var tableRows = 0;
            var tableHeaders = 1;
            var complexTypes = [];
            var dynamicTypes = {
                data: {
                    toHost: function (data) {
                        if (typeof data != "string" && data["tableRows"] !== undefined) {
                            var tableData = [];
                            tableData[tableRows] = data["tableRows"];
                            tableData[tableHeaders] = data["tableHeaders"];
                            data = tableData;
                        }
                        return data;
                    },
                    fromHost: function (hostArgs) {
                        var ret;
                        if (hostArgs.toArray) {
                            var dimensions = hostArgs.dimensions();
                            if (dimensions === 2) {
                                ret = _this.twoDVBArrayToJaggedArray(hostArgs);
                            }
                            else {
                                var array = hostArgs.toArray();
                                if (array.length === 2 && ((array[0] != null && array[0].toArray) || (array[1] != null && array[1].toArray))) {
                                    ret = {};
                                    ret["tableRows"] = _this.twoDVBArrayToJaggedArray(array[tableRows]);
                                    ret["tableHeaders"] = _this.twoDVBArrayToJaggedArray(array[tableHeaders]);
                                }
                                else {
                                    ret = array;
                                }
                            }
                        }
                        else {
                            ret = hostArgs;
                        }
                        return ret;
                    }
                }
            };
            _this = _super.call(this, complexTypes, dynamicTypes) || this;
            return _this;
        }
        SafeArraySpecialProcessor.prototype.unpack = function (param, arg) {
            var value;
            if (this.isComplexType(param) || OSF.ListType.isListType(param)) {
                var toArraySupported = arg !== undefined && arg.toArray !== undefined;
                value = toArraySupported ? arg.toArray() : arg || {};
            }
            else if (this.isDynamicType(param)) {
                value = this.dynamicTypes[param].fromHost(arg);
            }
            else {
                value = arg;
            }
            return value;
        };
        SafeArraySpecialProcessor.prototype.twoDVBArrayToJaggedArray = function (vbArr) {
            var ret;
            try {
                var rows = vbArr.ubound(1);
                var cols = vbArr.ubound(2);
                vbArr = vbArr.toArray();
                if (rows == 1 && cols == 1) {
                    ret = [vbArr];
                }
                else {
                    ret = [];
                    for (var row = 0; row < rows; row++) {
                        var rowArr = [];
                        for (var col = 0; col < cols; col++) {
                            var datum = vbArr[row * cols + col];
                            if (datum != "{66e7831f-81b2-42e2-823c-89e872d541b3}") {
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
        };
        return SafeArraySpecialProcessor;
    }(OSF.SpecialProcessor));
    OSF.SafeArraySpecialProcessor = SafeArraySpecialProcessor;
})(OSF || (OSF = {}));
var OSF;
(function (OSF) {
    var SyncMethodCall = (function () {
        function SyncMethodCall(requiredParameters, supportedOptions, privateStateCallbacks, checkCallArgs, displayName) {
            this._requiredCount = requiredParameters.length;
            this._apiMethods = new OSF.ApiMethodCall(requiredParameters, supportedOptions, privateStateCallbacks, checkCallArgs, displayName);
            this._supportedOptions = supportedOptions;
        }
        SyncMethodCall.prototype.verifyAndExtractCall = function (userArgs, caller, stateInfo) {
            var required = this._apiMethods.extractRequiredArguments(userArgs, caller, stateInfo);
            var options = this.extractOptions(userArgs, required, caller, stateInfo);
            var callArgs = this._apiMethods.constructCallArgs(required, options, caller, stateInfo);
            return callArgs;
        };
        SyncMethodCall.prototype.extractOptions = function (userArgs, requiredArgs, caller, stateInfo) {
            if (userArgs.length > this._requiredCount + 1) {
                throw OSF.Utility.createParameterException(Strings.OfficeOM.L_TooManyArguments);
            }
            var options;
            for (var i = userArgs.length - 1; i >= this._requiredCount; i--) {
                var argument = userArgs[i];
                switch (typeof argument) {
                    case "object":
                        if (options) {
                            throw OSF.Utility.createParameterException(Strings.OfficeOM.L_TooManyArguments);
                        }
                        else {
                            options = argument;
                        }
                        break;
                    default:
                        throw OSF.Utility.createArgumentException(Strings.OfficeOM.L_InValidOptionalArgument);
                }
            }
            options = this._apiMethods.fillOptions(options, requiredArgs, caller, stateInfo);
            this._apiMethods.verifyArguments(this._supportedOptions, options);
            return options;
        };
        return SyncMethodCall;
    }());
    OSF.SyncMethodCall = SyncMethodCall;
})(OSF || (OSF = {}));
var OSF;
(function (OSF) {
    var SyncMethodCalls;
    (function (SyncMethodCalls) {
        var syncMethodCalls = {};
        function define(callDefinition) {
            syncMethodCalls[callDefinition.method] = manufacture(callDefinition);
        }
        SyncMethodCalls.define = define;
        function get(method) {
            return syncMethodCalls[method];
        }
        SyncMethodCalls.get = get;
        function manufacture(params) {
            var supportedOptions = params.supportedOptions ? OSF.OUtil.createObject(params.supportedOptions) : [];
            return new OSF.SyncMethodCall(params.requiredArguments || [], supportedOptions, params.privateStateCallbacks, params.checkCallArgs, params.method.displayName);
        }
    })(SyncMethodCalls = OSF.SyncMethodCalls || (OSF.SyncMethodCalls = {}));
})(OSF || (OSF = {}));
var OSF;
(function (OSF) {
    OSF.SyncMethods = {
        MessageParent: "messageParent",
        MessageChild: "messageChild",
        SendMessage: "sendMessage",
        AddMessageHandler: "addEventHandler"
    };
})(OSF || (OSF = {}));
var OSF;
(function (OSF) {
    var V10ApiFeatureRegistry;
    (function (V10ApiFeatureRegistry) {
        var apiFeatures = [];
        function initialize() {
            apiFeatures.forEach(function (apiFeature) {
                apiFeature.defineMethodsFunc();
                if (OSF.OUtil.getHostPlatform() == OSF.HostInfoPlatform.web) {
                    if (typeof apiFeature.defineWebParameterMapFunc == "function") {
                        apiFeature.defineWebParameterMapFunc();
                    }
                }
                else {
                    if (typeof apiFeature.defineSafeArrayParameterMapFunc == "function") {
                        apiFeature.defineSafeArrayParameterMapFunc();
                    }
                }
                if (typeof apiFeature.initializeFunc == "function") {
                    apiFeature.initializeFunc();
                }
            });
        }
        V10ApiFeatureRegistry.initialize = initialize;
        function register(apiFeature) {
            apiFeatures.push(apiFeature);
        }
        V10ApiFeatureRegistry.register = register;
    })(V10ApiFeatureRegistry = OSF.V10ApiFeatureRegistry || (OSF.V10ApiFeatureRegistry = {}));
})(OSF || (OSF = {}));
var OSF;
(function (OSF) {
    var WACDelegate;
    (function (WACDelegate) {
        WACDelegate.version = 1;
        var UniqueArguments;
        (function (UniqueArguments) {
            UniqueArguments["Data"] = "Data";
            UniqueArguments["Properties"] = "Properties";
            UniqueArguments["BindingRequest"] = "DdaBindingsMethod";
            UniqueArguments["BindingResponse"] = "Bindings";
            UniqueArguments["SingleBindingResponse"] = "singleBindingResponse";
            UniqueArguments["GetData"] = "DdaGetBindingData";
            UniqueArguments["AddRowsColumns"] = "DdaAddRowsColumns";
            UniqueArguments["SetData"] = "DdaSetBindingData";
            UniqueArguments["ClearFormats"] = "DdaClearBindingFormats";
            UniqueArguments["SetFormats"] = "DdaSetBindingFormats";
            UniqueArguments["SettingsRequest"] = "DdaSettingsMethod";
            UniqueArguments["BindingEventSource"] = "ddaBinding";
            UniqueArguments["ArrayData"] = "ArrayData";
        })(UniqueArguments = WACDelegate.UniqueArguments || (WACDelegate.UniqueArguments = {}));
        function executeAsync(args) {
            if (!args.hostCallArgs) {
                args.hostCallArgs = {};
            }
            args.hostCallArgs["DdaMethod"] = {
                "ControlId": OSF._OfficeAppFactory.getId(),
                "Version": OSF.WACDelegate.version,
                "DispatchId": args.dispId
            };
            args.hostCallArgs["__timeout__"] = -1;
            if (args.onCalling) {
                args.onCalling();
            }
            if (!OSF.getClientEndPoint()) {
                return;
            }
            OSF.getClientEndPoint().invoke("executeMethod", function (xdmStatus, payload) {
                if (args.onReceiving) {
                    args.onReceiving();
                }
                var error;
                if (xdmStatus == 0) {
                    OSF.WACDelegate.version = payload["Version"];
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
                if (args.onComplete) {
                    args.onComplete(error, payload);
                }
            }, args.hostCallArgs);
        }
        WACDelegate.executeAsync = executeAsync;
        function getOnAfterRegisterEvent(register, args) {
            var startTime = (new Date()).getTime();
            return function (xdmStatus, payload) {
                if (args.onReceiving) {
                    args.onReceiving();
                }
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
                if (args.onComplete) {
                    args.onComplete(status);
                }
                if (OSF.AppTelemetry) {
                    OSF.AppTelemetry.onRegisterDone(register, args.dispId, Math.abs((new Date()).getTime() - startTime), status);
                }
            };
        }
        WACDelegate.getOnAfterRegisterEvent = getOnAfterRegisterEvent;
        function registerEventAsync(args) {
            if (args.onCalling) {
                args.onCalling();
            }
            if (!OSF.getClientEndPoint()) {
                return;
            }
            OSF.getClientEndPoint().registerForEvent(OSF.OUtil.getXdmEventName(args.targetId, args.eventType), function OSF_DDA_WACOMFacade$OnEvent(payload) {
                if (args.onEvent) {
                    args.onEvent(payload);
                }
                if (OSF.AppTelemetry) {
                    OSF.AppTelemetry.onEventDone(args.dispId);
                }
            }, getOnAfterRegisterEvent(true, args), {
                "controlId": OSF._OfficeAppFactory.getId(),
                "eventDispId": args.dispId,
                "targetId": args.targetId
            });
        }
        WACDelegate.registerEventAsync = registerEventAsync;
        function unregisterEventAsync(args) {
            if (args.onCalling) {
                args.onCalling();
            }
            if (!OSF.getClientEndPoint()) {
                return;
            }
            OSF.getClientEndPoint().unregisterForEvent(OSF.OUtil.getXdmEventName(args.targetId, args.eventType), getOnAfterRegisterEvent(false, args), {
                "controlId": OSF._OfficeAppFactory.getId(),
                "eventDispId": args.dispId,
                "targetId": args.targetId
            });
        }
        WACDelegate.unregisterEventAsync = unregisterEventAsync;
    })(WACDelegate = OSF.WACDelegate || (OSF.WACDelegate = {}));
})(OSF || (OSF = {}));
var OSF;
(function (OSF) {
    var WebSpecialProcessor = (function (_super) {
        __extends(WebSpecialProcessor, _super);
        function WebSpecialProcessor() {
            var _this = this;
            var complexTypes = [
                OSF.WACDelegate.UniqueArguments.SingleBindingResponse,
                OSF.WACDelegate.UniqueArguments.BindingRequest,
                OSF.WACDelegate.UniqueArguments.BindingResponse,
                OSF.WACDelegate.UniqueArguments.GetData,
                OSF.WACDelegate.UniqueArguments.AddRowsColumns,
                OSF.WACDelegate.UniqueArguments.SetData,
                OSF.WACDelegate.UniqueArguments.ClearFormats,
                OSF.WACDelegate.UniqueArguments.SetFormats,
                OSF.WACDelegate.UniqueArguments.SettingsRequest,
                OSF.WACDelegate.UniqueArguments.BindingEventSource
            ];
            var dynamicTypes = {};
            _this = _super.call(this, complexTypes, dynamicTypes) || this;
            return _this;
        }
        return WebSpecialProcessor;
    }(OSF.SpecialProcessor));
    OSF.WebSpecialProcessor = WebSpecialProcessor;
})(OSF || (OSF = {}));
var OfficeExt;
(function (OfficeExt) {
    var Container = (function () {
        function Container(parameters) {
        }
        return Container;
    }());
    OfficeExt.Container = Container;
})(OfficeExt || (OfficeExt = {}));
var OSF;
(function (OSF) {
    var Container;
    (function (Container) {
        function defineMethods() {
            OSF.AsyncMethodCalls.define({
                method: OSF.AsyncMethods.CloseContainerAsync,
                requiredArguments: [],
                supportedOptions: [],
                privateStateCallbacks: []
            });
        }
        function defineSafeArrayParameterMap() {
            OSF.HostParameterMap.define({
                type: 97,
                fromHost: [],
                toHost: []
            });
        }
        function defineWebParameterMap() {
            OSF.HostParameterMap.define({
                type: 97,
                fromHost: [],
                toHost: []
            });
        }
        function initialize() {
            var target = Office.context.ui;
            if (!OSF.OUtil.isDialog()) {
                if (OfficeExt.Container) {
                    OSF.DispIdHost.addAsyncMethods(target, [OSF.AsyncMethods.CloseContainerAsync]);
                }
            }
        }
        OSF.V10ApiFeatureRegistry.register({
            defineMethodsFunc: defineMethods,
            defineSafeArrayParameterMapFunc: defineSafeArrayParameterMap,
            defineWebParameterMapFunc: defineWebParameterMap,
            initializeFunc: initialize
        });
    })(Container || (Container = {}));
})(OSF || (OSF = {}));
var OSF;
(function (OSF) {
    var Marshaling;
    (function (Marshaling) {
        var DialogMessageReceivedEventKeys;
        (function (DialogMessageReceivedEventKeys) {
            DialogMessageReceivedEventKeys["MessageType"] = "messageType";
            DialogMessageReceivedEventKeys["MessageContent"] = "messageContent";
        })(DialogMessageReceivedEventKeys = Marshaling.DialogMessageReceivedEventKeys || (Marshaling.DialogMessageReceivedEventKeys = {}));
        ;
        var DialogParentMessageReceivedEventKeys;
        (function (DialogParentMessageReceivedEventKeys) {
            DialogParentMessageReceivedEventKeys["MessageType"] = "messageType";
            DialogParentMessageReceivedEventKeys["MessageContent"] = "messageContent";
        })(DialogParentMessageReceivedEventKeys = Marshaling.DialogParentMessageReceivedEventKeys || (Marshaling.DialogParentMessageReceivedEventKeys = {}));
        ;
        var MessageParentKeys;
        (function (MessageParentKeys) {
            MessageParentKeys["MessageToParent"] = "messageToParent";
        })(MessageParentKeys = Marshaling.MessageParentKeys || (Marshaling.MessageParentKeys = {}));
        ;
        var DialogNotificationShownEventType;
        (function (DialogNotificationShownEventType) {
            DialogNotificationShownEventType["DialogNotificationShown"] = "dialogNotificationShown";
        })(DialogNotificationShownEventType = Marshaling.DialogNotificationShownEventType || (Marshaling.DialogNotificationShownEventType = {}));
        ;
        var SendMessageKeys;
        (function (SendMessageKeys) {
            SendMessageKeys["MessageContent"] = "messageContent";
        })(SendMessageKeys = Marshaling.SendMessageKeys || (Marshaling.SendMessageKeys = {}));
        ;
    })(Marshaling = OSF.Marshaling || (OSF.Marshaling = {}));
})(OSF || (OSF = {}));
var OSF;
(function (OSF) {
    var EventDescriptors;
    (function (EventDescriptors) {
        EventDescriptors.DialogParentMessageReceivedEvent = "DialogParentMessageReceivedEvent";
        EventDescriptors.DialogMessageReceivedEvent = "DialogMessageReceivedEvent";
    })(EventDescriptors = OSF.EventDescriptors || (OSF.EventDescriptors = {}));
    OSF.DialogParentMessageEventDispatch = new OSF.EventDispatch([
        OSF.EventType.DialogParentMessageReceived,
        OSF.EventType.DialogParentEventReceived
    ]);
    var Dialog;
    (function (Dialog) {
        function defineMethods() {
            OSF.AsyncMethodCalls.define({
                method: OSF.AsyncMethods.DisplayDialogAsync,
                requiredArguments: [
                    {
                        "name": OSF.ParameterNames.Url,
                        "types": ["string"]
                    }
                ],
                supportedOptions: [
                    {
                        name: OSF.ParameterNames.Width,
                        value: {
                            "types": ["number"],
                            "defaultValue": 99
                        }
                    },
                    {
                        name: OSF.ParameterNames.Height,
                        value: {
                            "types": ["number"],
                            "defaultValue": 99
                        }
                    },
                    {
                        name: OSF.ParameterNames.RequireHTTPs,
                        value: {
                            "types": ["boolean"],
                            "defaultValue": true
                        }
                    },
                    {
                        name: OSF.ParameterNames.DisplayInIframe,
                        value: {
                            "types": ["boolean"],
                            "defaultValue": false
                        }
                    },
                    {
                        name: OSF.ParameterNames.HideTitle,
                        value: {
                            "types": ["boolean"],
                            "defaultValue": false
                        }
                    },
                    {
                        name: OSF.ParameterNames.UseDeviceIndependentPixels,
                        value: {
                            "types": ["boolean"],
                            "defaultValue": false
                        }
                    },
                    {
                        name: OSF.ParameterNames.PromptBeforeOpen,
                        value: {
                            "types": ["boolean"],
                            "defaultValue": true
                        }
                    },
                    {
                        name: OSF.ParameterNames.EnforceAppDomain,
                        value: {
                            "types": ["boolean"],
                            "defaultValue": true
                        }
                    },
                    {
                        name: OSF.ParameterNames.UrlNoHostInfo,
                        value: {
                            "types": ["boolean"],
                            "defaultValue": false
                        }
                    }
                ],
                privateStateCallbacks: [],
                onSucceeded: function (args, caller, callArgs) {
                    var targetId = args[OSF.ParameterNames.Id];
                    var eventDispatch = args[OSF.ParameterNames.Data];
                    var dialog = {};
                    var closeDialog = OSF.AsyncMethods.CloseAsync;
                    OSF.OUtil.defineEnumerableProperty(dialog, closeDialog, {
                        value: function () {
                            var closeDialogfunction = OSF.DispIdHost.CloseDialog;
                            closeDialogfunction(arguments, targetId, eventDispatch, dialog);
                        }
                    });
                    var addHandler = OSF.SyncMethods.AddMessageHandler;
                    OSF.OUtil.defineEnumerableProperty(dialog, addHandler, {
                        value: function () {
                            var syncMethodCall = OSF.SyncMethodCalls.get(OSF.SyncMethods.AddMessageHandler);
                            var callArgs = syncMethodCall.verifyAndExtractCall(arguments, dialog, eventDispatch);
                            var eventType = callArgs[OSF.ParameterNames.EventType];
                            var handler = callArgs[OSF.ParameterNames.Handler];
                            return eventDispatch.addEventHandlerAndFireQueuedEvent(eventType, handler);
                        }
                    });
                    if (OSF.EnableSendMessageDialogAPI === true) {
                        var sendMessage = OSF.SyncMethods.SendMessage;
                        OSF.OUtil.defineEnumerableProperty(dialog, sendMessage, {
                            value: function () {
                                var execute = OSF.DispIdHost.SendMessage;
                                return execute(arguments, eventDispatch, dialog);
                            }
                        });
                    }
                    if (OSF.EnableMessageChildDialogAPI === true) {
                        var messageChild = OSF.SyncMethods.MessageChild;
                        OSF.OUtil.defineEnumerableProperty(dialog, messageChild, {
                            value: function () {
                                var execute = OSF.DispIdHost.SendMessage;
                                return execute(arguments, eventDispatch, dialog);
                            }
                        });
                    }
                    return dialog;
                },
                checkCallArgs: function (callArgs, caller, stateInfo) {
                    if (callArgs[OSF.ParameterNames.Width] <= 0) {
                        callArgs[OSF.ParameterNames.Width] = 1;
                    }
                    if (!callArgs[OSF.ParameterNames.UseDeviceIndependentPixels] && callArgs[OSF.ParameterNames.Width] > 100) {
                        callArgs[OSF.ParameterNames.Width] = 99;
                    }
                    if (callArgs[OSF.ParameterNames.Height] <= 0) {
                        callArgs[OSF.ParameterNames.Height] = 1;
                    }
                    if (!callArgs[OSF.ParameterNames.UseDeviceIndependentPixels] && callArgs[OSF.ParameterNames.Height] > 100) {
                        callArgs[OSF.ParameterNames.Height] = 99;
                    }
                    if (!callArgs[OSF.ParameterNames.RequireHTTPs]) {
                        callArgs[OSF.ParameterNames.RequireHTTPs] = true;
                    }
                    return callArgs;
                }
            });
            OSF.AsyncMethodCalls.define({
                method: OSF.AsyncMethods.CloseAsync,
                requiredArguments: [],
                supportedOptions: [],
                privateStateCallbacks: []
            });
            OSF.SyncMethodCalls.define({
                method: OSF.SyncMethods.MessageParent,
                requiredArguments: [
                    {
                        "name": OSF.ParameterNames.MessageToParent,
                        "types": ["string", "number", "boolean"]
                    }
                ],
                supportedOptions: []
            });
            OSF.SyncMethodCalls.define({
                method: OSF.SyncMethods.AddMessageHandler,
                requiredArguments: [
                    {
                        "name": OSF.ParameterNames.EventType,
                        "enum": OSF.EventType,
                        "verify": function (eventType, caller, eventDispatch) { return eventDispatch.supportsEvent(eventType); }
                    },
                    {
                        "name": OSF.ParameterNames.Handler,
                        "types": ["function"]
                    }
                ],
                supportedOptions: []
            });
            OSF.SyncMethodCalls.define({
                method: OSF.SyncMethods.SendMessage,
                requiredArguments: [
                    {
                        "name": OSF.ParameterNames.MessageContent,
                        "types": ["string"]
                    }
                ],
                supportedOptions: [],
                privateStateCallbacks: []
            });
        }
        function defineSafeArrayParameterMap() {
            OSF.HostParameterMap.define({
                type: OSF.EventDispId.dispidDialogMessageReceivedEvent,
                fromHost: [
                    { name: OSF.EventDescriptors.DialogMessageReceivedEvent, value: OSF.HostParameterMap.self }
                ],
                isComplexType: true
            });
            OSF.HostParameterMap.define({
                type: OSF.EventDescriptors.DialogMessageReceivedEvent,
                fromHost: [
                    { name: OSF.PropertyDescriptors.MessageType, value: 0 },
                    { name: OSF.PropertyDescriptors.MessageContent, value: 1 }
                ],
                isComplexType: true
            });
            OSF.HostParameterMap.define({
                type: OSF.EventDispId.dispidDialogParentMessageReceivedEvent,
                fromHost: [
                    { name: OSF.EventDescriptors.DialogParentMessageReceivedEvent, value: OSF.HostParameterMap.self }
                ],
                isComplexType: true
            });
            OSF.HostParameterMap.define({
                type: OSF.EventDescriptors.DialogParentMessageReceivedEvent,
                fromHost: [
                    { name: OSF.PropertyDescriptors.MessageType, value: 0 },
                    { name: OSF.PropertyDescriptors.MessageContent, value: 1 }
                ],
                isComplexType: true
            });
        }
        function defineWebParameterMap() {
            OSF.HostParameterMap.define({
                type: OSF.EventDispId.dispidDialogMessageReceivedEvent,
                fromHost: [
                    { name: OSF.EventDescriptors.DialogMessageReceivedEvent, value: OSF.HostParameterMap.self }
                ]
            });
            OSF.HostParameterMap.addComplexType(OSF.EventDescriptors.DialogMessageReceivedEvent);
            OSF.HostParameterMap.define({
                type: OSF.EventDescriptors.DialogMessageReceivedEvent,
                fromHost: [
                    { name: OSF.PropertyDescriptors.MessageType, value: OSF.Marshaling.DialogMessageReceivedEventKeys.MessageType },
                    { name: OSF.PropertyDescriptors.MessageContent, value: OSF.Marshaling.DialogMessageReceivedEventKeys.MessageContent }
                ]
            });
            OSF.HostParameterMap.define({
                type: OSF.EventDispId.dispidDialogParentMessageReceivedEvent,
                fromHost: [
                    { name: OSF.EventDescriptors.DialogParentMessageReceivedEvent, value: OSF.HostParameterMap.self }
                ]
            });
            OSF.HostParameterMap.addComplexType(OSF.EventDescriptors.DialogParentMessageReceivedEvent);
            OSF.HostParameterMap.define({
                type: OSF.EventDescriptors.DialogParentMessageReceivedEvent,
                fromHost: [
                    { name: OSF.PropertyDescriptors.MessageType, value: OSF.Marshaling.DialogParentMessageReceivedEventKeys.MessageType },
                    { name: OSF.PropertyDescriptors.MessageContent, value: OSF.Marshaling.DialogParentMessageReceivedEventKeys.MessageContent }
                ]
            });
            OSF.HostParameterMap.define({
                type: 144,
                toHost: [
                    { name: OSF.ParameterNames.MessageToParent, value: OSF.Marshaling.MessageParentKeys.MessageToParent }
                ]
            });
            OSF.HostParameterMap.define({
                type: 145,
                toHost: [
                    { name: OSF.ParameterNames.MessageContent, value: OSF.Marshaling.SendMessageKeys.MessageContent }
                ]
            });
        }
        function initialize() {
            var isPopupWindow = OSF.OUtil.isPopupWindow();
            OSF.EnableMessageChildDialogAPI = true;
            var hostInfo = OSF._OfficeAppFactory.getHostInfo();
            if (hostInfo.hostType == "onenote") {
                OSF.EnableSendMessageDialogAPI = false;
            }
            else {
                OSF.EnableSendMessageDialogAPI = true;
            }
            var target = Office.context.ui;
            if (OSF.OUtil.isDialog()) {
                var messageParentName = OSF.SyncMethods.MessageParent;
                if (!target[messageParentName]) {
                    OSF.OUtil.defineEnumerableProperty(target, messageParentName, {
                        value: function () {
                            var messageParent = OSF.DispIdHost.MessageParent;
                            return messageParent(arguments, target);
                        }
                    });
                }
                var addEventHandler = OSF.SyncMethods.AddMessageHandler;
                if (!target[addEventHandler] && typeof OSF.DialogParentMessageEventDispatch != "undefined") {
                    OSF.DispIdHost.addEventSupport(target, OSF.DialogParentMessageEventDispatch, isPopupWindow);
                }
                if (isPopupWindow) {
                    OSF.WacDialogAction.registerMessageReceivedEvent();
                }
            }
            else {
                var eventDispatch;
                if (OSF.EventType.DialogParentMessageReceived != null) {
                    eventDispatch = new OSF.EventDispatch([
                        OSF.EventType.DialogMessageReceived,
                        OSF.EventType.DialogEventReceived,
                        OSF.EventType.DialogParentMessageReceived
                    ]);
                }
                else {
                    eventDispatch = new OSF.EventDispatch([
                        OSF.EventType.DialogMessageReceived,
                        OSF.EventType.DialogEventReceived
                    ]);
                }
                var openDialogName = OSF.AsyncMethods.DisplayDialogAsync;
                if (!target[openDialogName]) {
                    OSF.OUtil.defineEnumerableProperty(target, openDialogName, {
                        value: function () {
                            var openDialog = OSF.DispIdHost.OpenDialog;
                            openDialog(arguments, eventDispatch, target);
                        }
                    });
                }
            }
            OSF.OUtil.finalizeProperties(target);
        }
        OSF.V10ApiFeatureRegistry.register({
            defineMethodsFunc: defineMethods,
            defineSafeArrayParameterMapFunc: defineSafeArrayParameterMap,
            defineWebParameterMapFunc: defineWebParameterMap,
            initializeFunc: initialize
        });
    })(Dialog || (Dialog = {}));
})(OSF || (OSF = {}));
var OSF;
(function (OSF) {
    var ShowWindowDialogParameterKeys;
    (function (ShowWindowDialogParameterKeys) {
        ShowWindowDialogParameterKeys["Url"] = "url";
        ShowWindowDialogParameterKeys["Width"] = "width";
        ShowWindowDialogParameterKeys["Height"] = "height";
        ShowWindowDialogParameterKeys["DisplayInIframe"] = "displayInIframe";
        ShowWindowDialogParameterKeys["HideTitle"] = "hideTitle";
        ShowWindowDialogParameterKeys["UseDeviceIndependentPixels"] = "useDeviceIndependentPixels";
        ShowWindowDialogParameterKeys["PromptBeforeOpen"] = "promptBeforeOpen";
        ShowWindowDialogParameterKeys["EnforceAppDomain"] = "enforceAppDomain";
        ShowWindowDialogParameterKeys["UrlNoHostInfo"] = "urlNoHostInfo";
    })(ShowWindowDialogParameterKeys = OSF.ShowWindowDialogParameterKeys || (OSF.ShowWindowDialogParameterKeys = {}));
    var WacCommonUICssManager;
    (function (WacCommonUICssManager) {
        var hostType = {
            Excel: "excel",
            Word: "word",
            PowerPoint: "powerpoint",
            Outlook: "outlook",
            Visio: "visio"
        };
        function getDialogCssManager(applicationHostType) {
            switch (applicationHostType) {
                case hostType.Excel:
                case hostType.Word:
                case hostType.PowerPoint:
                case hostType.Outlook:
                case hostType.Visio:
                    return new DefaultDialogCSSManager();
                default:
                    return new DefaultDialogCSSManager();
            }
        }
        WacCommonUICssManager.getDialogCssManager = getDialogCssManager;
        var DefaultDialogCSSManager = (function () {
            function DefaultDialogCSSManager() {
                this.overlayElementCSS = [
                    "position: absolute",
                    "top: 0",
                    "left: 0",
                    "width: 100%",
                    "height: 100%",
                    "background-color: rgba(198, 198, 198, 0.5)",
                    "z-index: 99998"
                ];
                this.dialogNotificationPanelCSS = [
                    "width: 100%",
                    "height: 190px",
                    "position: absolute",
                    "z-index: 99999",
                    "background-color: rgba(255, 255, 255, 1)",
                    "left: 0px",
                    "top: 50%",
                    "margin-top: -95px"
                ];
                this.newWindowNotificationTextPanelCSS = [
                    "margin: 20px 14px",
                    "font-family: Segoe UI,Arial,Verdana,sans-serif",
                    "font-size: 14px",
                    "height: 100px",
                    "line-height: 100px"
                ];
                this.newWindowNotificationTextSpanCSS = [
                    "display: inline-block",
                    "line-height: normal",
                    "vertical-align: middle"
                ];
                this.crossZoneNotificationTextPanelCSS = [
                    "margin: 20px 14px",
                    "font-family: Segoe UI,Arial,Verdana,sans-serif",
                    "font-size: 14px",
                    "height: 100px",
                ];
                this.dialogNotificationButtonPanelCSS = "margin:0px 9px";
                this.buttonStyleCSS = [
                    "text-align: center",
                    "width: 70px",
                    "height: 25px",
                    "font-size: 14px",
                    "font-family: Segoe UI,Arial,Verdana,sans-serif",
                    "margin: 0px 5px",
                    "border-width: 1px",
                    "border-style: solid"
                ];
            }
            DefaultDialogCSSManager.prototype.getOverlayElementCSS = function () {
                return this.overlayElementCSS.join(";");
            };
            DefaultDialogCSSManager.prototype.getDialogNotificationPanelCSS = function () {
                return this.dialogNotificationPanelCSS.join(";");
            };
            DefaultDialogCSSManager.prototype.getNewWindowNotificationTextPanelCSS = function () {
                return this.newWindowNotificationTextPanelCSS.join(";");
            };
            DefaultDialogCSSManager.prototype.getNewWindowNotificationTextSpanCSS = function () {
                return this.newWindowNotificationTextSpanCSS.join(";");
            };
            DefaultDialogCSSManager.prototype.getCrossZoneNotificationTextPanelCSS = function () {
                return this.crossZoneNotificationTextPanelCSS.join(";");
            };
            DefaultDialogCSSManager.prototype.getDialogNotificationButtonPanelCSS = function () {
                return this.dialogNotificationButtonPanelCSS;
            };
            DefaultDialogCSSManager.prototype.getDialogButtonCSS = function () {
                return this.buttonStyleCSS.join(";");
            };
            return DefaultDialogCSSManager;
        }());
        WacCommonUICssManager.DefaultDialogCSSManager = DefaultDialogCSSManager;
    })(WacCommonUICssManager = OSF.WacCommonUICssManager || (OSF.WacCommonUICssManager = {}));
    var WacDialogAction;
    (function (WacDialogAction) {
        var windowInstance = null;
        var handler = null;
        var overlayElement = null;
        var dialogNotificationPanel = null;
        var closeDialogKey = "osfDialogInternal:action=closeDialog";
        var showDialogCallback = null;
        var hasCrossZoneNotification = false;
        var checkWindowDialogCloseInterval = -1;
        var messageParentKey = "messageParentKey";
        var hostThemeButtonStyle = null;
        var commonButtonBorderColor = "#ababab";
        var commonButtonBackgroundColor = "#ffffff";
        var commonEventInButtonBackgroundColor = "#ccc";
        var newWindowNotificationId = "newWindowNotificaiton";
        var crossZoneNotificationId = "crossZoneNotification";
        var configureBrowserLinkId = "configureBrowserLink";
        var dialogNotificationTextPanelId = "dialogNotificationTextPanel";
        var registerDialogNotificationShownArgs = {
            "dispId": OSF.EventDispId.dispidDialogNotificationShownInAddinEvent,
            "eventType": OSF.Marshaling.DialogNotificationShownEventType.DialogNotificationShown,
            "onComplete": null,
            "onCalling": null
        };
        function setHostThemeButtonStyle(args) {
            var hostThemeButtonStyleArgs = args.input;
            if (hostThemeButtonStyleArgs != null) {
                hostThemeButtonStyle = {
                    HostButtonBorderColor: hostThemeButtonStyleArgs[OSF.OUtil.HostThemeButtonStyleKeys.ButtonBorderColor],
                    HostButtonBackgroundColor: hostThemeButtonStyleArgs[OSF.OUtil.HostThemeButtonStyleKeys.ButtonBackgroundColor]
                };
            }
            args.completed();
        }
        WacDialogAction.setHostThemeButtonStyle = setHostThemeButtonStyle;
        function removeEventListenersForDialog(args) {
            addOrRemoveEventListenersForWindow(false);
            args.completed();
        }
        WacDialogAction.removeEventListenersForDialog = removeEventListenersForDialog;
        function handleNewWindowDialog(dialogInfo) {
            try {
                if (!checkAppDomain(dialogInfo)) {
                    showDialogCallback(12004);
                    return;
                }
                if (!dialogInfo[OSF.ShowWindowDialogParameterKeys.PromptBeforeOpen]) {
                    showDialog(dialogInfo);
                    return;
                }
                hasCrossZoneNotification = false;
                var ignoreButtonKeyDownClick = false;
                var hostInfoObj = OSF._OfficeAppFactory.getHostInfo();
                var dialogCssManager = OSF.WacCommonUICssManager.getDialogCssManager(hostInfoObj.hostType);
                var notificationText = OSF.OUtil.formatString(Strings.OfficeOM.L_ShowWindowDialogNotification, OSF._OfficeAppFactory.getOfficeAppContext().get_addinName());
                overlayElement = createOverlayElement(dialogCssManager);
                var docBodyChildren = removeAndStoreAllChildrenFromNode(document.body);
                document.body.appendChild(overlayElement);
                dialogNotificationPanel = createNotificationPanel(dialogCssManager, notificationText);
                dialogNotificationPanel.id = newWindowNotificationId;
                var dialogNotificationButtonPanel = createButtonPanel(dialogCssManager);
                var allowButton = createButtonControl(dialogCssManager, Strings.OfficeOM.L_ShowWindowDialogNotificationAllow);
                var ignoreButton = createButtonControl(dialogCssManager, Strings.OfficeOM.L_ShowWindowDialogNotificationIgnore);
                dialogNotificationButtonPanel.appendChild(allowButton);
                dialogNotificationButtonPanel.appendChild(ignoreButton);
                dialogNotificationPanel.appendChild(dialogNotificationButtonPanel);
                document.body.insertBefore(dialogNotificationPanel, document.body.firstChild);
                allowButton.onclick = function (event) {
                    showDialog(dialogInfo);
                    if (!hasCrossZoneNotification) {
                        dismissDialogNotification();
                    }
                    restoreChildrenToNode(document.body, docBodyChildren);
                    docBodyChildren = [];
                    event.preventDefault();
                    event.stopPropagation();
                };
                function removeAndStoreAllChildrenFromNode(node) {
                    var children = [];
                    try {
                        while (node.firstChild && node.firstChild != null) {
                            children.push(node.firstChild);
                            node.removeChild(node.firstChild);
                        }
                    }
                    catch (e) { }
                    return children;
                }
                function restoreChildrenToNode(node, children) {
                    try {
                        for (var i = 0; i < children.length; i++) {
                            node.appendChild(children[i]);
                        }
                    }
                    catch (e) { }
                }
                function ignoreButtonClickEventHandler(event) {
                    function unregisterDialogNotificationShownEventCallback(status) {
                        removeDialogNotificationElement();
                        setFocusOnFirstElement(status);
                        showDialogCallback(12009);
                    }
                    registerDialogNotificationShownArgs.onCalling = unregisterDialogNotificationShownEventCallback;
                    OSF.WACDelegate.unregisterEventAsync(registerDialogNotificationShownArgs);
                    restoreChildrenToNode(document.body, docBodyChildren);
                    docBodyChildren = [];
                    event.preventDefault();
                    event.stopPropagation();
                }
                ignoreButton.onclick = ignoreButtonClickEventHandler;
                allowButton.addEventListener("keydown", function (event) {
                    if (event.shiftKey && event.keyCode == 9) {
                        handleButtonControlEventOut(allowButton);
                        handleButtonControlEventIn(ignoreButton);
                        ignoreButton.focus();
                        event.preventDefault();
                        event.stopPropagation();
                    }
                }, false);
                ignoreButton.addEventListener("keydown", function (event) {
                    if (!event.shiftKey && event.keyCode == 9) {
                        handleButtonControlEventOut(ignoreButton);
                        handleButtonControlEventIn(allowButton);
                        allowButton.focus();
                        event.preventDefault();
                        event.stopPropagation();
                    }
                    else if (event.keyCode == 13) {
                        ignoreButtonKeyDownClick = true;
                        event.preventDefault();
                        event.stopPropagation();
                    }
                }, false);
                ignoreButton.addEventListener("keyup", function (event) {
                    if (event.keyCode == 13 && ignoreButtonKeyDownClick) {
                        ignoreButtonKeyDownClick = false;
                        ignoreButtonClickEventHandler(event);
                    }
                }, false);
                window.focus();
                function registerDialogNotificationShownEventCallback(status) {
                    allowButton.focus();
                }
                registerDialogNotificationShownArgs.onCalling = registerDialogNotificationShownEventCallback;
                OSF.WACDelegate.registerEventAsync(registerDialogNotificationShownArgs);
            }
            catch (e) {
                if (OSF.AppTelemetry) {
                    OSF.AppTelemetry.logAppException("Exception happens at new window dialog." + e);
                }
                showDialogCallback(5001);
            }
        }
        WacDialogAction.handleNewWindowDialog = handleNewWindowDialog;
        function closeDialog(callback) {
            try {
                if (windowInstance != null) {
                    var appDomains = OSF._OfficeAppFactory.getOfficeAppContext().get_appDomains();
                    if (appDomains) {
                        for (var i = 0; i < appDomains.length && appDomains[i].indexOf("://") !== -1; i++) {
                            windowInstance.postMessage(closeDialogKey, appDomains[i]);
                        }
                    }
                    if (windowInstance != null && !windowInstance.closed) {
                        windowInstance.close();
                    }
                    if (OSF.OUtil.shouldUseLocalStorageToPassMessage()) {
                        window.removeEventListener("storage", storageChangedHandler);
                    }
                    else {
                        window.removeEventListener("message", receiveMessage);
                    }
                    window.clearInterval(checkWindowDialogCloseInterval);
                    windowInstance = null;
                    callback(0);
                }
                else {
                    callback(5001);
                }
            }
            catch (e) {
                if (OSF.AppTelemetry) {
                    OSF.AppTelemetry.logAppException("Exception happens at close window dialog." + e);
                }
                callback(5001);
            }
        }
        WacDialogAction.closeDialog = closeDialog;
        function messageParent(params) {
            var message = params.hostCallArgs[OSF.ParameterNames.MessageToParent];
            if (OSF.OUtil.shouldUseLocalStorageToPassMessage()) {
                try {
                    var messageKey = OSF._OfficeAppFactory.getId() + messageParentKey;
                    window.localStorage.setItem(messageKey, message);
                }
                catch (e) {
                    if (OSF.AppTelemetry) {
                        OSF.AppTelemetry.logAppException("Error happened during messageParent method:" + e);
                    }
                }
            }
            else {
                postDialogMessage(window.opener, message);
            }
        }
        WacDialogAction.messageParent = messageParent;
        function sendMessage(params) {
            if (windowInstance != null) {
                var message = params.hostCallArgs;
                if (typeof message != "string") {
                    message = JSON.stringify(message);
                }
                postDialogMessage(windowInstance, message);
            }
        }
        WacDialogAction.sendMessage = sendMessage;
        function postDialogMessage(targetWindow, message) {
            var appDomains = OSF._OfficeAppFactory.getOfficeAppContext().get_appDomains();
            var currentOrigin = window.location.origin;
            if (!currentOrigin) {
                currentOrigin = window.location.protocol + "//"
                    + window.location.hostname
                    + (window.location.port ? ':' + window.location.port : '');
            }
            if (appDomains) {
                for (var i = 0; i < appDomains.length && appDomains[i].indexOf("://") !== -1; i++) {
                    targetWindow.postMessage(message, appDomains[i]);
                }
            }
            if (!OSF.XdmCommunicationManager.checkUrlWithAppDomains(appDomains, currentOrigin)) {
                targetWindow.postMessage(message, currentOrigin);
            }
        }
        WacDialogAction.postDialogMessage = postDialogMessage;
        function registerMessageReceivedEvent() {
            function receiveCloseDialogMessage(event) {
                if (event.source == window.opener) {
                    if (typeof event.data === "string" && event.data.indexOf(closeDialogKey) > -1) {
                        window.close();
                    }
                    else {
                        var messageContent = event.data, type = typeof messageContent;
                        if (messageContent && (type == "object" || type == "string")) {
                            if (type == "string") {
                                messageContent = JSON.parse(messageContent);
                            }
                            var eventArgs = OSF.manufactureEventArgs(OSF.EventType.DialogParentMessageReceived, null, messageContent);
                            OSF.DialogParentMessageEventDispatch.fireEvent(eventArgs);
                        }
                    }
                }
            }
            window.addEventListener("message", receiveCloseDialogMessage);
        }
        WacDialogAction.registerMessageReceivedEvent = registerMessageReceivedEvent;
        function setHandlerAndShowDialogCallback(onEventHandler, callback) {
            handler = onEventHandler;
            showDialogCallback = callback;
        }
        WacDialogAction.setHandlerAndShowDialogCallback = setHandlerAndShowDialogCallback;
        function escDismissDialogNotification() {
            try {
                if (dialogNotificationPanel && (dialogNotificationPanel.id == newWindowNotificationId) && showDialogCallback) {
                    showDialogCallback(12009);
                }
            }
            catch (e) {
                if (OSF.AppTelemetry) {
                    OSF.AppTelemetry.logAppException("Error happened during executing displayDialogAsync callback." + e);
                }
            }
            dismissDialogNotification();
        }
        WacDialogAction.escDismissDialogNotification = escDismissDialogNotification;
        function showCrossZoneNotification(windowUrl, hostType) {
            var okButtonKeyDownClick = false;
            var dialogCssManager = OSF.WacCommonUICssManager.getDialogCssManager(hostType);
            overlayElement = createOverlayElement(dialogCssManager);
            document.body.insertBefore(overlayElement, document.body.firstChild);
            dialogNotificationPanel = createNotificationPanelForCrossZoneIssue(dialogCssManager, windowUrl);
            dialogNotificationPanel.id = crossZoneNotificationId;
            var dialogNotificationButtonPanel = createButtonPanel(dialogCssManager);
            var okButton = createButtonControl(dialogCssManager, Strings.OfficeOM.L_DialogOK ? Strings.OfficeOM.L_DialogOK : "OK");
            dialogNotificationButtonPanel.appendChild(okButton);
            dialogNotificationPanel.appendChild(dialogNotificationButtonPanel);
            document.body.insertBefore(dialogNotificationPanel, document.body.firstChild);
            hasCrossZoneNotification = true;
            okButton.onclick = function () {
                dismissDialogNotification();
            };
            okButton.addEventListener("keydown", function (event) {
                if (event.keyCode == 9) {
                    document.getElementById(configureBrowserLinkId).focus();
                    event.preventDefault();
                    event.stopPropagation();
                }
                else if (event.keyCode == 13) {
                    okButtonKeyDownClick = true;
                    event.preventDefault();
                    event.stopPropagation();
                }
            }, false);
            okButton.addEventListener("keyup", function (event) {
                if (event.keyCode == 13 && okButtonKeyDownClick) {
                    okButtonKeyDownClick = false;
                    dismissDialogNotification();
                    event.preventDefault();
                    event.stopPropagation();
                }
            }, false);
            document.getElementById(configureBrowserLinkId).addEventListener("keydown", function (event) {
                if (event.keyCode == 9) {
                    okButton.focus();
                    event.preventDefault();
                    event.stopPropagation();
                }
            }, false);
            window.focus();
            okButton.focus();
        }
        WacDialogAction.showCrossZoneNotification = showCrossZoneNotification;
        function validateDialogDomain(dialogUrl, taskpaneUrl, allowSubdomains) {
            if (allowSubdomains === void 0) { allowSubdomains = true; }
            if (!dialogUrl || !taskpaneUrl) {
                return false;
            }
            var httpsIdentifyString = "https:";
            var parsedDialogUrl = OSF.OUtil.parseUrl(dialogUrl);
            var parsedTaskpaneUrl = OSF.OUtil.parseUrl(taskpaneUrl);
            var appDomains = OSF._OfficeAppFactory.getOfficeAppContext().get_appDomains();
            var isHttps = parsedDialogUrl.protocol === httpsIdentifyString;
            if (!isHttps) {
                return false;
            }
            var isSameDomain = parsedDialogUrl.protocol === parsedTaskpaneUrl.protocol
                && parsedDialogUrl.hostname === parsedTaskpaneUrl.hostname
                && parsedDialogUrl.port === parsedTaskpaneUrl.port;
            var isInAppDomains = OSF.XdmCommunicationManager.checkUrlWithAppDomains(appDomains, dialogUrl);
            var isTrustedDomain = isSameDomain || isInAppDomains;
            if (!isTrustedDomain && allowSubdomains) {
                isTrustedDomain = OSF.XdmCommunicationManager.isTargetSubdomainOfSourceLocation(taskpaneUrl, dialogUrl);
            }
            return isTrustedDomain;
        }
        function receiveMessage(event) {
            if (event.source == windowInstance) {
                try {
                    var dialogOrigin = event.origin;
                    var taskpaneUrl = OSF._OfficeAppFactory.getOfficeAppContext().get_docUrl();
                    var isTrustedDomain = validateDialogDomain(dialogOrigin, taskpaneUrl, true);
                    if (!isTrustedDomain) {
                        throw new Error("Received a message from a dialog with an untrusted domain.");
                    }
                    var dialogMessageReceivedArgs = {};
                    dialogMessageReceivedArgs[OSF.Marshaling.DialogMessageReceivedEventKeys.MessageType] = 0;
                    dialogMessageReceivedArgs[OSF.Marshaling.DialogMessageReceivedEventKeys.MessageContent] = event.data;
                    handler(dialogMessageReceivedArgs);
                }
                catch (e) {
                    if (OSF.AppTelemetry) {
                        OSF.AppTelemetry.logAppException("Error happened during receive message handler." + e);
                    }
                }
            }
        }
        function storageChangedHandler(event) {
            var messageKey = OSF._OfficeAppFactory.getId() + messageParentKey;
            if (event.key == messageKey) {
                try {
                    var dialogMessageReceivedArgs = {};
                    dialogMessageReceivedArgs[OSF.Marshaling.DialogMessageReceivedEventKeys.MessageType] = 0;
                    dialogMessageReceivedArgs[OSF.Marshaling.DialogMessageReceivedEventKeys.MessageContent] = event.newValue;
                    handler(dialogMessageReceivedArgs);
                }
                catch (e) {
                    if (OSF.AppTelemetry) {
                        OSF.AppTelemetry.logAppException("Error happened during storage changed handler." + e);
                    }
                }
            }
        }
        function checkAppDomain(dialogInfo) {
            var appDomains = OSF._OfficeAppFactory.getOfficeAppContext().get_appDomains();
            var url = dialogInfo["url"];
            var fInDomain = OSF.XdmCommunicationManager.checkUrlWithAppDomains(appDomains, url);
            if (!fInDomain) {
                return OSF._OfficeAppFactory.getOfficeAppContext().get_docUrl()
                    && OSF.XdmCommunicationManager.isTargetSubdomainOfSourceLocation(OSF._OfficeAppFactory.getOfficeAppContext().get_docUrl(), url);
            }
            return fInDomain;
        }
        function showDialog(dialogInfo) {
            var hostInfoObj = OSF._OfficeAppFactory.getHostInfo();
            var hostInfoVals = [
                hostInfoObj.hostType,
                hostInfoObj.hostPlatform,
                hostInfoObj.hostSpecificFileVersion,
                hostInfoObj.hostLocale,
                hostInfoObj.osfControlAppCorrelationId,
                "isDialog",
                hostInfoObj.disableLogging ? "disableLogging" : ""
            ];
            var URL_DELIM = "$";
            var hostInfo = hostInfoVals.join(URL_DELIM);
            var appContext = OSF._OfficeAppFactory.getOfficeAppContext();
            var windowUrl = dialogInfo["url"];
            if (!dialogInfo[OSF.ShowWindowDialogParameterKeys.UrlNoHostInfo]) {
                windowUrl = OSF.OUtil.addHostInfoAsQueryParam(windowUrl, hostInfo);
            }
            var windowName = JSON.parse(window.name);
            windowName["hostInfo"] = hostInfo;
            windowName["appContext"] = appContext;
            var width = dialogInfo[OSF.ShowWindowDialogParameterKeys.Width] * screen.width / 100;
            var height = dialogInfo[OSF.ShowWindowDialogParameterKeys.Height] * screen.height / 100;
            var left = appContext.get_clientWindowWidth() / 2 - width / 2;
            var top = appContext.get_clientWindowHeight() / 2 - height / 2;
            var windowSpecs = "width=" + width + ", height=" + height + ", left=" + left + ", top=" + top + ",channelmode=no,directories=no,fullscreen=no,location=no,menubar=no,resizable=yes,scrollbars=yes,status=no,titlebar=yes,toolbar=no";
            windowInstance = window.open(windowUrl, OSF.OUtil.serializeObjectToString(windowName), windowSpecs);
            if (windowInstance == null) {
                OSF.AppTelemetry.logAppCommonMessage("Encountered cross zone issue in displayDialogAsync api.");
                removeDialogNotificationElement();
                showCrossZoneNotification(windowUrl, hostInfoObj.hostType);
                showDialogCallback(12011);
                return;
            }
            if (OSF.OUtil.shouldUseLocalStorageToPassMessage()) {
                window.addEventListener("storage", storageChangedHandler);
            }
            else {
                window.addEventListener("message", receiveMessage);
            }
            function checkWindowClose() {
                try {
                    if (windowInstance == null || windowInstance.closed) {
                        window.clearInterval(checkWindowDialogCloseInterval);
                        if (OSF.OUtil.shouldUseLocalStorageToPassMessage()) {
                            window.removeEventListener("storage", storageChangedHandler);
                        }
                        else {
                            window.removeEventListener("message", receiveMessage);
                        }
                        var dialogClosedArgs = {};
                        dialogClosedArgs[OSF.Marshaling.DialogMessageReceivedEventKeys.MessageType] = 12006;
                        handler(dialogClosedArgs);
                    }
                }
                catch (e) {
                    if (OSF.AppTelemetry) {
                        OSF.AppTelemetry.logAppException("Error happened during check or handle window close." + e);
                    }
                }
            }
            checkWindowDialogCloseInterval = window.setInterval(checkWindowClose, 1000);
            if (showDialogCallback != null) {
                showDialogCallback(0);
            }
            else {
                if (OSF.AppTelemetry) {
                    OSF.AppTelemetry.logAppException("showDialogCallback can not be null.");
                }
            }
        }
        function createButtonControl(dialogCssManager, buttonValue) {
            var buttonControl = document.createElement("input");
            buttonControl.setAttribute("type", "button");
            buttonControl.style.cssText = dialogCssManager.getDialogButtonCSS();
            buttonControl.style.borderColor = commonButtonBorderColor;
            buttonControl.style.backgroundColor = commonButtonBackgroundColor;
            buttonControl.setAttribute("value", buttonValue);
            var buttonControlEventInHandler = function () {
                handleButtonControlEventIn(buttonControl);
            };
            var buttonControlEventOutHandler = function () {
                handleButtonControlEventOut(buttonControl);
            };
            buttonControl.addEventListener("mouseover", buttonControlEventInHandler);
            buttonControl.addEventListener("focus", buttonControlEventInHandler);
            buttonControl.addEventListener("mouseout", buttonControlEventOutHandler);
            buttonControl.addEventListener("focusout", buttonControlEventOutHandler);
            return buttonControl;
        }
        function handleButtonControlEventIn(buttonControl) {
            if (hostThemeButtonStyle != null) {
                buttonControl.style.borderColor = hostThemeButtonStyle.HostButtonBorderColor;
                buttonControl.style.backgroundColor = hostThemeButtonStyle.HostButtonBackgroundColor;
            }
            else if (OSF.OUtil.getCommonUI()) {
                buttonControl.style.borderColor = OSF.OUtil.getCommonUI().HostButtonBorderColor;
                buttonControl.style.backgroundColor = OSF.OUtil.getCommonUI().HostButtonBackgroundColor;
            }
            else {
                buttonControl.style.backgroundColor = commonEventInButtonBackgroundColor;
            }
        }
        function handleButtonControlEventOut(buttonControl) {
            buttonControl.style.borderColor = commonButtonBorderColor;
            buttonControl.style.backgroundColor = commonButtonBackgroundColor;
        }
        function dismissDialogNotification() {
            function unregisterDialogNotificationShownEventCallback(status) {
                removeDialogNotificationElement();
                setFocusOnFirstElement(status);
            }
            registerDialogNotificationShownArgs.onCalling = unregisterDialogNotificationShownEventCallback;
            OSF.WACDelegate.unregisterEventAsync(registerDialogNotificationShownArgs);
        }
        function removeDialogNotificationElement() {
            if (dialogNotificationPanel != null) {
                document.body.removeChild(dialogNotificationPanel);
                dialogNotificationPanel = null;
            }
            if (overlayElement != null) {
                document.body.removeChild(overlayElement);
                overlayElement = null;
            }
        }
        function createOverlayElement(dialogCssManager) {
            var overlayElement = document.createElement("div");
            overlayElement.style.cssText = dialogCssManager.getOverlayElementCSS();
            return overlayElement;
        }
        function createNotificationPanel(dialogCssManager, notificationString) {
            var dialogNotificationPanel = document.createElement("div");
            dialogNotificationPanel.style.cssText = dialogCssManager.getDialogNotificationPanelCSS();
            setAttributeForDialogNotificationPanel(dialogNotificationPanel);
            var dialogNotificationTextPanel = document.createElement("div");
            dialogNotificationTextPanel.style.cssText = dialogCssManager.getNewWindowNotificationTextPanelCSS();
            dialogNotificationTextPanel.id = dialogNotificationTextPanelId;
            if (document.documentElement.getAttribute("dir") == "rtl") {
                dialogNotificationTextPanel.style.paddingRight = "30px";
            }
            else {
                dialogNotificationTextPanel.style.paddingLeft = "30px";
            }
            var dialogNotificationTextSpan = document.createElement("span");
            dialogNotificationTextSpan.style.cssText = dialogCssManager.getNewWindowNotificationTextSpanCSS();
            dialogNotificationTextSpan.innerText = notificationString;
            dialogNotificationTextPanel.appendChild(dialogNotificationTextSpan);
            dialogNotificationPanel.appendChild(dialogNotificationTextPanel);
            return dialogNotificationPanel;
        }
        function createButtonPanel(dialogCssManager) {
            var dialogNotificationButtonPanel = document.createElement("div");
            dialogNotificationButtonPanel.style.cssText = dialogCssManager.getDialogNotificationButtonPanelCSS();
            if (document.documentElement.getAttribute("dir") == "rtl") {
                dialogNotificationButtonPanel.style.cssFloat = "left";
            }
            else {
                dialogNotificationButtonPanel.style.cssFloat = "right";
            }
            return dialogNotificationButtonPanel;
        }
        function setFocusOnFirstElement(status) {
            if (status != 0) {
                var list = document.querySelectorAll(OSF._OfficeAppFactory.getInitializationHelper().getTabbableElements());
                OSF.OUtil.focusToFirstTabbable(list, false);
            }
        }
        function createNotificationPanelForCrossZoneIssue(dialogCssManager, windowUrl) {
            var dialogNotificationPanel = document.createElement("div");
            dialogNotificationPanel.style.cssText = dialogCssManager.getDialogNotificationPanelCSS();
            setAttributeForDialogNotificationPanel(dialogNotificationPanel);
            var dialogNotificationTextPanel = document.createElement("div");
            dialogNotificationTextPanel.style.cssText = dialogCssManager.getCrossZoneNotificationTextPanelCSS();
            dialogNotificationTextPanel.id = dialogNotificationTextPanelId;
            var configureBrowserLink = document.createElement("a");
            configureBrowserLink.id = configureBrowserLinkId;
            configureBrowserLink.href = "#";
            configureBrowserLink.innerText = Strings.OfficeOM.L_NewWindowCrossZoneConfigureBrowserLink;
            configureBrowserLink.setAttribute("onclick", "window.open('https://support.microsoft.com/en-us/help/17479/windows-internet-explorer-11-change-security-privacy-settings', '_blank', 'fullscreen=1')");
            var dialogNotificationTextSpan = document.createElement("span");
            if (Strings.OfficeOM.L_NewWindowCrossZone) {
                dialogNotificationTextSpan.innerHTML = OSF.OUtil.formatString(Strings.OfficeOM.L_NewWindowCrossZone, configureBrowserLink.outerHTML, OSF.OUtil.getDomainForUrl(windowUrl));
            }
            dialogNotificationTextPanel.appendChild(dialogNotificationTextSpan);
            dialogNotificationPanel.appendChild(dialogNotificationTextPanel);
            return dialogNotificationPanel;
        }
        function setAttributeForDialogNotificationPanel(dialogNotificationDiv) {
            dialogNotificationDiv.setAttribute("role", "dialog");
            dialogNotificationDiv.setAttribute("aria-describedby", dialogNotificationTextPanelId);
        }
        function addOrRemoveEventListenersForWindow(isAdd) {
            var me = this;
            var onWindowFocus = function () {
                if (!OSF._OfficeAppFactory.getWebAppState().focused) {
                    OSF._OfficeAppFactory.getWebAppState().focused = true;
                }
                OSF.getClientEndPoint().invoke("ContextActivationManager_notifyHost", null, [me._webAppState.id, OSF.AgaveHostAction.Select]);
            };
            var onWindowBlur = function () {
                if (!OSF) {
                    return;
                }
                if (OSF._OfficeAppFactory.getWebAppState().focused) {
                    OSF._OfficeAppFactory.getWebAppState().focused = false;
                }
                OSF.getClientEndPoint().invoke("ContextActivationManager_notifyHost", null, [me._webAppState.id, OSF.AgaveHostAction.UnSelect]);
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
                    OSF.getClientEndPoint().invoke("ContextActivationManager_notifyHost", null, [me._webAppState.id, actionId]);
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
                                OSF.getClientEndPoint().invoke("ContextActivationManager_notifyHost", null, [me._webAppState.id, OSF.AgaveHostAction.TabExitShift]);
                            }
                            else {
                                OSF.getClientEndPoint().invoke("ContextActivationManager_notifyHost", null, [me._webAppState.id, OSF.AgaveHostAction.TabExit]);
                            }
                        }
                    }
                }
                else if (e.keyCode == 27) {
                    e.preventDefault();
                    escDismissDialogNotification();
                    OSF.getClientEndPoint().invoke("ContextActivationManager_notifyHost", null, [me._webAppState.id, OSF.AgaveHostAction.EscExit]);
                }
                else if (e.keyCode == 113) {
                    e.preventDefault();
                    OSF.getClientEndPoint().invoke("ContextActivationManager_notifyHost", null, [me._webAppState.id, OSF.AgaveHostAction.F2Exit]);
                }
                else if ((e.ctrlKey || e.metaKey || e.shiftKey || e.altKey) && e.keyCode >= 1 && e.keyCode <= 255) {
                    var params = {
                        "keyCode": e.keyCode,
                        "shiftKey": e.shiftKey,
                        "altKey": e.altKey,
                        "ctrlKey": e.ctrlKey,
                        "metaKey": e.metaKey
                    };
                    OSF.getClientEndPoint().invoke("ContextActivationManager_notifyHost", null, [me._webAppState.id, OSF.AgaveHostAction.KeyboardShortcuts, params]);
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
    })(WacDialogAction = OSF.WacDialogAction || (OSF.WacDialogAction = {}));
})(OSF || (OSF = {}));
var OSF;
(function (OSF) {
    var SafeArrayDelegate;
    (function (SafeArrayDelegate) {
        function openDialog(args) {
            try {
                if (args.onCalling) {
                    args.onCalling();
                }
                var callback = OSF.SafeArrayDelegate.getOnAfterRegisterEvent(true, args);
                OSF._OfficeAppFactory.getClientHostController().openDialog(args.dispId, undefined, args.targetId, function OSF_DDA_SafeArrayDelegate$RegisterEventAsync_OnEvent(eventDispId, payload) {
                    if (args.onEvent) {
                        args.onEvent(payload);
                    }
                    if (OSF.AppTelemetry) {
                        OSF.AppTelemetry.onEventDone(args.dispId);
                    }
                }, callback);
            }
            catch (ex) {
                OSF.SafeArrayDelegate.onException(ex, args);
            }
        }
        SafeArrayDelegate.openDialog = openDialog;
        function closeDialog(args) {
            if (args.onCalling) {
                args.onCalling();
            }
            var callback = OSF.SafeArrayDelegate.getOnAfterRegisterEvent(false, args);
            try {
                OSF._OfficeAppFactory.getClientHostController().closeDialog(args.dispId, undefined, args.targetId, callback);
            }
            catch (ex) {
                OSF.SafeArrayDelegate.onException(ex, args);
            }
        }
        SafeArrayDelegate.closeDialog = closeDialog;
        function messageParent(args) {
            try {
                if (args.onCalling) {
                    args.onCalling();
                }
                var startTime = (new Date()).getTime();
                var result = OSF._OfficeAppFactory.getClientHostController().messageParent(args.hostCallArgs);
                if (args.onReceiving) {
                    args.onReceiving();
                }
                if (OSF.AppTelemetry) {
                    OSF.AppTelemetry.onMethodDone(args.dispId, args.hostCallArgs, Math.abs((new Date()).getTime() - startTime), result);
                }
                return result;
            }
            catch (ex) {
                return OSF.SafeArrayDelegate.onExceptionSyncMethod(ex);
            }
        }
        SafeArrayDelegate.messageParent = messageParent;
        function sendMessage(args) {
            try {
                if (args.onCalling) {
                    args.onCalling();
                }
                var startTime = (new Date()).getTime();
                var result = OSF._OfficeAppFactory.getClientHostController().sendMessage(args.hostCallArgs);
                if (args.onReceiving) {
                    args.onReceiving();
                }
                return result;
            }
            catch (ex) {
                return OSF.SafeArrayDelegate.onExceptionSyncMethod(ex);
            }
        }
        SafeArrayDelegate.sendMessage = sendMessage;
    })(SafeArrayDelegate = OSF.SafeArrayDelegate || (OSF.SafeArrayDelegate = {}));
})(OSF || (OSF = {}));
var OSF;
(function (OSF) {
    var WACDelegate;
    (function (WACDelegate) {
        function openDialog(args) {
            var httpsIdentifyString = "https://";
            var httpIdentifyString = "http://";
            var dialogInfo = JSON.parse(args.targetId);
            var callback = WACDelegate.getOnAfterRegisterEvent(true, args);
            function showDialogCallback(status) {
                var payload = { "Error": status };
                try {
                    callback(0, payload);
                }
                catch (e) {
                    if (OSF.AppTelemetry) {
                        OSF.AppTelemetry.logAppException("Exception happens at showDialogCallback." + e);
                    }
                }
            }
            if (OSF.DialogShownStatus.hasDialogShown) {
                showDialogCallback(12007);
                return;
            }
            var dialogUrl = dialogInfo[OSF.ShowWindowDialogParameterKeys.Url].toLowerCase();
            var taskpaneUrl = (window.location.href).toLowerCase();
            if (OSF.AppTelemetry) {
                var isSameDomain = false;
                var parentIsSubdomain = false;
                var childIsSubdomain = false;
                var isAppDomain = false;
                var dialogUrlPortionAllowedToLog = "";
                var taskpaneUrlPortionAllowedToLog = "";
                if (OSF.OUtil) {
                    var parsedDialogUrl = OSF.OUtil.parseUrl(dialogUrl);
                    var parsedTaskpaneUrl = OSF.OUtil.parseUrl(taskpaneUrl);
                    isSameDomain = parsedDialogUrl.protocol === parsedTaskpaneUrl.protocol
                        && parsedDialogUrl.hostname === parsedTaskpaneUrl.hostname
                        && parsedDialogUrl.port === parsedTaskpaneUrl.port;
                    dialogUrlPortionAllowedToLog = OSF.OUtil.getHostnamePortionForLogging(parsedDialogUrl.hostname);
                    if (isSameDomain) {
                        taskpaneUrlPortionAllowedToLog = dialogUrlPortionAllowedToLog;
                    }
                    else {
                        taskpaneUrlPortionAllowedToLog = OSF.OUtil.getHostnamePortionForLogging(parsedTaskpaneUrl.hostname);
                        parentIsSubdomain = OSF.XdmCommunicationManager.isTargetSubdomainOfSourceLocation(dialogUrl, taskpaneUrl);
                        childIsSubdomain = OSF.XdmCommunicationManager.isTargetSubdomainOfSourceLocation(taskpaneUrl, dialogUrl);
                    }
                    var appDomains = OSF._OfficeAppFactory.getOfficeAppContext().get_appDomains();
                    isAppDomain = OSF.XdmCommunicationManager.checkUrlWithAppDomains(appDomains, dialogUrl);
                }
                var logJsonAsString = "openDialog isInline: " + dialogInfo[OSF.ShowWindowDialogParameterKeys.DisplayInIframe].toString() + ", " +
                    "taskpaneHostname: " + taskpaneUrlPortionAllowedToLog + ", " +
                    "dialogHostName: " + dialogUrlPortionAllowedToLog + ", " +
                    "isSameDomain: " + isSameDomain.toString() + ", " +
                    "parentIsSubdomain: " + parentIsSubdomain.toString() + ", " +
                    "childIsSubdomain: " + childIsSubdomain.toString() + ", " +
                    "isAppDomain: " + isAppDomain.toString();
                OSF.AppTelemetry.logAppCommonMessage(logJsonAsString);
            }
            if (dialogUrl == null || !(dialogUrl.substr(0, httpsIdentifyString.length) === httpsIdentifyString)) {
                if (dialogUrl.substr(0, httpIdentifyString.length) === httpIdentifyString) {
                    showDialogCallback(12005);
                }
                else {
                    showDialogCallback(12003);
                }
                return;
            }
            if (!dialogInfo[OSF.ShowWindowDialogParameterKeys.DisplayInIframe]) {
                OSF.DialogShownStatus.isWindowDialog = true;
                OSF.WacDialogAction.setHandlerAndShowDialogCallback(function OSF_DDA_WACDelegate$RegisterEventAsync_OnEvent(payload) {
                    if (args.onEvent) {
                        args.onEvent(payload);
                    }
                    if (OSF.AppTelemetry) {
                        OSF.AppTelemetry.onEventDone(args.dispId);
                    }
                }, showDialogCallback);
                OSF.WacDialogAction.handleNewWindowDialog(dialogInfo);
            }
            else {
                OSF.DialogShownStatus.isWindowDialog = false;
                OSF.WACDelegate.registerEventAsync(args);
            }
        }
        WACDelegate.openDialog = openDialog;
        function messageParent(args) {
            if (window.opener != null) {
                OSF.WacDialogAction.messageParent(args);
            }
            else {
                OSF.WACDelegate.executeAsync(args);
            }
        }
        WACDelegate.messageParent = messageParent;
        function sendMessage(args) {
            if (OSF.DialogShownStatus.hasDialogShown) {
                if (OSF.DialogShownStatus.isWindowDialog) {
                    OSF.WacDialogAction.sendMessage(args);
                }
                else {
                    OSF.WACDelegate.executeAsync(args);
                }
            }
        }
        WACDelegate.sendMessage = sendMessage;
        function closeDialog(args) {
            var callback = WACDelegate.getOnAfterRegisterEvent(false, args);
            function closeDialogCallback(status) {
                var payload = { "Error": status };
                try {
                    callback(0, payload);
                }
                catch (e) {
                    if (OSF.AppTelemetry) {
                        OSF.AppTelemetry.logAppException("Exception happens at closeDialogCallback." + e);
                    }
                }
            }
            if (!OSF.DialogShownStatus.hasDialogShown) {
                closeDialogCallback(12006);
            }
            else {
                if (OSF.DialogShownStatus.isWindowDialog) {
                    if (args.onCalling) {
                        args.onCalling();
                    }
                    OSF.WacDialogAction.closeDialog(closeDialogCallback);
                }
                else {
                    OSF.WACDelegate.unregisterEventAsync(args);
                }
            }
        }
        WACDelegate.closeDialog = closeDialog;
    })(WACDelegate = OSF.WACDelegate || (OSF.WACDelegate = {}));
})(OSF || (OSF = {}));
var OSF;
(function (OSF) {
    var Event;
    (function (Event) {
        function defineMethods() {
            OSF.AsyncMethodCalls.define({
                method: OSF.AsyncMethods.AddHandlerAsync,
                requiredArguments: [{
                        "name": OSF.ParameterNames.EventType,
                        "enum": OSF.EventType,
                        "verify": function (eventType, caller, eventDispatch) { return eventDispatch.supportsEvent(eventType); }
                    },
                    {
                        "name": OSF.ParameterNames.Handler,
                        "types": ["function"]
                    }
                ],
                supportedOptions: [],
                privateStateCallbacks: []
            });
            OSF.AsyncMethodCalls.define({
                method: OSF.AsyncMethods.RemoveHandlerAsync,
                requiredArguments: [
                    {
                        "name": OSF.ParameterNames.EventType,
                        "enum": OSF.EventType,
                        "verify": function (eventType, caller, eventDispatch) { return eventDispatch.supportsEvent(eventType); }
                    }
                ],
                supportedOptions: [
                    {
                        name: OSF.ParameterNames.Handler,
                        value: {
                            "types": ["function", "object"],
                            "defaultValue": null
                        }
                    }
                ],
                privateStateCallbacks: []
            });
        }
        OSF.V10ApiFeatureRegistry.register({
            defineMethodsFunc: defineMethods,
        });
    })(Event || (Event = {}));
})(OSF || (OSF = {}));
var OSF;
(function (OSF) {
    var OpenBrowserWindow;
    (function (OpenBrowserWindow) {
        function defineMethods() {
            OSF.AsyncMethodCalls.define({
                method: OSF.AsyncMethods.OpenBrowserWindow,
                requiredArguments: [
                    {
                        "name": OSF.ParameterNames.Url,
                        "types": ["string"]
                    }
                ],
                supportedOptions: [
                    {
                        name: OSF.ParameterNames.Reserved,
                        value: {
                            "types": ["number"],
                            "defaultValue": 0
                        }
                    }
                ],
                privateStateCallbacks: []
            });
        }
        function defineSafeArrayParameterMap() {
            OSF.HostParameterMap.define({
                type: 102,
                toHost: [
                    { name: OSF.ParameterNames.Reserved, value: 0 },
                    { name: OSF.ParameterNames.Url, value: 1 }
                ]
            });
        }
        function initialize() {
            if (OSF.OUtil.getHostPlatform() != OSF.HostInfoPlatform.web) {
                var target = Office.context.ui;
                OSF.DispIdHost.addAsyncMethods(target, [OSF.AsyncMethods.OpenBrowserWindow]);
            }
        }
        OSF.V10ApiFeatureRegistry.register({
            defineMethodsFunc: defineMethods,
            defineSafeArrayParameterMapFunc: defineSafeArrayParameterMap,
            initializeFunc: initialize
        });
    })(OpenBrowserWindow || (OpenBrowserWindow = {}));
})(OSF || (OSF = {}));
var OSF;
(function (OSF) {
    var AccountTypeFilter;
    (function (AccountTypeFilter) {
        AccountTypeFilter["NoFilter"] = "noFilter";
        AccountTypeFilter["AAD"] = "aad";
        AccountTypeFilter["MSA"] = "msa";
    })(AccountTypeFilter = OSF.AccountTypeFilter || (OSF.AccountTypeFilter = {}));
    ;
})(OSF || (OSF = {}));
var Office;
(function (Office) {
    var context;
    (function (context) {
        var auth;
        (function (auth) {
        })(auth = context.auth || (context.auth = {}));
    })(context = Office.context || (Office.context = {}));
})(Office || (Office = {}));
(function (OSF) {
    var Marshaling;
    (function (Marshaling) {
        var GetAccessTokenKeys;
        (function (GetAccessTokenKeys) {
            GetAccessTokenKeys["ForceConsent"] = "forceConsent";
            GetAccessTokenKeys["ForceAddAccount"] = "forceAddAccount";
            GetAccessTokenKeys["AuthChallenge"] = "authChallenge";
            GetAccessTokenKeys["AllowConsentPrompt"] = "allowConsentPrompt";
            GetAccessTokenKeys["ForMSGraphAccess"] = "forMSGraphAccess";
            GetAccessTokenKeys["AllowSignInPrompt"] = "allowSignInPrompt";
            GetAccessTokenKeys["EnableNewHosts"] = "enableNewHosts";
            GetAccessTokenKeys["AccountTypeFilter"] = "accountTypeFilter";
            GetAccessTokenKeys["AddinTrustId"] = "addinTrustId";
        })(GetAccessTokenKeys = Marshaling.GetAccessTokenKeys || (Marshaling.GetAccessTokenKeys = {}));
        ;
        var AccessTokenResultKeys;
        (function (AccessTokenResultKeys) {
            AccessTokenResultKeys["AccessToken"] = "accessToken";
        })(AccessTokenResultKeys = Marshaling.AccessTokenResultKeys || (Marshaling.AccessTokenResultKeys = {}));
        ;
    })(Marshaling = OSF.Marshaling || (OSF.Marshaling = {}));
})(OSF || (OSF = {}));
var OfficeExt;
(function (OfficeExt) {
    var SingleSignOn = (function () {
        function SingleSignOn(parameters) {
        }
        return SingleSignOn;
    }());
    OfficeExt.SingleSignOn = SingleSignOn;
})(OfficeExt || (OfficeExt = {}));
(function (OSF) {
    var SingleSignOn;
    (function (SingleSignOn) {
        function defineMethods() {
            OSF.AsyncMethodCalls.define({
                method: OSF.AsyncMethods.GetAccessTokenAsync,
                requiredArguments: [],
                supportedOptions: [
                    {
                        name: OSF.ParameterNames.ForceConsent,
                        value: {
                            "types": ["boolean"],
                            "defaultValue": false
                        }
                    },
                    {
                        name: OSF.ParameterNames.ForceAddAccount,
                        value: {
                            "types": ["boolean"],
                            "defaultValue": false
                        }
                    },
                    {
                        name: OSF.ParameterNames.AuthChallenge,
                        value: {
                            "types": ["string"],
                            "defaultValue": ""
                        }
                    },
                    {
                        name: OSF.ParameterNames.AllowConsentPrompt,
                        value: {
                            "types": ["boolean"],
                            "defaultValue": false
                        }
                    },
                    {
                        name: OSF.ParameterNames.ForMSGraphAccess,
                        value: {
                            "types": ["boolean"],
                            "defaultValue": false
                        }
                    },
                    {
                        name: OSF.ParameterNames.AllowSignInPrompt,
                        value: {
                            "types": ["boolean"],
                            "defaultValue": false
                        }
                    },
                    {
                        name: OSF.ParameterNames.EnableNewHosts,
                        value: {
                            "types": ["number"],
                            "defaultValue": 0
                        }
                    },
                    {
                        name: OSF.ParameterNames.AccountTypeFilter,
                        value: {
                            "enum": OSF.AccountTypeFilter,
                            "defaultValue": OSF.AccountTypeFilter.NoFilter
                        }
                    }
                ],
                checkCallArgs: function (callArgs, caller, stateInfo) {
                    var _a;
                    var appContext = OSF._OfficeAppFactory.getOfficeAppContext();
                    if (appContext && appContext.get_wopiHostOriginForSingleSignOn()) {
                        var addinTrustId = OSF.OUtil.Guid.generateNewGuid();
                        window.parent.parent.postMessage("{\"MessageId\":\"AddinTrustedOrigin\",\"AddinTrustId\":\"" + addinTrustId + "\"}", appContext.get_wopiHostOriginForSingleSignOn());
                        callArgs[OSF.ParameterNames.AddinTrustId] = addinTrustId;
                    }
                    if (window.Office.context.requirements.isSetSupported("JsonPayloadSSO")) {
                        var jsonParameterMap = (_a = {},
                            _a[OSF.ParameterNames.ForceConsent] = false,
                            _a[OSF.ParameterNames.ForceAddAccount] = false,
                            _a[OSF.ParameterNames.AuthChallenge] = true,
                            _a[OSF.ParameterNames.AllowConsentPrompt] = true,
                            _a[OSF.ParameterNames.ForMSGraphAccess] = true,
                            _a[OSF.ParameterNames.AllowSignInPrompt] = true,
                            _a[OSF.ParameterNames.EnableNewHosts] = true,
                            _a[OSF.ParameterNames.AccountTypeFilter] = true,
                            _a);
                        var jsonPayload = {};
                        for (var _i = 0, _b = Object.keys(jsonParameterMap); _i < _b.length; _i++) {
                            var key = _b[_i];
                            if (jsonParameterMap[key]) {
                                jsonPayload[key] = callArgs[key];
                            }
                            delete callArgs[key];
                        }
                        callArgs[OSF.ParameterNames.JsonPayload] = JSON.stringify(jsonPayload);
                    }
                    return callArgs;
                },
                onSucceeded: function (dataDescriptor, caller, callArgs) {
                    var data = dataDescriptor[OSF.ParameterNames.Data];
                    return data;
                }
            });
        }
        function defineSafeArrayParameterMap() {
            OSF.HostParameterMap.define({
                type: 98,
                toHost: [
                    { name: OSF.ParameterNames.JsonPayload, value: 0 },
                    { name: OSF.ParameterNames.ForceConsent, value: 0 },
                    { name: OSF.ParameterNames.ForceAddAccount, value: 1 },
                    { name: OSF.ParameterNames.AuthChallenge, value: 2 },
                    { name: OSF.ParameterNames.AllowConsentPrompt, value: 3 },
                    { name: OSF.ParameterNames.ForMSGraphAccess, value: 4 },
                    { name: OSF.ParameterNames.AllowSignInPrompt, value: 5 }
                ],
                fromHost: [
                    { name: OSF.ParameterNames.Data, value: OSF.HostParameterMap.self }
                ]
            });
        }
        function defineWebParameterMap() {
            OSF.HostParameterMap.define({
                type: 98,
                toHost: [
                    { name: OSF.ParameterNames.ForceConsent, value: OSF.Marshaling.GetAccessTokenKeys.ForceConsent },
                    { name: OSF.ParameterNames.ForceAddAccount, value: OSF.Marshaling.GetAccessTokenKeys.ForceAddAccount },
                    { name: OSF.ParameterNames.AuthChallenge, value: OSF.Marshaling.GetAccessTokenKeys.AuthChallenge },
                    { name: OSF.ParameterNames.AllowConsentPrompt, value: OSF.Marshaling.GetAccessTokenKeys.AllowConsentPrompt },
                    { name: OSF.ParameterNames.ForMSGraphAccess, value: OSF.Marshaling.GetAccessTokenKeys.ForMSGraphAccess },
                    { name: OSF.ParameterNames.AllowSignInPrompt, value: OSF.Marshaling.GetAccessTokenKeys.AllowSignInPrompt },
                    { name: OSF.ParameterNames.EnableNewHosts, value: OSF.Marshaling.GetAccessTokenKeys.EnableNewHosts },
                    { name: OSF.ParameterNames.AccountTypeFilter, value: OSF.Marshaling.GetAccessTokenKeys.AccountTypeFilter },
                    { name: OSF.ParameterNames.AddinTrustId, value: OSF.Marshaling.GetAccessTokenKeys.AddinTrustId }
                ],
                fromHost: [
                    { name: OSF.ParameterNames.Data, value: OSF.Marshaling.AccessTokenResultKeys.AccessToken }
                ]
            });
        }
        function initialize() {
            var target = Office.context.auth;
            OSF.DispIdHost.addAsyncMethods(target, [OSF.AsyncMethods.GetAccessTokenAsync]);
        }
        OSF.V10ApiFeatureRegistry.register({
            defineMethodsFunc: defineMethods,
            defineSafeArrayParameterMapFunc: defineSafeArrayParameterMap,
            defineWebParameterMapFunc: defineWebParameterMap,
            initializeFunc: initialize
        });
    })(SingleSignOn = OSF.SingleSignOn || (OSF.SingleSignOn = {}));
})(OSF || (OSF = {}));
var OSF;
(function (OSF) {
    OSF.BootStrapExtension.prepareApiSurface = function () {
        return new Promise(function () {
            OSF.V10ApiFeatureRegistry.initialize();
        });
    };
})(OSF || (OSF = {}));
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
var oteljs;
(function (oteljs) {
    function addContractField(dataFields, instanceName, contractName) {
        dataFields.push(oteljs.makeStringDataField("zC." + instanceName, contractName));
    }
    oteljs.addContractField = addContractField;
})(oteljs || (oteljs = {}));
var oteljs;
(function (oteljs) {
    var DataClassification;
    (function (DataClassification) {
        DataClassification[DataClassification["EssentialServiceMetadata"] = 1] = "EssentialServiceMetadata";
        DataClassification[DataClassification["AccountData"] = 2] = "AccountData";
        DataClassification[DataClassification["SystemMetadata"] = 4] = "SystemMetadata";
        DataClassification[DataClassification["OrganizationIdentifiableInformation"] = 8] = "OrganizationIdentifiableInformation";
        DataClassification[DataClassification["EndUserIdentifiableInformation"] = 16] = "EndUserIdentifiableInformation";
        DataClassification[DataClassification["CustomerContent"] = 32] = "CustomerContent";
        DataClassification[DataClassification["AccessControl"] = 64] = "AccessControl";
    })(DataClassification = oteljs.DataClassification || (oteljs.DataClassification = {}));
})(oteljs || (oteljs = {}));
var oteljs;
(function (oteljs) {
    function makeBooleanDataField(name, value) {
        return {
            name: name,
            dataType: oteljs.DataFieldType.Boolean,
            value: value,
            classification: oteljs.DataClassification.SystemMetadata
        };
    }
    oteljs.makeBooleanDataField = makeBooleanDataField;
    function makeInt64DataField(name, value) {
        return {
            name: name,
            dataType: oteljs.DataFieldType.Int64,
            value: value,
            classification: oteljs.DataClassification.SystemMetadata
        };
    }
    oteljs.makeInt64DataField = makeInt64DataField;
    function makeDoubleDataField(name, value) {
        return {
            name: name,
            dataType: oteljs.DataFieldType.Double,
            value: value,
            classification: oteljs.DataClassification.SystemMetadata
        };
    }
    oteljs.makeDoubleDataField = makeDoubleDataField;
    function makeStringDataField(name, value) {
        return {
            name: name,
            dataType: oteljs.DataFieldType.String,
            value: value,
            classification: oteljs.DataClassification.SystemMetadata
        };
    }
    oteljs.makeStringDataField = makeStringDataField;
    function makeGuidDataField(name, value) {
        return {
            name: name,
            dataType: oteljs.DataFieldType.Guid,
            value: value,
            classification: oteljs.DataClassification.SystemMetadata
        };
    }
    oteljs.makeGuidDataField = makeGuidDataField;
})(oteljs || (oteljs = {}));
var oteljs;
(function (oteljs) {
    var DataFieldType;
    (function (DataFieldType) {
        DataFieldType[DataFieldType["String"] = 0] = "String";
        DataFieldType[DataFieldType["Boolean"] = 1] = "Boolean";
        DataFieldType[DataFieldType["Int64"] = 2] = "Int64";
        DataFieldType[DataFieldType["Double"] = 3] = "Double";
        DataFieldType[DataFieldType["Guid"] = 4] = "Guid";
    })(DataFieldType = oteljs.DataFieldType || (oteljs.DataFieldType = {}));
})(oteljs || (oteljs = {}));
var oteljs;
(function (oteljs) {
    var Event = (function () {
        function Event() {
            this._listeners = [];
        }
        Event.prototype.fireEvent = function (args) {
            this._listeners.forEach(function (listener) { return listener(args); });
        };
        Event.prototype.addListener = function (listener) {
            if (listener) {
                this._listeners.push(listener);
            }
        };
        Event.prototype.removeListener = function (listener) {
            this._listeners = this._listeners.filter(function (h) { return h !== listener; });
        };
        Event.prototype.getListenerCount = function () {
            return this._listeners.length;
        };
        return Event;
    }());
    oteljs.Event = Event;
})(oteljs || (oteljs = {}));
var oteljs;
(function (oteljs) {
    function getEffectiveEventFlags(telemetryEvent) {
        var eventFlags = {
            costPriority: oteljs.CostPriority.Normal,
            samplingPolicy: oteljs.SamplingPolicy.Measure,
            persistencePriority: oteljs.PersistencePriority.Normal,
            dataCategories: oteljs.DataCategories.NotSet,
            diagnosticLevel: oteljs.DiagnosticLevel.FullEvent
        };
        if (!telemetryEvent.eventFlags || !telemetryEvent.eventFlags.dataCategories) {
            oteljs.logNotification(oteljs.LogLevel.Error, oteljs.Category.Core, function () { return 'Event is missing DataCategories event flag'; });
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
    oteljs.getEffectiveEventFlags = getEffectiveEventFlags;
})(oteljs || (oteljs = {}));
var oteljs;
(function (oteljs) {
    var SamplingPolicy;
    (function (SamplingPolicy) {
        SamplingPolicy[SamplingPolicy["NotSet"] = 0] = "NotSet";
        SamplingPolicy[SamplingPolicy["Measure"] = 1] = "Measure";
        SamplingPolicy[SamplingPolicy["Diagnostics"] = 2] = "Diagnostics";
        SamplingPolicy[SamplingPolicy["CriticalBusinessImpact"] = 191] = "CriticalBusinessImpact";
        SamplingPolicy[SamplingPolicy["CriticalCensus"] = 192] = "CriticalCensus";
        SamplingPolicy[SamplingPolicy["CriticalExperimentation"] = 193] = "CriticalExperimentation";
        SamplingPolicy[SamplingPolicy["CriticalUsage"] = 194] = "CriticalUsage";
    })(SamplingPolicy = oteljs.SamplingPolicy || (oteljs.SamplingPolicy = {}));
    var PersistencePriority;
    (function (PersistencePriority) {
        PersistencePriority[PersistencePriority["NotSet"] = 0] = "NotSet";
        PersistencePriority[PersistencePriority["Normal"] = 1] = "Normal";
        PersistencePriority[PersistencePriority["High"] = 2] = "High";
    })(PersistencePriority = oteljs.PersistencePriority || (oteljs.PersistencePriority = {}));
    var CostPriority;
    (function (CostPriority) {
        CostPriority[CostPriority["NotSet"] = 0] = "NotSet";
        CostPriority[CostPriority["Normal"] = 1] = "Normal";
        CostPriority[CostPriority["High"] = 2] = "High";
    })(CostPriority = oteljs.CostPriority || (oteljs.CostPriority = {}));
    var DataCategories;
    (function (DataCategories) {
        DataCategories[DataCategories["NotSet"] = 0] = "NotSet";
        DataCategories[DataCategories["SoftwareSetup"] = 1] = "SoftwareSetup";
        DataCategories[DataCategories["ProductServiceUsage"] = 2] = "ProductServiceUsage";
        DataCategories[DataCategories["ProductServicePerformance"] = 4] = "ProductServicePerformance";
        DataCategories[DataCategories["DeviceConfiguration"] = 8] = "DeviceConfiguration";
        DataCategories[DataCategories["InkingTypingSpeech"] = 16] = "InkingTypingSpeech";
    })(DataCategories = oteljs.DataCategories || (oteljs.DataCategories = {}));
    var DiagnosticLevel;
    (function (DiagnosticLevel) {
        DiagnosticLevel[DiagnosticLevel["ReservedDoNotUse"] = 0] = "ReservedDoNotUse";
        DiagnosticLevel[DiagnosticLevel["BasicEvent"] = 10] = "BasicEvent";
        DiagnosticLevel[DiagnosticLevel["FullEvent"] = 100] = "FullEvent";
        DiagnosticLevel[DiagnosticLevel["NecessaryServiceDataEvent"] = 110] = "NecessaryServiceDataEvent";
        DiagnosticLevel[DiagnosticLevel["AlwaysOnNecessaryServiceDataEvent"] = 120] = "AlwaysOnNecessaryServiceDataEvent";
    })(DiagnosticLevel = oteljs.DiagnosticLevel || (oteljs.DiagnosticLevel = {}));
})(oteljs || (oteljs = {}));
var oteljs;
(function (oteljs) {
    var onNotificationEvent = new oteljs.Event();
    var LogLevel;
    (function (LogLevel) {
        LogLevel[LogLevel["Error"] = 0] = "Error";
        LogLevel[LogLevel["Warning"] = 1] = "Warning";
        LogLevel[LogLevel["Info"] = 2] = "Info";
        LogLevel[LogLevel["Verbose"] = 3] = "Verbose";
    })(LogLevel = oteljs.LogLevel || (oteljs.LogLevel = {}));
    var Category;
    (function (Category) {
        Category[Category["Core"] = 0] = "Core";
        Category[Category["Sink"] = 1] = "Sink";
        Category[Category["Transport"] = 2] = "Transport";
    })(Category = oteljs.Category || (oteljs.Category = {}));
    function onNotification() {
        return onNotificationEvent;
    }
    oteljs.onNotification = onNotification;
    function logNotification(level, category, message) {
        onNotificationEvent.fireEvent({ level: level, category: category, message: message });
    }
    oteljs.logNotification = logNotification;
    function logError(category, message, error) {
        logNotification(LogLevel.Error, category, function () {
            var errorMessage = error instanceof Error ? error.message : '';
            return message + ": " + errorMessage;
        });
    }
    oteljs.logError = logError;
})(oteljs || (oteljs = {}));
var oteljs;
(function (oteljs) {
    oteljs.SuppressNexus = -1;
    var SimpleTelemetryLogger = (function () {
        function SimpleTelemetryLogger(parent, persistentDataFields, config) {
            var _a, _b;
            this.onSendEvent = new oteljs.Event();
            this.persistentDataFields = [];
            this.config = config || {};
            if (parent) {
                this.onSendEvent = parent.onSendEvent;
                (_a = this.persistentDataFields).push.apply(_a, parent.persistentDataFields);
                this.config = __assign(__assign({}, parent.getConfig()), this.config);
            }
            else {
                this.persistentDataFields.push(oteljs.makeStringDataField('OTelJS.Version', oteljs.oteljsVersion));
            }
            if (persistentDataFields) {
                (_b = this.persistentDataFields).push.apply(_b, persistentDataFields);
            }
        }
        SimpleTelemetryLogger.prototype.sendTelemetryEvent = function (event) {
            var localEvent;
            try {
                if (this.onSendEvent.getListenerCount() === 0) {
                    oteljs.logNotification(oteljs.LogLevel.Warning, oteljs.Category.Core, function () { return 'No telemetry sinks are attached.'; });
                    return;
                }
                localEvent = this.cloneEvent(event);
                this.processTelemetryEvent(localEvent);
            }
            catch (error) {
                oteljs.logError(oteljs.Category.Core, 'SendTelemetryEvent', error);
                return;
            }
            try {
                this.onSendEvent.fireEvent(localEvent);
            }
            catch (_e) {
            }
        };
        SimpleTelemetryLogger.prototype.processTelemetryEvent = function (event) {
            var _a;
            if (!event.telemetryProperties) {
                event.telemetryProperties = oteljs.TenantTokenManager.getTenantTokens(event.eventName);
            }
            if (event.dataFields && this.persistentDataFields) {
                (_a = event.dataFields).unshift.apply(_a, this.persistentDataFields);
            }
            if (!this.config.disableValidation) {
                oteljs.TelemetryEventValidator.validateTelemetryEvent(event);
            }
        };
        SimpleTelemetryLogger.prototype.addSink = function (sink) {
            this.onSendEvent.addListener(function (event) { return sink.sendTelemetryEvent(event); });
        };
        SimpleTelemetryLogger.prototype.setTenantToken = function (namespace, ariaTenantToken, nexusTenantToken) {
            oteljs.TenantTokenManager.setTenantToken(namespace, ariaTenantToken, nexusTenantToken);
        };
        SimpleTelemetryLogger.prototype.setTenantTokens = function (tokenTree) {
            oteljs.TenantTokenManager.setTenantTokens(tokenTree);
        };
        SimpleTelemetryLogger.prototype.cloneEvent = function (event) {
            var localEvent = { eventName: event.eventName, eventFlags: event.eventFlags };
            if (!!event.telemetryProperties) {
                localEvent.telemetryProperties = {
                    ariaTenantToken: event.telemetryProperties.ariaTenantToken,
                    nexusTenantToken: event.telemetryProperties.nexusTenantToken
                };
            }
            if (!!event.eventContract) {
                localEvent.eventContract = { name: event.eventContract.name, dataFields: event.eventContract.dataFields.slice() };
            }
            localEvent.dataFields = !!event.dataFields ? event.dataFields.slice() : [];
            return localEvent;
        };
        SimpleTelemetryLogger.prototype.getConfig = function () {
            return this.config;
        };
        return SimpleTelemetryLogger;
    }());
    oteljs.SimpleTelemetryLogger = SimpleTelemetryLogger;
})(oteljs || (oteljs = {}));
var oteljs;
(function (oteljs) {
    var TelemetryEventValidator;
    (function (TelemetryEventValidator) {
        var INT64_MIN = -9007199254740991;
        var INT64_MAX = 9007199254740991;
        var StartsWithCapitalRegex = /^[A-Z][a-zA-Z0-9]*$/;
        var AlphanumericRegex = /^[a-zA-Z0-9_\.]*$/;
        function validateTelemetryEvent(event) {
            if (!isEventNameValid(event.eventName)) {
                throw new Error('Invalid eventName');
            }
            if (event.eventContract && !isEventContractValid(event.eventContract)) {
                throw new Error('Invalid eventContract');
            }
            if (event.dataFields != null) {
                for (var i = 0; i < event.dataFields.length; i++) {
                    validateDataField(event.dataFields[i]);
                }
            }
        }
        TelemetryEventValidator.validateTelemetryEvent = validateTelemetryEvent;
        function isNamespaceValid(eventNamePieces) {
            return !!eventNamePieces && eventNamePieces.length >= 3 && eventNamePieces[0] === 'Office';
        }
        function isEventNodeValid(eventNode) {
            return eventNode !== undefined && StartsWithCapitalRegex.test(eventNode);
        }
        function isEventNameValid(eventName) {
            var maxEventNameLength = 98;
            if (!eventName || eventName.length > maxEventNameLength) {
                return false;
            }
            var eventNamePieces = eventName.split('.');
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
                throw new Error('Invalid dataField name');
            }
            if (dataField.dataType === oteljs.DataFieldType.Int64) {
                validateInt(dataField.value);
            }
        }
        function validateInt(value) {
            if (typeof value !== 'number' || !isFinite(value) || Math.floor(value) !== value || value < INT64_MIN || value > INT64_MAX) {
                throw new Error("Invalid integer " + JSON.stringify(value));
            }
        }
        TelemetryEventValidator.validateInt = validateInt;
    })(TelemetryEventValidator = oteljs.TelemetryEventValidator || (oteljs.TelemetryEventValidator = {}));
})(oteljs || (oteljs = {}));
var oteljs;
(function (oteljs) {
    var TokenType;
    (function (TokenType) {
        TokenType[TokenType["Aria"] = 0] = "Aria";
        TokenType[TokenType["Nexus"] = 1] = "Nexus";
    })(TokenType || (TokenType = {}));
    var TenantTokenManager;
    (function (TenantTokenManager) {
        var ariaTokenMap = {};
        var nexusTokenMap = {};
        var tenantTokens = {};
        function setTenantToken(namespace, ariaTenantToken, nexusTenantToken) {
            var parts = namespace.split('.');
            if (parts.length < 2 || parts[0] !== 'Office') {
                oteljs.logNotification(oteljs.LogLevel.Error, oteljs.Category.Core, function () {
                    return "Invalid namespace: " + namespace;
                });
                return;
            }
            var leaf = Object.create(Object.prototype);
            if (ariaTenantToken) {
                leaf['ariaTenantToken'] = ariaTenantToken;
            }
            if (nexusTenantToken) {
                leaf['nexusTenantToken'] = nexusTenantToken;
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
            if (typeof tokenTree !== 'object') {
                throw new Error('tokenTree must be an object');
            }
            tenantTokens = mergeTenantTokens(tenantTokens, tokenTree);
        }
        TenantTokenManager.setTenantTokens = setTenantTokens;
        function getTenantTokens(eventName) {
            var ariaTenantToken = getAriaTenantToken(eventName);
            var nexusTenantToken = getNexusTenantToken(eventName);
            if (!nexusTenantToken || !ariaTenantToken) {
                throw new Error('Could not find tenant token for ' + eventName);
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
            if (typeof ariaToken === 'string') {
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
            if (typeof nexusToken === 'number') {
                nexusTokenMap[eventName] = nexusToken;
                return nexusToken;
            }
            return undefined;
        }
        TenantTokenManager.getNexusTenantToken = getNexusTenantToken;
        function getTenantToken(eventName, tokenType) {
            var pieces = eventName.split('.');
            var node = tenantTokens;
            var token = undefined;
            if (!node) {
                return undefined;
            }
            for (var i = 0; i < pieces.length - 1; i++) {
                if (node[pieces[i]]) {
                    node = node[pieces[i]];
                    if (tokenType === TokenType.Aria && typeof node.ariaTenantToken === 'string') {
                        token = node.ariaTenantToken;
                    }
                    else if (tokenType === TokenType.Nexus && typeof node.nexusTenantToken === 'number') {
                        token = node.nexusTenantToken;
                    }
                }
            }
            return token;
        }
        function mergeTenantTokens(existingTokenTree, newTokenTree) {
            if (typeof newTokenTree !== 'object') {
                return newTokenTree;
            }
            for (var _i = 0, _a = Object.keys(newTokenTree); _i < _a.length; _i++) {
                var key = _a[_i];
                if (key in existingTokenTree && typeof (existingTokenTree[key] === 'object')) {
                    existingTokenTree[key] = mergeTenantTokens(existingTokenTree[key], newTokenTree[key]);
                }
                else {
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
    })(TenantTokenManager = oteljs.TenantTokenManager || (oteljs.TenantTokenManager = {}));
})(oteljs || (oteljs = {}));
var oteljs;
(function (oteljs) {
    oteljs.oteljsVersion = '3.1.46';
})(oteljs || (oteljs = {}));
var oteljs;
(function (oteljs) {
    var Contracts;
    (function (Contracts) {
        var Office;
        (function (Office) {
            var System;
            (function (System) {
                var SDX;
                (function (SDX) {
                    var contractName = 'Office.System.SDX';
                    function getFields(instanceName, contract) {
                        var dataFields = [];
                        if (contract.id !== undefined) {
                            dataFields.push(oteljs.makeStringDataField(instanceName + ".Id", contract.id));
                        }
                        if (contract.version !== undefined) {
                            dataFields.push(oteljs.makeStringDataField(instanceName + ".Version", contract.version));
                        }
                        if (contract.instanceId !== undefined) {
                            dataFields.push(oteljs.makeStringDataField(instanceName + ".InstanceId", contract.instanceId));
                        }
                        if (contract.name !== undefined) {
                            dataFields.push(oteljs.makeStringDataField(instanceName + ".Name", contract.name));
                        }
                        if (contract.marketplaceType !== undefined) {
                            dataFields.push(oteljs.makeStringDataField(instanceName + ".MarketplaceType", contract.marketplaceType));
                        }
                        if (contract.sessionId !== undefined) {
                            dataFields.push(oteljs.makeStringDataField(instanceName + ".SessionId", contract.sessionId));
                        }
                        if (contract.browserToken !== undefined) {
                            dataFields.push(oteljs.makeStringDataField(instanceName + ".BrowserToken", contract.browserToken));
                        }
                        if (contract.osfRuntimeVersion !== undefined) {
                            dataFields.push(oteljs.makeStringDataField(instanceName + ".OsfRuntimeVersion", contract.osfRuntimeVersion));
                        }
                        if (contract.officeJsVersion !== undefined) {
                            dataFields.push(oteljs.makeStringDataField(instanceName + ".OfficeJsVersion", contract.officeJsVersion));
                        }
                        if (contract.hostJsVersion !== undefined) {
                            dataFields.push(oteljs.makeStringDataField(instanceName + ".HostJsVersion", contract.hostJsVersion));
                        }
                        if (contract.assetId !== undefined) {
                            dataFields.push(oteljs.makeStringDataField(instanceName + ".AssetId", contract.assetId));
                        }
                        if (contract.providerName !== undefined) {
                            dataFields.push(oteljs.makeStringDataField(instanceName + ".ProviderName", contract.providerName));
                        }
                        if (contract.type !== undefined) {
                            dataFields.push(oteljs.makeStringDataField(instanceName + ".Type", contract.type));
                        }
                        oteljs.addContractField(dataFields, instanceName, contractName);
                        return dataFields;
                    }
                    SDX.getFields = getFields;
                })(SDX = System.SDX || (System.SDX = {}));
            })(System = Office.System || (Office.System = {}));
        })(Office = Contracts.Office || (Contracts.Office = {}));
    })(Contracts = oteljs.Contracts || (oteljs.Contracts = {}));
})(oteljs || (oteljs = {}));
/* Outlook.js API library */
/* office-js-api version: 20210406.1 */
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
var Outlook = typeof Outlook === "object" ? Outlook : {}; Outlook["OutlookAppOm"] =
/******/ (function(modules) { // webpackBootstrap
/******/ 	// The module cache
/******/ 	var installedModules = {};
/******/
/******/ 	// The require function
/******/ 	function __webpack_require__(moduleId) {
/******/
/******/ 		// Check if module is in cache
/******/ 		if(installedModules[moduleId]) {
/******/ 			return installedModules[moduleId].exports;
/******/ 		}
/******/ 		// Create a new module (and put it into the cache)
/******/ 		var module = installedModules[moduleId] = {
/******/ 			i: moduleId,
/******/ 			l: false,
/******/ 			exports: {}
/******/ 		};
/******/
/******/ 		// Execute the module function
/******/ 		modules[moduleId].call(module.exports, module, module.exports, __webpack_require__);
/******/
/******/ 		// Flag the module as loaded
/******/ 		module.l = true;
/******/
/******/ 		// Return the exports of the module
/******/ 		return module.exports;
/******/ 	}
/******/
/******/
/******/ 	// expose the modules object (__webpack_modules__)
/******/ 	__webpack_require__.m = modules;
/******/
/******/ 	// expose the module cache
/******/ 	__webpack_require__.c = installedModules;
/******/
/******/ 	// define getter function for harmony exports
/******/ 	__webpack_require__.d = function(exports, name, getter) {
/******/ 		if(!__webpack_require__.o(exports, name)) {
/******/ 			Object.defineProperty(exports, name, { enumerable: true, get: getter });
/******/ 		}
/******/ 	};
/******/
/******/ 	// define __esModule on exports
/******/ 	__webpack_require__.r = function(exports) {
/******/ 		if(typeof Symbol !== 'undefined' && Symbol.toStringTag) {
/******/ 			Object.defineProperty(exports, Symbol.toStringTag, { value: 'Module' });
/******/ 		}
/******/ 		Object.defineProperty(exports, '__esModule', { value: true });
/******/ 	};
/******/
/******/ 	// create a fake namespace object
/******/ 	// mode & 1: value is a module id, require it
/******/ 	// mode & 2: merge all properties of value into the ns
/******/ 	// mode & 4: return value when already ns object
/******/ 	// mode & 8|1: behave like require
/******/ 	__webpack_require__.t = function(value, mode) {
/******/ 		if(mode & 1) value = __webpack_require__(value);
/******/ 		if(mode & 8) return value;
/******/ 		if((mode & 4) && typeof value === 'object' && value && value.__esModule) return value;
/******/ 		var ns = Object.create(null);
/******/ 		__webpack_require__.r(ns);
/******/ 		Object.defineProperty(ns, 'default', { enumerable: true, value: value });
/******/ 		if(mode & 2 && typeof value != 'string') for(var key in value) __webpack_require__.d(ns, key, function(key) { return value[key]; }.bind(null, key));
/******/ 		return ns;
/******/ 	};
/******/
/******/ 	// getDefaultExport function for compatibility with non-harmony modules
/******/ 	__webpack_require__.n = function(module) {
/******/ 		var getter = module && module.__esModule ?
/******/ 			function getDefault() { return module['default']; } :
/******/ 			function getModuleExports() { return module; };
/******/ 		__webpack_require__.d(getter, 'a', getter);
/******/ 		return getter;
/******/ 	};
/******/
/******/ 	// Object.prototype.hasOwnProperty.call
/******/ 	__webpack_require__.o = function(object, property) { return Object.prototype.hasOwnProperty.call(object, property); };
/******/
/******/ 	// __webpack_public_path__
/******/ 	__webpack_require__.p = "/";
/******/
/******/
/******/ 	// Load entry module and return exports
/******/ 	return __webpack_require__(__webpack_require__.s = 2);
/******/ })
/************************************************************************/
/******/ ([
/* 0 */
/***/ (function(module, exports) {

module.exports = OSF;

/***/ }),
/* 1 */
/***/ (function(module, exports) {

module.exports = Microsoft;

/***/ }),
/* 2 */
/***/ (function(module, __webpack_exports__, __webpack_require__) {

"use strict";
__webpack_require__.r(__webpack_exports__);

// CONCATENATED MODULE: ./src/utils/isNullOrUndefined.ts
function isNullOrUndefined(value) {
  return value === null || value === undefined;
}
// CONCATENATED MODULE: ./src/types/ExtensibilityStrings.ts

var OfficeStringJS = "office_strings.js";
var OfficeStringDebugJS = "office_strings.debug.js";
var ExtensibilityStringJS = "outlook_strings.js";
var tempWindow = window;
var ExtensibilityStrings;
function getString(string) {
  return ExtensibilityStrings[string];
}
var ExtensibilityStrings_url = "";
var baseUrl = "";
var scriptElement = null;
var stringLoadedCallback;
var stringsAreLoaded = false;

function createScriptElement(url) {
  var scriptElement = document.createElement("script");
  scriptElement.type = "text/javascript";
  scriptElement.src = url;
  return scriptElement;
}

function loadLocalizedScript(initializeAppCallback) {
  stringLoadedCallback = initializeAppCallback;
  var officeIndex;
  var scripts = document.getElementsByTagName("script");

  for (var i = 0; i < scripts.length; i++) {
    var tag = scripts.item(i);

    if (tag && tag.src) {
      var filename = tag.src || "";
      filename = filename.toLowerCase();
      officeIndex = filename.indexOf(OfficeStringJS);

      if (filename && officeIndex > 0) {
        ExtensibilityStrings_url = filename.replace(OfficeStringJS, ExtensibilityStringJS);
        baseUrl = saveBaseUrl(baseUrl, officeIndex, filename);
        break;
      }

      officeIndex = filename.indexOf(OfficeStringDebugJS);

      if (filename && officeIndex > 0) {
        ExtensibilityStrings_url = filename.replace(OfficeStringDebugJS, ExtensibilityStringJS);
        baseUrl = saveBaseUrl(baseUrl, officeIndex, filename);
        break;
      }
    }
  }

  if (ExtensibilityStrings_url) {
    var head = document.getElementsByTagName("head")[0];
    scriptElement = createScriptElement(ExtensibilityStrings_url);
    scriptElement.onload = scriptElementCallback;
    scriptElement.onreadystatechange = scriptElementCallback;
    window.setTimeout(failureCallback, 2000);
    head.appendChild(scriptElement);
  }
}

function scriptElementCallback() {
  stringsAreLoaded = true;

  if (!isNullOrUndefined(stringLoadedCallback) && (isNullOrUndefined(scriptElement.readyState) || !isNullOrUndefined(scriptElement.readyState) && (scriptElement.readyState === "loaded" || scriptElement.readyState === "complete"))) {
    scriptElement.onload = null;
    scriptElement.onreadystatechange = null;

    if (typeof tempWindow._u !== "undefined") {
      ExtensibilityStrings = tempWindow._u.ExtensibilityStrings;
    }

    stringLoadedCallback();
  }
}

function failureCallback() {
  if (!stringsAreLoaded) {
    var head = document.getElementsByTagName("head")[0];
    var fallbackUrl = baseUrl + "en-us/" + ExtensibilityStringJS;
    scriptElement.onload = null;
    scriptElement.onreadystatechange = null;
    scriptElement = createScriptElement(fallbackUrl);
    scriptElement.onload = scriptElementCallback;
    scriptElement.onreadystatechange = scriptElementCallback;
    head.appendChild(scriptElement);
  }
}

function saveBaseUrl(baseUrl, officeIndex, filename) {
  var languageUrl = filename.substring(0, officeIndex);
  var lastIndexOfSlash = languageUrl.lastIndexOf("/", languageUrl.length - 2);

  if (lastIndexOfSlash === -1) {
    lastIndexOfSlash = languageUrl.lastIndexOf("\\", languageUrl.length - 2);
  }

  if (lastIndexOfSlash !== -1 && languageUrl.length > lastIndexOfSlash + 1) {
    baseUrl = languageUrl.substring(0, lastIndexOfSlash + 1);
  }

  return baseUrl;
}
// CONCATENATED MODULE: ./src/utils/ApiTelemetryConstants.ts
var ApiTelemetryCode = function () {
  function ApiTelemetryCode() {}

  ApiTelemetryCode.success = 0;
  ApiTelemetryCode.noResponseDictionary = -900;
  ApiTelemetryCode.noErrorCodeForStandardInvokeMethod = -901;
  ApiTelemetryCode.genericProxyError = -902;
  ApiTelemetryCode.genericLegacyApiError = -903;
  ApiTelemetryCode.genericUnknownError = -904;
  return ApiTelemetryCode;
}();


// CONCATENATED MODULE: ./src/utils/getErrorForTelemetry.ts


var getErrorForTelemetry_getErrorForTelemetry = function getErrorForTelemetry(resultCode, responseDictionary) {
  if (responseDictionary) {
    if ("error" in responseDictionary) {
      if (!responseDictionary["error"]) return ApiTelemetryCode.success;
      if ("errorCode" in responseDictionary) return responseDictionary["errorCode"];else return ApiTelemetryCode.noErrorCodeForStandardInvokeMethod;
    }

    if ("wasProxySuccessful" in responseDictionary) return responseDictionary["wasProxySuccessful"] ? ApiTelemetryCode.success : ApiTelemetryCode.genericProxyError;
    if ("wasSuccessful" in responseDictionary) return responseDictionary["wasSuccessful"] ? ApiTelemetryCode.success : ApiTelemetryCode.genericLegacyApiError;
  }

  if (!isNullOrUndefined(resultCode)) return resultCode;
  return ApiTelemetryCode.genericUnknownError;
};
// CONCATENATED MODULE: ./src/utils/isOwaOnly.ts
var isOwaOnly = function isOwaOnly(dispid) {
  switch (dispid) {
    case 402:
    case 401:
    case 400:
    case 403:
      return true;

    default:
      return false;
  }
};
// CONCATENATED MODULE: ./src/utils/InvokeResultCode.ts
var InvokeResultCode;

(function (InvokeResultCode) {
  InvokeResultCode[InvokeResultCode["noError"] = 0] = "noError";
  InvokeResultCode[InvokeResultCode["errorInRequest"] = -1] = "errorInRequest";
  InvokeResultCode[InvokeResultCode["errorHandlingRequest"] = -2] = "errorHandlingRequest";
  InvokeResultCode[InvokeResultCode["errorInResponse"] = -3] = "errorInResponse";
  InvokeResultCode[InvokeResultCode["errorHandlingResponse"] = -4] = "errorHandlingResponse";
  InvokeResultCode[InvokeResultCode["errorHandlingRequestAccessDenied"] = -5] = "errorHandlingRequestAccessDenied";
  InvokeResultCode[InvokeResultCode["errorHandlingMethodCallTimedout"] = -6] = "errorHandlingMethodCallTimedout";
})(InvokeResultCode || (InvokeResultCode = {}));
// CONCATENATED MODULE: ./src/utils/getErrorArgs.ts


var getErrorArgs_OSF = __webpack_require__(0);

var isInitialized = false;
function getErrorArgs(detailedErrorCode) {
  if (!isInitialized) {
    initialize();
  }

  return getErrorArgs_OSF.DDA.ErrorCodeManager.getErrorArgs(detailedErrorCode);
}
var totalRecipientsLimit = 500;
var sessionDataLengthLimit = 50000;
function initialize() {
  addErrorMessage(9000, "AttachmentSizeExceeded", getString("l_AttachmentExceededSize_Text"));
  addErrorMessage(9001, "NumberOfAttachmentsExceeded", getString("l_ExceededMaxNumberOfAttachments_Text"));
  addErrorMessage(9002, "InternalFormatError", getString("l_InternalFormatError_Text"));
  addErrorMessage(9003, "InvalidAttachmentId", getString("l_InvalidAttachmentId_Text"));
  addErrorMessage(9004, "InvalidAttachmentPath", getString("l_InvalidAttachmentPath_Text"));
  addErrorMessage(9005, "CannotAddAttachmentBeforeUpgrade", getString("l_CannotAddAttachmentBeforeUpgrade_Text"));
  addErrorMessage(9006, "AttachmentDeletedBeforeUploadCompletes", getString("l_AttachmentDeletedBeforeUploadCompletes_Text"));
  addErrorMessage(9007, "AttachmentUploadGeneralFailure", getString("l_AttachmentUploadGeneralFailure_Text"));
  addErrorMessage(9008, "AttachmentToDeleteDoesNotExist", getString("l_DeleteAttachmentDoesNotExist_Text"));
  addErrorMessage(9009, "AttachmentDeleteGeneralFailure", getString("l_AttachmentDeleteGeneralFailure_Text"));
  addErrorMessage(9010, "InvalidEndTime", getString("l_InvalidEndTime_Text"));
  addErrorMessage(9011, "HtmlSanitizationFailure", getString("l_HtmlSanitizationFailure_Text"));
  addErrorMessage(9012, "NumberOfRecipientsExceeded", getString("l_NumberOfRecipientsExceeded_Text").replace("{0}", totalRecipientsLimit));
  addErrorMessage(9013, "NoValidRecipientsProvided", getString("l_NoValidRecipientsProvided_Text"));
  addErrorMessage(9014, "CursorPositionChanged", getString("l_CursorPositionChanged_Text"));
  addErrorMessage(9016, "InvalidSelection", getString("l_InvalidSelection_Text"));
  addErrorMessage(9017, "AccessRestricted", "");
  addErrorMessage(9018, "GenericTokenError", "");
  addErrorMessage(9019, "GenericSettingsError", "");
  addErrorMessage(9020, "GenericResponseError", "");
  addErrorMessage(9021, "SaveError", getString("l_SaveError_Text"));
  addErrorMessage(9022, "MessageInDifferentStoreError", getString("l_MessageInDifferentStoreError_Text"));
  addErrorMessage(9023, "DuplicateNotificationKey", getString("l_DuplicateNotificationKey_Text"));
  addErrorMessage(9024, "NotificationKeyNotFound", getString("l_NotificationKeyNotFound_Text"));
  addErrorMessage(9025, "NumberOfNotificationsExceeded", getString("l_NumberOfNotificationsExceeded_Text"));
  addErrorMessage(9026, "PersistedNotificationArrayReadError", getString("l_PersistedNotificationArrayReadError_Text"));
  addErrorMessage(9027, "PersistedNotificationArraySaveError", getString("l_PersistedNotificationArraySaveError_Text"));
  addErrorMessage(9028, "CannotPersistPropertyInUnsavedDraftError", getString("l_CannotPersistPropertyInUnsavedDraftError_Text"));
  addErrorMessage(9029, "CanOnlyGetTokenForSavedItem", getString("l_CallSaveAsyncBeforeToken_Text"));
  addErrorMessage(9030, "APICallFailedDueToItemChange", getString("l_APICallFailedDueToItemChange_Text"));
  addErrorMessage(9031, "InvalidParameterValueError", getString("l_InvalidParameterValueError_Text"));
  addErrorMessage(9032, "ApiCallNotSupportedByExtensionPoint", getString("l_API_Not_Supported_By_ExtensionPoint_Error_Text"));
  addErrorMessage(9033, "SetRecurrenceOnInstanceError", getString("l_Recurrence_Error_Instance_SetAsync_Text"));
  addErrorMessage(9034, "InvalidRecurrenceError", getString("l_Recurrence_Error_Properties_Invalid_Text"));
  addErrorMessage(9035, "RecurrenceZeroOccurrences", getString("l_RecurrenceErrorZeroOccurrences_Text"));
  addErrorMessage(9036, "RecurrenceMaxOccurrences", getString("l_RecurrenceErrorMaxOccurrences_Text"));
  addErrorMessage(9037, "RecurrenceInvalidTimeZone", getString("l_RecurrenceInvalidTimeZone_Text"));
  addErrorMessage(9038, "InsufficientItemPermissionsError", getString("l_Insufficient_Item_Permissions_Text"));
  addErrorMessage(9039, "RecurrenceUnsupportedAlternateCalendar", getString("l_RecurrenceUnsupportedAlternateCalendar_Text"));
  addErrorMessage(9040, "HTTPRequestFailure", getString("l_Olk_Http_Error_Text"));
  addErrorMessage(9041, "NetworkError", getString("l_Internet_Not_Connected_Error_Text"));
  addErrorMessage(9042, "InternalServerError", getString("l_Internal_Server_Error_Text"));
  addErrorMessage(9043, "AttachmentTypeNotSupported", getString("l_AttachmentNotSupported_Text"));
  addErrorMessage(9044, "InvalidCategory", getString("l_Invalid_Category_Error_Text"));
  addErrorMessage(9045, "DuplicateCategory", getString("l_Duplicate_Category_Error_Text"));
  addErrorMessage(9046, "ItemNotSaved", getString("l_Item_Not_Saved_Error_Text"));
  addErrorMessage(9047, "MissingExtendedPermissionsForAPIError", getString("l_Missing_Extended_Permissions_For_API"));
  addErrorMessage(9048, "TokenAccessDenied", getString("l_TokenAccessDeniedWithoutItemContext_Text"));
  addErrorMessage(9049, "ItemNotFound", getString("l_ItemNotFound_Text"));
  addErrorMessage(9050, "KeyNotFound", getString("l_KeyNotFound_Text"));
  addErrorMessage(9051, "SessionObjectMaxLengthExceeded", getString("l_SessionDataObjectMaxLengthExceeded_Text").replace("{0}", sessionDataLengthLimit));
  addErrorMessage(9052, "AttachmentResourceNotFound", getString("l_Attachment_Resource_Not_Found"));
  addErrorMessage(9053, "AttachmentResourceUnAuthorizedAccess", getString("l_Attachment_Resource_UnAuthorizedAccess"));
  addErrorMessage(9054, "AttachmentDownloadFailed", getString("l_Attachment_Download_Failed_Generic_Error"));
  addErrorMessage(9055, "APINotSupportedForSharedFolders", getString("l_API_Not_Supported_For_Shared_Folders_Error"));
  isInitialized = true;
}
function addErrorMessage(code, error, message) {
  getErrorArgs_OSF.DDA.ErrorCodeManager.addErrorMessage(code, {
    name: error,
    message: message
  });
}
// CONCATENATED MODULE: ./src/utils/AdditionalGlobalParameters.ts
var additionalOutlookGlobalParameters;
var getAdditionalGlobalParametersSingleton = function getAdditionalGlobalParametersSingleton() {
  return additionalOutlookGlobalParameters;
};
var recreateAdditionalGlobalParametersSingleton = function recreateAdditionalGlobalParametersSingleton(parameterBlobSupported) {
  additionalOutlookGlobalParameters = new AdditionalGlobalParameters();
  additionalOutlookGlobalParameters.parameterBlobSupported = true;
  return additionalOutlookGlobalParameters;
};

var AdditionalGlobalParameters = function () {
  function AdditionalGlobalParameters() {
    this._parameterBlobSupported = true;
    this._itemNumber = 0;
    additionalOutlookGlobalParameters = this;
  }

  Object.defineProperty(AdditionalGlobalParameters.prototype, "parameterBlobSupported", {
    set: function set(supported) {
      this._parameterBlobSupported = supported;
    },
    enumerable: true,
    configurable: true
  });

  AdditionalGlobalParameters.prototype.setActionsDefinition = function (actionsDefinitionIn) {
    this._actionsDefinition = actionsDefinitionIn;
  };

  AdditionalGlobalParameters.prototype.setCurrentItemNumber = function (itemNumberIn) {
    if (itemNumberIn > 0) {
      this._itemNumber = itemNumberIn;
    }
  };

  Object.defineProperty(AdditionalGlobalParameters.prototype, "itemNumber", {
    get: function get() {
      return this._itemNumber;
    },
    enumerable: true,
    configurable: true
  });
  Object.defineProperty(AdditionalGlobalParameters.prototype, "actionsDefinition", {
    get: function get() {
      return this._actionsDefinition;
    },
    enumerable: true,
    configurable: true
  });

  AdditionalGlobalParameters.prototype.updateOutlookExecuteParameters = function (executeParameters, additionalApiParameters) {
    var outParameters = executeParameters;

    if (this._parameterBlobSupported) {
      if (this._itemNumber > 0) {
        additionalApiParameters.itemNumber = this._itemNumber.toString();
      }

      if (this._actionsDefinition != null) {
        additionalApiParameters.actions = this.actionsDefinition;
      }

      if (Object.keys(additionalApiParameters).length === 0) {
        return outParameters;
      }

      if (outParameters == null) {
        outParameters = [];
      }

      outParameters.push(JSON.stringify(additionalApiParameters));
    }

    return outParameters;
  };

  return AdditionalGlobalParameters;
}();


// CONCATENATED MODULE: ./src/utils/callOutlookNativeDispatcher.ts


var callOutlookNativeDispatcher_OSF = __webpack_require__(0);

var callOutlookNativeDispatcher = function callOutlookNativeDispatcher(dispid, data, responseCallback) {
  var executeParameters = callOutlookNativeDispatcher_convertToOutlookNativeParameters(dispid, data);
  callOutlookNativeDispatcher_OSF.ClientHostController.execute(dispid, executeParameters, function (nativeData, resultCode) {
    var responseData = nativeData.toArray();
    var deserializedData = callOutlookNativeDispatcher_deserializeResponseData(responseData);

    if (responseCallback != null) {
      responseCallback(resultCode, deserializedData);
    }
  });
};
var callOutlookNativeDispatcher_deserializeResponseData = function deserializeResponseData(responseData) {
  if (responseData.length == 0) {
    return null;
  }

  var itemNumberFromOutlookResponse = getItemNumberFromOutlookResponse(responseData);
  var isValidItemNumberFromOutlookResponse = itemNumberFromOutlookResponse > 0;
  var itemNumberInternal = 0;

  if (getAdditionalGlobalParametersSingleton()) {
    itemNumberInternal = getAdditionalGlobalParametersSingleton().itemNumber;
  }

  var isValidItemNumberInternal = itemNumberInternal > 0;
  var itemChanged = isValidItemNumberFromOutlookResponse && isValidItemNumberInternal && itemNumberFromOutlookResponse > itemNumberInternal;
  return createDeserializedData(responseData, itemChanged);
};
var callOutlookNativeDispatcher_convertToOutlookNativeParameters = function convertToOutlookNativeParameters(dispid, data) {
  var executeParameters = null;
  var optionalParameters = {};

  switch (dispid) {
    case 12:
      optionalParameters.isRest = data.isRest;
      break;

    case 4:
      {
        var jsonProperty = JSON.stringify(data.customProperties);
        executeParameters = [jsonProperty];
        break;
      }

    case 5:
      executeParameters = new Array(data.body);
      break;

    case 8:
    case 9:
    case 179:
    case 180:
      executeParameters = new Array(data.itemId);
      break;

    case 7:
    case 177:
      executeParameters = new Array(convertRecipientArrayParameterForOutlookForDisplayApi(data.requiredAttendees), convertRecipientArrayParameterForOutlookForDisplayApi(data.optionalAttendees), data.start, data.end, data.location, convertRecipientArrayParameterForOutlookForDisplayApi(data.resources), data.subject, data.body);
      break;

    case 44:
    case 178:
      executeParameters = [convertRecipientArrayParameterForOutlookForDisplayApi(data.toRecipients), convertRecipientArrayParameterForOutlookForDisplayApi(data.ccRecipients), convertRecipientArrayParameterForOutlookForDisplayApi(data.bccRecipients), data.subject, data.htmlBody, data.attachments];
      break;

    case 43:
      executeParameters = [data.ewsIdOrEmail];
      break;

    case 45:
      executeParameters = [data.module, data.queryString];
      break;

    case 40:
      executeParameters = [data.extensionId, data.consentState];
      break;

    case 11:
    case 10:
    case 184:
    case 183:
      executeParameters = [data.htmlBody];
      break;

    case 31:
    case 30:
    case 182:
    case 181:
      executeParameters = [data.htmlBody, data.attachments];
      break;

    case 23:
    case 13:
    case 38:
    case 29:
      executeParameters = [data.data, data.coercionType];
      break;

    case 37:
    case 28:
      executeParameters = [data.coercionType];
      break;

    case 17:
      executeParameters = [data.subject];
      break;

    case 15:
      executeParameters = [data.recipientField];
      break;

    case 22:
    case 21:
      executeParameters = [data.recipientField, convertComposeEmailDictionaryParameterForSetApi(data.recipientArray)];
      break;

    case 19:
      executeParameters = [data.itemId, data.name];
      break;

    case 16:
      executeParameters = [data.uri, data.name, data.isInline];
      break;

    case 148:
      executeParameters = [data.base64String, data.name, data.isInline];
      break;

    case 20:
      executeParameters = [data.attachmentIndex];
      break;

    case 25:
      executeParameters = [data.TimeProperty, data.time];
      break;

    case 24:
      executeParameters = [data.TimeProperty];
      break;

    case 27:
      executeParameters = [data.location];
      break;

    case 33:
    case 35:
      executeParameters = [data.key, data.type, data.persistent, data.message, data.icon];
      getAdditionalGlobalParametersSingleton().setActionsDefinition(data.actions);
      break;

    case 36:
      executeParameters = [data.key];
      break;

    default:
      optionalParameters = data || {};
      break;
  }

  if (dispid !== 1) {
    executeParameters = getAdditionalGlobalParametersSingleton().updateOutlookExecuteParameters(executeParameters, optionalParameters);
  }

  return executeParameters;
};

var convertRecipientArrayParameterForOutlookForDisplayApi = function convertRecipientArrayParameterForOutlookForDisplayApi(recipients) {
  return recipients != null ? recipients.join(";") : "";
};

var convertComposeEmailDictionaryParameterForSetApi = function convertComposeEmailDictionaryParameterForSetApi(recipients) {
  var results = [];

  if (recipients == null) {
    return results;
  }

  for (var i = 0; i < recipients.length; i++) {
    var newRecipient = [recipients[i].address, recipients[i].name];
    results.push(newRecipient);
  }

  return results;
};

var getItemNumberFromOutlookResponse = function getItemNumberFromOutlookResponse(responseData) {
  var itemNumber = 0;

  if (responseData.length > 2) {
    var extraParameters = JSON.parse(responseData[2]);

    if (!!extraParameters && typeof extraParameters === "object") {
      itemNumber = extraParameters.itemNumber;
    }
  }

  return itemNumber;
};
var createDeserializedData = function createDeserializedData(responseData, itemChanged) {
  var deserializedData = null;
  var returnValues = JSON.parse(responseData[0]);

  if (typeof returnValues === "number") {
    deserializedData = createDeserializedDataWithInt(responseData, itemChanged);
  } else if (!!returnValues && typeof returnValues === "object") {
    deserializedData = createDeserializedDataWithDictionary(responseData, itemChanged);
  } else {
    throw new Error("Return data type from host must be Object or Number");
  }

  return deserializedData;
};

var createDeserializedDataWithDictionary = function createDeserializedDataWithDictionary(responseData, itemChanged) {
  var deserializedData = JSON.parse(responseData[0]);

  if (itemChanged) {
    deserializedData.error = true;
    deserializedData.errorCode = 9030;
  } else if (responseData.length > 1 && responseData[1] !== 0) {
    deserializedData.error = true;
    deserializedData.errorCode = responseData[1];

    if (responseData.length > 2) {
      var diagnosticsData = JSON.parse(responseData[2]);
      deserializedData.diagnostics = diagnosticsData["Diagnostics"];
    }
  } else {
    deserializedData.error = false;
  }

  return deserializedData;
};

var createDeserializedDataWithInt = function createDeserializedDataWithInt(responseData, itemChanged) {
  var deserializedData = {};
  deserializedData.error = true;
  deserializedData.errorCode = responseData[0];
  return deserializedData;
};
// CONCATENATED MODULE: ./src/utils/isOutlookJs.ts
var outlookJs;
outlookJs = true;
var isOutlookJs = function isOutlookJs() {
  return outlookJs;
};
// CONCATENATED MODULE: ./src/api/standardInvokeHostMethod.ts








var standardInvokeHostMethod_OSF = __webpack_require__(0);

function standardInvokeHostMethod(dispid, userContext, callback, data, format, customResponse) {
  standardInvokeHostMethod_invokeHostMethod(dispid, data, function (resultCode, response) {
    if (callback) {
      var asyncResult = void 0;
      var wasSuccessful = true;

      if (typeof response === "object" && response !== null) {
        if (response.wasSuccessful !== undefined) {
          wasSuccessful = response.wasSuccessful;
        }

        if (response.error !== undefined || response.errorCode !== undefined || response.data !== undefined) {
          if (!response.error) {
            var formattedData = format ? format(response.data) : response.data;
            asyncResult = createAsyncResult(formattedData, standardInvokeHostMethod_OSF.DDA.AsyncResultEnum.ErrorCode.Success, 0, userContext);
          } else {
            var errorCode = response.errorCode;
            asyncResult = createAsyncResult(undefined, standardInvokeHostMethod_OSF.DDA.AsyncResultEnum.ErrorCode.Failed, errorCode, userContext);
          }
        }

        if (customResponse) {
          asyncResult = customResponse(response, userContext, resultCode);
        }

        if (!asyncResult && resultCode !== InvokeResultCode.noError) {
          asyncResult = createAsyncResult(undefined, standardInvokeHostMethod_OSF.DDA.AsyncResultEnum.ErrorCode.Failed, 9002, userContext);
        }

        if (!asyncResult && resultCode === InvokeResultCode.noError && wasSuccessful === false) {
          asyncResult = createAsyncResult(undefined, standardInvokeHostMethod_OSF.DDA.AsyncResultEnum.ErrorCode.Failed, standardInvokeHostMethod_OSF.DDA.ErrorCodeManager.errorCodes.ooeOperationNotSupported, userContext);
        }

        callback(asyncResult);
      }
    }
  });
}
function createAsyncResult(value, errorCode, detailedErrorCode, userContext, errorMessage) {
  var initArgs = {};
  var errorArgs;
  initArgs[standardInvokeHostMethod_OSF.DDA.AsyncResultEnum.Properties.Value] = value;
  initArgs[standardInvokeHostMethod_OSF.DDA.AsyncResultEnum.Properties.Context] = userContext;

  if (standardInvokeHostMethod_OSF.DDA.AsyncResultEnum.ErrorCode.Success !== errorCode) {
    errorArgs = {};
    var errorProperties = void 0;
    errorProperties = getErrorArgs(detailedErrorCode);
    errorArgs[standardInvokeHostMethod_OSF.DDA.AsyncResultEnum.ErrorProperties.Name] = errorProperties.name;
    errorArgs[standardInvokeHostMethod_OSF.DDA.AsyncResultEnum.ErrorProperties.Message] = !errorMessage ? errorProperties.message : errorMessage;
    errorArgs[standardInvokeHostMethod_OSF.DDA.AsyncResultEnum.ErrorProperties.Code] = detailedErrorCode;
  }

  return new standardInvokeHostMethod_OSF.DDA.AsyncResult(initArgs, errorArgs);
}
var standardInvokeHostMethod_invokeHostMethod = function invokeHostMethod(dispid, data, responseCallback) {
  if (isOutlookJs()) {
    standardInvokeHostMethod_invokeHostMethodOutlookJs(dispid, data, responseCallback);
  } else {
    standardInvokeHostMethod_invokeHostMethodInternal(dispid, data, responseCallback);
  }
};

var standardInvokeHostMethod_invokeHostMethodInternal = function invokeHostMethodInternal(dispid, data, responseCallback) {
  if (standardInvokeHostMethod_OSF.AppName.OutlookWebApp !== getAppName() && isOwaOnly(dispid)) {
    responseCallback(InvokeResultCode.errorHandlingRequest, null);
    return;
  }

  var start = performance && performance.now();

  var invokeResponseCallback = function invokeResponseCallback(resultCode, resultData) {
    standardInvokeHostMethod_logTelemetry(resultCode, resultData, dispid, start);

    if (responseCallback) {
      responseCallback(resultCode, resultData);
    }
  };

  if (standardInvokeHostMethod_OSF.AppName.OutlookWebApp === getAppName()) {
    var args = {
      ApiParams: data,
      MethodData: {
        ControlId: standardInvokeHostMethod_OSF._OfficeAppFactory.getId(),
        DispatchId: dispid
      }
    };

    if (dispid === 1) {
      standardInvokeHostMethod_OSF._OfficeAppFactory.getClientEndPoint().invoke("GetInitialData", invokeResponseCallback, args);
    } else {
      standardInvokeHostMethod_OSF._OfficeAppFactory.getClientEndPoint().invoke("ExecuteMethod", invokeResponseCallback, args);
    }
  } else {
    callOutlookNativeDispatcher(dispid, data, invokeResponseCallback);
  }
};

var standardInvokeHostMethod_invokeHostMethodOutlookJs = function invokeHostMethodOutlookJs(dispid, data, responseCallback) {
  if (standardInvokeHostMethod_OSF.AppName.OutlookWebApp !== getAppName() && isOwaOnly(dispid)) {
    responseCallback(InvokeResultCode.errorHandlingRequest, null);
    return;
  }

  var dataTransform = standardInvokeHostMethod_createDataTransform(dispid, data);
  var start = performance && performance.now();

  standardInvokeHostMethod_OSF._OfficeAppFactory.getAsyncMethodExecutor().executeAsync(dispid, dataTransform, function (resultCode, response) {
    standardInvokeHostMethod_logTelemetry(resultCode, response, dispid, start);

    if (responseCallback) {
      var deserializedData = response;

      if (standardInvokeHostMethod_OSF.AppName.OutlookWebApp !== getAppName()) {
        deserializedData = callOutlookNativeDispatcher_deserializeResponseData(response);
      }

      responseCallback(resultCode, deserializedData);
    }
  });
};

var standardInvokeHostMethod_logTelemetry = function logTelemetry(resultCode, response, dispid, start) {
  if (standardInvokeHostMethod_OSF.AppTelemetry) {
    var detailedErrorCode = getErrorForTelemetry_getErrorForTelemetry(resultCode, response);
    var end = performance && performance.now();
    standardInvokeHostMethod_OSF.AppTelemetry.onMethodDone(dispid, null, Math.round(end - start), detailedErrorCode);
  }
};

var standardInvokeHostMethod_createDataTransform = function createDataTransform(dispid, data) {
  return {
    toSafeArrayHost: function toSafeArrayHost() {
      return callOutlookNativeDispatcher_convertToOutlookNativeParameters(dispid, data);
    },
    fromSafeArrayHost: function fromSafeArrayHost(payload) {
      return payload;
    },
    toWebHost: function toWebHost() {
      return data;
    },
    fromWebHost: function fromWebHost(payload) {
      return payload;
    }
  };
};
// CONCATENATED MODULE: ./src/utils/getPermissionLevel.ts


var getPermissionLevel_getPermissionLevel = function getPermissionLevel() {
  var permissionLevel = getInitialDataProp("permissionLevel");

  if (isNullOrUndefined(permissionLevel)) {
    return -1;
  }

  return permissionLevel;
};
// CONCATENATED MODULE: ./src/utils/createError.ts
function createError(message, errorInfo) {
  var err = new Error(message);
  err.message = message || "";

  if (errorInfo) {
    for (var v in errorInfo) {
      err[v] = errorInfo[v];
    }
  }

  return err;
}
function createBetaError(featureName) {
  var displayMessage = "The feature {0}, is only enabled on the beta api endpoint".replace("{0}", featureName);
  var err = createError(displayMessage, {
    name: "Sys.FeatureNotEnabled"
  });
  return err;
}
function createParameterCountError(message) {
  var displayMessage = "Sys.ParameterCountException: " + (message ? message : "Parameter count mismatch.");
  var err = createError(displayMessage, {
    name: "Sys.ParameterCountException"
  });
  return err;
}
function createArgumentError(paramName, message) {
  var displayMessage = "Sys.ArgumentException: " + (message ? message : "Value does not fall within the expected range.");

  if (paramName) {
    displayMessage += "\n" + "Parameter name: {0}".replace("{0}", paramName);
  }

  var err = createError(displayMessage, {
    name: "Sys.ArgumentException",
    paramName: paramName
  });
  return err;
}
function createNullItemError(namespace) {
  var displayMessage = "Invalid operation ({0}) when Office.context.mailbox.item is null.".replace("{0}", namespace);
  var err = createError(displayMessage);
  return err;
}
function createNullArgumentError(paramName, message) {
  var displayMessage = "Sys.ArgumentNullException: " + (message ? message : "Value cannot be null.");

  if (paramName) {
    displayMessage += "\n" + "Parameter name: {0}".replace("{0}", paramName);
  }

  var err = createError(displayMessage, {
    name: "Sys.ArgumentNullException",
    paramName: paramName
  });
  return err;
}
function createArgumentOutOfRange(paramName, actualValue, message) {
  var displayMessage = "Sys.ArgumentOutOfRangeException: " + (message ? message : "Specified argument was out of the range of valid values.");

  if (paramName) {
    displayMessage += "\n" + "Parameter name: {0}".replace("{0}", paramName);
  }

  if (typeof actualValue !== "undefined" && actualValue !== null) {
    displayMessage += "\n" + "Actual value was {0}.".replace("{0}", actualValue);
  }

  var err = createError(displayMessage, {
    name: "Sys.ArgumentOutOfRangeException",
    paramName: paramName,
    actualValue: actualValue
  });
  return err;
}
function createArgumentTypeError(paramName, actualType, expectedType, message) {
  var displayMessage = "Sys.ArgumentTypeException: ";

  if (message) {
    displayMessage += message;
  } else if (actualType && expectedType) {
    displayMessage += "Object of type '{0}' cannot be converted to type '{1}'.".replace("{0}", actualType.getName ? actualType.getName() : actualType).replace("{1}", expectedType.getName ? expectedType.getName() : expectedType);
  } else {
    displayMessage += "Object cannot be converted to the required type.";
  }

  if (paramName) {
    displayMessage += "\n" + "Parameter name: {0}".replace("{0}", paramName);
  }

  var err = createError(displayMessage, {
    name: "Sys.ArgumentTypeException",
    paramName: paramName,
    actualType: actualType,
    expectedType: expectedType
  });
  return err;
}
// CONCATENATED MODULE: ./src/utils/checkPermissionsAndThrow.ts



function checkPermissionsAndThrow(permissions, namespace) {
  if (getPermissionLevel_getPermissionLevel() == -1) {
    throw createNullItemError(namespace);
  }

  if (getPermissionLevel_getPermissionLevel() < permissions) {
    throw createError(getString("l_ElevatedPermissionNeededForMethod_Text").replace("{0}", namespace));
  }
}
// CONCATENATED MODULE: ./src/utils/parseCommonArgs.ts


function parseCommonArgs(args, isCallbackRequired, tryLegacy) {
  var result = {};

  if (tryLegacy) {
    result = tryParseLegacy(args);

    if (result.callback) {
      return result;
    }
  }

  if (args.length === 1) {
    if (typeof args[0] === "function") {
      result.callback = args[0];
    } else if (typeof args[0] === "object") {
      result.options = args[0];
    } else {
      throw createArgumentTypeError();
    }
  } else if (args.length === 2) {
    if (typeof args[0] !== "object") {
      throw createArgumentError("options");
    }

    if (typeof args[1] !== "function") {
      throw createArgumentError("callback");
    }

    result.callback = args[1];
    result.options = args[0];
  } else if (args.length !== 0) {
    throw createParameterCountError(getString("l_ParametersNotAsExpected_Text"));
  }

  if (isCallbackRequired && !result.callback) {
    throw createNullArgumentError("callback");
  }

  if (result.options && result.options.asyncContext) {
    result.asyncContext = result.options.asyncContext;
  }

  return result;
}

function tryParseLegacy(args) {
  var result = {};

  if (args.length === 1 || args.length === 2) {
    if (typeof args[0] !== "function") {
      return result;
    }

    result.callback = args[0];

    if (args.length === 2) {
      result.asyncContext = args[1];
    }

    return result;
  }

  return result;
}
// CONCATENATED MODULE: ./src/validation/recipientConstants.ts
var RecipientFields;

(function (RecipientFields) {
  RecipientFields[RecipientFields["to"] = 0] = "to";
  RecipientFields[RecipientFields["cc"] = 1] = "cc";
  RecipientFields[RecipientFields["bcc"] = 2] = "bcc";
  RecipientFields[RecipientFields["requiredAttendees"] = 0] = "requiredAttendees";
  RecipientFields[RecipientFields["optionalAttendees"] = 1] = "optionalAttendees";
})(RecipientFields || (RecipientFields = {}));

var displayNameLengthLimit = 255;
var recipientsLimit = 100;
var recipientConstants_totalRecipientsLimit = 500;
var maxSmtpLength = 571;
// CONCATENATED MODULE: ./src/validation/displayConstants.ts
var maxLocationLength = 255;
var maxBodyLength = 32 * 1024;
var maxSubjectLength = 255;
var maxRecipients = 100;
var MaxAttachmentNameLength = 255;
var MaxUrlLength = 2048;
var MaxItemIdLength = 200;
var MaxRemoveIdLength = 200;
// CONCATENATED MODULE: ./src/utils/throwOnOutOfRange.ts

function throwOnOutOfRange(value, minValue, maxValue, argumentName) {
  if (value < minValue || value > maxValue) {
    throw createArgumentOutOfRange(String(argumentName));
  }
}
// CONCATENATED MODULE: ./src/utils/OutlookEnums.ts
var MailboxEnums = {};
MailboxEnums.EntityType = {
  MeetingSuggestion: "meetingSuggestion",
  TaskSuggestion: "taskSuggestion",
  Address: "address",
  EmailAddress: "emailAddress",
  Url: "url",
  PhoneNumber: "phoneNumber",
  Contact: "contact",
  FlightReservations: "flightReservations",
  ParcelDeliveries: "parcelDeliveries"
};
MailboxEnums.ItemType = {
  Message: "message",
  Appointment: "appointment"
};
MailboxEnums.ResponseType = {
  None: "none",
  Organizer: "organizer",
  Tentative: "tentative",
  Accepted: "accepted",
  Declined: "declined"
};
MailboxEnums.RecipientType = {
  Other: "other",
  DistributionList: "distributionList",
  User: "user",
  ExternalUser: "externalUser"
};
MailboxEnums.AttachmentType = {
  File: "file",
  Item: "item",
  Cloud: "cloud"
};
MailboxEnums.AttachmentStatus = {
  Added: "added",
  Removed: "removed"
};
MailboxEnums.AttachmentContentFormat = {
  Base64: "base64",
  Url: "url",
  Eml: "eml",
  ICalendar: "iCalendar"
};
MailboxEnums.BodyType = {
  Text: "text",
  Html: "html"
};
MailboxEnums.ItemNotificationMessageType = {
  ProgressIndicator: "progressIndicator",
  InformationalMessage: "informationalMessage",
  ErrorMessage: "errorMessage",
  InsightMessage: "insightMessage"
};
MailboxEnums.Folder = {
  Inbox: "inbox",
  Junk: "junk",
  DeletedItems: "deletedItems"
};
MailboxEnums.ComposeType = {
  Forward: "forward",
  NewMail: "newMail",
  Reply: "reply"
};
var CoercionType = {
  Text: "text",
  Html: "html"
};
MailboxEnums.UserProfileType = {
  Office365: "office365",
  OutlookCom: "outlookCom",
  Enterprise: "enterprise"
};
MailboxEnums.RestVersion = {
  v1_0: "v1.0",
  v2_0: "v2.0",
  Beta: "beta"
};
MailboxEnums.ModuleType = {
  Addins: "addins"
};
MailboxEnums.ActionType = {
  ShowTaskPane: "showTaskPane"
};
MailboxEnums.Days = {
  Mon: "mon",
  Tue: "tue",
  Wed: "wed",
  Thu: "thu",
  Fri: "fri",
  Sat: "sat",
  Sun: "sun",
  Weekday: "weekday",
  WeekendDay: "weekendDay",
  Day: "day"
};
MailboxEnums.WeekNumber = {
  First: "first",
  Second: "second",
  Third: "third",
  Fourth: "fourth",
  Last: "last"
};
MailboxEnums.RecurrenceType = {
  Daily: "daily",
  Weekday: "weekday",
  Weekly: "weekly",
  Monthly: "monthly",
  Yearly: "yearly"
};
MailboxEnums.Month = {
  Jan: "jan",
  Feb: "feb",
  Mar: "mar",
  Apr: "apr",
  May: "may",
  Jun: "jun",
  Jul: "jul",
  Aug: "aug",
  Sep: "sep",
  Oct: "oct",
  Nov: "nov",
  Dec: "dec"
};
MailboxEnums.DelegatePermissions = {
  Read: 0x00000001,
  Write: 0x00000002,
  DeleteOwn: 0x00000004,
  DeleteAll: 0x00000008,
  EditOwn: 0x00000010,
  EditAll: 0x00000020
};
MailboxEnums.TimeZone = {
  AfghanistanStandardTime: "Afghanistan Standard Time",
  AlaskanStandardTime: "Alaskan Standard Time",
  AleutianStandardTime: "Aleutian Standard Time",
  AltaiStandardTime: "Altai Standard Time",
  ArabStandardTime: "Arab Standard Time",
  ArabianStandardTime: "Arabian Standard Time",
  ArabicStandardTime: "Arabic Standard Time",
  ArgentinaStandardTime: "Argentina Standard Time",
  AstrakhanStandardTime: "Astrakhan Standard Time",
  AtlanticStandardTime: "Atlantic Standard Time",
  AUSCentralStandardTime: "AUS Central Standard Time",
  AusCentralWStandardTime: "Aus Central W. Standard Time",
  AUSEasternStandardTime: "AUS Eastern Standard Time",
  AzerbaijanStandardTime: "Azerbaijan Standard Time",
  AzoresStandardTime: "Azores Standard Time",
  BahiaStandardTime: "Bahia Standard Time",
  BangladeshStandardTime: "Bangladesh Standard Time",
  BelarusStandardTime: "Belarus Standard Time",
  BougainvilleStandardTime: "Bougainville Standard Time",
  CanadaCentralStandardTime: "Canada Central Standard Time",
  CapeVerdeStandardTime: "Cape Verde Standard Time",
  CaucasusStandardTime: "Caucasus Standard Time",
  CenAustraliaStandardTime: "Cen. Australia Standard Time",
  CentralAmericaStandardTime: "Central America Standard Time",
  CentralAsiaStandardTime: "Central Asia Standard Time",
  CentralBrazilianStandardTime: "Central Brazilian Standard Time",
  CentralEuropeStandardTime: "Central Europe Standard Time",
  CentralEuropeanStandardTime: "Central European Standard Time",
  CentralPacificStandardTime: "Central Pacific Standard Time",
  CentralStandardTime: "Central Standard Time",
  CentralStandardTime_Mexico: "Central Standard Time (Mexico)",
  ChathamIslandsStandardTime: "Chatham Islands Standard Time",
  ChinaStandardTime: "China Standard Time",
  CubaStandardTime: "Cuba Standard Time",
  DatelineStandardTime: "Dateline Standard Time",
  EAfricaStandardTime: "E. Africa Standard Time",
  EAustraliaStandardTime: "E. Australia Standard Time",
  EEuropeStandardTime: "E. Europe Standard Time",
  ESouthAmericaStandardTime: "E. South America Standard Time",
  EasterIslandStandardTime: "Easter Island Standard Time",
  EasternStandardTime: "Eastern Standard Time",
  EasternStandardTime_Mexico: "Eastern Standard Time (Mexico)",
  EgyptStandardTime: "Egypt Standard Time",
  EkaterinburgStandardTime: "Ekaterinburg Standard Time",
  FijiStandardTime: "Fiji Standard Time",
  FLEStandardTime: "FLE Standard Time",
  GeorgianStandardTime: "Georgian Standard Time",
  GMTStandardTime: "GMT Standard Time",
  GreenlandStandardTime: "Greenland Standard Time",
  GreenwichStandardTime: "Greenwich Standard Time",
  GTBStandardTime: "GTB Standard Time",
  HaitiStandardTime: "Haiti Standard Time",
  HawaiianStandardTime: "Hawaiian Standard Time",
  IndiaStandardTime: "India Standard Time",
  IranStandardTime: "Iran Standard Time",
  IsraelStandardTime: "Israel Standard Time",
  JordanStandardTime: "Jordan Standard Time",
  KaliningradStandardTime: "Kaliningrad Standard Time",
  KamchatkaStandardTime: "Kamchatka Standard Time",
  KoreaStandardTime: "Korea Standard Time",
  LibyaStandardTime: "Libya Standard Time",
  LineIslandsStandardTime: "Line Islands Standard Time",
  LordHoweStandardTime: "Lord Howe Standard Time",
  MagadanStandardTime: "Magadan Standard Time",
  MagallanesStandardTime: "Magallanes Standard Time",
  MarquesasStandardTime: "Marquesas Standard Time",
  MauritiusStandardTime: "Mauritius Standard Time",
  MidAtlanticStandardTime: "Mid-Atlantic Standard Time",
  MiddleEastStandardTime: "Middle East Standard Time",
  MontevideoStandardTime: "Montevideo Standard Time",
  MoroccoStandardTime: "Morocco Standard Time",
  MountainStandardTime: "Mountain Standard Time",
  MountainStandardTime_Mexico: "Mountain Standard Time (Mexico)",
  MyanmarStandardTime: "Myanmar Standard Time",
  NCentralAsiaStandardTime: "N. Central Asia Standard Time",
  NamibiaStandardTime: "Namibia Standard Time",
  NepalStandardTime: "Nepal Standard Time",
  NewZealandStandardTime: "New Zealand Standard Time",
  NewfoundlandStandardTime: "Newfoundland Standard Time",
  NorfolkStandardTime: "Norfolk Standard Time",
  NorthAsiaEastStandardTime: "North Asia East Standard Time",
  NorthAsiaStandardTime: "North Asia Standard Time",
  NorthKoreaStandardTime: "North Korea Standard Time",
  OmskStandardTime: "Omsk Standard Time",
  PacificSAStandardTime: "Pacific SA Standard Time",
  PacificStandardTime: "Pacific Standard Time",
  PacificStandardTime_Mexico: "Pacific Standard Time (Mexico)",
  PakistanStandardTime: "Pakistan Standard Time",
  ParaguayStandardTime: "Paraguay Standard Time",
  RomanceStandardTime: "Romance Standard Time",
  RussiaTimeZone10: "Russia Time Zone 10",
  RussiaTimeZone11: "Russia Time Zone 11",
  RussiaTimeZone3: "Russia Time Zone 3",
  RussianStandardTime: "Russian Standard Time",
  SAEasternStandardTime: "SA Eastern Standard Time",
  SAPacificStandardTime: "SA Pacific Standard Time",
  SAWesternStandardTime: "SA Western Standard Time",
  SaintPierreStandardTime: "Saint Pierre Standard Time",
  SakhalinStandardTime: "Sakhalin Standard Time",
  SamoaStandardTime: "Samoa Standard Time",
  SaratovStandardTime: "Saratov Standard Time",
  SEAsiaStandardTime: "SE Asia Standard Time",
  SingaporeStandardTime: "Singapore Standard Time",
  SouthAfricaStandardTime: "South Africa Standard Time",
  SriLankaStandardTime: "Sri Lanka Standard Time",
  SudanStandardTime: "Sudan Standard Time",
  SyriaStandardTime: "Syria Standard Time",
  TaipeiStandardTime: "Taipei Standard Time",
  TasmaniaStandardTime: "Tasmania Standard Time",
  TocantinsStandardTime: "Tocantins Standard Time",
  TokyoStandardTime: "Tokyo Standard Time",
  TomskStandardTime: "Tomsk Standard Time",
  TongaStandardTime: "Tonga Standard Time",
  TransbaikalStandardTime: "Transbaikal Standard Time",
  TurkeyStandardTime: "Turkey Standard Time",
  TurksAndCaicosStandardTime: "Turks And Caicos Standard Time",
  UlaanbaatarStandardTime: "Ulaanbaatar Standard Time",
  USEasternStandardTime: "US Eastern Standard Time",
  USMountainStandardTime: "US Mountain Standard Time",
  UTC: "UTC",
  UTCPLUS12: "UTC+12",
  UTCPLUS13: "UTC+13",
  UTCMINUS02: "UTC-02",
  UTCMINUS08: "UTC-08",
  UTCMINUS09: "UTC-09",
  UTCMINUS11: "UTC-11",
  VenezuelaStandardTime: "Venezuela Standard Time",
  VladivostokStandardTime: "Vladivostok Standard Time",
  WAustraliaStandardTime: "W. Australia Standard Time",
  WCentralAfricaStandardTime: "W. Central Africa Standard Time",
  WEuropeStandardTime: "W. Europe Standard Time",
  WMongoliaStandardTime: "W. Mongolia Standard Time",
  WestAsiaStandardTime: "West Asia Standard Time",
  WestBankStandardTime: "West Bank Standard Time",
  WestPacificStandardTime: "West Pacific Standard Time",
  YakutskStandardTime: "Yakutsk Standard Time"
};
MailboxEnums.LocationType = {
  Custom: "custom",
  Room: "room"
};
MailboxEnums.AppointmentSensitivityType = {
  Normal: "normal",
  Personal: "personal",
  Private: "private",
  Confidential: "confidential"
};
MailboxEnums.CategoryColor = {
  None: "None",
  Preset0: "Preset0",
  Preset1: "Preset1",
  Preset2: "Preset2",
  Preset3: "Preset3",
  Preset4: "Preset4",
  Preset5: "Preset5",
  Preset6: "Preset6",
  Preset7: "Preset7",
  Preset8: "Preset8",
  Preset9: "Preset9",
  Preset10: "Preset10",
  Preset11: "Preset11",
  Preset12: "Preset12",
  Preset13: "Preset13",
  Preset14: "Preset14",
  Preset15: "Preset15",
  Preset16: "Preset16",
  Preset17: "Preset17",
  Preset18: "Preset18",
  Preset19: "Preset19",
  Preset20: "Preset20",
  Preset21: "Preset21",
  Preset22: "Preset22",
  Preset23: "Preset23",
  Preset24: "Preset24"
};
// CONCATENATED MODULE: ./src/utils/throwOnInvalidRestVersion.ts


function throwOnInvalidRestVersion(restVersion) {
  if (restVersion === null || restVersion === undefined) {
    throw createNullArgumentError(restVersion);
  }

  if (restVersion !== MailboxEnums.RestVersion.v1_0 && restVersion !== MailboxEnums.RestVersion.v2_0 && restVersion !== MailboxEnums.RestVersion.Beta) {
    throw createArgumentError(restVersion);
  }
}
// CONCATENATED MODULE: ./src/utils/convertToRestId.ts


function convertToRestId(itemId, restVersion) {
  if (itemId === null || itemId === undefined) {
    throw createNullArgumentError(itemId);
  }

  throwOnInvalidRestVersion(restVersion);
  return itemId.replace(new RegExp("[/]", "g"), "-").replace(new RegExp("[+]", "g"), "_");
}
// CONCATENATED MODULE: ./src/utils/convertToEwsId.ts


function convertToEwsId(itemId, restVersion) {
  if (itemId === null || itemId === undefined) {
    throw createNullArgumentError(itemId);
  }

  throwOnInvalidRestVersion(restVersion);
  return itemId.replace(new RegExp("[-]", "g"), "/").replace(new RegExp("[_]", "g"), "+");
}
// CONCATENATED MODULE: ./src/validation/validateDisplayForms.ts









function validateRecipientEmails(emailset, name) {
  if (!Array.isArray(emailset)) {
    throw createArgumentTypeError("name");
  }

  throwOnOutOfRange(emailset.length, 0, maxRecipients, "{0}.length".replace("{0}", name));
}
function normalizeRecipientEmails(emailset, name) {
  var originalAttendees = emailset;
  var updatedAttendees = [];

  for (var i = 0; i < originalAttendees.length; i++) {
    if (typeof originalAttendees[i] === "object") {
      throwOnInvalidEmailAddressDetails(originalAttendees[i]);
      updatedAttendees[i] = originalAttendees[i].emailAddress;

      if (typeof updatedAttendees[i] !== "string") {
        throw createArgumentError("{0}[{1}]".replace(name, String(i)));
      }
    } else {
      if (!(typeof originalAttendees[i] === "string")) {
        throw createArgumentError("{0}[{1}]".replace(name, String(i)));
      }

      updatedAttendees[i] = originalAttendees[i];
    }
  }

  return updatedAttendees;
}
function throwOnInvalidEmailAddressDetails(originalAttendee) {
  if (!isNullOrUndefined(originalAttendee.displayName)) {
    if (typeof originalAttendee.displayName === "string" && originalAttendee.displayName.length > displayNameLengthLimit) {
      throw createArgumentOutOfRange("displayName");
    }
  }

  if (!isNullOrUndefined(originalAttendee.emailAddress)) {
    if (typeof originalAttendee.emailAddress === "string" && originalAttendee.emailAddress.length > maxSmtpLength) {
      throw createArgumentOutOfRange("emailAddress");
    }
  }

  if (!isNullOrUndefined(originalAttendee.appointmentResponse)) {
    if (typeof originalAttendee.appointmentResponse !== "string") {
      throw createArgumentOutOfRange("appointmentResponse");
    }
  }

  if (!isNullOrUndefined(originalAttendee.recipientType)) {
    if (typeof originalAttendee.recipientType !== "string") {
      throw createArgumentOutOfRange("recipientType");
    }
  }
}
function validateDisplayFormParameters(itemId) {
  if (typeof itemId === "string") {
    throwOnInvalidItemId(itemId);
  } else {
    throw createArgumentTypeError("itemId");
  }
}

function throwOnInvalidItemId(itemId) {
  if (isNullOrUndefined(itemId) || itemId === "") {
    throw createNullArgumentError("itemId");
  }
}

function getItemIdBasedOnHost(itemId) {
  if (getInitialDataProp("isRestIdSupported")) {
    return convertToRestId(itemId, MailboxEnums.RestVersion.v1_0);
  }

  return convertToEwsId(itemId, MailboxEnums.RestVersion.v1_0);
}
// CONCATENATED MODULE: ./src/methods/displayAppointmentForm.ts
var __spreadArrays = undefined && undefined.__spreadArrays || function () {
  for (var s = 0, i = 0, il = arguments.length; i < il; i++) {
    s += arguments[i].length;
  }

  for (var r = Array(s), k = 0, i = 0; i < il; i++) {
    for (var a = arguments[i], j = 0, jl = a.length; j < jl; j++, k++) {
      r[k] = a[j];
    }
  }

  return r;
};





function displayAppointmentForm(itemId) {
  var args = [];

  for (var _i = 1; _i < arguments.length; _i++) {
    args[_i - 1] = arguments[_i];
  }

  displayAppointmentFormHelper.apply(void 0, __spreadArrays([9, itemId], args));
}
function displayAppointmentFormAsync(itemId) {
  var args = [];

  for (var _i = 1; _i < arguments.length; _i++) {
    args[_i - 1] = arguments[_i];
  }

  displayAppointmentFormHelper.apply(void 0, __spreadArrays([180, itemId], args));
}

function displayAppointmentFormHelper(dispidToInvoke, itemId) {
  var args = [];

  for (var _i = 2; _i < arguments.length; _i++) {
    args[_i - 2] = arguments[_i];
  }

  checkPermissionsAndThrow(1, "mailbox.displayAppointmentForm");
  var commonParameters = parseCommonArgs(args, false, false);
  var parameters = {
    itemId: itemId
  };
  validateParameters(parameters);
  standardInvokeHostMethod(dispidToInvoke, commonParameters.asyncContext, commonParameters.callback, {
    itemId: getItemIdBasedOnHost(parameters.itemId)
  }, undefined);
}

function validateParameters(parameters) {
  validateDisplayFormParameters(parameters.itemId);
}
// CONCATENATED MODULE: ./src/methods/displayMessageForm.ts
var displayMessageForm_spreadArrays = undefined && undefined.__spreadArrays || function () {
  for (var s = 0, i = 0, il = arguments.length; i < il; i++) {
    s += arguments[i].length;
  }

  for (var r = Array(s), k = 0, i = 0; i < il; i++) {
    for (var a = arguments[i], j = 0, jl = a.length; j < jl; j++, k++) {
      r[k] = a[j];
    }
  }

  return r;
};





function displayMessageForm(itemId) {
  var args = [];

  for (var _i = 1; _i < arguments.length; _i++) {
    args[_i - 1] = arguments[_i];
  }

  displayMessageFormHelper.apply(void 0, displayMessageForm_spreadArrays([8, itemId], args));
}
function displayMessageFormAsync(itemId) {
  var args = [];

  for (var _i = 1; _i < arguments.length; _i++) {
    args[_i - 1] = arguments[_i];
  }

  displayMessageFormHelper.apply(void 0, displayMessageForm_spreadArrays([179, itemId], args));
}

function displayMessageFormHelper(dispidToInvoke, itemId) {
  var args = [];

  for (var _i = 2; _i < arguments.length; _i++) {
    args[_i - 2] = arguments[_i];
  }

  checkPermissionsAndThrow(1, "mailbox.displayMessageForm");
  var commonParameters = parseCommonArgs(args, false, false);
  var parameters = {
    itemId: itemId
  };
  displayMessageForm_validateParameters(parameters);
  standardInvokeHostMethod(dispidToInvoke, commonParameters.asyncContext, commonParameters.callback, {
    itemId: getItemIdBasedOnHost(parameters.itemId)
  }, undefined);
}

function displayMessageForm_validateParameters(parameters) {
  validateDisplayFormParameters(parameters.itemId);
}
// CONCATENATED MODULE: ./src/utils/validateOptionalStringParameter.ts


function validateOptionalStringParameter(value, minLength, maxlength, name) {
  if (typeof value === "string") {
    throwOnOutOfRange(value.length, minLength, maxlength, name);
  } else {
    throw createArgumentError(String(name));
  }
}
// CONCATENATED MODULE: ./src/utils/isDateObject.ts
var isDateObject = function isDateObject(objectIn) {
  return objectIn instanceof Date || Object.prototype.toString.call(objectIn) == "[object Date]";
};
// CONCATENATED MODULE: ./src/methods/displayNewAppointmentForm.ts
var displayNewAppointmentForm_spreadArrays = undefined && undefined.__spreadArrays || function () {
  for (var s = 0, i = 0, il = arguments.length; i < il; i++) {
    s += arguments[i].length;
  }

  for (var r = Array(s), k = 0, i = 0; i < il; i++) {
    for (var a = arguments[i], j = 0, jl = a.length; j < jl; j++, k++) {
      r[k] = a[j];
    }
  }

  return r;
};











function displayNewAppointmentForm(parameters) {
  var args = [];

  for (var _i = 1; _i < arguments.length; _i++) {
    args[_i - 1] = arguments[_i];
  }

  return displayNewAppointmentFormHelper.apply(void 0, displayNewAppointmentForm_spreadArrays([7, parameters], args));
}
function displayNewAppointmentFormAsync(parameters) {
  var args = [];

  for (var _i = 1; _i < arguments.length; _i++) {
    args[_i - 1] = arguments[_i];
  }

  return displayNewAppointmentFormHelper.apply(void 0, displayNewAppointmentForm_spreadArrays([177, parameters], args));
}

function displayNewAppointmentFormHelper(dispidToInvoke, parameters) {
  var args = [];

  for (var _i = 2; _i < arguments.length; _i++) {
    args[_i - 2] = arguments[_i];
  }

  checkPermissionsAndThrow(1, "mailbox.displayNewAppointmentForm");
  var commonParameters = parseCommonArgs(args, false, false);
  displayNewAppointmentForm_validateParameters(parameters);
  var updatedParameters = normalizeParameters(parameters);
  standardInvokeHostMethod(dispidToInvoke, commonParameters.asyncContext, commonParameters.callback, updatedParameters, undefined);
}

function displayNewAppointmentForm_validateParameters(parameters) {
  if (!isNullOrUndefined(parameters.requiredAttendees)) {
    validateRecipientEmails(parameters.requiredAttendees, "requiredAttendees");
  }

  if (!isNullOrUndefined(parameters.optionalAttendees)) {
    validateRecipientEmails(parameters.optionalAttendees, "optionalAttendees");
  }

  if (!isNullOrUndefined(parameters.location)) {
    validateOptionalStringParameter(parameters.location, 0, maxLocationLength, "location");
  }

  if (!isNullOrUndefined(parameters.body)) {
    validateOptionalStringParameter(parameters.body, 0, maxBodyLength, "body");
  }

  if (!isNullOrUndefined(parameters.subject)) {
    validateOptionalStringParameter(parameters.subject, 0, maxSubjectLength, "subject");
  }

  if (!isNullOrUndefined(parameters.start)) {
    if (!isDateObject(parameters.start)) {
      throw createArgumentError("start");
    }

    if (!isNullOrUndefined(parameters.end)) {
      if (!isDateObject(parameters.end)) {
        throw createArgumentError("end");
      }

      if (parameters.end && parameters.start && parameters.end < parameters.start) {
        throw createArgumentError("end", getString("l_InvalidEventDates_Text"));
      }
    }
  }
}

function normalizeParameters(parameters) {
  var normalizedRequiredAttendees = null;
  var normalizedOptionalAttendees = null;

  if (!isNullOrUndefined(parameters.requiredAttendees)) {
    normalizedRequiredAttendees = normalizeRecipientEmails(parameters.requiredAttendees, "requiredAttendees");
  }

  if (!isNullOrUndefined(parameters.optionalAttendees)) {
    normalizedOptionalAttendees = normalizeRecipientEmails(parameters.optionalAttendees, "optionalAttendees");
  }

  if (!isNullOrUndefined(parameters.start)) {
    var startDate = parameters.start;
    parameters.start = startDate.getTime();
  }

  if (!isNullOrUndefined(parameters.end)) {
    var endDate = parameters.end;
    parameters.end = endDate.getTime();
  }

  var updatedParameters = JSON.parse(JSON.stringify(parameters));

  if (normalizedRequiredAttendees || normalizedOptionalAttendees) {
    if (!isNullOrUndefined(parameters.requiredAttendees)) {
      updatedParameters.requiredAttendees = normalizedRequiredAttendees;
    }

    if (!isNullOrUndefined(parameters.optionalAttendees)) {
      updatedParameters.optionalAttendees = normalizedOptionalAttendees;
    }
  }

  return updatedParameters;
}
// CONCATENATED MODULE: ./src/methods/displayNewMessageForm.ts
var displayNewMessageForm_spreadArrays = undefined && undefined.__spreadArrays || function () {
  for (var s = 0, i = 0, il = arguments.length; i < il; i++) {
    s += arguments[i].length;
  }

  for (var r = Array(s), k = 0, i = 0; i < il; i++) {
    for (var a = arguments[i], j = 0, jl = a.length; j < jl; j++, k++) {
      r[k] = a[j];
    }
  }

  return r;
};











function displayNewMessageForm(parameters) {
  var args = [];

  for (var _i = 1; _i < arguments.length; _i++) {
    args[_i - 1] = arguments[_i];
  }

  return displayNewMessageFormHelper.apply(void 0, displayNewMessageForm_spreadArrays([44, parameters], args));
}
function displayNewMessageFormAsync(parameters) {
  var args = [];

  for (var _i = 1; _i < arguments.length; _i++) {
    args[_i - 1] = arguments[_i];
  }

  return displayNewMessageFormHelper.apply(void 0, displayNewMessageForm_spreadArrays([178, parameters], args));
}

function displayNewMessageFormHelper(dispidToInvoke, parameters) {
  var args = [];

  for (var _i = 2; _i < arguments.length; _i++) {
    args[_i - 2] = arguments[_i];
  }

  checkPermissionsAndThrow(1, "mailbox.displayNewMessageForm");
  var commonParameters = parseCommonArgs(args, false, false);
  displayNewMessageForm_validateParameters(parameters);
  var updatedParameters = normailzeParameters(parameters);
  standardInvokeHostMethod(dispidToInvoke, commonParameters.asyncContext, commonParameters.callback, updatedParameters === null || updatedParameters === undefined ? parameters : updatedParameters, undefined);
}

function displayNewMessageForm_validateParameters(parameters) {
  if (parameters !== null && parameters !== null) {
    if (!isNullOrUndefined(parameters.toRecipients)) {
      validateRecipientEmails(parameters.toRecipients, "toRecipients");
    }

    if (!isNullOrUndefined(parameters.ccRecipients)) {
      validateRecipientEmails(parameters.ccRecipients, "ccRecipients");
    }

    if (!isNullOrUndefined(parameters.bccRecipients)) {
      validateRecipientEmails(parameters.bccRecipients, "bccRecipients");
    }

    if (!isNullOrUndefined(parameters.htmlBody)) {
      validateOptionalStringParameter(parameters.htmlBody, 0, maxBodyLength, "htmlBody");
    }

    if (!isNullOrUndefined(parameters.subject)) {
      validateOptionalStringParameter(parameters.subject, 0, maxSubjectLength, "subject");
    }
  }
}

function normailzeParameters(parameters) {
  var updatedParameters = JSON.parse(JSON.stringify(parameters));

  if (!isNullOrUndefined(parameters)) {
    if (parameters.toRecipients) {
      updatedParameters.toRecipients = normalizeRecipientEmails(parameters.toRecipients, "toRecipients");
    }

    if (parameters.ccRecipients) {
      updatedParameters.ccRecipients = normalizeRecipientEmails(parameters.ccRecipients, "ccRecipients");
    }

    if (parameters.bccRecipients) {
      updatedParameters.bccRecipients = normalizeRecipientEmails(parameters.bccRecipients, "bccRecipients");
    }

    var attachments = getAttachments(parameters);

    if (parameters.attachments) {
      updatedParameters.attachments = createAttachmentsDataForHost(attachments);
    }
  }

  return updatedParameters;
}
function getAttachments(data) {
  var attachments = [];

  if (data.attachments) {
    attachments = data.attachments;
    throwOnInvalidAttachmentsArray(attachments);
  }

  return attachments;
}
function throwOnInvalidAttachmentsArray(attachments) {
  if (!isNullOrUndefined(attachments) && !Array.isArray(attachments)) {
    throw createArgumentError("attachments");
  }
}
function createAttachmentsDataForHost(attachments) {
  var attachmentsData = [];

  for (var i = 0; i < attachments.length; i++) {
    if (typeof attachments[i] === "object") {
      var attachment = attachments[i];
      throwOnInvalidAttachment(attachment);
      attachmentsData.push(createAttachmentData(attachment));
    } else {
      throw createArgumentError("attachments");
    }
  }

  return attachmentsData;
}
function throwOnInvalidAttachment(attachment) {
  if (typeof attachment !== "object") {
    throw createArgumentError("attachments");
  }

  if (!attachment.type || !attachment.name) {
    throw createArgumentError("attachments");
  }

  if (!attachment.url && !attachment.itemId) {
    throw createArgumentError("attachments");
  }
}
function createAttachmentData(attachment) {
  var attachmentData = null;

  if (attachment.type === MailboxEnums.AttachmentType.File) {
    var url = attachment.url;
    var name_1 = attachment.name;
    var isInline = !!attachment.isInline;
    throwOnInvalidAttachmentUrlOrName(url, name_1);
    attachmentData = [MailboxEnums.AttachmentType.File, name_1, url, isInline];
  } else if (attachment.type === MailboxEnums.AttachmentType.Item) {
    var itemId = getItemIdBasedOnHost(attachment.itemId);
    var name_2 = attachment.name;
    throwOnInvalidAttachmentItemIdOrName(itemId, name_2);
    attachmentData = [MailboxEnums.AttachmentType.Item, name_2, itemId];
  } else {
    throw createArgumentError("attachments");
  }

  return attachmentData;
}
function throwOnInvalidAttachmentUrlOrName(url, name) {
  if (!(typeof url === "string") && !(typeof name === "string")) {
    throw createArgumentError("attachments");
  }

  if (url.length > MaxUrlLength) {
    throw createArgumentOutOfRange("attachments", url.length, getString("l_AttachmentUrlTooLong_Text"));
  }

  throwOnInvalidAttachmentName(name);
}
function throwOnInvalidAttachmentName(name) {
  if (name.length > MaxAttachmentNameLength) {
    throw createArgumentOutOfRange("attachments", name.length, getString("l_AttachmentNameTooLong_Text"));
  }
}
function throwOnInvalidAttachmentItemIdOrName(itemId, name) {
  if (!(typeof itemId === "string") || !(typeof name === "string")) {
    throw createArgumentError("attachments");
  }

  if (itemId.length > MaxItemIdLength) {
    throw createArgumentOutOfRange("attachments", itemId.length, getString("l_AttachmentItemIdTooLong_Text"));
  }

  throwOnInvalidAttachmentName(name);
}
// CONCATENATED MODULE: ./src/utils/handleTokenResponse.ts





var handleTokenResponse_OSF = __webpack_require__(0);

function handleTokenResponse(response, context, resultCode) {
  var asyncResult = undefined;

  if (!!resultCode && resultCode !== InvokeResultCode.noError) {
    asyncResult = createAsyncResult(undefined, handleTokenResponse_OSF.DDA.AsyncResultEnum.ErrorCode.Failed, 9017, context, getString("l_InternalProtocolError_Text").replace("{0}", resultCode));

    if (!!asyncResult) {
      asyncResult.diagnostics = {
        InvokeCodeResult: resultCode
      };
    }
  } else {
    if (getAppName() === handleTokenResponse_OSF.AppName.Outlook && response.error !== undefined && response.errorCode !== undefined && !!response.error && response.errorCode === 9030) {
      asyncResult = createAsyncResult(undefined, handleTokenResponse_OSF.DDA.AsyncResultEnum.ErrorCode.Failed, response.errorCode, context, response.errorMessage);
    } else if (!!response.wasSuccessful) {
      asyncResult = createAsyncResult(response.token, handleTokenResponse_OSF.DDA.AsyncResultEnum.ErrorCode.Success, 0, context);
    } else {
      asyncResult = createAsyncResult(undefined, handleTokenResponse_OSF.DDA.AsyncResultEnum.ErrorCode.Failed, response.errorCode, context, response.errorMessage);
    }

    if (response.diagnostics) {
      asyncResult.diagnostics = response.diagnostics;
    }
  }

  return asyncResult;
}
// CONCATENATED MODULE: ./src/methods/getCallbackToken.ts








function getCallbackToken() {
  var args = [];

  for (var _i = 0; _i < arguments.length; _i++) {
    args[_i] = arguments[_i];
  }

  checkPermissionsAndThrow(1, "mailbox.getCallbackTokenAsync");
  var commonParameters = parseCommonArgs(args, true, true);
  var isRest = false;

  if (commonParameters.options && !!commonParameters.options.isRest) {
    isRest = true;
  }

  if (getIsNoItemContextWebExt()) {
    if (!isRest || getPermissionLevel_getPermissionLevel() < 3) {
      throw createError(getString("l_TokenAccessDeniedWithoutItemContext_Text"));
    }
  }

  standardInvokeHostMethod(12, commonParameters.asyncContext, commonParameters.callback, {
    isRest: isRest
  }, undefined, handleTokenResponse);
}
// CONCATENATED MODULE: ./src/methods/getUserIdentityToken.ts




function getUserIdentityToken() {
  var args = [];

  for (var _i = 0; _i < arguments.length; _i++) {
    args[_i] = arguments[_i];
  }

  checkPermissionsAndThrow(1, "mailbox.getUserIdentityToken");
  var commonParameters = parseCommonArgs(args, true, true);
  standardInvokeHostMethod(2, commonParameters.asyncContext, commonParameters.callback, undefined, undefined, handleTokenResponse);
}
// CONCATENATED MODULE: ./src/methods/makeEwsRequest.ts







var makeEwsRequest_OSF = __webpack_require__(0);

var maxEwsRequestSize = 1000000;
function makeEwsRequest(body) {
  var args = [];

  for (var _i = 1; _i < arguments.length; _i++) {
    args[_i - 1] = arguments[_i];
  }

  checkPermissionsAndThrow(3, "mailbox.makeEwsRequest");
  var commonParameters = parseCommonArgs(args, true, true);

  if (body === null || body === undefined) {
    throw createNullArgumentError("data");
  }

  if (typeof body !== "string") {
    throw createArgumentTypeError("data", typeof body, "string");
  }

  if (body.length > maxEwsRequestSize) {
    throw createArgumentError("data", getString("l_EwsRequestOversized_Text"));
  }

  standardInvokeHostMethod(5, commonParameters.asyncContext, commonParameters.callback, {
    body: body
  }, undefined, handleCustomResponse);
}

function handleCustomResponse(data, context, responseCode) {
  if (!!responseCode && responseCode !== InvokeResultCode.noError) {
    return createAsyncResult(undefined, makeEwsRequest_OSF.DDA.AsyncResultEnum.ErrorCode.Failed, 9017, context, getString("l_InternalProtocolError_Text").replace("{0}", responseCode));
  } else if (data.wasProxySuccessful === false) {
    return createAsyncResult(undefined, makeEwsRequest_OSF.DDA.AsyncResultEnum.ErrorCode.Failed, 9020, context, data.errorMessage);
  } else {
    return createAsyncResult(data.body, makeEwsRequest_OSF.DDA.AsyncResultEnum.ErrorCode.Success, 0, context);
  }
}
// CONCATENATED MODULE: ./src/utils/objectDefine.ts
var objectDefine = function objectDefine(o, props) {
  var keys = Object.keys(props);
  var values = keys.map(function (prop) {
    return {
      value: props[prop],
      writable: false
    };
  });
  var properties = {};
  keys.forEach(function (key, index) {
    properties[key] = values[index];
  });
  return Object.defineProperties(o, properties);
};
// CONCATENATED MODULE: ./src/api/getDiagnostics.ts



var getDiagnostics_OSF = __webpack_require__(0);

var getDiagnostics_getHostName = function getHostName() {
  var appName = getAppName();

  switch (appName) {
    case getDiagnostics_OSF.AppName.Outlook:
      return "Outlook";

    case getDiagnostics_OSF.AppName.OutlookWebApp:
      return "OutlookWebApp";

    case getDiagnostics_OSF.AppName.OutlookIOS:
      return "OutlookIOS";

    case getDiagnostics_OSF.AppName.OutlookAndroid:
      return "OutlookAndroid";

    default:
      return undefined;
  }
};
function getDiagnosticsSurface() {
  return objectDefine({}, {
    hostName: getDiagnostics_getHostName(),
    hostVersion: getInitialDataProp("hostVersion"),
    OWAView: getInitialDataProp("owaView")
  });
}
// CONCATENATED MODULE: ./src/api/getUserProfile.ts


function getUserProfileSurface() {
  return objectDefine({}, {
    accountType: getInitialDataProp("userProfileType"),
    displayName: getInitialDataProp("userDisplayName"),
    emailAddress: getInitialDataProp("userEmailAddress"),
    timeZone: getInitialDataProp("userTimeZone")
  });
}
// CONCATENATED MODULE: ./src/validation/categoryConstants.ts

var CategoryColor = MailboxEnums.CategoryColor;
var categoriesCharacterLimit = 255;
var colorPresets = [CategoryColor.None, CategoryColor.Preset0, CategoryColor.Preset1, CategoryColor.Preset2, CategoryColor.Preset3, CategoryColor.Preset4, CategoryColor.Preset5, CategoryColor.Preset6, CategoryColor.Preset7, CategoryColor.Preset8, CategoryColor.Preset9, CategoryColor.Preset10, CategoryColor.Preset11, CategoryColor.Preset12, CategoryColor.Preset13, CategoryColor.Preset14, CategoryColor.Preset15, CategoryColor.Preset16, CategoryColor.Preset17, CategoryColor.Preset18, CategoryColor.Preset19, CategoryColor.Preset20, CategoryColor.Preset21, CategoryColor.Preset22, CategoryColor.Preset23, CategoryColor.Preset24];
// CONCATENATED MODULE: ./src/validation/validateCategoryDetailsArray.ts


function validateCategoryDetailsArray(categoryDetails) {
  if (!categoryDetails) {
    throw createArgumentError("categoryDetails");
  }

  if (!Array.isArray(categoryDetails)) {
    throw createArgumentTypeError("categoryDetails", typeof categoryDetails, typeof []);
  }

  if (categoryDetails.length === 0) {
    throw createArgumentError("categoryDetails");
  }

  categoryDetails.forEach(validateCategoryDetails);
}

function validateCategoryDetails(categoryDetails) {
  if (!categoryDetails) {
    throw createArgumentError("categoryDetails");
  }

  if (!categoryDetails.color || !categoryDetails.displayName) {
    throw createArgumentError("categoryDetails");
  }

  if (typeof categoryDetails.color !== "string") {
    throw createArgumentTypeError("categoryDetails.color", typeof categoryDetails.color, "string");
  }

  if (typeof categoryDetails.displayName !== "string") {
    throw createArgumentTypeError("categoryDetails.displayName", typeof categoryDetails.displayName, "string");
  }

  if (categoryDetails.displayName.length > categoriesCharacterLimit) {
    throw createArgumentOutOfRange("categoryDetails.displayName", categoryDetails.displayName.length);
  }

  if (colorPresets.indexOf(categoryDetails.color) === -1) {
    throw createArgumentError("categoryDetails.color");
  }
}
// CONCATENATED MODULE: ./src/methods/addMasterCategories.ts




function addMasterCategories(categoryDetails) {
  var args = [];

  for (var _i = 1; _i < arguments.length; _i++) {
    args[_i - 1] = arguments[_i];
  }

  checkPermissionsAndThrow(3, "masterCategories.addAsync");
  var commonParameters = parseCommonArgs(args, false, false);
  var parameters = {
    categoryDetails: categoryDetails
  };
  validateCategoryDetailsArray(categoryDetails);
  standardInvokeHostMethod(161, commonParameters.asyncContext, commonParameters.callback, parameters, undefined);
}
// CONCATENATED MODULE: ./src/methods/getMasterCategories.ts



function getMasterCategories() {
  var args = [];

  for (var _i = 0; _i < arguments.length; _i++) {
    args[_i] = arguments[_i];
  }

  checkPermissionsAndThrow(3, "masterCategories.getAsync");
  var commonParameters = parseCommonArgs(args, true, false);
  standardInvokeHostMethod(160, commonParameters.asyncContext, commonParameters.callback, undefined, undefined);
}
// CONCATENATED MODULE: ./src/validation/validateCategoryArray.ts


function validateCategoryArray(categories) {
  if (!categories) {
    throw createArgumentError("categories");
  }

  if (!Array.isArray(categories)) {
    throw createArgumentTypeError("categories", typeof categories, typeof Array);
  }

  if (categories.length === 0) {
    throw createArgumentError("categories");
  }

  categories.forEach(validateCategory);
}

function validateCategory(category) {
  if (!category) {
    throw createArgumentError("categories");
  }

  if (typeof category !== "string") {
    throw createArgumentTypeError("categories", typeof category, "string");
  }

  if (category.length > categoriesCharacterLimit) {
    throw createArgumentOutOfRange("categories", category.length);
  }
}
// CONCATENATED MODULE: ./src/methods/removeMasterCategories.ts




function removeMasterCategories(categories) {
  var args = [];

  for (var _i = 1; _i < arguments.length; _i++) {
    args[_i - 1] = arguments[_i];
  }

  checkPermissionsAndThrow(3, "masterCategories.removeAsync");
  var commonParameters = parseCommonArgs(args, false, false);
  var parameters = {
    categories: categories
  };
  validateCategoryArray(categories);
  standardInvokeHostMethod(162, commonParameters.asyncContext, commonParameters.callback, parameters, undefined);
}
// CONCATENATED MODULE: ./src/api/getMasterCategoriesSurface.ts




function getMasterCategoriesSurface() {
  return objectDefine({}, {
    addAsync: addMasterCategories,
    getAsync: getMasterCategories,
    removeAsync: removeMasterCategories
  });
}
// CONCATENATED MODULE: ./src/methods/closeApp.ts

function closeApp() {
  standardInvokeHostMethod(42, undefined, undefined, undefined, undefined);
}
// CONCATENATED MODULE: ./src/utils/getHostItemType.ts

var getHostItemType_getHostItemType = function getHostItemType() {
  return getInitialDataProp("itemType");
};
// CONCATENATED MODULE: ./src/utils/HostItemType.ts
var HostItemType;

(function (HostItemType) {
  HostItemType[HostItemType["Message"] = 1] = "Message";
  HostItemType[HostItemType["Appointment"] = 2] = "Appointment";
  HostItemType[HostItemType["MeetingRequest"] = 3] = "MeetingRequest";
  HostItemType[HostItemType["MessageCompose"] = 4] = "MessageCompose";
  HostItemType[HostItemType["AppointmentCompose"] = 5] = "AppointmentCompose";
  HostItemType[HostItemType["ItemLess"] = 6] = "ItemLess";
})(HostItemType || (HostItemType = {}));
// CONCATENATED MODULE: ./src/methods/getInitializationContext.ts



function getInitializationContext() {
  var args = [];

  for (var _i = 0; _i < arguments.length; _i++) {
    args[_i] = arguments[_i];
  }

  checkPermissionsAndThrow(1, "item.getInitializationContext");
  var commonParameters = parseCommonArgs(args, true, false);
  standardInvokeHostMethod(99, commonParameters.asyncContext, commonParameters.callback, undefined, undefined);
}
// CONCATENATED MODULE: ./src/validation/customPropertiesConstants.ts
var DatePrefix = "Date(";
var DatePostfix = ")";
var MaxCustomPropertiesLength = 2500;
var CustomPropertyType;

(function (CustomPropertyType) {
  CustomPropertyType[CustomPropertyType["NonTransmittable"] = 0] = "NonTransmittable";
})(CustomPropertyType || (CustomPropertyType = {}));
// CONCATENATED MODULE: ./src/methods/saveCustomProperties.ts





function saveCustomProperties(customProperties) {
  var args = [];

  for (var _i = 1; _i < arguments.length; _i++) {
    args[_i - 1] = arguments[_i];
  }

  checkPermissionsAndThrow(1, "item.saveCustomProperties");
  var commonParameters = parseCommonArgs(args, false, true);
  saveCustomProperties_validateParameters(customProperties);
  standardInvokeHostMethod(4, commonParameters.asyncContext, commonParameters.callback, {
    customProperties: customProperties
  }, undefined);
}

function saveCustomProperties_validateParameters(customProperties) {
  if (JSON.stringify(customProperties).length > MaxCustomPropertiesLength) {
    throw createArgumentOutOfRange("customProperties");
  }
}
// CONCATENATED MODULE: ./src/api/CustomProperties.ts
var CustomProperties_spreadArrays = undefined && undefined.__spreadArrays || function () {
  for (var s = 0, i = 0, il = arguments.length; i < il; i++) {
    s += arguments[i].length;
  }

  for (var r = Array(s), k = 0, i = 0; i < il; i++) {
    for (var a = arguments[i], j = 0, jl = a.length; j < jl; j++, k++) {
      r[k] = a[j];
    }
  }

  return r;
};







var CustomProperties_CustomProperties = function () {
  function CustomProperties(deserializedData) {
    if (isNullOrUndefined(deserializedData)) {
      createNullArgumentError("data");
    }

    if (Array.isArray(deserializedData)) {
      var customropertiesArray = deserializedData;

      if (customropertiesArray.length > CustomPropertyType.NonTransmittable) {
        deserializedData = customropertiesArray[CustomPropertyType.NonTransmittable];
      } else {
        throw createArgumentError("data");
      }
    } else {
      this.rawData = deserializedData;
    }
  }

  CustomProperties.prototype.get = function (key) {
    var value = this.rawData[key];

    if (typeof value === "string") {
      var valueString = value;

      if (valueString.length > DatePrefix.length + DatePostfix.length && valueString.startsWith(DatePrefix) && valueString.endsWith(DatePostfix)) {
        var ticksString = valueString.substring(DatePrefix.length, valueString.length - 1);
        var ticks = parseInt(ticksString);

        if (!isNaN(ticks)) {
          var dateTimeValue = new Date(ticks);

          if (!isNullOrUndefined(dateTimeValue)) {
            value = dateTimeValue;
          }
        }
      }
    }

    return value;
  };

  CustomProperties.prototype.set = function (key, value) {
    if (isDateObject(value)) {
      value = DatePrefix + value.getTime() + DatePostfix;
    }

    this.rawData[key] = value;
  };

  CustomProperties.prototype.remove = function (key) {
    delete this.rawData[key];
  };

  CustomProperties.prototype.saveAsync = function () {
    var args = [];

    for (var _i = 0; _i < arguments.length; _i++) {
      args[_i] = arguments[_i];
    }

    saveCustomProperties.apply(void 0, CustomProperties_spreadArrays([this.rawData], args));
  };

  CustomProperties.prototype.getAll = function () {
    var _this = this;

    var dictionary = {};
    var keys = Object.keys(this.rawData);
    keys.forEach(function (key) {
      dictionary[key] = _this.get(key);
    });
    return dictionary;
  };

  return CustomProperties;
}();


// CONCATENATED MODULE: ./src/methods/loadCustomProperties.ts






var loadCustomProperties_OSF = __webpack_require__(0);

function loadCustomProperties() {
  var args = [];

  for (var _i = 0; _i < arguments.length; _i++) {
    args[_i] = arguments[_i];
  }

  var commonParameters = parseCommonArgs(args, true, true);
  standardInvokeHostMethod(3, commonParameters.asyncContext, commonParameters.callback, undefined, undefined, loadCustomProperties_handleCustomResponse);
}

function loadCustomProperties_handleCustomResponse(data, context, responseCode) {
  if (typeof responseCode !== "undefined" && responseCode !== InvokeResultCode.noError) {
    return createAsyncResult(undefined, loadCustomProperties_OSF.DDA.AsyncResultEnum.ErrorCode.Failed, 9017, context, getString("l_InternalProtocolError_Text").replace("{0}", responseCode));
  } else if (data.wasSuccessful) {
    var props = JSON.parse(data.customProperties);
    var value = new CustomProperties_CustomProperties(props);
    return createAsyncResult(value, loadCustomProperties_OSF.DDA.AsyncResultEnum.ErrorCode.Success, 0, context);
  } else {
    return createAsyncResult(undefined, loadCustomProperties_OSF.DDA.AsyncResultEnum.ErrorCode.Failed, 9020, context, data.errorMessage);
  }
}
// CONCATENATED MODULE: ./src/utils/bodyUtils.ts



var bodyUtils_OSF = __webpack_require__(0);

var HostCoercionType;

(function (HostCoercionType) {
  HostCoercionType[HostCoercionType["Text"] = 0] = "Text";
  HostCoercionType[HostCoercionType["Html"] = 3] = "Html";
})(HostCoercionType || (HostCoercionType = {}));

function addCoercionTypeParameter(parameters, args) {
  if (!!args.options && typeof args.options.coercionType === "string") {
    parameters.coercionType = getCoercionTypeFromString(args.options.coercionType);
  } else {
    parameters.coercionType = HostCoercionType.Text;
  }
}
function getCoercionTypeFromString(coercionType) {
  if (coercionType === CoercionType.Html) {
    return HostCoercionType.Html;
  } else if (coercionType === CoercionType.Text) {
    return HostCoercionType.Text;
  } else {
    return undefined;
  }
}
function invokeCallbackWithCoercionTypeError(args) {
  args.callback && args.callback(createAsyncResult(undefined, bodyUtils_OSF.DDA.AsyncResultEnum.ErrorCode.Failed, 1000, args.asyncContext));
}
// CONCATENATED MODULE: ./src/methods/getBody.ts





function getBody(coercionType) {
  var args = [];

  for (var _i = 1; _i < arguments.length; _i++) {
    args[_i - 1] = arguments[_i];
  }

  checkPermissionsAndThrow(1, "body.getAsync");
  var commonParameters = parseCommonArgs(args, true, false);
  var parameters = {
    coercionType: getCoercionTypeFromString(coercionType)
  };

  if (parameters.coercionType === undefined) {
    throw createArgumentError("coercionType");
  }

  standardInvokeHostMethod(37, commonParameters.asyncContext, commonParameters.callback, parameters, undefined);
}
// CONCATENATED MODULE: ./src/methods/getBodyType.ts



function getBodyType() {
  var args = [];

  for (var _i = 0; _i < arguments.length; _i++) {
    args[_i] = arguments[_i];
  }

  checkPermissionsAndThrow(1, "body.getTypeAsync");
  var commonParameters = parseCommonArgs(args, true, false);
  standardInvokeHostMethod(14, commonParameters.asyncContext, commonParameters.callback, undefined, undefined);
}
// CONCATENATED MODULE: ./src/validation/validateBodyParameters.ts

var maxDataLengthForBodyApi = 1000000;
var maxAppendOnSendLength = 5000;
var maxDataLengthForSignatureBodyApi = 120000;
function validateAppendOnSendBodyParamters(parameters) {
  if (typeof parameters.appendTxt !== "string") {
    throw createArgumentTypeError("data", typeof parameters.appendTxt, "string");
  }

  if (parameters.appendTxt.length > maxAppendOnSendLength) {
    throw createArgumentOutOfRange("data", parameters.appendTxt.length);
  }
}
function validateBodyParameters(parameters) {
  if (typeof parameters.data !== "string") {
    throw createArgumentTypeError("data", typeof parameters.data, "string");
  }

  if (parameters.data.length > maxDataLengthForBodyApi) {
    throw createArgumentOutOfRange("data", parameters.data.length);
  }
}
function validateSignatureBodyParameters(parameters) {
  if (typeof parameters.data !== "string") {
    throw createArgumentTypeError("data", typeof parameters.data, "string");
  }

  if (parameters.data.length > maxDataLengthForSignatureBodyApi) {
    throw createArgumentOutOfRange("data", parameters.data.length);
  }
}
// CONCATENATED MODULE: ./src/methods/setBody.ts





function setBody(data) {
  var args = [];

  for (var _i = 1; _i < arguments.length; _i++) {
    args[_i - 1] = arguments[_i];
  }

  checkPermissionsAndThrow(2, "body.setAsync");
  var commonParameters = parseCommonArgs(args, false, false);
  var parameters = {
    data: data
  };
  validateBodyParameters(parameters);
  addCoercionTypeParameter(parameters, commonParameters);

  if (parameters.coercionType === undefined) {
    invokeCallbackWithCoercionTypeError(commonParameters);
    return;
  }

  standardInvokeHostMethod(38, commonParameters.asyncContext, commonParameters.callback, parameters, undefined);
}
// CONCATENATED MODULE: ./src/methods/bodyPrepend.ts





function bodyPrepend(data) {
  var args = [];

  for (var _i = 1; _i < arguments.length; _i++) {
    args[_i - 1] = arguments[_i];
  }

  checkPermissionsAndThrow(2, "body.prependAsync");
  var commonParameters = parseCommonArgs(args, false, false);
  var parameters = {
    data: data
  };
  validateBodyParameters(parameters);
  addCoercionTypeParameter(parameters, commonParameters);

  if (parameters.coercionType === undefined) {
    invokeCallbackWithCoercionTypeError(commonParameters);
    return;
  }

  standardInvokeHostMethod(23, commonParameters.asyncContext, commonParameters.callback, parameters, undefined);
}
// CONCATENATED MODULE: ./src/methods/appendOnSend.ts






function appendOnSend(data) {
  var args = [];

  for (var _i = 1; _i < arguments.length; _i++) {
    args[_i - 1] = arguments[_i];
  }

  checkPermissionsAndThrow(2, "body.appendOnSendAsync");
  var commonParameters = parseCommonArgs(args, false, false);
  var parameters = {
    appendTxt: data
  };

  if (isNullOrUndefined(parameters.appendTxt)) {
    parameters.appendTxt = "";
  } else {
    validateAppendOnSendBodyParamters(parameters);
  }

  addCoercionTypeParameter(parameters, commonParameters);

  if (parameters.coercionType === undefined) {
    invokeCallbackWithCoercionTypeError(commonParameters);
    return;
  }

  standardInvokeHostMethod(100, commonParameters.asyncContext, commonParameters.callback, parameters, undefined);
}
// CONCATENATED MODULE: ./src/methods/setSelectedData.ts





function setSelectedData(dispid) {
  return function (data) {
    var args = [];

    for (var _i = 1; _i < arguments.length; _i++) {
      args[_i - 1] = arguments[_i];
    }

    checkPermissionsAndThrow(2, "body.setSelectedDataAsync");
    var commonParameters = parseCommonArgs(args, false, false);
    var parameters = {
      data: data
    };
    validateBodyParameters(parameters);
    addCoercionTypeParameter(parameters, commonParameters);

    if (parameters.coercionType === undefined) {
      invokeCallbackWithCoercionTypeError(commonParameters);
      return;
    }

    standardInvokeHostMethod(dispid, commonParameters.asyncContext, commonParameters.callback, parameters, undefined);
  };
}
// CONCATENATED MODULE: ./src/utils/RuntimeFlighting.ts

var beta = 2;
var production = 1;
var currentLevel;
currentLevel = beta;
var getCurrentLevel = function getCurrentLevel() {
  return currentLevel;
};
var Features = {
  featureSampleProduction: production,
  featureSampleBeta: beta,
  calendarItems: production,
  signature: production,
  replyCallback: beta,
  sessionData: beta
};
function isFeatureEnabled(feature) {
  return feature <= getCurrentLevel();
}
function checkFeatureEnabledAndThrow(feature, featureName) {
  if (!isFeatureEnabled(feature)) {
    throw createBetaError(featureName);
  }
}
// CONCATENATED MODULE: ./src/methods/setSignature.ts







function setSignature(data) {
  var args = [];

  for (var _i = 1; _i < arguments.length; _i++) {
    args[_i - 1] = arguments[_i];
  }

  checkPermissionsAndThrow(2, "item.body.setSignatureAsync");
  var commonParameters = parseCommonArgs(args, false, false);
  var parameters = {
    data: data
  };
  checkFeatureEnabledAndThrow(Features.signature, "setSignatureAsync");

  if (isNullOrUndefined(parameters.data)) {
    parameters.data = "";
  } else {
    validateSignatureBodyParameters(parameters);
  }

  addCoercionTypeParameter(parameters, commonParameters);

  if (parameters.coercionType === undefined) {
    invokeCallbackWithCoercionTypeError(commonParameters);
    return;
  }

  standardInvokeHostMethod(173, commonParameters.asyncContext, commonParameters.callback, parameters, undefined);
}
// CONCATENATED MODULE: ./src/api/getBodySurface.ts








function getBodySurface(isCompose) {
  var body = objectDefine({}, {
    getAsync: getBody
  });

  if (isCompose) {
    objectDefine(body, {
      appendOnSendAsync: appendOnSend,
      getTypeAsync: getBodyType,
      prependAsync: bodyPrepend,
      setAsync: setBody,
      setSelectedDataAsync: setSelectedData(13),
      setSignatureAsync: setSignature
    });
  }

  return body;
}
// CONCATENATED MODULE: ./src/methods/getAllInternetHeaders.ts



function getAllInternetHeaders() {
  var args = [];

  for (var _i = 0; _i < arguments.length; _i++) {
    args[_i] = arguments[_i];
  }

  checkPermissionsAndThrow(1, "item.getAllInternetHeadersAsync");
  var commonParameters = parseCommonArgs(args, true, false);
  standardInvokeHostMethod(168, commonParameters.asyncContext, commonParameters.callback, undefined, undefined);
}
// CONCATENATED MODULE: ./src/types/ItemNotificationMessageType.ts
var ItemNotificationMessageType;

(function (ItemNotificationMessageType) {
  ItemNotificationMessageType[ItemNotificationMessageType["informationalMessage"] = 0] = "informationalMessage";
  ItemNotificationMessageType[ItemNotificationMessageType["progressIndicator"] = 1] = "progressIndicator";
  ItemNotificationMessageType[ItemNotificationMessageType["errorMessage"] = 2] = "errorMessage";
  ItemNotificationMessageType[ItemNotificationMessageType["insightMessage"] = 3] = "insightMessage";
})(ItemNotificationMessageType || (ItemNotificationMessageType = {}));
// CONCATENATED MODULE: ./src/utils/validateString.ts


function validateStringParam(paramName, paramValue) {
  if (isNullOrUndefined(paramValue) || paramValue === "") {
    throw createNullArgumentError(paramName);
  }

  if (!(typeof paramValue === "string")) {
    throw createArgumentTypeError(paramName, typeof paramValue, "string");
  }
}
function validateStringParamWithEmptyAllowed(paramName, paramValue) {
  if (isNullOrUndefined(paramValue)) {
    throw createNullArgumentError(paramName);
  }

  if (!(typeof paramValue === "string")) {
    throw createArgumentTypeError(paramName, typeof paramValue, "string");
  }
}
// CONCATENATED MODULE: ./src/validation/notificationMessagesConstants.ts
var MaximumKeyLength = 32;
var MaximumIconLength = 32;
var MaximumMessageLength = 150;
var MaximumActionTextLength = 30;
var NotificationsKeyParameterName = "key";
var NotificationsTypeParameterName = "type";
var NotificationsIconParameterName = "icon";
var NotificationsMessageParameterName = "message";
var NotificationsPersistentParameterName = "persistent";
var NotificationsActionsDefinitionParameterName = "actions";
var NotificationsActionTypeParameterName = "actionType";
var NotificationsActionTextParameterName = "actionText";
var NotificationsActionCommandIdParameterName = "commandId";
var NotificationsActionShowTaskPaneActionId = "showTaskPane";
// CONCATENATED MODULE: ./src/validation/validateNotificationMessages.ts






function validateKey(key) {
  validateStringParam(NotificationsKeyParameterName, key);

  if (key.length > MaximumKeyLength) {
    throw createArgumentOutOfRange(NotificationsKeyParameterName, key.length);
  }
}
function validateData(data) {
  validateStringParam(NotificationsTypeParameterName, data.type);

  if (data.type === MailboxEnums.ItemNotificationMessageType.InformationalMessage) {
    validateStringParam(NotificationsIconParameterName, data.icon);

    if (data.icon.length > MaximumIconLength) {
      throw createArgumentOutOfRange(NotificationsIconParameterName, data.icon.length);
    }

    if (isNullOrUndefined(data.persistent)) {
      throw createNullArgumentError(NotificationsPersistentParameterName);
    }

    if (typeof data.persistent !== "boolean") {
      throw createArgumentTypeError(NotificationsPersistentParameterName, typeof data.persistent, "boolean");
    }

    if (!isNullOrUndefined(data.actions)) {
      throw createArgumentError(NotificationsActionsDefinitionParameterName, getString("l_ActionsDefinitionWrongNotificationMessageError_Text"));
    }
  } else if (data.type === MailboxEnums.ItemNotificationMessageType.InsightMessage) {
    validateInsightMessageParameters(data);
  } else {
    if (!isNullOrUndefined(data.icon)) {
      throw createArgumentError(NotificationsIconParameterName);
    }

    if (!isNullOrUndefined(data.persistent)) {
      throw createArgumentError(NotificationsPersistentParameterName);
    }

    if (!isNullOrUndefined(data.actions)) {
      throw createArgumentError(NotificationsActionsDefinitionParameterName, getString("l_ActionsDefinitionWrongNotificationMessageError_Text"));
    }
  }

  validateStringParam(NotificationsMessageParameterName, data.message);

  if (data.message.length > MaximumMessageLength) {
    throw createArgumentOutOfRange(NotificationsMessageParameterName, data.message.length);
  }
}

function validateInsightMessageParameters(data) {
  validateStringParam(NotificationsIconParameterName, data.icon);

  if (data.icon.length > MaximumIconLength) {
    throw createArgumentOutOfRange(NotificationsIconParameterName, data.icon.length);
  }

  if (!isNullOrUndefined(data.persistent)) {
    throw createArgumentError(NotificationsPersistentParameterName);
  }

  if (isNullOrUndefined(data.actions)) {
    throw createNullArgumentError(NotificationsActionsDefinitionParameterName);
  } else {
    validateActionsDefinitionBlob(data.actions);
  }
}

function validateActionsDefinitionBlob(actionsDefinitionBlob) {
  var actionsDefinition = extractActionsDefinition(actionsDefinitionBlob);

  if (isNullOrUndefined(actionsDefinition)) {
    return;
  }

  validateActionsDefinitionActionsType(actionsDefinition);
  validateActionsDefinitionActionsText(actionsDefinition);
}

function extractActionsDefinition(actionsDefinitionBlob) {
  var actionsDefinition = null;

  if (Array.isArray(actionsDefinitionBlob)) {
    if (actionsDefinitionBlob.length === 1) {
      actionsDefinition = actionsDefinitionBlob[0];
    } else if (actionsDefinitionBlob.length > 1) {
      throw createArgumentError(NotificationsActionsDefinitionParameterName, getString("l_ActionsDefinitionMultipleActionsError_Text"));
    }
  } else {
    throw createArgumentError(NotificationsActionsDefinitionParameterName);
  }

  return actionsDefinition;
}

function validateActionsDefinitionActionsType(actionsDefinition) {
  if (isNullOrUndefined(actionsDefinition.actionType)) {
    throw createNullArgumentError(NotificationsActionTypeParameterName);
  }

  if (NotificationsActionShowTaskPaneActionId !== actionsDefinition.actionType) {
    throw createArgumentError(NotificationsActionTypeParameterName, getString("l_InvalidActionType_Text"));
  } else {
    if (isNullOrUndefined(actionsDefinition.commandId) || typeof actionsDefinition.commandId !== "string" || actionsDefinition.commandId === "") {
      throw createArgumentError(NotificationsActionCommandIdParameterName, getString("l_InvalidCommandIdError_Text"));
    }
  }
}

function validateActionsDefinitionActionsText(actionsDefinition) {
  if (isNullOrUndefined(actionsDefinition.actionText) || actionsDefinition.actionText === "" || typeof actionsDefinition.actionText !== "string") {
    throw createNullArgumentError(NotificationsActionTextParameterName);
  }

  if (actionsDefinition.actionText.length > MaximumActionTextLength) {
    throw createArgumentOutOfRange(NotificationsActionTextParameterName, actionsDefinition.actionText.length);
  }
}
// CONCATENATED MODULE: ./src/methods/addNotificationMessage.ts







function addNotificationMessage(key, data) {
  var args = [];

  for (var _i = 2; _i < arguments.length; _i++) {
    args[_i - 2] = arguments[_i];
  }

  checkPermissionsAndThrow(1, "notificationMessages.addAsync");
  var commonParameters = parseCommonArgs(args, false, false);
  validateKey(key);
  validateData(data);
  var type = ItemNotificationMessageType[data.type];

  if (isNullOrUndefined(type)) {
    throw createArgumentError("type");
  }

  var message = data.message;
  var icon = data.icon;
  var persistent = data.persistent;
  var actions = data.actions;
  var parameters = {
    key: key,
    message: message,
    type: type,
    icon: icon,
    persistent: persistent,
    actions: actions
  };
  standardInvokeHostMethod(33, commonParameters.asyncContext, commonParameters.callback, parameters, undefined);
}
// CONCATENATED MODULE: ./src/methods/getAllNotificationMessages.ts



function getAllNotificationMessages() {
  var args = [];

  for (var _i = 0; _i < arguments.length; _i++) {
    args[_i] = arguments[_i];
  }

  checkPermissionsAndThrow(1, "notificationMessages.getAsync");
  var commonParameters = parseCommonArgs(args, true, false);
  standardInvokeHostMethod(34, commonParameters.asyncContext, commonParameters.callback, undefined, undefined);
}
// CONCATENATED MODULE: ./src/methods/removeNotificationMessage.ts




function removeNotificationMessage(key) {
  var args = [];

  for (var _i = 1; _i < arguments.length; _i++) {
    args[_i - 1] = arguments[_i];
  }

  checkPermissionsAndThrow(1, "notificationMessages.removeAsync");
  var commonParameters = parseCommonArgs(args, false, false);
  validateKey(key);
  var parameters = {
    key: key
  };
  standardInvokeHostMethod(36, commonParameters.asyncContext, commonParameters.callback, parameters, undefined);
}
// CONCATENATED MODULE: ./src/methods/replaceNotificationMessage.ts







function replaceNotificationMessage(key, data) {
  var args = [];

  for (var _i = 2; _i < arguments.length; _i++) {
    args[_i - 2] = arguments[_i];
  }

  checkPermissionsAndThrow(1, "notificationMessages.replaceAsync");
  var commonParameters = parseCommonArgs(args, false, false);
  validateKey(key);
  validateData(data);
  var type = ItemNotificationMessageType[data.type];

  if (isNullOrUndefined(type)) {
    throw createArgumentError("type");
  }

  var message = data.message;
  var icon = data.icon;
  var persistent = data.persistent;
  var actions = data.actions;
  var parameters = {
    key: key,
    message: message,
    type: type,
    icon: icon,
    persistent: persistent,
    actions: actions
  };
  standardInvokeHostMethod(35, commonParameters.asyncContext, commonParameters.callback, parameters, undefined);
}
// CONCATENATED MODULE: ./src/api/getNotificationMessagesSurface.ts





function getNotificationMessageSurface() {
  return objectDefine({}, {
    addAsync: addNotificationMessage,
    getAllAsync: getAllNotificationMessages,
    removeAsync: removeNotificationMessage,
    replaceAsync: replaceNotificationMessage
  });
}
// CONCATENATED MODULE: ./src/validation/validateDisplayReplyForm.ts





function validateStringParameters(formData) {
  if (!isNullOrUndefined(formData)) {
    throwOnOutOfRange(formData.length, 0, maxBodyLength, "htmlBody");
  }
}
function validateAndGetHtmlBody(data) {
  var htmlBody = "";

  if (data.htmlBody) {
    throwOnInvalidHtmlBody(data.htmlBody);
    htmlBody = data.htmlBody;
  }

  return htmlBody;
}
function throwOnInvalidHtmlBody(htmlBody) {
  if (!(typeof htmlBody === "string")) {
    throw createArgumentTypeError("htmlBody", typeof htmlBody, "string");
  }

  if (isNullOrUndefined(htmlBody)) {
    throw createNullArgumentError("htmlBody");
  }

  throwOnOutOfRange(htmlBody.length, 0, maxBodyLength, "htmlBody");
}
function validateAndGetAttachments(data) {
  var attachments = [];

  if (data.attachments) {
    attachments = data.attachments;
    throwOnInvalidAttachmentsArray(attachments);
  }

  return attachments;
}
// CONCATENATED MODULE: ./src/utils/getOptionsAndCallback.ts

function getOptionsAndCallback(data) {
  var args = [];

  if (!isNullOrUndefined(data.options)) {
    args[0] = data.options;
  }

  if (!isNullOrUndefined(data.callback)) {
    args[args.length] = data.callback;
  }

  return args;
}
// CONCATENATED MODULE: ./src/methods/displayReplyForm.ts
var displayReplyForm_spreadArrays = undefined && undefined.__spreadArrays || function () {
  for (var s = 0, i = 0, il = arguments.length; i < il; i++) {
    s += arguments[i].length;
  }

  for (var r = Array(s), k = 0, i = 0; i < il; i++) {
    for (var a = arguments[i], j = 0, jl = a.length; j < jl; j++, k++) {
      r[k] = a[j];
    }
  }

  return r;
};










function displayReplyForm(formData) {
  var args = [];

  for (var _i = 1; _i < arguments.length; _i++) {
    args[_i - 1] = arguments[_i];
  }

  return displayReplyFormHelper.apply(void 0, displayReplyForm_spreadArrays([false, false, formData], args));
}
function displayReplyAllForm(formData) {
  var args = [];

  for (var _i = 1; _i < arguments.length; _i++) {
    args[_i - 1] = arguments[_i];
  }

  return displayReplyFormHelper.apply(void 0, displayReplyForm_spreadArrays([true, false, formData], args));
}
function displayReplyFormAsync(formData) {
  var args = [];

  for (var _i = 1; _i < arguments.length; _i++) {
    args[_i - 1] = arguments[_i];
  }

  return displayReplyFormHelper.apply(void 0, displayReplyForm_spreadArrays([false, true, formData], args));
}
function displayReplyAllFormAsync(formData) {
  var args = [];

  for (var _i = 1; _i < arguments.length; _i++) {
    args[_i - 1] = arguments[_i];
  }

  return displayReplyFormHelper.apply(void 0, displayReplyForm_spreadArrays([true, true, formData], args));
}

function displayReplyFormHelper(isReplyAll, isAsync, formData) {
  var args = [];

  for (var _i = 3; _i < arguments.length; _i++) {
    args[_i - 3] = arguments[_i];
  }

  var dispidToInvoke;
  checkPermissionsAndThrow(1, "mailbox.displayReplyForm");
  var commonParameters = parseCommonArgs(getOptionsAndCallback(formData), false, false);

  if (isFeatureEnabled(Features.replyCallback)) {
    if (isNullOrUndefined(commonParameters) || commonParameters.options === undefined && commonParameters.callback === undefined) {
      commonParameters = parseCommonArgs(args, false, false);
    }
  }

  var parameters = {
    formData: formData
  };
  var updatedHtmlBody = null;
  var updateAttachments = null;

  if (typeof parameters.formData === "string") {
    if (isReplyAll) {
      if (isAsync) {
        dispidToInvoke = 184;
      } else {
        dispidToInvoke = 11;
      }
    } else {
      if (isAsync) {
        dispidToInvoke = 183;
      } else {
        dispidToInvoke = 10;
      }
    }

    validateStringParameters(parameters.formData);
    standardInvokeHostMethod(dispidToInvoke, commonParameters.asyncContext, commonParameters.callback, {
      htmlBody: parameters.formData
    }, undefined);
  } else if (typeof parameters.formData === "object") {
    updatedHtmlBody = validateAndGetHtmlBody(parameters.formData);
    updateAttachments = createAttachmentsDataForHost(validateAndGetAttachments(parameters.formData));

    if (isReplyAll) {
      if (isAsync) {
        dispidToInvoke = 182;
      } else {
        dispidToInvoke = 31;
      }
    } else {
      if (isAsync) {
        dispidToInvoke = 181;
      } else {
        dispidToInvoke = 30;
      }
    }

    standardInvokeHostMethod(dispidToInvoke, commonParameters.asyncContext, commonParameters.callback, {
      htmlBody: updatedHtmlBody,
      attachments: updateAttachments
    }, undefined);
  } else {
    throw createArgumentError();
  }
}
// CONCATENATED MODULE: ./src/methods/addCategories.ts




function addCategories(categories) {
  var args = [];

  for (var _i = 1; _i < arguments.length; _i++) {
    args[_i - 1] = arguments[_i];
  }

  checkPermissionsAndThrow(2, "categories.addAsync");
  var commonParameters = parseCommonArgs(args, false, false);
  var parameters = {
    categories: categories
  };
  validateCategoryArray(categories);
  standardInvokeHostMethod(158, commonParameters.asyncContext, commonParameters.callback, parameters, undefined);
}
// CONCATENATED MODULE: ./src/methods/getCategories.ts



function getCategories() {
  var args = [];

  for (var _i = 0; _i < arguments.length; _i++) {
    args[_i] = arguments[_i];
  }

  checkPermissionsAndThrow(1, "categories.getAsync");
  var commonParameters = parseCommonArgs(args, true, false);
  standardInvokeHostMethod(157, commonParameters.asyncContext, commonParameters.callback, undefined, undefined);
}
// CONCATENATED MODULE: ./src/methods/removeCategories.ts




function removeCategories(categories) {
  var args = [];

  for (var _i = 1; _i < arguments.length; _i++) {
    args[_i - 1] = arguments[_i];
  }

  checkPermissionsAndThrow(2, "categories.removeAsync");
  var commonParameters = parseCommonArgs(args, false, false);
  var parameters = {
    categories: categories
  };
  validateCategoryArray(categories);
  standardInvokeHostMethod(159, commonParameters.asyncContext, commonParameters.callback, parameters, undefined);
}
// CONCATENATED MODULE: ./src/api/getCategoriesSurface.ts




function getCategoriesSurface() {
  return objectDefine({}, {
    addAsync: addCategories,
    getAsync: getCategories,
    removeAsync: removeCategories
  });
}
// CONCATENATED MODULE: ./src/methods/getAttachmentContent.ts




function getAttachmentContent(id) {
  var args = [];

  for (var _i = 1; _i < arguments.length; _i++) {
    args[_i - 1] = arguments[_i];
  }

  checkPermissionsAndThrow(1, "item.getAttachmentContentAsync");
  var commonParameters = parseCommonArgs(args, true, false);
  var parameters = {
    id: id
  };
  getAttachmentContent_validateParameters(parameters);
  standardInvokeHostMethod(150, commonParameters.asyncContext, commonParameters.callback, parameters, undefined);
}

function getAttachmentContent_validateParameters(parameters) {
  validateStringParam("attachmentId", parameters.id);
}
// CONCATENATED MODULE: ./src/methods/moveToFolder.ts





var Folder = MailboxEnums.Folder;
function moveToFolder(destinationFolder) {
  var args = [];

  for (var _i = 1; _i < arguments.length; _i++) {
    args[_i - 1] = arguments[_i];
  }

  checkPermissionsAndThrow(3, "item.move");
  var commonParameters = parseCommonArgs(args, false, false);
  var parameters = {
    destinationFolder: destinationFolder
  };
  moveToFolder_validateParameters(destinationFolder);
  standardInvokeHostMethod(101, commonParameters.asyncContext, commonParameters.callback, parameters, undefined);
}

function moveToFolder_validateParameters(destinationFolder) {
  if (destinationFolder !== Folder.Inbox && destinationFolder !== Folder.Junk && destinationFolder !== Folder.DeletedItems) {
    throw createArgumentError("destinationFolder");
  }
}
// CONCATENATED MODULE: ./src/utils/createEmailAddressDetails.ts

var ResponseType = MailboxEnums.ResponseType;
var RecipientType = MailboxEnums.RecipientType;
var responseTypeMap = [ResponseType.None, ResponseType.Organizer, ResponseType.Tentative, ResponseType.Accepted, ResponseType.Declined];
var recipientTypeMap = [RecipientType.Other, RecipientType.DistributionList, RecipientType.User, RecipientType.ExternalUser];
var createEmailAddressDetails = function createEmailAddressDetails(input) {
  var response = input.appointmentResponse;
  var type = input.recipientType;
  var emailAddressDetails = {
    emailAddress: input.address,
    displayName: input.name
  };

  if (typeof input.appointmentResponse === "number") {
    emailAddressDetails.appointmentResponse = response < responseTypeMap.length ? responseTypeMap[response] : ResponseType.None;
  }

  if (typeof input.recipientType === "number") {
    emailAddressDetails.recipientType = type < recipientTypeMap.length ? recipientTypeMap[type] : RecipientType.Other;
  }

  return emailAddressDetails;
};
function createEmailAddressDetailsForEntity(data) {
  return createEmailAddressDetails({
    name: data.Name || "",
    address: data.UserId || ""
  });
}
// CONCATENATED MODULE: ./src/methods/getDelayDelivery.ts



function getDelayDelivery() {
  var args = [];

  for (var _i = 0; _i < arguments.length; _i++) {
    args[_i] = arguments[_i];
  }

  checkPermissionsAndThrow(1, "delayDeliveryTime.getAsync");
  var commonParameters = parseCommonArgs(args, true, false);
  standardInvokeHostMethod(166, commonParameters.asyncContext, commonParameters.callback, undefined, undefined);
}
// CONCATENATED MODULE: ./src/methods/setDelayDelivery.ts







function setDelayDelivery(dateTime) {
  var args = [];

  for (var _i = 1; _i < arguments.length; _i++) {
    args[_i - 1] = arguments[_i];
  }

  checkPermissionsAndThrow(2, "delayDeliveryTime.setAsync");
  var commonParameters = parseCommonArgs(args, false, false);
  validateParamerters(dateTime);
  standardInvokeHostMethod(167, commonParameters.asyncContext, commonParameters.callback, {
    time: dateTime.getTime()
  }, undefined);
}

function validateParamerters(dateTime) {
  if (isNullOrUndefined(dateTime)) {
    throw createNullArgumentError("dateTime", "You cannot conduct to a null dateTime");
  }

  if (!isDateObject(dateTime)) {
    throw createArgumentTypeError("dateTime", typeof dateTime, typeof Date);
  }

  if (isNaN(dateTime.getTime())) {
    throw createArgumentError("dateTime");
  }

  throwOnOutOfRange(dateTime.getTime(), -8640000000000000, 8640000000000000, "dateTime");
}
// CONCATENATED MODULE: ./src/api/getDelayDeliverySurface.ts



function getDelayDeliverySurface(isCompose) {
  var delayDelivery = objectDefine({}, {
    getAsync: getDelayDelivery
  });

  if (isCompose) {
    objectDefine(delayDelivery, {
      setAsync: setDelayDelivery
    });
  }

  return delayDelivery;
}
// CONCATENATED MODULE: ./src/utils/removeDuplicates.ts
function removeDuplicates(array, comparator) {
  for (var matchIndex1 = array.length - 1; matchIndex1 >= 0; matchIndex1--) {
    var removeMatch = false;

    for (var matchIndex2 = matchIndex1 - 1; matchIndex2 >= 0; matchIndex2--) {
      if (comparator(array[matchIndex1], array[matchIndex2])) {
        removeMatch = true;
        break;
      }
    }

    if (removeMatch) {
      array.splice(matchIndex1, 1);
    }
  }

  return array;
}
var stringComparator = function stringComparator(a, b) {
  return a === b;
};
var meetingComparator = function meetingComparator(a, b) {
  if (a === b) {
    return true;
  } else if (!a || !b) {
    return false;
  } else {
    return a.meetingString === b.meetingString;
  }
};
var taskComparator = function taskComparator(a, b) {
  if (a === b) {
    return true;
  } else if (!a || !b) {
    return false;
  } else {
    return a.taskString === b.taskString;
  }
};
var contactComparator = function contactComparator(a, b) {
  if (a === b) {
    return true;
  } else if (!a || !b) {
    return false;
  } else {
    return a.contactString === b.contactString;
  }
};
// CONCATENATED MODULE: ./src/utils/isLegacyEntityExtraction.ts

function isLegacyEntityExtraction() {
  return !!getInitialDataProp("entities") && getInitialDataProp("entities").IsLegacyExtraction !== undefined && getInitialDataProp("entities").IsLegacyExtraction;
}
// CONCATENATED MODULE: ./src/utils/resolveDate.ts


var totalBits = 18;
var typeBits = 3;
var preciseDateTypeBits = 3;
var yearBits = 7;
var monthBits = 4;
var dayBits = 5;
var modifierBits = 2;
var unitBits = 3;
var offsetBits = 6;
var tagBits = 4;
var preciseDateType = 0;
var relativeDateType = 1;
var oneDayInMilliseconds = 86400000;
var baseDate = new Date("0001-01-01T00:00:00Z");
function resolveDate(input, sentTime) {
  if (!sentTime) {
    return input;
  }

  var date = null;

  try {
    var sentDate = new Date(sentTime.getFullYear(), sentTime.getMonth(), sentTime.getDate(), 0, 0, 0, 0);
    var extractedDate = decode(input);

    if (!extractedDate) {
      return input;
    } else {
      var preciseDate = extractedDate;

      if (preciseDate.day && preciseDate.month && preciseDate.year !== undefined) {
        date = resolvePreciseDate(sentDate, extractedDate);
      } else {
        var relativeDate = extractedDate;

        if (relativeDate.modifier !== undefined && relativeDate.offset !== undefined && relativeDate.tag !== undefined && relativeDate.unit !== undefined) {
          date = resolveRelativeDate(sentDate, extractedDate);
        } else {
          date = sentDate;
        }
      }
    }

    if (isNaN(date.getTime())) {
      return sentTime;
    }

    date.setMilliseconds(date.getMilliseconds() + (isLegacyEntityExtraction() ? getTimeOfDayInMillisecondsUTC(input) : getTimeOfDayInMilliseconds(input)));
    return date;
  } catch (e) {
    return sentTime;
  }
}

function decode(input) {
  var dateValueMask = (1 << totalBits - typeBits) - 1;
  var time = 0;

  if (input == null) {
    return undefined;
  }

  if (isLegacyEntityExtraction()) {
    time = getTimeOfDayInMillisecondsUTC(input);
  } else {
    time = getTimeOfDayInMilliseconds(input);
  }

  var inDateAtMidnight = input.getTime() - time;
  var value = (inDateAtMidnight - baseDate.getTime()) / oneDayInMilliseconds;

  if (value < 0) {
    return undefined;
  } else if (value >= 1 << totalBits) {
    return undefined;
  } else {
    var type = value >> totalBits - typeBits;
    value = value & dateValueMask;

    switch (type) {
      case preciseDateType:
        return decodePreciseDate(value);

      case relativeDateType:
        return decodeRelativeDate(value);

      default:
        return undefined;
    }
  }
}

function decodePreciseDate(value) {
  var cSubTypeMask = (1 << preciseDateTypeBits) - 1;
  var cMonthMask = (1 << monthBits) - 1;
  var cDayMask = (1 << dayBits) - 1;
  var cYearMask = (1 << yearBits) - 1;
  var year = 0;
  var month = 0;
  var day = 0;
  var subType = value >> totalBits - typeBits - preciseDateTypeBits & cSubTypeMask;

  if ((subType & 4) == 4) {
    year = value >> totalBits - typeBits - preciseDateTypeBits - yearBits & cYearMask;

    if ((subType & 2) == 2) {
      if ((subType & 1) == 1) {
        return undefined;
      }

      month = value >> totalBits - typeBits - preciseDateTypeBits - yearBits - monthBits & cMonthMask;
    }
  } else {
    if ((subType & 2) == 2) {
      month = value >> totalBits - typeBits - preciseDateTypeBits - monthBits & cMonthMask;
    }

    if ((subType & 1) == 1) {
      day = value >> totalBits - typeBits - preciseDateTypeBits - monthBits - dayBits & cDayMask;
    }
  }

  return createPreciseDate(day, month, year);
}

function resolvePreciseDate(sentDate, precise) {
  var year = precise.year;
  var month = precise.month == 0 ? sentDate.getMonth() : precise.month - 1;
  var day = precise.day;

  if (day == 0) {
    return sentDate;
  }

  var candidate;

  if (isNullOrUndefined(year)) {
    candidate = new Date(sentDate.getFullYear(), month, day);

    if (candidate.getTime() < sentDate.getTime()) {
      candidate = new Date(sentDate.getFullYear() + 1, month, day);
    }
  } else {
    candidate = new Date(year < 50 ? 2000 + year : 1900 + year, month, day);
  }

  if (candidate.getMonth() != month) {
    return sentDate;
  }

  return candidate;
}

function resolveRelativeDate(sentDate, relative) {
  var date;

  switch (relative.unit) {
    case 0:
      date = new Date(sentDate.getFullYear(), sentDate.getMonth(), sentDate.getDate());
      date.setDate(date.getDate() + relative.offset);
      return date;

    case 5:
      return findBestDateForWeekDate(sentDate, relative.offset, relative.tag);

    case 2:
      {
        var days = 1;

        switch (relative.modifier) {
          case 1:
            break;

          case 2:
            days = 16;
            break;

          default:
            if (relative.offset == 0) {
              days = sentDate.getDate();
            }

            break;
        }

        date = new Date(sentDate.getFullYear(), sentDate.getMonth(), days);
        date.setMonth(date.getMonth() + relative.offset);

        if (date.getTime() < sentDate.getTime()) {
          date.setDate(date.getDate() + sentDate.getDate() - 1);
        }

        return date;
      }

    case 1:
      date = new Date(sentDate.getFullYear(), sentDate.getMonth(), sentDate.getDate());
      date.setDate(sentDate.getDate() + 7 * relative.offset);

      if (relative.modifier == 1 || relative.modifier == 0) {
        date.setDate(date.getDate() + 1 - date.getDay());

        if (date.getTime() < sentDate.getTime()) {
          return sentDate;
        }

        return date;
      } else if (relative.modifier == 2) {
        date.setDate(date.getDate() + 5 - date.getDay());
        return date;
      }

      break;

    case 4:
      return findBestDateForWeekOfMonthDate(sentDate, relative);

    case 3:
      if (relative.offset > 0) {
        return new Date(sentDate.getFullYear() + relative.offset, 0, 1);
      }

      break;

    default:
      break;
  }

  return sentDate;
}

function findBestDateForWeekDate(sentDate, offset, tag) {
  if (offset > -5 && offset < 5) {
    var dayOfWeek = (tag + 6) % 7 + 1;
    var days = 7 * offset + (dayOfWeek - sentDate.getDay());
    sentDate.setDate(sentDate.getDate() + days);
    return sentDate;
  } else {
    var days = (tag - sentDate.getDay()) % 7;

    if (days < 0) {
      days += 7;
    }

    sentDate.setDate(sentDate.getDate() + days);
    return sentDate;
  }
}

function findBestDateForWeekOfMonthDate(sentDate, relative) {
  var date;
  var firstDay;
  var newDate;
  date = sentDate;

  if (relative.tag <= 0 || relative.tag > 12 || relative.offset <= 0 || relative.offset > 5) {
    return sentDate;
  }

  var monthOffset = (12 + relative.tag - date.getMonth() - 1) % 12;
  firstDay = new Date(date.getFullYear(), date.getMonth() + monthOffset, 1);

  if (relative.modifier == 1) {
    if (relative.offset == 1 && firstDay.getDay() != 6 && firstDay.getDay() != 0) {
      return firstDay;
    } else {
      newDate = new Date(firstDay.getFullYear(), firstDay.getMonth(), firstDay.getDate());
      newDate.setDate(newDate.getDate() + (7 + (1 - firstDay.getDay())) % 7);

      if (firstDay.getDay() != 6 && firstDay.getDay() != 0 && firstDay.getDay() != 1) {
        newDate.setDate(newDate.getDate() - 7);
      }

      newDate.setDate(newDate.getDate() + 7 * (relative.offset - 1));

      if (newDate.getMonth() + 1 != relative.tag) {
        return sentDate;
      }

      return newDate;
    }
  } else {
    newDate = new Date(firstDay.getFullYear(), firstDay.getMonth(), daysInMonth(firstDay.getMonth(), firstDay.getFullYear()));
    var offset = 1 - newDate.getDay();

    if (offset > 0) {
      offset = offset - 7;
    }

    newDate.setDate(newDate.getDate() + offset);
    newDate.setDate(newDate.getDate() + 7 * (1 - relative.offset));

    if (newDate.getMonth() + 1 != relative.tag) {
      if (firstDay.getDay() != 6 && firstDay.getDay() != 0) {
        return firstDay;
      } else {
        return sentDate;
      }
    } else {
      return newDate;
    }
  }
}

function decodeRelativeDate(value) {
  var tagMask = (1 << tagBits) - 1;
  var offsetMask = (1 << offsetBits) - 1;
  var unitMask = (1 << unitBits) - 1;
  var modifierMask = (1 << modifierBits) - 1;
  var tag = value & tagMask;
  value >>= tagBits;
  var offset = fromComplement(value & offsetMask, offsetBits);
  value >>= offsetBits;
  var unit = value & unitMask;
  value >>= unitBits;
  var modifier = value & modifierMask;

  try {
    return createRelativeDate(modifier, offset, unit, tag);
  } catch (_a) {
    return undefined;
  }
}

function fromComplement(value, n) {
  var signed = 1 << n - 1;
  var mask = (1 << n) - 1;

  if ((value & signed) == signed) {
    return -((value ^ mask) + 1);
  } else {
    return value;
  }
}

function daysInMonth(month, year) {
  return 32 - new Date(year, month, 32).getDate();
}

function getTimeOfDayInMilliseconds(inputTime) {
  var timeOfDay = 0;
  timeOfDay += inputTime.getHours() * 3600;
  timeOfDay += inputTime.getMinutes() * 60;
  timeOfDay += inputTime.getSeconds();
  timeOfDay *= 1000;
  timeOfDay += inputTime.getMilliseconds();
  return timeOfDay;
}

function getTimeOfDayInMillisecondsUTC(inputTime) {
  var timeOfDay = 0;
  timeOfDay += inputTime.getUTCHours() * 3600;
  timeOfDay += inputTime.getUTCMinutes() * 60;
  timeOfDay += inputTime.getUTCSeconds();
  timeOfDay *= 1000;
  timeOfDay += inputTime.getUTCMilliseconds();
  return timeOfDay;
}

function createPreciseDate(day, month, year) {
  return {
    day: day,
    month: month,
    year: year % 100
  };
}

function createRelativeDate(modifier, offset, unit, tag) {
  return {
    modifier: modifier,
    offset: offset,
    unit: unit,
    tag: tag
  };
}
// CONCATENATED MODULE: ./src/utils/findOffset.ts



function findOffset(value) {
  var ranges = getInitialDataProp("timeZoneOffsets");

  for (var r = 0; r < ranges.length; r++) {
    var range = ranges[r];
    var start = parseInt(range.start);
    var end = parseInt(range.end);

    if (value.getTime() - start >= 0 && value.getTime() - end < 0) {
      return parseInt(range.offset);
    }
  }

  throw createArgumentError("input", getString("l_InvalidDate_Text"));
}
// CONCATENATED MODULE: ./src/utils/convertToUtcClientTime.ts





function convertToUtcClientTime(input) {
  var retValue = localClientTimeToDate(input);

  if (!isNullOrUndefined(getInitialDataProp("timeZoneOffsets"))) {
    var offset = findOffset(retValue);
    retValue.setUTCMinutes(retValue.getUTCMinutes() - offset);
    offset = !input["timezoneOffset"] ? retValue.getTimezoneOffset() * -1 : input["timezoneOffset"];
    retValue.setUTCMinutes(retValue.getUTCMinutes() + offset);
  }

  return retValue;
}
function localClientTimeToDate(input) {
  var retValue = new Date(input["year"], input["month"], input["date"], input["hours"], input["minutes"], input["seconds"], input["milliseconds"] === null ? 0 : input["milliseconds"]);

  if (isNaN(retValue.getTime())) {
    throw createArgumentError("input", getString("l_InvalidDate_Text"));
  }

  return retValue;
}
// CONCATENATED MODULE: ./src/utils/dateToDictionary.ts
function dateToDictionary(date) {
  return {
    month: date.getMonth(),
    date: date.getDate(),
    year: date.getFullYear(),
    hours: date.getHours(),
    minutes: date.getMinutes(),
    seconds: date.getSeconds(),
    milliseconds: date.getMilliseconds()
  };
}
// CONCATENATED MODULE: ./src/utils/createEntities.ts









var EntityKeys;

(function (EntityKeys) {
  EntityKeys["meetingSuggestion"] = "MeetingSuggestions";
  EntityKeys["taskSuggestion"] = "TaskSuggestions";
  EntityKeys["address"] = "Addresses";
  EntityKeys["emailAddress"] = "EmailAddresses";
  EntityKeys["url"] = "Urls";
  EntityKeys["phoneNumber"] = "PhoneNumbers";
  EntityKeys["contact"] = "Contacts";
  EntityKeys["flightReservations"] = "FlightReservations";
  EntityKeys["parcelDeliveries"] = "ParcelDeliveries";
})(EntityKeys || (EntityKeys = {}));

function createEntities(data) {
  if (isNullOrUndefined(data)) {
    return {
      addresses: [],
      emailAddresses: [],
      urls: [],
      taskSuggestions: [],
      meetingSuggestions: [],
      phoneNumbers: [],
      contacts: [],
      flightReservations: [],
      parcelDelivery: []
    };
  } else {
    return {
      addresses: createEntities_createAddresses(data[EntityKeys.address]),
      emailAddresses: createEntities_createEmailAddresses(data[EntityKeys.emailAddress]),
      urls: createUrls(data[EntityKeys.url]),
      taskSuggestions: createEntities_createTaskSuggestions(data[EntityKeys.taskSuggestion]),
      meetingSuggestions: createEntities_createMeetingSuggestions(data[EntityKeys.meetingSuggestion]),
      phoneNumbers: createPhoneNumbers(data[EntityKeys.phoneNumber]),
      contacts: createEntities_createContacts(data[EntityKeys.contact]),
      flightReservations: createEntities_createReadItemArray(data[EntityKeys.flightReservations]),
      parcelDelivery: createEntities_createReadItemArray(data[EntityKeys.parcelDeliveries])
    };
  }
}
function createFilteredEntities(data, name) {
  checkPermissionsAndThrow(1, "item.getFilteredEntitiesByName");
  var results = Object.keys(data).map(function (entities) {
    var results = data[entities][name];
    if (results) return {
      entityType: entities,
      name: name,
      entities: data[entities][name]
    };else return undefined;
  }).filter(function (results) {
    return results !== undefined;
  });

  if (results.length === 0) {
    return null;
  }

  var matchedRule = results[0];

  switch (matchedRule.entityType) {
    case EntityKeys.meetingSuggestion:
      return createEntities_createMeetingSuggestions(matchedRule.entities);

    case EntityKeys.address:
      return createEntities_createAddresses(matchedRule.entities);

    case EntityKeys.contact:
      return createEntities_createContacts(matchedRule.entities);

    case EntityKeys.emailAddress:
      return createEntities_createEmailAddresses(matchedRule.entities);

    case EntityKeys.phoneNumber:
      return createPhoneNumbers(matchedRule.entities);

    case EntityKeys.taskSuggestion:
      return createEntities_createTaskSuggestions(matchedRule.entities);

    case EntityKeys.url:
      return createUrls(matchedRule.entities);

    default:
      return createEntities_createReadItemArray(matchedRule.entities);
  }
}
var createEntities_createAddresses = function createAddresses(data) {
  var addresses = data || [];
  return removeDuplicates(addresses, stringComparator);
};
var createEntities_createEmailAddresses = function createEmailAddresses(data) {
  if (getPermissionLevel_getPermissionLevel() === 0) {
    return [];
  }

  return data || [];
};
var createUrls = function createUrls(data) {
  return data || [];
};
var createEntities_createTaskSuggestions = function createTaskSuggestions(data) {
  if (getPermissionLevel_getPermissionLevel() === 0) {
    return [];
  }

  var tasks = data || [];
  tasks = tasks.map(function (task) {
    return {
      assignees: (task.Assignees || []).map(createEmailAddressDetailsForEntity),
      taskString: task.TaskString
    };
  });
  return removeDuplicates(tasks, taskComparator);
};
var createEntities_createMeetingSuggestions = function createMeetingSuggestions(data) {
  if (getPermissionLevel_getPermissionLevel() === 0) {
    return [];
  }

  var meetings = data || [];
  meetings = meetings.map(function (meeting) {
    var start = meeting.StartTime !== "" ? getDate(meeting.StartTime) : undefined;
    var end = meeting.EndTime !== "" ? getDate(meeting.EndTime) : undefined;
    return {
      meetingString: meeting.MeetingString,
      attendees: (meeting.Attendees || []).map(createEmailAddressDetailsForEntity),
      location: meeting.Location,
      subject: meeting.Subject,
      start: meeting.StartTime !== undefined ? start : undefined,
      end: meeting.EndTime !== undefined ? end : undefined
    };
  });
  return removeDuplicates(meetings, meetingComparator);
};

function getDate(date) {
  var result = resolveDate(new Date(date), new Date(getInitialDataProp("dateTimeSent")));

  if (result.getTime() !== new Date(date).getTime()) {
    return convertToUtcClientTime(dateToDictionary(result));
  }

  return new Date(date);
}

var createPhoneNumbers = function createPhoneNumbers(data) {
  var phoneNumbers = data || [];
  return phoneNumbers.map(function (number) {
    return {
      phoneString: number.PhoneString,
      originalPhoneString: number.OriginalPhoneString,
      type: number.Type
    };
  });
};
var createEntities_createContacts = function createContacts(data) {
  if (getPermissionLevel_getPermissionLevel() === 0) {
    return [];
  }

  var contacts = data || [];
  contacts = contacts.map(function (contact) {
    return {
      personName: contact.PersonName,
      businessName: contact.BusinessName,
      phoneNumbers: createPhoneNumbers(contact.PhoneNumbers || []),
      emailAddresses: contact.EmailAddresses || [],
      urls: contact.Urls || [],
      addresses: contact.Addresses || [],
      contactString: contact.ContactString
    };
  });
  return removeDuplicates(contacts, contactComparator);
};
var createEntities_createReadItemArray = function createReadItemArray(data) {
  if (getPermissionLevel_getPermissionLevel() === 0) {
    return [];
  }

  return data || [];
};
// CONCATENATED MODULE: ./src/api/Entities.ts



var entityPermissions = {
  meetingSuggestion: 1,
  taskSuggestion: 1,
  address: 0,
  emailAddress: 1,
  url: 0,
  phoneNumber: 0,
  contact: 1,
  flightReservations: 1,
  parcelDeliveries: 1
};
var entityKeys = {
  meetingSuggestion: "meetingSuggestions",
  taskSuggestion: "taskSuggestions",
  address: "addresses",
  emailAddress: "emailAddresses",
  url: "urls",
  phoneNumber: "phoneNumbers",
  contact: "contacts",
  flightReservations: "flightReservations",
  parcelDeliveries: "parcelDeliveries"
};
var Entities_getEntities = function getEntities() {
  return createEntities(getInitialDataProp("entities"));
};
var Entities_getEntitiesByType = function getEntitiesByType(entityType) {
  var entities = createEntities(getInitialDataProp("entities"));
  checkPermissionsAndThrow(entityPermissions[entityType] !== undefined ? entityPermissions[entityType] : 1, entityType);
  var entityProperty = entityKeys[entityType];

  if (entityProperty === undefined) {
    return null;
  }

  return entities[entityProperty];
};
var Entities_getFilteredEntitiesByName = function getFilteredEntitiesByName(name) {
  return createFilteredEntities(getInitialDataProp("filteredEntities"), name);
};
var Entities_getRegExMatches = function getRegExMatches() {
  return getInitialDataProp("regExMatches");
};
var Entities_getRegExMatchesByName = function getRegExMatchesByName(name) {
  var regExMatches = getInitialDataProp("regExMatches") || {};
  return regExMatches[name];
};
var Entities_getSelectedEntities = function getSelectedEntities() {
  return createEntities(getInitialDataProp("selectedEntities"));
};
var Entities_getSelectedRegExMatches = function getSelectedRegExMatches() {
  return getInitialDataProp("selectedRegExMatches");
};
// CONCATENATED MODULE: ./src/utils/CustomJsonAttachmentsResponse.ts


function CustomJsonAttachmentsResponse(arrayOfAttachmentJsonData) {
  var customJsonResponse = [];

  if (getPermissionLevel_getPermissionLevel() === 0) {
    return [];
  }

  if (!!arrayOfAttachmentJsonData) {
    for (var i = 0; i < arrayOfAttachmentJsonData.length; i++) {
      if (!!arrayOfAttachmentJsonData[i]) {
        var newAttachment = convertAttachmentType(arrayOfAttachmentJsonData[i]);
        customJsonResponse.push(newAttachment);
      }
    }
  }

  return customJsonResponse;
}
function convertAttachmentType(attachmentDetails) {
  if (attachmentDetails.attachmentType !== null || attachmentDetails.attachmentType !== undefined) {
    switch (attachmentDetails.attachmentType) {
      case 0:
        {
          attachmentDetails.attachmentType = MailboxEnums.AttachmentType.File;
          break;
        }

      case 1:
        {
          attachmentDetails.attachmentType = MailboxEnums.AttachmentType.Item;
          break;
        }

      case 2:
        {
          attachmentDetails.attachmentType = MailboxEnums.AttachmentType.Cloud;
          break;
        }
    }
  }

  return attachmentDetails;
}
// CONCATENATED MODULE: ./src/methods/deepClone.ts
function deepClone(original) {
  return JSON.parse(JSON.stringify(original));
}
// CONCATENATED MODULE: ./src/validation/seriesTimeConstants.ts
var StartYearKey = "startYear";
var StartMonthKey = "startMonth";
var StartDayKey = "startDay";
var EndYearKey = "endYear";
var EndMonthKey = "endMonth";
var EndDayKey = "endDay";
var NoEndDateKey = "noEndDate";
var StartTimeMinKey = "startTimeMin";
var DurationMinKey = "durationMin";
// CONCATENATED MODULE: ./src/validation/recurrenceConstants.ts
var StartDateKey = "startDate";
var EndDateKey = "endDate";
var StartTimeKey = "startTime";
var EndTimeKey = "endTime";
var RecurrenceTypeKey = "recurrenceType";
var SeriesTimeKey = "seriesTime";
var SeriesTimeJsonKey = "seriesTimeJson";
var RecurrenceTimeZoneKey = "recurrenceTimeZone";
var RecurrenceTimeZoneName = "name";
var RecurrencePropertiesKey = "recurrenceProperties";
var IntervalKey = "interval";
var DaysKey = "days";
var DayOfMonthKey = "dayOfMonth";
var DayOfWeekKey = "dayOfWeek";
var WeekNumberKey = "weekNumber";
var MonthKey = "month";
var FirstDayOfWeekKey = "firstDayOfWeek";
// CONCATENATED MODULE: ./src/utils/seriesTimeUtils.ts



function prependZeroToString(number) {
  if (number < 0) {
    number = 1;
  }

  if (number < 10) {
    return "0" + number.toString();
  }

  return number.toString();
}
function throwOnInvalidDate(year, month, day) {
  if (!isValidDate(year, month, day)) {
    throw createArgumentError(SeriesTimeKey, getString("l_InvalidDate_Text"));
  }
}
function isValidDate(year, month, day) {
  if (year < 1601 || month < 1 || month > 12 || day < 1 || day > 31) {
    return false;
  }

  return true;
}
function throwOnInvalidDateString(dateString) {
  var regEx = new RegExp("^\\d{4}-(?:[0]\\d|1[0-2])-(?:[0-2]\\d|3[01])$");

  if (!regEx.test(dateString)) {
    throw createArgumentError(SeriesTimeKey, getString("l_InvalidDate_Text"));
  }
}
// CONCATENATED MODULE: ./src/api/SeriesTime.ts






var SeriesTime_SeriesTime = function () {
  function SeriesTime() {
    this.startYear = 0;
    this.startMonth = 0;
    this.startDay = 0;
    this.endYear = 0;
    this.endMonth = 0;
    this.endDay = 0;
    this.startTimeMinutes = 0;
    this.durationMinutes = 0;
  }

  SeriesTime.prototype.getDuration = function () {
    return this.durationMinutes;
  };

  SeriesTime.prototype.getEndTime = function () {
    var endTimeMinutes = this.startTimeMinutes + this.durationMinutes;
    var minutes = endTimeMinutes % 60;
    var hours = Math.floor(endTimeMinutes / 60) % 24;
    return "T" + prependZeroToString(hours) + ":" + prependZeroToString(minutes) + ":00.000";
  };

  SeriesTime.prototype.getEndDate = function () {
    if (this.endYear === 0 && this.endMonth === 0 && this.endDay === 0) {
      return null;
    }

    return this.endYear.toString() + "-" + prependZeroToString(this.endMonth) + "-" + prependZeroToString(this.endDay);
  };

  SeriesTime.prototype.getStartDate = function () {
    return this.startYear.toString() + "-" + prependZeroToString(this.startMonth) + "-" + prependZeroToString(this.startDay);
  };

  SeriesTime.prototype.getStartTime = function () {
    var minutes = this.startTimeMinutes % 60;
    var hours = Math.floor(this.startTimeMinutes / 60);
    return "T" + prependZeroToString(hours) + ":" + prependZeroToString(minutes) + ":00.000";
  };

  SeriesTime.prototype.setDuration = function (minutes) {
    if (minutes >= 0) {
      this.durationMinutes = minutes;
    } else {
      throw createArgumentError(undefined, getString("l_InvalidTime_Text"));
    }
  };

  SeriesTime.prototype.setEndDate = function (yearOrDateString, month, day) {
    if (yearOrDateString !== null && !isNullOrUndefined(month) && day !== null) {
      this.setDateHelper(false, yearOrDateString, month, day);
    } else if (yearOrDateString !== null) {
      this.setDateHelper(false, yearOrDateString);
    } else if (yearOrDateString == null) {
      this.endYear = 0;
      this.endMonth = 0;
      this.endDay = 0;
    }
  };

  SeriesTime.prototype.setStartDate = function (yearOrDateString, month, day) {
    if (yearOrDateString !== null && !isNullOrUndefined(month) && day !== null) {
      this.setDateHelper(true, yearOrDateString, month, day);
    } else if (yearOrDateString !== null) {
      this.setDateHelper(true, yearOrDateString);
    }
  };

  SeriesTime.prototype.setStartTime = function (hoursOrTimeString, minutes) {
    if (!isNullOrUndefined(hoursOrTimeString) && !isNullOrUndefined(minutes)) {
      var totalMinutes = hoursOrTimeString * 60 + minutes;

      if (totalMinutes >= 0) {
        this.startTimeMinutes = totalMinutes;
      } else {
        throw createArgumentError(undefined, getString("l_InvalidTime_Text"));
      }
    } else if (!isNullOrUndefined(hoursOrTimeString)) {
      var timeString = hoursOrTimeString;
      var newDateString = "2017-01-15" + timeString + "Z";
      var regEx = new RegExp("^T[0-2]\\d:[0-5]\\d:[0-5]\\d\\.\\d{3}$");

      if (!regEx.test(timeString)) {
        throw createArgumentError(undefined, getString("l_InvalidTime_Text"));
      }

      var dateObject = new Date(newDateString);

      if (!isNullOrUndefined(dateObject) && !isNaN(dateObject.getUTCHours()) && !isNaN(dateObject.getUTCMinutes())) {
        this.startTimeMinutes = dateObject.getUTCHours() * 60 + dateObject.getUTCMinutes();
      } else {
        throw createArgumentError(undefined, getString("l_InvalidTime_Text"));
      }
    }
  };

  SeriesTime.prototype.isValid = function () {
    if (!isValidDate(this.startYear, this.startMonth, this.startDay)) {
      return false;
    }

    if (this.endDay !== 0 && this.endMonth !== 0 && this.endYear !== 0) {
      if (!isValidDate(this.endYear, this.endMonth, this.endDay)) {
        return false;
      }
    }

    if (this.startTimeMinutes < 0 || this.durationMinutes <= 0) {
      return false;
    }

    return true;
  };

  SeriesTime.prototype.exportToSeriesTimeJson = function () {
    var result = {};
    result[StartYearKey] = this.startYear;
    result[StartMonthKey] = this.startMonth;
    result[StartDayKey] = this.startDay;

    if (this.endYear === 0 && this.endMonth === 0 && this.endDay === 0) {
      result[NoEndDateKey] = true;
    } else {
      result[EndYearKey] = this.endYear;
      result[EndMonthKey] = this.endMonth;
      result[EndDayKey] = this.endDay;
    }

    result[StartTimeMinKey] = this.startTimeMinutes;

    if (this.durationMinutes > 0) {
      result[DurationMinKey] = this.durationMinutes;
    }

    return result;
  };

  SeriesTime.prototype.importFromSeriesTimeJsonObject = function (jsonObject) {
    this.startYear = jsonObject[StartYearKey];
    this.startMonth = jsonObject[StartMonthKey];
    this.startDay = jsonObject[StartDayKey];

    if (jsonObject[NoEndDateKey] != null && typeof jsonObject[NoEndDateKey] === "boolean") {
      this.endYear = 0;
      this.endMonth = 0;
      this.endDay = 0;
    } else {
      this.endYear = jsonObject[EndYearKey];
      this.endMonth = jsonObject[EndMonthKey];
      this.endDay = jsonObject[EndDayKey];
    }

    this.startTimeMinutes = jsonObject[StartTimeMinKey];
    this.durationMinutes = jsonObject[DurationMinKey];
  };

  SeriesTime.prototype.setDateHelper = function (isStart, yearOrDateString, month, day) {
    var yearCalculated = 0;
    var monthCalculated = 0;
    var dayCalculated = 0;

    if (yearOrDateString !== null && !isNullOrUndefined(month) && day !== null) {
      throwOnInvalidDate(yearOrDateString, month + 1, day);
      yearCalculated = yearOrDateString;
      monthCalculated = month + 1;
      dayCalculated = day;
    } else if (yearOrDateString !== null) {
      var dateString = yearOrDateString;
      throwOnInvalidDateString(dateString);
      var dateObject = new Date(dateString);

      if (dateObject !== null && !isNaN(dateObject.getUTCFullYear()) && !isNaN(dateObject.getUTCMonth()) && !isNaN(dateObject.getUTCDate())) {
        throwOnInvalidDate(dateObject.getUTCFullYear(), dateObject.getUTCMonth() + 1, dateObject.getUTCDate());
        yearCalculated = dateObject.getUTCFullYear();
        monthCalculated = dateObject.getUTCMonth() + 1;
        dayCalculated = dateObject.getUTCDate();
      }
    }

    if (yearCalculated !== 0 && monthCalculated !== 0 && dayCalculated !== 0) {
      if (isStart) {
        this.startYear = yearCalculated;
        this.startMonth = monthCalculated;
        this.startDay = dayCalculated;
      } else {
        this.endYear = yearCalculated;
        this.endMonth = monthCalculated;
        this.endDay = dayCalculated;
      }
    }
  };

  SeriesTime.prototype.isEndAfterStart = function () {
    if (this.endYear === 0 && this.endMonth === 0 && this.endDay === 0) {
      return true;
    }

    var startDateTime = new Date();
    startDateTime.setFullYear(this.startYear);
    startDateTime.setMonth(this.startMonth - 1);
    startDateTime.setDate(this.startDay);
    var endDateTime = new Date();
    endDateTime.setFullYear(this.endYear);
    endDateTime.setMonth(this.endMonth - 1);
    endDateTime.setDate(this.endDay);
    return endDateTime >= startDateTime;
  };

  return SeriesTime;
}();


// CONCATENATED MODULE: ./src/utils/recurrenceUtils.ts



function copyRecurrenceObjectConvertSeriesTimeJson(recurrenceOriginal) {
  if (isNullOrUndefined(recurrenceOriginal) || isNullOrUndefined(recurrenceOriginal.seriesTimeJson)) {
    return recurrenceOriginal;
  }

  var recurrenceCopy = {
    recurrenceType: "",
    recurrenceProperties: null,
    recurrenceTimeZone: null
  };
  var newSeriesTime = new SeriesTime_SeriesTime();

  if (!isNullOrUndefined(recurrenceOriginal.recurrenceProperties)) {
    recurrenceCopy.recurrenceProperties = deepClone(recurrenceOriginal.recurrenceProperties);
  }

  recurrenceCopy.recurrenceType = recurrenceOriginal.recurrenceType;

  if (!isNullOrUndefined(recurrenceOriginal.recurrenceTimeZone)) {
    recurrenceCopy.recurrenceTimeZone = deepClone(recurrenceOriginal.recurrenceTimeZone);
  }

  newSeriesTime.importFromSeriesTimeJsonObject(recurrenceOriginal.seriesTimeJson);
  recurrenceCopy.seriesTime = newSeriesTime;
  return recurrenceCopy;
}
// CONCATENATED MODULE: ./src/api/getMessageRead.ts
















function getMessageRead() {
  var sender = getInitialDataProp("sender");
  var from = getInitialDataProp("from");
  var dateTimeCreated = getInitialDataProp("dateTimeCreated");
  var dateTimeModified = getInitialDataProp("dateTimeModified");
  var end = getInitialDataProp("end");
  var start = getInitialDataProp("start");
  var messageRead = objectDefine({}, {
    attachments: CustomJsonAttachmentsResponse(getInitialDataProp("attachments")),
    bcc: (getInitialDataProp("bcc") || []).map(createEmailAddressDetails),
    body: getBodySurface(false),
    categories: getCategoriesSurface(),
    cc: (getInitialDataProp("cc") || []).map(createEmailAddressDetails),
    conversationId: getInitialDataProp("conversationId"),
    dateTimeCreated: dateTimeCreated ? new Date(dateTimeCreated) : undefined,
    dateTimeModified: dateTimeModified ? new Date(dateTimeModified) : undefined,
    end: end ? new Date(end) : undefined,
    from: from ? createEmailAddressDetails(from) : undefined,
    getAllInternetHeadersAsync: getAllInternetHeaders,
    internetMessageId: getInitialDataProp("internetMessageId"),
    itemClass: getInitialDataProp("itemClass"),
    itemId: getInitialDataProp("id"),
    itemType: "message",
    location: getInitialDataProp("location"),
    move: moveToFolder,
    normalizedSubject: getInitialDataProp("normalizedSubject"),
    notificationMessages: getNotificationMessageSurface(),
    recurrence: copyRecurrenceObjectConvertSeriesTimeJson(getInitialDataProp("recurrence")),
    seriesId: getInitialDataProp("seriesId"),
    sender: sender ? createEmailAddressDetails(sender) : undefined,
    start: start ? new Date(start) : undefined,
    subject: getInitialDataProp("subject"),
    to: (getInitialDataProp("to") || []).map(createEmailAddressDetails),
    displayReplyForm: displayReplyForm,
    displayReplyFormAsync: displayReplyFormAsync,
    displayReplyAllForm: displayReplyAllForm,
    displayReplyAllFormAsync: displayReplyAllFormAsync,
    getAttachmentContentAsync: getAttachmentContent,
    getEntities: Entities_getEntities,
    getEntitiesByType: Entities_getEntitiesByType,
    getFilteredEntitiesByName: Entities_getFilteredEntitiesByName,
    getInitializationContextAsync: getInitializationContext,
    getRegExMatches: Entities_getRegExMatches,
    getRegExMatchesByName: Entities_getRegExMatchesByName,
    getSelectedEntities: Entities_getSelectedEntities,
    getSelectedRegExMatches: Entities_getSelectedRegExMatches,
    loadCustomPropertiesAsync: loadCustomProperties,
    delayDeliveryTime: getDelayDeliverySurface(false),
    isAllDayEvent: getInitialDataProp("isAllDayEvent"),
    sensitivity: getInitialDataProp("sensitivity")
  });
  return messageRead;
}
// CONCATENATED MODULE: ./src/validation/validateAttachments.ts




function validateAddFileAttachmentApis(attachmentName) {
  if (isNullOrUndefined(attachmentName) || attachmentName === "" || !(typeof attachmentName === "string")) {
    throw createArgumentError("attachmentName");
  }

  throwOnOutOfRange(attachmentName.length, 0, MaxAttachmentNameLength, "attachmentName");
}
// CONCATENATED MODULE: ./src/validation/attachmentsConstants.ts
var AddItemAttachmentClientEndPointTimeoutInMilliseconds = 600000;
var MaxBase64AttachmentSize = 27892122;
// CONCATENATED MODULE: ./src/methods/addFileAttachment.ts








function addFileAttachment(uri, attachmentName) {
  var args = [];

  for (var _i = 2; _i < arguments.length; _i++) {
    args[_i - 2] = arguments[_i];
  }

  checkPermissionsAndThrow(2, "item.addFileAttachmentAsync");
  var commonParameters = parseCommonArgs(args, false, false);
  var isInline = false;

  if (!!commonParameters.options) {
    isInline = !!commonParameters.options.isInline;
  }

  var name = attachmentName;
  var parameters = {
    uri: uri,
    name: name,
    isInline: isInline,
    __timeout__: AddItemAttachmentClientEndPointTimeoutInMilliseconds
  };
  addFileAttachment_validateParameters(parameters);
  standardInvokeHostMethod(16, commonParameters.asyncContext, commonParameters.callback, parameters, undefined);
}

function addFileAttachment_validateParameters(parameters) {
  validateStringParam("uri", parameters.uri);
  throwOnOutOfRange(parameters.uri.length, 0, MaxUrlLength, "uri");
  validateAddFileAttachmentApis(parameters.name);
}
// CONCATENATED MODULE: ./src/methods/addBase64FileAttachment.ts







function addBase64FileAttachment(base64String, name) {
  var args = [];

  for (var _i = 2; _i < arguments.length; _i++) {
    args[_i - 2] = arguments[_i];
  }

  checkPermissionsAndThrow(2, "item.addBase64FileAttachmentAsync");
  var commonParameters = parseCommonArgs(args, false, false);
  var isInline = false;

  if (!!commonParameters.options) {
    isInline = !!commonParameters.options.isInline;
  }

  var parameters = {
    base64String: base64String,
    name: name,
    isInline: isInline,
    __timeout__: AddItemAttachmentClientEndPointTimeoutInMilliseconds
  };
  addBase64FileAttachment_validateParameters(parameters);
  standardInvokeHostMethod(148, commonParameters.asyncContext, commonParameters.callback, parameters, undefined);
}

function addBase64FileAttachment_validateParameters(parameters) {
  validateStringParam("base64Encoded", parameters.base64String);
  throwOnOutOfRange(parameters.base64String.length, 0, MaxBase64AttachmentSize, "base64File");
  validateAddFileAttachmentApis(parameters.name);
}
// CONCATENATED MODULE: ./src/methods/addItemAttachment.ts








function addItemAttachment(itemId, name) {
  var args = [];

  for (var _i = 2; _i < arguments.length; _i++) {
    args[_i - 2] = arguments[_i];
  }

  checkPermissionsAndThrow(2, "item.addItemAttachmentAsync");
  var commonParameters = parseCommonArgs(args, false, false);
  var parameters = {
    itemId: itemId,
    name: name
  };
  addItemAttachment_validateParameters(parameters);
  standardInvokeHostMethod(19, commonParameters.asyncContext, commonParameters.callback, {
    itemId: getItemIdBasedOnHost(parameters.itemId),
    name: parameters.name,
    __timeout__: AddItemAttachmentClientEndPointTimeoutInMilliseconds
  }, undefined);
}

function addItemAttachment_validateParameters(parameters) {
  validateStringParam("itemId", parameters.itemId);
  validateStringParam("attachmentName", parameters.name);
  throwOnOutOfRange(parameters.itemId.length, 0, MaxItemIdLength, "itemId");
  throwOnOutOfRange(parameters.name.length, 0, MaxAttachmentNameLength, "attachmentName");
}
// CONCATENATED MODULE: ./src/methods/close.ts

function close_close() {
  standardInvokeHostMethod(41, undefined, undefined, undefined, undefined);
}
// CONCATENATED MODULE: ./src/methods/getAttachments.ts




function getAttachments_getAttachments() {
  var args = [];

  for (var _i = 0; _i < arguments.length; _i++) {
    args[_i] = arguments[_i];
  }

  checkPermissionsAndThrow(1, "item.getAttachmentsAsync");
  var commonParameters = parseCommonArgs(args, true, false);
  standardInvokeHostMethod(149, commonParameters.asyncContext, commonParameters.callback, undefined, CustomJsonAttachmentsResponse);
}
// CONCATENATED MODULE: ./src/methods/getSelectedData.ts





function getSelectedData(coercionType) {
  var args = [];

  for (var _i = 1; _i < arguments.length; _i++) {
    args[_i - 1] = arguments[_i];
  }

  checkPermissionsAndThrow(2, "item.getSelectedDataAsync");
  var commonParameters = parseCommonArgs(args, true, false);
  var parameters = {
    coercionType: getCoercionTypeFromString(coercionType)
  };

  if (parameters.coercionType === undefined) {
    throw createArgumentError("coercionType");
  }

  standardInvokeHostMethod(28, commonParameters.asyncContext, commonParameters.callback, parameters, undefined);
}
// CONCATENATED MODULE: ./src/methods/removeAttachment.ts






function removeAttachment(attachmentId) {
  var args = [];

  for (var _i = 1; _i < arguments.length; _i++) {
    args[_i - 1] = arguments[_i];
  }

  checkPermissionsAndThrow(2, "item.removeAttachmentAsync");
  var commonParameters = parseCommonArgs(args, false, false);
  var parameters = {
    attachmentIndex: attachmentId
  };
  removeAttachment_validateParameters(parameters);
  standardInvokeHostMethod(20, commonParameters.asyncContext, commonParameters.callback, parameters, undefined);
}

function removeAttachment_validateParameters(parameters) {
  validateStringParam("attachmentId", parameters.attachmentIndex);
  throwOnOutOfRange(parameters.attachmentIndex.length, 0, MaxRemoveIdLength, "attachmentId");
}
// CONCATENATED MODULE: ./src/methods/save.ts



function save() {
  var args = [];

  for (var _i = 0; _i < arguments.length; _i++) {
    args[_i] = arguments[_i];
  }

  checkPermissionsAndThrow(2, "item.saveAsync");
  var commonParameters = parseCommonArgs(args, false, false);
  standardInvokeHostMethod(32, commonParameters.asyncContext, commonParameters.callback, undefined, undefined);
}
// CONCATENATED MODULE: ./src/validation/validateRecipientParameters.ts




function validateRecipientParameters(parameters) {
  if (Array.isArray(parameters.recipientArray)) {
    if (parameters.recipientArray.length > recipientsLimit) {
      throw createArgumentOutOfRange("recipients", parameters.recipientArray.length);
    }

    var validatedRecipients = parameters.recipientArray.map(function (recipient) {
      if (isNullOrUndefined(recipient)) {
        throw createArgumentError("recipients");
      }

      if (typeof recipient === "string") {
        throwOnInvalidDisplayNameOrEmail(recipient, recipient);
        return createEmailAddressForHost(recipient, recipient);
      } else if (typeof recipient === "object") {
        throwOnInvalidDisplayNameOrEmail(recipient.displayName, recipient.emailAddress);
        return createEmailAddressForHost(recipient.displayName, recipient.emailAddress);
      } else {
        throw createArgumentError("recipients");
      }
    });
    parameters.recipientArray = validatedRecipients;
  } else {
    throw createArgumentError("recipients");
  }
}

function throwOnInvalidDisplayNameOrEmail(displayName, email) {
  if (!displayName && !email) {
    throw createArgumentError("recipients");
  } else if (typeof displayName === "string" && displayName.length > displayNameLengthLimit) {
    throw createArgumentOutOfRange("recipients", displayName.length, getString("l_DisplayNameTooLong_Text"));
  } else if (typeof email === "string" && email.length > maxSmtpLength) {
    throw createArgumentOutOfRange("recipients", email.length, getString("l_EmailAddressTooLong_Text"));
  } else if (typeof displayName !== "string" && typeof email !== "string") {
    throw createArgumentError("recipients");
  }
}

function createEmailAddressForHost(displayName, email) {
  return {
    address: email,
    name: displayName
  };
}
// CONCATENATED MODULE: ./src/methods/addRecipients.ts





function addRecipients(namespace) {
  return function (recipientArray) {
    var args = [];

    for (var _i = 1; _i < arguments.length; _i++) {
      args[_i - 1] = arguments[_i];
    }

    checkPermissionsAndThrow(2, namespace + ".addAsync");
    var commonParameters = parseCommonArgs(args, false, false);
    var parameters = {
      recipientField: RecipientFields[namespace],
      recipientArray: recipientArray
    };
    validateRecipientParameters(parameters);
    standardInvokeHostMethod(22, commonParameters.asyncContext, commonParameters.callback, parameters, undefined);
  };
}
// CONCATENATED MODULE: ./src/methods/getRecipients.ts





function getRecipients(namespace) {
  return function () {
    var args = [];

    for (var _i = 0; _i < arguments.length; _i++) {
      args[_i] = arguments[_i];
    }

    checkPermissionsAndThrow(1, namespace + ".getAsync");
    var commonParameters = parseCommonArgs(args, true, false);
    standardInvokeHostMethod(15, commonParameters.asyncContext, commonParameters.callback, {
      recipientField: RecipientFields[namespace]
    }, getRecipients_format);
  };
}

function getRecipients_format(rawInput) {
  if (rawInput === null || rawInput === undefined) {
    return [];
  }

  return rawInput.map(function (input) {
    return createEmailAddressDetails(input);
  });
}
// CONCATENATED MODULE: ./src/methods/setRecipients.ts





function setRecipients(namespace) {
  return function (recipientArray) {
    var args = [];

    for (var _i = 1; _i < arguments.length; _i++) {
      args[_i - 1] = arguments[_i];
    }

    checkPermissionsAndThrow(2, namespace + ".setAsync");
    var commonParameters = parseCommonArgs(args, false, false);
    var parameters = {
      recipientField: RecipientFields[namespace],
      recipientArray: recipientArray
    };
    validateRecipientParameters(parameters);
    standardInvokeHostMethod(21, commonParameters.asyncContext, commonParameters.callback, parameters, undefined);
  };
}
// CONCATENATED MODULE: ./src/api/getRecipientsSurface.ts




function getRecipientsSurface(namespace) {
  return objectDefine({}, {
    addAsync: addRecipients(namespace),
    getAsync: getRecipients(namespace),
    setAsync: setRecipients(namespace)
  });
}
// CONCATENATED MODULE: ./src/methods/getFrom.ts





function getFrom(namespace) {
  return function () {
    var args = [];

    for (var _i = 0; _i < arguments.length; _i++) {
      args[_i] = arguments[_i];
    }

    checkPermissionsAndThrow(1, namespace + ".getAsync");
    var commonParameters = parseCommonArgs(args, true, false);
    standardInvokeHostMethod(107, commonParameters.asyncContext, commonParameters.callback, undefined, getFrom_format);
  };
}

function getFrom_format(rawInput) {
  return isNullOrUndefined(rawInput) ? null : createEmailAddressDetails(rawInput);
}
// CONCATENATED MODULE: ./src/api/getFromSurface.ts


function getFromSurface(namespace) {
  return objectDefine({}, {
    getAsync: getFrom(namespace)
  });
}
// CONCATENATED MODULE: ./src/validation/validateInternetHeaders.ts



function validateInternetHeaderArray(internetHeaderArray) {
  if (isNullOrUndefined(internetHeaderArray)) {
    throw createArgumentError("internetHeaders");
  }

  if (!Array.isArray(internetHeaderArray)) {
    throw createArgumentTypeError("internetHeaders", typeof internetHeaderArray, "Array");
  }

  if (internetHeaderArray.length === 0) {
    throw createArgumentError("internetHeaders");
  }

  for (var _i = 0, internetHeaderArray_1 = internetHeaderArray; _i < internetHeaderArray_1.length; _i++) {
    var internetHeader = internetHeaderArray_1[_i];
    validateStringParam("internetHeaders", internetHeader);
  }
}
// CONCATENATED MODULE: ./src/methods/removeInternetHeaders.ts




function removeInternetHeaders(internetHeaderKeys) {
  var args = [];

  for (var _i = 1; _i < arguments.length; _i++) {
    args[_i - 1] = arguments[_i];
  }

  checkPermissionsAndThrow(2, "internetHeaders.removeAsync");
  var commonParameters = parseCommonArgs(args, false, false);
  var parameters = {
    internetHeaderKeys: internetHeaderKeys
  };
  removeInternetHeaders_validateParameters(parameters);
  standardInvokeHostMethod(153, commonParameters.asyncContext, commonParameters.callback, parameters, undefined);
}

function removeInternetHeaders_validateParameters(parameters) {
  validateInternetHeaderArray(parameters.internetHeaderKeys);
}
// CONCATENATED MODULE: ./src/methods/getInternetHeaders.ts




function getInternetHeaders(internetHeaderKeys) {
  var args = [];

  for (var _i = 1; _i < arguments.length; _i++) {
    args[_i - 1] = arguments[_i];
  }

  checkPermissionsAndThrow(1, "internetHeaders.getAsync");
  var commonParameters = parseCommonArgs(args, true, false);
  var parameters = {
    internetHeaderKeys: internetHeaderKeys
  };
  getInternetHeaders_validateParameters(parameters);
  standardInvokeHostMethod(151, commonParameters.asyncContext, commonParameters.callback, parameters, undefined);
}

function getInternetHeaders_validateParameters(parameters) {
  validateInternetHeaderArray(parameters.internetHeaderKeys);
}
// CONCATENATED MODULE: ./src/methods/setInternetHeaders.ts







var InternetHeadersLimit = 998;
function setInternetHeaders(internetHeaderNameValuePairs) {
  var args = [];

  for (var _i = 1; _i < arguments.length; _i++) {
    args[_i - 1] = arguments[_i];
  }

  checkPermissionsAndThrow(2, "internetHeaders.setAsync");
  var commonParameters = parseCommonArgs(args, false, false);
  var parameters = {
    internetHeaderNameValuePairs: internetHeaderNameValuePairs
  };
  setInternetHeaders_validateParameters(parameters);
  standardInvokeHostMethod(152, commonParameters.asyncContext, commonParameters.callback, parameters, undefined);
}

function setInternetHeaders_validateParameters(parameters) {
  if (isNullOrUndefined(parameters.internetHeaderNameValuePairs)) {
    throw createNullArgumentError("internetHeaders");
  }

  var keys = Object.keys(parameters.internetHeaderNameValuePairs);

  if (keys.length === 0) {
    throw createArgumentError("internetHeaders");
  }

  for (var _i = 0, keys_1 = keys; _i < keys_1.length; _i++) {
    var key = keys_1[_i];
    var value = parameters.internetHeaderNameValuePairs[key];
    validateStringParam("internetHeaders", key);

    if (!(typeof value === "string")) {
      throw createArgumentTypeError("internetHeaders", typeof value, "string");
    }

    throwOnOutOfRange(key.length + value.length, 0, InternetHeadersLimit, key);
  }
}
// CONCATENATED MODULE: ./src/api/getInternetHeadersSurface.ts




function getInternetHeadersSurface(isCompose) {
  var internetHeaders = objectDefine({}, {
    getAsync: getInternetHeaders
  });

  if (isCompose) {
    objectDefine(internetHeaders, {
      removeAsync: removeInternetHeaders,
      setAsync: setInternetHeaders
    });
  }

  return internetHeaders;
}
// CONCATENATED MODULE: ./src/methods/getSubject.ts



function getSubject() {
  var args = [];

  for (var _i = 0; _i < arguments.length; _i++) {
    args[_i] = arguments[_i];
  }

  checkPermissionsAndThrow(1, "subject.getAsync");
  var commonParameters = parseCommonArgs(args, true, false);
  standardInvokeHostMethod(18, commonParameters.asyncContext, commonParameters.callback, undefined, undefined);
}
// CONCATENATED MODULE: ./src/methods/setSubject.ts





var MaximumSubjectLength = 255;
function setSubject(subject) {
  var args = [];

  for (var _i = 1; _i < arguments.length; _i++) {
    args[_i - 1] = arguments[_i];
  }

  checkPermissionsAndThrow(2, "subject.setAsync");
  var commonParameters = parseCommonArgs(args, false, false);
  var parameters = {
    subject: subject
  };
  setSubject_validateParameters(parameters);
  standardInvokeHostMethod(17, commonParameters.asyncContext, commonParameters.callback, parameters, undefined);
}

function setSubject_validateParameters(parameters) {
  if (!(typeof parameters.subject === "string")) {
    throw createArgumentTypeError("subject", typeof parameters.subject, "string");
  }

  throwOnOutOfRange(parameters.subject.length, 0, MaximumSubjectLength, "subject");
}
// CONCATENATED MODULE: ./src/api/getSubjectSurface.ts



function getSubjectSurface() {
  return objectDefine({}, {
    getAsync: getSubject,
    setAsync: setSubject
  });
}
// CONCATENATED MODULE: ./src/methods/getItemId.ts



function getItemId() {
  var args = [];

  for (var _i = 0; _i < arguments.length; _i++) {
    args[_i] = arguments[_i];
  }

  checkPermissionsAndThrow(1, "item.getItemIdAsync");
  var commonParameters = parseCommonArgs(args, true, false);
  standardInvokeHostMethod(164, commonParameters.asyncContext, commonParameters.callback, undefined, undefined);
}
// CONCATENATED MODULE: ./src/methods/getComposeType.ts




function getComposeType() {
  var args = [];

  for (var _i = 0; _i < arguments.length; _i++) {
    args[_i] = arguments[_i];
  }

  checkPermissionsAndThrow(1, "item.getComposeTypeAsync");
  var commonParameters = parseCommonArgs(args, true, false);
  checkFeatureEnabledAndThrow(Features.signature, "getComposeTypeAsync");
  standardInvokeHostMethod(174, commonParameters.asyncContext, commonParameters.callback, undefined, undefined);
}
// CONCATENATED MODULE: ./src/methods/isClientSignatureEnabled.ts




function isClientSignatureEnabled() {
  var args = [];

  for (var _i = 0; _i < arguments.length; _i++) {
    args[_i] = arguments[_i];
  }

  checkPermissionsAndThrow(1, "isClientSignatureEnabledAsync");
  var commonParameters = parseCommonArgs(args, true, false);
  checkFeatureEnabledAndThrow(Features.signature, "isClientSignatureEnabledAsync");
  standardInvokeHostMethod(175, commonParameters.asyncContext, commonParameters.callback, undefined, undefined);
}
// CONCATENATED MODULE: ./src/methods/disableClientSignature.ts




function disableClientSignature() {
  var args = [];

  for (var _i = 0; _i < arguments.length; _i++) {
    args[_i] = arguments[_i];
  }

  checkPermissionsAndThrow(2, "disableClientSignatureAsync");
  var commonParameters = parseCommonArgs(args, true, false);
  checkFeatureEnabledAndThrow(Features.signature, "disableClientSignatureAsync");
  standardInvokeHostMethod(176, commonParameters.asyncContext, commonParameters.callback, undefined, undefined);
}
// CONCATENATED MODULE: ./src/methods/getSessionData.ts





function getSessionData(name) {
  var args = [];

  for (var _i = 1; _i < arguments.length; _i++) {
    args[_i - 1] = arguments[_i];
  }

  checkPermissionsAndThrow(2, "sessionData.getAsync");
  var commonParameters = parseCommonArgs(args, true, false);
  checkFeatureEnabledAndThrow(Features.sessionData, "sessionData.getAsync");
  var parameters = {
    name: name
  };
  getSessionData_validateParameters(parameters);
  standardInvokeHostMethod(186, commonParameters.asyncContext, commonParameters.callback, parameters, undefined);
}

function getSessionData_validateParameters(parameters) {
  validateStringParam("name", parameters.name);
}
// CONCATENATED MODULE: ./src/methods/setSessionData.ts





function setSessionData(name, value) {
  var args = [];

  for (var _i = 2; _i < arguments.length; _i++) {
    args[_i - 2] = arguments[_i];
  }

  checkPermissionsAndThrow(2, "sessionData.setAsync");
  var commonParameters = parseCommonArgs(args, false, false);
  var parameters = {
    name: name,
    value: value
  };
  checkFeatureEnabledAndThrow(Features.sessionData, "sessionData.setAsync");
  setSessionData_validateParameters(parameters);
  standardInvokeHostMethod(185, commonParameters.asyncContext, commonParameters.callback, parameters, undefined);
}

function setSessionData_validateParameters(parameters) {
  validateStringParam("name", parameters.name);
  validateStringParamWithEmptyAllowed("value", parameters.value);
}
// CONCATENATED MODULE: ./src/methods/getAllSessionData.ts




function getAllSessionData() {
  var args = [];

  for (var _i = 0; _i < arguments.length; _i++) {
    args[_i] = arguments[_i];
  }

  checkPermissionsAndThrow(2, "sessionData.getAllAsync");
  var commonParameters = parseCommonArgs(args, true, false);
  checkFeatureEnabledAndThrow(Features.sessionData, "sessionData.getAllAsync");
  standardInvokeHostMethod(187, commonParameters.asyncContext, commonParameters.callback, undefined, undefined);
}
// CONCATENATED MODULE: ./src/methods/clearSessionData.ts




function clearSessionData() {
  var args = [];

  for (var _i = 0; _i < arguments.length; _i++) {
    args[_i] = arguments[_i];
  }

  checkPermissionsAndThrow(2, "sessionData.clearAsync");
  var commonParameters = parseCommonArgs(args, false, false);
  checkFeatureEnabledAndThrow(Features.sessionData, "sessionData.clearAsync");
  standardInvokeHostMethod(188, commonParameters.asyncContext, commonParameters.callback, undefined, undefined);
}
// CONCATENATED MODULE: ./src/methods/removeSessionData.ts





function removeSessionData(name) {
  var args = [];

  for (var _i = 1; _i < arguments.length; _i++) {
    args[_i - 1] = arguments[_i];
  }

  checkPermissionsAndThrow(2, "sessionData.removeAsync");
  var commonParameters = parseCommonArgs(args, false, false);
  var parameters = {
    name: name
  };
  checkFeatureEnabledAndThrow(Features.sessionData, "sessionData.removeAsync");
  removeSessionData_validateParameters(parameters);
  standardInvokeHostMethod(189, commonParameters.asyncContext, commonParameters.callback, parameters, undefined);
}

function removeSessionData_validateParameters(parameters) {
  validateStringParam("name", parameters.name);
}
// CONCATENATED MODULE: ./src/api/getSessionDataSurface.ts






function getSessionDataSurface() {
  return objectDefine({}, {
    getAsync: getSessionData,
    setAsync: setSessionData,
    getAllAsync: getAllSessionData,
    clearAsync: clearSessionData,
    removeAsync: removeSessionData
  });
}
// CONCATENATED MODULE: ./src/api/getMessageCompose.ts



























function getMessageCompose() {
  var messageCompose = objectDefine({}, {
    bcc: getRecipientsSurface("bcc"),
    body: getBodySurface(true),
    categories: getCategoriesSurface(),
    cc: getRecipientsSurface("cc"),
    conversationId: getInitialDataProp("conversationId"),
    from: getFromSurface("from"),
    internetHeaders: getInternetHeadersSurface(true),
    itemType: "message",
    notificationMessages: getNotificationMessageSurface(),
    seriesId: getInitialDataProp("seriesId"),
    subject: getSubjectSurface(),
    to: getRecipientsSurface("to"),
    addFileAttachmentAsync: addFileAttachment,
    addFileAttachmentFromBase64Async: addBase64FileAttachment,
    addItemAttachmentAsync: addItemAttachment,
    close: close_close,
    getAttachmentsAsync: getAttachments_getAttachments,
    getAttachmentContentAsync: getAttachmentContent,
    getInitializationContextAsync: getInitializationContext,
    getItemIdAsync: getItemId,
    getSelectedDataAsync: getSelectedData,
    loadCustomPropertiesAsync: loadCustomProperties,
    removeAttachmentAsync: removeAttachment,
    saveAsync: save,
    setSelectedDataAsync: setSelectedData(29),
    delayDeliveryTime: getDelayDeliverySurface(true),
    getComposeTypeAsync: getComposeType,
    isClientSignatureEnabledAsync: isClientSignatureEnabled,
    disableClientSignatureAsync: disableClientSignature,
    sessionData: getSessionDataSurface()
  });
  return messageCompose;
}
// CONCATENATED MODULE: ./src/validation/validateEnhancedLocation.ts




function validateLocationIdentifiers(locationIdentifiers) {
  if (isNullOrUndefined(locationIdentifiers)) {
    throw createNullArgumentError("locationIdentifier");
  }

  if (!Array.isArray(locationIdentifiers)) {
    throw createArgumentTypeError("locationIdentifier", typeof locationIdentifiers, "Array");
  }

  if (locationIdentifiers.length === 0) {
    throw createArgumentError("locationIdentifier");
  }

  for (var _i = 0, locationIdentifiers_1 = locationIdentifiers; _i < locationIdentifiers_1.length; _i++) {
    var locationIdentifier = locationIdentifiers_1[_i];
    validateLocationIdentifier(locationIdentifier);
  }
}

function validateLocationIdentifier(locationIdentifier) {
  if (isNullOrUndefined(locationIdentifier) || isNullOrUndefined(locationIdentifier.id) || isNullOrUndefined(locationIdentifier.type)) {
    throw createNullArgumentError("locationIdentifier");
  }

  if (locationIdentifier.type === MailboxEnums.LocationType.Room || locationIdentifier.type === MailboxEnums.LocationType.Custom) {
    validateIdParameter(locationIdentifier.id, locationIdentifier.type);
  } else {
    throw createArgumentError("type");
  }
}

function validateIdParameter(id, type) {
  if (id === "") {
    throw createArgumentError("id");
  }

  if (type === MailboxEnums.LocationType.Room) {
    if (id.length > maxSmtpLength) {
      throw createArgumentError("id");
    }
  }
}
// CONCATENATED MODULE: ./src/methods/addEnhancedLocations.ts




function addEnhancedLocations(enhancedLocations) {
  var args = [];

  for (var _i = 1; _i < arguments.length; _i++) {
    args[_i - 1] = arguments[_i];
  }

  checkPermissionsAndThrow(2, "enhancedLocations.addAsync");
  var commonParameters = parseCommonArgs(args, false, false);
  var parameters = {
    enhancedLocations: enhancedLocations
  };
  addEnhancedLocations_validateParameters(parameters);
  standardInvokeHostMethod(155, commonParameters.asyncContext, commonParameters.callback, parameters, undefined);
}

function addEnhancedLocations_validateParameters(parameters) {
  validateLocationIdentifiers(parameters.enhancedLocations);
}
// CONCATENATED MODULE: ./src/methods/getEnhancedLocations.ts



function getEnhancedLocations() {
  var args = [];

  for (var _i = 0; _i < arguments.length; _i++) {
    args[_i] = arguments[_i];
  }

  checkPermissionsAndThrow(1, "enhancedLocations.getAsync");
  var commonParameters = parseCommonArgs(args, true, false);
  standardInvokeHostMethod(154, commonParameters.asyncContext, commonParameters.callback, undefined, undefined);
}
// CONCATENATED MODULE: ./src/methods/removeEnhancedLocations.ts




function removeEnhancedLocations(enhancedLocations) {
  var args = [];

  for (var _i = 1; _i < arguments.length; _i++) {
    args[_i - 1] = arguments[_i];
  }

  checkPermissionsAndThrow(2, "enhancedLocations.removeAsync");
  var commonParameters = parseCommonArgs(args, false, false);
  var parameters = {
    enhancedLocations: enhancedLocations
  };
  removeEnhancedLocations_validateParameters(parameters);
  standardInvokeHostMethod(156, commonParameters.asyncContext, commonParameters.callback, parameters, undefined);
}

function removeEnhancedLocations_validateParameters(parameters) {
  validateLocationIdentifiers(parameters.enhancedLocations);
}
// CONCATENATED MODULE: ./src/api/getEnhancedLocationSurface.ts




function getEnhancedLocationsSurface(isCompose) {
  var enhancedLocations = objectDefine({}, {
    getAsync: getEnhancedLocations
  });

  if (isCompose) {
    objectDefine(enhancedLocations, {
      addAsync: addEnhancedLocations,
      removeAsync: removeEnhancedLocations
    });
  }

  return enhancedLocations;
}
// CONCATENATED MODULE: ./src/api/getAppointmentRead.ts














function getAppointmentRead() {
  var organizer = getInitialDataProp("organizer");
  var dateTimeCreated = getInitialDataProp("dateTimeCreated");
  var dateTimeModified = getInitialDataProp("dateTimeModified");
  var end = getInitialDataProp("end");
  var start = getInitialDataProp("start");
  var appointmentRead = objectDefine({}, {
    attachments: CustomJsonAttachmentsResponse(getInitialDataProp("attachments")),
    body: getBodySurface(false),
    categories: getCategoriesSurface(),
    dateTimeCreated: dateTimeCreated ? new Date(dateTimeCreated) : undefined,
    dateTimeModified: dateTimeModified ? new Date(dateTimeModified) : undefined,
    end: end ? new Date(end) : undefined,
    enhancedLocation: getEnhancedLocationsSurface(false),
    itemClass: getInitialDataProp("itemClass"),
    itemId: getInitialDataProp("id"),
    itemType: "appointment",
    location: getInitialDataProp("location"),
    normalizedSubject: getInitialDataProp("normalizedSubject"),
    notificationMessages: getNotificationMessageSurface(),
    optionalAttendees: (getInitialDataProp("cc") || []).map(createEmailAddressDetails),
    organizer: organizer ? createEmailAddressDetails(organizer) : undefined,
    recurrence: copyRecurrenceObjectConvertSeriesTimeJson(getInitialDataProp("recurrence")),
    requiredAttendees: (getInitialDataProp("to") || []).map(createEmailAddressDetails),
    start: start ? new Date(start) : undefined,
    seriesId: getInitialDataProp("seriesId"),
    subject: getInitialDataProp("subject"),
    displayReplyForm: displayReplyForm,
    displayReplyFormAsync: displayReplyFormAsync,
    displayReplyAllForm: displayReplyAllForm,
    displayReplyAllFormAsync: displayReplyAllFormAsync,
    getAttachmentContentAsync: getAttachmentContent,
    getEntities: Entities_getEntities,
    getEntitiesByType: Entities_getEntitiesByType,
    getFilteredEntitiesByName: Entities_getFilteredEntitiesByName,
    getInitializationContextAsync: getInitializationContext,
    getRegExMatches: Entities_getRegExMatches,
    getRegExMatchesByName: Entities_getRegExMatchesByName,
    getSelectedEntities: Entities_getSelectedEntities,
    getSelectedRegExMatches: Entities_getSelectedRegExMatches,
    loadCustomPropertiesAsync: loadCustomProperties,
    isAllDayEvent: getInitialDataProp("isAllDayEvent"),
    sensitivity: getInitialDataProp("sensitivity")
  });
  return appointmentRead;
}
// CONCATENATED MODULE: ./src/validation/timeConstants.ts
var TimeType;

(function (TimeType) {
  TimeType[TimeType["start"] = 1] = "start";
  TimeType[TimeType["end"] = 2] = "end";
})(TimeType || (TimeType = {}));
// CONCATENATED MODULE: ./src/methods/getTime.ts




function getTime(namespace) {
  return function () {
    var args = [];

    for (var _i = 0; _i < arguments.length; _i++) {
      args[_i] = arguments[_i];
    }

    checkPermissionsAndThrow(1, namespace + ".getAsync");
    var commonParameters = parseCommonArgs(args, true, false);
    standardInvokeHostMethod(24, commonParameters.asyncContext, commonParameters.callback, {
      TimeProperty: TimeType[namespace]
    }, getTime_format);
  };
}

function getTime_format(rawInput) {
  var ticks = rawInput;
  return new Date(ticks);
}
// CONCATENATED MODULE: ./src/methods/setTime.ts






var maxTime = 8640000000000000;
var minTime = -8640000000000000;
function setTime(namespace) {
  return function (date) {
    var args = [];

    for (var _i = 1; _i < arguments.length; _i++) {
      args[_i - 1] = arguments[_i];
    }

    checkPermissionsAndThrow(2, namespace + ".setAsync");
    var commonParameters = parseCommonArgs(args, false, false);
    var parameters = {
      date: date
    };
    setTime_validateParameters(parameters);
    standardInvokeHostMethod(25, commonParameters.asyncContext, commonParameters.callback, {
      TimeProperty: TimeType[namespace],
      time: parameters.date.getTime()
    }, undefined);
  };
}

function setTime_validateParameters(parameters) {
  if (!isDateObject(parameters.date)) {
    throw createArgumentTypeError("dateTime", typeof parameters.date, typeof Date);
  }

  if (isNaN(parameters.date.getTime())) {
    throw createArgumentError("dateTime");
  }

  if (parameters.date.getTime() < minTime || parameters.date.getTime() > maxTime) {
    throw createArgumentOutOfRange("dateTime");
  }
}
// CONCATENATED MODULE: ./src/api/getTimeSurface.ts



function getTimeSurface(namespace) {
  return objectDefine({}, {
    getAsync: getTime(namespace),
    setAsync: setTime(namespace)
  });
}
// CONCATENATED MODULE: ./src/methods/getLocation.ts



function getLocation() {
  var args = [];

  for (var _i = 0; _i < arguments.length; _i++) {
    args[_i] = arguments[_i];
  }

  checkPermissionsAndThrow(1, "location.getAsync");
  var commonParameters = parseCommonArgs(args, true, false);
  standardInvokeHostMethod(26, commonParameters.asyncContext, commonParameters.callback, undefined, undefined);
}
// CONCATENATED MODULE: ./src/methods/setLocation.ts






var MaximumLocationLength = 255;
function setLocation(location) {
  var args = [];

  for (var _i = 1; _i < arguments.length; _i++) {
    args[_i - 1] = arguments[_i];
  }

  checkPermissionsAndThrow(2, "location.setAsync");
  var commonParameters = parseCommonArgs(args, false, false);
  var parameters = {
    location: location
  };
  setLocation_validateParameters(parameters);
  standardInvokeHostMethod(27, commonParameters.asyncContext, commonParameters.callback, parameters, undefined);
}

function setLocation_validateParameters(parameters) {
  if (!isNullOrUndefined(parameters.location)) {
    if (!(typeof parameters.location === "string")) {
      throw createArgumentTypeError("location", typeof parameters.location, "string");
    }

    throwOnOutOfRange(parameters.location.length, 0, MaximumLocationLength, "location");
  } else {
    throw createNullArgumentError("location");
  }
}
// CONCATENATED MODULE: ./src/api/getLocationSurface.ts



function getLocationSurface() {
  return objectDefine({}, {
    getAsync: getLocation,
    setAsync: setLocation
  });
}
// CONCATENATED MODULE: ./src/methods/getRecurrence.ts




function getRecurrence() {
  var args = [];

  for (var _i = 0; _i < arguments.length; _i++) {
    args[_i] = arguments[_i];
  }

  checkPermissionsAndThrow(1, "recurrenceProperties.getAsync");
  var commonParameters = parseCommonArgs(args, true, false);
  standardInvokeHostMethod(103, commonParameters.asyncContext, commonParameters.callback, undefined, seriesTimeJsonConverter);
}
function seriesTimeJsonConverter(rawInput) {
  if (rawInput !== null) {
    if (rawInput.seriesTimeJson !== null) {
      var seriesTime = new SeriesTime_SeriesTime();
      seriesTime.importFromSeriesTimeJsonObject(rawInput.seriesTimeJson);
      delete rawInput.seriesTimeJson;
      rawInput.seriesTime = seriesTime;
    }
  }

  return rawInput;
}
// CONCATENATED MODULE: ./src/validation/validateRecurrenceObject.ts






function validateRecurrenceObject(recurrenceObject) {
  if (isNullOrUndefined(recurrenceObject)) {
    return;
  }

  recurrenceObject = recurrenceObject;

  if (isNullOrUndefined(recurrenceObject.recurrenceType)) {
    throw createNullArgumentError(RecurrenceTypeKey);
  }

  if (isNullOrUndefined(recurrenceObject.seriesTime)) {
    throw createNullArgumentError(SeriesTimeKey);
  }

  if (!(recurrenceObject.seriesTime instanceof SeriesTime_SeriesTime) || !recurrenceObject.seriesTime.isValid()) {
    throw createArgumentError(SeriesTimeKey);
  }

  if (!recurrenceObject.seriesTime.isEndAfterStart()) {
    throw createArgumentError(SeriesTimeKey, getString("l_InvalidEventDates_Text"));
  }

  throwOnInvalidRecurrenceType(recurrenceObject.recurrenceType);

  if (recurrenceObject.recurrenceType !== MailboxEnums.RecurrenceType.Weekday) {
    if (isNullOrUndefined(recurrenceObject.recurrenceProperties)) {
      throw createNullArgumentError(RecurrenceTypeKey);
    }
  }

  if (!isNullOrUndefined(recurrenceObject.recurrenceTimeZone)) {
    if (isNullOrUndefined(recurrenceObject.recurrenceTimeZone.name)) {
      throw createNullArgumentError(RecurrenceTimeZoneName);
    }

    if (typeof recurrenceObject.recurrenceTimeZone.name !== "string") {
      throw createArgumentTypeError(RecurrenceTimeZoneName, typeof recurrenceObject.recurrenceTimeZone.name, "string");
    }
  }

  if (recurrenceObject.recurrenceType === MailboxEnums.RecurrenceType.Daily) {
    throwOnInvalidDailyRecurrence(recurrenceObject.recurrenceProperties);
  } else if (recurrenceObject.recurrenceType === MailboxEnums.RecurrenceType.Weekly) {
    throwOnInvalidWeeklyRecurrence(recurrenceObject.recurrenceProperties);
  } else if (recurrenceObject.recurrenceType === MailboxEnums.RecurrenceType.Monthly) {
    throwOnInvalidMonthlyRecurrence(recurrenceObject.recurrenceProperties);
  } else if (recurrenceObject.recurrenceType === MailboxEnums.RecurrenceType.Yearly) {
    throwOnInvalidYearlyRecurrence(recurrenceObject.recurrenceProperties);
  }
}

function throwOnInvalidRecurrenceType(recurrenceType) {
  if (recurrenceType !== MailboxEnums.RecurrenceType.Daily && recurrenceType !== MailboxEnums.RecurrenceType.Weekly && recurrenceType !== MailboxEnums.RecurrenceType.Weekday && recurrenceType !== MailboxEnums.RecurrenceType.Yearly && recurrenceType !== MailboxEnums.RecurrenceType.Monthly) {
    throw createArgumentError(RecurrenceTypeKey);
  }
}

function throwOnInvalidRecurrenceInterval(recurrenceProperties) {
  if (isNullOrUndefined(recurrenceProperties.interval)) {
    throw createNullArgumentError(IntervalKey);
  }

  if (typeof recurrenceProperties.interval !== "number") {
    throw createArgumentTypeError(IntervalKey, typeof recurrenceProperties.interval, "number");
  }

  if (recurrenceProperties.interval <= 0) {
    throw createArgumentError(IntervalKey);
  }
}

function throwOnInvalidDailyRecurrence(recurrenceProperties) {
  throwOnInvalidRecurrenceInterval(recurrenceProperties);
}

function throwOnInvalidWeeklyRecurrence(recurrenceProperties) {
  throwOnInvalidRecurrenceInterval(recurrenceProperties);

  if (isNullOrUndefined(recurrenceProperties.days)) {
    throw createArgumentTypeError(DaysKey);
  }

  if (!Array.isArray(recurrenceProperties.days)) {
    throw createArgumentTypeError(DaysKey);
  }

  throwOnInvalidDaysArray(recurrenceProperties.days);

  if (!isNullOrUndefined(recurrenceProperties.firstDayOfWeek)) {
    if (typeof recurrenceProperties.firstDayOfWeek !== "string") {
      throw createArgumentTypeError(FirstDayOfWeekKey);
    }

    if (!verifyDays(recurrenceProperties.firstDayOfWeek, false)) {
      throw createArgumentError(FirstDayOfWeekKey);
    }
  }
}

function throwOnInvalidDaysArray(daysArray) {
  for (var i = 0; i < daysArray.length; i++) {
    if (!verifyDays(daysArray[i], false)) {
      throw createArgumentError(DaysKey);
    }
  }
}

function verifyDays(dayEnum, checkGroupedDays) {
  var fRegularDay = dayEnum === MailboxEnums.Days.Mon || dayEnum === MailboxEnums.Days.Tue || dayEnum === MailboxEnums.Days.Wed || dayEnum === MailboxEnums.Days.Thu || dayEnum === MailboxEnums.Days.Fri || dayEnum === MailboxEnums.Days.Sat || dayEnum === MailboxEnums.Days.Sun;

  if (checkGroupedDays) {
    var fGroupedDay = dayEnum === MailboxEnums.Days.WeekendDay || dayEnum === MailboxEnums.Days.Weekday || dayEnum === MailboxEnums.Days.Day;
    return fGroupedDay || fRegularDay;
  } else {
    return fRegularDay;
  }
}

function throwOnInvalidMonthlyRecurrence(recurrenceProperties) {
  if (isNullOrUndefined(recurrenceProperties.interval)) {
    throw createNullArgumentError(IntervalKey);
  }

  if (typeof recurrenceProperties.interval !== "number") {
    throw createArgumentTypeError(IntervalKey, typeof recurrenceProperties.interval, "number");
  }

  if (!isNullOrUndefined(recurrenceProperties.dayOfMonth)) {
    if (typeof recurrenceProperties.dayOfMonth !== "number") {
      throw createArgumentTypeError(DayOfMonthKey, typeof recurrenceProperties.dayOfMonth, "number");
    }

    throwOnInvalidDayOfMonth(recurrenceProperties.dayOfMonth);
  } else if (!isNullOrUndefined(recurrenceProperties.dayOfWeek) && !isNullOrUndefined(recurrenceProperties.weekNumber)) {
    if (typeof recurrenceProperties.dayOfWeek !== "string") {
      throw createArgumentTypeError(DayOfWeekKey, typeof recurrenceProperties.dayOfWeek, "string");
    }

    if (!verifyDays(recurrenceProperties.dayOfWeek, true)) {
      throw createArgumentError(DayOfWeekKey);
    }

    if (typeof recurrenceProperties.weekNumber !== "string") {
      throw createArgumentTypeError(WeekNumberKey, typeof recurrenceProperties.weekNumber, "string");
    }

    throwOnInvalidWeekNumber(recurrenceProperties.weekNumber);
  } else {
    throw createArgumentError(undefined, getString("l_Recurrence_Error_Properties_Invalid_Text"));
  }
}

function throwOnInvalidWeekNumber(weekNumber) {
  if (weekNumber !== MailboxEnums.WeekNumber.First && weekNumber !== MailboxEnums.WeekNumber.Second && weekNumber !== MailboxEnums.WeekNumber.Third && weekNumber !== MailboxEnums.WeekNumber.Fourth && weekNumber !== MailboxEnums.WeekNumber.Last) {
    throw createArgumentError(WeekNumberKey);
  }
}

function throwOnInvalidDayOfMonth(iDayOfMonth) {
  if (iDayOfMonth < 1 || iDayOfMonth > 31) {
    throw createArgumentError(DayOfMonthKey);
  }
}

function throwOnInvalidYearlyRecurrence(recurrenceProperties) {
  if (isNullOrUndefined(recurrenceProperties.interval)) {
    throw createNullArgumentError(IntervalKey);
  }

  if (typeof recurrenceProperties.interval !== "number") {
    throw createArgumentTypeError(IntervalKey, typeof recurrenceProperties.interval, "number");
  }

  if (isNullOrUndefined(recurrenceProperties.month)) {
    throw createNullArgumentError(MonthKey);
  }

  if (typeof recurrenceProperties.month !== "string") {
    throw createArgumentTypeError(MonthKey, typeof recurrenceProperties.month, "string");
  }

  throwOnInvalidMonth(recurrenceProperties.month);

  if (!isNullOrUndefined(recurrenceProperties.dayOfMonth)) {
    if (typeof recurrenceProperties.dayOfMonth !== "number") {
      throw createArgumentTypeError(DayOfMonthKey, typeof recurrenceProperties.dayOfMonth, "number");
    }

    throwOnInvalidDayOfMonth(recurrenceProperties.dayOfMonth);
  } else if (!isNullOrUndefined(recurrenceProperties.weekNumber) && !isNullOrUndefined(recurrenceProperties.dayOfWeek)) {
    if (typeof recurrenceProperties.dayOfWeek !== "string") {
      throw createArgumentTypeError(DayOfWeekKey, typeof recurrenceProperties.dayOfWeek, "string");
    }

    if (!verifyDays(recurrenceProperties.dayOfWeek, true)) {
      throw createArgumentError(DayOfWeekKey);
    }

    if (typeof recurrenceProperties.weekNumber !== "string") {
      throw createArgumentTypeError(WeekNumberKey, typeof recurrenceProperties.weekNumber, "string");
    }

    throwOnInvalidWeekNumber(recurrenceProperties.weekNumber);
  } else {
    throw createArgumentError(undefined, getString("l_Recurrence_Error_Properties_Invalid_Text"));
  }
}

function throwOnInvalidMonth(month) {
  if (month !== MailboxEnums.Month.Jan && month !== MailboxEnums.Month.Feb && month !== MailboxEnums.Month.Mar && month !== MailboxEnums.Month.Apr && month !== MailboxEnums.Month.May && month !== MailboxEnums.Month.Jun && month !== MailboxEnums.Month.Jul && month !== MailboxEnums.Month.Aug && month !== MailboxEnums.Month.Sep && month !== MailboxEnums.Month.Oct && month !== MailboxEnums.Month.Nov && month !== MailboxEnums.Month.Dec) {
    throw createArgumentError(MonthKey);
  }
}
// CONCATENATED MODULE: ./src/methods/setRecurrence.ts









function setRecurrence(recurrencePattern) {
  var args = [];

  for (var _i = 1; _i < arguments.length; _i++) {
    args[_i - 1] = arguments[_i];
  }

  checkPermissionsAndThrow(2, "recurrenceProperties.setAsync");
  var seriesId = getAppointmentCompose().seriesId;

  if (!isNullOrUndefined(seriesId) && seriesId.length > 0) {
    throw createArgumentError(undefined, getString("l_Recurrence_Error_Instance_SetAsync_Text"));
  }

  validateRecurrenceObject(recurrencePattern);
  var commonParameters = parseCommonArgs(args, false, false);
  var recurrenceData = convertSeriesTime(recurrencePattern);
  var parameters = {
    recurrenceData: recurrenceData
  };
  standardInvokeHostMethod(104, commonParameters.asyncContext, commonParameters.callback, parameters, undefined);
}

function convertSeriesTime(recurrencePattern) {
  if (recurrencePattern !== null && recurrencePattern.seriesTime !== null) {
    if (recurrencePattern.seriesTime instanceof SeriesTime_SeriesTime) {
      var recurrencePatternCopy = {
        recurrenceProperties: recurrencePattern.recurrenceProperties,
        recurrenceTimeZone: recurrencePattern.recurrenceTimeZone,
        recurrenceType: recurrencePattern.recurrenceType,
        seriesTimeJson: recurrencePattern.seriesTime.exportToSeriesTimeJson()
      };
      return recurrencePatternCopy;
    }
  }

  return recurrencePattern;
}
// CONCATENATED MODULE: ./src/api/getRecurrenceSurface.ts



function getRecurrenceSurface(isCompose) {
  var recurrence = objectDefine({}, {
    getAsync: getRecurrence
  });

  if (isCompose) {
    objectDefine(recurrence, {
      setAsync: setRecurrence
    });
  }

  return recurrence;
}
// CONCATENATED MODULE: ./src/methods/getAllDayEvent.ts




function getAllDayEvent() {
  var args = [];

  for (var _i = 0; _i < arguments.length; _i++) {
    args[_i] = arguments[_i];
  }

  checkPermissionsAndThrow(1, "isAllDayEvent.getAsync");
  var commonParameters = parseCommonArgs(args, true, false);
  checkFeatureEnabledAndThrow(Features.calendarItems, "isAllDayEvent.getAsync");
  standardInvokeHostMethod(169, commonParameters.asyncContext, commonParameters.callback, undefined, undefined);
}
// CONCATENATED MODULE: ./src/methods/setAllDayEvent.ts






function setAllDayEvent(isAllDayEvent) {
  var args = [];

  for (var _i = 1; _i < arguments.length; _i++) {
    args[_i - 1] = arguments[_i];
  }

  checkPermissionsAndThrow(2, "isAllDayEvent.setAsync");
  var commonParameters = parseCommonArgs(args, true, false);
  var parameters = {
    isAllDayEvent: isAllDayEvent
  };
  checkFeatureEnabledAndThrow(Features.calendarItems, "isAllDayEvent.setAsync");
  setAllDayEvent_validateParameters(parameters);
  standardInvokeHostMethod(170, commonParameters.asyncContext, commonParameters.callback, parameters, undefined);
}

function setAllDayEvent_validateParameters(parameters) {
  if (isNullOrUndefined(parameters.isAllDayEvent)) {
    throw createNullArgumentError("isAllDayEvent");
  }

  if (typeof parameters.isAllDayEvent !== "boolean") {
    throw createArgumentTypeError("isAllDayEvent", typeof parameters.isAllDayEvent, "boolean");
  }
}
// CONCATENATED MODULE: ./src/api/getAllDayEventSurface.ts



function getAllDayEventSurface() {
  return objectDefine({}, {
    getAsync: getAllDayEvent,
    setAsync: setAllDayEvent
  });
}
// CONCATENATED MODULE: ./src/methods/setSensitivity.ts







function setSensitivity(sensitivity) {
  var args = [];

  for (var _i = 1; _i < arguments.length; _i++) {
    args[_i - 1] = arguments[_i];
  }

  checkPermissionsAndThrow(2, "sensitivity.setAsync");
  var commonParameters = parseCommonArgs(args, true, false);
  var parameters = {
    sensitivity: sensitivity
  };
  checkFeatureEnabledAndThrow(Features.calendarItems, "sensitivity.setAsync");
  setSensitivity_validateParameters(parameters);
  standardInvokeHostMethod(172, commonParameters.asyncContext, commonParameters.callback, parameters, undefined);
}

function setSensitivity_validateParameters(parameters) {
  validateStringParam("sensitivity", parameters.sensitivity);
  throwOnInvalidSensitivityType(parameters.sensitivity);
}

function throwOnInvalidSensitivityType(sensitivity) {
  if (sensitivity !== MailboxEnums.AppointmentSensitivityType.Normal && sensitivity !== MailboxEnums.AppointmentSensitivityType.Personal && sensitivity !== MailboxEnums.AppointmentSensitivityType.Private && sensitivity !== MailboxEnums.AppointmentSensitivityType.Confidential) {
    throw createArgumentError("sensitivity");
  }
}
// CONCATENATED MODULE: ./src/methods/getSensitivity.ts




function getSensitivity() {
  var args = [];

  for (var _i = 0; _i < arguments.length; _i++) {
    args[_i] = arguments[_i];
  }

  checkPermissionsAndThrow(1, "sensitivity.getAsync");
  var commonParameters = parseCommonArgs(args, true, false);
  checkFeatureEnabledAndThrow(Features.calendarItems, "sensitivity.getAsync");
  standardInvokeHostMethod(171, commonParameters.asyncContext, commonParameters.callback, undefined, undefined);
}
// CONCATENATED MODULE: ./src/api/getSensitivitySurface.ts



function getSensitivitySurface() {
  return objectDefine({}, {
    getAsync: getSensitivity,
    setAsync: setSensitivity
  });
}
// CONCATENATED MODULE: ./src/api/getAppointmentCompose.ts






























function getAppointmentCompose() {
  var appointmentCompose = objectDefine({}, {
    body: getBodySurface(true),
    categories: getCategoriesSurface(),
    end: getTimeSurface("end"),
    enhancedLocation: getEnhancedLocationsSurface(true),
    itemType: "appointment",
    location: getLocationSurface(),
    notificationMessages: getNotificationMessageSurface(),
    optionalAttendees: getRecipientsSurface("optionalAttendees"),
    organizer: getFromSurface("organizer"),
    recurrence: getRecurrenceSurface(true),
    requiredAttendees: getRecipientsSurface("requiredAttendees"),
    seriesId: getInitialDataProp("seriesId"),
    start: getTimeSurface("start"),
    subject: getSubjectSurface(),
    addFileAttachmentAsync: addFileAttachment,
    addFileAttachmentFromBase64Async: addBase64FileAttachment,
    addItemAttachmentAsync: addItemAttachment,
    close: close_close,
    getAttachmentsAsync: getAttachments_getAttachments,
    getAttachmentContentAsync: getAttachmentContent,
    getInitializationContextAsync: getInitializationContext,
    getItemIdAsync: getItemId,
    getSelectedDataAsync: getSelectedData,
    loadCustomPropertiesAsync: loadCustomProperties,
    removeAttachmentAsync: removeAttachment,
    saveAsync: save,
    setSelectedDataAsync: setSelectedData(29),
    isAllDayEvent: getAllDayEventSurface(),
    sensitivity: getSensitivitySurface(),
    isClientSignatureEnabledAsync: isClientSignatureEnabled,
    disableClientSignatureAsync: disableClientSignature,
    sessionData: getSessionDataSurface()
  });
  return appointmentCompose;
}
// CONCATENATED MODULE: ./src/utils/addEventSupport.ts
var addEventSupport_OSF = __webpack_require__(0);

var addEventSupport_Microsoft = __webpack_require__(1);

var addEventSupport = function addEventSupport(target) {
  addEventSupport_OSF.DDA.DispIdHost.addEventSupport(target, new addEventSupport_OSF.EventDispatch([addEventSupport_Microsoft.Office.WebExtension.EventType.RecipientsChanged, addEventSupport_Microsoft.Office.WebExtension.EventType.AppointmentTimeChanged, addEventSupport_Microsoft.Office.WebExtension.EventType.AttachmentsChanged, addEventSupport_Microsoft.Office.WebExtension.EventType.EnhancedLocationsChanged, addEventSupport_Microsoft.Office.WebExtension.EventType.InfobarClicked, addEventSupport_Microsoft.Office.WebExtension.EventType.RecurrenceChanged]));
};
// CONCATENATED MODULE: ./src/methods/registerConsent.ts



var ConsentStateType;

(function (ConsentStateType) {
  ConsentStateType[ConsentStateType["NotResponded"] = 0] = "NotResponded";
  ConsentStateType[ConsentStateType["NotConsented"] = 1] = "NotConsented";
  ConsentStateType[ConsentStateType["Consented"] = 2] = "Consented";
})(ConsentStateType || (ConsentStateType = {}));

function registerConsent(consentState) {
  var parameters = {
    consentState: consentState,
    extensionId: getInitialDataProp("extensionId")
  };
  registerConsent_validateParameters(consentState);
  standardInvokeHostMethod(40, undefined, undefined, parameters, undefined);
}

function registerConsent_validateParameters(consentState) {
  if (consentState !== ConsentStateType.Consented && consentState !== ConsentStateType.NotConsented && consentState !== ConsentStateType.NotResponded) {
    throw createArgumentOutOfRange("consentState");
  }
}
// CONCATENATED MODULE: ./src/methods/navigateToModule.ts





function navigateToModule(moduleName) {
  var args = [];

  for (var _i = 1; _i < arguments.length; _i++) {
    args[_i - 1] = arguments[_i];
  }

  var commonParameters = parseCommonArgs(args, false, false);
  var parameters = {
    module: moduleName
  };
  navigateToModule_validateParameters(moduleName);

  if (moduleName === MailboxEnums.ModuleType.Addins) {
    if (!!commonParameters.options && !!commonParameters.options.queryString) {
      parameters.queryString = commonParameters.options.queryString;
    } else {
      parameters.queryString = "";
    }
  }

  standardInvokeHostMethod(45, commonParameters.asyncContext, commonParameters.callback, parameters, undefined);
}

function navigateToModule_validateParameters(moduleName) {
  if (isNullOrUndefined(moduleName)) {
    throw createNullArgumentError("module");
  }

  if (moduleName === "") {
    throw createArgumentError("module");
  }

  if (moduleName !== MailboxEnums.ModuleType.Addins) {
    throw createArgumentError("module");
  }
}
// CONCATENATED MODULE: ./src/methods/recordDataPoint.ts



function recordDataPoint(data) {
  if (isNullOrUndefined(data)) {
    throw createNullArgumentError("data");
  }

  standardInvokeHostMethod(402, undefined, undefined, data, undefined);
}
// CONCATENATED MODULE: ./src/methods/recordTrace.ts



function recordTrace(data) {
  if (isNullOrUndefined(data)) {
    throw createNullArgumentError("data");
  }

  standardInvokeHostMethod(401, undefined, undefined, data, undefined);
}
// CONCATENATED MODULE: ./src/methods/trackCtq.ts



function trackCtq(data) {
  if (isNullOrUndefined(data)) {
    throw createNullArgumentError("data");
  }

  standardInvokeHostMethod(400, undefined, undefined, data, undefined);
}
// CONCATENATED MODULE: ./src/methods/windowOpenOverrideHandler.ts

function windowOpenOverrideHandler(url, target, features, replace) {
  standardInvokeHostMethod(403, undefined, undefined, {
    launchUrl: url
  }, undefined);
  return window;
}
// CONCATENATED MODULE: ./src/methods/logTelemetry.ts



function logTelemetry_logTelemetry(data) {
  if (isNullOrUndefined(data)) {
    throw createNullArgumentError("data");
  }

  standardInvokeHostMethod(163, undefined, undefined, {
    telemetryData: data
  }, undefined);
}
// CONCATENATED MODULE: ./src/methods/logCustomerContentTelemetry.ts



function logCustomerContentTelemetry(data) {
  if (isNullOrUndefined(data)) {
    throw createNullArgumentError("data");
  }

  standardInvokeHostMethod(193, undefined, undefined, {
    telemetryData: data
  }, undefined);
}
// CONCATENATED MODULE: ./src/utils/convertToLocalClientTime.ts
var __assign = undefined && undefined.__assign || function () {
  __assign = Object.assign || function (t) {
    for (var s, i = 1, n = arguments.length; i < n; i++) {
      s = arguments[i];

      for (var p in s) {
        if (Object.prototype.hasOwnProperty.call(s, p)) t[p] = s[p];
      }
    }

    return t;
  };

  return __assign.apply(this, arguments);
};







function convertToLocalClientTime(timeValue) {
  if (!isDateObject(timeValue)) {
    throw createArgumentError("timeValue");
  }

  var date = new Date(timeValue.getTime());
  var offset = date.getTimezoneOffset() * -1;

  if (!isNullOrUndefined(getInitialDataProp("timeZoneOffsets"))) {
    date.setUTCMinutes(date.getUTCMinutes() - offset);
    offset = findOffset(date);
    date.setUTCMinutes(date.getUTCMinutes() + offset);
  }

  var retValue = __assign({
    timezoneOffset: offset
  }, dateToDictionary(date));

  return retValue;
}
// CONCATENATED MODULE: ./src/methods/displayPersonaCardAsync.ts




function displayPersonaCardAsync(ewsIdOrEmail) {
  var args = [];

  for (var _i = 1; _i < arguments.length; _i++) {
    args[_i - 1] = arguments[_i];
  }

  var commonParameters = parseCommonArgs(args, false, false);
  var parameters = {
    ewsIdOrEmail: ewsIdOrEmail
  };
  displayPersonaCardAsync_validateParameters(parameters);
  standardInvokeHostMethod(43, commonParameters.asyncContext, commonParameters.callback, {
    ewsIdOrEmail: ewsIdOrEmail.trim()
  }, undefined);
}

function displayPersonaCardAsync_validateParameters(parameters) {
  if (!isNullOrUndefined(parameters.ewsIdOrEmail)) {
    displayPersonaCardAsync_throwOnInvalidItemId(parameters.ewsIdOrEmail);

    if (parameters.ewsIdOrEmail === "") {
      throw createArgumentError("ewsIdOrEmail", "ewsIdOrEmail cannot be empty.");
    }
  } else {
    throw createNullArgumentError("ewsIdOrEmail");
  }
}

function displayPersonaCardAsync_throwOnInvalidItemId(ewsIdOrEmail) {
  if (!(typeof ewsIdOrEmail === "string")) {
    throw createArgumentError("ewsIdOrEmail");
  }
}
// CONCATENATED MODULE: ./src/methods/getSharedProperties.ts



function getSharedProperties() {
  var args = [];

  for (var _i = 0; _i < arguments.length; _i++) {
    args[_i] = arguments[_i];
  }

  checkPermissionsAndThrow(1, "item.getSharedPropertiesAsync");
  var commonParameters = parseCommonArgs(args, true, false);
  standardInvokeHostMethod(108, commonParameters.asyncContext, commonParameters.callback, undefined, undefined);
}
// CONCATENATED MODULE: ./src/utils/addSharedPropertiesSupport.ts





var addSharedPropertiesSupport_addSharedPropertiesSupport = function addSharedPropertiesSupport(target) {
  if (target && getInitialDataProp("isFromSharedFolder") && getHostItemType_getHostItemType() !== HostItemType.ItemLess) {
    objectDefine(target, {
      getSharedPropertiesAsync: getSharedProperties
    });
  }
};
// CONCATENATED MODULE: ./src/api/prepareApiSurface.ts




































var prepareApiSurface_OSF = __webpack_require__(0);

var prepareApiSurface_createMailboxSurface = function createMailboxSurface(target) {
  objectDefine(target, {
    ewsUrl: getInitialDataProp("ewsUrl"),
    restUrl: getInitialDataProp("restUrl"),
    displayAppointmentForm: displayAppointmentForm,
    displayAppointmentFormAsync: displayAppointmentFormAsync,
    displayMessageForm: displayMessageForm,
    displayMessageFormAsync: displayMessageFormAsync,
    displayPersonaCardAsync: displayPersonaCardAsync,
    getCallbackTokenAsync: getCallbackToken,
    getUserIdentityTokenAsync: getUserIdentityToken,
    logTelemetry: logTelemetry_logTelemetry,
    logCustomerContentTelemetry: logCustomerContentTelemetry,
    makeEwsRequestAsync: makeEwsRequest,
    masterCategories: getMasterCategoriesSurface(),
    navigateToModuleAsync: navigateToModule,
    diagnostics: getDiagnosticsSurface(),
    userProfile: getUserProfileSurface(),
    convertToEwsId: convertToEwsId,
    convertToLocalClientTime: convertToLocalClientTime,
    convertToRestId: convertToRestId,
    convertToUtcClientTime: convertToUtcClientTime,
    RegisterConsentAsync: registerConsent,
    GetIsRead: function GetIsRead() {
      return getInitialDataProp("isRead");
    },
    GetEndPointUrl: function GetEndPointUrl() {
      return getInitialDataProp("endNodeUrl");
    },
    GetConsentMetaData: function GetConsentMetaData() {
      return getInitialDataProp("consentMetadata");
    },
    GetMarketplaceContentMarket: function GetMarketplaceContentMarket() {
      return getInitialDataProp("marketplaceContentMarket");
    },
    GetMarketplaceAssetId: function GetMarketplaceAssetId() {
      return getInitialDataProp("marketplaceAssetId");
    },
    GetExtensionId: function GetExtensionId() {
      return getInitialDataProp("extensionId");
    },
    CloseApp: closeApp,
    recordDataPoint: recordDataPoint,
    recordTrace: recordTrace,
    trackCtq: trackCtq
  });

  if (getHostItemType_getHostItemType() !== HostItemType.MessageCompose && getHostItemType_getHostItemType() !== HostItemType.AppointmentCompose) {
    objectDefine(target, {
      displayNewAppointmentForm: displayNewAppointmentForm,
      displayNewMessageForm: displayNewMessageForm,
      displayNewAppointmentFormAsync: displayNewAppointmentFormAsync,
      displayNewMessageFormAsync: displayNewMessageFormAsync
    });
  }

  if (getAppName() === prepareApiSurface_OSF.AppName.OutlookWebApp && getInitialDataProp("openWindowOpen")) {
    window.open = windowOpenOverrideHandler;
  }

  return target;
};
var prepareApiSurface_getItem = function getItem() {
  var item = undefined;

  switch (getHostItemType_getHostItemType()) {
    case HostItemType.Message:
      item = getMessageRead();
      break;

    case HostItemType.MessageCompose:
      item = getMessageCompose();
      break;

    case HostItemType.Appointment:
      item = getAppointmentRead();
      break;

    case HostItemType.AppointmentCompose:
      item = getAppointmentCompose();
      break;

    case HostItemType.MeetingRequest:
      item = getMessageRead();
      break;

    default:
      return undefined;
  }

  if (isOutlookJs()) {
    prepareApiSurface_OSF.OutlookInitializationHelper.addEventDispatchToTarget(item, prepareApiSurface_OSF.OutlookInitializationHelper.getMailboxItemEventDispatch());
  } else {
    addEventSupport(item);
  }

  addSharedPropertiesSupport_addSharedPropertiesSupport(item);
  return item;
};
// CONCATENATED MODULE: ./src/utils/isOutlook16.ts

var isOutlook16_isOutlook16OrGreater = function isOutlook16OrGreater(hostVersion) {
  var endIndex = 0;
  var majorVersionNumber = 0;

  if (!isNullOrUndefined(hostVersion)) {
    endIndex = hostVersion.indexOf(".");
    majorVersionNumber = parseInt(hostVersion.substring(0, endIndex));
  }

  return majorVersionNumber >= 16;
};
// CONCATENATED MODULE: ./src/utils/isApiVersionSupported.ts
var isApiVersionSupported = function isApiVersionSupported(requirementSet, officeAppContext) {
  var apiSupported = false;

  try {
    var requirementDict = JSON.parse(officeAppContext.get_requirementMatrix());
    var hostApiVersion = requirementDict["Mailbox"];
    var hostApiVersionParts = hostApiVersion.split(".");
    var requirementSetParts = requirementSet.split(".");

    if (parseInt(hostApiVersionParts[0]) > parseInt(requirementSetParts[0]) || parseInt(hostApiVersionParts[0]) === parseInt(requirementSetParts[0]) && parseInt(hostApiVersionParts[1]) >= parseInt(requirementSetParts[1])) {
      apiSupported = true;
    }
  } catch (_a) {}

  return apiSupported;
};
// CONCATENATED MODULE: ./src/api/OutlookAppOm.ts








var OutlookAppOm_OSF = __webpack_require__(0);

var appInstance;
var whenStringsFinish;
var getInitialDataProp = function getInitialDataProp(key) {
  return appInstance && appInstance.getInitialDataProp(key);
};
var getIsNoItemContextWebExt = function getIsNoItemContextWebExt() {
  return !appInstance || !appInstance.item;
};
var getAppName = function getAppName() {
  return appInstance && appInstance.getAppName();
};

var OutlookAppOm_OutlookAppOm = function () {
  function OutlookAppOm(appContext, targetWindow, appReadyCallback) {
    var _this = this;

    this.displayName = "mailbox";

    this.stringLoadedCallback = function () {
      if (!!_this.appReadyCallback) {
        if (!_this.officeAppContext.get_isDialog()) {
          standardInvokeHostMethod_invokeHostMethod(1, undefined, _this.onInitialDataResponse);
        } else {
          setTimeout(function () {
            return _this.appReadyCallback();
          }, 0);
        }
      }
    };

    this.initialize = function (data) {
      if (data === null || data === undefined) {
        recreateAdditionalGlobalParametersSingleton(true);
        _this.initialData = null;
        _this.item = null;
      } else {
        _this.initialData = data;
        _this.initialData.permissionLevel = calculatePermissionLevel();
        _this.item = prepareApiSurface_getItem();
        var supportsAdditionalParameters = false;
        supportsAdditionalParameters = getAppName() !== OutlookAppOm_OSF.AppName.Outlook || isOutlook16_isOutlook16OrGreater(getInitialDataProp("hostVersion")) || isApiVersionSupported("1.5", _this.officeAppContext);
        recreateAdditionalGlobalParametersSingleton(supportsAdditionalParameters);

        if (typeof data.itemNumber !== "undefined") {
          getAdditionalGlobalParametersSingleton().setCurrentItemNumber(data.itemNumber);
        }
      }
    };

    this.onInitialDataResponse = function (resultCode, data) {
      if (!!resultCode && resultCode !== InvokeResultCode.noError) {
        return;
      }

      _this.initialize(data);

      prepareApiSurface_createMailboxSurface(_this);
      setTimeout(function () {
        return _this.appReadyCallback();
      }, 0);
    };

    this.officeAppContext = appContext;
    this.targetWindow = window;
    this.appReadyCallback = appReadyCallback;
    appInstance = this;
    loadLocalizedScript(this.stringLoadedCallback);
  }

  OutlookAppOm.prototype.getAppName = function () {
    var retVal = -1;
    retVal = this.officeAppContext.get_appName();
    return retVal;
  };

  OutlookAppOm.prototype.getInitialDataProp = function (key) {
    return this.initialData && this.initialData[key];
  };

  OutlookAppOm.prototype.setCurrentItemNumber = function (newItemNumber) {
    getAdditionalGlobalParametersSingleton().setCurrentItemNumber(newItemNumber);
  };

  OutlookAppOm.addAdditionalArgs = function (dispid, hostCallingArgs) {
    return hostCallingArgs;
  };

  OutlookAppOm.shouldRunInitialDataResponse = function () {
    return true;
  };

  return OutlookAppOm;
}();



var calculatePermissionLevel = function calculatePermissionLevel() {
  var HostReadItem = 1;
  var HostReadWriteMailbox = 2;
  var HostReadWriteItem = 3;
  var permissionLevelFromHost = getInitialDataProp("permissionLevel");

  if (permissionLevelFromHost === undefined) {
    return 0;
  }

  switch (permissionLevelFromHost) {
    case HostReadItem:
      return 1;

    case HostReadWriteItem:
      return 2;

    case HostReadWriteMailbox:
      return 3;

    default:
      return 0;
  }
};
// CONCATENATED MODULE: ./src/methods/saveSettingsRequest.ts





var saveSettingsRequest_OSF = __webpack_require__(0);

var settingsMaxNumberOfCharacters = 32 * 1024;
function saveSettingsRequest(data) {
  var args = [];

  for (var _i = 1; _i < arguments.length; _i++) {
    args[_i - 1] = arguments[_i];
  }

  var commonParameters = parseCommonArgs(args, false, false);
  var serializedSettings = saveSettingsRequest_OSF.DDA.SettingsManager.serializeSettings(data);

  if (JSON.stringify(serializedSettings).length > settingsMaxNumberOfCharacters) {
    var asyncResult_1 = createAsyncResult(undefined, saveSettingsRequest_OSF.DDA.AsyncResultEnum.ErrorCode.Failed, 9019, commonParameters.asyncContext, "");

    if (!!commonParameters.callback) {
      setTimeout(function () {
        if (!!commonParameters.callback) commonParameters.callback(asyncResult_1);
      }, 0);
    }

    return;
  }

  if (saveSettingsRequest_OSF.AppName.OutlookWebApp === getAppName()) {
    saveSettingsForOwa(commonParameters, serializedSettings);
  } else {
    saveSettingsForOutlookDesktop(commonParameters, serializedSettings);
  }
}

function saveSettingsForOwa(commonParameters, serializedSettings) {
  standardInvokeHostMethod(404, commonParameters.asyncContext, commonParameters.callback, [serializedSettings], undefined);
}

function saveSettingsForOutlookDesktop(commonParameters, serializedSettings) {
  var detailedErrorCode = -1;
  var storedException = null;

  try {
    var jsonSettings = JSON.stringify(serializedSettings);
    var settingsObjectToSave = {};
    settingsObjectToSave.SettingsKey = jsonSettings;
    saveSettingsRequest_OSF.DDA.ClientSettingsManager.write(settingsObjectToSave);
  } catch (ex) {
    storedException = ex;
  }

  var asyncResult = undefined;

  if (storedException != null) {
    detailedErrorCode = 9019;
    asyncResult = createAsyncResult(undefined, saveSettingsRequest_OSF.DDA.AsyncResultEnum.ErrorCode.Failed, detailedErrorCode, commonParameters.asyncContext, storedException.Message);
  } else {
    detailedErrorCode = 0;
    asyncResult = createAsyncResult(undefined, saveSettingsRequest_OSF.DDA.AsyncResultEnum.ErrorCode.Success, detailedErrorCode, commonParameters.asyncContext);
  }

  if (!!commonParameters.callback) {
    commonParameters.callback(asyncResult);
  }
}
// CONCATENATED MODULE: ./src/api/Settings.ts
var Settings_spreadArrays = undefined && undefined.__spreadArrays || function () {
  for (var s = 0, i = 0, il = arguments.length; i < il; i++) {
    s += arguments[i].length;
  }

  for (var r = Array(s), k = 0, i = 0; i < il; i++) {
    for (var a = arguments[i], j = 0, jl = a.length; j < jl; j++, k++) {
      r[k] = a[j];
    }
  }

  return r;
};




var Settings_OSF = __webpack_require__(0);

var Settings_Settings = function () {
  function Settings(deserializedData) {
    this.rawData = deserializedData;
    this.settingsData = null;
  }

  Settings.prototype.getSettingsData = function () {
    if (this.settingsData == null) {
      this.settingsData = this.convertFromRawSettings(this.rawData);
      this.rawData = null;
    }

    return this.settingsData;
  };

  Settings.prototype.get = function (key) {
    return this.getSettingsData()[key];
  };

  Settings.prototype.set = function (key, value) {
    this.getSettingsData()[key] = value;
  };

  Settings.prototype.remove = function (key) {
    delete this.getSettingsData()[key];
  };

  Settings.prototype.saveAsync = function () {
    var args = [];

    for (var _i = 0; _i < arguments.length; _i++) {
      args[_i] = arguments[_i];
    }

    saveSettingsRequest.apply(void 0, Settings_spreadArrays([this.getSettingsData()], args));
  };

  Settings.prototype.convertFromRawSettings = function (rawSettings) {
    if (rawSettings == null) {
      return {};
    }

    if (getAppName() !== Settings_OSF.AppName.OutlookWebApp) {
      var outlookSettings = rawSettings.SettingsKey;

      if (!!outlookSettings) {
        return Settings_OSF.DDA.SettingsManager.deserializeSettings(outlookSettings);
      }
    }

    return rawSettings;
  };

  return Settings;
}();


// CONCATENATED MODULE: ./src/api/Intellisense.ts



var Intellisense = {
  toItemRead: function toItemRead(item) {
    var hostItemtype = getHostItemType_getHostItemType();

    if (hostItemtype === HostItemType.Message || hostItemtype === HostItemType.Appointment || hostItemtype === HostItemType.MeetingRequest) {
      return item;
    }

    throw createArgumentTypeError();
  },
  toItemCompose: function toItemCompose(item) {
    var hostItemtype = getHostItemType_getHostItemType();

    if (hostItemtype === HostItemType.MessageCompose || hostItemtype === HostItemType.AppointmentCompose) {
      return item;
    }

    throw createArgumentTypeError();
  },
  toMessage: function toMessage(item) {
    return Intellisense.toMessageRead(item);
  },
  toMessageRead: function toMessageRead(item) {
    if (getHostItemType_getHostItemType() === HostItemType.Message || getHostItemType_getHostItemType() === HostItemType.MeetingRequest) {
      return item;
    }

    throw createArgumentTypeError();
  },
  toMessageCompose: function toMessageCompose(item) {
    if (getHostItemType_getHostItemType() === HostItemType.MessageCompose) {
      return item;
    }

    throw createArgumentTypeError();
  },
  toMeetingRequest: function toMeetingRequest(item) {
    if (getHostItemType_getHostItemType() === HostItemType.MeetingRequest) {
      return item;
    }

    throw createArgumentTypeError();
  },
  toAppointment: function toAppointment(item) {
    if (getHostItemType_getHostItemType() === HostItemType.Appointment) {
      return item;
    }

    throw createArgumentTypeError();
  },
  toAppointmentRead: function toAppointmentRead(item) {
    if (getHostItemType_getHostItemType() === HostItemType.Appointment) {
      return item;
    }

    throw createArgumentTypeError();
  },
  toAppointmentCompose: function toAppointmentCompose(item) {
    if (getHostItemType_getHostItemType() === HostItemType.AppointmentCompose) {
      return item;
    }

    throw createArgumentTypeError();
  }
};
// CONCATENATED MODULE: ./src/api/OutlookBase.ts


var OutlookBase = {
  SeriesTimeJsonConverter: function SeriesTimeJsonConverter(rawInput) {
    if (rawInput !== null && typeof rawInput === "object") {
      if (rawInput.seriesTimeJson !== null) {
        var seriesTime = new SeriesTime_SeriesTime();
        seriesTime.importFromSeriesTimeJsonObject(rawInput.seriesTimeJson);
        delete rawInput["seriesTimeJson"];
        rawInput.seriesTime = seriesTime;
      }
    }

    return rawInput;
  },
  CreateAttachmentDetails: function CreateAttachmentDetails(data) {
    convertAttachmentType(data);
    return data;
  }
};
// CONCATENATED MODULE: ./src/index.tsx






OSF = typeof OSF === "object" ? OSF : {};
OSF.DDA = OSF.DDA || {};
OSF.DDA.Settings = Settings_Settings;
OSF = typeof OSF === "object" ? OSF : {};
OSF.DDA = OSF.DDA || {};
OSF.DDA.OutlookAppOm = OutlookAppOm_OutlookAppOm;
Office = typeof Office === "object" ? Office : {};
Office.cast = Office.cast || {};
Office.cast.item = Intellisense;
Microsoft.Office.WebExtension.MailboxEnums = MailboxEnums;
Microsoft.Office.WebExtension.CoercionType = CoercionType;
Microsoft.Office.WebExtension.SeriesTime = SeriesTime_SeriesTime;
Microsoft.Office.WebExtension.OutlookBase = OutlookBase;
/* harmony default export */ var src = __webpack_exports__["default"] = (OutlookAppOm_OutlookAppOm);
var hWindow = window;
hWindow.$h = typeof $h === "object" ? $h : {};
hWindow.$h.Message = $h.Message || {};
hWindow.$h.Appointment = $h.Appointment || {};

hWindow.$h.Message.isInstanceOfType = function (item) {
  return item && item.itemType === "message";
};

hWindow.$h.Appointment.isInstanceOfType = function (item) {
  return item && item.itemType === "appointment";
};

/***/ })
/******/ ])["default"];
//# sourceMappingURL=outlook.api.js.mapvar OSF;
(function (OSF) {
    function isNodeJs() {
        try {
            return (typeof process === 'object'
                && String(process) === '[object process]');
        }
        catch (e) {
            return false;
        }
    }
    if (!isNodeJs() && !OSF.isOfficeReactNative()) {
        try {
            OSF.Flights = OSF.OUtil.parseFlights(true);
        }
        catch (ex) { }
        OSF._OfficeAppFactory.bootstrap(function () { }, function (e) {
            if (e instanceof Error) {
                console.warn(e.message);
            }
            else {
                console.warn(JSON.stringify(e));
            }
        });
        function funcToRunWhenContentLoaded() {
            OSFPerformance.hostSpecificFileName = OSF.LoadScriptHelper.getHostBundleJsName();
            Office.onReadyInternal(function () {
                OSFPerfUtil.sendPerformanceTelemetry();
            });
            if (OSF._OfficeAppFactory.getHostInfo().hostLocale) {
                setTimeout(function () {
                    OSF.OUtil.ensureOfficeStringsJs().catch(function (ex) {
                        console.error(ex);
                    });
                }, 0);
            }
        }
        if (document.readyState === "complete"
            || document.readyState === "interactive") {
            funcToRunWhenContentLoaded();
        }
        else {
            window.addEventListener('DOMContentLoaded', function (event) {
                funcToRunWhenContentLoaded();
            });
        }
    }
})(OSF || (OSF = {}));
OSFPerformance.hostInitializationEnd = OSFPerformance.now();


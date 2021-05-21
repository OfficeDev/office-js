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
        function OfficeAppContext(id, appName, appVersion, appUILocale, dataLocale, docUrl, clientMode, settingsFunc, reason, osfControlType, eToken, correlationId, appInstanceId, touchEnabled, commerceAllowed, appMinorVersion, requirementMatrix, hostCustomMessage, hostFullVersion, clientWindowHeight, clientWindowWidth, addinName, appDomains, dialogRequirementMatrix, featureGates, officeThemeFunc, initialDisplayMode, isFromWacAutomation) {
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
            var returnedContext = new OSF.OfficeAppContext(id, appType, version, UILocale, dataLocale, docUrl, clientMode, settingsFunc, reason, osfControlType, eToken, correlationId, appInstanceId, touchEnabled, commerceAllowed, minorVersion, requirementMatrix, hostCustomMessage, hostFullVersion, undefined, undefined, undefined, undefined, dialogRequirementMatrix, sdxFeatureGates, officeThemeFunc, initialDisplayMode, undefined);
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
                    var returnedContext = new OSF.OfficeAppContext(appContext._id, appContext._appName, appContext._appVersion, appContext._appUILocale, appContext._dataLocale, appContext._docUrl, appContext._clientMode, settingsFunc, appContext._reason, appContext._osfControlType, appContext._eToken, appContext._correlationId, appInstanceId, touchEnabled, commerceAllowed, minorVersion, requirementMatrix, appContext._hostCustomMessage, appContext._hostFullVersion, appContext._clientWindowHeight, appContext._clientWindowWidth, appContext._addinName, appContext._appDomains, appContext._dialogRequirementMatrix, appContext._featureGates, undefined, appContext._initialDisplayMode, appContext._isFromWacAutomation);
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
                returnedContext = new OSF.OfficeAppContext(id, appType, version, UILocale, dataLocale, docUrl, clientMode, settingsFunc, reason, osfControlType, eToken, correlationId, appInstanceId, touchEnabled, commerceAllowed, minorVersion, requirementMatrix, hostCustomMessage, hostFullVersion, undefined, undefined, undefined, undefined, undefined, undefined, undefined, initialDisplayMode, undefined);
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
                returnedContext = new OSF.OfficeAppContext(id, appType, version, UILocale, dataLocale, docUrl, clientMode, settingsFunc, reason, osfControlType, eToken, correlationId, appInstanceId, touchEnabled, commerceAllowed, minorVersion, requirementMatrix, hostCustomMessage, hostFullVersion, undefined, undefined, undefined, undefined, undefined, undefined, undefined, initialDisplayMode, undefined);
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
                        httpStatusCode: xhr.status,
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
        HttpUtility.fetchSendRequestFunc = function (request) {
            var requestBodyText = CoreUtility._getRequestBodyText(request);
            if (requestBodyText === '') {
                requestBodyText = undefined;
            }
            return fetch(request.url, {
                method: request.method,
                headers: request.headers,
                body: requestBodyText
            })
                .then(function (resp) {
                return resp.text()
                    .then(function (body) {
                    var statusCode = resp.status;
                    var headers = {};
                    resp.headers.forEach(function (value, name) {
                        headers[name] = value;
                    });
                    var ret = { statusCode: statusCode, headers: headers, body: body };
                    return ret;
                });
            });
        };
        HttpUtility.sendRequest = function (request) {
            HttpUtility.validateAndNormalizeRequest(request);
            var func = HttpUtility.s_customSendRequestFunc;
            if (!func) {
                if (typeof (fetch) !== 'undefined') {
                    func = HttpUtility.fetchSendRequestFunc;
                }
                else {
                    func = HttpUtility.xhrSendRequestFunc;
                }
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
                    _this.httpStatusCode = error.httpStatusCode;
                    _this.data = error.data;
                }
                if (CoreUtility.isNullOrUndefined(_this.httpStatusCode) || _this.httpStatusCode === 200) {
                    var mapping = {};
                    mapping[CoreErrorCodes.accessDenied] = 401;
                    mapping[CoreErrorCodes.connectionFailure] = 500;
                    mapping[CoreErrorCodes.generalException] = 500;
                    mapping[CoreErrorCodes.invalidArgument] = 400;
                    mapping[CoreErrorCodes.invalidObjectPath] = 400;
                    mapping[CoreErrorCodes.invalidOrTimedOutSession] = 408;
                    mapping[CoreErrorCodes.invalidRequestContext] = 400;
                    mapping[CoreErrorCodes.timeout] = 408;
                    mapping[CoreErrorCodes.valueNotLoaded] = 400;
                    _this.httpStatusCode = mapping[_this.code];
                }
                if (CoreUtility.isNullOrUndefined(_this.httpStatusCode)) {
                    _this.httpStatusCode = 500;
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
                    httpStatusCode: 400,
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
        CoreErrorCodes.requestPayloadSizeLimitExceeded = 'RequestPayloadSizeLimitExceeded';
        CoreErrorCodes.writeNotSupportedWhenModalDialogOpen = 'WriteNotSupportedWhenModalDialogOpen';
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
        CoreResourceStrings.invalidSheetName = 'InvalidSheetName';
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
            var response = { HttpStatusCode: 200, ErrorCode: '', ErrorMessage: '', Headers: null, Body: null };
            response.Body = JSON.parse(responseBody);
            response.Headers = responseHeaders;
            return response;
        };
        RichApiMessageUtility.buildResponseOnError = function (errorCode, message) {
            var response = { HttpStatusCode: 500, ErrorCode: '', ErrorMessage: '', Headers: null, Body: null };
            response.ErrorCode = CoreErrorCodes.generalException;
            response.ErrorMessage = message;
            if (errorCode == RichApiMessageUtility.OfficeJsErrorCode_ooeNoCapability) {
                response.ErrorCode = CoreErrorCodes.accessDenied;
                response.HttpStatusCode = 401;
            }
            else if (errorCode == RichApiMessageUtility.OfficeJsErrorCode_ooeActivityLimitReached) {
                response.ErrorCode = CoreErrorCodes.activityLimitReached;
                response.HttpStatusCode = 429;
            }
            else if (errorCode == RichApiMessageUtility.OfficeJsErrorCode_ooeInvalidOrTimedOutSession) {
                response.ErrorCode = CoreErrorCodes.invalidOrTimedOutSession;
                response.HttpStatusCode = 408;
                response.ErrorMessage = CoreUtility._getResourceString(CoreResourceStrings.invalidOrTimedOutSessionMessage);
            }
            else if (errorCode == RichApiMessageUtility.OfficeJsErrorCode_ooeRequestPayloadSizeLimitExceeded) {
                response.ErrorCode = CoreErrorCodes.requestPayloadSizeLimitExceeded;
                response.HttpStatusCode = 400;
            }
            else if (errorCode == RichApiMessageUtility.OfficeJsErrorCode_ooeWriteNotSupportedWhenModalDialogOpen) {
                response.ErrorCode = CoreErrorCodes.writeNotSupportedWhenModalDialogOpen;
                response.HttpStatusCode = 400;
            }
            else if (errorCode == RichApiMessageUtility.OfficeJsErrorCode_ooeInvalidSheetName) {
                response.ErrorCode = CoreErrorCodes.invalidRequestContext;
                response.HttpStatusCode = 400;
                response.ErrorMessage = CoreUtility._getResourceString(CoreResourceStrings.invalidSheetName);
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
        RichApiMessageUtility.OfficeJsErrorCode_ooeRequestPayloadSizeLimitExceeded = 5103;
        RichApiMessageUtility.OfficeJsErrorCode_ooeWriteNotSupportedWhenModalDialogOpen = 5016;
        RichApiMessageUtility.OfficeJsErrorCode_ooeInvalidSheetName = 1014;
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
            var statusCode = responseInfo.statusCode.toString();
            if (CoreUtility.isNullOrUndefined(errorObj) || typeof errorObj !== 'object' || !errorObj.error) {
                return CoreUtility._createDefaultErrorResponse(statusCode);
            }
            var error = errorObj.error;
            var innerError = error.innerError;
            if (innerError && innerError.code) {
                return CoreUtility._createErrorResponse(innerError.code, statusCode, innerError.message);
            }
            if (error.code) {
                return CoreUtility._createErrorResponse(error.code, statusCode, error.message);
            }
            return CoreUtility._createDefaultErrorResponse(statusCode);
        };
        CoreUtility._createDefaultErrorResponse = function (statusCode) {
            return {
                errorCode: CoreErrorCodes.connectionFailure,
                errorMessage: CoreUtility._getResourceString(CoreResourceStrings.connectionFailureWithStatus, statusCode)
            };
        };
        CoreUtility._createErrorResponse = function (code, statusCode, message) {
            return {
                errorCode: code,
                errorMessage: CoreUtility._getResourceString(CoreResourceStrings.connectionFailureWithDetails, [
                    statusCode,
                    code,
                    message
                ])
            };
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
        enableUndoableFlag: true,
        appendTypeNameToObjectPathInfo: false
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
                                httpStatusCode: 400,
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
                            httpStatusCode: 400,
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
                    httpStatusCode: 400,
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
                    httpStatusCode: 400,
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
                        httpStatusCode: 400,
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
                    httpStatusCode: response.HttpStatusCode,
                    message: response.ErrorMessage
                });
            }
            if (response.Body && response.Body.Error) {
                return new _Internal.RuntimeError({
                    code: response.Body.Error.Code,
                    httpStatusCode: response.Body.Error.HttpStatusCode,
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
                        HttpStatusCode: responseInfo.statusCode,
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
                        HttpStatusCode: responseInfo.statusCode,
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
                        httpStatusCode: 400,
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
                                httpStatusCode: 400,
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
                    httpStatusCode: 400,
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
                    httpStatusCode: 404,
                    message: message,
                    debugInfo: { errorLocation: apiFullName }
                });
            }
        };
        CommonUtility.calculateApiFlags = function (apiFlags, undoableApiSetName, undoableApiSetVersion) {
            if (!CommonUtility.isSetSupported(undoableApiSetName, undoableApiSetVersion)) {
                apiFlags = apiFlags & (~2);
            }
            return apiFlags;
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
                if (OfficeExtension_1._internalConfig.appendTypeNameToObjectPathInfo && _this._objectPath.objectPathInfo && _this._className) {
                    _this._objectPath.objectPathInfo.T = _this._className;
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
                                httpStatusCode: 400,
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
                                httpStatusCode: 400,
                                message: CoreUtility._getResourceString(CommonResourceStrings.propertyDoesNotExist, prop),
                                debugInfo: {
                                    errorLocation: prop
                                }
                            });
                        }
                        if (throwOnReadOnly && !propertyDescriptor.set) {
                            throw new _Internal.RuntimeError({
                                code: CoreErrorCodes.invalidArgument,
                                httpStatusCode: 400,
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
                    httpStatusCode: 400,
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
                        HttpStatusCode: responseInfo.statusCode,
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
                        HttpStatusCode: responseInfo.statusCode,
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
                            httpStatusCode: 400,
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
                    httpStatusCode: response.HttpStatusCode,
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
                    httpStatusCode: response.Body.Error.HttpStatusCode,
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
                        httpStatusCode: 400,
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
        Constants.officeScriptFireRecordingEvent = 'X-OfficeScriptFireRecordingEvent';
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
                        return OSF.DDA.RichApi.richApiMessageManager.register(handler);
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
        InvalidSheetName: 'The request cannot be processed because the specified worksheet cannot be found. Please try again.',
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
                httpStatusCode: 400,
                message: CoreUtility._getResourceString(resourceId, arg),
                debugInfo: errorLocation ? { errorLocation: errorLocation } : undefined
            });
        };
        Utility.createRuntimeError = function (code, message, location, httpStatusCode, data) {
            return new _Internal.RuntimeError({
                code: code,
                httpStatusCode: httpStatusCode,
                message: message,
                debugInfo: { errorLocation: location },
                data: data
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
                httpStatusCode: 400,
                message: CoreUtility._getResourceString(ResourceStrings.propertyNotLoaded, propertyName),
                debugInfo: entityName ? { errorLocation: entityName + '.' + propertyName } : undefined
            });
        };
        Utility.createCannotUpdateReadOnlyPropertyException = function (entityName, propertyName) {
            return new _Internal.RuntimeError({
                code: ErrorCodes.cannotUpdateReadOnlyProperty,
                httpStatusCode: 400,
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
                if (OfficeExtension_1._internalConfig.appendTypeNameToObjectPathInfo) {
                    if (this._objectPath && this._objectPath.objectPathInfo && this._className) {
                        this._objectPath.objectPathInfo.T = this._className;
                    }
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
        LibraryBuilder.prototype.evaluateSimpleExpression = function (expression, thisObj) {
            if (Utility.isNullOrUndefined(expression)) {
                return null;
            }
            var paths = expression.split('.');
            if (paths.length === 3 && paths[0] === 'OfficeExtension' && paths[1] === 'Constants') {
                return Constants[paths[2]];
            }
            if (paths[0] === 'this') {
                var obj = thisObj;
                for (var i = 1; i < paths.length; i++) {
                    if (paths[i] == 'toString()') {
                        obj = obj.toString();
                    }
                    else if (paths[i].substr(paths[i].length - 2) === "()") {
                        obj = obj[paths[i].substr(0, paths[i].length - 2)]();
                    }
                    else {
                        obj = obj[paths[i]];
                    }
                }
                return obj;
            }
            throw new OfficeExtension_1.Error("Cannot evaluate: " + expression);
        };
        LibraryBuilder.prototype.evaluateEventTargetId = function (targetIdExpression, thisObj) {
            if (Utility.isNullOrEmptyString(targetIdExpression)) {
                return '';
            }
            return this.evaluateSimpleExpression(targetIdExpression, thisObj);
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
                        this.ensureArraySize(elem, 6);
                        typeInfo.scalarProperties[i] = {
                            name: this.getString(elem[0]),
                            behaviorFlags: elem[1],
                            apiSetInfoOrdinal: elem[2],
                            originalName: this.getString(elem[3]),
                            setMethodApiFlags: elem[4],
                            undoableApiSetInfoOrdinal: elem[5],
                        };
                    }
                    this.buildScalarProperty(type, typeInfo, typeInfo.scalarProperties[i]);
                }
            }
        };
        LibraryBuilder.prototype.calculateApiFlags = function (apiFlags, undoableApiSetInfoOrdinal) {
            if (undoableApiSetInfoOrdinal > 0) {
                var undoableApiSetInfo = this.m_apiSets[undoableApiSetInfoOrdinal - 1];
                if (undoableApiSetInfo) {
                    apiFlags = CommonUtility.calculateApiFlags(apiFlags, undoableApiSetInfo.name, undoableApiSetInfo.version);
                }
            }
            return apiFlags;
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
                    var apiFlags = thisBuilder.calculateApiFlags(propInfo.setMethodApiFlags, propInfo.undoableApiSetInfoOrdinal);
                    BatchApiHelper.invokeSetProperty(this, propInfo.originalName, value, apiFlags);
                };
            }
            Object.defineProperty(type.prototype, propInfo.name, descriptor);
        };
        LibraryBuilder.prototype.buildNavigationProperties = function (type, typeInfo) {
            if (Array.isArray(typeInfo.navigationProperties)) {
                for (var i = 0; i < typeInfo.navigationProperties.length; i++) {
                    var elem = typeInfo.navigationProperties[i];
                    if (Array.isArray(elem)) {
                        this.ensureArraySize(elem, 8);
                        typeInfo.navigationProperties[i] = {
                            name: this.getString(elem[0]),
                            behaviorFlags: elem[2],
                            apiSetInfoOrdinal: elem[3],
                            originalName: this.getString(elem[4]),
                            getMethodApiFlags: elem[5],
                            setMethodApiFlags: elem[6],
                            propertyTypeFullName: this.getString(elem[1]),
                            undoableApiSetInfoOrdinal: elem[7]
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
                    var apiFlags = thisBuilder.calculateApiFlags(propInfo.setMethodApiFlags, propInfo.undoableApiSetInfoOrdinal);
                    BatchApiHelper.invokeSetProperty(this, propInfo.originalName, value, apiFlags);
                };
            }
            Object.defineProperty(type.prototype, propInfo.name, descriptor);
        };
        LibraryBuilder.prototype.buildScalarMethods = function (type, typeInfo) {
            if (Array.isArray(typeInfo.scalarMethods)) {
                for (var i = 0; i < typeInfo.scalarMethods.length; i++) {
                    var elem = typeInfo.scalarMethods[i];
                    if (Array.isArray(elem)) {
                        this.ensureArraySize(elem, 7);
                        typeInfo.scalarMethods[i] = {
                            name: this.getString(elem[0]),
                            behaviorFlags: elem[2],
                            apiSetInfoOrdinal: elem[3],
                            originalName: this.getString(elem[5]),
                            apiFlags: elem[4],
                            parameterCount: elem[1],
                            undoableApiSetInfoOrdinal: elem[6]
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
                var apiFlags = thisBuilder.calculateApiFlags(methodInfo.apiFlags, methodInfo.undoableApiSetInfoOrdinal);
                return BatchApiHelper.invokeMethod(this, methodInfo.originalName, operationType, args, apiFlags, resultProcessType);
            };
        };
        LibraryBuilder.prototype.buildNavigationMethods = function (type, typeInfo) {
            if (Array.isArray(typeInfo.navigationMethods)) {
                for (var i = 0; i < typeInfo.navigationMethods.length; i++) {
                    var elem = typeInfo.navigationMethods[i];
                    if (Array.isArray(elem)) {
                        this.ensureArraySize(elem, 9);
                        typeInfo.navigationMethods[i] = {
                            name: this.getString(elem[0]),
                            behaviorFlags: elem[3],
                            apiSetInfoOrdinal: elem[4],
                            originalName: this.getString(elem[6]),
                            apiFlags: elem[5],
                            parameterCount: elem[2],
                            returnTypeFullName: this.getString(elem[1]),
                            returnObjectGetByIdMethodName: this.getString(elem[7]),
                            undoableApiSetInfoOrdinal: elem[8]
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
                    var apiFlags = thisBuilder.calculateApiFlags(methodInfo.apiFlags, methodInfo.undoableApiSetInfoOrdinal);
                    return BatchApiHelper.createMethodObject(thisBuilder.getFunction(methodInfo.returnTypeFullName), this, methodInfo.originalName, operationType, args, (methodInfo.behaviorFlags & 4) !== 0, (methodInfo.behaviorFlags & 8) !== 0, methodInfo.returnObjectGetByIdMethodName, apiFlags);
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
try {
    if (typeof OfficePlatformGlobal !== 'undefined' && typeof OfficePlatformGlobal.ReactNativeReka !== 'undefined') {
        var RichApiBridgeImpl = (function () {
            function RichApiBridgeImpl() {
            }
            RichApiBridgeImpl.prototype.sendMessageToHost = function (value) {
                OfficePlatformGlobal.ReactNativeReka.RekaInstance.InvokeNative('RichApiDataService', 'ProcessMessageToHost', value);
            };
            return RichApiBridgeImpl;
        }());
        var RichApiBridge_1 = new RichApiBridgeImpl();
        var RichApiDataService_1 = (function () {
            function RichApiDataService() {
            }
            RichApiDataService.prototype.Invoke = function (methodName, value) {
                if (methodName === 'ProcessMessageFromHost') {
                    if (RichApiBridge_1.onMessageFromHost) {
                        RichApiBridge_1.onMessageFromHost(value);
                    }
                }
            };
            return RichApiDataService;
        }());
        OfficePlatformGlobal.ReactNativeReka.initReka();
        OfficePlatformGlobal.ReactNativeReka.RekaInstance.RegisterService('RichApiDataService', function () { return new RichApiDataService_1(); });
        OfficeExtension.HostBridge.init(RichApiBridge_1);
        OfficeExtension.HostBridgeSession.getInstanceIfHostBridgeInited();
    }
}
catch (e) {
}
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
var OfficeRuntime;
(function (OfficeRuntime) {
    var ApiHelper;
    (function (ApiHelper) {
        function isSharedRuntime() {
            try {
                if (typeof (OSF) !== "undefined" &&
                    typeof (OSF._OfficeAppFactory) !== "undefined" &&
                    typeof (OSF._OfficeAppFactory.getHostInfo) !== "undefined") {
                    var flags = OSF._OfficeAppFactory.getHostInfo().flags;
                    if (flags & OSF.HostInfoFlags.SharedApp) {
                        return true;
                    }
                }
            }
            catch (e) { }
            return false;
        }
        function isReactNative() {
            try {
                if (typeof (OSF) !== "undefined" &&
                    typeof (OSF.isOfficeReactNative) !== "undefined" &&
                    OSF.isOfficeReactNative()) {
                    return true;
                }
            }
            catch (e) { }
            return false;
        }
        function getHostPlatform() {
            var hostPlatform = 'web';
            try {
                if (typeof (OSF) !== "undefined" &&
                    typeof (OSF._OfficeAppFactory) !== "undefined" &&
                    typeof (OSF._OfficeAppFactory.getHostInfo) !== "undefined") {
                    hostPlatform = OSF._OfficeAppFactory.getHostInfo().hostPlatform;
                }
            }
            catch (e) { }
            return hostPlatform;
        }
        function isNativeStorage() {
            var hostPlatform = getHostPlatform();
            if (!isSharedRuntime() &&
                (isReactNative() ||
                    hostPlatform === 'win32' ||
                    (hostPlatform === 'mac' && window._useSDXSandbox))) {
                return true;
            }
            return false;
        }
        ApiHelper.isNativeStorage = isNativeStorage;
        function isNativeApiInfomation() {
            var hostPlatform = getHostPlatform();
            if (isReactNative() ||
                hostPlatform === 'win32' ||
                (hostPlatform === 'mac' && window._useSDXSandbox)) {
                return true;
            }
            return false;
        }
        ApiHelper.isNativeApiInfomation = isNativeApiInfomation;
        function isNativeDialog() {
            var hostPlatform = getHostPlatform();
            if (isReactNative() || hostPlatform === 'win32') {
                return true;
            }
            return false;
        }
        ApiHelper.isNativeDialog = isNativeDialog;
    })(ApiHelper = OfficeRuntime.ApiHelper || (OfficeRuntime.ApiHelper = {}));
})(OfficeRuntime || (OfficeRuntime = {}));
var OfficeRuntime;
(function (OfficeRuntime) {
    if (OfficeRuntime.ApiHelper.isNativeApiInfomation()) {
        OfficeRuntime.apiInformation = apiInformationNative;
    }
    else {
        OfficeRuntime.apiInformation = apiInformationWeb;
    }
    var apiInformationNative = {
        isSetSupported: function (capability, version) {
            if (typeof capability != "string") {
                return false;
            }
            if (version == undefined) {
                version = "0.0.0";
            }
            if (global.__apiSets === undefined) {
                return false;
            }
            else {
                var sets = global.__apiSets;
                var index = Object.keys(sets).map(function (key) { return key.toLowerCase(); }).indexOf(capability.toLowerCase());
                if (index > -1) {
                    var setMaxVersion = sets[Object.keys(sets)[index]];
                    try {
                        var minVersionNum = getVersion(version);
                        var setMaxVersionNum = void 0;
                        if (Array.isArray(setMaxVersion)) {
                            setMaxVersionNum = getMatcheddMajorVersion(setMaxVersion, minVersionNum.major);
                        }
                        else {
                            setMaxVersionNum = getVersion(setMaxVersion);
                        }
                        if (setMaxVersionNum.major > 0) {
                            if (setMaxVersionNum.major > minVersionNum.major) {
                                return true;
                            }
                            if (setMaxVersionNum.major == minVersionNum.major &&
                                setMaxVersionNum.minor > minVersionNum.minor) {
                                return true;
                            }
                            if (setMaxVersionNum.major == minVersionNum.major &&
                                setMaxVersionNum.minor == minVersionNum.minor &&
                                setMaxVersionNum.patch >= minVersionNum.patch) {
                                return true;
                            }
                        }
                    }
                    catch (e) {
                        return false;
                    }
                }
                return false;
            }
        }
    };
    var apiInformationWeb = {
        isSetSupported: function (capability, version) {
            try {
                if (Office && Office.context && Office.context.requirements) {
                    return Office.context.requirements
                        .isSetSupported(capability, Number(version));
                }
                else {
                    return false;
                }
            }
            catch (e) {
                return false;
            }
        }
    };
    function getMatcheddMajorVersion(versions, majorVersion) {
        for (var i = 0; i < versions.length; i++) {
            var v = getVersion(versions[i]);
            if (v.major == majorVersion)
                return v;
        }
        return { major: 0, minor: 0, patch: 0 };
    }
    function getVersion(version) {
        var temp = version.split(".");
        var major = 0;
        var minor = 0;
        var patch = 0;
        if (temp.length < 2 && isNaN(Number(version))) {
            throw "version format incorrect";
        }
        else {
            major = Number(temp[0]);
            if (temp.length >= 2) {
                minor = Number(temp[1]);
            }
            if (temp.length >= 3) {
                patch = Number(temp[2]);
            }
            if (isNaN(major) || isNaN(minor) || isNaN(patch)) {
                throw "version format incorrect";
            }
        }
        var result = { "major": major, "minor": minor, "patch": patch };
        return result;
    }
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
var OfficeRuntime;
(function (OfficeRuntime) {
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
    var _calculateApiFlags = OfficeExtension.CommonUtility.calculateApiFlags;
    var _typeDialogService = "DialogService";
    var Dialog = (function () {
        function Dialog(dialogService) {
            this._dialogService = dialogService;
        }
        return Dialog;
    }());
    OfficeRuntime.Dialog = Dialog;
    var WebDialog = (function (_super) {
        __extends(WebDialog, _super);
        function WebDialog() {
            return _super !== null && _super.apply(this, arguments) || this;
        }
        WebDialog.prototype.close = function () {
            this._dialogService.close();
            return OfficeExtension.CoreUtility.Promise.resolve();
        };
        return WebDialog;
    }(Dialog));
    var NativeDialog = (function (_super) {
        __extends(NativeDialog, _super);
        function NativeDialog() {
            return _super !== null && _super.apply(this, arguments) || this;
        }
        NativeDialog.prototype.close = function () {
            this._dialogService.close();
            return this._dialogService.context.sync();
        };
        return NativeDialog;
    }(Dialog));
    function displayWebDialog(url, options) {
        if (options === void 0) { options = {}; }
        return new OfficeExtension.CoreUtility.Promise(function (resolve, reject) {
            if (options.width && options.height && (!isInt(options.width) || !isInt(options.height))) {
                throw new OfficeExtension.Error({ code: "InvalidArgument", message: 'Dimensions must be "number%" or number.' });
            }
            var dialogOptions = {
                width: options.width ? parseInt(options.width) : 50,
                height: options.height ? parseInt(options.height) : 50,
                displayInIFrame: options.displayInIFrame
            };
            if (OfficeRuntime.ApiHelper.isNativeDialog()) {
                displayDialogNative(dialogOptions, resolve, reject, url, options);
            }
            else {
                displayDialogWeb(dialogOptions, resolve, reject, url, options);
            }
        });
    }
    OfficeRuntime.displayWebDialog = displayWebDialog;
    function displayDialogWeb(dialogOptions, resolve, reject, url, options) {
        if (options === void 0) { options = {}; }
        var dialog;
        Office.context.ui.displayDialogAsync(url, dialogOptions, dialogCallback);
        function dialogCallback(asyncResult) {
            if (asyncResult.status === 'failed') {
                reject(new OfficeExtension.Error(lookupErrorCodeAndMessage(asyncResult.error.code)));
            }
            else {
                dialog = asyncResult.value;
                dialog.addEventHandler(Office.EventType.DialogMessageReceived, messageHandler);
                dialog.addEventHandler(Office.EventType.DialogEventReceived, eventHandler);
                resolve(new WebDialog(dialog));
            }
        }
        function messageHandler(args) {
            if (options.onMessage) {
                options.onMessage(args.message, dialog);
            }
        }
        function eventHandler(args) {
            if (args.error === 12006) {
                if (options.onClose) {
                    options.onClose();
                }
            }
            else {
                if (options.onRuntimeError) {
                    options.onRuntimeError(new OfficeExtension.Error(lookupErrorCodeAndMessage(args.error)), dialog);
                }
            }
        }
    }
    function displayDialogNative(dialogOptions, resolve, reject, url, options) {
        if (options === void 0) { options = {}; }
        var ctx = new OfficeExtension.ClientRequestContext();
        var dialogService = DialogService.newObject(ctx);
        var dialog = new NativeDialog(dialogService);
        var eventResult = dialogService.onDialogMessage.add(function (args) {
            OfficeExtension.Utility.log("dialogMessageHandler:" + JSON.stringify(args));
            switch (args.type) {
                case 17:
                    if (args.error) {
                        reject(args.error);
                    }
                    else {
                        resolve(dialog);
                    }
                    break;
                case 12:
                    if (options.onMessage) {
                        options.onMessage(args.message, dialog);
                    }
                    break;
                case 10:
                default:
                    if (args.originalErrorCode === 12006) {
                        if (eventResult) {
                            eventResult.remove();
                            ctx.sync();
                        }
                        if (options.onClose) {
                            options.onClose();
                        }
                    }
                    else {
                        if (options.onRuntimeError) {
                            options.onRuntimeError(args.error, dialog);
                            reject(args.error);
                        }
                    }
            }
            return OfficeExtension.CoreUtility.Promise.resolve();
        });
        return ctx.sync()
            .then(function () {
            dialogService.displayDialog(url, dialogOptions);
            return ctx.sync();
        })
            .catch(function (e) {
            reject(e);
        });
    }
    function isInt(value) {
        return (/^(\-|\+)?([0-9]+)%?$/.test(value));
    }
    function lookupErrorCodeAndMessage(internalCode) {
        var _a;
        var table = (_a = {},
            _a[12002] = { code: "InvalidUrl", message: "Cannot load URL, no such page or bad URL syntax." },
            _a[12003] = { code: "InvalidUrl", message: "HTTPS is required." },
            _a[12004] = { code: "Untrusted", message: "Domain is not trusted." },
            _a[12005] = { code: "InvalidUrl", message: "HTTPS is required." },
            _a[12007] = { code: "FailedToOpen", message: "Another dialog is already opened." },
            _a);
        if (table[internalCode]) {
            return table[internalCode];
        }
        else {
            return { code: "Unknown", message: "An unknown error has occured with code: " + internalCode };
        }
    }
    var DialogService = (function (_super) {
        __extends(DialogService, _super);
        function DialogService() {
            return _super !== null && _super.apply(this, arguments) || this;
        }
        Object.defineProperty(DialogService.prototype, "_className", {
            get: function () {
                return "DialogService";
            },
            enumerable: true,
            configurable: true
        });
        DialogService.prototype.close = function () {
            _invokeMethod(this, "Close", 1, [], 4, 0);
        };
        DialogService.prototype.displayDialog = function (url, options) {
            _invokeMethod(this, "DisplayDialog", 1, [url, options], 4, 0);
        };
        DialogService.prototype._handleResult = function (value) {
            _super.prototype._handleResult.call(this, value);
            if (_isNullOrUndefined(value))
                return;
            var obj = value;
            _fixObjectPathIfNecessary(this, obj);
        };
        DialogService.prototype._handleRetrieveResult = function (value, result) {
            _super.prototype._handleRetrieveResult.call(this, value, result);
            _processRetrieveResult(this, value, result);
        };
        DialogService.newObject = function (context) {
            return _createTopLevelServiceObject(OfficeRuntime.DialogService, context, "Microsoft.Dialog.DialogService", false, 4);
        };
        Object.defineProperty(DialogService.prototype, "onDialogMessage", {
            get: function () {
                if (!this.m_dialogMessage) {
                    this.m_dialogMessage = new OfficeExtension.GenericEventHandlers(this.context, this, "DialogMessage", {
                        eventType: 65536,
                        registerFunc: function () { return void {}; },
                        unregisterFunc: function () { return void {}; },
                        getTargetIdFunc: function () { return null; },
                        eventArgsTransformFunc: function (args) {
                            var transformedArgs;
                            try {
                                var parsedMessage = JSON.parse(args.message);
                                var error = parsedMessage.errorCode
                                    ? new OfficeExtension.Error(lookupErrorCodeAndMessage(parsedMessage.errorCode))
                                    : null;
                                transformedArgs = {
                                    originalErrorCode: parsedMessage.errorCode,
                                    type: parsedMessage.type,
                                    error: error,
                                    message: parsedMessage.message
                                };
                            }
                            catch (e) {
                                transformedArgs = {
                                    originalErrorCode: null,
                                    type: 17,
                                    error: new OfficeExtension.Error({ code: "GenericException", message: "Unknown error" }),
                                    message: e.message
                                };
                            }
                            return OfficeExtension.Utility._createPromiseFromResult(transformedArgs);
                        }
                    });
                }
                return this.m_dialogMessage;
            },
            enumerable: true,
            configurable: true
        });
        DialogService.prototype.toJSON = function () {
            return _toJson(this, {}, {});
        };
        return DialogService;
    }(OfficeExtension.ClientObject));
    OfficeRuntime.DialogService = DialogService;
    var DialogEventType;
    (function (DialogEventType) {
        DialogEventType[DialogEventType["dialogMessageReceived"] = 0] = "dialogMessageReceived";
        DialogEventType[DialogEventType["dialogEventReceived"] = 1] = "dialogEventReceived";
    })(DialogEventType || (DialogEventType = {}));
    var DialogErrorCodes;
    (function (DialogErrorCodes) {
        DialogErrorCodes["generalException"] = "GeneralException";
    })(DialogErrorCodes = OfficeRuntime.DialogErrorCodes || (OfficeRuntime.DialogErrorCodes = {}));
})(OfficeRuntime || (OfficeRuntime = {}));
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
var OfficeRuntime;
(function (OfficeRuntime) {
    var message;
    (function (message_1) {
        var listeners = {};
        function on(eventName, listener) {
            add(eventName, listener);
            return new Promise(function (resolve) {
                resolve();
            });
        }
        message_1.on = on;
        function off(eventName, listener) {
            remove(eventName, listener);
            return new Promise(function (resolve) {
                resolve();
            });
        }
        message_1.off = off;
        function emit(eventName, message) {
            send(eventName, message);
            return new Promise(function (resolve) {
                resolve();
            });
        }
        message_1.emit = emit;
        function setupReceive() {
            if (Office && Office.context && Office.context.messaging && !Office.context.messaging.onMessage) {
                Office.context.messaging.onMessage = receiveMessage.bind(this);
            }
        }
        function add(eventName, listener) {
            if (!listeners.hasOwnProperty(eventName)) {
                listeners[eventName] = new ListenerManager();
                setupReceive();
            }
            listeners[eventName].add(listener);
        }
        function remove(eventName, listener) {
            if (listeners.hasOwnProperty(eventName)) {
                if (listener) {
                    listeners[eventName].remove(listener);
                }
                else {
                    delete listeners[eventName];
                }
            }
        }
        function send(eventName, message) {
            var wrapped = {
                eventName: eventName,
                message: message
            };
            if (Office && Office.context && Office.context.messaging && Office.context.messaging.sendMessage) {
                Office.context.messaging.sendMessage(wrapped);
            }
        }
        function receiveMessage(wrapped) {
            var eventName = wrapped.eventName;
            var message = wrapped.message;
            if (listeners.hasOwnProperty(eventName)) {
                listeners[eventName].call(message);
            }
        }
    })(message = OfficeRuntime.message || (OfficeRuntime.message = {}));
    var ListenerManager = (function () {
        function ListenerManager() {
            this.listeners = [];
        }
        ListenerManager.prototype.add = function (listener) {
            this.listeners.push(listener);
        };
        ListenerManager.prototype.remove = function (listener) {
            var index = this.listeners.lastIndexOf(listener);
            if (index !== -1) {
                this.listeners.splice(index, 1);
            }
        };
        ListenerManager.prototype.call = function (message) {
            this.listeners.forEach(function (listener) { return listener(message); });
        };
        return ListenerManager;
    }());
    OfficeRuntime.ListenerManager = ListenerManager;
})(OfficeRuntime || (OfficeRuntime = {}));
var OfficeRuntime;
(function (OfficeRuntime) {
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
    var _calculateApiFlags = OfficeExtension.CommonUtility.calculateApiFlags;
    var _typePersistentKvStorageService = "PersistentKvStorageService";
    var prefix = '_Office_AsyncStorage_';
    var dummyUnusedKey = prefix + '|_unusedKey_';
    var storageNative = {
        getItem: function (key) {
            return callPersistentKvStorageManager(function (perStorage, invokeId) { return perStorage.multiGet(invokeId, JSON.stringify([key])); }, function (result) {
                var parsedResult = JSON.parse(result);
                return parsedResult && parsedResult[0] && parsedResult[0][1] ? parsedResult[0][1] : null;
            });
        },
        setItem: function (key, value) {
            return callPersistentKvStorageManager(function (perStorage, invokeId) { return perStorage.multiSet(invokeId, JSON.stringify([[key, value]])); }, function () { return null; });
        },
        removeItem: function (key) {
            return callPersistentKvStorageManager(function (perStorage, invokeId) { return perStorage.multiRemove(invokeId, JSON.stringify([key])); }, function () { return null; });
        },
        getItems: function (keys) {
            return callPersistentKvStorageManager(function (perStorage, invokeId) { return perStorage.multiGet(invokeId, JSON.stringify(keys)); }, function (result) {
                var keyValues = JSON.parse(result);
                var map = {};
                keyValues && keyValues.forEach(function (_a) {
                    var key = _a[0], value = _a[1];
                    map[key] = value;
                    return value;
                });
                keys && keys.forEach(function (key) { map[key] ? map[key] : null; });
                return map;
            });
        },
        setItems: function (keyValues) {
            var keyValuePairs = [];
            for (var key in keyValues) {
                if (keyValues.hasOwnProperty(key)) {
                    keyValuePairs.push([key, keyValues[key]]);
                }
            }
            return callPersistentKvStorageManager(function (storage, invokeId) { return storage.multiSet(invokeId, JSON.stringify(keyValuePairs)); }, function () { return null; });
        },
        removeItems: function (keys) {
            return callPersistentKvStorageManager(function (perStorage, invokeId) { return perStorage.multiRemove(invokeId, JSON.stringify(keys)); }, function () { return null; });
        },
        getKeys: function () {
            return callPersistentKvStorageManager(function (perStorage, invokeId) { return perStorage.getAllKeys(invokeId); }, function (result) { return JSON.parse(result); });
        }
    };
    var storageWeb = {
        getItem: function (key) {
            return performActionAndReturnResult(function () { return localStorageGetItem(key); });
        },
        setItem: function (key, value) {
            return performAction(function () { return localStorageSetItem(key, value); });
        },
        removeItem: function (key) {
            return performAction(function () { return window.localStorage.removeItem(prefix + key); });
        },
        getItems: function (keys) {
            return new Promise(function (resolve, reject) {
                var result = {};
                var errors = [];
                try {
                    ensureFreshLocalStorage();
                }
                catch (e) {
                    reject(e);
                }
                keys.forEach(function (key) {
                    try {
                        var value = localStorageGetItem(key);
                        result[key] = value ? value : null;
                    }
                    catch (e) {
                        errors.push(e);
                    }
                });
                if (errors.length > 0) {
                    reject(new Error("Unknown error."));
                }
                else {
                    resolve(result);
                }
            });
        },
        setItems: function (keyValues) {
            return performMultiAction(keyValues, function (key) { return localStorageSetItem(key, keyValues[key]); });
        },
        removeItems: function (keys) {
            return performMultiAction(keys, function (key) { return window.localStorage.removeItem(prefix + key); });
        },
        getKeys: function () {
            return performActionAndReturnResult(function () {
                return Object.keys(window.localStorage)
                    .filter(function (fullKey) { return fullKey.indexOf(prefix) === 0; })
                    .map(function (fullKey) { return fullKey.substr(prefix.length); });
            });
        }
    };
    if (OfficeRuntime.ApiHelper.isNativeStorage()) {
        OfficeRuntime.storage = storageNative;
    }
    else {
        OfficeRuntime.storage = storageWeb;
    }
    function localStorageGetItem(key) {
        return window.localStorage.getItem(prefix + key);
    }
    function localStorageSetItem(key, value) {
        window.localStorage.setItem(prefix + key, value);
    }
    function callPersistentKvStorageManager(nativeCall, getValueOnSuccess) {
        return new OfficeExtension.CoreUtility.Promise(function (resolve, reject) {
            var storageManager = PersistentKvStorageManager.getInstance();
            var invokeId = storageManager.setCallBack(function (result, error) {
                if (error) {
                    reject(error);
                    return;
                }
                var value = getValueOnSuccess(result);
                resolve(value);
            });
            storageManager.ctx
                .sync()
                .then(function () {
                var storageService = storageManager.getPersistentKvStorageService();
                nativeCall(storageService, invokeId);
                return storageManager.ctx.sync();
            })
                .catch(function (e) {
                reject(e);
            });
        });
    }
    function ensureFreshLocalStorage() {
        localStorageSetItem(dummyUnusedKey, null);
        window.localStorage.removeItem(dummyUnusedKey);
    }
    function performAction(action, callback) {
        if (callback === void 0) { callback = function () {
        }; }
        return new Promise(function (resolve, reject) {
            try {
                ensureFreshLocalStorage();
                action();
                callback(null);
                resolve();
            }
            catch (e) {
                callback(e);
                reject(e);
            }
        });
    }
    function performActionAndReturnResult(action, callback) {
        if (callback === void 0) { callback = function () {
        }; }
        return new Promise(function (resolve, reject) {
            try {
                ensureFreshLocalStorage();
                var result = action();
                callback(null, result);
                resolve(result);
            }
            catch (e) {
                callback(e, null);
                reject(e);
            }
        });
    }
    function performMultiAction(collection, action, callback) {
        if (callback === void 0) { callback = function () {
        }; }
        return new Promise(function (resolve, reject) {
            var errors = [];
            try {
                ensureFreshLocalStorage();
            }
            catch (e) {
                errors.push(e);
            }
            for (var key in collection) {
                if (collection.hasOwnProperty(key) || Array.isArray(collection)) {
                    try {
                        if (Array.isArray(collection)) {
                            action(collection[key]);
                        }
                        else {
                            action(key);
                        }
                    }
                    catch (e) {
                        errors.push(e);
                    }
                }
            }
            callback(errors);
            if (errors.length > 0) {
                reject(errors);
            }
            else {
                resolve();
            }
        });
    }
    var PersistentKvStorageManager = (function () {
        function PersistentKvStorageManager() {
            var _this = this;
            this._invokeId = 0;
            this._callDict = {};
            this.ctx = new OfficeExtension.ClientRequestContext();
            this._perkvstorService = PersistentKvStorageService.newObject(this.ctx);
            this._eventResult = this._perkvstorService.onPersistentStorageMessage.add(function (args) {
                OfficeExtension.Utility.log("persistentKvStoragegMessageHandler:" + JSON.stringify(args));
                var callback = _this._callDict[args.invokeId];
                if (callback) {
                    callback(args.message, args.error);
                    delete _this._callDict[args.invokeId];
                }
            });
        }
        PersistentKvStorageManager.getInstance = function () {
            if (PersistentKvStorageManager.instance === undefined) {
                PersistentKvStorageManager.instance = new PersistentKvStorageManager();
            }
            else {
                PersistentKvStorageManager.instance._perkvstorService = PersistentKvStorageService.newObject(PersistentKvStorageManager.instance.ctx);
            }
            return PersistentKvStorageManager.instance;
        };
        PersistentKvStorageManager.prototype.getPersistentKvStorageService = function () {
            return this._perkvstorService;
        };
        PersistentKvStorageManager.prototype.getCallBack = function (callId) {
            return this._callDict[callId];
        };
        PersistentKvStorageManager.prototype.setCallBack = function (callback) {
            var id = this._invokeId;
            this._callDict[this._invokeId++] = callback;
            return id;
        };
        return PersistentKvStorageManager;
    }());
    var PersistentKvStorageService = (function (_super) {
        __extends(PersistentKvStorageService, _super);
        function PersistentKvStorageService() {
            return _super !== null && _super.apply(this, arguments) || this;
        }
        Object.defineProperty(PersistentKvStorageService.prototype, "_className", {
            get: function () {
                return "PersistentKvStorageService";
            },
            enumerable: true,
            configurable: true
        });
        PersistentKvStorageService.prototype.clear = function (id) {
            _invokeMethod(this, "Clear", 1, [id], 4, 0);
        };
        PersistentKvStorageService.prototype.getAllKeys = function (id) {
            _invokeMethod(this, "GetAllKeys", 1, [id], 4, 0);
        };
        PersistentKvStorageService.prototype.multiGet = function (id, jsonKeys) {
            _invokeMethod(this, "MultiGet", 1, [id, jsonKeys], 4, 0);
        };
        PersistentKvStorageService.prototype.multiRemove = function (id, jsonKeys) {
            _invokeMethod(this, "MultiRemove", 1, [id, jsonKeys], 4, 0);
        };
        PersistentKvStorageService.prototype.multiSet = function (id, jsonKeyValue) {
            _invokeMethod(this, "MultiSet", 1, [id, jsonKeyValue], 4, 0);
        };
        PersistentKvStorageService.prototype._handleResult = function (value) {
            _super.prototype._handleResult.call(this, value);
            if (_isNullOrUndefined(value))
                return;
            var obj = value;
            _fixObjectPathIfNecessary(this, obj);
        };
        PersistentKvStorageService.prototype._handleRetrieveResult = function (value, result) {
            _super.prototype._handleRetrieveResult.call(this, value, result);
            _processRetrieveResult(this, value, result);
        };
        PersistentKvStorageService.newObject = function (context) {
            return _createTopLevelServiceObject(OfficeRuntime.PersistentKvStorageService, context, "Microsoft.PersistentKvStorage.PersistentKvStorageService", false, 4);
        };
        Object.defineProperty(PersistentKvStorageService.prototype, "onPersistentStorageMessage", {
            get: function () {
                if (!this.m_persistentStorageMessage) {
                    this.m_persistentStorageMessage = new OfficeExtension.GenericEventHandlers(this.context, this, "PersistentStorageMessage", {
                        eventType: 65537,
                        registerFunc: function () { return void {}; },
                        unregisterFunc: function () { return void {}; },
                        getTargetIdFunc: function () { return null; },
                        eventArgsTransformFunc: function (args) {
                            var perkvstorArgs;
                            try {
                                var parsedMessage = JSON.parse(args.message);
                                var hr = parseInt(parsedMessage.errorCode);
                                var error = hr != 0 ? new OfficeExtension.Error(getErrorCodeAndMessage(hr)) : null;
                                perkvstorArgs = {
                                    invokeId: parsedMessage.invokeId,
                                    message: parsedMessage.message,
                                    error: error
                                };
                            }
                            catch (e) {
                                perkvstorArgs = {
                                    invokeId: -1,
                                    message: e.message,
                                    error: new OfficeExtension.Error({ code: "GenericException", message: "Unknown error" })
                                };
                            }
                            return OfficeExtension.Utility._createPromiseFromResult(perkvstorArgs);
                            function getErrorCodeAndMessage(internalCode) {
                                var _a;
                                var table = (_a = {},
                                    _a[16389] = { code: "GenericException", message: "Unknown error." },
                                    _a[65535] = { code: "Unexcepted", message: "Catastrophic failure." },
                                    _a[14] = { code: "OutOfMemory", message: "Ran out of memory." },
                                    _a[87] = { code: "InvalidArg", message: "One or more arguments are invalid." },
                                    _a[16385] = { code: "NotImplemented", message: "Not implemented." },
                                    _a[6] = { code: "BadHandle", message: "File Handle is not Set." },
                                    _a[5] = { code: "AccessDenied", message: "Can't read the Storage File." },
                                    _a);
                                if (table[internalCode]) {
                                    return table[internalCode];
                                }
                                else {
                                    return { code: "Unknown", message: "An unknown error has occured" };
                                }
                            }
                        }
                    });
                }
                return this.m_persistentStorageMessage;
            },
            enumerable: true,
            configurable: true
        });
        PersistentKvStorageService.prototype.toJSON = function () {
            return _toJson(this, {}, {});
        };
        return PersistentKvStorageService;
    }(OfficeExtension.ClientObject));
    OfficeRuntime.PersistentKvStorageService = PersistentKvStorageService;
    var ErrorCodes;
    (function (ErrorCodes) {
        ErrorCodes["generalException"] = "GeneralException";
    })(ErrorCodes = OfficeRuntime.ErrorCodes || (OfficeRuntime.ErrorCodes = {}));
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
        Object.defineProperty(RequestContext.prototype, "ribbon", {
            get: function () {
                if (!this.m_ribbon) {
                    this.m_ribbon = OfficeCore.DynamicRibbon.newObject(this);
                }
                return this.m_ribbon;
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
var Office;
(function (Office) {
    var license;
    (function (license_1) {
        function _createRequestContext() {
            var context = new OfficeCore.RequestContext();
            if (OSF._OfficeAppFactory.getHostInfo().hostPlatform == 'web') {
                context._customData = 'WacPartition';
            }
            return context;
        }
        function isFeatureEnabled(feature, fallbackValue) {
            return __awaiter(this, void 0, void 0, function () {
                var context, license, isEnabled;
                return __generator(this, function (_a) {
                    switch (_a.label) {
                        case 0:
                            context = _createRequestContext();
                            license = OfficeCore.License.newObject(context);
                            isEnabled = license.isFeatureEnabled(feature, fallbackValue);
                            return [4, context.sync()];
                        case 1:
                            _a.sent();
                            return [2, isEnabled.value];
                    }
                });
            });
        }
        license_1.isFeatureEnabled = isFeatureEnabled;
        function getFeatureTier(feature, fallbackValue) {
            return __awaiter(this, void 0, void 0, function () {
                var context, license, tier;
                return __generator(this, function (_a) {
                    switch (_a.label) {
                        case 0:
                            context = _createRequestContext();
                            license = OfficeCore.License.newObject(context);
                            tier = license.getFeatureTier(feature, fallbackValue);
                            return [4, context.sync()];
                        case 1:
                            _a.sent();
                            return [2, tier.value];
                    }
                });
            });
        }
        license_1.getFeatureTier = getFeatureTier;
        function isFreemiumUpsellEnabled() {
            return __awaiter(this, void 0, void 0, function () {
                var context, license, isFreemiumUpsellEnabled;
                return __generator(this, function (_a) {
                    switch (_a.label) {
                        case 0:
                            context = _createRequestContext();
                            license = OfficeCore.License.newObject(context);
                            isFreemiumUpsellEnabled = license.isFreemiumUpsellEnabled();
                            return [4, context.sync()];
                        case 1:
                            _a.sent();
                            return [2, isFreemiumUpsellEnabled.value];
                    }
                });
            });
        }
        license_1.isFreemiumUpsellEnabled = isFreemiumUpsellEnabled;
        function launchUpsellExperience(experienceId) {
            return __awaiter(this, void 0, void 0, function () {
                var context, license;
                return __generator(this, function (_a) {
                    switch (_a.label) {
                        case 0:
                            context = _createRequestContext();
                            license = OfficeCore.License.newObject(context);
                            license.launchUpsellExperience(experienceId);
                            return [4, context.sync()];
                        case 1:
                            _a.sent();
                            return [2];
                    }
                });
            });
        }
        license_1.launchUpsellExperience = launchUpsellExperience;
        function onFeatureStateChanged(feature, listener) {
            return __awaiter(this, void 0, void 0, function () {
                var context, license, licenseFeature, removeListener;
                return __generator(this, function (_a) {
                    switch (_a.label) {
                        case 0:
                            context = _createRequestContext();
                            license = OfficeCore.License.newObject(context);
                            licenseFeature = license.getLicenseFeature(feature);
                            licenseFeature.onStateChanged.add(listener);
                            removeListener = function () {
                                licenseFeature.onStateChanged.remove(listener);
                                return null;
                            };
                            return [4, context.sync()];
                        case 1:
                            _a.sent();
                            return [2, removeListener];
                    }
                });
            });
        }
        license_1.onFeatureStateChanged = onFeatureStateChanged;
        function getMsaDeviceTicket(resource, policy, options) {
            return __awaiter(this, void 0, void 0, function () {
                var context, license, msaDeviceTicket;
                return __generator(this, function (_a) {
                    switch (_a.label) {
                        case 0:
                            context = _createRequestContext();
                            license = OfficeCore.License.newObject(context);
                            msaDeviceTicket = license.getMsaDeviceTicket(resource, policy, options);
                            return [4, context.sync()];
                        case 1:
                            _a.sent();
                            return [2, msaDeviceTicket.value];
                    }
                });
            });
        }
        license_1.getMsaDeviceTicket = getMsaDeviceTicket;
    })(license = Office.license || (Office.license = {}));
})(Office || (Office = {}));
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
    var _calculateApiFlags = OfficeExtension.CommonUtility.calculateApiFlags;
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
        "clientObjectTypes": [[1,
                0,
                0,
                0,
                [["executeAction",
                        3,
                        2,
                        0,
                        5],
                    ["notifyPaneEvent",
                        2,
                        2,
                        0,
                        5],
                    [2,
                        0,
                        0,
                        0,
                        1],
                    [3,
                        0,
                        0,
                        0,
                        1],
                    ["testFireEvent",
                        0,
                        0,
                        0,
                        1]],
                0,
                0,
                0,
                [["HostSkillEvent",
                        2,
                        0,
                        "65538",
                        "",
                        2,
                        3]],
                "Microsoft.SkillApi.Skill",
                4]] };
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
    var _setMockData = OfficeExtension.Utility.setMockData;
    var _calculateApiFlags = OfficeExtension.CommonUtility.calculateApiFlags;
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
        TelemetryService.prototype.sendCustomerContent = function (telemetryProperties, eventName, eventContract, eventFlags, value) {
            _throwIfApiNotSupported("TelemetryService.sendCustomerContent", "Telemetry", "1.3", _hostName);
            _invokeMethod(this, "SendCustomerContent", 1, [telemetryProperties, eventName, eventContract, eventFlags, value], 4, 0);
        };
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
    var WebAuthReplyUrlsStorageKey = "officeWebAuthReplyUrls";
    var loaded = false;
    OfficeFirstPartyAuth.authFlow = "implicit";
    OfficeFirstPartyAuth.autoPopup = false;
    OfficeFirstPartyAuth.debugging = false;
    function load(replyUrl) {
        return new OfficeExtension.CoreUtility.Promise(function (resolve, reject) {
            if (OSF.WebAuth && OSF._OfficeAppFactory.getHostInfo().hostPlatform == "web") {
                var retrievedAuthContext = false;
                var errorMessage;
                try {
                    if (!Office || !Office.context || !Office.context.webAuth) {
                        reject({
                            code: "GetAuthContextAsyncMissing",
                            message: (typeof (Strings) !== 'undefined' && Strings.OfficeOM.L_ImplicitGetAuthContextMissing) ? Strings.OfficeOM.L_ImplicitGetAuthContextMissing : ""
                        });
                        return;
                    }
                    Office.context.webAuth.getAuthContextAsync(function (result) {
                        if (result.status === "succeeded") {
                            retrievedAuthContext = true;
                            var authContext = result.value;
                            if (!authContext || authContext.isAnonymous) {
                                reject({
                                    code: "CannotGetAuthContext",
                                    message: (typeof (Strings) !== 'undefined' && Strings.OfficeOM.L_ImplicitGetAuthContextMissing) ? Strings.OfficeOM.L_ImplicitGetAuthContextMissing : ""
                                });
                                return;
                            }
                            var isMsa = authContext.authorityType.toLowerCase() === 'msa';
                            OSF.WebAuth.config = {
                                authFlow: OfficeFirstPartyAuth.authFlow,
                                authVersion: (OfficeFirstPartyAuth.authVersion) ? OfficeFirstPartyAuth.authVersion : null,
                                debugging: OfficeFirstPartyAuth.debugging,
                                authority: (OfficeFirstPartyAuth.authorityOverride) ? OfficeFirstPartyAuth.authorityOverride : authContext.authority,
                                idp: authContext.authorityType.toLowerCase(),
                                appIds: [isMsa ? (authContext.msaAppId) ? authContext.msaAppId : authContext.appId : authContext.appId],
                                redirectUri: (replyUrl) ? replyUrl : null,
                                upn: authContext.upn,
                                telemetryInstance: 'otel',
                                autoPopup: OfficeFirstPartyAuth.autoPopup,
                                enableConsoleLogging: OfficeFirstPartyAuth.debugging
                            };
                            OSF.WebAuth.load().then(function (loadResult) {
                                loaded = true;
                                logLoadEvent(loadResult, true);
                                resolve();
                            }).catch(function (loadError) {
                                logLoadEvent(loadError, false);
                                reject({
                                    code: "PackageNotLoaded",
                                    message: (typeof (Strings) !== 'undefined' && Strings.OfficeOM.L_ImplicitNotLoaded) ? Strings.OfficeOM.L_ImplicitNotLoaded : ""
                                });
                            });
                            var finalReplyUrl = (replyUrl) ? replyUrl : window.location.href.split("?")[0];
                            var replyUrls = sessionStorage.getItem(WebAuthReplyUrlsStorageKey);
                            if (replyUrls || replyUrls === "") {
                                replyUrls = finalReplyUrl;
                            }
                            else {
                                replyUrls += ", " + finalReplyUrl;
                            }
                            if (replyUrls)
                                sessionStorage.setItem(WebAuthReplyUrlsStorageKey, replyUrls);
                        }
                        else {
                            OSF.WebAuth.config = null;
                            errorMessage = JSON.stringify(result);
                            reject({
                                code: "CannotGetAuthContext",
                                message: errorMessage
                            });
                        }
                    });
                }
                catch (e) {
                    errorMessage = e;
                    OSF.WebAuth.config = null;
                    OSF.WebAuth.load().then(function () {
                        resolve();
                    }).catch(function () {
                        reject({
                            code: retrievedAuthContext ? "CannotGetAuthContext" : "FailedToLoad",
                            message: errorMessage
                        });
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
        return new OfficeExtension.CoreUtility.Promise(function (resolve, reject) {
            if (OSF._OfficeAppFactory.getHostInfo().hostPlatform == "web") {
                Office.context.webAuth.getAuthContextAsync(function (result) {
                    var supportsAuthToken = false;
                    if (result.status === "succeeded") {
                        var authContext = result.value;
                        if (authContext.supportsAuthToken) {
                            supportsAuthToken = true;
                        }
                    }
                    if (!supportsAuthToken) {
                        if (OSF.WebAuth && loaded) {
                            if (OSF.WebAuth.config.appIds[0]) {
                                OSF.WebAuth.getToken(options.resource, OSF.WebAuth.config.appIds[0], OSF._OfficeAppFactory.getHostInfo().osfControlAppCorrelationId, (behaviorOption && behaviorOption.popup) ? behaviorOption.popup : null).then(function (result) {
                                    logAcquireEvent(result, true, options.resource, (behaviorOption && behaviorOption.popup) ? behaviorOption.popup : null);
                                    resolve({
                                        accessToken: result.Token,
                                        tokenIdenityType: (OSF.WebAuth.config.idp.toLowerCase() == "msa")
                                            ? OfficeCore.IdentityType.microsoftAccount
                                            : OfficeCore.IdentityType.organizationAccount
                                    });
                                }).catch(function (result) {
                                    logAcquireEvent(result, false, options.resource, (behaviorOption && behaviorOption.popup) ? behaviorOption.popup : null, result.ErrorCode);
                                    reject({ code: result.ErrorCode, message: result.ErrorMessage });
                                });
                            }
                        }
                        else {
                            logUnexpectedAcquireEvent(loaded, OSF.WebAuth.loadAttempts);
                        }
                    }
                    else {
                        var context = new OfficeCore.RequestContext();
                        var auth = OfficeCore.AuthenticationService.newObject(context);
                        context._customData = "WacPartition";
                        var result_1 = auth.getAccessToken(options, null);
                        context.sync().then(function () {
                            resolve(result_1.value);
                        });
                    }
                });
            }
            else {
                var context_1 = new OfficeCore.RequestContext();
                var auth_1 = OfficeCore.AuthenticationService.newObject(context_1);
                var handler_1 = auth_1.onTokenReceived.add(function (arg) {
                    if (!OfficeExtension.CoreUtility.isNullOrUndefined(arg)) {
                        handler_1.remove();
                        context_1.sync().catch(function () {
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
                context_1.sync()
                    .then(function () {
                    var apiResult = auth_1.getAccessToken(options, auth_1._targetId);
                    return context_1.sync()
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
            }
        });
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
            console.log("Logging WebAuthJs load event");
        }
        if (typeof OTel !== "undefined") {
            OTel.OTelLogger.onTelemetryLoaded(function () {
                var telemetryData = [
                    oteljs.makeStringDataField('IdentityProvider', OSF.WebAuth.config.idp),
                    oteljs.makeStringDataField('AppId', OSF.WebAuth.config.appIds[0]),
                    oteljs.makeBooleanDataField('Js', (typeof Implicit !== "undefined" || typeof BrowserAuth !== "undefined") ? true : false),
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
            console.log("Logging WebAuthJs acquire event");
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
    function logUnexpectedAcquireEvent(loadResult, loadAttempts) {
        if (OfficeFirstPartyAuth.debugging) {
            console.log("Logging WebAuthJs unexpected acquire event");
        }
        if (typeof OTel !== "undefined") {
            OTel.OTelLogger.onTelemetryLoaded(function () {
                var telemetryData = [
                    oteljs.makeBooleanDataField('Loaded', loadResult),
                    oteljs.makeInt64DataField('LoadAttempts', loadAttempts)
                ];
                OTel.OTelLogger.sendTelemetryEvent({
                    eventName: "Office.Extensibility.OfficeJs.OfficeFirstPartyAuth.UnexpectedAcquire",
                    dataFields: telemetryData,
                    eventFlags: {
                        dataCategories: oteljs.DataCategories.ProductServiceUsage
                    }
                });
            });
        }
    }
    function loadWebAuthForReplyPage() {
        try {
            if (typeof (window) === "undefined" || !window.sessionStorage) {
                return;
            }
            var webAuthRedirectUrls = sessionStorage.getItem(WebAuthReplyUrlsStorageKey);
            if (webAuthRedirectUrls !== null && webAuthRedirectUrls.indexOf(window.location.origin + window.location.pathname) !== -1) {
                load();
            }
        }
        catch (ex) {
            console.error(ex);
        }
    }
    if (typeof (window) !== "undefined" && window.OSF) {
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
    var _calculateApiFlags = OfficeExtension.CommonUtility.calculateApiFlags;
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
        function LicenseFeature_StateChanged_EventArgsTransform(thisObj, args) {
            var newArgs = {
                feature: args.featureName,
                isEnabled: args.isEnabled,
                tier: args.tierName
            };
            if (args.tierName) {
                newArgs.tier = args.tierName == 0 ? LicenseFeatureTier.unknown :
                    args.tierName == 1 ? LicenseFeatureTier.basic :
                        args.tierName == 2 ? LicenseFeatureTier.premium :
                            args.tierName;
            }
            return newArgs;
        }
        _CC.LicenseFeature_StateChanged_EventArgsTransform = LicenseFeature_StateChanged_EventArgsTransform;
    })(_CC = OfficeCore._CC || (OfficeCore._CC = {}));
    var ErrorCodes;
    (function (ErrorCodes) {
        ErrorCodes["apiNotAvailable"] = "ApiNotAvailable";
        ErrorCodes["clientError"] = "ClientError";
        ErrorCodes["controlIdNotFound"] = "ControlIdNotFound";
        ErrorCodes["entryIdRequired"] = "EntryIdRequired";
        ErrorCodes["generalException"] = "GeneralException";
        ErrorCodes["hostRestartNeeded"] = "HostRestartNeeded";
        ErrorCodes["instanceNotFound"] = "InstanceNotFound";
        ErrorCodes["interactiveFlowAborted"] = "InteractiveFlowAborted";
        ErrorCodes["invalidArgument"] = "InvalidArgument";
        ErrorCodes["invalidGrant"] = "InvalidGrant";
        ErrorCodes["invalidResourceUrl"] = "InvalidResourceUrl";
        ErrorCodes["invalidRibbonDefinition"] = "InvalidRibbonDefinition";
        ErrorCodes["objectNotFound"] = "ObjectNotFound";
        ErrorCodes["resourceNotSupported"] = "ResourceNotSupported";
        ErrorCodes["serverError"] = "ServerError";
        ErrorCodes["serviceUrlNotFound"] = "ServiceUrlNotFound";
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
        "apiSets": [["1.2", "FirstPartyAuthentication"], ["1.3", "FirstPartyAuthentication"], ["1.2", "DynamicRibbon"]],
        "strings": ["AuthenticationService", "RoamingSetting", "RoamingSettingCollection", "ServiceUrlProvider", "LinkedIn", "NetworkUsage", "DynamicRibbon", "RibbonTab", "RibbonButton", "RibbonButtonCollection", "LocaleApi", "OfficeServicesManagerApi", "Comment", "CommentCollection", "MemberInfoList", "PersonaActions", "PersonaInfoSource", "PersonaInfo", "PersonaUnifiedCommunicationInfo", "PersonaPhotoInfo", "PersonaCollection", "PersonaOrganizationInfo", "UnifiedGroupInfo", "Persona", "PersonaLifetime", "LokiTokenProvider", "LokiTokenProviderFactory", "ServiceContext", "RichapiPcxFeatureChecks", "Tap", "AppRuntimePersistenceService", "AppRuntimeService", "License", "LicenseFeature", "MsaDeviceTicketOptions", "DialogPage", "null", "id", "getItem", "getCount", "isWarmedUp", "isWarmingUp", "displayName", "email", "emailAddresses", "sipAddresses", "birthday", "birthdays", "title", "jobInfoDepartment", "companyName", "office", "linkedTitles", "linkedDepartments", "linkedCompanyNames", "linkedOffices", "webSites", "notes", "getImageUri", "setPlaceholderColor", "getPlaceholderUri", "getImageUriWithMetadata", "instanceId", "dispose", "_RegisterPersonaUpdatedEvent", "_UnregisterPersonaUpdatedEvent", "this.instanceId", "_RegisterLokiTokenAvailableEvent", "_UnregisterLokiTokenAvailableEvent", "_RegisterIdentityUniqueIdAvailableEvent", "_UnregisterIdentityUniqueIdAvailableEvent", "_RegisterClientAccessTokenAvailableEvent", "_UnregisterClientAccessTokenAvailableEvent", "getLokiTokenProvider", "_RegisterStateChange", "_UnregisterStateChange", "registerOnShow", "unregisterOnShow"],
        "enumTypes": [["IdentityType", ["organizationAccount", "microsoftAccount", "unsupported"]],
            ["ServiceProvider", ["ariaBrowserPipeUrl", "ariaUploadUrl", "ariaVNextUploadUrl", "lokiAutoDiscoverUrl"]],
            ["TimeStringFormat", ["shortTime", "longTime", "shortDate", "longDate"]],
            ["CommentTextFormat", ["plain", "markdown", "delta"]],
            ["PersonaCardPerfPoint", ["placeHolderRendered", "initialCardRendered"]],
            ["MessageType", [], { "personaLifetimePersonaUpdatedEvent": 3502, "lokiTokenProviderLokiTokenAvailableEvent": 3503, "lokiTokenProviderIdentityUniqueIdAvailableEvent": 3504, "lokiTokenProviderClientAccessTokenAvailableEvent": 3505 }],
            ["UnifiedCommunicationAvailability", ["notSet", "free", "idle", "busy", "idleBusy", "doNotDisturb", "unalertable", "unavailable"]],
            ["UnifiedCommunicationStatus", ["online", "notOnline", "away", "busy", "beRightBack", "onThePhone", "outToLunch", "inAMeeting", "outOfOffice", "doNotDisturb", "inAConference", "getting", "notABuddy", "disconnected", "notInstalled", "urgentInterruptionsOnly", "mayBeAvailable", "idle", "inPresentation"]],
            ["UnifiedCommunicationPresence", ["free", "busy", "idle", "doNotDistrub", "blocked", "notSet", "outOfOffice"]],
            ["FreeBusyCalendarState", ["unknown", "free", "busy", "elsewhere", "tentative", "outOfOffice"]],
            ["PersonaType", ["unknown", "enterprise", "contact", "bot", "phoneOnly", "oneOff", "distributionList", "personalDistributionList", "anonymous", "unifiedGroup"]],
            ["PhoneType", ["workPhone", "homePhone", "mobilePhone", "businessFax", "otherPhone"]],
            ["AddressType", ["workAddress", "homeAddress", "otherAddress"]],
            ["MemberType", ["unknown", "individual", "group"]],
            ["PersonaDataUpdated", ["hostId", "type", "photo", "personaInfo", "unifiedCommunicationInfo", "organization", "unifiedGroupInfo", "members", "membership", "capabilities", "customizations", "viewableSources", "placeholder"]],
            ["CustomizedData", ["email", "workPhone", "workPhone2", "workFax", "mobilePhone", "homePhone", "homePhone2", "otherPhone", "sipAddress", "profile", "office", "company", "workAddress", "homeAddress", "otherAddress", "birthday"]],
            ["ObjectType", ["unknown", "chart", "smartArt", "table", "image", "slide", "text"], { "ole": "OLE" }],
            ["AppRuntimeState", ["inactive", "background", "visible"]],
            ["Visibility", ["hidden", "visible"]],
            ["LicenseFeatureTier", ["unknown", "basic", "premium"]],
            ["LicenseEventType", [], { "featureStateChanged": 1 }],
            ["DialogPageEventType", [], { "onShow": 1 }]],
        "clientObjectTypes": [[1,
                4,
                0,
                [["roamingSettings",
                        3,
                        2,
                        0,
                        0,
                        4]],
                [["getAccessToken",
                        2,
                        2,
                        0,
                        5],
                    ["getPrimaryIdentityInfo",
                        0,
                        2,
                        1,
                        5],
                    ["getIdentities",
                        0,
                        2,
                        2,
                        5]],
                0,
                0,
                0,
                [["TokenReceived",
                        2,
                        1,
                        "3001",
                        "this._targetId",
                        37,
                        37]],
                "Microsoft.Authentication.AuthenticationService",
                4],
            [2,
                0,
                [[38,
                        3],
                    ["value",
                        1]]],
            [3,
                0,
                0,
                0,
                0,
                [[39,
                        2,
                        1,
                        2,
                        0,
                        4],
                    ["getItemOrNullObject",
                        2,
                        1,
                        2,
                        0,
                        4]]],
            [4,
                0,
                0,
                0,
                [["getServiceUrl",
                        2,
                        2,
                        0,
                        4]],
                0,
                0,
                0,
                0,
                "Microsoft.DesktopCompliance.ServiceUrlProvider",
                4],
            [5,
                0,
                0,
                0,
                [["isEnabledForOffice",
                        0,
                        2,
                        0,
                        4],
                    ["recordLinkedInSettingsCompliance",
                        2]],
                0,
                0,
                0,
                0,
                "Microsoft.DesktopCompliance.LinkedIn",
                4],
            [6,
                0,
                0,
                0,
                [["isInOnlineMode",
                        0,
                        2,
                        0,
                        4],
                    ["isInDisconnectedMode",
                        0,
                        2,
                        0,
                        4]],
                0,
                0,
                0,
                0,
                "Microsoft.DesktopCompliance.NetworkUsage",
                4],
            [7,
                0,
                0,
                [["buttons",
                        10,
                        19,
                        0,
                        0,
                        4]],
                [["executeRequestUpdate",
                        1,
                        2,
                        0,
                        4],
                    ["executeRequestCreate",
                        1,
                        2,
                        3,
                        4]],
                [["getButton",
                        9,
                        1,
                        2,
                        0,
                        4],
                    ["getTab",
                        8,
                        1,
                        2,
                        0,
                        4]],
                0,
                0,
                0,
                "Microsoft.DynamicRibbon.DynamicRibbon",
                4],
            [8,
                0,
                [[38,
                        3]],
                0,
                [["setVisibility",
                        1]]],
            [9,
                0,
                [[38,
                        3],
                    ["enabled",
                        1],
                    ["label",
                        3]],
                0,
                [["setEnabled",
                        1]]],
            [10,
                1,
                0,
                0,
                [[40,
                        0,
                        2,
                        0,
                        4]],
                [[39,
                        9,
                        1,
                        18,
                        0,
                        4]],
                0,
                9],
            [11,
                0,
                0,
                0,
                [["getLocaleDateTimeFormattingInfo",
                        1,
                        2,
                        0,
                        4],
                    ["formatDateTimeString",
                        3,
                        2,
                        0,
                        4]],
                0,
                0,
                0,
                0,
                "Microsoft.LocaleApi.LocaleApi",
                4],
            [12,
                0,
                0,
                0,
                [["bindServiceToProfile",
                        3]],
                0,
                0,
                0,
                0,
                "Microsoft.OfficeServicesManager.OfficeServicesManagerApi",
                4],
            [13,
                0,
                [[38,
                        3],
                    ["text",
                        1],
                    ["created",
                        11],
                    ["level",
                        3],
                    ["resolved",
                        1],
                    ["author",
                        3],
                    ["mentions",
                        3]],
                [["parent",
                        13,
                        2,
                        0,
                        0,
                        4],
                    ["parentOrNullObject",
                        13,
                        2,
                        0,
                        0,
                        4],
                    ["replies",
                        14,
                        19,
                        0,
                        0,
                        4]],
                [["getRichText",
                        1,
                        2,
                        0,
                        4],
                    ["setRichText",
                        2],
                    ["delete"]],
                [["getParentOrSelf",
                        13,
                        0,
                        2,
                        0,
                        4],
                    ["reply",
                        13,
                        2]]],
            [14,
                1,
                0,
                0,
                [[40,
                        0,
                        2,
                        0,
                        4]],
                [[39,
                        13,
                        1,
                        18,
                        0,
                        4]],
                0,
                13],
            [15,
                0,
                [[41,
                        3],
                    [42,
                        3]],
                0,
                [["items",
                        0,
                        2,
                        0,
                        4]],
                [["getPersonaForMember",
                        24,
                        1,
                        2,
                        0,
                        4]]],
            [16,
                0,
                0,
                0,
                [["addContact"],
                    ["editContact"],
                    ["composeEmail",
                        1],
                    ["composeInstantMessage",
                        1],
                    ["callPhoneNumber",
                        1],
                    ["pinPersonaToQuickContacts"],
                    ["toggleTagForAlerts"],
                    ["scheduleMeeting"],
                    ["openLinkContactUx"],
                    ["editContactByIdentifier",
                        1],
                    ["showHoverCardForPersona",
                        6],
                    ["hideHoverCardForPersona"],
                    ["showContextMenu",
                        6],
                    ["showContactCard",
                        6],
                    ["showExpandedCard",
                        6],
                    ["openGroupCalendar"],
                    ["subscribeToGroup"],
                    ["unsubscribeFromGroup"],
                    ["getChangePhotoUrlAndOpenInBrowser"],
                    ["startAudioCall"],
                    ["startVideoCall"],
                    ["openOutlookProperties"],
                    ["editUnifiedGroup"],
                    ["joinUnifiedGroup"],
                    ["leaveUnifiedGroup"]]],
            [17,
                0,
                [[43,
                        3],
                    [44,
                        3],
                    [45,
                        3],
                    [46,
                        3],
                    [47,
                        3],
                    [48,
                        3],
                    [49,
                        3],
                    [50,
                        3],
                    [51,
                        3],
                    [52,
                        3],
                    [53,
                        3],
                    [54,
                        3],
                    [55,
                        3],
                    [56,
                        3],
                    ["phones",
                        3],
                    ["addresses",
                        3],
                    [57,
                        3],
                    [58,
                        3]]],
            [18,
                0,
                [[43,
                        3],
                    [44,
                        3],
                    [45,
                        3],
                    [46,
                        3],
                    [47,
                        11],
                    [48,
                        11],
                    [49,
                        3],
                    [50,
                        3],
                    [51,
                        3],
                    [52,
                        3],
                    [53,
                        3],
                    [54,
                        3],
                    [55,
                        3],
                    [56,
                        3],
                    [57,
                        3],
                    [58,
                        3],
                    ["isPersonResolved",
                        3]],
                [["sources",
                        17,
                        3,
                        0,
                        0,
                        4]],
                [["getPhones",
                        0,
                        2,
                        0,
                        4],
                    ["getAddresses",
                        0,
                        2,
                        0,
                        4]]],
            [19,
                0,
                [["availability",
                        3],
                    ["status",
                        3],
                    ["isSelf",
                        3],
                    ["isTagged",
                        3],
                    ["customStatusString",
                        3],
                    ["isBlocked",
                        3],
                    ["presenceTooltip",
                        3],
                    ["isOutOfOffice",
                        3],
                    ["outOfOfficeNote",
                        3],
                    ["timezone",
                        3],
                    ["meetingLocation",
                        3],
                    ["meetingSubject",
                        3],
                    ["timezoneBias",
                        3],
                    ["idleStartTime",
                        11],
                    ["overallCapability",
                        3],
                    ["isOnBuddyList",
                        3],
                    ["presenceNote",
                        3],
                    ["voiceMailUri",
                        3],
                    ["availabilityText",
                        3],
                    ["availabilityTooltip",
                        3],
                    ["isDurationInAvailabilityText",
                        3],
                    ["freeBusyStatus",
                        3],
                    ["calendarState",
                        3],
                    ["presence",
                        3]]],
            [20,
                0,
                0,
                0,
                [[59,
                        1,
                        2,
                        0,
                        4,
                        59],
                    [60,
                        1,
                        0,
                        0,
                        0,
                        60],
                    [61,
                        1,
                        2,
                        0,
                        4,
                        61],
                    [62,
                        1,
                        2,
                        0,
                        4,
                        62]]],
            [21,
                1,
                0,
                0,
                [[40,
                        0,
                        2,
                        0,
                        4]],
                [[39,
                        24,
                        1,
                        18,
                        0,
                        4]],
                0,
                24],
            [22,
                0,
                [[41,
                        3],
                    [42,
                        3]],
                [["hierarchy",
                        21,
                        18,
                        0,
                        0,
                        4],
                    ["manager",
                        24,
                        2,
                        0,
                        0,
                        4],
                    ["directReports",
                        21,
                        18,
                        0,
                        0,
                        4]]],
            [23,
                0,
                [["description",
                        1],
                    ["oneDrive",
                        1],
                    ["oneNote",
                        1],
                    ["isPublic",
                        1],
                    ["amIOwner",
                        1],
                    ["amIMember",
                        1],
                    ["amISubscribed",
                        1],
                    ["memberCount",
                        1],
                    ["ownerCount",
                        1],
                    ["hasGuests",
                        1],
                    ["site",
                        1],
                    ["planner",
                        1],
                    ["classification",
                        1],
                    ["subscriptionEnabled",
                        1]]],
            [24,
                4,
                [["hostId",
                        3],
                    ["type",
                        3],
                    ["capabilities",
                        3],
                    ["diagnosticId",
                        3],
                    [63,
                        3]],
                [["photo",
                        20,
                        3,
                        0,
                        0,
                        4],
                    ["personaInfo",
                        18,
                        3,
                        0,
                        0,
                        4],
                    ["unifiedCommunicationInfo",
                        19,
                        3,
                        0,
                        0,
                        4],
                    ["organization",
                        22,
                        3,
                        0,
                        0,
                        4],
                    ["unifiedGroupInfo",
                        23,
                        35,
                        0,
                        0,
                        4],
                    ["actions",
                        16,
                        2,
                        0,
                        0,
                        4]],
                [["getCustomizations",
                        0,
                        2,
                        0,
                        4],
                    ["warmup",
                        1],
                    [64],
                    ["getViewableSources",
                        0,
                        2,
                        0,
                        4],
                    ["reportTimeForRender",
                        2]],
                [["getMembers",
                        15,
                        0,
                        2,
                        0,
                        4],
                    ["getMembership",
                        15,
                        0,
                        2,
                        0,
                        4]]],
            [25,
                0,
                [[63,
                        3]],
                0,
                [["getPolicies",
                        0,
                        2,
                        0,
                        4],
                    [65],
                    [66],
                    ["getTextScaleFactor",
                        0,
                        2,
                        0,
                        4]],
                [["getPersona",
                        24,
                        1,
                        2,
                        0,
                        4],
                    ["getPersonaForOrgEntry",
                        24,
                        4,
                        2,
                        0,
                        4],
                    ["getPersonaForOrgByEntryId",
                        24,
                        4,
                        2,
                        0,
                        4]],
                0,
                0,
                [["PersonaUpdated",
                        0,
                        0,
                        "MessageType.personaLifetimePersonaUpdatedEvent",
                        67,
                        65,
                        66]]],
            [26,
                0,
                [["emailOrUpn",
                        3],
                    [63,
                        3]],
                0,
                [["requestToken"],
                    [68],
                    [69],
                    ["requestIdentityUniqueId"],
                    [70],
                    [71],
                    ["requestClientAccessToken"],
                    [72],
                    [73]],
                0,
                0,
                0,
                [["ClientAccessTokenAvailable",
                        0,
                        0,
                        "MessageType.lokiTokenProviderClientAccessTokenAvailableEvent",
                        67,
                        72,
                        73],
                    ["IdentityUniqueIdAvailable",
                        0,
                        0,
                        "MessageType.lokiTokenProviderIdentityUniqueIdAvailableEvent",
                        67,
                        70,
                        71],
                    ["LokiTokenAvailable",
                        0,
                        0,
                        "MessageType.lokiTokenProviderLokiTokenAvailableEvent",
                        67,
                        68,
                        69]]],
            [27,
                0,
                0,
                0,
                0,
                [[74,
                        26,
                        1,
                        2,
                        0,
                        4]],
                0,
                0,
                0,
                "Microsoft.People.LokiTokenProviderFactory",
                4],
            [28,
                0,
                0,
                0,
                [[64,
                        1],
                    ["accountEmailOrUpn",
                        1,
                        2,
                        0,
                        4],
                    ["getPersonaPolicies",
                        0,
                        2,
                        0,
                        4]],
                [[74,
                        26,
                        1,
                        2,
                        0,
                        4],
                    ["getPersonaLifetime",
                        25,
                        1,
                        2,
                        0,
                        4],
                    ["getInitialPersona",
                        24,
                        1,
                        2,
                        0,
                        4]],
                0,
                0,
                0,
                "Microsoft.People.ServiceContext",
                4],
            [29,
                0,
                0,
                0,
                [["isAddChangePhotoLinkOnLpcPersonaImageFlightEnabled",
                        0,
                        2,
                        0,
                        4]],
                0,
                0,
                0,
                0,
                "Microsoft.People.RichapiPcxFeatureChecks",
                4],
            [30,
                0,
                0,
                0,
                [["getEnterpriseUserInfo",
                        0,
                        2,
                        0,
                        5],
                    ["getMruFriendlyPath",
                        1,
                        2,
                        0,
                        5],
                    ["launchFileUrlInOfficeApp",
                        2,
                        2,
                        0,
                        5],
                    ["performLocalSearch",
                        4,
                        2,
                        0,
                        5],
                    ["readSearchCache",
                        3,
                        2,
                        0,
                        5],
                    ["writeSearchCache",
                        3,
                        2,
                        0,
                        5]],
                0,
                0,
                0,
                0,
                "Microsoft.TapRichApi.Tap",
                4],
            [31,
                0,
                0,
                0,
                [["setAppRuntimeStartState",
                        1],
                    ["getAppRuntimeStartState",
                        0,
                        2,
                        0,
                        4]],
                0,
                0,
                0,
                0,
                "Microsoft.AppRuntime.AppRuntimePersistenceService",
                4],
            [32,
                0,
                0,
                0,
                [["setAppRuntimeState",
                        1],
                    ["getAppRuntimeState",
                        0,
                        2,
                        0,
                        4]],
                0,
                0,
                0,
                [["VisibilityChanged",
                        0,
                        0,
                        "65539",
                        "",
                        37,
                        37]],
                "Microsoft.AppRuntime.AppRuntimeService",
                4],
            [33,
                0,
                0,
                0,
                [["isFeatureEnabled",
                        2,
                        2,
                        0,
                        4],
                    ["getFeatureTier",
                        2,
                        2,
                        0,
                        4],
                    ["isFreemiumUpsellEnabled",
                        0,
                        2,
                        0,
                        4],
                    ["launchUpsellExperience",
                        1,
                        2,
                        0,
                        4],
                    ["_TestFireStateChangedEvent",
                        1,
                        0,
                        0,
                        1],
                    ["getMsaDeviceTicket",
                        3,
                        2,
                        0,
                        5]],
                [["getLicenseFeature",
                        34,
                        1,
                        2,
                        0,
                        4]],
                0,
                0,
                0,
                "Microsoft.Office.Licensing.License",
                4],
            [34,
                0,
                [[38,
                        3]],
                0,
                [[75,
                        0,
                        2,
                        0,
                        4],
                    [76,
                        0,
                        2,
                        0,
                        4]],
                0,
                0,
                0,
                [["StateChanged",
                        2,
                        0,
                        "LicenseEventType.featureStateChanged",
                        "this.id",
                        75,
                        76]]],
            [35,
                0,
                [["scopes",
                        1]],
                0,
                0,
                0,
                0,
                0,
                0,
                "Microsoft.Office.Licensing.MsaDeviceTicketOptions",
                4],
            [36,
                0,
                [["_Id",
                        2]],
                0,
                [["close",
                        0,
                        2,
                        0,
                        4],
                    ["readyToShow",
                        0,
                        2,
                        0,
                        4],
                    [77,
                        0,
                        2,
                        0,
                        4],
                    [78,
                        0,
                        2,
                        0,
                        4],
                    ["sendMessageToHost",
                        1,
                        2,
                        0,
                        4]],
                0,
                0,
                0,
                [["OnShowEvent",
                        0,
                        0,
                        "DialogPageEventType.onShow",
                        "this._Id",
                        77,
                        78]],
                "Microsoft.Office.DialogPage.DialogPage",
                4]] };
    var _builder = new OfficeExtension.LibraryBuilder({ metadata: _libraryMetadataOfficeSharedApi, targetNamespaceObject: OfficeCore });
})(OfficeCore || (OfficeCore = {}));
var Office;
(function (Office) {
    var VisibilityMode;
    (function (VisibilityMode) {
        VisibilityMode["hidden"] = "Hidden";
        VisibilityMode["taskpane"] = "Taskpane";
    })(VisibilityMode = Office.VisibilityMode || (Office.VisibilityMode = {}));
    var StartupBehavior;
    (function (StartupBehavior) {
        StartupBehavior["none"] = "None";
        StartupBehavior["load"] = "Load";
    })(StartupBehavior = Office.StartupBehavior || (Office.StartupBehavior = {}));
    var addin;
    (function (addin) {
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
        function _convertVisibilityToVisibilityMode(visibility) {
            if (visibility === OfficeCore.Visibility.visible) {
                return VisibilityMode.taskpane;
            }
            return VisibilityMode.hidden;
        }
        function onVisibilityModeChanged(listener) {
            return __awaiter(this, void 0, void 0, function () {
                var eventService, registrationToken, ret;
                var _this = this;
                return __generator(this, function (_a) {
                    switch (_a.label) {
                        case 0:
                            eventService = _getAppRuntimeEventService();
                            registrationToken = eventService.onVisibilityChanged.add(function (args) {
                                if (listener) {
                                    var msg = {
                                        visibilityMode: _convertVisibilityToVisibilityMode(args.visibility)
                                    };
                                    listener(msg);
                                }
                                return null;
                            });
                            return [4, eventService.context.sync()];
                        case 1:
                            _a.sent();
                            ret = function () { return __awaiter(_this, void 0, void 0, function () {
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
                            }); };
                            return [2, ret];
                    }
                });
            });
        }
        addin.onVisibilityModeChanged = onVisibilityModeChanged;
    })(addin = Office.addin || (Office.addin = {}));
})(Office || (Office = {}));
var Office;
(function (Office) {
    var ribbon;
    (function (ribbon_1) {
        function _createRequestContext() {
            var context = new OfficeCore.RequestContext();
            if (OSF._OfficeAppFactory.getHostInfo().hostPlatform == 'web') {
                context._customData = 'WacPartition';
            }
            return context;
        }
        function requestUpdate(input) {
            var requestContext = _createRequestContext();
            var ribbon = requestContext.ribbon;
            function processControls(parent) {
                if (parent.controls !== undefined
                    && parent.controls.length !== undefined
                    && !!parent.controls.length) {
                    parent.controls
                        .filter(function (control) { return !(!control.id); })
                        .forEach(function (control) {
                        var ribbonControl = ribbon.getButton(control.id);
                        if (control.enabled !== undefined && control.enabled !== null) {
                            ribbonControl.enabled = control.enabled;
                        }
                    });
                }
            }
            input.tabs
                .filter(function (tab) { return !(!tab.id); })
                .forEach(function (tab) {
                var ribbonTab = ribbon.getTab(tab.id);
                if (tab.visible !== undefined && tab.visible !== null) {
                    ribbonTab.setVisibility(tab.visible);
                }
                if (!!tab.groups && !!tab.groups.length) {
                    tab.groups
                        .filter(function (group) { return !(!group.id); })
                        .forEach(function (group) {
                        processControls(group);
                    });
                }
                else {
                    processControls(tab);
                }
            });
            return requestContext.sync();
        }
        ribbon_1.requestUpdate = requestUpdate;
        function requestCreateControls(input) {
            var requestContext = _createRequestContext();
            var ribbon = requestContext.ribbon;
            var delay = function (milliseconds) {
                return new Promise(function (resolve, _) { return setTimeout(function () { return resolve(); }, milliseconds); });
            };
            ribbon.executeRequestCreate(JSON.stringify(input));
            return delay(250)
                .then(function () { return requestContext.sync(); });
        }
        ribbon_1.requestCreateControls = requestCreateControls;
    })(ribbon = Office.ribbon || (Office.ribbon = {}));
})(Office || (Office = {}));
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
    var _calculateApiFlags = OfficeExtension.CommonUtility.calculateApiFlags;
    var AddinInternalServiceErrorCodes;
    (function (AddinInternalServiceErrorCodes) {
        AddinInternalServiceErrorCodes["generalException"] = "GeneralException";
    })(AddinInternalServiceErrorCodes || (AddinInternalServiceErrorCodes = {}));
    var _libraryMetadataInternalServiceApi = { "version": "1.0.0",
        "name": "OfficeCore",
        "defaultApiSetName": "OfficeSharedApi",
        "hostName": "Office",
        "apiSets": [],
        "strings": ["AddinInternalService"],
        "enumTypes": [],
        "clientObjectTypes": [[1,
                0,
                0,
                0,
                [["notifyActionHandlerReady",
                        0,
                        2,
                        0,
                        4]],
                0,
                0,
                0,
                0,
                "Microsoft.InternalService.AddinInternalService",
                4]] };
    var _builder = new OfficeExtension.LibraryBuilder({ metadata: _libraryMetadataInternalServiceApi, targetNamespaceObject: OfficeCore });
})(OfficeCore || (OfficeCore = {}));
var Office;
(function (Office) {
    var actionProxy;
    (function (actionProxy) {
        var _isNullOrUndefined = OfficeExtension.Utility.isNullOrUndefined;
        var _association;
        var ActionMessageCategory = 2;
        var ActionDispatchMessageType = 1000;
        function init() {
            var context = new OfficeExtension.ClientRequestContext();
            return context.eventRegistration.register(5, "", _handleMessage);
        }
        function setActionAssociation(association) {
            _association = association;
        }
        function _getFunction(functionName) {
            if (functionName) {
                var nameUpperCase = functionName.toUpperCase();
                var call = _association.mappings[nameUpperCase];
                if (!_isNullOrUndefined(call) && typeof (call) === "function") {
                    return call;
                }
            }
            throw OfficeExtension.Utility.createRuntimeError("invalidOperation", "sourceData", "ActionProxy._getFunction");
        }
        function _handleMessage(args) {
            try {
                OfficeExtension.Utility.log('ActionProxy._handleMessage');
                OfficeExtension.Utility.checkArgumentNull(args, "args");
                var entryArray = args.entries;
                var invocationArray = [];
                for (var i = 0; i < entryArray.length; i++) {
                    if (entryArray[i].messageCategory !== ActionMessageCategory) {
                        continue;
                    }
                    if (typeof (entryArray[i].message) === 'string') {
                        entryArray[i].message = JSON.parse(entryArray[i].message);
                    }
                    if (entryArray[i].messageType === ActionDispatchMessageType) {
                        var actionsArgs = null;
                        var actionName = entryArray[i].message[0];
                        var call = _getFunction(actionName);
                        if (entryArray[i].message.length >= 2) {
                            var actionArgsJson = entryArray[i].message[1];
                            if (actionArgsJson) {
                                if (_isJsonObjectString(actionArgsJson)) {
                                    actionsArgs = JSON.parse(actionArgsJson);
                                }
                                else {
                                    actionsArgs = actionArgsJson;
                                }
                            }
                        }
                        if (typeof (OSF) !== 'undefined' &&
                            OSF.AppTelemetry &&
                            OSF.AppTelemetry.CallOnAppActivatedIfPending) {
                            OSF.AppTelemetry.CallOnAppActivatedIfPending();
                        }
                        call.apply(null, [actionsArgs]);
                    }
                    else {
                        OfficeExtension.Utility.log('ActionProxy._handleMessage unknown message type ' + entryArray[i].messageType);
                    }
                }
            }
            catch (ex) {
                _tryLog(ex);
                throw ex;
            }
            return OfficeExtension.Utility._createPromiseFromResult(null);
        }
        function _isJsonObjectString(value) {
            if (typeof value === 'string' && value[0] === '{') {
                return true;
            }
            return false;
        }
        function toLogMessage(ex) {
            var ret = 'Unknown Error';
            if (ex) {
                try {
                    if (ex.toString) {
                        ret = ex.toString();
                    }
                    ret = ret + ' ' + JSON.stringify(ex);
                }
                catch (otherEx) {
                    ret = 'Unexpected Error';
                }
            }
            return ret;
        }
        function _tryLog(ex) {
            var message = toLogMessage(ex);
            OfficeExtension.Utility.log(message);
        }
        function notifyActionHandlerReady() {
            var context = new OfficeExtension.ClientRequestContext();
            var addinInternalService = OfficeCore.AddinInternalService.newObject(context);
            context._customData = 'WacPartition';
            addinInternalService.notifyActionHandlerReady();
            return context.sync();
        }
        function handlerOnReadyInternal() {
            try {
                Microsoft.Office.WebExtension.onReadyInternal()
                    .then(function () {
                    return init();
                })
                    .then(function () {
                    var hostInfo = OSF._OfficeAppFactory.getHostInfo();
                    if (hostInfo.hostPlatform === "web" && hostInfo.hostType !== "word" && hostInfo.hostType !== "excel") {
                        return;
                    }
                    else {
                        return notifyActionHandlerReady();
                    }
                });
            }
            catch (ex) {
            }
        }
        function initFromHostBridge(hostBridge) {
            hostBridge.addHostMessageHandler(function (bridgeMessage) {
                if (bridgeMessage.type === 3) {
                    _handleMessage(bridgeMessage.message);
                }
            });
        }
        function initOnce() {
            OfficeExtension.Utility.log('ActionProxy.initOnce');
            if (typeof (Office.actions) != 'undefined') {
                setActionAssociation(Office.actions._association);
            }
            if (typeof (document) !== 'undefined') {
                if (document.readyState && document.readyState !== 'loading') {
                    OfficeExtension.Utility.log('ActionProxy.initOnce: document.readyState is not loading state');
                    handlerOnReadyInternal();
                }
                else if (document.addEventListener) {
                    document.addEventListener("DOMContentLoaded", function () {
                        OfficeExtension.Utility.log('ActionProxy.initOnce: DOMContentLoaded event triggered');
                        handlerOnReadyInternal();
                    });
                }
            }
            OfficeExtension.HostBridge.onInited(function (hostBridge) {
                initFromHostBridge(hostBridge);
            });
        }
        initOnce();
    })(actionProxy || (actionProxy = {}));
})(Office || (Office = {}));
var Office;
(function (Office) {
    var dialogPage;
    (function (dialogPage_1) {
        function close() {
            return __awaiter(this, void 0, void 0, function () {
                var context, dialogPage;
                return __generator(this, function (_a) {
                    switch (_a.label) {
                        case 0:
                            context = new OfficeCore.RequestContext();
                            dialogPage = OfficeCore.DialogPage.newObject(context);
                            dialogPage.close();
                            return [4, context.sync()];
                        case 1:
                            _a.sent();
                            return [2];
                    }
                });
            });
        }
        dialogPage_1.close = close;
        function readyToShow() {
            return __awaiter(this, void 0, void 0, function () {
                var context, dialogPage;
                return __generator(this, function (_a) {
                    switch (_a.label) {
                        case 0:
                            context = new OfficeCore.RequestContext();
                            dialogPage = OfficeCore.DialogPage.newObject(context);
                            dialogPage.readyToShow();
                            return [4, context.sync()];
                        case 1:
                            _a.sent();
                            return [2];
                    }
                });
            });
        }
        dialogPage_1.readyToShow = readyToShow;
        function onShow(callback) {
            return __awaiter(this, void 0, void 0, function () {
                var context, dialogPage, removeListener;
                return __generator(this, function (_a) {
                    switch (_a.label) {
                        case 0:
                            context = new OfficeCore.RequestContext();
                            dialogPage = OfficeCore.DialogPage.newObject(context);
                            dialogPage.onOnShowEvent.add(callback);
                            removeListener = function () {
                                dialogPage.onOnShowEvent.remove(callback);
                                return null;
                            };
                            return [4, context.sync()];
                        case 1:
                            _a.sent();
                            return [2, removeListener];
                    }
                });
            });
        }
        dialogPage_1.onShow = onShow;
        function sendMessageToHost(message) {
            return __awaiter(this, void 0, void 0, function () {
                var context, dialogPage;
                return __generator(this, function (_a) {
                    switch (_a.label) {
                        case 0:
                            context = new OfficeCore.RequestContext();
                            dialogPage = OfficeCore.DialogPage.newObject(context);
                            dialogPage.sendMessageToHost(message);
                            return [4, context.sync()];
                        case 1:
                            _a.sent();
                            return [2];
                    }
                });
            });
        }
        dialogPage_1.sendMessageToHost = sendMessageToHost;
    })(dialogPage = Office.dialogPage || (Office.dialogPage = {}));
})(Office || (Office = {}));
var __extends = (this && this.__extends) || (function () {
    var extendStatics = Object.setPrototypeOf ||
        ({ __proto__: [] } instanceof Array && function (d, b) { d.__proto__ = b; }) ||
        function (d, b) { for (var p in b) if (b.hasOwnProperty(p)) d[p] = b[p]; };
    return function (d, b) {
        extendStatics(d, b);
        function __() { this.constructor = d; }
        d.prototype = b === null ? Object.create(b) : (__.prototype = b.prototype, new __());
    };
})();
var Word;
(function (Word) {
    function _normalizeSearchOptions(context, searchOptions) {
        if (OfficeExtension.Utility.isNullOrUndefined(searchOptions)) {
            return null;
        }
        if (typeof (searchOptions) != "object") {
            OfficeExtension.Utility.throwError(OfficeExtension.ResourceStrings.invalidArgument, "searchOptions");
        }
        if (searchOptions instanceof Word.SearchOptions) {
            return searchOptions;
        }
        var newSearchOptions = Word.SearchOptions.newObject(context);
        for (var property in searchOptions) {
            if (searchOptions.hasOwnProperty(property)) {
                newSearchOptions[property] = searchOptions[property];
            }
        }
        return newSearchOptions;
    }
    var _hostName = "Word";
    var _defaultApiSetName = "WordApi";
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
    var _calculateApiFlags = OfficeExtension.CommonUtility.calculateApiFlags;
    var AnnotationCollectionCustom = (function () {
        function AnnotationCollectionCustom() {
        }
        AnnotationCollectionCustom.prototype.getDocument = function () {
            if (_isNullOrUndefined(this._document)) {
                this._document = this.context.document;
            }
            return this._document;
        };
        AnnotationCollectionCustom.prototype.getReferenceId = function () {
            if (_isNullOrUndefined(this._refId)) {
                this._refId = this._ReferenceId;
            }
            return this._refId;
        };
        AnnotationCollectionCustom.prototype._RegisterAddedEvent = function () {
            this.getDocument()._RegisterEvent(this.getReferenceId(), "AnnotationAdded");
        };
        AnnotationCollectionCustom.prototype._UnregisterAddedEvent = function () {
            this.getDocument()._UnregisterEvent(this.getReferenceId(), "AnnotationAdded");
        };
        AnnotationCollectionCustom.prototype._RegisterChangedEvent = function () {
            this.getDocument()._RegisterEvent(this.getReferenceId(), "AnnotationChanged");
        };
        AnnotationCollectionCustom.prototype._UnregisterChangedEvent = function () {
            this.getDocument()._UnregisterEvent(this.getReferenceId(), "AnnotationChanged");
        };
        AnnotationCollectionCustom.prototype._RegisterDeletedEvent = function () {
            this.getDocument()._RegisterEvent(this.getReferenceId(), "AnnotationDeleted");
        };
        AnnotationCollectionCustom.prototype._UnregisterDeletedEvent = function () {
            this.getDocument()._UnregisterEvent(this.getReferenceId(), "AnnotationDeleted");
        };
        return AnnotationCollectionCustom;
    }());
    Word.AnnotationCollectionCustom = AnnotationCollectionCustom;
    var _CC;
    (function (_CC) {
        function AnnotationCollection_AnnotationAdded_EventArgsTransform(thisObj, args) {
            var evt = {
                eventType: Word.EventType.annotationAdded,
                annotation: OfficeExtension.BatchApiHelper.createObjectFromReferenceId(Word.Annotation, thisObj.context, args)
            };
            evt.annotation.load();
            return evt;
        }
        _CC.AnnotationCollection_AnnotationAdded_EventArgsTransform = AnnotationCollection_AnnotationAdded_EventArgsTransform;
        function AnnotationCollection_AnnotationChanged_EventArgsTransform(thisObj, args) {
            var evt = {
                eventType: Word.EventType.annotationChanged,
                annotation: OfficeExtension.BatchApiHelper.createObjectFromReferenceId(Word.Annotation, thisObj.context, args)
            };
            evt.annotation.load();
            return evt;
        }
        _CC.AnnotationCollection_AnnotationChanged_EventArgsTransform = AnnotationCollection_AnnotationChanged_EventArgsTransform;
        function AnnotationCollection_AnnotationDeleted_EventArgsTransform(thisObj, args) {
            var evt = {
                eventType: Word.EventType.annotationDeleted,
                annotation: OfficeExtension.BatchApiHelper.createObjectFromReferenceId(Word.Annotation, thisObj.context, args)
            };
            evt.annotation.load();
            return evt;
        }
        _CC.AnnotationCollection_AnnotationDeleted_EventArgsTransform = AnnotationCollection_AnnotationDeleted_EventArgsTransform;
    })(_CC = Word._CC || (Word._CC = {}));
    (function (_CC) {
        function Body_Search(thisObj, searchText, searchOptions) {
            searchOptions = _normalizeSearchOptions(thisObj.context, searchOptions);
            var result = _createMethodObject(Word.RangeCollection, thisObj, "Search", 1, [searchText, searchOptions], true, false, null, 4);
            return { handled: true, result: result };
        }
        _CC.Body_Search = Body_Search;
    })(_CC = Word._CC || (Word._CC = {}));
    var ContentControlCustom = (function () {
        function ContentControlCustom() {
        }
        ContentControlCustom.prototype.getDocument = function () {
            if (_isNullOrUndefined(this._document)) {
                this._document = this.context.document;
            }
            return this._document;
        };
        ContentControlCustom.prototype.getReferenceId = function () {
            if (_isNullOrUndefined(this._refId)) {
                this._refId = this._ReferenceId;
            }
            return this._refId;
        };
        ContentControlCustom.prototype._RegisterDataChangedEvent = function () {
            this.getDocument()._RegisterEvent(this.getReferenceId(), "ContentControlDataChanged");
        };
        ContentControlCustom.prototype._UnregisterDataChangedEvent = function () {
            this.getDocument()._UnregisterEvent(this.getReferenceId(), "ContentControlDataChanged");
        };
        ContentControlCustom.prototype._RegisterDeletedEvent = function () {
            this.getDocument()._RegisterEvent(this.getReferenceId(), "ContentControlDeleted");
        };
        ContentControlCustom.prototype._UnregisterDeletedEvent = function () {
            this.getDocument()._UnregisterEvent(this.getReferenceId(), "ContentControlDeleted");
        };
        ContentControlCustom.prototype._RegisterSelectionChangedEvent = function () {
            this.getDocument()._RegisterEvent(this.getReferenceId(), "ContentControlSelectionChanged");
        };
        ContentControlCustom.prototype._UnregisterSelectionChangedEvent = function () {
            this.getDocument()._UnregisterEvent(this.getReferenceId(), "ContentControlSelectionChanged");
        };
        return ContentControlCustom;
    }());
    Word.ContentControlCustom = ContentControlCustom;
    (function (_CC) {
        function ContentControl_Search(thisObj, searchText, searchOptions) {
            searchOptions = _normalizeSearchOptions(thisObj.context, searchOptions);
            var result = _createMethodObject(Word.RangeCollection, thisObj, "Search", 1, [searchText, searchOptions], true, false, null, 4);
            return { handled: true, result: result };
        }
        _CC.ContentControl_Search = ContentControl_Search;
        function ContentControl_DataChanged_EventArgsTransform(thisObj, args) {
            var evt = {
                eventType: Word.EventType.contentControlDataChanged,
                contentControl: thisObj
            };
            return evt;
        }
        _CC.ContentControl_DataChanged_EventArgsTransform = ContentControl_DataChanged_EventArgsTransform;
        function ContentControl_Deleted_EventArgsTransform(thisObj, args) {
            var evt = {
                eventType: Word.EventType.contentControlDeleted,
                contentControl: thisObj
            };
            return evt;
        }
        _CC.ContentControl_Deleted_EventArgsTransform = ContentControl_Deleted_EventArgsTransform;
        function ContentControl_SelectionChanged_EventArgsTransform(thisObj, args) {
            var evt = {
                eventType: Word.EventType.contentControlSelectionChanged,
                contentControl: thisObj
            };
            return evt;
        }
        _CC.ContentControl_SelectionChanged_EventArgsTransform = ContentControl_SelectionChanged_EventArgsTransform;
    })(_CC = Word._CC || (Word._CC = {}));
    (function (_CC) {
        function CustomProperty_HandleResult(thisObj, value) {
            if (!_isUndefined(value["Value"]) && !_isUndefined(value["Type"]) && value["Type"] == "Date") {
                value["Value"] = new Date(value["Value"]);
            }
            ;
        }
        _CC.CustomProperty_HandleResult = CustomProperty_HandleResult;
    })(_CC = Word._CC || (Word._CC = {}));
    var DocumentCustom = (function () {
        function DocumentCustom() {
        }
        DocumentCustom.prototype._RegisterContentControlAddedEvent = function () {
            var document = this;
            document._RegisterEvent(document._ReferenceId, "ContentControlAdded");
        };
        DocumentCustom.prototype._UnregisterContentControlAddedEvent = function () {
            var document = this;
            document._RegisterEvent(document._ReferenceId, "ContentControlAdded");
        };
        return DocumentCustom;
    }());
    Word.DocumentCustom = DocumentCustom;
    (function (_CC) {
        function Document_ContentControlAdded_EventArgsTransform(thisObj, args) {
            var evt = {
                eventType: Word.EventType.contentControlAdded,
                contentControl: OfficeExtension.BatchApiHelper.createObjectFromReferenceId(Word.ContentControl, thisObj.context, args)
            };
            evt.contentControl.load();
            return evt;
        }
        _CC.Document_ContentControlAdded_EventArgsTransform = Document_ContentControlAdded_EventArgsTransform;
    })(_CC = Word._CC || (Word._CC = {}));
    (function (_CC) {
        function Paragraph_Search(thisObj, searchText, searchOptions) {
            searchOptions = _normalizeSearchOptions(thisObj.context, searchOptions);
            var result = _createMethodObject(Word.RangeCollection, thisObj, "Search", 1, [searchText, searchOptions], true, false, null, 4);
            return { handled: true, result: result };
        }
        _CC.Paragraph_Search = Paragraph_Search;
    })(_CC = Word._CC || (Word._CC = {}));
    (function (_CC) {
        function Range_Search(thisObj, searchText, searchOptions) {
            searchOptions = _normalizeSearchOptions(thisObj.context, searchOptions);
            var result = _createMethodObject(Word.RangeCollection, thisObj, "Search", 1, [searchText, searchOptions], true, false, null, 4);
            return { handled: true, result: result };
        }
        _CC.Range_Search = Range_Search;
    })(_CC = Word._CC || (Word._CC = {}));
    var SearchOptionsCustom = (function () {
        function SearchOptionsCustom() {
        }
        Object.defineProperty(SearchOptionsCustom.prototype, "matchWildCards", {
            get: function () {
                _throwIfNotLoaded("matchWildCards", this.m_matchWildcards);
                return this.m_matchWildcards;
            },
            set: function (value) {
                this.m_matchWildcards = value;
                _invokeSetProperty(this, "MatchWildCards", value, 0);
            },
            enumerable: true,
            configurable: true
        });
        return SearchOptionsCustom;
    }());
    Word.SearchOptionsCustom = SearchOptionsCustom;
    var SettingCustom = (function () {
        function SettingCustom() {
        }
        SettingCustom.replaceStringDateWithDate = function (value) {
            var strValue = JSON.stringify(value);
            value = JSON.parse(strValue, function dateReviver(k, v) {
                var d;
                if (typeof v === 'string' && v && v.length > 6 && v.slice(0, 5) === SettingCustom.DateJSONPrefix && v.slice(-1) === SettingCustom.DateJSONSuffix) {
                    d = new Date(parseInt(v.slice(5, -1)));
                    if (d) {
                        return d;
                    }
                }
                return v;
            });
            return value;
        };
        SettingCustom.replaceDateWithStringDate = function (value) {
            var strValue = JSON.stringify(value, function dateReplacer(k, v) {
                return (this[k] instanceof Date) ? (SettingCustom.DateJSONPrefix + this[k].getTime() + SettingCustom.DateJSONSuffix) : v;
            });
            value = JSON.parse(strValue);
            return value;
        };
        SettingCustom.DateJSONPrefix = "Date(";
        SettingCustom.DateJSONSuffix = ")";
        return SettingCustom;
    }());
    Word.SettingCustom = SettingCustom;
    (function (_CC) {
        function Setting_HandleResult(thisObj, value) {
            function dateReviver(key, val) {
                var re = /^(\d{4})-(\d{2})-(\d{2})T(\d{2}):(\d{2}):(\d{2}(?:\.\d*)?)Z$/;
                if (re.exec(val))
                    return new Date(val);
                return val;
            }
            if (!_isUndefined(value["Value"]) && typeof value["Value"] === "string") {
                var newValue = JSON.parse(value["Value"], dateReviver);
                value["Value"] = SettingCustom.replaceStringDateWithDate(newValue);
            }
        }
        _CC.Setting_HandleResult = Setting_HandleResult;
        function Setting_Value_Set(thisObj, value) {
            var newValue = JSON.stringify(SettingCustom.replaceDateWithStringDate(value));
            if (newValue !== null) {
                this.m_value = newValue;
                _invokeSetProperty(thisObj, "Value", newValue, 0);
                return { handled: true };
            }
        }
        _CC.Setting_Value_Set = Setting_Value_Set;
    })(_CC = Word._CC || (Word._CC = {}));
    (function (_CC) {
        function SettingCollection_Add(thisObj, key, value) {
            var newValue = JSON.stringify(SettingCustom.replaceDateWithStringDate(value));
            if (newValue !== null) {
                var result = _createMethodObject(Word.Setting, thisObj, "Add", 0, [key, newValue], false, false, null, 0);
                return { handled: true, result: result };
            }
            return { handled: false, result: null };
        }
        _CC.SettingCollection_Add = SettingCollection_Add;
    })(_CC = Word._CC || (Word._CC = {}));
    (function (_CC) {
        function Table_Search(thisObj, searchText, searchOptions) {
            searchOptions = _normalizeSearchOptions(thisObj.context, searchOptions);
            var result = _createMethodObject(Word.RangeCollection, thisObj, "Search", 1, [searchText, searchOptions], true, false, null, 4);
            return { handled: true, result: result };
        }
        _CC.Table_Search = Table_Search;
    })(_CC = Word._CC || (Word._CC = {}));
    (function (_CC) {
        function TableRow_Search(thisObj, searchText, searchOptions) {
            searchOptions = _normalizeSearchOptions(thisObj.context, searchOptions);
            var result = _createMethodObject(Word.RangeCollection, thisObj, "Search", 1, [searchText, searchOptions], true, false, null, 4);
            return { handled: true, result: result };
        }
        _CC.TableRow_Search = TableRow_Search;
    })(_CC = Word._CC || (Word._CC = {}));
    var ErrorCodes;
    (function (ErrorCodes) {
        ErrorCodes["accessDenied"] = "AccessDenied";
        ErrorCodes["generalException"] = "GeneralException";
        ErrorCodes["invalidArgument"] = "InvalidArgument";
        ErrorCodes["itemNotFound"] = "ItemNotFound";
        ErrorCodes["notImplemented"] = "NotImplemented";
        ErrorCodes["searchDialogIsOpen"] = "SearchDialogIsOpen";
        ErrorCodes["searchStringInvalidOrTooLong"] = "SearchStringInvalidOrTooLong";
    })(ErrorCodes = Word.ErrorCodes || (Word.ErrorCodes = {}));
    var Interfaces;
    (function (Interfaces) {
    })(Interfaces = Word.Interfaces || (Word.Interfaces = {}));
    var _libraryMetadataWdJscomApi = { "version": "1.0.0",
        "name": "Word",
        "defaultApiSetName": "WordApi",
        "hostName": "Word",
        "apiSets": [["1.3"], ["1.2"], ["1.4"], ["1.1", "WordOnline"], ["1.3", "WordApiHiddenDocument"], ["1.4", "WordApiHiddenDocument"]],
        "strings": ["Annotation", "AnnotationCollection", "Application", "Body", "ContentControl", "ContentControlCollection", "CustomProperty", "CustomPropertyCollection", "CustomXmlPart", "CustomXmlPartCollection", "CustomXmlPartScopedCollection", "Document", "DocumentCreated", "DocumentProperties", "Font", "InlinePicture", "InlinePictureCollection", "List", "ListCollection", "ListItem", "Paragraph", "ParagraphCollection", "Range", "RangeCollection", "SearchOptions", "Section", "SectionCollection", "Setting", "SettingCollection", "Table", "TableCollection", "TableRow", "TableRowCollection", "TableCell", "TableCellCollection", "TableBorder", "id", "_ReferenceId", "_KeepReference", "delete", "getItem", "getFirst", "getFirstOrNullObject", "this._ReferenceId", "_RegisterDeletedEvent", "_UnregisterDeletedEvent", "style", "text", "type", "styleBuiltIn", "paragraphs", "contentControls", "parentContentControl", "font", "inlinePictures", "parentBody", "lists", "tables", "parentContentControlOrNullObject", "insertBreak", "clear", "getHtml", "getOoxml", "select", "insertText", "insertHtml", "insertOoxml", "insertParagraph", "insertContentControl", "insertFileFromBase64", "insertInlinePictureFromBase64", "search", "getRange", "insertTable", "title", "color", "parentTableCell", "parentTable", "parentTableCellOrNullObject", "parentTableOrNullObject", "split", "getTextRanges", "getById", "getByIdOrNullObject", "key", "value", "_Id", "getCount", "deleteAll", "add", "getItemOrNullObject", "saved", "sections", "body", "properties", "settings", "customXmlParts", "save", "deleteBookmark", "getMetadata", "setMetadata", "getBookmarkRange", "getBookmarkRangeOrNullObject", "hyperlink", "width", "getNext", "getNextOrNullObject", "_GetItem", "alignment", "values", "shadingColor", "horizontalAlignment", "verticalAlignment", "getCellPadding", "setCellPadding", "getBorder", "rowIndex", "insertRows"],
        "enumTypes": [["NumericEventType", [], { "contentControlDeleted": 0, "contentControlSelectionChanged": 1, "contentControlDataChanged": 2, "contentControlAdded": 3, "annotationAdded": 4, "annotationChanged": 5, "annotationDeleted": 6 }],
            ["EventType", ["contentControlDeleted", "contentControlSelectionChanged", "contentControlDataChanged", "contentControlAdded", "annotationAdded", "annotationChanged", "annotationDeleted"]],
            ["ContentControlType", ["unknown", "richTextInline", "richTextParagraphs", "richTextTableCell", "richTextTableRow", "richTextTable", "plainTextInline", "plainTextParagraph", "picture", "buildingBlockGallery", "checkBox", "comboBox", "dropDownList", "datePicker", "repeatingSection", "richText", "plainText"]],
            ["ContentControlAppearance", ["boundingBox", "tags", "hidden"]],
            ["UnderlineType", ["mixed", "none", "hidden", "dotLine", "single", "word", "double", "thick", "dotted", "dottedHeavy", "dashLine", "dashLineHeavy", "dashLineLong", "dashLineLongHeavy", "dotDashLine", "dotDashLineHeavy", "twoDotDashLine", "twoDotDashLineHeavy", "wave", "waveHeavy", "waveDouble"]],
            ["BreakType", ["page", "next", "sectionNext", "sectionContinuous", "sectionEven", "sectionOdd", "line"]],
            ["InsertLocation", ["before", "after", "start", "end", "replace"]],
            ["Alignment", ["mixed", "unknown", "left", "centered", "right", "justified"]],
            ["HeaderFooterType", ["primary", "firstPage", "evenPages"]],
            ["BodyType", ["unknown", "mainDoc", "section", "header", "footer", "tableCell"]],
            ["SelectionMode", ["select", "start", "end"]],
            ["ImageFormat", ["unsupported", "undefined", "bmp", "jpeg", "gif", "tiff", "png", "icon", "exif", "wmf", "emf", "pict", "pdf", "svg"]],
            ["RangeLocation", ["whole", "start", "end", "before", "after", "content"]],
            ["LocationRelation", ["unrelated", "equal", "containsStart", "containsEnd", "contains", "insideStart", "insideEnd", "inside", "adjacentBefore", "overlapsBefore", "before", "adjacentAfter", "overlapsAfter", "after"]],
            ["BorderLocation", ["top", "left", "bottom", "right", "insideHorizontal", "insideVertical", "inside", "outside", "all"]],
            ["CellPaddingLocation", ["top", "left", "bottom", "right"]],
            ["BorderType", ["mixed", "none", "single", "double", "dotted", "dashed", "dotDashed", "dot2Dashed", "triple", "thinThickSmall", "thickThinSmall", "thinThickThinSmall", "thinThickMed", "thickThinMed", "thinThickThinMed", "thinThickLarge", "thickThinLarge", "thinThickThinLarge", "wave", "doubleWave", "dashedSmall", "dashDotStroked", "threeDEmboss", "threeDEngrave"]],
            ["VerticalAlignment", ["mixed", "top", "center", "bottom"]],
            ["ListLevelType", ["bullet", "number", "picture"]],
            ["ListBullet", ["custom", "solid", "hollow", "square", "diamonds", "arrow", "checkmark"]],
            ["ListNumbering", ["none", "arabic", "upperRoman", "lowerRoman", "upperLetter", "lowerLetter"]],
            ["Style", ["other", "normal", "heading1", "heading2", "heading3", "heading4", "heading5", "heading6", "heading7", "heading8", "heading9", "toc1", "toc2", "toc3", "toc4", "toc5", "toc6", "toc7", "toc8", "toc9", "footnoteText", "header", "footer", "caption", "footnoteReference", "endnoteReference", "endnoteText", "title", "subtitle", "hyperlink", "strong", "emphasis", "noSpacing", "listParagraph", "quote", "intenseQuote", "subtleEmphasis", "intenseEmphasis", "subtleReference", "intenseReference", "bookTitle", "bibliography", "tocHeading", "tableGrid", "plainTable1", "plainTable2", "plainTable3", "plainTable4", "plainTable5", "tableGridLight", "gridTable1Light", "gridTable1Light_Accent1", "gridTable1Light_Accent2", "gridTable1Light_Accent3", "gridTable1Light_Accent4", "gridTable1Light_Accent5", "gridTable1Light_Accent6", "gridTable2", "gridTable2_Accent1", "gridTable2_Accent2", "gridTable2_Accent3", "gridTable2_Accent4", "gridTable2_Accent5", "gridTable2_Accent6", "gridTable3", "gridTable3_Accent1", "gridTable3_Accent2", "gridTable3_Accent3", "gridTable3_Accent4", "gridTable3_Accent5", "gridTable3_Accent6", "gridTable4", "gridTable4_Accent1", "gridTable4_Accent2", "gridTable4_Accent3", "gridTable4_Accent4", "gridTable4_Accent5", "gridTable4_Accent6", "gridTable5Dark", "gridTable5Dark_Accent1", "gridTable5Dark_Accent2", "gridTable5Dark_Accent3", "gridTable5Dark_Accent4", "gridTable5Dark_Accent5", "gridTable5Dark_Accent6", "gridTable6Colorful", "gridTable6Colorful_Accent1", "gridTable6Colorful_Accent2", "gridTable6Colorful_Accent3", "gridTable6Colorful_Accent4", "gridTable6Colorful_Accent5", "gridTable6Colorful_Accent6", "gridTable7Colorful", "gridTable7Colorful_Accent1", "gridTable7Colorful_Accent2", "gridTable7Colorful_Accent3", "gridTable7Colorful_Accent4", "gridTable7Colorful_Accent5", "gridTable7Colorful_Accent6", "listTable1Light", "listTable1Light_Accent1", "listTable1Light_Accent2", "listTable1Light_Accent3", "listTable1Light_Accent4", "listTable1Light_Accent5", "listTable1Light_Accent6", "listTable2", "listTable2_Accent1", "listTable2_Accent2", "listTable2_Accent3", "listTable2_Accent4", "listTable2_Accent5", "listTable2_Accent6", "listTable3", "listTable3_Accent1", "listTable3_Accent2", "listTable3_Accent3", "listTable3_Accent4", "listTable3_Accent5", "listTable3_Accent6", "listTable4", "listTable4_Accent1", "listTable4_Accent2", "listTable4_Accent3", "listTable4_Accent4", "listTable4_Accent5", "listTable4_Accent6", "listTable5Dark", "listTable5Dark_Accent1", "listTable5Dark_Accent2", "listTable5Dark_Accent3", "listTable5Dark_Accent4", "listTable5Dark_Accent5", "listTable5Dark_Accent6", "listTable6Colorful", "listTable6Colorful_Accent1", "listTable6Colorful_Accent2", "listTable6Colorful_Accent3", "listTable6Colorful_Accent4", "listTable6Colorful_Accent5", "listTable6Colorful_Accent6", "listTable7Colorful", "listTable7Colorful_Accent1", "listTable7Colorful_Accent2", "listTable7Colorful_Accent3", "listTable7Colorful_Accent4", "listTable7Colorful_Accent5", "listTable7Colorful_Accent6"]],
            ["DocumentPropertyType", ["string", "number", "date", "boolean"]],
            ["TapObjectType", ["chart", "smartArt", "table", "image", "slide", "text"], { "ole": "OLE" }],
            ["FileContentFormat", ["base64", "html", "ooxml"]],
            ["AnnotationParentType", ["none", "document", "paragraph", "annotation"]],
            ["AnnotationState", ["undefined", "created", "sent", "duplicated", "seen", "tried", "kept", "rejected"]]],
        "clientObjectTypes": [[1,
                2,
                [["content",
                        3],
                    [37,
                        3],
                    [38,
                        2],
                    ["_State"]],
                0,
                [[39,
                        0,
                        2,
                        0,
                        4],
                    ["getParentType",
                        0,
                        2,
                        0,
                        4],
                    [40]],
                [["getParentAsParagraph",
                        21,
                        0,
                        2,
                        0,
                        4],
                    ["getParentAsAnnotation",
                        1,
                        0,
                        2,
                        0,
                        4]]],
            [2,
                7,
                [[38,
                        2]],
                0,
                [[39,
                        0,
                        2,
                        0,
                        4],
                    ["refresh",
                        0,
                        2,
                        0,
                        4]],
                [[41,
                        1,
                        1,
                        18,
                        0,
                        4],
                    [42,
                        1,
                        0,
                        2,
                        0,
                        4],
                    [43,
                        1,
                        0,
                        2,
                        0,
                        4]],
                0,
                1,
                [["AnnotationAdded",
                        2,
                        0,
                        "NumericEventType.annotationAdded",
                        44,
                        "_RegisterAddedEvent",
                        "_UnregisterAddedEvent"],
                    ["AnnotationChanged",
                        2,
                        0,
                        "NumericEventType.annotationChanged",
                        44,
                        "_RegisterChangedEvent",
                        "_UnregisterChangedEvent"],
                    ["AnnotationDeleted",
                        2,
                        0,
                        "NumericEventType.annotationDeleted",
                        44,
                        45,
                        46]]],
            [3,
                0,
                0,
                0,
                [["isTapEnabled",
                        0,
                        2,
                        0,
                        5],
                    ["getSharePointTenantRoot",
                        0,
                        2,
                        0,
                        5],
                    ["getEnterpriseUserInfo",
                        0,
                        2,
                        0,
                        5],
                    ["getMruFriendlyPath",
                        1,
                        2,
                        0,
                        5],
                    ["launchFileUrlInOfficeApp",
                        2,
                        2,
                        0,
                        5]],
                [["createDocument",
                        13,
                        1,
                        2,
                        0,
                        4]],
                0,
                0,
                0,
                "Microsoft.WordServices.Application",
                4],
            [4,
                2,
                [[38,
                        2],
                    [47,
                        1],
                    [48,
                        3],
                    [49,
                        3,
                        1],
                    [50,
                        1,
                        1]],
                [[51,
                        22,
                        19,
                        0,
                        0,
                        4],
                    [52,
                        6,
                        19,
                        0,
                        0,
                        4],
                    [53,
                        5,
                        2,
                        0,
                        0,
                        4],
                    [54,
                        15,
                        35,
                        0,
                        0,
                        4],
                    [55,
                        17,
                        19,
                        0,
                        0,
                        4],
                    [56,
                        4,
                        2,
                        1,
                        0,
                        4],
                    [57,
                        19,
                        19,
                        1,
                        0,
                        4],
                    [58,
                        31,
                        19,
                        1,
                        0,
                        4],
                    ["parentSection",
                        26,
                        2,
                        1,
                        0,
                        4],
                    [59,
                        5,
                        2,
                        1,
                        0,
                        4],
                    ["parentBodyOrNullObject",
                        4,
                        2,
                        1,
                        0,
                        4],
                    ["parentSectionOrNullObject",
                        26,
                        2,
                        1,
                        0,
                        4]],
                [[39,
                        0,
                        2,
                        0,
                        4],
                    [60,
                        2],
                    [61],
                    [62,
                        0,
                        2,
                        0,
                        4],
                    [63,
                        0,
                        2,
                        0,
                        4],
                    [64,
                        1,
                        2,
                        0,
                        4]],
                [[65,
                        23,
                        2,
                        8],
                    [66,
                        23,
                        2,
                        8],
                    [67,
                        23,
                        2,
                        8],
                    [68,
                        21,
                        2,
                        8],
                    [69,
                        5,
                        0,
                        8],
                    [70,
                        23,
                        2,
                        8],
                    [71,
                        16,
                        2,
                        8,
                        2],
                    [72,
                        24,
                        2,
                        7,
                        0,
                        4],
                    [73,
                        23,
                        1,
                        2,
                        1,
                        4],
                    [74,
                        30,
                        4,
                        8,
                        1]]],
            [5,
                6,
                [[37,
                        3],
                    [38,
                        2],
                    [75,
                        1],
                    ["tag",
                        1],
                    ["placeholderText",
                        1],
                    [49,
                        3],
                    ["appearance",
                        1],
                    [76,
                        1],
                    ["removeWhenEdited",
                        1],
                    ["cannotDelete",
                        1],
                    ["cannotEdit",
                        1],
                    [47,
                        1],
                    [48,
                        3],
                    ["subtype",
                        3,
                        1],
                    [50,
                        1,
                        1]],
                [[54,
                        15,
                        35,
                        0,
                        0,
                        4],
                    [51,
                        22,
                        19,
                        0,
                        0,
                        4],
                    [52,
                        6,
                        19,
                        0,
                        0,
                        4],
                    [53,
                        5,
                        2,
                        0,
                        0,
                        4],
                    [55,
                        17,
                        19,
                        0,
                        0,
                        4],
                    [57,
                        19,
                        19,
                        1,
                        0,
                        4],
                    [58,
                        31,
                        19,
                        1,
                        0,
                        4],
                    [77,
                        34,
                        2,
                        1,
                        0,
                        4],
                    [78,
                        30,
                        2,
                        1,
                        0,
                        4],
                    [56,
                        4,
                        2,
                        1,
                        0,
                        4],
                    [59,
                        5,
                        2,
                        1,
                        0,
                        4],
                    [79,
                        34,
                        2,
                        1,
                        0,
                        4],
                    [80,
                        30,
                        2,
                        1,
                        0,
                        4]],
                [[39,
                        0,
                        2,
                        0,
                        4],
                    [60,
                        2],
                    [61],
                    [40,
                        1],
                    [64,
                        1,
                        2,
                        0,
                        4],
                    [62,
                        0,
                        2,
                        0,
                        4],
                    [63,
                        0,
                        2,
                        0,
                        4]],
                [[65,
                        23,
                        2,
                        8],
                    [66,
                        23,
                        2,
                        8],
                    [67,
                        23,
                        2,
                        8],
                    [70,
                        23,
                        2,
                        8],
                    [68,
                        21,
                        2,
                        8],
                    [71,
                        16,
                        2,
                        8,
                        2],
                    [72,
                        24,
                        2,
                        7,
                        0,
                        4],
                    [73,
                        23,
                        1,
                        2,
                        1,
                        4],
                    [81,
                        24,
                        4,
                        6,
                        1,
                        4],
                    [74,
                        30,
                        4,
                        8,
                        1],
                    [82,
                        24,
                        2,
                        6,
                        1,
                        4]],
                0,
                0,
                [["DataChanged",
                        2,
                        3,
                        "NumericEventType.contentControlDataChanged",
                        44,
                        "_RegisterDataChangedEvent",
                        "_UnregisterDataChangedEvent"],
                    ["Deleted",
                        2,
                        3,
                        "NumericEventType.contentControlDeleted",
                        44,
                        45,
                        46],
                    ["SelectionChanged",
                        2,
                        3,
                        "NumericEventType.contentControlSelectionChanged",
                        44,
                        "_RegisterSelectionChangedEvent",
                        "_UnregisterSelectionChangedEvent"]]],
            [6,
                3,
                [[38,
                        2]],
                0,
                [[39,
                        0,
                        2,
                        0,
                        4]],
                [[41,
                        5,
                        1,
                        18,
                        0,
                        4],
                    [83,
                        5,
                        1,
                        2,
                        0,
                        4],
                    ["getByTitle",
                        6,
                        1,
                        6,
                        0,
                        4],
                    ["getByTag",
                        6,
                        1,
                        6,
                        0,
                        4],
                    ["getByTypes",
                        6,
                        1,
                        6,
                        1,
                        4],
                    [42,
                        5,
                        0,
                        2,
                        1,
                        4],
                    [84,
                        5,
                        1,
                        2,
                        1,
                        4],
                    [43,
                        5,
                        0,
                        2,
                        1,
                        4]],
                0,
                5],
            [7,
                10,
                [[38,
                        2],
                    [85,
                        3],
                    [86,
                        1],
                    [49,
                        3],
                    [87,
                        2]],
                0,
                [[39,
                        0,
                        2,
                        0,
                        4],
                    [40]]],
            [8,
                3,
                [[38,
                        2]],
                0,
                [[39,
                        0,
                        2,
                        0,
                        4],
                    [88,
                        0,
                        2,
                        0,
                        4],
                    [89]],
                [[41,
                        7,
                        1,
                        18,
                        0,
                        4],
                    [90,
                        7,
                        2,
                        8],
                    [91,
                        7,
                        1,
                        2,
                        0,
                        4]],
                0,
                7],
            [9,
                2,
                [[38,
                        2],
                    [37,
                        3],
                    ["namespaceUri",
                        3]],
                0,
                [[39,
                        0,
                        2,
                        0,
                        4],
                    [40],
                    ["getXml",
                        0,
                        2,
                        0,
                        4],
                    ["setXml",
                        1],
                    ["query",
                        2],
                    ["insertElement",
                        4],
                    ["updateElement",
                        3],
                    ["deleteElement",
                        2],
                    ["insertAttribute",
                        4],
                    ["updateAttribute",
                        4],
                    ["deleteAttribute",
                        3]]],
            [10,
                3,
                [[38,
                        2]],
                0,
                [[39,
                        0,
                        2,
                        0,
                        4],
                    [88,
                        0,
                        2,
                        0,
                        4]],
                [[41,
                        9,
                        1,
                        18,
                        0,
                        4],
                    [90,
                        9,
                        1,
                        8],
                    ["getByNamespace",
                        11,
                        1,
                        6,
                        0,
                        4],
                    [91,
                        9,
                        1,
                        2,
                        0,
                        4]],
                0,
                9],
            [11,
                3,
                [[38,
                        2]],
                0,
                [[39,
                        0,
                        2,
                        0,
                        4],
                    [88,
                        0,
                        2,
                        0,
                        4]],
                [[41,
                        9,
                        1,
                        18,
                        0,
                        4],
                    [91,
                        9,
                        1,
                        2,
                        0,
                        4],
                    ["getOnlyItem",
                        9,
                        0,
                        2,
                        0,
                        4],
                    ["getOnlyItemOrNullObject",
                        9,
                        0,
                        2,
                        0,
                        4]],
                0,
                9],
            [12,
                6,
                [[92,
                        3],
                    [38,
                        2],
                    ["allowCloseOnUntitled",
                        1]],
                [[93,
                        27,
                        19,
                        0,
                        0,
                        4],
                    [94,
                        4,
                        35,
                        0,
                        0,
                        4],
                    [52,
                        6,
                        19,
                        0,
                        0,
                        4],
                    [95,
                        14,
                        35,
                        1,
                        0,
                        4],
                    [96,
                        29,
                        19,
                        3,
                        0,
                        4],
                    [97,
                        10,
                        19,
                        3,
                        0,
                        4]],
                [["_GetObjectByReferenceId",
                        1,
                        2,
                        0,
                        4],
                    ["_GetObjectTypeNameByReferenceId",
                        1,
                        2,
                        0,
                        4],
                    ["_RemoveReference",
                        1,
                        2,
                        0,
                        4],
                    ["_RemoveAllReferences",
                        0,
                        2,
                        0,
                        4],
                    [98],
                    [39,
                        0,
                        2,
                        0,
                        4],
                    [99,
                        1,
                        0,
                        3],
                    [100,
                        1,
                        2,
                        0,
                        4],
                    [101,
                        2],
                    ["setMetadataOnTile",
                        3],
                    ["launchTapPane",
                        1,
                        2,
                        0,
                        5],
                    ["getNeighborhoodTextAroundSelection",
                        1,
                        2,
                        0,
                        5],
                    ["_RegisterEvent",
                        2,
                        2,
                        0,
                        4],
                    ["_UnregisterEvent",
                        2,
                        2,
                        0,
                        4],
                    ["setNavigationPaneVisibility",
                        1,
                        0,
                        4,
                        1]],
                [["getSelection",
                        23,
                        0,
                        10,
                        0,
                        4],
                    [102,
                        23,
                        1,
                        2,
                        3,
                        4],
                    [103,
                        23,
                        1,
                        2,
                        3,
                        4],
                    ["getAnnotationsByType",
                        2,
                        1,
                        6,
                        4,
                        4]],
                0,
                0,
                [["ContentControlAdded",
                        2,
                        3,
                        "NumericEventType.contentControlAdded",
                        44,
                        "_RegisterContentControlAddedEvent",
                        "_UnregisterContentControlAddedEvent"]]],
            [13,
                2,
                [[92,
                        3,
                        5],
                    [38,
                        2]],
                [[93,
                        27,
                        19,
                        5,
                        0,
                        4],
                    [94,
                        4,
                        35,
                        5,
                        0,
                        4],
                    [52,
                        6,
                        19,
                        5,
                        0,
                        4],
                    [95,
                        14,
                        35,
                        5,
                        0,
                        4],
                    [96,
                        29,
                        19,
                        6,
                        0,
                        4],
                    [97,
                        10,
                        19,
                        6,
                        0,
                        4]],
                [[98,
                        0,
                        0,
                        5],
                    [39,
                        0,
                        2,
                        0,
                        4],
                    ["open",
                        0,
                        2,
                        0,
                        4],
                    [99,
                        1,
                        0,
                        6]],
                [[102,
                        23,
                        1,
                        2,
                        6,
                        4],
                    [103,
                        23,
                        1,
                        2,
                        6,
                        4]]],
            [14,
                2,
                [[38,
                        2],
                    [75,
                        1],
                    ["subject",
                        1],
                    ["author",
                        1],
                    ["keywords",
                        1],
                    ["comments",
                        1],
                    ["template",
                        3],
                    ["lastAuthor",
                        3],
                    ["revisionNumber",
                        3],
                    ["applicationName",
                        3],
                    ["lastPrintDate",
                        11],
                    ["creationDate",
                        11],
                    ["lastSaveTime",
                        11],
                    ["security",
                        3],
                    ["category",
                        1],
                    ["format",
                        1],
                    ["manager",
                        1],
                    ["company",
                        1]],
                [["customProperties",
                        8,
                        19,
                        0,
                        0,
                        4]],
                [[39,
                        0,
                        2,
                        0,
                        4]]],
            [15,
                2,
                [[38,
                        2],
                    ["name",
                        1],
                    ["size",
                        1],
                    ["bold",
                        1],
                    ["italic",
                        1],
                    [76,
                        1],
                    ["underline",
                        1],
                    ["subscript",
                        1],
                    ["superscript",
                        1],
                    ["strikeThrough",
                        1],
                    ["doubleStrikeThrough",
                        1],
                    ["highlightColor",
                        1]],
                0,
                [[39,
                        0,
                        2,
                        0,
                        4]]],
            [16,
                2,
                [[87,
                        2],
                    [38,
                        2],
                    ["altTextDescription",
                        1],
                    ["altTextTitle",
                        1],
                    ["height",
                        1],
                    [104,
                        1],
                    ["lockAspectRatio",
                        1],
                    [105,
                        1],
                    ["imageFormat",
                        3,
                        3]],
                [[53,
                        5,
                        2,
                        0,
                        0,
                        4],
                    ["paragraph",
                        21,
                        2,
                        2,
                        0,
                        4],
                    [77,
                        34,
                        2,
                        1,
                        0,
                        4],
                    [78,
                        30,
                        2,
                        1,
                        0,
                        4],
                    [59,
                        5,
                        2,
                        1,
                        0,
                        4],
                    [79,
                        34,
                        2,
                        1,
                        0,
                        4],
                    [80,
                        30,
                        2,
                        1,
                        0,
                        4]],
                [[39,
                        0,
                        2,
                        0,
                        4],
                    ["getBase64ImageSrc",
                        0,
                        2,
                        0,
                        4],
                    [60,
                        2,
                        0,
                        2],
                    [40,
                        0,
                        0,
                        2],
                    [64,
                        1,
                        2,
                        2,
                        4]],
                [[69,
                        5,
                        0,
                        8],
                    [71,
                        16,
                        2,
                        8,
                        2],
                    [65,
                        23,
                        2,
                        8,
                        2],
                    [66,
                        23,
                        2,
                        8,
                        2],
                    [67,
                        23,
                        2,
                        8,
                        2],
                    [68,
                        21,
                        2,
                        8,
                        2],
                    [70,
                        23,
                        2,
                        8,
                        2],
                    [73,
                        23,
                        1,
                        2,
                        1,
                        4],
                    [106,
                        16,
                        0,
                        2,
                        1,
                        4],
                    [107,
                        16,
                        0,
                        2,
                        1,
                        4]]],
            [17,
                3,
                [[38,
                        2]],
                0,
                [[39,
                        0,
                        2,
                        0,
                        4]],
                [[108,
                        16,
                        1,
                        18,
                        0,
                        4],
                    [42,
                        16,
                        0,
                        2,
                        1,
                        4],
                    [43,
                        16,
                        0,
                        2,
                        1,
                        4]],
                0,
                16],
            [18,
                2,
                [[37,
                        3],
                    [38,
                        2],
                    ["levelTypes",
                        3],
                    ["levelExistences",
                        3]],
                [[51,
                        22,
                        19,
                        0,
                        0,
                        4]],
                [[39,
                        0,
                        2,
                        0,
                        4],
                    ["setLevelBullet",
                        4],
                    ["setLevelNumbering",
                        3],
                    ["getLevelString",
                        1,
                        2,
                        0,
                        4],
                    ["setLevelPicture",
                        2,
                        0,
                        3],
                    ["getLevelPicture",
                        1,
                        2,
                        3,
                        4],
                    ["resetLevelFont",
                        2,
                        0,
                        3],
                    ["setLevelAlignment",
                        2],
                    ["setLevelIndents",
                        3],
                    ["setLevelStartingNumber",
                        2]],
                [[68,
                        21,
                        2,
                        8],
                    ["getLevelParagraphs",
                        22,
                        1,
                        6,
                        0,
                        4],
                    ["getLevelFont",
                        15,
                        1,
                        2,
                        3,
                        4]]],
            [19,
                3,
                [[38,
                        2]],
                0,
                [[39,
                        0,
                        2,
                        0,
                        4]],
                [[41,
                        18,
                        1,
                        18,
                        0,
                        4],
                    [83,
                        18,
                        1,
                        2,
                        0,
                        4],
                    [42,
                        18,
                        0,
                        2,
                        0,
                        4],
                    [84,
                        18,
                        1,
                        2,
                        0,
                        4],
                    [43,
                        18,
                        0,
                        2,
                        0,
                        4]],
                0,
                18],
            [20,
                2,
                [[38,
                        2],
                    ["siblingIndex",
                        3],
                    ["listString",
                        3],
                    ["level",
                        1]],
                0,
                [[39,
                        0,
                        2,
                        0,
                        4]],
                [["getAncestor",
                        21,
                        1,
                        2,
                        0,
                        4],
                    ["getDescendants",
                        22,
                        1,
                        6,
                        0,
                        4],
                    ["getAncestorOrNullObject",
                        21,
                        1,
                        2,
                        0,
                        4]]],
            [21,
                2,
                [[87,
                        2],
                    [38,
                        2],
                    [47,
                        1],
                    [109,
                        1],
                    ["firstLineIndent",
                        1],
                    ["leftIndent",
                        1],
                    ["rightIndent",
                        1],
                    ["lineSpacing",
                        1],
                    ["outlineLevel",
                        1],
                    ["spaceBefore",
                        1],
                    ["spaceAfter",
                        1],
                    ["lineUnitBefore",
                        1],
                    ["lineUnitAfter",
                        1],
                    [48,
                        3],
                    ["isListItem",
                        3,
                        1],
                    ["tableNestingLevel",
                        3,
                        1],
                    ["isLastParagraph",
                        3,
                        1],
                    [50,
                        1,
                        1]],
                [[54,
                        15,
                        35,
                        0,
                        0,
                        4],
                    [52,
                        6,
                        18,
                        0,
                        0,
                        4],
                    [53,
                        5,
                        2,
                        0,
                        0,
                        4],
                    [55,
                        17,
                        19,
                        0,
                        0,
                        4],
                    [56,
                        4,
                        2,
                        1,
                        0,
                        4],
                    ["list",
                        18,
                        2,
                        1,
                        0,
                        4],
                    [77,
                        34,
                        2,
                        1,
                        0,
                        4],
                    [78,
                        30,
                        2,
                        1,
                        0,
                        4],
                    ["listItem",
                        20,
                        35,
                        1,
                        0,
                        4],
                    [59,
                        5,
                        2,
                        1,
                        0,
                        4],
                    [79,
                        34,
                        2,
                        1,
                        0,
                        4],
                    [80,
                        30,
                        2,
                        1,
                        0,
                        4],
                    ["listOrNullObject",
                        18,
                        2,
                        1,
                        0,
                        4],
                    ["listItemOrNullObject",
                        20,
                        35,
                        1,
                        0,
                        4]],
                [[39,
                        0,
                        2,
                        0,
                        4],
                    [60,
                        2],
                    [61],
                    [40],
                    [64,
                        1,
                        2,
                        0,
                        4],
                    [62,
                        0,
                        2,
                        0,
                        4],
                    [63,
                        0,
                        2,
                        0,
                        4],
                    ["detachFromList",
                        0,
                        0,
                        1],
                    [100,
                        1,
                        2,
                        0,
                        4],
                    [101,
                        2]],
                [[71,
                        16,
                        2,
                        8],
                    [69,
                        5,
                        0,
                        8],
                    [65,
                        23,
                        2,
                        8],
                    [66,
                        23,
                        2,
                        8],
                    [67,
                        23,
                        2,
                        8],
                    [70,
                        23,
                        2,
                        8],
                    [68,
                        21,
                        2,
                        8],
                    [72,
                        24,
                        2,
                        7,
                        0,
                        4],
                    [73,
                        23,
                        1,
                        2,
                        1,
                        4],
                    [81,
                        24,
                        3,
                        6,
                        1,
                        4],
                    [74,
                        30,
                        4,
                        8,
                        1],
                    [82,
                        24,
                        2,
                        6,
                        1,
                        4],
                    ["startNewList",
                        18,
                        0,
                        0,
                        1],
                    ["attachToList",
                        18,
                        2,
                        0,
                        1],
                    [106,
                        21,
                        0,
                        2,
                        1,
                        4],
                    ["getPrevious",
                        21,
                        0,
                        2,
                        1,
                        4],
                    [107,
                        21,
                        0,
                        2,
                        1,
                        4],
                    ["getPreviousOrNullObject",
                        21,
                        0,
                        2,
                        1,
                        4],
                    ["getSubrange",
                        23,
                        2,
                        2,
                        1,
                        4]]],
            [22,
                3,
                [[38,
                        2]],
                0,
                [[39,
                        0,
                        2,
                        0,
                        4]],
                [[108,
                        21,
                        1,
                        18,
                        0,
                        4],
                    [42,
                        21,
                        0,
                        2,
                        1,
                        4],
                    ["getLast",
                        21,
                        0,
                        2,
                        1,
                        4],
                    [43,
                        21,
                        0,
                        2,
                        1,
                        4],
                    ["getLastOrNullObject",
                        21,
                        0,
                        2,
                        1,
                        4]],
                0,
                21],
            [23,
                2,
                [[87,
                        2],
                    [38,
                        2],
                    [47,
                        1],
                    [48,
                        3],
                    ["isEmpty",
                        3,
                        1],
                    [104,
                        1,
                        1],
                    [50,
                        1,
                        1]],
                [[54,
                        15,
                        35,
                        0,
                        0,
                        4],
                    [51,
                        22,
                        18,
                        0,
                        0,
                        4],
                    [52,
                        6,
                        18,
                        0,
                        0,
                        4],
                    [53,
                        5,
                        2,
                        0,
                        0,
                        4],
                    [55,
                        17,
                        19,
                        2,
                        0,
                        4],
                    [57,
                        19,
                        18,
                        1,
                        0,
                        4],
                    [58,
                        31,
                        18,
                        1,
                        0,
                        4],
                    [77,
                        34,
                        2,
                        1,
                        0,
                        4],
                    [78,
                        30,
                        2,
                        1,
                        0,
                        4],
                    [56,
                        4,
                        2,
                        1,
                        0,
                        4],
                    [59,
                        5,
                        2,
                        1,
                        0,
                        4],
                    [79,
                        34,
                        2,
                        1,
                        0,
                        4],
                    [80,
                        30,
                        2,
                        1,
                        0,
                        4]],
                [[39,
                        0,
                        2,
                        0,
                        4],
                    [60,
                        2],
                    [61],
                    [40],
                    [64,
                        1,
                        2,
                        0,
                        4],
                    [62,
                        0,
                        2,
                        0,
                        4],
                    [63,
                        0,
                        2,
                        0,
                        4],
                    ["compareLocationWith",
                        1,
                        2,
                        1,
                        4],
                    ["getBookmarks",
                        2,
                        2,
                        3,
                        4],
                    ["insertBookmark",
                        1,
                        0,
                        3],
                    ["highlight",
                        1,
                        2,
                        0,
                        4],
                    ["removeHighlight",
                        0,
                        2,
                        0,
                        4],
                    ["previewTextReplacement",
                        1,
                        2,
                        0,
                        4],
                    ["endPreview",
                        0,
                        2,
                        0,
                        4]],
                [[69,
                        5,
                        0,
                        8],
                    [65,
                        23,
                        2,
                        8],
                    [66,
                        23,
                        2,
                        8],
                    [67,
                        23,
                        2,
                        8],
                    [70,
                        23,
                        2,
                        8],
                    [68,
                        21,
                        2,
                        8],
                    [71,
                        16,
                        2,
                        8,
                        2],
                    [72,
                        24,
                        2,
                        7,
                        0,
                        4],
                    [73,
                        23,
                        1,
                        2,
                        1,
                        4],
                    [81,
                        24,
                        4,
                        6,
                        1,
                        4],
                    ["expandTo",
                        23,
                        1,
                        0,
                        1],
                    ["intersectWith",
                        23,
                        1,
                        0,
                        1],
                    ["getNextTextRange",
                        23,
                        2,
                        2,
                        1,
                        4],
                    ["getHyperlinkRanges",
                        24,
                        0,
                        6,
                        1,
                        4],
                    [74,
                        30,
                        4,
                        8,
                        1],
                    [82,
                        24,
                        2,
                        6,
                        1,
                        4],
                    ["expandToOrNullObject",
                        23,
                        1,
                        0,
                        1],
                    ["intersectWithOrNullObject",
                        23,
                        1,
                        0,
                        1],
                    ["getNextTextRangeOrNullObject",
                        23,
                        2,
                        2,
                        1,
                        4],
                    ["insertTapObjectFromFileContent",
                        23,
                        3,
                        8,
                        0,
                        1]]],
            [24,
                3,
                [[38,
                        2]],
                0,
                [[39,
                        0,
                        2,
                        0,
                        4]],
                [[108,
                        23,
                        1,
                        18,
                        0,
                        4],
                    [42,
                        23,
                        0,
                        2,
                        1,
                        4],
                    [43,
                        23,
                        0,
                        2,
                        1,
                        4]],
                0,
                23],
            [25,
                4,
                [["ignorePunct",
                        1],
                    ["ignoreSpace",
                        1],
                    ["matchCase",
                        1],
                    ["matchPrefix",
                        1],
                    ["matchSuffix",
                        1],
                    ["matchWildcards",
                        1],
                    ["matchWholeWord",
                        1]],
                0,
                0,
                0,
                0,
                0,
                0,
                "Microsoft.WordServices.SearchOptions",
                4],
            [26,
                2,
                [[87,
                        2],
                    [38,
                        2]],
                [[94,
                        4,
                        35,
                        0,
                        0,
                        4]],
                [[39,
                        0,
                        2,
                        0,
                        4]],
                [["getHeader",
                        4,
                        1,
                        10,
                        0,
                        4],
                    ["getFooter",
                        4,
                        1,
                        10,
                        0,
                        4],
                    [106,
                        26,
                        0,
                        2,
                        1,
                        4],
                    [107,
                        26,
                        0,
                        2,
                        1,
                        4]]],
            [27,
                3,
                [[38,
                        2]],
                0,
                [[39,
                        0,
                        2,
                        0,
                        4]],
                [[108,
                        26,
                        1,
                        18,
                        0,
                        4],
                    [42,
                        26,
                        0,
                        2,
                        1,
                        4],
                    [43,
                        26,
                        0,
                        2,
                        1,
                        4]],
                0,
                26],
            [28,
                14,
                [[38,
                        2],
                    [85,
                        3],
                    [86,
                        5],
                    [87,
                        2]],
                0,
                [[39,
                        0,
                        2,
                        0,
                        4],
                    [40]]],
            [29,
                3,
                [[38,
                        2]],
                0,
                [[39,
                        0,
                        2,
                        0,
                        4],
                    [88,
                        0,
                        2,
                        0,
                        4],
                    [89]],
                [[41,
                        28,
                        1,
                        18,
                        0,
                        4],
                    [90,
                        28,
                        2,
                        9],
                    [91,
                        28,
                        1,
                        2,
                        0,
                        4]],
                0,
                28],
            [30,
                2,
                [[87,
                        2],
                    [38,
                        2],
                    ["isUniform",
                        3],
                    ["nestingLevel",
                        3],
                    [110,
                        1],
                    [47,
                        1],
                    ["rowCount",
                        3],
                    ["headerRowCount",
                        1],
                    ["styleTotalRow",
                        1],
                    ["styleFirstColumn",
                        1],
                    ["styleLastColumn",
                        1],
                    ["styleBandedRows",
                        1],
                    ["styleBandedColumns",
                        1],
                    [111,
                        1],
                    [112,
                        1],
                    [113,
                        1],
                    [105,
                        1],
                    [50,
                        1],
                    [109,
                        1]],
                [["rows",
                        33,
                        19,
                        0,
                        0,
                        4],
                    [58,
                        31,
                        19,
                        0,
                        0,
                        4],
                    [77,
                        34,
                        2,
                        0,
                        0,
                        4],
                    [78,
                        30,
                        2,
                        0,
                        0,
                        4],
                    [54,
                        15,
                        35,
                        0,
                        0,
                        4],
                    [53,
                        5,
                        2,
                        0,
                        0,
                        4],
                    [56,
                        4,
                        2,
                        0,
                        0,
                        4],
                    [79,
                        34,
                        2,
                        0,
                        0,
                        4],
                    [80,
                        30,
                        2,
                        0,
                        0,
                        4],
                    [59,
                        5,
                        2,
                        0,
                        0,
                        4]],
                [[39,
                        0,
                        2,
                        0,
                        4],
                    ["addColumns",
                        3],
                    [40],
                    [61],
                    ["deleteRows",
                        2],
                    ["deleteColumns",
                        2],
                    ["autoFitWindow"],
                    ["distributeColumns"],
                    [64,
                        1,
                        2,
                        0,
                        4],
                    [114,
                        1,
                        2,
                        0,
                        4],
                    [115,
                        2]],
                [["addRows",
                        33,
                        3,
                        4],
                    ["getCell",
                        34,
                        2,
                        2,
                        0,
                        4],
                    ["mergeCells",
                        34,
                        4,
                        8,
                        3],
                    [116,
                        36,
                        1,
                        2,
                        0,
                        4],
                    [72,
                        24,
                        2,
                        7,
                        0,
                        4],
                    [73,
                        23,
                        1,
                        2,
                        0,
                        4],
                    [69,
                        5,
                        0,
                        8],
                    [74,
                        30,
                        4,
                        8],
                    [68,
                        21,
                        2,
                        8],
                    [106,
                        30,
                        0,
                        2,
                        0,
                        4],
                    ["getParagraphBefore",
                        21,
                        0,
                        2,
                        0,
                        4],
                    ["getParagraphAfter",
                        21,
                        0,
                        2,
                        0,
                        4],
                    ["getCellOrNullObject",
                        34,
                        2,
                        2,
                        0,
                        4],
                    [107,
                        30,
                        0,
                        2,
                        0,
                        4],
                    ["getParagraphBeforeOrNullObject",
                        21,
                        0,
                        2,
                        0,
                        4],
                    ["getParagraphAfterOrNullObject",
                        21,
                        0,
                        2,
                        0,
                        4]]],
            [31,
                3,
                [[38,
                        2]],
                0,
                [[39,
                        0,
                        2,
                        0,
                        4]],
                [[108,
                        30,
                        1,
                        18,
                        0,
                        4],
                    [42,
                        30,
                        0,
                        2,
                        0,
                        4],
                    [43,
                        30,
                        0,
                        2,
                        0,
                        4]],
                0,
                30],
            [32,
                2,
                [[87,
                        2],
                    [38,
                        2],
                    ["cellCount",
                        3],
                    [117,
                        3],
                    [110,
                        1],
                    [111,
                        1],
                    [112,
                        1],
                    [113,
                        1],
                    ["isHeader",
                        3],
                    ["preferredHeight",
                        1]],
                [["cells",
                        35,
                        19,
                        0,
                        0,
                        4],
                    [78,
                        30,
                        2,
                        0,
                        0,
                        4],
                    [54,
                        15,
                        35,
                        0,
                        0,
                        4]],
                [[39,
                        0,
                        2,
                        0,
                        4],
                    [40],
                    [61],
                    [64,
                        1,
                        2,
                        0,
                        4],
                    [114,
                        1,
                        2,
                        0,
                        4],
                    [115,
                        2]],
                [[118,
                        33,
                        3,
                        6,
                        0,
                        4],
                    ["merge",
                        34,
                        0,
                        0,
                        3],
                    [72,
                        24,
                        2,
                        7,
                        0,
                        4],
                    [116,
                        36,
                        1,
                        2,
                        0,
                        4],
                    [106,
                        32,
                        0,
                        2,
                        0,
                        4],
                    [107,
                        32,
                        0,
                        2,
                        0,
                        4],
                    [69,
                        5,
                        0,
                        0,
                        3]]],
            [33,
                3,
                [[38,
                        2]],
                0,
                [[39,
                        0,
                        2,
                        0,
                        4]],
                [[108,
                        32,
                        1,
                        18,
                        0,
                        4],
                    [42,
                        32,
                        0,
                        2,
                        0,
                        4],
                    [43,
                        32,
                        0,
                        2,
                        0,
                        4]],
                0,
                32],
            [34,
                2,
                [[87,
                        2],
                    [38,
                        2],
                    [117,
                        3],
                    ["cellIndex",
                        3],
                    [86,
                        1],
                    [111,
                        1],
                    [112,
                        1],
                    [113,
                        1],
                    ["columnWidth",
                        1],
                    [105,
                        3]],
                [[78,
                        30,
                        2,
                        0,
                        0,
                        4],
                    ["parentRow",
                        32,
                        2,
                        0,
                        0,
                        4],
                    [94,
                        4,
                        35,
                        0,
                        0,
                        4]],
                [[39,
                        0,
                        2,
                        0,
                        4],
                    ["insertColumns",
                        3],
                    [81,
                        2,
                        0,
                        3],
                    ["deleteRow"],
                    ["deleteColumn"],
                    [114,
                        1,
                        2,
                        0,
                        4],
                    [115,
                        2]],
                [[118,
                        33,
                        3,
                        4],
                    [116,
                        36,
                        1,
                        2,
                        0,
                        4],
                    [106,
                        34,
                        0,
                        2,
                        0,
                        4],
                    [107,
                        34,
                        0,
                        2,
                        0,
                        4]]],
            [35,
                3,
                [[38,
                        2]],
                0,
                [[39,
                        0,
                        2,
                        0,
                        4]],
                [[108,
                        34,
                        1,
                        18,
                        0,
                        4],
                    [42,
                        34,
                        0,
                        2,
                        0,
                        4],
                    [43,
                        34,
                        0,
                        2,
                        0,
                        4]],
                0,
                34],
            [36,
                2,
                [[38,
                        2],
                    [76,
                        1],
                    [49,
                        1],
                    [105,
                        1]],
                0,
                [[39,
                        0,
                        2,
                        0,
                        4]]]] };
    var _builder = new OfficeExtension.LibraryBuilder({ metadata: _libraryMetadataWdJscomApi, targetNamespaceObject: Word });
})(Word || (Word = {}));
var Word;
(function (Word) {
    var RequestContext = (function (_super) {
        __extends(RequestContext, _super);
        function RequestContext(url) {
            var _this = _super.call(this, url) || this;
            _this.m_document = OfficeExtension.BatchApiHelper.createRootServiceObject(Word.Document, _this);
            _this._rootObject = _this.m_document;
            return _this;
        }
        Object.defineProperty(RequestContext.prototype, "document", {
            get: function () {
                return this.m_document;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(RequestContext.prototype, "application", {
            get: function () {
                if (this.m_application == null) {
                    this.m_application = OfficeExtension.BatchApiHelper.createTopLevelServiceObject(Word.Application, this, "Microsoft.WordServices.Application", false, 0);
                }
                return this.m_application;
            },
            enumerable: true,
            configurable: true
        });
        return RequestContext;
    }(OfficeCore.RequestContext));
    Word.RequestContext = RequestContext;
    function run(arg1, arg2) {
        return OfficeExtension.ClientRequestContext._runBatch("Word.run", arguments, function () { return new Word.RequestContext(); });
    }
    Word.run = run;
})(Word || (Word = {}));
var OSF;
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
        window.addEventListener('DOMContentLoaded', function (event) {
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
        });
    }
})(OSF || (OSF = {}));
OSFPerformance.hostInitializationEnd = OSFPerformance.now();


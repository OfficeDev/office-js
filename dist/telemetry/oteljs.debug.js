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
                throw {
                    message: "Invalid integer " + JSON.stringify(value)
                };
            }
        }
        TelemetryEventValidator.validateInt = validateInt;
    })(TelemetryEventValidator_TelemetryEventValidator || (TelemetryEventValidator_TelemetryEventValidator = {}));
    function makeBooleanDataField(name, value) {
        return {
            name: name,
            dataType: DataFieldType.Boolean,
            value: value
        };
    }
    function makeInt64DataField(name, value) {
        TelemetryEventValidator_TelemetryEventValidator.validateInt(value);
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
            if (contract.url !== undefined) {
                dataFields.push(makeStringDataField(instanceName + ".Url", contract.url));
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
    var __awaiter = undefined && undefined.__awaiter || function(thisArg, _arguments, P, generator) {
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
                result.done ? resolve(result.value) : new P(function(resolve) {
                    resolve(result.value);
                }).then(fulfilled, rejected);
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
                throw new Error("Could not find tenant token");
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
    var oteljsVersion = "3.1.7";
    var SuppressNexus = -1;
    var SimpleTelemetryLogger_SimpleTelemetryLogger = function() {
        function SimpleTelemetryLogger(parent, persistentDataFields) {
            var _a, _b;
            this.onSendEvent = new Event();
            this.telemetryEnabled = true;
            this.queriedForTelemetryEnabled = false;
            this.persistentDataFields = [];
            if (parent) {
                this.onSendEvent = parent.onSendEvent;
                (_a = this.persistentDataFields).push.apply(_a, parent.persistentDataFields);
            } else {
                this.persistentDataFields.push(makeStringDataField("OTelJS.Version", oteljsVersion));
            }
            if (persistentDataFields) {
                (_b = this.persistentDataFields).push.apply(_b, persistentDataFields);
            }
        }
        SimpleTelemetryLogger.prototype.sendTelemetryEvent = function(event) {
            try {
                if (!this.isTelemetryEnabled()) {
                    return;
                }
                if (this.onSendEvent.getListenerCount() === 0) {
                    logNotification(LogLevel.Warning, Category.Core, function() {
                        return "No telemetry sinks are attached.";
                    });
                    return;
                }
                var localEvent = this.cloneEvent(event);
                this.processTelemetryEvent(localEvent);
                this.onSendEvent.fireEvent(localEvent);
            } catch (error) {
                var errorMessage_1;
                if (error instanceof Error) {
                    errorMessage_1 = error.message;
                } else {
                    errorMessage_1 = JSON.stringify(error);
                }
                logNotification(LogLevel.Error, Category.Core, function() {
                    return errorMessage_1;
                });
            }
        };
        SimpleTelemetryLogger.prototype.processTelemetryEvent = function(event) {
            var _a;
            if (!event.telemetryProperties) {
                event.telemetryProperties = TenantTokenManager_TenantTokenManager.getTenantTokens(event.eventName);
            }
            (_a = event.dataFields).push.apply(_a, this.persistentDataFields);
            TelemetryEventValidator_TelemetryEventValidator.validateTelemetryEvent(event);
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
        SimpleTelemetryLogger.prototype.setIsTelemetryEnabled_TestOnly = function(enabled) {
            this.telemetryEnabled = enabled;
            this.queriedForTelemetryEnabled = true;
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
        SimpleTelemetryLogger.prototype.isTelemetryEnabled = function() {
            if (!this.queriedForTelemetryEnabled) {
                this.telemetryEnabled = this.isTelemetryEnabledInternal();
                this.queriedForTelemetryEnabled = true;
            }
            return this.telemetryEnabled;
        };
        SimpleTelemetryLogger.prototype.isTelemetryEnabledInternal = function() {
            if (typeof OSF !== "undefined") {
                if (typeof OSF.AppTelemetry === "undefined" || typeof OSF.AppTelemetry.enableTelemetry === "undefined" || OSF.AppTelemetry.enableTelemetry === false) {
                    logNotification(LogLevel.Warning, Category.Core, function() {
                        return "AppTelemetry is disabled for this platform.";
                    });
                    return false;
                }
            }
            return true;
        };
        return SimpleTelemetryLogger;
    }();
    var __extends = undefined && undefined.__extends || function() {
        var extendStatics = Object.setPrototypeOf || {
            __proto__: []
        } instanceof Array && function(d, b) {
            d.__proto__ = b;
        } || function(d, b) {
            for (var p in b) if (b.hasOwnProperty(p)) d[p] = b[p];
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
                result.done ? resolve(result.value) : new P(function(resolve) {
                    resolve(result.value);
                }).then(fulfilled, rejected);
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
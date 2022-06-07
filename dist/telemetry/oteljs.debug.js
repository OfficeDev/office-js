var oteljs = function(modules) {
    var installedModules = {};
    function __webpack_require__(moduleId) {
        if (installedModules[moduleId]) return installedModules[moduleId].exports;
        var module = installedModules[moduleId] = {
            i: moduleId,
            l: !1,
            exports: {}
        };
        return modules[moduleId].call(module.exports, module, module.exports, __webpack_require__), 
        module.l = !0, module.exports;
    }
    return __webpack_require__.m = modules, __webpack_require__.c = installedModules, 
    __webpack_require__.d = function(exports, name, getter) {
        __webpack_require__.o(exports, name) || Object.defineProperty(exports, name, {
            enumerable: !0,
            get: getter
        });
    }, __webpack_require__.r = function(exports) {
        "undefined" != typeof Symbol && Symbol.toStringTag && Object.defineProperty(exports, Symbol.toStringTag, {
            value: "Module"
        }), Object.defineProperty(exports, "__esModule", {
            value: !0
        });
    }, __webpack_require__.t = function(value, mode) {
        if (1 & mode && (value = __webpack_require__(value)), 8 & mode) return value;
        if (4 & mode && "object" == typeof value && value && value.__esModule) return value;
        var ns = Object.create(null);
        if (__webpack_require__.r(ns), Object.defineProperty(ns, "default", {
            enumerable: !0,
            value: value
        }), 2 & mode && "string" != typeof value) for (var key in value) __webpack_require__.d(ns, key, function(key) {
            return value[key];
        }.bind(null, key));
        return ns;
    }, __webpack_require__.n = function(module) {
        var getter = module && module.__esModule ? function() {
            return module.default;
        } : function() {
            return module;
        };
        return __webpack_require__.d(getter, "a", getter), getter;
    }, __webpack_require__.o = function(object, property) {
        return Object.prototype.hasOwnProperty.call(object, property);
    }, __webpack_require__.p = "", __webpack_require__(__webpack_require__.s = 19);
}([ function(module, __webpack_exports__, __webpack_require__) {
    "use strict";
    __webpack_require__.d(__webpack_exports__, "a", (function() {
        return makeBooleanDataField;
    })), __webpack_require__.d(__webpack_exports__, "d", (function() {
        return makeInt64DataField;
    })), __webpack_require__.d(__webpack_exports__, "b", (function() {
        return makeDoubleDataField;
    })), __webpack_require__.d(__webpack_exports__, "e", (function() {
        return makeStringDataField;
    })), __webpack_require__.d(__webpack_exports__, "c", (function() {
        return makeGuidDataField;
    }));
    var _DataFieldType__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(3), _DataClassification__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(4);
    function makeBooleanDataField(name, value) {
        return {
            name: name,
            dataType: _DataFieldType__WEBPACK_IMPORTED_MODULE_0__.a.Boolean,
            value: value,
            classification: _DataClassification__WEBPACK_IMPORTED_MODULE_1__.a.SystemMetadata
        };
    }
    function makeInt64DataField(name, value) {
        return {
            name: name,
            dataType: _DataFieldType__WEBPACK_IMPORTED_MODULE_0__.a.Int64,
            value: value,
            classification: _DataClassification__WEBPACK_IMPORTED_MODULE_1__.a.SystemMetadata
        };
    }
    function makeDoubleDataField(name, value) {
        return {
            name: name,
            dataType: _DataFieldType__WEBPACK_IMPORTED_MODULE_0__.a.Double,
            value: value,
            classification: _DataClassification__WEBPACK_IMPORTED_MODULE_1__.a.SystemMetadata
        };
    }
    function makeStringDataField(name, value) {
        return {
            name: name,
            dataType: _DataFieldType__WEBPACK_IMPORTED_MODULE_0__.a.String,
            value: value,
            classification: _DataClassification__WEBPACK_IMPORTED_MODULE_1__.a.SystemMetadata
        };
    }
    function makeGuidDataField(name, value) {
        return {
            name: name,
            dataType: _DataFieldType__WEBPACK_IMPORTED_MODULE_0__.a.Guid,
            value: value,
            classification: _DataClassification__WEBPACK_IMPORTED_MODULE_1__.a.SystemMetadata
        };
    }
}, function(module, __webpack_exports__, __webpack_require__) {
    "use strict";
    __webpack_require__.d(__webpack_exports__, "b", (function() {
        return LogLevel;
    })), __webpack_require__.d(__webpack_exports__, "a", (function() {
        return Category;
    })), __webpack_require__.d(__webpack_exports__, "e", (function() {
        return onNotification;
    })), __webpack_require__.d(__webpack_exports__, "d", (function() {
        return logNotification;
    })), __webpack_require__.d(__webpack_exports__, "c", (function() {
        return logError;
    }));
    var LogLevel, Category, onNotificationEvent = new (__webpack_require__(10).a);
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
        logNotification(LogLevel.Error, category, (function() {
            var errorMessage = error instanceof Error ? error.message : "";
            return message + ": " + errorMessage;
        }));
    }
    !function(LogLevel) {
        LogLevel[LogLevel.Error = 0] = "Error", LogLevel[LogLevel.Warning = 1] = "Warning", 
        LogLevel[LogLevel.Info = 2] = "Info", LogLevel[LogLevel.Verbose = 3] = "Verbose";
    }(LogLevel || (LogLevel = {})), function(Category) {
        Category[Category.Core = 0] = "Core", Category[Category.Sink = 1] = "Sink", Category[Category.Transport = 2] = "Transport";
    }(Category || (Category = {}));
}, function(module, __webpack_exports__, __webpack_require__) {
    "use strict";
    __webpack_require__.d(__webpack_exports__, "a", (function() {
        return addContractField;
    }));
    var _DataFieldHelper__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(0);
    function addContractField(dataFields, instanceName, contractName) {
        dataFields.push(Object(_DataFieldHelper__WEBPACK_IMPORTED_MODULE_0__.e)("zC." + instanceName, contractName));
    }
}, function(module, __webpack_exports__, __webpack_require__) {
    "use strict";
    var DataFieldType;
    __webpack_require__.d(__webpack_exports__, "a", (function() {
        return DataFieldType;
    })), function(DataFieldType) {
        DataFieldType[DataFieldType.String = 0] = "String", DataFieldType[DataFieldType.Boolean = 1] = "Boolean", 
        DataFieldType[DataFieldType.Int64 = 2] = "Int64", DataFieldType[DataFieldType.Double = 3] = "Double", 
        DataFieldType[DataFieldType.Guid = 4] = "Guid";
    }(DataFieldType || (DataFieldType = {}));
}, function(module, __webpack_exports__, __webpack_require__) {
    "use strict";
    var DataClassification;
    __webpack_require__.d(__webpack_exports__, "a", (function() {
        return DataClassification;
    })), function(DataClassification) {
        DataClassification[DataClassification.EssentialServiceMetadata = 1] = "EssentialServiceMetadata", 
        DataClassification[DataClassification.AccountData = 2] = "AccountData", DataClassification[DataClassification.SystemMetadata = 4] = "SystemMetadata", 
        DataClassification[DataClassification.OrganizationIdentifiableInformation = 8] = "OrganizationIdentifiableInformation", 
        DataClassification[DataClassification.EndUserIdentifiableInformation = 16] = "EndUserIdentifiableInformation", 
        DataClassification[DataClassification.CustomerContent = 32] = "CustomerContent", 
        DataClassification[DataClassification.AccessControl = 64] = "AccessControl";
    }(DataClassification || (DataClassification = {}));
}, function(module, __webpack_exports__, __webpack_require__) {
    "use strict";
    var SamplingPolicy, PersistencePriority, CostPriority, DataCategories, DiagnosticLevel;
    __webpack_require__.d(__webpack_exports__, "e", (function() {
        return SamplingPolicy;
    })), __webpack_require__.d(__webpack_exports__, "d", (function() {
        return PersistencePriority;
    })), __webpack_require__.d(__webpack_exports__, "a", (function() {
        return CostPriority;
    })), __webpack_require__.d(__webpack_exports__, "b", (function() {
        return DataCategories;
    })), __webpack_require__.d(__webpack_exports__, "c", (function() {
        return DiagnosticLevel;
    })), function(SamplingPolicy) {
        SamplingPolicy[SamplingPolicy.NotSet = 0] = "NotSet", SamplingPolicy[SamplingPolicy.Measure = 1] = "Measure", 
        SamplingPolicy[SamplingPolicy.Diagnostics = 2] = "Diagnostics", SamplingPolicy[SamplingPolicy.CriticalBusinessImpact = 191] = "CriticalBusinessImpact", 
        SamplingPolicy[SamplingPolicy.CriticalCensus = 192] = "CriticalCensus", SamplingPolicy[SamplingPolicy.CriticalExperimentation = 193] = "CriticalExperimentation", 
        SamplingPolicy[SamplingPolicy.CriticalUsage = 194] = "CriticalUsage";
    }(SamplingPolicy || (SamplingPolicy = {})), function(PersistencePriority) {
        PersistencePriority[PersistencePriority.NotSet = 0] = "NotSet", PersistencePriority[PersistencePriority.Normal = 1] = "Normal", 
        PersistencePriority[PersistencePriority.High = 2] = "High";
    }(PersistencePriority || (PersistencePriority = {})), function(CostPriority) {
        CostPriority[CostPriority.NotSet = 0] = "NotSet", CostPriority[CostPriority.Normal = 1] = "Normal", 
        CostPriority[CostPriority.High = 2] = "High";
    }(CostPriority || (CostPriority = {})), function(DataCategories) {
        DataCategories[DataCategories.NotSet = 0] = "NotSet", DataCategories[DataCategories.SoftwareSetup = 1] = "SoftwareSetup", 
        DataCategories[DataCategories.ProductServiceUsage = 2] = "ProductServiceUsage", 
        DataCategories[DataCategories.ProductServicePerformance = 4] = "ProductServicePerformance", 
        DataCategories[DataCategories.DeviceConfiguration = 8] = "DeviceConfiguration", 
        DataCategories[DataCategories.InkingTypingSpeech = 16] = "InkingTypingSpeech";
    }(DataCategories || (DataCategories = {})), function(DiagnosticLevel) {
        DiagnosticLevel[DiagnosticLevel.ReservedDoNotUse = 0] = "ReservedDoNotUse", DiagnosticLevel[DiagnosticLevel.BasicEvent = 10] = "BasicEvent", 
        DiagnosticLevel[DiagnosticLevel.FullEvent = 100] = "FullEvent", DiagnosticLevel[DiagnosticLevel.NecessaryServiceDataEvent = 110] = "NecessaryServiceDataEvent", 
        DiagnosticLevel[DiagnosticLevel.AlwaysOnNecessaryServiceDataEvent = 120] = "AlwaysOnNecessaryServiceDataEvent";
    }(DiagnosticLevel || (DiagnosticLevel = {}));
}, function(module, __webpack_exports__, __webpack_require__) {
    "use strict";
    __webpack_require__.d(__webpack_exports__, "a", (function() {
        return Contracts;
    }));
    var officeeventschema_tml_Result, officeeventschema_tml_Activity, Activity, officeeventschema_tml_Host, officeeventschema_tml_User, officeeventschema_tml_SDX, officeeventschema_tml_Funnel, officeeventschema_tml_UserAction, Office_System_Error_Error, DataFieldHelper = __webpack_require__(0), Contract = __webpack_require__(2);
    (officeeventschema_tml_Result || (officeeventschema_tml_Result = {})).getFields = function(instanceName, contract) {
        var dataFields = [];
        return dataFields.push(Object(DataFieldHelper.d)(instanceName + ".Code", contract.code)), 
        void 0 !== contract.type && dataFields.push(Object(DataFieldHelper.e)(instanceName + ".Type", contract.type)), 
        void 0 !== contract.tag && dataFields.push(Object(DataFieldHelper.d)(instanceName + ".Tag", contract.tag)), 
        void 0 !== contract.isExpected && dataFields.push(Object(DataFieldHelper.a)(instanceName + ".IsExpected", contract.isExpected)), 
        Object(Contract.a)(dataFields, instanceName, "Office.System.Result"), dataFields;
    }, (Activity = officeeventschema_tml_Activity || (officeeventschema_tml_Activity = {})).contractName = "Office.System.Activity", 
    Activity.getFields = function(contract) {
        var dataFields = [];
        return void 0 !== contract.cV && dataFields.push(Object(DataFieldHelper.e)("Activity.CV", contract.cV)), 
        dataFields.push(Object(DataFieldHelper.d)("Activity.Duration", contract.duration)), 
        dataFields.push(Object(DataFieldHelper.d)("Activity.Count", contract.count)), dataFields.push(Object(DataFieldHelper.d)("Activity.AggMode", contract.aggMode)), 
        void 0 !== contract.success && dataFields.push(Object(DataFieldHelper.a)("Activity.Success", contract.success)), 
        void 0 !== contract.result && dataFields.push.apply(dataFields, officeeventschema_tml_Result.getFields("Activity.Result", contract.result)), 
        Object(Contract.a)(dataFields, "Activity", Activity.contractName), dataFields;
    }, (officeeventschema_tml_Host || (officeeventschema_tml_Host = {})).getFields = function(instanceName, contract) {
        var dataFields = [];
        return void 0 !== contract.id && dataFields.push(Object(DataFieldHelper.e)(instanceName + ".Id", contract.id)), 
        void 0 !== contract.version && dataFields.push(Object(DataFieldHelper.e)(instanceName + ".Version", contract.version)), 
        void 0 !== contract.sessionId && dataFields.push(Object(DataFieldHelper.e)(instanceName + ".SessionId", contract.sessionId)), 
        Object(Contract.a)(dataFields, instanceName, "Office.System.Host"), dataFields;
    }, (officeeventschema_tml_User || (officeeventschema_tml_User = {})).getFields = function(instanceName, contract) {
        var dataFields = [];
        return void 0 !== contract.alias && dataFields.push(Object(DataFieldHelper.e)(instanceName + ".Alias", contract.alias)), 
        void 0 !== contract.primaryIdentityHash && dataFields.push(Object(DataFieldHelper.e)(instanceName + ".PrimaryIdentityHash", contract.primaryIdentityHash)), 
        void 0 !== contract.primaryIdentitySpace && dataFields.push(Object(DataFieldHelper.e)(instanceName + ".PrimaryIdentitySpace", contract.primaryIdentitySpace)), 
        void 0 !== contract.tenantId && dataFields.push(Object(DataFieldHelper.e)(instanceName + ".TenantId", contract.tenantId)), 
        void 0 !== contract.tenantGroup && dataFields.push(Object(DataFieldHelper.e)(instanceName + ".TenantGroup", contract.tenantGroup)), 
        void 0 !== contract.isAnonymous && dataFields.push(Object(DataFieldHelper.a)(instanceName + ".IsAnonymous", contract.isAnonymous)), 
        Object(Contract.a)(dataFields, instanceName, "Office.System.User"), dataFields;
    }, (officeeventschema_tml_SDX || (officeeventschema_tml_SDX = {})).getFields = function(instanceName, contract) {
        var dataFields = [];
        return void 0 !== contract.id && dataFields.push(Object(DataFieldHelper.e)(instanceName + ".Id", contract.id)), 
        void 0 !== contract.version && dataFields.push(Object(DataFieldHelper.e)(instanceName + ".Version", contract.version)), 
        void 0 !== contract.instanceId && dataFields.push(Object(DataFieldHelper.e)(instanceName + ".InstanceId", contract.instanceId)), 
        void 0 !== contract.name && dataFields.push(Object(DataFieldHelper.e)(instanceName + ".Name", contract.name)), 
        void 0 !== contract.marketplaceType && dataFields.push(Object(DataFieldHelper.e)(instanceName + ".MarketplaceType", contract.marketplaceType)), 
        void 0 !== contract.sessionId && dataFields.push(Object(DataFiel
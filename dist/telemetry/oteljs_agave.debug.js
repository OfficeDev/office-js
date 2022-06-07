var oteljs_agave = function(modules) {
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
    }, __webpack_require__.p = "", __webpack_require__(__webpack_require__.s = 31);
}([ function(module, exports, __webpack_require__) {
    "use strict";
    Object.defineProperty(exports, "__esModule", {
        value: !0
    }), function(AWTPropertyType) {
        AWTPropertyType[AWTPropertyType.Unspecified = 0] = "Unspecified", AWTPropertyType[AWTPropertyType.String = 1] = "String", 
        AWTPropertyType[AWTPropertyType.Int64 = 2] = "Int64", AWTPropertyType[AWTPropertyType.Double = 3] = "Double", 
        AWTPropertyType[AWTPropertyType.Boolean = 4] = "Boolean", AWTPropertyType[AWTPropertyType.Date = 5] = "Date";
    }(exports.AWTPropertyType || (exports.AWTPropertyType = {})), function(AWTPiiKind) {
        AWTPiiKind[AWTPiiKind.NotSet = 0] = "NotSet", AWTPiiKind[AWTPiiKind.DistinguishedName = 1] = "DistinguishedName", 
        AWTPiiKind[AWTPiiKind.GenericData = 2] = "GenericData", AWTPiiKind[AWTPiiKind.IPV4Address = 3] = "IPV4Address", 
        AWTPiiKind[AWTPiiKind.IPv6Address = 4] = "IPv6Address", AWTPiiKind[AWTPiiKind.MailSubject = 5] = "MailSubject", 
        AWTPiiKind[AWTPiiKind.PhoneNumber = 6] = "PhoneNumber", AWTPiiKind[AWTPiiKind.QueryString = 7] = "QueryString", 
        AWTPiiKind[AWTPiiKind.SipAddress = 8] = "SipAddress", AWTPiiKind[AWTPiiKind.SmtpAddress = 9] = "SmtpAddress", 
        AWTPiiKind[AWTPiiKind.Identity = 10] = "Identity", AWTPiiKind[AWTPiiKind.Uri = 11] = "Uri", 
        AWTPiiKind[AWTPiiKind.Fqdn = 12] = "Fqdn", AWTPiiKind[AWTPiiKind.IPV4AddressLegacy = 13] = "IPV4AddressLegacy";
    }(exports.AWTPiiKind || (exports.AWTPiiKind = {})), function(AWTCustomerContentKind) {
        AWTCustomerContentKind[AWTCustomerContentKind.NotSet = 0] = "NotSet", AWTCustomerContentKind[AWTCustomerContentKind.GenericContent = 1] = "GenericContent";
    }(exports.AWTCustomerContentKind || (exports.AWTCustomerContentKind = {})), function(AWTEventPriority) {
        AWTEventPriority[AWTEventPriority.Low = 1] = "Low", AWTEventPriority[AWTEventPriority.Normal = 2] = "Normal", 
        AWTEventPriority[AWTEventPriority.High = 3] = "High", AWTEventPriority[AWTEventPriority.Immediate_sync = 5] = "Immediate_sync";
    }(exports.AWTEventPriority || (exports.AWTEventPriority = {})), function(AWTEventsDroppedReason) {
        AWTEventsDroppedReason[AWTEventsDroppedReason.NonRetryableStatus = 1] = "NonRetryableStatus", 
        AWTEventsDroppedReason[AWTEventsDroppedReason.QueueFull = 3] = "QueueFull", AWTEventsDroppedReason[AWTEventsDroppedReason.MaxRetryLimit = 4] = "MaxRetryLimit";
    }(exports.AWTEventsDroppedReason || (exports.AWTEventsDroppedReason = {})), function(AWTEventsRejectedReason) {
        AWTEventsRejectedReason[AWTEventsRejectedReason.InvalidEvent = 1] = "InvalidEvent", 
        AWTEventsRejectedReason[AWTEventsRejectedReason.SizeLimitExceeded = 2] = "SizeLimitExceeded", 
        AWTEventsRejectedReason[AWTEventsRejectedReason.KillSwitch = 3] = "KillSwitch";
    }(exports.AWTEventsRejectedReason || (exports.AWTEventsRejectedReason = {}));
}, function(module, exports, __webpack_require__) {
    "use strict";
    Object.defineProperty(exports, "__esModule", {
        value: !0
    });
    var Enums_1 = __webpack_require__(0);
    exports.AWTPropertyType = Enums_1.AWTPropertyType, exports.AWTPiiKind = Enums_1.AWTPiiKind, 
    exports.AWTEventPriority = Enums_1.AWTEventPriority, exports.AWTEventsDroppedReason = Enums_1.AWTEventsDroppedReason, 
    exports.AWTEventsRejectedReason = Enums_1.AWTEventsRejectedReason, exports.AWTCustomerContentKind = Enums_1.AWTCustomerContentKind;
    var Enums_2 = __webpack_require__(6);
    exports.AWTUserIdType = Enums_2.AWTUserIdType, exports.AWTSessionState = Enums_2.AWTSessionState;
    var DataModels_1 = __webpack_require__(12);
    exports.AWT_BEST_EFFORT = DataModels_1.AWT_BEST_EFFORT, exports.AWT_NEAR_REAL_TIME = DataModels_1.AWT_NEAR_REAL_TIME, 
    exports.AWT_REAL_TIME = DataModels_1.AWT_REAL_TIME;
    var AWTEventProperties_1 = __webpack_require__(7);
    exports.AWTEventProperties = AWTEventProperties_1.default;
    var AWTLogger_1 = __webpack_require__(13);
    exports.AWTLogger = AWTLogger_1.default;
    var AWTLogManager_1 = __webpack_require__(17);
    exports.AWTLogManager = AWTLogManager_1.default;
    var AWTTransmissionManager_1 = __webpack_require__(30);
    exports.AWTTransmissionManager = AWTTransmissionManager_1.default;
    var AWTSerializer_1 = __webpack_require__(15);
    exports.AWTSerializer = AWTSerializer_1.default;
    var AWTSemanticContext_1 = __webpack_require__(9);
    exports.AWTSemanticContext = AWTSemanticContext_1.default, exports.AWT_COLLECTOR_URL_UNITED_STATES = "https://us.pipe.aria.microsoft.com/Collector/3.0/", 
    exports.AWT_COLLECTOR_URL_GERMANY = "https://de.pipe.aria.microsoft.com/Collector/3.0/", 
    exports.AWT_COLLECTOR_URL_JAPAN = "https://jp.pipe.aria.microsoft.com/Collector/3.0/", 
    exports.AWT_COLLECTOR_URL_AUSTRALIA = "https://au.pipe.aria.microsoft.com/Collector/3.0/", 
    exports.AWT_COLLECTOR_URL_EUROPE = "https://eu.pipe.aria.microsoft.com/Collector/3.0/", 
    exports.AWT_COLLECTOR_URL_USGOV_DOD = "https://pf.pipe.aria.microsoft.com/Collector/3.0", 
    exports.AWT_COLLECTOR_URL_USGOV_DOJ = "https://tb.pipe.aria.microsoft.com/Collector/3.0";
}, , function(module, exports, __webpack_require__) {
    "use strict";
    Object.defineProperty(exports, "__esModule", {
        value: !0
    });
    var microsoft_bond_primitives_1 = __webpack_require__(8), Enums_1 = __webpack_require__(0), GuidRegex = /[xy]/g;
    exports.EventNameAndTypeRegex = /^[a-zA-Z]([a-zA-Z0-9]|_){2,98}[a-zA-Z0-9]$/, exports.EventNameDotRegex = /\./g, 
    exports.PropertyNameRegex = /^[a-zA-Z](([a-zA-Z0-9|_|\.]){0,98}[a-zA-Z0-9])?$/, 
    exports.StatsApiKey = "a387cfcf60114a43a7699f9fbb49289e-9bceb9fe-1c06-460f-96c5-6a0b247358bc-7238";
    var beaconsSupported = null, uInt8ArraySupported = null, useXDR = null;
    function isString(value) {
        return "string" == typeof value;
    }
    function isNumber(value) {
        return "number" == typeof value;
    }
    function isBoolean(value) {
        return "boolean" == typeof value;
    }
    function isDate(value) {
        return value instanceof Date;
    }
    function msToTicks(timeInMs) {
        return 1e4 * (timeInMs + 621355968e5);
    }
    function isReactNative() {
        return !("undefined" == typeof navigator || !navigator.product) && "ReactNative" === navigator.product;
    }
    function isServiceWorkerGlobalScope() {
        return "object" == typeof self && "ServiceWorkerGlobalScope" === self.constructor.name;
    }
    function twoDigit(n) {
        return n < 10 ? "0" + n : n.toString();
    }
    function isNotDefined(value) {
        return null == value || "" === value;
    }
    exports.numberToBondInt64 = function(value) {
        var bond_value = new microsoft_bond_primitives_1.Int64("0");
        return bond_value.low = 4294967295 & value, bond_value.high = Math.floor(value / 4294967296), 
        bond_value;
    }, exports.newGuid = function() {
        return "xxxxxxxx-xxxx-4xxx-yxxx-xxxxxxxxxxxx".replace(GuidRegex, (function(c) {
            var r = 16 * Math.random() | 0;
            return ("x" === c ? r : 3 & r | 8).toString(16);
        }));
    }, exports.isString = isString, exports.isNumber = isNumber, exports.isBoolean = isBoolean, 
    exports.isDate = isDate, exports.msToTicks = msToTicks, exports.getTenantId = function(apiKey) {
        var indexTenantId = apiKey.indexOf("-");
        return indexTenantId > -1 ? apiKey.substring(0, indexTenantId) : "";
    }, exports.isBeaconsSupported = function() {
        return null === beaconsSupported && (beaconsSupported = "undefined" != typeof navigator && Boolean(navigator.sendBeacon)), 
        beaconsSupported;
    }, exports.isUint8ArrayAvailable = function() {
        return null === uInt8ArraySupported && (uInt8ArraySupported = "undefined" != typeof Uint8Array && !function() {
            if ("undefined" != typeof navigator && navigator.userAgent) {
                var ua = navigator.userAgent.toLowerCase();
                if ((ua.indexOf("safari") >= 0 || ua.indexOf("firefox") >= 0) && ua.indexOf("chrome") < 0) return !0;
            }
            return !1;
        }() && !isReactNative()), uInt8ArraySupported;
    }, exports.isPriority = function(value) {
        return !(!isNumber(value) || !(value >= 1 && value <= 3 || 5 === value));
    }, exports.sanitizeProperty = function(name, property) {
        return !exports.PropertyNameRegex.test(name) || isNotDefined(property) ? null : (isNotDefined(property.value) && (property = {
            value: property,
            type: Enums_1.AWTPropertyType.Unspecified
        }), property.type = function(value, type) {
            switch (type = function(value) {
                if (isNumber(value) && value >= 0 && value <= 4) return !0;
                return !1;
            }(type) ? type : Enums_1.AWTPropertyType.Unspecified) {
              case Enums_1.AWTPropertyType.Unspecified:
                return function(value) {
                    switch (typeof value) {
                      case "string":
                        return Enums_1.AWTPropertyType.String;

                      case "boolean":
                        return Enums_1.AWTPropertyType.Boolean;

                      case "number":
                        return Enums_1.AWTPropertyType.Double;

                      case "object":
                        return isDate(value) ? Enums_1.AWTPropertyType.Date : null;
                    }
                    return null;
                }(value);

              case Enums_1.AWTPropertyType.String:
                return isString(value) ? type : null;

              case Enums_1.AWTPropertyType.Boolean:
                return isBoolean(value) ? type : null;

              case Enums_1.AWTPropertyType.Date:
                return isDate(value) && NaN !== value.getTime() ? type : null;

              case Enums_1.AWTPropertyType.Int64:
                return isNumber(value) && value % 1 == 0 ? type : null;

              case Enums_1.AWTPropertyType.Double:
                return isNumber(value) ? type : null;
            }
            return null;
        }(property.value, property.type), property.type ? (isDate(property.value) && (property.value = msToTicks(property.value.getTime())), 
        property.pii > 0 && property.cc > 0 ? null : property.pii ? function(value) {
            if (isNumber(value) && value >= 0 && value <= 13) return !0;
            return !1;
        }(property.pii) ? property : null : property.cc ? function(value) {
            if (isNumber(value) && value >= 0 && value <= 1) return !0;
            return !1;
        }(property.cc) ? property : null : property) : null);
    }, exports.getISOString = function(date) {
        return date.getUTCFullYear() + "-" + twoDigit(date.getUTCMonth() + 1) + "-" + twoDigit(date.getUTCDate()) + "T" + twoDigit(date.getUTCHours()) + ":" + twoDigit(date.getUTCMinutes()) + ":" + twoDigit(date.getUTCSeconds()) + "." + function(n) {
            if (n < 10) return "00" + n;
            if (n < 100) return "0" + n;
            return n.toString();
        }(date.getUTCMilliseconds()) + "Z";
    }, exports.useXDomainRequest = function() {
        if (null === useXDR) {
            var conn = new XMLHttpRequest;
            useXDR = "undefined" == typeof conn.withCredentials && "undefined" != typeof XDomainRequest;
        }
        return useXDR;
    }, exports.useFetchRequest = function() {
        return isReactNative() || isServiceWorkerGlobalScope();
    }, exports.isReactNative = isReactNative, exports.isServiceWorkerGlobalScope = isServiceWorkerGlobalScope;
}, function(module, exports, __webpack_require__) {
    "use strict";
    Object.defineProperty(exports, "__esModule", {
        value: !0
    });
    var DataModels_1 = __webpack_require__(12), Enums_1 = __webpack_require__(0), AWTQueueManager_1 = __webpack_require__(11), AWTStatsManager_1 = __webpack_require__(14), AWTEventProperties_1 = __webpack_require__(7), AWTLogManager_1 = __webpack_require__(17), Utils = __webpack_require__(3), AWTTransmissionManagerCore = function() {
        function AWTTransmissionManagerCore() {}
        return AWTTransmissionManagerCore.setEventsHandler = function(eventsHandler) {
            this._eventHandler = eventsHandler;
        }, AWTTransmissionManagerCore.getEventsHandler = function() {
            return this._eventHandler;
        }, AWTTransmissionManagerCore.scheduleTimer = function() {
            var _this = this, timer = this._profiles[this._currentProfile][2];
            this._timeout < 0 && timer >= 0 && !this._paused && (this._eventHandler.hasEvents() ? (0 === timer && this._currentBackoffCount > 0 && (timer = 1), 
            this._timeout = setTimeout((function() {
                return _this._batchAndSendEvents();
            }), timer * (1 << this._currentBackoffCount) * 1e3)) : this._timerCount = 0);
        }, AWTTransmissionManagerCore.initialize = function(config) {
            var _this = this;
            this._newEventsAllowed = !0, this._config = config, this._eventHandler = new AWTQueueManager_1.default(config.collectorUri, config.cacheMemorySizeLimitInNumberOfEvents, config.httpXHROverride, config.clockSkewRefreshDurationInMins), 
            this._initializeProfiles(), AWTStatsManager_1.default.initialize((function(stats, tenantId) {
                if (_this._config.canSendStatEvent("awt_stats")) {
                    var event_1 = new AWTEventProperties_1.default("awt_stats");
                    for (var statKey in event_1.setEventPriority(Enums_1.AWTEventPriority.Hi
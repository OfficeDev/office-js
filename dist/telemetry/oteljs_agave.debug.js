var oteljs_agave = function(modules) {
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
    return __webpack_require__(__webpack_require__.s = 21);
}([ function(module, exports, __webpack_require__) {
    "use strict";
    Object.defineProperty(exports, "__esModule", {
        value: true
    });
    var AWTPropertyType;
    (function(AWTPropertyType) {
        AWTPropertyType[AWTPropertyType["Unspecified"] = 0] = "Unspecified";
        AWTPropertyType[AWTPropertyType["String"] = 1] = "String";
        AWTPropertyType[AWTPropertyType["Int64"] = 2] = "Int64";
        AWTPropertyType[AWTPropertyType["Double"] = 3] = "Double";
        AWTPropertyType[AWTPropertyType["Boolean"] = 4] = "Boolean";
        AWTPropertyType[AWTPropertyType["Date"] = 5] = "Date";
    })(AWTPropertyType = exports.AWTPropertyType || (exports.AWTPropertyType = {}));
    var AWTPiiKind;
    (function(AWTPiiKind) {
        AWTPiiKind[AWTPiiKind["NotSet"] = 0] = "NotSet";
        AWTPiiKind[AWTPiiKind["DistinguishedName"] = 1] = "DistinguishedName";
        AWTPiiKind[AWTPiiKind["GenericData"] = 2] = "GenericData";
        AWTPiiKind[AWTPiiKind["IPV4Address"] = 3] = "IPV4Address";
        AWTPiiKind[AWTPiiKind["IPv6Address"] = 4] = "IPv6Address";
        AWTPiiKind[AWTPiiKind["MailSubject"] = 5] = "MailSubject";
        AWTPiiKind[AWTPiiKind["PhoneNumber"] = 6] = "PhoneNumber";
        AWTPiiKind[AWTPiiKind["QueryString"] = 7] = "QueryString";
        AWTPiiKind[AWTPiiKind["SipAddress"] = 8] = "SipAddress";
        AWTPiiKind[AWTPiiKind["SmtpAddress"] = 9] = "SmtpAddress";
        AWTPiiKind[AWTPiiKind["Identity"] = 10] = "Identity";
        AWTPiiKind[AWTPiiKind["Uri"] = 11] = "Uri";
        AWTPiiKind[AWTPiiKind["Fqdn"] = 12] = "Fqdn";
        AWTPiiKind[AWTPiiKind["IPV4AddressLegacy"] = 13] = "IPV4AddressLegacy";
    })(AWTPiiKind = exports.AWTPiiKind || (exports.AWTPiiKind = {}));
    var AWTCustomerContentKind;
    (function(AWTCustomerContentKind) {
        AWTCustomerContentKind[AWTCustomerContentKind["NotSet"] = 0] = "NotSet";
        AWTCustomerContentKind[AWTCustomerContentKind["GenericContent"] = 1] = "GenericContent";
    })(AWTCustomerContentKind = exports.AWTCustomerContentKind || (exports.AWTCustomerContentKind = {}));
    var AWTEventPriority;
    (function(AWTEventPriority) {
        AWTEventPriority[AWTEventPriority["Low"] = 1] = "Low";
        AWTEventPriority[AWTEventPriority["Normal"] = 2] = "Normal";
        AWTEventPriority[AWTEventPriority["High"] = 3] = "High";
        AWTEventPriority[AWTEventPriority["Immediate_sync"] = 5] = "Immediate_sync";
    })(AWTEventPriority = exports.AWTEventPriority || (exports.AWTEventPriority = {}));
    var AWTEventsDroppedReason;
    (function(AWTEventsDroppedReason) {
        AWTEventsDroppedReason[AWTEventsDroppedReason["NonRetryableStatus"] = 1] = "NonRetryableStatus";
        AWTEventsDroppedReason[AWTEventsDroppedReason["QueueFull"] = 3] = "QueueFull";
    })(AWTEventsDroppedReason = exports.AWTEventsDroppedReason || (exports.AWTEventsDroppedReason = {}));
    var AWTEventsRejectedReason;
    (function(AWTEventsRejectedReason) {
        AWTEventsRejectedReason[AWTEventsRejectedReason["InvalidEvent"] = 1] = "InvalidEvent";
        AWTEventsRejectedReason[AWTEventsRejectedReason["SizeLimitExceeded"] = 2] = "SizeLimitExceeded";
        AWTEventsRejectedReason[AWTEventsRejectedReason["KillSwitch"] = 3] = "KillSwitch";
    })(AWTEventsRejectedReason = exports.AWTEventsRejectedReason || (exports.AWTEventsRejectedReason = {}));
}, , function(module, exports, __webpack_require__) {
    "use strict";
    Object.defineProperty(exports, "__esModule", {
        value: true
    });
    var microsoft_bond_primitives_1 = __webpack_require__(8);
    var Enums_1 = __webpack_require__(0);
    var GuidRegex = /[xy]/g;
    var MSTillUnixEpoch = 621355968e5;
    var MSToTicksMultiplier = 1e4;
    var NullValue = null;
    exports.EventNameAndTypeRegex = /^[a-zA-Z]([a-zA-Z0-9]|_){2,98}[a-zA-Z0-9]$/;
    exports.EventNameDotRegex = /\./g;
    exports.PropertyNameRegex = /^[a-zA-Z](([a-zA-Z0-9|_|\.]){0,98}[a-zA-Z0-9])?$/;
    exports.StatsApiKey = "a387cfcf60114a43a7699f9fbb49289e-9bceb9fe-1c06-460f-96c5-6a0b247358bc-7238";
    var beaconsSupported = NullValue;
    var uInt8ArraySupported = NullValue;
    var useXDR = NullValue;
    function numberToBondInt64(value) {
        var bond_value = new microsoft_bond_primitives_1.Int64("0");
        bond_value.low = value & 4294967295;
        bond_value.high = Math.floor(value / 4294967296);
        return bond_value;
    }
    exports.numberToBondInt64 = numberToBondInt64;
    function newGuid() {
        return "xxxxxxxx-xxxx-4xxx-yxxx-xxxxxxxxxxxx".replace(GuidRegex, function(c) {
            var r = Math.random() * 16 | 0, v = c === "x" ? r : r & 3 | 8;
            return v.toString(16);
        });
    }
    exports.newGuid = newGuid;
    function isString(value) {
        return typeof value === "string";
    }
    exports.isString = isString;
    function isNumber(value) {
        return typeof value === "number";
    }
    exports.isNumber = isNumber;
    function isBoolean(value) {
        return typeof value === "boolean";
    }
    exports.isBoolean = isBoolean;
    function isDate(value) {
        return value instanceof Date;
    }
    exports.isDate = isDate;
    function msToTicks(timeInMs) {
        return (timeInMs + MSTillUnixEpoch) * MSToTicksMultiplier;
    }
    exports.msToTicks = msToTicks;
    function getTenantId(apiKey) {
        var indexTenantId = apiKey.indexOf("-");
        if (indexTenantId > -1) {
            return apiKey.substring(0, indexTenantId);
        }
        return "";
    }
    exports.getTenantId = getTenantId;
    function isBeaconsSupported() {
        if (beaconsSupported === NullValue) {
            beaconsSupported = typeof navigator !== "undefined" && Boolean(navigator.sendBeacon);
        }
        return beaconsSupported;
    }
    exports.isBeaconsSupported = isBeaconsSupported;
    function isUint8ArrayAvailable() {
        if (uInt8ArraySupported === NullValue) {
            uInt8ArraySupported = typeof Uint8Array !== "undefined" && !isSafariOrFirefox() && !isReactNative();
        }
        return uInt8ArraySupported;
    }
    exports.isUint8ArrayAvailable = isUint8ArrayAvailable;
    function isPriority(value) {
        if (isNumber(value) && (value >= 1 && value <= 3 || value === 5)) {
            return true;
        }
        return false;
    }
    exports.isPriority = isPriority;
    function sanitizeProperty(name, property) {
        if (!exports.PropertyNameRegex.test(name) || isNotDefined(property)) {
            return NullValue;
        }
        if (isNotDefined(property.value)) {
            property = {
                value: property,
                type: Enums_1.AWTPropertyType.Unspecified
            };
        }
        property.type = sanitizePropertyType(property.value, property.type);
        if (!property.type) {
            return NullValue;
        }
        if (isDate(property.value)) {
            property.value = msToTicks(property.value.getTime());
        }
        if (property.pii > 0 && property.cc > 0) {
            return NullValue;
        }
        if (property.pii) {
            return isPii(property.pii) ? property : NullValue;
        }
        if (property.cc) {
            return isCustomerContent(property.cc) ? property : NullValue;
        }
        return property;
    }
    exports.sanitizeProperty = sanitizeProperty;
    function getISOString(date) {
        return date.getUTCFullYear() + "-" + twoDigit(date.getUTCMonth() + 1) + "-" + twoDigit(date.getUTCDate()) + "T" + twoDigit(date.getUTCHours()) + ":" + twoDigit(date.getUTCMinutes()) + ":" + twoDigit(date.getUTCSeconds()) + "." + threeDigit(date.getUTCMilliseconds()) + "Z";
    }
    exports.getISOString = getISOString;
    function useXDomainRequest() {
        if (useXDR === NullValue) {
            var conn = new XMLHttpRequest();
            if (typeof conn.withCredentials === "undefined" && typeof XDomainRequest !== "undefined") {
                useXDR = true;
            } else {
                useXDR = false;
            }
        }
        return useXDR;
    }
    exports.useXDomainRequest = useXDomainRequest;
    function isReactNative() {
        if (typeof navigator !== "undefined" && navigator.product) {
            return navigator.product === "ReactNative";
        }
        return false;
    }
    exports.isReactNative = isReactNative;
    function twoDigit(n) {
        return n < 10 ? "0" + n : n.toString();
    }
    function threeDigit(n) {
        if (n < 10) {
            return "00" + n;
        } else if (n < 100) {
            return "0" + n;
        }
        return n.toString();
    }
    function sanitizePropertyType(value, type) {
        type = !isPropertyType(type) ? Enums_1.AWTPropertyType.Unspecified : type;
        switch (type) {
          case Enums_1.AWTPropertyType.Unspecified:
            return getCorrectType(value);

          case Enums_1.AWTPropertyType.String:
            return isString(value) ? type : NullValue;

          case Enums_1.AWTPropertyType.Boolean:
            return isBoolean(value) ? type : NullValue;

          case Enums_1.AWTPropertyType.Date:
            return isDate(value) && value.getTime() !== NaN ? type : NullValue;

          case Enums_1.AWTPropertyType.Int64:
            return isNumber(value) && value % 1 === 0 ? type : NullValue;

          case Enums_1.AWTPropertyType.Double:
            return isNumber(value) ? type : NullValue;
        }
        return NullValue;
    }
    function getCorrectType(value) {
        switch (typeof value) {
          case "string":
            return Enums_1.AWTPropertyType.String;

          case "boolean":
            return Enums_1.AWTPropertyType.Boolean;

          case "number":
            return Enums_1.AWTPropertyType.Double;

          case "object":
            return isDate(value) ? Enums_1.AWTPropertyType.Date : NullValue;
        }
        return NullValue;
    }
    function isPii(value) {
        if (isNumber(value) && value >= 0 && value <= 13) {
            return true;
        }
        return false;
    }
    function isCustomerContent(value) {
        if (isNumber(value) && value >= 0 && value <= 1) {
            return true;
        }
        return false;
    }
    function isPropertyType(value) {
        if (isNumber(value) && value >= 0 && value <= 4) {
            return true;
        }
        return false;
    }
    function isSafariOrFirefox() {
        if (typeof navigator !== "undefined" && navigator.userAgent) {
            var ua = navigator.userAgent.toLowerCase();
            if ((ua.indexOf("safari") >= 0 || ua.indexOf("firefox") >= 0) && ua.indexOf("chrome") < 0) {
                return true;
            }
        }
        return false;
    }
    function isNotDefined(value) {
        return value === undefined || value === NullValue || value === "";
    }
}, function(module, exports, __webpack_require__) {
    "use strict";
    Object.defineProperty(exports, "__esModule", {
        value: true
    });
    var DataModels_1 = __webpack_require__(12);
    var Enums_1 = __webpack_require__(0);
    var AWTQueueManager_1 = __webpack_require__(11);
    var AWTStatsManager_1 = __webpack_require__(14);
    var AWTEventProperties_1 = __webpack_require__(7);
    var AWTLogManager_1 = __webpack_require__(17);
    var Utils = __webpack_require__(2);
    var MaxBackoffCount = 4;
    var MinDurationBetweenUploadNow = 3e4;
    var StatName = "awt_stats";
    var AWTTransmissionManagerCore = function() {
        function AWTTransmissionManagerCore() {}
        AWTTransmissionManagerCore.setEventsHandler = function(eventsHandler) {
            this._eventHandler = eventsHandler;
        };
        AWTTransmissionManagerCore.getEventsHandler = function() {
            return this._eventHandler;
        };
        AWTTransmissionManagerCore.scheduleTimer = function() {
            var _this = this;
            var timer = this._profiles[this._currentProfile][2];
            if (this._timeout < 0 && timer >= 0 && !this._paused) {
                if (this._eventHandler.hasEvents()) {
                    if (timer === 0 && this._currentBackoffCount > 0) {
                        timer = 1;
                    }
                    this._timeout = setTimeout(function() {
                        return _this._batchAndSendEvents();
                    }, timer * (1 << this._currentBackoffCount) * 1e3);
                } else {
                    this._timerCount = 0;
                }
            }
        };
        AWTTransmissionManagerCore.initialize = function(config) {
            var _this = this;
            this._newEventsAllowed = true;
            this._config = config;
            this._eventHandler = new AWTQueueManager_1.default(config.collectorUri, config.cacheMemorySizeLimitInNumberOfEvents, config.httpXHROverride, config.clockSkewRefreshDurationInMins);
            this._initializeProfiles();
            AWTStatsManager_1.default.initialize(function(stats, tenantId) {
                if (_this._config.canSendStatEvent(StatName)) {
                    var event_1 = new AWTEventProperties_1.default(StatName);
                    event_1.setEventPriority(Enums_1.AWTEventPriority.High);
                    event_1.setProperty("TenantId", tenantId);
                    for (var statKey in stats) {
                        if (stats.hasOwnProperty(statKey)) {
                            event_1.setProperty(statKey, stats[statKey].toString());
                        }
                    }
                    AWTLogManager_1.default.getLogger(Utils.StatsApiKey).logEvent(event_1);
                }
            });
        };
        AWTTransmissionManagerCore.setTransmitProfile = function(profileName) {
            if (this._currentProfile !== profileName && this._profiles[profileName] !== undefined) {
                this.clearTimeout();
                this._currentProfile = profileName;
                this.scheduleTimer();
            }
        };
        AWTTransmissionManagerCore.loadTransmitProfiles = function(profiles) {
            this._resetTransmitProfiles();
            for (var profileName in profiles) {
                if (profiles.hasOwnProperty(profileName)) {
                    if (profiles[profileName].length !== 3) {
                        continue;
                    }
                    for (var i = 2; i >= 0; --i) {
                        if (profiles[profileName][i] < 0) {
                            for (var j = i; j >= 0; --j) {
                                profiles[profileName][j] = -1;
                            }
                            break;
                        }
                    }
                    for (var i = 2; i > 0; --i) {
                        if (profiles[profileName][i] > 0 && profiles[profileName][i - 1] > 0) {
                            var timerMultiplier = profiles[profileName][i - 1] / profiles[profileName][i];
                            profiles[profileName][i - 1] = Math.ceil(timerMultiplier) * profiles[profileName][i];
                        }
                    }
                    this._profiles[profileName] = profiles[profileName];
                }
            }
        };
        AWTTransmissionManagerCore.sendEvent = function(event) {
            if (this._newEventsAllowed) {
                if (this._currentBackoffCount > 0 && event.priority === Enums_1.AWTEventPriority.Immediate_sync) {
                    event.priority = Enums_1.AWTEventPriority.High;
                }
                this._eventHandler.addEvent(event);
                this.scheduleTimer();
            }
        };
        AWTTransmissionManagerCore.flush = function(callback) {
            var currentTime = new Date().getTime();
            if (!this._paused && this._lastUploadNowCall + MinDurationBetweenUploadNow < currentTime) {
                this._lastUploadNowCall = currentTime;
                if (this._timeout > -1) {
                    clearTimeout(this._timeout);
                    this._timeout = -1;
                }
                this._eventHandler.uploadNow(callback);
            }
        };
        AWTTransmissionManagerCore.pauseTransmission = function() {
            if (!this._paused) {
                this.clearTimeout();
                this._eventHandler.pauseTransmission();
                this._paused = true;
            }
        };
        AWTTransmissionManagerCore.resumeTransmision = function() {
            if (this._paused) {
                this._paused = false;
                this._eventHandler.resumeTransmission();
                this.scheduleTimer();
            }
        };
        AWTTransmissionManagerCore.flushAndTeardown = function() {
            AWTStatsManager_1.default.teardown();
            this._newEventsAllowed = false;
            this.clearTimeout();
            this._eventHandler.teardown();
        };
        AWTTransmissionManagerCore.backOffTransmission = function() {
            if (this._currentBackoffCount < MaxBackoffCount) {
                this._currentBackoffCount++;
                this.clearTimeout();
                this.scheduleTimer();
            }
        };
        AWTTransmissionManagerCore.clearBackOff = function() {
            if (this._currentBackoffCount > 0) {
                this._currentBackoffCount = 0;
                this.clearTimeout();
                this.scheduleTimer();
            }
        };
        AWTTransmissionManagerCore._resetTransmitProfiles = function() {
            this.clearTimeout();
            this._initializeProfiles();
            this._currentProfile = DataModels_1.AWT_REAL_TIME;
            this.scheduleTimer();
        };
        AWTTransmissionManagerCore.clearTimeout = function() {
            if (this._timeout > 0) {
                clearTimeout(this._timeout);
                this._timeout = -1;
                this._timerCount = 0;
            }
        };
        AWTTransmissionManagerCore._batchAndSendEvents = function() {
            var priority = Enums_1.AWTEventPriority.High;
            this._timerCount++;
            if (this._timerCount * this._profiles[this._currentProfile][2] === this._profiles[this._currentProfile][0]) {
                priority = Enums_1.AWTEventPriority.Low;
                this._timerCount = 0;
            } else if (this._timerCount * this._profiles[this._currentProfile][2] === this._profiles[this._currentProfile][1]) {
                priority = Enums_1.AWTEventPriority.Normal;
            }
            this._eventHandler.sendEventsForPriorityAndAbove(priority);
            this._timeout = -1;
            this.scheduleTimer();
        };
        AWTTransmissionManagerCore._initializeProfiles = function() {
            this._profiles = {};
            this._profiles[DataModels_1.AWT_REAL_TIME] = [ 4, 2, 1 ];
            this._profiles[DataModels_1.AWT_NEAR_REAL_TIME] = [ 12, 6, 3 ];
            this._profiles[DataModels_1.AWT_BEST_EFFORT] = [ 36, 18, 9 ];
        };
        AWTTransmissionManagerCore._newEventsAllowed = false;
        AWTTransmissionManagerCore._currentProfile = DataModels_1.AWT_REAL_TIME;
        AWTTransmissionManagerCore._timeout = -1;
        AWTTransmissionManagerCore._currentBackoffCount = 0;
        AWTTransmissionManagerCore._paused = false;
        AWTTransmissionManagerCore._timerCount = 0;
        AWTTransmissionManagerCore._lastUploadNowCall = 0;
        return AWTTransmissionManagerCore;
    }();
    exports.default = AWTTransmissionManagerCore;
}, function(module, exports, __webpack_require__) {
    "use strict";
    Object.defineProperty(exports, "__esModule", {
        value: true
    });
    var AWTNotificationManager = function() {
        function AWTNotificationManager() {}
        AWTNotificationManager.addNotificationListener = function(listener) {
            this.listeners.push(listener);
        };
        AWTNotificationManager.removeNotificationListener = function(listener) {
            var index = this.listeners.indexOf(listener);
            while (index > -1) {
                this.listeners.splice(index, 1);
                index = this.listeners.indexOf(listener);
            }
        };
        AWTNotificationManager.eventsSent = function(events) {
            var _this = this;
            var _loop_1 = function(i) {
                if (this_1.listeners[i].eventsSent) {
                    setTimeout(function() {
                        return _this.listeners[i].eventsSent(events);
                    }, 0);
                }
            };
            var this_1 = this;
            for (var i = 0; i < this.listeners.length; ++i) {
                _loop_1(i);
            }
        };
        AWTNotificationManager.eventsDropped = function(events, reason) {
            var _this = this;
            var _loop_2 = function(i) {
                if (this_2.listeners[i].eventsDropped) {
                    setTimeout(function() {
                        return _this.listeners[i].eventsDropped(events, reason);
                    }, 0);
                }
            };
            var this_2 = this;
            for (var i = 0; i < this.listeners.length; ++i) {
                _loop_2(i);
            }
        };
        AWTNotificationManager.eventsRetrying = function(events) {
            var _this = this;
            var _loop_3 = function(i) {
                if (this_3.listeners[i].eventsRetrying) {
                    setTimeout(function() {
                        return _this.listeners[i].eventsRetrying(events);
                    }, 0);
                }
            };
            var this_3 = this;
            for (var i = 0; i < this.listeners.length; ++i) {
                _loop_3(i);
            }
        };
        AWTNotificationManager.eventsRejected = function(events, reason) {
            var _this = this;
            var _loop_4 = function(i) {
                if (this_4.listeners[i].eventsRejected) {
                    setTimeout(function() {
                        return _this.listeners[i].eventsRejected(events, reason);
                    }, 0);
                }
            };
            var this_4 = this;
            for (var i = 0; i < this.listeners.length; ++i) {
                _loop_4(i);
            }
        };
        AWTNotificationManager.listeners = [];
        return AWTNotificationManager;
    }();
    exports.default = AWTNotificationManager;
}, function(module, exports, __webpack_require__) {
    "use strict";
    Object.defineProperty(exports, "__esModule", {
        value: true
    });
    var Enums_1 = __webpack_require__(0);
    exports.AWTPropertyType = Enums_1.AWTPropertyType;
    exports.AWTPiiKind = Enums_1.AWTPiiKind;
    exports.AWTEventPriority = Enums_1.AWTEventPriority;
    exports.AWTEventsDroppedReason = Enums_1.AWTEventsDroppedReason;
    exports.AWTEventsRejectedReason = Enums_1.AWTEventsRejectedReason;
    exports.AWTCustomerContentKind = Enums_1.AWTCustomerContentKind;
    var Enums_2 = __webpack_require__(6);
    exports.AWTUserIdType = Enums_2.AWTUserIdType;
    exports.AWTSessionState = Enums_2.AWTSessionState;
    var DataModels_1 = __webpack_require__(12);
    exports.AWT_BEST_EFFORT = DataModels_1.AWT_BEST_EFFORT;
    exports.AWT_NEAR_REAL_TIME = DataModels_1.AWT_NEAR_REAL_TIME;
    exports.AWT_REAL_TIME = DataModels_1.AWT_REAL_TIME;
    var AWTEventProperties_1 = __webpack_require__(7);
    exports.AWTEventProperties = AWTEventProperties_1.default;
    var AWTLogger_1 = __webpack_require__(13);
    exports.AWTLogger = AWTLogger_1.default;
    var AWTLogManager_1 = __webpack_require__(17);
    exports.AWTLogManager = AWTLogManager_1.default;
    var AWTTransmissionManager_1 = __webpack_require__(33);
    exports.AWTTransmissionManager = AWTTransmissionManager_1.default;
    var AWTSerializer_1 = __webpack_require__(15);
    exports.AWTSerializer = AWTSerializer_1.default;
    var AWTSemanticContext_1 = __webpack_require__(9);
    exports.AWTSemanticContext = AWTSemanticContext_1.default;
    exports.AWT_COLLECTOR_URL_UNITED_STATES = "https://us.pipe.aria.microsoft.com/Collector/3.0/";
    exports.AWT_COLLECTOR_URL_GERMANY = "https://de.pipe.aria.microsoft.com/Collector/3.0/";
    exports.AWT_COLLECTOR_URL_JAPAN = "https://jp.pipe.aria.microsoft.com/Collector/3.0/";
    exports.AWT_COLLECTOR_URL_AUSTRALIA = "https://au.pipe.aria.microsoft.com/Collector/3.0/";
    exports.AWT_COLLECTOR_URL_EUROPE = "https://eu.pipe.aria.microsoft.com/Collector/3.0/";
    exports.AWT_COLLECTOR_URL_USGOV_DOD = "https://pf.pipe.aria.microsoft.com/Collector/3.0";
    exports.AWT_COLLECTOR_URL_USGOV_DOJ = "https://tb.pipe.aria.microsoft.com/Collector/3.0";
}, function(module, exports, __webpack_require__) {
    "use strict";
    Object.defineProperty(exports, "__esModule", {
        value: true
    });
    var AWTUserIdType;
    (function(AWTUserIdType) {
        AWTUserIdType[AWTUserIdType["Unknown"] = 0] = "Unknown";
        AWTUserIdType[AWTUserIdType["MSACID"] = 1] = "MSACID";
        AWTUserIdType[AWTUserIdType["MSAPUID"] = 2] = "MSAPUID";
        AWTUserIdType[AWTUserIdType["ANID"] = 3] = "ANID";
        AWTUserIdType[AWTUserIdType["OrgIdCID"] = 4] = "OrgIdCID";
        AWTUserIdType[AWTUserIdType["OrgIdPUID"] = 5] = "OrgIdPUID";
        AWTUserIdType[AWTUserIdType["UserObjectId"] = 6] = "UserObjectId";
        AWTUserIdType[AWTUserIdType["Skype"] = 7] = "Skype";
        AWTUserIdType[AWTUserIdType["Yammer"] = 8] = "Yammer";
        AWTUserIdType[AWTUserIdType["EmailAddress"] = 9] = "EmailAddress";
        AWTUserIdType[AWTUserIdType["PhoneNumber"] = 10] = "PhoneNumber";
        AWTUserIdType[AWTUserIdType["SipAddress"] = 11] = "SipAddress";
        AWTUserIdType[AWTUserIdType["MUID"] = 12] = "MUID";
    })(AWTUserIdType = exports.AWTUserIdType || (exports.AWTUserIdType = {}));
    var AWTSessionState;
    (function(AWTSessionState) {
        AWTSessionState[AWTSessionState["Started"] = 0] = "Started";
        AWTSessionState[AWTSessionState["Ended"] = 1] = "Ended";
    })(AWTSessionState = exports.AWTSessionState || (exports.AWTSessionState = {}));
}, function(module, exports, __webpack_require__) {
    "use strict";
    Object.defineProperty(exports, "__esModule", {
        value: true
    });
    var Utils = __webpack_require__(2);
    var Enums_1 = __webpack_require__(0);
    var AWTEventProperties = function() {
        function AWTEventProperties(name) {
            this._event = {
                name: "",
                properties: {}
            };
            if (name) {
                this.setName(name);
            }
        }
        AWTEventProperties.prototype.setName = function(name) {
            this._event.name = name;
        };
        AWTEventProperties.prototype.getName = function() {
            return this._event.name;
        };
        AWTEventProperties.prototype.setType = function(type) {
            this._event.type = type;
        };
        AWTEventProperties.prototype.getType = function() {
            return this._event.type;
        };
        AWTEventProperties.prototype.setTimestamp = function(timestampInEpochMillis) {
            this._event.timestamp = timestampInEpochMillis;
        };
        AWTEventProperties.prototype.getTimestamp = function() {
            return this._event.timestamp;
        };
        AWTEventProperties.prototype.setEventPriority = function(priority) {
            this._event.priority = priority;
        };
        AWTEventProperties.prototype.getEventPriority = function() {
            return this._event.priority;
        };
        AWTEventProperties.prototype.setProperty = function(name, value, type) {
            if (type === void 0) {
                type = Enums_1.AWTPropertyType.Unspecified;
            }
            var property = {
                value: value,
                type: type,
                pii: Enums_1.AWTPiiKind.NotSet,
                cc: Enums_1.AWTCustomerContentKind.NotSet
            };
            property = Utils.sanitizeProperty(name, property);
            if (property === null) {
                delete this._event.properties[name];
                return;
            }
            this._event.properties[name] = property;
        };
        AWTEventProperties.prototype.setPropertyWithPii = function(name, value, pii, type) {
            if (type === void 0) {
                type = Enums_1.AWTPropertyType.Unspecified;
            }
            var property = {
                value: value,
                type: type,
                pii: pii,
                cc: Enums_1.AWTCustomerContentKind.NotSet
            };
            property = Utils.sanitizeProperty(name, property);
            if (property === null) {
                delete this._event.properties[name];
                return;
            }
            this._event.properties[name] = property;
        };
        AWTEventProperties.prototype.setPropertyWithCustomerContent = function(name, value, customerContent, type) {
            if (type === void 0) {
                type = Enums_1.AWTPropertyType.Unspecified;
            }
            var property = {
                value: value,
                type: type,
                pii: Enums_1.AWTPiiKind.NotSet,
                cc: customerContent
            };
            property = Utils.sanitizeProperty(name, property);
            if (property === null) {
                delete this._event.properties[name];
                return;
            }
            this._event.properties[name] = property;
        };
        AWTEventProperties.prototype.getPropertyMap = function() {
            return this._event.properties;
        };
        AWTEventProperties.prototype.getEvent = function() {
            return this._event;
        };
        return AWTEventProperties;
    }();
    exports.default = AWTEventProperties;
}, function(module, exports, __webpack_require__) {
    "use strict";
    Object.defineProperty(exports, "__esModule", {
        value: true
    });
    var Int64 = function() {
        function Int64(numberStr) {
            this.low = 0;
            this.high = 0;
            this.low = parseInt(numberStr, 10);
            if (this.low < 0) {
                this.high = -1;
            }
        }
        Int64.prototype._Equals = function(numberStr) {
            var tmp = new Int64(numberStr);
            return this.low === tmp.low && this.high === tmp.high;
        };
        return Int64;
    }();
    exports.Int64 = Int64;
    var UInt64 = function() {
        function UInt64(numberStr) {
            this.low = 0;
            this.high = 0;
            this.low = parseInt(numberStr, 10);
        }
        UInt64.prototype._Equals = function(numberStr) {
            var tmp = new UInt64(numberStr);
            return this.low === tmp.low && this.high === tmp.high;
        };
        return UInt64;
    }();
    exports.UInt64 = UInt64;
    var Number = function() {
        function Number() {}
        Number._ToByte = function(value) {
            return this._ToUInt8(value);
        };
        Number._ToUInt8 = function(value) {
            return value & 255;
        };
        Number._ToInt32 = function(value) {
            var signMask = value & 2147483648;
            return value & 2147483647 | signMask;
        };
        Number._ToUInt32 = function(value) {
            return value & 4294967295;
        };
        return Number;
    }();
    exports.Number = Number;
}, function(module, exports, __webpack_require__) {
    "use strict";
    Object.defineProperty(exports, "__esModule", {
        value: true
    });
    var AWTAutoCollection_1 = __webpack_require__(10);
    var Enums_1 = __webpack_require__(0);
    var Enums_2 = __webpack_require__(6);
    var UI_IDTYPE = "UserInfo.IdType";
    var AWTSemanticContext = function() {
        function AWTSemanticContext(_allowDeviceFields, _properties) {
            this._allowDeviceFields = _allowDeviceFields;
            this._properties = _properties;
        }
        AWTSemanticContext.prototype.setAppId = function(appId) {
            this._addContext("AppInfo.Id", appId);
        };
        AWTSemanticContext.prototype.setAppVersion = function(appVersion) {
            this._addContext("AppInfo.Version", appVersion);
        };
        AWTSemanticContext.prototype.setAppLanguage = function(appLanguage) {
            this._addContext("AppInfo.Language", appLanguage);
        };
        AWTSemanticContext.prototype.setDeviceId = function(deviceId) {
            if (this._allowDeviceFields) {
                AWTAutoCollection_1.default.checkAndSaveDeviceId(deviceId);
                this._addContext("DeviceInfo.Id", deviceId);
            }
        };
        AWTSemanticContext.prototype.setDeviceOsName = function(deviceOsName) {
            if (this._allowDeviceFields) {
                this._addContext("DeviceInfo.OsName", deviceOsName);
            }
        };
        AWTSemanticContext.prototype.setDeviceOsVersion = function(deviceOsVersion) {
            if (this._allowDeviceFields) {
                this._addContext("DeviceInfo.OsVersion", deviceOsVersion);
            }
        };
        AWTSemanticContext.prototype.setDeviceBrowserName = function(deviceBrowserName) {
            if (this._allowDeviceFields) {
                this._addContext("DeviceInfo.BrowserName", deviceBrowserName);
            }
        };
        AWTSemanticContext.prototype.setDeviceBrowserVersion = function(deviceBrowserVersion) {
            if (this._allowDeviceFields) {
                this._addContext("DeviceInfo.BrowserVersion", deviceBrowserVersion);
            }
        };
        AWTSemanticContext.prototype.setDeviceMake = function(deviceMake) {
            if (this._allowDeviceFields) {
                this._addContext("DeviceInfo.Make", deviceMake);
            }
        };
        AWTSemanticContext.prototype.setDeviceModel = function(deviceModel) {
            if (this._allowDeviceFields) {
                this._addContext("DeviceInfo.Model", deviceModel);
            }
        };
        AWTSemanticContext.prototype.setUserId = function(userId, pii, userIdType) {
            if (!isNaN(userIdType) && userIdType !== null && userIdType >= 0 && userIdType <= 12) {
                this._addContext(UI_IDTYPE, userIdType.toString());
            } else {
                var inferredUserIdType = void 0;
                switch (pii) {
                  case Enums_1.AWTPiiKind.SipAddress:
                    inferredUserIdType = Enums_2.AWTUserIdType.SipAddress;
                    break;

                  case Enums_1.AWTPiiKind.PhoneNumber:
                    inferredUserIdType = Enums_2.AWTUserIdType.PhoneNumber;
                    break;

                  case Enums_1.AWTPiiKind.SmtpAddress:
                    inferredUserIdType = Enums_2.AWTUserIdType.EmailAddress;
                    break;

                  default:
                    inferredUserIdType = Enums_2.AWTUserIdType.Unknown;
                    break;
                }
                this._addContext(UI_IDTYPE, inferredUserIdType.toString());
            }
            if (isNaN(pii) || pii === null || pii === Enums_1.AWTPiiKind.NotSet || pii > 13) {
                switch (userIdType) {
                  case Enums_2.AWTUserIdType.Skype:
                    pii = Enums_1.AWTPiiKind.Identity;
                    break;

                  case Enums_2.AWTUserIdType.EmailAddress:
                    pii = Enums_1.AWTPiiKind.SmtpAddress;
                    break;

                  case Enums_2.AWTUserIdType.PhoneNumber:
                    pii = Enums_1.AWTPiiKind.PhoneNumber;
                    break;

                  case Enums_2.AWTUserIdType.SipAddress:
                    pii = Enums_1.AWTPiiKind.SipAddress;
                    break;

                  default:
                    pii = Enums_1.AWTPiiKind.NotSet;
                    break;
                }
            }
            this._addContextWithPii("UserInfo.Id", userId, pii);
        };
        AWTSemanticContext.prototype.setUserAdvertisingId = function(userAdvertisingId) {
            this._addContext("UserInfo.AdvertisingId", userAdvertisingId);
        };
        AWTSemanticContext.prototype.setUserTimeZone = function(userTimeZone) {
            this._addContext("UserInfo.TimeZone", userTimeZone);
        };
        AWTSemanticContext.prototype.setUserLanguage = function(userLanguage) {
            this._addContext("UserInfo.Language", userLanguage);
        };
        AWTSemanticContext.prototype._addContext = function(key, value) {
            if (typeof value === "string") {
                this._properties.setProperty(key, value);
            }
        };
        AWTSemanticContext.prototype._addContextWithPii = function(key, value, pii) {
            if (typeof value === "string") {
                this._properties.setPropertyWithPii(key, value, pii);
            }
        };
        return AWTSemanticContext;
    }();
    exports.default = AWTSemanticContext;
}, function(module, exports, __webpack_require__) {
    "use strict";
    Object.defineProperty(exports, "__esModule", {
        value: true
    });
    var Utils = __webpack_require__(2);
    var DEVICE_ID_COOKIE = "MicrosoftApplicationsTelemetryDeviceId";
    var FIRSTLAUNCHTIME_COOKIE = "MicrosoftApplicationsTelemetryFirstLaunchTime";
    var BROWSERS = {
        MSIE: "MSIE",
        CHROME: "Chrome",
        FIREFOX: "Firefox",
        SAFARI: "Safari",
        EDGE: "Edge",
        ELECTRON: "Electron",
        SKYPE_SHELL: "SkypeShell",
        PHANTOMJS: "PhantomJS",
        OPERA: "Opera"
    };
    var OPERATING_SYSTEMS = {
        WINDOWS: "Windows",
        MACOSX: "Mac OS X",
        WINDOWS_PHONE: "Windows Phone",
        WINDOWS_RT: "Windows RT",
        IOS: "iOS",
        ANDROID: "Android",
        LINUX: "Linux",
        CROS: "Chrome OS",
        UNKNOWN: "Unknown"
    };
    var OSNAMEREGEX = {
        WIN: /(windows|win32)/i,
        WINRT: / arm;/i,
        WINPHONE: /windows\sphone\s\d+\.\d+/i,
        OSX: /(macintosh|mac os x)/i,
        IOS: /(iPad|iPhone|iPod)(?=.*like Mac OS X)/i,
        LINUX: /(linux|joli|[kxln]?ubuntu|debian|[open]*suse|gentoo|arch|slackware|fedora|mandriva|centos|pclinuxos|redhat|zenwalk)/i,
        ANDROID: /android/i,
        CROS: /CrOS/i
    };
    var VERSION_MAPPINGS = {
        5.1: "XP",
        "6.0": "Vista",
        6.1: "7",
        6.2: "8",
        6.3: "8.1",
        "10.0": "10"
    };
    var REGEX_VERSION = "([\\d,.]+)";
    var REGEX_VERSION_MAC = "([\\d,_,.]+)";
    var UNKNOWN = "Unknown";
    var UNDEFINED = "undefined";
    var AWTAutoCollection = function() {
        function AWTAutoCollection() {}
        AWTAutoCollection.addPropertyStorageOverride = function(propertyStorage) {
            if (propertyStorage) {
                this._propertyStorage = propertyStorage;
                return true;
            }
            return false;
        };
        AWTAutoCollection.autoCollect = function(semanticContext, disableCookies, userAgent) {
            this._semanticContext = semanticContext;
            this._disableCookies = disableCookies;
            this._autoCollect();
            if (!userAgent && typeof navigator !== UNDEFINED) {
                userAgent = navigator.userAgent || "";
            }
            this._autoCollectFromUserAgent(userAgent);
            if (this._disableCookies && !this._propertyStorage) {
                this._deleteCookie(DEVICE_ID_COOKIE);
                this._deleteCookie(FIRSTLAUNCHTIME_COOKIE);
                return;
            }
            if (this._propertyStorage || this._areCookiesAvailable && !this._disableCookies) {
                this._autoCollectDeviceId();
            }
        };
        AWTAutoCollection.checkAndSaveDeviceId = function(deviceId) {
            if (deviceId) {
                var oldDeviceId = this._getData(DEVICE_ID_COOKIE);
                var flt = this._getData(FIRSTLAUNCHTIME_COOKIE);
                if (oldDeviceId !== deviceId) {
                    flt = Utils.getISOString(new Date());
                }
                this._saveData(DEVICE_ID_COOKIE, deviceId);
                this._saveData(FIRSTLAUNCHTIME_COOKIE, flt);
                this._setFirstLaunchTime(flt);
            }
        };
        AWTAutoCollection._autoCollectDeviceId = function() {
            var deviceId = this._getData(DEVICE_ID_COOKIE);
            if (!deviceId) {
                deviceId = Utils.newGuid();
            }
            this._semanticContext.setDeviceId(deviceId);
        };
        AWTAutoCollection._autoCollect = function() {
            if (typeof document !== UNDEFINED && document.documentElement) {
                this._semanticContext.setAppLanguage(document.documentElement.lang);
            }
            if (typeof navigator !== UNDEFINED) {
                this._semanticContext.setUserLanguage(navigator.userLanguage || navigator.language);
            }
            var timeZone = new Date().getTimezoneOffset();
            var minutes = timeZone % 60;
            var hours = (timeZone - minutes) / 60;
            var timeZonePrefix = "+";
            if (hours > 0) {
                timeZonePrefix = "-";
            }
            hours = Math.abs(hours);
            minutes = Math.abs(minutes);
            this._semanticContext.setUserTimeZone(timeZonePrefix + (hours < 10 ? "0" + hours : hours.toString()) + ":" + (minutes < 10 ? "0" + minutes : minutes.toString()));
        };
        AWTAutoCollection._autoCollectFromUserAgent = function(userAgent) {
            if (userAgent) {
                var browserName = this._getBrowserName(userAgent);
                this._semanticContext.setDeviceBrowserName(browserName);
                this._semanticContext.setDeviceBrowserVersion(this._getBrowserVersion(userAgent, browserName));
                var osName = this._getOsName(userAgent);
                this._semanticContext.setDeviceOsName(osName);
                this._semanticContext.setDeviceOsVersion(this._getOsVersion(userAgent, osName));
            }
        };
        AWTAutoCollection._getBrowserName = function(userAgent) {
            if (this._userAgentContainsString("OPR/", userAgent)) {
                return BROWSERS.OPERA;
            }
            if (this._userAgentContainsString(BROWSERS.PHANTOMJS, userAgent)) {
                return BROWSERS.PHANTOMJS;
            }
            if (this._userAgentContainsString(BROWSERS.EDGE, userAgent)) {
                return BROWSERS.EDGE;
            }
            if (this._userAgentContainsString(BROWSERS.ELECTRON, userAgent)) {
                return BROWSERS.ELECTRON;
            }
            if (this._userAgentContainsString(BROWSERS.CHROME, userAgent)) {
                return BROWSERS.CHROME;
            }
            if (this._userAgentContainsString("Trident", userAgent)) {
                return BROWSERS.MSIE;
            }
            if (this._userAgentContainsString(BROWSERS.FIREFOX, userAgent)) {
                return BROWSERS.FIREFOX;
            }
            if (this._userAgentContainsString(BROWSERS.SAFARI, userAgent)) {
                return BROWSERS.SAFARI;
            }
            if (this._userAgentContainsString(BROWSERS.SKYPE_SHELL, userAgent)) {
                return BROWSERS.SKYPE_SHELL;
            }
            return UNKNOWN;
        };
        AWTAutoCollection._setFirstLaunchTime = function(flt) {
            if (!isNaN(flt)) {
                var fltDate = new Date();
                fltDate.setTime(parseInt(flt, 10));
                flt = Utils.getISOString(fltDate);
            }
            this.firstLaunchTime = flt;
        };
        AWTAutoCollection._userAgentContainsString = function(searchString, userAgent) {
            return userAgent.indexOf(searchString) > -1;
        };
        AWTAutoCollection._getBrowserVersion = function(userAgent, browserName) {
            if (browserName === BROWSERS.MSIE) {
                return this._getIeVersion(userAgent);
            } else {
                return this._getOtherVersion(browserName, userAgent);
            }
        };
        AWTAutoCollection._getIeVersion = function(userAgent) {
            var classicIeVersionMatches = userAgent.match(new RegExp(BROWSERS.MSIE + " " + REGEX_VERSION));
            if (classicIeVersionMatches) {
                return classicIeVersionMatches[1];
            } else {
                var ieVersionMatches = userAgent.match(new RegExp("rv:" + REGEX_VERSION));
                if (ieVersionMatches) {
                    return ieVersionMatches[1];
                }
            }
        };
        AWTAutoCollection._getOtherVersion = function(browserString, userAgent) {
            if (browserString === BROWSERS.SAFARI) {
                browserString = "Version";
            }
            var matches = userAgent.match(new RegExp(browserString + "/" + REGEX_VERSION));
            if (matches) {
                return matches[1];
            }
            return UNKNOWN;
        };
        AWTAutoCollection._getOsName = function(userAgent) {
            if (userAgent.match(OSNAMEREGEX.WINPHONE)) {
                return OPERATING_SYSTEMS.WINDOWS_PHONE;
            }
            if (userAgent.match(OSNAMEREGEX.WINRT)) {
                return OPERATING_SYSTEMS.WINDOWS_RT;
            }
            if (userAgent.match(OSNAMEREGEX.IOS)) {
                return OPERATING_SYSTEMS.IOS;
            }
            if (userAgent.match(OSNAMEREGEX.ANDROID)) {
                return OPERATING_SYSTEMS.ANDROID;
            }
            if (userAgent.match(OSNAMEREGEX.LINUX)) {
                return OPERATING_SYSTEMS.LINUX;
            }
            if (userAgent.match(OSNAMEREGEX.OSX)) {
                return OPERATING_SYSTEMS.MACOSX;
            }
            if (userAgent.match(OSNAMEREGEX.WIN)) {
                return OPERATING_SYSTEMS.WINDOWS;
            }
            if (userAgent.match(OSNAMEREGEX.CROS)) {
                return OPERATING_SYSTEMS.CROS;
            }
            return UNKNOWN;
        };
        AWTAutoCollection._getOsVersion = function(userAgent, osName) {
            if (osName === OPERATING_SYSTEMS.WINDOWS) {
                return this._getGenericOsVersion(userAgent, "Windows NT");
            }
            if (osName === OPERATING_SYSTEMS.ANDROID) {
                return this._getGenericOsVersion(userAgent, osName);
            }
            if (osName === OPERATING_SYSTEMS.MACOSX) {
                return this._getMacOsxVersion(userAgent);
            }
            return UNKNOWN;
        };
        AWTAutoCollection._getGenericOsVersion = function(userAgent, osName) {
            var ntVersionMatches = userAgent.match(new RegExp(osName + " " + REGEX_VERSION));
            if (ntVersionMatches) {
                if (VERSION_MAPPINGS[ntVersionMatches[1]]) {
                    return VERSION_MAPPINGS[ntVersionMatches[1]];
                }
                return ntVersionMatches[1];
            }
            return UNKNOWN;
        };
        AWTAutoCollection._getMacOsxVersion = function(userAgent) {
            var macOsxVersionInUserAgentMatches = userAgent.match(new RegExp(OPERATING_SYSTEMS.MACOSX + " " + REGEX_VERSION_MAC));
            if (macOsxVersionInUserAgentMatches) {
                var versionString = macOsxVersionInUserAgentMatches[1].replace(/_/g, ".");
                if (versionString) {
                    var delimiter = this._getDelimiter(versionString);
                    if (delimiter) {
                        var components = versionString.split(delimiter);
                        return components[0];
                    } else {
                        return versionString;
                    }
                }
            }
            return UNKNOWN;
        };
        AWTAutoCollection._getDelimiter = function(versionString) {
            if (versionString.indexOf(".") > -1) {
                return ".";
            }
            if (versionString.indexOf("_") > -1) {
                return "_";
            }
            return null;
        };
        AWTAutoCollection._saveData = function(name, value) {
            if (this._propertyStorage) {
                this._propertyStorage.setProperty(name, value);
            } else if (this._areCookiesAvailable) {
                var date = new Date();
                date.setTime(date.getTime() + 31536e6);
                var expires = "expires=" + date.toUTCString();
                document.cookie = name + "=" + value + "; " + expires;
            }
        };
        AWTAutoCollection._getData = function(name) {
            if (this._propertyStorage) {
                return this._propertyStorage.getProperty(name) || "";
            } else if (this._areCookiesAvailable) {
                name = name + "=";
                var ca = document.cookie.split(";");
                for (var i = 0; i < ca.length; i++) {
                    var c = ca[i];
                    var j = 0;
                    while (c.charAt(j) === " ") {
                        j++;
                    }
                    c = c.substring(j);
                    if (c.indexOf(name) === 0) {
                        return c.substring(name.length, c.length);
                    }
                }
            }
            return "";
        };
        AWTAutoCollection._deleteCookie = function(name) {
            if (this._areCookiesAvailable) {
                document.cookie = name + "=;expires=Thu, 01 Jan 1970 00:00:01 GMT;";
            }
        };
        AWTAutoCollection._disableCookies = false;
        AWTAutoCollection._areCookiesAvailable = typeof document !== UNDEFINED && typeof document.cookie !== UNDEFINED;
        return AWTAutoCollection;
    }();
    exports.default = AWTAutoCollection;
}, function(module, exports, __webpack_require__) {
    "use strict";
    Object.defineProperty(exports, "__esModule", {
        value: true
    });
    var Enums_1 = __webpack_require__(0);
    var AWTHttpManager_1 = __webpack_require__(22);
    var AWTTransmissionManagerCore_1 = __webpack_require__(3);
    var AWTRecordBatcher_1 = __webpack_require__(32);
    var AWTNotificationManager_1 = __webpack_require__(4);
    var UploadNowCheckTimer = 250;
    var MaxNumberEventPerBatch = 500;
    var MaxSendAttempts = 6;
    var AWTQueueManager = function() {
        function AWTQueueManager(collectorUrl, _queueSizeLimit, xhrOverride, clockSkewRefreshDurationInMins) {
            this._queueSizeLimit = _queueSizeLimit;
            this._isCurrentlyUploadingNow = false;
            this._uploadNowQueue = [];
            this._shouldDropEventsOnPause = false;
            this._paused = false;
            this._queueSize = 0;
            this._outboundQueue = [];
            this._inboundQueues = {};
            this._inboundQueues[Enums_1.AWTEventPriority.High] = [];
            this._inboundQueues[Enums_1.AWTEventPriority.Normal] = [];
            this._inboundQueues[Enums_1.AWTEventPriority.Low] = [];
            this._addEmptyQueues();
            this._batcher = new AWTRecordBatcher_1.default(this._outboundQueue, MaxNumberEventPerBatch);
            this._httpManager = new AWTHttpManager_1.default(this._outboundQueue, collectorUrl, this, xhrOverride, clockSkewRefreshDurationInMins);
        }
        AWTQueueManager.prototype.addEvent = function(event) {
            if (event.priority === Enums_1.AWTEventPriority.Immediate_sync) {
                this._httpManager.sendSynchronousRequest(this._batcher.addEventToBatch(event), event.apiKey);
            } else if (this._queueSize < this._queueSizeLimit) {
                this._addEventToProperQueue(event);
            } else {
                if (this._dropEventWithPriorityOrLess(event.priority)) {
                    this._addEventToProperQueue(event);
                } else {
                    AWTNotificationManager_1.default.eventsDropped([ event ], Enums_1.AWTEventsDroppedReason.QueueFull);
                }
            }
        };
        AWTQueueManager.prototype.sendEventsForPriorityAndAbove = function(priority) {
            this._batchEvents(priority);
            this._httpManager.sendQueuedRequests();
        };
        AWTQueueManager.prototype.hasEvents = function() {
            return (this._inboundQueues[Enums_1.AWTEventPriority.High][0].length > 0 || this._inboundQueues[Enums_1.AWTEventPriority.Normal][0].length > 0 || this._inboundQueues[Enums_1.AWTEventPriority.Low][0].length > 0 || this._batcher.hasBatch()) && this._httpManager.hasIdleConnection();
        };
        AWTQueueManager.prototype.addBackRequest = function(request) {
            if (!this._paused || !this._shouldDropEventsOnPause) {
                for (var token in request) {
                    if (request.hasOwnProperty(token)) {
                        for (var i = 0; i < request[token].length; ++i) {
                            if (request[token][i].sendAttempt < MaxSendAttempts) {
                                this.addEvent(request[token][i]);
                            } else {
                                AWTNotificationManager_1.default.eventsDropped([ request[token][i] ], Enums_1.AWTEventsDroppedReason.NonRetryableStatus);
                            }
                        }
                    }
                }
                AWTTransmissionManagerCore_1.default.scheduleTimer();
            }
        };
        AWTQueueManager.prototype.teardown = function() {
            if (!this._paused) {
                this._batchEvents(Enums_1.AWTEventPriority.Low);
                this._httpManager.teardown();
            }
        };
        AWTQueueManager.prototype.uploadNow = function(callback) {
            var _this = this;
            this._addEmptyQueues();
            if (!this._isCurrentlyUploadingNow) {
                this._isCurrentlyUploadingNow = true;
                setTimeout(function() {
                    return _this._uploadNow(callback);
                }, 0);
            } else {
                this._uploadNowQueue.push(callback);
            }
        };
        AWTQueueManager.prototype.pauseTransmission = function() {
            this._paused = true;
            this._httpManager.pause();
            if (this.shouldDropEventsOnPause) {
                this._queueSize -= this._inboundQueues[Enums_1.AWTEventPriority.High][0].length + this._inboundQueues[Enums_1.AWTEventPriority.Normal][0].length + this._inboundQueues[Enums_1.AWTEventPriority.Low][0].length;
                this._inboundQueues[Enums_1.AWTEventPriority.High][0] = [];
                this._inboundQueues[Enums_1.AWTEventPriority.Normal][0] = [];
                this._inboundQueues[Enums_1.AWTEventPriority.Low][0] = [];
                this._httpManager.removeQueuedRequests();
            }
        };
        AWTQueueManager.prototype.resumeTransmission = function() {
            this._paused = false;
            this._httpManager.resume();
        };
        AWTQueueManager.prototype.shouldDropEventsOnPause = function(shouldDropEventsOnPause) {
            this._shouldDropEventsOnPause = shouldDropEventsOnPause;
        };
        AWTQueueManager.prototype._removeFirstQueues = function() {
            this._inboundQueues[Enums_1.AWTEventPriority.High].shift();
            this._inboundQueues[Enums_1.AWTEventPriority.Normal].shift();
            this._inboundQueues[Enums_1.AWTEventPriority.Low].shift();
        };
        AWTQueueManager.prototype._addEmptyQueues = function() {
            this._inboundQueues[Enums_1.AWTEventPriority.High].push([]);
            this._inboundQueues[Enums_1.AWTEventPriority.Normal].push([]);
            this._inboundQueues[Enums_1.AWTEventPriority.Low].push([]);
        };
        AWTQueueManager.prototype._addEventToProperQueue = function(event) {
            if (!this._paused || !this._shouldDropEventsOnPause) {
                this._queueSize++;
                this._inboundQueues[event.priority][this._inboundQueues[event.priority].length - 1].push(event);
            }
        };
        AWTQueueManager.prototype._dropEventWithPriorityOrLess = function(priority) {
            var currentPriority = Enums_1.AWTEventPriority.Low;
            while (currentPriority <= priority) {
                if (this._inboundQueues[currentPriority][this._inboundQueues[currentPriority].length - 1].length > 0) {
                    AWTNotificationManager_1.default.eventsDropped([ this._inboundQueues[currentPriority][this._inboundQueues[currentPriority].length - 1].shift() ], Enums_1.AWTEventsDroppedReason.QueueFull);
                    return true;
                }
                currentPriority++;
            }
            return false;
        };
        AWTQueueManager.prototype._batchEvents = function(priority) {
            var priorityToProcess = Enums_1.AWTEventPriority.High;
            while (priorityToProcess >= priority) {
                while (this._inboundQueues[priorityToProcess][0].length > 0) {
                    var event_1 = this._inboundQueues[priorityToProcess][0].pop();
                    this._queueSize--;
                    this._batcher.addEventToBatch(event_1);
                }
                priorityToProcess--;
            }
            this._batcher.flushBatch();
        };
        AWTQueueManager.prototype._uploadNow = function(callback) {
            var _this = this;
            if (this.hasEvents()) {
                this.sendEventsForPriorityAndAbove(Enums_1.AWTEventPriority.Low);
            }
            this._checkOutboundQueueEmptyAndSent(function() {
                _this._removeFirstQueues();
                if (callback !== null && callback !== undefined) {
                    callback();
                }
                if (_this._uploadNowQueue.length > 0) {
                    setTimeout(function() {
                        return _this._uploadNow(_this._uploadNowQueue.shift());
                    }, 0);
                } else {
                    _this._isCurrentlyUploadingNow = false;
                    if (_this.hasEvents()) {
                        AWTTransmissionManagerCore_1.default.scheduleTimer();
                    }
                }
            });
        };
        AWTQueueManager.prototype._checkOutboundQueueEmptyAndSent = function(callback) {
            var _this = this;
            if (this._httpManager.isCompletelyIdle()) {
                callback();
            } else {
                setTimeout(function() {
                    return _this._checkOutboundQueueEmptyAndSent(callback);
                }, UploadNowCheckTimer);
            }
        };
        return AWTQueueManager;
    }();
    exports.default = AWTQueueManager;
}, function(module, exports, __webpack_require__) {
    "use strict";
    Object.defineProperty(exports, "__esModule", {
        value: true
    });
    exports.AWT_REAL_TIME = "REAL_TIME";
    exports.AWT_NEAR_REAL_TIME = "NEAR_REAL_TIME";
    exports.AWT_BEST_EFFORT = "BEST_EFFORT";
}, function(module, exports, __webpack_require__) {
    "use strict";
    Object.defineProperty(exports, "__esModule", {
        value: true
    });
    var Enums_1 = __webpack_require__(0);
    var Enums_2 = __webpack_require__(6);
    var AWTEventProperties_1 = __webpack_require__(7);
    var Utils = __webpack_require__(2);
    var AWTStatsManager_1 = __webpack_require__(14);
    var AWTNotificationManager_1 = __webpack_require__(4);
    var AWTTransmissionManagerCore_1 = __webpack_require__(3);
    var AWTLogManagerSettings_1 = __webpack_require__(18);
    var Version = __webpack_require__(16);
    var AWTSemanticContext_1 = __webpack_require__(9);
    var AWTAutoCollection_1 = __webpack_require__(10);
    var AWTLogger = function() {
        function AWTLogger(_apiKey) {
            this._apiKey = _apiKey;
            this._contextProperties = new AWTEventProperties_1.default();
            this._semanticContext = new AWTSemanticContext_1.default(false, this._contextProperties);
            this._sessionStartTime = 0;
            this._createInitId();
        }
        AWTLogger.prototype.setContext = function(name, value, type) {
            if (type === void 0) {
                type = Enums_1.AWTPropertyType.Unspecified;
            }
            this._contextProperties.setProperty(name, value, type);
        };
        AWTLogger.prototype.setContextWithPii = function(name, value, pii, type) {
            if (type === void 0) {
                type = Enums_1.AWTPropertyType.Unspecified;
            }
            this._contextProperties.setPropertyWithPii(name, value, pii, type);
        };
        AWTLogger.prototype.setContextWithCustomerContent = function(name, value, customerContent, type) {
            if (type === void 0) {
                type = Enums_1.AWTPropertyType.Unspecified;
            }
            this._contextProperties.setPropertyWithCustomerContent(name, value, customerContent, type);
        };
        AWTLogger.prototype.getSemanticContext = function() {
            return this._semanticContext;
        };
        AWTLogger.prototype.logEvent = function(event) {
            if (AWTLogManagerSettings_1.default.loggingEnabled) {
                if (!this._apiKey) {
                    this._apiKey = AWTLogManagerSettings_1.default.defaultTenantToken;
                    this._createInitId();
                }
                var sanitizeProperties = true;
                if (Utils.isString(event)) {
                    event = {
                        name: event
                    };
                } else if (event instanceof AWTEventProperties_1.default) {
                    event = event.getEvent();
                    sanitizeProperties = false;
                }
                AWTStatsManager_1.default.eventReceived(this._apiKey);
                AWTLogger._logEvent(AWTLogger._getInternalEvent(event, this._apiKey, sanitizeProperties), this._contextProperties);
            }
        };
        AWTLogger.prototype.logSession = function(state, properties) {
            if (AWTLogManagerSettings_1.default.sessionEnabled) {
                var sessionEvent = {
                    name: "session",
                    type: "session",
                    properties: {}
                };
                AWTLogger._addPropertiesToEvent(sessionEvent, properties);
                sessionEvent.priority = Enums_1.AWTEventPriority.High;
                if (state === Enums_2.AWTSessionState.Started) {
                    if (this._sessionStartTime > 0) {
                        return;
                    }
                    this._sessionStartTime = new Date().getTime();
                    this._sessionId = Utils.newGuid();
                    this.setContext("Session.Id", this._sessionId);
                    sessionEvent.properties["Session.State"] = "Started";
                } else if (state === Enums_2.AWTSessionState.Ended) {
                    if (this._sessionStartTime === 0) {
                        return;
                    }
                    var sessionDurationSec = Math.floor((new Date().getTime() - this._sessionStartTime) / 1e3);
                    sessionEvent.properties["Session.Id"] = this._sessionId;
                    sessionEvent.properties["Session.State"] = "Ended";
                    sessionEvent.properties["Session.Duration"] = sessionDurationSec.toString();
                    sessionEvent.properties["Session.DurationBucket"] = AWTLogger._getSessionDurationFromTime(sessionDurationSec);
                    this._sessionStartTime = 0;
                    this.setContext("Session.Id", null);
                    this._sessionId = undefined;
                } else {
                    return;
                }
                sessionEvent.properties["Session.FirstLaunchTime"] = AWTAutoCollection_1.default.firstLaunchTime;
                this.logEvent(sessionEvent);
            }
        };
        AWTLogger.prototype.getSessionId = function() {
            return this._sessionId;
        };
        AWTLogger.prototype.logFailure = function(signature, detail, category, id, properties) {
            if (!signature || !detail) {
                return;
            }
            var failureEvent = {
                name: "failure",
                type: "failure",
                properties: {}
            };
            AWTLogger._addPropertiesToEvent(failureEvent, properties);
            failureEvent.properties["Failure.Signature"] = signature;
            failureEvent.properties["Failure.Detail"] = detail;
            if (category) {
                failureEvent.properties["Failure.Category"] = category;
            }
            if (id) {
                failureEvent.properties["Failure.Id"] = id;
            }
            failureEvent.priority = Enums_1.AWTEventPriority.High;
            this.logEvent(failureEvent);
        };
        AWTLogger.prototype.logPageView = function(id, pageName, category, uri, referrerUri, properties) {
            if (!id || !pageName) {
                return;
            }
            var pageViewEvent = {
                name: "pageview",
                type: "pageview",
                properties: {}
            };
            AWTLogger._addPropertiesToEvent(pageViewEvent, properties);
            pageViewEvent.properties["PageView.Id"] = id;
            pageViewEvent.properties["PageView.Name"] = pageName;
            if (category) {
                pageViewEvent.properties["PageView.Category"] = category;
            }
            if (uri) {
                pageViewEvent.properties["PageView.Uri"] = uri;
            }
            if (referrerUri) {
                pageViewEvent.properties["PageView.ReferrerUri"] = referrerUri;
            }
            this.logEvent(pageViewEvent);
        };
        AWTLogger.prototype._createInitId = function() {
            if (!AWTLogger._initIdMap[this._apiKey] && this._apiKey) {
                AWTLogger._initIdMap[this._apiKey] = Utils.newGuid();
            }
        };
        AWTLogger._addPropertiesToEvent = function(event, propertiesEvent) {
            if (propertiesEvent) {
                if (propertiesEvent instanceof AWTEventProperties_1.default) {
                    propertiesEvent = propertiesEvent.getEvent();
                }
                if (propertiesEvent.name) {
                    event.name = propertiesEvent.name;
                }
                if (propertiesEvent.priority) {
                    event.priority = propertiesEvent.priority;
                }
                for (var name_1 in propertiesEvent.properties) {
                    if (propertiesEvent.properties.hasOwnProperty(name_1)) {
                        event.properties[name_1] = propertiesEvent.properties[name_1];
                    }
                }
            }
        };
        AWTLogger._getSessionDurationFromTime = function(timeInSec) {
            if (timeInSec < 0) {
                return "Undefined";
            } else if (timeInSec <= 3) {
                return "UpTo3Sec";
            } else if (timeInSec <= 10) {
                return "UpTo10Sec";
            } else if (timeInSec <= 30) {
                return "UpTo30Sec";
            } else if (timeInSec <= 60) {
                return "UpTo60Sec";
            } else if (timeInSec <= 180) {
                return "UpTo3Min";
            } else if (timeInSec <= 600) {
                return "UpTo10Min";
            } else if (timeInSec <= 1800) {
                return "UpTo30Min";
            }
            return "Above30Min";
        };
        AWTLogger._logEvent = function(eventWithMetaData, contextProperties) {
            if (!eventWithMetaData.name || !Utils.isString(eventWithMetaData.name)) {
                AWTNotificationManager_1.default.eventsRejected([ eventWithMetaData ], Enums_1.AWTEventsRejectedReason.InvalidEvent);
                return;
            }
            eventWithMetaData.name = eventWithMetaData.name.toLowerCase();
            eventWithMetaData.name = eventWithMetaData.name.replace(Utils.EventNameDotRegex, "_");
            if (!eventWithMetaData.type || !Utils.isString(eventWithMetaData.type)) {
                eventWithMetaData.type = "custom";
            } else {
                eventWithMetaData.type = eventWithMetaData.type.toLowerCase();
            }
            if (!Utils.EventNameAndTypeRegex.test(eventWithMetaData.name) || !Utils.EventNameAndTypeRegex.test(eventWithMetaData.type)) {
                AWTNotificationManager_1.default.eventsRejected([ eventWithMetaData ], Enums_1.AWTEventsRejectedReason.InvalidEvent);
                return;
            }
            if (!Utils.isNumber(eventWithMetaData.timestamp) || eventWithMetaData.timestamp < 0) {
                eventWithMetaData.timestamp = new Date().getTime();
            }
            if (!eventWithMetaData.properties) {
                eventWithMetaData.properties = {};
            }
            this._addContextIfAbsent(eventWithMetaData, contextProperties.getPropertyMap());
            this._addContextIfAbsent(eventWithMetaData, AWTLogManagerSettings_1.default.logManagerContext.getPropertyMap());
            this._setDefaultProperty(eventWithMetaData, "EventInfo.InitId", this._getInitId(eventWithMetaData.apiKey));
            this._setDefaultProperty(eventWithMetaData, "EventInfo.Sequence", this._getSequenceId(eventWithMetaData.apiKey));
            this._setDefaultProperty(eventWithMetaData, "EventInfo.SdkVersion", Version.FullVersionString);
            this._setDefaultProperty(eventWithMetaData, "EventInfo.Name", eventWithMetaData.name);
            this._setDefaultProperty(eventWithMetaData, "EventInfo.Time", new Date(eventWithMetaData.timestamp).toISOString());
            if (!Utils.isPriority(eventWithMetaData.priority)) {
                eventWithMetaData.priority = Enums_1.AWTEventPriority.Normal;
            }
            this._sendEvent(eventWithMetaData);
        };
        AWTLogger._addContextIfAbsent = function(event, contextProperties) {
            if (contextProperties) {
                for (var name_2 in contextProperties) {
                    if (contextProperties.hasOwnProperty(name_2)) {
                        if (!event.properties[name_2]) {
                            event.properties[name_2] = contextProperties[name_2];
                        }
                    }
                }
            }
        };
        AWTLogger._setDefaultProperty = function(event, name, value) {
            event.properties[name] = {
                value: value,
                pii: Enums_1.AWTPiiKind.NotSet,
                type: Enums_1.AWTPropertyType.String
            };
        };
        AWTLogger._sendEvent = function(event) {
            AWTTransmissionManagerCore_1.default.sendEvent(event);
        };
        AWTLogger._getInternalEvent = function(event, apiKey, sanitizeProperties) {
            event.properties = event.properties || {};
            if (sanitizeProperties) {
                for (var name_3 in event.properties) {
                    if (event.properties.hasOwnProperty(name_3)) {
                        event.properties[name_3] = Utils.sanitizeProperty(name_3, event.properties[name_3]);
                        if (event.properties[name_3] === null) {
                            delete event.properties[name_3];
                        }
                    }
                }
            }
            var internalEvent = event;
            internalEvent.id = Utils.newGuid();
            internalEvent.apiKey = apiKey;
            return internalEvent;
        };
        AWTLogger._getInitId = function(apiKey) {
            return AWTLogger._initIdMap[apiKey];
        };
        AWTLogger._getSequenceId = function(apiKey) {
            if (AWTLogger._sequenceIdMap[apiKey] === undefined) {
                AWTLogger._sequenceIdMap[apiKey] = 0;
            }
            return (++AWTLogger._sequenceIdMap[apiKey]).toString();
        };
        AWTLogger._sequenceIdMap = {};
        AWTLogger._initIdMap = {};
        return AWTLogger;
    }();
    exports.default = AWTLogger;
}, function(module, exports, __webpack_require__) {
    "use strict";
    Object.defineProperty(exports, "__esModule", {
        value: true
    });
    var Utils = __webpack_require__(2);
    var AWTNotificationManager_1 = __webpack_require__(4);
    var Enums_1 = __webpack_require__(0);
    var StatsTimer = 6e4;
    var AWTStatsManager = function() {
        function AWTStatsManager() {}
        AWTStatsManager.initialize = function(sendStats) {
            var _this = this;
            this._sendStats = sendStats;
            this._isInitalized = true;
            AWTNotificationManager_1.default.addNotificationListener({
                eventsSent: function(events) {
                    _this._addStat("records_sent_count", events.length, events[0].apiKey);
                },
                eventsDropped: function(events, reason) {
                    switch (reason) {
                      case Enums_1.AWTEventsDroppedReason.NonRetryableStatus:
                        _this._addStat("d_send_fail", events.length, events[0].apiKey);
                        _this._addStat("records_dropped_count", events.length, events[0].apiKey);
                        break;

                      case Enums_1.AWTEventsDroppedReason.QueueFull:
                        _this._addStat("d_queue_full", events.length, events[0].apiKey);
                        break;
                    }
                },
                eventsRejected: function(events, reason) {
                    switch (reason) {
                      case Enums_1.AWTEventsRejectedReason.InvalidEvent:
                        _this._addStat("r_inv", events.length, events[0].apiKey);
                        break;

                      case Enums_1.AWTEventsRejectedReason.KillSwitch:
                        _this._addStat("r_kl", events.length, events[0].apiKey);
                        break;

                      case Enums_1.AWTEventsRejectedReason.SizeLimitExceeded:
                        _this._addStat("r_size", events.length, events[0].apiKey);
                        break;
                    }
                    _this._addStat("r_count", events.length, events[0].apiKey);
                },
                eventsRetrying: null
            });
            setTimeout(function() {
                return _this.flush();
            }, StatsTimer);
        };
        AWTStatsManager.teardown = function() {
            if (this._isInitalized) {
                this.flush();
                this._isInitalized = false;
            }
        };
        AWTStatsManager.eventReceived = function(apiKey) {
            AWTStatsManager._addStat("records_received_count", 1, apiKey);
        };
        AWTStatsManager.flush = function() {
            var _this = this;
            if (this._isInitalized) {
                for (var tenantId in this._stats) {
                    if (this._stats.hasOwnProperty(tenantId)) {
                        this._sendStats(this._stats[tenantId], tenantId);
                    }
                }
                this._stats = {};
                setTimeout(function() {
                    return _this.flush();
                }, StatsTimer);
            }
        };
        AWTStatsManager._addStat = function(statName, value, apiKey) {
            if (this._isInitalized && apiKey !== Utils.StatsApiKey) {
                var tenantId = Utils.getTenantId(apiKey);
                if (!this._stats[tenantId]) {
                    this._stats[tenantId] = {};
                }
                if (!this._stats[tenantId][statName]) {
                    this._stats[tenantId][statName] = value;
                } else {
                    this._stats[tenantId][statName] = this._stats[tenantId][statName] + value;
                }
            }
        };
        AWTStatsManager._isInitalized = false;
        AWTStatsManager._stats = {};
        return AWTStatsManager;
    }();
    exports.default = AWTStatsManager;
}, function(module, exports, __webpack_require__) {
    "use strict";
    Object.defineProperty(exports, "__esModule", {
        value: true
    });
    var Bond = __webpack_require__(23);
    var Enums_1 = __webpack_require__(0);
    var AWTNotificationManager_1 = __webpack_require__(4);
    var Utils = __webpack_require__(2);
    var RequestSizeLimitBytes = 2936012;
    var AWTSerializer = function() {
        function AWTSerializer() {}
        AWTSerializer.getPayloadBlob = function(requestDictionary, tokenCount) {
            var requestFull = false;
            var remainingRequest;
            var stream = new Bond.IO.MemoryStream();
            var writer = new Bond.CompactBinaryProtocolWriter(stream);
            writer._WriteFieldBegin(Bond._BondDataType._BT_MAP, 3, null);
            writer._WriteMapContainerBegin(tokenCount, Bond._BondDataType._BT_STRING, Bond._BondDataType._BT_LIST);
            for (var token in requestDictionary) {
                if (!requestFull) {
                    if (requestDictionary.hasOwnProperty(token)) {
                        writer._WriteString(token);
                        var dataPackage = requestDictionary[token];
                        writer._WriteContainerBegin(1, Bond._BondDataType._BT_STRUCT);
                        writer._WriteFieldBegin(Bond._BondDataType._BT_STRING, 2, null);
                        writer._WriteString("act_default_source");
                        writer._WriteFieldBegin(Bond._BondDataType._BT_STRING, 5, null);
                        writer._WriteString(Utils.newGuid());
                        writer._WriteFieldBegin(Bond._BondDataType._BT_INT64, 6, null);
                        writer._WriteInt64(Utils.numberToBondInt64(Date.now()));
                        writer._WriteFieldBegin(Bond._BondDataType._BT_LIST, 8, null);
                        var dpSizePos = stream._GetBuffer().length + 1;
                        writer._WriteContainerBegin(requestDictionary[token].length, Bond._BondDataType._BT_STRUCT);
                        var dpSizeSerialized = stream._GetBuffer().length - dpSizePos;
                        for (var i = 0; i < dataPackage.length; ++i) {
                            var currentStreamPos = stream._GetBuffer().length;
                            this.writeEvent(dataPackage[i], writer);
                            if (stream._GetBuffer().length - currentStreamPos > RequestSizeLimitBytes) {
                                AWTNotificationManager_1.default.eventsRejected([ dataPackage[i] ], Enums_1.AWTEventsRejectedReason.SizeLimitExceeded);
                                dataPackage.splice(i--, 1);
                                stream._GetBuffer().splice(currentStreamPos);
                                this._addNewDataPackageSize(dataPackage.length, stream, dpSizeSerialized, dpSizePos);
                                continue;
                            }
                            if (stream._GetBuffer().length > RequestSizeLimitBytes) {
                                stream._GetBuffer().splice(currentStreamPos);
                                if (!remainingRequest) {
                                    remainingRequest = {};
                                }
                                requestDictionary[token] = dataPackage.splice(0, i);
                                remainingRequest[token] = dataPackage;
                                this._addNewDataPackageSize(requestDictionary[token].length, stream, dpSizeSerialized, dpSizePos);
                                break;
                            }
                        }
                        writer._WriteStructEnd(false);
                    }
                } else {
                    if (!remainingRequest) {
                        remainingRequest = {};
                    }
                    remainingRequest[token] = requestDictionary[token];
                    delete requestDictionary[token];
                }
            }
            writer._WriteStructEnd(false);
            return {
                payloadBlob: stream._GetBuffer(),
                remainingRequest: remainingRequest
            };
        };
        AWTSerializer._addNewDataPackageSize = function(size, stream, oldDpSize, streamPos) {
            var newRecordCountSerialized = Bond._Encoding._Varint_GetBytes(Bond.Number._ToUInt32(size));
            for (var j = 0; j < oldDpSize; ++j) {
                if (j < newRecordCountSerialized.length) {
                    stream._GetBuffer()[streamPos + j] = newRecordCountSerialized[j];
                } else {
                    stream._GetBuffer().slice(streamPos + j, oldDpSize - j);
                    break;
                }
            }
        };
        AWTSerializer.writeEvent = function(eventData, writer) {
            writer._WriteFieldBegin(Bond._BondDataType._BT_STRING, 1, null);
            writer._WriteString(eventData.id);
            writer._WriteFieldBegin(Bond._BondDataType._BT_INT64, 3, null);
            writer._WriteInt64(Utils.numberToBondInt64(eventData.timestamp));
            writer._WriteFieldBegin(Bond._BondDataType._BT_STRING, 5, null);
            writer._WriteString(eventData.type);
            writer._WriteFieldBegin(Bond._BondDataType._BT_STRING, 6, null);
            writer._WriteString(eventData.name);
            var propsString = {};
            var propStringCount = 0;
            var propsInt64 = {};
            var propInt64Count = 0;
            var propsDouble = {};
            var propDoubleCount = 0;
            var propsBool = {};
            var propBoolCount = 0;
            var propsDate = {};
            var propDateCount = 0;
            var piiProps = {};
            var piiPropCount = 0;
            var ccProps = {};
            var ccPropCount = 0;
            for (var key in eventData.properties) {
                if (eventData.properties.hasOwnProperty(key)) {
                    var property = eventData.properties[key];
                    if (property.cc > 0) {
                        ccProps[key] = property;
                        ccPropCount++;
                    } else if (property.pii > 0) {
                        piiProps[key] = property;
                        piiPropCount++;
                    } else {
                        switch (property.type) {
                          case Enums_1.AWTPropertyType.String:
                            propsString[key] = property.value;
                            propStringCount++;
                            break;

                          case Enums_1.AWTPropertyType.Int64:
                            propsInt64[key] = property.value;
                            propInt64Count++;
                            break;

                          case Enums_1.AWTPropertyType.Double:
                            propsDouble[key] = property.value;
                            propDoubleCount++;
                            break;

                          case Enums_1.AWTPropertyType.Boolean:
                            propsBool[key] = property.value;
                            propBoolCount++;
                            break;

                          case Enums_1.AWTPropertyType.Date:
                            propsDate[key] = property.value;
                            propDateCount++;
                            break;
                        }
                    }
                }
            }
            if (propStringCount) {
                writer._WriteFieldBegin(Bond._BondDataType._BT_MAP, 13, null);
                writer._WriteMapContainerBegin(propStringCount, Bond._BondDataType._BT_STRING, Bond._BondDataType._BT_STRING);
                for (var key in propsString) {
                    if (propsString.hasOwnProperty(key)) {
                        var value = propsString[key];
                        writer._WriteString(key);
                        writer._WriteString(value.toString());
                    }
                }
            }
            if (piiPropCount) {
                writer._WriteFieldBegin(Bond._BondDataType._BT_MAP, 30, null);
                writer._WriteMapContainerBegin(piiPropCount, Bond._BondDataType._BT_STRING, Bond._BondDataType._BT_STRUCT);
                for (var key in piiProps) {
                    if (piiProps.hasOwnProperty(key)) {
                        var property = piiProps[key];
                        writer._WriteString(key);
                        writer._WriteFieldBegin(Bond._BondDataType._BT_INT32, 1, null);
                        writer._WriteInt32(1);
                        writer._WriteFieldBegin(Bond._BondDataType._BT_INT32, 2, null);
                        writer._WriteInt32(property.pii);
                        writer._WriteFieldBegin(Bond._BondDataType._BT_STRING, 3, null);
                        writer._WriteString(property.value.toString());
                        writer._WriteStructEnd(false);
                    }
                }
            }
            if (propBoolCount) {
                writer._WriteFieldBegin(Bond._BondDataType._BT_MAP, 31, null);
                writer._WriteMapContainerBegin(propBoolCount, Bond._BondDataType._BT_STRING, Bond._BondDataType._BT_BOOL);
                for (var key in propsBool) {
                    if (propsBool.hasOwnProperty(key)) {
                        var value = propsBool[key];
                        writer._WriteString(key);
                        writer._WriteBool(value);
                    }
                }
            }
            if (propDateCount) {
                writer._WriteFieldBegin(Bond._BondDataType._BT_MAP, 32, null);
                writer._WriteMapContainerBegin(propDateCount, Bond._BondDataType._BT_STRING, Bond._BondDataType._BT_INT64);
                for (var key in propsDate) {
                    if (propsDate.hasOwnProperty(key)) {
                        var value = propsDate[key];
                        writer._WriteString(key);
                        writer._WriteInt64(Utils.numberToBondInt64(value));
                    }
                }
            }
            if (propInt64Count) {
                writer._WriteFieldBegin(Bond._BondDataType._BT_MAP, 33, null);
                writer._WriteMapContainerBegin(propInt64Count, Bond._BondDataType._BT_STRING, Bond._BondDataType._BT_INT64);
                for (var key in propsInt64) {
                    if (propsInt64.hasOwnProperty(key)) {
                        var value = propsInt64[key];
                        writer._WriteString(key);
                        writer._WriteInt64(Utils.numberToBondInt64(value));
                    }
                }
            }
            if (propDoubleCount) {
                writer._WriteFieldBegin(Bond._BondDataType._BT_MAP, 34, null);
                writer._WriteMapContainerBegin(propDoubleCount, Bond._BondDataType._BT_STRING, Bond._BondDataType._BT_DOUBLE);
                for (var key in propsDouble) {
                    if (propsDouble.hasOwnProperty(key)) {
                        var value = propsDouble[key];
                        writer._WriteString(key);
                        writer._WriteDouble(value);
                    }
                }
            }
            if (ccPropCount) {
                writer._WriteFieldBegin(Bond._BondDataType._BT_MAP, 36, null);
                writer._WriteMapContainerBegin(ccPropCount, Bond._BondDataType._BT_STRING, Bond._BondDataType._BT_STRUCT);
                for (var key in ccProps) {
                    if (ccProps.hasOwnProperty(key)) {
                        var property = ccProps[key];
                        writer._WriteString(key);
                        writer._WriteFieldBegin(Bond._BondDataType._BT_INT32, 1, null);
                        writer._WriteInt32(property.cc);
                        writer._WriteFieldBegin(Bond._BondDataType._BT_STRING, 2, null);
                        writer._WriteString(property.value.toString());
                        writer._WriteStructEnd(false);
                    }
                }
            }
            writer._WriteStructEnd(false);
        };
        AWTSerializer.base64Encode = function(data) {
            return Bond._Encoding._Base64_GetString(data);
        };
        return AWTSerializer;
    }();
    exports.default = AWTSerializer;
}, function(module, exports, __webpack_require__) {
    "use strict";
    Object.defineProperty(exports, "__esModule", {
        value: true
    });
    exports.Version = "1.8.1";
    exports.FullVersionString = "AWT-Web-JS-" + exports.Version;
}, function(module, exports, __webpack_require__) {
    "use strict";
    Object.defineProperty(exports, "__esModule", {
        value: true
    });
    var Enums_1 = __webpack_require__(0);
    var Enums_2 = __webpack_require__(6);
    var AWTLogManagerSettings_1 = __webpack_require__(18);
    var AWTLogger_1 = __webpack_require__(13);
    var AWTTransmissionManagerCore_1 = __webpack_require__(3);
    var AWTNotificationManager_1 = __webpack_require__(4);
    var AWTAutoCollection_1 = __webpack_require__(10);
    var AWTLogManager = function() {
        function AWTLogManager() {}
        AWTLogManager.initialize = function(tenantToken, configuration) {
            if (configuration === void 0) {
                configuration = {};
            }
            if (this._isInitialized) {
                return;
            }
            this._isInitialized = true;
            AWTLogManagerSettings_1.default.defaultTenantToken = tenantToken;
            this._overrideValuesFromConfig(configuration);
            if (this._config.disableCookiesUsage && !this._config.propertyStorageOverride) {
                AWTLogManagerSettings_1.default.sessionEnabled = false;
            }
            AWTAutoCollection_1.default.addPropertyStorageOverride(this._config.propertyStorageOverride);
            AWTAutoCollection_1.default.autoCollect(AWTLogManagerSettings_1.default.semanticContext, this._config.disableCookiesUsage, this._config.userAgent);
            AWTTransmissionManagerCore_1.default.initialize(this._config);
            AWTLogManagerSettings_1.default.loggingEnabled = true;
            if (this._config.enableAutoUserSession) {
                this.getLogger().logSession(Enums_2.AWTSessionState.Started);
                window.addEventListener("beforeunload", this.flushAndTeardown);
            }
            return this.getLogger();
        };
        AWTLogManager.getSemanticContext = function() {
            return AWTLogManagerSettings_1.default.semanticContext;
        };
        AWTLogManager.flush = function(callback) {
            if (this._isInitialized && !this._isDestroyed) {
                AWTTransmissionManagerCore_1.default.flush(callback);
            }
        };
        AWTLogManager.flushAndTeardown = function() {
            if (this._isInitialized && !this._isDestroyed) {
                if (this._config.enableAutoUserSession) {
                    this.getLogger().logSession(Enums_2.AWTSessionState.Ended);
                }
                AWTTransmissionManagerCore_1.default.flushAndTeardown();
                AWTLogManagerSettings_1.default.loggingEnabled = false;
                this._isDestroyed = true;
            }
        };
        AWTLogManager.pauseTransmission = function() {
            if (this._isInitialized && !this._isDestroyed) {
                AWTTransmissionManagerCore_1.default.pauseTransmission();
            }
        };
        AWTLogManager.resumeTransmision = function() {
            if (this._isInitialized && !this._isDestroyed) {
                AWTTransmissionManagerCore_1.default.resumeTransmision();
            }
        };
        AWTLogManager.setTransmitProfile = function(profileName) {
            if (this._isInitialized && !this._isDestroyed) {
                AWTTransmissionManagerCore_1.default.setTransmitProfile(profileName);
            }
        };
        AWTLogManager.loadTransmitProfiles = function(profiles) {
            if (this._isInitialized && !this._isDestroyed) {
                AWTTransmissionManagerCore_1.default.loadTransmitProfiles(profiles);
            }
        };
        AWTLogManager.setContext = function(name, value, type) {
            if (type === void 0) {
                type = Enums_1.AWTPropertyType.Unspecified;
            }
            AWTLogManagerSettings_1.default.logManagerContext.setProperty(name, value, type);
        };
        AWTLogManager.setContextWithPii = function(name, value, pii, type) {
            if (type === void 0) {
                type = Enums_1.AWTPropertyType.Unspecified;
            }
            AWTLogManagerSettings_1.default.logManagerContext.setPropertyWithPii(name, value, pii, type);
        };
        AWTLogManager.setContextWithCustomerContent = function(name, value, customerContent, type) {
            if (type === void 0) {
                type = Enums_1.AWTPropertyType.Unspecified;
            }
            AWTLogManagerSettings_1.default.logManagerContext.setPropertyWithCustomerContent(name, value, customerContent, type);
        };
        AWTLogManager.getLogger = function(tenantToken) {
            var key = tenantToken;
            if (!key || key === AWTLogManagerSettings_1.default.defaultTenantToken) {
                key = "";
            }
            if (!this._loggers[key]) {
                this._loggers[key] = new AWTLogger_1.default(key);
            }
            return this._loggers[key];
        };
        AWTLogManager.addNotificationListener = function(listener) {
            AWTNotificationManager_1.default.addNotificationListener(listener);
        };
        AWTLogManager.removeNotificationListener = function(listener) {
            AWTNotificationManager_1.default.removeNotificationListener(listener);
        };
        AWTLogManager._overrideValuesFromConfig = function(config) {
            if (config.collectorUri) {
                this._config.collectorUri = config.collectorUri;
            }
            if (config.cacheMemorySizeLimitInNumberOfEvents > 0) {
                this._config.cacheMemorySizeLimitInNumberOfEvents = config.cacheMemorySizeLimitInNumberOfEvents;
            }
            if (config.httpXHROverride && config.httpXHROverride.sendPOST) {
                this._config.httpXHROverride = config.httpXHROverride;
            }
            if (config.propertyStorageOverride && config.propertyStorageOverride.getProperty && config.propertyStorageOverride.setProperty) {
                this._config.propertyStorageOverride = config.propertyStorageOverride;
            }
            if (config.userAgent) {
                this._config.userAgent = config.userAgent;
            }
            if (config.disableCookiesUsage) {
                this._config.disableCookiesUsage = config.disableCookiesUsage;
            }
            if (config.canSendStatEvent) {
                this._config.canSendStatEvent = config.canSendStatEvent;
            }
            if (config.enableAutoUserSession && typeof window !== "undefined" && window.addEventListener) {
                this._config.enableAutoUserSession = config.enableAutoUserSession;
            }
            if (config.clockSkewRefreshDurationInMins > 0) {
                this._config.clockSkewRefreshDurationInMins = config.clockSkewRefreshDurationInMins;
            }
        };
        AWTLogManager._loggers = {};
        AWTLogManager._isInitialized = false;
        AWTLogManager._isDestroyed = false;
        AWTLogManager._config = {
            collectorUri: "https://browser.pipe.aria.microsoft.com/Collector/3.0/",
            cacheMemorySizeLimitInNumberOfEvents: 1e4,
            disableCookiesUsage: false,
            canSendStatEvent: function(eventName) {
                return true;
            },
            clockSkewRefreshDurationInMins: 0
        };
        return AWTLogManager;
    }();
    exports.default = AWTLogManager;
}, function(module, exports, __webpack_require__) {
    "use strict";
    Object.defineProperty(exports, "__esModule", {
        value: true
    });
    var AWTEventProperties_1 = __webpack_require__(7);
    var AWTSemanticContext_1 = __webpack_require__(9);
    var AWTLogManagerSettings = function() {
        function AWTLogManagerSettings() {}
        AWTLogManagerSettings.logManagerContext = new AWTEventProperties_1.default();
        AWTLogManagerSettings.sessionEnabled = true;
        AWTLogManagerSettings.loggingEnabled = false;
        AWTLogManagerSettings.defaultTenantToken = "";
        AWTLogManagerSettings.semanticContext = new AWTSemanticContext_1.default(true, AWTLogManagerSettings.logManagerContext);
        return AWTLogManagerSettings;
    }();
    exports.default = AWTLogManagerSettings;
}, function(module, exports, __webpack_require__) {
    var rng = __webpack_require__(34);
    var bytesToUuid = __webpack_require__(35);
    function v4(options, buf, offset) {
        var i = buf && offset || 0;
        if (typeof options == "string") {
            buf = options === "binary" ? new Array(16) : null;
            options = null;
        }
        options = options || {};
        var rnds = options.random || (options.rng || rng)();
        rnds[6] = rnds[6] & 15 | 64;
        rnds[8] = rnds[8] & 63 | 128;
        if (buf) {
            for (var ii = 0; ii < 16; ++ii) {
                buf[i + ii] = rnds[ii];
            }
        }
        return buf || bytesToUuid(rnds);
    }
    module.exports = v4;
}, , function(module, exports, __webpack_require__) {
    module.exports = __webpack_require__(36);
}, function(module, exports, __webpack_require__) {
    "use strict";
    Object.defineProperty(exports, "__esModule", {
        value: true
    });
    var Enums_1 = __webpack_require__(0);
    var AWTSerializer_1 = __webpack_require__(15);
    var AWTRetryPolicy_1 = __webpack_require__(29);
    var AWTKillSwitch_1 = __webpack_require__(30);
    var AWTClockSkewManager_1 = __webpack_require__(31);
    var Version = __webpack_require__(16);
    var Utils = __webpack_require__(2);
    var AWTNotificationManager_1 = __webpack_require__(4);
    var AWTTransmissionManagerCore_1 = __webpack_require__(3);
    var MaxConnections = 2;
    var MaxRetries = 1;
    var Method = "POST";
    var AWTHttpManager = function() {
        function AWTHttpManager(_requestQueue, collectorUrl, _queueManager, _httpInterface, clockSkewRefreshDurationInMins) {
            var _this = this;
            this._requestQueue = _requestQueue;
            this._queueManager = _queueManager;
            this._httpInterface = _httpInterface;
            this._urlString = "?qsp=true&content-type=application%2Fbond-compact-binary&client-id=NO_AUTH&sdk-version=" + Version.FullVersionString;
            this._killSwitch = new AWTKillSwitch_1.default();
            this._paused = false;
            this._useBeacons = false;
            this._activeConnections = 0;
            this._clockSkewManager = new AWTClockSkewManager_1.default(clockSkewRefreshDurationInMins);
            if (!Utils.isUint8ArrayAvailable()) {
                this._urlString += "&content-encoding=base64";
            }
            this._urlString = collectorUrl + this._urlString;
            if (!this._httpInterface) {
                this._useBeacons = !Utils.isReactNative();
                this._httpInterface = {
                    sendPOST: function(urlString, data, ontimeout, onerror, onload, sync) {
                        try {
                            if (Utils.useXDomainRequest()) {
                                var xdr = new XDomainRequest();
                                xdr.open(Method, urlString);
                                xdr.onload = function() {
                                    onload(200, null);
                                };
                                xdr.onerror = function() {
                                    onerror(400, null);
                                };
                                xdr.ontimeout = function() {
                                    ontimeout(500, null);
                                };
                                xdr.send(data);
                            } else if (Utils.isReactNative()) {
                                fetch(urlString, {
                                    body: data,
                                    method: Method
                                }).then(function(response) {
                                    var headerMap = {};
                                    if (response.headers) {
                                        response.headers.forEach(function(value, name) {
                                            headerMap[name] = value;
                                        });
                                    }
                                    onload(response.status, headerMap);
                                }).catch(function(error) {
                                    onerror(0, {});
                                });
                            } else {
                                var xhr_1 = new XMLHttpRequest();
                                xhr_1.open(Method, urlString, !sync);
                                xhr_1.onload = function() {
                                    onload(xhr_1.status, _this._convertAllHeadersToMap(xhr_1.getAllResponseHeaders()));
                                };
                                xhr_1.onerror = function() {
                                    onerror(xhr_1.status, _this._convertAllHeadersToMap(xhr_1.getAllResponseHeaders()));
                                };
                                xhr_1.ontimeout = function() {
                                    ontimeout(xhr_1.status, _this._convertAllHeadersToMap(xhr_1.getAllResponseHeaders()));
                                };
                                xhr_1.send(data);
                            }
                        } catch (e) {
                            onerror(400, null);
                        }
                    }
                };
            }
        }
        AWTHttpManager.prototype.hasIdleConnection = function() {
            return this._activeConnections < MaxConnections;
        };
        AWTHttpManager.prototype.sendQueuedRequests = function() {
            while (this.hasIdleConnection() && !this._paused && this._requestQueue.length > 0 && this._clockSkewManager.allowRequestSending()) {
                this._activeConnections++;
                this._sendRequest(this._requestQueue.shift(), 0, false);
            }
            if (this.hasIdleConnection()) {
                AWTTransmissionManagerCore_1.default.scheduleTimer();
            }
        };
        AWTHttpManager.prototype.isCompletelyIdle = function() {
            return this._activeConnections === 0;
        };
        AWTHttpManager.prototype.teardown = function() {
            while (this._requestQueue.length > 0) {
                this._sendRequest(this._requestQueue.shift(), 0, true);
            }
        };
        AWTHttpManager.prototype.pause = function() {
            this._paused = true;
        };
        AWTHttpManager.prototype.resume = function() {
            this._paused = false;
            this.sendQueuedRequests();
        };
        AWTHttpManager.prototype.removeQueuedRequests = function() {
            this._requestQueue.length = 0;
        };
        AWTHttpManager.prototype.sendSynchronousRequest = function(request, token) {
            if (this._paused) {
                request[token][0].priority = Enums_1.AWTEventPriority.High;
            }
            this._activeConnections++;
            this._sendRequest(request, 0, false, true);
        };
        AWTHttpManager.prototype._sendRequest = function(request, retryCount, isTeardown, isSynchronous) {
            var _this = this;
            if (isSynchronous === void 0) {
                isSynchronous = false;
            }
            try {
                if (this._paused) {
                    this._activeConnections--;
                    this._queueManager.addBackRequest(request);
                    return;
                }
                var tokenCount_1 = 0;
                var apikey_1 = "";
                for (var token in request) {
                    if (request.hasOwnProperty(token)) {
                        if (!this._killSwitch.isTenantKilled(token)) {
                            if (apikey_1.length > 0) {
                                apikey_1 += ",";
                            }
                            apikey_1 += token;
                            tokenCount_1++;
                        } else {
                            AWTNotificationManager_1.default.eventsRejected(request[token], Enums_1.AWTEventsRejectedReason.KillSwitch);
                            delete request[token];
                        }
                    }
                }
                if (tokenCount_1 > 0) {
                    var payloadResult = AWTSerializer_1.default.getPayloadBlob(request, tokenCount_1);
                    if (payloadResult.remainingRequest) {
                        this._requestQueue.push(payloadResult.remainingRequest);
                    }
                    var urlString = this._urlString + "&x-apikey=" + apikey_1 + "&client-time-epoch-millis=" + Date.now().toString();
                    if (this._clockSkewManager.shouldAddClockSkewHeaders()) {
                        urlString = urlString + "&time-delta-to-apply-millis=" + this._clockSkewManager.getClockSkewHeaderValue();
                    }
                    var data = void 0;
                    if (!Utils.isUint8ArrayAvailable()) {
                        data = AWTSerializer_1.default.base64Encode(payloadResult.payloadBlob);
                    } else {
                        data = new Uint8Array(payloadResult.payloadBlob);
                    }
                    for (var token in request) {
                        if (request.hasOwnProperty(token)) {
                            for (var i = 0; i < request[token].length; ++i) {
                                request[token][i].sendAttempt > 0 ? request[token][i].sendAttempt++ : request[token][i].sendAttempt = 1;
                            }
                        }
                    }
                    if (this._useBeacons && isTeardown && Utils.isBeaconsSupported()) {
                        if (navigator.sendBeacon(urlString, data)) {
                            return;
                        }
                    }
                    this._httpInterface.sendPOST(urlString, data, function(status, headers) {
                        _this._retryRequestIfNeeded(status, headers, request, tokenCount_1, apikey_1, retryCount, isTeardown, isSynchronous);
                    }, function(status, headers) {
                        _this._retryRequestIfNeeded(status, headers, request, tokenCount_1, apikey_1, retryCount, isTeardown, isSynchronous);
                    }, function(status, headers) {
                        _this._retryRequestIfNeeded(status, headers, request, tokenCount_1, apikey_1, retryCount, isTeardown, isSynchronous);
                    }, isTeardown || isSynchronous);
                } else if (!isTeardown) {
                    this._handleRequestFinished(false, {}, isTeardown, isSynchronous);
                }
            } catch (e) {
                this._handleRequestFinished(false, {}, isTeardown, isSynchronous);
            }
        };
        AWTHttpManager.prototype._retryRequestIfNeeded = function(status, headers, request, tokenCount, apikey, retryCount, isTeardown, isSynchronous) {
            var _this = this;
            var shouldRetry = true;
            if (typeof status !== "undefined") {
                if (headers) {
                    var killedTokens = this._killSwitch.setKillSwitchTenants(headers["kill-tokens"], headers["kill-duration-seconds"]);
                    this._clockSkewManager.setClockSkew(headers["time-delta-millis"]);
                    for (var i = 0; i < killedTokens.length; ++i) {
                        AWTNotificationManager_1.default.eventsRejected(request[killedTokens[i]], Enums_1.AWTEventsRejectedReason.KillSwitch);
                        delete request[killedTokens[i]];
                        tokenCount--;
                    }
                } else {
                    this._clockSkewManager.setClockSkew(null);
                }
                if (status === 200) {
                    this._handleRequestFinished(true, request, isTeardown, isSynchronous);
                    return;
                }
                if (!AWTRetryPolicy_1.default.shouldRetryForStatus(status) || tokenCount <= 0) {
                    shouldRetry = false;
                }
            }
            if (shouldRetry) {
                if (isSynchronous) {
                    this._activeConnections--;
                    request[apikey][0].priority = Enums_1.AWTEventPriority.High;
                    this._queueManager.addBackRequest(request);
                } else if (retryCount < MaxRetries) {
                    for (var token in request) {
                        if (request.hasOwnProperty(token)) {
                            AWTNotificationManager_1.default.eventsRetrying(request[token]);
                        }
                    }
                    setTimeout(function() {
                        return _this._sendRequest(request, retryCount + 1, false);
                    }, AWTRetryPolicy_1.default.getMillisToBackoffForRetry(retryCount));
                } else {
                    this._activeConnections--;
                    AWTTransmissionManagerCore_1.default.backOffTransmission();
                    this._queueManager.addBackRequest(request);
                }
            } else {
                this._handleRequestFinished(false, request, isTeardown, isSynchronous);
            }
        };
        AWTHttpManager.prototype._handleRequestFinished = function(success, request, isTeardown, isSynchronous) {
            if (success) {
                AWTTransmissionManagerCore_1.default.clearBackOff();
            }
            for (var token in request) {
                if (request.hasOwnProperty(token)) {
                    if (success) {
                        AWTNotificationManager_1.default.eventsSent(request[token]);
                    } else {
                        AWTNotificationManager_1.default.eventsDropped(request[token], Enums_1.AWTEventsDroppedReason.NonRetryableStatus);
                    }
                }
            }
            this._activeConnections--;
            if (!isSynchronous && !isTeardown) {
                this.sendQueuedRequests();
            }
        };
        AWTHttpManager.prototype._convertAllHeadersToMap = function(headersString) {
            var headers = {};
            if (headersString) {
                var headersArray = headersString.split("\n");
                for (var i = 0; i < headersArray.length; ++i) {
                    var header = headersArray[i].split(": ");
                    headers[header[0]] = header[1];
                }
            }
            return headers;
        };
        return AWTHttpManager;
    }();
    exports.default = AWTHttpManager;
}, function(module, exports, __webpack_require__) {
    "use strict";
    Object.defineProperty(exports, "__esModule", {
        value: true
    });
    var bond_const_1 = __webpack_require__(24);
    exports._BondDataType = bond_const_1._BondDataType;
    var _Encoding = __webpack_require__(25);
    exports._Encoding = _Encoding;
    var IO = __webpack_require__(28);
    exports.IO = IO;
    var microsoft_bond_primitives_1 = __webpack_require__(8);
    exports.Int64 = microsoft_bond_primitives_1.Int64;
    exports.UInt64 = microsoft_bond_primitives_1.UInt64;
    exports.Number = microsoft_bond_primitives_1.Number;
    var CompactBinaryProtocolWriter = function() {
        function CompactBinaryProtocolWriter(stream) {
            this._stream = stream;
        }
        CompactBinaryProtocolWriter.prototype._WriteBlob = function(blob) {
            this._stream._Write(blob, 0, blob.length);
        };
        CompactBinaryProtocolWriter.prototype._WriteBool = function(value) {
            this._stream._WriteByte(value ? 1 : 0);
        };
        CompactBinaryProtocolWriter.prototype._WriteContainerBegin = function(size, elementType) {
            this._WriteUInt8(elementType);
            this._WriteUInt32(size);
        };
        CompactBinaryProtocolWriter.prototype._WriteMapContainerBegin = function(size, keyType, valueType) {
            this._WriteUInt8(keyType);
            this._WriteUInt8(valueType);
            this._WriteUInt32(size);
        };
        CompactBinaryProtocolWriter.prototype._WriteDouble = function(value) {
            var array = _Encoding._Double_GetBytes(value);
            this._stream._Write(array, 0, array.length);
        };
        CompactBinaryProtocolWriter.prototype._WriteFieldBegin = function(type, id, metadata) {
            if (id <= 5) {
                this._stream._WriteByte(type | id << 5);
            } else if (id <= 255) {
                this._stream._WriteByte(type | 6 << 5);
                this._stream._WriteByte(id);
            } else {
                this._stream._WriteByte(type | 7 << 5);
                this._stream._WriteByte(id);
                this._stream._WriteByte(id >> 8);
            }
        };
        CompactBinaryProtocolWriter.prototype._WriteInt32 = function(value) {
            value = _Encoding._Zigzag_EncodeZigzag32(value);
            this._WriteUInt32(value);
        };
        CompactBinaryProtocolWriter.prototype._WriteInt64 = function(value) {
            this._WriteUInt64(_Encoding._Zigzag_EncodeZigzag64(value));
        };
        CompactBinaryProtocolWriter.prototype._WriteString = function(value) {
            if (value === "") {
                this._WriteUInt32(0);
            } else {
                var array = _Encoding._Utf8_GetBytes(value);
                this._WriteUInt32(array.length);
                this._stream._Write(array, 0, array.length);
            }
        };
        CompactBinaryProtocolWriter.prototype._WriteStructEnd = function(isBase) {
            this._WriteUInt8(isBase ? bond_const_1._BondDataType._BT_STOP_BASE : bond_const_1._BondDataType._BT_STOP);
        };
        CompactBinaryProtocolWriter.prototype._WriteUInt32 = function(value) {
            var array = _Encoding._Varint_GetBytes(microsoft_bond_primitives_1.Number._ToUInt32(value));
            this._stream._Write(array, 0, array.length);
        };
        CompactBinaryProtocolWriter.prototype._WriteUInt64 = function(value) {
            var array = _Encoding._Varint64_GetBytes(value);
            this._stream._Write(array, 0, array.length);
        };
        CompactBinaryProtocolWriter.prototype._WriteUInt8 = function(value) {
            this._stream._WriteByte(microsoft_bond_primitives_1.Number._ToUInt8(value));
        };
        return CompactBinaryProtocolWriter;
    }();
    exports.CompactBinaryProtocolWriter = CompactBinaryProtocolWriter;
}, function(module, exports, __webpack_require__) {
    "use strict";
    Object.defineProperty(exports, "__esModule", {
        value: true
    });
    var _BondDataType;
    (function(_BondDataType) {
        _BondDataType[_BondDataType["_BT_STOP"] = 0] = "_BT_STOP";
        _BondDataType[_BondDataType["_BT_STOP_BASE"] = 1] = "_BT_STOP_BASE";
        _BondDataType[_BondDataType["_BT_BOOL"] = 2] = "_BT_BOOL";
        _BondDataType[_BondDataType["_BT_DOUBLE"] = 8] = "_BT_DOUBLE";
        _BondDataType[_BondDataType["_BT_STRING"] = 9] = "_BT_STRING";
        _BondDataType[_BondDataType["_BT_STRUCT"] = 10] = "_BT_STRUCT";
        _BondDataType[_BondDataType["_BT_LIST"] = 11] = "_BT_LIST";
        _BondDataType[_BondDataType["_BT_MAP"] = 13] = "_BT_MAP";
        _BondDataType[_BondDataType["_BT_INT32"] = 16] = "_BT_INT32";
        _BondDataType[_BondDataType["_BT_INT64"] = 17] = "_BT_INT64";
    })(_BondDataType = exports._BondDataType || (exports._BondDataType = {}));
}, function(module, exports, __webpack_require__) {
    "use strict";
    Object.defineProperty(exports, "__esModule", {
        value: true
    });
    var microsoft_bond_primitives_1 = __webpack_require__(8);
    var microsoft_bond_floatutils_1 = __webpack_require__(26);
    var microsoft_bond_utils_1 = __webpack_require__(27);
    function _Utf8_GetBytes(value) {
        var array = [];
        for (var i = 0; i < value.length; ++i) {
            var char = value.charCodeAt(i);
            if (char < 128) {
                array.push(char);
            } else if (char < 2048) {
                array.push(192 | char >> 6, 128 | char & 63);
            } else if (char < 55296 || char >= 57344) {
                array.push(224 | char >> 12, 128 | char >> 6 & 63, 128 | char & 63);
            } else {
                char = 65536 + ((char & 1023) << 10 | value.charCodeAt(++i) & 1023);
                array.push(240 | char >> 18, 128 | char >> 12 & 63, 128 | char >> 6 & 63, 128 | char & 63);
            }
        }
        return array;
    }
    exports._Utf8_GetBytes = _Utf8_GetBytes;
    function _Base64_GetString(inArray) {
        var lookup = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/";
        var output = [];
        var paddingBytes = inArray.length % 3;
        var toBase64 = function(num) {
            return [ lookup.charAt(num >> 18 & 63), lookup.charAt(num >> 12 & 63), lookup.charAt(num >> 6 & 63), lookup.charAt(num & 63) ].join("");
        };
        for (var i = 0, length_1 = inArray.length - paddingBytes; i < length_1; i += 3) {
            var temp = (inArray[i] << 16) + (inArray[i + 1] << 8) + inArray[i + 2];
            output.push(toBase64(temp));
        }
        switch (paddingBytes) {
          case 1:
            var temp = inArray[inArray.length - 1];
            output.push(lookup.charAt(temp >> 2));
            output.push(lookup.charAt(temp << 4 & 63));
            output.push("==");
            break;

          case 2:
            var temp2 = (inArray[inArray.length - 2] << 8) + inArray[inArray.length - 1];
            output.push(lookup.charAt(temp2 >> 10));
            output.push(lookup.charAt(temp2 >> 4 & 63));
            output.push(lookup.charAt(temp2 << 2 & 63));
            output.push("=");
            break;
        }
        return output.join("");
    }
    exports._Base64_GetString = _Base64_GetString;
    function _Varint_GetBytes(value) {
        var array = [];
        while (value & 4294967168) {
            array.push(value & 127 | 128);
            value >>>= 7;
        }
        array.push(value & 127);
        return array;
    }
    exports._Varint_GetBytes = _Varint_GetBytes;
    function _Varint64_GetBytes(value) {
        var low = value.low;
        var high = value.high;
        var array = [];
        while (high || 4294967168 & low) {
            array.push(low & 127 | 128);
            low = (high & 127) << 25 | low >>> 7;
            high >>>= 7;
        }
        array.push(low & 127);
        return array;
    }
    exports._Varint64_GetBytes = _Varint64_GetBytes;
    function _Double_GetBytes(value) {
        if (microsoft_bond_utils_1.BrowserChecker._IsDataViewSupport()) {
            var view = new DataView(new ArrayBuffer(8));
            view.setFloat64(0, value, true);
            var array = [];
            for (var i = 0; i < 8; ++i) {
                array.push(view.getUint8(i));
            }
            return array;
        } else {
            return microsoft_bond_floatutils_1.FloatUtils._ConvertNumberToArray(value, true);
        }
    }
    exports._Double_GetBytes = _Double_GetBytes;
    function _Zigzag_EncodeZigzag32(value) {
        value = microsoft_bond_primitives_1.Number._ToInt32(value);
        return value << 1 ^ value >> 4 * 8 - 1;
    }
    exports._Zigzag_EncodeZigzag32 = _Zigzag_EncodeZigzag32;
    function _Zigzag_EncodeZigzag64(value) {
        var low = value.low;
        var high = value.high;
        var tmpH = high << 1 | low >>> 31;
        var tmpL = low << 1;
        if (high & 2147483648) {
            tmpH = ~tmpH;
            tmpL = ~tmpL;
        }
        var res = new microsoft_bond_primitives_1.UInt64("0");
        res.low = tmpL;
        res.high = tmpH;
        return res;
    }
    exports._Zigzag_EncodeZigzag64 = _Zigzag_EncodeZigzag64;
}, function(module, exports, __webpack_require__) {
    "use strict";
    Object.defineProperty(exports, "__esModule", {
        value: true
    });
    var FloatUtils = function() {
        function FloatUtils() {}
        FloatUtils._ConvertNumberToArray = function(num, isDouble) {
            if (!num) {
                return isDouble ? this._doubleZero : this._floatZero;
            }
            var exponentBits = isDouble ? 11 : 8;
            var precisionBits = isDouble ? 52 : 23;
            var bias = (1 << exponentBits - 1) - 1;
            var minExponent = 1 - bias;
            var maxExponent = bias;
            var sign = num < 0 ? 1 : 0;
            num = Math.abs(num);
            var intPart = Math.floor(num);
            var floatPart = num - intPart;
            var len = 2 * (bias + 2) + precisionBits;
            var buffer = new Array(len);
            var i = 0;
            while (i < len) {
                buffer[i++] = 0;
            }
            i = bias + 2;
            while (i && intPart) {
                buffer[--i] = intPart % 2;
                intPart = Math.floor(intPart / 2);
            }
            i = bias + 1;
            while (i < len - 1 && floatPart > 0) {
                floatPart *= 2;
                if (floatPart >= 1) {
                    buffer[++i] = 1;
                    --floatPart;
                } else {
                    buffer[++i] = 0;
                }
            }
            var firstBit = 0;
            while (firstBit < len && !buffer[firstBit]) {
                firstBit++;
            }
            var exponent = bias + 1 - firstBit;
            var lastBit = firstBit + precisionBits;
            if (buffer[lastBit + 1]) {
                for (i = lastBit; i > firstBit; --i) {
                    buffer[i] = 1 - buffer[i];
                    if (buffer) {
                        break;
                    }
                }
                if (i === firstBit) {
                    ++exponent;
                }
            }
            if (exponent > maxExponent || intPart) {
                if (sign) {
                    return isDouble ? this._doubleNegInifinity : this._floatNegInifinity;
                } else {
                    return isDouble ? this._doubleInifinity : this._floatInifinity;
                }
            } else if (exponent < minExponent) {
                return isDouble ? this._doubleZero : this._floatZero;
            }
            if (isDouble) {
                var high = 0;
                for (i = 0; i < 20; ++i) {
                    high = high << 1 | buffer[++firstBit];
                }
                var low = 0;
                for (;i < 52; ++i) {
                    low = low << 1 | buffer[++firstBit];
                }
                high |= exponent + bias << 20;
                high = sign << 31 | high & 2147483647;
                var resArray = [ low & 255, low >> 8 & 255, low >> 16 & 255, low >>> 24, high & 255, high >> 8 & 255, high >> 16 & 255, high >>> 24 ];
                return resArray;
            } else {
                var result = 0;
                for (i = 0; i < 23; ++i) {
                    result = result << 1 | buffer[++firstBit];
                }
                result |= exponent + bias << 23;
                result = sign << 31 | result & 2147483647;
                var resArray = [ result & 255, result >> 8 & 255, result >> 16 & 255, result >>> 24 ];
                return resArray;
            }
        };
        FloatUtils._floatZero = [ 0, 0, 0, 0 ];
        FloatUtils._doubleZero = [ 0, 0, 0, 0, 0, 0, 0, 0 ];
        FloatUtils._floatInifinity = [ 0, 0, 128, 127 ];
        FloatUtils._floatNegInifinity = [ 0, 0, 128, 255 ];
        FloatUtils._doubleInifinity = [ 0, 0, 0, 0, 0, 0, 240, 127 ];
        FloatUtils._doubleNegInifinity = [ 0, 0, 0, 0, 0, 0, 240, 255 ];
        return FloatUtils;
    }();
    exports.FloatUtils = FloatUtils;
}, function(module, exports, __webpack_require__) {
    "use strict";
    Object.defineProperty(exports, "__esModule", {
        value: true
    });
    var BrowserChecker = function() {
        function BrowserChecker() {}
        BrowserChecker._IsDataViewSupport = function() {
            return typeof ArrayBuffer !== "undefined" && typeof DataView !== "undefined";
        };
        return BrowserChecker;
    }();
    exports.BrowserChecker = BrowserChecker;
}, function(module, exports, __webpack_require__) {
    "use strict";
    Object.defineProperty(exports, "__esModule", {
        value: true
    });
    var microsoft_bond_primitives_1 = __webpack_require__(8);
    var MemoryStream = function() {
        function MemoryStream() {
            this._buffer = [];
        }
        MemoryStream.prototype._WriteByte = function(byte) {
            this._buffer.push(microsoft_bond_primitives_1.Number._ToByte(byte));
        };
        MemoryStream.prototype._Write = function(buffer, offset, count) {
            while (count--) {
                this._WriteByte(buffer[offset++]);
            }
        };
        MemoryStream.prototype._GetBuffer = function() {
            return this._buffer;
        };
        return MemoryStream;
    }();
    exports.MemoryStream = MemoryStream;
}, function(module, exports, __webpack_require__) {
    "use strict";
    Object.defineProperty(exports, "__esModule", {
        value: true
    });
    var RandomizationLowerThreshold = .8;
    var RandomizationUpperThreshold = 1.2;
    var BaseBackoff = 3e3;
    var MaxBackoff = 12e4;
    var AWTRetryPolicy = function() {
        function AWTRetryPolicy() {}
        AWTRetryPolicy.shouldRetryForStatus = function(httpStatusCode) {
            return !(httpStatusCode >= 300 && httpStatusCode < 500 && httpStatusCode !== 408 || httpStatusCode === 501 || httpStatusCode === 505);
        };
        AWTRetryPolicy.getMillisToBackoffForRetry = function(retriesSoFar) {
            var waitDuration = 0;
            var minBackoff = BaseBackoff * RandomizationLowerThreshold;
            var maxBackoff = BaseBackoff * RandomizationUpperThreshold;
            var randomBackoff = Math.floor(Math.random() * (maxBackoff - minBackoff)) + minBackoff;
            waitDuration = Math.pow(4, retriesSoFar) * randomBackoff;
            return Math.min(waitDuration, MaxBackoff);
        };
        return AWTRetryPolicy;
    }();
    exports.default = AWTRetryPolicy;
}, function(module, exports, __webpack_require__) {
    "use strict";
    Object.defineProperty(exports, "__esModule", {
        value: true
    });
    var SecToMsMultiplier = 1e3;
    var AWTKillSwitch = function() {
        function AWTKillSwitch() {
            this._killedTokenDictionary = {};
        }
        AWTKillSwitch.prototype.setKillSwitchTenants = function(killTokens, killDuration) {
            if (killTokens && killDuration) {
                try {
                    var killedTokens = killTokens.split(",");
                    if (killDuration === "this-request-only") {
                        return killedTokens;
                    }
                    var durationMs = parseInt(killDuration, 10) * SecToMsMultiplier;
                    for (var i = 0; i < killedTokens.length; ++i) {
                        this._killedTokenDictionary[killedTokens[i]] = Date.now() + durationMs;
                    }
                } catch (ex) {
                    return [];
                }
            }
            return [];
        };
        AWTKillSwitch.prototype.isTenantKilled = function(tenantToken) {
            if (this._killedTokenDictionary[tenantToken] !== undefined && this._killedTokenDictionary[tenantToken] > Date.now()) {
                return true;
            }
            delete this._killedTokenDictionary[tenantToken];
            return false;
        };
        return AWTKillSwitch;
    }();
    exports.default = AWTKillSwitch;
}, function(module, exports, __webpack_require__) {
    "use strict";
    Object.defineProperty(exports, "__esModule", {
        value: true
    });
    var AWTClockSkewManager = function() {
        function AWTClockSkewManager(clockSkewRefreshDurationInMins) {
            this.clockSkewRefreshDurationInMins = clockSkewRefreshDurationInMins;
            this._reset();
        }
        AWTClockSkewManager.prototype.allowRequestSending = function() {
            if (this._isFirstRequest && !this._clockSkewSet) {
                this._isFirstRequest = false;
                this._allowRequestSending = false;
                return true;
            }
            return this._allowRequestSending;
        };
        AWTClockSkewManager.prototype.shouldAddClockSkewHeaders = function() {
            return this._shouldAddClockSkewHeaders;
        };
        AWTClockSkewManager.prototype.getClockSkewHeaderValue = function() {
            return this._clockSkewHeaderValue;
        };
        AWTClockSkewManager.prototype.setClockSkew = function(timeDeltaInMillis) {
            if (!this._clockSkewSet) {
                if (timeDeltaInMillis) {
                    this._clockSkewHeaderValue = timeDeltaInMillis;
                } else {
                    this._shouldAddClockSkewHeaders = false;
                }
                this._clockSkewSet = true;
                this._allowRequestSending = true;
            }
        };
        AWTClockSkewManager.prototype._reset = function() {
            var _this = this;
            this._isFirstRequest = true;
            this._clockSkewSet = false;
            this._allowRequestSending = true;
            this._shouldAddClockSkewHeaders = true;
            this._clockSkewHeaderValue = "use-collector-delta";
            if (this.clockSkewRefreshDurationInMins > 0) {
                setTimeout(function() {
                    return _this._reset();
                }, this.clockSkewRefreshDurationInMins * 6e4);
            }
        };
        return AWTClockSkewManager;
    }();
    exports.default = AWTClockSkewManager;
}, function(module, exports, __webpack_require__) {
    "use strict";
    Object.defineProperty(exports, "__esModule", {
        value: true
    });
    var Enums_1 = __webpack_require__(0);
    var AWTRecordBatcher = function() {
        function AWTRecordBatcher(_outboundQueue, _maxNumberOfEvents) {
            this._outboundQueue = _outboundQueue;
            this._maxNumberOfEvents = _maxNumberOfEvents;
            this._currentBatch = {};
            this._currentNumEventsInBatch = 0;
        }
        AWTRecordBatcher.prototype.addEventToBatch = function(event) {
            if (event.priority === Enums_1.AWTEventPriority.Immediate_sync) {
                var immediateBatch = {};
                immediateBatch[event.apiKey] = [ event ];
                return immediateBatch;
            } else {
                if (this._currentNumEventsInBatch >= this._maxNumberOfEvents) {
                    this.flushBatch();
                }
                if (this._currentBatch[event.apiKey] === undefined) {
                    this._currentBatch[event.apiKey] = [];
                }
                this._currentBatch[event.apiKey].push(event);
                this._currentNumEventsInBatch++;
            }
            return null;
        };
        AWTRecordBatcher.prototype.flushBatch = function() {
            if (this._currentNumEventsInBatch > 0) {
                this._outboundQueue.push(this._currentBatch);
                this._currentBatch = {};
                this._currentNumEventsInBatch = 0;
            }
        };
        AWTRecordBatcher.prototype.hasBatch = function() {
            return this._currentNumEventsInBatch > 0;
        };
        return AWTRecordBatcher;
    }();
    exports.default = AWTRecordBatcher;
}, function(module, exports, __webpack_require__) {
    "use strict";
    Object.defineProperty(exports, "__esModule", {
        value: true
    });
    var AWTTransmissionManagerCore_1 = __webpack_require__(3);
    var AWTTransmissionManager = function() {
        function AWTTransmissionManager() {}
        AWTTransmissionManager.setEventsHandler = function(eventsHandler) {
            AWTTransmissionManagerCore_1.default.setEventsHandler(eventsHandler);
        };
        AWTTransmissionManager.getEventsHandler = function() {
            return AWTTransmissionManagerCore_1.default.getEventsHandler();
        };
        AWTTransmissionManager.scheduleTimer = function() {
            AWTTransmissionManagerCore_1.default.scheduleTimer();
        };
        return AWTTransmissionManager;
    }();
    exports.default = AWTTransmissionManager;
}, function(module, exports) {
    var getRandomValues = typeof crypto != "undefined" && crypto.getRandomValues && crypto.getRandomValues.bind(crypto) || typeof msCrypto != "undefined" && typeof window.msCrypto.getRandomValues == "function" && msCrypto.getRandomValues.bind(msCrypto);
    if (getRandomValues) {
        var rnds8 = new Uint8Array(16);
        module.exports = function whatwgRNG() {
            getRandomValues(rnds8);
            return rnds8;
        };
    } else {
        var rnds = new Array(16);
        module.exports = function mathRNG() {
            for (var i = 0, r; i < 16; i++) {
                if ((i & 3) === 0) r = Math.random() * 4294967296;
                rnds[i] = r >>> ((i & 3) << 3) & 255;
            }
            return rnds;
        };
    }
}, function(module, exports) {
    var byteToHex = [];
    for (var i = 0; i < 256; ++i) {
        byteToHex[i] = (i + 256).toString(16).substr(1);
    }
    function bytesToUuid(buf, offset) {
        var i = offset || 0;
        var bth = byteToHex;
        return [ bth[buf[i++]], bth[buf[i++]], bth[buf[i++]], bth[buf[i++]], "-", bth[buf[i++]], bth[buf[i++]], "-", bth[buf[i++]], bth[buf[i++]], "-", bth[buf[i++]], bth[buf[i++]], "-", bth[buf[i++]], bth[buf[i++]], bth[buf[i++]], bth[buf[i++]], bth[buf[i++]], bth[buf[i++]] ].join("");
    }
    module.exports = bytesToUuid;
}, function(module, __webpack_exports__, __webpack_require__) {
    "use strict";
    __webpack_require__.r(__webpack_exports__);
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
    var DataFieldType;
    (function(DataFieldType) {
        DataFieldType[DataFieldType["String"] = 0] = "String";
        DataFieldType[DataFieldType["Boolean"] = 1] = "Boolean";
        DataFieldType[DataFieldType["Int64"] = 2] = "Int64";
        DataFieldType[DataFieldType["Double"] = 3] = "Double";
    })(DataFieldType || (DataFieldType = {}));
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
    function isWacAgave() {
        if (typeof Office !== "undefined" && typeof Office.context !== "undefined" && typeof Office.context.platform !== "undefined") {
            return Office.context.platform === Office.PlatformType.OfficeOnline;
        }
        return typeof OfficeExt !== "undefined" && typeof OfficeExt.HostName !== "undefined" && typeof OfficeExt.HostName.Host !== "undefined" && typeof OfficeExt.HostName.Host.getInstance === "function" && typeof OfficeExt.HostName.Host.getInstance().getPlatform === "function" && OfficeExt.HostName.Host.getInstance().getPlatform() === Office.PlatformType.OfficeOnline;
    }
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
    var DEFAULT_MINIMUM_MILLISECONDS_BETWEEN_CALLS = 1e3;
    var _minimumMillisecondsBeforeFirstCall = DEFAULT_MINIMUM_MILLISECONDS_BETWEEN_CALLS;
    var _minimumMillisecondsBetweenCalls = DEFAULT_MINIMUM_MILLISECONDS_BETWEEN_CALLS;
    var RichApiTelemetryQueue_RichApiTelemetryQueue = function() {
        function RichApiTelemetryQueue(onSendFirstEvent) {
            this._requestIsPending = false;
            this._items = [];
            this._sentFirstEvent = false;
            this._sentFirstEvent = false;
            this._onSendFirstEvent = onSendFirstEvent;
        }
        RichApiTelemetryQueue.prototype.add = function(item) {
            this._items.push(item);
            if (this._requestIsPending) {
                return;
            }
            this.processWorkBacklog();
        };
        RichApiTelemetryQueue.prototype.processWorkBacklog = function() {
            var _this = this;
            this._requestIsPending = true;
            var currentWork = this._items;
            this._items = [];
            this.pauseIfNecessary().then(function() {
                _this.processTelemetryEvents(currentWork);
                _this.waitAndProcessMore();
            }).catch(function(e) {
                logNotification(LogLevel.Error, Category.Sink, function() {
                    return JSON.stringify(e);
                });
                _this.waitAndProcessMore();
            });
        };
        RichApiTelemetryQueue.prototype.waitAndProcessMore = function() {
            var _this = this;
            pause(_minimumMillisecondsBetweenCalls).then(function() {
                if (_this._items.length > 0) {
                    setTimeout(function() {
                        return _this.processWorkBacklog();
                    }, 0);
                }
                _this._requestIsPending = false;
            });
        };
        RichApiTelemetryQueue.prototype.processTelemetryEvents = function(telemetryEvents) {
            var _this = this;
            var ctx = new OfficeCore.RequestContext();
            telemetryEvents.forEach(function(telemetryEvent) {
                if (!telemetryEvent.telemetryProperties) {
                    return;
                }
                var dataFields = [];
                _this.addDataFields(dataFields, telemetryEvent.dataFields);
                var contractName = !!telemetryEvent.eventContract ? telemetryEvent.eventContract.name : "";
                if (!!telemetryEvent.eventContract) {
                    _this.addDataFields(dataFields, telemetryEvent.eventContract.dataFields);
                }
                ctx.telemetry.sendTelemetryEvent(telemetryEvent.telemetryProperties, telemetryEvent.eventName, contractName, getEffectiveEventFlags(telemetryEvent), dataFields);
            });
            ctx.sync().then(function() {
                if (!_this._sentFirstEvent) {
                    _this._sentFirstEvent = true;
                    if (_this._onSendFirstEvent) {
                        _this._onSendFirstEvent(true);
                    }
                }
            }).catch(function(e) {
                if (!_this._sentFirstEvent) {
                    _this._sentFirstEvent = true;
                    if (_this._onSendFirstEvent) {
                        _this._onSendFirstEvent(false);
                    }
                } else {
                    logNotification(LogLevel.Error, Category.Sink, function() {
                        return JSON.stringify(e);
                    });
                }
            });
        };
        RichApiTelemetryQueue.prototype.addDataFields = function(richApiDataFields, dataFields) {
            if (dataFields) {
                dataFields.forEach(function(dataField) {
                    richApiDataFields.push({
                        name: dataField.name,
                        value: dataField.value,
                        classification: dataField.classification ? dataField.classification : DataClassification.SystemMetadata,
                        type: dataField.dataType
                    });
                });
            }
        };
        RichApiTelemetryQueue.prototype.pauseIfNecessary = function() {
            if (!this._sentFirstEvent) {
                return pause(_minimumMillisecondsBeforeFirstCall);
            }
            return OfficeExtension.Promise.resolve(undefined);
        };
        return RichApiTelemetryQueue;
    }();
    function pause(ms) {
        return new OfficeExtension.Promise(function(resolve) {
            return setTimeout(resolve, ms);
        });
    }
    var RichApiSink_RichApiSink = function() {
        function RichApiSink(queue) {
            this._queue = queue ? queue : new RichApiTelemetryQueue_RichApiTelemetryQueue();
        }
        RichApiSink.prototype.sendTelemetryEvent = function(event, timestamp) {
            try {
                this._queue.add(event);
            } catch (error) {
                logNotification(LogLevel.Error, Category.Sink, function() {
                    timestamp;
                    return "RichApiSink caught an error : " + JSON.stringify(error);
                });
            }
        };
        return RichApiSink;
    }();
    var _queriedForIsSupported = false;
    var _queue;
    var _richApiSink;
    var IsSupportedState;
    (function(IsSupportedState) {
        IsSupportedState[IsSupportedState["Unsupported"] = 0] = "Unsupported";
        IsSupportedState[IsSupportedState["Supported"] = 1] = "Supported";
        IsSupportedState[IsSupportedState["NotDetermined"] = 2] = "NotDetermined";
    })(IsSupportedState || (IsSupportedState = {}));
    function getRichApiSink(forceNew, onGetRichApiSink) {
        if (_queriedForIsSupported && !forceNew) {
            return onGetRichApiSink(_richApiSink);
        }
        _queue = undefined;
        _richApiSink = undefined;
        var isSupported = isSupportedSync();
        if (isSupported === IsSupportedState.NotDetermined) {
            return isSupportedAsync(onGetRichApiSink);
        }
        _queriedForIsSupported = true;
        if (isSupported) {
            _richApiSink = new RichApiSink_RichApiSink();
        }
        onGetRichApiSink(_richApiSink);
    }
    function isSupportedSync() {
        if (isWacAgave()) {
            logNotification(LogLevel.Info, Category.Sink, function() {
                return "RichApi telemetry is not supported on Office Online";
            });
            return IsSupportedState.Unsupported;
        }
        if (typeof OfficeCore === "undefined") {
            logNotification(LogLevel.Info, Category.Sink, function() {
                return "Can't get OfficeCore";
            });
            return IsSupportedState.Unsupported;
        }
        if (isTelemetryApiSetSupported()) {
            return IsSupportedState.Supported;
        }
        return IsSupportedState.NotDetermined;
    }
    function isTelemetryApiSetSupported() {
        return Office.context.requirements.isSetSupported("Telemetry", 1.1);
    }
    function isSupportedAsync(onGetRichApiSink) {
        var testEvent = {
            telemetryProperties: {
                nexusTenantToken: 1723,
                ariaTenantToken: "f998cc5ba4d448d6a1e8e913ff18be94-dd122e0a-fcf8-4dc5-9dbb-6afac5325183-7405"
            },
            eventName: "Office.Telemetry.RichApi.TestForSupport",
            eventFlags: {
                dataCategories: DataCategories.ProductServiceUsage,
                diagnosticLevel: DiagnosticLevel.FullEvent
            }
        };
        function onSendEvent(succeeded) {
            if (succeeded) {
                _richApiSink = new RichApiSink_RichApiSink(_queue);
            } else {
                logNotification(LogLevel.Info, Category.Sink, function() {
                    return "RichAPI SendTelemetryEvent is not supported on the host";
                });
            }
            onGetRichApiSink(_richApiSink);
        }
        _queue = new RichApiTelemetryQueue_RichApiTelemetryQueue(onSendEvent);
        _queue.add(testEvent);
    }
    var SdxWacSink_SdxWacSink = function() {
        function SdxWacSink() {}
        SdxWacSink.isSupported = function() {
            return isWacAgave() && typeof OSF === "object" && typeof OSF.getClientEndPoint === "function" && typeof OSF._OfficeAppFactory === "object" && typeof OSF._OfficeAppFactory.getId === "function" && typeof OSF.AgaveHostAction === "object" && typeof OSF.AgaveHostAction.SendTelemetryEvent === "number";
        };
        SdxWacSink.prototype.sendTelemetryEvent = function(event, timestamp) {
            try {
                if (event.dataFields && event.dataFields.filter(function(dataField) {
                    return dataField.classification && dataField.classification !== DataClassification.SystemMetadata;
                }).length > 0) {
                    return;
                }
                var id = OSF._OfficeAppFactory.getId();
                var SendTelemetryEventId = OSF.AgaveHostAction.SendTelemetryEvent;
                OSF.getClientEndPoint().invoke("ContextActivationManager_notifyHost", null, [ id, SendTelemetryEventId, event ]);
            } catch (error) {
                logNotification(LogLevel.Error, Category.Sink, function() {
                    timestamp;
                    return "AgaveWacSink caught an error : " + JSON.stringify(error);
                });
            }
        };
        return SdxWacSink;
    }();
    var AriaSDK = __webpack_require__(5);
    var AWTTransmissionManagerCore = __webpack_require__(3);
    var AWTTransmissionManagerCore_default = __webpack_require__.n(AWTTransmissionManagerCore);
    var AWTQueueManager = __webpack_require__(11);
    var AWTQueueManager_default = __webpack_require__.n(AWTQueueManager);
    var Enums = __webpack_require__(0);
    var EVENT_NAME_DOT_REPLACE_REGEX = /\./g;
    var SEPARATOR_TOKEN = ".";
    var DATA_TOKEN = "Data";
    var CONTRACT_TOKEN = "zC";
    var eventSequence = 0;
    function getAriaEvent(telemetryEvent, additionalDataFields, timestamp) {
        var ariaEvent = {
            name: getAriaEventName(telemetryEvent.eventName),
            properties: {}
        };
        if (!telemetryEvent.telemetryProperties || !telemetryEvent.telemetryProperties.ariaTenantToken) {
            throw new Error("Unable to find ariaTenantToken for namespace.");
        }
        ariaEvent.properties["Event.Sequence"] = {
            value: ++eventSequence,
            type: Enums["AWTPropertyType"].Int64
        };
        ariaEvent.properties["Event.Name"] = telemetryEvent.eventName;
        ariaEvent.properties["Event.Source"] = "OTelJS";
        var timestampLocal;
        if (timestamp) {
            timestampLocal = new Date(timestamp);
        } else {
            timestampLocal = new Date();
        }
        ariaEvent.properties["Event.Time"] = {
            value: timestampLocal,
            type: Enums["AWTPropertyType"].Date
        };
        if (telemetryEvent.eventContract) {
            ariaEvent.properties["Event.Contract"] = telemetryEvent.eventContract.name;
            addDataFields(ariaEvent, telemetryEvent.eventContract.dataFields, false);
        }
        addDataFields(ariaEvent, additionalDataFields, false);
        addDataFields(ariaEvent, telemetryEvent.dataFields, true);
        return ariaEvent;
    }
    function addDataFields(ariaEvent, fields, prependDataToken) {
        if (fields) {
            fields.forEach(function(field) {
                if (field.classification && field.classification !== DataClassification.SystemMetadata) {
                    return;
                }
                var _a = [ "", "", field.name ], metadataPrefix = _a[0], dataToken = _a[1], fieldName = _a[2];
                var firstSeparator = field.name.indexOf(SEPARATOR_TOKEN);
                if (firstSeparator > 0 && isMetadataPrefix(field.name.substr(0, firstSeparator))) {
                    metadataPrefix = field.name.substring(0, firstSeparator + 1);
                    fieldName = field.name.substring(firstSeparator + 1);
                }
                if (prependDataToken) {
                    dataToken = DATA_TOKEN + SEPARATOR_TOKEN;
                }
                var ariaFieldName = metadataPrefix + dataToken + fieldName;
                ariaEvent.properties[ariaFieldName] = {
                    value: field.value,
                    type: mapDataFieldTypeToAWTPropertyType(field.dataType)
                };
            });
        }
    }
    function mapDataFieldTypeToAWTPropertyType(otelType) {
        switch (otelType) {
          case DataFieldType.String:
            return Enums["AWTPropertyType"].String;

          case DataFieldType.Boolean:
            return Enums["AWTPropertyType"].Boolean;

          case DataFieldType.Int64:
            return Enums["AWTPropertyType"].Int64;

          case DataFieldType.Double:
            return Enums["AWTPropertyType"].Double;

          default:
            var _exhaustiveCheck = otelType;
            throw new Error(_exhaustiveCheck);
        }
    }
    function getAriaEventName(eventName) {
        return eventName.toLowerCase().replace(EVENT_NAME_DOT_REPLACE_REGEX, "_");
    }
    function isMetadataPrefix(prefix) {
        return prefix === CONTRACT_TOKEN;
    }
    var ARIA_INIT_TOKEN = "cd836626611c4caaa8fc5b2e728ee81d-3b6d6c45-6377-4bf5-9792-dbf8e1881088-7521";
    var otelEventsProcessed = 0;
    var ariaEventsSent = 0;
    var ariaEventsDropped = 0;
    var ariaEventsRejected = 0;
    var ariaEventsRetrying = 0;
    var awtInitialized = false;
    function sendEvent(telemetryEvent, additionalDataFields, timestamp) {
        otelEventsProcessed++;
        initialize();
        var ariaEvent;
        if (!telemetryEvent.telemetryProperties || !telemetryEvent.telemetryProperties.ariaTenantToken) {
            throw new Error("Unable to find ariaTenantToken for namespace.");
        }
        ariaEvent = getAriaEvent(telemetryEvent, additionalDataFields, timestamp);
        var logger = AriaSDK["AWTLogManager"].getLogger(telemetryEvent.telemetryProperties.ariaTenantToken);
        logger.logEvent(ariaEvent);
    }
    function initialize(configuration) {
        if (!awtInitialized) {
            hookUpAriaNotifications();
            AriaSDK["AWTLogManager"].initialize(ARIA_INIT_TOKEN, configuration);
            awtInitialized = true;
        }
    }
    function hookUpAriaNotifications() {
        AriaSDK["AWTLogManager"].addNotificationListener({
            eventsSent: function(events) {
                logNotification(LogLevel.Info, Category.Transport, function() {
                    return "Successfully sent " + events.length + " event(s)";
                });
                logNotification(LogLevel.Verbose, Category.Transport, function() {
                    return "Sent event(s) details : " + JSON.stringify(events, null, 2);
                });
                ariaEventsSent += events.length;
            },
            eventsDropped: function(events, reason) {
                logNotification(LogLevel.Error, Category.Transport, function() {
                    return "Dropped " + events.length + " event(s) because " + reason;
                });
                logNotification(LogLevel.Verbose, Category.Transport, function() {
                    return "Dropped event(s) details : " + JSON.stringify(events, null, 2);
                });
                ariaEventsDropped += events.length;
            },
            eventsRejected: function(events, reason) {
                logNotification(LogLevel.Error, Category.Transport, function() {
                    return "Rejected " + events.length + " event(s) because " + reason;
                });
                logNotification(LogLevel.Verbose, Category.Transport, function() {
                    return "Rejected event(s) details : " + JSON.stringify(events, null, 2);
                });
                ariaEventsRejected += events.length;
            },
            eventsRetrying: function(events) {
                logNotification(LogLevel.Warning, Category.Transport, function() {
                    return "Retrying " + events.length + " event(s)";
                });
                logNotification(LogLevel.Verbose, Category.Transport, function() {
                    return "Retrying event(s) details : " + JSON.stringify(events, null, 2);
                });
                ariaEventsRetrying += events.length;
            }
        });
    }
    function disableBeaconsApiAvailabilityForAriaSdk() {
        var ariaEventHander = AWTTransmissionManagerCore_default.a.getEventsHandler();
        if (ariaEventHander instanceof AWTQueueManager_default.a) {
            var ariaHttpManager = ariaEventHander._httpManager;
            if (ariaHttpManager && ariaHttpManager.hasOwnProperty("_useBeacons")) {
                ariaHttpManager._useBeacons = false;
                return true;
            }
        }
        return false;
    }
    function shutdown() {
        AriaSDK["AWTLogManager"].flushAndTeardown();
    }
    var FullEventProcessor_FullEventProcessor = function() {
        function FullEventProcessor() {
            this._fullEventsEnabled = false;
        }
        FullEventProcessor.prototype.processEvent = function(event) {
            return this._fullEventsEnabled || !!event.eventFlags && (event.eventFlags.diagnosticLevel === DiagnosticLevel.BasicEvent || event.eventFlags.diagnosticLevel === DiagnosticLevel.NecessaryServiceDataEvent || event.eventFlags.diagnosticLevel === DiagnosticLevel.AlwaysOnNecessaryServiceDataEvent);
        };
        FullEventProcessor.prototype.setFullEventsEnabled = function(enabled) {
            this._fullEventsEnabled = enabled;
        };
        return FullEventProcessor;
    }();
    var AriaSink_AriaSink = function() {
        function AriaSink(additionalDataFields, cacheMemorySizeLimitInNumberOfEvents) {
            if (additionalDataFields === void 0) {
                additionalDataFields = [];
            }
            this._preprocessors = [];
            this._additionalDataFields = additionalDataFields;
            this._fullEventProcessor = new FullEventProcessor_FullEventProcessor();
            this.addPreprocessor(this._fullEventProcessor);
            initialize({
                cacheMemorySizeLimitInNumberOfEvents: cacheMemorySizeLimitInNumberOfEvents
            });
        }
        AriaSink.prototype.sendTelemetryEvent = function(event, timestamp) {
            try {
                for (var i = 0; i < this._preprocessors.length; i++) {
                    if (!this._preprocessors[i].processEvent(event)) {
                        return;
                    }
                }
                sendEvent(event, this._additionalDataFields, timestamp);
            } catch (error) {
                logNotification(LogLevel.Error, Category.Sink, function() {
                    var errorMessage;
                    if (error instanceof Error) {
                        errorMessage = error.message;
                    } else {
                        errorMessage = JSON.stringify(error);
                    }
                    return "AriaSink caught an error : " + errorMessage;
                });
            }
        };
        AriaSink.prototype.addPreprocessor = function(preprocessor) {
            this._preprocessors.push(preprocessor);
        };
        AriaSink.prototype.setFullEventsEnabled = function(enabled) {
            this._fullEventProcessor.setFullEventsEnabled(enabled);
        };
        AriaSink.prototype.disableBeaconsApiAvailabilityForAriaSdk = function() {
            return disableBeaconsApiAvailabilityForAriaSdk();
        };
        AriaSink.prototype.shutdown = function() {
            shutdown();
        };
        return AriaSink;
    }();
    var v4 = __webpack_require__(19);
    var Utils_Utils;
    (function(Utils) {
        function newGuid() {
            return v4();
        }
        Utils.newGuid = newGuid;
    })(Utils_Utils || (Utils_Utils = {}));
    var OutlookSink_OutlookSink = function() {
        function OutlookSink() {}
        OutlookSink.isSupported = function() {
            return !!Office && Office.context.requirements.isSetSupported("OutlookTelemetry", 1);
        };
        OutlookSink.prototype.sendTelemetryEvent = function(event) {
            if (event.eventName.match(/^Office\.Extensibility\.OfficeJs\.[a-zA-Z]*$/)) {
                Office.context.mailbox.logTelemetry(JSON.stringify(event));
            } else {
                logNotification(LogLevel.Warning, Category.Sink, function() {
                    return "Outlook only accepts OfficeJS telemetry events";
                });
            }
        };
        return OutlookSink;
    }();
    var __assign = undefined && undefined.__assign || Object.assign || function(t) {
        for (var s, i = 1, n = arguments.length; i < n; i++) {
            s = arguments[i];
            for (var p in s) if (Object.prototype.hasOwnProperty.call(s, p)) t[p] = s[p];
        }
        return t;
    };
    var CLIENTID_LOCALSTORAGE_NAME = "OTelJS.ClientId";
    var clientId = null;
    var sessionId = null;
    var AgaveSink_AgaveSink = function() {
        function AgaveSink(ariaAdditionalContext, ariaSendEventEnabled) {
            this._isUsable = true;
            this._awaitingInitialization = false;
            this._eventQueue = [];
            this.defaultAriaContext = {
                "App.Name": "TBD",
                "App.Platform": "TBD",
                "App.Version": "TBD",
                "Device.OsBuild": "TBD",
                "Device.OsVersion": "TBD",
                "Host.Id": "",
                "Host.Version": "",
                "Client.Id": "",
                "Session.Id": "",
                "Release.Audience": "TBD",
                "Release.AudienceGroup": "TBD",
                "Release.Channel": "TBD",
                "Release.Fork": "TBD"
            };
            this._ariaAdditionalContext = ariaAdditionalContext;
            this._ariaSendEventEnabled = ariaSendEventEnabled;
            this.initialize();
        }
        AgaveSink.createInstance = function(ariaAdditionalContext, ariaSendEventEnabled) {
            if (ariaAdditionalContext === void 0) {
                ariaAdditionalContext = {};
            }
            if (ariaSendEventEnabled === void 0) {
                ariaSendEventEnabled = true;
            }
            var sink = new AgaveSink(ariaAdditionalContext, ariaSendEventEnabled);
            return sink;
        };
        AgaveSink.prototype.initialize = function() {
            if (isWacAgave() || typeof OfficeExtension === "undefined") {
                if (SdxWacSink_SdxWacSink.isSupported()) {
                    this.connectSdxWacSink();
                } else if (this._ariaSendEventEnabled) {
                    this.connectAriaSink();
                } else {
                    this.failToInitialize();
                }
            } else if (OutlookSink_OutlookSink.isSupported()) {
                this.connectOutlookSink();
            } else {
                this._awaitingInitialization = true;
                getRichApiSink(false, this.onGetRichApi.bind(this));
            }
        };
        AgaveSink.prototype.onGetRichApi = function(richApiSink) {
            var _this = this;
            if (richApiSink) {
                this.connectRichApiSink(richApiSink);
            } else if (this._ariaSendEventEnabled) {
                this.connectAriaSink();
            } else {
                this.failToInitialize();
            }
            this._awaitingInitialization = false;
            this._eventQueue.forEach(function(event) {
                _this.sendTelemetryEvent(event);
            });
        };
        AgaveSink.prototype.failToInitialize = function() {
            this._isUsable = false;
            this._awaitingInitialization = false;
            var errorMessage = "AgaveSink could not find a suitable sink to use";
            logNotification(LogLevel.Error, Category.Sink, function() {
                return errorMessage;
            });
            throw new Error(errorMessage);
        };
        AgaveSink.prototype.sendTelemetryEvent = function(event) {
            if (this._awaitingInitialization && this._isUsable) {
                this._eventQueue.push(event);
            } else if (this._sink) {
                try {
                    this._sink.sendTelemetryEvent(event);
                } catch (error) {
                    var errorMessage_1;
                    if (error instanceof Error) {
                        errorMessage_1 = error.message;
                    } else {
                        errorMessage_1 = JSON.stringify(error);
                    }
                    logNotification(LogLevel.Error, Category.Sink, function() {
                        return "AgaveSink caught an error : " + errorMessage_1;
                    });
                }
            } else {
                logNotification(LogLevel.Error, Category.Sink, function() {
                    return "AgaveSink does not have an underlying sink";
                });
            }
        };
        AgaveSink.prototype.connectOutlookSink = function() {
            this._sink = new OutlookSink_OutlookSink();
            logNotification(LogLevel.Info, Category.Sink, function() {
                return "AgaveSink is using OutlookSink";
            });
        };
        AgaveSink.prototype.connectRichApiSink = function(sink) {
            this._sink = sink;
            logNotification(LogLevel.Info, Category.Sink, function() {
                return "AgaveSink is using RichApiSink";
            });
        };
        AgaveSink.prototype.connectAriaSink = function() {
            this._ariaAdditionalContext["Client.Id"] = getClientId();
            if (!this._ariaAdditionalContext["Session.Id"]) {
                this._ariaAdditionalContext["Session.Id"] = getSessionId();
            }
            var additionalDataFields = this.convertContextToTypedDataFields(this._ariaAdditionalContext);
            this._sink = new AriaSink_AriaSink(additionalDataFields);
            logNotification(LogLevel.Info, Category.Sink, function() {
                return "AgaveSink is using AriaSink";
            });
        };
        AgaveSink.prototype.connectSdxWacSink = function() {
            this._sink = new SdxWacSink_SdxWacSink();
            logNotification(LogLevel.Info, Category.Sink, function() {
                return "AgaveSink is using SdxWacSink";
            });
        };
        AgaveSink.prototype.convertContextToTypedDataFields = function(additionalContext) {
            var context = __assign({}, this.defaultAriaContext, additionalContext);
            var additionalDataFields = [];
            Object.keys(context).forEach(function(key) {
                additionalDataFields.push({
                    name: key,
                    value: context[key],
                    dataType: DataFieldType.String
                });
            });
            return additionalDataFields;
        };
        return AgaveSink;
    }();
    function getClientId() {
        if (clientId != null) {
            return clientId;
        }
        if (typeof localStorage !== "undefined") {
            clientId = localStorage.getItem(CLIENTID_LOCALSTORAGE_NAME);
        }
        if (clientId == null) {
            clientId = Utils_Utils.newGuid();
            if (typeof localStorage !== "undefined") {
                localStorage.setItem(CLIENTID_LOCALSTORAGE_NAME, clientId);
            }
        }
        return clientId;
    }
    function getSessionId() {
        if (sessionId == null) {
            sessionId = Utils_Utils.newGuid();
        }
        return sessionId;
    }
    __webpack_require__.d(__webpack_exports__, "AgaveSink", function() {
        return AgaveSink_AgaveSink;
    });
} ]);
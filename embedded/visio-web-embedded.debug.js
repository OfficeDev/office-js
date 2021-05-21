/*
	Copyright (c) Microsoft Corporation.  All rights reserved.
*/

/*
	Your use of this file is governed by the Microsoft Services Agreement http://go.microsoft.com/fwlink/?LinkId=266419.
*/

/*
* @overview es6-promise - a tiny implementation of Promises/A+.
* @copyright Copyright (c) 2014 Yehuda Katz, Tom Dale, Stefan Penner and contributors (Conversion to ES6 API by Jake Archibald)
* @license   Licensed under MIT license
*            See https://raw.githubusercontent.com/jakearchibald/es6-promise/master/LICENSE
* @version   2.3.0
*/


// Sources:
// osfweb: none
// runtime: 16.0\13326.10000
// core: 16.0\13326.10000
// host: 16.0.13327.34950



var __extends = (this && this.__extends) || (function () {
    var extendStatics = function (d, b) {
        extendStatics = Object.setPrototypeOf ||
            ({ __proto__: [] } instanceof Array && function (d, b) { d.__proto__ = b; }) ||
            function (d, b) { for (var p in b)
                if (b.hasOwnProperty(p))
                    d[p] = b[p]; };
        return extendStatics(d, b);
    };
    return function (d, b) {
        extendStatics(d, b);
        function __() { this.constructor = d; }
        d.prototype = b === null ? Object.create(b) : (__.prototype = b.prototype, new __());
    };
})();
var OfficeExtension;
(function (OfficeExtension) {
    var _Internal;
    (function (_Internal) {
        _Internal.OfficeRequire = function () {
            return null;
        }();
    })(_Internal = OfficeExtension._Internal || (OfficeExtension._Internal = {}));
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
                            nextTick = window.setImmediate;
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
                        return new _Internal.Error('Array Methods must be provided an Array');
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
    })(_Internal = OfficeExtension._Internal || (OfficeExtension._Internal = {}));
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
    })(_Internal = OfficeExtension._Internal || (OfficeExtension._Internal = {}));
    var OfficePromise = _Internal.OfficePromise;
    OfficeExtension.Promise = OfficePromise;
})(OfficeExtension || (OfficeExtension = {}));
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
            if (resultHandler === void 0) {
                resultHandler = null;
            }
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
            })["catch"](function (ex) {
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
    var ClientRetrieveResult = (function (_super) {
        __extends(ClientRetrieveResult, _super);
        function ClientRetrieveResult(m_shouldPolyfill) {
            var _this = _super.call(this) || this;
            _this.m_shouldPolyfill = m_shouldPolyfill;
            return _this;
        }
        ClientRetrieveResult.prototype._handleResult = function (value) {
            _super.prototype._handleResult.call(this, value);
            if (this.m_shouldPolyfill) {
                this.m_value = CommonUtility.changePropertyNameToCamelLowerCase(this.m_value);
            }
            this.m_value = this.removeItemNodes(this.m_value);
        };
        ClientRetrieveResult.prototype.removeItemNodes = function (value) {
            if (typeof value === 'object' && value !== null && value[CommonConstants.itemsLowerCase]) {
                value = value[CommonConstants.itemsLowerCase];
            }
            return CommonUtility.purifyJson(value);
        };
        return ClientRetrieveResult;
    }(ClientResult));
    OfficeExtension_1.ClientRetrieveResult = ClientRetrieveResult;
    var TraceActionResultHandler = (function () {
        function TraceActionResultHandler(callback) {
            this.callback = callback;
        }
        TraceActionResultHandler.prototype._handleResult = function (value) {
            if (this.callback) {
                this.callback();
            }
        };
        return TraceActionResultHandler;
    }());
    var ClientResultCallback = (function (_super) {
        __extends(ClientResultCallback, _super);
        function ClientResultCallback(callback) {
            var _this = _super.call(this) || this;
            _this.callback = callback;
            return _this;
        }
        ClientResultCallback.prototype._handleResult = function (value) {
            _super.prototype._handleResult.call(this, value);
            this.callback();
        };
        return ClientResultCallback;
    }(ClientResult));
    OfficeExtension_1.ClientResultCallback = ClientResultCallback;
    var OperationalApiHelper = (function () {
        function OperationalApiHelper() {
        }
        OperationalApiHelper.invokeMethod = function (obj, methodName, operationType, args, flags, resultProcessType) {
            if (operationType === void 0) {
                operationType = 0;
            }
            if (args === void 0) {
                args = [];
            }
            if (flags === void 0) {
                flags = 0;
            }
            if (resultProcessType === void 0) {
                resultProcessType = 0;
            }
            return CoreUtility.createPromise(function (resolve, reject) {
                var result = new ClientResult();
                var actionInfo = {
                    Id: obj._context._nextId(),
                    ActionType: 3,
                    Name: methodName,
                    ObjectPathId: obj._objectPath.objectPathInfo.Id,
                    ArgumentInfo: {}
                };
                var referencedArgumentObjectPaths = CommonUtility.setMethodArguments(obj._context, actionInfo.ArgumentInfo, args);
                var action = new Action(actionInfo, operationType, flags);
                action.referencedObjectPath = obj._objectPath;
                action.referencedArgumentObjectPaths = referencedArgumentObjectPaths;
                obj._context._addServiceApiAction(action, result, resolve, reject);
            });
        };
        OperationalApiHelper.invokeMethodWithClientResultCallback = function (callback, obj, methodName) {
            var operationType = 0;
            var args = [];
            var flags = 0;
            return CoreUtility.createPromise(function (resolve, reject) {
                var result = new ClientResultCallback(callback);
                var actionInfo = {
                    Id: obj._context._nextId(),
                    ActionType: 3,
                    Name: methodName,
                    ObjectPathId: obj._objectPath.objectPathInfo.Id,
                    ArgumentInfo: {}
                };
                var referencedArgumentObjectPaths = CommonUtility.setMethodArguments(obj._context, actionInfo.ArgumentInfo, args);
                var action = new Action(actionInfo, operationType, flags);
                action.referencedObjectPath = obj._objectPath;
                action.referencedArgumentObjectPaths = referencedArgumentObjectPaths;
                obj._context._addServiceApiAction(action, result, resolve, reject);
            });
        };
        OperationalApiHelper.invokeRetrieve = function (obj, select) {
            var shouldPolyfill = OfficeExtension_1._internalConfig.alwaysPolyfillClientObjectRetrieveMethod;
            if (!shouldPolyfill) {
                shouldPolyfill = !CommonUtility.isSetSupported('RichApiRuntime', '1.1');
            }
            var option;
            if (typeof select[0] === 'object' && select[0].hasOwnProperty('$all')) {
                if (!select[0]['$all']) {
                    throw OfficeExtension_1.Error._createInvalidArgError({});
                }
                option = select[0];
            }
            else {
                option = OperationalApiHelper._parseSelectOption(select);
            }
            return obj._retrieve(option, new ClientRetrieveResult(shouldPolyfill));
        };
        OperationalApiHelper._parseSelectOption = function (select) {
            if (!select || !select[0]) {
                throw OfficeExtension_1.Error._createInvalidArgError({});
            }
            var parsedSelect = select[0] && typeof select[0] !== 'string' ? select[0] : select;
            return Array.isArray(parsedSelect) ? parsedSelect : OperationalApiHelper.parseRecursiveSelect(parsedSelect);
        };
        OperationalApiHelper.parseRecursiveSelect = function (select) {
            var deconstruct = function (selectObj) {
                return Object.keys(selectObj).reduce(function (scalars, name) {
                    var value = selectObj[name];
                    if (typeof value === 'object') {
                        return scalars.concat(deconstruct(value).map(function (postfix) { return name + "/" + postfix; }));
                    }
                    if (value) {
                        return scalars.concat(name);
                    }
                    return scalars;
                }, []);
            };
            return deconstruct(select);
        };
        OperationalApiHelper.invokeRecursiveUpdate = function (obj, properties) {
            return CoreUtility.createPromise(function (resolve, reject) {
                obj._recursivelyUpdate(properties);
                var actionInfo = {
                    Id: obj._context._nextId(),
                    ActionType: 5,
                    Name: 'Trace',
                    ObjectPathId: 0
                };
                var action = new Action(actionInfo, 1, 4);
                obj._context._addServiceApiAction(action, null, resolve, reject);
            });
        };
        OperationalApiHelper.createRootServiceObject = function (type, context) {
            var objectPathInfo = {
                Id: context._nextId(),
                ObjectPathType: 1,
                Name: ''
            };
            var objectPath = new ObjectPath(objectPathInfo, null, false, false, 1, 4);
            return new type(context, objectPath);
        };
        OperationalApiHelper.createTopLevelServiceObject = function (type, context, typeName, isCollection, flags) {
            var objectPathInfo = {
                Id: context._nextId(),
                ObjectPathType: 2,
                Name: typeName
            };
            var objectPath = new ObjectPath(objectPathInfo, null, isCollection, false, 1, flags | 4);
            return new type(context, objectPath);
        };
        OperationalApiHelper.createPropertyObject = function (type, parent, propertyName, isCollection, flags) {
            var objectPathInfo = {
                Id: parent._context._nextId(),
                ObjectPathType: 4,
                Name: propertyName,
                ParentObjectPathId: parent._objectPath.objectPathInfo.Id
            };
            var objectPath = new ObjectPath(objectPathInfo, parent._objectPath, isCollection, false, 1, flags | 4);
            return new type(parent._context, objectPath);
        };
        OperationalApiHelper.createIndexerObject = function (type, parent, args) {
            var objectPathInfo = {
                Id: parent._context._nextId(),
                ObjectPathType: 5,
                Name: '',
                ParentObjectPathId: parent._objectPath.objectPathInfo.Id,
                ArgumentInfo: {}
            };
            objectPathInfo.ArgumentInfo.Arguments = args;
            var objectPath = new ObjectPath(objectPathInfo, parent._objectPath, false, false, 1, 4);
            return new type(parent._context, objectPath);
        };
        OperationalApiHelper.createMethodObject = function (type, parent, methodName, operationType, args, isCollection, isInvalidAfterRequest, getByIdMethodName, flags) {
            var id = parent._context._nextId();
            var objectPathInfo = {
                Id: id,
                ObjectPathType: 3,
                Name: methodName,
                ParentObjectPathId: parent._objectPath.objectPathInfo.Id,
                ArgumentInfo: {}
            };
            var argumentObjectPaths = CommonUtility.setMethodArguments(parent._context, objectPathInfo.ArgumentInfo, args);
            var objectPath = new ObjectPath(objectPathInfo, parent._objectPath, isCollection, isInvalidAfterRequest, operationType, flags);
            objectPath.argumentObjectPaths = argumentObjectPaths;
            objectPath.getByIdMethodName = getByIdMethodName;
            var o = new type(parent._context, objectPath);
            return o;
        };
        OperationalApiHelper.createAndInstantiateMethodObject = function (type, parent, methodName, operationType, args, isCollection, isInvalidAfterRequest, getByIdMethodName, flags) {
            return CoreUtility.createPromise(function (resolve, reject) {
                var objectPathInfo = {
                    Id: parent._context._nextId(),
                    ObjectPathType: 3,
                    Name: methodName,
                    ParentObjectPathId: parent._objectPath.objectPathInfo.Id,
                    ArgumentInfo: {}
                };
                var argumentObjectPaths = CommonUtility.setMethodArguments(parent._context, objectPathInfo.ArgumentInfo, args);
                var objectPath = new ObjectPath(objectPathInfo, parent._objectPath, isCollection, isInvalidAfterRequest, operationType, flags);
                objectPath.argumentObjectPaths = argumentObjectPaths;
                objectPath.getByIdMethodName = getByIdMethodName;
                var result = new ClientResult();
                var actionInfo = {
                    Id: parent._context._nextId(),
                    ActionType: 1,
                    Name: '',
                    ObjectPathId: objectPath.objectPathInfo.Id,
                    QueryInfo: {}
                };
                var action = new Action(actionInfo, 1, 4);
                action.referencedObjectPath = objectPath;
                parent._context._addServiceApiAction(action, result, function () { return resolve(new type(parent._context, objectPath)); }, reject);
            });
        };
        OperationalApiHelper.createTraceAction = function (context, callback) {
            return CoreUtility.createPromise(function (resolve, reject) {
                var actionInfo = {
                    Id: context._nextId(),
                    ActionType: 5,
                    Name: 'Trace',
                    ObjectPathId: 0
                };
                var action = new Action(actionInfo, 1, 4);
                var result = new TraceActionResultHandler(callback);
                context._addServiceApiAction(action, result, resolve, reject);
            });
        };
        OperationalApiHelper.localDocumentContext = new ClientRequestContextBase();
        return OperationalApiHelper;
    }());
    OfficeExtension_1.OperationalApiHelper = OperationalApiHelper;
    var GenericEventRegistryOperational = (function () {
        function GenericEventRegistryOperational(eventId, targetId, eventArgumentTransform) {
            this.eventId = eventId;
            this.targetId = targetId;
            this.eventArgumentTransform = eventArgumentTransform;
            this.registeredCallbacks = [];
        }
        GenericEventRegistryOperational.prototype.add = function (callback) {
            if (this.hasZero()) {
                GenericEventRegistration.getGenericEventRegistration().register(this.eventId, this.targetId, this.registerCallback);
            }
            this.registeredCallbacks.push(callback);
        };
        GenericEventRegistryOperational.prototype.remove = function (callback) {
            var index = this.registeredCallbacks.lastIndexOf(callback);
            if (index !== -1) {
                this.registeredCallbacks.splice(index, 1);
            }
        };
        GenericEventRegistryOperational.prototype.removeAll = function () {
            this.registeredCallbacks = [];
            GenericEventRegistration.getGenericEventRegistration().unregister(this.eventId, this.targetId, this.registerCallback);
        };
        GenericEventRegistryOperational.prototype.hasZero = function () {
            return this.registeredCallbacks.length === 0;
        };
        Object.defineProperty(GenericEventRegistryOperational.prototype, "registerCallback", {
            get: function () {
                var i = this;
                if (!this.outsideCallback) {
                    this.outsideCallback = function (argument) {
                        i.call(argument);
                    };
                }
                return this.outsideCallback;
            },
            enumerable: true,
            configurable: true
        });
        GenericEventRegistryOperational.prototype.call = function (rawEventArguments) {
            var _this = this;
            this.eventArgumentTransform(rawEventArguments).then(function (eventArguments) {
                var promises = _this.registeredCallbacks.map(function (callback) { return GenericEventRegistryOperational.callCallback(callback, eventArguments); });
                CoreUtility.Promise.all(promises);
            });
        };
        GenericEventRegistryOperational.callCallback = function (callback, eventArguments) {
            return CoreUtility._createPromiseFromResult(null)
                .then(GenericEventRegistryOperational.wrapCallbackInFunction(callback, eventArguments))["catch"](function (e) {
                CoreUtility.log('Error when invoke handler: ' + JSON.stringify(e));
            });
        };
        GenericEventRegistryOperational.wrapCallbackInFunction = function (callback, args) {
            return function () { return callback(args); };
        };
        return GenericEventRegistryOperational;
    }());
    OfficeExtension_1.GenericEventRegistryOperational = GenericEventRegistryOperational;
    var GlobalEventRegistryOperational = (function () {
        function GlobalEventRegistryOperational() {
            this.eventToTargetToHandlerMap = {};
        }
        Object.defineProperty(GlobalEventRegistryOperational, "globalEventRegistry", {
            get: function () {
                if (!GlobalEventRegistryOperational.singleton) {
                    GlobalEventRegistryOperational.singleton = new GlobalEventRegistryOperational();
                }
                return GlobalEventRegistryOperational.singleton;
            },
            enumerable: true,
            configurable: true
        });
        GlobalEventRegistryOperational.getGlobalEventRegistry = function (eventId, targetId, eventArgumentTransform) {
            var global = GlobalEventRegistryOperational.globalEventRegistry;
            var mapGlobal = global.eventToTargetToHandlerMap;
            if (!mapGlobal.hasOwnProperty(eventId)) {
                mapGlobal[eventId] = {};
            }
            var mapEvent = mapGlobal[eventId];
            if (!mapEvent.hasOwnProperty(targetId)) {
                mapEvent[targetId] = new GenericEventRegistryOperational(eventId, targetId, eventArgumentTransform);
            }
            var target = mapEvent[targetId];
            return target;
        };
        GlobalEventRegistryOperational.singleton = undefined;
        return GlobalEventRegistryOperational;
    }());
    OfficeExtension_1.GlobalEventRegistryOperational = GlobalEventRegistryOperational;
    var GenericEventHandlerOperational = (function () {
        function GenericEventHandlerOperational(genericEventInfo) {
            this.genericEventInfo = genericEventInfo;
        }
        GenericEventHandlerOperational.prototype.add = function (callback) {
            var _this = this;
            var eventRegistered = undefined;
            var promise = CoreUtility.createPromise(function (resolve) {
                eventRegistered = resolve;
            });
            var addCallback = function () {
                var eventId = _this.genericEventInfo.eventType;
                var targetId = _this.genericEventInfo.getTargetIdFunc();
                var event = GlobalEventRegistryOperational.getGlobalEventRegistry(eventId, targetId, _this.genericEventInfo.eventArgsTransformFunc);
                event.add(callback);
                eventRegistered();
            };
            this.register();
            this.createTrace(addCallback);
            return promise;
        };
        GenericEventHandlerOperational.prototype.remove = function (callback) {
            var _this = this;
            var removeCallback = function () {
                var eventId = _this.genericEventInfo.eventType;
                var targetId = _this.genericEventInfo.getTargetIdFunc();
                var event = GlobalEventRegistryOperational.getGlobalEventRegistry(eventId, targetId, _this.genericEventInfo.eventArgsTransformFunc);
                event.remove(callback);
            };
            this.register();
            this.createTrace(removeCallback);
        };
        GenericEventHandlerOperational.prototype.removeAll = function () {
            var _this = this;
            var removeAllCallback = function () {
                var eventId = _this.genericEventInfo.eventType;
                var targetId = _this.genericEventInfo.getTargetIdFunc();
                var event = GlobalEventRegistryOperational.getGlobalEventRegistry(eventId, targetId, _this.genericEventInfo.eventArgsTransformFunc);
                event.removeAll();
            };
            this.unregister();
            this.createTrace(removeAllCallback);
        };
        GenericEventHandlerOperational.prototype.createTrace = function (callback) {
            OperationalApiHelper.createTraceAction(this.genericEventInfo.object._context, callback);
        };
        GenericEventHandlerOperational.prototype.register = function () {
            var operationType = 0;
            var args = [];
            var flags = 0;
            OperationalApiHelper.invokeMethod(this.genericEventInfo.object, this.genericEventInfo.register, operationType, args, flags);
            if (!GenericEventRegistration.getGenericEventRegistration().isReady) {
                GenericEventRegistration.getGenericEventRegistration().ready();
            }
        };
        GenericEventHandlerOperational.prototype.unregister = function () {
            OperationalApiHelper.invokeMethod(this.genericEventInfo.object, this.genericEventInfo.unregister);
        };
        return GenericEventHandlerOperational;
    }());
    OfficeExtension_1.GenericEventHandlerOperational = GenericEventHandlerOperational;
    var EventHelper = (function () {
        function EventHelper() {
        }
        EventHelper.invokeOn = function (eventHandler, callback, options) {
            var promiseResolve = undefined;
            var promise = CoreUtility.createPromise(function (resolve, reject) {
                promiseResolve = resolve;
            });
            eventHandler.add(callback).then(function () {
                promiseResolve({});
            });
            return promise;
        };
        EventHelper.invokeOff = function (genericEventHandlersOpObj, eventHandler, eventName, callback) {
            if (!eventName && !callback) {
                var allGenericEventHandlersOp = Object.keys(genericEventHandlersOpObj).map(function (eventName) { return genericEventHandlersOpObj[eventName]; });
                return EventHelper.invokeAllOff(allGenericEventHandlersOp);
            }
            if (!eventName) {
                return CoreUtility._createPromiseFromException(eventName + " must be supplied if handler is supplied.");
            }
            if (callback) {
                eventHandler.remove(callback);
            }
            else {
                eventHandler.removeAll();
            }
            return CoreUtility.createPromise(function (resolve, reject) { return resolve(); });
        };
        EventHelper.invokeAllOff = function (allGenericEventHandlersOperational) {
            allGenericEventHandlersOperational.forEach(function (genericEventHandlerOperational) {
                genericEventHandlerOperational.removeAll();
            });
            return CoreUtility.createPromise(function (resolve, reject) { return resolve(); });
        };
        return EventHelper;
    }());
    OfficeExtension_1.EventHelper = EventHelper;
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
            if (resultHandler === void 0) {
                resultHandler = null;
            }
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
                return _this.processPendingEventHandlers(req)["catch"](function (ex) {
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
                })["catch"](function (ex) {
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
            if (numCleanupAttempts === void 0) {
                numCleanupAttempts = 3;
            }
            if (retryDelay === void 0) {
                retryDelay = 5000;
            }
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
            if (numCleanupAttempts === void 0) {
                numCleanupAttempts = 3;
            }
            if (retryDelay === void 0) {
                retryDelay = 5000;
            }
            return ClientRequestContext._runBatchCommon(0, functionName, receivedRunArgs, ctxInitializer, numCleanupAttempts, retryDelay, onBeforeRun, onCleanupSuccess, onCleanupFailure);
        };
        ClientRequestContext._runExplicitBatch = function (functionName, receivedRunArgs, ctxInitializer, onBeforeRun, numCleanupAttempts, retryDelay, onCleanupSuccess, onCleanupFailure) {
            if (numCleanupAttempts === void 0) {
                numCleanupAttempts = 3;
            }
            if (retryDelay === void 0) {
                retryDelay = 5000;
            }
            return ClientRequestContext._runBatchCommon(1, functionName, receivedRunArgs, ctxInitializer, numCleanupAttempts, retryDelay, onBeforeRun, onCleanupSuccess, onCleanupFailure);
        };
        ClientRequestContext._runBatchCommon = function (batchMode, functionName, receivedRunArgs, ctxInitializer, numCleanupAttempts, retryDelay, onBeforeRun, onCleanupSuccess, onCleanupFailure) {
            if (numCleanupAttempts === void 0) {
                numCleanupAttempts = 3;
            }
            if (retryDelay === void 0) {
                retryDelay = 5000;
            }
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
            if (code === void 0) {
                code = CoreResourceStrings.invalidArgument;
            }
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
            })["catch"](function (error) {
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
                    })["catch"](function () {
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
                        .then(function () { return _this.m_eventInfo.unregisterFunc(_this.m_callback); })["catch"](function (ex) {
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
                    .then(this.createFireOneEventHandlerFunc(handler, args))["catch"](function (ex) {
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
                            events: elem[8]
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
                            undoableApiSetInfoOrdinal: elem[5]
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
    var versionToken = 1;
    var internalConfiguration = {
        invokeRequestModifier: function (request) {
            request.DdaMethod.Version = versionToken;
            return request;
        },
        invokeResponseModifier: function (args) {
            versionToken = args.Version;
            if (args.Error) {
                args.error = {};
                args.error.Code = args.Error;
            }
            return args;
        }
    };
    var CommunicationConstants;
    (function (CommunicationConstants) {
        CommunicationConstants["SendingId"] = "sId";
        CommunicationConstants["RespondingId"] = "rId";
        CommunicationConstants["CommandKey"] = "command";
        CommunicationConstants["SessionInfoKey"] = "sessionInfo";
        CommunicationConstants["ParamsKey"] = "params";
        CommunicationConstants["ApiReadyCommand"] = "apiready";
        CommunicationConstants["ExecuteMethodCommand"] = "executeMethod";
        CommunicationConstants["GetAppContextCommand"] = "getAppContext";
        CommunicationConstants["RegisterEventCommand"] = "registerEvent";
        CommunicationConstants["UnregisterEventCommand"] = "unregisterEvent";
        CommunicationConstants["FireEventCommand"] = "fireEvent";
    })(CommunicationConstants || (CommunicationConstants = {}));
    var EmbeddedConstants = (function () {
        function EmbeddedConstants() {
        }
        EmbeddedConstants.sessionContext = 'sc';
        EmbeddedConstants.embeddingPageOrigin = 'EmbeddingPageOrigin';
        EmbeddedConstants.embeddingPageSessionInfo = 'EmbeddingPageSessionInfo';
        return EmbeddedConstants;
    }());
    OfficeExtension_1.EmbeddedConstants = EmbeddedConstants;
    var EmbeddedSession = (function (_super) {
        __extends(EmbeddedSession, _super);
        function EmbeddedSession(url, options) {
            var _this = _super.call(this) || this;
            _this.m_chosenWindow = null;
            _this.m_chosenOrigin = null;
            _this.m_enabled = true;
            _this.m_onMessageHandler = _this._onMessage.bind(_this);
            _this.m_callbackList = {};
            _this.m_id = 0;
            _this.m_timeoutId = -1;
            _this.m_appContext = null;
            _this.m_url = url;
            _this.m_options = options;
            if (!_this.m_options) {
                _this.m_options = { sessionKey: Math.random().toString() };
            }
            if (!_this.m_options.sessionKey) {
                _this.m_options.sessionKey = Math.random().toString();
            }
            if (!_this.m_options.container) {
                _this.m_options.container = document.body;
            }
            if (!_this.m_options.timeoutInMilliseconds) {
                _this.m_options.timeoutInMilliseconds = 60000;
            }
            if (!_this.m_options.height) {
                _this.m_options.height = '400px';
            }
            if (!_this.m_options.width) {
                _this.m_options.width = '100%';
            }
            if (!(_this.m_options.webApplication &&
                _this.m_options.webApplication.accessToken &&
                _this.m_options.webApplication.accessTokenTtl)) {
                _this.m_options.webApplication = null;
            }
            return _this;
        }
        EmbeddedSession.prototype._getIFrameSrc = function () {
            var origin = window.location.protocol + '//' + window.location.host;
            var toAppend = EmbeddedConstants.embeddingPageOrigin +
                '=' +
                encodeURIComponent(origin) +
                '&' +
                EmbeddedConstants.embeddingPageSessionInfo +
                '=' +
                encodeURIComponent(this.m_options.sessionKey);
            var useHash = false;
            if (this.m_url.toLowerCase().indexOf('/_layouts/preauth.aspx') > 0 ||
                this.m_url.toLowerCase().indexOf('/_layouts/15/preauth.aspx') > 0) {
                useHash = true;
            }
            var a = document.createElement('a');
            a.href = this.m_url;
            if (this.m_options.webApplication) {
                var toAppendWAC = EmbeddedConstants.embeddingPageOrigin +
                    '=' +
                    origin +
                    '&' +
                    EmbeddedConstants.embeddingPageSessionInfo +
                    '=' +
                    this.m_options.sessionKey;
                if (a.search.length === 0 || a.search === '?') {
                    a.search = '?' + EmbeddedConstants.sessionContext + '=' + encodeURIComponent(toAppendWAC);
                }
                else {
                    a.search = a.search + '&' + EmbeddedConstants.sessionContext + '=' + encodeURIComponent(toAppendWAC);
                }
            }
            else if (useHash) {
                if (a.hash.length === 0 || a.hash === '#') {
                    a.hash = '#' + toAppend;
                }
                else {
                    a.hash = a.hash + '&' + toAppend;
                }
            }
            else {
                if (a.search.length === 0 || a.search === '?') {
                    a.search = '?' + toAppend;
                }
                else {
                    a.search = a.search + '&' + toAppend;
                }
            }
            var iframeSrc = a.href;
            return iframeSrc;
        };
        EmbeddedSession.prototype.init = function () {
            var _this = this;
            window.addEventListener('message', this.m_onMessageHandler);
            var iframeSrc = this._getIFrameSrc();
            return CoreUtility.createPromise(function (resolve, reject) {
                var iframeElement = document.createElement('iframe');
                if (_this.m_options.id) {
                    iframeElement.id = _this.m_options.id;
                    iframeElement.name = _this.m_options.id;
                }
                iframeElement.style.height = _this.m_options.height;
                iframeElement.style.width = _this.m_options.width;
                if (!_this.m_options.webApplication) {
                    iframeElement.src = iframeSrc;
                    _this.m_options.container.appendChild(iframeElement);
                }
                else {
                    var webApplicationForm = document.createElement('form');
                    webApplicationForm.setAttribute('action', iframeSrc);
                    webApplicationForm.setAttribute('method', 'post');
                    webApplicationForm.setAttribute('target', iframeElement.name);
                    _this.m_options.container.appendChild(webApplicationForm);
                    var token_input = document.createElement('input');
                    token_input.setAttribute('type', 'hidden');
                    token_input.setAttribute('name', 'access_token');
                    token_input.setAttribute('value', _this.m_options.webApplication.accessToken);
                    webApplicationForm.appendChild(token_input);
                    var token_ttl_input = document.createElement('input');
                    token_ttl_input.setAttribute('type', 'hidden');
                    token_ttl_input.setAttribute('name', 'access_token_ttl');
                    token_ttl_input.setAttribute('value', _this.m_options.webApplication.accessTokenTtl);
                    webApplicationForm.appendChild(token_ttl_input);
                    _this.m_options.container.appendChild(iframeElement);
                    webApplicationForm.submit();
                }
                _this.m_timeoutId = window.setTimeout(function () {
                    _this.close();
                    var err = Utility.createRuntimeError(CoreErrorCodes.timeout, CoreUtility._getResourceString(CoreResourceStrings.timeout), 'EmbeddedSession.init');
                    reject(err);
                }, _this.m_options.timeoutInMilliseconds);
                _this.m_promiseResolver = resolve;
            });
        };
        EmbeddedSession.prototype._invoke = function (method, callback, params) {
            if (!this.m_enabled) {
                callback(5001, null);
                return;
            }
            if (internalConfiguration.invokeRequestModifier) {
                params = internalConfiguration.invokeRequestModifier(params);
            }
            this._sendMessageWithCallback(this.m_id++, method, params, function (args) {
                if (internalConfiguration.invokeResponseModifier) {
                    args = internalConfiguration.invokeResponseModifier(args);
                }
                var errorCode = args['Error'];
                delete args['Error'];
                callback(errorCode || 0, args);
            });
        };
        EmbeddedSession.prototype.close = function () {
            window.removeEventListener('message', this.m_onMessageHandler);
            window.clearTimeout(this.m_timeoutId);
            this.m_enabled = false;
        };
        Object.defineProperty(EmbeddedSession.prototype, "eventRegistration", {
            get: function () {
                if (!this.m_sessionEventManager) {
                    this.m_sessionEventManager = new EventRegistration(this._registerEventImpl.bind(this), this._unregisterEventImpl.bind(this));
                }
                return this.m_sessionEventManager;
            },
            enumerable: true,
            configurable: true
        });
        EmbeddedSession.prototype._createRequestExecutorOrNull = function () {
            return new EmbeddedRequestExecutor(this);
        };
        EmbeddedSession.prototype._resolveRequestUrlAndHeaderInfo = function () {
            return CoreUtility._createPromiseFromResult(null);
        };
        EmbeddedSession.prototype._registerEventImpl = function (eventId, targetId) {
            var _this = this;
            return CoreUtility.createPromise(function (resolve, reject) {
                _this._sendMessageWithCallback(_this.m_id++, CommunicationConstants.RegisterEventCommand, { EventId: eventId, TargetId: targetId }, function () {
                    resolve(null);
                });
            });
        };
        EmbeddedSession.prototype._unregisterEventImpl = function (eventId, targetId) {
            var _this = this;
            return CoreUtility.createPromise(function (resolve, reject) {
                _this._sendMessageWithCallback(_this.m_id++, CommunicationConstants.UnregisterEventCommand, { EventId: eventId, TargetId: targetId }, function () {
                    resolve();
                });
            });
        };
        EmbeddedSession.prototype._onMessage = function (event) {
            var _this = this;
            if (!this.m_enabled) {
                return;
            }
            if (this.m_chosenWindow && (this.m_chosenWindow !== event.source || this.m_chosenOrigin !== event.origin)) {
                return;
            }
            var eventData = event.data;
            if (eventData && eventData[CommunicationConstants.CommandKey] === CommunicationConstants.ApiReadyCommand) {
                if (!this.m_chosenWindow &&
                    this._isValidDescendant(event.source) &&
                    eventData[CommunicationConstants.SessionInfoKey] === this.m_options.sessionKey) {
                    this.m_chosenWindow = event.source;
                    this.m_chosenOrigin = event.origin;
                    this._sendMessageWithCallback(this.m_id++, CommunicationConstants.GetAppContextCommand, null, function (appContext) {
                        _this._setupContext(appContext);
                        window.clearTimeout(_this.m_timeoutId);
                        _this.m_promiseResolver();
                    });
                }
                return;
            }
            if (eventData && eventData[CommunicationConstants.CommandKey] === CommunicationConstants.FireEventCommand) {
                var msg = eventData[CommunicationConstants.ParamsKey];
                var eventId = msg['EventId'];
                var targetId = msg['TargetId'];
                var data = msg['Data'];
                if (this.m_sessionEventManager) {
                    var handlers = this.m_sessionEventManager.getHandlers(eventId, targetId);
                    for (var i = 0; i < handlers.length; i++) {
                        handlers[i](data);
                    }
                }
                return;
            }
            if (eventData && eventData.hasOwnProperty(CommunicationConstants.RespondingId)) {
                var rId = eventData[CommunicationConstants.RespondingId];
                var callback = this.m_callbackList[rId];
                if (typeof callback === 'function') {
                    callback(eventData[CommunicationConstants.ParamsKey]);
                }
                delete this.m_callbackList[rId];
            }
        };
        EmbeddedSession.prototype._sendMessageWithCallback = function (id, command, data, callback) {
            this.m_callbackList[id] = callback;
            var message = {};
            message[CommunicationConstants.SendingId] = id;
            message[CommunicationConstants.CommandKey] = command;
            message[CommunicationConstants.ParamsKey] = data;
            this.m_chosenWindow.postMessage(JSON.stringify(message), this.m_chosenOrigin);
        };
        EmbeddedSession.prototype._isValidDescendant = function (wnd) {
            var container = this.m_options.container || document.body;
            function doesFrameWindow(containerWindow) {
                if (containerWindow === wnd) {
                    return true;
                }
                for (var i = 0, len = containerWindow.frames.length; i < len; i++) {
                    if (doesFrameWindow(containerWindow.frames[i])) {
                        return true;
                    }
                }
                return false;
            }
            var iframes = container.getElementsByTagName('iframe');
            for (var i = 0, len = iframes.length; i < len; i++) {
                if (doesFrameWindow(iframes[i].contentWindow)) {
                    return true;
                }
            }
            return false;
        };
        EmbeddedSession.prototype._setupContext = function (appContext) {
            if (!(this.m_appContext = appContext)) {
                return;
            }
        };
        return EmbeddedSession;
    }(SessionBase));
    OfficeExtension_1.EmbeddedSession = EmbeddedSession;
    var EmbeddedRequestExecutor = (function () {
        function EmbeddedRequestExecutor(session) {
            this.m_session = session;
        }
        EmbeddedRequestExecutor.prototype.executeAsync = function (customData, requestFlags, requestMessage) {
            var _this = this;
            var messageSafearray = RichApiMessageUtility.buildMessageArrayForIRequestExecutor(customData, requestFlags, requestMessage, EmbeddedRequestExecutor.SourceLibHeaderValue);
            return CoreUtility.createPromise(function (resolve, reject) {
                _this.m_session._invoke(CommunicationConstants.ExecuteMethodCommand, function (status, result) {
                    CoreUtility.log('Response:');
                    CoreUtility.log(JSON.stringify(result));
                    var response;
                    if (status == 0) {
                        response = RichApiMessageUtility.buildResponseOnSuccess(RichApiMessageUtility.getResponseBodyFromSafeArray(result.Data), RichApiMessageUtility.getResponseHeadersFromSafeArray(result.Data));
                    }
                    else {
                        response = RichApiMessageUtility.buildResponseOnError(result.error.Code, result.error.Message);
                    }
                    resolve(response);
                }, EmbeddedRequestExecutor._transformMessageArrayIntoParams(messageSafearray));
            });
        };
        EmbeddedRequestExecutor._transformMessageArrayIntoParams = function (msgArray) {
            return {
                ArrayData: msgArray,
                DdaMethod: {
                    DispatchId: EmbeddedRequestExecutor.DispidExecuteRichApiRequestMethod
                }
            };
        };
        EmbeddedRequestExecutor.DispidExecuteRichApiRequestMethod = 93;
        EmbeddedRequestExecutor.SourceLibHeaderValue = 'Embedded';
        return EmbeddedRequestExecutor;
    }());
})(OfficeExtension || (OfficeExtension = {}));
var __extends = (this && this.__extends) || (function () {
    var extendStatics = function (d, b) {
        extendStatics = Object.setPrototypeOf ||
            ({ __proto__: [] } instanceof Array && function (d, b) { d.__proto__ = b; }) ||
            function (d, b) { for (var p in b)
                if (b.hasOwnProperty(p))
                    d[p] = b[p]; };
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
        function fulfilled(value) { try {
            step(generator.next(value));
        }
        catch (e) {
            reject(e);
        } }
        function rejected(value) { try {
            step(generator["throw"](value));
        }
        catch (e) {
            reject(e);
        } }
        function step(result) { result.done ? resolve(result.value) : new P(function (resolve) { resolve(result.value); }).then(fulfilled, rejected); }
        step((generator = generator.apply(thisArg, _arguments || [])).next());
    });
};
var __generator = (this && this.__generator) || function (thisArg, body) {
    var _ = { label: 0, sent: function () { if (t[0] & 1)
            throw t[1]; return t[1]; }, trys: [], ops: [] }, f, y, t, g;
    return g = { next: verb(0), "throw": verb(1), "return": verb(2) }, typeof Symbol === "function" && (g[Symbol.iterator] = function () { return this; }), g;
    function verb(n) { return function (v) { return step([n, v]); }; }
    function step(op) {
        if (f)
            throw new TypeError("Generator is already executing.");
        while (_)
            try {
                if (f = 1, y && (t = op[0] & 2 ? y["return"] : op[0] ? y["throw"] || ((t = y["return"]) && t.call(y), 0) : y.next) && !(t = t.call(y, op[1])).done)
                    return t;
                if (y = 0, t)
                    op = [op[0] & 2, t.value];
                switch (op[0]) {
                    case 0:
                    case 1:
                        t = op;
                        break;
                    case 4:
                        _.label++;
                        return { value: op[1], done: false };
                    case 5:
                        _.label++;
                        y = op[1];
                        op = [0];
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
                        if (op[0] === 3 && (!t || (op[1] > t[0] && op[1] < t[3]))) {
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
                        if (t[2])
                            _.ops.pop();
                        _.trys.pop();
                        continue;
                }
                op = body.call(thisArg, _);
            }
            catch (e) {
                op = [6, e];
                y = 0;
            }
            finally {
                f = t = 0;
            }
        if (op[0] & 5)
            throw op[1];
        return { value: op[0] ? op[1] : void 0, done: true };
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
                "value": this._V
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
    var _typeSkill = "Skill";
    var Skill = (function (_super) {
        __extends(Skill, _super);
        function Skill() {
            return _super !== null && _super.apply(this, arguments) || this;
        }
        Object.defineProperty(Skill.prototype, "_className", {
            get: function () {
                return "Skill";
            },
            enumerable: true,
            configurable: true
        });
        Skill.prototype.executeAction = function (paneId, actionId, actionDescriptor) {
            return _invokeMethod(this, "ExecuteAction", 1, [paneId, actionId, actionDescriptor], 4 | 1, 0);
        };
        Skill.prototype.notifyPaneEvent = function (paneId, eventDescriptor) {
            _invokeMethod(this, "NotifyPaneEvent", 1, [paneId, eventDescriptor], 4 | 1, 0);
        };
        Skill.prototype.registerHostSkillEvent = function () {
            _invokeMethod(this, "RegisterHostSkillEvent", 0, [], 1, 0);
        };
        Skill.prototype.testFireEvent = function () {
            _invokeMethod(this, "TestFireEvent", 0, [], 1, 0);
        };
        Skill.prototype.unregisterHostSkillEvent = function () {
            _invokeMethod(this, "UnregisterHostSkillEvent", 0, [], 1, 0);
        };
        Skill.prototype._handleResult = function (value) {
            _super.prototype._handleResult.call(this, value);
            if (_isNullOrUndefined(value))
                return;
            var obj = value;
            _fixObjectPathIfNecessary(this, obj);
        };
        Skill.prototype._handleRetrieveResult = function (value, result) {
            _super.prototype._handleRetrieveResult.call(this, value, result);
            _processRetrieveResult(this, value, result);
        };
        Skill.newObject = function (context) {
            return _createTopLevelServiceObject(OfficeCore.Skill, context, "Microsoft.SkillApi.Skill", false, 4);
        };
        Object.defineProperty(Skill.prototype, "onHostSkillEvent", {
            get: function () {
                var _this = this;
                if (!this.m_hostSkillEvent) {
                    this.m_hostSkillEvent = new OfficeExtension.GenericEventHandlers(this.context, this, "HostSkillEvent", {
                        eventType: 65538,
                        registerFunc: function () { return _this.registerHostSkillEvent(); },
                        unregisterFunc: function () { return _this.unregisterHostSkillEvent(); },
                        getTargetIdFunc: function () { return ""; },
                        eventArgsTransformFunc: function (value) {
                            var event = _CC.Skill_HostSkillEvent_EventArgsTransform(_this, value);
                            return OfficeExtension.Utility._createPromiseFromResult(event);
                        }
                    });
                }
                return this.m_hostSkillEvent;
            },
            enumerable: true,
            configurable: true
        });
        Skill.prototype.toJSON = function () {
            return _toJson(this, {}, {});
        };
        return Skill;
    }(OfficeExtension.ClientObject));
    OfficeCore.Skill = Skill;
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
                        reject({ code: ErrorCode.GetAuthContextAsyncMissing, message: (typeof (Strings) !== 'undefined' && Strings.OfficeOM.L_ImplicitGetAuthContextMissing) ? Strings.OfficeOM.L_ImplicitGetAuthContextMissing : "" });
                    }
                    Office.context.webAuth.getAuthContextAsync(function (result) {
                        if (result.status === "succeeded") {
                            retrievedAuthContext = true;
                            var authContext = result.value;
                            if (!authContext || authContext.isAnonymous) {
                                reject({ code: ErrorCode.CannotGetAuthContext, message: (typeof (Strings) !== 'undefined' && Strings.OfficeOM.L_ImplicitGetAuthContextMissing) ? Strings.OfficeOM.L_ImplicitGetAuthContextMissing : "" });
                                return;
                            }
                            var isMsa = authContext.authorityType.toLowerCase() === 'msa';
                            OSF.WebAuth.config = {
                                idp: authContext.authorityType.toLowerCase(),
                                appIds: [isMsa ? (authContext.msaAppId) ? authContext.msaAppId : authContext.appId : authContext.appId],
                                authority: (OfficeFirstPartyAuth.authorityOverride) ? OfficeFirstPartyAuth.authorityOverride : authContext.authority,
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
                                reject({ code: ErrorCode.PackageNotLoaded, message: (typeof (Strings) !== 'undefined' && Strings.OfficeOM.L_ImplicitNotLoaded) ? Strings.OfficeOM.L_ImplicitNotLoaded : "" });
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
        return new OfficeExtension.CoreUtility.Promise(function (resolve, reject) {
            if (OSF._OfficeAppFactory.getHostInfo().hostPlatform == "web") {
                Office.context.webAuth.getAuthContextAsync(function (result) {
                    var supportsAuthToken = false;
                    if (result.status === "succeeded") {
                        retrievedAuthContext = true;
                        var authContext = result.value;
                        if (authContext.supportsAuthToken) {
                            supportsAuthToken = true;
                        }
                    }
                    if (!supportsAuthToken) {
                        if (OSF.WebAuth && OSF.WebAuth.loaded) {
                            if (behaviorOption && behaviorOption.forceRefresh) {
                                OSF.WebAuth.clearCache();
                            }
                            var identityType_1 = (OSF.WebAuth.config.idp.toLowerCase() == "msa")
                                ? OfficeCore.IdentityType.microsoftAccount
                                : OfficeCore.IdentityType.organizationAccount;
                            if (OSF.WebAuth.config.appIds[0]) {
                                OSF.WebAuth.getToken(options.resource, OSF.WebAuth.config.appIds[0], OSF._OfficeAppFactory.getHostInfo().osfControlAppCorrelationId, (behaviorOption && behaviorOption.popup) ? behaviorOption.popup : null).then(function (result) {
                                    logAcquireEvent(result, true, options.resource, (behaviorOption && behaviorOption.popup) ? behaviorOption.popup : null);
                                    resolve({ accessToken: result.Token, tokenIdenityType: identityType_1 });
                                })["catch"](function (result) {
                                    logAcquireEvent(result, false, options.resource, (behaviorOption && behaviorOption.popup) ? behaviorOption.popup : null, result.ErrorCode);
                                    reject({ code: result.ErrorCode, message: result.ErrorMessage });
                                });
                            }
                        }
                        else {
                            logUnexpectedAcquireEvent(OSF.WebAuth.loaded, OSF.WebAuth.loadAttempts);
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
                        context_1.sync()["catch"](function () {
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
                })["catch"](function (e) {
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
    function logUnexpectedAcquireEvent(loaded, loadAttempts) {
        if (OfficeFirstPartyAuth.debugging) {
            console.log("Logging Implicit unexpected acquire event");
        }
        if (typeof OTel !== "undefined") {
            OTel.OTelLogger.onTelemetryLoaded(function () {
                var telemetryData = [
                    oteljs.makeBooleanDataField('Loaded', loaded),
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
    var IdentityType;
    (function (IdentityType) {
        IdentityType["organizationAccount"] = "OrganizationAccount";
        IdentityType["microsoftAccount"] = "MicrosoftAccount";
        IdentityType["unsupported"] = "Unsupported";
    })(IdentityType = OfficeCore.IdentityType || (OfficeCore.IdentityType = {}));
    var _typeAuthenticationService = "AuthenticationService";
    var AuthenticationService = (function (_super) {
        __extends(AuthenticationService, _super);
        function AuthenticationService() {
            return _super !== null && _super.apply(this, arguments) || this;
        }
        Object.defineProperty(AuthenticationService.prototype, "_className", {
            get: function () {
                return "AuthenticationService";
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(AuthenticationService.prototype, "_navigationPropertyNames", {
            get: function () {
                return ["roamingSettings"];
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(AuthenticationService.prototype, "roamingSettings", {
            get: function () {
                if (!this._R) {
                    this._R = _createPropertyObject(OfficeCore.RoamingSettingCollection, this, "RoamingSettings", false, 4);
                }
                return this._R;
            },
            enumerable: true,
            configurable: true
        });
        AuthenticationService.prototype.getAccessToken = function (tokenParameters, targetId) {
            return _invokeMethod(this, "GetAccessToken", 1, [tokenParameters, targetId], 4 | 1, 0);
        };
        AuthenticationService.prototype.getIdentities = function () {
            _throwIfApiNotSupported("AuthenticationService.getIdentities", "FirstPartyAuthentication", "1.3", _hostName);
            return _invokeMethod(this, "GetIdentities", 1, [], 4 | 1, 0);
        };
        AuthenticationService.prototype.getPrimaryIdentityInfo = function () {
            _throwIfApiNotSupported("AuthenticationService.getPrimaryIdentityInfo", "FirstPartyAuthentication", "1.2", _hostName);
            return _invokeMethod(this, "GetPrimaryIdentityInfo", 1, [], 4 | 1, 0);
        };
        AuthenticationService.prototype._handleResult = function (value) {
            _super.prototype._handleResult.call(this, value);
            if (_isNullOrUndefined(value))
                return;
            var obj = value;
            _fixObjectPathIfNecessary(this, obj);
            _handleNavigationPropertyResults(this, obj, ["roamingSettings", "RoamingSettings"]);
        };
        AuthenticationService.prototype.load = function (options) {
            return _load(this, options);
        };
        AuthenticationService.prototype.retrieve = function (option) {
            return _retrieve(this, option);
        };
        AuthenticationService.prototype._handleRetrieveResult = function (value, result) {
            _super.prototype._handleRetrieveResult.call(this, value, result);
            _processRetrieveResult(this, value, result);
        };
        AuthenticationService.newObject = function (context) {
            return _createTopLevelServiceObject(OfficeCore.AuthenticationService, context, "Microsoft.Authentication.AuthenticationService", false, 4);
        };
        Object.defineProperty(AuthenticationService.prototype, "onTokenReceived", {
            get: function () {
                var _this = this;
                _throwIfApiNotSupported("AuthenticationService.onTokenReceived", "FirstPartyAuthentication", "1.2", _hostName);
                if (!this.m_tokenReceived) {
                    this.m_tokenReceived = new OfficeExtension.GenericEventHandlers(this.context, this, "TokenReceived", {
                        eventType: 3001,
                        registerFunc: function () { return OfficeExtension.Utility._createPromiseFromResult(null); },
                        unregisterFunc: function () { return OfficeExtension.Utility._createPromiseFromResult(null); },
                        getTargetIdFunc: function () { return _this._targetId; },
                        eventArgsTransformFunc: function (value) {
                            var event = _CC.AuthenticationService_TokenReceived_EventArgsTransform(_this, value);
                            return OfficeExtension.Utility._createPromiseFromResult(event);
                        }
                    });
                }
                return this.m_tokenReceived;
            },
            enumerable: true,
            configurable: true
        });
        AuthenticationService.prototype.toJSON = function () {
            return _toJson(this, {}, {});
        };
        return AuthenticationService;
    }(OfficeExtension.ClientObject));
    OfficeCore.AuthenticationService = AuthenticationService;
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
    OfficeExtension.Utility.applyMixin(AuthenticationService, AuthenticationServiceCustom);
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
    var _typeRoamingSetting = "RoamingSetting";
    var RoamingSetting = (function (_super) {
        __extends(RoamingSetting, _super);
        function RoamingSetting() {
            return _super !== null && _super.apply(this, arguments) || this;
        }
        Object.defineProperty(RoamingSetting.prototype, "_className", {
            get: function () {
                return "RoamingSetting";
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(RoamingSetting.prototype, "_scalarPropertyNames", {
            get: function () {
                return ["id", "value"];
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(RoamingSetting.prototype, "_scalarPropertyOriginalNames", {
            get: function () {
                return ["Id", "Value"];
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(RoamingSetting.prototype, "_scalarPropertyUpdateable", {
            get: function () {
                return [false, true];
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(RoamingSetting.prototype, "id", {
            get: function () {
                _throwIfNotLoaded("id", this._I, _typeRoamingSetting, this._isNull);
                return this._I;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(RoamingSetting.prototype, "value", {
            get: function () {
                _throwIfNotLoaded("value", this._V, _typeRoamingSetting, this._isNull);
                return this._V;
            },
            set: function (value) {
                this._V = value;
                _invokeSetProperty(this, "Value", value, 0);
            },
            enumerable: true,
            configurable: true
        });
        RoamingSetting.prototype.set = function (properties, options) {
            this._recursivelySet(properties, options, ["value"], [], []);
        };
        RoamingSetting.prototype.update = function (properties) {
            this._recursivelyUpdate(properties);
        };
        RoamingSetting.prototype._handleResult = function (value) {
            _super.prototype._handleResult.call(this, value);
            if (_isNullOrUndefined(value))
                return;
            var obj = value;
            _fixObjectPathIfNecessary(this, obj);
            if (!_isUndefined(obj["Id"])) {
                this._I = obj["Id"];
            }
            if (!_isUndefined(obj["Value"])) {
                this._V = obj["Value"];
            }
        };
        RoamingSetting.prototype.load = function (options) {
            return _load(this, options);
        };
        RoamingSetting.prototype.retrieve = function (option) {
            return _retrieve(this, option);
        };
        RoamingSetting.prototype._handleIdResult = function (value) {
            _super.prototype._handleIdResult.call(this, value);
            if (_isNullOrUndefined(value)) {
                return;
            }
            if (!_isUndefined(value["Id"])) {
                this._I = value["Id"];
            }
        };
        RoamingSetting.prototype._handleRetrieveResult = function (value, result) {
            _super.prototype._handleRetrieveResult.call(this, value, result);
            _processRetrieveResult(this, value, result);
        };
        RoamingSetting.prototype.toJSON = function () {
            return _toJson(this, {
                "id": this._I,
                "value": this._V
            }, {});
        };
        RoamingSetting.prototype.setMockData = function (data) {
            _setMockData(this, data);
        };
        RoamingSetting.prototype.ensureUnchanged = function (data) {
            _invokeEnsureUnchanged(this, data);
            return;
        };
        return RoamingSetting;
    }(OfficeExtension.ClientObject));
    OfficeCore.RoamingSetting = RoamingSetting;
    var _typeRoamingSettingCollection = "RoamingSettingCollection";
    var RoamingSettingCollection = (function (_super) {
        __extends(RoamingSettingCollection, _super);
        function RoamingSettingCollection() {
            return _super !== null && _super.apply(this, arguments) || this;
        }
        Object.defineProperty(RoamingSettingCollection.prototype, "_className", {
            get: function () {
                return "RoamingSettingCollection";
            },
            enumerable: true,
            configurable: true
        });
        RoamingSettingCollection.prototype.getItem = function (id) {
            return _createMethodObject(OfficeCore.RoamingSetting, this, "GetItem", 1, [id], false, false, null, 4);
        };
        RoamingSettingCollection.prototype.getItemOrNullObject = function (id) {
            return _createMethodObject(OfficeCore.RoamingSetting, this, "GetItemOrNullObject", 1, [id], false, false, null, 4);
        };
        RoamingSettingCollection.prototype._handleResult = function (value) {
            _super.prototype._handleResult.call(this, value);
            if (_isNullOrUndefined(value))
                return;
            var obj = value;
            _fixObjectPathIfNecessary(this, obj);
        };
        RoamingSettingCollection.prototype._handleRetrieveResult = function (value, result) {
            _super.prototype._handleRetrieveResult.call(this, value, result);
            _processRetrieveResult(this, value, result);
        };
        RoamingSettingCollection.prototype.toJSON = function () {
            return _toJson(this, {}, {});
        };
        return RoamingSettingCollection;
    }(OfficeExtension.ClientObject));
    OfficeCore.RoamingSettingCollection = RoamingSettingCollection;
    var ServiceProvider;
    (function (ServiceProvider) {
        ServiceProvider["ariaBrowserPipeUrl"] = "AriaBrowserPipeUrl";
        ServiceProvider["ariaUploadUrl"] = "AriaUploadUrl";
        ServiceProvider["ariaVNextUploadUrl"] = "AriaVNextUploadUrl";
        ServiceProvider["lokiAutoDiscoverUrl"] = "LokiAutoDiscoverUrl";
    })(ServiceProvider = OfficeCore.ServiceProvider || (OfficeCore.ServiceProvider = {}));
    var _typeServiceUrlProvider = "ServiceUrlProvider";
    var ServiceUrlProvider = (function (_super) {
        __extends(ServiceUrlProvider, _super);
        function ServiceUrlProvider() {
            return _super !== null && _super.apply(this, arguments) || this;
        }
        Object.defineProperty(ServiceUrlProvider.prototype, "_className", {
            get: function () {
                return "ServiceUrlProvider";
            },
            enumerable: true,
            configurable: true
        });
        ServiceUrlProvider.prototype.getServiceUrl = function (emailAddress, provider) {
            return _invokeMethod(this, "GetServiceUrl", 1, [emailAddress, provider], 4, 0);
        };
        ServiceUrlProvider.prototype._handleResult = function (value) {
            _super.prototype._handleResult.call(this, value);
            if (_isNullOrUndefined(value))
                return;
            var obj = value;
            _fixObjectPathIfNecessary(this, obj);
        };
        ServiceUrlProvider.prototype._handleRetrieveResult = function (value, result) {
            _super.prototype._handleRetrieveResult.call(this, value, result);
            _processRetrieveResult(this, value, result);
        };
        ServiceUrlProvider.newObject = function (context) {
            return _createTopLevelServiceObject(OfficeCore.ServiceUrlProvider, context, "Microsoft.DesktopCompliance.ServiceUrlProvider", false, 4);
        };
        ServiceUrlProvider.prototype.toJSON = function () {
            return _toJson(this, {}, {});
        };
        return ServiceUrlProvider;
    }(OfficeExtension.ClientObject));
    OfficeCore.ServiceUrlProvider = ServiceUrlProvider;
    var _typeLinkedIn = "LinkedIn";
    var LinkedIn = (function (_super) {
        __extends(LinkedIn, _super);
        function LinkedIn() {
            return _super !== null && _super.apply(this, arguments) || this;
        }
        Object.defineProperty(LinkedIn.prototype, "_className", {
            get: function () {
                return "LinkedIn";
            },
            enumerable: true,
            configurable: true
        });
        LinkedIn.prototype.isEnabledForOffice = function () {
            return _invokeMethod(this, "IsEnabledForOffice", 1, [], 4, 0);
        };
        LinkedIn.prototype.recordLinkedInSettingsCompliance = function (featureName, isEnabled) {
            _invokeMethod(this, "RecordLinkedInSettingsCompliance", 0, [featureName, isEnabled], 0, 0);
        };
        LinkedIn.prototype._handleResult = function (value) {
            _super.prototype._handleResult.call(this, value);
            if (_isNullOrUndefined(value))
                return;
            var obj = value;
            _fixObjectPathIfNecessary(this, obj);
        };
        LinkedIn.prototype._handleRetrieveResult = function (value, result) {
            _super.prototype._handleRetrieveResult.call(this, value, result);
            _processRetrieveResult(this, value, result);
        };
        LinkedIn.newObject = function (context) {
            return _createTopLevelServiceObject(OfficeCore.LinkedIn, context, "Microsoft.DesktopCompliance.LinkedIn", false, 4);
        };
        LinkedIn.prototype.toJSON = function () {
            return _toJson(this, {}, {});
        };
        return LinkedIn;
    }(OfficeExtension.ClientObject));
    OfficeCore.LinkedIn = LinkedIn;
    var _typeNetworkUsage = "NetworkUsage";
    var NetworkUsage = (function (_super) {
        __extends(NetworkUsage, _super);
        function NetworkUsage() {
            return _super !== null && _super.apply(this, arguments) || this;
        }
        Object.defineProperty(NetworkUsage.prototype, "_className", {
            get: function () {
                return "NetworkUsage";
            },
            enumerable: true,
            configurable: true
        });
        NetworkUsage.prototype.isInDisconnectedMode = function () {
            return _invokeMethod(this, "IsInDisconnectedMode", 1, [], 4, 0);
        };
        NetworkUsage.prototype.isInOnlineMode = function () {
            return _invokeMethod(this, "IsInOnlineMode", 1, [], 4, 0);
        };
        NetworkUsage.prototype._handleResult = function (value) {
            _super.prototype._handleResult.call(this, value);
            if (_isNullOrUndefined(value))
                return;
            var obj = value;
            _fixObjectPathIfNecessary(this, obj);
        };
        NetworkUsage.prototype._handleRetrieveResult = function (value, result) {
            _super.prototype._handleRetrieveResult.call(this, value, result);
            _processRetrieveResult(this, value, result);
        };
        NetworkUsage.newObject = function (context) {
            return _createTopLevelServiceObject(OfficeCore.NetworkUsage, context, "Microsoft.DesktopCompliance.NetworkUsage", false, 4);
        };
        NetworkUsage.prototype.toJSON = function () {
            return _toJson(this, {}, {});
        };
        return NetworkUsage;
    }(OfficeExtension.ClientObject));
    OfficeCore.NetworkUsage = NetworkUsage;
    var _typeDynamicRibbon = "DynamicRibbon";
    var DynamicRibbon = (function (_super) {
        __extends(DynamicRibbon, _super);
        function DynamicRibbon() {
            return _super !== null && _super.apply(this, arguments) || this;
        }
        Object.defineProperty(DynamicRibbon.prototype, "_className", {
            get: function () {
                return "DynamicRibbon";
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(DynamicRibbon.prototype, "_navigationPropertyNames", {
            get: function () {
                return ["buttons"];
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(DynamicRibbon.prototype, "buttons", {
            get: function () {
                if (!this._B) {
                    this._B = _createPropertyObject(OfficeCore.RibbonButtonCollection, this, "Buttons", true, 4);
                }
                return this._B;
            },
            enumerable: true,
            configurable: true
        });
        DynamicRibbon.prototype.executeRequestCreate = function (jsonCreate) {
            _throwIfApiNotSupported("DynamicRibbon.executeRequestCreate", "DynamicRibbon", "1.2", _hostName);
            _invokeMethod(this, "ExecuteRequestCreate", 1, [jsonCreate], 4, 0);
        };
        DynamicRibbon.prototype.executeRequestUpdate = function (jsonUpdate) {
            _invokeMethod(this, "ExecuteRequestUpdate", 1, [jsonUpdate], 4, 0);
        };
        DynamicRibbon.prototype.getButton = function (id) {
            return _createMethodObject(OfficeCore.RibbonButton, this, "GetButton", 1, [id], false, false, null, 4);
        };
        DynamicRibbon.prototype.getTab = function (id) {
            return _createMethodObject(OfficeCore.RibbonTab, this, "GetTab", 1, [id], false, false, null, 4);
        };
        DynamicRibbon.prototype._handleResult = function (value) {
            _super.prototype._handleResult.call(this, value);
            if (_isNullOrUndefined(value))
                return;
            var obj = value;
            _fixObjectPathIfNecessary(this, obj);
            _handleNavigationPropertyResults(this, obj, ["buttons", "Buttons"]);
        };
        DynamicRibbon.prototype.load = function (options) {
            return _load(this, options);
        };
        DynamicRibbon.prototype.retrieve = function (option) {
            return _retrieve(this, option);
        };
        DynamicRibbon.prototype._handleRetrieveResult = function (value, result) {
            _super.prototype._handleRetrieveResult.call(this, value, result);
            _processRetrieveResult(this, value, result);
        };
        DynamicRibbon.newObject = function (context) {
            return _createTopLevelServiceObject(OfficeCore.DynamicRibbon, context, "Microsoft.DynamicRibbon.DynamicRibbon", false, 4);
        };
        DynamicRibbon.prototype.toJSON = function () {
            return _toJson(this, {}, {
                "buttons": this._B
            });
        };
        return DynamicRibbon;
    }(OfficeExtension.ClientObject));
    OfficeCore.DynamicRibbon = DynamicRibbon;
    var _typeRibbonTab = "RibbonTab";
    var RibbonTab = (function (_super) {
        __extends(RibbonTab, _super);
        function RibbonTab() {
            return _super !== null && _super.apply(this, arguments) || this;
        }
        Object.defineProperty(RibbonTab.prototype, "_className", {
            get: function () {
                return "RibbonTab";
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(RibbonTab.prototype, "_scalarPropertyNames", {
            get: function () {
                return ["id"];
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(RibbonTab.prototype, "_scalarPropertyOriginalNames", {
            get: function () {
                return ["Id"];
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(RibbonTab.prototype, "id", {
            get: function () {
                _throwIfNotLoaded("id", this._I, _typeRibbonTab, this._isNull);
                return this._I;
            },
            enumerable: true,
            configurable: true
        });
        RibbonTab.prototype.setVisibility = function (visibility) {
            _invokeMethod(this, "SetVisibility", 0, [visibility], 0, 0);
        };
        RibbonTab.prototype._handleResult = function (value) {
            _super.prototype._handleResult.call(this, value);
            if (_isNullOrUndefined(value))
                return;
            var obj = value;
            _fixObjectPathIfNecessary(this, obj);
            if (!_isUndefined(obj["Id"])) {
                this._I = obj["Id"];
            }
        };
        RibbonTab.prototype.load = function (options) {
            return _load(this, options);
        };
        RibbonTab.prototype.retrieve = function (option) {
            return _retrieve(this, option);
        };
        RibbonTab.prototype._handleIdResult = function (value) {
            _super.prototype._handleIdResult.call(this, value);
            if (_isNullOrUndefined(value)) {
                return;
            }
            if (!_isUndefined(value["Id"])) {
                this._I = value["Id"];
            }
        };
        RibbonTab.prototype._handleRetrieveResult = function (value, result) {
            _super.prototype._handleRetrieveResult.call(this, value, result);
            _processRetrieveResult(this, value, result);
        };
        RibbonTab.prototype.toJSON = function () {
            return _toJson(this, {
                "id": this._I
            }, {});
        };
        RibbonTab.prototype.setMockData = function (data) {
            _setMockData(this, data);
        };
        RibbonTab.prototype.ensureUnchanged = function (data) {
            _invokeEnsureUnchanged(this, data);
            return;
        };
        return RibbonTab;
    }(OfficeExtension.ClientObject));
    OfficeCore.RibbonTab = RibbonTab;
    var _typeRibbonButton = "RibbonButton";
    var RibbonButton = (function (_super) {
        __extends(RibbonButton, _super);
        function RibbonButton() {
            return _super !== null && _super.apply(this, arguments) || this;
        }
        Object.defineProperty(RibbonButton.prototype, "_className", {
            get: function () {
                return "RibbonButton";
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(RibbonButton.prototype, "_scalarPropertyNames", {
            get: function () {
                return ["id", "enabled", "label"];
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(RibbonButton.prototype, "_scalarPropertyOriginalNames", {
            get: function () {
                return ["Id", "Enabled", "Label"];
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(RibbonButton.prototype, "_scalarPropertyUpdateable", {
            get: function () {
                return [false, true, false];
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(RibbonButton.prototype, "enabled", {
            get: function () {
                _throwIfNotLoaded("enabled", this._E, _typeRibbonButton, this._isNull);
                return this._E;
            },
            set: function (value) {
                this._E = value;
                _invokeSetProperty(this, "Enabled", value, 0);
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(RibbonButton.prototype, "id", {
            get: function () {
                _throwIfNotLoaded("id", this._I, _typeRibbonButton, this._isNull);
                return this._I;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(RibbonButton.prototype, "label", {
            get: function () {
                _throwIfNotLoaded("label", this._L, _typeRibbonButton, this._isNull);
                return this._L;
            },
            enumerable: true,
            configurable: true
        });
        RibbonButton.prototype.set = function (properties, options) {
            this._recursivelySet(properties, options, ["enabled"], [], []);
        };
        RibbonButton.prototype.update = function (properties) {
            this._recursivelyUpdate(properties);
        };
        RibbonButton.prototype.setEnabled = function (enabled) {
            _invokeMethod(this, "SetEnabled", 0, [enabled], 0, 0);
        };
        RibbonButton.prototype._handleResult = function (value) {
            _super.prototype._handleResult.call(this, value);
            if (_isNullOrUndefined(value))
                return;
            var obj = value;
            _fixObjectPathIfNecessary(this, obj);
            if (!_isUndefined(obj["Enabled"])) {
                this._E = obj["Enabled"];
            }
            if (!_isUndefined(obj["Id"])) {
                this._I = obj["Id"];
            }
            if (!_isUndefined(obj["Label"])) {
                this._L = obj["Label"];
            }
        };
        RibbonButton.prototype.load = function (options) {
            return _load(this, options);
        };
        RibbonButton.prototype.retrieve = function (option) {
            return _retrieve(this, option);
        };
        RibbonButton.prototype._handleIdResult = function (value) {
            _super.prototype._handleIdResult.call(this, value);
            if (_isNullOrUndefined(value)) {
                return;
            }
            if (!_isUndefined(value["Id"])) {
                this._I = value["Id"];
            }
        };
        RibbonButton.prototype._handleRetrieveResult = function (value, result) {
            _super.prototype._handleRetrieveResult.call(this, value, result);
            _processRetrieveResult(this, value, result);
        };
        RibbonButton.prototype.toJSON = function () {
            return _toJson(this, {
                "enabled": this._E,
                "id": this._I,
                "label": this._L
            }, {});
        };
        RibbonButton.prototype.setMockData = function (data) {
            _setMockData(this, data);
        };
        RibbonButton.prototype.ensureUnchanged = function (data) {
            _invokeEnsureUnchanged(this, data);
            return;
        };
        return RibbonButton;
    }(OfficeExtension.ClientObject));
    OfficeCore.RibbonButton = RibbonButton;
    var _typeRibbonButtonCollection = "RibbonButtonCollection";
    var RibbonButtonCollection = (function (_super) {
        __extends(RibbonButtonCollection, _super);
        function RibbonButtonCollection() {
            return _super !== null && _super.apply(this, arguments) || this;
        }
        Object.defineProperty(RibbonButtonCollection.prototype, "_className", {
            get: function () {
                return "RibbonButtonCollection";
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(RibbonButtonCollection.prototype, "_isCollection", {
            get: function () {
                return true;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(RibbonButtonCollection.prototype, "items", {
            get: function () {
                _throwIfNotLoaded("items", this.m__items, _typeRibbonButtonCollection, this._isNull);
                return this.m__items;
            },
            enumerable: true,
            configurable: true
        });
        RibbonButtonCollection.prototype.getCount = function () {
            return _invokeMethod(this, "GetCount", 1, [], 4, 0);
        };
        RibbonButtonCollection.prototype.getItem = function (key) {
            return _createIndexerObject(OfficeCore.RibbonButton, this, [key]);
        };
        RibbonButtonCollection.prototype._handleResult = function (value) {
            _super.prototype._handleResult.call(this, value);
            if (_isNullOrUndefined(value))
                return;
            var obj = value;
            _fixObjectPathIfNecessary(this, obj);
            if (!_isNullOrUndefined(obj[OfficeExtension.Constants.items])) {
                this.m__items = [];
                var _data = obj[OfficeExtension.Constants.items];
                for (var i = 0; i < _data.length; i++) {
                    var _item = _createChildItemObject(OfficeCore.RibbonButton, true, this, _data[i], i);
                    _item._handleResult(_data[i]);
                    this.m__items.push(_item);
                }
            }
        };
        RibbonButtonCollection.prototype.load = function (options) {
            return _load(this, options);
        };
        RibbonButtonCollection.prototype.retrieve = function (option) {
            return _retrieve(this, option);
        };
        RibbonButtonCollection.prototype._handleRetrieveResult = function (value, result) {
            var _this = this;
            _super.prototype._handleRetrieveResult.call(this, value, result);
            _processRetrieveResult(this, value, result, function (childItemData, index) { return _createChildItemObject(OfficeCore.RibbonButton, true, _this, childItemData, index); });
        };
        RibbonButtonCollection.prototype.toJSON = function () {
            return _toJson(this, {}, {}, this.m__items);
        };
        RibbonButtonCollection.prototype.setMockData = function (data) {
            var _this = this;
            _setMockData(this, data, function (childItemData, index) { return _createChildItemObject(OfficeCore.RibbonButton, true, _this, childItemData, index); }, function (items) { return _this.m__items = items; });
        };
        return RibbonButtonCollection;
    }(OfficeExtension.ClientObject));
    OfficeCore.RibbonButtonCollection = RibbonButtonCollection;
    var TimeStringFormat;
    (function (TimeStringFormat) {
        TimeStringFormat["shortTime"] = "ShortTime";
        TimeStringFormat["longTime"] = "LongTime";
        TimeStringFormat["shortDate"] = "ShortDate";
        TimeStringFormat["longDate"] = "LongDate";
    })(TimeStringFormat = OfficeCore.TimeStringFormat || (OfficeCore.TimeStringFormat = {}));
    var _typeLocaleApi = "LocaleApi";
    var LocaleApi = (function (_super) {
        __extends(LocaleApi, _super);
        function LocaleApi() {
            return _super !== null && _super.apply(this, arguments) || this;
        }
        Object.defineProperty(LocaleApi.prototype, "_className", {
            get: function () {
                return "LocaleApi";
            },
            enumerable: true,
            configurable: true
        });
        LocaleApi.prototype.formatDateTimeString = function (localeName, value, format) {
            return _invokeMethod(this, "FormatDateTimeString", 1, [localeName, value, format], 4, 0);
        };
        LocaleApi.prototype.getLocaleDateTimeFormattingInfo = function (localeName) {
            return _invokeMethod(this, "GetLocaleDateTimeFormattingInfo", 1, [localeName], 4, 0);
        };
        LocaleApi.prototype._handleResult = function (value) {
            _super.prototype._handleResult.call(this, value);
            if (_isNullOrUndefined(value))
                return;
            var obj = value;
            _fixObjectPathIfNecessary(this, obj);
        };
        LocaleApi.prototype._handleRetrieveResult = function (value, result) {
            _super.prototype._handleRetrieveResult.call(this, value, result);
            _processRetrieveResult(this, value, result);
        };
        LocaleApi.newObject = function (context) {
            return _createTopLevelServiceObject(OfficeCore.LocaleApi, context, "Microsoft.LocaleApi.LocaleApi", false, 4);
        };
        LocaleApi.prototype.toJSON = function () {
            return _toJson(this, {}, {});
        };
        return LocaleApi;
    }(OfficeExtension.ClientObject));
    OfficeCore.LocaleApi = LocaleApi;
    var _typeComment = "Comment";
    var Comment = (function (_super) {
        __extends(Comment, _super);
        function Comment() {
            return _super !== null && _super.apply(this, arguments) || this;
        }
        Object.defineProperty(Comment.prototype, "_className", {
            get: function () {
                return "Comment";
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(Comment.prototype, "_scalarPropertyNames", {
            get: function () {
                return ["id", "text", "created", "level", "resolved", "author", "mentions"];
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(Comment.prototype, "_scalarPropertyOriginalNames", {
            get: function () {
                return ["Id", "Text", "Created", "Level", "Resolved", "Author", "Mentions"];
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(Comment.prototype, "_scalarPropertyUpdateable", {
            get: function () {
                return [false, true, false, false, true, false, false];
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(Comment.prototype, "_navigationPropertyNames", {
            get: function () {
                return ["parent", "parentOrNullObject", "replies"];
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(Comment.prototype, "parent", {
            get: function () {
                if (!this._P) {
                    this._P = _createPropertyObject(OfficeCore.Comment, this, "Parent", false, 4);
                }
                return this._P;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(Comment.prototype, "parentOrNullObject", {
            get: function () {
                if (!this._Pa) {
                    this._Pa = _createPropertyObject(OfficeCore.Comment, this, "ParentOrNullObject", false, 4);
                }
                return this._Pa;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(Comment.prototype, "replies", {
            get: function () {
                if (!this._R) {
                    this._R = _createPropertyObject(OfficeCore.CommentCollection, this, "Replies", true, 4);
                }
                return this._R;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(Comment.prototype, "author", {
            get: function () {
                _throwIfNotLoaded("author", this._A, _typeComment, this._isNull);
                return this._A;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(Comment.prototype, "created", {
            get: function () {
                _throwIfNotLoaded("created", this._C, _typeComment, this._isNull);
                return this._C;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(Comment.prototype, "id", {
            get: function () {
                _throwIfNotLoaded("id", this._I, _typeComment, this._isNull);
                return this._I;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(Comment.prototype, "level", {
            get: function () {
                _throwIfNotLoaded("level", this._L, _typeComment, this._isNull);
                return this._L;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(Comment.prototype, "mentions", {
            get: function () {
                _throwIfNotLoaded("mentions", this._M, _typeComment, this._isNull);
                return this._M;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(Comment.prototype, "resolved", {
            get: function () {
                _throwIfNotLoaded("resolved", this._Re, _typeComment, this._isNull);
                return this._Re;
            },
            set: function (value) {
                this._Re = value;
                _invokeSetProperty(this, "Resolved", value, 0);
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(Comment.prototype, "text", {
            get: function () {
                _throwIfNotLoaded("text", this._T, _typeComment, this._isNull);
                return this._T;
            },
            set: function (value) {
                this._T = value;
                _invokeSetProperty(this, "Text", value, 0);
            },
            enumerable: true,
            configurable: true
        });
        Comment.prototype.set = function (properties, options) {
            this._recursivelySet(properties, options, ["text", "resolved"], [], [
                "parent",
                "parentOrNullObject",
                "replies"
            ]);
        };
        Comment.prototype.update = function (properties) {
            this._recursivelyUpdate(properties);
        };
        Comment.prototype["delete"] = function () {
            _invokeMethod(this, "Delete", 0, [], 0, 0);
        };
        Comment.prototype.getParentOrSelf = function () {
            return _createMethodObject(OfficeCore.Comment, this, "GetParentOrSelf", 1, [], false, false, null, 4);
        };
        Comment.prototype.getRichText = function (format) {
            return _invokeMethod(this, "GetRichText", 1, [format], 4, 0);
        };
        Comment.prototype.reply = function (text, format) {
            return _createMethodObject(OfficeCore.Comment, this, "Reply", 0, [text, format], false, false, null, 0);
        };
        Comment.prototype.setRichText = function (text, format) {
            return _invokeMethod(this, "SetRichText", 0, [text, format], 0, 0);
        };
        Comment.prototype._handleResult = function (value) {
            _super.prototype._handleResult.call(this, value);
            if (_isNullOrUndefined(value))
                return;
            var obj = value;
            _fixObjectPathIfNecessary(this, obj);
            if (!_isUndefined(obj["Author"])) {
                this._A = obj["Author"];
            }
            if (!_isUndefined(obj["Created"])) {
                this._C = _adjustToDateTime(obj["Created"]);
            }
            if (!_isUndefined(obj["Id"])) {
                this._I = obj["Id"];
            }
            if (!_isUndefined(obj["Level"])) {
                this._L = obj["Level"];
            }
            if (!_isUndefined(obj["Mentions"])) {
                this._M = obj["Mentions"];
            }
            if (!_isUndefined(obj["Resolved"])) {
                this._Re = obj["Resolved"];
            }
            if (!_isUndefined(obj["Text"])) {
                this._T = obj["Text"];
            }
            _handleNavigationPropertyResults(this, obj, ["parent", "Parent", "parentOrNullObject", "ParentOrNullObject", "replies", "Replies"]);
        };
        Comment.prototype.load = function (options) {
            return _load(this, options);
        };
        Comment.prototype.retrieve = function (option) {
            return _retrieve(this, option);
        };
        Comment.prototype._handleIdResult = function (value) {
            _super.prototype._handleIdResult.call(this, value);
            if (_isNullOrUndefined(value)) {
                return;
            }
            if (!_isUndefined(value["Id"])) {
                this._I = value["Id"];
            }
        };
        Comment.prototype._handleRetrieveResult = function (value, result) {
            _super.prototype._handleRetrieveResult.call(this, value, result);
            if (_isNullOrUndefined(value))
                return;
            var obj = value;
            if (!_isUndefined(obj["Created"])) {
                obj["created"] = _adjustToDateTime(obj["created"]);
            }
            _processRetrieveResult(this, value, result);
        };
        Comment.prototype.toJSON = function () {
            return _toJson(this, {
                "author": this._A,
                "created": this._C,
                "id": this._I,
                "level": this._L,
                "mentions": this._M,
                "resolved": this._Re,
                "text": this._T
            }, {
                "replies": this._R
            });
        };
        Comment.prototype.setMockData = function (data) {
            _setMockData(this, data);
        };
        Comment.prototype.ensureUnchanged = function (data) {
            _invokeEnsureUnchanged(this, data);
            return;
        };
        return Comment;
    }(OfficeExtension.ClientObject));
    OfficeCore.Comment = Comment;
    var _typeCommentCollection = "CommentCollection";
    var CommentCollection = (function (_super) {
        __extends(CommentCollection, _super);
        function CommentCollection() {
            return _super !== null && _super.apply(this, arguments) || this;
        }
        Object.defineProperty(CommentCollection.prototype, "_className", {
            get: function () {
                return "CommentCollection";
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(CommentCollection.prototype, "_isCollection", {
            get: function () {
                return true;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(CommentCollection.prototype, "items", {
            get: function () {
                _throwIfNotLoaded("items", this.m__items, _typeCommentCollection, this._isNull);
                return this.m__items;
            },
            enumerable: true,
            configurable: true
        });
        CommentCollection.prototype.getCount = function () {
            return _invokeMethod(this, "GetCount", 1, [], 4, 0);
        };
        CommentCollection.prototype.getItem = function (id) {
            return _createIndexerObject(OfficeCore.Comment, this, [id]);
        };
        CommentCollection.prototype._handleResult = function (value) {
            _super.prototype._handleResult.call(this, value);
            if (_isNullOrUndefined(value))
                return;
            var obj = value;
            _fixObjectPathIfNecessary(this, obj);
            if (!_isNullOrUndefined(obj[OfficeExtension.Constants.items])) {
                this.m__items = [];
                var _data = obj[OfficeExtension.Constants.items];
                for (var i = 0; i < _data.length; i++) {
                    var _item = _createChildItemObject(OfficeCore.Comment, true, this, _data[i], i);
                    _item._handleResult(_data[i]);
                    this.m__items.push(_item);
                }
            }
        };
        CommentCollection.prototype.load = function (options) {
            return _load(this, options);
        };
        CommentCollection.prototype.retrieve = function (option) {
            return _retrieve(this, option);
        };
        CommentCollection.prototype._handleRetrieveResult = function (value, result) {
            var _this = this;
            _super.prototype._handleRetrieveResult.call(this, value, result);
            _processRetrieveResult(this, value, result, function (childItemData, index) { return _createChildItemObject(OfficeCore.Comment, true, _this, childItemData, index); });
        };
        CommentCollection.prototype.toJSON = function () {
            return _toJson(this, {}, {}, this.m__items);
        };
        CommentCollection.prototype.setMockData = function (data) {
            var _this = this;
            _setMockData(this, data, function (childItemData, index) { return _createChildItemObject(OfficeCore.Comment, true, _this, childItemData, index); }, function (items) { return _this.m__items = items; });
        };
        return CommentCollection;
    }(OfficeExtension.ClientObject));
    OfficeCore.CommentCollection = CommentCollection;
    var CommentTextFormat;
    (function (CommentTextFormat) {
        CommentTextFormat["plain"] = "Plain";
        CommentTextFormat["markdown"] = "Markdown";
        CommentTextFormat["delta"] = "Delta";
    })(CommentTextFormat = OfficeCore.CommentTextFormat || (OfficeCore.CommentTextFormat = {}));
    var PersonaCardPerfPoint;
    (function (PersonaCardPerfPoint) {
        PersonaCardPerfPoint["placeHolderRendered"] = "PlaceHolderRendered";
        PersonaCardPerfPoint["initialCardRendered"] = "InitialCardRendered";
    })(PersonaCardPerfPoint = OfficeCore.PersonaCardPerfPoint || (OfficeCore.PersonaCardPerfPoint = {}));
    var UnifiedCommunicationAvailability;
    (function (UnifiedCommunicationAvailability) {
        UnifiedCommunicationAvailability["notSet"] = "NotSet";
        UnifiedCommunicationAvailability["free"] = "Free";
        UnifiedCommunicationAvailability["idle"] = "Idle";
        UnifiedCommunicationAvailability["busy"] = "Busy";
        UnifiedCommunicationAvailability["idleBusy"] = "IdleBusy";
        UnifiedCommunicationAvailability["doNotDisturb"] = "DoNotDisturb";
        UnifiedCommunicationAvailability["unalertable"] = "Unalertable";
        UnifiedCommunicationAvailability["unavailable"] = "Unavailable";
    })(UnifiedCommunicationAvailability = OfficeCore.UnifiedCommunicationAvailability || (OfficeCore.UnifiedCommunicationAvailability = {}));
    var UnifiedCommunicationStatus;
    (function (UnifiedCommunicationStatus) {
        UnifiedCommunicationStatus["online"] = "Online";
        UnifiedCommunicationStatus["notOnline"] = "NotOnline";
        UnifiedCommunicationStatus["away"] = "Away";
        UnifiedCommunicationStatus["busy"] = "Busy";
        UnifiedCommunicationStatus["beRightBack"] = "BeRightBack";
        UnifiedCommunicationStatus["onThePhone"] = "OnThePhone";
        UnifiedCommunicationStatus["outToLunch"] = "OutToLunch";
        UnifiedCommunicationStatus["inAMeeting"] = "InAMeeting";
        UnifiedCommunicationStatus["outOfOffice"] = "OutOfOffice";
        UnifiedCommunicationStatus["doNotDisturb"] = "DoNotDisturb";
        UnifiedCommunicationStatus["inAConference"] = "InAConference";
        UnifiedCommunicationStatus["getting"] = "Getting";
        UnifiedCommunicationStatus["notABuddy"] = "NotABuddy";
        UnifiedCommunicationStatus["disconnected"] = "Disconnected";
        UnifiedCommunicationStatus["notInstalled"] = "NotInstalled";
        UnifiedCommunicationStatus["urgentInterruptionsOnly"] = "UrgentInterruptionsOnly";
        UnifiedCommunicationStatus["mayBeAvailable"] = "MayBeAvailable";
        UnifiedCommunicationStatus["idle"] = "Idle";
        UnifiedCommunicationStatus["inPresentation"] = "InPresentation";
    })(UnifiedCommunicationStatus = OfficeCore.UnifiedCommunicationStatus || (OfficeCore.UnifiedCommunicationStatus = {}));
    var UnifiedCommunicationPresence;
    (function (UnifiedCommunicationPresence) {
        UnifiedCommunicationPresence["free"] = "Free";
        UnifiedCommunicationPresence["busy"] = "Busy";
        UnifiedCommunicationPresence["idle"] = "Idle";
        UnifiedCommunicationPresence["doNotDistrub"] = "DoNotDistrub";
        UnifiedCommunicationPresence["blocked"] = "Blocked";
        UnifiedCommunicationPresence["notSet"] = "NotSet";
        UnifiedCommunicationPresence["outOfOffice"] = "OutOfOffice";
    })(UnifiedCommunicationPresence = OfficeCore.UnifiedCommunicationPresence || (OfficeCore.UnifiedCommunicationPresence = {}));
    var FreeBusyCalendarState;
    (function (FreeBusyCalendarState) {
        FreeBusyCalendarState["unknown"] = "Unknown";
        FreeBusyCalendarState["free"] = "Free";
        FreeBusyCalendarState["busy"] = "Busy";
        FreeBusyCalendarState["elsewhere"] = "Elsewhere";
        FreeBusyCalendarState["tentative"] = "Tentative";
        FreeBusyCalendarState["outOfOffice"] = "OutOfOffice";
    })(FreeBusyCalendarState = OfficeCore.FreeBusyCalendarState || (OfficeCore.FreeBusyCalendarState = {}));
    var PersonaType;
    (function (PersonaType) {
        PersonaType["unknown"] = "Unknown";
        PersonaType["enterprise"] = "Enterprise";
        PersonaType["contact"] = "Contact";
        PersonaType["bot"] = "Bot";
        PersonaType["phoneOnly"] = "PhoneOnly";
        PersonaType["oneOff"] = "OneOff";
        PersonaType["distributionList"] = "DistributionList";
        PersonaType["personalDistributionList"] = "PersonalDistributionList";
        PersonaType["anonymous"] = "Anonymous";
        PersonaType["unifiedGroup"] = "UnifiedGroup";
    })(PersonaType = OfficeCore.PersonaType || (OfficeCore.PersonaType = {}));
    var PhoneType;
    (function (PhoneType) {
        PhoneType["workPhone"] = "WorkPhone";
        PhoneType["homePhone"] = "HomePhone";
        PhoneType["mobilePhone"] = "MobilePhone";
        PhoneType["businessFax"] = "BusinessFax";
        PhoneType["otherPhone"] = "OtherPhone";
    })(PhoneType = OfficeCore.PhoneType || (OfficeCore.PhoneType = {}));
    var AddressType;
    (function (AddressType) {
        AddressType["workAddress"] = "WorkAddress";
        AddressType["homeAddress"] = "HomeAddress";
        AddressType["otherAddress"] = "OtherAddress";
    })(AddressType = OfficeCore.AddressType || (OfficeCore.AddressType = {}));
    var MemberType;
    (function (MemberType) {
        MemberType["unknown"] = "Unknown";
        MemberType["individual"] = "Individual";
        MemberType["group"] = "Group";
    })(MemberType = OfficeCore.MemberType || (OfficeCore.MemberType = {}));
    var _typeMemberInfoList = "MemberInfoList";
    var MemberInfoList = (function (_super) {
        __extends(MemberInfoList, _super);
        function MemberInfoList() {
            return _super !== null && _super.apply(this, arguments) || this;
        }
        Object.defineProperty(MemberInfoList.prototype, "_className", {
            get: function () {
                return "MemberInfoList";
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(MemberInfoList.prototype, "_scalarPropertyNames", {
            get: function () {
                return ["isWarmedUp", "isWarmingUp"];
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(MemberInfoList.prototype, "_scalarPropertyOriginalNames", {
            get: function () {
                return ["IsWarmedUp", "IsWarmingUp"];
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(MemberInfoList.prototype, "isWarmedUp", {
            get: function () {
                _throwIfNotLoaded("isWarmedUp", this._I, _typeMemberInfoList, this._isNull);
                return this._I;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(MemberInfoList.prototype, "isWarmingUp", {
            get: function () {
                _throwIfNotLoaded("isWarmingUp", this._Is, _typeMemberInfoList, this._isNull);
                return this._Is;
            },
            enumerable: true,
            configurable: true
        });
        MemberInfoList.prototype.getPersonaForMember = function (memberCookie) {
            return _createMethodObject(OfficeCore.Persona, this, "GetPersonaForMember", 1, [memberCookie], false, false, null, 4);
        };
        MemberInfoList.prototype.items = function () {
            return _invokeMethod(this, "Items", 1, [], 4, 0);
        };
        MemberInfoList.prototype._handleResult = function (value) {
            _super.prototype._handleResult.call(this, value);
            if (_isNullOrUndefined(value))
                return;
            var obj = value;
            _fixObjectPathIfNecessary(this, obj);
            if (!_isUndefined(obj["IsWarmedUp"])) {
                this._I = obj["IsWarmedUp"];
            }
            if (!_isUndefined(obj["IsWarmingUp"])) {
                this._Is = obj["IsWarmingUp"];
            }
        };
        MemberInfoList.prototype.load = function (options) {
            return _load(this, options);
        };
        MemberInfoList.prototype.retrieve = function (option) {
            return _retrieve(this, option);
        };
        MemberInfoList.prototype._handleRetrieveResult = function (value, result) {
            _super.prototype._handleRetrieveResult.call(this, value, result);
            _processRetrieveResult(this, value, result);
        };
        MemberInfoList.prototype.toJSON = function () {
            return _toJson(this, {
                "isWarmedUp": this._I,
                "isWarmingUp": this._Is
            }, {});
        };
        MemberInfoList.prototype.setMockData = function (data) {
            _setMockData(this, data);
        };
        MemberInfoList.prototype.ensureUnchanged = function (data) {
            _invokeEnsureUnchanged(this, data);
            return;
        };
        return MemberInfoList;
    }(OfficeExtension.ClientObject));
    OfficeCore.MemberInfoList = MemberInfoList;
    var PersonaDataUpdated;
    (function (PersonaDataUpdated) {
        PersonaDataUpdated["hostId"] = "HostId";
        PersonaDataUpdated["type"] = "Type";
        PersonaDataUpdated["photo"] = "Photo";
        PersonaDataUpdated["personaInfo"] = "PersonaInfo";
        PersonaDataUpdated["unifiedCommunicationInfo"] = "UnifiedCommunicationInfo";
        PersonaDataUpdated["organization"] = "Organization";
        PersonaDataUpdated["unifiedGroupInfo"] = "UnifiedGroupInfo";
        PersonaDataUpdated["members"] = "Members";
        PersonaDataUpdated["membership"] = "Membership";
        PersonaDataUpdated["capabilities"] = "Capabilities";
        PersonaDataUpdated["customizations"] = "Customizations";
        PersonaDataUpdated["viewableSources"] = "ViewableSources";
        PersonaDataUpdated["placeholder"] = "Placeholder";
    })(PersonaDataUpdated = OfficeCore.PersonaDataUpdated || (OfficeCore.PersonaDataUpdated = {}));
    var _typePersonaActions = "PersonaActions";
    var PersonaActions = (function (_super) {
        __extends(PersonaActions, _super);
        function PersonaActions() {
            return _super !== null && _super.apply(this, arguments) || this;
        }
        Object.defineProperty(PersonaActions.prototype, "_className", {
            get: function () {
                return "PersonaActions";
            },
            enumerable: true,
            configurable: true
        });
        PersonaActions.prototype.addContact = function () {
            _invokeMethod(this, "AddContact", 0, [], 0, 0);
        };
        PersonaActions.prototype.callPhoneNumber = function (contactNumber) {
            _invokeMethod(this, "CallPhoneNumber", 0, [contactNumber], 0, 0);
        };
        PersonaActions.prototype.composeEmail = function (emailAddress) {
            _invokeMethod(this, "ComposeEmail", 0, [emailAddress], 0, 0);
        };
        PersonaActions.prototype.composeInstantMessage = function (sipAddress) {
            _invokeMethod(this, "ComposeInstantMessage", 0, [sipAddress], 0, 0);
        };
        PersonaActions.prototype.editContact = function () {
            _invokeMethod(this, "EditContact", 0, [], 0, 0);
        };
        PersonaActions.prototype.editContactByIdentifier = function (identifier) {
            _invokeMethod(this, "EditContactByIdentifier", 0, [identifier], 0, 0);
        };
        PersonaActions.prototype.getChangePhotoUrlAndOpenInBrowser = function () {
            _invokeMethod(this, "GetChangePhotoUrlAndOpenInBrowser", 0, [], 0, 0);
        };
        PersonaActions.prototype.hideHoverCardForPersona = function () {
            _invokeMethod(this, "HideHoverCardForPersona", 0, [], 0, 0);
        };
        PersonaActions.prototype.openGroupCalendar = function () {
            _invokeMethod(this, "OpenGroupCalendar", 0, [], 0, 0);
        };
        PersonaActions.prototype.openLinkContactUx = function () {
            _invokeMethod(this, "OpenLinkContactUx", 0, [], 0, 0);
        };
        PersonaActions.prototype.openOutlookProperties = function () {
            _invokeMethod(this, "OpenOutlookProperties", 0, [], 0, 0);
        };
        PersonaActions.prototype.pinPersonaToQuickContacts = function () {
            _invokeMethod(this, "PinPersonaToQuickContacts", 0, [], 0, 0);
        };
        PersonaActions.prototype.scheduleMeeting = function () {
            _invokeMethod(this, "ScheduleMeeting", 0, [], 0, 0);
        };
        PersonaActions.prototype.showContactCard = function (pointToShowX, pointToShowY, personaRectTop, personaRectLeft, personaRectWidth, personaRectHeight) {
            _invokeMethod(this, "ShowContactCard", 0, [pointToShowX, pointToShowY, personaRectTop, personaRectLeft, personaRectWidth, personaRectHeight], 0, 0);
        };
        PersonaActions.prototype.showContextMenu = function (pointToShowX, pointToShowY, personaRectTop, personaRectLeft, personaRectWidth, personaRectHeight) {
            _invokeMethod(this, "ShowContextMenu", 0, [pointToShowX, pointToShowY, personaRectTop, personaRectLeft, personaRectWidth, personaRectHeight], 0, 0);
        };
        PersonaActions.prototype.showExpandedCard = function (pointToShowX, pointToShowY, personaRectTop, personaRectLeft, personaRectWidth, personaRectHeight) {
            _invokeMethod(this, "ShowExpandedCard", 0, [pointToShowX, pointToShowY, personaRectTop, personaRectLeft, personaRectWidth, personaRectHeight], 0, 0);
        };
        PersonaActions.prototype.showHoverCardForPersona = function (pointToShowX, pointToShowY, personaRectTop, personaRectLeft, personaRectWidth, personaRectHeight) {
            _invokeMethod(this, "ShowHoverCardForPersona", 0, [pointToShowX, pointToShowY, personaRectTop, personaRectLeft, personaRectWidth, personaRectHeight], 0, 0);
        };
        PersonaActions.prototype.startAudioCall = function () {
            _invokeMethod(this, "StartAudioCall", 0, [], 0, 0);
        };
        PersonaActions.prototype.startVideoCall = function () {
            _invokeMethod(this, "StartVideoCall", 0, [], 0, 0);
        };
        PersonaActions.prototype.subscribeToGroup = function () {
            _invokeMethod(this, "SubscribeToGroup", 0, [], 0, 0);
        };
        PersonaActions.prototype.toggleTagForAlerts = function () {
            _invokeMethod(this, "ToggleTagForAlerts", 0, [], 0, 0);
        };
        PersonaActions.prototype.unsubscribeFromGroup = function () {
            _invokeMethod(this, "UnsubscribeFromGroup", 0, [], 0, 0);
        };
        PersonaActions.prototype._handleResult = function (value) {
            _super.prototype._handleResult.call(this, value);
            if (_isNullOrUndefined(value))
                return;
            var obj = value;
            _fixObjectPathIfNecessary(this, obj);
        };
        PersonaActions.prototype._handleRetrieveResult = function (value, result) {
            _super.prototype._handleRetrieveResult.call(this, value, result);
            _processRetrieveResult(this, value, result);
        };
        PersonaActions.prototype.toJSON = function () {
            return _toJson(this, {}, {});
        };
        return PersonaActions;
    }(OfficeExtension.ClientObject));
    OfficeCore.PersonaActions = PersonaActions;
    var _typePersonaInfoSource = "PersonaInfoSource";
    var PersonaInfoSource = (function (_super) {
        __extends(PersonaInfoSource, _super);
        function PersonaInfoSource() {
            return _super !== null && _super.apply(this, arguments) || this;
        }
        Object.defineProperty(PersonaInfoSource.prototype, "_className", {
            get: function () {
                return "PersonaInfoSource";
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(PersonaInfoSource.prototype, "_scalarPropertyNames", {
            get: function () {
                return ["displayName", "email", "emailAddresses", "sipAddresses", "birthday", "birthdays", "title", "jobInfoDepartment", "companyName", "office", "linkedTitles", "linkedDepartments", "linkedCompanyNames", "linkedOffices", "phones", "addresses", "webSites", "notes"];
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(PersonaInfoSource.prototype, "_scalarPropertyOriginalNames", {
            get: function () {
                return ["DisplayName", "Email", "EmailAddresses", "SipAddresses", "Birthday", "Birthdays", "Title", "JobInfoDepartment", "CompanyName", "Office", "LinkedTitles", "LinkedDepartments", "LinkedCompanyNames", "LinkedOffices", "Phones", "Addresses", "WebSites", "Notes"];
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(PersonaInfoSource.prototype, "addresses", {
            get: function () {
                _throwIfNotLoaded("addresses", this._A, _typePersonaInfoSource, this._isNull);
                return this._A;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(PersonaInfoSource.prototype, "birthday", {
            get: function () {
                _throwIfNotLoaded("birthday", this._B, _typePersonaInfoSource, this._isNull);
                return this._B;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(PersonaInfoSource.prototype, "birthdays", {
            get: function () {
                _throwIfNotLoaded("birthdays", this._Bi, _typePersonaInfoSource, this._isNull);
                return this._Bi;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(PersonaInfoSource.prototype, "companyName", {
            get: function () {
                _throwIfNotLoaded("companyName", this._C, _typePersonaInfoSource, this._isNull);
                return this._C;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(PersonaInfoSource.prototype, "displayName", {
            get: function () {
                _throwIfNotLoaded("displayName", this._D, _typePersonaInfoSource, this._isNull);
                return this._D;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(PersonaInfoSource.prototype, "email", {
            get: function () {
                _throwIfNotLoaded("email", this._E, _typePersonaInfoSource, this._isNull);
                return this._E;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(PersonaInfoSource.prototype, "emailAddresses", {
            get: function () {
                _throwIfNotLoaded("emailAddresses", this._Em, _typePersonaInfoSource, this._isNull);
                return this._Em;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(PersonaInfoSource.prototype, "jobInfoDepartment", {
            get: function () {
                _throwIfNotLoaded("jobInfoDepartment", this._J, _typePersonaInfoSource, this._isNull);
                return this._J;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(PersonaInfoSource.prototype, "linkedCompanyNames", {
            get: function () {
                _throwIfNotLoaded("linkedCompanyNames", this._L, _typePersonaInfoSource, this._isNull);
                return this._L;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(PersonaInfoSource.prototype, "linkedDepartments", {
            get: function () {
                _throwIfNotLoaded("linkedDepartments", this._Li, _typePersonaInfoSource, this._isNull);
                return this._Li;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(PersonaInfoSource.prototype, "linkedOffices", {
            get: function () {
                _throwIfNotLoaded("linkedOffices", this._Lin, _typePersonaInfoSource, this._isNull);
                return this._Lin;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(PersonaInfoSource.prototype, "linkedTitles", {
            get: function () {
                _throwIfNotLoaded("linkedTitles", this._Link, _typePersonaInfoSource, this._isNull);
                return this._Link;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(PersonaInfoSource.prototype, "notes", {
            get: function () {
                _throwIfNotLoaded("notes", this._N, _typePersonaInfoSource, this._isNull);
                return this._N;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(PersonaInfoSource.prototype, "office", {
            get: function () {
                _throwIfNotLoaded("office", this._O, _typePersonaInfoSource, this._isNull);
                return this._O;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(PersonaInfoSource.prototype, "phones", {
            get: function () {
                _throwIfNotLoaded("phones", this._P, _typePersonaInfoSource, this._isNull);
                return this._P;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(PersonaInfoSource.prototype, "sipAddresses", {
            get: function () {
                _throwIfNotLoaded("sipAddresses", this._S, _typePersonaInfoSource, this._isNull);
                return this._S;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(PersonaInfoSource.prototype, "title", {
            get: function () {
                _throwIfNotLoaded("title", this._T, _typePersonaInfoSource, this._isNull);
                return this._T;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(PersonaInfoSource.prototype, "webSites", {
            get: function () {
                _throwIfNotLoaded("webSites", this._W, _typePersonaInfoSource, this._isNull);
                return this._W;
            },
            enumerable: true,
            configurable: true
        });
        PersonaInfoSource.prototype._handleResult = function (value) {
            _super.prototype._handleResult.call(this, value);
            if (_isNullOrUndefined(value))
                return;
            var obj = value;
            _fixObjectPathIfNecessary(this, obj);
            if (!_isUndefined(obj["Addresses"])) {
                this._A = obj["Addresses"];
            }
            if (!_isUndefined(obj["Birthday"])) {
                this._B = obj["Birthday"];
            }
            if (!_isUndefined(obj["Birthdays"])) {
                this._Bi = obj["Birthdays"];
            }
            if (!_isUndefined(obj["CompanyName"])) {
                this._C = obj["CompanyName"];
            }
            if (!_isUndefined(obj["DisplayName"])) {
                this._D = obj["DisplayName"];
            }
            if (!_isUndefined(obj["Email"])) {
                this._E = obj["Email"];
            }
            if (!_isUndefined(obj["EmailAddresses"])) {
                this._Em = obj["EmailAddresses"];
            }
            if (!_isUndefined(obj["JobInfoDepartment"])) {
                this._J = obj["JobInfoDepartment"];
            }
            if (!_isUndefined(obj["LinkedCompanyNames"])) {
                this._L = obj["LinkedCompanyNames"];
            }
            if (!_isUndefined(obj["LinkedDepartments"])) {
                this._Li = obj["LinkedDepartments"];
            }
            if (!_isUndefined(obj["LinkedOffices"])) {
                this._Lin = obj["LinkedOffices"];
            }
            if (!_isUndefined(obj["LinkedTitles"])) {
                this._Link = obj["LinkedTitles"];
            }
            if (!_isUndefined(obj["Notes"])) {
                this._N = obj["Notes"];
            }
            if (!_isUndefined(obj["Office"])) {
                this._O = obj["Office"];
            }
            if (!_isUndefined(obj["Phones"])) {
                this._P = obj["Phones"];
            }
            if (!_isUndefined(obj["SipAddresses"])) {
                this._S = obj["SipAddresses"];
            }
            if (!_isUndefined(obj["Title"])) {
                this._T = obj["Title"];
            }
            if (!_isUndefined(obj["WebSites"])) {
                this._W = obj["WebSites"];
            }
        };
        PersonaInfoSource.prototype.load = function (options) {
            return _load(this, options);
        };
        PersonaInfoSource.prototype.retrieve = function (option) {
            return _retrieve(this, option);
        };
        PersonaInfoSource.prototype._handleRetrieveResult = function (value, result) {
            _super.prototype._handleRetrieveResult.call(this, value, result);
            _processRetrieveResult(this, value, result);
        };
        PersonaInfoSource.prototype.toJSON = function () {
            return _toJson(this, {
                "addresses": this._A,
                "birthday": this._B,
                "birthdays": this._Bi,
                "companyName": this._C,
                "displayName": this._D,
                "email": this._E,
                "emailAddresses": this._Em,
                "jobInfoDepartment": this._J,
                "linkedCompanyNames": this._L,
                "linkedDepartments": this._Li,
                "linkedOffices": this._Lin,
                "linkedTitles": this._Link,
                "notes": this._N,
                "office": this._O,
                "phones": this._P,
                "sipAddresses": this._S,
                "title": this._T,
                "webSites": this._W
            }, {});
        };
        PersonaInfoSource.prototype.setMockData = function (data) {
            _setMockData(this, data);
        };
        PersonaInfoSource.prototype.ensureUnchanged = function (data) {
            _invokeEnsureUnchanged(this, data);
            return;
        };
        return PersonaInfoSource;
    }(OfficeExtension.ClientObject));
    OfficeCore.PersonaInfoSource = PersonaInfoSource;
    var _typePersonaInfo = "PersonaInfo";
    var PersonaInfo = (function (_super) {
        __extends(PersonaInfo, _super);
        function PersonaInfo() {
            return _super !== null && _super.apply(this, arguments) || this;
        }
        Object.defineProperty(PersonaInfo.prototype, "_className", {
            get: function () {
                return "PersonaInfo";
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(PersonaInfo.prototype, "_scalarPropertyNames", {
            get: function () {
                return ["displayName", "email", "emailAddresses", "sipAddresses", "birthday", "birthdays", "title", "jobInfoDepartment", "companyName", "office", "linkedTitles", "linkedDepartments", "linkedCompanyNames", "linkedOffices", "webSites", "notes", "isPersonResolved"];
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(PersonaInfo.prototype, "_scalarPropertyOriginalNames", {
            get: function () {
                return ["DisplayName", "Email", "EmailAddresses", "SipAddresses", "Birthday", "Birthdays", "Title", "JobInfoDepartment", "CompanyName", "Office", "LinkedTitles", "LinkedDepartments", "LinkedCompanyNames", "LinkedOffices", "WebSites", "Notes", "IsPersonResolved"];
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(PersonaInfo.prototype, "_navigationPropertyNames", {
            get: function () {
                return ["sources"];
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(PersonaInfo.prototype, "sources", {
            get: function () {
                if (!this._So) {
                    this._So = _createPropertyObject(OfficeCore.PersonaInfoSource, this, "Sources", false, 4);
                }
                return this._So;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(PersonaInfo.prototype, "birthday", {
            get: function () {
                _throwIfNotLoaded("birthday", this._B, _typePersonaInfo, this._isNull);
                return this._B;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(PersonaInfo.prototype, "birthdays", {
            get: function () {
                _throwIfNotLoaded("birthdays", this._Bi, _typePersonaInfo, this._isNull);
                return this._Bi;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(PersonaInfo.prototype, "companyName", {
            get: function () {
                _throwIfNotLoaded("companyName", this._C, _typePersonaInfo, this._isNull);
                return this._C;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(PersonaInfo.prototype, "displayName", {
            get: function () {
                _throwIfNotLoaded("displayName", this._D, _typePersonaInfo, this._isNull);
                return this._D;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(PersonaInfo.prototype, "email", {
            get: function () {
                _throwIfNotLoaded("email", this._E, _typePersonaInfo, this._isNull);
                return this._E;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(PersonaInfo.prototype, "emailAddresses", {
            get: function () {
                _throwIfNotLoaded("emailAddresses", this._Em, _typePersonaInfo, this._isNull);
                return this._Em;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(PersonaInfo.prototype, "isPersonResolved", {
            get: function () {
                _throwIfNotLoaded("isPersonResolved", this._I, _typePersonaInfo, this._isNull);
                return this._I;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(PersonaInfo.prototype, "jobInfoDepartment", {
            get: function () {
                _throwIfNotLoaded("jobInfoDepartment", this._J, _typePersonaInfo, this._isNull);
                return this._J;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(PersonaInfo.prototype, "linkedCompanyNames", {
            get: function () {
                _throwIfNotLoaded("linkedCompanyNames", this._L, _typePersonaInfo, this._isNull);
                return this._L;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(PersonaInfo.prototype, "linkedDepartments", {
            get: function () {
                _throwIfNotLoaded("linkedDepartments", this._Li, _typePersonaInfo, this._isNull);
                return this._Li;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(PersonaInfo.prototype, "linkedOffices", {
            get: function () {
                _throwIfNotLoaded("linkedOffices", this._Lin, _typePersonaInfo, this._isNull);
                return this._Lin;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(PersonaInfo.prototype, "linkedTitles", {
            get: function () {
                _throwIfNotLoaded("linkedTitles", this._Link, _typePersonaInfo, this._isNull);
                return this._Link;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(PersonaInfo.prototype, "notes", {
            get: function () {
                _throwIfNotLoaded("notes", this._N, _typePersonaInfo, this._isNull);
                return this._N;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(PersonaInfo.prototype, "office", {
            get: function () {
                _throwIfNotLoaded("office", this._O, _typePersonaInfo, this._isNull);
                return this._O;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(PersonaInfo.prototype, "sipAddresses", {
            get: function () {
                _throwIfNotLoaded("sipAddresses", this._S, _typePersonaInfo, this._isNull);
                return this._S;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(PersonaInfo.prototype, "title", {
            get: function () {
                _throwIfNotLoaded("title", this._T, _typePersonaInfo, this._isNull);
                return this._T;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(PersonaInfo.prototype, "webSites", {
            get: function () {
                _throwIfNotLoaded("webSites", this._W, _typePersonaInfo, this._isNull);
                return this._W;
            },
            enumerable: true,
            configurable: true
        });
        PersonaInfo.prototype.getAddresses = function () {
            return _invokeMethod(this, "GetAddresses", 1, [], 4, 0);
        };
        PersonaInfo.prototype.getPhones = function () {
            return _invokeMethod(this, "GetPhones", 1, [], 4, 0);
        };
        PersonaInfo.prototype._handleResult = function (value) {
            _super.prototype._handleResult.call(this, value);
            if (_isNullOrUndefined(value))
                return;
            var obj = value;
            _fixObjectPathIfNecessary(this, obj);
            if (!_isUndefined(obj["Birthday"])) {
                this._B = _adjustToDateTime(obj["Birthday"]);
            }
            if (!_isUndefined(obj["Birthdays"])) {
                this._Bi = _adjustToDateTime(obj["Birthdays"]);
            }
            if (!_isUndefined(obj["CompanyName"])) {
                this._C = obj["CompanyName"];
            }
            if (!_isUndefined(obj["DisplayName"])) {
                this._D = obj["DisplayName"];
            }
            if (!_isUndefined(obj["Email"])) {
                this._E = obj["Email"];
            }
            if (!_isUndefined(obj["EmailAddresses"])) {
                this._Em = obj["EmailAddresses"];
            }
            if (!_isUndefined(obj["IsPersonResolved"])) {
                this._I = obj["IsPersonResolved"];
            }
            if (!_isUndefined(obj["JobInfoDepartment"])) {
                this._J = obj["JobInfoDepartment"];
            }
            if (!_isUndefined(obj["LinkedCompanyNames"])) {
                this._L = obj["LinkedCompanyNames"];
            }
            if (!_isUndefined(obj["LinkedDepartments"])) {
                this._Li = obj["LinkedDepartments"];
            }
            if (!_isUndefined(obj["LinkedOffices"])) {
                this._Lin = obj["LinkedOffices"];
            }
            if (!_isUndefined(obj["LinkedTitles"])) {
                this._Link = obj["LinkedTitles"];
            }
            if (!_isUndefined(obj["Notes"])) {
                this._N = obj["Notes"];
            }
            if (!_isUndefined(obj["Office"])) {
                this._O = obj["Office"];
            }
            if (!_isUndefined(obj["SipAddresses"])) {
                this._S = obj["SipAddresses"];
            }
            if (!_isUndefined(obj["Title"])) {
                this._T = obj["Title"];
            }
            if (!_isUndefined(obj["WebSites"])) {
                this._W = obj["WebSites"];
            }
            _handleNavigationPropertyResults(this, obj, ["sources", "Sources"]);
        };
        PersonaInfo.prototype.load = function (options) {
            return _load(this, options);
        };
        PersonaInfo.prototype.retrieve = function (option) {
            return _retrieve(this, option);
        };
        PersonaInfo.prototype._handleRetrieveResult = function (value, result) {
            _super.prototype._handleRetrieveResult.call(this, value, result);
            if (_isNullOrUndefined(value))
                return;
            var obj = value;
            if (!_isUndefined(obj["Birthday"])) {
                obj["birthday"] = _adjustToDateTime(obj["birthday"]);
            }
            if (!_isUndefined(obj["Birthdays"])) {
                obj["birthdays"] = _adjustToDateTime(obj["birthdays"]);
            }
            _processRetrieveResult(this, value, result);
        };
        PersonaInfo.prototype.toJSON = function () {
            return _toJson(this, {
                "birthday": this._B,
                "birthdays": this._Bi,
                "companyName": this._C,
                "displayName": this._D,
                "email": this._E,
                "emailAddresses": this._Em,
                "isPersonResolved": this._I,
                "jobInfoDepartment": this._J,
                "linkedCompanyNames": this._L,
                "linkedDepartments": this._Li,
                "linkedOffices": this._Lin,
                "linkedTitles": this._Link,
                "notes": this._N,
                "office": this._O,
                "sipAddresses": this._S,
                "title": this._T,
                "webSites": this._W
            }, {
                "sources": this._So
            });
        };
        PersonaInfo.prototype.setMockData = function (data) {
            _setMockData(this, data);
        };
        PersonaInfo.prototype.ensureUnchanged = function (data) {
            _invokeEnsureUnchanged(this, data);
            return;
        };
        return PersonaInfo;
    }(OfficeExtension.ClientObject));
    OfficeCore.PersonaInfo = PersonaInfo;
    var _typePersonaUnifiedCommunicationInfo = "PersonaUnifiedCommunicationInfo";
    var PersonaUnifiedCommunicationInfo = (function (_super) {
        __extends(PersonaUnifiedCommunicationInfo, _super);
        function PersonaUnifiedCommunicationInfo() {
            return _super !== null && _super.apply(this, arguments) || this;
        }
        Object.defineProperty(PersonaUnifiedCommunicationInfo.prototype, "_className", {
            get: function () {
                return "PersonaUnifiedCommunicationInfo";
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(PersonaUnifiedCommunicationInfo.prototype, "_scalarPropertyNames", {
            get: function () {
                return ["availability", "status", "isSelf", "isTagged", "customStatusString", "isBlocked", "presenceTooltip", "isOutOfOffice", "outOfOfficeNote", "timezone", "meetingLocation", "meetingSubject", "timezoneBias", "idleStartTime", "overallCapability", "isOnBuddyList", "presenceNote", "voiceMailUri", "availabilityText", "availabilityTooltip", "isDurationInAvailabilityText", "freeBusyStatus", "calendarState", "presence"];
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(PersonaUnifiedCommunicationInfo.prototype, "_scalarPropertyOriginalNames", {
            get: function () {
                return ["Availability", "Status", "IsSelf", "IsTagged", "CustomStatusString", "IsBlocked", "PresenceTooltip", "IsOutOfOffice", "OutOfOfficeNote", "Timezone", "MeetingLocation", "MeetingSubject", "TimezoneBias", "IdleStartTime", "OverallCapability", "IsOnBuddyList", "PresenceNote", "VoiceMailUri", "AvailabilityText", "AvailabilityTooltip", "IsDurationInAvailabilityText", "FreeBusyStatus", "CalendarState", "Presence"];
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(PersonaUnifiedCommunicationInfo.prototype, "availability", {
            get: function () {
                _throwIfNotLoaded("availability", this._A, _typePersonaUnifiedCommunicationInfo, this._isNull);
                return this._A;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(PersonaUnifiedCommunicationInfo.prototype, "availabilityText", {
            get: function () {
                _throwIfNotLoaded("availabilityText", this._Av, _typePersonaUnifiedCommunicationInfo, this._isNull);
                return this._Av;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(PersonaUnifiedCommunicationInfo.prototype, "availabilityTooltip", {
            get: function () {
                _throwIfNotLoaded("availabilityTooltip", this._Ava, _typePersonaUnifiedCommunicationInfo, this._isNull);
                return this._Ava;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(PersonaUnifiedCommunicationInfo.prototype, "calendarState", {
            get: function () {
                _throwIfNotLoaded("calendarState", this._C, _typePersonaUnifiedCommunicationInfo, this._isNull);
                return this._C;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(PersonaUnifiedCommunicationInfo.prototype, "customStatusString", {
            get: function () {
                _throwIfNotLoaded("customStatusString", this._Cu, _typePersonaUnifiedCommunicationInfo, this._isNull);
                return this._Cu;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(PersonaUnifiedCommunicationInfo.prototype, "freeBusyStatus", {
            get: function () {
                _throwIfNotLoaded("freeBusyStatus", this._F, _typePersonaUnifiedCommunicationInfo, this._isNull);
                return this._F;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(PersonaUnifiedCommunicationInfo.prototype, "idleStartTime", {
            get: function () {
                _throwIfNotLoaded("idleStartTime", this._I, _typePersonaUnifiedCommunicationInfo, this._isNull);
                return this._I;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(PersonaUnifiedCommunicationInfo.prototype, "isBlocked", {
            get: function () {
                _throwIfNotLoaded("isBlocked", this._Is, _typePersonaUnifiedCommunicationInfo, this._isNull);
                return this._Is;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(PersonaUnifiedCommunicationInfo.prototype, "isDurationInAvailabilityText", {
            get: function () {
                _throwIfNotLoaded("isDurationInAvailabilityText", this._IsD, _typePersonaUnifiedCommunicationInfo, this._isNull);
                return this._IsD;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(PersonaUnifiedCommunicationInfo.prototype, "isOnBuddyList", {
            get: function () {
                _throwIfNotLoaded("isOnBuddyList", this._IsO, _typePersonaUnifiedCommunicationInfo, this._isNull);
                return this._IsO;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(PersonaUnifiedCommunicationInfo.prototype, "isOutOfOffice", {
            get: function () {
                _throwIfNotLoaded("isOutOfOffice", this._IsOu, _typePersonaUnifiedCommunicationInfo, this._isNull);
                return this._IsOu;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(PersonaUnifiedCommunicationInfo.prototype, "isSelf", {
            get: function () {
                _throwIfNotLoaded("isSelf", this._IsS, _typePersonaUnifiedCommunicationInfo, this._isNull);
                return this._IsS;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(PersonaUnifiedCommunicationInfo.prototype, "isTagged", {
            get: function () {
                _throwIfNotLoaded("isTagged", this._IsT, _typePersonaUnifiedCommunicationInfo, this._isNull);
                return this._IsT;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(PersonaUnifiedCommunicationInfo.prototype, "meetingLocation", {
            get: function () {
                _throwIfNotLoaded("meetingLocation", this._M, _typePersonaUnifiedCommunicationInfo, this._isNull);
                return this._M;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(PersonaUnifiedCommunicationInfo.prototype, "meetingSubject", {
            get: function () {
                _throwIfNotLoaded("meetingSubject", this._Me, _typePersonaUnifiedCommunicationInfo, this._isNull);
                return this._Me;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(PersonaUnifiedCommunicationInfo.prototype, "outOfOfficeNote", {
            get: function () {
                _throwIfNotLoaded("outOfOfficeNote", this._O, _typePersonaUnifiedCommunicationInfo, this._isNull);
                return this._O;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(PersonaUnifiedCommunicationInfo.prototype, "overallCapability", {
            get: function () {
                _throwIfNotLoaded("overallCapability", this._Ov, _typePersonaUnifiedCommunicationInfo, this._isNull);
                return this._Ov;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(PersonaUnifiedCommunicationInfo.prototype, "presence", {
            get: function () {
                _throwIfNotLoaded("presence", this._P, _typePersonaUnifiedCommunicationInfo, this._isNull);
                return this._P;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(PersonaUnifiedCommunicationInfo.prototype, "presenceNote", {
            get: function () {
                _throwIfNotLoaded("presenceNote", this._Pr, _typePersonaUnifiedCommunicationInfo, this._isNull);
                return this._Pr;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(PersonaUnifiedCommunicationInfo.prototype, "presenceTooltip", {
            get: function () {
                _throwIfNotLoaded("presenceTooltip", this._Pre, _typePersonaUnifiedCommunicationInfo, this._isNull);
                return this._Pre;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(PersonaUnifiedCommunicationInfo.prototype, "status", {
            get: function () {
                _throwIfNotLoaded("status", this._S, _typePersonaUnifiedCommunicationInfo, this._isNull);
                return this._S;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(PersonaUnifiedCommunicationInfo.prototype, "timezone", {
            get: function () {
                _throwIfNotLoaded("timezone", this._T, _typePersonaUnifiedCommunicationInfo, this._isNull);
                return this._T;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(PersonaUnifiedCommunicationInfo.prototype, "timezoneBias", {
            get: function () {
                _throwIfNotLoaded("timezoneBias", this._Ti, _typePersonaUnifiedCommunicationInfo, this._isNull);
                return this._Ti;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(PersonaUnifiedCommunicationInfo.prototype, "voiceMailUri", {
            get: function () {
                _throwIfNotLoaded("voiceMailUri", this._V, _typePersonaUnifiedCommunicationInfo, this._isNull);
                return this._V;
            },
            enumerable: true,
            configurable: true
        });
        PersonaUnifiedCommunicationInfo.prototype._handleResult = function (value) {
            _super.prototype._handleResult.call(this, value);
            if (_isNullOrUndefined(value))
                return;
            var obj = value;
            _fixObjectPathIfNecessary(this, obj);
            if (!_isUndefined(obj["Availability"])) {
                this._A = obj["Availability"];
            }
            if (!_isUndefined(obj["AvailabilityText"])) {
                this._Av = obj["AvailabilityText"];
            }
            if (!_isUndefined(obj["AvailabilityTooltip"])) {
                this._Ava = obj["AvailabilityTooltip"];
            }
            if (!_isUndefined(obj["CalendarState"])) {
                this._C = obj["CalendarState"];
            }
            if (!_isUndefined(obj["CustomStatusString"])) {
                this._Cu = obj["CustomStatusString"];
            }
            if (!_isUndefined(obj["FreeBusyStatus"])) {
                this._F = obj["FreeBusyStatus"];
            }
            if (!_isUndefined(obj["IdleStartTime"])) {
                this._I = _adjustToDateTime(obj["IdleStartTime"]);
            }
            if (!_isUndefined(obj["IsBlocked"])) {
                this._Is = obj["IsBlocked"];
            }
            if (!_isUndefined(obj["IsDurationInAvailabilityText"])) {
                this._IsD = obj["IsDurationInAvailabilityText"];
            }
            if (!_isUndefined(obj["IsOnBuddyList"])) {
                this._IsO = obj["IsOnBuddyList"];
            }
            if (!_isUndefined(obj["IsOutOfOffice"])) {
                this._IsOu = obj["IsOutOfOffice"];
            }
            if (!_isUndefined(obj["IsSelf"])) {
                this._IsS = obj["IsSelf"];
            }
            if (!_isUndefined(obj["IsTagged"])) {
                this._IsT = obj["IsTagged"];
            }
            if (!_isUndefined(obj["MeetingLocation"])) {
                this._M = obj["MeetingLocation"];
            }
            if (!_isUndefined(obj["MeetingSubject"])) {
                this._Me = obj["MeetingSubject"];
            }
            if (!_isUndefined(obj["OutOfOfficeNote"])) {
                this._O = obj["OutOfOfficeNote"];
            }
            if (!_isUndefined(obj["OverallCapability"])) {
                this._Ov = obj["OverallCapability"];
            }
            if (!_isUndefined(obj["Presence"])) {
                this._P = obj["Presence"];
            }
            if (!_isUndefined(obj["PresenceNote"])) {
                this._Pr = obj["PresenceNote"];
            }
            if (!_isUndefined(obj["PresenceTooltip"])) {
                this._Pre = obj["PresenceTooltip"];
            }
            if (!_isUndefined(obj["Status"])) {
                this._S = obj["Status"];
            }
            if (!_isUndefined(obj["Timezone"])) {
                this._T = obj["Timezone"];
            }
            if (!_isUndefined(obj["TimezoneBias"])) {
                this._Ti = obj["TimezoneBias"];
            }
            if (!_isUndefined(obj["VoiceMailUri"])) {
                this._V = obj["VoiceMailUri"];
            }
        };
        PersonaUnifiedCommunicationInfo.prototype.load = function (options) {
            return _load(this, options);
        };
        PersonaUnifiedCommunicationInfo.prototype.retrieve = function (option) {
            return _retrieve(this, option);
        };
        PersonaUnifiedCommunicationInfo.prototype._handleRetrieveResult = function (value, result) {
            _super.prototype._handleRetrieveResult.call(this, value, result);
            if (_isNullOrUndefined(value))
                return;
            var obj = value;
            if (!_isUndefined(obj["IdleStartTime"])) {
                obj["idleStartTime"] = _adjustToDateTime(obj["idleStartTime"]);
            }
            _processRetrieveResult(this, value, result);
        };
        PersonaUnifiedCommunicationInfo.prototype.toJSON = function () {
            return _toJson(this, {
                "availability": this._A,
                "availabilityText": this._Av,
                "availabilityTooltip": this._Ava,
                "calendarState": this._C,
                "customStatusString": this._Cu,
                "freeBusyStatus": this._F,
                "idleStartTime": this._I,
                "isBlocked": this._Is,
                "isDurationInAvailabilityText": this._IsD,
                "isOnBuddyList": this._IsO,
                "isOutOfOffice": this._IsOu,
                "isSelf": this._IsS,
                "isTagged": this._IsT,
                "meetingLocation": this._M,
                "meetingSubject": this._Me,
                "outOfOfficeNote": this._O,
                "overallCapability": this._Ov,
                "presence": this._P,
                "presenceNote": this._Pr,
                "presenceTooltip": this._Pre,
                "status": this._S,
                "timezone": this._T,
                "timezoneBias": this._Ti,
                "voiceMailUri": this._V
            }, {});
        };
        PersonaUnifiedCommunicationInfo.prototype.setMockData = function (data) {
            _setMockData(this, data);
        };
        PersonaUnifiedCommunicationInfo.prototype.ensureUnchanged = function (data) {
            _invokeEnsureUnchanged(this, data);
            return;
        };
        return PersonaUnifiedCommunicationInfo;
    }(OfficeExtension.ClientObject));
    OfficeCore.PersonaUnifiedCommunicationInfo = PersonaUnifiedCommunicationInfo;
    var _typePersonaPhotoInfo = "PersonaPhotoInfo";
    var PersonaPhotoInfo = (function (_super) {
        __extends(PersonaPhotoInfo, _super);
        function PersonaPhotoInfo() {
            return _super !== null && _super.apply(this, arguments) || this;
        }
        Object.defineProperty(PersonaPhotoInfo.prototype, "_className", {
            get: function () {
                return "PersonaPhotoInfo";
            },
            enumerable: true,
            configurable: true
        });
        PersonaPhotoInfo.prototype.getImageUri = function (uriScheme) {
            return _invokeMethod(this, "getImageUri", 1, [uriScheme], 4, 0);
        };
        PersonaPhotoInfo.prototype.getImageUriWithMetadata = function (uriScheme) {
            return _invokeMethod(this, "getImageUriWithMetadata", 1, [uriScheme], 4, 0);
        };
        PersonaPhotoInfo.prototype.getPlaceholderUri = function (uriScheme) {
            return _invokeMethod(this, "getPlaceholderUri", 1, [uriScheme], 4, 0);
        };
        PersonaPhotoInfo.prototype.setPlaceholderColor = function (color) {
            _invokeMethod(this, "setPlaceholderColor", 0, [color], 0, 0);
        };
        PersonaPhotoInfo.prototype._handleResult = function (value) {
            _super.prototype._handleResult.call(this, value);
            if (_isNullOrUndefined(value))
                return;
            var obj = value;
            _fixObjectPathIfNecessary(this, obj);
        };
        PersonaPhotoInfo.prototype._handleRetrieveResult = function (value, result) {
            _super.prototype._handleRetrieveResult.call(this, value, result);
            _processRetrieveResult(this, value, result);
        };
        PersonaPhotoInfo.prototype.toJSON = function () {
            return _toJson(this, {}, {});
        };
        return PersonaPhotoInfo;
    }(OfficeExtension.ClientObject));
    OfficeCore.PersonaPhotoInfo = PersonaPhotoInfo;
    var _typePersonaCollection = "PersonaCollection";
    var PersonaCollection = (function (_super) {
        __extends(PersonaCollection, _super);
        function PersonaCollection() {
            return _super !== null && _super.apply(this, arguments) || this;
        }
        Object.defineProperty(PersonaCollection.prototype, "_className", {
            get: function () {
                return "PersonaCollection";
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(PersonaCollection.prototype, "_isCollection", {
            get: function () {
                return true;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(PersonaCollection.prototype, "items", {
            get: function () {
                _throwIfNotLoaded("items", this.m__items, _typePersonaCollection, this._isNull);
                return this.m__items;
            },
            enumerable: true,
            configurable: true
        });
        PersonaCollection.prototype.getCount = function () {
            return _invokeMethod(this, "GetCount", 1, [], 4, 0);
        };
        PersonaCollection.prototype.getItem = function (index) {
            return _createIndexerObject(OfficeCore.Persona, this, [index]);
        };
        PersonaCollection.prototype._handleResult = function (value) {
            _super.prototype._handleResult.call(this, value);
            if (_isNullOrUndefined(value))
                return;
            var obj = value;
            _fixObjectPathIfNecessary(this, obj);
            if (!_isNullOrUndefined(obj[OfficeExtension.Constants.items])) {
                this.m__items = [];
                var _data = obj[OfficeExtension.Constants.items];
                for (var i = 0; i < _data.length; i++) {
                    var _item = _createChildItemObject(OfficeCore.Persona, true, this, _data[i], i);
                    _item._handleResult(_data[i]);
                    this.m__items.push(_item);
                }
            }
        };
        PersonaCollection.prototype.load = function (options) {
            return _load(this, options);
        };
        PersonaCollection.prototype.retrieve = function (option) {
            return _retrieve(this, option);
        };
        PersonaCollection.prototype._handleRetrieveResult = function (value, result) {
            var _this = this;
            _super.prototype._handleRetrieveResult.call(this, value, result);
            _processRetrieveResult(this, value, result, function (childItemData, index) { return _createChildItemObject(OfficeCore.Persona, true, _this, childItemData, index); });
        };
        PersonaCollection.prototype.toJSON = function () {
            return _toJson(this, {}, {}, this.m__items);
        };
        PersonaCollection.prototype.setMockData = function (data) {
            var _this = this;
            _setMockData(this, data, function (childItemData, index) { return _createChildItemObject(OfficeCore.Persona, true, _this, childItemData, index); }, function (items) { return _this.m__items = items; });
        };
        return PersonaCollection;
    }(OfficeExtension.ClientObject));
    OfficeCore.PersonaCollection = PersonaCollection;
    var _typePersonaOrganizationInfo = "PersonaOrganizationInfo";
    var PersonaOrganizationInfo = (function (_super) {
        __extends(PersonaOrganizationInfo, _super);
        function PersonaOrganizationInfo() {
            return _super !== null && _super.apply(this, arguments) || this;
        }
        Object.defineProperty(PersonaOrganizationInfo.prototype, "_className", {
            get: function () {
                return "PersonaOrganizationInfo";
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(PersonaOrganizationInfo.prototype, "_scalarPropertyNames", {
            get: function () {
                return ["isWarmedUp", "isWarmingUp"];
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(PersonaOrganizationInfo.prototype, "_scalarPropertyOriginalNames", {
            get: function () {
                return ["IsWarmedUp", "IsWarmingUp"];
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(PersonaOrganizationInfo.prototype, "_navigationPropertyNames", {
            get: function () {
                return ["hierarchy", "manager", "directReports"];
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(PersonaOrganizationInfo.prototype, "directReports", {
            get: function () {
                if (!this._D) {
                    this._D = _createPropertyObject(OfficeCore.PersonaCollection, this, "DirectReports", true, 4);
                }
                return this._D;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(PersonaOrganizationInfo.prototype, "hierarchy", {
            get: function () {
                if (!this._H) {
                    this._H = _createPropertyObject(OfficeCore.PersonaCollection, this, "Hierarchy", true, 4);
                }
                return this._H;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(PersonaOrganizationInfo.prototype, "manager", {
            get: function () {
                if (!this._M) {
                    this._M = _createPropertyObject(OfficeCore.Persona, this, "Manager", false, 4);
                }
                return this._M;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(PersonaOrganizationInfo.prototype, "isWarmedUp", {
            get: function () {
                _throwIfNotLoaded("isWarmedUp", this._I, _typePersonaOrganizationInfo, this._isNull);
                return this._I;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(PersonaOrganizationInfo.prototype, "isWarmingUp", {
            get: function () {
                _throwIfNotLoaded("isWarmingUp", this._Is, _typePersonaOrganizationInfo, this._isNull);
                return this._Is;
            },
            enumerable: true,
            configurable: true
        });
        PersonaOrganizationInfo.prototype._handleResult = function (value) {
            _super.prototype._handleResult.call(this, value);
            if (_isNullOrUndefined(value))
                return;
            var obj = value;
            _fixObjectPathIfNecessary(this, obj);
            if (!_isUndefined(obj["IsWarmedUp"])) {
                this._I = obj["IsWarmedUp"];
            }
            if (!_isUndefined(obj["IsWarmingUp"])) {
                this._Is = obj["IsWarmingUp"];
            }
            _handleNavigationPropertyResults(this, obj, ["directReports", "DirectReports", "hierarchy", "Hierarchy", "manager", "Manager"]);
        };
        PersonaOrganizationInfo.prototype.load = function (options) {
            return _load(this, options);
        };
        PersonaOrganizationInfo.prototype.retrieve = function (option) {
            return _retrieve(this, option);
        };
        PersonaOrganizationInfo.prototype._handleRetrieveResult = function (value, result) {
            _super.prototype._handleRetrieveResult.call(this, value, result);
            _processRetrieveResult(this, value, result);
        };
        PersonaOrganizationInfo.prototype.toJSON = function () {
            return _toJson(this, {
                "isWarmedUp": this._I,
                "isWarmingUp": this._Is
            }, {});
        };
        PersonaOrganizationInfo.prototype.setMockData = function (data) {
            _setMockData(this, data);
        };
        PersonaOrganizationInfo.prototype.ensureUnchanged = function (data) {
            _invokeEnsureUnchanged(this, data);
            return;
        };
        return PersonaOrganizationInfo;
    }(OfficeExtension.ClientObject));
    OfficeCore.PersonaOrganizationInfo = PersonaOrganizationInfo;
    var CustomizedData;
    (function (CustomizedData) {
        CustomizedData["email"] = "Email";
        CustomizedData["workPhone"] = "WorkPhone";
        CustomizedData["workPhone2"] = "WorkPhone2";
        CustomizedData["workFax"] = "WorkFax";
        CustomizedData["mobilePhone"] = "MobilePhone";
        CustomizedData["homePhone"] = "HomePhone";
        CustomizedData["homePhone2"] = "HomePhone2";
        CustomizedData["otherPhone"] = "OtherPhone";
        CustomizedData["sipAddress"] = "SipAddress";
        CustomizedData["profile"] = "Profile";
        CustomizedData["office"] = "Office";
        CustomizedData["company"] = "Company";
        CustomizedData["workAddress"] = "WorkAddress";
        CustomizedData["homeAddress"] = "HomeAddress";
        CustomizedData["otherAddress"] = "OtherAddress";
        CustomizedData["birthday"] = "Birthday";
    })(CustomizedData = OfficeCore.CustomizedData || (OfficeCore.CustomizedData = {}));
    var _typeUnifiedGroupInfo = "UnifiedGroupInfo";
    var UnifiedGroupInfo = (function (_super) {
        __extends(UnifiedGroupInfo, _super);
        function UnifiedGroupInfo() {
            return _super !== null && _super.apply(this, arguments) || this;
        }
        Object.defineProperty(UnifiedGroupInfo.prototype, "_className", {
            get: function () {
                return "UnifiedGroupInfo";
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(UnifiedGroupInfo.prototype, "_scalarPropertyNames", {
            get: function () {
                return ["description", "oneDrive", "oneNote", "isPublic", "amIOwner", "amIMember", "amISubscribed", "memberCount", "ownerCount", "hasGuests", "site", "planner", "classification", "subscriptionEnabled"];
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(UnifiedGroupInfo.prototype, "_scalarPropertyOriginalNames", {
            get: function () {
                return ["Description", "OneDrive", "OneNote", "IsPublic", "AmIOwner", "AmIMember", "AmISubscribed", "MemberCount", "OwnerCount", "HasGuests", "Site", "Planner", "Classification", "SubscriptionEnabled"];
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(UnifiedGroupInfo.prototype, "_scalarPropertyUpdateable", {
            get: function () {
                return [true, true, true, true, true, true, true, true, true, true, true, true, true, true];
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(UnifiedGroupInfo.prototype, "amIMember", {
            get: function () {
                _throwIfNotLoaded("amIMember", this._A, _typeUnifiedGroupInfo, this._isNull);
                return this._A;
            },
            set: function (value) {
                this._A = value;
                _invokeSetProperty(this, "AmIMember", value, 0);
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(UnifiedGroupInfo.prototype, "amIOwner", {
            get: function () {
                _throwIfNotLoaded("amIOwner", this._Am, _typeUnifiedGroupInfo, this._isNull);
                return this._Am;
            },
            set: function (value) {
                this._Am = value;
                _invokeSetProperty(this, "AmIOwner", value, 0);
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(UnifiedGroupInfo.prototype, "amISubscribed", {
            get: function () {
                _throwIfNotLoaded("amISubscribed", this._AmI, _typeUnifiedGroupInfo, this._isNull);
                return this._AmI;
            },
            set: function (value) {
                this._AmI = value;
                _invokeSetProperty(this, "AmISubscribed", value, 0);
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(UnifiedGroupInfo.prototype, "classification", {
            get: function () {
                _throwIfNotLoaded("classification", this._C, _typeUnifiedGroupInfo, this._isNull);
                return this._C;
            },
            set: function (value) {
                this._C = value;
                _invokeSetProperty(this, "Classification", value, 0);
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(UnifiedGroupInfo.prototype, "description", {
            get: function () {
                _throwIfNotLoaded("description", this._D, _typeUnifiedGroupInfo, this._isNull);
                return this._D;
            },
            set: function (value) {
                this._D = value;
                _invokeSetProperty(this, "Description", value, 0);
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(UnifiedGroupInfo.prototype, "hasGuests", {
            get: function () {
                _throwIfNotLoaded("hasGuests", this._H, _typeUnifiedGroupInfo, this._isNull);
                return this._H;
            },
            set: function (value) {
                this._H = value;
                _invokeSetProperty(this, "HasGuests", value, 0);
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(UnifiedGroupInfo.prototype, "isPublic", {
            get: function () {
                _throwIfNotLoaded("isPublic", this._I, _typeUnifiedGroupInfo, this._isNull);
                return this._I;
            },
            set: function (value) {
                this._I = value;
                _invokeSetProperty(this, "IsPublic", value, 0);
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(UnifiedGroupInfo.prototype, "memberCount", {
            get: function () {
                _throwIfNotLoaded("memberCount", this._M, _typeUnifiedGroupInfo, this._isNull);
                return this._M;
            },
            set: function (value) {
                this._M = value;
                _invokeSetProperty(this, "MemberCount", value, 0);
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(UnifiedGroupInfo.prototype, "oneDrive", {
            get: function () {
                _throwIfNotLoaded("oneDrive", this._O, _typeUnifiedGroupInfo, this._isNull);
                return this._O;
            },
            set: function (value) {
                this._O = value;
                _invokeSetProperty(this, "OneDrive", value, 0);
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(UnifiedGroupInfo.prototype, "oneNote", {
            get: function () {
                _throwIfNotLoaded("oneNote", this._On, _typeUnifiedGroupInfo, this._isNull);
                return this._On;
            },
            set: function (value) {
                this._On = value;
                _invokeSetProperty(this, "OneNote", value, 0);
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(UnifiedGroupInfo.prototype, "ownerCount", {
            get: function () {
                _throwIfNotLoaded("ownerCount", this._Ow, _typeUnifiedGroupInfo, this._isNull);
                return this._Ow;
            },
            set: function (value) {
                this._Ow = value;
                _invokeSetProperty(this, "OwnerCount", value, 0);
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(UnifiedGroupInfo.prototype, "planner", {
            get: function () {
                _throwIfNotLoaded("planner", this._P, _typeUnifiedGroupInfo, this._isNull);
                return this._P;
            },
            set: function (value) {
                this._P = value;
                _invokeSetProperty(this, "Planner", value, 0);
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(UnifiedGroupInfo.prototype, "site", {
            get: function () {
                _throwIfNotLoaded("site", this._S, _typeUnifiedGroupInfo, this._isNull);
                return this._S;
            },
            set: function (value) {
                this._S = value;
                _invokeSetProperty(this, "Site", value, 0);
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(UnifiedGroupInfo.prototype, "subscriptionEnabled", {
            get: function () {
                _throwIfNotLoaded("subscriptionEnabled", this._Su, _typeUnifiedGroupInfo, this._isNull);
                return this._Su;
            },
            set: function (value) {
                this._Su = value;
                _invokeSetProperty(this, "SubscriptionEnabled", value, 0);
            },
            enumerable: true,
            configurable: true
        });
        UnifiedGroupInfo.prototype.set = function (properties, options) {
            this._recursivelySet(properties, options, ["description", "oneDrive", "oneNote", "isPublic", "amIOwner", "amIMember", "amISubscribed", "memberCount", "ownerCount", "hasGuests", "site", "planner", "classification", "subscriptionEnabled"], [], []);
        };
        UnifiedGroupInfo.prototype.update = function (properties) {
            this._recursivelyUpdate(properties);
        };
        UnifiedGroupInfo.prototype._handleResult = function (value) {
            _super.prototype._handleResult.call(this, value);
            if (_isNullOrUndefined(value))
                return;
            var obj = value;
            _fixObjectPathIfNecessary(this, obj);
            if (!_isUndefined(obj["AmIMember"])) {
                this._A = obj["AmIMember"];
            }
            if (!_isUndefined(obj["AmIOwner"])) {
                this._Am = obj["AmIOwner"];
            }
            if (!_isUndefined(obj["AmISubscribed"])) {
                this._AmI = obj["AmISubscribed"];
            }
            if (!_isUndefined(obj["Classification"])) {
                this._C = obj["Classification"];
            }
            if (!_isUndefined(obj["Description"])) {
                this._D = obj["Description"];
            }
            if (!_isUndefined(obj["HasGuests"])) {
                this._H = obj["HasGuests"];
            }
            if (!_isUndefined(obj["IsPublic"])) {
                this._I = obj["IsPublic"];
            }
            if (!_isUndefined(obj["MemberCount"])) {
                this._M = obj["MemberCount"];
            }
            if (!_isUndefined(obj["OneDrive"])) {
                this._O = obj["OneDrive"];
            }
            if (!_isUndefined(obj["OneNote"])) {
                this._On = obj["OneNote"];
            }
            if (!_isUndefined(obj["OwnerCount"])) {
                this._Ow = obj["OwnerCount"];
            }
            if (!_isUndefined(obj["Planner"])) {
                this._P = obj["Planner"];
            }
            if (!_isUndefined(obj["Site"])) {
                this._S = obj["Site"];
            }
            if (!_isUndefined(obj["SubscriptionEnabled"])) {
                this._Su = obj["SubscriptionEnabled"];
            }
        };
        UnifiedGroupInfo.prototype.load = function (options) {
            return _load(this, options);
        };
        UnifiedGroupInfo.prototype.retrieve = function (option) {
            return _retrieve(this, option);
        };
        UnifiedGroupInfo.prototype._handleRetrieveResult = function (value, result) {
            _super.prototype._handleRetrieveResult.call(this, value, result);
            _processRetrieveResult(this, value, result);
        };
        UnifiedGroupInfo.prototype.toJSON = function () {
            return _toJson(this, {
                "amIMember": this._A,
                "amIOwner": this._Am,
                "amISubscribed": this._AmI,
                "classification": this._C,
                "description": this._D,
                "hasGuests": this._H,
                "isPublic": this._I,
                "memberCount": this._M,
                "oneDrive": this._O,
                "oneNote": this._On,
                "ownerCount": this._Ow,
                "planner": this._P,
                "site": this._S,
                "subscriptionEnabled": this._Su
            }, {});
        };
        UnifiedGroupInfo.prototype.setMockData = function (data) {
            _setMockData(this, data);
        };
        UnifiedGroupInfo.prototype.ensureUnchanged = function (data) {
            _invokeEnsureUnchanged(this, data);
            return;
        };
        return UnifiedGroupInfo;
    }(OfficeExtension.ClientObject));
    OfficeCore.UnifiedGroupInfo = UnifiedGroupInfo;
    var _typePersona = "Persona";
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
    var Persona = (function (_super) {
        __extends(Persona, _super);
        function Persona() {
            return _super !== null && _super.apply(this, arguments) || this;
        }
        Object.defineProperty(Persona.prototype, "_className", {
            get: function () {
                return "Persona";
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(Persona.prototype, "_scalarPropertyNames", {
            get: function () {
                return ["hostId", "type", "capabilities", "diagnosticId", "instanceId"];
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(Persona.prototype, "_scalarPropertyOriginalNames", {
            get: function () {
                return ["HostId", "Type", "Capabilities", "DiagnosticId", "InstanceId"];
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(Persona.prototype, "_navigationPropertyNames", {
            get: function () {
                return ["photo", "personaInfo", "unifiedCommunicationInfo", "organization", "unifiedGroupInfo", "actions"];
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(Persona.prototype, "actions", {
            get: function () {
                if (!this._A) {
                    this._A = _createPropertyObject(OfficeCore.PersonaActions, this, "Actions", false, 4);
                }
                return this._A;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(Persona.prototype, "organization", {
            get: function () {
                if (!this._O) {
                    this._O = _createPropertyObject(OfficeCore.PersonaOrganizationInfo, this, "Organization", false, 4);
                }
                return this._O;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(Persona.prototype, "personaInfo", {
            get: function () {
                if (!this._P) {
                    this._P = _createPropertyObject(OfficeCore.PersonaInfo, this, "PersonaInfo", false, 4);
                }
                return this._P;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(Persona.prototype, "photo", {
            get: function () {
                if (!this._Ph) {
                    this._Ph = _createPropertyObject(OfficeCore.PersonaPhotoInfo, this, "Photo", false, 4);
                }
                return this._Ph;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(Persona.prototype, "unifiedCommunicationInfo", {
            get: function () {
                if (!this._U) {
                    this._U = _createPropertyObject(OfficeCore.PersonaUnifiedCommunicationInfo, this, "UnifiedCommunicationInfo", false, 4);
                }
                return this._U;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(Persona.prototype, "unifiedGroupInfo", {
            get: function () {
                if (!this._Un) {
                    this._Un = _createPropertyObject(OfficeCore.UnifiedGroupInfo, this, "UnifiedGroupInfo", false, 4);
                }
                return this._Un;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(Persona.prototype, "capabilities", {
            get: function () {
                _throwIfNotLoaded("capabilities", this._C, _typePersona, this._isNull);
                return this._C;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(Persona.prototype, "diagnosticId", {
            get: function () {
                _throwIfNotLoaded("diagnosticId", this._D, _typePersona, this._isNull);
                return this._D;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(Persona.prototype, "hostId", {
            get: function () {
                _throwIfNotLoaded("hostId", this._H, _typePersona, this._isNull);
                return this._H;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(Persona.prototype, "instanceId", {
            get: function () {
                _throwIfNotLoaded("instanceId", this._I, _typePersona, this._isNull);
                return this._I;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(Persona.prototype, "type", {
            get: function () {
                _throwIfNotLoaded("type", this._T, _typePersona, this._isNull);
                return this._T;
            },
            enumerable: true,
            configurable: true
        });
        Persona.prototype.set = function (properties, options) {
            this._recursivelySet(properties, options, [], ["unifiedGroupInfo"], [
                "actions",
                "organization",
                "personaInfo",
                "photo",
                "unifiedCommunicationInfo"
            ]);
        };
        Persona.prototype.update = function (properties) {
            this._recursivelyUpdate(properties);
        };
        Persona.prototype.dispose = function () {
            _invokeMethod(this, "Dispose", 0, [], 0, 0);
        };
        Persona.prototype.getCustomizations = function () {
            return _invokeMethod(this, "GetCustomizations", 1, [], 4, 0);
        };
        Persona.prototype.getMembers = function () {
            return _createMethodObject(OfficeCore.MemberInfoList, this, "GetMembers", 1, [], false, false, null, 4);
        };
        Persona.prototype.getMembership = function () {
            return _createMethodObject(OfficeCore.MemberInfoList, this, "GetMembership", 1, [], false, false, null, 4);
        };
        Persona.prototype.getViewableSources = function () {
            return _invokeMethod(this, "GetViewableSources", 1, [], 4, 0);
        };
        Persona.prototype.reportTimeForRender = function (perfpoint, millisecUTC) {
            _invokeMethod(this, "ReportTimeForRender", 0, [perfpoint, millisecUTC], 0, 0);
        };
        Persona.prototype.warmup = function (dataToWarmUp) {
            _invokeMethod(this, "Warmup", 0, [dataToWarmUp], 0, 0);
        };
        Persona.prototype._handleResult = function (value) {
            _super.prototype._handleResult.call(this, value);
            if (_isNullOrUndefined(value))
                return;
            var obj = value;
            _fixObjectPathIfNecessary(this, obj);
            if (!_isUndefined(obj["Capabilities"])) {
                this._C = obj["Capabilities"];
            }
            if (!_isUndefined(obj["DiagnosticId"])) {
                this._D = obj["DiagnosticId"];
            }
            if (!_isUndefined(obj["HostId"])) {
                this._H = obj["HostId"];
            }
            if (!_isUndefined(obj["InstanceId"])) {
                this._I = obj["InstanceId"];
            }
            if (!_isUndefined(obj["Type"])) {
                this._T = obj["Type"];
            }
            _handleNavigationPropertyResults(this, obj, ["actions", "Actions", "organization", "Organization", "personaInfo", "PersonaInfo", "photo", "Photo", "unifiedCommunicationInfo", "UnifiedCommunicationInfo", "unifiedGroupInfo", "UnifiedGroupInfo"]);
        };
        Persona.prototype.load = function (options) {
            return _load(this, options);
        };
        Persona.prototype.retrieve = function (option) {
            return _retrieve(this, option);
        };
        Persona.prototype._handleRetrieveResult = function (value, result) {
            _super.prototype._handleRetrieveResult.call(this, value, result);
            _processRetrieveResult(this, value, result);
        };
        Persona.prototype.toJSON = function () {
            return _toJson(this, {
                "capabilities": this._C,
                "diagnosticId": this._D,
                "hostId": this._H,
                "instanceId": this._I,
                "type": this._T
            }, {
                "organization": this._O,
                "personaInfo": this._P,
                "unifiedCommunicationInfo": this._U,
                "unifiedGroupInfo": this._Un
            });
        };
        Persona.prototype.setMockData = function (data) {
            _setMockData(this, data);
        };
        Persona.prototype.ensureUnchanged = function (data) {
            _invokeEnsureUnchanged(this, data);
            return;
        };
        return Persona;
    }(OfficeExtension.ClientObject));
    OfficeCore.Persona = Persona;
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
    OfficeExtension.Utility.applyMixin(Persona, PersonaCustom);
    var _typePersonaLifetime = "PersonaLifetime";
    var PersonaLifetime = (function (_super) {
        __extends(PersonaLifetime, _super);
        function PersonaLifetime() {
            return _super !== null && _super.apply(this, arguments) || this;
        }
        Object.defineProperty(PersonaLifetime.prototype, "_className", {
            get: function () {
                return "PersonaLifetime";
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(PersonaLifetime.prototype, "_scalarPropertyNames", {
            get: function () {
                return ["instanceId"];
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(PersonaLifetime.prototype, "_scalarPropertyOriginalNames", {
            get: function () {
                return ["InstanceId"];
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(PersonaLifetime.prototype, "instanceId", {
            get: function () {
                _throwIfNotLoaded("instanceId", this._I, _typePersonaLifetime, this._isNull);
                return this._I;
            },
            enumerable: true,
            configurable: true
        });
        PersonaLifetime.prototype.getPersona = function (hostId) {
            return _createMethodObject(OfficeCore.Persona, this, "GetPersona", 1, [hostId], false, false, null, 4);
        };
        PersonaLifetime.prototype.getPersonaForOrgByEntryId = function (entryId, name, sip, smtp) {
            return _createMethodObject(OfficeCore.Persona, this, "GetPersonaForOrgByEntryId", 1, [entryId, name, sip, smtp], false, false, null, 4);
        };
        PersonaLifetime.prototype.getPersonaForOrgEntry = function (name, sip, smtp, entryId) {
            return _createMethodObject(OfficeCore.Persona, this, "GetPersonaForOrgEntry", 1, [name, sip, smtp, entryId], false, false, null, 4);
        };
        PersonaLifetime.prototype.getPolicies = function () {
            return _invokeMethod(this, "GetPolicies", 1, [], 4, 0);
        };
        PersonaLifetime.prototype._RegisterPersonaUpdatedEvent = function () {
            _invokeMethod(this, "_RegisterPersonaUpdatedEvent", 0, [], 0, 0);
        };
        PersonaLifetime.prototype._UnregisterPersonaUpdatedEvent = function () {
            _invokeMethod(this, "_UnregisterPersonaUpdatedEvent", 0, [], 0, 0);
        };
        PersonaLifetime.prototype._handleResult = function (value) {
            _super.prototype._handleResult.call(this, value);
            if (_isNullOrUndefined(value))
                return;
            var obj = value;
            _fixObjectPathIfNecessary(this, obj);
            if (!_isUndefined(obj["InstanceId"])) {
                this._I = obj["InstanceId"];
            }
        };
        PersonaLifetime.prototype.load = function (options) {
            return _load(this, options);
        };
        PersonaLifetime.prototype.retrieve = function (option) {
            return _retrieve(this, option);
        };
        PersonaLifetime.prototype._handleRetrieveResult = function (value, result) {
            _super.prototype._handleRetrieveResult.call(this, value, result);
            _processRetrieveResult(this, value, result);
        };
        Object.defineProperty(PersonaLifetime.prototype, "onPersonaUpdated", {
            get: function () {
                var _this = this;
                if (!this.m_personaUpdated) {
                    this.m_personaUpdated = new OfficeExtension.GenericEventHandlers(this.context, this, "PersonaUpdated", {
                        eventType: 3502,
                        registerFunc: function () { return _this._RegisterPersonaUpdatedEvent(); },
                        unregisterFunc: function () { return _this._UnregisterPersonaUpdatedEvent(); },
                        getTargetIdFunc: function () { return _this.instanceId; },
                        eventArgsTransformFunc: function (value) {
                            var event = {
                                dataUpdated: value.dataUpdated,
                                sendingPersonaHostId: value.sendingPersonaHostId
                            };
                            return OfficeExtension.Utility._createPromiseFromResult(event);
                        }
                    });
                }
                return this.m_personaUpdated;
            },
            enumerable: true,
            configurable: true
        });
        PersonaLifetime.prototype.toJSON = function () {
            return _toJson(this, {
                "instanceId": this._I
            }, {});
        };
        PersonaLifetime.prototype.setMockData = function (data) {
            _setMockData(this, data);
        };
        PersonaLifetime.prototype.ensureUnchanged = function (data) {
            _invokeEnsureUnchanged(this, data);
            return;
        };
        return PersonaLifetime;
    }(OfficeExtension.ClientObject));
    OfficeCore.PersonaLifetime = PersonaLifetime;
    var _typeLokiTokenProvider = "LokiTokenProvider";
    var LokiTokenProvider = (function (_super) {
        __extends(LokiTokenProvider, _super);
        function LokiTokenProvider() {
            return _super !== null && _super.apply(this, arguments) || this;
        }
        Object.defineProperty(LokiTokenProvider.prototype, "_className", {
            get: function () {
                return "LokiTokenProvider";
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(LokiTokenProvider.prototype, "_scalarPropertyNames", {
            get: function () {
                return ["emailOrUpn", "instanceId"];
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(LokiTokenProvider.prototype, "_scalarPropertyOriginalNames", {
            get: function () {
                return ["EmailOrUpn", "InstanceId"];
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(LokiTokenProvider.prototype, "emailOrUpn", {
            get: function () {
                _throwIfNotLoaded("emailOrUpn", this._E, _typeLokiTokenProvider, this._isNull);
                return this._E;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(LokiTokenProvider.prototype, "instanceId", {
            get: function () {
                _throwIfNotLoaded("instanceId", this._I, _typeLokiTokenProvider, this._isNull);
                return this._I;
            },
            enumerable: true,
            configurable: true
        });
        LokiTokenProvider.prototype.requestClientAccessToken = function () {
            _invokeMethod(this, "RequestClientAccessToken", 0, [], 0, 0);
        };
        LokiTokenProvider.prototype.requestIdentityUniqueId = function () {
            _invokeMethod(this, "RequestIdentityUniqueId", 0, [], 0, 0);
        };
        LokiTokenProvider.prototype.requestToken = function () {
            _invokeMethod(this, "RequestToken", 0, [], 0, 0);
        };
        LokiTokenProvider.prototype._RegisterClientAccessTokenAvailableEvent = function () {
            _invokeMethod(this, "_RegisterClientAccessTokenAvailableEvent", 0, [], 0, 0);
        };
        LokiTokenProvider.prototype._RegisterIdentityUniqueIdAvailableEvent = function () {
            _invokeMethod(this, "_RegisterIdentityUniqueIdAvailableEvent", 0, [], 0, 0);
        };
        LokiTokenProvider.prototype._RegisterLokiTokenAvailableEvent = function () {
            _invokeMethod(this, "_RegisterLokiTokenAvailableEvent", 0, [], 0, 0);
        };
        LokiTokenProvider.prototype._UnregisterClientAccessTokenAvailableEvent = function () {
            _invokeMethod(this, "_UnregisterClientAccessTokenAvailableEvent", 0, [], 0, 0);
        };
        LokiTokenProvider.prototype._UnregisterIdentityUniqueIdAvailableEvent = function () {
            _invokeMethod(this, "_UnregisterIdentityUniqueIdAvailableEvent", 0, [], 0, 0);
        };
        LokiTokenProvider.prototype._UnregisterLokiTokenAvailableEvent = function () {
            _invokeMethod(this, "_UnregisterLokiTokenAvailableEvent", 0, [], 0, 0);
        };
        LokiTokenProvider.prototype._handleResult = function (value) {
            _super.prototype._handleResult.call(this, value);
            if (_isNullOrUndefined(value))
                return;
            var obj = value;
            _fixObjectPathIfNecessary(this, obj);
            if (!_isUndefined(obj["EmailOrUpn"])) {
                this._E = obj["EmailOrUpn"];
            }
            if (!_isUndefined(obj["InstanceId"])) {
                this._I = obj["InstanceId"];
            }
        };
        LokiTokenProvider.prototype.load = function (options) {
            return _load(this, options);
        };
        LokiTokenProvider.prototype.retrieve = function (option) {
            return _retrieve(this, option);
        };
        LokiTokenProvider.prototype._handleRetrieveResult = function (value, result) {
            _super.prototype._handleRetrieveResult.call(this, value, result);
            _processRetrieveResult(this, value, result);
        };
        Object.defineProperty(LokiTokenProvider.prototype, "onClientAccessTokenAvailable", {
            get: function () {
                var _this = this;
                if (!this.m_clientAccessTokenAvailable) {
                    this.m_clientAccessTokenAvailable = new OfficeExtension.GenericEventHandlers(this.context, this, "ClientAccessTokenAvailable", {
                        eventType: 3505,
                        registerFunc: function () { return _this._RegisterClientAccessTokenAvailableEvent(); },
                        unregisterFunc: function () { return _this._UnregisterClientAccessTokenAvailableEvent(); },
                        getTargetIdFunc: function () { return _this.instanceId; },
                        eventArgsTransformFunc: function (value) {
                            var event = {
                                clientAccessToken: value.clientAccessToken,
                                isAvailable: value.isAvailable,
                                tokenTTLInSeconds: value.tokenTTLInSeconds
                            };
                            return OfficeExtension.Utility._createPromiseFromResult(event);
                        }
                    });
                }
                return this.m_clientAccessTokenAvailable;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(LokiTokenProvider.prototype, "onIdentityUniqueIdAvailable", {
            get: function () {
                var _this = this;
                if (!this.m_identityUniqueIdAvailable) {
                    this.m_identityUniqueIdAvailable = new OfficeExtension.GenericEventHandlers(this.context, this, "IdentityUniqueIdAvailable", {
                        eventType: 3504,
                        registerFunc: function () { return _this._RegisterIdentityUniqueIdAvailableEvent(); },
                        unregisterFunc: function () { return _this._UnregisterIdentityUniqueIdAvailableEvent(); },
                        getTargetIdFunc: function () { return _this.instanceId; },
                        eventArgsTransformFunc: function (value) {
                            var event = {
                                isAvailable: value.isAvailable,
                                uniqueId: value.uniqueId
                            };
                            return OfficeExtension.Utility._createPromiseFromResult(event);
                        }
                    });
                }
                return this.m_identityUniqueIdAvailable;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(LokiTokenProvider.prototype, "onLokiTokenAvailable", {
            get: function () {
                var _this = this;
                if (!this.m_lokiTokenAvailable) {
                    this.m_lokiTokenAvailable = new OfficeExtension.GenericEventHandlers(this.context, this, "LokiTokenAvailable", {
                        eventType: 3503,
                        registerFunc: function () { return _this._RegisterLokiTokenAvailableEvent(); },
                        unregisterFunc: function () { return _this._UnregisterLokiTokenAvailableEvent(); },
                        getTargetIdFunc: function () { return _this.instanceId; },
                        eventArgsTransformFunc: function (value) {
                            var event = {
                                isAvailable: value.isAvailable,
                                lokiAutoDiscoverUrl: value.lokiAutoDiscoverUrl,
                                lokiToken: value.lokiToken
                            };
                            return OfficeExtension.Utility._createPromiseFromResult(event);
                        }
                    });
                }
                return this.m_lokiTokenAvailable;
            },
            enumerable: true,
            configurable: true
        });
        LokiTokenProvider.prototype.toJSON = function () {
            return _toJson(this, {
                "emailOrUpn": this._E,
                "instanceId": this._I
            }, {});
        };
        LokiTokenProvider.prototype.setMockData = function (data) {
            _setMockData(this, data);
        };
        LokiTokenProvider.prototype.ensureUnchanged = function (data) {
            _invokeEnsureUnchanged(this, data);
            return;
        };
        return LokiTokenProvider;
    }(OfficeExtension.ClientObject));
    OfficeCore.LokiTokenProvider = LokiTokenProvider;
    var _typeLokiTokenProviderFactory = "LokiTokenProviderFactory";
    var LokiTokenProviderFactory = (function (_super) {
        __extends(LokiTokenProviderFactory, _super);
        function LokiTokenProviderFactory() {
            return _super !== null && _super.apply(this, arguments) || this;
        }
        Object.defineProperty(LokiTokenProviderFactory.prototype, "_className", {
            get: function () {
                return "LokiTokenProviderFactory";
            },
            enumerable: true,
            configurable: true
        });
        LokiTokenProviderFactory.prototype.getLokiTokenProvider = function (accountName) {
            return _createMethodObject(OfficeCore.LokiTokenProvider, this, "GetLokiTokenProvider", 1, [accountName], false, false, null, 4);
        };
        LokiTokenProviderFactory.prototype._handleResult = function (value) {
            _super.prototype._handleResult.call(this, value);
            if (_isNullOrUndefined(value))
                return;
            var obj = value;
            _fixObjectPathIfNecessary(this, obj);
        };
        LokiTokenProviderFactory.prototype._handleRetrieveResult = function (value, result) {
            _super.prototype._handleRetrieveResult.call(this, value, result);
            _processRetrieveResult(this, value, result);
        };
        LokiTokenProviderFactory.newObject = function (context) {
            return _createTopLevelServiceObject(OfficeCore.LokiTokenProviderFactory, context, "Microsoft.People.LokiTokenProviderFactory", false, 4);
        };
        LokiTokenProviderFactory.prototype.toJSON = function () {
            return _toJson(this, {}, {});
        };
        return LokiTokenProviderFactory;
    }(OfficeExtension.ClientObject));
    OfficeCore.LokiTokenProviderFactory = LokiTokenProviderFactory;
    var _typeServiceContext = "ServiceContext";
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
    var ServiceContext = (function (_super) {
        __extends(ServiceContext, _super);
        function ServiceContext() {
            return _super !== null && _super.apply(this, arguments) || this;
        }
        Object.defineProperty(ServiceContext.prototype, "_className", {
            get: function () {
                return "ServiceContext";
            },
            enumerable: true,
            configurable: true
        });
        ServiceContext.prototype.accountEmailOrUpn = function (instanceId) {
            return _invokeMethod(this, "AccountEmailOrUpn", 1, [instanceId], 4, 0);
        };
        ServiceContext.prototype.dispose = function (instance) {
            _invokeMethod(this, "Dispose", 0, [instance], 0, 0);
        };
        ServiceContext.prototype.getInitialPersona = function (instanceId) {
            return _createMethodObject(OfficeCore.Persona, this, "GetInitialPersona", 1, [instanceId], false, false, null, 4);
        };
        ServiceContext.prototype.getLokiTokenProvider = function (instanceId) {
            return _createMethodObject(OfficeCore.LokiTokenProvider, this, "GetLokiTokenProvider", 1, [instanceId], false, false, null, 4);
        };
        ServiceContext.prototype.getPersonaLifetime = function (instanceId) {
            return _createMethodObject(OfficeCore.PersonaLifetime, this, "GetPersonaLifetime", 1, [instanceId], false, false, null, 4);
        };
        ServiceContext.prototype.getPersonaPolicies = function () {
            return _invokeMethod(this, "GetPersonaPolicies", 1, [], 4, 0);
        };
        ServiceContext.prototype._handleResult = function (value) {
            _super.prototype._handleResult.call(this, value);
            if (_isNullOrUndefined(value))
                return;
            var obj = value;
            _fixObjectPathIfNecessary(this, obj);
        };
        ServiceContext.prototype._handleRetrieveResult = function (value, result) {
            _super.prototype._handleRetrieveResult.call(this, value, result);
            _processRetrieveResult(this, value, result);
        };
        ServiceContext.newObject = function (context) {
            return _createTopLevelServiceObject(OfficeCore.ServiceContext, context, "Microsoft.People.ServiceContext", false, 4);
        };
        ServiceContext.prototype.toJSON = function () {
            return _toJson(this, {}, {});
        };
        return ServiceContext;
    }(OfficeExtension.ClientObject));
    OfficeCore.ServiceContext = ServiceContext;
    var _typeRichapiPcxFeatureChecks = "RichapiPcxFeatureChecks";
    var RichapiPcxFeatureChecks = (function (_super) {
        __extends(RichapiPcxFeatureChecks, _super);
        function RichapiPcxFeatureChecks() {
            return _super !== null && _super.apply(this, arguments) || this;
        }
        Object.defineProperty(RichapiPcxFeatureChecks.prototype, "_className", {
            get: function () {
                return "RichapiPcxFeatureChecks";
            },
            enumerable: true,
            configurable: true
        });
        RichapiPcxFeatureChecks.prototype.isAddChangePhotoLinkOnLpcPersonaImageFlightEnabled = function () {
            return _invokeMethod(this, "IsAddChangePhotoLinkOnLpcPersonaImageFlightEnabled", 1, [], 4, 0);
        };
        RichapiPcxFeatureChecks.prototype._handleResult = function (value) {
            _super.prototype._handleResult.call(this, value);
            if (_isNullOrUndefined(value))
                return;
            var obj = value;
            _fixObjectPathIfNecessary(this, obj);
        };
        RichapiPcxFeatureChecks.prototype._handleRetrieveResult = function (value, result) {
            _super.prototype._handleRetrieveResult.call(this, value, result);
            _processRetrieveResult(this, value, result);
        };
        RichapiPcxFeatureChecks.newObject = function (context) {
            return _createTopLevelServiceObject(OfficeCore.RichapiPcxFeatureChecks, context, "Microsoft.People.RichapiPcxFeatureChecks", false, 4);
        };
        RichapiPcxFeatureChecks.prototype.toJSON = function () {
            return _toJson(this, {}, {});
        };
        return RichapiPcxFeatureChecks;
    }(OfficeExtension.ClientObject));
    OfficeCore.RichapiPcxFeatureChecks = RichapiPcxFeatureChecks;
    var _typeTap = "Tap";
    var Tap = (function (_super) {
        __extends(Tap, _super);
        function Tap() {
            return _super !== null && _super.apply(this, arguments) || this;
        }
        Object.defineProperty(Tap.prototype, "_className", {
            get: function () {
                return "Tap";
            },
            enumerable: true,
            configurable: true
        });
        Tap.prototype.getEnterpriseUserInfo = function () {
            return _invokeMethod(this, "GetEnterpriseUserInfo", 1, [], 4 | 1, 0);
        };
        Tap.prototype.getMruFriendlyPath = function (documentUrl) {
            return _invokeMethod(this, "GetMruFriendlyPath", 1, [documentUrl], 4 | 1, 0);
        };
        Tap.prototype.launchFileUrlInOfficeApp = function (documentUrl, useUniversalAsBackup) {
            return _invokeMethod(this, "LaunchFileUrlInOfficeApp", 1, [documentUrl, useUniversalAsBackup], 4 | 1, 0);
        };
        Tap.prototype.performLocalSearch = function (query, numResultsRequested, supportedFileExtensions, documentUrlToExclude) {
            return _invokeMethod(this, "PerformLocalSearch", 1, [query, numResultsRequested, supportedFileExtensions, documentUrlToExclude], 4 | 1, 0);
        };
        Tap.prototype.readSearchCache = function (keyword, expiredHours, filterObjectType) {
            return _invokeMethod(this, "ReadSearchCache", 1, [keyword, expiredHours, filterObjectType], 4 | 1, 0);
        };
        Tap.prototype.writeSearchCache = function (fileContent, keyword, filterObjectType) {
            return _invokeMethod(this, "WriteSearchCache", 1, [fileContent, keyword, filterObjectType], 4 | 1, 0);
        };
        Tap.prototype._handleResult = function (value) {
            _super.prototype._handleResult.call(this, value);
            if (_isNullOrUndefined(value))
                return;
            var obj = value;
            _fixObjectPathIfNecessary(this, obj);
        };
        Tap.prototype._handleRetrieveResult = function (value, result) {
            _super.prototype._handleRetrieveResult.call(this, value, result);
            _processRetrieveResult(this, value, result);
        };
        Tap.newObject = function (context) {
            return _createTopLevelServiceObject(OfficeCore.Tap, context, "Microsoft.TapRichApi.Tap", false, 4);
        };
        Tap.prototype.toJSON = function () {
            return _toJson(this, {}, {});
        };
        return Tap;
    }(OfficeExtension.ClientObject));
    OfficeCore.Tap = Tap;
    var ObjectType;
    (function (ObjectType) {
        ObjectType["unknown"] = "Unknown";
        ObjectType["chart"] = "Chart";
        ObjectType["smartArt"] = "SmartArt";
        ObjectType["table"] = "Table";
        ObjectType["image"] = "Image";
        ObjectType["slide"] = "Slide";
        ObjectType["ole"] = "OLE";
        ObjectType["text"] = "Text";
    })(ObjectType = OfficeCore.ObjectType || (OfficeCore.ObjectType = {}));
    var _typeAppRuntimePersistenceService = "AppRuntimePersistenceService";
    var AppRuntimePersistenceService = (function (_super) {
        __extends(AppRuntimePersistenceService, _super);
        function AppRuntimePersistenceService() {
            return _super !== null && _super.apply(this, arguments) || this;
        }
        Object.defineProperty(AppRuntimePersistenceService.prototype, "_className", {
            get: function () {
                return "AppRuntimePersistenceService";
            },
            enumerable: true,
            configurable: true
        });
        AppRuntimePersistenceService.prototype.getAppRuntimeStartState = function () {
            return _invokeMethod(this, "GetAppRuntimeStartState", 1, [], 4, 0);
        };
        AppRuntimePersistenceService.prototype.setAppRuntimeStartState = function (appRuntimeState) {
            _invokeMethod(this, "SetAppRuntimeStartState", 0, [appRuntimeState], 0, 0);
        };
        AppRuntimePersistenceService.prototype._handleResult = function (value) {
            _super.prototype._handleResult.call(this, value);
            if (_isNullOrUndefined(value))
                return;
            var obj = value;
            _fixObjectPathIfNecessary(this, obj);
        };
        AppRuntimePersistenceService.prototype._handleRetrieveResult = function (value, result) {
            _super.prototype._handleRetrieveResult.call(this, value, result);
            _processRetrieveResult(this, value, result);
        };
        AppRuntimePersistenceService.newObject = function (context) {
            return _createTopLevelServiceObject(OfficeCore.AppRuntimePersistenceService, context, "Microsoft.AppRuntime.AppRuntimePersistenceService", false, 4);
        };
        AppRuntimePersistenceService.prototype.toJSON = function () {
            return _toJson(this, {}, {});
        };
        return AppRuntimePersistenceService;
    }(OfficeExtension.ClientObject));
    OfficeCore.AppRuntimePersistenceService = AppRuntimePersistenceService;
    var _typeAppRuntimeService = "AppRuntimeService";
    var AppRuntimeService = (function (_super) {
        __extends(AppRuntimeService, _super);
        function AppRuntimeService() {
            return _super !== null && _super.apply(this, arguments) || this;
        }
        Object.defineProperty(AppRuntimeService.prototype, "_className", {
            get: function () {
                return "AppRuntimeService";
            },
            enumerable: true,
            configurable: true
        });
        AppRuntimeService.prototype.getAppRuntimeState = function () {
            return _invokeMethod(this, "GetAppRuntimeState", 1, [], 4, 0);
        };
        AppRuntimeService.prototype.setAppRuntimeState = function (appRuntimeState) {
            _invokeMethod(this, "SetAppRuntimeState", 0, [appRuntimeState], 0, 0);
        };
        AppRuntimeService.prototype._handleResult = function (value) {
            _super.prototype._handleResult.call(this, value);
            if (_isNullOrUndefined(value))
                return;
            var obj = value;
            _fixObjectPathIfNecessary(this, obj);
        };
        AppRuntimeService.prototype._handleRetrieveResult = function (value, result) {
            _super.prototype._handleRetrieveResult.call(this, value, result);
            _processRetrieveResult(this, value, result);
        };
        AppRuntimeService.newObject = function (context) {
            return _createTopLevelServiceObject(OfficeCore.AppRuntimeService, context, "Microsoft.AppRuntime.AppRuntimeService", false, 4);
        };
        Object.defineProperty(AppRuntimeService.prototype, "onVisibilityChanged", {
            get: function () {
                if (!this.m_visibilityChanged) {
                    this.m_visibilityChanged = new OfficeExtension.GenericEventHandlers(this.context, this, "VisibilityChanged", {
                        eventType: 65539,
                        registerFunc: function () { return OfficeExtension.Utility._createPromiseFromResult(null); },
                        unregisterFunc: function () { return OfficeExtension.Utility._createPromiseFromResult(null); },
                        getTargetIdFunc: function () { return ""; },
                        eventArgsTransformFunc: function (value) {
                            var event = {
                                visibility: value.visibility
                            };
                            return OfficeExtension.Utility._createPromiseFromResult(event);
                        }
                    });
                }
                return this.m_visibilityChanged;
            },
            enumerable: true,
            configurable: true
        });
        AppRuntimeService.prototype.toJSON = function () {
            return _toJson(this, {}, {});
        };
        return AppRuntimeService;
    }(OfficeExtension.ClientObject));
    OfficeCore.AppRuntimeService = AppRuntimeService;
    var AppRuntimeState;
    (function (AppRuntimeState) {
        AppRuntimeState["inactive"] = "Inactive";
        AppRuntimeState["background"] = "Background";
        AppRuntimeState["visible"] = "Visible";
    })(AppRuntimeState = OfficeCore.AppRuntimeState || (OfficeCore.AppRuntimeState = {}));
    var Visibility;
    (function (Visibility) {
        Visibility["hidden"] = "Hidden";
        Visibility["visible"] = "Visible";
    })(Visibility = OfficeCore.Visibility || (OfficeCore.Visibility = {}));
    var LicenseFeatureTier;
    (function (LicenseFeatureTier) {
        LicenseFeatureTier["unknown"] = "Unknown";
        LicenseFeatureTier["basic"] = "Basic";
        LicenseFeatureTier["premium"] = "Premium";
    })(LicenseFeatureTier = OfficeCore.LicenseFeatureTier || (OfficeCore.LicenseFeatureTier = {}));
    var _typeLicense = "License";
    var License = (function (_super) {
        __extends(License, _super);
        function License() {
            return _super !== null && _super.apply(this, arguments) || this;
        }
        Object.defineProperty(License.prototype, "_className", {
            get: function () {
                return "License";
            },
            enumerable: true,
            configurable: true
        });
        License.prototype.getFeatureTier = function (feature, fallbackValue) {
            return _invokeMethod(this, "GetFeatureTier", 1, [feature, fallbackValue], 4, 0);
        };
        License.prototype.getLicenseFeature = function (feature) {
            return _createMethodObject(OfficeCore.LicenseFeature, this, "GetLicenseFeature", 1, [feature], false, false, null, 4);
        };
        License.prototype.isFeatureEnabled = function (feature, fallbackValue) {
            return _invokeMethod(this, "IsFeatureEnabled", 1, [feature, fallbackValue], 4, 0);
        };
        License.prototype.isFreemiumUpsellEnabled = function () {
            return _invokeMethod(this, "IsFreemiumUpsellEnabled", 1, [], 4, 0);
        };
        License.prototype.launchUpsellExperience = function (experienceId) {
            _invokeMethod(this, "LaunchUpsellExperience", 1, [experienceId], 4, 0);
        };
        License.prototype._TestFireStateChangedEvent = function (feature) {
            _invokeMethod(this, "_TestFireStateChangedEvent", 0, [feature], 1, 0);
        };
        License.prototype._handleResult = function (value) {
            _super.prototype._handleResult.call(this, value);
            if (_isNullOrUndefined(value))
                return;
            var obj = value;
            _fixObjectPathIfNecessary(this, obj);
        };
        License.prototype._handleRetrieveResult = function (value, result) {
            _super.prototype._handleRetrieveResult.call(this, value, result);
            _processRetrieveResult(this, value, result);
        };
        License.newObject = function (context) {
            return _createTopLevelServiceObject(OfficeCore.License, context, "Microsoft.Office.Licensing.License", false, 4);
        };
        License.prototype.toJSON = function () {
            return _toJson(this, {}, {});
        };
        return License;
    }(OfficeExtension.ClientObject));
    OfficeCore.License = License;
    var _typeLicenseFeature = "LicenseFeature";
    var LicenseFeature = (function (_super) {
        __extends(LicenseFeature, _super);
        function LicenseFeature() {
            return _super !== null && _super.apply(this, arguments) || this;
        }
        Object.defineProperty(LicenseFeature.prototype, "_className", {
            get: function () {
                return "LicenseFeature";
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(LicenseFeature.prototype, "_scalarPropertyNames", {
            get: function () {
                return ["id"];
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(LicenseFeature.prototype, "_scalarPropertyOriginalNames", {
            get: function () {
                return ["Id"];
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(LicenseFeature.prototype, "id", {
            get: function () {
                _throwIfNotLoaded("id", this._I, _typeLicenseFeature, this._isNull);
                return this._I;
            },
            enumerable: true,
            configurable: true
        });
        LicenseFeature.prototype._RegisterStateChange = function () {
            _invokeMethod(this, "_RegisterStateChange", 1, [], 4, 0);
        };
        LicenseFeature.prototype._UnregisterStateChange = function () {
            _invokeMethod(this, "_UnregisterStateChange", 1, [], 4, 0);
        };
        LicenseFeature.prototype._handleResult = function (value) {
            _super.prototype._handleResult.call(this, value);
            if (_isNullOrUndefined(value))
                return;
            var obj = value;
            _fixObjectPathIfNecessary(this, obj);
            if (!_isUndefined(obj["Id"])) {
                this._I = obj["Id"];
            }
        };
        LicenseFeature.prototype.load = function (options) {
            return _load(this, options);
        };
        LicenseFeature.prototype.retrieve = function (option) {
            return _retrieve(this, option);
        };
        LicenseFeature.prototype._handleIdResult = function (value) {
            _super.prototype._handleIdResult.call(this, value);
            if (_isNullOrUndefined(value)) {
                return;
            }
            if (!_isUndefined(value["Id"])) {
                this._I = value["Id"];
            }
        };
        LicenseFeature.prototype._handleRetrieveResult = function (value, result) {
            _super.prototype._handleRetrieveResult.call(this, value, result);
            _processRetrieveResult(this, value, result);
        };
        Object.defineProperty(LicenseFeature.prototype, "onStateChanged", {
            get: function () {
                var _this = this;
                if (!this.m_stateChanged) {
                    this.m_stateChanged = new OfficeExtension.GenericEventHandlers(this.context, this, "StateChanged", {
                        eventType: 1,
                        registerFunc: function () { return _this._RegisterStateChange(); },
                        unregisterFunc: function () { return _this._UnregisterStateChange(); },
                        getTargetIdFunc: function () { return _this.id; },
                        eventArgsTransformFunc: function (value) {
                            var event = _CC.LicenseFeature_StateChanged_EventArgsTransform(_this, value);
                            return OfficeExtension.Utility._createPromiseFromResult(event);
                        }
                    });
                }
                return this.m_stateChanged;
            },
            enumerable: true,
            configurable: true
        });
        LicenseFeature.prototype.toJSON = function () {
            return _toJson(this, {
                "id": this._I
            }, {});
        };
        LicenseFeature.prototype.setMockData = function (data) {
            _setMockData(this, data);
        };
        LicenseFeature.prototype.ensureUnchanged = function (data) {
            _invokeEnsureUnchanged(this, data);
            return;
        };
        return LicenseFeature;
    }(OfficeExtension.ClientObject));
    OfficeCore.LicenseFeature = LicenseFeature;
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
                            ret = function () {
                                return __awaiter(_this, void 0, void 0, function () {
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
                                });
                            };
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
                parent.controls
                    .filter(function (control) { return !(!control.id); })
                    .forEach(function (control) {
                    var ribbonControl = ribbon.getButton(control.id);
                    if (control.enabled !== undefined && control.enabled !== null) {
                        ribbonControl.enabled = control.enabled;
                    }
                });
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
            setActionAssociation(Office.actions._association);
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
                    if (OSF._OfficeAppFactory.getHostInfo().hostPlatform === "web" &&
                        OSF._OfficeAppFactory.getHostInfo().hostType !== "excel") {
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
        function initOnce() {
            OfficeExtension.Utility.log('ActionProxy.initOnce');
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
        }
        initOnce();
    })(actionProxy || (actionProxy = {}));
})(Office || (Office = {}));
var __extends = (this && this.__extends) || (function () {
    var extendStatics = Object.setPrototypeOf ||
        ({ __proto__: [] } instanceof Array && function (d, b) { d.__proto__ = b; }) ||
        function (d, b) { for (var p in b)
            if (b.hasOwnProperty(p))
                d[p] = b[p]; };
    return function (d, b) {
        extendStatics(d, b);
        function __() { this.constructor = d; }
        d.prototype = b === null ? Object.create(b) : (__.prototype = b.prototype, new __());
    };
})();
var Visio;
(function (Visio) {
    var _hostName = "Visio";
    var _defaultApiSetName = "";
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
    var _typeApplication = "Application";
    var Application = (function (_super) {
        __extends(Application, _super);
        function Application() {
            return _super !== null && _super.apply(this, arguments) || this;
        }
        Object.defineProperty(Application.prototype, "_className", {
            get: function () {
                return "Application";
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(Application.prototype, "_scalarPropertyNames", {
            get: function () {
                return ["showBorders", "showToolbars"];
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(Application.prototype, "_scalarPropertyOriginalNames", {
            get: function () {
                return ["ShowBorders", "ShowToolbars"];
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(Application.prototype, "_scalarPropertyUpdateable", {
            get: function () {
                return [true, true];
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(Application.prototype, "showBorders", {
            get: function () {
                _throwIfNotLoaded("showBorders", this._S, _typeApplication, this._isNull);
                return this._S;
            },
            set: function (value) {
                this._S = value;
                _invokeSetProperty(this, "ShowBorders", value, 0);
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(Application.prototype, "showToolbars", {
            get: function () {
                _throwIfNotLoaded("showToolbars", this._Sh, _typeApplication, this._isNull);
                return this._Sh;
            },
            set: function (value) {
                this._Sh = value;
                _invokeSetProperty(this, "ShowToolbars", value, 0);
            },
            enumerable: true,
            configurable: true
        });
        Application.prototype.set = function (properties, options) {
            this._recursivelySet(properties, options, ["showBorders", "showToolbars"], [], []);
        };
        Application.prototype.update = function (properties) {
            this._recursivelyUpdate(properties);
        };
        Application.prototype.showToolbar = function (id, show) {
            _invokeMethod(this, "ShowToolbar", 0, [id, show], 0, 0);
        };
        Application.prototype._handleResult = function (value) {
            _super.prototype._handleResult.call(this, value);
            if (_isNullOrUndefined(value))
                return;
            var obj = value;
            _fixObjectPathIfNecessary(this, obj);
            if (!_isUndefined(obj["ShowBorders"])) {
                this._S = obj["ShowBorders"];
            }
            if (!_isUndefined(obj["ShowToolbars"])) {
                this._Sh = obj["ShowToolbars"];
            }
        };
        Application.prototype.load = function (options) {
            return _load(this, options);
        };
        Application.prototype.retrieve = function (option) {
            return _retrieve(this, option);
        };
        Application.prototype._handleRetrieveResult = function (value, result) {
            _super.prototype._handleRetrieveResult.call(this, value, result);
            _processRetrieveResult(this, value, result);
        };
        Application.prototype.toJSON = function () {
            return _toJson(this, {
                "showBorders": this._S,
                "showToolbars": this._Sh
            }, {});
        };
        Application.prototype.setMockData = function (data) {
            _setMockData(this, data);
        };
        Application.prototype.ensureUnchanged = function (data) {
            _invokeEnsureUnchanged(this, data);
            return;
        };
        return Application;
    }(OfficeExtension.ClientObject));
    Visio.Application = Application;
    var _typeDocument = "Document";
    var Document = (function (_super) {
        __extends(Document, _super);
        function Document() {
            return _super !== null && _super.apply(this, arguments) || this;
        }
        Object.defineProperty(Document.prototype, "_className", {
            get: function () {
                return "Document";
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(Document.prototype, "_navigationPropertyNames", {
            get: function () {
                return ["view", "application", "pages"];
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(Document.prototype, "application", {
            get: function () {
                if (!this._A) {
                    this._A = _createPropertyObject(Visio.Application, this, "Application", false, 4);
                }
                return this._A;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(Document.prototype, "pages", {
            get: function () {
                if (!this._P) {
                    this._P = _createPropertyObject(Visio.PageCollection, this, "Pages", true, 4);
                }
                return this._P;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(Document.prototype, "view", {
            get: function () {
                if (!this._V) {
                    this._V = _createPropertyObject(Visio.DocumentView, this, "View", false, 4);
                }
                return this._V;
            },
            enumerable: true,
            configurable: true
        });
        Document.prototype.set = function (properties, options) {
            this._recursivelySet(properties, options, [], ["view", "application"], [
                "pages"
            ]);
        };
        Document.prototype.update = function (properties) {
            this._recursivelyUpdate(properties);
        };
        Document.prototype.getActivePage = function () {
            return _createMethodObject(Visio.Page, this, "GetActivePage", 1, [], false, false, null, 4);
        };
        Document.prototype.setActivePage = function (PageName) {
            _invokeMethod(this, "SetActivePage", 1, [PageName], 4, 0);
        };
        Document.prototype.showTaskPane = function (taskPaneType, initialProps, show) {
            _invokeMethod(this, "ShowTaskPane", 1, [taskPaneType, initialProps, show], 4, 0);
        };
        Document.prototype.startDataRefresh = function () {
            _invokeMethod(this, "StartDataRefresh", 1, [], 4, 0);
        };
        Document.prototype._RegisterDataVisualizerDiagramOperationCompletedEvent = function () {
            _invokeMethod(this, "_RegisterDataVisualizerDiagramOperationCompletedEvent", 0, [], 0, 0);
        };
        Document.prototype._UnregisterDataVisualizerDiagramOperationCompletedEvent = function () {
            _invokeMethod(this, "_UnregisterDataVisualizerDiagramOperationCompletedEvent", 0, [], 0, 0);
        };
        Document.prototype._handleResult = function (value) {
            _super.prototype._handleResult.call(this, value);
            if (_isNullOrUndefined(value))
                return;
            var obj = value;
            _fixObjectPathIfNecessary(this, obj);
            _handleNavigationPropertyResults(this, obj, ["application", "Application", "pages", "Pages", "view", "View"]);
        };
        Document.prototype.load = function (options) {
            return _load(this, options);
        };
        Document.prototype.retrieve = function (option) {
            return _retrieve(this, option);
        };
        Document.prototype._handleRetrieveResult = function (value, result) {
            _super.prototype._handleRetrieveResult.call(this, value, result);
            _processRetrieveResult(this, value, result);
        };
        Object.defineProperty(Document.prototype, "onDataRefreshComplete", {
            get: function () {
                var _this = this;
                if (!this.m_dataRefreshComplete) {
                    this.m_dataRefreshComplete = new OfficeExtension.EventHandlers(this.context, this, "DataRefreshComplete", {
                        registerFunc: function (handlerCallback) {
                            return _this.context.eventRegistration.register(3, "", handlerCallback);
                        },
                        unregisterFunc: function (handlerCallback) {
                            return _this.context.eventRegistration.unregister(3, "", handlerCallback);
                        },
                        eventArgsTransformFunc: function (args) {
                            var evt = {
                                document: this,
                                success: args.ddaBinding.Object.success
                            };
                            return OfficeExtension.Utility._createPromiseFromResult(evt);
                        }
                    });
                }
                return this.m_dataRefreshComplete;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(Document.prototype, "onDataVisualizerDiagramOperationCompleted", {
            get: function () {
                if (!this.m_dataVisualizerDiagramOperationCompleted) {
                    this.m_dataVisualizerDiagramOperationCompleted = new OfficeExtension.EventHandlers(this.context, this, "DataVisualizerDiagramOperationCompleted", null);
                }
                return this.m_dataVisualizerDiagramOperationCompleted;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(Document.prototype, "onDocumentError", {
            get: function () {
                var _this = this;
                if (!this.m_documentError) {
                    this.m_documentError = new OfficeExtension.EventHandlers(this.context, this, "DocumentError", {
                        registerFunc: function (handlerCallback) {
                            return _this.context.eventRegistration.register(15, "", handlerCallback);
                        },
                        unregisterFunc: function (handlerCallback) {
                            return _this.context.eventRegistration.unregister(15, "", handlerCallback);
                        },
                        eventArgsTransformFunc: function (args) {
                            return OfficeExtension.Utility._createPromiseFromResult(args.ddaBinding.Object);
                        }
                    });
                }
                return this.m_documentError;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(Document.prototype, "onDocumentLoadComplete", {
            get: function () {
                var _this = this;
                if (!this.m_documentLoadComplete) {
                    this.m_documentLoadComplete = new OfficeExtension.EventHandlers(this.context, this, "DocumentLoadComplete", {
                        registerFunc: function (handlerCallback) {
                            return _this.context.eventRegistration.register(7, "", handlerCallback);
                        },
                        unregisterFunc: function (handlerCallback) {
                            return _this.context.eventRegistration.unregister(7, "", handlerCallback);
                        },
                        eventArgsTransformFunc: function (args) {
                            var evt = {
                                success: args.ddaBinding.Object.success
                            };
                            return OfficeExtension.Utility._createPromiseFromResult(evt);
                        }
                    });
                }
                return this.m_documentLoadComplete;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(Document.prototype, "onPageLoadComplete", {
            get: function () {
                var _this = this;
                if (!this.m_pageLoadComplete) {
                    this.m_pageLoadComplete = new OfficeExtension.EventHandlers(this.context, this, "PageLoadComplete", {
                        registerFunc: function (handlerCallback) {
                            return _this.context.eventRegistration.register(1, "", handlerCallback);
                        },
                        unregisterFunc: function (handlerCallback) {
                            return _this.context.eventRegistration.unregister(1, "", handlerCallback);
                        },
                        eventArgsTransformFunc: function (args) {
                            return OfficeExtension.Utility._createPromiseFromResult(args.ddaBinding.Object);
                        }
                    });
                }
                return this.m_pageLoadComplete;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(Document.prototype, "onSelectionChanged", {
            get: function () {
                var _this = this;
                if (!this.m_selectionChanged) {
                    this.m_selectionChanged = new OfficeExtension.EventHandlers(this.context, this, "SelectionChanged", {
                        registerFunc: function (handlerCallback) {
                            return _this.context.eventRegistration.register(2, "", handlerCallback);
                        },
                        unregisterFunc: function (handlerCallback) {
                            return _this.context.eventRegistration.unregister(2, "", handlerCallback);
                        },
                        eventArgsTransformFunc: function (args) {
                            return OfficeExtension.Utility._createPromiseFromResult(args.ddaBinding.Object);
                        }
                    });
                }
                return this.m_selectionChanged;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(Document.prototype, "onShapeMouseEnter", {
            get: function () {
                var _this = this;
                if (!this.m_shapeMouseEnter) {
                    this.m_shapeMouseEnter = new OfficeExtension.EventHandlers(this.context, this, "ShapeMouseEnter", {
                        registerFunc: function (handlerCallback) {
                            return _this.context.eventRegistration.register(4, "", handlerCallback);
                        },
                        unregisterFunc: function (handlerCallback) {
                            return _this.context.eventRegistration.unregister(4, "", handlerCallback);
                        },
                        eventArgsTransformFunc: function (args) {
                            return OfficeExtension.Utility._createPromiseFromResult(args.ddaBinding.Object);
                        }
                    });
                }
                return this.m_shapeMouseEnter;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(Document.prototype, "onShapeMouseLeave", {
            get: function () {
                var _this = this;
                if (!this.m_shapeMouseLeave) {
                    this.m_shapeMouseLeave = new OfficeExtension.EventHandlers(this.context, this, "ShapeMouseLeave", {
                        registerFunc: function (handlerCallback) {
                            return _this.context.eventRegistration.register(5, "", handlerCallback);
                        },
                        unregisterFunc: function (handlerCallback) {
                            return _this.context.eventRegistration.unregister(5, "", handlerCallback);
                        },
                        eventArgsTransformFunc: function (args) {
                            return OfficeExtension.Utility._createPromiseFromResult(args.ddaBinding.Object);
                        }
                    });
                }
                return this.m_shapeMouseLeave;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(Document.prototype, "onTaskPaneStateChanged", {
            get: function () {
                var _this = this;
                if (!this.m_taskPaneStateChanged) {
                    this.m_taskPaneStateChanged = new OfficeExtension.EventHandlers(this.context, this, "TaskPaneStateChanged", {
                        registerFunc: function (handlerCallback) {
                            return _this.context.eventRegistration.register(18, "", handlerCallback);
                        },
                        unregisterFunc: function (handlerCallback) {
                            return _this.context.eventRegistration.unregister(18, "", handlerCallback);
                        },
                        eventArgsTransformFunc: function (args) {
                            return OfficeExtension.Utility._createPromiseFromResult(args.ddaBinding.Object);
                        }
                    });
                }
                return this.m_taskPaneStateChanged;
            },
            enumerable: true,
            configurable: true
        });
        Document.prototype.toJSON = function () {
            return _toJson(this, {}, {
                "application": this._A,
                "pages": this._P,
                "view": this._V
            });
        };
        Document.prototype.setMockData = function (data) {
            _setMockData(this, data);
        };
        Document.prototype.ensureUnchanged = function (data) {
            _invokeEnsureUnchanged(this, data);
            return;
        };
        return Document;
    }(OfficeExtension.ClientObject));
    Visio.Document = Document;
    var _typeDocumentView = "DocumentView";
    var DocumentView = (function (_super) {
        __extends(DocumentView, _super);
        function DocumentView() {
            return _super !== null && _super.apply(this, arguments) || this;
        }
        Object.defineProperty(DocumentView.prototype, "_className", {
            get: function () {
                return "DocumentView";
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(DocumentView.prototype, "_scalarPropertyNames", {
            get: function () {
                return ["disableHyperlinks", "disableZoom", "disablePan", "hideDiagramBoundary", "disablePanZoomWindow"];
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(DocumentView.prototype, "_scalarPropertyOriginalNames", {
            get: function () {
                return ["DisableHyperlinks", "DisableZoom", "DisablePan", "HideDiagramBoundary", "DisablePanZoomWindow"];
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(DocumentView.prototype, "_scalarPropertyUpdateable", {
            get: function () {
                return [true, true, true, true, true];
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(DocumentView.prototype, "disableHyperlinks", {
            get: function () {
                _throwIfNotLoaded("disableHyperlinks", this._D, _typeDocumentView, this._isNull);
                return this._D;
            },
            set: function (value) {
                this._D = value;
                _invokeSetProperty(this, "DisableHyperlinks", value, 0);
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(DocumentView.prototype, "disablePan", {
            get: function () {
                _throwIfNotLoaded("disablePan", this._Di, _typeDocumentView, this._isNull);
                return this._Di;
            },
            set: function (value) {
                this._Di = value;
                _invokeSetProperty(this, "DisablePan", value, 0);
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(DocumentView.prototype, "disablePanZoomWindow", {
            get: function () {
                _throwIfNotLoaded("disablePanZoomWindow", this._Dis, _typeDocumentView, this._isNull);
                return this._Dis;
            },
            set: function (value) {
                this._Dis = value;
                _invokeSetProperty(this, "DisablePanZoomWindow", value, 0);
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(DocumentView.prototype, "disableZoom", {
            get: function () {
                _throwIfNotLoaded("disableZoom", this._Disa, _typeDocumentView, this._isNull);
                return this._Disa;
            },
            set: function (value) {
                this._Disa = value;
                _invokeSetProperty(this, "DisableZoom", value, 0);
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(DocumentView.prototype, "hideDiagramBoundary", {
            get: function () {
                _throwIfNotLoaded("hideDiagramBoundary", this._H, _typeDocumentView, this._isNull);
                return this._H;
            },
            set: function (value) {
                this._H = value;
                _invokeSetProperty(this, "HideDiagramBoundary", value, 0);
            },
            enumerable: true,
            configurable: true
        });
        DocumentView.prototype.set = function (properties, options) {
            this._recursivelySet(properties, options, ["disableHyperlinks", "disableZoom", "disablePan", "hideDiagramBoundary", "disablePanZoomWindow"], [], []);
        };
        DocumentView.prototype.update = function (properties) {
            this._recursivelyUpdate(properties);
        };
        DocumentView.prototype._handleResult = function (value) {
            _super.prototype._handleResult.call(this, value);
            if (_isNullOrUndefined(value))
                return;
            var obj = value;
            _fixObjectPathIfNecessary(this, obj);
            if (!_isUndefined(obj["DisableHyperlinks"])) {
                this._D = obj["DisableHyperlinks"];
            }
            if (!_isUndefined(obj["DisablePan"])) {
                this._Di = obj["DisablePan"];
            }
            if (!_isUndefined(obj["DisablePanZoomWindow"])) {
                this._Dis = obj["DisablePanZoomWindow"];
            }
            if (!_isUndefined(obj["DisableZoom"])) {
                this._Disa = obj["DisableZoom"];
            }
            if (!_isUndefined(obj["HideDiagramBoundary"])) {
                this._H = obj["HideDiagramBoundary"];
            }
        };
        DocumentView.prototype.load = function (options) {
            return _load(this, options);
        };
        DocumentView.prototype.retrieve = function (option) {
            return _retrieve(this, option);
        };
        DocumentView.prototype._handleRetrieveResult = function (value, result) {
            _super.prototype._handleRetrieveResult.call(this, value, result);
            _processRetrieveResult(this, value, result);
        };
        DocumentView.prototype.toJSON = function () {
            return _toJson(this, {
                "disableHyperlinks": this._D,
                "disablePan": this._Di,
                "disablePanZoomWindow": this._Dis,
                "disableZoom": this._Disa,
                "hideDiagramBoundary": this._H
            }, {});
        };
        DocumentView.prototype.setMockData = function (data) {
            _setMockData(this, data);
        };
        DocumentView.prototype.ensureUnchanged = function (data) {
            _invokeEnsureUnchanged(this, data);
            return;
        };
        return DocumentView;
    }(OfficeExtension.ClientObject));
    Visio.DocumentView = DocumentView;
    var _typePage = "Page";
    var Page = (function (_super) {
        __extends(Page, _super);
        function Page() {
            return _super !== null && _super.apply(this, arguments) || this;
        }
        Object.defineProperty(Page.prototype, "_className", {
            get: function () {
                return "Page";
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(Page.prototype, "_scalarPropertyNames", {
            get: function () {
                return ["index", "name", "isBackground", "width", "height"];
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(Page.prototype, "_scalarPropertyOriginalNames", {
            get: function () {
                return ["Index", "Name", "IsBackground", "Width", "Height"];
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(Page.prototype, "_navigationPropertyNames", {
            get: function () {
                return ["shapes", "view", "comments", "allShapes", "dataVisualizerDiagrams"];
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(Page.prototype, "allShapes", {
            get: function () {
                if (!this._A) {
                    this._A = _createPropertyObject(Visio.ShapeCollection, this, "AllShapes", true, 4);
                }
                return this._A;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(Page.prototype, "comments", {
            get: function () {
                if (!this._C) {
                    this._C = _createPropertyObject(Visio.CommentCollection, this, "Comments", true, 4);
                }
                return this._C;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(Page.prototype, "dataVisualizerDiagrams", {
            get: function () {
                if (!this._D) {
                    this._D = _createPropertyObject(Visio.DataVisualizerDiagramCollection, this, "DataVisualizerDiagrams", true, 4);
                }
                return this._D;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(Page.prototype, "shapes", {
            get: function () {
                if (!this._S) {
                    this._S = _createPropertyObject(Visio.ShapeCollection, this, "Shapes", true, 4);
                }
                return this._S;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(Page.prototype, "view", {
            get: function () {
                if (!this._V) {
                    this._V = _createPropertyObject(Visio.PageView, this, "View", false, 4);
                }
                return this._V;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(Page.prototype, "height", {
            get: function () {
                _throwIfNotLoaded("height", this._H, _typePage, this._isNull);
                return this._H;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(Page.prototype, "index", {
            get: function () {
                _throwIfNotLoaded("index", this._I, _typePage, this._isNull);
                return this._I;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(Page.prototype, "isBackground", {
            get: function () {
                _throwIfNotLoaded("isBackground", this._Is, _typePage, this._isNull);
                return this._Is;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(Page.prototype, "name", {
            get: function () {
                _throwIfNotLoaded("name", this._N, _typePage, this._isNull);
                return this._N;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(Page.prototype, "width", {
            get: function () {
                _throwIfNotLoaded("width", this._W, _typePage, this._isNull);
                return this._W;
            },
            enumerable: true,
            configurable: true
        });
        Page.prototype.set = function (properties, options) {
            this._recursivelySet(properties, options, [], ["view"], [
                "allShapes",
                "comments",
                "dataVisualizerDiagrams",
                "shapes"
            ]);
        };
        Page.prototype.update = function (properties) {
            this._recursivelyUpdate(properties);
        };
        Page.prototype.activate = function () {
            _invokeMethod(this, "Activate", 1, [], 4, 0);
        };
        Page.prototype._handleResult = function (value) {
            _super.prototype._handleResult.call(this, value);
            if (_isNullOrUndefined(value))
                return;
            var obj = value;
            _fixObjectPathIfNecessary(this, obj);
            if (!_isUndefined(obj["Height"])) {
                this._H = obj["Height"];
            }
            if (!_isUndefined(obj["Index"])) {
                this._I = obj["Index"];
            }
            if (!_isUndefined(obj["IsBackground"])) {
                this._Is = obj["IsBackground"];
            }
            if (!_isUndefined(obj["Name"])) {
                this._N = obj["Name"];
            }
            if (!_isUndefined(obj["Width"])) {
                this._W = obj["Width"];
            }
            _handleNavigationPropertyResults(this, obj, ["allShapes", "AllShapes", "comments", "Comments", "dataVisualizerDiagrams", "DataVisualizerDiagrams", "shapes", "Shapes", "view", "View"]);
        };
        Page.prototype.load = function (options) {
            return _load(this, options);
        };
        Page.prototype.retrieve = function (option) {
            return _retrieve(this, option);
        };
        Page.prototype._handleRetrieveResult = function (value, result) {
            _super.prototype._handleRetrieveResult.call(this, value, result);
            _processRetrieveResult(this, value, result);
        };
        Page.prototype.toJSON = function () {
            return _toJson(this, {
                "height": this._H,
                "index": this._I,
                "isBackground": this._Is,
                "name": this._N,
                "width": this._W
            }, {
                "allShapes": this._A,
                "comments": this._C,
                "dataVisualizerDiagrams": this._D,
                "shapes": this._S,
                "view": this._V
            });
        };
        Page.prototype.setMockData = function (data) {
            _setMockData(this, data);
        };
        Page.prototype.ensureUnchanged = function (data) {
            _invokeEnsureUnchanged(this, data);
            return;
        };
        return Page;
    }(OfficeExtension.ClientObject));
    Visio.Page = Page;
    var _typePageView = "PageView";
    var PageView = (function (_super) {
        __extends(PageView, _super);
        function PageView() {
            return _super !== null && _super.apply(this, arguments) || this;
        }
        Object.defineProperty(PageView.prototype, "_className", {
            get: function () {
                return "PageView";
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(PageView.prototype, "_scalarPropertyNames", {
            get: function () {
                return ["zoom"];
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(PageView.prototype, "_scalarPropertyOriginalNames", {
            get: function () {
                return ["Zoom"];
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(PageView.prototype, "_scalarPropertyUpdateable", {
            get: function () {
                return [true];
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(PageView.prototype, "zoom", {
            get: function () {
                _throwIfNotLoaded("zoom", this._Z, _typePageView, this._isNull);
                return this._Z;
            },
            set: function (value) {
                this._Z = value;
                _invokeSetProperty(this, "Zoom", value, 0);
            },
            enumerable: true,
            configurable: true
        });
        PageView.prototype.set = function (properties, options) {
            this._recursivelySet(properties, options, ["zoom"], [], []);
        };
        PageView.prototype.update = function (properties) {
            this._recursivelyUpdate(properties);
        };
        PageView.prototype.centerViewportOnShape = function (ShapeId) {
            _invokeMethod(this, "CenterViewportOnShape", 1, [ShapeId], 4, 0);
        };
        PageView.prototype.fitToWindow = function () {
            _invokeMethod(this, "FitToWindow", 1, [], 4, 0);
        };
        PageView.prototype.getPosition = function () {
            return _invokeMethod(this, "GetPosition", 1, [], 4, 0);
        };
        PageView.prototype.getSelection = function () {
            return _createMethodObject(Visio.Selection, this, "GetSelection", 1, [], false, false, null, 4);
        };
        PageView.prototype.isShapeInViewport = function (Shape) {
            return _invokeMethod(this, "IsShapeInViewport", 1, [Shape], 4, 0);
        };
        PageView.prototype.setPosition = function (Position) {
            _invokeMethod(this, "SetPosition", 1, [Position], 4, 0);
        };
        PageView.prototype._handleResult = function (value) {
            _super.prototype._handleResult.call(this, value);
            if (_isNullOrUndefined(value))
                return;
            var obj = value;
            _fixObjectPathIfNecessary(this, obj);
            if (!_isUndefined(obj["Zoom"])) {
                this._Z = obj["Zoom"];
            }
        };
        PageView.prototype.load = function (options) {
            return _load(this, options);
        };
        PageView.prototype.retrieve = function (option) {
            return _retrieve(this, option);
        };
        PageView.prototype._handleRetrieveResult = function (value, result) {
            _super.prototype._handleRetrieveResult.call(this, value, result);
            _processRetrieveResult(this, value, result);
        };
        PageView.prototype.toJSON = function () {
            return _toJson(this, {
                "zoom": this._Z
            }, {});
        };
        PageView.prototype.setMockData = function (data) {
            _setMockData(this, data);
        };
        PageView.prototype.ensureUnchanged = function (data) {
            _invokeEnsureUnchanged(this, data);
            return;
        };
        return PageView;
    }(OfficeExtension.ClientObject));
    Visio.PageView = PageView;
    var _typePageCollection = "PageCollection";
    var PageCollection = (function (_super) {
        __extends(PageCollection, _super);
        function PageCollection() {
            return _super !== null && _super.apply(this, arguments) || this;
        }
        Object.defineProperty(PageCollection.prototype, "_className", {
            get: function () {
                return "PageCollection";
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(PageCollection.prototype, "_isCollection", {
            get: function () {
                return true;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(PageCollection.prototype, "items", {
            get: function () {
                _throwIfNotLoaded("items", this.m__items, _typePageCollection, this._isNull);
                return this.m__items;
            },
            enumerable: true,
            configurable: true
        });
        PageCollection.prototype.getCount = function () {
            return _invokeMethod(this, "GetCount", 1, [], 4, 0);
        };
        PageCollection.prototype.getItem = function (key) {
            return _createIndexerObject(Visio.Page, this, [key]);
        };
        PageCollection.prototype._handleResult = function (value) {
            _super.prototype._handleResult.call(this, value);
            if (_isNullOrUndefined(value))
                return;
            var obj = value;
            _fixObjectPathIfNecessary(this, obj);
            if (!_isNullOrUndefined(obj[OfficeExtension.Constants.items])) {
                this.m__items = [];
                var _data = obj[OfficeExtension.Constants.items];
                for (var i = 0; i < _data.length; i++) {
                    var _item = _createChildItemObject(Visio.Page, true, this, _data[i], i);
                    _item._handleResult(_data[i]);
                    this.m__items.push(_item);
                }
            }
        };
        PageCollection.prototype.load = function (options) {
            return _load(this, options);
        };
        PageCollection.prototype.retrieve = function (option) {
            return _retrieve(this, option);
        };
        PageCollection.prototype._handleRetrieveResult = function (value, result) {
            var _this = this;
            _super.prototype._handleRetrieveResult.call(this, value, result);
            _processRetrieveResult(this, value, result, function (childItemData, index) { return _createChildItemObject(Visio.Page, true, _this, childItemData, index); });
        };
        PageCollection.prototype.toJSON = function () {
            return _toJson(this, {}, {}, this.m__items);
        };
        PageCollection.prototype.setMockData = function (data) {
            var _this = this;
            _setMockData(this, data, function (childItemData, index) { return _createChildItemObject(Visio.Page, true, _this, childItemData, index); }, function (items) { return _this.m__items = items; });
        };
        return PageCollection;
    }(OfficeExtension.ClientObject));
    Visio.PageCollection = PageCollection;
    var _typeShapeCollection = "ShapeCollection";
    var ShapeCollection = (function (_super) {
        __extends(ShapeCollection, _super);
        function ShapeCollection() {
            return _super !== null && _super.apply(this, arguments) || this;
        }
        Object.defineProperty(ShapeCollection.prototype, "_className", {
            get: function () {
                return "ShapeCollection";
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(ShapeCollection.prototype, "_isCollection", {
            get: function () {
                return true;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(ShapeCollection.prototype, "items", {
            get: function () {
                _throwIfNotLoaded("items", this.m__items, _typeShapeCollection, this._isNull);
                return this.m__items;
            },
            enumerable: true,
            configurable: true
        });
        ShapeCollection.prototype.getCount = function () {
            return _invokeMethod(this, "GetCount", 1, [], 4, 0);
        };
        ShapeCollection.prototype.getItem = function (key) {
            return _createIndexerObject(Visio.Shape, this, [key]);
        };
        ShapeCollection.prototype._handleResult = function (value) {
            _super.prototype._handleResult.call(this, value);
            if (_isNullOrUndefined(value))
                return;
            var obj = value;
            _fixObjectPathIfNecessary(this, obj);
            if (!_isNullOrUndefined(obj[OfficeExtension.Constants.items])) {
                this.m__items = [];
                var _data = obj[OfficeExtension.Constants.items];
                for (var i = 0; i < _data.length; i++) {
                    var _item = _createChildItemObject(Visio.Shape, true, this, _data[i], i);
                    _item._handleResult(_data[i]);
                    this.m__items.push(_item);
                }
            }
        };
        ShapeCollection.prototype.load = function (options) {
            return _load(this, options);
        };
        ShapeCollection.prototype.retrieve = function (option) {
            return _retrieve(this, option);
        };
        ShapeCollection.prototype._handleRetrieveResult = function (value, result) {
            var _this = this;
            _super.prototype._handleRetrieveResult.call(this, value, result);
            _processRetrieveResult(this, value, result, function (childItemData, index) { return _createChildItemObject(Visio.Shape, true, _this, childItemData, index); });
        };
        ShapeCollection.prototype.toJSON = function () {
            return _toJson(this, {}, {}, this.m__items);
        };
        ShapeCollection.prototype.setMockData = function (data) {
            var _this = this;
            _setMockData(this, data, function (childItemData, index) { return _createChildItemObject(Visio.Shape, true, _this, childItemData, index); }, function (items) { return _this.m__items = items; });
        };
        return ShapeCollection;
    }(OfficeExtension.ClientObject));
    Visio.ShapeCollection = ShapeCollection;
    var _typeShape = "Shape";
    var Shape = (function (_super) {
        __extends(Shape, _super);
        function Shape() {
            return _super !== null && _super.apply(this, arguments) || this;
        }
        Object.defineProperty(Shape.prototype, "_className", {
            get: function () {
                return "Shape";
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(Shape.prototype, "_scalarPropertyNames", {
            get: function () {
                return ["name", "id", "text", "select", "isBoundToData"];
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(Shape.prototype, "_scalarPropertyOriginalNames", {
            get: function () {
                return ["Name", "Id", "Text", "Select", "IsBoundToData"];
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(Shape.prototype, "_scalarPropertyUpdateable", {
            get: function () {
                return [false, false, false, true, false];
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(Shape.prototype, "_navigationPropertyNames", {
            get: function () {
                return ["hyperlinks", "shapeDataItems", "view", "subShapes", "comments"];
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(Shape.prototype, "comments", {
            get: function () {
                if (!this._C) {
                    this._C = _createPropertyObject(Visio.CommentCollection, this, "Comments", true, 4);
                }
                return this._C;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(Shape.prototype, "hyperlinks", {
            get: function () {
                if (!this._H) {
                    this._H = _createPropertyObject(Visio.HyperlinkCollection, this, "Hyperlinks", true, 4);
                }
                return this._H;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(Shape.prototype, "shapeDataItems", {
            get: function () {
                if (!this._Sh) {
                    this._Sh = _createPropertyObject(Visio.ShapeDataItemCollection, this, "ShapeDataItems", true, 4);
                }
                return this._Sh;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(Shape.prototype, "subShapes", {
            get: function () {
                if (!this._Su) {
                    this._Su = _createPropertyObject(Visio.ShapeCollection, this, "SubShapes", true, 4);
                }
                return this._Su;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(Shape.prototype, "view", {
            get: function () {
                if (!this._V) {
                    this._V = _createPropertyObject(Visio.ShapeView, this, "View", false, 4);
                }
                return this._V;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(Shape.prototype, "id", {
            get: function () {
                _throwIfNotLoaded("id", this._I, _typeShape, this._isNull);
                return this._I;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(Shape.prototype, "isBoundToData", {
            get: function () {
                _throwIfNotLoaded("isBoundToData", this._Is, _typeShape, this._isNull);
                return this._Is;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(Shape.prototype, "name", {
            get: function () {
                _throwIfNotLoaded("name", this._N, _typeShape, this._isNull);
                return this._N;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(Shape.prototype, "select", {
            get: function () {
                _throwIfNotLoaded("select", this._S, _typeShape, this._isNull);
                return this._S;
            },
            set: function (value) {
                this._S = value;
                _invokeSetProperty(this, "Select", value, 0);
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(Shape.prototype, "text", {
            get: function () {
                _throwIfNotLoaded("text", this._T, _typeShape, this._isNull);
                return this._T;
            },
            enumerable: true,
            configurable: true
        });
        Shape.prototype.set = function (properties, options) {
            this._recursivelySet(properties, options, ["select"], ["view"], [
                "comments",
                "hyperlinks",
                "shapeDataItems",
                "subShapes"
            ]);
        };
        Shape.prototype.update = function (properties) {
            this._recursivelyUpdate(properties);
        };
        Shape.prototype.getBounds = function () {
            return _invokeMethod(this, "GetBounds", 1, [], 4, 0);
        };
        Shape.prototype._handleResult = function (value) {
            _super.prototype._handleResult.call(this, value);
            if (_isNullOrUndefined(value))
                return;
            var obj = value;
            _fixObjectPathIfNecessary(this, obj);
            if (!_isUndefined(obj["Id"])) {
                this._I = obj["Id"];
            }
            if (!_isUndefined(obj["IsBoundToData"])) {
                this._Is = obj["IsBoundToData"];
            }
            if (!_isUndefined(obj["Name"])) {
                this._N = obj["Name"];
            }
            if (!_isUndefined(obj["Select"])) {
                this._S = obj["Select"];
            }
            if (!_isUndefined(obj["Text"])) {
                this._T = obj["Text"];
            }
            _handleNavigationPropertyResults(this, obj, ["comments", "Comments", "hyperlinks", "Hyperlinks", "shapeDataItems", "ShapeDataItems", "subShapes", "SubShapes", "view", "View"]);
        };
        Shape.prototype.load = function (options) {
            return _load(this, options);
        };
        Shape.prototype.retrieve = function (option) {
            return _retrieve(this, option);
        };
        Shape.prototype._handleIdResult = function (value) {
            _super.prototype._handleIdResult.call(this, value);
            if (_isNullOrUndefined(value)) {
                return;
            }
            if (!_isUndefined(value["Id"])) {
                this._I = value["Id"];
            }
        };
        Shape.prototype._handleRetrieveResult = function (value, result) {
            _super.prototype._handleRetrieveResult.call(this, value, result);
            _processRetrieveResult(this, value, result);
        };
        Shape.prototype.toJSON = function () {
            return _toJson(this, {
                "id": this._I,
                "isBoundToData": this._Is,
                "name": this._N,
                "select": this._S,
                "text": this._T
            }, {
                "comments": this._C,
                "hyperlinks": this._H,
                "shapeDataItems": this._Sh,
                "subShapes": this._Su,
                "view": this._V
            });
        };
        Shape.prototype.setMockData = function (data) {
            _setMockData(this, data);
        };
        Shape.prototype.ensureUnchanged = function (data) {
            _invokeEnsureUnchanged(this, data);
            return;
        };
        return Shape;
    }(OfficeExtension.ClientObject));
    Visio.Shape = Shape;
    var _typeShapeView = "ShapeView";
    var ShapeView = (function (_super) {
        __extends(ShapeView, _super);
        function ShapeView() {
            return _super !== null && _super.apply(this, arguments) || this;
        }
        Object.defineProperty(ShapeView.prototype, "_className", {
            get: function () {
                return "ShapeView";
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(ShapeView.prototype, "_scalarPropertyNames", {
            get: function () {
                return ["highlight"];
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(ShapeView.prototype, "_scalarPropertyOriginalNames", {
            get: function () {
                return ["Highlight"];
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(ShapeView.prototype, "_scalarPropertyUpdateable", {
            get: function () {
                return [true];
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(ShapeView.prototype, "highlight", {
            get: function () {
                _throwIfNotLoaded("highlight", this._H, _typeShapeView, this._isNull);
                return this._H;
            },
            set: function (value) {
                this._H = value;
                _invokeSetProperty(this, "Highlight", value, 0);
            },
            enumerable: true,
            configurable: true
        });
        ShapeView.prototype.set = function (properties, options) {
            this._recursivelySet(properties, options, ["highlight"], [], []);
        };
        ShapeView.prototype.update = function (properties) {
            this._recursivelyUpdate(properties);
        };
        ShapeView.prototype.addOverlay = function (OverlayType, Content, OverlayHorizontalAlignment, OverlayVerticalAlignment, Width, Height) {
            return _invokeMethod(this, "AddOverlay", 1, [OverlayType, Content, OverlayHorizontalAlignment, OverlayVerticalAlignment, Width, Height], 4, 0);
        };
        ShapeView.prototype.removeOverlay = function (OverlayId) {
            _invokeMethod(this, "RemoveOverlay", 1, [OverlayId], 4, 0);
        };
        ShapeView.prototype.showOverlay = function (overlayId, show) {
            _invokeMethod(this, "ShowOverlay", 1, [overlayId, show], 4, 0);
        };
        ShapeView.prototype._handleResult = function (value) {
            _super.prototype._handleResult.call(this, value);
            if (_isNullOrUndefined(value))
                return;
            var obj = value;
            _fixObjectPathIfNecessary(this, obj);
            if (!_isUndefined(obj["Highlight"])) {
                this._H = obj["Highlight"];
            }
        };
        ShapeView.prototype.load = function (options) {
            return _load(this, options);
        };
        ShapeView.prototype.retrieve = function (option) {
            return _retrieve(this, option);
        };
        ShapeView.prototype._handleRetrieveResult = function (value, result) {
            _super.prototype._handleRetrieveResult.call(this, value, result);
            _processRetrieveResult(this, value, result);
        };
        ShapeView.prototype.toJSON = function () {
            return _toJson(this, {
                "highlight": this._H
            }, {});
        };
        ShapeView.prototype.setMockData = function (data) {
            _setMockData(this, data);
        };
        ShapeView.prototype.ensureUnchanged = function (data) {
            _invokeEnsureUnchanged(this, data);
            return;
        };
        return ShapeView;
    }(OfficeExtension.ClientObject));
    Visio.ShapeView = ShapeView;
    var _typeShapeDataItemCollection = "ShapeDataItemCollection";
    var ShapeDataItemCollection = (function (_super) {
        __extends(ShapeDataItemCollection, _super);
        function ShapeDataItemCollection() {
            return _super !== null && _super.apply(this, arguments) || this;
        }
        Object.defineProperty(ShapeDataItemCollection.prototype, "_className", {
            get: function () {
                return "ShapeDataItemCollection";
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(ShapeDataItemCollection.prototype, "_isCollection", {
            get: function () {
                return true;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(ShapeDataItemCollection.prototype, "items", {
            get: function () {
                _throwIfNotLoaded("items", this.m__items, _typeShapeDataItemCollection, this._isNull);
                return this.m__items;
            },
            enumerable: true,
            configurable: true
        });
        ShapeDataItemCollection.prototype.getCount = function () {
            return _invokeMethod(this, "GetCount", 1, [], 4, 0);
        };
        ShapeDataItemCollection.prototype.getItem = function (key) {
            return _createIndexerObject(Visio.ShapeDataItem, this, [key]);
        };
        ShapeDataItemCollection.prototype._handleResult = function (value) {
            _super.prototype._handleResult.call(this, value);
            if (_isNullOrUndefined(value))
                return;
            var obj = value;
            _fixObjectPathIfNecessary(this, obj);
            if (!_isNullOrUndefined(obj[OfficeExtension.Constants.items])) {
                this.m__items = [];
                var _data = obj[OfficeExtension.Constants.items];
                for (var i = 0; i < _data.length; i++) {
                    var _item = _createChildItemObject(Visio.ShapeDataItem, true, this, _data[i], i);
                    _item._handleResult(_data[i]);
                    this.m__items.push(_item);
                }
            }
        };
        ShapeDataItemCollection.prototype.load = function (options) {
            return _load(this, options);
        };
        ShapeDataItemCollection.prototype.retrieve = function (option) {
            return _retrieve(this, option);
        };
        ShapeDataItemCollection.prototype._handleRetrieveResult = function (value, result) {
            var _this = this;
            _super.prototype._handleRetrieveResult.call(this, value, result);
            _processRetrieveResult(this, value, result, function (childItemData, index) { return _createChildItemObject(Visio.ShapeDataItem, true, _this, childItemData, index); });
        };
        ShapeDataItemCollection.prototype.toJSON = function () {
            return _toJson(this, {}, {}, this.m__items);
        };
        ShapeDataItemCollection.prototype.setMockData = function (data) {
            var _this = this;
            _setMockData(this, data, function (childItemData, index) { return _createChildItemObject(Visio.ShapeDataItem, true, _this, childItemData, index); }, function (items) { return _this.m__items = items; });
        };
        return ShapeDataItemCollection;
    }(OfficeExtension.ClientObject));
    Visio.ShapeDataItemCollection = ShapeDataItemCollection;
    var _typeShapeDataItem = "ShapeDataItem";
    var ShapeDataItem = (function (_super) {
        __extends(ShapeDataItem, _super);
        function ShapeDataItem() {
            return _super !== null && _super.apply(this, arguments) || this;
        }
        Object.defineProperty(ShapeDataItem.prototype, "_className", {
            get: function () {
                return "ShapeDataItem";
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(ShapeDataItem.prototype, "_scalarPropertyNames", {
            get: function () {
                return ["label", "value", "format", "formattedValue"];
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(ShapeDataItem.prototype, "_scalarPropertyOriginalNames", {
            get: function () {
                return ["Label", "Value", "Format", "FormattedValue"];
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(ShapeDataItem.prototype, "format", {
            get: function () {
                _throwIfNotLoaded("format", this._F, _typeShapeDataItem, this._isNull);
                return this._F;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(ShapeDataItem.prototype, "formattedValue", {
            get: function () {
                _throwIfNotLoaded("formattedValue", this._Fo, _typeShapeDataItem, this._isNull);
                return this._Fo;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(ShapeDataItem.prototype, "label", {
            get: function () {
                _throwIfNotLoaded("label", this._L, _typeShapeDataItem, this._isNull);
                return this._L;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(ShapeDataItem.prototype, "value", {
            get: function () {
                _throwIfNotLoaded("value", this._V, _typeShapeDataItem, this._isNull);
                return this._V;
            },
            enumerable: true,
            configurable: true
        });
        ShapeDataItem.prototype._handleResult = function (value) {
            _super.prototype._handleResult.call(this, value);
            if (_isNullOrUndefined(value))
                return;
            var obj = value;
            _fixObjectPathIfNecessary(this, obj);
            if (!_isUndefined(obj["Format"])) {
                this._F = obj["Format"];
            }
            if (!_isUndefined(obj["FormattedValue"])) {
                this._Fo = obj["FormattedValue"];
            }
            if (!_isUndefined(obj["Label"])) {
                this._L = obj["Label"];
            }
            if (!_isUndefined(obj["Value"])) {
                this._V = obj["Value"];
            }
        };
        ShapeDataItem.prototype.load = function (options) {
            return _load(this, options);
        };
        ShapeDataItem.prototype.retrieve = function (option) {
            return _retrieve(this, option);
        };
        ShapeDataItem.prototype._handleRetrieveResult = function (value, result) {
            _super.prototype._handleRetrieveResult.call(this, value, result);
            _processRetrieveResult(this, value, result);
        };
        ShapeDataItem.prototype.toJSON = function () {
            return _toJson(this, {
                "format": this._F,
                "formattedValue": this._Fo,
                "label": this._L,
                "value": this._V
            }, {});
        };
        ShapeDataItem.prototype.setMockData = function (data) {
            _setMockData(this, data);
        };
        ShapeDataItem.prototype.ensureUnchanged = function (data) {
            _invokeEnsureUnchanged(this, data);
            return;
        };
        return ShapeDataItem;
    }(OfficeExtension.ClientObject));
    Visio.ShapeDataItem = ShapeDataItem;
    var _typeHyperlinkCollection = "HyperlinkCollection";
    var HyperlinkCollection = (function (_super) {
        __extends(HyperlinkCollection, _super);
        function HyperlinkCollection() {
            return _super !== null && _super.apply(this, arguments) || this;
        }
        Object.defineProperty(HyperlinkCollection.prototype, "_className", {
            get: function () {
                return "HyperlinkCollection";
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(HyperlinkCollection.prototype, "_isCollection", {
            get: function () {
                return true;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(HyperlinkCollection.prototype, "items", {
            get: function () {
                _throwIfNotLoaded("items", this.m__items, _typeHyperlinkCollection, this._isNull);
                return this.m__items;
            },
            enumerable: true,
            configurable: true
        });
        HyperlinkCollection.prototype.getCount = function () {
            return _invokeMethod(this, "GetCount", 1, [], 4, 0);
        };
        HyperlinkCollection.prototype.getItem = function (Key) {
            return _createIndexerObject(Visio.Hyperlink, this, [Key]);
        };
        HyperlinkCollection.prototype._handleResult = function (value) {
            _super.prototype._handleResult.call(this, value);
            if (_isNullOrUndefined(value))
                return;
            var obj = value;
            _fixObjectPathIfNecessary(this, obj);
            if (!_isNullOrUndefined(obj[OfficeExtension.Constants.items])) {
                this.m__items = [];
                var _data = obj[OfficeExtension.Constants.items];
                for (var i = 0; i < _data.length; i++) {
                    var _item = _createChildItemObject(Visio.Hyperlink, true, this, _data[i], i);
                    _item._handleResult(_data[i]);
                    this.m__items.push(_item);
                }
            }
        };
        HyperlinkCollection.prototype.load = function (options) {
            return _load(this, options);
        };
        HyperlinkCollection.prototype.retrieve = function (option) {
            return _retrieve(this, option);
        };
        HyperlinkCollection.prototype._handleRetrieveResult = function (value, result) {
            var _this = this;
            _super.prototype._handleRetrieveResult.call(this, value, result);
            _processRetrieveResult(this, value, result, function (childItemData, index) { return _createChildItemObject(Visio.Hyperlink, true, _this, childItemData, index); });
        };
        HyperlinkCollection.prototype.toJSON = function () {
            return _toJson(this, {}, {}, this.m__items);
        };
        HyperlinkCollection.prototype.setMockData = function (data) {
            var _this = this;
            _setMockData(this, data, function (childItemData, index) { return _createChildItemObject(Visio.Hyperlink, true, _this, childItemData, index); }, function (items) { return _this.m__items = items; });
        };
        return HyperlinkCollection;
    }(OfficeExtension.ClientObject));
    Visio.HyperlinkCollection = HyperlinkCollection;
    var _typeHyperlink = "Hyperlink";
    var Hyperlink = (function (_super) {
        __extends(Hyperlink, _super);
        function Hyperlink() {
            return _super !== null && _super.apply(this, arguments) || this;
        }
        Object.defineProperty(Hyperlink.prototype, "_className", {
            get: function () {
                return "Hyperlink";
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(Hyperlink.prototype, "_scalarPropertyNames", {
            get: function () {
                return ["address", "subAddress", "description", "extraInfo"];
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(Hyperlink.prototype, "_scalarPropertyOriginalNames", {
            get: function () {
                return ["Address", "SubAddress", "Description", "ExtraInfo"];
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(Hyperlink.prototype, "address", {
            get: function () {
                _throwIfNotLoaded("address", this._A, _typeHyperlink, this._isNull);
                return this._A;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(Hyperlink.prototype, "description", {
            get: function () {
                _throwIfNotLoaded("description", this._D, _typeHyperlink, this._isNull);
                return this._D;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(Hyperlink.prototype, "extraInfo", {
            get: function () {
                _throwIfNotLoaded("extraInfo", this._E, _typeHyperlink, this._isNull);
                return this._E;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(Hyperlink.prototype, "subAddress", {
            get: function () {
                _throwIfNotLoaded("subAddress", this._S, _typeHyperlink, this._isNull);
                return this._S;
            },
            enumerable: true,
            configurable: true
        });
        Hyperlink.prototype._handleResult = function (value) {
            _super.prototype._handleResult.call(this, value);
            if (_isNullOrUndefined(value))
                return;
            var obj = value;
            _fixObjectPathIfNecessary(this, obj);
            if (!_isUndefined(obj["Address"])) {
                this._A = obj["Address"];
            }
            if (!_isUndefined(obj["Description"])) {
                this._D = obj["Description"];
            }
            if (!_isUndefined(obj["ExtraInfo"])) {
                this._E = obj["ExtraInfo"];
            }
            if (!_isUndefined(obj["SubAddress"])) {
                this._S = obj["SubAddress"];
            }
        };
        Hyperlink.prototype.load = function (options) {
            return _load(this, options);
        };
        Hyperlink.prototype.retrieve = function (option) {
            return _retrieve(this, option);
        };
        Hyperlink.prototype._handleRetrieveResult = function (value, result) {
            _super.prototype._handleRetrieveResult.call(this, value, result);
            _processRetrieveResult(this, value, result);
        };
        Hyperlink.prototype.toJSON = function () {
            return _toJson(this, {
                "address": this._A,
                "description": this._D,
                "extraInfo": this._E,
                "subAddress": this._S
            }, {});
        };
        Hyperlink.prototype.setMockData = function (data) {
            _setMockData(this, data);
        };
        Hyperlink.prototype.ensureUnchanged = function (data) {
            _invokeEnsureUnchanged(this, data);
            return;
        };
        return Hyperlink;
    }(OfficeExtension.ClientObject));
    Visio.Hyperlink = Hyperlink;
    var _typeCommentCollection = "CommentCollection";
    var CommentCollection = (function (_super) {
        __extends(CommentCollection, _super);
        function CommentCollection() {
            return _super !== null && _super.apply(this, arguments) || this;
        }
        Object.defineProperty(CommentCollection.prototype, "_className", {
            get: function () {
                return "CommentCollection";
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(CommentCollection.prototype, "_isCollection", {
            get: function () {
                return true;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(CommentCollection.prototype, "items", {
            get: function () {
                _throwIfNotLoaded("items", this.m__items, _typeCommentCollection, this._isNull);
                return this.m__items;
            },
            enumerable: true,
            configurable: true
        });
        CommentCollection.prototype.getCount = function () {
            return _invokeMethod(this, "GetCount", 1, [], 4, 0);
        };
        CommentCollection.prototype.getItem = function (key) {
            return _createIndexerObject(Visio.Comment, this, [key]);
        };
        CommentCollection.prototype._handleResult = function (value) {
            _super.prototype._handleResult.call(this, value);
            if (_isNullOrUndefined(value))
                return;
            var obj = value;
            _fixObjectPathIfNecessary(this, obj);
            if (!_isNullOrUndefined(obj[OfficeExtension.Constants.items])) {
                this.m__items = [];
                var _data = obj[OfficeExtension.Constants.items];
                for (var i = 0; i < _data.length; i++) {
                    var _item = _createChildItemObject(Visio.Comment, true, this, _data[i], i);
                    _item._handleResult(_data[i]);
                    this.m__items.push(_item);
                }
            }
        };
        CommentCollection.prototype.load = function (options) {
            return _load(this, options);
        };
        CommentCollection.prototype.retrieve = function (option) {
            return _retrieve(this, option);
        };
        CommentCollection.prototype._handleRetrieveResult = function (value, result) {
            var _this = this;
            _super.prototype._handleRetrieveResult.call(this, value, result);
            _processRetrieveResult(this, value, result, function (childItemData, index) { return _createChildItemObject(Visio.Comment, true, _this, childItemData, index); });
        };
        CommentCollection.prototype.toJSON = function () {
            return _toJson(this, {}, {}, this.m__items);
        };
        CommentCollection.prototype.setMockData = function (data) {
            var _this = this;
            _setMockData(this, data, function (childItemData, index) { return _createChildItemObject(Visio.Comment, true, _this, childItemData, index); }, function (items) { return _this.m__items = items; });
        };
        return CommentCollection;
    }(OfficeExtension.ClientObject));
    Visio.CommentCollection = CommentCollection;
    var _typeComment = "Comment";
    var Comment = (function (_super) {
        __extends(Comment, _super);
        function Comment() {
            return _super !== null && _super.apply(this, arguments) || this;
        }
        Object.defineProperty(Comment.prototype, "_className", {
            get: function () {
                return "Comment";
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(Comment.prototype, "_scalarPropertyNames", {
            get: function () {
                return ["author", "text", "date"];
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(Comment.prototype, "_scalarPropertyOriginalNames", {
            get: function () {
                return ["Author", "Text", "Date"];
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(Comment.prototype, "_scalarPropertyUpdateable", {
            get: function () {
                return [true, true, true];
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(Comment.prototype, "author", {
            get: function () {
                _throwIfNotLoaded("author", this._A, _typeComment, this._isNull);
                return this._A;
            },
            set: function (value) {
                this._A = value;
                _invokeSetProperty(this, "Author", value, 0);
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(Comment.prototype, "date", {
            get: function () {
                _throwIfNotLoaded("date", this._D, _typeComment, this._isNull);
                return this._D;
            },
            set: function (value) {
                this._D = value;
                _invokeSetProperty(this, "Date", value, 0);
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(Comment.prototype, "text", {
            get: function () {
                _throwIfNotLoaded("text", this._T, _typeComment, this._isNull);
                return this._T;
            },
            set: function (value) {
                this._T = value;
                _invokeSetProperty(this, "Text", value, 0);
            },
            enumerable: true,
            configurable: true
        });
        Comment.prototype.set = function (properties, options) {
            this._recursivelySet(properties, options, ["author", "text", "date"], [], []);
        };
        Comment.prototype.update = function (properties) {
            this._recursivelyUpdate(properties);
        };
        Comment.prototype._handleResult = function (value) {
            _super.prototype._handleResult.call(this, value);
            if (_isNullOrUndefined(value))
                return;
            var obj = value;
            _fixObjectPathIfNecessary(this, obj);
            if (!_isUndefined(obj["Author"])) {
                this._A = obj["Author"];
            }
            if (!_isUndefined(obj["Date"])) {
                this._D = obj["Date"];
            }
            if (!_isUndefined(obj["Text"])) {
                this._T = obj["Text"];
            }
        };
        Comment.prototype.load = function (options) {
            return _load(this, options);
        };
        Comment.prototype.retrieve = function (option) {
            return _retrieve(this, option);
        };
        Comment.prototype._handleRetrieveResult = function (value, result) {
            _super.prototype._handleRetrieveResult.call(this, value, result);
            _processRetrieveResult(this, value, result);
        };
        Comment.prototype.toJSON = function () {
            return _toJson(this, {
                "author": this._A,
                "date": this._D,
                "text": this._T
            }, {});
        };
        Comment.prototype.setMockData = function (data) {
            _setMockData(this, data);
        };
        Comment.prototype.ensureUnchanged = function (data) {
            _invokeEnsureUnchanged(this, data);
            return;
        };
        return Comment;
    }(OfficeExtension.ClientObject));
    Visio.Comment = Comment;
    var _typeSelection = "Selection";
    var Selection = (function (_super) {
        __extends(Selection, _super);
        function Selection() {
            return _super !== null && _super.apply(this, arguments) || this;
        }
        Object.defineProperty(Selection.prototype, "_className", {
            get: function () {
                return "Selection";
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(Selection.prototype, "_navigationPropertyNames", {
            get: function () {
                return ["shapes"];
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(Selection.prototype, "shapes", {
            get: function () {
                if (!this._S) {
                    this._S = _createPropertyObject(Visio.ShapeCollection, this, "Shapes", true, 4);
                }
                return this._S;
            },
            enumerable: true,
            configurable: true
        });
        Selection.prototype._handleResult = function (value) {
            _super.prototype._handleResult.call(this, value);
            if (_isNullOrUndefined(value))
                return;
            var obj = value;
            _fixObjectPathIfNecessary(this, obj);
            _handleNavigationPropertyResults(this, obj, ["shapes", "Shapes"]);
        };
        Selection.prototype.load = function (options) {
            return _load(this, options);
        };
        Selection.prototype.retrieve = function (option) {
            return _retrieve(this, option);
        };
        Selection.prototype._handleRetrieveResult = function (value, result) {
            _super.prototype._handleRetrieveResult.call(this, value, result);
            _processRetrieveResult(this, value, result);
        };
        Selection.prototype.toJSON = function () {
            return _toJson(this, {}, {
                "shapes": this._S
            });
        };
        return Selection;
    }(OfficeExtension.ClientObject));
    Visio.Selection = Selection;
    var _typeDataVisualizerDiagramCollection = "DataVisualizerDiagramCollection";
    var DataVisualizerDiagramCollection = (function (_super) {
        __extends(DataVisualizerDiagramCollection, _super);
        function DataVisualizerDiagramCollection() {
            return _super !== null && _super.apply(this, arguments) || this;
        }
        Object.defineProperty(DataVisualizerDiagramCollection.prototype, "_className", {
            get: function () {
                return "DataVisualizerDiagramCollection";
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(DataVisualizerDiagramCollection.prototype, "_isCollection", {
            get: function () {
                return true;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(DataVisualizerDiagramCollection.prototype, "items", {
            get: function () {
                _throwIfNotLoaded("items", this.m__items, _typeDataVisualizerDiagramCollection, this._isNull);
                return this.m__items;
            },
            enumerable: true,
            configurable: true
        });
        DataVisualizerDiagramCollection.prototype.add = function (data, settings) {
            return _createMethodObject(Visio.DataVisualizerDiagram, this, "Add", 1, [data, settings], false, true, null, 4);
        };
        DataVisualizerDiagramCollection.prototype.addPreferred = function (data, diagramType) {
            return _createMethodObject(Visio.DataVisualizerDiagram, this, "AddPreferred", 0, [data, diagramType], false, false, null, 0);
        };
        DataVisualizerDiagramCollection.prototype.getCount = function () {
            return _invokeMethod(this, "GetCount", 1, [], 4, 0);
        };
        DataVisualizerDiagramCollection.prototype.getItem = function (key) {
            return _createIndexerObject(Visio.DataVisualizerDiagram, this, [key]);
        };
        DataVisualizerDiagramCollection.prototype.getItemAt = function (index) {
            return _createMethodObject(Visio.DataVisualizerDiagram, this, "GetItemAt", 1, [index], false, false, null, 4);
        };
        DataVisualizerDiagramCollection.prototype.getItemOrNullObject = function (key) {
            return _createMethodObject(Visio.DataVisualizerDiagram, this, "GetItemOrNullObject", 1, [key], false, false, null, 4);
        };
        DataVisualizerDiagramCollection.prototype._handleResult = function (value) {
            _super.prototype._handleResult.call(this, value);
            if (_isNullOrUndefined(value))
                return;
            var obj = value;
            _fixObjectPathIfNecessary(this, obj);
            if (!_isNullOrUndefined(obj[OfficeExtension.Constants.items])) {
                this.m__items = [];
                var _data = obj[OfficeExtension.Constants.items];
                for (var i = 0; i < _data.length; i++) {
                    var _item = _createChildItemObject(Visio.DataVisualizerDiagram, true, this, _data[i], i);
                    _item._handleResult(_data[i]);
                    this.m__items.push(_item);
                }
            }
        };
        DataVisualizerDiagramCollection.prototype.load = function (options) {
            return _load(this, options);
        };
        DataVisualizerDiagramCollection.prototype.retrieve = function (option) {
            return _retrieve(this, option);
        };
        DataVisualizerDiagramCollection.prototype._handleRetrieveResult = function (value, result) {
            var _this = this;
            _super.prototype._handleRetrieveResult.call(this, value, result);
            _processRetrieveResult(this, value, result, function (childItemData, index) { return _createChildItemObject(Visio.DataVisualizerDiagram, true, _this, childItemData, index); });
        };
        DataVisualizerDiagramCollection.prototype.toJSON = function () {
            return _toJson(this, {}, {}, this.m__items);
        };
        DataVisualizerDiagramCollection.prototype.setMockData = function (data) {
            var _this = this;
            _setMockData(this, data, function (childItemData, index) { return _createChildItemObject(Visio.DataVisualizerDiagram, true, _this, childItemData, index); }, function (items) { return _this.m__items = items; });
        };
        return DataVisualizerDiagramCollection;
    }(OfficeExtension.ClientObject));
    Visio.DataVisualizerDiagramCollection = DataVisualizerDiagramCollection;
    var _typeDataVisualizerDiagram = "DataVisualizerDiagram";
    var DataVisualizerDiagram = (function (_super) {
        __extends(DataVisualizerDiagram, _super);
        function DataVisualizerDiagram() {
            return _super !== null && _super.apply(this, arguments) || this;
        }
        Object.defineProperty(DataVisualizerDiagram.prototype, "_className", {
            get: function () {
                return "DataVisualizerDiagram";
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(DataVisualizerDiagram.prototype, "_scalarPropertyNames", {
            get: function () {
                return ["id", "mappings", "type", "dataHeaders"];
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(DataVisualizerDiagram.prototype, "_scalarPropertyOriginalNames", {
            get: function () {
                return ["ID", "Mappings", "Type", "DataHeaders"];
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(DataVisualizerDiagram.prototype, "_navigationPropertyNames", {
            get: function () {
                return ["page"];
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(DataVisualizerDiagram.prototype, "page", {
            get: function () {
                if (!this._P) {
                    this._P = _createPropertyObject(Visio.Page, this, "Page", false, 4);
                }
                return this._P;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(DataVisualizerDiagram.prototype, "dataHeaders", {
            get: function () {
                _throwIfNotLoaded("dataHeaders", this._D, _typeDataVisualizerDiagram, this._isNull);
                return this._D;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(DataVisualizerDiagram.prototype, "id", {
            get: function () {
                _throwIfNotLoaded("id", this._I, _typeDataVisualizerDiagram, this._isNull);
                return this._I;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(DataVisualizerDiagram.prototype, "mappings", {
            get: function () {
                _throwIfNotLoaded("mappings", this._M, _typeDataVisualizerDiagram, this._isNull);
                return this._M;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(DataVisualizerDiagram.prototype, "type", {
            get: function () {
                _throwIfNotLoaded("type", this._T, _typeDataVisualizerDiagram, this._isNull);
                return this._T;
            },
            enumerable: true,
            configurable: true
        });
        DataVisualizerDiagram.prototype.set = function (properties, options) {
            this._recursivelySet(properties, options, [], ["page"], []);
        };
        DataVisualizerDiagram.prototype["delete"] = function () {
            _invokeMethod(this, "Delete", 1, [], 4, 0);
        };
        DataVisualizerDiagram.prototype.getDataColumnValuesAsString = function (columnName) {
            return _invokeMethod(this, "GetDataColumnValuesAsString", 1, [columnName], 4, 0);
        };
        DataVisualizerDiagram.prototype.setConnection = function (connectionInfo) {
            _invokeMethod(this, "SetConnection", 1, [connectionInfo], 4, 0);
        };
        DataVisualizerDiagram.prototype.update = function (data, mappings, ignoreConflicts) {
            _invokeMethod(this, "Update", 1, [data, mappings, ignoreConflicts], 4, 0);
        };
        DataVisualizerDiagram.prototype.updateMappings = function (mappings) {
            _invokeMethod(this, "UpdateMappings", 1, [mappings], 4, 0);
        };
        DataVisualizerDiagram.prototype._handleResult = function (value) {
            _super.prototype._handleResult.call(this, value);
            if (_isNullOrUndefined(value))
                return;
            var obj = value;
            _fixObjectPathIfNecessary(this, obj);
            if (!_isUndefined(obj["DataHeaders"])) {
                this._D = obj["DataHeaders"];
            }
            if (!_isUndefined(obj["ID"])) {
                this._I = obj["ID"];
            }
            if (!_isUndefined(obj["Mappings"])) {
                this._M = obj["Mappings"];
            }
            if (!_isUndefined(obj["Type"])) {
                this._T = obj["Type"];
            }
            _handleNavigationPropertyResults(this, obj, ["page", "Page"]);
        };
        DataVisualizerDiagram.prototype.load = function (options) {
            return _load(this, options);
        };
        DataVisualizerDiagram.prototype.retrieve = function (option) {
            return _retrieve(this, option);
        };
        DataVisualizerDiagram.prototype._handleRetrieveResult = function (value, result) {
            _super.prototype._handleRetrieveResult.call(this, value, result);
            _processRetrieveResult(this, value, result);
        };
        DataVisualizerDiagram.prototype.toJSON = function () {
            return _toJson(this, {
                "dataHeaders": this._D,
                "id": this._I,
                "mappings": this._M,
                "type": this._T
            }, {
                "page": this._P
            });
        };
        DataVisualizerDiagram.prototype.setMockData = function (data) {
            _setMockData(this, data);
        };
        DataVisualizerDiagram.prototype.ensureUnchanged = function (data) {
            _invokeEnsureUnchanged(this, data);
            return;
        };
        return DataVisualizerDiagram;
    }(OfficeExtension.ClientObject));
    Visio.DataVisualizerDiagram = DataVisualizerDiagram;
    var OverlayHorizontalAlignment;
    (function (OverlayHorizontalAlignment) {
        OverlayHorizontalAlignment["left"] = "Left";
        OverlayHorizontalAlignment["center"] = "Center";
        OverlayHorizontalAlignment["right"] = "Right";
    })(OverlayHorizontalAlignment = Visio.OverlayHorizontalAlignment || (Visio.OverlayHorizontalAlignment = {}));
    var OverlayVerticalAlignment;
    (function (OverlayVerticalAlignment) {
        OverlayVerticalAlignment["top"] = "Top";
        OverlayVerticalAlignment["middle"] = "Middle";
        OverlayVerticalAlignment["bottom"] = "Bottom";
    })(OverlayVerticalAlignment = Visio.OverlayVerticalAlignment || (Visio.OverlayVerticalAlignment = {}));
    var OverlayType;
    (function (OverlayType) {
        OverlayType["text"] = "Text";
        OverlayType["image"] = "Image";
        OverlayType["html"] = "Html";
    })(OverlayType = Visio.OverlayType || (Visio.OverlayType = {}));
    var ToolBarType;
    (function (ToolBarType) {
        ToolBarType["commandBar"] = "CommandBar";
        ToolBarType["pageNavigationBar"] = "PageNavigationBar";
        ToolBarType["statusBar"] = "StatusBar";
    })(ToolBarType = Visio.ToolBarType || (Visio.ToolBarType = {}));
    var DataVisualizerDiagramResultType;
    (function (DataVisualizerDiagramResultType) {
        DataVisualizerDiagramResultType["success"] = "Success";
        DataVisualizerDiagramResultType["unexpected"] = "Unexpected";
        DataVisualizerDiagramResultType["validationError"] = "ValidationError";
        DataVisualizerDiagramResultType["conflictError"] = "ConflictError";
    })(DataVisualizerDiagramResultType = Visio.DataVisualizerDiagramResultType || (Visio.DataVisualizerDiagramResultType = {}));
    var DataVisualizerDiagramOperationType;
    (function (DataVisualizerDiagramOperationType) {
        DataVisualizerDiagramOperationType["unknown"] = "Unknown";
        DataVisualizerDiagramOperationType["create"] = "Create";
        DataVisualizerDiagramOperationType["updateMappings"] = "UpdateMappings";
        DataVisualizerDiagramOperationType["updateData"] = "UpdateData";
        DataVisualizerDiagramOperationType["update"] = "Update";
        DataVisualizerDiagramOperationType["delete"] = "Delete";
    })(DataVisualizerDiagramOperationType = Visio.DataVisualizerDiagramOperationType || (Visio.DataVisualizerDiagramOperationType = {}));
    var DataVisualizerDiagramType;
    (function (DataVisualizerDiagramType) {
        DataVisualizerDiagramType["unknown"] = "Unknown";
        DataVisualizerDiagramType["basicFlowchart"] = "BasicFlowchart";
        DataVisualizerDiagramType["crossFunctionalFlowchart_Horizontal"] = "CrossFunctionalFlowchart_Horizontal";
        DataVisualizerDiagramType["crossFunctionalFlowchart_Vertical"] = "CrossFunctionalFlowchart_Vertical";
        DataVisualizerDiagramType["audit"] = "Audit";
        DataVisualizerDiagramType["orgChart"] = "OrgChart";
        DataVisualizerDiagramType["network"] = "Network";
    })(DataVisualizerDiagramType = Visio.DataVisualizerDiagramType || (Visio.DataVisualizerDiagramType = {}));
    var ColumnType;
    (function (ColumnType) {
        ColumnType["unknown"] = "Unknown";
        ColumnType["string"] = "String";
        ColumnType["number"] = "Number";
        ColumnType["date"] = "Date";
        ColumnType["currency"] = "Currency";
    })(ColumnType = Visio.ColumnType || (Visio.ColumnType = {}));
    var DataSourceType;
    (function (DataSourceType) {
        DataSourceType["unknown"] = "Unknown";
        DataSourceType["excel"] = "Excel";
    })(DataSourceType = Visio.DataSourceType || (Visio.DataSourceType = {}));
    var CrossFunctionalFlowchartOrientation;
    (function (CrossFunctionalFlowchartOrientation) {
        CrossFunctionalFlowchartOrientation["horizontal"] = "Horizontal";
        CrossFunctionalFlowchartOrientation["vertical"] = "Vertical";
    })(CrossFunctionalFlowchartOrientation = Visio.CrossFunctionalFlowchartOrientation || (Visio.CrossFunctionalFlowchartOrientation = {}));
    var LayoutVariant;
    (function (LayoutVariant) {
        LayoutVariant["unknown"] = "Unknown";
        LayoutVariant["pageDefault"] = "PageDefault";
        LayoutVariant["flowchart_TopToBottom"] = "Flowchart_TopToBottom";
        LayoutVariant["flowchart_BottomToTop"] = "Flowchart_BottomToTop";
        LayoutVariant["flowchart_LeftToRight"] = "Flowchart_LeftToRight";
        LayoutVariant["flowchart_RightToLeft"] = "Flowchart_RightToLeft";
        LayoutVariant["wideTree_DownThenRight"] = "WideTree_DownThenRight";
        LayoutVariant["wideTree_DownThenLeft"] = "WideTree_DownThenLeft";
        LayoutVariant["wideTree_RightThenDown"] = "WideTree_RightThenDown";
        LayoutVariant["wideTree_LeftThenDown"] = "WideTree_LeftThenDown";
    })(LayoutVariant = Visio.LayoutVariant || (Visio.LayoutVariant = {}));
    var DataValidationErrorType;
    (function (DataValidationErrorType) {
        DataValidationErrorType["none"] = "None";
        DataValidationErrorType["columnNotMapped"] = "ColumnNotMapped";
        DataValidationErrorType["uniqueIdColumnError"] = "UniqueIdColumnError";
        DataValidationErrorType["swimlaneColumnError"] = "SwimlaneColumnError";
        DataValidationErrorType["delimiterError"] = "DelimiterError";
        DataValidationErrorType["connectorColumnError"] = "ConnectorColumnError";
        DataValidationErrorType["connectorColumnMappedElsewhere"] = "ConnectorColumnMappedElsewhere";
        DataValidationErrorType["connectorLabelColumnMappedElsewhere"] = "ConnectorLabelColumnMappedElsewhere";
        DataValidationErrorType["connectorColumnAndConnectorLabelMappedElsewhere"] = "ConnectorColumnAndConnectorLabelMappedElsewhere";
    })(DataValidationErrorType = Visio.DataValidationErrorType || (Visio.DataValidationErrorType = {}));
    var ConnectorDirection;
    (function (ConnectorDirection) {
        ConnectorDirection["fromTarget"] = "FromTarget";
        ConnectorDirection["toTarget"] = "ToTarget";
    })(ConnectorDirection = Visio.ConnectorDirection || (Visio.ConnectorDirection = {}));
    var TaskPaneType;
    (function (TaskPaneType) {
        TaskPaneType["none"] = "None";
        TaskPaneType["dataVisualizerProcessMappings"] = "DataVisualizerProcessMappings";
        TaskPaneType["dataVisualizerOrgChartMappings"] = "DataVisualizerOrgChartMappings";
    })(TaskPaneType = Visio.TaskPaneType || (Visio.TaskPaneType = {}));
    var EventType;
    (function (EventType) {
        EventType["dataVisualizerDiagramOperationCompleted"] = "DataVisualizerDiagramOperationCompleted";
    })(EventType = Visio.EventType || (Visio.EventType = {}));
    var ErrorCodes;
    (function (ErrorCodes) {
        ErrorCodes["accessDenied"] = "AccessDenied";
        ErrorCodes["generalException"] = "GeneralException";
        ErrorCodes["invalidArgument"] = "InvalidArgument";
        ErrorCodes["itemNotFound"] = "ItemNotFound";
        ErrorCodes["notImplemented"] = "NotImplemented";
        ErrorCodes["unsupportedOperation"] = "UnsupportedOperation";
    })(ErrorCodes = Visio.ErrorCodes || (Visio.ErrorCodes = {}));
    var Interfaces;
    (function (Interfaces) {
    })(Interfaces = Visio.Interfaces || (Visio.Interfaces = {}));
})(Visio || (Visio = {}));
Object.defineProperty(OfficeExtension.SessionBase, "_overrideSession", {
    get: function () {
        if (this._overrideSessionInternal) {
            return this._overrideSessionInternal;
        }
        if (OfficeExtension.ClientRequestContext) {
            return (OfficeExtension.ClientRequestContext)._overrideSession;
        }
        return undefined;
    },
    set: function (value) {
        this._overrideSessionInternal = value;
    },
    enumerable: true,
    configurable: true
});
var Visio;
(function (Visio) {
    var RequestContext = (function (_super) {
        __extends(RequestContext, _super);
        function RequestContext(url) {
            var _this = _super.call(this, url) || this;
            _this.m_document = new Visio.Document(_this, OfficeExtension.ObjectPathFactory.createGlobalObjectObjectPath(_this));
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
        return RequestContext;
    }(OfficeCore.RequestContext));
    Visio.RequestContext = RequestContext;
    function run(arg1, arg2, arg3) {
        return OfficeExtension.ClientRequestContext._runBatch("Visio.run", arguments, function (requestInfo) {
            var ret = new Visio.RequestContext(requestInfo);
            return ret;
        });
    }
    Visio.run = run;
})(Visio || (Visio = {}));
OfficeExtension.Utility._doApiNotSupportedCheck = true;

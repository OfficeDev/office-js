var BrowserAuth =
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
/******/ 	__webpack_require__.p = "";
/******/
/******/
/******/ 	// Load entry module and return exports
/******/ 	return __webpack_require__(__webpack_require__.s = 0);
/******/ })
/************************************************************************/
/******/ ({

/***/ "./node_modules/@azure/msal-browser/dist/_virtual/_tslib.js":
/*!******************************************************************!*\
  !*** ./node_modules/@azure/msal-browser/dist/_virtual/_tslib.js ***!
  \******************************************************************/
/*! exports provided: __assign, __awaiter, __extends, __generator, __read, __spread */
/***/ (function(module, __webpack_exports__, __webpack_require__) {

"use strict";
eval("__webpack_require__.r(__webpack_exports__);\n/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, \"__assign\", function() { return __assign; });\n/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, \"__awaiter\", function() { return __awaiter; });\n/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, \"__extends\", function() { return __extends; });\n/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, \"__generator\", function() { return __generator; });\n/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, \"__read\", function() { return __read; });\n/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, \"__spread\", function() { return __spread; });\n/*! @azure/msal-browser v2.19.0 2021-11-02 */\n\n/*! *****************************************************************************\r\nCopyright (c) Microsoft Corporation.\r\n\r\nPermission to use, copy, modify, and/or distribute this software for any\r\npurpose with or without fee is hereby granted.\r\n\r\nTHE SOFTWARE IS PROVIDED \"AS IS\" AND THE AUTHOR DISCLAIMS ALL WARRANTIES WITH\r\nREGARD TO THIS SOFTWARE INCLUDING ALL IMPLIED WARRANTIES OF MERCHANTABILITY\r\nAND FITNESS. IN NO EVENT SHALL THE AUTHOR BE LIABLE FOR ANY SPECIAL, DIRECT,\r\nINDIRECT, OR CONSEQUENTIAL DAMAGES OR ANY DAMAGES WHATSOEVER RESULTING FROM\r\nLOSS OF USE, DATA OR PROFITS, WHETHER IN AN ACTION OF CONTRACT, NEGLIGENCE OR\r\nOTHER TORTIOUS ACTION, ARISING OUT OF OR IN CONNECTION WITH THE USE OR\r\nPERFORMANCE OF THIS SOFTWARE.\r\n***************************************************************************** */\r\n/* global Reflect, Promise */\r\n\r\nvar extendStatics = function(d, b) {\r\n    extendStatics = Object.setPrototypeOf ||\r\n        ({ __proto__: [] } instanceof Array && function (d, b) { d.__proto__ = b; }) ||\r\n        function (d, b) { for (var p in b) if (Object.prototype.hasOwnProperty.call(b, p)) d[p] = b[p]; };\r\n    return extendStatics(d, b);\r\n};\r\n\r\nfunction __extends(d, b) {\r\n    extendStatics(d, b);\r\n    function __() { this.constructor = d; }\r\n    d.prototype = b === null ? Object.create(b) : (__.prototype = b.prototype, new __());\r\n}\r\n\r\nvar __assign = function() {\r\n    __assign = Object.assign || function __assign(t) {\r\n        for (var s, i = 1, n = arguments.length; i < n; i++) {\r\n            s = arguments[i];\r\n            for (var p in s) if (Object.prototype.hasOwnProperty.call(s, p)) t[p] = s[p];\r\n        }\r\n        return t;\r\n    };\r\n    return __assign.apply(this, arguments);\r\n};\r\n\r\nfunction __awaiter(thisArg, _arguments, P, generator) {\r\n    function adopt(value) { return value instanceof P ? value : new P(function (resolve) { resolve(value); }); }\r\n    return new (P || (P = Promise))(function (resolve, reject) {\r\n        function fulfilled(value) { try { step(generator.next(value)); } catch (e) { reject(e); } }\r\n        function rejected(value) { try { step(generator[\"throw\"](value)); } catch (e) { reject(e); } }\r\n        function step(result) { result.done ? resolve(result.value) : adopt(result.value).then(fulfilled, rejected); }\r\n        step((generator = generator.apply(thisArg, _arguments || [])).next());\r\n    });\r\n}\r\n\r\nfunction __generator(thisArg, body) {\r\n    var _ = { label: 0, sent: function() { if (t[0] & 1) throw t[1]; return t[1]; }, trys: [], ops: [] }, f, y, t, g;\r\n    return g = { next: verb(0), \"throw\": verb(1), \"return\": verb(2) }, typeof Symbol === \"function\" && (g[Symbol.iterator] = function() { return this; }), g;\r\n    function verb(n) { return function (v) { return step([n, v]); }; }\r\n    function step(op) {\r\n        if (f) throw new TypeError(\"Generator is already executing.\");\r\n        while (_) try {\r\n            if (f = 1, y && (t = op[0] & 2 ? y[\"return\"] : op[0] ? y[\"throw\"] || ((t = y[\"return\"]) && t.call(y), 0) : y.next) && !(t = t.call(y, op[1])).done) return t;\r\n            if (y = 0, t) op = [op[0] & 2, t.value];\r\n            switch (op[0]) {\r\n                case 0: case 1: t = op; break;\r\n                case 4: _.label++; return { value: op[1], done: false };\r\n                case 5: _.label++; y = op[1]; op = [0]; continue;\r\n                case 7: op = _.ops.pop(); _.trys.pop(); continue;\r\n                default:\r\n                    if (!(t = _.trys, t = t.length > 0 && t[t.length - 1]) && (op[0] === 6 || op[0] === 2)) { _ = 0; continue; }\r\n                    if (op[0] === 3 && (!t || (op[1] > t[0] && op[1] < t[3]))) { _.label = op[1]; break; }\r\n                    if (op[0] === 6 && _.label < t[1]) { _.label = t[1]; t = op; break; }\r\n                    if (t && _.label < t[2]) { _.label = t[2]; _.ops.push(op); break; }\r\n                    if (t[2]) _.ops.pop();\r\n                    _.trys.pop(); continue;\r\n            }\r\n            op = body.call(thisArg, _);\r\n        } catch (e) { op = [6, e]; y = 0; } finally { f = t = 0; }\r\n        if (op[0] & 5) throw op[1]; return { value: op[0] ? op[1] : void 0, done: true };\r\n    }\r\n}\r\n\r\nfunction __read(o, n) {\r\n    var m = typeof Symbol === \"function\" && o[Symbol.iterator];\r\n    if (!m) return o;\r\n    var i = m.call(o), r, ar = [], e;\r\n    try {\r\n        while ((n === void 0 || n-- > 0) && !(r = i.next()).done) ar.push(r.value);\r\n    }\r\n    catch (error) { e = { error: error }; }\r\n    finally {\r\n        try {\r\n            if (r && !r.done && (m = i[\"return\"])) m.call(i);\r\n        }\r\n        finally { if (e) throw e.error; }\r\n    }\r\n    return ar;\r\n}\r\n\r\nfunction __spread() {\r\n    for (var ar = [], i = 0; i < arguments.length; i++)\r\n        ar = ar.concat(__read(arguments[i]));\r\n    return ar;\r\n}\n\n\n\n\n//# sourceURL=webpack://BrowserAuth/./node_modules/@azure/msal-browser/dist/_virtual/_tslib.js?");

/***/ }),

/***/ "./node_modules/@azure/msal-browser/dist/app/ClientApplication.js":
/*!************************************************************************!*\
  !*** ./node_modules/@azure/msal-browser/dist/app/ClientApplication.js ***!
  \************************************************************************/
/*! exports provided: ClientApplication */
/***/ (function(module, __webpack_exports__, __webpack_require__) {

"use strict";
eval("__webpack_require__.r(__webpack_exports__);\n/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, \"ClientApplication\", function() { return ClientApplication; });\n/* harmony import */ var _virtual_tslib_js__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! ../_virtual/_tslib.js */ \"./node_modules/@azure/msal-browser/dist/_virtual/_tslib.js\");\n/* harmony import */ var _crypto_CryptoOps_js__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(/*! ../crypto/CryptoOps.js */ \"./node_modules/@azure/msal-browser/dist/crypto/CryptoOps.js\");\n/* harmony import */ var _azure_msal_common__WEBPACK_IMPORTED_MODULE_2__ = __webpack_require__(/*! @azure/msal-common */ \"./node_modules/@azure/msal-common/dist/index.js\");\n/* harmony import */ var _cache_BrowserCacheManager_js__WEBPACK_IMPORTED_MODULE_3__ = __webpack_require__(/*! ../cache/BrowserCacheManager.js */ \"./node_modules/@azure/msal-browser/dist/cache/BrowserCacheManager.js\");\n/* harmony import */ var _config_Configuration_js__WEBPACK_IMPORTED_MODULE_4__ = __webpack_require__(/*! ../config/Configuration.js */ \"./node_modules/@azure/msal-browser/dist/config/Configuration.js\");\n/* harmony import */ var _utils_BrowserConstants_js__WEBPACK_IMPORTED_MODULE_5__ = __webpack_require__(/*! ../utils/BrowserConstants.js */ \"./node_modules/@azure/msal-browser/dist/utils/BrowserConstants.js\");\n/* harmony import */ var _utils_BrowserUtils_js__WEBPACK_IMPORTED_MODULE_6__ = __webpack_require__(/*! ../utils/BrowserUtils.js */ \"./node_modules/@azure/msal-browser/dist/utils/BrowserUtils.js\");\n/* harmony import */ var _packageMetadata_js__WEBPACK_IMPORTED_MODULE_7__ = __webpack_require__(/*! ../packageMetadata.js */ \"./node_modules/@azure/msal-browser/dist/packageMetadata.js\");\n/* harmony import */ var _event_EventType_js__WEBPACK_IMPORTED_MODULE_8__ = __webpack_require__(/*! ../event/EventType.js */ \"./node_modules/@azure/msal-browser/dist/event/EventType.js\");\n/* harmony import */ var _error_BrowserConfigurationAuthError_js__WEBPACK_IMPORTED_MODULE_9__ = __webpack_require__(/*! ../error/BrowserConfigurationAuthError.js */ \"./node_modules/@azure/msal-browser/dist/error/BrowserConfigurationAuthError.js\");\n/* harmony import */ var _event_EventHandler_js__WEBPACK_IMPORTED_MODULE_10__ = __webpack_require__(/*! ../event/EventHandler.js */ \"./node_modules/@azure/msal-browser/dist/event/EventHandler.js\");\n/* harmony import */ var _interaction_client_PopupClient_js__WEBPACK_IMPORTED_MODULE_11__ = __webpack_require__(/*! ../interaction_client/PopupClient.js */ \"./node_modules/@azure/msal-browser/dist/interaction_client/PopupClient.js\");\n/* harmony import */ var _interaction_client_RedirectClient_js__WEBPACK_IMPORTED_MODULE_12__ = __webpack_require__(/*! ../interaction_client/RedirectClient.js */ \"./node_modules/@azure/msal-browser/dist/interaction_client/RedirectClient.js\");\n/* harmony import */ var _interaction_client_SilentIframeClient_js__WEBPACK_IMPORTED_MODULE_13__ = __webpack_require__(/*! ../interaction_client/SilentIframeClient.js */ \"./node_modules/@azure/msal-browser/dist/interaction_client/SilentIframeClient.js\");\n/* harmony import */ var _interaction_client_SilentRefreshClient_js__WEBPACK_IMPORTED_MODULE_14__ = __webpack_require__(/*! ../interaction_client/SilentRefreshClient.js */ \"./node_modules/@azure/msal-browser/dist/interaction_client/SilentRefreshClient.js\");\n/* harmony import */ var _cache_TokenCache_js__WEBPACK_IMPORTED_MODULE_15__ = __webpack_require__(/*! ../cache/TokenCache.js */ \"./node_modules/@azure/msal-browser/dist/cache/TokenCache.js\");\n/*! @azure/msal-browser v2.19.0 2021-11-02 */\n\n\n\n\n\n\n\n\n\n\n\n\n\n\n\n\n\n\n/*\r\n * Copyright (c) Microsoft Corporation. All rights reserved.\r\n * Licensed under the MIT License.\r\n */\r\nvar ClientApplication = /** @class */ (function () {\r\n    /**\r\n     * @constructor\r\n     * Constructor for the PublicClientApplication used to instantiate the PublicClientApplication object\r\n     *\r\n     * Important attributes in the Configuration object for auth are:\r\n     * - clientID: the application ID of your application. You can obtain one by registering your application with our Application registration portal : https://portal.azure.com/#blade/Microsoft_AAD_IAM/ActiveDirectoryMenuBlade/RegisteredAppsPreview\r\n     * - authority: the authority URL for your application.\r\n     * - redirect_uri: the uri of your application registered in the portal.\r\n     *\r\n     * In Azure AD, authority is a URL indicating the Azure active directory that MSAL uses to obtain tokens.\r\n     * It is of the form https://login.microsoftonline.com/{Enter_the_Tenant_Info_Here}\r\n     * If your application supports Accounts in one organizational directory, replace \"Enter_the_Tenant_Info_Here\" value with the Tenant Id or Tenant name (for example, contoso.microsoft.com).\r\n     * If your application supports Accounts in any organizational directory, replace \"Enter_the_Tenant_Info_Here\" value with organizations.\r\n     * If your application supports Accounts in any organizational directory and personal Microsoft accounts, replace \"Enter_the_Tenant_Info_Here\" value with common.\r\n     * To restrict support to Personal Microsoft accounts only, replace \"Enter_the_Tenant_Info_Here\" value with consumers.\r\n     *\r\n     * In Azure B2C, authority is of the form https://{instance}/tfp/{tenant}/{policyName
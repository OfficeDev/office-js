(function webpackUniversalModuleDefinition(root, factory) {
	if(typeof exports === 'object' && typeof module === 'object')
		module.exports = factory();
	else if(typeof define === 'function' && define.amd)
		define("Implicit", [], factory);
	else if(typeof exports === 'object')
		exports["Implicit"] = factory();
	else
		root["Implicit"] = factory();
})(window, function() {
return /******/ (function(modules) { // webpackBootstrap
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

/***/ "./packages/Microsoft.Office.WebAuth.Implicit/lib/api.js":
/***/ (function(module, exports, __webpack_require__) {

"use strict";


Object.defineProperty(exports, "__esModule", {
  value: true
});
exports.logUserAction = logUserAction;
exports.logActivity = logActivity;
exports.sendTelemetryEvent = sendTelemetryEvent;
exports.sendActivityEvent = sendActivityEvent;
exports.sendOtelEvent = sendOtelEvent;
exports.sendUserActionEvent = sendUserActionEvent;
exports.addNamespaceMapping = addNamespaceMapping;
exports.setEnabledState = setEnabledState;
exports.shutdown = shutdown;
exports.registerEventHandler = registerEventHandler; // Assume that telemetry is disabled and simply drop events on the floor unless the developer called initialize(true /*enabled*/).
// This should work well for component / unittest environments since nobody will end up listening to the events.
// The alternative is to cache them, but given that nobody will ever process them it might cause issues since the events would be cached forever.

var telemetryEnabled = false;
var events = [];
var eventHandler;
var numberOfDroppedEvents = 0;
var maxQueueSize = 20000;
var unknownStr = 'Unknown'; // Primary consumer public API
// ========================================================================================================================
// Call LogUserAction for logging a user action to Otel.
// This is similar to the bSqm actions that used to be logged earlier (deprecated now).
// Make sure you read the documentation below for userActionName and the Kusto table name implications.
// userActionName: Name of the user action, this should come from your app's commands,
//     for example: OneNoteCommands in office-online-ui\packages\onenote-online-ux\src\store\OneNoteCommands.ts (https://office.visualstudio.com/OC/_git/office-online-ui?path=%2Fpackages%2Fonenote-online-ux%2Fsrc%2Fstore%2FOneNoteCommands.ts&version=GBmaster)
//     Note that the userActionName will be the name of your table in Aria Kusto. So if 'ABC' is passed in for userActionName, the table in Kusto will be called Office_OneNote_Online_UserAction_ABC (or generically speaking Office_{AppName}_Online_UserAction_ABC )
//     Look at Kusto connection https://kusto.aria.microsoft.com and databases Office Word Online or Office OneNote Online, etc. and look at *UserAction* tables.
// success: Status of the user action (success is true, failure is false).
// parentNameStr: parent surface of the user action (example, tabView, tabHelp, Layout, etc).
// inputMethod: how the user action was performed (for example, via keyboard, or mouse, touch, etc.)
//             See the enum in /packages/app-commanding-ui/src/UISurfaces/controls/InputMethod.ts
//             Pass in this param as:  InputMethod.Keyboard.toString() instead of passing in "Keyboard"
// uiLocation: the surface where the user action was initiated from (example, ribbon, FileMenu, TellMe, etc).
//             See enum in /packages/app-commanding-ui/src/UISurfaces/controls/UILocation.ts
//             Pass in this param as:  UILocation.SingleLineRibbon.toString() instead of passing in "SingleLineRibbon"
// durationMsec: the time taken by the action (if relevant to the action)
// dataFieldArr: These are custom fields that you may want to add for your user action.
//               Example: InsertTable action may log custom data fields such as rowSize and colSize of the table inserted.
//                      Or in Excel, a cell related action may log the x and y coordinates of the cell.
// Note that things such as sessionID, data center, etc will be added to all user action logs.

function logUserAction(userActionName, success, parentNameStr, inputMethod, uiLocation, durationMsec, dataFieldArr) {
  if (success === void 0) {
    success = true;
  }

  if (parentNameStr === void 0) {
    parentNameStr = unknownStr;
  }

  if (inputMethod === void 0) {
    inputMethod = unknownStr;
  }

  if (uiLocation === void 0) {
    uiLocation = unknownStr;
  }

  if (durationMsec === void 0) {
    durationMsec = 0;
  }

  if (dataFieldArr === void 0) {
    dataFieldArr = [];
  } // passing null for 'name' field, which is the event table name. We will determine that in sendUserAction in full\api.ts as there we know what app we are, and hence what the event table name is


  sendUserActionEvent({
    name: null,
    actionName: userActionName,
    commandSurface: uiLocation,
    parentName: parentNameStr,
    triggerMethod: inputMethod,
    durationMs: durationMsec,
    succeeded: success,
    dataFields: dataFieldArr
  });
} //////////////////////////////////////////////////////////////////////////////////////////////////
// Call logActivity for logging an activity to Otel.
// This will be logged under Office {App} Online Data tenant
// For example, if your activity name is "ABC",
// it will go to a table called "Office_Word_Online_Data_Activity_ABC" for Word or "Office_OneNote_Online_Data_Activity_ABC" for OneNote.
// activityName: name of activity being logged
// success: Status of the activity (success is true, failure is false).
// durationMsec: the time taken by the action (if relevant to the action)
// dataFieldArr: These are custom fields that you may want to add for your activity, and will be added as columns to the activity table.
//               Example: dataFields has typingSpeedPerSec (integer) and dayOfWeek (string) in it, the activity table for this particular activity will contain these two custom fields.
// Note that things such as sessionID, data center, etc will be added to all user action logs.


function logActivity(activityName, success, durationMsec, dataFieldArr) {
  if (success === void 0) {
    success = true;
  }

  if (durationMsec === void 0) {
    durationMsec = 0;
  }

  if (dataFieldArr === void 0) {
    dataFieldArr = [];
  }

  sendActivityEvent({
    name: activityName,
    succeeded: success,
    durationMs: durationMsec,
    dataFields: dataFieldArr
  });
} // Call LogNonUserAction for logging a non user action to Otel.
// This is
// activityName: Name of the action (non user)
//     Note that the userActionName will be what your table will be named in Aria Kusto. So if 'ABC' is passed in for non user action, the table in Kusto will be called Office_OneNote_Online_NonUser_ABC
//     Look at Kusto connection https://kusto.aria.microsoft.com and databases Office Word Online or Office OneNote Online, etc. and look at *NonUser* tables.
// succeeded: Status of the user action (success is true, failure is false).
// parentName: parent surface of the user action (example, tabView, tabHelp, Layout, etc).
// inputMethod: how the user action was performed (for example, via keyboard, or mouse, touch, etc.)
// uiLocation: the surface where the user action was initiated from (example, ribbon, FileMenu, TellMe, etc)
// startTime: start time of the activity
// endTime: end time of the activity
// dataFields: These are custom fields that you may want to add for your user action.
//             Example: InsertTable action may log custom data fields such as rowSize and colSize of the table inserted.
//                      Or in Excel, a cell related action may log the x and y coordinates of the cell.
//             Note that things such as sessionID, data center, etc will be added to all user action logs.

/*
Being commented out as we dont think we should expose this API. But code is here in case someone educates us on why this should be exposed (it is being used historically in scriptsharp in OtelActionListener.cs (WsaListener.cs))

export const nonUserActionPrefix = 'non_user_action_'; // used in full\api.ts sendActivityEvent to determine if an activity is non user action or a regular activity

If this code is reinstated, then we need to add the following in full\api.ts:

import { nonUserActionPrefix } from '../core';
nonUserActionEventName = 'Office.Online.NonUserAction';
nonUserActionEventName = `Office.${settings.alwaysOnMetadata.name}.Online.NonUser.`;

function ContainsNonUserActionPrefix(eventName: string): boolean {
  return eventName.indexOf(nonUserActionPrefix) == 0;
}

And these lines in sendActivity function in full\api.ts

  if (event.name != null) {
    if (ContainsNonUserActionPrefix(event.name)) {
      event.name = nonUserActionEventName + event.name.substring(nonUserActionPrefix.length);
    }
  }

export function LogNonUserAction(
  activityName: string,
  succeeded: boolean = true,
  parentName: string = unknownStr,
  inputMethod: InputMethod = InputMethod.Unknown,
  uiLocation: UILocation | null = null,
  startTime: number = 0,
  endTime: number = 0,
  dataFields: DataField[] = []
) {
  let durationMs: number = Math.max(endTime - startTime, 0);

  dataFields!.push({ name: 'ParentName', string: parentName != null ? parentName : unknownStr });
  dataFields!.push({ name: 'TriggerMethod', string: inputMethod != null ? inputMethod.toString() : unknownStr });
  dataFields!.push({ name: 'CommandSurface', string: uiLocation != null ? uiLocation.toString() : unknownStr });
  dataFields!.push({ name: 'ActionName', string: activityName });
  dataFields!.push({ name: 'StartTime', double: startTime });
  dataFields!.push({ name: 'EndTime', double: endTime });
  dataFields!.push({ name: 'Succeeded', bool: succeeded });

  // add a sentinel prefix to activity name such that we know we need to add the non user activity event table name (instead of a regular activity event table name) in  sendActivityEvent in full\api.ts as there we know what app we are, and hence what the event table name is
  let activityNameWithPrefix = nonUserActionPrefix;
  if (activityName != null) {
    activityNameWithPrefix = activityNameWithPrefix.concat(activityName);
  }

  sendActivityEvent({
    name: activityNameWithPrefix,
    dataFields: dataFields,
    durationMs: durationMs,
    succeeded: succeeded
  });
}
*/


function sendTelemetryEvent(event) {
  raiseEvent({
    kind: 'event',
    event: event,
    timestamp: new Date().getTime()
  });
}

function sendActivityEvent(event) {
  raiseEvent({
    kind: 'activity',
    event: event,
    timestamp: new Date().getTime()
  });
}

function sendOtelEvent(event) {
  raiseEvent({
    kind: 'otel',
    event: event
  });
}

function sendUserActionEvent(event) {
  raiseEvent({
    kind: 'action',
    event: event,
    timestamp: new Date().getTime()
  });
}

function addNamespaceMapping(namespace, ariaTenantToken) {
  raiseEvent({
    kind: 'addNamespaceMapping',
    namespace: namespace,
    ariaTenantToken: ariaTenantToken
  });
} // Initialization / Shutdown
// ========================================================================================================================


function setEnabledState(enabled) {
  telemetryEnabled = enabled; // If the caller disables the queue, be sure to drop all of the outstanding events.
  // This can happen in cases where the slice with event processor functionality failed to load.

  if (!telemetryEnabled) {
    events = [];
  }
}

function shutdown() {
  raiseEvent({
    kind: 'shutdown'
  });
  return events.length + numberOfDroppedEvents;
}

function registerEventHandler(handler) {
  eventHandler = handler; // Then go through the queue and process the events in the order in which they were received
  // VSO.2533164: Push batch event processing to otelFull and add a lightweight queue

  events.forEach(function (event) {
    return raiseEvent(event);
  });
  events = [];
}

function raiseEvent(event) {
  if (!telemetryEnabled) {
    return;
  }

  if (eventHandler) {
    eventHandler(event);
  } else {
    if (events.length <= maxQueueSize) {
      events.push(event);
    } else {
      numberOfDroppedEvents += 1;
    }
  }
}


/***/ }),

/***/ "./packages/Microsoft.Office.WebAuth.Implic
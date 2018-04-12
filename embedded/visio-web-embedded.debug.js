/* Visio Web-specific API library */
/* Version: 16.0.9103.3000 */

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

var __extends=(this && this.__extends) || function (d, b) {
	for (var p in b) if (b.hasOwnProperty(p)) d[p]=b[p];
	function __() { this.constructor=d; }
	d.prototype=b===null ? Object.create(b) : (__.prototype=b.prototype, new __());
};
var OfficeExtension;
(function (OfficeExtension) {
	var Action=(function () {
		function Action(actionInfo, isWriteOperation, isRestrictedResourceAccess) {
			this.m_actionInfo=actionInfo;
			this.m_isWriteOperation=isWriteOperation;
			this.m_isRestrictedResourceAccess=isRestrictedResourceAccess;
		}
		Object.defineProperty(Action.prototype, "actionInfo", {
			get: function () {
				return this.m_actionInfo;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Action.prototype, "isWriteOperation", {
			get: function () {
				return this.m_isWriteOperation;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Action.prototype, "isRestrictedResourceAccess", {
			get: function () {
				return this.m_isRestrictedResourceAccess;
			},
			enumerable: true,
			configurable: true
		});
		return Action;
	}());
	OfficeExtension.Action=Action;
})(OfficeExtension || (OfficeExtension={}));
var OfficeExtension;
(function (OfficeExtension) {
	var TraceMarkerActionResultHandler=(function () {
		function TraceMarkerActionResultHandler(callback) {
			this.m_callback=callback;
		}
		TraceMarkerActionResultHandler.prototype._handleResult=function (value) {
			if (this.m_callback) {
				this.m_callback();
			}
		};
		return TraceMarkerActionResultHandler;
	}());
	var ActionFactory=(function () {
		function ActionFactory() {
		}
		ActionFactory.createSetPropertyAction=function (context, parent, propertyName, value) {
			OfficeExtension.Utility.validateObjectPath(parent);
			var actionInfo={
				Id: context._nextId(),
				ActionType: 4,
				Name: propertyName,
				ObjectPathId: parent._objectPath.objectPathInfo.Id,
				ArgumentInfo: {}
			};
			var args=[value];
			var referencedArgumentObjectPaths=OfficeExtension.Utility.setMethodArguments(context, actionInfo.ArgumentInfo, args);
			OfficeExtension.Utility.validateReferencedObjectPaths(referencedArgumentObjectPaths);
			context._pendingRequest.ensureInstantiateObjectPath(parent._objectPath);
			context._pendingRequest.ensureInstantiateObjectPaths(referencedArgumentObjectPaths);
			var ret=new OfficeExtension.Action(actionInfo, true, false);
			context._pendingRequest.addAction(ret);
			context._pendingRequest.addReferencedObjectPath(parent._objectPath);
			context._pendingRequest.addReferencedObjectPaths(referencedArgumentObjectPaths);
			ret.referencedObjectPath=parent._objectPath;
			ret.referencedArgumentObjectPaths=referencedArgumentObjectPaths;
			return ret;
		};
		ActionFactory.createMethodAction=function (context, parent, methodName, operationType, args, isRestrictedResourceAccess) {
			OfficeExtension.Utility.validateObjectPath(parent);
			var actionInfo={
				Id: context._nextId(),
				ActionType: 3,
				Name: methodName,
				ObjectPathId: parent._objectPath.objectPathInfo.Id,
				ArgumentInfo: {}
			};
			var referencedArgumentObjectPaths=OfficeExtension.Utility.setMethodArguments(context, actionInfo.ArgumentInfo, args);
			OfficeExtension.Utility.validateReferencedObjectPaths(referencedArgumentObjectPaths);
			context._pendingRequest.ensureInstantiateObjectPath(parent._objectPath);
			context._pendingRequest.ensureInstantiateObjectPaths(referencedArgumentObjectPaths);
			var isWriteOperation=operationType !=1;
			var ret=new OfficeExtension.Action(actionInfo, isWriteOperation, isRestrictedResourceAccess);
			context._pendingRequest.addAction(ret);
			context._pendingRequest.addReferencedObjectPath(parent._objectPath);
			context._pendingRequest.addReferencedObjectPaths(referencedArgumentObjectPaths);
			ret.referencedObjectPath=parent._objectPath;
			ret.referencedArgumentObjectPaths=referencedArgumentObjectPaths;
			return ret;
		};
		ActionFactory.createQueryAction=function (context, parent, queryOption) {
			OfficeExtension.Utility.validateObjectPath(parent);
			context._pendingRequest.ensureInstantiateObjectPath(parent._objectPath);
			var actionInfo={
				Id: context._nextId(),
				ActionType: 2,
				Name: "",
				ObjectPathId: parent._objectPath.objectPathInfo.Id,
			};
			actionInfo.QueryInfo=queryOption;
			var ret=new OfficeExtension.Action(actionInfo, false, false);
			context._pendingRequest.addAction(ret);
			context._pendingRequest.addReferencedObjectPath(parent._objectPath);
			ret.referencedObjectPath=parent._objectPath;
			return ret;
		};
		ActionFactory.createRecursiveQueryAction=function (context, parent, query) {
			OfficeExtension.Utility.validateObjectPath(parent);
			context._pendingRequest.ensureInstantiateObjectPath(parent._objectPath);
			var actionInfo={
				Id: context._nextId(),
				ActionType: 6,
				Name: "",
				ObjectPathId: parent._objectPath.objectPathInfo.Id,
				RecursiveQueryInfo: query
			};
			var ret=new OfficeExtension.Action(actionInfo, false, false);
			context._pendingRequest.addAction(ret);
			context._pendingRequest.addReferencedObjectPath(parent._objectPath);
			ret.referencedObjectPath=parent._objectPath;
			return ret;
		};
		ActionFactory.createQueryAsJsonAction=function (context, parent, queryOption) {
			OfficeExtension.Utility.validateObjectPath(parent);
			context._pendingRequest.ensureInstantiateObjectPath(parent._objectPath);
			var actionInfo={
				Id: context._nextId(),
				ActionType: 7,
				Name: "",
				ObjectPathId: parent._objectPath.objectPathInfo.Id,
			};
			actionInfo.QueryInfo=queryOption;
			var ret=new OfficeExtension.Action(actionInfo, false, false);
			context._pendingRequest.addAction(ret);
			context._pendingRequest.addReferencedObjectPath(parent._objectPath);
			ret.referencedObjectPath=parent._objectPath;
			return ret;
		};
		ActionFactory.createEnsureUnchangedAction=function (context, parent, objectState) {
			OfficeExtension.Utility.validateObjectPath(parent);
			context._pendingRequest.ensureInstantiateObjectPath(parent._objectPath);
			var actionInfo={
				Id: context._nextId(),
				ActionType: 8,
				Name: "",
				ObjectPathId: parent._objectPath.objectPathInfo.Id,
				ObjectState: objectState
			};
			var ret=new OfficeExtension.Action(actionInfo, false, false);
			context._pendingRequest.addAction(ret);
			context._pendingRequest.addReferencedObjectPath(parent._objectPath);
			ret.referencedObjectPath=parent._objectPath;
			return ret;
		};
		ActionFactory.createUpdateAction=function (context, parent, objectState) {
			OfficeExtension.Utility.validateObjectPath(parent);
			context._pendingRequest.ensureInstantiateObjectPath(parent._objectPath);
			var actionInfo={
				Id: context._nextId(),
				ActionType: 9,
				Name: "",
				ObjectPathId: parent._objectPath.objectPathInfo.Id,
				ObjectState: objectState
			};
			var ret=new OfficeExtension.Action(actionInfo, true, false);
			context._pendingRequest.addAction(ret);
			context._pendingRequest.addReferencedObjectPath(parent._objectPath);
			ret.referencedObjectPath=parent._objectPath;
			return ret;
		};
		ActionFactory.createInstantiateAction=function (context, obj) {
			OfficeExtension.Utility.validateObjectPath(obj);
			context._pendingRequest.ensureInstantiateObjectPath(obj._objectPath.parentObjectPath);
			context._pendingRequest.ensureInstantiateObjectPaths(obj._objectPath.argumentObjectPaths);
			var actionInfo={
				Id: context._nextId(),
				ActionType: 1,
				Name: "",
				ObjectPathId: obj._objectPath.objectPathInfo.Id
			};
			var ret=new OfficeExtension.Action(actionInfo, false, false);
			ret.referencedObjectPath=obj._objectPath;
			context._pendingRequest.addAction(ret);
			context._pendingRequest.addReferencedObjectPath(obj._objectPath);
			context._pendingRequest.addActionResultHandler(ret, new OfficeExtension.InstantiateActionResultHandler(obj));
			return ret;
		};
		ActionFactory.createTraceAction=function (context, message, addTraceMessage) {
			var actionInfo={
				Id: context._nextId(),
				ActionType: 5,
				Name: "Trace",
				ObjectPathId: 0
			};
			var ret=new OfficeExtension.Action(actionInfo, false, false);
			context._pendingRequest.addAction(ret);
			if (addTraceMessage) {
				context._pendingRequest.addTrace(actionInfo.Id, message);
			}
			return ret;
		};
		ActionFactory.createTraceMarkerForCallback=function (context, callback) {
			var action=ActionFactory.createTraceAction(context, null, false);
			context._pendingRequest.addActionResultHandler(action, new TraceMarkerActionResultHandler(callback));
		};
		return ActionFactory;
	}());
	OfficeExtension.ActionFactory=ActionFactory;
})(OfficeExtension || (OfficeExtension={}));
var OfficeExtension;
(function (OfficeExtension) {
	var ClientObject=(function () {
		function ClientObject(context, objectPath) {
			OfficeExtension.Utility.checkArgumentNull(context, "context");
			this.m_context=context;
			this.m_objectPath=objectPath;
			if (this.m_objectPath) {
				if (!context._processingResult) {
					OfficeExtension.ActionFactory.createInstantiateAction(context, this);
					if ((context._autoCleanup) && (this._KeepReference)) {
						context.trackedObjects._autoAdd(this);
					}
				}
			}
		}
		Object.defineProperty(ClientObject.prototype, "context", {
			get: function () {
				return this.m_context;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ClientObject.prototype, "_objectPath", {
			get: function () {
				return this.m_objectPath;
			},
			set: function (value) {
				this.m_objectPath=value;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ClientObject.prototype, "isNull", {
			get: function () {
				OfficeExtension.Utility.throwIfNotLoaded("isNull", this._isNull, null, this._isNull);
				return this._isNull;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ClientObject.prototype, "isNullObject", {
			get: function () {
				OfficeExtension.Utility.throwIfNotLoaded("isNullObject", this._isNull, null, this._isNull);
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
				this.m_isNull=value;
				if (value && this.m_objectPath) {
					this.m_objectPath._updateAsNullObject();
				}
			},
			enumerable: true,
			configurable: true
		});
		ClientObject.prototype._handleResult=function (value) {
			this._isNull=OfficeExtension.Utility.isNullOrUndefined(value);
			this.context.trackedObjects._autoTrackIfNecessaryWhenHandleObjectResultValue(this, value);
		};
		ClientObject.prototype._handleIdResult=function (value) {
			this._isNull=OfficeExtension.Utility.isNullOrUndefined(value);
			OfficeExtension.Utility.fixObjectPathIfNecessary(this, value);
			this.context.trackedObjects._autoTrackIfNecessaryWhenHandleObjectResultValue(this, value);
		};
		ClientObject.prototype._handleRetrieveResult=function (value, result) {
			this._handleIdResult(value);
		};
		ClientObject.prototype._recursivelySet=function (input, options, scalarWriteablePropertyNames, objectPropertyNames, notAllowedToBeSetPropertyNames) {
			var isClientObject=(input instanceof ClientObject);
			var originalInput=input;
			if (isClientObject) {
				if (Object.getPrototypeOf(this)===Object.getPrototypeOf(input)) {
					input=JSON.parse(JSON.stringify(input));
				}
				else {
					throw OfficeExtension._Internal.RuntimeError._createInvalidArgError({
						argumentName: 'properties',
						errorLocation: this._className+".set"
					});
				}
			}
			try {
				var prop;
				for (var i=0; i < scalarWriteablePropertyNames.length; i++) {
					prop=scalarWriteablePropertyNames[i];
					if (input.hasOwnProperty(prop)) {
						if (typeof input[prop] !=="undefined") {
							this[prop]=input[prop];
						}
					}
				}
				for (var i=0; i < objectPropertyNames.length; i++) {
					prop=objectPropertyNames[i];
					if (input.hasOwnProperty(prop)) {
						if (typeof input[prop] !=="undefined") {
							var dataToPassToSet=isClientObject ? originalInput[prop] : input[prop];
							this[prop].set(dataToPassToSet, options);
						}
					}
				}
				var throwOnReadOnly=!isClientObject;
				if (options && !OfficeExtension.Utility.isNullOrUndefined(throwOnReadOnly)) {
					throwOnReadOnly=options.throwOnReadOnly;
				}
				for (var i=0; i < notAllowedToBeSetPropertyNames.length; i++) {
					prop=notAllowedToBeSetPropertyNames[i];
					if (input.hasOwnProperty(prop)) {
						if (typeof input[prop] !=="undefined" && throwOnReadOnly) {
							throw new OfficeExtension._Internal.RuntimeError({
								code: OfficeExtension.ErrorCodes.invalidArgument,
								message: OfficeExtension.Utility._getResourceString(OfficeExtension.ResourceStrings.cannotApplyPropertyThroughSetMethod, prop),
								debugInfo: {
									errorLocation: prop
								}
							});
						}
					}
				}
				for (prop in input) {
					if (scalarWriteablePropertyNames.indexOf(prop) < 0 && objectPropertyNames.indexOf(prop) < 0) {
						var propertyDescriptor=Object.getOwnPropertyDescriptor(Object.getPrototypeOf(this), prop);
						if (!propertyDescriptor) {
							throw new OfficeExtension._Internal.RuntimeError({
								code: OfficeExtension.ErrorCodes.invalidArgument,
								message: OfficeExtension.Utility._getResourceString(OfficeExtension.ResourceStrings.propertyDoesNotExist, prop),
								debugInfo: {
									errorLocation: prop
								}
							});
						}
						if (throwOnReadOnly && !propertyDescriptor.set) {
							throw new OfficeExtension._Internal.RuntimeError({
								code: OfficeExtension.ErrorCodes.invalidArgument,
								message: OfficeExtension.Utility._getResourceString(OfficeExtension.ResourceStrings.attemptingToSetReadOnlyProperty, prop),
								debugInfo: {
									errorLocation: prop
								}
							});
						}
					}
				}
			}
			catch (innerError) {
				throw new OfficeExtension._Internal.RuntimeError({
					code: OfficeExtension.ErrorCodes.invalidArgument,
					message: OfficeExtension.Utility._getResourceString(OfficeExtension.ResourceStrings.invalidArgument, 'properties'),
					debugInfo: {
						errorLocation: this._className+".set"
					},
					innerError: innerError
				});
			}
		};
		ClientObject.prototype._recursivelyUpdate=function (properties) {
			var shouldPolyfill=OfficeExtension._internalConfig.alwaysPolyfillClientObjectUpdateMethod;
			if (!shouldPolyfill) {
				shouldPolyfill=!OfficeExtension.Utility.isSetSupported("RichApiRuntime", "1.2");
			}
			try {
				var scalarPropNames=this[OfficeExtension.Constants.scalarPropertyNames];
				if (!scalarPropNames) {
					scalarPropNames=[];
				}
				var scalarPropUpdatable=this[OfficeExtension.Constants.scalarPropertyUpdateable];
				if (!scalarPropUpdatable) {
					scalarPropUpdatable=[];
					for (var i=0; i < scalarPropNames.length; i++) {
						scalarPropUpdatable.push(false);
					}
				}
				var navigationPropNames=this[OfficeExtension.Constants.navigationPropertyNames];
				if (!navigationPropNames) {
					navigationPropNames=[];
				}
				var scalarProps={};
				var navigationProps={};
				var scalarPropCount=0;
				for (var propName in properties) {
					var index=scalarPropNames.indexOf(propName);
					if (index >=0) {
						if (!scalarPropUpdatable[index]) {
							throw new OfficeExtension._Internal.RuntimeError({
								code: OfficeExtension.ErrorCodes.invalidArgument,
								message: OfficeExtension.Utility._getResourceString(OfficeExtension.ResourceStrings.attemptingToSetReadOnlyProperty, propName),
								debugInfo: {
									errorLocation: propName
								}
							});
						}
						scalarProps[propName]=properties[propName];
++scalarPropCount;
					}
					else if (navigationPropNames.indexOf(propName) >=0) {
						navigationProps[propName]=properties[propName];
					}
					else {
						throw new OfficeExtension._Internal.RuntimeError({
							code: OfficeExtension.ErrorCodes.invalidArgument,
							message: OfficeExtension.Utility._getResourceString(OfficeExtension.ResourceStrings.propertyDoesNotExist, propName),
							debugInfo: {
								errorLocation: propName
							}
						});
					}
				}
				if (scalarPropCount > 0) {
					if (shouldPolyfill) {
						for (var i=0; i < scalarPropNames.length; i++) {
							var propName=scalarPropNames[i];
							var propValue=scalarProps[propName];
							if (!OfficeExtension.Utility.isUndefined(propValue)) {
								OfficeExtension.ActionFactory.createSetPropertyAction(this.context, this, propName, propValue);
							}
						}
					}
					else {
						OfficeExtension.ActionFactory.createUpdateAction(this.context, this, scalarProps);
					}
				}
				for (var propName in navigationProps) {
					var navigationPropProxy=this[propName];
					var navigationPropValue=navigationProps[propName];
					navigationPropProxy._recursivelyUpdate(navigationPropValue);
				}
			}
			catch (innerError) {
				throw new OfficeExtension._Internal.RuntimeError({
					code: OfficeExtension.ErrorCodes.invalidArgument,
					message: OfficeExtension.Utility._getResourceString(OfficeExtension.ResourceStrings.invalidArgument, 'properties'),
					debugInfo: {
						errorLocation: this._className+".update"
					},
					innerError: innerError
				});
			}
		};
		return ClientObject;
	}());
	OfficeExtension.ClientObject=ClientObject;
})(OfficeExtension || (OfficeExtension={}));
var OfficeExtension;
(function (OfficeExtension) {
	var ClientRequest=(function () {
		function ClientRequest(context) {
			this.m_context=context;
			this.m_actions=[];
			this.m_actionResultHandler={};
			this.m_referencedObjectPaths={};
			this.m_instantiatedObjectPaths={};
			this.m_flags=0;
			this.m_traceInfos={};
			this.m_pendingProcessEventHandlers=[];
			this.m_pendingEventHandlerActions={};
			this.m_responseTraceIds={};
			this.m_responseTraceMessages=[];
			this.m_preSyncPromises=[];
		}
		Object.defineProperty(ClientRequest.prototype, "flags", {
			get: function () {
				return this.m_flags;
			},
			enumerable: true,
			configurable: true
		});
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
		ClientRequest.prototype._setResponseTraceIds=function (value) {
			if (value) {
				for (var i=0; i < value.length; i++) {
					var traceId=value[i];
					this.m_responseTraceIds[traceId]=traceId;
					var message=this.m_traceInfos[traceId];
					if (!OfficeExtension.Utility.isNullOrUndefined(message)) {
						this.m_responseTraceMessages.push(message);
					}
				}
			}
		};
		ClientRequest.prototype.addAction=function (action) {
			if (this.m_context.batchMode===1) {
				var isSafeAction=false;
				if (action.actionInfo.ActionType===1 &&
					action.referencedObjectPath.objectPathInfo.ObjectPathType===4) {
					isSafeAction=true;
				}
				if (!isSafeAction) {
					this.m_context.ensureInProgressBatchIfBatchMode();
				}
			}
			if (action.isWriteOperation) {
				this.m_flags=this.m_flags | 1;
			}
			if (action.isRestrictedResourceAccess) {
				this.m_flags=this.m_flags | 2;
			}
			this.m_actions.push(action);
			if (action.actionInfo.ActionType==1) {
				this.m_instantiatedObjectPaths[action.actionInfo.ObjectPathId]=action;
			}
		};
		Object.defineProperty(ClientRequest.prototype, "hasActions", {
			get: function () {
				return this.m_actions.length > 0;
			},
			enumerable: true,
			configurable: true
		});
		ClientRequest.prototype._getLastAction=function () {
			return this.m_actions[this.m_actions.length - 1];
		};
		ClientRequest.prototype.addTrace=function (actionId, message) {
			this.m_traceInfos[actionId]=message;
		};
		ClientRequest.prototype.ensureInstantiateObjectPath=function (objectPath) {
			if (objectPath) {
				if (this.m_instantiatedObjectPaths[objectPath.objectPathInfo.Id]) {
					return;
				}
				this.ensureInstantiateObjectPath(objectPath.parentObjectPath);
				this.ensureInstantiateObjectPaths(objectPath.argumentObjectPaths);
				if (!this.m_instantiatedObjectPaths[objectPath.objectPathInfo.Id]) {
					var actionInfo={
						Id: this.m_context._nextId(),
						ActionType: 1,
						Name: "",
						ObjectPathId: objectPath.objectPathInfo.Id
					};
					var instantiateAction=new OfficeExtension.Action(actionInfo, false, false);
					instantiateAction.referencedObjectPath=objectPath;
					this.addReferencedObjectPath(objectPath);
					this.addAction(instantiateAction);
				}
			}
		};
		ClientRequest.prototype.ensureInstantiateObjectPaths=function (objectPaths) {
			if (objectPaths) {
				for (var i=0; i < objectPaths.length; i++) {
					this.ensureInstantiateObjectPath(objectPaths[i]);
				}
			}
		};
		ClientRequest.prototype.addReferencedObjectPath=function (objectPath) {
			if (this.m_referencedObjectPaths[objectPath.objectPathInfo.Id]) {
				return;
			}
			if (!objectPath.isValid) {
				throw new OfficeExtension._Internal.RuntimeError({
					code: OfficeExtension.ErrorCodes.invalidObjectPath,
					message: OfficeExtension.Utility._getResourceString(OfficeExtension.ResourceStrings.invalidObjectPath, OfficeExtension.Utility.getObjectPathExpression(objectPath)),
					debugInfo: {
						errorLocation: OfficeExtension.Utility.getObjectPathExpression(objectPath)
					}
				});
			}
			while (objectPath) {
				if (objectPath.isWriteOperation) {
					this.m_flags=this.m_flags | 1;
				}
				if (objectPath.isRestrictedResourceAccess) {
					this.m_flags=this.m_flags | 2;
				}
				this.m_referencedObjectPaths[objectPath.objectPathInfo.Id]=objectPath;
				if (objectPath.objectPathInfo.ObjectPathType==3) {
					this.addReferencedObjectPaths(objectPath.argumentObjectPaths);
				}
				objectPath=objectPath.parentObjectPath;
			}
		};
		ClientRequest.prototype.addReferencedObjectPaths=function (objectPaths) {
			if (objectPaths) {
				for (var i=0; i < objectPaths.length; i++) {
					this.addReferencedObjectPath(objectPaths[i]);
				}
			}
		};
		ClientRequest.prototype.addActionResultHandler=function (action, resultHandler) {
			this.m_actionResultHandler[action.actionInfo.Id]=resultHandler;
		};
		ClientRequest.prototype.buildRequestMessageBody=function () {
			if (OfficeExtension._internalConfig.enableEarlyDispose) {
				ClientRequest._calculateLastUsedObjectPathIds(this.m_actions);
			}
			var objectPaths={};
			for (var i in this.m_referencedObjectPaths) {
				objectPaths[i]=this.m_referencedObjectPaths[i].objectPathInfo;
			}
			var actions=[];
			for (var index=0; index < this.m_actions.length; index++) {
				actions.push(this.m_actions[index].actionInfo);
			}
			var ret={
				AutoKeepReference: this.m_context._autoCleanup,
				Actions: actions,
				ObjectPaths: objectPaths
			};
			return ret;
		};
		ClientRequest.prototype.processResponse=function (actionResults) {
			if (actionResults) {
				for (var i=0; i < actionResults.length; i++) {
					var actionResult=actionResults[i];
					var handler=this.m_actionResultHandler[actionResult.ActionId];
					if (handler) {
						handler._handleResult(actionResult.Value);
					}
				}
			}
		};
		ClientRequest.prototype.invalidatePendingInvalidObjectPaths=function () {
			for (var i in this.m_referencedObjectPaths) {
				if (this.m_referencedObjectPaths[i].isInvalidAfterRequest) {
					this.m_referencedObjectPaths[i].isValid=false;
				}
			}
		};
		ClientRequest.prototype._addPendingEventHandlerAction=function (eventHandlers, action) {
			if (!this.m_pendingEventHandlerActions[eventHandlers._id]) {
				this.m_pendingEventHandlerActions[eventHandlers._id]=[];
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
		ClientRequest.prototype._getPendingEventHandlerActions=function (eventHandlers) {
			return this.m_pendingEventHandlerActions[eventHandlers._id];
		};
		ClientRequest.prototype._addPreSyncPromise=function (value) {
			this.m_preSyncPromises.push(value);
		};
		Object.defineProperty(ClientRequest.prototype, "_preSyncPromises", {
			get: function () {
				return this.m_preSyncPromises;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ClientRequest.prototype, "_actions", {
			get: function () {
				return this.m_actions;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ClientRequest.prototype, "_objectPaths", {
			get: function () {
				return this.m_referencedObjectPaths;
			},
			enumerable: true,
			configurable: true
		});
		ClientRequest._updateLastUsedActionIdOfObjectPathId=function (lastUsedActionIdOfObjectPathId, objectPath, actionId) {
			while (objectPath) {
				if (lastUsedActionIdOfObjectPathId[objectPath.objectPathInfo.Id]) {
					return;
				}
				lastUsedActionIdOfObjectPathId[objectPath.objectPathInfo.Id]=actionId;
				var argumentObjectPaths=objectPath.argumentObjectPaths;
				if (argumentObjectPaths) {
					var argumentObjectPathsLength=argumentObjectPaths.length;
					for (var i=0; i < argumentObjectPathsLength; i++) {
						ClientRequest._updateLastUsedActionIdOfObjectPathId(lastUsedActionIdOfObjectPathId, argumentObjectPaths[i], actionId);
					}
				}
				objectPath=objectPath.parentObjectPath;
			}
		};
		ClientRequest._calculateLastUsedObjectPathIds=function (actions) {
			var lastUsedActionIdOfObjectPathId={};
			var actionsLength=actions.length;
			for (var index=actionsLength - 1; index >=0; --index) {
				var action=actions[index];
				var actionId=action.actionInfo.Id;
				if (action.referencedObjectPath) {
					ClientRequest._updateLastUsedActionIdOfObjectPathId(lastUsedActionIdOfObjectPathId, action.referencedObjectPath, actionId);
				}
				var referencedObjectPaths=action.referencedArgumentObjectPaths;
				if (referencedObjectPaths) {
					var referencedObjectPathsLength=referencedObjectPaths.length;
					for (var refIndex=0; refIndex < referencedObjectPathsLength; refIndex++) {
						ClientRequest._updateLastUsedActionIdOfObjectPathId(lastUsedActionIdOfObjectPathId, referencedObjectPaths[refIndex], actionId);
					}
				}
			}
			var lastUsedObjectPathIdsOfAction={};
			for (var key in lastUsedActionIdOfObjectPathId) {
				var actionId=lastUsedActionIdOfObjectPathId[key];
				var objectPathIds=lastUsedObjectPathIdsOfAction[actionId];
				if (!objectPathIds) {
					objectPathIds=[];
					lastUsedObjectPathIdsOfAction[actionId]=objectPathIds;
				}
				objectPathIds.push(parseInt(key));
			}
			for (var index=0; index < actionsLength; index++) {
				var action=actions[index];
				var lastUsedObjectPathIds=lastUsedObjectPathIdsOfAction[action.actionInfo.Id];
				if (lastUsedObjectPathIds && lastUsedObjectPathIds.length > 0) {
					action.actionInfo.L=lastUsedObjectPathIds;
				}
				else if (action.actionInfo.L) {
					delete action.actionInfo.L;
				}
			}
		};
		return ClientRequest;
	}());
	OfficeExtension.ClientRequest=ClientRequest;
})(OfficeExtension || (OfficeExtension={}));
var OfficeExtension;
(function (OfficeExtension) {
	OfficeExtension._internalConfig={
		showDisposeInfoInDebugInfo: false,
		enableEarlyDispose: true,
		alwaysPolyfillClientObjectUpdateMethod: false,
		alwaysPolyfillClientObjectRetrieveMethod: false
	};
	var SessionBase=(function () {
		function SessionBase() {
		}
		SessionBase.prototype._resolveRequestUrlAndHeaderInfo=function () {
			return OfficeExtension.Utility._createPromiseFromResult(null);
		};
		SessionBase.prototype._createRequestExecutorOrNull=function () {
			return null;
		};
		Object.defineProperty(SessionBase.prototype, "eventRegistration", {
			get: function () {
				return OfficeExtension._Internal.officeJsEventRegistration;
			},
			enumerable: true,
			configurable: true
		});
		return SessionBase;
	}());
	OfficeExtension.SessionBase=SessionBase;
	var ClientRequestContext=(function () {
		function ClientRequestContext(url) {
			this.m_customRequestHeaders={};
			this.m_batchMode=0;
			this._onRunFinishedNotifiers=[];
			this.m_nextId=0;
			if (ClientRequestContext._overrideSession) {
				this.m_requestUrlAndHeaderInfoResolver=ClientRequestContext._overrideSession;
			}
			else {
				if (OfficeExtension.Utility.isNullOrUndefined(url) || typeof (url)==="string" && url.length===0) {
					url=ClientRequestContext.defaultRequestUrlAndHeaders;
					if (!url) {
						url={ url: OfficeExtension.Constants.localDocument, headers: {} };
					}
				}
				if (typeof (url)==="string") {
					this.m_requestUrlAndHeaderInfo={ url: url, headers: {} };
				}
				else if (ClientRequestContext.isRequestUrlAndHeaderInfoResolver(url)) {
					this.m_requestUrlAndHeaderInfoResolver=url;
				}
				else if (ClientRequestContext.isRequestUrlAndHeaderInfo(url)) {
					var requestInfo=url;
					this.m_requestUrlAndHeaderInfo={ url: requestInfo.url, headers: {} };
					OfficeExtension.Utility._copyHeaders(requestInfo.headers, this.m_requestUrlAndHeaderInfo.headers);
				}
				else {
					throw OfficeExtension._Internal.RuntimeError._createInvalidArgError("url");
				}
			}
			if (this.m_requestUrlAndHeaderInfoResolver instanceof SessionBase) {
				this.m_session=this.m_requestUrlAndHeaderInfoResolver;
			}
			this._processingResult=false;
			this._customData=OfficeExtension.Constants.iterativeExecutor;
			this.sync=this.sync.bind(this);
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
				return OfficeExtension._Internal.officeJsEventRegistration;
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
				if (this.m_pendingRequest==null) {
					this.m_pendingRequest=new OfficeExtension.ClientRequest(this);
				}
				return this.m_pendingRequest;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ClientRequestContext.prototype, "debugInfo", {
			get: function () {
				var prettyPrinter=new OfficeExtension.RequestPrettyPrinter(this._rootObjectPropertyName, this._pendingRequest._objectPaths, this._pendingRequest._actions, OfficeExtension._internalConfig.showDisposeInfoInDebugInfo);
				var statements=prettyPrinter.process();
				return { pendingStatements: statements };
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ClientRequestContext.prototype, "trackedObjects", {
			get: function () {
				if (!this.m_trackedObjects) {
					this.m_trackedObjects=new OfficeExtension.TrackedObjects(this);
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
		ClientRequestContext.prototype.ensureInProgressBatchIfBatchMode=function () {
			if (this.m_batchMode===1 && !this.m_explicitBatchInProgress) {
				throw OfficeExtension.Utility.createRuntimeError(OfficeExtension.ErrorCodes.generalException, OfficeExtension.Utility._getResourceString(OfficeExtension.ResourceStrings.notInsideBatch), null);
			}
		};
		ClientRequestContext.prototype.load=function (clientObj, option) {
			OfficeExtension.Utility.validateContext(this, clientObj);
			var queryOption=ClientRequestContext._parseQueryOption(option);
			var action=OfficeExtension.ActionFactory.createQueryAction(this, clientObj, queryOption);
			this._pendingRequest.addActionResultHandler(action, clientObj);
		};
		ClientRequestContext.isLoadOption=function (loadOption) {
			if (!OfficeExtension.Utility.isUndefined(loadOption.select) && (typeof (loadOption.select)==="string" || Array.isArray(loadOption.select)))
				return true;
			if (!OfficeExtension.Utility.isUndefined(loadOption.expand) && (typeof (loadOption.expand)==="string" || Array.isArray(loadOption.expand)))
				return true;
			if (!OfficeExtension.Utility.isUndefined(loadOption.top) && typeof (loadOption.top)==="number")
				return true;
			if (!OfficeExtension.Utility.isUndefined(loadOption.skip) && typeof (loadOption.skip)==="number")
				return true;
			for (var i in loadOption) {
				return false;
			}
			return true;
		};
		ClientRequestContext.parseStrictLoadOption=function (option) {
			var ret={ Select: [] };
			ClientRequestContext.parseStrictLoadOptionHelper(ret, "", "option", option);
			return ret;
		};
		ClientRequestContext.combineQueryPath=function (pathPrefix, key, separator) {
			if (pathPrefix.length===0) {
				return key;
			}
			else {
				return pathPrefix+separator+key;
			}
		};
		ClientRequestContext.parseStrictLoadOptionHelper=function (queryInfo, pathPrefix, argPrefix, option) {
			for (var key in option) {
				var value=option[key];
				if (key==="$all") {
					if (typeof (value) !=="boolean") {
						throw OfficeExtension._Internal.RuntimeError._createInvalidArgError(ClientRequestContext.combineQueryPath(argPrefix, key, "."));
					}
					if (value) {
						queryInfo.Select.push(ClientRequestContext.combineQueryPath(pathPrefix, "*", "/"));
					}
				}
				else if (key==="$top") {
					if (typeof (value) !=="number" || pathPrefix.length > 0) {
						throw OfficeExtension._Internal.RuntimeError._createInvalidArgError(ClientRequestContext.combineQueryPath(argPrefix, key, "."));
					}
					queryInfo.Top=value;
				}
				else if (key==="$skip") {
					if (typeof (value) !=="number" || pathPrefix.length > 0) {
						throw OfficeExtension._Internal.RuntimeError._createInvalidArgError(ClientRequestContext.combineQueryPath(argPrefix, key, "."));
					}
					queryInfo.Skip=value;
				}
				else {
					if (typeof (value)==="boolean") {
						if (value) {
							queryInfo.Select.push(ClientRequestContext.combineQueryPath(pathPrefix, key, "/"));
						}
					}
					else if (typeof (value)==="object") {
						ClientRequestContext.parseStrictLoadOptionHelper(queryInfo, ClientRequestContext.combineQueryPath(pathPrefix, key, "/"), ClientRequestContext.combineQueryPath(argPrefix, key, "."), value);
					}
					else {
						throw OfficeExtension._Internal.RuntimeError._createInvalidArgError(ClientRequestContext.combineQueryPath(argPrefix, key, "."));
					}
				}
			}
		};
		ClientRequestContext._parseQueryOption=function (option) {
			var queryOption={};
			if (typeof (option)=="string") {
				var select=option;
				queryOption.Select=OfficeExtension.Utility._parseSelectExpand(select);
			}
			else if (Array.isArray(option)) {
				queryOption.Select=option;
			}
			else if (typeof (option)==="object") {
				var loadOption=option;
				if (ClientRequestContext.isLoadOption(loadOption)) {
					if (typeof (loadOption.select)=="string") {
						queryOption.Select=OfficeExtension.Utility._parseSelectExpand(loadOption.select);
					}
					else if (Array.isArray(loadOption.select)) {
						queryOption.Select=loadOption.select;
					}
					else if (!OfficeExtension.Utility.isNullOrUndefined(loadOption.select)) {
						throw OfficeExtension._Internal.RuntimeError._createInvalidArgError("option.select");
					}
					if (typeof (loadOption.expand)=="string") {
						queryOption.Expand=OfficeExtension.Utility._parseSelectExpand(loadOption.expand);
					}
					else if (Array.isArray(loadOption.expand)) {
						queryOption.Expand=loadOption.expand;
					}
					else if (!OfficeExtension.Utility.isNullOrUndefined(loadOption.expand)) {
						throw OfficeExtension._Internal.RuntimeError._createInvalidArgError("option.expand");
					}
					if (typeof (loadOption.top)==="number") {
						queryOption.Top=loadOption.top;
					}
					else if (!OfficeExtension.Utility.isNullOrUndefined(loadOption.top)) {
						throw OfficeExtension._Internal.RuntimeError._createInvalidArgError("option.top");
					}
					if (typeof (loadOption.skip)==="number") {
						queryOption.Skip=loadOption.skip;
					}
					else if (!OfficeExtension.Utility.isNullOrUndefined(loadOption.skip)) {
						throw OfficeExtension._Internal.RuntimeError._createInvalidArgError("option.skip");
					}
				}
				else {
					queryOption=ClientRequestContext.parseStrictLoadOption(option);
				}
			}
			else if (!OfficeExtension.Utility.isNullOrUndefined(option)) {
				throw OfficeExtension._Internal.RuntimeError._createInvalidArgError("option");
			}
			return queryOption;
		};
		ClientRequestContext.prototype.loadRecursive=function (clientObj, options, maxDepth) {
			if (!OfficeExtension.Utility.isPlainJsonObject(options)) {
				throw OfficeExtension._Internal.RuntimeError._createInvalidArgError("options");
			}
			var quries={};
			for (var key in options) {
				quries[key]=ClientRequestContext._parseQueryOption(options[key]);
			}
			var action=OfficeExtension.ActionFactory.createRecursiveQueryAction(this, clientObj, { Queries: quries, MaxDepth: maxDepth });
			this._pendingRequest.addActionResultHandler(action, clientObj);
		};
		ClientRequestContext.prototype.trace=function (message) {
			OfficeExtension.ActionFactory.createTraceAction(this, message, true);
		};
		ClientRequestContext.prototype._processOfficeJsErrorResponse=function (officeJsErrorCode, response) {
		};
		ClientRequestContext.prototype.ensureRequestUrlAndHeaderInfo=function () {
			var _this=this;
			return OfficeExtension.Utility._createPromiseFromResult(null)
				.then(function () {
				if (!_this.m_requestUrlAndHeaderInfo) {
					return _this.m_requestUrlAndHeaderInfoResolver._resolveRequestUrlAndHeaderInfo()
						.then(function (value) {
						_this.m_requestUrlAndHeaderInfo=value;
						if (!_this.m_requestUrlAndHeaderInfo) {
							_this.m_requestUrlAndHeaderInfo={ url: OfficeExtension.Constants.localDocument, headers: {} };
						}
						if (OfficeExtension.Utility.isNullOrEmptyString(_this.m_requestUrlAndHeaderInfo.url)) {
							_this.m_requestUrlAndHeaderInfo.url=OfficeExtension.Constants.localDocument;
						}
						if (!_this.m_requestUrlAndHeaderInfo.headers) {
							_this.m_requestUrlAndHeaderInfo.headers={};
						}
						if (typeof (_this.m_requestUrlAndHeaderInfoResolver._createRequestExecutorOrNull)==="function") {
							var executor=_this.m_requestUrlAndHeaderInfoResolver._createRequestExecutorOrNull();
							if (executor) {
								_this._requestExecutor=executor;
							}
						}
					});
				}
			});
		};
		ClientRequestContext.prototype.syncPrivateMain=function () {
			var _this=this;
			return this.ensureRequestUrlAndHeaderInfo()
				.then(function () {
				var req=_this._pendingRequest;
				_this.m_pendingRequest=null;
				return _this.processPreSyncPromises(req)
					.then(function () { return _this.syncPrivate(req); });
			});
		};
		ClientRequestContext.prototype.syncPrivate=function (req) {
			var _this=this;
			if (!req.hasActions) {
				return this.processPendingEventHandlers(req);
			}
			var msgBody=req.buildRequestMessageBody();
			var requestFlags=req.flags;
			if (!this._requestExecutor) {
				if (OfficeExtension.Utility._isLocalDocumentUrl(this.m_requestUrlAndHeaderInfo.url)) {
					this._requestExecutor=new OfficeExtension.OfficeJsRequestExecutor(this);
				}
				else {
					this._requestExecutor=new OfficeExtension.HttpRequestExecutor();
				}
			}
			var requestExecutor=this._requestExecutor;
			var headers={};
			OfficeExtension.Utility._copyHeaders(this.m_requestUrlAndHeaderInfo.headers, headers);
			OfficeExtension.Utility._copyHeaders(this.m_customRequestHeaders, headers);
			var requestExecutorRequestMessage={
				Url: this.m_requestUrlAndHeaderInfo.url,
				Headers: headers,
				Body: msgBody
			};
			req.invalidatePendingInvalidObjectPaths();
			var errorFromResponse=null;
			var errorFromProcessEventHandlers=null;
			this._lastSyncStart=performance.now();
			return requestExecutor.executeAsync(this._customData, requestFlags, requestExecutorRequestMessage)
				.then(function (response) {
				_this._lastSyncEnd=performance.now();
				errorFromResponse=_this.processRequestExecutorResponseMessage(req, response);
				return _this.processPendingEventHandlers(req)
					.catch(function (ex) {
					OfficeExtension.Utility.log("Error in processPendingEventHandlers");
					OfficeExtension.Utility.log(JSON.stringify(ex));
					errorFromProcessEventHandlers=ex;
				});
			})
				.then(function () {
				if (errorFromResponse) {
					OfficeExtension.Utility.log("Throw error from response: "+JSON.stringify(errorFromResponse));
					throw errorFromResponse;
				}
				if (errorFromProcessEventHandlers) {
					OfficeExtension.Utility.log("Throw error from ProcessEventHandler: "+JSON.stringify(errorFromProcessEventHandlers));
					var transformedError=null;
					if (errorFromProcessEventHandlers instanceof OfficeExtension._Internal.RuntimeError) {
						transformedError=errorFromProcessEventHandlers;
						transformedError.traceMessages=req._responseTraceMessages;
					}
					else {
						var message=null;
						if (typeof (errorFromProcessEventHandlers)==="string") {
							message=errorFromProcessEventHandlers;
						}
						else {
							message=errorFromProcessEventHandlers.message;
						}
						if (OfficeExtension.Utility.isNullOrEmptyString(message)) {
							message=OfficeExtension.Utility._getResourceString(OfficeExtension.ResourceStrings.cannotRegisterEvent);
						}
						transformedError=new OfficeExtension._Internal.RuntimeError({
							code: OfficeExtension.ErrorCodes.cannotRegisterEvent,
							message: message,
							traceMessages: req._responseTraceMessages
						});
					}
					throw transformedError;
				}
			});
		};
		ClientRequestContext.prototype.processRequestExecutorResponseMessage=function (req, response) {
			if (response.Body && response.Body.TraceIds) {
				req._setResponseTraceIds(response.Body.TraceIds);
			}
			var traceMessages=req._responseTraceMessages;
			var errorStatementInfo=null;
			if (response.Body) {
				if (response.Body.Error &&
					response.Body.Error.ActionIndex >=0) {
					var prettyPrinter=new OfficeExtension.RequestPrettyPrinter(this._rootObjectPropertyName, req._objectPaths, req._actions);
					errorStatementInfo=prettyPrinter.processForDebugStatementInfo(response.Body.Error.ActionIndex);
				}
				var actionResults=null;
				if (response.Body.Results) {
					actionResults=response.Body.Results;
				}
				else if (response.Body.ProcessedResults && response.Body.ProcessedResults.Results) {
					actionResults=response.Body.ProcessedResults.Results;
				}
				if (actionResults) {
					this._processingResult=true;
					try {
						req.processResponse(actionResults);
					}
					finally {
						this._processingResult=false;
					}
				}
			}
			if (!OfficeExtension.Utility.isNullOrEmptyString(response.ErrorCode)) {
				return new OfficeExtension._Internal.RuntimeError({
					code: response.ErrorCode,
					message: response.ErrorMessage,
					traceMessages: traceMessages
				});
			}
			else if (response.Body && response.Body.Error) {
				var debugInfo={
					errorLocation: response.Body.Error.Location
				};
				if (errorStatementInfo) {
					debugInfo.statement=errorStatementInfo.statement;
					debugInfo.surroundingStatements=errorStatementInfo.surroundingStatements;
				}
				return new OfficeExtension._Internal.RuntimeError({
					code: response.Body.Error.Code,
					message: response.Body.Error.Message,
					traceMessages: traceMessages,
					debugInfo: debugInfo
				});
			}
			return null;
		};
		ClientRequestContext.prototype.processPendingEventHandlers=function (req) {
			var ret=OfficeExtension.Utility._createPromiseFromResult(null);
			for (var i=0; i < req._pendingProcessEventHandlers.length; i++) {
				var eventHandlers=req._pendingProcessEventHandlers[i];
				ret=ret.then(this.createProcessOneEventHandlersFunc(eventHandlers, req));
			}
			return ret;
		};
		ClientRequestContext.prototype.createProcessOneEventHandlersFunc=function (eventHandlers, req) {
			return function () { return eventHandlers._processRegistration(req); };
		};
		ClientRequestContext.prototype.processPreSyncPromises=function (req) {
			var ret=OfficeExtension.Utility._createPromiseFromResult(null);
			for (var i=0; i < req._preSyncPromises.length; i++) {
				var p=req._preSyncPromises[i];
				ret=ret.then(this.createProcessOneProSyncFunc(p));
			}
			return ret;
		};
		ClientRequestContext.prototype.createProcessOneProSyncFunc=function (p) {
			return function () { return p; };
		};
		ClientRequestContext.prototype.sync=function (passThroughValue) {
			return this.syncPrivateMain().then(function () { return passThroughValue; });
		};
		ClientRequestContext.prototype.batch=function (batchBody) {
			var _this=this;
			if (this.m_batchMode !==1) {
				return OfficeExtension._Internal.OfficePromise.reject(OfficeExtension.Utility.createRuntimeError(OfficeExtension.ErrorCodes.generalException, null, null));
			}
			if (this.m_explicitBatchInProgress) {
				return OfficeExtension._Internal.OfficePromise.reject(OfficeExtension.Utility.createRuntimeError(OfficeExtension.ErrorCodes.generalException, OfficeExtension.Utility._getResourceString(OfficeExtension.ResourceStrings.pendingBatchInProgress), null));
			}
			if (OfficeExtension.Utility.isNullOrUndefined(batchBody)) {
				return OfficeExtension.Utility._createPromiseFromResult(null);
			}
			this.m_explicitBatchInProgress=true;
			var previousRequest=this.m_pendingRequest;
			this.m_pendingRequest=new OfficeExtension.ClientRequest(this);
			var batchBodyResult;
			try {
				batchBodyResult=batchBody(this._rootObject, this);
			}
			catch (ex) {
				this.m_explicitBatchInProgress=false;
				this.m_pendingRequest=previousRequest;
				return OfficeExtension._Internal.OfficePromise.reject(ex);
			}
			var request;
			var batchBodyResultPromise;
			if (typeof (batchBodyResult)==="object" &&
				batchBodyResult &&
				typeof (batchBodyResult.then)==="function") {
				batchBodyResultPromise=OfficeExtension.Utility._createPromiseFromResult(null)
					.then(function () {
					return batchBodyResult;
				})
					.then(function (result) {
					_this.m_explicitBatchInProgress=false;
					request=_this.m_pendingRequest;
					_this.m_pendingRequest=previousRequest;
					return result;
				})
					.catch(function (ex) {
					_this.m_explicitBatchInProgress=false;
					request=_this.m_pendingRequest;
					_this.m_pendingRequest=previousRequest;
					return OfficeExtension._Internal.OfficePromise.reject(ex);
				});
			}
			else {
				this.m_explicitBatchInProgress=false;
				request=this.m_pendingRequest;
				this.m_pendingRequest=previousRequest;
				batchBodyResultPromise=OfficeExtension.Utility._createPromiseFromResult(batchBodyResult);
			}
			return batchBodyResultPromise
				.then(function (result) {
				return _this.ensureRequestUrlAndHeaderInfo()
					.then(function () {
					return _this.syncPrivate(request);
				})
					.then(function () {
					return result;
				});
			});
		};
		ClientRequestContext._run=function (ctxInitializer, runBody, numCleanupAttempts, retryDelay, onCleanupSuccess, onCleanupFailure) {
			if (numCleanupAttempts===void 0) { numCleanupAttempts=3; }
			if (retryDelay===void 0) { retryDelay=5000; }
			return ClientRequestContext._runCommon("run", null, ctxInitializer, 0, runBody, numCleanupAttempts, retryDelay, onCleanupSuccess, onCleanupFailure);
		};
		ClientRequestContext.isRequestUrlAndHeaderInfo=function (value) {
			return (typeof (value)==="object" &&
				value !==null &&
				Object.getPrototypeOf(value)===Object.getPrototypeOf({}) &&
				!OfficeExtension.Utility.isNullOrUndefined(value.url));
		};
		ClientRequestContext.isRequestUrlAndHeaderInfoResolver=function (value) {
			return (typeof (value)==="object" &&
				value !==null &&
				typeof (value._resolveRequestUrlAndHeaderInfo)==="function");
		};
		ClientRequestContext._runBatch=function (functionName, receivedRunArgs, ctxInitializer, numCleanupAttempts, retryDelay, onCleanupSuccess, onCleanupFailure) {
			if (numCleanupAttempts===void 0) { numCleanupAttempts=3; }
			if (retryDelay===void 0) { retryDelay=5000; }
			return ClientRequestContext._runBatchCommon(0, functionName, receivedRunArgs, ctxInitializer, numCleanupAttempts, retryDelay, onCleanupSuccess, onCleanupFailure);
		};
		ClientRequestContext._runExplicitBatch=function (functionName, receivedRunArgs, ctxInitializer, numCleanupAttempts, retryDelay, onCleanupSuccess, onCleanupFailure) {
			if (numCleanupAttempts===void 0) { numCleanupAttempts=3; }
			if (retryDelay===void 0) { retryDelay=5000; }
			return ClientRequestContext._runBatchCommon(1, functionName, receivedRunArgs, ctxInitializer, numCleanupAttempts, retryDelay, onCleanupSuccess, onCleanupFailure);
		};
		ClientRequestContext._runBatchCommon=function (batchMode, functionName, receivedRunArgs, ctxInitializer, numCleanupAttempts, retryDelay, onCleanupSuccess, onCleanupFailure) {
			if (numCleanupAttempts===void 0) { numCleanupAttempts=3; }
			if (retryDelay===void 0) { retryDelay=5000; }
			var ctxRetriever;
			var batch;
			var requestInfo=null;
			var argOffset=0;
			if (receivedRunArgs.length > 0 &&
				(typeof (receivedRunArgs[0])==="string" ||
					ClientRequestContext.isRequestUrlAndHeaderInfo(receivedRunArgs[0]) ||
					ClientRequestContext.isRequestUrlAndHeaderInfoResolver(receivedRunArgs[0]))) {
				requestInfo=receivedRunArgs[0];
				argOffset=1;
			}
			if (receivedRunArgs.length==argOffset+1) {
				ctxRetriever=ctxInitializer;
				batch=receivedRunArgs[argOffset+0];
			}
			else if (receivedRunArgs.length==argOffset+2) {
				if (OfficeExtension.Utility.isNullOrUndefined(receivedRunArgs[argOffset+0])) {
					ctxRetriever=ctxInitializer;
				}
				else if (receivedRunArgs[argOffset+0] instanceof OfficeExtension.ClientObject) {
					ctxRetriever=function () { return receivedRunArgs[argOffset+0].context; };
				}
				else if (receivedRunArgs[argOffset+0] instanceof ClientRequestContext) {
					ctxRetriever=function () { return receivedRunArgs[argOffset+0]; };
				}
				else if (Array.isArray(receivedRunArgs[argOffset+0])) {
					var array=receivedRunArgs[argOffset+0];
					if (array.length==0) {
						return ClientRequestContext.createErrorPromise(functionName);
					}
					for (var i=0; i < array.length; i++) {
						if (!(array[i] instanceof OfficeExtension.ClientObject)) {
							return ClientRequestContext.createErrorPromise(functionName);
						}
						if (array[i].context !=array[0].context) {
							return ClientRequestContext.createErrorPromise(functionName, OfficeExtension.ResourceStrings.invalidRequestContext);
						}
					}
					ctxRetriever=function () { return array[0].context; };
				}
				else {
					return ClientRequestContext.createErrorPromise(functionName);
				}
				batch=receivedRunArgs[argOffset+1];
			}
			else {
				return ClientRequestContext.createErrorPromise(functionName);
			}
			return ClientRequestContext._runCommon(functionName, requestInfo, ctxRetriever, batchMode, batch, numCleanupAttempts, retryDelay, onCleanupSuccess, onCleanupFailure);
		};
		ClientRequestContext.createErrorPromise=function (functionName, code) {
			if (code===void 0) { code=OfficeExtension.ResourceStrings.invalidArgument; }
			return OfficeExtension._Internal.OfficePromise.reject(OfficeExtension.Utility.createRuntimeError(code, OfficeExtension.Utility._getResourceString(code), functionName));
		};
		ClientRequestContext._runCommon=function (functionName, requestInfo, ctxRetriever, batchMode, runBody, numCleanupAttempts, retryDelay, onCleanupSuccess, onCleanupFailure) {
			if (ClientRequestContext._overrideSession) {
				requestInfo=ClientRequestContext._overrideSession;
			}
			var starterPromise=new OfficeExtension._Internal.OfficePromise(function (resolve, reject) { resolve(); });
			var ctx;
			var succeeded=false;
			var resultOrError;
			var previousBatchMode;
			return starterPromise
				.then(function () {
				ctx=ctxRetriever(requestInfo);
				if (ctx._autoCleanup) {
					return new OfficeExtension._Internal.OfficePromise(function (resolve, reject) {
						ctx._onRunFinishedNotifiers.push(function () {
							ctx._autoCleanup=true;
							resolve();
						});
					});
				}
				else {
					ctx._autoCleanup=true;
				}
			})
				.then(function () {
				if (typeof runBody !=='function') {
					return ClientRequestContext.createErrorPromise(functionName);
				}
				previousBatchMode=ctx.m_batchMode;
				ctx.m_batchMode=batchMode;
				var runBodyResult;
				if (batchMode==1) {
					runBodyResult=runBody(ctx.batch.bind(ctx));
				}
				else {
					runBodyResult=runBody(ctx);
				}
				if (OfficeExtension.Utility.isNullOrUndefined(runBodyResult) || (typeof runBodyResult.then !=='function')) {
					OfficeExtension.Utility.throwError(OfficeExtension.ResourceStrings.runMustReturnPromise);
				}
				return runBodyResult;
			})
				.then(function (runBodyResult) {
				if (batchMode===1) {
					return runBodyResult;
				}
				else {
					return ctx.sync(runBodyResult);
				}
			})
				.then(function (result) {
				succeeded=true;
				resultOrError=result;
			})
				.catch(function (error) {
				resultOrError=error;
			})
				.then(function () {
				var itemsToRemove=ctx.trackedObjects._retrieveAndClearAutoCleanupList();
				ctx._autoCleanup=false;
				ctx.m_batchMode=previousBatchMode;
				for (var key in itemsToRemove) {
					itemsToRemove[key]._objectPath.isValid=false;
				}
				var cleanupCounter=0;
				if (OfficeExtension.Utility._synchronousCleanup || ClientRequestContext.isRequestUrlAndHeaderInfoResolver(requestInfo)) {
					return attemptCleanup();
				}
				else {
					attemptCleanup();
				}
				function attemptCleanup() {
					cleanupCounter++;
					var savedPendingRequest=ctx.m_pendingRequest;
					var savedBatchMode=ctx.m_batchMode;
					var request=new OfficeExtension.ClientRequest(ctx);
					ctx.m_pendingRequest=request;
					ctx.m_batchMode=0;
					try {
						for (var key in itemsToRemove) {
							ctx.trackedObjects.remove(itemsToRemove[key]);
						}
					}
					finally {
						ctx.m_batchMode=savedBatchMode;
						ctx.m_pendingRequest=savedPendingRequest;
					}
					return ctx.syncPrivate(request)
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
					var func=ctx._onRunFinishedNotifiers.shift();
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
		ClientRequestContext.prototype._nextId=function () {
			return++this.m_nextId;
		};
		return ClientRequestContext;
	}());
	OfficeExtension.ClientRequestContext=ClientRequestContext;
})(OfficeExtension || (OfficeExtension={}));
var OfficeExtension;
(function (OfficeExtension) {
	var ClientResult=(function () {
		function ClientResult(type) {
			this.m_type=type;
		}
		Object.defineProperty(ClientResult.prototype, "value", {
			get: function () {
				if (!this.m_isLoaded) {
					throw new OfficeExtension._Internal.RuntimeError({
						code: OfficeExtension.ErrorCodes.valueNotLoaded,
						message: OfficeExtension.Utility._getResourceString(OfficeExtension.ResourceStrings.valueNotLoaded),
						debugInfo: {
							errorLocation: "clientResult.value"
						}
					});
				}
				return this.m_value;
			},
			enumerable: true,
			configurable: true
		});
		ClientResult.prototype._handleResult=function (value) {
			this.m_isLoaded=true;
			if (typeof (value)==="object" && value && value._IsNull) {
				return;
			}
			if (this.m_type===1) {
				this.m_value=OfficeExtension.Utility.adjustToDateTime(value);
			}
			else {
				this.m_value=value;
			}
		};
		return ClientResult;
	}());
	OfficeExtension.ClientResult=ClientResult;
	var RetrieveResultImpl=(function () {
		function RetrieveResultImpl(m_proxy, m_shouldPolyfill) {
			this.m_proxy=m_proxy;
			this.m_shouldPolyfill=m_shouldPolyfill;
			var scalarPropertyNames=m_proxy[OfficeExtension.Constants.scalarPropertyNames];
			var navigationPropertyNames=m_proxy[OfficeExtension.Constants.navigationPropertyNames];
			var typeName=m_proxy[OfficeExtension.Constants.className];
			var isCollection=m_proxy[OfficeExtension.Constants.isCollection];
			if (scalarPropertyNames) {
				for (var i=0; i < scalarPropertyNames.length; i++) {
					OfficeExtension.Utility.definePropertyThrowUnloadedException(this, typeName, scalarPropertyNames[i]);
				}
			}
			if (navigationPropertyNames) {
				for (var i=0; i < navigationPropertyNames.length; i++) {
					OfficeExtension.Utility.definePropertyThrowUnloadedException(this, typeName, navigationPropertyNames[i]);
				}
			}
			if (isCollection) {
				OfficeExtension.Utility.definePropertyThrowUnloadedException(this, typeName, OfficeExtension.Constants.itemsLowerCase);
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
					throw new OfficeExtension._Internal.RuntimeError({
						code: OfficeExtension.ErrorCodes.valueNotLoaded,
						message: OfficeExtension.Utility._getResourceString(OfficeExtension.ResourceStrings.valueNotLoaded),
						debugInfo: {
							errorLocation: "retrieveResult.$isNullObject"
						}
					});
				}
				return this.m_isNullObject;
			},
			enumerable: true,
			configurable: true
		});
		RetrieveResultImpl.prototype.toJSON=function () {
			if (!this.m_isLoaded) {
				return undefined;
			}
			if (this.m_isNullObject) {
				return null;
			}
			if (OfficeExtension.Utility.isUndefined(this.m_json)) {
				this.m_json=this.purifyJson(this.m_value);
			}
			return this.m_json;
		};
		RetrieveResultImpl.prototype.toString=function () {
			return JSON.stringify(this.toJSON());
		};
		RetrieveResultImpl.prototype._handleResult=function (value) {
			this.m_isLoaded=true;
			if (value===null || typeof (value)==="object" && value && value._IsNull) {
				this.m_isNullObject=true;
				value=null;
			}
			else {
				this.m_isNullObject=false;
			}
			if (this.m_shouldPolyfill) {
				value=this.changePropertyNameToCamelLowerCase(value);
			}
			this.m_value=value;
			this.m_proxy._handleRetrieveResult(value, this);
		};
		RetrieveResultImpl.prototype.changePropertyNameToCamelLowerCase=function (value) {
			var charCodeUnderscore=95;
			if (Array.isArray(value)) {
				var ret=[];
				for (var i=0; i < value.length; i++) {
					ret.push(this.changePropertyNameToCamelLowerCase(value[i]));
				}
				return ret;
			}
			else if (typeof (value)==="object" && value !==null) {
				var ret={};
				for (var key in value) {
					var propValue=value[key];
					if (key===OfficeExtension.Constants.items) {
						ret={};
						ret[OfficeExtension.Constants.itemsLowerCase]=this.changePropertyNameToCamelLowerCase(propValue);
						break;
					}
					else {
						var propName=OfficeExtension.Utility._toCamelLowerCase(key);
						ret[propName]=this.changePropertyNameToCamelLowerCase(propValue);
					}
				}
				return ret;
			}
			else {
				return value;
			}
		};
		RetrieveResultImpl.prototype.purifyJson=function (value) {
			var charCodeUnderscore=95;
			if (Array.isArray(value)) {
				var ret=[];
				for (var i=0; i < value.length; i++) {
					ret.push(this.purifyJson(value[i]));
				}
				return ret;
			}
			else if (typeof (value)==="object" && value !==null) {
				var ret={};
				for (var key in value) {
					if (key.charCodeAt(0) !==charCodeUnderscore) {
						var propValue=value[key];
						if (typeof (propValue)==="object" &&
							propValue !==null &&
							Array.isArray(propValue["items"])) {
							propValue=propValue["items"];
						}
						ret[key]=this.purifyJson(propValue);
					}
				}
				return ret;
			}
			else {
				return value;
			}
		};
		return RetrieveResultImpl;
	}());
	OfficeExtension.RetrieveResultImpl=RetrieveResultImpl;
})(OfficeExtension || (OfficeExtension={}));
var OfficeExtension;
(function (OfficeExtension) {
	var Constants=(function () {
		function Constants() {
		}
		Constants.flags="flags";
		Constants.getItemAt="GetItemAt";
		Constants.id="Id";
		Constants.idLowerCase="id";
		Constants.idPrivate="_Id";
		Constants.index="_Index";
		Constants.items="_Items";
		Constants.iterativeExecutor="IterativeExecutor";
		Constants.localDocument="http://document.localhost/";
		Constants.localDocumentApiPrefix="http://document.localhost/_api/";
		Constants.processQuery="ProcessQuery";
		Constants.referenceId="_ReferenceId";
		Constants.isTracked="_IsTracked";
		Constants.sourceLibHeader="SdkVersion";
		Constants.embeddingPageOrigin="EmbeddingPageOrigin";
		Constants.embeddingPageSessionInfo="EmbeddingPageSessionInfo";
		Constants.eventMessageCategory=65536;
		Constants.eventWorkbookId="Workbook";
		Constants.eventSourceRemote="Remote";
		Constants.itemsLowerCase="items";
		Constants.proxy="$proxy";
		Constants.scalarPropertyNames="_scalarPropertyNames";
		Constants.navigationPropertyNames="_navigationPropertyNames";
		Constants.className="_className";
		Constants.isCollection="_isCollection";
		Constants.scalarPropertyUpdateable="_scalarPropertyUpdateable";
		Constants.collectionPropertyPath="_collectionPropertyPath";
		return Constants;
	}());
	OfficeExtension.Constants=Constants;
})(OfficeExtension || (OfficeExtension={}));
var OfficeExtension;
(function (OfficeExtension) {
	var versionToken=1;
	var internalConfiguration={
		invokeRequestModifier: function (request) {
			request.DdaMethod.Version=versionToken;
			return request;
		},
		invokeResponseModifier: function (args) {
			versionToken=args.Version;
			if (args.Error) {
				args.error={};
				args.error.Code=args.Error;
			}
			return args;
		}
	};
	var EmbeddedApiStatus;
	(function (EmbeddedApiStatus) {
		EmbeddedApiStatus[EmbeddedApiStatus["Success"]=0]="Success";
		EmbeddedApiStatus[EmbeddedApiStatus["Timeout"]=1]="Timeout";
		EmbeddedApiStatus[EmbeddedApiStatus["InternalError"]=5001]="InternalError";
	})(EmbeddedApiStatus || (EmbeddedApiStatus={}));
	var CommunicationConstants;
	(function (CommunicationConstants) {
		CommunicationConstants.SendingId="sId";
		CommunicationConstants.RespondingId="rId";
		CommunicationConstants.CommandKey="command";
		CommunicationConstants.SessionInfoKey="sessionInfo";
		CommunicationConstants.ParamsKey="params";
		CommunicationConstants.ApiReadyCommand="apiready";
		CommunicationConstants.ExecuteMethodCommand="executeMethod";
		CommunicationConstants.GetAppContextCommand="getAppContext";
		CommunicationConstants.RegisterEventCommand="registerEvent";
		CommunicationConstants.UnregisterEventCommand="unregisterEvent";
		CommunicationConstants.FireEventCommand="fireEvent";
	})(CommunicationConstants || (CommunicationConstants={}));
	var EmbeddedSession=(function (_super) {
		__extends(EmbeddedSession, _super);
		function EmbeddedSession(url, options) {
			_super.call(this);
			this.m_chosenWindow=null;
			this.m_chosenOrigin=null;
			this.m_enabled=true;
			this.m_onMessageHandler=this._onMessage.bind(this);
			this.m_callbackList={};
			this.m_id=0;
			this.m_timeoutId=-1;
			this.m_appContext=null;
			this.m_url=url;
			this.m_options=options;
			if (!this.m_options) {
				this.m_options={ sessionKey: Math.random().toString() };
			}
			if (!this.m_options.sessionKey) {
				this.m_options.sessionKey=Math.random().toString();
			}
			if (!this.m_options.container) {
				this.m_options.container=document.body;
			}
			if (!this.m_options.timeoutInMilliseconds) {
				this.m_options.timeoutInMilliseconds=60000;
			}
			if (!this.m_options.height) {
				this.m_options.height="400px";
			}
			if (!this.m_options.width) {
				this.m_options.width="100%";
			}
		}
		EmbeddedSession.prototype._getIFrameSrc=function () {
			var origin=window.location.protocol+"//"+window.location.host;
			var toAppend=OfficeExtension.Constants.embeddingPageOrigin+"="+encodeURIComponent(origin)+"&"+OfficeExtension.Constants.embeddingPageSessionInfo+"="+encodeURIComponent(this.m_options.sessionKey);
			var useHash=false;
			if (this.m_url.toLowerCase().indexOf("/_layouts/preauth.aspx") > 0 ||
				this.m_url.toLowerCase().indexOf("/_layouts/15/preauth.aspx") > 0) {
				useHash=true;
			}
			var a=document.createElement("a");
			a.href=this.m_url;
			if (useHash) {
				if (a.hash.length===0 || a.hash==="#") {
					a.hash="#"+toAppend;
				}
				else {
					a.hash=a.hash+"&"+toAppend;
				}
			}
			else {
				if (a.search.length===0 || a.search==="?") {
					a.search="?"+toAppend;
				}
				else {
					a.search=a.search+"&"+toAppend;
				}
			}
			var iframeSrc=a.href;
			return iframeSrc;
		};
		EmbeddedSession.prototype.init=function () {
			var _this=this;
			window.addEventListener("message", this.m_onMessageHandler);
			var iframeSrc=this._getIFrameSrc();
			return new OfficeExtension._Internal.OfficePromise(function (resolve, reject) {
				var iframeElement=document.createElement("iframe");
				if (_this.m_options.id) {
					iframeElement.id=_this.m_options.id;
				}
				iframeElement.style.height=_this.m_options.height;
				iframeElement.style.width=_this.m_options.width;
				iframeElement.src=iframeSrc;
				_this.m_options.container.appendChild(iframeElement);
				_this.m_timeoutId=setTimeout(function () {
					_this.close();
					var err=OfficeExtension.Utility.createRuntimeError(OfficeExtension.ErrorCodes.timeout, OfficeExtension.Utility._getResourceString(OfficeExtension.ResourceStrings.timeout), "EmbeddedSession.init");
					reject(err);
				}, _this.m_options.timeoutInMilliseconds);
				_this.m_promiseResolver=resolve;
			});
		};
		EmbeddedSession.prototype._invoke=function (method, callback, params) {
			if (!this.m_enabled) {
				callback(EmbeddedApiStatus.InternalError, null);
				return;
			}
			if (internalConfiguration.invokeRequestModifier) {
				params=internalConfiguration.invokeRequestModifier(params);
			}
			this._sendMessageWithCallback(this.m_id++, method, params, function (args) {
				if (internalConfiguration.invokeResponseModifier) {
					args=internalConfiguration.invokeResponseModifier(args);
				}
				var errorCode=args["Error"];
				delete args["Error"];
				callback(errorCode || EmbeddedApiStatus.Success, args);
			});
		};
		EmbeddedSession.prototype.close=function () {
			window.removeEventListener("message", this.m_onMessageHandler);
			window.clearTimeout(this.m_timeoutId);
			this.m_enabled=false;
		};
		Object.defineProperty(EmbeddedSession.prototype, "eventRegistration", {
			get: function () {
				if (!this.m_sessionEventManager) {
					this.m_sessionEventManager=new OfficeExtension.EventRegistration(this._registerEventImpl.bind(this), this._unregisterEventImpl.bind(this));
				}
				return this.m_sessionEventManager;
			},
			enumerable: true,
			configurable: true
		});
		EmbeddedSession.prototype._createRequestExecutorOrNull=function () {
			return new EmbeddedRequestExecutor(this);
		};
		EmbeddedSession.prototype._resolveRequestUrlAndHeaderInfo=function () {
			return OfficeExtension.Utility._createPromiseFromResult(null);
		};
		EmbeddedSession.prototype._registerEventImpl=function (eventId, targetId) {
			var _this=this;
			return new OfficeExtension._Internal.OfficePromise(function (resolve, reject) {
				_this._sendMessageWithCallback(_this.m_id++, CommunicationConstants.RegisterEventCommand, { EventId: eventId, TargetId: targetId }, function () {
					resolve(null);
				});
			});
		};
		EmbeddedSession.prototype._unregisterEventImpl=function (eventId, targetId) {
			var _this=this;
			return new OfficeExtension._Internal.OfficePromise(function (resolve, reject) {
				_this._sendMessageWithCallback(_this.m_id++, CommunicationConstants.UnregisterEventCommand, { EventId: eventId, TargetId: targetId }, function () {
					resolve();
				});
			});
		};
		EmbeddedSession.prototype._onMessage=function (event) {
			var _this=this;
			if (!this.m_enabled) {
				return;
			}
			if (this.m_chosenWindow
				&& (this.m_chosenWindow !==event.source || this.m_chosenOrigin !==event.origin)) {
				return;
			}
			var eventData=event.data;
			if (eventData && eventData[CommunicationConstants.CommandKey]===CommunicationConstants.ApiReadyCommand) {
				if (!this.m_chosenWindow
					&& this._isValidDescendant(event.source)
					&& eventData[CommunicationConstants.SessionInfoKey]===this.m_options.sessionKey) {
					this.m_chosenWindow=event.source;
					this.m_chosenOrigin=event.origin;
					this._sendMessageWithCallback(this.m_id++, CommunicationConstants.GetAppContextCommand, null, function (appContext) {
						_this._setupContext(appContext);
						window.clearTimeout(_this.m_timeoutId);
						_this.m_promiseResolver();
					});
				}
				return;
			}
			if (eventData && eventData[CommunicationConstants.CommandKey]===CommunicationConstants.FireEventCommand) {
				var msg=eventData[CommunicationConstants.ParamsKey];
				var eventId=msg["EventId"];
				var targetId=msg["TargetId"];
				var data=msg["Data"];
				if (this.m_sessionEventManager) {
					var handlers=this.m_sessionEventManager.getHandlers(eventId, targetId);
					for (var i=0; i < handlers.length; i++) {
						handlers[i](data);
					}
				}
				return;
			}
			if (eventData && eventData.hasOwnProperty(CommunicationConstants.RespondingId)) {
				var rId=eventData[CommunicationConstants.RespondingId];
				var callback=this.m_callbackList[rId];
				if (typeof callback==="function") {
					callback(eventData[CommunicationConstants.ParamsKey]);
				}
				delete this.m_callbackList[rId];
			}
		};
		EmbeddedSession.prototype._sendMessageWithCallback=function (id, command, data, callback) {
			this.m_callbackList[id]=callback;
			var message={};
			message[CommunicationConstants.SendingId]=id;
			message[CommunicationConstants.CommandKey]=command;
			message[CommunicationConstants.ParamsKey]=data;
			this.m_chosenWindow.postMessage(JSON.stringify(message), this.m_chosenOrigin);
		};
		EmbeddedSession.prototype._isValidDescendant=function (wnd) {
			var container=this.m_options.container || document.body;
			function doesFrameWindow(containerWindow) {
				if (containerWindow===wnd) {
					return true;
				}
				for (var i=0, len=containerWindow.frames.length; i < len; i++) {
					if (doesFrameWindow(containerWindow.frames[i])) {
						return true;
					}
				}
				return false;
			}
			var iframes=container.getElementsByTagName("iframe");
			for (var i=0, len=iframes.length; i < len; i++) {
				if (doesFrameWindow(iframes[i].contentWindow)) {
					return true;
				}
			}
			return false;
		};
		EmbeddedSession.prototype._setupContext=function (appContext) {
			if (!(this.m_appContext=appContext)) {
				return;
			}
		};
		return EmbeddedSession;
	}(OfficeExtension.SessionBase));
	OfficeExtension.EmbeddedSession=EmbeddedSession;
	var EmbeddedRequestExecutor=(function () {
		function EmbeddedRequestExecutor(session) {
			this.m_session=session;
		}
		EmbeddedRequestExecutor.prototype.executeAsync=function (customData, requestFlags, requestMessage) {
			var _this=this;
			var messageSafearray=OfficeExtension.RichApiMessageUtility.buildMessageArrayForIRequestExecutor(customData, requestFlags, requestMessage, EmbeddedRequestExecutor.SourceLibHeaderValue);
			return new OfficeExtension._Internal.OfficePromise(function (resolve, reject) {
				_this.m_session._invoke(CommunicationConstants.ExecuteMethodCommand, function (status, result) {
					OfficeExtension.Utility.log("Response:");
					OfficeExtension.Utility.log(JSON.stringify(result));
					var response;
					if (status==EmbeddedApiStatus.Success) {
						response=OfficeExtension.RichApiMessageUtility.buildResponseOnSuccess(OfficeExtension.RichApiMessageUtility.getResponseBodyFromSafeArray(result.Data), OfficeExtension.RichApiMessageUtility.getResponseHeadersFromSafeArray(result.Data));
					}
					else {
						response=OfficeExtension.RichApiMessageUtility.buildResponseOnError(result.error.Code, result.error.Message);
					}
					resolve(response);
				}, EmbeddedRequestExecutor._transformMessageArrayIntoParams(messageSafearray));
			});
		};
		EmbeddedRequestExecutor._transformMessageArrayIntoParams=function (msgArray) {
			return {
				ArrayData: msgArray,
				DdaMethod: {
					DispatchId: EmbeddedRequestExecutor.DispidExecuteRichApiRequestMethod
				}
			};
		};
		EmbeddedRequestExecutor.DispidExecuteRichApiRequestMethod=93;
		EmbeddedRequestExecutor.SourceLibHeaderValue="Embedded";
		return EmbeddedRequestExecutor;
	}());
})(OfficeExtension || (OfficeExtension={}));
var OfficeExtension;
(function (OfficeExtension) {
	var _Internal;
	(function (_Internal) {
		var RuntimeError=(function (_super) {
			__extends(RuntimeError, _super);
			function RuntimeError(error) {
				_super.call(this, (typeof error==="string") ? error : error.message);
				this.name="OfficeExtension.Error";
				if (typeof error==="string") {
					this.message=error;
				}
				else {
					this.code=error.code;
					this.message=error.message;
					this.traceMessages=error.traceMessages || [];
					this.innerError=error.innerError || null;
					this.debugInfo=this._createDebugInfo(error.debugInfo || {});
				}
			}
			RuntimeError.prototype.toString=function () {
				return this.code+': '+this.message;
			};
			RuntimeError.prototype._createDebugInfo=function (partialDebugInfo) {
				var debugInfo={
					code: this.code,
					message: this.message,
					toString: function () {
						return JSON.stringify(this);
					}
				};
				for (var key in partialDebugInfo) {
					debugInfo[key]=partialDebugInfo[key];
				}
				if (this.innerError) {
					if (this.innerError instanceof OfficeExtension._Internal.RuntimeError) {
						debugInfo.innerError=this.innerError.debugInfo;
					}
					else {
						debugInfo.innerError=this.innerError;
					}
				}
				return debugInfo;
			};
			RuntimeError._createInvalidArgError=function (error) {
				return new _Internal.RuntimeError({
					code: OfficeExtension.ErrorCodes.invalidArgument,
					message: (OfficeExtension.Utility.isNullOrEmptyString(error.argumentName) ?
						OfficeExtension.Utility._getResourceString(OfficeExtension.ResourceStrings.invalidArgumentGeneric) :
						OfficeExtension.Utility._getResourceString(OfficeExtension.ResourceStrings.invalidArgument, error.argumentName)),
					debugInfo: error.errorLocation ? { errorLocation: error.errorLocation } : {},
					innerError: error.innerError
				});
			};
			return RuntimeError;
		}(Error));
		_Internal.RuntimeError=RuntimeError;
	})(_Internal=OfficeExtension._Internal || (OfficeExtension._Internal={}));
	OfficeExtension.Error=_Internal.RuntimeError;
})(OfficeExtension || (OfficeExtension={}));
var OfficeExtension;
(function (OfficeExtension) {
	var ErrorCodes=(function () {
		function ErrorCodes() {
		}
		ErrorCodes.accessDenied="AccessDenied";
		ErrorCodes.generalException="GeneralException";
		ErrorCodes.activityLimitReached="ActivityLimitReached";
		ErrorCodes.invalidObjectPath="InvalidObjectPath";
		ErrorCodes.propertyNotLoaded="PropertyNotLoaded";
		ErrorCodes.valueNotLoaded="ValueNotLoaded";
		ErrorCodes.invalidRequestContext="InvalidRequestContext";
		ErrorCodes.invalidArgument="InvalidArgument";
		ErrorCodes.runMustReturnPromise="RunMustReturnPromise";
		ErrorCodes.cannotRegisterEvent="CannotRegisterEvent";
		ErrorCodes.apiNotFound="ApiNotFound";
		ErrorCodes.connectionFailure="ConnectionFailure";
		ErrorCodes.timeout="Timeout";
		ErrorCodes.invalidOrTimedOutSession="InvalidOrTimedOutSession";
		ErrorCodes.cannotUpdateReadOnlyProperty="CannotUpdateReadOnlyProperty";
		return ErrorCodes;
	}());
	OfficeExtension.ErrorCodes=ErrorCodes;
})(OfficeExtension || (OfficeExtension={}));
var OfficeExtension;
(function (OfficeExtension) {
	var EventHandlers=(function () {
		function EventHandlers(context, parentObject, name, eventInfo) {
			var _this=this;
			this.m_id=context._nextId();
			this.m_context=context;
			this.m_name=name;
			this.m_handlers=[];
			this.m_registered=false;
			this.m_eventInfo=eventInfo;
			this.m_callback=function (args) {
				_this.m_eventInfo.eventArgsTransformFunc(args)
					.then(function (newArgs) { return _this.fireEvent(newArgs); });
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
		EventHandlers.prototype.add=function (handler) {
			var action=OfficeExtension.ActionFactory.createTraceAction(this.m_context, null, false);
			this.m_context._pendingRequest._addPendingEventHandlerAction(this, { id: action.actionInfo.Id, handler: handler, operation: 0 });
			return new OfficeExtension.EventHandlerResult(this.m_context, this, handler);
		};
		EventHandlers.prototype.remove=function (handler) {
			var action=OfficeExtension.ActionFactory.createTraceAction(this.m_context, null, false);
			this.m_context._pendingRequest._addPendingEventHandlerAction(this, { id: action.actionInfo.Id, handler: handler, operation: 1 });
		};
		EventHandlers.prototype.removeAll=function () {
			var action=OfficeExtension.ActionFactory.createTraceAction(this.m_context, null, false);
			this.m_context._pendingRequest._addPendingEventHandlerAction(this, { id: action.actionInfo.Id, handler: null, operation: 2 });
		};
		EventHandlers.prototype._processRegistration=function (req) {
			var _this=this;
			var ret=OfficeExtension.Utility._createPromiseFromResult(null);
			var actions=req._getPendingEventHandlerActions(this);
			if (!actions) {
				return ret;
			}
			var handlersResult=[];
			for (var i=0; i < this.m_handlers.length; i++) {
				handlersResult.push(this.m_handlers[i]);
			}
			var hasChange=false;
			for (var i=0; i < actions.length; i++) {
				if (req._responseTraceIds[actions[i].id]) {
					hasChange=true;
					switch (actions[i].operation) {
						case 0:
							handlersResult.push(actions[i].handler);
							break;
						case 1:
							for (var index=handlersResult.length - 1; index >=0; index--) {
								if (handlersResult[index]===actions[i].handler) {
									handlersResult.splice(index, 1);
									break;
								}
							}
							break;
						case 2:
							handlersResult=[];
							break;
					}
				}
			}
			if (hasChange) {
				if (!this.m_registered && handlersResult.length > 0) {
					ret=ret
						.then(function () { return _this.m_eventInfo.registerFunc(_this.m_callback); })
						.then(function () { return (_this.m_registered=true); });
				}
				else if (this.m_registered && handlersResult.length==0) {
					ret=ret
						.then(function () { return _this.m_eventInfo.unregisterFunc(_this.m_callback); })
						.catch(function (ex) {
						OfficeExtension.Utility.log("Error when unregister event: "+JSON.stringify(ex));
					})
						.then(function () { return (_this.m_registered=false); });
				}
				ret=ret
					.then(function () { return (_this.m_handlers=handlersResult); });
			}
			return ret;
		};
		EventHandlers.prototype.fireEvent=function (args) {
			var promises=[];
			for (var i=0; i < this.m_handlers.length; i++) {
				var handler=this.m_handlers[i];
				var p=OfficeExtension.Utility._createPromiseFromResult(null)
					.then(this.createFireOneEventHandlerFunc(handler, args))
					.catch(function (ex) {
					OfficeExtension.Utility.log("Error when invoke handler: "+JSON.stringify(ex));
				});
				promises.push(p);
			}
			OfficeExtension._Internal.OfficePromise.all(promises);
		};
		EventHandlers.prototype.createFireOneEventHandlerFunc=function (handler, args) {
			return function () { return handler(args); };
		};
		return EventHandlers;
	}());
	OfficeExtension.EventHandlers=EventHandlers;
})(OfficeExtension || (OfficeExtension={}));
var OfficeExtension;
(function (OfficeExtension) {
	var EventHandlerResult=(function () {
		function EventHandlerResult(context, handlers, handler) {
			this.m_context=context;
			this.m_allHandlers=handlers;
			this.m_handler=handler;
		}
		Object.defineProperty(EventHandlerResult.prototype, "context", {
			get: function () {
				return this.m_context;
			},
			enumerable: true,
			configurable: true
		});
		EventHandlerResult.prototype.remove=function () {
			if (this.m_allHandlers && this.m_handler) {
				this.m_allHandlers.remove(this.m_handler);
				this.m_allHandlers=null;
				this.m_handler=null;
			}
		};
		return EventHandlerResult;
	}());
	OfficeExtension.EventHandlerResult=EventHandlerResult;
})(OfficeExtension || (OfficeExtension={}));
var OfficeExtension;
(function (OfficeExtension) {
	var _Internal;
	(function (_Internal) {
		var OfficeJsEventRegistration=(function () {
			function OfficeJsEventRegistration() {
			}
			OfficeJsEventRegistration.prototype.register=function (eventId, targetId, handler) {
				switch (eventId) {
					case 4:
						return OfficeExtension.Utility.promisify(function (callback) { return Office.context.document.bindings.getByIdAsync(targetId, callback); })
							.then(function (officeBinding) {
							return OfficeExtension.Utility.promisify(function (callback) { return officeBinding.addHandlerAsync(Office.EventType.BindingDataChanged, handler, callback); });
						});
					case 3:
						return OfficeExtension.Utility.promisify(function (callback) { return Office.context.document.bindings.getByIdAsync(targetId, callback); })
							.then(function (officeBinding) {
							return OfficeExtension.Utility.promisify(function (callback) { return officeBinding.addHandlerAsync(Office.EventType.BindingSelectionChanged, handler, callback); });
						});
					case 2:
						return OfficeExtension.Utility.promisify(function (callback) { return Office.context.document.addHandlerAsync(Office.EventType.DocumentSelectionChanged, handler, callback); });
					case 1:
						return OfficeExtension.Utility.promisify(function (callback) { return Office.context.document.settings.addHandlerAsync(Office.EventType.SettingsChanged, handler, callback); });
					case 5:
						return OfficeExtension.Utility.promisify(function (callback) { return OSF.DDA.RichApi.richApiMessageManager.addHandlerAsync("richApiMessage", handler, callback); });
					case 13:
						return OfficeExtension.Utility.promisify(function (callback) { return Office.context.document.addHandlerAsync(Office.EventType.ObjectDeleted, handler, { id: targetId }, callback); });
					case 14:
						return OfficeExtension.Utility.promisify(function (callback) { return Office.context.document.addHandlerAsync(Office.EventType.ObjectSelectionChanged, handler, { id: targetId }, callback); });
					case 15:
						return OfficeExtension.Utility.promisify(function (callback) { return Office.context.document.addHandlerAsync(Office.EventType.ObjectDataChanged, handler, { id: targetId }, callback); });
					case 16:
						return OfficeExtension.Utility.promisify(function (callback) { return Office.context.document.addHandlerAsync(Office.EventType.ContentControlAdded, handler, { id: targetId }, callback); });
					default:
						throw _Internal.RuntimeError._createInvalidArgError("eventId");
				}
			};
			OfficeJsEventRegistration.prototype.unregister=function (eventId, targetId, handler) {
				switch (eventId) {
					case 4:
						return OfficeExtension.Utility.promisify(function (callback) { return Office.context.document.bindings.getByIdAsync(targetId, callback); })
							.then(function (officeBinding) {
							return OfficeExtension.Utility.promisify(function (callback) { return officeBinding.removeHandlerAsync(Office.EventType.BindingDataChanged, { handler: handler }, callback); });
						});
					case 3:
						return OfficeExtension.Utility.promisify(function (callback) { return Office.context.document.bindings.getByIdAsync(targetId, callback); })
							.then(function (officeBinding) {
							return OfficeExtension.Utility.promisify(function (callback) { return officeBinding.removeHandlerAsync(Office.EventType.BindingSelectionChanged, { handler: handler }, callback); });
						});
					case 2:
						return OfficeExtension.Utility.promisify(function (callback) { return Office.context.document.removeHandlerAsync(Office.EventType.DocumentSelectionChanged, { handler: handler }, callback); });
					case 1:
						return OfficeExtension.Utility.promisify(function (callback) { return Office.context.document.settings.removeHandlerAsync(Office.EventType.SettingsChanged, { handler: handler }, callback); });
					case 5:
						return OfficeExtension.Utility.promisify(function (callback) { return OSF.DDA.RichApi.richApiMessageManager.removeHandlerAsync("richApiMessage", { handler: handler }, callback); });
					case 13:
						return OfficeExtension.Utility.promisify(function (callback) { return Office.context.document.removeHandlerAsync(Office.EventType.ObjectDeleted, { id: targetId, handler: handler }, callback); });
					case 14:
						return OfficeExtension.Utility.promisify(function (callback) { return Office.context.document.removeHandlerAsync(Office.EventType.ObjectSelectionChanged, { id: targetId, handler: handler }, callback); });
					case 15:
						return OfficeExtension.Utility.promisify(function (callback) { return Office.context.document.removeHandlerAsync(Office.EventType.ObjectDataChanged, { id: targetId, handler: handler }, callback); });
					case 16:
						return OfficeExtension.Utility.promisify(function (callback) { return Office.context.document.removeHandlerAsync(Office.EventType.ContentControlAdded, { id: targetId, handler: handler }, callback); });
					default:
						throw _Internal.RuntimeError._createInvalidArgError("eventId");
				}
			};
			return OfficeJsEventRegistration;
		}());
		_Internal.officeJsEventRegistration=new OfficeJsEventRegistration();
	})(_Internal=OfficeExtension._Internal || (OfficeExtension._Internal={}));
	var EventRegistration=(function () {
		function EventRegistration(registerEventImpl, unregisterEventImpl) {
			this.m_handlersByEventByTarget={};
			this.m_registerEventImpl=registerEventImpl;
			this.m_unregisterEventImpl=unregisterEventImpl;
		}
		EventRegistration.prototype.getHandlers=function (eventId, targetId) {
			if (OfficeExtension.Utility.isNullOrUndefined(targetId)) {
				targetId="";
			}
			var handlersById=this.m_handlersByEventByTarget[eventId];
			if (!handlersById) {
				handlersById={};
				this.m_handlersByEventByTarget[eventId]=handlersById;
			}
			var handlers=handlersById[targetId];
			if (!handlers) {
				handlers=[];
				handlersById[targetId]=handlers;
			}
			return handlers;
		};
		EventRegistration.prototype.register=function (eventId, targetId, handler) {
			if (!handler) {
				throw _Internal.RuntimeError._createInvalidArgError("handler");
			}
			var handlers=this.getHandlers(eventId, targetId);
			handlers.push(handler);
			if (handlers.length===1) {
				return this.m_registerEventImpl(eventId, targetId);
			}
			return OfficeExtension.Utility._createPromiseFromResult(null);
		};
		EventRegistration.prototype.unregister=function (eventId, targetId, handler) {
			if (!handler) {
				throw _Internal.RuntimeError._createInvalidArgError("handler");
			}
			var handlers=this.getHandlers(eventId, targetId);
			for (var index=handlers.length - 1; index >=0; index--) {
				if (handlers[index]===handler) {
					handlers.splice(index, 1);
					break;
				}
			}
			if (handlers.length===0) {
				return this.m_unregisterEventImpl(eventId, targetId);
			}
			return OfficeExtension.Utility._createPromiseFromResult(null);
		};
		return EventRegistration;
	}());
	OfficeExtension.EventRegistration=EventRegistration;
})(OfficeExtension || (OfficeExtension={}));
var OfficeExtension;
(function (OfficeExtension) {
	var GenericEventRegistration=(function () {
		function GenericEventRegistration() {
			this.m_eventRegistration=new OfficeExtension.EventRegistration(this._registerEventImpl.bind(this), this._unregisterEventImpl.bind(this));
			this.m_richApiMessageHandler=this._handleRichApiMessage.bind(this);
		}
		GenericEventRegistration.prototype.ready=function () {
			var _this=this;
			if (!this.m_ready) {
				if (GenericEventRegistration._testReadyImpl) {
					this.m_ready=GenericEventRegistration._testReadyImpl()
						.then(function () {
						_this.m_isReady=true;
					});
				}
				else {
					this.m_ready=OfficeExtension._Internal.officeJsEventRegistration.register(5, "", this.m_richApiMessageHandler)
						.then(function () {
						_this.m_isReady=true;
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
		GenericEventRegistration.prototype.register=function (eventId, targetId, handler) {
			var _this=this;
			return this.ready()
				.then(function () { return _this.m_eventRegistration.register(eventId, targetId, handler); });
		};
		GenericEventRegistration.prototype.unregister=function (eventId, targetId, handler) {
			var _this=this;
			return this.ready()
				.then(function () { return _this.m_eventRegistration.unregister(eventId, targetId, handler); });
		};
		GenericEventRegistration.prototype._registerEventImpl=function (eventId, targetId) {
			return OfficeExtension.Utility._createPromiseFromResult(null);
		};
		GenericEventRegistration.prototype._unregisterEventImpl=function (eventId, targetId) {
			return OfficeExtension.Utility._createPromiseFromResult(null);
		};
		GenericEventRegistration.prototype._handleRichApiMessage=function (msg) {
			if (msg && msg.entries) {
				for (var entryIndex=0; entryIndex < msg.entries.length; entryIndex++) {
					var entry=msg.entries[entryIndex];
					if (entry.messageCategory==OfficeExtension.Constants.eventMessageCategory) {
						if (OfficeExtension.Utility._logEnabled) {
							OfficeExtension.Utility.log(JSON.stringify(entry));
						}
						var funcs=this.m_eventRegistration.getHandlers(entry.messageType, entry.targetId);
						if (funcs.length > 0) {
							var arg=JSON.parse(entry.message);
							if (entry.isRemoteOverride) {
								arg.source=OfficeExtension.Constants.eventSourceRemote;
							}
							for (var i=0; i < funcs.length; i++) {
								funcs[i](arg);
							}
						}
					}
				}
			}
		};
		GenericEventRegistration.getGenericEventRegistration=function () {
			if (!GenericEventRegistration.s_genericEventRegistration) {
				GenericEventRegistration.s_genericEventRegistration=new GenericEventRegistration();
			}
			return GenericEventRegistration.s_genericEventRegistration;
		};
		GenericEventRegistration.richApiMessageEventCategory=65536;
		return GenericEventRegistration;
	}());
	function _testSetRichApiMessageReadyImpl(impl) {
		GenericEventRegistration._testReadyImpl=impl;
	}
	OfficeExtension._testSetRichApiMessageReadyImpl=_testSetRichApiMessageReadyImpl;
	function _testTriggerRichApiMessageEvent(msg) {
		GenericEventRegistration.getGenericEventRegistration()._handleRichApiMessage(msg);
	}
	OfficeExtension._testTriggerRichApiMessageEvent=_testTriggerRichApiMessageEvent;
	var GenericEventHandlers=(function (_super) {
		__extends(GenericEventHandlers, _super);
		function GenericEventHandlers(context, parentObject, name, eventInfo) {
			_super.call(this, context, parentObject, name, eventInfo);
			this.m_genericEventInfo=eventInfo;
		}
		GenericEventHandlers.prototype.add=function (handler) {
			var _this=this;
			if ((this._handlers.length==0) && this.m_genericEventInfo.registerFunc) {
				this.m_genericEventInfo.registerFunc();
			}
			if (!GenericEventRegistration.getGenericEventRegistration().isReady) {
				this._context._pendingRequest._addPreSyncPromise(GenericEventRegistration.getGenericEventRegistration().ready());
			}
			OfficeExtension.ActionFactory.createTraceMarkerForCallback(this._context, function () {
				_this._handlers.push(handler);
				if (_this._handlers.length==1) {
					GenericEventRegistration.getGenericEventRegistration().register(_this.m_genericEventInfo.eventType, _this.m_genericEventInfo.getTargetIdFunc(), _this._callback);
				}
			});
			return new OfficeExtension.EventHandlerResult(this._context, this, handler);
		};
		GenericEventHandlers.prototype.remove=function (handler) {
			var _this=this;
			if ((this._handlers.length==1) && this.m_genericEventInfo.unregisterFunc) {
				this.m_genericEventInfo.unregisterFunc();
			}
			OfficeExtension.ActionFactory.createTraceMarkerForCallback(this._context, function () {
				var handlers=_this._handlers;
				for (var index=handlers.length - 1; index >=0; index--) {
					if (handlers[index]===handler) {
						handlers.splice(index, 1);
						break;
					}
				}
				if (handlers.length==0) {
					GenericEventRegistration.getGenericEventRegistration().unregister(_this.m_genericEventInfo.eventType, _this.m_genericEventInfo.getTargetIdFunc(), _this._callback);
				}
			});
		};
		GenericEventHandlers.prototype.removeAll=function () {
		};
		return GenericEventHandlers;
	}(OfficeExtension.EventHandlers));
	OfficeExtension.GenericEventHandlers=GenericEventHandlers;
})(OfficeExtension || (OfficeExtension={}));
var OfficeExtension;
(function (OfficeExtension) {
	var HttpRequestExecutor=(function () {
		function HttpRequestExecutor() {
		}
		HttpRequestExecutor.prototype.executeAsync=function (customData, requestFlags, requestMessage) {
			var requestMessageText=JSON.stringify(requestMessage.Body);
			var url=requestMessage.Url;
			if (url.charAt(url.length - 1) !="/") {
				url=url+"/";
			}
			url=url+OfficeExtension.Constants.processQuery;
			url=url+"?"+OfficeExtension.Constants.flags+"="+requestFlags.toString();
			var requestInfo={
				method: "POST",
				url: url,
				headers: {},
				body: requestMessageText
			};
			requestInfo.headers[OfficeExtension.Constants.sourceLibHeader]=HttpRequestExecutor.SourceLibHeaderValue;
			requestInfo.headers["CONTENT-TYPE"]="application/json";
			if (requestMessage.Headers) {
				for (var key in requestMessage.Headers) {
					requestInfo.headers[key]=requestMessage.Headers[key];
				}
			}
			return OfficeExtension.HttpUtility.sendRequest(requestInfo)
				.then(function (responseInfo) {
				var response;
				if (responseInfo.statusCode===200) {
					response={ ErrorCode: null, ErrorMessage: null, Headers: responseInfo.headers, Body: JSON.parse(responseInfo.body) };
				}
				else {
					OfficeExtension.Utility.log("Error Response:"+responseInfo.body);
					var error=OfficeExtension.Utility._parseErrorResponse(responseInfo);
					response={
						ErrorCode: error.errorCode,
						ErrorMessage: error.errorMessage,
						Headers: responseInfo.headers,
						Body: null
					};
				}
				return response;
			});
		};
		HttpRequestExecutor.SourceLibHeaderValue="officejs-rest";
		return HttpRequestExecutor;
	}());
	OfficeExtension.HttpRequestExecutor=HttpRequestExecutor;
})(OfficeExtension || (OfficeExtension={}));
var OfficeExtension;
(function (OfficeExtension) {
	var HttpUtility=(function () {
		function HttpUtility() {
		}
		HttpUtility.setCustomSendRequestFunc=function (func) {
			HttpUtility.s_customSendRequestFunc=func;
		};
		HttpUtility.xhrSendRequestFunc=function (request) {
			return new OfficeExtension._Internal.OfficePromise(function (resolve, reject) {
				var xhr=new XMLHttpRequest();
				xhr.open(request.method, request.url);
				xhr.onload=function () {
					var resp={
						statusCode: xhr.status,
						headers: OfficeExtension.Utility._parseHttpResponseHeaders(xhr.getAllResponseHeaders()),
						body: xhr.responseText
					};
					resolve(resp);
				};
				xhr.onerror=function () {
					reject(new OfficeExtension._Internal.RuntimeError({
						code: OfficeExtension.ErrorCodes.connectionFailure,
						message: OfficeExtension.Utility._getResourceString(OfficeExtension.ResourceStrings.connectionFailureWithStatus, xhr.statusText)
					}));
				};
				if (request.headers) {
					for (var key in request.headers) {
						xhr.setRequestHeader(key, request.headers[key]);
					}
				}
				xhr.send(request.body);
			});
		};
		HttpUtility.sendRequest=function (request) {
			HttpUtility.validateAndNormalizeRequest(request);
			var func=HttpUtility.s_customSendRequestFunc;
			if (!func) {
				func=HttpUtility.xhrSendRequestFunc;
			}
			return func(request);
		};
		HttpUtility.setCustomSendLocalDocumentRequestFunc=function (func) {
			HttpUtility.s_customSendLocalDocumentRequestFunc=func;
		};
		HttpUtility.sendLocalDocumentRequest=function (request) {
			HttpUtility.validateAndNormalizeRequest(request);
			var func;
			func=HttpUtility.s_customSendLocalDocumentRequestFunc || HttpUtility.officeJsSendLocalDocumentRequestFunc;
			return func(request);
		};
		HttpUtility.officeJsSendLocalDocumentRequestFunc=function (request) {
			request=OfficeExtension.Utility._validateLocalDocumentRequest(request);
			var requestSafeArray=OfficeExtension.Utility._buildRequestMessageSafeArray(request);
			return new OfficeExtension._Internal.OfficePromise(function (resolve, reject) {
				OSF.DDA.RichApi.executeRichApiRequestAsync(requestSafeArray, function (asyncResult) {
					var response;
					if (asyncResult.status=="succeeded") {
						response=							{
								statusCode: OfficeExtension.RichApiMessageUtility.getResponseStatusCode(asyncResult),
								headers: OfficeExtension.RichApiMessageUtility.getResponseHeaders(asyncResult),
								body: OfficeExtension.RichApiMessageUtility.getResponseBody(asyncResult)
							};
					}
					else {
						response=OfficeExtension.RichApiMessageUtility.buildHttpResponseFromOfficeJsError(asyncResult.error.code, asyncResult.error.message);
					}
					OfficeExtension.Utility.log(JSON.stringify(response));
					resolve(response);
				});
			});
		};
		HttpUtility.validateAndNormalizeRequest=function (request) {
			if (OfficeExtension.Utility.isNullOrUndefined(request)) {
				throw OfficeExtension._Internal.RuntimeError._createInvalidArgError({
					argumentName: "request"
				});
			}
			if (OfficeExtension.Utility.isNullOrEmptyString(request.method)) {
				request.method="GET";
			}
			request.method=request.method.toUpperCase();
		};
		HttpUtility.logRequest=function (request) {
			if (OfficeExtension.Utility._logEnabled) {
				OfficeExtension.Utility.log("---HTTP Request---");
				OfficeExtension.Utility.log(request.method+" "+request.url);
				if (request.headers) {
					for (var key in request.headers) {
						OfficeExtension.Utility.log(key+": "+request.headers[key]);
					}
				}
				if (HttpUtility._logBody) {
					OfficeExtension.Utility.log(request.body);
				}
			}
		};
		HttpUtility.logResponse=function (response) {
			if (OfficeExtension.Utility._logEnabled) {
				OfficeExtension.Utility.log("---HTTP Response---");
				OfficeExtension.Utility.log(""+response.statusCode);
				if (response.headers) {
					for (var key in response.headers) {
						OfficeExtension.Utility.log(key+": "+response.headers[key]);
					}
				}
				if (HttpUtility._logBody) {
					OfficeExtension.Utility.log(response.body);
				}
			}
		};
		HttpUtility._logBody=false;
		return HttpUtility;
	}());
	OfficeExtension.HttpUtility=HttpUtility;
})(OfficeExtension || (OfficeExtension={}));
var OfficeExtension;
(function (OfficeExtension) {
	var InstantiateActionResultHandler=(function () {
		function InstantiateActionResultHandler(clientObject) {
			this.m_clientObject=clientObject;
		}
		InstantiateActionResultHandler.prototype._handleResult=function (value) {
			this.m_clientObject._handleIdResult(value);
		};
		return InstantiateActionResultHandler;
	}());
	OfficeExtension.InstantiateActionResultHandler=InstantiateActionResultHandler;
})(OfficeExtension || (OfficeExtension={}));
var OfficeExtension;
(function (OfficeExtension) {
	;
})(OfficeExtension || (OfficeExtension={}));
var OfficeExtension;
(function (OfficeExtension) {
	var HostBridgeRequestExecutor=(function () {
		function HostBridgeRequestExecutor(session) {
			this.m_session=session;
		}
		HostBridgeRequestExecutor.prototype.executeAsync=function (customData, requestFlags, requestMessage) {
			var httpRequestInfo={
				url: OfficeExtension.Constants.processQuery,
				method: "POST",
				headers: requestMessage.Headers,
				body: requestMessage.Body
			};
			var message={
				id: HostBridgeSession.nextId(),
				type: 1,
				flags: requestFlags,
				message: httpRequestInfo
			};
			OfficeExtension.Utility.log(JSON.stringify(message));
			return this.m_session.sendMessageToHost(message)
				.then(function (nativeBridgeResponse) {
				OfficeExtension.Utility.log("Received response: "+JSON.stringify(nativeBridgeResponse));
				var responseInfo=nativeBridgeResponse.message;
				var response;
				if (responseInfo.statusCode===200) {
					response={ ErrorCode: null, ErrorMessage: null, Headers: responseInfo.headers, Body: responseInfo.body };
				}
				else {
					OfficeExtension.Utility.log("Error Response:"+responseInfo.body);
					var error=OfficeExtension.Utility._parseErrorResponse(responseInfo);
					response={
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
	var HostBridgeSession=(function (_super) {
		__extends(HostBridgeSession, _super);
		function HostBridgeSession(bridge) {
			var _this=this;
			_super.call(this);
			this.m_promiseResolver={};
			this.m_bridge=bridge;
			this.m_bridge.onMessageFromHost=function (msg) {
				_this.onMessageFromHost(msg);
			};
		}
		HostBridgeSession.prototype._resolveRequestUrlAndHeaderInfo=function () {
			return OfficeExtension.Utility._createPromiseFromResult(null);
		};
		HostBridgeSession.prototype._createRequestExecutorOrNull=function () {
			OfficeExtension.Utility.log("NativeBridgeSession::CreateRequestExecutor");
			return new HostBridgeRequestExecutor(this);
		};
		Object.defineProperty(HostBridgeSession.prototype, "eventRegistration", {
			get: function () {
				return OfficeExtension._Internal.officeJsEventRegistration;
			},
			enumerable: true,
			configurable: true
		});
		HostBridgeSession.init=function (bridge) {
			if (bridge && typeof (bridge)==="object") {
				var session=new HostBridgeSession(bridge);
				OfficeExtension.ClientRequestContext._overrideSession=session;
				OfficeExtension.HttpUtility.setCustomSendLocalDocumentRequestFunc(function (request) {
					var bridgeMessage={
						id: HostBridgeSession.nextId(),
						type: 1,
						flags: 0,
						message: request
					};
					return session.sendMessageToHost(bridgeMessage)
						.then(function (bridgeResponse) {
						var responseInfo=bridgeResponse.message;
						return responseInfo;
					});
				});
			}
		};
		HostBridgeSession.prototype.sendMessageToHost=function (message) {
			var _this=this;
			this.m_bridge.sendMessageToHost(JSON.stringify(message));
			var ret=new OfficeExtension._Internal.OfficePromise(function (resolve, reject) {
				_this.m_promiseResolver[message.id]=resolve;
			});
			return ret;
		};
		HostBridgeSession.prototype.onMessageFromHost=function (messageText) {
			if (messageText==="test") {
				if (HostBridgeTest._testFunc) {
					HostBridgeTest._testFunc();
				}
			}
			else {
				var message=JSON.parse(messageText);
				if (typeof (message.id)==="number") {
					var resolve=this.m_promiseResolver[message.id];
					if (resolve) {
						resolve(message);
					}
					delete this.m_promiseResolver[message.id];
				}
			}
		};
		HostBridgeSession.nextId=function () {
			return HostBridgeSession.s_nextId++;
		};
		HostBridgeSession.s_nextId=1;
		return HostBridgeSession;
	}(OfficeExtension.SessionBase));
	var HostBridge=(function () {
		function HostBridge() {
		}
		HostBridge.init=function (bridge) {
			HostBridgeSession.init(bridge);
		};
		return HostBridge;
	}());
	OfficeExtension.HostBridge=HostBridge;
	if (typeof (_richApiNativeBridge)==="object" && _richApiNativeBridge) {
		HostBridge.init(_richApiNativeBridge);
	}
	var HostBridgeTest=(function () {
		function HostBridgeTest() {
		}
		HostBridgeTest.setTestFunc=function (func) {
			HostBridgeTest._testFunc=func;
		};
		return HostBridgeTest;
	}());
	OfficeExtension.HostBridgeTest=HostBridgeTest;
})(OfficeExtension || (OfficeExtension={}));
var OfficeExtension;
(function (OfficeExtension) {
	var ObjectPath=(function () {
		function ObjectPath(objectPathInfo, parentObjectPath, isCollection, isInvalidAfterRequest) {
			this.m_objectPathInfo=objectPathInfo;
			this.m_parentObjectPath=parentObjectPath;
			this.m_isWriteOperation=false;
			this.m_isCollection=isCollection;
			this.m_isInvalidAfterRequest=isInvalidAfterRequest;
			this.m_isValid=true;
		}
		Object.defineProperty(ObjectPath.prototype, "objectPathInfo", {
			get: function () {
				return this.m_objectPathInfo;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ObjectPath.prototype, "isWriteOperation", {
			get: function () {
				return this.m_isWriteOperation;
			},
			set: function (value) {
				this.m_isWriteOperation=value;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ObjectPath.prototype, "isRestrictedResourceAccess", {
			get: function () {
				return this.m_isRestrictedResourceAccess;
			},
			set: function (value) {
				this.m_isRestrictedResourceAccess=value;
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
				this.m_argumentObjectPaths=value;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ObjectPath.prototype, "isValid", {
			get: function () {
				return this.m_isValid;
			},
			set: function (value) {
				this.m_isValid=value;
				if (!value &&
					this.m_objectPathInfo.ObjectPathType===6 &&
					this.m_savedObjectPathInfo) {
					ObjectPath.copyObjectPathInfo(this.m_savedObjectPathInfo.pathInfo, this.m_objectPathInfo);
					this.m_parentObjectPath=this.m_savedObjectPathInfo.parent;
					this.m_isValid=true;
					this.m_savedObjectPathInfo=null;
				}
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ObjectPath.prototype, "getByIdMethodName", {
			get: function () {
				return this.m_getByIdMethodName;
			},
			set: function (value) {
				this.m_getByIdMethodName=value;
			},
			enumerable: true,
			configurable: true
		});
		ObjectPath.prototype._updateAsNullObject=function () {
			this.resetForUpdateUsingObjectData();
			this.m_objectPathInfo.ObjectPathType=7;
			this.m_objectPathInfo.Name="";
			this.m_parentObjectPath=null;
		};
		ObjectPath.prototype.updateUsingObjectData=function (value, clientObject) {
			var referenceId=value[OfficeExtension.Constants.referenceId];
			if (!OfficeExtension.Utility.isNullOrEmptyString(referenceId)) {
				if (!this.m_savedObjectPathInfo &&
					!this.isInvalidAfterRequest &&
					ObjectPath.isRestorableObjectPath(this.m_objectPathInfo.ObjectPathType)) {
					var pathInfo={};
					ObjectPath.copyObjectPathInfo(this.m_objectPathInfo, pathInfo);
					this.m_savedObjectPathInfo={
						pathInfo: pathInfo,
						parent: this.m_parentObjectPath
					};
				}
				this.resetForUpdateUsingObjectData();
				this.m_objectPathInfo.ObjectPathType=6;
				this.m_objectPathInfo.Name=referenceId;
				delete this.m_objectPathInfo.ParentObjectPathId;
				this.m_parentObjectPath=null;
				return;
			}
			var parentIsCollection=this.parentObjectPath && this.parentObjectPath.isCollection;
			var getByIdMethodName=this.getByIdMethodName;
			if (parentIsCollection || !OfficeExtension.Utility.isNullOrEmptyString(getByIdMethodName)) {
				var id=OfficeExtension.Utility.tryGetObjectIdFromLoadOrRetrieveResult(value);
				if (!OfficeExtension.Utility.isNullOrUndefined(id)) {
					this.resetForUpdateUsingObjectData();
					if (!OfficeExtension.Utility.isNullOrEmptyString(getByIdMethodName)) {
						this.m_objectPathInfo.ObjectPathType=3;
						this.m_objectPathInfo.Name=getByIdMethodName;
						this.m_getByIdMethodName=null;
					}
					else {
						this.m_objectPathInfo.ObjectPathType=5;
						this.m_objectPathInfo.Name="";
					}
					this.m_objectPathInfo.ArgumentInfo.Arguments=[id];
					return;
				}
			}
			var collectionPropertyPath=clientObject[OfficeExtension.Constants.collectionPropertyPath];
			if (this.m_objectPathInfo.ObjectPathType !==5 && !OfficeExtension.Utility.isNullOrEmptyString(collectionPropertyPath)) {
				var id=OfficeExtension.Utility.tryGetObjectIdFromLoadOrRetrieveResult(value);
				if (!OfficeExtension.Utility.isNullOrUndefined(id)) {
					var propNames=collectionPropertyPath.split(".");
					var parent_1=clientObject.context[propNames[0]];
					for (var i=1; i < propNames.length; i++) {
						parent_1=parent_1[propNames[i]];
					}
					this.resetForUpdateUsingObjectData();
					this.m_parentObjectPath=parent_1._objectPath;
					this.m_objectPathInfo.ParentObjectPathId=this.m_parentObjectPath.objectPathInfo.Id;
					this.m_objectPathInfo.ObjectPathType=5;
					this.m_objectPathInfo.Name="";
					this.m_objectPathInfo.ArgumentInfo.Arguments=[id];
					return;
				}
			}
		};
		ObjectPath.prototype.resetForUpdateUsingObjectData=function () {
			this.m_isInvalidAfterRequest=false;
			this.m_isValid=true;
			this.m_isWriteOperation=false;
			this.m_objectPathInfo.ArgumentInfo={};
			this.m_argumentObjectPaths=null;
		};
		ObjectPath.isRestorableObjectPath=function (objectPathType) {
			return (objectPathType===1 ||
				objectPathType===5 ||
				objectPathType===3 ||
				objectPathType===4);
		};
		ObjectPath.copyObjectPathInfo=function (src, dest) {
			dest.Id=src.Id;
			dest.ArgumentInfo=src.ArgumentInfo;
			dest.Name=src.Name;
			dest.ObjectPathType=src.ObjectPathType;
			dest.ParentObjectPathId=src.ParentObjectPathId;
		};
		return ObjectPath;
	}());
	OfficeExtension.ObjectPath=ObjectPath;
})(OfficeExtension || (OfficeExtension={}));
var OfficeExtension;
(function (OfficeExtension) {
	var ObjectPathFactory=(function () {
		function ObjectPathFactory() {
		}
		ObjectPathFactory.createGlobalObjectObjectPath=function (context) {
			var objectPathInfo={ Id: context._nextId(), ObjectPathType: 1, Name: "" };
			return new OfficeExtension.ObjectPath(objectPathInfo, null, false, false);
		};
		ObjectPathFactory.createNewObjectObjectPath=function (context, typeName, isCollection, isRestrictedResourceAccess) {
			var objectPathInfo={ Id: context._nextId(), ObjectPathType: 2, Name: typeName };
			var ret=new OfficeExtension.ObjectPath(objectPathInfo, null, isCollection, false);
			ret.isRestrictedResourceAccess=isRestrictedResourceAccess;
			return ret;
		};
		ObjectPathFactory.createPropertyObjectPath=function (context, parent, propertyName, isCollection, isInvalidAfterRequest, isRestrictedResourceAccess) {
			var objectPathInfo={
				Id: context._nextId(),
				ObjectPathType: 4,
				Name: propertyName,
				ParentObjectPathId: parent._objectPath.objectPathInfo.Id,
			};
			var ret=new OfficeExtension.ObjectPath(objectPathInfo, parent._objectPath, isCollection, isInvalidAfterRequest);
			ret.isRestrictedResourceAccess=isRestrictedResourceAccess;
			return ret;
		};
		ObjectPathFactory.createIndexerObjectPath=function (context, parent, args) {
			var objectPathInfo={
				Id: context._nextId(),
				ObjectPathType: 5,
				Name: "",
				ParentObjectPathId: parent._objectPath.objectPathInfo.Id,
				ArgumentInfo: {}
			};
			objectPathInfo.ArgumentInfo.Arguments=args;
			return new OfficeExtension.ObjectPath(objectPathInfo, parent._objectPath, false, false);
		};
		ObjectPathFactory.createIndexerObjectPathUsingParentPath=function (context, parentObjectPath, args) {
			var objectPathInfo={
				Id: context._nextId(),
				ObjectPathType: 5,
				Name: "",
				ParentObjectPathId: parentObjectPath.objectPathInfo.Id,
				ArgumentInfo: {}
			};
			objectPathInfo.ArgumentInfo.Arguments=args;
			return new OfficeExtension.ObjectPath(objectPathInfo, parentObjectPath, false, false);
		};
		ObjectPathFactory.createMethodObjectPath=function (context, parent, methodName, operationType, args, isCollection, isInvalidAfterRequest, getByIdMethodName, isRestrictedResourceAccess) {
			var objectPathInfo={
				Id: context._nextId(),
				ObjectPathType: 3,
				Name: methodName,
				ParentObjectPathId: parent._objectPath.objectPathInfo.Id,
				ArgumentInfo: {}
			};
			var argumentObjectPaths=OfficeExtension.Utility.setMethodArguments(context, objectPathInfo.ArgumentInfo, args);
			var ret=new OfficeExtension.ObjectPath(objectPathInfo, parent._objectPath, isCollection, isInvalidAfterRequest);
			ret.argumentObjectPaths=argumentObjectPaths;
			ret.isWriteOperation=(operationType !=1);
			ret.getByIdMethodName=getByIdMethodName;
			ret.isRestrictedResourceAccess=isRestrictedResourceAccess;
			return ret;
		};
		ObjectPathFactory.createReferenceIdObjectPath=function (context, referenceId) {
			var objectPathInfo={
				Id: context._nextId(),
				ObjectPathType: 6,
				Name: referenceId,
				ArgumentInfo: {}
			};
			var ret=new OfficeExtension.ObjectPath(objectPathInfo, null, false, false);
			return ret;
		};
		ObjectPathFactory.createChildItemObjectPathUsingIndexerOrGetItemAt=function (hasIndexerMethod, context, parent, childItem, index) {
			var id=OfficeExtension.Utility.tryGetObjectIdFromLoadOrRetrieveResult(childItem);
			if (hasIndexerMethod && !OfficeExtension.Utility.isNullOrUndefined(id)) {
				return ObjectPathFactory.createChildItemObjectPathUsingIndexer(context, parent, childItem);
			}
			else {
				return ObjectPathFactory.createChildItemObjectPathUsingGetItemAt(context, parent, childItem, index);
			}
		};
		ObjectPathFactory.createChildItemObjectPathUsingIndexer=function (context, parent, childItem) {
			var id=OfficeExtension.Utility.tryGetObjectIdFromLoadOrRetrieveResult(childItem);
			var objectPathInfo=objectPathInfo=				{
					Id: context._nextId(),
					ObjectPathType: 5,
					Name: "",
					ParentObjectPathId: parent._objectPath.objectPathInfo.Id,
					ArgumentInfo: {}
				};
			objectPathInfo.ArgumentInfo.Arguments=[id];
			return new OfficeExtension.ObjectPath(objectPathInfo, parent._objectPath, false, false);
		};
		ObjectPathFactory.createChildItemObjectPathUsingGetItemAt=function (context, parent, childItem, index) {
			var indexFromServer=childItem[OfficeExtension.Constants.index];
			if (indexFromServer) {
				index=indexFromServer;
			}
			var objectPathInfo={
				Id: context._nextId(),
				ObjectPathType: 3,
				Name: OfficeExtension.Constants.getItemAt,
				ParentObjectPathId: parent._objectPath.objectPathInfo.Id,
				ArgumentInfo: {}
			};
			objectPathInfo.ArgumentInfo.Arguments=[index];
			return new OfficeExtension.ObjectPath(objectPathInfo, parent._objectPath, false, false);
		};
		return ObjectPathFactory;
	}());
	OfficeExtension.ObjectPathFactory=ObjectPathFactory;
})(OfficeExtension || (OfficeExtension={}));
var OfficeExtension;
(function (OfficeExtension) {
	var OfficeJsRequestExecutor=(function () {
		function OfficeJsRequestExecutor(context) {
			this.m_context=context;
		}
		OfficeJsRequestExecutor.prototype.executeAsync=function (customData, requestFlags, requestMessage) {
			var _this=this;
			var messageSafearray=OfficeExtension.RichApiMessageUtility.buildMessageArrayForIRequestExecutor(customData, requestFlags, requestMessage, OfficeJsRequestExecutor.SourceLibHeaderValue);
			return new OfficeExtension._Internal.OfficePromise(function (resolve, reject) {
				OSF.DDA.RichApi.executeRichApiRequestAsync(messageSafearray, function (result) {
					OfficeExtension.Utility.log("Response:");
					OfficeExtension.Utility.log(JSON.stringify(result));
					var response;
					if (result.status=="succeeded") {
						response=OfficeExtension.RichApiMessageUtility.buildResponseOnSuccess(OfficeExtension.RichApiMessageUtility.getResponseBody(result), OfficeExtension.RichApiMessageUtility.getResponseHeaders(result));
					}
					else {
						response=OfficeExtension.RichApiMessageUtility.buildResponseOnError(result.error.code, result.error.message);
						_this.m_context._processOfficeJsErrorResponse(result.error.code, response);
					}
					resolve(response);
				});
			});
		};
		OfficeJsRequestExecutor.SourceLibHeaderValue="officejs";
		return OfficeJsRequestExecutor;
	}());
	OfficeExtension.OfficeJsRequestExecutor=OfficeJsRequestExecutor;
})(OfficeExtension || (OfficeExtension={}));
var OfficeExtension;
(function (OfficeExtension) {
	var _Internal;
	(function (_Internal) {
		_Internal.OfficeRequire=function () {
			return null;
		}();
	})(_Internal=OfficeExtension._Internal || (OfficeExtension._Internal={}));
	var _Internal;
	(function (_Internal) {
		var PromiseImpl;
		(function (PromiseImpl) {
			function Init() {
				return (function () {
					"use strict";
					function lib$es6$promise$utils$$objectOrFunction(x) {
						return typeof x==='function' || (typeof x==='object' && x !==null);
					}
					function lib$es6$promise$utils$$isFunction(x) {
						return typeof x==='function';
					}
					function lib$es6$promise$utils$$isMaybeThenable(x) {
						return typeof x==='object' && x !==null;
					}
					var lib$es6$promise$utils$$_isArray;
					if (!Array.isArray) {
						lib$es6$promise$utils$$_isArray=function (x) {
							return Object.prototype.toString.call(x)==='[object Array]';
						};
					}
					else {
						lib$es6$promise$utils$$_isArray=Array.isArray;
					}
					var lib$es6$promise$utils$$isArray=lib$es6$promise$utils$$_isArray;
					var lib$es6$promise$asap$$len=0;
					var lib$es6$promise$asap$$toString={}.toString;
					var lib$es6$promise$asap$$vertxNext;
					var lib$es6$promise$asap$$customSchedulerFn;
					var lib$es6$promise$asap$$asap=function asap(callback, arg) {
						lib$es6$promise$asap$$queue[lib$es6$promise$asap$$len]=callback;
						lib$es6$promise$asap$$queue[lib$es6$promise$asap$$len+1]=arg;
						lib$es6$promise$asap$$len+=2;
						if (lib$es6$promise$asap$$len===2) {
							if (lib$es6$promise$asap$$customSchedulerFn) {
								lib$es6$promise$asap$$customSchedulerFn(lib$es6$promise$asap$$flush);
							}
							else {
								lib$es6$promise$asap$$scheduleFlush();
							}
						}
					};
					function lib$es6$promise$asap$$setScheduler(scheduleFn) {
						lib$es6$promise$asap$$customSchedulerFn=scheduleFn;
					}
					function lib$es6$promise$asap$$setAsap(asapFn) {
						lib$es6$promise$asap$$asap=asapFn;
					}
					var lib$es6$promise$asap$$browserWindow=(typeof window !=='undefined') ? window : undefined;
					var lib$es6$promise$asap$$browserGlobal=lib$es6$promise$asap$$browserWindow || {};
					var lib$es6$promise$asap$$BrowserMutationObserver=lib$es6$promise$asap$$browserGlobal.MutationObserver || lib$es6$promise$asap$$browserGlobal.WebKitMutationObserver;
					var lib$es6$promise$asap$$isNode=typeof process !=='undefined' && {}.toString.call(process)==='[object process]';
					var lib$es6$promise$asap$$isWorker=typeof Uint8ClampedArray !=='undefined' &&
						typeof importScripts !=='undefined' &&
						typeof MessageChannel !=='undefined';
					function lib$es6$promise$asap$$useNextTick() {
						var nextTick=process.nextTick;
						var version=process.versions.node.match(/^(?:(\d+)\.)?(?:(\d+)\.)?(\*|\d+)$/);
						if (Array.isArray(version) && version[1]==='0' && version[2]==='10') {
							nextTick=setImmediate;
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
						var iterations=0;
						var observer=new lib$es6$promise$asap$$BrowserMutationObserver(lib$es6$promise$asap$$flush);
						var node=document.createTextNode('');
						observer.observe(node, { characterData: true });
						return function () {
							node.data=(iterations=++iterations % 2);
						};
					}
					function lib$es6$promise$asap$$useMessageChannel() {
						var channel=new MessageChannel();
						channel.port1.onmessage=lib$es6$promise$asap$$flush;
						return function () {
							channel.port2.postMessage(0);
						};
					}
					function lib$es6$promise$asap$$useSetTimeout() {
						return function () {
							setTimeout(lib$es6$promise$asap$$flush, 1);
						};
					}
					var lib$es6$promise$asap$$queue=new Array(1000);
					function lib$es6$promise$asap$$flush() {
						for (var i=0; i < lib$es6$promise$asap$$len; i+=2) {
							var callback=lib$es6$promise$asap$$queue[i];
							var arg=lib$es6$promise$asap$$queue[i+1];
							callback(arg);
							lib$es6$promise$asap$$queue[i]=undefined;
							lib$es6$promise$asap$$queue[i+1]=undefined;
						}
						lib$es6$promise$asap$$len=0;
					}
					var lib$es6$promise$asap$$scheduleFlush;
					if (lib$es6$promise$asap$$isNode) {
						lib$es6$promise$asap$$scheduleFlush=lib$es6$promise$asap$$useNextTick();
					}
					else if (lib$es6$promise$asap$$BrowserMutationObserver) {
						lib$es6$promise$asap$$scheduleFlush=lib$es6$promise$asap$$useMutationObserver();
					}
					else if (lib$es6$promise$asap$$isWorker) {
						lib$es6$promise$asap$$scheduleFlush=lib$es6$promise$asap$$useMessageChannel();
					}
					else {
						lib$es6$promise$asap$$scheduleFlush=lib$es6$promise$asap$$useSetTimeout();
					}
					function lib$es6$promise$$internal$$noop() { }
					var lib$es6$promise$$internal$$PENDING=void 0;
					var lib$es6$promise$$internal$$FULFILLED=1;
					var lib$es6$promise$$internal$$REJECTED=2;
					var lib$es6$promise$$internal$$GET_THEN_ERROR=new lib$es6$promise$$internal$$ErrorObject();
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
							lib$es6$promise$$internal$$GET_THEN_ERROR.error=error;
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
							var sealed=false;
							var error=lib$es6$promise$$internal$$tryThen(then, thenable, function (value) {
								if (sealed) {
									return;
								}
								sealed=true;
								if (thenable !==value) {
									lib$es6$promise$$internal$$resolve(promise, value);
								}
								else {
									lib$es6$promise$$internal$$fulfill(promise, value);
								}
							}, function (reason) {
								if (sealed) {
									return;
								}
								sealed=true;
								lib$es6$promise$$internal$$reject(promise, reason);
							}, 'Settle: '+(promise._label || ' unknown promise'));
							if (!sealed && error) {
								sealed=true;
								lib$es6$promise$$internal$$reject(promise, error);
							}
						}, promise);
					}
					function lib$es6$promise$$internal$$handleOwnThenable(promise, thenable) {
						if (thenable._state===lib$es6$promise$$internal$$FULFILLED) {
							lib$es6$promise$$internal$$fulfill(promise, thenable._result);
						}
						else if (thenable._state===lib$es6$promise$$internal$$REJECTED) {
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
						if (maybeThenable.constructor===promise.constructor) {
							lib$es6$promise$$internal$$handleOwnThenable(promise, maybeThenable);
						}
						else {
							var then=lib$es6$promise$$internal$$getThen(maybeThenable);
							if (then===lib$es6$promise$$internal$$GET_THEN_ERROR) {
								lib$es6$promise$$internal$$reject(promise, lib$es6$promise$$internal$$GET_THEN_ERROR.error);
							}
							else if (then===undefined) {
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
						if (promise===value) {
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
						if (promise._state !==lib$es6$promise$$internal$$PENDING) {
							return;
						}
						promise._result=value;
						promise._state=lib$es6$promise$$internal$$FULFILLED;
						if (promise._subscribers.length !==0) {
							lib$es6$promise$asap$$asap(lib$es6$promise$$internal$$publish, promise);
						}
					}
					function lib$es6$promise$$internal$$reject(promise, reason) {
						if (promise._state !==lib$es6$promise$$internal$$PENDING) {
							return;
						}
						promise._state=lib$es6$promise$$internal$$REJECTED;
						promise._result=reason;
						lib$es6$promise$asap$$asap(lib$es6$promise$$internal$$publishRejection, promise);
					}
					function lib$es6$promise$$internal$$subscribe(parent, child, onFulfillment, onRejection) {
						var subscribers=parent._subscribers;
						var length=subscribers.length;
						parent._onerror=null;
						subscribers[length]=child;
						subscribers[length+lib$es6$promise$$internal$$FULFILLED]=onFulfillment;
						subscribers[length+lib$es6$promise$$internal$$REJECTED]=onRejection;
						if (length===0 && parent._state) {
							lib$es6$promise$asap$$asap(lib$es6$promise$$internal$$publish, parent);
						}
					}
					function lib$es6$promise$$internal$$publish(promise) {
						var subscribers=promise._subscribers;
						var settled=promise._state;
						if (subscribers.length===0) {
							return;
						}
						var child, callback, detail=promise._result;
						for (var i=0; i < subscribers.length; i+=3) {
							child=subscribers[i];
							callback=subscribers[i+settled];
							if (child) {
								lib$es6$promise$$internal$$invokeCallback(settled, child, callback, detail);
							}
							else {
								callback(detail);
							}
						}
						promise._subscribers.length=0;
					}
					function lib$es6$promise$$internal$$ErrorObject() {
						this.error=null;
					}
					var lib$es6$promise$$internal$$TRY_CATCH_ERROR=new lib$es6$promise$$internal$$ErrorObject();
					function lib$es6$promise$$internal$$tryCatch(callback, detail) {
						try {
							return callback(detail);
						}
						catch (e) {
							lib$es6$promise$$internal$$TRY_CATCH_ERROR.error=e;
							return lib$es6$promise$$internal$$TRY_CATCH_ERROR;
						}
					}
					function lib$es6$promise$$internal$$invokeCallback(settled, promise, callback, detail) {
						var hasCallback=lib$es6$promise$utils$$isFunction(callback), value, error, succeeded, failed;
						if (hasCallback) {
							value=lib$es6$promise$$internal$$tryCatch(callback, detail);
							if (value===lib$es6$promise$$internal$$TRY_CATCH_ERROR) {
								failed=true;
								error=value.error;
								value=null;
							}
							else {
								succeeded=true;
							}
							if (promise===value) {
								lib$es6$promise$$internal$$reject(promise, lib$es6$promise$$internal$$cannotReturnOwn());
								return;
							}
						}
						else {
							value=detail;
							succeeded=true;
						}
						if (promise._state !==lib$es6$promise$$internal$$PENDING) {
						}
						else if (hasCallback && succeeded) {
							lib$es6$promise$$internal$$resolve(promise, value);
						}
						else if (failed) {
							lib$es6$promise$$internal$$reject(promise, error);
						}
						else if (settled===lib$es6$promise$$internal$$FULFILLED) {
							lib$es6$promise$$internal$$fulfill(promise, value);
						}
						else if (settled===lib$es6$promise$$internal$$REJECTED) {
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
						var enumerator=this;
						enumerator._instanceConstructor=Constructor;
						enumerator.promise=new Constructor(lib$es6$promise$$internal$$noop);
						if (enumerator._validateInput(input)) {
							enumerator._input=input;
							enumerator.length=input.length;
							enumerator._remaining=input.length;
							enumerator._init();
							if (enumerator.length===0) {
								lib$es6$promise$$internal$$fulfill(enumerator.promise, enumerator._result);
							}
							else {
								enumerator.length=enumerator.length || 0;
								enumerator._enumerate();
								if (enumerator._remaining===0) {
									lib$es6$promise$$internal$$fulfill(enumerator.promise, enumerator._result);
								}
							}
						}
						else {
							lib$es6$promise$$internal$$reject(enumerator.promise, enumerator._validationError());
						}
					}
					lib$es6$promise$enumerator$$Enumerator.prototype._validateInput=function (input) {
						return lib$es6$promise$utils$$isArray(input);
					};
					lib$es6$promise$enumerator$$Enumerator.prototype._validationError=function () {
						return new _Internal.Error('Array Methods must be provided an Array');
					};
					lib$es6$promise$enumerator$$Enumerator.prototype._init=function () {
						this._result=new Array(this.length);
					};
					var lib$es6$promise$enumerator$$default=lib$es6$promise$enumerator$$Enumerator;
					lib$es6$promise$enumerator$$Enumerator.prototype._enumerate=function () {
						var enumerator=this;
						var length=enumerator.length;
						var promise=enumerator.promise;
						var input=enumerator._input;
						for (var i=0; promise._state===lib$es6$promise$$internal$$PENDING && i < length; i++) {
							enumerator._eachEntry(input[i], i);
						}
					};
					lib$es6$promise$enumerator$$Enumerator.prototype._eachEntry=function (entry, i) {
						var enumerator=this;
						var c=enumerator._instanceConstructor;
						if (lib$es6$promise$utils$$isMaybeThenable(entry)) {
							if (entry.constructor===c && entry._state !==lib$es6$promise$$internal$$PENDING) {
								entry._onerror=null;
								enumerator._settledAt(entry._state, i, entry._result);
							}
							else {
								enumerator._willSettleAt(c.resolve(entry), i);
							}
						}
						else {
							enumerator._remaining--;
							enumerator._result[i]=entry;
						}
					};
					lib$es6$promise$enumerator$$Enumerator.prototype._settledAt=function (state, i, value) {
						var enumerator=this;
						var promise=enumerator.promise;
						if (promise._state===lib$es6$promise$$internal$$PENDING) {
							enumerator._remaining--;
							if (state===lib$es6$promise$$internal$$REJECTED) {
								lib$es6$promise$$internal$$reject(promise, value);
							}
							else {
								enumerator._result[i]=value;
							}
						}
						if (enumerator._remaining===0) {
							lib$es6$promise$$internal$$fulfill(promise, enumerator._result);
						}
					};
					lib$es6$promise$enumerator$$Enumerator.prototype._willSettleAt=function (promise, i) {
						var enumerator=this;
						lib$es6$promise$$internal$$subscribe(promise, undefined, function (value) {
							enumerator._settledAt(lib$es6$promise$$internal$$FULFILLED, i, value);
						}, function (reason) {
							enumerator._settledAt(lib$es6$promise$$internal$$REJECTED, i, reason);
						});
					};
					function lib$es6$promise$promise$all$$all(entries) {
						return new lib$es6$promise$enumerator$$default(this, entries).promise;
					}
					var lib$es6$promise$promise$all$$default=lib$es6$promise$promise$all$$all;
					function lib$es6$promise$promise$race$$race(entries) {
						var Constructor=this;
						var promise=new Constructor(lib$es6$promise$$internal$$noop);
						if (!lib$es6$promise$utils$$isArray(entries)) {
							lib$es6$promise$$internal$$reject(promise, new TypeError('You must pass an array to race.'));
							return promise;
						}
						var length=entries.length;
						function onFulfillment(value) {
							lib$es6$promise$$internal$$resolve(promise, value);
						}
						function onRejection(reason) {
							lib$es6$promise$$internal$$reject(promise, reason);
						}
						for (var i=0; promise._state===lib$es6$promise$$internal$$PENDING && i < length; i++) {
							lib$es6$promise$$internal$$subscribe(Constructor.resolve(entries[i]), undefined, onFulfillment, onRejection);
						}
						return promise;
					}
					var lib$es6$promise$promise$race$$default=lib$es6$promise$promise$race$$race;
					function lib$es6$promise$promise$resolve$$resolve(object) {
						var Constructor=this;
						if (object && typeof object==='object' && object.constructor===Constructor) {
							return object;
						}
						var promise=new Constructor(lib$es6$promise$$internal$$noop);
						lib$es6$promise$$internal$$resolve(promise, object);
						return promise;
					}
					var lib$es6$promise$promise$resolve$$default=lib$es6$promise$promise$resolve$$resolve;
					function lib$es6$promise$promise$reject$$reject(reason) {
						var Constructor=this;
						var promise=new Constructor(lib$es6$promise$$internal$$noop);
						lib$es6$promise$$internal$$reject(promise, reason);
						return promise;
					}
					var lib$es6$promise$promise$reject$$default=lib$es6$promise$promise$reject$$reject;
					var lib$es6$promise$promise$$counter=0;
					function lib$es6$promise$promise$$needsResolver() {
						throw new TypeError('You must pass a resolver function as the first argument to the promise constructor');
					}
					function lib$es6$promise$promise$$needsNew() {
						throw new TypeError("Failed to construct 'Promise': Please use the 'new' operator, this object constructor cannot be called as a function.");
					}
					var lib$es6$promise$promise$$default=lib$es6$promise$promise$$Promise;
					function lib$es6$promise$promise$$Promise(resolver) {
						this._id=lib$es6$promise$promise$$counter++;
						this._state=undefined;
						this._result=undefined;
						this._subscribers=[];
						if (lib$es6$promise$$internal$$noop !==resolver) {
							if (!lib$es6$promise$utils$$isFunction(resolver)) {
								lib$es6$promise$promise$$needsResolver();
							}
							if (!(this instanceof lib$es6$promise$promise$$Promise)) {
								lib$es6$promise$promise$$needsNew();
							}
							lib$es6$promise$$internal$$initializePromise(this, resolver);
						}
					}
					lib$es6$promise$promise$$Promise.all=lib$es6$promise$promise$all$$default;
					lib$es6$promise$promise$$Promise.race=lib$es6$promise$promise$race$$default;
					lib$es6$promise$promise$$Promise.resolve=lib$es6$promise$promise$resolve$$default;
					lib$es6$promise$promise$$Promise.reject=lib$es6$promise$promise$reject$$default;
					lib$es6$promise$promise$$Promise._setScheduler=lib$es6$promise$asap$$setScheduler;
					lib$es6$promise$promise$$Promise._setAsap=lib$es6$promise$asap$$setAsap;
					lib$es6$promise$promise$$Promise._asap=lib$es6$promise$asap$$asap;
					lib$es6$promise$promise$$Promise.prototype={
						constructor: lib$es6$promise$promise$$Promise,
						then: function (onFulfillment, onRejection) {
							var parent=this;
							var state=parent._state;
							if (state===lib$es6$promise$$internal$$FULFILLED && !onFulfillment || state===lib$es6$promise$$internal$$REJECTED && !onRejection) {
								return this;
							}
							var child=new this.constructor(lib$es6$promise$$internal$$noop);
							var result=parent._result;
							if (state) {
								var callback=arguments[state - 1];
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
			PromiseImpl.Init=Init;
		})(PromiseImpl=_Internal.PromiseImpl || (_Internal.PromiseImpl={}));
	})(_Internal=OfficeExtension._Internal || (OfficeExtension._Internal={}));
	var _Internal;
	(function (_Internal) {
		function isEdgeLessThan14() {
			var userAgent=window.navigator.userAgent;
			var versionIdx=userAgent.indexOf("Edge/");
			if (versionIdx >=0) {
				userAgent=userAgent.substring(versionIdx+5, userAgent.length);
				if (userAgent < "14.14393")
					return true;
				else
					return false;
			}
			return false;
		}
		function determinePromise() {
			if (typeof (window)==="undefined" && typeof (Promise)==="function") {
				return Promise;
			}
			if (typeof (window) !=="undefined" && window.Promise) {
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
		_Internal.OfficePromise=determinePromise();
	})(_Internal=OfficeExtension._Internal || (OfficeExtension._Internal={}));
	var OfficePromise=_Internal.OfficePromise;
	OfficeExtension.Promise=OfficePromise;
})(OfficeExtension || (OfficeExtension={}));
var OfficeExtension;
(function (OfficeExtension) {
	var TrackedObjects=(function () {
		function TrackedObjects(context) {
			this._autoCleanupList={};
			this.m_context=context;
		}
		TrackedObjects.prototype.add=function (param) {
			var _this=this;
			if (Array.isArray(param)) {
				param.forEach(function (item) { return _this._addCommon(item, true); });
			}
			else {
				this._addCommon(param, true);
			}
		};
		TrackedObjects.prototype._autoAdd=function (object) {
			this._addCommon(object, false);
			this._autoCleanupList[object._objectPath.objectPathInfo.Id]=object;
		};
		TrackedObjects.prototype._autoTrackIfNecessaryWhenHandleObjectResultValue=function (object, resultValue) {
			var shouldAutoTrack=(this.m_context._autoCleanup &&
				!object[OfficeExtension.Constants.isTracked] &&
				object !==this.m_context._rootObject &&
				resultValue &&
				!OfficeExtension.Utility.isNullOrEmptyString(resultValue[OfficeExtension.Constants.referenceId]));
			if (shouldAutoTrack) {
				this._autoCleanupList[object._objectPath.objectPathInfo.Id]=object;
				object[OfficeExtension.Constants.isTracked]=true;
			}
		};
		TrackedObjects.prototype._addCommon=function (object, isExplicitlyAdded) {
			if (object[OfficeExtension.Constants.isTracked]) {
				if (isExplicitlyAdded && this.m_context._autoCleanup) {
					delete this._autoCleanupList[object._objectPath.objectPathInfo.Id];
				}
				return;
			}
			var referenceId=object[OfficeExtension.Constants.referenceId];
			if (OfficeExtension.Utility.isNullOrEmptyString(referenceId) && object._KeepReference) {
				object._KeepReference();
				OfficeExtension.ActionFactory.createInstantiateAction(this.m_context, object);
				if (isExplicitlyAdded && this.m_context._autoCleanup) {
					delete this._autoCleanupList[object._objectPath.objectPathInfo.Id];
				}
				object[OfficeExtension.Constants.isTracked]=true;
			}
		};
		TrackedObjects.prototype.remove=function (param) {
			var _this=this;
			if (Array.isArray(param)) {
				param.forEach(function (item) { return _this._removeCommon(item); });
			}
			else {
				this._removeCommon(param);
			}
		};
		TrackedObjects.prototype._removeCommon=function (object) {
			var referenceId=object[OfficeExtension.Constants.referenceId];
			if (!OfficeExtension.Utility.isNullOrEmptyString(referenceId)) {
				var rootObject=this.m_context._rootObject;
				if (rootObject._RemoveReference) {
					rootObject._RemoveReference(referenceId);
				}
				delete object[OfficeExtension.Constants.isTracked];
			}
		};
		TrackedObjects.prototype._retrieveAndClearAutoCleanupList=function () {
			var list=this._autoCleanupList;
			this._autoCleanupList={};
			return list;
		};
		return TrackedObjects;
	}());
	OfficeExtension.TrackedObjects=TrackedObjects;
})(OfficeExtension || (OfficeExtension={}));
var OfficeExtension;
(function (OfficeExtension) {
	var RequestPrettyPrinter=(function () {
		function RequestPrettyPrinter(globalObjName, referencedObjectPaths, actions, showDispose) {
			if (!globalObjName) {
				globalObjName="root";
			}
			this.m_globalObjName=globalObjName;
			this.m_referencedObjectPaths=referencedObjectPaths;
			this.m_actions=actions;
			this.m_statements=[];
			this.m_variableNameForObjectPathMap={};
			this.m_variableNameToObjectPathMap={};
			this.m_declaredObjectPathMap={};
			this.m_showDispose=showDispose;
		}
		RequestPrettyPrinter.prototype.process=function () {
			if (this.m_showDispose) {
				OfficeExtension.ClientRequest._calculateLastUsedObjectPathIds(this.m_actions);
			}
			for (var i=0; i < this.m_actions.length; i++) {
				this.processOneAction(this.m_actions[i]);
			}
			return this.m_statements;
		};
		RequestPrettyPrinter.prototype.processForDebugStatementInfo=function (actionIndex) {
			if (this.m_showDispose) {
				OfficeExtension.ClientRequest._calculateLastUsedObjectPathIds(this.m_actions);
			}
			var surroundingCount=5;
			this.m_statements=[];
			var oneStatement="";
			var statementIndex=-1;
			for (var i=0; i < this.m_actions.length; i++) {
				this.processOneAction(this.m_actions[i]);
				if (actionIndex==i) {
					statementIndex=this.m_statements.length - 1;
				}
				if (statementIndex >=0 && this.m_statements.length > statementIndex+surroundingCount+1) {
					break;
				}
			}
			if (statementIndex < 0) {
				return null;
			}
			var startIndex=statementIndex - surroundingCount;
			if (startIndex < 0) {
				startIndex=0;
			}
			var endIndex=statementIndex+1+surroundingCount;
			if (endIndex > this.m_statements.length) {
				endIndex=this.m_statements.length;
			}
			var surroundingStatements=[];
			if (startIndex !=0) {
				surroundingStatements.push("...");
			}
			for (var i_1=startIndex; i_1 < statementIndex; i_1++) {
				surroundingStatements.push(this.m_statements[i_1]);
			}
			surroundingStatements.push("// >>>>>");
			surroundingStatements.push(this.m_statements[statementIndex]);
			surroundingStatements.push("// <<<<<");
			for (var i_2=statementIndex+1; i_2 < endIndex; i_2++) {
				surroundingStatements.push(this.m_statements[i_2]);
			}
			if (endIndex < this.m_statements.length) {
				surroundingStatements.push("...");
			}
			return {
				statement: this.m_statements[statementIndex],
				surroundingStatements: surroundingStatements
			};
		};
		RequestPrettyPrinter.prototype.processOneAction=function (action) {
			var actionInfo=action.actionInfo;
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
		RequestPrettyPrinter.prototype.processInstantiateAction=function (action) {
			var objId=action.actionInfo.ObjectPathId;
			var objPath=this.m_referencedObjectPaths[objId];
			var varName=this.getObjVarName(objId);
			if (!this.m_declaredObjectPathMap[objId]) {
				var statement="var "+varName+"="+this.buildObjectPathExpressionWithParent(objPath)+";";
				statement=this.appendDisposeCommentIfRelevant(statement, action);
				this.m_statements.push(statement);
				this.m_declaredObjectPathMap[objId]=varName;
			}
			else {
				var statement="// Instantiate {"+varName+"}";
				statement=this.appendDisposeCommentIfRelevant(statement, action);
				this.m_statements.push(statement);
			}
		};
		RequestPrettyPrinter.prototype.processMethodAction=function (action) {
			var methodName=action.actionInfo.Name;
			if (methodName==="_KeepReference") {
				methodName="track";
			}
			var statement=this.getObjVarName(action.actionInfo.ObjectPathId)+"."+OfficeExtension.Utility._toCamelLowerCase(methodName)+"("+this.buildArgumentsExpression(action.actionInfo.ArgumentInfo)+");";
			statement=this.appendDisposeCommentIfRelevant(statement, action);
			this.m_statements.push(statement);
		};
		RequestPrettyPrinter.prototype.processQueryAction=function (action) {
			var queryExp=this.buildQueryExpression(action);
			var statement=this.getObjVarName(action.actionInfo.ObjectPathId)+".load("+queryExp+");";
			statement=this.appendDisposeCommentIfRelevant(statement, action);
			this.m_statements.push(statement);
		};
		RequestPrettyPrinter.prototype.processQueryAsJsonAction=function (action) {
			var queryExp=this.buildQueryExpression(action);
			var statement=this.getObjVarName(action.actionInfo.ObjectPathId)+".retrieve("+queryExp+");";
			statement=this.appendDisposeCommentIfRelevant(statement, action);
			this.m_statements.push(statement);
		};
		RequestPrettyPrinter.prototype.processRecursiveQueryAction=function (action) {
			var queryExp="";
			if (action.actionInfo.RecursiveQueryInfo) {
				queryExp=JSON.stringify(action.actionInfo.RecursiveQueryInfo);
			}
			var statement=this.getObjVarName(action.actionInfo.ObjectPathId)+".loadRecursive("+queryExp+");";
			statement=this.appendDisposeCommentIfRelevant(statement, action);
			this.m_statements.push(statement);
		};
		RequestPrettyPrinter.prototype.processSetPropertyAction=function (action) {
			var statement=this.getObjVarName(action.actionInfo.ObjectPathId)+"."+OfficeExtension.Utility._toCamelLowerCase(action.actionInfo.Name)+"="+this.buildArgumentsExpression(action.actionInfo.ArgumentInfo)+";";
			statement=this.appendDisposeCommentIfRelevant(statement, action);
			this.m_statements.push(statement);
		};
		RequestPrettyPrinter.prototype.processTraceAction=function (action) {
			var statement="context.trace();";
			statement=this.appendDisposeCommentIfRelevant(statement, action);
			this.m_statements.push(statement);
		};
		RequestPrettyPrinter.prototype.processEnsureUnchangedAction=function (action) {
			var statement=this.getObjVarName(action.actionInfo.ObjectPathId)+".ensureUnchanged("+JSON.stringify(action.actionInfo.ObjectState)+");";
			statement=this.appendDisposeCommentIfRelevant(statement, action);
			this.m_statements.push(statement);
		};
		RequestPrettyPrinter.prototype.processUpdateAction=function (action) {
			var statement=this.getObjVarName(action.actionInfo.ObjectPathId)+".update("+JSON.stringify(action.actionInfo.ObjectState)+");";
			statement=this.appendDisposeCommentIfRelevant(statement, action);
			this.m_statements.push(statement);
		};
		RequestPrettyPrinter.prototype.appendDisposeCommentIfRelevant=function (statement, action) {
			var _this=this;
			if (this.m_showDispose) {
				var lastUsedObjectPathIds=action.actionInfo.L;
				if (lastUsedObjectPathIds && lastUsedObjectPathIds.length > 0) {
					var objectNamesToDispose=lastUsedObjectPathIds.map(function (item) { return _this.getObjVarName(item); }).join(", ");
					return statement+" // And then dispose {"+objectNamesToDispose+"}";
				}
			}
			return statement;
		};
		RequestPrettyPrinter.prototype.buildQueryExpression=function (action) {
			if (action.actionInfo.QueryInfo) {
				var option={};
				option.select=action.actionInfo.QueryInfo.Select;
				option.expand=action.actionInfo.QueryInfo.Expand;
				option.skip=action.actionInfo.QueryInfo.Skip;
				option.top=action.actionInfo.QueryInfo.Top;
				if (typeof (option.top)==="undefined" && typeof (option.skip)==="undefined" && typeof (option.expand)==="undefined") {
					if (typeof (option.select)==="undefined") {
						return "";
					}
					else {
						return JSON.stringify(option.select);
					}
				}
				else {
					return JSON.stringify(option);
				}
			}
			return "";
		};
		RequestPrettyPrinter.prototype.buildObjectPathExpressionWithParent=function (objPath) {
			var hasParent=objPath.objectPathInfo.ObjectPathType==5 ||
				objPath.objectPathInfo.ObjectPathType==3 ||
				objPath.objectPathInfo.ObjectPathType==4;
			if (hasParent && objPath.objectPathInfo.ParentObjectPathId) {
				return this.getObjVarName(objPath.objectPathInfo.ParentObjectPathId)+"."+this.buildObjectPathExpression(objPath);
			}
			return this.buildObjectPathExpression(objPath);
		};
		RequestPrettyPrinter.prototype.buildObjectPathExpression=function (objPath) {
			switch (objPath.objectPathInfo.ObjectPathType) {
				case 1:
					return "context."+this.m_globalObjName;
				case 5:
					return "getItem("+this.buildArgumentsExpression(objPath.objectPathInfo.ArgumentInfo)+")";
				case 3:
					return OfficeExtension.Utility._toCamelLowerCase(objPath.objectPathInfo.Name)+"("+this.buildArgumentsExpression(objPath.objectPathInfo.ArgumentInfo)+")";
				case 2:
					return objPath.objectPathInfo.Name+".newObject()";
				case 7:
					return "null";
				case 4:
					return OfficeExtension.Utility._toCamelLowerCase(objPath.objectPathInfo.Name);
				case 6:
					return "context."+this.m_globalObjName+"._getObjectByReferenceId("+JSON.stringify(objPath.objectPathInfo.Name)+")";
			}
		};
		RequestPrettyPrinter.prototype.buildArgumentsExpression=function (args) {
			var ret="";
			if (!args.Arguments) {
				return ret;
			}
			for (var i=0; i < args.Arguments.length; i++) {
				if (i > 0) {
					ret=ret+", ";
				}
				ret=ret+this.buildArgumentLiteral(args.Arguments[i], args.ReferencedObjectPathIds ? args.ReferencedObjectPathIds[i] : null);
			}
			if (ret==="undefined") {
				ret="";
			}
			return ret;
		};
		RequestPrettyPrinter.prototype.buildArgumentLiteral=function (value, objectPathId) {
			if (typeof value=="number" && value===objectPathId) {
				return this.getObjVarName(objectPathId);
			}
			else {
				return JSON.stringify(value);
			}
		};
		RequestPrettyPrinter.prototype.getObjVarNameBase=function (objectPathId) {
			var ret="v";
			var objPath=this.m_referencedObjectPaths[objectPathId];
			if (objPath) {
				switch (objPath.objectPathInfo.ObjectPathType) {
					case 1:
						ret=this.m_globalObjName;
						break;
					case 4:
						ret=OfficeExtension.Utility._toCamelLowerCase(objPath.objectPathInfo.Name);
						break;
					case 3:
						var methodName=objPath.objectPathInfo.Name;
						if (methodName.length > 3 && methodName.substr(0, 3)==="Get") {
							methodName=methodName.substr(3);
						}
						ret=OfficeExtension.Utility._toCamelLowerCase(methodName);
						break;
					case 5:
						var parentName=this.getObjVarNameBase(objPath.objectPathInfo.ParentObjectPathId);
						if (parentName.charAt(parentName.length - 1)==="s") {
							ret=parentName.substr(0, parentName.length - 1);
						}
						else {
							ret=parentName+"Item";
						}
						break;
				}
			}
			return ret;
		};
		RequestPrettyPrinter.prototype.getObjVarName=function (objectPathId) {
			if (this.m_variableNameForObjectPathMap[objectPathId]) {
				return this.m_variableNameForObjectPathMap[objectPathId];
			}
			var ret=this.getObjVarNameBase(objectPathId);
			if (!this.m_variableNameToObjectPathMap[ret]) {
				this.m_variableNameForObjectPathMap[objectPathId]=ret;
				this.m_variableNameToObjectPathMap[ret]=objectPathId;
				return ret;
			}
			var i=1;
			while (this.m_variableNameToObjectPathMap[ret+i.toString()]) {
				i++;
			}
			ret=ret+i.toString();
			this.m_variableNameForObjectPathMap[objectPathId]=ret;
			this.m_variableNameToObjectPathMap[ret]=objectPathId;
			return ret;
		};
		return RequestPrettyPrinter;
	}());
	OfficeExtension.RequestPrettyPrinter=RequestPrettyPrinter;
})(OfficeExtension || (OfficeExtension={}));
var OfficeExtension;
(function (OfficeExtension) {
	var ResourceStrings=(function () {
		function ResourceStrings() {
		}
		ResourceStrings.cannotRegisterEvent="CannotRegisterEvent";
		ResourceStrings.connectionFailureWithStatus="ConnectionFailureWithStatus";
		ResourceStrings.connectionFailureWithDetails="ConnectionFailureWithDetails";
		ResourceStrings.invalidObjectPath="InvalidObjectPath";
		ResourceStrings.invalidRequestContext="InvalidRequestContext";
		ResourceStrings.invalidArgument="InvalidArgument";
		ResourceStrings.invalidArgumentGeneric="InvalidArgumentGeneric";
		ResourceStrings.propertyNotLoaded="PropertyNotLoaded";
		ResourceStrings.runMustReturnPromise="RunMustReturnPromise";
		ResourceStrings.timeout="Timeout";
		ResourceStrings.propertyDoesNotExist="PropertyDoesNotExist";
		ResourceStrings.attemptingToSetReadOnlyProperty="AttemptingToSetReadOnlyProperty";
		ResourceStrings.moreInfoInnerError="MoreInfoInnerError";
		ResourceStrings.cannotApplyPropertyThroughSetMethod="CannotApplyPropertyThroughSetMethod";
		ResourceStrings.valueNotLoaded="ValueNotLoaded";
		ResourceStrings.invalidOrTimedOutSessionMessage="InvalidOrTimedOutSessionMessage";
		ResourceStrings.invalidOperationInCellEditMode="InvalidOperationInCellEditMode";
		ResourceStrings.customFunctionDefintionMissing="CustomFunctionDefintionMissing";
		ResourceStrings.customFunctionImplementationMissing="CustomFunctionImplementationMissing";
		ResourceStrings.customFunctionNameContainsBadChars="CustomFunctionNameContainsBadChars";
		ResourceStrings.customFunctionNameCannotSplit="CustomFunctionNameCannotSplit";
		ResourceStrings.customFunctionUnexpectedNumberOfEntriesInResultBatch="CustomFunctionUnexpectedNumberOfEntriesInResultBatch";
		ResourceStrings.customFunctionCancellationHandlerMissing="CustomFunctionCancellationHandlerMissing";
		ResourceStrings.apiNotFoundDetails="ApiNotFoundDetails";
		ResourceStrings.pendingBatchInProgress="PendingBatchInProgress";
		ResourceStrings.notInsideBatch="NotInsideBatch";
		ResourceStrings.cannotUpdateReadOnlyProperty="CannotUpdateReadOnlyProperty";
		return ResourceStrings;
	}());
	OfficeExtension.ResourceStrings=ResourceStrings;
})(OfficeExtension || (OfficeExtension={}));
var OfficeExtension;
(function (OfficeExtension) {
	var ResourceStringValues=(function () {
		function ResourceStringValues() {
		}
		ResourceStringValues.CannotRegisterEvent="The event handler cannot be registered.";
		ResourceStringValues.ConnectionFailureWithStatus="The request failed with status code of {0}.";
		ResourceStringValues.ConnectionFailureWithDetails="The request failed with status code of {0}, error code {1} and the following error message: {2}";
		ResourceStringValues.InvalidArgument="The argument '{0}' doesn't work for this situation, is missing, or isn't in the right format.";
		ResourceStringValues.InvalidObjectPath="The object path '{0}' isn't working for what you're trying to do. If you're using the object across multiple \"context.sync\" calls and outside the sequential execution of a \".run\" batch, please use the \"context.trackedObjects.add()\" and \"context.trackedObjects.remove()\" methods to manage the object's lifetime.";
		ResourceStringValues.InvalidRequestContext="Cannot use the object across different request contexts.";
		ResourceStringValues.PropertyNotLoaded="The property '{0}' is not available. Before reading the property's value, call the load method on the containing object and call \"context.sync()\" on the associated request context.";
		ResourceStringValues.RunMustReturnPromise="The batch function passed to the \".run\" method didn't return a promise. The function must return a promise, so that any automatically-tracked objects can be released at the completion of the batch operation. Typically, you return a promise by returning the response from \"context.sync()\".";
		ResourceStringValues.Timeout="The operation has timed out.";
		ResourceStringValues.ValueNotLoaded="The value of the result object has not been loaded yet. Before reading the value property, call \"context.sync()\" on the associated request context.";
		ResourceStringValues.InvalidOrTimedOutSessionMessage="Your Office Online session has expired or is invalid. To continue, refresh the page.";
		ResourceStringValues.InvalidOperationInCellEditMode="Excel is in cell-editing mode. Please exit the edit mode by pressing ENTER or TAB or selecting another cell, and then try again.";
		ResourceStringValues.CustomFunctionDefintionMissing="A property with this name that represents the function's definition must exist on Excel.CustomFunctions.";
		ResourceStringValues.CustomFunctionImplementationMissing="The property with this name on Excel.CustomFunctions that represents the function's definition must contain a 'call' property that implements the function.";
		ResourceStringValues.CustomFunctionNameContainsBadChars="The function name may only contain letters, digits, underscores, and periods.";
		ResourceStringValues.CustomFunctionNameCannotSplit="The function name must contain a non-empty namespace and a non-empty short name.";
		ResourceStringValues.CustomFunctionUnexpectedNumberOfEntriesInResultBatch="The batching function returned a number of results that doesn't match the number of parameter value sets that were passed into it.";
		ResourceStringValues.CustomFunctionCancellationHandlerMissing="The cancellation handler onCanceled is missing in the function. The handler must be present as the function is defined as cancelable.";
		ResourceStringValues.ApiNotFoundDetails="The method or property {0} is part of the {1} requirement set, which is not available in your version of {2}.";
		ResourceStringValues.PendingBatchInProgress="There is a pending batch in progress. The batch method may not be called inside another batch, or simultaneously with another batch.";
		ResourceStringValues.NotInsideBatch="Operations may not be invoked outside of a batch method.";
		ResourceStringValues.CannotUpdateReadOnlyProperty="The property '{0}' is read-only and it cannot be updated.";
		return ResourceStringValues;
	}());
	OfficeExtension.ResourceStringValues=ResourceStringValues;
})(OfficeExtension || (OfficeExtension={}));
var OfficeExtension;
(function (OfficeExtension) {
	var RichApiMessageUtility=(function () {
		function RichApiMessageUtility() {
		}
		RichApiMessageUtility.buildMessageArrayForIRequestExecutor=function (customData, requestFlags, requestMessage, sourceLibHeaderValue) {
			var requestMessageText=JSON.stringify(requestMessage.Body);
			OfficeExtension.Utility.log("Request:");
			OfficeExtension.Utility.log(requestMessageText);
			var headers={};
			headers[OfficeExtension.Constants.sourceLibHeader]=sourceLibHeaderValue;
			var messageSafearray=RichApiMessageUtility.buildRequestMessageSafeArray(customData, requestFlags, "POST", "ProcessQuery", headers, requestMessageText);
			return messageSafearray;
		};
		RichApiMessageUtility.buildResponseOnSuccess=function (responseBody, responseHeaders) {
			var response={ ErrorCode: '', ErrorMessage: '', Headers: null, Body: null };
			response.Body=JSON.parse(responseBody);
			response.Headers=responseHeaders;
			return response;
		};
		RichApiMessageUtility.buildResponseOnError=function (errorCode, message) {
			var response={ ErrorCode: '', ErrorMessage: '', Headers: null, Body: null };
			response.ErrorCode=OfficeExtension.ErrorCodes.generalException;
			response.ErrorMessage=message;
			if (errorCode==RichApiMessageUtility.OfficeJsErrorCode_ooeNoCapability) {
				response.ErrorCode=OfficeExtension.ErrorCodes.accessDenied;
			}
			else if (errorCode==RichApiMessageUtility.OfficeJsErrorCode_ooeActivityLimitReached) {
				response.ErrorCode=OfficeExtension.ErrorCodes.activityLimitReached;
			}
			else if (errorCode==RichApiMessageUtility.OfficeJsErrorCode_ooeInvalidOrTimedOutSession) {
				response.ErrorCode=OfficeExtension.ErrorCodes.invalidOrTimedOutSession;
				response.ErrorMessage=OfficeExtension.Utility._getResourceString(OfficeExtension.ResourceStrings.invalidOrTimedOutSessionMessage);
			}
			return response;
		};
		RichApiMessageUtility.buildHttpResponseFromOfficeJsError=function (errorCode, message) {
			var statusCode=500;
			var errorBody={};
			errorBody["error"]={};
			errorBody["error"]["code"]=OfficeExtension.ErrorCodes.generalException;
			errorBody["error"]["message"]=message;
			if (errorCode===RichApiMessageUtility.OfficeJsErrorCode_ooeNoCapability) {
				statusCode=403;
				errorBody["error"]["code"]=OfficeExtension.ErrorCodes.accessDenied;
			}
			else if (errorCode===RichApiMessageUtility.OfficeJsErrorCode_ooeActivityLimitReached) {
				statusCode=429;
				errorBody["error"]["code"]=OfficeExtension.ErrorCodes.activityLimitReached;
			}
			return { statusCode: statusCode, headers: {}, body: JSON.stringify(errorBody) };
		};
		RichApiMessageUtility.buildRequestMessageSafeArray=function (customData, requestFlags, method, path, headers, body) {
			var headerArray=[];
			if (headers) {
				for (var headerName in headers) {
					headerArray.push(headerName);
					headerArray.push(headers[headerName]);
				}
			}
			var appPermission=0;
			var solutionId="";
			var instanceId="";
			var marketplaceType="";
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
		RichApiMessageUtility.getResponseBody=function (result) {
			return RichApiMessageUtility.getResponseBodyFromSafeArray(result.value.data);
		};
		RichApiMessageUtility.getResponseHeaders=function (result) {
			return RichApiMessageUtility.getResponseHeadersFromSafeArray(result.value.data);
		};
		RichApiMessageUtility.getResponseBodyFromSafeArray=function (data) {
			var ret=data[2];
			if (typeof (ret)==="string") {
				return ret;
			}
			var arr=ret;
			return arr.join("");
		};
		RichApiMessageUtility.getResponseHeadersFromSafeArray=function (data) {
			var arrayHeader=data[1];
			if (!arrayHeader) {
				return null;
			}
			var headers={};
			for (var i=0; i < arrayHeader.length - 1; i+=2) {
				headers[arrayHeader[i]]=arrayHeader[i+1];
			}
			return headers;
		};
		RichApiMessageUtility.getResponseStatusCode=function (result) {
			return RichApiMessageUtility.getResponseStatusCodeFromSafeArray(result.value.data);
		};
		RichApiMessageUtility.getResponseStatusCodeFromSafeArray=function (data) {
			return data[0];
		};
		RichApiMessageUtility.OfficeJsErrorCode_ooeInvalidOrTimedOutSession=5012;
		RichApiMessageUtility.OfficeJsErrorCode_ooeActivityLimitReached=5102;
		RichApiMessageUtility.OfficeJsErrorCode_ooeNoCapability=7000;
		return RichApiMessageUtility;
	}());
	OfficeExtension.RichApiMessageUtility=RichApiMessageUtility;
})(OfficeExtension || (OfficeExtension={}));
var OfficeExtension;
(function (OfficeExtension) {
	var Utility=(function () {
		function Utility() {
		}
		Utility.checkArgumentNull=function (value, name) {
			if (Utility.isNullOrUndefined(value)) {
				throw OfficeExtension._Internal.RuntimeError._createInvalidArgError(name);
			}
		};
		Utility.isNullOrUndefined=function (value) {
			if (value===null) {
				return true;
			}
			if (typeof (value)==="undefined") {
				return true;
			}
			return false;
		};
		Utility.isUndefined=function (value) {
			if (typeof (value)==="undefined") {
				return true;
			}
			return false;
		};
		Utility.isNullOrEmptyString=function (value) {
			if (value===null) {
				return true;
			}
			if (typeof (value)==="undefined") {
				return true;
			}
			if (value.length==0) {
				return true;
			}
			return false;
		};
		Utility.isPlainJsonObject=function (value) {
			if (Utility.isNullOrUndefined(value)) {
				return false;
			}
			if (typeof (value) !=="object") {
				return false;
			}
			return Object.getPrototypeOf(value)===Object.getPrototypeOf({});
		};
		Utility.trim=function (str) {
			return str.replace(new RegExp("^\\s+|\\s+$", "g"), "");
		};
		Utility.caseInsensitiveCompareString=function (str1, str2) {
			if (Utility.isNullOrUndefined(str1)) {
				return Utility.isNullOrUndefined(str2);
			}
			else {
				if (Utility.isNullOrUndefined(str2)) {
					return false;
				}
				else {
					return str1.toUpperCase()==str2.toUpperCase();
				}
			}
		};
		Utility.adjustToDateTime=function (value) {
			if (Utility.isNullOrUndefined(value)) {
				return null;
			}
			if (typeof (value)==="string") {
				return new Date(value);
			}
			if (Array.isArray(value)) {
				var arr=value;
				for (var i=0; i < arr.length; i++) {
					arr[i]=Utility.adjustToDateTime(arr[i]);
				}
				return arr;
			}
			throw OfficeExtension._Internal.RuntimeError._createInvalidArgError("date");
		};
		Utility.isReadonlyRestRequest=function (method) {
			return Utility.caseInsensitiveCompareString(method, "GET");
		};
		Utility.setMethodArguments=function (context, argumentInfo, args) {
			if (Utility.isNullOrUndefined(args)) {
				return null;
			}
			var referencedObjectPaths=new Array();
			var referencedObjectPathIds=new Array();
			var hasOne=Utility.collectObjectPathInfos(context, args, referencedObjectPaths, referencedObjectPathIds);
			argumentInfo.Arguments=args;
			if (hasOne) {
				argumentInfo.ReferencedObjectPathIds=referencedObjectPathIds;
				return referencedObjectPaths;
			}
			return null;
		};
		Utility.collectObjectPathInfos=function (context, args, referencedObjectPaths, referencedObjectPathIds) {
			var hasOne=false;
			for (var i=0; i < args.length; i++) {
				if (args[i] instanceof OfficeExtension.ClientObject) {
					var clientObject=args[i];
					Utility.validateContext(context, clientObject);
					args[i]=clientObject._objectPath.objectPathInfo.Id;
					referencedObjectPathIds.push(clientObject._objectPath.objectPathInfo.Id);
					referencedObjectPaths.push(clientObject._objectPath);
					hasOne=true;
				}
				else if (Array.isArray(args[i])) {
					var childArrayObjectPathIds=new Array();
					var childArrayHasOne=Utility.collectObjectPathInfos(context, args[i], referencedObjectPaths, childArrayObjectPathIds);
					if (childArrayHasOne) {
						referencedObjectPathIds.push(childArrayObjectPathIds);
						hasOne=true;
					}
					else {
						referencedObjectPathIds.push(0);
					}
				}
				else {
					referencedObjectPathIds.push(0);
				}
			}
			return hasOne;
		};
		Utility.fixObjectPathIfNecessary=function (clientObject, value) {
			if (clientObject && clientObject._objectPath && value) {
				clientObject._objectPath.updateUsingObjectData(value, clientObject);
			}
		};
		Utility.tryGetObjectIdFromLoadOrRetrieveResult=function (value) {
			var id=value[OfficeExtension.Constants.id];
			if (Utility.isNullOrUndefined(id)) {
				id=value[OfficeExtension.Constants.idLowerCase];
			}
			if (Utility.isNullOrUndefined(id)) {
				id=value[OfficeExtension.Constants.idPrivate];
			}
			return id;
		};
		Utility.validateObjectPath=function (clientObject) {
			var objectPath=clientObject._objectPath;
			while (objectPath) {
				if (!objectPath.isValid) {
					throw new OfficeExtension._Internal.RuntimeError({
						code: OfficeExtension.ErrorCodes.invalidObjectPath,
						message: Utility._getResourceString(OfficeExtension.ResourceStrings.invalidObjectPath, Utility.getObjectPathExpression(objectPath)),
						debugInfo: {
							errorLocation: Utility.getObjectPathExpression(objectPath)
						}
					});
				}
				objectPath=objectPath.parentObjectPath;
			}
		};
		Utility.validateReferencedObjectPaths=function (objectPaths) {
			if (objectPaths) {
				for (var i=0; i < objectPaths.length; i++) {
					var objectPath=objectPaths[i];
					while (objectPath) {
						if (!objectPath.isValid) {
							throw new OfficeExtension._Internal.RuntimeError({
								code: OfficeExtension.ErrorCodes.invalidObjectPath,
								message: Utility._getResourceString(OfficeExtension.ResourceStrings.invalidObjectPath, Utility.getObjectPathExpression(objectPath))
							});
						}
						objectPath=objectPath.parentObjectPath;
					}
				}
			}
		};
		Utility.validateContext=function (context, obj) {
			if (obj && obj.context !==context) {
				throw new OfficeExtension._Internal.RuntimeError({
					code: OfficeExtension.ErrorCodes.invalidRequestContext,
					message: Utility._getResourceString(OfficeExtension.ResourceStrings.invalidRequestContext)
				});
			}
		};
		Utility.log=function (message) {
			if (Utility._logEnabled && typeof (console) !=="undefined" && console.log) {
				console.log(message);
			}
		};
		Utility.load=function (clientObj, option) {
			clientObj.context.load(clientObj, option);
			return clientObj;
		};
		Utility.loadAndSync=function (clientObj, option) {
			clientObj.context.load(clientObj, option);
			return clientObj.context.sync().then(function () { return clientObj; });
		};
		Utility.retrieve=function (clientObj, option) {
			var shouldPolyfill=OfficeExtension._internalConfig.alwaysPolyfillClientObjectRetrieveMethod;
			if (!shouldPolyfill) {
				shouldPolyfill=!Utility.isSetSupported("RichApiRuntime", "1.1");
			}
			var result=new OfficeExtension.RetrieveResultImpl(clientObj, shouldPolyfill);
			var queryOption=OfficeExtension.ClientRequestContext._parseQueryOption(option);
			var action;
			if (shouldPolyfill) {
				action=OfficeExtension.ActionFactory.createQueryAction(clientObj.context, clientObj, queryOption);
			}
			else {
				action=OfficeExtension.ActionFactory.createQueryAsJsonAction(clientObj.context, clientObj, queryOption);
			}
			clientObj.context._pendingRequest.addActionResultHandler(action, result);
			return result;
		};
		Utility.retrieveAndSync=function (clientObj, option) {
			var result=Utility.retrieve(clientObj, option);
			return clientObj.context.sync().then(function () { return result; });
		};
		Utility.isSetSupported=function (apiSetName, apiSetVersion) {
			if (typeof (window) !=="undefined" && window.Office && window.Office.context && window.Office.context.requirements) {
				return window.Office.context.requirements.isSetSupported(apiSetName, apiSetVersion);
			}
			return true;
		};
		Utility._parseSelectExpand=function (select) {
			var args=[];
			if (!Utility.isNullOrEmptyString(select)) {
				var propertyNames=select.split(",");
				for (var i=0; i < propertyNames.length; i++) {
					var propertyName=propertyNames[i];
					propertyName=sanitizeForAnyItemsSlash(propertyName.trim());
					if (propertyName.length > 0) {
						args.push(propertyName);
					}
				}
			}
			return args;
			function sanitizeForAnyItemsSlash(propertyName) {
				var propertyNameLower=propertyName.toLowerCase();
				if (propertyNameLower==="items" || propertyNameLower==="items/") {
					return '*';
				}
				var itemsSlashLength=6;
				var isItemsSlashOrItemsDot=propertyNameLower.substr(0, itemsSlashLength)==="items/" ||
					propertyNameLower.substr(0, itemsSlashLength)==="items.";
				if (isItemsSlashOrItemsDot) {
					propertyName=propertyName.substr(itemsSlashLength);
				}
				return propertyName.replace(new RegExp("[\/\.]items[\/\.]", "gi"), "/");
			}
		};
		Utility.toJson=function (clientObj, scalarProperties, navigationProperties, collectionItemsIfAny) {
			var result={};
			for (var prop in scalarProperties) {
				var value=scalarProperties[prop];
				if (typeof value !=="undefined") {
					result[prop]=value;
				}
			}
			for (var prop in navigationProperties) {
				var value=navigationProperties[prop];
				if (typeof value !=="undefined") {
					if (value[Utility.fieldName_isCollection] && (typeof value[Utility.fieldName_m__items] !=="undefined")) {
						result[prop]=value.toJSON()["items"];
					}
					else {
						result[prop]=value.toJSON();
					}
				}
			}
			if (collectionItemsIfAny) {
				result["items"]=collectionItemsIfAny.map(function (item) { return item.toJSON(); });
			}
			return result;
		};
		Utility.throwError=function (resourceId, arg, errorLocation) {
			throw new OfficeExtension._Internal.RuntimeError({
				code: resourceId,
				message: Utility._getResourceString(resourceId, arg),
				debugInfo: errorLocation ? { errorLocation: errorLocation } : undefined
			});
		};
		Utility.createRuntimeError=function (code, message, location) {
			return (new OfficeExtension._Internal.RuntimeError({
				code: code,
				message: message,
				debugInfo: { errorLocation: location }
			}));
		};
		Utility._getResourceString=function (resourceId, arg) {
			var ret;
			if (typeof (window) !=="undefined" && window.Strings && window.Strings.OfficeOM) {
				var stringName="L_"+resourceId;
				var stringValue=window.Strings.OfficeOM[stringName];
				if (stringValue) {
					ret=stringValue;
				}
			}
			if (!ret) {
				ret=OfficeExtension.ResourceStringValues[resourceId];
			}
			if (!ret) {
				ret=resourceId;
			}
			if (!Utility.isNullOrUndefined(arg)) {
				if (Array.isArray(arg)) {
					var arrArg=arg;
					ret=Utility._formatString(ret, arrArg);
				}
				else {
					ret=ret.replace("{0}", arg);
				}
			}
			return ret;
		};
		Utility._formatString=function (format, arrArg) {
			return format.replace(/\{\d\}/g, function (v) {
				var position=parseInt(v.substr(1, v.length - 2));
				if (position < arrArg.length) {
					return arrArg[position];
				}
				else {
					throw OfficeExtension._Internal.RuntimeError._createInvalidArgError("format");
				}
			});
		};
		Utility.throwIfNotLoaded=function (propertyName, fieldValue, entityName, isNull) {
			if (!isNull && Utility.isUndefined(fieldValue) && propertyName.charCodeAt(0) !=Utility.s_underscoreCharCode) {
				throw Utility.createPropertyNotLoadedException(entityName, propertyName);
			}
		};
		Utility.createPropertyNotLoadedException=function (entityName, propertyName) {
			return new OfficeExtension._Internal.RuntimeError({
				code: OfficeExtension.ErrorCodes.propertyNotLoaded,
				message: Utility._getResourceString(OfficeExtension.ResourceStrings.propertyNotLoaded, propertyName),
				debugInfo: entityName ? { errorLocation: entityName+"."+propertyName } : undefined
			});
		};
		Utility.createCannotUpdateReadOnlyPropertyException=function (entityName, propertyName) {
			return new OfficeExtension._Internal.RuntimeError({
				code: OfficeExtension.ErrorCodes.cannotUpdateReadOnlyProperty,
				message: Utility._getResourceString(OfficeExtension.ResourceStrings.cannotUpdateReadOnlyProperty, propertyName),
				debugInfo: entityName ? { errorLocation: entityName+"."+propertyName } : undefined
			});
		};
		Utility.throwIfApiNotSupported=function (apiFullName, apiSetName, apiSetVersion, hostName) {
			if (!Utility._doApiNotSupportedCheck) {
				return;
			}
			if (!Utility.isSetSupported(apiSetName, apiSetVersion)) {
				var message=Utility._getResourceString(OfficeExtension.ResourceStrings.apiNotFoundDetails, [apiFullName, apiSetName+" "+apiSetVersion, hostName]);
				throw new OfficeExtension._Internal.RuntimeError({
					code: OfficeExtension.ErrorCodes.apiNotFound,
					message: message,
					debugInfo: { errorLocation: apiFullName }
				});
			}
		};
		Utility.getObjectPathExpression=function (objectPath) {
			var ret="";
			while (objectPath) {
				switch (objectPath.objectPathInfo.ObjectPathType) {
					case 1:
						ret=ret;
						break;
					case 2:
						ret="new()"+(ret.length > 0 ? "." : "")+ret;
						break;
					case 3:
						ret=Utility.normalizeName(objectPath.objectPathInfo.Name)+"()"+(ret.length > 0 ? "." : "")+ret;
						break;
					case 4:
						ret=Utility.normalizeName(objectPath.objectPathInfo.Name)+(ret.length > 0 ? "." : "")+ret;
						break;
					case 5:
						ret="getItem()"+(ret.length > 0 ? "." : "")+ret;
						break;
					case 6:
						ret="_reference()"+(ret.length > 0 ? "." : "")+ret;
						break;
				}
				objectPath=objectPath.parentObjectPath;
			}
			return ret;
		};
		Utility._createPromiseFromResult=function (value) {
			return new OfficeExtension._Internal.OfficePromise(function (resolve, reject) {
				resolve(value);
			});
		};
		Utility._createTimeoutPromise=function (timeout) {
			return new OfficeExtension._Internal.OfficePromise(function (resolve, reject) {
				setTimeout(function () {
					resolve(null);
				}, timeout);
			});
		};
		Utility.promisify=function (action) {
			return new OfficeExtension._Internal.OfficePromise(function (resolve, reject) {
				var callback=function (result) {
					if (result.status=="failed") {
						reject(result.error);
					}
					else {
						resolve(result.value);
					}
				};
				action(callback);
			});
		};
		Utility._addActionResultHandler=function (clientObj, action, resultHandler) {
			clientObj.context._pendingRequest.addActionResultHandler(action, resultHandler);
		};
		Utility._handleNavigationPropertyResults=function (clientObj, objectValue, propertyNames) {
			for (var i=0; i < propertyNames.length - 1; i+=2) {
				if (!Utility.isUndefined(objectValue[propertyNames[i+1]])) {
					clientObj[propertyNames[i]]._handleResult(objectValue[propertyNames[i+1]]);
				}
			}
		};
		Utility.normalizeName=function (name) {
			return name.substr(0, 1).toLowerCase()+name.substr(1);
		};
		Utility._isLocalDocumentUrl=function (url) {
			return Utility._getLocalDocumentUrlPrefixLength(url) > 0;
		};
		Utility._getLocalDocumentUrlPrefixLength=function (url) {
			var localDocumentPrefixes=["http://document.localhost", "https://document.localhost", "//document.localhost"];
			var urlLower=url.toLowerCase().trim();
			for (var i=0; i < localDocumentPrefixes.length; i++) {
				if (urlLower===localDocumentPrefixes[i]) {
					return localDocumentPrefixes[i].length;
				}
				else if (urlLower.substr(0, localDocumentPrefixes[i].length+1)===localDocumentPrefixes[i]+"/") {
					return localDocumentPrefixes[i].length+1;
				}
			}
			return 0;
		};
		Utility._validateLocalDocumentRequest=function (request) {
			var index=Utility._getLocalDocumentUrlPrefixLength(request.url);
			if (index <=0) {
				throw OfficeExtension._Internal.RuntimeError._createInvalidArgError({
					argumentName: "request"
				});
			}
			var path=request.url.substr(index);
			var pathLower=path.toLowerCase();
			if (pathLower==="_api") {
				path="";
			}
			else if (pathLower.substr(0, "_api/".length)==="_api/") {
				path=path.substr("_api/".length);
			}
			return {
				method: request.method,
				url: path,
				headers: request.headers,
				body: request.body
			};
		};
		Utility._buildRequestMessageSafeArray=function (request) {
			var requestFlags=0;
			if (!Utility.isReadonlyRestRequest(request.method)) {
				requestFlags=1;
			}
			if (request.url.substr(0, OfficeExtension.Constants.processQuery.length).toLowerCase()===OfficeExtension.Constants.processQuery.toLowerCase()) {
				var index=request.url.indexOf("?");
				if (index > 0) {
					var queryString=request.url.substr(index+1);
					var parts=queryString.split("&");
					for (var i=0; i < parts.length; i++) {
						var keyvalue=parts[i].split("=");
						if (keyvalue[0].toLowerCase()===OfficeExtension.Constants.flags) {
							var flags=parseInt(keyvalue[1]);
							requestFlags=flags;
							requestFlags=requestFlags & 1;
							break;
						}
					}
				}
			}
			return OfficeExtension.RichApiMessageUtility.buildRequestMessageSafeArray("", requestFlags, request.method, request.url, request.headers, request.body);
		};
		Utility._parseHttpResponseHeaders=function (allResponseHeaders) {
			var responseHeaders={};
			if (!Utility.isNullOrEmptyString(allResponseHeaders)) {
				var regex=new RegExp("\r?\n");
				var entries=allResponseHeaders.split(regex);
				for (var i=0; i < entries.length; i++) {
					var entry=entries[i];
					if (entry !=null) {
						var index=entry.indexOf(':');
						if (index > 0) {
							var key=entry.substr(0, index);
							var value=entry.substr(index+1);
							key=Utility.trim(key);
							value=Utility.trim(value);
							responseHeaders[key.toUpperCase()]=value;
						}
					}
				}
			}
			return responseHeaders;
		};
		Utility._parseErrorResponse=function (responseInfo) {
			var errorObj=null;
			if (Utility.isPlainJsonObject(responseInfo.body)) {
				errorObj=responseInfo.body;
			}
			else if (!Utility.isNullOrEmptyString(responseInfo.body)) {
				var errorResponseBody=Utility.trim(responseInfo.body);
				try {
					errorObj=JSON.parse(errorResponseBody);
				}
				catch (e) {
					Utility.log("Error when parse "+errorResponseBody);
				}
			}
			var errorMessage;
			var errorCode;
			if (!Utility.isNullOrUndefined(errorObj) && typeof (errorObj)==="object" && errorObj.error) {
				errorCode=errorObj.error.code;
				errorMessage=Utility._getResourceString(OfficeExtension.ResourceStrings.connectionFailureWithDetails, [responseInfo.statusCode.toString(), errorObj.error.code, errorObj.error.message]);
			}
			else {
				errorMessage=Utility._getResourceString(OfficeExtension.ResourceStrings.connectionFailureWithStatus, responseInfo.statusCode.toString());
			}
			if (Utility.isNullOrEmptyString(errorCode)) {
				errorCode=OfficeExtension.ErrorCodes.connectionFailure;
			}
			return { errorCode: errorCode, errorMessage: errorMessage };
		};
		Utility._copyHeaders=function (src, dest) {
			if (src && dest) {
				for (var key in src) {
					dest[key]=src[key];
				}
			}
		};
		Utility._toCamelLowerCase=function (name) {
			if (Utility.isNullOrEmptyString(name)) {
				return name;
			}
			var index=0;
			while (index < name.length && name.charCodeAt(index) >=65 && name.charCodeAt(index) <=90) {
				index++;
			}
			if (index < name.length) {
				return name.substr(0, index).toLowerCase()+name.substr(index);
			}
			else {
				return name.toLowerCase();
			}
		};
		Utility.definePropertyThrowUnloadedException=function (obj, typeName, propertyName) {
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
		Utility.defineReadOnlyPropertyWithValue=function (obj, propertyName, value) {
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
		Utility.processRetrieveResult=function (proxy, value, result, childItemCreateFunc) {
			if (Utility.isNullOrUndefined(value)) {
				return;
			}
			if (childItemCreateFunc) {
				var data=value[OfficeExtension.Constants.itemsLowerCase];
				if (Array.isArray(data)) {
					var itemsResult=[];
					for (var i=0; i < data.length; i++) {
						var itemProxy=childItemCreateFunc(data[i], i);
						var itemResult={};
						itemResult[OfficeExtension.Constants.proxy]=itemProxy;
						itemProxy._handleRetrieveResult(data[i], itemResult);
						itemsResult.push(itemResult);
					}
					Utility.defineReadOnlyPropertyWithValue(result, OfficeExtension.Constants.itemsLowerCase, itemsResult);
				}
			}
			else {
				var scalarPropertyNames=proxy[OfficeExtension.Constants.scalarPropertyNames];
				var navigationPropertyNames=proxy[OfficeExtension.Constants.navigationPropertyNames];
				var typeName=proxy[OfficeExtension.Constants.className];
				if (scalarPropertyNames) {
					for (var i=0; i < scalarPropertyNames.length; i++) {
						var propName=scalarPropertyNames[i];
						var propValue=value[propName];
						if (Utility.isUndefined(propValue)) {
							Utility.definePropertyThrowUnloadedException(result, typeName, propName);
						}
						else {
							Utility.defineReadOnlyPropertyWithValue(result, propName, propValue);
						}
					}
				}
				if (navigationPropertyNames) {
					for (var i=0; i < navigationPropertyNames.length; i++) {
						var propName=navigationPropertyNames[i];
						var propValue=value[propName];
						if (Utility.isUndefined(propValue)) {
							Utility.definePropertyThrowUnloadedException(result, typeName, propName);
						}
						else {
							var propProxy=proxy[propName];
							var propResult={};
							propProxy._handleRetrieveResult(propValue, propResult);
							propResult[OfficeExtension.Constants.proxy]=propProxy;
							if (Array.isArray(propResult[OfficeExtension.Constants.itemsLowerCase])) {
								propResult=propResult[OfficeExtension.Constants.itemsLowerCase];
							}
							Utility.defineReadOnlyPropertyWithValue(result, propName, propResult);
						}
					}
				}
			}
		};
		Utility.fieldName_m__items="m__items";
		Utility.fieldName_isCollection="_isCollection";
		Utility._logEnabled=false;
		Utility._synchronousCleanup=false;
		Utility._doApiNotSupportedCheck=false;
		Utility.s_underscoreCharCode="_".charCodeAt(0);
		return Utility;
	}());
	OfficeExtension.Utility=Utility;
})(OfficeExtension || (OfficeExtension={}));

var __extends=(this && this.__extends) || (function () {
	var extendStatics=Object.setPrototypeOf ||
		({ __proto__: [] } instanceof Array && function (d, b) { d.__proto__=b; }) ||
		function (d, b) { for (var p in b) if (b.hasOwnProperty(p)) d[p]=b[p]; };
	return function (d, b) {
		extendStatics(d, b);
		function __() { this.constructor=d; }
		d.prototype=b===null ? Object.create(b) : (__.prototype=b.prototype, new __());
	};
})();
var OfficeCore;
(function (OfficeCore) {
	var _hostName="OfficeCore";
	var _defaultApiSetName="AgaveVisualApi";
	var _createPropertyObjectPath=OfficeExtension.ObjectPathFactory.createPropertyObjectPath;
	var _createMethodObjectPath=OfficeExtension.ObjectPathFactory.createMethodObjectPath;
	var _createIndexerObjectPath=OfficeExtension.ObjectPathFactory.createIndexerObjectPath;
	var _createNewObjectObjectPath=OfficeExtension.ObjectPathFactory.createNewObjectObjectPath;
	var _createChildItemObjectPathUsingIndexer=OfficeExtension.ObjectPathFactory.createChildItemObjectPathUsingIndexer;
	var _createChildItemObjectPathUsingGetItemAt=OfficeExtension.ObjectPathFactory.createChildItemObjectPathUsingGetItemAt;
	var _createChildItemObjectPathUsingIndexerOrGetItemAt=OfficeExtension.ObjectPathFactory.createChildItemObjectPathUsingIndexerOrGetItemAt;
	var _createMethodAction=OfficeExtension.ActionFactory.createMethodAction;
	var _createEnsureUnchangedAction=OfficeExtension.ActionFactory.createEnsureUnchangedAction;
	var _createSetPropertyAction=OfficeExtension.ActionFactory.createSetPropertyAction;
	var _isNullOrUndefined=OfficeExtension.Utility.isNullOrUndefined;
	var _isUndefined=OfficeExtension.Utility.isUndefined;
	var _throwIfNotLoaded=OfficeExtension.Utility.throwIfNotLoaded;
	var _throwIfApiNotSupported=OfficeExtension.Utility.throwIfApiNotSupported;
	var _load=OfficeExtension.Utility.load;
	var _loadAndSync=OfficeExtension.Utility.loadAndSync;
	var _retrieve=OfficeExtension.Utility.retrieve;
	var _retrieveAndSync=OfficeExtension.Utility.retrieveAndSync;
	var _toJson=OfficeExtension.Utility.toJson;
	var _fixObjectPathIfNecessary=OfficeExtension.Utility.fixObjectPathIfNecessary;
	var _addActionResultHandler=OfficeExtension.Utility._addActionResultHandler;
	var _handleNavigationPropertyResults=OfficeExtension.Utility._handleNavigationPropertyResults;
	var _adjustToDateTime=OfficeExtension.Utility.adjustToDateTime;
	var _typeBiShim="BiShim";
	var BiShim=(function (_super) {
		__extends(BiShim, _super);
		function BiShim() {
			return _super !==null && _super.apply(this, arguments) || this;
		}
		Object.defineProperty(BiShim.prototype, "_className", {
			get: function () {
				return "BiShim";
			},
			enumerable: true,
			configurable: true
		});
		BiShim.prototype.initialize=function (capabilities) {
			_createMethodAction(this.context, this, "Initialize", 0, [capabilities], false);
		};
		BiShim.prototype.uninitialize=function () {
			_createMethodAction(this.context, this, "Uninitialize", 0, [], false);
		};
		BiShim.prototype.getData=function () {
			var action=_createMethodAction(this.context, this, "getData", 1, [], false);
			var ret=new OfficeExtension.ClientResult();
			_addActionResultHandler(this, action, ret);
			return ret;
		};
		BiShim.prototype._handleResult=function (value) {
			_super.prototype._handleResult.call(this, value);
			if (_isNullOrUndefined(value))
				return;
			var obj=value;
			_fixObjectPathIfNecessary(this, obj);
		};
		BiShim.newObject=function (context) {
			var ret=new OfficeCore.BiShim(context, _createNewObjectObjectPath(context, "Microsoft.AgaveVisual.BiShim", false, false));
			return ret;
		};
		BiShim.prototype.toJSON=function () {
			return _toJson(this, {}, {});
		};
		return BiShim;
	}(OfficeExtension.ClientObject));
	OfficeCore.BiShim=BiShim;
	var ErrorCodes;
	(function (ErrorCodes) {
		ErrorCodes.generalException="GeneralException";
	})(ErrorCodes=OfficeCore.ErrorCodes || (OfficeCore.ErrorCodes={}));
})(OfficeCore || (OfficeCore={}));
var OfficeCore;
(function (OfficeCore) {
	var _hostName="OfficeCore";
	var _defaultApiSetName="ExperimentApi";
	var _createPropertyObjectPath=OfficeExtension.ObjectPathFactory.createPropertyObjectPath;
	var _createMethodObjectPath=OfficeExtension.ObjectPathFactory.createMethodObjectPath;
	var _createIndexerObjectPath=OfficeExtension.ObjectPathFactory.createIndexerObjectPath;
	var _createNewObjectObjectPath=OfficeExtension.ObjectPathFactory.createNewObjectObjectPath;
	var _createChildItemObjectPathUsingIndexer=OfficeExtension.ObjectPathFactory.createChildItemObjectPathUsingIndexer;
	var _createChildItemObjectPathUsingGetItemAt=OfficeExtension.ObjectPathFactory.createChildItemObjectPathUsingGetItemAt;
	var _createChildItemObjectPathUsingIndexerOrGetItemAt=OfficeExtension.ObjectPathFactory.createChildItemObjectPathUsingIndexerOrGetItemAt;
	var _createMethodAction=OfficeExtension.ActionFactory.createMethodAction;
	var _createSetPropertyAction=OfficeExtension.ActionFactory.createSetPropertyAction;
	var _isNullOrUndefined=OfficeExtension.Utility.isNullOrUndefined;
	var _isUndefined=OfficeExtension.Utility.isUndefined;
	var _throwIfNotLoaded=OfficeExtension.Utility.throwIfNotLoaded;
	var _throwIfApiNotSupported=OfficeExtension.Utility.throwIfApiNotSupported;
	var _load=OfficeExtension.Utility.load;
	var _fixObjectPathIfNecessary=OfficeExtension.Utility.fixObjectPathIfNecessary;
	var _addActionResultHandler=OfficeExtension.Utility._addActionResultHandler;
	var _handleNavigationPropertyResults=OfficeExtension.Utility._handleNavigationPropertyResults;
	var _adjustToDateTime=OfficeExtension.Utility.adjustToDateTime;
	var FlightingService=(function (_super) {
		__extends(FlightingService, _super);
		function FlightingService() {
			return _super !==null && _super.apply(this, arguments) || this;
		}
		Object.defineProperty(FlightingService.prototype, "_className", {
			get: function () {
				return "FlightingService";
			},
			enumerable: true,
			configurable: true
		});
		FlightingService.prototype.getClientSessionId=function () {
			var action=_createMethodAction(this.context, this, "GetClientSessionId", 1, []);
			var ret=new OfficeExtension.ClientResult();
			_addActionResultHandler(this, action, ret);
			return ret;
		};
		FlightingService.prototype.getDeferredFlights=function () {
			var action=_createMethodAction(this.context, this, "GetDeferredFlights", 1, []);
			var ret=new OfficeExtension.ClientResult();
			_addActionResultHandler(this, action, ret);
			return ret;
		};
		FlightingService.prototype.getFeature=function (featureName, type, defaultValue, possibleValues) {
			return new OfficeCore.ABType(this.context, _createMethodObjectPath(this.context, this, "GetFeature", 1, [featureName, type, defaultValue, possibleValues], false, false, null));
		};
		FlightingService.prototype.getFeatureGate=function (featureName, scope) {
			return new OfficeCore.ABType(this.context, _createMethodObjectPath(this.context, this, "GetFeatureGate", 1, [featureName, scope], false, false, null));
		};
		FlightingService.prototype.resetOverride=function (featureName) {
			_createMethodAction(this.context, this, "ResetOverride", 0, [featureName]);
		};
		FlightingService.prototype.setOverride=function (featureName, type, value) {
			_createMethodAction(this.context, this, "SetOverride", 0, [featureName, type, value]);
		};
		FlightingService.prototype._handleResult=function (value) {
			_super.prototype._handleResult.call(this, value);
			if (_isNullOrUndefined(value))
				return;
			var obj=value;
			_fixObjectPathIfNecessary(this, obj);
		};
		FlightingService.newObject=function (context) {
			var ret=new OfficeCore.FlightingService(context, _createNewObjectObjectPath(context, "Microsoft.Experiment.FlightingService", false));
			return ret;
		};
		FlightingService.prototype.toJSON=function () {
			return {};
		};
		return FlightingService;
	}(OfficeExtension.ClientObject));
	OfficeCore.FlightingService=FlightingService;
	var ABType=(function (_super) {
		__extends(ABType, _super);
		function ABType() {
			return _super !==null && _super.apply(this, arguments) || this;
		}
		Object.defineProperty(ABType.prototype, "_className", {
			get: function () {
				return "ABType";
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ABType.prototype, "value", {
			get: function () {
				_throwIfNotLoaded("value", this.m_value, "ABType", this._isNull);
				return this.m_value;
			},
			enumerable: true,
			configurable: true
		});
		ABType.prototype._handleResult=function (value) {
			_super.prototype._handleResult.call(this, value);
			if (_isNullOrUndefined(value))
				return;
			var obj=value;
			_fixObjectPathIfNecessary(this, obj);
			if (!_isUndefined(obj["Value"])) {
				this.m_value=obj["Value"];
			}
		};
		ABType.prototype.load=function (option) {
			_load(this, option);
			return this;
		};
		ABType.prototype.toJSON=function () {
			return {
				"value": this.m_value
			};
		};
		return ABType;
	}(OfficeExtension.ClientObject));
	OfficeCore.ABType=ABType;
	var FeatureType;
	(function (FeatureType) {
		FeatureType.boolean="Boolean";
		FeatureType.integer="Integer";
		FeatureType.string="String";
	})(FeatureType=OfficeCore.FeatureType || (OfficeCore.FeatureType={}));
	var ExperimentErrorCodes;
	(function (ExperimentErrorCodes) {
		ExperimentErrorCodes.generalException="GeneralException";
	})(ExperimentErrorCodes=OfficeCore.ExperimentErrorCodes || (OfficeCore.ExperimentErrorCodes={}));
})(OfficeCore || (OfficeCore={}));
var OfficeCore;
(function (OfficeCore) {
	var FirstPartyApis=(function () {
		function FirstPartyApis(context) {
			this.context=context;
		}
		Object.defineProperty(FirstPartyApis.prototype, "authentication", {
			get: function () {
				if (!this.m_authentication) {
					this.m_authentication=OfficeCore.AuthenticationService.newObject(this.context);
				}
				return this.m_authentication;
			},
			enumerable: true,
			configurable: true
		});
		return FirstPartyApis;
	}());
	OfficeCore.FirstPartyApis=FirstPartyApis;
	var RequestContext=(function (_super) {
		__extends(RequestContext, _super);
		function RequestContext(url) {
			return _super.call(this, url) || this;
		}
		Object.defineProperty(RequestContext.prototype, "firstParty", {
			get: function () {
				if (!this.m_firstPartyApis) {
					this.m_firstPartyApis=new FirstPartyApis(this);
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
					this.m_telemetry=OfficeCore.TelemetryService.newObject(this);
				}
				return this.m_telemetry;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(RequestContext.prototype, "bi", {
			get: function () {
				if (!this.m_biShim) {
					this.m_biShim=OfficeCore.BiShim.newObject(this);
				}
				return this.m_biShim;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(RequestContext.prototype, "flightingService", {
			get: function () {
				if (!this.m_flightingService) {
					this.m_flightingService=OfficeCore.FlightingService.newObject(this);
				}
				return this.m_flightingService;
			},
			enumerable: true,
			configurable: true
		});
		return RequestContext;
	}(OfficeExtension.ClientRequestContext));
	OfficeCore.RequestContext=RequestContext;
})(OfficeCore || (OfficeCore={}));
var OfficeCore;
(function (OfficeCore) {
	var _hostName="OfficeCore";
	var _defaultApiSetName="TelemetryApi";
	var _createPropertyObjectPath=OfficeExtension.ObjectPathFactory.createPropertyObjectPath;
	var _createMethodObjectPath=OfficeExtension.ObjectPathFactory.createMethodObjectPath;
	var _createIndexerObjectPath=OfficeExtension.ObjectPathFactory.createIndexerObjectPath;
	var _createNewObjectObjectPath=OfficeExtension.ObjectPathFactory.createNewObjectObjectPath;
	var _createChildItemObjectPathUsingIndexer=OfficeExtension.ObjectPathFactory.createChildItemObjectPathUsingIndexer;
	var _createChildItemObjectPathUsingGetItemAt=OfficeExtension.ObjectPathFactory.createChildItemObjectPathUsingGetItemAt;
	var _createChildItemObjectPathUsingIndexerOrGetItemAt=OfficeExtension.ObjectPathFactory.createChildItemObjectPathUsingIndexerOrGetItemAt;
	var _createMethodAction=OfficeExtension.ActionFactory.createMethodAction;
	var _createSetPropertyAction=OfficeExtension.ActionFactory.createSetPropertyAction;
	var _isNullOrUndefined=OfficeExtension.Utility.isNullOrUndefined;
	var _isUndefined=OfficeExtension.Utility.isUndefined;
	var _throwIfNotLoaded=OfficeExtension.Utility.throwIfNotLoaded;
	var _throwIfApiNotSupported=OfficeExtension.Utility.throwIfApiNotSupported;
	var _load=OfficeExtension.Utility.load;
	var _fixObjectPathIfNecessary=OfficeExtension.Utility.fixObjectPathIfNecessary;
	var _addActionResultHandler=OfficeExtension.Utility._addActionResultHandler;
	var _handleNavigationPropertyResults=OfficeExtension.Utility._handleNavigationPropertyResults;
	var _adjustToDateTime=OfficeExtension.Utility.adjustToDateTime;
	var _typeTelemetryService="TelemetryService";
	var TelemetryService=(function (_super) {
		__extends(TelemetryService, _super);
		function TelemetryService() {
			return _super !==null && _super.apply(this, arguments) || this;
		}
		Object.defineProperty(TelemetryService.prototype, "_className", {
			get: function () {
				return "TelemetryService";
			},
			enumerable: true,
			configurable: true
		});
		TelemetryService.prototype.sendTelemetryEvent=function (telemetryProperties, eventName, eventContract, eventFlags, value) {
			_createMethodAction(this.context, this, "SendTelemetryEvent", 1, [telemetryProperties, eventName, eventContract, eventFlags, value], false);
		};
		TelemetryService.prototype._handleResult=function (value) {
			_super.prototype._handleResult.call(this, value);
			if (_isNullOrUndefined(value))
				return;
			var obj=value;
			_fixObjectPathIfNecessary(this, obj);
		};
		TelemetryService.newObject=function (context) {
			var ret=new OfficeCore.TelemetryService(context, _createNewObjectObjectPath(context, "Microsoft.Telemetry.TelemetryService", false, false));
			return ret;
		};
		TelemetryService.prototype.toJSON=function () {
			return {};
		};
		return TelemetryService;
	}(OfficeExtension.ClientObject));
	OfficeCore.TelemetryService=TelemetryService;
	var TelemetryErrorCodes;
	(function (TelemetryErrorCodes) {
		TelemetryErrorCodes.generalException="GeneralException";
	})(TelemetryErrorCodes=OfficeCore.TelemetryErrorCodes || (OfficeCore.TelemetryErrorCodes={}));
})(OfficeCore || (OfficeCore={}));
var OfficeCore;
(function (OfficeCore) {
	var _hostName="Office";
	var _defaultApiSetName="OfficeSharedApi";
	var _createPropertyObjectPath=OfficeExtension.ObjectPathFactory.createPropertyObjectPath;
	var _createMethodObjectPath=OfficeExtension.ObjectPathFactory.createMethodObjectPath;
	var _createIndexerObjectPath=OfficeExtension.ObjectPathFactory.createIndexerObjectPath;
	var _createNewObjectObjectPath=OfficeExtension.ObjectPathFactory.createNewObjectObjectPath;
	var _createChildItemObjectPathUsingIndexer=OfficeExtension.ObjectPathFactory.createChildItemObjectPathUsingIndexer;
	var _createChildItemObjectPathUsingGetItemAt=OfficeExtension.ObjectPathFactory.createChildItemObjectPathUsingGetItemAt;
	var _createChildItemObjectPathUsingIndexerOrGetItemAt=OfficeExtension.ObjectPathFactory.createChildItemObjectPathUsingIndexerOrGetItemAt;
	var _createMethodAction=OfficeExtension.ActionFactory.createMethodAction;
	var _createEnsureUnchangedAction=OfficeExtension.ActionFactory.createEnsureUnchangedAction;
	var _createSetPropertyAction=OfficeExtension.ActionFactory.createSetPropertyAction;
	var _isNullOrUndefined=OfficeExtension.Utility.isNullOrUndefined;
	var _isUndefined=OfficeExtension.Utility.isUndefined;
	var _throwIfNotLoaded=OfficeExtension.Utility.throwIfNotLoaded;
	var _throwIfApiNotSupported=OfficeExtension.Utility.throwIfApiNotSupported;
	var _load=OfficeExtension.Utility.load;
	var _retrieve=OfficeExtension.Utility.retrieve;
	var _toJson=OfficeExtension.Utility.toJson;
	var _fixObjectPathIfNecessary=OfficeExtension.Utility.fixObjectPathIfNecessary;
	var _addActionResultHandler=OfficeExtension.Utility._addActionResultHandler;
	var _handleNavigationPropertyResults=OfficeExtension.Utility._handleNavigationPropertyResults;
	var _adjustToDateTime=OfficeExtension.Utility.adjustToDateTime;
	var _processRetrieveResult=OfficeExtension.Utility.processRetrieveResult;
	var IdentityType;
	(function (IdentityType) {
		IdentityType.organizationAccount="OrganizationAccount";
		IdentityType.microsoftAccount="MicrosoftAccount";
	})(IdentityType=OfficeCore.IdentityType || (OfficeCore.IdentityType={}));
	var _typeAuthenticationService="AuthenticationService";
	var AuthenticationService=(function (_super) {
		__extends(AuthenticationService, _super);
		function AuthenticationService() {
			return _super !==null && _super.apply(this, arguments) || this;
		}
		Object.defineProperty(AuthenticationService.prototype, "_className", {
			get: function () {
				return "AuthenticationService";
			},
			enumerable: true,
			configurable: true
		});
		AuthenticationService.prototype.getAccessToken=function (tokenParameters) {
			var action=_createMethodAction(this.context, this, "GetAccessToken", 1, [tokenParameters], true);
			var ret=new OfficeExtension.ClientResult();
			_addActionResultHandler(this, action, ret);
			return ret;
		};
		AuthenticationService.prototype.getPrimaryIdentityType=function () {
			var action=_createMethodAction(this.context, this, "GetPrimaryIdentityType", 1, [], true);
			var ret=new OfficeExtension.ClientResult();
			_addActionResultHandler(this, action, ret);
			return ret;
		};
		AuthenticationService.prototype.hasUserSignin=function () {
			var action=_createMethodAction(this.context, this, "HasUserSignin", 1, [], true);
			var ret=new OfficeExtension.ClientResult();
			_addActionResultHandler(this, action, ret);
			return ret;
		};
		AuthenticationService.prototype._handleResult=function (value) {
			_super.prototype._handleResult.call(this, value);
			if (_isNullOrUndefined(value))
				return;
			var obj=value;
			_fixObjectPathIfNecessary(this, obj);
		};
		AuthenticationService.prototype._handleRetrieveResult=function (value, result) {
			_super.prototype._handleRetrieveResult.call(this, value, result);
			_processRetrieveResult(this, value, result);
		};
		AuthenticationService.newObject=function (context) {
			var ret=new OfficeCore.AuthenticationService(context, _createNewObjectObjectPath(context, "Microsoft.Authentication.AuthenticationService", false, false));
			return ret;
		};
		AuthenticationService.prototype.toJSON=function () {
			return _toJson(this, {}, {});
		};
		return AuthenticationService;
	}(OfficeExtension.ClientObject));
	OfficeCore.AuthenticationService=AuthenticationService;
	var _typeComment="Comment";
	var Comment=(function (_super) {
		__extends(Comment, _super);
		function Comment() {
			return _super !==null && _super.apply(this, arguments) || this;
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
					this._P=new OfficeCore.Comment(this.context, _createPropertyObjectPath(this.context, this, "Parent", false, false, false));
				}
				return this._P;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Comment.prototype, "parentOrNullObject", {
			get: function () {
				if (!this._Pa) {
					this._Pa=new OfficeCore.Comment(this.context, _createPropertyObjectPath(this.context, this, "ParentOrNullObject", false, false, false));
				}
				return this._Pa;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Comment.prototype, "replies", {
			get: function () {
				if (!this._R) {
					this._R=new OfficeCore.CommentCollection(this.context, _createPropertyObjectPath(this.context, this, "Replies", true, false, false));
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
				this._Re=value;
				_createSetPropertyAction(this.context, this, "Resolved", value);
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
				this._T=value;
				_createSetPropertyAction(this.context, this, "Text", value);
			},
			enumerable: true,
			configurable: true
		});
		Comment.prototype.set=function (properties, options) {
			this._recursivelySet(properties, options, ["text", "resolved"], [], [
				"parent",
				"parentOrNullObject",
				"replies"
			]);
		};
		Comment.prototype.update=function (properties) {
			this._recursivelyUpdate(properties);
		};
		Comment.prototype.delete=function () {
			_createMethodAction(this.context, this, "Delete", 0, [], false);
		};
		Comment.prototype.getParentOrSelf=function () {
			return new OfficeCore.Comment(this.context, _createMethodObjectPath(this.context, this, "GetParentOrSelf", 1, [], false, false, null, false));
		};
		Comment.prototype.getRichText=function (format) {
			var action=_createMethodAction(this.context, this, "GetRichText", 1, [format], false);
			var ret=new OfficeExtension.ClientResult();
			_addActionResultHandler(this, action, ret);
			return ret;
		};
		Comment.prototype.reply=function (text, format) {
			return new OfficeCore.Comment(this.context, _createMethodObjectPath(this.context, this, "Reply", 0, [text, format], false, false, null, false));
		};
		Comment.prototype.setRichText=function (text, format) {
			var action=_createMethodAction(this.context, this, "SetRichText", 0, [text, format], false);
			var ret=new OfficeExtension.ClientResult();
			_addActionResultHandler(this, action, ret);
			return ret;
		};
		Comment.prototype._handleResult=function (value) {
			_super.prototype._handleResult.call(this, value);
			if (_isNullOrUndefined(value))
				return;
			var obj=value;
			_fixObjectPathIfNecessary(this, obj);
			if (!_isUndefined(obj["Author"])) {
				this._A=obj["Author"];
			}
			if (!_isUndefined(obj["Created"])) {
				this._C=_adjustToDateTime(obj["Created"]);
			}
			if (!_isUndefined(obj["Id"])) {
				this._I=obj["Id"];
			}
			if (!_isUndefined(obj["Level"])) {
				this._L=obj["Level"];
			}
			if (!_isUndefined(obj["Mentions"])) {
				this._M=obj["Mentions"];
			}
			if (!_isUndefined(obj["Resolved"])) {
				this._Re=obj["Resolved"];
			}
			if (!_isUndefined(obj["Text"])) {
				this._T=obj["Text"];
			}
			_handleNavigationPropertyResults(this, obj, ["parent", "Parent", "parentOrNullObject", "ParentOrNullObject", "replies", "Replies"]);
		};
		Comment.prototype.load=function (option) {
			return _load(this, option);
		};
		Comment.prototype.retrieve=function (option) {
			return _retrieve(this, option);
		};
		Comment.prototype._handleIdResult=function (value) {
			_super.prototype._handleIdResult.call(this, value);
			if (_isNullOrUndefined(value)) {
				return;
			}
			if (!_isUndefined(value["Id"])) {
				this._I=value["Id"];
			}
		};
		Comment.prototype._handleRetrieveResult=function (value, result) {
			_super.prototype._handleRetrieveResult.call(this, value, result);
			if (_isNullOrUndefined(value))
				return;
			var obj=value;
			if (!_isUndefined(obj["Created"])) {
				obj["created"]=_adjustToDateTime(obj["created"]);
			}
			_processRetrieveResult(this, value, result);
		};
		Comment.prototype.toJSON=function () {
			return _toJson(this, {
				"author": this._A,
				"created": this._C,
				"id": this._I,
				"level": this._L,
				"mentions": this._M,
				"resolved": this._Re,
				"text": this._T,
			}, {
				"replies": this._R,
			});
		};
		Comment.prototype.ensureUnchanged=function (data) {
			_createEnsureUnchangedAction(this.context, this, data);
			return;
		};
		return Comment;
	}(OfficeExtension.ClientObject));
	OfficeCore.Comment=Comment;
	var _typeCommentCollection="CommentCollection";
	var CommentCollection=(function (_super) {
		__extends(CommentCollection, _super);
		function CommentCollection() {
			return _super !==null && _super.apply(this, arguments) || this;
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
		CommentCollection.prototype.getCount=function () {
			var action=_createMethodAction(this.context, this, "GetCount", 1, [], false);
			var ret=new OfficeExtension.ClientResult();
			_addActionResultHandler(this, action, ret);
			return ret;
		};
		CommentCollection.prototype.getItem=function (id) {
			return new OfficeCore.Comment(this.context, _createIndexerObjectPath(this.context, this, [id]));
		};
		CommentCollection.prototype._handleResult=function (value) {
			_super.prototype._handleResult.call(this, value);
			if (_isNullOrUndefined(value))
				return;
			var obj=value;
			_fixObjectPathIfNecessary(this, obj);
			if (!_isNullOrUndefined(obj[OfficeExtension.Constants.items])) {
				this.m__items=[];
				var _data=obj[OfficeExtension.Constants.items];
				for (var i=0; i < _data.length; i++) {
					var _item=new OfficeCore.Comment(this.context, _createChildItemObjectPathUsingIndexerOrGetItemAt(true, this.context, this, _data[i], i));
					_item._handleResult(_data[i]);
					this.m__items.push(_item);
				}
			}
		};
		CommentCollection.prototype.load=function (option) {
			return _load(this, option);
		};
		CommentCollection.prototype.retrieve=function (option) {
			return _retrieve(this, option);
		};
		CommentCollection.prototype._handleRetrieveResult=function (value, result) {
			var _this=this;
			_super.prototype._handleRetrieveResult.call(this, value, result);
			_processRetrieveResult(this, value, result, function (childItemData, index) { return new OfficeCore.Comment(_this.context, _createChildItemObjectPathUsingIndexerOrGetItemAt(true, _this.context, _this, childItemData, index)); });
		};
		CommentCollection.prototype.toJSON=function () {
			return _toJson(this, {}, {}, this.m__items);
		};
		return CommentCollection;
	}(OfficeExtension.ClientObject));
	OfficeCore.CommentCollection=CommentCollection;
	var CommentTextFormat;
	(function (CommentTextFormat) {
		CommentTextFormat.plain="Plain";
		CommentTextFormat.markdown="Markdown";
		CommentTextFormat.delta="Delta";
	})(CommentTextFormat=OfficeCore.CommentTextFormat || (OfficeCore.CommentTextFormat={}));
	var ErrorCodes;
	(function (ErrorCodes) {
		ErrorCodes.apiNotAvailable="ApiNotAvailable";
		ErrorCodes.clientError="ClientError";
		ErrorCodes.generalException="GeneralException";
		ErrorCodes.invalidArgument="InvalidArgument";
		ErrorCodes.invalidGrant="InvalidGrant";
		ErrorCodes.invalidResourceUrl="InvalidResourceUrl";
		ErrorCodes.serverError="ServerError";
		ErrorCodes.unsupportedUserIdentity="UnsupportedUserIdentity";
		ErrorCodes.userNotSignIn="UserNotSignIn";
	})(ErrorCodes=OfficeCore.ErrorCodes || (OfficeCore.ErrorCodes={}));
})(OfficeCore || (OfficeCore={}));

var __extends=(this && this.__extends) || (function () {
	var extendStatics=Object.setPrototypeOf ||
		({ __proto__: [] } instanceof Array && function (d, b) { d.__proto__=b; }) ||
		function (d, b) { for (var p in b) if (b.hasOwnProperty(p)) d[p]=b[p]; };
	return function (d, b) {
		extendStatics(d, b);
		function __() { this.constructor=d; }
		d.prototype=b===null ? Object.create(b) : (__.prototype=b.prototype, new __());
	};
})();

var Visio;
(function (Visio) {

	var _hostName="Visio";
	var _defaultApiSetName="";
	var _createPropertyObjectPath=OfficeExtension.ObjectPathFactory.createPropertyObjectPath;
	var _createMethodObjectPath=OfficeExtension.ObjectPathFactory.createMethodObjectPath;
	var _createIndexerObjectPath=OfficeExtension.ObjectPathFactory.createIndexerObjectPath;
	var _createNewObjectObjectPath=OfficeExtension.ObjectPathFactory.createNewObjectObjectPath;
	var _createChildItemObjectPathUsingIndexer=OfficeExtension.ObjectPathFactory.createChildItemObjectPathUsingIndexer;
	var _createChildItemObjectPathUsingGetItemAt=OfficeExtension.ObjectPathFactory.createChildItemObjectPathUsingGetItemAt;
	var _createChildItemObjectPathUsingIndexerOrGetItemAt=OfficeExtension.ObjectPathFactory.createChildItemObjectPathUsingIndexerOrGetItemAt;
	var _createMethodAction=OfficeExtension.ActionFactory.createMethodAction;
	var _createEnsureUnchangedAction=OfficeExtension.ActionFactory.createEnsureUnchangedAction;
	var _createSetPropertyAction=OfficeExtension.ActionFactory.createSetPropertyAction;
	var _isNullOrUndefined=OfficeExtension.Utility.isNullOrUndefined;
	var _isUndefined=OfficeExtension.Utility.isUndefined;
	var _throwIfNotLoaded=OfficeExtension.Utility.throwIfNotLoaded;
	var _throwIfApiNotSupported=OfficeExtension.Utility.throwIfApiNotSupported;
	var _load=OfficeExtension.Utility.load;
	var _retrieve=OfficeExtension.Utility.retrieve;
	var _toJson=OfficeExtension.Utility.toJson;
	var _fixObjectPathIfNecessary=OfficeExtension.Utility.fixObjectPathIfNecessary;
	var _addActionResultHandler=OfficeExtension.Utility._addActionResultHandler;
	var _handleNavigationPropertyResults=OfficeExtension.Utility._handleNavigationPropertyResults;
	var _adjustToDateTime=OfficeExtension.Utility.adjustToDateTime;
	var _processRetrieveResult=OfficeExtension.Utility.processRetrieveResult;

	var _typeApplication="Application";

	var Application=(function (_super) {
		__extends(Application, _super);
		function Application() {
			return _super !==null && _super.apply(this, arguments) || this;
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

				this._S=value;
				_createSetPropertyAction(this.context, this, "ShowBorders", value);
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

				this._Sh=value;
				_createSetPropertyAction(this.context, this, "ShowToolbars", value);
			},
			enumerable: true,
			configurable: true
		});
		Application.prototype.set=function (properties, options) {

			this._recursivelySet(properties, options, ["showBorders", "showToolbars"] , [] , []);
		};

		Application.prototype.update=function (properties) {

			this._recursivelyUpdate(properties);
		};

		Application.prototype.showToolbar=function (id, show) {

			_createMethodAction(this.context, this, "ShowToolbar", 0 , [id, show], false);
		};

		Application.prototype._handleResult=function (value) {
			_super.prototype._handleResult.call(this, value);
			if (_isNullOrUndefined(value))
				return;
			var obj=value;
			_fixObjectPathIfNecessary(this, obj);

			if (!_isUndefined(obj["ShowBorders"])) {
				this._S=obj["ShowBorders"];

			}
			if (!_isUndefined(obj["ShowToolbars"])) {
				this._Sh=obj["ShowToolbars"];

			}
		};
		Application.prototype.load=function (option) {
			return _load(this, option);
		};
		Application.prototype.retrieve=function (option) {
			return _retrieve(this, option);
		};

		Application.prototype._handleRetrieveResult=function (value, result) {
			_super.prototype._handleRetrieveResult.call(this, value, result);

			_processRetrieveResult(this, value, result);
		};
		Application.prototype.toJSON=function () {
			return _toJson(this,
			 {
				"showBorders": this._S,
				"showToolbars": this._Sh,
			},
			 {});
		};

		Application.prototype.ensureUnchanged=function (data) {
			_createEnsureUnchangedAction(this.context, this, data);
			return;
		};
		return Application;
	}(OfficeExtension.ClientObject));
	Visio.Application=Application;

	var _typeDocument="Document";

	var Document=(function (_super) {
		__extends(Document, _super);
		function Document() {
			return _super !==null && _super.apply(this, arguments) || this;
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
					this._A=new Visio.Application(this.context, _createPropertyObjectPath(this.context, this, "Application", false , false , false));
				}

				return this._A;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Document.prototype, "pages", {

			get: function () {

				if (!this._P) {
					this._P=new Visio.PageCollection(this.context, _createPropertyObjectPath(this.context, this, "Pages", true , false , false));
				}

				return this._P;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Document.prototype, "view", {

			get: function () {

				if (!this._V) {
					this._V=new Visio.DocumentView(this.context, _createPropertyObjectPath(this.context, this, "View", false , false , false));
				}

				return this._V;
			},
			enumerable: true,
			configurable: true
		});
		Document.prototype.set=function (properties, options) {

			this._recursivelySet(properties, options, [] , ["view", "application"] , [
				"pages"
			]);
		};

		Document.prototype.update=function (properties) {

			this._recursivelyUpdate(properties);
		};

		Document.prototype.getActivePage=function () {

			return new Visio.Page(this.context, _createMethodObjectPath(this.context, this, "GetActivePage", 1 , [], false , false , null , false));
		};

		Document.prototype.setActivePage=function (PageName) {

			_createMethodAction(this.context, this, "SetActivePage", 1 , [PageName], false);
		};

		Document.prototype.startDataRefresh=function () {

			_createMethodAction(this.context, this, "StartDataRefresh", 1 , [], false);
		};

		Document.prototype._handleResult=function (value) {
			_super.prototype._handleResult.call(this, value);
			if (_isNullOrUndefined(value))
				return;
			var obj=value;
			_fixObjectPathIfNecessary(this, obj);

			_handleNavigationPropertyResults(this, obj, ["application", "Application", "pages", "Pages", "view", "View"]);
		};
		Document.prototype.load=function (option) {
			return _load(this, option);
		};
		Document.prototype.retrieve=function (option) {
			return _retrieve(this, option);
		};

		Document.prototype._handleRetrieveResult=function (value, result) {
			_super.prototype._handleRetrieveResult.call(this, value, result);

			_processRetrieveResult(this, value, result);
		};
		Object.defineProperty(Document.prototype, "onDataRefreshComplete", {

			get: function () {

				var _this=this;

				if (!this.m_dataRefreshComplete) {

					this.m_dataRefreshComplete=new OfficeExtension.EventHandlers(

					this.context, this, "DataRefreshComplete",

					{
						registerFunc: function (handlerCallback) {
							return _this.context.eventRegistration.register(3, "", handlerCallback);
						},
						unregisterFunc: function (handlerCallback) {
							return _this.context.eventRegistration.unregister(3, "", handlerCallback);
						},
						eventArgsTransformFunc: function (args) {
							var evt={
								document: this,
								success: args.ddaBinding.Object.success
							};
							return OfficeExtension.Utility._createPromiseFromResult(evt);
						}
					}

					);

				}
				return this.m_dataRefreshComplete;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Document.prototype, "onDocumentLoadComplete", {

			get: function () {

				var _this=this;

				if (!this.m_documentLoadComplete) {

					this.m_documentLoadComplete=new OfficeExtension.EventHandlers(

					this.context, this, "DocumentLoadComplete",

					{
						registerFunc: function (handlerCallback) {
							return _this.context.eventRegistration.register(7, "", handlerCallback);
						},
						unregisterFunc: function (handlerCallback) {
							return _this.context.eventRegistration.unregister(7, "", handlerCallback);
						},
						eventArgsTransformFunc: function (args) {
							var evt={
								success: args.ddaBinding.Object.success
							};
							return OfficeExtension.Utility._createPromiseFromResult(evt);
						}
					}

					);

				}
				return this.m_documentLoadComplete;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Document.prototype, "onPageLoadComplete", {

			get: function () {

				var _this=this;

				if (!this.m_pageLoadComplete) {

					this.m_pageLoadComplete=new OfficeExtension.EventHandlers(

					this.context, this, "PageLoadComplete",

					{
						registerFunc: function (handlerCallback) {
							return _this.context.eventRegistration.register(1, "", handlerCallback);
						},
						unregisterFunc: function (handlerCallback) {
							return _this.context.eventRegistration.unregister(1, "", handlerCallback);
						},
						eventArgsTransformFunc: function (args) {
							return OfficeExtension.Utility._createPromiseFromResult(args.ddaBinding.Object);
						}
					}

					);

				}
				return this.m_pageLoadComplete;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Document.prototype, "onSelectionChanged", {

			get: function () {

				var _this=this;

				if (!this.m_selectionChanged) {

					this.m_selectionChanged=new OfficeExtension.EventHandlers(

					this.context, this, "SelectionChanged",

					{
						registerFunc: function (handlerCallback) {
							return _this.context.eventRegistration.register(2, "", handlerCallback);
						},
						unregisterFunc: function (handlerCallback) {
							return _this.context.eventRegistration.unregister(2, "", handlerCallback);
						},
						eventArgsTransformFunc: function (args) {
							return OfficeExtension.Utility._createPromiseFromResult(args.ddaBinding.Object);
						}
					}

					);

				}
				return this.m_selectionChanged;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Document.prototype, "onShapeMouseEnter", {

			get: function () {

				var _this=this;

				if (!this.m_shapeMouseEnter) {

					this.m_shapeMouseEnter=new OfficeExtension.EventHandlers(

					this.context, this, "ShapeMouseEnter",

					{
						registerFunc: function (handlerCallback) {
							return _this.context.eventRegistration.register(4, "", handlerCallback);
						},
						unregisterFunc: function (handlerCallback) {
							return _this.context.eventRegistration.unregister(4, "", handlerCallback);
						},
						eventArgsTransformFunc: function (args) {
							return OfficeExtension.Utility._createPromiseFromResult(args.ddaBinding.Object);
						}
					}

					);

				}
				return this.m_shapeMouseEnter;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Document.prototype, "onShapeMouseLeave", {

			get: function () {

				var _this=this;

				if (!this.m_shapeMouseLeave) {

					this.m_shapeMouseLeave=new OfficeExtension.EventHandlers(

					this.context, this, "ShapeMouseLeave",

					{
						registerFunc: function (handlerCallback) {
							return _this.context.eventRegistration.register(5, "", handlerCallback);
						},
						unregisterFunc: function (handlerCallback) {
							return _this.context.eventRegistration.unregister(5, "", handlerCallback);
						},
						eventArgsTransformFunc: function (args) {
							return OfficeExtension.Utility._createPromiseFromResult(args.ddaBinding.Object);
						}
					}

					);

				}
				return this.m_shapeMouseLeave;
			},
			enumerable: true,
			configurable: true
		});
		Document.prototype.toJSON=function () {
			return _toJson(this,
			 {},
			 {
				"application": this._A,
				"pages": this._P,
				"view": this._V,
			});
		};

		Document.prototype.ensureUnchanged=function (data) {
			_createEnsureUnchangedAction(this.context, this, data);
			return;
		};
		return Document;
	}(OfficeExtension.ClientObject));
	Visio.Document=Document;

	var _typeDocumentView="DocumentView";

	var DocumentView=(function (_super) {
		__extends(DocumentView, _super);
		function DocumentView() {
			return _super !==null && _super.apply(this, arguments) || this;
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
				return ["disableHyperlinks", "disableZoom", "disablePan", "hideDiagramBoundary"];
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(DocumentView.prototype, "_scalarPropertyUpdateable", {
			get: function () {
				return [true, true, true, true];
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

				this._D=value;
				_createSetPropertyAction(this.context, this, "DisableHyperlinks", value);
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

				this._Di=value;
				_createSetPropertyAction(this.context, this, "DisablePan", value);
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(DocumentView.prototype, "disableZoom", {

			get: function () {

				_throwIfNotLoaded("disableZoom", this._Dis, _typeDocumentView, this._isNull);
				return this._Dis;
			},
			set: function (value) {

				this._Dis=value;
				_createSetPropertyAction(this.context, this, "DisableZoom", value);
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

				this._H=value;
				_createSetPropertyAction(this.context, this, "HideDiagramBoundary", value);
			},
			enumerable: true,
			configurable: true
		});
		DocumentView.prototype.set=function (properties, options) {

			this._recursivelySet(properties, options, ["disableHyperlinks", "disableZoom", "disablePan", "hideDiagramBoundary"] , [] , []);
		};

		DocumentView.prototype.update=function (properties) {

			this._recursivelyUpdate(properties);
		};

		DocumentView.prototype._handleResult=function (value) {
			_super.prototype._handleResult.call(this, value);
			if (_isNullOrUndefined(value))
				return;
			var obj=value;
			_fixObjectPathIfNecessary(this, obj);

			if (!_isUndefined(obj["DisableHyperlinks"])) {
				this._D=obj["DisableHyperlinks"];

			}
			if (!_isUndefined(obj["DisablePan"])) {
				this._Di=obj["DisablePan"];

			}
			if (!_isUndefined(obj["DisableZoom"])) {
				this._Dis=obj["DisableZoom"];

			}
			if (!_isUndefined(obj["HideDiagramBoundary"])) {
				this._H=obj["HideDiagramBoundary"];

			}
		};
		DocumentView.prototype.load=function (option) {
			return _load(this, option);
		};
		DocumentView.prototype.retrieve=function (option) {
			return _retrieve(this, option);
		};

		DocumentView.prototype._handleRetrieveResult=function (value, result) {
			_super.prototype._handleRetrieveResult.call(this, value, result);

			_processRetrieveResult(this, value, result);
		};
		DocumentView.prototype.toJSON=function () {
			return _toJson(this,
			 {
				"disableHyperlinks": this._D,
				"disablePan": this._Di,
				"disableZoom": this._Dis,
				"hideDiagramBoundary": this._H,
			},
			 {});
		};

		DocumentView.prototype.ensureUnchanged=function (data) {
			_createEnsureUnchangedAction(this.context, this, data);
			return;
		};
		return DocumentView;
	}(OfficeExtension.ClientObject));
	Visio.DocumentView=DocumentView;

	var _typePage="Page";

	var Page=(function (_super) {
		__extends(Page, _super);
		function Page() {
			return _super !==null && _super.apply(this, arguments) || this;
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
		Object.defineProperty(Page.prototype, "_navigationPropertyNames", {
			get: function () {
				return ["shapes", "view", "comments", "allShapes"];
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Page.prototype, "allShapes", {

			get: function () {

				if (!this._A) {
					this._A=new Visio.ShapeCollection(this.context, _createPropertyObjectPath(this.context, this, "AllShapes", true , false , false));
				}

				return this._A;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Page.prototype, "comments", {

			get: function () {

				if (!this._C) {
					this._C=new Visio.CommentCollection(this.context, _createPropertyObjectPath(this.context, this, "Comments", true , false , false));
				}

				return this._C;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Page.prototype, "shapes", {

			get: function () {

				if (!this._S) {
					this._S=new Visio.ShapeCollection(this.context, _createPropertyObjectPath(this.context, this, "Shapes", true , false , false));
				}

				return this._S;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Page.prototype, "view", {

			get: function () {

				if (!this._V) {
					this._V=new Visio.PageView(this.context, _createPropertyObjectPath(this.context, this, "View", false , false , false));
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
		Page.prototype.set=function (properties, options) {

			this._recursivelySet(properties, options, [] , ["view"] , [
				"allShapes" ,
				"comments" ,
				"shapes"
			]);
		};

		Page.prototype.update=function (properties) {

			this._recursivelyUpdate(properties);
		};

		Page.prototype.activate=function () {

			_createMethodAction(this.context, this, "Activate", 1 , [], false);
		};

		Page.prototype._handleResult=function (value) {
			_super.prototype._handleResult.call(this, value);
			if (_isNullOrUndefined(value))
				return;
			var obj=value;
			_fixObjectPathIfNecessary(this, obj);

			if (!_isUndefined(obj["Height"])) {
				this._H=obj["Height"];

			}
			if (!_isUndefined(obj["Index"])) {
				this._I=obj["Index"];

			}
			if (!_isUndefined(obj["IsBackground"])) {
				this._Is=obj["IsBackground"];

			}
			if (!_isUndefined(obj["Name"])) {
				this._N=obj["Name"];

			}
			if (!_isUndefined(obj["Width"])) {
				this._W=obj["Width"];

			}
			_handleNavigationPropertyResults(this, obj, ["allShapes", "AllShapes", "comments", "Comments", "shapes", "Shapes", "view", "View"]);
		};
		Page.prototype.load=function (option) {
			return _load(this, option);
		};
		Page.prototype.retrieve=function (option) {
			return _retrieve(this, option);
		};

		Page.prototype._handleRetrieveResult=function (value, result) {
			_super.prototype._handleRetrieveResult.call(this, value, result);

			_processRetrieveResult(this, value, result);
		};
		Page.prototype.toJSON=function () {
			return _toJson(this,
			 {
				"height": this._H,
				"index": this._I,
				"isBackground": this._Is,
				"name": this._N,
				"width": this._W,
			},
			 {
				"allShapes": this._A,
				"comments": this._C,
				"shapes": this._S,
				"view": this._V,
			});
		};

		Page.prototype.ensureUnchanged=function (data) {
			_createEnsureUnchangedAction(this.context, this, data);
			return;
		};
		return Page;
	}(OfficeExtension.ClientObject));
	Visio.Page=Page;

	var _typePageView="PageView";

	var PageView=(function (_super) {
		__extends(PageView, _super);
		function PageView() {
			return _super !==null && _super.apply(this, arguments) || this;
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

				this._Z=value;
				_createSetPropertyAction(this.context, this, "Zoom", value);
			},
			enumerable: true,
			configurable: true
		});
		PageView.prototype.set=function (properties, options) {

			this._recursivelySet(properties, options, ["zoom"] , [] , []);
		};

		PageView.prototype.update=function (properties) {

			this._recursivelyUpdate(properties);
		};

		PageView.prototype.centerViewportOnShape=function (ShapeId) {

			_createMethodAction(this.context, this, "CenterViewportOnShape", 1 , [ShapeId], false);
		};

		PageView.prototype.fitToWindow=function () {

			_createMethodAction(this.context, this, "FitToWindow", 1 , [], false);
		};

		PageView.prototype.getPosition=function () {

			var action=_createMethodAction(this.context, this, "GetPosition", 1 , [], false);
			var ret=new OfficeExtension.ClientResult();
			_addActionResultHandler(this, action, ret);
			return ret;
		};

		PageView.prototype.getSelection=function () {

			return new Visio.Selection(this.context, _createMethodObjectPath(this.context, this, "GetSelection", 1 , [], false , false , null , false));
		};

		PageView.prototype.isShapeInViewport=function (Shape) {

			var action=_createMethodAction(this.context, this, "IsShapeInViewport", 1 , [Shape], false);
			var ret=new OfficeExtension.ClientResult();
			_addActionResultHandler(this, action, ret);
			return ret;
		};

		PageView.prototype.setPosition=function (Position) {

			_createMethodAction(this.context, this, "SetPosition", 1 , [Position], false);
		};

		PageView.prototype._handleResult=function (value) {
			_super.prototype._handleResult.call(this, value);
			if (_isNullOrUndefined(value))
				return;
			var obj=value;
			_fixObjectPathIfNecessary(this, obj);

			if (!_isUndefined(obj["Zoom"])) {
				this._Z=obj["Zoom"];

			}
		};
		PageView.prototype.load=function (option) {
			return _load(this, option);
		};
		PageView.prototype.retrieve=function (option) {
			return _retrieve(this, option);
		};

		PageView.prototype._handleRetrieveResult=function (value, result) {
			_super.prototype._handleRetrieveResult.call(this, value, result);

			_processRetrieveResult(this, value, result);
		};
		PageView.prototype.toJSON=function () {
			return _toJson(this,
			 {
				"zoom": this._Z,
			},
			 {});
		};

		PageView.prototype.ensureUnchanged=function (data) {
			_createEnsureUnchangedAction(this.context, this, data);
			return;
		};
		return PageView;
	}(OfficeExtension.ClientObject));
	Visio.PageView=PageView;

	var _typePageCollection="PageCollection";

	var PageCollection=(function (_super) {
		__extends(PageCollection, _super);
		function PageCollection() {
			return _super !==null && _super.apply(this, arguments) || this;
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

		PageCollection.prototype.getCount=function () {

			var action=_createMethodAction(this.context, this, "GetCount", 1 , [], false);
			var ret=new OfficeExtension.ClientResult();
			_addActionResultHandler(this, action, ret);
			return ret;
		};

		PageCollection.prototype.getItem=function (key) {

			return new Visio.Page(this.context, _createIndexerObjectPath(this.context, this, [key]));
		};

		PageCollection.prototype._handleResult=function (value) {
			_super.prototype._handleResult.call(this, value);
			if (_isNullOrUndefined(value))
				return;
			var obj=value;
			_fixObjectPathIfNecessary(this, obj);

			if (!_isNullOrUndefined(obj[OfficeExtension.Constants.items])) {
				this.m__items=[];
				var _data=obj[OfficeExtension.Constants.items];
				for (var i=0; i < _data.length; i++) {
					var _item=new Visio.Page(this.context, _createChildItemObjectPathUsingIndexerOrGetItemAt(true, this.context, this, _data[i], i));
					_item._handleResult(_data[i]);
					this.m__items.push(_item);
				}
			}
		};
		PageCollection.prototype.load=function (option) {
			return _load(this, option);
		};
		PageCollection.prototype.retrieve=function (option) {
			return _retrieve(this, option);
		};

		PageCollection.prototype._handleRetrieveResult=function (value, result) {
			var _this=this;
			_super.prototype._handleRetrieveResult.call(this, value, result);

			_processRetrieveResult(this, value, result, function (childItemData, index) { return new Visio.Page(_this.context, _createChildItemObjectPathUsingIndexerOrGetItemAt(true, _this.context, _this, childItemData, index)); });
		};
		PageCollection.prototype.toJSON=function () {
			return _toJson(this,
			 {},
			 {}, this.m__items);
		};
		return PageCollection;
	}(OfficeExtension.ClientObject));
	Visio.PageCollection=PageCollection;

	var _typeShapeCollection="ShapeCollection";

	var ShapeCollection=(function (_super) {
		__extends(ShapeCollection, _super);
		function ShapeCollection() {
			return _super !==null && _super.apply(this, arguments) || this;
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

		ShapeCollection.prototype.getCount=function () {

			var action=_createMethodAction(this.context, this, "GetCount", 1 , [], false);
			var ret=new OfficeExtension.ClientResult();
			_addActionResultHandler(this, action, ret);
			return ret;
		};

		ShapeCollection.prototype.getItem=function (key) {

			return new Visio.Shape(this.context, _createIndexerObjectPath(this.context, this, [key]));
		};

		ShapeCollection.prototype._handleResult=function (value) {
			_super.prototype._handleResult.call(this, value);
			if (_isNullOrUndefined(value))
				return;
			var obj=value;
			_fixObjectPathIfNecessary(this, obj);

			if (!_isNullOrUndefined(obj[OfficeExtension.Constants.items])) {
				this.m__items=[];
				var _data=obj[OfficeExtension.Constants.items];
				for (var i=0; i < _data.length; i++) {
					var _item=new Visio.Shape(this.context, _createChildItemObjectPathUsingIndexerOrGetItemAt(true, this.context, this, _data[i], i));
					_item._handleResult(_data[i]);
					this.m__items.push(_item);
				}
			}
		};
		ShapeCollection.prototype.load=function (option) {
			return _load(this, option);
		};
		ShapeCollection.prototype.retrieve=function (option) {
			return _retrieve(this, option);
		};

		ShapeCollection.prototype._handleRetrieveResult=function (value, result) {
			var _this=this;
			_super.prototype._handleRetrieveResult.call(this, value, result);

			_processRetrieveResult(this, value, result, function (childItemData, index) { return new Visio.Shape(_this.context, _createChildItemObjectPathUsingIndexerOrGetItemAt(true, _this.context, _this, childItemData, index)); });
		};
		ShapeCollection.prototype.toJSON=function () {
			return _toJson(this,
			 {},
			 {}, this.m__items);
		};
		return ShapeCollection;
	}(OfficeExtension.ClientObject));
	Visio.ShapeCollection=ShapeCollection;

	var _typeShape="Shape";

	var Shape=(function (_super) {
		__extends(Shape, _super);
		function Shape() {
			return _super !==null && _super.apply(this, arguments) || this;
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
				return ["name", "id", "text", "select"];
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Shape.prototype, "_scalarPropertyUpdateable", {
			get: function () {
				return [false, false, false, true];
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
					this._C=new Visio.CommentCollection(this.context, _createPropertyObjectPath(this.context, this, "Comments", true , false , false));
				}

				return this._C;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Shape.prototype, "hyperlinks", {

			get: function () {

				if (!this._H) {
					this._H=new Visio.HyperlinkCollection(this.context, _createPropertyObjectPath(this.context, this, "Hyperlinks", true , false , false));
				}

				return this._H;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Shape.prototype, "shapeDataItems", {

			get: function () {

				if (!this._Sh) {
					this._Sh=new Visio.ShapeDataItemCollection(this.context, _createPropertyObjectPath(this.context, this, "ShapeDataItems", true , false , false));
				}

				return this._Sh;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Shape.prototype, "subShapes", {

			get: function () {

				if (!this._Su) {
					this._Su=new Visio.ShapeCollection(this.context, _createPropertyObjectPath(this.context, this, "SubShapes", true , false , false));
				}

				return this._Su;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Shape.prototype, "view", {

			get: function () {

				if (!this._V) {
					this._V=new Visio.ShapeView(this.context, _createPropertyObjectPath(this.context, this, "View", false , false , false));
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

				this._S=value;
				_createSetPropertyAction(this.context, this, "Select", value);
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
		Shape.prototype.set=function (properties, options) {

			this._recursivelySet(properties, options, ["select"] , ["view"] , [
				"comments" ,
				"hyperlinks" ,
				"shapeDataItems" ,
				"subShapes"
			]);
		};

		Shape.prototype.update=function (properties) {

			this._recursivelyUpdate(properties);
		};

		Shape.prototype.getBounds=function () {

			var action=_createMethodAction(this.context, this, "GetBounds", 1 , [], false);
			var ret=new OfficeExtension.ClientResult();
			_addActionResultHandler(this, action, ret);
			return ret;
		};

		Shape.prototype._handleResult=function (value) {
			_super.prototype._handleResult.call(this, value);
			if (_isNullOrUndefined(value))
				return;
			var obj=value;
			_fixObjectPathIfNecessary(this, obj);

			if (!_isUndefined(obj["Id"])) {
				this._I=obj["Id"];

			}
			if (!_isUndefined(obj["Name"])) {
				this._N=obj["Name"];

			}
			if (!_isUndefined(obj["Select"])) {
				this._S=obj["Select"];

			}
			if (!_isUndefined(obj["Text"])) {
				this._T=obj["Text"];

			}
			_handleNavigationPropertyResults(this, obj, ["comments", "Comments", "hyperlinks", "Hyperlinks", "shapeDataItems", "ShapeDataItems", "subShapes", "SubShapes", "view", "View"]);
		};
		Shape.prototype.load=function (option) {
			return _load(this, option);
		};
		Shape.prototype.retrieve=function (option) {
			return _retrieve(this, option);
		};

		Shape.prototype._handleIdResult=function (value) {
			_super.prototype._handleIdResult.call(this, value);
			if (_isNullOrUndefined(value)) {
				return;
			}
			if (!_isUndefined(value["Id"])) {

				this._I=value["Id"];
			}
		};

		Shape.prototype._handleRetrieveResult=function (value, result) {
			_super.prototype._handleRetrieveResult.call(this, value, result);

			_processRetrieveResult(this, value, result);
		};
		Shape.prototype.toJSON=function () {
			return _toJson(this,
			 {
				"id": this._I,
				"name": this._N,
				"select": this._S,
				"text": this._T,
			},
			 {
				"comments": this._C,
				"hyperlinks": this._H,
				"shapeDataItems": this._Sh,
				"subShapes": this._Su,
				"view": this._V,
			});
		};

		Shape.prototype.ensureUnchanged=function (data) {
			_createEnsureUnchangedAction(this.context, this, data);
			return;
		};
		return Shape;
	}(OfficeExtension.ClientObject));
	Visio.Shape=Shape;

	var _typeShapeView="ShapeView";

	var ShapeView=(function (_super) {
		__extends(ShapeView, _super);
		function ShapeView() {
			return _super !==null && _super.apply(this, arguments) || this;
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

				this._H=value;
				_createSetPropertyAction(this.context, this, "Highlight", value);
			},
			enumerable: true,
			configurable: true
		});
		ShapeView.prototype.set=function (properties, options) {

			this._recursivelySet(properties, options, ["highlight"] , [] , []);
		};

		ShapeView.prototype.update=function (properties) {

			this._recursivelyUpdate(properties);
		};

		ShapeView.prototype.addOverlay=function (OverlayType, Content, OverlayHorizontalAlignment, OverlayVerticalAlignment, Width, Height) {

			var action=_createMethodAction(this.context, this, "AddOverlay", 1 , [OverlayType, Content, OverlayHorizontalAlignment, OverlayVerticalAlignment, Width, Height], false);
			var ret=new OfficeExtension.ClientResult();
			_addActionResultHandler(this, action, ret);
			return ret;
		};

		ShapeView.prototype.removeOverlay=function (OverlayId) {

			_createMethodAction(this.context, this, "RemoveOverlay", 1 , [OverlayId], false);
		};

		ShapeView.prototype._handleResult=function (value) {
			_super.prototype._handleResult.call(this, value);
			if (_isNullOrUndefined(value))
				return;
			var obj=value;
			_fixObjectPathIfNecessary(this, obj);

			if (!_isUndefined(obj["Highlight"])) {
				this._H=obj["Highlight"];

			}
		};
		ShapeView.prototype.load=function (option) {
			return _load(this, option);
		};
		ShapeView.prototype.retrieve=function (option) {
			return _retrieve(this, option);
		};

		ShapeView.prototype._handleRetrieveResult=function (value, result) {
			_super.prototype._handleRetrieveResult.call(this, value, result);

			_processRetrieveResult(this, value, result);
		};
		ShapeView.prototype.toJSON=function () {
			return _toJson(this,
			 {
				"highlight": this._H,
			},
			 {});
		};

		ShapeView.prototype.ensureUnchanged=function (data) {
			_createEnsureUnchangedAction(this.context, this, data);
			return;
		};
		return ShapeView;
	}(OfficeExtension.ClientObject));
	Visio.ShapeView=ShapeView;

	var _typeShapeDataItemCollection="ShapeDataItemCollection";

	var ShapeDataItemCollection=(function (_super) {
		__extends(ShapeDataItemCollection, _super);
		function ShapeDataItemCollection() {
			return _super !==null && _super.apply(this, arguments) || this;
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

		ShapeDataItemCollection.prototype.getCount=function () {

			var action=_createMethodAction(this.context, this, "GetCount", 1 , [], false);
			var ret=new OfficeExtension.ClientResult();
			_addActionResultHandler(this, action, ret);
			return ret;
		};

		ShapeDataItemCollection.prototype.getItem=function (key) {

			return new Visio.ShapeDataItem(this.context, _createIndexerObjectPath(this.context, this, [key]));
		};

		ShapeDataItemCollection.prototype._handleResult=function (value) {
			_super.prototype._handleResult.call(this, value);
			if (_isNullOrUndefined(value))
				return;
			var obj=value;
			_fixObjectPathIfNecessary(this, obj);

			if (!_isNullOrUndefined(obj[OfficeExtension.Constants.items])) {
				this.m__items=[];
				var _data=obj[OfficeExtension.Constants.items];
				for (var i=0; i < _data.length; i++) {
					var _item=new Visio.ShapeDataItem(this.context, _createChildItemObjectPathUsingIndexerOrGetItemAt(true, this.context, this, _data[i], i));
					_item._handleResult(_data[i]);
					this.m__items.push(_item);
				}
			}
		};
		ShapeDataItemCollection.prototype.load=function (option) {
			return _load(this, option);
		};
		ShapeDataItemCollection.prototype.retrieve=function (option) {
			return _retrieve(this, option);
		};

		ShapeDataItemCollection.prototype._handleRetrieveResult=function (value, result) {
			var _this=this;
			_super.prototype._handleRetrieveResult.call(this, value, result);

			_processRetrieveResult(this, value, result, function (childItemData, index) { return new Visio.ShapeDataItem(_this.context, _createChildItemObjectPathUsingIndexerOrGetItemAt(true, _this.context, _this, childItemData, index)); });
		};
		ShapeDataItemCollection.prototype.toJSON=function () {
			return _toJson(this,
			 {},
			 {}, this.m__items);
		};
		return ShapeDataItemCollection;
	}(OfficeExtension.ClientObject));
	Visio.ShapeDataItemCollection=ShapeDataItemCollection;

	var _typeShapeDataItem="ShapeDataItem";

	var ShapeDataItem=(function (_super) {
		__extends(ShapeDataItem, _super);
		function ShapeDataItem() {
			return _super !==null && _super.apply(this, arguments) || this;
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

		ShapeDataItem.prototype._handleResult=function (value) {
			_super.prototype._handleResult.call(this, value);
			if (_isNullOrUndefined(value))
				return;
			var obj=value;
			_fixObjectPathIfNecessary(this, obj);

			if (!_isUndefined(obj["Format"])) {
				this._F=obj["Format"];

			}
			if (!_isUndefined(obj["FormattedValue"])) {
				this._Fo=obj["FormattedValue"];

			}
			if (!_isUndefined(obj["Label"])) {
				this._L=obj["Label"];

			}
			if (!_isUndefined(obj["Value"])) {
				this._V=obj["Value"];

			}
		};
		ShapeDataItem.prototype.load=function (option) {
			return _load(this, option);
		};
		ShapeDataItem.prototype.retrieve=function (option) {
			return _retrieve(this, option);
		};

		ShapeDataItem.prototype._handleRetrieveResult=function (value, result) {
			_super.prototype._handleRetrieveResult.call(this, value, result);

			_processRetrieveResult(this, value, result);
		};
		ShapeDataItem.prototype.toJSON=function () {
			return _toJson(this,
			 {
				"format": this._F,
				"formattedValue": this._Fo,
				"label": this._L,
				"value": this._V,
			},
			 {});
		};

		ShapeDataItem.prototype.ensureUnchanged=function (data) {
			_createEnsureUnchangedAction(this.context, this, data);
			return;
		};
		return ShapeDataItem;
	}(OfficeExtension.ClientObject));
	Visio.ShapeDataItem=ShapeDataItem;

	var _typeHyperlinkCollection="HyperlinkCollection";

	var HyperlinkCollection=(function (_super) {
		__extends(HyperlinkCollection, _super);
		function HyperlinkCollection() {
			return _super !==null && _super.apply(this, arguments) || this;
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

		HyperlinkCollection.prototype.getCount=function () {

			var action=_createMethodAction(this.context, this, "GetCount", 1 , [], false);
			var ret=new OfficeExtension.ClientResult();
			_addActionResultHandler(this, action, ret);
			return ret;
		};

		HyperlinkCollection.prototype.getItem=function (Key) {

			return new Visio.Hyperlink(this.context, _createIndexerObjectPath(this.context, this, [Key]));
		};

		HyperlinkCollection.prototype._handleResult=function (value) {
			_super.prototype._handleResult.call(this, value);
			if (_isNullOrUndefined(value))
				return;
			var obj=value;
			_fixObjectPathIfNecessary(this, obj);

			if (!_isNullOrUndefined(obj[OfficeExtension.Constants.items])) {
				this.m__items=[];
				var _data=obj[OfficeExtension.Constants.items];
				for (var i=0; i < _data.length; i++) {
					var _item=new Visio.Hyperlink(this.context, _createChildItemObjectPathUsingIndexerOrGetItemAt(true, this.context, this, _data[i], i));
					_item._handleResult(_data[i]);
					this.m__items.push(_item);
				}
			}
		};
		HyperlinkCollection.prototype.load=function (option) {
			return _load(this, option);
		};
		HyperlinkCollection.prototype.retrieve=function (option) {
			return _retrieve(this, option);
		};

		HyperlinkCollection.prototype._handleRetrieveResult=function (value, result) {
			var _this=this;
			_super.prototype._handleRetrieveResult.call(this, value, result);

			_processRetrieveResult(this, value, result, function (childItemData, index) { return new Visio.Hyperlink(_this.context, _createChildItemObjectPathUsingIndexerOrGetItemAt(true, _this.context, _this, childItemData, index)); });
		};
		HyperlinkCollection.prototype.toJSON=function () {
			return _toJson(this,
			 {},
			 {}, this.m__items);
		};
		return HyperlinkCollection;
	}(OfficeExtension.ClientObject));
	Visio.HyperlinkCollection=HyperlinkCollection;

	var _typeHyperlink="Hyperlink";

	var Hyperlink=(function (_super) {
		__extends(Hyperlink, _super);
		function Hyperlink() {
			return _super !==null && _super.apply(this, arguments) || this;
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

		Hyperlink.prototype._handleResult=function (value) {
			_super.prototype._handleResult.call(this, value);
			if (_isNullOrUndefined(value))
				return;
			var obj=value;
			_fixObjectPathIfNecessary(this, obj);

			if (!_isUndefined(obj["Address"])) {
				this._A=obj["Address"];

			}
			if (!_isUndefined(obj["Description"])) {
				this._D=obj["Description"];

			}
			if (!_isUndefined(obj["ExtraInfo"])) {
				this._E=obj["ExtraInfo"];

			}
			if (!_isUndefined(obj["SubAddress"])) {
				this._S=obj["SubAddress"];

			}
		};
		Hyperlink.prototype.load=function (option) {
			return _load(this, option);
		};
		Hyperlink.prototype.retrieve=function (option) {
			return _retrieve(this, option);
		};

		Hyperlink.prototype._handleRetrieveResult=function (value, result) {
			_super.prototype._handleRetrieveResult.call(this, value, result);

			_processRetrieveResult(this, value, result);
		};
		Hyperlink.prototype.toJSON=function () {
			return _toJson(this,
			 {
				"address": this._A,
				"description": this._D,
				"extraInfo": this._E,
				"subAddress": this._S,
			},
			 {});
		};

		Hyperlink.prototype.ensureUnchanged=function (data) {
			_createEnsureUnchangedAction(this.context, this, data);
			return;
		};
		return Hyperlink;
	}(OfficeExtension.ClientObject));
	Visio.Hyperlink=Hyperlink;

	var _typeCommentCollection="CommentCollection";

	var CommentCollection=(function (_super) {
		__extends(CommentCollection, _super);
		function CommentCollection() {
			return _super !==null && _super.apply(this, arguments) || this;
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

		CommentCollection.prototype.getCount=function () {

			var action=_createMethodAction(this.context, this, "GetCount", 1 , [], false);
			var ret=new OfficeExtension.ClientResult();
			_addActionResultHandler(this, action, ret);
			return ret;
		};

		CommentCollection.prototype.getItem=function (key) {

			return new Visio.Comment(this.context, _createIndexerObjectPath(this.context, this, [key]));
		};

		CommentCollection.prototype._handleResult=function (value) {
			_super.prototype._handleResult.call(this, value);
			if (_isNullOrUndefined(value))
				return;
			var obj=value;
			_fixObjectPathIfNecessary(this, obj);

			if (!_isNullOrUndefined(obj[OfficeExtension.Constants.items])) {
				this.m__items=[];
				var _data=obj[OfficeExtension.Constants.items];
				for (var i=0; i < _data.length; i++) {
					var _item=new Visio.Comment(this.context, _createChildItemObjectPathUsingIndexerOrGetItemAt(true, this.context, this, _data[i], i));
					_item._handleResult(_data[i]);
					this.m__items.push(_item);
				}
			}
		};
		CommentCollection.prototype.load=function (option) {
			return _load(this, option);
		};
		CommentCollection.prototype.retrieve=function (option) {
			return _retrieve(this, option);
		};

		CommentCollection.prototype._handleRetrieveResult=function (value, result) {
			var _this=this;
			_super.prototype._handleRetrieveResult.call(this, value, result);

			_processRetrieveResult(this, value, result, function (childItemData, index) { return new Visio.Comment(_this.context, _createChildItemObjectPathUsingIndexerOrGetItemAt(true, _this.context, _this, childItemData, index)); });
		};
		CommentCollection.prototype.toJSON=function () {
			return _toJson(this,
			 {},
			 {}, this.m__items);
		};
		return CommentCollection;
	}(OfficeExtension.ClientObject));
	Visio.CommentCollection=CommentCollection;

	var _typeComment="Comment";

	var Comment=(function (_super) {
		__extends(Comment, _super);
		function Comment() {
			return _super !==null && _super.apply(this, arguments) || this;
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

				this._A=value;
				_createSetPropertyAction(this.context, this, "Author", value);
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

				this._D=value;
				_createSetPropertyAction(this.context, this, "Date", value);
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

				this._T=value;
				_createSetPropertyAction(this.context, this, "Text", value);
			},
			enumerable: true,
			configurable: true
		});
		Comment.prototype.set=function (properties, options) {

			this._recursivelySet(properties, options, ["author", "text", "date"] , [] , []);
		};

		Comment.prototype.update=function (properties) {

			this._recursivelyUpdate(properties);
		};

		Comment.prototype._handleResult=function (value) {
			_super.prototype._handleResult.call(this, value);
			if (_isNullOrUndefined(value))
				return;
			var obj=value;
			_fixObjectPathIfNecessary(this, obj);

			if (!_isUndefined(obj["Author"])) {
				this._A=obj["Author"];

			}
			if (!_isUndefined(obj["Date"])) {
				this._D=obj["Date"];

			}
			if (!_isUndefined(obj["Text"])) {
				this._T=obj["Text"];

			}
		};
		Comment.prototype.load=function (option) {
			return _load(this, option);
		};
		Comment.prototype.retrieve=function (option) {
			return _retrieve(this, option);
		};

		Comment.prototype._handleRetrieveResult=function (value, result) {
			_super.prototype._handleRetrieveResult.call(this, value, result);

			_processRetrieveResult(this, value, result);
		};
		Comment.prototype.toJSON=function () {
			return _toJson(this,
			 {
				"author": this._A,
				"date": this._D,
				"text": this._T,
			},
			 {});
		};

		Comment.prototype.ensureUnchanged=function (data) {
			_createEnsureUnchangedAction(this.context, this, data);
			return;
		};
		return Comment;
	}(OfficeExtension.ClientObject));
	Visio.Comment=Comment;

	var _typeSelection="Selection";

	var Selection=(function (_super) {
		__extends(Selection, _super);
		function Selection() {
			return _super !==null && _super.apply(this, arguments) || this;
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
					this._S=new Visio.ShapeCollection(this.context, _createPropertyObjectPath(this.context, this, "Shapes", true , false , false));
				}

				return this._S;
			},
			enumerable: true,
			configurable: true
		});

		Selection.prototype._handleResult=function (value) {
			_super.prototype._handleResult.call(this, value);
			if (_isNullOrUndefined(value))
				return;
			var obj=value;
			_fixObjectPathIfNecessary(this, obj);

			_handleNavigationPropertyResults(this, obj, ["shapes", "Shapes"]);
		};
		Selection.prototype.load=function (option) {
			return _load(this, option);
		};
		Selection.prototype.retrieve=function (option) {
			return _retrieve(this, option);
		};

		Selection.prototype._handleRetrieveResult=function (value, result) {
			_super.prototype._handleRetrieveResult.call(this, value, result);

			_processRetrieveResult(this, value, result);
		};
		Selection.prototype.toJSON=function () {
			return _toJson(this,
			 {},
			 {
				"shapes": this._S,
			});
		};
		return Selection;
	}(OfficeExtension.ClientObject));
	Visio.Selection=Selection;

	var OverlayHorizontalAlignment;
	(function (OverlayHorizontalAlignment) {

		OverlayHorizontalAlignment.left="Left";

		OverlayHorizontalAlignment.center="Center";

		OverlayHorizontalAlignment.right="Right";
	})(OverlayHorizontalAlignment=Visio.OverlayHorizontalAlignment || (Visio.OverlayHorizontalAlignment={}));

	var OverlayVerticalAlignment;
	(function (OverlayVerticalAlignment) {

		OverlayVerticalAlignment.top="Top";

		OverlayVerticalAlignment.middle="Middle";

		OverlayVerticalAlignment.bottom="Bottom";
	})(OverlayVerticalAlignment=Visio.OverlayVerticalAlignment || (Visio.OverlayVerticalAlignment={}));

	var OverlayType;
	(function (OverlayType) {

		OverlayType.text="Text";

		OverlayType.image="Image";
	})(OverlayType=Visio.OverlayType || (Visio.OverlayType={}));

	var ToolBarType;
	(function (ToolBarType) {

		ToolBarType.commandBar="CommandBar";

		ToolBarType.pageNavigationBar="PageNavigationBar";

		ToolBarType.statusBar="StatusBar";
	})(ToolBarType=Visio.ToolBarType || (Visio.ToolBarType={}));

	var ErrorCodes;
	(function (ErrorCodes) {
		ErrorCodes.accessDenied="AccessDenied";
		ErrorCodes.generalException="GeneralException";
		ErrorCodes.invalidArgument="InvalidArgument";
		ErrorCodes.itemNotFound="ItemNotFound";
		ErrorCodes.notImplemented="NotImplemented";
		ErrorCodes.unsupportedOperation="UnsupportedOperation";

	})(ErrorCodes=Visio.ErrorCodes || (Visio.ErrorCodes={}));
})(Visio || (Visio={}));

var Visio;
(function (Visio) {

	var RequestContext=(function (_super) {
		__extends(RequestContext, _super);
		function RequestContext(url) {
			var _this=_super.call(this, url) || this;
			_this.m_document=new Visio.Document(_this, OfficeExtension.ObjectPathFactory.createGlobalObjectObjectPath(_this));
			_this._rootObject=_this.m_document;
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
	}(OfficeExtension.ClientRequestContext));
	Visio.RequestContext=RequestContext;
	function run(arg1, arg2, arg3) {
		return OfficeExtension.ClientRequestContext._runBatch("Visio.run", arguments, function (requestInfo) {
			var ret=new Visio.RequestContext(requestInfo);
			return ret;
		});
	}
	Visio.run=run;
})(Visio || (Visio={}));

OfficeExtension.Utility._doApiNotSupportedCheck=true;



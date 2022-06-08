/* Excel iOS specific API library */
/* Version: 15.0.4420.1017 Build Time: 05/04/2015 */
/*
	Copyright (c) Microsoft Corporation.  All rights reserved.
*/

/*
	Your use of this file is governed by the Microsoft Services Agreement http://go.microsoft.com/fwlink/?LinkId=266419.
*/

OSF.OUtil.augmentList(Microsoft.Office.WebExtension.FilterType, {
	OnlyVisible: "onlyVisible"
});
var OSF=OSF || {};
var OSFWebkit;
(function (OSFWebkit) {
	var WebkitSafeArray=(function () {
		function WebkitSafeArray(data) {
			this.data=data;
			this.safeArrayFlag=this.isSafeArray(data);
		}
		WebkitSafeArray.prototype.dimensions=function () {
			var dimensions=0;
			if (this.safeArrayFlag) {
				dimensions=this.data[0][0];
			} else if (this.isArray()) {
				dimensions=2;
			}
			return dimensions;
		};
		WebkitSafeArray.prototype.getItem=function () {
			var array=[];
			var element=null;
			if (this.safeArrayFlag) {
				array=this.toArray();
			} else {
				array=this.data;
			}
			element=array;
			for (var i=0; i < arguments.length; i++) {
				element=element[arguments[i]];
			}
			return element;
		};
		WebkitSafeArray.prototype.lbound=function (dimension) {
			return 0;
		};
		WebkitSafeArray.prototype.ubound=function (dimension) {
			var ubound=0;
			if (this.safeArrayFlag) {
				ubound=this.data[0][dimension];
			} else if (this.isArray()) {
				if (dimension==1) {
					return this.data.length;
				} else if (dimension==2) {
					if (OSF.OUtil.isArray(this.data[0])) {
						return this.data[0].length;
					} else if (this.data[0] !=null) {
						return 1;
					}
				}
			}
			return ubound;
		};
		WebkitSafeArray.prototype.toArray=function () {
			if (this.isArray()==false) {
				return this.data;
			}
			var arr=[];
			var startingIndex=this.safeArrayFlag ? 1 : 0;
			for (var i=startingIndex; i < this.data.length; i++) {
				var element=this.data[i];
				if (this.isSafeArray(element)) {
					arr.push(new WebkitSafeArray(element));
				} else {
					arr.push(element);
				}
			}
			return arr;
		};
		WebkitSafeArray.prototype.isArray=function () {
			return OSF.OUtil.isArray(this.data);
		};
		WebkitSafeArray.prototype.isSafeArray=function (obj) {
			var isSafeArray=false;
			if (OSF.OUtil.isArray(obj) && OSF.OUtil.isArray(obj[0])) {
				var bounds=obj[0];
				var dimensions=bounds[0];
				if (bounds.length !=dimensions+1) {
					return false;
				}
				var expectedArraySize=1;
				for (var i=1; i < bounds.length; i++) {
					var dimension=bounds[i];
					if (isFinite(dimension)==false) {
						return false;
					}
					expectedArraySize=expectedArraySize * dimension;
				}
				expectedArraySize++;
				isSafeArray=(expectedArraySize==obj.length);
			}
			return isSafeArray;
		};
		return WebkitSafeArray;
	})();
	OSFWebkit.WebkitSafeArray=WebkitSafeArray;
})(OSFWebkit || (OSFWebkit={}));
var OSFWebkit;
(function (OSFWebkit) {
	(function (ScriptMessaging) {
		var scriptMessenger=null;
		function agaveHostCallback(callbackId, params) {
			scriptMessenger.agaveHostCallback(callbackId, params);
		}
		ScriptMessaging.agaveHostCallback=agaveHostCallback;
		function agaveHostEventCallback(callbackId, params) {
			scriptMessenger.agaveHostEventCallback(callbackId, params);
		}
		ScriptMessaging.agaveHostEventCallback=agaveHostEventCallback;
		function GetScriptMessenger() {
			if (scriptMessenger==null) {
				scriptMessenger=new WebkitScriptMessaging("OSF.ScriptMessaging.agaveHostCallback", "OSF.ScriptMessaging.agaveHostEventCallback");
			}
			return scriptMessenger;
		}
		ScriptMessaging.GetScriptMessenger=GetScriptMessenger;
		var EventHandlerCallback=(function () {
			function EventHandlerCallback(id, targetId, handler) {
				this.id=id;
				this.targetId=targetId;
				this.handler=handler;
			}
			return EventHandlerCallback;
		})();
		var WebkitScriptMessaging=(function () {
			function WebkitScriptMessaging(methodCallbackName, eventCallbackName) {
				this.callingIndex=0;
				this.callbackList={};
				this.eventHandlerList={};
				this.asyncMethodCallbackFunctionName=methodCallbackName;
				this.eventCallbackFunctionName=eventCallbackName;
				this.conversationId=WebkitScriptMessaging.getCurrentTimeMS().toString();
			}
			WebkitScriptMessaging.prototype.invokeMethod=function (handlerName, methodId, params, callback) {
				var messagingArgs={};
				this.postWebkitMessage(messagingArgs, handlerName, methodId, params, callback);
			};
			WebkitScriptMessaging.prototype.registerEvent=function (handlerName, methodId, dispId, targetId, handler, callback) {
				var messagingArgs={
					eventCallbackFunction: this.eventCallbackFunctionName
				};
				var hostArgs={
					id: dispId,
					targetId: targetId
				};
				var correlationId=this.postWebkitMessage(messagingArgs, handlerName, methodId, hostArgs, callback);
				this.eventHandlerList[correlationId]=new EventHandlerCallback(dispId, targetId, handler);
			};
			WebkitScriptMessaging.prototype.unregisterEvent=function (handlerName, methodId, dispId, targetId, callback) {
				var hostArgs={
					id: dispId,
					targetId: targetId
				};
				for (var key in this.eventHandlerList) {
					if (this.eventHandlerList.hasOwnProperty(key)) {
						var eventCallback=this.eventHandlerList[key];
						if (eventCallback.id==dispId && eventCallback.targetId==targetId) {
							delete this.eventHandlerList[key];
						}
					}
				}
				this.invokeMethod(handlerName, methodId, hostArgs, callback);
			};
			WebkitScriptMessaging.prototype.agaveHostCallback=function (callbackId, params) {
				var callbackFunction=this.callbackList[callbackId];
				if (callbackFunction) {
					callbackFunction(params);
					delete this.callbackList[callbackId];
				}
			};
			WebkitScriptMessaging.prototype.agaveHostEventCallback=function (callbackId, params) {
				var eventCallback=this.eventHandlerList[callbackId];
				if (eventCallback) {
					eventCallback.handler(params);
				}
			};
			WebkitScriptMessaging.prototype.postWebkitMessage=function (messagingArgs, handlerName, methodId, params, callback) {
				var correlationId=this.generateCorrelationId();
				this.callbackList[correlationId]=callback;
				messagingArgs.methodId=methodId;
				messagingArgs.params=params;
				messagingArgs.callbackId=correlationId;
				messagingArgs.callbackFunction=this.asyncMethodCallbackFunctionName;
				var invokePostMessage=function () {
					window.webkit.messageHandlers[handlerName].postMessage(JSON.stringify(messagingArgs));
				};
				var currentTimestamp=WebkitScriptMessaging.getCurrentTimeMS();
				if (this.lastMessageTimestamp==null || (currentTimestamp - this.lastMessageTimestamp >=WebkitScriptMessaging.MESSAGE_TIME_DELTA)) {
					invokePostMessage();
					this.lastMessageTimestamp=currentTimestamp;
				} else {
					this.lastMessageTimestamp+=WebkitScriptMessaging.MESSAGE_TIME_DELTA;
					setTimeout(function () {
						invokePostMessage();
					}, this.lastMessageTimestamp - currentTimestamp);
				}
				return correlationId;
			};
			WebkitScriptMessaging.prototype.generateCorrelationId=function () {
++this.callingIndex;
				return this.conversationId+this.callingIndex;
			};
			WebkitScriptMessaging.getCurrentTimeMS=function () {
				return (new Date).getTime();
			};
			WebkitScriptMessaging.MESSAGE_TIME_DELTA=10;
			return WebkitScriptMessaging;
		})();
		ScriptMessaging.WebkitScriptMessaging=WebkitScriptMessaging;
	})(OSFWebkit.ScriptMessaging || (OSFWebkit.ScriptMessaging={}));
	var ScriptMessaging=OSFWebkit.ScriptMessaging;
})(OSFWebkit || (OSFWebkit={}));
OSF.ScriptMessaging=OSFWebkit.ScriptMessaging;
var OSFWebkit;
(function (OSFWebkit) {
	OSFWebkit.MessageHandlerName="Agave";
	(function (AppContextProperties) {
		AppContextProperties[AppContextProperties["Settings"]=0]="Settings";
		AppContextProperties[AppContextProperties["SolutionReferenceId"]=1]="SolutionReferenceId";
		AppContextProperties[AppContextProperties["AppType"]=2]="AppType";
		AppContextProperties[AppContextProperties["MajorVersion"]=3]="MajorVersion";
		AppContextProperties[AppContextProperties["MinorVersion"]=4]="MinorVersion";
		AppContextProperties[AppContextProperties["RevisionVersion"]=5]="RevisionVersion";
		AppContextProperties[AppContextProperties["APIVersionSequence"]=6]="APIVersionSequence";
		AppContextProperties[AppContextProperties["AppCapabilities"]=7]="AppCapabilities";
		AppContextProperties[AppContextProperties["APPUILocale"]=8]="APPUILocale";
		AppContextProperties[AppContextProperties["AppDataLocale"]=9]="AppDataLocale";
		AppContextProperties[AppContextProperties["BindingCount"]=10]="BindingCount";
		AppContextProperties[AppContextProperties["DocumentUrl"]=11]="DocumentUrl";
		AppContextProperties[AppContextProperties["ActivationMode"]=12]="ActivationMode";
		AppContextProperties[AppContextProperties["ControlIntegrationLevel"]=13]="ControlIntegrationLevel";
		AppContextProperties[AppContextProperties["SolutionToken"]=14]="SolutionToken";
		AppContextProperties[AppContextProperties["APISetVersion"]=15]="APISetVersion";
		AppContextProperties[AppContextProperties["CorrelationId"]=16]="CorrelationId";
		AppContextProperties[AppContextProperties["InstanceId"]=17]="InstanceId";
	})(OSFWebkit.AppContextProperties || (OSFWebkit.AppContextProperties={}));
	var AppContextProperties=OSFWebkit.AppContextProperties;
	(function (MethodId) {
		MethodId[MethodId["Execute"]=1]="Execute";
		MethodId[MethodId["RegisterEvent"]=2]="RegisterEvent";
		MethodId[MethodId["UnregisterEvent"]=3]="UnregisterEvent";
		MethodId[MethodId["WriteSettings"]=4]="WriteSettings";
		MethodId[MethodId["GetContext"]=5]="GetContext";
	})(OSFWebkit.MethodId || (OSFWebkit.MethodId={}));
	var MethodId=OSFWebkit.MethodId;
	var WebkitHostController=(function () {
		function WebkitHostController(hostScriptProxy) {
			this.hostScriptProxy=hostScriptProxy;
		}
		WebkitHostController.prototype.execute=function (id, params, callback) {
			var args=params;
			if (args==null) {
				args=[];
			}
			var hostParams={
				id: id,
				apiArgs: args
			};
			var agaveResponseCallback=function (payload) {
				var safeArraySource=payload;
				if (OSF.OUtil.isArray(payload) && payload.length >=2) {
					var hrStatus=payload[0];
					safeArraySource=payload[1];
				}
				if (callback) {
					callback(new OSFWebkit.WebkitSafeArray(safeArraySource));
				}
			};
			this.hostScriptProxy.invokeMethod(OSF.Webkit.MessageHandlerName, OSF.Webkit.MethodId.Execute, hostParams, agaveResponseCallback);
		};
		WebkitHostController.prototype.registerEvent=function (id, targetId, handler, callback) {
			var agaveEventHandlerCallback=function (payload) {
				var safeArraySource=payload;
				var eventId=0;
				if (OSF.OUtil.isArray(payload) && payload.length >=2) {
					safeArraySource=payload[0];
					eventId=payload[1];
				}
				if (handler) {
					handler(eventId, new OSFWebkit.WebkitSafeArray(safeArraySource));
				}
			};
			var agaveResponseCallback=function (payload) {
				if (callback) {
					callback(new OSFWebkit.WebkitSafeArray(payload));
				}
			};
			this.hostScriptProxy.registerEvent(OSF.Webkit.MessageHandlerName, OSF.Webkit.MethodId.RegisterEvent, id, targetId, agaveEventHandlerCallback, agaveResponseCallback);
		};
		WebkitHostController.prototype.unregisterEvent=function (id, targetId, callback) {
			var agaveResponseCallback=function (response) {
				callback(new OSFWebkit.WebkitSafeArray(response));
			};
			this.hostScriptProxy.unregisterEvent(OSF.Webkit.MessageHandlerName, OSF.Webkit.MethodId.UnregisterEvent, id, targetId, agaveResponseCallback);
		};
		return WebkitHostController;
	})();
	OSFWebkit.WebkitHostController=WebkitHostController;
})(OSFWebkit || (OSFWebkit={}));
OSF.Webkit=OSFWebkit;
OSF.ClientHostController=new OSFWebkit.WebkitHostController(OSF.ScriptMessaging.GetScriptMessenger());
OSF.ClientMode={
	ReadWrite: 0,
	ReadOnly: 1
}
OSF.DDA.RichInitializationReason={
	1: Microsoft.Office.WebExtension.InitializationReason.Inserted,
	2: Microsoft.Office.WebExtension.InitializationReason.DocumentOpened
};
Microsoft.Office.WebExtension.FileType={
	Text: "text",
	Compressed: "compressed"
};
OSF.DDA.File=function OSF_DDA_File(handle, fileSize, sliceSize) {
	OSF.OUtil.defineEnumerableProperties(this, {
		"size": {
			value: fileSize
		},
		"sliceCount": {
			value: Math.ceil(fileSize / sliceSize)
		}
	});
	var privateState={};
	privateState[OSF.DDA.FileProperties.Handle]=handle;
	privateState[OSF.DDA.FileProperties.SliceSize]=sliceSize;
	var am=OSF.DDA.AsyncMethodNames;
	OSF.DDA.DispIdHost.addAsyncMethods(
		this, [
			am.GetDocumentCopyChunkAsync,
			am.ReleaseDocumentCopyAsync
		],
		privateState
	);
}
OSF.DDA.FileSliceOffset="fileSliceoffset";
OSF.DDA.CustomXmlParts=function OSF_DDA_CustomXmlParts() {
	this._eventDispatches=[];
	var am=OSF.DDA.AsyncMethodNames;
	OSF.DDA.DispIdHost.addAsyncMethods(this, [
		am.AddDataPartAsync,
		am.GetDataPartByIdAsync,
		am.GetDataPartsByNameSpaceAsync
	]);
};
OSF.DDA.CustomXmlPart=function OSF_DDA_CustomXmlPart(customXmlParts, id, builtIn) {
	OSF.OUtil.defineEnumerableProperties(this, {
		"builtIn": {
			value: builtIn
		},
		"id": {
			value: id
		},
		"namespaceManager": {
			value: new OSF.DDA.CustomXmlPrefixMappings(id)
		}
	});
	var am=OSF.DDA.AsyncMethodNames;
	OSF.DDA.DispIdHost.addAsyncMethods(this, [
		am.DeleteDataPartAsync,
		am.GetPartNodesAsync,
		am.GetPartXmlAsync
	]);
	var customXmlPartEventDispatches=customXmlParts._eventDispatches;
	var dispatch=customXmlPartEventDispatches[id];
	if (!dispatch) {
		var et=Microsoft.Office.WebExtension.EventType;
		dispatch=new OSF.EventDispatch([
			et.DataNodeDeleted,
			et.DataNodeInserted,
			et.DataNodeReplaced
		]);
		customXmlPartEventDispatches[id]=dispatch;
	}
	OSF.DDA.DispIdHost.addEventSupport(this, dispatch);
};
OSF.DDA.CustomXmlPrefixMappings=function OSF_DDA_CustomXmlPrefixMappings(partId) {
	var am=OSF.DDA.AsyncMethodNames;
	OSF.DDA.DispIdHost.addAsyncMethods(
		this,
		[
			am.AddDataPartNamespaceAsync,
			am.GetDataPartNamespaceAsync,
			am.GetDataPartPrefixAsync
		],
		partId
	);
};
OSF.DDA.CustomXmlNode=function OSF_DDA_CustomXmlNode(handle, nodeType, ns, baseName) {
	OSF.OUtil.defineEnumerableProperties(this, {
		"baseName": {
			value: baseName
		},
		"namespaceUri": {
			value: ns
		},
		"nodeType": {
			value: nodeType
		}
	});
	var am=OSF.DDA.AsyncMethodNames;
	OSF.DDA.DispIdHost.addAsyncMethods(
		this,
		[
			am.GetRelativeNodesAsync,
			am.GetNodeValueAsync,
			am.GetNodeXmlAsync,
			am.SetNodeValueAsync,
			am.SetNodeXmlAsync
		],
		handle
	);
};
OSF.DDA.NodeInsertedEventArgs=function OSF_DDA_NodeInsertedEventArgs(newNode, inUndoRedo) {
	OSF.OUtil.defineEnumerableProperties(this, {
		"type": {
			value: Microsoft.Office.WebExtension.EventType.DataNodeInserted
		},
		"newNode": {
			value: newNode
		},
		"inUndoRedo": {
			value: inUndoRedo
		}
	});
};
OSF.DDA.NodeReplacedEventArgs=function OSF_DDA_NodeReplacedEventArgs(oldNode, newNode, inUndoRedo) {
	OSF.OUtil.defineEnumerableProperties(this, {
		"type": {
			value: Microsoft.Office.WebExtension.EventType.DataNodeReplaced
		},
		"oldNode": {
			value: oldNode
		},
		"newNode": {
			value: newNode
		},
		"inUndoRedo": {
			value: inUndoRedo
		}
	});
};
OSF.DDA.NodeDeletedEventArgs=function OSF_DDA_NodeDeletedEventArgs(oldNode, oldNextSibling, inUndoRedo) {
	OSF.OUtil.defineEnumerableProperties(this, {
		"type": {
			value: Microsoft.Office.WebExtension.EventType.DataNodeDeleted
		},
		"oldNode": {
			value: oldNode
		},
		"oldNextSibling": {
			value: oldNextSibling
		},
		"inUndoRedo": {
			value: inUndoRedo
		}
	});
};
OSF.DDA.RichClientSettingsManager=(funct
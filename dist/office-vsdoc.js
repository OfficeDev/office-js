/* Version: 16.0.15307.10000 */
/*
	Copyright (c) Microsoft Corporation.  All rights reserved.
*/

/*
	Your use of this file is governed by the Microsoft Services Agreement http://go.microsoft.com/fwlink/?LinkId=266419.
*/

if (!Office) {
    var Office = new function () {
        this._appContext = 127; // All hosts; see _officeintellisense.js in VS, with binary 1111111
        this._showAll = true;
        this._setContext = {};
    }
}

Office._ExcelMask = 0x1;
Office._WordMask = 0x2;
Office._ProjectMask = 0x4;
Office._OutlookMask = 0x8;
Office._PowerPointMask = 0x10;
Office._OutlookComposeMask = 0x20;
Office._AccessWebAppMask = 0x40;

{
    Office._extractedCallback = function (originalArgs, totalArgsCount, optionalArgsCount) {
        var optionalArgs = Array.prototype.slice.call(originalArgs, totalArgsCount - optionalArgsCount);
        var callback = function(){};
        for (var i = Math.min(optionalArgs.length, optionalArgsCount) - 1; i >= 0; i--) {
            if (typeof optionalArgs[i] == "function") {
                callback = optionalArgs[i];
                break;
            }
        }
        return callback;
    }

    Office._BindingDataChangedEvents = function (eventType) {
        this.binding = new Office._Binding(bindingType);
        this.type = eventType;
        this.startColumn = {};
        this.startRow = {};
    }

    Office._DocumentEventArgs = function (eventType) {
        Office._processContents(this, {
            type: {
                annotate: {
                    /// <field type="Office.EventType"></field>
                    type: undefined
                },
                value: eventType
            }
        });
        if (eventType == "activeViewChanged") {
            Office._processItem(
                this,
                {
                    annotate: {
                        /// <field type="Office.ActiveView"></field>
                        activeView: undefined
                    }
                },
                "activeView"
            );
        }
    }

    Office._DialogEventArgs = function (eventType) {
        Office._processContents(this, {
            type: {
                annotate: {
                    /// <field type="Office.EventType"></field>
                    type: undefined
                },
                value: eventType
            }
        });
        if (eventType == "dialogMessageReceived") {
            Office._processItem(
                this,
                {
                    annotate: {
                        /// <field type="string"></field>
                        message: undefined
                    }
                },
                "message"
            );
        } else if (eventType == "dialogEventReceived") {
            Office._processItem(
                this,
                {
                    annotate: {
                        /// <field type="Integer"></field>
                        error: undefined
                    }
                },
                "error"
            );
        }
    }

    Office._CustomXmlNodeEvents = function (eventType) {
        this.type = eventType;
        this.inUndoRedo = {};

        if (eventType == 'nodeDeleted') {
            this.oldNode = new Office._CustomXmlNode();
            this.oldNextSibling = new Office._CustomXmlNode();
        }

        if (eventType == 'nodeInserted') {
            this.newNode = new Office._CustomXmlNode();


        }
        if (eventType == 'nodeReplaced') {
            this.oldNode = new Office._CustomXmlNode();
            this.newNode = new Office._CustomXmlNode();

        }
    }

    Office._Error = function () {
        this.id = {};
        this.message = {};
        this.name = {};
    }

    Office._CustomXmlNode = function () {
        this.baseName = {};
        this.namespaceUri = {};
        this.nodeType = {};

        this.getNodesAsync = function (xPath, callback) {
            ///<summary> Gets the nodes associated with the xPath expression.  </summary>
            ///<param name="xPath" type="string">The xPath expression</param>
            ///<param name="callback" type="function" optional="true">The optional callback method</param>
            var result = new Office._AsyncResult("getNodesAsync");
            callback(result);
        };
        this.getNodeValueAsync = function (callback) {
            ///<summary> Gets the node value.  </summary>
            ///<param name="callback" type="function" optional="true">The optional callback method</param>

            var result = new Office._AsyncResult("getNodeValueAsync");
            callback(result);
        };
        this.getXmlAsync = function (callback) {
            ///<summary> Gets the node's XML.  </summary>
            ///<param name="callback" type="function" optional="true">The optional callback method</param>
            var result = new Office._AsyncResult("getXmlAsync");
            callback(result);
        };
        this.setNodeValueAsync = function (value, callback) {
            ///<summary> Sets the node value.  </summary>
            ///<param name="value" type="string">The value to be set on the node</param>
            ///<param name="callback" type="function" optional="true">The optional callback method</param>
            var result = new Office._AsyncResult("setNodeValueAsync");
            callback(result);
        };
        this.setXmlAsync = function (xml, callback) {
            ///<summary> Sets the node XML.  </summary>
            ///<param name="xml" type="string">The XML to be set on the node</param>
            ///<param name="callback" type="function" optional="true">The optional callback method</param>
            var result = new Office._AsyncResult("setXmlAsync");
            callback(result);
        };
    }

    Office._context_document_customXmlParts_customXmlPrefixMappings = function () {
        this.addNamespaceAsync = function (prefix, nsUri, callback) {
            ///<summary>Adds a namespace.  </summary>
            //////<param name="prefix" type="string">The namespace prefix</param>
            //////<param name="nsUri" type="string">The namespace URI</param>
            ///<param name="callback" type="function" optional="true">The optional callback method</param>

            var result = new Office._AsyncResult("addNamespaceAsync");
            callback(result);
        };
        this.getNamespaceAsync = function (prefix, callback) {
            ///<summary> Gets a namespace  with the specified prefix </summary>
            ///<param name="prefix" type="string">The namespace prefix</param>
            ///<param name="callback" type="function" optional="true">The optional callback method</param>
            var result = new Office._AsyncResult("getNamespaceAsync");
            callback(result);
        };
        this.getPrefixAsync = function (nsUri, callback) {
            ///<summary> Gets a prefix  for  the specified URI </summary>
            ///<param name="nsUri" type="string">The namespace URI</param>
            ///<param name="callback" type="function" optional="true">The optional callback method</param>

            var result = new Office._AsyncResult("getPrefixAsync");
            callback(result);
        };
    }

    Office._CustomXmlPart = function () {
        this.builtIn = {};
        this.id = {};
        this.namespaceManager = new Office._context_document_customXmlParts_customXmlPrefixMappings();

        this.deleteAsync = function (callback) {
            ///<summary> Deletes the Custom XML Part.  </summary>
            ///<param name="callback" type="function" optional="true">The optional callback method</param>
            var result = new Office._AsyncResult("deleteAsync");
            callback(result);
        };
        this.getNodesAsync = function (xPath, callback) {
            ///<summary> Gets the nodes associated with the xPath expression.  </summary>
            ///<param name="xPath" type="string">The xPath expression</param>
            ///<param name="callback" type="function" optional="true">The optional callback method</param>

            var result = new Office._AsyncResult("getNodesAsync");
            callback(result);
        };
        this.getXmlAsync = function (callback) {
            ///<summary> Gets the XML for the Custom XML Part.  </summary>
            ///<param name="callback" type="function" optional="true">The optional callback method</param>
            var result = new Office._AsyncResult("getXmlAsync");
            callback(result);
        };

        this.addHandlerAsync = function (eventType, handler, callback) {
            ///<summary> Adds an event handler to the object using the specified event type.  </summary>
            ///<param name="eventType" type="Office.EventType">The event type. For CustomXmlPartNode it can be 'nodeDeleted', 'nodeInserted' or 'nodeReplaced' </param>
            ///<param name="handler" type="function">The name of the handler </param>
            ///<param name="callback" type="function" optional="true">The optional callback method</param>


            var events = new Office._CustomXmlNodeEvents(eventType);
            handler(events);

            var result = new Office._AsyncResult("addHandlerAsync");
            callback(result);
        };

        this.removeHandlerAsync = function (eventType, options, callback) {
            ///<summary> Removes an event handler from the object using the specified event type.  </summary>
            ///<param name="eventType" type="Office.EventType">The event type. For CustomXmlPartNode it can be 'nodeDeleted', 'nodeInserted' or 'nodeReplaced' </param>
            ///<param name="options" type="Object" optional="true">
            ///    Syntax example: {handler:eventHandler}
            /// &#10;     handler: Indicates a specific handler to be removed, if not specified all handlers are removed
            /// &#10;     asyncContext: Object keeping state for the callback
            ///</param>
            ///<param name="callback" type="function" optional="true">The optional callback method</param>
            ///
            Office._extractedCallback(arguments, 3, 2)(new Office._AsyncResult("removeHandlerAsync"));
        }
    }

    Office._Binding = function (bindingType) {
        ///<field type="String" name="id">Id of the Binding</field>};
        this.id = {};

        this.type = {};
        this.document = {};

        this.setDataAsync = function (data, options, callback) {
            ///<summary> Writes the specified data into the current selection.</summary>
            ///<param name="data" type="Object">The data to be set. Either a string or value, 2d array or TableData object</param>
            ///<param name="options" type="Object" optional="true">
            ///    Syntax example: {coercionType:Office.CoercionType.Matrix} or {coercionType: 'matrix'}
            /// &#10;     coercionType: Explicitly sets the shape of the data object. Use Office.CoercionType or text value. If not supplied is inferred from the data type.
            /// &#10;     startRow: Used in partial set for table/matrix. Indicates the start row.
            /// &#10;     startColumn: Used in partial set for table/matrix. Indicates the start column.

            /// &#10;     asyncContext: Object keeping state for the callback
            ///</param>
            ///<param name="callback" type="function" optional="true">The optional callback method</param>
            Office._extractedCallback(arguments, 3, 2)(new Office._AsyncResult("setDataAsync"));
        };

        this.getDataAsync = function (options, callback) {
            ///<summary> Returns the current selection.</summary>
            ///<param name="options" type="Object" optional="true">
            ///    Syntax example: {coercionType: 'matrix,'valueFormat: 'formatted', filterType:'all'}
            /// &#10;     coercionType: The expected shape of the selection. If not specified returns the bindingType shape. Use Office.CoercionType or text value.
            /// &#10;     valueFormat: Get data with or without format. Use Office.ValueFormat or text value.
            /// &#10;     startRow: Used in partial get for table/matrix. Indicates the start row.
            /// &#10;     startColumn: Used in partial get for table/matrix. Indicates the start column.
            /// &#10;     rowCount: Used in partial get for table/matrix. Indicates the number of rows from the start row.
            /// &#10;     columnCount: Used in partial get for table/matrix. Indicates the number of columns from the start column.
            /// &#10;     filterType: Get the visible or all the data. Useful when filtering data. Use Office.FilterType or text value.
            /// &#10;     asyncContext: Object keeping state for the callback
            ///</param>
            ///<param name="callback" type="function" optional="true">The optional callback method</param>

            if (arguments.length == 1) {
                callback = options;
            }
            var result = new Office._AsyncResult("getDataAsync", options.coercionType);
            callback(result);
        };

        this.addHandlerAsync = function (eventType, handler, callback) {
            ///<summary> Adds an event handler to the object using the specified event type.  </summary>
            ///<param name="eventType" type="Office.EventType">The event type. For binding it can be 'bindingDataChanged' and 'bindingSelectionChanged' </param>
            ///<param name="handler" type="function">The name of the handler </param>
            ///<param name="callback" type="function" optional="true">The optional callback method</param>


            var events = new Office._BindingDataChangedEvents(eventType);
            handler(events);

            var result = new Office._AsyncResult("addHandlerAsync");
            callback(result);
        };

        this.removeHandlerAsync = function (eventType, options, callback) {
            ///<summary> Removes an event handler from the object using the specified event type.  </summary>
            ///<param name="eventType" type="Office.EventType">The event type. For binding can be 'bindingDataChanged' and 'bindingSelectionChanged' </param>
            ///<param name="options" type="Object" optional="true">
            ///    Syntax example: {handler:eventHandler}
            /// &#10;     handler: Indicates a specific handler to be removed, if not specified all handlers are removed
            /// &#10;     asyncContext: Object keeping state for the callback
            ///</param>
            ///<param name="callback" type="function" optional="true">The optional callback method</param>


            var events = new Office._BindingDataChangedEvents(eventType);
            handler(events);

            Office._extractedCallback(arguments, 3, 2)(new Office._AsyncResult("removeHandlerAsync"));
        };


        if ((bindingType == undefined) || (bindingType == Office.BindingType.Matrix) || (bindingType == Office.BindingType.Table)) {
            this.columnCount = {};
            this.rowCount = {};
        }
        if ((bindingType == undefined) || (bindingType == Office.BindingType.Table)) {
            Office._processContents(this, {
                hasHeaders: {
                    value: {}
                },
                addRowsAsync: {
                    value: function (data, callback) {
                        ///<summary> Adds the specified rows to the table  </summary>
                        ///<param name="data" type="Object"> A 2D array with the rows to
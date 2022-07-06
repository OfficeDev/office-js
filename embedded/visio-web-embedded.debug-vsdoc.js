
var Visio;
(function (Visio) {
	var Application = (function(_super) {
		__extends(Application, _super);
		function Application() {
			/// <summary> Represents the Application. [Api set:  1.1] </summary>
			/// <field name="context" type="Visio.RequestContext">The request context associated with this object.</field>
			/// <field name="isNull" type="Boolean">Returns a boolean value for whether the corresponding object is null. You must call "context.sync()" before reading the isNull property.</field>
			/// <field name="showBorders" type="Boolean">Show or hide the iFrame application borders. [Api set:  1.1]</field>
			/// <field name="showToolbars" type="Boolean">Show or hide the standard toolbars. [Api set:  1.1]</field>
		}

		Application.prototype.load = function(option) {
			/// <summary>
			/// Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
			/// </summary>
			/// <param name="option" type="string | string[] | OfficeExtension.LoadOption"/>
			/// <returns type="Visio.Application"/>
		}

		Application.prototype.set = function() {
			/// <signature>
			/// <summary>Sets multiple properties on the object at the same time, based on JSON input.</summary>
			/// <param name="properties" type="Visio.Interfaces.ApplicationUpdateData">Properties described by the Visio.Interfaces.ApplicationUpdateData interface.</param>
			/// <param name="options" type="string">Options of the form { throwOnReadOnly?: boolean }
			/// <br />
			/// * throwOnReadOnly: Throw an error if the passed-in property list includes read-only properties (default = true).
			/// </param>
			/// </signature>
			/// <signature>
			/// <summary>Sets multiple properties on the object at the same time, based on an existing loaded object.</summary>
			/// <param name="properties" type="Application">An existing Application object, with properties that have already been loaded and synced.</param>
			/// </signature>
		}
		Application.prototype.showToolbar = function(id, show) {
			/// <summary>
			/// Sets the visibility of a specific toolbar in the application. [Api set:  1.1]
			/// </summary>
			/// <param name="id" type="String">The type of the Toolbar</param>
			/// <param name="show" type="Boolean">Whether the toolbar is visibile or not.</param>
			/// <returns ></returns>
		}

		return Application;
	})(OfficeExtension.ClientObject);
	Visio.Application = Application;
})(Visio || (Visio = {__proto__: null}));

var Visio;
(function (Visio) {
	var Interfaces;
	(function (Interfaces) {
		var BoundingBox = (function() {
			function BoundingBox() {
				/// <summary> Represents the BoundingBox of the shape. [Api set:  1.1] </summary>
				/// <field name="height" type="Number">The distance between the top and bottom edges of the bounding box of the shape, excluding any data graphics associated with the shape. [Api set:  1.1]</field>
				/// <field name="width" type="Number">The distance between the left and right edges of the bounding box of the shape, excluding any data graphics associated with the shape. [Api set:  1.1]</field>
				/// <field name="x" type="Number">An integer that specifies the x-coordinate of the bounding box. [Api set:  1.1]</field>
				/// <field name="y" type="Number">An integer that specifies the y-coordinate of the bounding box. [Api set:  1.1]</field>
			}
			return BoundingBox;
		})();
		Interfaces.BoundingBox.__proto__ = null;
		Interfaces.BoundingBox = BoundingBox;
	})(Interfaces = Visio.Interfaces || (Visio.Interfaces = { __proto__: null}));
})(Visio || (Visio = {__proto__: null}));

var Visio;
(function (Visio) {
	/// <summary> Represents the type of column values. [Api set:  1.1] </summary>
	var ColumnType = {
		__proto__: null,
		"unknown": "unknown",
		"string": "string",
		"number": "number",
		"date": "date",
		"currency": "currency",
	}
	Visio.ColumnType = ColumnType;
})(Visio || (Visio = {__proto__: null}));

var Visio;
(function (Visio) {
	var Comment = (function(_super) {
		__extends(Comment, _super);
		function Comment() {
			/// <summary> Represents the Comment. [Api set:  1.1] </summary>
			/// <field name="context" type="Visio.RequestContext">The request context associated with this object.</field>
			/// <field name="isNull" type="Boolean">Returns a boolean value for whether the corresponding object is null. You must call "context.sync()" before reading the isNull property.</field>
			/// <field name="author" type="String">A string that specifies the name of the author of the comment. [Api set:  1.1]</field>
			/// <field name="date" type="String">A string that specifies the date when the comment was created. [Api set:  1.1]</field>
			/// <field name="text" type="String">A string that contains the comment text. [Api set:  1.1]</field>
		}

		Comment.prototype.load = function(option) {
			/// <summary>
			/// Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
			/// </summary>
			/// <param name="option" type="string | string[] | OfficeExtension.LoadOption"/>
			/// <returns type="Visio.Comment"/>
		}

		Comment.prototype.set = function() {
			/// <signature>
			/// <summary>Sets multiple properties on the object at the same time, based on JSON input.</summary>
			/// <param name="properties" type="Visio.Interfaces.CommentUpdateData">Properties described by the Visio.Interfaces.CommentUpdateData interface.</param>
			/// <param name="options" type="string">Options of the form { throwOnReadOnly?: boolean }
			/// <br />
			/// * throwOnReadOnly: Throw an error if the passed-in property list includes read-only properties (default = true).
			/// </param>
			/// </signature>
			/// <signature>
			/// <summary>Sets multiple properties on the object at the same time, based on an existing loaded object.</summary>
			/// <param name="properties" type="Comment">An existing Comment object, with properties that have already been loaded and synced.</param>
			/// </signature>
		}

		return Comment;
	})(OfficeExtension.ClientObject);
	Visio.Comment = Comment;
})(Visio || (Visio = {__proto__: null}));

var Visio;
(function (Visio) {
	var CommentCollection = (function(_super) {
		__extends(CommentCollection, _super);
		function CommentCollection() {
			/// <summary> Represents the CommentCollection for a given Shape. [Api set:  1.1] </summary>
			/// <field name="context" type="Visio.RequestContext">The request context associated with this object.</field>
			/// <field name="isNull" type="Boolean">Returns a boolean value for whether the corresponding object is null. You must call "context.sync()" before reading the isNull property.</field>
			/// <field name="items" type="Array" elementType="Visio.Comment">Gets the loaded child items in this collection.</field>
		}

		CommentCollection.prototype.load = function(option) {
			/// <summary>
			/// Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
			/// </summary>
			/// <param name="option" type="string | string[] | OfficeExtension.LoadOption"/>
			/// <returns type="Visio.CommentCollection"/>
		}
		CommentCollection.prototype.getCount = function() {
			/// <summary>
			/// Gets the number of Comments. [Api set:  1.1]
			/// </summary>
			/// <returns type="OfficeExtension.ClientResult&lt;number&gt;"></returns>
			var result = new OfficeExtension.ClientResult();
			result.__proto__ = null;
			result.value = 0;
			return result;
		}
		CommentCollection.prototype.getItem = function(key) {
			/// <summary>
			/// Gets the Comment using its name. [Api set:  1.1]
			/// </summary>
			/// <param name="key" >Key is the name of the Comment to be retrieved.</param>
			/// <returns type="Visio.Comment"></returns>
		}

		return CommentCollection;
	})(OfficeExtension.ClientObject);
	Visio.CommentCollection = CommentCollection;
})(Visio || (Visio = {__proto__: null}));

var Visio;
(function (Visio) {
	var Interfaces;
	(function (Interfaces) {
		var ConnectorBinding = (function() {
			function ConnectorBinding() {
				/// <summary> Connector bindings for data visualizer diagram. [Api set:  1.1] </summary>
				/// <field name="delimiter" type="String">Delimiter for TargetColumn. It should not have more then one character. [Api set:  1.1]</field>
			}
			return ConnectorBinding;
		})();
		Interfaces.ConnectorBinding.__proto__ = null;
		Interfaces.ConnectorBinding = ConnectorBinding;
	})(Interfaces = Visio.Interfaces || (Visio.Interfaces = { __proto__: null}));
})(Visio || (Visio = {__proto__: null}));

var Visio;
(function (Visio) {
	/// <summary> Direction of connector in DataVisualizer diagram. [Api set:  1.1] </summary>
	var ConnectorDirection = {
		__proto__: null,
		"fromTarget": "fromTarget",
		"toTarget": "toTarget",
	}
	Visio.ConnectorDirection = ConnectorDirection;
})(Visio || (Visio = {__proto__: null}));

var Visio;
(function (Visio) {
	/// <summary> Represents the orientation of the Cross Functional Flowchart diagram. [Api set:  1.1] </summary>
	var CrossFunctionalFlowchartOrientation = {
		__proto__: null,
		"horizontal": "horizontal",
		"vertical": "vertical",
	}
	Visio.CrossFunctionalFlowchartOrientation = CrossFunctionalFlowchartOrientation;
})(Visio || (Visio = {__proto__: null}));

var Visio;
(function (Visio) {
	var Interfaces;
	(function (Interfaces) {
		var DataRefreshCompleteEventArgs = (function() {
			function DataRefreshCompleteEventArgs() {
				/// <summary> Provides information about the document that raised the DataRefreshComplete event. [Api set:  1.1] </summary>
				/// <field name="document" type="Visio.Document">Gets the document object that raised the DataRefreshComplete event. [Api set:  1.1]</field>
				/// <field name="success" type="Boolean">Gets the success or failure of the DataRefreshComplete event. [Api set:  1.1]</field>
			}
			return DataRefreshCompleteEventArgs;
		})();
		Interfaces.DataRefreshCompleteEventArgs.__proto__ = null;
		Interfaces.DataRefreshCompleteEventArgs = DataRefreshCompleteEventArgs;
	})(Interfaces = Visio.Interfaces || (Visio.Interfaces = { __proto__: null}));
})(Visio || (Visio = {__proto__: null}));

var Visio;
(function (Visio) {
	/// <summary> Represents the type of source for the data connection. [Api set:  1.1] </summary>
	var DataSourceType = {
		__proto__: null,
		"unknown": "unknown",
		"excel": "excel",
	}
	Visio.DataSourceType = DataSourceType;
})(Visio || (Visio = {__proto__: null}));

var Visio;
(function (Visio) {
	/// <summary> Represents the types of data validation error. [Api set:  1.1] </summary>
	var DataValidationErrorType = {
		__proto__: null,
		"none": "none",
		"columnNotMapped": "columnNotMapped",
		"uniqueIdColumnError": "uniqueIdColumnError",
		"swimlaneColumnError": "swimlaneColumnError",
		"delimiterError": "delimiterError",
		"connectorColumnError": "connectorColumnError",
		"connectorColumnMappedElsewhere": "connectorColumnMappedElsewhere",
		"connectorLabelColumnMappedElsewhere": "connectorLabelColumnMappedElsewhere",
		"connectorColumnAndConnectorLabelMappedElsewhere": "connectorColumnAndConnectorLabelMappedElsewhere",
	}
	Visio.DataValidationErrorType = DataValidationErrorType;
})(Visio || (Visio = {__proto__: null}));

var Visio;
(function (Visio) {
	/// <summary> Type of the Data Visualizer Diagram operation [Api set:  1.1] </summary>
	var DataVisualizerDiagramOperationType = {
		__proto__: null,
		"unknown": "unknown",
		"create": "create",
		"updateMappings": "updateMappings",
		"updateData": "updateData",
		"update": "update",
		"delete": "delete",
	}
	Visio.DataVisualizerDiagramOperationType = DataVisualizerDiagramOperationType;
})(Visio || (Visio = {__proto__: null}));

var Visio;
(function (Visio) {
	/// <summary> Result of Data Visualizer Diagram operations. [Api set:  1.1] </summary>
	var DataVisualizerDiagramResultType = {
		__proto__: null,
		"success": "success",
		"unexpected": "unexpected",
		"validationError": "validationError",
		"conflictError": "conflictError",
	}
	Visio.DataVisualizerDiagramResultType = DataVisualizerDiagramResultType;
})(Visio || (Visio = {__proto__: null}));

var Visio;
(function (Visio) {
	/// <summary> DiagramType for Data Visualizer diagrams [Api set:  1.1] </summary>
	var DataVisualizerDiagramType = {
		__proto__: null,
		"unknown": "unknown",
		"basicFlowchart": "basicFlowchart",
		"crossFunctionalFlowchart_Horizontal": "crossFunctionalFlowchart_Horizontal",
		"crossFunctionalFlowchart_Vertical": "crossFunctionalFlowchart_Vertical",
		"audit": "audit",
		"orgChart": "orgChart",
		"network": "network",
	}
	Visio.DataVisualizerDiagramType = DataVisualizerDiagramType;
})(Visio || (Visio = {__proto__: null}));

var Visio;
(function (Visio) {
	var Document = (function(_super) {
		__extends(Document, _super);
		function Document() {
			/// <summary> Represents the Document class. [Api set:  1.1] </summary>
			/// <field name="context" type="Visio.RequestContext">The request context associated with this object.</field>
			/// <field name="isNull" type="Boolean">Returns a boolean value for whether the corresponding object is null. You must call "context.sync()" before reading the isNull property.</field>
			/// <field name="application" type="Visio.Application">Represents a Visio application instance that contains this document. Read-only. [Api set:  1.1]</field>
			/// <field name="pages" type="Visio.PageCollection">Represents a collection of pages associated with the document. Read-only. [Api set:  1.1]</field>
			/// <field name="view" type="Visio.DocumentView">Returns the DocumentView object. Read-only. [Api set:  1.1]</field>
			/// <field name="onDataRefreshComplete" type="OfficeExtension.EventHandlers">Occurs when the data is refreshed in the diagram. [Api set:  1.1]</field>
			/// <field name="onDocumentError" type="OfficeExtension.EventHandlers">Occurs when there is an expected or unexpected error occured in the session. [Api set:  1.1]</field>
			/// <field name="onDocumentLoadComplete" type="OfficeExtension.EventHandlers">Occurs when the Document is loaded, refreshed, or changed. [Api set:  1.1]</field>
			/// <field name="onPageLoadComplete" type="OfficeExtension.EventHandlers">Occurs when the page is finished loading. [Api set:  1.1]</field>
			/// <field name="onSelectionChanged" type="OfficeExtension.EventHandlers">Occurs when the current selection of shapes changes. [Api set:  1.1]</field>
			/// <field name="onShapeMouseEnter" type="OfficeExtension.EventHandlers">Occurs when the user moves the mouse pointer into the bounding box of a shape. [Api set:  1.1]</field>
			/// <field name="onShapeMouseLeave" type="OfficeExtension.EventHandlers">Occurs when the user moves the mouse out of the bounding box of a shape. [Api set:  1.1]</field>
			/// <field name="onTaskPaneStateChanged" type="OfficeExtension.EventHandlers">Occurs whenever a task pane state is changed [Api set:  1.1]</field>
		}

		Document.prototype.load = function(option) {
			/// <summary>
			/// Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
			/// </summary>
			/// <param name="option" type="string | string[] | OfficeExtension.LoadOption"/>
			/// <returns type="Visio.Document"/>
		}

		Document.prototype.set = function() {
			/// <signature>
			/// <summary>Sets multiple properties on the object at the same time, based on JSON input.</summary>
			/// <param name="properties" type="Visio.Interfaces.DocumentUpdateData">Properties described by the Visio.Interfaces.DocumentUpdateData interface.</param>
			/// <param name="options" type="string">Options of the form { throwOnReadOnly?: boolean }
			/// <br />
			/// * throwOnReadOnly: Throw an error if the passed-in property list includes read-only properties (default = true).
			/// </param>
			/// </signature>
			/// <signature>
			/// <summary>Sets multiple properties on the object at the same time, based on an existing loaded object.</summary>
			/// <param name="properties" type="Document">An existing Document object, with properties that have already been loaded and synced.</param>
			/// </signature>
		}
		Document.prototype.getActivePage = function() {
			/// <summary>
			/// Returns the Active Page of the document. [Api set:  1.1]
			/// </summary>
			/// <returns type="Visio.Page"></returns>
		}
		Document.prototype.setActivePage = function(PageName) {
			/// <summary>
			/// Set the Active Page of the document. [Api set:  1.1]
			/// </summary>
			/// <param name="PageName" type="String">Name of the page</param>
			/// <returns ></returns>
		}
		Document.prototype.showTaskPane = function(taskPaneType, initialProps, show) {
			/// <summary>
			/// Show or Hide a TaskPane.              This will be consumed by the DV Excel Add-In/Other third-party apps who embed the visio drawing to show/hide the task pane. [Api set:  1.1]
			/// </summary>
			/// <param name="taskPaneType" type="String">Type of the 1st Party TaskPane. It can take values from enum TaskPaneType</param>
			/// <param name="initialProps"  optional="true">Optional Parameter. This is a generic data structure which would be filled with initial data required to initialize the content of the Taskpane</param>
			/// <param name="show" type="Boolean" optional="true">Optional Parameter. If it is set to false, it will hide the specified taskpane</param>
			/// <returns ></returns>
		}
		Document.prototype.startDataRefresh = function() {
			/// <summary>
			/// Triggers the refresh of the data in the Diagram, for all pages. [Api set:  1.1]
			/// </summary>
			/// <returns ></returns>
		}
		Document.prototype.onDataRefreshComplete = {
			__proto__: null,
			add: function (handler) {
				/// <param name="handler" type="function(eventArgs: Visio.Interfaces.DataRefreshCompleteEventArgs)">Handler for the event. EventArgs: Provides information about the document that raised the DataRefreshComplete event. </param>
				/// <returns type="OfficeExtension.EventHandlerResult"></returns>
				var eventInfo = new Visio.Interfaces.DataRefreshCompleteEventArgs();
				eventInfo.__proto__ = null;
				handler(eventInfo);
			},
			remove: function (handler) {
				/// <param name="handler" type="function(eventArgs: Visio.Interfaces.DataRefreshCompleteEventArgs)">Handler for the event.</param>
				return;
			}
		};
		Document.prototype.onDocumentError = {
			__proto__: null,
			add: function (handler) {
				/// <param name="handler" type="function(eventArgs: Visio.Interfaces.DocumentErrorEventArgs)">Handler for the event. EventArgs: Provides information about DocumentError event </param>
				/// <returns type="OfficeExtension.EventHandlerResult"></returns>
				var eventInfo = new Visio.Interfaces.DocumentErrorEventArgs();
				eventInfo.__proto__ = null;
				handler(eventInfo);
			},
			remove: function (handler) {
				/// <param name="handler" type="function(eventArgs: Visio.Interfaces.DocumentErrorEventArgs)">Handler for the event.</param>
				return;
			}
		};
		Document.prototype.onDocumentLoadComplete = {
			__proto__: null,
			add: function (handler) {
				/// <param name="handler" type="function(eventArgs: Visio.Interfaces.DocumentLoadCompleteEventArgs)">Handler for the event. EventArgs: Provides information about the success or failure of the DocumentLoadComplete event. </param>
				/// <returns type="OfficeExtension.EventHandlerResult"></returns>
				var eventInfo = new Visio.Interfaces.DocumentLoadCompleteEventArgs();
				eventInfo.__proto__ = null;
				handler(eventInfo);
			},
			remove: function (handler) {
				/// <param name="handler" type="function(eventArgs: Visio.Interfaces.DocumentLoadCompleteEventArgs)">Handler for the event.</param>
				return;
			}
		};
		Document.prototype.onPageLoadComplete = {
			__proto__: null,
			add: function (handler) {
				/// <param name="handler" type="function(eventArgs: Visio.Interfaces.PageLoadCompleteEventArgs)">Handler for the event. EventArgs: Provides information about the page that raised the PageLoadComplete event. </param>
				/// <returns type="OfficeExtension.EventHandlerResult"></returns>
				var eventInfo = new Visio.Interfaces.PageLoadCompleteEventArgs();
				eventInfo.__proto__ = null;
				handler(eventInfo);
			},
			remove: function (handler) {
				/// <param name="handler" type="function(eventArgs: Visio.Interfaces.PageLoadCompleteEventArgs)">Handler for the event.</param>
				return;
			}
		};
		Document.prototype.onSelectionChanged = {
			__proto__: null,
			add: function (handler) {
				/// <param name="handler" type="function(eventArgs: Visio.Interfaces.SelectionChangedEventArgs)">Handler for the event. EventArgs: Provides information about the shape collection that raised the SelectionChanged event. </param>
				/// <returns type="OfficeExtension.EventHandlerResult"></returns>
				var eventInfo = new Visio.Interfaces.SelectionChangedEventArgs();
				eventInfo.__proto__ = null;
				handler(eventInfo);
			},
			remove: function (handler) {
				/// <param name="handler" type="function(eventArgs: Visio.Interfaces.SelectionChangedEventArgs)">Handler for the event.</param>
				return;
			}
		};
		Document.prototype.onShapeMouseEnter = {
			__proto__: null,
			add: function (handler) {
				/// <param name="handler" type="function(eventArgs: Visio.Interfaces.ShapeMouseEnterEventArgs)">Handler for the event. EventArgs: Provides information about the shape that raised the ShapeMouseEnter event. </param>
				/// <returns type="OfficeExtension.EventHandlerResult"></returns>
				var eventInfo = new Visio.Interfaces.ShapeMouseEnterEventArgs();
				eventInfo.__proto__ = null;
				handler(eventInfo);
			},
			remove: function (handler) {
				/// <param name="handler" type="function(eventArgs: Visio.Interfaces.ShapeMouseEnterEventArgs)">Handler for the event.</param>
				return;
			}
		};
		Document.prototype.onShapeMouseLeave = {
			__proto__: null,
			add: function (handler) {
				/// <param name="handler" type="function(eventArgs: Visio.Interfaces.ShapeMouseLeaveEventArgs)">Handler for the event. EventArgs: Provides information about the shape that raised the ShapeMouseLeave event. </param>
				/// <returns type="OfficeExtension.EventHandlerResult"></returns>
				var eventInfo = new Visio.Interfaces.ShapeMouseLeaveEventArgs();
				eventInfo.__proto__ = null;
				handler(eventInfo);
			},
			remove: function (handler) {
				/// <param name="handler" type="function(eventArgs: Visio.Interfaces.ShapeMouseLeaveEventArgs)">Handler for the event.</param>
				return;
			}
		};
		Document.prototype.onTaskPaneStateChanged = {
			__proto__: null,
			add: function (handler) {
				/// <param name="handler" type="function(eventArgs: Visio.Interfaces.TaskPaneStateChangedEventArgs)">Handler for the event. EventArgs: Provides information about the TaskPaneStateChanged event. </param>
				/// <returns type="OfficeExtension.EventHandlerResult"></returns>
				var eventInfo = new Visio.Interfaces.TaskPaneStateChangedEventArgs();
				eventInfo.__proto__ = null;
				handler(eventInfo);
			},
			remove: function (handler) {
				/// <param name="handler" type="function(eventArgs: Visio.Interfaces.TaskPaneStateChangedEventArgs)">Handler for the event.</param>
				return;
			}
		};

		return Document;
	})(OfficeExtension.ClientObject);
	Visio.Document = Document;
})(Visio || (Visio = {__proto__: null}));

var Visio;
(function (Visio) {
	var Interfaces;
	(function (Interfaces) {
		var DocumentErrorEventArgs = (function() {
			function DocumentErrorEventArgs() {
				/// <summary> Provides information about DocumentError event [Api set:  1.1] </summary>
				/// <field name="errorCode" type="Number">Visio Error code [Api set:  1.1]</field>
				/// <field name="errorMessage" type="String">Message about error that occured [Api set:  1.1]</field>
				/// <field name="isCritical" type="Boolean">Tells if the error is critical or not. If critical the session cannot continue. [Api set:  1.1]</field>
			}
			return DocumentErrorEventArgs;
		})();
		Interfaces.DocumentErrorEventArgs.__proto__ = null;
		Interfaces.DocumentErrorEventArgs = DocumentErrorEventArgs;
	})(Interfaces = Visio.Interfaces || (Visio.Interfaces = { __proto__: null}));
})(Visio || (Visio = {__proto__: null}));

var Visio;
(function (Visio) {
	var Interfaces;
	(function (Interfaces) {
		var DocumentLoadCompleteEventArgs = (function() {
			function DocumentLoadCompleteEventArgs() {
				/// <summary> Provides information about the success or failure of the DocumentLoadComplete event. [Api set:  1.1] </summary>
				/// <field name="success" type="Boolean">Gets the success or failure of the DocumentLoadComplete event. [Api set:  1.1]</field>
			}
			return DocumentLoadCompleteEventArgs;
		})();
		Interfaces.DocumentLoadCompleteEventArgs.__proto__ = null;
		Interfaces.DocumentLoadCompleteEventArgs = DocumentLoadCompleteEventArgs;
	})(Interfaces = Visio.Interfaces || (Visio.Interfaces = { __proto__: null}));
})(Visio || (Visio = {__proto__: null}));

var Visio;
(function (Visio) {
	var DocumentView = (function(_super) {
		__extends(DocumentView, _super);
		function DocumentView() {
			/// <summary> Represents the DocumentView class. [Api set:  1.1] </summary>
			/// <field name="context" type="Visio.RequestContext">The request context associated with this object.</field>
			/// <field name="isNull" type="Boolean">Returns a boolean value for whether the corresponding object is null. You must call "context.sync()" before reading the isNull property.</field>
			/// <field name="disableHyperlinks" type="Boolean">Disable Hyperlinks. [Api set:  1.1]</field>
			/// <field name="disablePan" type="Boolean">Disable Pan. [Api set:  1.1]</field>
			/// <field name="disablePanZoomWindow" type="Boolean">Disable PanZoomWindow. [Api set:  1.1]</field>
			/// <field name="disableZoom" type="Boolean">Disable Zoom. [Api set:  1.1]</field>
			/// <field name="hideDiagramBoundary" type="Boolean">Hide Diagram Boundary. [Api set:  1.1]</field>
		}

		DocumentView.prototype.load = function(option) {
			/// <summary>
			/// Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
			/// </summary>
			/// <param name="option" type="string | string[] | OfficeExtension.LoadOption"/>
			/// <returns type="Visio.DocumentView"/>
		}

		DocumentView.prototype.set = function() {
			/// <signature>
			/// <summary>Sets multiple properties on the object at the same time, based on JSON input.</summary>
			/// <param name="properties" type="Visio.Interfaces.DocumentViewUpdateData">Properties described by the Visio.Interfaces.DocumentViewUpdateData interface.</param>
			/// <param name="options" type="string">Options of the form { throwOnReadOnly?: boolean }
			/// <br />
			/// * throwOnReadOnly: Throw an error if the passed-in property list includes read-only properties (default = true).
			/// </param>
			/// </signature>
			/// <signature>
			/// <summary>Sets multiple properties on the object at the same time, based on an existing loaded object.</summary>
			/// <param name="properties" type="DocumentView">An existing DocumentView object, with properties that have already been loaded and synced.</param>
			/// </signature>
		}

		return DocumentView;
	})(OfficeExtension.ClientObject);
	Visio.DocumentView = DocumentView;
})(Visio || (Visio = {__proto__: null}));

var Visio;
(function (Visio) {
	/// <summary> EventType represents the type of the events Host supports [Api set:  1.1] </summary>
	var EventType = {
		__proto__: null,
		"dataVisualizerDiagramOperationCompleted": "dataVisualizerDiagramOperationCompleted",
	}
	Visio.EventType = EventType;
})(Visio || (Visio = {__proto__: null}));

var Visio;
(function (Visio) {
	var Interfaces;
	(function (Interfaces) {
		var Highlight = (function() {
			function Highlight() {
				/// <summary> Represents the highlight data added to the shape. [Api set:  1.1] </summary>
				/// <field name="color" type="String">A string that specifies the color of the highlight. It must have the form &quot;#RRGGBB&quot;, where each letter represents a hexadecimal digit between 0 and F, and where RR is the red value between 0 and 0xFF (255), GG the green value between 0 and 0xFF (255), and BB is the blue value between 0 and 0xFF (255). [Api set:  1.1]</field>
				/// <field name="width" type="Number">A positive integer that specifies the width of the highlight&apos;s stroke in pixels. [Api set:  1.1]</field>
			}
			return Highlight;
		})();
		Interfaces.Highlight.__proto__ = null;
		Interfaces.Highlight = Highlight;
	})(Interfaces = Visio.Interfaces || (Visio.Interfaces = { __proto__: null}));
})(Visio || (Visio = {__proto__: null}));

var Visio;
(function (Visio) {
	var Hyperlink = (function(_super) {
		__extends(Hyperlink, _super);
		function Hyperlink() {
			/// <summary> Represents the Hyperlink. [Api set:  1.1] </summary>
			/// <field name="context" type="Visio.RequestContext">The request context associated with this object.</field>
			/// <field name="isNull" type="Boolean">Returns a boolean value for whether the corresponding object is null. You must call "context.sync()" before reading the isNull property.</field>
			/// <field name="address" type="String">Gets the address of the Hyperlink object. Read-only. [Api set:  1.1]</field>
			/// <field name="description" type="String">Gets the description of a hyperlink. Read-only. [Api set:  1.1]</field>
			/// <field name="extraInfo" type="String">Gets the extra URL request information used to resolve the hyperlink&apos;s URL. Read-only. [Api set:  1.1]</field>
			/// <field name="subAddress" type="String">Gets the sub-address of the Hyperlink object. Read-only. [Api set:  1.1]</field>
		}

		Hyperlink.prototype.load = function(option) {
			/// <summary>
			/// Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
			/// </summary>
			/// <param name="option" type="string | string[] | OfficeExtension.LoadOption"/>
			/// <returns type="Visio.Hyperlink"/>
		}

		return Hyperlink;
	})(OfficeExtension.ClientObject);
	Visio.Hyperlink = Hyperlink;
})(Visio || (Visio = {__proto__: null}));

var Visio;
(function (Visio) {
	var HyperlinkCollection = (function(_super) {
		__extends(HyperlinkCollection, _super);
		function HyperlinkCollection() {
			/// <summary> Represents the Hyperlink Collection. [Api set:  1.1] </summary>
			/// <field name="context" type="Visio.RequestContext">The request context associated with this object.</field>
			/// <field name="isNull" type="Boolean">Returns a boolean value for whether the corresponding object is null. You must call "context.sync()" before reading the isNull property.</field>
			/// <field name="items" type="Array" elementType="Visio.Hyperlink">Gets the loaded child items in this collection.</field>
		}

		HyperlinkCollection.prototype.load = function(option) {
			/// <summary>
			/// Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
			/// </summary>
			/// <param name="option" type="string | string[] | OfficeExtension.LoadOption"/>
			/// <returns type="Visio.HyperlinkCollection"/>
		}
		HyperlinkCollection.prototype.getCount = function() {
			/// <summary>
			/// Gets the number of hyperlinks. [Api set:  1.1]
			/// </summary>
			/// <returns type="OfficeExtension.ClientResult&lt;number&gt;"></returns>
			var result = new OfficeExtension.ClientResult();
			result.__proto__ = null;
			result.value = 0;
			return result;
		}
		HyperlinkCollection.prototype.getItem = function(Key) {
			/// <summary>
			/// Gets a Hyperlink using its key (name or Id). [Api set:  1.1]
			/// </summary>
			/// <param name="Key" >Key is the name or index of the Hyperlink to be retrieved.</param>
			/// <returns type="Visio.Hyperlink"></returns>
		}

		return HyperlinkCollection;
	})(OfficeExtension.ClientObject);
	Visio.HyperlinkCollection = HyperlinkCollection;
})(Visio || (Visio = {__proto__: null}));

var Visio;
(function (Visio) {
	/// <summary> Represents the type of layout. [Api set:  1.1] </summary>
	var LayoutVariant = {
		__proto__: null,
		"unknown": "unknown",
		"pageDefault": "pageDefault",
		"flowchart_TopToBottom": "flowchart_TopToBottom",
		"flowchart_BottomToTop": "flowchart_BottomToTop",
		"flowchart_LeftToRight": "flowchart_LeftToRight",
		"flowchart_RightToLeft": "flowchart_RightToLeft",
		"wideTree_DownThenRight": "wideTree_DownThenRight",
		"wideTree_DownThenLeft": "wideTree_DownThenLeft",
		"wideTree_RightThenDown": "wideTree_RightThenDown",
		"wideTree_LeftThenDown": "wideTree_LeftThenDown",
	}
	Visio.LayoutVariant = LayoutVariant;
})(Visio || (Visio = {__proto__: null}));

var Visio;
(function (Visio) {
	/// <summary> MessageType represents the type of message when event is fired from Host [Api set:  1.1] </summary>
	var MessageType = {
		__proto__: null,
		"none": 0,
		"dataVisualizerDiagramOperationCompletedEvent": 1,
	}
	Visio.MessageType = MessageType;
})(Visio || (Visio = {__proto__: null}));

var Visio;
(function (Visio) {
	/// <summary> Represents the Horizontal Alignment of the Overlay relative to the shape. [Api set:  1.1] </summary>
	var OverlayHorizontalAlignment = {
		__proto__: null,
		"left": "left",
		"center": "center",
		"right": "right",
	}
	Visio.OverlayHorizontalAlignment = OverlayHorizontalAlignment;
})(Visio || (Visio = {__proto__: null}));

var Visio;
(function (Visio) {
	/// <summary> Represents the type of the overlay. [Api set:  1.1] </summary>
	var OverlayType = {
		__proto__: null,
		"text": "text",
		"image": "image",
		"html": "html",
	}
	Visio.OverlayType = OverlayType;
})(Visio || (Visio = {__proto__: null}));

var Visio;
(function (Visio) {
	/// <summary> Represents the Vertical Alignment of the Overlay relative to the shape. [Api set:  1.1] </summary>
	var OverlayVerticalAlignment = {
		__proto__: null,
		"top": "top",
		"middle": "middle",
		"bottom": "bottom",
	}
	Visio.OverlayVerticalAlignment = OverlayVerticalAlignment;
})(Visio || (Visio = {__proto__: null}));

var Visio;
(function (Visio) {
	var Page = (function(_super) {
		__extends(Page, _super);
		function Page() {
			/// <summary> Represents the Page class. [Api set:  1.1] </summary>
			/// <field name="context" type="Visio.RequestContext">The request context associated with this object.</field>
			/// <field name="isNull" type="Boolean">Returns a boolean value for whether the corresponding object is null. You must call "context.sync()" before reading the isNull property.</field>
			/// <field name="allShapes" type="Visio.ShapeCollection">All shapes in the Page, including subshapes. Read-only. [Api set:  1.1]</field>
			/// <field name="comments" type="Visio.CommentCollection">Returns the Comments Collection.  Read-only. [Api set:  1.1]</field>
			/// <field name="height" type="Number">Returns the height of the page. Read-only. [Api set:  1.1]</field>
			/// <field name="index" type="Number">Index of the Page. Read-only. [Api set:  1.1]</field>
			/// <field name="isBackground" type="Boolean">Whether the page is a background page or not. Read-only. [Api set:  1.1]</field>
			/// <field name="name" type="String">Page name. Read-only. [Api set:  1.1]</field>
			/// <field name="shapes" type="Visio.ShapeCollection">All top-level shapes in the Page.Read-only. [Api set:  1.1]</field>
			/// <field name="view" type="Visio.PageView">Returns the view of the page. Read-only. [Api set:  1.1]</field>
			/// <field name="width" type="Number">Returns the width of the page. Read-only. [Api set:  1.1]</field>
		}

		Page.prototype.load = function(option) {
			/// <summary>
			/// Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
			/// </summary>
			/// <param name="option" type="string | string[] | OfficeExtension.LoadOption"/>
			/// <returns type="Visio.Page"/>
		}

		Page.prototype.set = function() {
			/// <signature>
			/// <summary>Sets multiple properties on the object at the same time, based on JSON input.</summary>
			/// <param name="properties" type="Visio.Interfaces.PageUpdateData">Properties described by the Visio.Interfaces.PageUpdateData interface.</param>
			/// <param name="options" type="string">Options of the form { throwOnReadOnly?: boolean }
			/// <br />
			/// * throwOnReadOnly: Throw an error if the passed-in property list includes read-only properties (default = true).
			/// </param>
			/// </signature>
			/// <signature>
			/// <summary>Sets multiple properties on the object at the same time, based on an existing loaded object.</summary>
			/// <param name="properties" type="Page">An existing Page object, with properties that have already been loaded and synced.</param>
			/// </signature>
		}
		Page.prototype.activate = function() {
			/// <summary>
			/// Set the page as Active Page of the document. [Api set:  1.1]
			/// </summary>
			/// <returns ></returns>
		}

		return Page;
	})(OfficeExtension.ClientObject);
	Visio.Page = Page;
})(Visio || (Visio = {__proto__: null}));

var Visio;
(function (Visio) {
	var PageCollection = (function(_super) {
		__extends(PageCollection, _super);
		function PageCollection() {
			/// <summary> Represents a collection of Page objects that are part of the document. [Api set:  1.1] </summary>
			/// <field name="context" type="Visio.RequestContext">The request context associated with this object.</field>
			/// <field name="isNull" type="Boolean">Returns a boolean value for whether the corresponding object is null. You must call "context.sync()" before reading the isNull property.</field>
			/// <field name="items" type="Array" elementType="Visio.Page">Gets the loaded child items in this collection.</field>
		}

		PageCollection.prototype.load = function(option) {
			/// <summary>
			/// Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
			/// </summary>
			/// <param name="option" type="string | string[] | OfficeExtension.LoadOption"/>
			/// <returns type="Visio.PageCollection"/>
		}
		PageCollection.prototype.getCount = function() {
			/// <summary>
			/// Gets the number of pages in the collection. [Api set:  1.1]
			/// </summary>
			/// <returns type="OfficeExtension.ClientResult&lt;number&gt;"></returns>
			var result = new OfficeExtension.ClientResult();
			result.__proto__ = null;
			result.value = 0;
			return result;
		}
		PageCollection.prototype.getItem = function(key) {
			/// <summary>
			/// Gets a page using its key (name or Id). [Api set:  1.1]
			/// </summary>
			/// <param name="key" >Key is the name or Id of the page to be retrieved.</param>
			/// <returns type="Visio.Page"></returns>
		}

		return PageCollection;
	})(OfficeExtension.ClientObject);
	Visio.PageCollection = PageCollection;
})(Visio || (Visio = {__proto__: null}));

var Visio;
(function (Visio) {
	var Interfaces;
	(function (Interfaces) {
		var PageLoadCompleteEventArgs = (function() {
			function PageLoadCompleteEventArgs() {
				/// <summary> Provides information about the page that raised the PageLoadComplete event. [Api set:  1.1] </summary>
				/// <field name="pageName" type="String">Gets the name of the page that raised the PageLoad event. [Api set:  1.1]</field>
				/// <field name="success" type="Boolean">Gets the success or failure of the PageLoadComplete event. [Api set:  1.1]</field>
			}
			return PageLoadCompleteEventArgs;
		})();
		Interfaces.PageLoadCompleteEventArgs.__proto__ = null;
		Interfaces.PageLoadCompleteEventArgs = PageLoadCompleteEventArgs;
	})(Interfaces = Visio.Interfaces || (Visio.Interfaces = { __proto__: null}));
})(Visio || (Visio = {__proto__: null}));

var Visio;
(function (Visio) {
	var Interfaces;
	(function (Interfaces) {
		var PageRenderCompleteEventArgs = (function() {
			function PageRenderCompleteEventArgs() {
				/// <summary> Provides information about the page that raised the PageRenderComplete event. [Api set:  1.1] </summary>
				/// <field name="pageName" type="String">Gets the name of the page that raised the PageLoad event. [Api set:  1.1]</field>
				/// <field name="success" type="Boolean">Gets the success/failure of the PageRender event. [Api set:  1.1]</field>
			}
			return PageRenderCompleteEventArgs;
		})();
		Interfaces.PageRenderCompleteEventArgs.__proto__ = null;
		Interfaces.PageRenderCompleteEventArgs = PageRenderCompleteEventArgs;
	})(Interfaces = Visio.Interfaces || (Visio.Interfaces = { __proto__: null}));
})(Visio || (Visio = {__proto__: null}));

var Visio;
(function (Visio) {
	var PageView = (function(_super) {
		__extends(PageView, _super);
		function PageView() {
			/// <summary> Represents the PageView class. [Api set:  1.1] </summary>
			/// <field name="context" type="Visio.RequestContext">The request context associated with this object.</field>
			/// <field name="isNull" type="Boolean">Returns a boolean value for whether the corresponding object is null. You must call "context.sync()" before reading the isNull property.</field>
			/// <field name="zoom" type="Number">Get and set Page&apos;s Zoom level. The value can be between 10 and 400 and denotes the percentage of zoom. [Api set:  1.1]</field>
		}

		PageView.prototype.load = function(option) {
			/// <summary>
			/// Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
			/// </summary>
			/// <param name="option" type="string | string[] | OfficeExtension.LoadOption"/>
			/// <returns type="Visio.PageView"/>
		}

		PageView.prototype.set = function() {
			/// <signature>
			/// <summary>Sets multiple properties on the object at the same time, based on JSON input.</summary>
			/// <param name="properties" type="Visio.Interfaces.PageViewUpdateData">Properties described by the Visio.Interfaces.PageViewUpdateData interface.</param>
			/// <param name="options" type="string">Options of the form { throwOnReadOnly?: boolean }
			/// <br />
			/// * throwOnReadOnly: Throw an error if the passed-in property list includes read-only properties (default = true).
			/// </param>
			/// </signature>
			/// <signature>
			/// <summary>Sets multiple properties on the object at the same time, based on an existing loaded object.</summary>
			/// <param name="properties" type="PageView">An existing PageView object, with properties that have already been loaded and synced.</param>
			/// </signature>
		}
		PageView.prototype.centerViewportOnShape = function(ShapeId) {
			/// <summary>
			/// Pans the Visio drawing to place the specified shape in the center of the view. [Api set:  1.1]
			/// </summary>
			/// <param name="ShapeId" type="Number">ShapeId to be seen in the center.</param>
			/// <returns ></returns>
		}
		PageView.prototype.fitToWindow = function() {
			/// <summary>
			/// Fit Page to current window. [Api set:  1.1]
			/// </summary>
			/// <returns ></returns>
		}
		PageView.prototype.getPosition = function() {
			/// <summary>
			/// Returns the position object that specifies the position of the page in the view. [Api set:  1.1]
			/// </summary>
			/// <returns type="OfficeExtension.ClientResult&lt;Visio.Position&gt;"></returns>
			var result = new OfficeExtension.ClientResult();
			result.__proto__ = null;
			result.value = {};
			return result;
		}
		PageView.prototype.getSelection = function() {
			/// <summary>
			/// Represents the Selection in the page. [Api set:  1.1]
			/// </summary>
			/// <returns type="Visio.Selection"></returns>
		}
		PageView.prototype.isShapeInViewport = function(Shape) {
			/// <summary>
			/// To check if the shape is in view of the page or not. [Api set:  1.1]
			/// </summary>
			/// <param name="Shape" type="Visio.Shape">Shape to be checked.</param>
			/// <returns type="OfficeExtension.ClientResult&lt;boolean&gt;"></returns>
			var result = new OfficeExtension.ClientResult();
			result.__proto__ = null;
			result.value = false;
			return result;
		}
		PageView.prototype.setPosition = function(Position) {
			/// <summary>
			/// Sets the position of the page in the view. [Api set:  1.1]
			/// </summary>
			/// <param name="Position" type="Visio.Interfaces.Position">Position object that specifies the new position of the page in the view.</param>
			/// <returns ></returns>
		}

		return PageView;
	})(OfficeExtension.ClientObject);
	Visio.PageView = PageView;
})(Visio || (Visio = {__proto__: null}));

var Visio;
(function (Visio) {
	var Interfaces;
	(function (Interfaces) {
		var Position = (function() {
			function Position() {
				/// <summary> Represents the Position of the object in the view. [Api set:  1.1] </summary>
				/// <field name="x" type="Number">An integer that specifies the x-coordinate of the object, which is the signed value of the distance in pixels from the viewport&apos;s center to the left boundary of the page. [Api set:  1.1]</field>
				/// <field name="y" type="Number">An integer that specifies the y-coordinate of the object, which is the signed value of the distance in pixels from the viewport&apos;s center to the top boundary of the page. [Api set:  1.1]</field>
			}
			return Position;
		})();
		Interfaces.Position.__proto__ = null;
		Interfaces.Position = Position;
	})(Interfaces = Visio.Interfaces || (Visio.Interfaces = { __proto__: null}));
})(Visio || (Visio = {__proto__: null}));

var Visio;
(function (Visio) {
	var Selection = (function(_super) {
		__extends(Selection, _super);
		function Selection() {
			/// <summary> Represents the Selection in the page. [Api set:  1.1] </summary>
			/// <field name="context" type="Visio.RequestContext">The request context associated with this object.</field>
			/// <field name="isNull" type="Boolean">Returns a boolean value for whether the corresponding object is null. You must call "context.sync()" before reading the isNull property.</field>
			/// <field name="shapes" type="Visio.ShapeCollection">Gets the Shapes of the Selection. Read-only. [Api set:  1.1]</field>
		}

		Selection.prototype.load = function(option) {
			/// <summary>
			/// Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
			/// </summary>
			/// <param name="option" type="string | string[] | OfficeExtension.LoadOption"/>
			/// <returns type="Visio.Selection"/>
		}

		return Selection;
	})(OfficeExtension.ClientObject);
	Visio.Selection = Selection;
})(Visio || (Visio = {__proto__: null}));

var Visio;
(function (Visio) {
	var Interfaces;
	(function (Interfaces) {
		var SelectionChangedEventArgs = (function() {
			function SelectionChangedEventArgs() {
				/// <summary> Provides information about the shape collection that raised the SelectionChanged event. [Api set:  1.1] </summary>
				/// <field name="pageName" type="String">Gets the name of the page which has the ShapeCollection object that raised the SelectionChanged event. [Api set:  1.1]</field>
				/// <field name="shapeNames" type="Array" elementType="String">Gets the array of shape names that raised the SelectionChanged event. [Api set:  1.1]</field>
			}
			return SelectionChangedEventArgs;
		})();
		Interfaces.SelectionChangedEventArgs.__proto__ = null;
		Interfaces.SelectionChangedEventArgs = SelectionChangedEventArgs;
	})(Interfaces = Visio.Interfaces || (Visio.Interfaces = { __proto__: null}));
})(Visio || (Visio = {__proto__: null}));

var Visio;
(function (Visio) {
	var Shape = (function(_super) {
		__extends(Shape, _super);
		function Shape() {
			/// <summary> Represents the Shape class. [Api set:  1.1] </summary>
			/// <field name="context" type="Visio.RequestContext">The request context associated with this object.</field>
			/// <field name="isNull" type="Boolean">Returns a boolean value for whether the corresponding object is null. You must call "context.sync()" before reading the isNull property.</field>
			/// <field name="comments" type="Visio.CommentCollection">Returns the Comments Collection. Read-only. [Api set:  1.1]</field>
			/// <field name="hyperlinks" type="Visio.HyperlinkCollection">Returns the Hyperlinks collection for a Shape object. Read-only. [Api set:  1.1]</field>
			/// <field name="id" type="Number">Shape&apos;s identifier. Read-only. [Api set:  1.1]</field>
			/// <field name="name" type="String">Shape&apos;s name. Read-only. [Api set:  1.1]</field>
			/// <field name="select" type="Boolean">Returns true, if shape is selected. User can set true to select the shape explicitly. [Api set:  1.1]</field>
			/// <field name="shapeDataItems" type="Visio.ShapeDataItemCollection">Returns the Shape&apos;s Data Section. Read-only. [Api set:  1.1]</field>
			/// <field name="subShapes" type="Visio.ShapeCollection">Gets SubShape Collection. Read-only. [Api set:  1.1]</field>
			/// <field name="text" type="String">Shape&apos;s text. Read-only. [Api set:  1.1]</field>
			/// <field name="view" type="Visio.ShapeView">Returns the view of the shape. Read-only. [Api set:  1.1]</field>
		}

		Shape.prototype.load = function(option) {
			/// <summary>
			/// Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
			/// </summary>
			/// <param name="option" type="string | string[] | OfficeExtension.LoadOption"/>
			/// <returns type="Visio.Shape"/>
		}

		Shape.prototype.set = function() {
			/// <signature>
			/// <summary>Sets multiple properties on the object at the same time, based on JSON input.</summary>
			/// <param name="properties" type="Visio.Interfaces.ShapeUpdateData">Properties described by the Visio.Interfaces.ShapeUpdateData interface.</param>
			/// <param name="options" type="string">Options of the form { throwOnReadOnly?: boolean }
			/// <br />
			/// * throwOnReadOnly: Throw an error if the passed-in property list includes read-only properties (default = true).
			/// </param>
			/// </signature>
			/// <signature>
			/// <summary>Sets multiple properties on the object at the same time, based on an existing loaded object.</summary>
			/// <param name="properties" type="Shape">An existing Shape object, with properties that have already been loaded and synced.</param>
			/// </signature>
		}
		Shape.prototype.getBounds = function() {
			/// <summary>
			/// Returns the BoundingBox object that specifies bounding box of the shape. [Api set:  1.1]
			/// </summary>
			/// <returns type="OfficeExtension.ClientResult&lt;Visio.BoundingBox&gt;"></returns>
			var result = new OfficeExtension.ClientResult();
			result.__proto__ = null;
			result.value = {};
			return result;
		}

		return Shape;
	})(OfficeExtension.ClientObject);
	Visio.Shape = Shape;
})(Visio || (Visio = {__proto__: null}));

var Visio;
(function (Visio) {
	var Interfaces;
	(function (Interfaces) {
		var ShapeBinding = (function() {
			function ShapeBinding() {
				/// <summary> Shape binding informations required for data visualizer diagram [Api set:  1.1] </summary>
			}
			return ShapeBinding;
		})();
		Interfaces.ShapeBinding.__proto__ = null;
		Interfaces.ShapeBinding = ShapeBinding;
	})(Interfaces = Visio.Interfaces || (Visio.Interfaces = { __proto__: null}));
})(Visio || (Visio = {__proto__: null}));

var Visio;
(function (Visio) {
	var ShapeCollection = (function(_super) {
		__extends(ShapeCollection, _super);
		function ShapeCollection() {
			/// <summary> Represents the Shape Collection. [Api set:  1.1] </summary>
			/// <field name="context" type="Visio.RequestContext">The request context associated with this object.</field>
			/// <field name="isNull" type="Boolean">Returns a boolean value for whether the corresponding object is null. You must call "context.sync()" before reading the isNull property.</field>
			/// <field name="items" type="Array" elementType="Visio.Shape">Gets the loaded child items in this collection.</field>
		}

		ShapeCollection.prototype.load = function(option) {
			/// <summary>
			/// Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
			/// </summary>
			/// <param name="option" type="string | string[] | OfficeExtension.LoadOption"/>
			/// <returns type="Visio.ShapeCollection"/>
		}
		ShapeCollection.prototype.getCount = function() {
			/// <summary>
			/// Gets the number of Shapes in the collection. [Api set:  1.1]
			/// </summary>
			/// <returns type="OfficeExtension.ClientResult&lt;number&gt;"></returns>
			var result = new OfficeExtension.ClientResult();
			result.__proto__ = null;
			result.value = 0;
			return result;
		}
		ShapeCollection.prototype.getItem = function(key) {
			/// <summary>
			/// Gets a Shape using its key (name or Index). [Api set:  1.1]
			/// </summary>
			/// <param name="key" >Key is the Name or Index of the shape to be retrieved.</param>
			/// <returns type="Visio.Shape"></returns>
		}

		return ShapeCollection;
	})(OfficeExtension.ClientObject);
	Visio.ShapeCollection = ShapeCollection;
})(Visio || (Visio = {__proto__: null}));

var Visio;
(function (Visio) {
	var ShapeDataItem = (function(_super) {
		__extends(ShapeDataItem, _super);
		function ShapeDataItem() {
			/// <summary> Represents the ShapeDataItem. [Api set:  1.1] </summary>
			/// <field name="context" type="Visio.RequestContext">The request context associated with this object.</field>
			/// <field name="isNull" type="Boolean">Returns a boolean value for whether the corresponding object is null. You must call "context.sync()" before reading the isNull property.</field>
			/// <field name="format" type="String">A string that specifies the format of the shape data item. Read-only. [Api set:  1.1]</field>
			/// <field name="formattedValue" type="String">A string that specifies the formatted value of the shape data item. Read-only. [Api set:  1.1]</field>
			/// <field name="label" type="String">A string that specifies the label of the shape data item. Read-only. [Api set:  1.1]</field>
			/// <field name="value" type="String">A string that specifies the value of the shape data item. Read-only. [Api set:  1.1]</field>
		}

		ShapeDataItem.prototype.load = function(option) {
			/// <summary>
			/// Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
			/// </summary>
			/// <param name="option" type="string | string[] | OfficeExtension.LoadOption"/>
			/// <returns type="Visio.ShapeDataItem"/>
		}

		return ShapeDataItem;
	})(OfficeExtension.ClientObject);
	Visio.ShapeDataItem = ShapeDataItem;
})(Visio || (Visio = {__proto__: null}));

var Visio;
(function (Visio) {
	var ShapeDataItemCollection = (function(_super) {
		__extends(ShapeDataItemCollection, _super);
		function ShapeDataItemCollection() {
			/// <summary> Represents the ShapeDataItemCollection for a given Shape. [Api set:  1.1] </summary>
			/// <field name="context" type="Visio.RequestContext">The request context associated with this object.</field>
			/// <field name="isNull" type="Boolean">Returns a boolean value for whether the corresponding object is null. You must call "context.sync()" before reading the isNull property.</field>
			/// <field name="items" type="Array" elementType="Visio.ShapeDataItem">Gets the loaded child items in this collection.</field>
		}

		ShapeDataItemCollection.prototype.load = function(option) {
			/// <summary>
			/// Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
			/// </summary>
			/// <param name="option" type="string | string[] | OfficeExtension.LoadOption"/>
			/// <returns type="Visio.ShapeDataItemCollection"/>
		}
		ShapeDataItemCollection.prototype.getCount = function() {
			/// <summary>
			/// Gets the number of Shape Data Items. [Api set:  1.1]
			/// </summary>
			/// <returns type="OfficeExtension.ClientResult&lt;number&gt;"></returns>
			var result = new OfficeExtension.ClientResult();
			result.__proto__ = null;
			result.value = 0;
			return result;
		}
		ShapeDataItemCollection.prototype.getItem = function(key) {
			/// <summary>
			/// Gets the ShapeDataItem using its name. [Api set:  1.1]
			/// </summary>
			/// <param name="key" >Key is the name of the ShapeDataItem to be retrieved.</param>
			/// <returns type="Visio.ShapeDataItem"></returns>
		}

		return ShapeDataItemCollection;
	})(OfficeExtension.ClientObject);
	Visio.ShapeDataItemCollection = ShapeDataItemCollection;
})(Visio || (Visio = {__proto__: null}));

var Visio;
(function (Visio) {
	var Interfaces;
	(function (Interfaces) {
		var ShapeMouseEnterEventArgs = (function() {
			function ShapeMouseEnterEventArgs() {
				/// <summary> Provides information about the shape that raised the ShapeMouseEnter event. [Api set:  1.1] </summary>
				/// <field name="pageName" type="String">Gets the name of the page which has the shape object that raised the ShapeMouseEnter event. [Api set:  1.1]</field>
				/// <field name="shapeName" type="String">Gets the name of the shape object that raised the ShapeMouseEnter event. [Api set:  1.1]</field>
			}
			return ShapeMouseEnterEventArgs;
		})();
		Interfaces.ShapeMouseEnterEventArgs.__proto__ = null;
		Interfaces.ShapeMouseEnterEventArgs = ShapeMouseEnterEventArgs;
	})(Interfaces = Visio.Interfaces || (Visio.Interfaces = { __proto__: null}));
})(Visio || (Visio = {__proto__: null}));

var Visio;
(function (Visio) {
	var Interfaces;
	(function (Interfaces) {
		var ShapeMouseLeaveEventArgs = (function() {
			function ShapeMouseLeaveEventArgs() {
				/// <summary> Provides information about the shape that raised the ShapeMouseLeave event. [Api set:  1.1] </summary>
				/// <field name="pageName" type="String">Gets the name of the page which has the shape object that raised the ShapeMouseLeave event. [Api set:  1.1]</field>
				/// <field name="shapeName" type="String">Gets the name of the shape object that raised the ShapeMouseLeave event. [Api set:  1.1]</field>
			}
			return ShapeMouseLeaveEventArgs;
		})();
		Interfaces.ShapeMouseLeaveEventArgs.__proto__ = null;
		Interfaces.ShapeMouseLeaveEventArgs = ShapeMouseLeaveEventArgs;
	})(Interfaces = Visio.Interfaces || (Visio.Interfaces = { __proto__: null}));
})(Visio || (Visio = {__proto__: null}));

var Visio;
(function (Visio) {
	var ShapeView = (function(_super) {
		__extends(ShapeView, _super);
		function ShapeView() {
			/// <summary> Represents the ShapeView class. [Api set:  1.1] </summary>
			/// <field name="context" type="Visio.RequestContext">The request context associated with this object.</field>
			/// <field name="isNull" type="Boolean">Returns a boolean value for whether the corresponding object is null. You must call "context.sync()" before reading the isNull property.</field>
			/// <field name="highlight" type="Visio.Interfaces.Highlight">Represents the highlight around the shape. [Api set:  1.1]</field>
		}

		ShapeView.prototype.load = function(option) {
			/// <summary>
			/// Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
			/// </summary>
			/// <param name="option" type="string | string[] | OfficeExtension.LoadOption"/>
			/// <returns type="Visio.ShapeView"/>
		}

		ShapeView.prototype.set = function() {
			/// <signature>
			/// <summary>Sets multiple properties on the object at the same time, based on JSON input.</summary>
			/// <param name="properties" type="Visio.Interfaces.ShapeViewUpdateData">Properties described by the Visio.Interfaces.ShapeViewUpdateData interface.</param>
			/// <param name="options" type="string">Options of the form { throwOnReadOnly?: boolean }
			/// <br />
			/// * throwOnReadOnly: Throw an error if the passed-in property list includes read-only properties (default = true).
			/// </param>
			/// </signature>
			/// <signature>
			/// <summary>Sets multiple properties on the object at the same time, based on an existing loaded object.</summary>
			/// <param name="properties" type="ShapeView">An existing ShapeView object, with properties that have already been loaded and synced.</param>
			/// </signature>
		}
		ShapeView.prototype.addOverlay = function(OverlayType, Content, OverlayHorizontalAlignment, OverlayVerticalAlignment, Width, Height) {
			/// <summary>
			/// Adds an overlay on top of the shape. [Api set:  1.1]
			/// </summary>
			/// <param name="OverlayType" type="String">An Overlay Type. Can be &apos;Text&apos;, &apos;Image&apos; or &apos;Html&apos;.</param>
			/// <param name="Content" type="String">Content of Overlay.</param>
			/// <param name="OverlayHorizontalAlignment" type="String">Horizontal Alignment of Overlay. Can be &apos;Left&apos;, &apos;Center&apos;, or &apos;Right&apos;.</param>
			/// <param name="OverlayVerticalAlignment" type="String">Vertical Alignment of Overlay. Can be &apos;Top&apos;, &apos;Middle&apos;, &apos;Bottom&apos;.</param>
			/// <param name="Width" type="Number">Overlay Width.</param>
			/// <param name="Height" type="Number">Overlay Height.</param>
			/// <returns type="OfficeExtension.ClientResult&lt;number&gt;"></returns>
			var result = new OfficeExtension.ClientResult();
			result.__proto__ = null;
			result.value = 0;
			return result;
		}
		ShapeView.prototype.removeOverlay = function(OverlayId) {
			/// <summary>
			/// Removes particular overlay or all overlays on the Shape. [Api set:  1.1]
			/// </summary>
			/// <param name="OverlayId" type="Number">An Overlay Id. Removes the specific overlay id from the shape.</param>
			/// <returns ></returns>
		}
		ShapeView.prototype.setText = function(Text) {
			/// <summary>
			/// The purpose of SetText API is to update the text inside a visio Shape in run time. The updated text retains the existing formatting properties of the shape's text. [Api set:  1.1]
			/// </summary>
			/// <param name="Text" type="String">Text parameter is the 'Updated the text to display on the shape'</param>
			/// <returns ></returns>
		}
		ShapeView.prototype.showOverlay = function(overlayId, show) {
			/// <summary>
			/// Shows particular overlay on the Shape. [Api set:  1.1]
			/// </summary>
			/// <param name="overlayId" type="Number">overlay id in context</param>
			/// <param name="show" type="Boolean">to show or hide</param>
			/// <returns ></returns>
		}

		return ShapeView;
	})(OfficeExtension.ClientObject);
	Visio.ShapeView = ShapeView;
})(Visio || (Visio = {__proto__: null}));

var Visio;
(function (Visio) {
	var Interfaces;
	(function (Interfaces) {
		var TaskPaneStateChangedEventArgs = (function() {
			function TaskPaneStateChangedEventArgs() {
				/// <summary> Provides information about the TaskPaneStateChanged event. [Api set:  1.1] </summary>
				/// <field name="isVisible" type="Boolean">Current state of the taskpane [Api set:  1.1]</field>
				/// <field name="paneType" type="String">Type of the TaskPane. [Api set:  1.1]</field>
			}
			return TaskPaneStateChangedEventArgs;
		})();
		Interfaces.TaskPaneStateChangedEventArgs.__proto__ = null;
		Interfaces.TaskPaneStateChangedEventArgs = TaskPaneStateChangedEventArgs;
	})(Interfaces = Visio.Interfaces || (Visio.Interfaces = { __proto__: null}));
})(Visio || (Visio = {__proto__: null}));

var Visio;
(function (Visio) {
	/// <summary> TaskPaneType represents the types of the First Party TaskPanes that are supported by Host through APIs. Used in case of Show TaskPane API/ TaskPane State Changed Event etc [Api set:  1.1] </summary>
	var TaskPaneType = {
		__proto__: null,
		"none": "none",
		"dataVisualizerProcessMappings": "dataVisualizerProcessMappings",
		"dataVisualizerOrgChartMappings": "dataVisualizerOrgChartMappings",
	}
	Visio.TaskPaneType = TaskPaneType;
})(Visio || (Visio = {__proto__: null}));

var Visio;
(function (Visio) {
	/// <summary> Toolbar IDs of the app [Api set:  1.1] </summary>
	var ToolBarType = {
		__proto__: null,
		"commandBar": "commandBar",
		"pageNavigationBar": "pageNavigationBar",
		"statusBar": "statusBar",
	}
	Visio.ToolBarType = ToolBarType;
})(Visio || (Visio = {__proto__: null}));

var Visio;
(function (Visio) {
	var Interfaces;
	(function (Interfaces) {
		var ApplicationUpdateData = (function() {
			function ApplicationUpdateData() {
				/// <summary>An interface for updating data on the Application object, for use in "application.set({ ... })".</summary>
				/// <field name="showBorders" type="Boolean">Show or hide the iFrame application borders. [Api set:  1.1]</field>;
				/// <field name="showToolbars" type="Boolean">Show or hide the standard toolbars. [Api set:  1.1]</field>;
			}
			return ApplicationUpdateData;
		})();
		Interfaces.ApplicationUpdateData.__proto__ = null;
		Interfaces.ApplicationUpdateData = ApplicationUpdateData;
	})(Interfaces = Visio.Interfaces || (Visio.Interfaces = { __proto__: null}));
})(Visio || (Visio = {__proto__: null}));

var Visio;
(function (Visio) {
	var Interfaces;
	(function (Interfaces) {
		var DocumentUpdateData = (function() {
			function DocumentUpdateData() {
				/// <summary>An interface for updating data on the Document object, for use in "document.set({ ... })".</summary>
				/// <field name="application" type="Visio.Interfaces.ApplicationUpdateData">Represents a Visio application instance that contains this document. [Api set:  1.1]</field>
				/// <field name="view" type="Visio.Interfaces.DocumentViewUpdateData">Returns the DocumentView object. [Api set:  1.1]</field>
			}
			return DocumentUpdateData;
		})();
		Interfaces.DocumentUpdateData.__proto__ = null;
		Interfaces.DocumentUpdateData = DocumentUpdateData;
	})(Interfaces = Visio.Interfaces || (Visio.Interfaces = { __proto__: null}));
})(Visio || (Visio = {__proto__: null}));

var Visio;
(function (Visio) {
	var Interfaces;
	(function (Interfaces) {
		var DocumentViewUpdateData = (function() {
			function DocumentViewUpdateData() {
				/// <summary>An interface for updating data on the DocumentView object, for use in "documentView.set({ ... })".</summary>
				/// <field name="disableHyperlinks" type="Boolean">Disable Hyperlinks. [Api set:  1.1]</field>;
				/// <field name="disablePan" type="Boolean">Disable Pan. [Api set:  1.1]</field>;
				/// <field name="disablePanZoomWindow" type="Boolean">Disable PanZoomWindow. [Api set:  1.1]</field>;
				/// <field name="disableZoom" type="Boolean">Disable Zoom. [Api set:  1.1]</field>;
				/// <field name="hideDiagramBoundary" type="Boolean">Hide Diagram Boundary. [Api set:  1.1]</field>;
			}
			return DocumentViewUpdateData;
		})();
		Interfaces.DocumentViewUpdateData.__proto__ = null;
		Interfaces.DocumentViewUpdateData = DocumentViewUpdateData;
	})(Interfaces = Visio.Interfaces || (Visio.Interfaces = { __proto__: null}));
})(Visio || (Visio = {__proto__: null}));

var Visio;
(function (Visio) {
	var Interfaces;
	(function (Interfaces) {
		var PageUpdateData = (function() {
			function PageUpdateData() {
				/// <summary>An interface for updating data on the Page object, for use in "page.set({ ... })".</summary>
				/// <field name="view" type="Visio.Interfaces.PageViewUpdateData">Returns the view of the page. [Api set:  1.1]</field>
			}
			return PageUpdateData;
		})();
		Interfaces.PageUpdateData.__proto__ = null;
		Interfaces.PageUpdateData = PageUpdateData;
	})(Interfaces = Visio.Interfaces || (Visio.Interfaces = { __proto__: null}));
})(Visio || (Visio = {__proto__: null}));

var Visio;
(function (Visio) {
	var Interfaces;
	(function (Interfaces) {
		var PageViewUpdateData = (function() {
			function PageViewUpdateData() {
				/// <summary>An interface for updating data on the PageView object, for use in "pageView.set({ ... })".</summary>
				/// <field name="zoom" type="Number">Get and set Page&apos;s Zoom level. The value can be between 10 and 400 and denotes the percentage of zoom. [Api set:  1.1]</field>;
			}
			return PageViewUpdateData;
		})();
		Interfaces.PageViewUpdateData.__proto__ = null;
		Interfaces.PageViewUpdateData = PageViewUpdateData;
	})(Interfaces = Visio.Interfaces || (Visio.Interfaces = { __proto__: null}));
})(Visio || (Visio = {__proto__: null}));

var Visio;
(function (Visio) {
	var Interfaces;
	(function (Interfaces) {
		var ShapeUpdateData = (function() {
			function ShapeUpdateData() {
				/// <summary>An interface for updating data on the Shape object, for use in "shape.set({ ... })".</summary>
				/// <field name="view" type="Visio.Interfaces.ShapeViewUpdateData">Returns the view of the shape. [Api set:  1.1]</field>
				/// <field name="select" type="Boolean">Returns true, if shape is selected. User can set true to select the shape explicitly. [Api set:  1.1]</field>;
			}
			return ShapeUpdateData;
		})();
		Interfaces.ShapeUpdateData.__proto__ = null;
		Interfaces.ShapeUpdateData = ShapeUpdateData;
	})(Interfaces = Visio.Interfaces || (Visio.Interfaces = { __proto__: null}));
})(Visio || (Visio = {__proto__: null}));

var Visio;
(function (Visio) {
	var Interfaces;
	(function (Interfaces) {
		var ShapeViewUpdateData = (function() {
			function ShapeViewUpdateData() {
				/// <summary>An interface for updating data on the ShapeView object, for use in "shapeView.set({ ... })".</summary>
				/// <field name="highlight" type="Visio.Interfaces.Highlight">Represents the highlight around the shape. [Api set:  1.1]</field>;
			}
			return ShapeViewUpdateData;
		})();
		Interfaces.ShapeViewUpdateData.__proto__ = null;
		Interfaces.ShapeViewUpdateData = ShapeViewUpdateData;
	})(Interfaces = Visio.Interfaces || (Visio.Interfaces = { __proto__: null}));
})(Visio || (Visio = {__proto__: null}));

var Visio;
(function (Visio) {
	var Interfaces;
	(function (Interfaces) {
		var CommentUpdateData = (function() {
			function CommentUpdateData() {
				/// <summary>An interface for updating data on the Comment object, for use in "comment.set({ ... })".</summary>
				/// <field name="author" type="String">A string that specifies the name of the author of the comment. [Api set:  1.1]</field>;
				/// <field name="date" type="String">A string that specifies the date when the comment was created. [Api set:  1.1]</field>;
				/// <field name="text" type="String">A string that contains the comment text. [Api set:  1.1]</field>;
			}
			return CommentUpdateData;
		})();
		Interfaces.CommentUpdateData.__proto__ = null;
		Interfaces.CommentUpdateData = CommentUpdateData;
	})(Interfaces = Visio.Interfaces || (Visio.Interfaces = { __proto__: null}));
})(Visio || (Visio = {__proto__: null}));

var Visio;
(function (Visio) {
	var Interfaces;
	(function (Interfaces) {
		var DataVisualizerDiagramUpdateData = (function() {
			function DataVisualizerDiagramUpdateData() {
				/// <summary>An interface for updating data on the DataVisualizerDiagram object, for use in "dataVisualizerDiagram.set({ ... })".</summary>
				/// <field name="page" type="Visio.Interfaces.PageUpdateData">Returns the page object that is associated with this diagram object. [Api set:  1.1]</field>
			}
			return DataVisualizerDiagramUpdateData;
		})();
		Interfaces.DataVisualizerDiagramUpdateData.__proto__ = null;
		Interfaces.DataVisualizerDiagramUpdateData = DataVisualizerDiagramUpdateData;
	})(Interfaces = Visio.Interfaces || (Visio.Interfaces = { __proto__: null}));
})(Visio || (Visio = {__proto__: null}));
var Visio;
(function (Visio) {
	var RequestContext = (function (_super) {
		__extends(RequestContext, _super);
		function RequestContext() {
			/// <summary>
			/// The RequestContext object facilitates requests to the Visio application. Since the Office add-in and the Visio application run in two different processes, the request context is required to get access to the Visio object model from the add-in.
			/// </summary>
				/// <field name="document" type="Visio.Document">Root object for interacting with the document</field>
			_super.call(this, null);
		}
		return RequestContext;
	})(OfficeExtension.ClientRequestContext);
	Visio.RequestContext = RequestContext;

	Visio.run = function (batch) {
		/// <signature>
		/// <summary>
		/// Executes a batch script that performs actions on the Visio object model, using a new RequestContext. When the promise is resolved, any tracked objects that were automatically allocated during execution will be released.
		/// </summary>
		/// <param name="batch" type="function(context) { ... }">
		/// A function that takes in a RequestContext and returns a promise (typically, just the result of "context.sync()").
		/// <br />
		/// The context parameter facilitates requests to the Visio application. Since the Office add-in and the Visio application run in two different processes, the RequestContext is required to get access to the Visio object model from the add-in.
		/// </param>
		/// </signature>
		/// <signature>
		/// <summary>
		/// Executes a batch script that performs actions on the Visio object model, using the RequestContext of a previously-created API object. When the promise is resolved, any tracked objects that were automatically allocated during execution will be released.
		/// </summary>
		/// <param name="object" type="OfficeExtension.ClientObject">
		/// A previously-created API object. The batch will use the same RequestContext as the passed-in object, which means that any changes applied to the object will be picked up by "context.sync()".
		/// </param>
		/// <param name="batch" type="function(context) { ... }">
		/// A function that takes in a RequestContext and returns a promise (typically, just the result of "context.sync()").
		/// <br />
		/// The context parameter facilitates requests to the Visio application. Since the Office add-in and the Visio application run in two different processes, the RequestContext is required to get access to the Visio object model from the add-in.
		/// </param>
		/// </signature>
		/// <signature>
		/// <summary>
		/// Executes a batch script that performs actions on the Visio object model, using the RequestContext of a previously-created API object. When the promise is resolved, any tracked objects that were automatically allocated during execution will be released.
		/// </summary>
		/// <param name="objects" type="Array&lt;OfficeExtension.ClientObject&gt;">
		/// An array of previously-created API objects. The array will be validated to make sure that all of the objects share the same context. The batch will use this shared RequestContext, which means that any changes applied to these objects will be picked up by "context.sync()".
		/// </param>
		/// <param name="batch" type="function(context) { ... }">
		/// A function that takes in a RequestContext and returns a promise (typically, just the result of "context.sync()").
		/// <br />
		/// The context parameter facilitates requests to the Visio application. Since the Office add-in and the Visio application run in two different processes, the RequestContext is required to get access to the Visio object model from the add-in.
		/// </param>
		/// </signature>
		arguments[arguments.length - 1](new Visio.RequestContext());
		return new OfficeExtension.Promise();
	}
})(Visio || (Visio = {__proto__: null}));
Visio.__proto__ = null;


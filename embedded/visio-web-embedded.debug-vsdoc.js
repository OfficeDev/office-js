
var Visio;
(function (Visio) {
	var Application = (function(_super) {
		__extends(Application, _super);
		function Application() {
			/// <summary> Represents the Application. [Api set:  1.1] </summary>
			/// <field name="context" type="Visio.RequestContext">The request context associated with this object.</field>
			/// <field name="isNull" type="Boolean">Returns a boolean value for whether the corresponding object is null. You must call "context.sync()" before reading the isNull property.</field>
			/// <field name="showToolbars" type="Boolean">Show or Hide the standard toolbars. [Api set:  1.1]</field>
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
	var Comment = (function(_super) {
		__extends(Comment, _super);
		function Comment() {
			/// <summary> Represents the Comment. [Api set:  1.1] </summary>
			/// <field name="context" type="Visio.RequestContext">The request context associated with this object.</field>
			/// <field name="isNull" type="Boolean">Returns a boolean value for whether the corresponding object is null. You must call "context.sync()" before reading the isNull property.</field>
			/// <field name="author" type="String">A string that specifies the label of the shape data item. [Api set:  1.1]</field>
			/// <field name="date" type="String">A string that specifies the format of the shape data item. [Api set:  1.1]</field>
			/// <field name="text" type="String">A string that specifies the value of the shape data item. [Api set:  1.1]</field>
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
			/// Gets the number of Shape Data Items. [Api set:  1.1]
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
		var DataRefreshCompleteEventArgs = (function() {
			function DataRefreshCompleteEventArgs() {
				/// <summary> Provides information about the document that raised the DataRefreshComplete event. [Api set:  1.1] </summary>
				/// <field name="document" type="Visio.Document">Gets the document object that raised the DataRefreshComplete event. [Api set:  1.1]</field>
				/// <field name="success" type="Boolean">Gets the success/failure of the DataRefreshComplete event. [Api set:  1.1]</field>
			}
			return DataRefreshCompleteEventArgs;
		})();
		Interfaces.DataRefreshCompleteEventArgs.__proto__ = null;
		Interfaces.DataRefreshCompleteEventArgs = DataRefreshCompleteEventArgs;
	})(Interfaces = Visio.Interfaces || (Visio.Interfaces = { __proto__: null}));
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
			/// <field name="view" type="Visio.DocumentView">Returns the DocumentView object. [Api set:  1.1]</field>
			/// <field name="onDataRefreshComplete" type="OfficeExtension.EventHandlers">Occurs when the data is refreshed in the diagram. [Api set:  1.1]</field>
			/// <field name="onPageLoadComplete" type="OfficeExtension.EventHandlers">Occurs when the page is finished loading. [Api set:  1.1]</field>
			/// <field name="onSelectionChanged" type="OfficeExtension.EventHandlers">Occurs when the current selection of shapes changes. [Api set:  1.1]</field>
			/// <field name="onShapeMouseEnter" type="OfficeExtension.EventHandlers">Occurs when the user moves the mouse pointer into the bounding box of a shape. [Api set:  1.1]</field>
			/// <field name="onShapeMouseLeave" type="OfficeExtension.EventHandlers">Occurs when the user moves the mouse out of the bounding box of a shape. [Api set:  1.1]</field>
		}

		Document.prototype.load = function(option) {
			/// <summary>
			/// Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
			/// </summary>
			/// <param name="option" type="string | string[] | OfficeExtension.LoadOption"/>
			/// <returns type="Visio.Document"/>
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
			},
			removeAll: function () {
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
			},
			removeAll: function () {
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
			},
			removeAll: function () {
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
			},
			removeAll: function () {
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
			},
			removeAll: function () {
				return;
			}
		};

		return Document;
	})(OfficeExtension.ClientObject);
	Visio.Document = Document;
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
			/// <field name="disableZoom" type="Boolean">Disable Zoom. [Api set:  1.1]</field>
			/// <field name="hideDiagramBoundary" type="Boolean">Disable Hyperlinks. [Api set:  1.1]</field>
		}

		DocumentView.prototype.load = function(option) {
			/// <summary>
			/// Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
			/// </summary>
			///
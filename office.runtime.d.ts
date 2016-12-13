/** An abstract proxy object that represents an object in an Office document. You create proxy objects from the context (or from other proxy objects), add commands to a queue to act on the object, and then synchronize the proxy object state with the document by calling "context.sync()". */
export declare class ClientObject {
	/** The request context associated with the object */
	context: ClientRequestContext;
	/** Returns a boolean value for whether the corresponding object is a null object. You must call "context.sync()" before reading the isNullObject property. */
	isNullObject: boolean;
}

export declare interface LoadOption {
	select?: string | string[];
	expand?: string | string[];
	top?: number;
	skip?: number;
}

/** An abstract RequestContext object that facilitates requests to the host Office application. The "Excel.run" and "Word.run" methods provide a request context. */
export declare class ClientRequestContext {
	constructor(url?: string);
	/** Collection of objects that are tracked for automatic adjustments based on surrounding changes in the document. */
	trackedObjects: TrackedObjects;
	/** Request headers */
	requestHeaders: { [name: string]: string };
	/** Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties. */
	load(object: ClientObject, option?: string | string[] | LoadOption): void;
	/** Adds a trace message to the queue. If the promise returned by "context.sync()" is rejected due to an error, this adds a ".traceMessages" array to the OfficeExtension.Error object, containing all trace messages that were executed. These messages can help you monitor the program execution sequence and detect the cause of the error. */
	trace(message: string): void;
	/** Synchronizes the state between JavaScript proxy objects and the Office document, by executing instructions queued on the request context and retrieving properties of loaded Office objects for use in your code. This method returns a promise, which is resolved when the synchronization is complete. */
	sync<T>(passThroughValue?: T): IPromise<T>;
}

/** Contains the result for methods that return primitive types. The object's value property is retrieved from the document after "context.sync()" is invoked. */
export declare class ClientResult<T> {
	/** The value of the result that is retrieved from the document after "context.sync()" is invoked. */
	value: T;
}

/** The error object returned by "context.sync()", if a promise is rejected due to an error while processing the request. */
export declare class Error {
	/** Error name: "OfficeExtension.Error".*/
	name: string;
	/** The error message passed through from the host Office application. */
	message: string;
	/** Stack trace, if applicable. */
	stack: string;
	/** Error code string, such as "InvalidArgument". */
	code: string;
	/** Trace messages (if any) that were added via a "context.trace()" invocation before calling "context.sync()". If there was an error, this contains all trace messages that were executed before the error occurred. These messages can help you monitor the program execution sequence and detect the case of the error. */
	traceMessages: Array<string>;
	/** Debug info, if applicable. The ".errorLocation" property can describe the object and method or property that caused the error. */
	debugInfo: {
		/** If applicable, will return the object type and the name of the method or property that caused the error. */
		errorLocation?: string;
	};
}

export declare class ErrorCodes {
	static accessDenied: string;
	static generalException: string;
	static activityLimitReached: string;
}

/** An IPromise object that represents a deferred interaction with the host Office application. */
export declare interface IPromise<R> {
	/**
	 * This method will be called once the previous promise has been resolved.
	 * Both the onFulfilled on onRejected callbacks are optional.
	 * If either or both are omitted, the next onFulfilled/onRejected in the chain will be called called.

	 * @returns A new promise for the value or error that was returned from onFulfilled/onRejected.
	 */
	then<U>(onFulfilled?: (value: R) => IPromise<U>, onRejected?: (error: any) => IPromise<U>): IPromise<U>;

	/**
	 * This method will be called once the previous promise has been resolved.
	 * Both the onFulfilled on onRejected callbacks are optional.
	 * If either or both are omitted, the next onFulfilled/onRejected in the chain will be called called.

	 * @returns A new promise for the value or error that was returned from onFulfilled/onRejected.
	 */
	then<U>(onFulfilled?: (value: R) => IPromise<U>, onRejected?: (error: any) => U): IPromise<U>;

	/**
	 * This method will be called once the previous promise has been resolved.
	 * Both the onFulfilled on onRejected callbacks are optional.
	 * If either or both are omitted, the next onFulfilled/onRejected in the chain will be called called.

	 * @returns A new promise for the value or error that was returned from onFulfilled/onRejected.
	 */
	then<U>(onFulfilled?: (value: R) => IPromise<U>, onRejected?: (error: any) => void): IPromise<U>;

	/**
	 * This method will be called once the previous promise has been resolved.
	 * Both the onFulfilled on onRejected callbacks are optional.
	 * If either or both are omitted, the next onFulfilled/onRejected in the chain will be called called.

	 * @returns A new promise for the value or error that was returned from onFulfilled/onRejected.
	 */
	then<U>(onFulfilled?: (value: R) => U, onRejected?: (error: any) => IPromise<U>): IPromise<U>;

	/**
	 * This method will be called once the previous promise has been resolved.
	 * Both the onFulfilled on onRejected callbacks are optional.
	 * If either or both are omitted, the next onFulfilled/onRejected in the chain will be called called.

	 * @returns A new promise for the value or error that was returned from onFulfilled/onRejected.
	 */
	then<U>(onFulfilled?: (value: R) => U, onRejected?: (error: any) => U): IPromise<U>;

	/**
	 * This method will be called once the previous promise has been resolved.
	 * Both the onFulfilled on onRejected callbacks are optional.
	 * If either or both are omitted, the next onFulfilled/onRejected in the chain will be called called.

	 * @returns A new promise for the value or error that was returned from onFulfilled/onRejected.
	 */
	then<U>(onFulfilled?: (value: R) => U, onRejected?: (error: any) => void): IPromise<U>;


	/**
	 * Catches failures or exceptions from actions within the promise, or from an unhandled exception earlier in the call stack.
	 * @param onRejected function to be called if or when the promise rejects.
	 */
	catch<U>(onRejected?: (error: any) => IPromise<U>): IPromise<U>;

	/**
	 * Catches failures or exceptions from actions within the promise, or from an unhandled exception earlier in the call stack.
	 * @param onRejected function to be called if or when the promise rejects.
	 */
	catch<U>(onRejected?: (error: any) => U): IPromise<U>;

	/**
	 * Catches failures or exceptions from actions within the promise, or from an unhandled exception earlier in the call stack.
	 * @param onRejected function to be called if or when the promise rejects.
	 */
	catch<U>(onRejected?: (error: any) => void): IPromise<U>;
}


/** Collection of tracked objects, contained within a request context. See "context.trackedObjects" for more information. */
export declare class TrackedObjects {
	/** Track a new object for automatic adjustment based on surrounding changes in the document. Only some object types require this. If you are using an object across ".sync" calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you needed to have added the object to the tracked object collection when the object was first created. */
	add(object: ClientObject): void;
	/** Track a new object for automatic adjustment based on surrounding changes in the document. Only some object types require this. If you are using an object across ".sync" calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you needed to have added the object to the tracked object collection when the object was first created. */
	add(objects: ClientObject[]): void;
	/** Release the memory associated with an object that was previously added to this collection. Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You will need to call "context.sync()" before the memory release takes effect. */
	remove(object: ClientObject): void;
	/** Release the memory associated with an object that was previously added to this collection. Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You will need to call "context.sync()" before the memory release takes effect. */
	remove(objects: ClientObject[]): void;
}

export declare class EventHandlers<T> {
	constructor(context: ClientRequestContext, parentObject: ClientObject, name: string, eventInfo: EventInfo<T>);
	add(handler: (args: T) => IPromise<any>): EventHandlerResult<T>;
	remove(handler: (args: T) => IPromise<any>): void;
	removeAll(): void;
}

export declare class EventHandlerResult<T> {
	constructor(context: ClientRequestContext, handlers: EventHandlers<T>, handler: (args: T) => IPromise<any>);
	remove(): void;
}

export declare interface EventInfo<T> {
	registerFunc: (callback: (args: any) => void) => IPromise<any>;
	unregisterFunc: (callback: (args: any) => void) => IPromise<any>;
	eventArgsTransformFunc: (args: any) => IPromise<T>;
}

/**
* Request URL and headers 
*/
export declare interface RequestUrlAndHeaderInfo {
	/** Request URL */
	url: string;
	/** Request headers */
	headers?: {
		[name: string]: string;
	};
}

// Type definitions for Office.js 1.0
// Project: https://github.com/OfficeDev/office-js
// Definitions by: OfficeDev <https://github.com/OfficeDev>,
//                 Ricky Kirkham <https://github.com/Rick-Kirkham>,
//                 Alex Jerabek <https://github.com/AlexJerabek>,
//                 Elizabeth Samuel <https://github.com/ElizabethSamuel-MSFT>,
//                 Alison McKay <https://github.com/alison-mk>,
//                 Sam Ramon <https://github.com/samantharamon>
// Definitions: https://github.com/DefinitelyTyped/DefinitelyTyped
// TypeScript Version: 2.4

/*
office-js
Copyright (c) Microsoft Corporation
*/


////////////////////////////////////////////////////////////////
//////////////////// Begin Office namespace ////////////////////
////////////////////////////////////////////////////////////////

declare namespace Office {
    /** A Promise object. Promises can be chained via ".then", and errors can be caught via ".catch". 
     * When a browser-provided native Promise implementation is available, Office.Promise will switch to use the native Promise instead.
     */
    var Promise: IPromiseConstructor;

    // **Note**: this is a copy of the PromiseConstructor object from
    //     https://github.com/Microsoft/TypeScript/blob/master/lib/lib.es2015.promise.d.ts
    // It is necessary so that even with targeting "ES5" and not specifying any libs,
    //     developers will still get IntelliSense for "Office.Promise" just as they would with a regular Promise.
    // (because even though Promise is part of standard lib.d.ts, PromiseConstructor is not)
    interface IPromiseConstructor {
        /**
         * A reference to the prototype.
         */
        readonly prototype: Promise<any>;

        /**
         * Creates a new Promise.
         * @param executor A callback used to initialize the promise. This callback is passed two arguments:
         * a resolve callback used resolve the promise with a value or the result of another promise,
         * and a reject callback used to reject the promise with a provided reason or error.
         */
        new <T>(executor: (resolve: (value?: T | PromiseLike<T>) => void, reject: (reason?: any) => void) => void): Promise<T>;

        /**
         * Creates a Promise that is resolved with an array of results when all of the provided Promises
         * resolve, or rejected when any Promise is rejected.
         * @param values An array of Promises.
         * @returns A new Promise.
         */
        all<T1, T2, T3, T4, T5, T6, T7, T8, T9, T10>(values: [T1 | PromiseLike<T1>, T2 | PromiseLike<T2>, T3 | PromiseLike<T3>, T4 | PromiseLike<T4>, T5 | PromiseLike<T5>, T6 | PromiseLike<T6>, T7 | PromiseLike<T7>, T8 | PromiseLike<T8>, T9 | PromiseLike<T9>, T10 | PromiseLike<T10>]): Promise<[T1, T2, T3, T4, T5, T6, T7, T8, T9, T10]>;

        /**
         * Creates a Promise that is resolved with an array of results when all of the provided Promises
         * resolve, or rejected when any Promise is rejected.
         * @param values An array of Promises.
         * @returns A new Promise.
         */
        all<T1, T2, T3, T4, T5, T6, T7, T8, T9>(values: [T1 | PromiseLike<T1>, T2 | PromiseLike<T2>, T3 | PromiseLike<T3>, T4 | PromiseLike<T4>, T5 | PromiseLike<T5>, T6 | PromiseLike<T6>, T7 | PromiseLike<T7>, T8 | PromiseLike<T8>, T9 | PromiseLike<T9>]): Promise<[T1, T2, T3, T4, T5, T6, T7, T8, T9]>;

        /**
         * Creates a Promise that is resolved with an array of results when all of the provided Promises
         * resolve, or rejected when any Promise is rejected.
         * @param values An array of Promises.
         * @returns A new Promise.
         */
        all<T1, T2, T3, T4, T5, T6, T7, T8>(values: [T1 | PromiseLike<T1>, T2 | PromiseLike<T2>, T3 | PromiseLike<T3>, T4 | PromiseLike<T4>, T5 | PromiseLike<T5>, T6 | PromiseLike<T6>, T7 | PromiseLike<T7>, T8 | PromiseLike<T8>]): Promise<[T1, T2, T3, T4, T5, T6, T7, T8]>;

        /**
         * Creates a Promise that is resolved with an array of results when all of the provided Promises
         * resolve, or rejected when any Promise is rejected.
         * @param values An array of Promises.
         * @returns A new Promise.
         */
        all<T1, T2, T3, T4, T5, T6, T7>(values: [T1 | PromiseLike<T1>, T2 | PromiseLike<T2>, T3 | PromiseLike<T3>, T4 | PromiseLike<T4>, T5 | PromiseLike<T5>, T6 | PromiseLike<T6>, T7 | PromiseLike<T7>]): Promise<[T1, T2, T3, T4, T5, T6, T7]>;

        /**
         * Creates a Promise that is resolved with an array of results when all of the provided Promises
         * resolve, or rejected when any Promise is rejected.
         * @param values An array of Promises.
         * @returns A new Promise.
         */
        all<T1, T2, T3, T4, T5, T6>(values: [T1 | PromiseLike<T1>, T2 | PromiseLike<T2>, T3 | PromiseLike<T3>, T4 | PromiseLike<T4>, T5 | PromiseLike<T5>, T6 | PromiseLike<T6>]): Promise<[T1, T2, T3, T4, T5, T6]>;

        /**
         * Creates a Promise that is resolved with an array of results when all of the provided Promises
         * resolve, or rejected when any Promise is rejected.
         * @param values An array of Promises.
         * @returns A new Promise.
         */
        all<T1, T2, T3, T4, T5>(values: [T1 | PromiseLike<T1>, T2 | PromiseLike<T2>, T3 | PromiseLike<T3>, T4 | PromiseLike<T4>, T5 | PromiseLike<T5>]): Promise<[T1, T2, T3, T4, T5]>;

        /**
         * Creates a Promise that is resolved with an array of results when all of the provided Promises
         * resolve, or rejected when any Promise is rejected.
         * @param values An array of Promises.
         * @returns A new Promise.
         */
        all<T1, T2, T3, T4>(values: [T1 | PromiseLike<T1>, T2 | PromiseLike<T2>, T3 | PromiseLike<T3>, T4 | PromiseLike<T4>]): Promise<[T1, T2, T3, T4]>;

        /**
         * Creates a Promise that is resolved with an array of results when all of the provided Promises
         * resolve, or rejected when any Promise is rejected.
         * @param values An array of Promises.
         * @returns A new Promise.
         */
        all<T1, T2, T3>(values: [T1 | PromiseLike<T1>, T2 | PromiseLike<T2>, T3 | PromiseLike<T3>]): Promise<[T1, T2, T3]>;

        /**
         * Creates a Promise that is resolved with an array of results when all of the provided Promises
         * resolve, or rejected when any Promise is rejected.
         * @param values An array of Promises.
         * @returns A new Promise.
         */
        all<T1, T2>(values: [T1 | PromiseLike<T1>, T2 | PromiseLike<T2>]): Promise<[T1, T2]>;

        /**
         * Creates a Promise that is resolved with an array of results when all of the provided Promises
         * resolve, or rejected when any Promise is rejected.
         * @param values An array of Promises.
         * @returns A new Promise.
         */
        all<T>(values: (T | PromiseLike<T>)[]): Promise<T[]>;

        /**
         * Creates a Promise that is resolved or rejected when any of the provided Promises are resolved
         * or rejected.
         * @param values An array of Promises.
         * @returns A new Promise.
         */
        race<T1, T2, T3, T4, T5, T6, T7, T8, T9, T10>(values: [T1 | PromiseLike<T1>, T2 | PromiseLike<T2>, T3 | PromiseLike<T3>, T4 | PromiseLike<T4>, T5 | PromiseLike<T5>, T6 | PromiseLike<T6>, T7 | PromiseLike<T7>, T8 | PromiseLike<T8>, T9 | PromiseLike<T9>, T10 | PromiseLike<T10>]): Promise<T1 | T2 | T3 | T4 | T5 | T6 | T7 | T8 | T9 | T10>;

        /**
         * Creates a Promise that is resolved or rejected when any of the provided Promises are resolved
         * or rejected.
         * @param values An array of Promises.
         * @returns A new Promise.
         */
        race<T1, T2, T3, T4, T5, T6, T7, T8, T9>(values: [T1 | PromiseLike<T1>, T2 | PromiseLike<T2>, T3 | PromiseLike<T3>, T4 | PromiseLike<T4>, T5 | PromiseLike<T5>, T6 | PromiseLike<T6>, T7 | PromiseLike<T7>, T8 | PromiseLike<T8>, T9 | PromiseLike<T9>]): Promise<T1 | T2 | T3 | T4 | T5 | T6 | T7 | T8 | T9>;

        /**
         * Creates a Promise that is resolved or rejected when any of the provided Promises are resolved
         * or rejected.
         * @param values An array of Promises.
         * @returns A new Promise.
         */
        race<T1, T2, T3, T4, T5, T6, T7, T8>(values: [T1 | PromiseLike<T1>, T2 | PromiseLike<T2>, T3 | PromiseLike<T3>, T4 | PromiseLike<T4>, T5 | PromiseLike<T5>, T6 | PromiseLike<T6>, T7 | PromiseLike<T7>, T8 | PromiseLike<T8>]): Promise<T1 | T2 | T3 | T4 | T5 | T6 | T7 | T8>;

        /**
         * Creates a Promise that is resolved or rejected when any of the provided Promises are resolved
         * or rejected.
         * @param values An array of Promises.
         * @returns A new Promise.
         */
        race<T1, T2, T3, T4, T5, T6, T7>(values: [T1 | PromiseLike<T1>, T2 | PromiseLike<T2>, T3 | PromiseLike<T3>, T4 | PromiseLike<T4>, T5 | PromiseLike<T5>, T6 | PromiseLike<T6>, T7 | PromiseLike<T7>]): Promise<T1 | T2 | T3 | T4 | T5 | T6 | T7>;

        /**
         * Creates a Promise that is resolved or rejected when any of the provided Promises are resolved
         * or rejected.
         * @param values An array of Promises.
         * @returns A new Promise.
         */
        race<T1, T2, T3, T4, T5, T6>(values: [T1 | PromiseLike<T1>, T2 | PromiseLike<T2>, T3 | PromiseLike<T3>, T4 | PromiseLike<T4>, T5 | PromiseLike<T5>, T6 | PromiseLike<T6>]): Promise<T1 | T2 | T3 | T4 | T5 | T6>;

        /**
         * Creates a Promise that is resolved or rejected when any of the provided Promises are resolved
         * or rejected.
         * @param values An array of Promises.
         * @returns A new Promise.
         */
        race<T1, T2, T3, T4, T5>(values: [T1 | PromiseLike<T1>, T2 | PromiseLike<T2>, T3 | PromiseLike<T3>, T4 | PromiseLike<T4>, T5 | PromiseLike<T5>]): Promise<T1 | T2 | T3 | T4 | T5>;

        /**
         * Creates a Promise that is resolved or rejected when any of the provided Promises are resolved
         * or rejected.
         * @param values An array of Promises.
         * @returns A new Promise.
         */
        race<T1, T2, T3, T4>(values: [T1 | PromiseLike<T1>, T2 | PromiseLike<T2>, T3 | PromiseLike<T3>, T4 | PromiseLike<T4>]): Promise<T1 | T2 | T3 | T4>;

        /**
         * Creates a Promise that is resolved or rejected when any of the provided Promises are resolved
         * or rejected.
         * @param values An array of Promises.
         * @returns A new Promise.
         */
        race<T1, T2, T3>(values: [T1 | PromiseLike<T1>, T2 | PromiseLike<T2>, T3 | PromiseLike<T3>]): Promise<T1 | T2 | T3>;

        /**
         * Creates a Promise that is resolved or rejected when any of the provided Promises are resolved
         * or rejected.
         * @param values An array of Promises.
         * @returns A new Promise.
         */
        race<T1, T2>(values: [T1 | PromiseLike<T1>, T2 | PromiseLike<T2>]): Promise<T1 | T2>;

        /**
         * Creates a Promise that is resolved or rejected when any of the provided Promises are resolved
         * or rejected.
         * @param values An array of Promises.
         * @returns A new Promise.
         */
        race<T>(values: (T | PromiseLike<T>)[]): Promise<T>;

        /**
         * Creates a new rejected promise for the provided reason.
         * @param reason The reason the promise was rejected.
         * @returns A new rejected Promise.
         */
        reject(reason: any): Promise<never>;

        /**
         * Creates a new rejected promise for the provided reason.
         * @param reason The reason the promise was rejected.
         * @returns A new rejected Promise.
         */
        reject<T>(reason: any): Promise<T>;

        /**
         * Creates a new resolved promise for the provided value.
         * @param value A promise.
         * @returns A promise whose internal state matches the provided promise.
         */
        resolve<T>(value: T | PromiseLike<T>): Promise<T>;

        /**
         * Creates a new resolved promise.
         * @returns A resolved promise.
         */
        resolve(): Promise<void>;
    }

    /**
     * Gets the Context object that represents the runtime environment of the add-in and provides access to the top-level objects of the API.
     * 
     * @remarks
     * 
     * **Support details**
     * 
     * A capital Y in the following matrix indicates that this property is supported in the corresponding Office host application. 
     * An empty cell indicates that the Office host application doesn't support this enumeration.
     * 
     * For more information about Office host application and server requirements, see 
     * {@link https://docs.microsoft.com/office/dev/add-ins/concepts/requirements-for-running-office-add-ins | Requirements for running Office Add-ins}.
     * 
     * *Supported hosts, by platform*
     *  <table>
     *   <tr><th>                             </th><th> Office on Windows </th><th> Office in web browser </th><th> Office on iPad </th><th> Outlook on mobile devices </th><th> Office on Mac </th></tr>
     *   <tr><td><strong> Excel      </strong></td><td> Y                 </td><td> Y                     </td><td> Y              </td><td>                           </td><td> Y             </td></tr>
     *   <tr><td><strong> Outlook    </strong></td><td> Y                 </td><td> Y                     </td><td>                </td><td> Y                         </td><td> Y             </td></tr>
     *   <tr><td><strong> PowerPoint </strong></td><td> Y                 </td><td> Y                     </td><td> Y              </td><td>                           </td><td> Y             </td></tr>
     *   <tr><td><strong> Project    </strong></td><td> Y                 </td><td>                       </td><td>                </td><td>                           </td><td> Y             </td></tr>
     *   <tr><td><strong> Word       </strong></td><td> Y                 </td><td> Y                     </td><td> Y              </td><td>                           </td><td> Y             </td></tr>
     *  </table>
     */
    var context: Context;
    /**
     * Occurs when the runtime environment is loaded and the add-in is ready to start interacting with the application and hosted document. 
     * 
     * The reason parameter of the initialize event listener function returns an `InitializationReason` enumeration value that specifies how 
     * initialization occurred. A task pane or content add-in can be initialized in two ways:
     * 
     *  - The user just inserted it from Recently Used Add-ins section of the Add-in drop-down list on the Insert tab of the ribbon in the Office 
     * host application, or from Insert add-in dialog box.
     * 
     *  - The user opened a document that already contains the add-in.
     * 
     * *Note*: The reason parameter of the initialize event listener function only returns an `InitializationReason` enumeration value for task pane 
     * and content add-ins. It does not return a value for Outlook add-ins.
     * 
     * @remarks
     * 
     * **Support details**
     * 
     * A capital Y in the following matrix indicates that this method is supported in the corresponding Office host application. 
     * An empty cell indicates that the Office host application doesn't support this method.
     * 
     * For more information about Office host application and server requirements, see 
     * {@link https://docs.microsoft.com/office/dev/add-ins/concepts/requirements-for-running-office-add-ins | Requirements for running Office Add-ins}.
     * 
    
import * as OfficeExtension from './office.runtime';
export interface ThreeArrowsSet {
    [index: number]: Icon;
    redDownArrow: Icon;
    yellowSideArrow: Icon;
    greenUpArrow: Icon;
}
export interface ThreeArrowsGraySet {
    [index: number]: Icon;
    grayDownArrow: Icon;
    graySideArrow: Icon;
    grayUpArrow: Icon;
}
export interface ThreeFlagsSet {
    [index: number]: Icon;
    redFlag: Icon;
    yellowFlag: Icon;
    greenFlag: Icon;
}
export interface ThreeTrafficLights1Set {
    [index: number]: Icon;
    redCircleWithBorder: Icon;
    yellowCircle: Icon;
    greenCircle: Icon;
}
export interface ThreeTrafficLights2Set {
    [index: number]: Icon;
    redTrafficLight: Icon;
    yellowTrafficLight: Icon;
    greenTrafficLight: Icon;
}
export interface ThreeSignsSet {
    [index: number]: Icon;
    redDiamond: Icon;
    yellowTriangle: Icon;
    greenCircle: Icon;
}
export interface ThreeSymbolsSet {
    [index: number]: Icon;
    redCrossSymbol: Icon;
    yellowExclamationSymbol: Icon;
    greenCheckSymbol: Icon;
}
export interface ThreeSymbols2Set {
    [index: number]: Icon;
    redCross: Icon;
    yellowExclamation: Icon;
    greenCheck: Icon;
}
export interface FourArrowsSet {
    [index: number]: Icon;
    redDownArrow: Icon;
    yellowDownInclineArrow: Icon;
    yellowUpInclineArrow: Icon;
    greenUpArrow: Icon;
}
export interface FourArrowsGraySet {
    [index: number]: Icon;
    grayDownArrow: Icon;
    grayDownInclineArrow: Icon;
    grayUpInclineArrow: Icon;
    grayUpArrow: Icon;
}
export interface FourRedToBlackSet {
    [index: number]: Icon;
    blackCircle: Icon;
    grayCircle: Icon;
    pinkCircle: Icon;
    redCircle: Icon;
}
export interface FourRatingSet {
    [index: number]: Icon;
    oneBar: Icon;
    twoBars: Icon;
    threeBars: Icon;
    fourBars: Icon;
}
export interface FourTrafficLightsSet {
    [index: number]: Icon;
    blackCircleWithBorder: Icon;
    redCircleWithBorder: Icon;
    yellowCircle: Icon;
    greenCircle: Icon;
}
export interface FiveArrowsSet {
    [index: number]: Icon;
    redDownArrow: Icon;
    yellowDownInclineArrow: Icon;
    yellowSideArrow: Icon;
    yellowUpInclineArrow: Icon;
    greenUpArrow: Icon;
}
export interface FiveArrowsGraySet {
    [index: number]: Icon;
    grayDownArrow: Icon;
    grayDownInclineArrow: Icon;
    graySideArrow: Icon;
    grayUpInclineArrow: Icon;
    grayUpArrow: Icon;
}
export interface FiveRatingSet {
    [index: number]: Icon;
    noBars: Icon;
    oneBar: Icon;
    twoBars: Icon;
    threeBars: Icon;
    fourBars: Icon;
}
export interface FiveQuartersSet {
    [index: number]: Icon;
    whiteCircleAllWhiteQuarters: Icon;
    circleWithThreeWhiteQuarters: Icon;
    circleWithTwoWhiteQuarters: Icon;
    circleWithOneWhiteQuarter: Icon;
    blackCircle: Icon;
}
export interface ThreeStarsSet {
    [index: number]: Icon;
    silverStar: Icon;
    halfGoldStar: Icon;
    goldStar: Icon;
}
export interface ThreeTrianglesSet {
    [index: number]: Icon;
    redDownTriangle: Icon;
    yellowDash: Icon;
    greenUpTriangle: Icon;
}
export interface FiveBoxesSet {
    [index: number]: Icon;
    noFilledBoxes: Icon;
    oneFilledBox: Icon;
    twoFilledBoxes: Icon;
    threeFilledBoxes: Icon;
    fourFilledBoxes: Icon;
}
export interface IconCollections {
    threeArrows: ThreeArrowsSet;
    threeArrowsGray: ThreeArrowsGraySet;
    threeFlags: ThreeFlagsSet;
    threeTrafficLights1: ThreeTrafficLights1Set;
    threeTrafficLights2: ThreeTrafficLights2Set;
    threeSigns: ThreeSignsSet;
    threeSymbols: ThreeSymbolsSet;
    threeSymbols2: ThreeSymbols2Set;
    fourArrows: FourArrowsSet;
    fourArrowsGray: FourArrowsGraySet;
    fourRedToBlack: FourRedToBlackSet;
    fourRating: FourRatingSet;
    fourTrafficLights: FourTrafficLightsSet;
    fiveArrows: FiveArrowsSet;
    fiveArrowsGray: FiveArrowsGraySet;
    fiveRating: FiveRatingSet;
    fiveQuarters: FiveQuartersSet;
    threeStars: ThreeStarsSet;
    threeTriangles: ThreeTrianglesSet;
    fiveBoxes: FiveBoxesSet;
}
export declare var icons: IconCollections;
/**
 * Provides connection session for a remote workbook.
 */
export declare class Session {
    private static WorkbookSessionIdHeaderName;
    private static WorkbookSessionIdHeaderNameLower;
    constructor(workbookUrl?: string, requestHeaders?: {
        [name: string]: string;
    }, persisted?: boolean);
    /**
     * Close the session.
     */
    close(): OfficeExtension.IPromise<void>;
}
/**
 * The RequestContext object facilitates requests to the Excel application. Since the Office add-in and the Excel application run in two different processes, the request context is required to get access to the Excel object model from the add-in.
 */
export declare class RequestContext extends OfficeExtension.ClientRequestContext {
    constructor(url?: string | Session);
    workbook: Workbook;
}
/**
 * Executes a batch script that performs actions on the Excel object model, using a new RequestContext. When the promise is resolved, any tracked objects that were automatically allocated during execution will be released.
 * @param batch - A function that takes in a RequestContext and returns a promise (typically, just the result of "context.sync()"). The context parameter facilitates requests to the Excel application. Since the Office add-in and the Excel application run in two different processes, the RequestContext is required to get access to the Excel object model from the add-in.
 */
export declare function run<T>(batch: (context: RequestContext) => OfficeExtension.IPromise<T>): OfficeExtension.IPromise<T>;
/**
 * Executes a batch script that performs actions on the Excel object model, using a new remote RequestContext. When the promise is resolved, any tracked objects that were automatically allocated during execution will be released.
 * @param requestInfo - The URL of the remote workbook and the request headers to be sent.
 * @param batch - A function that takes in a RequestContext and returns a promise (typically, just the result of "context.sync()"). The context parameter facilitates requests to the Excel application. Since the Office add-in and the Excel application run in two different processes, the RequestContext is required to get access to the Excel object model from the add-in.
 */
export declare function run<T>(requestInfo: OfficeExtension.RequestUrlAndHeaderInfo | Session, batch: (context: RequestContext) => OfficeExtension.IPromise<T>): OfficeExtension.IPromise<T>;
/**
 * Executes a batch script that performs actions on the Excel object model, using the RequestContext of a previously-created API object. When the promise is resolved, any tracked objects that were automatically allocated during execution will be released.
 * @param object - A previously-created API object. The batch will use the same RequestContext as the passed-in object, which means that any changes applied to the object will be picked up by "context.sync()".
 * @param batch - A function that takes in a RequestContext and returns a promise (typically, just the result of "context.sync()"). The context parameter facilitates requests to the Excel application. Since the Office add-in and the Excel application run in two different processes, the RequestContext is required to get access to the Excel object model from the add-in.
 */
export declare function run<T>(object: OfficeExtension.ClientObject, batch: (context: RequestContext) => OfficeExtension.IPromise<T>): OfficeExtension.IPromise<T>;
/**
 * Executes a batch script that performs actions on the Excel object model, using the remote RequestContext of a previously-created API object. When the promise is resolved, any tracked objects that were automatically allocated during execution will be released.
 * @param requestInfo - The URL of the remote workbook and the request headers to be sent.
 * @param object - A previously-created API object. The batch will use the same RequestContext as the passed-in object, which means that any changes applied to the object will be picked up by "context.sync()".
 * @param batch - A function that takes in a RequestContext and returns a promise (typically, just the result of "context.sync()"). The context parameter facilitates requests to the Excel application. Since the Office add-in and the Excel application run in two different processes, the RequestContext is required to get access to the Excel object model from the add-in.
 */
export declare function run<T>(requestInfo: OfficeExtension.RequestUrlAndHeaderInfo | Session, object: OfficeExtension.ClientObject, batch: (context: RequestContext) => OfficeExtension.IPromise<T>): OfficeExtension.IPromise<T>;
/**
 * Executes a batch script that performs actions on the Excel object model, using the RequestContext of previously-created API objects.
 * @param objects - An array of previously-created API objects. The array will be validated to make sure that all of the objects share the same context. The batch will use this shared RequestContext, which means that any changes applied to these objects will be picked up by "context.sync()".
 * @param batch - A function that takes in a RequestContext and returns a promise (typically, just the result of "context.sync()"). The context parameter facilitates requests to the Excel application. Since the Office add-in and the Excel application run in two different processes, the RequestContext is required to get access to the Excel object model from the add-in.
 */
export declare function run<T>(objects: OfficeExtension.ClientObject[], batch: (context: RequestContext) => OfficeExtension.IPromise<T>): OfficeExtension.IPromise<T>;
/**
 * Executes a batch script that performs actions on the Excel object model, using the remote RequestContext of previously-created API objects.
 * @param requestInfo - The URL of the remote workbook and the request headers to be sent.
 * @param objects - An array of previously-created API objects. The array will be validated to make sure that all of the objects share the same context. The batch will use this shared RequestContext, which means that any changes applied to these objects will be picked up by "context.sync()".
 * @param batch - A function that takes in a RequestContext and returns a promise (typically, just the result of "context.sync()"). The context parameter facilitates requests to the Excel application. Since the Office add-in and the Excel application run in two different processes, the RequestContext is required to get access to the Excel object model from the add-in.
 */
export declare function run<T>(requestInfo: OfficeExtension.RequestUrlAndHeaderInfo | Session, objects: OfficeExtension.ClientObject[], batch: (context: RequestContext) => OfficeExtension.IPromise<T>): OfficeExtension.IPromise<T>;
export declare var _RedirectV1APIs: boolean;
export declare var _V1APIMap: {
    "GetDataAsync": {
        call: (ctx: any, callArgs: any) => any;
        postprocess: (response: any, callArgs: any) => any;
    };
    "GetSelectedDataAsync": {
        call: (ctx: any, callArgs: any) => any;
        postprocess: (response: any, callArgs: any) => any;
    };
    "GoToByIdAsync": {
        call: (ctx: any, callArgs: any) => any;
    };
    "AddColumnsAsync": {
        call: (ctx: any, callArgs: any) => any;
    };
    "AddFromSelectionAsync": {
        call: (ctx: any, callArgs: any) => any;
        postprocess: (response: any) => any;
    };
    "AddFromNamedItemAsync": {
        call: (ctx: any, callArgs: any) => any;
        postprocess: (response: any) => any;
    };
    "AddFromPromptAsync": {
        call: (ctx: any, callArgs: any) => any;
        postprocess: (response: any) => any;
    };
    "AddRowsAsync": {
        call: (ctx: any, callArgs: any) => any;
    };
    "GetByIdAsync": {
        call: (ctx: any, callArgs: any) => any;
        postprocess: (response: any) => any;
    };
    "ReleaseByIdAsync": {
        call: (ctx: any, callArgs: any) => any;
    };
    "GetAllAsync": {
        call: (ctx: any) => any;
        postprocess: (response: any) => any;
    };
    "DeleteAllDataValuesAsync": {
        call: (ctx: any, callArgs: any) => any;
    };
    "SetSelectedDataAsync": {
        preprocess: (callArgs: any) => any;
        call: (ctx: any, callArgs: any) => any;
    };
    "SetDataAsync": {
        preprocess: (callArgs: any) => any;
        call: (ctx: any, callArgs: any) => any;
    };
    "SetFormatsAsync": {
        preprocess: (callArgs: any) => any;
        call: (ctx: any, callArgs: any) => any;
    };
    "SetTableOptionsAsync": {
        call: (ctx: any, callArgs: any) => any;
    };
    "ClearFormatsAsync": {
        call: (ctx: any, callArgs: any) => any;
    };
};
/**
 *
 * Provides information about the binding that raised the SelectionChanged event.
 *
 * [Api set: ExcelApi 1.2]
 */
export interface BindingSelectionChangedEventArgs {
    /**
     *
     * Gets the Binding object that represents the binding that raised the SelectionChanged event.
     *
     * [Api set: ExcelApi 1.2]
     */
    binding: Binding;
    /**
     *
     * Gets the number of columns selected.
     *
     * [Api set: ExcelApi 1.2]
     */
    columnCount: number;
    /**
     *
     * Gets the number of rows selected.
     *
     * [Api set: ExcelApi 1.2]
     */
    rowCount: number;
    /**
     *
     * Gets the index of the first column of the selection (zero-based).
     *
     * [Api set: ExcelApi 1.2]
     */
    startColumn: number;
    /**
     *
     * Gets the index of the first row of the selection (zero-based).
     *
     * [Api set: ExcelApi 1.2]
     */
    startRow: number;
}
/**
 *
 * Provides information about the binding that raised the DataChanged event.
 *
 * [Api set: ExcelApi 1.2]
 */
export interface BindingDataChangedEventArgs {
    /**
     *
     * Gets the Binding object that represents the binding that raised the DataChanged event.
     *
     * [Api set: ExcelApi 1.2]
     */
    binding: Binding;
}
/**
 *
 * Provides information about the document that raised the SelectionChanged event.
 *
 * [Api set: ExcelApi 1.2]
 */
export interface SelectionChangedEventArgs {
    /**
     *
     * Gets the workbook object that raised the SelectionChanged event.
     *
     * [Api set: ExcelApi 1.2]
     */
    workbook: Workbook;
}
/**
 *
 * Provides information about the setting that raised the SettingsChanged event
 *
 * [Api set: ExcelApi 1.4]
 */
export interface SettingsChangedEventArgs {
    /**
     *
     * Gets the Setting object that represents the binding that raised the SettingsChanged event
     *
     * [Api set: ExcelApi 1.4]
     */
    settings: SettingCollection;
}
/**
 *
 * Represents the Excel application that manages the workbook.
 *
 * [Api set: ExcelApi 1.1]
 */
export declare class Application extends OfficeExtension.ClientObject {
    /**
     *
     * Returns the calculation mode used in the workbook. See CalculationMode for details. Read-only.
     *
     * [Api set: ExcelApi 1.1]
     */
    calculationMode: string;
    /**
     *
     * Recalculate all currently opened workbooks in
     *
     * [Api set: ExcelApi 1.1]
     *
     * @param calculationType Specifies the calculation type to use. See CalculationType for details.
     */
    calculate(calculationType: string): void;
    /**
     * Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
     */
    load(option?: string | string[] | OfficeExtension.LoadOption): Application;
    toJSON(): {
        "calculationMode": string;
    };
}
/**
 *
 * Workbook is the top level object which contains related workbook objects such as worksheets, tables, ranges, etc.
 *
 * [Api set: ExcelApi 1.1]
 */
export declare class Workbook extends OfficeExtension.ClientObject {
    /**
     *
     * Represents Excel application instance that contains this workbook. Read-only.
     *
     * [Api set: ExcelApi 1.1]
     */
    application: Application;
    /**
     *
     * Represents a collection of bindings that are part of the workbook. Read-only.
     *
     * [Api set: ExcelApi 1.1]
     */
    bindings: BindingCollection;
    /**
     *
     * Represents the collection of custom XML parts contained by this workbook. Read-only.
     *
     * [Api set: ExcelApi 1.4]
     */
    customXmlParts: CustomXmlPartCollection;
    /**
     *
     * Represents Excel application instance that contains this workbook. Read-only.
     *
     * [Api set: ExcelApi 1.2]
     */
    functions: Functions;
    /**
     *
     * Represents a collection of workbook scoped named items (named ranges and constants). Read-only.
     *
     * [Api set: ExcelApi 1.1]
     */
    names: NamedItemCollection;
    /**
     *
     * Represents a collection of PivotTables associated with the workbook. Read-only.
     *
     * [Api set: ExcelApi 1.3]
     */
    pivotTables: PivotTableCollection;
    /**
     *
     * Represents a collection of Settings associated with the workbook. Read-only.
     *
     * [Api set: ExcelApi 1.4]
     */
    settings: SettingCollection;
    /**
     *
     * Represents a collection of tables associated with the workbook. Read-only.
     *
     * [Api set: ExcelApi 1.1]
     */
    tables: TableCollection;
    /**
     *
     * Represents a collection of worksheets associated with the workbook. Read-only.
     *
     * [Api set: ExcelApi 1.1]
     */
    worksheets: WorksheetCollection;
    /**
     *
     * Gets the currently selected range from the workbook.
     *
     * [Api set: ExcelApi 1.1]
     */
    getSelectedRange(): Range;
    /**
     * Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
     */
    load(option?: string | string[] | OfficeExtension.LoadOption): Workbook;
    /**
     *
     * Occurs when the selection in the document is changed.
     *
     * [Api set: ExcelApi 1.2]
     */
    onSelectionChanged: OfficeExtension.EventHandlers<SelectionChangedEventArgs>;
    toJSON(): {};
}
/**
 *
 * An Excel worksheet is a grid of cells. It can contain data, tables, charts, etc.
 *
 * [Api set: ExcelApi 1.1]
 */
export declare class Worksheet extends OfficeExtension.ClientObject {
    /**
     *
     * Returns collection of charts that are part of the worksheet. Read-only.
     *
     * [Api set: ExcelApi 1.1]
     */
    charts: ChartCollection;
    /**
     *
     * Collection of names scoped to the current worksheet. Read-only.
     *
     * [Api set: ExcelApi 1.4]
     */
    names: NamedItemCollection;
    /**
     *
     * Collection of PivotTables that are part of the worksheet. Read-only.
     *
     * [Api set: ExcelApi 1.3]
     */
    pivotTables: PivotTableCollection;
    /**
     *
     * Returns sheet protection object for a worksheet.
     *
     * [Api set: ExcelApi 1.2]
     */
    protection: WorksheetProtection;
    /**
     *
     * Collection of tables that are part of the worksheet. Read-only.
     *
     * [Api set: ExcelApi 1.1]
     */
    tables: TableCollection;
    /**
     *
     * Returns a value that uniquely identifies the worksheet in a given workbook. The value of the identifier remains the same even when the worksheet is renamed or moved. Read-only.
     *
     * [Api set: ExcelApi 1.1]
     */
    id: string;
    /**
     *
     * The display name of the worksheet.
     *
     * [Api set: ExcelApi 1.1]
     */
    name: string;
    /**
     *
     * The zero-based position of the worksheet within the workbook.
     *
     * [Api set: ExcelApi 1.1]
     */
    position: number;
    /**
     *
     * The Visibility of the worksheet.
     *
     * [Api set: ExcelApi 1.1 for reading visibility; 1.2 for setting it.]
     */
    visibility: string;
    /**
     *
     * Activate the worksheet in the Excel UI.
     *
     * [Api set: ExcelApi 1.1]
     */
    activate(): void;
    /**
     *
     * Deletes the worksheet from the workbook.
     *
     * [Api set: ExcelApi 1.1]
     */
    delete(): void;
    /**
     *
     * Gets the range object containing the single cell based on row and column numbers. The cell can be outside the bounds of its parent range, so long as it's stays within the worksheet grid.
     *
     * [Api set: ExcelApi 1.1]
     *
     * @param row The row number of the cell to be retrieved. Zero-indexed.
     * @param column the column number of the cell to be retrieved. Zero-indexed.
     */
    getCell(row: number, column: number): Range;
    /**
     *
     * Gets the range object specified by the address or name.
     *
     * [Api set: ExcelApi 1.1]
     *
     * @param address The address or the name of the range. If not specified, the entire worksheet range is returned.
     */
    getRange(address?: string): Range;
    /**
     *
     * The used range is the smallest range that encompasses any cells that have a value or formatting assigned to them. If the entire worksheet is blank, this function will return the top left cell (i.e.,: it will *not* throw an error).
     *
     * [Api set: ExcelApi 1.1]
     *
     * @param valuesOnly Considers only cells with values as used cells (ignores formatting). [Api set: ExcelApi 1.2]
     */
    getUsedRange(valuesOnly?: boolean): Range;
    /**
     * Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
     */
    load(option?: string | string[] | OfficeExtension.LoadOption): Worksheet;
    toJSON(): {
        "id": string;
        "name": string;
        "position": number;
        "protection": WorksheetProtection;
        "visibility": string;
    };
}
/**
 *
 * Represents a collection of worksheet objects that are part of the workbook.
 *
 * [Api set: ExcelApi 1.1]
 */
export declare class WorksheetCollection extends OfficeExtension.ClientObject {
    /** Gets the loaded child items in this collection. */
    items: Array<Worksheet>;
    /**
     *
     * Adds a new worksheet to the workbook. The worksheet will be added at the end of existing worksheets. If you wish to activate the newly added worksheet, call ".activate() on it.
     *
     * [Api set: ExcelApi 1.1]
     *
     * @param name The name of the worksheet to be added. If specified, name should be unqiue. If not specified, Excel determines the name of the new worksheet.
     */
    add(name?: string): Worksheet;
    /**
     *
     * Gets the currently active worksheet in the workbook.
     *
     * [Api set: ExcelApi 1.1]
     */
    getActiveWorksheet(): Worksheet;
    /**
     *
     * Gets a worksheet object using its Name or ID.
     *
     * [Api set: ExcelApi 1.1]
     *
     * @param key The Name or ID of the worksheet.
     */
    getItem(key: string): Worksheet;
    /**
     *
     * Gets a worksheet object using its Name or ID. If the worksheet does not exist, will return a null object.
     *
     * [Api set: ExcelApi 1.4]
     *
     * @param key The Name or ID of the worksheet.
     */
    getItemOrNullObject(key: string): Worksheet;
    /**
     * Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
     */
    load(option?: string | string[] | OfficeExtension.LoadOption): WorksheetCollection;
    toJSON(): {};
}
/**
 *
 * Represents the protection of a sheet object.
 *
 * [Api set: ExcelApi 1.2]
 */
export declare class WorksheetProtection extends OfficeExtension.ClientObject {
    /**
     *
     * Sheet protection options. Read-Only.
     *
     * [Api set: ExcelApi 1.2]
     */
    options: WorksheetProtectionOptions;
    /**
     *
     * Indicates if the worksheet is protected. Read-Only.
     *
     * [Api set: ExcelApi 1.2]
     */
    protected: boolean;
    /**
     *
     * Protects a worksheet. Fails if the worksheet has been protected.
     *
     * [Api set: ExcelApi 1.2]
     *
     * @param options sheet protection options.
     */
    protect(options?: WorksheetProtectionOptions): void;
    /**
     *
     * Unprotects a worksheet.
     *
     * [Api set: ExcelApi 1.2]
     */
    unprotect(): void;
    /**
     * Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
     */
    load(option?: string | string[] | OfficeExtension.LoadOption): WorksheetProtection;
    toJSON(): {
        "options": WorksheetProtectionOptions;
        "protected": boolean;
    };
}
/**
 *
 * Represents the options in sheet protection.
 *
 * [Api set: ExcelApi 1.2]
 */
export interface WorksheetProtectionOptions {
    /**
     *
     * Represents the worksheet protection option of allowing using auto filter feature.
     *
     * [Api set: ExcelApi 1.2]
     */
    allowAutoFilter?: boolean;
    /**
     *
     * Represents the worksheet protection option of allowing deleting columns.
     *
     * [Api set: ExcelApi 1.2]
     */
    allowDeleteColumns?: boolean;
    /**
     *
     * Represents the worksheet protection option of allowing deleting rows.
     *
     * [Api set: ExcelApi 1.2]
     */
    allowDeleteRows?: boolean;
    /**
     *
     * Represents the worksheet protection option of allowing formatting cells.
     *
     * [Api set: ExcelApi 1.2]
     */
    allowFormatCells?: boolean;
    /**
     *
     * Represents the worksheet protection option of allowing formatting columns.
     *
     * [Api set: ExcelApi 1.2]
     */
    allowFormatColumns?: boolean;
    /**
     *
     * Represents the worksheet protection option of allowing formatting rows.
     *
     * [Api set: ExcelApi 1.2]
     */
    allowFormatRows?: boolean;
    /**
     *
     * Represents the worksheet protection option of allowing inserting columns.
     *
     * [Api set: ExcelApi 1.2]
     */
    allowInsertColumns?: boolean;
    /**
     *
     * Represents the worksheet protection option of allowing inserting hyperlinks.
     *
     * [Api set: ExcelApi 1.2]
     */
    allowInsertHyperlinks?: boolean;
    /**
     *
     * Represents the worksheet protection option of allowing inserting rows.
     *
     * [Api set: ExcelApi 1.2]
     */
    allowInsertRows?: boolean;
    /**
     *
     * Represents the worksheet protection option of allowing using PivotTable feature.
     *
     * [Api set: ExcelApi 1.2]
     */
    allowPivotTables?: boolean;
    /**
     *
     * Represents the worksheet protection option of allowing using sort feature.
     *
     * [Api set: ExcelApi 1.2]
     */
    allowSort?: boolean;
}
/**
 *
 * Range represents a set of one or more contiguous cells such as a cell, a row, a column, block of cells, etc.
 *
 * [Api set: ExcelApi 1.1]
 */
export declare class Range extends OfficeExtension.ClientObject {
    /**
     *
     * Collection of ConditionalFormats that intersect the range. Read-only.
     *
     * [Api set: ExcelApi 1.5]
     */
    conditionalFormats: ConditionalFormatCollection;
    /**
     *
     * Returns a format object, encapsulating the range's font, fill, borders, alignment, and other properties. Read-only.
     *
     * [Api set: ExcelApi 1.1]
     */
    format: RangeFormat;
    /**
     *
     * Represents the range sort of the current range.
     *
     * [Api set: ExcelApi 1.2]
     */
    sort: RangeSort;
    /**
     *
     * The worksheet containing the current range. Read-only.
     *
     * [Api set: ExcelApi 1.1]
     */
    worksheet: Worksheet;
    /**
     *
     * Represents the range reference in A1-style. Address value will contain the Sheet reference (e.g. Sheet1!A1:B4). Read-only.
     *
     * [Api set: ExcelApi 1.1]
     */
    address: string;
    /**
     *
     * Represents range reference for the specified range in the language of the user. Read-only.
     *
     * [Api set: ExcelApi 1.1]
     */
    addressLocal: string;
    /**
     *
     * Number of cells in the range. This API will return -1 if the cell count exceeds 2^31-1 (2,147,483,647). Read-only.
     *
     * [Api set: ExcelApi 1.1]
     */
    cellCount: number;
    /**
     *
     * Represents the total number of columns in the range. Read-only.
     *
     * [Api set: ExcelApi 1.1]
     */
    columnCount: number;
    /**
     *
     * Represents if all columns of the current range are hidden.
     *
     * [Api set: ExcelApi 1.2]
     */
    columnHidden: boolean;
    /**
     *
     * Represents the column number of the first cell in the range. Zero-indexed. Read-only.
     *
     * [Api set: ExcelApi 1.1]
     */
    columnIndex: number;
    /**
     *
     * Represents the formula in A1-style notation.
     *
     * [Api set: ExcelApi 1.1]
     */
    formulas: Array<Array<any>>;
    /**
     *
     * Represents the formula in A1-style notation, in the user's language and number-formatting locale.  For example, the English "=SUM(A1, 1.5)" formula would become "=SUMME(A1; 1,5)" in German.
     *
     * [Api set: ExcelApi 1.1]
     */
    formulasLocal: Array<Array<any>>;
    /**
     *
     * Represents the formula in R1C1-style notation.
     *
     * [Api set: ExcelApi 1.2]
     */
    formulasR1C1: Array<Array<any>>;
    /**
     *
     * Represents if all cells of the current range are hidden.
     *
     * [Api set: ExcelApi 1.2]
     */
    hidden: boolean;
    /**
     *
     * Represents Excel's number format code for the given cell.
     *
     * [Api set: ExcelApi 1.1]
     */
    numberFormat: Array<Array<any>>;
    /**
     *
     * Returns the total number of rows in the range. Read-only.
     *
     * [Api set: ExcelApi 1.1]
     */
    rowCount: number;
    /**
     *
     * Represents if all rows of the current range are hidden.
     *
     * [Api set: ExcelApi 1.2]
     */
    rowHidden: boolean;
    /**
     *
     * Returns the row number of the first cell in the range. Zero-indexed. Read-only.
     *
     * [Api set: ExcelApi 1.1]
     */
    rowIndex: number;
    /**
     *
     * Text values of the specified range. The Text value will not depend on the cell width. The # sign substitution that happens in Excel UI will not affect the text value returned by the API. Read-only.
     *
     * [Api set: ExcelApi 1.1]
     */
    text: Array<Array<any>>;
    /**
     *
     * Represents the type of data of each cell. Read-only.
     *
     * [Api set: ExcelApi 1.1]
     */
    valueTypes: Array<Array<string>>;
    /**
     *
     * Represents the raw values of the specified range. The data returned could be of type string, number, or a boolean. Cell that contain an error will return the error string.
     *
     * [Api set: ExcelApi 1.1]
     */
    values: Array<Array<any>>;
    /**
     *
     * Clear range values, format, fill, border, etc.
     *
     * [Api set: ExcelApi 1.1]
     *
     * @param applyTo Determines the type of clear action. See ClearApplyTo for details.
     */
    clear(applyTo?: string): void;
    /**
     *
     * Deletes the cells associated with the range.
     *
     * [Api set: ExcelApi 1.1]
     *
     * @param shift Specifies which way to shift the cells. See DeleteShiftDirection for details.
     */
    delete(shift: string): void;
    /**
     *
     * Gets the smallest range object that encompasses the given ranges. For example, the GetBoundingRect of "B2:C5" and "D10:E15" is "B2:E16".
     *
     * [Api set: ExcelApi 1.1]
     *
     * @param anotherRange The range object or address or range name.
     */
    getBoundingRect(anotherRange: Range | string): Range;
    /**
     *
     * Gets the range object containing the single cell based on row and column numbers. The cell can be outside the bounds of its parent range, so long as it's stays within the worksheet grid. The returned cell is located relative to the top left cell of the range.
     *
     * [Api set: ExcelApi 1.1]
     *
     * @param row Row number of the cell to be retrieved. Zero-indexed.
     * @param column Column number of the cell to be retrieved. Zero-indexed.
     */
    getCell(row: number, column: number): Range;
    /**
     *
     * Gets a column contained in the range.
     *
     * [Api set: ExcelApi 1.1]
     *
     * @param column Column number of the range to be retrieved. Zero-indexed.
     */
    getColumn(column: number): Range;
    /**
     *
     * Gets a certain number of columns to the right of the current Range object.
     *
     * [Api set: ExcelApi 1.2]
     *
     * @param count The number of columns to include in the resulting range. In general, use a positive number to create a range outside the current range. You can also use a negative number to create a range within the current range. The default value is 1.
     */
    getColumnsAfter(count?: number): Range;
    /**
     *
     * Gets a certain number of columns to the left of the current Range object.
     *
     * [Api set: ExcelApi 1.2]
     *
     * @param count The number of columns to include in the resulting range. In general, use a positive number to create a range outside the current range. You can also use a negative number to create a range within the current range. The default value is 1.
     */
    getColumnsBefore(count?: number): Range;
    /**
     *
     * Gets an object that represents the entire column of the range.
     *
     * [Api set: ExcelApi 1.1]
     */
    getEntireColumn(): Range;
    /**
     *
     * Gets an object that represents the entire row of the range.
     *
     * [Api set: ExcelApi 1.1]
     */
    getEntireRow(): Range;
    /**
     *
     * Gets the range object that represents the rectangular intersection of the given ranges.
     *
     * [Api set: ExcelApi 1.1]
     *
     * @param anotherRange The range object or range address that will be used to determine the intersection of ranges.
     */
    getIntersection(anotherRange: Range | string): Range;
    /**
     *
     * Gets the range object that represents the rectangular intersection of the given ranges. If no intersection is found, will return a null object.
     *
     * [Api set: ExcelApi 1.4]
     *
     * @param anotherRange The range object or range address that will be used to determine the intersection of ranges.
     */
    getIntersectionOrNullObject(anotherRange: Range | string): Range;
    /**
     *
     * Gets the last cell within the range. For example, the last cell of "B2:D5" is "D5".
     *
     * [Api set: ExcelApi 1.1]
     */
    getLastCell(): Range;
    /**
     *
     * Gets the last column within the range. For example, the last column of "B2:D5" is "D2:D5".
     *
     * [Api set: ExcelApi 1.1]
     */
    getLastColumn(): Range;
    /**
     *
     * Gets the last row within the range. For example, the last row of "B2:D5" is "B5:D5".
     *
     * [Api set: ExcelApi 1.1]
     */
    getLastRow(): Range;
    /**
     *
     * Gets an object which represents a range that's offset from the specified range. The dimension of the returned range will match this range. If the resulting range is forced outside the bounds of the worksheet grid, an error will be thrown.
     *
     * [Api set: ExcelApi 1.1]
     *
     * @param rowOffset The number of rows (positive, negative, or 0) by which the range is to be offset. Positive values are offset downward, and negative values are offset upward.
     * @param columnOffset The number of columns (positive, negative, or 0) by which the range is to be offset. Positive values are offset to the right, and negative values are offset to the left.
     */
    getOffsetRange(rowOffset: number, columnOffset: number): Range;
    /**
     *
     * Gets a Range object similar to the current Range object, but with its bottom-right corner expanded (or contracted) by some number of rows and columns.
     *
     * [Api set: ExcelApi 1.2]
     *
     * @param deltaRows The number of rows by which to expand the bottom-right corner, relative to the current range. Use a positive number to expand the range, or a negative number to decrease it.
     * @param deltaColumns The number of columnsby which to expand the bottom-right corner, relative to the current range. Use a positive number to expand the range, or a negative number to decrease it.
     */
    getResizedRange(deltaRows: number, deltaColumns: number): Range;
    /**
     *
     * Gets a row contained in the range.
     *
     * [Api set: ExcelApi 1.1]
     *
     * @param row Row number of the range to be retrieved. Zero-indexed.
     */
    getRow(row: number): Range;
    /**
     *
     * Gets a certain number of rows above the current Range object.
     *
     * [Api set: ExcelApi 1.2]
     *
     * @param count The number of rows to include in the resulting range. In general, use a positive number to create a range outside the current range. You can also use a negative number to create a range within the current range. The default value is 1.
     */
    getRowsAbove(count?: number): Range;
    /**
     *
     * Gets a certain number of rows below the current Range object.
     *
     * [Api set: ExcelApi 1.2]
     *
     * @param count The number of rows to include in the resulting range. In general, use a positive number to create a range outside the current range. You can also use a negative number to create a range within the current range. The default value is 1.
     */
    getRowsBelow(count?: number): Range;
    /**
     *
     * Returns the used range of the given range object. If there are no used cells within the range, this function will throw an ItemNotFound error.
     *
     * [Api set: ExcelApi 1.1]
     *
     * @param valuesOnly Considers only cells with values as used cells. [Api set: ExcelApi 1.2]
     */
    getUsedRange(valuesOnly?: boolean): Range;
    /**
     *
     * Represents the visible rows of the current range.
     *
     * [Api set: ExcelApi 1.3]
     */
    getVisibleView(): RangeView;
    /**
     *
     * Inserts a cell or a range of cells into the worksheet in place of this range, and shifts the other cells to make space. Returns a new Range object at the now blank space.
     *
     * [Api set: ExcelApi 1.1]
     *
     * @param shift Specifies which way to shift the cells. See InsertShiftDirection for details.
     */
    insert(shift: string): Range;
    /**
     *
     * Merge the range cells into one region in the worksheet.
     *
     * [Api set: ExcelApi 1.2]
     *
     * @param across Set true to merge cells in each row of the specified range as separate merged cells. The default value is false.
     */
    merge(across?: boolean): void;
    /**
     *
     * Selects the specified range in the Excel UI.
     *
     * [Api set: ExcelApi 1.1]
     */
    select(): void;
    /**
     *
     * Unmerge the range cells into separate cells.
     *
     * [Api set: ExcelApi 1.2]
     */
    unmerge(): void;
    /**
     * Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
     */
    load(option?: string | string[] | OfficeExtension.LoadOption): Range;
    /**
     * Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for context.trackedObjects.add(thisObject). If you are using this object across ".sync" calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you needed to have added the object to the tracked object collection when the object was first created.
     */
    track(): Range;
    /**
     * Release the memory associated with this object, if it has previously been tracked. This call is shorthand for context.trackedObjects.remove(thisObject). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You will need to call "context.sync()" before the memory release takes effect.
     */
    untrack(): Range;
    toJSON(): {
        "address": string;
        "addressLocal": string;
        "cellCount": number;
        "columnCount": number;
        "columnHidden": boolean;
        "columnIndex": number;
        "format": RangeFormat;
        "formulas": any[][];
        "formulasLocal": any[][];
        "formulasR1C1": any[][];
        "hidden": boolean;
        "numberFormat": any[][];
        "rowCount": number;
        "rowHidden": boolean;
        "rowIndex": number;
        "text": any[][];
        "values": any[][];
        "valueTypes": string[][];
    };
}
/**
 *
 * Represents a string reference of the form SheetName!A1:B5, or a global or local named range
 *
 * [Api set: ExcelApi 1.2]
 */
export interface RangeReference {
    address: string;
}
/**
 *
 * RangeView represents a set of visible cells of the parent range.
 *
 * [Api set: ExcelApi 1.3]
 */
export declare class RangeView extends OfficeExtension.ClientObject {
    /**
     *
     * Represents a collection of range views associated with the range. Read-only.
     *
     * [Api set: ExcelApi 1.3]
     */
    rows: RangeViewCollection;
    /**
     *
     * Represents the cell addresses of the RangeView.
     *
     * [Api set: ExcelApi 1.3]
     */
    cellAddresses: Array<Array<any>>;
    /**
     *
     * Returns the number of visible columns. Read-only.
     *
     * [Api set: ExcelApi 1.3]
     */
    columnCount: number;
    /**
     *
     * Represents the formula in A1-style notation.
     *
     * [Api set: ExcelApi 1.3]
     */
    formulas: Array<Array<any>>;
    /**
     *
     * Represents the formula in A1-style notation, in the user's language and number-formatting locale.  For example, the English "=SUM(A1, 1.5)" formula would become "=SUMME(A1; 1,5)" in German.
     *
     * [Api set: ExcelApi 1.3]
     */
    formulasLocal: Array<Array<any>>;
    /**
     *
     * Represents the formula in R1C1-style notation.
     *
     * [Api set: ExcelApi 1.3]
     */
    formulasR1C1: Array<Array<any>>;
    /**
     *
     * Returns a value that represents the index of the RangeView. Read-only.
     *
     * [Api set: ExcelApi 1.3]
     */
    index: number;
    /**
     *
     * Represents Excel's number format code for the given cell.
     *
     * [Api set: ExcelApi 1.3]
     */
    numberFormat: Array<Array<any>>;
    /**
     *
     * Returns the number of visible rows. Read-only.
     *
     * [Api set: ExcelApi 1.3]
     */
    rowCount: number;
    /**
     *
     * Text values of the specified range. The Text value will not depend on the cell width. The # sign substitution that happens in Excel UI will not affect the text value returned by the API. Read-only.
     *
     * [Api set: ExcelApi 1.3]
     */
    text: Array<Array<any>>;
    /**
     *
     * Represents the type of data of each cell. Read-only.
     *
     * [Api set: ExcelApi 1.3]
     */
    valueTypes: Array<Array<string>>;
    /**
     *
     * Represents the raw values of the specified range view. The data returned could be of type string, number, or a boolean. Cell that contain an error will return the error string.
     *
     * [Api set: ExcelApi 1.3]
     */
    values: Array<Array<any>>;
    /**
     *
     * Gets the parent range associated with the current RangeView.
     *
     * [Api set: ExcelApi 1.3]
     */
    getRange(): Range;
    /**
     * Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
     */
    load(option?: string | string[] | OfficeExtension.LoadOption): RangeView;
    toJSON(): {
        "cellAddresses": any[][];
        "columnCount": number;
        "formulas": any[][];
        "formulasLocal": any[][];
        "formulasR1C1": any[][];
        "index": number;
        "numberFormat": any[][];
        "rowCount": number;
        "text": any[][];
        "values": any[][];
        "valueTypes": string[][];
    };
}
/**
 *
 * Represents a collection of worksheet objects that are part of the workbook.
 *
 * [Api set: ExcelApi 1.3]
 */
export declare class RangeViewCollection extends OfficeExtension.ClientObject {
    /** Gets the loaded child items in this collection. */
    items: Array<RangeView>;
    /**
     *
     * Gets a RangeView Row via it's index. Zero-Indexed.
     *
     * [Api set: ExcelApi 1.3]
     *
     * @param index Index of the visible row.
     */
    getItemAt(index: number): RangeView;
    /**
     * Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
     */
    load(option?: string | string[] | OfficeExtension.LoadOption): RangeViewCollection;
    toJSON(): {};
}
/**
 *
 * Represents a collection of worksheet objects that are part of the workbook.
 *
 * [Api set: ExcelApi 1.4]
 */
export declare class SettingCollection extends OfficeExtension.ClientObject {
    /** Gets the loaded child items in this collection. */
    items: Array<Setting>;
    /**
     *
     * Sets or adds the specified setting to the workbook.
     *
     * [Api set: ExcelApi 1.4]
     *
     * @param key The Key of the new setting.
     * @param value The Value for the new setting.
     */
    add(key: string, value: string | number | boolean | Array<any> | any): Setting;
    /**
     *
     * Gets a Setting entry via the key.
     *
     * [Api set: ExcelApi 1.4]
     *
     * @param key Key of the setting.
     */
    getItem(key: string): Setting;
    /**
     *
     * Gets a Setting entry via the key. If the Setting does not exist, will return a null object.
     *
     * [Api set: ExcelApi 1.4]
     *
     * @param key The key of the setting.
     */
    getItemOrNullObject(key: string): Setting;
    /**
     * Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
     */
    load(option?: string | string[] | OfficeExtension.LoadOption): SettingCollection;
    /**
     *
     * Occurs when the Settings in the document are changed.
     *
     * [Api set: ExcelApi 1.4]
     */
    onSettingsChanged: OfficeExtension.EventHandlers<SettingsChangedEventArgs>;
    toJSON(): {};
}
/**
 *
 * Setting represents a key-value pair of a setting persisted to the document.
 *
 * [Api set: ExcelApi 1.4]
 */
export declare class Setting extends OfficeExtension.ClientObject {
    /**
     *
     * Returns the key that represents the id of the Setting. Read-only.
     *
     * [Api set: ExcelApi 1.4]
     */
    key: string;
    /**
     *
     * Represents the value stored for this setting.
     *
     * [Api set: ExcelApi 1.4]
     */
    value: any;
    /**
     *
     * Deletes the setting.
     *
     * [Api set: ExcelApi 1.4]
     */
    delete(): void;
    /**
     * Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
     */
    load(option?: string | string[] | OfficeExtension.LoadOption): Setting;
    toJSON(): {
        "key": string;
        "value": any;
    };
}
/**
 *
 * A collection of all the nameditem objects that are part of the workbook or worksheet, depending on how it was reached.
 *
 * [Api set: ExcelApi 1.1]
 */
export declare class NamedItemCollection extends OfficeExtension.ClientObject {
    /** Gets the loaded child items in this collection. */
    items: Array<NamedItem>;
    /**
     *
     * Adds a new name to the collection of the given scope.
     *
     * [Api set: ExcelApi 1.4]
     *
     * @param name The name of the named item.
     * @param reference The formula or the range that the name will refer to.
     * @param comment The comment associated with the named item
     * @returns
     */
    add(name: string, reference: Range | string, comment?: string): NamedItem;
    /**
     *
     * Adds a new name to the collection of the given scope using the user's locale for the formula.
     *
     * [Api set: ExcelApi 1.4]
     *
     * @param name The "name" of the named item.
     * @param formula The formula in the user's locale that the name will refer to.
     * @param comment The comment associated with the named item
     * @returns
     */
    addFormulaLocal(name: string, formula: string, comment?: string): NamedItem;
    /**
     *
     * Gets a nameditem object using its name
     *
     * [Api set: ExcelApi 1.1]
     *
     * @param name nameditem name.
     */
    getItem(name: string): NamedItem;
    /**
     *
     * Gets a nameditem object using its name. If the nameditem object does not exist, will return a null object.
     *
     * [Api set: ExcelApi 1.4]
     *
     * @param name nameditem name.
     */
    getItemOrNullObject(name: string): NamedItem;
    /**
     * Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
     */
    load(option?: string | string[] | OfficeExtension.LoadOption): NamedItemCollection;
    toJSON(): {};
}
/**
 *
 * Represents a defined name for a range of cells or value. Names can be primitive named objects (as seen in the type below), range object, reference to a range. This object can be used to obtain range object associated with names.
 *
 * [Api set: ExcelApi 1.1]
 */
export declare class NamedItem extends OfficeExtension.ClientObject {
    /**
     *
     * Returns the worksheet on which the named item is scoped to. Throws an error if the items is scoped to the workbook instead.
     *
     * [Api set: ExcelApi 1.4]
     */
    worksheet: Worksheet;
    /**
     *
     * Returns the worksheet on which the named item is scoped to. Returns a null object if the item is scoped to the workbook instead.
     *
     * [Api set: ExcelApi 1.4]
     */
    worksheetOrNullObject: Worksheet;
    /**
     *
     * Represents the comment associated with this name.
     *
     * [Api set: ExcelApi 1.4]
     */
    comment: string;
    /**
     *
     * The name of the object. Read-only.
     *
     * [Api set: ExcelApi 1.1]
     */
    name: string;
    /**
     *
     * Indicates whether the name is scoped to the workbook or to a specific worksheet. Read-only.
     *
     * [Api set: ExcelApi 1.4]
     */
    scope: string;
    /**
     *
     * Indicates the type of the value returned by the name's formula. See NamedItemType for details. Read-only.
     *
     * [Api set: ExcelApi 1.1]
     */
    type: string;
    /**
     *
     * Represents the value computed by the name's formula. For a named range, will return the range address. Read-only.
     *
     * [Api set: ExcelApi 1.1]
     */
    value: any;
    /**
     *
     * Specifies whether the object is visible or not.
     *
     * [Api set: ExcelApi 1.1]
     */
    visible: boolean;
    /**
     *
     * Deletes the given name.
     *
     * [Api set: ExcelApi 1.4]
     */
    delete(): void;
    /**
     *
     * Returns the range object that is associated with the name. Throws an error if the named item's type is not a range.
     *
     * [Api set: ExcelApi 1.4]
     */
    getRange(): Range;
    /**
     *
     * Returns the range object that is associated with the name. Returns a null object if the named item's type is not a range.
     *
     * [Api set: ExcelApi 1.4]
     */
    getRangeOrNullObject(): Range;
    /**
     * Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
     */
    load(option?: string | string[] | OfficeExtension.LoadOption): NamedItem;
    toJSON(): {
        "comment": string;
        "name": string;
        "scope": string;
        "type": string;
        "value": any;
        "visible": boolean;
    };
}
/**
 *
 * Represents an Office.js binding that is defined in the workbook.
 *
 * [Api set: ExcelApi 1.1]
 */
export declare class Binding extends OfficeExtension.ClientObject {
    /**
     *
     * Represents binding identifier. Read-only.
     *
     * [Api set: ExcelApi 1.1]
     */
    id: string;
    /**
     *
     * Returns the type of the binding. See BindingType for details. Read-only.
     *
     * [Api set: ExcelApi 1.1]
     */
    type: string;
    /**
     *
     * Deletes the binding.
     *
     * [Api set: ExcelApi 1.3]
     */
    delete(): void;
    /**
     *
     * Returns the range represented by the binding. Will throw an error if binding is not of the correct type.
     *
     * [Api set: ExcelApi 1.1]
     */
    getRange(): Range;
    /**
     *
     * Returns the table represented by the binding. Will throw an error if binding is not of the correct type.
     *
     * [Api set: ExcelApi 1.1]
     */
    getTable(): Table;
    /**
     *
     * Returns the text represented by the binding. Will throw an error if binding is not of the correct type.
     *
     * [Api set: ExcelApi 1.1]
     */
    getText(): OfficeExtension.ClientResult<string>;
    /**
     * Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
     */
    load(option?: string | string[] | OfficeExtension.LoadOption): Binding;
    /**
     *
     * Occurs when data or formatting within the binding is changed.
     *
     * [Api set: ExcelApi 1.2]
     */
    onDataChanged: OfficeExtension.EventHandlers<BindingDataChangedEventArgs>;
    /**
     *
     * Occurs when the selection is changed within the binding.
     *
     * [Api set: ExcelApi 1.2]
     */
    onSelectionChanged: OfficeExtension.EventHandlers<BindingSelectionChangedEventArgs>;
    toJSON(): {
        "id": string;
        "type": string;
    };
}
/**
 *
 * Represents the collection of all the binding objects that are part of the workbook.
 *
 * [Api set: ExcelApi 1.1]
 */
export declare class BindingCollection extends OfficeExtension.ClientObject {
    /** Gets the loaded child items in this collection. */
    items: Array<Binding>;
    /**
     *
     * Returns the number of bindings in the collection. Read-only.
     *
     * [Api set: ExcelApi 1.1]
     */
    count: number;
    /**
     *
     * Add a new binding to a particular Range.
     *
     * [Api set: ExcelApi 1.3]
     *
     * @param range Range to bind the binding to. May be an Excel Range object, or a string. If string, must contain the full address, including the sheet name
     * @param bindingType Type of binding. See BindingType.
     * @param id Name of binding.
     */
    add(range: Range | string, bindingType: string, id: string): Binding;
    /**
     *
     * Add a new binding based on a named item in the workbook.
     *
     * [Api set: ExcelApi 1.3]
     *
     * @param name Name from which to create binding.
     * @param bindingType Type of binding. See BindingType.
     * @param id Name of binding.
     */
    addFromNamedItem(name: string, bindingType: string, id: string): Binding;
    /**
     *
     * Add a new binding based on the current selection.
     *
     * [Api set: ExcelApi 1.3]
     *
     * @param bindingType Type of binding. See BindingType.
     * @param id Name of binding.
     */
    addFromSelection(bindingType: string, id: string): Binding;
    /**
     *
     * Gets a binding object by ID.
     *
     * [Api set: ExcelApi 1.1]
     *
     * @param id Id of the binding object to be retrieved.
     */
    getItem(id: string): Binding;
    /**
     *
     * Gets a binding object based on its position in the items array.
     *
     * [Api set: ExcelApi 1.1]
     *
     * @param index Index value of the object to be retrieved. Zero-indexed.
     */
    getItemAt(index: number): Binding;
    /**
     *
     * Gets a binding object by ID. If the binding object does not exist, will return a null object.
     *
     * [Api set: ExcelApi 1.4]
     *
     * @param id Id of the binding object to be retrieved.
     */
    getItemOrNullObject(id: string): Binding;
    /**
     * Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
     */
    load(option?: string | string[] | OfficeExtension.LoadOption): BindingCollection;
    toJSON(): {
        "count": number;
    };
}
/**
 *
 * Represents a collection of all the tables that are part of the workbook or worksheet, depending on how it was reached.
 *
 * [Api set: ExcelApi 1.1]
 */
export declare class TableCollection extends OfficeExtension.ClientObject {
    /** Gets the loaded child items in this collection. */
    items: Array<Table>;
    /**
     *
     * Returns the number of tables in the workbook. Read-only.
     *
     * [Api set: ExcelApi 1.1]
     */
    count: number;
    /**
     *
     * Create a new table. The range object or source address determines the worksheet under which the table will be added. If the table cannot be added (e.g., because the address is invalid, or the table would overlap with another table), an error will be thrown.
     *
     * [Api set: ExcelApi 1.1]
     *
     * @param address A Range object, or a string address or name of the range representing the data source. If the address does not contain a sheet name, the currently-active sheet is used. [Api set: ExcelApi 1.1 for string parameter; 1.3 for accepting a Range object as well]
     * @param hasHeaders Boolean value that indicates whether the data being imported has column labels. If the source does not contain headers (i.e,. when this property set to false), Excel will automatically generate header shifting the data down by one row.
     */
    add(address: Range | string, hasHeaders: boolean): Table;
    /**
     *
     * Gets a table by Name or ID.
     *
     * [Api set: ExcelApi 1.1]
     *
     * @param key Name or ID of the table to be retrieved.
     */
    getItem(key: number | string): Table;
    /**
     *
     * Gets a table based on its position in the collection.
     *
     * [Api set: ExcelApi 1.1]
     *
     * @param index Index value of the object to be retrieved. Zero-indexed.
     */
    getItemAt(index: number): Table;
    /**
     *
     * Gets a table by Name or ID. If the table does not exist, will return a null object.
     *
     * [Api set: ExcelApi 1.4]
     *
     * @param key Name or ID of the table to be retrieved.
     */
    getItemOrNullObject(key: number | string): Table;
    /**
     * Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
     */
    load(option?: string | string[] | OfficeExtension.LoadOption): TableCollection;
    toJSON(): {
        "count": number;
    };
}
/**
 *
 * Represents an Excel table.
 *
 * [Api set: ExcelApi 1.1]
 */
export declare class Table extends OfficeExtension.ClientObject {
    /**
     *
     * Represents a collection of all the columns in the table. Read-only.
     *
     * [Api set: ExcelApi 1.1]
     */
    columns: TableColumnCollection;
    /**
     *
     * Represents a collection of all the rows in the table. Read-only.
     *
     * [Api set: ExcelApi 1.1]
     */
    rows: TableRowCollection;
    /**
     *
     * Represents the sorting for the table.
     *
     * [Api set: ExcelApi 1.2]
     */
    sort: TableSort;
    /**
     *
     * The worksheet containing the current table. Read-only.
     *
     * [Api set: ExcelApi 1.2]
     */
    worksheet: Worksheet;
    /**
     *
     * Indicates whether the first column contains special formatting.
     *
     * [Api set: ExcelApi 1.3]
     */
    highlightFirstColumn: boolean;
    /**
     *
     * Indicates whether the last column contains special formatting.
     *
     * [Api set: ExcelApi 1.3]
     */
    highlightLastColumn: boolean;
    /**
     *
     * Returns a value that uniquely identifies the table in a given workbook. The value of the identifier remains the same even when the table is renamed. Read-only.
     *
     * [Api set: ExcelApi 1.1]
     */
    id: number;
    /**
     *
     * Name of the table.
     *
     * [Api set: ExcelApi 1.1]
     */
    name: string;
    /**
     *
     * Indicates whether the columns show banded formatting in which odd columns are highlighted differently from even ones to make reading the table easier.
     *
     * [Api set: ExcelApi 1.3]
     */
    showBandedColumns: boolean;
    /**
     *
     * Indicates whether the rows show banded formatting in which odd rows are highlighted differently from even ones to make reading the table easier.
     *
     * [Api set: ExcelApi 1.3]
     */
    showBandedRows: boolean;
    /**
     *
     * Indicates whether the filter buttons are visible at the top of each column header. Setting this is only allowed if the table contains a header row.
     *
     * [Api set: ExcelApi 1.3]
     */
    showFilterButton: boolean;
    /**
     *
     * Indicates whether the header row is visible or not. This value can be set to show or remove the header row.
     *
     * [Api set: ExcelApi 1.1]
     */
    showHeaders: boolean;
    /**
     *
     * Indicates whether the total row is visible or not. This value can be set to show or remove the total row.
     *
     * [Api set: ExcelApi 1.1]
     */
    showTotals: boolean;
    /**
     *
     * Constant value that represents the Table style. Possible values are: TableStyleLight1 thru TableStyleLight21, TableStyleMedium1 thru TableStyleMedium28, TableStyleStyleDark1 thru TableStyleStyleDark11. A custom user-defined style present in the workbook can also be specified.
     *
     * [Api set: ExcelApi 1.1]
     */
    style: string;
    /**
     *
     * Clears all the filters currently applied on the table.
     *
     * [Api set: ExcelApi 1.2]
     */
    clearFilters(): void;
    /**
     *
     * Converts the table into a normal range of cells. All data is preserved.
     *
     * [Api set: ExcelApi 1.2]
     */
    convertToRange(): Range;
    /**
     *
     * Deletes the table.
     *
     * [Api set: ExcelApi 1.1]
     */
    delete(): void;
    /**
     *
     * Gets the range object associated with the data body of the table.
     *
     * [Api set: ExcelApi 1.1]
     */
    getDataBodyRange(): Range;
    /**
     *
     * Gets the range object associated with header row of the table.
     *
     * [Api set: ExcelApi 1.1]
     */
    getHeaderRowRange(): Range;
    /**
     *
     * Gets the range object associated with the entire table.
     *
     * [Api set: ExcelApi 1.1]
     */
    getRange(): Range;
    /**
     *
     * Gets the range object associated with totals row of the table.
     *
     * [Api set: ExcelApi 1.1]
     */
    getTotalRowRange(): Range;
    /**
     *
     * Reapplies all the filters currently on the table.
     *
     * [Api set: ExcelApi 1.2]
     */
    reapplyFilters(): void;
    /**
     * Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
     */
    load(option?: string | string[] | OfficeExtension.LoadOption): Table;
    toJSON(): {
        "highlightFirstColumn": boolean;
        "highlightLastColumn": boolean;
        "id": number;
        "name": string;
        "showBandedColumns": boolean;
        "showBandedRows": boolean;
        "showFilterButton": boolean;
        "showHeaders": boolean;
        "showTotals": boolean;
        "style": string;
    };
}
/**
 *
 * Represents a collection of all the columns that are part of the table.
 *
 * [Api set: ExcelApi 1.1]
 */
export declare class TableColumnCollection extends OfficeExtension.ClientObject {
    /** Gets the loaded child items in this collection. */
    items: Array<TableColumn>;
    /**
     *
     * Returns the number of columns in the table. Read-only.
     *
     * [Api set: ExcelApi 1.1]
     */
    count: number;
    /**
     *
     * Adds a new column to the table.
     *
     * [Api set: ExcelApi 1.1 requires an index smaller than the total column count; 1.4 allows index to be optional (null or -1) and will append a column at the end; 1.4 allows name parameter at creation time.]
     *
     * @param index Specifies the relative position of the new column. If null or -1, the addition happens at the end. Columns with a higher index will be shifted to the side. Zero-indexed.
     * @param values A 2-dimensional array of unformatted values of the table column.
     * @param name Specifies the name of the new column. If null, the default name will be used.
     */
    add(index?: number, values?: Array<Array<boolean | string | number>> | boolean | string | number, name?: string): TableColumn;
    /**
     *
     * Gets a column object by Name or ID.
     *
     * [Api set: ExcelApi 1.1]
     *
     * @param key Column Name or ID.
     */
    getItem(key: number | string): TableColumn;
    /**
     *
     * Gets a column based on its position in the collection.
     *
     * [Api set: ExcelApi 1.1]
     *
     * @param index Index value of the object to be retrieved. Zero-indexed.
     */
    getItemAt(index: number): TableColumn;
    /**
     *
     * Gets a column object by Name or ID. If the column does not exist, will return a null object.
     *
     * [Api set: ExcelApi 1.4]
     *
     * @param key Column Name or ID.
     */
    getItemOrNullObject(key: number | string): TableColumn;
    /**
     * Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
     */
    load(option?: string | string[] | OfficeExtension.LoadOption): TableColumnCollection;
    toJSON(): {
        "count": number;
    };
}
/**
 *
 * Represents a column in a table.
 *
 * [Api set: ExcelApi 1.1]
 */
export declare class TableColumn extends OfficeExtension.ClientObject {
    /**
     *
     * Retrieve the filter applied to the column.
     *
     * [Api set: ExcelApi 1.2]
     */
    filter: Filter;
    /**
     *
     * Returns a unique key that identifies the column within the table. Read-only.
     *
     * [Api set: ExcelApi 1.1]
     */
    id: number;
    /**
     *
     * Returns the index number of the column within the columns collection of the table. Zero-indexed. Read-only.
     *
     * [Api set: ExcelApi 1.1]
     */
    index: number;
    /**
     *
     * Represents the name of the table column.
     *
     * [Api set: ExcelApi 1.1 for getting the name; 1.4 for setting it.]
     */
    name: string;
    /**
     *
     * Represents the raw values of the specified range. The data returned could be of type string, number, or a boolean. Cell that contain an error will return the error string.
     *
     * [Api set: ExcelApi 1.1]
     */
    values: Array<Array<any>>;
    /**
     *
     * Deletes the column from the table.
     *
     * [Api set: ExcelApi 1.1]
     */
    delete(): void;
    /**
     *
     * Gets the range object associated with the data body of the column.
     *
     * [Api set: ExcelApi 1.1]
     */
    getDataBodyRange(): Range;
    /**
     *
     * Gets the range object associated with the header row of the column.
     *
     * [Api set: ExcelApi 1.1]
     */
    getHeaderRowRange(): Range;
    /**
     *
     * Gets the range object associated with the entire column.
     *
     * [Api set: ExcelApi 1.1]
     */
    getRange(): Range;
    /**
     *
     * Gets the range object associated with the totals row of the column.
     *
     * [Api set: ExcelApi 1.1]
     */
    getTotalRowRange(): Range;
    /**
     * Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
     */
    load(option?: string | string[] | OfficeExtension.LoadOption): TableColumn;
    toJSON(): {
        "id": number;
        "index": number;
        "name": string;
        "values": any[][];
    };
}
/**
 *
 * Represents a collection of all the rows that are part of the table.
 *
 * [Api set: ExcelApi 1.1]
 */
export declare class TableRowCollection extends OfficeExtension.ClientObject {
    /** Gets the loaded child items in this collection. */
    items: Array<TableRow>;
    /**
     *
     * Returns the number of rows in the table. Read-only.
     *
     * [Api set: ExcelApi 1.1]
     */
    count: number;
    /**
     *
     * Adds one or more rows to the table. The return object will be the top of the newly added row(s).
     *
     * [Api set: ExcelApi 1.1 for adding a single row; 1.4 allows adding of multiple rows.]
     *
     * @param index Specifies the relative position of the new row. If null or -1, the addition happens at the end. Any rows below the inserted row are shifted downwards. Zero-indexed.
     * @param values A 2-dimensional array of unformatted values of the table row.
     */
    add(index?: number, values?: Array<Array<boolean | string | number>> | boolean | string | number): TableRow;
    /**
     *
     * Gets a row based on its position in the collection.
     *
     * [Api set: ExcelApi 1.1]
     *
     * @param index Index value of the object to be retrieved. Zero-indexed.
     */
    getItemAt(index: number): TableRow;
    /**
     * Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
     */
    load(option?: string | string[] | OfficeExtension.LoadOption): TableRowCollection;
    toJSON(): {
        "count": number;
    };
}
/**
 *
 * Represents a row in a table.
 *
 * [Api set: ExcelApi 1.1]
 */
export declare class TableRow extends OfficeExtension.ClientObject {
    /**
     *
     * Returns the index number of the row within the rows collection of the table. Zero-indexed. Read-only.
     *
     * [Api set: ExcelApi 1.1]
     */
    index: number;
    /**
     *
     * Represents the raw values of the specified range. The data returned could be of type string, number, or a boolean. Cell that contain an error will return the error string.
     *
     * [Api set: ExcelApi 1.1]
     */
    values: Array<Array<any>>;
    /**
     *
     * Deletes the row from the table.
     *
     * [Api set: ExcelApi 1.1]
     */
    delete(): void;
    /**
     *
     * Returns the range object associated with the entire row.
     *
     * [Api set: ExcelApi 1.1]
     */
    getRange(): Range;
    /**
     * Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
     */
    load(option?: string | string[] | OfficeExtension.LoadOption): TableRow;
    toJSON(): {
        "index": number;
        "values": any[][];
    };
}
/**
 *
 * A format object encapsulating the range's font, fill, borders, alignment, and other properties.
 *
 * [Api set: ExcelApi 1.1]
 */
export declare class RangeFormat extends OfficeExtension.ClientObject {
    /**
     *
     * Collection of border objects that apply to the overall range. Read-only.
     *
     * [Api set: ExcelApi 1.1]
     */
    borders: RangeBorderCollection;
    /**
     *
     * Returns the fill object defined on the overall range. Read-only.
     *
     * [Api set: ExcelApi 1.1]
     */
    fill: RangeFill;
    /**
     *
     * Returns the font object defined on the overall range. Read-only.
     *
     * [Api set: ExcelApi 1.1]
     */
    font: RangeFont;
    /**
     *
     * Returns the format protection object for a range.
     *
     * [Api set: ExcelApi 1.2]
     */
    protection: FormatProtection;
    /**
     *
     * Gets or sets the width of all colums within the range. If the column widths are not uniform, null will be returned.
     *
     * [Api set: ExcelApi 1.2]
     */
    columnWidth: number;
    /**
     *
     * Represents the horizontal alignment for the specified object. See HorizontalAlignment for details.
     *
     * [Api set: ExcelApi 1.1]
     */
    horizontalAlignment: string;
    /**
     *
     * Gets or sets the height of all rows in the range. If the row heights are not uniform null will be returned.
     *
     * [Api set: ExcelApi 1.2]
     */
    rowHeight: number;
    /**
     *
     * Represents the vertical alignment for the specified object. See VerticalAlignment for details.
     *
     * [Api set: ExcelApi 1.1]
     */
    verticalAlignment: string;
    /**
     *
     * Indicates if Excel wraps the text in the object. A null value indicates that the entire range doesn't have uniform wrap setting
     *
     * [Api set: ExcelApi 1.1]
     */
    wrapText: boolean;
    /**
     *
     * Changes the width of the columns of the current range to achieve the best fit, based on the current data in the columns.
     *
     * [Api set: ExcelApi 1.2]
     */
    autofitColumns(): void;
    /**
     *
     * Changes the height of the rows of the current range to achieve the best fit, based on the current data in the columns.
     *
     * [Api set: ExcelApi 1.2]
     */
    autofitRows(): void;
    /**
     * Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
     */
    load(option?: string | string[] | OfficeExtension.LoadOption): RangeFormat;
    toJSON(): {
        "columnWidth": number;
        "fill": RangeFill;
        "font": RangeFont;
        "horizontalAlignment": string;
        "protection": FormatProtection;
        "rowHeight": number;
        "verticalAlignment": string;
        "wrapText": boolean;
    };
}
/**
 *
 * Represents the format protection of a range object.
 *
 * [Api set: ExcelApi 1.2]
 */
export declare class FormatProtection extends OfficeExtension.ClientObject {
    /**
     *
     * Indicates if Excel hides the formula for the cells in the range. A null value indicates that the entire range doesn't have uniform formula hidden setting.
     *
     * [Api set: ExcelApi 1.2]
     */
    formulaHidden: boolean;
    /**
     *
     * Indicates if Excel locks the cells in the object. A null value indicates that the entire range doesn't have uniform lock setting.
     *
     * [Api set: ExcelApi 1.2]
     */
    locked: boolean;
    /**
     * Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
     */
    load(option?: string | string[] | OfficeExtension.LoadOption): FormatProtection;
    toJSON(): {
        "formulaHidden": boolean;
        "locked": boolean;
    };
}
/**
 *
 * Represents the background of a range object.
 *
 * [Api set: ExcelApi 1.1]
 */
export declare class RangeFill extends OfficeExtension.ClientObject {
    /**
     *
     * HTML color code representing the color of the border line, of the form #RRGGBB (e.g. "FFA500") or as a named HTML color (e.g. "orange")
     *
     * [Api set: ExcelApi 1.1]
     */
    color: string;
    /**
     *
     * Resets the range background.
     *
     * [Api set: ExcelApi 1.1]
     */
    clear(): void;
    /**
     * Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
     */
    load(option?: string | string[] | OfficeExtension.LoadOption): RangeFill;
    toJSON(): {
        "color": string;
    };
}
/**
 *
 * Represents the border of an object.
 *
 * [Api set: ExcelApi 1.1]
 */
export declare class RangeBorder extends OfficeExtension.ClientObject {
    /**
     *
     * HTML color code representing the color of the border line, of the form #RRGGBB (e.g. "FFA500") or as a named HTML color (e.g. "orange").
     *
     * [Api set: ExcelApi 1.1]
     */
    color: string;
    /**
     *
     * Constant value that indicates the specific side of the border. See BorderIndex for details. Read-only.
     *
     * [Api set: ExcelApi 1.1]
     */
    sideIndex: string;
    /**
     *
     * One of the constants of line style specifying the line style for the border. See BorderLineStyle for details.
     *
     * [Api set: ExcelApi 1.1]
     */
    style: string;
    /**
     *
     * Specifies the weight of the border around a range. See BorderWeight for details.
     *
     * [Api set: ExcelApi 1.1]
     */
    weight: string;
    /**
     * Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
     */
    load(option?: string | string[] | OfficeExtension.LoadOption): RangeBorder;
    toJSON(): {
        "color": string;
        "sideIndex": string;
        "style": string;
        "weight": string;
    };
}
/**
 *
 * Represents the border objects that make up range border.
 *
 * [Api set: ExcelApi 1.1]
 */
export declare class RangeBorderCollection extends OfficeExtension.ClientObject {
    /** Gets the loaded child items in this collection. */
    items: Array<RangeBorder>;
    /**
     *
     * Number of border objects in the collection. Read-only.
     *
     * [Api set: ExcelApi 1.1]
     */
    count: number;
    /**
     *
     * Gets a border object using its name
     *
     * [Api set: ExcelApi 1.1]
     *
     * @param index Index value of the border object to be retrieved. See BorderIndex for details.
     */
    getItem(index: string): RangeBorder;
    /**
     *
     * Gets a border object using its index
     *
     * [Api set: ExcelApi 1.1]
     *
     * @param index Index value of the object to be retrieved. Zero-indexed.
     */
    getItemAt(index: number): RangeBorder;
    /**
     * Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
     */
    load(option?: string | string[] | OfficeExtension.LoadOption): RangeBorderCollection;
    toJSON(): {
        "count": number;
    };
}
/**
 *
 * This object represents the font attributes (font name, font size, color, etc.) for an object.
 *
 * [Api set: ExcelApi 1.1]
 */
export declare class RangeFont extends OfficeExtension.ClientObject {
    /**
     *
     * Represents the bold status of font.
     *
     * [Api set: ExcelApi 1.1]
     */
    bold: boolean;
    /**
     *
     * HTML color code representation of the text color. E.g. #FF0000 represents Red.
     *
     * [Api set: ExcelApi 1.1]
     */
    color: string;
    /**
     *
     * Represents the italic status of the font.
     *
     * [Api set: ExcelApi 1.1]
     */
    italic: boolean;
    /**
     *
     * Font name (e.g. "Calibri")
     *
     * [Api set: ExcelApi 1.1]
     */
    name: string;
    /**
     *
     * Font size.
     *
     * [Api set: ExcelApi 1.1]
     */
    size: number;
    /**
     *
     * Type of underline applied to the font. See RangeUnderlineStyle for details.
     *
     * [Api set: ExcelApi 1.1]
     */
    underline: string;
    /**
     * Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
     */
    load(option?: string | string[] | OfficeExtension.LoadOption): RangeFont;
    toJSON(): {
        "bold": boolean;
        "color": string;
        "italic": boolean;
        "name": string;
        "size": number;
        "underline": string;
    };
}
/**
 *
 * A collection of all the chart objects on a worksheet.
 *
 * [Api set: ExcelApi 1.1]
 */
export declare class ChartCollection extends OfficeExtension.ClientObject {
    /** Gets the loaded child items in this collection. */
    items: Array<Chart>;
    /**
     *
     * Returns the number of charts in the worksheet. Read-only.
     *
     * [Api set: ExcelApi 1.1]
     */
    count: number;
    /**
     *
     * Creates a new chart.
     *
     * [Api set: ExcelApi 1.1]
     *
     * @param type Represents the type of a chart. See ChartType for details.
     * @param sourceData The Range object corresponding to the source data.
     * @param seriesBy Specifies the way columns or rows are used as data series on the chart. See ChartSeriesBy for details.
     */
    add(type: string, sourceData: Range, seriesBy?: string): Chart;
    /**
     *
     * Gets a chart using its name. If there are multiple charts with the same name, the first one will be returned.
     *
     * [Api set: ExcelApi 1.1]
     *
     * @param name Name of the chart to be retrieved.
     */
    getItem(name: string): Chart;
    /**
     *
     * Gets a chart based on its position in the collection.
     *
     * [Api set: ExcelApi 1.1]
     *
     * @param index Index value of the object to be retrieved. Zero-indexed.
     */
    getItemAt(index: number): Chart;
    /**
     *
     * Gets a chart using its name. If there are multiple charts with the same name, the first one will be returned.
        If the chart does not exist, will return a null object.
     *
     * [Api set: ExcelApi 1.4]
     *
     * @param name Name of the chart to be retrieved.
     */
    getItemOrNullObject(name: string): Chart;
    /**
     * Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
     */
    load(option?: string | string[] | OfficeExtension.LoadOption): ChartCollection;
    toJSON(): {
        "count": number;
    };
}
/**
 *
 * Represents a chart object in a workbook.
 *
 * [Api set: ExcelApi 1.1]
 */
export declare class Chart extends OfficeExtension.ClientObject {
    /**
     *
     * Represents chart axes. Read-only.
     *
     * [Api set: ExcelApi 1.1]
     */
    axes: ChartAxes;
    /**
     *
     * Represents the datalabels on the chart. Read-only.
     *
     * [Api set: ExcelApi 1.1]
     */
    dataLabels: ChartDataLabels;
    /**
     *
     * Encapsulates the format properties for the chart area. Read-only.
     *
     * [Api set: ExcelApi 1.1]
     */
    format: ChartAreaFormat;
    /**
     *
     * Represents the legend for the chart. Read-only.
     *
     * [Api set: ExcelApi 1.1]
     */
    legend: ChartLegend;
    /**
     *
     * Represents either a single series or collection of series in the chart. Read-only.
     *
     * [Api set: ExcelApi 1.1]
     */
    series: ChartSeriesCollection;
    /**
     *
     * Represents the title of the specified chart, including the text, visibility, position and formating of the title. Read-only.
     *
     * [Api set: ExcelApi 1.1]
     */
    title: ChartTitle;
    /**
     *
     * The worksheet containing the current chart. Read-only.
     *
     * [Api set: ExcelApi 1.2]
     */
    worksheet: Worksheet;
    /**
     *
     * Represents the height, in points, of the chart object.
     *
     * [Api set: ExcelApi 1.1]
     */
    height: number;
    /**
     *
     * The distance, in points, from the left side of the chart to the worksheet origin.
     *
     * [Api set: ExcelApi 1.1]
     */
    left: number;
    /**
     *
     * Represents the name of a chart object.
     *
     * [Api set: ExcelApi 1.1]
     */
    name: string;
    /**
     *
     * Represents the distance, in points, from the top edge of the object to the top of row 1 (on a worksheet) or the top of the chart area (on a chart).
     *
     * [Api set: ExcelApi 1.1]
     */
    top: number;
    /**
     *
     * Represents the width, in points, of the chart object.
     *
     * [Api set: ExcelApi 1.1]
     */
    width: number;
    /**
     *
     * Deletes the chart object.
     *
     * [Api set: ExcelApi 1.1]
     */
    delete(): void;
    /**
     *
     * Renders the chart as a base64-encoded image by scaling the chart to fit the specified dimensions.
        The aspect ratio is preserved as part of the resizing.
     *
     * [Api set: ExcelApi 1.2]
     *
     * @param height (Optional) The desired height of the resulting image.
     * @param width (Optional) The desired width of the resulting image.
     * @param fittingMode (Optional) The method used to scale the chart to the specified to the specified dimensions (if both height and width are set)."
     */
    getImage(width?: number, height?: number, fittingMode?: string): OfficeExtension.ClientResult<string>;
    /**
     *
     * Resets the source data for the chart.
     *
     * [Api set: ExcelApi 1.1]
     *
     * @param sourceData The Range object corresponding to the source data.
     * @param seriesBy Specifies the way columns or rows are used as data series on the chart. Can be one of the following: Auto (default), Rows, Columns. See ChartSeriesBy for details.
     */
    setData(sourceData: Range, seriesBy?: string): void;
    /**
     *
     * Positions the chart relative to cells on the worksheet.
     *
     * [Api set: ExcelApi 1.1]
     *
     * @param startCell The start cell. This is where the chart will be moved to. The start cell is the top-left or top-right cell, depending on the user's right-to-left display settings.
     * @param endCell (Optional) The end cell. If specified, the chart's width and height will be set to fully cover up this cell/range.
     */
    setPosition(startCell: Range | string, endCell?: Range | string): void;
    /**
     * Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
     */
    load(option?: string | string[] | OfficeExtension.LoadOption): Chart;
    toJSON(): {
        "axes": ChartAxes;
        "dataLabels": ChartDataLabels;
        "format": ChartAreaFormat;
        "height": number;
        "left": number;
        "legend": ChartLegend;
        "name": string;
        "title": ChartTitle;
        "top": number;
        "width": number;
    };
}
/**
 *
 * Encapsulates the format properties for the overall chart area.
 *
 * [Api set: ExcelApi 1.1]
 */
export declare class ChartAreaFormat extends OfficeExtension.ClientObject {
    /**
     *
     * Represents the fill format of an object, which includes background formatting information. Read-only.
     *
     * [Api set: ExcelApi 1.1]
     */
    fill: ChartFill;
    /**
     *
     * Represents the font attributes (font name, font size, color, etc.) for the current object. Read-only.
     *
     * [Api set: ExcelApi 1.1]
     */
    font: ChartFont;
    /**
     * Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
     */
    load(option?: string | string[] | OfficeExtension.LoadOption): ChartAreaFormat;
    toJSON(): {
        "fill": ChartFill;
        "font": ChartFont;
    };
}
/**
 *
 * Represents a collection of chart series.
 *
 * [Api set: ExcelApi 1.1]
 */
export declare class ChartSeriesCollection extends OfficeExtension.ClientObject {
    /** Gets the loaded child items in this collection. */
    items: Array<ChartSeries>;
    /**
     *
     * Returns the number of series in the collection. Read-only.
     *
     * [Api set: ExcelApi 1.1]
     */
    count: number;
    /**
     *
     * Retrieves a series based on its position in the collection
     *
     * [Api set: ExcelApi 1.1]
     *
     * @param index Index value of the object to be retrieved. Zero-indexed.
     */
    getItemAt(index: number): ChartSeries;
    /**
     * Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
     */
    load(option?: string | string[] | OfficeExtension.LoadOption): ChartSeriesCollection;
    toJSON(): {
        "count": number;
    };
}
/**
 *
 * Represents a series in a chart.
 *
 * [Api set: ExcelApi 1.1]
 */
export declare class ChartSeries extends OfficeExtension.ClientObject {
    /**
     *
     * Represents the formatting of a chart series, which includes fill and line formatting. Read-only.
     *
     * [Api set: ExcelApi 1.1]
     */
    format: ChartSeriesFormat;
    /**
     *
     * Represents a collection of all points in the series. Read-only.
     *
     * [Api set: ExcelApi 1.1]
     */
    points: ChartPointsCollection;
    /**
     *
     * Represents the name of a series in a chart.
     *
     * [Api set: ExcelApi 1.1]
     */
    name: string;
    /**
     * Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
     */
    load(option?: string | string[] | OfficeExtension.LoadOption): ChartSeries;
    toJSON(): {
        "format": ChartSeriesFormat;
        "name": string;
    };
}
/**
 *
 * encapsulates the format properties for the chart series
 *
 * [Api set: ExcelApi 1.1]
 */
export declare class ChartSeriesFormat extends OfficeExtension.ClientObject {
    /**
     *
     * Represents the fill format of a chart series, which includes background formating information. Read-only.
     *
     * [Api set: ExcelApi 1.1]
     */
    fill: ChartFill;
    /**
     *
     * Represents line formatting. Read-only.
     *
     * [Api set: ExcelApi 1.1]
     */
    line: ChartLineFormat;
    /**
     * Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
     */
    load(option?: string | string[] | OfficeExtension.LoadOption): ChartSeriesFormat;
    toJSON(): {
        "fill": ChartFill;
        "line": ChartLineFormat;
    };
}
/**
 *
 * A collection of all the chart points within a series inside a chart.
 *
 * [Api set: ExcelApi 1.1]
 */
export declare class ChartPointsCollection extends OfficeExtension.ClientObject {
    /** Gets the loaded child items in this collection. */
    items: Array<ChartPoint>;
    /**
     *
     * Returns the number of chart points in the collection. Read-only.
     *
     * [Api set: ExcelApi 1.1]
     */
    count: number;
    /**
     *
     * Retrieve a point based on its position within the series.
     *
     * [Api set: ExcelApi 1.1]
     *
     * @param index Index value of the object to be retrieved. Zero-indexed.
     */
    getItemAt(index: number): ChartPoint;
    /**
     * Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
     */
    load(option?: string | string[] | OfficeExtension.LoadOption): ChartPointsCollection;
    toJSON(): {
        "count": number;
    };
}
/**
 *
 * Represents a point of a series in a chart.
 *
 * [Api set: ExcelApi 1.1]
 */
export declare class ChartPoint extends OfficeExtension.ClientObject {
    /**
     *
     * Encapsulates the format properties chart point. Read-only.
     *
     * [Api set: ExcelApi 1.1]
     */
    format: ChartPointFormat;
    /**
     *
     * Returns the value of a chart point. Read-only.
     *
     * [Api set: ExcelApi 1.1]
     */
    value: any;
    /**
     * Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
     */
    load(option?: string | string[] | OfficeExtension.LoadOption): ChartPoint;
    toJSON(): {
        "format": ChartPointFormat;
        "value": any;
    };
}
/**
 *
 * Represents formatting object for chart points.
 *
 * [Api set: ExcelApi 1.1]
 */
export declare class ChartPointFormat extends OfficeExtension.ClientObject {
    /**
     *
     * Represents the fill format of a chart, which includes background formating information. Read-only.
     *
     * [Api set: ExcelApi 1.1]
     */
    fill: ChartFill;
    /**
     * Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
     */
    load(option?: string | string[] | OfficeExtension.LoadOption): ChartPointFormat;
    toJSON(): {
        "fill": ChartFill;
    };
}
/**
 *
 * Represents the chart axes.
 *
 * [Api set: ExcelApi 1.1]
 */
export declare class ChartAxes extends OfficeExtension.ClientObject {
    /**
     *
     * Represents the category axis in a chart. Read-only.
     *
     * [Api set: ExcelApi 1.1]
     */
    categoryAxis: ChartAxis;
    /**
     *
     * Represents the series axis of a 3-dimensional chart. Read-only.
     *
     * [Api set: ExcelApi 1.1]
     */
    seriesAxis: ChartAxis;
    /**
     *
     * Represents the value axis in an axis. Read-only.
     *
     * [Api set: ExcelApi 1.1]
     */
    valueAxis: ChartAxis;
    /**
     * Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
     */
    load(option?: string | string[] | OfficeExtension.LoadOption): ChartAxes;
    toJSON(): {
        "categoryAxis": ChartAxis;
        "seriesAxis": ChartAxis;
        "valueAxis": ChartAxis;
    };
}
/**
 *
 * Represents a single axis in a chart.
 *
 * [Api set: ExcelApi 1.1]
 */
export declare class ChartAxis extends OfficeExtension.ClientObject {
    /**
     *
     * Represents the formatting of a chart object, which includes line and font formatting. Read-only.
     *
     * [Api set: ExcelApi 1.1]
     */
    format: ChartAxisFormat;
    /**
     *
     * Returns a gridlines object that represents the major gridlines for the specified axis. Read-only.
     *
     * [Api set: ExcelApi 1.1]
     */
    majorGridlines: ChartGridlines;
    /**
     *
     * Returns a Gridlines object that represents the minor gridlines for the specified axis. Read-only.
     *
     * [Api set: ExcelApi 1.1]
     */
    minorGridlines: ChartGridlines;
    /**
     *
     * Represents the axis title. Read-only.
     *
     * [Api set: ExcelApi 1.1]
     */
    title: ChartAxisTitle;
    /**
     *
     * Represents the interval between two major tick marks. Can be set to a numeric value or an empty string.  The returned value is always a number.
     *
     * [Api set: ExcelApi 1.1]
     */
    majorUnit: any;
    /**
     *
     * Represents the maximum value on the value axis.  Can be set to a numeric value or an empty string (for automatic axis values).  The returned value is always a number.
     *
     * [Api set: ExcelApi 1.1]
     */
    maximum: any;
    /**
     *
     * Represents the minimum value on the value axis. Can be set to a numeric value or an empty string (for automatic axis values).  The returned value is always a number.
     *
     * [Api set: ExcelApi 1.1]
     */
    minimum: any;
    /**
     *
     * Represents the interval between two minor tick marks. "Can be set to a numeric value or an empty string (for automatic axis values). The returned value is always a number.
     *
     * [Api set: ExcelApi 1.1]
     */
    minorUnit: any;
    /**
     * Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
     */
    load(option?: string | string[] | OfficeExtension.LoadOption): ChartAxis;
    toJSON(): {
        "format": ChartAxisFormat;
        "majorGridlines": ChartGridlines;
        "majorUnit": any;
        "maximum": any;
        "minimum": any;
        "minorGridlines": ChartGridlines;
        "minorUnit": any;
        "title": ChartAxisTitle;
    };
}
/**
 *
 * Encapsulates the format properties for the chart axis.
 *
 * [Api set: ExcelApi 1.1]
 */
export declare class ChartAxisFormat extends OfficeExtension.ClientObject {
    /**
     *
     * Represents the font attributes (font name, font size, color, etc.) for a chart axis element. Read-only.
     *
     * [Api set: ExcelApi 1.1]
     */
    font: ChartFont;
    /**
     *
     * Represents chart line formatting. Read-only.
     *
     * [Api set: ExcelApi 1.1]
     */
    line: ChartLineFormat;
    /**
     * Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
     */
    load(option?: string | string[] | OfficeExtension.LoadOption): ChartAxisFormat;
    toJSON(): {
        "font": ChartFont;
        "line": ChartLineFormat;
    };
}
/**
 *
 * Represents the title of a chart axis.
 *
 * [Api set: ExcelApi 1.1]
 */
export declare class ChartAxisTitle extends OfficeExtension.ClientObject {
    /**
     *
     * Represents the formatting of chart axis title. Read-only.
     *
     * [Api set: ExcelApi 1.1]
     */
    format: ChartAxisTitleFormat;
    /**
     *
     * Represents the axis title.
     *
     * [Api set: ExcelApi 1.1]
     */
    text: string;
    /**
     *
     * A boolean that specifies the visibility of an axis title.
     *
     * [Api set: ExcelApi 1.1]
     */
    visible: boolean;
    /**
     * Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
     */
    load(option?: string | string[] | OfficeExtension.LoadOption): ChartAxisTitle;
    toJSON(): {
        "format": ChartAxisTitleFormat;
        "text": string;
        "visible": boolean;
    };
}
/**
 *
 * Represents the chart axis title formatting.
 *
 * [Api set: ExcelApi 1.1]
 */
export declare class ChartAxisTitleFormat extends OfficeExtension.ClientObject {
    /**
     *
     * Represents the font attributes, such as font name, font size, color, etc. of chart axis title object. Read-only.
     *
     * [Api set: ExcelApi 1.1]
     */
    font: ChartFont;
    /**
     * Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
     */
    load(option?: string | string[] | OfficeExtension.LoadOption): ChartAxisTitleFormat;
    toJSON(): {
        "font": ChartFont;
    };
}
/**
 *
 * Represents a collection of all the data labels on a chart point.
 *
 * [Api set: ExcelApi 1.1]
 */
export declare class ChartDataLabels extends OfficeExtension.ClientObject {
    /**
     *
     * Represents the format of chart data labels, which includes fill and font formatting. Read-only.
     *
     * [Api set: ExcelApi 1.1]
     */
    format: ChartDataLabelFormat;
    /**
     *
     * DataLabelPosition value that represents the position of the data label. See ChartDataLabelPosition for details.
     *
     * [Api set: ExcelApi 1.1]
     */
    position: string;
    /**
     *
     * String representing the separator used for the data labels on a chart.
     *
     * [Api set: ExcelApi 1.1]
     */
    separator: string;
    /**
     *
     * Boolean value representing if the data label bubble size is visible or not.
     *
     * [Api set: ExcelApi 1.1]
     */
    showBubbleSize: boolean;
    /**
     *
     * Boolean value representing if the data label category name is visible or not.
     *
     * [Api set: ExcelApi 1.1]
     */
    showCategoryName: boolean;
    /**
     *
     * Boolean value representing if the data label legend key is visible or not.
     *
     * [Api set: ExcelApi 1.1]
     */
    showLegendKey: boolean;
    /**
     *
     * Boolean value representing if the data label percentage is visible or not.
     *
     * [Api set: ExcelApi 1.1]
     */
    showPercentage: boolean;
    /**
     *
     * Boolean value representing if the data label series name is visible or not.
     *
     * [Api set: ExcelApi 1.1]
     */
    showSeriesName: boolean;
    /**
     *
     * Boolean value representing if the data label value is visible or not.
     *
     * [Api set: ExcelApi 1.1]
     */
    showValue: boolean;
    /**
     * Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
     */
    load(option?: string | string[] | OfficeExtension.LoadOption): ChartDataLabels;
    toJSON(): {
        "format": ChartDataLabelFormat;
        "position": string;
        "separator": string;
        "showBubbleSize": boolean;
        "showCategoryName": boolean;
        "showLegendKey": boolean;
        "showPercentage": boolean;
        "showSeriesName": boolean;
        "showValue": boolean;
    };
}
/**
 *
 * Encapsulates the format properties for the chart data labels.
 *
 * [Api set: ExcelApi 1.1]
 */
export declare class ChartDataLabelFormat extends OfficeExtension.ClientObject {
    /**
     *
     * Represents the fill format of the current chart data label. Read-only.
     *
     * [Api set: ExcelApi 1.1]
     */
    fill: ChartFill;
    /**
     *
     * Represents the font attributes (font name, font size, color, etc.) for a chart data label. Read-only.
     *
     * [Api set: ExcelApi 1.1]
     */
    font: ChartFont;
    /**
     * Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
     */
    load(option?: string | string[] | OfficeExtension.LoadOption): ChartDataLabelFormat;
    toJSON(): {
        "fill": ChartFill;
        "font": ChartFont;
    };
}
/**
 *
 * Represents major or minor gridlines on a chart axis.
 *
 * [Api set: ExcelApi 1.1]
 */
export declare class ChartGridlines extends OfficeExtension.ClientObject {
    /**
     *
     * Represents the formatting of chart gridlines. Read-only.
     *
     * [Api set: ExcelApi 1.1]
     */
    format: ChartGridlinesFormat;
    /**
     *
     * Boolean value representing if the axis gridlines are visible or not.
     *
     * [Api set: ExcelApi 1.1]
     */
    visible: boolean;
    /**
     * Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
     */
    load(option?: string | string[] | OfficeExtension.LoadOption): ChartGridlines;
    toJSON(): {
        "format": ChartGridlinesFormat;
        "visible": boolean;
    };
}
/**
 *
 * Encapsulates the format properties for chart gridlines.
 *
 * [Api set: ExcelApi 1.1]
 */
export declare class ChartGridlinesFormat extends OfficeExtension.ClientObject {
    /**
     *
     * Represents chart line formatting. Read-only.
     *
     * [Api set: ExcelApi 1.1]
     */
    line: ChartLineFormat;
    /**
     * Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
     */
    load(option?: string | string[] | OfficeExtension.LoadOption): ChartGridlinesFormat;
    toJSON(): {
        "line": ChartLineFormat;
    };
}
/**
 *
 * Represents the legend in a chart.
 *
 * [Api set: ExcelApi 1.1]
 */
export declare class ChartLegend extends OfficeExtension.ClientObject {
    /**
     *
     * Represents the formatting of a chart legend, which includes fill and font formatting. Read-only.
     *
     * [Api set: ExcelApi 1.1]
     */
    format: ChartLegendFormat;
    /**
     *
     * Boolean value for whether the chart legend should overlap with the main body of the chart.
     *
     * [Api set: ExcelApi 1.1]
     */
    overlay: boolean;
    /**
     *
     * Represents the position of the legend on the chart. See ChartLegendPosition for details.
     *
     * [Api set: ExcelApi 1.1]
     */
    position: string;
    /**
     *
     * A boolean value the represents the visibility of a ChartLegend object.
     *
     * [Api set: ExcelApi 1.1]
     */
    visible: boolean;
    /**
     * Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
     */
    load(option?: string | string[] | OfficeExtension.LoadOption): ChartLegend;
    toJSON(): {
        "format": ChartLegendFormat;
        "overlay": boolean;
        "position": string;
        "visible": boolean;
    };
}
/**
 *
 * Encapsulates the format properties of a chart legend.
 *
 * [Api set: ExcelApi 1.1]
 */
export declare class ChartLegendFormat extends OfficeExtension.ClientObject {
    /**
     *
     * Represents the fill format of an object, which includes background formating information. Read-only.
     *
     * [Api set: ExcelApi 1.1]
     */
    fill: ChartFill;
    /**
     *
     * Represents the font attributes such as font name, font size, color, etc. of a chart legend. Read-only.
     *
     * [Api set: ExcelApi 1.1]
     */
    font: ChartFont;
    /**
     * Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
     */
    load(option?: string | string[] | OfficeExtension.LoadOption): ChartLegendFormat;
    toJSON(): {
        "fill": ChartFill;
        "font": ChartFont;
    };
}
/**
 *
 * Represents a chart title object of a chart.
 *
 * [Api set: ExcelApi 1.1]
 */
export declare class ChartTitle extends OfficeExtension.ClientObject {
    /**
     *
     * Represents the formatting of a chart title, which includes fill and font formatting. Read-only.
     *
     * [Api set: ExcelApi 1.1]
     */
    format: ChartTitleFormat;
    /**
     *
     * Boolean value representing if the chart title will overlay the chart or not.
     *
     * [Api set: ExcelApi 1.1]
     */
    overlay: boolean;
    /**
     *
     * Represents the title text of a chart.
     *
     * [Api set: ExcelApi 1.1]
     */
    text: string;
    /**
     *
     * A boolean value the represents the visibility of a chart title object.
     *
     * [Api set: ExcelApi 1.1]
     */
    visible: boolean;
    /**
     * Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
     */
    load(option?: string | string[] | OfficeExtension.LoadOption): ChartTitle;
    toJSON(): {
        "format": ChartTitleFormat;
        "overlay": boolean;
        "text": string;
        "visible": boolean;
    };
}
/**
 *
 * Provides access to the office art formatting for chart title.
 *
 * [Api set: ExcelApi 1.1]
 */
export declare class ChartTitleFormat extends OfficeExtension.ClientObject {
    /**
     *
     * Represents the fill format of an object, which includes background formating information. Read-only.
     *
     * [Api set: ExcelApi 1.1]
     */
    fill: ChartFill;
    /**
     *
     * Represents the font attributes (font name, font size, color, etc.) for an object. Read-only.
     *
     * [Api set: ExcelApi 1.1]
     */
    font: ChartFont;
    /**
     * Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
     */
    load(option?: string | string[] | OfficeExtension.LoadOption): ChartTitleFormat;
    toJSON(): {
        "fill": ChartFill;
        "font": ChartFont;
    };
}
/**
 *
 * Represents the fill formatting for a chart element.
 *
 * [Api set: ExcelApi 1.1]
 */
export declare class ChartFill extends OfficeExtension.ClientObject {
    /**
     * Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
     */
    load(option?: string | string[] | OfficeExtension.LoadOption): ChartFill;
    /**
     *
     * Clear the fill color of a chart element.
     *
     * [Api set: ExcelApi 1.1]
     */
    clear(): void;
    /**
     *
     * Sets the fill formatting of a chart element to a uniform color.
     *
     * [Api set: ExcelApi 1.1]
     *
     * @param color HTML color code representing the color of the border line, of the form #RRGGBB (e.g. "FFA500") or as a named HTML color (e.g. "orange").
     */
    setSolidColor(color: string): void;
    toJSON(): {};
}
/**
 *
 * Enapsulates the formatting options for line elements.
 *
 * [Api set: ExcelApi 1.1]
 */
export declare class ChartLineFormat extends OfficeExtension.ClientObject {
    /**
     *
     * HTML color code representing the color of lines in the chart.
     *
     * [Api set: ExcelApi 1.1]
     */
    color: string;
    /**
     *
     * Clear the line format of a chart element.
     *
     * [Api set: ExcelApi 1.1]
     */
    clear(): void;
    /**
     * Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
     */
    load(option?: string | string[] | OfficeExtension.LoadOption): ChartLineFormat;
    toJSON(): {
        "color": string;
    };
}
/**
 *
 * This object represents the font attributes (font name, font size, color, etc.) for a chart object.
 *
 * [Api set: ExcelApi 1.1]
 */
export declare class ChartFont extends OfficeExtension.ClientObject {
    /**
     *
     * Represents the bold status of font.
     *
     * [Api set: ExcelApi 1.1]
     */
    bold: boolean;
    /**
     *
     * HTML color code representation of the text color. E.g. #FF0000 represents Red.
     *
     * [Api set: ExcelApi 1.1]
     */
    color: string;
    /**
     *
     * Represents the italic status of the font.
     *
     * [Api set: ExcelApi 1.1]
     */
    italic: boolean;
    /**
     *
     * Font name (e.g. "Calibri")
     *
     * [Api set: ExcelApi 1.1]
     */
    name: string;
    /**
     *
     * Size of the font (e.g. 11)
     *
     * [Api set: ExcelApi 1.1]
     */
    size: number;
    /**
     *
     * Type of underline applied to the font. See ChartUnderlineStyle for details.
     *
     * [Api set: ExcelApi 1.1]
     */
    underline: string;
    /**
     * Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
     */
    load(option?: string | string[] | OfficeExtension.LoadOption): ChartFont;
    toJSON(): {
        "bold": boolean;
        "color": string;
        "italic": boolean;
        "name": string;
        "size": number;
        "underline": string;
    };
}
/**
 *
 * Manages sorting operations on Range objects.
 *
 * [Api set: ExcelApi 1.2]
 */
export declare class RangeSort extends OfficeExtension.ClientObject {
    /**
     *
     * Perform a sort operation.
     *
     * [Api set: ExcelApi 1.2]
     *
     * @param fields The list of conditions to sort on.
     * @param matchCase Whether to have the casing impact string ordering.
     * @param hasHeaders Whether the range has a header.
     * @param orientation Whether the operation is sorting rows or columns.
     * @param method The ordering method used for Chinese characters.
     */
    apply(fields: Array<SortField>, matchCase?: boolean, hasHeaders?: boolean, orientation?: string, method?: string): void;
    toJSON(): {};
}
/**
 *
 * Manages sorting operations on Table objects.
 *
 * [Api set: ExcelApi 1.2]
 */
export declare class TableSort extends OfficeExtension.ClientObject {
    /**
     *
     * Represents the current conditions used to last sort the table.
     *
     * [Api set: ExcelApi 1.2]
     */
    fields: Array<SortField>;
    /**
     *
     * Represents whether the casing impacted the last sort of the table.
     *
     * [Api set: ExcelApi 1.2]
     */
    matchCase: boolean;
    /**
     *
     * Represents Chinese character ordering method last used to sort the table.
     *
     * [Api set: ExcelApi 1.2]
     */
    method: string;
    /**
     *
     * Perform a sort operation.
     *
     * [Api set: ExcelApi 1.2]
     *
     * @param fields The list of conditions to sort on.
     * @param matchCase Whether to have the casing impact string ordering.
     * @param method The ordering method used for Chinese characters.
     */
    apply(fields: Array<SortField>, matchCase?: boolean, method?: string): void;
    /**
     *
     * Clears the sorting that is currently on the table. While this doesn't modify the table's ordering, it clears the state of the header buttons.
     *
     * [Api set: ExcelApi 1.2]
     */
    clear(): void;
    /**
     *
     * Reapplies the current sorting parameters to the table.
     *
     * [Api set: ExcelApi 1.2]
     */
    reapply(): void;
    /**
     * Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
     */
    load(option?: string | string[] | OfficeExtension.LoadOption): TableSort;
    toJSON(): {
        "fields": SortField[];
        "matchCase": boolean;
        "method": string;
    };
}
/**
 *
 * Represents a condition in a sorting operation.
 *
 * [Api set: ExcelApi 1.2]
 */
export interface SortField {
    /**
     *
     * Represents whether the sorting is done in an ascending fashion.
     *
     * [Api set: ExcelApi 1.2]
     */
    ascending?: boolean;
    /**
     *
     * Represents the color that is the target of the condition if the sorting is on font or cell color.
     *
     * [Api set: ExcelApi 1.2]
     */
    color?: string;
    /**
     *
     * Represents additional sorting options for this field.
     *
     * [Api set: ExcelApi 1.2]
     */
    dataOption?: string;
    /**
     *
     * Represents the icon that is the target of the condition if the sorting is on the cell's icon.
     *
     * [Api set: ExcelApi 1.2]
     */
    icon?: Icon;
    /**
     *
     * Represents the column (or row, depending on the sort orientation) that the condition is on. Represented as an offset from the first column (or row).
     *
     * [Api set: ExcelApi 1.2]
     */
    key: number;
    /**
     *
     * Represents the type of sorting of this condition.
     *
     * [Api set: ExcelApi 1.2]
     */
    sortOn?: string;
}
/**
 *
 * Manages the filtering of a table's column.
 *
 * [Api set: ExcelApi 1.2]
 */
export declare class Filter extends OfficeExtension.ClientObject {
    /**
     *
     * The currently applied filter on the given column.
     *
     * [Api set: ExcelApi 1.2]
     */
    criteria: FilterCriteria;
    /**
     *
     * Apply the given filter criteria on the given column.
     *
     * [Api set: ExcelApi 1.2]
     *
     * @param criteria The criteria to apply.
     */
    apply(criteria: FilterCriteria): void;
    /**
     *
     * Apply a "Bottom Item" filter to the column for the given number of elements.
     *
     * [Api set: ExcelApi 1.2]
     *
     * @param count The number of elements from the bottom to show.
     */
    applyBottomItemsFilter(count: number): void;
    /**
     *
     * Apply a "Bottom Percent" filter to the column for the given percentage of elements.
     *
     * [Api set: ExcelApi 1.2]
     *
     * @param percent The percentage of elements from the bottom to show.
     */
    applyBottomPercentFilter(percent: number): void;
    /**
     *
     * Apply a "Cell Color" filter to the column for the given color.
     *
     * [Api set: ExcelApi 1.2]
     *
     * @param color The background color of the cells to show.
     */
    applyCellColorFilter(color: string): void;
    /**
     *
     * Apply a "Icon" filter to the column for the given criteria strings.
     *
     * [Api set: ExcelApi 1.2]
     *
     * @param criteria1 The first criteria string.
     * @param criteria2 The second criteria string.
     * @param oper The operator that describes how the two criteria are joined.
     */
    applyCustomFilter(criteria1: string, criteria2?: string, oper?: string): void;
    /**
     *
     * Apply a "Dynamic" filter to the column.
     *
     * [Api set: ExcelApi 1.2]
     *
     * @param criteria The dynamic criteria to apply.
     */
    applyDynamicFilter(criteria: string): void;
    /**
     *
     * Apply a "Font Color" filter to the column for the given color.
     *
     * [Api set: ExcelApi 1.2]
     *
     * @param color The font color of the cells to show.
     */
    applyFontColorFilter(color: string): void;
    /**
     *
     * Apply a "Icon" filter to the column for the given icon.
     *
     * [Api set: ExcelApi 1.2]
     *
     * @param icon The icons of the cells to show.
     */
    applyIconFilter(icon: Icon): void;
    /**
     *
     * Apply a "Top Item" filter to the column for the given number of elements.
     *
     * [Api set: ExcelApi 1.2]
     *
     * @param count The number of elements from the top to show.
     */
    applyTopItemsFilter(count: number): void;
    /**
     *
     * Apply a "Top Percent" filter to the column for the given percentage of elements.
     *
     * [Api set: ExcelApi 1.2]
     *
     * @param percent The percentage of elements from the top to show.
     */
    applyTopPercentFilter(percent: number): void;
    /**
     *
     * Apply a "Values" filter to the column for the given values.
     *
     * [Api set: ExcelApi 1.2]
     *
     * @param values The list of values to show.
     */
    applyValuesFilter(values: Array<string | FilterDatetime>): void;
    /**
     *
     * Clear the filter on the given column.
     *
     * [Api set: ExcelApi 1.2]
     */
    clear(): void;
    /**
     * Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
     */
    load(option?: string | string[] | OfficeExtension.LoadOption): Filter;
    toJSON(): {
        "criteria": FilterCriteria;
    };
}
/**
 *
 * Represents the filtering criteria applied to a column.
 *
 * [Api set: ExcelApi 1.2]
 */
export interface FilterCriteria {
    /**
     *
     * The HTML color string used to filter cells. Used with "cellColor" and "fontColor" filtering.
     *
     * [Api set: ExcelApi 1.2]
     */
    color?: string;
    /**
     *
     * The first criterion used to filter data. Used as an operator in the case of "custom" filtering.
         For example ">50" for number greater than 50 or "=*s" for values ending in "s".
        
         Used as a number in the case of top/bottom items/percents. E.g. "5" for the top 5 items if filterOn is set to "topItems"
     *
     * [Api set: ExcelApi 1.2]
     */
    criterion1?: string;
    /**
     *
     * The second criterion used to filter data. Only used as an operator in the case of "custom" filtering.
     *
     * [Api set: ExcelApi 1.2]
     */
    criterion2?: string;
    /**
     *
     * The dynamic criteria from the DynamicFilterCriteria set to apply on this column. Used with "dynamic" filtering.
     *
     * [Api set: ExcelApi 1.2]
     */
    dynamicCriteria?: string;
    /**
     *
     * The property used by the filter to determine whether the values should stay visible.
     *
     * [Api set: ExcelApi 1.2]
     */
    filterOn: string;
    /**
     *
     * The icon used to filter cells. Used with "icon" filtering.
     *
     * [Api set: ExcelApi 1.2]
     */
    icon?: Icon;
    /**
     *
     * The operator used to combine criterion 1 and 2 when using "custom" filtering.
     *
     * [Api set: ExcelApi 1.2]
     */
    operator?: string;
    /**
     *
     * The set of values to be used as part of "values" filtering.
     *
     * [Api set: ExcelApi 1.2]
     */
    values?: Array<string | FilterDatetime>;
}
/**
 *
 * Represents how to filter a date when filtering on values.
 *
 * [Api set: ExcelApi 1.2]
 */
export interface FilterDatetime {
    /**
     *
     * The date in ISO8601 format used to filter data.
     *
     * [Api set: ExcelApi 1.2]
     */
    date: string;
    /**
     *
     * How specific the date should be used to keep data. For example, if the date is 2005-04-02 and the specifity is set to "month", the filter operation will keep all rows with a date in the month of april 2009.
     *
     * [Api set: ExcelApi 1.2]
     */
    specificity: string;
}
/**
 *
 * Represents a cell icon.
 *
 * [Api set: ExcelApi 1.2]
 */
export interface Icon {
    /**
     *
     * Represents the index of the icon in the given set.
     *
     * [Api set: ExcelApi 1.2]
     */
    index: number;
    /**
     *
     * Represents the set that the icon is part of.
     *
     * [Api set: ExcelApi 1.2]
     */
    set: string;
}
/**
 *
 * A scoped collection of custom XML parts.
        A scoped collection is the result of some operation, e.g. filtering by namespace.
        A scoped collection cannot be scoped any further.
 *
 * [Api set: ExcelApi 1.4]
 */
export declare class CustomXmlPartScopedCollection extends OfficeExtension.ClientObject {
    /** Gets the loaded child items in this collection. */
    items: Array<CustomXmlPart>;
    /**
     *
     * Gets the number of items in the collection.
     *
     * [Api set: ExcelApi 1.4]
     */
    getCount(): OfficeExtension.ClientResult<number>;
    /**
     *
     * Gets a custom XML part based on its ID.
     *
     * [Api set: ExcelApi 1.4]
     *
     * @param id ID of the object to be retrieved.
     */
    getItem(id: string): CustomXmlPart;
    /**
     *
     * Gets a custom XML part based on its ID.
        If the CustomXmlPart does not exist, the return object's isNull property will be true.
     *
     * [Api set: ExcelApi 1.4]
     *
     * @param id ID of the object to be retrieved.
     */
    getItemOrNullObject(id: string): CustomXmlPart;
    /**
     *
     * If the collection contains exactly one item, this method returns it.
        Otherwise, this method produces an error.
     *
     * [Api set: ExcelApi 1.4]
     */
    getOnlyItem(): CustomXmlPart;
    /**
     *
     * If the collection contains exactly one item, this method returns it.
        Otherwise, this method returns Null.
     *
     * [Api set: ExcelApi 1.4]
     */
    getOnlyItemOrNullObject(): CustomXmlPart;
    /**
     * Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
     */
    load(option?: string | string[] | OfficeExtension.LoadOption): CustomXmlPartScopedCollection;
    toJSON(): {};
}
/**
 *
 * A collection of custom XML parts.
 *
 * [Api set: ExcelApi 1.4]
 */
export declare class CustomXmlPartCollection extends OfficeExtension.ClientObject {
    /** Gets the loaded child items in this collection. */
    items: Array<CustomXmlPart>;
    /**
     *
     * Adds a new custom XML part to the workbook.
     *
     * [Api set: ExcelApi 1.4]
     *
     * @param xml XML content. Must be a valid XML fragment.
     */
    add(xml: string): CustomXmlPart;
    /**
     *
     * Gets a new scoped collection of custom XML parts whose namespaces match the given namespace.
     *
     * [Api set: ExcelApi 1.4]
     *
     * @param namespaceUri
     */
    getByNamespace(namespaceUri: string): CustomXmlPartScopedCollection;
    /**
     *
     * Gets the number of items in the collection.
     *
     * [Api set: ExcelApi 1.4]
     */
    getCount(): OfficeExtension.ClientResult<number>;
    /**
     *
     * Gets a custom XML part based on its ID.
     *
     * [Api set: ExcelApi 1.4]
     *
     * @param id ID of the object to be retrieved.
     */
    getItem(id: string): CustomXmlPart;
    /**
     *
     * Gets a custom XML part based on its ID.
        If the CustomXmlPart does not exist, the return object's isNull property will be true.
     *
     * [Api set: ExcelApi 1.4]
     *
     * @param id ID of the object to be retrieved.
     */
    getItemOrNullObject(id: string): CustomXmlPart;
    /**
     * Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
     */
    load(option?: string | string[] | OfficeExtension.LoadOption): CustomXmlPartCollection;
    toJSON(): {};
}
/**
 *
 * Represents a custom XML part object in a workbook.
 *
 * [Api set: ExcelApi 1.4]
 */
export declare class CustomXmlPart extends OfficeExtension.ClientObject {
    /**
     *
     * The custom XML part's ID. Read-only.
     *
     * [Api set: ExcelApi 1.4]
     */
    id: string;
    /**
     *
     * The custom XML part's namespace URI. Read-only.
     *
     * [Api set: ExcelApi 1.4]
     */
    namespaceUri: string;
    /**
     *
     * Deletes the custom XML part.
     *
     * [Api set: ExcelApi 1.4]
     */
    delete(): void;
    /**
     *
     * Gets the custom XML part's full XML content.
     *
     * [Api set: ExcelApi 1.4]
     */
    getXml(): OfficeExtension.ClientResult<string>;
    /**
     *
     * Sets the custom XML part's full XML content.
     *
     * [Api set: ExcelApi 1.4]
     */
    setXml(xml: string): void;
    /**
     * Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
     */
    load(option?: string | string[] | OfficeExtension.LoadOption): CustomXmlPart;
    toJSON(): {
        "id": string;
        "namespaceUri": string;
    };
}
/**
 * [Api set: ExcelApi 1.3]
 */
export declare class _V1Api extends OfficeExtension.ClientObject {
    bindingAddColumns(input: any): OfficeExtension.ClientResult<any>;
    bindingAddFromNamedItem(input: any): OfficeExtension.ClientResult<any>;
    bindingAddFromPrompt(input: any): OfficeExtension.ClientResult<any>;
    bindingAddFromSelection(input: any): OfficeExtension.ClientResult<any>;
    bindingAddRows(input: any): OfficeExtension.ClientResult<any>;
    bindingClearFormats(input: any): OfficeExtension.ClientResult<any>;
    bindingDeleteAllDataValues(input: any): OfficeExtension.ClientResult<any>;
    bindingGetAll(): OfficeExtension.ClientResult<any>;
    bindingGetById(input: any): OfficeExtension.ClientResult<any>;
    bindingGetData(input: any): OfficeExtension.ClientResult<any>;
    bindingReleaseById(input: any): OfficeExtension.ClientResult<any>;
    bindingSetData(input: any): OfficeExtension.ClientResult<any>;
    bindingSetFormats(input: any): OfficeExtension.ClientResult<any>;
    bindingSetTableOptions(input: any): OfficeExtension.ClientResult<any>;
    getSelectedData(input: any): OfficeExtension.ClientResult<any>;
    gotoById(input: any): OfficeExtension.ClientResult<any>;
    setSelectedData(input: any): OfficeExtension.ClientResult<any>;
    toJSON(): {};
}
/**
 *
 * Represents a collection of all the PivotTables that are part of the workbook or worksheet.
 *
 * [Api set: ExcelApi 1.3]
 */
export declare class PivotTableCollection extends OfficeExtension.ClientObject {
    /** Gets the loaded child items in this collection. */
    items: Array<PivotTable>;
    /**
     *
     * Gets a PivotTable by name.
     *
     * [Api set: ExcelApi 1.3]
     *
     * @param name Name of the PivotTable to be retrieved.
     */
    getItem(name: string): PivotTable;
    /**
     *
     * Gets a PivotTable by name. If the PivotTable does not exist, will return a null object.
     *
     * [Api set: ExcelApi 1.4]
     *
     * @param name Name of the PivotTable to be retrieved.
     */
    getItemOrNullObject(name: string): PivotTable;
    /**
     *
     * Refreshes all the PivotTables in the collection.
     *
     * [Api set: ExcelApi 1.3]
     */
    refreshAll(): void;
    /**
     * Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
     */
    load(option?: string | string[] | OfficeExtension.LoadOption): PivotTableCollection;
    toJSON(): {};
}
/**
 *
 * Represents an Excel PivotTable.
 *
 * [Api set: ExcelApi 1.3]
 */
export declare class PivotTable extends OfficeExtension.ClientObject {
    /**
     *
     * The worksheet containing the current PivotTable. Read-only.
     *
     * [Api set: ExcelApi 1.3]
     */
    worksheet: Worksheet;
    /**
     *
     * Name of the PivotTable.
     *
     * [Api set: ExcelApi 1.3]
     */
    name: string;
    /**
     *
     * Refreshes the PivotTable.
     *
     * [Api set: ExcelApi 1.3]
     */
    refresh(): void;
    /**
     * Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
     */
    load(option?: string | string[] | OfficeExtension.LoadOption): PivotTable;
    toJSON(): {
        "name": string;
    };
}
/**
 *
 * Represents a collection of all the conditional formats that are overlap the range.
 *
 * [Api set: ExcelApi 1.5]
 */
export declare class ConditionalFormatCollection extends OfficeExtension.ClientObject {
    /** Gets the loaded child items in this collection. */
    items: Array<ConditionalFormat>;
    /**
     *
     * Adds a new conditional format to the collection at the first/top priority.
     *
     * [Api set: ExcelApi 1.5]
     *
     * @param type The type of conditional format being added. See ConditionalFormatType for details.
     */
    add(type: string): ConditionalFormat;
    /**
     *
     * Clears all conditional formats active on the current specified range.
     *
     * [Api set: ExcelApi 1.5]
     */
    clearAll(): void;
    /**
     *
     * Returns the number of conditional formats in the workbook. Read-only.
     *
     * [Api set: ExcelApi 1.5]
     */
    getCount(): OfficeExtension.ClientResult<number>;
    /**
     *
     * Returns a conditional format at the given index.
     *
     * [Api set: ExcelApi 1.5]
     *
     * @param index Index of the conditional formats to be retrieved.
     */
    getItemAt(index: number): ConditionalFormat;
    /**
     * Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
     */
    load(option?: string | string[] | OfficeExtension.LoadOption): ConditionalFormatCollection;
    toJSON(): {};
}
/**
 *
 * An object encapsulating a conditional format's range, format, rule, and other properties.
 *
 * [Api set: ExcelApi 1.5]
 */
export declare class ConditionalFormat extends OfficeExtension.ClientObject {
    /**
     *
     * Represents a Color Scale conditional format.
     *
     * [Api set: ExcelApi 1.4]
     */
    colorScale: ColorScaleConditionalFormat;
    /**
     *
     * Represents a Color Scale conditional format.
     *
     * [Api set: ExcelApi 1.4]
     */
    colorScaleOrNullObject: ColorScaleConditionalFormat;
    /**
     *
     * A custom conditional format and rule.
     *
     * [Api set: ExcelApi 1.5]
     */
    custom: CustomConditionalFormat;
    /**
     *
     * A custom conditional format and rule.
     *
     * [Api set: ExcelApi 1.5]
     */
    customOrNullObject: CustomConditionalFormat;
    /**
     *
     * Represents databars with customizable color, gradient, axis, and range format options.
        If no properties are set, a databar is created with the automatic default settings.
     *
     * [Api set: ExcelApi 1.5]
     */
    dataBar: DataBarConditionalFormat;
    /**
     *
     * Represents databars with customizable color, gradient, axis, and range format options.
        If no properties are set, a databar is created with the automatic default settings.
     *
     * [Api set: ExcelApi 1.5]
     */
    dataBarOrNullObject: DataBarConditionalFormat;
    /**
     *
     * Represents an IconSet conditional format.
     *
     * [Api set: ExcelApi 1.5]
     */
    iconSet: IconSetConditionalFormat;
    /**
     *
     * Represents an IconSet conditional format.
     *
     * [Api set: ExcelApi 1.5]
     */
    iconSetOrNullObject: IconSetConditionalFormat;
    /**
     *
     * The priority (or index) within the conditional format collection that this conditional format currently exists in. Changing this also
        changes other conditional formats' priorities, to allow for a contiguous priority order.
        Use a negative priority to begin from the back.
        Priorities greater than than bounds will get and set to the maximum (or minimum if negative) priority.
     *
     * [Api set: ExcelApi 1.5]
     */
    priority: number;
    /**
     *
     * If the conditions of this conditional format are met, no lower-priority formats shall take effect on that cell.
        Null on databars, icon sets, and colorscales as there's no concept of StopIfTrue for these
     *
     * [Api set: ExcelApi 1.5]
     */
    stopIfTrue: boolean;
    /**
     *
     * A type of conditional format. Only one can be set at a time. Read-Only.
     *
     * [Api set: ExcelApi 1.5]
     */
    type: string;
    /**
     *
     * Deletes this conditional format.
     *
     * [Api set: ExcelApi 1.5]
     */
    delete(): void;
    /**
     *
     * Returns the range the conditonal format is applied to or a null object if the range is discontiguous. Read-only.
     *
     * [Api set: ExcelApi 1.5]
     */
    getRange(): Range;
    /**
     *
     * Returns the range the conditonal format is applied to or a null object if the range is discontiguous. Read-only.
     *
     * [Api set: ExcelApi 1.5]
     */
    getRangeOrNullObject(): Range;
    /**
     * Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
     */
    load(option?: string | string[] | OfficeExtension.LoadOption): ConditionalFormat;
    toJSON(): {
        "colorScale": ColorScaleConditionalFormat;
        "colorScaleOrNullObject": ColorScaleConditionalFormat;
        "custom": CustomConditionalFormat;
        "customOrNullObject": CustomConditionalFormat;
        "dataBar": DataBarConditionalFormat;
        "dataBarOrNullObject": DataBarConditionalFormat;
        "iconSet": IconSetConditionalFormat;
        "iconSetOrNullObject": IconSetConditionalFormat;
        "priority": number;
        "stopIfTrue": boolean;
        "type": string;
    };
}
/**
 *
 * Represents an Excel Conditional Data Bar Type.
 *
 * [Api set: ExcelApi 1.5]
 */
export declare class DataBarConditionalFormat extends OfficeExtension.ClientObject {
    /**
     *
     * Representation of all values to the left of the axis in an Excel data bar.
     *
     * [Api set: ExcelApi 1.5]
     */
    negativeFormat: ConditionalDataBarNegativeFormat;
    /**
     *
     * Representation of all values to the right of the axis in an Excel data bar.
     *
     * [Api set: ExcelApi 1.5]
     */
    positiveFormat: ConditionalDataBarPositiveFormat;
    /**
     *
     * HTML color code representing the color of the Axis line, of the form #RRGGBB (e.g. "FFA500") or as a named HTML color (e.g. "orange").
        "" (empty string) if no axis is present or set.
     *
     * [Api set: ExcelApi 1.5]
     */
    axisColor: string;
    /**
     *
     * Representation of how the axis is determined for an Excel data bar.
     *
     * [Api set: ExcelApi 1.5]
     */
    axisFormat: string;
    /**
     *
     * Represents the direction that the data bar graphic should be based on.
     *
     * [Api set: ExcelApi 1.5]
     */
    barDirection: string;
    /**
     *
     * The rule for what consistutes the lower bound (and how to calculate it, if applicable) for a data bar.
     *
     * [Api set: ExcelApi 1.5]
     */
    lowerBoundRule: ConditionalDataBarRule;
    /**
     *
     * If true, hides the values from the cells where the data bar is applied.
     *
     * [Api set: ExcelApi 1.5]
     */
    showDataBarOnly: boolean;
    /**
     *
     * The rule for what constitutes the upper bound (and how to calculate it, if applicable) for a data bar.
     *
     * [Api set: ExcelApi 1.5]
     */
    upperBoundRule: ConditionalDataBarRule;
    /**
     * Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
     */
    load(option?: string | string[] | OfficeExtension.LoadOption): DataBarConditionalFormat;
    toJSON(): {
        "axisColor": string;
        "axisFormat": string;
        "barDirection": string;
        "lowerBoundRule": ConditionalDataBarRule;
        "negativeFormat": ConditionalDataBarNegativeFormat;
        "positiveFormat": ConditionalDataBarPositiveFormat;
        "showDataBarOnly": boolean;
        "upperBoundRule": ConditionalDataBarRule;
    };
}
/**
 *
 * Represents a conditional format DataBar Format for the positive side of the data bar.
 *
 * [Api set: ExcelApi 1.5]
 */
export declare class ConditionalDataBarPositiveFormat extends OfficeExtension.ClientObject {
    /**
     *
     * HTML color code representing the color of the border line, of the form #RRGGBB (e.g. "FFA500") or as a named HTML color (e.g. "orange").
        "" (empty string) if no border is present or set.
     *
     * [Api set: ExcelApi 1.5]
     */
    borderColor: string;
    /**
     *
     * HTML color code representing the fill color, of the form #RRGGBB (e.g. "FFA500") or as a named HTML color (e.g. "orange").
     *
     * [Api set: ExcelApi 1.5]
     */
    fillColor: string;
    /**
     *
     * Boolean representation of whether or not the DataBar has a gradient.
     *
     * [Api set: ExcelApi 1.5]
     */
    gradientFill: boolean;
    /**
     * Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
     */
    load(option?: string | string[] | OfficeExtension.LoadOption): ConditionalDataBarPositiveFormat;
    toJSON(): {
        "borderColor": string;
        "fillColor": string;
        "gradientFill": boolean;
    };
}
/**
 *
 * Represents a conditional format DataBar Format for the negative side of the data bar.
 *
 * [Api set: ExcelApi 1.5]
 */
export declare class ConditionalDataBarNegativeFormat extends OfficeExtension.ClientObject {
    /**
     *
     * HTML color code representing the color of the border line, of the form #RRGGBB (e.g. "FFA500") or as a named HTML color (e.g. "orange").
        "Empty String" if no border is present or set.
     *
     * [Api set: ExcelApi 1.5]
     */
    borderColor: string;
    /**
     *
     * HTML color code representing the fill color, of the form #RRGGBB (e.g. "FFA500") or as a named HTML color (e.g. "orange").
     *
     * [Api set: ExcelApi 1.5]
     */
    fillColor: string;
    /**
     *
     * Boolean representation of whether or not the negative DataBar has the same border color as the positive DataBar.
     *
     * [Api set: ExcelApi 1.5]
     */
    matchPositiveBorderColor: boolean;
    /**
     *
     * Boolean representation of whether or not the negative DataBar has the same fill color as the positive DataBar.
     *
     * [Api set: ExcelApi 1.5]
     */
    matchPositiveFillColor: boolean;
    /**
     * Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
     */
    load(option?: string | string[] | OfficeExtension.LoadOption): ConditionalDataBarNegativeFormat;
    toJSON(): {
        "borderColor": string;
        "fillColor": string;
        "matchPositiveBorderColor": boolean;
        "matchPositiveFillColor": boolean;
    };
}
/**
 *
 * Represents a rule-type for a Data Bar.
 *
 * [Api set: ExcelApi 1.5]
 */
export interface ConditionalDataBarRule {
    /**
     *
     * The formula, if required, to evaluate the databar rule on.
     *
     * [Api set: ExcelApi 1.5]
     */
    formula?: string;
    /**
     *
     * The type of rule for the databar.
     *
     * [Api set: ExcelApi 1.5]
     */
    type: string;
}
/**
 *
 * Represents a custom conditional format type.
 *
 * [Api set: ExcelApi 1.5]
 */
export declare class CustomConditionalFormat extends OfficeExtension.ClientObject {
    /**
     *
     * Returns a format object, encapsulating the conditional formats font, fill, borders, and other properties. Read-only.
     *
     * [Api set: ExcelApi 1.5]
     */
    format: ConditionalRangeFormat;
    /**
     *
     * Represents the Rule object on this conditional format.
     *
     * [Api set: ExcelApi 1.5]
     */
    rule: ConditionalFormatRule;
    /**
     * Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
     */
    load(option?: string | string[] | OfficeExtension.LoadOption): CustomConditionalFormat;
    toJSON(): {
        "format": ConditionalRangeFormat;
        "rule": ConditionalFormatRule;
    };
}
/**
 *
 * Represents a rule, for all traditional rule/format pairings.
 *
 * [Api set: ExcelApi 1.5]
 */
export declare class ConditionalFormatRule extends OfficeExtension.ClientObject {
    /**
     *
     * The formula, if required, to evaluate the conditional format rule on.
     *
     * [Api set: ExcelApi 1.5]
     */
    formula1: string;
    /**
     *
     * The formula, if required, to evaluate the conditional format rule on in the user's language.
     *
     * [Api set: ExcelApi 1.5]
     */
    formula1Local: string;
    /**
     *
     * The formula, if required, to evaluate the conditional format rule on in R1C1-style notation.
     *
     * [Api set: ExcelApi 1.5]
     */
    formula1R1C1: string;
    /**
     *
     * The formula, if required, to evaluate the conditional format rule on.
     *
     * [Api set: ExcelApi 1.5]
     */
    formula2: string;
    /**
     *
     * The formula, if required, to evaluate the conditional format rule on in the user's language.
     *
     * [Api set: ExcelApi 1.5]
     */
    formula2Local: string;
    /**
     *
     * The formula, if required, to evaluate the conditional format rule on in R1C1-style notation.
     *
     * [Api set: ExcelApi 1.5]
     */
    formula2R1C1: string;
    /**
     * Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
     */
    load(option?: string | string[] | OfficeExtension.LoadOption): ConditionalFormatRule;
    toJSON(): {
        "formula1": string;
        "formula1Local": string;
        "formula1R1C1": string;
        "formula2": string;
        "formula2Local": string;
        "formula2R1C1": string;
    };
}
/**
 *
 * Represents an IconSet criteria for conditional formatting.
 *
 * [Api set: ExcelApi 1.5]
 */
export declare class IconSetConditionalFormat extends OfficeExtension.ClientObject {
    /**
     *
     * An array of Criteria and IconSets for the rules and potential custom icons for conditional icons. Note that for the first criterion only the custom icon can be modified, while type, formula and operator will be ignored when set.
     *
     * [Api set: ExcelApi 1.5]
     */
    criteria: Array<ConditionalIconCriterion>;
    /**
     *
     * If true, reverses the icon orders for the IconSet. Note that this cannot be set if custom icons are used.
     *
     * [Api set: ExcelApi 1.5]
     */
    reverseIconOrder: boolean;
    /**
     *
     * If true, hides the values and only shows icons.
     *
     * [Api set: ExcelApi 1.5]
     */
    showIconOnly: boolean;
    /**
     *
     * If set, displays the IconSet option for the conditional format.
     *
     * [Api set: ExcelApi 1.5]
     */
    style: string;
    /**
     * Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
     */
    load(option?: string | string[] | OfficeExtension.LoadOption): IconSetConditionalFormat;
    toJSON(): {
        "criteria": ConditionalIconCriterion[];
        "reverseIconOrder": boolean;
        "showIconOnly": boolean;
        "style": string;
    };
}
/**
 *
 * Represents an Icon Criterion which contains a type, value, an Operator, and an optional custom icon, if not using an iconset.
 *
 * [Api set: ExcelApi 1.5]
 */
export interface ConditionalIconCriterion {
    /**
     *
     * The custom icon for the current criterion if different from the default IconSet, else null will be returned.
     *
     * [Api set: ExcelApi 1.5]
     */
    customIcon?: Icon;
    /**
     *
     * A number or a formula depending on the type.
     *
     * [Api set: ExcelApi 1.5]
     */
    formula: string;
    /**
     *
     * GreaterThan or GreaterThanOrEqual for each of the rule type for the Icon conditional format.
     *
     * [Api set: ExcelApi 1.5]
     */
    operator: string;
    /**
     *
     * What the icon conditional formula should be based on.
     *
     * [Api set: ExcelApi 1.5]
     */
    type: string;
}
/**
 *
 * Represents an IconSet criteria for conditional formatting.
 *
 * [Api set: ExcelApi 1.4]
 */
export declare class ColorScaleConditionalFormat extends OfficeExtension.ClientObject {
    /**
     *
     * The criteria of the color scale. Midpoint is optional when using a two point color scale.
     *
     * [Api set: ExcelApi 1.4]
     */
    criteria: ConditionalColorScaleCriteria;
    /**
     *
     * If true the color scale will have three points (minimum, midpoint, maximum), otherwise it will have two (minimum, maximum).
     *
     * [Api set: ExcelApi 1.4]
     */
    threeColorScale: boolean;
    /**
     * Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
     */
    load(option?: string | string[] | OfficeExtension.LoadOption): ColorScaleConditionalFormat;
    toJSON(): {
        "criteria": ConditionalColorScaleCriteria;
        "threeColorScale": boolean;
    };
}
/**
 *
 * Represents the criteria of the color scale.
 *
 * [Api set: ExcelApi 1.4]
 */
export interface ConditionalColorScaleCriteria {
    /**
     *
     * The maximum point Color Scale Criterion.
     *
     * [Api set: ExcelApi 1.4]
     */
    maximum: ConditionalColorScaleCriterion;
    /**
     *
     * The midpoint Color Scale Criterion if the color scale is a 3-color scale.
     *
     * [Api set: ExcelApi 1.4]
     */
    midpoint?: ConditionalColorScaleCriterion;
    /**
     *
     * The minimum point Color Scale Criterion.
     *
     * [Api set: ExcelApi 1.4]
     */
    minimum: ConditionalColorScaleCriterion;
}
/**
 *
 * Represents a Color Scale Criterion which contains a type, value and a color.
 *
 * [Api set: ExcelApi 1.4]
 */
export interface ConditionalColorScaleCriterion {
    /**
     *
     * HTML color code representation of the color scale color. E.g. #FF0000 represents Red.
     *
     * [Api set: ExcelApi 1.4]
     */
    color?: string;
    /**
     *
     * A number, a formula, or null (if Type is LowestValue).
     *
     * [Api set: ExcelApi 1.4]
     */
    formula?: string;
    /**
     *
     * What the icon conditional formula should be based on.
     *
     * [Api set: ExcelApi 1.4]
     */
    type: string;
}
/**
 *
 * A format object encapsulating the conditional formats range's font, fill, borders, and other properties.
 *
 * [Api set: ExcelApi 1.5]
 */
export declare class ConditionalRangeFormat extends OfficeExtension.ClientObject {
    /**
     *
     * Collection of border objects that apply to the overall conditional format range. Read-only.
     *
     * [Api set: ExcelApi 1.5]
     */
    borders: ConditionalRangeBorderCollection;
    /**
     *
     * Returns the fill object defined on the overall conditional format range. Read-only.
     *
     * [Api set: ExcelApi 1.5]
     */
    fill: ConditionalRangeFill;
    /**
     *
     * Returns the font object defined on the overall conditional format range. Read-only.
     *
     * [Api set: ExcelApi 1.5]
     */
    font: ConditionalRangeFont;
    /**
     *
     * Represents Excel's number format code for the given range. Cleared if null is passed in.
     *
     * [Api set: ExcelApi 1.5]
     */
    numberFormat: any;
    /**
     * Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
     */
    load(option?: string | string[] | OfficeExtension.LoadOption): ConditionalRangeFormat;
    toJSON(): {
        "numberFormat": any;
    };
}
/**
 *
 * This object represents the font attributes (font style,, color, etc.) for an object.
 *
 * [Api set: ExcelApi 1.5]
 */
export declare class ConditionalRangeFont extends OfficeExtension.ClientObject {
    /**
     *
     * Represents the bold status of font.
     *
     * [Api set: ExcelApi 1.5]
     */
    bold: boolean;
    /**
     *
     * HTML color code representation of the text color. E.g. #FF0000 represents Red.
     *
     * [Api set: ExcelApi 1.5]
     */
    color: string;
    /**
     *
     * Represents the italic status of the font.
     *
     * [Api set: ExcelApi 1.5]
     */
    italic: boolean;
    /**
     *
     * Represents the strikethrough status of the font.
     *
     * [Api set: ExcelApi 1.5]
     */
    strikethrough: boolean;
    /**
     *
     * Type of underline applied to the font. See ConditionalRangeFontUnderlineStyle for details.
     *
     * [Api set: ExcelApi 1.5]
     */
    underline: string;
    /**
     *
     * Resets the font formats.
     *
     * [Api set: ExcelApi 1.5]
     */
    clear(): void;
    /**
     * Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
     */
    load(option?: string | string[] | OfficeExtension.LoadOption): ConditionalRangeFont;
    toJSON(): {
        "bold": boolean;
        "color": string;
        "italic": boolean;
        "strikethrough": boolean;
        "underline": string;
    };
}
/**
 *
 * Represents the background of a conditional range object.
 *
 * [Api set: ExcelApi 1.5]
 */
export declare class ConditionalRangeFill extends OfficeExtension.ClientObject {
    /**
     *
     * HTML color code representing the color of the fill, of the form #RRGGBB (e.g. "FFA500") or as a named HTML color (e.g. "orange").
     *
     * [Api set: ExcelApi 1.5]
     */
    color: string;
    /**
     *
     * Resets the fill.
     *
     * [Api set: ExcelApi 1.5]
     */
    clear(): void;
    /**
     * Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
     */
    load(option?: string | string[] | OfficeExtension.LoadOption): ConditionalRangeFill;
    toJSON(): {
        "color": string;
    };
}
/**
 *
 * Represents the border of an object.
 *
 * [Api set: ExcelApi 1.5]
 */
export declare class ConditionalRangeBorder extends OfficeExtension.ClientObject {
    /**
     *
     * HTML color code representing the color of the border line, of the form #RRGGBB (e.g. "FFA500") or as a named HTML color (e.g. "orange").
     *
     * [Api set: ExcelApi 1.5]
     */
    color: string;
    /**
     *
     * Constant value that indicates the specific side of the border. See ConditionalRangeBorderIndex for details. Read-only.
     *
     * [Api set: ExcelApi 1.5]
     */
    sideIndex: string;
    /**
     *
     * One of the constants of line style specifying the line style for the border. See BorderLineStyle for details.
     *
     * [Api set: ExcelApi 1.5]
     */
    style: string;
    /**
     * Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
     */
    load(option?: string | string[] | OfficeExtension.LoadOption): ConditionalRangeBorder;
    toJSON(): {
        "color": string;
        "sideIndex": string;
        "style": string;
    };
}
/**
 *
 * Represents the border objects that make up range border.
 *
 * [Api set: ExcelApi 1.5]
 */
export declare class ConditionalRangeBorderCollection extends OfficeExtension.ClientObject {
    /**
     *
     * Gets the top border
     *
     * [Api set: ExcelApi 1.5]
     */
    bottom: ConditionalRangeBorder;
    /**
     *
     * Gets the top border
     *
     * [Api set: ExcelApi 1.5]
     */
    left: ConditionalRangeBorder;
    /**
     *
     * Gets the top border
     *
     * [Api set: ExcelApi 1.5]
     */
    right: ConditionalRangeBorder;
    /**
     *
     * Gets the top border
     *
     * [Api set: ExcelApi 1.5]
     */
    top: ConditionalRangeBorder;
    /** Gets the loaded child items in this collection. */
    items: Array<ConditionalRangeBorder>;
    /**
     *
     * Number of border objects in the collection. Read-only.
     *
     * [Api set: ExcelApi 1.5]
     */
    count: number;
    /**
     *
     * Gets a border object using its name
     *
     * [Api set: ExcelApi 1.5]
     *
     * @param index Index value of the border object to be retrieved. See ConditionalRangeBorderIndex for details.
     */
    getItem(index: string): ConditionalRangeBorder;
    /**
     *
     * Gets a border object using its index
     *
     * [Api set: ExcelApi 1.5]
     *
     * @param index Index value of the object to be retrieved. Zero-indexed.
     */
    getItemAt(index: number): ConditionalRangeBorder;
    /**
     * Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
     */
    load(option?: string | string[] | OfficeExtension.LoadOption): ConditionalRangeBorderCollection;
    toJSON(): {
        "count": number;
    };
}
/**
 * [Api set: ExcelApi 1.1]
 */
export declare namespace BindingType {
    var range: string;
    var table: string;
    var text: string;
}
/**
 * [Api set: ExcelApi 1.1]
 */
export declare namespace BorderIndex {
    var edgeTop: string;
    var edgeBottom: string;
    var edgeLeft: string;
    var edgeRight: string;
    var insideVertical: string;
    var insideHorizontal: string;
    var diagonalDown: string;
    var diagonalUp: string;
}
/**
 * [Api set: ExcelApi 1.1]
 */
export declare namespace BorderLineStyle {
    var none: string;
    var continuous: string;
    var dash: string;
    var dashDot: string;
    var dashDotDot: string;
    var dot: string;
    var double: string;
    var slantDashDot: string;
}
/**
 * [Api set: ExcelApi 1.1]
 */
export declare namespace BorderWeight {
    var hairline: string;
    var thin: string;
    var medium: string;
    var thick: string;
}
/**
 * [Api set: ExcelApi 1.1]
 */
export declare namespace CalculationMode {
    var automatic: string;
    var automaticExceptTables: string;
    var manual: string;
}
/**
 * [Api set: ExcelApi 1.1]
 */
export declare namespace CalculationType {
    /**
     *
     * This is a soft recalculation and is mainly for backwards compatibilty. To recalculate all cells use Full or FullRebuild.
     *
     */
    var recalculate: string;
    /**
     *
     * Recalculates all cells that Excel has marked as dirty, that is, dependents of volatile or changed data, and cells programmatically marked as dirty.
     *
     */
    var full: string;
    /**
     *
     * Recalculates all cells in all open workbooks.
     *
     */
    var fullRebuild: string;
}
/**
 * [Api set: ExcelApi 1.1]
 */
export declare namespace ClearApplyTo {
    var all: string;
    var formats: string;
    var contents: string;
}
/**
 * [Api set: ExcelApi 1.1]
 */
export declare namespace ChartDataLabelPosition {
    var invalid: string;
    var none: string;
    var center: string;
    var insideEnd: string;
    var insideBase: string;
    var outsideEnd: string;
    var left: string;
    var right: string;
    var top: string;
    var bottom: string;
    var bestFit: string;
    var callout: string;
}
/**
 * [Api set: ExcelApi 1.1]
 */
export declare namespace ChartLegendPosition {
    var invalid: string;
    var top: string;
    var bottom: string;
    var left: string;
    var right: string;
    var corner: string;
    var custom: string;
}
/**
 *
 * Specifies whether the series are by rows or by columns. On Desktop, the "auto" option will inspect the source data shape to automatically guess whether the data is by rows or columns; on Excel Online, "auto" will simply default to "columns".
 *
 * [Api set: ExcelApi 1.1]
 */
export declare namespace ChartSeriesBy {
    /**
     *
     * On Desktop, the "auto" option will inspect the source data shape to automatically guess whether the data is by rows or columns; on Excel Online, "auto" will simply default to "columns".
     *
     */
    var auto: string;
    var columns: string;
    var rows: string;
}
/**
 * [Api set: ExcelApi 1.1]
 */
export declare namespace ChartType {
    var invalid: string;
    var columnClustered: string;
    var columnStacked: string;
    var columnStacked100: string;
    var _3DColumnClustered: string;
    var _3DColumnStacked: string;
    var _3DColumnStacked100: string;
    var barClustered: string;
    var barStacked: string;
    var barStacked100: string;
    var _3DBarClustered: string;
    var _3DBarStacked: string;
    var _3DBarStacked100: string;
    var lineStacked: string;
    var lineStacked100: string;
    var lineMarkers: string;
    var lineMarkersStacked: string;
    var lineMarkersStacked100: string;
    var pieOfPie: string;
    var pieExploded: string;
    var _3DPieExploded: string;
    var barOfPie: string;
    var xyscatterSmooth: string;
    var xyscatterSmoothNoMarkers: string;
    var xyscatterLines: string;
    var xyscatterLinesNoMarkers: string;
    var areaStacked: string;
    var areaStacked100: string;
    var _3DAreaStacked: string;
    var _3DAreaStacked100: string;
    var doughnutExploded: string;
    var radarMarkers: string;
    var radarFilled: string;
    var surface: string;
    var surfaceWireframe: string;
    var surfaceTopView: string;
    var surfaceTopViewWireframe: string;
    var bubble: string;
    var bubble3DEffect: string;
    var stockHLC: string;
    var stockOHLC: string;
    var stockVHLC: string;
    var stockVOHLC: string;
    var cylinderColClustered: string;
    var cylinderColStacked: string;
    var cylinderColStacked100: string;
    var cylinderBarClustered: string;
    var cylinderBarStacked: string;
    var cylinderBarStacked100: string;
    var cylinderCol: string;
    var coneColClustered: string;
    var coneColStacked: string;
    var coneColStacked100: string;
    var coneBarClustered: string;
    var coneBarStacked: string;
    var coneBarStacked100: string;
    var coneCol: string;
    var pyramidColClustered: string;
    var pyramidColStacked: string;
    var pyramidColStacked100: string;
    var pyramidBarClustered: string;
    var pyramidBarStacked: string;
    var pyramidBarStacked100: string;
    var pyramidCol: string;
    var _3DColumn: string;
    var line: string;
    var _3DLine: string;
    var _3DPie: string;
    var pie: string;
    var xyscatter: string;
    var _3DArea: string;
    var area: string;
    var doughnut: string;
    var radar: string;
}
/**
 * [Api set: ExcelApi 1.1]
 */
export declare namespace ChartUnderlineStyle {
    var none: string;
    var single: string;
}
/**
 *
 * Represents the format options for a Data Bar Axis.
 *
 * [Api set: ExcelApi 1.5]
 */
export declare namespace ConditionalDataBarAxisFormat {
    var automatic: string;
    var none: string;
    var cellMidPoint: string;
}
/**
 *
 * Represents the Data Bar direction within a cell.
 *
 * [Api set: ExcelApi 1.5]
 */
export declare namespace ConditionalDataBarDirection {
    var context: string;
    var leftToRight: string;
    var rightToLeft: string;
}
/**
 *
 * Represents the direction for a selection.
 *
 * [Api set: ExcelApi 1.5]
 */
export declare namespace ConditionalFormatDirection {
    var top: string;
    var bottom: string;
}
/**
 * [Api set: ExcelApi 1.5]
 */
export declare namespace ConditionalFormatType {
    var custom: string;
    var dataBar: string;
    var colorScale: string;
    var iconSet: string;
}
/**
 *
 * Represents the types of conditional format values.
 *
 * [Api set: ExcelApi 1.5]
 */
export declare namespace ConditionalFormatRuleType {
    var invalid: string;
    var automatic: string;
    var lowestValue: string;
    var highestValue: string;
    var number: string;
    var percent: string;
    var formula: string;
    var percentile: string;
}
/**
 *
 * Represents the types of conditional format values.
 *
 * [Api set: ExcelApi 1.5]
 */
export declare module ConditionalFormatIconRuleType {
    var invalid: string;
    var number: string;
    var percent: string;
    var formula: string;
    var percentile: string;
}
/**
 *
 * Represents the types of conditional format values.
 *
 * [Api set: ExcelApi 1.4]
 */
export declare module ConditionalFormatColorCriterionType {
    var invalid: string;
    var lowestValue: string;
    var highestValue: string;
    var number: string;
    var percent: string;
    var formula: string;
    var percentile: string;
}
/**
 *
 * Represents the operator for each icon criteria.
 *
 * [Api set: ExcelApi 1.5]
 */
export declare module ConditionalIconCriterionOperator {
    var invalid: string;
    var greaterThan: string;
    var greaterThanOrEqual: string;
}
/**
 *
 * Represents all of the potential rule types for formats.
 *
 * [Api set: ExcelApi 1.5]
 */
export declare namespace ConditionalRangeFormatRuleType {
    var blank: string;
    var expression: string;
    var between: string;
    var notBetween: string;
    var count: string;
    var percent: string;
    var average: string;
    var unique: string;
    var error: string;
    var textContains: string;
    var dateOccurring: string;
}
/**
 *
 * Represents all of the potential rule types for formats.
 *
 * [Api set: ExcelApi 1.5]
 */
export declare namespace ConditionalFormatCustomRuleType {
    var formula: string;
    var between: string;
    var notBetween: string;
    var count: string;
    var percent: string;
    var average: string;
    var blank: string;
    var unique: string;
    var error: string;
    var textContains: string;
    var dateOccurring: string;
}
/**
 * [Api set: ExcelApi 1.5]
 */
export declare namespace ConditionalRangeBorderIndex {
    var edgeTop: string;
    var edgeBottom: string;
    var edgeLeft: string;
    var edgeRight: string;
}
/**
 * [Api set: ExcelApi 1.5]
 */
export declare namespace ConditionalRangeBorderLineStyle {
    var none: string;
    var continuous: string;
    var dash: string;
    var dashDot: string;
    var dashDotDot: string;
    var dot: string;
}
/**
 * [Api set: ExcelApi 1.5]
 */
export declare namespace ConditionalRangeFontUnderlineStyle {
    var none: string;
    var single: string;
    var double: string;
}
/**
 * [Api set: ExcelApi 1.1]
 */
export declare namespace DeleteShiftDirection {
    var up: string;
    var left: string;
}
/**
 * [Api set: ExcelApi 1.2]
 */
export declare namespace DynamicFilterCriteria {
    var unknown: string;
    var aboveAverage: string;
    var allDatesInPeriodApril: string;
    var allDatesInPeriodAugust: string;
    var allDatesInPeriodDecember: string;
    var allDatesInPeriodFebruray: string;
    var allDatesInPeriodJanuary: string;
    var allDatesInPeriodJuly: string;
    var allDatesInPeriodJune: string;
    var allDatesInPeriodMarch: string;
    var allDatesInPeriodMay: string;
    var allDatesInPeriodNovember: string;
    var allDatesInPeriodOctober: string;
    var allDatesInPeriodQuarter1: string;
    var allDatesInPeriodQuarter2: string;
    var allDatesInPeriodQuarter3: string;
    var allDatesInPeriodQuarter4: string;
    var allDatesInPeriodSeptember: string;
    var belowAverage: string;
    var lastMonth: string;
    var lastQuarter: string;
    var lastWeek: string;
    var lastYear: string;
    var nextMonth: string;
    var nextQuarter: string;
    var nextWeek: string;
    var nextYear: string;
    var thisMonth: string;
    var thisQuarter: string;
    var thisWeek: string;
    var thisYear: string;
    var today: string;
    var tomorrow: string;
    var yearToDate: string;
    var yesterday: string;
}
/**
 * [Api set: ExcelApi 1.2]
 */
export declare namespace FilterDatetimeSpecificity {
    var year: string;
    var month: string;
    var day: string;
    var hour: string;
    var minute: string;
    var second: string;
}
/**
 * [Api set: ExcelApi 1.2]
 */
export declare namespace FilterOn {
    var bottomItems: string;
    var bottomPercent: string;
    var cellColor: string;
    var dynamic: string;
    var fontColor: string;
    var values: string;
    var topItems: string;
    var topPercent: string;
    var icon: string;
    var custom: string;
}
/**
 * [Api set: ExcelApi 1.2]
 */
export declare namespace FilterOperator {
    var and: string;
    var or: string;
}
/**
 * [Api set: ExcelApi 1.1]
 */
export declare namespace HorizontalAlignment {
    var general: string;
    var left: string;
    var center: string;
    var right: string;
    var fill: string;
    var justify: string;
    var centerAcrossSelection: string;
    var distributed: string;
}
/**
 * [Api set: ExcelApi 1.2]
 */
export declare namespace IconSet {
    var invalid: string;
    var threeArrows: string;
    var threeArrowsGray: string;
    var threeFlags: string;
    var threeTrafficLights1: string;
    var threeTrafficLights2: string;
    var threeSigns: string;
    var threeSymbols: string;
    var threeSymbols2: string;
    var fourArrows: string;
    var fourArrowsGray: string;
    var fourRedToBlack: string;
    var fourRating: string;
    var fourTrafficLights: string;
    var fiveArrows: string;
    var fiveArrowsGray: string;
    var fiveRating: string;
    var fiveQuarters: string;
    var threeStars: string;
    var threeTriangles: string;
    var fiveBoxes: string;
}
/**
 * [Api set: ExcelApi 1.2]
 */
export declare namespace ImageFittingMode {
    var fit: string;
    var fitAndCenter: string;
    var fill: string;
}
/**
 * [Api set: ExcelApi 1.1]
 */
export declare namespace InsertShiftDirection {
    var down: string;
    var right: string;
}
/**
 * [Api set: ExcelApi 1.4]
 */
export declare module NamedItemScope {
    var worksheet: string;
    var workbook: string;
}
/**
 * [Api set: ExcelApi 1.1]
 */
export declare namespace NamedItemType {
    var string: string;
    var integer: string;
    var double: string;
    var boolean: string;
    var range: string;
    var error: string;
}
/**
 * [Api set: ExcelApi 1.1]
 */
export declare namespace RangeUnderlineStyle {
    var none: string;
    var single: string;
    var double: string;
    var singleAccountant: string;
    var doubleAccountant: string;
}
/**
 * [Api set: ExcelApi 1.1]
 */
export declare namespace SheetVisibility {
    var visible: string;
    var hidden: string;
    var veryHidden: string;
}
/**
 * [Api set: ExcelApi 1.1]
 */
export declare namespace RangeValueType {
    var unknown: string;
    var empty: string;
    var string: string;
    var integer: string;
    var double: string;
    var boolean: string;
    var error: string;
}
/**
 * [Api set: ExcelApi 1.2]
 */
export declare namespace SortOrientation {
    var rows: string;
    var columns: string;
}
/**
 * [Api set: ExcelApi 1.2]
 */
export declare namespace SortOn {
    var value: string;
    var cellColor: string;
    var fontColor: string;
    var icon: string;
}
/**
 * [Api set: ExcelApi 1.2]
 */
export declare namespace SortDataOption {
    var normal: string;
    var textAsNumber: string;
}
/**
 * [Api set: ExcelApi 1.2]
 */
export declare namespace SortMethod {
    var pinYin: string;
    var strokeCount: string;
}
/**
 * [Api set: ExcelApi 1.1]
 */
export declare namespace VerticalAlignment {
    var top: string;
    var center: string;
    var bottom: string;
    var justify: string;
    var distributed: string;
}
/**
 *
 * An object containing the result of a function-evaluation operation
 *
 * [Api set: ExcelApi 1.2]
 */
export declare class FunctionResult<T> extends OfficeExtension.ClientObject {
    /**
     *
     * Error value (such as "#DIV/0") representing the error. If the error string is not set, then the function succeeded, and its result is written to the Value field. The error is always in the English locale.
     *
     * [Api set: ExcelApi 1.2]
     */
    error: string;
    /**
     *
     * The value of function evaluation. The value field will be populated only if no error has occurred (i.e., the Error property is not set).
     *
     * [Api set: ExcelApi 1.2]
     */
    value: T;
    /**
     * Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
     */
    load(option?: string | string[] | OfficeExtension.LoadOption): FunctionResult<T>;
    toJSON(): {
        "error": string;
        "value": T;
    };
}
/**
 *
 * An object for evaluating Excel functions.
 *
 * [Api set: ExcelApi 1.2]
 */
export declare class Functions extends OfficeExtension.ClientObject {
    /**
     *
     * Returns the absolute value of a number, a number without its sign.
     *
     * [Api set: ExcelApi 1.2]
     *
     * @param number Is the real number for which you want the absolute value.
     */
    abs(number: number | Range | RangeReference | FunctionResult<any>): FunctionResult<number>;
    /**
     *
     * Returns the accrued interest for a security that pays periodic interest.
     *
     * [Api set: ExcelApi 1.2]
     *
     * @param issue Is the security's issue date, expressed as a serial date number.
     * @param firstInterest Is the security's first interest date, expressed as a serial date number.
     * @param settlement Is the security's settlement date, expressed as a serial date number.
     * @param rate Is the security's annual coupon rate.
     * @param par Is the security's par value.
     * @param frequency Is the number of coupon payments per year.
     * @param basis Is the type of day count basis to use.
     * @param calcMethod Is a logical value: to accrued interest from issue date = TRUE or omitted; to calculate from last coupon payment date = FALSE.
     */
    accrInt(issue: number | string | boolean | Range | RangeReference | FunctionResult<any>, firstInterest: number | string | boolean | Range | RangeReference | FunctionResult<any>, settlement: number | string | boolean | Range | RangeReference | FunctionResult<any>, rate: number | string | boolean | Range | RangeReference | FunctionResult<any>, par: number | string | boolean | Range | RangeReference | FunctionResult<any>, frequency: number | string | boolean | Range | RangeReference | FunctionResult<any>, basis?: number | string | boolean | Range | RangeReference | FunctionResult<any>, calcMethod?: number | string | boolean | Range | RangeReference | FunctionResult<any>): FunctionResult<number>;
    /**
     *
     * Returns the accrued interest for a security that pays interest at maturity.
     *
     * [Api set: ExcelApi 1.2]
     *
     * @param issue Is the security's issue date, expressed as a serial date number.
     * @param settlement Is the security's maturity date, expressed as a serial date number.
     * @param rate Is the security's annual coupon rate.
     * @param par Is the security's par value.
     * @param basis Is the type of day count basis to use.
     */
    accrIntM(issue: number | string | boolean | Range | RangeReference | FunctionResult<any>, settlement: number | string | boolean | Range | RangeReference | FunctionResult<any>, rate: number | string | boolean | Range | RangeReference | FunctionResult<any>, par: number | string | boolean | Range | RangeReference | FunctionResult<any>, basis?: number | string | boolean | Range | RangeReference | FunctionResult<any>): FunctionResult<number>;
    /**
     *
     * Returns the arccosine of a number, in radians in the range 0 to Pi. The arccosine is the angle whose cosine is Number.
     *
     * [Api set: ExcelApi 1.2]
     *
     * @param number Is the cosine of the angle you want and must be from -1 to 1.
     */
    acos(number: number | Range | RangeReference | FunctionResult<any>): FunctionResult<number>;
    /**
     *
     * Returns the inverse hyperbolic cosine of a number.
     *
     * [Api set: ExcelApi 1.2]
     *
     * @param number Is any real number equal to or greater than 1.
     */
    acosh(number: number | Range | RangeReference | FunctionResult<any>): FunctionResult<number>;
    /**
     *
     * Returns the arccotangent of a number, in radians in the range 0 to Pi.
     *
     * [Api set: ExcelApi 1.2]
     *
     * @param number Is the cotangent of the angle you want.
     */
    acot(number: number | Range | RangeReference | FunctionResult<any>): FunctionResult<number>;
    /**
     *
     * Returns the inverse hyperbolic cotangent of a number.
     *
     * [Api set: ExcelApi 1.2]
     *
     * @param number Is the hyperbolic cotangent of the angle that you want.
     */
    acoth(number: number | Range | RangeReference | FunctionResult<any>): FunctionResult<number>;
    /**
     *
     * Returns the prorated linear depreciation of an asset for each accounting period.
     *
     * [Api set: ExcelApi 1.2]
     *
     * @param cost Is the cost of the asset.
     * @param datePurchased Is the date the asset is purchased.
     * @param firstPeriod Is the date of the end of the first period.
     * @param salvage Is the salvage value at the end of life of the asset.
     * @param period Is the period.
     * @param rate Is the rate of depreciation.
     * @param basis Year_basis : 0 for year of 360 days, 1 for actual, 3 for year of 365 days.
     */
    amorDegrc(cost: number | string | boolean | Range | RangeReference | FunctionResult<any>, datePurchased: number | string | boolean | Range | RangeReference | FunctionResult<any>, firstPeriod: number | string | boolean | Range | RangeReference | FunctionResult<any>, salvage: number | string | boolean | Range | RangeReference | FunctionResult<any>, period: number | string | boolean | Range | RangeReference | FunctionResult<any>, rate: number | string | boolean | Range | RangeReference | FunctionResult<any>, basis?: number | string | boolean | Range | RangeReference | FunctionResult<any>): FunctionResult<number>;
    /**
     *
     * Returns the prorated linear depreciation of an asset for each accounting period.
     *
     * [Api set: ExcelApi 1.2]
     *
     * @param cost Is the cost of the asset.
     * @param datePurchased Is the date the asset is purchased.
     * @param firstPeriod Is the date of the end of the first period.
     * @param salvage Is the salvage value at the end of life of the asset.
     * @param period Is the period.
     * @param rate Is the rate of depreciation.
     * @param basis Year_basis : 0 for year of 360 days, 1 for actual, 3 for year of 365 days.
     */
    amorLinc(cost: number | string | boolean | Range | RangeReference | FunctionResult<any>, datePurchased: number | string | boolean | Range | RangeReference | FunctionResult<any>, firstPeriod: number | string | boolean | Range | RangeReference | FunctionResult<any>, salvage: number | string | boolean | Range | RangeReference | FunctionResult<any>, period: number | string | boolean | Range | RangeReference | FunctionResult<any>, rate: number | string | boolean | Range | RangeReference | FunctionResult<any>, basis?: number | string | boolean | Range | RangeReference | FunctionResult<any>): FunctionResult<number>;
    /**
     *
     * Checks whether all arguments are TRUE, and returns TRUE if all arguments are TRUE.
     *
     * [Api set: ExcelApi 1.2]
     *
     * @param values List of parameters, whose elements are 1 to 255 conditions you want to test that can be either TRUE or FALSE and can be logical values, arrays, or references.
     */
    and(...values: Array<boolean | Range | RangeReference | FunctionResult<any>>): FunctionResult<boolean>;
    /**
     *
     * Converts a Roman numeral to Arabic.
     *
     * [Api set: ExcelApi 1.2]
     *
     * @param text Is the Roman numeral you want to convert.
     */
    arabic(text: string | Range | RangeReference | FunctionResult<any>): FunctionResult<number>;
    /**
     *
     * Returns the number of areas in a reference. An area is a range of contiguous cells or a single cell.
     *
     * [Api set: ExcelApi 1.2]
     *
     * @param reference Is a reference to a cell or range of cells and can refer to multiple areas.
     */
    areas(reference: Range | RangeReference | FunctionResult<any>): FunctionResult<number>;
    /**
     *
     * Changes full-width (double-byte) characters to half-width (single-byte) characters. Use with double-byte character sets (DBCS).
     *
     * [Api set: ExcelApi 1.2]
     *
     * @param text Is a text, or a reference to a cell containing a text.
     */
    asc(text: string | Range | RangeReference | FunctionResult<any>): FunctionResult<string>;
    /**
     *
     * Returns the arcsine of a number in radians, in the range -Pi/2 to Pi/2.
     *
     * [Api set: ExcelApi 1.2]
     *
     * @param number Is the sine of the angle you want and must be from -1 to 1.
     */
    asin(number: number | Range | RangeReference | FunctionResult<any>): FunctionResult<number>;
    /**
     *
     * Returns the inverse hyperbolic sine of a number.
     *
     * [Api set: ExcelApi 1.2]
     *
     * @param number Is any real number equal to or greater than 1.
     */
    asinh(number: number | Range | RangeReference | FunctionResult<any>): FunctionResult<number>;
    /**
     *
     * Returns the arctangent of a number in radians, in the range -Pi/2 to Pi/2.
     *
     * [Api set: ExcelApi 1.2]
     *
     * @param number Is the tangent of the angle you want.
     */
    atan(number: number | Range | RangeReference | FunctionResult<any>): FunctionResult<number>;
    /**
     *
     * Returns the arctangent of the specified x- and y- coordinates, in radians between -Pi and Pi, excluding -Pi.
     *
     * [Api set: ExcelApi 1.2]
     *
     * @param xNum Is the x-coordinate of the point.
     * @param yNum Is the y-coordinate of the point.
     */
    atan2(xNum: number | Range | RangeReference | FunctionResult<any>, yNum: number | Range | RangeReference | FunctionResult<any>): FunctionResult<number>;
    /**
     *
     * Returns the inverse hyperbolic tangent of a number.
     *
     * [Api set: ExcelApi 1.2]
     *
     * @param number Is any real number between -1 and 1 excluding -1 and 1.
     */
    atanh(number: number | Range | RangeReference | FunctionResult<any>): FunctionResult<number>;
    /**
     *
     * Returns the average of the absolute deviations of data points from their mean. Arguments can be numbers or names, arrays, or references that contain numbers.
     *
     * [Api set: ExcelApi 1.2]
     *
     * @param values List of parameters, whose elements are 1 to 255 arguments for which you want the average of the absolute deviations.
     */
    aveDev(...values: Array<number | Range | RangeReference | FunctionResult<any>>): FunctionResult<number>;
    /**
     *
     * Returns the average (arithmetic mean) of its arguments, which can be numbers or names, arrays, or references that contain numbers.
     *
     * [Api set: ExcelApi 1.2]
     *
     * @param values List of parameters, whose elements are 1 to 255 numeric arguments for which you want the average.
     */
    average(...values: Array<number | Range | RangeReference | FunctionResult<any>>): FunctionResult<number>;
    /**
     *
     * Returns the average (arithmetic mean) of its arguments, evaluating text and FALSE in arguments as 0; TRUE evaluates as 1. Arguments can be numbers, names, arrays, or references.
     *
     * [Api set: ExcelApi 1.2]
     *
     * @param values List of parameters, whose elements are 1 to 255 arguments for which you want the average.
     */
    averageA(...values: Array<number | Range | RangeReference | FunctionResult<any>>): FunctionResult<number>;
    /**
     *
     * Finds average(arithmetic mean) for the cells specified by a given condition or criteria.
     *
     * [Api set: ExcelApi 1.2]
     *
     * @param range Is the range of cells you want evaluated.
     * @param criteria Is the condition or criteria in the form of a number, expression, or text that defines which cells will be used to find the average.
     * @param averageRange Are the actual cells to be used to find the average. If omitted, the cells in range are used.
     */
    averageIf(range: Range | RangeReference | FunctionResult<any>, criteria: number | string | boolean | Range | RangeReference | FunctionResult<any>, averageRange?: Range | RangeReference | FunctionResult<any>): FunctionResult<number>;
    /**
     *
     * Finds average(arithmetic mean) for the cells specified by a given set of conditions or criteria.
     *
     * [Api set: ExcelApi 1.2]
     *
     * @param averageRange Are the actual cells to be used to find the average.
     * @param values List of parameters, where the first element of each pair is the Is the range of cells you want evaluated for the particular condition , and the second element is is the condition or criteria in the form of a number, expression, or text that defines which cells will be used to find the average.
     */
    averageIfs(averageRange: Range | RangeReference | FunctionResult<any>, ...values: Array<Range | RangeReference | FunctionResult<any> | number | string | boolean>): FunctionResult<number>;
    /**
     *
     * Converts a number to text (baht).
     *
     * [Api set: ExcelApi 1.2]
     *
     * @param number Is a number that you want to convert.
     */
    bahtText(number: number | Range | RangeReference | FunctionResult<any>): FunctionResult<string>;
    /**
     *
     * Converts a number into a text representation with the given radix (base).
     *
     * [Api set: ExcelApi 1.2]
     *
     * @param number Is the number that you want to convert.
     * @param radix Is the base Radix that you want to convert the number into.
     * @param minLength Is the minimum length of the returned string.  If omitted leading zeros are not added.
     */
    base(number: number | Range | RangeReference | FunctionResult<any>, radix: number | Range | RangeReference | FunctionResult<any>, minLength?: number | Range | RangeReference | FunctionResult<any>): FunctionResult<string>;
    /**
     *
     * Returns the modified Bessel function In(x).
     *
     * [Api set: ExcelApi 1.2]
     *
     * @param x Is the value at which to evaluate the function.
     * @param n Is the order of the Bessel function.
     */
    besselI(x: number | string | boolean | Range | RangeReference | FunctionResult<any>, n: number | string | boolean | Range | RangeReference | FunctionResult<any>): FunctionResult<number>;
    /**
     *
     * Returns the Bessel function Jn(x).
     *
     * [Api set: ExcelApi 1.2]
     *
     * @param x Is the value at which to evaluate the function.
     * @param n Is the order of the Bessel function.
     */
    besselJ(x: number | string | boolean | Range | RangeReference | FunctionResult<any>, n: number | string | boolean | Range | RangeReference | FunctionResult<any>): FunctionResult<number>;
    /**
     *
     * Returns the modified Bessel function Kn(x).
     *
     * [Api set: ExcelApi 1.2]
     *
     * @param x Is the value at which to evaluate the function.
     * @param n Is the order of the function.
     */
    besselK(x: number | string | boolean | Range | RangeReference | FunctionResult<any>, n: number | string | boolean | Range | RangeReference | FunctionResult<any>): FunctionResult<number>;
    /**
     *
     * Returns the Bessel function Yn(x).
     *
     * [Api set: ExcelApi 1.2]
     *
     * @param x Is the value at which to evaluate the function.
     * @param n Is the order of the function.
     */
    besselY(x: number | string | boolean | Range | RangeReference | FunctionResult<any>, n: number | string | boolean | Range | RangeReference | FunctionResult<any>): FunctionResult<number>;
    /**
     *
     * Returns the beta probability distribution function.
     *
     * [Api set: ExcelApi 1.2]
     *
     * @param x Is the value between A and B at which to evaluate the function.
     * @param alpha Is a parameter to the distribution and must be greater than 0.
     * @param beta Is a parameter to the distribution and must be greater than 0.
     * @param cumulative Is a logical value: for the cumulative distribution function, use TRUE; for the probability density function, use FALSE.
     * @param A Is an optional lower bound to the interval of x. If omitted, A = 0.
     * @param B Is an optional upper bound to the interval of x. If omitted, B = 1.
     */
    beta_Dist(x: number | Range | RangeReference | FunctionResult<any>, alpha: number | Range | RangeReference | FunctionResult<any>, beta: number | Range | RangeReference | FunctionResult<any>, cumulative: boolean | Range | RangeReference | FunctionResult<any>, A?: number | Range | RangeReference | FunctionResult<any>, B?: number | Range | RangeReference | FunctionResult<any>): FunctionResult<number>;
    /**
     *
     * Returns the inverse of the cumulative beta probability density function (BETA.DIST).
     *
     * [Api set: ExcelApi 1.2]
     *
     * @param probability Is a probability associated with the beta distribution.
     * @param alpha Is a parameter to the distribution and must be greater than 0.
     * @param beta Is a parameter to the distribution and must be greater than 0.
     * @param A Is an optional lower bound to the interval of x. If omitted, A = 0.
     * @param B Is an optional upper bound to the interval of x. If omitted, B = 1.
     */
    beta_Inv(probability: number | Range | RangeReference | FunctionResult<any>, alpha: number | Range | RangeReference | FunctionResult<any>, beta: number | Range | RangeReference | FunctionResult<any>, A?: number | Range | RangeReference | FunctionResult<any>, B?: number | Range | RangeReference | FunctionResult<any>): FunctionResult<number>;
    /**
     *
     * Converts a binary number to decimal.
     *
     * [Api set: ExcelApi 1.2]
     *
     * @param number Is the binary number you want to convert.
     */
    bin2Dec(number: number | string | boolean | Range | RangeReference | FunctionResult<any>): FunctionResult<number>;
    /**
     *
     * Converts a binary number to hexadecimal.
     *
     * [Api set: ExcelApi 1.2]
     *
     * @param number Is the binary number you want to convert.
     * @param places Is the number of characters to use.
     */
    bin2Hex(number: number | string | boolean | Range | RangeReference | FunctionResult<any>, places?: number | string | boolean | Range | RangeReference | FunctionResult<any>): FunctionResult<number>;
    /**
     *
     * Converts a binary number to octal.
     *
     * [Api set: ExcelApi 1.2]
     *
     * @param number Is the binary number you want to convert.
     * @param places Is the number of characters to use.
     */
    bin2Oct(number: number | string | boolean | Range | RangeReference | FunctionResult<any>, places?: number | string | boolean | Range | RangeReference | FunctionResult<any>): FunctionResult<number>;
    /**
     *
     * Returns the individual term binomial distribution probability.
     *
     * [Api set: ExcelApi 1.2]
     *
     * @param numberS Is the number of successes in trials.
     * @param trials Is the number of independent trials.
     * @param probabilityS Is the probability of success on each trial.
     * @param cumulative Is a logical value: for the cumulative distribution function, use TRUE; for the probability mass function, use FALSE.
     */
    binom_Dist(numberS: number | Range | RangeReference | FunctionResult<any>, trials: number | Range | RangeReference | FunctionResult<any>, probabilityS: number | Range | RangeReference | FunctionResult<any>, cumulative: boolean | Range | RangeReference | FunctionResult<any>): FunctionResult<number>;
    /**
     *
     * Returns the probability of a trial result using a binomial distribution.
     *
     * [Api set: ExcelApi 1.2]
     *
     * @param trials Is the number of independent trials.
     * @param probabilityS Is the probability of success on each trial.
     * @param numberS Is the number of successes in trials.
     * @param numberS2 If provided this function returns the probability that the number of successful trials shall lie between numberS and numberS2.
     */
    binom_Dist_Range(trials: number | Range | RangeReference | FunctionResult<any>, probabilityS: number | Range | RangeReference | FunctionResult<any>, numberS: number | Range | RangeReference | FunctionResult<any>, numberS2?: number | Range | RangeReference | FunctionResult<any>): FunctionResult<number>;
    /**
     *
     * Returns the smallest value for which the cumulative binomial distribution is greater than or equal to a criterion value.
     *
     * [Api set: ExcelApi 1.2]
     *
     * @param trials Is the number of Bernoulli trials.
     * @param probabilityS Is the probability of success on each trial, a number between 0 and 1 inclusive.
     * @param alpha Is the criterion value, a number between 0 and 1 inclusive.
     */
    binom_Inv(trials: number | Range | RangeReference | FunctionResult<any>, probabilityS: number | Range | RangeReference | FunctionResult<any>, alpha: number | Range | RangeReference | FunctionResult<any>): FunctionResult<number>;
    /**
     *
     * Returns a bitwise 'And' of two numbers.
     *
     * [Api set: ExcelApi 1.2]
     *
     * @param number1 Is the decimal representation of the binary number you want to evaluate.
     * @param number2 Is the decimal representation of the binary number you want to evaluate.
     */
    bitand(number1: number | Range | RangeReference | FunctionResult<any>, number2: number | Range | RangeReference | FunctionResult<any>): FunctionResult<number>;
    /**
     *
     * Returns a number shifted left by shift_amount bits.
     *
     * [Api set: ExcelApi 1.2]
     *
     * @param number Is the decimal representation of the binary number you want to evaluate.
     * @param shiftAmount Is the number of bits that you want to shift Number left by.
     */
    bitlshift(number: number | Range | RangeReference | FunctionResult<any>, shiftAmount: number | Range | RangeReference | FunctionResult<any>): FunctionResult<number>;
    /**
     *
     * Returns a bitwise 'Or' of two numbers.
     *
     * [Api set: ExcelApi 1.2]
     *
     * @param number1 Is the decimal representation of the binary number you want to evaluate.
     * @param number2 Is the decimal representation of the binary number you want to evaluate.
     */
    bitor(number1: number | Range | RangeReference | FunctionResult<any>, number2: number | Range | RangeReference | FunctionResult<any>): FunctionResult<number>;
    /**
     *
     * Returns a number shifted right by shift_amount bits.
     *
     * [Api set: ExcelApi 1.2]
     *
     * @param number Is the decimal representation of the binary number you want to evaluate.
     * @param shiftAmount Is the number of bits that you want to shift Number right by.
     */
    bitrshift(number: number | Range | RangeReference | FunctionResult<any>, shiftAmount: number | Range | RangeReference | FunctionResult<any>): FunctionResult<number>;
    /**
     *
     * Returns a bitwise 'Exclusive Or' of two numbers.
     *
     * [Api set: ExcelApi 1.2]
     *
     * @param number1 Is the decimal representation of the binary number you want to evaluate.
     * @param number2 Is the decimal representation of the binary number you want to evaluate.
     */
    bitxor(number1: number | Range | RangeReference | FunctionResult<any>, number2: number | Range | RangeReference | FunctionResult<any>): FunctionResult<number>;
    /**
     *
     * Rounds a number up, to the nearest integer or to the nearest multiple of significance.
     *
     * [Api set: ExcelApi 1.2]
     *
     * @param number Is the value you want to round.
     * @param significance Is the multiple to which you want to round.
     * @param mode When given and nonzero this function will round away from zero.
     */
    ceiling_Math(number: number | Range | RangeReference | FunctionResult<any>, significance?: number | Range | RangeReference | FunctionResult<any>, mode?: number | Range | RangeReference | FunctionResult<any>): FunctionResult<number>;
    /**
     *
     * Rounds a number up, to the nearest integer or to the nearest multiple of significance.
     *
     * [Api set: ExcelApi 1.2]
     *
     * @param number Is the value you want to round.
     * @param significance Is the multiple to which you want to round.
     */
    ceiling_Precise(number: number | Range | RangeReference | FunctionResult<any>, significance?: number | Range | RangeReference | FunctionResult<any>): FunctionResult<number>;
    /**
     *
     * Returns the character specified by the code number from the character set for your computer.
     *
     * [Api set: ExcelApi 1.2]
     *
     * @param number Is a number between 1 and 255 specifying which character you want.
     */
    char(number: number | Range | RangeReference | FunctionResult<any>): FunctionResult<string>;
    /**
     *
     * Returns the left-tailed probability of the chi-squared distribution.
     *
     * [Api set: ExcelApi 1.2]
     *
     * @param x Is the value at which you want to evaluate the distribution, a nonnegative number.
     * @param degFreedom Is the number of degrees of freedom, a number between 1 and 10^10, excluding 10^10.
     * @param cumulative Is a logical value for the function to return: the cumulative distribution function = TRUE; the probability density function = FALSE.
     */
    chiSq_Dist(x: number | Range | RangeReference | FunctionResult<any>, degFreedom: number | Range | RangeReference | FunctionResult<any>, cumulative: boolean | Range | RangeReference | FunctionResult<any>): FunctionResult<number>;
    /**
     *
     * Returns the right-tailed probability of the chi-squared distribution.
     *
     * [Api set: ExcelApi 1.2]
     *
     * @param x Is the value at which you want to evaluate the distribution, a nonnegative number.
     * @param degFreedom Is the number of degrees of freedom, a number between 1 and 10^10, excluding 10^10.
     */
    chiSq_Dist_RT(x: number | Range | RangeReference | FunctionResult<any>, degFreedom: number | Range | RangeReference | FunctionResult<any>): FunctionResult<number>;
    /**
     *
     * Returns the inverse of the left-tailed probability of the chi-squared distribution.
     *
     * [Api set: ExcelApi 1.2]
     *
     * @param probability Is a probability associated with the chi-squared distribution, a value between 0 and 1 inclusive.
     * @param degFreedom Is the number of degrees of freedom, a number between 1 and 10^10, excluding 10^10.
     */
    chiSq_Inv(probability: number | Range | RangeReference | FunctionResult<any>, degFreedom: number | Range | RangeReference | FunctionResult<any>): FunctionResult<number>;
    /**
     *
     * Returns the inverse of the right-tailed probability of the chi-squared distribution.
     *
     * [Api set: ExcelApi 1.2]
     *
     * @param probability Is a probability associated with the chi-squared distribution, a value between 0 and 1 inclusive.
     * @param degFreedom Is the number of degrees of freedom, a number between 1 and 10^10, excluding 10^10.
     */
    chiSq_Inv_RT(probability: number | Range | RangeReference | FunctionResult<any>, degFreedom: number | Range | RangeReference | FunctionResult<any>): FunctionResult<number>;
    /**
     *
     * Chooses a value or action to perform from a list of values, based on an index number.
     *
     * [Api set: ExcelApi 1.2]
     *
     * @param indexNum Specifies which value argument is selected. indexNum must be between 1 and 254, or a formula or a reference to a number between 1 and 254.
     * @param values List of parameters, whose elements are 1 to 254 numbers, cell references, defined names, formulas, functions, or text arguments from which CHOOSE selects.
     */
    choose(indexNum: number | Range | RangeReference | FunctionResult<any>, ...values: Array<Range | number | string | boolean | RangeReference | FunctionResult<any>>): FunctionResult<number | string | boolean>;
    /**
     *
     * Removes all nonprintable characters from text.
     *
     * [Api set: ExcelApi 1.2]
     *
     * @param text Is any worksheet information from which you want to remove nonprintable characters.
     */
    clean(text: string | Range | RangeReference | FunctionResult<any>): FunctionResult<string>;
    /**
     *
     * Returns a numeric code for the first character in a text string, in the character set used by your computer.
     *
     * [Api set: ExcelApi 1.2]
     *
     * @param text Is the text for which you want the code of the first character.
     */
    code(text: string | Range | RangeReference | FunctionResult<any>): FunctionResult<number>;
    /**
     *
     * Returns the number of columns in an array or reference.
     *
     * [Api set: ExcelApi 1.2]
     *
     * @param array Is an array or array formula, or a reference to a range of cells for which you want the number of columns.
     */
    columns(array: Range | RangeReference | FunctionResult<any>): FunctionResult<number>;
    /**
     *
     * Returns the number of combinations for a given number of items.
     *
     * [Api set: ExcelApi 1.2]
     *
     * @param number Is the total number of items.
     * @param numberChosen Is the number of items in each combination.
     */
    combin(number: number | Range | RangeReference | FunctionResult<any>, numberChosen: number | Range | RangeReference | FunctionResult<any>): FunctionResult<number>;
    /**
     *
     * Returns the number of combinations with repetitions for a given number of items.
     *
     * [Api set: ExcelApi 1.2]
     *
     * @param number Is the total number of items.
     * @param numberChosen Is the number of items in each combination.
     */
    combina(number: number | Range | RangeReference | FunctionResult<any>, numberChosen: number | Range | RangeReference | FunctionResult<any>): FunctionResult<number>;
    /**
     *
     * Converts real and imaginary coefficients into a complex number.
     *
     * [Api set: ExcelApi 1.2]
     *
     * @param realNum Is the real coefficient of the complex number.
     * @param iNum Is the imaginary coefficient of the complex number.
     * @param suffix Is the suffix for the imaginary component of the complex number.
     */
    complex(realNum: number | string | boolean | Range | RangeReference | FunctionResult<any>, iNum: number | string | boolean | Range | RangeReference | FunctionResult<any>, suffix?: number | string | boolean | Range | RangeReference | FunctionResult<any>): FunctionResult<number>;
    /**
     *
     * Joins several text strings into one text string.
     *
     * [Api set: ExcelApi 1.2]
     *
     * @param values List of parameters, whose elements are 1 to 255 text strings to be joined into a single text string and can be text strings, numbers, or single-cell references.
     */
    concatenate(...values: Array<string | Range | RangeReference | FunctionResult<any>>): FunctionResult<string>;
    /**
     *
     * Returns the confidence interval for a population mean, using a normal distribution.
     *
     * [Api set: ExcelApi 1.2]
     *
     * @param alpha Is the significance level used to compute the confidence level, a number greater than 0 and less than 1.
     * @param standardDev Is the population standard deviation for the data range and is assumed to be known. standardDev must be greater than 0.
     * @param size Is the sample size.
     */
    confidence_Norm(alpha: number | Range | RangeReference | FunctionResult<any>, standardDev: number | Range | RangeReference | FunctionResult<any>, size: number | Range | RangeReference | FunctionResult<any>): FunctionResult<number>;
    /**
     *
     * Returns the confidence interval for a population mean, using a Student's T distribution.
     *
     * [Api set: ExcelApi 1.2]
     *
     * @param alpha Is the significance level used to compute the confidence level, a number greater than 0 and less than 1.
     * @param standardDev Is the population standard deviation for the data range and is assumed to be known. standardDev must be greater than 0.
     * @param size Is the sample size.
     */
    confidence_T(alpha: number | Range | RangeReference | FunctionResult<any>, standardDev: number | Range | RangeReference | FunctionResult<any>, size: number | Range | RangeReference | FunctionResult<any>): FunctionResult<number>;
    /**
     *
     * Converts a number from one measurement system to another.
     *
     * [Api set: ExcelApi 1.2]
     *
     * @param number Is the value in from_units to convert.
     * @param fromUnit Is the units for number.
     * @param toUnit Is the units for the result.
     */
    convert(number: number | string | boolean | Range | RangeReference | FunctionResult<any>, fromUnit: number | string | boolean | Range | RangeReference | FunctionResult<any>, toUnit: number | string | boolean | Range | RangeReference | FunctionResult<any>): FunctionResult<number>;
    /**
     *
     * Returns the cosine of an angle.
     *
     * [Api set: ExcelApi 1.2]
     *
     * @param number Is the angle in radians for which you want the cosine.
     */
    cos(number: number | Range | RangeReference | FunctionResult<any>): FunctionResult<number>;
    /**
     *
     * Returns the hyperbolic cosine of a number.
     *
     * [Api set: ExcelApi 1.2]
     *
     * @param number Is any real number.
     */
    cosh(number: number | Range | RangeReference | FunctionResult<any>): FunctionResult<number>;
    /**
     *
     * Returns the cotangent of an angle.
     *
     * [Api set: ExcelApi 1.2]
     *
     * @param number Is the angle in radians for which you want the cotangent.
     */
    cot(number: number | Range | RangeReference | FunctionResult<any>): FunctionResult<number>;
    /**
     *
     * Returns the hyperbolic cotangent of a number.
     *
     * [Api set: ExcelApi 1.2]
     *
     * @param number Is the angle in radians for which you want the hyperbolic cotangent.
     */
    coth(number: number | Range | RangeReference | FunctionResult<any>): FunctionResult<number>;
    /**
     *
     * Counts the number of cells in a range that contain numbers.
     *
     * [Api set: ExcelApi 1.2]
     *
     * @param values List of parameters, whose elements are 1 to 255 arguments that can contain or refer to a variety of different types of data, but only numbers are counted.
     */
    count(...values: Array<number | Range | RangeReference | FunctionResult<any>>): FunctionResult<number>;
    /**
     *
     * Counts the number of cells in a range that are not empty.
     *
     * [Api set: ExcelApi 1.2]
     *
     * @param values List of parameters, whose elements are 1 to 255 arguments representing the values and cells you want to count. Values can be any type of information.
     */
    countA(...values: Array<number | Range | RangeReference | FunctionResult<any>>): FunctionResult<number>;
    /**
     *
     * Counts the number of empty cells in a specified range of cells.
     *
     * [Api set: ExcelApi 1.2]
     *
     * @param range Is the range from which you want to count the empty cells.
     */
    countBlank(range: Range | RangeReference | FunctionResult<any>): FunctionResult<number>;
    /**
     *
     * Counts the number of cells within a range that meet the given condition.
     *
     * [Api set: ExcelApi 1.2]
     *
     * @param range Is the range of cells from which you want to count nonblank cells.
     * @param criteria Is the condition in the form of a number, expression, or text that defines which cells will be counted.
     */
    countIf(range: Range | RangeReference | FunctionResult<any>, criteria: number | string | boolean | Range | RangeReference | FunctionResult<any>): FunctionResult<number>;
    /**
     *
     * Counts the number of cells specified by a given set of conditions or criteria.
     *
     * [Api set: ExcelApi 1.2]
     *
     * @param values List of parameters, where the first element of each pair is the Is the range of cells you want evaluated for the particular condition , and the second element is is the condition in the form of a number, expression, or text that defines which cells will be counted.
     */
    countIfs(...values: Array<Range | RangeReference | FunctionResult<any> | number | string | boolean>): FunctionResult<number>;
    /**
     *
     * Returns the number of days from the beginning of the coupon period to the settlement date.
     *
     * [Api set: ExcelApi 1.2]
     *
     * @param settlement Is the security's settlement date, expressed as a serial date number.
     * @param maturity Is the security's maturity date, expressed as a serial date number.
     * @param frequency Is the number of coupon payments per year.
     * @param basis Is the type of day count basis to use.
     */
    coupDayBs(settlement: number | string | boolean | Range | RangeReference | FunctionResult<any>, maturity: number | string | boolean | Range | RangeReference | FunctionResult<any>, frequency: number | string | boolean | Range | RangeReference | FunctionResult<any>, basis?: number | string | boolean | Range | RangeReference | FunctionResult<any>): FunctionResult<number>;
    /**
     *
     * Returns the number of days in the coupon period that contains the settlement date.
     *
     * [Api set: ExcelApi 1.2]
     *
     * @param settlement Is the security's settlement date, expressed as a serial date number.
     * @param maturity Is the security's maturity date, expressed as a serial date number.
     * @param frequency Is the number of coupon payments per year.
     * @param basis Is the type of day count basis to use.
     */
    coupDays(settlement: number | string | boolean | Range | RangeReference | FunctionResult<any>, maturity: number | string | boolean | Range | RangeReference | FunctionResult<any>, frequency: number | string | boolean | Range | RangeReference | FunctionResult<any>, basis?: number | string | boolean | Range | RangeReference | FunctionResult<any>): FunctionResult<number>;
    /**
     *
     * Returns the number of days from the settlement date to the next coupon date.
     *
     * [Api set: ExcelApi 1.2]
     *
     * @param settlement Is the security's settlement date, expressed as a serial date number.
     * @param maturity Is the security's maturity date, expressed as a serial date number.
     * @param frequency Is the number of coupon payments per year.
     * @param basis Is the type of day count basis to use.
     */
    coupDaysNc(settlement: number | string | boolean | Range | RangeReference | FunctionResult<any>, maturity: number | string | boolean | Range | RangeReference | FunctionResult<any>, frequency: number | string | boolean | Range | RangeReference | FunctionResult<any>, basis?: number | string | boolean | Range | RangeReference | FunctionResult<any>): FunctionResult<number>;
    /**
     *
     * Returns the next coupon date after the settlement date.
     *
     * [Api set: ExcelApi 1.2]
     *
     * @param settlement Is the security's settlement date, expressed as a serial date number.
     * @param maturity Is the security's maturity date, expressed as a serial date number.
     * @param frequency Is the number of coupon payments per year.
     * @param basis Is the type of day count basis to use.
     */
    coupNcd(settlement: number | string | boolean | Range | RangeReference | FunctionResult<any>, maturity: number | string | boolean | Range | RangeReference | FunctionResult<any>, frequency: number | string | boolean | Range | RangeReference | FunctionResult<any>, basis?: number | string | boolean | Range | RangeReference | FunctionResult<any>): FunctionResult<number>;
    /**
     *
     * Returns the number of coupons payable between the settlement date and maturity date.
     *
     * [Api set: ExcelApi 1.2]
     *
     * @param settlement Is the security's settlement date, expressed as a serial date number.
     * @param maturity Is the security's maturity date, expressed as a serial date number.
     * @param frequency Is the number of coupon payments per year.
     * @param basis Is the type of day count basis to use.
     */
    coupNum(settlement: number | string | boolean | Range | RangeReference | FunctionResult<any>, maturity: number | string | boolean | Range | RangeReference | FunctionResult<any>, frequency: number | string | boolean | Range | RangeReference | FunctionResult<any>, basis?: number | string | boolean | Range | RangeReference | FunctionResult<any>): FunctionResult<number>;
    /**
     *
     * Returns the previous coupon date before the settlement date.
     *
     * [Api set: ExcelApi 1.2]
     *
     * @param settlement Is the security's settlement date, expressed as a serial date number.
     * @param maturity Is the security's maturity date, expressed as a serial date number.
     * @param frequency Is the number of coupon payments per year.
     * @param basis Is the type of day count basis to use.
     */
    coupPcd(settlement: number | string | boolean | Range | RangeReference | FunctionResult<any>, maturity: number | string | boolean | Range | RangeReference | FunctionResult<any>, frequency: number | string | boolean | Range | RangeReference | FunctionResult<any>, basis?: number | string | boolean | Range | RangeReference | FunctionResult<any>): FunctionResult<number>;
    /**
     *
     * Returns the cosecant of an angle.
     *
     * [Api set: ExcelApi 1.2]
     *
     * @param number Is the angle in radians for which you want the cosecant.
     */
    csc(number: number | Range | RangeReference | FunctionResult<any>): FunctionResult<number>;
    /**
     *
     * Returns the hyperbolic cosecant of an angle.
     *
     * [Api set: ExcelApi 1.2]
     *
     * @param number Is the angle in radians for which you want the hyperbolic cosecant.
     */
    csch(number: number | Range | RangeReference | FunctionResult<any>): FunctionResult<number>;
    /**
     *
     * Returns the cumulative interest paid between two periods.
     *
     * [Api set: ExcelApi 1.2]
     *
     * @param rate Is the interest rate.
     * @param nper Is the total number of payment periods.
     * @param pv Is the present value.
     * @param startPeriod Is the first period in the calculation.
     * @param endPeriod Is the last period in the calculation.
     * @param type Is the timing of the payment.
     */
    cumIPmt(rate: number | string | boolean | Range | RangeReference | FunctionResult<any>, nper: number | string | boolean | Range | RangeReference | FunctionResult<any>, pv: number | string | boolean | Range | RangeReference | FunctionResult<any>, startPeriod: number | string | boolean | Range | RangeReference | FunctionResult<any>, endPeriod: number | string | boolean | Range | RangeReference | FunctionResult<any>, type: number | string | boolean | Range | RangeReference | FunctionResult<any>): FunctionResult<number>;
    /**
     *
     * Returns the cumulative principal paid on a loan between two periods.
     *
     * [Api set: ExcelApi 1.2]
     *
     * @param rate Is the interest rate.
     * @param nper Is the total number of payment periods.
     * @param pv Is the present value.
     * @param startPeriod Is the first period in the calculation.
     * @param endPeriod Is the last period in the calculation.
     * @param type Is the timing of the payment.
     */
    cumPrinc(rate: number | string | boolean | Range | RangeReference | FunctionResult<any>, nper: number | string | boolean | Range | RangeReference | FunctionResult<any>, pv: number | string | boolean | Range | RangeReference | FunctionResult<any>, startPeriod: number | string | boolean | Range | RangeReference | FunctionResult<any>, endPeriod: number | string | boolean | Range | RangeReference | FunctionResult<any>, type: number | string | boolean | Range | RangeReference | FunctionResult<any>): FunctionResult<number>;
    /**
     *
     * Averages the values in a column in a list or database that match conditions you specify.
     *
     * [Api set: ExcelApi 1.2]
     *
     * @param database Is the range of cells that makes up the list or database. A database is a list of related data.
     * @param field Is either the label of the column in double quotation marks or a number that represents the column's position in the list.
     * @param criteria Is the range of cells that contains the conditions you specify. The range includes a column label and one cell below the label for a condition.
     */
    daverage(database: Range | RangeReference | FunctionResult<any>, field: number | string | Range | RangeReference | FunctionResult<any>, criteria: string | Range | RangeReference | FunctionResult<any>): FunctionResult<number>;
    /**
     *
     * Counts the cells containing numbers in the field (column) of records in the database that match the conditions you specify.
     *
     * [Api set: ExcelApi 1.2]
     *
     * @param database Is the range of cells that makes up the list or database. A database is a list of related data.
     * @param field Is either the label of the column in double quotation marks or a number that represents the column's position in the list.
     * @param criteria Is the range of cells that contains the conditions you specify. The range includes a column label and one cell below the label for a condition.
     */
    dcount(database: Range | RangeReference | FunctionResult<any>, field: number | string | Range | RangeReference | FunctionResult<any>, criteria: string | Range | RangeReference | FunctionResult<any>): FunctionResult<number>;
    /**
     *
     * Counts nonblank cells in the field (column) of records in the database that match the conditions you specify.
     *
     * [Api set: ExcelApi 1.2]
     *
     * @param database Is the range of cells that makes up the list or database. A database is a list of related data.
     * @param field Is either the label of the column in double quotation marks or a number that represents the column's position in the list.
     * @param criteria Is the range of cells that contains the conditions you specify. The range includes a column label and one cell below the label for a condition.
     */
    dcountA(database: Range | RangeReference | FunctionResult<any>, field: number | string | Range | RangeReference | FunctionResult<any>, criteria: string | Range | RangeReference | FunctionResult<any>): FunctionResult<number>;
    /**
     *
     * Extracts from a database a single record that matches the conditions you specify.
     *
     * [Api set: ExcelApi 1.2]
     *
     * @param database Is the range of cells that makes up the list or database. A database is a list of related data.
     * @param field Is either the label of the column in double quotation marks or a number that represents the column's position in the list.
     * @param criteria Is the range of cells that contains the conditions you specify. The range includes a column label and one cell below the label for a condition.
     */
    dget(database: Range | RangeReference | FunctionResult<any>, field: number | string | Range | RangeReference | FunctionResult<any>, criteria: string | Range | RangeReference | FunctionResult<any>): FunctionResult<number | boolean | string>;
    /**
     *
     * Returns the largest number in the field (column) of records in the database that match the conditions you specify.
     *
     * [Api set: ExcelApi 1.2]
     *
     * @param database Is the range of cells that makes up the list or database. A database is a list of related data.
     * @param field Is either the label of the column in double quotation marks or a number that represents the column's position in the list.
     * @param criteria Is the range of cells that contains the conditions you specify. The range includes a column label and one cell below the label for a condition.
     */
    dmax(database: Range | RangeReference | FunctionResult<any>, field: number | string | Range | RangeReference | FunctionResult<any>, criteria: string | Range | RangeReference | FunctionResult<any>): FunctionResult<number>;
    /**
     *
     * Returns the smallest number in the field (column) of records in the database that match the conditions you specify.
     *
     * [Api set: ExcelApi 1.2]
     *
     * @param database Is the range of cells that makes up the list or database. A database is a list of related data.
     * @param field Is either the label of the column in double quotation marks or a number that represents the column's position in the list.
     * @param criteria Is the range of cells that contains the conditions you specify. The range includes a column label and one cell below the label for a condition.
     */
    dmin(database: Range | RangeReference | FunctionResult<any>, field: number | string | Range | RangeReference | FunctionResult<any>, criteria: string | Range | RangeReference | FunctionResult<any>): FunctionResult<number>;
    /**
     *
     * Multiplies the values in the field (column) of records in the database that match the conditions you specify.
     *
     * [Api set: ExcelApi 1.2]
     *
     * @param database Is the range of cells that makes up the list or database. A database is a list of related data.
     * @param field Is either the label of the column in double quotation marks or a number that represents the column's position in the list.
     * @param criteria Is the range of cells that contains the conditions you specify. The range includes a column label and one cell below the label for a condition.
     */
    dproduct(database: Range | RangeReference | FunctionResult<any>, field: number | string | Range | RangeReference | FunctionResult<any>, criteria: string | Range | RangeReference | FunctionResult<any>): FunctionResult<number>;
    /**
     *
     * Estimates the standard deviation based on a sample from selected database entries.
     *
     * [Api set: ExcelApi 1.2]
     *
     * @param database Is the range of cells that makes up the list or database. A database is a list of related data.
     * @param field Is either the label of the column in double quotation marks or a number that represents the column's position in the list.
     * @param criteria Is the range of cells that contains the conditions you specify. The range includes a column label and one cell below the label for a condition.
     */
    dstDev(database: Range | RangeReference | FunctionResult<any>, field: number | string | Range | RangeReference | FunctionResult<any>, criteria: string | Range | RangeReference | FunctionResult<any>): FunctionResult<number>;
    /**
     *
     * Calculates the standard deviation based on the entire population of selected database entries.
     *
     * [Api set: ExcelApi 1.2]
     *
     * @param database Is the range of cells that makes up the list or database. A database is a list of related data.
     * @param field Is either the label of the column in double quotation marks or a number that represents the column's position in the list.
     * @param criteria Is the range of cells that contains the conditions you specify. The range includes a column label and one cell below the label for a condition.
     */
    dstDevP(database: Range | RangeReference | FunctionResult<any>, field: number | string | Range | RangeReference | FunctionResult<any>, criteria: string | Range | RangeReference | FunctionResult<any>): FunctionResult<number>;
    /**
     *
     * Adds the numbers in the field (column) of records in the database that match the conditions you specify.
     *
     * [Api set: ExcelApi 1.2]
     *
     * @param database Is the range of cells that makes up the list or database. A database is a list of related data.
     * @param field Is either the label of the column in double quotation marks or a number that represents the column's position in the list.
     * @param criteria Is the range of cells that contains the conditions you specify. The range includes a column label and one cell below the label for a condition.
     */
    dsum(database: Range | RangeReference | FunctionResult<any>, field: number | string | Range | RangeReference | FunctionResult<any>, criteria: string | Range | RangeReference | FunctionResult<any>): FunctionResult<number>;
    /**
     *
     * Estimates variance based on a sample from selected database entries.
     *
     * [Api set: ExcelApi 1.2]
     *
     * @param database Is the range of cells that makes up the list or database. A database is a list of related data.
     * @param field Is either the label of the column in double quotation marks or a number that represents the column's position in the list.
     * @param criteria Is the range of cells that contains the conditions you specify. The range includes a column label and one cell below the label for a condition.
     */
    dvar(database: Range | RangeReference | FunctionResult<any>, field: number | string | Range | RangeReference | FunctionResult<any>, criteria: string | Range | RangeReference | FunctionResult<any>): FunctionResult<number>;
    /**
     *
     * Calculates variance based on the entire population of selected database entries.
     *
     * [Api set: ExcelApi 1.2]
     *
     * @param database Is the range of cells that makes up the list or database. A database is a list of related data.
     * @param field Is either the label of the column in double quotation marks or a number that represents the column's position in the list.
     * @param criteria Is the range of cells that contains the conditions you specify. The range includes a column label and one cell below the label for a condition.
     */
    dvarP(database: Range | RangeReference | FunctionResult<any>, field: number | string | Range | RangeReference | FunctionResult<any>, criteria: string | Range | RangeReference | FunctionResult<any>): FunctionResult<number>;
    /**
     *
     * Returns the number that represents the date in Microsoft Excel date-time code.
     *
     * [Api set: ExcelApi 1.2]
     *
     * @param year Is a number from 1900 or 1904 (depending on the workbook's date system) to 9999.
     * @param month Is a number from 1 to 12 representing the month of the year.
     * @param day Is a number from 1 to 31 representing the day of the month.
     */
    date(year: number | Range | RangeReference | FunctionResult<any>, month: number | Range | RangeReference | FunctionResult<any>, day: number | Range | RangeReference | FunctionResult<any>): FunctionResult<number>;
    /**
     *
     * Converts a date in the form of text to a number that represents the date in Microsoft Excel date-time code.
     *
     * [Api set: ExcelApi 1.2]
     *
     * @param dateText Is text that represents a date in a Microsoft Excel date format, between 1/1/1900 or 1/1/1904 (depending on the workbook's date system) and 12/31/9999.
     */
    datevalue(dateText: string | number | Range | RangeReference | FunctionResult<any>): FunctionResult<number>;
    /**
     *
     * Returns the day of the month, a number from 1 to 31.
     *
     * [Api set: ExcelApi 1.2]
     *
     * @param serialNumber Is a number in the date-time code used by Microsoft
     */
    day(serialNumber: number | Range | RangeReference | FunctionResult<any>): FunctionResult<number>;
    /**
     *
     * Returns the number of days between the two dates.
     *
     * [Api set: ExcelApi 1.2]
     *
     * @param endDate startDate and endDate are the two dates between which you want to know the number of days.
     * @param startDate startDate and endDate are the two dates between which you want to know the number of days.
     */
    days(endDate: string | number | Range | RangeReference | FunctionResult<any>, startDate: string | number | Range | RangeReference | FunctionResult<any>): FunctionResult<number>;
    /**
     *
     * Returns the number of days between two dates based on a 360-day year (twelve 30-day months).
     *
     * [Api set: ExcelApi 1.2]
     *
     * @param startDate startDate and endDate are the two dates between which you want to know the number of days.
     * @param endDate startDate and endDate are the two dates between which you want to know the number of days.
     * @param method Is a logical value specifying the calculation method: U.S. (NASD) = FALSE or omitted; European = TRUE.
     */
    days360(startDate: number | Range | RangeReference | FunctionResult<any>, endDate: number | Range | RangeReference | FunctionResult<any>, method?: boolean | Range | RangeReference | FunctionResult<any>): FunctionResult<number>;
    /**
     *
     * Returns the depreciation of an asset for a specified period using the fixed-declining balance method.
     *
     * [Api set: ExcelApi 1.2]
     *
     * @param cost Is the initial cost of the asset.
     * @param salvage Is the salvage value at the end of the life of the asset.
     * @param life Is the number of periods over which the asset is being depreciated (sometimes called the useful life of the asset).
     * @param period Is the period for which you want to calculate the depreciation. Period must use the same units as Life.
     * @param month Is the number of months in the first year. If month is omitted, it is assumed to be 12.
     */
    db(cost: number | Range | RangeReference | FunctionResult<any>, salvage: number | Range | RangeReference | FunctionResult<any>, life: number | Range | RangeReference | FunctionResult<any>, period: number | Range | RangeReference | FunctionResult<any>, month?: number | Range | RangeReference | FunctionResult<any>): FunctionResult<number>;
    /**
     *
     * Changes half-width (single-byte) characters within a character string to full-width (double-byte) characters. Use with double-byte character sets (DBCS).
     *
     * [Api set: ExcelApi 1.2]
     *
     * @param text Is a text, or a reference to a cell containing a text.
     */
    dbcs(text: string | Range | RangeReference | FunctionResult<any>): FunctionResult<string>;
    /**
     *
     * Returns the depreciation of an asset for a specified period using the double-declining balance method or some other method you specify.
     *
     * [Api set: ExcelApi 1.2]
     *
     * @param cost Is the initial cost of the asset.
     * @param salvage Is the salvage value at the end of the life of the asset.
     * @param life Is the number of periods over which the asset is being depreciated (sometimes called the useful life of the asset).
     * @param period Is the period for which you want to calculate the depreciation. Period must use the same units as Life.
     * @param factor Is the rate at which the balance declines. If Factor is omitted, it is assumed to be 2 (the double-declining balance method).
     */
    ddb(cost: number | Range | RangeReference | FunctionResult<any>, salvage: number | Range | RangeReference | FunctionResult<any>, life: number | Range | RangeReference | FunctionResult<any>, period: number | Range | RangeReference | FunctionResult<any>, factor?: number | Range | RangeReference | FunctionResult<any>): FunctionResult<number>;
    /**
     *
     * Converts a decimal number to binary.
     *
     * [Api set: ExcelApi 1.2]
     *
     * @param number Is the decimal integer you want to convert.
     * @param places Is the number of characters to use.
     */
    dec2Bin(number: number | string | boolean | Range | RangeReference | FunctionResult<any>, places?: number | string | boolean | Range | RangeReference | FunctionResult<any>): FunctionResult<number>;
    /**
     *
     * Converts a decimal number to hexadecimal.
     *
     * [Api set: ExcelApi 1.2]
     *
     * @param number Is the decimal integer you want to convert.
     * @param places Is the number of characters to use.
     */
    dec2Hex(number: number | string | boolean | Range | RangeReference | FunctionResult<any>, places?: number | string | boolean | Range | RangeReference | FunctionResult<any>): FunctionResult<number>;
    /**
     *
     * Converts a decimal number to octal.
     *
     * [Api set: ExcelApi 1.2]
     *
     * @param number Is the decimal integer you want to convert.
     * @param places Is the number of characters to use.
     */
    dec2Oct(number: number | string | boolean | Range | RangeReference | FunctionResult<any>, places?: number | string | boolean | Range | RangeReference | FunctionResult<any>): FunctionResult<number>;
    /**
     *
     * Converts a text representation of a number in a given base into a decimal number.
     *
     * [Api set: ExcelApi 1.2]
     *
     * @param number Is the number that you want to convert.
     * @param radix Is the base Radix of the number you are converting.
     */
    decimal(number: string | Range | RangeReference | FunctionResult<any>, radix: number | Range | RangeReference | FunctionResult<any>): FunctionResult<number>;
    /**
     *
     * Converts radians to degrees.
     *
     * [Api set: ExcelApi 1.2]
     *
     * @param angle Is the angle in radians that you want to convert.
     */
    degrees(angle: number | Range | RangeReference | FunctionResult<any>): FunctionResult<number>;
    /**
     *
     * Tests whether two numbers are equal.
     *
     * [Api set: ExcelApi 1.2]
     *
     * @param number1 Is the first number.
     * @param number2 Is the second number.
     */
    delta(number1: number | string | boolean | Range | RangeReference | FunctionResult<any>, number2?: number | string | boolean | Range | RangeReference | FunctionResult<any>): FunctionResult<number>;
    /**
     *
     * Returns the sum of squares of deviations of data points from their sample mean.
     *
     * [Api set: ExcelApi 1.2]
     *
     * @param values List of parameters, whose elements are 1 to 255 arguments, or an array or array reference, on which you want DEVSQ to calculate.
     */
    devSq(...values: Array<number | Range | RangeReference | FunctionResult<any>>): FunctionResult<number>;
    /**
     *
     * Returns the discount rate for a security.
     *
     * [Api set: ExcelApi 1.2]
     *
     * @param settlement Is the security's settlement date, expressed as a serial date number.
     * @param maturity Is the security's maturity date, expressed as a serial date number.
     * @param pr Is the security's price per $100 face value.
     * @param redemption Is the security's redemption value per $100 face value.
     * @param basis Is the type of day count basis to use.
     */
    disc(settlement: number | string | boolean | Range | RangeReference | FunctionResult<any>, maturity: number | string | boolean | Range | RangeReference | FunctionResult<any>, pr: number | string | boolean | Range | RangeReference | FunctionResult<any>, redemption: number | string | boolean | Range | RangeReference | FunctionResult<any>, basis?: number | string | boolean | Range | RangeReference | FunctionResult<any>): FunctionResult<number>;
    /**
     *
     * Converts a number to text, using currency format.
     *
     * [Api set: ExcelApi 1.2]
     *
     * @param number Is a number, a reference to a cell containing a number, or a formula that evaluates to a number.
     * @param decimals Is the number of digits to the right of the decimal point. The number is rounded as necessary; if omitted, Decimals = 2.
     */
    dollar(number: number | Range | RangeReference | FunctionResult<any>, decimals?: number | Range | RangeReference | FunctionResult<any>): FunctionResult<string>;
    /**
     *
     * Converts a dollar price, expressed as a fraction, into a dollar price, expressed as a decimal number.
     *
     * [Api set: ExcelApi 1.2]
     *
     * @param fractionalDollar Is a number expressed as a fraction.
     * @param fraction Is the integer to use in the denominator of the fraction.
     */
    dollarDe(fractionalDollar: number | string | boolean | Range | RangeReference | FunctionResult<any>, fraction: number | string | boolean | Range | RangeReference | FunctionResult<any>): FunctionResult<number>;
    /**
     *
     * Converts a dollar price, expressed as a decimal number, into a dollar price, expressed as a fraction.
     *
     * [Api set: ExcelApi 1.2]
     *
     * @param decimalDollar Is a decimal number.
     * @param fraction Is the integer to use in the denominator of a fraction.
     */
    dollarFr(decimalDollar: number | string | boolean | Range | RangeReference | FunctionResult<any>, fraction: number | string | boolean | Range | RangeReference | FunctionResult<any>): FunctionResult<number>;
    /**
     *
     * Returns the annual duration of a security with periodic interest payments.
     *
     * [Api set: ExcelApi 1.2]
     *
     * @param settlement Is the security's settlement date, expressed as a serial date number.
     * @param maturity Is the security's maturity date, expressed as a serial date number.
     * @param coupon Is the security's annual coupon rate.
     * @param yld Is the security's annual yield.
     * @param frequency Is the number of coupon payments per year.
     * @param basis Is the type of day count basis to use.
     */
    duration(settlement: number | string | boolean | Range | RangeReference | FunctionResult<any>, maturity: number | string | boolean | Range | RangeReference | FunctionResult<any>, coupon: number | string | boolean | Range | RangeReference | FunctionResult<any>, yld: number | string | boolean | Range | RangeReference | FunctionResult<any>, frequency: number | string | boolean | Range | RangeReference | FunctionResult<any>, basis?: number | string | boolean | Range | RangeReference | FunctionResult<any>): FunctionResult<number>;
    /**
     *
     * Rounds a number up, to the nearest integer or to the nearest multiple of significance.
     *
     * [Api set: ExcelApi 1.2]
     *
     * @param number Is the value you want to round.
     * @param significance Is the multiple to which you want to round.
     */
    ecma_Ceiling(number: number | Range | RangeReference | FunctionResult<any>, significance: number | Range | RangeReference | FunctionResult<any>): FunctionResult<number>;
    /**
     *
     * Returns the serial number of the date that is the indicated number of months before or after the start date.
     *
     * [Api set: ExcelApi 1.2]
     *
     * @param startDate Is a serial date number that represents the start date.
     * @param months Is the number of months before or after startDate.
     */
    edate(startDate: number | string | boolean | Range | RangeReference | FunctionResult<any>, months: number | string | boolean | Range | RangeReference | FunctionResult<any>): FunctionResult<number>;
    /**
     *
     * Returns the effective annual interest rate.
     *
     * [Api set: ExcelApi 1.2]
     *
     * @param nominalRate Is the nominal interest rate.
     * @param npery Is the number of compounding periods per year.
     */
    effect(nominalRate: number | string | boolean | Range | RangeReference | FunctionResult<any>, npery: number | string | boolean | Range | RangeReference | FunctionResult<any>): FunctionResult<number>;
    /**
     *
     * Returns the serial number of the last day of the month before or after a specified number of months.
     *
     * [Api set: ExcelApi 1.2]
     *
     * @param startDate Is a serial date number that represents the start date.
     * @param months Is the number of months before or after the startDate.
     */
    eoMonth(startDate: number | string | boolean | Range | RangeReference | FunctionResult<any>, months: number | string | boolean | Range | RangeReference | FunctionResult<any>): FunctionResult<number>;
    /**
     *
     * Returns the error function.
     *
     * [Api set: ExcelApi 1.2]
     *
     * @param lowerLimit Is the lower bound for integrating ERF.
     * @param upperLimit Is the upper bound for integrating ERF.
     */
    erf(lowerLimit: number | string | boolean | Range | RangeReference | FunctionResult<any>, upperLimit?: number | string | boolean | Range | RangeReference | FunctionResult<any>): FunctionResult<number>;
    /**
     *
     * Returns the complementary error function.
     *
     * [Api set: ExcelApi 1.2]
     *
     * @param x Is the lower bound for integrating ERF.
     */
    erfC(x: number | string | boolean | Range | RangeReference | FunctionResult<any>): FunctionResult<number>;
    /**
     *
     * Returns the complementary error function.
     *
     * [Api set: ExcelApi 1.2]
     *
     * @param X Is the lower bound for integrating ERFC.PRECISE.
     */
    erfC_Precise(X: number | string | boolean | Range | RangeReference | FunctionResult<any>): FunctionResult<number>;
    /**
     *
     * Returns the error function.
     *
     * [Api set: ExcelApi 1.2]
     *
     * @param X Is the lower bound for integrating ERF.PRECISE.
     */
    erf_Precise(X: number | string | boolean | Range | RangeReference | FunctionResult<any>): FunctionResult<number>;
    /**
     *
     * Returns a number matching an error value.
     *
     * [Api set: ExcelApi 1.2]
     *
     * @param errorVal Is the error value for which you want the identifying number, and can be an actual error value or a reference to a cell containing an error value.
     */
    error_Type(errorVal: string | number | boolean | Range | RangeReference | FunctionResult<any>): FunctionResult<number>;
    /**
     *
     * Rounds a positive number up and negative number down to the nearest even integer.
     *
     * [Api set: ExcelApi 1.2]
     *
     * @param number Is the value to round.
     */
    even(number: number | Range | RangeReference | FunctionResult<any>): FunctionResult<number>;
    /**
     *
     * Checks whether two text strings are exactly the same, and returns TRUE or FALSE. EXACT is case-sensitive.
     *
     * [Api set: ExcelApi 1.2]
     *
     * @param text1 Is the first text string.
     * @param text2 Is the second text string.
     */
    exact(text1: string | Range | RangeReference | FunctionResult<any>, text2: string | Range | RangeReference | FunctionResult<any>): FunctionResult<boolean>;
    /**
     *
     * Returns e raised to the power of a given number.
     *
     * [Api set: ExcelApi 1.2]
     *
     * @param number Is the exponent applied to the base e. The constant e equals 2.71828182845904, the base of the natural logarithm.
     */
    exp(number: number | Range | RangeReference | FunctionResult<any>): FunctionResult<number>;
    /**
     *
     * Returns the exponential distribution.
     *
     * [Api set: ExcelApi 1.2]
     *
     * @param x Is the value of the function, a nonnegative number.
     * @param lambda Is the parameter value, a positive number.
     * @param cumulative Is a logical value for the function to return: the cumulative distribution function = TRUE; the probability density function = FALSE.
     */
    expon_Dist(x: number | Range | RangeReference | FunctionResult<any>, lambda: number | Range | RangeReference | FunctionResult<any>, cumulative: boolean | Range | RangeReference | FunctionResult<any>): FunctionResult<number>;
    /**
     *
     * Returns the future value of an initial principal after applying a series of compound interest rates.
     *
     * [Api set: ExcelApi 1.2]
     *
     * @param principal Is the present value.
     * @param schedule Is an array of interest rates to apply.
     */
    fvschedule(principal: number | string | boolean | Range | RangeReference | FunctionResult<any>, schedule: number | string | Range | boolean | RangeReference | FunctionResult<any>): FunctionResult<number>;
    /**
     *
     * Returns the (left-tailed) F probability distribution (degree of diversity) for two data sets.
     *
     * [Api set: ExcelApi 1.2]
     *
     * @param x Is the value at which to evaluate the function, a nonnegative number.
     * @param degFreedom1 Is the numerator degrees of freedom, a number between 1 and 10^10, excluding 10^10.
     * @param degFreedom2 Is the denominator degrees of freedom, a number between 1 and 10^10, excluding 10^10.
     * @param cumulative Is a logical value for the function to return: the cumulative distribution function = TRUE; the probability density function = FALSE.
     */
    f_Dist(x: number | Range | RangeReference | FunctionResult<any>, degFreedom1: number | Range | RangeReference | FunctionResult<any>, degFreedom2: number | Range | RangeReference | FunctionResult<any>, cumulative: boolean | Range | RangeReference | FunctionResult<any>): FunctionResult<number>;
    /**
     *
     * Returns the (right-tailed) F probability distribution (degree of diversity) for two data sets.
     *
     * [Api set: ExcelApi 1.2]
     *
     * @param x Is the value at which to evaluate the function, a nonnegative number.
     * @param degFreedom1 Is the numerator degrees of freedom, a number between 1 and 10^10, excluding 10^10.
     * @param degFreedom2 Is the denominator degrees of freedom, a number between 1 and 10^10, excluding 10^10.
     */
    f_Dist_RT(x: number | Range | RangeReference | FunctionResult<any>, degFreedom1: number | Range | RangeReference | FunctionResult<any>, degFreedom2: number | Range | RangeReference | FunctionResult<any>): FunctionResult<number>;
    /**
     *
     * Returns the inverse of the (left-tailed) F probability distribution: if p = F.DIST(x,...), then F.INV(p,...) = x.
     *
     * [Api set: ExcelApi 1.2]
     *
     * @param probability Is a probability associated with the F cumulative distribution, a number between 0 and 1 inclusive.
     * @param degFreedom1 Is the numerator degrees of freedom, a number between 1 and 10^10, excluding 10^10.
     * @param degFreedom2 Is the denominator degrees of freedom, a number between 1 and 10^10, excluding 10^10.
     */
    f_Inv(probability: number | Range | RangeReference | FunctionResult<any>, degFreedom1: number | Range | RangeReference | FunctionResult<any>, degFreedom2: number | Range | RangeReference | FunctionResult<any>): FunctionResult<number>;
    /**
     *
     * Returns the inverse of the (right-tailed) F probability distribution: if p = F.DIST.RT(x,...), then F.INV.RT(p,...) = x.
     *
     * [Api set: ExcelApi 1.2]
     *
     * @param probability Is a probability associated with the F cumulative distribution, a number between 0 and 1 inclusive.
     * @param degFreedom1 Is the numerator degrees of freedom, a number between 1 and 10^10, excluding 10^10.
     * @param degFreedom2 Is the denominator degrees of freedom, a number between 1 and 10^10, excluding 10^10.
     */
    f_Inv_RT(probability: number | Range | RangeReference | FunctionResult<any>, degFreedom1: number | Range | RangeReference | FunctionResult<any>, degFreedom2: number | Range | RangeReference | FunctionResult<any>): FunctionResult<number>;
    /**
     *
     * Returns the factorial of a number, equal to 1*2*3*...* Number.
     *
     * [Api set: ExcelApi 1.2]
     *
     * @param number Is the nonnegative number you want the factorial of.
     */
    fact(number: number | Range | RangeReference | FunctionResult<any>): FunctionResult<number>;
    /**
     *
     * Returns the double factorial of a number.
     *
     * [Api set: ExcelApi 1.2]
     *
     * @param number Is the value for which to return the double factorial.
     */
    factDouble(number: number | string | boolean | Range | RangeReference | FunctionResult<any>): FunctionResult<number>;
    /**
     *
     * Returns the logical value FALSE.
     *
     * [Api set: ExcelApi 1.2]
     */
    false(): FunctionResult<boolean>;
    /**
     *
     * Returns the starting position of one text string within another text string. FIND is case-sensitive.
     *
     * [Api set: ExcelApi 1.2]
     *
     * @param findText Is the text you want to find. Use double quotes (empty text) to match the first character in withinText; wildcard characters not allowed.
     * @param withinText Is the text containing the text you want to find.
     * @param startNum Specifies the character at which to start the search. The first character in withinText is character number 1. If omitted, startNum = 1.
     */
    find(findText: string | Range | RangeReference | FunctionResult<any>, withinText: string | Range | RangeReference | FunctionResult<any>, startNum?: number | Range | RangeReference | FunctionResult<any>): FunctionResult<number>;
    /**
     *
     * Finds the starting position of one text string within another text string. FINDB is case-sensitive. Use with double-byte character sets (DBCS).
     *
     * [Api set: ExcelApi 1.2]
     *
     * @param findText Is the text you want to find.
     * @param withinText Is the text containing the text you want to find.
     * @param startNum Specifies the character at which to start the search.
     */
    findB(findText: string | Range | RangeReference | FunctionResult<any>, withinText: string | Range | RangeReference | FunctionResult<any>, startNum?: number | Range | RangeReference | FunctionResult<any>): FunctionResult<number>;
    /**
     *
     * Returns the Fisher transformation.
     *
     * [Api set: ExcelApi 1.2]
     *
     * @param x Is the value for which you want the transformation, a number between -1 and 1, excluding -1 and 1.
     */
    fisher(x: number | Range | RangeReference | FunctionResult<any>): FunctionResult<number>;
    /**
     *
     * Returns the inverse of the Fisher transformation: if y = FISHER(x), then FISHERINV(y) = x.
     *
     * [Api set: ExcelApi 1.2]
     *
     * @param y Is the value for which you want to perform the inverse of the transformation.
     */
    fisherInv(y: number | Range | RangeReference | FunctionResult<any>): FunctionResult<number>;
    /**
     *
     * Rounds a number to the specified number of decimals and returns the result as text with or without commas.
     *
     * [Api set: ExcelApi 1.2]
     *
     * @param number Is the number you want to round and convert to text.
     * @param decimals Is the number of digits to the right of the decimal point. If omitted, Decimals = 2.
     * @param noCommas Is a logical value: do not display commas in the returned text = TRUE; do display commas in the returned text = FALSE or omitted.
     */
    fixed(number: number | Range | RangeReference | FunctionResult<any>, decimals?: number | Range | RangeReference | FunctionResult<any>, noCommas?: boolean | Range | RangeReference | FunctionResult<any>): FunctionResult<string>;
    /**
     *
     * Rounds a number down, to the nearest integer or to the nearest multiple of significance.
     *
     * [Api set: ExcelApi 1.2]
     *
     * @param number Is the value you want to round.
     * @param significance Is the multiple to which you want to round.
     * @param mode When given and nonzero this function will round towards zero.
     */
    floor_Math(number: number | Range | RangeReference | FunctionResult<any>, significance?: number | Range | RangeReference | FunctionResult<any>, mode?: number | Range | RangeReference | FunctionResult<any>): FunctionResult<number>;
    /**
     *
     * Rounds a number down, to the nearest integer or to the nearest multiple of significance.
     *
     * [Api set: ExcelApi 1.2]
     *
     * @param number Is the numeric value you want to round.
     * @param significance Is the multiple to which you want to round.
     */
    floor_Precise(number: number | Range | RangeReference | FunctionResult<any>, significance?: number | Range | RangeReference | FunctionResult<any>): FunctionResult<number>;
    /**
     *
     * Returns the future value of an investment based on periodic, constant payments and a constant interest rate.
     *
     * [Api set: ExcelApi 1.2]
     *
     * @param rate Is the interest rate per period. For example, use 6%/4 for quarterly payments at 6% APR.
     * @param nper Is the total number of payment periods in the investment.
     * @param pmt Is the payment made each period; it cannot change over the life of the investment.
     * @param pv Is the present value, or the lump-sum amount that a series of future payments is worth now. If omitted, Pv = 0.
     * @param type Is a value representing the timing of payment: payment at the beginning of the period = 1; payment at the end of the period = 0 or omitted.
     */
    fv(rate: number | Range | RangeReference | FunctionResult<any>, nper: number | Range | RangeReference | FunctionResult<any>, pmt: number | Range | RangeReference | FunctionResult<any>, pv?: number | Range | RangeReference | FunctionResult<any>, type?: number | Range | RangeReference | FunctionResult<any>): FunctionResult<number>;
    /**
     *
     * Returns the Gamma function value.
     *
     * [Api set: ExcelApi 1.2]
     *
     * @param x Is the value for which you want to calculate Gamma.
     */
    gamma(x: number | Range | RangeReference | FunctionResult<any>): FunctionResult<number>;
    /**
     *
     * Returns the natural logarithm of the gamma function.
     *
     * [Api set: ExcelApi 1.2]
     *
     * @param x Is the value for which you want to calculate GAMMALN, a positive number.
     */
    gammaLn(x: number | Range | RangeReference | FunctionResult<any>): FunctionResult<number>;
    /**
     *
     * Returns the natural logarithm of the gamma function.
     *
     * [Api set: ExcelApi 1.2]
     *
     * @param x Is the value for which you want to calculate GAMMALN.PRECISE, a positive number.
     */
    gammaLn_Precise(x: number | Range | RangeReference | FunctionResult<any>): FunctionResult<number>;
    /**
     *
     * Returns the gamma distribution.
     *
     * [Api set: ExcelApi 1.2]
     *
     * @param x Is the value at which you want to evaluate the distribution, a nonnegative number.
     * @param alpha Is a parameter to the distribution, a positive number.
     * @param beta Is a parameter to the distribution, a positive number. If beta = 1, GAMMA.DIST returns the standard gamma distribution.
     * @param cumulative Is a logical value: return the cumulative distribution function = TRUE; return the probability mass function = FALSE or omitted.
     */
    gamma_Dist(x: number | Range | RangeReference | FunctionResult<any>, alpha: number | Range | RangeReference | FunctionResult<any>, beta: number | Range | RangeReference | FunctionResult<any>, cumulative: boolean | Range | RangeReference | FunctionResult<any>): FunctionResult<number>;
    /**
     *
     * Returns the inverse of the gamma cumulative distribution: if p = GAMMA.DIST(x,...), then GAMMA.INV(p,...) = x.
     *
     * [Api set: ExcelApi 1.2]
     *
     * @param probability Is the probability associated with the gamma distribution, a number between 0 and 1, inclusive.
     * @param alpha Is a parameter to the distribution, a positive number.
     * @param beta Is a parameter to the distribution, a positive number. If beta = 1, GAMMA.INV returns the inverse of the standard gamma distribution.
     */
    gamma_Inv(probability: number | Range | RangeReference | FunctionResult<any>, alpha: number | Range | RangeReference | FunctionResult<any>, beta: number | Range | RangeReference | FunctionResult<any>): FunctionResult<number>;
    /**
     *
     * Returns 0.5 less than the standard normal cumulative distribution.
     *
     * [Api set: ExcelApi 1.2]
     *
     * @param x Is the value for which you want the distribution.
     */
    gauss(x: number | Range | RangeReference | FunctionResult<any>): FunctionResult<number>;
    /**
     *
     * Returns the greatest common divisor.
     *
     * [Api set: ExcelApi 1.2]
     *
     * @param values List of parameters, whose elements are 1 to 255 values.
     */
    gcd(...values: Array<number | string | Range | boolean | RangeReference | FunctionResult<any>>): FunctionResult<number>;
    /**
     *
     * Tests whether a number is greater than a threshold value.
     *
     * [Api set: ExcelApi 1.2]
     *
     * @param number Is the value to test against step.
     * @param step Is the threshold value.
     */
    geStep(number: number | string | boolean | Range | RangeReference | FunctionResult<any>, step?: number | string | boolean | Range | RangeReference | FunctionResult<any>): FunctionResult<number>;
    /**
     *
     * Returns the geometric mean of an array or range of positive numeric data.
     *
     * [Api set: ExcelApi 1.2]
     *
     * @param values List of parameters, whose elements are 1 to 255 numbers or names, arrays, or references that contain numbers for which you want the mean.
     */
    geoMean(...values: Array<number | Range | RangeReference | FunctionResult<any>>): FunctionResult<number>;
    /**
     *
     * Looks for a value in the top row of a table or array of values and returns the value in the same column from a row you specify.
     *
     * [Api set: ExcelApi 1.2]
     *
     * @param lookupValue Is the value to be found in the first row of the table and can be a value, a reference, or a text string.
     * @param tableArray Is a table of text, numbers, or logical values in which data is looked up. tableArray can be a reference to a range or a range name.
     * @param rowIndexNum Is the row number in tableArray from which the matching value should be returned. The first row of values in the table is row 1.
     * @param rangeLookup Is a logical value: to find the closest match in the top row (sorted in ascending order) = TRUE or omitted; find an exact match = FALSE.
     */
    hlookup(lookupValue: number | string | boolean | Range | RangeReference | FunctionResult<any>, tableArray: Range | number | RangeReference | FunctionResult<any>, rowIndexNum: Range | number | RangeReference | FunctionResult<any>, rangeLookup?: boolean | Range | RangeReference | FunctionResult<any>): FunctionResult<number | string | boolean>;
    /**
     *
     * Returns the harmonic mean of a data set of positive numbers: the reciprocal of the arithmetic mean of reciprocals.
     *
     * [Api set: ExcelApi 1.2]
     *
     * @param values List of parameters, whose elements are 1 to 255 numbers or names, arrays, or references that contain numbers for which you want the harmonic mean.
     */
    harMean(...values: Array<number | Range | RangeReference | FunctionResult<any>>): FunctionResult<number>;
    /**
     *
     * Converts a Hexadecimal number to binary.
     *
     * [Api set: ExcelApi 1.2]
     *
     * @param number Is the hexadecimal number you want to convert.
     * @param places Is the number of characters to use.
     */
    hex2Bin(number: number | string | boolean | Range | RangeReference | FunctionResult<any>, places?: number | string | boolean | Range | RangeReference | FunctionResult<any>): FunctionResult<number>;
    /**
     *
     * Converts a hexadecimal number to decimal.
     *
     * [Api set: ExcelApi 1.2]
     *
     * @param number Is the hexadecimal number you want to convert.
     */
    hex2Dec(number: number | string | boolean | Range | RangeReference | FunctionResult<any>): FunctionResult<number>;
    /**
     *
     * Converts a hexadecimal number to octal.
     *
     * [Api set: ExcelApi 1.2]
     *
     * @param number Is the hexadecimal number you want to convert.
     * @param places Is the number of characters to use.
     */
    hex2Oct(number: number | string | boolean | Range | RangeReference | FunctionResult<any>, places?: number | string | boolean | Range | RangeReference | FunctionResult<any>): FunctionResult<number>;
    /**
     *
     * Returns the hour as a number from 0 (12:00 A.M.) to 23 (11:00 P.M.).
     *
     * [Api set: ExcelApi 1.2]
     *
     * @param serialNumber Is a number in the date-time code used by Microsoft Excel, or text in time format, such as 16:48:00 or 4:48:00 PM.
     */
    hour(serialNumber: number | Range | RangeReference | FunctionResult<any>): FunctionResult<number>;
    /**
     *
     * Returns the hypergeometric distribution.
     *
     * [Api set: ExcelApi 1.2]
     *
     * @param sampleS Is the number of successes in the sample.
     * @param numberSample Is the size of the sample.
     * @param populationS Is the number of successes in the population.
     * @param numberPop Is the population size.
     * @param cumulative Is a logical value: for the cumulative distribution function, use TRUE; for the probability density function, use FALSE.
     */
    hypGeom_Dist(sampleS: number | Range | RangeReference | FunctionResult<any>, numberSample: number | Range | RangeReference | FunctionResult<any>, populationS: number | Range | RangeReference | FunctionResult<any>, numberPop: number | Range | RangeReference | FunctionResult<any>, cumulative: boolean | Range | RangeReference | FunctionResult<any>): FunctionResult<number>;
    /**
     *
     * Creates a shortcut or jump that opens a document stored on your hard drive, a network server, or on the Internet.
     *
     * [Api set: ExcelApi 1.2]
     *
     * @param linkLocation Is the text giving the path and file name to the document to be opened, a hard drive location, UNC address, or URL path.
     * @param friendlyName Is text or a number that is displayed in the cell. If omitted, the cell displays the linkLocation text.
     */
    hyperlink(linkLocation: string | Range | RangeReference | FunctionResult<any>, friendlyName?: number | string | boolean | Range | RangeReference | FunctionResult<any>): FunctionResult<number | string | boolean>;
    /**
     *
     * Rounds a number up, to the nearest integer or to the nearest multiple of significance.
     *
     * [Api set: ExcelApi 1.2]
     *
     * @param number Is the value you want to round.
     * @param significance Is the optional multiple to which you want to round.
     */
    iso_Ceiling(number: number | Range | RangeReference | FunctionResult<any>, significance?: number | Range | RangeReference | FunctionResult<any>): FunctionResult<number>;
    /**
     *
     * Checks whether a condition is met, and returns one value if TRUE, and another value if FALSE.
     *
     * [Api set: ExcelApi 1.2]
     *
     * @param logicalTest Is any value or expression that can be evaluated to TRUE or FALSE.
     * @param valueIfTrue Is the value that is returned if logicalTest is TRUE. If omitted, TRUE is returned. You can nest up to seven IF functions.
     * @param valueIfFalse Is the value that is returned if logicalTest is FALSE. If omitted, FALSE is returned.
     */
    if(logicalTest: boolean | Range | RangeReference | FunctionResult<any>, valueIfTrue?: Range | number | string | boolean | RangeReference | FunctionResult<any>, valueIfFalse?: Range | number | string | boolean | RangeReference | FunctionResult<any>): FunctionResult<number | string | boolean>;
    /**
     *
     * Returns the absolute value (modulus) of a complex number.
     *
     * [Api set: ExcelApi 1.2]
     *
     * @param inumber Is a complex number for which you want the absolute value.
     */
    imAbs(inumber: number | string | boolean | Range | RangeReference | FunctionResult<any>): FunctionResult<number>;
    /**
     *
     * Returns the argument q, an angle expressed in radians.
     *
     * [Api set: ExcelApi 1.2]
     *
     * @param inumber Is a complex number for which you want the argument.
     */
    imArgument(inumber: number | string | boolean | Range | RangeReference | FunctionResult<any>): FunctionResult<number>;
    /**
     *
     * Returns the complex conjugate of a complex number.
     *
     * [Api set: ExcelApi 1.2]
     *
     * @param inumber Is a complex number for which you want the conjugate.
     */
    imConjugate(inumber: number | string | boolean | Range | RangeReference | FunctionResult<any>): FunctionResult<number>;
    /**
     *
     * Returns the cosine of a complex number.
     *
     * [Api set: ExcelApi 1.2]
     *
     * @param inumber Is a complex number for which you want the cosine.
     */
    imCos(inumber: number | string | boolean | Range | RangeReference | FunctionResult<any>): FunctionResult<number>;
    /**
     *
     * Returns the hyperbolic cosine of a complex number.
     *
     * [Api set: ExcelApi 1.2]
     *
     * @param inumber Is a complex number for which you want the hyperbolic cosine.
     */
    imCosh(inumber: number | string | boolean | Range | RangeReference | FunctionResult<any>): FunctionResult<number>;
    /**
     *
     * Returns the cotangent of a complex number.
     *
     * [Api set: ExcelApi 1.2]
     *
     * @param inumber Is a complex number for which you want the cotangent.
     */
    imCot(inumber: number | string | boolean | Range | RangeReference | FunctionResult<any>): FunctionResult<number>;
    /**
     *
     * Returns the cosecant of a complex number.
     *
     * [Api set: ExcelApi 1.2]
     *
     * @param inumber Is a complex number for which you want the cosecant.
     */
    imCsc(inumber: number | string | boolean | Range | RangeReference | FunctionResult<any>): FunctionResult<number>;
    /**
     *
     * Returns the hyperbolic cosecant of a complex number.
     *
     * [Api set: ExcelApi 1.2]
     *
     * @param inumber Is a complex number for which you want the hyperbolic cosecant.
     */
    imCsch(inumber: number | string | boolean | Range | RangeReference | FunctionResult<any>): FunctionResult<number>;
    /**
     *
     * Returns the quotient of two complex numbers.
     *
     * [Api set: ExcelApi 1.2]
     *
     * @param inumber1 Is the complex numerator or dividend.
     * @param inumber2 Is the complex denominator or divisor.
     */
    imDiv(inumber1: number | string | boolean | Range | RangeReference | FunctionResult<any>, inumber2: number | string | boolean | Range | RangeReference | FunctionResult<any>): FunctionResult<number>;
    /**
     *
     * Returns the exponential of a complex number.
     *
     * [Api set: ExcelApi 1.2]
     *
     * @param inumber Is a complex number for which you want the exponential.
     */
    imExp(inumber: number | string | boolean | Range | RangeReference | FunctionResult<any>): FunctionResult<number>;
    /**
     *
     * Returns the natural logarithm of a complex number.
     *
     * [Api set: ExcelApi 1.2]
     *
     * @param inumber Is a complex number for which you want the natural logarithm.
     */
    imLn(inumber: number | string | boolean | Range | RangeReference | FunctionResult<any>): FunctionResult<number>;
    /**
     *
     * Returns the base-10 logarithm of a complex number.
     *
     * [Api set: ExcelApi 1.2]
     *
     * @param inumber Is a complex number for which you want the common logarithm.
     */
    imLog10(inumber: number | string | boolean | Range | RangeReference | FunctionResult<any>): FunctionResult<number>;
    /**
     *
     * Returns the base-2 logarithm of a complex number.
     *
     * [Api set: ExcelApi 1.2]
     *
     * @param inumber Is a complex number for which you want the base-2 logarithm.
     */
    imLog2(inumber: number | string | boolean | Range | RangeReference | FunctionResult<any>): FunctionResult<number>;
    /**
     *
     * Returns a complex number raised to an integer power.
     *
     * [Api set: ExcelApi 1.2]
     *
     * @param inumber Is a complex number you want to raise to a power.
     * @param number Is the power to which you want to raise the complex number.
     */
    imPower(inumber: number | string | boolean | Range | RangeReference | FunctionResult<any>, number: number | string | boolean | Range | RangeReference | FunctionResult<any>): FunctionResult<number>;
    /**
     *
     * Returns the product of 1 to 255 complex numbers.
     *
     * [Api set: ExcelApi 1.2]
     *
     * @param values Inumber1, Inumber2,... are from 1 to 255 complex numbers to multiply.
     */
    imProduct(...values: Array<Range | number | string | boolean | RangeReference | FunctionResult<any>>): FunctionResult<number>;
    /**
     *
     * Returns the real coefficient of a complex number.
     *
     * [Api set: ExcelApi 1.2]
     *
     * @param inumber Is a complex number for which you want the real coefficient.
     */
    imReal(inumber: number | string | boolean | Range | RangeReference | FunctionResult<any>): FunctionResult<number>;
    /**
     *
     * Returns the secant of a complex number.
     *
     * [Api set: ExcelApi 1.2]
     *
     * @param inumber Is a complex number for which you want the secant.
     */
    imSec(inumber: number | string | boolean | Range | RangeReference | FunctionResult<any>): FunctionResult<number>;
    /**
     *
     * Returns the hyperbolic secant of a complex number.
     *
     * [Api set: ExcelApi 1.2]
     *
     * @param inumber Is a complex number for which you want the hyperbolic secant.
     */
    imSech(inumber: number | string | boolean | Range | RangeReference | FunctionResult<any>): FunctionResult<number>;
    /**
     *
     * Returns the sine of a complex number.
     *
     * [Api set: ExcelApi 1.2]
     *
     * @param inumber Is a complex number for which you want the sine.
     */
    imSin(inumber: number | string | boolean | Range | RangeReference | FunctionResult<any>): FunctionResult<number>;
    /**
     *
     * Returns the hyperbolic sine of a complex number.
     *
     * [Api set: ExcelApi 1.2]
     *
     * @param inumber Is a complex number for which you want the hyperbolic sine.
     */
    imSinh(inumber: number | string | boolean | Range | RangeReference | FunctionResult<any>): FunctionResult<number>;
    /**
     *
     * Returns the square root of a complex number.
     *
     * [Api set: ExcelApi 1.2]
     *
     * @param inumber Is a complex number for which you want the square root.
     */
    imSqrt(inumber: number | string | boolean | Range | RangeReference | FunctionResult<any>): FunctionResult<number>;
    /**
     *
     * Returns the difference of two complex numbers.
     *
     * [Api set: ExcelApi 1.2]
     *
     * @param inumber1 Is the complex number from which to subtract inumber2.
     * @param inumber2 Is the complex number to subtract from inumber1.
     */
    imSub(inumber1: number | string | boolean | Range | RangeReference | FunctionResult<any>, inumber2: number | string | boolean | Range | RangeReference | FunctionResult<any>): FunctionResult<number>;
    /**
     *
     * Returns the sum of complex numbers.
     *
     * [Api set: ExcelApi 1.2]
     *
     * @param values List of parameters, whose elements are from 1 to 255 complex numbers to add.
     */
    imSum(...values: Array<Range | number | string | boolean | RangeReference | FunctionResult<any>>): FunctionResult<number>;
    /**
     *
     * Returns the tangent of a complex number.
     *
     * [Api set: ExcelApi 1.2]
     *
     * @param inumber Is a complex number for which you want the tangent.
     */
    imTan(inumber: number | string | boolean | Range | RangeReference | FunctionResult<any>): FunctionResult<number>;
    /**
     *
     * Returns the imaginary coefficient of a complex number.
     *
     * [Api set: ExcelApi 1.2]
     *
     * @param inumber Is a complex number for which you want the imaginary coefficient.
     */
    imaginary(inumber: number | string | boolean | Range | RangeReference | FunctionResult<any>): FunctionResult<number>;
    /**
     *
     * Rounds a number down to the nearest integer.
     *
     * [Api set: ExcelApi 1.2]
     *
     * @param number Is the real number you want to round down to an integer.
     */
    int(number: number | Range | RangeReference | FunctionResult<any>): FunctionResult<number>;
    /**
     *
     * Returns the interest rate for a fully invested security.
     *
     * [Api set: ExcelApi 1.2]
     *
     * @param settlement Is the security's settlement date, expressed as a serial date number.
     * @param maturity Is the security's maturity date, expressed as a serial date number.
     * @param investment Is the amount invested in the security.
     * @param redemption Is the amount to be received at maturity.
     * @param basis Is the type of day count basis to use.
     */
    intRate(settlement: number | string | boolean | Range | RangeReference | FunctionResult<any>, maturity: number | string | boolean | Range | RangeReference | FunctionResult<any>, investment: number | string | boolean | Range | RangeReference | FunctionResult<any>, redemption: number | string | boolean | Range | RangeReference | FunctionResult<any>, basis?: number | string | boolean | Range | RangeReference | FunctionResult<any>): FunctionResult<number>;
    /**
     *
     * Returns the interest payment for a given period for an investment, based on periodic, constant payments and a constant interest rate.
     *
     * [Api set: ExcelApi 1.2]
     *
     * @param rate Is the interest rate per period. For example, use 6%/4 for quarterly payments at 6% APR.
     * @param per Is the period for which you want to find the interest and must be in the range 1 to Nper.
     * @param nper Is the total number of payment periods in an investment.
     * @param pv Is the present value, or the lump-sum amount that a series of future payments is worth now.
     * @param fv Is the future value, or a cash balance you want to attain after the last payment is made. If omitted, Fv = 0.
     * @param type Is a logical value representing the timing of payment: at the end of the period = 0 or omitted, at the beginning of the period = 1.
     */
    ipmt(rate: number | Range | RangeReference | FunctionResult<any>, per: number | Range | RangeReference | FunctionResult<any>, nper: number | Range | RangeReference | FunctionResult<any>, pv: number | Range | RangeReference | FunctionResult<any>, fv?: number | Range | RangeReference | FunctionResult<any>, type?: number | Range | RangeReference | FunctionResult<any>): FunctionResult<number>;
    /**
     *
     * Returns the internal rate of return for a series of cash flows.
     *
     * [Api set: ExcelApi 1.2]
     *
     * @param values Is an array or a reference to cells that contain numbers for which you want to calculate the internal rate of return.
     * @param guess Is a number that you guess is close to the result of IRR; 0.1 (10 percent) if omitted.
     */
    irr(values: Range | RangeReference | FunctionResult<any>, guess?: number | Range | RangeReference | FunctionResult<any>): FunctionResult<number>;
    /**
     *
     * Checks whether a value is an error (#VALUE!, #REF!, #DIV/0!, #NUM!, #NAME?, or #NULL!) excluding #N/A, and returns TRUE or FALSE.
     *
     * [Api set: ExcelApi 1.2]
     *
     * @param value Is the value you want to test. Value can refer to a cell, a formula, or a name that refers to a cell, formula, or value.
     */
    isErr(value: number | string | boolean | Range | RangeReference | FunctionResult<any>): FunctionResult<boolean>;
    /**
     *
     * Checks whether a value is an error (#N/A, #VALUE!, #REF!, #DIV/0!, #NUM!, #NAME?, or #NULL!), and returns TRUE or FALSE.
     *
     * [Api set: ExcelApi 1.2]
     *
     * @param value Is the value you want to test. Value can refer to a cell, a formula, or a name that refers to a cell, formula, or value.
     */
    isError(value: number | string | boolean | Range | RangeReference | FunctionResult<any>): FunctionResult<boolean>;
    /**
     *
     * Returns TRUE if the number is even.
     *
     * [Api set: ExcelApi 1.2]
     *
     * @param number Is the value to test.
     */
    isEven(number: number | string | boolean | Range | RangeReference | FunctionResult<any>): FunctionResult<number>;
    /**
     *
     * Checks whether a reference is to a cell containing a formula, and returns TRUE or FALSE.
     *
     * [Api set: ExcelApi 1.2]
     *
     * @param reference Is a reference to the cell you want to test.  Reference can be a cell reference, a formula, or name that refers to a cell.
     */
    isFormula(reference: Range | RangeReference | FunctionResult<any>): FunctionResult<boolean>;
    /**
     *
     * Checks whether a value is a logical value (TRUE or FALSE), and returns TRUE or FALSE.
     *
     * [Api set: ExcelApi 1.2]
     *
     * @param value Is the value you want to test. Value can refer to a cell, a formula, or a name that refers to a cell, formula, or value.
     */
    isLogical(value: number | string | boolean | Range | RangeReference | FunctionResult<any>): FunctionResult<boolean>;
    /**
     *
     * Checks whether a value is #N/A, and returns TRUE or FALSE.
     *
     * [Api set: ExcelApi 1.2]
     *
     * @param value Is the value you want to test. Value can refer to a cell, a formula, or a name that refers to a cell, formula, or value.
     */
    isNA(value: number | string | boolean | Range | RangeReference | FunctionResult<any>): FunctionResult<boolean>;
    /**
     *
     * Checks whether a value is not text (blank cells are not text), and returns TRUE or FALSE.
     *
     * [Api set: ExcelApi 1.2]
     *
     * @param value Is the value you want tested: a cell; a formula; or a name referring to a cell, formula, or value.
     */
    isNonText(value: number | string | boolean | Range | RangeReference | FunctionResult<any>): FunctionResult<boolean>;
    /**
     *
     * Checks whether a value is a number, and returns TRUE or FALSE.
     *
     * [Api set: ExcelApi 1.2]
     *
     * @param value Is the value you want to test. Value can refer to a cell, a formula, or a name that refers to a cell, formula, or value.
     */
    isNumber(value: number | string | boolean | Range | RangeReference | FunctionResult<any>): FunctionResult<boolean>;
    /**
     *
     * Returns TRUE if the number is odd.
     *
     * [Api set: ExcelApi 1.2]
     *
     * @param number Is the value to test.
     */
    isOdd(number: number | string | boolean | Range | RangeReference | FunctionResult<any>): FunctionResult<number>;
    /**
     *
     * Checks whether a value is text, and returns TRUE or FALSE.
     *
     * [Api set: ExcelApi 1.2]
     *
     * @param value Is the value you want to test. Value can refer to a cell, a formula, or a name that refers to a cell, formula, or value.
     */
    isText(value: number | string | boolean | Range | RangeReference | FunctionResult<any>): FunctionResult<boolean>;
    /**
     *
     * Returns the ISO week number in the year for a given date.
     *
     * [Api set: ExcelApi 1.2]
     *
     * @param date Is the date-time code used by Microsoft Excel for date and time calculation.
     */
    isoWeekNum(date: number | Range | RangeReference | FunctionResult<any>): FunctionResult<number>;
    /**
     *
     * Returns the interest paid during a specific period of an investment.
     *
     * [Api set: ExcelApi 1.2]
     *
     * @param rate Interest rate per period. For example, use 6%/4 for quarterly payments at 6% APR.
     * @param per Period for which you want to find the interest.
     * @param nper Number of payment periods in an investment.
     * @param pv Lump sum amount that a series of future payments is right now.
     */
    ispmt(rate: number | Range | RangeReference | FunctionResult<any>, per: number | Range | RangeReference | FunctionResult<any>, nper: number | Range | RangeReference | FunctionResult<any>, pv: number | Range | RangeReference | FunctionResult<any>): FunctionResult<number>;
    /**
     *
     * Checks whether a value is a reference, and returns TRUE or FALSE.
     *
     * [Api set: ExcelApi 1.2]
     *
     * @param value Is the value you want to test. Value can refer to a cell, a formula, or a name that refers to a cell, formula, or value.
     */
    isref(value: Range | number | string | boolean | RangeReference | FunctionResult<any>): FunctionResult<boolean>;
    /**
     *
     * Returns the kurtosis of a data set.
     *
     * [Api set: ExcelApi 1.2]
     *
     * @param values List of parameters, whose elements are 1 to 255 numbers or names, arrays, or references that contain numbers for which you want the kurtosis.
     */
    kurt(...values: Array<number | Range | RangeReference | FunctionResult<any>>): FunctionResult<number>;
    /**
     *
     * Returns the k-th largest value in a data set. For example, the fifth largest number.
     *
     * [Api set: ExcelApi 1.2]
     *
     * @param array Is the array or range of data for which you want to determine the k-th largest value.
     * @param k Is the position (from the largest) in the array or cell range of the value to return.
     */
    large(array: number | Range | RangeReference | FunctionResult<any>, k: number | Range | RangeReference | FunctionResult<any>): FunctionResult<number>;
    /**
     *
     * Returns the least common multiple.
     *
     * [Api set: ExcelApi 1.2]
     *
     * @param values List of parameters, whose elements are 1 to 255 values for which you want the least common multiple.
     */
    lcm(...values: Array<number | string | Range | boolean | RangeReference | FunctionResult<any>>): FunctionResult<number>;
    /**
     *
     * Returns the specified number of characters from the start of a text string.
     *
     * [Api set: ExcelApi 1.2]
     *
     * @param text Is the text string containing the characters you want to extract.
     * @param numChars Specifies how many characters you want LEFT to extract; 1 if omitted.
     */
    left(text: string | Range | RangeReference | FunctionResult<any>, numChars?: number | Range | RangeReference | FunctionResult<any>): FunctionResult<string>;
    /**
     *
     * Returns the specified number of characters from the start of a text string. Use with double-byte character sets (DBCS).
     *
     * [Api set: ExcelApi 1.2]
     *
     * @param text Is the text string containing the characters you want to extract.
     * @param numBytes Specifies how many characters you want LEFT to return.
     */
    leftb(text: string | Range | RangeReference | FunctionResult<any>, numBytes?: number | Range | RangeReference | FunctionResult<any>): FunctionResult<string>;
    /**
     *
     * Returns the number of characters in a text string.
     *
     * [Api set: ExcelApi 1.2]
     *
     * @param text Is the text whose length you want to find. Spaces count as characters.
     */
    len(text: string | Range | RangeReference | FunctionResult<any>): FunctionResult<number>;
    /**
     *
     * Returns the number of characters in a text string. Use with double-byte character sets (DBCS).
     *
     * [Api set: ExcelApi 1.2]
     *
     * @param text Is the text whose length you want to find.
     */
    lenb(text: string | Range | RangeReference | FunctionResult<any>): FunctionResult<number>;
    /**
     *
     * Returns the natural logarithm of a number.
     *
     * [Api set: ExcelApi 1.2]
     *
     * @param number Is the positive real number for which you want the natural logarithm.
     */
    ln(number: number | Range | RangeReference | FunctionResult<any>): FunctionResult<number>;
    /**
     *
     * Returns the logarithm of a number to the base you specify.
     *
     * [Api set: ExcelApi 1.2]
     *
     * @param number Is the positive real number for which you want the logarithm.
     * @param base Is the base of the logarithm; 10 if omitted.
     */
    log(number: number | Range | RangeReference | FunctionResult<any>, base?: number | Range | RangeReference | FunctionResult<any>): FunctionResult<number>;
    /**
     *
     * Returns the base-10 logarithm of a number.
     *
     * [Api set: ExcelApi 1.2]
     *
     * @param number Is the positive real number for which you want the base-10 logarithm.
     */
    log10(number: number | Range | RangeReference | FunctionResult<any>): FunctionResult<number>;
    /**
     *
     * Returns the lognormal distribution of x, where ln(x) is normally distributed with parameters Mean and Standard_dev.
     *
     * [Api set: ExcelApi 1.2]
     *
     * @param x Is the value at which to evaluate the function, a positive number.
     * @param mean Is the mean of ln(x).
     * @param standardDev Is the standard deviation of ln(x), a positive number.
     * @param cumulative Is a logical value: for the cumulative distribution function, use TRUE; for the probability density function, use FALSE.
     */
    logNorm_Dist(x: number | Range | RangeReference | FunctionResult<any>, mean: number | Range | RangeReference | FunctionResult<any>, standardDev: number | Range | RangeReference | FunctionResult<any>, cumulative: boolean | Range | RangeReference | FunctionResult<any>): FunctionResult<number>;
    /**
     *
     * Returns the inverse of the lognormal cumulative distribution function of x, where ln(x) is normally distributed with parameters Mean and Standard_dev.
     *
     * [Api set: ExcelApi 1.2]
     *
     * @param probability Is a probability associated with the lognormal distribution, a number between 0 and 1, inclusive.
     * @param mean Is the mean of ln(x).
     * @param standardDev Is the standard deviation of ln(x), a positive number.
     */
    logNorm_Inv(probability: number | Range | RangeReference | FunctionResult<any>, mean: number | Range | RangeReference | FunctionResult<any>, standardDev: number | Range | RangeReference | FunctionResult<any>): FunctionResult<number>;
    /**
     *
     * Looks up a value either from a one-row or one-column range or from an array. Provided for backward compatibility.
     *
     * [Api set: ExcelApi 1.2]
     *
     * @param lookupValue Is a value that LOOKUP searches for in lookupVector and can be a number, text, a logical value, or a name or reference to a value.
     * @param lookupVector Is a range that contains only one row or one column of text, numbers, or logical values, placed in ascending order.
     * @param resultVector Is a range that contains only one row or column, the same size as lookupVector.
     */
    lookup(lookupValue: number | string | boolean | Range | RangeReference | FunctionResult<any>, lookupVector: Range | RangeReference | FunctionResult<any>, resultVector?: Range | RangeReference | FunctionResult<any>): FunctionResult<number | string | boolean>;
    /**
     *
     * Converts all letters in a text string to lowercase.
     *
     * [Api set: ExcelApi 1.2]
     *
     * @param text Is the text you want to convert to lowercase. Characters in Text that are not letters are not changed.
     */
    lower(text: string | Range | RangeReference | FunctionResult<any>): FunctionResult<string>;
    /**
     *
     * Returns the Macauley modified duration for a security with an assumed par value of $100.
     *
     * [Api set: ExcelApi 1.2]
     *
     * @param settlement Is the security's settlement date, expressed as a serial date number.
     * @param maturity Is the security's maturity date, expressed as a serial date number.
     * @param coupon Is the security's annual coupon rate.
     * @param yld Is the security's annual yield.
     * @param frequency Is the number of coupon payments per year.
     * @param basis Is the type of day count basis to use.
     */
    mduration(settlement: number | string | boolean | Range | RangeReference | FunctionResult<any>, maturity: number | string | boolean | Range | RangeReference | FunctionResult<any>, coupon: number | string | boolean | Range | RangeReference | FunctionResult<any>, yld: number | string | boolean | Range | RangeReference | FunctionResult<any>, frequency: number | string | boolean | Range | RangeReference | FunctionResult<any>, basis?: number | string | boolean | Range | RangeReference | FunctionResult<any>): FunctionResult<number>;
    /**
     *
     * Returns the internal rate of return for a series of periodic cash flows, considering both cost of investment and interest on reinvestment of cash.
     *
     * [Api set: ExcelApi 1.2]
     *
     * @param values Is an array or a reference to cells that contain numbers that represent a series of payments (negative) and income (positive) at regular periods.
     * @param financeRate Is the interest rate you pay on the money used in the cash flows.
     * @param reinvestRate Is the interest rate you receive on the cash flows as you reinvest them.
     */
    mirr(values: Range | RangeReference | FunctionResult<any>, financeRate: number | Range | RangeReference | FunctionResult<any>, reinvestRate: number | Range | RangeReference | FunctionResult<any>): FunctionResult<number>;
    /**
     *
     * Returns a number rounded to the desired multiple.
     *
     * [Api set: ExcelApi 1.2]
     *
     * @param number Is the value to round.
     * @param multiple Is the multiple to which you want to round number.
     */
    mround(number: number | string | boolean | Range | RangeReference | FunctionResult<any>, multiple: number | string | boolean | Range | RangeReference | FunctionResult<any>): FunctionResult<number>;
    /**
     *
     * Returns the relative position of an item in an array that matches a specified value in a specified order.
     *
     * [Api set: ExcelApi 1.2]
     *
     * @param lookupValue Is the value you use to find the value you want in the array, a number, text, or logical value, or a reference to one of these.
     * @param lookupArray Is a contiguous range of cells containing possible lookup values, an array of values, or a reference to an array.
     * @param matchType Is a number 1, 0, or -1 indicating which value to return.
     */
    match(lookupValue: number | string | boolean | Range | RangeReference | FunctionResult<any>, lookupArray: number | Range | RangeReference | FunctionResult<any>, matchType?: number | Range | RangeReference | FunctionResult<any>): FunctionResult<number>;
    /**
     *
     * Returns the largest value in a set of values. Ignores logical values and text.
     *
     * [Api set: ExcelApi 1.2]
     *
     * @param values List of parameters, whose elements are 1 to 255 numbers, empty cells, logical values, or text numbers for which you want the maximum.
     */
    max(...values: Array<number | Range | RangeReference | FunctionResult<any>>): FunctionResult<number>;
    /**
     *
     * Returns the largest value in a set of values. Does not ignore logical values and text.
     *
     * [Api set: ExcelApi 1.2]
     *
     * @param values List of parameters, whose elements are 1 to 255 numbers, empty cells, logical values, or text numbers for which you want the maximum.
     */
    maxA(...values: Array<number | Range | RangeReference | FunctionResult<any>>): FunctionResult<number>;
    /**
     *
     * Returns the median, or the number in the middle of the set of given numbers.
     *
     * [Api set: ExcelApi 1.2]
     *
     * @param values List of parameters, whose elements are 1 to 255 numbers or names, arrays, or references that contain numbers for which you want the median.
     */
    median(...values: Array<number | Range | RangeReference | FunctionResult<any>>): FunctionResult<number>;
    /**
     *
     * Returns the characters from the middle of a text string, given a starting position and length.
     *
     * [Api set: ExcelApi 1.2]
     *
     * @param text Is the text string from which you want to extract the characters.
     * @param startNum Is the position of the first character you want to extract. The first character in Text is 1.
     * @param numChars Specifies how many characters to return from Text.
     */
    mid(text: string | Range | RangeReference | FunctionResult<any>, startNum: number | Range | RangeReference | FunctionResult<any>, numChars: number | Range | RangeReference | FunctionResult<any>): FunctionResult<string>;
    /**
     *
     * Returns characters from the middle of a text string, given a starting position and length. Use with double-byte character sets (DBCS).
     *
     * [Api set: ExcelApi 1.2]
     *
     * @param text Is the text string containing the characters you want to extract.
     * @param startNum Is the position of the first character you want to extract in text.
     * @param numBytes Specifies how many characters to return from text.
     */
    midb(text: string | Range | RangeReference | FunctionResult<any>, startNum: number | Range | RangeReference | FunctionResult<any>, numBytes: number | Range | RangeReference | FunctionResult<any>): FunctionResult<string>;
    /**
     *
     * Returns the smallest number in a set of values. Ignores logical values and text.
     *
     * [Api set: ExcelApi 1.2]
     *
     * @param values List of parameters, whose elements are 1 to 255 numbers, empty cells, logical values, or text numbers for which you want the minimum.
     */
    min(...values: Array<number | Range | RangeReference | FunctionResult<any>>): FunctionResult<number>;
    /**
     *
     * Returns the smallest value in a set of values. Does not ignore logical values and text.
     *
     * [Api set: ExcelApi 1.2]
     *
     * @param values List of parameters, whose elements are 1 to 255 numbers, empty cells, logical values, or text numbers for which you want the minimum.
     */
    minA(...values: Array<number | Range | RangeReference | FunctionResult<any>>): FunctionResult<number>;
    /**
     *
     * Returns the minute, a number from 0 to 59.
     *
     * [Api set: ExcelApi 1.2]
     *
     * @param serialNumber Is a number in the date-time code used by Microsoft Excel or text in time format, such as 16:48:00 or 4:48:00 PM.
     */
    minute(serialNumber: number | Range | RangeReference | FunctionResult<any>): FunctionResult<number>;
    /**
     *
     * Returns the remainder after a number is divided by a divisor.
     *
     * [Api set: ExcelApi 1.2]
     *
     * @param number Is the number for which you want to find the remainder after the division is performed.
     * @param divisor Is the number by which you want to divide Number.
     */
    mod(number: number | Range | RangeReference | FunctionResult<any>, divisor: number | Range | RangeReference | FunctionResult<any>): FunctionResult<number>;
    /**
     *
     * Returns the month, a number from 1 (January) to 12 (December).
     *
     * [Api set: ExcelApi 1.2]
     *
     * @param serialNumber Is a number in the date-time code used by Microsoft
     */
    month(serialNumber: number | Range | RangeReference | FunctionResult<any>): FunctionResult<number>;
    /**
     *
     * Returns the multinomial of a set of numbers.
     *
     * [Api set: ExcelApi 1.2]
     *
     * @param values List of parameters, whose elements are 1 to 255 values for which you want the multinomial.
     */
    multiNomial(...values: Array<number | string | Range | boolean | RangeReference | FunctionResult<any>>): FunctionResult<number>;
    /**
     *
     * Converts non-number value to a number, dates to serial numbers, TRUE to 1, anything else to 0 (zero).
     *
     * [Api set: ExcelApi 1.2]
     *
     * @param value Is the value you want converted.
     */
    n(value: number | string | boolean | Range | RangeReference | FunctionResult<any>): FunctionResult<number>;
    /**
     *
     * Returns the number of periods for an investment based on periodic, constant payments and a constant interest rate.
     *
     * [Api set: ExcelApi 1.2]
     *
     * @param rate Is the interest rate per period. For example, use 6%/4 for quarterly payments at 6% APR.
     * @param pmt Is the payment made each period; it cannot change over the life of the investment.
     * @param pv Is the present value, or the lump-sum amount that a series of future payments is worth now.
     * @param fv Is the future value, or a cash balance you want to attain after the last payment is made. If omitted, zero is used.
     * @param type Is a logical value: payment at the beginning of the period = 1; payment at the end of the period = 0 or omitted.
     */
    nper(rate: number | Range | RangeReference | FunctionResult<any>, pmt: number | Range | RangeReference | FunctionResult<any>, pv: number | Range | RangeReference | FunctionResult<any>, fv?: number | Range | RangeReference | FunctionResult<any>, type?: number | Range | RangeReference | FunctionResult<any>): FunctionResult<number>;
    /**
     *
     * Returns the error value #N/A (value not available).
     *
     * [Api set: ExcelApi 1.2]
     */
    na(): FunctionResult<number | string>;
    /**
     *
     * Returns the negative binomial distribution, the probability that there will be Number_f failures before the Number_s-th success, with Probability_s probability of a success.
     *
     * [Api set: ExcelApi 1.2]
     *
     * @param numberF Is the number of failures.
     * @param numberS Is the threshold number of successes.
     * @param probabilityS Is the probability of a success; a number between 0 and 1.
     * @param cumulative Is a logical value: for the cumulative distribution function, use TRUE; for the probability mass function, use FALSE.
     */
    negBinom_Dist(numberF: number | Range | RangeReference | FunctionResult<any>, numberS: number | Range | RangeReference | FunctionResult<any>, probabilityS: number | Range | RangeReference | FunctionResult<any>, cumulative: boolean | Range | RangeReference | FunctionResult<any>): FunctionResult<number>;
    /**
     *
     * Returns the number of whole workdays between two dates.
     *
     * [Api set: ExcelApi 1.2]
     *
     * @param startDate Is a serial date number that represents the start date.
     * @param endDate Is a serial date number that represents the end date.
     * @param holidays Is an optional set of one or more serial date numbers to exclude from the working calendar, such as state and federal holidays and floating holidays.
     */
    networkDays(startDate: number | string | boolean | Range | RangeReference | FunctionResult<any>, endDate: number | string | boolean | Range | RangeReference | FunctionResult<any>, holidays?: number | string | Range | boolean | RangeReference | FunctionResult<any>): FunctionResult<number>;
    /**
     *
     * Returns the number of whole workdays between two dates with custom weekend parameters.
     *
     * [Api set: ExcelApi 1.2]
     *
     * @param startDate Is a serial date number that represents the start date.
     * @param endDate Is a serial date number that represents the end date.
     * @param weekend Is a number or string specifying when weekends occur.
     * @param holidays Is an optional set of one or more serial date numbers to exclude from the working calendar, such as state and federal holidays and floating holidays.
     */
    networkDays_Intl(startDate: number | string | boolean | Range | RangeReference | FunctionResult<any>, endDate: number | string | boolean | Range | RangeReference | FunctionResult<any>, weekend?: number | string | Range | RangeReference | FunctionResult<any>, holidays?: number | string | Range | boolean | RangeReference | FunctionResult<any>): FunctionResult<number>;
    /**
     *
     * Returns the annual nominal interest rate.
     *
     * [Api set: ExcelApi 1.2]
     *
     * @param effectRate Is the effective interest rate.
     * @param npery Is the number of compounding periods per year.
     */
    nominal(effectRate: number | string | boolean | Range | RangeReference | FunctionResult<any>, npery: number | string | boolean | Range | RangeReference | FunctionResult<any>): FunctionResult<number>;
    /**
     *
     * Returns the normal distribution for the specified mean and standard deviation.
     *
     * [Api set: ExcelApi 1.2]
     *
     * @param x Is the value for which you want the distribution.
     * @param mean Is the arithmetic mean of the distribution.
     * @param standardDev Is the standard deviation of the distribution, a positive number.
     * @param cumulative Is a logical value: for the cumulative distribution function, use TRUE; for the probability density function, use FALSE.
     */
    norm_Dist(x: number | Range | RangeReference | FunctionResult<any>, mean: number | Range | RangeReference | FunctionResult<any>, standardDev: number | Range | RangeReference | FunctionResult<any>, cumulative: boolean | Range | RangeReference | FunctionResult<any>): FunctionResult<number>;
    /**
     *
     * Returns the inverse of the normal cumulative distribution for the specified mean and standard deviation.
     *
     * [Api set: ExcelApi 1.2]
     *
     * @param probability Is a probability corresponding to the normal distribution, a number between 0 and 1 inclusive.
     * @param mean Is the arithmetic mean of the distribution.
     * @param standardDev Is the standard deviation of the distribution, a positive number.
     */
    norm_Inv(probability: number | Range | RangeReference | FunctionResult<any>, mean: number | Range | RangeReference | FunctionResult<any>, standardDev: number | Range | RangeReference | FunctionResult<any>): FunctionResult<number>;
    /**
     *
     * Returns the standard normal distribution (has a mean of zero and a standard deviation of one).
     *
     * [Api set: ExcelApi 1.2]
     *
     * @param z Is the value for which you want the distribution.
     * @param cumulative Is a logical value for the function to return: the cumulative distribution function = TRUE; the probability density function = FALSE.
     */
    norm_S_Dist(z: number | Range | RangeReference | FunctionResult<any>, cumulative: boolean | Range | RangeReference | FunctionResult<any>): FunctionResult<number>;
    /**
     *
     * Returns the inverse of the standard normal cumulative distribution (has a mean of zero and a standard deviation of one).
     *
     * [Api set: ExcelApi 1.2]
     *
     * @param probability Is a probability corresponding to the normal distribution, a number between 0 and 1 inclusive.
     */
    norm_S_Inv(probability: number | Range | RangeReference | FunctionResult<any>): FunctionResult<number>;
    /**
     *
     * Changes FALSE to TRUE, or TRUE to FALSE.
     *
     * [Api set: ExcelApi 1.2]
     *
     * @param logical Is a value or expression that can be evaluated to TRUE or FALSE.
     */
    not(logical: boolean | Range | RangeReference | FunctionResult<any>): FunctionResult<boolean>;
    /**
     *
     * Returns the current date and time formatted as a date and time.
     *
     * [Api set: ExcelApi 1.2]
     */
    now(): FunctionResult<number>;
    /**
     *
     * Returns the net present value of an investment based on a discount rate and a series of future payments (negative values) and income (positive values).
     *
     * [Api set: ExcelApi 1.2]
     *
     * @param rate Is the rate of discount over the length of one period.
     * @param values List of parameters, whose elements are 1 to 254 payments and income, equally spaced in time and occurring at the end of each period.
     */
    npv(rate: number | Range | RangeReference | FunctionResult<any>, ...values: Array<number | Range | RangeReference | FunctionResult<any>>): FunctionResult<number>;
    /**
     *
     * Converts text to number in a locale-independent manner.
     *
     * [Api set: ExcelApi 1.2]
     *
     * @param text Is the string representing the number you want to convert.
     * @param decimalSeparator Is the character used as the decimal separator in the string.
     * @param groupSeparator Is the character used as the group separator in the string.
     */
    numberValue(text: string | Range | RangeReference | FunctionResult<any>, decimalSeparator?: string | Range | RangeReference | FunctionResult<any>, groupSeparator?: string | Range | RangeReference | FunctionResult<any>): FunctionResult<number>;
    /**
     *
     * Converts an octal number to binary.
     *
     * [Api set: ExcelApi 1.2]
     *
     * @param number Is the octal number you want to convert.
     * @param places Is the number of characters to use.
     */
    oct2Bin(number: number | string | boolean | Range | RangeReference | FunctionResult<any>, places?: number | string | boolean | Range | RangeReference | FunctionResult<any>): FunctionResult<number>;
    /**
     *
     * Converts an octal number to decimal.
     *
     * [Api set: ExcelApi 1.2]
     *
     * @param number Is the octal number you want to convert.
     */
    oct2Dec(number: number | string | boolean | Range | RangeReference | FunctionResult<any>): FunctionResult<number>;
    /**
     *
     * Converts an octal number to hexadecimal.
     *
     * [Api set: ExcelApi 1.2]
     *
     * @param number Is the octal number you want to convert.
     * @param places Is the number of characters to use.
     */
    oct2Hex(number: number | string | boolean | Range | RangeReference | FunctionResult<any>, places?: number | string | boolean | Range | RangeReference | FunctionResult<any>): FunctionResult<number>;
    /**
     *
     * Rounds a positive number up and negative number down to the nearest odd integer.
     *
     * [Api set: ExcelApi 1.2]
     *
     * @param number Is the value to round.
     */
    odd(number: number | Range | RangeReference | FunctionResult<any>): FunctionResult<number>;
    /**
     *
     * Returns the price per $100 face value of a security with an odd first period.
     *
     * [Api set: ExcelApi 1.2]
     *
     * @param settlement Is the security's settlement date, expressed as a serial date number.
     * @param maturity Is the security's maturity date, expressed as a serial date number.
     * @param issue Is the security's issue date, expressed as a serial date number.
     * @param firstCoupon Is the security's first coupon date, expressed as a serial date number.
     * @param rate Is the security's interest rate.
     * @param yld Is the security's annual yield.
     * @param redemption Is the security's redemption value per $100 face value.
     * @param frequency Is the number of coupon payments per year.
     * @param basis Is the type of day count basis to use.
     */
    oddFPrice(settlement: number | string | boolean | Range | RangeReference | FunctionResult<any>, maturity: number | string | boolean | Range | RangeReference | FunctionResult<any>, issue: number | string | boolean | Range | RangeReference | FunctionResult<any>, firstCoupon: number | string | boolean | Range | RangeReference | FunctionResult<any>, rate: number | string | boolean | Range | RangeReference | FunctionResult<any>, yld: number | string | boolean | Range | RangeReference | FunctionResult<any>, redemption: number | string | boolean | Range | RangeReference | FunctionResult<any>, frequency: number | string | boolean | Range | RangeReference | FunctionResult<any>, basis?: number | string | boolean | Range | RangeReference | FunctionResult<any>): FunctionResult<number>;
    /**
     *
     * Returns the yield of a security with an odd first period.
     *
     * [Api set: ExcelApi 1.2]
     *
     * @param settlement Is the security's settlement date, expressed as a serial date number.
     * @param maturity Is the security's maturity date, expressed as a serial date number.
     * @param issue Is the security's issue date, expressed as a serial date number.
     * @param firstCoupon Is the security's first coupon date, expressed as a serial date number.
     * @param rate Is the security's interest rate.
     * @param pr Is the security's price.
     * @param redemption Is the security's redemption value per $100 face value.
     * @param frequency Is the number of coupon payments per year.
     * @param basis Is the type of day count basis to use.
     */
    oddFYield(settlement: number | string | boolean | Range | RangeReference | FunctionResult<any>, maturity: number | string | boolean | Range | RangeReference | FunctionResult<any>, issue: number | string | boolean | Range | RangeReference | FunctionResult<any>, firstCoupon: number | string | boolean | Range | RangeReference | FunctionResult<any>, rate: number | string | boolean | Range | RangeReference | FunctionResult<any>, pr: number | string | boolean | Range | RangeReference | FunctionResult<any>, redemption: number | string | boolean | Range | RangeReference | FunctionResult<any>, frequency: number | string | boolean | Range | RangeReference | FunctionResult<any>, basis?: number | string | boolean | Range | RangeReference | FunctionResult<any>): FunctionResult<number>;
    /**
     *
     * Returns the price per $100 face value of a security with an odd last period.
     *
     * [Api set: ExcelApi 1.2]
     *
     * @param settlement Is the security's settlement date, expressed as a serial date number.
     * @param maturity Is the security's maturity date, expressed as a serial date number.
     * @param lastInterest Is the security's last coupon date, expressed as a serial date number.
     * @param rate Is the security's interest rate.
     * @param yld Is the security's annual yield.
     * @param redemption Is the security's redemption value per $100 face value.
     * @param frequency Is the number of coupon payments per year.
     * @param basis Is the type of day count basis to use.
     */
    oddLPrice(settlement: number | string | boolean | Range | RangeReference | FunctionResult<any>, maturity: number | string | boolean | Range | RangeReference | FunctionResult<any>, lastInterest: number | string | boolean | Range | RangeReference | FunctionResult<any>, rate: number | string | boolean | Range | RangeReference | FunctionResult<any>, yld: number | string | boolean | Range | RangeReference | FunctionResult<any>, redemption: number | string | boolean | Range | RangeReference | FunctionResult<any>, frequency: number | string | boolean | Range | RangeReference | FunctionResult<any>, basis?: number | string | boolean | Range | RangeReference | FunctionResult<any>): FunctionResult<number>;
    /**
     *
     * Returns the yield of a security with an odd last period.
     *
     * [Api set: ExcelApi 1.2]
     *
     * @param settlement Is the security's settlement date, expressed as a serial date number.
     * @param maturity Is the security's maturity date, expressed as a serial date number.
     * @param lastInterest Is the security's last coupon date, expressed as a serial date number.
     * @param rate Is the security's interest rate.
     * @param pr Is the security's price.
     * @param redemption Is the security's redemption value per $100 face value.
     * @param frequency Is the number of coupon payments per year.
     * @param basis Is the type of day count basis to use.
     */
    oddLYield(settlement: number | string | boolean | Range | RangeReference | FunctionResult<any>, maturity: number | string | boolean | Range | RangeReference | FunctionResult<any>, lastInterest: number | string | boolean | Range | RangeReference | FunctionResult<any>, rate: number | string | boolean | Range | RangeReference | FunctionResult<any>, pr: number | string | boolean | Range | RangeReference | FunctionResult<any>, redemption: number | string | boolean | Range | RangeReference | FunctionResult<any>, frequency: number | string | boolean | Range | RangeReference | FunctionResult<any>, basis?: number | string | boolean | Range | RangeReference | FunctionResult<any>): FunctionResult<number>;
    /**
     *
     * Checks whether any of the arguments are TRUE, and returns TRUE or FALSE. Returns FALSE only if all arguments are FALSE.
     *
     * [Api set: ExcelApi 1.2]
     *
     * @param values List of parameters, whose elements are 1 to 255 conditions that you want to test that can be either TRUE or FALSE.
     */
    or(...values: Array<boolean | Range | RangeReference | FunctionResult<any>>): FunctionResult<boolean>;
    /**
     *
     * Returns the number of periods required by an investment to reach a specified value.
     *
     * [Api set: ExcelApi 1.2]
     *
     * @param rate Is the interest rate per period.
     * @param pv Is the present value of the investment.
     * @param fv Is the desired future value of the investment.
     */
    pduration(rate: number | Range | RangeReference | FunctionResult<any>, pv: number | Range | RangeReference | FunctionResult<any>, fv: number | Range | RangeReference | FunctionResult<any>): FunctionResult<number>;
    /**
     *
     * Returns the rank of a value in a data set as a percentage of the data set as a percentage (0..1, exclusive) of the data set.
     *
     * [Api set: ExcelApi 1.2]
     *
     * @param array Is the array or range of data with numeric values that defines relative standing.
     * @param x Is the value for which you want to know the rank.
     * @param significance Is an optional value that identifies the number of significant digits for the returned percentage, three digits if omitted (0.xxx%).
     */
    percentRank_Exc(array: number | Range | RangeReference | FunctionResult<any>, x: number | Range | RangeReference | FunctionResult<any>, significance?: number | Range | RangeReference | FunctionResult<any>): FunctionResult<number>;
    /**
     *
     * Returns the rank of a value in a data set as a percentage of the data set as a percentage (0..1, inclusive) of the data set.
     *
     * [Api set: ExcelApi 1.2]
     *
     * @param array Is the array or range of data with numeric values that defines relative standing.
     * @param x Is the value for which you want to know the rank.
     * @param significance Is an optional value that identifies the number of significant digits for the returned percentage, three digits if omitted (0.xxx%).
     */
    percentRank_Inc(array: number | Range | RangeReference | FunctionResult<any>, x: number | Range | RangeReference | FunctionResult<any>, significance?: number | Range | RangeReference | FunctionResult<any>): FunctionResult<number>;
    /**
     *
     * Returns the k-th percentile of values in a range, where k is in the range 0..1, exclusive.
     *
     * [Api set: ExcelApi 1.2]
     *
     * @param array Is the array or range of data that defines relative standing.
     * @param k Is the percentile value that is between 0 through 1, inclusive.
     */
    percentile_Exc(array: number | Range | RangeReference | FunctionResult<any>, k: number | Range | RangeReference | FunctionResult<any>): FunctionResult<number>;
    /**
     *
     * Returns the k-th percentile of values in a range, where k is in the range 0..1, inclusive.
     *
     * [Api set: ExcelApi 1.2]
     *
     * @param array Is the array or range of data that defines relative standing.
     * @param k Is the percentile value that is between 0 through 1, inclusive.
     */
    percentile_Inc(array: number | Range | RangeReference | FunctionResult<any>, k: number | Range | RangeReference | FunctionResult<any>): FunctionResult<number>;
    /**
     *
     * Returns the number of permutations for a given number of objects that can be selected from the total objects.
     *
     * [Api set: ExcelApi 1.2]
     *
     * @param number Is the total number of objects.
     * @param numberChosen Is the number of objects in each permutation.
     */
    permut(number: number | Range | RangeReference | FunctionResult<any>, numberChosen: number | Range | RangeReference | FunctionResult<any>): FunctionResult<number>;
    /**
     *
     * Returns the number of permutations for a given number of objects (with repetitions) that can be selected from the total objects.
     *
     * [Api set: ExcelApi 1.2]
     *
     * @param number Is the total number of objects.
     * @param numberChosen Is the number of objects in each permutation.
     */
    permutationa(number: number | Range | RangeReference | FunctionResult<any>, numberChosen: number | Range | RangeReference | FunctionResult<any>): FunctionResult<number>;
    /**
     *
     * Returns the value of the density function for a standard normal distribution.
     *
     * [Api set: ExcelApi 1.2]
     *
     * @param x Is the number for which you want the density of the standard normal distribution.
     */
    phi(x: number | Range | RangeReference | FunctionResult<any>): FunctionResult<number>;
    /**
     *
     * Returns the value of Pi, 3.14159265358979, accurate to 15 digits.
     *
     * [Api set: ExcelApi 1.2]
     */
    pi(): FunctionResult<number>;
    /**
     *
     * Calculates the payment for a loan based on constant payments and a constant interest rate.
     *
     * [Api set: ExcelApi 1.2]
     *
     * @param rate Is the interest rate per period for the loan. For example, use 6%/4 for quarterly payments at 6% APR.
     * @param nper Is the total number of payments for the loan.
     * @param pv Is the present value: the total amount that a series of future payments is worth now.
     * @param fv Is the future value, or a cash balance you want to attain after the last payment is made, 0 (zero) if omitted.
     * @param type Is a logical value: payment at the beginning of the period = 1; payment at the end of the period = 0 or omitted.
     */
    pmt(rate: number | Range | RangeReference | FunctionResult<any>, nper: number | Range | RangeReference | FunctionResult<any>, pv: number | Range | RangeReference | FunctionResult<any>, fv?: number | Range | RangeReference | FunctionResult<any>, type?: number | Range | RangeReference | FunctionResult<any>): FunctionResult<number>;
    /**
     *
     * Returns the Poisson distribution.
     *
     * [Api set: ExcelApi 1.2]
     *
     * @param x Is the number of events.
     * @param mean Is the expected numeric value, a positive number.
     * @param cumulative Is a logical value: for the cumulative Poisson probability, use TRUE; for the Poisson probability mass function, use FALSE.
     */
    poisson_Dist(x: number | Range | RangeReference | FunctionResult<any>, mean: number | Range | RangeReference | FunctionResult<any>, cumulative: boolean | Range | RangeReference | FunctionResult<any>): FunctionResult<number>;
    /**
     *
     * Returns the result of a number raised to a power.
     *
     * [Api set: ExcelApi 1.2]
     *
     * @param number Is the base number, any real number.
     * @param power Is the exponent, to which the base number is raised.
     */
    power(number: number | Range | RangeReference | FunctionResult<any>, power: number | Range | RangeReference | FunctionResult<any>): FunctionResult<number>;
    /**
     *
     * Returns the payment on the principal for a given investment based on periodic, constant payments and a constant interest rate.
     *
     * [Api set: ExcelApi 1.2]
     *
     * @param rate Is the interest rate per period. For example, use 6%/4 for quarterly payments at 6% APR.
     * @param per Specifies the period and must be in the range 1 to nper.
     * @param nper Is the total number of payment periods in an investment.
     * @param pv Is the present value: the total amount that a series of future payments is worth now.
     * @param fv Is the future value, or cash balance you want to attain after the last payment is made.
     * @param type Is a logical value: payment at the beginning of the period = 1; payment at the end of the period = 0 or omitted.
     */
    ppmt(rate: number | Range | RangeReference | FunctionResult<any>, per: number | Range | RangeReference | FunctionResult<any>, nper: number | Range | RangeReference | FunctionResult<any>, pv: number | Range | RangeReference | FunctionResult<any>, fv?: number | Range | RangeReference | FunctionResult<any>, type?: number | Range | RangeReference | FunctionResult<any>): FunctionResult<number>;
    /**
     *
     * Returns the price per $100 face value of a security that pays periodic interest.
     *
     * [Api set: ExcelApi 1.2]
     *
     * @param settlement Is the security's settlement date, expressed as a serial date number.
     * @param maturity Is the security's maturity date, expressed as a serial date number.
     * @param rate Is the security's annual coupon rate.
     * @param yld Is the security's annual yield.
     * @param redemption Is the security's redemption value per $100 face value.
     * @param frequency Is the number of coupon payments per year.
     * @param basis Is the type of day count basis to use.
     */
    price(settlement: number | string | boolean | Range | RangeReference | FunctionResult<any>, maturity: number | string | boolean | Range | RangeReference | FunctionResult<any>, rate: number | string | boolean | Range | RangeReference | FunctionResult<any>, yld: number | string | boolean | Range | RangeReference | FunctionResult<any>, redemption: number | string | boolean | Range | RangeReference | FunctionResult<any>, frequency: number | string | boolean | Range | RangeReference | FunctionResult<any>, basis?: number | string | boolean | Range | RangeReference | FunctionResult<any>): FunctionResult<number>;
    /**
     *
     * Returns the price per $100 face value of a discounted security.
     *
     * [Api set: ExcelApi 1.2]
     *
     * @param settlement Is the security's settlement date, expressed as a serial date number.
     * @param maturity Is the security's maturity date, expressed as a serial date number.
     * @param discount Is the security's discount rate.
     * @param redemption Is the security's redemption value per $100 face value.
     * @param basis Is the type of day count basis to use.
     */
    priceDisc(settlement: number | string | boolean | Range | RangeReference | FunctionResult<any>, maturity: number | string | boolean | Range | RangeReference | FunctionResult<any>, discount: number | string | boolean | Range | RangeReference | FunctionResult<any>, redemption: number | string | boolean | Range | RangeReference | FunctionResult<any>, basis?: number | string | boolean | Range | RangeReference | FunctionResult<any>): FunctionResult<number>;
    /**
     *
     * Returns the price per $100 face value of a security that pays interest at maturity.
     *
     * [Api set: ExcelApi 1.2]
     *
     * @param settlement Is the security's settlement date, expressed as a serial date number.
     * @param maturity Is the security's maturity date, expressed as a serial date number.
     * @param issue Is the security's issue date, expressed as a serial date number.
     * @param rate Is the security's interest rate at date of issue.
     * @param yld Is the security's annual yield.
     * @param basis Is the type of day count basis to use.
     */
    priceMat(settlement: number | string | boolean | Range | RangeReference | FunctionResult<any>, maturity: number | string | boolean | Range | RangeReference | FunctionResult<any>, issue: number | string | boolean | Range | RangeReference | FunctionResult<any>, rate: number | string | boolean | Range | RangeReference | FunctionResult<any>, yld: number | string | boolean | Range | RangeReference | FunctionResult<any>, basis?: number | string | boolean | Range | RangeReference | FunctionResult<any>): FunctionResult<number>;
    /**
     *
     * Multiplies all the numbers given as arguments.
     *
     * [Api set: ExcelApi 1.2]
     *
     * @param values List of parameters, whose elements are 1 to 255 numbers, logical values, or text representations of numbers that you want to multiply.
     */
    product(...values: Array<number | Range | RangeReference | FunctionResult<any>>): FunctionResult<number>;
    /**
     *
     * Converts a text string to proper case; the first letter in each word to uppercase, and all other letters to lowercase.
     *
     * [Api set: ExcelApi 1.2]
     *
     * @param text Is text enclosed in quotation marks, a formula that returns text, or a reference to a cell containing text to partially capitalize.
     */
    proper(text: string | Range | RangeReference | FunctionResult<any>): FunctionResult<string>;
    /**
     *
     * Returns the present value of an investment: the total amount that a series of future payments is worth now.
     *
     * [Api set: ExcelApi 1.2]
     *
     * @param rate Is the interest rate per period. For example, use 6%/4 for quarterly payments at 6% APR.
     * @param nper Is the total number of payment periods in an investment.
     * @param pmt Is the payment made each period and cannot change over the life of the investment.
     * @param fv Is the future value, or a cash balance you want to attain after the last payment is made.
     * @param type Is a logical value: payment at the beginning of the period = 1; payment at the end of the period = 0 or omitted.
     */
    pv(rate: number | Range | RangeReference | FunctionResult<any>, nper: number | Range | RangeReference | FunctionResult<any>, pmt: number | Range | RangeReference | FunctionResult<any>, fv?: number | Range | RangeReference | FunctionResult<any>, type?: number | Range | RangeReference | FunctionResult<any>): FunctionResult<number>;
    /**
     *
     * Returns the quartile of a data set, based on percentile values from 0..1, exclusive.
     *
     * [Api set: ExcelApi 1.2]
     *
     * @param array Is the array or cell range of numeric values for which you want the quartile value.
     * @param quart Is a number: minimum value = 0; 1st quartile = 1; median value = 2; 3rd quartile = 3; maximum value = 4.
     */
    quartile_Exc(array: number | Range | RangeReference | FunctionResult<any>, quart: number | Range | RangeReference | FunctionResult<any>): FunctionResult<number>;
    /**
     *
     * Returns the quartile of a data set, based on percentile values from 0..1, inclusive.
     *
     * [Api set: ExcelApi 1.2]
     *
     * @param array Is the array or cell range of numeric values for which you want the quartile value.
     * @param quart Is a number: minimum value = 0; 1st quartile = 1; median value = 2; 3rd quartile = 3; maximum value = 4.
     */
    quartile_Inc(array: number | Range | RangeReference | FunctionResult<any>, quart: number | Range | RangeReference | FunctionResult<any>): FunctionResult<number>;
    /**
     *
     * Returns the integer portion of a division.
     *
     * [Api set: ExcelApi 1.2]
     *
     * @param numerator Is the dividend.
     * @param denominator Is the divisor.
     */
    quotient(numerator: number | string | boolean | Range | RangeReference | FunctionResult<any>, denominator: number | string | boolean | Range | RangeReference | FunctionResult<any>): FunctionResult<number>;
    /**
     *
     * Converts degrees to radians.
     *
     * [Api set: ExcelApi 1.2]
     *
     * @param angle Is an angle in degrees that you want to convert.
     */
    radians(angle: number | Range | RangeReference | FunctionResult<any>): FunctionResult<number>;
    /**
     *
     * Returns a random number greater than or equal to 0 and less than 1, evenly distributed (changes on recalculation).
     *
     * [Api set: ExcelApi 1.2]
     */
    rand(): FunctionResult<number>;
    /**
     *
     * Returns a random number between the numbers you specify.
     *
     * [Api set: ExcelApi 1.2]
     *
     * @param bottom Is the smallest integer RANDBETWEEN will return.
     * @param top Is the largest integer RANDBETWEEN will return.
     */
    randBetween(bottom: number | string | boolean | Range | RangeReference | FunctionResult<any>, top: number | string | boolean | Range | RangeReference | FunctionResult<any>): FunctionResult<number>;
    /**
     *
     * Returns the rank of a number in a list of numbers: its size relative to other values in the list; if more than one value has the same rank, the average rank is returned.
     *
     * [Api set: ExcelApi 1.2]
     *
     * @param number Is the number for which you want to find the rank.
     * @param ref Is an array of, or a reference to, a list of numbers. Nonnumeric values are ignored.
     * @param order Is a number: rank in the list sorted descending = 0 or omitted; rank in the list sorted ascending = any nonzero value.
     */
    rank_Avg(number: number | Range | RangeReference | FunctionResult<any>, ref: Range | RangeReference | FunctionResult<any>, order?: boolean | Range | RangeReference | FunctionResult<any>): FunctionResult<number>;
    /**
     *
     * Returns the rank of a number in a list of numbers: its size relative to other values in the list; if more than one value has the same rank, the top rank of that set of values is returned.
     *
     * [Api set: ExcelApi 1.2]
     *
     * @param number Is the number for which you want to find the rank.
     * @param ref Is an array of, or a reference to, a list of numbers. Nonnumeric values are ignored.
     * @param order Is a number: rank in the list sorted descending = 0 or omitted; rank in the list sorted ascending = any nonzero value.
     */
    rank_Eq(number: number | Range | RangeReference | FunctionResult<any>, ref: Range | RangeReference | FunctionResult<any>, order?: boolean | Range | RangeReference | FunctionResult<any>): FunctionResult<number>;
    /**
     *
     * Returns the interest rate per period of a loan or an investment. For example, use 6%/4 for quarterly payments at 6% APR.
     *
     * [Api set: ExcelApi 1.2]
     *
     * @param nper Is the total number of payment periods for the loan or investment.
     * @param pmt Is the payment made each period and cannot change over the life of the loan or investment.
     * @param pv Is the present value: the total amount that a series of future payments is worth now.
     * @param fv Is the future value, or a cash balance you want to attain after the last payment is made. If omitted, uses Fv = 0.
     * @param type Is a logical value: payment at the beginning of the period = 1; payment at the end of the period = 0 or omitted.
     * @param guess Is your guess for what the rate will be; if omitted, Guess = 0.1 (10 percent).
     */
    rate(nper: number | Range | RangeReference | FunctionResult<any>, pmt: number | Range | RangeReference | FunctionResult<any>, pv: number | Range | RangeReference | FunctionResult<any>, fv?: number | Range | RangeReference | FunctionResult<any>, type?: number | Range | RangeReference | FunctionResult<any>, guess?: number | Range | RangeReference | FunctionResult<any>): FunctionResult<number>;
    /**
     *
     * Returns the amount received at maturity for a fully invested security.
     *
     * [Api set: ExcelApi 1.2]
     *
     * @param settlement Is the security's settlement date, expressed as a serial date number.
     * @param maturity Is the security's maturity date, expressed as a serial date number.
     * @param investment Is the amount invested in the security.
     * @param discount Is the security's discount rate.
     * @param basis Is the type of day count basis to use.
     */
    received(settlement: number | string | boolean | Range | RangeReference | FunctionResult<any>, maturity: number | string | boolean | Range | RangeReference | FunctionResult<any>, investment: number | string | boolean | Range | RangeReference | FunctionResult<any>, discount: number | string | boolean | Range | RangeReference | FunctionResult<any>, basis?: number | string | boolean | Range | RangeReference | FunctionResult<any>): FunctionResult<number>;
    /**
     *
     * Replaces part of a text string with a different text string.
     *
     * [Api set: ExcelApi 1.2]
     *
     * @param oldText Is text in which you want to replace some characters.
     * @param startNum Is the position of the character in oldText that you want to replace with newText.
     * @param numChars Is the number of characters in oldText that you want to replace.
     * @param newText Is the text that will replace characters in oldText.
     */
    replace(oldText: string | Range | RangeReference | FunctionResult<any>, startNum: number | Range | RangeReference | FunctionResult<any>, numChars: number | Range | RangeReference | FunctionResult<any>, newText: string | Range | RangeReference | FunctionResult<any>): FunctionResult<string>;
    /**
     *
     * Replaces part of a text string with a different text string. Use with double-byte character sets (DBCS).
     *
     * [Api set: ExcelApi 1.2]
     *
     * @param oldText Is text in which you want to replace some characters.
     * @param startNum Is the position of the character in oldText that you want to replace with newText.
     * @param numBytes Is the number of characters in oldText that you want to replace with newText.
     * @param newText Is the text that will replace characters in oldText.
     */
    replaceB(oldText: string | Range | RangeReference | FunctionResult<any>, startNum: number | Range | RangeReference | FunctionResult<any>, numBytes: number | Range | RangeReference | FunctionResult<any>, newText: string | Range | RangeReference | FunctionResult<any>): FunctionResult<string>;
    /**
     *
     * Repeats text a given number of times. Use REPT to fill a cell with a number of instances of a text string.
     *
     * [Api set: ExcelApi 1.2]
     *
     * @param text Is the text you want to repeat.
     * @param numberTimes Is a positive number specifying the number of times to repeat text.
     */
    rept(text: string | Range | RangeReference | FunctionResult<any>, numberTimes: number | Range | RangeReference | FunctionResult<any>): FunctionResult<string>;
    /**
     *
     * Returns the specified number of characters from the end of a text string.
     *
     * [Api set: ExcelApi 1.2]
     *
     * @param text Is the text string that contains the characters you want to extract.
     * @param numChars Specifies how many characters you want to extract, 1 if omitted.
     */
    right(text: string | Range | RangeReference | FunctionResult<any>, numChars?: number | Range | RangeReference | FunctionResult<any>): FunctionResult<string>;
    /**
     *
     * Returns the specified number of characters from the end of a text string. Use with double-byte character sets (DBCS).
     *
     * [Api set: ExcelApi 1.2]
     *
     * @param text Is the text string containing the characters you want to extract.
     * @param numBytes Specifies how many characters you want to extract.
     */
    rightb(text: string | Range | RangeReference | FunctionResult<any>, numBytes?: number | Range | RangeReference | FunctionResult<any>): FunctionResult<string>;
    /**
     *
     * Converts an Arabic numeral to Roman, as text.
     *
     * [Api set: ExcelApi 1.2]
     *
     * @param number Is the Arabic numeral you want to convert.
     * @param form Is the number specifying the type of Roman numeral you want.
     */
    roman(number: number | Range | RangeReference | FunctionResult<any>, form?: boolean | number | Range | RangeReference | FunctionResult<any>): FunctionResult<string>;
    /**
     *
     * Rounds a number to a specified number of digits.
     *
     * [Api set: ExcelApi 1.2]
     *
     * @param number Is the number you want to round.
     * @param numDigits Is the number of digits to which you want to round. Negative rounds to the left of the decimal point; zero to the nearest integer.
     */
    round(number: number | Range | RangeReference | FunctionResult<any>, numDigits: number | Range | RangeReference | FunctionResult<any>): FunctionResult<number>;
    /**
     *
     * Rounds a number down, toward zero.
     *
     * [Api set: ExcelApi 1.2]
     *
     * @param number Is any real number that you want rounded down.
     * @param numDigits Is the number of digits to which you want to round. Negative rounds to the left of the decimal point; zero or omitted, to the nearest integer.
     */
    roundDown(number: number | Range | RangeReference | FunctionResult<any>, numDigits: number | Range | RangeReference | FunctionResult<any>): FunctionResult<number>;
    /**
     *
     * Rounds a number up, away from zero.
     *
     * [Api set: ExcelApi 1.2]
     *
     * @param number Is any real number that you want rounded up.
     * @param numDigits Is the number of digits to which you want to round. Negative rounds to the left of the decimal point; zero or omitted, to the nearest integer.
     */
    roundUp(number: number | Range | RangeReference | FunctionResult<any>, numDigits: number | Range | RangeReference | FunctionResult<any>): FunctionResult<number>;
    /**
     *
     * Returns the number of rows in a reference or array.
     *
     * [Api set: ExcelApi 1.2]
     *
     * @param array Is an array, an array formula, or a reference to a range of cells for which you want the number of rows.
     */
    rows(array: Range | RangeReference | FunctionResult<any>): FunctionResult<number>;
    /**
     *
     * Returns an equivalent interest rate for the growth of an investment.
     *
     * [Api set: ExcelApi 1.2]
     *
     * @param nper Is the number of periods for the investment.
     * @param pv Is the present value of the investment.
     * @param fv Is the future value of the investment.
     */
    rri(nper: number | Range | RangeReference | FunctionResult<any>, pv: number | Range | RangeReference | FunctionResult<any>, fv: number | Range | RangeReference | FunctionResult<any>): FunctionResult<number>;
    /**
     *
     * Returns the secant of an angle.
     *
     * [Api set: ExcelApi 1.2]
     *
     * @param number Is the angle in radians for which you want the secant.
     */
    sec(number: number | Range | RangeReference | FunctionResult<any>): FunctionResult<number>;
    /**
     *
     * Returns the hyperbolic secant of an angle.
     *
     * [Api set: ExcelApi 1.2]
     *
     * @param number Is the angle in radians for which you want the hyperbolic secant.
     */
    sech(number: number | Range | RangeReference | FunctionResult<any>): FunctionResult<number>;
    /**
     *
     * Returns the second, a number from 0 to 59.
     *
     * [Api set: ExcelApi 1.2]
     *
     * @param serialNumber Is a number in the date-time code used by Microsoft Excel or text in time format, such as 16:48:23 or 4:48:47 PM.
     */
    second(serialNumber: number | Range | RangeReference | FunctionResult<any>): FunctionResult<number>;
    /**
     *
     * Returns the sum of a power series based on the formula.
     *
     * [Api set: ExcelApi 1.2]
     *
     * @param x Is the input value to the power series.
     * @param n Is the initial power to which you want to raise x.
     * @param m Is the step by which to increase n for each term in the series.
     * @param coefficients Is a set of coefficients by which each successive power of x is multiplied.
     */
    seriesSum(x: number | string | boolean | Range | RangeReference | FunctionResult<any>, n: number | string | boolean | Range | RangeReference | FunctionResult<any>, m: number | string | boolean | Range | RangeReference | FunctionResult<any>, coefficients: Range | string | number | boolean | RangeReference | FunctionResult<any>): FunctionResult<number>;
    /**
     *
     * Returns the sheet number of the referenced sheet.
     *
     * [Api set: ExcelApi 1.2]
     *
     * @param value Is the name of a sheet or a reference that you want the sheet number of.  If omitted the number of the sheet containing the function is returned.
     */
    sheet(value?: Range | string | RangeReference | FunctionResult<any>): FunctionResult<number>;
    /**
     *
     * Returns the number of sheets in a reference.
     *
     * [Api set: ExcelApi 1.2]
     *
     * @param reference Is a reference for which you want to know the number of sheets it contains.  If omitted the number of sheets in the workbook containing the function is returned.
     */
    sheets(reference?: Range | RangeReference | FunctionResult<any>): FunctionResult<number>;
    /**
     *
     * Returns the sign of a number: 1 if the number is positive, zero if the number is zero, or -1 if the number is negative.
     *
     * [Api set: ExcelApi 1.2]
     *
     * @param number Is any real number.
     */
    sign(number: number | Range | RangeReference | FunctionResult<any>): FunctionResult<number>;
    /**
     *
     * Returns the sine of an angle.
     *
     * [Api set: ExcelApi 1.2]
     *
     * @param number Is the angle in radians for which you want the sine. Degrees * PI()/180 = radians.
     */
    sin(number: number | Range | RangeReference | FunctionResult<any>): FunctionResult<number>;
    /**
     *
     * Returns the hyperbolic sine of a number.
     *
     * [Api set: ExcelApi 1.2]
     *
     * @param number Is any real number.
     */
    sinh(number: number | Range | RangeReference | FunctionResult<any>): FunctionResult<number>;
    /**
     *
     * Returns the skewness of a distribution: a characterization of the degree of asymmetry of a distribution around its mean.
     *
     * [Api set: ExcelApi 1.2]
     *
     * @param values List of parameters, whose elements are 1 to 255 numbers or names, arrays, or references that contain numbers for which you want the skewness.
     */
    skew(...values: Array<number | Range | RangeReference | FunctionResult<any>>): FunctionResult<number>;
    /**
     *
     * Returns the skewness of a distribution based on a population: a characterization of the degree of asymmetry of a distribution around its mean.
     *
     * [Api set: ExcelApi 1.2]
     *
     * @param values List of parameters, whose elements are 1 to 254 numbers or names, arrays, or references that contain numbers for which you want the population skewness.
     */
    skew_p(...values: Array<number | Range | RangeReference | FunctionResult<any>>): FunctionResult<number>;
    /**
     *
     * Returns the straight-line depreciation of an asset for one period.
     *
     * [Api set: ExcelApi 1.2]
     *
     * @param cost Is the initial cost of the asset.
     * @param salvage Is the salvage value at the end of the life of the asset.
     * @param life Is the number of periods over which the asset is being depreciated (sometimes called the useful life of the asset).
     */
    sln(cost: number | Range | RangeReference | FunctionResult<any>, salvage: number | Range | RangeReference | FunctionResult<any>, life: number | Range | RangeReference | FunctionResult<any>): FunctionResult<number>;
    /**
     *
     * Returns the k-th smallest value in a data set. For example, the fifth smallest number.
     *
     * [Api set: ExcelApi 1.2]
     *
     * @param array Is an array or range of numerical data for which you want to determine the k-th smallest value.
     * @param k Is the position (from the smallest) in the array or range of the value to return.
     */
    small(array: number | Range | RangeReference | FunctionResult<any>, k: number | Range | RangeReference | FunctionResult<any>): FunctionResult<number>;
    /**
     *
     * Returns the square root of a number.
     *
     * [Api set: ExcelApi 1.2]
     *
     * @param number Is the number for which you want the square root.
     */
    sqrt(number: number | Range | RangeReference | FunctionResult<any>): FunctionResult<number>;
    /**
     *
     * Returns the square root of (number * Pi).
     *
     * [Api set: ExcelApi 1.2]
     *
     * @param number Is the number by which p is multiplied.
     */
    sqrtPi(number: number | string | boolean | Range | RangeReference | FunctionResult<any>): FunctionResult<number>;
    /**
     *
     * Estimates standard deviation based on a sample, including logical values and text. Text and the logical value FALSE have the value 0; the logical value TRUE has the value 1.
     *
     * [Api set: ExcelApi 1.2]
     *
     * @param values List of parameters, whose elements are 1 to 255 values corresponding to a sample of a population and can be values or names or references to values.
     */
    stDevA(...values: Array<number | Range | RangeReference | FunctionResult<any>>): FunctionResult<number>;
    /**
     *
     * Calculates standard deviation based on an entire population, including logical values and text. Text and the logical value FALSE have the value 0; the logical value TRUE has the value 1.
     *
     * [Api set: ExcelApi 1.2]
     *
     * @param values List of parameters, whose elements are 1 to 255 values corresponding to a population and can be values, names, arrays, or references that contain values.
     */
    stDevPA(...values: Array<number | Range | RangeReference | FunctionResult<any>>): FunctionResult<number>;
    /**
     *
     * Calculates standard deviation based on the entire population given as arguments (ignores logical values and text).
     *
     * [Api set: ExcelApi 1.2]
     *
     * @param values List of parameters, whose elements are 1 to 255 numbers corresponding to a population and can be numbers or references that contain numbers.
     */
    stDev_P(...values: Array<number | Range | RangeReference | FunctionResult<any>>): FunctionResult<number>;
    /**
     *
     * Estimates standard deviation based on a sample (ignores logical values and text in the sample).
     *
     * [Api set: ExcelApi 1.2]
     *
     * @param values List of parameters, whose elements are 1 to 255 numbers corresponding to a sample of a population and can be numbers or references that contain numbers.
     */
    stDev_S(...values: Array<number | Range | RangeReference | FunctionResult<any>>): FunctionResult<number>;
    /**
     *
     * Returns a normalized value from a distribution characterized by a mean and standard deviation.
     *
     * [Api set: ExcelApi 1.2]
     *
     * @param x Is the value you want to normalize.
     * @param mean Is the arithmetic mean of the distribution.
     * @param standardDev Is the standard deviation of the distribution, a positive number.
     */
    standardize(x: number | Range | RangeReference | FunctionResult<any>, mean: number | Range | RangeReference | FunctionResult<any>, standardDev: number | Range | RangeReference | FunctionResult<any>): FunctionResult<number>;
    /**
     *
     * Replaces existing text with new text in a text string.
     *
     * [Api set: ExcelApi 1.2]
     *
     * @param text Is the text or the reference to a cell containing text in which you want to substitute characters.
     * @param oldText Is the existing text you want to replace. If the case of oldText does not match the case of text, SUBSTITUTE will not replace the text.
     * @param newText Is the text you want to replace oldText with.
     * @param instanceNum Specifies which occurrence of oldText you want to replace. If omitted, every instance of oldText is replaced.
     */
    substitute(text: string | Range | RangeReference | FunctionResult<any>, oldText: string | Range | RangeReference | FunctionResult<any>, newText: string | Range | RangeReference | FunctionResult<any>, instanceNum?: string | Range | RangeReference | FunctionResult<any>): FunctionResult<string>;
    /**
     *
     * Returns a subtotal in a list or database.
     *
     * [Api set: ExcelApi 1.2]
     *
     * @param functionNum Is the number 1 to 11 that specifies the summary function for the subtotal.
     * @param values List of parameters, whose elements are 1 to 254 ranges or references for which you want the subtotal.
     */
    subtotal(functionNum: number | Range | RangeReference | FunctionResult<any>, ...values: Array<Range | RangeReference | FunctionResult<any>>): FunctionResult<number>;
    /**
     *
     * Adds all the numbers in a range of cells.
     *
     * [Api set: ExcelApi 1.2]
     *
     * @param values List of parameters, whose elements are 1 to 255 numbers to sum. Logical values and text are ignored in cells, included if typed as arguments.
     */
    sum(...values: Array<number | Range | RangeReference | FunctionResult<any>>): FunctionResult<number>;
    /**
     *
     * Adds the cells specified by a given condition or criteria.
     *
     * [Api set: ExcelApi 1.2]
     *
     * @param range Is the range of cells you want evaluated.
     * @param criteria Is the condition or criteria in the form of a number, expression, or text that defines which cells will be added.
     * @param sumRange Are the actual cells to sum. If omitted, the cells in range are used.
     */
    sumIf(range: Range | RangeReference | FunctionResult<any>, criteria: number | string | boolean | Range | RangeReference | FunctionResult<any>, sumRange?: Range | RangeReference | FunctionResult<any>): FunctionResult<number>;
    /**
     *
     * Adds the cells specified by a given set of conditions or criteria.
     *
     * [Api set: ExcelApi 1.2]
     *
     * @param sumRange Are the actual cells to sum.
     * @param values List of parameters, where the first element of each pair is the Is the range of cells you want evaluated for the particular condition , and the second element is is the condition or criteria in the form of a number, expression, or text that defines which cells will be added.
     */
    sumIfs(sumRange: Range | RangeReference | FunctionResult<any>, ...values: Array<Range | RangeReference | FunctionResult<any> | number | string | boolean>): FunctionResult<number>;
    /**
     *
     * Returns the sum of the squares of the arguments. The arguments can be numbers, arrays, names, or references to cells that contain numbers.
     *
     * [Api set: ExcelApi 1.2]
     *
     * @param values List of parameters, whose elements are 1 to 255 numbers, arrays, names, or references to arrays for which you want the sum of the squares.
     */
    sumSq(...values: Array<number | Range | RangeReference | FunctionResult<any>>): FunctionResult<number>;
    /**
     *
     * Returns the sum-of-years' digits depreciation of an asset for a specified period.
     *
     * [Api set: ExcelApi 1.2]
     *
     * @param cost Is the initial cost of the asset.
     * @param salvage Is the salvage value at the end of the life of the asset.
     * @param life Is the number of periods over which the asset is being depreciated (sometimes called the useful life of the asset).
     * @param per Is the period and must use the same units as Life.
     */
    syd(cost: number | Range | RangeReference | FunctionResult<any>, salvage: number | Range | RangeReference | FunctionResult<any>, life: number | Range | RangeReference | FunctionResult<any>, per: number | Range | RangeReference | FunctionResult<any>): FunctionResult<number>;
    /**
     *
     * Checks whether a value is text, and returns the text if it is, or returns double quotes (empty text) if it is not.
     *
     * [Api set: ExcelApi 1.2]
     *
     * @param value Is the value to test.
     */
    t(value: number | string | boolean | Range | RangeReference | FunctionResult<any>): FunctionResult<string>;
    /**
     *
     * Returns the bond-equivalent yield for a treasury bill.
     *
     * [Api set: ExcelApi 1.2]
     *
     * @param settlement Is the Treasury bill's settlement date, expressed as a serial date number.
     * @param maturity Is the Treasury bill's maturity date, expressed as a serial date number.
     * @param discount Is the Treasury bill's discount rate.
     */
    tbillEq(settlement: number | string | boolean | Range | RangeReference | FunctionResult<any>, maturity: number | string | boolean | Range | RangeReference | FunctionResult<any>, discount: number | string | boolean | Range | RangeReference | FunctionResult<any>): FunctionResult<number>;
    /**
     *
     * Returns the price per $100 face value for a treasury bill.
     *
     * [Api set: ExcelApi 1.2]
     *
     * @param settlement Is the Treasury bill's settlement date, expressed as a serial date number.
     * @param maturity Is the Treasury bill's maturity date, expressed as a serial date number.
     * @param discount Is the Treasury bill's discount rate.
     */
    tbillPrice(settlement: number | string | boolean | Range | RangeReference | FunctionResult<any>, maturity: number | string | boolean | Range | RangeReference | FunctionResult<any>, discount: number | string | boolean | Range | RangeReference | FunctionResult<any>): FunctionResult<number>;
    /**
     *
     * Returns the yield for a treasury bill.
     *
     * [Api set: ExcelApi 1.2]
     *
     * @param settlement Is the Treasury bill's settlement date, expressed as a serial date number.
     * @param maturity Is the Treasury bill's maturity date, expressed as a serial date number.
     * @param pr Is the Treasury Bill's price per $100 face value.
     */
    tbillYield(settlement: number | string | boolean | Range | RangeReference | FunctionResult<any>, maturity: number | string | boolean | Range | RangeReference | FunctionResult<any>, pr: number | string | boolean | Range | RangeReference | FunctionResult<any>): FunctionResult<number>;
    /**
     *
     * Returns the left-tailed Student's t-distribution.
     *
     * [Api set: ExcelApi 1.2]
     *
     * @param x Is the numeric value at which to evaluate the distribution.
     * @param degFreedom Is an integer indicating the number of degrees of freedom that characterize the distribution.
     * @param cumulative Is a logical value: for the cumulative distribution function, use TRUE; for the probability density function, use FALSE.
     */
    t_Dist(x: number | Range | RangeReference | FunctionResult<any>, degFreedom: number | Range | RangeReference | FunctionResult<any>, cumulative: boolean | Range | RangeReference | FunctionResult<any>): FunctionResult<number>;
    /**
     *
     * Returns the two-tailed Student's t-distribution.
     *
     * [Api set: ExcelApi 1.2]
     *
     * @param x Is the numeric value at which to evaluate the distribution.
     * @param degFreedom Is an integer indicating the number of degrees of freedom that characterize the distribution.
     */
    t_Dist_2T(x: number | Range | RangeReference | FunctionResult<any>, degFreedom: number | Range | RangeReference | FunctionResult<any>): FunctionResult<number>;
    /**
     *
     * Returns the right-tailed Student's t-distribution.
     *
     * [Api set: ExcelApi 1.2]
     *
     * @param x Is the numeric value at which to evaluate the distribution.
     * @param degFreedom Is an integer indicating the number of degrees of freedom that characterize the distribution.
     */
    t_Dist_RT(x: number | Range | RangeReference | FunctionResult<any>, degFreedom: number | Range | RangeReference | FunctionResult<any>): FunctionResult<number>;
    /**
     *
     * Returns the left-tailed inverse of the Student's t-distribution.
     *
     * [Api set: ExcelApi 1.2]
     *
     * @param probability Is the probability associated with the two-tailed Student's t-distribution, a number between 0 and 1 inclusive.
     * @param degFreedom Is a positive integer indicating the number of degrees of freedom to characterize the distribution.
     */
    t_Inv(probability: number | Range | RangeReference | FunctionResult<any>, degFreedom: number | Range | RangeReference | FunctionResult<any>): FunctionResult<number>;
    /**
     *
     * Returns the two-tailed inverse of the Student's t-distribution.
     *
     * [Api set: ExcelApi 1.2]
     *
     * @param probability Is the probability associated with the two-tailed Student's t-distribution, a number between 0 and 1 inclusive.
     * @param degFreedom Is a positive integer indicating the number of degrees of freedom to characterize the distribution.
     */
    t_Inv_2T(probability: number | Range | RangeReference | FunctionResult<any>, degFreedom: number | Range | RangeReference | FunctionResult<any>): FunctionResult<number>;
    /**
     *
     * Returns the tangent of an angle.
     *
     * [Api set: ExcelApi 1.2]
     *
     * @param number Is the angle in radians for which you want the tangent. Degrees * PI()/180 = radians.
     */
    tan(number: number | Range | RangeReference | FunctionResult<any>): FunctionResult<number>;
    /**
     *
     * Returns the hyperbolic tangent of a number.
     *
     * [Api set: ExcelApi 1.2]
     *
     * @param number Is any real number.
     */
    tanh(number: number | Range | RangeReference | FunctionResult<any>): FunctionResult<number>;
    /**
     *
     * Converts a value to text in a specific number format.
     *
     * [Api set: ExcelApi 1.2]
     *
     * @param value Is a number, a formula that evaluates to a numeric value, or a reference to a cell containing a numeric value.
     * @param formatText Is a number format in text form from the Category box on the Number tab in the Format Cells dialog box (not General).
     */
    text(value: number | string | boolean | Range | RangeReference | FunctionResult<any>, formatText: string | Range | RangeReference | FunctionResult<any>): FunctionResult<string>;
    /**
     *
     * Converts hours, minutes, and seconds given as numbers to an Excel serial number, formatted with a time format.
     *
     * [Api set: ExcelApi 1.2]
     *
     * @param hour Is a number from 0 to 23 representing the hour.
     * @param minute Is a number from 0 to 59 representing the minute.
     * @param second Is a number from 0 to 59 representing the second.
     */
    time(hour: number | Range | RangeReference | FunctionResult<any>, minute: number | Range | RangeReference | FunctionResult<any>, second: number | Range | RangeReference | FunctionResult<any>): FunctionResult<number>;
    /**
     *
     * Converts a text time to an Excel serial number for a time, a number from 0 (12:00:00 AM) to 0.999988426 (11:59:59 PM). Format the number with a time format after entering the formula.
     *
     * [Api set: ExcelApi 1.2]
     *
     * @param timeText Is a text string that gives a time in any one of the Microsoft Excel time formats (date information in the string is ignored).
     */
    timevalue(timeText: string | number | Range | RangeReference | FunctionResult<any>): FunctionResult<number>;
    /**
     *
     * Returns the current date formatted as a date.
     *
     * [Api set: ExcelApi 1.2]
     */
    today(): FunctionResult<number>;
    /**
     *
     * Removes all spaces from a text string except for single spaces between words.
     *
     * [Api set: ExcelApi 1.2]
     *
     * @param text Is the text from which you want spaces removed.
     */
    trim(text: string | Range | RangeReference | FunctionResult<any>): FunctionResult<string>;
    /**
     *
     * Returns the mean of the interior portion of a set of data values.
     *
     * [Api set: ExcelApi 1.2]
     *
     * @param array Is the range or array of values to trim and average.
     * @param percent Is the fractional number of data points to exclude from the top and bottom of the data set.
     */
    trimMean(array: number | Range | RangeReference | FunctionResult<any>, percent: number | Range | RangeReference | FunctionResult<any>): FunctionResult<number>;
    /**
     *
     * Returns the logical value TRUE.
     *
     * [Api set: ExcelApi 1.2]
     */
    true(): FunctionResult<boolean>;
    /**
     *
     * Truncates a number to an integer by removing the decimal, or fractional, part of the number.
     *
     * [Api set: ExcelApi 1.2]
     *
     * @param number Is the number you want to truncate.
     * @param numDigits Is a number specifying the precision of the truncation, 0 (zero) if omitted.
     */
    trunc(number: number | Range | RangeReference | FunctionResult<any>, numDigits?: number | Range | RangeReference | FunctionResult<any>): FunctionResult<number>;
    /**
     *
     * Returns an integer representing the data type of a value: number = 1; text = 2; logical value = 4; error value = 16; array = 64.
     *
     * [Api set: ExcelApi 1.2]
     *
     * @param value Can be any value.
     */
    type(value: boolean | string | number | Range | RangeReference | FunctionResult<any>): FunctionResult<number>;
    /**
     *
     * Converts a number to text, using currency format.
     *
     * [Api set: ExcelApi 1.2]
     *
     * @param number Is a number, a reference to a cell containing a number, or a formula that evaluates to a number.
     * @param decimals Is the number of digits to the right of the decimal point.
     */
    usdollar(number: number | Range | RangeReference | FunctionResult<any>, decimals?: number | Range | RangeReference | FunctionResult<any>): FunctionResult<string>;
    /**
     *
     * Returns the Unicode character referenced by the given numeric value.
     *
     * [Api set: ExcelApi 1.2]
     *
     * @param number Is the Unicode number representing a character.
     */
    unichar(number: number | Range | RangeReference | FunctionResult<any>): FunctionResult<string>;
    /**
     *
     * Returns the number (code point) corresponding to the first character of the text.
     *
     * [Api set: ExcelApi 1.2]
     *
     * @param text Is the character that you want the Unicode value of.
     */
    unicode(text: string | Range | RangeReference | FunctionResult<any>): FunctionResult<number>;
    /**
     *
     * Converts a text string to all uppercase letters.
     *
     * [Api set: ExcelApi 1.2]
     *
     * @param text Is the text you want converted to uppercase, a reference or a text string.
     */
    upper(text: string | Range | RangeReference | FunctionResult<any>): FunctionResult<string>;
    /**
     *
     * Looks for a value in the leftmost column of a table, and then returns a value in the same row from a column you specify. By default, the table must be sorted in an ascending order.
     *
     * [Api set: ExcelApi 1.2]
     *
     * @param lookupValue Is the value to be found in the first column of the table, and can be a value, a reference, or a text string.
     * @param tableArray Is a table of text, numbers, or logical values, in which data is retrieved. tableArray can be a reference to a range or a range name.
     * @param colIndexNum Is the column number in tableArray from which the matching value should be returned. The first column of values in the table is column 1.
     * @param rangeLookup Is a logical value: to find the closest match in the first column (sorted in ascending order) = TRUE or omitted; find an exact match = FALSE.
     */
    vlookup(lookupValue: number | string | boolean | Range | RangeReference | FunctionResult<any>, tableArray: Range | number | RangeReference | FunctionResult<any>, colIndexNum: Range | number | RangeReference | FunctionResult<any>, rangeLookup?: boolean | Range | RangeReference | FunctionResult<any>): FunctionResult<number | string | boolean>;
    /**
     *
     * Converts a text string that represents a number to a number.
     *
     * [Api set: ExcelApi 1.2]
     *
     * @param text Is the text enclosed in quotation marks or a reference to a cell containing the text you want to convert.
     */
    value(text: string | boolean | number | Range | RangeReference | FunctionResult<any>): FunctionResult<number>;
    /**
     *
     * Estimates variance based on a sample, including logical values and text. Text and the logical value FALSE have the value 0; the logical value TRUE has the value 1.
     *
     * [Api set: ExcelApi 1.2]
     *
     * @param values List of parameters, whose elements are 1 to 255 value arguments corresponding to a sample of a population.
     */
    varA(...values: Array<number | Range | RangeReference | FunctionResult<any>>): FunctionResult<number>;
    /**
     *
     * Calculates variance based on the entire population, including logical values and text. Text and the logical value FALSE have the value 0; the logical value TRUE has the value 1.
     *
     * [Api set: ExcelApi 1.2]
     *
     * @param values List of parameters, whose elements are 1 to 255 value arguments corresponding to a population.
     */
    varPA(...values: Array<number | Range | RangeReference | FunctionResult<any>>): FunctionResult<number>;
    /**
     *
     * Calculates variance based on the entire population (ignores logical values and text in the population).
     *
     * [Api set: ExcelApi 1.2]
     *
     * @param values List of parameters, whose elements are 1 to 255 numeric arguments corresponding to a population.
     */
    var_P(...values: Array<number | Range | RangeReference | FunctionResult<any>>): FunctionResult<number>;
    /**
     *
     * Estimates variance based on a sample (ignores logical values and text in the sample).
     *
     * [Api set: ExcelApi 1.2]
     *
     * @param values List of parameters, whose elements are 1 to 255 numeric arguments corresponding to a sample of a population.
     */
    var_S(...values: Array<number | Range | RangeReference | FunctionResult<any>>): FunctionResult<number>;
    /**
     *
     * Returns the depreciation of an asset for any period you specify, including partial periods, using the double-declining balance method or some other method you specify.
     *
     * [Api set: ExcelApi 1.2]
     *
     * @param cost Is the initial cost of the asset.
     * @param salvage Is the salvage value at the end of the life of the asset.
     * @param life Is the number of periods over which the asset is being depreciated (sometimes called the useful life of the asset).
     * @param startPeriod Is the starting period for which you want to calculate the depreciation, in the same units as Life.
     * @param endPeriod Is the ending period for which you want to calculate the depreciation, in the same units as Life.
     * @param factor Is the rate at which the balance declines, 2 (double-declining balance) if omitted.
     * @param noSwitch Switch to straight-line depreciation when depreciation is greater than the declining balance = FALSE or omitted; do not switch = TRUE.
     */
    vdb(cost: number | Range | RangeReference | FunctionResult<any>, salvage: number | Range | RangeReference | FunctionResult<any>, life: number | Range | RangeReference | FunctionResult<any>, startPeriod: number | Range | RangeReference | FunctionResult<any>, endPeriod: number | Range | RangeReference | FunctionResult<any>, factor?: number | Range | RangeReference | FunctionResult<any>, noSwitch?: boolean | Range | RangeReference | FunctionResult<any>): FunctionResult<number>;
    /**
     *
     * Returns the week number in the year.
     *
     * [Api set: ExcelApi 1.2]
     *
     * @param serialNumber Is the date-time code used by Microsoft Excel for date and time calculation.
     * @param returnType Is a number (1 or 2) that determines the type of the return value.
     */
    weekNum(serialNumber: number | string | boolean | Range | RangeReference | FunctionResult<any>, returnType?: number | string | boolean | Range | RangeReference | FunctionResult<any>): FunctionResult<number>;
    /**
     *
     * Returns a number from 1 to 7 identifying the day of the week of a date.
     *
     * [Api set: ExcelApi 1.2]
     *
     * @param serialNumber Is a number that represents a date.
     * @param returnType Is a number: for Sunday=1 through Saturday=7, use 1; for Monday=1 through Sunday=7, use 2; for Monday=0 through Sunday=6, use 3.
     */
    weekday(serialNumber: number | Range | RangeReference | FunctionResult<any>, returnType?: number | Range | RangeReference | FunctionResult<any>): FunctionResult<number>;
    /**
     *
     * Returns the Weibull distribution.
     *
     * [Api set: ExcelApi 1.2]
     *
     * @param x Is the value at which to evaluate the function, a nonnegative number.
     * @param alpha Is a parameter to the distribution, a positive number.
     * @param beta Is a parameter to the distribution, a positive number.
     * @param cumulative Is a logical value: for the cumulative distribution function, use TRUE; for the probability mass function, use FALSE.
     */
    weibull_Dist(x: number | Range | RangeReference | FunctionResult<any>, alpha: number | Range | RangeReference | FunctionResult<any>, beta: number | Range | RangeReference | FunctionResult<any>, cumulative: boolean | Range | RangeReference | FunctionResult<any>): FunctionResult<number>;
    /**
     *
     * Returns the serial number of the date before or after a specified number of workdays.
     *
     * [Api set: ExcelApi 1.2]
     *
     * @param startDate Is a serial date number that represents the start date.
     * @param days Is the number of nonweekend and non-holiday days before or after startDate.
     * @param holidays Is an optional array of one or more serial date numbers to exclude from the working calendar, such as state and federal holidays and floating holidays.
     */
    workDay(startDate: number | string | boolean | Range | RangeReference | FunctionResult<any>, days: number | string | boolean | Range | RangeReference | FunctionResult<any>, holidays?: number | string | Range | boolean | RangeReference | FunctionResult<any>): FunctionResult<number>;
    /**
     *
     * Returns the serial number of the date before or after a specified number of workdays with custom weekend parameters.
     *
     * [Api set: ExcelApi 1.2]
     *
     * @param startDate Is a serial date number that represents the start date.
     * @param days Is the number of nonweekend and non-holiday days before or after startDate.
     * @param weekend Is a number or string specifying when weekends occur.
     * @param holidays Is an optional array of one or more serial date numbers to exclude from the working calendar, such as state and federal holidays and floating holidays.
     */
    workDay_Intl(startDate: number | string | boolean | Range | RangeReference | FunctionResult<any>, days: number | string | boolean | Range | RangeReference | FunctionResult<any>, weekend?: number | string | Range | RangeReference | FunctionResult<any>, holidays?: number | string | Range | boolean | RangeReference | FunctionResult<any>): FunctionResult<number>;
    /**
     *
     * Returns the internal rate of return for a schedule of cash flows.
     *
     * [Api set: ExcelApi 1.2]
     *
     * @param values Is a series of cash flows that correspond to a schedule of payments in dates.
     * @param dates Is a schedule of payment dates that corresponds to the cash flow payments.
     * @param guess Is a number that you guess is close to the result of XIRR.
     */
    xirr(values: number | string | Range | boolean | RangeReference | FunctionResult<any>, dates: number | string | Range | boolean | RangeReference | FunctionResult<any>, guess?: number | string | boolean | Range | RangeReference | FunctionResult<any>): FunctionResult<number>;
    /**
     *
     * Returns the net present value for a schedule of cash flows.
     *
     * [Api set: ExcelApi 1.2]
     *
     * @param rate Is the discount rate to apply to the cash flows.
     * @param values Is a series of cash flows that correspond to a schedule of payments in dates.
     * @param dates Is a schedule of payment dates that corresponds to the cash flow payments.
     */
    xnpv(rate: number | string | boolean | Range | RangeReference | FunctionResult<any>, values: number | string | Range | boolean | RangeReference | FunctionResult<any>, dates: number | string | Range | boolean | RangeReference | FunctionResult<any>): FunctionResult<number>;
    /**
     *
     * Returns a logical 'Exclusive Or' of all arguments.
     *
     * [Api set: ExcelApi 1.2]
     *
     * @param values List of parameters, whose elements are 1 to 254 conditions you want to test that can be either TRUE or FALSE and can be logical values, arrays, or references.
     */
    xor(...values: Array<boolean | Range | RangeReference | FunctionResult<any>>): FunctionResult<boolean>;
    /**
     *
     * Returns the year of a date, an integer in the range 1900 - 9999.
     *
     * [Api set: ExcelApi 1.2]
     *
     * @param serialNumber Is a number in the date-time code used by Microsoft
     */
    year(serialNumber: number | Range | RangeReference | FunctionResult<any>): FunctionResult<number>;
    /**
     *
     * Returns the year fraction representing the number of whole days between start_date and end_date.
     *
     * [Api set: ExcelApi 1.2]
     *
     * @param startDate Is a serial date number that represents the start date.
     * @param endDate Is a serial date number that represents the end date.
     * @param basis Is the type of day count basis to use.
     */
    yearFrac(startDate: number | string | boolean | Range | RangeReference | FunctionResult<any>, endDate: number | string | boolean | Range | RangeReference | FunctionResult<any>, basis?: number | string | boolean | Range | RangeReference | FunctionResult<any>): FunctionResult<number>;
    /**
     *
     * Returns the yield on a security that pays periodic interest.
     *
     * [Api set: ExcelApi 1.2]
     *
     * @param settlement Is the security's settlement date, expressed as a serial date number.
     * @param maturity Is the security's maturity date, expressed as a serial date number.
     * @param rate Is the security's annual coupon rate.
     * @param pr Is the security's price per $100 face value.
     * @param redemption Is the security's redemption value per $100 face value.
     * @param frequency Is the number of coupon payments per year.
     * @param basis Is the type of day count basis to use.
     */
    yield(settlement: number | string | boolean | Range | RangeReference | FunctionResult<any>, maturity: number | string | boolean | Range | RangeReference | FunctionResult<any>, rate: number | string | boolean | Range | RangeReference | FunctionResult<any>, pr: number | string | boolean | Range | RangeReference | FunctionResult<any>, redemption: number | string | boolean | Range | RangeReference | FunctionResult<any>, frequency: number | string | boolean | Range | RangeReference | FunctionResult<any>, basis?: number | string | boolean | Range | RangeReference | FunctionResult<any>): FunctionResult<number>;
    /**
     *
     * Returns the annual yield for a discounted security. For example, a treasury bill.
     *
     * [Api set: ExcelApi 1.2]
     *
     * @param settlement Is the security's settlement date, expressed as a serial date number.
     * @param maturity Is the security's maturity date, expressed as a serial date number.
     * @param pr Is the security's price per $100 face value.
     * @param redemption Is the security's redemption value per $100 face value.
     * @param basis Is the type of day count basis to use.
     */
    yieldDisc(settlement: number | string | boolean | Range | RangeReference | FunctionResult<any>, maturity: number | string | boolean | Range | RangeReference | FunctionResult<any>, pr: number | string | boolean | Range | RangeReference | FunctionResult<any>, redemption: number | string | boolean | Range | RangeReference | FunctionResult<any>, basis?: number | string | boolean | Range | RangeReference | FunctionResult<any>): FunctionResult<number>;
    /**
     *
     * Returns the annual yield of a security that pays interest at maturity.
     *
     * [Api set: ExcelApi 1.2]
     *
     * @param settlement Is the security's settlement date, expressed as a serial date number.
     * @param maturity Is the security's maturity date, expressed as a serial date number.
     * @param issue Is the security's issue date, expressed as a serial date number.
     * @param rate Is the security's interest rate at date of issue.
     * @param pr Is the security's price per $100 face value.
     * @param basis Is the type of day count basis to use.
     */
    yieldMat(settlement: number | string | boolean | Range | RangeReference | FunctionResult<any>, maturity: number | string | boolean | Range | RangeReference | FunctionResult<any>, issue: number | string | boolean | Range | RangeReference | FunctionResult<any>, rate: number | string | boolean | Range | RangeReference | FunctionResult<any>, pr: number | string | boolean | Range | RangeReference | FunctionResult<any>, basis?: number | string | boolean | Range | RangeReference | FunctionResult<any>): FunctionResult<number>;
    /**
     *
     * Returns the one-tailed P-value of a z-test.
     *
     * [Api set: ExcelApi 1.2]
     *
     * @param array Is the array or range of data against which to test X.
     * @param x Is the value to test.
     * @param sigma Is the population (known) standard deviation. If omitted, the sample standard deviation is used.
     */
    z_Test(array: number | Range | RangeReference | FunctionResult<any>, x: number | Range | RangeReference | FunctionResult<any>, sigma?: number | Range | RangeReference | FunctionResult<any>): FunctionResult<number>;
    toJSON(): {};
}
export declare namespace ErrorCodes {
    var accessDenied: string;
    var apiNotFound: string;
    var generalException: string;
    var insertDeleteConflict: string;
    var invalidArgument: string;
    var invalidBinding: string;
    var invalidOperation: string;
    var invalidReference: string;
    var invalidSelection: string;
    var itemAlreadyExists: string;
    var itemNotFound: string;
    var notImplemented: string;
    var unsupportedOperation: string;
}

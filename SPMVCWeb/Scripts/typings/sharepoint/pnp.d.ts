/// <reference path="../microsoft-ajax/microsoft.ajax.d.ts" />
/// <reference path="../es6-promise/es6-promise.d.ts" />
/// <reference path="../fetch/fetch.d.ts" />

declare module "utils/util" {
    export class Util {
        /**
         * Gets a callback function which will maintain context across async calls.
         * Allows for the calling pattern getCtxCallback(thisobj, method, methodarg1, methodarg2, ...)
         *
         * @param context The object that will be the 'this' value in the callback
         * @param method The method to which we will apply the context and parameters
         * @param params Optional, additional arguments to supply to the wrapped method when it is invoked
         */
        static getCtxCallback(context: any, method: Function, ...params: any[]): Function;
        /**
         * Tests if a url param exists
         *
         * @param name The name of the url paramter to check
         */
        static urlParamExists(name: string): boolean;
        /**
         * Gets a url param value by name
         *
         * @param name The name of the paramter for which we want the value
         */
        static getUrlParamByName(name: string): string;
        /**
         * Gets a url param by name and attempts to parse a bool value
         *
         * @param name The name of the paramter for which we want the boolean value
         */
        static getUrlParamBoolByName(name: string): boolean;
        /**
         * Inserts the string s into the string target as the index specified by index
         *
         * @param target The string into which we will insert s
         * @param index The location in target to insert s (zero based)
         * @param s The string to insert into target at position index
         */
        static stringInsert(target: string, index: number, s: string): string;
        /**
         * Adds a value to a date
         *
         * @param date The date to which we will add units, done in local time
         * @param interval The name of the interval to add, one of: ['year', 'quarter', 'month', 'week', 'day', 'hour', 'minute', 'second']
         * @param units The amount to add to date of the given interval
         *
         * http://stackoverflow.com/questions/1197928/how-to-add-30-minutes-to-a-javascript-date-object
         */
        static dateAdd(date: Date, interval: string, units: number): Date;
        /**
         * Loads a stylesheet into the current page
         *
         * @param path The url to the stylesheet
         * @param avoidCache If true a value will be appended as a query string to avoid browser caching issues
         */
        static loadStylesheet(path: string, avoidCache: boolean): void;
        /**
         * Combines an arbitrary set of paths ensuring that the slashes are normalized
         *
         * @param paths 0 to n path parts to combine
         */
        static combinePaths(...paths: string[]): string;
        /**
         * Gets a random string of chars length
         *
         * @param chars The length of the random string to generate
         */
        static getRandomString(chars: number): string;
        /**
         * Gets a random GUID value
         *
         * http://stackoverflow.com/questions/105034/create-guid-uuid-in-javascript
         */
        static getGUID(): string;
        /**
         * Determines if a given value is a function
         *
         * @param candidateFunction The thing to test for being a function
         */
        static isFunction(candidateFunction: any): boolean;
        /**
         * @returns whether the provided parameter is a JavaScript Array or not.
        */
        static isArray(array: any): boolean;
        /**
         * Determines if a string is null or empty or undefined
         *
         * @param s The string to test
         */
        static stringIsNullOrEmpty(s: string): boolean;
        /**
         * Provides functionality to extend the given object by doing a shallow copy
         *
         * @param target The object to which properties will be copied
         * @param source The source object from which properties will be copied
         * @param noOverwrite If true existing properties on the target are not overwritten from the source
         *
         */
        static extend<T, S>(target: T, source: S, noOverwrite?: boolean): T & S;
        /**
         * Applies one or more mixins to the supplied target
         *
         * @param derivedCtor The classto which we will apply the mixins
         * @param baseCtors One or more mixin classes to apply
         */
        static applyMixins(derivedCtor: any, ...baseCtors: any[]): void;
        /**
         * Determines if a given url is absolute
         *
         * @param url The url to check to see if it is absolute
         */
        static isUrlAbsolute(url: string): boolean;
        /**
         * Attempts to make the supplied relative url absolute based on the _spPageContextInfo object, if available
         *
         * @param url The relative url to make absolute
         */
        static makeUrlAbsolute(url: string): string;
    }
}
declare module "utils/storage" {
    /**
     * A wrapper class to provide a consistent interface to browser based storage
     *
     */
    export class PnPClientStorageWrapper implements PnPClientStore {
        private store;
        defaultTimeoutMinutes: number;
        /**
         * True if the wrapped storage is available; otherwise, false
         */
        enabled: boolean;
        /**
         * Creates a new instance of the PnPClientStorageWrapper class
         *
         * @constructor
         */
        constructor(store: Storage, defaultTimeoutMinutes?: number);
        /**
         * Get a value from storage, or null if that value does not exist
         *
         * @param key The key whose value we want to retrieve
         */
        get<T>(key: string): T;
        /**
         * Adds a value to the underlying storage
         *
         * @param key The key to use when storing the provided value
         * @param o The value to store
         * @param expire Optional, if provided the expiration of the item, otherwise the default is used
         */
        put(key: string, o: any, expire?: Date): void;
        /**
         * Deletes a value from the underlying storage
         *
         * @param key The key of the pair we want to remove from storage
         */
        delete(key: string): void;
        /**
         * Gets an item from the underlying storage, or adds it if it does not exist using the supplied getter function
         *
         * @param key The key to use when storing the provided value
         * @param getter A function which will upon execution provide the desired value
         * @param expire Optional, if provided the expiration of the item, otherwise the default is used
         */
        getOrPut<T>(key: string, getter: () => Promise<T>, expire?: Date): Promise<T>;
        /**
         * Used to determine if the wrapped storage is available currently
         */
        private test();
        /**
         * Creates the persistable to store
         */
        private createPersistable(o, expire?);
    }
    /**
     * Interface which defines the operations provided by a client storage object
     */
    export interface PnPClientStore {
        /**
         * True if the wrapped storage is available; otherwise, false
         */
        enabled: boolean;
        /**
         * Get a value from storage, or null if that value does not exist
         *
         * @param key The key whose value we want to retrieve
         */
        get(key: string): any;
        /**
         * Adds a value to the underlying storage
         *
         * @param key The key to use when storing the provided value
         * @param o The value to store
         * @param expire Optional, if provided the expiration of the item, otherwise the default is used
         */
        put(key: string, o: any, expire?: Date): void;
        /**
         * Deletes a value from the underlying storage
         *
         * @param key The key of the pair we want to remove from storage
         */
        delete(key: string): void;
        /**
         * Gets an item from the underlying storage, or adds it if it does not exist using the supplied getter function
         *
         * @param key The key to use when storing the provided value
         * @param getter A function which will upon execution provide the desired value
         * @param expire Optional, if provided the expiration of the item, otherwise the default is used
         */
        getOrPut(key: string, getter: Function, expire?: Date): any;
    }
    /**
     * A class that will establish wrappers for both local and session storage
     */
    export class PnPClientStorage {
        /**
         * Provides access to the local storage of the browser
         */
        local: PnPClientStore;
        /**
         * Provides access to the session storage of the browser
         */
        session: PnPClientStore;
        /**
         * Creates a new instance of the PnPClientStorage class
         *
         * @constructor
         */
        constructor();
    }
}
declare module "collections/collections" {
    /**
     * Interface defining an object with a known property type
     */
    export interface TypedHash<T> {
        [key: string]: T;
    }
    /**
     * Generic dictionary
     */
    export class Dictionary<T> {
        /**
         * The array used to store all the keys
         */
        private keys;
        /**
         * The array used to store all the values
         */
        private values;
        /**
         * Creates a new instance of the Dictionary<T> class
         *
         * @constructor
         */
        constructor();
        /**
         * Gets a value from the collection using the specified key
         *
         * @param key The key whose value we want to return, returns null if the key does not exist
         */
        get(key: string): T;
        /**
         * Adds the supplied key and value to the dictionary
         *
         * @param key The key to add
         * @param o The value to add
         */
        add(key: string, o: T): void;
        /**
         * Merges the supplied typed hash into this dictionary instance. Existing values are updated and new ones are created as appropriate.
         */
        merge(source: TypedHash<T> | Dictionary<T>): void;
        /**
         * Removes a value from the dictionary
         *
         * @param key The key of the key/value pair to remove. Returns null if the key was not found.
         */
        remove(key: string): T;
        /**
         * Returns all the keys currently in the dictionary as an array
         */
        getKeys(): string[];
        /**
         * Returns all the values currently in the dictionary as an array
         */
        getValues(): T[];
        /**
         * Clears the current dictionary
         */
        clear(): void;
        /**
         * Gets a count of the items currently in the dictionary
         */
        count(): number;
    }
}
declare module "configuration/providers/cachingConfigurationProvider" {
    import { IConfigurationProvider } from "configuration/configuration";
    import { TypedHash } from "collections/collections";
    import * as storage from "utils/storage";
    /**
     * A caching provider which can wrap other non-caching providers
     *
     */
    export default class CachingConfigurationProvider implements IConfigurationProvider {
        private wrappedProvider;
        private store;
        private cacheKey;
        /**
         * Creates a new caching configuration provider
         * @constructor
         * @param {IConfigurationProvider} wrappedProvider Provider which will be used to fetch the configuration
         * @param {string} cacheKey Key that will be used to store cached items to the cache
         * @param {IPnPClientStore} cacheStore OPTIONAL storage, which will be used to store cached settings.
         */
        constructor(wrappedProvider: IConfigurationProvider, cacheKey: string, cacheStore?: storage.PnPClientStore);
        /**
         * Gets the wrapped configuration providers
         *
         * @return {IConfigurationProvider} Wrapped configuration provider
         */
        getWrappedProvider(): IConfigurationProvider;
        /**
         * Loads the configuration values either from the cache or from the wrapped provider
         *
         * @return {Promise<TypedHash<string>>} Promise of loaded configuration values
         */
        getConfiguration(): Promise<TypedHash<string>>;
        private selectPnPCache();
    }
}
declare module "utils/logging" {
    /**
     * A set of logging levels
     *
     */
    export enum LogLevel {
        Verbose = 0,
        Info = 1,
        Warning = 2,
        Error = 3,
        Off = 99,
    }
    /**
     * Interface that defines a log entry
     *
     */
    export interface LogEntry {
        /**
         * The main message to be logged
         */
        message: string;
        /**
         * The level of information this message represents
         */
        level: LogLevel;
        /**
         * Any associated data that a given logging listener may choose to log or ignore
         */
        data?: any;
    }
    /**
     * Interface that defines a log listner
     *
     */
    export interface LogListener {
        /**
         * Any associated data that a given logging listener may choose to log or ignore
         *
         * @param entry The information to be logged
         */
        log(entry: LogEntry): void;
    }
    /**
     * Class used to subscribe ILogListener and log messages throughout an application
     *
     */
    export class Logger {
        private static _instance;
        static activeLogLevel: LogLevel;
        private static readonly instance;
        /**
         * Adds an ILogListener instance to the set of subscribed listeners
         *
         * @param listeners One or more listeners to subscribe to this log
         */
        static subscribe(...listeners: LogListener[]): void;
        /**
         * Clears the subscribers collection, returning the collection before modifiction
         */
        static clearSubscribers(): LogListener[];
        /**
         * Gets the current subscriber count
         */
        static readonly count: number;
        /**
         * Writes the supplied string to the subscribed listeners
         *
         * @param message The message to write
         * @param level [Optional] if supplied will be used as the level of the entry (Default: LogLevel.Verbose)
         */
        static write(message: string, level?: LogLevel): void;
        /**
         * Logs the supplied entry to the subscribed listeners
         *
         * @param entry The message to log
         */
        static log(entry: LogEntry): void;
        /**
         * Logs performance tracking data for the the execution duration of the supplied function using console.profile
         *
         * @param name The name of this profile boundary
         * @param f The function to execute and track within this performance boundary
         */
        static measure<T>(name: string, f: () => T): T;
    }
    /**
     * Implementation of ILogListener which logs to the browser console
     *
     */
    export class ConsoleListener implements LogListener {
        /**
         * Any associated data that a given logging listener may choose to log or ignore
         *
         * @param entry The information to be logged
         */
        log(entry: LogEntry): void;
        /**
         * Formats the message
         *
         * @param entry The information to format into a string
         */
        private format(entry);
    }
    /**
     * Implementation of ILogListener which logs to Azure Insights
     *
     */
    export class AzureInsightsListener implements LogListener {
        private azureInsightsInstrumentationKey;
        /**
         * Creats a new instance of the AzureInsightsListener class
         *
         * @constructor
         * @param azureInsightsInstrumentationKey The instrumentation key created when the Azure Insights instance was created
         */
        constructor(azureInsightsInstrumentationKey: string);
        /**
         * Any associated data that a given logging listener may choose to log or ignore
         *
         * @param entry The information to be logged
         */
        log(entry: LogEntry): void;
        /**
         * Formats the message
         *
         * @param entry The information to format into a string
         */
        private format(entry);
    }
    /**
     * Implementation of ILogListener which logs to the supplied function
     *
     */
    export class FunctionListener implements LogListener {
        private method;
        /**
         * Creates a new instance of the FunctionListener class
         *
         * @constructor
         * @param  method The method to which any logging data will be passed
         */
        constructor(method: (entry: LogEntry) => void);
        /**
         * Any associated data that a given logging listener may choose to log or ignore
         *
         * @param entry The information to be logged
         */
        log(entry: LogEntry): void;
    }
}
declare module "net/fetchclient" {
    import { HttpClientImpl } from "net/httpclient";
    /**
     * Makes requests using the fetch API
     */
    export class FetchClient implements HttpClientImpl {
        fetch(url: string, options: any): Promise<Response>;
    }
}
declare module "configuration/pnplibconfig" {
    import { TypedHash } from "collections/collections";
    export interface NodeClientData {
        clientId: string;
        clientSecret: string;
        siteUrl: string;
    }
    export interface LibraryConfiguration {
        /**
         * Any headers to apply to all requests
         */
        headers?: TypedHash<string>;
        /**
         * Allows caching to be global disabled, default: false
         */
        globalCacheDisable?: boolean;
        /**
         * Defines the default store used by the usingCaching method, default: session
         */
        defaultCachingStore?: "session" | "local";
        /**
         * Defines the default timeout in seconds used by the usingCaching method, default 30
         */
        defaultCachingTimeoutSeconds?: number;
        /**
         * If true the SP.RequestExecutor will be used to make the requests, you must include the required external libs
         */
        useSPRequestExecutor?: boolean;
        /**
         * If set the library will use node-fetch, typically for use with testing but works with any valid client id/secret pair
         */
        nodeClientOptions?: NodeClientData;
    }
    export class RuntimeConfigImpl {
        private _headers;
        private _defaultCachingStore;
        private _defaultCachingTimeoutSeconds;
        private _globalCacheDisable;
        private _useSPRequestExecutor;
        private _useNodeClient;
        private _nodeClientData;
        constructor();
        set(config: LibraryConfiguration): void;
        readonly headers: TypedHash<string>;
        readonly defaultCachingStore: "session" | "local";
        readonly defaultCachingTimeoutSeconds: number;
        readonly globalCacheDisable: boolean;
        readonly useSPRequestExecutor: boolean;
        readonly useNodeFetchClient: boolean;
        readonly nodeRequestOptions: NodeClientData;
    }
    export let RuntimeConfig: RuntimeConfigImpl;
    export function setRuntimeConfig(config: LibraryConfiguration): void;
}
declare module "sharepoint/rest/odata" {
    import { QueryableConstructor } from "sharepoint/rest/queryable";
    export function extractOdataId(candidate: any): string;
    export interface ODataParser<T, U> {
        parse(r: Response): Promise<U>;
    }
    export abstract class ODataParserBase<T, U> implements ODataParser<T, U> {
        parse(r: Response): Promise<U>;
        protected parseODataJSON<U>(json: any): U;
    }
    export class ODataDefaultParser extends ODataParserBase<any, any> {
    }
    export class ODataRawParserImpl implements ODataParser<any, any> {
        parse(r: Response): Promise<any>;
    }
    export let ODataRaw: ODataRawParserImpl;
    export function ODataValue<T>(): ODataParser<any, T>;
    export function ODataEntity<T>(factory: QueryableConstructor<T>): ODataParser<T, T>;
    export function ODataEntityArray<T>(factory: QueryableConstructor<T>): ODataParser<T, T[]>;
    /**
     * Manages a batch of OData operations
     */
    export class ODataBatch {
        private baseUrl;
        private _batchId;
        private _batchDependencies;
        private _requests;
        constructor(baseUrl: string, _batchId?: string);
        /**
         * Adds a request to a batch (not designed for public use)
         *
         * @param url The full url of the request
         * @param method The http method GET, POST, etc
         * @param options Any options to include in the request
         * @param parser The parser that will hadle the results of the request
         */
        add<U>(url: string, method: string, options: any, parser: ODataParser<any, U>): Promise<U>;
        addBatchDependency(): () => void;
        /**
         * Execute the current batch and resolve the associated promises
         *
         * @returns A promise which will be resolved once all of the batch's child promises have resolved
         */
        execute(): Promise<void>;
        private executeImpl();
        /**
         * Parses the response from a batch request into an array of Response instances
         *
         * @param body Text body of the response from the batch request
         */
        private _parseResponse(body);
    }
}
declare module "net/digestcache" {
    import { Dictionary } from "collections/collections";
    import { HttpClient } from "net/httpclient";
    export class CachedDigest {
        expiration: Date;
        value: string;
    }
    export class DigestCache {
        private _httpClient;
        private _digests;
        constructor(_httpClient: HttpClient, _digests?: Dictionary<CachedDigest>);
        getDigest(webUrl: string): Promise<string>;
        clear(): void;
    }
}
declare module "net/sprequestexecutorclient" {
    import { HttpClientImpl } from "net/httpclient";
    /**
     * Makes requests using the SP.RequestExecutor library.
     */
    export class SPRequestExecutorClient implements HttpClientImpl {
        /**
         * Fetches a URL using the SP.RequestExecutor library.
         */
        fetch(url: string, options: any): Promise<Response>;
        /**
         * Converts a SharePoint REST API response to a fetch API response.
         */
        private convertToResponse;
    }
}
declare module "net/nodefetchclient" {
    import { HttpClientImpl } from "net/httpclient";
    export interface AuthToken {
        token_type: string;
        expires_in: string;
        not_before: string;
        expires_on: string;
        resource: string;
        access_token: string;
    }
    /**
     * Fetch client for use within nodejs, requires you register a client id and secret with app only permissions
     */
    export class NodeFetchClient implements HttpClientImpl {
        siteUrl: string;
        private _clientId;
        private _clientSecret;
        private _realm;
        private static SharePointServicePrincipal;
        private token;
        constructor(siteUrl: string, _clientId: string, _clientSecret: string, _realm?: string);
        fetch(url: string, options: any): Promise<Response>;
        /**
         * Gets an add-in only authentication token based on the supplied site url, client id and secret
         */
        getAddInOnlyAccessToken(): Promise<AuthToken>;
        private getRealm();
        private getAuthUrl(realm);
        private getFormattedPrincipal(principalName, hostName, realm);
        private toDate(epoch);
    }
}
declare module "net/httpclient" {
    export interface FetchOptions {
        method?: string;
        headers?: any/*HeadersInit*/ | {
            [index: string]: string;
        };
        body?: BodyInit;
        mode?: string | RequestMode;
        credentials?: string | RequestCredentials;
        cache?: string | RequestCache;
    }
    export class HttpClient {
        private _digestCache;
        private _impl;
        constructor();
        fetch(url: string, options?: FetchOptions): Promise<Response>;
        fetchRaw(url: string, options?: FetchOptions): Promise<Response>;
        get(url: string, options?: FetchOptions): Promise<Response>;
        post(url: string, options?: FetchOptions): Promise<Response>;
        patch(url: string, options?: FetchOptions): Promise<Response>;
        delete(url: string, options?: FetchOptions): Promise<Response>;
        protected getFetchImpl(): HttpClientImpl;
        private mergeHeaders(target, source);
    }
    export interface HttpClientImpl {
        fetch(url: string, options: any): Promise<Response>;
    }
}
declare module "sharepoint/rest/caching" {
    import { ODataParser } from "sharepoint/rest/odata";
    import { PnPClientStore, PnPClientStorage } from "utils/storage";
    export interface ICachingOptions {
        expiration?: Date;
        storeName?: "session" | "local";
        key: string;
    }
    export class CachingOptions implements ICachingOptions {
        key: string;
        protected static storage: PnPClientStorage;
        expiration: Date;
        storeName: "session" | "local";
        constructor(key: string);
        readonly store: PnPClientStore;
    }
    export class CachingParserWrapper<T, U> implements ODataParser<T, U> {
        private _parser;
        private _cacheOptions;
        constructor(_parser: ODataParser<T, U>, _cacheOptions: CachingOptions);
        parse(response: Response): Promise<U>;
    }
}
declare module "sharepoint/rest/queryable" {
    import { Dictionary } from "collections/collections";
    import { FetchOptions } from "net/httpclient";
    import { ODataParser, ODataBatch } from "sharepoint/rest/odata";
    import { ICachingOptions } from "sharepoint/rest/caching";
    export interface QueryableConstructor<T> {
        new (baseUrl: string | Queryable, path?: string): T;
    }
    /**
     * Queryable Base Class
     *
     */
    export class Queryable {
        /**
         * Tracks the query parts of the url
         */
        protected _query: Dictionary<string>;
        /**
         * Tracks the batch of which this query may be part
         */
        private _batch;
        /**
         * Tracks the url as it is built
         */
        private _url;
        /**
         * Stores the parent url used to create this instance, for recursing back up the tree if needed
         */
        private _parentUrl;
        /**
         * Explicitly tracks if we are using caching for this request
         */
        private _useCaching;
        /**
         * Any options that were supplied when caching was enabled
         */
        private _cachingOptions;
        /**
         * Directly concatonates the supplied string to the current url, not normalizing "/" chars
         *
         * @param pathPart The string to concatonate to the url
         */
        concat(pathPart: string): void;
        /**
         * Appends the given string and normalizes "/" chars
         *
         * @param pathPart The string to append
         */
        protected append(pathPart: string): void;
        /**
         * Blocks a batch call from occuring, MUST be cleared by calling the returned function
         */
        protected addBatchDependency(): () => void;
        /**
         * Indicates if the current query has a batch associated
         *
         */
        protected readonly hasBatch: boolean;
        /**
         * Gets the parent url used when creating this instance
         *
         */
        protected readonly parentUrl: string;
        /**
         * Provides access to the query builder for this url
         *
         */
        readonly query: Dictionary<string>;
        /**
         * Creates a new instance of the Queryable class
         *
         * @constructor
         * @param baseUrl A string or Queryable that should form the base part of the url
         *
         */
        constructor(baseUrl: string | Queryable, path?: string);
        /**
         * Adds this query to the supplied batch
         *
         * @example
         * ```
         *
         * let b = pnp.sp.createBatch();
         * pnp.sp.web.inBatch(b).get().then(...);
         * b.execute().then(...)
         * ```
         */
        inBatch(batch: ODataBatch): this;
        /**
         * Enables caching for this request
         *
         * @param options Defines the options used when caching this request
         */
        usingCaching(options?: ICachingOptions): this;
        /**
         * Gets the currentl url, made absolute based on the availability of the _spPageContextInfo object
         *
         */
        toUrl(): string;
        /**
         * Gets the full url with query information
         *
         */
        toUrlAndQuery(): string;
        /**
         * Executes the currently built request
         *
         */
        get(parser?: ODataParser<any, any>, getOptions?: FetchOptions): Promise<any>;
        getAs<T, U>(parser?: ODataParser<T, U>, getOptions?: FetchOptions): Promise<U>;
        protected post(postOptions?: FetchOptions, parser?: ODataParser<any, any>): Promise<any>;
        protected postAs<T, U>(postOptions?: FetchOptions, parser?: ODataParser<T, U>): Promise<U>;
        protected patch(patchOptions?: FetchOptions, parser?: ODataParser<any, any>): Promise<any>;
        protected delete(deleteOptions?: FetchOptions, parser?: ODataParser<any, any>): Promise<any>;
        /**
         * Gets a parent for this instance as specified
         *
         * @param factory The contructor for the class to create
         */
        protected getParent<T extends Queryable>(factory: {
            new (q: string | Queryable, path?: string): T;
        }, baseUrl?: string | Queryable, path?: string): T;
        private getImpl<U>(getOptions, parser);
        private postImpl<U>(postOptions, parser);
        private patchImpl<U>(patchOptions, parser);
        private deleteImpl<U>(deleteOptions, parser);
        private processHttpClientResponse<U>(response, parser);
    }
    /**
     * Represents a REST collection which can be filtered, paged, and selected
     *
     */
    export class QueryableCollection extends Queryable {
        /**
         * Filters the returned collection (https://msdn.microsoft.com/en-us/library/office/fp142385.aspx#bk_supported)
         *
         * @param filter The string representing the filter query
         */
        filter(filter: string): this;
        /**
         * Choose which fields to return
         *
         * @param selects One or more fields to return
         */
        select(...selects: string[]): this;
        /**
         * Expands fields such as lookups to get additional data
         *
         * @param expands The Fields for which to expand the values
         */
        expand(...expands: string[]): this;
        /**
         * Orders based on the supplied fields ascending
         *
         * @param orderby The name of the field to sort on
         * @param ascending If false DESC is appended, otherwise ASC (default)
         */
        orderBy(orderBy: string, ascending?: boolean): this;
        /**
         * Skips the specified number of items
         *
         * @param skip The number of items to skip
         */
        skip(skip: number): this;
        /**
         * Limits the query to only return the specified number of items
         *
         * @param top The query row limit
         */
        top(top: number): this;
    }
    /**
     * Represents an instance that can be selected
     *
     */
    export class QueryableInstance extends Queryable {
        /**
         * Choose which fields to return
         *
         * @param selects One or more fields to return
         */
        select(...selects: string[]): this;
        /**
         * Expands fields such as lookups to get additional data
         *
         * @param expands The Fields for which to expand the values
         */
        expand(...expands: string[]): this;
    }
}
declare module "sharepoint/rest/types" {
    /**
     * Represents the unique sequential location of a change within the change log.
     */
    export interface ChangeToken {
        /**
         * Gets or sets a string value that contains the serialized representation of the change token generated by the protocol server.
         */
        StringValue: string;
    }
    /**
     * Defines a query that is performed against the change log.
     */
    export interface ChangeQuery {
        /**
         * Gets or sets a value that specifies whether add changes are included in the query.
         */
        Add?: boolean;
        /**
         * Gets or sets a value that specifies whether changes to alerts are included in the query.
         */
        Alert?: boolean;
        /**
         * Gets or sets a value that specifies the end date and end time for changes that are returned through the query.
         */
        ChangeTokenEnd?: ChangeToken;
        /**
         * Gets or sets a value that specifies the start date and start time for changes that are returned through the query.
         */
        ChangeTokenStart?: ChangeToken;
        /**
         * Gets or sets a value that specifies whether changes to content types are included in the query.
         */
        ContentType?: boolean;
        /**
         * Gets or sets a value that specifies whether deleted objects are included in the query.
         */
        DeleteObject?: boolean;
        /**
         * Gets or sets a value that specifies whether changes to fields are included in the query.
         */
        Field?: boolean;
        /**
         * Gets or sets a value that specifies whether changes to files are included in the query.
         */
        File?: boolean;
        /**
         * Gets or sets value that specifies whether changes to folders are included in the query.
         */
        Folder?: boolean;
        /**
         * Gets or sets a value that specifies whether changes to groups are included in the query.
         */
        Group?: boolean;
        /**
         * Gets or sets a value that specifies whether adding users to groups is included in the query.
         */
        GroupMembershipAdd?: boolean;
        /**
         * Gets or sets a value that specifies whether deleting users from the groups is included in the query.
         */
        GroupMembershipDelete?: boolean;
        /**
         * Gets or sets a value that specifies whether general changes to list items are included in the query.
         */
        Item?: boolean;
        /**
         * Gets or sets a value that specifies whether changes to lists are included in the query.
         */
        List?: boolean;
        /**
         * Gets or sets a value that specifies whether move changes are included in the query.
         */
        Move?: boolean;
        /**
         * Gets or sets a value that specifies whether changes to the navigation structure of a site collection are included in the query.
         */
        Navigation?: boolean;
        /**
         * Gets or sets a value that specifies whether renaming changes are included in the query.
         */
        Rename?: boolean;
        /**
         * Gets or sets a value that specifies whether restoring items from the recycle bin or from backups is included in the query.
         */
        Restore?: boolean;
        /**
         * Gets or sets a value that specifies whether adding role assignments is included in the query.
         */
        RoleAssignmentAdd?: boolean;
        /**
         * Gets or sets a value that specifies whether adding role assignments is included in the query.
         */
        RoleAssignmentDelete?: boolean;
        /**
         * Gets or sets a value that specifies whether adding role assignments is included in the query.
         */
        RoleDefinitionAdd?: boolean;
        /**
         * Gets or sets a value that specifies whether adding role assignments is included in the query.
         */
        RoleDefinitionDelete?: boolean;
        /**
         * Gets or sets a value that specifies whether adding role assignments is included in the query.
         */
        RoleDefinitionUpdate?: boolean;
        /**
         * Gets or sets a value that specifies whether modifications to security policies are included in the query.
         */
        SecurityPolicy?: boolean;
        /**
         * Gets or sets a value that specifies whether changes to site collections are included in the query.
         */
        Site?: boolean;
        /**
         * Gets or sets a value that specifies whether updates made using the item SystemUpdate method are included in the query.
         */
        SystemUpdate?: boolean;
        /**
         * Gets or sets a value that specifies whether update changes are included in the query.
         */
        Update?: boolean;
        /**
         * Gets or sets a value that specifies whether changes to users are included in the query.
         */
        User?: boolean;
        /**
         * Gets or sets a value that specifies whether changes to views are included in the query.
         */
        View?: boolean;
        /**
         * Gets or sets a value that specifies whether changes to Web sites are included in the query.
         */
        Web?: boolean;
    }
    /**
     * Specifies a Collaborative Application Markup Language (CAML) query on a list or joined lists.
     */
    export interface CamlQuery {
        /**
         * Gets or sets a value that indicates whether the query returns dates in Coordinated Universal Time (UTC) format.
         */
        DatesInUtc?: boolean;
        /**
         * Gets or sets a value that specifies the server relative URL of a list folder from which results will be returned.
         */
        FolderServerRelativeUrl?: string;
        /**
         * Gets or sets a value that specifies the information required to get the next page of data for the list view.
         */
        ListItemCollectionPosition?: ListItemCollectionPosition;
        /**
         * Gets or sets value that specifies the XML schema that defines the list view.
         */
        ViewXml?: string;
    }
    /**
     * Specifies the information required to get the next page of data for a list view.
     */
    export interface ListItemCollectionPosition {
        /**
         * Gets or sets a value that specifies information, as name-value pairs, required to get the next page of data for a list view.
         */
        PagingInfo: string;
    }
    /**
     * Represents the input parameter of the GetListItemChangesSinceToken method.
     */
    export interface ChangeLogitemQuery {
        /**
         * The change token for the request.
         */
        ChangeToken?: string;
        /**
         * The XML element that defines custom filtering for the query.
         */
        Contains?: string;
        /**
         * The records from the list to return and their return order.
         */
        Query?: string;
        /**
         * The options for modifying the query.
         */
        QueryOptions?: string;
        /**
         * RowLimit
         */
        RowLimit?: string;
        /**
         * The names of the fields to include in the query result.
         */
        ViewFields?: string;
        /**
         * The GUID of the view.
         */
        ViewName?: string;
    }
    /**
     * Determines the display mode of the given control or view
     */
    export enum ControlMode {
        Display = 1,
        Edit = 2,
        New = 3,
    }
    /**
     * Represents properties of a list item field and its value.
     */
    export interface ListItemFormUpdateValue {
        /**
         * The error message result after validating the value for the field.
         */
        ErrorMessage?: string;
        /**
         * The internal name of the field.
         */
        FieldName?: string;
        /**
         * The value of the field, in string format.
         */
        FieldValue?: string;
        /**
         * Indicates whether there was an error result after validating the value for the field.
         */
        HasException?: boolean;
    }
    /**
     * Specifies the type of the field.
     */
    export enum FieldTypes {
        Invalid = 0,
        Integer = 1,
        Text = 2,
        Note = 3,
        DateTime = 4,
        Counter = 5,
        Choice = 6,
        Lookup = 7,
        Boolean = 8,
        Number = 9,
        Currency = 10,
        URL = 11,
        Computed = 12,
        Threading = 13,
        Guid = 14,
        MultiChoice = 15,
        GridChoice = 16,
        Calculated = 17,
        File = 18,
        Attachments = 19,
        User = 20,
        Recurrence = 21,
        CrossProjectLink = 22,
        ModStat = 23,
        Error = 24,
        ContentTypeId = 25,
        PageSeparator = 26,
        ThreadIndex = 27,
        WorkflowStatus = 28,
        AllDayEvent = 29,
        WorkflowEventType = 30,
    }
    export enum DateTimeFieldFormatType {
        DateOnly = 0,
        DateTime = 1,
    }
    /**
     * Specifies the control settings while adding a field.
     */
    export enum AddFieldOptions {
        /**
         *  Specify that a new field added to the list must also be added to the default content type in the site collection
         */
        DefaultValue = 0,
        /**
         * Specify that a new field added to the list must also be added to the default content type in the site collection.
         */
        AddToDefaultContentType = 1,
        /**
         * Specify that a new field must not be added to any other content type
         */
        AddToNoContentType = 2,
        /**
         *  Specify that a new field that is added to the specified list must also be added to all content types in the site collection
         */
        AddToAllContentTypes = 4,
        /**
         * Specify adding an internal field name hint for the purpose of avoiding possible database locking or field renaming operations
         */
        AddFieldInternalNameHint = 8,
        /**
         * Specify that a new field that is added to the specified list must also be added to the default list view
         */
        AddFieldToDefaultView = 16,
        /**
         * Specify to confirm that no other field has the same display name
         */
        AddFieldCheckDisplayName = 32,
    }
    export interface XmlSchemaFieldCreationInformation {
        Options?: AddFieldOptions;
        SchemaXml: string;
    }
    export enum CalendarType {
        Gregorian = 1,
        Japan = 3,
        Taiwan = 4,
        Korea = 5,
        Hijri = 6,
        Thai = 7,
        Hebrew = 8,
        GregorianMEFrench = 9,
        GregorianArabic = 10,
        GregorianXLITEnglish = 11,
        GregorianXLITFrench = 12,
        KoreaJapanLunar = 14,
        ChineseLunar = 15,
        SakaEra = 16,
        UmAlQura = 23,
    }
    export enum UrlFieldFormatType {
        Hyperlink = 0,
        Image = 1,
    }
    export interface BasePermissions {
        Low: string;
        High: string;
    }
    export interface FollowedContent {
        FollowedDocumentsUrl: string;
        FollowedSitesUrl: string;
    }
    export interface UserProfile {
        /**
         * An object containing the user's FollowedDocumentsUrl and FollowedSitesUrl.
         */
        FollowedContent?: FollowedContent;
        /**
         * The account name of the user. (SharePoint Online only)
         */
        AccountName?: string;
        /**
         * The display name of the user. (SharePoint Online only)
         */
        DisplayName?: string;
        /**
         * The FirstRun flag of the user. (SharePoint Online only)
         */
        O15FirstRunExperience?: number;
        /**
         * The personal site of the user.
         */
        PersonalSite?: string;
        /**
         * The capabilities of the user's personal site. Represents a bitwise PersonalSiteCapabilities value:
         * None = 0; Profile Value = 1; Social Value = 2; Storage Value = 4; MyTasksDashboard Value = 8; Education Value = 16; Guest Value = 32.
         */
        PersonalSiteCapabilities?: number;
        /**
         * The error thrown when the user's personal site was first created, if any. (SharePoint Online only)
         */
        PersonalSiteFirstCreationError?: string;
        /**
         * The date and time when the user's personal site was first created. (SharePoint Online only)
         */
        PersonalSiteFirstCreationTime?: Date;
        /**
         * The status for the state of the personal site instantiation
         */
        PersonalSiteInstantiationState?: number;
        /**
         * The date and time when the user's personal site was last created. (SharePoint Online only)
         */
        PersonalSiteLastCreationTime?: Date;
        /**
         * The number of attempts made to create the user's personal site. (SharePoint Online only)
         */
        PersonalSiteNumberOfRetries?: number;
        /**
         * Indicates whether the user's picture is imported from Exchange.
         */
        PictureImportEnabled?: boolean;
        /**
         * The public URL of the personal site of the current user. (SharePoint Online only)
         */
        PublicUrl?: string;
        /**
         * The URL used to create the user's personal site.
         */
        UrlToCreatePersonalSite?: string;
    }
    export interface HashTag {
        /**
         * The hash tag's internal name.
         */
        Name?: string;
        /**
         * The number of times that the hash tag is used.
         */
        UseCount?: number;
    }
    export interface HashTagCollection {
        Items: HashTag[];
    }
    export interface UserIdInfo {
        NameId?: string;
        NameIdIssuer?: string;
    }
    export enum PrincipalType {
        None = 0,
        User = 1,
        DistributionList = 2,
        SecurityGroup = 4,
        SharePointGroup = 8,
        All = 15,
    }
    export interface DocumentLibraryInformation {
        AbsoluteUrl?: string;
        Modified?: Date;
        ModifiedFriendlyDisplay?: string;
        ServerRelativeUrl?: string;
        Title?: string;
    }
    export interface ContextInfo {
        FormDigestTimeoutSeconds?: number;
        FormDigestValue?: number;
        LibraryVersion?: string;
        SiteFullUrl?: string;
        SupportedSchemaVersions?: string[];
        WebFullUrl?: string;
    }
    export interface RenderListData {
        Row: any[];
        FirstRow: number;
        FolderPermissions: string;
        LastRow: number;
        FilterLink: string;
        ForceNoHierarchy: string;
        HierarchyHasIndention: string;
    }
    export enum PageType {
        Invalid = -1,
        DefaultView = 0,
        NormalView = 1,
        DialogView = 2,
        View = 3,
        DisplayForm = 4,
        DisplayFormDialog = 5,
        EditForm = 6,
        EditFormDialog = 7,
        NewForm = 8,
        NewFormDialog = 9,
        SolutionForm = 10,
        PAGE_MAXITEMS = 11,
    }
    export interface ListFormData {
        ContentType?: string;
        Title?: string;
        Author?: string;
        Editor?: string;
        Created?: Date;
        Modified: Date;
        Attachments?: any;
        ListSchema?: any;
        FormControlMode?: number;
        FieldControlModes?: {
            Title?: number;
            Author?: number;
            Editor?: number;
            Created?: number;
            Modified?: number;
            Attachments?: number;
        };
        WebAttributes?: {
            WebUrl?: string;
            EffectivePresenceEnabled?: boolean;
            AllowScriptableWebParts?: boolean;
            PermissionCustomizePages?: boolean;
            LCID?: number;
            CurrentUserId?: number;
        };
        ItemAttributes?: {
            Id?: number;
            FsObjType?: number;
            ExternalListItem?: boolean;
            Url?: string;
            EffectiveBasePermissionsLow?: number;
            EffectiveBasePermissionsHigh?: number;
        };
        ListAttributes?: {
            Id?: string;
            BaseType?: number;
            Direction?: string;
            ListTemplateType?: number;
            DefaultItemOpen?: number;
            EnableVersioning?: boolean;
        };
        CSRCustomLayout?: boolean;
        PostBackRequired?: boolean;
        PreviousPostBackHandled?: boolean;
        UploadMode?: boolean;
        SubmitButtonID?: string;
        ItemContentTypeName?: string;
        ItemContentTypeId?: string;
        JSLinks?: string;
    }
}
declare module "sharepoint/rest/siteusers" {
    import { Queryable, QueryableInstance, QueryableCollection } from "sharepoint/rest/queryable";
    import { SiteGroups } from "sharepoint/rest/sitegroups";
    import { UserIdInfo, PrincipalType } from "sharepoint/rest/types";
    /**
     * Properties that provide a getter, but no setter.
     *
     */
    export interface UserReadOnlyProperties {
        id?: number;
        isHiddenInUI?: boolean;
        loginName?: string;
        principalType?: PrincipalType;
        userIdInfo?: UserIdInfo;
    }
    /**
     * Properties that provide both a getter, and a setter.
     *
     */
    export interface UserWriteableProperties {
        isSiteAdmin?: string;
        email?: string;
        title?: string;
    }
    /**
     * Properties that provide both a getter, and a setter.
     *
     */
    export interface UserUpdateResult {
        user: SiteUser;
        data: any;
    }
    export interface UserProps extends UserReadOnlyProperties, UserWriteableProperties {
        __metadata: {
            id?: string;
            url?: string;
            type?: string;
        };
    }
    /**
     * Describes a collection of all site collection users
     *
     */
    export class SiteUsers extends QueryableCollection {
        /**
         * Creates a new instance of the Users class
         *
         * @param baseUrl The url or Queryable which forms the parent of this user collection
         */
        constructor(baseUrl: string | Queryable, path?: string);
        /**
         * Gets a user from the collection by email
         *
         * @param email The email of the user
         */
        getByEmail(email: string): SiteUser;
        /**
         * Gets a user from the collection by id
         *
         * @param id The id of the user
         */
        getById(id: number): SiteUser;
        /**
         * Gets a user from the collection by login name
         *
         * @param loginName The email address of the user
         */
        getByLoginName(loginName: string): SiteUser;
        /**
         * Removes a user from the collection by id
         *
         * @param id The id of the user
         */
        removeById(id: number | Queryable): Promise<void>;
        /**
         * Removes a user from the collection by login name
         *
         * @param loginName The login name of the user
         */
        removeByLoginName(loginName: string): Promise<any>;
        /**
         * Add a user to a group
         *
         * @param loginName The login name of the user to add to the group
         *
         */
        add(loginName: string): Promise<SiteUser>;
    }
    /**
     * Describes a single user
     *
     */
    export class SiteUser extends QueryableInstance {
        /**
         * Creates a new instance of the User class
         *
         * @param baseUrl The url or Queryable which forms the parent of this fields collection
         * @param path Optional, passes the path to the user
         */
        constructor(baseUrl: string | Queryable, path?: string);
        /**
         * Get's the groups for this user.
         *
         */
        readonly groups: SiteGroups;
        /**
        * Updates this user instance with the supplied properties
        *
        * @param properties A plain object of property names and values to update for the user
        */
        update(properties: UserWriteableProperties): Promise<UserUpdateResult>;
        /**
         * Delete this user
         *
         */
        delete(): Promise<void>;
    }
}
declare module "sharepoint/rest/sitegroups" {
    import { Queryable, QueryableInstance, QueryableCollection } from "sharepoint/rest/queryable";
    import { SiteUser, SiteUsers } from "sharepoint/rest/siteusers";
    /**
     * Properties that provide a getter, but no setter.
     *
     */
    export interface GroupReadOnlyProperties {
        canCurrentUserEditMembership?: boolean;
        canCurrentUserManageGroup?: boolean;
        canCurrentUserViewMembership?: boolean;
        id?: number;
        isHiddenInUI?: boolean;
        loginName?: string;
        ownerTitle?: string;
        principalType?: PrincipalType;
        users?: SiteUsers;
    }
    /**
     * Properties that provide both a getter, and a setter.
     *
     */
    export interface GroupWriteableProperties {
        allowMembersEditMembership?: boolean;
        allowRequestToJoinLeave?: boolean;
        autoAcceptRequestToJoinLeave?: boolean;
        description?: string;
        onlyAllowMembersViewMembership?: boolean;
        owner?: number | SiteUser | SiteGroup;
        requestToJoinLeaveEmailSetting?: string;
        title?: string;
    }
    /**
     * Group Properties
     *
     */
    export interface GroupProperties extends GroupReadOnlyProperties, GroupWriteableProperties {
        __metadata: {
            id?: string;
            url?: string;
            type?: string;
        };
    }
    /**
     * Principal Type enum
     *
     */
    export enum PrincipalType {
        None = 0,
        User = 1,
        DistributionList = 2,
        SecurityGroup = 4,
        SharePointGroup = 8,
        All = 15,
    }
    /**
     * Result from adding a group.
     *
     */
    export interface GroupUpdateResult {
        group: SiteGroup;
        data: any;
    }
    /**
     * Results from updating a group
     *
     */
    export interface GroupAddResult {
        group: SiteGroup;
        data: any;
    }
    /**
     * Describes a collection of site users
     *
     */
    export class SiteGroups extends QueryableCollection {
        /**
         * Creates a new instance of the SiteUsers class
         *
         * @param baseUrl The url or Queryable which forms the parent of this user collection
         */
        constructor(baseUrl: string | Queryable, path?: string);
        /**
         * Adds a new group to the site collection
         *
         * @param props The properties to be updated
         */
        add(properties: GroupWriteableProperties): Promise<GroupAddResult>;
        /**
         * Gets a group from the collection by name
         *
         * @param email The name of the group
         */
        getByName(groupName: string): SiteGroup;
        /**
         * Gets a group from the collection by id
         *
         * @param id The id of the group
         */
        getById(id: number): SiteGroup;
        /**
         * Removes the group with the specified member ID from the collection.
         *
         * @param id The id of the group to remove
         */
        removeById(id: number): Promise<void>;
        /**
         * Removes a user from the collection by login name
         *
         * @param loginName The login name of the user
         */
        removeByLoginName(loginName: string): Promise<any>;
    }
    /**
     * Describes a single group
     *
     */
    export class SiteGroup extends QueryableInstance {
        /**
         * Creates a new instance of the Group class
         *
         * @param baseUrl The url or Queryable which forms the parent of this site group
         * @param path Optional, passes the path to the group
         */
        constructor(baseUrl: string | Queryable, path?: string);
        /**
         * Get's the users for this group
         *
         */
        readonly users: SiteUsers;
        /**
        * Updates this group instance with the supplied properties
        *
        * @param properties A GroupWriteableProperties object of property names and values to update for the user
        */
        update(properties: GroupWriteableProperties): Promise<GroupUpdateResult>;
    }
    export interface SiteGroupAddResult {
        group: SiteGroup;
        data: GroupProperties;
    }
}
declare module "sharepoint/rest/roles" {
    import { Queryable, QueryableInstance, QueryableCollection } from "sharepoint/rest/queryable";
    import { SiteGroups } from "sharepoint/rest/sitegroups";
    import { BasePermissions } from "sharepoint/rest/types";
    import { TypedHash } from "collections/collections";
    /**
     * Describes a set of role assignments for the current scope
     *
     */
    export class RoleAssignments extends QueryableCollection {
        /**
         * Creates a new instance of the RoleAssignments class
         *
         * @param baseUrl The url or Queryable which forms the parent of this fields collection
         */
        constructor(baseUrl: string | Queryable, path?: string);
        /**
         * Adds a new role assignment with the specified principal and role definitions to the collection.
         *
         * @param principalId The ID of the user or group to assign permissions to
         * @param roleDefId The ID of the role definition that defines the permissions to assign
         *
         */
        add(principalId: number, roleDefId: number): Promise<void>;
        /**
         * Removes the role assignment with the specified principal and role definition from the collection
         *
         * @param principalId The ID of the user or group in the role assignment.
         * @param roleDefId The ID of the role definition in the role assignment
         *
         */
        remove(principalId: number, roleDefId: number): Promise<void>;
        /**
         * Gets the role assignment associated with the specified principal ID from the collection.
         *
         * @param id The id of the role assignment
         */
        getById(id: number): RoleAssignment;
    }
    export class RoleAssignment extends QueryableInstance {
        /**
     * Creates a new instance of the RoleAssignment class
     *
     * @param baseUrl The url or Queryable which forms the parent of this fields collection
     */
        constructor(baseUrl: string | Queryable, path?: string);
        readonly groups: SiteGroups;
        /**
         * Get the role definition bindings for this role assignment
         *
         */
        readonly bindings: RoleDefinitionBindings;
        /**
         * Delete this role assignment
         *
         */
        delete(): Promise<void>;
    }
    export class RoleDefinitions extends QueryableCollection {
        /**
         * Creates a new instance of the RoleDefinitions class
         *
         * @param baseUrl The url or Queryable which forms the parent of this fields collection
         * @param path
         *
         */
        constructor(baseUrl: string | Queryable, path?: string);
        /**
         * Gets the role definition with the specified ID from the collection.
         *
         * @param id The ID of the role definition.
         *
         */
        getById(id: number): RoleDefinition;
        /**
         * Gets the role definition with the specified name.
         *
         * @param name The name of the role definition.
         *
         */
        getByName(name: string): RoleDefinition;
        /**
         * Gets the role definition with the specified type.
         *
         * @param name The name of the role definition.
         *
         */
        getByType(roleTypeKind: number): RoleDefinition;
        /**
         * Create a role definition
         *
         * @param name The new role definition's name
         * @param description The new role definition's description
         * @param order The order in which the role definition appears
         * @param basePermissions The permissions mask for this role definition
         *
         */
        add(name: string, description: string, order: number, basePermissions: BasePermissions): Promise<RoleDefinitionAddResult>;
    }
    export class RoleDefinition extends QueryableInstance {
        constructor(baseUrl: string | Queryable, path?: string);
        /**
         * Updates this web intance with the supplied properties
         *
         * @param properties A plain object hash of values to update for the web
         */
        update(properties: TypedHash<string | number | boolean | BasePermissions>): Promise<RoleDefinitionUpdateResult>;
        /**
         * Delete this role definition
         *
         */
        delete(): Promise<void>;
    }
    export interface RoleDefinitionUpdateResult {
        definition: RoleDefinition;
        data: any;
    }
    export interface RoleDefinitionAddResult {
        definition: RoleDefinition;
        data: any;
    }
    export class RoleDefinitionBindings extends QueryableCollection {
        constructor(baseUrl: string | Queryable, path?: string);
    }
}
declare module "sharepoint/rest/queryablesecurable" {
    import { RoleAssignments } from "sharepoint/rest/roles";
    import { Queryable, QueryableInstance } from "sharepoint/rest/queryable";
    export class QueryableSecurable extends QueryableInstance {
        /**
         * Gets the set of role assignments for this item
         *
         */
        readonly roleAssignments: RoleAssignments;
        /**
         * Gets the closest securable up the security hierarchy whose permissions are applied to this list item
         *
         */
        readonly firstUniqueAncestorSecurableObject: QueryableInstance;
        /**
         * Gets the effective permissions for the user supplied
         *
         * @param loginName The claims username for the user (ex: i:0#.f|membership|user@domain.com)
         */
        getUserEffectivePermissions(loginName: string): Queryable;
        /**
         * Breaks the security inheritance at this level optinally copying permissions and clearing subscopes
         *
         * @param copyRoleAssignments If true the permissions are copied from the current parent scope
         * @param clearSubscopes Optional. true to make all child securable objects inherit role assignments from the current object
         */
        breakRoleInheritance(copyRoleAssignments?: boolean, clearSubscopes?: boolean): Promise<any>;
        /**
         * Breaks the security inheritance at this level optinally copying permissions and clearing subscopes
         *
         */
        resetRoleInheritance(): Promise<any>;
    }
}
declare module "sharepoint/rest/files" {
    import { Queryable, QueryableCollection, QueryableInstance } from "sharepoint/rest/queryable";
    import { Item } from "sharepoint/rest/items";
    import { ODataParser } from "sharepoint/rest/odata";
    export interface ChunkedFileUploadProgressData {
        stage: "starting" | "continue" | "finishing";
        blockNumber: number;
        totalBlocks: number;
        chunkSize: number;
        currentPointer: number;
        fileSize: number;
    }
    /**
     * Describes a collection of File objects
     *
     */
    export class Files extends QueryableCollection {
        /**
         * Creates a new instance of the Files class
         *
         * @param baseUrl The url or Queryable which forms the parent of this fields collection
         */
        constructor(baseUrl: string | Queryable, path?: string);
        /**
         * Gets a File by filename
         *
         * @param name The name of the file, including extension.
         */
        getByName(name: string): File;
        /**
         * Uploads a file.
         *
         * @param url The folder-relative url of the file.
         * @param content The file contents blob.
         * @param shouldOverWrite Should a file with the same name in the same location be overwritten? (default: true)
         * @returns The new File and the raw response.
         */
        add(url: string, content: Blob, shouldOverWrite?: boolean): Promise<FileAddResult>;
        /**
         * Uploads a file.
         *
         * @param url The folder-relative url of the file.
         * @param content The Blob file content to add
         * @param progress A callback function which can be used to track the progress of the upload
         * @param shouldOverWrite Should a file with the same name in the same location be overwritten? (default: true)
         * @param chunkSize The size of each file slice, in bytes (default: 10485760)
         * @returns The new File and the raw response.
         */
        addChunked(url: string, content: Blob, progress?: (data: ChunkedFileUploadProgressData) => void, shouldOverWrite?: boolean, chunkSize?: number): Promise<FileAddResult>;
        /**
         * Adds a ghosted file to an existing list or document library.
         *
         * @param fileUrl The server-relative url where you want to save the file.
         * @param templateFileType The type of use to create the file.
         * @returns The template file that was added and the raw response.
         */
        addTemplateFile(fileUrl: string, templateFileType: TemplateFileType): Promise<FileAddResult>;
    }
    /**
     * Describes a single File instance
     *
     */
    export class File extends QueryableInstance {
        /**
         * Creates a new instance of the File class
         *
         * @param baseUrl The url or Queryable which forms the parent of this fields collection
         * @param path Optional, if supplied will be appended to the supplied baseUrl
         */
        constructor(baseUrl: string | Queryable, path?: string);
        /**
         * Gets a value that specifies the list item field values for the list item corresponding to the file.
         *
         */
        readonly listItemAllFields: Item;
        /**
         * Gets a collection of versions
         *
         */
        readonly versions: Versions;
        /**
         * Approves the file submitted for content approval with the specified comment.
         * Only documents in lists that are enabled for content approval can be approved.
         *
         * @param comment The comment for the approval.
         */
        approve(comment: string): Promise<void>;
        /**
         * Stops the chunk upload session without saving the uploaded data.
         * If the file doesn’t already exist in the library, the partially uploaded file will be deleted.
         * Use this in response to user action (as in a request to cancel an upload) or an error or exception.
         * Use the uploadId value that was passed to the StartUpload method that started the upload session.
         * This method is currently available only on Office 365.
         *
         * @param uploadId The unique identifier of the upload session.
         */
        cancelUpload(uploadId: string): Promise<void>;
        /**
         * Checks the file in to a document library based on the check-in type.
         *
         * @param comment A comment for the check-in. Its length must be <= 1023.
         * @param checkinType The check-in type for the file.
         */
        checkin(comment?: string, checkinType?: CheckinType): Promise<void>;
        /**
         * Checks out the file from a document library.
         */
        checkout(): Promise<void>;
        /**
         * Copies the file to the destination url.
         *
         * @param url The absolute url or server relative url of the destination file path to copy to.
         * @param shouldOverWrite Should a file with the same name in the same location be overwritten?
         */
        copyTo(url: string, shouldOverWrite?: boolean): Promise<void>;
        /**
         * Delete this file.
         *
         * @param eTag Value used in the IF-Match header, by default "*"
         */
        delete(eTag?: string): Promise<void>;
        /**
         * Denies approval for a file that was submitted for content approval.
         * Only documents in lists that are enabled for content approval can be denied.
         *
         * @param comment The comment for the denial.
         */
        deny(comment?: string): Promise<void>;
        /**
         * Specifies the control set used to access, modify, or add Web Parts associated with this Web Part Page and view.
         * An exception is thrown if the file is not an ASPX page.
         *
         * @param scope The WebPartsPersonalizationScope view on the Web Parts page.
         */
        getLimitedWebPartManager(scope?: WebPartsPersonalizationScope): Queryable;
        /**
         * Moves the file to the specified destination url.
         *
         * @param url The absolute url or server relative url of the destination file path to move to.
         * @param moveOperations The bitwise MoveOperations value for how to move the file.
         */
        moveTo(url: string, moveOperations?: MoveOperations): Promise<void>;
        /**
         * Submits the file for content approval with the specified comment.
         *
         * @param comment The comment for the published file. Its length must be <= 1023.
         */
        publish(comment?: string): Promise<void>;
        /**
         * Moves the file to the Recycle Bin and returns the identifier of the new Recycle Bin item.
         *
         * @returns The GUID of the recycled file.
         */
        recycle(): Promise<string>;
        /**
         * Reverts an existing checkout for the file.
         *
         */
        undoCheckout(): Promise<void>;
        /**
         * Removes the file from content approval or unpublish a major version.
         *
         * @param comment The comment for the unpublish operation. Its length must be <= 1023.
         */
        unpublish(comment?: string): Promise<void>;
        /**
         * Gets the contents of the file as text
         *
         */
        getText(): Promise<string>;
        /**
         * Gets the contents of the file as a blob, does not work in Node.js
         *
         */
        getBlob(): Promise<Blob>;
        /**
         * Gets the contents of a file as an ArrayBuffer, works in Node.js
         */
        getBuffer(): Promise<ArrayBuffer>;
        /**
         * Sets the content of a file, for large files use setContentChunked
         *
         * @param content The file content
         *
         */
        setContent(content: string | ArrayBuffer | Blob): Promise<File>;
        /**
         * Sets the contents of a file using a chunked upload approach
         *
         * @param file The file to upload
         * @param progress A callback function which can be used to track the progress of the upload
         * @param chunkSize The size of each file slice, in bytes (default: 10485760)
         */
        setContentChunked(file: Blob, progress?: (data: ChunkedFileUploadProgressData) => void, chunkSize?: number): Promise<File>;
        /**
         * Starts a new chunk upload session and uploads the first fragment.
         * The current file content is not changed when this method completes.
         * The method is idempotent (and therefore does not change the result) as long as you use the same values for uploadId and stream.
         * The upload session ends either when you use the CancelUpload method or when you successfully
         * complete the upload session by passing the rest of the file contents through the ContinueUpload and FinishUpload methods.
         * The StartUpload and ContinueUpload methods return the size of the running total of uploaded data in bytes,
         * so you can pass those return values to subsequent uses of ContinueUpload and FinishUpload.
         * This method is currently available only on Office 365.
         *
         * @param uploadId The unique identifier of the upload session.
         * @param fragment The file contents.
         * @returns The size of the total uploaded data in bytes.
         */
        private startUpload(uploadId, fragment);
        /**
         * Continues the chunk upload session with an additional fragment.
         * The current file content is not changed.
         * Use the uploadId value that was passed to the StartUpload method that started the upload session.
         * This method is currently available only on Office 365.
         *
         * @param uploadId The unique identifier of the upload session.
         * @param fileOffset The size of the offset into the file where the fragment starts.
         * @param fragment The file contents.
         * @returns The size of the total uploaded data in bytes.
         */
        private continueUpload(uploadId, fileOffset, fragment);
        /**
         * Uploads the last file fragment and commits the file. The current file content is changed when this method completes.
         * Use the uploadId value that was passed to the StartUpload method that started the upload session.
         * This method is currently available only on Office 365.
         *
         * @param uploadId The unique identifier of the upload session.
         * @param fileOffset The size of the offset into the file where the fragment starts.
         * @param fragment The file contents.
         * @returns The newly uploaded file.
         */
        private finishUpload(uploadId, fileOffset, fragment);
    }
    export class TextFileParser implements ODataParser<any, string> {
        parse(r: Response): Promise<string>;
    }
    export class BlobFileParser implements ODataParser<any, Blob> {
        parse(r: Response): Promise<Blob>;
    }
    export class BufferFileParser implements ODataParser<any, ArrayBuffer> {
        parse(r: any): Promise<ArrayBuffer>;
    }
    /**
     * Describes a collection of Version objects
     *
     */
    export class Versions extends QueryableCollection {
        /**
         * Creates a new instance of the File class
         *
         * @param baseUrl The url or Queryable which forms the parent of this fields collection
         */
        constructor(baseUrl: string | Queryable, path?: string);
        /**
         * Gets a version by id
         *
         * @param versionId The id of the version to retrieve
         */
        getById(versionId: number): Version;
        /**
         * Deletes all the file version objects in the collection.
         *
         */
        deleteAll(): Promise<void>;
        /**
         * Deletes the specified version of the file.
         *
         * @param versionId The ID of the file version to delete.
         */
        deleteById(versionId: number): Promise<void>;
        /**
         * Deletes the file version object with the specified version label.
         *
         * @param label The version label of the file version to delete, for example: 1.2
         */
        deleteByLabel(label: string): Promise<void>;
        /**
         * Creates a new file version from the file specified by the version label.
         *
         * @param label The version label of the file version to restore, for example: 1.2
         */
        restoreByLabel(label: string): Promise<void>;
    }
    /**
     * Describes a single Version instance
     *
     */
    export class Version extends QueryableInstance {
        /**
         * Creates a new instance of the Version class
         *
         * @param baseUrl The url or Queryable which forms the parent of this fields collection
         * @param path Optional, if supplied will be appended to the supplied baseUrl
         */
        constructor(baseUrl: string | Queryable, path?: string);
        /**
        * Delete a specific version of a file.
        *
        * @param eTag Value used in the IF-Match header, by default "*"
        */
        delete(eTag?: string): Promise<void>;
    }
    export enum CheckinType {
        Minor = 0,
        Major = 1,
        Overwrite = 2,
    }
    export interface FileAddResult {
        file: File;
        data: any;
    }
    export enum WebPartsPersonalizationScope {
        User = 0,
        Shared = 1,
    }
    export enum MoveOperations {
        Overwrite = 1,
        AllowBrokenThickets = 8,
    }
    export enum TemplateFileType {
        StandardPage = 0,
        WikiPage = 1,
        FormPage = 2,
    }
}
declare module "sharepoint/rest/folders" {
    import { Queryable, QueryableCollection, QueryableInstance } from "sharepoint/rest/queryable";
    import { Files } from "sharepoint/rest/files";
    import { Item } from "sharepoint/rest/items";
    /**
     * Describes a collection of Folder objects
     *
     */
    export class Folders extends QueryableCollection {
        /**
         * Creates a new instance of the Folders class
         *
         * @param baseUrl The url or Queryable which forms the parent of this fields collection
         */
        constructor(baseUrl: string | Queryable, path?: string);
        /**
         * Gets a folder by folder name
         *
         */
        getByName(name: string): Folder;
        /**
         * Adds a new folder to the current folder (relative) or any folder (absolute)
         *
         * @param url The relative or absolute url where the new folder will be created. Urls starting with a forward slash are absolute.
         * @returns The new Folder and the raw response.
         */
        add(url: string): Promise<FolderAddResult>;
    }
    /**
     * Describes a single Folder instance
     *
     */
    export class Folder extends QueryableInstance {
        /**
         * Creates a new instance of the Folder class
         *
         * @param baseUrl The url or Queryable which forms the parent of this fields collection
         * @param path Optional, if supplied will be appended to the supplied baseUrl
         */
        constructor(baseUrl: string | Queryable, path?: string);
        /**
         * Specifies the sequence in which content types are displayed.
         *
         */
        readonly contentTypeOrder: QueryableCollection;
        /**
         * Gets this folder's files
         *
         */
        readonly files: Files;
        /**
         * Gets this folder's sub folders
         *
         */
        readonly folders: Folders;
        /**
         * Gets this folder's list item
         *
         */
        readonly listItemAllFields: Item;
        /**
         * Gets the parent folder, if available
         *
         */
        readonly parentFolder: Folder;
        /**
         * Gets this folder's properties
         *
         */
        readonly properties: QueryableInstance;
        /**
         * Gets this folder's server relative url
         *
         */
        readonly serverRelativeUrl: Queryable;
        /**
         * Gets a value that specifies the content type order.
         *
         */
        readonly uniqueContentTypeOrder: QueryableCollection;
        /**
        * Delete this folder
        *
        * @param eTag Value used in the IF-Match header, by default "*"
        */
        delete(eTag?: string): Promise<void>;
        /**
         * Moves the folder to the Recycle Bin and returns the identifier of the new Recycle Bin item.
         */
        recycle(): Promise<string>;
    }
    export interface FolderAddResult {
        folder: Folder;
        data: any;
    }
}
declare module "sharepoint/rest/contenttypes" {
    import { TypedHash } from "collections/collections";
    import { Queryable, QueryableCollection, QueryableInstance } from "sharepoint/rest/queryable";
    /**
     * Describes a collection of content types
     *
     */
    export class ContentTypes extends QueryableCollection {
        /**
         * Creates a new instance of the ContentTypes class
         *
         * @param baseUrl The url or Queryable which forms the parent of this content types collection
         */
        constructor(baseUrl: string | Queryable, path?: string);
        /**
         * Gets a ContentType by content type id
         */
        getById(id: string): ContentType;
        /**
         * Adds an existing contenttype to a content type collection
         *
         * @param contentTypeId in the following format, for example: 0x010102
         */
        addAvailableContentType(contentTypeId: string): Promise<ContentTypeAddResult>;
        /**
         * Adds a new content type to the collection
         *
         * @param id The desired content type id for the new content type (also determines the parent content type)
         * @param name The name of the content type
         * @param description The description of the content type
         * @param group The group in which to add the content type
         * @param additionalSettings Any additional settings to provide when creating the content type
         *
         */
        add(id: string, name: string, description?: string, group?: string, additionalSettings?: TypedHash<string | number | boolean>): Promise<ContentTypeAddResult>;
    }
    /**
     * Describes a single ContentType instance
     *
     */
    export class ContentType extends QueryableInstance {
        /**
         * Creates a new instance of the ContentType class
         *
         * @param baseUrl The url or Queryable which forms the parent of this content type instance
         */
        constructor(baseUrl: string | Queryable, path?: string);
        /**
         * Gets the column (also known as field) references in the content type.
        */
        readonly fieldLinks: Queryable;
        /**
         * Gets a value that specifies the collection of fields for the content type.
         */
        readonly fields: Queryable;
        /**
         * Gets the parent content type of the content type.
         */
        readonly parent: ContentType;
        /**
         * Gets a value that specifies the collection of workflow associations for the content type.
         */
        readonly workflowAssociations: Queryable;
    }
    export interface ContentTypeAddResult {
        contentType: ContentType;
        data: any;
    }
}
declare module "sharepoint/rest/items" {
    import { Queryable, QueryableCollection, QueryableInstance } from "sharepoint/rest/queryable";
    import { QueryableSecurable } from "sharepoint/rest/queryablesecurable";
    import { Folder } from "sharepoint/rest/folders";
    import { ContentType } from "sharepoint/rest/contenttypes";
    import { TypedHash } from "collections/collections";
    import * as Types from "sharepoint/rest/types";
    /**
     * Describes a collection of Item objects
     *
     */
    export class Items extends QueryableCollection {
        /**
         * Creates a new instance of the Items class
         *
         * @param baseUrl The url or Queryable which forms the parent of this fields collection
         */
        constructor(baseUrl: string | Queryable, path?: string);
        /**
         * Gets an Item by id
         *
         * @param id The integer id of the item to retrieve
         */
        getById(id: number): Item;
        /**
         * Skips the specified number of items (https://msdn.microsoft.com/en-us/library/office/fp142385.aspx#sectionSection6)
         *
         * @param skip The starting id where the page should start, use with top to specify pages
         */
        skip(skip: number): this;
        /**
         * Gets a collection designed to aid in paging through data
         *
         */
        getPaged(): Promise<PagedItemCollection<any>>;
        /**
         * Adds a new item to the collection
         *
         * @param properties The new items's properties
         */
        add(properties?: TypedHash<string | number | boolean>): Promise<ItemAddResult>;
    }
    /**
     * Descrines a single Item instance
     *
     */
    export class Item extends QueryableSecurable {
        /**
         * Creates a new instance of the Items class
         *
         * @param baseUrl The url or Queryable which forms the parent of this fields collection
         */
        constructor(baseUrl: string | Queryable, path?: string);
        /**
         * Gets the set of attachments for this item
         *
         */
        readonly attachmentFiles: QueryableCollection;
        /**
         * Gets the content type for this item
         *
         */
        readonly contentType: ContentType;
        /**
         * Gets the effective base permissions for the item
         *
         */
        readonly effectiveBasePermissions: Queryable;
        /**
         * Gets the effective base permissions for the item in a UI context
         *
         */
        readonly effectiveBasePermissionsForUI: Queryable;
        /**
         * Gets the field values for this list item in their HTML representation
         *
         */
        readonly fieldValuesAsHTML: QueryableInstance;
        /**
         * Gets the field values for this list item in their text representation
         *
         */
        readonly fieldValuesAsText: QueryableInstance;
        /**
         * Gets the field values for this list item for use in editing controls
         *
         */
        readonly fieldValuesForEdit: QueryableInstance;
        /**
         * Gets the folder associated with this list item (if this item represents a folder)
         *
         */
        readonly folder: Folder;
        /**
         * Updates this list intance with the supplied properties
         *
         * @param properties A plain object hash of values to update for the list
         * @param eTag Value used in the IF-Match header, by default "*"
         */
        update(properties: TypedHash<string | number | boolean>, eTag?: string): Promise<ItemUpdateResult>;
        /**
         * Delete this item
         *
         * @param eTag Value used in the IF-Match header, by default "*"
         */
        delete(eTag?: string): Promise<void>;
        /**
         * Moves the list item to the Recycle Bin and returns the identifier of the new Recycle Bin item.
         */
        recycle(): Promise<string>;
        /**
         * Gets a string representation of the full URL to the WOPI frame.
         * If there is no associated WOPI application, or no associated action, an empty string is returned.
         *
         * @param action Display mode: 0: view, 1: edit, 2: mobileView, 3: interactivePreview
         */
        getWopiFrameUrl(action?: number): Promise<string>;
        /**
         * Validates and sets the values of the specified collection of fields for the list item.
         *
         * @param formValues The fields to change and their new values.
         * @param newDocumentUpdate true if the list item is a document being updated after upload; otherwise false.
         */
        validateUpdateListItem(formValues: Types.ListItemFormUpdateValue[], newDocumentUpdate?: boolean): Promise<Types.ListItemFormUpdateValue[]>;
    }
    export interface ItemAddResult {
        item: Item;
        data: any;
    }
    export interface ItemUpdateResult {
        item: Item;
        data: any;
    }
    /**
     * Provides paging functionality for list items
     */
    export class PagedItemCollection<T> {
        private nextUrl;
        results: T;
        constructor(nextUrl: string, results: T);
        /**
         * If true there are more results available in the set, otherwise there are not
         */
        readonly hasNext: boolean;
        /**
         * Gets the next set of results, or resolves to null if no results are available
         */
        getNext(): Promise<PagedItemCollection<any>>;
    }
}
declare module "sharepoint/rest/views" {
    import { Queryable, QueryableCollection, QueryableInstance } from "sharepoint/rest/queryable";
    import { TypedHash } from "collections/collections";
    /**
     * Describes the views available in the current context
     *
     */
    export class Views extends QueryableCollection {
        /**
         * Creates a new instance of the Views class
         *
         * @param baseUrl The url or Queryable which forms the parent of this fields collection
         */
        constructor(baseUrl: string | Queryable);
        /**
         * Gets a view by guid id
         *
         * @param id The GUID id of the view
         */
        getById(id: string): View;
        /**
         * Gets a view by title (case-sensitive)
         *
         * @param title The case-sensitive title of the view
         */
        getByTitle(title: string): View;
        /**
         * Adds a new view to the collection
         *
         * @param title The new views's title
         * @param personalView True if this is a personal view, otherwise false, default = false
         * @param additionalSettings Will be passed as part of the view creation body
         */
        add(title: string, personalView?: boolean, additionalSettings?: TypedHash<string | number | boolean>): Promise<ViewAddResult>;
    }
    /**
     * Describes a single View instance
     *
     */
    export class View extends QueryableInstance {
        /**
         * Creates a new instance of the View class
         *
         * @param baseUrl The url or Queryable which forms the parent of this fields collection
         */
        constructor(baseUrl: string | Queryable, path?: string);
        readonly fields: ViewFields;
        /**
         * Updates this view intance with the supplied properties
         *
         * @param properties A plain object hash of values to update for the view
         */
        update(properties: TypedHash<string | number | boolean>): Promise<ViewUpdateResult>;
        /**
         * Delete this view
         *
         */
        delete(): Promise<void>;
        /**
         * Returns the list view as HTML.
         *
         */
        renderAsHtml(): Promise<string>;
    }
    export class ViewFields extends QueryableCollection {
        constructor(baseUrl: string | Queryable, path?: string);
        /**
         * Gets a value that specifies the XML schema that represents the collection.
         */
        getSchemaXml(): Promise<string>;
        /**
         * Adds the field with the specified field internal name or display name to the collection.
         *
         * @param fieldTitleOrInternalName The case-sensitive internal name or display name of the field to add.
         */
        add(fieldTitleOrInternalName: string): Promise<void>;
        /**
         * Moves the field with the specified field internal name to the specified position in the collection.
         *
         * @param fieldInternalName The case-sensitive internal name of the field to move.
         * @param index The zero-based index of the new position for the field.
         */
        move(fieldInternalName: string, index: number): Promise<void>;
        /**
         * Removes all the fields from the collection.
         */
        removeAll(): Promise<void>;
        /**
         * Removes the field with the specified field internal name from the collection.
         *
         * @param fieldInternalName The case-sensitive internal name of the field to remove from the view.
         */
        remove(fieldInternalName: string): Promise<void>;
    }
    export interface ViewAddResult {
        view: View;
        data: any;
    }
    export interface ViewUpdateResult {
        view: View;
        data: any;
    }
}
declare module "sharepoint/rest/fields" {
    import { Queryable, QueryableCollection, QueryableInstance } from "sharepoint/rest/queryable";
    import { TypedHash } from "collections/collections";
    import * as Types from "sharepoint/rest/types";
    /**
     * Describes a collection of Field objects
     *
     */
    export class Fields extends QueryableCollection {
        /**
         * Creates a new instance of the Fields class
         *
         * @param baseUrl The url or Queryable which forms the parent of this fields collection
         */
        constructor(baseUrl: string | Queryable, path?: string);
        /**
         * Gets a field from the collection by title
         *
         * @param title The case-sensitive title of the field
         */
        getByTitle(title: string): Field;
        /**
         * Gets a field from the collection by using internal name or title
         *
         * @param name The case-sensitive internal name or title of the field
         */
        getByInternalNameOrTitle(name: string): Field;
        /**
         * Gets a list from the collection by guid id
         *
         * @param title The Id of the list
         */
        getById(id: string): Field;
        /**
         * Creates a field based on the specified schema
         */
        createFieldAsXml(xml: string | Types.XmlSchemaFieldCreationInformation): Promise<FieldAddResult>;
        /**
         * Adds a new list to the collection
         *
         * @param title The new field's title
         * @param fieldType The new field's type (ex: SP.FieldText)
         * @param properties Differ by type of field being created (see: https://msdn.microsoft.com/en-us/library/office/dn600182.aspx)
         */
        add(title: string, fieldType: string, properties?: TypedHash<string | number | boolean>): Promise<FieldAddResult>;
        /**
         * Adds a new SP.FieldText to the collection
         *
         * @param title The field title
         * @param maxLength The maximum number of characters allowed in the value of the field.
         * @param properties Differ by type of field being created (see: https://msdn.microsoft.com/en-us/library/office/dn600182.aspx)
         */
        addText(title: string, maxLength?: number, properties?: TypedHash<string | number | boolean>): Promise<FieldAddResult>;
        /**
         * Adds a new SP.FieldCalculated to the collection
         *
         * @param title The field title.
         * @param formula The formula for the field.
         * @param dateFormat The date and time format that is displayed in the field.
         * @param outputType Specifies the output format for the field. Represents a FieldType value.
         * @param properties Differ by type of field being created (see: https://msdn.microsoft.com/en-us/library/office/dn600182.aspx)
         */
        addCalculated(title: string, formula: string, dateFormat: Types.DateTimeFieldFormatType, outputType?: Types.FieldTypes, properties?: TypedHash<string | number | boolean>): Promise<FieldAddResult>;
        /**
         * Adds a new SP.FieldDateTime to the collection
         *
         * @param title The field title
         * @param displayFormat The format of the date and time that is displayed in the field.
         * @param calendarType Specifies the calendar type of the field.
         * @param properties Differ by type of field being created (see: https://msdn.microsoft.com/en-us/library/office/dn600182.aspx)
         */
        addDateTime(title: string, displayFormat?: Types.DateTimeFieldFormatType, calendarType?: Types.CalendarType, friendlyDisplayFormat?: number, properties?: TypedHash<string | number | boolean>): Promise<FieldAddResult>;
        /**
         * Adds a new SP.FieldNumber to the collection
         *
         * @param title The field title
         * @param minValue The field's minimum value
         * @param maxValue The field's maximum value
         * @param properties Differ by type of field being created (see: https://msdn.microsoft.com/en-us/library/office/dn600182.aspx)
         */
        addNumber(title: string, minValue?: number, maxValue?: number, properties?: TypedHash<string | number | boolean>): Promise<FieldAddResult>;
        /**
         * Adds a new SP.FieldCurrency to the collection
         *
         * @param title The field title
         * @param minValue The field's minimum value
         * @param maxValue The field's maximum value
         * @param currencyLocalId Specifies the language code identifier (LCID) used to format the value of the field
         * @param properties Differ by type of field being created (see: https://msdn.microsoft.com/en-us/library/office/dn600182.aspx)
         */
        addCurrency(title: string, minValue?: number, maxValue?: number, currencyLocalId?: number, properties?: TypedHash<string | number | boolean>): Promise<FieldAddResult>;
        /**
         * Adds a new SP.FieldMultiLineText to the collection
         *
         * @param title The field title
         * @param numberOfLines Specifies the number of lines of text to display for the field.
         * @param richText Specifies whether the field supports rich formatting.
         * @param restrictedMode Specifies whether the field supports a subset of rich formatting.
         * @param appendOnly Specifies whether all changes to the value of the field are displayed in list forms.
         * @param allowHyperlink Specifies whether a hyperlink is allowed as a value of the field.
         * @param properties Differ by type of field being created (see: https://msdn.microsoft.com/en-us/library/office/dn600182.aspx)
         *
         */
        addMultilineText(title: string, numberOfLines?: number, richText?: boolean, restrictedMode?: boolean, appendOnly?: boolean, allowHyperlink?: boolean, properties?: TypedHash<string | number | boolean>): Promise<FieldAddResult>;
        /**
         * Adds a new SP.FieldUrl to the collection
         *
         * @param title The field title
         */
        addUrl(title: string, displayFormat?: Types.UrlFieldFormatType, properties?: TypedHash<string | number | boolean>): Promise<FieldAddResult>;
    }
    /**
     * Describes a single of Field instance
     *
     */
    export class Field extends QueryableInstance {
        /**
         * Creates a new instance of the Field class
         *
         * @param baseUrl The url or Queryable which forms the parent of this field instance
         */
        constructor(baseUrl: string | Queryable, path?: string);
        /**
         * Updates this field intance with the supplied properties
         *
         * @param properties A plain object hash of values to update for the list
         * @param fieldType The type value, required to update child field type properties
         */
        update(properties: TypedHash<string | number | boolean>, fieldType?: string): Promise<FieldUpdateResult>;
        /**
         * Delete this fields
         *
         */
        delete(): Promise<void>;
        /**
         * Sets the value of the ShowInDisplayForm property for this field.
         */
        setShowInDisplayForm(show: boolean): Promise<void>;
        /**
         * Sets the value of the ShowInEditForm property for this field.
         */
        setShowInEditForm(show: boolean): Promise<void>;
        /**
         * Sets the value of the ShowInNewForm property for this field.
         */
        setShowInNewForm(show: boolean): Promise<void>;
    }
    export interface FieldAddResult {
        data: any;
        field: Field;
    }
    export interface FieldUpdateResult {
        data: any;
        field: Field;
    }
}
declare module "sharepoint/rest/forms" {
    import { Queryable, QueryableCollection, QueryableInstance } from "sharepoint/rest/queryable";
    /**
     * Describes a collection of Field objects
     *
     */
    export class Forms extends QueryableCollection {
        /**
         * Creates a new instance of the Fields class
         *
         * @param baseUrl The url or Queryable which forms the parent of this fields collection
         */
        constructor(baseUrl: string | Queryable, path?: string);
        /**
         * Gets a form by id
         *
         * @param id The guid id of the item to retrieve
         */
        getById(id: string): Form;
    }
    /**
     * Describes a single of Form instance
     *
     */
    export class Form extends QueryableInstance {
        /**
         * Creates a new instance of the Form class
         *
         * @param baseUrl The url or Queryable which is the parent of this form instance
         */
        constructor(baseUrl: string | Queryable, path?: string);
    }
}
declare module "sharepoint/rest/subscriptions" {
    import { Queryable, QueryableCollection, QueryableInstance } from "sharepoint/rest/queryable";
    /**
     * Describes a collection of webhook subscriptions
     *
     */
    export class Subscriptions extends QueryableCollection {
        /**
         * Creates a new instance of the Subscriptions class
         *
         * @param baseUrl - The url or Queryable which forms the parent of this webhook subscriptions collection
         */
        constructor(baseUrl: string | Queryable, path?: string);
        /**
         * Returns all the webhook subscriptions or the specified webhook subscription
         *
         */
        getById(subscriptionId: string): Subscription;
        /**
         * Create a new webhook subscription
         *
         */
        add(notificationUrl: string, expirationDate: string, clientState?: string): Promise<SubscriptionAddResult>;
    }
    /**
     * Describes a single webhook subscription instance
     *
     */
    export class Subscription extends QueryableInstance {
        /**
         * Creates a new instance of the Subscription class
         *
         * @param baseUrl - The url or Queryable which forms the parent of this webhook subscription instance
         */
        constructor(baseUrl: string | Queryable, path?: string);
        /**
         * Update a webhook subscription
         *
         */
        update(expirationDate: string): Promise<SubscriptionUpdateResult>;
        /**
         * Remove a webhook subscription
         *
         */
        delete(): Promise<void>;
    }
    export interface SubscriptionAddResult {
        subscription: Subscription;
        data: any;
    }
    export interface SubscriptionUpdateResult {
        subscription: Subscription;
        data: any;
    }
}
declare module "sharepoint/rest/usercustomactions" {
    import { Queryable, QueryableInstance, QueryableCollection } from "sharepoint/rest/queryable";
    import { TypedHash } from "collections/collections";
    export class UserCustomActions extends QueryableCollection {
        constructor(baseUrl: string | Queryable, path?: string);
        /**
         * Returns the custom action with the specified identifier.
         *
         * @param id The GUID ID of the user custom action to get.
         */
        getById(id: string): UserCustomAction;
        /**
         * Create a custom action
         *
         * @param creationInfo The information which defines the new custom action
         *
         */
        add(properties: TypedHash<string | boolean | number>): Promise<UserCustomActionAddResult>;
        /**
         * Deletes all custom actions in the collection.
         *
         */
        clear(): Promise<void>;
    }
    export class UserCustomAction extends QueryableInstance {
        constructor(baseUrl: string | Queryable, path?: string);
        update(properties: TypedHash<string | boolean | number>): Promise<UserCustomActionUpdateResult>;
    }
    export interface UserCustomActionAddResult {
        data: any;
        action: UserCustomAction;
    }
    export interface UserCustomActionUpdateResult {
        data: any;
        action: UserCustomAction;
    }
}
declare module "sharepoint/rest/lists" {
    import { Items } from "sharepoint/rest/items";
    import { Views, View } from "sharepoint/rest/views";
    import { ContentTypes } from "sharepoint/rest/contenttypes";
    import { Fields } from "sharepoint/rest/fields";
    import { Forms } from "sharepoint/rest/forms";
    import { Subscriptions } from "sharepoint/rest/subscriptions";
    import { Queryable, QueryableInstance, QueryableCollection } from "sharepoint/rest/queryable";
    import { QueryableSecurable } from "sharepoint/rest/queryablesecurable";
    import { TypedHash } from "collections/collections";
    import { ControlMode, RenderListData, ChangeQuery, CamlQuery, ChangeLogitemQuery, ListFormData } from "sharepoint/rest/types";
    import { UserCustomActions } from "sharepoint/rest/usercustomactions";
    /**
     * Describes a collection of List objects
     *
     */
    export class Lists extends QueryableCollection {
        /**
         * Creates a new instance of the Lists class
         *
         * @param baseUrl The url or Queryable which forms the parent of this fields collection
         */
        constructor(baseUrl: string | Queryable, path?: string);
        /**
         * Gets a list from the collection by title
         *
         * @param title The title of the list
         */
        getByTitle(title: string): List;
        /**
         * Gets a list from the collection by guid id
         *
         * @param title The Id of the list
         */
        getById(id: string): List;
        /**
         * Adds a new list to the collection
         *
         * @param title The new list's title
         * @param description The new list's description
         * @param template The list template value
         * @param enableContentTypes If true content types will be allowed and enabled, otherwise they will be disallowed and not enabled
         * @param additionalSettings Will be passed as part of the list creation body
         */
        add(title: string, description?: string, template?: number, enableContentTypes?: boolean, additionalSettings?: TypedHash<string | number | boolean>): Promise<ListAddResult>;
        /**
         * Ensures that the specified list exists in the collection (note: settings are not updated if the list exists,
         * not supported for batching)
         *
         * @param title The new list's title
         * @param description The new list's description
         * @param template The list template value
         * @param enableContentTypes If true content types will be allowed and enabled, otherwise they will be disallowed and not enabled
         * @param additionalSettings Will be passed as part of the list creation body
         */
        ensure(title: string, description?: string, template?: number, enableContentTypes?: boolean, additionalSettings?: TypedHash<string | number | boolean>): Promise<ListEnsureResult>;
        /**
         * Gets a list that is the default asset location for images or other files, which the users upload to their wiki pages.
         */
        ensureSiteAssetsLibrary(): Promise<List>;
        /**
         * Gets a list that is the default location for wiki pages.
         */
        ensureSitePagesLibrary(): Promise<List>;
    }
    /**
     * Describes a single List instance
     *
     */
    export class List extends QueryableSecurable {
        /**
         * Creates a new instance of the Lists class
         *
         * @param baseUrl The url or Queryable which forms the parent of this fields collection
         * @param path Optional, if supplied will be appended to the supplied baseUrl
         */
        constructor(baseUrl: string | Queryable, path?: string);
        /**
         * Gets the content types in this list
         *
         */
        readonly contentTypes: ContentTypes;
        /**
         * Gets the items in this list
         *
         */
        readonly items: Items;
        /**
         * Gets the views in this list
         *
         */
        readonly views: Views;
        /**
         * Gets the fields in this list
         *
         */
        readonly fields: Fields;
        /**
         * Gets the forms in this list
         *
         */
        readonly forms: Forms;
        /**
         * Gets the default view of this list
         *
         */
        readonly defaultView: QueryableInstance;
        /**
         * Get all custom actions on a site collection
         *
         */
        readonly userCustomActions: UserCustomActions;
        /**
         * Gets the effective base permissions of this list
         *
         */
        readonly effectiveBasePermissions: Queryable;
        /**
         * Gets the event receivers attached to this list
         *
         */
        readonly eventReceivers: QueryableCollection;
        /**
         * Gets the related fields of this list
         *
         */
        readonly relatedFields: Queryable;
        /**
         * Gets the IRM settings for this list
         *
         */
        readonly informationRightsManagementSettings: Queryable;
        /**
         * Gets the webhook subscriptions of this list
         *
         */
        readonly subscriptions: Subscriptions;
        /**
         * Gets a view by view guid id
         *
         */
        getView(viewId: string): View;
        /**
         * Updates this list intance with the supplied properties
         *
         * @param properties A plain object hash of values to update for the list
         * @param eTag Value used in the IF-Match header, by default "*"
         */
        update(properties: TypedHash<string | number | boolean>, eTag?: string): Promise<ListUpdateResult>;
        /**
         * Delete this list
         *
         * @param eTag Value used in the IF-Match header, by default "*"
         */
        delete(eTag?: string): Promise<void>;
        /**
         * Returns the collection of changes from the change log that have occurred within the list, based on the specified query.
         */
        getChanges(query: ChangeQuery): Promise<any>;
        /**
         * Returns a collection of items from the list based on the specified query.
         *
         * @param CamlQuery The Query schema of Collaborative Application Markup
         * Language (CAML) is used in various ways within the context of Microsoft SharePoint Foundation
         * to define queries against list data.
         * see:
         *
         * https://msdn.microsoft.com/en-us/library/office/ms467521.aspx
         *
         * @param expands A URI with a $expand System Query Option indicates that Entries associated with
         * the Entry or Collection of Entries identified by the Resource Path
         * section of the URI must be represented inline (i.e. eagerly loaded).
         * see:
         *
         * https://msdn.microsoft.com/en-us/library/office/fp142385.aspx
         *
         * http://www.odata.org/documentation/odata-version-2-0/uri-conventions/#ExpandSystemQueryOption
         */
        getItemsByCAMLQuery(query: CamlQuery, ...expands: string[]): Promise<any>;
        /**
         * See: https://msdn.microsoft.com/en-us/library/office/dn292554.aspx
         */
        getListItemChangesSinceToken(query: ChangeLogitemQuery): Promise<string>;
        /**
         * Moves the list to the Recycle Bin and returns the identifier of the new Recycle Bin item.
         */
        recycle(): Promise<string>;
        /**
         * Renders list data based on the view xml provided
         */
        renderListData(viewXml: string): Promise<RenderListData>;
        /**
         * Gets the field values and field schema attributes for a list item.
         */
        renderListFormData(itemId: number, formId: string, mode: ControlMode): Promise<ListFormData>;
        /**
         * Reserves a list item ID for idempotent list item creation.
         */
        reserveListItemId(): Promise<number>;
    }
    export interface ListAddResult {
        list: List;
        data: any;
    }
    export interface ListUpdateResult {
        list: List;
        data: any;
    }
    export interface ListEnsureResult {
        list: List;
        created: boolean;
        data: any;
    }
}
declare module "sharepoint/rest/quicklaunch" {
    import { Queryable } from "sharepoint/rest/queryable";
    /**
     * Describes the quick launch navigation
     *
     */
    export class QuickLaunch extends Queryable {
        /**
         * Creates a new instance of the Lists class
         *
         * @param baseUrl The url or Queryable which forms the parent of this fields collection
         */
        constructor(baseUrl: string | Queryable);
    }
}
declare module "sharepoint/rest/topnavigationbar" {
    import { Queryable, QueryableInstance } from "sharepoint/rest/queryable";
    /**
     * Describes the top navigation on the site
     *
     */
    export class TopNavigationBar extends QueryableInstance {
        /**
         * Creates a new instance of the SiteUsers class
         *
         * @param baseUrl The url or Queryable which forms the parent of this fields collection
         */
        constructor(baseUrl: string | Queryable);
    }
}
declare module "sharepoint/rest/navigation" {
    import { Queryable } from "sharepoint/rest/queryable";
    import { QuickLaunch } from "sharepoint/rest/quicklaunch";
    import { TopNavigationBar } from "sharepoint/rest/topnavigationbar";
    /**
     * Exposes the navigation components
     *
     */
    export class Navigation extends Queryable {
        /**
         * Creates a new instance of the Lists class
         *
         * @param baseUrl The url or Queryable which forms the parent of this fields collection
         */
        constructor(baseUrl: string | Queryable);
        /**
         * Gets the quicklaunch navigation for the current context
         *
         */
        readonly quicklaunch: QuickLaunch;
        /**
         * Gets the top bar navigation navigation for the current context
         *
         */
        readonly topNavigationBar: TopNavigationBar;
    }
}
declare module "sharepoint/rest/webs" {
    import { Queryable, QueryableCollection } from "sharepoint/rest/queryable";
    import { QueryableSecurable } from "sharepoint/rest/queryablesecurable";
    import { Lists } from "sharepoint/rest/lists";
    import { Fields } from "sharepoint/rest/fields";
    import { Navigation } from "sharepoint/rest/navigation";
    import { SiteGroups } from "sharepoint/rest/sitegroups";
    import { ContentTypes } from "sharepoint/rest/contenttypes";
    import { Folders, Folder } from "sharepoint/rest/folders";
    import { RoleDefinitions } from "sharepoint/rest/roles";
    import { File } from "sharepoint/rest/files";
    import { TypedHash } from "collections/collections";
    import * as Types from "sharepoint/rest/types";
    import { List } from "sharepoint/rest/lists";
    import { SiteUsers, SiteUser } from "sharepoint/rest/siteusers";
    import { UserCustomActions } from "sharepoint/rest/usercustomactions";
    import { ODataBatch } from "sharepoint/rest/odata";
    export class Webs extends QueryableCollection {
        constructor(baseUrl: string | Queryable, webPath?: string);
        /**
         * Adds a new web to the collection
         *
         * @param title The new web's title
         * @param url The new web's relative url
         * @param description The web web's description
         * @param template The web's template
         * @param language The language code to use for this web
         * @param inheritPermissions If true permissions will be inherited from the partent web
         * @param additionalSettings Will be passed as part of the web creation body
         */
        add(title: string, url: string, description?: string, template?: string, language?: number, inheritPermissions?: boolean, additionalSettings?: TypedHash<string | number | boolean>): Promise<WebAddResult>;
    }
    /**
     * Describes a web
     *
     */
    export class Web extends QueryableSecurable {
        constructor(baseUrl: string | Queryable, path?: string);
        readonly webs: Webs;
        /**
         * Get the content types available in this web
         *
         */
        readonly contentTypes: ContentTypes;
        /**
         * Get the lists in this web
         *
         */
        readonly lists: Lists;
        /**
         * Gets the fields in this web
         *
         */
        readonly fields: Fields;
        /**
         * Gets the available fields in this web
         *
         */
        readonly availablefields: Fields;
        /**
         * Get the navigation options in this web
         *
         */
        readonly navigation: Navigation;
        /**
         * Gets the site users
         *
         */
        readonly siteUsers: SiteUsers;
        /**
         * Gets the site groups
         *
         */
        readonly siteGroups: SiteGroups;
        /**
         * Get the folders in this web
         *
         */
        readonly folders: Folders;
        /**
         * Get all custom actions on a site
         *
         */
        readonly userCustomActions: UserCustomActions;
        /**
         * Gets the collection of RoleDefinition resources.
         *
         */
        readonly roleDefinitions: RoleDefinitions;
        /**
         * Creates a new batch for requests within the context of context this web
         *
         */
        createBatch(): ODataBatch;
        /**
         * Get a folder by server relative url
         *
         * @param folderRelativeUrl the server relative path to the folder (including /sites/ if applicable)
         */
        getFolderByServerRelativeUrl(folderRelativeUrl: string): Folder;
        /**
         * Get a file by server relative url
         *
         * @param fileRelativeUrl the server relative path to the file (including /sites/ if applicable)
         */
        getFileByServerRelativeUrl(fileRelativeUrl: string): File;
        /**
         * Get a list by server relative url (list's root folder)
         *
         * @param listRelativeUrl the server relative path to the list's root folder (including /sites/ if applicable)
         */
        getList(listRelativeUrl: string): List;
        /**
         * Updates this web intance with the supplied properties
         *
         * @param properties A plain object hash of values to update for the web
         */
        update(properties: TypedHash<string | number | boolean>): Promise<WebUpdateResult>;
        /**
         * Delete this web
         *
         */
        delete(): Promise<void>;
        /**
         * Applies the theme specified by the contents of each of the files specified in the arguments to the site.
         *
         * @param colorPaletteUrl Server-relative URL of the color palette file.
         * @param fontSchemeUrl Server-relative URL of the font scheme.
         * @param backgroundImageUrl Server-relative URL of the background image.
         * @param shareGenerated true to store the generated theme files in the root site, or false to store them in this site.
         */
        applyTheme(colorPaletteUrl: string, fontSchemeUrl: string, backgroundImageUrl: string, shareGenerated: boolean): Promise<void>;
        /**
         * Applies the specified site definition or site template to the Web site that has no template applied to it.
         *
         * @param template Name of the site definition or the name of the site template
         */
        applyWebTemplate(template: string): Promise<void>;
        /**
         * Returns whether the current user has the given set of permissions.
         *
         * @param perms The high and low permission range.
         */
        doesUserHavePermissions(perms: Types.BasePermissions): Promise<boolean>;
        /**
         * Checks whether the specified login name belongs to a valid user in the site. If the user doesn't exist, adds the user to the site.
         *
         * @param loginName The login name of the user (ex: i:0#.f|membership|user@domain.onmicrosoft.com)
         */
        ensureUser(loginName: string): Promise<any>;
        /**
         * Returns a collection of site templates available for the site.
         *
         * @param language The LCID of the site templates to get.
         * @param true to include language-neutral site templates; otherwise false
         */
        availableWebTemplates(language?: number, includeCrossLanugage?: boolean): QueryableCollection;
        /**
         * Returns the list gallery on the site.
         *
         * @param type The gallery type - WebTemplateCatalog = 111, WebPartCatalog = 113 ListTemplateCatalog = 114,
         * MasterPageCatalog = 116, SolutionCatalog = 121, ThemeCatalog = 123, DesignCatalog = 124, AppDataCatalog = 125
         */
        getCatalog(type: number): Promise<List>;
        /**
         * Returns the collection of changes from the change log that have occurred within the list, based on the specified query.
         */
        getChanges(query: Types.ChangeQuery): Promise<any>;
        /**
         * Gets the custom list templates for the site.
         *
         */
        readonly customListTemplate: QueryableCollection;
        /**
         * Returns the user corresponding to the specified member identifier for the current site.
         *
         * @param id The ID of the user.
         */
        getUserById(id: number): SiteUser;
        /**
         * Returns the name of the image file for the icon that is used to represent the specified file.
         *
         * @param filename The file name. If this parameter is empty, the server returns an empty string.
         * @param size The size of the icon: 16x16 pixels = 0, 32x32 pixels = 1.
         * @param progId The ProgID of the application that was used to create the file, in the form OLEServerName.ObjectName
         */
        mapToIcon(filename: string, size?: number, progId?: string): Promise<string>;
    }
    export interface WebAddResult {
        data: any;
        web: Web;
    }
    export interface WebUpdateResult {
        data: any;
        web: Web;
    }
    export interface GetCatalogResult {
        data: any;
        list: List;
    }
}
declare module "configuration/providers/spListConfigurationProvider" {
    import { IConfigurationProvider } from "configuration/configuration";
    import { TypedHash } from "collections/collections";
    import { default as CachingConfigurationProvider } from "configuration/providers/cachingConfigurationProvider";
    import { Web } from "sharepoint/rest/webs";
    /**
     * A configuration provider which loads configuration values from a SharePoint list
     *
     */
    export default class SPListConfigurationProvider implements IConfigurationProvider {
        private sourceWeb;
        private sourceListTitle;
        /**
         * Creates a new SharePoint list based configuration provider
         * @constructor
         * @param {string} webUrl Url of the SharePoint site, where the configuration list is located
         * @param {string} listTitle Title of the SharePoint list, which contains the configuration settings (optional, default = "config")
         */
        constructor(sourceWeb: Web, sourceListTitle?: string);
        /**
         * Gets the url of the SharePoint site, where the configuration list is located
         *
         * @return {string} Url address of the site
         */
        readonly web: Web;
        /**
         * Gets the title of the SharePoint list, which contains the configuration settings
         *
         * @return {string} List title
         */
        readonly listTitle: string;
        /**
         * Loads the configuration values from the SharePoint list
         *
         * @return {Promise<TypedHash<string>>} Promise of loaded configuration values
         */
        getConfiguration(): Promise<TypedHash<string>>;
        /**
         * Wraps the current provider in a cache enabled provider
         *
         * @return {CachingConfigurationProvider} Caching providers which wraps the current provider
         */
        asCaching(): CachingConfigurationProvider;
    }
}
declare module "configuration/providers/providers" {
    import { default as cachingConfigurationProvider } from "configuration/providers/cachingConfigurationProvider";
    import { default as spListConfigurationProvider } from "configuration/providers/spListConfigurationProvider";
    export let CachingConfigurationProvider: typeof cachingConfigurationProvider;
    export let SPListConfigurationProvider: typeof spListConfigurationProvider;
}
declare module "configuration/configuration" {
    import * as Collections from "collections/collections";
    import * as providers from "configuration/providers/providers";
    /**
     * Interface for configuration providers
     *
     */
    export interface IConfigurationProvider {
        /**
         * Gets the configuration from the provider
         */
        getConfiguration(): Promise<Collections.TypedHash<string>>;
    }
    /**
     * Class used to manage the current application settings
     *
     */
    export class Settings {
        /**
         * Set of pre-defined providers which are available from this library
         */
        Providers: typeof providers;
        /**
         * The settings currently stored in this instance
         */
        private _settings;
        /**
         * Creates a new instance of the settings class
         *
         * @constructor
         */
        constructor();
        /**
         * Adds a new single setting, or overwrites a previous setting with the same key
         *
         * @param {string} key The key used to store this setting
         * @param {string} value The setting value to store
         */
        add(key: string, value: string): void;
        /**
         * Adds a JSON value to the collection as a string, you must use getJSON to rehydrate the object when read
         *
         * @param {string} key The key used to store this setting
         * @param {any} value The setting value to store
         */
        addJSON(key: string, value: any): void;
        /**
         * Applies the supplied hash to the setting collection overwriting any existing value, or created new values
         *
         * @param {Collections.TypedHash<any>} hash The set of values to add
         */
        apply(hash: Collections.TypedHash<any>): Promise<void>;
        /**
         * Loads configuration settings into the collection from the supplied provider and returns a Promise
         *
         * @param {IConfigurationProvider} provider The provider from which we will load the settings
         */
        load(provider: IConfigurationProvider): Promise<void>;
        /**
         * Gets a value from the configuration
         *
         * @param {string} key The key whose value we want to return. Returns null if the key does not exist
         * @return {string} string value from the configuration
         */
        get(key: string): string;
        /**
         * Gets a JSON value, rehydrating the stored string to the original object
         *
         * @param {string} key The key whose value we want to return. Returns null if the key does not exist
         * @return {any} object from the configuration
         */
        getJSON(key: string): any;
    }
}
declare module "sharepoint/rest/search" {
    import { Queryable, QueryableInstance } from "sharepoint/rest/queryable";
    /**
     * Describes the SearchQuery interface
     */
    export interface SearchQuery {
        /**
         * A string that contains the text for the search query.
         */
        Querytext: string;
        /**
         * A string that contains the text that replaces the query text, as part of a query transform.
         */
        QueryTemplate?: string;
        /**
         * A Boolean value that specifies whether the result tables that are returned for
         * the result block are mixed with the result tables that are returned for the original query.
         */
        EnableInterleaving?: boolean;
        /**
         * A Boolean value that specifies whether stemming is enabled.
         */
        EnableStemming?: boolean;
        /**
         * A Boolean value that specifies whether duplicate items are removed from the results.
         */
        TrimDuplicates?: boolean;
        /**
         * A Boolean value that specifies whether the exact terms in the search query are used to find matches, or if nicknames are used also.
         */
        EnableNicknames?: boolean;
        /**
         * A Boolean value that specifies whether the query uses the FAST Query Language (FQL).
         */
        EnableFql?: boolean;
        /**
         * A Boolean value that specifies whether the phonetic forms of the query terms are used to find matches.
         */
        EnablePhonetic?: boolean;
        /**
         * A Boolean value that specifies whether to perform result type processing for the query.
         */
        BypassResultTypes?: boolean;
        /**
         * A Boolean value that specifies whether to return best bet results for the query.
         * This parameter is used only when EnableQueryRules is set to true, otherwise it is ignored.
         */
        ProcessBestBets?: boolean;
        /**
         * A Boolean value that specifies whether to enable query rules for the query.
         */
        EnableQueryRules?: boolean;
        /**
         * A Boolean value that specifies whether to sort search results.
         */
        EnableSorting?: boolean;
        /**
         * Specifies whether to return block rank log information in the BlockRankLog property of the interleaved result table.
         * A block rank log contains the textual information on the block score and the documents that were de-duplicated.
         */
        GenerateBlockRankLog?: boolean;
        /**
         * The result source ID to use for executing the search query.
         */
        SourceId?: string;
        /**
         * The ID of the ranking model to use for the query.
         */
        RankingModelId?: string;
        /**
         * The first row that is included in the search results that are returned.
         * You use this parameter when you want to implement paging for search results.
         */
        StartRow?: number;
        /**
         * The maximum number of rows overall that are returned in the search results.
         * Compared to RowsPerPage, RowLimit is the maximum number of rows returned overall.
         */
        RowLimit?: number;
        /**
         * The maximum number of rows to return per page.
         * Compared to RowLimit, RowsPerPage refers to the maximum number of rows to return per page,
         * and is used primarily when you want to implement paging for search results.
         */
        RowsPerPage?: number;
        /**
         * The managed properties to return in the search results.
         */
        SelectProperties?: string[];
        /**
         * The locale ID (LCID) for the query.
         */
        Culture?: number;
        /**
         * The set of refinement filters used when issuing a refinement query (FQL)
         */
        RefinementFilters?: string[];
        /**
         * The set of refiners to return in a search result.
         */
        Refiners?: string;
        /**
         * The additional query terms to append to the query.
         */
        HiddenConstraints?: string;
        /**
         * The list of properties by which the search results are ordered.
         */
        SortList?: Sort[];
        /**
         * The amount of time in milliseconds before the query request times out.
         */
        Timeout?: number;
        /**
         * The properties to highlight in the search result summary when the property value matches the search terms entered by the user.
         */
        HithighlightedProperties?: string[];
        /**
         * The type of the client that issued the query.
         */
        ClientType?: string;
        /**
         * The GUID for the user who submitted the search query.
         */
        PersonalizationData?: string;
        /**
         * The URL for the search results page.
         */
        ResultsURL?: string;
        /**
         * Custom tags that identify the query. You can specify multiple query tags
         */
        QueryTag?: string[];
        /**
         * Properties to be used to configure the search query
         */
        Properties?: SearchProperty[];
        /**
         *  A Boolean value that specifies whether to return personal favorites with the search results.
         */
        ProcessPersonalFavorites?: boolean;
        /**
         * The location of the queryparametertemplate.xml file. This file is used to enable anonymous users to make Search REST queries.
         */
        QueryTemplatePropertiesUrl?: string;
        /**
         * Special rules for reordering search results.
         * These rules can specify that documents matching certain conditions are ranked higher or lower in the results.
         * This property applies only when search results are sorted based on rank.
         */
        ReorderingRules?: ReorderingRule[];
        /**
         * The number of properties to show hit highlighting for in the search results.
         */
        HitHighlightedMultivaluePropertyLimit?: number;
        /**
         * A Boolean value that specifies whether the hit highlighted properties can be ordered.
         */
        EnableOrderingHitHighlightedProperty?: boolean;
        /**
         * The managed properties that are used to determine how to collapse individual search results.
         * Results are collapsed into one or a specified number of results if they match any of the individual collapse specifications.
         * In a collapse specification, results are collapsed if their properties match all individual properties in the collapse specification.
         */
        CollapseSpecification?: string;
        /**
         * The locale identifier (LCID) of the user interface
         */
        UIlanguage?: number;
        /**
         * The preferred number of characters to display in the hit-highlighted summary generated for a search result.
         */
        DesiredSnippetLength?: number;
        /**
         * The maximum number of characters to display in the hit-highlighted summary generated for a search result.
         */
        MaxSnippetLength?: number;
        /**
         * The number of characters to display in the result summary for a search result.
         */
        SummaryLength?: number;
    }
    /**
     * Describes the search API
     *
     */
    export class Search extends QueryableInstance {
        /**
         * Creates a new instance of the Search class
         *
         * @param baseUrl The url for the search context
         * @param query The SearchQuery object to execute
         */
        constructor(baseUrl: string | Queryable, path?: string);
        /**
         * .......
         * @returns Promise
         */
        execute(query: SearchQuery): Promise<SearchResults>;
    }
    /**
     * Describes the SearchResults class, which returns the formatted and raw version of the query response
     */
    export class SearchResults {
        PrimarySearchResults: any;
        RawSearchResults: any;
        RowCount: number;
        TotalRows: number;
        TotalRowsIncludingDuplicates: number;
        ElapsedTime: number;
        /**
         * Creates a new instance of the SearchResult class
         *
         */
        constructor(rawResponse: any);
        /**
         * Formats a search results array
         *
         * @param rawResults The array to process
         */
        protected formatSearchResults(rawResults: Array<any> | any): SearchResult[];
    }
    /**
     * Describes the SearchResult class
     */
    export class SearchResult {
        /**
         * Creates a new instance of the SearchResult class
         *
         */
        constructor(rawItem: Array<any> | any);
    }
    /**
     * Defines how search results are sorted.
     */
    export interface Sort {
        /**
         * The name for a property by which the search results are ordered.
         */
        Property: string;
        /**
         * The direction in which search results are ordered.
         */
        Direction: SortDirection;
    }
    /**
     * Defines one search property
     */
    export interface SearchProperty {
        Name: string;
        Value: SearchPropertyValue;
    }
    /**
     * Defines one search property value
     */
    export interface SearchPropertyValue {
        StrVal: string;
        QueryPropertyValueTypeIndex: QueryPropertyValueType;
    }
    /**
     * defines the SortDirection enum
     */
    export enum SortDirection {
        Ascending = 0,
        Descending = 1,
        FQLFormula = 2,
    }
    /**
     * Defines how ReorderingRule interface, used for reordering results
     */
    export interface ReorderingRule {
        /**
         * The value to match on
         */
        MatchValue: string;
        /**
         * The rank boosting
         */
        Boost: number;
        /**
        * The rank boosting
        */
        MatchType: ReorderingRuleMatchType;
    }
    /**
     * defines the ReorderingRuleMatchType  enum
     */
    export enum ReorderingRuleMatchType {
        ResultContainsKeyword = 0,
        TitleContainsKeyword = 1,
        TitleMatchesKeyword = 2,
        UrlStartsWith = 3,
        UrlExactlyMatches = 4,
        ContentTypeIs = 5,
        FileExtensionMatches = 6,
        ResultHasTag = 7,
        ManualCondition = 8,
    }
    /**
     * Specifies the type value for the property
     */
    export enum QueryPropertyValueType {
        None = 0,
        StringType = 1,
        Int32TYpe = 2,
        BooleanType = 3,
        StringArrayType = 4,
        UnSupportedType = 5,
    }
}
declare module "sharepoint/rest/searchsuggest" {
    import { Queryable, QueryableInstance } from "sharepoint/rest/queryable";
    /**
     * Defines a query execute against the search/suggest endpoint (see https://msdn.microsoft.com/en-us/library/office/dn194079.aspx)
     */
    export interface SearchSuggestQuery {
        /**
         * A string that contains the text for the search query.
         */
        querytext: string;
        /**
         * The number of query suggestions to retrieve. Must be greater than zero (0). The default value is 5.
         */
        count?: number;
        /**
         * The number of personal results to retrieve. Must be greater than zero (0). The default value is 5.
         */
        personalCount?: number;
        /**
         * A Boolean value that specifies whether to retrieve pre-query or post-query suggestions. true to return pre-query suggestions; otherwise, false. The default value is false.
         */
        preQuery?: boolean;
        /**
         * A Boolean value that specifies whether to hit-highlight or format in bold the query suggestions. true to format in bold the terms in the returned query suggestions
         * that match terms in the specified query; otherwise, false. The default value is true.
         */
        hitHighlighting?: boolean;
        /**
         * A Boolean value that specifies whether to capitalize the first letter in each term in the returned query suggestions. true to capitalize the first letter in each term;
         * otherwise, false. The default value is false.
         */
        capitalize?: boolean;
        /**
         * The locale ID (LCID) for the query (see https://msdn.microsoft.com/en-us/library/cc233982.aspx).
         */
        culture?: string;
        /**
         * A Boolean value that specifies whether stemming is enabled. true to enable stemming; otherwise, false. The default value is true.
         */
        stemming?: boolean;
        /**
         * A Boolean value that specifies whether to include people names in the returned query suggestions. true to include people names in the returned query suggestions;
         * otherwise, false. The default value is true.
         */
        includePeople?: boolean;
        /**
         * A Boolean value that specifies whether to turn on query rules for this query. true to turn on query rules; otherwise, false. The default value is true.
         */
        queryRules?: boolean;
        /**
         * A Boolean value that specifies whether to return query suggestions for prefix matches. true to return query suggestions based on prefix matches, otherwise, false when
         * query suggestions should match the full query word.
         */
        prefixMatch?: boolean;
    }
    export class SearchSuggest extends QueryableInstance {
        constructor(baseUrl: string | Queryable, path?: string);
        execute(query: SearchSuggestQuery): Promise<SearchSuggestResult>;
        private mapQueryToQueryString(query);
    }
    export class SearchSuggestResult {
        PeopleNames: string[];
        PersonalResults: PersonalResultSuggestion[];
        Queries: any[];
        constructor(json: any);
    }
    export interface PersonalResultSuggestion {
        HighlightedTitle?: string;
        IsBestBet?: boolean;
        Title?: string;
        TypeId?: string;
        Url?: string;
    }
}
declare module "sharepoint/rest/site" {
    import { Queryable, QueryableInstance } from "sharepoint/rest/queryable";
    import { Web } from "sharepoint/rest/webs";
    import { UserCustomActions } from "sharepoint/rest/usercustomactions";
    import { ContextInfo, DocumentLibraryInformation } from "sharepoint/rest/types";
    import { ODataBatch } from "sharepoint/rest/odata";
    /**
     * Describes a site collection
     *
     */
    export class Site extends QueryableInstance {
        /**
         * Creates a new instance of the RoleAssignments class
         *
         * @param baseUrl The url or Queryable which forms the parent of this fields collection
         */
        constructor(baseUrl: string | Queryable, path?: string);
        /**
         * Gets the root web of the site collection
         *
         */
        readonly rootWeb: Web;
        /**
         * Get all custom actions on a site collection
         *
         */
        readonly userCustomActions: UserCustomActions;
        /**
         * Gets the context information for the site.
         */
        getContextInfo(): Promise<ContextInfo>;
        /**
         * Gets the document libraries on a site. Static method. (SharePoint Online only)
         *
         * @param absoluteWebUrl The absolute url of the web whose document libraries should be returned
         */
        getDocumentLibraries(absoluteWebUrl: string): Promise<DocumentLibraryInformation[]>;
        /**
         * Gets the site URL from a page URL.
         *
         * @param absolutePageUrl The absolute url of the page
         */
        getWebUrlFromPageUrl(absolutePageUrl: string): Promise<string>;
        /**
         * Creates a new batch for requests within the context of context this site
         *
         */
        createBatch(): ODataBatch;
    }
}
declare module "utils/files" {
    /**
     * Reads a blob as text
     *
     * @param blob The data to read
     */
    export function readBlobAsText(blob: Blob): Promise<string>;
    /**
     * Reads a blob into an array buffer
     *
     * @param blob The data to read
     */
    export function readBlobAsArrayBuffer(blob: Blob): Promise<ArrayBuffer>;
}
declare module "sharepoint/rest/userprofiles" {
    import { Queryable, QueryableInstance, QueryableCollection } from "sharepoint/rest/queryable";
    import * as Types from "sharepoint/rest/types";
    export class UserProfileQuery extends QueryableInstance {
        private profileLoader;
        constructor(baseUrl: string | Queryable, path?: string);
        /**
         * The URL of the edit profile page for the current user.
         */
        readonly editProfileLink: Promise<string>;
        /**
         * A Boolean value that indicates whether the current user's People I'm Following list is public.
         */
        readonly isMyPeopleListPublic: Promise<boolean>;
        /**
         * A Boolean value that indicates whether the current user's People I'm Following list is public.
         *
         * @param loginName The account name of the user
         */
        amIFollowedBy(loginName: string): Promise<boolean>;
        /**
         * Checks whether the current user is following the specified user.
         *
         * @param loginName The account name of the user
         */
        amIFollowing(loginName: string): Promise<boolean>;
        /**
         * Gets tags that the user is following.
         *
         * @param maxCount The maximum number of tags to get.
         */
        getFollowedTags(maxCount?: number): Promise<string[]>;
        /**
         * Gets the people who are following the specified user.
         *
         * @param loginName The account name of the user.
         */
        getFollowersFor(loginName: string): Promise<any[]>;
        /**
         * Gets the people who are following the current user.
         *
         */
        readonly myFollowers: QueryableCollection;
        /**
         * Gets user properties for the current user.
         *
         */
        readonly myProperties: QueryableInstance;
        /**
         * Gets the people who the specified user is following.
         *
         * @param loginName The account name of the user.
         */
        getPeopleFollowedBy(loginName: string): Promise<any[]>;
        /**
         * Gets user properties for the specified user.
         *
         * @param loginName The account name of the user.
         */
        getPropertiesFor(loginName: string): Promise<any[]>;
        /**
         * Gets the most popular tags.
         *
         */
        readonly trendingTags: Promise<Types.HashTagCollection>;
        /**
         * Gets the specified user profile property for the specified user.
         *
         * @param loginName The account name of the user.
         * @param propertyName The case-sensitive name of the property to get.
         */
        getUserProfilePropertyFor(loginName: string, propertyName: string): Promise<string>;
        /**
         * Removes the specified user from the user's list of suggested people to follow.
         *
         * @param loginName The account name of the user.
         */
        hideSuggestion(loginName: string): Promise<void>;
        /**
         * Checks whether the first user is following the second user.
         *
         * @param follower The account name of the user who might be following followee.
         * @param followee The account name of the user who might be followed.
         */
        isFollowing(follower: string, followee: string): Promise<boolean>;
        /**
         * Uploads and sets the user profile picture
         *
         * @param profilePicSource Blob data representing the user's picture
         */
        setMyProfilePic(profilePicSource: Blob): Promise<void>;
        /**
         * Provisions one or more users' personal sites. (My Site administrator on SharePoint Online only)
         *
         * @param emails The email addresses of the users to provision sites for
         */
        createPersonalSiteEnqueueBulk(...emails: string[]): Promise<void>;
        /**
         * Gets the user profile of the site owner.
         *
         */
        readonly ownerUserProfile: Promise<Types.UserProfile>;
        /**
         * Gets the user profile that corresponds to the current user.
         */
        readonly userProfile: Promise<any>;
        /**
         * Enqueues creating a personal site for this user, which can be used to share documents, web pages, and other files.
         *
         * @param interactiveRequest true if interactively (web) initiated request, or false if non-interactively (client) initiated request
         */
        createPersonalSite(interactiveRequest?: boolean): Promise<void>;
        /**
         * Sets the privacy settings for this profile.
         *
         * @param share true to make all social data public; false to make all social data private.
         */
        shareAllSocialData(share: any): Promise<void>;
    }
}
declare module "sharepoint/rest/rest" {
    import { SearchQuery, SearchResults } from "sharepoint/rest/search";
    import { SearchSuggestQuery, SearchSuggestResult } from "sharepoint/rest/searchsuggest";
    import { Site } from "sharepoint/rest/site";
    import { Web } from "sharepoint/rest/webs";
    import { UserProfileQuery } from "sharepoint/rest/userprofiles";
    import { ODataBatch } from "sharepoint/rest/odata";
    /**
     * Root of the SharePoint REST module
     */
    export class Rest {
        /**
         * Executes a search against this web context
         *
         * @param query The SearchQuery definition
         */
        searchSuggest(query: string | SearchSuggestQuery): Promise<SearchSuggestResult>;
        /**
         * Executes a search against this web context
         *
         * @param query The SearchQuery definition
         */
        search(query: string | SearchQuery): Promise<SearchResults>;
        /**
         * Begins a site collection scoped REST request
         *
         */
        readonly site: Site;
        /**
         * Begins a web scoped REST request
         *
         */
        readonly web: Web;
        /**
         * Access to user profile methods
         *
         */
        readonly profiles: UserProfileQuery;
        /**
         * Creates a new batch object for use with the Queryable.addToBatch method
         *
         */
        createBatch(): ODataBatch;
        /**
         * Begins a cross-domain, host site scoped REST request, for use in add-in webs
         *
         * @param addInWebUrl The absolute url of the add-in web
         * @param hostWebUrl The absolute url of the host web
         */
        crossDomainSite(addInWebUrl: string, hostWebUrl: string): Site;
        /**
         * Begins a cross-domain, host web scoped REST request, for use in add-in webs
         *
         * @param addInWebUrl The absolute url of the add-in web
         * @param hostWebUrl The absolute url of the host web
         */
        crossDomainWeb(addInWebUrl: string, hostWebUrl: string): Web;
        /**
         * Implements the creation of cross domain REST urls
         *
         * @param factory The constructor of the object to create Site | Web
         * @param addInWebUrl The absolute url of the add-in web
         * @param hostWebUrl The absolute url of the host web
         * @param urlPart String part to append to the url "site" | "web"
         */
        private _cdImpl<T>(factory, addInWebUrl, hostWebUrl, urlPart);
    }
}
declare module "sharepoint/rest/index" {
    export * from "sharepoint/rest/caching";
    export { FieldAddResult, FieldUpdateResult } from "sharepoint/rest/fields";
    export { CheckinType, FileAddResult, WebPartsPersonalizationScope, MoveOperations, TemplateFileType, TextFileParser, BlobFileParser, BufferFileParser, ChunkedFileUploadProgressData } from "sharepoint/rest/files";
    export { FolderAddResult } from "sharepoint/rest/folders";
    export { ItemAddResult, ItemUpdateResult, PagedItemCollection } from "sharepoint/rest/items";
    export { ListAddResult, ListUpdateResult, ListEnsureResult } from "sharepoint/rest/lists";
    export { extractOdataId, ODataParser, ODataParserBase, ODataDefaultParser, ODataRaw, ODataValue, ODataEntity, ODataEntityArray } from "sharepoint/rest/odata";
    export { RoleDefinitionUpdateResult, RoleDefinitionAddResult, RoleDefinitionBindings } from "sharepoint/rest/roles";
    export { Search, SearchProperty, SearchPropertyValue, SearchQuery, SearchResult, SearchResults, Sort, SortDirection, ReorderingRule, ReorderingRuleMatchType, QueryPropertyValueType } from "sharepoint/rest/search";
    export { SearchSuggest, SearchSuggestQuery, SearchSuggestResult, PersonalResultSuggestion } from "sharepoint/rest/searchsuggest";
    export { Site } from "sharepoint/rest/site";
    export { SiteGroupAddResult } from "sharepoint/rest/sitegroups";
    export { UserUpdateResult, UserProps } from "sharepoint/rest/siteusers";
    export { SubscriptionAddResult, SubscriptionUpdateResult } from "sharepoint/rest/subscriptions";
    export * from "sharepoint/rest/types";
    export { UserCustomActionAddResult, UserCustomActionUpdateResult } from "sharepoint/rest/usercustomactions";
    export { ViewAddResult, ViewUpdateResult } from "sharepoint/rest/views";
    export { Web, WebAddResult, WebUpdateResult, GetCatalogResult } from "sharepoint/rest/webs";
}
declare module "types/index" {
    export * from "sharepoint/rest/index";
    export { FetchOptions, HttpClient } from "net/httpclient";
    export { IConfigurationProvider } from "configuration/configuration";
    export { NodeClientData, LibraryConfiguration } from "configuration/pnplibconfig";
    export { TypedHash, Dictionary } from "collections/collections";
    export { Util } from "utils/util";
    export * from "utils/logging";
}
declare module "pnp" {
    import { Util } from "utils/util";
    import { PnPClientStorage } from "utils/storage";
    import { Settings } from "configuration/configuration";
    import { Logger } from "utils/logging";
    import { Rest } from "sharepoint/rest/rest";
    import { LibraryConfiguration } from "configuration/pnplibconfig";
    /**
     * Root class of the Patterns and Practices namespace, provides an entry point to the library
     */
    /**
     * Utility methods
     */
    export const util: typeof Util;
    /**
     * Provides access to the REST interface
     */
    export const sp: Rest;
    /**
     * Provides access to local and session storage
     */
    export const storage: PnPClientStorage;
    /**
     * Global configuration instance to which providers can be added
     */
    export const config: Settings;
    /**
     * Global logging instance to which subscribers can be registered and messages written
     */
    export const log: typeof Logger;
    /**
     * Allows for the configuration of the library
     */
    export const setup: (config: LibraryConfiguration) => void;
    /**
     * Expose a subset of classes from the library for public consumption
     */
    export * from "types/index";
    let Def: {
        config: Settings;
        log: typeof Logger;
        setup: (config: LibraryConfiguration) => void;
        sp: Rest;
        storage: PnPClientStorage;
        util: typeof Util;
    };
    export default Def;
}
declare module "net/nodefetchclientbrowser" {
    import { HttpClientImpl } from "net/httpclient";
    /**
     * This module is substituted for the NodeFetchClient.ts during the packaging process. This helps to reduce the pnp.js file size by
     * not including all of the node dependencies
     */
    export class NodeFetchClient implements HttpClientImpl {
        siteUrl: string;
        private _clientId;
        private _clientSecret;
        private _realm;
        constructor(siteUrl: string, _clientId: string, _clientSecret: string, _realm?: string);
        /**
         * Always throws an error that NodeFetchClient is not supported for use in the browser
         */
        fetch(url: string, options: any): Promise<Response>;
    }
}
declare module "types/locale" {
    export enum Locale {
        AfrikaansSouthAfrica = 1078,
        AlbanianAlbania = 1052,
        Alsatian = 1156,
        AmharicEthiopia = 1118,
        ArabicSaudiArabia = 1025,
        ArabicAlgeria = 5121,
        ArabicBahrain = 15361,
        ArabicEgypt = 3073,
        ArabicIraq = 2049,
        ArabicJordan = 11265,
        ArabicKuwait = 13313,
        ArabicLebanon = 12289,
        ArabicLibya = 4097,
        ArabicMorocco = 6145,
        ArabicOman = 8193,
        ArabicQatar = 16385,
        ArabicSyria = 10241,
        ArabicTunisia = 7169,
        ArabicUAE = 14337,
        ArabicYemen = 9217,
        ArmenianArmenia = 1067,
        Assamese = 1101,
        AzeriCyrillic = 2092,
        AzeriLatin = 1068,
        Bashkir = 1133,
        Basque = 1069,
        Belarusian = 1059,
        BengaliIndia = 1093,
        BengaliBangladesh = 2117,
        BosnianBosniaHerzegovina = 5146,
        Breton = 1150,
        Bulgarian = 1026,
        Burmese = 1109,
        Catalan = 1027,
        CherokeeUnitedStates = 1116,
        ChinesePeoplesRepublicofChina = 2052,
        ChineseSingapore = 4100,
        ChineseTaiwan = 1028,
        ChineseHongKongSAR = 3076,
        ChineseMacaoSAR = 5124,
        Corsican = 1155,
        Croatian = 1050,
        CroatianBosniaHerzegovina = 4122,
        Czech = 1029,
        Danish = 1030,
        Dari = 1164,
        Divehi = 1125,
        DutchNetherlands = 1043,
        DutchBelgium = 2067,
        Edo = 1126,
        EnglishUnitedStates = 1033,
        EnglishUnitedKingdom = 2057,
        EnglishAustralia = 3081,
        EnglishBelize = 10249,
        EnglishCanada = 4105,
        EnglishCaribbean = 9225,
        EnglishHongKongSAR = 15369,
        EnglishIndia = 16393,
        EnglishIndonesia = 14345,
        EnglishIreland = 6153,
        EnglishJamaica = 8201,
        EnglishMalaysia = 17417,
        EnglishNewZealand = 5129,
        EnglishPhilippines = 13321,
        EnglishSingapore = 18441,
        EnglishSouthAfrica = 7177,
        EnglishTrinidad = 11273,
        EnglishZimbabwe = 12297,
        Estonian = 1061,
        Faroese = 1080,
        Farsi = 1065,
        Filipino = 1124,
        Finnish = 1035,
        FrenchFrance = 1036,
        FrenchBelgium = 2060,
        FrenchCameroon = 11276,
        FrenchCanada = 3084,
        FrenchDemocraticRepofCongo = 9228,
        FrenchCotedIvoire = 12300,
        FrenchHaiti = 15372,
        FrenchLuxembourg = 5132,
        FrenchMali = 13324,
        FrenchMonaco = 6156,
        FrenchMorocco = 14348,
        FrenchNorthAfrica = 58380,
        FrenchReunion = 8204,
        FrenchSenegal = 10252,
        FrenchSwitzerland = 4108,
        FrenchWestIndies = 7180,
        FrisianNetherlands = 1122,
        FulfuldeNigeria = 1127,
        FYROMacedonian = 1071,
        Galician = 1110,
        Georgian = 1079,
        GermanGermany = 1031,
        GermanAustria = 3079,
        GermanLiechtenstein = 5127,
        GermanLuxembourg = 4103,
        GermanSwitzerland = 2055,
        Greek = 1032,
        Greenlandic = 1135,
        GuaraniParaguay = 1140,
        Gujarati = 1095,
        HausaNigeria = 1128,
        HawaiianUnitedStates = 1141,
        Hebrew = 1037,
        Hindi = 1081,
        Hungarian = 1038,
        IbibioNigeria = 1129,
        Icelandic = 1039,
        IgboNigeria = 1136,
        Indonesian = 1057,
        Inuktitut = 1117,
        Irish = 2108,
        ItalianItaly = 1040,
        ItalianSwitzerland = 2064,
        Japanese = 1041,
        Kiche = 1158,
        Kannada = 1099,
        KanuriNigeria = 1137,
        Kashmiri = 2144,
        KashmiriArabic = 1120,
        Kazakh = 1087,
        Khmer = 1107,
        Kinyarwanda = 1159,
        Konkani = 1111,
        Korean = 1042,
        KyrgyzCyrillic = 1088,
        Lao = 1108,
        Latin = 1142,
        Latvian = 1062,
        Lithuanian = 1063,
        Luxembourgish = 1134,
        MalayMalaysia = 1086,
        MalayBruneiDarussalam = 2110,
        Malayalam = 1100,
        Maltese = 1082,
        Manipuri = 1112,
        MaoriNewZealand = 1153,
        Mapudungun = 1146,
        Marathi = 1102,
        Mohawk = 1148,
        MongolianCyrillic = 1104,
        MongolianMongolian = 2128,
        Nepali = 1121,
        NepaliIndia = 2145,
        NorwegianBokmål = 1044,
        NorwegianNynorsk = 2068,
        Occitan = 1154,
        Oriya = 1096,
        Oromo = 1138,
        Papiamentu = 1145,
        Pashto = 1123,
        Polish = 1045,
        PortugueseBrazil = 1046,
        PortuguesePortugal = 2070,
        Punjabi = 1094,
        PunjabiPakistan = 2118,
        QuechaBolivia = 1131,
        QuechaEcuador = 2155,
        QuechaPeru = 3179,
        RhaetoRomanic = 1047,
        Romanian = 1048,
        RomanianMoldava = 2072,
        Russian = 1049,
        RussianMoldava = 2073,
        SamiLappish = 1083,
        Sanskrit = 1103,
        ScottishGaelic = 1084,
        Sepedi = 1132,
        SerbianCyrillic = 3098,
        SerbianLatin = 2074,
        SindhiIndia = 1113,
        SindhiPakistan = 2137,
        SinhaleseSriLanka = 1115,
        Slovak = 1051,
        Slovenian = 1060,
        Somali = 1143,
        Sorbian = 1070,
        SpanishSpainModernSort = 3082,
        SpanishSpainTraditionalSort = 1034,
        SpanishArgentina = 11274,
        SpanishBolivia = 16394,
        SpanishChile = 13322,
        SpanishColombia = 9226,
        SpanishCostaRica = 5130,
        SpanishDominicanRepublic = 7178,
        SpanishEcuador = 12298,
        SpanishElSalvador = 17418,
        SpanishGuatemala = 4106,
        SpanishHonduras = 18442,
        SpanishLatinAmerica = 22538,
        SpanishMexico = 2058,
        SpanishNicaragua = 19466,
        SpanishPanama = 6154,
        SpanishParaguay = 15370,
        SpanishPeru = 10250,
        SpanishPuertoRico = 20490,
        SpanishUnitedStates = 21514,
        SpanishUruguay = 14346,
        SpanishVenezuela = 8202,
        Sutu = 1072,
        Swahili = 1089,
        Swedish = 1053,
        SwedishFinland = 2077,
        Syriac = 1114,
        Tajik = 1064,
        TamazightArabic = 1119,
        TamazightLatin = 2143,
        Tamil = 1097,
        Tatar = 1092,
        Telugu = 1098,
        Thai = 1054,
        TibetanBhutan = 2129,
        TibetanPeoplesRepublicofChina = 1105,
        TigrignaEritrea = 2163,
        TigrignaEthiopia = 1139,
        Tsonga = 1073,
        Tswana = 1074,
        Turkish = 1055,
        Turkmen = 1090,
        UighurChina = 1152,
        Ukrainian = 1058,
        Urdu = 1056,
        UrduIndia = 2080,
        UzbekCyrillic = 2115,
        UzbekLatin = 1091,
        Venda = 1075,
        Vietnamese = 1066,
        Welsh = 1106,
        Wolof = 1160,
        Xhosa = 1076,
        Yakut = 1157,
        Yi = 1144,
        Yiddish = 1085,
        Yoruba = 1130,
        Zulu = 1077,
        HIDHumanInterfaceDevice = 1279,
    }
}
declare module "sharepoint/provisioning/provisioningstep" {
    /**
     * Describes a ProvisioningStep
     */
    export class ProvisioningStep {
        private name;
        private index;
        private objects;
        private parameters;
        private handler;
        /**
         * Executes the ProvisioningStep function
         *
         * @param dependentPromise The promise the ProvisioningStep is dependent on
         */
        execute(dependentPromise?: any): any;
        /**
         * Creates a new instance of the ProvisioningStep class
         */
        constructor(name: string, index: number, objects: any, parameters: any, handler: any);
    }
}
declare module "sharepoint/provisioning/util" {
    export class Util {
        /**
         * Make URL relative to host
         *
         * @param url The URL to make relative
         */
        static getRelativeUrl(url: string): string;
        /**
         * Replaces URL tokens in a string
         */
        static replaceUrlTokens(url: string): string;
        static encodePropertyKey(propKey: any): string;
    }
}
declare module "sharepoint/provisioning/objecthandlers/objecthandlerbase" {
    import { HttpClient } from "net/httpclient";
    /**
     * Describes the Object Handler Base
     */
    export class ObjectHandlerBase {
        httpClient: HttpClient;
        private name;
        /**
         * Creates a new instance of the ObjectHandlerBase class
         */
        constructor(name: string);
        /**
         * Provisioning objects
         */
        ProvisionObjects(objects: any, parameters?: any): Promise<{}>;
        /**
         * Writes to Logger when scope has started
         */
        scope_started(): void;
        /**
         * Writes to Logger when scope has stopped
         */
        scope_ended(): void;
    }
}
declare module "sharepoint/provisioning/schema/inavigationnode" {
    export interface INavigationNode {
        Title: string;
        Url: string;
        Children: Array<INavigationNode>;
    }
}
declare module "sharepoint/provisioning/schema/inavigation" {
    import { INavigationNode } from "sharepoint/provisioning/schema/inavigationnode";
    export interface INavigation {
        UseShared: boolean;
        QuickLaunch: Array<INavigationNode>;
    }
}
declare module "sharepoint/provisioning/objecthandlers/objectnavigation" {
    import { ObjectHandlerBase } from "sharepoint/provisioning/objecthandlers/objecthandlerbase";
    import { INavigation } from "sharepoint/provisioning/schema/inavigation";
    /**
     * Describes the Navigation Object Handler
     */
    export class ObjectNavigation extends ObjectHandlerBase {
        /**
         * Creates a new instance of the ObjectNavigation class
         */
        constructor();
        /**
         * Provision Navigation nodes
         *
         * @param object The navigation settings and nodes to provision
         */
        ProvisionObjects(object: INavigation): Promise<{}>;
        /**
         * Retrieves the node with the given title from a collection of SP.NavigationNode
         */
        private getNodeFromCollectionByTitle(nodeCollection, title);
        private ConfigureQuickLaunch(nodes, clientContext, httpClient, navigation);
    }
}
declare module "sharepoint/provisioning/schema/IPropertyBagEntry" {
    export interface IPropertyBagEntry {
        Key: string;
        Value: string;
        Indexed: boolean;
    }
}
declare module "sharepoint/provisioning/objecthandlers/objectpropertybagentries" {
    import { ObjectHandlerBase } from "sharepoint/provisioning/objecthandlers/objecthandlerbase";
    import { IPropertyBagEntry } from "sharepoint/provisioning/schema/IPropertyBagEntry";
    /**
     * Describes the Property Bag Entries Object Handler
     */
    export class ObjectPropertyBagEntries extends ObjectHandlerBase {
        /**
         * Creates a new instance of the ObjectPropertyBagEntries class
         */
        constructor();
        /**
         * Provision Property Bag Entries
         *
         * @param entries The entries to provision
         */
        ProvisionObjects(entries: Array<IPropertyBagEntry>): Promise<{}>;
    }
}
declare module "sharepoint/provisioning/schema/IFeature" {
    export interface IFeature {
        ID: string;
        Deactivate: boolean;
        Description: string;
    }
}
declare module "sharepoint/provisioning/objecthandlers/objectfeatures" {
    import { ObjectHandlerBase } from "sharepoint/provisioning/objecthandlers/objecthandlerbase";
    import { IFeature } from "sharepoint/provisioning/schema/IFeature";
    /**
     * Describes the Features Object Handler
     */
    export class ObjectFeatures extends ObjectHandlerBase {
        /**
         * Creates a new instance of the ObjectFeatures class
         */
        constructor();
        /**
         * Provisioning features
         *
         * @paramm features The features to provision
         */
        ProvisionObjects(features: Array<IFeature>): Promise<{}>;
    }
}
declare module "sharepoint/provisioning/schema/IWebSettings" {
    export interface IWebSettings {
        WelcomePage: string;
        AlternateCssUrl: string;
        SaveSiteAsTemplateEnabled: boolean;
        MasterUrl: string;
        CustomMasterUrl: string;
        RecycleBinEnabled: boolean;
        TreeViewEnabled: boolean;
        QuickLaunchEnabled: boolean;
    }
}
declare module "sharepoint/provisioning/objecthandlers/objectwebsettings" {
    import { ObjectHandlerBase } from "sharepoint/provisioning/objecthandlers/objecthandlerbase";
    import { IWebSettings } from "sharepoint/provisioning/schema/IWebSettings";
    /**
     * Describes the Web Settings Object Handler
     */
    export class ObjectWebSettings extends ObjectHandlerBase {
        /**
         * Creates a new instance of the ObjectWebSettings class
         */
        constructor();
        /**
         * Provision Web Settings
         *
         * @param object The Web Settings to provision
         */
        ProvisionObjects(object: IWebSettings): Promise<{}>;
    }
}
declare module "sharepoint/provisioning/schema/IComposedLook" {
    export interface IComposedLook {
        ColorPaletteUrl: string;
        FontSchemeUrl: string;
        BackgroundImageUrl: string;
    }
}
declare module "sharepoint/provisioning/objecthandlers/objectcomposedlook" {
    import { IComposedLook } from "sharepoint/provisioning/schema/IComposedLook";
    import { ObjectHandlerBase } from "sharepoint/provisioning/objecthandlers/objecthandlerbase";
    /**
     * Describes the Composed Look Object Handler
     */
    export class ObjectComposedLook extends ObjectHandlerBase {
        /**
         * Creates a new instance of the ObjectComposedLook class
         */
        constructor();
        /**
         * Provisioning Composed Look
         *
         * @param object The Composed Look to provision
         */
        ProvisionObjects(object: IComposedLook): Promise<{}>;
    }
}
declare module "sharepoint/provisioning/schema/ICustomAction" {
    export interface ICustomAction {
        CommandUIExtension: any;
        Description: string;
        Group: string;
        ImageUrl: string;
        Location: string;
        Name: string;
        RegistrationId: string;
        RegistrationType: any;
        Rights: any;
        ScriptBlock: string;
        ScriptSrc: string;
        Sequence: number;
        Title: string;
        Url: string;
    }
}
declare module "sharepoint/provisioning/objecthandlers/objectcustomactions" {
    import { ObjectHandlerBase } from "sharepoint/provisioning/objecthandlers/objecthandlerbase";
    import { ICustomAction } from "sharepoint/provisioning/schema/ICustomAction";
    /**
     * Describes the Custom Actions Object Handler
     */
    export class ObjectCustomActions extends ObjectHandlerBase {
        /**
         * Creates a new instance of the ObjectCustomActions class
         */
        constructor();
        /**
         * Provisioning Custom Actions
         *
         * @param customactions The Custom Actions to provision
         */
        ProvisionObjects(customactions: Array<ICustomAction>): Promise<{}>;
    }
}
declare module "sharepoint/provisioning/schema/IContents" {
    export interface IContents {
        Xml: string;
        FileUrl: string;
    }
}
declare module "sharepoint/provisioning/schema/IWebPart" {
    import { IContents } from "sharepoint/provisioning/schema/IContents";
    export interface IWebPart {
        Title: string;
        Order: number;
        Zone: string;
        Row: number;
        Column: number;
        Contents: IContents;
    }
}
declare module "sharepoint/provisioning/schema/IHiddenView" {
    export interface IHiddenView {
        List: string;
        Url: string;
        Paged: boolean;
        Query: string;
        RowLimit: number;
        Scope: number;
        ViewFields: Array<string>;
    }
}
declare module "sharepoint/provisioning/schema/IFile" {
    import { IWebPart } from "sharepoint/provisioning/schema/IWebPart";
    import { IHiddenView } from "sharepoint/provisioning/schema/IHiddenView";
    export interface IFile {
        Overwrite: boolean;
        Dest: string;
        Src: string;
        Properties: Object;
        RemoveExistingWebParts: boolean;
        WebParts: Array<IWebPart>;
        Views: Array<IHiddenView>;
    }
}
declare module "sharepoint/provisioning/objecthandlers/objectfiles" {
    import { ObjectHandlerBase } from "sharepoint/provisioning/objecthandlers/objecthandlerbase";
    import { IFile } from "sharepoint/provisioning/schema/IFile";
    /**
     * Describes the Files Object Handler
     */
    export class ObjectFiles extends ObjectHandlerBase {
        /**
         * Creates a new instance of the ObjectFiles class
         */
        constructor();
        /**
         * Provisioning Files
         *
         * @param objects The files to provisiion
         */
        ProvisionObjects(objects: Array<IFile>): Promise<{}>;
        private RemoveWebPartsFromFileIfSpecified(clientContext, limitedWebPartManager, shouldRemoveExisting);
        private GetWebPartXml(webParts);
        private AddWebPartsToWebPartPage(dest, src, webParts, shouldRemoveExisting);
        private ApplyFileProperties(dest, fileProperties);
        private GetViewFromCollectionByUrl(viewCollection, url);
        private ModifyHiddenViews(objects);
        private GetFolderFromFilePath(filePath);
        private GetFilenameFromFilePath(filePath);
    }
}
declare module "sharepoint/provisioning/sequencer/sequencer" {
    /**
     * Descibes a Sequencer
     */
    export class Sequencer {
        private functions;
        private parameter;
        private scope;
        /**
         * Creates a new instance of the Sequencer class, and declare private variables
         */
        constructor(functions: Array<any>, parameter: any, scope: any);
        /**
         * Executes the functions in sequence using DeferredObject
         */
        execute(progressFunction?: (s: Sequencer, index: number, functions: any[]) => void): Promise<void>;
    }
}
declare module "sharepoint/provisioning/schema/IFolder" {
    export interface IFolder {
        Name: string;
        DefaultValues: Object;
    }
}
declare module "sharepoint/provisioning/schema/IListInstanceFieldRef" {
    export interface IListInstanceFieldRef {
        Name: string;
    }
}
declare module "sharepoint/provisioning/schema/IField" {
    export interface IField {
        ShowInDisplayForm: boolean;
        ShowInEditForm: boolean;
        ShowInNewForm: boolean;
        CanBeDeleted: boolean;
        DefaultValue: string;
        Description: string;
        EnforceUniqueValues: boolean;
        Direction: string;
        EntityPropertyName: string;
        FieldTypeKind: any;
        Filterable: boolean;
        Group: string;
        Hidden: boolean;
        ID: string;
        Indexed: boolean;
        InternalName: string;
        JsLink: string;
        ReadOnlyField: boolean;
        Required: boolean;
        SchemaXml: string;
        StaticName: string;
        Title: string;
        TypeAsString: string;
        TypeDisplayName: string;
        TypeShortDescription: string;
        ValidationFormula: string;
        ValidationMessage: string;
        Type: string;
        Formula: string;
    }
}
declare module "sharepoint/provisioning/schema/IView" {
    export interface IView {
        Title: string;
        Paged: boolean;
        PersonalView: boolean;
        Query: string;
        RowLimit: number;
        Scope: number;
        SetAsDefaultView: boolean;
        ViewFields: Array<string>;
        ViewTypeKind: string;
    }
}
declare module "sharepoint/provisioning/schema/IRoleAssignment" {
    export interface IRoleAssignment {
        Principal: string;
        RoleDefinition: any;
    }
}
declare module "sharepoint/provisioning/schema/ISecurity" {
    import { IRoleAssignment } from "sharepoint/provisioning/schema/IRoleAssignment";
    export interface ISecurity {
        BreakRoleInheritance: boolean;
        CopyRoleAssignments: boolean;
        ClearSubscopes: boolean;
        RoleAssignments: Array<IRoleAssignment>;
    }
}
declare module "sharepoint/provisioning/schema/IContentTypeBinding" {
    export interface IContentTypeBinding {
        ContentTypeId: string;
    }
}
declare module "sharepoint/provisioning/schema/IListInstance" {
    import { IFolder } from "sharepoint/provisioning/schema/IFolder";
    import { IListInstanceFieldRef } from "sharepoint/provisioning/schema/IListInstanceFieldRef";
    import { IField } from "sharepoint/provisioning/schema/IField";
    import { IView } from "sharepoint/provisioning/schema/IView";
    import { ISecurity } from "sharepoint/provisioning/schema/ISecurity";
    import { IContentTypeBinding } from "sharepoint/provisioning/schema/IContentTypeBinding";
    export interface IListInstance {
        Title: string;
        Url: string;
        Description: string;
        DocumentTemplate: string;
        OnQuickLaunch: boolean;
        TemplateType: number;
        EnableVersioning: boolean;
        EnableMinorVersions: boolean;
        EnableModeration: boolean;
        EnableFolderCreation: boolean;
        EnableAttachments: boolean;
        RemoveExistingContentTypes: boolean;
        RemoveExistingViews: boolean;
        NoCrawl: boolean;
        DefaultDisplayFormUrl: string;
        DefaultEditFormUrl: string;
        DefaultNewFormUrl: string;
        DraftVersionVisibility: string;
        ImageUrl: string;
        Hidden: boolean;
        ForceCheckout: boolean;
        ContentTypeBindings: Array<IContentTypeBinding>;
        FieldRefs: Array<IListInstanceFieldRef>;
        Fields: Array<IField>;
        Folders: Array<IFolder>;
        Views: Array<IView>;
        DataRows: Array<Object>;
        Security: ISecurity;
    }
}
declare module "sharepoint/provisioning/objecthandlers/objectlists" {
    import { ObjectHandlerBase } from "sharepoint/provisioning/objecthandlers/objecthandlerbase";
    import { IListInstance } from "sharepoint/provisioning/schema/IListInstance";
    /**
     * Describes the Lists Object Handler
     */
    export class ObjectLists extends ObjectHandlerBase {
        /**
         * Creates a new instance of the ObjectLists class
         */
        constructor();
        /**
         * Provision Lists
         *
         * @param objects The lists to provision
         */
        ProvisionObjects(objects: Array<IListInstance>): Promise<{}>;
        private EnsureLocationBasedMetadataDefaultsReceiver(clientContext, list);
        private CreateFolders(params);
        private ApplyContentTypeBindings(params);
        private ApplyListInstanceFieldRefs(params);
        private ApplyFields(params);
        private ApplyLookupFields(params);
        private GetFieldXmlAttr(fieldXml, attr);
        private GetFieldXml(field, lists, list);
        private ApplyListSecurity(params);
        private CreateViews(params);
        private InsertDataRows(params);
    }
}
declare module "sharepoint/provisioning/schema/ISiteSchema" {
    import { IListInstance } from "sharepoint/provisioning/schema/IListInstance";
    import { ICustomAction } from "sharepoint/provisioning/schema/ICustomAction";
    import { IFeature } from "sharepoint/provisioning/schema/IFeature";
    import { IFile } from "sharepoint/provisioning/schema/IFile";
    import { INavigation } from "sharepoint/provisioning/schema/inavigation";
    import { IComposedLook } from "sharepoint/provisioning/schema/IComposedLook";
    import { IWebSettings } from "sharepoint/provisioning/schema/IWebSettings";
    export interface SiteSchema {
        Lists: Array<IListInstance>;
        Files: Array<IFile>;
        Navigation: INavigation;
        CustomActions: Array<ICustomAction>;
        ComposedLook: IComposedLook;
        PropertyBagEntries: Object;
        Parameters: Object;
        WebSettings: IWebSettings;
        Features: Array<IFeature>;
    }
}
declare module "sharepoint/provisioning/provisioning" {
    /**
     * Root class of Provisioning
     */
    export class Provisioning {
        private handlers;
        private httpClient;
        private startTime;
        private queueItems;
        /**
         * Creates a new instance of the Provisioning class
         */
        constructor();
        /**
         * Applies a JSON template to the current web
         *
         * @param path URL to the template file
         */
        applyTemplate(path: string): Promise<any>;
        /**
         * Starts the provisioning
         *
         * @param json The parsed template in JSON format
         * @param queue Array of Object Handlers to run
         */
        private start(json, queue);
    }
}
declare module "sharepoint/provisioning/schema/IContentType" {
    export interface IContentType {
        Name: string;
    }
}

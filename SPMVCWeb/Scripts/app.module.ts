/// <reference path="typings/jquery/jquery.d.ts" />
/// <reference path="typings/sharepoint/SharePoint.d.ts" />
/// <reference path="typings/sharepoint/pnp.d.ts" />
/// <reference path="typings/camljs/index.d.ts" />

import * as $pnp from "pnp";

export interface IFilter {
    field: string;
    value: Array<string> | Array<number> | Array<Date> | string | number | boolean;
    operation: Caml.FilterOperation;
    lookupId: boolean;
}

export interface IQueryStringFilter extends IFilter {
    data: string;
}

export interface ICamlFilter extends IFilter {
    fieldType: SP.FieldType;
}


export module Caml {
    "use strict";

    export enum FilterOperation {
        Eq,
        Neq,
        Gt,
        Lt,
        Geq,
        Leq,
        BeginsWith,
        Contains,
        In
    }

    export enum FilterClause {
        None,
        And,
        Or
    }

    export class Builder {

        private _camlBuilder: CamlBuilder.IFieldExpression;
        private _expression: CamlBuilder.IExpression;
        private _viewXml;
        private _clause: FilterClause;
        private _lastClause: FilterClause;

        constructor(viewXml: string, clause?: FilterClause) {
            if (viewXml) {
                this._viewXml = viewXml;
                this._clause = clause;
                switch (clause) {
                    case FilterClause.And:
                        this._camlBuilder = CamlBuilder.FromXml(viewXml).ModifyWhere().AppendAnd();
                        this._lastClause = FilterClause.And;
                        break;
                    case FilterClause.Or:
                        this._camlBuilder = CamlBuilder.FromXml(viewXml).ModifyWhere().AppendOr();
                        this._lastClause = FilterClause.Or;
                        break;
                    default:
                        this._camlBuilder = CamlBuilder.FromXml(viewXml).ReplaceWhere();
                        this._lastClause = FilterClause.None;
                        break;
                }
            }
            else {
                this._camlBuilder = new CamlBuilder().View().Query().Where();
            }
        }

        private getFilterLookupValueExpression(field: string,
            value: string | boolean | number,
            operation: FilterOperation,
            fieldType: SP.FieldType) {
            var fieldExpression = CamlBuilder.Expression();
            var expression: CamlBuilder.IExpression;
            var numberFieldExpression: CamlBuilder.INumberFieldExpression;
            switch (fieldType) {
                case SP.FieldType.lookup:
                default:
                    numberFieldExpression = fieldExpression.LookupField(field).Id();
                    break;
                case SP.FieldType.user:
                    numberFieldExpression = fieldExpression.UserField(field).Id();
                    break;
            }
            switch (operation) {
                case FilterOperation.Gt:
                    expression = numberFieldExpression.GreaterThan(<number>value);
                    break;
                case FilterOperation.Geq:
                    expression = numberFieldExpression.GreaterThanOrEqualTo(<number>value);
                    break;
                case FilterOperation.Neq:
                    if ($pnp.util.stringIsNullOrEmpty(<string>value)) {
                        expression = numberFieldExpression.IsNotNull();
                    }
                    else {
                        expression = numberFieldExpression.NotEqualTo(<number>value);
                    }
                    break;
                case FilterOperation.Eq:
                default:
                    if ($pnp.util.stringIsNullOrEmpty(<string>value)) {
                        expression = numberFieldExpression.IsNull();
                    } else {
                        expression = numberFieldExpression.EqualTo(<number>value);
                    }
                    break;
                case FilterOperation.Lt:
                    expression = numberFieldExpression.LessThan(<number>value);
                    break;
                case FilterOperation.Leq:
                    expression = numberFieldExpression.LessThanOrEqualTo(<number>value);
                    break;
            }
            return expression;
        }

        private getFilterValueExpression(field: string, value: string | boolean | number, operation: FilterOperation, isLookupId: boolean, fieldType: SP.FieldType) {
            var self = this;
            var expression: CamlBuilder.IExpression;
            if (Boolean(isLookupId)) {
                expression = self.getFilterLookupValueExpression(field, value, operation, fieldType);
            } else {
                var fieldExpression = CamlBuilder.Expression();
                switch (fieldType) {
                    case SP.FieldType.text:
                    default:
                        var textFieldExpression = fieldExpression.TextField(field);
                        switch (operation) {
                            case FilterOperation.BeginsWith:
                                expression = textFieldExpression.BeginsWith(<string>value);
                                break;
                            case FilterOperation.Contains:
                                expression = textFieldExpression.Contains(<string>value);
                                break;
                            case FilterOperation.Neq:
                                if ($pnp.util.stringIsNullOrEmpty(<string>value)) {
                                    expression = textFieldExpression.IsNotNull();
                                }
                                else {
                                    expression = textFieldExpression.NotEqualTo(<string>value);
                                }
                                break;
                            case FilterOperation.Eq:
                            default:
                                if ($pnp.util.stringIsNullOrEmpty(<string>value)) {
                                    expression = textFieldExpression.IsNull();
                                } else {
                                    expression = textFieldExpression.EqualTo(<string>value);
                                }
                                break;
                        }
                        break;
                    case SP.FieldType.lookup:
                        var lookupFieldExpression = fieldExpression.LookupField(field);
                        switch (operation) {
                            case FilterOperation.BeginsWith:
                                expression = lookupFieldExpression.ValueAsText().BeginsWith(<string>value);
                                break;
                            case FilterOperation.Contains:
                                expression = lookupFieldExpression.ValueAsText().Contains(<string>value);
                                break;
                            case FilterOperation.Neq:
                                if ($pnp.util.stringIsNullOrEmpty(<string>value)) {
                                    expression = lookupFieldExpression.ValueAsText().IsNotNull();
                                }
                                else {
                                    expression = lookupFieldExpression.ValueAsText().NotEqualTo(<string>value);
                                }
                                break;
                            case FilterOperation.Eq:
                            default:
                                if ($pnp.util.stringIsNullOrEmpty(<string>value)) {
                                    expression = lookupFieldExpression.ValueAsText().IsNull();
                                } else {
                                    expression = lookupFieldExpression.ValueAsText()
                                        .EqualTo(<string>value);
                                }
                                break;
                        }
                        break;
                    case SP.FieldType.dateTime:
                        var dateFieldExpression = fieldExpression.DateField(field);
                        switch (operation) {
                            case FilterOperation.Gt:
                                expression = dateFieldExpression.GreaterThan(<string>value);
                                break;
                            case FilterOperation.Geq:
                                expression = dateFieldExpression.GreaterThanOrEqualTo(<string>value);
                                break;
                            case FilterOperation.Neq:
                                if ($pnp.util.stringIsNullOrEmpty(<string>value)) {
                                    expression = dateFieldExpression.IsNotNull();
                                }
                                else {
                                    expression = dateFieldExpression.NotEqualTo(<string>value);
                                }
                                break;
                            case FilterOperation.Eq:
                            default:
                                if ($pnp.util.stringIsNullOrEmpty(<string>value)) {
                                    expression = dateFieldExpression.IsNull();
                                } else {
                                    expression = dateFieldExpression.EqualTo(<string>value);
                                }
                                break;
                            case FilterOperation.Lt:
                                expression = dateFieldExpression.LessThan(<string>value);
                                break;
                            case FilterOperation.Leq:
                                expression = dateFieldExpression.LessThanOrEqualTo(<string>value);
                                break;
                        }
                        break;
                    case <SP.FieldType>101: //datetime 
                        var dateTimeFieldExpression = fieldExpression.DateTimeField(field);
                        switch (operation) {
                            case FilterOperation.Gt:
                                expression = dateTimeFieldExpression.GreaterThan(<string>value);
                                break;
                            case FilterOperation.Geq:
                                expression = dateTimeFieldExpression
                                    .GreaterThanOrEqualTo(<string>value);
                                break;
                            case FilterOperation.Neq:
                                if ($pnp.util.stringIsNullOrEmpty(<string>value)) {
                                    expression = dateTimeFieldExpression.IsNotNull();
                                }
                                else {
                                    expression = dateTimeFieldExpression.NotEqualTo(<string>value);
                                }
                                break;
                            case FilterOperation.Eq:
                            default:
                                if ($pnp.util.stringIsNullOrEmpty(<string>value)) {
                                    expression = dateTimeFieldExpression.IsNull();
                                } else {
                                    expression = dateTimeFieldExpression.EqualTo(<string>value);
                                }
                                break;
                            case FilterOperation.Lt:
                                expression = dateTimeFieldExpression.LessThan(<string>value);
                                break;
                            case FilterOperation.Leq:
                                expression = dateTimeFieldExpression.LessThanOrEqualTo(<string>value);
                                break;
                        }
                        break;
                    case SP.FieldType.counter:
                        var counterFieldExpression = fieldExpression.CounterField(field);
                        switch (operation) {
                            case FilterOperation.Gt:
                                expression = counterFieldExpression.GreaterThan(<number>value);
                                break;
                            case FilterOperation.Geq:
                                expression = counterFieldExpression
                                    .GreaterThanOrEqualTo(<number>value);
                                break;
                            case FilterOperation.Neq:
                                if ($pnp.util.stringIsNullOrEmpty(<string>value)) {
                                    expression = counterFieldExpression.IsNotNull();
                                }
                                else {
                                    expression = counterFieldExpression.NotEqualTo(<number>value);
                                }
                                break;
                            case FilterOperation.Eq:
                            default:
                                if ($pnp.util.stringIsNullOrEmpty(<string>value)) {
                                    expression = counterFieldExpression.IsNull();
                                } else {
                                    expression = counterFieldExpression.EqualTo(<number>value);
                                }
                                break;
                            case FilterOperation.Lt:
                                expression = counterFieldExpression.LessThan(<number>value);
                                break;
                            case FilterOperation.Leq:
                                expression = counterFieldExpression.LessThanOrEqualTo(<number>value);
                                break;
                        }
                        break;
                    case SP.FieldType.integer:
                        var integerFieldExpression = fieldExpression.IntegerField(field);
                        switch (operation) {
                            case FilterOperation.Gt:
                                expression = integerFieldExpression.GreaterThan(<number>value);
                                break;
                            case FilterOperation.Geq:
                                expression = integerFieldExpression
                                    .GreaterThanOrEqualTo(<number>value);
                                break;
                            case FilterOperation.Neq:
                                if ($pnp.util.stringIsNullOrEmpty(<string>value)) {
                                    expression = integerFieldExpression.IsNotNull();
                                }
                                else {
                                    expression = integerFieldExpression.NotEqualTo(<number>value);
                                }
                                break;
                            case FilterOperation.Eq:
                            default:
                                if ($pnp.util.stringIsNullOrEmpty(<string>value)) {
                                    expression = integerFieldExpression.IsNull();
                                } else {
                                    expression = integerFieldExpression.EqualTo(<number>value);
                                }
                                break;
                            case FilterOperation.Lt:
                                expression = integerFieldExpression.LessThan(<number>value);
                                break;
                            case FilterOperation.Leq:
                                expression = integerFieldExpression.LessThanOrEqualTo(<number>value);
                                break;
                        }
                        break;
                    case SP.FieldType.modStat:
                        var modStatFieldExpression = fieldExpression.ModStatField(field).ValueAsText();
                        switch (operation) {
                            case FilterOperation.BeginsWith:
                                expression = modStatFieldExpression.BeginsWith(<string>value);
                                break;
                            case FilterOperation.Contains:
                                expression = modStatFieldExpression.Contains(<string>value);
                                break;
                            case FilterOperation.Neq:
                                if ($pnp.util.stringIsNullOrEmpty(<string>value)) {
                                    expression = modStatFieldExpression.IsNotNull();
                                }
                                else {
                                    expression = modStatFieldExpression.NotEqualTo(<string>value);
                                }
                                break;
                            case FilterOperation.Eq:
                            default:
                                if ($pnp.util.stringIsNullOrEmpty(<string>value)) {
                                    expression = modStatFieldExpression.IsNull();
                                } else {
                                    expression = modStatFieldExpression.EqualTo(<string>value);
                                }
                                break;
                        }
                        break;
                    case SP.FieldType.number:
                        var numberFieldExpression = fieldExpression.NumberField(field);
                        switch (operation) {
                            case FilterOperation.Gt:
                                expression = numberFieldExpression.GreaterThan(<number>value);
                                break;
                            case FilterOperation.Geq:
                                expression = numberFieldExpression.GreaterThanOrEqualTo(<number>value);
                                break;
                            case FilterOperation.Neq:
                                if ($pnp.util.stringIsNullOrEmpty(<string>value)) {
                                    expression = numberFieldExpression.IsNotNull();
                                }
                                else {
                                    expression = numberFieldExpression.NotEqualTo(<number>value);
                                }
                                break;
                            case FilterOperation.Eq:
                            default:
                                if ($pnp.util.stringIsNullOrEmpty(<string>value)) {
                                    expression = numberFieldExpression.IsNull();
                                } else {
                                    expression = numberFieldExpression.EqualTo(<number>value);
                                }
                                break;
                            case FilterOperation.Lt:
                                expression = numberFieldExpression.LessThan(<number>value);
                                break;
                            case FilterOperation.Leq:
                                expression = numberFieldExpression.LessThanOrEqualTo(<number>value);
                                break;
                        }
                        break;
                    case SP.FieldType.URL:
                        var urlFieldExpression = fieldExpression.UrlField(field);
                        switch (operation) {
                            case FilterOperation.BeginsWith:
                                expression = urlFieldExpression.BeginsWith(<string>value);
                                break;
                            case FilterOperation.Contains:
                                expression = urlFieldExpression.Contains(<string>value);
                                break;
                            case FilterOperation.Neq:
                                if ($pnp.util.stringIsNullOrEmpty(<string>value)) {
                                    expression = urlFieldExpression.IsNotNull();
                                }
                                else {
                                    expression = urlFieldExpression.NotEqualTo(<string>value);
                                }
                                break;
                            case FilterOperation.Eq:
                            default:
                                if ($pnp.util.stringIsNullOrEmpty(<string>value)) {
                                    expression = urlFieldExpression.IsNull();
                                } else {
                                    expression = urlFieldExpression.EqualTo(<string>value);
                                }
                                break;
                        }
                        break;
                    case SP.FieldType.user:
                        var userFieldExpression = fieldExpression.UserField(field).ValueAsText();
                        switch (operation) {
                            case FilterOperation.BeginsWith:
                                expression = userFieldExpression.BeginsWith(<string>value);
                                break;
                            case FilterOperation.Contains:
                                expression = userFieldExpression.Contains(<string>value);
                                break;
                            case FilterOperation.Neq:
                                if ($pnp.util.stringIsNullOrEmpty(<string>value)) {
                                    expression = userFieldExpression.IsNotNull();
                                }
                                else {
                                    expression = userFieldExpression.NotEqualTo(<string>value);
                                }
                                break;
                            case FilterOperation.Eq:
                            default:
                                if ($pnp.util.stringIsNullOrEmpty(<string>value)) {
                                    expression = userFieldExpression.IsNull();
                                } else {
                                    expression = userFieldExpression.EqualTo(<string>value);
                                }
                                break;
                        }
                        break;
                    case SP.FieldType.boolean:
                        var booleanFieldExpression = fieldExpression.BooleanField(field);
                        if ($pnp.util.stringIsNullOrEmpty(<string>value)) {
                            if (operation == FilterOperation.Neq) {
                                expression = booleanFieldExpression.IsNotNull();
                            }
                            else {
                                expression = booleanFieldExpression.IsNull();
                            }
                        } else {
                            if (Boolean(value) === true) {
                                expression = booleanFieldExpression.IsTrue();
                            } else {
                                expression = booleanFieldExpression.IsFalse();
                            }
                        }
                        break;
                }
            }
            return expression;
        }

        private getFilterMultiValueExpression(filter: ICamlFilter) {
            var self = this;
            var fieldExpression = CamlBuilder.Expression();
            var expression: CamlBuilder.IExpression;
            if ($pnp.util.isArray(filter.value)) {
                if ((<any[]>filter.value).length === 1) {
                    expression = self.getFilterValueExpression(filter.field,
                        filter.value[0],
                        filter.operation,
                        filter.lookupId,
                        filter.fieldType);
                } else if ((<any[]>filter.value).length > 1) {
                    if (Boolean(filter.lookupId)) {
                        switch (filter.fieldType) {
                            case SP.FieldType.lookup:
                            default:
                                expression = fieldExpression.LookupMultiField(filter.field).IncludesSuchItemThat().Id()
                                    .In(<number[]>filter.value);
                                break;
                            case SP.FieldType.user:
                                expression = fieldExpression.UserMultiField(filter.field).IncludesSuchItemThat().Id()
                                    .In(<number[]>filter.value);
                                break;
                        }
                    } else {
                        switch (filter.fieldType) {
                            case SP.FieldType.lookup:
                                expression = fieldExpression.LookupMultiField(filter.field).IncludesSuchItemThat()
                                    .ValueAsText().In(<string[]>filter.value);
                                break;
                            case SP.FieldType.text:
                            default:
                                expression = fieldExpression.TextField(filter.field).In(<string[]>filter.value);
                                break;
                            case SP.FieldType.dateTime:
                                expression = fieldExpression.DateField(filter.field).In(<Date[]>filter.value);
                                break;
                            case <SP.FieldType>101: //datetime                                        
                                expression = fieldExpression.DateTimeField(filter.field).In(<Date[]>filter.value);
                                break;
                            case SP.FieldType.counter:
                                expression = fieldExpression.CounterField(filter.field).In(<number[]>filter.value);
                                break;
                            case SP.FieldType.integer:
                                expression = fieldExpression.IntegerField(filter.field).In(<number[]>filter.value);
                                break;
                            case SP.FieldType.modStat:
                                expression = fieldExpression.ModStatField(filter.field).ValueAsText().In(<string[]>filter.value);
                                break;
                            case SP.FieldType.number:
                                expression = fieldExpression.NumberField(filter.field).In(<number[]>filter.value);
                                break;
                            case SP.FieldType.URL:
                                expression = fieldExpression.UrlField(filter.field).In(<string[]>filter.value);
                                break;
                            case SP.FieldType.user:
                                expression = fieldExpression.UserMultiField(filter.field).IncludesSuchItemThat()
                                    .ValueAsText().In(<string[]>filter.value);
                                break;
                        }
                    }
                }
            } else {
                expression = self.getFilterValueExpression(filter.field,
                    <any>filter.value,
                    filter.operation,
                    filter.lookupId,
                    filter.fieldType);
            }
            return expression;
        }

        private getExpressions(filters: Array<ICamlFilter>) {
            var expressions = new Array<CamlBuilder.IExpression>();
            if (filters) {
                for (var i in filters) {
                    if (filters.hasOwnProperty(i)) {
                        var filter = filters[i];
                        if (filter) {
                            var expression = this.getFilterMultiValueExpression(filter);
                            if (expression) {
                                expressions.push(expression);
                            }
                        }
                    }
                }
            }
            return expressions;
        }

        public Append(clause: FilterClause, ...filters: Array<ICamlFilter>) {
            if (!clause) {
                clause = this._lastClause;
            }
            switch (clause) {
                case FilterClause.None:
                default:
                case FilterClause.And:
                    this.AppendAnd.apply(this, [FilterClause.And].concat(<any>filters));
                    break;
                case FilterClause.Or:
                    this.AppendOr.apply(this, [FilterClause.Or].concat(<any>filters));
                    break;
            }

            return this;
        }

        public AppendAnd(clause?: FilterClause, ...filters: Array<ICamlFilter>) {
            if (!clause) {
                clause = this._lastClause;
            }
            var expressions = this.getExpressions(filters);
            if (expressions.length > 0) {
                this._lastClause = FilterClause.And;
                this._expression = this._expression ? (clause == FilterClause.Or ? this._expression.And().Any(expressions) : this._expression.And().All(expressions))
                    : (clause == FilterClause.Or ? this._camlBuilder.Any(expressions) : this._camlBuilder.All(expressions));
            }
            return this;
        }

        public AppendOr(clause?: FilterClause, ...filters: Array<ICamlFilter>) {
            if (!clause) {
                clause = this._lastClause;
            }
            var expressions = this.getExpressions(filters);
            if (expressions.length > 0) {
                this._lastClause = FilterClause.Or;
                this._expression = this._expression ? (clause == FilterClause.And ? this._expression.Or().All(expressions) : this._expression.Or().Any(expressions))
                    : (clause == FilterClause.And ? this._camlBuilder.All(expressions) : this._camlBuilder.Any(expressions));
            }
            return this;
        }

        public Clear() {
            this._expression = null;
            this._lastClause = this._clause;
        }

        public ToString() {
            if (this._expression) {
                this._viewXml = this._expression.ToString();
            }
            return this._viewXml;
        }
    }
}

window["Caml"] = Caml;

export namespace App {
    "use strict";
    export class AppBase {

        public appWebUrl: string;
        public hostWebUrl: string;
        public scriptBase: string;
        public $: JQueryStatic;

        private _initialized: boolean;
        private _scriptPromises;

        constructor() {
            this._initialized = false;
            this._scriptPromises = {};
        }

        public init(preloadedScripts: any[]) {
            var self = this;
            if (preloadedScripts) {
                let $ = preloadedScripts["jquery"];
                self.$ = $;
                if (self.$) {
                    (<any>self.$).cachedScript = (url, options) => {
                        options = self.$.extend(options || {}, {
                            dataType: "script",
                            cache: true,
                            url: url
                        });
                        return self.$.ajax(options);
                    };
                } else {
                    throw "jQuery is not loaded!";
                }
            }
            this.hostWebUrl = (<any>window)._spPageContextInfo && !$pnp.util.stringIsNullOrEmpty((<any>window)._spPageContextInfo.webAbsoluteUrl) ? (<any>window)._spPageContextInfo.webAbsoluteUrl : $pnp.util.getUrlParamByName("SPHostUrl");
            if ($pnp.util.stringIsNullOrEmpty(this.hostWebUrl)) {
                throw "SPHostUrl url parameter must be specified!";
            }
            this.appWebUrl = (<any>window)._spPageContextInfo && !$pnp.util.stringIsNullOrEmpty((<any>window)._spPageContextInfo.appWebUrl) ? (<any>window)._spPageContextInfo.appWebUrl : $pnp.util.getUrlParamByName("SPAppWebUrl");
            if ($pnp.util.stringIsNullOrEmpty(this.appWebUrl)) {
                throw "SPAppWebUrl url parameter must be specified!";
            }
            this.scriptBase = $pnp.util.combinePaths(this.hostWebUrl, (<any>window)._spPageContextInfo && !$pnp.util.stringIsNullOrEmpty((<any>window)._spPageContextInfo.layoutsUrl) ? (<any>window)._spPageContextInfo.layoutsUrl : "_layouts/15");
            this._initialized = true;
            self.$(self).trigger("app-init");
        }

        public ensureScript(url): JQueryXHR {
            if (url) {
                url = url.toLowerCase().replace("~sphost", this.hostWebUrl)
                    .replace("~spapp", this.appWebUrl)
                    .replace("~splayouts", this.scriptBase);
                var scriptPromise = this._scriptPromises[url];
                if (!scriptPromise) {
                    scriptPromise = (<any>this.$).cachedScript(url);
                    this._scriptPromises[url] = scriptPromise;
                }
                return scriptPromise;
            }
            return null;
        }

        public delay = (() => {
            var timer = 0;
            return (callback: () => void, ms: number) => {
                clearTimeout(timer);
                timer = setTimeout(callback, ms);
            };
        })();

        public render(modules: IModule[]) {
            var self = this;
            if (!self._initialized) {
                throw "App is not initialized!";
            }
            self.ensureScript("~splayouts/MicrosoftAjax.js").then(() => {
                self.ensureScript("~splayouts/SP.Runtime.js").then(() => {
                    self.ensureScript("~splayouts/SP.RequestExecutor.js").then(() => {
                        self.ensureScript("~splayouts/SP.js").then(() => {
                            if ($pnp.util.isArray(modules)) {
                                self.$.each(modules, (i: number, module: IModule) => {
                                    module.render();
                                });
                            }
                            self.$(self).trigger("app-render");
                        });
                    });
                });
            });
        }

        public is_initialized(): boolean {
            var self = this;
            return self._initialized;
        }

        public get_hostWebUrl(): string {
            var self = this;
            return self.hostWebUrl;
        }

        public get_appWebUrl(): string {
            var self = this;
            return self.appWebUrl;
        }

        public get_BasePermissions(permMask: string): SP.BasePermissions {
            var permissions = new SP.BasePermissions();
            if (permMask) {
                var permMaskHigh = permMask.length <= 10 ? 0 : parseInt(permMask.substring(2, permMask.length - 8), 16);
                var permMaskLow = permMask.length <= 10 ? parseInt(permMask) : parseInt(permMask.substring(permMask.length - 8, permMask.length), 16);
                permissions.initPropertiesFromJson({ "High": permMaskHigh, "Low": permMaskLow });
            }
            return permissions;
        }

        public getQueryParam(url: string, name: string) {
            var self = this;
            if (url) {
                var search = null;
                if (URL) {
                    search = new URL(url).search;
                }
                else {
                    var a = document.createElement('a');
                    a.href = url;
                    search = a.search;
                }
                if (search) {
                    while (search.startsWith("?") || search.startsWith("&")) {
                        search = search.slice(1, search.length);
                    }
                    var qParameters = <Array<string>>search.split("&");
                    for (var i in qParameters) {
                        if (qParameters.hasOwnProperty(i)) {
                            var qParameter = qParameters[i].split("=");
                            var key = decodeURIComponent(<any>self.$(qParameter).get(0));
                            if (key && key.toUpperCase() === name.toUpperCase()) {
                                var value = decodeURIComponent(<any>self.$(qParameter).get(1));
                                return value;
                            }
                        }
                    }
                }
            }
            return "";
        }
    }

    export interface ISPService {
        getFormDigest();
    }

    export interface IModuleOptions {
        controllerName: string;
    }

    export interface IModule {
        get_app(): AppBase;
        get_options(): IModuleOptions;
        render();
    }

    export module Module {

        export interface IListViewOptions extends IModuleOptions {
            listTitle: string;
            listId?: string;
            listUrl?: string;
            viewId?: string;
            viewXml?: string,
            orderBy?: string;
            sortAsc?: boolean;
            //filter?: string;
            limit?: number;
            expands?: string[];
            paged?: boolean,
            rootFolder?: string,
            //fields?: string[];
            appendRows?: boolean;
            renderMethod?: RenderMethod;
            renderOptions?: number;
            queryStringFilters?: Array<IQueryStringFilter>;
            queryBuilder?: Caml.Builder;
        }

        export enum RenderMethod {
            Default,
            RenderListDataAsStream,
            RenderListFilterData,
            RenderListData,
            GetItems
        }

        export class ListViewBase implements IModule {
            private _options: IListViewOptions;
            private _app: AppBase;

            constructor(app: AppBase, options: IListViewOptions) {
                if (!app) {
                    throw "App must be specified for ListView!";
                }
                this._app = app;
                this._options = $pnp.util.extend(<IListViewOptions>{
                    controllerName: null,
                    listTitle: null,
                    listId: null,
                    listUrl: null,
                    viewId: null,
                    viewXml: null,
                    orderBy: null,
                    sortAsc: null,
                    limit: null,
                    paged: null,
                    rootFolder: null,
                    appendRows: null,
                    renderMethod: null,
                    renderOptions: null,
                    queryStringFilters: null,
                    queryBuilder: new Caml.Builder(options.viewXml, Caml.FilterClause.And)
                }, options);
            }

            private addToken(query: $pnp.Dictionary<string>, token: string) {
                var self = this;
                if (token) {
                    while ((<any>token).startsWith("?") || (<any>token).startsWith("&")) {
                        token = token.slice(1, token.length);
                    }
                    var qParameters = <Array<string>>token.split("&");
                    for (var i in qParameters) {
                        if (qParameters.hasOwnProperty(i)) {
                            var qParameter = qParameters[i].split("=");
                            var key = self._app.$(qParameter).get(0);
                            var value = self._app.$(qParameter).get(1);
                            if (key && value) {
                                query.add(String(key), encodeURIComponent(String(value)));
                            }
                        }
                    }
                }
            }

            private addFilterQuery(query: $pnp.Dictionary<string>) {
                var self = this;
                if (self._app.$.isArray(self._options.queryStringFilters)) {
                    for (var i = 0; i < self._options.queryStringFilters.length; i++) {
                        var filter = self._options.queryStringFilters[i];
                        if (filter) {
                            if (self._app.$.isArray(filter.value) && (<any[]>filter.value).length > 1) {
                                query.add("FilterFields" + (i + 1), encodeURIComponent(filter.field));
                                query.add("FilterValues" + (i + 1),
                                    encodeURIComponent((<Array<string>>filter.value).join(";#")));
                            } else {
                                query.add("FilterField" + (i + 1), encodeURIComponent(filter.field));
                                query.add("FilterValue" + (i + 1), encodeURIComponent(<string>filter.value));
                            }
                            if (Boolean(filter.lookupId)) {
                                query.add("FilterLookupId" + (i + 1), String(1));
                            }
                            if (filter.operation != Caml.FilterOperation.Eq) {
                                query.add("FilterOp" + (i + 1), encodeURIComponent(Caml.FilterOperation[filter.operation]));
                            }
                            if (filter.data) {
                                query.add("FilterData" + (i + 1), encodeURIComponent(String(filter.data)));
                            }
                        }
                    }
                }
            }

            private getList() {
                var self = this;
                if (!$pnp.util.stringIsNullOrEmpty(self._options.listId)) {
                    return $pnp.sp.crossDomainWeb(self._app.appWebUrl, self._app.hostWebUrl).lists.getById(self._options.listId);
                } else if (!$pnp.util.stringIsNullOrEmpty(self._options.listUrl)) {
                    return $pnp.sp.crossDomainWeb(self._app.appWebUrl, self._app.hostWebUrl).getList(self._options.listUrl);
                } else if (!$pnp.util.stringIsNullOrEmpty(self._options.listTitle)) {
                    return $pnp.sp.crossDomainWeb(self._app.appWebUrl, self._app.hostWebUrl).lists.getByTitle(self._options.listId);
                }
                throw "List is not specified.";
            }

            private renderListDataAsStream(token: string/*, prevItemId?: number, pageLastRow?: number*/) {
                var self = this;
                var deferred = self._app.$.Deferred();
                try {
                    var query = self.getList();
                    query.concat("/RenderListDataAsStream");
                    if (!$pnp.util.stringIsNullOrEmpty(token)) {
                        self.addToken(query.query, token);
                    } else {
                        if (!$pnp.util.stringIsNullOrEmpty(self._options.viewId)) {
                            query.query.add("View", encodeURIComponent(self._options.viewId));
                        }
                        if (!$pnp.util.stringIsNullOrEmpty(self._options.orderBy)) {
                            query.query.add("SortField", encodeURIComponent(self._options.orderBy));
                        }
                        if (!$pnp.util.stringIsNullOrEmpty(<any>self._options.sortAsc)) {
                            query.query.add("SortDir", self._options.sortAsc ? "Asc" : "Desc");
                        }
                    }
                    self.addFilterQuery(query.query);
                    var viewXml = self._options.viewXml;
                    if (self._options.queryBuilder) {
                        viewXml = self._options.queryBuilder.ToString();
                    }
                    //if (!$pnp.util.stringIsNullOrEmpty(<any>prevItemId)) {
                    //    query.query.add("p_ID", prevItemId);
                    //}
                    //if (!$pnp.util.stringIsNullOrEmpty(<any>pageLastRow)) {
                    //    query.query.add("PageLastRow", pageLastRow);
                    //}
                    var url = query.toUrlAndQuery();
                    var parameters = <any>{ "__metadata": { "type": "SP.RenderListDataParameters" }, "RenderOptions": self._options.renderOptions };
                    if (!$pnp.util.stringIsNullOrEmpty(viewXml)) {
                        parameters.ViewXml = viewXml;
                    }
                    if (!$pnp.util.stringIsNullOrEmpty(<any>self._options.paged)) {
                        //parameters.Paging = self._options.paged == true ? "TRUE" : undefined;
                    }
                    if (!$pnp.util.stringIsNullOrEmpty(self._options.rootFolder)) {
                        parameters.FolderServerRelativeUrl = self._options.rootFolder;
                    }
                    var postBody = JSON.stringify({ "parameters": parameters });
                    var executor = new SP.RequestExecutor(self._app.appWebUrl);
                    executor.executeAsync(<SP.RequestInfo>{
                        url: url,
                        method: "POST",
                        headers: {
                            "accept": "application/json;odata=verbose",
                            "content-Type": "application/json;odata=verbose"
                        },
                        body: postBody,
                        success: (data) => {
                            var result = null;
                            if (data.body) {
                                result = JSON.parse(<string>data.body);
                            }
                            deferred.resolve(result);
                        },
                        error: (data, errorCode, errorMessage) => {
                            if (data.body) {
                                try {
                                    var error = JSON.parse(<string>data.body);
                                    if (error && error.error) {
                                        errorMessage = error.error.message.value;
                                    }
                                }
                                catch (e) { }
                            }
                            self._app.$(self._app).trigger("app-error", [errorMessage]);
                            deferred.reject(data, errorCode, errorMessage);
                        }
                    });
                } catch (e) {
                    self._app.$(self._app).trigger("app-error", [e]);
                    deferred.reject(e);
                }
                return deferred.promise();
            }

            private renderListData(token: string) {
                var self = this;
                var deferred = self._app.$.Deferred();
                try {
                    var query = self.getList();
                    query.concat("/renderListData(@viewXml)");
                    query.query.add("@viewXml", "'" + encodeURIComponent(self._options.viewXml) + "'");
                    if (!$pnp.util.stringIsNullOrEmpty(token)) {
                        self.addToken(query.query, token);
                    } else {
                        if (!$pnp.util.stringIsNullOrEmpty(self._options.viewId)) {
                            query.query.add("View", encodeURIComponent(self._options.viewId));
                        }
                        if (!$pnp.util.stringIsNullOrEmpty(self._options.orderBy)) {
                            query.query.add("SortField", encodeURIComponent(self._options.orderBy));
                        }
                        if (!$pnp.util.stringIsNullOrEmpty(<any>self._options.sortAsc)) {
                            query.query.add("SortDir", self._options.sortAsc ? "Asc" : "Desc");
                        }
                    }
                    self.addFilterQuery(query.query);
                    var url = query.toUrlAndQuery();
                    var executor = new SP.RequestExecutor(self._app.appWebUrl);
                    executor.executeAsync(<SP.RequestInfo>{
                        url: url,
                        method: "POST",
                        headers: {
                            "accept": "application/json;odata=verbose",
                            "content-Type": "application/json;odata=verbose"
                        },
                        success: (data) => {
                            var result = null;
                            if (data.body) {
                                result = JSON.parse(JSON.parse(<string>data.body).d.RenderListData);
                            }
                            deferred.resolve({ ListData: result });
                        },
                        error: (data, errorCode, errorMessage) => {
                            if (data.body) {
                                try {
                                    var error = JSON.parse(<string>data.body);
                                    if (error && error.error) {
                                        errorMessage = error.error.message.value;
                                    }
                                }
                                catch (e) { }
                            }
                            self._app.$(self._app).trigger("app-error", [errorMessage]);
                            deferred.reject(data, errorCode, errorMessage);
                        }
                    });
                } catch (e) {
                    self._app.$(self._app).trigger("app-error", [e]);
                    deferred.reject(e);
                }
                return deferred.promise();
            }

            private getItems() {
                var self = this;
                var deferred = self._app.$.Deferred();
                try {
                    var query = self.getList();
                    query.concat("/GetItems");
                    if ($pnp.util.isArray(self._options.expands)) {
                        query = query.expand(<any>self._options.expands);
                    }
                    var url = query.toUrlAndQuery();
                    var postBody = JSON.stringify({ "query": { "__metadata": { "type": "SP.CamlQuery" }, "ViewXml": self._options.viewXml } });
                    var executor = new SP.RequestExecutor(self._app.appWebUrl);
                    executor.executeAsync(<SP.RequestInfo>{
                        url: url,
                        method: "POST",
                        body: postBody,
                        headers: {
                            "accept": "application/json;odata=verbose",
                            "content-Type": "application/json;odata=verbose"
                        },
                        success: (data) => {
                            var listData = null;
                            if (data.body) {
                                var d = JSON.parse(<string>data.body).d;
                                listData = { Row: d.results, NextHref: null, PrevHref: null };
                            }
                            deferred.resolve({ ListData: listData });
                        },
                        error: (data, errorCode, errorMessage) => {
                            if (data.body) {
                                try {
                                    var error = JSON.parse(<string>data.body);
                                    if (error && error.error) {
                                        errorMessage = error.error.message.value;
                                    }
                                }
                                catch (e) { }
                            }
                            self._app.$(self._app).trigger("app-error", [errorMessage]);
                            deferred.reject(data, errorCode, errorMessage);
                        }
                    });
                } catch (e) {
                    self._app.$(self._app).trigger("app-error", [e]);
                    deferred.reject(e);
                }
                return deferred.promise();
            }

            private getItemsREST(token: string) {
                var self = this;
                var deferred = self._app.$.Deferred();
                try {
                    var query = self.getList().items;
                    if (!$pnp.util.stringIsNullOrEmpty(self._options.orderBy)) {
                        query = query.orderBy(self._options.orderBy, self._options.sortAsc);
                    }
                    //if (!$pnp.util.stringIsNullOrEmpty(self._options.filter)) {
                    //    query = query.filter(self._options.filter);
                    //}
                    if (self._options.limit > 0) {
                        query = query.top(self._options.limit);
                    }
                    if ($pnp.util.isArray(self._options.expands)) {
                        query = query.expand(<any>self._options.expands);
                    }
                    if (!$pnp.util.stringIsNullOrEmpty(token)) {
                        query.query.add("$skiptoken", encodeURIComponent(token));
                    }
                    var url = query.toUrlAndQuery();
                    var executor = new SP.RequestExecutor(self._app.appWebUrl);
                    executor.executeAsync(<SP.RequestInfo>{
                        url: url,
                        method: "GET",
                        headers: {
                            "accept": "application/json;odata=verbose",
                            "content-Type": "application/json;odata=verbose"
                        },
                        success: (data) => {
                            var listData = null;
                            if (data.body) {
                                var d = JSON.parse(<string>data.body).d;
                                listData = { Row: d.results, NextHref: null, PrevHref: null };
                                listData.NextHref = self._app.getQueryParam(d["__next"], "$skiptoken");
                                //listData.PrevHref = self._app.getQueryParam(d["__prev"], "$skiptoken");
                            }
                            deferred.resolve({ ListData: listData });
                        },
                        error: (data, errorCode, errorMessage) => {
                            if (data.body) {
                                try {
                                    var error = JSON.parse(<string>data.body);
                                    if (error && error.error) {
                                        errorMessage = error.error.message.value;
                                    }
                                }
                                catch (e) { }
                            }
                            self._app.$(self._app).trigger("app-error", [errorMessage]);
                            deferred.reject(data, errorCode, errorMessage);
                        }
                    });
                } catch (e) {
                    self._app.$(self._app).trigger("app-error", [e]);
                    deferred.reject(e);
                }
                return deferred.promise();
            }

            private getItemsCSOM(token: string) {
                var self = this;
                var deferred = self._app.$.Deferred();
                try {
                    var context = new SP.ClientContext(self._app.appWebUrl);
                    var factory = new SP.ProxyWebRequestExecutorFactory(self._app.appWebUrl);
                    context.set_webRequestExecutorFactory(factory);
                    var appContextSite = new SP.AppContextSite(context, self._app.hostWebUrl);
                    var web = appContextSite.get_web();
                    var list: SP.List = null;
                    if (!$pnp.util.stringIsNullOrEmpty(self._options.listId)) {
                        list = web.get_lists().getById(self._options.listId);
                    } else if (!$pnp.util.stringIsNullOrEmpty(self._options.listUrl)) {
                        list = web.getList(self._options.listUrl);
                    } else if (!$pnp.util.stringIsNullOrEmpty(self._options.listTitle)) {
                        list = web.get_lists().getByTitle(self._options.listTitle);
                    } else {
                        throw "List is not specified.";
                    }
                    var camlQuery = new SP.CamlQuery();
                    camlQuery.set_viewXml(self._options.viewXml);
                    if (!$pnp.util.stringIsNullOrEmpty(token)) {
                        var position = new SP.ListItemCollectionPosition();
                        position.set_pagingInfo(token);
                        camlQuery.set_listItemCollectionPosition(position);
                    }
                    var items = list.getItems(camlQuery);
                    context.load(items);
                    context.executeQueryAsync(() => {
                        var listData = { Row: self._app.$.map(items.get_data(), (item: SP.ListItem) => { return item.get_fieldValues(); }), NextHref: null, PrevHref: null };
                        var position = items.get_listItemCollectionPosition();
                        if (position) {
                            listData.NextHref = position.get_pagingInfo();
                        }
                        deferred.resolve({ ListData: listData });
                    }, (sender, args) => {
                        self._app.$(self._app).trigger("app-error", [args.get_message(), args.get_stackTrace()]);
                        deferred.reject(sender, args);
                    });
                } catch (e) {
                    self._app.$(self._app).trigger("app-error", [e]);
                    deferred.reject(e);
                }
                return deferred.promise();
            }

            public getListItems(token: string) {
                var self = this;
                switch (self._options.renderMethod) {
                    case RenderMethod.RenderListDataAsStream:
                        return self.renderListDataAsStream(token);
                    case RenderMethod.RenderListData:
                        return self.renderListData(token);
                    case RenderMethod.GetItems:
                        return self.getItems();
                    case RenderMethod.Default:
                        return self.getItemsREST(token);
                    default:
                        return self.getItemsCSOM(token);
                }
            }

            public get_app(): AppBase {
                return this._app;
            }

            public get_options() {
                return this._options;
            }

            public render() {
                this._app.$(self).trigger("model-render");
            }
        }

        export interface IListsViewOptions extends IModuleOptions {
            delay: number;
        }

        export class ListsViewBase implements IModule {
            private _options: IListsViewOptions;
            private _app: AppBase;

            constructor(app: AppBase, options: IListsViewOptions) {
                if (!app) {
                    throw "App must be specified for ListView!";
                }
                this._app = app;
                this._options = $pnp.util.extend(options, { delay: 1000 });
            }

            public getLists() {
                var self = this;
                var deferred = self._app.$.Deferred();
                var url = $pnp.sp.crossDomainWeb(self._app.appWebUrl, self._app.hostWebUrl).lists.select("Id", "Title", "BaseType", "ItemCount", "Description", "Hidden", "EffectiveBasePermissions").toUrlAndQuery();
                var executor = new SP.RequestExecutor(self._app.appWebUrl);
                executor.executeAsync(<SP.RequestInfo>{
                    url: url,
                    method: "GET",
                    headers: {
                        "accept": "application/json;odata=verbose",
                        "content-Type": "application/json;odata=verbose"
                    },
                    success: (data) => {
                        var lists = JSON.parse(<string>data.body).d.results;
                        deferred.resolve(lists);
                    },
                    error: (data, errorCode, errorMessage) => {
                        if (data.body) {
                            try {
                                var error = JSON.parse(<string>data.body);
                                if (error && error.message) {
                                    errorMessage = error.message.value;
                                }
                            }
                            catch (e) { }
                        }
                        self._app.$(self._app).trigger("app-error", [errorMessage]);
                        deferred.reject(data, errorCode, errorMessage);
                    }
                });
                return deferred.promise();
            }

            public getList(listId) {
                var self = this;
                var deferred = self._app.$.Deferred();
                var url = $pnp.sp.crossDomainWeb(self._app.appWebUrl, self._app.hostWebUrl).lists.getById(listId).select("Id", "Title", "BaseType", "ItemCount", "Description", "Hidden", "EffectiveBasePermissions").toUrlAndQuery();
                var executor = new SP.RequestExecutor(self._app.appWebUrl);
                executor.executeAsync(<SP.RequestInfo>{
                    url: url,
                    method: "GET",
                    headers: {
                        "accept": "application/json;odata=verbose",
                        "content-Type": "application/json;odata=verbose"
                    },
                    success: (data) => {
                        var list = JSON.parse(<string>data.body).d;
                        deferred.resolve(list);
                    },
                    error: (data, errorCode, errorMessage) => {
                        if (data.body) {
                            try {
                                var error = JSON.parse(<string>data.body);
                                if (error && error.error) {
                                    errorMessage = error.error.message.value;
                                }
                            }
                            catch (e) { }
                        }
                        self._app.$(self._app).trigger("app-error", [errorMessage]);
                        deferred.reject(data, errorCode, errorMessage);
                    }
                });
                return deferred.promise();
            }

            public updateList(listId, properties, digestValue) {
                var self = this;
                var deferred = self._app.$.Deferred();
                var url = $pnp.sp.crossDomainWeb(self._app.appWebUrl, self._app.hostWebUrl).lists.getById(listId).toUrlAndQuery();
                var body = JSON.stringify($pnp.util.extend({
                    "__metadata": { "type": "SP.List" },
                }, properties));

                var executor = new SP.RequestExecutor(self._app.appWebUrl);
                executor.executeAsync(<SP.RequestInfo>{
                    body: body,
                    url: url,
                    method: "POST",
                    headers: {
                        "accept": "application/json;odata=verbose",
                        "content-Type": "application/json;odata=verbose",
                        "IF-Match": "*",
                        "X-HTTP-Method": "MERGE",
                        "X-RequestDigest": digestValue
                    },
                    success: (data) => {
                        deferred.resolve();
                    },
                    error: (data, errorCode, errorMessage) => {
                        if (data.body) {
                            try {
                                var error = JSON.parse(<string>data.body);
                                if (error && error.error) {
                                    errorMessage = error.error.message.value;
                                }
                            }
                            catch (e) { }
                        }
                        self._app.$(self._app).trigger("app-error", [errorMessage]);
                        deferred.reject(data, errorCode, errorMessage);
                    }
                });
                return deferred.promise();
            }

            public get_options() {
                return this._options;
            }

            public get_app(): AppBase {
                return this._app;
            }

            public render() {
                this._app.$(self).trigger("model-render");
            }
        }
    }
}
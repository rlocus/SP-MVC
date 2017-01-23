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

    export class Builder {

        private _expression: CamlBuilder.IExpression;
        private _condition: CamlBuilder.IExpression;
        private _viewXml;
        private _replace?: boolean;

        constructor(viewXml: string, replace?: boolean) {
            this._expression == null;
            this._viewXml = viewXml;
            this._replace = replace;
        }

        private getFilterLookupCondition(field: string,
            value: string | boolean | number,
            operation: FilterOperation,
            fieldType: SP.FieldType) {
            var fieldExpression = CamlBuilder.Expression();
            var condition: CamlBuilder.IExpression;
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
                    condition = numberFieldExpression.GreaterThan(<number>value);
                    break;
                case FilterOperation.Geq:
                    condition = numberFieldExpression.GreaterThanOrEqualTo(<number>value);
                    break;
                case FilterOperation.Neq:
                    if ($pnp.util.stringIsNullOrEmpty(<string>value)) {
                        condition = numberFieldExpression.IsNotNull();
                    }
                    else {
                        condition = numberFieldExpression.NotEqualTo(<number>value);
                    }
                    break;
                case FilterOperation.Eq:
                default:
                    if ($pnp.util.stringIsNullOrEmpty(<string>value)) {
                        condition = numberFieldExpression.IsNull();
                    } else {
                        condition = numberFieldExpression.EqualTo(<number>value);
                    }
                    break;
                case FilterOperation.Lt:
                    condition = numberFieldExpression.LessThan(<number>value);
                    break;
                case FilterOperation.Leq:
                    condition = numberFieldExpression.LessThanOrEqualTo(<number>value);
                    break;
            }
            return condition;
        }

        private getFilterCondition(field: string, value: string | boolean | number, operation: FilterOperation, isLookupId: boolean, fieldType: SP.FieldType) {
            var self = this;
            var condition: CamlBuilder.IExpression;
            if (Boolean(isLookupId)) {
                condition = self.getFilterLookupCondition(field, value, operation, fieldType);
            } else {
                var fieldExpression = CamlBuilder.Expression();
                switch (fieldType) {
                    case SP.FieldType.text:
                    default:
                        var textFieldExpression = fieldExpression.TextField(field);
                        switch (operation) {
                            case FilterOperation.BeginsWith:
                                condition = textFieldExpression.BeginsWith(<string>value);
                                break;
                            case FilterOperation.Contains:
                                condition = textFieldExpression.Contains(<string>value);
                                break;
                            case FilterOperation.Neq:
                                if ($pnp.util.stringIsNullOrEmpty(<string>value)) {
                                    condition = textFieldExpression.IsNotNull();
                                }
                                else {
                                    condition = textFieldExpression.NotEqualTo(<string>value);
                                }
                                break;
                            case FilterOperation.Eq:
                            default:
                                if ($pnp.util.stringIsNullOrEmpty(<string>value)) {
                                    condition = textFieldExpression.IsNull();
                                } else {
                                    condition = textFieldExpression.EqualTo(<string>value);
                                }
                                break;
                        }
                        break;
                    case SP.FieldType.lookup:
                        var lookupFieldExpression = fieldExpression.LookupField(field);
                        switch (operation) {
                            case FilterOperation.BeginsWith:
                                condition = lookupFieldExpression.ValueAsText().BeginsWith(<string>value);
                                break;
                            case FilterOperation.Contains:
                                condition = lookupFieldExpression.ValueAsText().Contains(<string>value);
                                break;
                            case FilterOperation.Neq:
                                if ($pnp.util.stringIsNullOrEmpty(<string>value)) {
                                    condition = lookupFieldExpression.ValueAsText().IsNotNull();
                                }
                                else {
                                    condition = lookupFieldExpression.ValueAsText().NotEqualTo(<string>value);
                                }
                                break;
                            case FilterOperation.Eq:
                            default:
                                if ($pnp.util.stringIsNullOrEmpty(<string>value)) {
                                    condition = lookupFieldExpression.ValueAsText().IsNull();
                                } else {
                                    condition = lookupFieldExpression.ValueAsText()
                                        .EqualTo(<string>value);
                                }
                                break;
                        }
                        break;
                    case SP.FieldType.dateTime:
                        var dateFieldExpression = fieldExpression.DateField(field);
                        switch (operation) {
                            case FilterOperation.Gt:
                                condition = dateFieldExpression.GreaterThan(<string>value);
                                break;
                            case FilterOperation.Geq:
                                condition = dateFieldExpression.GreaterThanOrEqualTo(<string>value);
                                break;
                            case FilterOperation.Neq:
                                if ($pnp.util.stringIsNullOrEmpty(<string>value)) {
                                    condition = dateFieldExpression.IsNotNull();
                                }
                                else {
                                    condition = dateFieldExpression.NotEqualTo(<string>value);
                                }
                                break;
                            case FilterOperation.Eq:
                            default:
                                if ($pnp.util.stringIsNullOrEmpty(<string>value)) {
                                    condition = dateFieldExpression.IsNull();
                                } else {
                                    condition = dateFieldExpression.EqualTo(<string>value);
                                }
                                break;
                            case FilterOperation.Lt:
                                condition = dateFieldExpression.LessThan(<string>value);
                                break;
                            case FilterOperation.Leq:
                                condition = dateFieldExpression.LessThanOrEqualTo(<string>value);
                                break;
                        }
                        break;
                    case <SP.FieldType>101: //datetime 
                        var dateTimeFieldExpression = fieldExpression.DateTimeField(field);
                        switch (operation) {
                            case FilterOperation.Gt:
                                condition = dateTimeFieldExpression.GreaterThan(<string>value);
                                break;
                            case FilterOperation.Geq:
                                condition = dateTimeFieldExpression
                                    .GreaterThanOrEqualTo(<string>value);
                                break;
                            case FilterOperation.Neq:
                                if ($pnp.util.stringIsNullOrEmpty(<string>value)) {
                                    condition = dateTimeFieldExpression.IsNotNull();
                                }
                                else {
                                    condition = dateTimeFieldExpression.NotEqualTo(<string>value);
                                }
                                break;
                            case FilterOperation.Eq:
                            default:
                                if ($pnp.util.stringIsNullOrEmpty(<string>value)) {
                                    condition = dateTimeFieldExpression.IsNull();
                                } else {
                                    condition = dateTimeFieldExpression.EqualTo(<string>value);
                                }
                                break;
                            case FilterOperation.Lt:
                                condition = dateTimeFieldExpression.LessThan(<string>value);
                                break;
                            case FilterOperation.Leq:
                                condition = dateTimeFieldExpression.LessThanOrEqualTo(<string>value);
                                break;
                        }
                        break;
                    case SP.FieldType.counter:
                        var counterFieldExpression = fieldExpression.CounterField(field);
                        switch (operation) {
                            case FilterOperation.Gt:
                                condition = counterFieldExpression.GreaterThan(<number>value);
                                break;
                            case FilterOperation.Geq:
                                condition = counterFieldExpression
                                    .GreaterThanOrEqualTo(<number>value);
                                break;
                            case FilterOperation.Neq:
                                if ($pnp.util.stringIsNullOrEmpty(<string>value)) {
                                    condition = counterFieldExpression.IsNotNull();
                                }
                                else {
                                    condition = counterFieldExpression.NotEqualTo(<number>value);
                                }
                                break;
                            case FilterOperation.Eq:
                            default:
                                if ($pnp.util.stringIsNullOrEmpty(<string>value)) {
                                    condition = counterFieldExpression.IsNull();
                                } else {
                                    condition = counterFieldExpression.EqualTo(<number>value);
                                }
                                break;
                            case FilterOperation.Lt:
                                condition = counterFieldExpression.LessThan(<number>value);
                                break;
                            case FilterOperation.Leq:
                                condition = counterFieldExpression.LessThanOrEqualTo(<number>value);
                                break;
                        }
                        break;
                    case SP.FieldType.integer:
                        var integerFieldExpression = fieldExpression.IntegerField(field);
                        switch (operation) {
                            case FilterOperation.Gt:
                                condition = integerFieldExpression.GreaterThan(<number>value);
                                break;
                            case FilterOperation.Geq:
                                condition = integerFieldExpression.GreaterThanOrEqualTo(<number>value);
                                break;
                            case FilterOperation.Neq:
                                if ($pnp.util.stringIsNullOrEmpty(<string>value)) {
                                    condition = integerFieldExpression.IsNotNull();
                                }
                                else {
                                    condition = integerFieldExpression.NotEqualTo(<number>value);
                                }
                                break;
                            case FilterOperation.Eq:
                            default:
                                if ($pnp.util.stringIsNullOrEmpty(<string>value)) {
                                    condition = integerFieldExpression.IsNull();
                                } else {
                                    condition = integerFieldExpression.EqualTo(<number>value);
                                }
                                break;
                            case FilterOperation.Lt:
                                condition = integerFieldExpression.LessThan(<number>value);
                                break;
                            case FilterOperation.Leq:
                                condition = integerFieldExpression.LessThanOrEqualTo(<number>value);
                                break;
                        }
                        break;
                    case SP.FieldType.modStat:
                        var modStatFieldExpression = fieldExpression.ModStatField(field).ValueAsText();
                        switch (operation) {
                            case FilterOperation.BeginsWith:
                                condition = modStatFieldExpression.BeginsWith(<string>value);
                                break;
                            case FilterOperation.Contains:
                                condition = modStatFieldExpression.Contains(<string>value);
                                break;
                            case FilterOperation.Neq:
                                if ($pnp.util.stringIsNullOrEmpty(<string>value)) {
                                    condition = modStatFieldExpression.IsNotNull();
                                }
                                else {
                                    condition = modStatFieldExpression.NotEqualTo(<string>value);
                                }
                                break;
                            case FilterOperation.Eq:
                            default:
                                if ($pnp.util.stringIsNullOrEmpty(<string>value)) {
                                    condition = modStatFieldExpression.IsNull();
                                } else {
                                    condition = modStatFieldExpression.EqualTo(<string>value);
                                }
                                break;
                        }
                        break;
                    case SP.FieldType.number:
                        var numberFieldExpression = fieldExpression.NumberField(field);
                        switch (operation) {
                            case FilterOperation.Gt:
                                condition = numberFieldExpression.GreaterThan(<number>value);
                                break;
                            case FilterOperation.Geq:
                                condition = numberFieldExpression.GreaterThanOrEqualTo(<number>value);
                                break;
                            case FilterOperation.Neq:
                                if ($pnp.util.stringIsNullOrEmpty(<string>value)) {
                                    condition = numberFieldExpression.IsNotNull();
                                }
                                else {
                                    condition = numberFieldExpression.NotEqualTo(<number>value);
                                }
                                break;
                            case FilterOperation.Eq:
                            default:
                                if ($pnp.util.stringIsNullOrEmpty(<string>value)) {
                                    condition = numberFieldExpression.IsNull();
                                } else {
                                    condition = numberFieldExpression.EqualTo(<number>value);
                                }
                                break;
                            case FilterOperation.Lt:
                                condition = numberFieldExpression.LessThan(<number>value);
                                break;
                            case FilterOperation.Leq:
                                condition = numberFieldExpression.LessThanOrEqualTo(<number>value);
                                break;
                        }
                        break;
                    case SP.FieldType.URL:
                        var urlFieldExpression = fieldExpression.UrlField(field);
                        switch (operation) {
                            case FilterOperation.BeginsWith:
                                condition = urlFieldExpression.BeginsWith(<string>value);
                                break;
                            case FilterOperation.Contains:
                                condition = urlFieldExpression.Contains(<string>value);
                                break;
                            case FilterOperation.Neq:
                                if ($pnp.util.stringIsNullOrEmpty(<string>value)) {
                                    condition = urlFieldExpression.IsNotNull();
                                }
                                else {
                                    condition = urlFieldExpression.NotEqualTo(<string>value);
                                }
                                break;
                            case FilterOperation.Eq:
                            default:
                                if ($pnp.util.stringIsNullOrEmpty(<string>value)) {
                                    condition = urlFieldExpression.IsNull();
                                } else {
                                    condition = urlFieldExpression.EqualTo(<string>value);
                                }
                                break;
                        }
                        break;
                    case SP.FieldType.user:
                        var userFieldExpression = fieldExpression.UserField(field).ValueAsText();
                        switch (operation) {
                            case FilterOperation.BeginsWith:
                                condition = userFieldExpression.BeginsWith(<string>value);
                                break;
                            case FilterOperation.Contains:
                                condition = userFieldExpression.Contains(<string>value);
                                break;
                            case FilterOperation.Neq:
                                if ($pnp.util.stringIsNullOrEmpty(<string>value)) {
                                    condition = userFieldExpression.IsNotNull();
                                }
                                else {
                                    condition = userFieldExpression.NotEqualTo(<string>value);
                                }
                                break;
                            case FilterOperation.Eq:
                            default:
                                if ($pnp.util.stringIsNullOrEmpty(<string>value)) {
                                    condition = userFieldExpression.IsNull();
                                } else {
                                    condition = userFieldExpression.EqualTo(<string>value);
                                }
                                break;
                        }
                        break;
                    case SP.FieldType.boolean:
                        var booleanFieldExpression = fieldExpression.BooleanField(field);
                        if ($pnp.util.stringIsNullOrEmpty(<string>value)) {
                            if (operation == FilterOperation.Neq) {
                                condition = booleanFieldExpression.IsNotNull();
                            }
                            else {
                                condition = booleanFieldExpression.IsNull();
                            }
                        } else {
                            if (Boolean(value) === true) {
                                condition = booleanFieldExpression.IsTrue();
                            } else {
                                condition = booleanFieldExpression.IsFalse();
                            }
                        }
                        break;
                }
            }
            return condition;
        }

        private getFilterMultiCondition(filter: ICamlFilter) {
            var self = this;
            var fieldExpression = CamlBuilder.Expression();
            var condition: CamlBuilder.IExpression;
            if ($pnp.util.isArray(filter.value)) {
                if ((<any[]>filter.value).length === 1) {
                    condition = self.getFilterCondition(filter.field,
                        filter.value[0],
                        filter.operation,
                        filter.lookupId,
                        filter.fieldType);
                } else if ((<any[]>filter.value).length > 1) {
                    if (Boolean(filter.lookupId)) {
                        switch (filter.fieldType) {
                            case SP.FieldType.lookup:
                            default:
                                condition = fieldExpression.LookupMultiField(filter.field).IncludesSuchItemThat().Id()
                                    .In(<number[]>filter.value);
                                break;
                            case SP.FieldType.user:
                                condition = fieldExpression.UserMultiField(filter.field).IncludesSuchItemThat().Id()
                                    .In(<number[]>filter.value);
                                break;
                        }
                    } else {
                        switch (filter.fieldType) {
                            case SP.FieldType.lookup:
                                condition = fieldExpression.LookupMultiField(filter.field).IncludesSuchItemThat()
                                    .ValueAsText().In(<string[]>filter.value);
                                break;
                            case SP.FieldType.text:
                            default:
                                condition = fieldExpression.TextField(filter.field).In(<string[]>filter.value);
                                break;
                            case SP.FieldType.dateTime:
                                condition = fieldExpression.DateField(filter.field).In(<Date[]>filter.value);
                                break;
                            case <SP.FieldType>101: //datetime                                        
                                condition = fieldExpression.DateTimeField(filter.field).In(<Date[]>filter.value);
                                break;
                            case SP.FieldType.counter:
                                condition = fieldExpression.CounterField(filter.field).In(<number[]>filter.value);
                                break;
                            case SP.FieldType.integer:
                                condition = fieldExpression.IntegerField(filter.field).In(<number[]>filter.value);
                                break;
                            case SP.FieldType.modStat:
                                condition = fieldExpression.ModStatField(filter.field).ValueAsText().In(<string[]>filter.value);
                                break;
                            case SP.FieldType.number:
                                condition = fieldExpression.NumberField(filter.field).In(<number[]>filter.value);
                                break;
                            case SP.FieldType.URL:
                                condition = fieldExpression.UrlField(filter.field).In(<string[]>filter.value);
                                break;
                            case SP.FieldType.user:
                                condition = fieldExpression.UserMultiField(filter.field).IncludesSuchItemThat()
                                    .ValueAsText().In(<string[]>filter.value);
                                break;
                        }
                    }
                }
            } else {
                condition = self.getFilterCondition(filter.field,
                    <any>filter.value,
                    filter.operation,
                    filter.lookupId,
                    filter.fieldType);
            }
            return condition;
        }

        private getAndFieldExpression(): CamlBuilder.IFieldExpression {
            if (this._expression) {
                return this._expression.And();
            }

            if ($pnp.util.stringIsNullOrEmpty(this._viewXml)) {
                return new CamlBuilder().View().Query().Where();
            }

            if (this._replace) {
                return CamlBuilder.FromXml(this._viewXml).ReplaceWhere();
            }
            return CamlBuilder.FromXml(this._viewXml).ModifyWhere().AppendAnd();
        }

        private getOrFieldExpression(): CamlBuilder.IFieldExpression {
            if (this._expression) {
                return this._expression.Or();
            }

            if ($pnp.util.stringIsNullOrEmpty(this._viewXml)) {
                return new CamlBuilder().View().Query().Where();
            }

            if (this._replace) {
                return CamlBuilder.FromXml(this._viewXml).ReplaceWhere();
            }

            return CamlBuilder.FromXml(this._viewXml).ModifyWhere().AppendOr();
        }

        private getConditions(filters: Array<ICamlFilter>) {
            var conditions = new Array<CamlBuilder.IExpression>();
            if (filters) {
                for (var i in filters) {
                    if (filters.hasOwnProperty(i)) {
                        var filter = filters[i];
                        if (filter) {
                            var condition = this.getFilterMultiCondition(filter);
                            if (condition) {
                                conditions.push(condition);
                            }
                        }
                    }
                }
            }
            return conditions;
        }

        private getAndWithAllCondition(conditions: Array<CamlBuilder.IExpression>) {
            if (this._condition) {
                return this._condition.And().All(conditions);
            }
            return CamlBuilder.Expression().All(conditions);
        }

        private getAndWithAnyCondition(conditions: Array<CamlBuilder.IExpression>) {
            if (this._condition) {
                return this._condition.And().Any(conditions);
            }
            return CamlBuilder.Expression().Any(conditions);
        }

        private getOrWithAllCondition(conditions: Array<CamlBuilder.IExpression>) {
            if (this._condition) {
                return this._condition.Or().All(conditions);
            }
            return CamlBuilder.Expression().All(conditions);
        }

        private getOrWithAnyCondition(conditions: Array<CamlBuilder.IExpression>) {
            if (this._condition) {
                return this._condition.And().Any(conditions);
            }
            return CamlBuilder.Expression().Any(conditions);
        }

        public appendOr(...filters: Array<ICamlFilter>) {
            var conditions = this.getConditions(filters);
            if (conditions.length > 0) {
                this._expression = this.getOrFieldExpression().Any(conditions);
                this._condition = this.getOrWithAnyCondition(this.getConditions(filters));
            }
            return this;
        }

        public appendAnd(...filters: Array<ICamlFilter>) {
            var conditions = this.getConditions(filters);
            if (conditions.length > 0) {
                this._expression = this.getAndFieldExpression().All(conditions);
                this._condition = this.getAndWithAllCondition(this.getConditions(filters));
            }
            return this;
        }

        public appendOrWithAll(...filters: Array<ICamlFilter>) {
            var conditions = this.getConditions(filters);
            if (conditions.length > 0) {
                this._expression = this.getOrFieldExpression().All(conditions);
                this._condition = this.getOrWithAllCondition(this.getConditions(filters));
            }
            return this;
        }

        public appendAndWithAny(...filters: Array<ICamlFilter>) {
            var conditions = this.getConditions(filters);
            if (conditions.length > 0) {
                this._expression = this.getAndFieldExpression().Any(conditions);
                this._condition = this.getAndWithAnyCondition(this.getConditions(filters));
            }
            return this;
        }

        public combineAll(...builders: Array<Builder>) {
            var conditions = new Array<CamlBuilder.IExpression>();
            for (var i in builders) {
                var builder = builders[i];
                if (builder && builder._condition) {
                    conditions.push(builder._condition);
                }
            }
            if (conditions.length > 0) {
                this._expression = this.getAndFieldExpression().All(conditions);
                this._condition = this.getAndWithAllCondition(conditions);
            }
            return this;
        }

        public combineAny(...builders: Array<Builder>) {
            var conditions = new Array<CamlBuilder.IExpression>();
            for (var i in builders) {
                var builder = builders[i];
                if (builder && builder._condition) {
                    conditions.push(builder._condition);
                }
            }
            if (conditions.length > 0) {
                this._expression = this.getOrFieldExpression().Any(conditions);
                this._condition = this.getOrWithAnyCondition(conditions);
            }
            return this;
        }

        public clear() {
            this._expression = null;
            this._condition = null;
        }

        public toString() {
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
                    queryBuilder: new Caml.Builder(options.viewXml, false)
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
                        viewXml = self._options.queryBuilder.toString();
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
                    var viewXml = self._options.viewXml;
                    if (self._options.queryBuilder) {
                        viewXml = self._options.queryBuilder.toString();
                    }
                    query.query.add("@viewXml", "'" + encodeURIComponent(viewXml) + "'");
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
                    var viewXml = self._options.viewXml;
                    if (self._options.queryBuilder) {
                        viewXml = self._options.queryBuilder.toString();
                    }
                    var postBody = JSON.stringify({ "query": { "__metadata": { "type": "SP.CamlQuery" }, "ViewXml": viewXml } });
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
                    var viewXml = self._options.viewXml;
                    if (self._options.queryBuilder) {
                        viewXml = self._options.queryBuilder.toString();
                    }
                    camlQuery.set_viewXml(viewXml);
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
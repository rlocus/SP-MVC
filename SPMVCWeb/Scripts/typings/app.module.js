"use strict";
var __extends = (this && this.__extends) || (function () {
    var extendStatics = function (d, b) {
        extendStatics = Object.setPrototypeOf ||
            ({ __proto__: [] } instanceof Array && function (d, b) { d.__proto__ = b; }) ||
            function (d, b) { for (var p in b) if (b.hasOwnProperty(p)) d[p] = b[p]; };
        return extendStatics(d, b);
    };
    return function (d, b) {
        extendStatics(d, b);
        function __() { this.constructor = d; }
        d.prototype = b === null ? Object.create(b) : (__.prototype = b.prototype, new __());
    };
})();
exports.__esModule = true;
var $pnp = require("pnp");
var Caml;
(function (Caml) {
    "use strict";
    var FilterOperation;
    (function (FilterOperation) {
        FilterOperation[FilterOperation["Eq"] = 0] = "Eq";
        FilterOperation[FilterOperation["Neq"] = 1] = "Neq";
        FilterOperation[FilterOperation["Gt"] = 2] = "Gt";
        FilterOperation[FilterOperation["Lt"] = 3] = "Lt";
        FilterOperation[FilterOperation["Geq"] = 4] = "Geq";
        FilterOperation[FilterOperation["Leq"] = 5] = "Leq";
        FilterOperation[FilterOperation["BeginsWith"] = 6] = "BeginsWith";
        FilterOperation[FilterOperation["Contains"] = 7] = "Contains";
        FilterOperation[FilterOperation["In"] = 8] = "In";
    })(FilterOperation = Caml.FilterOperation || (Caml.FilterOperation = {}));
    var Builder = (function () {
        function Builder(limit, paged, orderBy, sortAsc, scope, viewFields) {
            this._expression == null;
            this._paged = paged;
            this._limit = limit;
            this._orderBy = orderBy;
            this._sortAsc = sortAsc;
            this._scope = scope;
            this._viewFields = viewFields;
        }
        Builder.prototype.getFilterLookupCondition = function (field, value, operation, fieldType) {
            if ($pnp.util.stringIsNullOrEmpty(field)) {
                throw "Field Internal Name cannot be empty.";
            }
            var fieldExpression = CamlBuilder.Expression();
            var condition;
            var numberFieldExpression;
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
                    condition = numberFieldExpression.GreaterThan(value);
                    break;
                case FilterOperation.Geq:
                    condition = numberFieldExpression.GreaterThanOrEqualTo(value);
                    break;
                case FilterOperation.Neq:
                    if ($pnp.util.stringIsNullOrEmpty(value)) {
                        condition = numberFieldExpression.IsNotNull();
                    }
                    else {
                        condition = numberFieldExpression.NotEqualTo(value);
                    }
                    break;
                case FilterOperation.Eq:
                default:
                    if ($pnp.util.stringIsNullOrEmpty(value)) {
                        condition = numberFieldExpression.IsNull();
                    }
                    else {
                        condition = numberFieldExpression.EqualTo(value);
                    }
                    break;
                case FilterOperation.Lt:
                    condition = numberFieldExpression.LessThan(value);
                    break;
                case FilterOperation.Leq:
                    condition = numberFieldExpression.LessThanOrEqualTo(value);
                    break;
                case FilterOperation.BeginsWith:
                case FilterOperation.Contains:
                    throw "Lookup field: '" + field + "'. Operation " + operation + " is not supported.";
            }
            return condition;
        };
        Builder.prototype.getFilterCondition = function (field, value, operation, isLookupId, fieldType) {
            var self = this;
            if ($pnp.util.stringIsNullOrEmpty(field)) {
                throw "Field Internal Name cannot be empty.";
            }
            var condition;
            if (Boolean(isLookupId)) {
                condition = self.getFilterLookupCondition(field, value, operation, fieldType);
            }
            else {
                var fieldExpression = CamlBuilder.Expression();
                switch (fieldType) {
                    case SP.FieldType.text:
                    default:
                        var textFieldExpression = fieldExpression.TextField(field);
                        switch (operation) {
                            case FilterOperation.BeginsWith:
                                condition = textFieldExpression.BeginsWith(value);
                                break;
                            case FilterOperation.Contains:
                                condition = textFieldExpression.Contains(value);
                                break;
                            case FilterOperation.Neq:
                                if ($pnp.util.stringIsNullOrEmpty(value)) {
                                    condition = textFieldExpression.IsNotNull();
                                }
                                else {
                                    condition = textFieldExpression.NotEqualTo(value);
                                }
                                break;
                            case FilterOperation.Eq:
                            default:
                                if ($pnp.util.stringIsNullOrEmpty(value)) {
                                    condition = textFieldExpression.IsNull();
                                }
                                else {
                                    condition = textFieldExpression.EqualTo(value);
                                }
                                break;
                        }
                        break;
                    case SP.FieldType.lookup:
                        var lookupFieldExpression = fieldExpression.LookupField(field);
                        switch (operation) {
                            case FilterOperation.BeginsWith:
                                condition = lookupFieldExpression.ValueAsText().BeginsWith(value);
                                break;
                            case FilterOperation.Contains:
                                condition = lookupFieldExpression.ValueAsText().Contains(value);
                                break;
                            case FilterOperation.Neq:
                                if ($pnp.util.stringIsNullOrEmpty(value)) {
                                    condition = lookupFieldExpression.ValueAsText().IsNotNull();
                                }
                                else {
                                    condition = lookupFieldExpression.ValueAsText().NotEqualTo(value);
                                }
                                break;
                            case FilterOperation.Eq:
                            default:
                                if ($pnp.util.stringIsNullOrEmpty(value)) {
                                    condition = lookupFieldExpression.ValueAsText().IsNull();
                                }
                                else {
                                    condition = lookupFieldExpression.ValueAsText().EqualTo(value);
                                }
                                break;
                        }
                        break;
                    case SP.FieldType.dateTime:
                        var dateFieldExpression = fieldExpression.DateField(field);
                        switch (operation) {
                            case FilterOperation.Gt:
                                condition = dateFieldExpression.GreaterThan(value);
                                break;
                            case FilterOperation.Geq:
                                condition = dateFieldExpression.GreaterThanOrEqualTo(value);
                                break;
                            case FilterOperation.Neq:
                                if ($pnp.util.stringIsNullOrEmpty(value)) {
                                    condition = dateFieldExpression.IsNotNull();
                                }
                                else {
                                    condition = dateFieldExpression.NotEqualTo(value);
                                }
                                break;
                            case FilterOperation.Eq:
                            default:
                                if ($pnp.util.stringIsNullOrEmpty(value)) {
                                    condition = dateFieldExpression.IsNull();
                                }
                                else {
                                    condition = dateFieldExpression.EqualTo(value);
                                }
                                break;
                            case FilterOperation.Lt:
                                condition = dateFieldExpression.LessThan(value);
                                break;
                            case FilterOperation.Leq:
                                condition = dateFieldExpression.LessThanOrEqualTo(value);
                                break;
                            case FilterOperation.BeginsWith:
                            case FilterOperation.Contains:
                                throw "Date field: '" + field + "'. Operation " + operation + " is not supported.";
                        }
                        break;
                    case 101:
                        var dateTimeFieldExpression = fieldExpression.DateTimeField(field);
                        switch (operation) {
                            case FilterOperation.Gt:
                                condition = dateTimeFieldExpression.GreaterThan(value);
                                break;
                            case FilterOperation.Geq:
                                condition = dateTimeFieldExpression
                                    .GreaterThanOrEqualTo(value);
                                break;
                            case FilterOperation.Neq:
                                if ($pnp.util.stringIsNullOrEmpty(value)) {
                                    condition = dateTimeFieldExpression.IsNotNull();
                                }
                                else {
                                    condition = dateTimeFieldExpression.NotEqualTo(value);
                                }
                                break;
                            case FilterOperation.Eq:
                            default:
                                if ($pnp.util.stringIsNullOrEmpty(value)) {
                                    condition = dateTimeFieldExpression.IsNull();
                                }
                                else {
                                    condition = dateTimeFieldExpression.EqualTo(value);
                                }
                                break;
                            case FilterOperation.Lt:
                                condition = dateTimeFieldExpression.LessThan(value);
                                break;
                            case FilterOperation.Leq:
                                condition = dateTimeFieldExpression.LessThanOrEqualTo(value);
                                break;
                            case FilterOperation.BeginsWith:
                            case FilterOperation.Contains:
                                throw "Date time field: '" + field + "'. Operation " + operation + " is not supported.";
                        }
                        break;
                    case SP.FieldType.counter:
                        var counterFieldExpression = fieldExpression.CounterField(field);
                        switch (operation) {
                            case FilterOperation.Gt:
                                condition = counterFieldExpression.GreaterThan(value);
                                break;
                            case FilterOperation.Geq:
                                condition = counterFieldExpression
                                    .GreaterThanOrEqualTo(value);
                                break;
                            case FilterOperation.Neq:
                                if ($pnp.util.stringIsNullOrEmpty(value)) {
                                    condition = counterFieldExpression.IsNotNull();
                                }
                                else {
                                    condition = counterFieldExpression.NotEqualTo(value);
                                }
                                break;
                            case FilterOperation.Eq:
                            default:
                                if ($pnp.util.stringIsNullOrEmpty(value)) {
                                    condition = counterFieldExpression.IsNull();
                                }
                                else {
                                    condition = counterFieldExpression.EqualTo(value);
                                }
                                break;
                            case FilterOperation.Lt:
                                condition = counterFieldExpression.LessThan(value);
                                break;
                            case FilterOperation.Leq:
                                condition = counterFieldExpression.LessThanOrEqualTo(value);
                                break;
                            case FilterOperation.BeginsWith:
                            case FilterOperation.Contains:
                                throw "Counter field: '" + field + "'. Operation " + operation + " is not supported.";
                        }
                        break;
                    case SP.FieldType.integer:
                        var integerFieldExpression = fieldExpression.IntegerField(field);
                        switch (operation) {
                            case FilterOperation.Gt:
                                condition = integerFieldExpression.GreaterThan(value);
                                break;
                            case FilterOperation.Geq:
                                condition = integerFieldExpression.GreaterThanOrEqualTo(value);
                                break;
                            case FilterOperation.Neq:
                                if ($pnp.util.stringIsNullOrEmpty(value)) {
                                    condition = integerFieldExpression.IsNotNull();
                                }
                                else {
                                    condition = integerFieldExpression.NotEqualTo(value);
                                }
                                break;
                            case FilterOperation.Eq:
                            default:
                                if ($pnp.util.stringIsNullOrEmpty(value)) {
                                    condition = integerFieldExpression.IsNull();
                                }
                                else {
                                    condition = integerFieldExpression.EqualTo(value);
                                }
                                break;
                            case FilterOperation.Lt:
                                condition = integerFieldExpression.LessThan(value);
                                break;
                            case FilterOperation.Leq:
                                condition = integerFieldExpression.LessThanOrEqualTo(value);
                                break;
                            case FilterOperation.BeginsWith:
                            case FilterOperation.Contains:
                                throw "Integer field: '" + field + "'. Operation " + operation + " is not supported.";
                        }
                        break;
                    case SP.FieldType.modStat:
                        var modStatFieldExpression = fieldExpression.ModStatField(field).ValueAsText();
                        switch (operation) {
                            case FilterOperation.BeginsWith:
                                condition = modStatFieldExpression.BeginsWith(value);
                                break;
                            case FilterOperation.Contains:
                                condition = modStatFieldExpression.Contains(value);
                                break;
                            case FilterOperation.Neq:
                                if ($pnp.util.stringIsNullOrEmpty(value)) {
                                    condition = modStatFieldExpression.IsNotNull();
                                }
                                else {
                                    condition = modStatFieldExpression.NotEqualTo(value);
                                }
                                break;
                            case FilterOperation.Eq:
                            default:
                                if ($pnp.util.stringIsNullOrEmpty(value)) {
                                    condition = modStatFieldExpression.IsNull();
                                }
                                else {
                                    condition = modStatFieldExpression.EqualTo(value);
                                }
                                break;
                        }
                        break;
                    case SP.FieldType.number:
                        var numberFieldExpression = fieldExpression.NumberField(field);
                        switch (operation) {
                            case FilterOperation.Gt:
                                condition = numberFieldExpression.GreaterThan(value);
                                break;
                            case FilterOperation.Geq:
                                condition = numberFieldExpression.GreaterThanOrEqualTo(value);
                                break;
                            case FilterOperation.Neq:
                                if ($pnp.util.stringIsNullOrEmpty(value)) {
                                    condition = numberFieldExpression.IsNotNull();
                                }
                                else {
                                    condition = numberFieldExpression.NotEqualTo(value);
                                }
                                break;
                            case FilterOperation.Eq:
                            default:
                                if ($pnp.util.stringIsNullOrEmpty(value)) {
                                    condition = numberFieldExpression.IsNull();
                                }
                                else {
                                    condition = numberFieldExpression.EqualTo(value);
                                }
                                break;
                            case FilterOperation.Lt:
                                condition = numberFieldExpression.LessThan(value);
                                break;
                            case FilterOperation.Leq:
                                condition = numberFieldExpression.LessThanOrEqualTo(value);
                                break;
                            case FilterOperation.BeginsWith:
                            case FilterOperation.Contains:
                                throw "Number field: '" + field + "'. Operation " + operation + " is not supported.";
                        }
                        break;
                    case SP.FieldType.URL:
                        var urlFieldExpression = fieldExpression.UrlField(field);
                        switch (operation) {
                            case FilterOperation.BeginsWith:
                                condition = urlFieldExpression.BeginsWith(value);
                                break;
                            case FilterOperation.Contains:
                                condition = urlFieldExpression.Contains(value);
                                break;
                            case FilterOperation.Neq:
                                if ($pnp.util.stringIsNullOrEmpty(value)) {
                                    condition = urlFieldExpression.IsNotNull();
                                }
                                else {
                                    condition = urlFieldExpression.NotEqualTo(value);
                                }
                                break;
                            case FilterOperation.Eq:
                            default:
                                if ($pnp.util.stringIsNullOrEmpty(value)) {
                                    condition = urlFieldExpression.IsNull();
                                }
                                else {
                                    condition = urlFieldExpression.EqualTo(value);
                                }
                                break;
                        }
                        break;
                    case SP.FieldType.user:
                        var userFieldExpression = fieldExpression.UserField(field).ValueAsText();
                        switch (operation) {
                            case FilterOperation.BeginsWith:
                                condition = userFieldExpression.BeginsWith(value);
                                break;
                            case FilterOperation.Contains:
                                condition = userFieldExpression.Contains(value);
                                break;
                            case FilterOperation.Neq:
                                if ($pnp.util.stringIsNullOrEmpty(value)) {
                                    condition = userFieldExpression.IsNotNull();
                                }
                                else {
                                    condition = userFieldExpression.NotEqualTo(value);
                                }
                                break;
                            case FilterOperation.Eq:
                            default:
                                if ($pnp.util.stringIsNullOrEmpty(value)) {
                                    condition = userFieldExpression.IsNull();
                                }
                                else {
                                    condition = userFieldExpression.EqualTo(value);
                                }
                                break;
                        }
                        break;
                    case SP.FieldType.boolean:
                        if (operation == FilterOperation.BeginsWith || operation == FilterOperation.Contains) {
                            throw "Boolean field: '" + field + "'. Operation " + operation + " is not supported.";
                        }
                        var booleanFieldExpression = fieldExpression.BooleanField(field);
                        if ($pnp.util.stringIsNullOrEmpty(value)) {
                            if (operation == FilterOperation.Neq) {
                                condition = booleanFieldExpression.IsNotNull();
                            }
                            else {
                                condition = booleanFieldExpression.IsNull();
                            }
                        }
                        else {
                            if (Boolean(value) === true) {
                                condition = booleanFieldExpression.IsTrue();
                            }
                            else {
                                condition = booleanFieldExpression.IsFalse();
                            }
                        }
                        break;
                }
            }
            return condition;
        };
        Builder.prototype.getFilterMultiCondition = function (filter) {
            var self = this;
            var fieldExpression = CamlBuilder.Expression();
            var condition;
            if ($pnp.util.stringIsNullOrEmpty(filter.field)) {
                throw "Field Internal Name cannot be empty.";
            }
            if ($pnp.util.isArray(filter.value)) {
                if (filter.value.length === 1) {
                    condition = self.getFilterCondition(filter.field, filter.value[0], filter.operation, filter.lookupId, filter.fieldType);
                }
                else if (filter.value.length > 1) {
                    if (!$pnp.util.stringIsNullOrEmpty(filter.operation) && filter.operation != FilterOperation.In) {
                        throw "Field: '" + filter.field + "'. Operation " + filter.operation + " is not supported.";
                    }
                    if (Boolean(filter.lookupId)) {
                        switch (filter.fieldType) {
                            case SP.FieldType.lookup:
                            default:
                                condition = fieldExpression.LookupMultiField(filter.field).IncludesSuchItemThat().Id().In(filter.value);
                                break;
                            case SP.FieldType.user:
                                condition = fieldExpression.UserMultiField(filter.field).IncludesSuchItemThat().Id().In(filter.value);
                                break;
                        }
                    }
                    else {
                        switch (filter.fieldType) {
                            case SP.FieldType.lookup:
                                condition = fieldExpression.LookupMultiField(filter.field).IncludesSuchItemThat().ValueAsText().In(filter.value);
                                break;
                            case SP.FieldType.text:
                            default:
                                condition = fieldExpression.TextField(filter.field).In(filter.value);
                                break;
                            case SP.FieldType.dateTime:
                                condition = fieldExpression.DateField(filter.field).In(filter.value);
                                break;
                            case 101:
                                condition = fieldExpression.DateTimeField(filter.field).In(filter.value);
                                break;
                            case SP.FieldType.counter:
                                condition = fieldExpression.CounterField(filter.field).In(filter.value);
                                break;
                            case SP.FieldType.integer:
                                condition = fieldExpression.IntegerField(filter.field).In(filter.value);
                                break;
                            case SP.FieldType.modStat:
                                condition = fieldExpression.ModStatField(filter.field).ValueAsText().In(filter.value);
                                break;
                            case SP.FieldType.number:
                                condition = fieldExpression.NumberField(filter.field).In(filter.value);
                                break;
                            case SP.FieldType.URL:
                                condition = fieldExpression.UrlField(filter.field).In(filter.value);
                                break;
                            case SP.FieldType.user:
                                condition = fieldExpression.UserMultiField(filter.field).IncludesSuchItemThat()
                                    .ValueAsText().In(filter.value);
                                break;
                        }
                    }
                }
            }
            else {
                condition = self.getFilterCondition(filter.field, filter.value, filter.operation, filter.lookupId, filter.fieldType);
            }
            return condition;
        };
        Builder.prototype.getQuery = function () {
            var view = new CamlBuilder().View(this._viewFields);
            if (!$pnp.util.stringIsNullOrEmpty(this._limit) || !$pnp.util.stringIsNullOrEmpty(this._paged)) {
                view = view.RowLimit(this._limit, this._paged);
            }
            if (!$pnp.util.stringIsNullOrEmpty(this._scope)) {
                view = view.Scope(this._scope);
            }
            var query = view.Query();
            if (!$pnp.util.stringIsNullOrEmpty(this._orderBy)) {
                query.OrderBy(this._orderBy);
            }
            return query;
        };
        Builder.prototype.getAndFieldExpression = function () {
            if (this._expression) {
                return this._expression.And();
            }
            return this.getQuery().Where();
        };
        Builder.prototype.getOrFieldExpression = function () {
            if (this._expression) {
                return this._expression.Or();
            }
            return this.getQuery().Where();
        };
        Builder.prototype.getConditions = function (filters) {
            var conditions = new Array();
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
        };
        Builder.prototype.getAndWithAllCondition = function (conditions) {
            if (this._condition) {
                return this._condition.And().All(conditions);
            }
            return CamlBuilder.Expression().All(conditions);
        };
        Builder.prototype.getAndWithAnyCondition = function (conditions) {
            if (this._condition) {
                return this._condition.And().Any(conditions);
            }
            return CamlBuilder.Expression().Any(conditions);
        };
        Builder.prototype.getOrWithAllCondition = function (conditions) {
            if (this._condition) {
                return this._condition.Or().All(conditions);
            }
            return CamlBuilder.Expression().All(conditions);
        };
        Builder.prototype.getOrWithAnyCondition = function (conditions) {
            if (this._condition) {
                return this._condition.And().Any(conditions);
            }
            return CamlBuilder.Expression().Any(conditions);
        };
        Builder.prototype.appendOr = function () {
            var filters = [];
            for (var _i = 0; _i < arguments.length; _i++) {
                filters[_i] = arguments[_i];
            }
            var conditions = this.getConditions(filters);
            if (conditions.length > 0) {
                this._expression = this.getOrFieldExpression().Any(conditions);
                this._condition = this.getOrWithAnyCondition(this.getConditions(filters));
            }
            return this;
        };
        Builder.prototype.appendAnd = function () {
            var filters = [];
            for (var _i = 0; _i < arguments.length; _i++) {
                filters[_i] = arguments[_i];
            }
            var conditions = this.getConditions(filters);
            if (conditions.length > 0) {
                this._expression = this.getAndFieldExpression().All(conditions);
                this._condition = this.getAndWithAllCondition(this.getConditions(filters));
            }
            return this;
        };
        Builder.prototype.appendOrWithAll = function () {
            var filters = [];
            for (var _i = 0; _i < arguments.length; _i++) {
                filters[_i] = arguments[_i];
            }
            var conditions = this.getConditions(filters);
            if (conditions.length > 0) {
                this._expression = this.getOrFieldExpression().All(conditions);
                this._condition = this.getOrWithAllCondition(this.getConditions(filters));
            }
            return this;
        };
        Builder.prototype.appendAndWithAny = function () {
            var filters = [];
            for (var _i = 0; _i < arguments.length; _i++) {
                filters[_i] = arguments[_i];
            }
            var conditions = this.getConditions(filters);
            if (conditions.length > 0) {
                this._expression = this.getAndFieldExpression().Any(conditions);
                this._condition = this.getAndWithAnyCondition(this.getConditions(filters));
            }
            return this;
        };
        Builder.prototype.combineAll = function () {
            var builders = [];
            for (var _i = 0; _i < arguments.length; _i++) {
                builders[_i] = arguments[_i];
            }
            var conditions = new Array();
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
        };
        Builder.prototype.combineAny = function () {
            var builders = [];
            for (var _i = 0; _i < arguments.length; _i++) {
                builders[_i] = arguments[_i];
            }
            var conditions = new Array();
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
        };
        Builder.prototype.clear = function () {
            delete this._expression;
            delete this._condition;
            delete this._viewXml;
            this._expression = null;
            this._condition = null;
        };
        Builder.prototype.toString = function () {
            if (this._expression) {
                var viewXml = this._expression.ToString();
                this.clear();
                this._viewXml = viewXml;
            }
            if ($pnp.util.stringIsNullOrEmpty(this._viewXml)) {
                this._viewXml = this.getQuery().ToString();
            }
            return this._viewXml;
        };
        return Builder;
    }());
    Caml.Builder = Builder;
    var ReBuilder = (function (_super) {
        __extends(ReBuilder, _super);
        function ReBuilder(viewXml, replace) {
            var _this = _super.call(this) || this;
            _this._originalViewXml = viewXml;
            _this._replace = replace;
            return _this;
        }
        ReBuilder.prototype.getAndFieldExpression = function () {
            if (this._expression) {
                return this._expression.And();
            }
            if ($pnp.util.stringIsNullOrEmpty(this._originalViewXml)) {
                return _super.prototype.getAndFieldExpression.call(this);
            }
            if (this._replace) {
                return CamlBuilder.FromXml(this._originalViewXml).ReplaceWhere();
            }
            return CamlBuilder.FromXml(this._originalViewXml).ModifyWhere().AppendAnd();
        };
        ReBuilder.prototype.getOrFieldExpression = function () {
            if (this._expression) {
                return this._expression.Or();
            }
            if ($pnp.util.stringIsNullOrEmpty(this._originalViewXml)) {
                return _super.prototype.getOrFieldExpression.call(this);
            }
            if (this._replace) {
                return CamlBuilder.FromXml(this._originalViewXml).ReplaceWhere();
            }
            return CamlBuilder.FromXml(this._originalViewXml).ModifyWhere().AppendOr();
        };
        ReBuilder.prototype.toString = function () {
            if (this._expression) {
                var viewXml = this._expression.ToString();
                this.clear();
                this._viewXml = viewXml;
            }
            if ($pnp.util.stringIsNullOrEmpty(this._viewXml)) {
                if (!$pnp.util.stringIsNullOrEmpty(this._originalViewXml)) {
                    this._viewXml = this._originalViewXml;
                }
                else {
                    _super.prototype.toString.call(this);
                }
            }
            return this._viewXml;
        };
        return ReBuilder;
    }(Builder));
    Caml.ReBuilder = ReBuilder;
})(Caml = exports.Caml || (exports.Caml = {}));
window["Caml"] = Caml;
var App;
(function (App) {
    "use strict";
    var AppBase = (function () {
        function AppBase() {
            this.delay = (function () {
                var timer = 0;
                return function (callback, ms) {
                    clearTimeout(timer);
                    timer = setTimeout(callback, ms);
                };
            })();
            this._initialized = false;
            this._scriptPromises = {};
        }
        AppBase.prototype.init = function (preloadedScripts) {
            var self = this;
            if (preloadedScripts) {
                var $_1 = preloadedScripts["jquery"];
                self.$ = $_1;
                if (self.$) {
                    self.$.cachedScript = function (url, options) {
                        options = self.$.extend(options || {}, {
                            dataType: "script",
                            cache: true,
                            url: url
                        });
                        return self.$.ajax(options);
                    };
                }
                else {
                    throw "jQuery is not loaded!";
                }
            }
            this.hostWebUrl = window._spPageContextInfo && !$pnp.util.stringIsNullOrEmpty(window._spPageContextInfo.webAbsoluteUrl) ? window._spPageContextInfo.webAbsoluteUrl : $pnp.util.getUrlParamByName("SPHostUrl");
            if ($pnp.util.stringIsNullOrEmpty(this.hostWebUrl)) {
                throw "SPHostUrl url parameter must be specified!";
            }
            this.appWebUrl = window._spPageContextInfo && !$pnp.util.stringIsNullOrEmpty(window._spPageContextInfo.appWebUrl) ? window._spPageContextInfo.appWebUrl : $pnp.util.getUrlParamByName("SPAppWebUrl");
            if ($pnp.util.stringIsNullOrEmpty(this.appWebUrl)) {
                throw "SPAppWebUrl url parameter must be specified!";
            }
            this.scriptBase = $pnp.util.combinePaths(this.hostWebUrl, window._spPageContextInfo && !$pnp.util.stringIsNullOrEmpty(window._spPageContextInfo.layoutsUrl) ? window._spPageContextInfo.layoutsUrl : "_layouts/15");
            this._initialized = true;
            self.$(self).trigger("app-init");
        };
        AppBase.prototype.ensureScript = function (url) {
            if (url) {
                url = url.toLowerCase().replace("~sphost", this.hostWebUrl)
                    .replace("~spapp", this.appWebUrl)
                    .replace("~splayouts", this.scriptBase);
                var scriptPromise = this._scriptPromises[url];
                if (!scriptPromise) {
                    scriptPromise = this.$.cachedScript(url);
                    this._scriptPromises[url] = scriptPromise;
                }
                return scriptPromise;
            }
            return null;
        };
        AppBase.prototype.render = function (modules) {
            var self = this;
            if (!self._initialized) {
                throw "App is not initialized!";
            }
            self.ensureScript("~splayouts/MicrosoftAjax.js").then(function () {
                self.ensureScript("~splayouts/SP.Runtime.js").then(function () {
                    self.ensureScript("~splayouts/SP.RequestExecutor.js").then(function () {
                        self.ensureScript("~splayouts/SP.js").then(function () {
                            if ($pnp.util.isArray(modules)) {
                                self.$.each(modules, function (i, module) {
                                    module.render();
                                });
                            }
                            self.$(self).trigger("app-render");
                        });
                    });
                });
            });
        };
        AppBase.prototype.is_initialized = function () {
            var self = this;
            return self._initialized;
        };
        AppBase.prototype.get_hostWebUrl = function () {
            var self = this;
            return self.hostWebUrl;
        };
        AppBase.prototype.get_appWebUrl = function () {
            var self = this;
            return self.appWebUrl;
        };
        AppBase.prototype.get_BasePermissions = function (permMask) {
            var permissions = new SP.BasePermissions();
            if (permMask) {
                var permMaskHigh = permMask.length <= 10 ? 0 : parseInt(permMask.substring(2, permMask.length - 8), 16);
                var permMaskLow = permMask.length <= 10 ? parseInt(permMask) : parseInt(permMask.substring(permMask.length - 8, permMask.length), 16);
                permissions.initPropertiesFromJson({ "High": permMaskHigh, "Low": permMaskLow });
            }
            return permissions;
        };
        AppBase.prototype.getQueryParam = function (url, name) {
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
                    var qParameters = search.split("&");
                    for (var i in qParameters) {
                        if (qParameters.hasOwnProperty(i)) {
                            var qParameter = qParameters[i].split("=");
                            var key = decodeURIComponent(self.$(qParameter).get(0));
                            if (key && key.toUpperCase() === name.toUpperCase()) {
                                var value = decodeURIComponent(self.$(qParameter).get(1));
                                return value;
                            }
                        }
                    }
                }
            }
            return "";
        };
        return AppBase;
    }());
    App.AppBase = AppBase;
    var Module;
    (function (Module) {
        var RenderMethod;
        (function (RenderMethod) {
            RenderMethod[RenderMethod["Default"] = 0] = "Default";
            RenderMethod[RenderMethod["RenderListDataAsStream"] = 1] = "RenderListDataAsStream";
            RenderMethod[RenderMethod["RenderListFilterData"] = 2] = "RenderListFilterData";
            RenderMethod[RenderMethod["RenderListData"] = 3] = "RenderListData";
            RenderMethod[RenderMethod["GetItems"] = 4] = "GetItems";
        })(RenderMethod = Module.RenderMethod || (Module.RenderMethod = {}));
        var ListViewBase = (function () {
            function ListViewBase(app, options) {
                if (!app) {
                    throw "App must be specified for ListView!";
                }
                this._app = app;
                this._options = $pnp.util.extend({
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
                    scope: null,
                    rootFolder: null,
                    appendRows: null,
                    renderMethod: null,
                    renderOptions: null,
                    queryStringFilters: null,
                    queryBuilder: !$pnp.util.stringIsNullOrEmpty(options.viewXml)
                        ? new Caml.ReBuilder(options.viewXml, false)
                        : new Caml.Builder(options.limit, options.paged, options.orderBy, options.sortAsc, options.scope, options.viewFields)
                }, options);
            }
            ListViewBase.prototype.addToken = function (query, token) {
                var self = this;
                if (token) {
                    while (token.startsWith("?") || token.startsWith("&")) {
                        token = token.slice(1, token.length);
                    }
                    var qParameters = token.split("&");
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
            };
            ListViewBase.prototype.addFilterQuery = function (query) {
                var self = this;
                if (self._app.$.isArray(self._options.queryStringFilters)) {
                    for (var i = 0; i < self._options.queryStringFilters.length; i++) {
                        var filter = self._options.queryStringFilters[i];
                        if (filter) {
                            if (self._app.$.isArray(filter.value) && filter.value.length > 1) {
                                query.add("FilterFields" + (i + 1), encodeURIComponent(filter.field));
                                query.add("FilterValues" + (i + 1), encodeURIComponent(filter.value.join(";#")));
                            }
                            else {
                                query.add("FilterField" + (i + 1), encodeURIComponent(filter.field));
                                query.add("FilterValue" + (i + 1), encodeURIComponent(filter.value));
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
            };
            ListViewBase.prototype.getList = function () {
                var self = this;
                if (!$pnp.util.stringIsNullOrEmpty(self._options.listId)) {
                    return $pnp.sp.crossDomainWeb(self._app.appWebUrl, self._app.hostWebUrl).lists.getById(self._options.listId);
                }
                else if (!$pnp.util.stringIsNullOrEmpty(self._options.listUrl)) {
                    return $pnp.sp.crossDomainWeb(self._app.appWebUrl, self._app.hostWebUrl).getList(self._options.listUrl);
                }
                else if (!$pnp.util.stringIsNullOrEmpty(self._options.listTitle)) {
                    return $pnp.sp.crossDomainWeb(self._app.appWebUrl, self._app.hostWebUrl).lists.getByTitle(self._options.listId);
                }
                throw "List is not specified.";
            };
            ListViewBase.prototype.renderListDataAsStream = function (token) {
                var self = this;
                var deferred = self._app.$.Deferred();
                try {
                    var query = self.getList();
                    query.concat("/RenderListDataAsStream");
                    if (!$pnp.util.stringIsNullOrEmpty(token)) {
                        self.addToken(query.query, token);
                    }
                    else {
                        if (!$pnp.util.stringIsNullOrEmpty(self._options.viewId)) {
                            query.query.add("View", encodeURIComponent(self._options.viewId));
                        }
                        if (!$pnp.util.stringIsNullOrEmpty(self._options.orderBy)) {
                            query.query.add("SortField", encodeURIComponent(self._options.orderBy));
                        }
                        if (!$pnp.util.stringIsNullOrEmpty(self._options.sortAsc)) {
                            query.query.add("SortDir", self._options.sortAsc ? "Asc" : "Desc");
                        }
                    }
                    self.addFilterQuery(query.query);
                    var viewXml = self._options.viewXml;
                    if (self._options.queryBuilder) {
                        viewXml = self._options.queryBuilder.toString();
                    }
                    var url = query.toUrlAndQuery();
                    var parameters = { "__metadata": { "type": "SP.RenderListDataParameters" }, "RenderOptions": self._options.renderOptions };
                    if (!$pnp.util.stringIsNullOrEmpty(viewXml)) {
                        parameters.ViewXml = viewXml;
                    }
                    if (!$pnp.util.stringIsNullOrEmpty(self._options.rootFolder)) {
                        parameters.FolderServerRelativeUrl = self._options.rootFolder;
                    }
                    var postBody = JSON.stringify({ "parameters": parameters });
                    var executor = new SP.RequestExecutor(self._app.appWebUrl);
                    executor.executeAsync({
                        url: url,
                        method: "POST",
                        headers: {
                            "accept": "application/json;odata=verbose",
                            "content-Type": "application/json;odata=verbose"
                        },
                        body: postBody,
                        success: function (data) {
                            var result = null;
                            if (data.body) {
                                result = JSON.parse(data.body);
                            }
                            deferred.resolve(result);
                        },
                        error: function (data, errorCode, errorMessage) {
                            if (data.body) {
                                try {
                                    var error = JSON.parse(data.body);
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
                }
                catch (e) {
                    self._app.$(self._app).trigger("app-error", [e]);
                    deferred.reject(e);
                }
                return deferred.promise();
            };
            ListViewBase.prototype.renderListData = function (token) {
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
                    }
                    else {
                        if (!$pnp.util.stringIsNullOrEmpty(self._options.viewId)) {
                            query.query.add("View", encodeURIComponent(self._options.viewId));
                        }
                        if (!$pnp.util.stringIsNullOrEmpty(self._options.orderBy)) {
                            query.query.add("SortField", encodeURIComponent(self._options.orderBy));
                        }
                        if (!$pnp.util.stringIsNullOrEmpty(self._options.sortAsc)) {
                            query.query.add("SortDir", self._options.sortAsc ? "Asc" : "Desc");
                        }
                    }
                    self.addFilterQuery(query.query);
                    var url = query.toUrlAndQuery();
                    var executor = new SP.RequestExecutor(self._app.appWebUrl);
                    executor.executeAsync({
                        url: url,
                        method: "POST",
                        headers: {
                            "accept": "application/json;odata=verbose",
                            "content-Type": "application/json;odata=verbose"
                        },
                        success: function (data) {
                            var result = null;
                            if (data.body) {
                                result = JSON.parse(JSON.parse(data.body).d.RenderListData);
                            }
                            deferred.resolve({ ListData: result });
                        },
                        error: function (data, errorCode, errorMessage) {
                            if (data.body) {
                                try {
                                    var error = JSON.parse(data.body);
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
                }
                catch (e) {
                    self._app.$(self._app).trigger("app-error", [e]);
                    deferred.reject(e);
                }
                return deferred.promise();
            };
            ListViewBase.prototype.getItems = function () {
                var self = this;
                var deferred = self._app.$.Deferred();
                try {
                    var query = self.getList();
                    query.concat("/GetItems");
                    if ($pnp.util.isArray(self._options.expands)) {
                        query = query.expand(self._options.expands);
                    }
                    var url = query.toUrlAndQuery();
                    var viewXml = self._options.viewXml;
                    if (self._options.queryBuilder) {
                        viewXml = self._options.queryBuilder.toString();
                    }
                    var postBody = JSON.stringify({ "query": { "__metadata": { "type": "SP.CamlQuery" }, "ViewXml": viewXml } });
                    var executor = new SP.RequestExecutor(self._app.appWebUrl);
                    executor.executeAsync({
                        url: url,
                        method: "POST",
                        body: postBody,
                        headers: {
                            "accept": "application/json;odata=verbose",
                            "content-Type": "application/json;odata=verbose"
                        },
                        success: function (data) {
                            var listData = null;
                            if (data.body) {
                                var d = JSON.parse(data.body).d;
                                listData = { Row: d.results, NextHref: null, PrevHref: null };
                            }
                            deferred.resolve({ ListData: listData });
                        },
                        error: function (data, errorCode, errorMessage) {
                            if (data.body) {
                                try {
                                    var error = JSON.parse(data.body);
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
                }
                catch (e) {
                    self._app.$(self._app).trigger("app-error", [e]);
                    deferred.reject(e);
                }
                return deferred.promise();
            };
            ListViewBase.prototype.getItemsREST = function (token) {
                var self = this;
                var deferred = self._app.$.Deferred();
                try {
                    var query = self.getList().items;
                    if (!$pnp.util.stringIsNullOrEmpty(self._options.orderBy)) {
                        query = query.orderBy(self._options.orderBy, self._options.sortAsc);
                    }
                    if (self._options.limit > 0) {
                        query = query.top(self._options.limit);
                    }
                    if ($pnp.util.isArray(self._options.expands)) {
                        query = query.expand(self._options.expands);
                    }
                    if (!$pnp.util.stringIsNullOrEmpty(token)) {
                        query.query.add("$skiptoken", encodeURIComponent(token));
                    }
                    var url = query.toUrlAndQuery();
                    var executor = new SP.RequestExecutor(self._app.appWebUrl);
                    executor.executeAsync({
                        url: url,
                        method: "GET",
                        headers: {
                            "accept": "application/json;odata=verbose",
                            "content-Type": "application/json;odata=verbose"
                        },
                        success: function (data) {
                            var listData = null;
                            if (data.body) {
                                var d = JSON.parse(data.body).d;
                                listData = { Row: d.results, NextHref: null, PrevHref: null };
                                listData.NextHref = self._app.getQueryParam(d["__next"], "$skiptoken");
                            }
                            deferred.resolve({ ListData: listData });
                        },
                        error: function (data, errorCode, errorMessage) {
                            if (data.body) {
                                try {
                                    var error = JSON.parse(data.body);
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
                }
                catch (e) {
                    self._app.$(self._app).trigger("app-error", [e]);
                    deferred.reject(e);
                }
                return deferred.promise();
            };
            ListViewBase.prototype.getItemsCSOM = function (token) {
                var self = this;
                var deferred = self._app.$.Deferred();
                try {
                    var context = new SP.ClientContext(self._app.appWebUrl);
                    var factory = new SP.ProxyWebRequestExecutorFactory(self._app.appWebUrl);
                    context.set_webRequestExecutorFactory(factory);
                    var appContextSite = new SP.AppContextSite(context, self._app.hostWebUrl);
                    var web = appContextSite.get_web();
                    var list = null;
                    if (!$pnp.util.stringIsNullOrEmpty(self._options.listId)) {
                        list = web.get_lists().getById(self._options.listId);
                    }
                    else if (!$pnp.util.stringIsNullOrEmpty(self._options.listUrl)) {
                        list = web.getList(self._options.listUrl);
                    }
                    else if (!$pnp.util.stringIsNullOrEmpty(self._options.listTitle)) {
                        list = web.get_lists().getByTitle(self._options.listTitle);
                    }
                    else {
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
                    context.executeQueryAsync(function () {
                        var listData = { Row: self._app.$.map(items.get_data(), function (item) { return item.get_fieldValues(); }), NextHref: null, PrevHref: null, FirstRow: 0, LastRow: 0 };
                        var position = items.get_listItemCollectionPosition();
                        if (position) {
                            listData.NextHref = position.get_pagingInfo();
                        }
                        deferred.resolve({ ListData: listData });
                    }, function (sender, args) {
                        self._app.$(self._app).trigger("app-error", [args.get_message(), args.get_stackTrace()]);
                        deferred.reject(sender, args);
                    });
                }
                catch (e) {
                    self._app.$(self._app).trigger("app-error", [e]);
                    deferred.reject(e);
                }
                return deferred.promise();
            };
            ListViewBase.prototype.getListItems = function (token) {
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
            };
            ListViewBase.prototype.get_app = function () {
                return this._app;
            };
            ListViewBase.prototype.get_options = function () {
                return this._options;
            };
            ListViewBase.prototype.render = function () {
                this._app.$(self).trigger("model-render");
            };
            return ListViewBase;
        }());
        Module.ListViewBase = ListViewBase;
        var ListsViewBase = (function () {
            function ListsViewBase(app, options) {
                if (!app) {
                    throw "App must be specified for ListView!";
                }
                this._app = app;
                this._options = options;
            }
            ListsViewBase.prototype.getLists = function () {
                var self = this;
                var deferred = self._app.$.Deferred();
                var url = $pnp.sp.crossDomainWeb(self._app.appWebUrl, self._app.hostWebUrl).lists.select("Id", "Title", "BaseType", "ItemCount", "Description", "Hidden", "EffectiveBasePermissions").toUrlAndQuery();
                var executor = new SP.RequestExecutor(self._app.appWebUrl);
                executor.executeAsync({
                    url: url,
                    method: "GET",
                    headers: {
                        "accept": "application/json;odata=verbose",
                        "content-Type": "application/json;odata=verbose"
                    },
                    success: function (data) {
                        var lists = JSON.parse(data.body).d.results;
                        deferred.resolve(lists);
                    },
                    error: function (data, errorCode, errorMessage) {
                        if (data.body) {
                            try {
                                var error = JSON.parse(data.body);
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
            };
            ListsViewBase.prototype.getList = function (listId) {
                var self = this;
                var deferred = self._app.$.Deferred();
                var url = $pnp.sp.crossDomainWeb(self._app.appWebUrl, self._app.hostWebUrl).lists.getById(listId).select("Id", "Title", "BaseType", "ItemCount", "Description", "Hidden", "EffectiveBasePermissions").toUrlAndQuery();
                var executor = new SP.RequestExecutor(self._app.appWebUrl);
                executor.executeAsync({
                    url: url,
                    method: "GET",
                    headers: {
                        "accept": "application/json;odata=verbose",
                        "content-Type": "application/json;odata=verbose"
                    },
                    success: function (data) {
                        var list = JSON.parse(data.body).d;
                        deferred.resolve(list);
                    },
                    error: function (data, errorCode, errorMessage) {
                        if (data.body) {
                            try {
                                var error = JSON.parse(data.body);
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
            };
            ListsViewBase.prototype.updateList = function (listId, properties, digestValue) {
                var self = this;
                var deferred = self._app.$.Deferred();
                var url = $pnp.sp.crossDomainWeb(self._app.appWebUrl, self._app.hostWebUrl).lists.getById(listId).toUrlAndQuery();
                var body = JSON.stringify($pnp.util.extend({
                    "__metadata": { "type": "SP.List" }
                }, properties));
                var executor = new SP.RequestExecutor(self._app.appWebUrl);
                executor.executeAsync({
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
                    success: function (data) {
                        deferred.resolve();
                    },
                    error: function (data, errorCode, errorMessage) {
                        if (data.body) {
                            try {
                                var error = JSON.parse(data.body);
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
            };
            ListsViewBase.prototype.get_options = function () {
                return this._options;
            };
            ListsViewBase.prototype.get_app = function () {
                return this._app;
            };
            ListsViewBase.prototype.render = function () {
                this._app.$(self).trigger("model-render");
            };
            return ListsViewBase;
        }());
        Module.ListsViewBase = ListsViewBase;
    })(Module = App.Module || (App.Module = {}));
})(App = exports.App || (exports.App = {}));

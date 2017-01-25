/// <reference path="typings/angularjs/angular.d.ts" />
/// <reference path="typings/sharepoint/SharePoint.d.ts" />
/// <reference path="typings/sharepoint/pnp.d.ts" />
/// <reference path="typings/microsoft-ajax/microsoft.ajax.d.ts" />

import * as $pnp from "pnp";
import * as $ from "jquery";
import * as angular from "angular";
import * as app from "./app.module";

namespace Angular {

    "use strict";
    export class App extends app.App.AppBase {

        public $angular: angular.IAngularStatic;
        public ngModule: ng.IModule;
        static SharePointAppName = "SharePointApp";
        static SPServiceName = "SPService";

        public init(preloadedScripts: any[]) {
            var self = this;
            let $angular = preloadedScripts["angular"];
            self.$angular = $angular;
            if (!self.$angular) {
                throw "angular is not loaded!";
            }
            self.ngModule = self.$angular.module(App.SharePointAppName, [
                //'ngSanitize',
                "officeuifabric.core",
                "officeuifabric.components"
            ]).service(App.SPServiceName, function ($http, $q) {
                this.getFormDigest = () => {
                    var deferred = self.$.Deferred();
                    var url = $pnp.util.combinePaths(self.appWebUrl, "_api/contextinfo");
                    var executor = new SP.RequestExecutor(self.appWebUrl);
                    executor.executeAsync(<SP.RequestInfo>{
                        url: url,
                        method: "POST",
                        headers: {
                            "accept": "application/json;odata=verbose",
                            "content-Type": "application/json;odata=verbose"
                        },
                        success: (data) => {
                            var formDigestValue = JSON.parse(<string>data.body).d.GetContextWebInformation.FormDigestValue;
                            deferred.resolve(formDigestValue);
                        },
                        error: (data, errorCode, errorMessage) => {
                            if (data.body) {
                                try {
                                    var error = JSON.parse(<string>data.body);
                                    if (error && error.error) {
                                        errorMessage = error.error.message.value;
                                    }
                                } catch (e) {
                                }
                            }
                            self.$(self).trigger("app-error", [errorMessage]);
                            deferred.reject(data, errorCode, errorMessage);
                        }
                    });
                    return deferred.promise();
                };
            }).filter("unsafe", function ($sce) {
                return $sce.trustAsHtml;
            }).directive("compile", [
                "$compile", ($compile) => {
                    return (scope, element, attrs) => {
                        scope.$watch(
                            (scope) => {
                                // watch the 'compile' expression for changes
                                return scope.$eval((<any>attrs).compile);
                            },
                            (value) => {
                                // when the 'compile' expression changes
                                // assign it into the current DOM
                                element.html(value);
                                // compile the new DOM and link it to the current
                                // scope.
                                // NOTE: we only compile .childNodes so that
                                // we don't get into infinite loop compiling ourselves
                                $compile(element.contents())(scope);
                            }
                        );
                    };
                }
            ]);
            super.init(preloadedScripts);
        }

        public render(modules: app.App.IModule[]) {
            var self = this;
            self.$(self).on("app-render", () => {
                self.$angular.element(() => {
                    self.$angular.bootstrap(document, [App.SharePointAppName]);
                });
            });
            super.render(modules);
        }

        public get_ListView(options: app.App.Module.IListViewOptions): App.Module.ListView {
            var self = this;
            //if (!self.is_initialized()) {
            //    throw "App is not initialized!";
            //}
            return new App.Module.ListView(self, options);
        }

        public get_ListsView(options: App.Module.IListsViewOptions): App.Module.ListsView {
            var self = this;
            //if (!self.is_initialized()) {
            //    throw "App is not initialized!";
            //}
            return new App.Module.ListsView(self, options);
        }
    }

    export module App.Module {

        interface IListViewOptions extends app.App.Module.IListViewOptions {
            delay?: number;
        }

        export class ListView extends app.App.Module.ListViewBase {

            constructor(app: App, options: IListViewOptions) {
                if (!options.delay) {
                    options.delay = 1000;
                }
                super(<app.App.AppBase>app, options);
            }

            private getEntity(listItem) {
                var permMask = listItem["PermMask"];
                var permissions = this.get_app().get_BasePermissions(permMask);
                var $permissions = {
                    edit: permissions.has(SP.PermissionKind.editListItems),
                    delete: permissions.has(SP.PermissionKind.deleteListItems)
                };
                var $events = { menuOpened: false, isSelected: false };
                return { $data: listItem, $events: $events, $permissions: $permissions };
            }

            public render() {
                var self = this;
                var allTokens = [];
                var app = <App>this.get_app();
                var options = <IListViewOptions>self.get_options();
                app.ngModule.factory("ListViewFactory", ($q, $http) => {
                    var factory = {} as IListViewFactory;
                    factory.listItems = [];
                    factory.getToken = (offset = 0) => {
                        var token = null;
                        if (offset < 0) {
                            token = factory.$prevToken;
                            var skipNext = 1 - offset;
                            if (allTokens.length > skipNext) {
                                token = allTokens[(allTokens.length - 1) - skipNext];
                            }
                        } else if (offset > 0) {
                            token = factory.$nextToken;
                            if (offset > 0) {
                                var index = allTokens.indexOf(factory.$nextToken);
                                if ((index + offset - 1) < allTokens.length) {
                                    token = allTokens[index + offset - 1];
                                }
                            }
                        } else {
                            token = factory.$currentToken;
                        }
                        return token;
                    };
                    factory.getListItems = (token?: string) => {
                        var deferred = $q.defer();
                        self.getListItems(token).then((data: any) => {
                            //factory.listItems.splice(0, factory.listItems.length);
                            factory.listItems = [];
                            if (data.ListData) {
                                app.$.each(data.ListData.Row, ((i, listItem) => {
                                    factory.listItems.push(self.getEntity(listItem));
                                }));
                                factory.$nextToken = data.ListData.NextHref;
                                //factory.$prevToken = data.ListData.PrevHref;
                                factory.$currentToken = token;
                                factory.$first = data.ListData.FirstRow ? Number(data.ListData.FirstRow) : 0;
                                factory.$last = data.ListData.LastRow ? Number(data.ListData.LastRow) : 0;
                            }
                            deferred.resolve();
                        }, deferred.reject);
                        return deferred.promise;
                    };
                    return factory;
                });

                app.ngModule.controller(options.controllerName, [
                    "$scope", "ListViewFactory", App.SPServiceName, function ($scope: ng.IScope, factory: IListViewFactory, service: app.App.ISPService) {
                        (<any>$scope).loading = true;
                        (<any>$scope).listItems = [];
                        factory.getListItems().then(() => {
                            allTokens.push(factory.$nextToken);
                            (<any>$scope).listItems = app.$angular.copy(factory.listItems);
                            (<any>$scope).selection.pager.first = factory.$first > 0 ? factory.$first : ((Math.max(allTokens.length, 1) - 1) * options.limit + Math.min(1, factory.listItems.length));
                            (<any>$scope).selection.pager.last = factory.$last > 0 ? factory.$last : ((<any>$scope).selection.pager.first + factory.listItems.length - (factory.listItems.length > 0 ? 1 : 0));
                            (<any>$scope).selection.pager.nextEnabled = !$pnp.util.stringIsNullOrEmpty(factory.$nextToken);
                            (<any>$scope).loading = false;
                        });
                        (<any>$scope).selection = {
                            commandBar: {
                                searchTerm: null,
                                selectionText: "",
                                createEnabled: false,
                                viewEnabled: false,
                                deleteEnabled: false,
                                view: function (listItem) {
                                    if (!listItem) {
                                        var selectedItems = (<any>$scope).table.selectedItems;
                                        listItem = app.$(selectedItems).get(0);
                                    }
                                    if (listItem) {
                                    }
                                },
                                delete: function (listItem) {
                                    if (!listItem) {
                                        var selectedItems = (<any>$scope).table.selectedItems;

                                    }
                                },
                                clearSelection: function () {
                                    var selectedItems = (<any>$scope).table.selectedItems;
                                    if (selectedItems.length > 0) {
                                        app.$.each((<any>$scope).table.rows, (i, item) => {
                                            if (item.selected) {
                                                item.selected = false;
                                            }
                                        });
                                    }
                                }
                            },
                            openMenu: (listItem) => {
                                if (listItem) {
                                    if (!listItem.$events.menuOpened) {
                                        app.$.each((<any>$scope).listItems, ((i, listItem) => {
                                            listItem.$events.menuOpened = false;
                                        }));
                                    }
                                    listItem.$events.menuOpened = !listItem.$events.menuOpened;
                                }
                            },
                            pager: {
                                first: 0,
                                last: 0,
                                prevEnabled: false,
                                nextEnabled: false,
                                refresh: () => {
                                    if ((<any>$scope).loading) { return; }
                                    var token = options.appendRows === true ? null : factory.getToken(0);
                                    (<any>$scope).table.rows.splice(0, (<any>$scope).table.rows.length);
                                    (<any>$scope).selection.pager.prevEnabled = false;
                                    (<any>$scope).selection.pager.nextEnabled = false;
                                    factory.getListItems(token).then(() => {
                                        //(<any>$scope).selection.commandBar.clearSelection();                                       
                                        if (!token) {
                                            allTokens = [];
                                        } else {
                                            allTokens.pop();
                                        }
                                        allTokens.push(factory.$nextToken);
                                        var prevToken = factory.getToken(-1);
                                        if ($pnp.util.stringIsNullOrEmpty(token) || !options.appendRows) {
                                            (<any>$scope).listItems = app.$angular.copy(factory.listItems);
                                        } else {
                                            (<any>$scope).listItems = app.$angular.copy([].concat((<any>$scope).listItems).concat(factory.listItems));
                                        }
                                        (<any>$scope).selection.pager.first = factory.$first > 0 ? factory.$first : ((Math.max(allTokens.length, 1) - 1) * options.limit + Math.min(1, factory.listItems.length));
                                        (<any>$scope).selection.pager.last = factory.$last > 0 ? factory.$last : ((<any>$scope).selection.pager.first + factory.listItems.length - (factory.listItems.length > 0 ? 1 : 0));
                                        (<any>$scope).selection.pager.nextEnabled = !$pnp.util.stringIsNullOrEmpty(factory.$nextToken);
                                        (<any>$scope).selection.pager.prevEnabled = options.appendRows !== true && ($pnp.util.stringIsNullOrEmpty(factory.$prevToken) ? !$pnp.util.stringIsNullOrEmpty(factory.getToken(0)) : true);
                                    });
                                },
                                next: (offset = 1) => {
                                    if ((<any>$scope).loading) { return; }
                                    if (!(<any>$scope).selection.pager.nextEnabled) return;
                                    (<any>$scope).table.rows.splice(0, (<any>$scope).table.rows.length);
                                    var token = factory.getToken(offset);
                                    (<any>$scope).selection.pager.prevEnabled = false;
                                    (<any>$scope).selection.pager.nextEnabled = false;
                                    factory.getListItems(token).then(() => {
                                        //(<any>$scope).selection.commandBar.clearSelection();
                                        if ($pnp.util.stringIsNullOrEmpty(token)) {
                                            allTokens = [];
                                        }
                                        allTokens.push(factory.$nextToken);
                                        var prevToken = factory.getToken(-1);
                                        if ($pnp.util.stringIsNullOrEmpty(token) || !options.appendRows) {
                                            (<any>$scope).listItems = app.$angular.copy(factory.listItems);
                                        } else {
                                            (<any>$scope).listItems = app.$angular.copy([].concat((<any>$scope).listItems).concat(factory.listItems));
                                        }
                                        (<any>$scope).selection.pager.first = factory.$first > 0 ? factory.$first : ((Math.max(allTokens.length, 1) - 1) * options.limit + Math.min(1, factory.listItems.length));
                                        (<any>$scope).selection.pager.last = factory.$last > 0 ? factory.$last : ((<any>$scope).selection.pager.first + factory.listItems.length - (factory.listItems.length > 0 ? 1 : 0));
                                        (<any>$scope).selection.pager.nextEnabled = !$pnp.util.stringIsNullOrEmpty(factory.$nextToken);
                                        (<any>$scope).selection.pager.prevEnabled = options.appendRows !== true && ($pnp.util.stringIsNullOrEmpty(factory.$prevToken) ? !$pnp.util.stringIsNullOrEmpty(factory.getToken(0)) : true);
                                    });
                                },
                                prev: (offset = 1) => {
                                    if ((<any>$scope).loading) return;
                                    if (!(<any>$scope).selection.pager.prevEnabled) return;
                                    (<any>$scope).table.rows.splice(0, (<any>$scope).table.rows.length);
                                    offset = Math.min(-1, -offset);
                                    var token = options.appendRows === true ? null : factory.getToken(offset);
                                    (<any>$scope).selection.pager.prevEnabled = false;
                                    (<any>$scope).selection.pager.nextEnabled = false;
                                    factory.getListItems(token).then(() => {
                                        //(<any>$scope).selection.commandBar.clearSelection();                                       
                                        if ($pnp.util.stringIsNullOrEmpty(token)) {
                                            allTokens = [];
                                        }
                                        else {
                                            var skipNext = 1 - offset;
                                            while (skipNext > 0) {
                                                allTokens.pop();
                                                skipNext--;
                                            }
                                        }
                                        allTokens.push(factory.$nextToken);
                                        var prevToken = factory.getToken(-1);
                                        if ($pnp.util.stringIsNullOrEmpty(token) || !options.appendRows) {
                                            (<any>$scope).listItems = app.$angular.copy(factory.listItems);
                                        } else {
                                            (<any>$scope).listItems = app.$angular.copy([].concat((<any>$scope).listItems).concat(factory.listItems));
                                        }
                                        (<any>$scope).selection.pager.first = factory.$first > 0 ? factory.$first : ((Math.max(allTokens.length, 1) - 1) * options.limit + Math.min(1, factory.listItems.length));
                                        (<any>$scope).selection.pager.last = factory.$last > 0 ? factory.$last : ((<any>$scope).selection.pager.first + factory.listItems.length - (factory.listItems.length > 0 ? 1 : 0));
                                        (<any>$scope).selection.pager.nextEnabled = !$pnp.util.stringIsNullOrEmpty(factory.$nextToken);
                                        (<any>$scope).selection.pager.prevEnabled = options.appendRows !== true && ($pnp.util.stringIsNullOrEmpty(factory.$prevToken) ? !$pnp.util.stringIsNullOrEmpty(factory.getToken(0)) : true);
                                    });
                                },
                            }
                        };
                        $scope.$watch("table.selectedItems", (newValue: Array<any>, oldValue: Array<any>) => {
                            (<any>$scope).selection.commandBar.viewEnabled = newValue.length === 1;
                            (<any>$scope).selection.commandBar.deleteEnabled = newValue.length > 0;
                            (<any>$scope).selection.commandBar.selectionText = newValue.length > 0 ? newValue.length + " selected" : null;
                        }, true);
                        $scope.$watch("selection.commandBar.searchTerm", (newValue: string, oldValue: string) => {
                            (<any>$scope).table.rows.splice(0, (<any>$scope).table.rows.length);
                            app.delay(() => {
                                $scope.$apply(() => {
                                    var listItems;
                                    if (newValue !== oldValue) {
                                        options.queryBuilder.clear();
                                        if (newValue) {
                                            var filters = new Array<app.Caml.ICamlFilter>();
                                            //filters.push({ field: "Title", fieldType: SP.FieldType.text, value: newValue, operation: 7 });
                                            app.$(self).trigger("search-item", [filters, newValue, $scope, factory]);
                                            options.queryBuilder.appendAndWithAny.apply(options.queryBuilder, filters);
                                        }
                                        (<any>$scope).selection.pager.prevEnabled = false;
                                        (<any>$scope).selection.pager.nextEnabled = false;
                                        factory.getListItems(null).then(() => {
                                            allTokens = [];
                                            allTokens.push(factory.$nextToken);
                                            (<any>$scope).listItems = app.$angular.copy(factory.listItems);
                                            (<any>$scope).selection.pager.first = factory.$first > 0 ? factory.$first : ((Math.max(allTokens.length, 1) - 1) * options.limit + Math.min(1, factory.listItems.length));
                                            (<any>$scope).selection.pager.last = factory.$last > 0 ? factory.$last : ((<any>$scope).selection.pager.first + factory.listItems.length - (factory.listItems.length > 0 ? 1 : 0));
                                            (<any>$scope).selection.pager.nextEnabled = !$pnp.util.stringIsNullOrEmpty(factory.$nextToken);
                                        });
                                    }
                                });
                            }, options.delay);
                        }, false);
                        app.$(self).trigger("model-render", [$scope, factory]);
                    }
                ]);
            }
        }

        interface IListViewFactory {
            listItems: Array<any>;
            $nextToken: string;
            $prevToken: string;
            $currentToken: string;
            $first: number,
            $last: number,
            getToken(offset?: number);
            getListItems(token?: string);
        }

        interface IListsViewFactory {
            lists: Array<any>;
            getLists();
            updateList(list, service: app.App.ISPService);
        }

        export interface IListsViewOptions extends app.App.Module.IListsViewOptions {
            delay: number;
        }

        export class ListsView extends app.App.Module.ListsViewBase {

            constructor(app: App, options: IListsViewOptions) {
                if (!options.delay) {
                    options.delay = 1000;
                }
                super(<app.App.AppBase>app, options);
            }

            private getEntity(list) {
                switch (list.BaseType) {
                    case 0:
                        list.Type = "List";
                        break;
                    case 1:
                        list.Type = "Document Library";
                        break;
                    case 2:
                        list.Type = "Unused";
                        break;
                    case 3:
                        list.Type = "Discussion Board";
                        break;
                    case 4:
                        list.Type = "Survey";
                        break;
                    case 5:
                        list.Type = "Issue";
                        break;
                    default:
                        list.Type = "None";
                        break;
                }
                var permissions = new SP.BasePermissions();
                permissions.initPropertiesFromJson(list["EffectiveBasePermissions"]);
                var $permissions = {
                    manage: permissions.has(SP.PermissionKind.manageLists)
                };
                var $events = { menuOpened: false, isSelected: false };
                return { $data: list, $events: $events, $permissions: $permissions };
            }

            public render() {
                var self = this;
                var app = <App>this.get_app();
                var options = <IListsViewOptions>self.get_options();
                app.ngModule.factory("ListsViewFactory", ($q, $http) => {
                    var factory = {} as IListsViewFactory;
                    factory.lists = [];
                    factory.getLists = () => {
                        var deferred = $q.defer();
                        self.getLists().then((data: Array<any>) => {
                            factory.lists.splice(0, factory.lists.length);
                            app.$.each(data, (function (i, list) {
                                if (!list.Hidden) {
                                    var entity = self.getEntity(list);
                                    factory.lists.push(entity);
                                }
                            }));
                            deferred.resolve(factory.lists);
                        }, deferred.reject);
                        return deferred.promise;
                    };
                    factory.updateList = (list, service: app.App.ISPService) => {
                        var deferred = $q.defer();
                        service.getFormDigest().done((digestValue: string) => {
                            var properties = {
                                "Title": list.Title,
                                "Description": list.Description
                            };
                            self.updateList(list.Id, properties, digestValue).done(() => {
                                self.getList(list.Id).done((data) => {
                                    deferred.resolve(data);
                                });
                            }).fail(deferred.reject);
                        }).fail(deferred.reject);
                        return deferred.promise;
                    };
                    return factory;
                });
                var deferred = app.$.Deferred();
                app.ngModule.controller(options.controllerName, ["$scope", "ListsViewFactory", App.SPServiceName, function ($scope: ng.IScope, factory: IListsViewFactory, service: app.App.ISPService) {
                    (<any>$scope).loading = true;
                    (<any>$scope).lists = [];
                    factory.getLists().then(() => {
                        (<any>$scope).lists = app.$angular.copy(factory.lists);
                        (<any>$scope).loading = false;
                        deferred.resolve();
                    }, () => {
                        (<any>$scope).loading = false;
                        deferred.reject();
                    });
                    (<any>$scope).selection = {
                        settings: {
                            opened: false,
                            data: [],
                            editMode: false,
                            onEdit: () => {
                                (<any>$scope).selection.settings.editMode = true;
                            },
                            onSave: () => {
                                return factory.updateList((<any>$scope).selection.settings.data, service).then((data) => {
                                    (<any>$scope).selection.settings.data = app.$.extend(true, {}, data);
                                    (<any>$scope).selection.settings.editMode = false;
                                    var entity = self.getEntity(data);
                                    $.each((<any>$scope).lists, (i, list) => {
                                        if (list.$data.Id === entity.$data.Id) {
                                            list.$data = entity.$data;
                                            list.$permissions = entity.$permissions;
                                        }
                                    });
                                });
                            }
                        },
                        commandBar: {
                            searchTerm: null,
                            selectionText: "",
                            createEnabled: false,
                            viewEnabled: false,
                            deleteEnabled: false,
                            settingsEnabled: false,
                            openSettings: function (list) {
                                if (!list) {
                                    var selectedItems = (<any>$scope).table.selectedItems;
                                    list = app.$(selectedItems).get(0);
                                }
                                if (list) {
                                    if (!(<any>$scope).selection.settings.opened) {
                                        (<any>$scope).selection.settings.editMode = false;
                                        (<any>$scope).selection.settings.canEdit = list.$permissions.manage;
                                        (<any>$scope).selection.settings.data.Id = list.$data.Id;
                                        (<any>$scope).selection.settings.data.Title = list.$data.Title;
                                        (<any>$scope).selection.settings.data.Description = list.$data.Description;
                                    } else {
                                        (<any>$scope).selection.settings.data = [];
                                    }
                                    (<any>$scope).selection.settings.opened = !(<any>$scope).selection.settings.opened;
                                }
                            },
                            view: function (list) {
                                if (!list) {
                                    var selectedItems = (<any>$scope).table.selectedItems;
                                    list = app.$(selectedItems).get(0);
                                }
                                if (list) {
                                    window.location.href = "/Home/List?ListId=" + list.$data.Id + "&SPHostUrl=" + decodeURIComponent(app.hostWebUrl) + "&SPAppWebUrl=" + decodeURIComponent(app.appWebUrl);
                                }
                            },
                            delete: function (list) {
                                if (!list) {
                                    var selectedItems = (<any>$scope).table.selectedItems;

                                }
                            },
                            clearSelection: function () {
                                var selectedItems = (<any>$scope).table.selectedItems;
                                if (selectedItems.length > 0) {
                                    app.$.each((<any>$scope).table.rows, (i, item) => {
                                        if (item.selected) {
                                            item.selected = false;
                                        }
                                    });
                                }
                            }
                        },
                        openMenu: (list) => {
                            if (list) {
                                if (!list.$events.menuOpened) {
                                    app.$.each((<any>$scope).lists, ((i, list) => {
                                        list.$events.menuOpened = false;
                                    }));
                                }
                                list.$events.menuOpened = !list.$events.menuOpened;
                            }
                        }
                    };
                    $scope.$watch("table.selectedItems", (newValue: Array<any>, oldValue: Array<any>) => {
                        (<any>$scope).selection.commandBar.viewEnabled = newValue.length === 1;
                        (<any>$scope).selection.commandBar.deleteEnabled = newValue.length > 0;
                        (<any>$scope).selection.commandBar.settingsEnabled = newValue.length === 1;
                        (<any>$scope).selection.commandBar.selectionText = newValue.length > 0 ? newValue.length + " selected" : null;
                    }, true);
                    $scope.$watch("selection.commandBar.searchTerm", function (newValue: string, oldValue: string) {
                        (<any>$scope).table.rows.splice(0, (<any>$scope).table.rows.length);
                        app.delay(() => {
                            $scope.$apply(function () {
                                var lists;
                                if (newValue && newValue !== oldValue) {
                                    lists = [];
                                    app.$.each(factory.lists, (i, list) => {
                                        if ((<any>list).$data && new RegExp(newValue, "i").test((<any>list).$data.Title)) {
                                            lists.push(list);
                                        }
                                    });
                                } else {
                                    lists = factory.lists;
                                }
                                (<any>$scope).lists = app.$angular.copy(lists);
                            });
                        }, options.delay);
                    }, false);
                }]);
                return deferred.promise();
            }
        }
    }
}

export = new Angular.App();
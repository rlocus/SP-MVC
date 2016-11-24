require.config({
    "baseUrl": "scripts",
    "paths": {
        "app": 'app',
        //"fetch": 'fetch',
        //"es6-promise": 'es6-promise.min',
        "pnp": 'pnp.min',
        "angular": 'angular.min',
        "ngOfficeUiFabric": 'ngOfficeUiFabric.min'
    },
    shim: {
        jquery: {
            exports: ['jQuery', '$']
        },
        angular: {
            exports: 'angular'
        },
        ngOfficeUiFabric: ['angular'],
        app: ['jquery', 'pnp', 'angular', 'ngOfficeUiFabric'],
    }
});

if (typeof window.jQuery == "undefined") {
} else {
    if (typeof define === "function" && define.amd) {
        define('jquery', [], function () {
            return window.jQuery; /* window.jQuery.noConflict();*/
        });
    }
}

// you can add additional requirements in here but you would need to manually add them to the preloaded modules object
// we are also showing how to include poly-fills for fetch and es6 promises if needed.
require(["jquery", "angular", "app"], function ($, angular, app) {
    //$(function () {
        app.init({
            "jquery": $,
            "angular": angular
        });
    //});
});
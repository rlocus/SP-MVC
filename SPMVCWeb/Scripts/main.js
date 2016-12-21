require.config({
    "baseUrl": "/Scripts",
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
    app.init({
        "jquery": $,
        "angular": angular
    });
    //$('body').fadeIn();
    app.ensureScript(app.scriptBase + "/MicrosoftAjax.js").then(function () {
        app.ensureScript(app.scriptBase + "/SP.Runtime.js").then(function () {
            app.ensureScript(app.scriptBase + "/SP.js").then(function () {
                //Execute the correct script based on the isDialog
                //Load the SP.UI.Controls.js file to render the App Chrome
                app.ensureScript(app.scriptBase + "/SP.UI.Controls.js").then(function() {
                    //Set the chrome options for launching Help, Account, and Contact pages
                    var options = {
                        'siteUrl': app.hostWebUrl,
                        'siteTitle': "Back to Host Web",
                        'appTitle': "SP MVC",
                        'onCssLoaded': 'chromeLoaded()'
                    };
                    renderSPChrome(options);
                });
            });
        });
    });
    //function callback to render chrome after SP.UI.Controls.js loads
    function renderSPChrome(options) {
        $(function () {
            window.chromeLoaded = function () {
                $('body').fadeIn();
                $('.body-content').css("padding-top", $('#divSPChrome').height());
                $('body').trigger("chrome-loaded");
            };
            //Load the Chrome Control in the divSPChrome element of the page
            var chromeNavigation = new SP.UI.Controls.Navigation('divSPChrome', options);
            $("#chromeControl_bottomheader").remove();
            $('body').css('overflow', 'auto');
            chromeNavigation.setVisible(true);
        });
    }
});
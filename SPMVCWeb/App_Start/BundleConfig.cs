﻿using System.Web;
using System.Web.Optimization;

namespace SPMVCWeb
{
    public class BundleConfig
    {
        // For more information on bundling, visit http://go.microsoft.com/fwlink/?LinkId=301862
        public static void RegisterBundles(BundleCollection bundles)
        {
            //BundleTable.EnableOptimizations = true;
            bundles.Add(new ScriptBundle("~/bundles/jquery").Include(
                        "~/Scripts/jquery-{version}.js"));

            bundles.Add(new ScriptBundle("~/bundles/jqueryval").Include(
                        "~/Scripts/jquery.validate*"));

            // Use the development version of Modernizr to develop with and learn from. Then, when you're
            // ready for production, use the build tool at http://modernizr.com to pick only the tests you need.
            bundles.Add(new ScriptBundle("~/bundles/modernizr").Include(
                        "~/Scripts/modernizr-*"));

            bundles.Add(new ScriptBundle("~/bundles/bootstrap").Include(
                      "~/Scripts/bootstrap.js",
                      "~/Scripts/respond.js"));

            bundles.Add(new ScriptBundle("~/bundles/require").Include(
                      "~/Scripts/require.js"));

            bundles.Add(new ScriptBundle("~/bundles/main").Include(
                    "~/Scripts/camljs.js",
                    "~/Scripts/main.js"));

            bundles.Add(new ScriptBundle("~/bundles/ng-office").Include(
                    //"~/Scripts/es6-promise.min.js",
                    //"~/Scripts/fetch.js",
                    "~/Scripts/pnp.min.js",
                    "~/Scripts/ngOfficeUiFabric.min",
                    "~/Scripts/app.module.js",
                    "~/Scripts/app.angular.js"));

            bundles.Add(new ScriptBundle("~/bundles/angular").Include(
                 "~/Scripts/angular.min.js"));

            bundles.Add(new StyleBundle("~/Content/css").Include(
                      "~/Content/bootstrap.css",
                      "~/Content/site.css"));
            bundles.Add(new StyleBundle("~/Content/ng-office-css").Include(
                    "~/Content/fabric.min.css",
                    "~/Content/fabric.components.min.css",
                    "~/Content/angular-csp.css"));
        }
    }
}

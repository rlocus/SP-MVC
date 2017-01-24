using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using Microsoft.SharePoint.Client;
using System.Text.RegularExpressions;

namespace SPMVCWeb.Models
{
    public class ListInformation
    {
        public ListInformation(List list, View view)
        {
            if (list == null) throw new ArgumentNullException(nameof(list));
            if (view == null) throw new ArgumentNullException(nameof(view));
            if (list.IsPropertyAvailable("Id"))
            {
                Id = list.Id;
            }
            if (list.IsPropertyAvailable("Title"))
            {
                Title = list.Title;
            }
            if (list.IsPropertyAvailable("BaseTemplate"))
            {
                ListTemplate = list.BaseTemplate;
            }
            if (list.IsPropertyAvailable("BaseType"))
            {
                ListType = (int) list.BaseType;
            }
            if (view.IsPropertyAvailable("Id"))
            {
                ViewId = view.Id;
            }
            if (view.ViewFields.AreItemsAvailable && list.Fields.AreItemsAvailable)
            {
                List<Field> fields = new List<Field>();
                foreach (string fieldName in view.ViewFields)
                {
                    var field = list.Fields.FirstOrDefault(f => f.InternalName == fieldName);
                    if (field != null)
                    {
                        fields.Add(field);
                    }
                }
                Fields = fields.Select(FieldInformation.GetInformation).ToArray();
            }
            if (view.IsPropertyAvailable("ViewJoins"))
            {
                ViewJoins = HttpUtility.HtmlEncode(view.ViewJoins);
            }
            if (view.IsPropertyAvailable("ViewProjectedFields"))
            {
                ViewProjectedFields = HttpUtility.HtmlEncode(view.ViewProjectedFields);
            }
            if (view.IsPropertyAvailable("ViewQuery"))
            {
                ViewQuery = HttpUtility.HtmlEncode(view.ViewQuery);
            }
            if (view.IsPropertyAvailable("ListViewXml"))
            {
                ViewSchema = HttpUtility.HtmlEncode(view.ListViewXml);
            }
            if (view.IsPropertyAvailable("Paged"))
            {
                Paged = view.Paged;
            }
            if (view.IsPropertyAvailable("RowLimit"))
            {
                RowLimit = view.RowLimit;
            }
            if (view.IsPropertyAvailable("Title"))
            {
                ViewTitle = view.Title;
            }
            if (view.IsPropertyAvailable("ServerRelativeUrl"))
            {
                ViewUrl = view.ServerRelativeUrl;
            }

            if (view.IsPropertyAvailable("Scope"))
            {
                ViewScope = (int)view.Scope;
            }
            if (list.IsPropertyAvailable("ItemCount"))
            {
                ItemCount = list.ItemCount;
            }
            if (list.RootFolder.IsPropertyAvailable("ServerRelativeUrl"))
            {
                ListUrl = list.RootFolder.ServerRelativeUrl;
            }
        }

        public string Title { get; private set; }
        public Guid Id { get; private set; }
        public int ListTemplate { get; private set; }
        public int ListType { get; private set; }
        public Guid ViewId { get; private set; }
        public FieldInformation[] Fields { get; private set; }
        public string ViewJoins { get; private set; }
        public string ViewProjectedFields { get; private set; }
        public string ViewQuery { get; private set; }
        public bool Paged { get; private set; }
        public uint RowLimit { get; private set; }
        public string ViewTitle { get; private set; }
        public string ViewUrl { get; private set; }
        public string ListUrl { get; private set; }
        public string ViewSchema { get; private set; }
        public int ViewScope { get; private set; }
        public int ItemCount { get; private set; }
    }
}
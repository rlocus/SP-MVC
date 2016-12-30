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
            if (list == null) throw new ArgumentNullException("list");
            if (view == null) throw new ArgumentNullException("view");
            Id = list.Id;
            Title = list.Title;
            ViewId = view.Id;
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
            ViewJoins = HttpUtility.HtmlEncode(view.ViewJoins);
            ViewProjectedFields = HttpUtility.HtmlEncode(view.ViewProjectedFields);
            ViewQuery = HttpUtility.HtmlEncode(view.ViewQuery);
            ViewXml = HttpUtility.HtmlEncode(view.ListViewXml);
            Paged = view.Paged;
            RowLimit = view.RowLimit;
            ViewTitle = view.Title;
            ViewUrl = view.ServerRelativeUrl;
            ListUrl = list.RootFolder.ServerRelativeUrl;
        }

        public string Title { get; private set; }
        public Guid Id { get; private set; }
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
        public string ViewXml { get; private set; }
    }
}
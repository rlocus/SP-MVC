using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using Microsoft.SharePoint.Client;

namespace SPMVCWeb.Models
{
    public class FieldInformation
    {
        public FieldInformation(Field field)
        {
            Id = field.Id;
            Name = field.InternalName;
            Title = HttpUtility.HtmlEncode(field.Title);
            Description = HttpUtility.HtmlEncode(field.Description);
            IsReadOnly = field.ReadOnlyField;
            TypeKind = (uint)field.FieldTypeKind;
            TypeName = field.TypeAsString;
            Required = field.Required;
            Filterable = field.Filterable;
            Sortable = field.Sortable;
            DefaultValue = HttpUtility.HtmlEncode(field.DefaultValue);
        }

        public Guid Id { get; private set; }
        public string Name { get; private set; }
        public string Title { get; private set; }
        public string Description { get; private set; }
        public bool IsReadOnly { get; private set; }
        public uint TypeKind { get; private set; }
        public string TypeName { get; private set; }
        public bool Required { get; private set; }
        public bool Filterable { get; private set; }
        public bool Sortable { get; private set; }
        public string DefaultValue { get; private set; }

        public static FieldInformation GetInformation(Field field)
        {
            Type type = field.GetType();
            if (typeof(FieldDateTime) == type)
            {
                return new FieldDateTimeInformation((FieldDateTime)field);
            }
            if (typeof(FieldCurrency) == type)
            {
                return new FieldCurrencyInformation((FieldCurrency)field);
            }
            if (typeof(FieldNumber) == type)
            {
                return new FieldNumberInformation((FieldNumber)field);
            }
            if (typeof(FieldText) == type)
            {
                return new FieldTextInformation((FieldText)field);
            }
            if (typeof(FieldUrl) == type)
            {
                return new FieldUrlInformation((FieldUrl)field);
            }
            if (typeof(FieldUser) == type)
            {
                return new FieldUserInformation((FieldUser)field);
            }
            if (typeof(FieldChoice) == type)
            {
                return new FieldChoiceInformation((FieldChoice)field);
            }
            if (typeof(FieldMultiChoice) == type)
            {
                return new FieldMultiChoiceInformation((FieldMultiChoice)field);
            }
            if (typeof(FieldLookup) == type)
            {
                return new FieldLookupInformation((FieldLookup)field);
            }
            if (typeof(FieldMultiLineText) == type)
            {
                return new FieldMultiLineTextInformation((FieldMultiLineText)field);
            }
            return new FieldInformation(field);
        }
    }

    public class FieldTextInformation : FieldInformation
    {
        public FieldTextInformation(FieldText field) : base(field)
        {
            MaxLength = field.MaxLength;
        }

        public int MaxLength { get; private set; }
    }

    public class FieldNumberInformation : FieldInformation
    {
        public FieldNumberInformation(FieldNumber field) : base(field)
        {
            MinimumValue = field.MinimumValue;
            MaximumValue = field.MaximumValue;
        }

        public double MaximumValue { get; private set; }
        public double MinimumValue { get; private set; }
    }

    public class FieldUrlInformation : FieldInformation
    {
        public FieldUrlInformation(FieldUrl field) : base(field)
        {
            Format = (uint)field.DisplayFormat;
        }

        public uint Format { get; private set; }
    }

    public class FieldUserInformation : FieldInformation
    {
        public FieldUserInformation(FieldUser field) : base(field)
        {
            Group = field.SelectionGroup;
            Mode = (uint)field.SelectionMode;
            AllowDisplay = field.AllowDisplay;
            Presence = field.Presence;
        }

        public uint Mode { get; private set; }
        public int Group { get; private set; }
        public bool AllowDisplay { get; private set; }
        public bool Presence { get; private set; }
    }

    public class FieldMultiChoiceInformation : FieldInformation
    {
        public FieldMultiChoiceInformation(FieldMultiChoice field) : base(field)
        {
            Choices = field.Choices;
            FillInChoice = field.FillInChoice;
            Mappings = field.Mappings;
        }

        public string[] Choices { get; private set; }
        public bool FillInChoice { get; private set; }
        public string Mappings { get; private set; }
    }

    public class FieldChoiceInformation : FieldMultiChoiceInformation
    {
        public FieldChoiceInformation(FieldChoice field) : base(field)
        {
            EditFormat = (uint)field.EditFormat;

        }
        public uint EditFormat { get; private set; }
    }

    public class FieldCurrencyInformation : FieldNumberInformation
    {
        public FieldCurrencyInformation(FieldCurrency field) : base(field)
        {
            CurrencyLocaleId = field.CurrencyLocaleId;
        }

        public int CurrencyLocaleId { get; private set; }
    }

    public class FieldDateTimeInformation : FieldInformation
    {
        public FieldDateTimeInformation(FieldDateTime field) : base(field)
        {
            DateTimeCalendarType = (uint)field.DateTimeCalendarType;
            DisplayFormat = (uint)field.DisplayFormat;
            FriendlyDisplayFormat = (uint)field.FriendlyDisplayFormat;
        }

        public uint DateTimeCalendarType { get; private set; }
        public uint DisplayFormat { get; private set; }
        public uint FriendlyDisplayFormat { get; private set; }
    }

    public class FieldMultiLineTextInformation : FieldInformation
    {
        public FieldMultiLineTextInformation(FieldMultiLineText field) : base(field)
        {
            AllowHyperlink = field.AllowHyperlink;
            AppendOnly = field.AppendOnly;
            NumberOfLines = field.NumberOfLines;
            RichText = field.RichText;
            RestrictedMode = field.RestrictedMode;
            WikiLinking = field.WikiLinking;
        }

        public bool AllowHyperlink { get; set; }
        public bool AppendOnly { get; set; }
        public int NumberOfLines { get; set; }
        public bool RestrictedMode { get; set; }
        public bool RichText { get; set; }
        public bool WikiLinking { get; }
    }


    public class FieldLookupInformation : FieldInformation
    {
        public FieldLookupInformation(FieldLookup field) : base(field)
        {
            AllowMultipleValues = field.AllowMultipleValues;
            IsRelationship = field.IsRelationship;
            RelationshipDeleteBehavior = (uint)field.RelationshipDeleteBehavior;
            LookupField = field.LookupField;
            LookupList = field.LookupList;
            LookupWebId = field.LookupWebId;
            PrimaryFieldId = field.PrimaryFieldId;
        }

        public bool AllowMultipleValues { get; set; }
        public bool IsRelationship { get; set; }
        public string LookupField { get; set; }
        public string LookupList { get; set; }
        public Guid LookupWebId { get; set; }
        public string PrimaryFieldId { get; set; }
        public uint RelationshipDeleteBehavior { get; set; }
    }
}
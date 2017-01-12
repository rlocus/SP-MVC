using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace SPMVCWeb.Models
{
    public class SPRegionalInfo
    {
        [JsonProperty("adjustHijriDays")]
        public short AdjustHijriDays { get; set; }

        [JsonProperty("alternateCalendarType")]
        public short AlternateCalendarType { get; set; }

        [JsonProperty("am")]
        public string AM { get; set; }

        [JsonProperty("calendarType")]
        public short CalendarType { get; set; }

        [JsonProperty("collation")]
        public short Collation { get; set; }

        [JsonProperty("collationLCID")]
        public uint CollationLCID { get; set; }

        [JsonProperty("dateFormat")]
        public uint DateFormat { get; set; }

        [JsonProperty("dateSeparator")]
        public string DateSeparator { get; set; }

        [JsonProperty("decimalSeparator")]
        public string DecimalSeparator { get; set; }

        [JsonProperty("digitGrouping")]
        public string DigitGrouping { get; set; }

        [JsonProperty("firstDayOfWeek")]
        public uint FirstDayOfWeek { get; set; }

        [JsonProperty("firstWeekOfYear")]
        public short FirstWeekOfYear { get; set; }

        [JsonProperty("isEastAsia")]
        public bool IsEastAsia { get; set; }

        [JsonProperty("isRightToLeft")]
        public bool IsRightToLeft { get; set; }

        [JsonProperty("isUIRightToLeft")]
        public bool IsUIRightToLeft { get; set; }

        [JsonProperty("listSeparator")]
        public string ListSeparator { get; set; }

        [JsonProperty("localeId")]
        public uint LocaleId { get; set; }

        [JsonProperty("negativeSign")]
        public string NegativeSign { get; set; }

        [JsonProperty("negNumberMode")]
        public uint NegNumberMode { get; set; }

        [JsonProperty("pm")]
        public string PM { get; set; }

        [JsonProperty("positiveSign")]
        public string PositiveSign { get; set; }

        [JsonProperty("showWeeks")]
        public bool ShowWeeks { get; set; }

        [JsonProperty("thousandSeparator")]
        public string ThousandSeparator { get; set; }

        [JsonProperty("time24")]
        public bool Time24 { get; set; }

        [JsonProperty("timeMarkerPosition")]
        public uint TimeMarkerPosition { get; set; }

        [JsonProperty("timeSeparator")]
        public string TimeSeparator { get; set; }

        [JsonProperty("workDayEnd")]
        public short WorkDayEndHour { get; set; }

        [JsonProperty("workDays")]
        public short WorkDays { get; set; }

        [JsonProperty("workDayStart")]
        public short WorkDayStartHour { get; set; }
    }
}
using Newtonsoft.Json;

namespace SPMVCWeb.Models
{
    public class SPPageContextInfo
    {
        [JsonProperty("webServerRelativeUrl")]
        public string WebServerRelativeUrl { get; set; }

        [JsonProperty("webAbsoluteUrl")]
        public string WebAbsoluteUrl { get; set; }

        [JsonProperty("siteServerRelativeUrl")]
        public string SiteServerRelativeUrl { get; set; }

        [JsonProperty("siteAbsoluteUrl")]
        public string SiteAbsoluteUrl { get; set; }

        [JsonProperty("layoutsUrl")]
        public string LayoutsUrl { get; set; }

        [JsonProperty("webTitle")]
        public string WebTitle { get; set; }

        [JsonProperty("webLogoUrl")]
        public string WebLogoUrl { get; set; }

        [JsonProperty("webLanguage")]
        public uint WebLanguage { get; set; }

        //[JsonProperty("currentLanguage")]
        //public string CurrentLanguage { get; set; }

        //[JsonProperty("currentUICultureName")]
        //public string CurrentUICultureName { get; set; }

        //[JsonProperty("currentCultureName")]
        //public string CurrentCultureName { get; set; }

        [JsonProperty("userId")]
        public int UserId { get; set; }

        [JsonProperty("userLoginName")]
        public string UserLoginName { get; set; }

        [JsonProperty("webPermMasks")]
        public int[] WebPermMasks { get; set; }

        [JsonProperty("webUIVersion")]
        public int WebUIVersion { get; set; }

        [JsonProperty("appWebUrl")]
        public string AppWebUrl { get; set; }

        [JsonProperty("regionalSettings")]
        public SPRegionalInfo RegionalInfo { get; set; }
    }
}
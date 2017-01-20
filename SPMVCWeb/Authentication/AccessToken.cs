using System;

namespace AspNet.Owin.SharePoint.Addin.Authentication
{
	public class AccessToken
	{
		public string Value { get; set; }
		public DateTime ExpiresOn { get; set; }

        public bool IsValid()
        {
            return !string.IsNullOrEmpty(this.Value) && this.ExpiresOn > DateTime.UtcNow;
        }
    }
}

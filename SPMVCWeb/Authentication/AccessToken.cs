using System;

namespace AspNet.Owin.SharePoint.Addin.Authentication
{
	public class AccessToken
	{
		public string Value { get; set; }
		public DateTime ExpiredOn { get; set; }
	}
}

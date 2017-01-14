﻿using System;
using Microsoft.SharePoint.Client;
using System.Text.RegularExpressions;

namespace SPMVCWeb.Models
{
    public class UserInformation
    {
        internal UserInformation(User spUser)
        {
            if (spUser == null) throw new ArgumentNullException("spUser");
            Id = spUser.Id;
            Initials = new Regex(@"(\b[a-zA-Z])[a-zA-Z]* ?").Replace(spUser.Title, "$1");
            Name = spUser.Title;
            Login = spUser.LoginName;
            IsSiteAdmin = spUser.IsSiteAdmin;
            Email = spUser.Email;
        }

        public int Id { get; private set; }
        public string Initials { get; private set; }
        public string Name { get; private set; }
        public string Login { get; private set; }
        public string Email { get; private set; }
        public bool IsSiteAdmin { get; private set; }
    }
}
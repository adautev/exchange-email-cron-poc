using System;
using System.Collections.Generic;
namespace Saorsa.Outlook.Mail
{
    class UserEmailBox
    {
        private string userId;
        private string body;

        public string Body { get => body; set => body = value; }

        public string UserId { get => userId; set => userId = value; }
        
    }
}
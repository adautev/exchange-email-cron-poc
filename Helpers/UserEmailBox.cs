using System;
using System.Collections.Generic;
namespace Saorsa.Outlook.Mail
{
    class UserEmailBox
    {
        private string userId;

        private List<EmailMessage> value = new List<EmailMessage>();

        public string UserId { get => userId; set => userId = value; }
        
        public List<EmailMessage> GetUserEmail()
        {
            return value;
        }

        public void SetUserEmail(EmailMessage message)
        {
            value.Add(message);
        }
    }
}
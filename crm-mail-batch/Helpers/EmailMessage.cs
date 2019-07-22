using System;
using Microsoft.Graph;
namespace crm_mail_batch
{
    class EmailMessage
    {

        private string subject;
        private EmailAddress senderEmail;
        private string message;

        public string Subject { get => subject; set => subject = value; }

        public EmailAddress GetSenderEmail()
        {
            return senderEmail;
        }

        public void SetSenderEmail(EmailAddress value)
        {
            senderEmail = value;
        }

        public string Message { get => message; set => message = value; }

    }
}
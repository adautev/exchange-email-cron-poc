using System;
using Microsoft.Graph;
namespace Saorsa.Outlook.Mail
{
    class EmailMessage
    {

        private string subject;
        private EmailAddress senderEmail;
        private string message;

        private string itemId;

        private string conversationId;

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
        public string ItemId { get => itemId; set => itemId = value; }
        public string ConversationId { get => conversationId; set => conversationId = value; }
    }
}
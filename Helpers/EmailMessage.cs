using System;
using System.Runtime.Serialization;
using System.Text.RegularExpressions;
using Microsoft.Graph;
namespace Saorsa.Outlook.Mail {
    [DataContract]
    class EmailMessage {
        private string _opportunityMatch { get; set; }
        public EmailMessage (string opportunityMatch) {
            this._opportunityMatch = opportunityMatch;

        }
        private string subject;
        private EmailAddress senderEmail;
        private string message;

        private string itemId;

        private string conversationId;
        [DataMember (Name = "subject")]
        public string Subject { get => subject; set => subject = value; }

        public EmailAddress SenderEmail { get => senderEmail; set => senderEmail = value; }

        public string Email { get => senderEmail.Address; }

        [DataMember (Name = "mailBody")]
        public string Message { get => message; set => message = value; }

        [DataMember (Name = "mailItemId")]
        public string ItemId { get => itemId; set => itemId = value; }

        [DataMember (Name = "conversationId")]
        public string ConversationId { get => conversationId; set => conversationId = value; }

        [DataMember (Name = "opportunityId")]
        public string OpportunityId { get => Regex.Match(Subject, String.Concat(_opportunityMatch.Replace("\"",""),@"\d+")).Value; }
    }
}
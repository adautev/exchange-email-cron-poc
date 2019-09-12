using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.Serialization;
using System.Text.RegularExpressions;
using Microsoft.Graph;
namespace Saorsa.Outlook.Mail {
    [DataContract]
    public class EmailMessage {
        public class Recipient {
            [DataMember (Name = "displayName")]
            public readonly string DisplayName;

            [DataMember (Name = "emailAddress")]
            public readonly string EmailAddress;

            public Recipient (string displayName, string emailAddress) {
                this.EmailAddress = emailAddress;
                this.DisplayName = displayName;
            }

        }
        private string _opportunityMatch { get; set; }
        public EmailMessage (string opportunityMatch) {
            this._opportunityMatch = opportunityMatch;
            this.recipients = new List<Recipient>();
        }
        private string subject;
        private EmailAddress senderEmail;
        private string message;

        private string itemId;

        private string conversationId;
        private List<Recipient> recipients;

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

        [DataMember (Name = "recipients")]
        public IEnumerable<Recipient> Recipients { get => recipients; }
        
        public void SetRecipient(IEnumerable<Microsoft.Graph.Recipient> recipients) {
            if (recipients != null && recipients.Count() > 0) {
                this.recipients.AddRange(recipients.Select(r=> new Recipient(r.EmailAddress.Name, r.EmailAddress.Address)));
            }
        }

        [DataMember (Name = "opportunityId")]
        public string OpportunityId { get => Regex.Match (Subject, String.Concat (_opportunityMatch.Replace ("\"", ""), @"\d+")).Value; }
    }
}
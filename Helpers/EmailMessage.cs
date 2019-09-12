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
            this._recipients = new List<Recipient>();
        }
        private string _subject;
        private EmailAddress _senderEmail;
        private string _message;

        private string _itemId;

        private string _conversationId;
        private List<Recipient> _recipients;

        [DataMember (Name = "subject")]
        public string Subject { get => _subject; set => _subject = value; }

        public EmailAddress SenderEmail { get => _senderEmail; set => _senderEmail = value; }

        public string Email { get => _senderEmail.Address; }

        [DataMember (Name = "mailBody")]
        public string Message { get => _message; set => _message = value; }

        [DataMember (Name = "mailItemId")]
        public string ItemId { get => _itemId; set => _itemId = value; }

        [DataMember (Name = "conversationId")]
        public string ConversationId { get => _conversationId; set => _conversationId = value; }

        [DataMember (Name = "recipients")]
        public IEnumerable<Recipient> Recipients { get => _recipients; }
        
        public void SetRecipient(IEnumerable<Microsoft.Graph.Recipient> recipients) {
            if (recipients != null && recipients.Count() > 0) {
                this._recipients.AddRange(recipients.Select(r=> new Recipient(r.EmailAddress.Name, r.EmailAddress.Address)));
            }
        }

        [DataMember (Name = "opportunityId")]
        public string OpportunityId { get => Regex.Match (Subject, String.Concat (_opportunityMatch.Replace ("\"", ""), @"\d+")).Value; }
    }
}
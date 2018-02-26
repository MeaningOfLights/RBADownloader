using Microsoft.Exchange.WebServices.Data;
using System;
using System.Configuration;

namespace FileDownloader
{
    public class Email
    {
        #region "Constants & Variables"

        private ExchangeService exchangeService = new ExchangeService(ExchangeVersion.Exchange2010_SP2);

        #endregion

        #region "Constructor"

        public Email()
        {
            try
            {
                exchangeService.Url = new Uri("https://webmail.YourExchangeServer.com/EWS/Exchange.asmx");

#if DEBUG
                exchangeService.Credentials = new WebCredentials(Environment.UserName, "password", "DOMAIN");
#else
				exchangeService.UseDefaultCredentials = true;
#endif

            }
            catch (System.Runtime.InteropServices.COMException ex)
            {
            }
        }

        #endregion


        /// <summary>
        ///  Send an email to draft
        /// </summary>
        /// <param name="emailTo"></param>
        /// <param name="emailCc"></param>
        /// <param name="emailBcc"></param>
        /// <param name="emailSubject"></param>
        /// <param name="emailBody"></param>
        /// <param name="emailFileAttachments"></param>
        /// <param name="emailFromSharedMailbox"></param>
        /// <remarks></remarks>
        public void SendEmail(string emailTo, string emailCc, string emailBcc, string emailSubject, string emailBody, string[] emailFileAttachments, bool emailFromSharedMailbox = false, bool saveToDraftAndDontSend = false)
        {
            EmailMessage message = default(EmailMessage);
            message = new EmailMessage(exchangeService);

            emailTo = emailTo.TrimEnd(';', ',');
            string[] emailArr = emailTo.Split(';', ',');
            if (emailTo.Length > 0)
                message.ToRecipients.AddRange(emailArr);

            emailCc = emailCc.TrimEnd(';', ',');
            emailArr = emailCc.Split(';', ',');
            if (emailCc.Length > 0)
                message.CcRecipients.AddRange(emailArr);

            emailBcc = emailBcc.TrimEnd(';', ',');
            emailArr = emailBcc.Split(';', ',');
            if (emailBcc.Length > 0)
                message.BccRecipients.AddRange(emailArr);

#if DEBUG
            emailSubject = "IGNORE - TESTING ONLY - " + emailSubject;
#endif
            message.Subject = emailSubject;

            if ((emailFileAttachments == null) == false)
            {
                foreach (string fileAttachment in emailFileAttachments)
                {
                    if (string.IsNullOrEmpty(fileAttachment) == false)
                        message.Attachments.AddFileAttachment(fileAttachment);
                }
            }

            message.Sensitivity = Sensitivity.Private;
            message.Body = new MessageBody();
            message.Body.BodyType = BodyType.HTML;
            message.Body.Text = emailBody;
            //if (emailFromSharedMailbox) message.Body.Text += "<br><br>" + _sharedMailSignature;

            //Save the email message to the Drafts folder (where it can be retrieved, updated, and sent at a later time). 
            if (saveToDraftAndDontSend)
            {
                message.Save(WellKnownFolderName.Drafts);
            }
            else
            {
                try
                {
                    message.Send();
                }
                catch (Exception ex)
                {
                    System.Windows.Forms.MessageBox.Show("FAILED TO SEND EMAIL - PLEASE CONTACT " + ConfigurationManager.AppSettings["EmailSupport"].ToString() + "!", "Email sending failed", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Exclamation);
                }

            }
        }
    }
}

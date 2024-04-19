using MailKit;
using MailKit.Net.Imap;
using MailKit.Search;
using MailKit.Security;
using Microsoft.Identity.Client;
using MimeKit;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.SqlClient;
using System.Data.SqlTypes;

namespace WInFormsImapMailKit
{
    public partial class ImapForm : Form
    {
        string strConnectionString;
        public ImapForm()
        {
            InitializeComponent();
        }

        public async void ImapForm_Load(object sender, EventArgs e)
        {
            var options = new PublicClientApplicationOptions
            {
                ClientId = "clientId", 
                TenantId = "tenantId", 
                RedirectUri = "RedirectUri",

                
            };

            var publicClientApplication = PublicClientApplicationBuilder
                .CreateWithApplicationOptions(options)
                .Build();

            var scopes = new string[] {
            "email",
            "offline_access",
            "openid",
            "profile",
            "https://outlook.office.com/IMAP.AccessAsUser.All", // Only needed for IMAP
            //"https://outlook.office.com/POP.AccessAsUser.All",  // Only needed for POP
            //"https://outlook.office.com/SMTP.Send", // Only needed for SMTP
            };

          
            var authToken = await publicClientApplication.AcquireTokenInteractive(scopes).ExecuteAsync();
            var oauth = new SaslMechanismOAuth2(authToken.Account.Username, authToken.AccessToken);

            checa_emails2(oauth);
        }

        private async void ImapForm_Activated(object sender, EventArgs e)
        {
           

        }

        private async void checa_emails2(SaslMechanismOAuth2 oauth)
        {
            using (var client = new ImapClient())
            {
                
                await client.ConnectAsync("outlook.office365.com", 993, SecureSocketOptions.SslOnConnect);
                await client.AuthenticateAsync(oauth);

                Start:

                // The Inbox folder is always available on all IMAP servers...
                var inbox = client.Inbox;
                inbox.Open(FolderAccess.ReadWrite);


                var query = SearchQuery.DeliveredAfter(DateTime.Parse("30-06-2023")).And(SearchQuery.NotSeen);
                //var query = SearchQuery.NotSeen;
                var uids = client.Inbox.Search(SearchQuery.NotSeen);

                //MessageBox.Show(uids.ToString());

                // virgula para lista
                List<string> uidsList = uids.ToString().Split(',').ToList();

                bool Entrou = false;

                foreach (var uid in inbox.Search(query))
                {
                    Entrou = true;
                    var message = inbox.GetMessage(uid);
                    ////Console.WriteLine("Anexo: {1}", uid, message.Attachments);
                    //inbox.AddFlags(uid, MessageFlags.Seen, true);


                    
                    string assunto = message.Subject;//obter assunto
                                                     //string origem = message.From.ToString(); //obter origem
                                                     //string mensagem1 = message.TextBody;

                    MessageBox.Show(assunto);

                    string mensagem = message.Body.ToString().Replace("\r\n", "");

                    string origem = message.From.OfType<MailboxAddress>().Single().Address;

                    DateTime data = Convert.ToDateTime(message.Date.ToString()); // .Substring(0, 10)

                    ////string uid = NewMessage[y].Headers.MessageId; //obter id da mensagem

                    //Guardar mensagem em disco
                    string file_name = uid.ToString() + ".eml";
                }


                if (Entrou == true)
                {
                    inbox.SetFlags(uids, MessageFlags.Seen, true);
                }

                MessageBox.Show("pausa 60 segundos e depois reinicia aplicacacao");

                System.Threading.Thread.Sleep(20000);

                //await client.DisconnectAsync(false);
                client.Disconnect(false);

                goto Start;

            }

           
            //Application.Exit();
        }
    }
}
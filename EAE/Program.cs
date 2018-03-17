using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Net;
using System.Net.Security;
using System.Security.Cryptography.X509Certificates;
using Microsoft.Exchange.WebServices.Data;
using System.Configuration;
using System.Globalization;
using System.Threading;
using System.Web.Script.Serialization;
using Microsoft.Exchange.WebServices.Autodiscover;
using Microsoft.IdentityModel.Clients.ActiveDirectory;
using System.IO;

namespace EAE
{
    class Program
    {
        static void Main(string[] args)
        {
            string authority = ConfigurationManager.AppSettings["authority"];
            string clientId = ConfigurationManager.AppSettings["clientID"];
            Uri clientAppUri = new Uri(ConfigurationManager.AppSettings["clientAppUri"]);
            string serverName = ConfigurationManager.AppSettings["serverName"];
            string mailAccountToWatch = ConfigurationManager.AppSettings["mailAccountToWatch"];
            string locationForExtraction = ConfigurationManager.AppSettings["locationForExtraction"];
            int waitTimeBetweenChecks = 120000;
            ExchangeService exchangeService = new ExchangeService(ExchangeVersion.Exchange2013);
            AuthenticationResult authenticationResult = null;
            AuthenticationContext authenticationContext = new AuthenticationContext(authority, false);

            Console.WriteLine("Acquiring the access token now...");
            string errorMessage = null;
            try
            {
                //authenticationResult = authenticationContext.AcquireToken(serverName, clientId, new UserCredential(mailLogin, mailPassword));
                authenticationResult = authenticationContext.AcquireToken(serverName, clientId, clientAppUri);

            }
            catch (AdalException ex)
            {
                errorMessage = ex.Message;
                if (ex.InnerException != null)
                {
                    errorMessage += "\nInnerException : " + ex.InnerException.Message;
                }
            }
            catch (ArgumentException ex)
            {
                errorMessage = ex.Message;
            }
            if (!string.IsNullOrEmpty(errorMessage))
            {
                Console.WriteLine("Failed: {0}, press enter to exit. Error message: " + errorMessage);
                Console.ReadLine();
                Environment.Exit(0);
                return;
            }
            Console.WriteLine("Got the token!");
            string currentToken = authenticationResult.AccessToken;
            //setAppSetting("credentials", authenticationResult.AccessToken);
            Console.WriteLine("--------END");
            Console.WriteLine("\nMaking the protocol call\n");
            exchangeService.Url = new Uri("https://outlook.office365.com/EWS/Exchange.asmx");
            exchangeService.TraceEnabled = true;
            exchangeService.TraceFlags = TraceFlags.DebugMessage;
            exchangeService.Credentials = new OAuthCredentials(authenticationResult.AccessToken);

            // Test 1: Write an Email
            //exchangeService.FindFolders(WellKnownFolderName.Root, new FolderView(10));
            //EmailMessage email = new EmailMessage(exchangeService);
            //email.ToRecipients.Add("test@domain.comtttt");
            //email.Subject = "HelloWorld";
            //email.Body = new MessageBody("This is the first email I've sent by using the EWS Managed API");
            //email.Send();

            // Test 2: Get Folders and their ID's withing mailbox (ID's are required for the EWS move if not a wellknonfoldername)
            FolderView view = new FolderView(100);
            view.PropertySet = new PropertySet(BasePropertySet.IdOnly);
            view.PropertySet.Add(FolderSchema.DisplayName);
            view.Traversal = FolderTraversal.Deep;
            FindFoldersResults findFolderResults = exchangeService.FindFolders(new FolderId(WellKnownFolderName.Root, mailAccountToWatch), view);
            foreach (Folder f in findFolderResults)
            {
                    Console.WriteLine("Foldername: " + f.DisplayName + " ID: " + f.Id);
            }

            // Test 3: Extract the Attachments of the first message in the inbox of the defined mail account and move it to processed mails folder
            // get the inbox folder of the tdb shared mail account
            FolderId folderToAccess = new FolderId(WellKnownFolderName.Inbox, mailAccountToWatch);
            // find all the mails in the inbox
            FindItemsResults<Item> findResults = exchangeService.FindItems(folderToAccess, new ItemView(1)); // take the next x mail items in the inbox
            foreach (Item item in findResults.Items)
            {
                Console.WriteLine(item.Subject);
                EmailMessage message = EmailMessage.Bind(exchangeService, item.Id, new PropertySet(ItemSchema.Attachments));
                foreach (Attachment attachment in message.Attachments)
                {
                    if (attachment is FileAttachment)
                    {
                        FileAttachment fileAttachment = attachment as FileAttachment;
                        fileAttachment.Load(locationForExtraction + fileAttachment.Name);
                        Console.WriteLine("Extracted the following file: " + fileAttachment.Name);
                    }
                }
                //item.Delete(DeleteMode.MoveToDeletedItems);
                item.Move(WellKnownFolderName.ArchiveMsgFolderRoot);
            }

            // Infinity loop to extract attachments of incoming mail and move the mails to the processed mails folder
            while (true)
            {
                // refreshing access token, expires usually after 60 minutes wich would give error 401, hence we always request a new one
                authenticationResult = authenticationContext.AcquireToken(serverName, clientId, clientAppUri);
                exchangeService.Credentials = new OAuthCredentials(authenticationResult.AccessToken);
                if (currentToken != authenticationResult.AccessToken)
                {
                    currentToken = authenticationResult.AccessToken;
                    Console.WriteLine(DateTime.Now + " Got a new token!");
                }
                findResults = exchangeService.FindItems(folderToAccess, new ItemView(1));
                foreach (Item item in findResults.Items)
                {
                    Console.WriteLine(DateTime.Now + " -- " + item.Subject);
                }
                //Thread.Sleep(4200000); //70 minutes
                Console.WriteLine("Waiting for: " + waitTimeBetweenChecks / 1000 / 60 + " minutes");
                Thread.Sleep(waitTimeBetweenChecks); 
                waitTimeBetweenChecks = waitTimeBetweenChecks + 1200000;
            }
        }

        static void setAppSetting(string key, string value)
        {
            //Laden der AppSettings
            Configuration config = ConfigurationManager.OpenExeConfiguration(System.Reflection.Assembly.GetExecutingAssembly().Location);
            //Überprüfen ob Key existiert
            if (config.AppSettings.Settings[key] != null)
            {
                //Key existiert. Löschen des Keys zum "überschreiben"
                config.AppSettings.Settings.Remove(key);
            }
            //Anlegen eines neuen KeyValue-Paars
            config.AppSettings.Settings.Add(key, value);
            //Speichern der aktualisierten AppSettings
            config.Save(ConfigurationSaveMode.Modified);
        }
    }
}

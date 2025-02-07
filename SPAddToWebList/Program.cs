using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text;
using System.Threading.Tasks;
using Microsoft.SharePoint.Client;
using System.IO;
using static System.Net.Mime.MediaTypeNames;
using Microsoft.Exchange.WebServices.Data;


namespace SPAddToWebList
{
    internal class Program
    {
        static void Main(string[] args)
        {

            if (args.Length == 1 && args[0] == "check")
            {
                CheckSiteAlive();
            }

            if (args.Length == 4 && args[0] == "create")
            {
                // Erst nextNumber holen
                var nextNumber = GetNextNumber();
                // Item anlegen
                addSiteToList(args[1], args[2], args[3], nextNumber.ToString());
            }

        }
        private static void addSiteToList(string Title, string Link, string Image, string Number)
        {
            using (var cc = new ClientContext(@"https://sharepoint.url.local"))
            {



                var List = cc.Web.Lists.GetByTitle("AllSites");
                ListItemCreationInformation myItemCreationInfo = new ListItemCreationInformation();
                ListItem newItem = List.AddItem(myItemCreationInfo);

                newItem["Title"] = Title;
                newItem["BackgroundImageLocation"] = Image;
                newItem["LinkLocation"] = Link;
                newItem["TileOrder"] = Number;

                newItem.Update();

                try
                {
                    if (Convert.ToInt32(Number) > 0)
                    {
                        cc.ExecuteQuery();
                        Console.WriteLine("Eintrag in AllSite Liste erzeugt!");
                    }

                    else { Console.WriteLine("Fehler: TileOrder Number ist 0. Item wird nicht angelegt."); }

                }
                catch (Exception e)
                {

                    Console.WriteLine($"Error: {e.Message} ");
                }



            }
        }
        private static int GetNextNumber()
        {
            using (var cc = new ClientContext(@"https://sharepoint.url.local"))
            {

                var ourList = cc.Web.Lists.GetByTitle("AllSites");

                // CAML Query erstellen, die keine Bedingungen enthält (=> wir holen alle Items)
                // Wenn man fiese Queries bauen möchte, kann man dieses Tool benutzen um die Abfragen sauber zu erstellen:
                // https://github.com/konradsikorski/smartCAML

                CamlQuery camlQuery = new CamlQuery();
                ListItemCollection itemCollection = ourList.GetItems(camlQuery);

                // Wir holen nur das was wir brauchen
                cc.Load(
                itemCollection,
                items => items.Include(
                item => item["TileOrder"]));

                cc.ExecuteQuery();

                // In unserer Liste aller Items "itemCollection", das Item finden, das den höchsten Wert bei "TileOrder" besitzt
                var listItem = itemCollection.Max(i => i["TileOrder"]);


                try
                {
                    var high = Convert.ToInt32(listItem.ToString());
                    return high + 1;
                }
                catch (Exception e)
                {
                    Console.WriteLine($"Error: {e.Message} ");
                    Console.WriteLine("ItemCollection has " + itemCollection.Count + " Elements!");
                }

                // Wenn wir ein Problem haben geben wir 0 zurück
                return 0;


            }
        }

        private static void CheckSiteAlive()
        {
            List<string> list = GetSitesFromSpList();

            foreach (string s in list) {

                try
                {
                    using (var cc = new ClientContext(s))
                    {
                        var List = cc.Web.Lists.GetByTitle("FormServerTemplates");
                        cc.ExecuteQuery();
                    }

                }
                catch (Exception ex) {

                    if (ex.Message.Contains("The remote server returned an error: (404) Not Found."))
                    {
                        using (var cc = new ClientContext(@"https://sharepoint.url.local"))
                        {
                            var ourList = cc.Web.Lists.GetByTitle("AllSites");
                            CamlQuery camlQuery = new CamlQuery();
                            ListItemCollection itemCollection = ourList.GetItems(camlQuery);
                            cc.Load(itemCollection,
                                    items => items.Include(
                                    item => item.Id,
                                    item => item.DisplayName,
                                    item => item["LinkLocation"]));

                            cc.ExecuteQuery();

                            // Convert to a list to enable LINQ processing
                            var itemsList = itemCollection.ToList();

                            // Find the matching item

                            var matchedItem = itemsList.FirstOrDefault(item =>
                            {

                                if (item["LinkLocation"] is FieldUrlValue urlValue)
                                {
                                    return urlValue.Url == s;
                                }
                                return false;
                            });

                            if (matchedItem != null)
                            {
                                // Save details for email notification
                                int deletedItemId = matchedItem.Id;
                                string deletedItemName = matchedItem.DisplayName;

                                // Delete the item
                                matchedItem.DeleteObject();
                                cc.ExecuteQuery(); // Commit deletion to SharePoint

                                Console.WriteLine($"Item deleted: ID = {matchedItem.Id}, DisplayName = {matchedItem.DisplayName}");
                                SendEmailNotification(deletedItemId, deletedItemName, s);
                            }
                            else
                            {
                                Console.WriteLine("No matching item found to delete.");
                            }
                        }
                    }
                }
            }
        }

        static void SendEmailNotification(int itemId, string itemName, string itemUrl)
        {
            try
            {
                // Setup ExchangeService
                ExchangeService service = new ExchangeService(ExchangeVersion.Exchange2013)
                {
                    Credentials = new WebCredentials("", "", "")
                };

                // Set Exchange server URL
                service.Url = new Uri("https://exchangehostname.local/EWS/Exchange.asmx");

                // Create email message
                EmailMessage email = new EmailMessage(service)
                {
                    Subject = "SharePoint List Item Deletion Notification",
                    Body = $"Hello,\n\nAn item has been deleted from the SharePoint list 'AllSites'.\n\n" +
                           $"📌 **Deleted Item Details:**\n" +
                           $"- **ID:** {itemId}\n" +
                           $"- **Name:** {itemName}\n" +
                           $"- **Link Location:** {itemUrl}\n\n" +
                           $"This action was performed automatically.\n\nBest regards,\nYour SharePoint Admin Bot"
                };

                // Add recipient
                email.ToRecipients.Add("verteilerspadmins@koeln-bonn-airport.de");

                // Send the email
                email.Send();
                Console.WriteLine("Notification email sent successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error sending email: " + ex.Message);
            }
        }
    
        private static List<string> GetSitesFromSpList()
        {
            var sitesList = new List<string>();
            using (var cc = new ClientContext(@"https://sharepoint.url.local"))
            {
                var ourList = cc.Web.Lists.GetByTitle("AllSites");
                CamlQuery camlQuery = new CamlQuery();
                ListItemCollection itemCollection = ourList.GetItems(camlQuery);
                cc.Load(itemCollection,
                        items => items.Include(
                        item => item["LinkLocation"]));

                cc.ExecuteQuery();
                
                foreach (var item in itemCollection)
                {
                    if (item["LinkLocation"] is FieldUrlValue urlValue)
                    {
                        sitesList.Add(urlValue.Url);
                    }
                }

            }
            foreach (var site in sitesList) {Console.WriteLine(site.ToString());}
            return sitesList;
        }
}
}

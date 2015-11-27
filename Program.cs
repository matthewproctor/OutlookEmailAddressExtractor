///
/// Outlook Email Address Extractor 
/// Version 0.1
/// Build 2015-Nov-18
/// Written by Matthew Proctor
/// www.matthewproctor.com
///
using System;
using System.Linq;
using System.IO;
using System.Collections.Generic;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace OutlookEmailAddressExtractor
{
    class Program
    {

        // Array to store email addresses and counts
        public static List<string> emailAddresses = new List<string>();
        public static List<int> emailAddressesCounter = new List<int>();

        static void Main(string[] args)
        {
            EnumerateAccounts();
            //EnumerateFoldersInDefaultStore();
            //Console.WriteLine("Total file size:" + totalfilesize);
        }

        static void EnumerateFoldersInDefaultStore()
        {
            Outlook.Application Application = new Outlook.Application();
            Outlook.Folder root = Application.Session.DefaultStore.GetRootFolder() as Outlook.Folder;
            EnumerateFolders(root);
        }

        // Uses recursion to enumerate Outlook subfolders.
        static void EnumerateFolders(Outlook.Folder folder)
        {
            Outlook.Folders childFolders = folder.Folders;
            if (childFolders.Count > 0)
            {
                foreach (Outlook.Folder childFolder in childFolders)
                {
                    // We only want Inbox folders - ignore Contacts and others
                    if (childFolder.FolderPath.Contains("Inbox"))
                    {
                        // Write the folder path.
                        Console.WriteLine(childFolder.FolderPath);
                        // Call EnumerateFolders using childFolder, to see if there are any sub-folders within this one
                        EnumerateFolders(childFolder);
                    }
                }
            }
            Console.WriteLine("Checking in " + folder.FolderPath);
            IterateMessages(folder);
        }

        static void IterateMessages(Outlook.Folder folder)
        {
            // attachment extensions to save
            string[] extensionsArray = { ".pdf", ".doc", ".xls", ".ppt", ".vsd", ".zip", ".rar", ".txt", ".csv", ".proj" };

            // Iterate through all items ("messages") in a folder
            var fi = folder.Items;
            if (fi != null)
            {

                try
                {
                    foreach (Object item in fi)
                    {
                        Outlook.MailItem mailitem = (Outlook.MailItem)item;

                        string senderAddress = mailitem.Sender.Address;
                        add_address_to_list(senderAddress);

                        Outlook.Recipients recipients = mailitem.Recipients;
                        foreach (Outlook.Recipient recipient in recipients)
                        {
                            add_address_to_list(recipient.Address);
                        }


                    }
                }
                catch (Exception e)
                {
                    //Console.WriteLine("An error occurred: '{0}'", e);
                }
            }
        }

        static void add_address_to_list(string emailAddress)
        {
            if (emailAddress.Contains("@") && emailAddress.Contains("."))
            {
                bool found = false;
                for (int i = 0; i < emailAddresses.Count; i++)
                {
                    if (emailAddresses[i] == emailAddress)
                    {
                        // email address was found, so just increment it's counter
                        found = true;
                        emailAddressesCounter[i]++;
                    }
                }
                if (!found)
                {
                    // email address wasn't found, so add it to the array
                    emailAddresses.Add(emailAddress);
                    emailAddressesCounter.Add(1); //starts with a count of 1
                    Console.WriteLine(emailAddresses.Count + ": Added " + emailAddress);
                }
            }
        }

        // Retrieves the email address for a given account object
        static string EnumerateAccountEmailAddress(Outlook.Account account)
        {
            try
            {
                if (string.IsNullOrEmpty(account.SmtpAddress) || string.IsNullOrEmpty(account.UserName))
                {
                    Outlook.AddressEntry oAE = account.CurrentUser.AddressEntry as Outlook.AddressEntry;
                    if (oAE.Type == "EX")
                    {
                        Outlook.ExchangeUser oEU = oAE.GetExchangeUser() as Outlook.ExchangeUser;
                        return oEU.PrimarySmtpAddress;
                    }
                    else
                    {
                        return oAE.Address;
                    }
                }
                else
                {
                    return account.SmtpAddress;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                return "";
            }
        }

        static void EnumerateAccounts()
        {
            Console.Clear();
            Console.WriteLine("Outlook Email Address Extractor v0.1");
            Console.WriteLine("------------------------------------");
            int id;
            Outlook.Application Application = new Outlook.Application();
            Outlook.Accounts accounts = Application.Session.Accounts;

            string response = "";
            while (true == true)
            {

                id = 1;
                foreach (Outlook.Account account in accounts)
                {
                    Console.WriteLine(id + ":" + EnumerateAccountEmailAddress(account));
                    id++;
                }
                Console.WriteLine("Q: Quit Application");

                response = Console.ReadLine().ToUpper();
                if (response == "Q")
                {
                    Console.WriteLine("Quitting");
                    return;
                }
                if (response != "")
                {
                    if (Int32.Parse(response.Trim()) >= 1 && Int32.Parse(response.Trim()) < id)
                    {
                        Console.WriteLine("Processing: " + accounts[Int32.Parse(response.Trim())].DisplayName);
                        Console.WriteLine("Processing: " + EnumerateAccountEmailAddress(accounts[Int32.Parse(response.Trim())]));

                        Outlook.Folder selectedFolder = Application.Session.DefaultStore.GetRootFolder() as Outlook.Folder;
                        selectedFolder = GetFolder(@"\\" + accounts[Int32.Parse(response.Trim())].DisplayName);
                        EnumerateFolders(selectedFolder);
                        Console.WriteLine("Sorting results.");
                        sort_email_addresses();
                        Console.WriteLine("Saving results.");
                        save_email_addresses();
                        Console.WriteLine("Finished Processing " + accounts[Int32.Parse(response.Trim())].DisplayName);
                        Console.WriteLine("Addresses Found " + emailAddresses.Count);
                        Console.WriteLine("");
                    }
                    else
                    {
                        Console.WriteLine("Invalid Account Selected");
                    }
                }
            }

        }

        // Saves the output as a CSV file in the format emailaddress,counter 
        // in the current directory        
        static void save_email_addresses()
        {
            Console.WriteLine("Saving to: " + Directory.GetCurrentDirectory() + @"\output.csv");
            using (StreamWriter writetext = new StreamWriter(Directory.GetCurrentDirectory() + @"\output.csv"))
            {
                writetext.WriteLine("emailaddress,counter");
                for (int i = 0; i < emailAddresses.Count; i++)
                {
                    writetext.WriteLine(emailAddresses[i] + "," + emailAddressesCounter[i]);
                }
            }
        }

        // Uses a basic bubble sort to order the results by the email address
        // and persisting the position of the counter. 
        static void sort_email_addresses()
        {
            for (int i = 1; i < emailAddresses.Count; i++)
            {
                for (int d = 0; d < i; d++)
                {
                    if (String.Compare(emailAddresses[d], emailAddresses[i]) > 0)
                    {
                        string tempEmailAddress = emailAddresses[d];
                        emailAddresses[d] = emailAddresses[i];
                        emailAddresses[i] = tempEmailAddress;
                        int tempEmailAddressCount = emailAddressesCounter[d];
                        emailAddressesCounter[d] = emailAddressesCounter[i];
                        emailAddressesCounter[i] = tempEmailAddressCount;
                    }
                }
            }
        }

        // Returns Folder object based on folder path
        static Outlook.Folder GetFolder(string folderPath)
        {
            Console.WriteLine("Looking for: " + folderPath);
            Outlook.Folder folder;
            string backslash = @"\";
            try
            {
                if (folderPath.StartsWith(@"\\"))
                {
                    folderPath = folderPath.Remove(0, 2);
                }
                String[] folders = folderPath.Split(backslash.ToCharArray());
                Outlook.Application Application = new Outlook.Application();
                folder = Application.Session.Folders[folders[0]] as Outlook.Folder;
                if (folder != null)
                {
                    for (int i = 1; i <= folders.GetUpperBound(0); i++)
                    {
                        Outlook.Folders subFolders = folder.Folders;
                        folder = subFolders[folders[i]] as Outlook.Folder;
                        if (folder == null)
                        {
                            return null;
                        }
                    }
                }
                return folder;
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                return null;
            }
        }

    }

}

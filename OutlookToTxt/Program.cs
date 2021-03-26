using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using Microsoft.Office.Interop.Outlook;
using System.Runtime.InteropServices;

using System.IO;
using ShellProgressBar;
using System.CommandLine;
using System.CommandLine.Invocation;

namespace OutlookToTxt
{
    class Program
    {
        static int Main(string[] args)
        {
            var rootCommand = new RootCommand();

            rootCommand.Add(
                new Option<string>("--output")
                {
                    IsRequired = true,
                    Description = "Output folder",
                    Argument = new Argument<string>() { Arity = ArgumentArity.ExactlyOne }
                }
            );


            rootCommand.Add(
                new Option<bool>(
                    "--inbox",
                    getDefaultValue: () => true,
                    description: "Inbox"
                )
            );

            rootCommand.Add(
                new Option<bool>(
                    "--sent",
                    getDefaultValue: () => false,
                    description: "Sent Items"
                )
            );

            rootCommand.Add(
                new Option<string>("--folder") 
                { 
                    Description = "Folder name",
                    Argument = new Argument<string> { Arity = ArgumentArity.OneOrMore } 
                }
                
                ); ;


            rootCommand.Description = "Export an Outlook mailbox as plain text";

            rootCommand.Handler = CommandHandler.Create<string, bool, bool, List<string>>(ReadMailItems);

            return rootCommand.InvokeAsync(args).Result;
        }

        private static void ReadMailItems(string output, bool inbox, bool sent, List<string> folders)
        {
            Application outlookApplication = null;
            NameSpace outlookNamespace = null;

            try
            {
                outlookApplication = new Application();
                outlookNamespace = outlookApplication.GetNamespace("MAPI");

                var mapiFolders = new Dictionary<string, MAPIFolder>();

                if (inbox)
                {
                    mapiFolders.Add("inbox", outlookNamespace.GetDefaultFolder(OlDefaultFolders.olFolderInbox));
                }
            
                if (sent)
                {
                    mapiFolders.Add("sent", outlookNamespace.GetDefaultFolder(OlDefaultFolders.olFolderSentMail));
                }

                foreach (string folder in folders)
                {
                    mapiFolders.Add(folder, outlookNamespace.GetDefaultFolder(OlDefaultFolders.olFolderInbox).Parent.Folders[folder]);
                }

                foreach (var e in mapiFolders)
                {
                    DumpFolder(output, e.Value, e.Key);
                    ReleaseComObject(e.Value);
                }
            }
            catch { }
            finally
            {
                ReleaseComObject(outlookNamespace);
                ReleaseComObject(outlookApplication);

            }
                
        }


        private static void DumpFolder(string output, MAPIFolder folder, string folderName)
        {
            Items mailItems = null;

            try
            {
                mailItems = folder.Items;

                using (var pbar = new ProgressBar(mailItems.Count, "Starting"))
                {
                    foreach (object item in mailItems)
                    {
                        MailItem mail = item as MailItem;
                        if (mail == null) { continue; }
                        try
                        {
                            SaveMessage(output, mail, folderName);
                            pbar.Tick(mail.Subject);
                        }
                        catch { }
                        finally
                        {

                            Marshal.ReleaseComObject(item);
                        }

                    }
                }

            }
            catch { }
            finally
            {
                ReleaseComObject(mailItems);
            }
        }


        private static void SaveMessage(string output, MailItem item, string folder)
        {
            string path = Path.Combine(output, folder);
            string fn = Path.Combine(path, item.EntryID + ".txt");

            try { Directory.CreateDirectory(path); } catch { }

            using (StreamWriter file = new StreamWriter(fn))
            {
                file.WriteLine("From: " + item.SenderName);
                file.WriteLine("To: " + item.To);
                file.WriteLine("CC: " + item.CC);
                file.WriteLine("Subject: " + item.Subject);
                file.WriteLine("Sent: " + item.ReceivedTime.ToString("dddd, dd MMMM yyyy hh:mm tt"));
                file.WriteLine("");
                file.WriteLine(item.Body);
            }
        }



        private static void ReleaseComObject(object obj)
        {
            if (obj != null)
            {
                Marshal.ReleaseComObject(obj);
                obj = null;
            }
        }

    }
}

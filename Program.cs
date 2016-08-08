using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace UploadCSV
{
    using System;
    using System.Collections.Generic;
    using System.Configuration;
    using System.Linq;
    using System.Net;
    using System.Runtime.CompilerServices;
    using System.Text;

    using HtmlAgilityPack;

    using Microsoft.SharePoint;
    class Program
    {
        static void Main(string[] args)
        {
             Console.WriteLine("");
            Console.WriteLine("==================================================");
            Console.WriteLine("");
            Console.WriteLine("  Customer Quality Specification ShareBank Sync");
            Console.WriteLine("");
            Console.WriteLine("==================================================");
            Console.WriteLine("");

            // STEP 1:
            // -------
            // Get everything ready
            Console.Write("  [1] - Preparing to sync");
            var fileUrl = ConfigurationManager.AppSettings["srcFile"];
            var listViewUrl = ConfigurationManager.AppSettings["srcView"];
             string siteCollectionUrl = ConfigurationManager.AppSettings["destSiteCol"];
            var webUrl = ConfigurationManager.AppSettings["destWeb"];
            var libraryUrl = ConfigurationManager.AppSettings["destLib"];
            var filename = fileUrl.TrimEnd('/').Split('/').Last();
            decimal d = 0.0M;
            //var destinationPath =
            //    string.Join("/", new[]
            //    {
            //        siteCollectionUrl.TrimEnd(','),
            //        libraryUrl.Trim('/'),
            //        Uri.EscapeDataString(filename)
            //    });

         string destinationPath = "http://dsl.riotinto.org/documentlibraries/controlleddocuments/Controlled%20Documents%20Library/Marketing%20and%20Sales%20(RH)/" + Uri.UnescapeDataString(filename);

            ConsoleWriteOk();

            // STEP 2:
            // -------
            // Download the page and parse it to extract metadata.
            Console.WriteLine("  [2] - Extracting relevant metadata from ShareBank");
            Console.Write("     [a] - Parsing 'Sales and Admin' document library page");
            var client = new WebClient { UseDefaultCredentials = true };
            var page = client.DownloadData(listViewUrl);
            var html = new HtmlDocument();
            html.LoadHtml(Encoding.UTF8.GetString(page));
            ConsoleWriteOk();

            // Get the document library list view table,
            // and get its headers and rows content
            Console.Write("     [b] - Extracting relevant metadata");
            var table = html.DocumentNode.SelectSingleNode("//table[@id='onetidDoclibViewTbl0']");
            var rows = table.ChildNodes;
            var headers = table.ChildNodes[0].ChildNodes.Where(q => q.Name == "th").ToList();

            // Build a table column mapping to make it easier to extract metadata from list view.
            var columnMap = new Dictionary<string, int>();
            for (var i = 0; i < headers.Count(); i++)
            {
                columnMap.Add(headers[i].InnerText, i);
            }

            // Get document metadata
            var row = rows.First(q => q.InnerHtml.Contains(filename));
            var currentVersion =  row.ChildNodes[columnMap["Version"]].InnerText;
            var lastModified = row.ChildNodes[columnMap["Modified"]].InnerText;
            var lastModifiedBy = row.ChildNodes[columnMap["Modified By"]].InnerText;
            ConsoleWriteOk();
            Console.WriteLine();
           
            // Output results
            Console.WriteLine("           Filename   : {0}", Uri.UnescapeDataString(filename));
            Console.WriteLine("           Version    : {0}", currentVersion);
            Console.WriteLine("           Modified   : {0}", lastModified);
            Console.WriteLine("           Modified By: {0}", lastModifiedBy);
            Console.WriteLine();
            
            // STEP 3:
            // -------
            // Download the file and save it to a temporary location.
            Console.Write("  [3] - Downloading Customer Quality Specification file");
            var data = client.DownloadData(fileUrl);
            ConsoleWriteOk();

            // Upload to document library
            // Loop through PMIS Project Sites and update as required
            
            Console.WriteLine("  [4] - Uploading Customer Quality Specification file");
            Console.Write("     [a] - Connecting to the document library");
            
            
            try
            {
                using (SPSite site = new SPSite(siteCollectionUrl))
                {
                    Console.WriteLine("Web URL path is : {0} ", webUrl);
                    using (SPWeb web = site.OpenWeb(webUrl))
                    {
                        ConsoleWriteOk();

                        Console.Write("     [b] - Uploading the document");
                        var comments = string.Format("{0}: Document Updated From ShareBank", DateTime.Now);
                        
                        try
                        {

                            SPFile exsfile = web.GetFile(destinationPath);
                            if (exsfile.Exists)
                            {
                                Console.WriteLine("File is exist");

                                //Get Current file version exist
                                var descurversion = exsfile.Versions.File.UIVersionLabel;
                                
                                Console.WriteLine(" destination current version : {0}", descurversion);
                                if (currentVersion == descurversion)
                                {
                                    Console.WriteLine(" Version Already exists ");
                                    //Do Nothing and exit the code
                                }
                                else
                                {
                                    //For New Version checkout the code and check-in back
                                    exsfile.CheckOut();
                                    web.Files.Add(destinationPath, data, true, comments, false);
                                    Console.WriteLine("File is added to : {0} ", destinationPath);
                                    
                                    //Check for Major or Minor Version
                                    Decimal.TryParse(currentVersion, out d);
                                    if ((d % 1) == 0)
                                    {
                                        exsfile.CheckIn("New version from CAUPL:  "+currentVersion, SPCheckinType.MajorCheckIn);
                                    }

                                    else
                                    {

                                        exsfile.CheckIn("New version from CAUPL:  "+currentVersion, SPCheckinType.MinorCheckIn);
                                    }
                                }
                            }
                            else
                            {
                                SPFile file = web.Files.Add(destinationPath, data, true, comments, false);
                                Console.WriteLine("File is added to : {0} ", destinationPath);
                                file.CheckIn("New version from CAUPL: "+currentVersion, SPCheckinType.MinorCheckIn);
                            
                            }
                            
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine("Error is  : {0}", ex.Message);
                        }
                        ConsoleWriteOk();
                    }
                }
            }
            catch (Exception e)
            {

                Console.WriteLine("Error : {0} ", e.Message);
            }
            Console.WriteLine("  Done. Press any key to exit...");
            Console.ReadKey();
        }
        /// <summary>
        /// Write an OK mark on the console.
        /// </summary>
        private static void ConsoleWriteOk()
        {
            for (var i = Console.CursorLeft + 2; i <= 75; i++)
            {
                Console.Write("·");
            }

            Console.ForegroundColor = ConsoleColor.Green;
            Console.WriteLine("[OK]");
            Console.ResetColor();

        }
    }
}

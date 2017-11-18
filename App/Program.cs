using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using iTextSharp.text.pdf;
using System.IO;
using iTextSharp.text.pdf.parser;
using System.Text.RegularExpressions;

namespace PdfReader
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Enter the pdf path >");
            var pdfPath = Console.ReadLine();
            ReadPdfFile(pdfPath); // @"C:\Workarea\PdfReader\App\Data\Contacts.pdf");
            Console.ReadLine();
        }

        private static void ReadPdfFile(string fileName)
        {
            
            var pageTitle = "REALTOR Association of Acadiana\nActive RAA Member Roster\n{0}\n";
            var contactResult = new List<Contact>();

            if (File.Exists(fileName))
            {
                Console.WriteLine("Processing {0}", fileName);
                var pdfReader = new iTextSharp.text.pdf.PdfReader(fileName);
                var partialContactContent = string.Empty;
                for (int page = 1; page <= pdfReader.NumberOfPages; page++)
                {
                    ITextExtractionStrategy strategy = new SimpleTextExtractionStrategy();
                    string pageContents = PdfTextExtractor.GetTextFromPage(pdfReader, page, strategy);
                    pageContents = Encoding.UTF8.GetString(ASCIIEncoding.Convert(Encoding.Default, Encoding.UTF8, Encoding.Default.GetBytes(pageContents)));
                    pageContents = pageContents.Replace(string.Format(pageTitle, page), "");
                    pageContents = partialContactContent + "\n" + pageContents; // append previous page partial contact details.
                    contactResult.AddRange(ReadPdfContents(pageContents));
                    partialContactContent = FindPartialContact(pageContents);
                }
                pdfReader.Close();

                ExportToText(contactResult);
            }
            else
            {
                Console.WriteLine("File not found!");
            }
        }

        private static void ExportToText(List<Contact> contacts)
        {
            var outputFileName = string.Format("Contacts_{0}.txt", DateTime.Now.ToString("yyyy-dd-MM_HH-mm-ss"));
            using (TextWriter tw = new StreamWriter(outputFileName))
            {
                tw.WriteLine("Email\tLastName\tFirstName\tOfficePhone\tCellPhone\tCompanyName");
                foreach (var contact in contacts)
                {
                    tw.WriteLine(string.Format("{0}\t{1}\t{2}\t{3}\t{4}\t{5}", 
                        contact.Email, contact.LastName, contact.FirstName, contact.WorkPhone, 
                        (!string.IsNullOrWhiteSpace(contact.PrefPhone) ? contact.PrefPhone : contact.CellPhone), 
                        contact.CompanyName));
                }
            }
            Console.WriteLine(@"Contacts exported to {0}\{1}", AppDomain.CurrentDomain.BaseDirectory, outputFileName);
        }
    
        private static string FindPartialContact(string pageContents)
        {
            Regex rx = new Regex(@"\w+([-+.']\w+)*@\w+([-.]\w+)*\.\w+([-.]\w+)*");
            var matches = rx.Matches(pageContents);
            var lastContactEmailMatch = matches[matches.Count - 1];
            var startIndex = lastContactEmailMatch.Index + lastContactEmailMatch.Value.Length;
            var remainingContent = pageContents.Substring(startIndex);
            return remainingContent;
        }

        private static List<Contact> ReadPdfContents(string content)
        {
            Regex rx = new Regex(@"\w+([-+.']\w+)*@\w+([-.]\w+)*\.\w+([-.]\w+)*");
            var contacts = new List<Contact>();
            var startIndex = 0;

            foreach (Match match in rx.Matches(content))
            {
                int i = match.Index;
                var endIndex = i + match.Value.Length;
                var ct = content.Substring(startIndex, endIndex - startIndex);
                startIndex = endIndex;
                contacts.Add(ExtractValues(ct, match.Value));
            }
            return contacts;
        }
        private static Contact ExtractValues(string content, string email)
        {

            var contact = new Contact { Email = email };
            var splits = content.Split(new[] { "\n" }, StringSplitOptions.RemoveEmptyEntries);
            var name = string.Empty;
            if (Regex.IsMatch(splits[0], "[0-9]{3}-[0-9]{3}-[0-9]{4}"))
            {
                // row with lastname, firstname workphone
                var workPhoneregex = new Regex(@"[0-9]{3}-[0-9]{3}-[0-9]{4}");
                var match = workPhoneregex.Match(splits[0]);
                contact.WorkPhone = match.Value;
                name = splits[0].Substring(0, match.Index);
            }
            else
            {
                name = splits[0];
            }
            FillLastAndFirstName(name, contact);

            if (!FetchPhone(splits[1], contact)) // Work Phone
            {
                // its company name
                contact.CompanyName = splits[1];
                // you are done with the contact. no phones available return.
                return contact;
            }
            if (!FetchPhone(splits[2], contact)) // Cell Phone
            {
                contact.CompanyName = splits[2];
                // you are done with the contact. no phones available return.
                return contact;
            }
            if (!FetchPhone(splits[3], contact)) // Pref Phone
            {
                contact.CompanyName = splits[3];
            }

            return contact;
        }

        private static void FillLastAndFirstName(string name, Contact contact)
        {
            var splits = name.Split(new[] { ", " }, StringSplitOptions.RemoveEmptyEntries);
            contact.LastName = splits[0];
            contact.FirstName = splits.Length > 1 ? splits[1] : string.Empty;
        }
        private static bool FetchPhone(string content, Contact contact)
        {
            var findPhoneMatch = new Regex(@"[0-9]{3}-[0-9]{3}-[0-9]{4}");
            var checkMatch = findPhoneMatch.Match(content);
            if (checkMatch.Success)
            {
                // phone
                if (content.Contains("Cell"))
                {
                    contact.CellPhone = checkMatch.Value;
                }
                else if (content.Contains("Pref"))
                {
                    contact.PrefPhone = checkMatch.Value;
                }
                else
                {
                    contact.WorkPhone = checkMatch.Value;
                }
            }
            return checkMatch.Success;
        }
    }

    public class Contact
    {
        public string LastName { get; set; }
        public string FirstName { get; set; }
        public string WorkPhone { get; set; }
        public string CellPhone { get; set; }
        public string PrefPhone { get; set; }
        public string CompanyName { get; set; }
        public string Email { get; set; }
    }
}

// Minimal Test for ClosedXML Email
// Platform:  Visual Studio 2015, .NET 4.6.1
// Included:  ClosedXML v0.87.1 from NuGet
// Apr 18, 2017 released to public domain
using ClosedXML.Excel;
using System;
using System.IO;
using System.Linq;
using System.Net.Mail;
using System.Net.Mime;
using System.Text;

namespace ClosedXML_Email_Test01
{
  class Program
  {
    static void Main(string[] args)
    {
      Console.WriteLine("Test sending ClosedXML XLSX file as an email attachment ...");

      // Create the XLSX workbook file
      string myBookname = "Mars.xlsx";
      CreateXlBook(myBookname);

      // Send the XLSX file by email
      string recipientAddress = "kitty@gmail.com";
      string senderAddress = "astro@planet.earth";
      SendXlBookEmail(myBookname, recipientAddress, senderAddress);

      Console.WriteLine("Done!");
      Console.Write("Press [Enter] to Exit -> ");
      string exitMessage = Console.ReadLine();
    }


    private static void CreateXlBook(string bookFilename)
    {
      // Initialize a new Workbook
      using (XLWorkbook myXlBook = new XLWorkbook())
      {
        string mySheetname = "VENUS";
        using (IXLWorksheet myXlSheet = myXlBook.Worksheets.Add(mySheetname))
        {
          try
          {
            // Check to see if a specific worksheet exists
            if (myXlBook.Worksheets.Where(w => w.Name == mySheetname).Count() == 1)
            {
              Console.WriteLine(" + worksheet created = " + mySheetname);
            }

            // Add worksheet content ...
            myXlSheet.Cell(1, 1).Value = "Hello World!";

            // Save the Workbook
            if (File.Exists(bookFilename))
            {
              Console.Write(" - deleting old file ... ");
              File.Delete(bookFilename);
              Console.WriteLine(" old file deleted.");
            }
            Console.Write(" * Saving XL Workbook -> " + bookFilename + " ... ");
            myXlBook.SaveAs(bookFilename);
            Console.WriteLine("Workbook Saved!");
          }
          catch (Exception xlException)
          {
            Console.WriteLine("XLSX Construction or Save Error -> " + xlException.Message);
            if (xlException.InnerException != null)
            {
              Console.WriteLine(" --> " + xlException.InnerException.Message);
            }
          }
        } // dispose using IXLWorksheet myXlSheet
      } // dispose using XLWorkbook myXlBook
    }


    private static void SendXlBookEmail(string bookFilename, string emailTo, string emailFrom)
    {
      Console.WriteLine("Preparing Email from " + emailFrom + " to " + emailTo + " ...");

      MailAddress mailFromAddress = new MailAddress(emailFrom, emailFrom);
      MailAddress mailToAddress = new MailAddress(emailTo, emailTo);

      MailMessage mailMess = new MailMessage(mailFromAddress, mailToAddress);
      SmtpClient mailClient = new SmtpClient();
      try
      {
        mailMess.SubjectEncoding = Encoding.UTF8;
        mailMess.Subject = "ClosedXML Worksheets Attached # " + DateTime.Now.ToString("HH:mm:ss MM/dd/yyyy");
        mailMess.BodyEncoding = Encoding.UTF8;
        mailMess.Body = "Can you view my worksheets?  one is mime-type octet, the other is openxml.";
        mailMess.IsBodyHtml = false;

        Console.WriteLine(" + Encoding Attachment Files ...");
        // Create the file attachments for this e-mail message
        //   -> encode as application/octet
        string filenameOctet = bookFilename.Replace(".xlsx", "-octet.xlsx");
        if (File.Exists(filenameOctet))
        {
          File.Delete(filenameOctet);
        }
        File.Copy(bookFilename, filenameOctet);
        Attachment attFileOctet = new Attachment(filenameOctet, 
                                                 MediaTypeNames.Application.Octet);
        attFileOctet.TransferEncoding = TransferEncoding.Base64;

        //   -> encode as openxml
        string filenameOpenxml = bookFilename.Replace(".xlsx", "-openxml.xlsx");
        if (File.Exists(filenameOpenxml))
        {
          File.Delete(filenameOpenxml);
        }
        File.Copy(bookFilename, filenameOpenxml);
        Attachment attFileOpenXml = new Attachment(filenameOpenxml,
                                                   new ContentType("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"));
        attFileOpenXml.TransferEncoding = TransferEncoding.Base64;

        // Add the attachments to the message
        mailMess.Attachments.Add(attFileOctet);
        mailMess.Attachments.Add(attFileOpenXml);

        // Setup the SMTP Server endpoint
        mailClient.Host = "mailster.planet.earth";
        mailClient.Port = 25;
        mailClient.DeliveryMethod = SmtpDeliveryMethod.Network;
        mailClient.EnableSsl = false;
        mailClient.Credentials = new System.Net.NetworkCredential(emailFrom, "");

        Console.Write(" * Sending Email with Attachments ... ");
        mailClient.Send(mailMess);
        Console.WriteLine("Email Sent!");

        if (mailMess.Attachments.Count > 0)
        {
          // Must Dispose the Attachment(s) or else .NET Process will lock the file(s).
          foreach (Attachment att in mailMess.Attachments)
          {
            att.Dispose();
          }
          mailMess.Attachments.Clear();
        }
      }
      catch (Exception emailException)
      {
        Console.WriteLine("Email Error -> " + emailException.Message);
        if (emailException.InnerException != null)
        {
          Console.WriteLine(" --> " + emailException.InnerException.Message);
        }
      }

      // Clean up
      if (mailMess != null)
      {
        mailMess.Dispose();
      }
      if (mailClient != null)
      {
        mailClient.Dispose();
      }
    }

  }
}

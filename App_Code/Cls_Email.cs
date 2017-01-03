using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Net.Mail;

public class Cls_Email
{
    public Cls_Email()
    {
        //
        // TODO: Add constructor logic here
        //
    }
    public static string GetNextEmailFileName()
    {
        return (AppDomain.CurrentDomain.BaseDirectory + (@"Emails\SavedMail") + GetNextEmailNumber().ToString() + @".eml");
    }
    public static int GetNextEmailNumber()
    {
        string str_NextEmailNumber = Cls_Gen.GetSettings("NextEmailIndex");
        int int_NextEmailNumber = Convert.ToInt32(str_NextEmailNumber);
        return int_NextEmailNumber;
    }
    public static void IncrementEmailNumberinSettings()
    {
        string str_NextEmailNumber = Cls_Gen.GetSettings("NextEmailIndex");
        int int_NextEmailNumber = Convert.ToInt32(str_NextEmailNumber);
        int_NextEmailNumber++;
        Cls_Gen.SetSettings("NextEmailIndex", int_NextEmailNumber.ToString());        
    }
    public static void Chp50_SendingEmailswithGmail()
    {
        //This requires "using System.Net.Mail"
        /* Activate with:
            Templates objTemplate = new Templates();
            objTemplate.Chp21_SendingEmails();
         */
        try
        {
            MailMessage mail = new MailMessage();
            SmtpClient SmtpServer = new SmtpClient("smtp.gmail.com");

            mail.From = new MailAddress("daniel.sk.sim@gmail.com");
            mail.To.Add("cstechassist@gmail.com");
            mail.Subject = "C# Test Mail sent on " + DateTime.Now;
            mail.IsBodyHtml = true; //Set this to false if sending text message.
            mail.Body = "This is an experiment message sent on behalf of daniel.sk.sim@gmail.com, by C# template robotic assistant on " + DateTime.Now + "\n" + "End of message.";

            System.Net.Mail.Attachment attachment;
            attachment = new System.Net.Mail.Attachment(@"C:\Users\User\Documents\Visual Studio 2012\Projects\Beginner\Beginner\DocXExample.docx");
            mail.Attachments.Add(attachment);

            SmtpServer.Port = 587;
            SmtpServer.Credentials = new System.Net.NetworkCredential("daniel.sk.sim@gmail.com", "i9mrqci9mrqc");
            SmtpServer.EnableSsl = true;
            SmtpServer.Send(mail);
            System.Diagnostics.Debug.WriteLine("Mail Sent!");
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine(ex.ToString());
        }

    }
    public static void Chp50a_SendingEmailwithExchange()
    {
        //This requires "using System.Net.Mail"
        /* Activate with:
            Templates objTemplate = new Templates();
            objTemplate.Chp21_SendingEmails();
         */
        try
        {
            System.Net.Mail.MailMessage mail = new System.Net.Mail.MailMessage();
            System.Net.Mail.SmtpClient SmtpServer = new System.Net.Mail.SmtpClient("smtp.office365.com", 587);

            mail.From = new System.Net.Mail.MailAddress("xxx@domain.com.sg");
            mail.To.Add("cstechassist@gmail.com");
            mail.Subject = "C# Test Mail sent on " + DateTime.Now;
            mail.IsBodyHtml = true; //Set this to false if sending text message.
            mail.Body = "This is an experiment message sent on behalf of xxx@domain.com.sg, by C# template robotic assistant on " + DateTime.Now + "\n" + "End of message.";

            //System.Net.Mail.Attachment attachment;
            //attachment = new System.Net.Mail.Attachment(@"C:\Users\User\Documents\Visual Studio 2012\Projects\Beginner\Beginner\DocXExample.docx");
            //mail.Attachments.Add(attachment);

            SmtpServer.UseDefaultCredentials = false;
            SmtpServer.Credentials = new System.Net.NetworkCredential("xxx@domain", "password");
            SmtpServer.EnableSsl = true;
            SmtpServer.Send(mail);
            System.Diagnostics.Debug.WriteLine("Mail Sent!");
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine(ex.ToString());
        }

    }
    static public List<OpenPop.Mime.Message> Chp51_ReadingEmails(string username = @"cstechassist@gmail.com", string password = "i9mrqci9mrqc", string hostname = @"pop.gmail.com", int port = 995, bool useSsl = true)
    {
        /// <summary>
        /// Example showing how to fetch all messages from a POP3 server
        /// </summary>
        /// <param name="hostname">Hostname of the server. For example: pop3.live.com</param>
        /// <param name="port">Host port to connect to. Normally: 110 for plain POP3, 995 for SSL POP3</param>
        /// <param name="useSsl">Whether or not to use SSL to connect to server</param>
        /// <param name="username">Username of the user on the server</param>
        /// <param name="password">Password of the user on the server</param>
        /// <returns>All Messages on the POP3 server</returns>

        //To use this, must add nuget package OpenPop.Net
        //This function gets all the mail from daniel.sk.sim@gmail.com mailbox and prints the subject
        //Beware, this function takes long to complete.

        // The client disconnects from the server when being disposed
        OpenPop.Pop3.Pop3Client client = new OpenPop.Pop3.Pop3Client();
        // Connect to the server
        client.Connect(hostname, port, useSsl);

        // Authenticate ourselves towards the server
        client.Authenticate(username, password);

        // Get the number of messages in the inbox
        int messageCount = client.GetMessageCount();

        // We want to download all messages
        List<OpenPop.Mime.Message> allMessages = new List<OpenPop.Mime.Message>(messageCount);

        // Messages are numbered in the interval: [1, messageCount]
        // Ergo: message numbers are 1-based.
        // Most servers give the latest message the highest number
        for (int i = messageCount; i > 0; i--)
        {
            allMessages.Add(client.GetMessage(i));
        }

        //foreach (OpenPop.Mime.Message msg in allMessages)
        //{
        //    System.Diagnostics.Debug.Print(msg.Headers.Subject);
        //    //System.Diagnostics.Debug.Print(Chp51d_GetTextInMessage(msg));
        //    //Chp51e_GetAttachmentInMessage(msg);
        //}

        client.Dispose();
        return allMessages;
    }

    /// <summary>
    /// Example showing:
    ///  - how to fetch only headers from a POP3 server
    ///  - how to examine some of the headers
    ///  - how to fetch a full message
    ///  - how to find a specific attachment and save it to a file
    /// </summary>
    /// <param name="hostname">Hostname of the server. For example: pop3.live.com</param>
    /// <param name="port">Host port to connect to. Normally: 110 for plain POP3, 995 for SSL POP3</param>
    /// <param name="useSsl">Whether or not to use SSL to connect to server</param>
    /// <param name="username">Username of the user on the server</param>
    /// <param name="password">Password of the user on the server</param>
    /// <param name="messageNumber">
    /// The number of the message to examine.
    /// Must be in range [1, messageCount] where messageCount is the number of messages on the server.
    /// </param>
    public static void HeadersFromAndSubject(int messageNumber, string hostname = @"pop.gmail.com", int port = 995, bool useSsl = true, string username = @"cstechassist@gmail.com", string password = "i9mrqci9mrqc")
    {
        // The client disconnects from the server when being disposed
        using (OpenPop.Pop3.Pop3Client client = new OpenPop.Pop3.Pop3Client())
        {
            // Connect to the server
            client.Connect(hostname, port, useSsl);

            // Authenticate ourselves towards the server
            client.Authenticate(username, password);

            // We want to check the headers of the message before we download
            // the full message
            OpenPop.Mime.Header.MessageHeader headers = client.GetMessageHeaders(messageNumber);

            OpenPop.Mime.Header.RfcMailAddress from = headers.From;
            string subject = headers.Subject;
            System.Diagnostics.Debug.WriteLine(subject);
            // Only want to download message if:
            //  - is from test@xample.com
            //  - has subject "Some subject"
            if (from.HasValidMailAddress && from.Address.Equals("test@example.com") && "Some subject".Equals(subject))
            {
                // Download the full message
                OpenPop.Mime.Message message = client.GetMessage(messageNumber);

                // We know the message contains an attachment with the name "useful.pdf".
                // We want to save this to a file with the same name
                foreach (OpenPop.Mime.MessagePart attachment in message.FindAllAttachments())
                {
                    if (attachment.FileName.Equals("useful.pdf"))
                    {
                        // Save the raw bytes to a file                            
                        System.IO.File.WriteAllBytes(attachment.FileName, attachment.Body);
                    }
                }
            }
        }
    }
    static public int GetMessageCount(string hostname = @"pop.gmail.com", int port = 995, bool useSsl = true, string username = @"cstechassist@gmail.com", string password = "i9mrqci9mrqc")
    {
        // The client disconnects from the server when being disposed
        OpenPop.Pop3.Pop3Client client = new OpenPop.Pop3.Pop3Client();
        // Connect to the server
        client.Connect(hostname, port, useSsl);

        // Authenticate ourselves towards the server
        client.Authenticate(username, password);

        // Get the number of messages in the inbox
        return (client.GetMessageCount());
    }
    static public OpenPop.Mime.Message Chp51a_ReadLatestMail()
    {
        //To use this, must add nuget package OpenPop.Net
        //This function gets all the mail from daniel.sk.sim@gmail.com mailbox and prints the subject
        //Returns the latest message if there is one, else return null
        string hostname = @"pop.gmail.com";
        int port = 995;
        bool useSsl = true;
        string username = @"cstechassist@gmail.com";
        string password = "i9mrqci9mrqc";
        OpenPop.Mime.Message latestmsg = null;
        // The client disconnects from the server when being disposed
        OpenPop.Pop3.Pop3Client client = new OpenPop.Pop3.Pop3Client();
        // Connect to the server
        client.Connect(hostname, port, useSsl);

        // Authenticate ourselves towards the server
        client.Authenticate(username, password);

        // Get the number of messages in the inbox
        int messageCount = client.GetMessageCount();
        System.Diagnostics.Debug.WriteLine("Message Count = " + messageCount);
        // Messages are numbered in the interval: [1, messageCount]
        // Ergo: message numbers are 1-based.
        // Most servers give the latest message the highest number
        if (messageCount != 0)
        {
            latestmsg = client.GetMessage(messageCount);
            Chp51b_SaveFullMessage(latestmsg, "savedmail.eml");
            System.Diagnostics.Debug.Print(latestmsg.Headers.Subject + latestmsg.Headers.DateSent);
        }

        client.Dispose();
        return latestmsg;

    }

    public static string Chp51b_SaveFullMessage(OpenPop.Mime.Message message, string str_FileName)
    {
        /// <summary>
        /// Example showing:
        ///  - how to save a message to a file
        ///  - how to load a message from a file at a later point
        /// </summary>
        /// <param name="message">The message to save and load at a later point</param>
        /// <returns>The Message, but loaded from the file system</returns>
        // FileInfo about the location to save/load message\
        // For illustrative purpose, the message is retrieved. In actual application, 
        // can include message as a parameter to make it more flexible.

        System.IO.FileInfo file = new System.IO.FileInfo(str_FileName);

        // Save the full message to some file
        if (message != null)
        {
            message.Save(file);
            return "Success";
        }
        else
        {
            return "Failed";
        }

    }
    public static OpenPop.Mime.Message Chp51b_LoadFullMessage(string str_FileName)
    {
        System.IO.FileInfo file = new System.IO.FileInfo(str_FileName);
        return OpenPop.Mime.Message.Load(file);
    }

    public static void Chp51c_DeleteMessageOnServer(int messageNumber)
    {
        /// <summary>
        /// Example showing:
        ///  - how to delete a specific message from a server
        /// </summary>
        /// <param name="hostname">Hostname of the server. For example: pop3.live.com</param>
        /// <param name="port">Host port to connect to. Normally: 110 for plain POP3, 995 for SSL POP3</param>
        /// <param name="useSsl">Whether or not to use SSL to connect to server</param>
        /// <param name="username">Username of the user on the server</param>
        /// <param name="password">Password of the user on the server</param>
        /// <param name="messageNumber">
        /// The number of the message to delete.
        /// Must be in range [1, messageCount] where messageCount is the number of messages on the server.
        /// </param>
        // The client disconnects from the server when being disposed

        string hostname = @"pop.gmail.com";
        int port = 995;
        bool useSsl = true;
        string username = @"cstechassist@gmail.com";
        string password = "i9mrqci9mrqc";

        using (OpenPop.Pop3.Pop3Client client = new OpenPop.Pop3.Pop3Client())
        {
            // Connect to the server
            client.Connect(hostname, port, useSsl);

            // Authenticate ourselves towards the server
            client.Authenticate(username, password);

            // Mark the message as deleted
            // Notice that it is only MARKED as deleted
            // POP3 requires you to "commit" the changes
            // which is done by sending a QUIT command to the server
            // You can also reset all marked messages, by sending a RSET command.
            client.DeleteMessage(messageNumber);

            // When a QUIT command is sent to the server, the connection between them are closed.
            // When the client is disposed, the QUIT command will be sent to the server
            // just as if you had called the Disconnect method yourself.
        }
    }

    public static string Chp51d_GetTextInMessage(OpenPop.Mime.Message message)
    {
        string Body = "";
        OpenPop.Mime.MessagePart plainText = message.FindFirstPlainTextVersion();
        if (plainText != null)
        {
            // We found some plaintext!
            Body = plainText.GetBodyAsText();

        }
        else
        {
            // Might include a part holding html instead
            OpenPop.Mime.MessagePart html = message.FindFirstHtmlVersion();
            if (html != null)
            {
                // We found some html!
                Body = html.GetBodyAsText();
            }
        }
        return Body;
    }

    public static void Chp51e_GetAttachmentInMessage(OpenPop.Mime.Message message)
    {
        string _path = "";
        foreach (OpenPop.Mime.MessagePart attachment in message.FindAllAttachments())
        {
            string filePath = System.IO.Path.Combine(_path, attachment.FileName);

            System.IO.FileStream Stream = new System.IO.FileStream(filePath, System.IO.FileMode.Create);
            System.IO.BinaryWriter BinaryStream = new System.IO.BinaryWriter(Stream);
            BinaryStream.Write(attachment.Body);
            Stream.Close();
            BinaryStream.Close();

        }
    }
}
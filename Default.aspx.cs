using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

public partial class _Default : System.Web.UI.Page
{
    List<OpenPop.Mime.Message> allMessages;

    protected void Page_Load(object sender, EventArgs e)
    {
      PageLoaderFunction();
    }

    protected void PageLoaderFunction()
    {
        p_Status.InnerText = "Status text";
        allMessages = Cls_Email.Chp51_ReadingEmails();
        foreach (OpenPop.Mime.Message msg in allMessages)
        {
            if (Cls_Email.Chp51b_SaveFullMessage(msg,Cls_Email.GetNextEmailFileName()) == "Success")
            {
                Cls_Email.IncrementEmailNumberinSettings();
                p_Status.InnerText = "Saving email successful";
            }           
        }
    }
    protected void btn_Test_Click(object sender, EventArgs e)
    {
        p_Status.InnerText = "1.43pm";
            
    }



}
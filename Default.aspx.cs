﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

public partial class _Default : System.Web.UI.Page
{
    protected void Page_Load(object sender, EventArgs e)
    {
        p_Status.InnerText = "This is a start";
    }
    protected void btn_Test_Click(object sender, EventArgs e)
    {
        p_Status.InnerText = "Test activated";
    }
}
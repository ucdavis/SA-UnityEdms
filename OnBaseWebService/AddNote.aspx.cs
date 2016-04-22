﻿using Hyland.Unity;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Web;
using System.Web.Helpers;
using System.Web.UI;
using System.Web.UI.WebControls;

namespace UCDavis.StudentAffairs.OnBaseWebService
{
    public partial class AddNote : System.Web.UI.Page
    {
        protected void Page_Load(object sender, EventArgs e)
        {
            Response.Clear();
            Response.ContentType = "application/json;charset=utf-8";
            Response.ContentEncoding = Encoding.UTF8;
            Response.Write(Json.Encode(ServiceWrapper.AddNote(Request)));
            Response.End();

        }
    }
}
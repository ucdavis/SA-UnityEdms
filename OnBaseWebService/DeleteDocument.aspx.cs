﻿using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

namespace UCDavis.StudentAffairs.OnBaseWebService
{
    public partial class DeleteDocument : System.Web.UI.Page
    {
        protected void Page_Load(object sender, EventArgs e)
        {
            Response.Clear();
            Response.ContentType = "application/json;charset=utf-8";
            Response.ContentEncoding = Encoding.UTF8;
            Response.Write(JsonConvert.SerializeObject(ServiceWrapper.DeleteDocument(Request)));
            Response.End();
        }
    }
}
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Web.Helpers;
using System.Configuration;
using System.IO;
using System.Text;

namespace UCDavis.StudentAffairs.OnBaseWebService
{
    public partial class AddDocument : System.Web.UI.Page
    {
        protected void Page_Load(object sender, EventArgs e)
        {
            Response.Clear();
            Response.ContentType = "application/json;charset=utf-8";
            Response.ContentEncoding = Encoding.UTF8;
            Response.Write(Json.Encode(ServiceWrapper.AddDocument(Request)));
            Response.End();
        }
    }
}
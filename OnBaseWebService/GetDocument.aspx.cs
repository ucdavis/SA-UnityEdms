using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Web;
using System.Web.Helpers;
using System.Web.Script.Serialization;
using System.Web.UI;
using System.Web.UI.WebControls;

namespace UCDavis.StudentAffairs.OnBaseWebService
{
    public partial class GetDocument : System.Web.UI.Page
    {
        protected void Page_Load(object sender, EventArgs e)
        {
            Response.Clear();
            Response.ContentType = "application/json;charset=utf-8";
            Response.ContentEncoding = Encoding.UTF8;
            JavaScriptSerializer serializer = new JavaScriptSerializer();
            
            serializer.MaxJsonLength = int.MaxValue;            
            Response.Write(serializer.Serialize(ServiceWrapper.GetDocument(Request)));
            Response.End();
        }
    }
}
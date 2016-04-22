using System;
using System.Text;
using System.Web.Helpers;

namespace UCDavis.StudentAffairs.OnBaseWebService
{
    public partial class GetDocumentTypeDetails : System.Web.UI.Page
    {
        protected void Page_Load(object sender, EventArgs e)
        {
            Response.Clear();
            Response.ContentType = "application/json;charset=utf-8";
            Response.ContentEncoding = Encoding.UTF8;
            Response.Write(Json.Encode(ServiceWrapper.GetDocumentTypeDetails(Request)));
            Response.End();
        }
    }
}
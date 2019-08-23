using Newtonsoft.Json;
using System;
using System.Text;

namespace UCDavis.StudentAffairs.OnBaseWebService
{
    public partial class GetDocumentTypeDetails : System.Web.UI.Page
    {
        protected void Page_Load(object sender, EventArgs e)
        {
            Response.Clear();
            Response.ContentType = "application/json;charset=utf-8";
            Response.ContentEncoding = Encoding.UTF8;
            Response.Write(JsonConvert.SerializeObject(ServiceWrapper.GetDocumentTypeDetails(Request)));
            Response.End();
        }
    }
}
using Hyland.Unity;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

namespace UCDavis.StudentAffairs.OnBaseWebService
{
    public partial class UpdateDocumentTypeKeys : System.Web.UI.Page
    {
        protected void Page_Load(object sender, EventArgs e)
        {
            Response.Clear();
            Response.ContentType = "application/json";
            Response.ContentEncoding = Encoding.UTF8;
            Response.Write(JsonConvert.SerializeObject(ServiceWrapper.UpdateDocumentTypeKeys(Request)));
            Response.End();
        }        
    }
}
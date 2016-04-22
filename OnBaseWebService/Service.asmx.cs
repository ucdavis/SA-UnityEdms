using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.Script.Services;
using System.Web.Services;
using System.Web.Helpers;
using System.Configuration;

namespace UCDavis.StudentAffairs.OnBaseWebService
{
    /// <summary>
    /// Summary description for Service
    /// </summary>
    [WebService]
    [WebServiceBinding(ConformsTo = WsiProfiles.BasicProfile1_1)]
    [System.ComponentModel.ToolboxItem(false)]
    // To allow this Web Service to be called from script, using ASP.NET AJAX, uncomment the following line. 
    // [System.Web.Script.Services.ScriptService]
    public class Service : System.Web.Services.WebService
    {

        [WebMethod]
        public DocumentContract GetDocument(string UserName, string Password, string ActualUser, string ActualUserIP,
            long DocID, string AlternateFormat, int? applicationid, string appspecificid)
        {
            return ServiceWrapper.GetDocument(UserName, Password, Context.Request.UserHostAddress,
                ActualUser, ActualUserIP, DocID, AlternateFormat, applicationid, appspecificid);
        }

        [WebMethod]
        public DocumentTypeContract GetDocumentTypeDetails(string UserName, string Password, string ActualUser, string ActualUserIP,
            long DocTypeID, int? applicationid, string appspecificid)
        {
            return ServiceWrapper.GetDocumentTypeDetails(UserName, Password, Context.Request.UserHostAddress,
                ActualUser, ActualUserIP, DocTypeID, applicationid, appspecificid);
        }


        [WebMethod]
        public OnBaseContract DeleteDocument(string UserName, string Password, string ActualUser, string ActualUserIP,
            long DocID, int? applicationid, string appspecificid)
        {
            return ServiceWrapper.DeleteDocument(UserName, Password, Context.Request.UserHostAddress,
                ActualUser, ActualUserIP, DocID, applicationid, appspecificid);
        }


        [WebMethod]
        public UserDetailsContract GetUserDetails(string UserName, string Password, string ActualUser, string ActualUserIP,
            string LookupAccountName, int ? applicationid, string appspecificid)
        {
            return ServiceWrapper.GetUserDetails(UserName, Password, Context.Request.UserHostAddress, 
                ActualUser, ActualUserIP, LookupAccountName, applicationid, appspecificid);            
        }

        [WebMethod]
        public AddNoteContract AddNote(string UserName, string Password, string ActualUser, string ActualUserIP, 
            long DocID, long NoteType, string Text, long? Width, long? Height, long? XPosition, long? YPosition, long? PageNumber, int ? applicationId, string appSpecificId)
        {
            return ServiceWrapper.AddNote(UserName, Password, Context.Request.UserHostAddress, ActualUser, ActualUserIP,
                DocID, NoteType, Text, Width, Height, XPosition, YPosition, PageNumber, applicationId, appSpecificId);
        }

        [WebMethod]
        public AddDocumentContract AddDocument(string username, string password, string actualUser, string actualIP,
            long DocType, List<byte[]> FileContent, List<string> FileExtensions,
            List<string> Keys, List<string> KeyValues, int? applicationId, string appSpecificId)
        {
            return ServiceWrapper.AddDocument(username, password, Context.Request.UserHostAddress, actualUser, actualIP,
                DocType, FileContent, FileExtensions, Keys, KeyValues, applicationId, appSpecificId);
        }

        [WebMethod]
        public OnBaseContract UpdateDocumentTypeKeys(string username, string password, string actualUser, string actualIP,
            long DocID, long? DocType, List<string> Keys, List<string> KeyValues, int? applicationId, string appSpecificId)
        {
            return ServiceWrapper.UpdateDocumentTypeKeys(username, password, Context.Request.UserHostAddress, actualUser, actualIP,
                DocID, DocType, Keys, KeyValues, applicationId, appSpecificId);
        }
       
    }
}

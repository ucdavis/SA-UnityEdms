using Hyland.Unity;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Web;
using System.Runtime.Serialization;
using System.Data.SqlClient;

namespace UCDavis.StudentAffairs.OnBaseWebService
{
    [Serializable]
    [DataContract]
    public class OnBaseContract
    {
        public OnBaseContract() { }

        [DataMember]
        public bool Success { get; set; }
        [DataMember]
        public string ErrorMessage { get; set; }
        [DataMember]
        public string ErrorDetails { get; set; }
    }
    [Serializable]
    [DataContract]
    public class AddDocumentContract : OnBaseContract
    {
        public AddDocumentContract() { }
        public AddDocumentContract(Document aDocument)
        {
            Success = true;
            DocID = aDocument.ID;
        }

        [DataMember]
        public long DocID { get; set; }
    }

    [Serializable]
    [DataContract]
    public class AddNoteContract : OnBaseContract
    {
        public AddNoteContract() { }
        public AddNoteContract(Note aNote)
        {
            Success = true;
            NoteID = aNote.ID;
        }

        [DataMember]
        public long NoteID { get; set; }
    }

    [Serializable]
    [CollectionDataContract(Name = "GroupMembership", ItemName = "GroupName")]
    public class GroupMembershipContract : List<string>
    {
        public GroupMembershipContract() { }
        public GroupMembershipContract(IEnumerable<UserGroup> lGroups)
            : base(lGroups.Select(x => x.Name))
        { }
    }
    [Serializable]
    [DataContract(Name = "UserDetails")]
    public class UserDetailsContract : OnBaseContract
    {
        public UserDetailsContract() { }
        public UserDetailsContract(User aUser)
        {
            Success = true;
            RealName = aUser.RealName;
            DisplayName = aUser.DisplayName;
            EmailAddress = aUser.EmailAddress;
            GroupMembership = new GroupMembershipContract(aUser.GetUserGroups());
        }
        [DataMember]
        public string DisplayName { get; set; }
        [DataMember]
        public GroupMembershipContract GroupMembership { get; set; }
        [DataMember]
        public string RealName { get; set; }
        [DataMember]
        public string EmailAddress { get; set; }
    }
    [Serializable]
    [DataContract(Name = "Keyword")]
    public class KeywordContract
    {
        public KeywordContract() { }
        public KeywordContract(Keyword aKeyword)
        {
            Value = aKeyword.ToString();
            DataType = (int)aKeyword.KeywordType.DataType;
            DataTypeName = Enum.GetName(typeof(KeywordDataType), aKeyword.KeywordType.DataType);
            TypeName = aKeyword.KeywordType.Name;
            TypeId = aKeyword.KeywordType.ID;
            if (aKeyword.KeywordType.CurrencyFormat != null)
                CurrencyFormatId = aKeyword.KeywordType.CurrencyFormat.ID;
            else CurrencyFormatId = 0;
        }
        [DataMember]
        public string TypeName { get; set; }
        [DataMember]
        public string Value { get; set; }
        [DataMember]
        public long TypeId { get; set; }
        [DataMember]
        public int DataType { get; set; }
        [DataMember]
        public string DataTypeName { get; set; }
        [DataMember]
        public long? CurrencyFormatId { get; set; }
    }
    [Serializable]
    [CollectionDataContract(Name = "Keywords", ItemName = "Keyword")]
    public class KeywordListContract : List<KeywordContract>
    {
        public KeywordListContract() { }
        public KeywordListContract(IEnumerable<Keyword> aKeywords)
            : base(aKeywords.Select(x => new KeywordContract(x)))
        { }
    }
    [Serializable]
    [DataContract(Name = "Note")]
    public class NoteContract
    {
        public NoteContract() { }
        public NoteContract(Note aNote)
        {
            CreationDate = aNote.CreationDate.ToString("G");
            NoteId = aNote.ID;
            NoteTypeId = aNote.NoteType.ID;
            NoteType = aNote.NoteType.Name;
            NoteFlavor = aNote.NoteType.Flavor.ToString();
            OtherStapleId = aNote.OtherStapleID;
            PageNumber = aNote.PageNumber;
            Height = aNote.Size.Height;
            Width = aNote.Size.Width;
            XPosition = aNote.Position.X;
            YPosition = aNote.Position.Y;
            Title = aNote.Title;
            Text = aNote.Text;
            UserId = aNote.CreatedBy.ID;
            UserName = aNote.CreatedBy.DisplayName;
            UserAccount = aNote.CreatedBy.Name;
        }
        [DataMember]
        public string CreationDate { get; set; }
        [DataMember]
        public long NoteId { get; set; }
        [DataMember]
        public long NoteTypeId { get; set; }
        [DataMember]
        public string NoteType { get; set; }
        [DataMember]
        public string NoteFlavor { get; set; }
        [DataMember]
        public long OtherStapleId { get; set; }
        [DataMember]
        public long PageNumber { get; set; }
        [DataMember]
        public long Width { get; set; }
        [DataMember]
        public long Height { get; set; }
        [DataMember]
        public long XPosition { get; set; }
        [DataMember]
        public long YPosition { get; set; }
        [DataMember]
        public string Title { get; set; }
        [DataMember]
        public string Text { get; set; }
        [DataMember]
        public long UserId { get; set; }
        [DataMember]
        public string UserAccount { get; set; }
        [DataMember]
        public string UserName { get; set; }
    }
    [Serializable]
    [CollectionDataContract(Name = "Notes", ItemName = "Note")]
    public class NoteListContract : List<NoteContract>
    {
        public NoteListContract() { }
        public NoteListContract(IEnumerable<Note> aNotes)
            : base(aNotes.Select(x => new NoteContract(x)))
        { }
    }
    [Serializable]
    [DataContract(Name = "Page")]
    public class PageContract
    {
        public PageContract() { }
        public PageContract(PageData aPage)
        {
            BinaryReader lReader = new BinaryReader(aPage.Stream);
            int lBufferLength = 1024;

            int lCount = 0;
            List<byte> lBuffer = new List<byte>();
            do
            {
                byte[] lBytesRead = lReader.ReadBytes(lBufferLength);
                lBuffer.AddRange(lBytesRead);
                lCount = lBytesRead.Length;
            } while (lCount == lBufferLength);

            byte[] ContentBytes = lBuffer.ToArray();


            //System.Drawing.Image image = null;
            //try
            //{

            //    using (MemoryStream ms = new MemoryStream(ContentBytes))
            //    {
            //        image = System.Drawing.Image.FromStream(ms);
            //        ImageCodecInfo codecInfo = ImageCodecInfo.GetImageEncoders().First(x => x.FormatID == image.RawFormat.Guid);
            //        Extension = codecInfo.FilenameExtension.Split(new char[] { ';', '.', '*' }, StringSplitOptions.RemoveEmptyEntries).FirstOrDefault();
            //        if (string.IsNullOrWhiteSpace(Extension))
            //            Extension = aPage.Extension;

            //    }
            //}
            //catch (ArgumentException e)
            //{
            //    //failed to convert to image, thats ok.
            //    Extension = aPage.Extension;
            //}

            //if (image != null)
            //{
            //    Bitmap lBitmap = new Bitmap(image);
            //    ImageCodecInfo codecInfo = ImageCodecInfo.GetImageEncoders().First(x => x.FormatID == image.RawFormat.Guid);                
            //    EncoderParameters encodingParameters = new EncoderParameters();
            //    //encodingParameters.Param[0] = new EncoderParameter(System.Drawing.Imaging.Encoder.Quality, 90L);
            //    //encodingParameters.Param[1] = new EncoderParameter(System.Drawing.Imaging.Encoder.Compression, (long)EncoderValue.CompressionLZW);

            //    using (MemoryStream ms = new MemoryStream())
            //    {
            //        lBitmap.Save(ms, codecInfo, encodingParameters);
            //        ContentBytes = ms.ToArray();
            //    }                    
            //}            
            Content = Convert.ToBase64String(ContentBytes);
            Extension = aPage.Extension;
        }

        //[IgnoreDataMember]
        //public byte [] ContentBytes { get; set; }
        [DataMember]
        public string Content { get; set; }
        [DataMember]
        public string Extension { get; set; }
    }
    [Serializable]
    [CollectionDataContract(Name = "Pages", ItemName = "Page")]
    public class PageListContract : List<PageContract>
    {
        public PageListContract() { }
        public PageListContract(IEnumerable<PageData> aPages)
            : base(aPages.Select(x => new PageContract(x)))
        { }
    }

    [Serializable]
    [DataContract(Name = "Document")]
    public class DocumentContract : OnBaseContract
    {
        public DocumentContract() { }
        public DocumentContract(Document aDocument, List<PageData> aPages)
        {
            Success = true;
            DocumentTypeGroupID = aDocument.DocumentType.DocumentTypeGroup.ID;
            DocumentTypeID = aDocument.DocumentType.ID;
            DocumentTypeName = aDocument.DocumentType.Name;
            DocumentName = aDocument.Name;
            DocumentId = aDocument.ID;
            DocumentDate = aDocument.DocumentDate.ToString("G");
            DateStored = aDocument.DateStored.ToString("G");
            CreatedUserName = aDocument.CreatedBy.DisplayName;
            CreatedUserAccount = aDocument.CreatedBy.Name;
            CreatedUserID = aDocument.CreatedBy.ID;
            Status = (int)aDocument.Status;
            Keywords = new KeywordListContract(aDocument.GetStandaloneKeywords());
            Notes = new NoteListContract(aDocument.Notes);
            Pages = new PageListContract(aPages);
        }
        [DataMember]
        public PageListContract Pages { get; set; }
        [DataMember]
        public KeywordListContract Keywords { get; set; }
        [DataMember]
        public NoteListContract Notes { get; set; }
        [DataMember]
        public long DocumentTypeID { get; set; }
        [DataMember]
        public string DocumentTypeName { get; set; }
        [DataMember]
        public long DocumentTypeGroupID { get; set; }
        [DataMember]
        public long DocumentId { get; set; }
        [DataMember]
        public string DocumentName { get; set; }
        [DataMember]
        public string DocumentDate { get; set; }
        [DataMember]
        public string DateStored { get; set; }
        [DataMember]
        public string CreatedUserName { get; set; }
        [DataMember]
        public string CreatedUserAccount { get; set; }
        [DataMember]
        public long CreatedUserID { get; set; }
        [DataMember]
        public int Status { get; set; }
    }

    [Serializable]
    [DataContract(Name = "KeywordType")]
    public class KeywordTypeContract
    {
        [DataMember]
        public long ID { get; set; }
        [DataMember]
        public string Name { get; set; }
        [DataMember]
        public int DataType { get; set; }
        [DataMember]
        public string DataTypeName { get; set; }

        public KeywordTypeContract() { }
        public KeywordTypeContract(KeywordType aType)
        {
            ID = aType.ID;
            Name = aType.Name;
            DataType = (int)aType.DataType;
            DataTypeName = Enum.GetName(typeof(KeywordDataType), aType.DataType);            
        }
    }
    [Serializable]
    [CollectionDataContract(Name = "KeywordTypes", ItemName="Keyword")]
    public class KeywordTypeListContract : List<KeywordTypeContract>
    {
        public KeywordTypeListContract() { }
        public KeywordTypeListContract(IEnumerable<KeywordType> aKeywordTypes)
            : base(aKeywordTypes.Select(x => new KeywordTypeContract(x)))
        { }
    }
    [Serializable]
    [DataContract(Name = "DocumentType")]
    public class DocumentTypeContract : OnBaseContract
    {
        public DocumentTypeContract() { Success = false; }

        [DataMember]
        public long DocumentTypeGroupID { get;  set; }
        [DataMember]
        public long DocumentTypeID { get;  set; }

        [DataMember]
        public string DocumentTypeName { get;  set; }

        [DataMember]
        public KeywordTypeListContract AssociatedKeywordTypes { get;  set; }

        [DataMember]
        public KeywordTypeListContract RequiredKeywordTypes { get;  set; }


        public DocumentTypeContract(string aErrorMessage, string aErrorDetails)
        {
            Success = false;
            ErrorMessage = aErrorMessage;
            ErrorDetails = aErrorDetails;
        }
        public DocumentTypeContract(DocumentType aDocumentType, IEnumerable<KeywordType> aAssociatedKeywordTypes, IEnumerable<KeywordType> aRequiredKeywordTypes)            
        {
            Success = true;
            DocumentTypeGroupID = aDocumentType.DocumentTypeGroup.ID;
            DocumentTypeID = aDocumentType.ID;
            DocumentTypeName = aDocumentType.Name;
            AssociatedKeywordTypes = new KeywordTypeListContract(aAssociatedKeywordTypes);
            RequiredKeywordTypes = new KeywordTypeListContract(aAssociatedKeywordTypes);
        }
    }



    public static class ServiceWrapper
    {
        private static string ConnectionString {get{return System.Configuration.ConfigurationManager.ConnectionStrings["OnBase"].ConnectionString;}}
        private static string Url { get { return System.Configuration.ConfigurationManager.AppSettings["OnBaseUrl"]; } }
        private static string DataSource { get { return System.Configuration.ConfigurationManager.AppSettings["OnBaseDataSource"]; } }

        public static DocumentContract GetDocument(HttpRequest request)
        {
            try
            {
                string username = request.Form.Get("Username");
                string password = request.Form.Get("Password");
                string actualUser = request.Form.Get("ActualUser");
                string actualUserIp = request.Form.Get("ActualUserIP");
                string docIdString = request.Form.Get("DocID");
                string alternateFormat = request.Form.Get("AlternateFormat");
                string applicationIdString = request.Form.Get("applicationid");
                string appSpecificId = request.Form.Get("appspecificid");

                int? applicationId = null;
                if (!string.IsNullOrWhiteSpace(applicationIdString))
                {
                    try
                    {
                        applicationId = int.Parse(applicationIdString);
                    }
                    catch
                    {         
                        throw new ArgumentException("ApplicationId must be a number", "ApplicationId");                        
                    }
                }
                long DocID;
                if (!long.TryParse(docIdString, out DocID))
                {
                    throw new ArgumentException("DocID must be a number", "DocID");
                }
                return GetDocument(username, password, request.UserHostAddress, actualUser, actualUserIp, 
                    DocID, alternateFormat, applicationId, appSpecificId);
            }
            catch (Exception ex)
            {
                return new DocumentContract() { Success = false, ErrorMessage = ex.Message, ErrorDetails = ex.ToString() };
            }
        }
        public static DocumentContract GetDocument(string username, string password, string appIp, string actualUser, string actualIP,
            long DocID, string alternateFormat, int? applicationId, string appSpecificId)
        {
            try
            {
                OnBaseHelper lHelper = GetHelper(username, password);

                Document lDocument = lHelper.GetDocument(DocID, DocumentRetrievalOptions.LoadKeywords | DocumentRetrievalOptions.LoadNotes | DocumentRetrievalOptions.LoadRevisionsAndVersions);

                //lDocument.DefaultFileType.DisplayType;
                lDocument.LatestRevision.Renditions.Count();
                List<PageData> lPages = lHelper.GetAllPages(alternateFormat, lDocument.DefaultRenditionOfLatestRevision);
                LogUsage(username, actualUser, appIp, actualIP, lDocument.ID, "GetDoc", null, DateTime.Now, applicationId, appSpecificId);
                return new DocumentContract(lDocument, lPages);
            }
            catch (UnityAPIException unityEx)
            {
                return new DocumentContract() { Success = false, ErrorMessage = unityEx.ServerMessage, ErrorDetails = unityEx.ToString() };
            }
            catch (Exception ex)
            {
                return new DocumentContract() { Success = false, ErrorMessage = ex.Message, ErrorDetails = ex.ToString() };
            }            
        }

        public static AddDocumentContract AddDocument(HttpRequest request)
        {
            try
            {
                string username = request.Form.Get("Username");
                string password = request.Form.Get("Password");
                string actualUser = request.Form.Get("ActualUser");
                string actualUserIp = request.Form.Get("ActualUserIP");
                string docTypeString = request.Form.Get("DocType");
                string keys = request.Form.Get("Keys");
                string keyValues = request.Form.Get("KeyValues");
                string applicationIdString = request.Form.Get("applicationid");
                string appSpecificId = request.Form.Get("appspecificid");

                int? applicationId = null;
                if (!string.IsNullOrWhiteSpace(applicationIdString))
                {
                    try
                    {
                        applicationId = int.Parse(applicationIdString);
                    }
                    catch
                    {
                        throw new ArgumentException("ApplicationId must be a number", "ApplicationId");
                    }
                }
                if (string.IsNullOrWhiteSpace(docTypeString))
                {
                    throw new ArgumentNullException("DocType", "DocType is a required parameter");
                }
                long DocTypeID;
                if (!long.TryParse(docTypeString, out DocTypeID))
                {
                    throw new ArgumentException("DocType must be a number", "DocType");
                }

                List<byte[]> lPageContent = new List<byte[]>();
                List<string> lPageExtensions = new List<string>();
                for (int i = 0; i < request.Files.Count; ++i)
                {
                    HttpPostedFile lFile = request.Files[i];
                    lPageExtensions.Add(Path.GetExtension(lFile.FileName).TrimStart('.'));

                    if (lFile.InputStream.Length > 5000000)
                        throw new Exception("Page content may not exceed 5MB");
                    BinaryReader lreader = new BinaryReader(lFile.InputStream);
                    
                    int lLength = Convert.ToInt32(lFile.InputStream.Length);
                    byte [] lContent = lreader.ReadBytes(lLength);                    

                    //lFile.InputStream.Write(lContent, 0, lContent.Length);
                    lPageContent.Add(lContent);
                    lFile.InputStream.Close();
                }
                IEnumerable<string> lKeywordTypesNames = (string.IsNullOrWhiteSpace(keys)) ? null : keys.Split(',');
                IEnumerable<string> lKeywordValues = (string.IsNullOrWhiteSpace(keyValues)) ? null : keyValues.Split(',');
                return AddDocument(username, password, request.UserHostAddress, actualUser, actualUserIp, 
                   DocTypeID, lPageContent, lPageExtensions, lKeywordTypesNames, lKeywordValues, 
                   applicationId, appSpecificId);                 
            }
            catch (Exception ex)
            {
                return new AddDocumentContract() { Success = false, ErrorMessage = ex.Message, ErrorDetails = ex.ToString() };
            }
        }

        public static AddDocumentContract AddDocument(string username, string password, string appIp, string actualUser, string actualIP,
            long DocType, IEnumerable<Stream> FileContent, IEnumerable<string> FileExtensions, IEnumerable<string> Keys, IEnumerable<string> KeyValues, int? applicationId, string appSpecificId)
        {
            List<byte[]> lPageContent = new List<byte []>();
            foreach (Stream lStream in FileContent)
            {
                if(lStream.Length > Int32.MaxValue)
                    throw new Exception("Page content may not exceed 2Gb");

                byte [] lContent = new byte[(int)Math.Max(Int32.MaxValue, lStream.Length)];

                lStream.Write(lContent, 0, lContent.Length);
                lPageContent.Add(lContent);                
            }
            return AddDocument(username, password, appIp, actualUser, actualIP,
                DocType, lPageContent, FileExtensions, Keys, KeyValues, applicationId, appSpecificId);
        }


        public static AddDocumentContract AddDocument(string username, string password, string appIp, string actualUser, string actualIP, 
            long DocType, IEnumerable<byte[]> FileContent, IEnumerable<string> FileExtensions, IEnumerable<string> Keys, IEnumerable<string> KeyValues, int ? applicationId, string appSpecificId)
        {
            if (string.IsNullOrWhiteSpace(username))
            {
                throw new ArgumentNullException("Username", "Username is a required parameter");
            }
            if (string.IsNullOrWhiteSpace(password))
            {
                throw new ArgumentNullException("Password", "Password is a required parameter");
            }
            if (string.IsNullOrWhiteSpace(actualUser))
            {
                throw new ArgumentNullException("ActualUser", "ActualUser is a required parameter");
            }
            if (string.IsNullOrWhiteSpace(actualIP))
            {
                throw new ArgumentNullException("ActualUserIP", "ActualUserIP is a required parameter");
            }


            List<PageData> lPageData = new List<PageData>();
            try
            {
                OnBaseHelper lHelper = GetHelper(username, password);

                DocumentType docType = lHelper.GetDocumentType(DocType);
                bool AreAllImages = true;
               
                List<FileType> lFileTypes = new List<FileType>();
                foreach (var inputPage in FileContent.Zip(FileExtensions, (x, y) => new { data = x, extension = y }))
                {
                    MemoryStream stream = new MemoryStream();
                    
                    stream.Write(inputPage.data, 0, inputPage.data.Length);
                    stream.Flush();
                    stream.Position = 0;

                    FileType lFileType = lHelper.GetFileTypeByExtension(inputPage.extension);
                    if (lFileType == null)
                        lFileType = docType.DefaultFileType;
                    lPageData.Add(lHelper.CreatePageData(stream, inputPage.extension));
                    lFileTypes.Add(lFileType);
                    AreAllImages &= lFileType.DisplayType == DisplayType.Image;
                    
                }
                if (lPageData.Count > 1 && !AreAllImages)
                    throw new NotSupportedException("Only image documents may have multiple files");
                
                IEnumerable<Keyword> lKeywords = lHelper.CreateKeywordCollection(Keys, KeyValues, docType);                                    

                Document lDocument = lHelper.StoreNewDocument(docType, lPageData, lKeywords, null, lFileTypes.FirstOrDefault(), true, true);
                LogUsage(username, actualUser, appIp, actualIP, lDocument.ID, "AddDoc", null, lDocument.DateStored, applicationId, appSpecificId);
                return new AddDocumentContract(lDocument);
            }
            catch (UnityAPIException unityEx)
            {
                return new AddDocumentContract() { Success = false, ErrorMessage = unityEx.ServerMessage, ErrorDetails = unityEx.ToString() };
            }
            catch (Exception ex)
            {
                return new AddDocumentContract() { Success = false, ErrorMessage = ex.Message, ErrorDetails = ex.ToString() };
            }
            finally
            {
                if (lPageData != null)
                {
                    foreach (PageData lPage in lPageData)
                    {
                        if (lPage != null)
                            lPage.Dispose();
                    }
                }                
            }
        }

        public static OnBaseContract UpdateDocumentTypeKeys(HttpRequest request)
        {
            try
            {
                string username = request.Form.Get("Username");
                string password = request.Form.Get("Password");
                string actualUser = request.Form.Get("ActualUser");
                string actualUserIp = request.Form.Get("ActualUserIP");
                string DocIdString = request.Form.Get("DocID");
                string docTypeString = request.Form.Get("DocType");
                string keys = request.Form.Get("Keys");
                string keyValues = request.Form.Get("KeyValues");
                string applicationIdString = request.Form.Get("applicationid");
                string appSpecificId = request.Form.Get("appspecificid");
                int? applicationId = null;
                if (!string.IsNullOrWhiteSpace(applicationIdString))
                {
                    try
                    {
                        applicationId = int.Parse(applicationIdString);
                    }
                    catch
                    {
                        throw new ArgumentException("ApplicationId must be a number", "ApplicationId");
                    }
                }
                long? DocType = null;
                if (!string.IsNullOrWhiteSpace(docTypeString))
                {
                    try
                    {
                        DocType = long.Parse(docTypeString);
                    }
                    catch
                    {
                        throw new ArgumentException("DocType must be a number", "DocType");
                    }
                }
                if (string.IsNullOrWhiteSpace(DocIdString))
                {
                    throw new ArgumentNullException("DocID", "DocID is a required parameter");
                }
                long DocID;
                if (!long.TryParse(DocIdString, out DocID))
                {
                    throw new ArgumentException("DocID must be a number", "DocID");
                }
                IEnumerable<string> lKeywordTypesNames = (string.IsNullOrWhiteSpace(keys)) ? null : keys.Split(',');
                IEnumerable<string> lKeywordValues = (string.IsNullOrWhiteSpace(keyValues)) ? null : keyValues.Split(',');
                return UpdateDocumentTypeKeys(username, password, request.UserHostAddress, actualUser, actualUserIp,
                    DocID, DocType, lKeywordTypesNames, lKeywordValues, applicationId, appSpecificId);
            }
            catch (Exception ex)
            {
                return new OnBaseContract(){ Success = false, ErrorMessage = ex.Message, ErrorDetails = ex.ToString() };
            }
        }

        public static OnBaseContract UpdateDocumentTypeKeys(string username, string password, string appIp, string actualUser, string actualIP,
            long DocID, long ? DocType, IEnumerable<string> Keys, IEnumerable<string> KeyValues, int ? applicationId, string appSpecificId)
        {
            try{
                OnBaseHelper lHelper = GetHelper(username, password);

                Document lDocument = lHelper.GetDocument(DocID, DocumentRetrievalOptions.LoadKeywords);
                DocumentType docType = (DocType.HasValue) ? lHelper.GetDocumentType(DocType.Value) : lDocument.DocumentType;
                IEnumerable<Keyword> lKeywords = lHelper.CreateKeywordCollection(Keys, KeyValues, docType);

                lDocument = lHelper.ChangeDocumentType(lDocument, docType, lKeywords, true, true);
                
                LogUsage(username, actualUser, appIp, actualIP, lDocument.ID, "ReIndexDoc", null, lDocument.DateStored, applicationId, appSpecificId);
                return new OnBaseContract(){ Success = true };
            }
            catch (UnityAPIException unityEx)
            {
                return new OnBaseContract(){ Success = false, ErrorMessage = unityEx.ServerMessage, ErrorDetails = unityEx.ToString() };
            }
            catch (Exception ex)
            {
                return new OnBaseContract(){ Success = false, ErrorMessage = ex.Message, ErrorDetails = ex.ToString() };
            }            

        }
        public static AddNoteContract AddNote(HttpRequest request)
        {
            try
            {

                string username = request.Form.Get("Username");
                string password = request.Form.Get("Password");
                string actualUser = request.Form.Get("ActualUser");
                string actualUserIp = request.Form.Get("ActualUserIP");
                string DocIdString = request.Form.Get("DocID");
                string NoteTypeIdString = request.Form.Get("NoteTypeID");
                string Text = request.Form.Get("Text");
                string applicationIdString = request.Form.Get("applicationid");
                string appSpecificId = request.Form.Get("appspecificid");

                int? applicationId = null;
                if (!string.IsNullOrWhiteSpace(applicationIdString))
                {
                    try
                    {
                        applicationId = int.Parse(applicationIdString);
                    }
                    catch
                    {
                        throw new ArgumentException("ApplicationId must be a number", "ApplicationId");
                    }
                }

                if (string.IsNullOrWhiteSpace(DocIdString))
                {
                    throw new ArgumentNullException("DocID", "DocID is a required parameter");
                }
                if (string.IsNullOrWhiteSpace(NoteTypeIdString))
                {
                    throw new ArgumentNullException("NoteTypeID", "NoteTypeID is a required parameter");
                }
                long NoteTypeID;
                if (!long.TryParse(NoteTypeIdString, out NoteTypeID))
                {
                    throw new ArgumentException("NoteTypeID must be a number", "NoteTypeID");
                }
                long DocID;
                if (!long.TryParse(DocIdString, out DocID))
                {
                    throw new ArgumentException("DocID must be a number", "DocID");
                }

                long? Height = null;
                string HeightString = request.Form.Get("Height");
                if (!string.IsNullOrWhiteSpace(HeightString))
                    Height = long.Parse(HeightString);
                long? Width = null;
                string WidthString = request.Form.Get("Width");
                if (!string.IsNullOrWhiteSpace(WidthString))
                    Width = long.Parse(WidthString);

                long? XPosition = null;
                string xPosString = request.Form.Get("XPosition");
                if (!string.IsNullOrWhiteSpace(xPosString))
                    XPosition = long.Parse(xPosString);
                long? YPosition = null;
                string yPosString = request.Form.Get("YPosition");
                if (!string.IsNullOrWhiteSpace(yPosString))
                    YPosition = long.Parse(yPosString);


                long? PageNumber = null;
                string pageString = request.Form.Get("PageNumber");
                if (!string.IsNullOrWhiteSpace(pageString))
                {
                    PageNumber = long.Parse(pageString);
                }
                return AddNote(username, password, request.UserHostAddress, actualUser, actualUserIp, 
                    DocID, NoteTypeID, Text, Width, Height, XPosition, YPosition, PageNumber, applicationId, appSpecificId);
            }
            catch (Exception ex)
            {
                return new AddNoteContract() { Success = false, ErrorMessage = ex.Message, ErrorDetails = ex.ToString() };
            }
        }
        public static AddNoteContract AddNote(string username, string password, string appIp, string actualUser, string actualIP, long DocID, long NoteType, string text, long? width, long? height, long? xPos, long? yPos, long? pageNumber, int? applicationId, string appSpecificId)
        {
            try
            {
                if (string.IsNullOrWhiteSpace(username))
                {
                    throw new ArgumentNullException("Username", "Username is a required parameter");
                }
                if (string.IsNullOrWhiteSpace(password))
                {
                    throw new ArgumentNullException("Password", "Password is a required parameter");
                }
                if (string.IsNullOrWhiteSpace(actualUser))
                {
                    throw new ArgumentNullException("ActualUser", "ActualUser is a required parameter");
                }
                if (string.IsNullOrWhiteSpace(actualIP))
                {
                    throw new ArgumentNullException("ActualUserIP", "ActualUserIP is a required parameter");
                }


                OnBaseHelper lHelper = GetHelper(username, password);

                Note lNote = lHelper.AddNote(DocID, NoteType, text, width, height, xPos, yPos, pageNumber);

                LogUsage(username, actualUser, appIp, actualIP, DocID, "AddNote", null, DateTime.Now, applicationId, appSpecificId);
                return new AddNoteContract(lNote);
            }
            catch (UnityAPIException unityEx)
            {
                return new AddNoteContract() { Success = false, ErrorMessage = unityEx.ServerMessage, ErrorDetails = unityEx.ToString() };
            }
            catch (Exception ex)
            {
                return new AddNoteContract() { Success = false, ErrorMessage = ex.Message, ErrorDetails = ex.ToString() };
            }
        }

        public static UserDetailsContract GetUserDetails(HttpRequest request)
        {
            try
            {
                string username = request.Form.Get("Username");
                string password = request.Form.Get("Password");
                string actualUser = request.Form.Get("ActualUser");
                string actualUserIp = request.Form.Get("ActualUserIP");
                string userAcountName = request.Form.Get("UserAccountName");
                string applicationIdString = request.Form.Get("applicationid");
                string appSpecificId = request.Form.Get("appspecificid");

                int? applicationId = null;
                if (!string.IsNullOrWhiteSpace(applicationIdString))
                {
                    try
                    {
                        applicationId = int.Parse(applicationIdString);
                    }
                    catch
                    {
                        throw new ArgumentException("ApplicationId must be a number", "ApplicationId");
                    }
                }
                return GetUserDetails(username, password, request.UserHostAddress, actualUser, actualUserIp, userAcountName, applicationId, appSpecificId);
            }
            catch (Exception ex)
            {
                return new UserDetailsContract() { Success = false, ErrorMessage = ex.Message, ErrorDetails = ex.ToString() };
            }
        }
        public static UserDetailsContract GetUserDetails(string username, string password, string appIp, string actualUser, string actualIP, string lookupAccountName, int? applicationId, string appSpecificId)
        {

            try
            {
                if (string.IsNullOrWhiteSpace(username))
                {
                    throw new ArgumentNullException("Username", "Username is a required parameter");
                }
                if (string.IsNullOrWhiteSpace(password))
                {
                    throw new ArgumentNullException("Password", "Password is a required parameter");
                }
                if (string.IsNullOrWhiteSpace(actualUser))
                {
                    throw new ArgumentNullException("ActualUser", "ActualUser is a required parameter");
                }
                if (string.IsNullOrWhiteSpace(actualIP))
                {
                    throw new ArgumentNullException("ActualUserIP", "ActualUserIP is a required parameter");
                }
                if (string.IsNullOrWhiteSpace(lookupAccountName))
                {
                    throw new ArgumentNullException("UserAccountName", "UserAccountName is a required parameter");
                }


                OnBaseHelper lHelper = GetHelper(username, password);

                Hyland.Unity.User lUser = lHelper.GetUserData(lookupAccountName);
                LogUsage(username, actualUser, appIp, actualIP, null,
                    "UserDetails", "Lookup details for " + lookupAccountName, DateTime.Now, applicationId, appSpecificId);                
                return new UserDetailsContract(lUser);
            }
            catch (UnityAPIException unityEx)
            {
                return new UserDetailsContract() { Success = false, ErrorMessage = unityEx.ServerMessage, ErrorDetails = unityEx.ToString() };
            }
            catch (Exception ex)
            {
                return new UserDetailsContract() { Success = false, ErrorMessage = ex.Message, ErrorDetails = ex.ToString() };
            }


        }

        public static OnBaseContract DeleteDocument(HttpRequest request)
        {
            try
            {
                string username = request.Form.Get("Username");
                string password = request.Form.Get("Password");
                string actualUser = request.Form.Get("ActualUser");
                string actualUserIp = request.Form.Get("ActualUserIP");
                string DocIdString = request.Form.Get("DocID");
                string applicationIdString = request.Form.Get("applicationid");
                string appSpecificId = request.Form.Get("appspecificid");
                int? applicationId = null;
                if (!string.IsNullOrWhiteSpace(applicationIdString))
                {
                    try
                    {
                        applicationId = int.Parse(applicationIdString);
                    }
                    catch
                    {
                        throw new ArgumentException("ApplicationId must be a number", "ApplicationId");
                    }
                }
                if (string.IsNullOrWhiteSpace(DocIdString))
                {
                    throw new ArgumentNullException("DocID", "DocID is a required parameter");
                }
                long DocID;
                if (!long.TryParse(DocIdString, out DocID))
                {
                    throw new ArgumentException("DocID must be a number", "DocID");
                }
                return DeleteDocument(username, password, request.UserHostAddress, actualUser, actualUserIp, DocID, applicationId, appSpecificId);
            }
            catch (Exception ex)
            {
                return new OnBaseContract(){ Success = false, ErrorMessage = ex.Message, ErrorDetails = ex.ToString() };
            }
           
        }
        public static OnBaseContract DeleteDocument(string username, string password, string appIp, string actualUser, string actualIP, long DocID, int? applicationId, string appSpecificId)
        {
            try
            {
                if (string.IsNullOrWhiteSpace(username))
                {
                    throw new ArgumentNullException("Username", "Username is a required parameter");
                }
                if (string.IsNullOrWhiteSpace(password))
                {
                    throw new ArgumentNullException("Password", "Password is a required parameter");
                }
                if (string.IsNullOrWhiteSpace(actualUser))
                {
                    throw new ArgumentNullException("ActualUser", "ActualUser is a required parameter");
                }
                if (string.IsNullOrWhiteSpace(actualIP))
                {
                    throw new ArgumentNullException("ActualUserIP", "ActualUserIP is a required parameter");
                }


                OnBaseHelper lHelper = GetHelper(username, password);

                lHelper.DeleteDocument(DocID);
                LogUsage(username, actualUser, appIp, actualIP, DocID, "DeleteDoc", null, DateTime.Now, applicationId, appSpecificId);
                return new OnBaseContract(){ Success = true };
            }
            catch (UnityAPIException unityEx)
            {
                return new OnBaseContract(){ Success = false, ErrorMessage = unityEx.ServerMessage, ErrorDetails = unityEx.ToString() };
            }
            catch (Exception ex)
            {
                return new OnBaseContract(){ Success = false, ErrorMessage = ex.Message, ErrorDetails = ex.ToString() };
            }

        }

        public static OnBaseHelper GetHelper(string aUsername, string aPassword)
        {
            OnBaseHelper lHelper = null;
            string lSessionId = null;
            if (GetSession(aUsername, out lSessionId))
            {
                try
                {
                    lHelper = new OnBaseHelper(Url, lSessionId);
                }
                catch
                {
                    DeleteSession(aUsername);
                }
            }
            if (lHelper == null)
            {
                lHelper = new OnBaseHelper(Url, aUsername, aPassword, DataSource);
                SaveSession(aUsername, lHelper.SessionId);
            }
            return lHelper;
        }


        public static DocumentTypeContract GetDocumentTypeDetails(HttpRequest request)
        {
            try
            {
                string username = request.Form.Get("Username");
                string password = request.Form.Get("Password");
                string actualUser = request.Form.Get("ActualUser");
                string actualUserIp = request.Form.Get("ActualUserIP");
                string docTypeIdString= request.Form.Get("DocTypeId");
                string applicationIdString = request.Form.Get("applicationid");
                string appSpecificId = request.Form.Get("appspecificid");

                int? applicationId = null;
                if (!string.IsNullOrWhiteSpace(applicationIdString))
                {
                    try
                    {
                        applicationId = int.Parse(applicationIdString);
                    }
                    catch
                    {
                        throw new ArgumentException("ApplicationId must be a number", "ApplicationId");
                    }
                }
                long docTypeId;
                if (!long.TryParse(docTypeIdString, out docTypeId))
                {
                    throw new ArgumentException("doctypeId must be a number", "DocTypeId");
                }
                return GetDocumentTypeDetails(username, password, request.UserHostAddress, actualUser, actualUserIp, docTypeId, applicationId, appSpecificId);
            }
            catch (Exception ex)
            {
                return new DocumentTypeContract(ex.Message, ex.ToString() );
            }
        }
        public static DocumentTypeContract GetDocumentTypeDetails(string username, string password, string appIp, string actualUser, string actualIP, long DocTypeId, int? applicationId, string appSpecificId)
        {
            try
            {
                if (string.IsNullOrWhiteSpace(username))
                {
                    throw new ArgumentNullException("Username", "Username is a required parameter");
                }
                if (string.IsNullOrWhiteSpace(password))
                {
                    throw new ArgumentNullException("Password", "Password is a required parameter");
                }
                if (string.IsNullOrWhiteSpace(actualUser))
                {
                    throw new ArgumentNullException("ActualUser", "ActualUser is a required parameter");
                }
                if (string.IsNullOrWhiteSpace(actualIP))
                {
                    throw new ArgumentNullException("ActualUserIP", "ActualUserIP is a required parameter");
                }


                OnBaseHelper lHelper = GetHelper(username, password);
                DocumentType aType = lHelper.GetDocumentType(DocTypeId);
                                
                LogUsage(username, actualUser, appIp, actualIP, null, 
                    "GetDocTypeInfo", string.Format("requesting information about DocumentType={0}", DocTypeId), 
                    DateTime.Now, applicationId, appSpecificId);
                return new DocumentTypeContract(aType, aType.GetDefinedKeywordTypes(), aType.KeywordTypesRequiredForArchival);
            }
            catch (UnityAPIException unityEx)
            {
                return new DocumentTypeContract(unityEx.ServerMessage, unityEx.ToString());
            }
            catch (Exception ex)
            {
                return new DocumentTypeContract(ex.Message, ex.ToString());
            }
        }

        public static bool GetSession(string aUsername, out string aSessionId)
        {
            aSessionId = null;
            using (SqlConnection lConnection = new SqlConnection(ConnectionString))
            {
                lConnection.Open();
                const string query = "select SessionID from ActiveSessions where UserName=:aUserName and Url=:aUrl";
                using (SqlCommand lCommand = new SqlCommand(query, lConnection))
                {
                    lCommand.Parameters.AddWithValue("aUserName", aUsername);
                    lCommand.Parameters.AddWithValue("aUrl", Url);

                    using(SqlDataReader lReader = lCommand.ExecuteReader())
                    {
                        if (lReader.Read())
                        {
                            aSessionId = lReader[0].ToString();
                            return true;
                        }
                    }
                }
            }
            return false;
        }
        public static void DeleteSession(string aUsername)
        {            
            using (SqlConnection lConnection = new SqlConnection(ConnectionString))
            {
                lConnection.Open();
                const string query = "delete from ActiveSessions where UserName=:aUserName and Url=:aUrl";
                using (SqlCommand lCommand = new SqlCommand(query, lConnection))
                {
                    lCommand.Parameters.AddWithValue("aUserName", aUsername);
                    lCommand.Parameters.AddWithValue("aUrl", Url);
                    lCommand.ExecuteNonQuery();
                }
            }
        }
        public static void SaveSession(string aUsername, string aSessionId)
        {            
            using (SqlConnection lConnection = new SqlConnection(ConnectionString))
            {
                lConnection.Open();
                const string query = "insert into ActiveSessions(Username, Url, SessionId) Values(:aUserName, :aUrl, :sessionid)";
                using (SqlCommand lCommand = new SqlCommand(query, lConnection))
                {
                    lCommand.Parameters.AddWithValue("aUserName", aUsername);
                    lCommand.Parameters.AddWithValue("aUrl", Url);
                    lCommand.Parameters.AddWithValue("sessionid", aSessionId);
                    lCommand.ExecuteNonQuery();
                }
            }            
        }
        public static void LogUsage(string aApplicationLoginId, string aLoginId, string aApplicationIp, string aActualIp, long ? docid, string action, string description, DateTime aTimeStamp, int ? applicationid, string appspecificid)
        {            
            using (SqlConnection lConnection = new SqlConnection(ConnectionString))
            {
                lConnection.Open();
                const string query = "insert into SIS_EDMS_AUDIT(dept_user,actual_user, timestamp, dept_request_ip, user_request_ip, docid, action,description,applicationid,appspecificid) " +
                    "values(:dept_user, :actual_user, :aTimeStamp, :dept_request_ip, :user_request_ip, :docid, :action, :description, :applicationid, :appspecificid)";

                using (SqlCommand lCommand = new SqlCommand(query, lConnection))
                {
                    lCommand.Parameters.AddWithValue("dept_user", aApplicationLoginId);
                    lCommand.Parameters.AddWithValue("actual_user", aLoginId);
                    lCommand.Parameters.AddWithValue("aTimeStamp", aTimeStamp);
                    lCommand.Parameters.AddWithValue("dept_request_ip", aApplicationIp);
                    lCommand.Parameters.AddWithValue("user_request_ip", aActualIp);
                    if (docid.HasValue)
                        lCommand.Parameters.AddWithValue("docid", docid.Value);
                    else lCommand.Parameters.AddWithValue("docid", DBNull.Value);
                    lCommand.Parameters.AddWithValue("action", action);
                    if (string.IsNullOrWhiteSpace(description))
                        lCommand.Parameters.AddWithValue("description", DBNull.Value);
                    else lCommand.Parameters.AddWithValue("description", description);
                    if(applicationid.HasValue)
                        lCommand.Parameters.AddWithValue("applicationid", applicationid.Value);
                    else lCommand.Parameters.AddWithValue("applicationid", DBNull.Value);
                    if (string.IsNullOrWhiteSpace(appspecificid))
                        lCommand.Parameters.AddWithValue("appspecificid", DBNull.Value);
                    else lCommand.Parameters.AddWithValue("appspecificid", appspecificid);

                    lCommand.ExecuteNonQuery();
                }
            }        
        }       
    }
}
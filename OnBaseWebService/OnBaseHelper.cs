using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using Hyland.Unity;
using System.IO;
using System.Data.Common;
using Hyland;
using System.Reflection;

namespace UCDavis.StudentAffairs.OnBaseWebService
{
    public class DocumentLockedException : Exception
    {
        public Document Document { get; private set; }
        public User UserHoldingLock { get; private set; }
        public DocumentLockStatus LockStatus { get; private set; }

        const string ErrorFormat = "Unable to acquire lock for document {0} ({1}). Document locked by {2}";

        public DocumentLockedException(Document document, DocumentLock documentLock)
            : this(string.Format(ErrorFormat, document.Name, document.ID, documentLock.UserHoldingLock.Name), document, documentLock, null)
        { }

        public DocumentLockedException(string message, Document document, DocumentLock documentLock)
            : this(message, document, documentLock, null)
        { }

        public DocumentLockedException(string message, Document document, DocumentLock documentLock, Exception innerException)
            : base(message, innerException)
        {
            this.LockStatus = documentLock.Status;
            this.UserHoldingLock = documentLock.UserHoldingLock;
            this.Document = document;
        }
    }

    public class KeywordUpdate
    {
        public Keyword OldValue { get; private set; }
        public Keyword NewValue { get; private set; }

        private KeywordUpdate() { }
        public KeywordUpdate(OnBaseHelper aHelper, long aKeywordType, string aOldValue, string aNewValue, DocumentType aDocumentType = null)
        {
            KeywordType lType = aHelper.GetKeywordType(aKeywordType, aDocumentType);
            OldValue = OnBaseHelper.CreateKeyword(lType, aOldValue);
            NewValue = OnBaseHelper.CreateKeyword(lType, aNewValue);
        }
        public KeywordUpdate(OnBaseHelper aHelper, string aKeywordType, string aOldValue, string aNewValue, DocumentType aDocumentType = null)
        {
            KeywordType lType = aHelper.GetKeywordType(aKeywordType, aDocumentType);
            OldValue = OnBaseHelper.CreateKeyword(lType, aOldValue);
            NewValue = OnBaseHelper.CreateKeyword(lType, aNewValue);
        }
        public KeywordUpdate(Keyword aOldValue, Keyword aNewValue)
        {
            OldValue = aOldValue;
            NewValue = aNewValue;
        }
    }

    public static class DocumentExtensions
    {
        public static IEnumerable<Keyword> GetStandaloneKeywords(this Document aDocument)
        {
            KeywordRecord aRecord = aDocument.KeywordRecords.Find(x => x.KeywordRecordType.RecordType == RecordType.StandAlone);
            if (aRecord != null)
                return aRecord.Keywords;
            else return new List<Keyword>();
        }

        public static bool IsKeywordDefined(this Document aDocument, Keyword aKeyword)
        {
            if (aKeyword == null)
                return false;
            if (aDocument == null)
                throw new ArgumentNullException("Cannot validate keyword definition against a null document");
            return aDocument.DocumentType.IsDefined(aKeyword.KeywordType);
        }
    }

    public static class DocumentTypeExtensions
    {
        public static bool IsRequired(this DocumentType aDocumentType, Keyword aKeyword)
        {
            if (aKeyword == null)
                return false;
            return IsRequired(aDocumentType, aKeyword.KeywordType);
        }
        public static bool IsRequired(this DocumentType aDocumentType, KeywordType aKeywordType)
        {
            if (aKeywordType == null)
                return false;
            if (aDocumentType == null)
                throw new ArgumentNullException("Cannot validate keyword requirements against a null document type");
            return aDocumentType.KeywordTypesRequiredForArchival.Contains(aKeywordType);
        }
        public static bool IsDefined(this DocumentType aDocumentType, Keyword aKeyword)
        {
            if (aKeyword == null)
                return false;
            return IsDefined(aDocumentType, aKeyword.KeywordType);
        }
        public static bool IsDefined(this DocumentType aDocumentType, KeywordType aKeywordType)
        {
            if (aKeywordType == null)
                return false;
            if (aDocumentType == null)
                throw new ArgumentNullException("Cannot validate keyword definition against a null document type");
            return KeywordType.Equals(aKeywordType, aDocumentType.KeywordRecordTypes.FindKeywordType(aKeywordType.ID));
        }

        public static List<KeywordType> GetDefinedKeywordTypes(this DocumentType aDocumentType)
        {
            List<KeywordType> result = new List<KeywordType>();
            foreach (KeywordRecordType recordType in aDocumentType.KeywordRecordTypes)
            {
                result.AddRange(recordType.KeywordTypes);
            }
            result.AddRange(aDocumentType.KeywordRecordTypes.Find(OnBaseHelper.STANDALONE_KEYWORD_RECORD_TYPE).KeywordTypes);
            return result;
        }
    }
    public class OnBaseHelper
    {
        internal const long STANDALONE_KEYWORD_RECORD_TYPE = 0;

        public Application Application { get; private set; }
        //public string UserName { get; private set; }
        //public string Url { get; private set; }
        //public string DataSource { get; private set; }
        public string SessionId { get { return Application.SessionID; } }
        public bool IsConnected { get { return Application != null && Application.IsConnected; } }
        public OnBaseHelper(string aUrl, string aSessionID)
        {
            SessionIDAuthenticationProperties lauthentication = Application.CreateSessionIDAuthenticationProperties(aUrl, aSessionID, false);

            Application = Application.Connect(lauthentication);
        }

        public OnBaseHelper(string aUrl, string aUserName, string aPassword, string aDataSource)
        {
            AuthenticationProperties lAuthentication = Application.CreateOnBaseAuthenticationProperties(aUrl, aUserName, aPassword, aDataSource);           
            Application = Application.Connect(lAuthentication);
        }

        public void Disconnect()
        {
            Application.Disconnect();
        }            

        public User GetUserData(string aUser)
        {
            if (string.IsNullOrWhiteSpace(aUser))
                return Application.CurrentUser;
            else return Application.Core.GetUser(aUser);
        }
        public User GetUserData(long? aUser)
        {
            if (!aUser.HasValue)
                return Application.CurrentUser;
            else return Application.Core.GetUser(aUser.Value);
        }
        public Document GetDocument(long aDocumentID, DocumentRetrievalOptions lOptions)
        {
            Document lDocument = Application.Core.GetDocumentByID(aDocumentID, lOptions);
            if (lDocument == null)
                throw new KeyNotFoundException("No Matching Document found with ID: " + aDocumentID);
            return lDocument;
        }
        public List<PageData> GetDocumentData(Document aDocument , string AlternateFormat)
        {                        
            return GetAllPages(AlternateFormat, aDocument.DefaultRenditionOfLatestRevision);            
        }
        public Document StoreNewDocument(DocumentType aDocumentType, IEnumerable<PageData> aPageData, IEnumerable<Keyword> aKeywords, DateTime? aDocumentDate, FileType aFileType, bool aExpandKeyset = true, bool aUseDefaultForRequiredKeys = true)
        {
            if (aDocumentType == null)
            {
                throw new ArgumentNullException("aDocumentType", "Must specify a Document Type to store a new document");
            }
            if (aFileType == null)
            {
                throw new ArgumentNullException("aFileType", "Must specify a File Type to store a new document");
            }
            if (!aDocumentDate.HasValue)
                aDocumentDate = DateTime.Now;
            StoreNewDocumentProperties lStoreProperties = Application.Core.Storage.CreateStoreNewDocumentProperties(aDocumentType, aFileType);
            lStoreProperties.DocumentDate = aDocumentDate;

            if (aKeywords != null)
            {
                IEnumerable<Keyword> definedKeywords = aKeywords.Where(x => aDocumentType.IsDefined(x));
                List<Keyword> lExpandedKeywordList;
                if (aExpandKeyset)
                {
                    lExpandedKeywordList = ExpandAutoFillKeyset(definedKeywords, aDocumentType, aUseDefaultForRequiredKeys);
                }
                else
                {
                    lExpandedKeywordList = new List<Keyword>();
                }

                IEnumerable<Keyword> lFullList = definedKeywords.Concat(lExpandedKeywordList);
                IEnumerable<Keyword> lRequiredDefaults = VerifyRequiredKeyword(lFullList, aDocumentType.KeywordTypesRequiredForArchival, aUseDefaultForRequiredKeys);
                AddKeywords(lStoreProperties, lFullList);
                AddKeywords(lStoreProperties, lRequiredDefaults);
            }
            return Application.Core.Storage.StoreNewDocument(aPageData, lStoreProperties);
        }

        public void DeleteDocument(long aDocument)
        {
            DeleteDocument(GetDocument(aDocument, DocumentRetrievalOptions.None));
        }
        public void DeleteDocument(Document aDocument)
        {
            Application.Core.Storage.DeleteDocument(aDocument);
        }

        public Note AddNote(long aDocument, string aNoteType, string aText, long? width, long? height, long? xPos, long? yPos, long? aPage = null)
        {
            Document lDocument = GetDocument(aDocument, DocumentRetrievalOptions.None);
            NoteType lNoteType = GetNoteType(aNoteType);
            return AddNote(lDocument, lNoteType, aText, width, height, xPos, yPos, aPage);
        }
        public Note AddNote(long aDocument, long aNoteType, string aText, long? width, long? height, long? xPos, long? yPos, long? aPage = null)
        {
            Document lDocument = GetDocument(aDocument, DocumentRetrievalOptions.None);
            NoteType lNoteType = GetNoteType(aNoteType);
            return AddNote(lDocument, lNoteType, aText, width, height, xPos, yPos, aPage);
        }
        public Note AddNote(Document aDocument, NoteType aNoteType, string aText, long?width, long?height, long?xPos, long ?yPos, long? aPage = null)
        {
            
            using (DocumentLock lDocumentLock = aDocument.LockDocument())
            {
                if (lDocumentLock.Status != DocumentLockStatus.LockObtained)
                {
                    throw new DocumentLockedException(aDocument, lDocumentLock);
                }

                NoteModifier lModifier = aDocument.CreateNoteModifier();
                NoteProperties lProperties = lModifier.CreateNoteProperties();
                if (width.HasValue && height.HasValue)
                    lProperties.Size = lModifier.CreateNoteSize(width.Value, height.Value);

                lProperties.Position = lModifier.CreateNotePosition(xPos.GetValueOrDefault(0), yPos.GetValueOrDefault(0));                
                lProperties.PageNumber = aPage;
                lProperties.Text = aText;
                Note lNote = aNoteType.CreateNote(lProperties);
                lModifier.AddNote(lNote);
                lModifier.ApplyChanges();
                return lNote;
            }
        }
        public Document UpdateDocumentKeywords(
            long aDocumentId,
            IEnumerable<Keyword> aKeywords = null,
            bool aExpandKeyset = true, bool aUseDefaultForRequiredKeys = true)
        {
            Document lDocument = GetDocument(aDocumentId, DocumentRetrievalOptions.LoadKeywords);
            return UpdateDocumentKeywords(lDocument, aKeywords, aExpandKeyset, aUseDefaultForRequiredKeys);
        }
        public Document UpdateDocumentKeywords(
            Document aDocument,
            IEnumerable<Keyword> aKeywords = null,
            bool aExpandKeyset = true, bool aUseDefaultForRequiredKeys = true)
        {
            using (DocumentLock lDocumentLock = aDocument.LockDocument())
            {
                if (lDocumentLock.Status == DocumentLockStatus.AlreadyLocked)
                {
                    throw new DocumentLockedException(aDocument, lDocumentLock);
                }

                KeywordModifier lKeywordModifier = aDocument.CreateKeywordModifier();
                EditKeywords(lKeywordModifier, aDocument, aDocument.DocumentType, aKeywords, aExpandKeyset, aUseDefaultForRequiredKeys);
                lKeywordModifier.ApplyChanges();
            }
            return aDocument;
        }


        public Document ChangeDocumentType(
            long aDocumentId, long aNewDocumentTypeId,
            IEnumerable<Keyword> aKeywords = null,
            bool aExpandKeyset = true, bool aUseDefaultForRequiredKeys = true)
        {
            Document lDocument = GetDocument(aDocumentId, DocumentRetrievalOptions.LoadKeywords);
            DocumentType lNewDocType = GetDocumentType(aNewDocumentTypeId);
            return ChangeDocumentType(lDocument, lNewDocType, aKeywords, aExpandKeyset, aUseDefaultForRequiredKeys);
        }
        public Document ChangeDocumentType(
            Document aDocument, DocumentType aNewDocumentType,
            IEnumerable<Keyword> aKeywords = null,
            bool aExpandKeyset = true, bool aUseDefaultForRequiredKeys = true)
        {
            if (aNewDocumentType==null || aNewDocumentType.Equals(aDocument.DocumentType))
            {
                return UpdateDocumentKeywords(aDocument, aKeywords, aExpandKeyset, aUseDefaultForRequiredKeys);
            }

            using (DocumentLock lDocumentLock = aDocument.LockDocument())
            {
                if (lDocumentLock.Status == DocumentLockStatus.AlreadyLocked)
                {
                    throw new DocumentLockedException(aDocument, lDocumentLock);
                }

                ReindexProperties lReindexProperties = Application.Core.Storage.CreateReindexProperties(aDocument, aNewDocumentType);
                EditKeywords(lReindexProperties, aDocument, aNewDocumentType, aKeywords, aExpandKeyset, aUseDefaultForRequiredKeys);
                return Application.Core.Storage.ReindexDocument(lReindexProperties);
            }            
        }
   
        public NoteType GetNoteType(string aNoteType)
        {
            NoteType lType = Application.Core.NoteTypes.Find(aNoteType);
            if (lType == null)
                throw new ArgumentOutOfRangeException("aNoteType", aNoteType, "Invalid Note Type: " + aNoteType);
            return lType;
        }
        public NoteType GetNoteType(long aNoteType)
        {
            NoteType lType = Application.Core.NoteTypes.Find(aNoteType);
            if (lType == null)
                throw new ArgumentOutOfRangeException("aNoteType", aNoteType, "Invalid Note Type: " + aNoteType);
            return lType;
        }
        public KeywordType GetKeywordType(string aKeywordType, DocumentType aDocumentType = null)
        {
            KeywordType lType;

            if (aDocumentType != null)
                lType = aDocumentType.KeywordRecordTypes.FindKeywordType(aKeywordType);
            else
                lType = Application.Core.KeywordTypes.Find(aKeywordType);

            if (lType == null)
                throw new ArgumentOutOfRangeException("aKeywordType", aKeywordType, "Invalid Keyword type: " + aKeywordType);

            return lType;
        }
        public KeywordType GetKeywordType(long aKeywordType, DocumentType aDocumentType = null)
        {
            KeywordType lType;

            if (aDocumentType != null)
                lType = aDocumentType.KeywordRecordTypes.FindKeywordType(aKeywordType);
            else
                lType = Application.Core.KeywordTypes.Find(aKeywordType);

            if (lType == null)
                throw new ArgumentOutOfRangeException("aKeywordType", aKeywordType, "Invalid Keyword type: " + aKeywordType);

            return lType;
        }
        public KeywordRecordType GetKeywordRecordType(string aKeywordRecordType, DocumentType aDocumentType = null)
        {
            KeywordRecordTypeList lKeywordRecordTypeList;
            if (aDocumentType == null)
                lKeywordRecordTypeList = Application.Core.KeywordRecordTypes;
            else
                lKeywordRecordTypeList = aDocumentType.KeywordRecordTypes;

            KeywordRecordType lType = lKeywordRecordTypeList.Find(aKeywordRecordType);
            if (lType == null)
                throw new ArgumentOutOfRangeException("Invalid Keyword Record Type: " + aKeywordRecordType);

            return lType;
        }
        public KeywordRecordType GetKeywordRecordType(long aKeywordRecordType, DocumentType aDocumentType = null)
        {
            KeywordRecordTypeList lKeywordRecordTypeList;
            if (aDocumentType == null)
                lKeywordRecordTypeList = Application.Core.KeywordRecordTypes;
            else
                lKeywordRecordTypeList = aDocumentType.KeywordRecordTypes;

            KeywordRecordType lType = lKeywordRecordTypeList.Find(aKeywordRecordType);
            if (lType == null)
                throw new ArgumentOutOfRangeException("Invalid Keyword Record Type: " + aKeywordRecordType);

            return lType;
        }

        public DocumentType GetDocumentType(string aDocumentType)
        {
            DocumentType lDocumentType = Application.Core.DocumentTypes.Find(aDocumentType);
            if (lDocumentType == null)
                throw new ArgumentOutOfRangeException("aDocumentType", aDocumentType, "Invalid Document Type: " + aDocumentType);
            return lDocumentType;
        }
        public DocumentType GetDocumentType(long aDocumentType)
        {
            DocumentType lDocumentType = Application.Core.DocumentTypes.Find(aDocumentType);
            if (lDocumentType == null)
                throw new ArgumentOutOfRangeException("aDocumentType", aDocumentType, "Invalid Document Type: " + aDocumentType);
            return lDocumentType;
        }


        public FileType GetFileType(DocumentType aDocumentType, string aFileType = null)
        {
            if (string.IsNullOrWhiteSpace(aFileType))
            {
                return aDocumentType.DefaultFileType;
            }
            else
            {
                FileType lFileType = Application.Core.FileTypes.Find(aFileType);
                if (lFileType == null)
                    throw new ArgumentOutOfRangeException("aFileType", aFileType, "Invalid File Type: " + aDocumentType);
                return lFileType;
            }
        }
        public FileType GetFileTypeByExtension(string aExtension)
        {
            Hyland.Common.Core.ExtensionToFileTypeMimeType lTypeMapping = null;
            if (!Hyland.Common.Core.Utility.FileTypeMappingUtility.GetBaseExtensionToFileTypeMimeTypeMappings().TryGetValue(aExtension, out lTypeMapping)
                && !Hyland.Common.Core.Utility.FileTypeMappingUtility.GetBaseExtensionToFileTypeMimeTypeMappings().TryGetValue(aExtension.ToUpper(), out lTypeMapping)
                && !Hyland.Common.Core.Utility.FileTypeMappingUtility.GetBaseExtensionToFileTypeMimeTypeMappings().TryGetValue(aExtension.ToLower(), out lTypeMapping))
            {
                throw new ArgumentOutOfRangeException("aExtension", aExtension, "Invalid File Extension: " + aExtension);
            }
            return GetFileType(null, lTypeMapping.FileTypeID);
        }
        public FileType GetFileType(DocumentType aDocumentType, long? aFileType = null)
        {
            if (!aFileType.HasValue)
            {
                return aDocumentType.DefaultFileType;
            }
            else
            {
                FileType lFileType = Application.Core.FileTypes.Find(aFileType.Value);
                if (lFileType == null)
                    throw new ArgumentOutOfRangeException("aFileType", aFileType, "Invalid File Type: " + aDocumentType);
                return lFileType;
            }
        }
        protected DataProvider GetDataProvider(string providerName)
        {
            if (string.IsNullOrWhiteSpace(providerName))
                return null;
            //use reflection to grab the right aProvider   
            try
            {
                PropertyInfo propertyInfo = Application.Core.Retrieval.GetType().GetProperty(providerName);
                if (propertyInfo == null)
                    throw new ArgumentOutOfRangeException("Provider Name", "Invalid data format");

                return propertyInfo.GetValue(Application.Core.Retrieval, null) as DataProvider;
            }
            catch
            {
                throw new ArgumentOutOfRangeException("Provider Name", "Invalid data format");
            }            
        }

        public List<PageData> GetAllPages(string aProviderName, Rendition rendition)
        {
            if (string.IsNullOrWhiteSpace(aProviderName))
            {
                if(rendition.GetNativeFormatAllowed())
                    aProviderName = "Native";
                else aProviderName = "Default"; 
                
            }
            DataProvider lProvider = GetDataProvider(aProviderName);
            
            return GetAllPages(lProvider, rendition);
        }
        public List<PageData> GetAllPages(DataProvider aProvider, Rendition rendition)
        {
            if (aProvider == null)
                throw new ArgumentNullException("provider", "No DataProvider specified");
            if (rendition == null)
                throw new ArgumentNullException("rendition", "No document rendition specified");
            List<PageData> pagesResults = new List<PageData>();

            PageRangeSet lRangeSet = aProvider.CreatePageRangeSet();
            lRangeSet.AddRange(1, rendition.NumberOfPages);
            MethodInfo GetPagesMethod = aProvider.GetType().GetMethod("GetPages", new Type[] { typeof(Rendition), typeof(PageRangeSet) });
            if (GetPagesMethod == null)
                throw new Exception("Specified DataProvider does not define file access");
            try
            {
                pagesResults.AddRange(GetPagesMethod.Invoke(aProvider, new object[] { rendition, lRangeSet }) as PageDataList);                    
            }
            catch (TargetInvocationException ex)
            {
                throw ex.InnerException;
            }
            
            return pagesResults;
        }

        public PageData CreatePageData(Stream aPage, string aExtension)
        {
            return Application.Core.Storage.CreatePageData(aPage, aExtension);
        }
        public PageData CreatePageData(Stream aPage, FileType aFileType)
        {
            return Application.Core.Storage.CreatePageData(aPage, aFileType.Extension);
        }
        public PageData CreatePageData(string aPage)
        {
            return Application.Core.Storage.CreatePageData(aPage);
        }
        public List<PageData> CreatePageData(IEnumerable<Stream> aPageList, IEnumerable<string> aExtensions)
        {
            List<PageData> lPageDatas = new List<PageData>();
            if (aPageList != null && aExtensions != null)
            {
                foreach (PageData lPage in aPageList.Zip(aExtensions, (x, y) => CreatePageData(x, y)))
                {
                    lPageDatas.Add(lPage);
                }
            }
            return lPageDatas;
        }
        public List<PageData> CreatePageData(IEnumerable<Stream> aPageList, IEnumerable<FileType> aFileTypes)
        {
            return CreatePageData(aPageList, aFileTypes.Select(x => x.Extension));
        }
        public List<PageData> CreatePageData(IEnumerable<Stream> aPageList, FileType aFileType)
        {
            List<PageData> lPageDatas = new List<PageData>(); ;
            if (aPageList != null)
            {
                foreach (Stream lPage in aPageList)
                {
                    lPageDatas.Add(CreatePageData(lPage, aFileType));
                }
            }
            return lPageDatas;
        }
        public List<PageData> CreatePageData(IEnumerable<string> aPageList)
        {
            List<PageData> lPageDatas = new List<PageData>(); ;
            if (aPageList != null)
            {
                foreach (string lPage in aPageList)
                {
                    lPageDatas.Add(CreatePageData(lPage));
                }
            }
            return lPageDatas;
        }

        private List<Keyword> ExpandAutoFillKeyset(IEnumerable<Keyword> aKeywordList, DocumentType aDocumentType, bool aExpandDefault)
        {
            List<Keyword> lResult = new List<Keyword>();
            Keyset lKeyset = (aDocumentType == null) ? null : aDocumentType.AutoFillKeyset;            
            if (aKeywordList == null || lKeyset == null || !aDocumentType.IsDefined(lKeyset.PrimaryKeywordType))
                return lResult;
            KeywordType lPrimaryKeywordType = lKeyset.PrimaryKeywordType;

            //if the primary keyword type is a required keyword type, do not expand the default
            aExpandDefault &= aDocumentType.IsRequired(lPrimaryKeywordType);

            
            List<Keyword> lPrimaryKeys = aKeywordList.Where(x => x != null && x.KeywordType.Equals(lPrimaryKeywordType)).ToList();
            if (!lPrimaryKeys.Any() && aExpandDefault)
            {
                lPrimaryKeys.Add(CreateKeyword(lPrimaryKeywordType, lPrimaryKeywordType.DefaultValue));
            }
            if (lPrimaryKeys.Any())
            {
                KeysetDataList lKeysetDataList = lKeyset.GetKeysetData(lPrimaryKeys);
                foreach (KeysetData lKeysetData in lKeysetDataList)
                {
                    foreach (Keyword lAutofillKeyword in lKeysetData.Keywords)
                    {
                        if (!lAutofillKeyword.KeywordType.Equals(lKeyset.PrimaryKeywordType) && aDocumentType.IsDefined(lAutofillKeyword))
                        {
                            lResult.Add(lAutofillKeyword);
                        }
                    }
                }
            }
            
            return lResult;
        }


        
        /// <summary>
        /// Verify that all required key words are present in the collection.
        /// If UseDefaults is true, return collection of defaults to add to satisfy requirements
        /// If UseDefaults is false and a required keyword is missing, an exception will be thrown
        /// </summary>
        /// <param name="aKeywords">Collection of keywords to check</param>
        /// <param name="aRequiredKeywordTypes">Collection of REquired Keyword Types</param>        
        /// <param name="aUseDefault">If false, fail on missing required keyword. Else fill result collection with default values for missing keys</param>
        /// <returns>Collection of missing keywords created with default values.</returns>
        private List<Keyword> VerifyRequiredKeyword(IEnumerable<Keyword> aKeywords, KeywordTypeList aRequiredKeywordTypes, bool aUseDefault)
        {
            List<Keyword> lNeededDefaults = new List<Keyword>();
            //nothing to verify
            if (aRequiredKeywordTypes == null)
                return lNeededDefaults;

            if (aKeywords == null)
                aKeywords = new List<Keyword>();

            ILookup<KeywordType, Keyword> lLookup = aKeywords.ToLookup(x => x.KeywordType);
            foreach (KeywordType lRequiredKeywordType in aRequiredKeywordTypes)
            {
                if (!lLookup.Contains(lRequiredKeywordType))
                {
                    //required word not found in input or result set
                    if (aUseDefault)
                    {
                        //use default key value
                        Keyword lDefaultKeyword = CreateKeyword(lRequiredKeywordType, lRequiredKeywordType.DefaultValue);
                        lNeededDefaults.Add(lDefaultKeyword);
                    }
                    else
                    {
                        //required key is missing
                        const string lErrorFormat = "Required keyword {0} ({1}) not present in keyword set";
                        throw new Exception(string.Format(lErrorFormat, lRequiredKeywordType.Name, lRequiredKeywordType.ID));
                    }
                }
            }
            return lNeededDefaults;
        }


        /// <summary>
        /// Performs all the necessary operations on the EditableKeywordModifier
        /// Removes will not affect updates or adds.
        /// Meaning a request to update the old value of a keyword will not be affected by a request to remove the old value
        /// Similarly a request to update a keyword that is in the add request collection, will not succeed.
        /// the method also handles expansion of autofill keysets if requested, and using defaults to satisfy 
        /// retuired keywords if requested.
        /// </summary>
        /// <param name="aModifier">EditableKeywordModifier on which to operate</param>
        /// <param name="aDocument">Document on which modifications will take place</param>
        /// <param name="aUpdateRequest">Collection of keyword update requests</param>
        /// <param name="aAddRequest">Collection of keyword addition requests</param>
        /// <param name="aRemoveRequest">Collection of keyword removal requests</param>
        /// <param name="aExpandKeyset">If true, autofill keysets will be expanded</param>
        /// <param name="aUseDefaultForRequiredKeys">If true, default keyword values will be used to satisfy any missing required keywords</param>       
        private void EditKeywords(EditableKeywordModifier aModifier, Document aDocument,
            IEnumerable<KeywordUpdate> aUpdateRequest, IEnumerable<Keyword> aAddRequest, IEnumerable<Keyword> aRemoveRequest,
            bool aExpandKeyset, bool aUseDefaultForRequiredKeys)
        {            
            if (aUpdateRequest == null)
                aUpdateRequest = new List<KeywordUpdate>();
            if (aAddRequest == null)
                aAddRequest = new List<Keyword>();
            if (aRemoveRequest == null)
                aRemoveRequest = new List<Keyword>();
            DocumentType lDocumentType = aDocument.DocumentType;
            KeywordRecordType lStandaloneKeywordRecordType = lDocumentType.KeywordRecordTypes.Find(STANDALONE_KEYWORD_RECORD_TYPE);
            KeywordRecord lStandaloneKeywordRecord = aDocument.KeywordRecords.Find(lStandaloneKeywordRecordType);

            List<Keyword> lExpectedUnmodified = new List<Keyword>();
            List<Keyword> lExpectedRemoved = new List<Keyword>();
            List<Keyword> lExpectedAddedByUpdate = new List<Keyword>();
            ILookup<Keyword, Keyword> lLookupUpdate = aUpdateRequest.ToLookup(x => x.OldValue, y => y.NewValue);

            foreach (Keyword lExistingKeyword in lStandaloneKeywordRecord.Keywords)
            {
                IEnumerable<Keyword> lNewValues = lLookupUpdate[lExistingKeyword];
                if (lNewValues.Any())
                {
                    lExpectedAddedByUpdate.AddRange(lNewValues);
                }
                else if (aRemoveRequest.Contains(lExistingKeyword))
                {
                    lExpectedRemoved.Add(lExistingKeyword);
                }
                else
                {
                    lExpectedUnmodified.Add(lExistingKeyword);
                }
            }
            IEnumerable<Keyword> lUpdateAndAdd = lExpectedAddedByUpdate.Concat(aAddRequest);
            IEnumerable<Keyword> lExpanded = (aExpandKeyset) ? ExpandAutoFillKeyset(lUpdateAndAdd, lDocumentType, aUseDefaultForRequiredKeys) : new List<Keyword>();
            IEnumerable<Keyword> lExpectedResult = lExpectedUnmodified.Concat(lUpdateAndAdd).Concat(lExpanded);
            IEnumerable<Keyword> lRequiredDefault = VerifyRequiredKeyword(lExpectedResult, lDocumentType.KeywordTypesRequiredForArchival, aUseDefaultForRequiredKeys);

            foreach (Keyword lKeyword in lExpectedRemoved)
            {
                aModifier.RemoveKeyword(lKeyword);
            }
            foreach (KeywordUpdate lUpdate in aUpdateRequest)
            {
                aModifier.UpdateKeyword(lUpdate.OldValue, lUpdate.NewValue);
            }
            AddKeywords(aModifier, aAddRequest);
            AddKeywords(aModifier, lExpanded);
            AddKeywords(aModifier, lRequiredDefault);
        }

        private void EditKeywords(EditableKeywordModifier aModifier, Document aDocument, DocumentType aNewDocumentType, IEnumerable<Keyword> aKeywords, bool aExpandKeyset, bool aUseDefaultForRequiredKeys)
        {
            if (aKeywords == null)
                aKeywords = new List<Keyword>();

            DocumentType lDocumentType = aDocument.DocumentType;
            KeywordRecordType lStandaloneKeywordRecordType = aNewDocumentType.KeywordRecordTypes.Find(STANDALONE_KEYWORD_RECORD_TYPE);

            IEnumerable<Keyword> lDefinedRequestedKeywords = aKeywords.Where(x => aNewDocumentType.IsDefined(x));
            IEnumerable<Keyword> lExpansion = (aExpandKeyset) ? ExpandAutoFillKeyset(lDefinedRequestedKeywords, aNewDocumentType, aUseDefaultForRequiredKeys) : new List<Keyword>();
            IEnumerable<Keyword> lFullRequest = aKeywords.Concat(lExpansion);
            ILookup<KeywordType, Keyword> lNewKeywordLookup = lFullRequest.ToLookup(x => x.KeywordType);
            ILookup<KeywordType, Keyword> lExistingKeywordLookup = aDocument.GetStandaloneKeywords().ToLookup(x => x.KeywordType);
            List<Keyword> lExpectedUnmodified = new List<Keyword>();
            List<Keyword> lExpectedRemoved = new List<Keyword>();


            foreach (IGrouping<KeywordType, Keyword> lGroup in lExistingKeywordLookup)
            {
                if (!aNewDocumentType.IsDefined(lGroup.Key) || lNewKeywordLookup.Contains(lGroup.Key))
                {
                    lExpectedRemoved.AddRange(lGroup);
                }
                else
                {
                    lExpectedUnmodified.AddRange(lGroup);
                }
            } 

            IEnumerable<Keyword> lRequestedAndUnmodified = lFullRequest.Concat(lExpectedUnmodified);
            IEnumerable<Keyword> lRequiredDefault = VerifyRequiredKeyword(lRequestedAndUnmodified, lDocumentType.KeywordTypesRequiredForArchival, aUseDefaultForRequiredKeys);

            foreach (Keyword lKeyword in lExpectedRemoved)
            {
                aModifier.RemoveKeyword(lKeyword);
            }

            AddKeywords(aModifier, lFullRequest);
            AddKeywords(aModifier, lRequiredDefault);
        }
        private void AddKeywords(AddOnlyKeywordModifier aModifier, IEnumerable<Keyword> aAddKeywordList)
        {
            foreach (Keyword lKeyword in aAddKeywordList)
            {
                aModifier.AddKeyword(lKeyword);
            }
        }

        public Keyword CreateKeyword(long aKeywordType, string aValue, DocumentType aDocumentType = null)
        {
            return CreateKeyword(GetKeywordType(aKeywordType, aDocumentType), aValue);
        }

        public Keyword CreateKeyword(string aKeywordType, string aValue, DocumentType aDocumentType = null)
        {
            return CreateKeyword(GetKeywordType(aKeywordType, aDocumentType), aValue);
        }


        public List<Keyword> CreateKeywordCollection(
            IEnumerable<string> aKeywordTypeCollection, IEnumerable<string> aKeywordValueCollection,
            DocumentType aDocumentType = null
        )
        {
            if (aKeywordTypeCollection == null && aKeywordValueCollection == null)
                return new List<Keyword>();
            else if (aKeywordTypeCollection.Count() != aKeywordValueCollection.Count())
            {
                throw new Exception("Keyword type and value collections have different number of elements");
            }
            List<Keyword> result = new List<Keyword>();
            foreach (KeyValuePair<string,string> item in aKeywordTypeCollection.Zip(aKeywordValueCollection, (x,y) => new KeyValuePair<string,string>(x, y)))
            {
                try
                {
                    KeywordType lType = GetKeywordType(item.Key, aDocumentType);
                    Keyword lKeyword = CreateKeyword(lType, item.Value);
                    result.Add(lKeyword);
                }
                catch (ArgumentOutOfRangeException )
                {
                    //skip undefined keywords
                    continue;
                }                
            }
            return result;
        }

        public List<Keyword> CreateKeywordCollection(
            IEnumerable<long> aKeywordTypeColection, IEnumerable<string> aKeywordValueCollection,
            DocumentType aDocumentType = null
        )
        {
            return aKeywordTypeColection.Zip(aKeywordValueCollection, (x, y) => CreateKeyword(x, y, aDocumentType)).ToList();
        }

        public static Keyword CreateKeyword(KeywordType aKeywordType, string aValue)
        {
            Keyword key = null;
            if (string.IsNullOrWhiteSpace(aValue))
            {
                key = aKeywordType.CreateBlankKeyword();
            }
            else
            {
                switch (aKeywordType.DataType)
                {
                    case KeywordDataType.Currency:
                    case KeywordDataType.Numeric20:
                        decimal decVal = decimal.Parse(aValue);
                        key = aKeywordType.CreateKeyword(decVal);
                        break;
                    case KeywordDataType.Date:
                    case KeywordDataType.DateTime:
                        DateTime dateVal = DateTime.Parse(aValue);
                        key = aKeywordType.CreateKeyword(dateVal);
                        break;
                    case KeywordDataType.FloatingPoint:
                        double dblVal = double.Parse(aValue);
                        key = aKeywordType.CreateKeyword(dblVal);
                        break;
                    case KeywordDataType.Numeric9:
                        long lngVal = long.Parse(aValue);
                        key = aKeywordType.CreateKeyword(lngVal);
                        break;
                    default:
                        key = aKeywordType.CreateKeyword(aValue);
                        break;
                }
            }
            //validate keyword dataset
            if (aKeywordType.KeywordMustExistInDataSet)
            {                
                if (key.IsBlank || aKeywordType.GetKeywordDataSet().Find(x => x.Value.Equals(key.Value)) == null)
                {
                    if(key.IsBlank)
                        aValue = "<null>";

                    const string lErrorFormat = "{0} is not a valid value for keyword {1} ({2})";
                    throw new ArgumentOutOfRangeException("aValue", aValue, string.Format(lErrorFormat, aValue, aKeywordType.Name, aKeywordType.ID));
                }
            }
            return key;
        }
    }
    

}
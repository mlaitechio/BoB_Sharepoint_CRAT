export class Constants {
  public static readonly headerText = "Circular Repository Access Tool";
  public static readonly searchText = "Search";
  //public static readonly documentCategoryList = "DocumentCategory";
  public static readonly circularList = "CircularRepository";

  public static readonly commentsAuditLogs = `CommentsAuditLogs`;

  public static readonly emailList = "EmailTemplates";

  public static readonly sharedDocuments = "Shared Documents";

  public static readonly configurationList = "Configuration";
  public static readonly circularNo = "BOB:BR:{0}:";
  public static readonly filterString = "CircularStatus eq 'Published' or CircularStatus eq 'Archived'";
  public static readonly filterCircularNumber = `CircularNumber eq '{0}'`;


  public static readonly makerGroup = "Maker";
  public static readonly checkerGroup = "Checker";
  public static readonly complianceGroup = "Compliance";
  public static readonly rejectedGroup = "Rejected";
  public static readonly headerCircularUpload = "{0} Circular Upload";

  /**
  |--------------------------------------------------
  | Edit Circular ,Pending Compliance & Pending Checker Filter string one more filter criteria department needs to be added
  |--------------------------------------------------
  */
  public static readonly editCircularFilterString = "(Department eq '{0}' or Department eq null) and (CircularStatus eq 'Draft' or CircularStatus eq 'Submitted to Compliance' or CircularStatus eq 'Submitted to Checker' or CircularStatus eq 'Comments from Compliance' or CircularStatus eq 'Comments from Checker')"
  public static readonly compliancePendingFilterString = "(Department eq '{0}' or Department eq null) and (CircularStatus eq 'Submitted to Compliance') and (Compliance eq 'Yes')";
  public static readonly checkerPendingFilterString = "(Department eq '{0}' or Department eq null) and (CircularStatus eq 'Approved by Compliance' or CircularStatus eq 'Submitted to Checker')";
  public static readonly rejectedFilterString = "(Department eq '{0}' or Department eq null) and (CircularStatus eq 'Rejected')";
  /**
  |--------------------------------------------------
  | MegaMenu Constants
  |--------------------------------------------------
  */

  public static readonly lblHome = "Home";
  public static readonly lblAddCircular = "Add Circular";
  public static readonly lblEditCircular = "Edit Circular";
  public static readonly lblViewCircular = "View Circular";
  public static readonly lblPendingCompliance = "Pending Compliance Request";
  public static readonly lblPendingChecker = "Pending Checker Request";
  public static readonly lblRejectedRequest = "Rejected Request";

  /**
  |--------------------------------------------------
  | Form Field Display Name
  |--------------------------------------------------
  */

  public static readonly circularNumber = `Circular Number`;
  public static readonly subFileNo = `SubFile Code`;
  public static readonly circularInitator = `Circular Initiator`;
  public static readonly issuedFor = `Issued For`;
  public static readonly category = `Category`;
  public static readonly classification = `Classification`;
  public static readonly type = "Type";
  public static readonly expiry = "Expiry";
  public static readonly subject = `Subject`;
  public static readonly keyWords = `Keywords`;
  public static readonly department = `Department`;
  public static readonly gist = `Gist`;
  public static readonly compliance = `Regulatory Compliance`;
  public static readonly faqs = "FAQs";
  public static readonly sop = "SOP";
  public static readonly supportingDocument = `Supporting Documents`;
  public static readonly limited = "Limited";
  public static readonly unlimited = "Unlimited";
  public static readonly lblCompliance = "This circular will go to compliance department";
  public static readonly lblYes = "Yes";
  public static readonly lblNo = "No";
  public static readonly lblCommentsMaker = "Comments Maker";
  public static readonly lblCommentsChecker = "Comments Checker";
  public static readonly lblCommentsCompliance = "Comments Compliance";
  public static readonly goBack = `Go Back`;
  public static readonly delete = `Delete`;
  public static readonly publishedYear = "Published Year";
  public static readonly validationAlertTitle = "Validation Alert!";
  public static readonly validationAlertMessage = "Please input all fields marked as *";
  public static readonly validationAlertMessageFileSize = "File Size is greater than 5MB";
  public static readonly validationAlertMessageFileType = "File type is not .pdf or .docx";
  public static readonly validationCircularNumber = "Circular Number Already exist.";
  public static readonly deleteCircularTitle = `Delete Circular File`;
  public static readonly deleteCircularMessage = `Are you sure you want to delete the file?`;

  public static readonly searchSupportingCirculars = `Search Circulars`;


  public static readonly colCircularRepository = "Id,Subject,ArchivalDate,PublishedDate,CircularStatus,Category,IssuedFor,MigratedDepartment,Department,MakerCommentsHistory,CheckerCommentsHistory,ComplianceCommentsHistory,IsMigrated,CircularNumber,Classification,MigratedOriginator,Author/Title,Author/Id,Author/EMail"
  public static readonly expandColCircularRepository = "Author";
  public static readonly adSelectedColumns = "id,mail,region,title,position,grade,zone,displayName,department,employeeId,extensions,businessPhones,title,employeeNumber";
  public static readonly configSelectColumns = `Title,Limit,ID,ToolTip`;
  public static readonly configVal = {
    SupportingDocuments: "SupportingDocuments",
    SOPFileUpload: "SOPFileUpload",
    SOPFileMaxSizeinMB: "SOPFileMaxSizeinMB",
    SubjectMaxWord: "SubjectMaxWord",
    GistMaxWord: "GistMaxWord",
    FAQMaxWord: "FAQMaxWord",
    MakerCommentsMaxWord: "MakerCommentsMaxWord",
    ComplianceCommentsMaxWord: "ComplianceCommentsMaxWord",
    CheckerCommentsMaxWord: "CheckerCommentsMaxWord",
    SupportingDocUpload: "SupportingDocUpload",
    SupportingDocSizeinMB: "SupportingDocSizeinMB",
    TemplateToolTipText: "TemplateToolTipText",
    KeyWordsToolTipText: "KeyWordsToolTipText",
    SupportingDocToolTipText: "SupportingDocToolTipText",
    AllYearsToolTipText: "AllYearsToolTipText",
    PreviousYearToolTipText: "PreviousYearToolTipText"
  }
  public static readonly hindiBarodaPedia = "बड़ौदापीडिया";
  public static readonly hindiSearchHeader = "परिपत्र का विवरण दर्ज करें";
  public static readonly engSearchHeader = "Enter details of circular";
  public static readonly hindiSearchCircular = "परिपत्र खोजें";
  public static readonly engSearchCircular = "Search Circular";

  public static readonly sorting = "Sorting";

  public static readonly infoPDFText = "This data is as on {0} of view/download.For updated information refer to circular/master circular on portal.";

  //Archived Status circulars with limited period after expiry date will have this text on header of pdf with red font
  public static readonly warninglimitedPDFText = "The circular stands archived with effect from {0} consequent upon validity expiry.";

  //Archived Status circulars with Unlimited period after incorporation in master circular will have this text on header of pdf with red font
  public static readonly warningUnlimitedPDFText = "The circular stands archived with effect from {0} consequent upon its incorporation in the master circular."

  /**
  |--------------------------------------------------
  | Check box Label
  |--------------------------------------------------
  */

  public static readonly lblContains = "Contains";
  public static readonly lblStartsWith = "Starts With";
  public static readonly lblEndsWith = "Ends With";
  public static readonly lblMaster = "Master";
  public static readonly lblCircular = "Circular";
  public static readonly lblIndia = "India";
  public static readonly lblGlobal = "Global";
  public static readonly lblComplianceYes = "Yes";
  public static readonly lblComplianceNo = "No";
  public static readonly lblIntimation = "Intimation";
  public static readonly lblInformation = "Information";
  public static readonly lblAction = "Action";

  /**
  |--------------------------------------------------
  | RefinableString00 -> CircularNumber
    RefinableString01 -> Subject
    RefinableString02 -> Migrated Department
    RefinableString03 -> Department
    RefinableString04 -> Category
    RefinableString05 -> IsMigrated 
    RefinableString06 -> Classification
    RefinableDate00 -> PublishedDate  
    RefinableString07 -> CircularStatus
    RefinableString08 -> IssuedFor
    RefinableString09 -> Compliance
    RefinableString10 -> Keywords
   RefinableString100 ->Gist(Summary)
   RefinableString101 ->FAQ
  |--------------------------------------------------
  */
  public static readonly selectedSearchProperties = ["ListItemID", "Modified", "LastModifiedTime", "RefinableString00", "RefinableString01", "RefinableString02", "RefinableString03", "RefinableString04", "RefinableString05", "RefinableString06", "RefinableDate00", "Created", "RefinableString07", "RefinableString08", "RefinableString09", "RefinableString10", "RefinableString100", "RefinableString101", "RefinableString102"]
  public static readonly filterSearchProperties = ["RefinableString00", "RefinableString01", "RefinableString02", "RefinableString03", "RefinableDate00", "RefinableString07", "RefinableString08", "RefinableString09", "RefinableString10", "RefinableString100", "RefinableString101", "RefinableString102"];

  /**
   * Managed Metadata Properties
 |--------------------------------------------------
 | RefinableString00 -> CircularNumber
   RefinableString01 -> Subject
   RefinableString02 -> Migrated Department
   RefinableString03 -> Department
   RefinableString04 -> Category
   RefinableString05 -> IsMigrated 
   RefinableString06 -> Classification
   RefinableDate00 -> PublishedDate  
   RefinableString07 -> CircularStatus
   RefinableString08 -> IssuedFor
   RefinableString09 -> Compliance
   RefinableString10 -> Keywords
   RefinableString100 ->Gist(Summary)
   RefinableString101 ->FAQ
 |--------------------------------------------------
 */

  public static readonly managePropListItemID = "ListItemID";
  public static readonly managePropCircularNumber = "RefinableString00"
  public static readonly managePropSubject = "RefinableString01"
  public static readonly managePropMigratedDepartment = "RefinableString02"
  public static readonly managePropDepartment = "RefinableString03"
  public static readonly managePropCategory = "RefinableString04";
  public static readonly managePropIsMigrated = "RefinableString05";
  public static readonly managePropClassification = "RefinableString06";
  public static readonly managePropCircularStatus = "RefinableString07";
  public static readonly managePropIssuedFor = "RefinableString08";
  public static readonly managePropCompliance = "RefinableString09";
  public static readonly managePropKeywords = "RefinableString10";
  public static readonly managePropSummary = "RefinableString100";
  public static readonly managePropFAQ = "RefinableString101";
  public static readonly managePropCircularType = "RefinableString102";
  public static readonly managePropPublishedDate = "RefinableDate00"

  /**
  |--------------------------------------------------
  | SharePoint Fields used in ListView
  |--------------------------------------------------
  */

  public static readonly colSubject = "Subject";
  public static readonly colPublishedDate = "PublishedDate";
  public static readonly colMigratedDepartment = "MigratedDepartment";
  public static readonly colCircularNumber = "CircularNumber";
  public static readonly colClassification = "Classification";
  public static readonly colIssuedFor = "IssuedFor";
  public static readonly colCategory = "Category";
  public static readonly colSummary = "Summary";
  public static readonly colType = "Type";
  public static readonly colSupportingDoc = "Supporting Documents";
  public static readonly lblTemplate = "Template";
  public static readonly templateFolder = "Shared Documents/Template";

  /**
  |--------------------------------------------------
  | SharePoint Column Display Name
  |--------------------------------------------------
  */

  public static readonly fieldSubject = "Subject";
  public static readonly fieldCircularNumber = "Circular Number";
  public static readonly fieldCircularType = "Circular Type";
  public static readonly fieldClassification = "Classification";
  public static readonly fieldDepartment = "Department";
  public static readonly fieldCircularStatus = "Circular Status";
  public static readonly fieldCompliance = "Compliance";
  public static readonly fieldPublishedDate = "Published Date";
  public static readonly fieldCircularContent = "Circular Content";
  public static readonly fieldCircularSOP = "Circular SOP";
  public static readonly fieldCircularFAQ = "Circular FAQ";

  /**
  |--------------------------------------------------
  | Circular Status
  |--------------------------------------------------
  */

  public static readonly lblNew = `New`
  public static readonly draft = "Draft";
  public static readonly sbmtCompliance = "Submitted to Compliance";
  public static readonly sbmtChecker = "Submitted to Checker";
  public static readonly appCompliance = "Approved by Compliance";
  public static readonly appChecker = "Approved by Checker";
  public static readonly cmmtCompliance = "Comments from Compliance";
  public static readonly cmmtChecker = "Comments from Checker";
  public static readonly published = "Published";
  public static readonly archived = "Archived";
  public static readonly deleted = "Deleted";
  public static readonly expired = "Expired";
  public static readonly rejected = "Rejected";


  public static readonly loreumIPSUM = "Lorem ipsum dolor sit amet, consectetur adipiscing elit. Aliquam eget libero nec tellus facilisis blandit at at magna. Donec sed dui finibus, tincidunt ante a, malesuada ligula. Class aptent taciti sociosqu ad litora torquent per conubia nostra, per inceptos himenaeos. Nunc non iaculis erat, a semper dui. Integer porta in nunc sed molestie. Suspendisse aliquet hendrerit justo. Nullam sit amet tortor non nisl viverra venenatis."


  /**
  |--------------------------------------------------
  | Category , Keywords ,Template Tooltip
  |--------------------------------------------------
  */

  public static readonly categoryToolTip = ["Intimation: (ROI Change, Signature number, List of transporters, etc.) that require attention of user but does not direct any system or manual intervention by user.",
    "Information: (New Finacle menu, launch of new product, etc.) that require attention of user but does not direct any system related updation / intervention by user, however the functionality or product are updated at the back end.",
    "Action: (Change in SOP, Updation of customer details in finacle as per new regulatory developments, etc.) which directs any system or manual intervention by user."
  ];

  public static readonly keywordsToolTip = ["When you use keywords, search in SharePoint returns results based on exact matches of your words"]

  public static readonly templateToolTip = [""];


  /**
   * File Types
   */

  public static readonly fileTypes: string[] = ['png', 'jpg', 'jpeg', 'gif', 'doc', 'docx', 'xls', 'xlsx', 'ppt', 'pptx', 'mp4', 'pdf', 'js', 'css', 'txt', 'rtf'];
  public static readonly imageTypes: string[] = ['png', 'jpg', 'jpeg', 'gif'];
  public static readonly pdfFileType: string[] = ['pdf'];
  public static readonly videoTypes: string[] = ['mp4'];
  public static readonly officeFileTypes: string[] = ['doc', 'docx', 'xls', 'xlsx', 'ppt', 'pptx', 'csv', 'rtf'];
  public static readonly otherFileTypes: string[] = ['js', 'txt', 'css'];// 'pdf'

  public static readonly ListItemPickerSelectValue = 'Select value';
  public static readonly ListItemAttachmentsActionDeleteIconTitle = 'Delete';
  public static readonly ListItemAttachmentsactionDeleteTitle = 'Delete';
  public static readonly ListItemAttachmentsfileDeletedMsg = 'File {0} deleted';
  public static readonly ListItemAttachmentsfileDeleteError = 'Error on delete file= {0}; reason {1}';
  public static readonly ListItemAttachmentserrorLoadAttachments = 'Error on load list item attachment; reason= {0}';
  public static readonly ListItemAttachmentsconfirmDelete = 'Are you sure you want send the attachment {0} to the site recycle bin?';
  public static readonly ListItemAttachmentsdialogTitle = 'List Item Attachment';
  public static readonly ListItemAttachmentsdialogOKbuttonLabel = 'OK';
  public static readonly ListItemAttachmentsdialogCancelButtonLabel = 'Cancel';
  public static readonly ListItemAttachmentsdialogOKbuttonLabelOnDelete = 'Delete';
  public static readonly ListItemAttachmentsuploadAttachmentDialogTitle = 'Add Attachment';
  public static readonly ListItemAttachmentsuploadAttachmentButtonLabel = 'Add Attachment';
  public static readonly ListItemAttachmentsuploadAttachmentErrorMsg = 'The file {0} not attached; reason= {1}';
  public static readonly ListItemAttachmentsCommandBarAddAttachmentLabel = 'Add Attachment';
  public static readonly ListItemAttachmentsloadingMessage = 'Uploading file ...';
  public static readonly ListItemAttachmentslPlaceHolderIconText = 'List Item Attachment';
  public static readonly ListItemAttachmentslPlaceHolderDescription = 'Please Add Attachment';
  public static readonly ListItemAttachmentslPlaceHolderButtonLabel = 'Add';
  public static readonly ListReadPermission = 'Read';
  public static readonly ListContriPermission = 'Contribute Without Delete';
  public static readonly ListFullPermission = 'Full Control';
  public static readonly OwnerGroupID = 3;
  public static readonly EveryoneID = 10;
}
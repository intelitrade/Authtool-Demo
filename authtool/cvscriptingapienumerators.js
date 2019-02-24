/*ANNOTATIONEDITFLAGS*/
EDIT_ANN_NOTE =	0x0001;//Edit Note Annotation
EDIT_ANN_TICKMARK = 0x0002;//Edit Tickmark Annotation
EDIT_ANN_REFERENCE = 0x0004;//Edit Reference Annoation

/*APPBASEUNITTYPE*/
APP_BASEUNIT_INCHES	= 0;//Application Member
APP_BASEUNIT_CENTIMETERS = 1;//Application Member

/*BASEUNITTYPE*/
BASEUNIT_INCHES1000	= 0;//1000ths of an Inch
BASEUNIT_INCHES10 = 1;//10ths of an Inch
BASEUNIT_CM1000	= 2;//1000ths of a Centimeter
BASEUNIT_CM10 = 3;//10ths of a Centimeter

/*BGDOCUMENTFLAGS*/
BGDOCFLAGS_NONE = 0;//OpenBgDocument Flags
BGDOCFLAGS_READONLYCVDB = 1;//OpenBgDocument Flags
BGDOCFLAGS_READONLY = 2;//OpenBgDocument Flags
BGDOCFLAGS_NOREFCOUNT = 4;//OpenBgDocument Flags

/*BLOCK_RENUMBER_FLAGS*/
BLOCK_RENUMBER_NONE = 0;//None
BLOCK_RENUMBER_UPDATE_HIDE_SKIP = 1;//Update Hide/Skip

/*BOOLFLAG*/
FLAG_TRUE = 1;//Note Number
FLAG_FALSE = 0;//Note Number

/*BULLETTYPES*/
BULLET_CHARACTERED = 0;//Bullet Member
BULLET_NUMBERED	Bullet Member = 1;//Bullet Member

/*BUTTONBITMAPID*/
BUTTONBITMAP_LEFTARROWUP = 0;//Note Number 
BUTTONBITMAP_LEFTARROWDOWN = 1;//Note Number
BUTTONBITMAP_RIGHTARROWUP =	2;//Note Number
BUTTONBITMAP_RIGHTARROWDOWN	= 3;//Note Number
BUTTONBITMAP_UPARROWUP = 4;//Note Number
BUTTONBITMAP_UPARROWDOWN = 5;//Note Number
BUTTONBITMAP_DOWNARROWUP = 6;//Note Number
BUTTONBITMAP_DOWNARROWDOWN = 7;//Note Number
BUTTONBITMAP_GRAPHUP = 8;//Note Number
BUTTONBITMAP_GRAPHDOWN = 9;//Note Number
BUTTONBITMAP_LIGHTBULBUP = 10;//Note Number
BUTTONBITMAP_LIGHTBULBDOWN = 11;//Note Number
BUTTONBITMAP_OPENBOOKUP = 12;//Note Number
BUTTONBITMAP_OPENBOOKDOWN = 13;//Note Number
BUTTONBITMAP_REDEXCLUP = 14;//Note Number
BUTTONBITMAP_REDEXCLDOWN = 15;//Note Number
BUTTONBITMAP_INFOUP	Note = 16;//Note Number
BUTTONBITMAP_INFODOWN = 17;//Note Number
BUTTONBITMAP_FACEUP = 18;//Note Number
BUTTONBITMAP_FACEDOWN = 19;//Note Number
BUTTONBITMAP_SADFACEUP = 20;//Note Number
BUTTONBITMAP_SADFACEDOWN = 21;//Note Number
BUTTONBITMAP_QUESTIONUP = 22;//Note Number
BUTTONBITMAP_QUESTIONDOWN = 23;//Note Number
BUTTONBITMAP_CHECKUP = 24;//Note Number
BUTTONBITMAP_CHECKDOWN = 25;//Note Number
BUTTONBITMAP_ISOUP = 26;//Note Number
BUTTONBITMAP_ISODOWN = 27;//Note Number
BUTTONBITMAP_XUP = 28;//Note Number
BUTTONBITMAP_XDOWN = 29;//Note Number
BUTTONBITMAP_REDSQUAREDOWN = 30;//Note Number
BUTTONBITMAP_REDSQUAREUP = 31;//Note Number
BUTTONBITMAP_GREENSQUAREDOWN = 32;//Note Number
BUTTONBITMAP_GREENSQUAREUP = 33;//Note Number
BUTTONBITMAP_LOGODOWN = 34;//Note Number
BUTTONBITMAP_LOGOUP = 35;//Note Number

/*BUTTONCELLTYPE*/
BUTTONTYPE_TEXT	Text Button: text button	0
BUTTONTYPE_BITMAP = 1;//Custom Image Button: this is the original name (same as BUTTONTYPE_CUSTOM)
BUTTONTYPE_CUSTOM = 1;//Custom Image Button: button with static image (one of the built-in images)
BUTTONTYPE_DYNAMIC = 2;//Dynamic Image Button: button with dynamic image (one of images managed by image-manager)

/*CELLBORDERTYPE*/
DRAWBORDERTYPE_SINGLE = 1;//Border Types
DRAWBORDERTYPE_DOTTED = 2;//Border Types
DRAWBORDERTYPE_DASHEDS = 3;//Border Types
DRAWBORDERTYPE_DASHEDL = 4;//Border Types
DRAWBORDERTYPE_DOUBLES = 5;//Border Types
DRAWBORDERTYPE_DOUBLEM = 6;//Border Types
DRAWBORDERTYPE_DOUBLEL = 7;//Border Types
DRAWBORDERTYPE_DASHEDM = 8;//Border Types

/*CELLDASHPOSITION*/
DASHPOSITION_FROMLEFT = 0;//Note Number
DASHPOSITION_LEFTOFCENTER = 1;//Note Number
DASHPOSITION_RIGHTOFCENTER = 2;//Note Number
DASHPOSITION_FROMRIGHT = 3;//Note Number

/*CELLTYPE*/
TYPE_NUMERIC = 0;//Numeric Cell
TYPE_ALPHA = 1;//Alphanumeric Cell
TYPE_DATE = 2;//Date Cell
TYPE_CHECKBOX = 3;//Checkbox Cell
TYPE_RADIOBUTTON = 4;//Radio Button Cell
TYPE_BUTTON = 5;//Button Cell

/*CHECKCELLTYPE*/
CHECKTYPE_NORMAL = 0;//Normal Checkbox: normal checkbox
CHECKTYPE_PLUSMINUS = 1;//Plus/Minus Checkbox: +/- checkbox
CHECKTYPE_FOLDER = //Folder Checkbox: folder checkbox (opened/closed folder)	2;
CHECKTYPE_CUSTOM = 3;//Custom Checkbox: checkbox with static image (one of the built-in images)
CHECKTYPE_DYNAMIC = 4;//Dynamic Checkbox: checkbox with dynamic image (one of the images managed by image-manager)

/*CIRCLINGLINESTYLES*/
CIRCLING_LINE_SOLID = 0;//Script Circle Member
CIRCLING_LINE_DASH = 1;//Script Circle Member
CIRCLING_LINE_DOT = 2;//Script Circle Member
CIRCLING_LINE_DASHDOT = 3;//Script Circle Member
CIRCLING_LINE_DASHDOTDOT = 4;//Script Circle Member

/*CIRCLINGSHAPES*/
CIRCLING_SHAPE_DEFAULT = 0;//Script Circle Member
CIRCLING_SHAPE_CIRCLE = 1;//Script Circle Member
CIRCLING_SHAPE_RECTANGLE = 2;//Script Circle Member
CIRCLING_SHAPE_H_BRACKETS = 3;//Script Circle Member
CIRCLING_SHAPE_V_BRACKETS = 4;//Script Circle Member

/*CURRENCYSOURCE*/
CURRENCYSOURCE_REGIONAL = 0;//Cell Member
CURRENCYSOURCE_CELLOPTIONS = 1;//Cell Member
CURRENCYSOURCE_OVERRIDE = 2;//Cell Member
CURRENCYSOURCE_PERCENTAGE = 3;//Cell Member

/*CVCOLORS*/
BLACK = 0x00000000;//Black
LTGRAY = 0x00c0c0c0;//Light Gray
MEDGRAY = 0x00a0a0a0;//Medium Gray
DKGRAY = 0x00808080;//Dark Gray
RED = 0x000000ff;//Red
BLUE = 0x00ff3030;//Blue
LIME = 0x0000ff00;//Lime
MAROON = 0x00000084;//Maroon
GREEN = 0x00008200;//Green
OLIVE = 0x00008284;//Olive
NAVY = 0x00840000;//Navy
PURPLE = 0x00840084;//Purple
TEAL = 0x00848200;//Teal
YELLOW = 0x0000ffff;//Yellow
FUSCHIA = 0x00ff00ff;//Fuschia
AQUA = 0x00ffff00;//Aqua
WHITE = 0x00ffffff;//White

/*CVDOCUMENTOPENTYPES*/
CVDOC_NOT_OPEN = 0;//Application Member
CVDOC_OPEN_WINDOW = 1;//Application Member
CVDOC_OPEN_SCRIPT = 2;//Application Member
CVDOC_OPEN_KNOWLEDGELIB = 3;//Application Member

/*DECIMALSTYPE*/
DECTYPE_INFINITE = 1000;//Infinite Precision: Constant decimalsInfinite
DECTYPE_AUTO = 1001;//Automatic (Using Rounding and Cell Decimal Options: Constant decimalsAutomatic

/*DEFAULTOPENDOCMODE*/
DOCMODE_OPEN_DEFAULT = -1;//Default Mode
DOCMODE_OPEN_FORM = 0;//Form Mode
DOCMODE_OPEN_FORMENABLEDEDIT = 1;//Form mode with editing enabled
DOCMODE_OPEN_DESIGN = 2;//Design Mode

/*DIAGNOSTICFAILFLAGS*/
DGNFAIL_SHOWMESSAGE = 1;//Note Number
DGNFAIL_NOPRINTING = 2;//Note Number

/*DIAGNOSTICRUNFLAGS*/
DGNRUN_PRECALC = 1;//Note Number
DGNRUN_PREPRINT = 4;//Note Number
DGNRUN_ONLOADING = 16;//Note Number
DGNRUN_ONSAVING = 8;//Note Number
DGNRUN_INFORMLIST = 32;//Note Number
DGNRUN_INDESIGNLIST = 64;//Note Number

/*DOCMODETYPE*/
DOCMODE_FORM = 0;//Form Mode
DOCMODE_ENABLEEDITING = 1;//Enable Editing Mode
DOCMODE_DESIGN = 2;//Design Mode

/*DOCUMENTOPENFLAGS*/
DOCOPEN_FLAGS_NONE = 0;//Document Open Flags: no flag at all
DOCOPEN_FLAGS_READONLY = 1;//Document Open Flags: force read-only
DOCOPEN_FLAGS_NOOPENINGSCRIPT = 2;//Document Open Flags: bypass on-opening system script
DOCOPEN_FLAGS_NOSYSTEMSCRIPTS = 4;//Document Open Falgs: disable all system scripts

/*DOCUMENTTYPE*/
DOCUMENT_CLIENT	= 0;//Client Associated Document
DOCUMENT_LIBRARY = 1;//Library Document
DOCUMENT_KNOWLEDGELIBRARY = 2;//Knowledge Library Document
DOCUMENT_STANDALONE = 3;//Standalone Document

/*DOLLARPOSTION*/
CURRENCYPOSITION_LEFT = 1;//Cell Function
CURRENCYPOSITION_RIGHT = 2;//Cell Function

/*DOLLARSYMBOL*/
CURRENCYSYMBOL_DEFAULT = 0;//Cell Function

/*EXPORTTYPE*/
EXPTYPE_XBRL = 0;//XBRL Export
EXPTYPE_INLINEXBRL = 1;//Inline XBRL Export

/*EXTDATACOPY*/
EXTDATACOPY_NEVER = 0;//Note Number
EXTDATACOPY_ADD	Note = 1;//Note Number
EXTDATACOPY_REPLACE	Note = 2;//Note Number
EXTDATACOPY_REMOVE	Note = 3;//Note Number
EXTDATACOPY_REPLACEOLD = 4;//Note Number

/*EXTDATAPOS*/
EXTDATAPOS_CURRENT = 1;//Note Number
EXTDATAPOS_RECNO = 2;//Note Number
EXTDATAPOS_FILTER = 3;//Note Number

/*EXTDATATYPE*/
EXTDATATYPE_INTERNAL = 0;//Non external data
EXTDATATYPE_CELLDATA = 1;//CaseView DB Global cell data
EXTDATATYPE_FORMCELLDATA = 2;//CaseView DB Form cell data
EXTDATATYPE_CELLDATASYS = 3;//System DB Global cell data
EXTDATATYPE_FORMCELLDATASYS = 4;//System DB Form cell data
EXTDATATYPE_CELLDATA_L = 5;//CaseView DB Global cell data with long IDs
EXTDATATYPE_FORMCELLDATA_L = 6;//CaseView DB Form cell data with long IDs

/*FEATURESTYPE*/
FEAT_OLDLIBRARIES = 0;//Old Libraries
FEAT_KNOWLEDGELIBS = 1;//Knowledge Libraries
FEAT_AUTOCOPY = 2;//Auto Copy
FEAT_USERDEFINEDDATA = 3;//User Defined Data
FEAT_EDIEXPORT = 4;//EDI INFENT Export
FEAT_XBRLEXPORT = 5;//XBRL Export
FEAT_PAGINATION = 6;//Pagination
FEAT_HYPHENATION = 7;//Hyphenation
FEAT_BROWSES = 8;//Browses
FEAT_GLOBALFONTCHANGE = 9;//Global Font Change
FEAT_CARRYFORWARD = 10;//Carry Forward
FEAT_CELLGROUPS = 11;//Cell Groups
FEAT_SCRIPTING = 12;//Scripting
FEAT_HISTORY = 13;//History and Milestones
FEAT_ISSUES = 14;//Issues
FEAT_CALCULATIONTOOL = 15;//Calculation Tool

/*FILEAUTOSAVETYPES*/
DOCFILE_AUTOSAVE_PROMPT = 0;//Prompt (Default)
DOCFILE_AUTOSAVE_AUTOMATIC = 1;//Automatically Save
DOCFILE_AUTOSAVE_NEVER = 2;//Never Save

/*FILEBACKUPTYPES*/
DOCFILE_BACKUP_KEEP = 0;//Always keep a backup copy
DOCFILE_BACKUP_REMOVE = 1;//Don't keep a backup copy
DOCFILE_BACKUP_CLIENTDEFAULT = 2;//Use client defaults

/*FONTWIDTH*/
FW_DONTCARE = 0;//Note Number
FW_THIN	Note = 100;//Note Number
FW_EXTRALIGHT = 200;//Note Number
FW_ULTRALIGHT = 200;//Note Number
FW_LIGHT = 300;//Note Number
FW_NORMAL = 400;//Note Number
FW_REGULAR = 400;//Note Number
FW_MEDIUM = 500;//Note Number
FW_SEMIBOLD = 600;//Note Number
FW_DEMIBOLD = 600;//Note Number
FW_BOLD	Note = 700;//Note Number
FW_EXTRABOLD = 800;//Note Number
FW_ULTRABOLD = 800;//Note Number
FW_HEAVY = 900;//Note Number
FW_BLACK = 900;//Note Number

/*FREEZE_EDITING_OPTIONS*/
NO_INPUT_PERMITTED = 0;//No input permitted
ENABLE_INPUT_DISPLAY_SELECTION = 1;//Enable input and display selection
ENABLE_INPUT_HIDE_SELECTION = 2;//Enable input but hide selection

/*HEIGHTTYPE*/
HEIGHT_VARIABLE = 0;//Height Variable (Auto)
CELL_HEIGHT_VARIABLE = 0;//Height Variable (Auto)
HEIGHT_EXACT = 1;//Height Exactly
CELL_HEIGHT_EXACT = 1;//Height Exactly
HEIGHT_ATLEAST = 2;//Height At Least
CELL_HEIGHT_ATLEAST = 2;//Height At Least
HEIGHT_ATMOST = 3;//Height At Most
CELL_HEIGHT_ATMOST = 3;//Height At Most

/*HFATTACHTYPE*/
HFATTACHTYPE_NOHF = 0;//Header/Footer Attach Types
HFATTACHTYPE_ABSOLUTE = 1;//Header/Footer Attach Types
HFATTACHTYPE_CONDITIONAL = 2;//Header/Footer Attach Types

/*HFSETTYPE*/
//HFSETTYPE_NONE = only;//Header/Footer Conditional Set Types
HFSETTYPE_LANDSCAPE = 0;//Header/Footer Conditional Set Types
HFSETTYPE_PORTRAIT = 1;//Header/Footer Conditional Set Types
HFSETTYPE_ODD = 2;//Header/Footer Conditional Set Types
HFSETTYPE_EVEN = 3;//Header/Footer Conditional Set Types
HFSETTYPE_FIRST_PAGE = 4;//Header/Footer Conditional Set Types
HFSETTYPE_SUBSEQUENT = 5;//Header/Footer Conditional Set Types
HFSETTYPE_LAST = 6;//Header/Footer Conditional Set Types

/*HIDEACTION*/
ACTION_NEVER_HIDE = 0;//Note Number
ACTION_HIDE_IF_BUTTON = 1;//Note Number
ACTION_ALWAYS_HIDE = 2;//Note Number
ACTION_HIDE_IF_COND = 3;//Note Number
ACTION_NOAPPLY = 3;//Note Number

/*HORZALIGN*/
HORZ_DEFAULT = 0;//Note Number
HORZ_LEFT = 1;//Note Number
HORZ_CENTER = 2;//Note Number
HORZ_RIGHT = 3;//Note Number
HORZ_JUSTIFY = 4;//Note Number
HORZ_FILL = 5;//Note Number

/*IGNORESIGN*/
IGNORESIGN_DEFAULT = 0;//Note Number
IGNORESIGN_YES = 1;//Note Number
IGNORESIGN_NO = 2;//Note Number

/*IMAGEFORMATS*/
IMAGE_FORMAT_BMP = 0;//Image Member
IMAGE_FORMAT_JPG = 1;//Image Member
IMAGE_FORMAT_JFF = 2;//Image Member
IMAGE_FORMAT_JTF = 3;//Image Member
IMAGE_FORMAT_PNG = 4;//Image Member
IMAGE_FORMAT_TIF = 5;//Image Member
IMAGE_FORMAT_CMP = 6;//Image Member
IMAGE_FORMAT_MSP = 7;//Image Member

/*INPUTDIALOG*/
INPUTDIALOG_TEXT = 0;//A dialog where the user can enter text
INPUTDIALOG_INTEGER = 1;//A dialog where the user uses up/down arrows to choose a number
INPUTDIALOG_LENGTH = 2;//A dialog where the user uses up/down arrows to choose a number (In Inches)
INPUTDIALOG_POINTS = 3;//A dialog where the user uses up/down arrows to choose a number (In Points)

/*INPUT_FORMAT_TYPE*/
INPUT_FORMAT_NONE = 0;//No input format
INPUT_FORMAT_POSTALCODE = 1;//ANA NAN format
INPUT_FORMAT_ZIPCODE_USA = 2;//*NNNNN format
INPUT_FORMAT_ZIPCODE4_USA = 3;//*N-NNNN format
INPUT_FORMAT_ZIPCODE_NUM = 4;//*N format
INPUT_FORMAT_ZIPCODE_HYPHEN = 5;//*N-*N format
INPUT_FORMAT_ZIPCODE_GENERAL = 6;//*N or *N-*N format
INPUT_FORMAT_PHONENO_CAN = 7;//*N-NNN-NNNN format
INPUT_FORMAT_PHONENO_USA = 8;//(*N) NNN-NNNN format
INPUT_FORMAT_PHONENO_GENERAL = 9;//*N-*N-*N format
INPUT_FORMAT_PHONENO_INTL = 10;//+ *N *N *N *N format
INPUT_FORMAT_SIN_CAN = 11;//*N NNN NNN format
INPUT_FORMAT_SIN_USA = 12;//*N-NN-NNNN format
INPUT_FORMAT_VALUE_X = 13;//X or' ' values
INPUT_FORMAT_VALUE_YN = 14;//Y,N or values
INPUT_FORMAT_SPECIFIC_LIST = 15;//Specific list values
INPUT_FORMAT_CUSTOM = 16;//Custom format

/*INPUT_RESPONSE*/
INPUT_RESPONSE_NONE = 0;//No input response
INPUT_RESPONSE_TOOLTIP = 1;//Tooltip input response
INPUT_RESPONSE_INFODLG = 2;//Info dialog input response
INPUT_RESPONSE_OKCANDLG = 3;//OK/Cancel dialog input response
INPUT_RESPONSE_YESNODLG = 4;//Yes/No dialog input response

/*INPUT_VALIDATION*/
INPUT_VALIDATE_NONE = 0;//No input validation
INPUT_VALIDATE_LESS = 1;//Less than input validation
INPUT_VALIDATE_LESSEQ = 2;//Less than or equal input validation
INPUT_VALIDATE_GT = 3;//Greater than input validation
INPUT_VALIDATE_GTEQ = 4;//Greater than or equal input validation
INPUT_VALIDATE_NOTEQ = 5;//Not equal input validation
INPUT_VALIDATE_RANGE = 6;//In range input validation
INPUT_VALIDATE_NOTEMPTY = 7;//Not empty input validation
INPUT_VALIDATE_MATCHFMT = 8;//Match format input validation

/*INSTABLECOLFLAG*/
INSTABLECOLFLAG_COPYCONTENTS = 1;//Border Types
INSTABLECOLFLAG_BEFORE = 8;//Border Types
INSTABLECOLFLAG_MODIFIABLE = 16;//Border Types

/*KLICUSTOMPROPERTYTYPES*/
KLI_CUSTOM_PROPERTY_FIRSTTYPE = 0;//KLI Custom Property Member
KLI_CUSTOM_PROPERTY_CHARACTER = 0;//Text
KLI_CUSTOM_PROPERTY_BOOLEAN = 1;//Boolean
KLI_CUSTOM_PROPERTY_INTEGER = 2;//Integer
KLI_CUSTOM_PROPERTY_SINGLE = 3;//Numeric
KLI_CUSTOM_PROPERTY_DOUBLE = 4;//Large Numeric
KLI_CUSTOM_PROPERTY_CURRENCY = 5;//Monetary
KLI_CUSTOM_PROPERTY_DATE = 6;//Date
KLI_CUSTOM_PROPERTY_LISTTEXT = 7;//Drop Down List
KLI_CUSTOM_PROPERTY_LASTTYPE = 7;//KLI Custom Property Member

/*KLINSERTFLAGS*/
KLL_INSERT_FLAGS_NONE = 0;//KLL Insert no flags
KLL_INSERT_RENUM_NONTABLE_CELLS = 1;//Renumber non table cells

/*KLIRECORDSETMEMFLDINDICE*/
KLI_RECORDSET_MEM_FLD_CHECK	KLI = 0;//Form Recordset Member
KLI_RECORDSET_MEM_FLD_INCLUDE	KLI = 1;//Form Recordset Member
KLI_RECORDSET_MEM_FLD_MERGE	KLI Form = 2;//Recordset Member

/*KLISAVERETURNCODE*/
KLISYNC_SAVE_SUCCEED = 1;//KLI Member
KLISYNC_SAVE_FAILED = 0;//KLI Member
KLISYNC_ERROR_UNKNOWN = -1;//KLI Member
KLISYNC_NOT_KLIFILE = -2;//KLI Member
KLISYNC_KLI_NOTFOUND = -3;//KLI Member
KLISYNC_INVALID_USER = -4;//KLI Member
KLISYNC_LOGIN_CANCELED = -5;//KLI Member

/*LABELCELLTYPE*/
LABEL_UNRESTRICTED = 0;//Note Number
LABEL_POSTAL = 1;//Note Number
LABEL_YESNO = 2;//Note Number
LABEL_XONLY = 3;//Note Number
LABEL_ZIP = 4;//Note Number

/*LEVELFORMAT*/
LVF_NONE = -1;//No level format
LVF_ARABIC = 0;//Arabic level format
LVF_ROMAN_LARGE = 1;//Large Roman level format
LVF_ROMAN_SMALL = 2;//Small Roman level format
LVF_LETTERS_LARGE = 3;//Large letters level format
LVF_LETTERS_SMALL = 4;//Small letters level format
LVF_SIMPLE_CHINESE_1 = 5;//Simple Chinese level format (simple number)
LVF_SIMPLE_CHINESE_2 = 6;//Simple Chinese level format (traditional number)
LVF_SIMPLE_CHINESE_3 = 7;//Simple Chinese level format (traditional year)
LVF_SIMPLE_CHINESE_4 = 8;//Simple Chinese level format (traditional animal)
LVF_TRADITION_CHINESE_1 = 9;//Traditional Chinese level format (simple number)
LVF_TRADITION_CHINESE_2 = 10;//Traditional Chinese level format (traditional number)
LVF_TRADITION_CHINESE_3 = 11;//Traditional Chinese level format (traditional year)
LVF_TRADITION_CHINESE_4 = 12;//Traditional Chinese level format (traditional animal)
LVF_GREEK_LETTERS_LARGE = 13;//Greek large letters level format
LVF_GREEK_LETTERS_SMALL = 14;//Greek small letters level format

/*LINESYLEEXTENTTYPES*/
LINE_EXTENT_INDENTATION = 0;//Line Style Member
LINE_EXTENT_FIXED = 1;//Line Style Member
LINE_EXTENT_MARGIN = 2;//Line Style Member

/*LINKSETSORTTYPES*/
LINKSET_SECTION_GROUP = 0;//Link Set
LINKSET_SECTION = 1;//Link Set
LINKSET_TABLE_SORT_GROUP = 2;//Link Set

/*MESSAGEBOXRESULT*/
MESSAGE_RESULT_OK = 1;//Message box result clicked Ok
MESSAGE_RESULT_CANCEL = 2;//Message box result clicked Cancel
MESSAGE_RESULT_YES = 6;//Message box result clicked Yes
MESSAGE_RESULT_NO = 7;//Message box result clicked No
MESSAGE_RESULT_YESALL = 20;//Message box result clicked Yes to All
MESSAGE_RESULT_NOALL = 21;//Message box result clicked No to All

/*MESSAGEBOXTYPE*/
MESSAGE_OK = 0;//Message box with Ok Button
MESSAGE_OKCANCEL = 1;//Message box with Ok and Cancel Button
MESSAGE_YESNO = 2;//Message box with Yes/No Buttons
MESSAGE_YESNOCANCEL = 3;//Message box with Yes/No and Cancel Button
MESSAGE_YESNOALLCANCEL = 4;//Message box with Yes/YestoAll/No/NotoAll and Cancel Button

/*MISSING_REFERENCE_DRAW_OP*/
NORMALLY = 0;//Draw reference annotation normally
HIDDEN = 1;//Hide reference annotation
CUSTOM_COLOR = 2;//Draw reference annotation in different color

/*MODEACCESSIBILITY*/
DOCMODE_ALWAYS_VIEWLOCKED = -1;//Always accessibility (View Locked)
DOCMODE_NOT_ACCESSIBLE = 0;//Document Member
DOCMODE_ALWAYS_ACCESSIBLE = 1;//Document Member
DOCMODE_ACCESS_LEVELONE = 2;//Document Member
DOCMODE_ACCESS_LEVELTWO = 3;//Document Member
DOCMODE_ACCESS_LEVELTHREE = 4;//Document Member
DOCMODE_ACCESS_LEVELFOUR = 5;//Document Member
DOCMODE_ACCESS_LEVELFIVE = 6;//Document Member
DOCMODE_ACCESS_LEVELSIX = 7;//Document Member
DOCMODE_ACCESS_LEVELSEVEN = 8;//Document Member
DOCMODE_ACCESS_LEVELEIGHT = 9;//Document Member
DOCMODE_ACCESS_LEVELNINE = 10;//Document Member
DOCMODE_ACCESS_LEVELTEN = 11;//Document Member

/*NEGATIVEASTYPE*/
NEGATIVEAS_BRACKET = 0;//Note Number
NEGATIVEAS_NEG = 1;//Note Number
NEGATIVEAS_CR = 2;//Note Number
NEGATIVEAS_RBRACKET = 4;//Note Number
NEGATIVEAS_POS = 5;//Note Number
NEGATIVEAS_BLANK = 6;//Note Number

/*NET_INCOME_GROUPS*/
NET_INCOME_MAP	Map	0;//
NET_INCOME_GROUP1 = 1;//Group 1
NET_INCOME_GROUP2 = 2;//Group 2
NET_INCOME_GROUP3 = 3;//Group 3
NET_INCOME_GROUP4 = 4;//Group 4
NET_INCOME_GROUP5 = 5;//Group 5
NET_INCOME_GROUP6 = 6;//Group 6
NET_INCOME_GROUP7 = 7;//Group 7
NET_INCOME_GROUP8 = 8;//Group 8
NET_INCOME_GROUP9 = 9;//Group 9
NET_INCOME_GROUP10 = 10;//Group 10

/*NOTENUMFORMAT*/
NNF_NONE = -1;//No note number format
NNF_ARABIC = 0;//Arabic note number format
NNF_ROMAN_LARGE = 1;//Large Roman note number format
NNF_ROMAN_SMALL = 2;//Small Roman note number format
NNF_LETTERS_LARGE = 3;//Large letters note number format
NNF_LETTERS_SMALL = 4;//Small letters note number format
NNF_SIMPLE_CHINESE_1 = 5;//Simple Chinese note number format (simple number)
NNF_SIMPLE_CHINESE_2 = 6;//Simple Chinese note number format (traditional number)
NNF_SIMPLE_CHINESE_3 = 7;//Simple Chinese note number format (traditional year)
NNF_SIMPLE_CHINESE_4 = 8;//Simple Chinese note number format (traditional animal)
NNF_TRADITION_CHINESE_1 = 9;//Traditional Chinese note number format (simple number)
NNF_TRADITION_CHINESE_2 = 10;//Traditional Chinese note number format (traditional number)
NNF_TRADITION_CHINESE_3 = 11;//Traditional Chinese note number format (traditional year)
NNF_TRADITION_CHINESE_4 = 12;//Traditional Chinese note number format (traditional animal)
NNF_GREEK_LETTERS_LARGE = 13;//Greek large letters note number format
NNF_GREEK_LETTERS_SMALL = 14;//Grrek small letters note number format

/*NOTENUMSUBTYPE*/
NNT_NONE = -1;//No sub type
NNT_DEPTH_ONE_SHOW_ALL = 0;//Level 1, Show All
NNT_DEPTH_TWO_SHOW_ALL = 1;//Level 2, Show All
NNT_DEPTH_THREE_SHOW_ALL = 2;//Level 3, Show All
NNT_DEPTH_TWO_SHOW_ALLFLUSH = 3;//Level 2, Show All, Flush
NNT_DEPTH_TWO_SHOW_ALLBRACKET = 4;//Level 2, Show All, Bracketed 2 Only
NNT_DEPTH_TWO_SHOW_TWO = 5;//Level 2, Show 2 Only
NNT_DEPTH_THREE_SHOW_THREE = 6;//Level 3, Show 3 Only
NNT_DEPTH_ONE_SHOW_ALLBRACKET = 7;//Level 1, Show 1 Only, Bracketed 1 Only
NNT_DEPTH_TWO_SHOW_TWOBRACKET = 8;//Level 2, Show 2 Only, Bracketed 2 Only
NNT_DEPTH_THREE_SHOW_THREEBRACKET = 9;//Level 3, Show 3 Only, Bracketed 3 Only

/*NOTENUMTYPE*/
NOTE_NUMBER = 0;//Note number
NOTE_NUMBER_XREF = 1;//Note number cross reference

/*NUMBERTYPES*/
NUMBER_ARABIC = 0;//Paragraph/Bullet Member
NUMBER_ROMAN_LARGE = 1;//Paragraph/Bullet Member
NUMBER_ROMAN_SMALL = 2;//Paragraph/Bullet Member
NUMBER_LETTER_LARGE = 3;//Paragraph/Bullet Member
NUMBER_LETTER_SMALL = 4;//Paragraph/Bullet Member
NUMBER_SIMPLIFIED_CHINESE_1 = 5;//Paragraph/Bullet Member
NUMBER_SIMPLIFIED_CHINESE_2 = 6;//Paragraph/Bullet Member
NUMBER_SIMPLIFIED_CHINESE_3 = 7;//Paragraph/Bullet Member
NUMBER_SIMPLIFIED_CHINESE_4 = 8;//Paragraph/Bullet Member
NUMBER_TRADITIONAL_CHINESE_1 = 9;//Paragraph/Bullet Member
NUMBER_TRADITIONAL_CHINESE_2 = 10;//Paragraph/Bullet Member
NUMBER_TRADITIONAL_CHINESE_3 = 11;//Paragraph/Bullet Member
NUMBER_TRADITIONAL_CHINESE_4 = 12;//Paragraph/Bullet Member

/*OBJECTTYPEID*/
ID_CELL = 1;//Cell Object
ID_PAGE_NUMBER = 3;//Page Number
ID_LINE = 4;//Line Style
ID_NOTE_NUMBER = 5;//Note Number
ID_PAGE_XREF = 6;//Page Number Cross reference
ID_PICTURE = 8;//Picture
ID_HARD_HYPHEN = 10;//Hard Hyphen
ID_GRAPH = 11;//Graph
ID_PAGEBREAK = 12;//Page Break
ID_OLE = 13;//OLE Object
ID_ANNOTATION = 15;//Annotation
TYPEID_ANYOBJECT = 0;//OLE Object
TYPEID_CELL = 1;//OLE Object
TYPEID_PAGE = 3;//OLE Object
TYPEID_LINE = 4;//OLE Object
TYPEID_NOTENUM = 5;//OLE Object
TYPEID_PAGEXREF = 6;//OLE Object
TYPEID_PICTURE = 8;//OLE Object
TYPEID_HARDHYPHEN = 10;//OLE Object
TYPEID_GRAPH = 11;//OLE Object
TYPEID_PAGEBREAK = 12;//OLE Object
TYPEID_OLE = 13;//OLE Object
TYPEID_ANNOTATION = 15;//Annotation
TYPEID_PDF_EMBED = 16;//PDF Embed
TYPEID_EXTERNAL_DOC = 16;//External Documnt Embed
TYPEID_HYPERLINKS = 17;//Hyperlinks Object

/*OLDCVCOLORS*/
COLOR_BLACK = 0x00000000;//Black
COLOR_LTGRAY = 0x00c0c0c0;//Light Gray
COLOR_MEDGRAY = 0x00a0a0a0;//Medium Gray
COLOR_DKGRAY = 0x00808080;//Dark Gray
COLOR_RED = 0x000000ff;//Red
COLOR_BLUE = 0x00ff3030;//Blue
COLOR_LIME = 0x0000ff00;//Lime
COLOR_MAROON = 0x00000084;//Maroon
COLOR_GREEN = 0x00008200;//Green
COLOR_OLIVE = 0x00008284;//Olive
COLOR_NAVY = 0x00840000;//Navy
COLOR_PURPLE = 0x00840084;//Purple
COLOR_TEAL = 0x00848200;//Teal
COLOR_YELLOW = 0x0000ffff;//Yellow
COLOR_FUSCHIA = 0x00ff00ff;//Fuschia
COLOR_AQUA = 0x00ffff00;//Aqua
COLOR_WHITE = 0x00ffffff;//White

/*OPENDOC_ACCESS*/
OPENDOC_READWRITE = 0;//Application Member
OPENDOC_READONLY = 1;//Application Member
OPENDOC_READWRITEONLY = 2;//Application Member

/*PAGEBREAKTYPES*/
PAGEBREAK_NEXTPAGE = 1;//Page Break Member
PAGEBREAK_NEXTEVENPAGE = 2;//Page Break Member
PAGEBREAK_NEXTODDPAGE = 3;//Page Break Member
PAGEBREAK_THISORNEXTEVENPAGE = 4;//Page Break Member
PAGEBREAK_THISORNEXTODDPAGE = 5;//Page Break Member

/*PAGENUMFORMATTYPES*/
PAGENUM_FORMAT_ARABIC_NUM = 0;//Page Number Member
PAGENUM_FORMAT_LARGE_ROMAN = 1;//Page Number Member
PAGENUM_FORMAT_SMALL_ROMAN = 2;//Page Number Member
PAGENUM_FORMAT_LARGE_LETTER = 3;//Page Number Member
PAGENUM_FORMAT_SMALL_LETTER = 4;//Page Number Member

/*PAGENUMTOUSETYPES*/
PAGENUM_USE_CURRENT_PAGE = 0;//Page Number Member
PAGENUM_USE_NEXT_PAGE = 1;//Page Number Member
PAGENUM_USE_PREVIOUS_PAGE = 2;//Page Number Member

/*PAGESETUPNUMSCHEME*/
PGSETUP_DEFAULT = 0;//Default page numbering
PGSETUP_SETPAGEGNUMTO = 1;//Set page number to
PGSETUP_INCPAGENUMBY = 2;//Increase page number by

/*PAPERSIZES*/
PAPER_LETTER = 1;//Note Number
PAPER_LETTERSMALL = 2;//Note Number
PAPER_TABLOID = 3;//Note Number
PAPER_LEDGER = 4;//Note Number
PAPER_LEGAL = 5;//Note Number
PAPER_STATEMENT = 6;//Note Number
PAPER_EXECUTIVE = 7;//Note Number
PAPER_A3 = 8;//Note Number
PAPER_A4 = 9;//Note Number
PAPER_A4SMALL = 10;//Note Number
PAPER_A5 = 11;//Note Number
PAPER_B4 = 12;//Note Number
PAPER_B5 = 13;//Note Number
PAPER_FOLIO = 14;//Note Number
PAPER_QUARTO = 15;//Note Number
PAPER_10X14 = 16;//Note Number
PAPER_11X17 = 17;//Note Number
PAPER_NOTE = 18;//Note Number
PAPER_ENV_9 = 19;//Note Number
PAPER_ENV_10 = 20;//Note Number
PAPER_ENV_11 = 21;//Note Number
PAPER_ENV_12 = 22;//Note Number
PAPER_ENV_14 = 23;//Note Number
PAPER_CSHEET = 24;//Note Number
PAPER_DSHEET = 25;//Note Number
PAPER_ESHEET = 26;//Note Number
PAPER_ENV_DL = 27;//Note Number
PAPER_ENV_C5 = 28;//Note Number
PAPER_ENV_C3 = 29;//Note Number
PAPER_ENV_C4 = 30;//Note Number
PAPER_ENV_C6 = 31;//Note Number
PAPER_ENV_C65 = 32;//Note Number
PAPER_ENV_B4 = 33;//Note Number
PAPER_ENV_B5 = 34;//Note Number
PAPER_ENV_B6 = 35;//Note Number
PAPER_ENV_ITALY = 36;//Note Number
PAPER_ENV_MONARCH = 37;//Note Number
PAPER_ENV_PERSONAL = 38;//Note Number
PAPER_FANFOLD_US = 39;//Note Number
PAPER_FANFOLD_STD_GERMAN = 40;//Note Number
PAPER_FANFOLD_LGL_GERMAN = 41;//Note Number

/*PARAGRAPHALIGN*/
ALIGN_NONE = 0;//Note Number
ALIGN_AT = 1;//Note Number
ALIGN_WITH = 2;//Note Number

/*PARAGRAPHJUSTIFICATION*/
JUST_LEFT = 0;//Left
JUST_FLUSH = 1;//Flush
JUST_CENTER = 2;//Center
JUST_CENTRE = 2;//Center
JUST_RIGHT = 3;//Right

/*PARAHEIGHTTYPE*/
PARA_HEIGHT_AUTO = 0;//Paragraph Number
PARA_HEIGHT_ATLEAST = 1;//Paragraph Number
PARA_HEIGHT_ATMOST = 2;//Paragraph Number
PARA_HEIGHT_EXACT = 3;//Paragraph Number

/*PARASPACINGNEW*/
SPACING_SINGLE = 1;//Single
SPACING_ONEANDHALF = 2;//1 1/2
SPACING_DOUBLE = 3;//Double
SPACING_ATLEAST = 4;//At least
SPACING_ATMOST = 5;//At most
SPACING_EXACTLY = 6;//Exactly
SPACING_MULTIPLE = 7;//Multiple

/*PCHOOSERNODETYPES*/
PCHOOSER_NODE_CHECKBOX = 0;//
PCHOOSER_NODE_FOLDER = 1;//
PCHOOSER_NODE_RADIOBUTTON = 2;//

/*PDFPRINT_CODE*/
PDFPRINT_CODE_FNAME = 0;//With file name
PDFPRINT_CODE_DEFFNAME = 1;//With default file name
PDFPRINT_CODE_DLGFNAME = 2;//With file dialog

/*PDF_SAVEFLAGS*/
PDF_SAVEFLAG_FULLEMBEDFONTS = 1;//Full embedded fonts
PDF_SAVEFLAG_PDFA = 2;//PDF/A
PDF_SAVEFLAG_OPENPDFFILE = 4;//Open generated PDF file
PDF_SAVEFLAG_MASTERDOC = 8;//Use master document information if available

/*PERIODRANGETYPE*/
PERIODRANGE_FULL = 0;//Full Period
PERIODRANGE_START = 1;//Period Start
PERIODRANGE_END = 2;//Period End

/*PERIODTYPE*/
PERIOD_YEARLY = 0;//Yearly Period
PERIOD_QUARTER_FIRST = 1;//Quarterly Period - First Quarter
PERIOD_QUARTER_SECOND = 2;//Quarterly Period - Second Quarter
PERIOD_QUARTER_THIRD = 3;//Quarterly Period - Fourth Quarter
PERIOD_QUARTER_FOURTH = 4;//
PERIOD_SEMI_FIRST = 5;//Semi-Annual Period - First Half
PERIOD_SEMI_SECOND = 6;//Semi-Annual Period - Second Quarter
PERIOD_THIRD_FIRST = 7;//Thirdly Period - First Third
PERIOD_THIRD_SECOND = 8;//Thirdly Period - Second Third
PERIOD_THIRD_THIRD = 9;//Thirdly Period - Third Third
PERIOD_MONTHLY_FIRST = 10;//Monthly Period - First Month
PERIOD_MONTHLY_SECOND = 11;//Monthly Period - Second Month
PERIOD_MONTHLY_THIRD = 12;//Monthly Period - Third Month
PERIOD_MONTHLY_FOURTH = 13;//Monthly Period - Fourth Month
PERIOD_MONTHLY_FIFTH = 14;//Monthly Period - Fifth Month
PERIOD_MONTHLY_SIXTH = 15;//Monthly Period - Sixth Month
PERIOD_MONTHLY_SEVENTH = 16;//Monthly Period - Seventh Month
PERIOD_MONTHLY_EIGHTH = 17;//Monthly Period - Eighth Month
PERIOD_MONTHLY_NINTH = 18;//Monthly Period - Ninth Month
PERIOD_MONTHLY_TENTH = 19;//Monthly Period - Tenth Month
PERIOD_MONTHLY_ELEVENTH = 20;//Monthly Period - Eleventh Month
PERIOD_MONTHLY_TWELFTH = 21;//Monthly Period - Twelfth Month
PERIOD_BIMONTHLY_FIRST_SECOND = 22;//Bi-Monthly Period - First/Second Months
PERIOD_BIMONTHLY_THIRD_FOURTH = 23;//Bi-Monthly Period - Third/Fourth Months
PERIOD_BIMONTHLY_FIFTH_SIXTH = 24;//Bi-Monthly Period - Fifth/Sixth Months
PERIOD_BIMONTHLY_SEVENTH_EIGHTH = 25;//Bi-Monthly Period - Seventh/Eighth Months
PERIOD_BIMONTHLY_NINTH_TENTH = 26;//Bi-Monthly Period - Ninth/Tenth Months
PERIOD_BIMONTHLY_ELEVENTH_TWELFTH = 27;//Bi-Monthly Period - Eleventh/Twelfth Months
PERIOD_13_PERIOD1 = 28;//13 Periods - Period 1
PERIOD_13_PERIOD2 = 29;//13 Periods - Period 2
PERIOD_13_PERIOD3 = 30;//13 Periods - Period 3
PERIOD_13_PERIOD4 = 31;//13 Periods - Period 4
PERIOD_13_PERIOD5 = 32;//13 Periods - Period 5
PERIOD_13_PERIOD6 = 33;//13 Periods - Period 6
PERIOD_13_PERIOD7 = 34;//13 Periods - Period 7
PERIOD_13_PERIOD8 = 35;//13 Periods - Period 8
PERIOD_13_PERIOD9 = 36;//13 Periods - Period 9
PERIOD_13_PERIOD10 = 37;//13 Periods - Period 10
PERIOD_13_PERIOD11 = 38;//13 Periods - Period 11
PERIOD_13_PERIOD12 = 39;//13 Periods - Period 12
PERIOD_13_PERIOD13 = 40;//13 Periods - Period 13
PERIOD_RANDOM_PERIOD1 = 41;//Random Period - Period 1
PERIOD_RANDOM_PERIOD2 = 42;//Random Period - Period 2
PERIOD_RANDOM_PERIOD3 = 43;//Random Period - Period 3
PERIOD_RANDOM_PERIOD4 = 44;//Random Period - Period 4
PERIOD_RANDOM_PERIOD5 = 45;//Random Period - Period 5
PERIOD_RANDOM_PERIOD6 = 46;//Random Period - Period 6
PERIOD_RANDOM_PERIOD7 = 47;//Random Period - Period 7
PERIOD_RANDOM_PERIOD8 = 48;//Random Period - Period 8
PERIOD_RANDOM_PERIOD9 = 49;//Random Period - Period 9
PERIOD_RANDOM_PERIOD10 = 50;//Random Period - Period 10
PERIOD_RANDOM_PERIOD11 = 51;//Random Period - Period 11
PERIOD_RANDOM_PERIOD12 = 52;//Random Period - Period 12
PERIOD_RANDOM_PERIOD13 = 53;//Random Period - Period 13
PERIOD_ACTIVE = 54;//Active Period
PERIOD_PRIOR_ACTIVE = 55;//Prior Active Period
PERIOD_FULL_DATE = 56;//Fully Specified Date
PERIOD_FULL_PERIOD = 57;//Fully Specified Period
PERIOD_FOREVER = 58;//Forever

/*PERMISSIONINDEX*/
PERM_MODIFY_CELL_CALCS = 0;//Security Configuration Member
PERM_MODIFY_CELL_LOGIC = 0;//Security Configuration Member
PERM_MODIFY_CELL_FORMAT = 1;//Security Configuration Member
PERM_MODIFY_CELL_ATTRIBUTES = 1;//Security Configuration Member
PERM_DELETE_CELL = 2;//Security Configuration Member
PERM_MODIFY_PARA_FORMAT = 3;//Security Configuration Member
PERM_MODIFY_PARA_CONTENTS = 4;//Security Configuration Member
PERM_MODIFY_PARA_SKIPS = 5;//Security Configuration Member
PERM_MODIFY_SECTION_SKIP = 6;//Security Configuration Member
PERM_MODIFY_SECTION_FORMAT = 7;//Security Configuration Member
PERM_MODIFY_ROUNDING = 8;//Security Configuration Member
PERM_MODIFY_TABLE_FORMAT = 9;//Security Configuration Member
PERM_MODIFY_TABLE_LOGIC = 10;//Security Configuration Member
PERM_INSERT_COLUMN = 11;//Security Configuration Member
PERM_DELETE_COLUMN = 12;//Security Configuration Member
PERM_MODIFY_COLUMN_FORMAT = 13;//Security Configuration Member
PERM_MODIFY_COLUMN_LOGIC = 14;//Security Configuration Member
PERM_INSERT_ROW = 15;//Security Configuration Member
PERM_DELETE_ROW = 16;//Security Configuration Member
PERM_MODIFY_ROW_FORMAT = 17;//Security Configuration Member
PERM_MODIFY_ROW_LOGIC = 18;//Security Configuration Member
PERM_MODIFY_CHARLEVERL_FORMAT = 19;//Security Configuration Member
PERM_APPLY_STYLE = 20;//Security Configuration Member
PERM_INSERT_KNOWLEDGELIBRARY = 21;//Security Configuration Member
PERM_INSERT_AUTOMATICTOC = 22;//Security Configuration Member
PERM_MODIFY_PAGE_BREAK = 23;//Security Configuration Member
PERM_MODIFY_CELLOPTIONS = 24;//Security Configuration Member
PERM_CONFIGURE_DIAGNOSTICS = 25;//Security Configuration Member
PERM_UPDATE_KNOWLEDGELIBRARY = 26;//Security Configuration Member
PERM_DELETE_KNOWLEDGELIBRARY = 27;//Security Configuration Member
PERM_INSERT_SUBDOCUMENT = 28;//Security Configuration Member
PERM_ALLOW_FULL_COPY = 29;//Security Configuration Member
PERM_MODIFY_PAGE_SETUP = 30;//Security Configuration Member
PERM_MODIFY_WIDE_TABLE = 31;//Security Configuration Member
PERM_RESIZE_SECTIONS = 32;//Security Configuration Member

/*PERMISSIONINDEXRANGE*/
PERM_INDEX_MIN = 0;//Security Configuration Member
PERM_INDEX_MAX = 32;//Security Configuration Member

/*PERMISSIONVALUE*/
PERM_USE_DOC = -2;//Security Configuration Member
PERM_NOT_RESTRICTED = -1;//Security Configuration Member
PERM_LEVEL_ONE = 0;//Security Configuration Member
PERM_LEVEL_TWO = 1;//Security Configuration Member
PERM_LEVEL_THREE = 2;//Security Configuration Member
PERM_LEVEL_FOUR = 3;//Security Configuration Member
PERM_LEVEL_FIVE = 4;//Security Configuration Member
PERM_LEVEL_SIX = 5;//Security Configuration Member
PERM_LEVEL_SEVEN = 6;//Security Configuration Member
PERM_LEVEL_EIGHT = 7;//Security Configuration Member
PERM_LEVEL_NINE = 8;//Security Configuration Member
PERM_LEVEL_TEN = 9;//Security Configuration Member
PERM_PROHIBITED = 10;//Security Configuration Member

/*POSITIVEASTYPE*/
POSITIVEAS_POS = 0;//Note Number
POSITIVEAS_DR = 1;//Note Number
POSITIVEAS_NEG = 3;//Note Number
POSITIVEAS_BLANK = 4;//Note Number

/*PRINTDOC_FLAGS*/
PRINTDOC_NONE = 0;//Use default options
PRINTDOC_MASTERDOC = 1;//Use master document information if available

/*PRINTDOUBLESIDED*/
PRINTER_DS_DEFAULT = 0;//Page Layout
PRINTER_DS_NONE = 1;//Page Layout
PRINTER_DS_HORIZONTAL = 2;//Page Layout
PRINTER_DS_VERTICAL = 3;//Page Layout

/*PRINTORIENTATION*/
PRINTER_PORTRAIT = 1;//
PRINTER_LANDSCAPE = 2;//

/*PROCESSSTATE*/
DOC_STATE_NORMAL = 0;//Document Member
DOC_STATE_ROLLFORWARD = 1;//Document Member
DOC_STATE_COPYTEMPLATE = 2;//Document Member
DOC_STATE_PRINT = 3;//Document Member
DOC_STATE_EXPORTHTML = 4;//Document Member
DOC_STATE_NORMALFROMSCRIPT = 5;//Document Member
DOC_STATE_CONVERT = 6;//Document Member
DOC_STATE_DESIGNTIME = 7;//Document Member
DOC_STATE_VIEWTIME = 8;//Document Member
DOC_STATE_BACKGROUND = 9;//Document Member
DOC_STATE_GENERATEPDF = 10;//Document Member
DOC_STATE_INTECHECK = 11;//Document Member
DOC_STATE_CONVDLL = 12;//Document Member
DOC_STATE_COMPARE = 13;//Document Member
DOC_STATE_PREVIEW = 14;//Document Member
DOC_STATE_KL_UPDATE = 15;//Document Member
DOC_STATE_CLEANUP = 16;//Document Member: EC20130131 CV-8289

/*RADIOCELLTYPE*/
RADIOTYPE_NORMAL = 0;//Normal Radio-Button: normal radio button
RADIOTYPE_PLUSMINUS = 1;//Plus/Minus Radio-Button: +/- radio button
RADIOTYPE_FOLDER = 2;//Folder Radio-Button: folder radio button
RADIOTYPE_CUSTOM = 3;//Custom Radio-Button: radio button with static image (one of the built-in images)
RADIOTYPE_DYNAMIC = 4;//Dynamic Radio-Button: radio button with dynamic image (one of the images managed by image-manager)

/*RENUMBERCELLFLAGS*/
RENUMBER_ALLCELLS = 1;//Renumber all cells
RENUMBER_REFCALCS = 2;//Rereference calculations
RENUMBER_REFDIAG = 4;//Rereference Diagnostics
RENUMBER_EXISTSCALC = 8;//Cell exists. Maintain original calculation
RENUMBER_EXISTSDATA = 16;//Cell exists. Maintain original data
RENUMBER_EXISTSATTR = 32;//Cell exists. Maintain original attributes
//RENUMBER_REFHIDEALL = apply;//Rereference all skip/hide conditions and graphs to new cell number
RENUMBER_REFHIDECURPARA = 128;//Rereference hide/skip conditions and graphs in current paragraph/range
RENUMBER_REFHIDEENONE = 256;//Do not rereference any hide/skip conditions and graphs

/*RESERVESPACE*/
RESERVESPACE_DEFAULT = 0;//Note Number
RESERVESPACE_SPACE = 1;//Note Number
RESERVESPACE_NONE = 2;//Note Number

/*RIBBONBAR_FIXED_STATE*/
/*
FIXED_STATE_NONE

Status of fixed ribbon bar for scripting: None.

FIXED_STATE_COLLAPSED

Status of fixed ribbon bar for scripting: Collapsed.

FIXED_STATE_MINIMIZED

Status of fixed ribbon bar for scripting: Minimized.

FIXED_STATE_MAXIMIZED

Status of fixed ribbon bar for scripting: Maximized.

*/

/*ROLLFORWARDTYPE*/
ROLLFORWARD_NONE = 0;//Roll Forward None
ROLLFORWARD_DIRECT = 1;//Roll Forward Direct
ROLLFORWARD_RELATIVE = 2;//Roll Forward Relative
ROLLFORWARD_EQUATION = 3;//Roll Forward Equation
ROLLFORWARD_RELATIVECOL = 4;//Roll Forward Relative Column
ROLLFORWARD_MAXVALUE = 4;//Roll Forward Max Value

/*ROUNDTO*/
ROUND_1 = 0;//Round to 1
ROUND_1000 = 1;//Round to 1,000
ROUND_1000000 = 2;//Round to 1,000,000
ROUND_1000000000 = 3;//Round to 1,000,000,000
ROUND_10000 = 4;//Round to 10,000
ROUND_100000000 = 5;//Round to 100,000,000

/*SCALEFACTOR*/
SCALE_THOUSANDTHS = -3;//Thousandths
SCALE_HUNDREDTHS = -2;//Hundredths
SCALE_TENTHS = -1;//Tenths
SCALE_DEFAULT = 0;//Document Default
SCALE_NONE = 1;//None
SCALE_TENS = 2;//Tens
SCALE_HUNDREDS = 3;//Hundreds
SCALE_THOUSANDS = 4;//Thousands
SCALE_TEN_THOUSAND = 5;//Ten Thousands
SCALE_HUNDRED_THOUSAND = 6;//Hundred Thousands
SCALE_MILLIONS = 7;//Millions
SCALE_TEN_MILLIONS = 8;//Ten Millions
SCALE_HUNDRED_MILLIONS = 9;//Hundred Millions
SCALE_BILLIONS = 10;//Billions
SCALE_TEN_BILLIONS = 11;//Ten Billions

/*SCRIPTLANGUAGE*/
SCRIPT_JAVA = 0;//JavaScript
SCRIPT_VBSCRIPT = 1;//VBScript

/*SCRIPTTYPE*/
SCRIPTTYPE_SYSTEM = 1;//System Script
SCRIPTTYPE_DOCUMENT = 2;//Document Script
SCRIPTTYPE_GLOBAL = 3;//Global Script

/*SECTIONBOXSIDE*/
BOX_LEFT = 1;//Note Number
BOX_TOP = 2;//Note Number
BOX_RIGHT = 4;//Note Number
BOX_BOTTOM = 8;//Note Number

/*SECTIONSORTORDER*/
SECTIONSORT_ASCENDING = 1;//Note Number
SECTIONSORT_DESCENDING = 0;//Note Number
SECTIONSORT_MANUAL = 2;//Note Number

/*SECTIONSORTRTNCODE*/
SECTIONSORT_OK = 0;//Note Number
SECTIONSORT_NONE = 1;//Note Number
SECTIONSORT_OVERLAP = 2;//Note Number
SECTIONSORT_PARTIAL = 3;//Note Number
SECTIONSORT_SYNTAX = 4;//Note Number
SECTIONSORT_TYPEMISMATCH = 5;//Note Number
SECTIONSORT_BADTYPE = 6;//Note Number
SECTIONSORT_INVALIDGROUP = 7;//Note Number
SECTIONSORT_CANCELLED = 8;//Note Number
SECTIONSORT_INVALIDTABLE = 9;//Note Number
SECTIONSORT_INVALIDTABLECOLUMN = 10;//Note Number

/*SECTLINKTYPE*/
SECT_LINK_NONE = 0;//Not linked
SECT_LINK_KL = 1;//Linked to a KL section
SECT_LINK_KLI = 2;//Linked to a KLI section

/*SECTSKIPTYPE*/
SECTSKIP_NONE = 0;//Skip None
SECTSKIP_NUMERICS_ZERO = 1;//Skip if numeric, radio button and check box cells zero
SECTSKIP_CELLS_BLANK = 2;//Skip if all cells blank
SECTSKIP_ALWAYS = 3;//Skip Section Always
SECTSKIP_NUMERICS_ZERONOCONTROLS = 4;//Skip if all numeric cells zero

/*SHADETYPE*/
SHADE_NONE = 0;//Note Number: Depricated DM07-B6E-3038
SHADE_SDEFAULT = 0;//Note Number
SHADE_SOLID = 1;//Note Number
SHADE_SBDIAG = 2;//Note Number: Depricated DB09-30E-1BB4
SHADE_SDIAGFWD = 2;//Note Number
SHADE_SCROSS = 3;//Note Number
SHADE_SDCROSS = 4;//Note Number
SHADE_SFDIAG = 5;//Note Number: Depricated DB09-30E-1BB4
SHADE_SDIAGBACK = 5;//Note Number
SHADE_SHORIZ = 6;//Note Number
SHADE_SVERT = 7;//Note Number
SHADE_TRANSPARENT = 8;//Note Number
SHADE_DEFAULT = 9;//Note Number: Depricated DM07-B6E-3038
SHADE_SNONE = 9;//Note Number

/*SHADINGSPEC*/
SHADING_DEFAULT = 0;//Note Number
SHADING_ON = 1;//Note Number
SHADING_OFF = 2;//Note Number

/*SHOWCELLINFO*/
SHOWCELL_NONE = 0;//Do not show cell info
SHOWCELL_NUMBER = 1;//Show cell number above cell
SHOWCELL_NUMBER_INSIDE = 2;//Show cell number inside cell
SHOWCELL_SUMGROUP = 3;//Show cell sum group above cell
SHOWCELL_SUMGROUP_INSIDE = 4;//Show cell sum group inside cell

/*SKIPRPINTTYPE*/
SKIP_NEVER = 0;//Note Number
SKIP_NUMERICS_ZERO = 1;//Note Number
SKIP_CELLS_BLANK = 2;//Note Number
SKIP_ALWAYS = 3;//Note Number
SKIP_NOAPPLY = 4;//Note Number
SKIP_NUMERICS_ZERONOCONTROLS = 5;//Note Number

/*SORTKEYTYPE*/
SORTKEYTYPE_CALCULATION = 0;//Note Number
SORTKEYTYPE_PARAGRAPHTEXT = 1;//Note Number

/*SYSTEMSCRIPTINDEX*/
SYS_SCRIPT_ONOPENING = 0;//Document Member
SYS_SCRIPT_ONSAVING = 1;//Document Member
SYS_SCRIPT_BEFORERECALC = 2;//Document Member
SYS_SCRIPT_AFTERRECALC = 3;//Document Member
SYS_SCRIPT_ONCLOSING = 4;//Document Member
SYS_SCRIPT_ONACTIVATE = 5;//Document Member
SYS_SCRIPT_ONRECORDSELECT = 6;//Document Member
SYS_SCRIPT_ONCOPYTEMPLATE = 7;//Document Member
SYS_SCRIPT_ONMODECHANGED = 8;//Document Member
SYS_SCRIPT_ONROLLFORWARD = 9;//Document Member
SYS_SCRIPT_ONCELLLINK = 10;//Document Member
SYS_SCRIPT_TASKFUNCTIONS = 11;//Document Member
SYS_SCRIPT_ONNEWKLENTRY = 12;//Document Member
SYS_SCRIPT_ONGOTOKLENTRY = 13;//Document Member
SYS_SCRIPT_ONDELETEKLENTRY = 14;//Document Member
SYS_SCRIPT_ONPRINT = 15;//Document Member
SYS_SCRIPT_ONPROGRAMEVENT = 16;//Document Member
SYS_SCRIPT_ONXMLSTORAGE = 17;//Document Member
SYS_SCRIPT_ONXBRLEXPORT = 18;//Document Member
SYS_SCRIPT_ONCLEANUP = 19;//Document Member

/*SYSTEMSCRIPTINDEXRANGE*/
SYS_SCRIPT_INDEX_MIN = 0;//Document Member
SYS_SCRIPT_INDEX_MAX = 10;//Document Member

/*TABJUSTIFICATION*/
TAB_JUST_LEFT = 0;//Tab Member
TAB_JUST_CENTER = 1;//Tab Member
TAB_JUST_RIGHT = 2;//Tab Member

/*TABLEADERTYPES*/
TAB_LEADER_NONE = 0;//Tab Member
TAB_LEADER_LOWDOT = 1;//Tab Member
TAB_LEADER_HIGHDOT = 2;//Tab Member
TAB_LEADER_UNDERLINE = 3;//Tab Member

/*TABLEBORDERPOSITION*/
TABLEBORDER_ALL = 0;//Note Number
TABLEBORDER_LEFT = 1;//Note Number
TABLEBORDER_RIGHT = 2;//Note Number
TABLEBORDER_TOP = 3;//Note Number
TABLEBORDER_BOTTOM = 4;//Note Number
TABLEBORDER_VINTERIOR = 5;//Note Number
TABLEBORDER_HINTERIOR = 6;//Note Number

/*TABLEBORDERSTYLE*/
TABLEBORDER_NONE = 0;//Note Number
TABLEBORDER_SINGLE = 1;//Note Number
TABLEBORDER_SHORT = 2;//Note Number
TABLEBORDER_MEDIUM = 3;//Note Number
TABLEBORDER_LONG = 4;//Note Number
TABLEBORDER_DOUBLE = 5;//Note Number

/*TABLECOLUMNBREAKTYPE*/
COLUMN_BREAK_NO_BREAK = 1;//Note Number
COLUMN_BREAK_VERTICAL_BREAK = 2;//Note Number
COLUMN_BREAK_PAGE_BREAK = 3;//Note Number

/*TABLECOLUMNREPEATTYPE*/
COLUMN_REPEAT_NO_REPEAT = 1;//Note Number
COLUMN_REPEAT_EVERY_PAGE = 2;//Note Number
COLUMN_REPEAT_ODD_PAGES = 3;//Note Number
COLUMN_REPEAT_FIRST_AND_EVEN_PAGES = 4;//Note Number
COLUMN_REPEAT_EVERY_FOURTH_PAGE = 5;//Note Number
COLUMN_REPEAT_FIRST_AND_EVERY_SECOND_EVEN_PAGE = 6;//Note Number

/*TABLECOLUMNTYPE*/
COLTYPE_GENERAL	Note Number	0;//Note Number
//COLTYPE_GENERAL_MOD = Modifiable
COLTYPE_NUMERIC = 2;//Note Number
COLTYPE_FIXED_ALPHA = 3;//Note Number
//COLTYPE_VARIABLE_APHA = compatibility
COLTYPE_VARIABLE_ALPHA = 4;//Note Number

/*TABLECOLUMNTYPES*/
COLUMN_GENERAL = 0;//Table Member
COLUMN_NUMERIC_CELL = 1;//Table Member
COLUMN_ALPHANUM_CELL_COL_WIDTH = 2;//Table Member
COLUMN_ALPHANUM_CELL_VAR_WIDTH = 3;//Table Member

/*TABLEMODFLAGS*/
TABLEMODFLAGS_COLSPACESEPARATE = 0x00008000;//TableMod Flags

/*TABLESORTORDER*/
TABLESORT_ASCENDING = 0;//Note Number
TABLESORT_DESCENDING = 1;//Note Number
TABLESORT_MANUAL = 2;//Note Number


/*TIMECASEVEIWDOCUMENTTYPES*/
TIME_CVDOC_UNKNOWN = -1;//Document Member
TIME_CVDOC_LETTER = 0;//Document Member
TIME_CVDOC_INVOICE = 1;//Document Member
TIME_CVDOC_MISCELLANEOUSINVOICE = 2;//Document Member
TIME_CVDOC_CREDITNOTE = 3;//Document Member
TIME_CVDOC_STATEMENT = 4;//Document Member
TIME_CVDOC_GROUPSTATEMENT = 5;//Document Member
TIME_CVDOC_TAXINVOICE = 6;//Document Member

/*TOCFORMATCOUNTERFORMATS*/
COUNTER_FORMAT_CUSTOM = 0;//TOC Format Member
COUNTER_FORMAT_NONE = 1;//TOC Format Member
COUNTER_FORMAT_DOT_SUFFIX = 2;//TOC Format Member
COUNTER_FORMAT_BRACKET_SUFFIX = 3;//TOC Format Member
COUNTER_FORMAT_BRACKETS = 4;//TOC Format Member
COUNTER_FORMAT_HYPHENS = 5;//TOC Format Member
COUNTER_FORMAT_SQUARE_BRACKETS = 6;//TOC Format Member


/*TOCPAGENUMLOCATION*/
TOC_PAGENUM_SUPPRESS = 0;//Table of Contents Member
TOC_PAGENUM_MARGIN = 1;//Table of Contents Member
TOC_PAGENUM_ATTEXT = 2;//Table of Contents Member
TOC_PAGENUM_LEFTAT = 3;//Table of Contents Member
TOC_PAGENUM_CENTREAT = 4;//Table of Contents Member
TOC_PAGENUM_RIGHTAT = 5;//Table of Contents Member

/*TOCPAGETITLELOCATION*/
TOC_PAGETITLE_MARGIN = 0;//Table of Contents Member
TOC_PAGETITLE_LEFTAT = 1;//Table of Contents Member
TOC_PAGETITLE_CENTREAT = 2;//Table of Contents Member
TOC_PAGETITLE_RIGHTAT = 3;//Table of Contents Member

/*TOCSKIPTYPE*/
TOC_SKIP_WHEN_PARA_SKIPPED = 0;//Skip this TOC marker when the paragraph is skipped
TOC_SKIP_NEVER = 1;//Never skips this TOC marker
TOC_SKIP_USING_CONDITION = 2;//Skips using the TOC Skip Condition

/*UNDERLINEDEFINITIONSINDEX*/
UNDERLINEDEFINITIONS_MININDEX = 0;//Cell Line Options Member
UNDERLINE_DEF_SINGLE = 0;//Cell Line Options Member
UNDERLINE_DEF_DOUBLE = 1;//Cell Line Options Member
UNDERLINE_DEF_SINGLEDESCENDED = 2;//Cell Line Options Member
UNDERLINE_DEF_DOUBLEDESCENDED = 3;//Cell Line Options Member
UNDERLINE_DEF_SINGLE2 = 4;//Cell Line Options Member
UNDERLINE_DEF_DOUBLE2 = 5;//Cell Line Options Member
UNDERLINE_DEF_SINGLEDESCENDED2 = 6;//Cell Line Options Member
UNDERLINE_DEF_DOUBLEDESCENDED2 = 7;//Cell Line Options Member
UNDERLINE_DEF_USER_1 = 8;//Cell Line Options Member
UNDERLINE_DEF_USER_2 = 9;//Cell Line Options Member
UNDERLINE_DEF_USER_3 = 10;//Cell Line Options Member
UNDERLINE_DEF_USER_4 = 11;//Cell Line Options Member
UNDERLINE_DEF_USER_5 = 12;//Cell Line Options Member
UNDERLINE_DEF_USER_6 = 13;//Cell Line Options Member
UNDERLINE_DEF_USER_7 = 14;//Cell Line Options Member
UNDERLINEDEFINITIONS_MAXINDEX = 14;//Cell Line Options Member

/*UNDERLINEDEFINITIONTYPE*/
UNDERLINETYPE_SINGLE = 0;//Cell Line Options Member
UNDERLINETYPE_DOUBLE = 1;//Cell Line Options Member
UNDERLINETYPE_TRIPLE = 2;//Cell Line Options Member

/*UNDERLINEINDEX*/
FIRST_UNDERLINE = 0;//Cell Line Options Member
SECOND_UNDERLINE = 1;//Cell Line Options Member
THIRD_UNDERLINE = 2;//Cell Line Options Member

/*UNDERLINESPEC*/
UNDERLINE_DEFAULT = 0;//Note Number
UNDERLINE_ON = 1;//Note Number
UNDERLINE_OFF = 2;//Note Number

/*UNDERLINESTYLE*/
UNDERLINE_SOLID = 0;//Cell Line Options Member
UNDERLINE_DASH = 1;//Cell Line Options Member
UNDERLINE_DOTTED = 2;//Cell Line Options Member
UNDERLINE_FINEDOTS = 3;//Cell Line Options Member

/*UNITTYPE*/
UNIT_MONETARY = 0;//Monetary
UNIT_OTHER = 1;//Other
UNIT_DIVISION = 2;//Division
UNIT_MULTIPLE = 3;//Multiple

/*VERTALIGN*/
VERT_DEFAULT = 0;//Note Number
VERT_TOP = 1;//Note Number
VERT_CENTER = 2;//Note Number
VERT_BOTTOM = 3;//Note Number
VERT_JUSTIFY = 4;//Note Number

/*WIDTHTYPE*/
WIDTH_VARIABLE = 0;//Cell Number
CELL_WIDTH_VARIABLE = 0;//Cell Number
WIDTH_EXACT = 1;//Cell Number
CELL_WIDTH_EXACT = 1;//Cell Number
WIDTH_ATLEAST = 2;//Cell Number
CELL_WIDTH_ATLEAST = 2;//Cell Number
WIDTH_ATMOST = 3;//Cell Number
CELL_WIDTH_ATMOST = 3;//Cell Number
CELL_WIDTH_COLUMN = 4;//Cell Number

/*XBRLDGNMODE*/
XBRLDGN_NEVER = 0;//Never Display
XBRLDGN_ERRORS = 1;//Errors Only
XBRLDGN_BOTH = 2;//Errors or Warnings
XBRLDGN_ALWAYS = 3;//Always Display

/*XBRLMULTTAGACTION*/
XBRLMTA_NONE = 0;//No action specified
XBRLMTA_ADD = 1;//Add values in duplicate tags together
XBRLMTA_USEIFEQUAL = 4;//Use the tag if all values are equal

/*XBRLVER*/
XBRLVER_20 = 0;//XBRL v2.0
XBRLVER_21 = 1;//XBRL v2.1

/*ZEROASTYPE*/
ZEROAS_DASH = 0;//Zero As Type Dash
ZEROAS_ZERO = 1;//Zero As Type Zero
ZEROAS_NIL = 2;//Zero As Type NIL
ZEROAS_BLANK = 3;//Zero As Type Blank
ZEROAS_NA = 4;//Zero As Type N/A
ZEROAS_LCNIL = 5;//Zero As Type nil
ZEROAS_DASHES = 6;//Zero As Type Dashes
ZEROAS_NIL2 = 7;//As Type Nil Zero 
ZEROAS_PLACEHLDTXT = 8;//Zero As Type Placeholder Text


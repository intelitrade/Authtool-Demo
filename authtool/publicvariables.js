var iPos = 0;//Tracks position of a list item
var sAction = "";
var sCurrentlyElementSelected = "";
var sObjectType = ""
var sCurrentTable = "";
var oRowElement;
var oSectionElement;
var oTableElement;
var oParaElement
var oRefElement;
var sRefElement = "";
var bMultiSelect = false;//Tracks if multiple items have been selected
var aSelectedItems = [];
var aCopiedItems = [];
var aDeletedItems = [];
var iTest = 1;
/*var oElement;
var oRefElement;
var sRefElement = "";

*/
var CCOLTYPE = "CCOLTYPE";
var CROWTYPE = "CROWTYPE";
var CROWID = "CROWID"
//CM 27 May 2011
var HORIZONTAL_NOTETABLE8 = "N21";//Horizontal Note table 8

//Column custom properties
var CALC_COL		= "CA"; // Calculation column
var SPACE_COL		= "CS"; // Space column
var DESC_COL		= "CD"; // Description column
var VARIANCE_COL	= "CV"; // Variance column
var TOTAL_COL		= "CT"; // Total column
var HIDDEN_COL		= "CC"; // Hidden (controls) column
var MAP_COL		= "CM"; // Map Number column -- also the link to the Note
var NOTEREF_COL		= "CN"; // Note reference column
var PERCENT_COL		= "PC"; // Percentage Column
var LS_COL		= "LS"; // Lead sheet column
var PRINT_CONT_COL	= "CP"; // Print control column
var ST_REF_COL		= "ST"; // Statement reference column
var CTRLCELLCOL		= "CR"; // Column holding control cells
var ANNOTATION_COL	= "AN";	// Annotations column
var PAGENO_COL		= "PG"
var CACCBASIS_COL	= "CACCBASIS"//Basis of accounting column
var PYC_COL	= "PYC"//Company prior year column
var AYC_COL	="AYC"//Company current year column
var PYG_COL	= "PYG"
var AYG_COL	= "AYG"
var PYG2_COL	= "PYG2" // Second comparative - group
var PYC2_COL	= "PYC2" // second comparative - company

//Row custom properties values
var CONTROL_ROW			= "RH"; // Control row
var HEADING_ROW			= "D1"; // Heading row
var SUBHEADING_ROW		= "D2"; // Subheading row
var CALC1_ROW			= "C1"; // Calculation row  [grp]	->	grp(c1,ayg,entity(groupid,rc(-4)))
var CALC2_ROW			= "C2"; // Calculation row  [-grp]	->	-grp(c1,ayg,entity(groupid,rc(-4)))
var CALC3_ROW			= "C3"; // Calculation row [grpent,1]	->	grpent(c1,ayg,entity(groupid,rc(-4)),1)
var CALC4_ROW			= "C4"; // Calculation row [grpent,-1]	->	grpent(c1,ayg,entity(groupid,rc(-4)),-1)
var CALC5_ROW			= "C5"; // Cell reference calculation row / Multiple calculation row
var CALC6_ROW			= "C6"; // [grp] with Description based on map col less last 4 characters used in PPEV recon note
var CALC7_ROW			= "C7"; // [grp] with additional terms and conditions paragraph below
var CALC8_ROW			= "C8"; // pos of grp calcultion
var CALC9_ROW			= "C9"; // pos of -grp calculation
var CALC10_ROW			= "C0"; // Movement calculation - PY - CY
var CALCB_ROW			= "CB"; // PY value pulls through to the CY col. PY col is input
var CALCC_ROW			= "CC"; // Movement calculation - PY - CY -- Negative
var INPUT_ROW			= "I1"; // Input row
var TEXTONLY_ROW		= "I2";	// Text only row - skips if row above is skipped
var TEXTCALC1_ROW		= "I3"; // Input row with opening balance being last years total
var INPUTPERCENT_ROW		= "I4"; // Input row with percentage columns
var INPUTROLLFORWARD_ROW	= "I5"; // Input row roll forward to PY
var INPUTDESCNOT_ROW		= "I6";	// Input row, description columns not input
var INPUTBULLET_ROW		= "I7"; // Input row with bullet, description pulls through from Map number
var INPUTTOTAL_ROW		= "I8"; // Input row with last column being a total of other 3 columns
var INPUTANDTEXT_ROW		= "I9"; // Input row with additional paragraph input text
//I10 needs to be added it is an input row with the cells having 2 decimal places
var ACTCALC1_ROW		= "A1"; // Calculation row pulling through from account numbers
var ACTCALC2_ROW		= "A2"; // Negative Calculation row pulling through from account numbers
var SUBTOTAL_ROW		= "S1"; // Subtotal row
var THINLINE_ROW		= "L1"; // Thin line row
var THINKLINE_ROW		= "L2"; // Thick line row
var TOTAL_ROW			= "T1"; // Total row
var LINKTOTAL_ROW		= "T2"; // Total row with linkage
var LINKTOTALSKIP_ROW		= "T3"; // Total row with linkage --Always skipped and hidden
var LINKSUBTOTAL_ROW		= "T4"; // Total row based on row position in other tables
var YEARHEADER_ROW		= "RY"; // Year heading row
var NOTE_ROW			= "RN"; // Note linkage row
var PAGEBREAK_ROW		= "RP"; // Page break row
var BALCHK_ROW			= "BL";	// Balance check row
var CHEADDEFAULTPRINT		= "CHEADDEFAULTPRINT"
var CHEADSKIPDEFAULT		= "CHEADSKIPDEFAULT"
var CACCPOLLINK			= "CACCPOLLINK" //sets which accouting policies are associated with this note
//var HEADING_ROW			= "HEADING_ROW" //Heading row
var COL_MAPPING_ROW		= "COL_MAPPING_ROW"	//Columng mapping row

// Table custom properties
var GENERIC_TABLE		= "G2"; // Generic table
var GENERIC_TABLE_SCHED		= "G5"; // Generic schedules table
var CONTROL_TABLE		= "PC"; // Controls table
var SECTIONHEAD_TABLE		= "H1"; // section Header table
var SECTIONSUBHEAD_TABLE	= "H2"; // section Header table
var BS_TABLE			= "B1"; // Balance sheet table
var IS_TABLE			= "I1"; // Income statemment table
var DETIS_TABLE			= "D1" // Detailed Income Statement table
var DETIS_SUBTABLE		= "D2" // Detailed income statement subtable
var CF_TABLE			= "C1"; // Cash flow statement table
var EQUITY_TABLE		= "E1"; // Statement of changes in equity
var FARM_TABLE			= "F1"; // Farming table
var NOTECONTROL_TABLE		= "N1"; // Note number / control table
var NOTE1_TABLE			= "N2"; // Note type 1 table - Standard note table
var NOTE2_TABLE			= "N3"; // Note type 2 table - Note Total table
var NOTE3_TABLE			= "N4"; // Note type 3 table - balance check
var INVEST_NOTE_TABLE		= "N5"; // Note Investment type table - used in investment to subs, JV's & Associates
var NOTE9_TABLE			= "N9"; // Note type 9 table - carrying value table
var HORIZONTAL_NOTE_TABLE	= "N10"; // Horizontal Note table - used in horizontal notes
var NOTE_HIDE_TABLE		= "N11"; // Hidden note table used for linking
var NOTE_TOTAL3_TABLE		= "N12"; // as note total  
var CDIRTABLE			= "CDIRTABLE" //Custom property that indicates if directors table to be sorted

var CTABLETYPE = "CTABLETYPE";

//table uses
var MODTABLE			= "MT";	// Table is modifiable and as such has the various controls available, e.g sort rows
var STATEMENTTABLE		= "ST"; // Statement table e.g. Balance sheet
var NOTETABLE			= "NO"; // Note Table
var HEADERCTRL			= "HC"; //Header control table - The following table hold controls associated with the header 

var HEADING_TABLE		= "TH"; //Heading table
var FOOTER_TABLE		= "FT"; // Footer table

var INDEX_TABLE			= "IT" // Index table	

//Mouse pointer or cursor
var sTableLineItemCursor = "hand";
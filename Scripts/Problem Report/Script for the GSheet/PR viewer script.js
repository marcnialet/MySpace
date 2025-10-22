var SHEET_CONFIGURATION                       = "Config"
var SHEET_PR_INFO                             = "Content";
var SHEET_PROBLEMREPORT_SUMMARY               = "Summary";
var SHEET_BACKEND_LOG                         = "Backend Log";
var SHEET_AUDITS_AND_MESSAGES                 = "Audits and Messages";
var SHEET_ORACLE_LOG                          = "Oracle Log";
var SHEET_SOFTWARESETTINGS                    = "Software Settings";
var SHEET_LIS_SETTINGS                        = "LIS Settings";
var SHEET_SYSTEM_CONFIG_SETTINGS              = "System Configuration  Settings";
var SHEET_TNS_CONFIG_SETTINGS                 = "Tsn Configuration Settings";
var SHEET_EXECUTION_LOG                       = "Execution Log";
var SHEET_CYCLE_JOB_TRACE                     = "Cycle job trace";
var SHEET_SAMPLE_TEST_RESULTS_COAG            = "Sample Test results coag";
var SHEET_CALIBRATION_TEST_RESULTS_COAG       = "Calibration Test results coag";
var SHEET_SAMPLE_TEST_RESULTS_CC              = "Sample Test results cc";
var SHEET_CALIBRATION_TEST_RESULTS_CC         = "Calibration Test results cc";
var SHEET_SERVICE_HOSTING_COMITSERVICES_LOG   = "Service.Hosting.ComitServices Log";
var SHEET_SYSTEM_SERVICEHOST_COMITSERVICES_LOG= "System.ServiceHost.ComitServices.log";
var SHEET_UI_TRACE_LOG                        = "UI Log";
var SHEET_CALCULATION_LOG                     = "Calculation Log";


var NAMED_RANGE_FILES_AND_FOLDERS             			="ProblemReportFilesAndFolders";
var NAMED_RANGE_BACKENDLOG                    			="BackendLog";
var NAMED_RANGE_AUDITSANDMESSAGES             			="AuditsAndMessages";
var NAMED_RANGE_ORACLE_LOG                    			="OracleLog"; 
var NAMED_RANGE_SOFTWARESETTINGS              			="SoftwareSettings"; 
var NAMED_RANGE_LIS_SETTINGS                        = "LISSettings";
var NAMED_RANGE_SYSTEM_CONFIG_SETTINGS              = "SystemConfigurationSettings";
var NAMED_RANGE_TNS_CONFIG_SETTINGS                 = "TsnConfigurationSettings";
var NAMED_RANGE_EXECUTION_LOG                 			="ExecutionLog"; 
var NAMED_RANGE_PROBLEMREPORT_SUMMARY         			= "ProblemReportSummary"; 
var NAMED_RANGE_CYCLE_JOB_TRACE               			= "CycleJobTrace";
var NAMED_RANGE_SAMPLE_TEST_RESULTS_COAG      			= "SampleTestResultsCoag";
var NAMED_RANGE_CALIBRATION_TEST_RESULTS_COAG 			= "CalibrationTestResultsCoag";
var NAMED_RANGE_SAMPLE_TEST_RESULTS_CC        			= "SampleTestResultsCC";
var NAMED_RANGE_CALIBRATION_TEST_RESULTS_CC       	= "CalibrationTestResultsCC";
var NAMED_RANGE_SERVICE_HOSTING_COMITSERVICES_LOG   = "ServiceHostingComitServicesLog"
var NAMED_RANGE_SYSTEM_SERVICEHOST_COMITSERVICES_LOG= "SystemServiceHostComitServicesLog"
var NAMED_RANGE_UI_TRACE_LOG                        = "UITraceLog"
var NAMED_RANGE_CALCULATION_LOG                     = "CalculationLog";

function onOpen() {
  var ui = SpreadsheetApp.getUi();
  // Or DocumentApp or FormApp.
  ui.createMenu('LabCoag')
      .addItem('Load Problem Report', 'menuLoadProblemReport')  
      .addSeparator()
      .addItem('Load File Information', 'menuLoadFileInformation')  
      .addItem('Load Summary', 'menuLoadSummary')  
      .addItem('Load Backend Log (last file only)', 'menuLoadBackendLog')      
      .addItem('Load UI Log (last file only)', 'menuUITraceLog')      
      .addItem('Load Audit&Messages Log', 'menuLoadAuditsAndMessages') 
      .addItem('Load Oracle Log', 'menuLoadOracleLog') 
      .addItem('Load Software Settings', 'menuLoadSoftwareSettings')       
      .addItem('Load Calculation log', 'menuLoadCalculationLog') 
      .addItem('Load ServiceHosting.ComitServices.log', 'menuLoadServiceHostingComitServicesLog') 
      .addItem('Load SystemServiceHost.ComitServices.log', 'menuLoadSystemServiceHostComitServicesLog')       
      .addSeparator()
      .addItem('Analyze content', 'menuAnalyzeFileContent')  
      .addSeparator()
      .addSubMenu(ui.createMenu('Delete contents')
        .addItem('Delete All', 'menuDeleteAll')  
        .addSeparator()
        .addItem('Delete Content log', 'menuDeleteContentlog')  
        .addItem('Delete Summary log', 'menuDeleteSummarylog')        
        .addItem('Delete Backend log', 'menuDeleteBackendlog')  
        .addItem('Delete Audits&Message log', 'menuDeleteAuditsAndMessageslog')
        .addItem('Delete Oracle log', 'menuDeleteOraclelog')
        .addItem('Delete Commit logs', 'menuDeleteCommitLogs')
        .addItem('Delete Software Settings', 'menuDeleteSoftwareSettings')
        .addItem('Delete Execution log', 'menuDeleteExecutionlog')
        .addItem('Delete Calculationlog', 'menuDeleteCalculationLog')
        .addItem('Delete Sample test results coag', 'menuDeleteSampleTestResultsCoag')
        .addItem('Delete Calibration test results coag', 'menuDeleteCalibrationTestResultsCoag')
        .addItem('Delete Calibration test results cc', 'menuDeleteSampleTestResultsCC')
        .addItem('Delete Calibration test results cc', 'menuDeleteCalibrationTestResultsCC'))  
      .addSeparator()
      .addSubMenu(ui.createMenu('Load Results')
        .addItem('Load all results', 'menuLoadAllResults')  
        .addSeparator()
        .addItem('Cycle job trace', 'menuLoadCycleJobTrace') 
        .addItem('Sample test results coag', 'menuLoadSampleTestResultsCoag') 
        .addItem('Calibration test results coag', 'menuLoadCalibrationTestResultsCoag') 
        .addItem('Sample test results cc', 'menuLoadSampleTestResultsCC') 
        .addItem('Calibration test results cc', 'menuLoadCalibrationTestResultsCC'))
      .addToUi();
}
function menuLoadProblemReport() 
{ 
  var starttime = logMethodStart("Load problem report");
   
  SpreadsheetApp.flush();
  deleteAllSheets();

  //SpreadsheetApp.flush();
  //createAllSheets();
  
  menuLoadFileInformation();
  menuLoadSummary();
  menuLoadBackendLog();  
  menuUITraceLog();
  menuLoadAuditsAndMessages();
  menuLoadOracleLog();
  menuLoadSoftwareSettings();
  menuLoadServiceHostingComitServicesLog();
  menuLoadSystemServiceHostComitServicesLog();     
  menuLoadCalculationLog();      

  logMethodEnd(starttime,"Load problem report");
  AppendExecutionLogToLogsSheet();
}
function menuLoadAllResults()
{
  menuLoadCycleJobTrace();
  menuLoadSampleTestResultsCoag();
  menuLoadCalibrationTestResultsCoag();
  menuLoadSampleTestResultsCC();
  menuLoadCalibrationTestResultsCC();
}
function menuDeleteAll()
{
  menuDeleteContentlog();
  menuDeleteSummarylog();  
  menuDeleteBackendlog();
  menuDeleteAuditsAndMessageslog();
  menuDeleteCommitLogs();
  menuDeleteOraclelog();  
  menuDeleteSoftwareSettings();
  menuDeleteExecutionlog();  
  menuDeleteCalculationLog();
  menuDeleteSampleTestResultsCoag();
  menuDeleteCalibrationTestResultsCoag();
  menuDeleteSampleTestResultsCC();
  menuDeleteCalibrationTestResultsCC();
}
function menuDeleteCalculationLog()
{  
  deleteDataRangeContent(SHEET_CALCULATION_LOG, NAMED_RANGE_CALCULATION_LOG);
}
function menuDeleteSampleTestResultsCoag()
{  
  deleteDataRangeContent(SHEET_SAMPLE_TEST_RESULTS_COAG, NAMED_RANGE_SAMPLE_TEST_RESULTS_COAG);
}
function menuDeleteCalibrationTestResultsCoag()
{  
  deleteDataRangeContent(SHEET_CALIBRATION_TEST_RESULTS_COAG, NAMED_RANGE_CALIBRATION_TEST_RESULTS_COAG);
}
function menuDeleteSampleTestResultsCC()
{  
  deleteDataRangeContent(SHEET_SAMPLE_TEST_RESULTS_CC, NAMED_RANGE_SAMPLE_TEST_RESULTS_CC);
}
function menuDeleteCalibrationTestResultsCC()
{  
  deleteDataRangeContent(SHEET_CALIBRATION_TEST_RESULTS_CC, NAMED_RANGE_CALIBRATION_TEST_RESULTS_CC);
}
function menuDeleteSoftwareSettings()
{  
  deleteDataRangeContent(SHEET_SOFTWARESETTINGS, NAMED_RANGE_SOFTWARESETTINGS);
  deleteDataRangeContent(SHEET_LIS_SETTINGS, NAMED_RANGE_LIS_SETTINGS);
  deleteDataRangeContent(SHEET_SYSTEM_CONFIG_SETTINGS, NAMED_RANGE_SYSTEM_CONFIG_SETTINGS);
  deleteDataRangeContent(SHEET_TNS_CONFIG_SETTINGS, NAMED_RANGE_TNS_CONFIG_SETTINGS);
}
function menuDeleteContentlog()
{  
  deleteDataRangeContent(SHEET_PR_INFO, NAMED_RANGE_FILES_AND_FOLDERS);
}
function menuDeleteCommitLogs()
{  
  deleteDataRangeContent(SHEET_SERVICE_HOSTING_COMITSERVICES_LOG, NAMED_RANGE_SERVICE_HOSTING_COMITSERVICES_LOG);
  deleteDataRangeContent(SHEET_SYSTEM_SERVICEHOST_COMITSERVICES_LOG, NAMED_RANGE_SYSTEM_SERVICEHOST_COMITSERVICES_LOG);
}
function menuDeleteOraclelog()
{  
  deleteDataRangeContent(SHEET_ORACLE_LOG, NAMED_RANGE_ORACLE_LOG);
}
function menuDeleteSummarylog()
{  
  deleteDataRangeContent(SHEET_PROBLEMREPORT_SUMMARY, NAMED_RANGE_PROBLEMREPORT_SUMMARY);
}
function menuDeleteExecutionlog()
{  
  deleteDataRangeContent(SHEET_EXECUTION_LOG, NAMED_RANGE_EXECUTION_LOG);
}
function menuDeleteBackendlog()
{  
  deleteDataRangeContent(SHEET_BACKEND_LOG, NAMED_RANGE_BACKENDLOG);
}
function menuDeleteAuditsAndMessageslog()
{  
  deleteDataRangeContent(SHEET_AUDITS_AND_MESSAGES, NAMED_RANGE_AUDITSANDMESSAGES);
}
function menuAnalyzeFileContent()
{  
  AnalyzeFileContent();
  AppendExecutionLogToLogsSheet();
}
function logMethodStart(methodname)
{   
  Logger.log("-------- Start: " + methodname + " --------");  
  return new Date().getTime(); // Record the start time in milliseconds
}
function logMethodEnd(startTime, methodname)
{
  var endTime = new Date().getTime(); // Record the end time in milliseconds
  var elapsedTime = endTime - startTime; // Calculate the elapsed time
  Logger.log("-------- End ["+formatTime(elapsedTime)+"]: "+methodname+" --------");
}

function formatTime(milliseconds) {
  var minutes = Math.floor(milliseconds / 60000);
  var seconds = ((milliseconds % 60000) / 1000).toFixed(0);
  return minutes + ' min:' + (seconds < 10 ? '0' : '') + seconds+" sec";
}

function menuLoadFileInformation()
{ 
  Utilities.sleep(150);
  var starttime = logMethodStart("LoadFileInformation");

  CreateSheet(SHEET_PR_INFO);
  LoadProblemReportInformation();
  FormatProblemReportInfo();  
  
  logMethodEnd(starttime,"LoadFileInformation");

  AppendExecutionLogToLogsSheet();
}
function menuLoadSummary()
{  
  Utilities.sleep(150);
  var starttime = logMethodStart("LoadSummary");
  CreateSheet(SHEET_PROBLEMREPORT_SUMMARY);
  ReadProblemReportSummary();
  FormatProblemReportSummary();
  logMethodEnd(starttime,"LoadSummary");

  AppendExecutionLogToLogsSheet();
}
function menuUITraceLog() 
{  
  Utilities.sleep(150);
  var starttime = logMethodStart("UITraceLog");
  
  CreateSheet(SHEET_UI_TRACE_LOG) ;
  LoadUITraceLogFromCSV();
  FormatUITracelog();
  logMethodEnd(starttime,"UITraceLog");

  AppendExecutionLogToLogsSheet();
}
function menuLoadBackendLog() 
{  
  Utilities.sleep(150);
  var starttime = logMethodStart("LoadBackendLog");
  
  CreateSheet(SHEET_BACKEND_LOG) ;
  LoadBackendLogFromCSV();
  FormatBackendlog();
  logMethodEnd(starttime,"LoadBackendLog");

  AppendExecutionLogToLogsSheet();

}
function menuLoadCycleJobTrace()
{ 
  Utilities.sleep(150);
  var starttime = logMethodStart("LoadCycleJobTrace");
  
  CreateSheet(SHEET_CYCLE_JOB_TRACE);  
  if(LoadResultsFromCSV("CycleJobTrace", SHEET_CYCLE_JOB_TRACE,','))
  {
    CreateFilterAndDataRangeInSheet(SHEET_CYCLE_JOB_TRACE, NAMED_RANGE_CYCLE_JOB_TRACE);
  }
  logMethodEnd(starttime,"LoadCycleJobTrace");
  AppendExecutionLogToLogsSheet();
}
function menuLoadSampleTestResultsCoag()
{ 
  Utilities.sleep(150);
  var starttime = logMethodStart("LoadSampleTestResultsCoag");
  
  CreateSheet(SHEET_SAMPLE_TEST_RESULTS_COAG);  
  if(LoadResultsFromCSV("sampleTestResults_Coag", SHEET_SAMPLE_TEST_RESULTS_COAG,';'))
  {
    CreateFilterAndDataRangeInSheet(SHEET_SAMPLE_TEST_RESULTS_COAG, NAMED_RANGE_SAMPLE_TEST_RESULTS_COAG);
  }
  logMethodEnd(starttime,"LoadSampleTestResultsCoag");

  AppendExecutionLogToLogsSheet();
}
function menuLoadCalibrationTestResultsCoag()
{  
  Utilities.sleep(150);
  var starttime = logMethodStart("LoadCalibrationTestResultsCoag");
  
  CreateSheet(SHEET_CALIBRATION_TEST_RESULTS_COAG);  
  if(LoadResultsFromCSV("calibrationResults_Coag", SHEET_CALIBRATION_TEST_RESULTS_COAG,';'))
  {
    CreateFilterAndDataRangeInSheet(SHEET_CALIBRATION_TEST_RESULTS_COAG, NAMED_RANGE_CALIBRATION_TEST_RESULTS_COAG);
  }
  logMethodEnd(starttime,"LoadCalibrationTestResultsCoag");

  AppendExecutionLogToLogsSheet();
}
function menuLoadSampleTestResultsCC()
{  
  Utilities.sleep(150);
  var starttime = logMethodStart("LoadSampleTestResultsCC");
  
  CreateSheet(SHEET_SAMPLE_TEST_RESULTS_CC);   
  if(LoadResultsFromCSV("sampleTestResults_CC", SHEET_SAMPLE_TEST_RESULTS_CC,';'))
  {
    CreateFilterAndDataRangeInSheet(SHEET_SAMPLE_TEST_RESULTS_CC, NAMED_RANGE_SAMPLE_TEST_RESULTS_CC);
  }
  logMethodEnd(starttime,"LoadSampleTestResultsCC");

  AppendExecutionLogToLogsSheet();
}
function menuLoadCalibrationTestResultsCC()
{  
  Utilities.sleep(150);
  var starttime = logMethodStart("LoadCalibrationTestResultsCC");
  
  CreateSheet(SHEET_CALIBRATION_TEST_RESULTS_CC);  
  if( LoadResultsFromCSV("calibrationResults_CC", SHEET_CALIBRATION_TEST_RESULTS_CC,';') )
  {
    CreateFilterAndDataRangeInSheet(SHEET_CALIBRATION_TEST_RESULTS_CC, NAMED_RANGE_CALIBRATION_TEST_RESULTS_CC);
  }
  logMethodEnd(starttime,"LoadCalibrationTestResultsCC");

  AppendExecutionLogToLogsSheet();
}
function menuLoadServiceHostingComitServicesLog()
{
  Utilities.sleep(150);
  var starttime = logMethodStart("LoadServiceHostingComitServicesLog");
  
  CreateSheet(SHEET_SERVICE_HOSTING_COMITSERVICES_LOG);  
  LoadServiceHostingComitServicesLog();
  CreateFilterAndDataRangeInSheet(SHEET_SERVICE_HOSTING_COMITSERVICES_LOG,NAMED_RANGE_SERVICE_HOSTING_COMITSERVICES_LOG);
  FormaServiceHostingComitServicesLog();
  logMethodEnd(starttime,"LoadServiceHostingComitServicesLog");

  AppendExecutionLogToLogsSheet();
}
function menuLoadSystemServiceHostComitServicesLog()
{
  Utilities.sleep(150);
  var starttime = logMethodStart("LoadSystemServiceHostComitServicesLog");
  
  CreateSheet(SHEET_SYSTEM_SERVICEHOST_COMITSERVICES_LOG);  
  LoadSystemServiceHostComitServicesLog();
  CreateFilterAndDataRangeInSheet(SHEET_SYSTEM_SERVICEHOST_COMITSERVICES_LOG,NAMED_RANGE_SYSTEM_SERVICEHOST_COMITSERVICES_LOG);
  FormatSystemServiceHostComitServicesLog();
  logMethodEnd(starttime,"LoadSystemServiceHostComitServicesLog");
  
  AppendExecutionLogToLogsSheet();
}
function menuLoadCalculationLog()
{
  Utilities.sleep(150);
  var starttime = logMethodStart("LoadCalculationLog");
  
  CreateSheet(SHEET_CALCULATION_LOG); 
  LoadCalculationLog();
  CreateFilterAndDataRangeInSheet(SHEET_CALCULATION_LOG,NAMED_RANGE_CALCULATION_LOG);
  FormatCalculationLog();
  logMethodEnd(starttime,"LoadCalculationLog");

  AppendExecutionLogToLogsSheet();
}

function menuLoadOracleLog()
{
  Utilities.sleep(150);
  var starttime = logMethodStart("LoadOracleLog");
  
  CreateSheet(SHEET_ORACLE_LOG); 
  LoadOracleDBLog();
  CreateFilterAndDataRangeInSheet(SHEET_ORACLE_LOG,NAMED_RANGE_ORACLE_LOG);
  FormatOracleLog();
  logMethodEnd(starttime,"LoadOracleLog");

  AppendExecutionLogToLogsSheet();
}
function FormaServiceHostingComitServicesLog()
{
  setAlternativeColorToRange(SHEET_SERVICE_HOSTING_COMITSERVICES_LOG, NAMED_RANGE_SERVICE_HOSTING_COMITSERVICES_LOG, SpreadsheetApp.BandingTheme.CYAN)

  var sheet = SpreadsheetApp.getActive().getSheetByName(SHEET_SERVICE_HOSTING_COMITSERVICES_LOG);
  sheet.setColumnWidth(1,1500);  
  sheet.getRange('A:A').setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP); // Set the wrapping to clip for column A
 }
function FormatSystemServiceHostComitServicesLog()
{
  setAlternativeColorToRange(SHEET_SYSTEM_SERVICEHOST_COMITSERVICES_LOG, NAMED_RANGE_SYSTEM_SERVICEHOST_COMITSERVICES_LOG, SpreadsheetApp.BandingTheme.CYAN)

  var sheet = SpreadsheetApp.getActive().getSheetByName(SHEET_SYSTEM_SERVICEHOST_COMITSERVICES_LOG);
  sheet.setColumnWidth(1,1500);  
  sheet.getRange('A:A').setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP); // Set the wrapping to clip for column A
}
function FormatOracleLog()
{
  setAlternativeColorToRange(SHEET_ORACLE_LOG, NAMED_RANGE_ORACLE_LOG, SpreadsheetApp.BandingTheme.CYAN)

  var sheet = SpreadsheetApp.getActive().getSheetByName(SHEET_ORACLE_LOG);
  sheet.setColumnWidth(1,800);  
  sheet.getRange('A:A').setWrapStrategy(SpreadsheetApp.WrapStrategy.OVERFLOW); // Set the wrapping to clip for column A
}
function FormatSoftwareSettings(sheetname,namedrange)
{
  setAlternativeColorToRange(sheetname, namedrange, SpreadsheetApp.BandingTheme.CYAN)

  var sheet = SpreadsheetApp.getActive().getSheetByName(sheetname);
  sheet.setColumnWidth(1,1100);  
  sheet.getRange('A:A').setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP); // Set the wrapping to clip for column A
}
function menuLoadSoftwareSettings()
{  
  loadSettings("ProblemReportExportConfiguration.json",SHEET_SOFTWARESETTINGS, NAMED_RANGE_SOFTWARESETTINGS);
  loadSettings("LisSettings.xml",SHEET_LIS_SETTINGS, NAMED_RANGE_LIS_SETTINGS);
  loadSettings("SystemConfigurationSettings.xml",SHEET_SYSTEM_CONFIG_SETTINGS, NAMED_RANGE_SYSTEM_CONFIG_SETTINGS);
  loadSettings("TsnSystemConfiguration.xml",SHEET_TNS_CONFIG_SETTINGS, NAMED_RANGE_TNS_CONFIG_SETTINGS);
}
function loadSettings(filename, sheetname, namedrange)
{
  Utilities.sleep(150);
  var starttime = logMethodStart("loadSettings");
  
  CreateSheet(sheetname);
  LoadSettingsFile(filename, sheetname);
  CreateFilterAndDataRangeInSheet(sheetname,namedrange);
  FormatSoftwareSettings(sheetname,namedrange);
  AppendExecutionLogToLogsSheet();
  logMethodEnd(starttime,"loadSettings");
}
function menuLoadAuditsAndMessages() 
{  
  Utilities.sleep(150);
  var starttime = logMethodStart("LoadAuditsAndMessages");
  
  CreateSheet(SHEET_AUDITS_AND_MESSAGES);
  LoadAuditsAndMessages();
  AddFormulaColumnsToAuditsAndMessages();
  FormatLogsAndMessages();
  SortNamedRangeAscending(SHEET_AUDITS_AND_MESSAGES,NAMED_RANGE_AUDITSANDMESSAGES, 3);
  logMethodEnd(starttime,"LoadAuditsAndMessages");

  AppendExecutionLogToLogsSheet();
}
function FormatProblemReportSummary()
{
  Logger.log("Applying format to summary"); 
  //formatHeader(SHEET_PROBLEMREPORT_SUMMARY,'A1:B1','#4a86e8');
  setAlternativeColorToRange(SHEET_PROBLEMREPORT_SUMMARY, NAMED_RANGE_PROBLEMREPORT_SUMMARY, SpreadsheetApp.BandingTheme.CYAN)

  var sheet = SpreadsheetApp.getActive().getSheetByName(SHEET_PROBLEMREPORT_SUMMARY);
  sheet.setColumnWidth(1,200);
  sheet.setColumnWidth(2,800);
  sheet.getRange('A:A').setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP); // Set the wrapping to clip for column A
  sheet.getRange('B:B').setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP); // Set the wrapping to clip for column B
}
function FormatProblemReportInfo()
{
  Logger.log("Applying format to problem report info"); 
  //formatHeader(SHEET_PR_INFO,'A1:F1','#4a86e8');
  setAlternativeColorToRange(SHEET_PR_INFO, NAMED_RANGE_FILES_AND_FOLDERS, SpreadsheetApp.BandingTheme.CYAN)

  SetConditionalRule(SHEET_PR_INFO, "E:F",'=AND(ISNUMBER($E1), $E1=0, ISNUMBER($F1), $F1=0)','#b6d7a8'); // Green - OK  
  SetConditionalRule(SHEET_PR_INFO, "E:F",'=OR(AND(ISNUMBER($E1), $E1>0), AND(ISNUMBER($F1), $F1>0))','#ea9999'); // Red - Contains error/exception text
  SetConditionalRule(SHEET_PR_INFO, "E:F",'=OR(AND(ISNUMBER($E1), $E1=-2), AND(ISNUMBER($F1), $F1=-2))','#9900ff'); // Lila - Exception

  SetConditionalRule(SHEET_PR_INFO, NAMED_RANGE_FILES_AND_FOLDERS,'=$B1="File"','#cccccc'); // Gray
  SetConditionalRule(SHEET_PR_INFO, NAMED_RANGE_FILES_AND_FOLDERS,'=$B1="Folder"','#ffe599'); // Yellow

  var sheet = SpreadsheetApp.getActive().getSheetByName(SHEET_PR_INFO);
  sheet.setColumnWidth(1,300);
  sheet.setColumnWidth(2,100);
  sheet.setColumnWidth(3,300);
  sheet.setColumnWidth(4,100);
  sheet.setColumnWidth(5,100);
  sheet.setColumnWidth(6,100);
  sheet.getRange('A:A').setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP); // Set the wrapping to clip for column A
}
function FormatLogsAndMessages()
{   
  Logger.log("Starting loading of FormatLogsAndMessages"); 
  
  //formatHeader(SHEET_AUDITS_AND_MESSAGES,'A1:K1','#4a86e8');
  setAlternativeColorToRange(SHEET_AUDITS_AND_MESSAGES, NAMED_RANGE_AUDITSANDMESSAGES, SpreadsheetApp.BandingTheme.CYAN)

  var sheet = SpreadsheetApp.getActive().getSheetByName(SHEET_AUDITS_AND_MESSAGES);
  sheet.setColumnWidth(1,100);
  sheet.setColumnWidth(2,125);
  sheet.setColumnWidth(3,200);
  sheet.setColumnWidth(4,100);
  sheet.setColumnWidth(5,250);
  sheet.setColumnWidth(6,550);
  sheet.setColumnWidth(7,100);
  sheet.setColumnWidth(8,100);
  sheet.setColumnWidth(9,100);
  sheet.setColumnWidth(10,100);
  sheet.setColumnWidth(11,100);
  sheet.getRange('A:A').setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP); // Set the wrapping to clip for column A
  sheet.getRange('B:B').setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP); // Set the wrapping to clip for column A
  sheet.getRange('C:C').setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP); // Set the wrapping to clip for column A
  sheet.getRange('D:D').setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP); // Set the wrapping to clip for column A
  sheet.getRange('E:E').setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP); // Set the wrapping to clip for column A
  sheet.getRange('F:F').setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP); // Set the wrapping to clip for column A
  sheet.getRange('G:G').setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP); // Set the wrapping to clip for column A
  sheet.getRange('H:H').setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP); // Set the wrapping to clip for column A
  sheet.getRange('I:I').setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP); // Set the wrapping to clip for column A
  sheet.getRange('J:J').setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP); // Set the wrapping to clip for column A
  sheet.getRange('K:K').setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP); // Set the wrapping to clip for column A
  
  //SetConditionalRule(SHEET_AUDITS_AND_MESSAGES, NAMED_RANGE_AUDITSANDMESSAGES,'=$H1=1','#cccccc');
  SetConditionalRule(SHEET_AUDITS_AND_MESSAGES, NAMED_RANGE_AUDITSANDMESSAGES,'=$H1=2','#ffe599'); //Yellow
  SetConditionalRule(SHEET_AUDITS_AND_MESSAGES, NAMED_RANGE_AUDITSANDMESSAGES,'=$H1=3','#ea9999'); //Red
}

function deleteDataRangeContent(sheet_name, named_range)
{
  var sheet = SpreadsheetApp.getActive().getSheetByName(sheet_name);
  if(sheet)
  {
    var range = sheet.getRange(named_range);
    if(range)
    {
      var numRows = range.getNumRows();
      if(numRows>1)
      {
        var numCols = range.getNumColumns();
        // Exclude the first row
        var contentRange = range.offset(1, 0, numRows - 1, numCols);
        // Clear the content of the range
        contentRange.clearContent();
        CreateFilterAndDataRangeInSheet(sheet_name,named_range);
      }
    }
  }
}
function formatHeader(sheet_name, header_range, background_color) 
{
  var sheet = SpreadsheetApp.getActive().getSheetByName(sheet_name);
  var range = sheet.getRange(header_range);
  
  //var headerStyle = range.getFontWeight().setFontWeight('bold');
  //var headerSize = range.getFontSize().setFontSize(12);
  //var headerColor = range.getBackground().setBackground(background_color);
  
  //range.setFontWeight(headerStyle);
  //range.setFontSize(headerSize);
  range.setBackground(background_color);
}

function SetConditionalRule(sheet_name, named_range, formula, formulacolor)
{
  var sheet = SpreadsheetApp.getActive().getSheetByName(sheet_name);
  var range = sheet.getRange(named_range);
  //sheet.clearConditionalFormatRules();
  
  var ruleexist = false;
  var rules = sheet.getConditionalFormatRules();
  // Find the rule
  for (var i = rules.length - 1; i >= 0; i--) 
  {
    var rule = rules[i];
    var ranges = rule.getRanges();
    var ruleFormula = rule.getBooleanCondition().getCriteriaValues()[0]
    if (ranges.length === 1 && ranges[0].getA1Notation() === range.getA1Notation() && ruleFormula === formula) 
    {
      ruleexist = true;
    }
  }
  if(!ruleexist)
  {
    var rule = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied(formula)
    .setBackground(formulacolor)
    .setRanges([range])
    .build();
  
    rules.push(rule);
    sheet.setConditionalFormatRules(rules);
  }
}
function LoadProblemReportInformation()
{
  Logger.log("Starting loading of problem report for shared folder");
  var sheet = SpreadsheetApp.getActive().getSheetByName(SHEET_CONFIGURATION);
  var cell = sheet.getRange("B1");
  var sharedPath = cell.getValue();
  var folderId = extractFileIdFromPath(sharedPath);
  var folder = DriveApp.getFolderById(folderId);
  sheet.getRange("B2").setValue(folder.getName());
  
  ListAllFilesAndFoldersInProblemReport(folderId);
  CreateFilterAndDataRangeInSheet(SHEET_PR_INFO, NAMED_RANGE_FILES_AND_FOLDERS);
}
 
function FindFileInFolder(folderID, filename) {
  var folder = DriveApp.getFolderById(folderID);
  var files = folder.getFilesByName(filename);
  if (files.hasNext()) {
    var file = files.next();
    var fileId = file.getId();
    return fileId;
  } else {
    return null; // Return null instead of "File not found"
  }
}

function FormatUITracelog()
{
  //Date	Thread	Level	Logger	Message	Exception	Correlation Id [EOL]


  var sheet = SpreadsheetApp.getActive().getSheetByName(SHEET_UI_TRACE_LOG);
  sheet.setColumnWidth(1,150);
  sheet.setColumnWidth(2,200);
  sheet.setColumnWidth(3,150);
  sheet.setColumnWidth(4,200);
  sheet.setColumnWidth(5,600);
  sheet.setColumnWidth(6,200);
  sheet.setColumnWidth(7,200);

  sheet.getRange('A:A').setWrapStrategy(SpreadsheetApp.WrapStrategy.OVERFLOW); // Set the wrapping to clip for column A
  sheet.getRange('B:B').setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP); // Set the wrapping to clip for column B
  sheet.getRange('C:C').setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP); // Set the wrapping to clip for column B
  sheet.getRange('D:D').setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP); // Set the wrapping to clip for column B
  sheet.getRange('E:E').setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP); // Set the wrapping to clip for column B
  sheet.getRange('E:F').setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP); // Set the wrapping to clip for column B
  sheet.getRange('E:G').setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP); // Set the wrapping to clip for column B
  
  sheet.getRange('A:A').setNumberFormat('yyyy-MM-dd HH:mm:ss'); // Set the number format to date and time


  //formatHeader(SHEET_BACKEND_LOG,'A1:G1','#4a86e8');
  setAlternativeColorToRange(SHEET_UI_TRACE_LOG, NAMED_RANGE_UI_TRACE_LOG, SpreadsheetApp.BandingTheme.CYAN)
  
  //SetConditionalRule(SHEET_UI_TRACE_LOG, NAMED_RANGE_UI_TRACE_LOG,'=$C1="INFO"','#cccccc'); //Gray
  //SetConditionalRule(SHEET_UI_TRACE_LOG, NAMED_RANGE_UI_TRACE_LOG,'=$C1="DEBUG"','#cccccc'); //Gray
  SetConditionalRule(SHEET_UI_TRACE_LOG, NAMED_RANGE_UI_TRACE_LOG,'=$C1="WARN"','#f9cb9c'); // Orange
  SetConditionalRule(SHEET_UI_TRACE_LOG, NAMED_RANGE_UI_TRACE_LOG,'=$C1="ERROR"','#ea9999'); // Red
}
function FormatBackendlog()
{
  var sheet = SpreadsheetApp.getActive().getSheetByName(SHEET_BACKEND_LOG);
  sheet.setColumnWidth(1,180);
  sheet.setColumnWidth(2,200);
  sheet.setColumnWidth(3,100);
  sheet.setColumnWidth(4,200);
  sheet.setColumnWidth(5,700);
  sheet.setColumnWidth(6,200);
  sheet.setColumnWidth(7,200);
  sheet.getRange('A:A').setWrapStrategy(SpreadsheetApp.WrapStrategy.OVERFLOW); // Set the wrapping to clip for column A
  sheet.getRange('B:B').setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP); // Set the wrapping to clip for column B
  sheet.getRange('C:C').setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP); // Set the wrapping to clip for column B
  sheet.getRange('D:D').setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP); // Set the wrapping to clip for column B
  sheet.getRange('E:E').setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP); // Set the wrapping to clip for column B
  sheet.getRange('F:F').setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP); // Set the wrapping to clip for column B
  sheet.getRange('G:G').setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP); // Set the wrapping to clip for column B

  sheet.getRange('A:A').setNumberFormat('yyyy-MM-dd HH:mm:ss'); // Set the number format to date and time

  //formatHeader(SHEET_BACKEND_LOG,'A1:G1','#4a86e8');
  setAlternativeColorToRange(SHEET_BACKEND_LOG, NAMED_RANGE_BACKENDLOG, SpreadsheetApp.BandingTheme.CYAN)
  
  //SetConditionalRule(SHEET_BACKEND_LOG, NAMED_RANGE_BACKENDLOG,'=$C1="INFO"','#cccccc'); //Gray
  //SetConditionalRule(SHEET_BACKEND_LOG, NAMED_RANGE_BACKENDLOG,'=$C1="DEBUG"','#cccccc'); //Gray
  SetConditionalRule(SHEET_BACKEND_LOG, NAMED_RANGE_BACKENDLOG,'=$C1="WARN"','#f9cb9c'); // Orange
  SetConditionalRule(SHEET_BACKEND_LOG, NAMED_RANGE_BACKENDLOG,'=$C1="ERROR"','#ea9999'); // Red
}
function setAlternativeColorToRange(sheetname, namedrange, theme)
{
  var sheet = SpreadsheetApp.getActive().getSheetByName(sheetname);
  var range = sheet.getRange(namedrange); // Use the named range NAMED_RANGE_BACKENDLOG  
  range.getBandings().forEach(function (banding) {
    banding.remove();
  });
  range.applyRowBanding(theme); // Apply alternating colors to the range
}
function LoadResultsFromCSV(filepattern, sheetname,separator)
{   
  Logger.log("Starting loading results for pattern: "+filepattern+" on the sheet: "+sheetname);  
  var results = FindListOfFilesInPR(filepattern);
  if (results.length <= 0) 
  {
    Logger.log("No matching rows found. File not found:"+filepattern);    
    //SpreadsheetApp.getUi().alert("Error", "Find files in PR: no matching files found. File name:"+filepattern, SpreadsheetApp.getUi().ButtonSet.OK);
    return false;
  }
  var sheet = SpreadsheetApp.getActive().getSheetByName(sheetname);
  sheet.clear();
  for (var j = 0; j < results.length; j++) 
  {
    var append = true;
    if(j==0) append = false;
    var fileID = results[j][2];
    LoadResultsDataFromCVSFile(fileID, sheetname,append,separator)
  }    
  return true;
}
function LoadResultsDataFromCVSFile(fileID, sheetName, append,separator) {
  var file = DriveApp.getFileById(fileID);
  var fileName = file.getName(); // Get the name of the file
  Logger.log("Loading CSV results from cvs file: "+fileName);  
  
  var csvData = file.getBlob().getDataAsString();
  var parsedData = Utilities.parseCsv(csvData, separator); // Specify the semicolon delimiter
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  
  if (!append) {
    sheet.clear();
    sheet.getRange(1, 1, parsedData.length, parsedData[0].length).setValues(parsedData);

    // Add the column with header and file name
    sheet.insertColumnAfter(parsedData[0].length);
    sheet.getRange(1, parsedData[0].length + 1).setValue("Source file");
    sheet.getRange(2, parsedData[0].length + 1, parsedData.length-1).setValue(fileName);
  } else {
    parsedData.shift(); // Remove the first row
    sheet.getRange(sheet.getLastRow() + 1, 1, parsedData.length, parsedData[0].length).setValues(parsedData);

    // Add the file name in the last column for the appended data
    sheet.getRange(sheet.getLastRow() + 1 - parsedData.length, parsedData[0].length + 1, parsedData.length).setValue(fileName);
  }
}

/*
function LoadCVSFile(fileID, sheetName, append) {
  var file = DriveApp.getFileById(fileID);
  var csvData = file.getBlob().getDataAsString();
  var parsedData = Utilities.parseCsv(csvData, ';'); // Specify the semicolon delimiter
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  
  if (!append) {    
    sheet.clear();
    sheet.getRange(1, 1, parsedData.length, parsedData[0].length).setValues(parsedData);
  } else {
    parsedData.shift(); // Remove the first row
    sheet.getRange(sheet.getLastRow() + 1, 1, parsedData.length, parsedData[0].length).setValues(parsedData);
  }
}
*/

function LoadUITraceLogFromCSV()
{   
  Logger.log("Starting loading of the UI trace log");
  var filename = "UiTraces.log";
  var results = FindListOfFilesInPR(filename);
  if (results.length <= 0) 
  {
    Logger.log("No matching rows found. File not found:"+filename);    
    SpreadsheetApp.getUi().alert("Error", "Find files in PR: no matching files found. File name:"+filename, SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }

  var sheet = SpreadsheetApp.getActive().getSheetByName(SHEET_UI_TRACE_LOG);
  sheet.clear();
  for (var j = 0; j < results.length; j++) 
  {
    var filename = results[j][0];
    if(filename === "UiTraces.log")
    {
      var FILE_ID = results[j][2];    
      LoadDataFromCSV(FILE_ID, SHEET_UI_TRACE_LOG,null); 
      break;
    }
  }
  CreateFilterAndDataRangeInSheet(SHEET_UI_TRACE_LOG, NAMED_RANGE_UI_TRACE_LOG);  
}
function LoadBackendLogFromCSV()
{   
  Logger.log("Starting loading of the backendlog for the last file");
  var filename = "Roche.C4C.ServiceHosting.BackEnd.log.000.part";
  var results = FindListOfFilesInPR(filename);
  if (results.length <= 0) 
  {
    Logger.log("No matching rows found. File not found:"+filename);    
    SpreadsheetApp.getUi().alert("Error", "Find files in PR: no matching files found. File name:"+filename, SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }

  var sheet = SpreadsheetApp.getActive().getSheetByName(SHEET_BACKEND_LOG);
  sheet.clear();
  for (var j = 0; j < results.length; j++) 
  {
    var FILE_ID = results[j][2];
    LoadDataFromCSV(FILE_ID, SHEET_BACKEND_LOG,null)
  }
  CreateFilterAndDataRangeInSheet(SHEET_BACKEND_LOG, NAMED_RANGE_BACKENDLOG);  
}
function LoadLoggingExportTextFile(FILE_ID, SHEET_NAME) 
{  
  var sheet = SpreadsheetApp.getActive().getSheetByName(SHEET_NAME);
  
  var file = DriveApp.getFileById(FILE_ID);
  Logger.log("Loading Text file " + file.getName());
  
  var text = file.getBlob().getDataAsString();
  var lines = text.split("\n");
  var sheetData = [];
  
  sheetData.push(["Description", ""]);
  for (var i = 0; i < lines.length; i++) {
    var line = lines[i];
    var formattedLine = "'" + line; // Add the apostrophe to the beginning of each line
    sheetData.push([formattedLine, ""]);
  }
  
  // Clear existing data in the sheet  
  sheet.clear();  
  var lastRow = sheet.getLastRow();
  var range = sheet.getRange(lastRow + 1, 1, sheetData.length, 2);
  range.setValues(sheetData);
}


function LoadDataFromCSV(FILE_ID, SHEET_NAME, header)
{  
  Logger.log("Loading Backend LOG from CSV "+ FILE_ID); 
  var sheet = SpreadsheetApp.getActive().getSheetByName(SHEET_NAME);  
  
  var file = DriveApp.getFileById(FILE_ID);
  var csvData = file.getBlob().getDataAsString();
  //var csvData = file.getBlob().getAs('text/csv');   //var csvData = file.getBlob().getAs('text/plain');   //var csvData = Utilities.newBlob(file.getBlob().getBytes()).getDataAsString('UTF-8');
  var csvRows = csvData.split("\n");
  var sheetData = [];
  var maxColumns = 0;
  if(header)
  {
    var headerdata = header.split("<!>");  
    sheetData.push(headerdata);  
  }
  for (var i = 0; i < csvRows.length; i++) {
    var rowData = csvRows[i].split("<!>");    
    // Check if the number of columns in the current row is greater than the previous maximum
    if (rowData.length > maxColumns) {
      maxColumns = rowData.length;
    }    
    sheetData.push(rowData);
  }
  // Pad each row with empty values to match the maximum number of columns
  for (var i = 0; i < sheetData.length; i++) {
    var row = sheetData[i];
    while (row.length < maxColumns) {
      row.push("");
    }
  } 

  var lastRow = sheet.getLastRow();
  var range = sheet.getRange(lastRow + 1, 1, sheetData.length, maxColumns);
  range.setValues(sheetData);  
}
function ReadProblemReportSummary()
{
  Logger.log("Starting loading the problem report summay");
  var FileToFind = "ProblemReportSummary.txt";
  var results = FindListOfFilesInPR(FileToFind);
  if (results.length <= 0) 
  {
    Logger.log("No matching rows found. File not found: "+FileToFind);    
    SpreadsheetApp.getUi().alert("Error", "No matching rows found in the problem report list of files. File not found: "+FileToFind, SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }

  var FILE_ID = results[0][2];
  ReadAndLoadProblemReportSummary(FILE_ID, SHEET_PROBLEMREPORT_SUMMARY);
  CreateFilterAndDataRangeInSheet(SHEET_PROBLEMREPORT_SUMMARY, NAMED_RANGE_PROBLEMREPORT_SUMMARY); 
}

function ReadAndLoadProblemReportSummary(FILE_ID, SHEET_NAME) 
{  
  var file = DriveApp.getFileById(FILE_ID);
  Logger.log("Loading ProblemReportSummary XML File: " + file.getName());

  var xmlContent = file.getBlob().getDataAsString();
  var xmlDoc = XmlService.parse(xmlContent);
  var rootElement = xmlDoc.getRootElement();

  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  sheet.clearContents(); // Clears the content of the sheet

  // Create header with two columns named "Property" and "Value"
  sheet.getRange(1, 1, 1, 2).setValues([["Property", "Value"]]);

  var data = [];

  var properties = ["Actor", "AdditionalInformation", "ApplicationDescription", "ComponentType", "OSVersion", "SerialNumber", "SoftwareVersion", "SystemCode", "CreationDate", "IsComplete"];
  for (var i = 0; i < properties.length; i++) {
    var property = properties[i];
    var value = rootElement.getChildText(property);
    data.push([property, value]);
  }

  var containedFiles = rootElement.getChild("ContainedFiles").getChildren("FilePath");
  for (var i = 0; i < containedFiles.length; i++) {
    var filePath = containedFiles[i].getText();
    data.push(["ContainedFiles", filePath]);
  }

  sheet.getRange(2, 1, data.length, 2).setValues(data); // Start writing data from row 2
}
function LoadSettingsFile(settingsfilename, sheetname) {
  
  Logger.log("Starting loading of Settings File");  
  var results = FindListOfFilesInPR(settingsfilename);

  if (results.length > 0) {
    for (var j = 0; j < results.length; j++) {
      Logger.log(results[j]);
    }    
    var fileId = results[0][2];   
    Logger.log("Loading settings file: "+settingsfilename+"from file ID: "+ fileId);    
    LoadLoggingExportTextFile(fileId,sheetname);
  } else {
    Logger.log("No matching rows found. File not found: "+settingsfilename);    
    SpreadsheetApp.getUi().alert("Error", "Cannot load settings, file not found: "+settingsfilename, SpreadsheetApp.getUi().ButtonSet.OK);
  }  
}
function LoadServiceHostingComitServicesLog()
{
  Logger.log("Starting loading of LoadServiceHostingComitServicesLog");  
  var results = FindListOfFilesInPR("ServiceHosting.ComitServices");
  
  if (results.length > 0) {
    for (var j = 0; j < results.length; j++) {
      Logger.log(results[j]);
    }    
    var FILE_ID = results[0][2];   
    Logger.log("Loading ServiceHosting.ComitServices logs from file: "+ FILE_ID);
    LoadLoggingExportTextFile(FILE_ID,SHEET_SERVICE_HOSTING_COMITSERVICES_LOG);
  } else {
    Logger.log("No matching rows found. File not found: ServiceHosting.ComitServices");    
    SpreadsheetApp.getUi().alert("Error", "File not found: ServiceHosting.ComitServices", SpreadsheetApp.getUi().ButtonSet.OK);
  }  
}
function LoadSystemServiceHostComitServicesLog()
{
  Logger.log("Starting loading of LoadSystemServiceHostComitServicesLog");  
  var results = FindListOfFilesInPR("SystemServiceHost.ComitServices");
  
  if (results.length > 0) {
    for (var j = 0; j < results.length; j++) {
      Logger.log(results[j]);
    }    
    var FILE_ID = results[0][2];   
    Logger.log("Loading SystemServiceHost.ComitServices logs from file: "+ FILE_ID);
    LoadLoggingExportTextFile(FILE_ID,SHEET_SYSTEM_SERVICEHOST_COMITSERVICES_LOG);
  } else {
    Logger.log("No matching rows found. File not found: SystemServiceHost.ComitServices");    
    SpreadsheetApp.getUi().alert("Error", "File not found: SystemServiceHost.ComitServices", SpreadsheetApp.getUi().ButtonSet.OK);
  }  
}
function LoadCalculationLog()
{
  Logger.log("Starting loading of LoadCalculationLog");  
  var results = FindListOfFilesInPR("calculation.log");
  
  if (results.length > 0) {
    for (var j = 0; j < results.length; j++) {
      Logger.log(results[j]);
    }    
    var FILE_ID = results[0][2];   
    Logger.log("Loading calculation logs from file: "+ FILE_ID);
    LoadLoggingExportTextFile(FILE_ID,SHEET_CALCULATION_LOG);
  } else {
    Logger.log("No matching rows found. File not found: calculation.log");    
    SpreadsheetApp.getUi().alert("Error", "File not found: calculation.log", SpreadsheetApp.getUi().ButtonSet.OK);
  }  
}
function FormatCalculationLog()
{
  setAlternativeColorToRange(SHEET_CALCULATION_LOG, NAMED_RANGE_CALCULATION_LOG, SpreadsheetApp.BandingTheme.CYAN)

  var sheet = SpreadsheetApp.getActive().getSheetByName(SHEET_CALCULATION_LOG);
  sheet.setColumnWidth(1,800);  
  sheet.getRange('A:A').setWrapStrategy(SpreadsheetApp.WrapStrategy.OVERFLOW); // Set the wrapping to clip for column A
}

function LoadOracleDBLog()
{
  Logger.log("Starting loading of LoadOracleDBLog");  
  var results = FindListOfFilesInPR("alert_c4cdb.log");
  
  if (results.length > 0) {
    for (var j = 0; j < results.length; j++) {
      Logger.log(results[j]);
    }    
    var FILE_ID = results[0][2];   
    Logger.log("Loading oracle db logs from file: "+ FILE_ID);
    LoadLoggingExportTextFile(FILE_ID,SHEET_ORACLE_LOG);
  } else {
    Logger.log("No matching rows found. File not found: alert_c4cdb.log");    
    SpreadsheetApp.getUi().alert("Error", "File not found: alert_c4cdb.log", SpreadsheetApp.getUi().ButtonSet.OK);
  }  
}

function LoadAuditsAndMessages()
{
  Logger.log("Starting loading of LoadAuditsAndMessages");  
  var results = FindListOfFilesInPR("LoggingExport.xml");  
  
  if (results.length > 0) {
    for (var j = 0; j < results.length; j++) {
      Logger.log(results[j]);
    }    
    var FILE_ID = results[0][2];   
    Logger.log("Loading audis and messages from file: "+ FILE_ID);                 
    LoadLoggingExportXMLFile(FILE_ID,SHEET_AUDITS_AND_MESSAGES,NAMED_RANGE_AUDITSANDMESSAGES);     
  } else {
    Logger.log("No matching rows found. File not found: Roche.C4C.ServiceHosting.BackEnd.log");    
    SpreadsheetApp.getUi().alert("Error", "File not found: Roche.C4C.ServiceHosting.BackEnd.log", SpreadsheetApp.getUi().ButtonSet.OK);
  }  
}


function ListAllFilesAndFoldersInProblemReport(FOLDER_ID) {
  var folder = DriveApp.getFolderById(FOLDER_ID);
  Logger.log("Listing all files and folders of the problem report in folder: " + folder.getName() + " with ID: " + FOLDER_ID);

  var fileData = [];
  var folderData = [];

  fileData.push(["Name", "Type", "File ID", "Index", "#Errors", "#Exceptions"]);
  fileData.push([folder.getName(), "Folder", folder.getId(), 1, "n/a", "n/a"]);

  ListFilesAndFoldersRecursive(FOLDER_ID, fileData);

  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spreadsheet.getSheetByName(SHEET_PR_INFO);
  sheet.clear(); // Clear existing data in the sheet

  // Set the file data
  sheet.getRange(1, 1, fileData.length, fileData[0].length).setValues(fileData);  

   // Add the formula for hyperlinks in the "File ID" column
  var formulas = [];
  for (var i = 1; i < fileData.length; i++) {
    if(fileData[i][1]==="File")
      formulas.push(["=HYPERLINK(\"https://drive.google.com/uc?export=view&id=" + fileData[i][2] + "\", \"" + fileData[i][2] + "\")"]);
    else
      formulas.push(["=HYPERLINK(\"drive.google.com/drive/folders/" + fileData[i][2] + "\", \"" + fileData[i][2] + "\")"]);      
  }
  sheet.getRange(2, 3, formulas.length, 1).setFormulas(formulas);
  
}

function ListFilesAndFoldersRecursive(folderId, fileData) {
  var folder = DriveApp.getFolderById(folderId);
  var files = folder.getFiles();
  var folders = folder.getFolders();

  while (files.hasNext()) {
    var file = files.next();
    fileData.push([file.getName(), "File", file.getId(), fileData.length , "n/a", "n/a"]);
  }

  while (folders.hasNext()) {
    var subFolder = folders.next();
    fileData.push([subFolder.getName(), "Folder", subFolder.getId(), fileData.length + 1, "n/a", "n/a"]);
    ListFilesAndFoldersRecursive(subFolder.getId(), fileData, fileData);
  }
}
function extractFileIdFromPath(url) 
{ 
  return url.match(/[-\w]{25,}/); 
}
function LoadLoggingExportXMLFile(FILE_ID, SHEET_NAME, NAMED_RANGE) 
{  
  var file = DriveApp.getFileById(FILE_ID);
  Logger.log("Loading Logging XML File: " + file.getName());

  var xmlContent = file.getBlob().getDataAsString();
  var xmlDoc = XmlService.parse(xmlContent);
  var rootElement = xmlDoc.getRootElement();  
  var logEvents = rootElement.getChild("LogEvents").getChildren("LogEvent");  
  var data = [];
  
  // Add column headers
  data.push(["ActorId", "ModuleId", "Timestamp", "EventCode", "Title", "EventText", "IsAuditTrail", "Severity", "LogEventDefinitionId"]);  
  // Parse XML and add data to the table
  logEvents.forEach(function(logEvent) {
    var actorId = logEvent.getChildText("ActorId");
    var moduleId = logEvent.getChildText("ModuleId");
    var timestamp = logEvent.getChildText("Timestamp");
    var eventCode = logEvent.getChildText("EventCode");
    var title = logEvent.getChildText("Title");
    var eventText = logEvent.getChildText("EventText");
    var isAuditTrail = logEvent.getChildText("IsAuditTrail");
    var severity = logEvent.getChildText("Severity");
    var logEventDefinitionId = logEvent.getChildText("LogEventDefinitionId");
    
    data.push([actorId, moduleId, timestamp, eventCode, title, eventText, isAuditTrail, severity, logEventDefinitionId]);
  });
  
  // Clear existing data in the sheet
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spreadsheet.getSheetByName(SHEET_NAME);
  sheet.clear();  
  // Set the values in the sheet
  sheet.getRange(1, 1, data.length, data[0].length).setValues(data);
  CreateFilterAndDataRangeInSheet(SHEET_NAME, NAMED_RANGE);  
}

function CreateFilterAndDataRangeInSheet(SHEET_NAME, NAMED_RANGE)
{
  Logger.log("Create filter in sheet: "+SHEET_NAME+" for named range: "+NAMED_RANGE);
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spreadsheet.getSheetByName(SHEET_NAME);
  var dataRange = sheet.getDataRange();
  var numRows = dataRange.getNumRows();
  var numColumns = dataRange.getNumColumns();  
  var destrange = sheet.getRange(1, 1, numRows, numColumns);

  var filter = sheet.getFilter();
  if(filter!=null)
  {
    filter.remove();
  }
  destrange.createFilter();
  SpreadsheetApp.getActive().setNamedRange(NAMED_RANGE, destrange);
}
function AnalyzeFileContent()
{
  Logger.log("Start Analyze problem report content");
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  ss.toast("[1/5] Analyzing content for files of type: Roche.C4C.ProblemReporting", "Analyzing file content", 2);
  AnalyzeFileContentWithErrors("Roche.C4C.ProblemReporting");

  ss.toast("[1/5] Analyzing content for files of type: Roche.C4C.ProblemReporting", "Analyzing file content", 2);
  AnalyzeFileContentWithErrors("ServiceHosting.ComitServices");

  ss.toast("[1/5] Analyzing content for files of type: Roche.C4C.ProblemReporting", "Analyzing file content", 2);
  AnalyzeFileContentWithErrors("SystemServiceHost.ComitServices");

  ss.toast("[1/5] Analyzing content for files of type: Roche.C4C.ProblemReporting", "Analyzing file content", 2);
  AnalyzeFileContentWithErrors("WindowsEvents");  

  ss.toast("[1/5] Analyzing content for files of type: Roche.C4C.ProblemReporting", "Analyzing file content", 2);
  AnalyzeFileContentWithErrors("Roche.C4C.ServiceHosting.BackEnd");
  
  ss.toast("[5/5] Analisis finished", "Analyzing file content", 1);

  AppendExecutionLogToLogsSheet();
}
function AnalyzeFileContentWithErrors(filetofind)
{
   var ss = SpreadsheetApp.getActiveSpreadsheet();
   //ss.toast("Analyze file content for file {}", "Analyzing file content.", -1);
  var sheet = SpreadsheetApp.getActive().getSheetByName(SHEET_PR_INFO);
  
  var results = FindListOfFilesInPR(filetofind);
  if (results.length <= 0)   {
    Logger.log("No matching rows found. File not found: "+filetofind);            
  }  
  for (var i= 0; i < results.length; i++)
  { 
    var fileID = results[i][2];
    var rowindex = results[i][3];
    var file = DriveApp.getFileById(fileID);

    ss.toast("Progress: "+i+"/"+results.length+". Analyzing file: "+file.getName(), "Analyzing file content", -1);
    var numerrors     = CountOfWordOccurrencesInfile(file, "error");
    var numexceptions = CountOfWordOccurrencesInfile(file, "exception");
    sheet.getRange(rowindex,5).setValue(numerrors);
    sheet.getRange(rowindex,6).setValue(numexceptions);
  }
}
function FindListOfFilesInPR(filenametosearch)
{
  var SHEET_NAME = SHEET_PR_INFO;
  var NAMED_RANGE = NAMED_RANGE_FILES_AND_FOLDERS;

  Logger.log("Looking for file in problem report: "+filenametosearch);
  var sheet = SpreadsheetApp.getActive().getSheetByName(SHEET_NAME);
  var data = sheet.getRange(NAMED_RANGE).getValues();
  var results = [];  
   for (var i = 0; i < data.length; i++) {
    if (data[i][0].indexOf(filenametosearch) !== -1) {
      results.push(data[i]);
    }
  }
  if (results.length > 0) {
    results.sort(function(a, b) {
      return a[0].localeCompare(b[0]);
    });
  }
  return results;
}
function AppendExecutionLogToLogsSheet() 
{
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var logsSheet = spreadsheet.getSheetByName(SHEET_EXECUTION_LOG);
  if (!logsSheet) {  // If the "LOGS" sheet doesn't exist, create a new sheet
    logsSheet = spreadsheet.insertSheet(SHEET_EXECUTION_LOG);
  }  
  // Get the existing data in the "LOGS" sheet and adds a header if it is empty
  var existingData = logsSheet.getDataRange().getValues();    
  if (existingData.length < 2) {
    // Add header with three rows
    var header = [
      ["Date", "Type", "Information"]      
    ];
    logsSheet.getRange(1, 1, 1, 3).setValues(header);
  }
  var existingData = logsSheet.getDataRange().getValues();

  // Get the execution log
  var executionLog = Logger.getLog();  

  /* Version 1: one block 
  // Append the execution log to the existing data
  //existingData.push([executionLog]);  
  */
  
  /* 
  Version 2: one n lines  Append each line as a separate row in the existingData array
  var logLines = executionLog.split("\n"); // Split the executionLog into an array of lines
  for (var i = 0; i < logLines.length; i++) {
    existingData.push([logLines[i]]);
  }
  */
  
  // Version 3: N lines and 3 columns
  // Append each line as a separate row in the existingData array, splitting the line into three columns
  var logLines = executionLog.split("\n"); // Split the executionLog into an array of lines
  var formattedData = []; // Create a new array to store the formatted log information
  for (var i = 0; i < logLines.length; i++) {
    var line = logLines[i];
    var infotype = "INFO:";
    var infoIndex = line.indexOf(infotype);
    var infoLength = infotype.length;
    if (infoIndex !== -1) {
      var dateraw = line.substring(0, infoIndex).trim();
      var substringToRemove = " CEST 2024";  
      dateraw = dateraw.replace(substringToRemove, "");

      var date = new Date(dateraw);
      if (isNaN(date)) {
        date = dateraw;
      }
      var info = line.substring(infoIndex,infoIndex+infoLength-1).trim();
      var description = line.substring(infoIndex+infoLength).trim();
      
      formattedData.push([date, info, description]);
    }
  }
  
  // Concatenate existingData and formattedData arrays
  var newData = existingData.concat(formattedData);

  // Write the updated data to the "LOGS" sheet
  logsSheet.clear();
  logsSheet.getRange(1, 1, newData.length, newData[0].length).setValues(newData);
  
   Logger.clear();

   CreateFilterAndDataRangeInSheet(SHEET_EXECUTION_LOG, NAMED_RANGE_EXECUTION_LOG);
   FormatExecutionLog();
}
function FormatExecutionLog()
{
  var sheet = SpreadsheetApp.getActive().getSheetByName(SHEET_EXECUTION_LOG);
  sheet.setColumnWidth(1,200);
  sheet.setColumnWidth(2,100);
  sheet.setColumnWidth(3,800);
  sheet.getRange('A:A').setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP); // Set the wrapping to clip for column A
  sheet.getRange('B:B').setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP); // Set the wrapping to clip for column B
  sheet.getRange('C:C').setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP); // Set the wrapping to clip for column B
  setAlternativeColorToRange(SHEET_EXECUTION_LOG, NAMED_RANGE_EXECUTION_LOG, SpreadsheetApp.BandingTheme.CYAN)
  
  SetConditionalRule(SHEET_EXECUTION_LOG, "A:C",'=SEARCH("---", $C1)>=0','#fff2cc'); // Yellow. (#b6d7a8 - Green)

  sheet.getRange('A:A').setNumberFormat('yyyy-MM-dd HH:mm:ss'); // Set the number format to date and time
  var sort=false;
  if(sort)
  {
    var range = sheet.getRange(NAMED_RANGE_EXECUTION_LOG); // Use the named range NAMED_RANGE_BACKENDLOG  
    range.sort({column: 1, ascending: false, sortRange: range});
  }
}
function CreateSheet(SHEET_NAME) 
{
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  // Check if the sheet already exists
  var sheet = spreadsheet.getSheetByName(SHEET_NAME);
  if (!sheet) 
  {// If the sheet doesn't exist, create a new sheet
    spreadsheet.insertSheet(SHEET_NAME);    
    Logger.log("Sheet created: " + SHEET_NAME);
  } 
  else 
  {
    Logger.log("Sheet already exists: " + SHEET_NAME);
  }  
}

function convertToUTF8(file_id) {
  var file = DriveApp.getFileById(file_id); // Replace YOUR_XML_FILE_ID with the ID of your XML file
  var content = file.getBlob().getDataAsString();
  var utf8Content = Utilities.newBlob('').setDataFromString(content, 'UTF-8');

  DriveApp.createFile(utf8Content);
}
function AddFormulaColumnsToAuditsAndMessages()
{
  Logger.log("Starting loading of formulas to LoadAuditsAndMessages");  
  var columnname = "Instrument Status";
  var formula = '=IF(IF(IFERROR(FIND("System status changed to:",F2),-1)<=0,"",RIGHT(F2,LEN(F2)-LEN ("System status changed to:")))<>"", IF(IFERROR(FIND("System status changed to:",F2),-1)<=0,"",RIGHT(F2,LEN(F2)-LEN ("System status changed to:"))), IF(ROW()=2,"Unknown",INDIRECT("J" & ROW()-1)))';
  
  AddColumnWithFormulaToNamedRange(SHEET_AUDITS_AND_MESSAGES,NAMED_RANGE_AUDITSANDMESSAGES, columnname, formula);

   columnname = "LABAC Status";
   formula = '=IF(IF(IFERROR(FIND("Connection module is connected. Operating mode:",F2),-1)<=0,"",RIGHT(F2,LEN(F2)-LEN ("Connection module is connected. Operating mode:")))<>"", IF(IFERROR(FIND("Connection module is connected. Operating mode:",F2),-1)<=0,"",RIGHT(F2,LEN(F2)-LEN ("Connection module is connected. Operating mode:"))), IF(ROW()=2,"Unknown",INDIRECT("K" & ROW()-1)))';
  
  AddColumnWithFormulaToNamedRange(SHEET_AUDITS_AND_MESSAGES,NAMED_RANGE_AUDITSANDMESSAGES, columnname, formula);
  

  CreateFilterAndDataRangeInSheet(SHEET_AUDITS_AND_MESSAGES, NAMED_RANGE_AUDITSANDMESSAGES);
}
function SortNamedRangeAscending(sheet_name, named_range, columnnumber) {

   Logger.log("SortNamedRangeAscending Sheet: "+sheet_name+" For named range: "+named_range+" and column: "+columnnumber);  

  //var sheet = SpreadsheetApp.getActive().getSheetByName(sheet_name);  
  //var namedRange = sheet.getRange(named_range);
  //namedRange.sort({column: columnnumber, ascending: true});

  var sheet = SpreadsheetApp.getActive().getSheetByName(sheet_name);  
  var dataRange = sheet.getRange(named_range);
  var numRows = dataRange.getNumRows();
  var numColumns = dataRange.getNumColumns();
  
  var headerRange = sheet.getRange(1, 1, 1, numColumns); // Get the header row range
  var dataRowsRange = sheet.getRange(2, 1, numRows - 1, numColumns); // Get the data rows range
  
  // Sort the data rows range by column 3
  dataRowsRange.sort({column: 3, ascending: true});
  
  // Reassemble the sorted data range, including the header row
  var sortedRange = headerRange.getValues().concat(dataRowsRange.getValues());
  
  // Update the sheet with the sorted range
  sheet.getRange(1, 1, numRows, numColumns).setValues(sortedRange);

}

function AddColumnWithFormulaToNamedRange(sheet_name, named_range, columnname, formula) {
  var sheet = SpreadsheetApp.getActive().getSheetByName(sheet_name);
  var namedRange = sheet.getRange(named_range);
  var numRows = namedRange.getNumRows();
  
  // Insert a new column to the right of the named range
  var statusColumn = namedRange.getLastColumn() + 1;
  sheet.insertColumnAfter(statusColumn);
  
  // Set the header for the new column
  sheet.getRange(1, statusColumn).setValue(columnname);
  
  // Populate the new column with the formula
  var numRows = namedRange.getNumRows();
  
  var range = sheet.getRange(2, statusColumn, numRows - 1, 1);
  range.setFormula(formula);


  // Add the new column to the named range
  var newRange = sheet.getRange(1, 1, numRows, statusColumn);
  var namedRanges = sheet.getNamedRanges();
  for (var i = 0; i < namedRanges.length; i++) {
    if (namedRanges[i].getName() === named_range) {
      namedRanges[i].setRange(newRange);
      break;
    }
  }
}
 function CountOfWordOccurrencesInfile(file, WORD) 
 {  
  var occurrences = -1;
  try 
  {
    var textContent = file.getBlob().getDataAsString();
    var regex = new RegExp("\\b" + WORD + "\\b", "gi"); //new RegExp(WORD, "gi")
    var matches = textContent.match(regex);
    var occurrences = matches ? matches.length : 0;
    Logger.log("Word ["+WORD+ "] appear :" + occurrences+" times in file: "+file.getName());      
  } catch (error) {
    occurrences = -2;
    Logger.log("Unable to find Word ["+WORD+ "] in file: "+file.getName()+"Reason: "+error);  
  }
  return occurrences;
}
function deleteAllSheets() 
{
  Utilities.sleep(100);
  var starttime = logMethodStart("deleteAllSheets");
  
  var sheetNames = [SHEET_EXECUTION_LOG, SHEET_PR_INFO,SHEET_PROBLEMREPORT_SUMMARY,SHEET_BACKEND_LOG, SHEET_UI_TRACE_LOG,SHEET_AUDITS_AND_MESSAGES,SHEET_ORACLE_LOG,SHEET_SOFTWARESETTINGS,SHEET_LIS_SETTINGS,SHEET_SYSTEM_CONFIG_SETTINGS, SHEET_TNS_CONFIG_SETTINGS,SHEET_SERVICE_HOSTING_COMITSERVICES_LOG, SHEET_SYSTEM_SERVICEHOST_COMITSERVICES_LOG,SHEET_CALCULATION_LOG];  
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  for (var i = 0; i < sheetNames.length; i++) 
  {   
    var SHEET_NAME = sheetNames[i];
    var sheet = spreadsheet.getSheetByName(SHEET_NAME);        
    if (sheet!=null) 
    {
      spreadsheet.deleteSheet(sheet);      
      Logger.log("Sheet deleted: " + SHEET_NAME);
    }
    Utilities.sleep(50);
  }
  logMethodEnd(starttime,"deleteAllSheets");
  
}
function createAllSheets() {
   
  Utilities.sleep(100);
  var starttime = logMethodStart("createAllSheets");
  
  var sheetNames = [SHEET_CONFIGURATION, SHEET_EXECUTION_LOG, SHEET_PR_INFO,SHEET_PROBLEMREPORT_SUMMARY,SHEET_BACKEND_LOG, SHEET_UI_TRACE_LOG,SHEET_AUDITS_AND_MESSAGES,SHEET_ORACLE_LOG,SHEET_SOFTWARESETTINGS,SHEET_LIS_SETTINGS,SHEET_SYSTEM_CONFIG_SETTINGS, SHEET_TNS_CONFIG_SETTINGS,SHEET_SERVICE_HOSTING_COMITSERVICES_LOG, SHEET_SYSTEM_SERVICEHOST_COMITSERVICES_LOG,SHEET_CALCULATION_LOG];  
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  for (var i = 0; i < sheetNames.length; i++) {
    // Check if the sheet already exists
    var SHEET_NAME = sheetNames[i];
    var sheet = spreadsheet.getSheetByName(SHEET_NAME);        
    if (!sheet) 
    {// If the sheet doesn't exist, create a new sheet
      spreadsheet.insertSheet(SHEET_NAME);      
      Logger.log("Sheet created: " + SHEET_NAME);
    } 
    else 
    {
      Logger.log("Sheet already exists: " + SHEET_NAME);
    }  
     Utilities.sleep(50);
  }
  logMethodEnd(starttime,"createAllSheets");
}
function sortSheets() {
 
  Utilities.sleep(100);
  var starttime = logMethodStart("sortSheets");
  
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheetNames = [SHEET_CONFIGURATION, SHEET_EXECUTION_LOG, SHEET_PR_INFO,SHEET_PROBLEMREPORT_SUMMARY,SHEET_BACKEND_LOG, SHEET_UI_TRACE_LOG,SHEET_AUDITS_AND_MESSAGES,SHEET_ORACLE_LOG,SHEET_SOFTWARESETTINGS,SHEET_LIS_SETTINGS,SHEET_SYSTEM_CONFIG_SETTINGS, SHEET_TNS_CONFIG_SETTINGS,SHEET_SERVICE_HOSTING_COMITSERVICES_LOG, SHEET_SYSTEM_SERVICEHOST_COMITSERVICES_LOG,SHEET_CALCULATION_LOG, SHEET_CYCLE_JOB_TRACE,SHEET_SAMPLE_TEST_RESULTS_COAG,SHEET_CALIBRATION_TEST_RESULTS_COAG,SHEET_SAMPLE_TEST_RESULTS_CC,SHEET_CALIBRATION_TEST_RESULTS_CC ];

  for (var i = 0; i < sheetNames.length; i++) {
    var sheet = spreadsheet.getSheetByName(sheetNames[i]);    
     if (sheet) {
      spreadsheet.setActiveSheet(sheet);
      spreadsheet.moveActiveSheet(i + 1);
     }
  }
  logMethodEnd(starttime,"sortSheets");
}

/*
function SelectPRFolder() {
  var ui = SpreadsheetApp.getUi();
  var result = ui.prompt('Select an uncompressed Problem Report Folder', 'Please enter the folder ID:', ui.ButtonSet.OK_CANCEL);
  
  if (result.getSelectedButton() == ui.Button.OK) {
    var folderId = result.getResponseText();
    ListAllFilesAndFoldersInProblemReport(folderId);
  }
}
function SelectPRFolder() {
  var htmlOutput = HtmlService.createHtmlOutputFromFile('folderPicker')
    .setWidth(400)
    .setHeight(300);
  
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Select PR Folder');
}

function LoadDataFromCSVFilteredAndFast(FILE_ID, SHEET_NAME, NAMED_RANGE)
{
   var sheet = SpreadsheetApp.getActive().getSheetByName(SHEET_NAME);
  sheet.clear();
  
  var file = DriveApp.getFileById(FILE_ID);
  var csvData = file.getBlob().getDataAsString();
  var csvRows = csvData.split("\n");

  var data = csvRows.map(function(row) {
    var rowData = row.split("<!>");
    if (rowData.length === 7) { // Example condition for valid rows
      return rowData;
    } else {
      return null; // Skip invalid rows
    }
  }).filter(function(row) {
    return row !== null; // Remove null values (skipped rows)
  });


  sheet.getRange(1, 1, data.length, data[0].length).setValues(data);  
  var destrange = sheet.getDataRange();

  var filter = sheet.getFilter();
  if(filter!=null)
  {
    filter.remove();
  }
  destrange.createFilter();
  SpreadsheetApp.getActive().setNamedRange(NAMED_RANGE, destrange);
}
function LoadDataFromCSVWithChar(FILE_ID, SHEET_NAME, NAMED_RANGE)
{
  var file = DriveApp.getFileById(FILE_ID);
  var csvraw = file.getBlob().getDataAsString();
  var csvdata = Utilities.parseCsv(csvraw,"!");
  var sheet = SpreadsheetApp.getActive().getSheetByName(SHEET_NAME);
  sheet.clear();
  var destrange = sheet.getRange(1, 1, csvdata.length, csvdata[0].length);
  destrange.setValues(csvdata);
 
  var filter = sheet.getFilter();
  if(filter!=null)
  {
    filter.remove();
  }
  destrange.createFilter();
  SpreadsheetApp.getActive().setNamedRange(NAMED_RANGE, destrange);
}
*/

//OLD VERSIONS
/*
function ListAllFilesAndFoldersInProblemReport(FOLDER_ID) 
{ 
  var folder = DriveApp.getFolderById(FOLDER_ID);
  Logger.log("Listing all files and folders of the problem report in folder: "+folder.getName()+"with ID: "+ FOLDER_ID);

  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spreadsheet.getSheetByName(SHEET_PR_INFO);
  sheet.clear();// Clear existing data in the sheet

  // Set headers in the sheet
  sheet.getRange(1, 1).setValue("Name");
  sheet.getRange(1, 2).setValue("Type");
  sheet.getRange(1, 3).setValue("File ID");
  sheet.getRange(1, 4).setValue("Index");
  sheet.getRange(1, 5).setValue("#Errors");
  sheet.getRange(1, 6).setValue("#Exceptions");
  //Set the root folder:
  sheet.getRange(2, 1).setValue(folder.getName());
  sheet.getRange(2, 2).setValue("Folder");
  sheet.getRange(2, 3).setValue(folder.getId()).setFormula('=HYPERLINK("https://drive.google.com/uc?export=view&id=' + folder.getId() + '","' + folder.getId() + '")');
  sheet.getRange(2, 4).setValue(row);
  sheet.getRange(2, 5).setValue("n/a");
  sheet.getRange(2, 6).setValue("n/a");

  var row = 3; // Start from row 3 and start the recursive listing  
  
  ListFilesAndFoldersRecursive(FOLDER_ID, sheet, row);
}
function ListFilesAndFoldersRecursive(folderId, sheet, row) {
  var folder = DriveApp.getFolderById(folderId);
  var files = folder.getFiles();
  var folders = folder.getFolders();

  while (files.hasNext()) {
    var file = files.next();
    sheet.getRange(row, 1).setValue(file.getName());
    sheet.getRange(row, 2).setValue("File");
    sheet.getRange(row, 3).setValue(file.getId()).setFormula('=HYPERLINK("https://drive.google.com/uc?export=view&id=' + file.getId() + '","' + file.getId() + '")');
    sheet.getRange(row, 4).setValue(row);
    sheet.getRange(row, 5).setValue("n/a");
    sheet.getRange(row, 6).setValue("n/a");
    row++;
  }

  while (folders.hasNext()) {
    var subFolder = folders.next();
    sheet.getRange(row, 1).setValue(subFolder.getName());
    sheet.getRange(row, 2).setValue("Folder");
    sheet.getRange(row, 3).setValue(subFolder.getId()).setFormula('=HYPERLINK("https://drive.google.com/uc?export=view&id=' + subFolder.getId() + '","' + subFolder.getId() + '")');
    sheet.getRange(row, 4).setValue(row);
    sheet.getRange(row, 5).setValue("n/a");
    sheet.getRange(row, 6).setValue("n/a");
    row++;

    // Recursively call the function for subfolders
    row = ListFilesAndFoldersRecursive(subFolder.getId(), sheet, row);
  }
  return row;
}
*/
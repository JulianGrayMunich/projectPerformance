using databaseAPI;

using EASendMail;

using GNAchartingtools;

using GNAgeneraltools;

using GNAspreadsheettools;

using Microsoft.VisualBasic;

using OfficeOpenXml;

using System;
using System.Configuration;
using System.Data;
using System.Data.SqlTypes;
using System.Globalization;
using System.IO.Enumeration;
using System.Net.NetworkInformation;
using System.Security;
using System.Security.Cryptography;

using static OfficeOpenXml.ExcelErrorValue;

namespace projectPerformance
{






    class Program
    {
        static void Main(string[] args)
        {


#pragma warning disable CS0162
#pragma warning disable CS0164
#pragma warning disable CS8600
#pragma warning disable CS8601
#pragma warning disable CS8602
#pragma warning disable CS8604
#pragma warning disable CA1416




            gnaTools gnaT = new gnaTools();
            dbAPI gnaDBAPI = new dbAPI();
            spreadsheetAPI gnaSpreadsheetAPI = new spreadsheetAPI();
            GNAchartingAPI chartingAPI = new GNAchartingAPI();

            //==== System config variables
            string strSoftware = ConfigurationManager.AppSettings["Software"];
            string strDBconnection = ConfigurationManager.ConnectionStrings["DBconnectionString"].ConnectionString;
            string strProjectTitle = ConfigurationManager.AppSettings["ProjectTitle"];
            string strContractTitle = ConfigurationManager.AppSettings["ContractTitle"];
            string strReportType = ConfigurationManager.AppSettings["ReportType"];

            string strExcelPath = ConfigurationManager.AppSettings["ExcelPath"];
            string strExcelFile = ConfigurationManager.AppSettings["ExcelFile"];

            string strCheckWorksheetsExist = ConfigurationManager.AppSettings["checkWorksheetsExist"];

            string strReferenceWorksheet = ConfigurationManager.AppSettings["ReferenceWorksheet"];
            string strSurveyWorksheet = ConfigurationManager.AppSettings["SurveyWorksheet"];
            string strHourlyPerformanceWorksheet = ConfigurationManager.AppSettings["HourlyPerformanceWorksheet"];
            string strDailyPerformanceWorksheet = ConfigurationManager.AppSettings["DailyPerformanceWorksheet"];

            string strPerformanceWorksheet = ConfigurationManager.AppSettings["DailyPerformanceWorksheet"];

            string strFirstDataRow = ConfigurationManager.AppSettings["FirstDataRow"];
            string strFirstDataCol = ConfigurationManager.AppSettings["FirstDataCol"];
            string strFirstOutputRow = ConfigurationManager.AppSettings["FirstOutputRow"];

            int iFirstOutputRow = Convert.ToInt16(strFirstOutputRow);
            int iFirstDataCol = Convert.ToInt16(strFirstDataCol);


            string strSendEmails = ConfigurationManager.AppSettings["SendEmails"];
            string strAddAttachment = ConfigurationManager.AppSettings["AddAttachment"];
            string strEmailLogin = ConfigurationManager.AppSettings["EmailLogin"];
            string strEmailPassword = ConfigurationManager.AppSettings["EmailPassword"];
            string strEmailFrom = ConfigurationManager.AppSettings["EmailFrom"];
            string strEmailRecipients = ConfigurationManager.AppSettings["EmailRecipients"];

            string strMasterWorkbookFullPath = strExcelPath + strExcelFile;
            string strExcelWorkingFileFullPath = strExcelPath + strProjectTitle + "_" + strReportType + "_" + DateTime.Now.ToString("yyyyMMdd_HHmm") + ".xlsx";

            string strTimeBlockStartLocal = "";
            string strTimeBlockEndLocal = "";
            string strTimeBlockStartUTC = "";
            string strTimeBlockEndUTC = "";

            string strTimeBlockType = ConfigurationManager.AppSettings["TimeBlockType"];
            string strManualBlockStart = ConfigurationManager.AppSettings["manualBlockStart"];
            string strManualBlockEnd = ConfigurationManager.AppSettings["manualBlockEnd"];
            string strTimeOffsetHrs = ConfigurationManager.AppSettings["TimeOffsetHrs"];
            string strBlockSizeHrs = ConfigurationManager.AppSettings["BlockSizeHrs"];

            string[,] strSensorID = new string[5000, 2];


            PrismStats[] prismStats = new PrismStats[5000];       // this class is defined in spreadsheetAPS.cs
            Prism[] prism = new Prism[500];  // prisms observed by the ATS


            ATSstats[] atsStats = new ATSstats[200];

            string[] strATS = new String[11];
            strATS[0] = "";
            strATS[1] = ConfigurationManager.AppSettings["ATS1"];
            strATS[2] = ConfigurationManager.AppSettings["ATS2"];
            strATS[3] = ConfigurationManager.AppSettings["ATS3"];
            strATS[4] = ConfigurationManager.AppSettings["ATS4"];
            strATS[5] = ConfigurationManager.AppSettings["ATS5"];
            strATS[6] = ConfigurationManager.AppSettings["ATS6"];
            strATS[7] = ConfigurationManager.AppSettings["ATS7"];
            strATS[8] = ConfigurationManager.AppSettings["ATS8"];
            strATS[9] = ConfigurationManager.AppSettings["ATS9"];
            strATS[10] = ConfigurationManager.AppSettings["ATS10"];

            //==== Set the EPPlus license
            ExcelPackage.LicenseContext = LicenseContext.Commercial;


            //====[ Main Program ]====================================================================================

            gnaT.WelcomeMessage("projectPerformance 20221111");

            string strSoftwareLicenseTag = "PJTPFM";
            gnaT.checkLicenseValidity(strSoftwareLicenseTag, strProjectTitle, strEmailLogin, strEmailPassword, strSendEmails);





            //goto ThatsAllFolks;

            //==== Environment check

            Console.WriteLine("");
            Console.WriteLine("1. Check system environment");
            Console.WriteLine("   Project: " + strProjectTitle);
            Console.WriteLine("   Master workbook: " + strMasterWorkbookFullPath);

            gnaDBAPI.testDBconnection(strDBconnection);
            int i = 1;

            if (strCheckWorksheetsExist == "Yes")
            {
                Console.WriteLine("   Check Existance of workbook & worksheets");
                gnaSpreadsheetAPI.checkWorksheetExists(strMasterWorkbookFullPath, strSurveyWorksheet);
                gnaSpreadsheetAPI.checkWorksheetExists(strMasterWorkbookFullPath, strHourlyPerformanceWorksheet);
                gnaSpreadsheetAPI.checkWorksheetExists(strMasterWorkbookFullPath, strDailyPerformanceWorksheet);
            }
            else
            {
                Console.WriteLine("   Existance of workbook & worksheets is not checked");
            }

            // Find the first output rows
            int iFirstEmpty24hrRow = gnaSpreadsheetAPI.findFirstEmptyRow(strMasterWorkbookFullPath, strDailyPerformanceWorksheet, strFirstOutputRow, "2");

            // generating the time blocks WARNING: NOT TO BE REPLICATED (This is very specific to this program)

            string strStart24hrsUTC, strEnd24hrsUTC;
            string strStart24hrsLocal, strEnd24hrsLocal;
            int iDays = 0;


            Console.WriteLine("   Generate time blocks");

            switch (strTimeBlockType)
            {
                case "Manual":
                    strTimeBlockStartUTC = gnaT.convertLocalToUTC(strManualBlockStart);
                    strTimeBlockEndUTC = gnaT.convertLocalToUTC(strManualBlockEnd);
                    strStart24hrsUTC = strTimeBlockStartUTC.Replace("'", "").Trim();
                    strEnd24hrsUTC = strTimeBlockEndUTC.Replace("'", "").Trim();
                    strStart24hrsUTC = "'" + strStart24hrsUTC.Substring(0, 10) + " 00:00:00' ";
                    strEnd24hrsUTC = "'" + strEnd24hrsUTC.Substring(0, 10) + " 23:59:59' ";
                    strStart24hrsLocal = strStart24hrsUTC;  // not true but leave
                    strEnd24hrsLocal = strEnd24hrsUTC;
                    strTimeBlockStartLocal = strStart24hrsLocal;
                    strTimeBlockEndLocal = strEnd24hrsLocal;
                    iDays = gnaT.daysBetweenDates(strStart24hrsUTC, strEnd24hrsUTC);
                    break;

                case "Schedule":
                    double dblStartTimeOffset = -1.0 * Convert.ToDouble(strTimeOffsetHrs);
                    double dblEndTimeOffset = dblStartTimeOffset - Convert.ToDouble(strBlockSizeHrs);
                    strTimeBlockStartLocal = " '" + DateTime.Now.AddHours(dblEndTimeOffset).ToString("yyyy-MM-dd HH:mm:ss") + "' ";
                    strTimeBlockEndLocal = " '" + DateTime.Now.AddHours(dblStartTimeOffset).ToString("yyyy-MM-dd HH:mm:ss") + "' ";
                    strTimeBlockStartUTC = gnaT.convertLocalToUTC(strTimeBlockStartLocal);
                    strTimeBlockEndUTC = gnaT.convertLocalToUTC(strTimeBlockEndLocal);
                    strStart24hrsUTC = strTimeBlockStartUTC.Replace("'", "").Trim();
                    strEnd24hrsUTC = strTimeBlockEndUTC.Replace("'", "").Trim();
                    strStart24hrsUTC = "'" + strStart24hrsUTC.Substring(0, 10) + " 00:00:01' ";
                    strEnd24hrsUTC = "'" + strEnd24hrsUTC.Substring(0, 10) + " 23:59:59' ";
                    iDays = gnaT.daysBetweenDates(strStart24hrsUTC, strEnd24hrsUTC);
                    break;

                default:
                    Console.WriteLine("\nError in Timeblock Type");
                    Console.WriteLine("   Time block type: " + strTimeBlockType);
                    Console.WriteLine("   Must be Manual or Schedule");
                    Console.WriteLine("\nPress key to exit..."); Console.ReadKey();
                    goto ThatsAllFolks;
                    break;
            }

            string strDateTime = DateTime.Now.ToString("yyyyMMdd_HHmm");
            string strDateTimeUTC = DateTime.UtcNow.ToString("yyyy-MM-dd HH:mm");   //2022-07-26 13:45:15

            string strExportFile = strExcelPath + strContractTitle + "_" + strReportType + "_" + strDateTime + ".xlsx";
            Console.WriteLine("");

            string strTimeStamp = strTimeBlockEndLocal + "\n(local)";

            Console.WriteLine("   Time block type: " + strTimeBlockType);
            Console.WriteLine("     " + strTimeBlockStartLocal.Replace("'", "") + " Local");
            Console.WriteLine("     " + strTimeBlockEndLocal.Replace("'", "") + " Local");
            Console.WriteLine("     " + strStart24hrsUTC.Replace("'", "").Trim() + "  24hr start");
            Console.WriteLine("     " + strEnd24hrsUTC.Replace("'", "").Trim() + "  24hr end");
            Console.WriteLine("     Days: " + iDays);

            // generate the 24hr time windows

            CultureInfo provider = CultureInfo.InvariantCulture;
            DateTime dtStart24hr = DateTime.ParseExact(strStart24hrsUTC.Replace("'", "").Trim(), "yyyy-MM-dd HH:mm:ss", provider);
            DateTime dtEnd24hr = DateTime.ParseExact(strEnd24hrsUTC.Replace("'", "").Trim(), "yyyy-MM-dd HH:mm:ss", provider);
            DateTime dtStartTime = dtStart24hr;


            string[,] str24hrBlocks = new string[720, 2];

            for (int j = 1; j <= iDays; j++)
            {
                DateTime dtEndTime = dtStartTime.AddHours(24.0);
                if (dtEndTime > dtEnd24hr)
                {
                    dtEndTime = dtEnd24hr;
                }
                str24hrBlocks[j, 0] = "'" + dtStartTime.ToString("yyyy-MM-dd HH:mm:ss") + "'";      // start of the time block
                str24hrBlocks[j, 1] = "'" + dtEndTime.ToString("yyyy-MM-dd HH:mm:ss") + "'";        // end of the time block
                dtStartTime = dtEndTime;
            }

//**** later here must go the hours time blocks
//Console.WriteLine("\n2. Extract point names");
//string[] strPointNames = gnaSpreadsheetAPI.readPointNames(strMasterWorkbookFullPath, strSurveyWorksheet, strFirstDataRow);
//Console.WriteLine("3. Extract SensorID");
//strSensorID = gnaDBAPI.getSensorIDfromDB(strDBconnection, strPointNames, strProjectTitle);

//Console.WriteLine("4. Write SensorID to workbook");
//gnaSpreadsheetAPI.writeSensorID(strMasterWorkbookFullPath, strSurveyWorksheet, strSensorID, strFirstDataRow);

EntryPoint:
            Console.WriteLine("4. Count the prisms");
            int iPrismsTotal = gnaSpreadsheetAPI.countPrisms(strMasterWorkbookFullPath, strSurveyWorksheet, strFirstDataRow) - 1;
            Console.WriteLine("     Prisms: " + iPrismsTotal);


            Console.WriteLine("5. Process each ATS/day");

            //          
            //              
            //      iATScolumn=??
            //      make copy of master spreadsheet
            //
            //   24hr ATS data
            //      For each ATS
            //          From Survey Worksheet
            //              Extract the prisms for that ATS -> Prism
            //              Count no of prisms observed by that ATS: iNoOfPrisms    
            //          For each Day (TimeblockStart - TimeblockEnd) in the time window
            //              Extract the max no of readings possible on a prism : iMax
            //              Compute the total max number of readings for the Day: iTotal=iNoOfPrisms x iMax
            //              Clear the PrismStats array
            //              Reset the observationCounter
            //              For each prism read by the ATS         
            //                  Extract number of readings from DB: iSuccessfulObs
            //                  Increment the iObservationCounter: iObservationCounter+iSuccessfulObs
            //                  next prism
            //              Compute %: (iObservationCounter/iTotal)*100
            //              Find first available row in the ATS column
            //              Write to the ATSstats 
            //                  Date
            //                  Successful obs: iObservationCounter
            //                  iFailed Obs:  iTotal-iObservationCounter
            //                  % success
            //              Increment the iATScolumn
            //              next Day
            //
            //   Prism stats
            //      Define no of days for readings (1,2,3)
            //      Define time slot size
            //      Generate the time slots
            //      For each ATS
            //          For each Timeslot (TimeblockStart - TimeblockEnd) in the time window
            //              Extract the max no of readings possible on a prism : iMax
            //              For each prism read by the ATS         
            //                  Extract number of readings from DB: iSuccessfulObs
            //                  Compute %: (iSuccessfulObs/iMax)*100
            //              Write to the ATSstats 
            //                  % success
            //              Increment the iATScolumn     
            //              Next Timeslot
            //          next ATS
            //
            //      email Spreadsheet
            //      Update logs



            int iRow = 0;
            string strNextPrism = "";




            FileInfo newFile = new FileInfo(strMasterWorkbookFullPath);
            using (ExcelPackage package = new ExcelPackage(newFile))
            {
                ExcelWorksheet namedWorksheet = package.Workbook.Worksheets[strSurveyWorksheet];
                int iMaxCount = 0;
                int k = 1;
                int iATScol = -2;

                //   24hr ATS data
                do   // Fort each ATs
                {
                    string ATS = strATS[k];
                    Console.WriteLine("     " + ATS);
                    iRow = Convert.ToInt16(strFirstDataRow);
                    int iPrismCounter = 0;
                    iMaxCount = 0;



                    // Extract the prisms observed by this ATS -> Prism
                    do
                    {
                        if (Convert.ToString(namedWorksheet.Cells[iRow, 10].Value) == ATS)
                        {
                            prism[iPrismCounter] = new Prism();
                            prism[iPrismCounter].SensorID = Convert.ToString(namedWorksheet.Cells[iRow, 1].Value);
                            prism[iPrismCounter].Name = Convert.ToString(namedWorksheet.Cells[iRow, 2].Value);
                            prism[iPrismCounter].ReplacementName = Convert.ToString(namedWorksheet.Cells[iRow, 9].Value);
                            prism[iPrismCounter].ATS = Convert.ToString(namedWorksheet.Cells[iRow, 10].Value);
                            iPrismCounter++;
                        }
                        iRow++;
                        strNextPrism = Convert.ToString(namedWorksheet.Cells[iRow, 2].Value);
                    } while (strNextPrism != "");

                    prism[iPrismCounter] = new Prism();
                    prism[iPrismCounter].SensorID = "Missing";
                    prism[iPrismCounter].Name = "TheEnd";
                    prism[iPrismCounter].ReplacementName = "TheEnd";
                    prism[iPrismCounter].ATS = "TheEnd";

                    int iNumberOfPrisms = iPrismCounter;

                    //=================== checked to here =============================


                    string strTimeBlockStart = "";
                    string strTimeBlockEnd = "";


                    //For each Day (TimeblockStart - TimeblockEnd) in the time window
                    for (int iDayNo = 1; iDayNo <= iDays; iDayNo++)
                    {
                        strTimeBlockStart = str24hrBlocks[iDayNo, 0];
                        strTimeBlockEnd = str24hrBlocks[iDayNo, 1];
                        iMaxCount = 0;
                        int iActualObservations = 0;
                        // For each prism count the number of observations per time slot
                        for (i = 0; i <= iNumberOfPrisms; i++)
                        {
                            prismStats[i] = new PrismStats();
                            prismStats[i].Name = prism[i].Name;
                            prismStats[i].SensorID = prism[i].SensorID;
                            prismStats[i].ReplacementName = prism[i].ReplacementName;
                            prismStats[i].TimeBlockStart = strTimeBlockStart;
                            prismStats[i].TimeBlockEnd = strTimeBlockEnd;

                            if (prism[i].SensorID == "Missing")
                            {
                                prismStats[i].ObservationCount = 0;
                            }
                            else
                            {
                                prismStats[i].ObservationCount = gnaDBAPI.getNoOfObservations(strDBconnection, prismStats[i].SensorID, strTimeBlockStart, strTimeBlockEnd);

                                if (prismStats[i].ObservationCount > iMaxCount)
                                {
                                    iMaxCount = prismStats[i].ObservationCount;
                                }
                                iActualObservations = iActualObservations + prismStats[i].ObservationCount;
                            }
                        } // end of the day block loop

                        double dblTotal = Convert.ToDouble(iMaxCount * iNumberOfPrisms); // 100% possible observations in this day
                        double dblPercentageSuccess;

                        if ((iActualObservations > 0) && (dblTotal > 0.0))
                        {
                            dblPercentageSuccess = Math.Round((iActualObservations / dblTotal) * 100.0, 1);
                        }
                        else
                        {
                            dblPercentageSuccess = 0;
                        }

                        atsStats[iDayNo] = new ATSstats();
                        atsStats[iDayNo].ATSname = strATS[k];
                        atsStats[iDayNo].MaxPossibleObs = iMaxCount * iNumberOfPrisms;
                        atsStats[iDayNo].SuccessfulObs = iActualObservations;
                        atsStats[iDayNo].PercentageSuccess = dblPercentageSuccess;
                        atsStats[iDayNo].TimeBlockStart = strTimeBlockStart.Replace("'", "");
                        atsStats[iDayNo].TimeBlockEnd = strTimeBlockEnd.Replace("'", "");
                        atsStats[iDayNo].Date = strTimeBlockStart.Replace("'", "").Substring(0,11);

                    } // end of days

                    // Write this data to the ATSdata worksheet

                    iATScol = iATScol + 4;
                    gnaSpreadsheetAPI.writeATSPerformanceData(strMasterWorkbookFullPath, strDailyPerformanceWorksheet, atsStats, iATScol, iDays);

                    

                    k++;
                } while (strATS[k] != "blank");

                atsStats[k] = new ATSstats();
                atsStats[k].ATSname = "TheEnd";

                goto ThatsAllFolks;

                //gnaSpreadsheetAPI.writePerformanceData(strMasterWorkbookFullPath, strDailyPerformanceWorksheet, atsStats, iFirstOutputRow, 2);

                //******* write the data in writePerformanceData

            }







ThatsAllFolks:

            string strFreezeScreen = ConfigurationManager.AppSettings["freezeScreen"];

            if (strFreezeScreen == "Yes")
            {
                Console.WriteLine("\nfreezeScreen set to Yes");
                Console.WriteLine("press key to exit..");
                Console.ReadKey();
            }

            Environment.Exit(0);
            Console.WriteLine("\nTask Complete....");

        }

    }
}

using System;
using System.Collections.Generic;
using Advent.Geneva.WFM.Framework.BaseImplementation;
using Advent.Geneva.WFM.Framework.Interfaces;
using Advent.Geneva.WFM.SQLDataAccess;
using Advent.Geneva.WFM.GenevaDataAccess;
using System.IO;
using SSRStoPDF.RE2005;
using System.Text.RegularExpressions;


namespace SSRStoPDF
{
    class Program
    {
        private static ReportExecutionService rsExec;

        private static string status = "";
        private static string outputFolder = "";

        private static string portfolio;
        private static DateTime dtStartDate;
        private static DateTime dtEndDate;
        private static DateTime dtKnowledgeDate;
        private static DateTime dtPriorKnowledgeDate;

        static void Main(string[] args)
        {
            //SaveReport();
            //ShowReport();

            //Exception ex = RunSSRS("DEECEF", "2015-05-01", "2015-05-31");

            Run();

        }


        private static void Run()
        {
            UpdateStatus("#######################################", false);
            UpdateStatus("########## Begin Processing ###########", false);
            UpdateStatus(" ", false);

            try
            {
                outputFolder = "D:\\Share\\Temp";  // GetSettingValue("OutputFolder");


                //#########  Set User Parameters #############
                //activityRun.CurrentStep = "Set Parameters";
                UpdateStatus("Set User Parameters", false);

                //Set Activity Start Time
                //activityRun.StartDateTime = DateTime.Now;

                //Read Activty Parameter
                portfolio = "DEECEF"; // activityRun.GetParameterValue("Portfolio");
                string strStartDate = "01/09/2016 00:00:00";// activityRun.GetParameterValue("StartDate");
                string strEndDate = "01/15/2016 23:59:59"; // activityRun.GetParameterValue("EndDate");
                string strKnowledgeDate = "01/18/2016 15:42:05";// activityRun.GetParameterValue("StartDate");
                string strPriorKnowledgeDate = "01/11/2016 16:58:34"; // activityRun.GetParameterValue("EndDate");
                dtStartDate = DateTime.ParseExact(strStartDate, "MM/dd/yyyy HH:mm:ss", System.Globalization.CultureInfo.InvariantCulture);
                dtEndDate = DateTime.ParseExact(strEndDate, "MM/dd/yyyy HH:mm:ss", System.Globalization.CultureInfo.InvariantCulture);
                dtKnowledgeDate = DateTime.ParseExact(strKnowledgeDate, "MM/dd/yyyy HH:mm:ss", System.Globalization.CultureInfo.InvariantCulture);
                dtPriorKnowledgeDate = DateTime.ParseExact(strPriorKnowledgeDate, "MM/dd/yyyy HH:mm:ss", System.Globalization.CultureInfo.InvariantCulture);

                //Creating Generic (base) ReportParameter List; add required parameter to this object
                ReportParameterList base_parameters = new ReportParameterList();
                base_parameters.Add(new Advent.Geneva.WFM.GenevaDataAccess.ReportParameter("ConnectionString", Properties.Settings.Default.GenevaConnection));
                base_parameters.Add(new Advent.Geneva.WFM.GenevaDataAccess.ReportParameter("Portfolio", portfolio));
                base_parameters.Add(new Advent.Geneva.WFM.GenevaDataAccess.ReportParameter("PeriodStartDate", dtStartDate));
                base_parameters.Add(new Advent.Geneva.WFM.GenevaDataAccess.ReportParameter("PeriodEndDate", dtEndDate));
                base_parameters.Add(new Advent.Geneva.WFM.GenevaDataAccess.ReportParameter("KnowledgeDate", dtKnowledgeDate));
                base_parameters.Add(new Advent.Geneva.WFM.GenevaDataAccess.ReportParameter("PriorKnowledgeDate", dtPriorKnowledgeDate));
                base_parameters.Add(new Advent.Geneva.WFM.GenevaDataAccess.ReportParameter("AccountingRunType", "ClosedPeriod"));
                base_parameters.Add(new Advent.Geneva.WFM.GenevaDataAccess.ReportParameter("AccountingCalendar", portfolio));
                base_parameters.Add(new Advent.Geneva.WFM.GenevaDataAccess.ReportParameter("RegionalSettings", "en-IE"));
                base_parameters.Add(new Advent.Geneva.WFM.GenevaDataAccess.ReportParameter("DisableHyperlinks", "True"));
                base_parameters.Add(new Advent.Geneva.WFM.GenevaDataAccess.ReportParameter("QuantityPrecision", "4"));
                

                

                //#########  Get Report List PDF #############
                UpdateStatus("Get Report List PDF ", false);
                List<ReportParameters> reportListPDF = getReportNames("PDF", base_parameters);

                foreach (ReportParameters report in reportListPDF)
                {
                    UpdateStatus("Run PDF - " + report.FileName, false);

                    Exception rdlException = RunSSRSReport(outputFolder, report);
                    if (rdlException != null)
                    {
                       // activityRun.UpdateFailedActivityRun(rdlException);
                        //activityRun.UpdateActivityStep(false, "Fail");
                        //activityRun.Save();
                        throw rdlException;
                    }
                    UpdateStatus("Success PDF - " + report.FileName, false);
                }

                WriteToLog();


            }
            catch (Exception exp)
            {
                //Set Failure Flag
                Exception e = new Exception(status + Environment.NewLine + "-----------Exception Message--------------" + Environment.NewLine + exp.Message);
                //activityRun.UpdateFailedActivityRun(e);
                //activityRun.UpdateActivityStep(false, "Fail");
                //activityRun.Save();
            }
            finally
            {
                //Set Activity End Time and Save Activity
                //activityRun.EndDateTime = DateTime.Now;
                //activityRun.UpdateSuccessfulActivityRun();
                //activityRun.Save();
            }
        }


        private static Exception RunSSRSReport(string OutputFolder, ReportParameters Report)
        {
            rsExec = new ReportExecutionService();
            rsExec.Url = Properties.Settings.Default.SSRStoPDF_RE2005_ReportExecutionService;
            //rsExec.Url =  "http://dubbtvm03/reportserver/reportexecution2005.asmx";

            rsExec.UseDefaultCredentials = true;

            string historyID = null;
            string deviceInfo = null;
            string format = "PDF";
            Byte[] results;
            string encoding = String.Empty;
            string mimeType = String.Empty;
            string extension = String.Empty;
            Warning[] warnings = null;
            string[] streamIDs = null;

            var p = "";
            for (int i = 0; i < Report.Parameters.Count; i++)
            {
                p = "Param[" + i.ToString() + "]" + Report.Parameters[i].Name + "|" + Report.Parameters[i].Value.ToString() + Environment.NewLine + p;

            }
            UpdateStatus(p, true);


            // Path of the Report - XLS, PDF etc.
            string FilePath = OutputFolder + "\\" + GetOutputFileName(Report) + ".pdf";

            UpdateStatus(Report.FileName + ", Output : " + FilePath, true);

            // Name of the report - Please note this is not the RDL file.
            string _reportName = @"/GenevaReports/" + Report.Name;

            UpdateStatus(Report.FileName + ", _reportName : " + _reportName, true);
            UpdateStatus("###### Report: " + Report.Name + " #########", false);

            try
            {
                UpdateStatus("Load Report",false);

                ExecutionInfo ei = rsExec.LoadReport(_reportName, historyID);
                ParameterValue[] parameters = new ParameterValue[Report.Parameters.Count];

                UpdateStatus("Set Parameters " + Report.Parameters.Count, true);

                for (int i = 0; i < Report.Parameters.Count; i++)
                {

                    parameters[i] = new ParameterValue();
                    parameters[i].Name = Report.Parameters[i].Name;
                    parameters[i].Value = (string)Report.Parameters[i].Value;

                    UpdateStatus(Report.Name + "|" + parameters[i].GetType().ToString() + "|" + Report.Parameters[i].Name + "|" + Report.Parameters[i].Value.ToString(), true);
                }

                rsExec.SetExecutionParameters(parameters, "en-GB");

                DataSourceCredentials dataSourceCredentials2 = new DataSourceCredentials();
                dataSourceCredentials2.DataSourceName = Properties.Settings.Default.DataSourceName; 
                dataSourceCredentials2.UserName = Properties.Settings.Default.GenevaUser;
                dataSourceCredentials2.Password = Properties.Settings.Default.GenevaPass;


                DataSourceCredentials[] _credentials2 = new DataSourceCredentials[] { dataSourceCredentials2 };

                var c = "";
                for (int i = 0; i < _credentials2.Length; i++)
                {
                    c = "_credentials2[" + i.ToString() + "]:" + _credentials2[i].DataSourceName + "|" + 
                                                                      _credentials2[i].UserName + "|" + 
                                                                      _credentials2[i].Password + "|" + Environment.NewLine + c;
                }
                UpdateStatus(c, true);

                rsExec.SetExecutionCredentials(_credentials2);
                //rsExec.UseDefaultCredentials = true;

                UpdateStatus("Pre Render Report..." +
                         "\tformat: " + format +
                         "\tdeviceInfo: " + deviceInfo, true);

                results = rsExec.Render(format, deviceInfo, out extension,
                                                            out encoding,
                                                            out mimeType,
                                                            out warnings,
                                                            out streamIDs);

                UpdateStatus("Post Render Report..." +
                         "\tdeviceInfo: " + deviceInfo +
                         "\textension: " + extension +
                         "\tencoding: " + encoding +
                         "\tmimeType: " + mimeType, true);

                UpdateStatus("###### PDF Size: " + results.Length.ToString() + " #########", true);

                using (FileStream stream = File.OpenWrite(FilePath))
                {
                    stream.Write(results, 0, results.Length);
                }

            }
            catch (Exception ex)
            {
                UpdateStatus("--- ERROR ---" + Environment.NewLine + ex.Message, false);
                return new Exception(status);
            }

            return null;

        }

        private static List<ReportParameters> getReportNames(string ReportType, ReportParameterList base_ParameterList)
        {
            List<ReportParameters> reports = new List<ReportParameters>();


            //reports.Add(new ReportParameters("Fund Allocation Percentages",
            //                                 "fundalloc.rsl"));
            //reports[0].AddParameterList(base_ParameterList);
            //reports[0].AddParameters("AccountingRunType", "NAV");



            reports.Add(new ReportParameters("Custom Unsettled Income Report #0008",
                                                 "0008CustomUnsettledIncomeReport.rsl"));
            reports[0].AddParameterList(base_ParameterList);

            reports.Add(new ReportParameters("Custom Portfolio Valuation Report #0014",
                                             "0014CustomPortfolioValuationReport.rsl"));
            reports[1].AddParameterList(base_ParameterList);

            reports.Add(new ReportParameters("Custom Unsettled Transactions Report #0012",
                                             "0012CustomUnsettledTransactionsReport.rsl"));
            reports[2].AddParameterList(base_ParameterList);

            reports.Add(new ReportParameters("Custom Profit and Loss Report #0015",
                                             "0015CustomProfitandLossReport.rsl"));
            reports[3].AddParameterList(base_ParameterList);
            reports[3].AddParameters("AccountingCalendar", portfolio);

            reports.Add(new ReportParameters("Custom Realised Gain Loss Ledger #0005",
                                             "0005CustomRealisedGainLossLedger.rsl"));
            reports[4].AddParameterList(base_ParameterList);

            reports.Add(new ReportParameters("Custom Appreciation and Depreciation on Foreign Currency Contracts #0018",
                                             "0018CustomAppreciationandDepreciationonFCC.rsl"));
            reports[5].AddParameterList(base_ParameterList);

            reports.Add(new ReportParameters("Custom Other Assets And Liabilities #0050",
                                             "0050CustomOtherAssetsAndLiabilities.rsl"));
            reports[6].AddParameterList(base_ParameterList);

            reports.Add(new ReportParameters("Custom Performance Attribution Report #0056",
                                             "0056CustomPerformanceAttributionReport.rsl"));
            reports[7].AddParameterList(base_ParameterList);

            reports.Add(new ReportParameters("Trial Balance",
                                             "glmap_fundtrialbal.rsl"));
            reports[8].AddParameterList(base_ParameterList);
            reports[8].AddParameters("FundLegalEntity", portfolio);

            reports.Add(new ReportParameters("Cash Appraisal",
                                             "glmap_cashapp.rsl"));
            reports[9].AddParameterList(base_ParameterList);

            reports.Add(new ReportParameters("Fund Structure NAV",
                                             "nav.rsl"));
            reports[10].AddParameterList(base_ParameterList);
            reports[10].AddParameters("AccountingRunType", "NAV");
            reports[10].AddParameters("KnowledgeDate", DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"));

            reports.Add(new ReportParameters("Fund Capital Ledger",
                                             "fundcapldgr.rsl"));
            reports[11].AddParameterList(base_ParameterList);
            reports[11].AddParameters("AccountingRunType", "NAV");
            reports[11].AddParameters("KnowledgeDate", DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"));

            reports.Add(new ReportParameters("Fund Allocation Percentages",
                                             "fundalloc.rsl"));
            reports[12].AddParameterList(base_ParameterList);
            reports[12].AddParameters("AccountingRunType", "NAV");
            reports[12].AddParameters("KnowledgeDate", DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"));

            reports.Add(new ReportParameters("Fund Allocated Income Detail",
                                             "fundincdet.rsl"));
            reports[13].AddParameterList(base_ParameterList);
            reports[13].AddParameters("AccountingRunType", "NAV");
            reports[13].AddParameters("KnowledgeDate", DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"));

            reports.Add(new ReportParameters("Statement of Net Assets",
                                             "glmap_netassets.rsl"));
            reports[14].AddParameterList(base_ParameterList);

            reports.Add(new ReportParameters("Statement of Changes in Net Assets",
                                             "glmap_chginassets.rsl"));
            reports[15].AddParameterList(base_ParameterList);

            reports.Add(new ReportParameters("Local Position Appraisal",
                                             "locposapp.rsl"));
            reports[16].AddParameterList(base_ParameterList);
            reports[16].AddParameters("Consolidate", "None");



            if (ReportType == "CSV")
            {

                reports.Add(new ReportParameters("Custom Other Assets And Liabilities #0050",
                                                 "0050CustomOtherAssetsAndLiabilities.rsl"));
                reports[17].AddParameterList(base_ParameterList);
            }

            return reports;
        }

        
        private static void UpdateStatus(string Status, bool Debug)
        {
            if (Debug)
            {
                if (Properties.Settings.Default.Debug)
                {
                    status = status + Environment.NewLine + Status;
                }
            }
            else
            {
                status = status + Environment.NewLine + Status;
            }
        }

        private static void WriteToLog()
        {
            if(Properties.Settings.Default.Logging)
            {
                TextWriter RunWriter = new StreamWriter(outputFolder + "\\" + "Run.log");

                RunWriter.Write(status);
                RunWriter.Close();
            }
        }

        private static string GetOutputFileName(ReportParameters Report)
        {
            string OutputFileName = portfolio + "_";

            Regex rgx = new Regex("[^a-zA-Z0-9]");
            string cleanReportName = rgx.Replace(Report.Name, "");

            OutputFileName += cleanReportName + "_" + dtEndDate.ToString("yyyyMMdd") + "_" + DateTime.Now.ToString("yyyyMMddHHmmss");

            return OutputFileName;
        }

    }
}

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Advent.Geneva.WFM.GenevaDataAccess;

namespace SSRStoPDF
{
    class ReportParameters
    {
        public string Name { get; set; }
        public string FileName { get; set; }
        public List<ReportParameter> Parameters { get; set; }
        public ReportParameterList ParameterList { get; set; }

        public ReportParameters(string name, string filename)
        {
            Name = name;
            FileName = filename;

            Parameters = new List<Advent.Geneva.WFM.GenevaDataAccess.ReportParameter>();
            ParameterList = new ReportParameterList();

            getDefaultParameters();
        }

        public void AddParameters(string name, string value)
        {
            //Check if the Parameter already exists and remove first.
            for (int i = 0; i < this.Parameters.Count; i++)
            {
                if (this.Parameters[i].Name == name )
                {
                    var dataType = this.Parameters[i].GetType();
                    ReportParameter param = null;

                    if (dataType.ToString() == "DateTime")
                    {
                        param = new ReportParameter(this.Parameters[i].Name, (DateTime)this.Parameters[i].Value);
                    }
                    else
                    {
                        param = new ReportParameter(this.Parameters[i].Name, this.Parameters[i].Value.ToString());
                    }

                    this.Parameters.RemoveAt(i) ; 
                    this.ParameterList.RemoveReportParameter(name);
                    break;
                }
                
            }
            
            Parameters.Add(new ReportParameter(name, value));
            ParameterList.Add(new ReportParameter(name, value));
            
        }


        public void AddParameterList(ReportParameterList paramList)
        {
            foreach (var param in paramList.GetParameterList())
            {
                if (param.Key.ToString().Contains("Date"))
                {
                    DateTime dt = DateTime.ParseExact(param.Value.ToString(), "yyyy-MM-dd:HH:mm:ss", System.Globalization.CultureInfo.InvariantCulture);
                    this.AddParameters(param.Key.ToString(), dt.ToString("yyyy-MM-dd HH:mm:ss"));
                }
                else
                {
                    this.AddParameters(param.Key.ToString(), (string)param.Value);
                }
            }

        }

        private void getDefaultParameters()
        {
            RS2005.ReportingService2005 rs = new RS2005.ReportingService2005();
            rs.UseDefaultCredentials = true;
            rs.Url = Properties.Settings.Default.SSRStoPDF_RS2005_ReportingService2005;

            string historyID = null;
            string deviceInfo = null;
            string format = "EXCEL";
            Byte[] results;
            string encoding = String.Empty;
            string mimeType = String.Empty;
            string extension = String.Empty;
            RE2005.Warning[] warnings = null;
            string[] streamIDs = null;

            RS2005.DataSourceCredentials dataSourceCredentials2 = new RS2005.DataSourceCredentials();
            dataSourceCredentials2.DataSourceName = Properties.Settings.Default.DataSourceName;
            dataSourceCredentials2.UserName = Properties.Settings.Default.GenevaUser;
            dataSourceCredentials2.Password = Properties.Settings.Default.GenevaPass;
            RS2005.DataSourceCredentials[] _credentials = new RS2005.DataSourceCredentials[] { dataSourceCredentials2 };
            RS2005.ReportParameter[] reportParameters = null;

            try
            {
                reportParameters = rs.GetReportParameters(@"/GenevaReports/" + this.Name, null, false, null, _credentials);
                //RE2005.ExecutionInfo ei = rsExec.LoadReport(_reportName, historyID);

                foreach (var param in reportParameters)
                {
                    if (param.DefaultValues != null 
                        && param.DefaultValues[0] != null
                        && param.PromptUser)
                        //&& param.Name != "InventoryState"
                        //&& param.Name != "TranslationTable")
                    {
                        this.AddParameters(param.Name, param.DefaultValues[0].ToString());
                    }

                }

            }
            catch (Exception ex)
            {
                throw ex;
            }
        }


    }
}

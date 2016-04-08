using Microsoft.Reporting.WebForms;
using System;
using System.Collections.Generic;
using System.Text;

namespace OnReport
{
    /// <summary>
    /// Definition of the report.
    /// </summary>
    public class OnReportConfig
    {
        /// <summary>
        /// Path 'rdlc' file.
        /// </summary>
        public string ReportPath { get; set; }

        /// <summary>
        /// Name of data source in 'rdlc' file.
        /// </summary>
        public string DataSource { get; set; }

        /// <summary>
        /// Data of the report.
        /// </summary>
        public object ReportData { get; set; }

        /// <summary>
        /// Name of the file after output.
        /// </summary>
        public string OutputFileName { get; set; }

        /// <summary>
        /// Ouput format of the report.
        /// </summary>
        public OnReportOutputFormat OutputFormat { get; set; }

        /// <summary>
        /// Size and margins of the report.
        /// </summary>
        public OnReportSize ReportSize { get; set; }

        /// <summary>
        /// Report's parameters.
        /// </summary>
        public IList<OnReportParameter> Parameters { get; set; }

        /// <summary>
        /// Method that renders the report.
        /// </summary>
        /// <returns>OnReportResult</returns>
        public OnReportResult Render()
        {
            try
            {
                //Validations
                if (string.IsNullOrEmpty(ReportPath) || string.IsNullOrEmpty(DataSource) || ReportData == null)
                    throw new ArgumentNullException();

                OutputFileName = string.IsNullOrEmpty(OutputFileName) ? "Report" : OutputFileName;

                if (ReportSize == null)
                    ReportSize = new OnReportSize
                    {
                        PageWidth = "21cm",
                        PageHeight = "29.7cm",
                        MarginTop = "0.5cm",
                        MarginBottom = "0.5cm",
                        MarginLeft = "0.5cm",
                        MarginRight = "0.5cm"
                    };
                else
                {
                    ReportSize.PageWidth = ReportSize.PageWidth ?? "21cm";
                    ReportSize.PageHeight = ReportSize.PageHeight ?? "29.7cm";
                    ReportSize.MarginTop = ReportSize.MarginTop ?? "0.5cm";
                    ReportSize.MarginBottom = ReportSize.MarginBottom ?? "0.5cm";
                    ReportSize.MarginLeft = ReportSize.MarginLeft ?? "0.5cm";
                    ReportSize.MarginRight = ReportSize.MarginRight ?? "0.5cm";
                }

                //Method Execution
                using (var relat = new LocalReport())
                {
                    relat.ReportPath = ReportPath;

                    //Parameters
                    if (Parameters != null)
                    {
                        var parameters = new List<ReportParameter>();

                        foreach (var p in Parameters)
                        {
                            parameters.Add(new ReportParameter(p.Name, p.Value));
                        }

                        relat.SetParameters(parameters.ToArray());
                    }

                    //Data
                    var ds = new ReportDataSource
                    {
                        Name = DataSource,
                        Value = ReportData
                    };

                    relat.DataSources.Add(ds);

                    var reportType = OutputFormat.ToString();
                    string mimeType;
                    string encoding;
                    string fileNameExtension;

                    var sb = new StringBuilder();

                    sb.Append("<DeviceInfo>");
                    sb.Append($"<OutputFormat>{reportType}</OutputFormat>");
                    sb.Append($"<PageWidth>{ReportSize.PageWidth}</PageWidth>");
                    sb.Append($"<PageHeight>{ReportSize.PageHeight}</PageHeight>");
                    sb.Append($"<MarginTop>{ReportSize.MarginTop}</MarginTop>");
                    sb.Append($"<MarginBottom>{ReportSize.MarginBottom}</MarginBottom>");
                    sb.Append($"<MarginLeft>{ReportSize.MarginLeft}</MarginLeft>");
                    sb.Append($"<MarginRight>{ReportSize.MarginRight}</MarginRight>");
                    sb.Append("</DeviceInfo>");

                    var deviceInfo = sb.ToString();
                    Warning[] warnings;
                    string[] streams;
                    byte[] bytes;

                    bytes = relat.Render(
                        reportType,
                        deviceInfo,
                        out mimeType,
                        out encoding,
                        out fileNameExtension,
                        out streams,
                        out warnings);

                    var orr = new OnReportResult
                    {
                        Bytes = bytes,
                        MimeType = mimeType,
                        FileName = $"{OutputFileName}.{fileNameExtension}"
                    };

                    return orr;
                }
            }
            catch (Exception)
            {
                throw new OnReportRenderException("Could not render the report.");
            }
        }
    }
}

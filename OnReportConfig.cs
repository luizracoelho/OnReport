using Microsoft.Reporting.WebForms;
using System;
using System.Collections.Generic;
using System.Linq;
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
        /// Reports's DataSource
        /// </summary>
        public OnReportDataSource DataSource { get; set; }

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
        /// Enable or disabled external images.
        /// </summary>
        public bool EnableExternalImages { get; set; }

        /// <summary>
        /// Report's parameters.
        /// </summary>
        public IList<OnReportParameter> Parameters { get; set; }

        /// <summary>
        /// Report's SubReport Render
        /// </summary>
        public SubreportProcessingEventHandler SubReportProcessing { get; set; }

        /// <summary>
        /// Method that renders the report.
        /// </summary>
        /// <returns>OnReportResult</returns>
        public OnReportResult Render()
        {
            try
            {
                //Validations
                if (string.IsNullOrEmpty(ReportPath))
                    throw new ArgumentNullException();

                if (string.IsNullOrEmpty(DataSource.Name) || DataSource.Value == null)
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
                var relat = new LocalReport()
                {
                    ReportPath = ReportPath,

                    //ExternalImages
                    EnableExternalImages = EnableExternalImages
                };

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
                relat.DataSources.Add(new ReportDataSource
                {
                    Name = DataSource.Name,
                    Value = DataSource.Value
                });

                //Subreport
                relat.SubreportProcessing += SubReportProcessing;

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
            catch (Exception ex)
            {
                throw new OnReportRenderException(ex.InnerException?.InnerException?.Message);
            }
        }
    }
}

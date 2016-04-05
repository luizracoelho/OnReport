namespace OnReport
{
    /// <summary>
    /// Result of OnReport.
    /// </summary>
    public class OnReportResult
    {
        /// <summary>
        /// File in bytes.
        /// </summary>
        public byte[] Bytes { get; set; }

        /// <summary>
        /// Mime type.
        /// </summary>
        public string MimeType { get; set; }

        /// <summary>
        /// File Name
        /// </summary>
        public string FileName { get; set; }
    }
}

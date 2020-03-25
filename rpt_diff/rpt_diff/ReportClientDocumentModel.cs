using CrystalDecisions.ReportAppServer.ClientDoc;
using ExtensionMethods;
using System.Text.Json;

namespace rpt_diff
{
    /*
    *  ReportClientDocumentModel
    *  - newer
    *  - contains all atributes availible from rpt file
    */
    class ReportClientDocumentModel
    {
        public static void ProcessReport(ISCDReportClientDocument report, Utf8JsonWriter jsonw, string reportDoc = "ReportClientDocument")
        {
            jsonw.WritePropertyName(reportDoc);
            jsonw.WriteStartObject();
            jsonw.WriteString("DisplayName", report.DisplayName);
            jsonw.WriteString("IsModified", report.IsModified.ToStringSafe());
            jsonw.WriteString("IsOpen", report.IsOpen.ToStringSafe());
            jsonw.WriteString("IsReadOnly", report.IsReadOnly.ToStringSafe());
            jsonw.WriteString("LocaleID", report.LocaleID.ToStringSafe());
            jsonw.WriteString("MajorVersion", report.MajorVersion.ToStringSafe());
            jsonw.WriteString("MinorVersion", report.MinorVersion.ToStringSafe());
            jsonw.WriteString("Path", report.Path);
            jsonw.WriteString("PreferredViewingLocaleID", report.PreferredViewingLocaleID.ToStringSafe());
            jsonw.WriteString("ProductLocaleID", report.ProductLocaleID.ToStringSafe());
            jsonw.WriteString("ReportAppServer", report.ReportAppServer);
            Controllers.ProcessCustomFunctionController(report.CustomFunctionController, jsonw);
            Controllers.ProcessDatabaseController(report.DatabaseController, jsonw);
            Controllers.ProcessDataDefController(report.DataDefController, jsonw);
            Controllers.ProcessPrintOutputController(report.PrintOutputController,jsonw);
            Controllers.ProcessReportDefController(report.ReportDefController, jsonw);
            ReportDefModel.ProcessReportOptions(report.ReportOptions, jsonw);
            Controllers.ProcessSubreportController(report.SubreportController, jsonw);
            DataDefModel.ProcessSummaryInfo(report.SummaryInfo, jsonw);
            jsonw.WriteEndObject();
        }   
    }
}

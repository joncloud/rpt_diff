using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;

using CrystalDecisions.ReportAppServer.ClientDoc;

using ExtensionMethods;

namespace rpt_diff
{
    /*
    *  ReportClientDocumentModel
    *  - newer
    *  - contains all atributes availible from rpt file
    *  - can be used in future to merge changes or back convert from xml to rpt
    */
    class ReportClientDocumentModel
    {
        public static void ProcessReport(ISCDReportClientDocument report, XmlWriter xmlw, string reportDoc = "ReportClientDocument")
        {
            xmlw.WriteStartElement(reportDoc);
            xmlw.WriteElementString("DisplayName", report.DisplayName);
            xmlw.WriteElementString("IsModified", report.IsModified.ToStringSafe());
            xmlw.WriteElementString("IsOpen", report.IsOpen.ToStringSafe());
            xmlw.WriteElementString("IsReadOnly", report.IsReadOnly.ToStringSafe());
            xmlw.WriteElementString("LocaleID", report.LocaleID.ToStringSafe());
            xmlw.WriteElementString("MajorVersion", report.MajorVersion.ToStringSafe());
            xmlw.WriteElementString("MinorVersion", report.MinorVersion.ToStringSafe());
            xmlw.WriteElementString("Path", report.Path);
            xmlw.WriteElementString("PreferredViewingLocaleID", report.PreferredViewingLocaleID.ToStringSafe());
            xmlw.WriteElementString("ProductLocaleID", report.ProductLocaleID.ToStringSafe());
            xmlw.WriteElementString("ReportAppServer", report.ReportAppServer);
            Controllers.ProcessCustomFunctionController(report.CustomFunctionController, xmlw);
            Controllers.ProcessDatabaseController(report.DatabaseController, xmlw);
            Controllers.ProcessDataDefController(report.DataDefController, xmlw);
            Controllers.ProcessPrintOutputController(report.PrintOutputController,xmlw);
            Controllers.ProcessReportDefController(report.ReportDefController, xmlw);
            ReportDefModel.ProcessReportOptions(report.ReportOptions, xmlw);
            Controllers.ProcessSubreportController(report.SubreportController, xmlw);
            DataDefModel.ProcessSummaryInfo(report.SummaryInfo, xmlw);
            xmlw.WriteEndElement();
        }   
    }
}

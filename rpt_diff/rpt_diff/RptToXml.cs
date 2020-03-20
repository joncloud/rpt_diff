using CrystalDecisions.CrystalReports.Engine;
using CrystalDecisions.Shared;
using System;
using System.IO;
using System.Text;
using System.Xml;

namespace rpt_diff
{
    class RptToXml
    {
        class TemporaryFile : IDisposable
        {
            public string FilePath { get; }
            public TemporaryFile()
            {
                FilePath = Path.GetTempFileName();
            }

            public void CopyFrom(Stream stream)
            {
                using (var target = new FileStream(FilePath, FileMode.Truncate))
                {
                    stream.CopyTo(target);
                }
            }

            public void Dispose()
            {
                if (File.Exists(FilePath))
                {
                    File.Delete(FilePath);
                }
            }
        }

        public static void ConvertRptToXml(Stream reportBinary, Stream reportDefinition, int model)
        {
            using (var temporaryFile = new TemporaryFile())
            using (var report = new ReportDocument())
            using (var xmlw = new XmlTextWriter(reportDefinition, Encoding.UTF8) { Formatting = Formatting.Indented })
            {
                temporaryFile.CopyFrom(reportBinary);

                report.Load(temporaryFile.FilePath, OpenReportMethod.OpenReportByTempCopy);

                xmlw.WriteStartDocument();
                if (model == 0)
                {
                    ReportDocumentModel.ProcessReport(report, xmlw);
                }
                else
                {
                    ReportClientDocumentModel.ProcessReport(report.ReportClientDocument, xmlw);
                }

                xmlw.WriteEndDocument();
                xmlw.Flush();
            }
        }

        /*
         * ConvertRptToXml
         * Opens report file found in rptPath and converts it to xml using model specified by parameter model.
         * params:
         *  rptPath - full path to rpt file to be converted to xml
         *  model   - specifies which object model use to convert 
         *          - 0 = ReportDocumentModel
         *          - 1 = ReportClientDocumentModel (RAS)
         */
        public static void ConvertRptToXml(string reportBinaryPath, string reportDefinitionPath, int model)
        {
            using (var source = File.OpenRead(reportBinaryPath))
            using (var target = new FileStream(reportDefinitionPath, File.Exists(reportDefinitionPath) ? FileMode.Truncate : FileMode.CreateNew))
            {
                ConvertRptToXml(source, target, model);
            }
        }


    }
}

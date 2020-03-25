using CrystalDecisions.CrystalReports.Engine;
using CrystalDecisions.Shared;
using System;
using System.IO;
using System.Text.Json;

namespace rpt_diff
{
    class RptToJson
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

        public static void ConvertRptToJson(Stream reportBinary, Stream reportDefinition)
        {
            using (var temporaryFile = new TemporaryFile())
            using (var report = new ReportDocument())
            using (var jsonw = new Utf8JsonWriter(reportDefinition, new JsonWriterOptions { Indented = true }))
            {
                temporaryFile.CopyFrom(reportBinary);

                report.Load(temporaryFile.FilePath, OpenReportMethod.OpenReportByTempCopy);

                jsonw.WriteStartObject();

                ReportClientDocumentModel.ProcessReport(report.ReportClientDocument, jsonw);

                jsonw.WriteEndObject();
                jsonw.Flush();
            }
        }

        /*
         * ConvertRptToJson
         * Opens report file found in rptPath and converts it to json
         * params:
         *  reportBinaryPath - full path to rpt file to be converted to json
         *  reportDefinitionPath - full path to the json file to write
         */
        public static void ConvertRptToJson(string reportBinaryPath, string reportDefinitionPath)
        {
            using (var source = File.OpenRead(reportBinaryPath))
            using (var target = new FileStream(reportDefinitionPath, File.Exists(reportDefinitionPath) ? FileMode.Truncate : FileMode.CreateNew))
            {
                ConvertRptToJson(source, target);
            }
        }


    }
}

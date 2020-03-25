using CrystalDecisions.ReportAppServer.Controllers;
using ExtensionMethods;
using System.Text.Json;

namespace rpt_diff
{
    class Controllers
    {
        public static void ProcessCustomFunctionController(CustomFunctionController cfc, Utf8JsonWriter jsonw)
        {
            DataDefModel.ProcessCustomFunctions(cfc.GetCustomFunctions(), jsonw); 
        }
        public static void ProcessDatabaseController(DatabaseController dc, Utf8JsonWriter jsonw)
        {
            DataDefModel.ProcessDatabase(dc.Database, jsonw);
        }

        public static void ProcessDataDefController(DataDefController ddc, Utf8JsonWriter jsonw)
        {
            DataDefModel.ProcessDataDefinition(ddc.DataDefinition, jsonw);
        }


        public static void ProcessPrintOutputController(PrintOutputController poc, Utf8JsonWriter jsonw)
        {
            ReportDefModel.ProcessPrintOptions(poc.GetPrintOptions(), jsonw);
            ReportDefModel.ProcessSavedXMLExportFormats(poc.GetSavedXMLExportFormats(), jsonw);
        }

        public static void ProcessReportDefController(ReportDefController2 rdc, Utf8JsonWriter jsonw)
        {
            ReportDefModel.ProcessReportDefinition(rdc.ReportDefinition, jsonw);
        }

        public static void ProcessSubreportController(SubreportController sc, Utf8JsonWriter jsonw)
        {
            jsonw.WritePropertyName("Subreports");
            jsonw.WriteStartArray();
            foreach (string Subreport in sc.GetSubreportNames())
            {
                jsonw.WriteStartObject();
                ProcessSubreportClientDocument(sc.GetSubreport(Subreport),jsonw);
                ReportDefModel.ProcessSubreportLinks(sc.GetSubreportLinks(Subreport), jsonw);
                jsonw.WriteEndObject();
            }
            jsonw.WriteEndArray();
        }

        private static void ProcessSubreportClientDocument(SubreportClientDocument scd, Utf8JsonWriter jsonw)
        {
            jsonw.WriteString("EnableOnDemand", scd.EnableOnDemand.ToStringSafe());
            jsonw.WriteString("EnableReimport", scd.EnableReimport.ToStringSafe());
            jsonw.WriteString("IsImported", scd.IsImported.ToStringSafe());
            jsonw.WriteString("Name", scd.Name);
            jsonw.WriteString("SubreportLocation", scd.SubreportLocation);
            ProcessDatabaseController(scd.DatabaseController, jsonw);
            ProcessDataDefController(scd.DataDefController, jsonw);
            ProcessReportDefController(scd.ReportDefController, jsonw);
            ReportDefModel.ProcessReportOptions(scd.ReportOptions, jsonw);
        }

    }
}

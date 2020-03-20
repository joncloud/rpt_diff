using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Drawing;
using System.Diagnostics;
using System.IO;

namespace rpt_diff
{
    static class Program
    {
        /*
         * All return codes
         */
        enum ExitCode : int
        {
            Success = 0,
            WrongDiffApp = 1,
            WrongRptFile = 2,
            WrongArgs    = 3,
            WrongModel   = 4,
            ConvertError = 5
        }
        /*
         * The main entry point for the application.
         */
        [STAThread]
        static int Main(string[] args)
        {
            int ModelNumber = 0;
            switch (args.Length)
            {
                case 3:
                    {
                        if (!File.Exists(args[0]))
                        {
                            Console.Error.WriteLine("Error: Can't find RPT file - Bad RPTPath1");
                            WriteUsage();
                            return (int)ExitCode.WrongRptFile;
                        }
                        if (!Int32.TryParse(args[2], out ModelNumber) || (ModelNumber != 0 && ModelNumber != 1))
                        {
                            Console.Error.WriteLine("Error: Wrong ModelNumber select - must be 0 or 1");
                            WriteUsage();
                            return (int)ExitCode.WrongModel;
                        }
                        Console.WriteLine("Using object model: \"" + ((ModelNumber == 0) ? "ReportDocument" : "ReportClientDocument") + "\"");
                        Console.WriteLine("Converting file: \"" + args[0] + "\"");
                        try
                        {
                            RptToXml.ConvertRptToXml(args[0], args[1], ModelNumber);
                        }
                        catch (Exception e)
                        {
                            Console.Error.WriteLine("Error: Convert to XML error");
                            Console.Error.WriteLine(e);
                            return (int)ExitCode.ConvertError;
                        }
                        Console.WriteLine("File \"" + args[0] + "\" converted to \"" + args[1] + "\"");
                        return (int)ExitCode.Success;
                    }
                default:
                    {
                        Console.Error.WriteLine("Error: Wrong number of parameters");
                        WriteUsage();
                        return (int)ExitCode.WrongArgs;
                    }
            }
        }
        static void WriteUsage()
        {
            Console.WriteLine("Usage: rpt_diff.exe DiffUtilPath RPTPath1 [RPTPath2] ModelNumber");
            Console.WriteLine("       DiffUtilPath - Full path to external diff application .exe file that can compare two xml files (for example KDiff)");
            Console.WriteLine("       RPTPath1 - Full path to first .rpt file to be converted to xml");
            Console.WriteLine("       RPTPath2 - Full path to second .rpt file to be converted to xml and compared with first file");
            Console.WriteLine("       ModelNumber - Select between object model 0 - ReportDocument or 1 (recomended)- ReportClientDocument ");
        }
    }
}

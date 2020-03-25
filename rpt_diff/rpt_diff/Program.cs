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
            switch (args.Length)
            {
                case 2:
                    {
                        if (!File.Exists(args[0]))
                        {
                            Console.Error.WriteLine("Error: Can't find RPT file - Bad RPTPath");
                            WriteUsage();
                            return (int)ExitCode.WrongRptFile;
                        }
                        Console.WriteLine("Converting file: \"" + args[0] + "\"");
                        try
                        {
                            RptToJson.ConvertRptToJson(args[0], args[1]);
                        }
                        catch (Exception e)
                        {
                            Console.Error.WriteLine("Error: Convert to JSON error");
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
            Console.WriteLine("Usage: rpt_diff.exe RPTPath JSONPath");
            Console.WriteLine("       RPTPath - Full path to .rpt file to be converted to json");
            Console.WriteLine("       JSONPath - Full path to where the json file is written to");
        }
    }
}

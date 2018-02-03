using System;
using System.Collections.Generic;
using CommandLine;

namespace PPT2PDF
{
    class Program
    {

        public static Convertion convertion;

        static void Main(string[] args)
        {
            var result = Parser.Default.ParseArguments<Configuration>(args).WithParsed(options =>
            {
                convertion = Convertion.build(options.frameSlides, options.handout, options.verticalFirst, options.onlyPPT);
                convertion.convert(options.inputFiles, options.outputFile);
            });
        }
    }
}

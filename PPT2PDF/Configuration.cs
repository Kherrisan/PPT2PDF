using CommandLine;
using System;
using System.Collections.Generic;
using System.Text;


namespace PPT2PDF
{
    class Configuration
    {
        [Option('i', "input-files", HelpText = "Input ppt or pptx files or directory", Required = true,Min =1, Separator = ' ')]
        public IEnumerable<string> inputFiles { set; get; }

        [Option('n', "number", HelpText = "Number of slides per page.", Default = 6)]
        public int number { set; get; }

        [Option('o', "output-file", HelpText = "Path of the generated PDF file.", Default = "output.pdf")]
        public string outputFile { set; get; }
    }
}

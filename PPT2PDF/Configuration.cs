using CommandLine;
using System;
using System.Collections.Generic;
using System.Text;


namespace PPT2PDF
{
    class Configuration
    {

        [Option('i', "input-files", HelpText = "Input ppt or pptx files or directory", Required = true, Min = 1, Separator = ' ')]
        public IEnumerable<string> inputFiles { set; get; }

        [Option('o', "output-file", HelpText = "Path of the generated PDF file.", Default = "output.pdf")]
        public string outputFile { set; get; }

        [Option('v', "vertical-first", HelpText = "The order in which the handout should be printed.", Default = false)]
        public bool verticalFirst { set; get; }

        public enum Handout
        {
            TWO, THREE, SIX, FOUR, NINE, SINGLE
        }

        [Option('h', "handout", HelpText = "A value that indicates how many slides to be printed in one page.", Default = Handout.SIX)]
        public Handout handout { set; get; }

        [Option('f', "frame-slides", HelpText = "Whether the slides to be exported should be bordered by a frame.", Default = false)]
        public bool frameSlides { set; get; }

        [Option('p', "only-output-ppt", HelpText = "Output merged PPT instead of PDF file.", Default = false)]
        public bool onlyPPT { set; get; }
    }
}

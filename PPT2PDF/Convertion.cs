using System;
using System.IO;
using System.Collections.Generic;
using System.Text;
using MSPowerPoint = Microsoft.Office.Interop.PowerPoint;
using MSCore = Microsoft.Office.Core;
using ShellProgressBar;

namespace PPT2PDF
{
    class Convertion
    {
        private MSCore.MsoTriState frameSlides;
        private MSPowerPoint.PpPrintHandoutOrder handoutOrder;
        private MSPowerPoint.PpPrintOutputType outputType;
        private bool onlyPPT;

        public static ProgressBarOptions progressBarOptions = new ProgressBarOptions
        {
            ProgressCharacter = '-',
            DisplayTimeInRealTime = true,
            ProgressBarOnBottom = true,
            BackgroundColor = ConsoleColor.Green,
            ForeGroundColor = ConsoleColor.Gray,
            ForeGroundColorDone = ConsoleColor.Yellow,

        };

        public static Convertion build(bool frameSlides, Configuration.Handout handout, bool verticalFirst, bool onlyPPT)
        {
            Convertion convertion = new Convertion();
            convertion.frameSlides = frameSlides ? MSCore.MsoTriState.msoTrue : MSCore.MsoTriState.msoFalse;
            convertion.handoutOrder = verticalFirst ? MSPowerPoint.PpPrintHandoutOrder.ppPrintHandoutVerticalFirst : MSPowerPoint.PpPrintHandoutOrder.ppPrintHandoutHorizontalFirst;
            switch (handout)
            {
                case Configuration.Handout.SINGLE:
                    convertion.outputType = MSPowerPoint.PpPrintOutputType.ppPrintOutputOneSlideHandouts;
                    break;
                case Configuration.Handout.TWO:
                    convertion.outputType = MSPowerPoint.PpPrintOutputType.ppPrintOutputTwoSlideHandouts;
                    break;
                case Configuration.Handout.THREE:
                    convertion.outputType = MSPowerPoint.PpPrintOutputType.ppPrintOutputThreeSlideHandouts;
                    break;
                case Configuration.Handout.FOUR:
                    convertion.outputType = MSPowerPoint.PpPrintOutputType.ppPrintOutputFourSlideHandouts;
                    break;
                case Configuration.Handout.SIX:
                    convertion.outputType = MSPowerPoint.PpPrintOutputType.ppPrintOutputSixSlideHandouts;
                    break;
                case Configuration.Handout.NINE:
                    convertion.outputType = MSPowerPoint.PpPrintOutputType.ppPrintOutputNineSlideHandouts;
                    break;
                default:
                    convertion.outputType = MSPowerPoint.PpPrintOutputType.ppPrintOutputSixSlideHandouts;
                    break;
            }
            convertion.onlyPPT = onlyPPT;
            return convertion;
        }

        public Convertion()
        {

        }

        private string checkFile(string file)
        {
            if (!File.Exists(file))
                return null;
            FileInfo info = new FileInfo(file);
            if (info.Extension.Equals(".ppt") || info.Extension.Equals(".pptx"))
                return info.FullName;
            return null;
        }

        private void handlePPT(MSPowerPoint.Application application, string input, MSPowerPoint.Presentation output, ProgressBar bar)
        {
            if ((input = checkFile(input)) != null)
            {
                output.Slides.InsertFromFile(input, output.Slides.Count);
                bar.Tick();
            }
        }

        private int countFile(string file)
        {
            if (checkFile(file) != null)
                return 1;
            return 0;
        }

        private int countAllFiles(IEnumerable<string> inputs)
        {
            Console.WriteLine("Conting down all PPT files to be merged.");
            int counter = 0;
            foreach (string inputFile in inputs)
            {
                if (Directory.Exists(inputFile))
                {
                    foreach (string eachFile in Directory.EnumerateFiles(inputFile))
                    {
                        counter += countFile(eachFile);
                    }
                }
                else
                {
                    counter += countFile(inputFile);
                }
            }
            return counter;
        }

        public void convert(IEnumerable<string> inputs, string output)
        {
            var app = new MSPowerPoint.Application();
            var outputPPT = app.Presentations.Add(WithWindow: Microsoft.Office.Core.MsoTriState.msoFalse);
            try
            {
                int total = countAllFiles(inputs);
                Console.WriteLine(total + " PPT files need to be merged.");
                using (var bar = new ProgressBar(total, "Emmmm", progressBarOptions))
                {
                    foreach (string inputFile in inputs)
                    {
                        if (Directory.Exists(inputFile))
                        {
                            foreach (string eachFile in Directory.EnumerateFiles(inputFile))
                            {
                                handlePPT(app, eachFile, outputPPT, bar);
                            }
                        }
                        else
                        {
                            handlePPT(app, inputFile, outputPPT, bar);
                        }
                    }
                }
                Console.WriteLine("Finished merging all PPT files.");
                if (onlyPPT)
                {
                    Console.WriteLine("Now generate target PPT file.");
                    outputPPT.SaveAs(output);
                }
                else
                {
                    Console.WriteLine("Now generate target PDF file.Progress will be shown in a new dialog prompted by MSOffice.");
                    outputPPT.ExportAsFixedFormat2(output,
                    MSPowerPoint.PpFixedFormatType.ppFixedFormatTypePDF,
                    MSPowerPoint.PpFixedFormatIntent.ppFixedFormatIntentPrint,
                    frameSlides,
                    handoutOrder,
                    outputType,
                    Microsoft.Office.Core.MsoTriState.msoFalse);
                }
            }
            catch (Exception e)
            {

            }
            finally
            {
                outputPPT.Close();
                app.Quit();
                Console.WriteLine("Finished all.");
            }
        }
    }
}

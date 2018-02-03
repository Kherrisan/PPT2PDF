using System;
using System.IO;
using System.Collections.Generic;
using System.Text;
using MSPowerPoint = Microsoft.Office.Interop.PowerPoint;
using MSCore = Microsoft.Office.Core;

namespace PPT2PDF
{
    class Convertion
    {
        private MSCore.MsoTriState frameSlides;
        private MSPowerPoint.PpPrintHandoutOrder handoutOrder;
        private MSPowerPoint.PpPrintOutputType outputType;
        private bool onlyPPT;

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

        private void handlePPT(MSPowerPoint.Application application, string input, MSPowerPoint.Presentation output)
        {
            try
            {
                output.Slides.InsertFromFile(input, output.Slides.Count);
            }
            catch (Exception e)
            {

            }
            finally
            {
                //inputPPT.Close();
            }
        }

        public void convert(IEnumerable<string> inputs, string output)
        {
            var app = new MSPowerPoint.Application();
            var outputPPT = app.Presentations.Add(WithWindow: Microsoft.Office.Core.MsoTriState.msoFalse);
            try
            {
                foreach (string inputFile in inputs)
                {
                    if (Directory.Exists(inputFile))
                    {
                        foreach (string eachFile in Directory.EnumerateFiles(inputFile))
                        {
                            handlePPT(app, eachFile, outputPPT);
                        }
                    }
                    else
                    {
                        FileInfo fileInfo = new FileInfo(inputFile);
                        if (fileInfo.Extension.Equals(".ppt") || fileInfo.Extension.Equals(".pptx"))
                        {
                            handlePPT(app, fileInfo.FullName, outputPPT);
                        }
                    }
                }
                Console.WriteLine("Finished merging all ppt files.")
                if (onlyPPT)
                {
                    Console.WriteLine("Finished merging all ppt files.")
                    outputPPT.SaveAs(output);
                }
                else
                {
                    Console.WriteLine("Finished merging all ppt files.")
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
                Console.WriteLine("Finished all.")
            }
        }
    }
}

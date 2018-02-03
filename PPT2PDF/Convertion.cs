using System;
using System.IO;
using System.Collections.Generic;
using System.Text;
using MSPowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace PPT2PDF
{
    class Convertion
    { 
        private int number;

        public static Convertion build(int number)
        {
            Convertion convertion = new Convertion(number);
            return convertion;
        }

        public Convertion(int number)
        {
            this.number = number;
        }

        private void handlePPT(string input,MSPowerPoint.Presentation output)
        {
            output.Slides.InsertFromFile(input, output.Slides.Count);
        }

        public void convert(IEnumerable<string> inputs,string output)
        {
            var app = new MSPowerPoint.Application();
            var outputPPT = app.Presentations.Add();
            app.Visible = false;
            try
            { 
                foreach(string inputFile in inputs)
                {
                    if (Directory.Exists(inputFile))
                    {
                        foreach(string eachFile in Directory.EnumerateFiles(inputFile))
                        {
                            handlePPT(eachFile, outputPPT);
                        }
                    }
                    else
                    {
                        FileInfo fileInfo = new FileInfo(inputFile);
                        if (fileInfo.Extension.Equals("ppt") || fileInfo.Extension.Equals("pptx"))
                        {
                            handlePPT(inputFile, outputPPT);
                        }
                    }
                }
            }
            catch (Exception e)
            {
                
            }
            finally
            {
                app.Quit();
            }
        }
    }
}

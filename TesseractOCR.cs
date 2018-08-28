using System;
using Tesseract;


namespace SFBTesting.LibraryFunctions
{
    public  class Tesseract_OCR
    {
        private  string dataPath = @"C:\SFB_Automation\DataFiles\tessdata";
        private  string text = null;
        
        public  string ImageExtraction(string ImagePath)
        {
            try
            {
                TesseractEnviornment.CustomSearchPath = @"C:\SFB_Automation";
                using (var tEngine = new TesseractEngine(dataPath, "eng", EngineMode.Default)) //creating the tesseract OCR engine with English as the language
                {
                    using (var img = Pix.LoadFromFile(ImagePath)) // Load of the image file from the Pix object which is a wrapper for Leptonica PIX structure
                    {
                        using (var page = tEngine.Process(img)) //process the specified image
                        {
                             text= page.GetText();
                           //Gets the image's content as plain text.
                        }                       
                    }
                   
                }             
            }
            catch (Exception e)
            {
                throw e;
            }
            return text;
        }
    }
}

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Acrobat;
using System.IO;
using System.Diagnostics;
using AFORMAUTLib;

namespace Adhoc_Adobe_Library_3_5
{
    /// <summary>
    /// custom class for opening, saving, merging, and splitting a specified PDF document
    /// </summary>
    public class PDFDoc
    {
        private CAcroAVDoc pdf;
        private AcroApp app;
        private string defaultPath;
        private string defaultFileName;
        private string originalFilePathFull;
        private readonly string[] fileNameAddons = new string[] { "new", "newer", "newest", "newerest", "mostNewest", "theMostNewerest", "evenMoreNewerest", "thisIsGettingRediculous", "nowYouAreJustScrewingWithMe", "reallyQuestionMark", "fineYoureGettingCloseToTheEndHere", "imSeriousDontTryAndSaveThisFileAgainWithoutDeletingSomething", "lastWarning" };
       
        #region cosntructor, closeDoc, dispose, basic functions
        /// <summary>
        /// base constructor, filePath must be valid system path to a PDF document
        /// </summary>
        /// <param name="filePath">path of file to load</param>
        public PDFDoc(string filePath)
        {
            if (File.Exists(filePath))
            {
                this.pdf = new AcroAVDoc();
                this.app = new AcroApp();
                this.app.Hide();
                this.pdf.Open(filePath, Path.GetFileName(filePath));
                this.pdf.SetTitle(Path.GetFileName(filePath));
                this.defaultPath = Path.GetDirectoryName(filePath) + @"\";
                this.defaultFileName = Path.GetFileNameWithoutExtension(filePath);
                this.originalFilePathFull = filePath;
            }
            else
            {
                throw new PDFDocException("The Specified File Path is not Valid");
            }
        }

        /// <summary>
        /// returns total number of pages in the loaded document
        /// </summary>
        /// <returns></returns>
        public int getPageCount()
        {
            try
            {
                CAcroPDDoc doc = (AcroPDDoc)this.pdf.GetPDDoc();
                return doc.GetNumPages();
            }
            catch (Exception e)
            {
                
                throw new PDFDocException("Unable To Get Page Count, base error message: " + e.Message);
            }
        }

        /// <summary>
        /// closes the loaded document, does not save
        /// </summary>
        public void closeDoc()
        {
            try
            {
                this.app.CloseAllDocs();
                //this.pdf.Close(1);
            }
            catch (Exception e)
            {
                throw new PDFDocException("Unable To Close Document, base error message: " + e.Message);
            }
        }

        /// <summary>
        /// save current loaded document to specificed path, extension must be .pdf or .PDF
        /// </summary>
        /// <param name="path"></param>
        public void saveAs(string path)
        {
            if (Path.GetExtension(path).Equals(".pdf") || Path.GetExtension(path).Equals(".PDF"))
            {
                try
                {
                    CAcroPDDoc doc = (AcroPDDoc)this.pdf.GetPDDoc();
                    doc.Save(1, path);
                }
                catch (Exception e)
                {
                    throw new PDFDocException("Unable To Save Document, base error message: " + e.Message);
                }
            }
            else
            {
                throw new PDFDocException("Extension of the save path must be .pdf or .PDF");
            }            
        }

        /// <summary>
        /// returns the total word count for a specified page
        /// </summary>
        /// <param name="pageNumber">the 1-based index of the page to get the count for</param>
        /// <returns></returns>
        public int getWordCountForPage(int pageNumber)
        {
            //get doc and json 
            CAcroPDDoc doc = (CAcroPDDoc)this.pdf.GetPDDoc();
            object jsObj = doc.GetJSObject();
            //fix pageNumber to be 0 index based
            pageNumber--;

            object[] param = new object[] { pageNumber.ToString() };
            Type t = jsObj.GetType();
            object count = t.InvokeMember("getPageNumWords", System.Reflection.BindingFlags.InvokeMethod | System.Reflection.BindingFlags.Public |
                    System.Reflection.BindingFlags.Instance, null, jsObj, param);
            string temp = count.ToString();
            return Int32.Parse(temp);
        }
        #endregion

        #region merge and add pages functions
        /// <summary>
        /// merges current doc with specified file 
        /// </summary>
        /// <param name="pageToInsertAfter">where to begin inserting specified pages, 0 merges before 1st page of loaded document</param>
        /// <param name="pathToInsertFrom">file path to insert pages from</param>
        /// <param name="directoryToSaveTo">directory output should be save to, if not saving pass in blank string</param>
        /// <param name="startPageForInsertPDF">page start of the insert doc</param>
        /// <param name="finalPageForInsertPDF">page end of the doc</param>
        /// <param name="saveAsFileName">file name to use to save doc, does not need extension, if not saving pass in blank string</param>
        /// <param name="save">indicates if doc should save after merging</param>
        public void mergePDFS(int pageToInsertAfter, string pathToInsertFrom, string directoryToSaveTo, int startPageForInsertPDF, int finalPageForInsertPDF, string saveAsFileName, bool save)
        {
            try
            {
                //make sure last char of saveAsPathBase is '\'
                if (directoryToSaveTo.Length>0 && !directoryToSaveTo.Substring(directoryToSaveTo.Length - 1).Equals(@"\"))
                {
                    directoryToSaveTo += @"\";
                }

                //get doc and json object
                CAcroPDDoc doc = (AcroPDDoc)this.pdf.GetPDDoc();
                object jsObject = doc.GetJSObject();

                if (finalPageForInsertPDF == 0)
                {
                    finalPageForInsertPDF = startPageForInsertPDF;
                }
                /*
                 *  nPage - 0-based index of the page after which to insert source doc, -1 inserts before first page
                 *  cPath - the path of the doc to insert from
                 *  nStart - 0-based inclusive index of page to start inserting from the source doc
                 *  nEnd - 0-based inclusive index of pages to stop inserting from the source doc
                 * */

                object[] param = new object[] { pageToInsertAfter - 1, pathToInsertFrom, startPageForInsertPDF - 1, finalPageForInsertPDF - 1 };
                Type t = jsObject.GetType();
                t.InvokeMember("insertPages", System.Reflection.BindingFlags.InvokeMethod | System.Reflection.BindingFlags.Public |
                        System.Reflection.BindingFlags.Instance, null, jsObject, param);

                if (save)
                {

                    //save the new doc
                    string filePath = directoryToSaveTo + saveAsFileName + ".pdf";

                    if (File.Exists(filePath))
                    {
                        filePath = getSafeFileNameToSaveTo(filePath, 0);
                    }

                    doc.Save(1, filePath);
                }

            }
            catch (Exception e)
            {
                throw new PDFDocException("Could Not Combine Files.  Ensure the Integrity and File Paths of All Documents, base exception: " + e.Message);
            }
        }

        /// <summary>
        /// inserts a blank page after the specified point in the loaded doc, use 0 to indicate insert at the front of the doc
        /// </summary>
        /// <param name="pageToInsertBlankAfter">1-based page number to insert new page after</param>
        public void addBlankPageToDoc(int pageToInsertBlankAfter)
        {
            CAcroPDDoc doc = (AcroPDDoc)this.pdf.GetPDDoc();
            object jsObj = doc.GetJSObject();

            CAcroPDPage page = (AcroPDPage)doc.AcquirePage(0);
            CAcroPoint points = (AcroPoint)page.GetSize();

            /**
             * nPage - page after which to insert new page, 1 based index
             * nWidth - width of new page
             * nHeight - height of new page
             * */
            object[] param = new object[] { pageToInsertBlankAfter, points.x, points.y };
            Type t = jsObj.GetType();
            t.InvokeMember("newPage", System.Reflection.BindingFlags.InvokeMethod | System.Reflection.BindingFlags.Public |
                        System.Reflection.BindingFlags.Instance, null, jsObj, param);

        }
        #endregion

        #region split methods
        /// <summary>
        /// splits specified pages from the loaded document
        /// </summary>
        /// <param name="startPage"></param>
        /// <param name="numberOfPagesToSplit"></param>
        /// <param name="pathToSaveSplitPages"></param>
        /// <returns>true if split was succesfull, false if page range was not valid</returns>
        public bool splitPagesFromDocument(int startPage, int numberOfPagesToSplit, string pathToSaveSplitPages)
        {
            //make sure we aren't going to grab a page outside of the doc length
            if (startPage + numberOfPagesToSplit - 1 <= this.getPageCount())
            {
                //get doc and json object
                CAcroPDDoc doc = (AcroPDDoc)this.pdf.GetPDDoc();
                object jsObject = doc.GetJSObject();

                /**
                * nStart = 0-based index that defines the start range of pages to extract
                * nEnd = 0-based index that defines the end of the range of pages to extract
                * cPath = device-independent path to save new doc
                * */
                int zeroBasedStartPage = startPage - 1;
                object[] param = new object[]{ zeroBasedStartPage, zeroBasedStartPage + numberOfPagesToSplit - 1, pathToSaveSplitPages };
                Type t = jsObject.GetType();
                t.InvokeMember("ExtractPages", System.Reflection.BindingFlags.InvokeMethod | System.Reflection.BindingFlags.Public |
                    System.Reflection.BindingFlags.Instance, null, jsObject, param);

                return true;
            }
            else
            {
                //throw new PDFDocException("Start Page + Number of Pages to Split is outside the range of the current PDF");
                return false;
            }
        }

        /// <summary>
        /// splits loaded document 
        /// </summary>
        /// <param name="startPage">page to begin spliting</param>
        /// <param name="numberOfPagesToSplit">total number of pages per record to split out</param>
        /// <param name="saveToPathBase">base path where to save, _ and record number will be appended to this</param>
        /// <param name="formatedRecordNumberLength">total length of the foramted record number to appended</param>
        /// <param name="recordStart">which record to begin the count at</param>
        /// <returns>recordStart incremented to the total number of records split out of document</returns>
        public int splitDocument(int startPage, int numberOfPagesToSplit, string saveToPathBase, int formatedRecordNumberLength, int recordStart, bool appendRecordNumber)
        {
            //get doc and json object
            CAcroPDDoc doc = (AcroPDDoc)this.pdf.GetPDDoc();
            object jsObject = doc.GetJSObject();

            int pageCount = this.getPageCount();

            if (startPage + numberOfPagesToSplit - 1 <= pageCount)
            {
                /**
                 * nStart = 0-based index that defines the start range of pages to extract
                 * nEnd = 0-based index that defines the end of the range of pages to extract
                 * cPath = device-independent path to save new doc
                 * */
                int zeroBasedStartPage = startPage - 1;

                while (zeroBasedStartPage + numberOfPagesToSplit <= pageCount)
                {
                    string saveAsPath;
                    if (appendRecordNumber)
                    {
                        saveAsPath = saveToPathBase + "_" + this.getRecordNumberFormatedWithLeadingZeros(formatedRecordNumberLength, recordStart) + ".pdf";
                    }
                    else
                    {
                        saveAsPath = saveToPathBase + ".pdf";
                    }
                     
                    object[] param = { zeroBasedStartPage, zeroBasedStartPage + numberOfPagesToSplit - 1, saveAsPath };
                    Type t = jsObject.GetType();
                    t.InvokeMember("ExtractPages", System.Reflection.BindingFlags.InvokeMethod | System.Reflection.BindingFlags.Public |
                        System.Reflection.BindingFlags.Instance, null, jsObject, param);
                    zeroBasedStartPage += numberOfPagesToSplit;
                    recordStart++;
                }
            }
            else
            {
                throw new PDFDocException("Start Page + Number of Pages to Split is outside the range of the current PDF");
            }

            return recordStart;
        }

        /// <summary>
        /// splits the loaded document into new pdf docs based on the counts in the queue
        /// </summary>
        /// <param name="pageCountQueue">queue to determine how many pages should be split out per record</param>
        /// <param name="saveAsPathBase">base path to save each pdf to, _ and record number will be appended to this</param>
        /// <param name="formatRecordNumberLength">total length of the foramted record number to appended</param>
        public void splitDocumentWithVaryingPageLengths(Queue<int> pageCountQueue, string saveAsPathBase, int formatRecordNumberLength)
        {
            //get doc and json object
            CAcroPDDoc doc = (AcroPDDoc)this.pdf.GetPDDoc();
            object jsObject = doc.GetJSObject();
            int rec = 1;

            int startPage = 1;

            while (pageCountQueue.Count > 0)
            {
                int pageLength = pageCountQueue.Dequeue();
                string saveAsPath = saveAsPathBase + "_" + this.getRecordNumberFormatedWithLeadingZeros(formatRecordNumberLength, rec) + ".pdf";
                this.splitPagesFromDocument(startPage, pageLength, saveAsPath);
                startPage += pageLength;
                rec++;
            }
        }
        #endregion

        #region save functions

        /// <summary>
        /// saves loaded doc to eps, uses same path and filename as current loaded doc
        /// </summary>
        public void saveToEPS()
        {
            this.saveToEPS(this.defaultPath, this.defaultFileName);
        }

        /// <summary>
        /// flattens the PDF by converting all annotations and fields to text
        /// </summary>
        public void flattenPDF()
        {
            CAcroPDDoc doc = (AcroPDDoc)this.pdf.GetPDDoc();
            object jsObj = doc.GetJSObject();
            Type t = jsObj.GetType();
            object[] paramaters = new object[] { 0, this.getPageCount() -1, 0 };
            t.InvokeMember("flattenPages", System.Reflection.BindingFlags.InvokeMethod | System.Reflection.BindingFlags.Public |
                    System.Reflection.BindingFlags.Instance, null, jsObj, paramaters);
        }

        /// <summary>
        /// saves loaded doc to eps
        /// </summary>
        /// <param name="directoryToSaveTo">where to save eps files</param>
        /// <param name="fileName">filename base for the eps</param>
        public void saveToEPS(string directoryToSaveTo, string fileName)
        {
            try
            {
                //make sure last char of saveAsPathBase is '\'
                if (!directoryToSaveTo.Substring(directoryToSaveTo.Length - 1).Equals(@"\"))
                {
                    directoryToSaveTo += @"\";
                }

                //get doc and json object
                CAcroPDDoc doc = (AcroPDDoc)this.pdf.GetPDDoc();
                object jsObject = doc.GetJSObject();

                //set the destination string
                string destination = directoryToSaveTo + fileName + ".eps";

                //check if the output files exist, if so delete since overwrite does not work
                int count = this.getPageCount();
                deleteEPSFilesForOverwrite(count, destination, fileName, directoryToSaveTo);

                //code i found for dealing with COM objects
                /*
                 * SaveAs method params - 
                 *      cPath - path to save at
                 *      cconvID - conversion ID, defaults to PDF
                 *      cFS - source file system name, either ""(default) or "CHTTP"
                 *      bCopy - save the PDf file as copy, boolean, defaults to false
                 *      bPromptToOverwrite - prompts user to overwrite, boolean, defaults to false
                 * */
                object[] saveAsParam = { destination, "com.adobe.acrobat.eps", "", false, true };
                Type T = jsObject.GetType();
                T.InvokeMember("saveAs", System.Reflection.BindingFlags.InvokeMethod | System.Reflection.BindingFlags.Public |
                    System.Reflection.BindingFlags.Instance, null, jsObject, saveAsParam);
            }
            catch (Exception e)
            {
                throw new PDFDocException(e.Message);
            }
        }

        /// <summary>
        /// saves loaded doc to PNG, saves in same directory and with the same name as the loaded doc
        /// </summary>
        public void saveToPNG()
        {
            this.saveToPNG(this.defaultPath, this.defaultFileName);
        }

        /// <summary>
        /// save loaded doc to PNG
        /// </summary>
        /// <param name="directoryToSaveTo">where to save to</param>
        /// <param name="fileName">file name for each PNG</param>
        public void saveToPNG(string directoryToSaveTo, string fileName)
        {
            try
            {
                //get doc and json object
                CAcroPDDoc doc = (AcroPDDoc)this.pdf.GetPDDoc();
                object jsObject = doc.GetJSObject();

                //set the destination string
                string destination = directoryToSaveTo + fileName + ".png";

                //delete old files
                int count = this.getPageCount();
                deleteFilesForOverwrite(count, destination, fileName, directoryToSaveTo, ".png");

                //code i found for dealing with COM objects
                object[] saveAsParam = { destination, "com.adobe.acrobat.png" };
                Type T = jsObject.GetType();
                T.InvokeMember("SaveAs", System.Reflection.BindingFlags.InvokeMethod | System.Reflection.BindingFlags.Public |
                    System.Reflection.BindingFlags.Instance, null, jsObject, saveAsParam);
            }
            catch (Exception e)
            {

                throw new PDFDocException(e.Message);
            }
        }

        /// <summary>
        /// saves loaded doc to JPG, saves in same directory and with the same name as the loaded doc
        /// </summary>
        public void saveToJPG()
        {
            this.saveToJPG(this.defaultPath, this.defaultFileName);
        }

        /// <summary>
        /// saves loaded doc to JPG
        /// </summary>
        /// <param name="directoryToSaveTo">where to save JPGs</param>
        /// <param name="fileName">base name for each JPG saved</param>
        public void saveToJPG(string directoryToSaveTo, string fileName)
        {
            try
            {
                //get doc and json object
                CAcroPDDoc doc = (AcroPDDoc)this.pdf.GetPDDoc();
                object jsObject = doc.GetJSObject();

                //set the destination string
                string destination = directoryToSaveTo + fileName + ".jpeg";

                //delete old files
                int count = this.getPageCount();
                deleteFilesForOverwrite(count, destination, fileName, directoryToSaveTo, ".jpeg");

                //code i found for dealing with COM objects
                object[] saveAsParam = { destination, "com.adobe.acrobat.jpeg" };
                Type T = jsObject.GetType();
                T.InvokeMember("SaveAs", System.Reflection.BindingFlags.InvokeMethod | System.Reflection.BindingFlags.Public |
                    System.Reflection.BindingFlags.Instance, null, jsObject, saveAsParam);
            }
            catch (Exception e)
            {
                throw new PDFDocException(e.Message);
            }
        }

        /// <summary>
        /// save loaded doc to TIFF, saves in same directory and with the same filename as the loaded doc
        /// </summary>
        public void saveToTIFF()
        {
            this.saveToTIFF(this.defaultPath, this.defaultFileName);
        }

        /// <summary>
        /// saves the loaded doc to TIFF
        /// </summary>
        /// <param name="directoryToSaveTo">directory to save TIFFs to</param>
        /// <param name="fileName">filename base for each TIFF</param>
        public void saveToTIFF(string directoryToSaveTo, string fileName)
        {
            try
            {
                //get doc and json object
                CAcroPDDoc doc = (AcroPDDoc)this.pdf.GetPDDoc();
                object jsObject = doc.GetJSObject();

                //set the destination string
                string destination = directoryToSaveTo + fileName + ".tiff";

                //delete old files
                int count = this.getPageCount();
                deleteFilesForOverwrite(count, destination, fileName, directoryToSaveTo, ".tiff");

                //code i found for dealing with COM objects
                object[] saveAsParam = { destination, "com.adobe.acrobat.tiff" };
                Type T = jsObject.GetType();
                T.InvokeMember("SaveAs", System.Reflection.BindingFlags.InvokeMethod | System.Reflection.BindingFlags.Public |
                    System.Reflection.BindingFlags.Instance, null, jsObject, saveAsParam);
            }
            catch (Exception e)
            {
                throw new PDFDocException(e.Message);
            }
            
        }
        #endregion

        #region private methods
        private void deleteEPSFilesForOverwrite(int count, string destination, string fileName, string directoryToSaveTo)
        {
            string epsSpaces = "";
            if (count > 10 && count <= 100)
            {
                epsSpaces = " ";
            }
            else if (count > 100 && count <= 1000)
            {
                epsSpaces = "  ";
            }
            else if (count > 1000 && count <= 10000)
            {
                epsSpaces = "   ";
            }
            else if (count > 10000)
            {
                epsSpaces = "    ";
            }

            for (int i = 1; i <= count; i++)
            {
                if (count == 1)
                {
                    if (File.Exists(destination))
                    {
                        File.Delete(destination);
                    }
                }
                else
                {
                    //remove a space at this point
                    if ((i == 10 || i == 100 || i == 1000 || i == 10000) && epsSpaces.Length > 0)
                    {
                        epsSpaces = epsSpaces.Substring(0, epsSpaces.Length - 1);
                    }

                    if (File.Exists(directoryToSaveTo + fileName + "_" + epsSpaces + i + ".eps"))
                    {
                        File.Delete(directoryToSaveTo + fileName + "_" + epsSpaces + i + ".eps");
                    }
                }
            }
        }

        private void deleteFilesForOverwrite(int count, string destination, string fileName, string directoryToSaveTo, string extension)
        {
            string leadingZeros = "";
            if (count > 10 && count <= 100)
            {
                leadingZeros = "0";
            }
            else if (count > 100 && count <= 1000)
            {
                leadingZeros = "00";
            }
            else if (count > 1000 && count <= 10000)
            {
                leadingZeros = "000";
            }
            else if (count > 10000)
            {
                leadingZeros = "0000";
            }
            
            for (int i = 1; i <= count; i++)
            {
                if ((i == 10 || i == 100 || i == 1000 || i == 10000)&& leadingZeros.Length > 0)
                {
                    leadingZeros = leadingZeros.Substring(0, leadingZeros.Length - 1);
                }
                if (File.Exists(directoryToSaveTo + fileName + "_Page_" + leadingZeros + i + extension))
                {
                    File.Delete(directoryToSaveTo + fileName + "_Page_" + leadingZeros + i + extension);
                }
            }
        }

        private string getRecordNumberFormatedWithLeadingZeros(int totalStringLength, int record)
        {
            string recordString = record.ToString();
            while (recordString.Length < totalStringLength)
            {
                recordString = "0" + recordString;
            }
            return recordString;
        }

        private string getSafeFileNameToSaveTo(string filePath, int tries)
        {            
            //build the new path
            string name = Path.GetFileNameWithoutExtension(filePath);
            string path = Path.GetDirectoryName(filePath) + @"\";
            string newFilePath = path + name + "_" + this.fileNameAddons[tries] + ".pdf";

            while (File.Exists(newFilePath))
            {
                newFilePath = getSafeFileNameToSaveTo(filePath, tries + 1);
            }

            return newFilePath;
        }
        #endregion

        #region preflight methods
        /// <summary>
        /// run preflight droplet for embeding fonts, saves over the loaded doc
        /// </summary>
        public void embedFonts()
        {
            try
            {
                //CAcroPDDoc doc = (AcroPDDoc)this.pdf.GetPDDoc();
                //object jsObj = doc.GetJSObject();

                //object[] param = new object[] { "embedFont",
                //"var oProfile = Preflight.getProfileByName(\"Embed Fonts\"); if (oProfile != undefined) {this.preflight(oProfile);}"                
                //};
                //Type t = jsObj.GetType();
                //t.InvokeMember("addScript", System.Reflection.BindingFlags.InvokeMethod | System.Reflection.BindingFlags.Public |
                //    System.Reflection.BindingFlags.Instance, null, jsObj, param);
                Process.Start(@"\\oh50ms05\Development\Preflight_Droplets\Embed Fonts.exe", this.defaultPath + this.defaultFileName + ".pdf");
            }
            catch (Exception e)
            {
                throw new PDFDocException("Could Not Embed Fonts.  Base error: " + e.Message);
            }
        }

        public void embedFonts(string[] filesToEmbedFonts)
        {
            try
            {
                ProcessStartInfo startInfo = new ProcessStartInfo();
                
                string command = "";
                foreach (string s in filesToEmbedFonts)
                {
                    command = command + " " + s;                     
                }
                startInfo.Arguments = command;
                startInfo.FileName = @"\\oh50ms05\Development\Preflight_Droplets\Embed Fonts.exe";
                Process.Start(startInfo);
                //Process p = Process.Start(@"\\oh50ms05\Development\Preflight_Droplets\Embed Fonts.exe", command);
                
            }
            catch (Exception e)
            {
                throw new PDFDocException("Could Not Embed Fonts.  Base error: " + e.Message);
            }
        }

        /// <summary>
        /// run preflight droplet for covnerting to greyscale, saves over the loaded doc
        /// </summary>
        public void convertToGreyScale()
        {
            try
            {
                Process.Start(@"\\oh50ms05\Development\Preflight_Droplets\Convert to grayscale.exe", this.defaultPath + this.defaultFileName + ".pdf");
            }
            catch (Exception e)
            {
                throw new PDFDocException("Could Not Convert To Gre(a)yscale.  Base error: " + e.Message);
            }
        }
        #endregion

        #region text reading methods

        /// <summary>
        /// returns the text of a page in a specified rectangle 
        /// </summary>
        /// <param name="left_x_inches">left most x point in inches</param>
        /// <param name="right_x_inches">right most x point in inches</param>
        /// <param name="top_y_inches">top most y point in inches</param>
        /// <param name="bottom_y_inches">bottom most y point in inches</param>
        /// <param name="page">page to read</param>
        /// <returns>string value of text iside the area of the rectangle</returns>
        public string getTextInSpecifiedArea(double left_x_inches, double right_x_inches, double top_y_inches, double bottom_y_inches, int page)
        {
            //covert from inches to PDF units 72units/inch
            short left_x = (short)(left_x_inches * 72);
            short right_x =(short)(right_x_inches * 72);
            short top_y = (short)(top_y_inches * 72);
            short bottom_y = (short)(bottom_y_inches * 72);
            string text = "";

            //make the selection rectangle
            CAcroRect rect = new AcroRect();
            rect.Left = left_x;
            rect.right = right_x;
            rect.Top = top_y;
            rect.bottom = bottom_y;

            //get doc and make selection
            CAcroPDDoc doc = (AcroPDDoc)this.pdf.GetPDDoc();
            CAcroPDTextSelect result = (AcroPDTextSelect)doc.CreateTextSelect(page-1, rect);
            if (result != null)//successfully made selection
            {
                int count = result.GetNumText();
                for (int i = 0; i < count; i++)
                {
                    text += result.GetText(i);
                }
            }
            return text;
        }

        /// <summary>
        /// determines if the supplied text is within the text of the specified page
        /// </summary>
        /// <param name="page">the page to search</param>
        /// <param name="text">the text to determine if it is present</param>
        /// <param name="matchCase">whether the text must match with regardes to letter casing</param>
        /// <returns>bool if the text is on the page</returns>
        public bool findTextOnPage(int page, string text, bool matchCase)
        {
            bool found = false;
            //get doc 
            CAcroPDDoc doc = (AcroPDDoc)this.pdf.GetPDDoc();
            object jsObj = doc.GetJSObject();
            object[] param = new object[] { page };
            Type t = jsObj.GetType();
            //get the word count on the page
            double count = (double)t.InvokeMember("getPageNumWords", System.Reflection.BindingFlags.InvokeMethod | System.Reflection.BindingFlags.Public | System.Reflection.BindingFlags.Instance,
                null, jsObj, param);
            //iterate over each word and compare to the supplied text
            for (int i = 0; i < count; i++)
            {
                object[] wordIndex = new object[] { page, i, true };
                object w = t.InvokeMember("getPageNthWord", System.Reflection.BindingFlags.InvokeMethod | System.Reflection.BindingFlags.Public | System.Reflection.BindingFlags.Instance,
                    null, jsObj, wordIndex);
                string word = w.ToString();

                if (matchCase)
                {
                    found = text.Equals(word);
                }
                else
                {
                    string uWord = word.ToUpper();
                    string uText = text.ToUpper();
                    found = uText.Equals(uWord);
                }
                if (found)
                {
                    break;
                }
            }
            return found;
        }
        #endregion

        #region text writing methods
        /// <summary>
        /// creates a field at the specified position and writes the text to the given pt size and font
        /// </summary>
        /// <param name="text">the text to write</param>
        /// <param name="textSize">point size of text</param>
        /// <param name="font">desired font</param>
        /// <param name="left_x_coord">top left x coord of bounding rectangle</param>
        /// <param name="left_y_coord">top left y coord of bounding rectangle</param>
        /// <param name="right_x_coord">top right x coord of bounding rectangle</param>
        /// <param name="right_y_coord">top right y coord of bounding rectangle</param>
        /// <param name="page">the page to right the text to</param>
        public void writeTextToForm(string text, int textSize, string font, double left_x_coord, double left_y_coord, double right_x_coord, double right_y_coord, int page)
        {
            int zeroBasedPageIndex = page - 1;
            //covert from inches to PDF units 72units/inch
            float left_x = (float)(left_x_coord * 72);
            float left_y = (float)(left_y_coord * 72);
            float right_x = (float)(right_x_coord * 72);
            float right_y = (float)(right_y_coord * 72);
            object[] rectArr = new object[] { left_x, left_y, right_x, right_y };
            string guid = new Guid().ToString();
            object[] param1 = new object[] { guid, "text", zeroBasedPageIndex, rectArr };

            //get json object
            CAcroPDDoc doc = (AcroPDDoc)this.pdf.GetPDDoc();
            object jsObj = doc.GetJSObject();
            Type t = jsObj.GetType();
            t.InvokeMember("addField", System.Reflection.BindingFlags.InvokeMethod | System.Reflection.BindingFlags.Public | System.Reflection.BindingFlags.Instance,
                null, jsObj, param1);

            IAFormApp app = new AFormAppClass();

            IFields fields = (IFields)app.Fields;
            //fields.Add("data", "TEXT", (short)page, left_x, left_y, right_x, right_y);
            fields = (IFields)app.Fields;
            System.Collections.IEnumerator myEnumerator = fields.GetEnumerator();
            while (myEnumerator.MoveNext())
            {
                IField myField = (IField)myEnumerator.Current;
                if (myField.Name.Equals(guid))
                {
                    myField.Value = text;
                    myField.TextFont = font;
                    myField.TextSize = (short)textSize;
                }
            }
        }

        public void add2DToPDf(string imagePath, double top_x, double top_y, int page)
        {
            int zeroBasedPageIndex = page - 1;
            ////covert from inches to PDF units 72units/inch
            //float left_x = (float)(top_x * 72);
            //float left_y = (float)(top_y * 72);
            //float right_x = (float)((top_x + .25) * 72);
            //float right_y = (float)((top_y -.25) * 72);
            //object[] rectArr = new object[] { left_x, left_y, right_x, right_y };
            //string guid = new Guid().ToString();
            object[] param1 = new object[] { imagePath, 0, zeroBasedPageIndex, zeroBasedPageIndex, true, true, true, 0, 3, 10, -10, false, 0, 1.0  };
            //addWatermarkFromFile(SampleImageFilePath, 0, 0, 0, True, True, True, 0, 3, 10, -10, False, 0.4, False, 0, 0.7)
            ////get json object
            CAcroPDDoc doc = (AcroPDDoc)this.pdf.GetPDDoc();
            object jsObj = doc.GetJSObject();
            Type t = jsObj.GetType();
            t.InvokeMember("addWatermarkFromFile", System.Reflection.BindingFlags.InvokeMethod | System.Reflection.BindingFlags.Public | System.Reflection.BindingFlags.Instance,
                null, jsObj, param1);

            //IAFormApp app = new AFormAppClass();

            //IFields fields = (IFields)app.Fields;
            //fields = (IFields)app.Fields;
            //System.Collections.IEnumerator myEnumerator = fields.GetEnumerator();
            //while (myEnumerator.MoveNext())
            //{
            //    IField myField = (IField)myEnumerator.Current;
            //    if (myField.Name.Equals(guid))
            //    {
            //        myField.SetButtonIcon("N", imagePath, (short)page);
            //    }
            //}            
        }

        //public void writePremergeDataToFields(string dataPath, int row)
        //{
        //    IAFormApp formApp = new AFormAppClass();
        //    IFields myFields = (IFields)formApp.Fields;
        //    System.Collections.IEnumerator enumerator = myFields.GetEnumerator();
        //    int record = 1;
        //    while (enumerator.MoveNext())
        //    {
        //        string dataString = record < 10 ? "000" + record.ToString() + " 12345" : "0000" + record.ToString() + " 12345";
        //        IField f = (IField)enumerator.Current;
        //        if (f.Name.Equals("data"))
        //        {
        //            f.Value = dataString;
        //            f.TextFont = "Arial";
        //            f.TextSize = 12;
        //            record++;
        //        }
        //    }
        //    //CAcroPDDoc doc = (AcroPDDoc)this.pdf.GetPDDoc();
        //    //object jsObj = doc.GetJSObject();
        //    //Type t = jsObj.GetType();

        //    //object[] param3 = new object[] { dataPath, row };
        //    //t.InvokeMember("importTextData", System.Reflection.BindingFlags.InvokeMethod | System.Reflection.BindingFlags.Public | System.Reflection.BindingFlags.Instance,
        //    //    null, jsObj, param3);
        //}
        #endregion
    }
}

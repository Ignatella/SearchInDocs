using Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using Rectangle = System.Drawing.Rectangle;
using Word = Microsoft.Office.Interop.Word;

namespace WordLibrary
{
    public static class Search
    {
        private static readonly string[] supportedFileExtensions = { ".doc", ".docx" };

        private static object syncObject = new object();
        private static CancellationTokenSource cancellationToken = new CancellationTokenSource();
        private static Word.Application[] wordApps { get; set; }
        private static Document[] wordDocs { get; set; }
        private static List<int>[] pagesNumbersLists { get; set; }
        private static DirectoryInfo dir { get; set; }
        private static List<FileInfo> Files { get; set; }
        private static string strToSearchFor { get; set; }

        public static event EventHandler<ErrorOccuredEventArgs> ErrorOccured; //need to be fixed (due to multithread)

        public static void SearchInFilesAndConvertPagesToJpg(string strToSearchFor, string path, Action action)
        {
            dir = new DirectoryInfo(path);
            Search.strToSearchFor = strToSearchFor.ToLower();


            Files = new List<FileInfo>();
            dir.GetFiles().Where(fileInfo => !fileInfo.Name.Contains("~") &&
                supportedFileExtensions.Contains(fileInfo.Extension))
                    .Foreach(fileInfo => Files.Add(fileInfo));


            wordApps = new Application[Files.Count];
            wordDocs = new Document[Files.Count];
            pagesNumbersLists = new List<int>[Files.Count];

            KillAllRequiredWordProcesses();

            try 
            { 

                Parallel.For(0, Files.Count, new ParallelOptions()
                {
                    MaxDegreeOfParallelism = 3,
                    CancellationToken = cancellationToken.Token
                }, (int q) =>
                {
                    try
                    {
                        wordApps[q] = new Word.Application();
                        wordDocs[q] = wordApps[q].Documents.OpenNoRepairDialog(FileName: Files[q].FullName, ReadOnly: false,
                        AddToRecentFiles: false);
                        pagesNumbersLists[q] = new List<int>();

                        SearchInTextBox(Search.strToSearchFor, q);
                        SearchInParagraphs(Search.strToSearchFor, q);
                        ConvertDocToJpeg(Search.strToSearchFor, Files[q].Name, q, pagesNumbersLists[q].ToArray());

                        wordApps[q].Visible = false;
                    }
                    catch (Exception ex)
                    {
                        ErrorOccured?.Invoke(null, new ErrorOccuredEventArgs(Files[q].Name, ex));
                    }
                    finally
                    {
                        wordDocs[q].Close(false);
                        wordApps[q].Quit(false);
                        pagesNumbersLists[q].Clear();

                        action.BeginInvoke(action.EndInvoke, null);
                    }
                });
            }
            catch(OperationCanceledException) { }
            finally
            {
                cancellationToken.Dispose();
            }

            KillAllRequiredWordProcesses();
        }

        public static int GetFileCount(string path) =>
             (new DirectoryInfo(path)).GetFiles().Where(fileInfo => !fileInfo.Name.Contains("~") &&
                supportedFileExtensions.Contains(fileInfo.Extension)).Count();

        public static void Cancel()
        {
            cancellationToken.Cancel();
        }

        private static void SearchInTextBox(string wordToSearchFor, int threadNum)
        {
            foreach (Shape shape in wordApps[threadNum].ActiveDocument.Shapes)
            {
                if (shape.TextFrame.HasText == -1 && shape.TextFrame.TextRange.Text.ToLower().Contains(wordToSearchFor))
                {
                    Range rang = shape.TextFrame.TextRange;
                    string textFrameText = rang.Text.ToLower();
                    int wordPosition = rang.Text.ToLower().IndexOf(wordToSearchFor);

                    while (wordPosition > -1)
                    {
                        rang.Select();
                        wordApps[threadNum].Selection.MoveLeft(WdUnits.wdCharacter, 1, WdMovementType.wdMove);
                        wordApps[threadNum].Selection.MoveRight(WdUnits.wdCharacter, wordPosition, WdMovementType.wdMove);
                        wordApps[threadNum].Selection.MoveRight(WdUnits.wdCharacter, wordToSearchFor.Length, WdMovementType.wdExtend);
                        wordApps[threadNum].Selection.Range.HighlightColorIndex = WdColorIndex.wdYellow;

                        wordPosition = rang.Text.ToLower().IndexOf(wordToSearchFor, wordPosition + wordToSearchFor.Length);
                    }

                    AddPageTopagesNumberList(threadNum);
                }
            }
        }

        private static void KillAllRequiredWordProcesses() //not the best solution
        {
            Process[] processes = Process.GetProcesses(".")
                .Where(process => process.ProcessName.ToLower().Contains("word")).ToArray();

            //kill if no window or if required file is opened by user.
            foreach (FileInfo file in Files)
            {
                foreach (Process process in processes
                    .Where(process => !process.HasExited && (process.MainWindowHandle == new IntPtr(0) ||
                        process.MainWindowTitle.Contains(file.Name))))
                {
                    try
                    {
                        process.Kill();
                        break;
                    }
                    catch { }
                }
            }
        }

        private static void SearchInParagraphs(string wordToSearchFor, int threadNum)
        {
            for (int i = 1; i <= wordDocs[threadNum].Paragraphs.Count; i++)
            {
                string parText = wordDocs[threadNum].Paragraphs[i].Range.Text;
                if (parText.ToLower().Contains(wordToSearchFor))
                {
                    int wordPosition = parText.ToLower().IndexOf(wordToSearchFor);
                    while (wordPosition > -1)
                    {
                        wordDocs[threadNum].Paragraphs[i].Range.Select();
                        wordApps[threadNum].Selection.MoveLeft(WdUnits.wdCharacter, 1, WdMovementType.wdMove);

                        wordApps[threadNum].Selection.MoveRight(WdUnits.wdCharacter, wordPosition, WdMovementType.wdMove);
                        wordApps[threadNum].Selection.MoveRight(WdUnits.wdCharacter, wordToSearchFor.Length, WdMovementType.wdExtend);
                        wordApps[threadNum].Selection.Range.HighlightColorIndex = WdColorIndex.wdYellow;

                        wordPosition = parText.ToLower().IndexOf(wordToSearchFor, wordPosition + wordToSearchFor.Length); ;
                    }

                    AddPageTopagesNumberList(threadNum);
                }
            }
        }

        private static void AddPageTopagesNumberList(int threadNum)
        {
            int currentPageNumber = wordApps[threadNum].Selection.Information[WdInformation.wdActiveEndPageNumber];
            if (!pagesNumbersLists[threadNum].Contains(currentPageNumber))
            {
                pagesNumbersLists[threadNum].Add(currentPageNumber);
            }
        }

        private static void ConvertDocToJpeg(string searchWord, string fileName, int threadNum, params int[] pagesNumbersToBeConverted)
        {
            DirectoryInfo wordDir = dir.CreateSubdirectory(searchWord);

            foreach (int pageNumber in pagesNumbersToBeConverted)
            {
                wordApps[threadNum].Selection.GoTo(WdGoToItem.wdGoToPage, WdGoToDirection.wdGoToAbsolute, pageNumber);

                var page = wordDocs[threadNum].ActiveWindow.ActivePane.Pages[pageNumber];
                var bits = page.EnhMetaFileBits;

                string target = wordDir.FullName + @"\" + searchWord + "_" + pageNumber + "_" + fileName;

                lock (syncObject)
                {
                    using (var ms = new MemoryStream((byte[])(bits)))
                    {
                        using (Bitmap bmp1 = new Bitmap(ms))
                        {
                            var pngTarget = Path.ChangeExtension(target, "jpeg");

                            using (Bitmap bmp2 = new Bitmap(bmp1.Width, bmp1.Height))
                            {
                                Transparent2Color(bmp1, bmp2, Color.White).Save(pngTarget, ImageFormat.Jpeg);
                            }
                        }
                    }
                }
            }
        }

        private static Bitmap Transparent2Color(Bitmap bmp1, Bitmap bmp2, Color target)
        {
            Rectangle rect = new Rectangle(System.Drawing.Point.Empty, bmp1.Size);
            using (Graphics G = Graphics.FromImage(bmp2))
            {
                G.Clear(target);
                G.DrawImageUnscaledAndClipped(bmp1, rect);
            }
            return bmp2;
        }
    }
}

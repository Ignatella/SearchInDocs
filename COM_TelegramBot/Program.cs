using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;
using Microsoft.Office.Interop.Word;
using Rectangle = System.Drawing.Rectangle;
using Word = Microsoft.Office.Interop.Word;

namespace COM_TelegramBot
{
    class Program
    {
        public static Word.Application[] wordApps;
        public static Document[] wordDocs;

        public static List<int>[] pagesNumbersLists;

        public static object syncObject = new object();

        public static DirectoryInfo dir = new DirectoryInfo(@"C:\Users\Mi\Desktop\bot\data");
        public static List<FileInfo> FileNames = new List<FileInfo>();

        public static string strToSearhFor = "grupa";
        public static void Main(string[] args)
        {
            dir.GetFiles().Where(fileInfo => !fileInfo.Name.Contains("~")).Foreach(fileInfo => FileNames.Add(fileInfo));


            wordApps = new Application[FileNames.Count];
            wordDocs = new Document[FileNames.Count];
            pagesNumbersLists = new List<int>[FileNames.Count];

            Parallel.For(0, FileNames.Count, new ParallelOptions() { MaxDegreeOfParallelism = 3 }, (int q) =>
            {
                try
                {
                    wordApps[q] = new Word.Application();
                    wordDocs[q] = wordApps[q].Documents.OpenNoRepairDialog(FileName: FileNames[q].FullName, ReadOnly: false);
                    pagesNumbersLists[q] = new List<int>();

                    SearchInTextBox(strToSearhFor, q);
                    SearchInParagraphs(strToSearhFor, q);
                    ConvertDocToJpeg(strToSearhFor, FileNames[q].Name, q, pagesNumbersLists[q].ToArray());

                    wordApps[q].Visible = false;
                }
                finally
                {
                    wordDocs[q].Close(false);
                    wordApps[q].Quit(false);
                    pagesNumbersLists[q].Clear();
                }
            });

            Console.WriteLine("finished");
            Console.ReadLine();
        }

        public static void ConvertDocToJpeg(string searchWord, string fileName, int threadNum, params int[] pagesNumbersToBeConverted)
        {
            DirectoryInfo wordDir = dir.Parent.CreateSubdirectory(searchWord);

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
                        Bitmap bitmap = new Bitmap(ms);

                        var pngTarget = Path.ChangeExtension(target, "jpeg");

                        Transparent2Color(bitmap, Color.White).Save(pngTarget, ImageFormat.Jpeg);
                    }
                }
            }
        }

        public static Bitmap Transparent2Color(Bitmap bmp1, Color target)
        {
            Bitmap bmp2 = new Bitmap(bmp1.Width, bmp1.Height);

            Rectangle rect = new Rectangle(System.Drawing.Point.Empty, bmp1.Size);
            using (Graphics G = Graphics.FromImage(bmp2))
            {
                G.Clear(target);
                G.DrawImageUnscaledAndClipped(bmp1, rect);
            }
            return bmp2;

        }

        public static void SearchInTextBox(string wordToSearchFor, int threadNum)
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

        public static void SearchInParagraphs(string wordToSearchFor, int threadNum)
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

        public static void AddPageTopagesNumberList(int threadNum)
        {
            int currentPageNumber = wordApps[threadNum].Selection.Information[WdInformation.wdActiveEndPageNumber];
            if (!pagesNumbersLists[threadNum].Contains(currentPageNumber))
            {
                pagesNumbersLists[threadNum].Add(currentPageNumber);
            }
        }
    }
}

#region FindTheWordInText
//Word.Range range = wordApp.ActiveDocument.Content;
//Word.Find find = range.Find;
//find.Text = "xxx";
//                find.ClearFormatting();
//                find.ClearAllFuzzyOptions();
//                find.MatchControl = true;
//                Console.WriteLine(find.Execute()); 
#endregion

#region SaveImage


//var image = System.Drawing.Image.FromStream(ms);
//var pngTarget = Path.ChangeExtension(target, "png");
//image.Save(pngTarget, ImageFormat.Png);

#endregion
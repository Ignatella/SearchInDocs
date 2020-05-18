using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Runtime.CompilerServices;
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
        public static Word.Application wordApp;
        public static Document wordDoc;

        public static List<int> pagesNumbersList = new List<int>();

        public static DirectoryInfo dir = new DirectoryInfo(@"C:\Users\Mi\Desktop\bot\data");
        public  static List<FileInfo> FileNames = new List<FileInfo>();

        public static string strToSearhFor = "reklamacj";
        public static void Main(string[] args)
        { 
            dir.GetFiles().Foreach(fileInfo => FileNames.Add(fileInfo));

            foreach (FileInfo fileName in FileNames)
            {
                try
                {
                    wordApp = new Word.Application();
                    wordDoc = wordApp.Documents.OpenNoRepairDialog(FileName: fileName.FullName, ReadOnly: false);

                    SearchInTextBox(strToSearhFor);
                    SearchInParagraphs(strToSearhFor);
                    ConvertDocToJpeg(strToSearhFor, fileName.Name, pagesNumbersList.ToArray());

                    wordApp.Visible = false;
                }
                finally
                {
                    wordDoc.Close(false);
                    wordApp.Quit(false);
                    pagesNumbersList.Clear();
                }
            }
            Console.WriteLine("finished");
            Console.ReadLine();
        }

        public static void ConvertDocToJpeg(string searchWord, string fileName, params int[] pagesNumbersToBeConverted)
        {
            DirectoryInfo wordDir = dir.Parent.CreateSubdirectory(searchWord); //tmp

            foreach (int pageNumber in pagesNumbersToBeConverted)
            {
                wordApp.Selection.GoTo(WdGoToItem.wdGoToPage, WdGoToDirection.wdGoToAbsolute, pageNumber);

                var page = wordDoc.ActiveWindow.ActivePane.Pages[pageNumber];
                var bits = page.EnhMetaFileBits;

                string target = wordDir.FullName + @"\" + searchWord + "_" + pageNumber +"_" + fileName;
               
                using (var ms = new MemoryStream((byte[])(bits)))
                {
                    Bitmap bitmap = new Bitmap(ms);
                    
                    var pngTarget = Path.ChangeExtension(target, "jpeg");

                    Transparent2Color(bitmap, Color.White).Save(pngTarget, ImageFormat.Jpeg);
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

        public static void SearchInTextBox(string wordToSearchFor)
        {
            foreach (Shape shape in wordApp.ActiveDocument.Shapes)
            {
                if (shape.TextFrame.HasText == -1 && shape.TextFrame.TextRange.Text.ToLower().Contains(wordToSearchFor))
                {
                    Range rang = shape.TextFrame.TextRange;
                    string textFrameText = rang.Text.ToLower();
                    int wordPosition = rang.Text.ToLower().IndexOf(wordToSearchFor);

                    while (wordPosition > -1)
                    {
                        rang.Select();
                        wordApp.Selection.MoveLeft(WdUnits.wdCharacter, 1, WdMovementType.wdMove);
                        wordApp.Selection.MoveRight(WdUnits.wdCharacter, wordPosition, WdMovementType.wdMove);
                        wordApp.Selection.MoveRight(WdUnits.wdCharacter, wordToSearchFor.Length, WdMovementType.wdExtend);
                        wordApp.Selection.Range.HighlightColorIndex = WdColorIndex.wdYellow;

                        wordPosition = rang.Text.ToLower().IndexOf(wordToSearchFor, wordPosition + wordToSearchFor.Length);
                    }
                    AddPageTopagesNumberList();
                }
            }
        }

        public static void SearchInParagraphs(string wordToSearchFor)
        {
            for (int i = 1; i <= wordDoc.Paragraphs.Count; i++)
            {
                string parText = wordDoc.Paragraphs[i].Range.Text;
                if (parText.ToLower().Contains(wordToSearchFor))
                {
                    int wordPosition = parText.ToLower().IndexOf(wordToSearchFor);
                    while (wordPosition > -1)
                    {
                        wordDoc.Paragraphs[i].Range.Select();
                        wordApp.Selection.MoveLeft(WdUnits.wdCharacter, 1, WdMovementType.wdMove);

                        wordApp.Selection.MoveRight(WdUnits.wdCharacter, wordPosition, WdMovementType.wdMove);
                        wordApp.Selection.MoveRight(WdUnits.wdCharacter, wordToSearchFor.Length, WdMovementType.wdExtend);
                        wordApp.Selection.Range.HighlightColorIndex = WdColorIndex.wdYellow;

                        wordPosition = parText.ToLower().IndexOf(wordToSearchFor, wordPosition + wordToSearchFor.Length); ;
                    }
                    AddPageTopagesNumberList();
                }
            }
        }

        public static void AddPageTopagesNumberList()
        {
            int currentPageNumber = wordApp.Selection.Information[WdInformation.wdActiveEndPageNumber];
            if (!pagesNumbersList.Contains(currentPageNumber))
            {
                pagesNumbersList.Add(currentPageNumber);
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
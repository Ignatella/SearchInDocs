using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text;
using System.Threading.Tasks;
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
        public static void Main(string[] args)
        {
            try
            {
                wordApp = new Word.Application();
                wordDoc = wordApp.Documents.OpenNoRepairDialog(FileName: @"C:\Users\Mi\Desktop\a.doc", ReadOnly: false);

                string strToSearhFor = "protest";



                //SearchInTextBox(strToSearhFor);
                //SearchInParagraphs(strToSearhFor);
                //ConvertDocToJpeg(pagesNumbersList.ToArray());

                ConvertDocToJpeg(1);

                Console.WriteLine("finished");
                wordApp.Visible = false;

                Console.ReadLine();
            }
            finally
            {
                wordDoc.Close(false);
                wordApp.Quit(false);
            }
        }

        public static void ConvertDocToJpeg(params int[] pagesNumbersToBeConverted)
        {
            foreach (int pageNumber in pagesNumbersToBeConverted)
            {
                wordApp.Selection.GoTo(WdGoToItem.wdGoToPage, WdGoToDirection.wdGoToAbsolute, pageNumber);

                var page = wordDoc.ActiveWindow.ActivePane.Pages[pageNumber];
                var bits = page.EnhMetaFileBits;
                var target = @"C:\Users\Mi\Desktop\a" + "_image" + pageNumber + ".doc";

                using (var ms = new MemoryStream((byte[])(bits)))
                {
                    using (Bitmap bitmap = new Bitmap(ms))
                    {
                        var pngTarget = Path.ChangeExtension(target, "jpeg");


                        using (Bitmap targetBmp = bitmap.Clone(new Rectangle(0, 0, bitmap.Width, bitmap.Height), PixelFormat.Format32bppArgb))
                        {
                            targetBmp.Save(pngTarget, ImageFormat.Jpeg);
                        }
                    }

                   
                    







                    //var image = System.Drawing.Image.FromStream(ms);
                    //var pngTarget = Path.ChangeExtension(target, "png");
                    //image.Save(pngTarget, ImageFormat.Png);
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

            //Transparent2Color(bitmap, Color.White).Save(pngTarget, ImageFormat.Jpeg);
        }

        public static void SearchInTextBox(string wordToSearchFor)
        {
            foreach (Shape shape in wordApp.ActiveDocument.Shapes)
            {
                if (shape.TextFrame.HasText == -1 && shape.TextFrame.TextRange.Text.ToLower().Contains(wordToSearchFor))
                {
                    Range rang = shape.TextFrame.TextRange;
                    int wordPosition = shape.TextFrame.TextRange.Text.IndexOf(wordToSearchFor);
                    rang.Select();
                    wordApp.Selection.HomeKey(WdUnits.wdLine, WdMovementType.wdMove);
                    wordApp.Selection.MoveRight(WdUnits.wdCharacter, wordPosition, WdMovementType.wdMove);
                    wordApp.Selection.MoveRight(WdUnits.wdWord, 1, WdMovementType.wdExtend);
                    wordApp.Selection.Range.HighlightColorIndex = WdColorIndex.wdYellow;

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
                    wordApp.Selection.HomeKey(WdUnits.wdStory, WdMovementType.wdMove);
                    wordApp.Selection.MoveDown(WdUnits.wdParagraph, i - 1, WdMovementType.wdMove);
                    wordApp.Selection.MoveRight(WdUnits.wdCharacter, wordPosition, WdMovementType.wdMove);
                    wordApp.Selection.MoveRight(WdUnits.wdWord, 1, WdMovementType.wdExtend);
                    wordApp.Selection.Range.HighlightColorIndex = WdColorIndex.wdYellow;

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
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Word = Microsoft.Office.Interop.Word;

namespace COM_TelegramBot
{
    class Program
    {
        public static Word.Application wordApp;
        public static Document wordDoc;
        public static void Main(string[] args)
        {
            try
            {
                wordApp = new Word.Application();

                wordDoc = wordApp.Documents.OpenNoRepairDialog(FileName: @"C:\Users\Mi\Desktop\a.doc", ReadOnly: false);

                string strToSearhFor = "protest";

                SearchInTextBox(strToSearhFor);


                wordApp.Visible = true;
                Console.ReadLine();
            }
            finally
            {
                wordDoc.Close(false);
                wordApp.Quit(false);
            }
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
                }
            }
        }
    }
}

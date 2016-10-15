using System;
using Microsoft.Office.Interop.Word;

namespace Alegor.DocxConcat
{
    public class Program
    {
        public Properties Properties { get; }

        public Program(Properties properties)
        {
            Properties = properties;
        }

        public void Run()
        {
            Application application = null;
            try
            {
                application = new Application();
                var document = application.Documents
                    .Open(Properties.InputDocumentPathList[Properties.BaseInputDocumentIndex]);

                var selection = application.Selection;
                selection.GoTo(WdGoToItem.wdGoToSection, WdGoToDirection.wdGoToFirst);

                for (var i = Properties.BaseInputDocumentIndex - 1; i >= 0; i--)
                {
                    var inputDocumentPath = Properties.InputDocumentPathList[i];

                    selection.WholeStory();
                    selection.MoveLeft(WdUnits.wdCharacter, 1);
                    selection.InsertParagraph();
                    selection.MoveLeft(WdUnits.wdCharacter, 1);
                    selection.InsertFile(inputDocumentPath);
                }

                for (var i = 0; i < Properties.InputDocumentPathList.Count; i++)
                {
                    if (i == Properties.BaseInputDocumentIndex)
                    {
                        continue;
                    }

                    var inputDocumentPath = Properties.InputDocumentPathList[i];

                    if (i < Properties.BaseInputDocumentIndex)
                    {
                        //  TODO: Написать реализацию
                    }
                    else
                    {
                        selection.WholeStory();
                        selection.MoveRight(WdUnits.wdCharacter, 1);
                        selection.InsertParagraph();
                        selection.MoveRight(WdUnits.wdCharacter, 1);
                        selection.InsertFile(inputDocumentPath);
                    }

                    document.SaveAs(Properties.OutputDocumentPath);
                }
            }
            catch (Exception e)
            {
                Console.WriteLine(e);
                Console.WriteLine(e.StackTrace);
                Environment.ExitCode = 1;
            }
            finally
            {
                application?.Quit();
            }
        }
    }
}
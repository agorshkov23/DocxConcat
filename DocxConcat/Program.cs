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
                        selection.GoTo(WdGoToItem.wdGoToSection, WdGoToDirection.wdGoToLast);
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
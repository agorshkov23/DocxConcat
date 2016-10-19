using System.Collections.Generic;

namespace Alegor.DocxConcat
{
    public class Properties
    {
        public List<string> InputDocumentPathList { get; } = new List<string>();
        public int BaseInputDocumentIndex { get; set; } = 0;
        public string OutputDocumentPath { get; set; } = null;

        public bool Validate()
        {
            var errorList = new List<string>();

            return errorList.Count == 0;
        }
    }
}
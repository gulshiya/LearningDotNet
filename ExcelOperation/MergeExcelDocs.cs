using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Activities;
using System.ComponentModel;

namespace ExcelOperation
{
    public class MergeExcelDocs : CodeActivity
    {
        [Category("Input")]
        [RequiredArgument]
        public InArgument<string> InputFolderPath { get; set; }

        [Category("Input")]
        [RequiredArgument]
        public InArgument<string> OutputFolderPath { get; set; }

        [Category("Input")]
        [RequiredArgument]
        public InArgument<string> OutPutFileName { get; set; }

        protected override void Execute(CodeActivityContext context)
        {
            string inputFolderPath = InputFolderPath.Get(context);
            string outputFolderPath = OutputFolderPath.Get(context);
            string outPutFileName = OutPutFileName.Get(context);
            HelperMethod method = new HelperMethod();
            method.mergeExcels(inputFolderPath, outputFolderPath, outPutFileName);
        }
    }
}

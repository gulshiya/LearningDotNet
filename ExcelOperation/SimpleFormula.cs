using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Activities;
using System.ComponentModel;

namespace ExcelOperation
{
    public class SimpleFormula : CodeActivity
    {
        [Category("Input")]
        [RequiredArgument]
        public InArgument<double> FirstNumber { get; set; }

        [Category("Input")]
        public InArgument<double> SecondNumber { get; set; }

        [Category("Output")]
        public OutArgument<double> ResultNumber { get; set; }

        protected override void Execute(CodeActivityContext context)
        {
            double firstNumber = FirstNumber.Get(context);
            double secondNumber = SecondNumber.Get(context);

            HelperMethod method = new HelperMethod();
            var result = method.sumTwoNumber(firstNumber, secondNumber);
            ResultNumber.Set(context, result);

        }
    }
}

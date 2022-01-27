using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Activities;
using System.ComponentModel;
using _Excel = Microsoft.Office.Interop.Excel;
using _PPT =Microsoft.Office.Interop.PowerPoint;

namespace image_xlsxToppt
{
    public class XLSXToPPT : CodeActivity
    {
        [Category("Input")]
        [RequiredArgument]
        public InArgument<string> ExcelPath { get; set; }

        [Category("Output")]
        [RequiredArgument]
        public OutArgument<double> ResultNumber { get; set; }

        protected override void Execute(CodeActivityContext context)
        {
            var excelPath = ExcelPath.Get(context);

            _Excel.Application excel = new _Excel.Application();

            if(excel == null)
            {
                //Error Excel couldn't be started!!
                //TO-DO
            }

            excel.Visible = true;

            _Excel.Workbook wb = excel.Workbooks.Open(excelPath);


            ResultNumber.Set(context, null); ;
        }
    }
}

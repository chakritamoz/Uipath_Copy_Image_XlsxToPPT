using System;
using System.Activities;
using System.ComponentModel;
using _Excel = Microsoft.Office.Interop.Excel;
using _PPT =Microsoft.Office.Interop.PowerPoint;
using System.Windows.Forms;
using System.IO;
using System.Collections.Generic;

namespace CopyGroupExcelToPPT
{
    [Category("Arutu")]
    [DisplayName("Copy Group Excel to PPT")]
    [Description("Copy group from excel to powerpoint")]
    public class Main : CodeActivity
    {
        [Category("Input")]
        [RequiredArgument]
        [Description("Enter excel path")]
        public InArgument<string> ExcelPath { get; set; }

        [Category("Input")]
        [Description("Enter powerpoint path")]
        [RequiredArgument]
        public InArgument<string> PptPath { get; set; }

        [Category("Input")]
        [RequiredArgument]
        [Description("Enter start page layout powerpoint")]
        public InArgument<int> LayoutPage { get; set; }

        [Category("Input")]
        [Description("Enter shape name")]
        [RequiredArgument]
        public InArgument<string> ShapeName { get; set; }

        [Category("Input")]
        [Description("Enter list sheet want to delete")]
        public InArgument<List<string>> SheetBanList { get; set; }

        protected override void Execute(CodeActivityContext context)
        {
            var excelPath = ExcelPath.Get(context);
            var pptPath = PptPath.Get(context);
            var layoutPage = LayoutPage.Get(context);
            var shapeName = ShapeName.Get(context);
            var sheetBanList = SheetBanList.Get(context);

            _Excel.Application excel = new _Excel.Application();
            _PPT.Application ppt = new _PPT.Application();

            _Excel.Workbook wb = readWorkBook(excel, excelPath);
            excel.DisplayAlerts = false;
            if(sheetBanList != null)
            {
                wb = deleteWorkSheet(wb, sheetBanList);
            }
            _PPT.Presentation pptPresentation = readPowerpoint(ppt, pptPath);
            _PPT.Slides slides = pptPresentation.Slides;
            if (File.Exists(pptPath))
            {
                addAlreadyPPT(wb, shapeName, slides, layoutPage);
            }
            else
            {
                layoutPage = 1;
                _PPT.CustomLayout customLayout = pptPresentation.SlideMaster.CustomLayouts[_PPT.PpSlideLayout.ppLayoutTitle];
                addNonPPT(wb, shapeName, slides, customLayout, layoutPage);
            }
            closeApplication(excel, wb, ppt, pptPresentation, pptPath);
        }

        public static _Excel.Workbook readWorkBook(_Excel.Application excel, string excelPath)
        {
            if (excel == null)
            {
                //Error Excel couldn't be started!!
                //TO-DO
            }
            excel.Visible = true;
            _Excel.Workbook wb = excel.Workbooks.Open(excelPath, 0, true, 5, "", "", true, _Excel.XlPlatform.xlWindows, "\t", false, false, 0, true);

            return wb;
        }

        public static _PPT.Presentation readPowerpoint(_PPT.Application ppt, string pptPath)
        {
            _PPT.Presentation pptPresentation = null;
            if (File.Exists(pptPath))
            {
                 pptPresentation = ppt.Presentations.Open(pptPath);

            }
            else
            {
                pptPresentation = ppt.Presentations.Add(Microsoft.Office.Core.MsoTriState.msoTrue);
            }

            return pptPresentation;
        }

        public static void addAlreadyPPT(_Excel.Workbook wb, string shapeName, _PPT.Slides slides, int layoutPage)
        {
            foreach (_Excel.Worksheet sheet in wb.Worksheets)
            {
                bool flag = true;
                _PPT.Slide slide = slides._Index(layoutPage);
                sheet.Shapes.Item(shapeName).Copy();
                while (Clipboard.GetDataObject() == null)
                {
                    sheet.Shapes.Item(shapeName).Copy();
                }
                while (flag)
                {
                    try
                    {
                        _PPT.ShapeRange shapeRange = slide.Shapes.Paste();
                        flag = false;
                    }
                    catch
                    {
                        flag = true;
                    }
                }
                layoutPage++;
            }
        }

        public static void addNonPPT(_Excel.Workbook wb, string shapeName, _PPT.Slides slides, _PPT.CustomLayout customLayout, int layoutPage = 1)
        {
            //_Excel.ChartObjects chartObjs = (_Excel.ChartObjects)(sheet.ChartObjects());
            /*foreach (_Excel.ChartObject chartObj in chartObjs)
            {
                chartObj.CopyPicture();
                _PPT.ShapeRange shapeRange = slide.Shapes.Paste();
            }*/
            //slide.Shapes[1].Width = 500;


            foreach (_Excel.Worksheet sheet in wb.Worksheets)
            {
                bool flag = true;
                sheet.Shapes.Item(shapeName).Copy();
                while (Clipboard.GetDataObject() == null)
                {
                    sheet.Shapes.Item(shapeName).Copy();
                }
                _PPT._Slide slide = slides.AddSlide(layoutPage, customLayout);
                while (slide.Shapes.Count > 0)
                {
                    slide.Shapes[1].Delete();
                }
                while (flag)
                {
                    try
                    {
                        _PPT.ShapeRange shapeRange = slide.Shapes.Paste();
                        flag = false;
                    }
                    catch
                    {
                        flag = true;
                    }
                }
                layoutPage++;
            }
        }

        public static _Excel.Workbook deleteWorkSheet(_Excel.Workbook wb, List<string> sheetBanLists)
        {
            foreach(_Excel.Worksheet sheet in wb.Worksheets)
            {
                if (sheetBanLists.Contains(sheet.Name.ToString()))
                {
                    sheet.Delete();
                }
            }

            return wb;
        }

        public static void closeApplication(_Excel.Application excel, _Excel.Workbook wb, _PPT.Application ppt, _PPT.Presentation pptPresentation, string pptPath)
        {
            wb.Close();
            excel.Quit();
            if (File.Exists(pptPath))
            {
                pptPresentation.Save();
            }else
            {
                pptPresentation.SaveAs(pptPath);
            }
            pptPresentation.Close();
            ppt.Quit();
        }
    }
}

using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.ComponentModel.DataAnnotations;
using System.Data;
using System.IO;
using System.Linq;
using System.Reflection;
using CadDev.Excel.OpenXML;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using ekkodale.Cutter.Model;
using MessageBox = System.Windows.Forms.Message;


namespace CadDev.Excel.OpenXML
{
    public class SlugGeneratorByGKey<T> : ExcelGenerator<T>
    {
        private dynamic _model;

        public SlugGeneratorByGKey(string filePath, dynamic model) : base(filePath)
        {
            _model = model;
        }

        /// <summary>
        /// Creates an excel file from a list.
        /// </summary>
        /// <param name="data">List with desired object type.</param>
        /// <param name="OutPutFileDirectory">Output File Directory where the excel file should be placed.</param>
        public override string CreateExcelFile(List<T> data, string selectedPrefabJob)
        {
            var datetime = DateTime.Now.ToString().Replace("/", "_").Replace(":", "_");
            try
            {

                string fileFullname = Path.Combine(FilePath, "Cutter_Export_" + datetime + "_" + selectedPrefabJob + ".xlsx");
                FilePathGenerated = fileFullname;
                document = SpreadsheetDocument.Create(fileFullname, SpreadsheetDocumentType.Workbook);
                using (document)
                {

                    WorkbookPart workbookPart1 = document.AddWorkbookPart();

                    int counter = 0;

                    foreach (IGrouping<string, T> groupedSlugs in _model)
                    //foreach (var inheritedSlugs in GroupedList)
                    {
                        if (counter <= 2)
                        {
                            counter++;
                            CreatePartsForExcelTest(document, groupedSlugs.ToList(), groupedSlugs.Key, workbookPart1, counter);
                        }
                    }
                    //CreatePartsForExcel(document, data);
                    SaveAsCsv(document, Path.ChangeExtension(fileFullname, "csv"));
                }
                return fileFullname;
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
                return "";
            }
        }


        /// <summary>
        /// Creates the SheetDatas and Workbooks.
        /// </summary>
        /// <param name="document">List with desired object type.</param>
        /// <param name="data"></param>
        protected void CreatePartsForExcelTest(SpreadsheetDocument document, List<T> data, string sheetName, WorkbookPart workbookPart1, int counter)
        {


            SheetData partSheetData = GenerateSheetdataForDetails(data);



            GenerateWorkbookPartContent(workbookPart1, sheetName);

            // WorkbookStylesPart workbookStylesPart1 = workbookPart1.AddNewPart<WorkbookStylesPart>($"rId{counter}");
            // GenerateWorkbookStylesPartContent(workbookStylesPart1);

            WorksheetPart worksheetPart1 = workbookPart1.AddNewPart<WorksheetPart>($"xId{counter}");
            GenerateWorksheetPartContent(worksheetPart1, partSheetData);


        }

        /// <summary>
        /// Creates the workbook.
        /// </summary>
        /// <param name="workbookPart1"></param>
        protected override void GenerateWorkbookPartContent(WorkbookPart workbookPart1, string sheetName)
        {
            try
            {
                Workbook workbook1 = new Workbook();
                Sheets sheets1 = new Sheets();
                Sheet sheet1 = new Sheet() { Name = sheetName, SheetId = (UInt32Value)1U, Id = "rId1" };
                sheets1.Append(sheet1);
                workbook1.Append(sheets1);
                workbookPart1.Workbook = workbook1;
            }
            catch (Exception ec)
            {

                throw;
            }

        }

        /// <summary>
        /// Creates the headers in the excel content.
        /// </summary>
        /// <param name="type">Desired object type.</param>
        /// <returns>Row with all properties as headers of its given type.</returns>
        /// 
        protected override Row CreateHeaderRowForExcel(Type type)
        {
            Row workRow = new Row();
            try
            {
                foreach (PropertyInfo propertyInfo in type.GetProperties())
                {
                    if (propertyInfo.Name == "NominalDiameter" || propertyInfo.Name == "Material" || propertyInfo.Name == "Typ")
                    {
                        workRow.Append(CreateCell(GetNameOfProperty(propertyInfo), 2U));
                    }
                }
                foreach (PropertyInfo propertyInfo in type.GetProperties())
                {
                    if (propertyInfo.Name == "System" || propertyInfo.Name == "Length" || propertyInfo.Name == "CutLoss" || propertyInfo.Name == "FreeLength")
                    {
                        workRow.Append(CreateCell(GetNameOfProperty(propertyInfo), 2U));
                    }
                }
                workRow.Append(CreateCell("Cuts", 2U));
                return workRow;
            }
            catch (Exception ex)
            {
                return null;
            }
        }
        /// <summary>
        /// Generates the content of an property out of its values.
        /// </summary>
        /// <param name="model">Generic object model which will be used.</param>
        /// <returns>Row with content of all properties of its given generic object.</returns>
        protected override Row GenerateRowForChildPartDetail(T model)
        {
            Row tRow = new Row();

            foreach (PropertyInfo propertyInfo in model.GetType().GetProperties())
            {
                if (propertyInfo.Name == "NominalDiameter" || propertyInfo.Name == "Material" || propertyInfo.Name == "Typ")
                {
                    var value = propertyInfo.GetValue(model);
                    tRow.Append(CreateCell(value));
                }

            }

            foreach (PropertyInfo propertyInfo in model.GetType().GetProperties())
            {
                if (propertyInfo.Name == "System" || propertyInfo.Name == "Length" || propertyInfo.Name == "CutLoss" || propertyInfo.Name == "FreeLength")
                {
                    var value = propertyInfo.GetValue(model);
                    tRow.Append(CreateCell(value));
                }
            }

            dynamic slug = model;
            foreach (var cut in slug.Cuts)
            {
                var line = $"Postion Number: {cut.Key.PositionNumber} &";
                line += $"& Cut:{cut.Value}";
                tRow.Append(CreateCell(line));
            }
            return tRow;
        }


        protected string GetNameOfProperty(PropertyInfo pinfo)
        {
            string result = "N/A";

            if (!String.IsNullOrEmpty(GetDisplayNameOfProperty(pinfo)))
                result = GetDisplayNameOfProperty(pinfo);
            else
                result = pinfo.Name;

            return result;
        }

        protected string GetDisplayNameOfProperty(PropertyInfo pinfo)
        {
            string result = "N/A";

            var attribute = pinfo.GetCustomAttributes(typeof(DisplayNameAttribute), true).Cast<DisplayNameAttribute>().Single();
            result = attribute.DisplayName;

            return result;
        }


        /// <summary>
        /// Generates the excel content (SheetData).
        /// </summary>
        /// <param name="data">List with desired object type.</param>
        /// <returns>Excel Content as SheetData.</returns>
        protected override SheetData GenerateSheetdataForDetails(List<T> data)
        {
            SheetData sheetData1 = new SheetData();
            sheetData1.Append(CreateHeaderRowForExcel(data.FirstOrDefault().GetType()));


            foreach (T model in data)
            {
                Row partsRows = GenerateRowForChildPartDetail(model);
                sheetData1.Append(partsRows);
            }
            return sheetData1;
        }


        protected override EnumValue<CellValues> ResolveCellDataTypeOnValue(Type type)
        {
            if (type == typeof(int) || type == typeof(double) || type == typeof(float))
            {
                return CellValues.Number;
            }
            ////
            else if (type == typeof(Dictionary<string, int>))
            {

                return CellValues.Error;
            }

            else if (type == typeof(bool))
            {
                return CellValues.Boolean;

            }

            else if (type == typeof(DateTime))
            {
                return CellValues.Date;
            }
            else
            {
                return CellValues.String;
            }
        }

    }
}

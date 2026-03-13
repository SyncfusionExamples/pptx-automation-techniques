using Syncfusion.Presentation;
using Syncfusion.XlsIO;
using Syncfusion.Drawing;
using System.Data;

namespace Create_Table_with_Dynamic_Data
{
    internal class ExcelToPresentationHelper
    {
        private const string SLIDE_TITLE    = "Financial Year Budget & Expense Report — FY 2024–2025";
        private const string EXCEL_FILE_NAME = "Sample.xlsx";

        // Table layout constants
        private const double TABLE_LEFT   = 150;
        private const double TABLE_TOP    = 140;
        private const double TABLE_WIDTH  = 660;
        private const double ROW_HEIGHT   = 28;

        /// <summary>
        /// Reads financial data from the bundled Excel file and renders it as a
        /// formatted table on the first slide of <paramref name="presentation"/>.
        /// </summary>
        /// <param name="presentation">The loaded <see cref="IPresentation"/> to populate.</param>
        public void AddFinancialDataToPresentation(IPresentation presentation)
        {
            //Resolve the Excel file path relative to the application base directory
            string dataFolderPath = Path.Combine(AppContext.BaseDirectory, "Data");
            string excelFilePath  = Path.Combine(dataFolderPath, EXCEL_FILE_NAME);

            //Load Excel data
            DataTable financialData = LoadExcelData(excelFilePath);

            //Get the first slide
            ISlide slide = presentation.Slides[0];

            //Add title to the slide
            AddSlideTitle(slide, SLIDE_TITLE);

            //Add table with financial data
            AddFinancialTable(slide, financialData);
        }

        private DataTable LoadExcelData(string excelFilePath)
        {
            using (ExcelEngine excelEngine = new ExcelEngine())
            {
                IApplication excelApplication = excelEngine.Excel;
                excelApplication.DefaultVersion = ExcelVersion.Xlsx;

                //Open the Excel workbook
                using (FileStream excelStream = new FileStream(excelFilePath, FileMode.Open, FileAccess.Read))
                {
                    IWorkbook workbook = excelApplication.Workbooks.Open(excelStream);
                    IWorksheet worksheet = workbook.Worksheets[0];

                    //Export data to DataTable
                    DataTable dataTable = worksheet.ExportDataTable(
                        worksheet.UsedRange,
                        ExcelExportDataTableOptions.ColumnNames | ExcelExportDataTableOptions.PreserveOleDate
                    );

                    //Convert DateTime values back to display text from Excel
                    for (int rowIndex = 0; rowIndex < dataTable.Rows.Count; rowIndex++)
                    {
                        for (int columnIndex = 0; columnIndex < dataTable.Columns.Count; columnIndex++)
                        {
                            //Get the display text from Excel cell
                            string cellDisplayText = worksheet.Range[rowIndex + 2, columnIndex + 1].DisplayText;
                            dataTable.Rows[rowIndex][columnIndex] = cellDisplayText;
                        }
                    }
                    return dataTable;
                }
            }
        }

        private void AddSlideTitle(ISlide slide, string titleText)
        {
            //Remove default placeholders
            for (int shapeIndex = slide.Shapes.Count - 1; shapeIndex >= 0; shapeIndex--)
            {
                var slideItem = slide.Shapes[shapeIndex];
                if (slideItem is Syncfusion.Presentation.IShape shape && shape.PlaceholderFormat != null)
                {
                    slide.Shapes.Remove(slideItem);
                }
            }

            //Add custom title shape at the top
            Syncfusion.Presentation.IShape titleShape = slide.Shapes.AddTextBox(150, 25, 680, 40);
            
            //Set title text
            ITextBody titleTextBody = titleShape.TextBody;
            IParagraph titleParagraph = titleTextBody.AddParagraph(titleText);
            titleParagraph.HorizontalAlignment = HorizontalAlignmentType.Center;
            
            //Format title text
            ITextPart titleTextPart = titleParagraph.TextParts[0];
            titleTextPart.Font.FontSize = 36;
            titleTextPart.Font.Bold = true;
            titleTextPart.Font.Color = ColorObject.FromArgb(21, 66, 115);

            //Add decorative rule line beneath the title — no border, solid terracotta fill
            Syncfusion.Presentation.IShape titleRule = slide.Shapes.AddShape(Syncfusion.Presentation.AutoShapeType.Rectangle, 48, 120, 864, 6);
            titleRule.Fill.SolidFill.Color = ColorObject.FromArgb(215, 100, 67);
            titleRule.LineFormat.Fill.FillType = FillType.None;
        }

        private void AddFinancialTable(ISlide slide, DataTable financialData)
        {
            int rowCount    = financialData.Rows.Count + 1; // +1 for header row
            int columnCount = financialData.Columns.Count;
            double tableHeight = rowCount * ROW_HEIGHT;

            //Add table to slide using named layout constants
            ITable table = slide.Shapes.AddTable(rowCount, columnCount, TABLE_LEFT, TABLE_TOP, TABLE_WIDTH, tableHeight);

            //Populate header row
            PopulateHeaderRow(table, financialData);

            //Populate data rows
            PopulateDataRows(table, financialData);

            //Apply table styling
            ApplyTableStyling(table);
        }

        private void PopulateHeaderRow(ITable table, DataTable financialData)
        {
            for (int columnIndex = 0; columnIndex < financialData.Columns.Count; columnIndex++)
            {
                ICell headerCell = table[0, columnIndex];
                headerCell.TextBody.AddParagraph(financialData.Columns[columnIndex].ColumnName);

                //Header cell formatting
                ITextPart headerTextPart = headerCell.TextBody.Paragraphs[0].TextParts[0];
                headerTextPart.Font.FontName = "Calibri";
                headerTextPart.Font.FontSize = 11;
                headerTextPart.Font.Bold = true;
                headerTextPart.Font.Color = ColorObject.White;

                //Header cell background
                headerCell.Fill.SolidFill.Color = ColorObject.FromArgb(192, 80, 77);
            }
        }

        private void PopulateDataRows(ITable table, DataTable financialData)
        {
            for (int rowIndex = 0; rowIndex < financialData.Rows.Count; rowIndex++)
            {
                for (int columnIndex = 0; columnIndex < financialData.Columns.Count; columnIndex++)
                {
                    ICell dataCell = table[rowIndex + 1, columnIndex];
                    object? cellValue = financialData.Rows[rowIndex][columnIndex];
                    string cellText = cellValue?.ToString() ?? string.Empty;
                    
                    dataCell.TextBody.AddParagraph(cellText);

                    //Data cell formatting
                    ITextPart dataTextPart = dataCell.TextBody.Paragraphs[0].TextParts[0];
                    dataTextPart.Font.FontName = "Calibri";
                    dataTextPart.Font.FontSize = 10;
                    dataTextPart.Font.Color = ColorObject.Black;

                    //Alternate row colors
                    if (rowIndex % 2 == 0)
                    {
                        dataCell.Fill.SolidFill.Color = ColorObject.FromArgb(242, 220, 219);
                    }
                    else
                    {
                        dataCell.Fill.SolidFill.Color = ColorObject.White;
                    }
                }
            }
        }

        private void ApplyTableStyling(ITable table)
        {
            //Distribute columns evenly across the total table width
            double columnWidth = TABLE_WIDTH / table.ColumnsCount;
            for (int columnIndex = 0; columnIndex < table.ColumnsCount; columnIndex++)
            {
                table.Columns[columnIndex].Width = columnWidth;
            }

            //Apply uniform row height
            for (int rowIndex = 0; rowIndex < table.Rows.Count; rowIndex++)
            {
                table.Rows[rowIndex].Height = ROW_HEIGHT;
            }

            //Apply cell alignment
            foreach (IRow row in table.Rows)
            {
                foreach (ICell cell in row.Cells)
                {
                    cell.TextBody.Paragraphs[0].HorizontalAlignment = HorizontalAlignmentType.Center;
                    cell.TextBody.VerticalAlignment = VerticalAlignmentType.Middle;
                }
            }
        }
    }
}

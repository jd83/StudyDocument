using NPOI.SS.UserModel;
using NPOI.SS.Util;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.IO;
namespace PrepareForWork
{
    public class Export
    {
        static void Main()
        {
            var requestID = 1;
            var fileName = string.Format("Export/OrderTransferResult_{0}_{1}.xlsx", requestID, string.Format("{0:yyyyMMddHHmmssffff}", DateTime.Now));
            var resultList = new List<OrderTransferRequestTransaction>
            {
                new OrderTransferRequestTransaction { SONumber = 1, FromWH = "07", ToWH = "09",Phase=OrderTransferRequestPhase.Transferred },
                new OrderTransferRequestTransaction { SONumber = 2, FromWH = "07", ToWH = "09",Phase=OrderTransferRequestPhase.Hold,ExceptionMessage="Q4S rollback&deduct failed." },
                new OrderTransferRequestTransaction { SONumber = 3, FromWH = "07", ToWH = "09" ,Phase=OrderTransferRequestPhase.Transferred},
                new OrderTransferRequestTransaction { SONumber = 4, FromWH = "07", ToWH = "09",Phase=OrderTransferRequestPhase.Q4S,ExceptionMessage="Transferred failed." }
            };
            IWorkbook workbook = new XSSFWorkbook();
            ISheet sheet = workbook.CreateSheet("Batch Transfer Result");
            
            int rowIndex = 1;
            resultList.ForEach(x =>
            {
                var dataRow = sheet.CreateRow(rowIndex);
                dataRow.CreateCell(0).SetCellValue(x.SONumber);
                dataRow.CreateCell(1).SetCellValue(x.FromWH);
                dataRow.CreateCell(2).SetCellValue(x.ToWH);
                dataRow.CreateCell(3).SetCellValue(x.Phase.ToString());
                dataRow.CreateCell(4).SetCellValue(x.ExceptionMessage);
                rowIndex++;
            });
            ConditionFormat(sheet, rowIndex);

            var headerStyle = GetHeaderStyle(workbook);
            IRow header = sheet.CreateRow(0);
            int colIndex = 0;
            new List<string> { "SONumber", "FromWH", "ToWH", "Phase", "ExceptionMessage" }.ForEach(y =>
              {
                  var cell = header.CreateCell(colIndex);
                  cell.SetCellValue(y);
                  cell.CellStyle = headerStyle;
                  sheet.AutoSizeColumn(colIndex);
                  colIndex++;
              });
            if (!Directory.Exists("Export"))
            {
                Directory.CreateDirectory("Export");
            }
            using (var fs = File.Create(fileName))
            {
                workbook.Write(fs);
            }
        }

        public static void ConditionFormat(ISheet sheet,int rowIndex)
        {
            ISheetConditionalFormatting sheetCF = sheet.SheetConditionalFormatting;

            IConditionalFormattingRule rule = sheetCF.CreateConditionalFormattingRule("$D2<>\"Transferred\"");
            IPatternFormatting fill = rule.CreatePatternFormatting();
            fill.FillBackgroundColor = (IndexedColors.Red.Index);
            fill.FillPattern = FillPattern.SolidForeground;

            CellRangeAddress[] regions = {
                CellRangeAddress.ValueOf($"D2:D{rowIndex}")
            };

            sheetCF.AddConditionalFormatting(regions, rule);
        }

        public static ICellStyle GetHeaderStyle(IWorkbook workbook)
        {
            ICellStyle headerStyle = workbook.CreateCellStyle();
            headerStyle.FillForegroundColor = IndexedColors.Grey25Percent.Index;
            headerStyle.FillPattern = FillPattern.SolidForeground;

            IFont font = workbook.CreateFont();
            font.IsBold = true;
            font.FontHeightInPoints = 12;

            headerStyle.SetFont(font);
            return headerStyle;
        }
    }
}

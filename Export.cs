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
            var result = new OrderTransferRequestMaster { FromWH = "07", ToWH = "09" };
            var resultList = new List<OrderTransferRequestTransaction>
            {
                new OrderTransferRequestTransaction { SONumber = 1, Phase=OrderTransferRequestPhase.Transferred },
                new OrderTransferRequestTransaction { SONumber = 2,Phase=OrderTransferRequestPhase.Hold,ExceptionMessage="Q4S rollback&deduct failed." },
                new OrderTransferRequestTransaction { SONumber = 3, Phase=OrderTransferRequestPhase.Transferred},
                new OrderTransferRequestTransaction { SONumber = 4, Phase=OrderTransferRequestPhase.Q4S,ExceptionMessage="Transferred failed." }
            };
            for (int i = 0; i < 1000; i++)
            {
                resultList.Add(new OrderTransferRequestTransaction { SONumber = i+5, Phase = OrderTransferRequestPhase.Transferred });
            }
            IWorkbook workbook = new XSSFWorkbook();
            ISheet sheet = workbook.CreateSheet("Batch Transfer Result");
            
            //body data
            int rowIndex = 2;
            var bodyStyle = GetBodyStyle(workbook);
            resultList.ForEach(x =>
            {
                var dataRow = sheet.CreateRow(rowIndex);

                var cell1 = dataRow.CreateCell(0);
                cell1.SetCellValue(x.SONumber);
                cell1.CellStyle = bodyStyle;
                var cell2 = dataRow.CreateCell(1);
                cell2.SetCellValue(x.Phase.ToString());
                cell2.CellStyle = bodyStyle;
                var cell3 = dataRow.CreateCell(2);
                cell3.SetCellValue(x.ExceptionMessage);
                cell3.CellStyle = bodyStyle;
                rowIndex++;
            });
            ConditionFormat(sheet, rowIndex);

            //title2
            var titleStyle = GetTitleStyle(workbook);
            IRow titleRow2 = sheet.CreateRow(1);
            int colIndex = 0;
            new List<string> { "SONumber", "Phase", "ExceptionMessage" }.ForEach(y =>
              {
                  var row2Cell = titleRow2.CreateCell(colIndex);
                  row2Cell.SetCellValue(y);
                  row2Cell.CellStyle = titleStyle;
                  sheet.AutoSizeColumn(colIndex);
                  colIndex++;
              });
            //title1
            IRow titleRow1 = sheet.CreateRow(0);
            var row1Cell = titleRow1.CreateCell(0);
            row1Cell.SetCellValue(string.Format("FromWH:{0}-->ToWH:{1}", result.FromWH, result.ToWH));
            row1Cell.CellStyle = titleStyle;
            sheet.AddMergedRegion(CellRangeAddress.ValueOf("A1:C1"));

            sheet.CreateFreezePane(0, 2);

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

            IConditionalFormattingRule rule = sheetCF.CreateConditionalFormattingRule("$B3<>\"Transferred\"");
            IPatternFormatting fill = rule.CreatePatternFormatting();
            fill.FillBackgroundColor = (IndexedColors.Red.Index);
            fill.FillPattern = FillPattern.SolidForeground;

            CellRangeAddress[] regions = {
                CellRangeAddress.ValueOf($"B3:B{rowIndex}")
            };

            sheetCF.AddConditionalFormatting(regions, rule);
        }

        public static ICellStyle GetTitleStyle(IWorkbook workbook)
        {
            ICellStyle titleStyle = workbook.CreateCellStyle();
            titleStyle.FillForegroundColor = IndexedColors.Grey25Percent.Index;
            titleStyle.FillPattern = FillPattern.SolidForeground;
            titleStyle.BorderBottom = BorderStyle.Thin;
            titleStyle.BorderTop = BorderStyle.Thin;
            titleStyle.BorderLeft = BorderStyle.Thin;
            titleStyle.BorderRight = BorderStyle.Thin;

            IFont font = workbook.CreateFont();
            font.IsBold = true;
            font.FontName = "Times New Roman";
            font.FontHeightInPoints = 13;

            titleStyle.SetFont(font);
            return titleStyle;
        }
        public static ICellStyle GetBodyStyle(IWorkbook workbook)
        {
            ICellStyle bodyStyle = workbook.CreateCellStyle();
            IFont font = workbook.CreateFont();
            font.FontName = "Times New Roman";
            font.FontHeightInPoints = 12;

            bodyStyle.SetFont(font);
            return bodyStyle;
        }
    }
}

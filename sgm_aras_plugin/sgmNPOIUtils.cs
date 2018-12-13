using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using NPOI.SS.UserModel;
using NPOI.XSSF.Util;
using NPOI.HSSF.Util;
using NPOI.HSSF.UserModel;
using NPOI.SS.Util;


namespace sgm_aras_plugin
{
    public class sgmNPOIUtils
    {
        public static ICellStyle currentStyle(string styleCase,HSSFWorkbook hssfworkbook)
        {
           
            ICellStyle style = hssfworkbook.CreateCellStyle();
            style.Alignment = NPOI.SS.UserModel.HorizontalAlignment.Center;
            style.VerticalAlignment = VerticalAlignment.Center;
            style.WrapText = true;

            IFont font = hssfworkbook.CreateFont();
            switch (styleCase)
            {
                case "title":
                    style.BorderBottom = NPOI.SS.UserModel.BorderStyle.Thin;
                    style.BorderLeft = NPOI.SS.UserModel.BorderStyle.Thin;
                    style.BorderRight = NPOI.SS.UserModel.BorderStyle.Thin;
                    style.BorderTop = NPOI.SS.UserModel.BorderStyle.Thin;
                    font.FontHeightInPoints = 11;
                    font.Boldweight = (short)NPOI.SS.UserModel.FontBoldWeight.Bold;
                    font.FontName = "Arial";
                    style.SetFont(font);
                    style.IsLocked = true;
                    break;

                case "mix":
                    style.BorderBottom = NPOI.SS.UserModel.BorderStyle.Thin;
                    style.BorderLeft = NPOI.SS.UserModel.BorderStyle.Thin;
                    style.BorderRight = NPOI.SS.UserModel.BorderStyle.Thin;
                    style.BorderTop = NPOI.SS.UserModel.BorderStyle.Thin;
                    style.DataFormat = HSSFDataFormat.GetBuiltinFormat("0%");
                    font.FontHeightInPoints = 11;
                    font.Boldweight = (short)NPOI.SS.UserModel.FontBoldWeight.Bold;
                    font.FontName = "Arial";
                    style.SetFont(font);
                    style.IsLocked = true;
                    break;

                case "note":
                    style.BorderBottom = NPOI.SS.UserModel.BorderStyle.Thin;
                    style.BorderLeft = NPOI.SS.UserModel.BorderStyle.Thin;
                    style.BorderRight = NPOI.SS.UserModel.BorderStyle.Thin;
                    style.BorderTop = NPOI.SS.UserModel.BorderStyle.Thin;
                    style.FillPattern = FillPattern.SolidForeground;
                    style.FillForegroundColor = NPOI.HSSF.Util.HSSFColor.Black.Index;
                    style.IsLocked = true;
                    font.FontHeightInPoints = 10;
                    font.FontName = "Arial";
                    font.Color = NPOI.HSSF.Util.HSSFColor.White.Index;
                    style.SetFont(font);
                    break;
                case "userMix":
                    style.BorderBottom = NPOI.SS.UserModel.BorderStyle.Thin;
                    style.BorderLeft = NPOI.SS.UserModel.BorderStyle.Thin;
                    style.BorderRight = NPOI.SS.UserModel.BorderStyle.Thin;
                    style.BorderTop = NPOI.SS.UserModel.BorderStyle.Thin;
                    style.DataFormat = HSSFDataFormat.GetBuiltinFormat("0%");
                    style.IsLocked = false;
               //     font.FontHeightInPoints = 10;
               //     font.FontName = "Arial";
                    style.SetFont(font);
                    break;
                case "userQPU":
                    style.BorderBottom = NPOI.SS.UserModel.BorderStyle.Thin;
                    style.BorderLeft = NPOI.SS.UserModel.BorderStyle.Thin;
                    style.BorderRight = NPOI.SS.UserModel.BorderStyle.Thin;
                    style.BorderTop = NPOI.SS.UserModel.BorderStyle.Thin;
                 //   style.DataFormat = HSSFDataFormat.GetBuiltinFormat("2");
                    style.IsLocked = false;
               //     font.FontHeightInPoints = 10;
              //      font.FontName = "Arial";
                    style.SetFont(font);
                    break;
                case "modelYear":
                    style.BorderBottom = NPOI.SS.UserModel.BorderStyle.Thin;
                    style.BorderLeft = NPOI.SS.UserModel.BorderStyle.Thin;
                    style.BorderRight = NPOI.SS.UserModel.BorderStyle.Thin;
                    style.BorderTop = NPOI.SS.UserModel.BorderStyle.Thin;
                    font.FontHeightInPoints = 10;
                    font.Boldweight = (short)NPOI.SS.UserModel.FontBoldWeight.Bold;
                    font.FontName = "Arial";
                    font.Color = NPOI.HSSF.Util.HSSFColor.Red.Index;
                    style.SetFont(font);
                    style.IsLocked = true;
                    break;

                case "lcr":
                    style.BorderBottom = NPOI.SS.UserModel.BorderStyle.Thin;
                    style.BorderLeft = NPOI.SS.UserModel.BorderStyle.Thin;
                    style.BorderRight = NPOI.SS.UserModel.BorderStyle.Thin;
                    style.BorderTop = NPOI.SS.UserModel.BorderStyle.Thin;
                    style.DataFormat = HSSFDataFormat.GetBuiltinFormat("0.00");
                    font.FontHeightInPoints = 10;
                    font.Boldweight = (short)NPOI.SS.UserModel.FontBoldWeight.Bold;
                    font.FontName = "Arial";
                    style.SetFont(font);
                    style.IsLocked = true;
                    break;
                case "formula":
                    font.FontHeightInPoints = 10;
                    font.FontName = "Arial";
                    font.Color = NPOI.HSSF.Util.HSSFColor.White.Index;
                    style.SetFont(font);
                    style.IsLocked = true;
                    break;
            }


            return style;

        }

        public static string ConvertColumnIndexToColumnName(int index)
        {
            index = index + 1;
            int system = 26;
            char[] digArray = new char[100];
            int i = 0;
            while (index > 0)
            {
                int mod = index % system;
                if (mod == 0)
                    mod = system;
                digArray[i++] = (char)(mod - 1 + 'A');
                index = (index - 1) / 26;
            }
            StringBuilder sb = new StringBuilder(i);
            for (int j = i - 1; j >= 0; j--)
            {
                sb.Append(digArray[j]);
            }

            return sb.ToString();
        }

        public static void CopyRow(ISheet sheet, int startRowIndex, int sourceRowIndex, int insertCount)
            {
                IRow sourceRow = sheet.GetRow(sourceRowIndex);
                int sourceCellCount = sourceRow.Cells.Count;

                //1. 批量移动行,清空插入区域
                sheet.ShiftRows(startRowIndex, //开始行
                                sheet.LastRowNum, //结束行
                                insertCount, //插入行总数
                                true,        //是否复制行高
                                false        //是否重置行高
                                );

                int startMergeCell = -1; //记录每行的合并单元格起始位置
                for (int i = startRowIndex; i < startRowIndex + insertCount; i++)
                {
                    IRow targetRow = null;
                    ICell sourceCell = null;
                    ICell targetCell = null;

                    targetRow = sheet.CreateRow(i);
                    targetRow.Height = sourceRow.Height;//复制行高

                    for (int m = sourceRow.FirstCellNum; m < sourceRow.LastCellNum; m++)
                    {
                        sourceCell = sourceRow.GetCell(m);
                        if (sourceCell == null)
                            continue;
                        targetCell = targetRow.CreateCell(m);
                        targetCell.CellStyle = sourceCell.CellStyle;//赋值单元格格式
                        targetCell.SetCellType(sourceCell.CellType);

                        //以下为复制模板行的单元格合并格式
                        if (sourceCell.IsMergedCell)
                        {
                            if (startMergeCell <= 0)
                                startMergeCell = m;
                            else if (startMergeCell > 0 && sourceCellCount == m + 1)
                            {
                                sheet.AddMergedRegion(new CellRangeAddress(i, i, startMergeCell, m));
                                startMergeCell = -1;
                            }
                        }
                        else
                        {
                            if (startMergeCell >= 0)
                            {
                                sheet.AddMergedRegion(new CellRangeAddress(i, i, startMergeCell, m - 1));
                                startMergeCell = -1;
                            }
                        }
                    }
                }
            }
     }

}

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using NPOI.SS.UserModel;
using NPOI.XSSF.Util;
using NPOI.HSSF.Util;
using NPOI.HSSF.UserModel;
using NPOI.HSSF.Util;

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
                    font.FontHeightInPoints = 10;
                    font.FontName = "Arial";
                    style.SetFont(font);
                    break;
                case "userQPU":
                    style.BorderBottom = NPOI.SS.UserModel.BorderStyle.Thin;
                    style.BorderLeft = NPOI.SS.UserModel.BorderStyle.Thin;
                    style.BorderRight = NPOI.SS.UserModel.BorderStyle.Thin;
                    style.BorderTop = NPOI.SS.UserModel.BorderStyle.Thin;
                    style.DataFormat = HSSFDataFormat.GetBuiltinFormat("0");
                    style.IsLocked = false;
                    font.FontHeightInPoints = 10;
                    font.FontName = "Arial";
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

    }
}

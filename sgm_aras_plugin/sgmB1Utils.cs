using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Aras.IOM;
using NPOI.HPSF;
using NPOI.HSSF.UserModel;
using NPOI.HSSF.Util;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using System.Collections;
using System.IO;
using System.Globalization;
using System.Windows.Forms; //测试弹框

namespace sgm_aras_plugin
{
    public class sgmB1Utils
    {
        static string sgm_EndRow;

        public string getEndRow(Innovator inn)
        {
            string endRow = "";
            Item variableItem = inn.newItem("Variable", "get");
            variableItem.setAttribute("select", "value,name");
            variableItem.setProperty("name", "sgm_B1EndRow");
            variableItem = variableItem.apply();


            if (variableItem.isError())
            {
                endRow = variableItem.getItemByIndex(0).getProperty("value");
            }
            else
            {
                endRow = "300";
            }

            return endRow;
        }
        public List<YearLCR> yearLCR = new List<YearLCR>();
        public static List<ModelLine> getAllModelLine(Innovator inn, Item modelyear, string is_stretch)
        {
            List<ModelLine> modelline = new List<ModelLine>();
            string aml = "<Item type='ModelYear' action='get' select='year,lcr,stretch_lcr'>";
            aml += "     <id>" + modelyear.getID() + "</id>";
            aml += "	<Relationships>";
            aml += "		<Item type='Year Model' select='related_id,sort_order,id,mix,package_lcr,stretch_package_lcr,comments'>";
            aml += "			<related_id>";
            aml += "				<Item type='ModelLine' select='umd,package'/>";
            aml += "            </related_id>";
            aml += "		</Item>";
            aml += "	</Relationships>";
            aml += "</Item>";

            Item temp = inn.newItem();
            temp.loadAML(aml);
            Item results = temp.apply();
            if (results.getItemCount() > 0)
            {
                Item result = results.getItemByIndex(0);
                Item ym = result.getItemsByXPath("Relationships/Item[@type ='Year Model']");
                //  Item ml= result.getItemsByXPath("Relationships/Item[@type ='Year Model']/related_id/Item[@type='ModelLine']");

                for (int i = 0; i < ym.getItemCount(); i++)
                {
                    ModelLine m = new ModelLine();
                    m.Comments = ym.getItemByIndex(i).getProperty("comments");
                    m.PackageLCR = ym.getItemByIndex(i).getProperty("package_lcr");
                    m.Mix = ym.getItemByIndex(i).getProperty("mix");
                    m.StretchPackageLCR = ym.getItemByIndex(i).getProperty("stretch_package_lcr");
                    Item mls = ym.getItemByIndex(i).getItemsByXPath("related_id/Item[@type='ModelLine']");
                    Item ml = mls.getItemByIndex(0);
                    m.UMD = ml.getProperty("umd");
                    m.Package = ml.getProperty("package");
                    m.IS_Stretch = is_stretch;
                    modelline.Add(m);

                }
            }
            return modelline;
        }

        public static List<ModelYear> getAllModelYear(Innovator inn, Item country, string is_stretch)
        {
            List<ModelYear> modelyear = new List<ModelYear>();
            string aml = "<Item type='Location' action='get' select='id,name'>";
            aml += "    <id>" + country.getID() + "</id>";
            aml += "    <Relationships>";
            aml += "		<Item type='Location Year' select='related_id,sort_order,id'>";
            aml += "			<related_id>";
            aml += "				<Item type='ModelYear' select='year,lcr,stretch_lcr' />";
            aml += "            </related_id>";
            aml += "        </Item>";
            aml += "    </Relationships>";
            aml += "</Item>";

            Item temp = inn.newItem();
            temp.loadAML(aml);
            Item results = temp.apply();
            if (results.getItemCount() > 0)
            {
                Item result = results.getItemByIndex(0);
                Item my = result.getItemsByXPath("Relationships/Item[@type ='Location Year']/related_id/Item[@type='ModelYear']");
                for (int i = 0; i < my.getItemCount(); i++)
                {
                    ModelYear m = new ModelYear();
                    m.LCR = my.getItemByIndex(i).getProperty("lcr");
                    m.Year = my.getItemByIndex(i).getProperty("year");
                    m.StretchLCR = my.getItemByIndex(i).getProperty("stretch_lcr") ?? "0";
                    List<ModelLine> ml = getAllModelLine(inn, my.getItemByIndex(i), is_stretch);
                    m.ModelLine = ml;
                    modelyear.Add(m);
                }

            }

            return modelyear;
        }
        public static List<CountryGroup> getAllCountryGroup(Innovator inn, Item B1, string is_stretch)
        {
            List<CountryGroup> countryGroup = new List<CountryGroup>();
            //string aml = "";
            string aml = "<Item type='B1Template' action='get' select='id,program_code'>";
            aml += "     <id>" + B1.getID() + "</id>";
            aml += "    <Relationships>";
            aml += "        <Item type='B1 Location' action='get' initial_action='GetItemConfig' select='related_id,sort_order'>";
            aml += "            <related_id>";
            aml += "  	            <Item type='Location' select='id,name'>";
            aml += "		        </Item>";
            aml += "            </related_id>";
            aml += "        </Item>";
            aml += "    </Relationships>";
            aml += "</Item>";

            Item temp = inn.newItem();
            temp.loadAML(aml);
            Item results = temp.apply();

            if (results.getItemCount() > 0)
            {
                Item result = results.getItemByIndex(0);
                Item Locations = result.getItemsByXPath("Relationships/Item[@type ='B1 Location']/related_id /Item[@type ='Location']");
                for (int i = 0; i < Locations.getItemCount(); i++)
                {
                    CountryGroup cg = new CountryGroup();
                    cg.Country = Locations.getItemByIndex(i).getProperty("name");

                    List<ModelYear> my = getAllModelYear(inn, Locations.getItemByIndex(i), is_stretch);
                    cg.ModelYear = my;
                    countryGroup.Add(cg);
                }
            }
            return countryGroup;
        }
        public static List<B1Content> getAllB1Content(Innovator inn, string cartID)
        {
            List<B1Content> b1 = new List<B1Content>();

            string aml = "<Item type='B1TemplateCart' action='get' select='id,source_id,related_id'> ";
            aml += "    <id>" + cartID + "</id>";
            aml += "    <Relationships>";
            aml += "         <Item type='B1 Template Cart' select='related_id'>";
            aml += "            <related_id>";
            aml += "                <Item type='B1Template' action='get' select='id,program_code,is_stretch '>";
            aml += "                </Item>";
            aml += "            </related_id>";
            aml += "          </Item>";
            aml += "    </Relationships>";
            aml += "</Item>";


            Item temp = inn.newItem();
            temp.loadAML(aml);
            Item results = temp.apply();
            if (results.getItemCount() > 0)
            {
                Item result = results.getItemByIndex(0);
                Item b1Template = result.getItemsByXPath("Relationships/Item[@type ='B1 Template Cart']/related_id /Item[@type ='B1Template']");

                for (int i = 0; i < b1Template.getItemCount(); i++)
                {
                    B1Content b1Content = new B1Content();
                    b1Content.Programe = b1Template.getItemByIndex(i).getProperty("program_code");
                    b1Content.Is_Stretch = b1Template.getItemByIndex(i).getProperty("is_stretch");
                    List<CountryGroup> countrygroup = getAllCountryGroup(inn, b1Template.getItemByIndex(i), b1Template.getItemByIndex(i).getProperty("is_stretch"));
                    b1Content.CountryGroup = countrygroup;
                    b1.Add(b1Content);

                }

            }


            return b1;
        }

        public static List<ModelYear>
            getAllExportModelYear(B1Content b1Content)
        {
            List<ModelYear> modelyear = new List<ModelYear>();
            List<ModelYear> tempMY = new List<ModelYear>();

            for (int i = 0; i < b1Content.CountryGroup.Count; i++)
            {
                CountryGroup cg = b1Content.CountryGroup[i];
                if (cg.Country != "China")  //不为china的才加入出口综合
                {
                    for (int j = 0; j < cg.ModelYear.Count; j++)
                    {
                        ModelYear my = cg.ModelYear[j];
                        tempMY.Add(my);
                    }
                }

            }

            List<string> templist = tempMY.Select(p => p.Year).Distinct().ToList();  //year去重
            templist.Sort();

            for (int m = 0; m < templist.Count; m++)
            {
                string year = templist[m];
                List<ModelYear> thisModelYear = tempMY.FindAll(s => s.Year == year);
                List<ModelLine> tempML = new List<ModelLine>();
                ModelYear nMY = new ModelYear();
                nMY.Year = year;

                double lcr = 0;
                double stretchLCR = 0;
                for (int n = 0; n < thisModelYear.Count; n++)
                {
                    ModelYear thisMY = thisModelYear[n];
                    tempML = tempML.Concat(thisMY.ModelLine).ToList<ModelLine>();
                    lcr = lcr + double.Parse(thisMY.LCR);
                    stretchLCR = stretchLCR + double.Parse(thisMY.StretchLCR);

                }
                nMY.LCR = lcr.ToString();
                nMY.StretchLCR = stretchLCR.ToString();

                List<string> packagelist = tempML.Select(p => p.Package).Distinct().ToList();  //package去重
                packagelist.Sort();

                List<ModelLine> tempML3 = new List<ModelLine>();

                decimal sumMix = 0;
                for (int x = 0; x < packagelist.Count; x++)
                {
                    string package = packagelist[x];
                    ModelLine ml = tempML.Find(s => s.Package == package);
                    string umd = ml.UMD;
                    List<ModelLine> tempML2 = tempML.FindAll(s => s.Package == package);
                    double pLCR = 0;
                    double spLCR = 0;
                    for (int y = 0; y < tempML2.Count; y++)
                    {
                        pLCR = pLCR + double.Parse(tempML2[y].PackageLCR);
                        spLCR = spLCR + double.Parse(tempML2[y].StretchPackageLCR);
                    }
                    ModelLine nML = new ModelLine();
                    decimal mix = 0;
                    if (x == packagelist.Count - 1)
                    {

                        mix = 1 - sumMix;

                    }
                    else
                    {
                        mix = decimal.Parse(Math.Round(pLCR / lcr, 2).ToString());
                        sumMix = sumMix + mix;
                    }


                    nML.UMD = umd;
                    nML.Package = package;
                    mix = mix * 100;
                    nML.Mix = mix.ToString();
                    nML.PackageLCR = pLCR.ToString();
                    nML.StretchPackageLCR = spLCR.ToString();
                    tempML3.Add(nML);

                }

                nMY.ModelLine = tempML3;
                modelyear.Add(nMY);
            }


            return modelyear;
        }

        public string AppendixB1Excel(Innovator inn, string cartId, string templatefolder, string templatefile)
        {
            //创建工作簿对象
            HSSFWorkbook hssfworkbook;
            string xlsT = templatefolder + templatefile; //服务器模板本地路径

            // 打开模板文件到文件流中
            using (FileStream file = new FileStream(xlsT, FileMode.Open, FileAccess.Read))
            {
                //将文件流中模板加载到工作簿对象中
                hssfworkbook = new HSSFWorkbook(file);
            }

            OutPutXLS(hssfworkbook, inn, cartId);

            List<B1Content> b1Content = new List<B1Content>();
            b1Content = getAllB1Content(inn, cartId);
            string pgCollection = "";
            if (b1Content.Count > 0)
            {
                for (int i = 0; i < b1Content.Count; i++)
                {
                    pgCollection = pgCollection + "_" + b1Content[i].Programe;
                }
            }
            string tradeTime = DateTime.Now.ToString("yyyyMMddHHmmss", DateTimeFormatInfo.InvariantInfo);

            string fpath = templatefolder + "AppendixB1-" + pgCollection + "-" + tradeTime + ".xls";

            FileStream outfile = new FileStream(fpath, FileMode.Create);
            hssfworkbook.Write(outfile);
            outfile.Close();  //关闭文件流  
            hssfworkbook.Close();



            return fpath;

        }

        public int ExportCount(B1Content b1content)
        {
            List<CountryGroup> cg = b1content.CountryGroup;

            List<CountryGroup> export = cg.FindAll(delegate (CountryGroup c) { return c.Country != "China"; });

            List<string> countryexport = export.Select(p => p.Country).Distinct().ToList(); //去重，暂时不用

            return export.Count;
        }

        bool is_export(List<B1Content> b1content)
        {
            bool f = false;
            for (int i = 0; i < b1content.Count; i++)
            {

                int exportcount = ExportCount(b1content[i]);
                if (exportcount > 2)
                {
                    f = f | true;
                }
            }


            return f;
        }
        public void OutPutXLS(HSSFWorkbook hssfworkbook, Innovator inn, string cartID)
        {
            sgm_EndRow = getEndRow(inn);

            List<B1Content> b1Content = new List<B1Content>();
            b1Content = getAllB1Content(inn, cartID);
            if (is_export(b1Content))
            {
                OutPutXLSExport(hssfworkbook, b1Content);
            }
            else
            {
                OutPutXLSLocal(hssfworkbook, b1Content);
            }


        }

        public void OutPutXLSLocal(HSSFWorkbook hssfworkbook, List<B1Content> b1Content)
        {
            ISheet sheet1;
            sheet1 = hssfworkbook.GetSheet("B1 Form");

            int startColIndex = 4;


            string v1 = sheet1.GetRow(0).GetCell(5).ToString();
            string v2 = sheet1.GetRow(0).GetCell(6).ToString();
            if (v1.Trim() != "")
            {
                startColIndex = int.Parse(v1) - 1;
            }

            int colIndex = startColIndex;
            v2 = v2.Replace("，", ",");
            string[] v3 = v2.Split(',');

            int PGRow = 8;
            int CGRow = 9;
            int MYRow = 10;
            int UMDRow = 11;
            int PRow = 12;
            int mixRow = 13;
            int CRow = 14;
            int noteRow = 15;
            int inputRow = 17;//此处不是Index


            if (v3.Length == 9)
            {
                PGRow = int.Parse(v3[0]) - 1;
                CGRow = int.Parse(v3[1]) - 1;
                MYRow = int.Parse(v3[2]) - 1;
                UMDRow = int.Parse(v3[3]) - 1;
                PRow = int.Parse(v3[4]) - 1;
                mixRow = int.Parse(v3[5]) - 1;
                CRow = int.Parse(v3[6]) - 1;
                noteRow = int.Parse(v3[7]) - 1;
                inputRow = int.Parse(v3[8]);//此处不是Index

            }

            int yearStart = colIndex;
            int countryStart = colIndex;
            int programeStart = colIndex;

            ICellStyle titleStyle = sgmNPOIUtils.currentStyle("title", hssfworkbook);
            ICellStyle style1 = sgmNPOIUtils.currentStyle("note", hssfworkbook);


            ICellStyle mixStyle = sgmNPOIUtils.currentStyle("userMix", hssfworkbook);
            ICellStyle QPUStyle = sgmNPOIUtils.currentStyle("userQPU", hssfworkbook);
            ICellStyle lcrStyle = sgmNPOIUtils.currentStyle("lcr", hssfworkbook);
            ICellStyle forStyle = sgmNPOIUtils.currentStyle("formula", hssfworkbook);


            ICell fCell = sheet1.GetRow(0).CreateCell(4);
            fCell.CellStyle = forStyle;
            string fCol = "";

            if (b1Content.Count > 0)
            {
                for (int i = 0; i < b1Content.Count; i++)
                {
                    B1Content b1 = b1Content[i];
                    List<CountryGroup> countrygroup = b1.CountryGroup;
                    for (int j = 0; j < countrygroup.Count; j++)
                    {
                        CountryGroup cg = countrygroup[j];
                        List<ModelYear> modelyear = cg.ModelYear;
                        for (int m = 0; m < modelyear.Count; m++)
                        {
                            ModelYear my = modelyear[m];
                            List<ModelLine> modelline = my.ModelLine;
                            YearLCR yl = new YearLCR(my.Year);
                            string fn = "SUM(";
                            string stretchFn = "SUM(";

                            for (int n = 0; n < modelline.Count; n++)
                            {
                                ModelLine ml = modelline[n];

                                sheet1.AddMergedRegion(new CellRangeAddress(UMDRow, UMDRow, colIndex, colIndex + 1));
                                sheet1.GetRow(UMDRow).CreateCell(colIndex).SetCellValue(ml.UMD);
                                sheet1.GetRow(UMDRow).GetCell(colIndex).CellStyle = titleStyle;
                                sheet1.GetRow(UMDRow).CreateCell(colIndex + 1).CellStyle = titleStyle;

                                sheet1.AddMergedRegion(new CellRangeAddress(PRow, PRow, colIndex, colIndex + 1));
                                sheet1.GetRow(PRow).CreateCell(colIndex).SetCellValue(ml.Package);
                                sheet1.GetRow(PRow).GetCell(colIndex).CellStyle = titleStyle;
                                sheet1.GetRow(PRow).CreateCell(colIndex + 1).CellStyle = titleStyle;


                                sheet1.AddMergedRegion(new CellRangeAddress(mixRow, mixRow, colIndex, colIndex + 1));
                                sheet1.GetRow(mixRow).CreateCell(colIndex).SetCellValue(ml.Mix + "%");
                                sheet1.GetRow(mixRow).GetCell(colIndex).CellStyle = titleStyle;
                                sheet1.GetRow(mixRow).CreateCell(colIndex + 1).CellStyle = titleStyle;

                                sheet1.AddMergedRegion(new CellRangeAddress(CRow, CRow, colIndex, colIndex + 1));
                                sheet1.GetRow(CRow).CreateCell(colIndex).SetCellValue("");
                                sheet1.GetRow(CRow).GetCell(colIndex).CellStyle = titleStyle;
                                sheet1.GetRow(CRow).CreateCell(colIndex + 1).CellStyle = titleStyle;


                                sheet1.GetRow(noteRow).CreateCell(colIndex).SetCellValue("mix");
                                sheet1.GetRow(noteRow).GetCell(colIndex).CellStyle = style1;
                                sheet1.GetRow(noteRow).CreateCell(colIndex + 1).SetCellValue("QPU");
                                sheet1.GetRow(noteRow).GetCell(colIndex + 1).CellStyle = style1;

                                string colNm1 = sgmNPOIUtils.ConvertColumnIndexToColumnName(colIndex);
                                string colNm2 = sgmNPOIUtils.ConvertColumnIndexToColumnName(colIndex + 1);

                                //fn = fn + ml.PackageLCR + "*"+colNm1 + inputRow.ToString()+"*" + colNm2 + inputRow.ToString()+"+";
                                //stretchFn = stretchFn+ ml.StretchPackageLCR + "*" + colNm1 + inputRow.ToString()+"*" + colNm2 + inputRow.ToString() + "+";
                                fn = fn + ml.PackageLCR + "*" + colNm1 + "{0}*" + colNm2 + "{0}+";
                                stretchFn = stretchFn + ml.StretchPackageLCR + "*" + colNm1 + "{0}*" + colNm2 + "{0}+";

                                colIndex = colIndex + 2;

                            }
                            fn = fn.Substring(0, fn.Length - 1);
                            fn = fn + ")";
                            stretchFn = stretchFn.Substring(0, stretchFn.Length - 1);

                            stretchFn = stretchFn + ")";
                            yl.Formula = fn;
                            yl.StretchFormula = stretchFn;
                            yearLCR.Add(yl);

                            sheet1.AddMergedRegion(new CellRangeAddress(MYRow, MYRow, yearStart, colIndex - 1));
                            sheet1.GetRow(MYRow).CreateCell(yearStart).SetCellValue(my.Year);
                            sheet1.GetRow(MYRow).GetCell(yearStart).CellStyle = titleStyle;
                            for (int x = yearStart + 1; x < colIndex; x++)
                            {
                                sheet1.GetRow(MYRow).CreateCell(x);
                                sheet1.GetRow(MYRow).GetCell(x).CellStyle = titleStyle;
                            }

                            yearStart = colIndex;

                        }


                        sheet1.AddMergedRegion(new CellRangeAddress(CGRow, CGRow, countryStart, yearStart - 1));
                        sheet1.GetRow(CGRow).CreateCell(countryStart).SetCellValue(cg.Country);
                        sheet1.GetRow(CGRow).GetCell(countryStart).CellStyle = titleStyle;
                        for (int y = countryStart + 1; y < yearStart; y++)
                        {
                            sheet1.GetRow(CGRow).CreateCell(y);
                            sheet1.GetRow(CGRow).GetCell(y).CellStyle = titleStyle;
                        }

                        countryStart = yearStart;
                    }

                    sheet1.AddMergedRegion(new CellRangeAddress(PGRow, PGRow, programeStart, countryStart - 1));
                    sheet1.GetRow(PGRow).CreateCell(programeStart).SetCellValue(b1.Programe);
                    sheet1.GetRow(PGRow).GetCell(programeStart).CellStyle = titleStyle;
                    for (int z = programeStart + 1; z < countryStart; z++)
                    {
                        sheet1.GetRow(PGRow).CreateCell(z);
                        sheet1.GetRow(PGRow).GetCell(z).CellStyle = titleStyle;
                    }

                    programeStart = countryStart;

                }

                List<YearLCR> newYearLCR = splitList(yearLCR);

                int myColIndex = colIndex;

                ICellStyle myStyle = sgmNPOIUtils.currentStyle("modelYear", hssfworkbook);
                sheet1.GetRow(CRow).CreateCell(colIndex).SetCellValue("Model Year");
                int columnWidth = sheet1.GetColumnWidth(colIndex);
                sheet1.SetColumnWidth(colIndex, columnWidth * 2);
                sheet1.GetRow(CRow).GetCell(colIndex).CellStyle = myStyle;
                sheet1.GetRow(noteRow).CreateCell(colIndex);
                sheet1.GetRow(noteRow).GetCell(colIndex).CellStyle = style1;
                colIndex = colIndex + 1;
                int endInputRow = int.Parse(sgm_EndRow);
                for (int q = 0; q < endInputRow - 1; q++)
                {
                    sgmNPOIUtils.CopyRow(sheet1, inputRow + q, inputRow - 1, 1);
                }


                for (int a = 0; a < newYearLCR.Count; a++)
                {
                    sheet1.GetRow(CRow).CreateCell(colIndex).SetCellValue(newYearLCR[a].Year);
                    sheet1.SetColumnWidth(colIndex, columnWidth * 2);
                    sheet1.GetRow(CRow).GetCell(colIndex).CellStyle = titleStyle;
                    sheet1.GetRow(noteRow).CreateCell(colIndex);
                    sheet1.GetRow(noteRow).GetCell(colIndex).CellStyle = style1;
                    //sheet1.GetRow(inputRow - 1).CreateCell(colIndex).SetCellFormula(newYearLCR[a].Formula);

                    string xFn = newYearLCR[a].Formula;
                    xFn = "IF(AND(TRIM(A{0})=\"\",TRIM(B{0})=\"\"),\"\"," + xFn + ")";
                    for (int f = 0; f < endInputRow; f++)
                    {
                        string yFn = string.Format(xFn, inputRow + f);
                        int r = inputRow - 1 + f;
                        sheet1.GetRow(r).CreateCell(colIndex).SetCellFormula(yFn);
                    }


                    string colNm = sgmNPOIUtils.ConvertColumnIndexToColumnName(colIndex);
                    fCol = fCol + colNm + ",";

                    colIndex = colIndex + 1;
                }



                if (is_stretch(b1Content))
                {
                    for (int b = 0; b < newYearLCR.Count; b++)
                    {
                        sheet1.GetRow(CRow).CreateCell(colIndex).SetCellValue(newYearLCR[b].Year + "\n" + "stretch");
                        sheet1.SetColumnWidth(colIndex, columnWidth * 2);
                        sheet1.GetRow(CRow).GetCell(colIndex).CellStyle = titleStyle;
                        sheet1.GetRow(noteRow).CreateCell(colIndex);
                        sheet1.GetRow(noteRow).GetCell(colIndex).CellStyle = style1;
                        //    sheet1.GetRow(inputRow - 1).CreateCell(colIndex).SetCellFormula(newYearLCR[b].StretchFormula);

                        //          int endInputRow = int.Parse(sgm_EndRow);
                        string xFn = newYearLCR[b].Formula;
                        xFn = "IF(AND(TRIM(A{0})=\"\",TRIM(B{0})=\"\"),\"\"," + xFn + ")";
                        for (int f = 0; f < endInputRow; f++)
                        {
                            string yFn = string.Format(xFn, inputRow + f);
                            int r = inputRow - 1 + f;
                            //        sgmNPOIUtils.CopyRow(sheet1, r - 1, r, 1);

                            sheet1.GetRow(r).CreateCell(colIndex).SetCellFormula(yFn);
                        }

                        string colNm = sgmNPOIUtils.ConvertColumnIndexToColumnName(colIndex);
                        fCol = fCol + colNm + ",";

                        colIndex = colIndex + 1;
                    }
                }



                sheet1.AddMergedRegion(new CellRangeAddress(mixRow, mixRow, myColIndex, colIndex - 1));
                sheet1.GetRow(mixRow).CreateCell(myColIndex).SetCellValue("LCR");
                sheet1.GetRow(mixRow).GetCell(myColIndex).CellStyle = titleStyle;
                for (int a = myColIndex + 1; a < colIndex; a++)
                {
                    sheet1.GetRow(mixRow).CreateCell(a);
                    sheet1.GetRow(mixRow).GetCell(a).CellStyle = titleStyle;
                }

                for (int c = inputRow - 1; c < inputRow - 1 + endInputRow; c++)
                {
                    for (int d = startColIndex; d < colIndex; d++)
                    {
                        //    if (c == inputRow - 1 && d > myColIndex)
                        if (d > myColIndex)
                        {
                            sheet1.GetRow(c).GetCell(d).CellStyle = lcrStyle; ;
                        }
                        else if (d == myColIndex)
                        {
                            sheet1.GetRow(c).CreateCell(d).CellStyle = style1;
                        }
                        else
                        {
                            if (d % 2 == 0)
                            {
                                sheet1.GetRow(c).CreateCell(d).CellStyle = mixStyle;
                            }
                            else
                            {
                                sheet1.GetRow(c).CreateCell(d).CellStyle = QPUStyle;
                            }

                        }

                    }
                }
                sheet1.GetRow(0).GetCell(5).SetCellValue("");
                sheet1.GetRow(0).GetCell(6).SetCellValue("");

                fCol = fCol.Substring(0, fCol.Length - 1);
                fCell.SetCellValue(fCol);
                sheet1.CreateFreezePane(4, 0, 5, 0);
                sheet1.ProtectSheet("password");//设置密码保护
            }
        }

        public void OutPutXLSExport(HSSFWorkbook hssfworkbook, List<B1Content> b1Content)
        {
            ISheet sheet1;
            sheet1 = hssfworkbook.GetSheet("B1 Form");
            int startColIndex = 4;

            int endInputRow = int.Parse(sgm_EndRow);

            string v1 = sheet1.GetRow(0).GetCell(5).ToString();
            string v2 = sheet1.GetRow(0).GetCell(6).ToString();
            if (v1.Trim() != "")
            {
                startColIndex = int.Parse(v1) - 1;
            }

            int colIndex = startColIndex;
            v2 = v2.Replace("，", ",");
            string[] v3 = v2.Split(',');

            int PGRow = 8;
            int CGRow = 9;
            int MYRow = 10;
            int UMDRow = 11;
            int PRow = 12;
            int mixRow = 13;
            int CRow = 14;
            int noteRow = 15;
            int inputRow = 17;//此处不是Index


            if (v3.Length == CGRow)
            {
                PGRow = int.Parse(v3[0]) - 1;
                CGRow = int.Parse(v3[1]) - 1;
                MYRow = int.Parse(v3[2]) - 1;
                UMDRow = int.Parse(v3[3]) - 1;
                PRow = int.Parse(v3[4]) - 1;
                mixRow = int.Parse(v3[5]) - 1;
                CRow = int.Parse(v3[6]) - 1;
                noteRow = int.Parse(v3[7]) - 1;
                inputRow = int.Parse(v3[PGRow]);//此处不是Index

            }

            int yearStart = colIndex;
            int countryStart = colIndex;
            int programeStart = colIndex;

            ICellStyle titleStyle = sgmNPOIUtils.currentStyle("title", hssfworkbook);
            ICellStyle mStyle = sgmNPOIUtils.currentStyle("mix", hssfworkbook);

            ICellStyle style1 = sgmNPOIUtils.currentStyle("note", hssfworkbook);


            ICellStyle mixStyle = sgmNPOIUtils.currentStyle("userMix", hssfworkbook);
            ICellStyle QPUStyle = sgmNPOIUtils.currentStyle("userQPU", hssfworkbook);
            ICellStyle lcrStyle = sgmNPOIUtils.currentStyle("lcr", hssfworkbook);
            ICellStyle forStyle = sgmNPOIUtils.currentStyle("formula", hssfworkbook);

            ICell fCell = sheet1.GetRow(0).CreateCell(4);
            fCell.CellStyle = forStyle;
            string fCol = "";


            List<B1Content> exportB1 = new List<B1Content>();
            List<B1Content> localB1 = new List<B1Content>();

            List<int> exportCols = new List<int>();
            List<int> mixCols = new List<int>();
            List<int> QPUCols = new List<int>();

            if (b1Content.Count > 0)
            {

                for (int q = 0; q < endInputRow - 1; q++)
                {
                    sgmNPOIUtils.CopyRow(sheet1, inputRow + q, inputRow - 1, 1);
                }


                for (int i = 0; i < b1Content.Count; i++)
                {
                    if (ExportCount(b1Content[i]) > 2)  //原先为>1，B1表中大于2的才是出口综合，修改此处
                    {
                        exportB1.Add(b1Content[i]);
                    }
                    else
                    {
                        localB1.Add(b1Content[i]);
                    }
                }

                if (localB1.Count > 0)
                {
                    for (int i = 0; i < localB1.Count; i++)
                    {
                        B1Content b1 = localB1[i];
                        List<CountryGroup> countrygroup = b1.CountryGroup;
                        for (int j = 0; j < countrygroup.Count; j++)
                        {
                            CountryGroup cg = countrygroup[j];
                            List<ModelYear> modelyear = cg.ModelYear;
                            for (int m = 0; m < modelyear.Count; m++)
                            {
                                ModelYear my = modelyear[m];
                                List<ModelLine> modelline = my.ModelLine;
                                YearLCR yl = new YearLCR(my.Year);
                                string fn = "SUM(";
                                string stretchFn = "SUM(";

                                for (int n = 0; n < modelline.Count; n++)
                                {
                                    ModelLine ml = modelline[n];

                                    sheet1.AddMergedRegion(new CellRangeAddress(UMDRow, UMDRow, colIndex, colIndex + 1));
                                    sheet1.GetRow(UMDRow).CreateCell(colIndex).SetCellValue(ml.UMD);
                                    sheet1.GetRow(UMDRow).GetCell(colIndex).CellStyle = titleStyle;
                                    sheet1.GetRow(UMDRow).CreateCell(colIndex + 1).CellStyle = titleStyle;

                                    sheet1.AddMergedRegion(new CellRangeAddress(PRow, PRow, colIndex, colIndex + 1));
                                    sheet1.GetRow(PRow).CreateCell(colIndex).SetCellValue(ml.Package);
                                    sheet1.GetRow(PRow).GetCell(colIndex).CellStyle = titleStyle;
                                    sheet1.GetRow(PRow).CreateCell(colIndex + 1).CellStyle = titleStyle;


                                    sheet1.AddMergedRegion(new CellRangeAddress(mixRow, mixRow, colIndex, colIndex + 1));
                                    sheet1.GetRow(mixRow).CreateCell(colIndex).CellStyle = mStyle;
                                    sheet1.GetRow(mixRow).CreateCell(colIndex + 1).CellStyle = mStyle;
                                    sheet1.GetRow(mixRow).GetCell(colIndex).SetCellValue(double.Parse(ml.Mix) / 100);


                                    sheet1.AddMergedRegion(new CellRangeAddress(CRow, CRow, colIndex, colIndex + 1));
                                    sheet1.GetRow(CRow).CreateCell(colIndex).SetCellValue("");
                                    sheet1.GetRow(CRow).GetCell(colIndex).CellStyle = titleStyle;
                                    sheet1.GetRow(CRow).CreateCell(colIndex + 1).CellStyle = titleStyle;

                                    //    sheet1.AddMergedRegion(new CellRangeAddress(noteRow, noteRow, colIndex, colIndex + 1));
                                    sheet1.GetRow(noteRow).CreateCell(colIndex).SetCellValue("mix");
                                    sheet1.GetRow(noteRow).GetCell(colIndex).CellStyle = style1;
                                    mixCols.Add(colIndex);
                                    sheet1.GetRow(noteRow).CreateCell(colIndex + 1).SetCellValue("QPU");
                                    sheet1.GetRow(noteRow).GetCell(colIndex + 1).CellStyle = style1;
                                    QPUCols.Add(colIndex + 1);

                                    string colNm1 = sgmNPOIUtils.ConvertColumnIndexToColumnName(colIndex);
                                    string colNm2 = sgmNPOIUtils.ConvertColumnIndexToColumnName(colIndex + 1);

                                    //     fn = fn + ml.PackageLCR + "*" + colNm1 + inputRow.ToString() + "*" + colNm2 + inputRow.ToString() + "+";
                                    //      stretchFn = stretchFn + ml.StretchPackageLCR + "*" + colNm1 + inputRow.ToString() + "*" + colNm2 + inputRow.ToString() + "+";

                                    fn = fn + ml.PackageLCR + "*" + colNm1 + "{0}*" + colNm2 + "{0}+";
                                    stretchFn = stretchFn + ml.StretchPackageLCR + "*" + colNm1 + "{0}*" + colNm2 + "{0}+";


                                    colIndex = colIndex + 2;

                                }
                                fn = fn.Substring(0, fn.Length - 1);
                                fn = fn + ")";
                                stretchFn = stretchFn.Substring(0, stretchFn.Length - 1);

                                stretchFn = stretchFn + ")";
                                yl.Formula = fn;
                                yl.StretchFormula = stretchFn;
                                yearLCR.Add(yl);

                                sheet1.AddMergedRegion(new CellRangeAddress(MYRow, MYRow, yearStart, colIndex - 1));
                                sheet1.GetRow(MYRow).CreateCell(yearStart).SetCellValue(my.Year);
                                sheet1.GetRow(MYRow).GetCell(yearStart).CellStyle = titleStyle;
                                for (int x = yearStart + 1; x < colIndex; x++)
                                {
                                    sheet1.GetRow(MYRow).CreateCell(x);
                                    sheet1.GetRow(MYRow).GetCell(x).CellStyle = titleStyle;
                                }

                                yearStart = colIndex;

                            }


                            sheet1.AddMergedRegion(new CellRangeAddress(CGRow, CGRow, countryStart, yearStart - 1));
                            sheet1.GetRow(CGRow).CreateCell(countryStart).SetCellValue(cg.Country);
                            sheet1.GetRow(CGRow).GetCell(countryStart).CellStyle = titleStyle;
                            for (int y = countryStart + 1; y < yearStart; y++)
                            {
                                sheet1.GetRow(CGRow).CreateCell(y);
                                sheet1.GetRow(CGRow).GetCell(y).CellStyle = titleStyle;
                            }

                            countryStart = yearStart;
                        }

                        sheet1.AddMergedRegion(new CellRangeAddress(PGRow, PGRow, programeStart, countryStart - 1));
                        sheet1.GetRow(PGRow).CreateCell(programeStart).SetCellValue(b1.Programe);
                        sheet1.GetRow(PGRow).GetCell(programeStart).CellStyle = titleStyle;
                        for (int z = programeStart + 1; z < countryStart; z++)
                        {
                            sheet1.GetRow(PGRow).CreateCell(z);
                            sheet1.GetRow(PGRow).GetCell(z).CellStyle = titleStyle;
                        }

                        programeStart = countryStart;

                    }
                }

                if (exportB1.Count > 0)
                {
                    for (int i = 0; i < exportB1.Count; i++)
                    {
                        List<ExportModel> emList = new List<ExportModel>();
                        List<ExportModel> exportModels = new List<ExportModel>();
                        B1Content b1 = exportB1[i];


                        List<ModelYear> myList = getAllExportModelYear(b1);  //综合版部分

                        for (int m = 0; m < myList.Count; m++)
                        {
                            ModelYear my = myList[m];
                            List<ModelLine> modelline = my.ModelLine;

                            for (int n = 0; n < modelline.Count; n++)
                            {
                                ExportModel em = new ExportModel();
                                ModelLine ml = modelline[n];

                                em.UDM = ml.UMD;
                                em.Package = ml.Package;
                                //        em.ColIndex = colIndex;
                                em.Formula = "";
                                em.Year = my.Year;
                                //       exportCols.Add(colIndex);

                                emList.Add(em);

                                //int colWidth = sheet1.GetColumnWidth(colIndex);
                                //sheet1.SetColumnWidth(colIndex, colWidth * 2);

                                //sheet1.GetRow(UMDRow).CreateCell(colIndex).SetCellValue(ml.UMD);
                                //sheet1.GetRow(UMDRow).GetCell(colIndex).CellStyle = titleStyle;
                                ////sheet1.GetRow(UMDRow).CreateCell(colIndex + 1).CellStyle = titleStyle;

                                //sheet1.GetRow(PRow).CreateCell(colIndex).SetCellValue(ml.Package);
                                //sheet1.GetRow(PRow).GetCell(colIndex).CellStyle = titleStyle;
                                ////sheet1.GetRow(PRow).CreateCell(colIndex + 1).CellStyle = titleStyle;

                                //sheet1.GetRow(mixRow).CreateCell(colIndex).CellStyle = mStyle;
                                ////sheet1.GetRow(mixRow).CreateCell(colIndex + 1).CellStyle = mStyle;
                                //sheet1.GetRow(mixRow).GetCell(colIndex).SetCellValue(double.Parse(ml.Mix) / 100);

                                //sheet1.GetRow(CRow).CreateCell(colIndex).SetCellValue("");
                                //sheet1.GetRow(CRow).GetCell(colIndex).CellStyle = titleStyle;
                                ////sheet1.GetRow(CRow).CreateCell(colIndex + 1).CellStyle = titleStyle;

                                //sheet1.GetRow(noteRow).CreateCell(colIndex).SetCellValue("QPU");
                                //sheet1.GetRow(noteRow).GetCell(colIndex).CellStyle = style1;
                                ////sheet1.GetRow(noteRow).CreateCell(colIndex + 1).CellStyle = style1;

                                //string colNm1 = sgmNPOIUtils.ConvertColumnIndexToColumnName(colIndex);
                                ////string colNm2 = sgmNPOIUtils.ConvertColumnIndexToColumnName(colIndex + 1);

                                //colIndex = colIndex + 1;

                            }


                            //sheet1.AddMergedRegion(new CellRangeAddress(MYRow, MYRow, yearStart, colIndex - 1));
                            //sheet1.GetRow(MYRow).CreateCell(yearStart).SetCellValue(my.Year);
                            //sheet1.GetRow(MYRow).GetCell(yearStart).CellStyle = titleStyle;
                            //for (int x = yearStart + 1; x < colIndex; x++)
                            //{
                            //    sheet1.GetRow(MYRow).CreateCell(x);
                            //    sheet1.GetRow(MYRow).GetCell(x).CellStyle = titleStyle;
                            //}

                            //yearStart = colIndex;

                        }
                        //sheet1.AddMergedRegion(new CellRangeAddress(CGRow, CGRow, countryStart, yearStart - 1));
                        //sheet1.GetRow(CGRow).CreateCell(countryStart).SetCellValue("出口综合版");
                        //sheet1.GetRow(CGRow).GetCell(countryStart).CellStyle = titleStyle;
                        //for (int y = countryStart + 1; y < yearStart; y++)
                        //{
                        //    sheet1.GetRow(CGRow).CreateCell(y);
                        //    sheet1.GetRow(CGRow).GetCell(y).CellStyle = titleStyle;
                        //}

                        //               countryStart = yearStart;
                        countryStart = colIndex;

                        //========以下是出口正常输出的部分

                        List<CountryGroup> countrygroup = b1.CountryGroup;
                        for (int j = 0; j < countrygroup.Count; j++)
                        {
                            CountryGroup cg = countrygroup[j];

                            List<ModelYear> modelyear = cg.ModelYear;
                            for (int m = 0; m < modelyear.Count; m++)
                            {
                                ModelYear my = modelyear[m];
                                List<ModelLine> modelline = my.ModelLine;
                                YearLCR yl = new YearLCR(my.Year);
                                string fn = "SUM(";
                                string stretchFn = "SUM(";

                                for (int n = 0; n < modelline.Count; n++)
                                {
                                    ModelLine ml = modelline[n];

                                    sheet1.AddMergedRegion(new CellRangeAddress(UMDRow, UMDRow, colIndex, colIndex + 1));
                                    sheet1.GetRow(UMDRow).CreateCell(colIndex).SetCellValue(ml.UMD);
                                    sheet1.GetRow(UMDRow).GetCell(colIndex).CellStyle = titleStyle;
                                    sheet1.GetRow(UMDRow).CreateCell(colIndex + 1).CellStyle = titleStyle;

                                    sheet1.AddMergedRegion(new CellRangeAddress(PRow, PRow, colIndex, colIndex + 1));
                                    sheet1.GetRow(PRow).CreateCell(colIndex).SetCellValue(ml.Package);
                                    sheet1.GetRow(PRow).GetCell(colIndex).CellStyle = titleStyle;
                                    sheet1.GetRow(PRow).CreateCell(colIndex + 1).CellStyle = titleStyle;


                                    sheet1.AddMergedRegion(new CellRangeAddress(mixRow, mixRow, colIndex, colIndex + 1));
                                    sheet1.GetRow(mixRow).CreateCell(colIndex).CellStyle = mStyle;
                                    sheet1.GetRow(mixRow).CreateCell(colIndex + 1).CellStyle = mStyle;
                                    sheet1.GetRow(mixRow).GetCell(colIndex).SetCellValue(double.Parse(ml.Mix) / 100);


                                    sheet1.AddMergedRegion(new CellRangeAddress(CRow, CRow, colIndex, colIndex + 1));
                                    sheet1.GetRow(CRow).CreateCell(colIndex).SetCellValue("");
                                    sheet1.GetRow(CRow).GetCell(colIndex).CellStyle = titleStyle;
                                    sheet1.GetRow(CRow).CreateCell(colIndex + 1).CellStyle = titleStyle;

                                    sheet1.GetRow(noteRow).CreateCell(colIndex).SetCellValue("mix");
                                    mixCols.Add(colIndex);
                                    sheet1.GetRow(noteRow).GetCell(colIndex).CellStyle = style1;
                                    sheet1.GetRow(noteRow).CreateCell(colIndex + 1).SetCellValue("QPU");
                                    QPUCols.Add(colIndex + 1);
                                    sheet1.GetRow(noteRow).GetCell(colIndex + 1).CellStyle = style1;

                                    string colNm1 = sgmNPOIUtils.ConvertColumnIndexToColumnName(colIndex);
                                    string colNm2 = sgmNPOIUtils.ConvertColumnIndexToColumnName(colIndex + 1);

                                    if (cg.Country != "China")
                                    {
                                        ExportModel EM = emList.Find(s => s.Year == my.Year && s.Package == ml.Package);
                                        if (EM != null)
                                        {
                                            string exportFn = EM.Formula;

                                            exportFn = exportFn + ml.PackageLCR + "*" + colNm1 + "{0}*" + colNm2 + "{0}+";

                                            EM.Formula = exportFn;
                                        }
                                    }


                                    fn = fn + ml.PackageLCR + "*" + colNm1 + "{0}*" + colNm2 + "{0}+";
                                    stretchFn = stretchFn + ml.StretchPackageLCR + "*" + colNm1 + "{0}*" + colNm2 + "{0}+";


                                    colIndex = colIndex + 2;

                                }
                                fn = fn.Substring(0, fn.Length - 1);
                                fn = fn + ")";
                                stretchFn = stretchFn.Substring(0, stretchFn.Length - 1);

                                stretchFn = stretchFn + ")";
                                yl.Formula = fn;
                                yl.StretchFormula = stretchFn;
                                yearLCR.Add(yl);

                                sheet1.AddMergedRegion(new CellRangeAddress(MYRow, MYRow, yearStart, colIndex - 1));
                                sheet1.GetRow(MYRow).CreateCell(yearStart).SetCellValue(my.Year);
                                sheet1.GetRow(MYRow).GetCell(yearStart).CellStyle = titleStyle;
                                for (int x = yearStart + 1; x < colIndex; x++)
                                {
                                    sheet1.GetRow(MYRow).CreateCell(x);
                                    sheet1.GetRow(MYRow).GetCell(x).CellStyle = titleStyle;
                                }

                                yearStart = colIndex;

                            }


                            sheet1.AddMergedRegion(new CellRangeAddress(CGRow, CGRow, countryStart, yearStart - 1));
                            sheet1.GetRow(CGRow).CreateCell(countryStart).SetCellValue(cg.Country);
                            sheet1.GetRow(CGRow).GetCell(countryStart).CellStyle = titleStyle;
                            for (int y = countryStart + 1; y < yearStart; y++)
                            {
                                sheet1.GetRow(CGRow).CreateCell(y);
                                sheet1.GetRow(CGRow).GetCell(y).CellStyle = titleStyle;
                            }

                            countryStart = yearStart;
                        }

                        //sheet1.AddMergedRegion(new CellRangeAddress(PGRow, PGRow, programeStart, countryStart - 1));
                        //sheet1.GetRow(PGRow).CreateCell(programeStart).SetCellValue(b1.Programe);
                        //sheet1.GetRow(PGRow).GetCell(programeStart).CellStyle = titleStyle;
                        //for (int z = programeStart + 1; z < countryStart; z++)
                        //{
                        //    sheet1.GetRow(PGRow).CreateCell(z);
                        //    sheet1.GetRow(PGRow).GetCell(z).CellStyle = titleStyle;
                        //}


                        //调整输出位置
                        int exportStart = colIndex;
                        for (int m = 0; m < myList.Count; m++)
                        {
                            ModelYear my = myList[m];
                            List<ModelLine> modelline = my.ModelLine;

                            for (int n = 0; n < modelline.Count; n++)
                            {
                                ModelLine ml = modelline[n];

                                int colWidth = sheet1.GetColumnWidth(colIndex);
                                sheet1.SetColumnWidth(colIndex, colWidth * 2);

                                sheet1.GetRow(UMDRow).CreateCell(colIndex).SetCellValue(ml.UMD);
                                sheet1.GetRow(UMDRow).GetCell(colIndex).CellStyle = titleStyle;
                                //sheet1.GetRow(UMDRow).CreateCell(colIndex + 1).CellStyle = titleStyle;

                                sheet1.GetRow(PRow).CreateCell(colIndex).SetCellValue(ml.Package);
                                sheet1.GetRow(PRow).GetCell(colIndex).CellStyle = titleStyle;
                                //sheet1.GetRow(PRow).CreateCell(colIndex + 1).CellStyle = titleStyle;

                                sheet1.GetRow(mixRow).CreateCell(colIndex).CellStyle = mStyle;
                                //sheet1.GetRow(mixRow).CreateCell(colIndex + 1).CellStyle = mStyle;
                                sheet1.GetRow(mixRow).GetCell(colIndex).SetCellValue(double.Parse(ml.Mix) / 100);

                                sheet1.GetRow(CRow).CreateCell(colIndex).SetCellValue("");
                                sheet1.GetRow(CRow).GetCell(colIndex).CellStyle = titleStyle;
                                //sheet1.GetRow(CRow).CreateCell(colIndex + 1).CellStyle = titleStyle;

                                sheet1.GetRow(noteRow).CreateCell(colIndex).SetCellValue("QPU");
                                sheet1.GetRow(noteRow).GetCell(colIndex).CellStyle = style1;
                                //sheet1.GetRow(noteRow).CreateCell(colIndex + 1).CellStyle = style1;

                                string colNm1 = sgmNPOIUtils.ConvertColumnIndexToColumnName(colIndex);
                                //string colNm2 = sgmNPOIUtils.ConvertColumnIndexToColumnName(colIndex + 1);

                                colIndex = colIndex + 1;

                            }


                            sheet1.AddMergedRegion(new CellRangeAddress(MYRow, MYRow, yearStart, colIndex - 1));
                            sheet1.GetRow(MYRow).CreateCell(yearStart).SetCellValue(my.Year);
                            sheet1.GetRow(MYRow).GetCell(yearStart).CellStyle = titleStyle;
                            for (int x = yearStart + 1; x < colIndex; x++)
                            {
                                sheet1.GetRow(MYRow).CreateCell(x);
                                sheet1.GetRow(MYRow).GetCell(x).CellStyle = titleStyle;
                            }

                            yearStart = colIndex;

                        }
                        sheet1.AddMergedRegion(new CellRangeAddress(CGRow, CGRow, countryStart, yearStart - 1));
                        sheet1.GetRow(CGRow).CreateCell(countryStart).SetCellValue("出口综合版");
                        sheet1.GetRow(CGRow).GetCell(countryStart).CellStyle = titleStyle;
                        for (int y = countryStart + 1; y < yearStart; y++)
                        {
                            sheet1.GetRow(CGRow).CreateCell(y);
                            sheet1.GetRow(CGRow).GetCell(y).CellStyle = titleStyle;
                        }

                        //调整结束

                        for (int j = 0; j < emList.Count; j++)
                        {

                            exportCols.Add(exportStart);  //记录综合部分的列
                            ExportModel em = emList[j];
                            string year = em.Year;
                            ModelYear my = myList.Find(s => s.Year == year);
                            List<ModelLine> mlList = my.ModelLine;
                            ModelLine ml = mlList.Find(s => s.Package == em.Package);
                            string pLCR = ml.PackageLCR;
                            string fn = em.Formula;
                            fn = fn.Substring(0, fn.Length - 1);
                            fn = "IF(AND(TRIM(A{0})=\"\",TRIM(B{0})=\"\"),\"\",(" + fn + ")/" + pLCR + ")";


                            for (int f = 0; f < endInputRow; f++)
                            {
                                string yFn = string.Format(fn, inputRow + f);
                                int r = inputRow - 1 + f;
                                sheet1.GetRow(r).CreateCell(exportStart).SetCellFormula(yFn);
                            }

                            sheet1.GetRow(16).GetCell(exportStart).CellStyle = lcrStyle;

                            string colNm = sgmNPOIUtils.ConvertColumnIndexToColumnName(exportStart);
                            fCol = fCol + colNm + ",";

                            exportStart = exportStart + 1;

                        }

                        //调整输出位置
                        sheet1.AddMergedRegion(new CellRangeAddress(PGRow, PGRow, programeStart, countryStart - 1));
                        sheet1.GetRow(PGRow).CreateCell(programeStart).SetCellValue(b1.Programe);
                        sheet1.GetRow(PGRow).GetCell(programeStart).CellStyle = titleStyle;
                        for (int z = programeStart + 1; z < countryStart; z++)
                        {
                            sheet1.GetRow(PGRow).CreateCell(z);
                            sheet1.GetRow(PGRow).GetCell(z).CellStyle = titleStyle;
                        }

                        //         programeStart = countryStart;
                        programeStart = exportStart;

                    }


                }

                List<YearLCR> newYearLCR = splitList(yearLCR);

                int myColIndex = colIndex;

                ICellStyle myStyle = sgmNPOIUtils.currentStyle("modelYear", hssfworkbook);
                sheet1.GetRow(CRow).CreateCell(colIndex).SetCellValue("Model Year");
                int columnWidth = sheet1.GetColumnWidth(colIndex);
                sheet1.SetColumnWidth(colIndex, columnWidth * 2);
                sheet1.GetRow(CRow).GetCell(colIndex).CellStyle = myStyle;
                sheet1.GetRow(noteRow).CreateCell(colIndex);
                sheet1.GetRow(noteRow).GetCell(colIndex).CellStyle = style1;
                colIndex = colIndex + 1;

                for (int a = 0; a < newYearLCR.Count; a++)
                {
                    sheet1.GetRow(CRow).CreateCell(colIndex).SetCellValue(newYearLCR[a].Year);
                    sheet1.SetColumnWidth(colIndex, columnWidth * 2);
                    sheet1.GetRow(CRow).GetCell(colIndex).CellStyle = titleStyle;
                    sheet1.GetRow(noteRow).CreateCell(colIndex);
                    sheet1.GetRow(noteRow).GetCell(colIndex).CellStyle = style1;
                    //  sheet1.GetRow(16).CreateCell(colIndex).SetCellFormula(newYearLCR[a].Formula);

                    string xFn = newYearLCR[a].Formula;
                    xFn = "IF(AND(TRIM(A{0})=\"\",TRIM(B{0})=\"\"),\"\"," + xFn + ")";
                    for (int f = 0; f < endInputRow; f++)
                    {
                        string yFn = string.Format(xFn, inputRow + f);
                        int r = inputRow - 1 + f;
                        sheet1.GetRow(r).CreateCell(colIndex).SetCellFormula(yFn);
                    }


                    string colNm = sgmNPOIUtils.ConvertColumnIndexToColumnName(colIndex);
                    fCol = fCol + colNm + ",";

                    colIndex = colIndex + 1;
                }

                if (is_stretch(b1Content))
                {
                    for (int b = 0; b < newYearLCR.Count; b++)
                    {
                        sheet1.GetRow(CRow).CreateCell(colIndex).SetCellValue(newYearLCR[b].Year + "\n" + "stretch");
                        sheet1.SetColumnWidth(colIndex, columnWidth * 2);
                        sheet1.GetRow(CRow).GetCell(colIndex).CellStyle = titleStyle;
                        sheet1.GetRow(noteRow).CreateCell(colIndex);
                        sheet1.GetRow(noteRow).GetCell(colIndex).CellStyle = style1;
                        //   sheet1.GetRow(16).CreateCell(colIndex).SetCellFormula(newYearLCR[b].StretchFormula);

                        string xFn = newYearLCR[b].Formula;
                        xFn = "IF(AND(TRIM(A{0})=\"\",TRIM(B{0})=\"\"),\"\"," + xFn + ")";
                        for (int f = 0; f < endInputRow; f++)
                        {
                            string yFn = string.Format(xFn, inputRow + f);
                            int r = inputRow - 1 + f;
                            sheet1.GetRow(r).CreateCell(colIndex).SetCellFormula(yFn);
                        }


                        string colNm = sgmNPOIUtils.ConvertColumnIndexToColumnName(colIndex);
                        fCol = fCol + colNm + ",";

                        colIndex = colIndex + 1;
                    }
                }

                sheet1.AddMergedRegion(new CellRangeAddress(mixRow, mixRow, myColIndex, colIndex - 1));
                sheet1.GetRow(mixRow).CreateCell(myColIndex).SetCellValue("LCR");
                sheet1.GetRow(mixRow).GetCell(myColIndex).CellStyle = titleStyle;
                for (int a = myColIndex + 1; a < colIndex; a++)
                {
                    sheet1.GetRow(mixRow).CreateCell(a);
                    sheet1.GetRow(mixRow).GetCell(a).CellStyle = titleStyle;
                }

                for (int c = inputRow - 1; c < inputRow - 1 + endInputRow; c++)
                {
                    for (int d = startColIndex; d < colIndex; d++)
                    {

                        if (exportCols.Exists(s => s == d))
                        {
                            sheet1.GetRow(c).GetCell(d).CellStyle = lcrStyle;
                            continue;
                        }

                        if (d > myColIndex)
                        {
                            sheet1.GetRow(c).GetCell(d).CellStyle = lcrStyle;
                            continue;
                        }
                        if (d > myColIndex)
                        {
                            sheet1.GetRow(c).CreateCell(d).CellStyle = QPUStyle;
                            continue;
                        }
                        if (d == myColIndex)
                        {
                            sheet1.GetRow(c).CreateCell(d).CellStyle = style1;
                            continue;
                        }
                        if (mixCols.Exists(s => s == d))
                        {
                            sheet1.GetRow(c).CreateCell(d).CellStyle = mixStyle;
                            continue;
                        }

                        if (QPUCols.Exists(s => s == d))
                        {
                            sheet1.GetRow(c).CreateCell(d).CellStyle = QPUStyle;
                            continue;
                        }


                        //}

                    }
                }
                sheet1.GetRow(0).GetCell(5).SetCellValue("");
                sheet1.GetRow(0).GetCell(6).SetCellValue("");
                fCol = fCol.Substring(0, fCol.Length - 1);
                fCell.SetCellValue(fCol);
                sheet1.CreateFreezePane(4, 0, 5, 0);
                sheet1.ProtectSheet("password");//设置密码保护
            }
        }


        public List<YearLCR> splitList(List<YearLCR> yearlcr)
        {
            List<YearLCR> tempList = new List<YearLCR>();
            List<YearLCR> newList = new List<YearLCR>();
            for (int i = 0; i < yearlcr.Count; i++)
            {
                YearLCR yl = yearlcr[i];
                string year = yl.Year;
                string formula = yl.Formula;
                string sFormula = yl.StretchFormula;
                year = year.Replace("，", ","); //防止中文逗号
                string[] yearArray = year.Split(',');

                for (int j = 0; j < yearArray.Length; j++)
                {
                    YearLCR newYL = new YearLCR();
                    newYL.Year = yearArray[j];
                    newYL.Formula = formula;
                    newYL.StretchFormula = sFormula;
                    tempList.Add(newYL);


                }
            }
            List<string> yearList = tempList.Select(p => p.Year).Distinct().ToList();  //year去重
            yearList.Sort();

            for (int m = 0; m < yearList.Count; m++)
            {
                string year = yearList[m];
                List<YearLCR> tempYL = tempList.FindAll(s => s.Year == year);
                YearLCR nYL = new YearLCR();
                nYL.Year = year;
                string f = " ";
                string sf = " ";
                for (int n = 0; n < tempYL.Count; n++)
                {
                    f = f + "+" + tempYL[n].Formula;
                    sf = sf + "+" + tempYL[n].StretchFormula;

                }
                f = f.Substring(2, f.Length - 2);
                sf = sf.Substring(2, sf.Length - 2);
                nYL.Formula = f;
                nYL.StretchFormula = sf;
                newList.Add(nYL);
            }


            return newList;

        }
        public List<string> Test()
        {
            List<string> templist = yearLCR.Select(p => p.Year).Distinct().ToList();  //year去重
            templist.Sort();
            return templist;

        }

        public bool is_stretch(List<B1Content> b1content)
        {
            bool f = false;
            for (int i = 0; i < b1content.Count; i++)
            {

                if (b1content[i].Is_Stretch == "1")
                {
                    f = f | true;
                }
            }


            return f;
        }

    }
}

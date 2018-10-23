using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Aras.IOM;

namespace sgm_aras_plugin
{
    public class ModelLine
    {

        public string UMD { get; set; }
        public string Package { get; set; }
        public string Mix { get; set; }
        public string Comments { get; set; }
        public string PackageLCR { get; set; }
        public string StretchPackageLCR { get; set; }

        public string IS_Stretch { get; set; }

        public ModelLine() { }
        public ModelLine(string umd, string package, string mix, string comments, string packagelcr,string stretchpackagelcr)
        {
            this.UMD = umd;
            this.Package = package;
            this.Mix = mix;
            this.Comments = comments;
            this.PackageLCR = packagelcr;
            this.StretchPackageLCR = stretchpackagelcr;
        }
        public ModelLine(string umd, string package, string mix, string comments, string packagelcr, string stretchpackagelcr,string is_stretch)
        {
            this.UMD = umd;
            this.Package = package;
            this.Mix = mix;
            this.Comments = comments;
            this.PackageLCR = packagelcr;
            this.StretchPackageLCR = stretchpackagelcr;
            this.IS_Stretch = is_stretch;
        }
    }

    public class ModelYear
    {
        public string Year { get; set; }
        public List<ModelLine> ModelLine { get; set; }
        public string LCR { get; set; }
        public string StretchLCR { get; set; }

        public string Formula { get; set; }

        public ModelYear() { }

        public ModelYear(string year,string lcr,string stretchlcr,List<ModelLine> modelline)
        {
            this.Year = year;
            this.ModelLine = ModelLine;
            this.LCR = lcr;
            this.StretchLCR = stretchlcr;
        }

        public ModelYear(string year, List<ModelLine> modelline,string formula)
        {
            this.Year = year;
            this.ModelLine = ModelLine;
            this.Formula = formula;
        }
    }

    public class CountryGroup
    {
        public string Country { get; set; }
        public List<ModelYear> ModelYear { get; set; }

        public CountryGroup() { }

        public CountryGroup(string country, List<ModelYear> modelyear)
        {
            this.Country = country;
            this.ModelYear = modelyear;
        }
    }
    public class B1Content
    {
        public string Programe { get; set; }
        public List<CountryGroup> CountryGroup { get; set; }
        public string Is_Stretch { get; set; }
        //public int CountryCount;

        public B1Content() { }
        public B1Content(string programe,List<CountryGroup> countrygroup)
        {
            this.Programe = programe;
            this.CountryGroup = countrygroup;
        }

        public B1Content(string programe, List<CountryGroup> countrygroup,string is_stretch)
        {
            this.Programe = programe;
            this.CountryGroup = countrygroup;
            this.Is_Stretch = is_stretch;
        }

        //public void setCountryCount()
        //{
           
        //}
    }

    public class YearLCR
    {
        public string Year { get; set; }
        public string Formula { get; set; }
        public string StretchFormula { get; set; }

        public YearLCR() { }
        public YearLCR(string year)
        {
            this.Year = year;
        }

        public YearLCR(string year,string formula)
        {
            this.Formula = formula;
            this.Year = year;
        }

        public YearLCR(string year, string formula,string stretchformula)
        {
            this.Formula = formula;
            this.Year = year;
            this.StretchFormula = stretchformula;
        }

    }

    public class ExportModel
    {
        public string UDM { get; set; }
        public string Package  { get; set; }
        public int ColIndex { get; set; }
        public string Formula { get; set; }
        public string Year { get; set; }

        public string packageLCR { get; set; }

        public ExportModel() { }

        public ExportModel(string package,int colIndex)
        {
            this.Package = package;
            this.ColIndex = colIndex;
        }

        public ExportModel(string package, int colIndex,string formula)
        {
            this.Package = package;
            this.ColIndex = colIndex;
            this.Formula = formula;
        }
    }


}

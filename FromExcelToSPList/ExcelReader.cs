using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Excel = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;

namespace FromExcelToSPList
{
    public class ExcelReader
    {
        private Excel.Application xlApp;
        private Excel.Workbook xlWorkbook;
        private Excel._Worksheet xlWorkSheet;
        private Excel.Range xlRange;
        private List<Isik> isikud;
        private List<Noustamine> noustamised;
        private Isik isik;
        private Noustamine noustamine;

        public ExcelReader(string workbookLocation) 
        {
            xlApp = new Excel.Application();
            xlWorkbook = xlApp.Workbooks.Open(workbookLocation);
            isikud = new List<Isik>();
            noustamised = new List<Noustamine>();
        }

        private void LoeNoustamised()
        {
            xlWorkSheet = (Excel._Worksheet)xlWorkbook.Sheets[1];

            //Loetleb kuni viimase reani millest andemid loeb, andmed algavad 4-ndalt realt, lõppevad 3539 realt
            for (int j = 4; j <= 10; j++)
            {
                noustamine = new Noustamine();
                noustamine.pealkiri = "Konsultatsioon";
                noustamine.isik = (string)(xlWorkSheet.Cells[j, 3] as Excel.Range).Value;

                string noustamiskeskus = (string)(xlWorkSheet.Cells[j, 6] as Excel.Range).Value;
                if (noustamiskeskus.Contains(","))
                {
                    int l = noustamiskeskus.IndexOf(",");
                    noustamiskeskus = noustamiskeskus.Substring(0, l).Trim();
                }
                noustamine.noustamiskeskus = noustamiskeskus;

                noustamine.esmakylastus = (string)(xlWorkSheet.Cells[j, 7] as Excel.Range).Value;
                noustamine.algus = (string)Convert.ToString((xlWorkSheet.Cells[j, 8] as Excel.Range).Value);
                noustamine.lopp = (string)Convert.ToString((xlWorkSheet.Cells[j, 9] as Excel.Range).Value);
                noustamine.valdkond = (string)(xlWorkSheet.Cells[j, 14] as Excel.Range).Value;
                noustamine.tapsemKusimus = (string)(xlWorkSheet.Cells[j, 15] as Excel.Range).Value;
                noustamine.kaua_eestis = (string)(xlWorkSheet.Cells[j, 25] as Excel.Range).Value;
                noustamine.kustSaiInfot = (string)(xlWorkSheet.Cells[j, 50] as Excel.Range).Value;
                noustamine.noustaja = (string)(xlWorkSheet.Cells[j, 51] as Excel.Range).Value;
                noustamine.kaib = "Lõppenud";

                //Kohanemise motiveerimines
                string kohanemiseMotiveerimine = (string)(xlWorkSheet.Cells[j, 17] as Excel.Range).Value;
                if (kohanemiseMotiveerimine == "x" || kohanemiseMotiveerimine == "X")
                {
                    kohanemiseMotiveerimine = "Jah";
                }
                else
                {
                    kohanemiseMotiveerimine = "Ei";
                }
                noustamine.kohanemiseMotiveerimine = kohanemiseMotiveerimine;

                //Osalemine NK üritustel
                string osalemineNK = (string)(xlWorkSheet.Cells[j, 49] as Excel.Range).Value;
                if (osalemineNK == "x" || osalemineNK == "X")
                {
                    osalemineNK = "Jah";
                }
                else
                {
                    osalemineNK = "Ei";
                }
                noustamine.osalemineNK = osalemineNK;

                //Tööhõive TAT'ga liitumisel 18-22
                string toohoive = "";
                string toohoiveVaartus;
                for (int k = 18; k <= 22; k++)
                {
                    toohoiveVaartus = (string)(xlWorkSheet.Cells[j, k] as Excel.Range).Value;

                    if (toohoiveVaartus != null)
                        toohoive += "x";
                    else
                        toohoive += "o";
                }
                noustamine.toohoiveTATgaLiitumisel = toohoive;

                //Ebasoodsad olud
                string ebasoodsad = "";
                string ebasoodsadVaartus;
                for (int l = 26; l <= 32; l++)
                {
                    ebasoodsadVaartus = (string)(xlWorkSheet.Cells[j, l] as Excel.Range).Value;

                    if (ebasoodsadVaartus != null)
                        ebasoodsad += "x";
                    else
                        ebasoodsad += "o";
                }
                noustamine.ebasoodsadOlud = ebasoodsad;

                //Olukord peale TAT
                string pealeTAT = "";
                string pealeTATVaartus;
                for (int m = 33; m <= 39; m++)
                {
                    pealeTATVaartus = (string)(xlWorkSheet.Cells[j, m] as Excel.Range).Value;

                    if (pealeTATVaartus != null)
                        pealeTAT += "x";
                    else
                        pealeTAT += "o";
                }
                noustamine.olukordPealeTAT = pealeTAT;

                //Olukord peale x-kuud peale TAT
                string xPealeTAT = "";
                string xPealeTATVaartus;
                for (int n = 40; n <= 48; n++)
                {
                    xPealeTATVaartus = (string)(xlWorkSheet.Cells[j, n] as Excel.Range).Value;

                    if (xPealeTATVaartus != null)
                        xPealeTAT += "x";
                    else
                        xPealeTAT += "o";
                }
                noustamine.olukordXkuudPealeTAT = xPealeTAT;

                noustamised.Add(noustamine);
            }
        }

        private void LoeIsikud()
        {
            xlWorkSheet = (Excel._Worksheet)xlWorkbook.Sheets[1];

                //Loetleb kuni viimase reani millest andeid loed, andmed algavad 4-ndalt realt, , lõppevad 3539 realt
                for (int j = 4; j <= 10; j++)
                {
                    isik = new Isik();
                    isik.isikukood = (string)Convert.ToString((xlWorkSheet.Cells[j, 2] as Excel.Range).Value);
                    isik.nimi = (string)(xlWorkSheet.Cells[j, 3] as Excel.Range).Value;
                    isik.elukoht = (string)(xlWorkSheet.Cells[j, 4] as Excel.Range).Value;
                    isik.epost = (string)(xlWorkSheet.Cells[j, 10] as Excel.Range).Value;
                    isik.telnr = (string)Convert.ToString((xlWorkSheet.Cells[j, 11] as Excel.Range).Value);
                    isik.vanus = (string)(xlWorkSheet.Cells[j, 13] as Excel.Range).Value;

                    //Sugu
                    string sugu = (string)(xlWorkSheet.Cells[j, 12] as Excel.Range).Value;
                    if (sugu.Equals("Mees") || sugu.Equals("Naine"))
                    {
                        isik.sugu = sugu;
                    }
                    else
                    {
                        isik.sugu = "";
                    }

                    //Kodakondsus
                    if ((string)(xlWorkSheet.Cells[j, 5] as Excel.Range).Value != null)
                    {
                        string kodakondsus = (string)(xlWorkSheet.Cells[j, 5] as Excel.Range).Value;
                        if (kodakondsus.Equals("kodakondsuseta", StringComparison.InvariantCultureIgnoreCase) || kodakondsus.Equals("kodakonsuseta", StringComparison.InvariantCultureIgnoreCase))
                            isik.kodakondsus = "Kodakonduseta";
                        else
                            isik.kodakondsus = kodakondsus;
                    }
                    else
                        isik.kodakondsus = "";

                    //SIM kohanemisprogramm
                    if ((string)(xlWorkSheet.Cells[j, 16] as Excel.Range).Value != null)
                    {
                        string sim_kohanemisprogramm = (string)(xlWorkSheet.Cells[j, 16] as Excel.Range).Value;
                        if (sim_kohanemisprogramm.Equals("x", StringComparison.InvariantCultureIgnoreCase))
                            isik.simKohanemisprogramm = "Jah";
                        else
                            isik.simKohanemisprogramm = "Ei";
                    }
                    else
                        isik.simKohanemisprogramm = "Ei";

                    //Haridus
                    if ((string)(xlWorkSheet.Cells[j, 23] as Excel.Range).Value != null)
                    {
                        string haridus = (string)(xlWorkSheet.Cells[j, 23] as Excel.Range).Value;
                        if (haridus == "")
                            isik.haridus = "Ei ole teada";
                        else
                            isik.haridus = haridus;
                    }
                    else
                        isik.haridus = "Ei ole teada";

                    //Soov õppida eesti keelt
                    if ((string)(xlWorkSheet.Cells[j, 24] as Excel.Range).Value != null)
                    {
                        string soov_eesti_keelt = (string)(xlWorkSheet.Cells[j, 24] as Excel.Range).Value;
                        if (soov_eesti_keelt.Equals("x", StringComparison.InvariantCultureIgnoreCase))
                            isik.soovEestiKeelt = "Jah";
                        else
                            isik.soovEestiKeelt = "Ei";
                    }
                    else
                        isik.soovEestiKeelt = "Ei";

                    isikud.Add(isik);
                }
        }

        public List<Isik> SaaIsikud() 
        {
            LoeIsikud();
            return isikud;
        }

        public List<Noustamine> SaaNoustamised()
        {
            LoeNoustamised();
            return noustamised;
        }
    }
}
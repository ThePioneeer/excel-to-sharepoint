using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using OfficeOpenXml;
using System.IO;

namespace FromExcelToSPList
{
    class ExcelPackageReader
    {
        private List<Isik> isikud;
        private List<Noustamine> noustamised;
        private Isik isik;
        private Noustamine noustamine;
        private ExcelPackage xlPackage;
        private ExcelWorksheet xlWorkSheet;
        private int end;

        public ExcelPackageReader(string workbookLocation, int end)
        {
            FileInfo newFile = new FileInfo(workbookLocation);
            isikud = new List<Isik>();
            noustamised = new List<Noustamine>();
            xlPackage = new ExcelPackage(newFile);
            xlWorkSheet = xlPackage.Workbook.Worksheets[1];
            this.end = end;
        }

        private void LoeNoustamised()
        {            
            //Loetleb kuni viimase reani millest andemid loeb, andmed algavad 1-ndalt realt, lõppevad x realt
            for (int j = 1; j <= end; j++)
            {
                noustamine = new Noustamine();
                noustamine.pealkiri = "Konsultatsioon";
                noustamine.isik = xlWorkSheet.Cell(j, 3).Value;

                string noustamiskeskus = xlWorkSheet.Cell(j, 6).Value;
                if (noustamiskeskus.Contains(","))
                {
                    int l = noustamiskeskus.IndexOf(",");
                    noustamiskeskus = noustamiskeskus.Substring(0, l).Trim();
                }
                noustamine.noustamiskeskus = noustamiskeskus;

                noustamine.esmakylastus = xlWorkSheet.Cell(j, 7).Value;
                noustamine.algus = xlWorkSheet.Cell(j, 8).Value;
                noustamine.lopp = xlWorkSheet.Cell(j, 9).Value;
                noustamine.valdkond = xlWorkSheet.Cell(j, 14).Value;
                noustamine.tapsemKusimus = xlWorkSheet.Cell(j, 15).Value;
                noustamine.kaua_eestis = xlWorkSheet.Cell(j, 25).Value;
                noustamine.kustSaiInfot = xlWorkSheet.Cell(j, 50).Value;
                noustamine.noustaja = xlWorkSheet.Cell(j, 51).Value;
                noustamine.kaib = "Lõppenud";
                noustamine.rahastaja = "KUM/ESF";

                //Kohanemise motiveerimines
                string kohanemiseMotiveerimine = xlWorkSheet.Cell(j, 17).Value;
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
                string osalemineNK = xlWorkSheet.Cell(j, 49).Value;
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
                    toohoiveVaartus = xlWorkSheet.Cell(j, k).Value;

                    if (toohoiveVaartus != "")
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
                    ebasoodsadVaartus = xlWorkSheet.Cell(j, l).Value;

                    if (ebasoodsadVaartus != "")
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
                    pealeTATVaartus = xlWorkSheet.Cell(j, m).Value;

                    if (pealeTATVaartus != "")
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
                    xPealeTATVaartus = xlWorkSheet.Cell(j, n).Value;

                    if (xPealeTATVaartus != "")
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
            //Loetleb kuni viimase reani millest andeid loed, andmed algavad 4-ndalt realt, , lõppevad 3539 realt
            for (int j = 1; j <= end; j++)
            {
                isik = new Isik();
                isik.isikukood = Convert.ToString(xlWorkSheet.Cell(j, 2).Value);
                isik.nimi = xlWorkSheet.Cell(j, 3).Value;
                isik.elukoht = xlWorkSheet.Cell(j, 4).Value;
                isik.epost = xlWorkSheet.Cell(j, 10).Value;
                isik.telnr = Convert.ToString(xlWorkSheet.Cell(j, 11).Value);
                isik.vanus = xlWorkSheet.Cell(j, 13).Value;

                //Sugu
                string sugu = xlWorkSheet.Cell(j, 12).Value;
                if (sugu.Equals("Mees") || sugu.Equals("Naine"))
                {
                    isik.sugu = sugu;
                }
                else
                {
                    isik.sugu = "";
                }

                //Kodakondsus
                if (xlWorkSheet.Cell(j, 5).Value != "")
                {
                    string kodakondsus = xlWorkSheet.Cell(j, 5).Value;
                    if (kodakondsus.Equals("kodakondsuseta", StringComparison.InvariantCultureIgnoreCase) || kodakondsus.Equals("kodakonsuseta", StringComparison.InvariantCultureIgnoreCase))
                        isik.kodakondsus = "Kodakonduseta";
                    else
                        isik.kodakondsus = kodakondsus;
                }
                else
                    isik.kodakondsus = "";

                //SIM kohanemisprogramm
                if (xlWorkSheet.Cell(j, 16).Value != "")
                {
                    string sim_kohanemisprogramm = xlWorkSheet.Cell(j, 16).Value;
                    if (sim_kohanemisprogramm.Equals("x", StringComparison.InvariantCultureIgnoreCase))
                        isik.simKohanemisprogramm = "Jah";
                    else
                        isik.simKohanemisprogramm = "Ei";
                }
                else
                    isik.simKohanemisprogramm = "Ei";

                //Haridus
                if (xlWorkSheet.Cell(j, 23).Value != "")
                {
                    string haridus = xlWorkSheet.Cell(j, 23).Value;
                    if (haridus == "")
                        isik.haridus = "Ei ole teada";
                    else
                        isik.haridus = haridus;
                }
                else
                    isik.haridus = "Ei ole teada";

                //Soov õppida eesti keelt
                if (xlWorkSheet.Cell(j, 24).Value != "")
                {
                    string soov_eesti_keelt = xlWorkSheet.Cell(j, 24).Value;
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

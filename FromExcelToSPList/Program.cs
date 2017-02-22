using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.SharePoint;
using System.Web;
using System.Net;
using Microsoft.SharePoint.Client;

namespace FromExcelToSPList
{
    class Program
    {
        static void Main(string[] args)
        {
            
            string fileLocation1 = @"";
            string fileLocation = @"";
            
            ExcelPackageReader reader = new ExcelPackageReader(fileLocation, 994);
            
            //ExcelReader reader = new ExcelReader(fileLocation);
            Console.WriteLine("Loen andmeid");
            DateTime start = DateTime.Now;
            List<Isik> listIsikud = reader.SaaIsikud();
            List<Noustamine> listNoustamised = reader.SaaNoustamised();
            DateTime end = DateTime.Now;
            Console.WriteLine("Andmed loetud," + (end - start));

            /
            ExcelPackageReader reader1 = new ExcelPackageReader(fileLocation1, 324);

            //ExcelReader reader = new ExcelReader(fileLocation);
            Console.WriteLine("Loen andmeid");
            start = DateTime.Now;
            List<Isik> listIsikud1 = reader1.SaaIsikud();
            List<Noustamine> listNoustamised1 = reader1.SaaNoustamised();
            end = DateTime.Now;
            Console.WriteLine("Andmed loetud," + (end - start));

            string siteUrl = "";
            //string siteUrl = "https://moss2010-v2/";

            SPSite site = new SPSite(siteUrl);           
            SPWeb web = site.OpenWeb();
            Console.WriteLine("Serveriga ühendatud");

            LisaIsikud(web, listIsikud);
            LisaIsikud(web, listIsikud1);
            Console.WriteLine("Isikud lisatud");
            LisaNoustamised(web, listNoustamised);
            LisaNoustamised(web, listNoustamised1);
            Console.WriteLine("Nõustamised lisatud");

            Console.ReadLine();
        }

        private static void LisaIsikud(SPWeb web, List<Isik> list)
        { 
            SPListItemCollection listItems;
            bool hasDuplicate;
            int itemCount;
 
            //Loetleb iga inimese läbi
            for (int i = 1; i <= list.Count; i++)
			{
                //Kui on sama nimega olemas siis ei lisata listi
                hasDuplicate = false;
                listItems = web.Lists["Isikud"].Items;
                itemCount = listItems.Count;
                //Loetleb iga listi objekti läbi ja kontrollib kas eelnevalt sama nimega listi üksust pole. 
                //Kui isiku nimi on "Pole teada", siis lisatakse kasutaja ikkagist.
			    for (int j = 1; j <= itemCount; j++)
                {
                    SPListItem item = listItems[j-1];
                    if (list[i-1].nimi.Equals(item["Title"].ToString()) && !list[i-1].nimi.Equals("Pole teada"))
                    {
                        hasDuplicate = true;
                        break;
                    }
                }

                if(!hasDuplicate)
                {
                    SPListItem newItem = listItems.Add();                    
                    newItem["Title"] = list[i-1].nimi;
                    newItem["Isikukood"] = list[i-1].isikukood;
                    newItem["Elukoht_x002c__x0020_linn"] = list[i - 1].elukoht;
                    newItem["Kodakondsus"] = list[i - 1].kodakondsus;
                    newItem["E_x002d_post"] = list[i - 1].epost;
                    newItem["Tel_x0020_nr"] = list[i - 1].telnr;
                    newItem["Sugu"] = list[i - 1].sugu;
                    newItem["Vanus"] = list[i - 1].vanus;
                    newItem["SIM_x0020_kohanemisprogramm"] = list[i - 1].simKohanemisprogramm;
                    newItem["Haridus"] = list[i - 1].haridus;
                    newItem["Soov_x0020__x00f5_ppida_x0020_ee"] = list[i - 1].soovEestiKeelt;
                    newItem.Update();
                }
			}        
        }
        
        private static void LisaNoustamised(SPWeb web, List<Noustamine> list)
        {
            SPListItemCollection listItems = web.Lists["Nõustamised"].Items;
            for (int i = 1; i <= list.Count; i++)
            {
                int id = GetLookupID(web, "Isikud", "Title", list[i-1].isik);
                SPListItem newItem = listItems.Add();

                newItem["Title"] = "Konsultatsioon";
                newItem["Isik"] = new SPFieldLookupValue(id.ToString());
                newItem["N_x00f5_ustamiskeskus"] = list[i - 1].noustamiskeskus;
                newItem["Esmak_x00fc_lastus"] = list[i - 1].esmakylastus;
                try
                {
                    int date = Convert.ToInt32(list[i - 1].algus);
                    DateTime algus = DateTime.FromOADate(date);
                    newItem["Algus"] = algus;
                }
                catch (Exception)
                {
                    
                }
                try
                {
                    int date = Convert.ToInt32(list[i - 1].lopp);
                    DateTime lopp = DateTime.FromOADate(date);
                    newItem["L_x00f5_pp"] = lopp;
                }
                catch (Exception)
                {
                    
                }
                newItem["T_x00e4_psem_x0020_k_x00fc_simus"] = list[i - 1].tapsemKusimus;
                newItem["Kui_x0020_kaua_x0020_on_x0020_Ee"] = list[i - 1].kaua_eestis;
                newItem["Rahastaja"] = list[i - 1].rahastaja;
                newItem["Osalemine_x0020_NK_x0020__x00fc_"] = list[i - 1].osalemineNK;
                newItem["Kohanemise_x0020_motiveerimine"] = list[i - 1].kohanemiseMotiveerimine;
                newItem["Staatus"] = list[i - 1].kaib;
                
                //Valdkond
                var valdkond = new SPFieldMultiChoiceValue();                
                if (list[i - 1].valdkond != null)
                {
                    var valdkondList = list[i - 1].valdkond.Split(',').ToList<string>();
                    foreach (string v in valdkondList)
                    {
                        valdkond.Add(v);
                    }
                }
                newItem["Valdkond"] = valdkond;

                //Kust sai infot
                var infot = new SPFieldMultiChoiceValue();
                if (list[i - 1].kustSaiInfot != null)
                {
                    var infotList = list[i - 1].kustSaiInfot.Split(',').ToList<string>();
                    foreach (string v in infotList)
                    {
                        infot.Add(v);
                    }
                }
                newItem["Kust_x0020_sai_x0020_infot_x0020"] = infot;

                //Tööhõive TAT liitumusel
                var toohoive = new SPFieldMultiChoiceValue();
                string toohoiveStr = list[i - 1].toohoiveTATgaLiitumisel;
                if (toohoiveStr != null)
                {
                    if (toohoiveStr[0] == 'x') toohoive.Add("Töötav, sh FIE");
                    if (toohoiveStr[1] == 'x') toohoive.Add("Töötu, sh pikaajaline töötu");
                    if (toohoiveStr[2] == 'x') toohoive.Add("Töötu - pikaajaline töötu");
                    if (toohoiveStr[3] == 'x') toohoive.Add("Mitteaktiivne");
                    if (toohoiveStr[4] == 'x') toohoive.Add("Mitteaktiivne, kes ei ole hariduses ega koolitusel");
                }
                newItem["T_x00f6__x00f6_h_x00f5_ive_x0020"] = toohoive;

                //Ebasoodsad olud
                var eabsoodsad = new SPFieldMultiChoiceValue();
                string eabsoodsadStr = list[i - 1].ebasoodsadOlud;
                if (eabsoodsadStr != null)
                {
                    if (eabsoodsadStr[0] == 'x') eabsoodsad.Add("Osaleja, kes elab tööta leibkonnas");
                    if (eabsoodsadStr[1] == 'x') eabsoodsad.Add("Osaleja, kes elab ülalpeetavate lastega tööta leibkonnas");
                    if (eabsoodsadStr[2] == 'x') eabsoodsad.Add("Osaleja, kes elab ühe täiskasvanuga leibkonnas koos ülalpeetavate lastega");
                    if (eabsoodsadStr[3] == 'x') eabsoodsad.Add("Osaleja, kelle emakeel ei ole eesti keel, sh mittekodanikust alaline elanik, teise riigi taustaga isik või rahvusvähemuste hulka kuuluv kodanik (sh marginaliseeritud kogukonnad nagu romad)");
                    if (eabsoodsadStr[4] == 'x') eabsoodsad.Add("Puudega inimene");
                    if (eabsoodsadStr[5] == 'x') eabsoodsad.Add("Muus ebasoodsas olukorras olev inimene");
                    if (eabsoodsadStr[6] == 'x') eabsoodsad.Add("Kodutu või eluasemeturult tõrjutud osaleja");
                }
                newItem["Ebasoodsad_x0020_olud"] = eabsoodsad;

                //Olukord peale TAT
                var pealeTAT = new SPFieldMultiChoiceValue();
                string pealeTATStr = list[i - 1].olukordPealeTAT;
                if (pealeTATStr != null)
                {
                    if (pealeTATStr[0] == 'x') pealeTAT.Add("Mitteaktiivne, kes on asunud tööd otsima");
                    if (pealeTATStr[1] == 'x') pealeTAT.Add("Osaleja, kes on hariduses või koolitusel");
                    if (pealeTATStr[2] == 'x') pealeTAT.Add("Osaleja, kes sai kutsekvalifikatsiooni");
                    if (pealeTATStr[3] == 'x') pealeTAT.Add("Osaleja, kes on hõives, sh FIE");
                    if (pealeTATStr[4] == 'x') pealeTAT.Add("Osaleja, kes jätkab tööl");
                    if (pealeTATStr[5] == 'x') pealeTAT.Add("Ebasoodsas olukorras olev osaleja, kes on asunud tööd otsima, haridust või koolitust saama, omandanud kutsekvalifikatsiooni või liikunud hõivesse, sh FIEna");
                    if (pealeTATStr[6] == 'x') pealeTAT.Add("Pole teada");
                }
                newItem["Olukord_x0020_vahetult_x0020_pea"] = pealeTAT;

                //Olukord x kuud peale TAT
                var xPealeTAT = new SPFieldMultiChoiceValue();
                string xPealeTATStr = list[i - 1].olukordXkuudPealeTAT;
                if (xPealeTATStr != null)
                {
                    if (xPealeTATStr[0] == 'x') xPealeTAT.Add("Osaleja, kes 6 kuud pärast lapsehoiu ja/või puuetega laste tugiteenuse saamise algust on  tööturul");
                    if (xPealeTATStr[1] == 'x') xPealeTAT.Add("Osaleja, kes 6 kuud pärast hoolekandeteenuse saamise algust on tööturul");
                    if (xPealeTATStr[2] == 'x') xPealeTAT.Add("Osaleja, kes 1 kuu pärast vanglast vabanenutele suunatud tugiteenuse saamist on tööturul");
                    if (xPealeTATStr[3] == 'x') xPealeTAT.Add("Osaleja, kelle hulgas on alkoholi liigtarvitamise riskitase vähenenud 6 kuud pärast alkoholi tarvitamise vähendamisele suunatud teenuste osutamise algust vähenenud");
                    if (xPealeTATStr[4] == 'x') xPealeTAT.Add("Töötavad inimesed, kelle töövõimet hinnati osaliseks ning kes on säilitanud oma töökoha 12. kuu möödumisel peale hindamist");
                    if (xPealeTATStr[5] == 'x') xPealeTAT.Add("Mittetöötavad inimesed, kelle töövõimet hinnati osaliseks ning kes on liikunud hõivesse 12. kuu möödumisel peale hindamist");
                    if (xPealeTATStr[6] == 'x') xPealeTAT.Add("Osaleja, kes  on asunud kuue kuu jooksul pärast programmist lahkumist tööle, sh füüsilisest isikust ettevõtjana");
                    if (xPealeTATStr[7] == 'x') xPealeTAT.Add("Osaleja, kes sai aasta jooksul ESFist  toetatud hoolekandeteenuseid ning kelle toimetulek seeläbi paranes või kelle puhul välditi ööpäevaringsele institutsionaalsele teenusele suundumist");
                    if (xPealeTATStr[8] == 'x') xPealeTAT.Add("Pole teada");
                }
                newItem["Olukord_x0020_x_x002d_kuud_x0020"] = xPealeTAT;                
                newItem["N_x00f5_ustaja"] = ConvertLoginName(list[i-1].noustaja, web); 
  
                newItem.Update();
            }
        }

        private static SPFieldUserValue ConvertLoginName(string username, SPWeb web)
        {
            string userid;
            switch (username)
            {
                case "Ion Braga":
                    userid = "71";
                    break;
                case "Irina Rakova":
                    userid = "81";
                    break;
                case "Kätlin Kõverik":
                    userid = "73";
                    break;
                case "Natalja Vovdenko":
                    userid = "80";
                    break;
                case "Anna Kuznetsova":
                    userid = "82";
                    break;
                default:
                    return null;
            }

            SPFieldUserValue uservalue = new SPFieldUserValue(web, userid);
            return uservalue;
        }
        
        private static int GetLookupID(SPWeb web, string listName, string lookupColumnName, string lookupValue)
        {
            try
            {
                SPList list = web.Lists[listName];

                string strCAMLQuery = "<Where><Eq><FieldRef Name='" + lookupColumnName + "' /><Value Type='Text'>" + System.Web.HttpUtility.HtmlEncode(lookupValue) + "</Value></Eq></Where>";

                SPQuery query = new SPQuery();
                query.Query = strCAMLQuery;
                query.ViewFields = string.Concat(
                            "<FieldRef Name='ID' />",
                            "<FieldRef Name='" + lookupColumnName + "' />");

                SPListItemCollection items = list.GetItems(query);

                if (items.Count > 0)
                { 
                    int id = items[0].ID;
                    return id;
                }
                else
                {
                    return 0;
                }

            }
            catch (Exception)
            {
                return 0;
            }            
        }
    }
}
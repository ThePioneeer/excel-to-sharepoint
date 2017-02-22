using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace FromExcelToSPList
{
    public class Noustamine
    {
        public string pealkiri { get; set; }
        public string isik { get; set; }
        public string noustamiskeskus { get; set; }
        public string esmakylastus { get; set; }
        public string algus { get; set; }
        public string lopp { get; set; }
        public string valdkond { get; set; }
        public string tapsemKusimus { get; set; }
        public string toohoiveTATgaLiitumisel { get; set; } //xoxox kus x on jah o on ei
        public string kaua_eestis { get; set; }
        public string ebasoodsadOlud { get; set; } //xoxox kus x on jah o on ei
        public string olukordPealeTAT { get; set; } //xoxox kus x on jah o on ei
        public string kohanemiseMotiveerimine { get; set; }
        public string olukordXkuudPealeTAT { get; set; } //xoxox kus x on jah o on ei
        public string osalemineNK { get; set; }
        public string kustSaiInfot { get; set; }
        public string rahastaja { get; set; }
        public string noustaja { get; set; }
        public string kaib { get; set; }
    }
}

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ONE012
{
    public struct One012Data
    {
        public string PracKod { set; get; }     //Akronim                  DOH DOH_DodId -> DOD DOD_DodId  DOD_PraId-> PRA_PraId 
        public string DOH_Nazwa { set; get; }//Nazwa Składnika
        public DateTime DOH_WaznyOd { set; get; }

        public DateTime DOH_WaznyDo { set; get; }

        public decimal DOH_Wartosc1 { set; get; }
        //public string ListaPlac { set; get; }    //symbol tyup skladnika

    }
}

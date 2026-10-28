using System;
using System.IO;
using System.Windows.Forms;
using System.Runtime.InteropServices;
using CDNBase;
//using OfficeOpenXml;
using System.Linq;
//using OfficeOpenXml.Table;
using ExcelDataReader;
using System.Threading;

using System.Collections.Generic;
using System.Data;
//using Excel = Microsoft.Office.Interop.Excel; dad
//using GemBox.Spreadsheet;


namespace ONE012
{
    public partial class ONE012WIN : Form
    {

        public static string Operator = "DARIUSZ"; //"BARTEK";
        public static string Haslo = "NE!w0rk12";       // operator nie ma hasła
        public static string Firma = "RAINBOW TOURS";//"RAINBOW TOURS";//Program.args1; //"ERP_TOMMA_KADRY_TEST";     // nazwa firmy
        protected static CDNBase.Application Application = null;
        protected static ILogin Login = null;
        public ONE012WIN()
        {
            //MessageBox.Show(Firma);
            InitializeComponent();
            //MessageBox.Show("Test2");

        }

        #region funkcje optimy

        static protected void LogowanieAutomatyczne()//do walniecia logowanie 
        {



            object[] hPar = new object[] {
                         0,  0,   0,  0,  0,   0,  0,    0,   0,   0,   0,   0,   0,   0,  1,   0,  0 };    // do jakich modułów się logujemy
                                                                                                            /* Kolejno: KP, KH, KHP, ST, FA, MAG, PK, PKXL, CRM, ANL, DET, BIU, SRW, ODB, KB, KBP, HAP
                                                                                                             */


            // katalog, gdzie jest zainstalowana Optima (bez ustawienia tej zmiennej nie zadziała, chyba że program odpalimy z katalogu O!)
            System.Environment.CurrentDirectory = @"C:\Program Files (x86)\Comarch ERP Optima";
            // tworzymy nowy obiekt apliakcji
            Application = new CDNBase.Application();
            // blokujemy
            Application.LockApp(513, 5000, null, null, null, null);
            // logujemy się do podanej Firmy, na danego operatora, do podanych modułów
            Login = Application.Login(vUser: Operator, vChk: Haslo, vFirm: Firma, _ModulKP: hPar[0], _ModulKH: hPar[1], _ModulKHP: hPar[2], _ModulST: hPar[3], _ModulFA: hPar[4], _ModulMAG: hPar[5], _ModulPK: hPar[6], _ModulPKXL: hPar[7], _ModulCRM: hPar[8], _ModulANL: hPar[9], _ModulDET: hPar[10], _ModulBIU: hPar[11], _ModulSRW: hPar[12], _ModulOBD: hPar[13], _ModulKB: hPar[14], _ModulKBP: hPar[15], _ModulHAP: hPar[16]);

            //problem jest taki, że nei ma w tej metodzie dostępu do KBP do którego chc się zalogować


            // tu jesteśmy zalogowani do O!
            //MessageBox.Show("Jesteśmy zalogowani do O!");
        }

        protected static void Wylogowanie()
        {
            // niszczymy Login
            Login = null;
            // odblokowanie (wylogowanie) O!
            Application.UnlockApp();
            // niszczymy obiekt Aplikacji
            Application = null;
        }
        #endregion

        private void Button1_Click(object sender, EventArgs e)
        {
            #region excel5
            Cursor.Current = Cursors.WaitCursor;
            if(tB1.Text == "nie wybrano pliku")
            {
                MessageBox.Show("Nie wybrano pliku arkusza.", "Komunikat",
                MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            string filePath = tB1.Text;
            DataTable dataTable;
            using (var stream = File.Open(filePath, FileMode.Open, FileAccess.Read))
            {
                IExcelDataReader reader;
                reader = ExcelDataReader.ExcelReaderFactory.CreateReader(stream);

                // reader.IsFirstRowAsColumnNames
                var conf = new ExcelDataSetConfiguration
                {
                    ConfigureDataTable = _ => new ExcelDataTableConfiguration
                    {
                        //UseHeaderRow = false
                    }
                };

                var dataSet = reader.AsDataSet(conf);

                //Now you can get data from each sheet by its index or its "name"
                dataTable = dataSet.Tables[0];
                //dataTable.Rows;
                //...
            }

            //Przypisywanie danych z datasetu do wygodnej listy pracowników.
            List<One012Data> listaPrac = new List<One012Data>();
            if (dataTable.Rows.Count == 0)
            {
                MessageBox.Show("W arkuszu brak pracownikóW.", "Komunikat",
                MessageBoxButtons.OK, MessageBoxIcon.Error);//błąd
                return;
            }

            


            for (int i = 1; i < dataTable.Rows.Count; i++)
            {
                for (int k = 5; k < dataTable.Columns.Count; k++)
                {
                    
                    if (Convert.ToString(dataTable.Rows[0][k]).Contains("#") && !String.IsNullOrEmpty(dataTable.Rows[i][k].ToString().Trim()) && !dataTable.Rows[i][k].ToString().Contains("-") && !String.IsNullOrEmpty(dataTable.Rows[i][4].ToString()))
                    {
                        //MessageBox.Show("Dodaję pole dla akronimu:  "+ Convert.ToString(dataTable.Rows[i][4]), "Komunikat",
                //MessageBoxButtons.OK, MessageBoxIcon.Error);//błąd
                  //      MessageBox.Show("Data od: " + dataTable.Rows[i][0].ToString());
                    //    MessageBox.Show("Data do: " + dataTable.Rows[i][1].ToString());

                        try
                        {
                            listaPrac.Add(new One012Data()
                            {
                                PracKod = Convert.ToString(dataTable.Rows[i][4]),
                                DOH_Nazwa = Convert.ToString(dataTable.Rows[0][k]).Substring(1),
                                DOH_WaznyOd = Convert.ToDateTime(dataTable.Rows[i][0]),
                                DOH_WaznyDo = Convert.ToDateTime(dataTable.Rows[i][1]),
                                DOH_Wartosc1 = Convert.ToDecimal(dataTable.Rows[i][k])
                                //ListaPlac = Convert.ToString(dataTable.Rows[i][5])
                            }); ;
                        }
                        catch (Exception exp)
                        {
                            MessageBox.Show("Niepoprawny format zmiennych podanych w pliku arkusza.\n" + exp.Message+"\nDla wiersza: ["+(i+1)+"]["+(k+1)+"]\nTen wiersz zostanie pominięty przy imporcie", "Komunikat",
                            MessageBoxButtons.OK, MessageBoxIcon.Warning);//blad
                            
                        }
                    }
                }


            }
            Cursor.Current = Cursors.Default;
            #endregion



            #region dzialanie na optimie
            #region logowanie
            //Logowanie do O! bez wyświetlania okienka logowania
            Cursor.Current = Cursors.WaitCursor;
            try
            {
                LogowanieAutomatyczne();
            }
            catch (COMException eo)
            {
                if (eo.InnerException != null)
                    MessageBox.Show("Błąd logowania: " + eo.InnerException.Message, "Komunikat",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
                else
                    MessageBox.Show("Błąd logowania: " + eo.Message, "Komunikat",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            catch (Exception ei)
            {
                if (ei.InnerException != null)
                    MessageBox.Show("Błąd logowania: " + ei.InnerException.Message, "Komunikat",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
                else
                    MessageBox.Show("Błąd logowania: " + ei.Message, "Komunikat",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            #endregion
            CDNBase.AdoSession Sesja = Login.CreateSession();


            //Dodawanie rejestru/aktualizacja doh
            foreach (One012Data prac in listaPrac)
            {
                CDNPrac.IPracownik rPracownik;
                //Sprawdzam pracownika
                try
                {
                    rPracownik = (CDNPrac.IPracownik)Sesja.CreateObject("CDN.Pracownicy", "PRA_Kod = '" + prac.PracKod + "'");
                }
                catch (Exception)
                {
                    if (MessageBox.Show("Nie znaleziono pracownika o akronimie: " + prac.PracKod + ".\nCzy kontynuować proces dodawanie z pominięciem tego przypadku?", "Komunikat", MessageBoxButtons.YesNo, MessageBoxIcon.Error) == DialogResult.Yes)
                    {
                        // user clicked yes
                        continue;
                    }
                    else
                    {
                        // user clicked no
                        Wylogowanie();
                        return;
                    }
                }

                ICollection rDodatki = rPracownik.Dodatki;
                //Dobieranie odpowiedniego dodoatku do którego dodany zostanie dodatek historyczny
                int ktory = -1; //ktory dodatek
                for (int i = 0; i < rDodatki.Count; i++)
                {
                    CDNPrac.IDodatek rPrzedDodatek = (CDNPrac.IDodatek)rDodatki[i];
                    if (rPrzedDodatek.Nazwa == prac.DOH_Nazwa &&  rPrzedDodatek.Nieaktywny==0)
                    {
                        
                        ktory = i;
                        break;
                    }
                }
                if (ktory == -1)
                {
                    /*  if (MessageBox.Show("Pracownik o akronimie " + prac.PracKod + " nie posiada dodatku o nazwie: " + prac.DOH_Nazwa + ".\nCzy kontynuować proces dodawanie z pominięciem tego przypadku?", "Komunikat", MessageBoxButtons.YesNo, MessageBoxIcon.Error) == DialogResult.Yes)
                      {
                          // user clicked yes
                          continue;
                      }
                      else
                      {
                          // user clicked no
                          Wylogowanie();
                          return;
                      }*/
                    //tutaj dodajemy nowy dodatek
                    try
                    {



                        var rDods = rPracownik.Dodatki;
                        var rDod = rDods.AddNew(); //tutaj jednocześnie jest wykonywany AddNew na zapisach historycznych
                        var rDodsHis = rDod.Historia;
                        var rDodHis = rDodsHis[rDod.ID]; //pobieramy zapis historyczny

                        var rSklad = (CDNPlcfg.ITypSkladnika)Sesja.CreateObject("CDNPlcfg.TypySkladnikow", "TWP_Nazwa='" + prac.DOH_Nazwa + "'");

                        rDodHis.Skladnik.TypSkladnika = rSklad;
                        rDodHis.Wartosc1 = prac.DOH_Wartosc1;
                        rDodHis.Wartosc2 = 0;

                        rDod.Nazwa = rSklad.Nazwa;

                        var dDodOd = prac.DOH_WaznyOd;
                        var dDodDo = prac.DOH_WaznyDo;

                        var wazny = rDodHis.Wazny;
                        wazny.Od = dDodOd;
                        wazny.Do = dDodDo;

                        Sesja.Save();



                        // Zapisanie obiektow 
                        Sesja.Save();
                        rDodatki = rPracownik.Dodatki;
                    }
                    catch (Exception en)
                    {
                        MessageBox.Show("Błąd dodawania nowego dodatku: " + prac.DOH_Nazwa + " dla pracownika: " + prac.PracKod + "\n" + en.Message, "Komunikat", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }

                }
                else
                {
                    CDNPrac.IDodatek rDodatek = (CDNPrac.IDodatek)rDodatki[ktory];
                    DateTime pracWaznyDo = DateTime.MaxValue;

                    //Sprawdzam aktualność istniejących dodatków hist
                    bool flag = false;
                    ICollection rDodatkiHis = rDodatek.Historia;
                    int temp = rDodatkiHis.Count;


                    if (prac.DOH_WaznyOd == null)
                    {
                        #region przypadki waznych od 
                        for (int i = 0; i < rDodatkiHis.Count; i++)
                        {
                            CDNPrac.IDodatekHis tempDOH = (CDNPrac.IDodatekHis)rDodatkiHis[i];
                            //MessageBox.Show(Convert.ToString(tempDOH.ID)); przyp 5 i 8 
                            if (tempDOH.Wazny.Do > prac.DOH_WaznyOd && tempDOH.Wazny.Od > prac.DOH_WaznyOd && tempDOH.Wazny.Od.Subtract(prac.DOH_WaznyOd) > new TimeSpan(25, 0, 0, 0, 0))
                            {

                                //ważny do naszego dodawanego ma byc do starego wazny od -1 dzien
                                if (pracWaznyDo > tempDOH.Wazny.Od.AddDays(-1))
                                {
                                    pracWaznyDo = tempDOH.Wazny.Od.AddDays(-1);
                                }
                            }//5 i 8
                            else if (tempDOH.Wazny.Do > prac.DOH_WaznyOd && tempDOH.Wazny.Od > prac.DOH_WaznyOd && (tempDOH.Wazny.Od.Subtract(prac.DOH_WaznyOd) < new TimeSpan(25, 0, 0, 0, 0) && tempDOH.Wazny.Od.Subtract(prac.DOH_WaznyOd) > new TimeSpan(0, 0, 0, 0, 0)))
                            {

                                if (MessageBox.Show("Dla pracownika o akronimie " + prac.PracKod + " istnieje już dodatek historyczny dla dodatku " + rDodatek.Nazwa + " o dacie następującej po dacie wprowadzanego dodatku w tym samym miesiącu rozliczeniowym. \nCzy kontynułować przypisywanie dodatków?", "Komunikat", MessageBoxButtons.YesNo, MessageBoxIcon.Information) == DialogResult.Yes)
                                {
                                    continue;
                                }
                                else
                                {
                                    Wylogowanie();
                                    return;
                                }
                            }
                            //nie ma kolidacji w przyszłosci, ale poprzedni dodatek ma nachodzacy czas 2, 3, 7
                            else if (tempDOH.Wazny.Do >= prac.DOH_WaznyOd && tempDOH.Wazny.Od < prac.DOH_WaznyOd)
                            {

                                tempDOH.Wazny.Do = prac.DOH_WaznyOd.AddDays(-1);
                                Sesja.Save();
                            }//4 i 6
                            else if (tempDOH.Wazny.Od == prac.DOH_WaznyOd)
                            {

                                if (MessageBox.Show("Dla pracownika o akronimie " + prac.PracKod + " istnieje już dodatek historyczny " + rDodatek.Nazwa + " o tej samej dacie \"WaznyOd\". Czy zaktualizować istniejący zapis?", "Komunikat", MessageBoxButtons.YesNo, MessageBoxIcon.Information) == DialogResult.Yes)
                                {
                                    tempDOH.Wartosc1 = prac.DOH_Wartosc1;
                                    //tempDOH.Wazny.Do = DateTime.MaxValue;
                                    flag = true;
                                    Sesja.Save();

                                }
                                else
                                {
                                    // user clicked no
                                    flag = true;
                                    continue;
                                }

                            }//1
                            else if (tempDOH.Wazny.Do < prac.DOH_WaznyOd && tempDOH.Wazny.Od < prac.DOH_WaznyOd)
                            {

                                continue;
                            }
                            else
                            {
                                MessageBox.Show("Nieprzewidziany przypadek.");
                            }

                        }
                        #endregion
                    }
                    else
                    {
                        #region przypadki waznych Od i Do
                        for (int i = 0; i < rDodatkiHis.Count; i++)
                        {
                            CDNPrac.IDodatekHis tempDOH = (CDNPrac.IDodatekHis)rDodatkiHis[i];
                            //MessageBox.Show(Convert.ToString(tempDOH.ID)); przyp 6 10, 11
                            //
                            if (tempDOH.Wazny.Od > prac.DOH_WaznyOd && tempDOH.Wazny.Od < prac.DOH_WaznyDo)
                            {

                                if (MessageBox.Show("Dla pracownika o akronimie " + prac.PracKod + " istnieje już dodatek historyczny dla dodatku " + rDodatek.Nazwa + " o dacie następującej po dacie wprowadzanego dodatku w tym samym miesiącu rozliczeniowym. \nCzy kontynułować przypisywanie dodatków?", "Komunikat", MessageBoxButtons.YesNo, MessageBoxIcon.Information) == DialogResult.Yes)
                                {
                                    continue;
                                }
                                else
                                {
                                    Wylogowanie();
                                    return;
                                }
                            }
                            //nie ma kolidacji w przyszłosci, ale poprzedni dodatek ma nachodzacy czas 2,3,,4,8
                            else if (tempDOH.Wazny.Do >= prac.DOH_WaznyOd && tempDOH.Wazny.Od < prac.DOH_WaznyOd)
                            {

                                tempDOH.Wazny.Do = prac.DOH_WaznyOd.AddDays(-1);
                                Sesja.Save();
                            }//5,7,8,9
                            else if (tempDOH.Wazny.Od == prac.DOH_WaznyOd)
                            {

                                if (MessageBox.Show("Dla pracownika o akronimie " + prac.PracKod + " istnieje już dodatek historyczny " + rDodatek.Nazwa + " o tej samej dacie \"WaznyOd\". Czy zaktualizować istniejący zapis?", "Komunikat", MessageBoxButtons.YesNo, MessageBoxIcon.Information) == DialogResult.Yes)
                                {
                                    tempDOH.Wartosc1 = prac.DOH_Wartosc1;
                                    tempDOH.Wazny.Do = prac.DOH_WaznyDo;
                                    flag = true;
                                    Sesja.Save();

                                }
                                else
                                {
                                    // user clicked no
                                    flag = true;
                                    continue;
                                }

                            }//1
                            else if (tempDOH.Wazny.Do < prac.DOH_WaznyOd && tempDOH.Wazny.Od < prac.DOH_WaznyOd)
                            {

                                continue;
                            }//12
                            else if (tempDOH.Wazny.Od > prac.DOH_WaznyDo)
                            {
                                continue;
                            }
                            else
                            {
                                MessageBox.Show("Nieprzewidziany przypadek.");
                            }


                               

                        }

                        #endregion
                    }

                    if (flag)
                    {
                        continue;
                    }

                    //MessageBox.Show("dodaje nowy");
                    CDNPrac.IDodatekHis rDodatekHis2 = (CDNPrac.IDodatekHis)rDodatkiHis.AddNew();

                    //Przypisywanie nazwy/Typu składnika
                    var rSkladnik = (CDNPlcfg.ITypSkladnika)Sesja.CreateObject("CDNPlcfg.TypySkladnikow", "TWP_Nazwa='" + prac.DOH_Nazwa + "'");
                    rDodatekHis2.Skladnik.TypSkladnika = (CDNPlcfg.TypSkladnika)rSkladnik;
                    if (rSkladnik.NazwaPelna != "")
                    {
                        rDodatekHis2.Nazwa = rSkladnik.NazwaPelna;
                    }
                    else
                    {
                        rDodatekHis2.Nazwa = rSkladnik.Nazwa;
                    }

                    //Przypiswanie wartosci1
                    rDodatekHis2.Wartosc1 = prac.DOH_Wartosc1;
                    rDodatekHis2.Wartosc2 = 0;

                    ////Przypisanie symbolu listy płac
                    //try
                    //{
                    //    CDNHeal.IDefinicjaDokumentu rDD = (CDNHeal.IDefinicjaDokumentu)Sesja.CreateObject("CDN.DefinicjeDokumentow", "DDF_Symbol = '" + prac.ListaPlac + "'");
                    //    rDodatekHis2.ListaPlacSymbol = (CDNHeal.DefinicjaDokumentu)rDD;
                    //}
                    //catch (Exception)
                    //{
                    //    if (MessageBox.Show("Nie znaleziono symbolu listy plac: " + prac.ListaPlac+".\nCzy kontynuować proces dodawanie z pominięciem tego przypadku?", "Komunikat", MessageBoxButtons.YesNo, MessageBoxIcon.Error) == DialogResult.Yes)
                    //    {
                    //        // user clicked yes
                    //        continue;
                    //    }
                    //    else
                    //    {
                    //        // user clicked no
                    //        Wylogowanie();
                    //        return;
                    //    }
                    //}


                    //Przypisanie WażnnyOd
                    DateTime dDodatekOd = Convert.ToDateTime(prac.DOH_WaznyOd);
                    PeriodExt rOkres = rDodatekHis2.Okres;
                    rOkres.Od = dDodatekOd;
                    //if(prac.DOH_WaznyDo != null)
                    //    rOkres.Do = Convert.ToDateTime(prac.DOH_WaznyDo);
                    PeriodExt rWazny = rDodatekHis2.Wazny;
                    rWazny.Od = dDodatekOd;
                    if (prac.DOH_WaznyDo != null )
                        rWazny.Do = Convert.ToDateTime(prac.DOH_WaznyDo);


                    Sesja.Save();
                }
            }
            Cursor.Current = Cursors.Default;
            Wylogowanie();
            MessageBox.Show("Zakończono dodawanie dodatków historycznych.", "Komunikat",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
            #endregion

        }
        #region wybranie pliku
        void WybraniePliku()
        {
            OpenFileDialog openFileDialog1 = new OpenFileDialog()
            {
                InitialDirectory = Directory.GetCurrentDirectory(),
                DefaultExt = "xlsx",
                Filter = "Pliki xlsx (*.xlsx)|*.xlsx|Pliki xls (*.xls)|*.xls|Wszystkie pliki (*.*)|*.*",
            };
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                tB1.Text = openFileDialog1.FileName;
                //MessageBox.Show(path);
            }
            return;
        }

        private void Button2_Click(object sender, EventArgs e)
        {
            WybraniePliku();
        }
        private void TextBox1_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            WybraniePliku();
        }
        #endregion
    }
}

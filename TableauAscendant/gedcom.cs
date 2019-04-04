
using PA.Utilities;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Diagnostics;
using System.IO;
using System.Runtime.CompilerServices;

namespace TableauAscendant
{
    class GEDCOMClass
    {
        /// <summary>
        /// nom du fichier de log
        /// </summary>
        public bool LOGACTIF = false;
        /// <summary>
        /// nom du fichier de log
        /// </summary>
        public string FICHIERLOG = "01TA-szUejmCjMh.log";
        /// <summary>
        /// nom du fichier de log
        /// </summary>
        public string DossierPDF = "C:/Users/dapam/Documents/TableauAscendant/";


        
        private string[] dataGEDCOM;

        public class ListeIndividu : IEquatable<ListeIndividu>
        {
            public int Reference { get; set; }
            public string ID { get; set; }
            public string Patronyme { get; set; }
            public string Prenom { get; set; }
            public string Sex { get; set; }
            public string Titre { get; set; }
            public string DateNaissance { get; set; }
            public string LieuNaissance { get; set; }
            public string VilleNaissance { get; set; }
            public string FormatPhoto { get; set; }
            public string FichierPhoto { get; set; }
            public string DateDeces { get; set; }
            public string LieuDeces { get; set; }
            public string VilleDeces { get; set; }
            public string DateInhumation { get; set; }
            public string LieuInhumation { get; set; }
            public string VilleInhumation { get; set; }
            public string DateBapteme { get; set; }
            public string LieuBapteme { get; set; }
            public string VilleBapteme { get; set; }
            public string Note { get; set; }
            public string FamilleEpoux { get; set; }
            public string FamilleEnfant { get; set; }


            public override string ToString()
            {
                return "ID: " + ID + "   Prenom: " + Prenom + "   Patronyme: " + Patronyme;
            }
            public override bool Equals(object obj)
            {
                if (obj == null) return false;
                if (!(obj is ListeIndividu objAsPart)) return false;
                else return Equals(objAsPart);
            }
            public override int GetHashCode()
            {
                return Reference;
            }
            public bool Equals(ListeIndividu other)
            {
                if (other == null) return false;
                return (this.ID.Equals(other.ID));
            }
        }
        public class ListeFamille : IEquatable<ListeIndividu>
        {
            public int Reference { get; set; }
            public string ID { get; set; }
            public string IDEpoux { get; set; }
            public string IDEpouse { get; set; }
            public string DateMariage { get; set; }
            public string LieuMariage { get; set; }
            public string VilleMariage { get; set; }
            public string IDEnfant { get; set; }



            public override string ToString()
            {
                return "ID: " + ID + "   Époux: " + IDEpoux + "   Epouse: " + IDEpouse;
            }
            public override bool Equals(object obj)
            {
                if (obj == null) return false;
                //ListeIndividu objAsPart = obj as ListeIndividu;


                if (!(obj is ListeIndividu objAsPart)) return false;
                else return Equals(objAsPart);
            }
            public override int GetHashCode()
            {
                return Reference;
            }
            public bool Equals(ListeIndividu other)
            {
                if (other == null) return false;
                return (this.ID.Equals(other.ID));
            }
        }
        List<ListeIndividu> listeIndividu = new List<ListeIndividu>();
        List<ListeFamille> listeFamille = new List<ListeFamille>();

        private string ExtraireID(string s) {
            int p1 = s.IndexOf("@") + 2 ;
            int p2 = s.IndexOf("@", s.IndexOf('@') + 1);
            if ((p1 >= 0) && (p2>= 0)) 
            {
                // @I0001@
                // 2345678
                return s.Substring( p1-1 , p2 - p1 + 1 );
            } else
            {
                return "";
            }
        }

        private string[] ExtrairePatronyme(string s)
        {
            // 1 NAME Nnn Nnnn  /Nnnnnnnn/
            int p1 = s.IndexOf("/");
            int p2 = s.IndexOf("/", s.IndexOf('/') + 1);
            string[] a = new string[2];
            if (p1 == 7 && p2 == 8)
            {
                a[0] = "";
                a[1] = "";
                return a;
            }
            a[0] = s.Substring(7, p1 - 8);
            a[1] = s.Substring(p1 + 1, p2-p1-1);
            return  a;
        }
        public bool Individu()
        {
            // Individu ############################################################
            
            int i = 1;
            
            string ID;
            //string IDs;
            bool loop;
            int objet = 0;
            bool photoTrouver = false;

            //try
            {
                listeIndividu.ToArray();
                do
                {
                    if (dataGEDCOM[i].Contains("0 @I"))                                                    // Individu
                    {
                        string patronyme = "";
                        string prenom = "";
                        string sex = "";
                        string titre;
                        string dateNaissance = "";
                        string lieuNaissance = "";
                        string villeNaissance = "";
                        string formatPhoto = "";
                        string fichierPhoto = "";
                        string dateDeces = "";
                        string lieuDeces = "";
                        string villeDeces = "";
                        string dateInhumation = "";
                        string lieuInhumation = "";
                        string villeInhumation = "";
                        string dateBapteme = "";
                        string lieuBapteme = "";
                        string villeBapteme = "";
                        string note = "";
                        string familleEpoux = "";
                        string familleEnfant = "";
                        if (dataGEDCOM[i].Length == 5)
                        {
                            dataGEDCOM[i] = dataGEDCOM[i] + " ";
                        }
                        familleEpoux = "";
                        note = "";
                        ID = ExtraireID(dataGEDCOM[i]);
                        //IDs = ID.ToString();
                        i++;
                        photoTrouver = false;
                        while ((int)Char.GetNumericValue(dataGEDCOM[i][0]) > 0)
                        {
                            dataGEDCOM[i] = dataGEDCOM[i] + " ";
                            if ( dataGEDCOM[i].Contains("1 BIRT"))                                              // Naissance
                            {
                                i++;
                                loop = false;
                                while ((int)Char.GetNumericValue(dataGEDCOM[i][0]) > 1)
                                {
                                    loop = true;
                                    if (dataGEDCOM[i].Contains("2 DATE") && dataGEDCOM[i].Length > 7 )          // date
                                    {
                                        // 2 DATE 17 JUN 1871
                                        dateNaissance = dataGEDCOM[i].Substring(7, dataGEDCOM[i].Length - 7);
                                    }
                                    if (dataGEDCOM[i].Contains("2 PLAC") && dataGEDCOM[i].Length > 7)           // lieu
                                    {
                                        lieuNaissance = dataGEDCOM[i].Substring(7, dataGEDCOM[i].Length - 7);
                                    }
                                    if (dataGEDCOM[i].Contains("3 CITY") && dataGEDCOM[i].Length > 7 )          // ville
                                    {
                                        villeNaissance = dataGEDCOM[i].Substring(7, dataGEDCOM[i].Length - 7);
                                    }
                                    i++;
                                }
                                if (loop)
                                    i--;
                            }
                            if (dataGEDCOM[i].Contains("1 NAME") && patronyme == "" && prenom == ""  )                // Nom
                            {
                                string[] temp = ExtrairePatronyme(dataGEDCOM[i]);
                                prenom = temp[0];
                                patronyme = temp[1];
                            }
                            if (dataGEDCOM[i].Contains("1 SEX") && dataGEDCOM[i].Length > 6 )                   // Sex
                            {
                                // 1 SEX F
                                sex = dataGEDCOM[i].Substring(6, 1);
                            }
                            if (dataGEDCOM[i].Contains("1 TITL") && dataGEDCOM[i].Length > 7 )                  // Titre
                            {
                                    titre = dataGEDCOM[i].Substring(7, dataGEDCOM[i].Length - 8);
                            }
                            if (dataGEDCOM[i].Contains("1 OBJE") )                                              //  Photo individu, la première
                            {
                                i++;
                                objet++;
                                loop = false;
                                while ((int)Char.GetNumericValue(dataGEDCOM[i][0]) > 1)
                                {
                                    loop = true;
                                    if (dataGEDCOM[i].Contains("2 FORM URL") && dataGEDCOM[i].Length > 11)
                                    {
                                        string temp = "AA";
                                        temp = temp + "BB";
                                        // code pour map
                                    }
                                    else if (dataGEDCOM[i].Contains("4 FORM BMP"))                         // format photo BMP
                                    {
                                        formatPhoto = dataGEDCOM[i].Substring(7, dataGEDCOM[i].Length - 11);
                                        photoTrouver = true;
                                    }
                                    else if (dataGEDCOM[i].Contains("4 FORM bmp"))                         // format photo bmp
                                    {
                                        formatPhoto = dataGEDCOM[i].Substring(7, dataGEDCOM[i].Length - 11);
                                        photoTrouver = true;
                                    }
                                    else if (dataGEDCOM[i].Contains("4 FORM PNG"))                         // format photo PNG
                                    {
                                        formatPhoto = dataGEDCOM[i].Substring(7, dataGEDCOM[i].Length - 11);
                                        photoTrouver = true;
                                    }
                                    else if (dataGEDCOM[i].Contains("4 FORM image/x-png"))                 // # format photo image/x-png
                                    {
                                        formatPhoto = dataGEDCOM[i].Substring(7, dataGEDCOM[i].Length - 11);
                                        photoTrouver = true;
                                    }
                                    else if (dataGEDCOM[i].Contains("4 FORM JPEG"))                        // format photo JPEG
                                    {
                                        formatPhoto = dataGEDCOM[i].Substring(7, dataGEDCOM[i].Length - 11);
                                        photoTrouver = true;
                                    }
                                    else if (dataGEDCOM[i].Contains("4 FORM image/pjpeg"))                 // format photo image/pjpeg
                                    {
                                        formatPhoto = dataGEDCOM[i].Substring(7, dataGEDCOM[i].Length - 19);
                                        photoTrouver = true;
                                    }
                                    else if (dataGEDCOM[i].Contains("FORM GIF"))                           // format photo GIF
                                    {
                                        formatPhoto = dataGEDCOM[i].Substring(7, dataGEDCOM[i].Length - 10);
                                        photoTrouver = true;
                                    }
                                    else if (dataGEDCOM[i].Contains("4 FORM gif"))                         // format photo gif
                                    {
                                        formatPhoto = dataGEDCOM[i].Substring(7, dataGEDCOM[i].Length - 11);
                                        photoTrouver = true;
                                    }
                                    else if ((dataGEDCOM[i].Contains("4 FILE")) && photoTrouver && objet > 0)// # fichier photo
                                    {
                                        fichierPhoto = dataGEDCOM[i].Substring(7, dataGEDCOM[i].Length -7);
                                        photoTrouver = false;
                                    }
                                    i++;
                                }
                                if (loop)
                                {
                                    i--;
                                }
                            }
                            if (dataGEDCOM[i].Contains("1 DEAT"))                                           // Décès
                            {
                                i++;
                                loop = false;
                                while ((int)Char.GetNumericValue(dataGEDCOM[i][0]) > 1)
                                {
                                    loop = true;
                                    if (dataGEDCOM[i].Contains("2 DATE") && dataGEDCOM[i].Length > 7)       //date
                                    {
                                        dateDeces = dataGEDCOM[i].Substring(7, dataGEDCOM[i].Length - 7);
                                    }
                                    if (dataGEDCOM[i].Contains("2 PLAC") && dataGEDCOM[i].Length > 7)       // lieu
                                    {
                                        lieuDeces = dataGEDCOM[i].Substring(7, dataGEDCOM[i].Length - 7);
                                    }
                                    if (dataGEDCOM[i].Contains("3 CITY") && dataGEDCOM[i].Length > 7 )      // ville
                                    {
                                        villeDeces = dataGEDCOM[i].Substring(7, dataGEDCOM[i].Length - 7);
                                    }
                                    i++;
                                }
                                if (loop)
                                {
                                    i--;
                                }
                            }
                            if (dataGEDCOM[i].Contains("1 BURI"))                                               // Inhumation
                            {
                                i++;
                                loop = false;
                                while ((int)Char.GetNumericValue(dataGEDCOM[i][0]) > 1)
                                {
                                    loop = true;
                                    if (dataGEDCOM[i].Contains("2 DATE") && dataGEDCOM[i].Length > 7)           // date
                                    {
                                        dateInhumation = dataGEDCOM[i].Substring(7, dataGEDCOM[i].Length - 7);
                                    }
                                    if (dataGEDCOM[i].Contains("2 PLAC") && dataGEDCOM[i].Length > 7)           // lieu
                                    {
                                        lieuInhumation = dataGEDCOM[i].Substring(7, dataGEDCOM[i].Length - 7);
                                    }
                                    if (dataGEDCOM[i].Contains("3 CITY") )                                      // ville
                                    {
                                        villeInhumation = dataGEDCOM[i].Substring(7, dataGEDCOM[i].Length - 7);
                                    }
                                    i++;
                                }
                                if (loop)
                                {
                                    i--;
                                }
                            }
                            if (dataGEDCOM[i].Contains("1 BAPM"))                                               // Bapthème
                            {
                                i++;
                                loop = false;
                                while ((int)Char.GetNumericValue(dataGEDCOM[i][0]) > 1)
                                {
                                    loop = true;
                                    if (dataGEDCOM[i].Contains("2 DATE") && dataGEDCOM[i].Length > 7)           // date
                                    {
                                        dateBapteme = dataGEDCOM[i].Substring(7, dataGEDCOM[i].Length - 7);
                                    }
                                    else  if (dataGEDCOM[i].Contains("2 PLAC") && dataGEDCOM[i].Length > 7)     // lieu
                                    {
                                        lieuBapteme = dataGEDCOM[i].Substring(7, dataGEDCOM[i].Length - 7);
                                    }
                                    else if (dataGEDCOM[i].Contains("3 PLAC") && dataGEDCOM[i].Length > 7)      // ville
                                    {
                                        villeBapteme = dataGEDCOM[i].Substring(7, dataGEDCOM[i].Length - 7);
                                    }
                                    i++;
                                }
                                if (loop)
                                {
                                    i--;
                                }
                            }
                            if (dataGEDCOM[i].Contains("1 NOTE") && dataGEDCOM[i].Length > 7)                   // Note
                            {
                                string tmp = ExtraireID(dataGEDCOM[i]);
                                if (note == "")
                                {
                                    note = tmp.ToString();
                                }
                                else
                                {
                                    note = note + " " + tmp.ToString();
                                }
                            }
                            if (dataGEDCOM[i].Contains("1 FAMS") && dataGEDCOM[i].Length > 7 )                  // FAMS famille des époux
                            {
                                string tmp = ExtraireID(dataGEDCOM[i]);
                                if (familleEpoux == "")
                                {
                                    familleEpoux = tmp.ToString();
                                }
                                else
                                {
                                    familleEpoux = familleEpoux + " " + tmp.ToString();
                                }
                            }
                            if (dataGEDCOM[i].Contains("1 FAMC") && dataGEDCOM[i].Length > 7)                   // FAMC enfant de la famille
                            {
                                familleEnfant = ExtraireID(dataGEDCOM[i]).ToString();
                            }
                            i++;
                        }
                        i--;
                        // fin de while int(dataGEDCOM[i][0]) <> 0:
                        listeIndividu.Add(new ListeIndividu() {
                            ID = ID,
                            Prenom = prenom,
                            Patronyme = patronyme,
                            Sex = sex,
                            DateNaissance = dateNaissance,
                            LieuNaissance = lieuNaissance,
                            VilleNaissance = villeNaissance,
                            DateDeces = dateDeces,
                            LieuDeces = lieuDeces,
                            VilleDeces = villeDeces,
                            DateInhumation = dateInhumation,
                            LieuInhumation = lieuInhumation,
                            VilleInhumation = villeInhumation,
                            FamilleEpoux = familleEpoux,
                            FamilleEnfant = familleEnfant
                        });
                    }
                    i++;
                } while ( i < dataGEDCOM.Length);
                
                //foreach (ListeIndividu aPart in listeIndividu)
                //{
                //    if (aPart.ID == 1444) {
                //    }
                //}
                return true;
            } //catch (Exception msg)
            //{
            //    //SystemSounds.Beep.Play();
            //    MessageBox.Show("En lisant la ligne " + i.ToString() + " du fichier GEDCOM.\r\n\r\n" + msg.Message, "Problème ?",
            //                     MessageBoxButtons.OK,
            //                     MessageBoxIcon.Warning);
            //    return false;
            //}
        }
        public bool Famille()
        {
            int i = 1;
            string ID = "";
            int loop = 0;
            try
            {
                listeIndividu.ToArray();
                do
                {
                    if (dataGEDCOM[i].Contains("0 @F"))                                                     // Famille
                    {
                        string IDenfant = "";
                        string IDepoux = "";
                        string IDepouse = "";
                        string dateMariage = "";
                        string lieuMariage = "";
                        string villeMariage = "";
                        ID = ExtraireID(dataGEDCOM[i]);
                        i++;
                        while ((int)Char.GetNumericValue(dataGEDCOM[i][0]) > 0)
                        {
                            if (dataGEDCOM[i].Contains("1 HUSB") && dataGEDCOM[i].Length > 7)               // conjoint
                            {
                                IDepoux = ExtraireID(dataGEDCOM[i]);
                                i++;
                                //while ((int)Char.GetNumericValue(dataGEDCOM[i][0]) > 1)
                                //{
                                //}
                            }
                            if (dataGEDCOM[i].Contains("1 WIFE") && dataGEDCOM[i].Length > 7)               // conjointe
                            {
                                IDepouse = ExtraireID(dataGEDCOM[i]);
                                i++;
                                //while ((int)Char.GetNumericValue(dataGEDCOM[i][0]) > 1)
                                //{
                                //    i++;
                                //}
                            }


                            if (dataGEDCOM[i].Contains("1 MARR"))               //  mariage
                            {
                                i++;
                                loop = 0;
                                while ((int)Char.GetNumericValue(dataGEDCOM[i][0]) > 1)
                                {
                                    loop = 0;
                                    if (dataGEDCOM[i].Contains("2 DATE") && dataGEDCOM[i].Length > 7)       // date mariage
                                    {
                                        dateMariage = dataGEDCOM[i].Substring(7, dataGEDCOM[i].Length - 7);
                                    }
                                    if (dataGEDCOM[i].Contains("2 PLAC") && dataGEDCOM[i].Length > 7)       // lieu mariage
                                        lieuMariage = dataGEDCOM[i].Substring(7, dataGEDCOM[i].Length - 7);
                                    if (dataGEDCOM[i].Contains("3 CITY") && dataGEDCOM[i].Length > 7)       // ville
                                        // Ville mariage
                                        villeMariage  = dataGEDCOM[i].Substring(7, dataGEDCOM[i].Length - 7);
                                    i++;
                                        }
                                    if (loop > 0) {
                                        i--;
                                    }
                            }



                            if (dataGEDCOM[i].Contains("1 CHIL") && dataGEDCOM[i].Length > 7)                   // enfant
                            {
                                string tmp = ExtraireID(dataGEDCOM[i]);
                                if (IDenfant == "")
                                {
                                    IDenfant = tmp.ToString();
                                    i++;
                                }
                                else
                                {
                                    IDenfant = IDenfant + " " + tmp.ToString();
                                    i++;
                                }
                            }

                            if (loop == i)
                            {
                                i++;
                            }
                            else
                            {
                                loop = i;
                            }
                        }
                        i--;

                        //self._listeFamille[ID] =[\epoux, epouse, enfant]

                        listeFamille.Add(new ListeFamille()
                        {
                            ID = ID,
                            IDEpoux = IDepoux,
                            IDEpouse = IDepouse,
                            DateMariage = dateMariage,
                            LieuMariage = lieuMariage,
                            VilleMariage = villeMariage,
                            IDEnfant = IDenfant

                        });
                    }
                    i++;
                } while (i < dataGEDCOM.Length) ;


            } catch (Exception msg)
              {
                  //SystemSounds.Beep.Play();
                  MessageBox.Show("En lisant la ligne pour famille" + i.ToString() + " du fichier GEDCOM.\r\n\r\n" + msg.Message, "Problème ?",
                                   MessageBoxButtons.OK,
                                   MessageBoxIcon.Warning);
                  return false;
              }
            return true;
        }

        public void LireGEDCOM(string fichier)
        {
            dataGEDCOM = System.IO.File.ReadAllLines(@fichier);
        }

        public string AvoirCoupleID(string IDm, string IDf)
        {
            foreach (ListeFamille info in listeFamille)
            {
                if (info.IDEpoux  == IDm && info.IDEpouse == IDf)
                {
                    return info.ID;
                }
            }
            return "";
        }

        public string AvoirDateDeces(string ID)
        {
            foreach (ListeIndividu info in listeIndividu)
            {
                if (info.ID == ID)
                {
                    return info.DateDeces;
                }
            }
            return "";
        }

        public string AvoirDateInhumation(string ID)
        {
            foreach (ListeIndividu info in listeIndividu)
            {
                if (info.ID == ID)
                {
                    return info.DateInhumation;
                }
            }
            return "";
        }

        public string AvoirDateNaissance(string ID)
        {
            foreach (ListeIndividu info in listeIndividu)
            {
                if (info.ID == ID)
                {
                    return info.DateNaissance;
                }
            }
            return "";
        }
        public string AvoirEnfant(string ID)
        {
            foreach (ListeFamille info in listeFamille)
            {
                if (info.ID == ID)
                {
                    return info.IDEnfant;
                }
            }
            return "";
        }
        public string AvoirEpouse(string ID)
        {
            foreach (ListeFamille info in listeFamille)
            {
                if (info.ID == ID)
                {
                    return info.IDEpouse;
                }
            }
            return "";
        }
        public string AvoirEpoux(string ID)
        {
            foreach (ListeFamille info in listeFamille)
            {
                if (info.ID == ID)
                {
                    return info.IDEpoux;
                }
            }
            return "";
        }
        public string AvoirFamilleEnfant(string ID)
        {
            foreach (ListeIndividu info in listeIndividu)
            {
                if (info.ID == ID)
                {
                    return info.FamilleEnfant;
                }
            }
            return "";
        }
        public string AvoirFamilleEpoux(string ID)
        {
            foreach (ListeIndividu info in listeIndividu)
            {
                if (info.ID == ID)
                {
                    return info.FamilleEpoux;
                }
            }
            return "";
        }

        public string AvoirDateMariage(string ID)
        {
            foreach (ListeFamille info in listeFamille)
            {
                if (info.ID == ID)
                {
                    return info.DateMariage;
                }
            }
            return "";
        }

        public string AvoirPatronyme(string ID)
        {
            foreach (ListeIndividu info in listeIndividu)
            {
                if (info.ID == ID)
                {
                    return info.Patronyme;
                }
            }
            return "";
        }

        public string AvoirLieuDeces(string ID)
        {
            foreach (ListeIndividu info in listeIndividu)
            {
                if (info.ID == ID)
                {
                    return info.LieuDeces;
                }
            }
            return "";
        }

        public string AvoirLieuMariage(string ID)
        {
            foreach (ListeFamille info in listeFamille)
            {
                if (info.ID == ID)
                {
                    return info.LieuMariage;
                }
            }
            return "";
        }

        public string AvoirLieuNaissance(string ID)
        {
            foreach (ListeIndividu info in listeIndividu)
            {
                if (info.ID == ID)
                {
                    return info.LieuNaissance;
                }
            }
            return "";
        }
        public string AvoirLieuInhumation(string ID)
        {
            foreach (ListeIndividu info in listeIndividu)
            {
                if (info.ID == ID)
                {
                    return info.LieuInhumation ;
                }
            }
            return "";
        }
        public string AvoirEndroitDeces(string ID)
        {
            foreach (ListeIndividu info in listeIndividu)
            {
                string s =  AvoirVilleDeces(ID);
                if (s != "")
                {
                    return s;
                }
                else
                {
                    return AvoirLieuDeces(ID);
                }
            }
            return "";
        }
        public string AvoirEndroitMariage(string ID)
        {
            foreach (ListeIndividu info in listeIndividu)
            {
                string s = AvoirVilleMariage(ID);
                if (s != "")
                {
                    return s;
                }
                else
                {
                    return AvoirLieuMariage(ID);
                }
            }
            return "";
        }
        public string AvoirEndroitNaissance(string ID)
        {
            foreach (ListeIndividu info in listeIndividu)
            {
                string s = AvoirVilleNaissance(ID);
                if ( s != "" )
                {
                    return s;
                } else
                {
                    return AvoirLieuNaissance(ID);
                }
            }
            return "";
        }

        public string AvoirPrenom(string ID)
        {
            foreach (ListeIndividu info in listeIndividu)
            {
                if (info.ID == ID)
                {
                    return info.Prenom;
                }
            }
            return "";
        }
        public string AvoirSex(string ID)
        {
            foreach (ListeIndividu info in listeIndividu)
            {
                if (info.ID == ID)
                {
                    return info.Sex;
                }
            }
            return "";
        }
        public string AvoirVilleBapteme(string ID)
        {
            foreach (ListeIndividu info in listeIndividu)
            {
                if (info.ID == ID)
                {
                    return info.VilleBapteme;
                }
            }
            return "";
        }
        public string AvoirVilleNaissance(string ID)
        {
            foreach (ListeIndividu info in listeIndividu)
            {
                ;
                if (info.ID == ID)
                {
                    return info.VilleNaissance;
                }
            }
            return "";
        }
        public string AvoirVilleDeces(string ID)
        {
            foreach (ListeIndividu info in listeIndividu)
            {
                if (info.ID == ID)
                {
                    return info.VilleDeces;
                }
            }
            return "";
        }

            public string AvoirVilleMariage(string ID)
        {
            foreach (ListeFamille info in listeFamille)
            {
                if (info.ID == ID)
                {
                    return info.VilleMariage;
                }
            }
            return "";
        }

        public void EffacerDataGEDCOM()
        {
            for (int f = listeIndividu.Count - 1; f > 0; f--)
            {
                listeIndividu.RemoveAt(f);
            }
            listeFamille.Clear();
            for (int f = listeFamille.Count - 1; f > 0; f--)
            {
                listeIndividu.RemoveAt(f);
            }
            listeFamille.Clear();

        }
        public List<string> RechercheIndividu(string nom, string prenom)
        {
            List<string> ID = new List<string>();
            string n = nom.ToLower();
            string p = prenom.ToLower();
            foreach (ListeIndividu info in listeIndividu)
            {
                string nn = info.Patronyme.ToLower();
                string pp = info.Prenom.ToLower();
                if (nn.Contains(n) && pp.Contains(p))
                {
                    ID.Add(info.ID);
                }
            }
            return ID;
        }
        static void M(string message, [CallerLineNumber] int lineNumber = 0, [CallerMemberName] string caller = null)
        {
            Console.WriteLine(lineNumber + " " + caller + " " + message);
        }
        private void ZXCV(string message, [CallerLineNumber] int lineNumber = 0, [CallerMemberName] string caller = null)
        {
            if (LOGACTIF && Environment.UserName == "dapam")
            {
                string fichier;
                try
                {
                    if (Directory.Exists(DossierPDF))
                    {
                        fichier = DossierPDF + "\\" + FICHIERLOG;
                    }
                    else
                    {
                        fichier = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + "\\" + FICHIERLOG;
                    }

                    using (StreamWriter ligne = File.AppendText(fichier))
                    {
                        ligne.WriteLine(lineNumber + " " + caller + " " + message);
                    }
                }
                catch { }
            }
        }
    }
}

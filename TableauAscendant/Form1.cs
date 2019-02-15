// compile with: -doc:Form1.xml 
using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Windows.Forms;
using System.Diagnostics;
using System.IO;
//using System.Linq;
using System.Drawing;
using System.Drawing.Drawing2D;
//using System.Text;
using System.Runtime.CompilerServices;
using System.Media;
//using System.Reflection;
using PdfSharp;
using PdfSharp.Drawing;
using PdfSharp.Drawing.Layout;
using PdfSharp.Pdf;
using TableauAscendant;
using Microsoft.Win32;

namespace WindowsFormsApp1
{
    ///<Summary>
    /// 
    ///</Summary>
    public partial class TableauAscendant : Form
    {   /// <summary>
        /// nom du fichier de log
        /// </summary>
        public bool LOGACTIF = true;
        /// <summary>
        /// position de la colonne SOSA dans le tableau grille
        /// </summary>
        public const int SOSA =              0;
        /// <summary>
        /// position de la colonne PAGE dans le tableau grille
        /// </summary>
        public const int PAGE =              1;
        /// <summary>
        /// position de la colonne GENERATION dans le tableau grille
        /// </summary>
        public const int GENERATION =        2;
        /// <summary>
        /// position de la colonne TABLEAU dans le tableau grille
        /// </summary>
        public const int TABLEAU =           3;
        /// <summary>
        /// position de la colonne PATRONYME dans le tableau grille
        /// </summary>
        public const int PATRONYME =         4;
        /// <summary>
        /// position de la colonne PRENOM dans le tableau grille
        /// </summary>
        public const int PRENOM =            5;
        /// <summary>
        /// position de la colonne NOMTRI dans le tableau grille
        /// </summary>
        public const int NOMTRI =            6;
        /// <summary>
        /// position de la colonne NELE dans le tableau grille
        /// </summary>
        public const int NELE =              7;
        /// <summary>
        /// position de la colonne NELIEU dans le tableau grille
        /// </summary>
        public const int NELIEU =            8;
        /// <summary>
        /// position de la colonne DELE dans le tableau grille
        /// </summary>
        public const int DELE =              9;
        /// <summary>
        /// position de la colonne DELIEU dans le tableau grille
        /// </summary>
        public const int DELIEU =           10;
        /// <summary>
        /// position de la colonne MALE dans le tableau grille
        /// </summary>
        public const int MALE =             11;
        /// <summary>
        /// position de la colonne MALIEU dans le tableau grille
        /// </summary>
        public const int MALIEU =           12;
        /// <summary>
        /// position de la colonne IDg dans le tableau grille
        /// </summary>
        public const int IDg =              13;
        /// <summary>
        /// position de la colonne IDFAMILLEENFANT dans le tableau grille
        /// </summary>
        public const int IDFAMILLEENFANT =  14;
        /// <summary>
        /// position de la colonne NOTE1 dans le tableau grille
        /// </summary>
        public const int NOTE1 =            15;
        /// <summary>
        /// position de la colonne Note2 dans le tableau grille
        /// </summary>
        public const int NOTE2 =            16;
        /// <summary>
        /// Nom du programme
        /// </summary>
        public const string NomPrograme = "TableauAscendant";
        /// <summary>
        /// largeur maximun de nom dans fiche
        /// </summary>
        public const int LARGEURNOMFICHE = 144;
        /// <summary>
        /// largeur maximun du texte dans fiche
        /// </summary>
        public const int LARGEURTEXTEFICHE = 138;
        /// <summary>
        /// numero du SOSA courant
        /// </summary>
        public int sosaCourant = 0;
        /// <summary>
        /// nom du fichier de log
        /// </summary>
        public string  FICHIERLOG = "01TA-szUejmCjMh.log";
        /// <summary>
        /// nom du fichier de la grille
        /// </summary>
        public string FICHIERGRILLE = "01TA-grille.txt";



        /// <summary>
        /// liste qui contient toutes les informations pour généré les tableaux
        /// </summary>
        public string[,] liste = new string[512,17];
        
        /// <summary>
        /// Vrai si la grille à été modifier
        /// </summary>
        public Boolean Modifier = false;
        /// <summary>
        /// liste de recherche
        /// </summary>
        public int[] rechercheListe = new int[512]; // 512 lignes

        /// <summary>
        /// Nom du fichier courant
        /// </summary>
        public string FichierCourant ="";
        /// <summary>
        /// nom du fichier GEDCOM
        /// </summary>
        public string FichierGEDCOM = "";
        /// <summary>
        /// 
        /// </summary>
        public string argument = "";


        // pdf

        XFont font8 = new XFont("Arial", 8, XFontStyle.Regular);
        XFont font7 = new XFont("Arial", 7, XFontStyle.Regular);
        XFont font6 = new XFont("Arial", 6, XFontStyle.Regular);
        XFont font5 = new XFont("Arial", 5, XFontStyle.Regular);
        XFont font8B = new XFont("Arial", 8, XFontStyle.Bold);
        /// <summary>
        /// nom du dossier où seront enregistrer les fichiers PDF
        /// </summary>
        public string DossierPDF = "";
        /// <summary>
        /// valeur pour un pouce
        /// </summary>
        public XUnit POUCE = XUnit.FromInch(1);

        //couleur
        //Color arrierePlanForm = Color.FromArgb(102, 204, 255);
        //Color arrierePlanBoite = Color.FromArgb(99, 255, 255);
        Color couleurChamp = Color.White;
        Color couleurTextTropLong = Color.Yellow;

       GEDCOMClass GEDCOM = new GEDCOMClass();

        //  Function **************************************************************************************************************************

        private void    AfficherData()
        {
            if (ChoixSosaComboBox.Text == "")
            {
                // enlève les cases
                Sosa1PatronymeTextBox.Visible = false;
                Sosa1PrenomTextBox.Visible = false;
                Sosa1NeTextBox.Visible = false;
                Sosa1NeEndroitTextBox.Visible = false;
                Sosa1DeTextBox.Visible = false;
                Sosa1DeEndroitTextBox.Visible = false;
                Sosa1MaTextBox.Visible = false;
                Sosa1MaEndroitTextBox.Visible = false;
                Sosa1MaTextBox.Visible = false;
                Sosa1MaEndroitTextBox.Visible = false;
                sosa1LigneVertical.Visible = true;
                RectangleSosaConjoint1.Visible = true;
                Sosa1MaEtiquettetBox.Visible = true;
                Sosa1LieuEtiquettetBox.Visible = true;
                Note1.Visible = false;
                Note2.Visible = false;
                NoteDuHautLb.Visible = false;
                NoteDuBasLb.Visible = false;
                RectangleSosa1.BorderColor = Color.Black;

                Sosa2Label.Visible = false;
                Sosa2PatronymeTextBox.Visible = false;
                Sosa2PrenomTextBox.Visible = false;
                Sosa2NeTextBox.Visible = false;
                Sosa2NeEndroitTextBox.Visible = false;
                Sosa2DeTextBox.Visible = false;
                Sosa2DeEndroitTextBox.Visible = false;
                Sosa23MaTextBox.Visible = false;
                Sosa23MaEndroitTextBox.Visible = false;
                RectangleSosa2.BorderColor = Color.Black;

                Sosa3Label.Visible = false;
                Sosa3PatronymeTextBox.Visible = false;
                Sosa3PrenomTextBox.Visible = false;
                Sosa3NeTextBox.Visible = false;
                Sosa3NeEndroitTextBox.Visible = false;
                Sosa3DeTextBox.Visible = false;
                Sosa3DeEndroitTextBox.Visible = false;
                RectangleSosa3.BorderColor = Color.Black;

                Sosa4Label.Visible = false;
                Sosa4PatronymeTextBox.Visible = false;
                Sosa4PrenomTextBox.Visible = false;
                Sosa4NeTextBox.Visible = false;
                Sosa4NeEndroitTextBox.Visible = false;
                Sosa4DeTextBox.Visible = false;
                Sosa4DeEndroitTextBox.Visible = false;
                Sosa45MaTextBox.Visible = false;
                Sosa45MaLEndroitTextBox.Visible = false;
                RectangleSosa4.BorderColor = Color.Black;

                Sosa5Label.Visible = false;
                Sosa5PatronymeTextBox.Visible = false;
                Sosa5NeTextBox.Visible = false;
                Sosa5NeEndroitTextBox.Visible = false;
                Sosa5DeTextBox.Visible = false;
                Sosa5DeEndroitTextBox.Visible = false;
                RectangleSosa5.BorderColor = Color.Black;

                Sosa6Label.Visible = false;
                Sosa6PatronymeTextBox.Visible = false;
                Sosa6PrenomTextBox.Visible = false;
                Sosa6NeTextBox.Visible = false;
                Sosa6NeEndroitTextBox.Visible = false;
                Sosa6DeTextBox.Visible = false;
                Sosa6DeEndroitTextBox.Visible = false;
                Sosa67MaTextBox.Visible = false;
                Sosa67MaEndroitTextBox.Visible = false;
                RectangleSosa6.BorderColor = Color.Black;

                Sosa7Label.Visible = false;
                Sosa7PatronymeTextBox.Visible = false;
                Sosa7PrenomTextBox.Visible = false;
                Sosa7NeTextBox.Visible = false;
                Sosa7NeEndroitTextBox.Visible = false;
                Sosa7DeTextBox.Visible = false;
                Sosa7DeEndroitTextBox.Visible = false;
                RectangleSosa7.BorderColor = Color.Black;

                GoSosa4Btn.Visible = false;
                GoSosa5Btn.Visible = false;
                GoSosa6Btn.Visible = false;
                GoSosa7Btn.Visible = false;
                this.ResumeLayout();

            }
            else
            {
                // affiche les cases
                sosaCourant = Int32.Parse(ChoixSosaComboBox.Text);
                Sosa1PatronymeTextBox.Visible = true;
                Sosa1PrenomTextBox.Visible = true;
                Sosa1NeTextBox.Visible = true;
                Sosa1NeEndroitTextBox.Visible = true;
                Sosa1DeTextBox.Visible = true;
                Sosa1DeEndroitTextBox.Visible = true;
                if ((sosaCourant > 1) &&  (sosaCourant % 2 == 0) )
                {
                    Sosa1MaTextBox.Visible = true;
                    Sosa1MaEndroitTextBox.Visible = true;
                    sosa1LigneVertical.Visible = true;
                    SosaConjoint1Label.Visible = true;
                    RectangleSosaConjoint1.Visible = true;
                    Sosa1MaEtiquettetBox.Visible = true;
                    Sosa1LieuEtiquettetBox.Visible = true;
                } else
                {
                    Sosa1MaTextBox.Visible = false;
                    Sosa1MaEndroitTextBox.Visible = false;
                    sosa1LigneVertical.Visible = false;
                    SosaConjoint1Label.Visible = false;
                    SosaConjoint1PatronymeTextBox.Visible = false;
                    SosaConjoint1PrenomTextBox.Visible = false;
                    Conjoint1Lbl.Visible = false;
                    RectangleSosaConjoint1.Visible = false;
                    Sosa1MaEtiquettetBox.Visible = false;
                    Sosa1LieuEtiquettetBox.Visible = false;
                    
                }
                Note1.Visible = true;
                Note2.Visible = true;
                NoteDuHautLb.Visible = true;
                NoteDuBasLb.Visible = true;
                RectangleSosa1.BorderColor = Color.Black;

                Sosa2Label.Visible = true;
                Sosa2PatronymeTextBox.Visible = true;
                Sosa2PrenomTextBox.Visible = true;
                Sosa2NeTextBox.Visible = true;
                Sosa2NeEndroitTextBox.Visible = true;
                Sosa2DeTextBox.Visible = true;
                Sosa2DeEndroitTextBox.Visible = true;
                Sosa23MaTextBox.Visible = true;
                Sosa23MaEndroitTextBox.Visible = true;
                RectangleSosa2.BorderColor = Color.Black;

                Sosa3Label.Visible = true;
                Sosa3PatronymeTextBox.Visible = true;
                Sosa3PrenomTextBox.Visible = true;
                Sosa3NeTextBox.Visible = true;
                Sosa3NeEndroitTextBox.Visible = true;
                Sosa3DeTextBox.Visible = true;
                Sosa3DeEndroitTextBox.Visible = true;
                RectangleSosa3.BorderColor = Color.Black;

                Sosa4Label.Visible = true;
                Sosa4PatronymeTextBox.Visible = true;
                Sosa4PrenomTextBox.Visible = true;
                Sosa4NeTextBox.Visible = true;
                Sosa4NeEndroitTextBox.Visible = true;
                Sosa4DeTextBox.Visible = true;
                Sosa4DeEndroitTextBox.Visible = true;
                Sosa45MaTextBox.Visible = true;
                Sosa45MaLEndroitTextBox.Visible = true;
                RectangleSosa4.BorderColor = Color.Black;

                Sosa5Label.Visible = true;
                Sosa5PatronymeTextBox.Visible = true;
                Sosa5PrenomTextBox.Visible = true;
                Sosa5NeTextBox.Visible = true;
                Sosa5NeEndroitTextBox.Visible = true;
                Sosa5DeTextBox.Visible = true;
                Sosa5DeEndroitTextBox.Visible = true;
                RectangleSosa5.BorderColor = Color.Black;

                Sosa6Label.Visible = true;
                Sosa6PatronymeTextBox.Visible = true;
                Sosa6PrenomTextBox.Visible = true;
                Sosa6NeTextBox.Visible = true;
                Sosa6NeEndroitTextBox.Visible = true;
                Sosa6DeTextBox.Visible = true;
                Sosa6DeEndroitTextBox.Visible = true;
                Sosa67MaTextBox.Visible = true;
                Sosa67MaEndroitTextBox.Visible = true;
                RectangleSosa6.BorderColor = Color.Black;

                Sosa7Label.Visible = true;
                Sosa7PatronymeTextBox.Visible = true;
                Sosa7PrenomTextBox.Visible = true;
                Sosa7NeTextBox.Visible = true;
                Sosa7NeEndroitTextBox.Visible = true;
                Sosa7DeTextBox.Visible = true;
                Sosa7DeEndroitTextBox.Visible = true;
                RectangleSosa7.BorderColor = Color.Black;
                                
                GoSosa4Btn.Visible = true;
                GoSosa5Btn.Visible = true;
                GoSosa6Btn.Visible = true;
                GoSosa7Btn.Visible = true;
            }
            if (sosaCourant > 1)
            {
                goSosa1Btn.Visible = true;
                double temp = System.Convert.ToInt32(sosaCourant);
                double SosaPrecedent = Math.Floor(temp / 8);
                goSosa1Btn.Text = Convert.ToString(SosaPrecedent);
            }
            else
            {
                goSosa1Btn.Visible = false;
            }
            if (sosaCourant != 0)
            {
                if (sosaCourant % 2 == 0)
                {
                    GoSosaConjoint1Btn.Visible = true;
                    GoSosaConjoint1Btn.Text = Convert.ToString(sosaCourant + 1);
                }
                else
                {
                    GoSosaConjoint1Btn.Visible = false;
                }
            }
            if (sosaCourant == 0)
            {
                GoSosaConjoint1Btn.Visible = false;
            }

            int Sosa4 = (sosaCourant * 4) * 2;
            int Sosa5 = (sosaCourant * 4 + 1) *2;
            int Sosa6 = (sosaCourant * 4 + 2) *2;
            int Sosa7 = (sosaCourant * 4 + 3) *2;
            if (Sosa4 > 6 && Sosa4 < 128 )
            {
                GoSosa4Btn.Visible = true;
                GoSosa4Btn.Text = Convert.ToString(Sosa4);
            } else
            {
                GoSosa4Btn.Visible = false;
            }
            if (Sosa5 > 5 && Sosa5 < 128)
            {
                GoSosa5Btn.Visible = true;
                GoSosa5Btn.Text = Convert.ToString(Sosa5);
            }
            else
            {
                GoSosa5Btn.Visible = false;
            }
            if (Sosa6 > 6 && Sosa6 < 128)
            {
                GoSosa6Btn.Visible = true;
                GoSosa6Btn.Text = Convert.ToString(Sosa6);
            }
            else
            {
                GoSosa6Btn.Visible = false;
            }
            if (Sosa7 > 6 && Sosa7 < 128)
            {
                GoSosa7Btn.Visible = true;
                GoSosa7Btn.Text = Convert.ToString(Sosa7);
            }
            else
            {
                GoSosa7Btn.Visible = false;
            }
            ChoixSosaComboBox.Focus();
        }
        private void    AvoirDossierrapport()
        {
            FolderBrowserDialog folderBrowserDialog1 = new FolderBrowserDialog
            {
                Description = "Ou enregister les rapports"
            };

            if (folderBrowserDialog1.ShowDialog() == DialogResult.OK)
            {
                DossierPDF = folderBrowserDialog1.SelectedPath;
                DossierPDFToolStripMenuItem.Text = "D&ossier PDF -> " + DossierPDF;
            }
        }
        private string  AssemblerNom(string prenom, string patronyme)
        {
            if (prenom == "" && patronyme =="" ) {
                return "";
            }
            if (prenom == "" && patronyme !="" ) {
            return patronyme;
            }
            if (prenom != "" && patronyme =="" ) {
                return  prenom;
            }

            return prenom + " " + patronyme;
        }
        private void    ChangerCouleurBloc(Color rgb)
        {
            SosaConjoint1Label.BackColor = rgb;
            Sosa2Label.BackColor = rgb;
            Sosa3Label.BackColor = rgb;
            Sosa4Label.BackColor = rgb;
            Sosa5Label.BackColor = rgb;
            Sosa6Label.BackColor = rgb;
            Sosa7Label.BackColor = rgb;

            RectangleGénérationA.BackColor = rgb;
            RectangleGénérationB.BackColor = rgb;
            RectangleGénérationC.BackColor = rgb;
            GenerationAlb.BackColor = rgb;
            GenerationBlb.BackColor = rgb;
            GenerationClb.BackColor = rgb;
            Generation1lb.BackColor = rgb;
            Generation2lb.BackColor = rgb;
            Generation3lb.BackColor = rgb;

            RectangleSosa1.FillColor = rgb;
            RectangleSosaConjoint1.FillColor = rgb;
            SosaConjoint1PatronymeTextBox.BackColor = rgb;
            SosaConjoint1PrenomTextBox.BackColor = rgb;
            Conjoint1Lbl.BackColor = rgb;
            RectangleSosa2.FillColor = rgb;
            RectangleSosa3.FillColor = rgb;
            RectangleSosa4.FillColor = rgb;
            RectangleSosa5.FillColor = rgb;
            RectangleSosa6.FillColor = rgb;
            RectangleSosa7.FillColor = rgb;
          
            Nele1Lbl.BackColor = rgb;
            Nele2Lbl.BackColor = rgb;
            Nele3Lbl.BackColor = rgb;
            Nele4Lbl.BackColor = rgb;
            Nele5Lbl.BackColor = rgb;
            Nele6Lbl.BackColor = rgb;
            Nele7Lbl.BackColor = rgb;

            NeEndroit1Lbl.BackColor = rgb;
            NeEndroit2Lbl.BackColor = rgb;
            NeEndroit3Lbl.BackColor = rgb;
            NeEndroit4Lbl.BackColor = rgb;
            NeEndroit5Lbl.BackColor = rgb;
            NeEndroit6Lbl.BackColor = rgb;
            NeEndroit7Lbl.BackColor = rgb;

            Dele1Lbl.BackColor = rgb;
            Dele2Lbl.BackColor = rgb;
            Dele3Lbl.BackColor = rgb;
            Dele4Lbl.BackColor = rgb;
            Dele5Lbl.BackColor = rgb;
            Dele6Lbl.BackColor = rgb;
            Dele7Lbl.BackColor = rgb;

            DeEndroit1Lbl.BackColor = rgb;
            DeEndroit2Lbl.BackColor = rgb;
            DeEndroit3Lbl.BackColor = rgb;
            DeEndroit4Lbl.BackColor = rgb;
            DeEndroit5Lbl.BackColor = rgb;
            DeEndroit6Lbl.BackColor = rgb;
            DeEndroit7Lbl.BackColor = rgb;
        }
        private void    ChoixChanger()
        {
            if (ChoixSosaComboBox.Text == "" && sosaCourant == 0)
            {
                return;
            }
            
            if (ChoixSosaComboBox.Text != "" && sosaCourant != 0)
            {
                if (Int32.Parse(ChoixSosaComboBox.Text) == sosaCourant)
                {
                    return;
                }
            }
            int[] listePage = new int[] { 0, 1, 8, 9, 10, 11, 12, 13, 14, 15, 64, 65, 66, 67, 68, 69, 70, 71, 72, 73, 74, 75, 76, 77, 78, 79, 80, 81, 82, 83, 84, 85, 86, 87, 88, 89, 90, 91, 92, 93, 94, 95, 96, 97, 98, 99, 100, 101, 102, 103, 104, 105, 106, 107, 108, 109, 110, 111, 112, 113, 114, 115, 116, 117, 118, 119, 120, 121, 122, 123, 124, 125, 126, 127 };

            int Page;
            DateTime Maintenant = DateTime.Today;
            DateLb.Text = Maintenant.ToString("yyyy/MM/dd");

            if (ChoixSosaComboBox.Text == "")
            {
                Page = 0;
                GenerationAlb.Text = "";
                GenerationBlb.Text = "";
                GenerationClb.Text = "";
            }
            else
            {
                try
                {
                    Page = Int32.Parse(ChoixSosaComboBox.Text);
                }
                catch
                {
                    ChoixSosaComboBox.BackColor = Color.Gray;
                    return;
                }
            }
            for (int i = 0; i < 74; i++)
            {
                if (Page == listePage[i])
                {
                    SosaChanger();
                    ChoixSosaComboBox.BackColor = Color.White;

                    return;
                }
                else
                {
                    ChoixSosaComboBox.BackColor = Color.Red;
                }
            }
        }
        private string  ConvertirDate(string date)
        {
            if ( date == "" || date == null)
            {
                return "";
            }
            char[] s = { ' ' };
            char zero = '0';
            string[] d = date.Split(s);
            int l = d.Length;
            if (l == 1 )
            {
                return date;
            }
            if (l == 2)
            {
                if (d[0].ToUpper() == "ABT")
                {
                    return "autour " + d[1];
                }
                if (d[0].ToUpper() == "BEF")
                {
                    return "avant " + d[1];
                }

                if (d[0].ToUpper() == "EST")
                {
                    return "estimé " + d[1];
                }

                d[0] = ConvertirMois(d[0]);
                return d[1] + "-" + d[0];
            }
            if (l == 3 )
            {
                if (d[0].ToUpper() == "ABT")
                {
                    return "autour " + d[2] + ConvertirMois(d[1]);
                }
                if (d[1].ToUpper() == "ABT")
                {
                    return "autour + " + d[2] + "-" + d[1];
                }
                if (d[1].ToUpper() == "BEF")
                {
                    return "autour + " + d[2] + "-" + d[1];
                }
                d[1] = ConvertirMois(d[1]);
                return d[2] + "-" + d[1] + "-" + d[0].PadLeft(2,zero);

            }
            if (l == 4)
            {
                if (d[0].ToUpper() == "BEF" )
                {
                    return "avant " + d[3] + "-" + d[2] + "-" + d[1];
                }
                if (d[0].ToUpper() == "BET" && d[2].ToUpper() == "AND")
                {
                    return "entre " + d[1] + " et " + d[3];
                }
                if (d[0].ToUpper() == "FROM" && d[2].ToUpper() == "TO")
                {
                    return "De " + d[1] + " à " + d[3];
                }

            }

            if (l == 6)
            {
                if (d[0].ToUpper() == "BEF" && d[4].ToUpper() == "AND")
                {
                    return "entre " + d[3] + "-" + d[2] + "-" + d[1] + " et " + d[5];
                }
            }

            if (l == 8)
            {
                if (d[0].ToUpper() == "BEF" && d[4].ToUpper() == "AND")
                {
                    return "entre " + d[3] + "-" + d[2] + "-" + d[1] + " et " + d[7] + "-" + d[6] + "-" + d[5];
                }
            }

            return "";
        }
        private string  ConvertirMois(string mois)
        {

            string m = mois.ToUpper();
            if (m == "JAN")
            {
                m = "01";
            }
            if (m == "FEB")
            {
                m = "02";
            }
            if (m == "MAR")
            {
                m = "03";
            }
            if (m == "APR")
            {
                m = "04";
            }
            if (m == "MAY")
            {
                m = "05";
            }
            if (m == "JUN")
            {
                m = "06";
            }
            if (m == "JUL")
            {
                m = "07";
            }
            if (m == "AUG")
            {
                m = "08";
            }
            if (m == "SEP")
            {
                m = "09";
            }
            if (m == "OCT")
            {
                m = "10";
            }
            if (m == "NOV")
            {
                m = "11";
            }
            if (m == "DEC")
            {
                m = "12";
            }
            return m;

        }
        private void    Classer(int Colonne)
        {
           
        }
        private void    Continuer()
        {
            ChoixPersonne.Visible = false;
            ChoixPersonne.Enabled = false;
            EffacerData();
            FichierCourant = "";
            sosaCourant = 1;
            ChoixSosaComboBox.SelectedIndex = 1;
            Modifier = false;
            string ID = null;
            try
            {

                ID = ChoixLV.SelectedItems[0].SubItems[0].Text;
            } catch
            {
                SystemSounds.Beep.Play();
                MessageBox.Show("Il faut choisir un nom dans la liste.\r\n\r\n", "Problème ?",
                                 MessageBoxButtons.OK,
                                 MessageBoxIcon.Warning);
                return;
            }
            //int ID = Int32.Parse(id);
            liste[1, IDg] = ID;
            liste[1, PATRONYME] = GEDCOM.AvoirPatronyme(ID);
            liste[1, PRENOM] = GEDCOM.AvoirPrenom(ID);
            liste[1, NOMTRI] = liste[1, PATRONYME] + " " + liste[1, PRENOM];
            string dateN = ConvertirDate(GEDCOM.AvoirDateNaissance(ID));
            liste[1, NELE] = dateN;
            liste[1, NELIEU] = GEDCOM.AvoirEndroitNaissance(ID);
            string sex = GEDCOM.AvoirSex(ID);
            string IDFamilleEpoux = GEDCOM.AvoirFamilleEpoux(ID);
            string[] IDListeFamilleEpoux = IDFamilleEpoux.Split(' ');
            string IDConjoint ="";
            if (sex == "M")
            {
                IDConjoint = GEDCOM.AvoirEpouse(IDListeFamilleEpoux[0]);
            }
            if (sex == "F")
            {
                IDConjoint = GEDCOM.AvoirEpoux(IDListeFamilleEpoux[0]);
            }
            liste[1, MALE] = ConvertirDate(GEDCOM.AvoirDateMariage(IDListeFamilleEpoux[0]));
            liste[1, MALIEU] = GEDCOM.AvoirEndroitMariage(IDListeFamilleEpoux[0]);
            liste[1, IDg] = ID.ToString();
            string IDFamilleEnfant = GEDCOM.AvoirFamilleEnfant(ID);
            liste[1, IDFAMILLEENFANT] = GEDCOM.AvoirFamilleEnfant(ID);
            liste[0, PATRONYME] = GEDCOM.AvoirPatronyme(IDConjoint);
            liste[0, PRENOM] = GEDCOM.AvoirPrenom(IDConjoint);
            liste[0, NOMTRI] = "";
            if (liste[0, PATRONYME] != "" && liste[0, PRENOM] != "")
            {
                liste[0, NOMTRI] = liste[0, PATRONYME] + " " + liste[0, PRENOM];
            }
            if (liste[0, PATRONYME] != "" && liste[0, PRENOM] == "")
            {
                liste[0, NOMTRI] = liste[0, PATRONYME] ;
            }
            if (liste[0, PATRONYME] == "" && liste[0, PRENOM] != "")
            {
                liste[0, NOMTRI] = " " + liste[0, PRENOM];
            }

            for (int f = 2; f < 512; f += 2)
            {
                int a = f / 2;
                string ss = liste[f / 2, IDFAMILLEENFANT];

                //string IDFamilleEnfant = GEDCOM.AvoirFamilleEnfant(ID);
                if (ss != "")
                {
                    liste[f, IDg] = GEDCOM.AvoirEpoux(ss);
                    liste[f, IDFAMILLEENFANT] = GEDCOM.AvoirFamilleEnfant(liste[f, IDg]);
                    liste[f + 1, IDg] = GEDCOM.AvoirEpouse(ss);
                    liste[f + 1, IDFAMILLEENFANT] = GEDCOM.AvoirFamilleEnfant(liste[f + 1, IDg]);
                }
            }
            for (int f = 2; f < 512; f += 2)
            {
                //liste[f, IDg] = sosaID[f].ToString();
                ID = liste[f, IDg];
                liste[f, PATRONYME] = GEDCOM.AvoirPatronyme(ID);
                liste[f, PRENOM] = GEDCOM.AvoirPrenom(ID);
                liste[f, NOMTRI] = "";
                if (liste[f, PATRONYME] != "" && liste[f, PRENOM] != "")
                {
                    liste[f, NOMTRI] = liste[f, PATRONYME] + " " + liste[f, PRENOM];
                }
                if (liste[f, PATRONYME] != "" && liste[f, PRENOM] == "")
                {
                    liste[f, NOMTRI] = liste[f, PATRONYME];
                }
                if (liste[f, PATRONYME] == "" && liste[f, PRENOM] != "")
                {
                    liste[f, NOMTRI] = " " + liste[f, PRENOM];
                }
                liste[f, NELE] = ConvertirDate(GEDCOM.AvoirDateNaissance(ID));
                liste[f, NELIEU] = GEDCOM.AvoirEndroitNaissance(ID);
                liste[f, DELE] = ConvertirDate(GEDCOM.AvoirDateDeces(ID));
                liste[f, DELIEU] = GEDCOM.AvoirEndroitDeces(ID);
                liste[f, IDFAMILLEENFANT] = GEDCOM.AvoirFamilleEnfant(liste[f, IDg]);
                string ss = liste[f / 2, IDFAMILLEENFANT];
                liste[f, MALE] = ConvertirDate(GEDCOM.AvoirDateMariage(ss));
                liste[f, MALIEU] = GEDCOM.AvoirEndroitMariage(ss);
                int ff = f + 1;
                ID = liste[ff, IDg];
                liste[ff, PATRONYME] = GEDCOM.AvoirPatronyme(ID);
                liste[ff, PRENOM] = GEDCOM.AvoirPrenom(ID);
                liste[ff, NOMTRI] = "";
                if (liste[ff, PATRONYME] != "" && liste[ff, PRENOM] != "")
                {
                    liste[ff, NOMTRI] = liste[ff, PATRONYME] + " " + liste[ff, PRENOM];
                }
                if (liste[ff, PATRONYME] != "" && liste[ff, PRENOM] == "")
                {
                    liste[ff, NOMTRI] = liste[ff, PATRONYME];
                }
                if (liste[ff, PATRONYME] == "" && liste[ff, PRENOM] != "")
                {
                    liste[ff, NOMTRI] = " " + liste[ff, PRENOM];
                }
                liste[ff, NELE] = ConvertirDate(GEDCOM.AvoirDateNaissance(ID));
                liste[ff, NELIEU] = GEDCOM.AvoirEndroitNaissance(ID);
                liste[ff, DELE] = ConvertirDate(GEDCOM.AvoirDateDeces(ID));
                liste[ff, DELIEU] = GEDCOM.AvoirEndroitDeces(ID);
            }
            PatronymeRecherche.Text = "";
            PrenomRecherche.Text = "";
            RafraichirData();
            AfficherData();
        }
        private bool    DataModifier()
        {
            if (Modifier)
            {
                DialogResult resultat = MessageBox.Show("Enregister le tableau avant ?", "Attention",
                            MessageBoxButtons.YesNoCancel, MessageBoxIcon.Warning);
                if (resultat == DialogResult.Yes)
                {
                    if (FichierCourant == "")
                    {
                         bool Ok = EnregistrerDataSous();
                        if (!Ok)
                        {
                            return false;
                        }
                    }
                    else
                    {
                        EnregistrerData();
                    }
                    if (Modifier)
                    {
                        EffacerData();
                    }

                    return true;
                }
                if (resultat == DialogResult.No)
                {
                    EffacerData();
                    return true; 
                }
                if (resultat == DialogResult.Cancel)
                {
                    return false;
                }
            }
            return true;
        }
        private bool    EnregistrerData()
        {
            int index;
            try
            {
                if (File.Exists(FichierCourant))
                {
                    File.Delete(FichierCourant);
                }
                //Création du fichier Texte
                using (StreamWriter ligne = File.CreateText(FichierCourant))
                {
                    ligne.WriteLine("[ver**]");
                    ligne.WriteLine("Ver   =3.0");
                    for (index = 0; index < 512; index++)
                    {
                        if ((liste[index, PATRONYME] != "" || liste[index, PRENOM] != "" || liste[index, NELE] != "" || liste[index, NELIEU] != "" || liste[index, DELE] != ""
                             || liste[index, DELIEU] != "" || liste[index, MALE] != "" || liste[index, MALIEU] != "" ||
                             liste[index, NOTE1] != "" || liste[index, NOTE2] != "") && liste[index, SOSA] != "0")
                        {
                            ligne.WriteLine("[sosa*]");
                            ligne.WriteLine("No    =" + liste[index, SOSA]);
                            if (liste[index, PATRONYME] != "") ligne.WriteLine("Nom   =" + liste[index, PATRONYME]);
                            if (liste[index, PRENOM] != "") ligne.WriteLine("Prenom=" + liste[index, PRENOM]);
                            if (liste[index, NELE] != "") ligne.WriteLine("NeLe  =" + liste[index, NELE]);
                            if (liste[index, NELIEU] != "") ligne.WriteLine("NeLieu=" + liste[index, NELIEU]);
                            if (liste[index, DELE] != "") ligne.WriteLine("DeLe  =" + liste[index, DELE]);
                            if (liste[index, DELIEU] != "") ligne.WriteLine("DeLieu=" + liste[index, DELIEU]);
                            if (liste[index, MALE] != "") ligne.WriteLine("MaLe  =" + liste[index, MALE]);
                            if (liste[index, MALIEU] != "") ligne.WriteLine("MaLieu=" + liste[index, MALIEU]);
                            if (liste[index, NOTE1].Length > 0)
                            {
                                ligne.WriteLine("NoteH =");
                                ligne.WriteLine(liste[index, NOTE1]);
                                ligne.WriteLine("##FIN##");
                            }
                            if (liste[index, NOTE2].Length > 0)
                            {
                                ligne.WriteLine("NoteB =");
                                ligne.WriteLine(liste[index, NOTE2]);
                                ligne.WriteLine("##FIN##");
                            }
                        }
                    }
                    ligne.WriteLine("[par**]");
                    ligne.WriteLine("Par   =" + PreparerPar.Text);
                    ligne.WriteLine("[Asc**]");
                    ligne.WriteLine("Ascend=" + AscendantDeTb.Text);
                    ligne.WriteLine("[FIN**]");
                    ligne.Close();
                    this.Text = NomPrograme + "   " + FichierCourant;
                    Modifier = false;
                    return true;
                }
            }
            catch (Exception m)
            {
                SystemSounds.Beep.Play();
                MessageBox.Show("Ne peut pas enregister la configuration.\r\n\r\n" + m.Message, "Problème ?",
                                 MessageBoxButtons.OK,
                                 MessageBoxIcon.Warning);
                return false;
            }
        }
        private void    EffacerData()
        {
            //SystemSounds.Beep.Play();
            sosaCourant = 0;
            int[]pageListe = new int[] { 0,1, 8, 9, 10, 11, 12, 13, 14, 15, 64, 65, 66, 67, 68, 69, 70, 71, 72, 73, 74, 75, 76, 77, 78, 79, 80, 81, 82, 83, 84, 85, 86, 87, 88, 89, 90, 91, 92, 93, 94, 95, 96, 97, 98, 99, 100, 101, 102, 103, 104, 105, 106, 107, 108, 109, 110, 111, 112, 113, 114, 115, 116, 117, 118, 119, 120, 121, 122, 123, 124, 125, 126, 127 };
            for (int f = 0; f < 512; f++)
            {
                liste[f, SOSA] = f.ToString();
                liste[f, PAGE] = "";
                liste[f, GENERATION] = "";
                liste[f, TABLEAU] = "";
                liste[f, PATRONYME] = "";
                liste[f, PRENOM] = "";
                liste[f, NOMTRI] = "";
                liste[f, NELE] = "";
                liste[f, NELIEU] = "";
                liste[f, DELE] = "";
                liste[f, DELIEU] = "";
                liste[f, MALE] = "";
                liste[f, MALIEU] = "";
                liste[f, IDg] = "";
                liste[f, IDFAMILLEENFANT] = "";
                liste[f, NOTE1] = "";
                liste[f, NOTE2] = "";
            }
            liste[1, GENERATION] = "1";
            liste[2, GENERATION] = "2";
            liste[3, GENERATION] = "3";
            for (int f = 4; f < 8; f++)
            {
                liste[f, GENERATION] = "3";
            }
            for (int f = 8; f < 16; f++)
            {
                liste[f, GENERATION] = "4";
            }
            for (int f = 16; f < 32; f++)
            {
                liste[f, GENERATION] = "5";
            }
            for (int f = 32; f < 64; f++)
            {
                liste[f, GENERATION] = "6";
            }
            for (int f = 64; f < 128; f++)
            {
                liste[f, GENERATION] = "7";
            }
            for (int f = 128; f < 256; f++)
            {
                liste[f, GENERATION] = "8";
            }
            for (int f = 256; f < 512; f++)
            {
                liste[f, GENERATION] = "9";
            }
            for (int f = 1; f < 8; f++)
            {
                liste[f, TABLEAU] = "1";

            }

            liste[ 8, TABLEAU] = "2";
            liste[ 9, TABLEAU] = "3";
            liste[10, TABLEAU] = "4";
            liste[11, TABLEAU] = "5";
            liste[12, TABLEAU] = "6";
            liste[13, TABLEAU] = "7";
            liste[14, TABLEAU] = "8";
            liste[15, TABLEAU] = "9";
            liste[16, TABLEAU] = "2";
            liste[17, TABLEAU] = "2";
            liste[18, TABLEAU] = "3";
            liste[19, TABLEAU] = "3";
            liste[20, TABLEAU] = "4";
            liste[21, TABLEAU] = "4";
            liste[22, TABLEAU] = "5";
            liste[23, TABLEAU] = "5";
            liste[24, TABLEAU] = "6";
            liste[25, TABLEAU] = "6";
            liste[26, TABLEAU] = "7";
            liste[27, TABLEAU] = "7";
            liste[28, TABLEAU] = "8";
            liste[29, TABLEAU] = "8";
            liste[30, TABLEAU] = "9";
            liste[31, TABLEAU] = "9";

            for (int f = 32; f < 36; f++)
            {
                liste[f, TABLEAU] = "2";
            }
            for (int f = 36; f < 40; f++)
            {
                liste[f, TABLEAU] = "3";
            }
            for (int f = 40; f < 44; f++)
            {
                liste[f, TABLEAU] = "4";
            }
            for (int f = 44; f < 48; f++)
            {
                liste[f, TABLEAU] = "5";
            }
            for (int f = 48; f < 52; f++)
            {
                liste[f, TABLEAU] = "6";
            }
            for (int f = 52; f < 56; f++)
            {
                liste[f, TABLEAU] = "7";
            }
            for (int f = 56; f < 60; f++)
            {
                liste[f, TABLEAU] = "8";
            }
            for (int f = 60; f < 64; f++)
            {
                liste[f, TABLEAU] = "9";
            }
            liste[64, TABLEAU] = "10";
            liste[65, TABLEAU] = "11";
            liste[66, TABLEAU] = "12";
            liste[67, TABLEAU] = "13";
            liste[68, TABLEAU] = "14";
            liste[69, TABLEAU] = "15";
            liste[70, TABLEAU] = "16";
            liste[71, TABLEAU] = "17";
            liste[72, TABLEAU] = "18";
            liste[73, TABLEAU] = "19";
            liste[74, TABLEAU] = "20";
            liste[75, TABLEAU] = "21";
            liste[76, TABLEAU] = "22";
            liste[77, TABLEAU] = "23";
            liste[78, TABLEAU] = "24";
            liste[79, TABLEAU] = "25";
            liste[80, TABLEAU] = "26";
            liste[81, TABLEAU] = "27";
            liste[82, TABLEAU] = "28";
            liste[83, TABLEAU] = "29";
            liste[84, TABLEAU] = "30";
            liste[85, TABLEAU] = "31";
            liste[86, TABLEAU] = "32";
            liste[87, TABLEAU] = "33";
            liste[88, TABLEAU] = "34";
            liste[89, TABLEAU] = "35";
            liste[90, TABLEAU] = "36";
            liste[91, TABLEAU] = "37";
            liste[92, TABLEAU] = "38";
            liste[93, TABLEAU] = "39";
            liste[94, TABLEAU] = "40";
            liste[95, TABLEAU] = "41";
            liste[96, TABLEAU] = "42";
            liste[97, TABLEAU] = "43";
            liste[98, TABLEAU] = "44";
            liste[99, TABLEAU] = "45";
            liste[100, TABLEAU] = "46";
            liste[101, TABLEAU] = "47";
            liste[102, TABLEAU] = "48";
            liste[103, TABLEAU] = "49";
            liste[104, TABLEAU] = "50";
            liste[105, TABLEAU] = "51";
            liste[106, TABLEAU] = "52";
            liste[107, TABLEAU] = "53";
            liste[108, TABLEAU] = "54";
            liste[109, TABLEAU] = "55";
            liste[110, TABLEAU] = "56";
            liste[111, TABLEAU] = "57";
            liste[112, TABLEAU] = "58";
            liste[113, TABLEAU] = "59";
            liste[114, TABLEAU] = "60";
            liste[115, TABLEAU] = "61";
            liste[116, TABLEAU] = "62";
            liste[117, TABLEAU] = "63";
            liste[118, TABLEAU] = "64";
            liste[119, TABLEAU] = "65";
            liste[120, TABLEAU] = "66";
            liste[121, TABLEAU] = "67";
            liste[122, TABLEAU] = "68";
            liste[123, TABLEAU] = "69";
            liste[124, TABLEAU] = "70";
            liste[125, TABLEAU] = "71";
            liste[126, TABLEAU] = "72";
            liste[127, TABLEAU] = "73";
            liste[128, TABLEAU] = "10";
            liste[129, TABLEAU] = "10";
            liste[130, TABLEAU] = "11";
            liste[131, TABLEAU] = "11";
            liste[132, TABLEAU] = "12";
            liste[133, TABLEAU] = "12";
            liste[134, TABLEAU] = "13";
            liste[135, TABLEAU] = "13";
            liste[136, TABLEAU] = "14";
            liste[137, TABLEAU] = "14";
            liste[138, TABLEAU] = "15";
            liste[139, TABLEAU] = "15";
            liste[140, TABLEAU] = "16";
            liste[141, TABLEAU] = "16";
            liste[142, TABLEAU] = "17";
            liste[143, TABLEAU] = "17";
            liste[144, TABLEAU] = "18";
            liste[145, TABLEAU] = "18";
            liste[146, TABLEAU] = "19";
            liste[147, TABLEAU] = "19";
            liste[148, TABLEAU] = "20";
            liste[149, TABLEAU] = "20";
            liste[150, TABLEAU] = "21";
            liste[151, TABLEAU] = "21";
            liste[152, TABLEAU] = "22";
            liste[153, TABLEAU] = "22";
            liste[154, TABLEAU] = "23";
            liste[155, TABLEAU] = "23";
            liste[156, TABLEAU] = "24";
            liste[157, TABLEAU] = "24";
            liste[158, TABLEAU] = "25";
            liste[159, TABLEAU] = "25";
            liste[160, TABLEAU] = "26";
            liste[161, TABLEAU] = "26";
            liste[162, TABLEAU] = "27";
            liste[163, TABLEAU] = "27";
            liste[164, TABLEAU] = "28";
            liste[165, TABLEAU] = "28";
            liste[166, TABLEAU] = "29";
            liste[167, TABLEAU] = "29";
            liste[168, TABLEAU] = "30";
            liste[169, TABLEAU] = "30";
            liste[170, TABLEAU] = "31";
            liste[171, TABLEAU] = "31";
            liste[172, TABLEAU] = "32";
            liste[173, TABLEAU] = "32";
            liste[174, TABLEAU] = "33";
            liste[175, TABLEAU] = "33";
            liste[176, TABLEAU] = "34";
            liste[177, TABLEAU] = "34";
            liste[178, TABLEAU] = "35";
            liste[179, TABLEAU] = "35";
            liste[180, TABLEAU] = "36";
            liste[181, TABLEAU] = "36";
            liste[182, TABLEAU] = "37";
            liste[183, TABLEAU] = "37";
            liste[184, TABLEAU] = "38";
            liste[185, TABLEAU] = "38";
            liste[186, TABLEAU] = "39";
            liste[187, TABLEAU] = "39";
            liste[188, TABLEAU] = "40";
            liste[189, TABLEAU] = "40";
            liste[190, TABLEAU] = "41";
            liste[191, TABLEAU] = "41";
            liste[192, TABLEAU] = "42";
            liste[193, TABLEAU] = "42";
            liste[194, TABLEAU] = "43";
            liste[195, TABLEAU] = "43";
            liste[196, TABLEAU] = "44";
            liste[197, TABLEAU] = "44";
            liste[198, TABLEAU] = "45";
            liste[199, TABLEAU] = "45";
            liste[200, TABLEAU] = "46";
            liste[201, TABLEAU] = "46";
            liste[202, TABLEAU] = "47";
            liste[203, TABLEAU] = "47";
            liste[204, TABLEAU] = "48";
            liste[205, TABLEAU] = "48";
            liste[206, TABLEAU] = "49";
            liste[207, TABLEAU] = "49";
            liste[208, TABLEAU] = "50";
            liste[209, TABLEAU] = "50";
            liste[210, TABLEAU] = "51";
            liste[211, TABLEAU] = "51";
            liste[212, TABLEAU] = "52";
            liste[213, TABLEAU] = "52";
            liste[214, TABLEAU] = "53";
            liste[215, TABLEAU] = "53";
            liste[216, TABLEAU] = "54";
            liste[217, TABLEAU] = "54";
            liste[218, TABLEAU] = "55";
            liste[219, TABLEAU] = "55";
            liste[220, TABLEAU] = "56";
            liste[221, TABLEAU] = "56";
            liste[222, TABLEAU] = "57";
            liste[223, TABLEAU] = "57";
            liste[224, TABLEAU] = "58";
            liste[225, TABLEAU] = "58";
            liste[226, TABLEAU] = "59";
            liste[227, TABLEAU] = "59";
            liste[228, TABLEAU] = "60";
            liste[229, TABLEAU] = "60";
            liste[230, TABLEAU] = "61";
            liste[231, TABLEAU] = "61";
            liste[232, TABLEAU] = "62";
            liste[233, TABLEAU] = "62";
            liste[234, TABLEAU] = "63";
            liste[235, TABLEAU] = "63";
            liste[236, TABLEAU] = "64";
            liste[237, TABLEAU] = "64";
            liste[238, TABLEAU] = "65";
            liste[239, TABLEAU] = "65";
            liste[240, TABLEAU] = "66";
            liste[241, TABLEAU] = "66";
            liste[242, TABLEAU] = "67";
            liste[243, TABLEAU] = "67";
            liste[244, TABLEAU] = "68";
            liste[245, TABLEAU] = "68";
            liste[246, TABLEAU] = "69";
            liste[247, TABLEAU] = "69";
            liste[248, TABLEAU] = "70";
            liste[249, TABLEAU] = "70";
            liste[250, TABLEAU] = "71";
            liste[251, TABLEAU] = "71";
            liste[252, TABLEAU] = "72";
            liste[253, TABLEAU] = "72";
            liste[254, TABLEAU] = "73";
            liste[255, TABLEAU] = "73";


            liste[256, TABLEAU] = "10";
            liste[257, TABLEAU] = "10";
            liste[258, TABLEAU] = "10";
            liste[259, TABLEAU] = "10";
            liste[260, TABLEAU] = "11";
            liste[261, TABLEAU] = "11";
            liste[262, TABLEAU] = "11";
            liste[263, TABLEAU] = "11";
            liste[264, TABLEAU] = "12";
            liste[265, TABLEAU] = "12";
            liste[266, TABLEAU] = "12";
            liste[267, TABLEAU] = "12";
            liste[268, TABLEAU] = "13";
            liste[269, TABLEAU] = "13";
            liste[270, TABLEAU] = "13";
            liste[271, TABLEAU] = "13";
            liste[272, TABLEAU] = "14";
            liste[273, TABLEAU] = "14";
            liste[274, TABLEAU] = "14";
            liste[275, TABLEAU] = "14";
            liste[276, TABLEAU] = "15";
            liste[277, TABLEAU] = "15";
            liste[278, TABLEAU] = "15";
            liste[279, TABLEAU] = "15";
            liste[280, TABLEAU] = "16";
            liste[281, TABLEAU] = "16";
            liste[282, TABLEAU] = "16";
            liste[283, TABLEAU] = "16";
            liste[284, TABLEAU] = "17";
            liste[285, TABLEAU] = "17";
            liste[286, TABLEAU] = "17";
            liste[287, TABLEAU] = "17";
            liste[288, TABLEAU] = "18";
            liste[289, TABLEAU] = "18";
            liste[290, TABLEAU] = "18";
            liste[291, TABLEAU] = "18";
            liste[292, TABLEAU] = "19";
            liste[293, TABLEAU] = "19";
            liste[294, TABLEAU] = "19";
            liste[295, TABLEAU] = "19";
            liste[296, TABLEAU] = "20";
            liste[297, TABLEAU] = "20";
            liste[298, TABLEAU] = "20";
            liste[299, TABLEAU] = "20";
            liste[300, TABLEAU] = "21";
            liste[301, TABLEAU] = "21";
            liste[302, TABLEAU] = "21";
            liste[303, TABLEAU] = "21";
            liste[304, TABLEAU] = "22";
            liste[305, TABLEAU] = "22";
            liste[306, TABLEAU] = "22";
            liste[307, TABLEAU] = "22";
            liste[308, TABLEAU] = "23";
            liste[309, TABLEAU] = "23";
            liste[310, TABLEAU] = "23";
            liste[311, TABLEAU] = "23";
            liste[312, TABLEAU] = "24";
            liste[313, TABLEAU] = "24";
            liste[314, TABLEAU] = "24";
            liste[315, TABLEAU] = "24";
            liste[316, TABLEAU] = "25";
            liste[317, TABLEAU] = "25";
            liste[318, TABLEAU] = "25";
            liste[319, TABLEAU] = "25";
            liste[320, TABLEAU] = "26";
            liste[321, TABLEAU] = "26";
            liste[322, TABLEAU] = "26";
            liste[323, TABLEAU] = "26";
            liste[324, TABLEAU] = "27";
            liste[325, TABLEAU] = "27";
            liste[326, TABLEAU] = "27";
            liste[327, TABLEAU] = "27";
            liste[328, TABLEAU] = "28";
            liste[329, TABLEAU] = "28";
            liste[330, TABLEAU] = "28";
            liste[331, TABLEAU] = "28";
            liste[332, TABLEAU] = "29";
            liste[333, TABLEAU] = "29";
            liste[334, TABLEAU] = "29";
            liste[335, TABLEAU] = "29";
            liste[336, TABLEAU] = "30";
            liste[337, TABLEAU] = "30";
            liste[338, TABLEAU] = "30";
            liste[339, TABLEAU] = "30";
            liste[340, TABLEAU] = "31";
            liste[341, TABLEAU] = "31";
            liste[342, TABLEAU] = "31";
            liste[343, TABLEAU] = "31";
            liste[344, TABLEAU] = "32";
            liste[345, TABLEAU] = "32";
            liste[346, TABLEAU] = "32";
            liste[347, TABLEAU] = "32";
            liste[348, TABLEAU] = "33";
            liste[349, TABLEAU] = "33";
            liste[350, TABLEAU] = "33";
            liste[351, TABLEAU] = "33";
            liste[352, TABLEAU] = "34";
            liste[353, TABLEAU] = "34";
            liste[354, TABLEAU] = "34";
            liste[355, TABLEAU] = "34";
            liste[356, TABLEAU] = "35";
            liste[357, TABLEAU] = "35";
            liste[358, TABLEAU] = "35";
            liste[359, TABLEAU] = "35";
            liste[360, TABLEAU] = "36";
            liste[361, TABLEAU] = "36";
            liste[362, TABLEAU] = "36";
            liste[363, TABLEAU] = "36";
            liste[364, TABLEAU] = "37";
            liste[365, TABLEAU] = "37";
            liste[366, TABLEAU] = "37";
            liste[367, TABLEAU] = "37";
            liste[368, TABLEAU] = "38";
            liste[369, TABLEAU] = "38";
            liste[370, TABLEAU] = "38";
            liste[371, TABLEAU] = "38";
            liste[372, TABLEAU] = "39";
            liste[373, TABLEAU] = "39";
            liste[374, TABLEAU] = "39";
            liste[375, TABLEAU] = "39";
            liste[376, TABLEAU] = "40";
            liste[377, TABLEAU] = "40";
            liste[378, TABLEAU] = "40";
            liste[379, TABLEAU] = "40";
            liste[380, TABLEAU] = "41";
            liste[381, TABLEAU] = "41";
            liste[382, TABLEAU] = "41";
            liste[383, TABLEAU] = "41";
            liste[384, TABLEAU] = "42";
            liste[385, TABLEAU] = "42";
            liste[386, TABLEAU] = "42";
            liste[387, TABLEAU] = "42";
            liste[388, TABLEAU] = "43";
            liste[389, TABLEAU] = "43";
            liste[390, TABLEAU] = "43";
            liste[391, TABLEAU] = "43";
            liste[392, TABLEAU] = "44";
            liste[393, TABLEAU] = "44";
            liste[394, TABLEAU] = "44";
            liste[395, TABLEAU] = "44";
            liste[396, TABLEAU] = "45";
            liste[397, TABLEAU] = "45";
            liste[398, TABLEAU] = "45";
            liste[399, TABLEAU] = "45";
            liste[400, TABLEAU] = "46";
            liste[401, TABLEAU] = "46";
            liste[402, TABLEAU] = "46";
            liste[403, TABLEAU] = "46";
            liste[404, TABLEAU] = "47";
            liste[405, TABLEAU] = "47";
            liste[406, TABLEAU] = "47";
            liste[407, TABLEAU] = "47";
            liste[408, TABLEAU] = "48";
            liste[409, TABLEAU] = "48";
            liste[410, TABLEAU] = "48";
            liste[411, TABLEAU] = "48";
            liste[412, TABLEAU] = "49";
            liste[413, TABLEAU] = "49";
            liste[414, TABLEAU] = "49";
            liste[415, TABLEAU] = "49";
            liste[416, TABLEAU] = "50";
            liste[417, TABLEAU] = "50";
            liste[418, TABLEAU] = "50";
            liste[419, TABLEAU] = "50";
            liste[420, TABLEAU] = "51";
            liste[421, TABLEAU] = "51";
            liste[422, TABLEAU] = "51";
            liste[423, TABLEAU] = "51";
            liste[424, TABLEAU] = "52";
            liste[425, TABLEAU] = "52";
            liste[426, TABLEAU] = "52";
            liste[427, TABLEAU] = "52";
            liste[428, TABLEAU] = "53";
            liste[429, TABLEAU] = "53";
            liste[430, TABLEAU] = "53";
            liste[431, TABLEAU] = "53";
            liste[432, TABLEAU] = "54";
            liste[433, TABLEAU] = "54";
            liste[434, TABLEAU] = "54";
            liste[435, TABLEAU] = "54";
            liste[436, TABLEAU] = "55";
            liste[437, TABLEAU] = "55";
            liste[438, TABLEAU] = "55";
            liste[439, TABLEAU] = "55";
            liste[440, TABLEAU] = "56";
            liste[441, TABLEAU] = "56";
            liste[442, TABLEAU] = "56";
            liste[443, TABLEAU] = "56";
            liste[444, TABLEAU] = "57";
            liste[445, TABLEAU] = "57";
            liste[446, TABLEAU] = "57";
            liste[447, TABLEAU] = "57";
            liste[448, TABLEAU] = "58";
            liste[449, TABLEAU] = "58";
            liste[450, TABLEAU] = "58";
            liste[451, TABLEAU] = "58";
            liste[452, TABLEAU] = "59";
            liste[453, TABLEAU] = "59";
            liste[454, TABLEAU] = "59";
            liste[455, TABLEAU] = "59";
            liste[456, TABLEAU] = "60";
            liste[457, TABLEAU] = "60";
            liste[458, TABLEAU] = "60";
            liste[459, TABLEAU] = "60";
            liste[460, TABLEAU] = "61";
            liste[461, TABLEAU] = "61";
            liste[462, TABLEAU] = "61";
            liste[463, TABLEAU] = "61";
            liste[464, TABLEAU] = "62";
            liste[465, TABLEAU] = "62";
            liste[466, TABLEAU] = "62";
            liste[467, TABLEAU] = "62";
            liste[468, TABLEAU] = "63";
            liste[469, TABLEAU] = "63";
            liste[470, TABLEAU] = "63";
            liste[471, TABLEAU] = "63";
            liste[472, TABLEAU] = "64";
            liste[473, TABLEAU] = "64";
            liste[474, TABLEAU] = "64";
            liste[475, TABLEAU] = "64";
            liste[476, TABLEAU] = "65";
            liste[477, TABLEAU] = "65";
            liste[478, TABLEAU] = "65";
            liste[479, TABLEAU] = "65";
            liste[480, TABLEAU] = "66";
            liste[481, TABLEAU] = "66";
            liste[482, TABLEAU] = "66";
            liste[483, TABLEAU] = "66";
            liste[484, TABLEAU] = "67";
            liste[485, TABLEAU] = "67";
            liste[486, TABLEAU] = "67";
            liste[487, TABLEAU] = "67";
            liste[488, TABLEAU] = "68";
            liste[489, TABLEAU] = "68";
            liste[490, TABLEAU] = "68";
            liste[491, TABLEAU] = "68";
            liste[492, TABLEAU] = "69";
            liste[493, TABLEAU] = "69";
            liste[494, TABLEAU] = "69";
            liste[495, TABLEAU] = "69";
            liste[496, TABLEAU] = "70";
            liste[497, TABLEAU] = "70";
            liste[498, TABLEAU] = "70";
            liste[499, TABLEAU] = "70";
            liste[500, TABLEAU] = "71";
            liste[501, TABLEAU] = "71";
            liste[502, TABLEAU] = "71";
            liste[503, TABLEAU] = "71";
            liste[504, TABLEAU] = "72";
            liste[505, TABLEAU] = "72";
            liste[506, TABLEAU] = "72";
            liste[507, TABLEAU] = "72";
            liste[508, TABLEAU] = "73";
            liste[509, TABLEAU] = "73";
            liste[510, TABLEAU] = "73";
            liste[511, TABLEAU] = "73";
            


            foreach (int s in pageListe)
            {
                liste[s, PAGE] = s.ToString();
                liste[s * 1, PAGE] = s.ToString();
                liste[s * 2, PAGE] = s.ToString();
                liste[s * 2 + 1, PAGE] = s.ToString();
                liste[s * 4, PAGE] = s.ToString();
                liste[s * 4 + 1, PAGE] = s.ToString();
                liste[s * 4 + 2, PAGE] = s.ToString();
                liste[s * 4 + 3, PAGE] = s.ToString();
            }
            this.Text = NomPrograme;
            ChoixSosaComboBox.Text = "";
            AscendantDeTb.Text = "";

            PreparerPar.Text = "";
            Modifier = false;
            int rowLength = liste.Length;



            
            AfficherData();
        }
        private bool    EnregistrerDataSous()
        {
            SaveFileDialog EnregisterDialog = new SaveFileDialog
            {
                Filter = "Fichier|*.tas",
                Title = "Enregister sous ..."
            };
            DialogResult dr = new DialogResult();
            dr = EnregisterDialog.ShowDialog();
            if (dr == DialogResult.Cancel)
            {
                return false;
            }
            if (EnregisterDialog.FileName != "")
            {
                FichierCourant = EnregisterDialog.FileName;
                return EnregistrerData();

            }
            return false;
        }
        /**************************************************************************************************************
            enregistre la grille dans ficier
        **************************************************************************************************************/
        private void    EnregisterGrille()
        {
            try
            {
                string fichier;
                if (Directory.Exists(DossierPDF))
                {
                    fichier = DossierPDF + "\\" + FICHIERGRILLE;
                }
                else
                {
                    fichier = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + "\\" + FICHIERGRILLE;
                }

                
                using (StreamWriter ligne = File.CreateText(fichier))
                    //ligne.WriteLine("SOSA" + " " + "PAGE");

                    for (int f = 0; f < 512; f++)
                    {
                        ligne.WriteLine(
                            "SOSA " + liste[f, SOSA] + " " + 
                            "PAGE " + liste[f, PAGE] + " " + 
                            liste[f, PATRONYME] +  " " + 
                            liste[f, PRENOM] + " " +
                            "MALE " + liste[f, MALE]  +" " +
                            "MALIEU " + liste[f, MALIEU ] + " " 
                            );
                    }
            } catch {}
        }
        /**************************************************************************************************************/
        private void    EnteteTableMatiere(ref PdfDocument document, ref XGraphics gfx, ref PdfPage page)
        {
            XFont font8 = new XFont("Arial", 8, XFontStyle.Bold);
            double x = POUCE * .5;
            double xx = POUCE * 4.5;
            double y = POUCE * 1;

            /**************************************************************************/
            /* Pour le développement marge millieu table matière                      */   
            /**************************************************************************/
            /*
            XPen penG = new XPen(XColor.FromArgb(150, 150, 150),1);
            gfx.DrawRectangle(penG, POUCE * 4,0, POUCE *.5, POUCE * 11);
            */
            // FIN

            XTextFormatter et = new XTextFormatter(gfx);
            XRect rect = new XRect();                                                           
            rect = new XRect(x, y, 170, 10);
            et.DrawString("SOSA", font8, XBrushes.Black, rect, XStringFormats.TopLeft);
            rect = new XRect(x + 100, y, 170, 10);
            et.DrawString("Patronyme", font8, XBrushes.Black, rect, XStringFormats.TopLeft);
            rect = new XRect(x + 219, y, 50, 10);
            et.DrawString("Tableau", font8, XBrushes.Black, rect, XStringFormats.TopLeft);

            rect = new XRect(xx, y, 170, 10);
            et.DrawString("SOSA", font8, XBrushes.Black, rect, XStringFormats.TopLeft);
            rect = new XRect(xx + 100, y, 170, 10);
            et.DrawString("Patronyme", font8, XBrushes.Black, rect, XStringFormats.TopLeft);
            rect = new XRect(xx + 219, y, 50, 10);
            et.DrawString("Tableau", font8, XBrushes.Red, rect, XStringFormats.TopLeft);
        }
        private void    FichierTest()
        {
            string FichierTest = "fichier_test.tas";
            if (File.Exists(FichierTest))
            {
                File.Delete(FichierTest);
            }
            //Création du fichier Texte
            try
            {
                using (StreamWriter ligne = File.CreateText(FichierTest))
                {
                    ligne.WriteLine("[ver**]");
                    ligne.WriteLine("Ver   =3.0");
                    for (int index = 0; index < 512; index++)
                    {
                        ligne.WriteLine("[sosa*]");
                        ligne.WriteLine("No    =" + index);
                        ligne.WriteLine("Nom   =" + "Patronyme" + index.ToString());
                        ligne.WriteLine("Prenom=" + "Prenom" + index.ToString());
                        ligne.WriteLine("NeLe  =" + "Nele " + index.ToString());
                        ligne.WriteLine("NeLieu=" + "NeLieu " + index.ToString());
                        ligne.WriteLine("DeLe  =" + "Dele " + index.ToString());
                        ligne.WriteLine("DeLieu=" + "DeLieu " + index.ToString());
                        ligne.WriteLine("MaLe  =" + "MaLe " + index.ToString());
                        ligne.WriteLine("MaLieu=" + "MaLieu " + index.ToString());
                        ligne.WriteLine("NoteH =");
                        ligne.WriteLine("Note Haute " + index.ToString());
                        ligne.WriteLine("##FIN##");
                        ligne.WriteLine("NoteB =");
                        ligne.WriteLine("NoteBase " + index.ToString());
                        ligne.WriteLine("##FIN##");
                    }
                    ligne.WriteLine("[par**]");
                    ligne.WriteLine("Par   = Daniel");
                    ligne.WriteLine("[Asc**]");
                    ligne.WriteLine("Ascend=" + "Daniel Pambrun");
                    ligne.WriteLine("[FIN**]");
                    ligne.Close();
                    this.Text = FichierTest;
                    Modifier = false;
                }
            }
            catch (Exception m)
            {
                SystemSounds.Beep.Play();
                MessageBox.Show("Ne peut pas enregister le fichier test " + FichierTest + ".\r\n\r\n" + m.Message, "Problème ?",
                                 MessageBoxButtons.OK,
                                 MessageBoxIcon.Warning);
            }
        }
        private void    FlecheDroite(XGraphics gfx, XFont font, double x, double y,  double hauteur, int sosa)
        {
            XImage img = global::TableauAscendant.Properties.Resources.flecheDroite;
            y = y + (hauteur / 2 - 16);
            if (sosa == 0 || (sosa > 7 && sosa < 128)) {
                gfx.DrawImage(img, x, y, 32, 32);
            }
            y = y + 11; 
            XRect rect = new XRect(x, y, 25, 10);
            if (sosa > 7 && sosa < 128)
            {
                gfx.DrawString(liste[sosa, TABLEAU], font, XBrushes.Black, rect, XStringFormats.Center);
            }
            return;
        }
        private void    FlecheGauche(XGraphics gfx, XFont font, double x, double y, double hauteur, int sosa)
        {
            decimal sosaD = sosa;
            XImage img = global::TableauAscendant.Properties.Resources.flecheGauche;
            y = y + (hauteur / 2 - 16);
            if (sosa == 0)
            {
                gfx.DrawImage(img, x, y, 32, 32);
            }
            if (sosa > 1)
            {
                sosaD = Math.Floor(sosaD / 8);
                gfx.DrawImage(img, x, y, 32, 32);
                x = x + 6;
                y = y + 11;
                XRect rect = new XRect(x, y, 25, 10);
                sosa = Convert.ToInt32(sosaD);
                gfx.DrawString(liste[sosa, TABLEAU], font, XBrushes.Black, rect, XStringFormats.Center);
            }
            return;
        }
        private string  DessinerPage(ref PdfDocument document, ref XGraphics gfx, int sosa, bool fleche, bool tous)
        {
            //int inch = 72 // 72 pointCreatePage
            XUnit pouce = XUnit.FromInch(1);
            XPen pen = new XPen(XColor.FromArgb(0, 0, 0),2);
            XPen penG = new XPen(XColor.FromArgb(100, 100, 100), 1);
            XPen penM = new XPen(XColor.FromArgb(0, 0, 255), 1);
            XPen penB = new XPen(XColor.FromArgb(255, 255, 255), 1);
            XFont fontT = new XFont("Arial", 14, XFontStyle.Regular);
            XFont font8 = new XFont("Arial", 8, XFontStyle.Regular);
            XFont font8B = new XFont("Arial", 8, XFontStyle.Bold);
            XBrush gris = new XSolidBrush(XColor.FromArgb(255, 255, 255));
            XBrush CouleurBloc = new XSolidBrush(XColor.FromArgb(RectangleSosa1.FillColor.R, RectangleSosa1.FillColor.G, RectangleSosa1.FillColor.B));

            XTextFormatter tf = new XTextFormatter(gfx);
            double xx = pouce * .5;
            double yy = pouce * .5;
            string str;
            string numeroTableau = "";
            int f; // pour for
            //int numeroPage = 0;
            double y = 0;
            /**************************************************************************/
            //  Pour le développement                                                 */   
            /**************************************************************************/
            /*
            gfx.DrawRectangle(pen, pouce * 0.5, 0.5 * pouce, pouce * 10, pouce * 7.5); //' x1,y1,x2,y2  cadrage de page
            */
            /**************************************************************************/
            /* FIN                                                                    */
            /**************************************************************************/
            double hauteurLigne = pouce * .1; //ok .125;
            double largeurBoite = pouce * 2.05;
            double hauteurBoite = hauteurLigne * 11;
            double hauteurBoiteMini = hauteurLigne * 4;
            double EspaceEntreBoite = pouce * .25;
            double Col1 = pouce * .5;               // Flèche Gauche
            double Col2 = pouce;                    // Boite coté gauche sosa 1
            double Col3 = Col2 + largeurBoite;      // Boite coté droite sosa 1
            double Col4 = Col3 + EspaceEntreBoite;  // Boite coté gauche sosa 2 
            double Col5 = Col4 + largeurBoite;      // Boite coté droite sosa 2 
            double Col6 = Col5 + EspaceEntreBoite;  // Boite coté gauche sosa 4
            double Col7 = Col6 + largeurBoite;      // Boite coté droite sosa 4
            double Col8 = Col7 + EspaceEntreBoite;  // Boite coté gauche sosa 8
            double Col9 = Col8 + largeurBoite;      // Boite coté droite sosa 8
            double Col10 = Col9 + 5;                // Flèche Droite

            
            double positionLieu = .16 * pouce; // position Lieu par rapport date mariage = .16 * pouce; // par rapport date mariage
            // position des ligne au 1/4 pouce
            double[] Ligne = new double[100];
            for (f = 0; f < 100; f++)
            {
                double l = hauteurLigne * f;
                Ligne[f] = l;
            }
            /**************************************************************************/
            /* Pour le développement dessine colonnes                                 */
            /**************************************************************************/
            /*
            XPen penLigne = new XPen(XColor.FromArgb(200, 200, 255), 0.5);
            
            for (f = 0; f < 100; f++)
            {
                gfx.DrawString("V" + f, font8, XBrushes.Black, 0, Ligne[f]);
                gfx.DrawLine(penLigne, 0, Ligne[f], pouce * 11, Ligne[f]);
            }
            gfx.DrawString("<1", font8, XBrushes.Black, Col1, y + 10);
            gfx.DrawLine(penLigne, Col1, 0, Col1, pouce * 8.5);
            gfx.DrawString("2", font8, XBrushes.Black, Col2, y + 10);
            gfx.DrawLine(penLigne, Col2, 0, Col2, pouce * 8.5);
            gfx.DrawString("<3", font8, XBrushes.Black, Col3, y + 10);
            gfx.DrawLine(penLigne, Col3, 0, Col3, pouce * 8.5);
            gfx.DrawString("<4", font8, XBrushes.Black, Col4, y + 10);
            gfx.DrawLine(penLigne, Col4, 0, Col4, pouce * 8.5);
            gfx.DrawString("<5", font8, XBrushes.Black, Col5, y + 10);
            gfx.DrawLine(penLigne, Col5, 0, Col5, pouce * 8.5);
            gfx.DrawString("<6", font8, XBrushes.Black, Col6, y + 10);
            gfx.DrawLine(penLigne, Col6, 0, Col6, pouce * 8.5);
            gfx.DrawString("<7", font8, XBrushes.Black, Col7, y + 10);
            gfx.DrawLine(penLigne, Col7, 0, Col7, pouce * 8.5);
            gfx.DrawString("<8", font8, XBrushes.Black, Col8, y + 10);
            gfx.DrawLine(penLigne, Col8, 0, Col8, pouce * 8.5);
            gfx.DrawString("<9", font8, XBrushes.Black, Col9, y + 10);
            gfx.DrawLine(penLigne, Col9, 0, Col9, pouce * 8.5);
            gfx.DrawString("<10", font8, XBrushes.Black, Col10, y + 20);
            gfx.DrawLine(penLigne, Col10, 0, Col10, pouce * 8.5);
            */
            /**************************************************************************/
            /* FIN                                                                    */
            /**************************************************************************/

            /**************************************************************************/
            // Position des boites
            double[,] positionBoite = new double[16, 2]; // en pouce
            {
                // boite 1
                positionBoite[1, 0] = Col2 + 2;
                positionBoite[1, 1] = Ligne[39];
                // boite 2
                positionBoite[2, 0] = Col4 + 2;
                positionBoite[2, 1] = Ligne[24];
                // boite 3
                positionBoite[3, 0] = Col4 + 2;
                positionBoite[3, 1] = Ligne[56];
                // boite 4
                positionBoite[4, 0] = Col6 + 2;
                positionBoite[4, 1] = Ligne[16];
                // boite 5
                positionBoite[5, 0] = Col6 + 2;
                positionBoite[5, 1] = Ligne[32];
                // boite 6
                positionBoite[6, 0] = Col6 + 2;
                positionBoite[6, 1] = Ligne[48];
                // boite 7
                positionBoite[7, 0] = Col6 + 2;
                positionBoite[7, 1] = Ligne[64];
                // boite 8
                positionBoite[8, 0] = Col8 + 2;
                positionBoite[8, 1] = Ligne[15];
                // boite 9
                positionBoite[9, 0] = Col8 + 2;
                positionBoite[9, 1] = Ligne[24];
                // boite 10
                positionBoite[10, 0] = Col8 + 2;
                positionBoite[10, 1] = Ligne[31];
                // boite 11
                positionBoite[11, 0] = Col8 + 2;
                positionBoite[11, 1] = Ligne[40];
                // boite 12
                positionBoite[12, 0] = Col8 + 2;
                positionBoite[12, 1] = Ligne[47];
                // boite 13
                positionBoite[13, 0] = Col8 + 2;
                positionBoite[13, 1] = Ligne[56];
                // boite 14
                positionBoite[14, 0] = Col8 + 2;
                positionBoite[14, 1] = Ligne[63];
                // boite 15
                positionBoite[15, 0] = Col8 + 2;
                positionBoite[15, 1] = Ligne[72];
            }
            
            // position mariage
            double[,] positionMariagexx = new double[7, 2]; // en pouce
            int[] sosaIndex = new int[16];
            sosaIndex[1] = sosa;
            sosaIndex[2] = sosa * 2;
            sosaIndex[3] = sosa * 2 + 1;
            sosaIndex[4] = sosa * 4;
            sosaIndex[5] = sosa * 4 + 1;
            sosaIndex[6] = sosa * 4 + 2;
            sosaIndex[7] = sosa * 4 + 3;
            sosaIndex[8] = sosa * 8;
            sosaIndex[9] = sosa * 8 + 1;
            sosaIndex[10] = sosa * 8 + 2;
            sosaIndex[11] = sosa * 8 + 3;
            sosaIndex[12] = sosa * 8 + 4;
            sosaIndex[13] = sosa * 8 + 5;
            sosaIndex[14] = sosa * 8 + 6;
            sosaIndex[15] = sosa * 8 + 7;
            string[] sosaBoite = new string[16];
            for (f = 1; f < 16; f++)
            {
                sosaBoite[f] = sosaIndex[f].ToString();
            }
            y = pouce;
            double HauteurGeneration = pouce * .25;
            double Rond = 10;

            // haut de page
            str = "Tableau ascendant de ";
            XSize textLargeur = gfx.MeasureString(str, fontT);
            gfx.DrawString(str, fontT, XBrushes.Black, Col1, POUCE * .75);
            textLargeur = gfx.MeasureString(str, fontT);
            if (AscendantDeTb.Text != "")
            {
                str = AscendantDeTb.Text;
                gfx.DrawString(str, fontT, XBrushes.Black, Col1 + textLargeur.Width + 5, POUCE * .75);
            }
            else
            {
                str = "____________________________________";
                gfx.DrawString(str, fontT, XBrushes.Black, Col1 + textLargeur.Width + 5, POUCE * .75);
            }
            
            // dessine génération
            {
                //génération 1
                XRect g = new XRect(Col2, 50, largeurBoite, 20);
                gfx.DrawRoundedRectangle(pen, CouleurBloc, Col2, Ligne[10], largeurBoite, HauteurGeneration, Rond, Rond);
                str = "Génération " + liste[sosa, GENERATION];
                textLargeur = gfx.MeasureString(str, font8);
                if (liste[sosa, GENERATION] == "")
                {
                    gfx.DrawLine(penG, Col2 + (largeurBoite / 2) + (textLargeur.Width / 2) +2, Ligne[10] + 12, Col2 + (largeurBoite / 2) + (textLargeur.Width / 2) + 20, Ligne[10] + 12);
                }
                gfx.DrawString(str, font8, XBrushes.Black, Col2 + (largeurBoite / 2) - textLargeur.Width / 2, Ligne[10] + 12);

                //génération 2
                g = new XRect(Col4, 50, largeurBoite, 20);
                gfx.DrawRoundedRectangle(pen, CouleurBloc, Col4, Ligne[10], largeurBoite, HauteurGeneration, Rond, Rond);
                str = "Génération " + liste[sosa * 2, GENERATION];
                textLargeur = gfx.MeasureString(str, font8);
                if (liste[sosa * 2, GENERATION] == "")
                {
                    gfx.DrawLine(penG, Col4 + (largeurBoite / 2) + (textLargeur.Width / 2) + 2, Ligne[10] + 12, Col4 + (largeurBoite / 2) + (textLargeur.Width / 2) + 20, Ligne[10] + 12);
                }
                gfx.DrawString(str, font8, XBrushes.Black, Col4 + (largeurBoite / 2) - textLargeur.Width / 2, Ligne[10] + 12);

                //génération 3
                g = new XRect(Col6, 50, largeurBoite, 20);
                gfx.DrawRoundedRectangle(pen, CouleurBloc, Col6, Ligne[10], largeurBoite, HauteurGeneration, Rond, Rond);
                str = "Génération " + liste[sosa * 4, GENERATION];
                textLargeur = gfx.MeasureString(str, font8);
                if (liste[sosa * 4, GENERATION] == "")
                {
                    gfx.DrawLine(penG, Col6 + (largeurBoite / 2) + (textLargeur.Width / 2) + 2, Ligne[10] + 12, Col6 + (largeurBoite / 2) + (textLargeur.Width / 2) + 20, Ligne[10] + 12);
                }
                gfx.DrawString(str, font8, XBrushes.Black, Col6 + (largeurBoite / 2) - textLargeur.Width / 2, Ligne[10] + 12);

                //génération 4
                int s = sosa * 8;
                if (s < 512)
                {
                    g = new XRect(Col8, 50, largeurBoite, 20);
                    gfx.DrawRoundedRectangle(pen, CouleurBloc, Col8, Ligne[10], largeurBoite, HauteurGeneration, Rond, Rond);
                    str = "Génération " + liste[s, GENERATION];
                    textLargeur = gfx.MeasureString(str, font8);
                    if (liste[s, GENERATION] == "")
                    {
                        gfx.DrawLine(penG, Col8 + (largeurBoite / 2) + (textLargeur.Width / 2) + 2, y + 12, Col8 + (largeurBoite / 2) + (textLargeur.Width / 2) + 20, y + 12);
                    }
                    gfx.DrawString(str, font8, XBrushes.Black, Col8 + (largeurBoite / 2) - textLargeur.Width / 2, y + 12);
                }
            }
            
            // dessine boite
            {
                gfx.DrawRoundedRectangle(pen, CouleurBloc, Col2, positionBoite[1, 1], largeurBoite, hauteurBoite, 10, 10); // Boite sosa 1
                if(sosa != 1) {
                    gfx.DrawRoundedRectangle(pen, CouleurBloc, Col2, positionBoite[1, 1] + hauteurLigne* 16, largeurBoite, hauteurBoiteMini, 10, 10); //  Boite sosa 1 conjoint
                }
                gfx.DrawRoundedRectangle(pen, CouleurBloc, Col4, positionBoite[2, 1], largeurBoite, hauteurBoite, 10, 10); // Boite sosa 2
                gfx.DrawRoundedRectangle(pen, CouleurBloc, Col4, positionBoite[3, 1], largeurBoite, hauteurBoite, 10, 10); // Boite sosa 3
                gfx.DrawRoundedRectangle(pen, CouleurBloc, Col6, positionBoite[4, 1], largeurBoite, hauteurBoite, 10, 10); // Boite sosa 4
                gfx.DrawRoundedRectangle(pen, CouleurBloc, Col6, positionBoite[5, 1], largeurBoite, hauteurBoite, 10, 10); // Boite sosa 5
                gfx.DrawRoundedRectangle(pen, CouleurBloc, Col6, positionBoite[6, 1], largeurBoite, hauteurBoite, 10, 10); // Boite sosa 6
                gfx.DrawRoundedRectangle(pen, CouleurBloc, Col6, positionBoite[7, 1], largeurBoite, hauteurBoite, 10, 10); // Boite sosa 7
                int s = sosa * 8;
                if (s < 512)
                {
                    gfx.DrawRoundedRectangle(pen, CouleurBloc, Col8, positionBoite[ 8, 1], largeurBoite, hauteurBoiteMini, 10, 10); // Boite sosa 8
                    gfx.DrawRoundedRectangle(pen, CouleurBloc, Col8, positionBoite[ 9, 1], largeurBoite, hauteurBoiteMini, 10, 10); // Boite sosa 9
                    gfx.DrawRoundedRectangle(pen, CouleurBloc, Col8, positionBoite[10, 1], largeurBoite, hauteurBoiteMini, 10, 10); // Boite sosa 10
                    gfx.DrawRoundedRectangle(pen, CouleurBloc, Col8, positionBoite[11, 1], largeurBoite, hauteurBoiteMini, 10, 10); // Boite sosa 11
                    gfx.DrawRoundedRectangle(pen, CouleurBloc, Col8, positionBoite[12, 1], largeurBoite, hauteurBoiteMini, 10, 10); // Boite sosa 12
                    gfx.DrawRoundedRectangle(pen, CouleurBloc, Col8, positionBoite[13, 1], largeurBoite, hauteurBoiteMini, 10, 10); // Boite sosa 13
                    gfx.DrawRoundedRectangle(pen, CouleurBloc, Col8, positionBoite[14, 1], largeurBoite, hauteurBoiteMini, 10, 10); // Boite sosa 14
                    gfx.DrawRoundedRectangle(pen, CouleurBloc, Col8, positionBoite[15, 1], largeurBoite, hauteurBoiteMini, 10, 10); // Boite sosa 15
                }
            }
            // dessine ligne entre les boites
            { 
                gfx.DrawLine(pen, Col3, positionBoite[1, 1] + hauteurBoite / 2, Col4 + 8, positionBoite[1, 1] + hauteurBoite / 2); // Horizontal 1
                gfx.DrawLine(pen, Col5, positionBoite[2, 1] + hauteurBoite / 2, Col6 + 8, positionBoite[2, 1] + hauteurBoite / 2); // Horizontal 2
                gfx.DrawLine(pen, Col5, positionBoite[3, 1] + hauteurBoite / 2, Col6 + 8, positionBoite[3, 1] + hauteurBoite / 2); // Horizontal 3

                int s = sosa * 8;
                if (s < 512)
                {
                    gfx.DrawLine(pen, Col7, positionBoite[4, 1] + hauteurBoite / 2, Col8 + 8, positionBoite[4, 1] + hauteurBoite / 2); // Horizontal 4
                    gfx.DrawLine(pen, Col7, positionBoite[5, 1] + hauteurBoite / 2, Col8 + 8, positionBoite[5, 1] + hauteurBoite / 2); // Horizontal 5
                    gfx.DrawLine(pen, Col7, positionBoite[6, 1] + hauteurBoite / 2, Col8 + 8, positionBoite[6, 1] + hauteurBoite / 2); // Horizontal 6
                    gfx.DrawLine(pen, Col7, positionBoite[7, 1] + hauteurBoite / 2, Col8 + 8, positionBoite[7, 1] + hauteurBoite / 2); // Horizontal 7
                }
                if(sosa != 1) {
                    gfx.DrawLine(pen, Col2 + 8, positionBoite[1, 1] + hauteurBoite, Col2 + 8, positionBoite[1, 1] + hauteurBoite + hauteurLigne * 5); // vertical sosa 1 conjoint
                }
                gfx.DrawLine(pen, Col4 + 8, positionBoite[2, 1] + hauteurBoite, Col4 + 8, positionBoite[3, 1]); // vertical 2 3
                gfx.DrawLine(pen, Col6 + 8, positionBoite[4, 1] + hauteurBoite, Col6 + 8, positionBoite[5, 1]); // vertical 4 5
                gfx.DrawLine(pen, Col6 + 8, positionBoite[6, 1] + hauteurBoite, Col6 + 8, positionBoite[7, 1]); // vertical 6 7

                if (s < 512)
                {
                    gfx.DrawLine(pen, Col8 + 8, positionBoite[8, 1] + hauteurBoiteMini, Col8 + 8, positionBoite[9, 1]); // vertical 8 9
                    gfx.DrawLine(pen, Col8 + 8, positionBoite[10, 1] + hauteurBoiteMini, Col8 + 8, positionBoite[11, 1]); // vertical 10 11
                    gfx.DrawLine(pen, Col8 + 8, positionBoite[12, 1] + hauteurBoiteMini, Col8 + 8, positionBoite[13, 1]); // vertical 12 13
                    gfx.DrawLine(pen, Col8 + 8, positionBoite[14, 1] + hauteurBoiteMini, Col8 + 8, positionBoite[15, 1]); // vertical 14 15
                }
            }
            //
            tf.Alignment = XParagraphAlignment.Right;
            int RetraitSosa = 20;
            XRect rect = new XRect();
            
            // Dessiner les informations des boites
            // sosa conjoint

            if (sosa > 1)
            {
                int sosaConjoint;
                if (sosa % 2 == 0)
                {
                    sosaConjoint = sosa + 1;
                }
                else
                {
                    sosaConjoint = sosa - 1;
                }
                rect = new XRect(positionBoite[1, 0] - RetraitSosa, positionBoite[1, 1] + hauteurLigne * 16, 15, 10);
                if (sosa < 2)
                {
                    gfx.DrawLine(penG, positionBoite[1, 0] - RetraitSosa + 3, positionBoite[1, 1] + 10, positionBoite[1, 0] - RetraitSosa + 11, positionBoite[1, 1] + 10);
                }
                else
                    tf.DrawString(sosaConjoint.ToString(), font8B, XBrushes.Black, rect, XStringFormats.TopLeft);

            }

                // sosa 1 à 7
                for (f = 1; f < 8; f++)
            {
                rect = new XRect(positionBoite[f, 0] - RetraitSosa, positionBoite[f, 1], 15, 10);
                if ( sosa == 0 )
                {
                    gfx.DrawLine(penG, positionBoite[f, 0] - RetraitSosa + 3, positionBoite[f, 1] + 10, positionBoite[f, 0] - RetraitSosa + 11, positionBoite[f, 1] + 10);
                }
                else
                {
                    tf.DrawString(sosaBoite[f], font8B, XBrushes.Black, rect, XStringFormats.TopLeft);
                }
            
                gfx.DrawString("N", font8B, XBrushes.Black, positionBoite[f, 0], positionBoite[f, 1] + hauteurLigne * 4);
                gfx.DrawString("L", font8B, XBrushes.Black, positionBoite[f, 0], positionBoite[f, 1] + hauteurLigne * 6);
                gfx.DrawString("D", font8B, XBrushes.Black, positionBoite[f, 0], positionBoite[f, 1] + hauteurLigne * 8);
                gfx.DrawString("L", font8B, XBrushes.Black, positionBoite[f, 0], positionBoite[f, 1] + hauteurLigne * 10);
            }
            // sosa 8 à 15
            for (f = 8; f < 16; f++)
            {
                int s = sosa * 8;
                if (s < 512)
                {
                    rect = new XRect(positionBoite[f, 0] - RetraitSosa, positionBoite[f, 1], 15, 10);
                    if (sosa == 0)
                    {
                        gfx.DrawLine(penG, positionBoite[f, 0] - RetraitSosa + 3, positionBoite[f, 1] + 10, positionBoite[f, 0] - RetraitSosa + 11,positionBoite[f, 1] + 10);
                    }
                    else
                    {
                        tf.DrawString(sosaBoite[f], font8B, XBrushes.Black, rect, XStringFormats.TopLeft);
                    }
                }
            }

            if (sosa != 1) {
                gfx.DrawString("M", font8B, XBrushes.Black, Col2 + 10, positionBoite[1, 1] + (hauteurLigne * 13), XStringFormats.Default);  // sosa 1 conjoint
                gfx.DrawString("L", font8B, XBrushes.Black, Col2 + 10, positionBoite[1, 1] + (hauteurLigne * 15), XStringFormats.Default);  // sosa 1 conjoint
            }
            gfx.DrawString("M", font8B, XBrushes.Black, Col4 + 10, positionBoite[2, 1] + hauteurLigne * 20, XStringFormats.Default);      // sosa 02-03
            gfx.DrawString("L", font8B, XBrushes.Black, Col4 + 10, positionBoite[2, 1] + hauteurLigne * 22, XStringFormats.Default);      // sosa 02-03
            gfx.DrawString("M", font8B, XBrushes.Black, Col6 + 10, positionBoite[4, 1] + hauteurLigne * 13, XStringFormats.Default);      // sosa 04-05
            gfx.DrawString("L", font8B, XBrushes.Black, Col6 + 10, positionBoite[4, 1] + hauteurLigne * 15, XStringFormats.Default);      // sosa 04-05
            gfx.DrawString("M", font8B, XBrushes.Black, Col6 + 10, positionBoite[6, 1] + hauteurLigne * 13, XStringFormats.Default);      // sosa 06-07
            gfx.DrawString("L", font8B, XBrushes.Black, Col6 + 10, positionBoite[6, 1] + hauteurLigne * 15, XStringFormats.Default);      // sosa 06-07  

            if (sosa * 8 < 512)
            {
                gfx.DrawString("M", font8B, XBrushes.Black, Col8 + 10, positionBoite[ 8, 1] + hauteurLigne * 6, XStringFormats.Default);  // sosa 08-09
                gfx.DrawString("L", font8B, XBrushes.Black, Col8 + 10, positionBoite[ 8, 1] + hauteurLigne * 8, XStringFormats.Default);  // sosa 08-09
                gfx.DrawString("M", font8B, XBrushes.Black, Col8 + 10, positionBoite[10, 1] + hauteurLigne * 6, XStringFormats.Default);  // sosa 10-11
                gfx.DrawString("L", font8B, XBrushes.Black, Col8 + 10, positionBoite[10, 1] + hauteurLigne * 8, XStringFormats.Default);  // sosa 10-11
                gfx.DrawString("M", font8B, XBrushes.Black, Col8 + 10, positionBoite[12, 1] + hauteurLigne * 6, XStringFormats.Default);  // sosa 12-13
                gfx.DrawString("L", font8B, XBrushes.Black, Col8 + 10, positionBoite[12, 1] + hauteurLigne * 8, XStringFormats.Default);  // sosa 12-13
                gfx.DrawString("M", font8B, XBrushes.Black, Col8 + 10, positionBoite[14, 1] + hauteurLigne * 6, XStringFormats.Default);  // sosa 14-15
                gfx.DrawString("L", font8B, XBrushes.Black, Col8 + 10, positionBoite[14, 1] + hauteurLigne * 8, XStringFormats.Default);  // sosa 14-15
            }
            //}

            int xInfo = 7;
            if (sosa == 0)
            {
                int largeurLigne = 135; // 
                for (f = 1; f < 8; f++)
                {
                    gfx.DrawLine(penG, positionBoite[f, 0], positionBoite[f, 1] + hauteurLigne * 1, positionBoite[f, 0] + 142, positionBoite[f, 1] + hauteurLigne * 1);
                    gfx.DrawLine(penG, positionBoite[f, 0] + xInfo, positionBoite[f, 1] + hauteurLigne * 4, positionBoite[f, 0] + xInfo + largeurLigne, positionBoite[f, 1] + hauteurLigne * 4);
                    gfx.DrawLine(penG, positionBoite[f, 0] + xInfo, positionBoite[f, 1] + hauteurLigne * 6, positionBoite[f, 0] + xInfo + largeurLigne, positionBoite[f, 1] + hauteurLigne * 6);
                    gfx.DrawLine(penG, positionBoite[f, 0] + xInfo, positionBoite[f, 1] + hauteurLigne * 8, positionBoite[f, 0] + xInfo + largeurLigne, positionBoite[f, 1] + hauteurLigne * 8);
                    gfx.DrawLine(penG, positionBoite[f, 0] + xInfo, positionBoite[f, 1] + hauteurLigne * 10, positionBoite[f, 0] + xInfo + largeurLigne, positionBoite[f, 1] + hauteurLigne * 10);
                }
                int p = 18;
                int l = 140;
                gfx.DrawLine(penG, Col4 + p, Ligne[44], Col4 + l, Ligne[44]);
                gfx.DrawLine(penG, Col4 + p, Ligne[46], Col4 + l, Ligne[46]);

                gfx.DrawLine(penG, Col6 + p, Ligne[29], Col6 + l, Ligne[29]);
                gfx.DrawLine(penG, Col6 + p, Ligne[31], Col6 + l, Ligne[31]);
                gfx.DrawLine(penG, Col6 + p, Ligne[61], Col6 + l, Ligne[61]);
                gfx.DrawLine(penG, Col6 + p, Ligne[63], Col6 + l, Ligne[63]);

                gfx.DrawLine(penG, Col8 + p, Ligne[21], Col8 + l, Ligne[21]);
                gfx.DrawLine(penG, Col8 + p, Ligne[23], Col8 + l, Ligne[23]);
                gfx.DrawLine(penG, Col8 + p, Ligne[37], Col8 + l, Ligne[37]);
                gfx.DrawLine(penG, Col8 + p, Ligne[39], Col8 + l, Ligne[39]);
                gfx.DrawLine(penG, Col8 + p, Ligne[53], Col8 + l, Ligne[53]);
                gfx.DrawLine(penG, Col8 + p, Ligne[55], Col8 + l, Ligne[55]);
                gfx.DrawLine(penG, Col8 + p, Ligne[69], Col8 + l, Ligne[69]);
                gfx.DrawLine(penG, Col8 + p, Ligne[71], Col8 + l, Ligne[71]);

            }
            else
            {
                //info des boites
                int largeurLigne = 142; // 
                {
                    if (sosa != 1){
                        int sosaConjoint;
                        if (sosa%2 == 0 ) {
                            sosaConjoint = sosa + 1;
                        } else {
                            sosaConjoint = sosa - 1;
                        }
                        string nom = AssemblerNom(liste[sosaConjoint, PRENOM], liste[sosaConjoint, PATRONYME]);
                        if (nom == "" )
                        {
                            gfx.DrawLine(penG, positionBoite[1, 0], positionBoite[1, 1] + hauteurLigne * 19, Col2 + 142, positionBoite[1, 1] + hauteurLigne * 19);
                        }
                        string rt = RacoucirNom(nom, ref gfx);
                        gfx.DrawString(rt, font8B, XBrushes.Black, Col2 + 2, positionBoite[1, 1] + hauteurLigne * 19, XStringFormats.Default);

                    }
                    for (f = 1; f < 8; f++)
                    {

                        // Nom
                        string nom = AssemblerNom(liste[sosaIndex[f], PRENOM], liste[sosaIndex[f], PATRONYME]);
                        if (nom == "")
                        {
                            gfx.DrawLine(penG, positionBoite[f, 0], positionBoite[f, 1] + hauteurLigne * 2, positionBoite[f, 0] + 142, positionBoite[f, 1] + hauteurLigne * 2);
                        }
                        string rt = RacoucirNom(nom, ref gfx);
                        gfx.DrawString(rt, font8B, XBrushes.Black, positionBoite[f, 0], positionBoite[f, 1] + hauteurLigne * 2, XStringFormats.Default);
                        // Né le 
                        if (liste[sosaIndex[f], NELE] == "")
                        {
                            gfx.DrawLine(penG, positionBoite[f, 0] + 9, positionBoite[f, 1] + hauteurLigne * 4, positionBoite[f, 0] + largeurLigne, positionBoite[f, 1] + hauteurLigne * 4);
                        }
                        rt = RacoucirTexte(liste[sosaIndex[f], NELE], ref gfx);
                        gfx.DrawString(rt, font8, XBrushes.Black, positionBoite[f, 0] + xInfo, positionBoite[f, 1] + hauteurLigne * 4, XStringFormats.Default);
                        // Né endroit
                        if (liste[sosaIndex[f], NELIEU] == "")
                        {
                            gfx.DrawLine(penG, positionBoite[f, 0] + 9, positionBoite[f, 1] + hauteurLigne * 6, positionBoite[f, 0] + largeurLigne, positionBoite[f, 1] + hauteurLigne * 6);
                        }
                        rt = RacoucirTexte(liste[sosaIndex[f], NELIEU], ref gfx);
                        gfx.DrawString(rt, font8, XBrushes.Black, positionBoite[f, 0] + xInfo, positionBoite[f, 1] + hauteurLigne * 6, XStringFormats.Default);
                        // Décédé le 
                        if (liste[sosaIndex[f], DELE] == "")
                        {
                            gfx.DrawLine(penG, positionBoite[f, 0] + 9, positionBoite[f, 1] + hauteurLigne * 8, positionBoite[f, 0] + largeurLigne, positionBoite[f, 1] + hauteurLigne * 8);
                        }
                        rt = RacoucirTexte(liste[sosaIndex[f], DELE], ref gfx);
                        gfx.DrawString(rt, font8, XBrushes.Black, positionBoite[f, 0] + xInfo, positionBoite[f, 1] + hauteurLigne * 8, XStringFormats.Default);
                        // Décédé endroit
                        if (liste[sosaIndex[f], DELIEU] == "")
                        {
                            gfx.DrawLine(penG, positionBoite[f, 0] + 9, positionBoite[f, 1] + hauteurLigne * 10, positionBoite[f, 0] + largeurLigne, positionBoite[f, 1] + hauteurLigne * 10);
                        }
                        rt = RacoucirTexte(liste[sosaIndex[f], DELIEU], ref gfx);
                        gfx.DrawString(rt, font8, XBrushes.Black, positionBoite[f, 0] + xInfo, positionBoite[f, 1] + hauteurLigne * 10, XStringFormats.Default);
                    }
                    for (f = 8; f < 16; f++)
                    {
                        if (sosaIndex[f] < 512)
                        {
                            string nom = AssemblerNom(liste[sosaIndex[f], PRENOM], liste[sosaIndex[f], PATRONYME]);
                            if (nom == "")
                            {
                                gfx.DrawLine(penG, positionBoite[f, 0] + 2, positionBoite[f, 1] + hauteurLigne * 3, positionBoite[f, 0] + 2 + 140, positionBoite[f, 1] + hauteurLigne * 3);
                            }
                            string rt = RacoucirNom(nom, ref gfx);
                            gfx.DrawString(rt, font8B, XBrushes.Black, positionBoite[f, 0], positionBoite[f, 1] + hauteurLigne * 2 + 3, XStringFormats.Default);
                        }
                    }
                    int p = 18;
                    int l = 140;


                    // mariage 1 sosa = 1
                    if (sosa != 1 ) {
                        if (sosa%2 == 0 ) {
                        
                            if (liste[sosa, MALE] == "")
                            {
                                gfx.DrawLine(penG, Col2 + p, positionBoite[1, 1] + hauteurLigne* 13, Col2 + l, positionBoite[1, 1] + hauteurLigne * 13);
                            }
                            gfx.DrawString(liste[sosa, MALE], font8, XBrushes.Black, Col2 + p, positionBoite[1, 1] + hauteurLigne * 13, XStringFormats.Default);
                            if (liste[sosa, MALIEU] == "")
                            {
                                gfx.DrawLine(penG, Col2 + p, positionBoite[1, 1] + hauteurLigne * 15, Col2 + l, positionBoite[1, 1] + hauteurLigne * 15);
                            }
                            gfx.DrawString(liste[sosa, MALIEU], font8, XBrushes.Black, Col2 + p, positionBoite[1, 1] + hauteurLigne * 15, XStringFormats.Default);
                        } else {
                            if (liste[sosa-1, MALE] == "") 
                            {
                                gfx.DrawLine(penG, Col2 + p, positionBoite[1, 1] + hauteurLigne * 13, Col2 + l, positionBoite[1, 1] + hauteurLigne * 13);
                            }
                            gfx.DrawString(liste[sosa-1, MALE], font8, XBrushes.Black, Col2 + p, positionBoite[1, 1] + hauteurLigne * 13, XStringFormats.Default);
                            if (liste[sosa-1, MALIEU] == "")
                            {
                                gfx.DrawLine(penG, Col2 + p, positionBoite[1, 1] + hauteurLigne * 15, Col2 + l, positionBoite[1, 1] + hauteurLigne * 15);
                            }
                            gfx.DrawString(liste[sosa-1, MALIEU], font8, XBrushes.Black, Col2 + p, positionBoite[1, 1] + hauteurLigne * 15, XStringFormats.Default);
                        }

                    }

                    // mariage 2 3
                    if (liste[sosa * 2, MALE] == "")
                    {
                        gfx.DrawLine(penG, Col4 + p, positionBoite[2, 1] + hauteurLigne * 20, Col4 + l, positionBoite[2, 1] + hauteurLigne * 20);
                    }
                    gfx.DrawString(liste[sosa * 2, MALE], font8, XBrushes.Black, Col4 + p, positionBoite[2, 1] + hauteurLigne * 20, XStringFormats.Default);
                    if (liste[sosa * 2, MALIEU] == "")
                    {
                        gfx.DrawLine(penG, Col4 + p, positionBoite[2, 1] + hauteurLigne * 22, Col4 + l, positionBoite[2, 1] + hauteurLigne * 22);
                    }
                    gfx.DrawString(liste[sosa * 2, MALIEU], font8, XBrushes.Black, Col4 + p, positionBoite[2, 1] + hauteurLigne * 22, XStringFormats.Default);

                    // mariage 4 5
                    if (liste[sosa * 4, MALE] == "")
                    {
                        gfx.DrawLine(penG, Col6 + p, positionBoite[4, 1] + hauteurLigne * 13, Col6 + l, positionBoite[4, 1] + hauteurLigne * 13);
                    }
                    gfx.DrawString(liste[sosa * 4, MALE], font8, XBrushes.Black, Col6 + p , positionBoite[4, 1] + hauteurLigne * 13, XStringFormats.Default);

                    if (liste[sosa * 4, MALIEU] == "")
                    {
                        gfx.DrawLine(penG, Col6 + p, positionBoite[4, 1] + hauteurLigne * 15, Col6 + l, positionBoite[4, 1] + hauteurLigne * 15);
                    }
                    gfx.DrawString(liste[sosa * 4, MALIEU], font8, XBrushes.Black, Col6 + p, positionBoite[4, 1] + hauteurLigne * 15, XStringFormats.Default);
                    // mariage 6 7
                    if (liste[sosa * 4 + 2, MALE] == "")
                    {
                        gfx.DrawLine(penG, Col6 + p, positionBoite[6, 1] + hauteurLigne * 13, Col6 + l, positionBoite[6, 1] + hauteurLigne * 13);
                    }
                    gfx.DrawString(liste[sosa * 4 + 2, MALE], font8, XBrushes.Black, Col6 + p, positionBoite[6, 1] + hauteurLigne * 13, XStringFormats.Default);
                    if (liste[sosa * 4 + 2, MALIEU] == "")
                    {
                        gfx.DrawLine(penG, Col6 + p, positionBoite[6, 1] + hauteurLigne * 15, Col6 + l, positionBoite[6, 1] + hauteurLigne * 15);
                    }
                    gfx.DrawString(liste[sosa * 4 + 2, MALIEU], font8, XBrushes.Black, Col6 + p, positionBoite[6, 1] + hauteurLigne * 15, XStringFormats.Default);
                    // mariage 8 9
                    int s = sosa * 8;
                    if (s < 512)
                    {
                        if (liste[s, MALE] == "")
                        {
                            gfx.DrawLine(penG, Col8 + p, positionBoite[8, 1] + hauteurLigne * 6, Col8 + l, positionBoite[8, 1] + hauteurLigne * 6);
                        }
                        gfx.DrawString(liste[s, MALE], font8, XBrushes.Black, Col8 + p, positionBoite[8, 1] + hauteurLigne * 6, XStringFormats.Default);
                        if (liste[s, MALIEU] == "")
                        {
                            gfx.DrawLine(penG, Col8 + p, positionBoite[8, 1] + hauteurLigne * 8, Col8 + l, positionBoite[8, 1] + hauteurLigne * 8);
                        }
                        gfx.DrawString(liste[s, MALIEU], font8, XBrushes.Black, Col8 + p, positionBoite[8, 1] + hauteurLigne * 8, XStringFormats.Default);
                    }
                    // mariage 10 11
                    s = sosa * 8 + 2;
                    if (s < 512)
                    {
                        if (liste[s, MALE] == "")
                        {
                            gfx.DrawLine(penG, Col8 + p, positionBoite[10, 1] + hauteurLigne * 6, Col8 + l, positionBoite[10, 1] + hauteurLigne * 6);
                        }
                        gfx.DrawString(liste[s, MALE], font8, XBrushes.Black, Col8 + p, positionBoite[10, 1] + hauteurLigne * 6, XStringFormats.Default);
                        if (liste[s, MALIEU] == "")
                        {
                            gfx.DrawLine(penG, Col8 + p, positionBoite[10, 1] + hauteurLigne * 8, Col8 + l, positionBoite[10, 1] + hauteurLigne * 8);
                        }
                        gfx.DrawString(liste[s, MALIEU], font8, XBrushes.Black, Col8 + p, positionBoite[10, 1] + hauteurLigne * 8, XStringFormats.Default);
                    }
                    // mariage 12 13
                    s = sosa * 8 + 4;
                    if (s < 512)
                    {
                        if (liste[s, MALE] == "")
                        {
                            gfx.DrawLine(penG, Col8 + p, positionBoite[12, 1] + hauteurLigne * 6, Col8 + l, positionBoite[12, 1] + hauteurLigne * 6);
                        }
                        gfx.DrawString(liste[s, MALE], font8, XBrushes.Black, Col8 + p, positionBoite[12, 1] + hauteurLigne * 6, XStringFormats.Default);
                        if (liste[s, MALIEU] == "")
                        {
                            gfx.DrawLine(penG, Col8 + p, positionBoite[12, 1] + hauteurLigne * 8, Col8 + l, positionBoite[12, 1] + hauteurLigne * 8);
                        }
                        gfx.DrawString(liste[s, MALIEU], font8, XBrushes.Black, Col8 + p, positionBoite[12, 1] + hauteurLigne * 8, XStringFormats.Default);
                    }
                    // mariage 14 15
                    s = sosa * 8 + 6;
                    if (s < 512)
                    {
                        if (liste[s, MALE] == "") 
                        {
                            gfx.DrawLine(penG, Col8 + p, positionBoite[14, 1] + hauteurLigne * 6 , Col8 + l, positionBoite[14, 1] + hauteurLigne * 6);
                        }
                        gfx.DrawString(liste[s, MALE], font8, XBrushes.Black, Col8 + p, positionBoite[14, 1] + hauteurLigne * 6, XStringFormats.Default);
                        if (liste[s, MALIEU] == "")
                        {
                            gfx.DrawLine(penG, Col8 + p, positionBoite[14, 1] + hauteurLigne * 8, Col8 + l, positionBoite[14, 1] + hauteurLigne * 8);
                        }
                        gfx.DrawString(liste[s, MALIEU], font8, XBrushes.Black, Col8 + p, positionBoite[14, 1] + hauteurLigne * 8, XStringFormats.Default);
                    }
                }
            }
            //dessiner  flèche
            if (fleche)
            {
                // flèche gauche
                FlecheGauche(gfx, font8, Col1, positionBoite[1, 1], hauteurBoite, sosa);

                // flèche doite
                if (sosa == 0)
                {
                    FlecheDroite(gfx, font8, Col10, positionBoite[08, 1], hauteurBoiteMini, 0);
                    FlecheDroite(gfx, font8, Col10, positionBoite[09, 1], hauteurBoiteMini, 0);
                    FlecheDroite(gfx, font8, Col10, positionBoite[10, 1], hauteurBoiteMini, 0);
                    FlecheDroite(gfx, font8, Col10, positionBoite[11, 1], hauteurBoiteMini, 0);
                    FlecheDroite(gfx, font8, Col10, positionBoite[12, 1], hauteurBoiteMini, 0);
                    FlecheDroite(gfx, font8, Col10, positionBoite[13, 1], hauteurBoiteMini, 0);
                    FlecheDroite(gfx, font8, Col10, positionBoite[14, 1], hauteurBoiteMini, 0);
                    FlecheDroite(gfx, font8, Col10, positionBoite[15, 1], hauteurBoiteMini, 0);
                }
                else
                {
                    FlecheDroite(gfx, font8, Col10, positionBoite[08, 1], hauteurBoiteMini, sosa * 8);
                    FlecheDroite(gfx, font8, Col10, positionBoite[09, 1], hauteurBoiteMini, sosa * 8 + 1);
                    FlecheDroite(gfx, font8, Col10, positionBoite[10, 1], hauteurBoiteMini, sosa * 8 + 2);
                    FlecheDroite(gfx, font8, Col10, positionBoite[11, 1], hauteurBoiteMini, sosa * 8 + 3);
                    FlecheDroite(gfx, font8, Col10, positionBoite[12, 1], hauteurBoiteMini, sosa * 8 + 4);
                    FlecheDroite(gfx, font8, Col10, positionBoite[13, 1], hauteurBoiteMini, sosa * 8 + 5);
                    FlecheDroite(gfx, font8, Col10, positionBoite[14, 1], hauteurBoiteMini, sosa * 8 + 6);
                    FlecheDroite(gfx, font8, Col10, positionBoite[15, 1], hauteurBoiteMini, sosa * 8 + 7);
                }
            }

            // Note 1
            rect = new XRect(Col1, Ligne[14], Col3 - Col1, hauteurLigne  * 24);
            //gfx.DrawRectangle(penM, rect);
            tf.Alignment = XParagraphAlignment.Justify;
            tf.DrawString(liste[sosa, NOTE1], font8, XBrushes.Black, rect, XStringFormats.TopLeft);

            // Note 2
            rect = new XRect(Col1, Ligne[61], Col3 - Col1, hauteurLigne  * 18);
            //gfx.DrawRectangle(penM, rect);
            tf.Alignment = XParagraphAlignment.Justify;
            tf.DrawString(liste[sosa, NOTE2], font8, XBrushes.Black, rect, XStringFormats.TopLeft);
            
            // bas de page
            if (fleche)
            {
                numeroTableau = liste[sosa, TABLEAU];
                gfx.DrawString("Tableau", font8, XBrushes.Black, Col8 + 100, Ligne[80], XStringFormats.TopLeft);

                if (numeroTableau != "")
                {
                    gfx.DrawString(numeroTableau, font8, XBrushes.Black, Col8 + 135, Ligne[80], XStringFormats.TopLeft);
                }
                else
                {
                    gfx.DrawString("_____", font8, XBrushes.Black, Col8 + 135, Ligne[80], XStringFormats.TopLeft);
                }
            }

            if (PreparerPar.Text != "" && !tous)
            {
                gfx.DrawString("Préparé par " + PreparerPar.Text + " le " + DateLb.Text, font8, XBrushes.Black, Col1, Ligne[81], XStringFormats.Default);
            }

            // version à adfficher pour beta
            gfx.DrawString("Version " + Application.ProductVersion + "B", font8, XBrushes.Black, Col1, Ligne[82], XStringFormats.Default);
            // Logo
            XImage img = global::TableauAscendant.Properties.Resources.dapamv5_32png;
            
            XPen penDapam = new XPen(XColor.FromArgb(0, 0, 0), 2);
            XFont fontDapam = new XFont("Arial", 14, XFontStyle.Bold);
            XFont fontDesign = new XFont("Arial", 5.5, XFontStyle.Italic);
            gfx.DrawRoundedRectangle(penDapam,gris, pouce * 8.03, pouce * 7.82, 59, 20, 15, 15);
            gfx.DrawString("DAPAM", fontDapam, XBrushes.Black, pouce * 8.08, pouce * 8.025);
            gfx.DrawString("Design", fontDesign, XBrushes.Black, pouce * 7.75, pouce * 7.9);
            
            return numeroTableau;
        }
        private void ZXCV(string message, [CallerLineNumber] int lineNumber = 0, [CallerMemberName] string caller = null)
        {
            if (LOGACTIF) { 
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
        private void    NouvellePage(ref PdfDocument document, ref XGraphics gfx, ref PdfPage page, string Orientation)
        {
            page = document.AddPage();
            page.Size = PageSize.Letter;
            if (Orientation == "L")
            {
                page.Orientation = PdfSharp.PageOrientation.Landscape;
            }
            else
            {
                page.Orientation = PdfSharp.PageOrientation.Portrait;
            }
            gfx = XGraphics.FromPdfPage(page);

        }
        private Boolean LireData()
        {
            int index=0;
            string s;
            string crochet = "[]";
            string version = "";
            if (File.Exists(FichierCourant))
            {
                try
                {
                    using (StreamReader sr = File.OpenText(FichierCourant))
                    {
                        EffacerData();
                        s = sr.ReadLine();
                        if (s != "[ver**]")
                        {
                            SystemSounds.Beep.Play();
                            MessageBox.Show("Les fichiers de version précédente ne sont pas valides.\r\n\r\n", "Fichier non valide",
                                             MessageBoxButtons.OK,
                                             MessageBoxIcon.Warning);
                            return false;
                        }
                        s = sr.ReadLine();
                        version = s.Substring(7);
                        s = sr.ReadLine();
                        while (s != "[FIN**]")
                        {
                            if (s == "[sosa*]")
                            {
                                s = sr.ReadLine();
                                while (s[0] != crochet[0])
                                {
                                    if ( s.Substring ( 0,7) == "No    =")
                                    {
                                        index = Int32.Parse(s.Substring(7));
                                        if (s.Length > 7) liste[index, SOSA] = s.Substring(7);

                                        s = sr.ReadLine();
                                    }
                                    if (s.Substring(0, 7) == "Nom   =")
                                    {
                                        if (s.Length > 7) liste[index, PATRONYME] = s.Substring(7);
                                        s = sr.ReadLine();
                                    }
                                    if (s.Substring(0, 7) == "Prenom=")
                                    {
                                        if (s.Length > 7) liste[index, PRENOM] = s.Substring(7);
                                        s = sr.ReadLine();
                                    }
                                    if ( s.Substring( 0,7) == "NeLe  =")
                                    {
                                        if (s.Length > 7) liste[index, NELE] = s.Substring(7);
                                        s = sr.ReadLine();
                                    }
                                    if (s.Substring(0, 7) == "NeLieu=")
                                    {
                                        if (s.Length > 7) liste[index, NELIEU] = s.Substring(7);
                                        s = sr.ReadLine();
                                    }
                                    if (s.Substring(0, 7) == "DeLe  =")
                                    {
                                        if (s.Length > 7) liste[index, DELE] = s.Substring(7);
                                        s = sr.ReadLine();
                                    }
                                    if (s.Substring(0, 7) == "DeLieu=")
                                    {
                                        if (s.Length > 7) liste[index, DELIEU] = s.Substring(7);
                                        s = sr.ReadLine();
                                    }
                                    if (s.Substring(0, 7) == "MaLe  =")
                                    {
                                        if (s.Length > 7) liste[index, MALE] = s.Substring(7);
                                        s = sr.ReadLine();
                                    }
                                    if (s.Substring(0, 7) == "MaLieu=")
                                    {
                                        if (s.Length > 7) liste[index, MALIEU] = s.Substring(7);
                                        s = sr.ReadLine();
                                    }
                                    if (s.Substring(0, 7) == "NoteH =")
                                    {
                                        s = sr.ReadLine();
                                        while (s != "##FIN##")
                                        {
                                            if (liste[index, NOTE1] == "")
                                            {
                                                liste[index, NOTE1] = s;
                                                s = sr.ReadLine();
                                            }
                                            else
                                            {
                                                liste[index, NOTE1] = liste[index, NOTE1] + "\r\n" + s;
                                                s = sr.ReadLine();
                                            }
                                        }
                                        s = sr.ReadLine();
                                    }
                                    if (s.Substring(0, 7) == "NoteB =")
                                    {
                                        s = sr.ReadLine();
                                        while (s != "##FIN##")
                                        {
                                            if (liste[index, NOTE2] == "")
                                            {
                                                liste[index, NOTE2] = s;
                                                s = sr.ReadLine();
                                            }
                                            else
                                            {
                                                liste[index, NOTE2] = liste[index, NOTE2] + "\r\n" + s;
                                                s = sr.ReadLine();
                                            }
                                        }
                                        s = sr.ReadLine();
                                    }
                                    liste[index, NOMTRI] = "";
                                    if (liste[index, PATRONYME] != "" && liste[index, PRENOM] != "")
                                    {
                                        liste[index, NOMTRI] = liste[index, PATRONYME] + " " + liste[index, PRENOM];
                                    }
                                    if (liste[index, PATRONYME] != "" && liste[index, PRENOM] == "")
                                    {
                                        liste[index, NOMTRI] = liste[index, PATRONYME];
                                    }
                                    if (liste[index, PATRONYME] == "" && liste[index, PRENOM] != "")
                                    {
                                        liste[index, NOMTRI] = " " + liste[index, PRENOM];
                                    }
                                }
                            }
                            if (s == "[par**]")
                            {
                                s = sr.ReadLine();
                                while (s[0] != crochet[0])
                                {
                                    if (s.Substring(0, 7) == "Par   =")
                                    {
                                        if (s.Length > 7) PreparerPar.Text = s.Substring(7);
                                        s = sr.ReadLine();
                                    }
                                }
                            }
                            if (s == "[Asc**]")
                            {
                                s = sr.ReadLine();
                                while (s[0] != crochet[0])
                                {
                                    if (s.Substring(0, 7) == "Ascend=")
                                    {
                                        if (s.Length > 7) AscendantDeTb.Text = s.Substring(7);
                                        s = sr.ReadLine();
                                    }
                                }
                            }
                        }
                    }
                    ChoixSosaComboBox.Text = "1";
                    Modifier = false;
                    this.Text = NomPrograme + "   " + FichierCourant;
                    return true;
                }
                catch (Exception m)
                {
                    SystemSounds.Beep.Play();
                    MessageBox.Show("Ne peut pas lire le fichier du data.\r\n\r\n" + m.Message, "Problème ?",
                                     MessageBoxButtons.OK,
                                     MessageBoxIcon.Warning);
                    return false;
                }

            }
            return true;
        }
        private Boolean LongeurNomtOk(string nom)
        {
            PdfDocument doc = new PdfDocument();
            PdfPage page = doc.AddPage();
            XGraphics gfx = XGraphics.FromPdfPage(page);
            XSize nomInfo = gfx.MeasureString(nom, font8B);
            if (nomInfo.Width <= LARGEURNOMFICHE)
            {
                return true;
            }
            else
            {
                return false;
            }
        }
        private Boolean LongeurTextOk( string text)
        {
            PdfDocument doc = new PdfDocument();
            PdfPage page = doc.AddPage();
            XGraphics gfx = XGraphics.FromPdfPage(page);
            XSize textInfo = gfx.MeasureString(text, font8);
            if (textInfo.Width <= LARGEURTEXTEFICHE) {
                return true;
            } 
            else {
                return false;
            }
        }
        private void    PDFEcrire(ref XGraphics gfx, string text, double X, double Y, double L)
        {

            XSize textInfo = gfx.MeasureString(text, font8);
            if (textInfo.Width <= L)
            { 
                gfx.DrawString(text, font8, XBrushes.Black, X, Y);
                return;
            }
            textInfo = gfx.MeasureString(text, font7);
            if (textInfo.Width <= L)
            {
                gfx.DrawString(text, font7, XBrushes.Black, X, Y);
                return;
            }
            textInfo = gfx.MeasureString(text, font6);
            if (textInfo.Width <= L)
            {
                gfx.DrawString(text, font6, XBrushes.Black, X, Y);
                return;
            }
            gfx.DrawString(text, font5, XBrushes.Black, X, Y);
         }
        private void    PDFEcrireCentrer(ref XGraphics gfx, string text, double X, double Y, double XX)
        {
            double L = XX - X; // largeur

            //page.Width / 2 - textLargeur.Width / 2
            XSize textInfo = gfx.MeasureString(text, font8);
            if (textInfo.Width <= L)
            {
                gfx.DrawString(text, font8, XBrushes.Black, X + ((XX-X) /2) - (textInfo.Width / 2), Y);
                return;
            }
            textInfo = gfx.MeasureString(text, font7);
            if (textInfo.Width <= L)
            {
                gfx.DrawString(text, font7, XBrushes.Black, X + 5 + L / 2, Y);
                return;
            }
            textInfo = gfx.MeasureString(text, font6);
            if (textInfo.Width <= L)
            {
                gfx.DrawString(text, font6, XBrushes.Black, X + 5 + L / 2, Y);
                return;
            }
            gfx.DrawString(text, font5, XBrushes.Black, X + 5 + L / 2, Y);
        }
        private string  RacoucirNom(string nom,  ref XGraphics gfx)
        {
            XSize nomInfo = gfx.MeasureString(nom, font8B);
            if (nomInfo.Width <= LARGEURNOMFICHE)
            {
                return nom;
            }
            do
            {
                nom = nom.Remove(nom.Length - 1);
                nomInfo = gfx.MeasureString( nom + "...", font8B);
            } while (nomInfo.Width > LARGEURNOMFICHE);
            return nom + "...";
        }
        private string  RacoucirTexte (string text, ref XGraphics gfx )
        {
            //XFont font8 = new XFont("Arial", 8, XFontStyle.Regular);
            XSize textInfo = gfx.MeasureString(text, font8);
            if (textInfo.Width <= LARGEURTEXTEFICHE)
            {
                return text;
            }
            do
            {
                text = text.Substring(1);
                textInfo = gfx.MeasureString("..." + text, font8);
            } while (textInfo.Width  > LARGEURTEXTEFICHE);
            return "..." + text;
        }
        private void    RafraichirData()
        {
            sosaCourant = Int32.Parse(ChoixSosaComboBox.Text);
            int index;
            // affiche les informations
            index = sosaCourant;

            if (index == 0)
            {
                Note1.Visible = false;
                Note2.Visible = false;
            }
            Sosa1PatronymeTextBox.Text = liste[index, PATRONYME];
            Sosa1PrenomTextBox.Text = liste[index, PRENOM];
            Sosa1NeTextBox.Text = liste[index, NELE];
            Sosa1NeEndroitTextBox.Text = liste[index, NELIEU];
            Sosa1DeTextBox.Text = liste[index, DELE];
            Sosa1DeEndroitTextBox.Text = liste[index, DELIEU];
            Sosa1MaTextBox.Text = liste[index, MALE];
            Sosa1MaEndroitTextBox.Text = liste[index, MALIEU];
            if (index > 1)
            {
                int i = index % 2;
                if (index % 2 == 0)
                {
                    SosaConjoint1PatronymeTextBox.Text = liste[index + 1, PATRONYME];
                    SosaConjoint1PrenomTextBox.Text = liste[index + 1, PRENOM];
                    SosaConjoint1PatronymeTextBox.Visible = true;
                    Conjoint1Lbl.Visible = true;
                    SosaConjoint1PrenomTextBox.Visible = true;
                    SosaConjoint1Label.Text = (index + 1).ToString();
                    SosaConjoint1Label.Visible = true;
                }
                else
                {
                    SosaConjoint1PatronymeTextBox.Visible = false;
                    SosaConjoint1PrenomTextBox.Visible = false;
                    SosaConjoint1Label.Visible = false;
                    Conjoint1Lbl.Visible = false;
                }
            }
            Note1.Text = liste[index, NOTE1];
            Note2.Text = liste[index, NOTE2];
            GenerationAlb.Text = liste[index, GENERATION];

            index = sosaCourant * 2;
            Sosa2Label.Text = liste[index, SOSA];
            Sosa2PatronymeTextBox.Text = liste[index, PATRONYME];
            Sosa2PrenomTextBox.Text = liste[index, PRENOM];
            Sosa2NeTextBox.Text = liste[index, NELE];
            Sosa2NeEndroitTextBox.Text = liste[index, NELIEU];
            Sosa2DeTextBox.Text = liste[index, DELE];
            Sosa2DeEndroitTextBox.Text = liste[index, DELIEU];
            Sosa23MaTextBox.Text = liste[index, MALE];
            Sosa23MaEndroitTextBox.Text = liste[index, MALIEU];
            GenerationBlb.Text = liste[index, GENERATION];

            index = sosaCourant * 2 + 1;
            Sosa3Label.Text = liste[index, SOSA];
            Sosa3PatronymeTextBox.Text = liste[index, PATRONYME];
            Sosa3PrenomTextBox.Text = liste[index, PRENOM];
            Sosa3NeTextBox.Text = liste[index, NELE];
            Sosa3NeEndroitTextBox.Text = liste[index, NELIEU];
            Sosa3DeTextBox.Text = liste[index, DELE];
            Sosa3DeEndroitTextBox.Text = liste[index, DELIEU];

            index = sosaCourant * 4;
            Sosa4Label.Text = liste[index, SOSA];
            Sosa4PatronymeTextBox.Text = liste[index, PATRONYME];
            Sosa4PrenomTextBox.Text = liste[index, PRENOM];
            Sosa4NeTextBox.Text = liste[index, NELE];
            Sosa4NeEndroitTextBox.Text = liste[index, NELIEU];
            Sosa4DeTextBox.Text = liste[index, DELE];
            Sosa4DeEndroitTextBox.Text = liste[index, DELIEU];
            Sosa45MaTextBox.Text = liste[index, MALE];
            Sosa45MaLEndroitTextBox.Text = liste[index, MALIEU];
            GenerationClb.Text = liste[index, GENERATION];

            index = sosaCourant * 4 + 1;
            Sosa5Label.Text = liste[index, SOSA];
            Sosa5PatronymeTextBox.Text = liste[index, PATRONYME];
            Sosa5PrenomTextBox.Text = liste[index, PRENOM];
            Sosa5NeTextBox.Text = liste[index, NELE];
            Sosa5NeEndroitTextBox.Text = liste[index, NELIEU];
            Sosa5DeTextBox.Text = liste[index, DELE];
            Sosa5DeEndroitTextBox.Text = liste[index, DELIEU];

            index = sosaCourant * 4 + 2;
            Sosa6Label.Text = liste[index, SOSA];
            Sosa6PatronymeTextBox.Text = liste[index, PATRONYME];
            Sosa6PrenomTextBox.Text = liste[index, PRENOM];
            Sosa6NeTextBox.Text = liste[index, NELE];
            Sosa6NeEndroitTextBox.Text = liste[index, NELIEU];
            Sosa6DeTextBox.Text = liste[index, DELE];
            Sosa6DeEndroitTextBox.Text = liste[index, DELIEU];
            Sosa67MaTextBox.Text = liste[index, MALE];
            Sosa67MaEndroitTextBox.Text = liste[index, MALIEU];

            index = sosaCourant * 4 + 3;
            Sosa7Label.Text = liste[index, SOSA];
            Sosa7PatronymeTextBox.Text = liste[index, PATRONYME];
            Sosa7PrenomTextBox.Text = liste[index, PRENOM];
            Sosa7NeTextBox.Text = liste[index, NELE];
            Sosa7NeEndroitTextBox.Text = liste[index, NELIEU];
            Sosa7DeTextBox.Text = liste[index, DELE];
            Sosa7DeEndroitTextBox.Text = liste[index, DELIEU];
        }
        private void    RechercheID()
        {
            DataTable listeAChoisir = new DataTable();
            listeAChoisir.Columns.Add("ID", typeof(string));
            listeAChoisir.Columns.Add("Nom", typeof(string));
            listeAChoisir.Columns.Add("Naissance", typeof(string));
            listeAChoisir.Columns.Add("Deces", typeof(string));
            ContinuerBtn.Visible = false;
            List<string> IDListe = new List<string>();
            IDListe = GEDCOM.RechercheIndividu(PatronymeRecherche.Text, PrenomRecherche.Text);
            foreach (string info in IDListe)
            {
                listeAChoisir.Rows.Add(info, GEDCOM.AvoirPatronyme(info) + " " + GEDCOM.AvoirPrenom(info),
                    ConvertirDate(GEDCOM.AvoirDateNaissance(info)), ConvertirDate(GEDCOM.AvoirDateDeces(info)));
            }
            DataView trier = new DataView(listeAChoisir)
            {
                Sort = "Nom ASC"
            };
            ListViewItem itm;
            ChoixLV.Items.Clear();
            for (int f = 0; f < trier.Count; f++)
            {
                string[] ligne = new string[4];
                ligne[0] = trier[f]["ID"].ToString();
                ligne[1] = trier[f]["Nom"].ToString();
                ligne[2] = trier[f]["Naissance"].ToString();
                ligne[3] = trier[f]["Deces"].ToString();
                itm = new ListViewItem(ligne);
                ChoixLV.Items.Add(itm);
            }
        }
        private void    SosaChanger()
        {
            int index;
            if (sosaCourant > 0)
            {
                index = sosaCourant;
                liste[index, SOSA] = index.ToString();
                liste[index, PATRONYME] = Sosa1PatronymeTextBox.Text;
                liste[index, PRENOM] = Sosa1PrenomTextBox.Text;
                liste[index, NELE] = Sosa1NeTextBox.Text;
                liste[index, NELIEU] = Sosa1NeEndroitTextBox.Text;
                liste[index, DELE] = Sosa1DeTextBox.Text;
                liste[index, DELIEU] = Sosa1DeEndroitTextBox.Text;
                liste[index, NOTE1] = Note1.Text;
                liste[index, NOTE2] = Note2.Text;
                index = sosaCourant * 2;
                liste[index, SOSA] = index.ToString();
                liste[index, PATRONYME] = Sosa2PatronymeTextBox.Text;
                liste[index, PRENOM] = Sosa2PrenomTextBox.Text;
                liste[index, NELE] = Sosa2NeTextBox.Text;
                liste[index, NELIEU] = Sosa2NeEndroitTextBox.Text;
                liste[index, DELE] = Sosa2DeTextBox.Text;
                liste[index, DELIEU] = Sosa2DeEndroitTextBox.Text;
                liste[index, MALE] = Sosa23MaTextBox.Text;
                liste[index, MALIEU] = Sosa23MaEndroitTextBox.Text;

                index = sosaCourant * 2 + 1;
                liste[index, SOSA] = index.ToString();
                liste[index, PATRONYME] = Sosa3PatronymeTextBox.Text;
                liste[index, PRENOM] = Sosa3PrenomTextBox.Text;
                liste[index, NELE] = Sosa3NeTextBox.Text;
                liste[index, NELIEU] = Sosa3NeEndroitTextBox.Text;
                liste[index, DELE] = Sosa3DeTextBox.Text;
                liste[index, DELIEU] = Sosa3DeEndroitTextBox.Text;

                index = sosaCourant * 4;
                liste[index, SOSA] = index.ToString();
                liste[index, PATRONYME] = Sosa4PatronymeTextBox.Text;
                liste[index, PRENOM] = Sosa4PrenomTextBox.Text;
                liste[index, NELE] = Sosa4NeTextBox.Text;
                liste[index, NELIEU] = Sosa4NeEndroitTextBox.Text;
                liste[index, DELE] = Sosa4DeTextBox.Text;
                liste[index, DELIEU] = Sosa4DeEndroitTextBox.Text;
                liste[index, MALE] = Sosa45MaTextBox.Text;
                liste[index, MALIEU] = Sosa45MaLEndroitTextBox.Text;

                index = sosaCourant * 4 + 1;

                liste[index, SOSA] = index.ToString();
                liste[index, PATRONYME] = Sosa5PatronymeTextBox.Text;
                liste[index, PRENOM] = Sosa5PrenomTextBox.Text;
                liste[index, NELE] = Sosa5NeTextBox.Text;
                liste[index, NELIEU] = Sosa5NeEndroitTextBox.Text;
                liste[index, DELE] = Sosa5DeTextBox.Text;
                liste[index, DELIEU] = Sosa5DeEndroitTextBox.Text;
                liste[index, MALE] = "";
                liste[index, MALIEU] = "";

                index = sosaCourant * 4 + 2;
                liste[index, SOSA] = index.ToString();
                liste[index, PATRONYME] = Sosa6PatronymeTextBox.Text;
                liste[index, PRENOM] = Sosa6PrenomTextBox.Text;
                liste[index, NELE] = Sosa6NeTextBox.Text;
                liste[index, NELIEU] = Sosa6NeEndroitTextBox.Text;
                liste[index, DELE] = Sosa6DeTextBox.Text;
                liste[index, DELIEU] = Sosa6DeEndroitTextBox.Text;
                liste[index, MALE] = Sosa67MaTextBox.Text;
                liste[index, MALIEU] = Sosa67MaEndroitTextBox.Text;

                index = sosaCourant * 4 + 3;
                liste[index, SOSA] = index.ToString();
                liste[index, PATRONYME] = Sosa7PatronymeTextBox.Text;
                liste[index, PRENOM] = Sosa7PrenomTextBox.Text;
                liste[index, NELE] = Sosa7NeTextBox.Text;
                liste[index, NELIEU] = Sosa7NeEndroitTextBox.Text;
                liste[index, DELE] = Sosa7DeTextBox.Text;
                liste[index, DELIEU] = Sosa7DeEndroitTextBox.Text;
            }
            if (ChoixSosaComboBox.Text == "")
            {
                sosaCourant = 0;
            }
            else
            {
                RafraichirData();
            }
            AfficherData();
        }
        private string  StrDate(string date)
        {
            date = date.ToLower();
            date = date.Replace("and", "et").Replace("bet", "entre").Replace("abt", "vers").Replace("aft", "après").Replace("bef", "avant").Replace("abt", "autour");
            if (date.Contains("et") || date.Contains("entre") || date.Contains("vers") || date.Contains("après") || date.Contains("avant")  || date.Contains("autour"))
            {
                return date;
            }
            date = date.ToLower();
            string[] d = date.Split(' ');
            if (d.Length == 1)
            {
                return date;
            }
                if (d.Length == 3) {
                if (d[1] == "jan") d[1] = "01";
                if (d[1] == "feb") d[1] = "02";
                if (d[1] == "mar") d[1] = "03";
                if (d[1] == "apr") d[1] = "04";
                if (d[1] == "may") d[1] = "05";
                if (d[1] == "jun") d[1] = "06";
                if (d[1] == "jul") d[1] = "07";
                if (d[1] == "aug") d[1] = "08";
                if (d[1] == "sep") d[1] = "09";
                if (d[1] == "oct") d[1] = "10";
                if (d[1] == "nov") d[1] = "11";
                if (d[1] == "dec") d[1] = "12";
                return d[2] + "-" + d[1] + d[0];
            }
            return "";
        }
        private void    TableMatiere(ref PdfDocument document, ref XGraphics gfx, ref PdfPage page)
        {
            DataTable listeATrier = new DataTable();
            listeATrier.Columns.Add("Sosa", typeof(string));
            listeATrier.Columns.Add("Nom", typeof(string));
            listeATrier.Columns.Add("Tableau", typeof(string));
            for (int f = 0; f < 512; f++)
            {
                if (liste[f, NOMTRI] != "" && liste[f, SOSA] != "0") {
                    listeATrier.Rows.Add(liste[f, SOSA], liste[f, NOMTRI],liste[f, TABLEAU]);
                }
                
            }
            DataView trier = new DataView(listeATrier)
            {
                Sort = "Nom ASC"
            };
            XFont font32 = new XFont("arial", 32, XFontStyle.Bold);
            XFont font8 = new XFont("Arial", 8, XFontStyle.Regular);
            XPen pen = new XPen(XColor.FromArgb(0, 0, 0));
            XPen penD = new XPen(XColor.FromArgb(150, 150, 150), 0.25)
            {
                DashStyle = XDashStyle.Dot
            };
            double x = POUCE * .5;
            double y = POUCE * .5;
            XRect rect = new XRect();
            //string str;
            EnteteTableMatiere(ref document, ref gfx, ref page);

            XTextFormatter tf = new XTextFormatter(gfx)
            {
                Alignment = XParagraphAlignment.Center
            };
            rect = new XRect(POUCE * .5, POUCE * .5, POUCE * 7.5, 10);
            tf.DrawString("Table des matières" , font32, XBrushes.Black, rect, XStringFormats.TopLeft);
            double Top = 90;
            y = Top;
            int MaxLigne = 63;
            int NombrePage = 1;
            int NumeroLigne = 0;
            int Col = 1;
            for (int f = 1; f < trier.Count; f++)
            {
                tf = new XTextFormatter(gfx);
                if (Col == 1) x = POUCE * .5;
                if (Col == 2) x = POUCE * 4.5;
                tf.Alignment = XParagraphAlignment.Left;
                string nom = trier[f]["Nom"].ToString();
                if (nom.Length > 0 && nom != "0")
                {
                    // largeur maximum nom 240
                    XSize textLargeur = gfx.MeasureString(nom, font8);
                    if (textLargeur.Width > 220)
                    {
                        textLargeur.Width = 220;
                    }
                    
                    // sosa
                    tf.Alignment = XParagraphAlignment.Right;
                    rect = new XRect(x, y, 15, 10);
                    string sosa = trier[f]["Sosa"].ToString();
                    tf.DrawString(sosa, font8, XBrushes.Black, rect, XStringFormats.TopLeft);
                    gfx.DrawLine(penD, x , y + 8, x + 15, y + 8);

                    //nom
                    tf.Alignment = XParagraphAlignment.Left;
                    rect = new XRect(x + 18, y, 220, 10);
                    tf.DrawString(nom, font8, XBrushes.Black, rect, XStringFormats.TopLeft);
                    gfx.DrawLine(penD, x + 18 , y + 8, x + 235, y + 8);

                    // tableau 
                    tf.Alignment = XParagraphAlignment.Right;
                    rect = new XRect(x + 240, y, 10, 10);
                    string tableau = trier[f]["Tableau"].ToString();
                    tf.DrawString(tableau, font8, XBrushes.Black, rect, XStringFormats.TopLeft);
                    y = y + 10;
                    NumeroLigne++;
                    if ((NumeroLigne > MaxLigne) && Col == 1)
                    {
                        Col = 2;
                        NumeroLigne = 0;
                        y = Top;
                    }
                    if ((NumeroLigne > MaxLigne) && Col == 2)
                    {
                        Col = 1;
                        NumeroLigne = 0;
                        string s = "";
                        if (NombrePage == 1) s = "I";
                        if (NombrePage == 2) s = "II";
                        if (NombrePage == 3) s = "III";
                        if (NombrePage == 4) s = "IV";
                        if (NombrePage == 5) s = "V";
                        if (NombrePage == 6) s = "VI";
                        rect = new XRect(POUCE * 7.25 , POUCE * 10.5, 20, 10);
                        tf.DrawString(s, font8, XBrushes.Black, rect, XStringFormats.TopLeft);
                        NouvellePage(ref document, ref gfx, ref page, "P");
                        NombrePage++;
                        EnteteTableMatiere(ref document, ref gfx, ref page);
                        y = Top;
                    }
                }
            }
            string ss = "";
            if (NombrePage == 1) ss = "I";
            if (NombrePage == 2) ss = "II";
            if (NombrePage == 3) ss = "III";
            if (NombrePage == 4) ss = "IV";
            if (NombrePage == 5) ss = "V";
            if (NombrePage == 6) ss = "VI";
            rect = new XRect(POUCE * 7.25, POUCE * 10.5, 20, 10);
            tf.DrawString(ss, font8, XBrushes.Black, rect, XStringFormats.TopLeft);
            if (!(NombrePage % 2 == 0))
            {
                NouvellePage(ref document, ref gfx, ref page, "P");
            }
        }
        /// <summary>
        /// Valide les champde date
        /// </summary>

        /// <param name="date">
        /// Date du champ à vérifier 
        /// </param>

        /// <returns>
        /// Retourne Vrai si date valide ou contien l'un de ces mot et entre vers après avant autour
        /// </returns> 
        private bool    ValiderDate( string date)
        {
             if (date == "")
            { 
                return true;
            }
            date = date.ToLower();
            if (date.Contains("et")  || date.Contains("entre") || date.Contains("vers") || date.Contains("après") || date.Contains("avant") || date.Contains("autour"))
            {
                return true;
            }
            if (date.Length == 4)
            {
                if (int.TryParse(date, out int i))
                {
                    return true;
                }
                else
                {
                    return true;
                }
            }
            if (date.Length == 7)
            {
                string str2 = date + "-01";
                DateTime.TryParseExact(str2, "yyyy-MM-dd", null, System.Globalization.DateTimeStyles.None, out DateTime a);
                if (a == DateTime.MinValue)
                {
                    return false;
                }
                else
                {
                    return true;
                }
            }
            if (date.Length == 10)
            {
                DateTime.TryParseExact(date, "yyyy-MM-dd", null, System.Globalization.DateTimeStyles.None, out DateTime a);
                if (a == DateTime.MinValue)
                {
                    return false;
                }
                else
                {
                    return true;
                }
            }
            return false;
        }
        //  Function ************************************************************************************************************************** Fin
        /// <summary>
        /// 
        /// </summary>
        public          TableauAscendant(string a)
        {
            argument = a;
            InitializeComponent();
        }
        private void    Form1_Load(object sender, EventArgs e)
        {
            DoubleBuffered = true;
            string Ligne = "";
            EffacerData();
            DateTime Maintenant = DateTime.Today;
            DateLb.Text = Maintenant.ToString("yyyy/MM/dd");
            
            string fichier = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData) + "\\TableauAscendant\\TableauAscendant.ini";
            if (File.Exists(fichier))
            {
                try
                {
                    using (StreamReader sr = File.OpenText(fichier))
                    {
                        while ((Ligne = sr.ReadLine()) != null)
                        {
                            if (Ligne == "[DossierPDF]")
                            {
                                DossierPDF = sr.ReadLine();
                            }
                            if (Ligne == "[CouleurBloc]")
                            {
                                byte r = Convert.ToByte(sr.ReadLine());
                                byte g = Convert.ToByte(sr.ReadLine());
                                byte b = Convert.ToByte(sr.ReadLine());

                                Color c = Color.FromArgb(r, g, b);
                                ChangerCouleurBloc(c);
                            }

                        }
                    }
                    if (DossierPDF != "")
                    {
                        DossierPDFToolStripMenuItem.Text = "D&ossier PDF -> " + DossierPDF;
                    }
                }
                catch (Exception msg)
                {
                    SystemSounds.Beep.Play();
                    MessageBox.Show("Ne peut pas lire les paramètres.\r\n\r\n" + msg.Message, "Problème ?",
                                     MessageBoxButtons.OK,
                                     MessageBoxIcon.Warning);
                }
                if (LOGACTIF)
                {
                    try
                    {
                        fichier = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + "\\" + FICHIERGRILLE;
                        if (File.Exists(fichier))
                        {
                            File.Delete(fichier);
                        }
                        fichier = DossierPDF + "\\" + FICHIERGRILLE;
                        if (File.Exists(fichier))
                        {
                            File.Delete(fichier);
                        }


                        fichier = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + "\\" + FICHIERLOG;
                        if (File.Exists(fichier))
                        {
                            File.Delete(fichier);
                        }
                        fichier = DossierPDF + "\\" + FICHIERLOG;
                        if (File.Exists(fichier))
                        {
                            File.Delete(fichier);
                        }



                        if (Directory.Exists(DossierPDF))
                        {
                            fichier = fichier = DossierPDF + "\\" + FICHIERLOG;
                        }
                        else
                        {
                            fichier = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + "\\" + FICHIERLOG;
                        }
                        using (StreamWriter ligne = File.AppendText(fichier))
                        {
                            ligne.WriteLine("************************************************");
                            ligne.WriteLine("  Log de TableauAscendant");
                            ligne.WriteLine("  " + "Version " + Application.ProductVersion + "B");
                            ligne.WriteLine("  " + DateTime.Now);
                            RegistryKey registryKey = Registry.LocalMachine.OpenSubKey(@"SOFTWARE\Microsoft\Windows NT\CurrentVersion");
                            string ProductName = registryKey.GetValue("ProductName").ToString(); //Windows Home
                            string ReleaseId = registryKey.GetValue("ReleaseId").ToString(); //1809
                            string CurrentBuild = registryKey.GetValue("CurrentBuild").ToString(); //17763
                            string buildNumber = registryKey.GetValue("UBR").ToString(); //316
                            ligne.WriteLine("  " + ProductName  + " " + ReleaseId + " " +  CurrentBuild + "." + buildNumber); 
                            ligne.WriteLine("************************************************");
                        }
                    }
                    catch { }
                }
            }
            ChoixSosaComboBox.Text  = "";
            GenerationAlb.Text = "";
            GenerationBlb.Text = "";
            GenerationClb.Text = "";
            AfficherData();
            if (argument != "")
            {
                FichierCourant = argument;
                LireData();
            }
            FlecheGaucheRechercheButton.Visible = false;
            FlecheDroiteRechercheButton.Visible = false ;

            // pour le développement
            // FichierTest();
        }
        private void    TableauAscendant_Paint(object sender, PaintEventArgs e)
        {
            Pen p = new Pen(Color.Black, 5)
            {
                EndCap = LineCap.ArrowAnchor
            };
            e.Graphics.DrawLine(p, 140, 340, 140, 383);

            //flèche droite

            Pen arrowPen = new Pen(Color.Black, 2);
        }
        private void    Sosa1PatronymeTextBox_TextChanged(object sender, EventArgs e)
        {
            liste[sosaCourant, PATRONYME] = Sosa1PatronymeTextBox.Text;
            if (!LongeurNomtOk(liste[sosaCourant, PATRONYME] + " " + liste[sosaCourant, PRENOM]))
            {
                Sosa1PatronymeTextBox.BackColor = couleurTextTropLong;
                Sosa1PrenomTextBox.BackColor = couleurTextTropLong;
            }
            else
            {
                Sosa1PatronymeTextBox.BackColor = couleurChamp;
                Sosa1PrenomTextBox.BackColor = couleurChamp;
            }
        }
        private void    Sosa1PatronymeTextBox_KeyPress(object sender, KeyPressEventArgs e)
        {
            Modifier = true;
            this.Text = NomPrograme + "   *" + FichierCourant;
        }
        private void    Sosa1PrenomTextBox_TextChanged(object sender, EventArgs e)
        {
            liste[sosaCourant, PRENOM] = Sosa1PrenomTextBox.Text;
            if (!LongeurNomtOk(liste[sosaCourant, PATRONYME] + " " + liste[sosaCourant, PRENOM]))
            {
                Sosa1PatronymeTextBox.BackColor = couleurTextTropLong;
                Sosa1PrenomTextBox.BackColor = couleurTextTropLong;
            }
            else
            {
                Sosa1PatronymeTextBox.BackColor = couleurChamp;
                Sosa1PrenomTextBox.BackColor = couleurChamp;
            }
        }
        private void    Sosa1PrenomTextBox_KeyPress(object sender, KeyPressEventArgs e)
        {
            Modifier = true;
            this.Text = NomPrograme + "   *" + FichierCourant;
        }
        private void    Sosa1NeTextBox_TextChanged(object sender, EventArgs e)
        {
            liste[sosaCourant, NELE] = Sosa1NeTextBox.Text;
            bool rep  = ValiderDate(Sosa1NeTextBox.Text);
            if (rep)
            {
                Sosa1NeTextBox.BackColor = Color.White;
                if (!LongeurTextOk(liste[sosaCourant, NELE]))
                {
                    Sosa1NeTextBox.BackColor = couleurTextTropLong;
                }
                else
                {
                    Sosa1NeTextBox.BackColor = couleurChamp;
                }
            }
            else
            {
                Sosa1NeTextBox.BackColor = Color.Red;
            }
        }
        private void    Sosa1NeTextBox_KeyPress(object sender, KeyPressEventArgs e)
        {
            Modifier = true;
            this.Text = NomPrograme + "   *" + FichierCourant;
        }
        private void    Sosa1NeEndroitTextBox_TextChanged(object sender, EventArgs e)
        {
            liste[sosaCourant, NELIEU] = Sosa1NeEndroitTextBox.Text;
            if (!LongeurTextOk(liste[sosaCourant, NELIEU]))
            {
                Sosa1NeEndroitTextBox.BackColor = couleurTextTropLong;
            } else
            {
                Sosa1NeEndroitTextBox.BackColor = couleurChamp;
            }
        }
        private void    Sosa1NeEndroitTextBox_KeyPress(object sender, KeyPressEventArgs e)
        {
            Modifier = true;
            this.Text = NomPrograme + "   *" + FichierCourant;
        }
        private void    Sosa1DeTextBox_TextChanged(object sender, EventArgs e)
        {
            liste[sosaCourant, DELE] = Sosa1DeTextBox.Text;
            bool rep = ValiderDate(Sosa1DeTextBox.Text);
            if (rep)
            {
                Sosa1DeTextBox.BackColor = Color.White;
                if (!LongeurTextOk(liste[sosaCourant, DELE]))
                {
                    Sosa1DeTextBox.BackColor = couleurTextTropLong;
                }
                else
                {
                    Sosa1DeTextBox.BackColor = couleurChamp;
                }
            }
            else
            {
                Sosa1DeTextBox.BackColor = Color.Red;
            }
        }
        private void    Sosa1DeTextBox_KeyPress(object sender, KeyPressEventArgs e)
        {
            Modifier = true;
            this.Text = NomPrograme + "   *" + FichierCourant;
        }
        private void    Sosa1DeEndroitTextBox_TextChanged(object sender, EventArgs e)
        {
            liste[sosaCourant, DELIEU] = Sosa1DeEndroitTextBox.Text;
            if (!LongeurTextOk(liste[sosaCourant, DELIEU]))
            {
                Sosa1DeEndroitTextBox.BackColor = couleurTextTropLong;
            }
            else
            {
                Sosa1DeEndroitTextBox.BackColor = couleurChamp;
            }
        }
        private void    Sosa1DeEndroitTextBox_KeyPress(object sender, KeyPressEventArgs e)
        {
            Modifier = true;
            this.Text = NomPrograme + "   *" + FichierCourant;
        }
        private void Sosa1MaTextBox_TextChanged(object sender, EventArgs e)
        {
            liste[sosaCourant, MALE] = Sosa1MaTextBox.Text;
            bool rep = ValiderDate(Sosa1MaTextBox.Text);
            if (rep)
            {
                Sosa1MaTextBox.BackColor = Color.White;
                if (!LongeurTextOk(liste[sosaCourant, MALE]))
                {
                    Sosa1MaTextBox.BackColor = couleurTextTropLong;
                }
                else
                {
                    Sosa1MaTextBox.BackColor = couleurChamp;
                }
            }
            else
            {
                Sosa1MaTextBox.BackColor = Color.Red;
            }
        }

        private void Sosa1MaTextBox_KeyPress(object sender, KeyPressEventArgs e)
        {
            Modifier = true;
            this.Text = NomPrograme + "   *" + FichierCourant;
        }

        private void Sosa1MaEndroitTextBox_TextChanged(object sender, EventArgs e)
        {
            liste[sosaCourant, MALIEU] = Sosa1MaEndroitTextBox.Text;
            if (!LongeurTextOk(Sosa1MaEndroitTextBox.Text))
            {
                Sosa1MaEndroitTextBox.BackColor = couleurTextTropLong;
            }
            else
            {
                Sosa1MaEndroitTextBox.BackColor = couleurChamp;
            }
        }

        private void Sosa1MaEndroitTextBox_KeyPress(object sender, KeyPressEventArgs e)
        {
            Modifier = true;
            this.Text = NomPrograme + "   *" + FichierCourant;
        }
        private void    Sosa2PatronymeTextBox_TextChanged(object sender, EventArgs e)
        {
            liste[sosaCourant * 2, PATRONYME] = Sosa2PatronymeTextBox.Text;
            Sosa2PatronymeTextBox.BackColor = Color.White;
            if (!LongeurNomtOk(liste[sosaCourant * 2, PATRONYME] + " " + liste[sosaCourant * 2, PRENOM]))
            {
                Sosa2PatronymeTextBox.BackColor = couleurTextTropLong;
                Sosa2PrenomTextBox.BackColor = couleurTextTropLong;
            }
            else
            {
                Sosa2PatronymeTextBox.BackColor = couleurChamp;
                Sosa2PrenomTextBox.BackColor = couleurChamp;
            }
        }
        private void    Sosa2PatronymeTextBox_KeyPress(object sender, KeyPressEventArgs e)
        {
            Modifier = true;
            this.Text = NomPrograme + "   *" + FichierCourant;
        }
        private void    Sosa2PrenomTextBox_TextChanged(object sender, EventArgs e)
        {
            liste[sosaCourant * 2, PRENOM] = Sosa2PrenomTextBox.Text;
            if (!LongeurNomtOk(liste[sosaCourant * 2, PRENOM] + " " + liste[sosaCourant * 2, PATRONYME]))
            {
                Sosa2PrenomTextBox.BackColor = couleurTextTropLong;
                Sosa2PatronymeTextBox.BackColor = couleurTextTropLong;
            }
            else
            {
                Sosa2PrenomTextBox.BackColor = couleurChamp;
                Sosa2PatronymeTextBox.BackColor = couleurChamp;
            }
        }
        private void    Sosa2PrenomTextBox_KeyPress(object sender, KeyPressEventArgs e)
        {
            Modifier = true;
            this.Text = NomPrograme + "   *" + FichierCourant;
        }
        private void    Sosa2NeTextBox_TextChanged(object sender, EventArgs e)
        {
            liste[sosaCourant * 2, NELE] = Sosa2NeTextBox.Text;
            bool rep = ValiderDate(Sosa2NeTextBox.Text);
            if (rep)
            {
                Sosa2NeTextBox.BackColor = Color.White;
                if (!LongeurTextOk(liste[sosaCourant * 2, NELE]))
                {
                    Sosa2NeTextBox.BackColor = couleurTextTropLong;
                }
                else
                {
                    Sosa2NeTextBox.BackColor = couleurChamp;
                }
            }
            else
            {
                Sosa2NeTextBox.BackColor = Color.Red;
            }
            
        }
        private void    Sosa2NeTextBox_KeyPress(object sender, KeyPressEventArgs e)
        {
            Modifier = true;
            this.Text = NomPrograme + "   *" + FichierCourant;
        }
        private void    Sosa2NeEndroitTextBox_TextChanged(object sender, EventArgs e)
        {
            liste[sosaCourant * 2, NELIEU] = Sosa2NeEndroitTextBox.Text;
            if (!LongeurTextOk(Sosa2NeEndroitTextBox.Text))
            {
                Sosa2NeEndroitTextBox.BackColor = couleurTextTropLong;
            }
            else
            {
                Sosa2NeEndroitTextBox.BackColor = couleurChamp;
            }
        }
        private void    Sosa2NeEndroitTextBox_KeyPress(object sender, KeyPressEventArgs e)
        {
            Modifier = true;
            this.Text = NomPrograme + "   *" + FichierCourant;
        }
        private void    Sosa2DeTextBox_TextChanged(object sender, EventArgs e)
        {
            liste[sosaCourant * 2, DELE] = Sosa2DeTextBox.Text;
            bool rep = ValiderDate(Sosa2DeTextBox.Text);
            if (rep)
            {
                Sosa2DeTextBox.BackColor = Color.White;
                if (!LongeurTextOk(liste[sosaCourant, DELE]))
                {
                    Sosa2DeTextBox.BackColor = couleurTextTropLong;
                }
                else
                {
                    Sosa2DeTextBox.BackColor = couleurChamp;
                }
            }
            else
            {
                Sosa2DeTextBox.BackColor = Color.Red;
            }
        }
        private void    Sosa2DeTextBox_KeyPress(object sender, KeyPressEventArgs e)
        {
            Modifier = true;
            this.Text = NomPrograme + "   *" + FichierCourant;
        }
        private void    Sosa2DeEndroitTextBox_TextChanged(object sender, EventArgs e)
        {
            liste[sosaCourant * 2, DELIEU] = Sosa2DeEndroitTextBox.Text;
            if (!LongeurTextOk(Sosa2DeEndroitTextBox.Text))
            {
                Sosa2DeEndroitTextBox.BackColor = couleurTextTropLong;
            }
            else
            {
                Sosa2DeEndroitTextBox.BackColor = couleurChamp;
            }
        }
        private void    Sosa2DeEndroitTextBox_KeyPress(object sender, KeyPressEventArgs e)
        {
            Modifier = true;
            this.Text = NomPrograme + "   *" + FichierCourant;
        }
        private void    Sosa23MaTextBox_TextChanged(object sender, EventArgs e)
        {
            liste[sosaCourant * 2, MALE] = Sosa23MaTextBox.Text;
            bool rep = ValiderDate(Sosa23MaTextBox.Text);
            if (rep)
            {
                Sosa23MaTextBox.BackColor = Color.White;
                if (!LongeurTextOk(liste[sosaCourant * 2, MALE]))
                {
                    Sosa23MaTextBox.BackColor = couleurTextTropLong;
                }
                else
                {
                    Sosa23MaTextBox.BackColor = couleurChamp;
                }
            }
            else
            {
                Sosa23MaTextBox.BackColor = Color.Red;
            }
        }
        private void    Sosa23MaTextBox_KeyPress(object sender, KeyPressEventArgs e)
        {
            Modifier = true;
            this.Text = NomPrograme + "   *" + FichierCourant;
        }
        private void    Sosa23MaEndroitTextBox_TextChanged(object sender, EventArgs e)
        {
            liste[sosaCourant * 2, MALIEU] = Sosa23MaEndroitTextBox.Text;
            if (!LongeurTextOk(Sosa23MaEndroitTextBox.Text))
            {
                Sosa23MaEndroitTextBox.BackColor = couleurTextTropLong;
            }
            else
            {
                Sosa23MaEndroitTextBox.BackColor = couleurChamp;
            }
        }
        private void    Sosa23MaEndroitBox_KeyPress(object sender, KeyPressEventArgs e)
        {
            Modifier = true;
            this.Text = NomPrograme + "   *" + FichierCourant;
        }
        private void    Sosa3PatronymeTextBox_TextChanged(object sender, EventArgs e)
        {
            liste[sosaCourant * 2 + 1, PATRONYME] = Sosa3PatronymeTextBox.Text;
            if (!LongeurNomtOk(liste[sosaCourant * 2 + 1, PATRONYME] + " " + liste[sosaCourant * 2 + 1, PRENOM]))
            {
                Sosa3PatronymeTextBox.BackColor = couleurTextTropLong;
                Sosa3PrenomTextBox.BackColor = couleurTextTropLong;
            }
            else
            {
                Sosa3PatronymeTextBox.BackColor = couleurChamp;
                Sosa3PrenomTextBox.BackColor = couleurChamp;
            }
        }
        private void    Sosa3PatronymeTextBox_KeyPress(object sender, KeyPressEventArgs e)
        {
            Modifier = true;
            this.Text = NomPrograme + "   *" + FichierCourant;
        }
        private void    Sosa3PrenomTextBox_TextChanged(object sender, EventArgs e)
        {
            liste[sosaCourant * 2 + 1, PRENOM] = Sosa3PrenomTextBox.Text;
            Sosa3PrenomTextBox.BackColor = Color.White;
            if (!LongeurNomtOk(liste[sosaCourant * 2 + 1, PRENOM] + " " + liste[sosaCourant * 2 + 1, PATRONYME]))
            {
                Sosa3PrenomTextBox.BackColor = couleurTextTropLong;
                Sosa3PatronymeTextBox.BackColor = couleurTextTropLong;
            }
            else
            {
                Sosa3PrenomTextBox.BackColor = couleurChamp;
                Sosa3PatronymeTextBox.BackColor = couleurChamp;
            }
        }
        private void    Sosa3PrenomTextBox_KeyPress(object sender, KeyPressEventArgs e)
        {
            Modifier = true;
            this.Text = NomPrograme + "   *" + FichierCourant;
        }
        private void    Sosa3NeTextBox_TextChanged(object sender, EventArgs e)
        {
            liste[sosaCourant * 2 + 1, NELE] = Sosa3NeTextBox.Text;
            bool rep = ValiderDate(Sosa3NeTextBox.Text);
            if (rep)
            {
                Sosa3NeTextBox.BackColor = Color.White;
                if (!LongeurTextOk(liste[sosaCourant * 2 + 1, NELE]))
                {
                    Sosa3NeTextBox.BackColor = couleurTextTropLong;
                }
                else
                {
                    Sosa3NeTextBox.BackColor = couleurChamp;
                }
            }
            else
            {
                Sosa3NeTextBox.BackColor = Color.Red;
            }
        }
        private void    Sosa3NeTextBox_KeyPress(object sender, KeyPressEventArgs e)
        {
            Modifier = true;
            this.Text = NomPrograme + "   *" + FichierCourant;
        }
        private void    Sosa3NeEndroitTextBox_TextChanged(object sender, EventArgs e)
        {
            liste[sosaCourant * 2 + 1, NELIEU] = Sosa3NeEndroitTextBox.Text;
            if (!LongeurTextOk(Sosa3NeEndroitTextBox.Text))
            {
                Sosa3NeEndroitTextBox.BackColor = couleurTextTropLong;
            }
            else
            {
                Sosa3NeEndroitTextBox.BackColor = couleurChamp;
            }
        }
        private void    Sosa3NeEndroitTextBox_KeyPress(object sender, KeyPressEventArgs e)
        {
            Modifier = true;
            this.Text = NomPrograme + "   *" + FichierCourant;
        }
        private void    Sosa3DeTextBox_TextChanged(object sender, EventArgs e)
        {
            liste[sosaCourant * 2 + 1, DELE] = Sosa3DeTextBox.Text;
            bool rep = ValiderDate(Sosa3DeTextBox.Text);
            if (rep)
            {
                Sosa3DeTextBox.BackColor = Color.White;
                if (!LongeurTextOk(liste[sosaCourant * 2 + 1, DELE]))
                {
                    Sosa3DeTextBox.BackColor = couleurTextTropLong;
                }
                else
                {
                    Sosa3DeTextBox.BackColor = couleurChamp;
                }
            }
            else
            {
                Sosa3DeTextBox.BackColor = Color.Red;
            }
        }
        private void    Sosa3DeTextBox_KeyPress(object sender, KeyPressEventArgs e)
        {
            Modifier = true;
            this.Text = NomPrograme + "   *" + FichierCourant;
        }
        private void    Sosa3DeEndroitTextBox_TextChanged(object sender, EventArgs e)
        {
            liste[sosaCourant * 2 + 1, DELIEU] = Sosa3DeEndroitTextBox.Text;
            if (!LongeurTextOk(Sosa3DeEndroitTextBox.Text))
            {
                Sosa3DeEndroitTextBox.BackColor = couleurTextTropLong;
            }
            else
            {
                Sosa3DeEndroitTextBox.BackColor = couleurChamp;
            }
        }
        private void    Sosa3DeEndroitTextBox_KeyPress(object sender, KeyPressEventArgs e)
        {
            Modifier = true;
            this.Text = NomPrograme + "   *" + FichierCourant;
        }
        private void    Sosa4PatronymeTextBox_TextChanged(object sender, EventArgs e)
        {
            liste[sosaCourant * 4, PATRONYME] = Sosa4PatronymeTextBox.Text;
            if (!LongeurNomtOk(liste[sosaCourant * 4, PATRONYME] + " " + liste[sosaCourant * 4, PRENOM]))
            {
                Sosa4PatronymeTextBox.BackColor = couleurTextTropLong;
                Sosa4PrenomTextBox.BackColor = couleurTextTropLong;
            }
            else
            {
                Sosa4PatronymeTextBox.BackColor = couleurChamp;
                Sosa4PrenomTextBox.BackColor = couleurChamp;
            }
        }
        private void    Sosa4PatronymeTextBox_KeyPress(object sender, KeyPressEventArgs e)
        {
            Modifier = true;
            this.Text = NomPrograme + "   *" + FichierCourant;
        }
        private void    Sosa4PrenomTextBox_TextChanged(object sender, EventArgs e)
        {
            liste[sosaCourant * 4, PRENOM] = Sosa4PrenomTextBox.Text;
            Sosa4PrenomTextBox.BackColor = Color.White;
            if (!LongeurNomtOk(liste[sosaCourant * 4, PRENOM] + " " + liste[sosaCourant * 4, PATRONYME]))
            {
                Sosa4PrenomTextBox.BackColor = couleurTextTropLong;
                Sosa4PatronymeTextBox.BackColor = couleurTextTropLong;
            }
            else
            {
                Sosa4PrenomTextBox.BackColor = couleurChamp;
                Sosa4PatronymeTextBox.BackColor = couleurChamp;
            }
        }
        private void    Sosa4PrenomTextBox_KeyPress(object sender, KeyPressEventArgs e)
        {
            Modifier = true;
            this.Text = NomPrograme + "   *" + FichierCourant;
        }
        private void    Sosa4NeTextBox_TextChanged(object sender, EventArgs e)
        {
            liste[sosaCourant * 4, NELE] = Sosa4PrenomTextBox.Text;
            bool rep = ValiderDate(Sosa4PrenomTextBox.Text);
            if (rep)
            {
                Sosa4PrenomTextBox.BackColor = Color.White;
                if (!LongeurTextOk(liste[sosaCourant * 4, NELE]))
                {
                    Sosa4PrenomTextBox.BackColor = couleurTextTropLong;
                }
                else
                {
                    Sosa4PrenomTextBox.BackColor = couleurChamp;
                }
            }
            else
            {
                Sosa4PrenomTextBox.BackColor = Color.Red;
            }
        }
        private void    Sosa4NeTextBox_KeyPress(object sender, KeyPressEventArgs e)
        {
            Modifier = true;
            this.Text = NomPrograme + "   *" + FichierCourant;
        }
        private void    Sosa4NeEndroitTextBox_TextChanged(object sender, EventArgs e)
        {
            liste[sosaCourant * 4, NELIEU] = Sosa4NeEndroitTextBox.Text;
            if (!LongeurTextOk(Sosa4NeEndroitTextBox.Text))
            {
                Sosa4NeEndroitTextBox.BackColor = couleurTextTropLong;
            }
            else
            {
                Sosa4NeEndroitTextBox.BackColor = couleurChamp;
            }
        }
        private void    Sosa4NeEndroitTextBox_KeyPress(object sender, KeyPressEventArgs e)
        {
            Modifier = true;
            this.Text = NomPrograme + "   *" + FichierCourant;
        }
        private void    Sosa4DeTextBox_TextChanged(object sender, EventArgs e)
        {
            liste[sosaCourant * 4, DELE] = Sosa4DeTextBox.Text;
            bool rep = ValiderDate(Sosa4DeTextBox.Text);
            if (rep)
            {
                Sosa4DeTextBox.BackColor = Color.White;
                if (!LongeurTextOk(liste[sosaCourant * 4, DELE]))
                {
                    Sosa4DeTextBox.BackColor = couleurTextTropLong;
                }
                else
                {
                    Sosa4DeTextBox.BackColor = couleurChamp;
                }
            }
            else
            {
                Sosa4DeTextBox.BackColor = Color.Red;
            }
        }
        private void    Sosa4DeTextBox_KeyPress(object sender, KeyPressEventArgs e)
        {
            Modifier = true;
            this.Text = NomPrograme + "   *" + FichierCourant;
        }
        private void    Sosa4DeEndroitTextBox_TextChanged(object sender, EventArgs e)
        {
            liste[sosaCourant * 4, DELIEU] = Sosa4DeEndroitTextBox.Text;
            if (!LongeurTextOk(Sosa4DeEndroitTextBox.Text))
            {
                Sosa4DeEndroitTextBox.BackColor = couleurTextTropLong;
            }
            else
            {
                Sosa4DeEndroitTextBox.BackColor = couleurChamp;
            }
        }
        private void    Sosa4DeEndroitTextBox_KeyPress(object sender, KeyPressEventArgs e)
        {
            Modifier = true;
            this.Text = NomPrograme + "   *" + FichierCourant;
        }
        private void    Sosa45MaTextBox_TextChanged(object sender, EventArgs e)
        {
            liste[sosaCourant * 4, MALE] = Sosa45MaTextBox.Text;
            bool rep = ValiderDate(Sosa45MaTextBox.Text);
            if (rep)
            {
                Sosa45MaTextBox.BackColor = Color.White;
                if (!LongeurTextOk(liste[sosaCourant * 4, MALE]))
                {
                    Sosa45MaTextBox.BackColor = couleurTextTropLong;
                }
                else
                {
                    Sosa45MaTextBox.BackColor = couleurChamp;
                }
            }
            else
            {
                Sosa45MaTextBox.BackColor = Color.Red;
            }
            EnregisterGrille();
        }
        private void    Sosa45MaTextBox_KeyPress(object sender, KeyPressEventArgs e)
        {
            Modifier = true;
            this.Text = NomPrograme + "   *" + FichierCourant;
        }
        private void    Sosa45MaEndroitTextBox_TextChanged(object sender, EventArgs e)
        {
            liste[sosaCourant * 4, MALIEU] = Sosa45MaLEndroitTextBox.Text;
            if (!LongeurTextOk(Sosa45MaLEndroitTextBox.Text))
            {
                Sosa45MaLEndroitTextBox.BackColor = couleurTextTropLong;
            }
            else
            {
                Sosa45MaLEndroitTextBox.BackColor = couleurChamp;
            }
        }
        private void    Sosa45MaEndroitNomTextBox_KeyPress(object sender, KeyPressEventArgs e)
        {
            Modifier = true;
            this.Text = NomPrograme + "   *" + FichierCourant;
        }
        private void    Sosa5PatronymeTextBox_TextChanged(object sender, EventArgs e)
        {
            liste[sosaCourant * 4 + 1, PATRONYME] = Sosa5PatronymeTextBox.Text;
            if (!LongeurNomtOk(liste[sosaCourant * 4 + 1, PATRONYME] + " " + liste[sosaCourant * 4 + 1, PRENOM]))
            {
                Sosa5PatronymeTextBox.BackColor = couleurTextTropLong;
                Sosa5PrenomTextBox.BackColor = couleurTextTropLong;
            }
            else
            {
                Sosa5PatronymeTextBox.BackColor = couleurChamp;
                Sosa5PrenomTextBox.BackColor = couleurChamp;
            }
        }
        private void    Sosa5PatronymeTextBox_KeyPress(object sender, KeyPressEventArgs e)
        {
            Modifier = true;
            this.Text = NomPrograme + "   *" + FichierCourant;
        }
        private void    Sosa5PrenomTextBox_TextChanged(object sender, EventArgs e)
        {
            liste[sosaCourant * 4 + 1, PRENOM] = Sosa5PrenomTextBox.Text;
            Sosa5PrenomTextBox.BackColor = Color.White;
            if (!LongeurNomtOk(liste[sosaCourant * 4 + 1, PRENOM] + " " + liste[sosaCourant * 4 + 1, PATRONYME]))
            {
                Sosa5PrenomTextBox.BackColor = couleurTextTropLong;
                Sosa5PatronymeTextBox.BackColor = couleurTextTropLong;
            }
            else
            {
                Sosa5PrenomTextBox.BackColor = couleurChamp;
                Sosa5PatronymeTextBox.BackColor = couleurChamp;
            }
        }
        private void    Sosa5PrenomTextBox_KeyPress(object sender, KeyPressEventArgs e)
        {
            Modifier = true;
            this.Text = NomPrograme + "   *" + FichierCourant;
        }

        private void    Sosa5NeTextBox_TextChanged(object sender, EventArgs e)
        {
            liste[sosaCourant * 4 + 1, NELE] = Sosa5NeTextBox.Text;
            bool rep = ValiderDate(Sosa5NeTextBox.Text);
            if (rep)
            {
                Sosa5NeTextBox.BackColor = Color.White;
                if (!LongeurTextOk(liste[sosaCourant * 4 + 1, NELE]))
                {
                    Sosa5NeTextBox.BackColor = couleurTextTropLong;
                }
                else
                {
                    Sosa5DeTextBox.BackColor = couleurChamp;
                }
            }
            else
            {
                Sosa5NeTextBox.BackColor = Color.Red;
            }
        }
        private void    Sosa5NeTextBox_KeyPress(object sender, KeyPressEventArgs e)
        {
            Modifier = true;
            this.Text = NomPrograme + "   *" + FichierCourant;
        }
        private void    Sosa5NeEndroitTextBox_TextChanged(object sender, EventArgs e)
        {
            liste[sosaCourant * 4 + 1, NELIEU] = Sosa5NeEndroitTextBox.Text;
            if (!LongeurTextOk(Sosa5NeEndroitTextBox.Text))
            {
                Sosa5NeEndroitTextBox.BackColor = couleurTextTropLong;
            }
            else
            {
                Sosa5NeEndroitTextBox.BackColor = couleurChamp;
            }
        }
        private void    Sosa5NeEndroit1NomTextBox_KeyPress(object sender, KeyPressEventArgs e)
        {
            Modifier = true;
            this.Text = NomPrograme + "   *" + FichierCourant;
        }
        private void    Sosa5DeTextBox_TextChanged(object sender, EventArgs e)
        {
            {
                liste[sosaCourant * 4 + 1, DELE] = Sosa5DeTextBox.Text;
                bool rep = ValiderDate(Sosa5DeTextBox.Text);
                if (rep)
                {
                    Sosa5DeTextBox.BackColor = Color.White;
                    if (!LongeurTextOk(liste[sosaCourant * 4 + 1, DELE]))
                    {
                        Sosa5DeTextBox.BackColor = couleurTextTropLong;
                    }
                    else
                    {
                        Sosa5DeTextBox.BackColor = couleurChamp;
                    }
                }
                else
                {
                    Sosa5DeTextBox.BackColor = Color.Red;
                }
            }
        }
        private void    Sosa5DeTextBox_KeyPress(object sender, KeyPressEventArgs e)
        {
            Modifier = true;
            this.Text = NomPrograme + "   *" + FichierCourant;
        }
        private void    Sosa5DeEndroitTextBox_TextChanged(object sender, EventArgs e)
        {
            liste[sosaCourant * 4 + 1, DELIEU] = Sosa5DeEndroitTextBox.Text;
            if (!LongeurTextOk(Sosa5DeEndroitTextBox.Text))
            {
                Sosa5DeEndroitTextBox.BackColor = couleurTextTropLong;
            }
            else
            {
                Sosa5DeEndroitTextBox.BackColor = couleurChamp;
            }
        }
        private void    Sosa5DeEndroitTextBox_KeyPress(object sender, KeyPressEventArgs e)
        {
            Modifier = true;
            this.Text = NomPrograme + "   *" + FichierCourant;
        }
        private void    Sosa6PatronymeTextBox_TextChanged(object sender, EventArgs e)
        {
            liste[sosaCourant * 4 + 2, PATRONYME] = Sosa6PatronymeTextBox.Text;
            if (!LongeurNomtOk(liste[sosaCourant * 4 + 2, PATRONYME] + " " + liste[sosaCourant * 4 + 2, PRENOM]))
            {
                Sosa6PatronymeTextBox.BackColor = couleurTextTropLong;
                Sosa6PrenomTextBox.BackColor = couleurTextTropLong;
            }
            else
            {
                Sosa6PatronymeTextBox.BackColor = couleurChamp;
                Sosa6PrenomTextBox.BackColor = couleurChamp;
            }
        }
        private void    Sosa6PatronymeTextBox_KeyPress(object sender, KeyPressEventArgs e)
        {
            Modifier = true;
            this.Text = NomPrograme + "   *" + FichierCourant;
        }
        private void    Sosa6PrenomTextBox_TextChanged(object sender, EventArgs e)
        {
            liste[sosaCourant * 4 + 2, PRENOM] = Sosa6PrenomTextBox.Text;
            Sosa6PrenomTextBox.BackColor = Color.White;
            if (!LongeurNomtOk(liste[sosaCourant * 4 + 2, PRENOM] + " " + liste[sosaCourant * 4 + 2, PATRONYME]))
            {
                Sosa6PrenomTextBox.BackColor = couleurTextTropLong;
                Sosa6PatronymeTextBox.BackColor = couleurTextTropLong;
            }
            else
            {
                Sosa6PrenomTextBox.BackColor = couleurChamp;
                Sosa6PatronymeTextBox.BackColor = couleurChamp;
            }
        }
        private void    Sosa6PrenomTextBox_KeyPress(object sender, KeyPressEventArgs e)
        {
            Modifier = true;
            this.Text = NomPrograme + "   *" + FichierCourant;
        }
        private void    Sosa6NeTextBox_TextChanged(object sender, EventArgs e)
        {
            liste[sosaCourant * 4 + 2, NELE] = Sosa6NeTextBox.Text;
            bool rep = ValiderDate(Sosa6NeTextBox.Text);
            if (rep)
            {
                Sosa6NeTextBox.BackColor = Color.White;
                if (!LongeurTextOk(liste[sosaCourant * 4 + 2, NELE]))
                {
                    Sosa6NeTextBox.BackColor = couleurTextTropLong;
                }
                else
                {
                    Sosa6NeTextBox.BackColor = couleurChamp;
                }
            }
            else
            {
                Sosa6NeTextBox.BackColor = Color.Red;
            }
        }
        private void    Sosa6NeTextBox_KeyPress(object sender, KeyPressEventArgs e)
        {
            Modifier = true;
            this.Text = NomPrograme + "   *" + FichierCourant;
        }
        private void    Sosa6NeEndroitTextBox_TextChanged(object sender, EventArgs e)
        {
            liste[sosaCourant * 4 + 2, NELIEU] = Sosa6NeEndroitTextBox.Text;
            if (!LongeurTextOk(Sosa6NeEndroitTextBox.Text))
            {
                Sosa6NeEndroitTextBox.BackColor = couleurTextTropLong;
            }
            else
            {
                Sosa6NeEndroitTextBox.BackColor = couleurChamp;
            }
        }
        private void    Sosa6NeEndroitTextBox_KeyPress(object sender, KeyPressEventArgs e)
        {
            Modifier = true;
            this.Text = NomPrograme + "   *" + FichierCourant;
        }
        private void    Sosa6DeTextBox_TextChanged(object sender, EventArgs e)
        {
            liste[sosaCourant * 4 + 2, DELE] = Sosa6DeTextBox.Text;
            bool rep = ValiderDate(Sosa6DeTextBox.Text);
            if (rep)
            {
                Sosa6DeTextBox.BackColor = Color.White;
                if (!LongeurTextOk(liste[sosaCourant * 4 + 2, DELE]))
                {
                    Sosa6DeTextBox.BackColor = couleurTextTropLong;
                }
                else
                {
                    Sosa6DeTextBox.BackColor = couleurChamp;
                }
            }
            else
            {
                Sosa6DeTextBox.BackColor = Color.Red;
            }
        }
        private void    Sosa6DeTextBox_KeyPress(object sender, KeyPressEventArgs e)
        {
            Modifier = true;
            this.Text = NomPrograme + "   *" + FichierCourant;
        }
        private void    Sosa6DeEndroitTextBox_TextChanged(object sender, EventArgs e)
        {
            liste[sosaCourant * 4 + 2, DELIEU] = Sosa6DeEndroitTextBox.Text;
            if (!LongeurTextOk(Sosa6DeEndroitTextBox.Text))
            {
                Sosa6DeEndroitTextBox.BackColor = couleurTextTropLong;
            }
            else
            {
                Sosa6DeEndroitTextBox.BackColor = couleurChamp;
            }
        }
        private void    Sosa6DeEndroitTextBox_KeyPress(object sender, KeyPressEventArgs e)
        {
            Modifier = true;
            this.Text = NomPrograme + "   *" + FichierCourant;
        }
        private void    Sosa67MaTextBox_TextChanged(object sender, EventArgs e)
        {
            liste[sosaCourant * 4 + 2, MALE] = Sosa67MaTextBox.Text;
            bool rep = ValiderDate(Sosa67MaTextBox.Text);
            if (rep)
            {
                Sosa67MaTextBox.BackColor = Color.White;
                if (!LongeurTextOk(liste[sosaCourant * 2, MALE]))
                {
                    Sosa67MaTextBox.BackColor = couleurTextTropLong;
                }
                else
                {
                    Sosa67MaTextBox.BackColor = couleurChamp;
                }
            }
            else
            {
                Sosa67MaTextBox.BackColor = Color.Red;
            }
        }
        private void    Sosa67MaTextBox_KeyPress(object sender, KeyPressEventArgs e)
        {
            Modifier = true;
            this.Text = NomPrograme + "   *" + FichierCourant;
        }
        private void    Sosa67MaEndroitTextBox_TextChanged(object sender, EventArgs e)
        {
            liste[sosaCourant * 4 + 2, MALIEU] = Sosa67MaEndroitTextBox.Text;
            if (!LongeurTextOk(Sosa67MaEndroitTextBox.Text))
            {
                Sosa67MaEndroitTextBox.BackColor = couleurTextTropLong;
            }
            else
            {
                Sosa67MaEndroitTextBox.BackColor = couleurChamp;
            }
        }
        private void    Sosa67MAEndroitTextBox_KeyPress(object sender, KeyPressEventArgs e)
        {
            Modifier = true;
            this.Text = NomPrograme + "   *" + FichierCourant;
        }
        private void    Sosa7PatronymeTextBox_TextChanged(object sender, EventArgs e)
        {
            liste[sosaCourant * 4 + 3, PATRONYME] = Sosa7PatronymeTextBox.Text;
            if (!LongeurNomtOk(liste[sosaCourant * 4 + 3, PATRONYME] + " " + liste[sosaCourant * 4 + 3, PRENOM]))
            {
                Sosa7PatronymeTextBox.BackColor = couleurTextTropLong;
                Sosa7PrenomTextBox.BackColor = couleurTextTropLong;
            }
            else
            {
                Sosa7PatronymeTextBox.BackColor = couleurChamp;
                Sosa7PrenomTextBox.BackColor = couleurChamp;
            }
        }
        private void    Sosa7PatronymeTextBox_KeyPress(object sender, KeyPressEventArgs e)
        {
            Modifier = true;
            this.Text = NomPrograme + "   *" + FichierCourant;
        }
        private void    Sosa7PrenomTextBox_TextChanged(object sender, EventArgs e)
        {
            Sosa7PrenomTextBox.BackColor = Color.White;
            if (!LongeurNomtOk(liste[sosaCourant * 4 + 3, PRENOM] + " " + liste[sosaCourant * 4 + 3, PATRONYME]))
            {
                Sosa7PrenomTextBox.BackColor = couleurTextTropLong;
                Sosa7PatronymeTextBox.BackColor = couleurTextTropLong;
            }
            else
            {
                Sosa7PrenomTextBox.BackColor = couleurChamp;
                Sosa7PatronymeTextBox.BackColor = couleurChamp;
            }
        }
        private void    Sosa7PrenomTextBox_KeyPress(object sender, KeyPressEventArgs e)
        {
            Modifier = true;
            this.Text = NomPrograme + "   *" + FichierCourant;
        }
        private void    Sosa7NeTextBox_TextChanged(object sender, EventArgs e)
        {
            liste[sosaCourant * 4 + 3, NELE] = Sosa7NeTextBox.Text;
            bool rep = ValiderDate(Sosa7NeTextBox.Text);
            if (rep)
            {
                Sosa7NeTextBox.BackColor = Color.White;
                if (!LongeurTextOk(liste[sosaCourant * 4 + 3, NELE]))
                {
                    Sosa7NeTextBox.BackColor = couleurTextTropLong;
                }
                else
                {
                    Sosa7NeTextBox.BackColor = couleurChamp;
                }
            }
            else
            {
                Sosa7NeTextBox.BackColor = Color.Red;
            }
        }
        private void    Sosa7NeTextBox_KeyPress(object sender, KeyPressEventArgs e)
        {
            Modifier = true;
            this.Text = NomPrograme + "   *" + FichierCourant;
        }
        private void    Sosa7NeEndroitTextBox_TextChanged(object sender, EventArgs e)
        {
            liste[sosaCourant * 4 + 3, NELIEU] = Sosa7NeEndroitTextBox.Text;
            if (!LongeurTextOk(Sosa7NeEndroitTextBox.Text))
            {
                Sosa7NeEndroitTextBox.BackColor = couleurTextTropLong;
            }
            else
            {
                Sosa7NeEndroitTextBox.BackColor = couleurChamp;
            }
        }
        private void    Sosa7NeEndroitTextBox_KeyPress(object sender, KeyPressEventArgs e)
        {
            Modifier = true;
            this.Text = NomPrograme + "   *" + FichierCourant;
        }
        private void    Sosa7DeTextBox_TextChanged(object sender, EventArgs e)
        {
            liste[sosaCourant * 4 + 3, DELE] = Sosa7DeTextBox.Text;
            bool rep = ValiderDate(Sosa7DeTextBox.Text);
            if (rep)
            {
                Sosa7DeTextBox.BackColor = Color.White;
                if (!LongeurTextOk(liste[sosaCourant * 4 + 3, DELE]))
                {
                    Sosa7DeTextBox.BackColor = couleurTextTropLong;
                }
                else
                {
                    Sosa7DeTextBox.BackColor = couleurChamp;
                }
            }
            else
            {
                Sosa7DeTextBox.BackColor = Color.Red;
            }
        }
        private void    Sosa7DeTextBox_KeyPress(object sender, KeyPressEventArgs e)
        {
            Modifier = true;
            this.Text = NomPrograme + "   *" + FichierCourant;
        }
        private void    Sosa7DeEndroitTextBox_TextChanged(object sender, EventArgs e)
        {
            liste[sosaCourant * 4 + 3, DELIEU] = Sosa7DeEndroitTextBox.Text;
            if (!LongeurTextOk(Sosa7DeEndroitTextBox.Text))
            {
                Sosa7DeEndroitTextBox.BackColor = couleurTextTropLong;
            }
            else
            {
                Sosa7DeEndroitTextBox.BackColor = couleurChamp;
            }
        }
        private void    Sosa7DeEndroitTextBox_KeyPress(object sender, KeyPressEventArgs e)
        {
            Modifier = true;
            this.Text = NomPrograme + "   *" + FichierCourant;
        }
        private void EnregisterToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if(FichierCourant =="")
            {
                EnregistrerDataSous();
            } else
            {
                EnregistrerData();
            }
            
        }
        private void MenuMs_ItemClicked(object sender, ToolStripItemClickedEventArgs e)
        {

        }
        private void OuvrirToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (DataModifier())
            {

                ChoixPersonne.Visible = false;
                ChoixPersonne.Enabled = false;
                OpenFileDialog LireDialog = new OpenFileDialog
                {
                    Filter = "Fichier|*.tas",
                    Title = "Lire le fichier"
                };
                LireDialog.ShowDialog();

                if (LireDialog.FileName != "")
                {
                    FichierCourant = LireDialog.FileName;
                    LireData();
                    this.Text = NomPrograme + "   " + FichierCourant;
                    Modifier = false;

                }
            }
        }
        private void EnregistrerSousToolStripMenuItem_Click(object sender, EventArgs e)
        {
            EnregistrerDataSous();
        }
        private void NouveauToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (DataModifier())
            {
                ChoixPersonne.Visible = false;
                ChoixPersonne.Enabled = false;
                EffacerData();
                FichierCourant = "";
                this.Text = NomPrograme;
            }
        }
        private void CreerPageCouranteToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (DossierPDF == "")
            {
                AvoirDossierrapport();
            }
            int[] listePage = new int[] { 0, 1, 8, 9, 10, 11, 12, 13, 14, 15, 64, 65, 66, 67, 68, 69, 70, 71, 72, 73, 74, 75, 76, 77, 78, 79, 80, 81, 82, 83, 84, 85, 86, 87, 88, 89, 90, 91, 92, 93, 94, 95, 96, 97, 98, 99, 100, 101, 102, 103, 104, 105, 106, 107, 108, 109, 110, 111, 112, 113, 114, 115, 116, 117, 118, 119, 120, 121, 122, 123, 124, 125, 126, 127 };
            int sosa;
            if (ChoixSosaComboBox.Text == "")
            {
                sosa = 0;
            }
            else
            {
                try
                {
                    sosa = Int32.Parse(ChoixSosaComboBox.Text);
                    for (int i = 0; i < 74; i++)
                    {
                        if (sosa == listePage[i])
                        {
                            SosaChanger();
                            ChoixSosaComboBox.BackColor = Color.White;
                            break;
                        }
                    }
                }
                catch
                {
                    ChoixSosaComboBox.BackColor = Color.Gray;
                    return;
                }
            }
            PdfDocument document = new PdfDocument();
            document.Info.Title = "Tableau Ascendance";
            PdfPage page = document.AddPage();
            page.Size = PageSize.Letter;
            page.Orientation = PdfSharp.PageOrientation.Landscape;
            XGraphics gfx = XGraphics.FromPdfPage(page);
            //NouvellePage(ref document, ref gfx, ref page);
            string numeroTableau = DessinerPage(ref document, ref gfx, sosa, true, false);
            string FichierPage = numeroTableau.ToString() + ".pdf";
            if (FichierPage == "0.pdf" || FichierPage == ".pdf")
            {
                FichierPage = "vide.pdf";

            }
            string Fichier = DossierPDF + "\\" + FichierPage;
            try
            {
                document.Save(Fichier);
                Process.Start(Fichier);
            }
            catch (Exception p)
            {
                SystemSounds.Beep.Play();
                MessageBox.Show("Ne peut pas afficher la page.\r\n\r\n" + p.Message, "Problème ?",
                                 MessageBoxButtons.OK,
                                 MessageBoxIcon.Warning);
            }
        }
        private void CreerToutesLesPagesToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (DossierPDF == "")
            {
                AvoirDossierrapport();
            }
            int[] listePage = new int[] { 1, 8, 9, 10, 11, 12, 13, 14, 15, 64, 65, 66, 67, 68, 69, 70, 71, 72, 73, 74, 75, 76, 77, 78, 79, 80, 81, 82, 83, 84, 85, 86, 87, 88, 89, 90, 91, 92, 93, 94, 95, 96, 97, 98, 99, 100, 101, 102, 103, 104, 105, 106, 107, 108, 109, 110, 111, 112, 113, 114, 115, 116, 117, 118, 119, 120, 121, 122, 123, 124, 125, 126, 127 };
            XUnit xpouce = XUnit.FromInch(1);
            XFont font32 = new XFont("arial", 32, XFontStyle.Bold);
            XFont font18 = new XFont("arial", 18, XFontStyle.Bold);
            XPen pen = new XPen(XColor.FromArgb(0, 0, 0));
            XFont font8 = new XFont("arial", 8, XFontStyle.Bold);
            PdfDocument document = new PdfDocument();
            document.Info.Title = "Tabeau Ascendance";
            PdfPage page = document.AddPage();
            page.Size = PageSize.Letter;
            page.Orientation = PdfSharp.PageOrientation.Portrait;
            XGraphics gfx = XGraphics.FromPdfPage(page);
            XTextFormatter tm = new XTextFormatter(gfx);

            // cadre
            XImage img = global::TableauAscendant.Properties.Resources.cadre;
            gfx.DrawImage(img, 0, 0);

            string str = "Tableau asendant";
            XSize textLargeur = gfx.MeasureString(str, font32);
            gfx.DrawString(str, font32, XBrushes.Black, page.Width / 2 - textLargeur.Width / 2, POUCE * 4);
            str = "de";
            textLargeur = gfx.MeasureString(str, font32);
            gfx.DrawString(str, font32, XBrushes.Black, page.Width / 2 - textLargeur.Width / 2, POUCE * 5);
            if (AscendantDeTb.Text != "")
            {
                textLargeur = gfx.MeasureString(AscendantDeTb.Text, font32);
                tm.Alignment = XParagraphAlignment.Center;
                XRect rect = new XRect();
                rect = new XRect(POUCE * .83, POUCE * 6, POUCE * 6.84 , POUCE);
                tm.DrawString(AscendantDeTb.Text, font32, XBrushes.Black, rect);
            } else
            {
                str = "____________________________________";
                textLargeur = gfx.MeasureString(str, font18);
                gfx.DrawString(str, font18, XBrushes.Black, page.Width / 2 - textLargeur.Width / 2, POUCE * 6);
            }

            // Préparé par 
            str = "Préparé par";
            textLargeur = gfx.MeasureString(str, font18);
            gfx.DrawString(str, font18, XBrushes.Black, page.Width / 2 - textLargeur.Width / 2, POUCE * 8);
            if (PreparerPar.Text != "")
            {
                str = PreparerPar.Text;
                textLargeur = gfx.MeasureString(str, font18);
                gfx.DrawString(str, font18, XBrushes.Black, page.Width / 2 - textLargeur.Width / 2, POUCE * 8.3);

                str = DateLb.Text ;
                textLargeur = gfx.MeasureString(str, font18);
                gfx.DrawString(str, font18, XBrushes.Black, page.Width / 2 - textLargeur.Width / 2, POUCE * 8.6);
            }
            else
            {
                str = "____________________________________";
                textLargeur = gfx.MeasureString(str, font18);
                gfx.DrawString(str, font18, XBrushes.Black, page.Width / 2 - textLargeur.Width / 2, POUCE * 8.3);
                str = "______/___/___";
                textLargeur = gfx.MeasureString(str, font18);
                gfx.DrawString(str, font18, XBrushes.Black, page.Width / 2 - textLargeur.Width / 2, POUCE * 8.6);
            }

            NouvellePage(ref document, ref gfx, ref page, "P");
            NouvellePage(ref document, ref gfx, ref page, "P");

            
            TableMatiere(ref document, ref gfx, ref page);

            foreach (int sosa in listePage)
            {
                NouvellePage(ref document, ref gfx, ref page,"L");
                DessinerPage(ref document, ref gfx, sosa, true, true);
            }

            string FichierPage = "TableauAscendant.pdf";
            string Fichier = DossierPDF + "\\" + FichierPage;
            try
            {
                document.Save(Fichier);
                Process.Start(Fichier);
            }
            catch (Exception p)
            {
                SystemSounds.Beep.Play();
                MessageBox.Show("Ne peut pas afficher les pages.\r\n\r\n" + p.Message, "Problème ?",
                                 MessageBoxButtons.OK,
                                 MessageBoxIcon.Warning);
            }

        }
        private void PDFToolStripMenuItem_Click(object sender, EventArgs e)
        {
             
        }
        private void EffacerToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (Modifier)
            {
                DialogResult resulta = MessageBox.Show("Enregister avant ?", "Attention",
                        MessageBoxButtons.YesNoCancel, MessageBoxIcon.Warning);
                if (resulta == DialogResult.Yes)
                {
                    if (FichierCourant == "")
                    {
                        EnregistrerDataSous();
                    }
                    else
                    {
                        EnregistrerData();
                    }
                    EffacerData();
                    FichierCourant = "";
                    this.Text = NomPrograme;
                }
                else if (resulta == DialogResult.No)
                {
                    EffacerData();
                    FichierCourant = "";
                    return;
                }
                else if (resulta == DialogResult.Cancel)
                {
                    return;
                }
            }
            else
            {
                EffacerData();
                FichierCourant = "";
            }
        }
        private void DossierPDFToolStripMenuItem_Click(object sender, EventArgs e)
        {
            AvoirDossierrapport();
        }
        private void ChoixSosaComboBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            ChoixChanger();
        }
        private void ChoixSosaComboBox_SelectedValueChanged(object sender, EventArgs e)
        {
            ChoixChanger();

        }
        private void ChoixSosaComboBox_KeyPress(object sender, KeyPressEventArgs e)
        {
            ChoixChanger();

        }
        private void ChoixSosaComboBox_TextChanged(object sender, EventArgs e)
        {
            ChoixChanger();
        }
        private void PreparerPar_KeyPress(object sender, KeyPressEventArgs e)
        {
            Modifier = true;
            this.Text = NomPrograme + "   *" + FichierCourant;
        }
        private void TableauAscendant_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (Modifier)
            {
                DialogResult resulta = MessageBox.Show("Enregister avant ?", "Attention",
                        MessageBoxButtons.YesNoCancel, MessageBoxIcon.Warning);
                if (resulta == DialogResult.Yes)
                {
                    if (FichierCourant == "")
                    {
                        EnregistrerDataSous();
                    }
                    else
                    {
                        EnregistrerData();
                    }
                    //EffacerData();
                    //FichierCourant = "";
                    //this.Text = NomPrograme + "   " + FichierCourant;
                }
                else if (resulta == DialogResult.No)
                {
                    return;
                }
                else if (resulta == DialogResult.Cancel)
                {
                    e.Cancel = true;
                }
            }

            try
            {
                string Dossier = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData) + "\\TableauAscendant";
                DirectoryInfo di = Directory.CreateDirectory(@Dossier);
                string Fichier = Dossier + "\\TableauAscendant.ini";


                if (File.Exists(Fichier))
                {
                    File.Delete(Fichier);
                }
                using (StreamWriter ligne = File.CreateText(Fichier))
                {
                    ligne.WriteLine("[DossierPDF]");
                    ligne.WriteLine(DossierPDF);
                    ligne.WriteLine("[CouleurBloc]");
                    ligne.WriteLine(RectangleSosa1.FillColor.R.ToString());
                    ligne.WriteLine(RectangleSosa1.FillColor.G.ToString());
                    ligne.WriteLine(RectangleSosa1.FillColor.B.ToString());
                }
            }
            catch (Exception msg)
            {
                {
                    SystemSounds.Beep.Play();
                    MessageBox.Show("Ne peut pas écrire les paramètres.\r\n\r\n" + msg.Message, "Problème ?",
                                     MessageBoxButtons.OK,
                                     MessageBoxIcon.Warning);
                }
            }
        }
        private void VersionToolStripMenuItem_Click(object sender, EventArgs e)
        {

        }
        private void AideToolStripMenuItem_Click(object sender, EventArgs e)
        {
            
            
        }
        private void AideToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            Form l = new AideFm();
            l.Show();
        }
        private void VersionToolStripMenuItem_Click_1(object sender, EventArgs e)
        {
            Form l = new Form2();
            l.ShowDialog(this);
        }
        private void Button1_Click(object sender, EventArgs e)
        {

        }
        private void GoSosa1Btn_Click(object sender, EventArgs e)
        {
            double a = System.Convert.ToInt32(sosaCourant);
            a =   Math.Floor(a / 8);
            int b = System.Convert.ToInt32(a);
            ChoixSosaComboBox.Text = Convert.ToString(b) ;
        }
        private void GoSosa4Btn_Click(object sender, EventArgs e)
        {
            ChoixSosaComboBox.Text = GoSosa4Btn.Text; // changer
        }
        private void GoSosa5Btn_Click(object sender, EventArgs e)
        {
            ChoixSosaComboBox.Text = GoSosa5Btn.Text; // changer
        }
        private void GoSosa6Btn_Click(object sender, EventArgs e)
        {
            ChoixSosaComboBox.Text = GoSosa6Btn.Text; // changer
        }
        private void GoSosa7Btn_Click(object sender, EventArgs e)
        {
            ChoixSosaComboBox.Text = GoSosa7Btn.Text; // changer
        }
        private void QuitterToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (Modifier)
            {
                DialogResult resulta = MessageBox.Show("Enregister avant de Quitter?", "Attention",
                        MessageBoxButtons.YesNoCancel, MessageBoxIcon.Warning);
                if (resulta == DialogResult.Yes)
                {
                    if (FichierCourant == "")
                    {
                        EnregistrerDataSous();
                    }
                    else
                    {
                        EnregistrerData();
                    }
                    Modifier = false;
                    Application.Exit();
                }
                else if (resulta == DialogResult.No)
                {
                    Modifier = false;
                    Application.Exit();
                }
                else if (resulta == DialogResult.Cancel)
                {
                    return;
                }
            }
            else
            {
                Application.Exit();
            }
        }
        private void OuvrirFichierGEDCOMToolStripMenuItem_Click(object sender, EventArgs e)
        {
            FichierGEDCOM = "";
            DataModifier();
            EffacerData();
            OpenFileDialog LireDialog = new OpenFileDialog
            {
                Filter = "Fichier|*.ged",
                Title = "Lire le fichier"
            };
            LireDialog.ShowDialog();
            if (LireDialog.FileName != "")
            {
                FichierGEDCOM  = LireDialog.FileName;
                //LireData();
            }
            NomFichierGedcomLb.Text = Path.GetFileName(FichierGEDCOM);

            if (FichierGEDCOM != "")
            {
                GEDCOM.EffacerDataGEDCOM();
                GEDCOM.LireGEDCOM(FichierGEDCOM);
                GEDCOM.Individu();
                GEDCOM.Famille();

                ChoixLV.View = View.Details;
                ChoixLV.GridLines = true;
                ChoixLV.FullRowSelect = true;
                for (int f = ChoixLV.Items.Count - 1; f > 0; f--)
                {
                    ChoixLV.Items.RemoveAt(f);
                }
                ChoixLV.Items.Clear();
                ChoixLV.Columns.Add("ID", 50);
                ChoixLV.Columns.Add("Nom", 200);
                ChoixLV.Columns.Add("Naissance", 58);
                ChoixLV.Columns.Add("Décès", 58);
                ChoixPersonne.Visible = true;
                ChoixPersonne.Enabled = true;
                FichierCourant = "";
                this.Text = NomPrograme;

            }
            
        }
        private void Recherche_Click(object sender, EventArgs e)
        {
            RechercheID();
        }
        private void ChoixLV_SelectedIndexChanged(object sender, EventArgs e)
        {
            ContinuerBtn.Visible = true;
        }
        private void ChoixLV_MouseDoubleClick(object sender, System.EventArgs e)
        {
            Continuer();
        }
        private void AnnulerBtn_Click(object sender, EventArgs e)
        {
            ChoixPersonne.Visible = false;
            ChoixPersonne.Enabled = false;
        }
        private void AscendantDeTb_KeyPress(object sender, KeyPressEventArgs e)
        {
            Modifier = true;
            this.Text = NomPrograme + "   *" + FichierCourant;
        }
        private void PatronymeRecherche_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                RechercheID();
            }
        }
        private void PrenomRecherche_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                RechercheID();
            }
        }
        private void CouleurDesBlocsToolStripMenuItem_Click(object sender, EventArgs e)
        {
            ColorDialog MyDialog = new ColorDialog
            {
                AllowFullOpen = false,
                ShowHelp = true
            };
            MyDialog.AllowFullOpen = true;
            MyDialog.Color = RectangleSosa1.FillColor;
            if (MyDialog.ShowDialog() == DialogResult.OK)
                ChangerCouleurBloc(MyDialog.Color);
        }
        private void CreerPage4GénérationsToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (DossierPDF == "")
            {
                AvoirDossierrapport();
            }
            int sosa = 1;
            PdfDocument document = new PdfDocument();
            document.Info.Title = "Tableau Ascendance";
            PdfPage page = document.AddPage();
            page.Size = PageSize.Letter;
            page.Orientation = PdfSharp.PageOrientation.Landscape;
            XGraphics gfx = XGraphics.FromPdfPage(page);
            string numeroTableau = DessinerPage(ref document, ref gfx, sosa, false, false);
            string FichierPage = "4Génération.pdf";
            string Fichier = DossierPDF + "\\" + FichierPage;
            try
            {
                document.Save(Fichier);
                Process.Start(Fichier);
            }
            catch (Exception p)
            {
                SystemSounds.Beep.Play();
                MessageBox.Show("Ne peut pas afficher la page.\r\n\r\n" + p.Message, "Problème ?",
                                 MessageBoxButtons.OK,
                                 MessageBoxIcon.Warning);
            }
        }
        private void Note1_TextChanged(object sender, EventArgs e)
        {
            liste[sosaCourant, NOTE1] = Note1.Text;
        }
        private void Note1_KeyPress(object sender, KeyPressEventArgs e)
        {
            Modifier = true;
            this.Text = NomPrograme + "   *" + FichierCourant;
        }
        private void Note2_TextChanged(object sender, EventArgs e)
        {
            liste[sosaCourant, NOTE2] = Note2.Text;
        }
        private void Note2_KeyPress(object sender, KeyPressEventArgs e)
        {
            Modifier = true;
            this.Text = NomPrograme + "   *" + FichierCourant;
        }
        private void RechercheSosaButton_Click(object sender, EventArgs e)
        {
            DummyButton.Focus();
            RectangleSosa1.BorderColor = Color.Black;
            RectangleSosa2.BorderColor = Color.Black;
            RectangleSosa3.BorderColor = Color.Black;
            RectangleSosa4.BorderColor = Color.Black;
            RectangleSosa5.BorderColor = Color.Black;
            RectangleSosa6.BorderColor = Color.Black;
            RectangleSosa7.BorderColor = Color.Black;
            if (RechercheSosaTextBox.Text == "")
            {
                FlecheGaucheRechercheButton.Visible = false;
                FlecheDroiteRechercheButton.Visible = false;
                return;
            }

            if (Int32.TryParse(RechercheSosaTextBox.Text, out int sosa))
            {
                if (sosa < 0 || sosa > 511) return;
                if (sosa == 0) {
                    ChoixSosaComboBox.Text = "";
                    return;
                }
                ChoixSosaComboBox.Text = liste[sosa, PAGE];
                FlecheGaucheRechercheButton.Visible = false;
                FlecheDroiteRechercheButton.Visible = false;

                if (liste[rechercheListe[0], SOSA] == liste[sosa, PAGE]) RectangleSosa1.BorderColor = Color.White;
                if (Int32.Parse(liste[sosa, SOSA]) == Int32.Parse(liste[sosa, PAGE]) * 2) RectangleSosa2.BorderColor = Color.White;
                if (Int32.Parse(liste[sosa, SOSA]) == Int32.Parse(liste[sosa, PAGE]) * 2 + 1) RectangleSosa3.BorderColor = Color.White;
                if (Int32.Parse(liste[sosa, SOSA]) == Int32.Parse(liste[sosa, PAGE]) * 4) RectangleSosa4.BorderColor = Color.White;
                if (Int32.Parse(liste[sosa, SOSA]) == Int32.Parse(liste[sosa, PAGE]) * 4 + 1) RectangleSosa5.BorderColor = Color.White;
                if (Int32.Parse(liste[sosa, SOSA]) == Int32.Parse(liste[sosa, PAGE]) * 4 + 2) RectangleSosa6.BorderColor = Color.White;
                if (Int32.Parse(liste[sosa, SOSA]) == Int32.Parse(liste[sosa, PAGE]) * 4 + 3) RectangleSosa7.BorderColor = Color.White;

            } else
            {
                string s;
                s = RechercheSosaTextBox.Text;
                if (s == "") return;
                string[] mots = s.Split(' ');
                bool rep;
                int trouver;

                Array.Clear(rechercheListe, 0, rechercheListe.Length);

                rechercheListe[0] = 0;
                trouver = 0;
                for (int f = 1; f < 512; f++)
                {
                    rep = true;
                    foreach (string m in mots)
                    {
                        if (!liste[f, NOMTRI].ToLower().Contains(m.ToLower())) rep = false;
                    }
                    if (rep)
                    {
                        rechercheListe[Int32.Parse(liste[f,SOSA])] = 1;
                        if (rechercheListe[0] == 0) rechercheListe[0] = Int32.Parse(liste[f,SOSA]);
                        trouver = trouver + 1;
                    }
                }
                if (trouver == 0) return;
                ChoixSosaComboBox.Text = liste[rechercheListe[0], PAGE];
                if (liste[rechercheListe[0], SOSA] == liste[rechercheListe[0], PAGE]) RectangleSosa1.BorderColor = Color.White;
                if (Int32.Parse(liste[rechercheListe[0], SOSA]) == Int32.Parse(liste[rechercheListe[0], PAGE]) * 2) RectangleSosa2.BorderColor = Color.White;
                if (Int32.Parse(liste[rechercheListe[0], SOSA]) == Int32.Parse(liste[rechercheListe[0], PAGE]) * 2 + 1) RectangleSosa3.BorderColor = Color.White;
                if (Int32.Parse(liste[rechercheListe[0], SOSA]) == Int32.Parse(liste[rechercheListe[0], PAGE]) * 4) RectangleSosa4.BorderColor = Color.White;
                if (Int32.Parse(liste[rechercheListe[0], SOSA]) == Int32.Parse(liste[rechercheListe[0], PAGE]) * 4 + 1) RectangleSosa5.BorderColor = Color.White;
                if (Int32.Parse(liste[rechercheListe[0], SOSA]) == Int32.Parse(liste[rechercheListe[0], PAGE]) * 4 + 2) RectangleSosa6.BorderColor = Color.White;
                if (Int32.Parse(liste[rechercheListe[0], SOSA]) == Int32.Parse(liste[rechercheListe[0], PAGE]) * 4 + 3) RectangleSosa7.BorderColor = Color.White;
                if (trouver > 1)
                {
                    FlecheGaucheRechercheButton.Visible = true;
                    FlecheDroiteRechercheButton.Visible = true;
                }
                return;
            }
        }
        private void RechercheSosaTextBox_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                e.SuppressKeyPress = true;
                RechercheSosaButton.PerformClick();
            }
        }
        private void FlecheGaucheRechercheButton_Click(object sender, EventArgs e)
        {
            DummyButton.Focus();
            if (rechercheListe[0] == 0) return;
            for (int f = rechercheListe[0] - 1; f > 0 ; f--)
            {
                if (rechercheListe[f] == 1)
                {

                    rechercheListe[0] = f;
                    ChoixSosaComboBox.Text = "";
                    ChoixSosaComboBox.Text = liste[f, PAGE];
                    if (liste[f, SOSA] == liste[f, PAGE]) RectangleSosa1.BorderColor = Color.White;
                    if (Int32.Parse(liste[f, SOSA]) == Int32.Parse(liste[f, PAGE]) * 2) RectangleSosa2.BorderColor = Color.White;
                    if (Int32.Parse(liste[f, SOSA]) == Int32.Parse(liste[f, PAGE]) * 2 + 1) RectangleSosa3.BorderColor = Color.White;
                    if (Int32.Parse(liste[f, SOSA]) == Int32.Parse(liste[f, PAGE]) * 4) RectangleSosa4.BorderColor = Color.White;
                    if (Int32.Parse(liste[f, SOSA]) == Int32.Parse(liste[f, PAGE]) * 4 + 1) RectangleSosa5.BorderColor = Color.White;
                    if (Int32.Parse(liste[f, SOSA]) == Int32.Parse(liste[f, PAGE]) * 4 + 2) RectangleSosa6.BorderColor = Color.White;
                    if (Int32.Parse(liste[f, SOSA]) == Int32.Parse(liste[f, PAGE]) * 4 + 3) RectangleSosa7.BorderColor = Color.White;
                    return;
                }
            }
        }
        private void FlecheDroiteRechercheButton_Click(object sender, EventArgs e)
        {
            DummyButton.Focus();
            if (rechercheListe[0] == 0) return;
            for (int f = rechercheListe[0] + 1; f < 512; f++)
            {
                if (rechercheListe[f] == 1)
                {
                    rechercheListe[0] = f;
                    ChoixSosaComboBox.Text = "";
                    ChoixSosaComboBox.Text = liste[f, PAGE];
                    if (liste[f, SOSA] == liste[f, PAGE]) RectangleSosa1.BorderColor = Color.White;
                    if (Int32.Parse(liste[f, SOSA]) == Int32.Parse(liste[f, PAGE]) * 2) RectangleSosa2.BorderColor = Color.White;
                    if (Int32.Parse(liste[f, SOSA]) == Int32.Parse(liste[f, PAGE]) * 2 + 1) RectangleSosa3.BorderColor = Color.White;
                    if (Int32.Parse(liste[f, SOSA]) == Int32.Parse(liste[f, PAGE]) * 4) RectangleSosa4.BorderColor = Color.White;
                    if (Int32.Parse(liste[f, SOSA]) == Int32.Parse(liste[f, PAGE]) * 4 + 1) RectangleSosa5.BorderColor = Color.White;
                    if (Int32.Parse(liste[f, SOSA]) == Int32.Parse(liste[f, PAGE]) * 4 + 2) RectangleSosa6.BorderColor = Color.White;
                    if (Int32.Parse(liste[f, SOSA]) == Int32.Parse(liste[f, PAGE]) * 4 + 3) RectangleSosa7.BorderColor = Color.White;
                    return;
                }
            }
        }
        private void RechercheSosaTextBox_TextChanged(object sender, EventArgs e)
        {
            FlecheGaucheRechercheButton.Visible = false;
            FlecheDroiteRechercheButton.Visible = false;
        }
        private void CreerUnePagePatrilinéaireToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (DossierPDF == "")
            {
                AvoirDossierrapport();
            }
            XUnit pouce = XUnit.FromInch(1);
            XPen pen1 = new XPen(XColor.FromArgb(0, 0, 0), 1);
            XFont fontNom = new XFont("Arial", 10, XFontStyle.Bold);
            XFont fontDate = new XFont("Arial", 8, XFontStyle.Regular);
            XFont fontB = new XFont("Arial", 8, XFontStyle.Bold);
            XFont font32 = new XFont("arial", 24, XFontStyle.Bold);

            //string dateMariage;
            string nomParent;
            double largeur = 7 * pouce;
            double margin = .75 * pouce;
            double col1 = margin;
            double col2 = margin + .75 * pouce;
            double col3 = col1 + ((largeur / 3));
            double col4 = col3 + ((largeur / 3));
            double col5 = col4 + .75 * pouce;
            double hauteur = .51 * pouce;
            double hauteurLigne = 10;
            double espace = .26 * pouce;
            double padding = 5;

            PdfDocument document = new PdfDocument();
            document.Info.Title = "Lignée patrilinéaire";
            PdfPage page = document.AddPage();
            page.Size = PageSize.Letter;
            page.Orientation = PdfSharp.PageOrientation.Portrait;
            XGraphics gfx = XGraphics.FromPdfPage(page);

            // cadre
            XImage img = global::TableauAscendant.Properties.Resources.cadre;
            gfx.DrawImage(img, 0, 0);

            // cameo droite
            img = global::TableauAscendant.Properties.Resources.male_G_512;
            gfx.DrawImage(img, margin + largeur - 48, 1.25 * pouce, 48, 64);

            // cameo gauche
            img = global::TableauAscendant.Properties.Resources.male_D_512;
            gfx.DrawImage(img, margin, 1.25 * pouce, 48, 64);


            string str = "Titre d'ascendance de";
            XSize textLargeur = gfx.MeasureString(str, font32);
            gfx.DrawString(str, font32, XBrushes.Black, page.Width / 2 - textLargeur.Width / 2, 1.1 * POUCE);
            str = liste[1, PATRONYME];
            textLargeur = gfx.MeasureString(str, font32);
            gfx.DrawString(str, font32, XBrushes.Black, page.Width / 2 - textLargeur.Width / 2, 1.5 * POUCE);
            str = "patrilinéaire";
            textLargeur = gfx.MeasureString(str, font32);
            gfx.DrawString(str, font32, XBrushes.Black, page.Width / 2 - textLargeur.Width / 2, 1.9 * POUCE);

            double Y;

            Y = 2.2 * pouce;
            // entète
            gfx.DrawRectangle(pen1, margin, Y, largeur, hauteur);
            gfx.DrawLine(pen1, col3, Y, col3, Y + hauteur);
            gfx.DrawLine(pen1, col4, Y, col4, Y + hauteur);
            //gfx.DrawLine(pen1, col5, Y, col5, Y + hauteur);
            // col 1 ligne 1 nom
            gfx.DrawString("Non", fontNom, XBrushes.Black, col1 + padding, Y + hauteurLigne);
            // col 1 ligne 2 Date et lieu de naissance
            gfx.DrawString("Date et lieu de naissance", font8, XBrushes.Black, col1 + padding, Y + hauteurLigne * 2);
            // col 1 ligne 3 Date et lieu du décès
            gfx.DrawString("Date et lieu du décès", font8, XBrushes.Black, col1 + padding, Y + hauteurLigne * 3);
            // col 3 ligne 1 Date du mariage
            PDFEcrireCentrer(ref gfx, "Date du mariage", col3, Y + hauteurLigne, col4);
            // col 3 ligne 2 Lieu du mariage
            PDFEcrireCentrer(ref gfx, "Lieu du mariage", col3, Y + hauteurLigne * 2, col4);
            // col 3 ligne 1 Non de la conjointe
            gfx.DrawString("Non de la conjointe", font8, XBrushes.Black, col4 + padding, Y + hauteurLigne);
            // col 4 ligne 2 Nom des parents de la conjointe
            gfx.DrawString("Nom des parents de la conjointe", font8, XBrushes.Black, col4 + padding, Y + hauteurLigne * 2);
            // col 4 ligne 3 Date et lieu du mariage des parents de la conjointe
            //PDFEcrire(ref gfx, "Date et lieu du mariage des parents de la conjointe", col4 + padding, Y + hauteurLigne * 3, 2.3 * pouce);



            int [] sosaListe = new int[] { 256, 128, 64, 32, 16, 8, 4, 2, 1 };
            foreach (int sosa in sosaListe)
            {
                Y = Y + hauteur + espace;
                if (sosa == 256) str = "9e génération";
                if (sosa == 128) str = "8e génération";
                if (sosa == 64) str = "7e génération";
                if (sosa == 32) str = "6e génération";
                if (sosa == 16) str = "5e génération";
                if (sosa == 8) str = "4e génération";
                if (sosa == 4) str = "3e génération";
                if (sosa == 2) str = "2e génération";
                if (sosa == 1) str = "1ère génération";
                textLargeur = gfx.MeasureString(str, fontB);
                gfx.DrawString(str, fontB, XBrushes.Black, page.Width / 2 - textLargeur.Width / 2, Y - 4);
                gfx.DrawRectangle(pen1, margin, Y, largeur, hauteur);
                gfx.DrawLine(pen1, col3, Y, col3, Y + hauteur);
                gfx.DrawLine(pen1, col4, Y, col4, Y + hauteur);
                // col 1 ligne 1 nom
                string nom = AssemblerNom(liste[sosa, PRENOM], liste[sosa, PATRONYME]);
                gfx.DrawString(nom, fontNom, XBrushes.Black, col1 + padding, Y + hauteurLigne);
                // col 1 ligne 2
                if (liste[sosa, NELE] != "" || liste[sosa, NELIEU] != "")
                    gfx.DrawString("°", fontDate, XBrushes.Black, col1 + padding + 1, Y + hauteurLigne * 2);
                PDFEcrire(ref gfx, liste[sosa, NELE], col1 + padding + 6, Y + hauteurLigne * 2, .5 * pouce);
                PDFEcrire(ref gfx, liste[sosa, NELIEU], col2, Y + hauteurLigne * 2, 1 * pouce);
                // col 1 ligne3
                if (liste[sosa, DELE] != "" || liste[sosa, DELIEU] != "")
                    gfx.DrawString("+", fontDate, XBrushes.Black, col1 + padding, Y + hauteurLigne * 3);
                PDFEcrire(ref gfx, liste[sosa, DELE], col1 + padding + 6, Y + hauteurLigne * 3, .5 * pouce);
                PDFEcrire(ref gfx, liste[sosa, DELIEU], col2, Y + hauteurLigne * 3, 1.5 * pouce);

                if (sosa > 1)
                {
                    // col 3 ligne 1
                    if (liste[sosa, MALE] != "")
                    {
                        gfx.DrawString("X ", fontDate, XBrushes.Black, col3 + .83 * pouce, Y + hauteurLigne);
                        gfx.DrawString(liste[sosa, MALE], fontDate, XBrushes.Black, col3 + 6 + .83 * pouce, Y + hauteurLigne);
                    }
                    // col 3 ligne 2
                    PDFEcrireCentrer(ref gfx, liste[sosa, MALIEU], col3, Y + hauteurLigne * 2, col4);

                }
                if (sosa > 1)
                {
                    // col 4 ligne 1 // nom conjoint
                    //PDFEcrire(ref gfx, liste[sosa + 1][PATRONYME], col4 + padding, Y + hauteurLigne, 2.3 * pouce);
                    nom = AssemblerNom(liste[sosa + 1, PRENOM], liste[sosa + 1, PATRONYME]);
                    gfx.DrawString(nom, fontNom, XBrushes.Black, col4 + padding, Y + hauteurLigne);
                    // col 4 ligne 2 ET 3
                    int sosaParent = (sosa + 1) * 2;
                    if (sosaParent < 512 && sosaParent > 0) 
                    {
                        string nomPere = AssemblerNom(liste[sosaParent, PRENOM], liste[sosaParent, PATRONYME]);
                        string nomMere = AssemblerNom(liste[sosaParent + 1, PRENOM], liste[sosaParent + 1, PATRONYME]);
                        if (nomPere != "" && nomMere != "") {
                            nomParent = nomPere + " et " + nomMere;
                            PDFEcrire(ref gfx, nomParent, col4 + padding, Y + hauteurLigne * 2, 2.3 * pouce);
                            if (liste[sosaParent, MALE] != "" || liste[sosaParent, MALIEU] != "")
                            {
                                gfx.DrawString("X", fontDate, XBrushes.Black, col4 + padding + 1, Y + hauteurLigne * 3);
                                PDFEcrire(ref gfx, liste[sosaParent, MALE], col4 + padding + 8, Y + hauteurLigne * 3, .5 * pouce);
                                PDFEcrire(ref gfx, liste[sosaParent, MALIEU], col5, Y + hauteurLigne * 3, 1 * pouce);
                            }
                        }
                    }
                }
            }
            //string numeroTableau = DessinerPage(ref document, ref gfx, sosa, false);
            string FichierPage = "lignee_patrilineairel.pdf";
            string Fichier = DossierPDF + "\\" + FichierPage;
            try
            {
                document.Save(Fichier);
                Process.Start(Fichier);
            }
            catch (Exception p)
            {
                SystemSounds.Beep.Play();
                MessageBox.Show("Ne peut pas afficher la page.\r\n\r\n" + p.Message, "Problème ?",
                                 MessageBoxButtons.OK,
                                 MessageBoxIcon.Warning);
            }
        }
        private void CreerUnePageMatrilénéaireToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (DossierPDF == "")
            {
                AvoirDossierrapport();
            }
            XUnit pouce = XUnit.FromInch(1);
            XPen pen1 = new XPen(XColor.FromArgb(0, 0, 0), 1);
            XFont fontNom = new XFont("Arial", 10, XFontStyle.Regular);
            XFont fontDate = new XFont("Arial", 8, XFontStyle.Regular);
            XFont fontB = new XFont("Arial", 8, XFontStyle.Bold);
            XFont font32 = new XFont("arial", 24, XFontStyle.Bold);

            //string dateMariage;
            string nomParent;
            //int sosaMariage;
            double largeur = 7 * pouce;
            double margin = .75 * pouce;
            double col1 = margin;
            double col2 = margin + .75 * pouce;
            double col3 = col1 + ((largeur / 3));
            double col4 = col3 + ((largeur / 3));
            double col5 = col4 + .75 * pouce;
            double hauteur = .51 * pouce;
            double hauteurLigne = 10;
            double espace = .26 * pouce;
            double padding = 5;

            PdfDocument document = new PdfDocument();
            document.Info.Title = "Lignée matrilinéaire";
            PdfPage page = document.AddPage();
            page.Size = PageSize.Letter;
            page.Orientation = PdfSharp.PageOrientation.Portrait;
            XGraphics gfx = XGraphics.FromPdfPage(page);

            // cadre
            XImage img = global::TableauAscendant.Properties.Resources.cadre;
            gfx.DrawImage(img, 0, 0);

            // cameo droite
            img = global::TableauAscendant.Properties.Resources.femelle_G_512;
            gfx.DrawImage(img, margin + largeur - 48, 1.25 * pouce, 48, 64);

            // cameo gauche
            img = global::TableauAscendant.Properties.Resources.femelle_D_512;
            gfx.DrawImage(img, margin, 1.25 * pouce, 48, 64);


            string str = "Titre d'ascendance de";
            XSize textLargeur = gfx.MeasureString(str, font32);
            gfx.DrawString(str, font32, XBrushes.Black, page.Width / 2 - textLargeur.Width / 2, 1.1 * POUCE);
            str = liste[1, PATRONYME];
            textLargeur = gfx.MeasureString(str, font32);
            gfx.DrawString(str, font32, XBrushes.Black, page.Width / 2 - textLargeur.Width / 2, 1.5 * POUCE);
            str = "matrilinéaire";
            textLargeur = gfx.MeasureString(str, font32);
            gfx.DrawString(str, font32, XBrushes.Black, page.Width / 2 - textLargeur.Width / 2, 1.9 * POUCE);

            double Y;

            Y = 2.2 * pouce;
            // entète
            gfx.DrawRectangle(pen1, margin, Y, largeur, hauteur);
            gfx.DrawLine(pen1, col3, Y, col3, Y + hauteur);
            gfx.DrawLine(pen1, col4, Y, col4, Y + hauteur);

            // col 1 ligne 1 nom
            gfx.DrawString("Non", fontNom, XBrushes.Black, col1 + padding, Y + hauteurLigne);
            // col 1 ligne 2 ° Date et lieu de naissance
            gfx.DrawString("° Date et lieu de naissance", font8, XBrushes.Black, col1 + padding, Y + hauteurLigne * 2);
            // col 1 ligne 3 + Date et lieu du décès
            gfx.DrawString("+ Date et lieu du décès", font8, XBrushes.Black, col1 + padding, Y + hauteurLigne * 3);
            // col 3 ligne 1 X Date du mariage
            PDFEcrireCentrer(ref gfx, "X Date du mariage", col3, Y + hauteurLigne, col4);
            // col 3 ligne 2 Lieu du mariage
            PDFEcrireCentrer(ref gfx, "Lieu du mariage", col3, Y + hauteurLigne * 2, col4);
            // col 3 ligne 1 Non du conjoint
            gfx.DrawString("Non du conjoint", font8, XBrushes.Black, col4 + padding, Y + hauteurLigne);
            // col 4 ligne 2 Nom des parents du conjoint
            gfx.DrawString("Nom des parents du conjoint", font8, XBrushes.Black, col4 + padding, Y + hauteurLigne * 2);
            // col 4 ligne 3 Date et lieu du mariage des parents du conjoint
            PDFEcrire(ref gfx, "Date et lieu du mariage des parents du conjoint", col4 + padding, Y + hauteurLigne * 3, 2.3 * pouce);

            int[] sosaListe = new int[] { 511, 255, 127, 63, 31, 15, 7, 3, 1 };
            foreach (int sosa in sosaListe)
            {
                Y = Y + hauteur + espace;
                if (sosa == 511) str = "9e génération";
                if (sosa == 255) str = "8e génération";
                if (sosa == 127) str = "7e génération";
                if (sosa == 63) str = "6e génération";
                if (sosa == 31) str = "5e génération";
                if (sosa == 15) str = "4e génération";
                if (sosa == 7) str = "3e génération";
                if (sosa == 3) str = "2e génération";
                if (sosa == 1) str = "1ère génération";
                textLargeur = gfx.MeasureString(str, fontB);
                gfx.DrawString(str, fontB, XBrushes.Black, page.Width / 2 - textLargeur.Width / 2, Y - 4);
                gfx.DrawRectangle(pen1, margin, Y, largeur, hauteur);
                gfx.DrawLine(pen1, col3, Y, col3, Y + hauteur);
                gfx.DrawLine(pen1, col4, Y, col4, Y + hauteur);
                // col 1 ligne 1 nom
                string nom = AssemblerNom(liste[sosa, PRENOM], liste[sosa, PATRONYME]);
                gfx.DrawString(nom, fontNom, XBrushes.Black, col1 + padding, Y + hauteurLigne);
                // col 1 ligne 2
                if (liste[sosa, NELE] != "" || liste[sosa, NELIEU] != "")
                    gfx.DrawString("°", fontDate, XBrushes.Black, col1 + padding + 1, Y + hauteurLigne * 2);
                PDFEcrire(ref gfx, liste[sosa, NELE], col1 + padding + 6, Y + hauteurLigne * 2, .5 * pouce);
                PDFEcrire(ref gfx, liste[sosa, NELIEU], col2, Y + hauteurLigne * 2, 1 * pouce);
                // col 1 ligne3
                if (liste[sosa, DELE] != "" || liste[sosa, DELIEU] != "")
                    gfx.DrawString("+", fontDate, XBrushes.Black, col1 + padding, Y + hauteurLigne * 3);
                PDFEcrire(ref gfx, liste[sosa, DELE], col1 + padding + 6, Y + hauteurLigne * 3, .5 * pouce);
                PDFEcrire(ref gfx, liste[sosa, DELIEU], col2, Y + hauteurLigne * 3, 1.5 * pouce);
                
                // col 3 ligne 1
                if (sosa > 1)
                {
                    if (liste[sosa-1, MALE] != "")
                    {
                        gfx.DrawString("X ", fontDate, XBrushes.Black, col3 + .83 * pouce, Y + hauteurLigne);
                        gfx.DrawString(liste[sosa-1, MALE], fontDate, XBrushes.Black, col3 + 6 + .83 * pouce, Y + hauteurLigne);
                    }
                    // col 3 ligne 2
                    PDFEcrireCentrer(ref gfx, liste[sosa-1, MALIEU], col3, Y + hauteurLigne * 2, col4);
                }
                if (sosa > 1)
                {
                // col 4 ligne 1 // nom conjoint
                    //PDFEcrire(ref gfx, liste[sosa - 1][PATRONYME], col4 + padding, Y + hauteurLigne, 2.3 * pouce);
                    nom = AssemblerNom(liste[sosa - 1, PRENOM], liste[sosa - 1, PATRONYME]);
                    gfx.DrawString(nom, fontNom, XBrushes.Black, col4 + padding, Y + hauteurLigne);
                    // col 4 ligne 2 ET 3

                    int sosaParent = (sosa - 1) * 2;
                    if ((sosaParent < 512 && sosaParent > 0))
                    {
                        string nomPere = AssemblerNom(liste[sosaParent, PRENOM], liste[sosaParent, PATRONYME]);
                        string nomMere = AssemblerNom(liste[sosaParent + 1, PRENOM], liste[sosaParent + 1, PATRONYME]);
                        if (nomPere != "" && nomMere != "")
                        {
                            nomParent = nomPere + " et " + nomMere;
                            PDFEcrire(ref gfx, nomParent, col4 + padding, Y + hauteurLigne * 2, 2.3 * pouce);
                            if (liste[sosaParent, MALE] != "" || liste[sosaParent, MALIEU] != "")
                            {
                                gfx.DrawString("X", fontDate, XBrushes.Black, col4 + padding + 1, Y + hauteurLigne * 3);
                                PDFEcrire(ref gfx, liste[sosaParent, MALE], col4 + padding + 8, Y + hauteurLigne * 3, .5 * pouce);
                                PDFEcrire(ref gfx, liste[sosaParent, MALIEU], col5, Y + hauteurLigne * 3, 1 * pouce);
                            }
                        }
                    }
                }
            }
            //string numeroTableau = DessinerPage(ref document, ref gfx, sosa, false);
            string FichierPage = "lignee_matrilineairel.pdf";
            string Fichier = DossierPDF + "\\" + FichierPage;
            try
            {
                document.Save(Fichier);
                Process.Start(Fichier);
            }
            catch (Exception p)
            {
                SystemSounds.Beep.Play();
                MessageBox.Show("Ne peut pas afficher la page.\r\n\r\n" + p.Message, "Problème ?",
                                 MessageBoxButtons.OK,
                                 MessageBoxIcon.Warning);
            }
        }
        private void ContinuerBtn_Click(object sender, EventArgs e)
        {
            Continuer();
        }
        private void GoSosaConjoint1Btn_Click(object sender, EventArgs e)
        {
            double a = System.Convert.ToInt32(sosaCourant + 1);
            //a = Math.Floor(a / 8);
            //int b = System.Convert.ToInt32(a);
            ChoixSosaComboBox.Text = Convert.ToString(a);
        }

        
    }

    internal class List<T1, T2>
    {
    }
}

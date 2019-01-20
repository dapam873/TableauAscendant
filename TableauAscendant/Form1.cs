// compile with: -doc:Form1.xml 
using System; 

using System.Collections.Generic;
using System.Data;
using System.Windows.Forms;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Text;
using System.Runtime.CompilerServices;
using System.Media;
using System.Reflection;
using PdfSharp;
using PdfSharp.Drawing;
using PdfSharp.Drawing.Layout;
using PdfSharp.Pdf;
//using PdfSharp.Forms;
//using PdfSharp.Charting;
//using PdfSharp.Forms;
//using PdfSharp.Pdf.IO;
using TableauAscendant;

namespace WindowsFormsApp1
{
    ///<Summary>
    /// 
    ///</Summary>
    public partial class TableauAscendant : Form
    {
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
        /// position de la colonne NOM dans le tableau grille
        /// </summary>
        public const int NOM =               4;
        /// <summary>
        /// position de la colonne NELE dans le tableau grille
        /// </summary>
        public const int NELE =              5;
        /// <summary>
        /// position de la colonne NELIEU dans le tableau grille
        /// </summary>
        public const int NELIEU =            6;
        /// <summary>
        /// position de la colonne DELE dans le tableau grille
        /// </summary>
        public const int DELE =              7;
        /// <summary>
        /// position de la colonne DELIEU dans le tableau grille
        /// </summary>
        public const int DELIEU =            8;
        /// <summary>
        /// position de la colonne MALE dans le tableau grille
        /// </summary>
        public const int MALE =              9;
        /// <summary>
        /// position de la colonne MALIEU dans le tableau grille
        /// </summary>
        public const int MALIEU =           10;
        /// <summary>
        /// position de la colonne IDg dans le tableau grille
        /// </summary>
        public const int IDg =              11;
        /// <summary>
        /// position de la colonne IDFAMILLEENFANT dans le tableau grille
        /// </summary>
        public const int IDFAMILLEENFANT =  12;
        /// <summary>
        /// position de la colonne NOTE1 dans le tableau grille
        /// </summary>
        public const int NOTE1 =            13;
        /// <summary>
        /// position de la colonne Note2 dans le tableau grille
        /// </summary>
        public const int NOTE2 =            14;
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
        /// numero du SOSA courantlargeur maximun de nom dans fiche
        /// </summary>
        public int sosaCourant = 0;
        /// <summary>
        /// grille qui contient toutes les informations pour généré les tableaux
        /// </summary>
        public string[][] grille = new string[512][]; // 512 lignes
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

        //{
        //    _nomFichier = "";
        //};
        //  Function **************************************************************************************************************************

        private void    AfficherData()
        {
            if (ChoixSosaComboBox.Text == "")
            {
                // enlève les cases

                Sosa1NomTextBox.Visible = false;
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
                Sosa2NomTextBox.Visible = false;
                Sosa2NeTextBox.Visible = false;
                Sosa2NeEndroitTextBox.Visible = false;
                Sosa2DeTextBox.Visible = false;
                Sosa2DeEndroitTextBox.Visible = false;
                Sosa23MaTextBox.Visible = false;
                Sosa23MaEndroitTextBox.Visible = false;
                RectangleSosa2.BorderColor = Color.Black;

                Sosa3Label.Visible = false;
                Sosa3NomTextBox.Visible = false;
                Sosa3NeTextBox.Visible = false;
                Sosa3NeEndroitTextBox.Visible = false;
                Sosa3DeTextBox.Visible = false;
                Sosa3DeEndroitTextBox.Visible = false;
                RectangleSosa3.BorderColor = Color.Black;

                Sosa4Label.Visible = false;
                Sosa4NomTextBox.Visible = false;
                Sosa4NeTextBox.Visible = false;
                Sosa4NeEndroitTextBox.Visible = false;
                Sosa4DeTextBox.Visible = false;
                Sosa4DeEndroitTextBox.Visible = false;
                Sosa45MaTextBox.Visible = false;
                Sosa45MaLEndroitTextBox.Visible = false;
                RectangleSosa4.BorderColor = Color.Black;

                Sosa5Label.Visible = false;
                Sosa5NomTextBox.Visible = false;
                Sosa5NeTextBox.Visible = false;
                Sosa5NeEndroitTextBox.Visible = false;
                Sosa5DeTextBox.Visible = false;
                Sosa5DeEndroitTextBox.Visible = false;
                RectangleSosa5.BorderColor = Color.Black;

                Sosa6Label.Visible = false;
                Sosa6NomTextBox.Visible = false;
                Sosa6NeTextBox.Visible = false;
                Sosa6NeEndroitTextBox.Visible = false;
                Sosa6DeTextBox.Visible = false;
                Sosa6DeEndroitTextBox.Visible = false;
                Sosa67MaTextBox.Visible = false;
                Sosa67MaEndroitTextBox.Visible = false;
                RectangleSosa6.BorderColor = Color.Black;

                Sosa7Label.Visible = false;
                Sosa7NomTextBox.Visible = false;
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
                Sosa1NomTextBox.Visible = true;
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
                Sosa2NomTextBox.Visible = true;
                Sosa2NeTextBox.Visible = true;
                Sosa2NeEndroitTextBox.Visible = true;
                Sosa2DeTextBox.Visible = true;
                Sosa2DeEndroitTextBox.Visible = true;
                Sosa23MaTextBox.Visible = true;
                Sosa23MaEndroitTextBox.Visible = true;
                RectangleSosa2.BorderColor = Color.Black;

                Sosa3Label.Visible = true;
                Sosa3NomTextBox.Visible = true;
                Sosa3NeTextBox.Visible = true;
                Sosa3NeEndroitTextBox.Visible = true;
                Sosa3DeTextBox.Visible = true;
                Sosa3DeEndroitTextBox.Visible = true;
                RectangleSosa3.BorderColor = Color.Black;

                Sosa4Label.Visible = true;
                Sosa4NomTextBox.Visible = true;
                Sosa4NeTextBox.Visible = true;
                Sosa4NeEndroitTextBox.Visible = true;
                Sosa4DeTextBox.Visible = true;
                Sosa4DeEndroitTextBox.Visible = true;
                Sosa45MaTextBox.Visible = true;
                Sosa45MaLEndroitTextBox.Visible = true;
                RectangleSosa4.BorderColor = Color.Black;

                Sosa5Label.Visible = true;
                Sosa5NomTextBox.Visible = true;
                Sosa5NeTextBox.Visible = true;
                Sosa5NeEndroitTextBox.Visible = true;
                Sosa5DeTextBox.Visible = true;
                Sosa5DeEndroitTextBox.Visible = true;
                RectangleSosa5.BorderColor = Color.Black;

                Sosa6Label.Visible = true;
                Sosa6NomTextBox.Visible = true;
                Sosa6NeTextBox.Visible = true;
                Sosa6NeEndroitTextBox.Visible = true;
                Sosa6DeTextBox.Visible = true;
                Sosa6DeEndroitTextBox.Visible = true;
                Sosa67MaTextBox.Visible = true;
                Sosa67MaEndroitTextBox.Visible = true;
                RectangleSosa6.BorderColor = Color.Black;

                Sosa7Label.Visible = true;
                Sosa7NomTextBox.Visible = true;
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
            SosaConjoint1NomTextBox.BackColor = rgb;
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
            grille[1][IDg] = ID;

            string nom = GEDCOM.AvoirPrenom(ID) + " " + GEDCOM.AvoirNom(ID);
            grille[1][NOM] = nom;

            string dateN = ConvertirDate(GEDCOM.AvoirDateNaissance(ID));
            grille[1][NELE] = dateN;
            grille[1][NELIEU] = GEDCOM.AvoirEndroitNaissance(ID);
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
            grille[1][MALE] = ConvertirDate(GEDCOM.AvoirDateMariage(IDListeFamilleEpoux[0]));
            grille[1][MALIEU] = GEDCOM.AvoirEndroitMariage(IDListeFamilleEpoux[0]);
            grille[1][IDg] = ID.ToString();
            string IDFamilleEnfant = GEDCOM.AvoirFamilleEnfant(ID);
            grille[1][IDFAMILLEENFANT] = GEDCOM.AvoirFamilleEnfant(ID);
            grille[0][NOM] = GEDCOM.AvoirPrenom(IDConjoint) + " " + GEDCOM.AvoirNom(IDConjoint);
            for (int f = 2; f < 510; f += 2)
            {
                int a = f / 2;
                string ss = grille[f / 2][IDFAMILLEENFANT];

                //string IDFamilleEnfant = GEDCOM.AvoirFamilleEnfant(ID);
                if (ss != "")
                {
                    grille[f][IDg] = GEDCOM.AvoirEpoux(ss);
                    grille[f][IDFAMILLEENFANT] = GEDCOM.AvoirFamilleEnfant(grille[f][IDg]);
                    grille[f + 1][IDg] = GEDCOM.AvoirEpouse(ss);
                    grille[f + 1][IDFAMILLEENFANT] = GEDCOM.AvoirFamilleEnfant(grille[f + 1][IDg]);
                }
            }
            for (int f = 2; f < 510; f += 2)
            {
                //grille[f, IDg] = sosaID[f].ToString();
                ID = grille[f][IDg];
                string n = GEDCOM.AvoirNom(ID);
                string p = GEDCOM.AvoirPrenom(ID);
                string np = "";
                if (n != "" && p != "") np = GEDCOM.AvoirPrenom(ID) + " " + GEDCOM.AvoirNom(ID);
                if (n != "" && p == "") np = GEDCOM.AvoirNom(ID);
                if (n == "" && p != "") np = GEDCOM.AvoirPrenom(ID);
                if (n == "" && p == "") np = "";
                grille[f][NOM] = np;
                grille[f][NELE] = ConvertirDate(GEDCOM.AvoirDateNaissance(ID));
                grille[f][NELIEU] = GEDCOM.AvoirEndroitNaissance(ID);
                grille[f][DELE] = ConvertirDate(GEDCOM.AvoirDateDeces(ID));
                grille[f][DELIEU] = GEDCOM.AvoirEndroitDeces(ID);
                grille[f][IDFAMILLEENFANT] = GEDCOM.AvoirFamilleEnfant(grille[f][IDg]);
                string ss = grille[f / 2][IDFAMILLEENFANT];
                grille[f][MALE] = ConvertirDate(GEDCOM.AvoirDateMariage(ss));
                grille[f][MALIEU] = GEDCOM.AvoirEndroitMariage(ss);
                int ff = f + 1;
                ID = grille[ff][IDg];
                n = GEDCOM.AvoirNom(ID);
                p = GEDCOM.AvoirPrenom(ID);
                np = "";
                if (n != "" && p != "") np = GEDCOM.AvoirPrenom(ID) + " " + GEDCOM.AvoirNom(ID);
                if (n != "" && p == "") np = GEDCOM.AvoirNom(ID);
                if (n == "" && p != "") np = GEDCOM.AvoirPrenom(ID);
                if (n == "" && p == "") np = "";
                grille[ff][NOM] = np;
                grille[ff][NELE] = ConvertirDate(GEDCOM.AvoirDateNaissance(ID));
                grille[ff][NELIEU] = GEDCOM.AvoirEndroitNaissance(ID);
                grille[ff][DELE] = ConvertirDate(GEDCOM.AvoirDateDeces(ID));
                grille[ff][DELIEU] = GEDCOM.AvoirEndroitDeces(ID);
            }
            NomRecherche.Text = "";
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
                        if ((grille[index][NOM] != "" || grille[index][NELE] != "" || grille[index][NELIEU] != "" || grille[index][DELE] != ""
                             || grille[index][DELIEU] != "" || grille[index][MALE] != "" || grille[index][MALIEU] != "" ||
                             grille[index][NOTE1] != "" || grille[index][NOTE2] != "") && grille[index][SOSA] != "0")
                        {
                            ligne.WriteLine("[sosa*]");
                            ligne.WriteLine("No    =" + grille[index][SOSA]);
                            if (grille[index][NOM] != "") ligne.WriteLine("Nom   =" + grille[index][NOM]);
                            if (grille[index][NELE] != "") ligne.WriteLine("NeLe  =" + grille[index][NELE]);
                            if (grille[index][NELIEU] != "") ligne.WriteLine("NeLieu=" + grille[index][NELIEU]);
                            if (grille[index][DELE] != "") ligne.WriteLine("DeLe  =" + grille[index][DELE]);
                            if (grille[index][DELIEU] != "") ligne.WriteLine("DeLieu=" + grille[index][DELIEU]);
                            if (grille[index][MALE] != "") ligne.WriteLine("MaLe  =" + grille[index][MALE]);
                            if (grille[index][MALIEU] != "") ligne.WriteLine("MaLieu=" + grille[index][MALIEU]);
                            if (grille[index][NOTE1].Length > 0)
                            {
                                ligne.WriteLine("NoteH =");
                                ligne.WriteLine(grille[index][NOTE1]);
                                ligne.WriteLine("##FIN##");
                            }
                            if (grille[index][NOTE2].Length > 0)
                            {
                                ligne.WriteLine("NoteB =");
                                ligne.WriteLine(grille[index][NOTE2]);
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
                grille[f] = new string[15];
                grille[f][SOSA] = f.ToString();
                grille[f][PAGE] = "";
                grille[f][GENERATION] = "";
                grille[f][TABLEAU] = "";
                grille[f][NOM] = "";
                grille[f][NELE] = "";
                grille[f][NELIEU] = "";
                grille[f][DELE] = "";
                grille[f][DELIEU] = "";
                grille[f][MALE] = "";
                grille[f][MALIEU] = "";
                grille[f][IDg] = "";
                grille[f][IDFAMILLEENFANT] = "";
                grille[f][NOTE1] = "";
                grille[f][NOTE2] = "";
            }
            grille[1][GENERATION] = "1";
            grille[2][GENERATION] = "2";
            grille[3][GENERATION] = "3";
            for (int f = 4; f < 8; f++)
            {
                grille[f][GENERATION] = "3";
            }
            for (int f = 8; f < 16; f++)
            {
                grille[f][GENERATION] = "4";
            }
            for (int f = 16; f < 32; f++)
            {
                grille[f][GENERATION] = "5";
            }
            for (int f = 32; f < 64; f++)
            {
                grille[f][GENERATION] = "6";
            }
            for (int f = 64; f < 128; f++)
            {
                grille[f][GENERATION] = "7";
            }
            for (int f = 128; f < 256; f++)
            {
                grille[f][GENERATION] = "8";
            }
            for (int f = 256; f < 512; f++)
            {
                grille[f][GENERATION] = "9";
            }
            for (int f = 1; f < 8; f++)
            {
                grille[f][TABLEAU] = "1";

            }

            grille[ 8][TABLEAU] = "2";
            grille[ 9][TABLEAU] = "3";
            grille[10][TABLEAU] = "4";
            grille[11][TABLEAU] = "5";
            grille[12][TABLEAU] = "6";
            grille[13][TABLEAU] = "7";
            grille[14][TABLEAU] = "8";
            grille[15][TABLEAU] = "9";
            grille[16][TABLEAU] = "2";
            grille[17][TABLEAU] = "2";
            grille[18][TABLEAU] = "3";
            grille[19][TABLEAU] = "3";
            grille[20][TABLEAU] = "4";
            grille[21][TABLEAU] = "4";
            grille[22][TABLEAU] = "5";
            grille[23][TABLEAU] = "5";
            grille[24][TABLEAU] = "6";
            grille[25][TABLEAU] = "6";
            grille[26][TABLEAU] = "7";
            grille[27][TABLEAU] = "7";
            grille[28][TABLEAU] = "8";
            grille[29][TABLEAU] = "8";
            grille[30][TABLEAU] = "9";
            grille[31][TABLEAU] = "9";

            for (int f = 32; f < 36; f++)
            {
                grille[f][TABLEAU] = "2";
            }
            for (int f = 36; f < 40; f++)
            {
                grille[f][TABLEAU] = "3";
            }
            for (int f = 40; f < 44; f++)
            {
                grille[f][TABLEAU] = "4";
            }
            for (int f = 44; f < 48; f++)
            {
                grille[f][TABLEAU] = "5";
            }
            for (int f = 48; f < 52; f++)
            {
                grille[f][TABLEAU] = "6";
            }
            for (int f = 52; f < 56; f++)
            {
                grille[f][TABLEAU] = "7";
            }
            for (int f = 56; f < 60; f++)
            {
                grille[f][TABLEAU] = "8";
            }
            for (int f = 60; f < 64; f++)
            {
                grille[f][TABLEAU] = "9";
            }
            grille[64][TABLEAU] = "10";
            grille[65][TABLEAU] = "11";
            grille[66][TABLEAU] = "12";
            grille[67][TABLEAU] = "13";
            grille[68][TABLEAU] = "14";
            grille[69][TABLEAU] = "15";
            grille[70][TABLEAU] = "16";
            grille[71][TABLEAU] = "17";
            grille[72][TABLEAU] = "18";
            grille[73][TABLEAU] = "19";
            grille[74][TABLEAU] = "20";
            grille[75][TABLEAU] = "21";
            grille[76][TABLEAU] = "22";
            grille[77][TABLEAU] = "23";
            grille[78][TABLEAU] = "24";
            grille[79][TABLEAU] = "25";
            grille[80][TABLEAU] = "26";
            grille[81][TABLEAU] = "27";
            grille[82][TABLEAU] = "28";
            grille[83][TABLEAU] = "29";
            grille[84][TABLEAU] = "30";
            grille[85][TABLEAU] = "31";
            grille[86][TABLEAU] = "32";
            grille[87][TABLEAU] = "33";
            grille[88][TABLEAU] = "34";
            grille[89][TABLEAU] = "35";
            grille[90][TABLEAU] = "36";
            grille[91][TABLEAU] = "37";
            grille[92][TABLEAU] = "38";
            grille[93][TABLEAU] = "39";
            grille[94][TABLEAU] = "40";
            grille[95][TABLEAU] = "41";
            grille[96][TABLEAU] = "42";
            grille[97][TABLEAU] = "43";
            grille[98][TABLEAU] = "44";
            grille[99][TABLEAU] = "45";
            grille[100][TABLEAU] = "46";
            grille[101][TABLEAU] = "47";
            grille[102][TABLEAU] = "48";
            grille[103][TABLEAU] = "49";
            grille[104][TABLEAU] = "50";
            grille[105][TABLEAU] = "51";
            grille[106][TABLEAU] = "52";
            grille[107][TABLEAU] = "53";
            grille[108][TABLEAU] = "54";
            grille[109][TABLEAU] = "55";
            grille[110][TABLEAU] = "56";
            grille[111][TABLEAU] = "57";
            grille[112][TABLEAU] = "58";
            grille[113][TABLEAU] = "59";
            grille[114][TABLEAU] = "60";
            grille[115][TABLEAU] = "61";
            grille[116][TABLEAU] = "62";
            grille[117][TABLEAU] = "63";
            grille[118][TABLEAU] = "64";
            grille[119][TABLEAU] = "65";
            grille[120][TABLEAU] = "66";
            grille[121][TABLEAU] = "67";
            grille[122][TABLEAU] = "68";
            grille[123][TABLEAU] = "69";
            grille[124][TABLEAU] = "70";
            grille[125][TABLEAU] = "71";
            grille[126][TABLEAU] = "72";
            grille[127][TABLEAU] = "73";
            grille[128][TABLEAU] = "10";
            grille[129][TABLEAU] = "10";
            grille[130][TABLEAU] = "11";
            grille[131][TABLEAU] = "11";
            grille[132][TABLEAU] = "12";
            grille[133][TABLEAU] = "12";
            grille[134][TABLEAU] = "13";
            grille[135][TABLEAU] = "13";
            grille[136][TABLEAU] = "14";
            grille[137][TABLEAU] = "14";
            grille[138][TABLEAU] = "15";
            grille[139][TABLEAU] = "15";
            grille[140][TABLEAU] = "16";
            grille[141][TABLEAU] = "16";
            grille[142][TABLEAU] = "17";
            grille[143][TABLEAU] = "17";
            grille[144][TABLEAU] = "18";
            grille[145][TABLEAU] = "18";
            grille[146][TABLEAU] = "19";
            grille[147][TABLEAU] = "19";
            grille[148][TABLEAU] = "20";
            grille[149][TABLEAU] = "20";
            grille[150][TABLEAU] = "21";
            grille[151][TABLEAU] = "21";
            grille[152][TABLEAU] = "22";
            grille[153][TABLEAU] = "22";
            grille[154][TABLEAU] = "23";
            grille[155][TABLEAU] = "23";
            grille[156][TABLEAU] = "24";
            grille[157][TABLEAU] = "24";
            grille[158][TABLEAU] = "25";
            grille[159][TABLEAU] = "25";
            grille[160][TABLEAU] = "26";
            grille[161][TABLEAU] = "26";
            grille[162][TABLEAU] = "27";
            grille[163][TABLEAU] = "27";
            grille[164][TABLEAU] = "28";
            grille[165][TABLEAU] = "28";
            grille[166][TABLEAU] = "29";
            grille[167][TABLEAU] = "29";
            grille[168][TABLEAU] = "30";
            grille[169][TABLEAU] = "30";
            grille[170][TABLEAU] = "31";
            grille[171][TABLEAU] = "31";
            grille[172][TABLEAU] = "32";
            grille[173][TABLEAU] = "32";
            grille[174][TABLEAU] = "33";
            grille[175][TABLEAU] = "33";
            grille[176][TABLEAU] = "34";
            grille[177][TABLEAU] = "34";
            grille[178][TABLEAU] = "35";
            grille[179][TABLEAU] = "35";
            grille[180][TABLEAU] = "36";
            grille[181][TABLEAU] = "36";
            grille[182][TABLEAU] = "37";
            grille[183][TABLEAU] = "37";
            grille[184][TABLEAU] = "38";
            grille[185][TABLEAU] = "38";
            grille[186][TABLEAU] = "39";
            grille[187][TABLEAU] = "39";
            grille[188][TABLEAU] = "40";
            grille[189][TABLEAU] = "40";
            grille[190][TABLEAU] = "41";
            grille[191][TABLEAU] = "41";
            grille[192][TABLEAU] = "42";
            grille[193][TABLEAU] = "42";
            grille[194][TABLEAU] = "43";
            grille[195][TABLEAU] = "43";
            grille[196][TABLEAU] = "44";
            grille[197][TABLEAU] = "44";
            grille[198][TABLEAU] = "45";
            grille[199][TABLEAU] = "45";
            grille[200][TABLEAU] = "46";
            grille[201][TABLEAU] = "46";
            grille[202][TABLEAU] = "47";
            grille[203][TABLEAU] = "47";
            grille[204][TABLEAU] = "48";
            grille[205][TABLEAU] = "48";
            grille[206][TABLEAU] = "49";
            grille[207][TABLEAU] = "49";
            grille[208][TABLEAU] = "50";
            grille[209][TABLEAU] = "50";
            grille[210][TABLEAU] = "51";
            grille[211][TABLEAU] = "51";
            grille[212][TABLEAU] = "52";
            grille[213][TABLEAU] = "52";
            grille[214][TABLEAU] = "53";
            grille[215][TABLEAU] = "53";
            grille[216][TABLEAU] = "54";
            grille[217][TABLEAU] = "54";
            grille[218][TABLEAU] = "55";
            grille[219][TABLEAU] = "55";
            grille[220][TABLEAU] = "56";
            grille[221][TABLEAU] = "56";
            grille[222][TABLEAU] = "57";
            grille[223][TABLEAU] = "57";
            grille[224][TABLEAU] = "58";
            grille[225][TABLEAU] = "58";
            grille[226][TABLEAU] = "59";
            grille[227][TABLEAU] = "59";
            grille[228][TABLEAU] = "60";
            grille[229][TABLEAU] = "60";
            grille[230][TABLEAU] = "61";
            grille[231][TABLEAU] = "61";
            grille[232][TABLEAU] = "62";
            grille[233][TABLEAU] = "62";
            grille[234][TABLEAU] = "63";
            grille[235][TABLEAU] = "63";
            grille[236][TABLEAU] = "64";
            grille[237][TABLEAU] = "64";
            grille[238][TABLEAU] = "65";
            grille[239][TABLEAU] = "65";
            grille[240][TABLEAU] = "66";
            grille[241][TABLEAU] = "66";
            grille[242][TABLEAU] = "67";
            grille[243][TABLEAU] = "67";
            grille[244][TABLEAU] = "68";
            grille[245][TABLEAU] = "68";
            grille[246][TABLEAU] = "69";
            grille[247][TABLEAU] = "69";
            grille[248][TABLEAU] = "70";
            grille[249][TABLEAU] = "70";
            grille[250][TABLEAU] = "71";
            grille[251][TABLEAU] = "71";
            grille[252][TABLEAU] = "72";
            grille[253][TABLEAU] = "72";
            grille[254][TABLEAU] = "73";
            grille[255][TABLEAU] = "73";


            grille[256][TABLEAU] = "10";
            grille[257][TABLEAU] = "10";
            grille[258][TABLEAU] = "10";
            grille[259][TABLEAU] = "10";
            grille[260][TABLEAU] = "11";
            grille[261][TABLEAU] = "11";
            grille[262][TABLEAU] = "11";
            grille[263][TABLEAU] = "11";
            grille[264][TABLEAU] = "12";
            grille[265][TABLEAU] = "12";
            grille[266][TABLEAU] = "12";
            grille[267][TABLEAU] = "12";
            grille[268][TABLEAU] = "13";
            grille[269][TABLEAU] = "13";
            grille[270][TABLEAU] = "13";
            grille[271][TABLEAU] = "13";
            grille[272][TABLEAU] = "14";
            grille[273][TABLEAU] = "14";
            grille[274][TABLEAU] = "14";
            grille[275][TABLEAU] = "14";
            grille[276][TABLEAU] = "15";
            grille[277][TABLEAU] = "15";
            grille[278][TABLEAU] = "15";
            grille[279][TABLEAU] = "15";
            grille[280][TABLEAU] = "16";
            grille[281][TABLEAU] = "16";
            grille[282][TABLEAU] = "16";
            grille[283][TABLEAU] = "16";
            grille[284][TABLEAU] = "17";
            grille[285][TABLEAU] = "17";
            grille[286][TABLEAU] = "17";
            grille[287][TABLEAU] = "17";
            grille[288][TABLEAU] = "18";
            grille[289][TABLEAU] = "18";
            grille[290][TABLEAU] = "18";
            grille[291][TABLEAU] = "18";
            grille[292][TABLEAU] = "19";
            grille[293][TABLEAU] = "19";
            grille[294][TABLEAU] = "19";
            grille[295][TABLEAU] = "19";
            grille[296][TABLEAU] = "20";
            grille[297][TABLEAU] = "20";
            grille[298][TABLEAU] = "20";
            grille[299][TABLEAU] = "20";
            grille[300][TABLEAU] = "21";
            grille[301][TABLEAU] = "21";
            grille[302][TABLEAU] = "21";
            grille[303][TABLEAU] = "21";
            grille[304][TABLEAU] = "22";
            grille[305][TABLEAU] = "22";
            grille[306][TABLEAU] = "22";
            grille[307][TABLEAU] = "22";
            grille[308][TABLEAU] = "23";
            grille[309][TABLEAU] = "23";
            grille[310][TABLEAU] = "23";
            grille[311][TABLEAU] = "23";
            grille[312][TABLEAU] = "24";
            grille[313][TABLEAU] = "24";
            grille[314][TABLEAU] = "24";
            grille[315][TABLEAU] = "24";
            grille[316][TABLEAU] = "25";
            grille[317][TABLEAU] = "25";
            grille[318][TABLEAU] = "25";
            grille[319][TABLEAU] = "25";
            grille[320][TABLEAU] = "26";
            grille[321][TABLEAU] = "26";
            grille[322][TABLEAU] = "26";
            grille[323][TABLEAU] = "26";
            grille[324][TABLEAU] = "27";
            grille[325][TABLEAU] = "27";
            grille[326][TABLEAU] = "27";
            grille[327][TABLEAU] = "27";
            grille[328][TABLEAU] = "28";
            grille[329][TABLEAU] = "28";
            grille[330][TABLEAU] = "28";
            grille[331][TABLEAU] = "28";
            grille[332][TABLEAU] = "29";
            grille[333][TABLEAU] = "29";
            grille[334][TABLEAU] = "29";
            grille[335][TABLEAU] = "29";
            grille[336][TABLEAU] = "30";
            grille[337][TABLEAU] = "30";
            grille[338][TABLEAU] = "30";
            grille[339][TABLEAU] = "30";
            grille[340][TABLEAU] = "31";
            grille[341][TABLEAU] = "31";
            grille[342][TABLEAU] = "31";
            grille[343][TABLEAU] = "31";
            grille[344][TABLEAU] = "32";
            grille[345][TABLEAU] = "32";
            grille[346][TABLEAU] = "32";
            grille[347][TABLEAU] = "32";
            grille[348][TABLEAU] = "33";
            grille[349][TABLEAU] = "33";
            grille[350][TABLEAU] = "33";
            grille[351][TABLEAU] = "33";
            grille[352][TABLEAU] = "34";
            grille[353][TABLEAU] = "34";
            grille[354][TABLEAU] = "34";
            grille[355][TABLEAU] = "34";
            grille[356][TABLEAU] = "35";
            grille[357][TABLEAU] = "35";
            grille[358][TABLEAU] = "35";
            grille[359][TABLEAU] = "35";
            grille[360][TABLEAU] = "36";
            grille[361][TABLEAU] = "36";
            grille[362][TABLEAU] = "36";
            grille[363][TABLEAU] = "36";
            grille[364][TABLEAU] = "37";
            grille[365][TABLEAU] = "37";
            grille[366][TABLEAU] = "37";
            grille[367][TABLEAU] = "37";
            grille[368][TABLEAU] = "38";
            grille[369][TABLEAU] = "38";
            grille[370][TABLEAU] = "38";
            grille[371][TABLEAU] = "38";
            grille[372][TABLEAU] = "39";
            grille[373][TABLEAU] = "39";
            grille[374][TABLEAU] = "39";
            grille[375][TABLEAU] = "39";
            grille[376][TABLEAU] = "40";
            grille[377][TABLEAU] = "40";
            grille[378][TABLEAU] = "40";
            grille[379][TABLEAU] = "40";
            grille[380][TABLEAU] = "41";
            grille[381][TABLEAU] = "41";
            grille[382][TABLEAU] = "41";
            grille[383][TABLEAU] = "41";
            grille[384][TABLEAU] = "42";
            grille[385][TABLEAU] = "42";
            grille[386][TABLEAU] = "42";
            grille[387][TABLEAU] = "42";
            grille[388][TABLEAU] = "43";
            grille[389][TABLEAU] = "43";
            grille[390][TABLEAU] = "43";
            grille[391][TABLEAU] = "43";
            grille[392][TABLEAU] = "44";
            grille[393][TABLEAU] = "44";
            grille[394][TABLEAU] = "44";
            grille[395][TABLEAU] = "44";
            grille[396][TABLEAU] = "45";
            grille[397][TABLEAU] = "45";
            grille[398][TABLEAU] = "45";
            grille[399][TABLEAU] = "45";
            grille[400][TABLEAU] = "46";
            grille[401][TABLEAU] = "46";
            grille[402][TABLEAU] = "46";
            grille[403][TABLEAU] = "46";
            grille[404][TABLEAU] = "47";
            grille[405][TABLEAU] = "47";
            grille[406][TABLEAU] = "47";
            grille[407][TABLEAU] = "47";
            grille[408][TABLEAU] = "48";
            grille[409][TABLEAU] = "48";
            grille[410][TABLEAU] = "48";
            grille[411][TABLEAU] = "48";
            grille[412][TABLEAU] = "49";
            grille[413][TABLEAU] = "49";
            grille[414][TABLEAU] = "49";
            grille[415][TABLEAU] = "49";
            grille[416][TABLEAU] = "50";
            grille[417][TABLEAU] = "50";
            grille[418][TABLEAU] = "50";
            grille[419][TABLEAU] = "50";
            grille[420][TABLEAU] = "51";
            grille[421][TABLEAU] = "51";
            grille[422][TABLEAU] = "51";
            grille[423][TABLEAU] = "51";
            grille[424][TABLEAU] = "52";
            grille[425][TABLEAU] = "52";
            grille[426][TABLEAU] = "52";
            grille[427][TABLEAU] = "52";
            grille[428][TABLEAU] = "53";
            grille[429][TABLEAU] = "53";
            grille[430][TABLEAU] = "53";
            grille[431][TABLEAU] = "53";
            grille[432][TABLEAU] = "54";
            grille[433][TABLEAU] = "54";
            grille[434][TABLEAU] = "54";
            grille[435][TABLEAU] = "54";
            grille[436][TABLEAU] = "55";
            grille[437][TABLEAU] = "55";
            grille[438][TABLEAU] = "55";
            grille[439][TABLEAU] = "55";
            grille[440][TABLEAU] = "56";
            grille[441][TABLEAU] = "56";
            grille[442][TABLEAU] = "56";
            grille[443][TABLEAU] = "56";
            grille[444][TABLEAU] = "57";
            grille[445][TABLEAU] = "57";
            grille[446][TABLEAU] = "57";
            grille[447][TABLEAU] = "57";
            grille[448][TABLEAU] = "58";
            grille[449][TABLEAU] = "58";
            grille[450][TABLEAU] = "58";
            grille[451][TABLEAU] = "58";
            grille[452][TABLEAU] = "59";
            grille[453][TABLEAU] = "59";
            grille[454][TABLEAU] = "59";
            grille[455][TABLEAU] = "59";
            grille[456][TABLEAU] = "60";
            grille[457][TABLEAU] = "60";
            grille[458][TABLEAU] = "60";
            grille[459][TABLEAU] = "60";
            grille[460][TABLEAU] = "61";
            grille[461][TABLEAU] = "61";
            grille[462][TABLEAU] = "61";
            grille[463][TABLEAU] = "61";
            grille[464][TABLEAU] = "62";
            grille[465][TABLEAU] = "62";
            grille[466][TABLEAU] = "62";
            grille[467][TABLEAU] = "62";
            grille[468][TABLEAU] = "63";
            grille[469][TABLEAU] = "63";
            grille[470][TABLEAU] = "63";
            grille[471][TABLEAU] = "63";
            grille[472][TABLEAU] = "64";
            grille[473][TABLEAU] = "64";
            grille[474][TABLEAU] = "64";
            grille[475][TABLEAU] = "64";
            grille[476][TABLEAU] = "65";
            grille[477][TABLEAU] = "65";
            grille[478][TABLEAU] = "65";
            grille[479][TABLEAU] = "65";
            grille[480][TABLEAU] = "66";
            grille[481][TABLEAU] = "66";
            grille[482][TABLEAU] = "66";
            grille[483][TABLEAU] = "66";
            grille[484][TABLEAU] = "67";
            grille[485][TABLEAU] = "67";
            grille[486][TABLEAU] = "67";
            grille[487][TABLEAU] = "67";
            grille[488][TABLEAU] = "68";
            grille[489][TABLEAU] = "68";
            grille[490][TABLEAU] = "68";
            grille[491][TABLEAU] = "68";
            grille[492][TABLEAU] = "69";
            grille[493][TABLEAU] = "69";
            grille[494][TABLEAU] = "69";
            grille[495][TABLEAU] = "69";
            grille[496][TABLEAU] = "70";
            grille[497][TABLEAU] = "70";
            grille[498][TABLEAU] = "70";
            grille[499][TABLEAU] = "70";
            grille[500][TABLEAU] = "71";
            grille[501][TABLEAU] = "71";
            grille[502][TABLEAU] = "71";
            grille[503][TABLEAU] = "71";
            grille[504][TABLEAU] = "72";
            grille[505][TABLEAU] = "72";
            grille[506][TABLEAU] = "72";
            grille[507][TABLEAU] = "72";
            grille[508][TABLEAU] = "73";
            grille[509][TABLEAU] = "73";
            grille[510][TABLEAU] = "73";
            grille[511][TABLEAU] = "73";
            


            foreach (int s in pageListe)
            {
                grille[s][PAGE] = s.ToString();
                grille[s * 1][PAGE] = s.ToString();
                grille[s * 2][PAGE] = s.ToString();
                grille[s * 2 + 1][PAGE] = s.ToString();
                grille[s * 4][PAGE] = s.ToString();
                grille[s * 4 + 1][PAGE] = s.ToString();
                grille[s * 4 + 2][PAGE] = s.ToString();
                grille[s * 4 + 3][PAGE] = s.ToString();
            }
            this.Text = NomPrograme;
            ChoixSosaComboBox.Text = "";
            AscendantDeTb.Text = "";

            PreparerPar.Text = "";
            Modifier = false;
            int rowLength = grille.Length;



            
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
            string Fichier = "grille.txt";
            using (StreamWriter ligne = File.CreateText(Fichier))
                //ligne.WriteLine("SOSA" + " " + "PAGE");

                for (int f = 0; f < 512; f++)
                {
                    ligne.WriteLine(grille[f][SOSA] + " " + grille[f][PAGE] + " " + grille[f][NOM]);
                }

        }
/**************************************************************************************************************/
        private void    Entete(ref PdfDocument document, ref XGraphics gfx, ref PdfPage page)
        {
            XFont font8 = new XFont("Arial", 8, XFontStyle.Bold);
            double x = POUCE * .5;
            double xx = POUCE * 4.5;
            double y = POUCE * 1;

            /**************************************************************************/
            //* Pour le développement marge millieu table matière                     */   
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
            et.DrawString("Nom", font8, XBrushes.Black, rect, XStringFormats.TopLeft);
            rect = new XRect(x + 219, y, 50, 10);
            et.DrawString("Tableau", font8, XBrushes.Black, rect, XStringFormats.TopLeft);

            rect = new XRect(xx, y, 170, 10);
            et.DrawString("SOSA", font8, XBrushes.Black, rect, XStringFormats.TopLeft);
            rect = new XRect(xx + 100, y, 170, 10);
            et.DrawString("Nom", font8, XBrushes.Black, rect, XStringFormats.TopLeft);
            rect = new XRect(xx + 219, y, 50, 10);
            et.DrawString("Tableau", font8, XBrushes.Black, rect, XStringFormats.TopLeft);
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
                        ligne.WriteLine("Nom   =" + "Nom " + index.ToString());
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
                gfx.DrawString(grille[sosa][TABLEAU], font, XBrushes.Black, rect, XStringFormats.Center);
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
                gfx.DrawString(grille[sosa][TABLEAU], font, XBrushes.Black, rect, XStringFormats.Center);
            }
            return;
        }
        private string  DessinerPage(ref PdfDocument document, ref XGraphics gfx, int sosa, bool fleche)
        {
            //int inch = 72 // 72 pointCreatePage
            XUnit pouce = XUnit.FromInch(1);
            XPen pen = new XPen(XColor.FromArgb(0, 0, 0),2);
            XPen penG = new XPen(XColor.FromArgb(100, 100, 100), 1);
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
            //* Pour le développement                                                 */   
            /**************************************************************************/
            /*
            gfx.DrawRectangle(pen, pouce * 0.5, 0.5 * pouce, pouce * 10, pouce * 7.5); //' x1,y1,x2,y2  cadrage de page
            */
            // FIN
            double largeurBoite = pouce * 2.05;
            double hauteurBoite = pouce * .75;
            double hauteurBoiteMini = pouce * .25;
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

            double hauteurLigne = pouce * .125;
            double positionLieu = .16 * pouce; // position Lieu par rapport date mariage = .16 * pouce; // par rapport date mariage
            // position des ligne au 1/4 pouce
            double[] Ligne = new double[65];
            for (f = 0; f < 65; f++)
            {
                double l = hauteurLigne * f;
                Ligne[f] = l;
            }
            /**************************************************************************/
            //* Pour le développement dessine colonnes                                */   
            /**************************************************************************/
            /*
            XPen penLigne = new XPen(XColor.FromArgb(200, 200, 255), 0.5);
            for (f = 0; f < 65; f++)
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
            // FIN

            /**************************************************************************/
            // Position des boites
            double[,] positionBoite = new double[16, 2]; // en pouce
            {
                // boite 1
                positionBoite[1, 0] = Col2 + 2;
                positionBoite[1, 1] = Ligne[32] + (pouce * .25 / 2);
                // boite 2
                positionBoite[2, 0] = Col4 + 2;
                positionBoite[2, 1] = Ligne[18] + (pouce * .25 / 2);
                // boite 3
                positionBoite[3, 0] = Col4 + 2;
                positionBoite[3, 1] = Ligne[46] + (pouce * .25 / 2);
                // boite 4
                positionBoite[4, 0] = Col6 + 2;
                positionBoite[4, 1] = Ligne[12];
                // boite 5
                positionBoite[5, 0] = Col6 + 2;
                positionBoite[5, 1] = Ligne[26];
                // boite 6
                positionBoite[6, 0] = Col6 + 2;
                positionBoite[6, 1] = Ligne[40];
                // boite 7
                positionBoite[7, 0] = Col6 + 2;
                positionBoite[7, 1] = Ligne[54];
                
                // boite 8
                positionBoite[8, 0] = Col8 + 2;
                positionBoite[8, 1] = Ligne[11];
                // boite 9
                positionBoite[9, 0] = Col8 + 2;
                positionBoite[9, 1] = Ligne[17];
                // boite 10
                positionBoite[10, 0] = Col8 + 2;
                positionBoite[10, 1] = Ligne[25];
                // boite 11
                positionBoite[11, 0] = Col8 + 2;
                positionBoite[11, 1] = Ligne[31];
                // boite 12
                positionBoite[12, 0] = Col8 + 2;
                positionBoite[12, 1] = Ligne[39];
                // boite 13
                positionBoite[13, 0] = Col8 + 2;
                positionBoite[13, 1] = Ligne[45];
                // boite 14
                positionBoite[14, 0] = Col8 + 2;
                positionBoite[14, 1] = Ligne[53];
                // boite 15
                positionBoite[15, 0] = Col8 + 2;
                positionBoite[15, 1] = y = Ligne[59];
            }
            // position mariage
            double[,] positionMariagexx = new double[7, 2]; // en pouce
            {
                int p = 10;
                positionMariagexx[0, 0] = Col4 + p;   // 2 3
                positionMariagexx[0, 1] = Ligne[34] + 17;
                positionMariagexx[1, 0] = Col6 + p;   // 4 5
                positionMariagexx[1, 1] = Ligne[20] + 17;
                positionMariagexx[2, 0] = Col6 + p;   // 6 7
                positionMariagexx[2, 1] = Ligne[48] + 17;
                positionMariagexx[3, 0] = Col8 + 12;   // 8 9
                positionMariagexx[3, 1] = Ligne[14] + 9;
                positionMariagexx[4, 0] = Col8 + p;   // 10 11
                positionMariagexx[4, 1] = Ligne[28] + 9;
                positionMariagexx[5, 0] = Col8 + 12;   // 12 13
                positionMariagexx[5, 1] = Ligne[42] + 9;
                positionMariagexx[6, 0] = Col8 + 12;   // 14 15
                positionMariagexx[6, 1] = Ligne[56] + 9;
            }
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
            // au de page
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
                gfx.DrawRoundedRectangle(pen, CouleurBloc, Col2, Ligne[8], largeurBoite, HauteurGeneration, Rond, Rond);
                str = "Génération " + grille[sosa][GENERATION];
                textLargeur = gfx.MeasureString(str, font8);
                if (grille[sosa][GENERATION] == "")
                {
                    gfx.DrawLine(penG, Col2 + (largeurBoite / 2) + (textLargeur.Width / 2) +2, y + 12, Col2 + (largeurBoite / 2) + (textLargeur.Width / 2) + 20, y + 12);
                }
                gfx.DrawString(str, font8, XBrushes.Black, Col2 + (largeurBoite / 2) - textLargeur.Width / 2, y + 12);

                //génération 2
                g = new XRect(Col4, 50, largeurBoite, 20);
                gfx.DrawRoundedRectangle(pen, CouleurBloc, Col4, y, largeurBoite, HauteurGeneration, Rond, Rond);
                str = "Génération " + grille[sosa * 2][GENERATION];
                textLargeur = gfx.MeasureString(str, font8);
                if (grille[sosa * 2][GENERATION] == "")
                {
                    gfx.DrawLine(penG, Col4 + (largeurBoite / 2) + (textLargeur.Width / 2) + 2, y + 12, Col4 + (largeurBoite / 2) + (textLargeur.Width / 2) + 20, y + 12);
                }
                gfx.DrawString(str, font8, XBrushes.Black, Col4 + (largeurBoite / 2) - textLargeur.Width / 2, y + 12);

                //génération 3
                g = new XRect(Col6, 50, largeurBoite, 20);
                gfx.DrawRoundedRectangle(pen, CouleurBloc, Col6, y, largeurBoite, HauteurGeneration, Rond, Rond);
                str = "Génération " + grille[sosa * 4][GENERATION];
                textLargeur = gfx.MeasureString(str, font8);
                if (grille[sosa * 4][GENERATION] == "")
                {
                    gfx.DrawLine(penG, Col6 + (largeurBoite / 2) + (textLargeur.Width / 2) + 2, y + 12, Col6 + (largeurBoite / 2) + (textLargeur.Width / 2) + 20, y + 12);
                }
                gfx.DrawString(str, font8, XBrushes.Black, Col6 + (largeurBoite / 2) - textLargeur.Width / 2, y + 12);

                //génération 4
                int s = sosa * 8;
                if (s < 512)
                {
                    g = new XRect(Col8, 50, largeurBoite, 20);
                    gfx.DrawRoundedRectangle(pen, CouleurBloc, Col8, y, largeurBoite, HauteurGeneration, Rond, Rond);
                    str = "Génération " + grille[s][GENERATION];
                    textLargeur = gfx.MeasureString(str, font8);
                    if (grille[s][GENERATION] == "")
                    {
                        gfx.DrawLine(penG, Col8 + (largeurBoite / 2) + (textLargeur.Width / 2) + 2, y + 12, Col8 + (largeurBoite / 2) + (textLargeur.Width / 2) + 20, y + 12);
                    }
                    gfx.DrawString(str, font8, XBrushes.Black, Col8 + (largeurBoite / 2) - textLargeur.Width / 2, y + 12);
                }
            }
            // dessine boite
            {
                gfx.DrawRoundedRectangle(pen, CouleurBloc, Col2, Ligne[32] + (pouce * .25 / 2), largeurBoite, hauteurBoite, 10, 10); // Boite sosa 1
                if(sosa != 1) {
                    gfx.DrawRoundedRectangle(pen, CouleurBloc, Col2, Ligne[43], largeurBoite, hauteurBoiteMini, 10, 10); //  Boite sosa 1 conjoint
                }
                gfx.DrawRoundedRectangle(pen, CouleurBloc, Col4, Ligne[18] + (pouce * .25 / 2), largeurBoite, hauteurBoite, 10, 10); // Boite sosa 2
                gfx.DrawRoundedRectangle(pen, CouleurBloc, Col4, Ligne[46] + (pouce * .25 / 2), largeurBoite, hauteurBoite, 10, 10); // Boite sosa 3
                gfx.DrawRoundedRectangle(pen, CouleurBloc, Col6, Ligne[12], largeurBoite, hauteurBoite, 10, 10); // Boite sosa 4
                gfx.DrawRoundedRectangle(pen, CouleurBloc, Col6, Ligne[26], largeurBoite, hauteurBoite, 10, 10); // Boite sosa 5
                gfx.DrawRoundedRectangle(pen, CouleurBloc, Col6, Ligne[40], largeurBoite, hauteurBoite, 10, 10); // Boite sosa 6
                gfx.DrawRoundedRectangle(pen, CouleurBloc, Col6, Ligne[54], largeurBoite, hauteurBoite, 10, 10); // Boite sosa 7
                int s = sosa * 8;
                if (s < 512)
                {
                    gfx.DrawRoundedRectangle(pen, CouleurBloc, Col8, Ligne[11], largeurBoite, hauteurBoiteMini, 10, 10); // Boite sosa 8
                    gfx.DrawRoundedRectangle(pen, CouleurBloc, Col8, Ligne[17], largeurBoite, hauteurBoiteMini, 10, 10); // Boite sosa 9
                    gfx.DrawRoundedRectangle(pen, CouleurBloc, Col8, Ligne[25], largeurBoite, hauteurBoiteMini, 10, 10); // Boite sosa 10
                    gfx.DrawRoundedRectangle(pen, CouleurBloc, Col8, Ligne[31], largeurBoite, hauteurBoiteMini, 10, 10); // Boite sosa 11
                    gfx.DrawRoundedRectangle(pen, CouleurBloc, Col8, Ligne[39], largeurBoite, hauteurBoiteMini, 10, 10); // Boite sosa 12
                    gfx.DrawRoundedRectangle(pen, CouleurBloc, Col8, Ligne[45], largeurBoite, hauteurBoiteMini, 10, 10); // Boite sosa 13
                    gfx.DrawRoundedRectangle(pen, CouleurBloc, Col8, Ligne[53], largeurBoite, hauteurBoiteMini, 10, 10); // Boite sosa 14
                    gfx.DrawRoundedRectangle(pen, CouleurBloc, Col8, Ligne[59], largeurBoite, hauteurBoiteMini, 10, 10); // Boite sosa 15
                }
            }
            // dessine ligne
            { 
                gfx.DrawLine(pen, Col3, Ligne[36], Col4 + 8, Ligne[36]); // Horizontal 1
                gfx.DrawLine(pen, Col5, Ligne[22], Col6 + 8, Ligne[22]); // Horizontal 2
                gfx.DrawLine(pen, Col5, Ligne[50], Col6 + 8, Ligne[50]); // Horizontal 3

                int s = sosa * 8;
                if (s < 512)
                {
                    gfx.DrawLine(pen, Col7, Ligne[14] + (pouce * .25 / 2), Col8 + 8, Ligne[14] + (pouce * .25 / 2)); // Horizontal 4
                    gfx.DrawLine(pen, Col7, Ligne[28] + (pouce * .25 / 2), Col8 + 8, Ligne[28] + (pouce * .25 / 2)); // Horizontal 5
                    gfx.DrawLine(pen, Col7, Ligne[42] + (pouce * .25 / 2), Col8 + 8, Ligne[42] + (pouce * .25 / 2)); // Horizontal 6
                    gfx.DrawLine(pen, Col7, Ligne[56] + (pouce * .25 / 2), Col8 + 8, Ligne[56] + (pouce * .25 / 2)); // Horizontal 7
                }
                if(sosa != 1) {
                    gfx.DrawLine(pen, Col2 + 8, Ligne[38] + 9, Col2 + 8, Ligne[44] - 9); // vertical sosa 1 conjoint
                }
                gfx.DrawLine(pen, Col4 + 8, Ligne[24] + 9, Col4 + 8, Ligne[48] - 9); // vertical 2 3
                gfx.DrawLine(pen, Col6 + 8, Ligne[18],  Col6 + 8, Ligne[26]); // vertical 4 5
                gfx.DrawLine(pen, Col6 + 8, Ligne[46], Col6 + 8, Ligne[54]); // vertical 6 7

                if (s < 512)
                {
                    gfx.DrawLine(pen, Col8 + 8, Ligne[13], Col8 + 8, Ligne[17]); // vertical 8 9
                    gfx.DrawLine(pen, Col8 + 8, Ligne[27], Col8 + 8, Ligne[31]); // vertical 10 11
                    gfx.DrawLine(pen, Col8 + 8, Ligne[41], Col8 + 8, Ligne[45]); // vertical 12 13
                    gfx.DrawLine(pen, Col8 + 8, Ligne[55], Col8 + 8, Ligne[59]); // vertical 14 15
                }
            }
            tf.Alignment = XParagraphAlignment.Right;
            int RetraitSosa = 20;
            XRect rect = new XRect();
            // Dessiner les boites
            // sosa 1 à 7
            for (f = 1; f < 8; f++) {
                rect = new XRect(positionBoite[f, 0] - RetraitSosa, positionBoite[f, 1], 15, 10);
                if ( sosa == 0 )
                {
                    gfx.DrawLine(penG, positionBoite[f, 0] - RetraitSosa + 3, positionBoite[f, 1] + 10, positionBoite[f, 0] - RetraitSosa + 11, positionBoite[f, 1] + 10);
                }
                else
                {
                    tf.DrawString(sosaBoite[f], font8B, XBrushes.Black, rect, XStringFormats.TopLeft);
                }
                gfx.DrawString("N", font8B, XBrushes.Black, positionBoite[f, 0], positionBoite[f, 1] + 20);
                gfx.DrawString("L", font8B, XBrushes.Black, positionBoite[f, 0], positionBoite[f, 1] + 30);
                gfx.DrawString("D", font8B, XBrushes.Black, positionBoite[f, 0], positionBoite[f, 1] + 40);
                gfx.DrawString("L", font8B, XBrushes.Black, positionBoite[f, 0], positionBoite[f, 1] + 50);
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
                gfx.DrawString("M", font8B, XBrushes.Black, Col2 + 10, Ligne[40] + 6, XStringFormats.Default);  // sosa 1 conjoint
                gfx.DrawString("L", font8B, XBrushes.Black, Col2 + 10, Ligne[40] + 6 + positionLieu, XStringFormats.Default);  // sosa 1 conjoint
            }
            gfx.DrawString("M", font8B, XBrushes.Black, Col4 + 10, Ligne[35] + 6, XStringFormats.Default);      // sosa 02-03
            gfx.DrawString("L", font8B, XBrushes.Black, Col4 + 10, Ligne[35] + 6 + positionLieu, XStringFormats.Default);      // sosa 02-03
            gfx.DrawString("M", font8B, XBrushes.Black, Col6 + 10, Ligne[21] + 6, XStringFormats.Default);      // sosa 04-05
            gfx.DrawString("L", font8B, XBrushes.Black, Col6 + 10, Ligne[21] + 6 + positionLieu, XStringFormats.Default);      // sosa 04-05
            gfx.DrawString("M", font8B, XBrushes.Black, Col6 + 10, Ligne[49] + 6, XStringFormats.Default);      // sosa 06-07
            gfx.DrawString("L", font8B, XBrushes.Black, Col6 + 10, Ligne[49] + 6 + positionLieu, XStringFormats.Default);      // sosa 06-07  

            if (sosa * 8 < 512)
            {
                gfx.DrawString("M", font8B, XBrushes.Black, Col8 + 10, Ligne[14] + 6, XStringFormats.Default);  // sosa 08-09
                gfx.DrawString("L", font8B, XBrushes.Black, Col8 + 10, Ligne[14] + 6 + positionLieu, XStringFormats.Default);  // sosa 08-09
                gfx.DrawString("M", font8B, XBrushes.Black, Col8 + 10, Ligne[28] + 6, XStringFormats.Default);  // sosa 10-11
                gfx.DrawString("L", font8B, XBrushes.Black, Col8 + 10, Ligne[28] + 6 + positionLieu, XStringFormats.Default);  // sosa 10-11
                gfx.DrawString("M", font8B, XBrushes.Black, Col8 + 10, Ligne[42] + 6, XStringFormats.Default);  // sosa 12-13
                gfx.DrawString("L", font8B, XBrushes.Black, Col8 + 10, Ligne[42] + 6 + positionLieu, XStringFormats.Default);  // sosa 12-13
                gfx.DrawString("M", font8B, XBrushes.Black, Col8 + 10, Ligne[56] + 6, XStringFormats.Default);  // sosa 14-15
                gfx.DrawString("L", font8B, XBrushes.Black, Col8 + 10, Ligne[56] + 6 + positionLieu, XStringFormats.Default);  // sosa 14-15
            }
            //}

            int xInfo = 7;
            if (sosa == 0)
            {
                int largeurLigne = 135; // 
                for (f = 1; f < 8; f++)
                {
                    gfx.DrawLine(penG, positionBoite[f, 0], positionBoite[f, 1] + 11, positionBoite[f, 0] + 142, positionBoite[f, 1] + 11);
                    
                    gfx.DrawLine(penG, positionBoite[f, 0] + xInfo, positionBoite[f, 1] + 21, positionBoite[f, 0] + xInfo + largeurLigne, positionBoite[f, 1] + 21);
                    
                    gfx.DrawLine(penG, positionBoite[f, 0] + xInfo, positionBoite[f, 1] + 31, positionBoite[f, 0] + xInfo + largeurLigne, positionBoite[f, 1] + 31);
                    
                    gfx.DrawLine(penG, positionBoite[f, 0] + xInfo, positionBoite[f, 1] + 41, positionBoite[f, 0] + xInfo + largeurLigne, positionBoite[f, 1] + 41);
                    
                    gfx.DrawLine(penG, positionBoite[f, 0] + xInfo, positionBoite[f, 1] + 51, positionBoite[f, 0] + xInfo + largeurLigne, positionBoite[f, 1] + 51);
                }
                int p = 18;
                int l = 140;
                gfx.DrawLine(penG, Col4 + p, Ligne[35], Col4 + l, Ligne[35]);
                gfx.DrawLine(penG, Col4 + p, Ligne[37], Col4 + l, Ligne[37]);

                gfx.DrawLine(penG, Col6 + p, Ligne[21], Col6 + l, Ligne[21]);
                gfx.DrawLine(penG, Col6 + p, Ligne[23], Col6 + l, Ligne[23]);
                gfx.DrawLine(penG, Col6 + p, Ligne[49], Col6 + l, Ligne[49]);
                gfx.DrawLine(penG, Col6 + p, Ligne[51], Col6 + l, Ligne[51]);


                gfx.DrawLine(penG, Col8 + p, Ligne[14] + 5, Col8 + l, Ligne[14] + 5);
                gfx.DrawLine(penG, Col8 + p, Ligne[16] + 3, Col8 + l, Ligne[16] + 3);
                gfx.DrawLine(penG, Col8 + p, Ligne[28] + 5, Col8 + l, Ligne[28] + 5);
                gfx.DrawLine(penG, Col8 + p, Ligne[30] + 3, Col8 + l, Ligne[30] + 3);
                gfx.DrawLine(penG, Col8 + p, Ligne[42] + 5, Col8 + l, Ligne[42] + 5);
                gfx.DrawLine(penG, Col8 + p, Ligne[44] + 3, Col8 + l, Ligne[44] + 3);
                gfx.DrawLine(penG, Col8 + p, Ligne[56] + 5, Col8 + l, Ligne[56] + 5);
                gfx.DrawLine(penG, Col8 + p, Ligne[58] + 3, Col8 + l, Ligne[58] + 3);

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
                        if (grille[sosaConjoint][NOM] == "")
                        {
                            gfx.DrawLine(penG, Col2 + 10, Ligne[60], Col2 + 142, Ligne[60]);
                        }
                        string rt = RacoucirNom(grille[sosaConjoint][NOM], ref gfx);
                        gfx.DrawString(rt, font8B, XBrushes.Black, Col2 + 2, Ligne[43] + 12, XStringFormats.Default);

                    }
                    for (f = 1; f < 8; f++)
                    {

                        // Nom
                        if (grille[sosaIndex[f]][NOM] == "")
                        {
                            gfx.DrawLine(penG, positionBoite[f, 0], positionBoite[f, 1] + 10, positionBoite[f, 0] + 142, positionBoite[f, 1] + 10);
                        }
                        string rt = RacoucirNom(grille[sosaIndex[f]][NOM], ref gfx);
                        gfx.DrawString(rt, font8B, XBrushes.Black, positionBoite[f, 0], positionBoite[f, 1] + 10, XStringFormats.Default);
                        // Né le 
                        if (grille[sosaIndex[f]][NELE] == "")
                        {
                            gfx.DrawLine(penG, positionBoite[f, 0] + 9, positionBoite[f, 1] + 20, positionBoite[f, 0] + largeurLigne, positionBoite[f, 1] + 20);
                        }
                        rt = RacoucirTexte(grille[sosaIndex[f]][NELE], ref gfx);
                        gfx.DrawString(rt, font8, XBrushes.Black, positionBoite[f, 0] + xInfo, positionBoite[f, 1] + 20, XStringFormats.Default);
                        // Né endroit
                        if (grille[sosaIndex[f]][NELIEU] == "")
                        {
                            gfx.DrawLine(penG, positionBoite[f, 0] + 9, positionBoite[f, 1] + 30, positionBoite[f, 0] + largeurLigne, positionBoite[f, 1] + 30);
                        }
                        rt = RacoucirTexte(grille[sosaIndex[f]][NELIEU], ref gfx);
                        gfx.DrawString(rt, font8, XBrushes.Black, positionBoite[f, 0] + xInfo, positionBoite[f, 1] + 30, XStringFormats.Default);
                        // Décédé le 
                        if (grille[sosaIndex[f]][DELE] == "")
                        {
                            gfx.DrawLine(penG, positionBoite[f, 0] + 9, positionBoite[f, 1] + 40, positionBoite[f, 0] + largeurLigne, positionBoite[f, 1] + 40);
                        }
                        rt = RacoucirTexte(grille[sosaIndex[f]][DELE], ref gfx);
                        gfx.DrawString(rt, font8, XBrushes.Black, positionBoite[f, 0] + xInfo, positionBoite[f, 1] + 40, XStringFormats.Default);
                        // Décédé endroit
                        if (grille[sosaIndex[f]][DELIEU] == "")
                        {
                            gfx.DrawLine(penG, positionBoite[f, 0] + 9, positionBoite[f, 1] + 50, positionBoite[f, 0] + largeurLigne, positionBoite[f, 1] + 50);
                        }
                        rt = RacoucirTexte(grille[sosaIndex[f]][DELIEU], ref gfx);
                        gfx.DrawString(rt, font8, XBrushes.Black, positionBoite[f, 0] + xInfo, positionBoite[f, 1] + 50, XStringFormats.Default);
                    }
                    for (f = 8; f < 16; f++)
                    {
                        if (sosaIndex[f] < 512)
                        {
                            if (grille[f][NOM] == "")
                            {
                                gfx.DrawLine(penG, positionBoite[f, 0] + 2, positionBoite[f, 1] + 15, positionBoite[f, 0] + 2 + 140, positionBoite[f, 1] + 15);
                            }
                            string rt = RacoucirNom(grille[sosaIndex[f]][NOM], ref gfx);
                            gfx.DrawString(rt, font8B, XBrushes.Black, positionBoite[f, 0], positionBoite[f, 1] + 12, XStringFormats.Default);
                        }
                    }
                    int p = 18;
                    int l = 140;


                    // mariage 1 sosa = 1
                    if (sosa != 1 ) {
                        if (sosa%2 == 0 ) {
                        
                            if (grille[sosa][MALE] == "")
                            {
                                gfx.DrawLine(penG, Col2 + p, Ligne[39] + 9 + 6, Col2 + l, Ligne[39] + 9 + 6 );
                            }
                            gfx.DrawString(grille[sosa][MALE], font8, XBrushes.Black, Col2 + p, Ligne[40] + 6, XStringFormats.Default);
                            if (grille[sosa][MALIEU] == "")
                            {
                                gfx.DrawLine(penG, Col2 + p, Ligne[42], Col2 + l, Ligne[42]);
                            }
                            gfx.DrawString(grille[sosa][MALIEU], font8, XBrushes.Black, Col2 + p, Ligne[40] + 6 + positionLieu, XStringFormats.Default);
                        } else {
                            if (grille[sosa-1][MALE] == "") 
                            {
                                gfx.DrawLine(penG, Col2 + p, Ligne[39] + 9 + 6, Col2 + l, Ligne[39] + 9 + 6 );
                            }
                            gfx.DrawString(grille[sosa-1][MALE], font8, XBrushes.Black, Col2 + p, Ligne[40] + 6, XStringFormats.Default);
                            if (grille[sosa-1][MALIEU] == "")
                            {
                                gfx.DrawLine(penG, Col2 + p, Ligne[42], Col2 + l, Ligne[42]);
                            }
                            gfx.DrawString(grille[sosa-1][MALIEU], font8, XBrushes.Black, Col2 + p, Ligne[40] + 6 + positionLieu, XStringFormats.Default);
                        }

                    }

                    // mariage 2 3
                    if (grille[sosa * 2][MALE] == "")
                    {
                        gfx.DrawLine(penG, Col4 + p, Ligne[35] + 6, Col4 + l, Ligne[35] + 6 );
                    }
                    gfx.DrawString(grille[sosa * 2][MALE], font8, XBrushes.Black, Col4 + p, Ligne[35] + 6, XStringFormats.Default);
                    if (grille[sosa * 2][MALIEU] == "")
                    {
                        gfx.DrawLine(penG, Col4 + p, Ligne[37], Col4 + l, Ligne[37]);
                    }
                    gfx.DrawString(grille[sosa * 2][MALIEU], font8, XBrushes.Black, Col4 + p, Ligne[35] + 6 + positionLieu, XStringFormats.Default);

                    // mariage 4 5
                    if (grille[sosa * 4][MALE] == "")
                    {
                        gfx.DrawLine(penG, Col6 + p, Ligne[21] + 6, Col6 + l, Ligne[21] + 6);
                    }
                    gfx.DrawString(grille[sosa * 4][MALE], font8, XBrushes.Black, Col6 + p , Ligne[21] + 6, XStringFormats.Default);

                    if (grille[sosa * 4][MALIEU] == "")
                    {
                        gfx.DrawLine(penG, Col6 + p, Ligne[21] + 6 + positionLieu, Col6 + l, Ligne[21] + 6 + positionLieu);
                    }
                    gfx.DrawString(grille[sosa * 4][MALIEU], font8, XBrushes.Black, Col6 + p, Ligne[21] + 6 + positionLieu, XStringFormats.Default);
                    // mariage 6 7
                    if (grille[sosa * 4 + 2][MALE] == "")
                    {
                        gfx.DrawLine(penG, Col6 + p, Ligne[49] + 6, Col6 + l, Ligne[49] + 6);
                    }
                    gfx.DrawString(grille[sosa * 4 + 2][MALE], font8, XBrushes.Black, Col6 + p, Ligne[49] + 6, XStringFormats.Default);
                    if (grille[sosa * 4 + 2][MALIEU] == "")
                    {
                        gfx.DrawLine(penG, Col6 + p, Ligne[49] + 6 + positionLieu, Col6 + l, Ligne[49] + 6 + positionLieu);
                    }
                    gfx.DrawString(grille[sosa * 4 + 2][MALIEU], font8, XBrushes.Black, Col6 + p, Ligne[49] + 6 + positionLieu, XStringFormats.Default);
                    // mariage 8 9
                    int s = sosa * 8;
                    if (s < 512)
                    {
                        if (grille[s][MALE] == "")
                        {
                            gfx.DrawLine(penG, Col8 + p, Ligne[14] + 6, Col8 + l, Ligne[14] + 5);
                        }
                        gfx.DrawString(grille[s][MALE], font8, XBrushes.Black, Col8 + p, Ligne[14] + 6, XStringFormats.Default);
                        if (grille[s][MALIEU] == "")
                        {
                            gfx.DrawLine(penG, Col8 + p, Ligne[14] + 6 + positionLieu, Col8 + l, Ligne[14] + 6 + positionLieu);
                        }
                        gfx.DrawString(grille[s][MALIEU], font8, XBrushes.Black, Col8 + p, Ligne[14] + 6 + positionLieu, XStringFormats.Default);
                    }
                    // mariage 10 11
                    s = sosa * 8 + 2;
                    if (s < 512)
                    {
                        if (grille[s][MALE] == "")
                        {
                            gfx.DrawLine(penG, Col8 + p, Ligne[28] + 6, Col8 + l, Ligne[28] + 6);
                        }
                        gfx.DrawString(grille[s][MALE], font8, XBrushes.Black, Col8 + p, Ligne[28] + 6, XStringFormats.Default);
                        if (grille[s][MALIEU] == "")
                        {
                            gfx.DrawLine(penG, Col8 + p, Ligne[28] + 6 + positionLieu, Col8 + l, Ligne[28] + 6 + positionLieu);
                        }
                        gfx.DrawString(grille[s][MALIEU], font8, XBrushes.Black, Col8 + p, Ligne[28] + 6 + positionLieu, XStringFormats.Default);
                    }
                    // mariage 12 13
                    s = sosa * 8 + 4;
                    if (s < 512)
                    {
                        if (grille[s][MALE] == "")
                        {
                            gfx.DrawLine(penG, Col8 + p, Ligne[42] + 6, Col8 + l, Ligne[42] + 6);
                        }
                        gfx.DrawString(grille[sosa * 8 + 4][MALE], font8, XBrushes.Black, Col8 + p, Ligne[42] + 6, XStringFormats.Default);
                        if (grille[s][MALIEU] == "")
                        {
                            gfx.DrawLine(penG, Col8 + p, Ligne[42] + 6 + positionLieu, Col8 + l, Ligne[42] + 6 + positionLieu);
                        }
                        gfx.DrawString(grille[sosa * 8 + 4][MALIEU], font8, XBrushes.Black, Col8 + p, Ligne[42] + 6 + positionLieu, XStringFormats.Default);
                    }
                    // mariage 14 15
                    s = sosa * 8 + 6;
                    if (s < 512)
                    {
                        if (grille[s][MALE] == "") 
                        {
                            gfx.DrawLine(penG, Col8 + p, Ligne[56] + 6, Col8 + l, Ligne[56] +6);
                        }
                        gfx.DrawString(grille[s][MALE], font8, XBrushes.Black, Col8 + p, Ligne[56] + 5, XStringFormats.Default);
                        if (grille[s][MALIEU] == "")
                        {
                            gfx.DrawLine(penG, Col8 + p, Ligne[56] + 6 + positionLieu, Col8 + l, Ligne[56] + 6 + positionLieu);
                        }
                        gfx.DrawString(grille[s][MALIEU], font8, XBrushes.Black, Col8 + p, Ligne[56] + 6 + positionLieu, XStringFormats.Default);
                    }
                }
            }
            //dessiner  flèche
            if (fleche)
            {
                FlecheGauche(gfx, font8, Col1, positionBoite[1, 1], hauteurBoite, sosa);
                if (sosa == 0)
                {
                    FlecheDroite(gfx, font8, Col10, positionBoite[08, 1], hauteurBoiteMini, 0);
                    FlecheDroite(gfx, font8, Col10, positionBoite[09, 1], hauteurBoiteMini, 0);
                    FlecheDroite(gfx, font8, Col10, positionBoite[10, 1], hauteurBoiteMini, 0);
                    FlecheDroite(gfx, font8, Col10, positionBoite[11, 1], hauteurBoiteMini, 0);
                    FlecheDroite(gfx, font8, Col10, positionBoite[12, 1], hauteurBoiteMini, 0);
                    FlecheDroite(gfx, font8, Col10, positionBoite[13, 1], hauteurBoiteMini, 0);
                    FlecheDroite(gfx, font8, Col10, Ligne[53], hauteurBoiteMini, 0);
                    FlecheDroite(gfx, font8, Col10, Ligne[59], hauteurBoiteMini, 0);
                }
                else
                {
                    FlecheDroite(gfx, font8, Col10, positionBoite[08, 1], hauteurBoiteMini, sosa * 8);
                    FlecheDroite(gfx, font8, Col10, positionBoite[09, 1], hauteurBoiteMini, sosa * 8 + 1);
                    FlecheDroite(gfx, font8, Col10, positionBoite[10, 1], hauteurBoiteMini, sosa * 8 + 2);
                    FlecheDroite(gfx, font8, Col10, positionBoite[11, 1], hauteurBoiteMini, sosa * 8 + 3);
                    FlecheDroite(gfx, font8, Col10, positionBoite[12, 1], hauteurBoiteMini, sosa * 8 + 4);
                    FlecheDroite(gfx, font8, Col10, positionBoite[13, 1], hauteurBoiteMini, sosa * 8 + 5);
                    FlecheDroite(gfx, font8, Col10, Ligne[53], hauteurBoiteMini, sosa * 8 + 6);
                    FlecheDroite(gfx, font8, Col10, Ligne[59], hauteurBoiteMini, sosa * 8 + 7);
                }
            }
            // Note 1
            rect = new XRect(Col1, Ligne[13], Col3 - Col1, hauteurLigne  * 18);
            //gfx.DrawRectangle(penB, rect);
            tf.Alignment = XParagraphAlignment.Justify;
            tf.DrawString(grille[sosa][NOTE1], font8, XBrushes.Black, rect, XStringFormats.TopLeft);




            // Note 2
            rect = new XRect(Col1, Ligne[46], Col3 - Col1, hauteurLigne  * 18);
            //gfx.DrawRectangle(penB, rect);
            tf.Alignment = XParagraphAlignment.Justify;
            tf.DrawString(grille[sosa][NOTE2], font8, XBrushes.Black, rect, XStringFormats.TopLeft);
            // bas de page
            if (fleche)
            {
                numeroTableau = grille[sosa][TABLEAU];
                gfx.DrawString("Tableau", font8, XBrushes.Black, Col8 + 100, Ligne[64], XStringFormats.TopLeft);

                if (numeroTableau != "")
                {
                    gfx.DrawString(numeroTableau, font8, XBrushes.Black, Col8 + 135, Ligne[64], XStringFormats.TopLeft);
                }
                else
                {
                    gfx.DrawString("_____", font8, XBrushes.Black, Col8 + 135, Ligne[64], XStringFormats.TopLeft);
                }
            }
              if (PreparerPar.Text != "")
            {
                gfx.DrawString("Préparé par " + PreparerPar.Text + " le " + DateLb.Text, font8, XBrushes.Black, Col1, Ligne[64], XStringFormats.Default);
            }
            XImage img = global::TableauAscendant.Properties.Resources.dapamv5_32png;
            
            XPen penDapam = new XPen(XColor.FromArgb(0, 0, 0), 2);
            XFont fontDapam = new XFont("Arial", 14, XFontStyle.Bold);
            XFont fontDesign = new XFont("Arial", 5.5, XFontStyle.Italic);
            gfx.DrawRoundedRectangle(penDapam,gris, pouce * 8.03, pouce * 7.82, 59, 20, 15, 15);
            gfx.DrawString("DAPAM", fontDapam, XBrushes.Black, pouce * 8.08, pouce * 8.025);
            gfx.DrawString("Design", fontDesign, XBrushes.Black, pouce * 7.75, pouce * 7.9);
            
            return numeroTableau;
        }
        private string  DessinerPatrilineairexxx(ref PdfDocument document, ref XGraphics gfx, int sosa, bool fleche)
        {
            //int inch = 72 // 72 pointCreatePage
            XUnit pouce = XUnit.FromInch(1);
            XPen pen = new XPen(XColor.FromArgb(0, 0, 0), 2);
            XPen penG = new XPen(XColor.FromArgb(100, 100, 100), 0.5);
            XPen penB = new XPen(XColor.FromArgb(255, 255, 255), 0.5);
            XFont fontT = new XFont("Arial", 14, XFontStyle.Regular);
            XFont font8 = new XFont("Arial", 8, XFontStyle.Regular);
            XFont font8B = new XFont("Arial", 8, XFontStyle.Bold);
            XBrush gris = new XSolidBrush(XColor.FromArgb(255, 255, 255));
            XBrush CouleurBloc = new XSolidBrush(XColor.FromArgb(RectangleSosa1.FillColor.R, RectangleSosa1.FillColor.G, RectangleSosa1.FillColor.B));

            XTextFormatter tf = new XTextFormatter(gfx);
            return "";
        }
        static void     ZXCV(string message, [CallerLineNumber] int lineNumber = 0, [CallerMemberName] string caller = null)
        {
            Console.WriteLine(lineNumber + " " + caller + " " + message);

            string Fichier = "01log.txt";
            using (StreamWriter ligne = File.AppendText(Fichier))
            {
                ligne.WriteLine(lineNumber + " " + caller + " " + message);
                
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
                                        if (s.Length > 7) grille[index][SOSA] = s.Substring(7);
                                        s = sr.ReadLine();
                                    }
                                    if (s.Substring(0, 7) == "Nom   =")
                                    {
                                        if (s.Length > 7 ) grille[index][NOM] = s.Substring(7);
                                        s = sr.ReadLine();
                                    }
                                    if ( s.Substring( 0,7) == "NeLe  =")
                                    {
                                        if (s.Length > 7) grille[index][NELE] = s.Substring(7);
                                        s = sr.ReadLine();
                                    }
                                    if (s.Substring(0, 7) == "NeLieu=")
                                    {
                                        if (s.Length > 7) grille[index][NELIEU] = s.Substring(7);
                                        s = sr.ReadLine();
                                    }
                                    if (s.Substring(0, 7) == "DeLe  =")
                                    {
                                        if (s.Length > 7) grille[index][DELE] = s.Substring(7);
                                        s = sr.ReadLine();
                                    }
                                    if (s.Substring(0, 7) == "DeLieu=")
                                    {
                                        if (s.Length > 7) grille[index][DELIEU] = s.Substring(7);
                                        s = sr.ReadLine();
                                    }
                                    if (s.Substring(0, 7) == "MaLe  =")
                                    {
                                        if (s.Length > 7) grille[index][MALE] = s.Substring(7);
                                        s = sr.ReadLine();
                                    }
                                    if (s.Substring(0, 7) == "MaLieu=")
                                    {
                                        if (s.Length > 7) grille[index][MALIEU] = s.Substring(7);
                                        s = sr.ReadLine();
                                    }
                                    if (s.Substring(0, 7) == "NoteH =")
                                    {
                                        s = sr.ReadLine();
                                        while (s != "##FIN##")
                                        {
                                            if (grille[index][NOTE1] == "")
                                            {
                                                grille[index][NOTE1] = s;
                                                s = sr.ReadLine();
                                            }
                                            else
                                            {
                                                grille[index][NOTE1] = grille[index][NOTE1] + "\r\n" + s;
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
                                            if (grille[index][NOTE2] == "")
                                            {
                                                grille[index][NOTE2] = s;
                                                s = sr.ReadLine();
                                            }
                                            else
                                            {
                                                grille[index][NOTE2] = grille[index][NOTE2] + "\r\n" + s;
                                                s = sr.ReadLine();
                                            }
                                        }
                                        s = sr.ReadLine();
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
        private void RafraichirData()
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
            Sosa1NomTextBox.Text = grille[index][NOM];
            Sosa1NeTextBox.Text = grille[index][NELE];
            Sosa1NeEndroitTextBox.Text = grille[index][NELIEU];
            Sosa1DeTextBox.Text = grille[index][DELE];
            Sosa1DeEndroitTextBox.Text = grille[index][DELIEU];
            Sosa1MaTextBox.Text = grille[index][MALE];
            Sosa1MaEndroitTextBox.Text = grille[index][MALIEU];
            if (index > 1)
            {
                int i = index % 2;
                if (index % 2 == 0)
                {
                    SosaConjoint1NomTextBox.Text = grille[index + 1][NOM];
                    SosaConjoint1NomTextBox.Visible = true;
                    SosaConjoint1Label.Text = (index + 1).ToString();
                    SosaConjoint1Label.Visible = true;
                }
                else
                {
                    SosaConjoint1NomTextBox.Visible = false;
                    SosaConjoint1Label.Visible = false;
                }
            }
            Note1.Text = grille[index][NOTE1];
            Note2.Text = grille[index][NOTE2];
            GenerationAlb.Text = grille[index][GENERATION];

            index = sosaCourant * 2;
            Sosa2Label.Text = grille[index][SOSA];
            Sosa2NomTextBox.Text = grille[index][NOM];
            Sosa2NeTextBox.Text = grille[index][NELE];
            Sosa2NeEndroitTextBox.Text = grille[index][NELIEU];
            Sosa2DeTextBox.Text = grille[index][DELE];
            Sosa2DeEndroitTextBox.Text = grille[index][DELIEU];
            Sosa23MaTextBox.Text = grille[index][MALE];
            Sosa23MaEndroitTextBox.Text = grille[index][MALIEU];
            GenerationBlb.Text = grille[index][GENERATION];

            index = sosaCourant * 2 + 1;
            Sosa3Label.Text = grille[index][SOSA];
            Sosa3NomTextBox.Text = grille[index][NOM];
            Sosa3NeTextBox.Text = grille[index][NELE];
            Sosa3NeEndroitTextBox.Text = grille[index][NELIEU];
            Sosa3DeTextBox.Text = grille[index][DELE];
            Sosa3DeEndroitTextBox.Text = grille[index][DELIEU];

            index = sosaCourant * 4;
            Sosa4Label.Text = grille[index][SOSA];
            Sosa4NomTextBox.Text = grille[index][NOM];
            Sosa4NeTextBox.Text = grille[index][NELE];
            Sosa4NeEndroitTextBox.Text = grille[index][NELIEU];
            Sosa4DeTextBox.Text = grille[index][DELE];
            Sosa4DeEndroitTextBox.Text = grille[index][DELIEU];
            Sosa45MaTextBox.Text = grille[index][MALE];
            Sosa45MaLEndroitTextBox.Text = grille[index][MALIEU];
            GenerationClb.Text = grille[index][GENERATION];

            index = sosaCourant * 4 + 1;
            Sosa5Label.Text = grille[index][SOSA];
            Sosa5NomTextBox.Text = grille[index][NOM];
            Sosa5NeTextBox.Text = grille[index][NELE];
            Sosa5NeEndroitTextBox.Text = grille[index][NELIEU];
            Sosa5DeTextBox.Text = grille[index][DELE];
            Sosa5DeEndroitTextBox.Text = grille[index][DELIEU];

            index = sosaCourant * 4 + 2;
            Sosa6Label.Text = grille[index][SOSA];
            Sosa6NomTextBox.Text = grille[index][NOM];
            Sosa6NeTextBox.Text = grille[index][NELE];
            Sosa6NeEndroitTextBox.Text = grille[index][NELIEU];
            Sosa6DeTextBox.Text = grille[index][DELE];
            Sosa6DeEndroitTextBox.Text = grille[index][DELIEU];
            Sosa67MaTextBox.Text = grille[index][MALE];
            Sosa67MaEndroitTextBox.Text = grille[index][MALIEU];

            index = sosaCourant * 4 + 3;
            Sosa7Label.Text = grille[index][SOSA];
            Sosa7NomTextBox.Text = grille[index][NOM];
            Sosa7NeTextBox.Text = grille[index][NELE];
            Sosa7NeEndroitTextBox.Text = grille[index][NELIEU];
            Sosa7DeTextBox.Text = grille[index][DELE];
            Sosa7DeEndroitTextBox.Text = grille[index][DELIEU];
        }
        private void    RechercheID()
        {
            ContinuerBtn.Visible = false;
            List<string> IDListe = new List<string>();
            IDListe = GEDCOM.RechercheIndividu(NomRecherche.Text, PrenomRecherche.Text);
            string[] ligne = new string[5];
            ListViewItem itm;
            ChoixLV.Items.Clear();
            foreach (string info in IDListe)
            {

                ligne[0] = info;
                ligne[1] = GEDCOM.AvoirNom(info);
                ligne[2] = GEDCOM.AvoirPrenom(info);
                ligne[3] = ConvertirDate(GEDCOM.AvoirDateNaissance(info));
                ligne[4] = GEDCOM.AvoirEndroitNaissance(info);
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
                grille[index][SOSA] = index.ToString();
                grille[index][NOM] = Sosa1NomTextBox.Text;
                grille[index][NELE] = Sosa1NeTextBox.Text;
                grille[index][NELIEU] = Sosa1NeEndroitTextBox.Text;
                grille[index][DELE] = Sosa1DeTextBox.Text;
                grille[index][DELIEU] = Sosa1DeEndroitTextBox.Text;
                grille[index][NOTE1] = Note1.Text;
                grille[index][NOTE2] = Note2.Text;

                index = sosaCourant * 2;
                grille[index][SOSA] = index.ToString();
                grille[index][NOM] = Sosa2NomTextBox.Text;
                grille[index][NELE] = Sosa2NeTextBox.Text;
                grille[index][NELIEU] = Sosa2NeEndroitTextBox.Text;
                grille[index][DELE] = Sosa2DeTextBox.Text;
                grille[index][DELIEU] = Sosa2DeEndroitTextBox.Text;
                grille[index][MALE] = Sosa23MaTextBox.Text;
                grille[index][MALIEU] = Sosa23MaEndroitTextBox.Text;

                index = sosaCourant * 2 + 1;
                grille[index][SOSA] = index.ToString();
                grille[index][NOM] = Sosa3NomTextBox.Text;
                grille[index][NELE] = Sosa3NeTextBox.Text;
                grille[index][NELIEU] = Sosa3NeEndroitTextBox.Text;
                grille[index][DELE] = Sosa3DeTextBox.Text;
                grille[index][DELIEU] = Sosa3DeEndroitTextBox.Text;

                index = sosaCourant * 4;
                grille[index][SOSA] = index.ToString();
                grille[index][NOM] = Sosa4NomTextBox.Text;
                grille[index][NELE] = Sosa4NeTextBox.Text;
                grille[index][NELIEU] = Sosa4NeEndroitTextBox.Text;
                grille[index][DELE] = Sosa4DeTextBox.Text;
                grille[index][DELIEU] = Sosa4DeEndroitTextBox.Text;
                grille[index][MALE] = Sosa45MaTextBox.Text;
                grille[index][MALIEU] = Sosa45MaLEndroitTextBox.Text;

                index = sosaCourant * 4 + 1;

                grille[index][SOSA] = index.ToString();
                grille[index][NOM] = Sosa5NomTextBox.Text;
                grille[index][NELE] = Sosa5NeTextBox.Text;
                grille[index][NELIEU] = Sosa5NeEndroitTextBox.Text;
                grille[index][DELE] = Sosa5DeTextBox.Text;
                grille[index][DELIEU] = Sosa5DeEndroitTextBox.Text;
                grille[index][MALE] = "";
                grille[index][MALIEU] = "";

                index = sosaCourant * 4 + 2;
                grille[index][SOSA] = index.ToString();
                grille[index][NOM] = Sosa6NomTextBox.Text;
                grille[index][NELE] = Sosa6NeTextBox.Text;
                grille[index][NELIEU] = Sosa6NeEndroitTextBox.Text;
                grille[index][DELE] = Sosa6DeTextBox.Text;
                grille[index][DELIEU] = Sosa6DeEndroitTextBox.Text;
                grille[index][MALE] = Sosa67MaTextBox.Text;
                grille[index][MALIEU] = Sosa67MaEndroitTextBox.Text;

                index = sosaCourant * 4 + 3;
                grille[index][SOSA] = index.ToString();
                grille[index][NOM] = Sosa7NomTextBox.Text;
                grille[index][NELE] = Sosa7NeTextBox.Text;
                grille[index][NELIEU] = Sosa7NeEndroitTextBox.Text;
                grille[index][DELE] = Sosa7DeTextBox.Text;
                grille[index][DELIEU] = Sosa7DeEndroitTextBox.Text;
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
            Entete(ref document, ref gfx, ref page);

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
            for (int f = 1; f < (grille.GetLength(0)) ; f++)
            {
                tf = new XTextFormatter(gfx);
                if (Col == 1) x = POUCE * .5;
                if (Col == 2) x = POUCE * 4.5;
                tf.Alignment = XParagraphAlignment.Left;
                //rect = new XRect(x, y, 240, 10);
                //str = grille[f][NOM];
                if (grille[f][NOM].Length > 0 && grille[f][SOSA] != "0")
                {
                    // largeur maximum nom 240
                    XSize textLargeur = gfx.MeasureString(grille[f][NOM], font8);
                    if (textLargeur.Width > 220)
                    {
                        textLargeur.Width = 220;
                    }

                    // sosa
                    tf.Alignment = XParagraphAlignment.Right;
                    rect = new XRect(x, y, 15, 10);
                    tf.DrawString(grille[f][SOSA], font8, XBrushes.Black, rect, XStringFormats.TopLeft);
                    gfx.DrawLine(penD, x , y + 8, x + 15, y + 8);

                    //nom
                    tf.Alignment = XParagraphAlignment.Left;
                    rect = new XRect(x + 18, y, 220, 10);
                    tf.DrawString(grille[f][NOM], font8, XBrushes.Black, rect, XStringFormats.TopLeft);
                    gfx.DrawLine(penD, x + 18 , y + 8, x + 235, y + 8);

                    // tableau 
                    tf.Alignment = XParagraphAlignment.Right;
                    rect = new XRect(x + 240, y, 10, 10);
                    tf.DrawString(grille[f][TABLEAU], font8, XBrushes.Black, rect, XStringFormats.TopLeft);
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
                        Entete(ref document, ref gfx, ref page);

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
        private void    Triage(string[][] grille, int col)
        {
            Array.Sort(grille, delegate (object[] x, object[] y)
            {
                return (x[col] as IComparable).CompareTo(y[col]);
            }
            );
            for (int f = 1; f < 512; f++)
            {

                grille[f][NOM].TrimStart(" ".ToCharArray());
                grille[f][SOSA].TrimStart(" ".ToCharArray());
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
        public TableauAscendant(string a)
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
            
            string Fichier = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData) + "\\TableauAscendant\\TableauAscendant.ini";
  
            if (File.Exists(Fichier))
            {
                try
                {
                    using (StreamReader sr = File.OpenText(Fichier))
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
            }
            ChoixSosaComboBox.Text  = "";
            /*
            Sosa1MaLelb.Text = "";
            Sosa1MaEndroitTextlb.Text = "";
            Sosa1MaAvecTextlb.Text = "";
            */
            GenerationAlb.Text = "";
            GenerationBlb.Text = "";
            GenerationClb.Text = "";

            this.ChoixLV.MouseDoubleClick += new MouseEventHandler(ChoixLV_MouseDoubleClick);
            AfficherData();
            if (argument != "")
            {
                FichierCourant = argument;
                LireData();
            }
            FlecheGaucheRechercheButton.Visible = false;
            FlecheDroiteRechercheButton.Visible = false ;

            //FichierTest();

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
        private void    Sosa1NomTextBox_TextChanged(object sender, EventArgs e)
        {
            grille[sosaCourant][NOM] = Sosa1NomTextBox.Text;
            if (!LongeurNomtOk(grille[sosaCourant][NOM]))
            {
                Sosa1NomTextBox.BackColor = couleurTextTropLong;
            }
            else
            {
                Sosa1NomTextBox.BackColor = couleurChamp;
            }
        }
        private void Sosa1NomTextBox_KeyPress(object sender, KeyPressEventArgs e)
        {
            Modifier = true;
            this.Text = NomPrograme + "   *" + FichierCourant;
        }
        private void    Sosa1NeTextBox_TextChanged(object sender, EventArgs e)
        {
            grille[sosaCourant][NELE] = Sosa1NeTextBox.Text;
            bool rep  = ValiderDate(Sosa1NeTextBox.Text);
            if (rep)
            {
                Sosa1NeTextBox.BackColor = Color.White;
                if (!LongeurTextOk(grille[sosaCourant][NELE]))
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
        private void Sosa1NeTextBox_KeyPress(object sender, KeyPressEventArgs e)
        {
            Modifier = true;
            this.Text = NomPrograme + "   *" + FichierCourant;
        }
        private void    Sosa1NeEndroitTextBox_TextChanged(object sender, EventArgs e)
        {
            grille[sosaCourant][NELIEU] = Sosa1NeEndroitTextBox.Text;
            if (!LongeurTextOk(grille[sosaCourant][NELIEU]))
            {
                Sosa1NeEndroitTextBox.BackColor = couleurTextTropLong;
            } else
            {
                Sosa1NeEndroitTextBox.BackColor = couleurChamp;
            }
        }
        private void Sosa1NeEndroitTextBox_KeyPress(object sender, KeyPressEventArgs e)
        {
            Modifier = true;
            this.Text = NomPrograme + "   *" + FichierCourant;
        }
        private void    Sosa1DeTextBox_TextChanged(object sender, EventArgs e)
        {
            grille[sosaCourant][DELE] = Sosa1DeTextBox.Text;
            bool rep = ValiderDate(Sosa1DeTextBox.Text);
            if (rep)
            {
                Sosa1DeTextBox.BackColor = Color.White;
                if (!LongeurTextOk(grille[sosaCourant][DELE]))
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
        private void Sosa1DeTextBox_KeyPress(object sender, KeyPressEventArgs e)
        {
            Modifier = true;
            this.Text = NomPrograme + "   *" + FichierCourant;
        }
        private void Sosa1DeEndroitTextBox_TextChanged(object sender, EventArgs e)
        {
            grille[sosaCourant][DELIEU] = Sosa1DeEndroitTextBox.Text;
            if (!LongeurTextOk(grille[sosaCourant][DELIEU]))
            {
                Sosa1DeEndroitTextBox.BackColor = couleurTextTropLong;
            }
            else
            {
                Sosa1DeEndroitTextBox.BackColor = couleurChamp;
            }
        }
        private void Sosa1DeEndroitTextBox_KeyPress(object sender, KeyPressEventArgs e)
        {
            Modifier = true;
            this.Text = NomPrograme + "   *" + FichierCourant;
        }
        private void Sosa2NomTextBox_TextChanged(object sender, EventArgs e)
        {
            grille[sosaCourant * 2][NOM] = Sosa2NomTextBox.Text;
            Sosa2NomTextBox.BackColor = Color.White;
            if (!LongeurNomtOk(grille[sosaCourant * 2][NOM]))
            {
                Sosa2NomTextBox.BackColor = couleurTextTropLong;
            }
            else
            {
                Sosa2NomTextBox.BackColor = couleurChamp;
            }
        }
        private void Sosa2NomTextBox_KeyPress(object sender, KeyPressEventArgs e)
        {
            Modifier = true;
            this.Text = NomPrograme + "   *" + FichierCourant;
        }
        private void Sosa2NeTextBox_TextChanged(object sender, EventArgs e)
        {
            grille[sosaCourant * 2][NELE] = Sosa2NeTextBox.Text;
            bool rep = ValiderDate(Sosa2NeTextBox.Text);
            if (rep)
            {
                Sosa2NeTextBox.BackColor = Color.White;
                if (!LongeurTextOk(grille[sosaCourant][NELE]))
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
        private void Sosa2NeTextBox_KeyPress(object sender, KeyPressEventArgs e)
        {
            Modifier = true;
            this.Text = NomPrograme + "   *" + FichierCourant;
        }
        private void Sosa2NeEndroitTextBox_TextChanged(object sender, EventArgs e)
        {
            grille[sosaCourant * 2][NELIEU] = Sosa2NeEndroitTextBox.Text;
            if (!LongeurTextOk(Sosa2NeEndroitTextBox.Text))
            {
                Sosa2NeEndroitTextBox.BackColor = couleurTextTropLong;
            }
            else
            {
                Sosa2NeEndroitTextBox.BackColor = couleurChamp;
            }
        }
        private void Sosa2NeEndroitTextBox_KeyPress(object sender, KeyPressEventArgs e)
        {
            Modifier = true;
            this.Text = NomPrograme + "   *" + FichierCourant;
        }
        private void Sosa2DeTextBox_TextChanged(object sender, EventArgs e)
        {
            grille[sosaCourant * 2][DELE] = Sosa2DeTextBox.Text;
            bool rep = ValiderDate(Sosa2DeTextBox.Text);
            if (rep)
            {
                Sosa2DeTextBox.BackColor = Color.White;
                if (!LongeurTextOk(grille[sosaCourant][DELE]))
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
        private void Sosa2DeTextBox_KeyPress(object sender, KeyPressEventArgs e)
        {
            Modifier = true;
            this.Text = NomPrograme + "   *" + FichierCourant;
        }
        private void Sosa2DeEndroitTextBox_TextChanged(object sender, EventArgs e)
        {
            grille[sosaCourant * 2][DELIEU] = Sosa2DeEndroitTextBox.Text;
            if (!LongeurTextOk(Sosa2DeEndroitTextBox.Text))
            {
                Sosa2DeEndroitTextBox.BackColor = couleurTextTropLong;
            }
            else
            {
                Sosa2DeEndroitTextBox.BackColor = couleurChamp;
            }
        }
        private void Sosa2DeEndroitTextBox_KeyPress(object sender, KeyPressEventArgs e)
        {
            Modifier = true;
            this.Text = NomPrograme + "   *" + FichierCourant;
        }
        private void Sosa23MaTextBox_TextChanged(object sender, EventArgs e)
        {
            grille[sosaCourant * 2][MALE] = Sosa23MaTextBox.Text;
            bool rep = ValiderDate(Sosa23MaTextBox.Text);
            if (rep)
            {
                Sosa23MaTextBox.BackColor = Color.White;
                if (!LongeurTextOk(grille[sosaCourant * 2][MALE]))
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
        private void Sosa23MaTextBox_KeyPress(object sender, KeyPressEventArgs e)
        {
            Modifier = true;
            this.Text = NomPrograme + "   *" + FichierCourant;
        }
        private void Sosa23MaEndroitTextBox_TextChanged(object sender, EventArgs e)
        {
            grille[sosaCourant * 2][MALIEU] = Sosa23MaEndroitTextBox.Text;
            if (!LongeurTextOk(Sosa23MaEndroitTextBox.Text))
            {
                Sosa23MaEndroitTextBox.BackColor = couleurTextTropLong;
            }
            else
            {
                Sosa23MaEndroitTextBox.BackColor = couleurChamp;
            }
        }
        private void Sosa23MaEndroitBox_KeyPress(object sender, KeyPressEventArgs e)
        {
            Modifier = true;
            this.Text = NomPrograme + "   *" + FichierCourant;
        }
        private void Sosa3NomTextBox_TextChanged(object sender, EventArgs e)
        {
            grille[sosaCourant * 2 + 1][NOM] = Sosa3NomTextBox.Text;
            if (!LongeurNomtOk(grille[sosaCourant * 2 + 1][NOM]))
            {
                Sosa3NomTextBox.BackColor = couleurTextTropLong;
            }
            else
            {
                Sosa3NomTextBox.BackColor = couleurChamp;
            }
        }
        private void Sosa3NomTextBox_KeyPress(object sender, KeyPressEventArgs e)
        {
            Modifier = true;
            this.Text = NomPrograme + "   *" + FichierCourant;
        }
        private void Sosa3NeTextBox_TextChanged(object sender, EventArgs e)
        {
            grille[sosaCourant * 2 + 1][NELE] = Sosa3NeTextBox.Text;
            bool rep = ValiderDate(Sosa3NeTextBox.Text);
            if (rep)
            {
                Sosa3NeTextBox.BackColor = Color.White;
                if (!LongeurTextOk(grille[sosaCourant * 2 + 1][NELE]))
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
        private void Sosa3NeTextBox_KeyPress(object sender, KeyPressEventArgs e)
        {
            Modifier = true;
            this.Text = NomPrograme + "   *" + FichierCourant;
        }
        private void Sosa3NeEndroitTextBox_TextChanged(object sender, EventArgs e)
        {
            grille[sosaCourant * 2 + 1][NELIEU] = Sosa3NeEndroitTextBox.Text;
            if (!LongeurTextOk(Sosa3NeEndroitTextBox.Text))
            {
                Sosa3NeEndroitTextBox.BackColor = couleurTextTropLong;
            }
            else
            {
                Sosa3NeEndroitTextBox.BackColor = couleurChamp;
            }
        }
        private void Sosa3NeEndroitTextBox_KeyPress(object sender, KeyPressEventArgs e)
        {
            Modifier = true;
            this.Text = NomPrograme + "   *" + FichierCourant;
        }
        private void Sosa3DeTextBox_TextChanged(object sender, EventArgs e)
        {
            grille[sosaCourant * 2 + 1][DELE] = Sosa3DeTextBox.Text;
            bool rep = ValiderDate(Sosa3DeTextBox.Text);
            if (rep)
            {
                Sosa3DeTextBox.BackColor = Color.White;
                if (!LongeurTextOk(grille[sosaCourant * 2 + 1][DELE]))
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
        private void Sosa3DeTextBox_KeyPress(object sender, KeyPressEventArgs e)
        {
            Modifier = true;
            this.Text = NomPrograme + "   *" + FichierCourant;
        }
        private void Sosa3DeEndroitTextBox_TextChanged(object sender, EventArgs e)
        {
            grille[sosaCourant * 2 + 1][DELIEU] = Sosa3DeEndroitTextBox.Text;
            if (!LongeurTextOk(Sosa3DeEndroitTextBox.Text))
            {
                Sosa3DeEndroitTextBox.BackColor = couleurTextTropLong;
            }
            else
            {
                Sosa3DeEndroitTextBox.BackColor = couleurChamp;
            }
        }
        private void Sosa3DeEndroitTextBox_KeyPress(object sender, KeyPressEventArgs e)
        {
            Modifier = true;
            this.Text = NomPrograme + "   *" + FichierCourant;
        }
        private void Sosa4NomTextBox_TextChanged(object sender, EventArgs e)
        {
            grille[sosaCourant * 4][NOM] = Sosa4NomTextBox.Text;
            if (!LongeurNomtOk(grille[sosaCourant * 4][NOM]))
            {
                Sosa4NomTextBox.BackColor = couleurTextTropLong;
            }
            else
            {
                Sosa4NomTextBox.BackColor = couleurChamp;
            }
        }
        private void Sosa4NomTextBox_KeyPress(object sender, KeyPressEventArgs e)
        {
            Modifier = true;
            this.Text = NomPrograme + "   *" + FichierCourant;
        }
        private void Sosa4NeTextBox_TextChanged(object sender, EventArgs e)
        {
            grille[sosaCourant * 4][NELE] = Sosa4NeTextBox.Text;
            bool rep = ValiderDate(Sosa4NeTextBox.Text);
            if (rep)
            {
                Sosa4NeTextBox.BackColor = Color.White;
                if (!LongeurTextOk(grille[sosaCourant * 4][NELE]))
                {
                    Sosa4NeTextBox.BackColor = couleurTextTropLong;
                }
                else
                {
                    Sosa4NeTextBox.BackColor = couleurChamp;
                }
            }
            else
            {
                Sosa4NeTextBox.BackColor = Color.Red;
            }
        }
        private void Sosa4NeTextBox_KeyPress(object sender, KeyPressEventArgs e)
        {
            Modifier = true;
            this.Text = NomPrograme + "   *" + FichierCourant;
        }
        private void Sosa4NeEndroitTextBox_TextChanged(object sender, EventArgs e)
        {
            grille[sosaCourant * 4][NELIEU] = Sosa4NeEndroitTextBox.Text;
            if (!LongeurTextOk(Sosa4NeEndroitTextBox.Text))
            {
                Sosa4NeEndroitTextBox.BackColor = couleurTextTropLong;
            }
            else
            {
                Sosa4NeEndroitTextBox.BackColor = couleurChamp;
            }
        }
        private void Sosa4NeEndroitTextBox_KeyPress(object sender, KeyPressEventArgs e)
        {
            Modifier = true;
            this.Text = NomPrograme + "   *" + FichierCourant;
        }
        private void Sosa4DeTextBox_TextChanged(object sender, EventArgs e)
        {
            grille[sosaCourant * 4][DELE] = Sosa4DeTextBox.Text;
            bool rep = ValiderDate(Sosa4DeTextBox.Text);
            if (rep)
            {
                Sosa4DeTextBox.BackColor = Color.White;
                if (!LongeurTextOk(grille[sosaCourant * 4][DELE]))
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
        private void Sosa4DeTextBox_KeyPress(object sender, KeyPressEventArgs e)
        {
            Modifier = true;
            this.Text = NomPrograme + "   *" + FichierCourant;
        }
        private void Sosa4DeEndroitTextBox_TextChanged(object sender, EventArgs e)
        {
            grille[sosaCourant * 4][DELIEU] = Sosa4DeEndroitTextBox.Text;
            if (!LongeurTextOk(Sosa4DeEndroitTextBox.Text))
            {
                Sosa4DeEndroitTextBox.BackColor = couleurTextTropLong;
            }
            else
            {
                Sosa4DeEndroitTextBox.BackColor = couleurChamp;
            }
        }
        private void Sosa4DeEndroitTextBox_KeyPress(object sender, KeyPressEventArgs e)
        {
            Modifier = true;
            this.Text = NomPrograme + "   *" + FichierCourant;
        }
        private void Sosa45MaTextBox_TextChanged(object sender, EventArgs e)
        {
            grille[sosaCourant * 4][MALE] = Sosa45MaTextBox.Text;
            bool rep = ValiderDate(Sosa45MaTextBox.Text);
            if (rep)
            {
                Sosa45MaTextBox.BackColor = Color.White;
                if (!LongeurTextOk(grille[sosaCourant * 4][MALE]))
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
        }
        private void Sosa45MaTextBox_KeyPress(object sender, KeyPressEventArgs e)
        {
            Modifier = true;
            this.Text = NomPrograme + "   *" + FichierCourant;
        }
        private void Sosa45MaEndroitTextBox_TextChanged(object sender, EventArgs e)
        {
            grille[sosaCourant * 4][MALIEU] = Sosa45MaLEndroitTextBox.Text;
            if (!LongeurTextOk(Sosa45MaLEndroitTextBox.Text))
            {
                Sosa45MaLEndroitTextBox.BackColor = couleurTextTropLong;
            }
            else
            {
                Sosa45MaLEndroitTextBox.BackColor = couleurChamp;
            }
        }
        private void Sosa45MaEndroitNomTextBox_KeyPress(object sender, KeyPressEventArgs e)
        {
            Modifier = true;
            this.Text = NomPrograme + "   *" + FichierCourant;
        }
        private void Sosa5NomTextBox_TextChanged(object sender, EventArgs e)
        {
            grille[sosaCourant * 4 + 1][NOM] = Sosa5NomTextBox.Text;
            if (!LongeurNomtOk(grille[sosaCourant * 4 + 1][NOM]))
            {
                Sosa5NomTextBox.BackColor = couleurTextTropLong;
            }
            else
            {
                Sosa5NomTextBox.BackColor = couleurChamp;
            }
        }
        private void Sosa5NomTextBox_KeyPress(object sender, KeyPressEventArgs e)
        {
            Modifier = true;
            this.Text = NomPrograme + "   *" + FichierCourant;
        }
        private void Sosa5NeTextBox_TextChanged(object sender, EventArgs e)
        {
            grille[sosaCourant * 4 + 1][NELE] = Sosa5NeTextBox.Text;
            bool rep = ValiderDate(Sosa5NeTextBox.Text);
            if (rep)
            {
                Sosa5NeTextBox.BackColor = Color.White;
                if (!LongeurTextOk(grille[sosaCourant * 4 + 1][NELE]))
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
        private void Sosa5NeTextBox_KeyPress(object sender, KeyPressEventArgs e)
        {
            Modifier = true;
            this.Text = NomPrograme + "   *" + FichierCourant;
        }
        private void Sosa5NeEndroitTextBox_TextChanged(object sender, EventArgs e)
        {
            grille[sosaCourant * 4 + 1][NELIEU] = Sosa5NeEndroitTextBox.Text;
            if (!LongeurTextOk(Sosa5NeEndroitTextBox.Text))
            {
                Sosa5NeEndroitTextBox.BackColor = couleurTextTropLong;
            }
            else
            {
                Sosa5NeEndroitTextBox.BackColor = couleurChamp;
            }
        }
        private void Sosa5NeEndroit1NomTextBox_KeyPress(object sender, KeyPressEventArgs e)
        {
            Modifier = true;
            this.Text = NomPrograme + "   *" + FichierCourant;
        }
        private void Sosa5DeTextBox_TextChanged(object sender, EventArgs e)
        {
            {
                grille[sosaCourant * 4 + 1][DELE] = Sosa5DeTextBox.Text;
                bool rep = ValiderDate(Sosa5DeTextBox.Text);
                if (rep)
                {
                    Sosa5DeTextBox.BackColor = Color.White;
                    if (!LongeurTextOk(grille[sosaCourant * 4 + 1][DELE]))
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
        private void Sosa5DeTextBox_KeyPress(object sender, KeyPressEventArgs e)
        {
            Modifier = true;
            this.Text = NomPrograme + "   *" + FichierCourant;
        }
        private void Sosa5DeEndroitTextBox_TextChanged(object sender, EventArgs e)
        {
            grille[sosaCourant * 4 + 1][DELIEU] = Sosa5DeEndroitTextBox.Text;
            if (!LongeurTextOk(Sosa5DeEndroitTextBox.Text))
            {
                Sosa5DeEndroitTextBox.BackColor = couleurTextTropLong;
            }
            else
            {
                Sosa5DeEndroitTextBox.BackColor = couleurChamp;
            }
        }
        private void Sosa5DeEndroitTextBox_KeyPress(object sender, KeyPressEventArgs e)
        {
            Modifier = true;
            this.Text = NomPrograme + "   *" + FichierCourant;
        }
        private void Sosa6NomTextBox_TextChanged(object sender, EventArgs e)
        {
            grille[sosaCourant * 4 + 2][NOM] = Sosa6NomTextBox.Text;
            if (!LongeurNomtOk(grille[sosaCourant * 4 + 2][NOM]))
            {
                Sosa6NomTextBox.BackColor = couleurTextTropLong;
            }
            else
            {
                Sosa6NomTextBox.BackColor = couleurChamp;
            }
        }
        private void Sosa6NomTextBox_KeyPress(object sender, KeyPressEventArgs e)
        {
            Modifier = true;
            this.Text = NomPrograme + "   *" + FichierCourant;
        }
        private void Sosa6NeTextBox_TextChanged(object sender, EventArgs e)
        {
            grille[sosaCourant * 4 + 2][NELE] = Sosa6NeTextBox.Text;
            bool rep = ValiderDate(Sosa6NeTextBox.Text);
            if (rep)
            {
                Sosa6NeTextBox.BackColor = Color.White;
                if (!LongeurTextOk(grille[sosaCourant * 4 + 2][NELE]))
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
        private void Sosa6NeTextBox_KeyPress(object sender, KeyPressEventArgs e)
        {
            Modifier = true;
            this.Text = NomPrograme + "   *" + FichierCourant;
        }
        private void Sosa6NeEndroitTextBox_TextChanged(object sender, EventArgs e)
        {
            grille[sosaCourant * 4 + 2][NELIEU] = Sosa6NeEndroitTextBox.Text;
            if (!LongeurTextOk(Sosa6NeEndroitTextBox.Text))
            {
                Sosa6NeEndroitTextBox.BackColor = couleurTextTropLong;
            }
            else
            {
                Sosa6NeEndroitTextBox.BackColor = couleurChamp;
            }
        }
        private void Sosa6NeEndroitTextBox_KeyPress(object sender, KeyPressEventArgs e)
        {
            Modifier = true;
            this.Text = NomPrograme + "   *" + FichierCourant;
        }
        private void Sosa6DeTextBox_TextChanged(object sender, EventArgs e)
        {
            grille[sosaCourant * 4 + 2][DELE] = Sosa6DeTextBox.Text;
            bool rep = ValiderDate(Sosa6DeTextBox.Text);
            if (rep)
            {
                Sosa6DeTextBox.BackColor = Color.White;
                if (!LongeurTextOk(grille[sosaCourant * 4 + 2][DELE]))
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
        private void Sosa6DeTextBox_KeyPress(object sender, KeyPressEventArgs e)
        {
            Modifier = true;
            this.Text = NomPrograme + "   *" + FichierCourant;
        }
        private void Sosa6DeEndroitTextBox_TextChanged(object sender, EventArgs e)
        {
            grille[sosaCourant * 4 + 2][DELIEU] = Sosa6DeEndroitTextBox.Text;
            if (!LongeurTextOk(Sosa6DeEndroitTextBox.Text))
            {
                Sosa6DeEndroitTextBox.BackColor = couleurTextTropLong;
            }
            else
            {
                Sosa6DeEndroitTextBox.BackColor = couleurChamp;
            }
        }
        private void Sosa6DeEndroitTextBox_KeyPress(object sender, KeyPressEventArgs e)
        {
            Modifier = true;
            this.Text = NomPrograme + "   *" + FichierCourant;
        }
        private void Sosa67MaTextBox_TextChanged(object sender, EventArgs e)
        {
            grille[sosaCourant * 4 + 2][MALE] = Sosa67MaTextBox.Text;
            bool rep = ValiderDate(Sosa67MaTextBox.Text);
            if (rep)
            {
                Sosa67MaTextBox.BackColor = Color.White;
                if (!LongeurTextOk(grille[sosaCourant * 2][MALE]))
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
        private void Sosa67MaTextBox_KeyPress(object sender, KeyPressEventArgs e)
        {
            Modifier = true;
            this.Text = NomPrograme + "   *" + FichierCourant;
        }
        private void Sosa67MaEndroitTextBox_TextChanged(object sender, EventArgs e)
        {
            grille[sosaCourant * 4 + 2][MALIEU] = Sosa67MaEndroitTextBox.Text;
            if (!LongeurTextOk(Sosa67MaEndroitTextBox.Text))
            {
                Sosa67MaEndroitTextBox.BackColor = couleurTextTropLong;
            }
            else
            {
                Sosa67MaEndroitTextBox.BackColor = couleurChamp;
            }
        }
        private void Sosa67MAEndroitTextBox_KeyPress(object sender, KeyPressEventArgs e)
        {
            Modifier = true;
            this.Text = NomPrograme + "   *" + FichierCourant;
        }
        private void Sosa7NomTextBox_TextChanged(object sender, EventArgs e)
        {
            grille[sosaCourant * 4 + 3][NOM] = Sosa7NomTextBox.Text;
            if (!LongeurNomtOk(grille[sosaCourant * 4 + 3][NOM]))
            {
                Sosa7NomTextBox.BackColor = couleurTextTropLong;
            }
            else
            {
                Sosa7NomTextBox.BackColor = couleurChamp;
            }
        }
        private void Sosa7NomTextBox_KeyPress(object sender, KeyPressEventArgs e)
        {
            Modifier = true;
            this.Text = NomPrograme + "   *" + FichierCourant;
        }
        private void Sosa7NeTextBox_TextChanged(object sender, EventArgs e)
        {
            grille[sosaCourant * 4 + 3][NELE] = Sosa7NeTextBox.Text;
            bool rep = ValiderDate(Sosa7NeTextBox.Text);
            if (rep)
            {
                Sosa7NeTextBox.BackColor = Color.White;
                if (!LongeurTextOk(grille[sosaCourant * 4 + 3][NELE]))
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
        private void Sosa7NeTextBox_KeyPress(object sender, KeyPressEventArgs e)
        {
            Modifier = true;
            this.Text = NomPrograme + "   *" + FichierCourant;
        }
        private void Sosa7NeEndroitTextBox_TextChanged(object sender, EventArgs e)
        {
            grille[sosaCourant * 4 + 3][NELIEU] = Sosa7NeTextBox.Text;
            if (!LongeurTextOk(Sosa7NeEndroitTextBox.Text))
            {
                Sosa7NeEndroitTextBox.BackColor = couleurTextTropLong;
            }
            else
            {
                Sosa7NeEndroitTextBox.BackColor = couleurChamp;
            }
        }
        private void Sosa7NeEndroitTextBox_KeyPress(object sender, KeyPressEventArgs e)
        {
            Modifier = true;
            this.Text = NomPrograme + "   *" + FichierCourant;
        }
        private void Sosa7DeTextBox_TextChanged(object sender, EventArgs e)
        {
            grille[sosaCourant * 4 + 3][DELE] = Sosa7DeTextBox.Text;
            bool rep = ValiderDate(Sosa7DeTextBox.Text);
            if (rep)
            {
                Sosa7DeTextBox.BackColor = Color.White;
                if (!LongeurTextOk(grille[sosaCourant * 4 + 3][DELE]))
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
        private void Sosa7DeTextBox_KeyPress(object sender, KeyPressEventArgs e)
        {
            Modifier = true;
            this.Text = NomPrograme + "   *" + FichierCourant;
        }
        private void Sosa7DeEndroitTextBox_TextChanged(object sender, EventArgs e)
        {
            grille[sosaCourant * 4 + 3][DELIEU] = Sosa7DeEndroitTextBox.Text;
            if (!LongeurTextOk(Sosa7DeEndroitTextBox.Text))
            {
                Sosa7DeEndroitTextBox.BackColor = couleurTextTropLong;
            }
            else
            {
                Sosa7DeEndroitTextBox.BackColor = couleurChamp;
            }
        }
        private void Sosa7DeEndroitTextBox_KeyPress(object sender, KeyPressEventArgs e)
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
            string numeroTableau = DessinerPage(ref document, ref gfx, sosa, true);
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
                //str = grille[1][ NOM];
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
            // trier par nom si sosa défini
            if (ChoixSosaComboBox.Text != "")
            {
                Triage(grille, NOM);
            }
            TableMatiere(ref document, ref gfx, ref page);
            for (int f = 0; f < 512; f++)
            {
                grille[f][SOSA] = grille[f][SOSA].PadLeft(5, '0');
            }
            // trier pas SOSA
            Triage(grille, SOSA);
            //Trier(grille, SOSA, "ASC");
            for (int f = 0; f < 512; f++)
            {
                //Console.WriteLine("3614>" + grille[f, SOSA]);
                grille[f][SOSA] =    grille[f][SOSA].TrimStart('0');
                //Console.WriteLine("3617>" + grille[f, SOSA]);
            }
            foreach (int sosa in listePage)
            {
                NouvellePage(ref document, ref gfx, ref page,"L");
                DessinerPage(ref document, ref gfx, sosa, true);
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
                ChoixLV.Columns.Add("Nom", 80);
                ChoixLV.Columns.Add("Prénom", 150);
                ChoixLV.Columns.Add("Date naissance", 100);
                ChoixLV.Columns.Add("Lieu naissance", 100);
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
        private void NomRecherche_KeyDown(object sender, KeyEventArgs e)
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
            string numeroTableau = DessinerPage(ref document, ref gfx, sosa, false);
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
            grille[sosaCourant][NOTE1] = Note1.Text;
        }
        private void Note1_KeyPress(object sender, KeyPressEventArgs e)
        {
            Modifier = true;
            this.Text = NomPrograme + "   *" + FichierCourant;
        }
        private void Note2_TextChanged(object sender, EventArgs e)
        {
            grille[sosaCourant][NOTE2] = Note2.Text;
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
                ChoixSosaComboBox.Text = grille[sosa][PAGE];
                FlecheGaucheRechercheButton.Visible = false;
                FlecheDroiteRechercheButton.Visible = false;

                if (grille[rechercheListe[0]][SOSA] == grille[sosa][PAGE]) RectangleSosa1.BorderColor = Color.White;
                if (Int32.Parse(grille[sosa][SOSA]) == Int32.Parse(grille[sosa][PAGE]) * 2) RectangleSosa2.BorderColor = Color.White;
                if (Int32.Parse(grille[sosa][SOSA]) == Int32.Parse(grille[sosa][PAGE]) * 2 + 1) RectangleSosa3.BorderColor = Color.White;
                if (Int32.Parse(grille[sosa][SOSA]) == Int32.Parse(grille[sosa][PAGE]) * 4) RectangleSosa4.BorderColor = Color.White;
                if (Int32.Parse(grille[sosa][SOSA]) == Int32.Parse(grille[sosa][PAGE]) * 4 + 1) RectangleSosa5.BorderColor = Color.White;
                if (Int32.Parse(grille[sosa][SOSA]) == Int32.Parse(grille[sosa][PAGE]) * 4 + 2) RectangleSosa6.BorderColor = Color.White;
                if (Int32.Parse(grille[sosa][SOSA]) == Int32.Parse(grille[sosa][PAGE]) * 4 + 3) RectangleSosa7.BorderColor = Color.White;

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
                foreach (string[] info in grille)
                {
                    rep = true;
                    
                    foreach (string m in mots)
                    {
                        if (!info[NOM].ToLower().Contains(m.ToLower())) rep = false;
                    }
                    if (rep)
                    {
                        rechercheListe[Int32.Parse(info[SOSA])] = 1;
                        if (rechercheListe[0] == 0) rechercheListe[0] = Int32.Parse(info[SOSA]);
                        trouver = trouver + 1;
                    }
                }
                if (trouver == 0) return;

                ChoixSosaComboBox.Text = grille[rechercheListe[0]][PAGE];
                if (grille[rechercheListe[0]][SOSA] == grille[rechercheListe[0]][PAGE]) RectangleSosa1.BorderColor = Color.White;
                if (Int32.Parse(grille[rechercheListe[0]][SOSA]) == Int32.Parse(grille[rechercheListe[0]][PAGE]) * 2) RectangleSosa2.BorderColor = Color.White;
                if (Int32.Parse(grille[rechercheListe[0]][SOSA]) == Int32.Parse(grille[rechercheListe[0]][PAGE]) * 2 + 1) RectangleSosa3.BorderColor = Color.White;
                if (Int32.Parse(grille[rechercheListe[0]][SOSA]) == Int32.Parse(grille[rechercheListe[0]][PAGE]) * 4) RectangleSosa4.BorderColor = Color.White;
                if (Int32.Parse(grille[rechercheListe[0]][SOSA]) == Int32.Parse(grille[rechercheListe[0]][PAGE]) * 4 + 1) RectangleSosa5.BorderColor = Color.White;
                if (Int32.Parse(grille[rechercheListe[0]][SOSA]) == Int32.Parse(grille[rechercheListe[0]][PAGE]) * 4 + 2) RectangleSosa6.BorderColor = Color.White;
                if (Int32.Parse(grille[rechercheListe[0]][SOSA]) == Int32.Parse(grille[rechercheListe[0]][PAGE]) * 4 + 3) RectangleSosa7.BorderColor = Color.White;
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
                    ChoixSosaComboBox.Text = grille[f][PAGE];
                    if (grille[f][SOSA] == grille[f][PAGE]) RectangleSosa1.BorderColor = Color.White;
                    if (Int32.Parse(grille[f][SOSA]) == Int32.Parse(grille[f][PAGE]) * 2) RectangleSosa2.BorderColor = Color.White;
                    if (Int32.Parse(grille[f][SOSA]) == Int32.Parse(grille[f][PAGE]) * 2 + 1) RectangleSosa3.BorderColor = Color.White;
                    if (Int32.Parse(grille[f][SOSA]) == Int32.Parse(grille[f][PAGE]) * 4) RectangleSosa4.BorderColor = Color.White;
                    if (Int32.Parse(grille[f][SOSA]) == Int32.Parse(grille[f][PAGE]) * 4 + 1) RectangleSosa5.BorderColor = Color.White;
                    if (Int32.Parse(grille[f][SOSA]) == Int32.Parse(grille[f][PAGE]) * 4 + 2) RectangleSosa6.BorderColor = Color.White;
                    if (Int32.Parse(grille[f][SOSA]) == Int32.Parse(grille[f][PAGE]) * 4 + 3) RectangleSosa7.BorderColor = Color.White;
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
                    ChoixSosaComboBox.Text = grille[f][PAGE];
                    if (grille[f][SOSA] == grille[f][PAGE]) RectangleSosa1.BorderColor = Color.White;
                    if (Int32.Parse(grille[f][SOSA]) == Int32.Parse(grille[f][PAGE]) * 2) RectangleSosa2.BorderColor = Color.White;
                    if (Int32.Parse(grille[f][SOSA]) == Int32.Parse(grille[f][PAGE]) * 2 + 1) RectangleSosa3.BorderColor = Color.White;
                    if (Int32.Parse(grille[f][SOSA]) == Int32.Parse(grille[f][PAGE]) * 4) RectangleSosa4.BorderColor = Color.White;
                    if (Int32.Parse(grille[f][SOSA]) == Int32.Parse(grille[f][PAGE]) * 4 + 1) RectangleSosa5.BorderColor = Color.White;
                    if (Int32.Parse(grille[f][SOSA]) == Int32.Parse(grille[f][PAGE]) * 4 + 2) RectangleSosa6.BorderColor = Color.White;
                    if (Int32.Parse(grille[f][SOSA]) == Int32.Parse(grille[f][PAGE]) * 4 + 3) RectangleSosa7.BorderColor = Color.White;
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
            str = grille[1][NOM];
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
            // gfx.DrawString("Date et lieu de naissance", font8, XBrushes.Black, col1 + padding, Y + hauteurLigne * 2);
            // col 1 ligne 3 Date et lieu du décès
            //gfx.DrawString("Date et lieu du décès", font8, XBrushes.Black, col1 + padding, Y + hauteurLigne * 3);
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
                gfx.DrawString(grille[sosa][NOM], fontNom, XBrushes.Black, col1 + padding, Y + hauteurLigne);
                // col 1 ligne 2
                if (grille[sosa][NELE] != "" || grille[sosa][NELIEU] != "")
                    gfx.DrawString("°", fontDate, XBrushes.Black, col1 + padding + 1, Y + hauteurLigne * 2);
                PDFEcrire(ref gfx, grille[sosa][NELE], col1 + padding + 6, Y + hauteurLigne * 2, .5 * pouce);
                PDFEcrire(ref gfx, grille[sosa][NELIEU], col2, Y + hauteurLigne * 2, 1 * pouce);
                // col 1 ligne3
                if (grille[sosa][DELE] != "" || grille[sosa][DELIEU] != "")
                    gfx.DrawString("+", fontDate, XBrushes.Black, col1 + padding, Y + hauteurLigne * 3);
                PDFEcrire(ref gfx, grille[sosa][DELE], col1 + padding + 6, Y + hauteurLigne * 3, .5 * pouce);
                PDFEcrire(ref gfx, grille[sosa][DELIEU], col2, Y + hauteurLigne * 3, 1.5 * pouce);

                if (sosa > 1)
                {
                    // col 3 ligne 1
                    if (grille[sosa][MALE] != "")
                    {
                        gfx.DrawString("X", fontDate, XBrushes.Black, col3 + .83 * pouce, Y + hauteurLigne);
                        gfx.DrawString(grille[sosa][MALE], fontDate, XBrushes.Black, col3 + 6 + .83 * pouce, Y + hauteurLigne);
                    }
                    // col 3 ligne 2
                    PDFEcrireCentrer(ref gfx, grille[sosa][MALIEU], col3, Y + hauteurLigne * 2, col4);

                }
                if (sosa > 1)
                {
                    // col 4 ligne 1 // nom conjoint
                    PDFEcrire(ref gfx, grille[sosa + 1][NOM], col4 + padding, Y + hauteurLigne, 2.3 * pouce);
                    // col 4 ligne 2 ET 3
                    int sosaParent = (sosa + 1) * 2;
                    if ((sosaParent < 512 && sosaParent > 0) && grille[sosaParent][NOM] != "" && grille[(sosaParent + 1)][NOM] != "")
                    {
                        nomParent = grille[sosaParent][NOM] + " et " + grille[(sosaParent + 1)][NOM];
                        PDFEcrire(ref gfx, nomParent, col4 + padding, Y + hauteurLigne * 2, 2.3 * pouce);
                        if (grille[sosaParent][MALE] != "" || grille[sosaParent][MALIEU] != "")
                        {
                            gfx.DrawString("X", fontDate, XBrushes.Black, col4 + padding + 1, Y + hauteurLigne * 3);
                            PDFEcrire(ref gfx, grille[sosaParent][MALE], col4 + padding + 6, Y + hauteurLigne * 3, .5 * pouce);
                            PDFEcrire(ref gfx, grille[sosaParent][MALIEU], col5, Y + hauteurLigne * 3, 1 * pouce);
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
            str = grille[1][NOM];
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
                gfx.DrawString(grille[sosa][NOM], fontNom, XBrushes.Black, col1 + padding, Y + hauteurLigne);
                // col 1 ligne 2
                if (grille[sosa][NELE] != "" || grille[sosa][NELIEU] != "")
                    gfx.DrawString("°", fontDate, XBrushes.Black, col1 + padding + 1, Y + hauteurLigne * 2);
                PDFEcrire(ref gfx, grille[sosa][NELE], col1 + padding + 6, Y + hauteurLigne * 2, .5 * pouce);
                PDFEcrire(ref gfx, grille[sosa][NELIEU], col2, Y + hauteurLigne * 2, 1 * pouce);
                // col 1 ligne3
                if (grille[sosa][DELE] != "" || grille[sosa][DELIEU] != "")
                    gfx.DrawString("+", fontDate, XBrushes.Black, col1 + padding, Y + hauteurLigne * 3);
                PDFEcrire(ref gfx, grille[sosa][DELE], col1 + padding + 6, Y + hauteurLigne * 3, .5 * pouce);
                PDFEcrire(ref gfx, grille[sosa][DELIEU], col2, Y + hauteurLigne * 3, 1.5 * pouce);
                
                // col 3 ligne 1
                if (sosa > 1)
                {
                    if (grille[sosa-1][MALE] != "")
                    {
                        gfx.DrawString("X", fontDate, XBrushes.Black, col3 + .83 * pouce, Y + hauteurLigne);
                        gfx.DrawString(grille[sosa-1][MALE], fontDate, XBrushes.Black, col3 + 6 + .83 * pouce, Y + hauteurLigne);
                    }
                    // col 3 ligne 2
                    PDFEcrireCentrer(ref gfx, grille[sosa-1][MALIEU], col3, Y + hauteurLigne * 2, col4);
                }
                if (sosa > 1)
                {
                // col 4 ligne 1 // nom conjoint
                    PDFEcrire(ref gfx, grille[sosa - 1][NOM], col4 + padding, Y + hauteurLigne, 2.3 * pouce);
                // col 4 ligne 2 ET 3

                    int sosaParent = (sosa - 1) * 2;
                    if ((sosaParent < 512 && sosaParent > 0) && grille[sosaParent][NOM] != "" && grille[(sosaParent + 1)][NOM] != "")
                    {
                        nomParent = grille[sosaParent][NOM] + " et " + grille[(sosaParent + 1)][NOM];
                        PDFEcrire(ref gfx, nomParent, col4 + padding, Y + hauteurLigne * 2, 2.3 * pouce);
                        if (grille[sosaParent][MALE] != "" || grille[sosaParent][MALIEU] != "") {
                            gfx.DrawString("X", fontDate, XBrushes.Black, col4 + padding + 1, Y + hauteurLigne * 3);
                            PDFEcrire(ref gfx, grille[sosaParent][MALE], col4 + padding + 6, Y + hauteurLigne * 3, .5 * pouce);
                            PDFEcrire(ref gfx, grille[sosaParent][MALIEU], col5, Y + hauteurLigne * 3, 1 * pouce);
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
 }

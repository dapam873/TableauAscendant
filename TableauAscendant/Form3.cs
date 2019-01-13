using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;


using System.IO;
using System.Diagnostics;

//using System.ComponentModel;






using System.Runtime.CompilerServices;


namespace TableauAscendant
{
    public partial class AideFm : Form
    {
        static void ZXCV(string message, [CallerLineNumber] int lineNumber = 0, [CallerMemberName] string caller = null)
        {
            Console.WriteLine(lineNumber + " " + caller + " " + message);

            string Fichier = "01log.txt";
            


            using (StreamWriter ligne = File.AppendText(Fichier))
            {
                //ligne.WriteLine(lineNumber + " " + caller + " " + message);
                ligne.WriteLine( message);
            }


        }
        public AideFm()
        {
            InitializeComponent();
        }

        private void RichTextBox1_TextChanged(object sender, EventArgs e) 
        {
            
        }

        private void AideFm_Load(object sender, EventArgs e)

        {
             if (File.Exists("01log.txt"))
            {
                File.Delete("01log.txt");
            }
            int col1 = 25;
            int col2 = 45;
            int col3 = 65;
            int col4 = 340;
            int y = 10;
            int hauteur;
            titre_lb.Location = new Point(288, y);
            y = y + 30;
            Intro_lb.Location = new Point(col1, y);
            hauteur = 60;
            Intro_lb.Height = hauteur;
            Intro_lb.Text = "Ce logiciel permet de produire un ou des fichiers PDF de TableauAscendance. Les Tableaux peuvent être créés en PDF vierge aussi bien que remplis. Les données entrées peuvent être enregistrées. Chaque rectangle représente la fiche d'une personne.";
            y = y + hauteur  + 20;
            lineShape1.X1 = col1;
            lineShape1.X2 = 790;
            lineShape1.Y1 = y;
            lineShape1.Y2 = y;
            y = y + 20;
            groupeMenu_lb.Location = new Point(col1, y);
            y = y + 20;
            menuFichier_lb.Location = new Point(col2, y);
            y = y + 20;
            menuFichierNouveau_lb.Location = new Point(col3, y);
            menuFichierNouveauDesc_lb.Location = new Point(col4, y);
            y = y + 20;
            menuFichierOuvrir_lb.Location = new Point(col3, y);
            menuFichierOuvrirDesc_lb.Location = new Point(col4, y);
            y = y + 20;
            menuFichierGEDCOM_lb.Location = new Point(col3, y);
            menuFichierGEDCOMDesc_lb.Location = new Point(col4, y);
            y = y + 30;
            fenetre1GEDCOMlb.Location = new Point(col3 + 20, y);
            fenetre1GEDCOMlb.Text = "Dans cette fenêtre, entrez le nom et/ ou le prénom du probant ( au complet ou partiel) à rechercher. Choisir dans la liste, la personne qui sera le probant. C'est-à-dire le SOSA 1";
            y = y + 60;
            recherchePb.Location = new Point(col3 + 20, y);
            y = y + 320;
            menuFichierEnregister_lb.Location = new Point(col3, y);
            menuFichierEnregisterDesc_lb.Location = new Point(col4, y);
            y = y + 20;
            menuFichierEnregisterSous_lb.Location = new Point(col3, y);
            menuFichierEnregisterSousDesc_lb.Location = new Point(col4, y);
            y = y + 20;
            menuFichierQuiter_lb.Location = new Point(col3, y);
            menuFichierQuiterDesc_lb.Location = new Point(col4, y);
            y = y + 20;
            menuEditer_lb.Location = new Point(col2, y);
            y = y + 20;
            menuEffacerTout_lb.Location = new Point(col3, y);
            menuEffacerToutDesc_lb.Location = new Point(col4, y);
            y = y + 20;
            menuRapport_lb.Location = new Point(col2, y);
            y = y + 20;
            menuRapport4Generation_lb.Location = new Point(col3, y);
            menuRapport4GenerationDesc_lb.Location = new Point(col4, y);
            y = y + 20;
            menuRapportCourant_lb.Location = new Point(col3, y);
            y = y + 20;
            menuRapportToutes_lb.Location = new Point(col3, y);
            hauteur = 40;
            menuRapportToutesDesc_lb.Height = hauteur;
            menuRapportToutesDesc_lb.Location = new Point(col4, y);
            menuRapportToutesDesc_lb.Text = "Produit un rapport de tous les tableaux avec la page couverture et la table des matières.";
            y = y + hauteur;
            menuRapportPatri_lb.Location = new Point(col3, y);
            menuRapportPatriDesc_lb.Height = hauteur;
            menuRapportPatriDesc_lb.Location = new Point(col4, y);
            menuRapportPatriDesc_lb.Text = "Produit un rapport de titre d'ascendance patrilinéaire de l'individu qui à le sosa 1 et cela sur neuf génération.";
            y = y + hauteur;
            menuRapportMatri_lb.Location = new Point(col3, y);
            menuRapportMatriDesc_lb.Height = hauteur;
            menuRapportMatriDesc_lb.Location = new Point(col4, y);
            menuRapportMatriDesc_lb.Text = "Produit un rapport de titre d'ascendance matrilinéaire de l'individu qui à le sosa 1 et cela sur neuf génération.";
            y = y + hauteur + 20;
            menuParametre_lb.Location = new Point(col2, y);
            y = y + 20;
            menuParametreDossierPDF_lb.Location = new Point(col3, y);
            menuParametreDossierPDFDesc_lb.Location = new Point(col4, y);
            y = y + 20;
            menuParametreCouleurPDF_lb.Location = new Point(col3, y);
            menuParametreCouleurPDFDesc_lb.Location = new Point(col4, y);

            y = y + 80;
            lineShape2.X1 = col1;
            lineShape2.X2 = 790;
            lineShape2.Y1 = y;
            lineShape2.Y2 = y;
            // Goupe champ
            y = y + 30;
            fichePb.Location = new Point(col1, y);
            // SOSA
            y = y + 180;
            groupSosa_lb.Location = new Point(col2, y);
            y = y + 20;
            textSosa_lb.Location = new Point(col3, y);
            hauteur = 160;
            textSosa_lb.Height = hauteur;
            textSosa_lb.Text = "Dans la première fiche à l'extrême droite, dans le champ gauche de la première ligne, on tape ou on sélectionne le numéro SOSA, le probant étant 1.\n\nSeuls les SOSA disponibles pour cette position sont disponibles dans la liste. Toute erreur dans ce champ produira un fond rouge dans ce champ.\n\nPar exemple, si on tape 2, on aura une erreur.\n\nLe SOSA des autres fiches sera automatiquement mise à jour.";

            // Nom
            y = y + hauteur + 20;
            groupeNomlb.Location = new Point(col2, y);
            y = y + 20;
            textNom_lb.Location = new Point(col3, y);
            hauteur = 40;
            textNom_lb.Height = hauteur;
            textNom_lb.Text = "Si la couleur de l'arrière-plan du champ de nom est jaune, c'est que le texte est trop long. Cela n'empêchera pas de continuer et de faire un rapport. Si c'est jaune, il sera tronqué sur la droite.";
            // Date
            y = y + hauteur + 20;
            groupeDatelb.Location = new Point(col2, y);
            y = y + 20;
            textDate_lb.Location = new Point(col3, y);
            hauteur = 90;
            textDate_lb.Height = hauteur;
            textDate_lb.Text = "Si la couleur de l'arrière-plan du champ date est rouge, c'est que le format de la date n'est pas valide. Format valide AAAA, AAAAA-MM et AAAA-MM-JJ. Si par contre le texte contient l'un des mots «après» «avant» «entre» «et» ou «vers», il sera considéré valide. Si la couleur de l'arrière-plan du champ de date est jaune, c'est que le texte est trop long. Que l'arrière-plan soit rouge ou jaune, cela n'empêchera pas de continuer et de faire un rapport. Si c'est jaune, il sera tronqué sur la gauche.";

            //lieu
            y = y + hauteur + 20;
            groupLieu_lb.Location = new Point(col2, y);
            y = y + 20;
            hauteur = 140;
            textLieu_lb.Height = hauteur;
            textLieu_lb.Text = "Si la couleur de l'arrière-plan du champ de lieu est jaune, c'est que le texte est trop long. Cela n'empêchera pas de continuer et de faire un rapport, mais il sera tronqué sur la gauche.\n\nSi vos données sont entrées à partir d'un fichier GEDCOM généré par votre logiciel de généalogie, vous pouvez, comme le fait le logiciel de généalogie GRAMP, divisez votre lieu en partie, soit, l'adresse, ville, province/état, pays et TableauAscence utilisera ville comme partie seulement. C'est-à-dire que si vous utilisez comme partie, village, commune, municipalité ou autre, ils ne seront pas utilisés et l'adresse complète sera utilisée. Ceci est dû au fait que la norme GEDCOM utilise le mot-clé «CITY» dans son fichier pour enregistrer cette portion de lieu.";
            textLieu_lb.Location = new Point(col3, y);
            //Rapport
            y = y + hauteur + 20;
            groupRapport_lb.Location = new Point(col2, y);
            y = y + 20;
            hauteur = 280;
            textRapport_lb.Text = "Si on indique pour le champ SOSA le numéro 0 ou on le laisse vide, le rapport sortira vierge dans le PDF.\nPour sortir les rapports:\n\n    Première étape, choisir le dossier ou sera enregistré les fichiers PDF.\n        Dans le menu en haut de la fenêtre, ->Paramètre->Dossier PDF et suivre la procédure normale de\n        Windows pour choisir le dossier.\n\n    Deuxième étape, choisir le type de rapport.\n        Une page\n            Dans le menu en haut de la fenêtre, -> PDF->Créer page courante.\n                Ceci enregistrera, puis affichera le fichier PDF.\n                Le fichier portera pour nom le numéro de la page.\n        Toutes les pages\n            Dans le menu en haut de la fenêtre, -> PDF->Créer toutes les pages.\n                Ceci enregistrera, puis affichera le fichier PDF.\n                Le fichier portera pour nom TableauAscendant.\n";
            textRapport_lb.Location = new Point(col3, y);

            //Mariage
            y = y + hauteur + 20;
            groupMariage_lb.Location = new Point(col2, y);
            y = y + 20;
            hauteur =150;
            textMariage_lb.Text = "Pour le SOSA 1, on doit entrer pour le mariage, le lieu, la date et la personne avec qui on est marié\n\nPour les SOSA où le numéro est pair, on doit entrer le lieu et la date. La personne avec qui on est marié sera affiché lorsque que l'on entrera le nom dans la fiche du SOSA suivant.\n\nPour les SOSA où le numéro est impair, aucune donnée n'est entrée, par contre elles seront affichées si les données sont entrées dans la fiche du SOSA précédent.";
            textMariage_lb.Location = new Point(col3, y);
            textMariage_lb.Height = hauteur;

            // Note

            y = y + hauteur + 30;
            lineShape3.X1 = col1;
            lineShape3.X2 = 790;
            lineShape3.Y1 = y;
            lineShape3.Y2 = y;
          
            y = y + 30;
            groupNote_lb.Location = new Point(col2, y);
            y = y + 20;
            hauteur = 160;
            textNote_lb.Text = "Deux sections dont disponibles pour écrire des notes sur chaque page. Elle sont placées au-dessus et en dessous de la boite du premier sosa.\r\n\nSi dans chaque section de note, le texte ne parait pas au complet, il faudra raccourcir le texte ou le partager entre les deux sections.\r\n\nSi le texte concerne une des personnes de cette page, commencez la phrase par  «Sosa x », Le x représente le numéro du sosa de la personne.";
            textNote_lb.Location = new Point(col3, y);
            textNote_lb.Height = hauteur;

            y = y + hauteur + 20;
            lineShape4.X1 = col1;
            lineShape4.X2 = 790;
            lineShape4.Y1 = y;
            lineShape4.Y2 = y;

            //recherche
            y = y + 20;
            groupRecherche_lb.Location = new Point(col2, y);
            y = y + 20;
            hauteur = 205;
            textRecherche_lb.Text = "Pour faire une recherche dans les différents tableaux, on inscrit en bas, à gauche du tableau, le numéro sosa ou le nom d\'une personne à rechercher.\r\n\nLe numéro sosa doit être un nombre entre 1 à 511.\r\n\n On peut faire une recherche avec un prénom ou un nom et aussi avec les deux. Chaque mot entré peut être une partie du prénom ou d\'un nom et pas nécessairement au début du prénom ou du nom. c\'est à dire si je cherche avec ces mots pet et ri, le résultat de la recherche comprendra Mario Montpetit et Richard Petrit s\'il sont les seuls à être dans les tableaux.\r\n\n Si la recherche produit plusieurs noms, les flèches pour aller de l\'un à l\'autre.";
            textRecherche_lb.Location = new Point(col3, y);
            textRecherche_lb.Height = hauteur;
        }
    }
}

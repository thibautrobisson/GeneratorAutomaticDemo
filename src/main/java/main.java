import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import javax.swing.*;
import java.io.*;
import java.nio.file.Files;
import java.nio.file.Paths;
import static java.nio.file.StandardCopyOption.REPLACE_EXISTING;

public class main
{
    //Public variables
    public static File repertoireFichierSources;
    public static File repertoireFichierTarget;
    public static File fichierSourceDureesCyclesCible;
    public static File fichierSourceDureesCyclesReference;
    public static File fichierSourceDifferencesTables;
    public static String moisSolde;
    public static String versionCible;
    public static String versionDeReference;


    public static void main(String[] args) throws IOException
    {
        //Voir avec Benjamin où le récupérer sur amboise et les mettre en argument au moment du déploiement
        fichierSourceDureesCyclesCible = new File("Durées cycles juillet 2019 LVS_07.20.00.d.r02 VS LVS_v07.19.02.r.r01 DEMO.xlsx");
        fichierSourceDureesCyclesReference = new File("Durées cycles juin 2019 LVS_07.19.00.d.r01 VS LVS_v07.19.01.r.r01 DEMO.xlsx");
        fichierSourceDifferencesTables = new File("Différences tables Octobre 2019 LVS_07.19.00.d.r01 VS LVS_v07.20.00.d.r02 DEMO.xlsx");
        moisSolde = "Mois Inconnu";
        versionCible = "07.20.00.d.r02 + 07.19.02.r.r01";
        versionDeReference = "07.19.00.d.r.01 + 01.19.00.r.r01";
        repertoireFichierSources = new File("C:\\Users\\thiba\\IdeaProjects\\Test");
        System.out.println(Utils.ANSI_PURPLE + "Version Cible : " + Utils.ANSI_CYAN + versionCible + Utils.ANSI_RESET);
        System.out.println(Utils.ANSI_PURPLE + "Version de Référence : " + Utils.ANSI_CYAN + versionDeReference + Utils.ANSI_RESET);
        System.out.println(Utils.ANSI_PURPLE + "Chemin fichier Durées cycle version de référence : " + Utils.ANSI_CYAN +  fichierSourceDureesCyclesReference + Utils.ANSI_RESET);
        System.out.println(Utils.ANSI_PURPLE + "Chemin fichier Durées cycle version cible : " +  Utils.ANSI_CYAN + fichierSourceDureesCyclesCible + Utils.ANSI_RESET);
        System.out.println(Utils.ANSI_PURPLE + "Répertoire de base des sources: " +  Utils.ANSI_CYAN + repertoireFichierSources + Utils.ANSI_RESET);

        // ------------------------------------------------------------------------------------------------------------------------------------
        //Afficher une box DialogDirectory demandant à l'utilisateur où enregistrer le poerpoint/images/excel
        repertoireFichierTarget = Utils.OpenBrowserDialogTargetPpt();

        if (repertoireFichierTarget == null)
        {
            JFrame f = new JFrame("frame");
            JOptionPane.showMessageDialog(f,
                    "Le chemin spécifié pour le fichier CR ppt est incorrect...Consultez les logs.",
                    "Erreur",
                    JOptionPane.ERROR_MESSAGE);
        }



        // ------------------------------------------------------------------------------------------------------------------------------------
        //On créer un répertoire dans lequel se trouvera tous nos fichiers de sorties (CR ppt, fichier excel généré et images)
        //ainsi que tous nos fichiers d'entrées (Durées cycle) que l'on va copier coller dans ce même répertoire
        //(Sachant que l'on déposera également le CR ppt dans le repertoire de l'intégration où sont stockés tous leurs CR)
        repertoireFichierTarget = new File(repertoireFichierTarget + "\\CR DEMO " + versionCible);
        System.out.println("Répertoire de travail : " + repertoireFichierTarget);

        try
        {
            repertoireFichierTarget.mkdirs(); //permets de créer le dossier
        }catch (Exception e)
        {
            JFrame f = new JFrame("frame");
            JOptionPane.showMessageDialog(f,
                    "Impossible de créer le dossier : " + repertoireFichierTarget.getPath(),
                    "Erreur",
                    JOptionPane.ERROR_MESSAGE);
        }

        // ------------------------------------------------------------------------------------------------------------------------------------
        //On copie colle ensuite nos fichiers durées cycle de la version de référence et de la version cible ainsi que le fichier DifferencesTables dans le repertoire target
        try
        {
            File fichierSourceDureesCyclesReferenceCopy = new File(repertoireFichierTarget + "\\" + fichierSourceDureesCyclesReference.getName());
            Files.copy(fichierSourceDureesCyclesReference.toPath(), Paths.get(fichierSourceDureesCyclesReferenceCopy.getPath()),REPLACE_EXISTING);
            //Et on indique le nouveau chemin du fichierSourceDureesCyclesReference
            fichierSourceDureesCyclesReference = new File(repertoireFichierTarget + "\\" + fichierSourceDureesCyclesReference.getName());
            System.out.println("Copie du fichier Durées Cycles de la version de référence effectuée...");

            File fichierSourceDureesCyclesCibleCopy = new File(repertoireFichierTarget + "\\" +  fichierSourceDureesCyclesCible.getName());
            Files.copy(fichierSourceDureesCyclesCible.toPath(), Paths.get(fichierSourceDureesCyclesCibleCopy.getPath()),REPLACE_EXISTING);
            //Et on indique le nouveau chemin du fichierSourceDureesCyclesReference
            fichierSourceDureesCyclesCible = new File(repertoireFichierTarget + "\\" + fichierSourceDureesCyclesCible.getName());
            System.out.println("Copie du fichier Durées Cycles de la version de cible effectuée...");

            File fichierSourceDifferencesTablesCopy = new File(repertoireFichierTarget + "\\" +  fichierSourceDifferencesTables.getName());
            Files.copy(fichierSourceDifferencesTables.toPath(), Paths.get(fichierSourceDifferencesTablesCopy.getPath()),REPLACE_EXISTING);
            //Et on indique le nouveau chemin du fichierSourceDifferencesTables
            fichierSourceDifferencesTables = new File(repertoireFichierTarget + "\\" + fichierSourceDifferencesTables.getName());

            System.out.println("Copie du fichier Différences Tables effectuée...");
        }catch (IOException e)
        {
            e.printStackTrace();
        }

        // ------------------------------------------------------------------------------------------------------------------------------------
        //On récupère ensuite le mois de solde de la version cible grâce au nom du fichier Durées Cycle qui est du type:
        //"Durées cycles juin 2019 LVS_07.19.00.d.r01 VS LVS_v07.19.01.r.r01 DEMO.xlsx"
        moisSolde = Utils.getMoisSoldeVersionCible(fichierSourceDureesCyclesCible.getName());
        System.out.println(Utils.ANSI_PURPLE + "Mois Solde : " + Utils.ANSI_CYAN +  moisSolde + Utils.ANSI_RESET);


        // ------------------------------------------------------------------------------------------------------------------------------------
        //Récupération des données et création de tous les graphiques via un fichier excel
        //puis convertion de tous les graphiques en image png
        FillXlsx.CreateGraphAndTabSyntheseXslx();

        // ------------------------------------------------------------------------------------------------------------------------------------
        //Création / Remplissage / Sauvegarde du Compte-rendu Powerpoint
        FillPpt.CreateFillSaveCrPowerpoint(versionCible, moisSolde, repertoireFichierTarget.getPath());
        System.out.println(Utils.ANSI_GREEN + "^_^ === Succès de l'opération === ^_^" + Utils.ANSI_RESET);
    }

}

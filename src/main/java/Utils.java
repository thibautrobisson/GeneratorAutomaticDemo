import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import javax.swing.*;
import javax.swing.filechooser.FileSystemView;
import java.io.File;
import java.io.FileInputStream;
import java.io.FilenameFilter;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

public class Utils {

    public static final String ANSI_RESET = "\u001B[0m";
    public static final String ANSI_BLACK = "\u001B[30m";
    public static final String ANSI_RED = "\u001B[31m";
    public static final String ANSI_GREEN = "\u001B[32m";
    public static final String ANSI_YELLOW = "\u001B[33m";
    public static final String ANSI_BLUE = "\u001B[34m";
    public static final String ANSI_PURPLE = "\u001B[35m";
    public static final String ANSI_CYAN = "\u001B[36m";
    public static final String ANSI_WHITE = "\u001B[37m";

    private static FilenameFilter xlsxFileFilter = new FilenameFilter() {
        public boolean accept(File dir, String name) {
            return name.endsWith(".xlsx");
        }
    };

    //Méthode permettant de récupérer la liste des fichiers xlsx dans un répoertoire
    public static List<String> ListerFichiersDansRepertoire(File nomDuRepertoire)
    {
        //On récupère la liste des fichiers xlsx contenus dans ce répertoire
        File[] files = nomDuRepertoire.listFiles(xlsxFileFilter);
        //On initialise notre liste de fichiers
        List<String> listeDesFichiers = new ArrayList<String>();

        for (File fichier:files)
        {
            listeDesFichiers.add(fichier.getName());
        }
        return listeDesFichiers;
    }


    public static boolean VerifierPresenceFichierDansListeFichiers(List<String> listeFichiers, String nomFichierRecherche)
    {
        for (String fichier:listeFichiers )
        {
            if (fichier.contains(nomFichierRecherche))
            {
                return true;
            }
        }
        return false;
    }


    private static File GetFichierDansListeFichiers(File nomDuRepertoire, String nomFichierRecherche)
    {
        File[] files = nomDuRepertoire.listFiles(xlsxFileFilter);
        for (File file:files )
        {
            if (file.getName().contains(nomFichierRecherche))
            {
                return file;
            }
        }
        return null;
    }



    //Méthode plus utilisée.....
    private static File OpenBrowserDialogSourceFiles(File repertoireFichierSources, File fichierSourceSynthese, File fichierSourceDifferencesTables, File fichierSourceDureesCyclesCible)
    {
        JFileChooser jfc = new JFileChooser(FileSystemView.getFileSystemView().getHomeDirectory());
        jfc.setDialogTitle("Indiquez le dossier dans lequel se trouvent les fichiers sources excel");
        jfc.setFileSelectionMode(JFileChooser.DIRECTORIES_ONLY);

        int returnValue = jfc.showSaveDialog(null);
        if (returnValue == JFileChooser.APPROVE_OPTION) {
            if (jfc.getSelectedFile().isDirectory())
            {
                //On récupère le répertoire des fichiers sources indiqué par l'utilisateur
                repertoireFichierSources = jfc.getSelectedFile();
                System.out.println("Les fichiers sources se trouvent dans le dossier : " + repertoireFichierSources);
                System.out.println("Vérification de la présence des fichiers sources...");
                List<String> listeFichierRepertoireSource = ListerFichiersDansRepertoire(repertoireFichierSources);

                if (!VerifierPresenceFichierDansListeFichiers(listeFichierRepertoireSource, "Synthèse"))
                {
                    System.out.println("Il manque le fichier source Synthèse dans le répertoire indiqué...");
                    return null;
                }
                //On renseigne le nom du fichier Synthese
                fichierSourceSynthese = GetFichierDansListeFichiers(repertoireFichierSources, "Synthèse");

                if (!VerifierPresenceFichierDansListeFichiers(listeFichierRepertoireSource, "Différences tables"))
                {
                    System.out.println("Il manque le fichier source Différences tables dans le répertoire indiqué...");
                    return null;
                }
                //On renseigne le nom du fichier DifferencesTables
                fichierSourceDifferencesTables = GetFichierDansListeFichiers(repertoireFichierSources, "Différences tables");

                if (!VerifierPresenceFichierDansListeFichiers(listeFichierRepertoireSource, "Durées cycles"))
                {
                    System.out.println("Il manque le fichier source Durées cycles dans le répertoire indiqué...");
                    return null;
                }
                //On renseigne le nom du fichier DureesCycles
                fichierSourceDureesCyclesCible = GetFichierDansListeFichiers(repertoireFichierSources, "Durées cycles");

                return repertoireFichierSources;
            }
            else
            {
                System.out.println("Le répertoire en question n'est pas un dossier...");
            }
        }
        return null;
    }


    public static String getMoisSoldeVersionCible(String fichierDureesCycleVersionCible)
    {
        //Pour rappel le nom du fichier est de ce type:
        // =>  Durées cycles juillet 2019 LVS_07.20.00.d.r02 VS LVS_v07.19.02.r.r01 DEMO.xlsx <=
        //le mois corresponds au troisieme mot du nom du fichier
        //l'année corresponds au quatrieme du nom du fichier
        int pos = fichierDureesCycleVersionCible.indexOf(" ");
        int pos1 = fichierDureesCycleVersionCible.indexOf(" ", pos+1);
        int pos2 = fichierDureesCycleVersionCible.indexOf(" ", pos1+1);
        int pos3 = fichierDureesCycleVersionCible.indexOf(" ", pos2+1);
        String mois = fichierDureesCycleVersionCible.substring(pos1+1, pos2);
        String annee = fichierDureesCycleVersionCible.substring(pos2+1, pos3);
        return mois + " " + annee;
    }


    public static File OpenBrowserDialogTargetPpt()
    {
        JFileChooser jfc = new JFileChooser(FileSystemView.getFileSystemView().getHomeDirectory());
        jfc.setDialogTitle("Choisissez le dossier dans lequel sera crée le CR PPT ");
        jfc.setFileSelectionMode(JFileChooser.DIRECTORIES_ONLY);

        int returnValue = jfc.showSaveDialog(null);
        if (returnValue == JFileChooser.APPROVE_OPTION) {
            if (jfc.getSelectedFile().isDirectory())
            {
                return jfc.getSelectedFile();
            }
            else
            {
                System.out.println("Le répertoire en question n'est pas un dossier...");
            }
        }
        return null;
    }


    // -------------------------------------------------------
    //Méthode plus utilisée qui va permettre de récupérer dans le fichier source excel Synthèse tout ce qu'il y a d'important ! PLUS UTILISEE car toutes les infos
    // se trouvent dans le fichier DuréesCYcle
    // -------------------------------------------------------
    private static void GetAllMainvaluesFromFichierSynthese(File fichierSourceSynthese, String moisSolde, String versionCible, String versionDeReference) throws IOException {
        //On récupère la version cible via le fichier source excel Synthèse
        //On commence par récupérer le mois de solde via le titre du fichier qui est du type Synthèse Octobre 2019 DEMO
        Pattern pattern = Pattern.compile("\\s([A-Za-z]+)");
        Matcher matcher = pattern.matcher(fichierSourceSynthese.getName());
        if (matcher.find()) {
            moisSolde = matcher.group(1);
        }


        //Ensuite on va entrer dans le fichier source excel synthèse pour récupérer dans un premier temps
        //la version cible complète et la version de référence complète qui se trouve sur la sheet 0 respectivement
        // en E4  et C4
        FileInputStream inputStream = new FileInputStream(fichierSourceSynthese);
        org.apache.poi.ss.usermodel.Workbook workbook = new XSSFWorkbook(inputStream);
        Sheet sheet = workbook.getSheetAt(0);
        // --> Version cible en C4
        versionCible = UtilsXlsx.getValueStringOfCellAdress(sheet,"E4");
        // --> Version de référence en C4
        versionDeReference = UtilsXlsx.getValueStringOfCellAdress(sheet,"C4");

        inputStream.close();
    }
}

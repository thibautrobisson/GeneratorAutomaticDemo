import org.apache.commons.io.IOUtils;
import org.apache.poi.sl.usermodel.PictureData;
import org.apache.poi.xslf.usermodel.XMLSlideShow;
import org.apache.poi.xslf.usermodel.XSLFPictureShape;
import org.apache.poi.xslf.usermodel.XSLFSlide;

import java.awt.*;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Paths;

import static java.nio.file.StandardCopyOption.REPLACE_EXISTING;

public class FillPpt extends main {



    // ----------------------------------------------------------------------------------------------------------------
    //Méthode principale permettant de créer/remplir et sauvegarder le compte rendu Powerpoint
    // ----------------------------------------------------------------------------------------------------------------
    public static void CreateFillSaveCrPowerpoint(String versionCibleppt, String moisSoldeppt, String repertoireFichierTargetppt) throws IOException {

        //On récupère le modele powerpoint dans les ressources
        String fileNamePptx = "modele.pptx";
        ClassLoader classLoaderPptx = new main().getClass().getClassLoader();
        classLoaderPptx.getResource(fileNamePptx);
        File fileSourcePptx = new File(classLoaderPptx.getResource(fileNamePptx).getFile());

        //On copie le fichier source et on le colle dans le dossier indiqué par l'utilisateur via la dialogbox
        //Ci-dessous on créer notre nouveau fichier avec le chemin indiqué par l'utilisateur et le nom du fichier ppt
        File filePptx = new File(repertoireFichierTarget + "\\" + fileSourcePptx.getName());
        try
        {
            Files.copy(fileSourcePptx.toPath(), Paths.get(filePptx.getPath()),REPLACE_EXISTING);
            System.out.println("Initialisation du CR pptx...");
        }catch (IOException e)
        {
            e.printStackTrace();
        }


        //On modifie le powerpoint
        FileInputStream inputstream = new FileInputStream(filePptx);
        XMLSlideShow ppt = new XMLSlideShow(inputstream);
        //Remplissage du Compte-rendu Powerpoint
        FillPpt.RemplissagePowerpoint(ppt);
        //enregistrer les modifications
        FileOutputStream fis = new FileOutputStream(filePptx);
        ppt.write(fis);
        fis.close();

    }

    // ----------------------------------------------------------------------------------------------------------------
    //Méthode permettant de remplir le Compte rendu powerpoint
    // ----------------------------------------------------------------------------------------------------------------
    private static void RemplissagePowerpoint(XMLSlideShow ppt) throws IOException {

        //on récupère les slide
        XSLFSlide slide1 = ppt.getSlides().get(0);
        XSLFSlide  slide2 = ppt.getSlides().get(1);
        XSLFSlide  slide3 = ppt.getSlides().get(2);
        XSLFSlide  slide4 = ppt.getSlides().get(3);
        XSLFSlide  slide5 = ppt.getSlides().get(4);
        XSLFSlide  slide6 = ppt.getSlides().get(5);
        XSLFSlide  slide7 = ppt.getSlides().get(6);
        XSLFSlide  slide8 = ppt.getSlides().get(7);
        XSLFSlide  slide9 = ppt.getSlides().get(8);

        //Remplissage de la slide 1
        RemplissageSlide1(slide1);
        //Remplissage de la slide 2
        RemplissageSlide2(ppt, slide2);
        //Remplissage de la slide 3
        RemplissageSlide3(ppt, slide3);
        //Remplissage de la slide 4
        RemplissageSlide4(ppt, slide4);
        //Remplissage de la slide 5
        RemplissageSlide5(ppt, slide5);
        //Remplissage de la slide 6
        RemplissageSlide6(ppt, slide6);
        //Remplissage de la slide 7
        RemplissageSlide7(ppt, slide7);
        //Remplissage de la slide 8
        RemplissageSlide8(ppt, slide8);
        //Remplissage de la slide 9
        RemplissageSlide9(ppt, slide9);
    }

    // ----------------------------------------------------------------------------------------------------------------
    //Méthode permettant de remplir la Slide 1 du compte rendu Powerpoint
    // ----------------------------------------------------------------------------------------------------------------
    private static void RemplissageSlide1(XSLFSlide  slide1) throws IOException
    {
        //Récupération du titre et écriture du titre
        UtilsPpt.WriteTittle("Rapport DEMO Version " + versionCible + " solde " + moisSolde, slide1);

        // On oublie pas de rajouter le pied de page
        UtilsPpt.AddFooterSlide(slide1, versionCible, moisSolde);
        System.out.println(Utils.ANSI_BLUE+ "Création/Remplissage Slide 1  du CR pptx Terminée..." + Utils.ANSI_RESET);
    }

    // ----------------------------------------------------------------------------------------------------------------
    //Méthode permettant de remplir la Slide 2 du compte rendu Powerpoint
    // ----------------------------------------------------------------------------------------------------------------
    private static void RemplissageSlide2(XMLSlideShow ppt, XSLFSlide  slide2) throws IOException
    {
        //Ajout de l'image du graphique radar (graphe1.png)
        XSLFPictureShape pic = UtilsPpt.AddImageToSlide(repertoireFichierTarget + "\\" + "Graphe1.png", ppt, slide2);
        pic.setAnchor(new Rectangle(103,60,570,435));


        // On oublie pas de rajouter le pied de page
        UtilsPpt.AddFooterSlide(slide2, versionCible, moisSolde);
        System.out.println(Utils.ANSI_BLUE + "Création/Remplissage Slide 2  du CR pptx Terminée..."+ Utils.ANSI_RESET);
    }

    // ----------------------------------------------------------------------------------------------------------------
    //Méthode permettant de remplir la Slide 3 du compte rendu Powerpoint
    // ----------------------------------------------------------------------------------------------------------------
    private static void RemplissageSlide3(XMLSlideShow ppt, XSLFSlide  slide3) throws IOException
    {
        //Ajout de l'image de l'image synthese (Synthese.png)
        XSLFPictureShape pic = UtilsPpt.AddImageToSlide(repertoireFichierTarget + "\\" + "Synthese.png", ppt, slide3);
        pic.setAnchor(new Rectangle(40,95,710,180));

        // On oublie pas de rajouter le pied de page
        UtilsPpt.AddFooterSlide(slide3, versionCible, moisSolde);
        System.out.println(Utils.ANSI_BLUE + "Création/Remplissage Slide 3  du CR pptx Terminée..." + Utils.ANSI_RESET);
    }

    // ----------------------------------------------------------------------------------------------------------------
    //Méthode permettant de remplir la Slide 4 du compte rendu Powerpoint
    // ----------------------------------------------------------------------------------------------------------------
    private static void RemplissageSlide4(XMLSlideShow ppt, XSLFSlide  slide4) throws IOException
    {

        //Ajout de l'image du graphique 2 (graphe2.png)
        XSLFPictureShape pic2 = UtilsPpt.AddImageToSlide(repertoireFichierTarget + "\\" + "Graphe2.png", ppt, slide4);
        pic2.setAnchor(new Rectangle(60,95,320,220));

        //Ajout de l'image du graphique 3 (graphe3.png)
        XSLFPictureShape pic3 = UtilsPpt.AddImageToSlide(repertoireFichierTarget + "\\" + "Graphe3.png", ppt, slide4);
        pic3.setAnchor(new Rectangle(395,95,320,220));

        // On oublie pas de rajouter le pied de page
        UtilsPpt.AddFooterSlide(slide4, versionCible, moisSolde);
        System.out.println(Utils.ANSI_BLUE + "Création/Remplissage Slide 4  du CR pptx Terminée..." + Utils.ANSI_RESET);
    }

    // ----------------------------------------------------------------------------------------------------------------
    //Méthode permettant de remplir la Slide 5 du compte rendu Powerpoint
    // ----------------------------------------------------------------------------------------------------------------
    private static void RemplissageSlide5(XMLSlideShow ppt, XSLFSlide  slide5) throws IOException
    {

        //Ajout de l'image de la version (LegendeVersions.png)
        XSLFPictureShape picVersion = UtilsPpt.AddImageToSlide(repertoireFichierTarget + "\\" + "LegendeVersions.png", ppt, slide5);
        picVersion.setAnchor(new Rectangle(200,75,350,21));

        //Ajout de l'image du graphique 7 (graphe7.png)
        XSLFPictureShape pic7 = UtilsPpt.AddImageToSlide(repertoireFichierTarget + "\\" + "Graphe7.png", ppt, slide5);
        pic7.setAnchor(new Rectangle(110,95,250,220)); //coordonnées où insérer l'image

        //Ajout de l'image du graphique 8 (graphe8.png)
        XSLFPictureShape pic8 = UtilsPpt.AddImageToSlide(repertoireFichierTarget + "\\" + "Graphe8.png", ppt, slide5);
        pic8.setAnchor(new Rectangle(430,95,250,220));

        // On oublie pas de rajouter le pied de page
        UtilsPpt.AddFooterSlide(slide5, versionCible, moisSolde);
        System.out.println(Utils.ANSI_BLUE + "Création/Remplissage Slide 5  du CR pptx Terminée..." + Utils.ANSI_RESET);
    }

    // ----------------------------------------------------------------------------------------------------------------
    //Méthode permettant de remplir la Slide 6 du compte rendu Powerpoint
    // ----------------------------------------------------------------------------------------------------------------
    private static void RemplissageSlide6(XMLSlideShow ppt, XSLFSlide  slide6) throws IOException
    {
        XSLFPictureShape picVersion = UtilsPpt.AddImageToSlide(repertoireFichierTarget + "\\" + "LegendeVersions.png", ppt, slide6);
        picVersion.setAnchor(new Rectangle(200,75,350,21));


        //Ajout de l'image du graphique 4 (graphe4.png)
        XSLFPictureShape pic4 = UtilsPpt.AddImageToSlide(repertoireFichierTarget + "\\" + "Graphe4.png", ppt, slide6);
        pic4.setAnchor(new Rectangle(60,95,210,220));


        //Ajout de l'image du graphique 5 (graphe5.png)
        XSLFPictureShape pic5 = UtilsPpt.AddImageToSlide(repertoireFichierTarget + "\\" + "Graphe5.png", ppt, slide6);
        pic5.setAnchor(new Rectangle(280,95,210,220));


        //Ajout de l'image du graphique 6 (graphe6.png)
        XSLFPictureShape pic6 = UtilsPpt.AddImageToSlide(repertoireFichierTarget + "\\" + "Graphe6.png", ppt, slide6);
        pic6.setAnchor(new Rectangle(500,95,210,220));

        // On oublie pas de rajouter le pied de page
        UtilsPpt.AddFooterSlide(slide6, versionCible, moisSolde);
        System.out.println(Utils.ANSI_BLUE + "Création/Remplissage Slide 6  du CR pptx Terminée..." + Utils.ANSI_RESET);
    }

    // ----------------------------------------------------------------------------------------------------------------
    //Méthode permettant de remplir la Slide 7 du compte rendu Powerpoint
    // ----------------------------------------------------------------------------------------------------------------
    private static void RemplissageSlide7(XMLSlideShow ppt, XSLFSlide  slide7) throws IOException
    {
        XSLFPictureShape picVersionRef = UtilsPpt.AddImageToSlide(repertoireFichierTarget + "\\" + "VersionRef.png", ppt, slide7);
        picVersionRef.setAnchor(new Rectangle(200,100,150,31));

        XSLFPictureShape picVersionCible = UtilsPpt.AddImageToSlide(repertoireFichierTarget + "\\" + "VersionCible.png", ppt, slide7);
        picVersionCible.setAnchor(new Rectangle(505,100,150,31));

        //Ajout de l'image du TabAvantFluxMasse (TabAvantFluxMasse.png)
        XSLFPictureShape pic4 = UtilsPpt.AddImageToSlide(repertoireFichierTarget + "\\" + "TabAvantFluxMasse.png", ppt, slide7);
        pic4.setAnchor(new Rectangle(30,130,730,120));

        // On oublie pas de rajouter le pied de page
        UtilsPpt.AddFooterSlide(slide7, versionCible, moisSolde);
        System.out.println(Utils.ANSI_BLUE + "Création/Remplissage Slide 7  du CR pptx Terminée..." + Utils.ANSI_RESET);
    }

    // ----------------------------------------------------------------------------------------------------------------
    //Méthode permettant de remplir la Slide 8 du compte rendu Powerpoint
    // ----------------------------------------------------------------------------------------------------------------
    private static void RemplissageSlide8(XMLSlideShow ppt, XSLFSlide  slide8) throws IOException
    {
        XSLFPictureShape picVersionRef = UtilsPpt.AddImageToSlide(repertoireFichierTarget + "\\" + "VersionRef.png", ppt, slide8);
        picVersionRef.setAnchor(new Rectangle(200,100,150,31));

        XSLFPictureShape picVersionCible = UtilsPpt.AddImageToSlide(repertoireFichierTarget + "\\" + "VersionCible.png", ppt, slide8);
        picVersionCible.setAnchor(new Rectangle(505,100,150,31));

        //Ajout de l'image du TabApresFluxMasse (TabApresFluxMasse.png)
        XSLFPictureShape pic4 = UtilsPpt.AddImageToSlide(repertoireFichierTarget + "\\" + "TabApresFluxMasse.png", ppt, slide8);
        pic4.setAnchor(new Rectangle(30,130,730,120));

        // On oublie pas de rajouter le pied de page
        UtilsPpt.AddFooterSlide(slide8, versionCible, moisSolde);
        System.out.println(Utils.ANSI_BLUE + "Création/Remplissage Slide 8  du CR pptx Terminée..." + Utils.ANSI_RESET);
    }

    // ----------------------------------------------------------------------------------------------------------------
    //Méthode permettant de remplir la Slide 9 du compte rendu Powerpoint
    // ----------------------------------------------------------------------------------------------------------------
    private static void RemplissageSlide9(XMLSlideShow ppt, XSLFSlide  slide9) throws IOException
    {
        XSLFPictureShape picVersionRef = UtilsPpt.AddImageToSlide(repertoireFichierTarget + "\\" + "VersionRef.png", ppt, slide9);
        picVersionRef.setAnchor(new Rectangle(200,100,150,31));

        XSLFPictureShape picVersionCible = UtilsPpt.AddImageToSlide(repertoireFichierTarget + "\\" + "VersionCible.png", ppt, slide9);
        picVersionCible.setAnchor(new Rectangle(505,100,150,31));

        //Ajout de l'image du TabApresFluxValo (TabApresFluxValo.png)
        XSLFPictureShape pic4 = UtilsPpt.AddImageToSlide(repertoireFichierTarget + "\\" + "TabApresFluxValo.png", ppt, slide9);
        pic4.setAnchor(new Rectangle(30,130,730,120));

        // On oublie pas de rajouter le pied de page
        UtilsPpt.AddFooterSlide(slide9, versionCible, moisSolde);
        System.out.println(Utils.ANSI_BLUE + "Création/Remplissage Slide 9  du CR pptx Terminée..." + Utils.ANSI_RESET);
    }
}

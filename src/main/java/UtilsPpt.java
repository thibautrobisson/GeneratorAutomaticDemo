import org.apache.commons.io.IOUtils;
import org.apache.poi.sl.usermodel.PictureData;
import org.apache.poi.sl.usermodel.Placeholder;
import org.apache.poi.sl.usermodel.VerticalAlignment;
import org.apache.poi.xslf.usermodel.*;
import java.awt.*;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

public class UtilsPpt {



    //Méthode permettant d'ajouter le footer d'une slide ppt
    public static void AddFooterSlide(XSLFSlide slide1, String versionCible, String moisSolde)
    {
        //on récupère la hauteur du slide au cas où si le modèle change entre temps afin de bien placer le
        // footer tout en bas...
        int hauteurSlide = slide1.getSlideShow().getPageSize().height;
        XSLFTextShape footer= slide1.createTextBox();
        footer.setPlaceholder(Placeholder.SUBTITLE);
        footer.setAnchor(new Rectangle(50, hauteurSlide-40, 700, 10));
        footer.setText("Tests Performance DEMO         " + "Version : " + versionCible
                + "         Dump : Pré-clôture " + moisSolde)
                .setFontSize(10.0);
    }

    public static XSLFPictureShape AddImageToSlide(String cheminFichier, XMLSlideShow ppt, XSLFSlide  slide) throws IOException {
        //Ajout de l'image de la version
        //lecture de l'image
        File image = new File(cheminFichier);
        //conversion de l'image en byte
        byte[] picture = IOUtils.toByteArray(new FileInputStream(image));
        //ajout de l'image a la presentation
        PictureData picdataVersion = ppt.addPicture(picture, PictureData.PictureType.PNG);
        XSLFPictureShape pic = slide.createPicture(picdataVersion);
        return pic;
    }

    public static void WriteTittle(String tittle, XSLFSlide  slide)
    {
        //Récupération du titre et écriture du titre
        XSLFTextShape title1 = slide.getPlaceholder(0);
        XSLFTextRun textTittle = title1.setText(tittle);
        title1.setVerticalAlignment(VerticalAlignment.TOP);
        textTittle.setFontSize(25.0);
    }

    public static void WriteSubTittle(String subtittletext, XSLFSlide  slide)
    {
        //récupération du sous titre et écriture du sous titre
        XSLFTextShape subtittle = slide.getPlaceholder(1);
        XSLFTextRun textSubtittle = subtittle.setText(subtittletext);
        textSubtittle.setFontSize(15.0);
    }
}

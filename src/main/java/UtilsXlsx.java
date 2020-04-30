import com.spire.xls.Worksheet;
import com.sun.corba.se.spi.orbutil.threadpool.Work;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellAddress;
import javax.imageio.ImageIO;
import java.awt.*;
import java.awt.image.BufferedImage;
import java.io.File;
import java.io.IOException;

public class UtilsXlsx {

    // -------------------------------------------------------
    //Méthode permettant de convertir du rgb en hsb
    // -------------------------------------------------------
    public static Color convertRgbToHsb(int red, int green, int blue)
    {
        float[] hsb = Color.RGBtoHSB(red,green,blue,null);
        float hue = hsb[0];
        float saturation = hsb[1];
        float brightness = hsb[2];
        return Color.getHSBColor(hue, saturation, brightness);
    }

    // -------------------------------------------------------
    //Méthode permettant de convertir une durée en un nbre d'heure décimal: 01:15:15 donnera 1.25 heure
    // -------------------------------------------------------
    public static Double ConvertohhmmssToNbreHeureDecimal(String value)
    {
        if(value.isEmpty() || value == null)
        {
            return 0.00;
        }
        value = ConvertDateTohhmmss(value);
        //on extrait les heures et minutes
        String hh = value.substring(0, 2);
        String mm = value.substring(3, 5);

        double nbreHeure = Double.parseDouble(hh);
        double nbreMin = Double.parseDouble(mm);
        double pourcentageMin=nbreMin/60;
        double result = nbreHeure + pourcentageMin;
        result = Math.round(result * 100.0) / 100.0 ;//arrondi  avec 2 décimals
        return result;
    }


    // -------------------------------------------------------
    //Méthode permettant de convertir un string de type 0.3556 (pourcentage) en format double 35,56
    // -------------------------------------------------------
    public static int ConvertStringEvaluatedValuePourcentToDouble(String value)
    {
        if(value.isEmpty() || value == null)
        {
            return 0;
        }
        double valuechanged = Double.parseDouble(value);
        int valuechangedint = (int)Math.round(valuechanged*100);
        if(valuechangedint == 0)
        {
            return 1;
        }
        return valuechangedint;
    }

    // -------------------------------------------------------
    //Méthode permettant de découper une image à partir des coordonées x et y selon une largeur w et une hauteur h
    // -------------------------------------------------------
    public static void ConvertGraphExcelToImgPng(String repertoireFichierTarget, String image, int x, int y, int w, int h) throws IOException
    {
        BufferedImage graphe = ImageIO.read(new File(repertoireFichierTarget + "\\" + image));
        graphe = graphe.getSubimage(x,y,w,h);
        //Puis on construit notre nouvelle image redimensionnée (on écrase au passage l'ancienne crée à l'étape d'avant)...
        ImageIO.write(graphe, "png", new File(repertoireFichierTarget + "\\" + image));
        System.out.println(Utils.ANSI_BLUE + "Création " + image + " terminée" +  Utils.ANSI_RESET);
    }

    // -------------------------------------------------------
    //Méthode permettant de récupérer la valeur d'une cellule (col/line) et si valeur null on returne DefaultValue
    // -------------------------------------------------------
    public static String GetEnvalutedValue(String col, int line, Worksheet sheet, String DefaultValueNull)
    {
        String value;
        try{
            value = sheet.getCellRange(col + line).getEnvalutedValue();
        }catch (Exception e)
        {
            return DefaultValueNull;
        }
        return value;
    }

    // -------------------------------------------------------
    //Méthode permettant de récupérer la valeur d'une cellule (col/line) et si valeur null on returne DefaultValue via apachePoi
    // -------------------------------------------------------
    public static int GetEnvalutedValuePoi(int col, int line, Sheet sheet, int DefaultValueNull)
    {
        int value = DefaultValueNull;
        try{
            value =  (int)sheet.getRow(line).getCell(col).getNumericCellValue();
        }catch (Exception e)
        {
            return DefaultValueNull;
        }
        return value;
    }

    // -------------------------------------------------------
    ///Méthode permettant de récupérer la valeur d'une cellule dans un string
    // -------------------------------------------------------
    public static String getValueStringOfCellAdress(Sheet sheet, String cellAdresse)
    {
        CellAddress cellAddress = new CellAddress(cellAdresse);
        Row row = sheet.getRow(cellAddress.getRow());
        Cell cell = row.getCell(cellAddress.getColumn());
        return cell.getStringCellValue();
    }


    public static int getSumValueOfRangeCell(Worksheet sheet, String col, int firstLine, int lastLine)
    {
        int result = 0;
        for (int i= firstLine; i <= lastLine;i++)
        {
            result += sheet.getCellRange(col + i).getNumberValue();
        }
        return result;
    }

    //Private Methods

    // -------------------------------------------------------
    //Méthode permettant de convertir une date au format dd/mm/yy hh:mm:ss en hh:mm:ss
    // -------------------------------------------------------
    private static String ConvertDateTohhmmss(String value)
    {
        if (value.length() > 8)
        {
            return value.substring(value.length() - 8);
        }
        return "00:00:00";
    }




}

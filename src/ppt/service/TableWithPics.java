package ppt.service;

import java.awt.Rectangle;
import java.awt.geom.Rectangle2D;
import java.io.FileInputStream;
import java.io.FileOutputStream;

import org.apache.poi.POIXMLDocumentPart.RelationPart;
import org.apache.poi.util.IOUtils;
import org.apache.poi.xslf.usermodel.XMLSlideShow;
import org.apache.poi.xslf.usermodel.XSLFHyperlink;
import org.apache.poi.xslf.usermodel.XSLFPictureData;
import org.apache.poi.xslf.usermodel.XSLFRelation;
import org.apache.poi.xslf.usermodel.XSLFSlide;
import org.apache.poi.xslf.usermodel.XSLFTable;
import org.apache.poi.xslf.usermodel.XSLFTableCell;
import org.apache.poi.xslf.usermodel.XSLFTableRow;
import org.apache.poi.xslf.usermodel.XSLFTextBox;
import org.apache.poi.xslf.usermodel.XSLFTextRun;
import org.openxmlformats.schemas.drawingml.x2006.main.CTBlip;
import org.openxmlformats.schemas.drawingml.x2006.main.CTBlipFillProperties;
import org.openxmlformats.schemas.drawingml.x2006.main.CTRelativeRect;
import org.openxmlformats.schemas.drawingml.x2006.main.CTTableCell;

public class TableWithPics {
    public static void main(String[] args) throws Exception {
        XMLSlideShow pptx = new XMLSlideShow();
        // Imagine
        byte[] pictureData = IOUtils.toByteArray(new FileInputStream("leaf.jpg"));
        XSLFPictureData pd = pptx.addPicture(pictureData, org.apache.poi.sl.usermodel.PictureData.PictureType.PNG);

        XSLFSlide slide = pptx.createSlide();
        XSLFSlide slide2 = pptx.createSlide();

        //Let us add Text
        
        XSLFTextBox shape1 = slide.createTextBox();
        shape1.setAnchor(new Rectangle(50, 50, 200, 50));
        XSLFTextRun r1 = shape1.addNewTextParagraph().addNewTextRun();
        XSLFHyperlink link1 = r1.createHyperlink();
        r1.setText("Example Test"); // visible text
        link1.setAddress("http://poi.apache.org");  // link address

        XSLFTextBox shape2 = slide.createTextBox();
        shape2.setAnchor(new Rectangle(300, 50, 200, 50));
        XSLFTextRun r2 = shape2.addNewTextParagraph().addNewTextRun();
        XSLFHyperlink link2 = r2.createHyperlink();
        r2.setText("Go to the second slide"); // visible text
        link2.linkToSlide(slide2);  // link address
        
        
        XSLFTable table = slide.createTable();
        table.setAnchor(new Rectangle2D.Double(50, 150, 500, 20));

        XSLFTableRow row = table.addRow();
        XSLFTableCell cell = row.addCell();
        row.addCell().setText("100000");
      //  cell.setText("Cell 2");

        CTBlipFillProperties blipPr = ((CTTableCell)cell.getXmlObject()).getTcPr().addNewBlipFill();
        blipPr.setDpi(72);
        // http://officeopenxml.com/drwPic-ImageData.php
        CTBlip blib = blipPr.addNewBlip();
        blipPr.addNewSrcRect();
        CTRelativeRect fillRect = blipPr.addNewStretch().addNewFillRect();
        fillRect.setL(30000);
        fillRect.setR(30000);

        RelationPart rp = slide.addRelation(null, XSLFRelation.IMAGES, pd);
        blib.setEmbed(rp.getRelationship().getId());

        FileOutputStream fos = new FileOutputStream("TableWithPics.pptx");
        pptx.write(fos);
        fos.close();
    }
}
package ppt.service;

import java.awt.Color;
import java.awt.Rectangle;
import java.io.FileInputStream;
import java.io.FileOutputStream;

import org.apache.poi.util.IOUtils;
import org.apache.poi.xslf.usermodel.SlideLayout;
import org.apache.poi.xslf.usermodel.XMLSlideShow;
import org.apache.poi.xslf.usermodel.XSLFAutoShape;
import org.apache.poi.xslf.usermodel.XSLFHyperlink;
import org.apache.poi.xslf.usermodel.XSLFPictureData;
import org.apache.poi.xslf.usermodel.XSLFPictureShape;
import org.apache.poi.xslf.usermodel.XSLFShape;
import org.apache.poi.xslf.usermodel.XSLFSlide;
import org.apache.poi.xslf.usermodel.XSLFSlideLayout;
import org.apache.poi.xslf.usermodel.XSLFSlideMaster;
import org.apache.poi.xslf.usermodel.XSLFTable;
import org.apache.poi.xslf.usermodel.XSLFTableCell;
import org.apache.poi.xslf.usermodel.XSLFTableRow;
import org.apache.poi.xslf.usermodel.XSLFTextBox;
import org.apache.poi.xslf.usermodel.XSLFTextParagraph;
import org.apache.poi.xslf.usermodel.XSLFTextRun;
import org.apache.poi.xslf.usermodel.XSLFTextShape;

public class PptCreator {

    public static void main(String[] tsr) {
        try {
            XMLSlideShow ppt = new XMLSlideShow();
            ppt.createSlide();

            XSLFSlideMaster defaultMaster = ppt.getSlideMasters().get(0);
            XSLFSlideLayout layout = defaultMaster.getLayout(SlideLayout.TITLE_AND_CONTENT);
            XSLFSlide slide = ppt.createSlide(layout);

            // Slide2
            XSLFSlide imageSlide = ppt.createSlide();
            
            // Slide3
            XSLFSlide listSlide = ppt.createSlide(layout);

            // Slide3
            XSLFSlide tableSlide = ppt.createSlide();
            
            XSLFTextShape titleShape = listSlide.getPlaceholder(0);
            XSLFTextShape contentShape = listSlide.getPlaceholder(1);

            for (XSLFShape shape : slide.getShapes()) {
                if (shape instanceof XSLFAutoShape) {
                    // this is a template placeholder
                }
            }

            // Text

            XSLFTextBox shape = slide.createTextBox();
            XSLFTextParagraph p = shape.addNewTextParagraph();
            XSLFTextRun r = p.addNewTextRun();
            r.setText("Baeldung");
            r.setFontColor(Color.green);
            r.setFontSize(24.);

            // Link
            XSLFHyperlink link = r.createHyperlink();
            link.setAddress("http://www.baeldung.com");

            // Imagine
            byte[] pictureData = IOUtils.toByteArray(new FileInputStream("leaf.jpg"));

            XSLFPictureData pd = ppt.addPicture(pictureData, org.apache.poi.sl.usermodel.PictureData.PictureType.PNG);
            XSLFPictureShape picture = imageSlide.createPicture(pd);
            
            //List 
            XSLFTextShape content = listSlide.getPlaceholder(1);
            
            /*
            XSLFTextParagraph p2 = content.addNewTextParagraph();
            p2.setIndentLevel(0);
            p2.setBullet(true);
            XSLFTextRun r2 = p.addNewTextRun();
            r2 = p2.addNewTextRun();
            r2.setText("Bullet");*/
            
            XSLFTextParagraph p3 = content.addNewTextParagraph();
            p3.setBulletAutoNumber(org.apache.poi.sl.usermodel.AutoNumberingScheme.alphaLcParenRight, 1);
            p3.setIndentLevel(1);
            XSLFTextRun r4 = p3.addNewTextRun();
            r4.setText("Numbered List Item - 1");
            
            
            
            // Table
            XSLFTable tbl = tableSlide.createTable();
            tbl.setAnchor(new Rectangle(50, 50, 450, 300));

            int numColumns = 3;
            XSLFTableRow headerRow = tbl.addRow();
            headerRow.setHeight(50);

            for (int i = 0; i < numColumns; i++) {
                XSLFTableCell th = headerRow.addCell();
                XSLFTextParagraph p1 = th.addNewTextParagraph();
                p1.setTextAlign(org.apache.poi.sl.usermodel.TextParagraph.TextAlign.CENTER);
                XSLFTextRun r1 = p1.addNewTextRun();
                r1.setText("Header " + (i + 1));
                tbl.setColumnWidth(i, 150);
            }
            
            //Let us add an Image to table
            XSLFTableRow headerRow1 = tbl.addRow();
            headerRow1.setHeight(50);
            XSLFTableCell th = headerRow.addCell();
            
            

            FileOutputStream out;
            out = new FileOutputStream("powerpoint.pptx");
            ppt.write(out);
            out.close();
        } catch (Exception e) {
            // TODO Auto-generated catch block
            e.printStackTrace();
        }

    }
}

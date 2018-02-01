package com.dokyme.ppt2pdf;

import com.itextpdf.text.*;
import com.itextpdf.text.Image;
import com.itextpdf.text.Rectangle;
import com.itextpdf.text.pdf.PdfPCell;
import com.itextpdf.text.pdf.PdfPTable;
import com.itextpdf.text.pdf.PdfWriter;
import org.apache.poi.xslf.usermodel.XMLSlideShow;
import org.apache.poi.xslf.usermodel.XSLFSlide;
import org.apache.poi.xslf.usermodel.XSLFSlideShow;

import java.awt.*;
import java.awt.geom.AffineTransform;
import java.awt.geom.Rectangle2D;
import java.awt.image.AffineTransformOp;
import java.awt.image.BufferedImage;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 * @author dokym
 */
public class Convertion {

    private boolean isVertical;
    private int columns;
    private int rows;

    private List<XMLSlideShow> inputPPTs;
    private Map<String, Integer> indexes;
    private Document document;
    private int index;
    private PdfPTable table;
    private Dimension pgSize;
    private float[] slideSize;
    private float[] margins;
    private AffineTransformOp ato;

    public static Convertion build(boolean isVertical, int columns, int rows) {
        Convertion convertion = new Convertion();
        convertion.isVertical = isVertical;
        convertion.columns = columns;
        convertion.rows = rows;
        return convertion;
    }

    private Convertion() {
        index = 0;
        indexes = new HashMap<>();
        document = new Document();

    }

    private void initPgSize(XMLSlideShow ppt) {
        //一个页面中，每行slide的宽度之和为整个page宽度的90%，其余留空，且间距相等。
        //高度同上。
        if (pgSize == null) {
            pgSize = ppt.getPageSize();
            slideSize = new float[2];
            margins = new float[2];
            float width = document.getPageSize().getWidth() * 0.9f / columns;
            float height = pgSize.height * width / pgSize.width;
            margins[0] = (document.getPageSize().getWidth() - width * columns) / (columns + 1);
            margins[1] = (document.getPageSize().getHeight() - height * rows) / (rows + 1);
            slideSize[0] = width;
            slideSize[1] = height;
            ato = new AffineTransformOp(AffineTransform.getScaleInstance(width / pgSize.width, height / pgSize.height), null);
        }
    }

    private void handleSlides(File input) {
        indexes.put(input.getName(), index);
        try {
            XMLSlideShow iPPT = new XMLSlideShow(new FileInputStream(input));
            initPgSize(iPPT);
            for (XSLFSlide slide : iPPT.getSlides()) {
                handleSlide(slide);
            }
        } catch (Exception e) {
            e.printStackTrace();
            System.exit(1);
        }
    }

    private void handleSlide(XSLFSlide slide) {
        BufferedImage bufferedImage = new BufferedImage(pgSize.width, pgSize.height, BufferedImage.TYPE_INT_RGB);
        Graphics2D graphics = bufferedImage.createGraphics();
        graphics.setPaint(Color.WHITE);
        graphics.fill(new Rectangle2D.Float(0, 0, pgSize.width, pgSize.height));
        try {
            if (index % (columns * rows) == (columns * rows - 1)) {
                table = new PdfPTable(columns);
                table.setWidthPercentage(90);
                table.setHorizontalAlignment(Element.ALIGN_CENTER);
                document.newPage();
                document.add(table);
            }
            slide.draw(graphics);
            java.awt.Image scaledImage = bufferedImage.getScaledInstance((int) slideSize[0], (int) slideSize[1], BufferedImage.SCALE_SMOOTH);
            scaledImage = ato.filter(bufferedImage, null);
            Image image = Image.getInstance(scaledImage, Color.WHITE);
            PdfPCell pCell = new PdfPCell(image);
            pCell.setBorder(Rectangle.NO_BORDER);
            pCell.setBackgroundColor(BaseColor.WHITE);
            table.addCell(new PdfPCell(image));
        } catch (Exception e) {
            e.printStackTrace();
            System.exit(1);
        }
    }

    public void convert(List<String> inputs, String output) {
        inputPPTs = new ArrayList<>();
        try {
            PdfWriter.getInstance(document, new FileOutputStream(output));
            document.setPageSize(PageSize.A4);
            document.open();

//            document.newPage();
            table = new PdfPTable(columns);
            table.setWidthPercentage(90);
            table.setHorizontalAlignment(Element.ALIGN_CENTER);
            for (String input : inputs) {
                File temp = new File(input);
                if (temp.isFile() && temp.getName().endsWith("pptx")) {
                    handleSlides(temp);
                } else if (temp.isDirectory()) {
                    for (File each : temp.listFiles()) {
                        if (each.getName().endsWith(".pptx")) {
                            handleSlides(each);
                        }
                    }
                }
            }
            document.add(table);
            document.close();

        } catch (Exception e) {
            e.printStackTrace();
            System.exit(1);
        }


    }

}

package com.dokyme.ppt2pdf;

import com.itextpdf.text.pdf.PdfWriter;
import org.apache.poi.sl.usermodel.SlideShow;
import org.apache.poi.sl.usermodel.SlideShowFactory;
import org.apache.poi.xslf.usermodel.XMLSlideShow;
import org.apache.poi.xslf.usermodel.XSLFSlide;
import org.apache.poi.xslf.usermodel.XSLFSlideShow;
import org.apache.poi.xwpf.usermodel.Document;
import org.apache.poi.xwpf.usermodel.XWPFDocument;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

/**
 * @author dokym
 */
public class Convertion {

    private boolean isVertical;
    private int slidesPerPage;

    public static Convertion build(boolean isVertical, int slidesPerPage) {
        Convertion convertion = new Convertion();
        convertion.isVertical = isVertical;
        convertion.slidesPerPage = slidesPerPage;
    }

    private Convertion() {

    }

    public void convert(List<String> inputs, String output) {
        List<XMLSlideShow> inputPPTs = new ArrayList<>();
        XMLSlideShow outputPPT = new XMLSlideShow();
        try {
            for (String input : inputs) {
                File temp = new File(input);
                if (temp.isFile() && temp.getName().endsWith("pptx")) {
                    inputPPTs.add(new XMLSlideShow());
                } else if (temp.isDirectory()) {
                    for (File each : temp.listFiles()) {
                        if (each.getName().endsWith(".pptx")) {
                            inputPPTs.add(new XMLSlideShow(new FileInputStream(each)));
                        }
                    }
                }
            }
            for (XMLSlideShow iPPT : inputPPTs) {
                for (XSLFSlide srcSlide : iPPT.getSlides()) {
                    outputPPT.createSlide().importContent(srcSlide);
                }
            }
            PdfWriter
        } catch (Exception e) {
            e.printStackTrace();
            System.exit(1);
        }


    }

}

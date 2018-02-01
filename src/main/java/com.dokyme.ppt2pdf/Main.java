package com.dokyme.ppt2pdf;

import com.itextpdf.text.Document;
import com.itextpdf.text.Paragraph;
import com.itextpdf.text.pdf.PdfWriter;
import org.apache.commons.cli.*;

import java.io.FileOutputStream;
import java.util.ArrayList;
import java.util.List;

/**
 * @author dokym
 */
public class Main {

    private boolean isVertical;
    private List<String> inputs;
    private String output;
    private int slidesPerPage;

    public static void main(String[] args) {
        Main main = new Main(args);
        main.run();
//        main.test();
    }

    private Main(String[] args) {
//        parseCmdArgs(args);
    }

    private void parseCmdArgs(String[] args) {
        Options options = new Options();
        options.addOption(Option.builder().argName("i").longOpt("input").desc("input file or directory").optionalArg(false).hasArg().required().type(String.class).valueSeparator(' ').build());
        options.addOption(Option.builder().argName("o").longOpt("output").desc("output file").optionalArg(true).hasArg().required(false).type(String.class).build());
        options.addOption(Option.builder().argName("v").longOpt("vertical").desc("Emmmmmmmmm").optionalArg(true).hasArg(false).required(false).build());
        options.addOption(Option.builder().argName("p").longOpt("per").desc("number of slides per page").optionalArg(true).hasArg(true).required(false).type(Integer.class).build());

        CommandLineParser parser = new DefaultParser();
        try {
            CommandLine cmd = parser.parse(options, args);
            if (cmd.hasOption("o")) {
                output = cmd.getOptionValue("o");
            } else {
                output = "./output.pdf";
            }
            if (cmd.hasOption("v")) {
                isVertical = true;
            } else {
                isVertical = false;
            }
            if (cmd.hasOption("p")) {
                slidesPerPage = Integer.valueOf(cmd.getOptionValue("p"));
            } else {
                slidesPerPage = 6;
            }
            inputs = new ArrayList<>();
            for (String i : cmd.getOptionValues("i")) {
                inputs.add(i);
            }
        } catch (ParseException e) {
            e.printStackTrace();
            System.out.println("Command line arguements parsing error.Please check arguements and try again.");
            System.exit(1);
        } catch (Exception e) {
            e.printStackTrace();
            System.exit(1);
        }
    }

    private void run() {
        Convertion convertion = Convertion.build(true, 3, 2);
        List<String> input = new ArrayList<>();
        input.add("test.pptx");
        convertion.convert(input, "output.pdf");
    }

    private void test() {
        try {
            //Step 1—Create a Document.
            Document document = new Document();
            //Step 2—Get a PdfWriter instance.
            PdfWriter.getInstance(document, new FileOutputStream("createSamplePDF.pdf"));
            //Step 3—Open the Document.
            document.open();
            //Step 4—Add content.
            document.add(new Paragraph("Hello World"));
            //Step 5—Close the Document.
            document.close();
        }catch (Exception e){
            e.printStackTrace();
        }
    }

}

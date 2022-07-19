package main;

import org.apache.poi.xwpf.converter.pdf.PdfConverter;
import org.apache.poi.xwpf.converter.pdf.PdfOptions;
import org.apache.poi.xwpf.usermodel.*;

import java.io.*;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.List;

public class Main {
    static String ORG_NM = "삼성카드";
    static String ORG_CD = "CD000010";
    static String CLIENT_ID = "zmffkdldjsxmid";
    static String CLIENT_SECRET = "zmffkdldjsxmtlzmflt";
    static String ALGORITHM = "AES/CBC/PKCS5Padding";
    static String ENC_KEY = "dkaghghkzldlqslek";
    static String ENC_IV = "dkghghkiv";

    public static void main(String[] args) throws IOException {
//        generateSample();
        modifySample();
    }

    private static void generateSample() throws IOException {

        XWPFDocument document = new XWPFDocument();

        XWPFParagraph titleParagraph = document.createParagraph();
        titleParagraph.setAlignment(ParagraphAlignment.CENTER);

        XWPFRun run = titleParagraph.createRun();
        run.setBold(true);
        run.setUnderline(UnderlinePatterns.SINGLE);
        run.setFontSize(26);
        run.setText("플러그인 발급확인서");
        run.addCarriageReturn();
        run.addBreak();

        XWPFParagraph nameParagraph = document.createParagraph();
        nameParagraph.setAlignment(ParagraphAlignment.LEFT);

        XWPFRun nameRun = nameParagraph.createRun();
        nameRun.setFontSize(15);
        nameRun.setText("고객사명: "+ ORG_NM);
        nameRun.addBreak();

        XWPFRun nameRun2 = nameParagraph.createRun();
        nameRun2.setFontSize(15);
        nameRun2.setText("발급용도: 마이데이터 플러그인 서비스 이용");
        nameRun2.addCarriageReturn();
        nameRun2.addBreak();
        //create table
        XWPFTable table = document.createTable();
        table.setWidthType(TableWidthType.AUTO);
        table.setWidth(XWPFTable.DEFAULT_PERCENTAGE_WIDTH);

        //create first row
        XWPFTableRow tableRowOne = table.getRow(0);
        tableRowOne.getCell(0).setText("CLIENT ID");
        tableRowOne.getCell(0).setWidth("1000");
        tableRowOne.addNewTableCell().setText("askdjalsdkjaklsd");

        //create second row
        XWPFTableRow tableRowTwo = table.createRow();
        tableRowTwo.getCell(0).setText("CLIENT_SECRET");
        tableRowTwo.getCell(1).setText("askdljasdklajsdlkajdsl");

        //create table
        XWPFTable table2 = document.createTable();

        //create first row
        XWPFTableRow table2RowOne = table2.getRow(0);
        table2RowOne.getCell(0).setText("ENC KEY");
        table2RowOne.addNewTableCell().setText("askdjalsdkjaklsd");

        //create second row
        XWPFTableRow table2RowTwo = table2.createRow();
        table2RowTwo.getCell(0).setText("ENC IV");
        table2RowTwo.getCell(1).setText("askdljasdklajsdlkajdsl");

        XWPFParagraph paragraph2 = document.createParagraph();
        paragraph2.setAlignment(ParagraphAlignment.RIGHT);
        XWPFRun run2 = paragraph2.createRun();
        run2.setBold(true);
        run2.setFontSize(14);

        Date today = new Date();
        SimpleDateFormat format = new SimpleDateFormat("yyyy년 MM월 dd일");
        run2.setText(format.format(today));

        run2.addCarriageReturn();
        run2.addBreak();

        XWPFRun run3 = paragraph2.createRun();
        run3.setBold(true);
        run3.setFontSize(14);
        run3.setText("상기의 플러그인 정보발급을 확인합니다.");

        String fileName = ORG_NM+".docx";
        FileOutputStream out = new FileOutputStream(fileName);
        document.write(out);
        out.close();

        System.out.println(fileName+ " written");
    }

    private static void modifySample() throws IOException {
        File f = new File("양식.docx"); // 경로 변경 필요
        FileInputStream fis = new FileInputStream(f.getAbsolutePath());

        XWPFDocument document = new XWPFDocument(fis);
        XWPFHeader header = document.getHeaderArray(0);
        for (XWPFParagraph paragraph : document.getParagraphs()) {
            for(XWPFRun run : paragraph.getRuns()) {
                if(run.text().contains("고객사명")) {
                    run.setText(" "+ORG_NM);
                }
                else if(run.text().contains("20")) {
                    Date today = new Date();
                    SimpleDateFormat format = new SimpleDateFormat("yy년 MM월 dd일");
                    run.setText(format.format(today));
                }
                System.out.print(run.text()+"\t");
            }
        }

        List<XWPFTable> tableArrayList = document.getTables();

        // 토큰 관련 항목
        XWPFTable tokenTable =  tableArrayList.get(0);
        XWPFTableCell orgCdCell = tokenTable.getRow(0).getCell(1);
        orgCdCell.removeParagraph(0);
        XWPFRun orgCdRun = orgCdCell.addParagraph().createRun();
        orgCdRun = setRunDefaultStyle(orgCdRun);
        orgCdRun.setText(ORG_CD);

        XWPFTableCell clientIdCell = tokenTable.getRow(1).getCell(1);
        clientIdCell.removeParagraph(0);
        XWPFRun clientIdRun = clientIdCell.addParagraph().createRun();
        clientIdRun = setRunDefaultStyle(clientIdRun);
        clientIdRun.setText(CLIENT_ID);

        XWPFTableCell clientSecretCell = tokenTable.getRow(2).getCell(1);
        clientSecretCell.removeParagraph(0);
        XWPFRun clientSecretRun = clientSecretCell.addParagraph().createRun();
        clientSecretRun = setRunDefaultStyle(clientSecretRun);
        clientSecretRun.setText(CLIENT_SECRET);

        // 암호화 관련 항목
        XWPFTable encryptTable = tableArrayList.get(1);
        XWPFTableCell encKeyCell = encryptTable.getRow(1).getCell(1);
        encKeyCell.removeParagraph(0);
        XWPFRun enckeyRun = encKeyCell.addParagraph().createRun();
        enckeyRun = setRunDefaultStyle(enckeyRun);
        enckeyRun.setText(ENC_KEY);

        XWPFTableCell encIvCell = encryptTable.getRow(2).getCell(1);
        encIvCell.removeParagraph(0);
        XWPFRun encIvRun = encIvCell.addParagraph().createRun();
        encIvRun = setRunDefaultStyle(encIvRun);
        encIvRun.setText(ENC_IV);

        String fileName = ORG_NM+".docx";
        String pdfFileName = ORG_NM+".pdf";
        System.out.println(fileName +" written");
        FileOutputStream out = new FileOutputStream(fileName);
        document.write(out);
        out.close();

//        PdfOptions options = PdfOptions.getDefault();
//        OutputStream out2 = new FileOutputStream(pdfFileName);
//        PdfConverter.getInstance().convert(document, out2, options);
//        out2.close();
//        convertDocxToPdf(fileName,pdfFileName);
    }

    private static XWPFRun setRunDefaultStyle(XWPFRun xwpfRun) {
        if (xwpfRun == null) {
            throw new NullPointerException();
        }
        xwpfRun.setFontSize(15);
        xwpfRun.setBold(true);
        xwpfRun.setFontFamily("Gulim");

        return xwpfRun;
    }

    private static void convertDocxToPdf (String docPath, String pdfPath) {
        try {
            InputStream doc = new FileInputStream(docPath);
            XWPFDocument document = new XWPFDocument(doc);
            PdfOptions options = PdfOptions.getDefault();
            OutputStream out = new FileOutputStream(pdfPath);
            PdfConverter.getInstance().convert(document, out, options);
        } catch (IOException ex) {
            System.out.println(ex.getMessage());
        }
    }
}

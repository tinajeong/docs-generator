package main;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.xwpf.usermodel.*;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.Date;

public class Main {
    public static void main(String[] args) throws IOException {
        XWPFDocument document = new XWPFDocument();


        XWPFParagraph paragraph = document.createParagraph();
        paragraph.setAlignment(ParagraphAlignment.CENTER);

        XWPFRun run = paragraph.createRun();
        run.setBold(true);
        run.setFontSize(24);
        run.setText("플러그인 발급확인서");
        run.addCarriageReturn();
        run.addBreak();

        //create table
        XWPFTable table = document.createTable();
        table.setWidthType(TableWidthType.AUTO);

        //create first row
        XWPFTableRow tableRowOne = table.getRow(0);
        tableRowOne.getCell(0).setText("CLIENT ID");
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

        FileOutputStream out = new FileOutputStream(new File("ORG_NM.docx"));
        document.write(out);
        out.close();

        System.out.println("ORG_NM.docx written");

    }
}

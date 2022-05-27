package org.example;


import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.List;
import java.util.*;

import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.xwpf.model.XWPFHeaderFooterPolicy;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFFooter;
import org.apache.poi.xwpf.usermodel.XWPFHeader;

import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;

import org.apache.commons.lang3.StringUtils;

//https://github.com/jeecgboot/autopoi/blob/master/autopoi/src/main/java/org/jeecgframework/poi/word/parse/ParseWord07.java
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFFooter;
import org.apache.poi.xwpf.usermodel.XWPFHeader;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;

import org.apache.poi.xwpf.usermodel.*;
//import org.apache.poi.xwpf.sprm.*;

import org.jeecgframework.poi.cache.WordCache;
import org.jeecgframework.poi.util.PoiPublicUtil;
import org.jeecgframework.poi.word.entity.MyXWPFDocument;
import org.jeecgframework.poi.word.entity.WordImageEntity;
import org.jeecgframework.poi.word.entity.params.ExcelListEntity;
import org.jeecgframework.poi.word.parse.excel.ExcelEntityParse;
import org.jeecgframework.poi.word.parse.excel.ExcelMapParse;


import java.util.Iterator;
import java.util.List;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.xwpf.usermodel.IBodyElement;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFTable;

//import org.apache.xmlbeans.impl.piccolo.io.FileFormatException;


import java.util.Timer;
import java.util.TimerTask;
import java.util.concurrent.TimeUnit;

import java.io.BufferedReader;
import java.io.IOException;
import java.io.InputStreamReader;

import java.lang.reflect.Field;
import java.lang.reflect.Method;

public class WordBot {
    //public static void main(String[] args){

    public static void main(String[] args){
        Date date_start = new Date();
        String[] arrOfStr;

        XWPFRun run;
        XWPFRun currentRun = null;
        String currentText = "";
        String text;
        Boolean isfinde = false;
        List<Integer> runIndex = new ArrayList<Integer>();
        XWPFParagraph paragraph;
        Map<String, Object> map;
        BufferedReader reader = new BufferedReader(new InputStreamReader(System.in));

        // display time and date using toString()
        System.out.println(date_start.toString());

        System.out.println("Reading doc file:");
        try {
            //FileInputStream fis = new FileInputStream("X:/travail/SMTV_DVD2SMTV/UK6-french.docx");
            //FileInputStream fis = new FileInputStream("X:\\travail\\smtv-translation-bot\\AP 1342 p145 (Return of the King - SMCH) sf6 - table fix3 Priority-HUN - Copy2.docx");
            FileInputStream fis = new FileInputStream("X:\\travail\\smtv-translation-bot\\NWN 1479 sf4 - table fix1 - Highlight Gray Ignore.docx");
            XWPFDocument xdoc=new XWPFDocument(OPCPackage.open(fis));


            Date date2 = new Date();

            // display time and date using toString()
            System.out.println(date2.toString());

            XWPFHeaderFooterPolicy policy = new XWPFHeaderFooterPolicy(xdoc);
            //read header

            //readTablesDataInDocx(xdoc);

            for (XWPFTable table : xdoc.getTables()) {
                System.out.println(table.getRows().size());

                //in case you want to do more with the table cells...
                int rownInt = 0;
                for (XWPFTableRow row : table.getRows()) {
                    System.out.println("");
                    for (XWPFTableCell cell : row.getTableCells()) {
                        for (XWPFParagraph para : cell.getParagraphs()) {
                            String cellStr = para.getText();
                            arrOfStr = cellStr.split("\r");
                            //System.out.println(para.getText());
                            System.out.println(rownInt + "-0 : " + arrOfStr[0]);

                            paragraph = para;
                            //for (XWPFRun rn : paragraph.getRuns()) {
                            for (int i = 0; i < paragraph.getRuns().size(); i++) {

                                //https://www.tabnine.com/code/java/classes/org.apache.poi.hwpf.usermodel.CharacterRun


                                run = paragraph.getRuns().get(i);
                                text = run.getText(0);
                                // https://stackoverflow.com/questions/35253030/is-there-any-way-to-identify-character-styles-with-apache-poi-xwpf-documents

                                System.out.println(run.isBold());
                                System.out.println(run.isHighlighted());
                                if (run.isHighlighted() == true){
                                    System.out.println("Run's highlighted color: " + run.getTextHightlightColor());
                                    // https://ostack.cn/?qa=2500281/&show=2500282
                                    //System.out.println("run.getCTR().getHighlight()=" + run.getCTR().getRPr().getHighlight());
                                    //run.getCTR().addNewRPr().addNewHighlight().setVal(STHighlightColor.YELLOW);

                                    //https://coderedirect.com/questions/241656/how-can-i-set-background-colour-of-a-run-a-word-in-line-or-a-paragraph-in-a-do

                                    System.out.println("Found '" + text + "' highlighted, press enter to continue");
                                    //String name = reader.readLine();

                                }
                                System.out.println(run.isCapitalized());
                                System.out.println(run.getFontSize());

/*                                CTRPr ctrpr = run.getCTR().getRPr();
                                if (ctrpr != null && ctrpr.isSetHighlight()) {
                                    //This is highlighted
                                    System.out("Found highlighted text");
                                    System.out.println("Highlighted found :" + text);
                                    String name = reader.readLine("Press enter");

                                    // Printing the read line
                                }*/
                            }


                        }
                    }
                    //Arrays.toString
                    rownInt+=1;
                }
            }
            fis.close();

        } catch(Exception ex) {
            ex.printStackTrace();
        }
        Date date_end = new Date();

        // display time and date using toString()
        System.out.println(date_end.toString());

        long duration = date_end.getTime() - date_start.getTime();
        System.out.println("Time elapsed : " + formatDuration(duration));

    }

    private static String formatDuration(long duration) {
        long hours = TimeUnit.MILLISECONDS.toHours(duration);
        long minutes = TimeUnit.MILLISECONDS.toMinutes(duration) % 60;
        long seconds = TimeUnit.MILLISECONDS.toSeconds(duration) % 60;
        long milliseconds = duration % 1000;
        return String.format("%02d:%02d:%02d,%03d", hours, minutes, seconds, milliseconds);
    }


    private void changeValues(
            XWPFParagraph paragraph,
            XWPFRun currentRun,
            String currentText,
            List<Integer> runIndex,
            Map<String, Object> map)
            throws Exception {
        Object obj = PoiPublicUtil.getRealValue(currentText, map);
        if (obj instanceof WordImageEntity) { // 如果是图片就设置为图片
            currentRun.setText("", 0);
            //addAnImage((WordImageEntity) obj, currentRun);
        } else {
            currentText = obj.toString();
            currentRun.setText(currentText, 0);
        }
        for (int k = 0; k < runIndex.size(); k++) {
            paragraph.getRuns().get(runIndex.get(k)).setText("", 0);
        }
        runIndex.clear();
    }

}

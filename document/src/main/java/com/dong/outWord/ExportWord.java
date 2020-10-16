package com.dong.outWord;

import org.apache.poi.xwpf.usermodel.ParagraphAlignment;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;

import java.awt.*;
import java.awt.font.TextAttribute;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.util.HashMap;

public class ExportWord {

    public XWPFDocument createXWPFDocument() {
        XWPFDocument doc = new XWPFDocument();
        createTitleParagraph(doc);
        createTitleParagraph1(doc);
        createTitleParagraph2(doc);
        createTitleParagraph3(doc);
        createTitleParagraph4(doc);
        createTitleParagraph5(doc);
        createTitleParagraph6(doc);
        return doc;
    }

    public void createTitleParagraph(XWPFDocument document) {
        XWPFParagraph titleParagraph = document.createParagraph();    //新建一个标题段落对象（就是一段文字）
        titleParagraph.setAlignment(ParagraphAlignment.CENTER);//样式居中
        XWPFRun titleFun = titleParagraph.createRun();    //创建文本对象
        titleFun.setText("离职证明"); //设置标题的名字
        titleFun.setBold(true); //加粗
        titleFun.setColor("000000");//设置颜色
        titleFun.setFontSize(18);    //字体大小
        titleFun.addBreak();    //换行
    }
    public void createTitleParagraph1(XWPFDocument document) {
        for (int i = 1; i <= 1; i++) {
            XWPFParagraph titleParagraph = document.createParagraph();    //新建一个标题段落对象（就是一段文字）
            /*titleParagraph.setAlignment(ParagraphAlignment.RIGHT);//样式居中*/
            XWPFRun titleFun = titleParagraph.createRun();    //创建文本对象
            titleFun.setFontSize(14);
            titleFun.setFontFamily("宋体(中文)");
            titleFun.addTab();
            /*String test = "qqqqqq";*/
            //涉及下划线
           /* setParagraphFontInfoAndUnderLineStyle(titleParagraph, test, "宋体", "1D8C56",
                    "36", *//*false,*//* false, false, true, i,
                    "000000", false, 0,
                    null);*/
            titleFun.setText(" MMMMMM ");
            titleFun.setText("同志（身份证号码：XXXXXXXXXXXXX");
            String test1 = "XXXX";
            //涉及下划线
            /*setParagraphFontInfoAndUnderLineStyle(titleParagraph, test1, "宋体", "1D8C56",
                    "36", *//*false,*//* false, false, true, i,
                    "000000", false, 0,
                    null);*/
            titleFun.setText("   ），入");
            //字体大小
        }
    }
    public void createTitleParagraph2(XWPFDocument document) {
        XWPFParagraph titleParagraph = document.createParagraph();    //新建一个标题段落对象（就是一段文字）
        titleParagraph.setAlignment(ParagraphAlignment.LEFT);
        XWPFRun titleFun = titleParagraph.createRun();    //创建文本对象
        titleFun.setFontSize(14);
        titleFun.setFontFamily("宋体(中文)");
        titleFun.setText("职日期为");
        titleFun.setText("2019 ");
        titleFun.setText("年");
        titleFun.setText("03 ");
        titleFun.setText("月");
        titleFun.setText("11 ");
        titleFun.setText("日，因个人原因向公司提出离职，离");
        //字体大小
    }
    public void createTitleParagraph3(XWPFDocument document) {
        XWPFParagraph titleParagraph = document.createParagraph();    //新建一个标题段落对象（就是一段文字）
        titleParagraph.setAlignment(ParagraphAlignment.LEFT);
        XWPFRun titleFun = titleParagraph.createRun();    //创建文本对象
        titleFun.setFontSize(14);
        titleFun.setFontFamily("宋体(中文)");
        titleFun.setText("职时间为");
        titleFun.setText("2020 ");
        titleFun.setText("年");
        titleFun.setText("03 ");
        titleFun.setText("月");
        titleFun.setText("27 ");
        titleFun.setText("日，已与我公司解除劳动关系。");
    }
    public void createTitleParagraph4(XWPFDocument document) {
        XWPFParagraph titleParagraph = document.createParagraph();    //新建一个标题段落对象（就是一段文字）
        titleParagraph.setAlignment(ParagraphAlignment.LEFT);
        XWPFRun titleFun = titleParagraph.createRun();    //创建文本对象
        titleFun.setFontSize(14);
        titleFun.setFontFamily("宋体(中文)");
        titleFun.addTab();
        titleFun.setText("特此证明！");
        titleFun.addBreak();    //换行
        titleFun.addBreak();    //换行
        titleFun.addBreak();    //换行
    }
    public void createTitleParagraph5(XWPFDocument document) {
        XWPFParagraph titleParagraph = document.createParagraph();    //新建一个标题段落对象（就是一段文字）
        titleParagraph.setAlignment(ParagraphAlignment.RIGHT);
        XWPFRun titleFun = titleParagraph.createRun();    //创建文本对象
        titleFun.setFontSize(14);
        titleFun.setFontFamily("宋体(中文)");
        titleFun.addTab();
        HashMap<TextAttribute, Object> hm = new HashMap<TextAttribute, Object>();
        hm.put(TextAttribute.UNDERLINE, TextAttribute.UNDERLINE_ON); // 定义是否有下划线
        hm.put(TextAttribute.SIZE, 12); // 定义字号
        hm.put(TextAttribute.FAMILY, "Simsun"); // 定义字体名
        Font font = new Font(hm); // 生成字号为12，字体为宋体，字形带有下划线的字体
        String name = font.getName();
        titleFun.setText("公司名称：   "+name);
        titleFun.setText("XXXX   ");
    }
    public void createTitleParagraph6(XWPFDocument document) {
        XWPFParagraph titleParagraph = document.createParagraph();    //新建一个标题段落对象（就是一段文字）
        titleParagraph.setAlignment(ParagraphAlignment.RIGHT);
        XWPFRun titleFun = titleParagraph.createRun();    //创建文本对象
        titleFun.setFontSize(14);
        titleFun.setFontFamily("宋体(中文)");
        titleFun.setText("2020 ");
        titleFun.setText("年");
        titleFun.setText("03 ");
        titleFun.setText("月");
        titleFun.setText("27 ");
        titleFun.setText("日");
    }
    public void exportCheckWord(XWPFDocument document, String savePath) throws IOException {
        saveDocument(document, savePath);
    }
    public void saveDocument(XWPFDocument document, String savePath) throws IOException {
        OutputStream os = new FileOutputStream(savePath);
        document.write(os);
        os.close();
    }
    /*public void setParagraphFontInfoAndUnderLineStyle(XWPFParagraph p,
                                                      String content, String fontFamily, String colorVal,
                                                      String fontSize, *//*boolean isBlod, *//*boolean isItalic,
                                                      boolean isStrike, boolean isUnderLine, int underLineStyle,
                                                      String underLineColor, boolean isShd, int shdValue, String shdColor) {
        XWPFRun pRun = null;
        if (p.getRuns() != null && p.getRuns().size() > 0) {
            pRun = p.getRuns().get(0);
        } else {
            pRun = p.createRun();
        }
        pRun.setText(content);

        CTRPr pRpr = null;
        if (pRun.getCTR() != null) {
            pRpr = pRun.getCTR().getRPr();
            if (pRpr == null) {
                pRpr = pRun.getCTR().addNewRPr();
            }
        }
        // 设置字体
        CTFonts fonts = pRpr.isSetRFonts() ? pRpr.getRFonts() : pRpr
                .addNewRFonts();
        fonts.setAscii(fontFamily);
        fonts.setEastAsia(fontFamily);
        fonts.setHAnsi(fontFamily);
        // 设置字体大小
        CTHpsMeasure sz = pRpr.isSetSz() ? pRpr.getSz() : pRpr.addNewSz();
        sz.setVal(new BigInteger(fontSize));
        CTHpsMeasure szCs = pRpr.isSetSzCs() ? pRpr.getSzCs() : pRpr
                .addNewSzCs();
        szCs.setVal(new BigInteger(fontSize));
        // 设置下划线样式
        if (isUnderLine) {
            CTUnderline u = pRpr.isSetU() ? pRpr.getU() : pRpr.addNewU();
            u.setVal(STUnderline.Enum.forInt(Math.abs(underLineStyle % 100)));
            if (underLineColor != null) {
                u.setColor(underLineColor);
            }
        }
    }*/
}
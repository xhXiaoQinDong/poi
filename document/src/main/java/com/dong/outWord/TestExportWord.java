package com.dong.outWord;

import org.apache.poi.xwpf.usermodel.XWPFDocument;


public class TestExportWord {
    
    public static void main(String[] args) throws Exception {
        ExportWord ew = new ExportWord();
        XWPFDocument document = ew.createXWPFDocument();
        ew.exportCheckWord( document, "D:/expWord.docx");
        System.out.println("文档生成成功");
    }
}
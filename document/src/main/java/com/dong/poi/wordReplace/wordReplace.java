package com.dong.poi.wordReplace;

import org.apache.poi.POIXMLDocument;
import org.apache.poi.xwpf.usermodel.*;

import java.io.FileOutputStream;
import java.util.*;
/**
* @Author: 雪浪风尘
* @Date: 2020/10/16
*/
public class wordReplace {
    /**
    * @Author: 雪浪风尘
     * 1、思路就是使用map存放要修改的值以及修改后的值。遍历整个word文档，当遇到的字段与map中的key相同时，就替换。所以这个key没有必要
     * 加个${}之类的，反而更容易出现错误，怎样方便怎样写
     * 2、可以会存在一部分替换了，而另一部分却没有替换，可能是格式的问题，将那些能够替换的复制到不能替换的应该就可以了。
    */
    public static void searchAndReplace(String srcPath, String destPath, Map<String, String> map) {
        try {
            XWPFDocument document = new XWPFDocument(POIXMLDocument.openPackage(srcPath));
            /**
             * 替换段落中的指定文字
             */
            Iterator<XWPFParagraph> itPara = document.getParagraphsIterator();
            while (itPara.hasNext()) {
                XWPFParagraph paragraph = (XWPFParagraph) itPara.next();
                Set<String> set = map.keySet();
                Iterator<String> iterator = set.iterator();
                while (iterator.hasNext()) {
                    String key = iterator.next();
                    List<XWPFRun> run=paragraph.getRuns();
                    for(int i=0;i<run.size();i++)
                    {
                        if(run.get(i).getText(run.get(i).getTextPosition())!=null &&
                                run.get(i).getText(run.get(i).getTextPosition()).equals(key))
                        {
                            /**
                             * 参数0表示生成的文字是要从哪一个地方开始放置,设置文字从位置0开始
                             * 就可以把原来的文字全部替换掉了
                             */
                            run.get(i).setText(map.get(key),0);
                        }
                    }
                }
            }

            /**
             * 替换表格中的指定文字
             */
            // 替换表格中的指定文字
            Iterator<XWPFTable> itTable = document.getTablesIterator();//获得Word的表格
            while (itTable.hasNext()) { //遍历表格
                XWPFTable table = (XWPFTable) itTable.next();
                int count = table.getNumberOfRows();//获得表格总行数
                for (int i = 0; i < count; i++) { //遍历表格的每一行
                    XWPFTableRow row = table.getRow(i);//获得表格的行
                    List<XWPFTableCell> cells = row.getTableCells();//在行元素中，获得表格的单元格
                    for (XWPFTableCell cell : cells) {   //遍历单元格
                        for (Map.Entry<String, String> e : map.entrySet()) {
                            if (cell.getText().equals(e.getKey())) {//如果单元格中的变量和‘键’相等，就用‘键’所对应的‘值’代替。
                                cell.removeParagraph(0);//所以这里就要求每一个单元格只能有唯一的变量。
                                cell.setText(e.getValue());
                            }
                        }
                    }
                }
            }
            FileOutputStream outStream = null;
            outStream = new FileOutputStream(destPath);
            document.write(outStream);
            outStream.close();
        } catch (Exception e) {
            e.printStackTrace();
        }
    }
    public static void main(String[] args) {
        Map<String, String> map = new HashMap<>();
        map.put("no1", "雪浪风尘");
        map.put("no2", "千寻一醉");
        map.put("no3", "江山如画");
        map.put("no4", "融化黑暗之温暖");
        String srcPath = "D:\\1.docx";
        String destPath = "D:\\m.doc";
        searchAndReplace(srcPath, destPath, map);
    }
}

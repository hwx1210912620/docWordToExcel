package com.hwx.docwordtoexcel;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;

import java.io.*;
import java.util.List;
import java.util.logging.Logger;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

public class Demo3 {
    public static void main(String[] args) throws Exception{
        String filePath = "/Users/hehebangjia/Desktop/files/11";
        String outPath = "/Users/hehebangjia/Downloads/nearSatisfied/study/NiXiang/docWordToExcel/target/template.xls";
        File file = new File(filePath);
        File[] files = file.listFiles();
        Workbook workbook = new HSSFWorkbook(new FileInputStream(new File(outPath)));
        Sheet sheet = workbook.getSheetAt(0);
        for (int i = 0; i < files.length; i++) {
            sheet.createRow(i+1);
            if(files[i].getName().endsWith(".docx")) {
                String content = getContent(files[i], sheet.getRow(i + 1));
            }

        }
        OutputStream out = null;
        try {
            out = new FileOutputStream(outPath);
            workbook.write(out);
            out.close();
        } catch (Exception e){
            e.printStackTrace();
        }
    }

    private static boolean setRegx(boolean flag,Row row,String text){
        //Logger logger = new Logger("textLogger");
        //原告赢
        Pattern compile5 = Pattern.compile("被告[\\u4E00-\\u9FA5、]{0,}担[\\u4E00-\\u9FA5]{0,}案[\\u4E00-\\u9FA5]{0,}诉讼费");
        Matcher matcher5 = compile5.matcher(text);
        while (matcher5.find()){
            flag = false;
            row.createCell(16);
            Cell cell = row.getCell(16);
            cell.setCellValue(row.getCell(2).getStringCellValue());
        }
        Pattern compile7 = Pattern.compile("被告[\\u4E00-\\u9FA5、]{0,}担[\\u4E00-\\u9FA5]{0,}案[\\u4E00-\\u9FA5]{0,}受理费");
        Matcher matcher7 = compile7.matcher(text);
        while (matcher7.find()){
            flag = false;
            row.createCell(16);
            Cell cell = row.getCell(16);
            cell.setCellValue(row.getCell(2).getStringCellValue());
        }
        Pattern compile1 = Pattern.compile("案诉讼费[\\u4E00-\\u9FA50-9、（）]{0,}由[\\u4E00-\\u9FA5]{0,}被告[\\u4E00-\\u9FA5、]{0,}担");
        Matcher matcher = compile1.matcher(text);
        while (matcher.find()){
            flag = false;
            row.createCell(16);
            Cell cell = row.getCell(16);
            cell.setCellValue(row.getCell(2).getStringCellValue());
        }
        Pattern compile2 = Pattern.compile("案[\\u4E00-\\u9FA5]{0,}受理费[\\u4E00-\\u9FA50-9、]{0,}由[\\u4E00-\\u9FA5]{0,}被告[\\u4E00-\\u9FA5、]{0,}担");
        Matcher matcher2 = compile2.matcher(text);
        while (matcher2.find()){
            flag=false;
            row.createCell(16);
            Cell cell = row.getCell(16);
            cell.setCellValue(row.getCell(2).getStringCellValue());
        }
        //被告赢
        Pattern compile3 = Pattern.compile("案诉讼费[\\u4E00-\\u9FA50-9、]{0,}由[\\u4E00-\\u9FA5]{0,}原告[\\u4E00-\\u9FA5、]{0,}担");
        Matcher matcher3 = compile3.matcher(text);
        while (matcher3.find()){
            flag=false;
            row.createCell(16);
            Cell cell = row.getCell(16);
            cell.setCellValue(row.getCell(6).getStringCellValue());
        }
        Pattern compile4 = Pattern.compile("案[\\u4E00-\\u9FA5]{0,}受理费[\\u4E00-\\u9FA50-9、]{0,}由[\\u4E00-\\u9FA5]{0,}原告[\\u4E00-\\u9FA5、]{0,}担");
        Matcher matcher4 = compile4.matcher(text);
        while (matcher4.find()){
            flag=false;
            row.createCell(16);
            Cell cell = row.getCell(16);
            cell.setCellValue(row.getCell(6).getStringCellValue());
        }
        Pattern compile6 = Pattern.compile("原告[\\u4E00-\\u9FA5、]{0,}担[\\u4E00-\\u9FA5]{0,}案[\\u4E00-\\u9FA5]{0,}诉讼费");
        Matcher matcher6 = compile6.matcher(text);
        while (matcher6.find()){
            flag = false;
            row.createCell(16);
            Cell cell = row.getCell(16);
            cell.setCellValue(row.getCell(2).getStringCellValue());
        }
        return flag;
    }

    private static String getContent(File file, Row row) {
        //List<Policy_content> list = new ArrayList<>();
        try {
            InputStream is = new FileInputStream(file);  //需要将文件路更改为word文档所在路径。
            row.createCell(0);
            Cell fileNameCell = row.getCell(0);
            fileNameCell.setCellValue(file.getName());
            //POIFSFileSystem fs = new POIFSFileSystem(is);
            XWPFDocument doc = new XWPFDocument(is);
            List<XWPFParagraph> paras = doc.getParagraphs();
            StringBuilder sbYuan = new StringBuilder();
            StringBuilder sbBei = new StringBuilder();
            int j=0;//1原告，2被告
            if(paras.size()>2) {
                for (int i = 0; i < paras.size(); i++) {
                    String text = paras.get(i).getParagraphText();
                    System.out.println(text);
                    //胜诉方被告胜利
                /*if(text.contains("驳回")&&row.getCell(6)!=null){
                    row.createCell(16);
                    Cell cell = row.getCell(16);
                    cell.setCellValue(row.getCell(6).getStringCellValue());
                }*/
                    //胜诉方
                    if ((text.contains("案诉讼费") || text.contains("案件受理费") || text.contains("案受理费")) && row.getCell(16) == null) {
                        //第一种情况查询文中是否本来就存在诉讼费由原告或者被告承担的句子
                        boolean flag = true;
                        flag = setRegx(flag, row, text);
                        //文中不存在诉讼费由原告或者被告承担的句子，通过比较诉讼费大小来判断胜诉方
                        text = text.replaceAll("。", "");
                        String[] fuDanStr = new String[]{};
                        if (text.contains("承担")) {
                            fuDanStr = text.split("承担");
                        }
                        if (text.contains("负担")) {
                            fuDanStr = text.split("负担");
                        }
                        //第二种情况，判断两个负担费用比较小的是胜诉方
                        if (fuDanStr.length > 2 && flag) {
                            Integer res = Integer.MAX_VALUE;
                            String resStr = "";
                            for (int i1 = 0; i1 < fuDanStr.length; i1++) {
                                String s = fuDanStr[i1];
                                //if(s.indexOf("元")==-1) continue;
                                String str = s.substring(0, s.indexOf("元"));
                                Pattern compile = Pattern.compile("([0-9]{1,})");
                                Matcher matcher1 = compile.matcher(str);
                                Integer resIn = 0;
                                while (matcher1.find()) {
                                    resIn = Integer.valueOf(matcher1.group());
                                }
                                if (i1 > 0) {
                                    if (res > resIn) {
                                        res = resIn;
                                        resStr = fuDanStr[i1 - 1];
                                    }
                                } else {
                                    res = resIn;
                                    resStr = fuDanStr[i1];
                                }
                            }
                            String yuanSubStr = matchSubStr(resStr, row.getCell(2).getStringCellValue() != null ? row.getCell(2).getStringCellValue() : "");
                            String beiSubStr = matchSubStr(resStr, row.getCell(6).getStringCellValue() != null ? row.getCell(6).getStringCellValue() : "");
                            String resAgency = yuanSubStr.length() > beiSubStr.length() ? row.getCell(2).getStringCellValue() : row.getCell(6).getStringCellValue();
                            //越小越精准
                            if (yuanSubStr.length() == beiSubStr.length()) {
                                resAgency = row.getCell(2).getStringCellValue().length() < row.getCell(6).getStringCellValue().length() ? row.getCell(2).getStringCellValue() : row.getCell(6).getStringCellValue();
                            }
                            row.createCell(16);
                            Cell cell = row.getCell(16);
                            cell.setCellValue(resAgency);
                        }
                        //第二种情况，只有一个负担的就是败诉方
                        if (fuDanStr.length <= 2 && row.getCell(6) != null && flag) {
                            String resStr2 = "";
                            if (text.contains("负担")) {
                                resStr2 = text.substring(text.indexOf("，由") + 2, text.indexOf("负担"));
                            }
                            if (text.contains("承担")) {
                                resStr2 = text.substring(text.indexOf("，由") + 2, text.indexOf("承担"));
                            }
                            String agency = resStr2.split("、")[0];
                            if (row.getCell(6).getStringCellValue().contains(agency)) {
                                row.createCell(16);
                                Cell cell = row.getCell(16);
                                cell.setCellValue(row.getCell(2).getStringCellValue());
                            } else {
                                row.createCell(16);
                                Cell cell = row.getCell(16);
                                cell.setCellValue(row.getCell(6).getStringCellValue());
                            }
                        }
                    }
                    //案号
                    if (i == 2) {
                        String s = text.replaceAll(" ", "");
                        row.createCell(1);
                        Cell cell = row.getCell(1);
                        cell.setCellValue(s);
                    }
                    //原告
                    Pattern compile = Pattern.compile("原告（[\\u4E00-\\u9FA50-9、，]{0,}）[：|:]{0,}[\\u4E00-\\u9FA5]{0,}");
                    Matcher matcher = compile.matcher(text);
                    while (matcher.find()) {
                        j = 1;
                        String group = matcher.group();
                        String resStr = "";
                        if (group.contains("："))
                            resStr = group.substring(group.indexOf("）：") + 2);
                        if (group.contains(":"))
                            resStr = group.substring(group.indexOf("）:") + 2);
                        if (!group.contains(":") && !group.contains("：")) {
                            resStr = group.substring(group.indexOf("）") + 1);
                        }
                        sbYuan.append(resStr).append("\n");
                        row.createCell(2);
                        Cell cell = row.getCell(2);
                        cell.setCellValue(sbYuan.toString());
                    }
                    //处理不含冒号的原告被告，如果第二段文本不含冒号就是属于原告直接+名称的形式
                    String paragraphText1 = paras.get(3).getParagraphText();
                    if (paragraphText1.startsWith("原告") && !paragraphText1.contains(":") && !paragraphText1.contains("：") && text.length() < 60) {
                        if (text.startsWith("原告")&&text.contains("，")) {
                            if (text.contains("（") && text.contains("）")) {
                                text = text.replaceAll(text.substring(text.indexOf("（"), text.indexOf("）") + 1), "");
                            }
                            if ((text.substring(2, text.indexOf("，")).length() <= 4)) {
                                String yuanGao = "";
                                if (text.contains(",")) {
                                    yuanGao = text.substring(2, text.indexOf(","));
                                } else if (text.contains("，")) {
                                    yuanGao = text.substring(2, text.indexOf("，"));
                                } else
                                    yuanGao = text.substring(2);
                                sbYuan.append(yuanGao).append("\n");
                            }
                            row.createCell(2);
                            Cell cell = row.getCell(2);
                            cell.setCellValue(sbYuan.toString());
                        }
                        if (text.startsWith("被告")&&text.contains("，")) {
                            if (text.contains("（") && text.contains("）")) {
                                text = text.replaceAll(text.substring(text.indexOf("（"), text.indexOf("）") + 1), "");
                            }
                            if (text.substring(2, text.indexOf("，")).length() <= 4) {
                                String beiGao = "";
                                if (text.contains(",")) {
                                    beiGao = text.substring(2, text.indexOf(","));
                                } else if (text.contains("，")) {
                                    beiGao = text.substring(2, text.indexOf("，"));
                                } else
                                    beiGao = text.substring(2);
                                sbBei.append(beiGao).append("\n");
                            }
                            row.createCell(6);
                            Cell cell = row.getCell(6);
                            cell.setCellValue(sbBei.toString());
                        }
                    }
                    if (text.startsWith("原告：")) {
                        j = 1;
                        String yuanGao = "";
                        if (text.contains("，")) {
                            yuanGao = text.substring(text.indexOf("：") + 1, text.indexOf("，"));
                        } else {
                            yuanGao = text.substring(text.indexOf("：") + 1);
                        }
                        sbYuan.append(yuanGao).append("\n");
                        row.createCell(2);
                        Cell cell = row.getCell(2);
                        cell.setCellValue(sbYuan.toString());
                    }
                    //原告
                    if (text.startsWith("原告:")) {
                        j = 1;
                        String yuanGao = "";
                        if (text.contains("，")) {
                            yuanGao = text.substring(text.indexOf(":") + 1, text.indexOf("，"));
                        } else {
                            yuanGao = text.substring(text.indexOf(":") + 1);
                        }
                        sbYuan.append(yuanGao).append("\n");
                        row.createCell(2);
                        Cell cell = row.getCell(2);
                        cell.setCellValue(sbYuan.toString());
                    }
                    //原告法定代理人
                    if (j == 1 && text.contains("法定代表人：")) {
                        String faDing = text.substring(text.indexOf("：") + 1);
                        row.createCell(3);
                        Cell cell = row.getCell(3);
                        cell.setCellValue(faDing);
                    }
                    //原告委托诉讼人1
                    if (j == 1 && text.contains("代理人：") && row.getCell(4) == null) {
                        String weituo = text.substring(text.indexOf("：") + 1);
                        row.createCell(4);
                        Cell cell = row.getCell(4);
                        cell.setCellValue(weituo);
                        continue;
                    }
                    //原告委托诉讼人2
                    if (j == 1 && text.contains("代理人：") && row.getCell(5) == null) {
                        String weituo = text.substring(text.indexOf("：") + 1);
                        row.createCell(5);
                        Cell cell = row.getCell(5);
                        cell.setCellValue(weituo);
                        continue;
                    }
                    //被告
                    Pattern compileBei = Pattern.compile("被告（[\\u4E00-\\u9FA50-9、，]{0,}）[：|:]{0,}[\\u4E00-\\u9FA5]{0,}");
                    Matcher matcherBei = compileBei.matcher(text);
                    while (matcherBei.find()) {
                        j = 2;
                        String group = matcherBei.group();
                        String resStr = "";
                        if (group.contains("：")) {
                            resStr = group.substring(group.indexOf("）：") + 2);
                        }
                        if (group.contains(":")) {
                            resStr = group.substring(group.indexOf("）:") + 2);
                        }
                        if (!group.contains(":") && !group.contains("：")) {
                            resStr = group.substring(group.indexOf("）") + 1);
                        }
                        sbYuan.append(resStr).append("\n");
                        row.createCell(6);
                        Cell cell = row.getCell(6);
                        cell.setCellValue(sbYuan.toString());
                    }
                    if (text.startsWith("被告：")) {
                        j = 2;
                        String beiGao = "";
                        if (text.contains("，")) {
                            beiGao = text.substring(text.indexOf("：") + 1, text.indexOf("，"));
                        } else {
                            beiGao = text.substring(text.indexOf("：") + 1);
                        }
                        sbBei.append(beiGao).append("\n");
                        row.createCell(6);
                        Cell cell = row.getCell(6);
                        cell.setCellValue(sbBei.toString());
                    }
                    if (text.startsWith("被告:")) {
                        j = 2;
                        String beiGao = "";
                        if (text.contains("，")) {
                            beiGao = text.substring(text.indexOf(":") + 1, text.indexOf("，"));
                        } else {
                            beiGao = text.substring(text.indexOf(":") + 1);
                        }
                        sbBei.append(beiGao).append("\n");
                        row.createCell(6);
                        Cell cell = row.getCell(6);
                        cell.setCellValue(sbBei.toString());
                    }
                    //被告法定代理人
                    if (j == 2 && text.contains("法定代表人：")) {
                        String faDing = text.substring(text.indexOf("：") + 1);
                        row.createCell(7);
                        Cell cell = row.getCell(7);
                        cell.setCellValue(faDing);
                    }
                    //被告诉讼代理人1
                    if (j == 2 && text.contains("代理人：") && row.getCell(8) == null) {
                        String beiGaoSuSong = text.substring(text.indexOf("：") + 1);
                        row.createCell(8);
                        Cell cell = row.getCell(8);
                        cell.setCellValue(beiGaoSuSong);
                        continue;
                    }
                    //被告诉讼代理人2
                    if (j == 2 && text.contains("代理人：") && row.getCell(9) == null) {
                        String beiGaoSuSong = text.substring(text.indexOf("：") + 1);
                        row.createCell(9);
                        Cell cell = row.getCell(9);
                        cell.setCellValue(beiGaoSuSong);
                        continue;
                    }
                    //审判长
                    String s1 = text.replaceAll(" ", "");
                    String s = s1.replaceAll("　", "");
                    if (s.contains("审判长")) {
                        String resStr = s.replaceAll("审判长", "");
                        row.createCell(10);
                        Cell cell = row.getCell(10);
                        cell.setCellValue(resStr);
                    }
                    if (s.contains("审判员") && !s.contains("代理审判员")) {
                        String shenPanYuan = s.replaceAll("审判员", "");
                        row.createCell(11);
                        Cell cell = row.getCell(11);
                        cell.setCellValue(shenPanYuan);
                    }
                    if (s.contains("代理审判员")) {
                        String daiLiShenPanYuan = s.replaceAll("代理审判员", "");
                        row.createCell(12);
                        Cell cell = row.getCell(12);
                        cell.setCellValue(daiLiShenPanYuan);
                    }
                    if (s.contains("人民陪审员")) {
                        String renMinPeiShen = s.replaceAll("人民陪审员", "");
                        row.createCell(13);
                        Cell cell = row.getCell(13);
                        cell.setCellValue(renMinPeiShen);
                    }
                    if (s.contains("书记员") && !s.contains("代书记员")) {
                        String shuJiYuan = s.replaceAll("书记员", "");
                        row.createCell(14);
                        Cell cell = row.getCell(14);
                        cell.setCellValue(shuJiYuan);
                    }
                    if (s.contains("代书记员")) {
                        String daoShuJiYuan = s.replaceAll("代书记员", "");
                        row.createCell(15);
                        Cell cell = row.getCell(15);
                        cell.setCellValue(daoShuJiYuan);
                    }

                }
            }
        } catch (Exception e) {
            e.printStackTrace();
        }
        return "";
    }

    private static String matchSubStr(String s1,String s2){
        String max = s1.length() >= s2.length()?s1:s2;
        String min = s1.length() >= s2.length()?s2:s1;
        int l = 0;
        String s ="";
        for(int i=0;i<min.length();i++){
            for(int j=i+1;j<=min.length();j++){
                if(max.contains(min.substring(i,j)) && j-i>l){
                    l=j-i;
                    s=min.substring(i,j);
                }
            }
        }
        return s;
    }
}

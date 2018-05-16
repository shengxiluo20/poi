package com.chi;

import org.apache.commons.codec.binary.StringUtils;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.util.StringUtil;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;

import java.io.*;
import java.util.*;
import java.util.regex.Matcher;

/**
 * @author chi  2018-05-15 14:29
 **/
public class POI {

    public static void main(String[] args) throws Exception {

        String property = System.getProperty("user.dir");
        System.out.println(property);

        System.out.println("输入word文件名：");
        Scanner scan = new Scanner(System.in);
        // 判断是否还有输入
        String str1="";
        if (scan.hasNext()) {
            str1= scan.next();
            System.out.println("输入word文件名数据为：" + str1);
        }

        System.out.println("输入excel文件名：");
        // 判断是否还有输入
        String str2="";
        if (scan.hasNext()) {
            str2= scan.next();
            System.out.println("输入excel文件名数据为：" + str2);
        }
        scan.close();


        POI poi = new POI();
        String filePath = property + "/" + str1;
        InputStream is = new FileInputStream(filePath);
        XWPFDocument doc = new XWPFDocument(is);

        /*StringBuffer sb= new StringBuffer("");
        FileReader reader = new FileReader("D:\\sdlfj\\word list 1.txt");
        BufferedReader br = new BufferedReader(reader);
        String str = null;
        while((str = br.readLine()) != null) {
            sb.append(str+"/n");
            //System.out.println(str);
        }*/

        XlsMain ddl = new XlsMain();
        Workbook wb = null;
        try {
            wb = ddl.createWorkbook(property + "/" + str2);
        } catch (IOException e) {
            e.printStackTrace();
        }
        Map<String, Integer> paramsa = ddl.doSomething(wb);




        poi.replace(doc, paramsa);


        OutputStream os = new FileOutputStream("D:\\sdlfj\\write.docx");
        doc.write(os);
        poi.close(os);
        poi.close(is);

    }

    private void replace(XWPFDocument doc, Map<String, Integer> params) {
        Iterator<XWPFParagraph> iterator = doc.getParagraphsIterator();
        XWPFParagraph para;

        TreeSet<Object> objects = new TreeSet<Object>();

        while (iterator.hasNext()) {
            para = iterator.next();
            List<XWPFRun> runs;
            Matcher matcher;

            String paragraphText = para.getParagraphText();
            String[] split = paragraphText.split(" ");
            for (String s : split) {
                s = s.replaceAll("[\\pP‘’“”]", "").toLowerCase();
                if (params.containsKey(s)) {
                    Integer integer = params.get(s);
                    params.put(s,++integer);
                    //System.out.println(s+"出现过");
                    objects.add(s);
                }

            }
        }
        System.out.println("所有出现过的的单词:"+objects.toString());
        System.out.println("共匹配出" + objects.size() + "个单词");
        System.out.println("==================================");
        Iterator it = params.entrySet().iterator();
        while (it.hasNext()) {
            Map.Entry entry = (Map.Entry) it.next();
            String key = (String)entry.getKey();
            Integer value = (Integer)entry.getValue();
            if(!"".equals(key) && value!=null && value>0){
                System.out.println(key +"    出现了 " +value +" 次");
            }
        }
    }

    /**
     * 关闭输入流
     * @param is
     */
    private void close(InputStream is) {
        if (is != null) {
            try {
                is.close();
            } catch (IOException e) {
                e.printStackTrace();
            }
        }
    }

    /**
     * 关闭输出流
     * @param os
     */
    private void close(OutputStream os) {
        if (os != null) {
            try {
                os.close();
            } catch (IOException e) {
                e.printStackTrace();
            }
        }
    }

}

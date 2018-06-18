package com.baiyu.word2html;

import org.apache.poi.hwpf.HWPFDocument;
import org.apache.poi.hwpf.usermodel.Range;
import org.apache.poi.xwpf.extractor.XWPFWordExtractor;
import org.apache.poi.xwpf.usermodel.XWPFDocument;

import java.io.BufferedWriter;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileWriter;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;

public class Word2Html {

    private String tailStr="         \n" +
            "        </div>\n" +
            "      </div>\n" +
            "    </div>\n" +
            "    <!--#include file=\"./footer.shtml\"-->\n" +
            "    <script data-main=\"../js/app/news6.js\" src=\"../js/require.js\"></script>\n" +
            "  </body>\n" +
            "</html>";
    private String headStr="<!DOCTYPE html>\n" +
            "<html>\n" +
            "  <head>\n" +
            "    <meta name=\"generator\"\n" +
            "    content=\"HTML Tidy for HTML5 (experimental) for Windows https://github.com/w3c/tidy-html5/tree/c63cc39\" />\n" +
            "    <title>风险提示</title>\n" +
            "    <meta charset=\"UTF-8\" />\n" +
            "    <meta name=\"format-detection\" content=\"telephone=no\" />\n" +
            "    <link rel=\"stylesheet\" type=\"text/css\" href=\"../css/common.css?v=0201\" />\n" +
            "    <link rel=\"stylesheet\" type=\"text/css\" href=\"../css/news8.css?v=0201\" />\n" +
            "    <link rel=\"stylesheet\" type=\"text/css\" href=\"../css/bxnews.css?v=0201\" />\n" +
            "    <script async=\"\" src=\"https://hm.baidu.com/hm.js?7c0e9c96a74719fdef5910598c47bac8\"></script>\n" +
            "  </head>\n" +
            "  <body>\n" +
            "    <div class=\"bodyer\">\n" +
            "      <!--#include file=\"./header.shtml\"-->\n" +
            "      <div class=\"container\">\n" +
            "        <div class=\"bybody\">\n";

    public void wordList2Html(List<WordEntity> wordEntityList){
        for(WordEntity wordEntity:wordEntityList) {
            word2Html(wordEntity,"shtml");
        }
    }
    public void word2Html(WordEntity wordEntity,String postfix){
        try{
        String path = wordEntity.fileDir+"\\"+wordEntity.fileName+"."+ postfix;
        File file = new File(path);
        if(!file.exists()){
            file.getParentFile().mkdirs();
        }
        file.createNewFile();
        FileWriter fw = new FileWriter(file, true);
        BufferedWriter bw = new BufferedWriter(fw);
        String contentStr=headStr;
        contentStr+="<p class=\"bytitle\">"+wordEntity.tilte+"</p>\n";
        for (int i=0;i<wordEntity.paragraphs.size();i++) {
            contentStr+="<p class=\"bycontent\">"+ mytrim(wordEntity.paragraphs.get(i))+"</p>\n";
        }
        contentStr+=tailStr;
        bw.write(contentStr);
        bw.flush();
        bw.close();
        fw.close();
        }catch (Exception e){e.printStackTrace();}
    }
    public List<WordEntity> getWordEntityList(String filepath){
        List<WordEntity> wordEntityList=new ArrayList<>();
        File file = new File(filepath);
        if (!file.isDirectory()) {
            WordEntity wordEntity=new WordEntity();
            wordEntity.fileDir=filepath;
            wordEntity.fileName=file.getName().split(".")[0];
            wordEntity.postfix=file.getName().split(".")[1];
            getWordEntity(wordEntity);
            wordEntityList.add(wordEntity);
            System.out.println("path=" + file.getPath());
            /*System.out.println("文件");

            System.out.println("absolutepath=" + file.getAbsolutePath());
            System.out.println("name=" + file.getName());*/

        } else if (file.isDirectory()) {

            String[] filelist = file.list();
            for (int i = 0; i < filelist.length; i++) {
                File readfile = new File(filepath + "\\" + filelist[i]);
                if (!readfile.isDirectory()) {
                    WordEntity wordEntity=new WordEntity();
                    wordEntity.fileDir=filepath;
                    wordEntity.fileName=filelist[i].split("\\.")[0];
                    wordEntity.postfix=filelist[i].split("\\.")[1];
                    wordEntity=getWordEntity(wordEntity);
                    wordEntityList.add(wordEntity);
                    System.out.println("path=" + readfile.getPath());
                    /*
                    System.out.println("absolutepath="
                            + readfile.getAbsolutePath());
                    System.out.println("name=" + readfile.getName());*/

                } else if (readfile.isDirectory()) {
                    wordEntityList.addAll(getWordEntityList(filepath + "\\" + filelist[i]));
                }
            }

        }
        return wordEntityList;
    }
    public WordEntity getWordEntity(WordEntity wordEntity){
        return getWordEntity(wordEntity.fileDir,wordEntity.fileName,wordEntity.postfix);
    }
    public WordEntity getWordEntity(String fileDir, String fileName,String postfix){
        WordEntity wordEntity=new WordEntity();
        wordEntity.fileName=fileName;
        wordEntity.postfix=postfix;
        wordEntity.fileDir=fileDir;
        File file = new File(fileDir+"\\"+fileName+"."+postfix);
        String str = "";
        try {
            FileInputStream fis = new FileInputStream(file);
            if(postfix.equals("doc")){
                HWPFDocument doc = new HWPFDocument(fis);
                String docStr = doc.getDocumentText();
                List<String> docStrs=Arrays.asList(docStr.split("\n"));
                wordEntity.tilte=docStrs.get(0);
                wordEntity.paragraphs=docStrs.subList(1,docStrs.size());

                /*System.out.println(doc1);
                StringBuilder doc2 = doc.getText();
                System.out.println(doc2);
                Range rang = doc.getRange();
                String doc3 = rang.text();
                System.out.println(doc3);*/
                fis.close();
            }else if(postfix.equals("docx")){
                XWPFDocument xdoc = new XWPFDocument(fis);
                XWPFWordExtractor extractor = new XWPFWordExtractor(xdoc);
                String docStr = extractor.getText();
                //System.out.println(doc1);
                //String docStr = doc.getDocumentText();
                List<String> docStrs=Arrays.asList(docStr.split("\n"));
                wordEntity.tilte=docStrs.get(0);
                wordEntity.paragraphs=docStrs.subList(1,docStrs.size());
                fis.close();
            }else{

            }

        } catch (Exception e) {
            e.printStackTrace();
        }
        return  wordEntity;
    }
    public String mytrim(String s){
        String result = "";
        if(null!=s && !"".equals(s)){
            result = s.replaceAll("^[　*| *| *|//s*]*", "").replaceAll("[　*| *| *|//s*]*$", "");
        }
        return result;
    }
}

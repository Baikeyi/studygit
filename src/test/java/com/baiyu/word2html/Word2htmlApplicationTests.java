package com.baiyu.word2html;

import org.junit.Test;
import org.junit.runner.RunWith;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.boot.test.context.SpringBootTest;
import org.springframework.test.context.junit4.SpringRunner;

import java.util.List;

@RunWith(SpringRunner.class)
@SpringBootTest
public class Word2htmlApplicationTests {


    @Test
    public void contextLoads() {
        Word2Html word2Html=new Word2Html();
//        word2Html.getWordEntity("E:\\Downloads\\附件1：保健会风险教育类文章","关于“开门红”保险销售的风险提示","docx");
        List<WordEntity> wordEntityList=word2Html.getWordEntityList("E:\\Documents\\保险");
        word2Html.wordList2Html(wordEntityList);
    }

}

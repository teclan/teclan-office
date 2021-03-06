package com.teclan.office.word;

import com.teclan.office.OpenOffice;
import com.teclan.office.POIFactory;
import com.teclan.office.WordFactory;
import jdk.nashorn.internal.runtime.regexp.joni.constants.OPCode;
import org.apache.poi.xwpf.converter.pdf.PdfOptions;
import org.junit.Test;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.OutputStream;
import java.util.HashMap;
import java.util.Map;

public class WordTest {

    @Test
    public void tableOutPutTest() throws Exception {
        String templatePat = "template\\word\\车辆信息1.ftl";
        String outputFile = "output/车辆信息1.docx";

        Map<String,Object> map = new HashMap<String, Object>();
        map.put("licensePlate","桂C88888");
        map.put("regDate","2019年03月18日");
        map.put("engineNo","1234567890");
        map.put("frameNo","ABCDEFGHIJKLMN");
        map.put("owner","张三");
        map.put("drivingLicense","12345qwert");
        map.put("phone","13366668888");
        WordFactory.export(templatePat,map,outputFile,true);
    }


    @Test
    public void workProveTest() throws Exception {
        String templatePat = "template\\word\\工作证明.ftl";
        String outputFile = "output/工作证明.docx";

        Map<String,Object> map = new HashMap<String, Object>();
        map.put("name","谭炳健");
        map.put("sex","男");
        map.put("id","12345678");
        map.put("y1","2019");
        map.put("m1","03");
        map.put("d1","19");
        map.put("y2","2020");
        map.put("m2","07");
        map.put("d2","30");
        WordFactory.export(templatePat,map,outputFile,true);
    }

    @Test
    public void doc2Pdf() throws Exception {
        String doc = "output/车辆信息1.docx";

        String html = "output/车辆信息1.html";

        String pdf = "output/车辆信息1.pdf";
//
        POIFactory.docx2Html(doc,"/output/images",html);

        WordFactory.html2Pdf(html,pdf,null);

//        OpenOffice.convert(doc,pdf,true);
    }

}

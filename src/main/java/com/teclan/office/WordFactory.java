package com.teclan.office;

import com.itextpdf.text.Document;
import com.itextpdf.text.DocumentException;
import com.itextpdf.text.pdf.PdfWriter;
import com.itextpdf.tool.xml.XMLWorkerFontProvider;
import com.itextpdf.tool.xml.XMLWorkerHelper;
import freemarker.template.Configuration;
import freemarker.template.Template;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.*;
import java.nio.charset.Charset;
import java.util.Locale;
import java.util.Map;

public class WordFactory {
    private static final Logger LOGGER = LoggerFactory.getLogger(WordFactory.class);

    /**
     * 将 content 输出到 word 文档
     *
     * @param templatePath 模板文件绝对路径
     * @param content
     * @param outputFile
     * @throws Exception
     */
    public static void export(String templatePath, Map<String, Object> content, String outputFile) throws Exception {
        export(templatePath,content,outputFile,false);
    }

    /**
     * 将 content 输出到 word 文档
     *
     * @param templatePath 模板文件绝对路径
     * @param content
     * @param outputFile
     * @param cover 是否覆盖目标输出文件，若模板文件已存在。若目标文件存在但不覆盖则会抛出异常
     * @throws Exception
     */
    public static void export(String templatePath, Map<String, Object> content, String outputFile,Boolean cover) throws Exception {
        Configuration configuration = new Configuration(Configuration.VERSION_2_3_29);
//        configuration.setDefaultEncoding("GBK");
        File outFile = new File(outputFile);
        try {
            File template = new File(templatePath);
            if (!template.exists()) {
                throw new Exception(String.format("模板文件[%s]不存在！", template.getAbsolutePath()));
            }

            if (template.isDirectory()) {
                throw new Exception(String.format("模板文件[%s]异常，期望是一个文件，实际是一个目录！", template.getAbsolutePath()));
            }

            configuration.setDirectoryForTemplateLoading(new File(template.getParent())); // FTL文件所存在的位置
            Template t = configuration.getTemplate(template.getName()); // 模板文件名

            if(outFile.exists() && !cover){
                throw new Exception(String.format("输出文件 %s 已存在，请使用重载方法设置是否替换目标文件",outFile.getAbsolutePath()));
            }
            outFile.getParentFile().mkdirs();

            Writer out = new BufferedWriter(new OutputStreamWriter(new FileOutputStream(outFile),"GBK"));
            t.process(content, out);
            LOGGER.error("导出word文档成功，模板 {},输出路径 {}", templatePath, outFile.getAbsolutePath());
        } catch (Exception e) {
            LOGGER.error("导出word文档出错，模板 {},输出路径 {}", templatePath, outFile.getAbsolutePath());
            LOGGER.error(e.getMessage(), e);
            throw e;
        }
    }

    public static void html2Pdf(String html,String pdf,String font) throws Exception {
        LOGGER.info("FreeMarker 文档转换开始,源 {},目标 {}", html, pdf);
        if(font==null){
            font="simhei.ttf";
        }

        Document document = new Document();
        PdfWriter writer = PdfWriter.getInstance(document, new FileOutputStream(pdf));
        document.open();
        XMLWorkerFontProvider fontImp = new XMLWorkerFontProvider(XMLWorkerFontProvider.DONTLOOKFORFONTS);
//        fontImp.register(font);
        XMLWorkerHelper.getInstance().parseXHtml(writer, document,
                new FileInputStream(html), null, Charset.forName("UTF-8"), fontImp);
        document.close();
        LOGGER.info("FreeMarker 文档转换完成,源 {},目标 {}", html, pdf);
    }


}

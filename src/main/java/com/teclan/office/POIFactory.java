package com.teclan.office;

import java.io.*;
import java.util.List;
import java.util.Map;
import javax.xml.parsers.DocumentBuilderFactory;
import javax.xml.transform.*;
import javax.xml.transform.dom.DOMSource;
import javax.xml.transform.stream.StreamResult;
import javax.xml.transform.stream.StreamSource;

import com.itextpdf.text.pdf.BaseFont;
import fr.opensagres.xdocreport.core.utils.StringUtils;
import fr.opensagres.xdocreport.itext.extension.font.IFontProvider;
import org.apache.commons.io.FileUtils;
import org.apache.poi.hwpf.HWPFDocument;
import org.apache.poi.hwpf.converter.PicturesManager;
import org.apache.poi.hwpf.converter.WordToHtmlConverter;
import org.apache.poi.hwpf.usermodel.PictureType;
import org.apache.poi.poifs.filesystem.DirectoryEntry;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.util.IOUtils;
import org.apache.poi.xwpf.converter.core.FileImageExtractor;
import org.apache.poi.xwpf.converter.core.FileURIResolver;
import org.apache.poi.xwpf.converter.pdf.PdfConverter;
import org.apache.poi.xwpf.converter.pdf.PdfOptions;
import org.apache.poi.xwpf.converter.xhtml.XHTMLConverter;
import org.apache.poi.xwpf.converter.xhtml.XHTMLOptions;
import org.apache.poi.xwpf.usermodel.*;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.w3c.dom.Document;

public class POIFactory {
    private static final Logger LOGGER = LoggerFactory.getLogger(WordFactory.class);

    public static void word2Html(String word, final String imgDir, String html) throws Exception {
        if(word.endsWith(".DOC")||word.endsWith(".doc")){
            doc2Html(word,imgDir,html);
        }else {
            docx2Html(word,imgDir,html);
        }
    }

    /**
     * 2007版本word转换成html
     * @param docx 源文件路径
     * @param imgDir 转换html后的图片存储路径
     * @param html 输出的 html 文件路径
     * @throws Exception 可能的异常
     */
    public static void docx2Html(String docx, final String imgDir, String html) throws Exception {
        docx2Html(new File(docx),new File(imgDir),new File(html),false);
    }

    /**
     * 2007版本word转换成html
     * @param docx 源文件
     * @param imgDir 转换html后的图片存储路径
     * @param html 输出的 html 文件
     * @throws Exception 可能的异常
     */
    public static void docx2Html(File docx, final File imgDir, File html,Boolean cover) throws Exception {
        LOGGER.info("POI 文档转换开始,源 {},目标 {}", docx.getAbsolutePath(), html.getAbsolutePath());
        if (!docx.exists()) {
            LOGGER.error("文件不存在:{}", docx.getAbsolutePath());
            throw new IOException(String.format("文件不存在:%s", docx.getAbsolutePath()));
        }
        if (!(docx.getName().endsWith(".docx") || docx.getName().endsWith(".DOCX"))) {
            LOGGER.error("不支持的文件格式，文件扩展名必须是 .docx 或 .DOCX ，实际文件：{}", docx);
            throw new IOException(String.format("不支持的文件格式，文件扩展名必须是 .docx 或 .DOCX ，实际文件：%s", docx));
        }

        if(html.exists() && !cover){
            throw new Exception(String.format("输出文件 %s 已存在，请使用重载方法设置是否替换目标文件", html.getAbsolutePath()));
        }

        InputStream in = new FileInputStream(docx.getAbsolutePath());
        XWPFDocument document = new XWPFDocument(in);

        XHTMLOptions options = XHTMLOptions.create().URIResolver(new FileURIResolver(imgDir));
        options.setExtractor(new FileImageExtractor(imgDir));
        options.setIgnoreStylesIfUnused(false);
        options.setFragment(true);

        OutputStream out = new FileOutputStream(html);
        XHTMLConverter.getInstance().convert(document, out, options);
        LOGGER.info("POI 文档转换完成,源 {},目标 {}", docx.getAbsolutePath(), html.getAbsolutePath());
    }

    /**
     * 2003版本word转换成html
     * @param doc 源文件路径
     * @param imageDir 转换html后的图片存储路径
     * @param html 输出的 html 文件路径
     * @throws Exception 能的异常
     */
    public static void doc2Html(String doc, final String imageDir, String html) throws Exception {
        doc2Html(new File(doc),new File(imageDir),new File(html),false);
    }

    /**
     * 2003版本word转换成html
     * @param doc 源文件路径
     * @param imageDir 转换html后的图片存储路径
     * @param html 输出的 html 文件路径
     * @throws Exception 能的异常
     */
    public static void doc2Html(String doc, final String imageDir, String html,Boolean cover) throws Exception {
        doc2Html(new File(doc),new File(imageDir),new File(html),cover);
    }


    /**
     * 2003版本word转换成html
     * @param doc 源文件路径
     * @param imgDir 转换html后的图片存储路径
     * @param html 输出的 html 文件路径
     * @param cover 是否覆盖生成目标文件
     * @throws Exception 能的异常
     */
    public static void doc2Html(File doc, final File imgDir, File html,Boolean cover) throws Exception {
        LOGGER.info("POI 文档转换开始,源 {},目标 {}", doc.getAbsolutePath(), html.getAbsolutePath());
        if (!doc.exists()) {
            LOGGER.error("文件不存在:{}", doc);
            throw new IOException(String.format("文件不存在:%s", doc));
        }
        if (!(doc.getName().endsWith(".doc") || doc.getName().endsWith(".DOC"))) {
            LOGGER.error("不支持的文件格式，文件扩展名必须是 .doc 或 .DOC ，实际文件：{}", doc);
            throw new IOException(String.format("不支持的文件格式，文件扩展名必须是 .doc 或 .DOC ，实际文件：%s", doc));
        }

        if(html.exists() && !cover){
            throw new Exception(String.format("输出文件 %s 已存在，请使用重载方法设置是否替换目标文件", html.getAbsolutePath()));
        }


        OutputStream outStream = new FileOutputStream(html);

        try {
            InputStream input = new FileInputStream(doc);
            HWPFDocument wordDocument = new HWPFDocument(input);
            WordToHtmlConverter wordToHtmlConverter = new WordToHtmlConverter(DocumentBuilderFactory.newInstance().newDocumentBuilder().newDocument());
            //设置图片存放的位置
            wordToHtmlConverter.setPicturesManager(new PicturesManager() {
                public String savePicture(byte[] content, PictureType pictureType, String suggestedName, float widthInches, float heightInches) {
                    if (!imgDir.exists()) {//图片目录不存在则创建
                        imgDir.mkdirs();
                    }
                    File file = new File(imgDir + suggestedName);
                    try {
                        OutputStream os = new FileOutputStream(file);
                        os.write(content);
                        os.close();
                    } catch (FileNotFoundException e) {
                        LOGGER.error(e.getMessage(), e);
                    } catch (IOException e) {
                        LOGGER.error(e.getMessage(), e);
                    }
                    return file.getAbsolutePath();
                }
            });

            //解析word文档
            wordToHtmlConverter.processDocument(wordDocument);
            Document htmlDocument = wordToHtmlConverter.getDocument();

            DOMSource domSource = new DOMSource(htmlDocument);
            StreamResult streamResult = new StreamResult(outStream);

            TransformerFactory factory = TransformerFactory.newInstance();
            Transformer serializer = factory.newTransformer();
            serializer.setOutputProperty(OutputKeys.ENCODING, "utf-8");
            serializer.setOutputProperty(OutputKeys.INDENT, "yes");
            serializer.setOutputProperty(OutputKeys.METHOD, "html");
            serializer.transform(domSource, streamResult);

            LOGGER.info("POI 文档转换完成,源 {},目标 {}", doc.getAbsolutePath(), html.getAbsolutePath());
        }catch (Exception e){
            LOGGER.error(e.getMessage(), e);
        }finally {
            IOUtils.closeQuietly(outStream);
        }
    }

    /**
     * html 转  word
     * @param html
     * @param word
     */
    public static void html2word(String html, String word) throws Exception {
       html2word(new File(html),new File(word),false);
    }

    /**
     * html 转  word
     * @param html
     * @param word
     */
    public static void html2word(File html, File word,Boolean cover) throws Exception {
        LOGGER.info("POI 文档转换开始,源 {},目标 {}", html.getAbsolutePath(), word.getAbsolutePath());

        if (!html.exists()) {
            LOGGER.error("文件不存在:{}", html.getAbsolutePath());
            throw new IOException(String.format("文件不存在:%s", html.getAbsolutePath()));
        }

        if(html.exists() && !cover){
            throw new Exception(String.format("输出文件 %s 已存在，请使用重载方法设置是否替换目标文件", word.getAbsolutePath()));
        }

        POIFSFileSystem poifs = new POIFSFileSystem();
        FileOutputStream ostream = null;
        ByteArrayInputStream bais = null;
        try {
            //HTML内容必须被<html><body></body></html>包装
            byte[] b = ("<html><body>" + FileUtils.readFileToString(html, "GBK") + "</body></html>").getBytes();
            bais = new ByteArrayInputStream(b);
            DirectoryEntry directory = poifs.getRoot();
            //WordDocument名称不允许修改
            directory.createDocument("WordDocument", bais);
            ostream = new FileOutputStream(word);
            poifs.writeFilesystem(ostream);
            LOGGER.info("POI 文档转换结束,源 {},目标 {}", html.getAbsolutePath(), word.getAbsolutePath());
        } catch (Exception e) {
            LOGGER.error(e.getMessage(), e);
        } finally {
            IOUtils.closeQuietly(ostream);
            IOUtils.closeQuietly(bais);
        }
    }

    /**
     * doc 和 docx 相互转换
     * @param src
     * @param des
     * @throws Exception
     */
    public static void wordConvert(String src,String des) throws Exception {

        if(!src.endsWith(".DOCX") && !src.endsWith(".DOCX")){
            throw new IOException(String.format("源文件扩展名必须是 .doc 或者 .DOCX，当前文件：%s",src));
        }

        if(!des.endsWith(".DOCX") && !des.endsWith(".DOCX")){
            throw new IOException(String.format("目标文件扩展名必须是 .doc 或者 .DOCX，当前文件：%s",src));
        }

        if(src.substring(0,src.lastIndexOf(".")).equalsIgnoreCase(des.substring(0,des.lastIndexOf(".")))){
            throw new IOException(String.format("源文件和目标文件不能是同一种格式，当前源文件：%s，目标文件:%s",src,des));
        }

        if (src.endsWith(".DOCX")||src.endsWith(".docx")){
            docx2Html(src, "output/images",src+".html");
        }else {
            doc2Html(src, "output/images",src+".html");
        }
        html2word(src+".html",des);
    }

    public static void wordConverterToPdf(InputStream source, OutputStream target, PdfOptions options,
                                          Map<String, String> params) throws Exception {
        XWPFDocument doc = new XWPFDocument(source);
        paragraphReplace(doc.getParagraphs(), params);
        for (XWPFTable table : doc.getTables()) {
            for (XWPFTableRow row : table.getRows()) {
                for (XWPFTableCell cell : row.getTableCells()) {
                    paragraphReplace(cell.getParagraphs(), params);
                }
            }
        }
        // 中文字体处理
        options.fontProvider(new IFontProvider() {
            public com.lowagie.text.Font getFont(String familyName, String encoding, float size, int style, java.awt.Color color) {

                try {
                    com.lowagie.text.pdf.BaseFont bfChinese = com.lowagie.text.pdf.BaseFont.createFont("STSong-Light", "UniGB-UCS2-H", BaseFont.NOT_EMBEDDED);
                    com.lowagie.text.Font fontChinese = new com.lowagie.text.Font(bfChinese, size, style, color);
                    if (familyName != null) {
                        fontChinese.setFamily(familyName);
                    }
                    return fontChinese;
                } catch (Exception e) {
                    LOGGER.error(e.getMessage(), e);
                    return null;
                }
            }

        });
        PdfConverter.getInstance().convert(doc, target, options);
    }

    /**
     * 替换段落中内容
     */
    private static void paragraphReplace(List<XWPFParagraph> paragraphs, Map<String, String> params) {
        if (params != null && !params.isEmpty()) {
            for (XWPFParagraph p : paragraphs) {
                for (XWPFRun r : p.getRuns()) {
                    String content = r.getText(r.getTextPosition());
                    if (StringUtils.isNotEmpty(content) && params.containsKey(content)) {
                        r.setText(content.replace("keyword", params.get(content)), 0);
                    }
                }
            }
        }
    }


    public static void xml2Html(String xml,String html) throws Exception{
        //创建XML的文件输入流
        FileInputStream fis=new FileInputStream("F:/123.xml");
        Source source=new StreamSource(fis);

        //创建XSL文件的输入流
        FileInputStream fis1=new FileInputStream("F:/123.xsl");
        Source template=new StreamSource(fis1);

        PrintStream stm=new PrintStream(new File("F:/123.html"));
        Result result=new StreamResult(stm);
        //根据XSL文件创建准个转换对象
        Transformer transformer=TransformerFactory.newInstance().newTransformer(template);
        //处理xml进行交换
        transformer.transform(source, result);
        fis1.close();
        fis.close();
    }
}




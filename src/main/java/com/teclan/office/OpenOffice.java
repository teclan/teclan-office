package com.teclan.office;

import com.artofsolving.jodconverter.DocumentConverter;
import com.artofsolving.jodconverter.openoffice.connection.OpenOfficeConnection;
import com.artofsolving.jodconverter.openoffice.connection.SocketOpenOfficeConnection;
import com.artofsolving.jodconverter.openoffice.converter.OpenOfficeDocumentConverter;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.*;

public class OpenOffice {
    private static final Logger LOGGER = LoggerFactory.getLogger(OpenOffice.class);

    public static void convert(String src, String des) throws Exception {
        convert(new File(src), new File(des), false);
    }

    public static void convert(String src, String des, Boolean cover) throws Exception {
        convert(new File(src), new File(des), cover);
    }


    public static void convert(File src, File des, Boolean cover) throws Exception {
        LOGGER.info("OpneOffice 文档转换开始,源 {},目标 {}", src.getAbsolutePath(), des.getAbsolutePath());
        if (des.exists() && !cover) {
            throw new Exception(String.format("输出文件 %s 已存在，请使用重载方法设置是否替换目标文件", des.getAbsolutePath()));
        }
        des.getParentFile().mkdirs();
        OpenOfficeConnection connection = new SocketOpenOfficeConnection(
                "127.0.0.1", 8100);

        try {
            connection.connect();
            DocumentConverter converter = new OpenOfficeDocumentConverter(connection);
            converter.convert(src, des);
        } catch (Exception e) {
            LOGGER.error("OpneOffice 文档转换异常,源 {},目标 {}", src.getAbsolutePath(), des.getAbsolutePath());
            LOGGER.error(e.getMessage(), e);
            throw e;
        } finally {
            connection.disconnect();
        }
        LOGGER.info("OpneOffice 文档转换成功,源 {},目标 {}", src.getAbsolutePath(), des.getAbsolutePath());
    }
}

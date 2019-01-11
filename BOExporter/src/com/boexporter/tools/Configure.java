package com.boexporter.tools;

import java.io.*;
import java.util.*;




public class Configure {

//    private static final Log log = LogFactory.getLog(ServerConfig.class);
    private static final long serialVersionUID = 1L;
    
    private static Properties config = null;    
    
    public Configure() {
        config = new Properties();
    }
    
    public Configure(String filePath) {
        config = new Properties();
        InputStream in = null;
        try {
        	/*
            ClassLoader CL = this.getClass().getClassLoader();
            InputStream in;
            if (CL != null) {
                in = CL.getResourceAsStream(filePath);
            }else {
                in = ClassLoader.getSystemResourceAsStream(filePath);
            }
            */
        	in = new FileInputStream(filePath);
            config.load(in);
        //    in.close();
        } catch (FileNotFoundException e) {
        //    log.error("服务器配置文件没有找到");
        	e.printStackTrace();
        } catch (Exception e) {
        //    log.error("服务器配置信息读取错误");
        	e.printStackTrace();
        }
    }
    
    public String getValue(String key) {
        if (config.containsKey(key)) {
            String value = config.getProperty(key);
            return value;
        }else {
            return "";
        }
    }
    
    public int getValueInt(String key) {
        String value = getValue(key);
        int valueInt = 0;
        try {
            valueInt = Integer.parseInt(value);
        } catch (NumberFormatException e) {
            e.printStackTrace();
            return valueInt;
        }
        return valueInt;
    }    
}
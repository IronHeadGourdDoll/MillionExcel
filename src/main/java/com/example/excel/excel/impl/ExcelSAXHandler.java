package com.example.excel.excel.impl;

import lombok.extern.slf4j.Slf4j;
import org.apache.poi.openxml4j.opc.OPCPackage;


import org.apache.poi.xssf.eventusermodel.XSSFReader;
import org.apache.poi.xssf.model.SharedStrings;
import org.apache.poi.xssf.model.SharedStringsTable;
import org.apache.poi.util.IOUtils;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.xml.sax.Attributes;
import org.xml.sax.InputSource;
import org.xml.sax.SAXException;
import org.xml.sax.XMLReader;
import org.xml.sax.helpers.DefaultHandler;
import org.xml.sax.helpers.XMLReaderFactory;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.lang.reflect.Field;
import java.math.BigDecimal;
import java.sql.Timestamp;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.HashMap;
import java.util.Iterator;
import java.util.List;
import java.util.Map;

/**
 * SAX方式处理Excel大文件
 */
@Slf4j
public class ExcelSAXHandler<T> extends DefaultHandler {
    private final Class<T> clazz;
    private final List<T> dataList = new ArrayList<>(100000);
    private final Map<String, Field> fieldMap = new HashMap<>();
    private final Map<String, Integer> headerMap = new HashMap<>();
    private T currentObject;
    private StringBuilder currentValue;
    private Field[] fields;
    private int currentCol = -1;
    private SharedStrings sst;
    private boolean isSSTString;
    private boolean isHeaderRow = true;

    public ExcelSAXHandler(Class<T> clazz) {
        this.clazz = clazz;
        this.fields = clazz.getDeclaredFields();
        
        // 初始化字段映射
        for (Field field : fields) {
            fieldMap.put(field.getName().toLowerCase(), field);
        }
    }

    public List<T> parse(InputStream inputStream) throws Exception {
        // 大文件检测
        long fileSize = inputStream.available();
        if (fileSize > 100_000_000) {
            log.warn("处理超大Excel文件(大小: {}MB)", fileSize/1024/1024);
        }

        // 创建临时文件以避免内存溢出
        File tempFile = null;
        FileInputStream tempFis = null;
        try {
            // 创建临时文件存储输入流内容
            tempFile = File.createTempFile("excel_import_", ".xlsx");
            try (FileOutputStream fos = new FileOutputStream(tempFile)) {
                byte[] buffer = new byte[8192];
                int bytesRead;
                while ((bytesRead = inputStream.read(buffer)) != -1) {
                    fos.write(buffer, 0, bytesRead);
                }
            }
            
            // 从临时文件读取数据
            tempFis = new FileInputStream(tempFile);
            OPCPackage pkg = OPCPackage.open(tempFis);
            XSSFReader reader = new XSSFReader(pkg);
            sst = reader.getSharedStringsTable();

            // 使用安全的XMLReader配置
            XMLReader parser = XMLReaderFactory.createXMLReader();
            parser.setFeature("http://apache.org/xml/features/disallow-doctype-decl", true);
            parser.setFeature("http://xml.org/sax/features/external-general-entities", false);
            parser.setFeature("http://apache.org/xml/features/nonvalidating/load-external-dtd", false);
            parser.setContentHandler(this);
            
            // 处理工作表数据
            Iterator<InputStream> sheetsIterator = reader.getSheetsData();
            if (sheetsIterator.hasNext()) {
                InputStream sheetStream = sheetsIterator.next();
                try {
                    InputSource sheetSource = new InputSource(sheetStream);
                    parser.parse(sheetSource);
                } finally {
                    IOUtils.closeQuietly(sheetStream);
                }
            }
            
            return dataList;
        } finally {
            // 清理资源
            IOUtils.closeQuietly(tempFis);
            if (tempFile != null && !tempFile.delete()) {
                log.warn("无法删除临时文件: {}", tempFile.getAbsolutePath());
                // 设置临时文件为JVM退出时删除
                tempFile.deleteOnExit();
            }
        }
    }

    @Override
    public void startElement(String uri, String localName, String qName,
                            Attributes attributes) throws SAXException {
        if ("row".equals(qName)) {
            currentCol = -1;
            if (!isHeaderRow) {
                try {
                    currentObject = clazz.getDeclaredConstructor().newInstance();
                } catch (Exception e) {
                    throw new SAXException("创建实例失败", e);
                }
            }
        } else if ("c".equals(qName)) {
            currentCol++;
            currentValue = new StringBuilder();
            isSSTString = "s".equals(attributes.getValue("t"));
        }
    }

    @Override
    public void characters(char[] ch, int start, int length) {
        if (currentValue != null) {
            currentValue.append(ch, start, length);
        }
    }

    @Override
    public void endElement(String uri, String localName, String qName) {
        try {
            if ("row".equals(qName)) {
                if (isHeaderRow) {
                    isHeaderRow = false;
                    // 表头处理完成后，清理一些不需要的变量
                    if (dataList.size() == 0) {
                        log.debug("表头解析完成，共解析 {} 个列名", headerMap.size());
                    }
                } else if (currentObject != null) {
                    // 检查基本类型必填字段
                    boolean isValid = true;
                    try {
                        for (Field field : fields) {
                            field.setAccessible(true);
                            if (field.getType().isPrimitive() && field.get(currentObject) == null) {
                                log.warn("必填字段 {} 为空，已跳过该行数据", field.getName());
                                isValid = false;
                                break;
                            }
                        }
                    } catch (Exception e) {
                        log.warn("数据校验异常: {}", e.getMessage());
                        isValid = false;
                    }
                    
                    if (isValid) {
                        dataList.add(currentObject);
                    }
                    
                    currentObject = null;
                    
                    // 每1万条数据执行GC，减少内存占用
                    if (dataList.size() % 10000 == 0) {
                        log.debug("已处理 {} 条数据，执行GC以释放内存", dataList.size());
                        System.gc();
                    }
                }
            } else if ("v".equals(qName)) {
                String value = currentValue.toString();
                if (isSSTString) {
                    try {
                        int index = Integer.parseInt(value);
                        if (index >= 0 && index < sst.getCount()) {
                            value = new XSSFRichTextString(String.valueOf(sst.getItemAt(index))).toString();
                        } else {
                            log.debug("共享字符串索引超出范围: {}", index);
                        }
                    } catch (NumberFormatException e) {
                        log.warn("共享字符串索引解析失败: {}", value);
                    }
                }
                value = value.trim();
                
                if (isHeaderRow) {
                    // 处理表头，使用trim和toLowerCase规范化列名
                    if (!value.isEmpty()) {
                        String normalizedHeader = value.trim().toLowerCase();
                        headerMap.put(normalizedHeader, currentCol);
                    }
                } else if (currentObject != null) {
                    // 使用更高效的方式查找列对应的字段
                    String headerName = findHeaderByColumnIndex(currentCol);
                    if (!headerName.isEmpty()) {
                        Field field = fieldMap.get(headerName.toLowerCase());
                        if (field != null) {
                            field.setAccessible(true);
                            // 即使值为空也尝试设置，避免字段为null
                            setFieldValue(field, currentObject, value);
                        }
                    }
                }
            }
        } catch (Exception e) {
            log.warn("处理Excel元素时发生异常: {}", e.getMessage());
            // 继续处理，避免整个导入过程中断
        } finally {
            // 清空当前值，避免内存泄漏
            if (qName.equals("v")) {
                currentValue.setLength(0);
            }
        }
    }
    
    /**
     * 根据列索引查找对应的表头名称
     * 优化性能，避免每次都使用stream操作
     */
    private String findHeaderByColumnIndex(int columnIndex) {
        for (Map.Entry<String, Integer> entry : headerMap.entrySet()) {
            if (entry.getValue() == columnIndex) {
                return entry.getKey();
            }
        }
        return "";
    }

    private void setFieldValue(Field field, T instance, String value) throws Exception {
        if (value == null) {
            value = "";
        }

        Class<?> fieldType = field.getType();
        try {
            if (String.class.equals(fieldType)) {
                field.set(instance, value);
            } else if (Integer.class.equals(fieldType) || int.class.equals(fieldType)) {
                field.set(instance, parseInteger(value, 0));
            } else if (Long.class.equals(fieldType) || long.class.equals(fieldType)) {
                field.set(instance, parseLong(value, 0L));
            } else if (Double.class.equals(fieldType) || double.class.equals(fieldType)) {
                field.set(instance, parseDouble(value, 0.0));
            } else if (Boolean.class.equals(fieldType) || boolean.class.equals(fieldType)) {
                field.set(instance, "1".equals(value) || "true".equalsIgnoreCase(value) || "yes".equalsIgnoreCase(value));
            } else if (Date.class.equals(fieldType)) {
                field.set(instance, parseDate(value));
            } else if (Timestamp.class.equals(fieldType)) {
                Date date = parseDate(value);
                field.set(instance, date != null ? new Timestamp(date.getTime()) : null);
            } else if (BigDecimal.class.equals(fieldType)) {
                field.set(instance, parseBigDecimal(value));
            } else {
                // 对于其他类型，尝试直接设置字符串值
                field.set(instance, value);
            }
        } catch (Exception e) {
            log.warn("设置字段 [{}] 值 [{}] 失败: {}", field.getName(), value, e.getMessage());
            // 对于原始类型，设置默认值避免NPE
            if (fieldType.isPrimitive()) {
                if (int.class.equals(fieldType)) {
                    field.set(instance, 0);
                } else if (long.class.equals(fieldType)) {
                    field.set(instance, 0L);
                } else if (double.class.equals(fieldType)) {
                    field.set(instance, 0.0);
                } else if (boolean.class.equals(fieldType)) {
                    field.set(instance, false);
                }
            }
        }
    }
    
    /**
     * 安全地解析Integer，失败时返回默认值
     */
    private Integer parseInteger(String value, Integer defaultValue) {
        try {
            return value.isEmpty() ? defaultValue : Integer.parseInt(value.trim());
        } catch (NumberFormatException e) {
            log.debug("解析整数失败: {}", value);
            return defaultValue;
        }
    }
    
    /**
     * 安全地解析Long，失败时返回默认值
     */
    private Long parseLong(String value, Long defaultValue) {
        try {
            return value.isEmpty() ? defaultValue : Long.parseLong(value.trim());
        } catch (NumberFormatException e) {
            log.debug("解析长整数失败: {}", value);
            return defaultValue;
        }
    }
    
    /**
     * 安全地解析Double，失败时返回默认值
     */
    private Double parseDouble(String value, Double defaultValue) {
        try {
            return value.isEmpty() ? defaultValue : Double.parseDouble(value.trim());
        } catch (NumberFormatException e) {
            log.debug("解析浮点数失败: {}", value);
            return defaultValue;
        }
    }
    
    /**
     * 尝试解析日期
     */
    private Date parseDate(String value) {
        if (value.isEmpty()) {
            return null;
        }
        
        // 尝试解析常见的日期格式
        String[] patterns = {"yyyy-MM-dd", "yyyy/MM/dd", "yyyy.MM.dd", 
                             "yyyy-MM-dd HH:mm:ss", "yyyy/MM/dd HH:mm:ss", 
                             "yyyy.MM.dd HH:mm:ss", "yyyyMMdd"};
        
        for (String pattern : patterns) {
            try {
                SimpleDateFormat sdf = new SimpleDateFormat(pattern);
                sdf.setLenient(false);
                return sdf.parse(value);
            } catch (Exception e) {
                // 尝试下一种格式
            }
        }
        
        log.debug("无法解析日期格式: {}", value);
        return null;
    }
    
    /**
     * 解析BigDecimal
     */
    private BigDecimal parseBigDecimal(String value) {
        try {
            return value.isEmpty() ? BigDecimal.ZERO : new BigDecimal(value.trim());
        } catch (NumberFormatException e) {
            log.debug("解析BigDecimal失败: {}", value);
            return BigDecimal.ZERO;
        }
    }
}
/**
 * 本类库是解决EXCEL与XML之间互相转换的问题
 * 开发此类库是为了解决EXCEL在subversion等版本管理软件中合并版本时的问题，因EXCEL是二进制文件，
 * 因此在合并项目时无法自动合并，因此建立中间件将excel转换成明文的文本模式，从而利于管理和维护。
 * 
 * 本项目是开源项目，可任意修改和使用
 * 
 * @author Yanbin Zhu<haker-haker@163.com>
 * @date 2012-08-24 10:28:24
 * @version 1.0.0
 */
package com.zhuyanbin.je2x;

import java.io.IOException;

import javax.xml.parsers.DocumentBuilder;
import javax.xml.parsers.DocumentBuilderFactory;
import javax.xml.parsers.ParserConfigurationException;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.w3c.dom.Document;
import org.w3c.dom.Node;
import org.w3c.dom.NodeList;
import org.xml.sax.SAXException;

/**
 * Xml读取器<br/>
 * 用于由Excel文件转换过来的数据,并转换成Excel格式的数据
 */
public class XmlReader
{
    /**
     * 需要转换的xml文件地址
     */
    private String _fileName = null;

    /**
     * work book
     */
    private HSSFWorkbook _wb       = null;

    /**
     * 默认构造函数
     */
    public XmlReader()
    {
    }

    /**
     * 构造函数
     * 
     * @param fileName
     *            要解析(转换为Excel)的xml文件
     */
    public XmlReader(String fileName)
    {
        setXmlFile(fileName);
    }

    /**
     * 设置转换成功之后的workbook对象
     * 
     * @param wb
     *            转换成功之后的workbook对象
     */
    private void setWorkBook(HSSFWorkbook wb)
    {
        _wb = wb;
    }

    /**
     * 获取转换成功之后的workbook对象
     * 
     * @return 转换成功之后的workbook对象
     */
    public HSSFWorkbook getWorkBook()
    {
        return _wb;
    }

    /**
     * 设置要转换的xml文件
     * 
     * @param fileName
     *            要转换的xml文件名
     */
    private void setXmlFile(String fileName)
    {
        _fileName = fileName;
    }

    /**
     * 获取要转换的xml文件
     * 
     * @return 要转换的xml文件
     */
    public String getXmlFile()
    {
        return _fileName;
    }

    /**
     * 解析实现方法
     * 
     * @throws ParserConfigurationException
     * @throws IOException
     * @throws SAXException
     */
    public void load() throws ParserConfigurationException, SecurityException, SAXException, IOException
    {
        setWorkBook(null);
        DocumentBuilder db = DocumentBuilderFactory.newInstance().newDocumentBuilder();
        Document doc = db.parse(getXmlFile());
        Node nwb = doc.getChildNodes().item(0);
        NodeList nl_sheets = nwb.getChildNodes();
        int len = nl_sheets.getLength();
        for (int i = 0; i < len; i++)
        {
            // System.out.println("name==" + nl_sheets.);
        }
        System.out.println("length==" + len);
    }

    /**
     * 解析实现方法
     * 
     * @param fileName
     *            要解析的xml文件名
     */
    public void load(String fileName) throws ParserConfigurationException, SecurityException, SAXException, IOException
    {
        setXmlFile(fileName);
        load();
    }

    /**
     * 要保存的Excel文件名
     * 
     * @param fileName
     *            Excel文件名
     * @return 成功返回true,失败返回false
     */
    public boolean save(String fileName)
    {
        return true;
    }
}

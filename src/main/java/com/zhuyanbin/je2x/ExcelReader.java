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

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.text.DecimalFormat;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;

/**
 * Excel读取器<br/>
 * 用于读取Excel文件中的数据,并转换成xml格式的数据
 */
public class ExcelReader
{
    /**
     * 需要处理的excel文件名
     */
    private String _fileName = null;

    private static DecimalFormat format    = null;
    /**
     * 转换成xml格式的字符串
     */
    private String               _xmlString = null;

    /**
     * 默认构造方法
     */
    public ExcelReader()
    {
    }

    /**
     * 构造方法
     * 
     * @param fileName
     *            需要处理的Excel路径(相对路径或绝对路径)
     */
    public ExcelReader(String fileName)
    {
        setFileName(fileName);
    }

    /**
     * 设置文件名
     * 
     * @param filename
     *            需要处理的Excel路径(相对路径或绝对路径)
     */
    private void setFileName(String filename)
    {
        _fileName = filename;
    }

    /**
     * 获取需要处理的Excel路径(相对路径或绝对路径)
     * 
     * @return 需要处理的Excel路径
     */
    public String getFileName()
    {
        return _fileName;
    }

    /**
     * 获取将excel文件转换好之后的xml数据
     * 
     * @return 转换好只好的xml数据
     */
    public String getXmlString()
    {
        return _xmlString;
    }

    /**
     * 将整个Excel文件转换的实现方法
     * 
     * @param wb
     *            excel文件操作对象
     * @return 转换之后的xml数据
     */
    private String parseWorkBook2XmlString(HSSFWorkbook wb)
    {
        String result = "<?xml version=\"1.0\" encoding=\"UTF-8\"?>\n";
        result += "<Workbook>\n";

        if (wb instanceof HSSFWorkbook)
        {
            int as_num = wb.getNumberOfSheets();
            for (int i = 0; i < as_num; i++)
            {
                result += parseSheet2XmlString(wb.getSheetAt(i));
            }
        }

        result += "</Workbook>\n";
        return result;
    }

    /**
     * 将Excel文件的工作表(sheet)转换的实现方法
     * 
     * @param sheet
     *            Excel文件的工作表(sheet)
     * @return 转换之后的xml数据
     */
    private String parseSheet2XmlString(HSSFSheet sheet)
    {
        String result = "";

        if (sheet instanceof HSSFSheet)
        {
            result = "<worksheet name=\"" + sheet.getSheetName() + "\">\n";
            int rowCount = sheet.getLastRowNum();
            for (int i = 0; i < rowCount; i++)
            {
                result += parseRow2XmlString(sheet.getRow(i));
            }

            result += "</worksheet>\n";
        }

        return result;
    }

    /**
     * 将Excel文件的工作表(sheet)中的某行(row)转换的实现方法
     * 
     * @param row
     *            Excel文件的工作表(sheet)中的某行(row)
     * @return 转换之后的xml数据
     */
    private String parseRow2XmlString(HSSFRow row)
    {
        String result = "";
        if (row instanceof HSSFRow)
        {
            result = "    <row Height=\"" + row.getHeight() + "\">\n";
            int cells = row.getLastCellNum();
            for (int i = 0; i < cells; i++)
            {
                result += parseCell2XmlString(row.getCell(i));
            }

            result += "    </row>\n";
        }
        return result;
    }

    /**
     * 将Excel文件的工作表(sheet)中的某行(row)中的某个单元格(cell)转换的实现方法
     * 
     * @param cell
     *            Excel文件的工作表(sheet)中的某行(row)中的某个单元格(cell)
     * @return 转换之后的xml数据
     */
    private String parseCell2XmlString(HSSFCell cell)
    {
        String result = "";
        if (cell instanceof HSSFCell)
        {
            result = "        <cell >" + getCellValue(cell) + "</cell>\n";
        }

        return result;
    }

    /**
     * 设置转换之后的xml数据
     * 
     * @param xml
     *            xml数据
     */
    private void setXmlString(String xml)
    {
        _xmlString = xml;
    }

    /**
     * 解析Excel文件，并转换成xml数据 <br/>
     * 如果文件不存在或无法打开Excel文件则抛出异常
     * 
     * @throws FileNotFoundException
     * @throws IOException
     */
    public void load() throws FileNotFoundException, IOException
    {
        setXmlString(null);
        FileInputStream fis = new FileInputStream(getFileName());

        try
        {
            POIFSFileSystem fs = new POIFSFileSystem(fis);
            setXmlString(parseWorkBook2XmlString(new HSSFWorkbook(fs)));
        }
        catch (IOException ex)
        {
            if (null != fis)
            {
                fis.close();
                fis = null;
            }

            throw ex;
        }
        
        if (null != fis)
        {
            fis.close();
            fis = null;
        }
    }
    
    /**
     * 解析Excel文件，并转换成xml数据 <br/>
     * 如果文件不存在或无法打开Excel文件则抛出异常
     * 
     * @param fileName
     *            需要解析的excel文件
     * @throws FileNotFoundException
     * @throws IOException
     */
    public void load(String fileName) throws FileNotFoundException, IOException
    {
        setFileName(fileName);
        load();
    }

    /**
     * 数值格式化器
     * 
     * @return
     */
    private static DecimalFormat getFormater()
    {
        if (null == format)
        {
            format = new DecimalFormat();
            format.setDecimalSeparatorAlwaysShown(false);
            format.setGroupingUsed(false);
        }

        return format;
    }

    /**
     * 获取Excel中单元格(cell)数据
     * 
     * @param cell
     *            Excel中的单元格
     * @return Excel中单元格(cell)数据
     */
    private String getCellValue(HSSFCell cell)
    {
        String result = "";
        if (cell instanceof HSSFCell)
        {
            switch (cell.getCellType())
            {
                case HSSFCell.CELL_TYPE_NUMERIC:
                    result = "<data type=\"Number\">" + getFormater().format(cell.getNumericCellValue()) + "</data>";
                    break;
                case HSSFCell.CELL_TYPE_BOOLEAN:
                    result = "<data type=\"Boolean\">" + cell.toString() + "</data>";
                    break;
                case HSSFCell.CELL_TYPE_STRING:
                default:
                    result = "<data type=\"String\">" + cell.toString() + "</data>";
            }
        }

        return result;
    }
}

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

/**
 * XML 标签
 */
final public class XmlType
{
    /**
     * Excel中的Work Book
     */
    public final static String WorkBook = "WorkBook";
    /**
     * Excel中的Work Sheet
     */
    public final static String WorkSheet = "WorkSheet";
    /**
     * Excel中的Row
     */
    public final static String Row       = "Row";
    /**
     * Excel中的Cell
     */
    public final static String Cell      = "Cell";

    /**
     * Excel中的Cell Data
     */
    public final static String Data      = "Data";
}

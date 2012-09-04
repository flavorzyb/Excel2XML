package com.zhuyanbin.je2x;

import java.io.FileNotFoundException;
import java.io.IOException;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.easymock.EasyMock;
import org.jdom2.Element;
import org.junit.After;
import org.junit.AfterClass;
import org.junit.Assert;
import org.junit.Before;
import org.junit.BeforeClass;
import org.junit.Test;

public class ExcelReaderTest
{
    private final String fileName    = "src/test/excel/test.xls";
    private final String xmlFileName = "src/test/xml/test.xml";

    private ExcelReader    classRelection;
    @BeforeClass
    public static void setUpBeforeClass() throws Exception
    {
    }

    @AfterClass
    public static void tearDownAfterClass() throws Exception
    {
    }

    @Before
    public void setUp() throws Exception
    {
        classRelection = new ExcelReader();
    }

    @After
    public void tearDown() throws Exception
    {
        classRelection = null;
    }

    @Test
    public void testGetFileName()
    {
        Assert.assertSame(null, classRelection.getFileName());
        classRelection = null;
        classRelection = new ExcelReader(fileName);
        Assert.assertSame(fileName, classRelection.getFileName());
    }

    @Test
    public void testLoadSuccess() throws FileNotFoundException, IOException
    {
        Assert.assertFalse(classRelection.output(xmlFileName));
        classRelection.load(fileName);
        Assert.assertTrue(classRelection.getXml() instanceof Element);
        Assert.assertTrue(classRelection.output(xmlFileName));
    }
    
    @Test
    public void testLoadWithMock() throws FileNotFoundException, IOException
    {
        // mock parseSheet2Xml
        classRelection = EasyMock.createMockBuilder(ExcelReader.class).addMockedMethod("parseSheet2Xml", HSSFSheet.class).createMock();
        EasyMock.expect(classRelection.parseSheet2Xml(EasyMock.anyObject(HSSFSheet.class))).andReturn(null).anyTimes();
        EasyMock.replay(classRelection);

        classRelection.load(fileName);
        EasyMock.verify(classRelection);

        // mock parseRow2Xml
        classRelection = EasyMock.createMockBuilder(ExcelReader.class).addMockedMethod("parseRow2Xml", HSSFRow.class).createMock();
        EasyMock.expect(classRelection.parseRow2Xml(EasyMock.anyObject(HSSFRow.class))).andReturn(null).anyTimes();
        EasyMock.replay(classRelection);

        classRelection.load(fileName);
        EasyMock.verify(classRelection);

        // mock parseCell2Xml
        classRelection = EasyMock.createMockBuilder(ExcelReader.class).addMockedMethod("parseCell2Xml", HSSFCell.class).createMock();
        EasyMock.expect(classRelection.parseCell2Xml(EasyMock.anyObject(HSSFCell.class))).andReturn(null).anyTimes();
        EasyMock.replay(classRelection);

        classRelection.load(fileName);
        EasyMock.verify(classRelection);

        // mock getCellValue
        classRelection = EasyMock.createMockBuilder(ExcelReader.class).addMockedMethod("getCellValue", HSSFCell.class).createMock();
        EasyMock.expect(classRelection.getCellValue(EasyMock.anyObject(HSSFCell.class))).andReturn(null).anyTimes();
        EasyMock.replay(classRelection);

        classRelection.load(fileName);
        EasyMock.verify(classRelection);
    }
    
    @Test
    public void testLoadWithCellMock() throws FileNotFoundException, IOException
    {
        // mock getCellValue
        classRelection = EasyMock.createMockBuilder(ExcelReader.class).addMockedMethod("getCellValue", HSSFCell.class).createMock();
        EasyMock.expect(classRelection.getCellValue(EasyMock.anyObject(HSSFCell.class))).andReturn(null).anyTimes();
        EasyMock.replay(classRelection);

        classRelection.load(fileName);
        EasyMock.verify(classRelection);
    }
}

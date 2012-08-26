package com.zhuyanbin.je2x;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import org.junit.After;
import org.junit.AfterClass;
import org.junit.Assert;
import org.junit.Before;
import org.junit.BeforeClass;
import org.junit.Test;

public class ExcelReaderTest
{
    private final String fileName = "src/test/excel/test.xls";

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
        classRelection.load(fileName);
        Assert.assertTrue(classRelection.getXmlString().length() > 0);
        FileOutputStream fos = new FileOutputStream("test.xml");
        fos.write(classRelection.getXmlString().getBytes());
        fos.flush();
        fos.close();
    }
}

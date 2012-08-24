package com.zhuyanbin.je2x;

import org.junit.After;
import org.junit.AfterClass;
import org.junit.Assert;
import org.junit.Before;
import org.junit.BeforeClass;
import org.junit.Test;

import com.zhuyanbin.je2x.Excel2Xml;

public class Excel2XmlTest
{
    private final String fileName = "test.xls";
    private Excel2Xml    classRelection;
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
        classRelection = new Excel2Xml(fileName);
    }

    @After
    public void tearDown() throws Exception
    {
        classRelection = null;
    }

    @Test
    public void testGetFileName()
    {
        Assert.assertSame(fileName, classRelection.getFileName());
        classRelection = new Excel2Xml(null);
        Assert.assertSame(null, classRelection.getFileName());
    }
}

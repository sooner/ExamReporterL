﻿using Microsoft.Office.Interop.Excel;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;

namespace statistic_test
{
    
    
    /// <summary>
    ///This is a test class for RangeTest and is intended
    ///to contain all RangeTest Unit Tests
    ///</summary>
    [TestClass()]
    public class RangeTest
    {


        private TestContext testContextInstance;

        /// <summary>
        ///Gets or sets the test context which provides
        ///information about and functionality for the current test run.
        ///</summary>
        public TestContext TestContext
        {
            get
            {
                return testContextInstance;
            }
            set
            {
                testContextInstance = value;
            }
        }

        #region Additional test attributes
        // 
        //You can use the following additional attributes as you write your tests:
        //
        //Use ClassInitialize to run code before running the first test in the class
        //[ClassInitialize()]
        //public static void MyClassInitialize(TestContext testContext)
        //{
        //}
        //
        //Use ClassCleanup to run code after all tests in a class have run
        //[ClassCleanup()]
        //public static void MyClassCleanup()
        //{
        //}
        //
        //Use TestInitialize to run code before running each test
        //[TestInitialize()]
        //public void MyTestInitialize()
        //{
        //}
        //
        //Use TestCleanup to run code after each test has run
        //[TestCleanup()]
        //public void MyTestCleanup()
        //{
        //}
        //
        #endregion


        internal virtual Range CreateRange()
        {
            // TODO: Instantiate an appropriate concrete class.
            Range target = null;
            return target;
        }

        /// <summary>
        ///A test for _Default
        ///</summary>
        [TestMethod()]
        public void _DefaultTest()
        {
            Range target = CreateRange(); // TODO: Initialize to an appropriate value
            object RowIndex = null; // TODO: Initialize to an appropriate value
            object ColumnIndex = null; // TODO: Initialize to an appropriate value
            object expected = null; // TODO: Initialize to an appropriate value
            object actual;
            target[RowIndex, ColumnIndex] = expected;
            actual = target[RowIndex, ColumnIndex];
            Assert.AreEqual(expected, actual);
            Assert.Inconclusive("Verify the correctness of this test method.");
        }

        /// <summary>
        ///A test for Value2
        ///</summary>
        [TestMethod()]
        public void Value2Test()
        {
            Range target = CreateRange(); // TODO: Initialize to an appropriate value
            object expected = null; // TODO: Initialize to an appropriate value
            object actual;
            target.Value2 = expected;
            actual = target.Value2;
            Assert.AreEqual(expected, actual);
            Assert.Inconclusive("Verify the correctness of this test method.");
        }

        /// <summary>
        ///A test for Text
        ///</summary>
        [TestMethod()]
        public void TextTest()
        {
            Range target = CreateRange(); // TODO: Initialize to an appropriate value
            object actual;
            actual = target.Text;
            Assert.Inconclusive("Verify the correctness of this test method.");
        }

        /// <summary>
        ///A test for Rows
        ///</summary>
        [TestMethod()]
        public void RowsTest()
        {
            Range target = CreateRange(); // TODO: Initialize to an appropriate value
            Range actual;
            actual = target.Rows;
            Assert.Inconclusive("Verify the correctness of this test method.");
        }

        /// <summary>
        ///A test for Count
        ///</summary>
        [TestMethod()]
        public void CountTest()
        {
            Range target = CreateRange(); // TODO: Initialize to an appropriate value
            int actual;
            actual = target.Count;
            Assert.Inconclusive("Verify the correctness of this test method.");
        }
    }
}

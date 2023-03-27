using Microsoft.VisualStudio.TestTools.UnitTesting;
using Potolki;
using System;

namespace UnitTest
{
    [TestClass]
    public class UnitTest
    {
        [TestMethod]
        public void TestMethod_calculation_meter1()
        {
            string width = "0";
            string height = "0";
            string expected = "null";
            string actual = Potolki.Base_form.calculation_meter(width, height);
            Assert.AreEqual(expected, actual);
        }
        [TestMethod]
        public void TestMethod_calculation_meter2()
        {
            string width = "";
            string height = "";
            string expected = "error";
            string actual = Potolki.Base_form.calculation_meter(width, height);
            Assert.AreEqual(expected, actual);
        }
        [TestMethod]
        public void TestMethod_calculation_meter3()
        {
            string width = "-10";
            string height = "-5";
            string expected = "null";
            string actual = Potolki.Base_form.calculation_meter(width, height);
            Assert.AreEqual(expected, actual);
        }
    }
}

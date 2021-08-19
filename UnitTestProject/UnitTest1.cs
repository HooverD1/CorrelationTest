using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace CorrelationStringTests
{
    [TestClass]
    public class ConstructFromStringValue
    {
        [TestMethod]
        public void TestMethod1_Fail()
        {
            try
            {
                var cs = CorrelationTest.Data.CorrelationString.ConstructFromStringValue("");
            }
            catch(System.FormatException)
            {
                Assert.IsTrue(true);
                return;
            }
            Assert.IsTrue(false);
        }
        [TestMethod]
        public void TestMethod2_Fail()
        {
            try
            {
                var cs = CorrelationTest.Data.CorrelationString.ConstructFromStringValue("a,a&");
            }
            catch (System.FormatException)
            {
                Assert.IsTrue(true);
                return;
            }
            Assert.IsTrue(false);
        }
        [TestMethod]
        public void TestMethod3_Fail()
        {
            try
            {
                var cs = CorrelationTest.Data.CorrelationString.ConstructFromStringValue("a,CM,a&");
            }
            catch (System.FormatException)
            {
                Assert.IsTrue(true);
                return;
            }
            Assert.IsTrue(false);
        }
        [TestMethod]
        public void TestMethod4_Fail()
        {
            try
            {
                var cs = CorrelationTest.Data.CorrelationString.ConstructFromStringValue("a,CP,a&");
            }
            catch (System.FormatException)
            {
                Assert.IsTrue(true);
                return;
            }
            Assert.IsTrue(false);
        }
        [TestMethod]
        public void TestMethod5_Pass()
        {
            try
            {
                    var cs = CorrelationTest.Data.CorrelationString.ConstructFromStringValue(
                        "2,CP," +
                        "DH|E|David Hoover|19072119:23:30.302," +
                        "DH|E|David Hoover|19072119:23:30.303," +
                        "DH|E|David Hoover|19072119:23:30.304&" +
                        "0,0");
            }
            catch (System.FormatException)
            {
                Assert.IsTrue(false);
                return;
            }
            catch(System.Exception)
            {
                Assert.IsTrue(false);
                return;
            }
            Assert.IsTrue(true);
        }

        [TestMethod]
        public void TestMethod6_Pass()
        {
            try
            {
                var cs = CorrelationTest.Data.CorrelationString.ConstructFromStringValue(
                    "2,CM," +
                    "DH|E|David Hoover|19072119:23:30.302," +
                    "DH|E|David Hoover|19072119:23:30.303," +
                    "DH|E|David Hoover|19072119:23:30.304&" +
                    "0");
            }
            catch (System.FormatException)
            {
                Assert.IsTrue(false);
                return;
            }
            catch (System.Exception)
            {
                Assert.IsTrue(false);
                return;
            }
            Assert.IsTrue(true);
        }

        [TestMethod]
        public void TestMethod7_Pass()
        {
            try
            {
                var cs = CorrelationTest.Data.CorrelationString.ConstructFromStringValue(
                    "2,DM," +
                    "DH|E|David Hoover|19072119:23:30.302," +
                    "DH|E|David Hoover|19072119:23:30.303," +
                    "DH|E|David Hoover|19072119:23:30.304&" +
                    "0");
            }
            catch (System.FormatException)
            {
                Assert.IsTrue(false);
                return;
            }
            catch (System.Exception)
            {
                Assert.IsTrue(false);
                return;
            }
            Assert.IsTrue(true);
        }

        [TestMethod]
        public void TestMethod8_Pass()
        {
            try
            {
                var cs = CorrelationTest.Data.CorrelationString.ConstructFromStringValue(
                    "2,DP," +
                    "DH|E|David Hoover|19072119:23:30.302," +
                    "DH|E|David Hoover|19072119:23:30.303," +
                    "DH|E|David Hoover|19072119:23:30.304&" +
                    "0,0");
            }
            catch (System.FormatException)
            {
                Assert.IsTrue(false);
                return;
            }
            catch (System.Exception)
            {
                Assert.IsTrue(false);
                return;
            }
            Assert.IsTrue(true);
        }

        [TestMethod]
        public void TestMethod9_Pass()
        {
            try
            {
                var cs = CorrelationTest.Data.CorrelationString.ConstructFromStringValue(
                    "2,PP," +
                    "DH|E|David Hoover|19072119:23:30.302," +
                    "DH|E|David Hoover|19072119:23:30.303," +
                    "DH|E|David Hoover|19072119:23:30.304&" +
                    "0,0");
            }
            catch (System.FormatException)
            {
                Assert.IsTrue(false);
                return;
            }
            catch (System.Exception)
            {
                Assert.IsTrue(false);
                return;
            }
            Assert.IsTrue(true);
        }

        [TestMethod]
        public void TestMethod10_Fail()
        {
            try
            {
                var cs = CorrelationTest.Data.CorrelationString.ConstructFromStringValue(
                    "1,CP," +
                    "DH|E|David Hoover|19072119:23:30.302," +
                    "DH|E|David Hoover|19072119:23:30.303," +
                    "DH|E|David Hoover|19072119:23:30.304&" +
                    "0,0");
            }
            catch (System.FormatException)
            {
                Assert.IsTrue(true);
                return;
            }
            catch (System.Exception)
            {
                Assert.IsTrue(false);
                return;
            }
            Assert.IsTrue(false);
        }

        [TestMethod]
        public void TestMethod11_Fail()
        {
            try
            {
                var cs = CorrelationTest.Data.CorrelationString.ConstructFromStringValue(
                    "1,CM," +
                    "DH|E|David Hoover|19072119:23:30.302," +
                    "DH|E|David Hoover|19072119:23:30.303," +
                    "DH|E|David Hoover|19072119:23:30.304&" +
                    "0");
            }
            catch (System.FormatException)
            {
                Assert.IsTrue(true);
                return;
            }
            catch (System.Exception)
            {
                Assert.IsTrue(false);
                return;
            }
            Assert.IsTrue(false);
        }
        [TestMethod]
        public void TestMethod12_Fail()
        {
            try
            {
                var cs = CorrelationTest.Data.CorrelationString.ConstructFromStringValue(
                    "3,CM," +
                    "DH|E|David Hoover|19072119:23:30.302," +
                    "DH|E|David Hoover|19072119:23:30.303," +
                    "DH|E|David Hoover|19072119:23:30.304&" +
                    "0,0&0");
            }
            catch (System.FormatException)
            {
                Assert.IsTrue(true);
                return;
            }
            catch (System.Exception)
            {
                Assert.IsTrue(false);
                return;
            }
            Assert.IsTrue(false);
        }
    }
}

using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Data;
using Shoe_Organizer__Excel_;

namespace OrganizerTest
{
    [TestClass]
    public class OrganizerTests
    {
        [TestMethod]
        public void TestShoeSum_2and100_2and150_500current()
        {
            decimal x = 1000;
            decimal y = 10;
            decimal expected = 1100;
            Procents c = new Procents();
            decimal act = c.Prc(x, y);
            Assert.AreEqual(expected, act);
        }
    }
}

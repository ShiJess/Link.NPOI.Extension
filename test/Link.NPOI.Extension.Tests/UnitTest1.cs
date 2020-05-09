using System;
using Xunit;

namespace Link.NPOI.Extension.Tests
{
    public class UnitTest1
    {
        [Fact]
        public void Test1()
        {
            Class1 a = new Class1();
            a.a();
         
            Assert.NotNull(a);
        }
    }
}

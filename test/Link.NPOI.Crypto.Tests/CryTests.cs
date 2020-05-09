using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Xunit;

namespace OfficeOpenXmlCrypto.Test
{
    public class CryTests
    {
        [Fact]
        public void EncryptedFile()
        {
            using (OfficeCryptoStream s = OfficeCryptoStream.Open("test.xlsx"))
            {
                s.Password = "123";
                s.Save();
            }
        }
    }
}

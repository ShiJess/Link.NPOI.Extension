using System;
using System.Collections.Generic;
using System.Text;
using System.IO;
using Xunit;

namespace OfficeOpenXmlCrypto.Test
{
    public class OfficeCryptoStreamTest
    {
        String TestFile = "test.xlsx";
               
        [Fact]
        public void PlaintextFile()
        {
            CreateTestWorkbook(null);
            AssertFileCorrect(null);
        }

        [Fact]
        public void EncryptedFile()
        {
            CreateTestWorkbook("foo");
            //AssertFileCorrect("foo");
        }

        [Fact]
        // This works -- if file is saved later, 
        // we encrypt it and apply the password.
        public void PlaintextFileUsingPassword()
        {
            CreateTestWorkbook(null);
            AssertFileCorrect("bar");
        }

        [Fact]
        //[ExpectedException(typeof(InvalidPasswordException))]
        public void EncryptedFileEmptyPassword()
        {
            CreateTestWorkbook("foo");
            AssertFileCorrect(null);
        }

        [Fact]
        public void IsFilePlaintext()
        {
            CreateTestWorkbook(null);
            //Assert.Equal(true, OfficeCryptoStream.IsPlaintext(TestFile), "IsPlaintext");
            //Assert.Equal(false, OfficeCryptoStream.IsEncrypted(TestFile), "IsEncrypted");
        }

        [Fact]
        public void IsFileEncrypted()
        {
            CreateTestWorkbook("foo");
            //Assert.Equal(false, OfficeCryptoStream.IsPlaintext(TestFile), "IsPlaintext");
            //Assert.Equal(true, OfficeCryptoStream.IsEncrypted(TestFile), "IsEncrypted");
        }

        [Fact]
        public void IsFileNotOfficePackage()
        {
            File.WriteAllText(TestFile, "Not an office package");
            //Assert.Equal(false, OfficeCryptoStream.IsPlaintext(TestFile), "IsPlaintext");
            //Assert.Equal(false, OfficeCryptoStream.IsEncrypted(TestFile), "IsEncrypted");
        }

        [Fact]
        public void TestFormatEncrypted()
        {
            CreateTestWorkbook("foo");
            //Assert.Equal(false, OfficeCryptoStream.IsPlaintext(TestFile), "IsPlaintext");
            //Assert.Equal(true, OfficeCryptoStream.IsEncrypted(TestFile), "IsEncrypted");
        }

        [Fact]
        public void TryOpen()
        {
            CreateTestWorkbook("foo");
            //OfficeCryptoStream ocs;
            //Assert.IsFalse(OfficeCryptoStream.TryOpen(TestFile, null, out ocs));
            //Assert.IsFalse(OfficeCryptoStream.TryOpen(TestFile, "bar", out ocs));
            //Assert.IsTrue(OfficeCryptoStream.TryOpen(TestFile, "foo", out ocs));
            //ocs.Close();
        }

        [Fact]
        public void EncryptedFileInvalidPassword()
        {
            CreateTestWorkbook("foo");
            AssertFileCorrect("bar");
        }

        [Fact]
        public void PasswordChange()
        {
            CreateTestWorkbook("foo");
            ChangePassword("foo", "bar");
            AssertFileCorrect("bar");
        }

        [Fact]
        public void PasswordAddToPlaintext()
        {
            CreateTestWorkbook(null);
            ChangePassword(null, "bar");
            AssertFileCorrect("bar");
        }

        [Fact]
        public void PasswordRemoveFromEncrypted()
        {
            CreateTestWorkbook("foo");
            ChangePassword("foo", null);
            AssertFileCorrect(null);
        }

        void AssertFileCorrect(String password)
        {
            using (OfficeCryptoStream s = OfficeCryptoStream.Open(TestFile, password))
            {
                //using (ExcelPackage p = new ExcelPackage(s))
                //{
                //    Assert.IsNotNull(p, "Cannot create package.");
                //    ExcelWorksheet ws = p.Workbook.Worksheets["Test"];
                //    Assert.IsNotNull(ws, "No Test worksheet.");
                //    String cval = ws.Cell(1, 1).Value;
                //    Assert.AreEqual("Test Cell", cval, "First cell value incorrect.");
                //}
            }
        }

        void ChangePassword(String oldPassword, String newPassword)
        {
            using (OfficeCryptoStream s = OfficeCryptoStream.Open(TestFile, oldPassword))
            {
                s.Password = newPassword;
                s.Save();
            }
        }

        void CreateTestWorkbook(String password)
        {
            using (OfficeCryptoStream s = OfficeCryptoStream.Open(TestFile))
            {
                s.Password = password;
               
                s.Save();
            }
        }

    }
}

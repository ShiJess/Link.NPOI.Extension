using System;
using System.Collections.Generic;
using System.Text;
using System.IO;

namespace OfficeOpenXmlCrypto.Test
{
    public static class SampleCode
    {

        public static void AccessEncryptedFile()
        {
            using (OfficeCryptoStream stream = OfficeCryptoStream.Open("a.xlsx", "password"))
            {
                DoStuff(stream);
                stream.Save(); // Skip this line if you don't want to save/encrypt
            }
        }

        public static void AccessEncryptedFileManualSave()
        {
            // Create stream, decrypt
            OfficeCryptoStream stream = OfficeCryptoStream.Open("a.xlsx", "password");

            // Do whatever is needed in your program 
            DoStuff(stream);

            // When done, save and close the encrypted stream
            stream.Save();
            stream.Close();
        }

        public static void AccessPlaintextFile()
        {
            using (OfficeCryptoStream stream = OfficeCryptoStream.Open("a.xlsx"))
            {
                DoStuff(stream);
                stream.Save();
            }
        }

        public static void CreateEncryptedFile()
        {
            using (OfficeCryptoStream stream = OfficeCryptoStream.Create("a.xlsx"))
            {
                DoStuff(stream);

                // Set or change the password anytime before the save. 
                // Set to null to save as plaintext.
                stream.Password = "password";

                stream.Save();
            }
        }

        public static OfficeCryptoStream OpenPasswordProtectedFile(String file)
        {
            String password = null;
            OfficeCryptoStream stream = null;
            while (OfficeCryptoStream.TryOpen(file, password, out stream) == false)
            {
                // Replace with your own input method (e.g. a dialog box)
                Console.Write("Enter password: ");
                password = Console.ReadLine();
            }
            return stream;
        }

        public static void DoStuff(Stream stream)
        {
            /*
            // Create the package based on the stream
            ExcelPackage package = new ExcelPackage(stream);

            // Do whatever is needed in your program (read/write/change)...

            // When you're done, save and close the package
            package.Save();
            package.Dispose();
            */
        }

    }
}

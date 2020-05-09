using System;
using System.Collections.Generic;
using System.Text;
using System.IO;

/*
 * Wrapper coded by Danilo Mirkovic, Oct 2009
 * License: Open source, GPL 
 *  
 * Note: 
 * - OfficeCrypto class is LGPL2/Apache license.
 *   http://www.lyquidity.com/devblog/?p=35
 * - NPOI is Apache 2.0 license
 *   http://npoi.codeplex.com/
 */
namespace OfficeOpenXmlCrypto
{
    /// <summary>
    /// Provides Office 2007 encryption service.
    /// Usage: pass as a plain-text stream to Package constructor.
    /// </summary>
    public class OfficeCryptoStream : MemoryStream
    {
        static readonly byte[] HeaderEncrypted = 
            new byte[] { 0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1 };
        static readonly byte[] 
            HeaderPlaintext = new byte[] { 0x50, 0x4B, 0x03, 0x04 };

        String _password = null;

        // Encrypted or plaintext stream (of the underlying storage file)
        Stream _storage;

        #region Creator methods (helper)

        /// <summary>
        /// Create a new file and stream based on it.
        /// Set the Password field to encrypt it.
        /// </summary>
        /// <param name="newFile">Name of the new file.</param>
        /// <returns>Empty stream for writing to the file.</returns>
        public static OfficeCryptoStream Create(String newFile)
        {
            return new OfficeCryptoStream(newFile, FileMode.Create, null);
        }

        /// <summary>
        /// Open an existing plaintext file.
        /// </summary>
        /// <param name="file"></param>
        /// <returns>Stream for accessing the file</returns>
        /// <throws>InvalidPasswordException if file is encrypted</throws>
        /// <throws>FileNotFoundException if file does not exist</throws>
        /// <throws>Other exceptions (FileFormatException, IOException)</throws>
        public static OfficeCryptoStream Open(String file)
        {
            return Open(file, null);
        }

        /// <summary>
        /// Open an existing encrypted file. 
        /// </summary>
        /// <param name="file">Path to the existing file</param>
        /// <param name="password">Password to decrypt the file</param>
        /// <returns>Stream for accessing the file</returns>
        /// <throws>InvalidPasswordException if password is incorrect</throws>
        /// <throws>FileNotFoundException if file does not exist</throws>
        /// <throws>Other exceptions (FileFormatException, IOException)</throws>
        public static OfficeCryptoStream Open(String file, String password)
        {
            if (!File.Exists(file))
            {
                throw new FileNotFoundException("", file);
            }
            return new OfficeCryptoStream(file, FileMode.Open, password);
        }

        /// <summary>
        /// Try to open an encrypted, or plaintext file.
        /// </summary>
        /// <param name="file"></param>
        /// <param name="password"></param>
        /// <param name="stream"></param>
        /// <returns>True if file was opened, false otherwise.</returns>
        public static bool TryOpen(String file, String password, out OfficeCryptoStream stream)
        {
            stream = null;
            try
            {
                stream = Open(file, password);

            }
            catch (InvalidPasswordException)
            {
                return false;
            }
            return true;
        }

        #endregion

        private OfficeCryptoStream(String file, FileMode mode, String password)
            : this (new FileStream(file, mode), password) { }

        /// <summary>
        /// Create based on a stream.
        /// </summary>
        /// <param name="storageStream">Storage stream, usually FileStream</param>
        /// <param name="password">Password. Pass null for plaintext.</param>
        /// <throws>InvalidPasswordException if password is incorrect</throws>
        /// <throws>FileNotFoundException if file does not exist</throws>
        /// <throws>FileFormatException: file is in the wrong format</throws>
        public OfficeCryptoStream(Stream storageStream, String password)
        {
            _storage = storageStream;
            Password = password;

            if (storageStream.Length == 0) 
            {
                // No need to decrypt, stream is already 0-length 
                return;
            }

            // Check if the file is actually encrypted or plaintext
            bool isPlain = IsPlaintext(storageStream);
            bool isEncrypted = IsEncrypted(storageStream);
            if (!isPlain && !isEncrypted)
            {
                Close(); // In ctor, cannot rely on client/using to close
                throw new FileFormatException("File is neither plaintext package nor Office 2007 encrypted.");
            }

            // Read the file
            byte[] contents = new byte[storageStream.Length];
            storageStream.Read(contents, 0, contents.Length);

            // Decrypt if needed
            if (isEncrypted)
            {
                if (String.IsNullOrEmpty(Password))
                {
                    Close(); // In ctor, cannot rely on client/using to close
                    throw new InvalidPasswordException("Password not provided.");
                }

                try
                {
                    OfficeCrypto oc = new OfficeCrypto();
                    contents = oc.DecryptToArray(contents, password);
                }
                catch (Exception)
                {
                    Close(); // In ctor, cannot rely on client/using to close
                    throw;
                }
            }

            // Write out to the memory stream
            base.Write(contents, 0, contents.Length);
            base.Flush();
            base.Position = 0;
        }

        #region File format checks

        /// <summary>
        /// Checks if given file is a plaintext Office 2007 package.
        /// </summary>
        /// <param name="file"></param>
        /// <returns></returns>
        public static bool IsPlaintext(String file)
        {
            using (FileStream fs = new FileStream(file, FileMode.Open))
            {
                return IsPlaintext(fs);
            }
        }

        /// <summary>
        /// Checks if given stream is a plaintext Office 2007 package.
        /// </summary>
        /// <param name="file"></param>
        /// <returns></returns>
        public static bool IsPlaintext(Stream s)
        {
            return ContainsHeader(s, HeaderPlaintext);
        }

        /// <summary>
        /// Checks if given file is an encrypted Office 2007 package.
        /// </summary>
        /// <param name="file"></param>
        /// <returns></returns>
        public static bool IsEncrypted(String file)
        {
            using (FileStream fs = new FileStream(file, FileMode.Open))
            {
                return IsEncrypted(fs);
            }
        }

        /// <summary>
        /// Checks if given stream is an encrypted Office 2007 package.
        /// </summary>
        /// <param name="file"></param>
        /// <returns></returns>
        public static bool IsEncrypted(Stream s)
        {
            return ContainsHeader(s, HeaderEncrypted);
        }

        /// <summary>
        /// Checks the header without messing up the stream
        /// </summary>
        /// <param name="s"></param>
        /// <param name="header"></param>
        /// <returns></returns>
        static bool ContainsHeader(Stream s, byte[] header)
        {
            long pos = s.Position;
            try
            {
                foreach (byte hb in header)
                {
                    if (s.ReadByte() != (int)hb) { return false; }
                }
            }
            finally
            {
                s.Position = pos;
            }
            return true;
        }

        #endregion

        /// <summary>
        /// True if stream is encrypted (has a password), false otherwise.
        /// </summary>
        public bool Encrypted
        {
            get { return !String.IsNullOrEmpty(_password); } 
        }

        /// <summary>
        /// Gets or sets the password. Set to null for plaintext (no encryption).
        /// Throws InvalidOperationException if stream is read-only or does 
        /// not support seeking.
        /// </summary>
        public String Password
        {
            get { return _password; }
            set 
            {
                // Throw exception if closed
                if (!base.CanWrite || !base.CanSeek)
                {
                    throw new InvalidOperationException("Cannot set password. Underlying stream does not support seek or write. Make sure it was not closed.");
                }
                _password = value; 
            }
        }

        /// <summary>
        /// Encrypt and write out to a new file stream.
        /// NOTE: Don't forget to call Close().
        /// If file exists, this overwrites it -- check before calling.
        /// </summary>
        /// <param name="filename"></param>
        public void SaveAs(String filename)
        {
            _storage.Close();
            _storage = new FileStream(filename, FileMode.Create);
            Save();
        }

        /// <summary>
        /// Encrypt and write out to storage stream.
        /// NOTE: Don't forget to call Close()
        /// </summary>
        public void Save()
        {
            _storage.Seek(0, SeekOrigin.Begin);
            _storage.SetLength(0);
            _storage.Position = 0;

            if (Encrypted)
            {
                // Encrypt this to the storage stream
                OfficeCrypto oc = new OfficeCrypto();
                oc.EncryptToStream(base.ToArray(), Password, _storage);
            }
            else
            {
                // Just write the contents to storage stream
                base.WriteTo(_storage);
            }
        }

        /// <summary>
        /// Close the stream and save the file.
        /// NOTE: Call Save() to encrypt the storage stream. 
        /// </summary>
        public override void Close()
        {
            base.Close();
            _storage.Close();
        }
    }
}

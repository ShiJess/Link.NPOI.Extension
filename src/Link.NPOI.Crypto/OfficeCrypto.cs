/******************************************************************************* 
 *  _                      _     _ _ _         
 * | |   _   _  __ _ _   _(_) __| (_) |_ _   _ 
 * | |  | | | |/ _` | | | | |/ _` | | __| | | |
 * | |__| |_| | (_| | |_| | | (_| | | |_| |_| |
 * |_____\__, |\__, |\__,_|_|\__,_|_|\__|\__, |
 *       |___/    |_|                    |___/ 
 * 
 *  Decrypytion library for Office Open XML files
 *  API Version: 2008-12-12
 *  Generated: Wed Dec 12 12:33:00 GMT 2008 
 * 
 * ***************************************************************************** 
 *  Copyright Lyquidity Solutions Limited 2008
 *  Licensed under the Apache License, Version 2.0 (the "License"); 
 *  
 *  You may not use this file except in compliance with permission LGPL2 license. 
 *  You may obtain a copy of the License at: 
 *  
 *  http://creativecommons.org/licenses/by-sa/3.0/
 *
 *  This file is distributed on an "AS IS" BASIS, WITHOUT WARRANTIES OR 
 *  CONDITIONS OF ANY KIND, either express or implied. See the License for the 
 *  specific language governing permissions and limitations under the License.
 * ***************************************************************************** 
 * 
 * Algorithms in this code file are based on the MS-OFFCRYPT.PDF provided by
 * Microsoft as part of its Open Specification Promise (OSP) program and which 
 * is available here:
 * 
 * http://msdn.microsoft.com/en-us/library/cc313071.aspx
 * 
 */

/*
 * Modified by Danilo Mirkovic, Oct 2009
 * - works with NPOI (for OLE Compound File access)
 * - added a few methods for convenience (e.g. EncryptToFile, EncryptToStream)
 */

using System;
using System.Collections.Generic;
using System.Text;
using System.Runtime.InteropServices;
using System.IO.Packaging;
using System.IO;

namespace OfficeOpenXmlCrypto
{
	/// <summary>
	/// This class tries to implement the algorithms documents in MS-OFFCRYPTO 2.3.4.7-2.3.4.9.
	/// Usage: OfficeCryptoTest.OfficeCrypto.OfficePasswordHash();
	/// 
	/// It is intended to be used with .NET 3.0 or later and it is assumes the WindowsBase 
	/// assembly is referenced by the project to include the System.IO.Packaging namespace.
	/// </summary>
	/// <remarks>
	/// -------------------------------------------------------------------------
	///
	/// This is the content of the EncryptionInfo stream created by Excel 2007
	/// when saving a password protected xlsb file.  The password used is:
	/// "password"
	///
	/// 00000000 03 00 02 00                                         				Version
	/// 00000000             24 00 00 00						            		Flags (fCryptoAPI + fAES)
	/// 00000000                         A4 00 00 00			                	Header length
	/// 00000000                                     24 00 00 00                 	Flags (again)
	/// 00000010 00 00 00 00                                         				Size extra
	/// 00000010             0E 66 00 00                                 			Alg ID 0x0000660E = 128-bit AES,0x0000660F  192-bit AES, 0x00006610  256-bit AES
	/// 00000010                         04 80 00 00                         		Alg hash ID 0x00008004 SHA1
	/// 00000010                                     80 00 00 00                 	Key size AES = 0x00000080, 0x000000C0, 0x00000100  128, 192 or 256-bit 
	/// 00000020 18 00 00 00                                         				Provider type 0x00000018 AES
	/// 00000020             A0 C7 DC 02 00 00 00 00                         		Reserved
	/// 00000020                                     4D 00 69 00             M?i?	CSP Name
	/// 00000030 63 00 72 00 6F 00 73 00 6F 00 66 00 74 00 20 00 c?r?o?s?o?f?t? ?
	/// 00000040 45 00 6E 00 68 00 61 00 6E 00 63 00 65 00 64 00 E?n?h?a?n?c?e?d?
	/// 00000050 20 00 52 00 53 00 41 00 20 00 61 00 6E 00 64 00  ?R?S?A? ?a?n?d?
	/// 00000060 20 00 41 00 45 00 53 00 20 00 43 00 72 00 79 00  ?A?E?S? ?C?r?y?
	/// 00000070 70 00 74 00 6F 00 67 00 72 00 61 00 70 00 68 00 p?t?o?g?r?a?p?h?
	/// 00000080 69 00 63 00 20 00 50 00 72 00 6F 00 76 00 69 00 i?c? ?P?r?o?v?i?
	/// 00000090 64 00 65 00 72 00 20 00 28 00 50 00 72 00 6F 00 d?e?r? ?(?P?r?o?
	/// 000000A0 74 00 6F 00 74 00 79 00 70 00 65 00 29 00 00 00 t?o?t?y?p?e?)
	/// 
	/// 000000B0 10 00 00 00                                         				Key size
	/// 000000B0             90 AC 68 0E 76 F9 43 2B 8D 13 B7 1D                 	Salt
	/// 000000C0 B7 C0 FC 0D                                     			
	/// 000000C0             43 8B 34 B2 C6 0A A1 E1 0C 40 81 CE                 	Encrypted verifier
	/// 000000D0 83 78 F4 7A                                    
	/// 000000D0             14 00 00 00                                 			Hash length
	/// 000000D0                         48 BF F0 D6 C1 54 5C 40                 	EncryptedVerifierHash
	/// 000000E0 FE 7D 59 0F 8A D7 10 B4 C5 60 F7 73 99 2F 3C 8F 
	/// 000000F0 2C F5 6F AB 3E FB 0A D5                        
	///
	/// -------------------------------------------------------------------------
	/// </remarks>
	public class OfficeCrypto
	{
		#region The entry points

		/// <summary>
		/// Static test function exposed by the class
		/// </summary>
		public static Package OfficePasswordHash(string filename, string password)
		{
			OfficeCrypto officeCrypto = new OfficeCrypto();

			try
			{
				Package package = officeCrypto.DecryptToPackage(filename, password);
				Console.WriteLine("Package decrypted and opened");
				return package;
			}
			catch (Exception ex)
			{
				Console.WriteLine(ex.Message);
			}

			return null;
		}

		public static bool EncryptPackageFile(string filename, string password, out byte[] encryptionInfo, out byte[] encryptedPackage)
		{
			OfficeCrypto officeCrypto = new OfficeCrypto();

			encryptionInfo = null;
			encryptedPackage = null;

			try
			{
				officeCrypto.EncryptPackage(filename, password, out encryptionInfo, out encryptedPackage);
				Console.WriteLine("Package encrypted");
				officeCrypto.TestEncrytion(password, encryptionInfo, encryptedPackage);
				return true;
			}
			catch (Exception ex)
			{
				Console.WriteLine(ex.Message);
			}

			return false;
		}

		#endregion

		#region Everything else

		#region Prolog

		#region Enumerations

		[Flags]
		enum EncryptionFlags
		{
			None		= 0,
			Reserved1	= 1,	// MUST be 0, and MUST be ignored.
			Reserved2	= 2,	// MUST be 0, and MUST be ignored.
			fCryptoAPI	= 4,	// A flag that specifies whether CryptoAPI RC4 or [ECMA-376] encryption is used. MUST be 1 unless fExternal is 1. If fExternal is 1, MUST be 0. 
			fDocProps	= 8,	// MUST be 0 if document properties are encrypted. Encryption of document properties is specified in section 2.3.5.4. 
			fExternal	= 16,	// If extensible encryption is used, MUST be 1. If this field is 1, all other fields in this structure MUST be 0. 
			fAES		= 32,	// If the protected content is an [ECMA-376] document, MUST be 1. If the fAES bit is 1, the fCryptoAPI bit MUST also be 1.
		}

		enum AlgId
		{
			ByFlags	= 0x00,
			RC4		= 0x00006801,
			AES128	= 0x0000660E,
			AES192	= 0x0000660F, 
			AES256	= 0x00006610
		}

		enum AlgHashId
		{
			Any		= 0x00,
			RC4		= 0x00008000,
			SHA1	= 0x00008004
		}

		enum ProviderType
		{
			Any = 0x00000000,
			RC4 = 0x00000001,
			AES = 0x00000018
		}

		#endregion

		#region Constants

		internal const string csEncryptionInfoStreamName = "EncryptionInfo";
		internal const string csEncryptedPackageStreamName = "EncryptedPackage";

		internal const string csNotStorage = "The file is not an OLE Storage document.";
		internal const string csNoEntry = "The file does not contain an entry called {0}.";
		internal const string csNoStream = "The file does not contain a stream called {0}.";
		internal const string csNotStream = "The entry in the file called {0} is not a stream.";

		#endregion

		#region Class variables

		ushort			versionMajor		= 0;
		ushort			versionMinor		= 0;
		EncryptionFlags encryptionFlags			= EncryptionFlags.fCryptoAPI | EncryptionFlags.fAES;

		// Encryption header
		uint			sizeExtra				= 0;
		AlgId			algId					= AlgId.AES128;
		AlgHashId		algHashId				= AlgHashId.SHA1;
		int				keySize					= 0x80; // AES = 0x00000080, 0x000000C0, 0x00000100  128, 192 or 256-bit 
		ProviderType	providerType			= ProviderType.AES;
		// 8-bytes reserved here
		string			CSPName					= "";

		// Encryption verifier
		int				saltSize				= 0x10;  // Default
		byte[]			salt					= null;
		byte[]			encryptedVerifier		= null;
		int				verifierHashSize		= 0x14;
		byte[]			encryptedVerifierHash	= null;

		byte[]			encryptedPackage		= null;

		#endregion

		#region Constructor

		public OfficeCrypto()
		{
			// Check the SHA1 and AES functions work
			// if (!TestSHA1()) return;
			// if (!TestAES()) return;

			// byte[] input = new byte[0x10];
			// byte[] result = AESCryptoAPIEncrypt(input, new byte[0x10]);
			// Console.WriteLine(result.Length);
		}

		#endregion

		#endregion

		#region To package

		/// <summary>
		/// Validates and opens the storage containing the encrypted package.
		/// Reads the encryption information and encrypted package
		/// Parses the encryption information
		/// Generates a decrypytion key and validates it against the password
		/// Decrypts the encrypted package and creates and Packaging.Package.
		/// </summary>
		/// <param name="filename">The name of the storage file containing the encrypted package</param>
		/// <param name="password">The password to decrypt the package</param>
		/// <returns>System.IO.Packaging.Package instance</returns>
		public Package DecryptToPackage(string filename, string password)
		{
			return CreatePackage(DecryptToArray(filename, password));
		}

		/// <summary>
		/// Validates and opens the storage containing the encrypted package.
		/// Reads the encryption information and encrypted package
		/// Parses the encryption information
		/// Generates a decrypytion key and validates it against the password
		/// Decrypts the encrypted package and creates and Packaging.Package.
		/// </summary>
		/// <param name="filename">A stream of a storage file containing the encrypted package</param>
		/// <param name="password">The password to decrypt the package</param>
		/// <returns>System.IO.Packaging.Package instance</returns>
        public Package DecryptToPackage(byte[] contents, string password)
		{
			return CreatePackage(DecryptToArray(contents, password));
		}

		/// <summary>
		/// Reads the encryption information and encrypted package
		/// Parses the encryption information
		/// Generates a decrypytion key and validates it against the password
		/// Decrypts the encrypted package and creates and Packaging.Package.
		/// </summary>
		/// <param name="filename">The storage file containing the encrypted package</param>
		/// <param name="password">The password to decrypt the package</param>
		/// <returns>System.IO.Packaging.Package instance</returns>
		public Package DecryptToPackage(OleStorage stgRoot, string password)
		{
			return CreatePackage(DecryptToArray(stgRoot, password));
		}

		#endregion

		#region To MemoryStream

		/// <summary>
		/// Validates and opens the storage containing the encrypted package.
		/// Reads the encryption information and encrypted package
		/// Parses the encryption information
		/// Generates a decrypytion key and validates it against the password
		/// Decrypts the encrypted package and creates and Packaging.Package.
		/// </summary>
		/// <param name="filename">The name of the storage file containing the encrypted package</param>
		/// <param name="password">The password to decrypt the package</param>
		/// <returns>System.IO.MemoryStream instance</returns>
		public MemoryStream DecryptToStream(string filename, string password)
		{
			return CreateStream(DecryptToArray(filename, password));
		}

		/// <summary>
		/// Validates and opens the storage containing the encrypted package.
		/// Reads the encryption information and encrypted package
		/// Parses the encryption information
		/// Generates a decrypytion key and validates it against the password
		/// Decrypts the encrypted package and creates and Packaging.Package.
		/// </summary>
		/// <param name="filename">A stream of a storage file containing the encrypted package</param>
		/// <param name="password">The password to decrypt the package</param>
		/// <returns>System.IO.MemoryStream instance</returns>
        public MemoryStream DecryptToStream(byte[] contents, string password)
		{
			return CreateStream(DecryptToArray(contents, password));
		}

		/// <summary>
		/// Reads the encryption information and encrypted package
		/// Parses the encryption information
		/// Generates a decrypytion key and validates it against the password
		/// Decrypts the encrypted package and creates and Packaging.Package.
		/// </summary>
		/// <param name="filename">The storage file containing the encrypted package</param>
		/// <param name="password">The password to decrypt the package</param>
		/// <returns>System.IO.MemoryStream instance</returns>
		public MemoryStream DecryptToStream(OleStorage stgRoot, string password)
		{
			return CreateStream(DecryptToArray(stgRoot, password));
		}

		#endregion

		#region To Byte array

		/// <summary>
		/// Validates and opens the storage containing the encrypted package.
		/// Reads the encryption information and encrypted package
		/// Parses the encryption information
		/// Generates a decrypytion key and validates it against the password
		/// Decrypts the encrypted package and creates and Packaging.Package.
		/// </summary>
		/// <param name="filename">The name of the storage file containing the encrypted package</param>
		/// <param name="password">The password to decrypt the package</param>
		/// <returns>System.IO.Packaging.Package instance</returns>
		public byte[] DecryptToArray(string filename, string password)
		{
			Console.WriteLine("Open the storage");			
            OleStorage stgRoot = new OleStorage(filename);

			return DecryptToArray(stgRoot, password);
		}

        /// <summary>
        /// Validates and opens the storage containing the encrypted package.
        /// Reads the encryption information and encrypted package
        /// Parses the encryption information
        /// Generates a decrypytion key and validates it against the password
        /// Decrypts the encrypted package and creates and Packaging.Package.
        /// </summary>
        /// <param name="filename">The name of the storage file containing the encrypted package</param>
        /// <param name="password">The password to decrypt the package</param>
        /// <returns>System.IO.Packaging.Package instance</returns>
        public byte[] DecryptToArray(byte[] contents, string password)
        {
            Console.WriteLine("Open the storage");
            OleStorage stgRoot = new OleStorage(contents);

            return DecryptToArray(stgRoot, password);
        }

		/// <summary>
		/// Reads the encryption information and encrypted package
		/// Parses the encryption information
		/// Generates a decrypytion key and validates it against the password
		/// Decrypts the encrypted package and creates and Packaging.Package.
		/// </summary>
		/// <param name="filename">The storage file containing the encrypted package</param>
		/// <param name="password">The password to decrypt the package</param>
		/// <returns>System.IO.Packaging.Package instance</returns>
		public byte[] DecryptToArray(OleStorage stgRoot, string password)
		{
            // TODO: check if the EncryptedPackage exists (if it's a valid doc)
            encryptedPackage = stgRoot.ReadStream(csEncryptedPackageStreamName);
			
            // TODO: check if the EncryptionInfo exists (if it's a valid doc)
            byte[] encryptionInfo = stgRoot.ReadStream(csEncryptionInfoStreamName);

			// Delegate the rest to this common function
			return DecryptInternal(password, encryptionInfo, encryptedPackage);
		}

		#endregion


        #region Encrypt to File

        /// <summary>
        /// Encryptes the package to a file, using the given password. 
        /// </summary>
        /// <param name="packageContents">Plaintext contents of the package.</param>
        /// <param name="password">Password to use to encrypt.</param>
        /// <param name="encryptedFilename">Name of the encrypted file to save</param>
        public void EncryptToFile(byte[] packageContents, string password, string encryptedFilename)
        {
            byte[] encryptionInfo;
            byte[] encryptedPackage;
            EncryptPackage(packageContents, password, out encryptionInfo, out encryptedPackage);

            OleStorage storage = new OleStorage();
            storage.WriteStream(csEncryptionInfoStreamName, encryptionInfo);
            storage.WriteStream(csEncryptedPackageStreamName, encryptedPackage);
            storage.Save(encryptedFilename);
        }

        /// <summary>
        /// Encryptes the package to a stream, using the given password. 
        /// </summary>
        /// <param name="packageContents">Plaintext contents of the package.</param>
        /// <param name="password">Password to use to encrypt.</param>
        /// <param name="encryptedFilename">Name of the encrypted stream write to</param>
        public void EncryptToStream(byte[] packageContents, string password, Stream encryptedStream)
        {
            byte[] encryptionInfo;
            byte[] encryptedPackage;
            EncryptPackage(packageContents, password, out encryptionInfo, out encryptedPackage);

            OleStorage storage = new OleStorage();
            storage.WriteStream(csEncryptionInfoStreamName, encryptionInfo);
            storage.WriteStream(csEncryptedPackageStreamName, encryptedPackage);
            storage.Save(encryptedStream);
        }


        #endregion

        #region Public methods

		/// <summary>
		/// Encrypts a package (zip) file using a supplied password and returns 
		/// an array to create an encryption information stream and a byte array 
		/// of the encrypted package.
		/// </summary>
		/// <param name="filename">The package (zip) file to be encrypted</param>
		/// <param name="password">The password to decrypt the package</param>
		/// <param name="encryptionInfo">An array of bytes containing the encrption info</param>
		/// <param name="encryptedPackage">The encrpyted package</param>
		/// <returns></returns>
		public void EncryptPackage(string filename, string password, out byte[] encryptionInfo, out byte[] encryptedPackage)
		{
			if (!File.Exists(filename)) throw new ArgumentException("Package file does not exist");

			// Grab the package contents and encrypt
			byte[] packageContents = null;
			using(System.IO.FileStream fs = new System.IO.FileStream(filename, System.IO.FileMode.Open, System.IO.FileAccess.Read, System.IO.FileShare.ReadWrite))
			{
				packageContents = new byte[fs.Length];
				fs.Read(packageContents, 0, packageContents.Length);
			}

			EncryptPackage(packageContents, password, out encryptionInfo, out encryptedPackage);
		}

		/// <summary>
		/// Encrypts a package (zip) file using a supplied password and returns 
		/// an array to create an encryption information stream and a byte array 
		/// of the encrypted package.
		/// </summary>
		/// <param name="packageContents">The package (zip) file to be encrypted</param>
		/// <param name="password">The password to decrypt the package</param>
		/// <param name="encryptionInfo">An array of bytes containing the encrption info</param>
		/// <param name="encryptedPackage">The encrpyted package</param>
		/// <returns></returns>
		public void EncryptPackage(byte[] packageContents, string password, out byte[] encryptionInfo, out byte[] encryptedPackage)
		{
			versionMajor = 3;
			versionMinor = 2;

			// Console.WriteLine(password);

			// this.algId = AlgId.AES256;
			// this.keySize = 0x100;

			// Generate a salt
			System.Security.Cryptography.RijndaelManaged aes = new System.Security.Cryptography.RijndaelManaged();
			saltSize = 0x10;
			byte[] tempSalt = SHA1Hash(aes.IV);
			aes = null;
			this.verifierHashSize = tempSalt.Length;

			salt = new byte[saltSize];
			Array.Copy(tempSalt, salt, saltSize);
			
			// Generate a key from salt and password
			byte[] key = GeneratePasswordHashUsingSHA1(password);

			CreateVerifier(key);

			int originalLength = packageContents.Length;

			// Pad the array to the nearest 16 byte boundary
			int remainder = packageContents.Length % 0x10;
			if (remainder != 0)
			{
				byte[] tempContents = new byte[packageContents.Length + 0x10 - remainder];
				Array.Copy(packageContents, tempContents, packageContents.Length);
				packageContents = tempContents;
			}

			byte[] encryptionResult = AESEncrypt(packageContents, key);

			// Need to prepend the original package size as a Int64 (8 byte) field
			encryptedPackage = new byte[encryptionResult.Length +  8];
			// MUST record the original length here
			Array.Copy(BitConverter.GetBytes((long)originalLength), encryptedPackage, 8);  
			Array.Copy(encryptionResult, 0, encryptedPackage, 8, encryptionResult.Length);

			byte[] encryptionHeader = null;

			// Generate the encryption header structure
			using (System.IO.MemoryStream ms = new System.IO.MemoryStream())
			{
				System.IO.BinaryWriter br = new System.IO.BinaryWriter(ms);
				br.Write((int)this.encryptionFlags);
				br.Write((int)this.sizeExtra);
				br.Write((int)this.algId);
				br.Write((int)this.algHashId);
				br.Write((int)this.keySize);
				br.Write((int)this.providerType);
				br.Write(new byte[] { 0xA0, 0xC7, 0xDC, 0x02, 0x00, 0x00, 0x00, 0x00 } );
				br.Write(System.Text.UnicodeEncoding.Unicode.GetBytes("Microsoft Enhanced RSA and AES Cryptographic Provider (Prototype)\0"));

				ms.Flush();
				encryptionHeader = ms.ToArray();
			}

			byte[] encryptionVerifier = null;

			// Generate the encryption header structure
			using (System.IO.MemoryStream ms = new System.IO.MemoryStream())
			{
				System.IO.BinaryWriter br = new System.IO.BinaryWriter(ms);
				br.Write((int)salt.Length);
				br.Write(this.salt);
				br.Write(this.encryptedVerifier);
				br.Write(this.verifierHashSize); // Hash length
				br.Write(this.encryptedVerifierHash);

				ms.Flush();
				encryptionVerifier = ms.ToArray();
			}

			// Now generate the encryption info structure
			using (System.IO.MemoryStream ms = new System.IO.MemoryStream())
			{
				System.IO.BinaryWriter br = new System.IO.BinaryWriter(ms);
				br.Write(versionMajor);
				br.Write(versionMinor);
				br.Write((int)this.encryptionFlags);
				br.Write((int)encryptionHeader.Length);
				br.Write(encryptionHeader);
				br.Write(encryptionVerifier);

				ms.Flush();
				encryptionInfo = ms.ToArray();
			}

			// Console.WriteLine( Lyquidity.OLEStorage.Utilities.FormatData(encryptionInfo, 128, 0) );
		}

		public void TestEncrytion(string password, byte[] encryptionInfo, byte[] encryptedPackage)
		{
			byte[] array = DecryptInternal(password, encryptionInfo, encryptedPackage);
		}

		#endregion

		#region Private functions

		#region 2.3.4.7 & 2.3.4.9 algorithm implementation methods

		private void CreateVerifier(byte[] key)
		{
			// Much of the commmentary in this function is taken from 2.3.3
			// The EncryptionVerifier structure MUST be set using the following process: 

			// 1)	Random data are generated and written into the Salt field. 
			// 2)	The encryption key is derived from the password and salt, as specified in section 
			//		2.3.4.7 or 2.3.5.2, with block number 0.

			//		This is passed in as parameter key

			// 3)	Generate 16 bytes of additional random data as the Verifier.

			System.Security.Cryptography.RijndaelManaged aes = new System.Security.Cryptography.RijndaelManaged();
			byte[] verifier = aes.IV;
			aes = null;

			// Console.WriteLine("Verifier");
			// Console.WriteLine( Lyquidity.OLEStorage.Utilities.FormatData(verifier, 32, 0) );
			// Console.WriteLine();

			// 4)	Results of step 3 are encrypted and written into the EncryptedVerifier field.

			encryptedVerifier = AESEncrypt( verifier, key);

			// Console.WriteLine("encryptedVerifier");
			// Console.WriteLine( Lyquidity.OLEStorage.Utilities.FormatData(encryptedVerifier, 32, 0) );
			// Console.WriteLine();

			// 5)	For the hashing algorithm chosen, obtain the size of the hash data and write this value 
			//		into the VerifierHashSize field.
 
			// Not applicable right now

			// 6)	Obtain the hashing algorithm output using an input of data generated in step 3. 

			byte[] verifierHash = SHA1Hash( verifier );
			// Console.WriteLine("verifierHash");
			// Console.WriteLine( Lyquidity.OLEStorage.Utilities.FormatData(verifierHash, 32, 0) );
			// Console.WriteLine();

			// 7)	Encrypt the hashing algorithm output from step 6 using the encryption algorithm 
			//		chosen, and write the output into the EncryptedVerifierHash field.

			// First pad to 32 bytes
			byte[] tempHash = new byte[0x20];
			Array.Copy(verifierHash, tempHash, verifierHash.Length);
			verifierHash = tempHash;

			encryptedVerifierHash = AESEncrypt( verifierHash, key );

			// Console.WriteLine("encryptedVerifierHash");
			// Console.WriteLine( Lyquidity.OLEStorage.Utilities.FormatData(encryptedVerifierHash, 32, 0) );
			// Console.WriteLine();
		}

		/// <summary>
		/// Implements the password verifier process
		/// </summary>
		/// <param name="key"></param>
		/// <param name="encryptedVerifier">An array of the encryptedVerifier bytes</param>
		/// <param name="encryptedVerifierHash">An array of the encryptedVerifierHash bytes</param>
		/// <returns>True if the password is a match</returns>
		private bool PasswordVerifier(byte[] key)
		{
			// Decrypt the encrypted verifier...
			byte[] decryptedVerifier = AESDecrypt( encryptedVerifier, key);

			// Truncate
			byte[] data = new byte[16];
			Array.Copy(decryptedVerifier, data, data.Length);
			decryptedVerifier = data;

			// ... and hash
			byte[] decryptedVerifierHash = AESDecrypt( encryptedVerifierHash, key );

			// Hash the decrypted verifier (2.3.4.9)
			byte[] checkHash = SHA1Hash(decryptedVerifier);

			// Check the 
			for (int i = 0; i < checkHash.Length; i++)
			{
				if (decryptedVerifierHash[i] != checkHash[i])
					return false;
			}

			return true;
		}

		/// <summary>
		/// Implements (tries to) the hash key generation algorithm in section 2.3.4.7
		/// The 
		/// </summary>
		/// <param name="salt">A salt (taken from the EncrptionInfo stream)</param>
		/// <param name="password">The password used to decode the stream</param>
		/// <param name="sha1HashSize">Size of the hash (taken from theEncrptionInfo stream)</param>
		/// <param name="keySize">The keysize (taken from theEncrptionInfo stream)</param>
		/// <returns>The derived encryption key byte array</returns>
		private byte[] GeneratePasswordHashUsingSHA1(string password)
		{
			byte[] hashBuf = null;

			try
			{
				// H(0) = H(salt, password);
				hashBuf = SHA1Hash(salt, password);

				for (int i = 0; i < 50000; i++)
				{
					// Generate each hash in turn
					// H(n) = H(i, H(n-1))
					hashBuf = SHA1Hash(i, hashBuf);
				}

				// Finally, append "block" (0) to H(n)
				hashBuf = SHA1Hash(hashBuf, 0);

				// The algorithm in this 'DeriveKey' function is the bit that's not clear from the documentation
				byte[] key = DeriveKey(hashBuf);

				// Should handle the case of longer key lengths as shown in 2.3.4.9
				// Grab the key length bytes of the final hash as the encrypytion key
				byte[] final = new byte[keySize/8];
				Array.Copy(key, final, final.Length);

				return final;

			}
			catch (Exception ex)
			{
				Console.WriteLine(ex.Message);
			}

			return null;
		}

		#endregion

		internal byte[] DecryptInternal(string password, byte[] encryptionInfo, byte[] encryptedPackage)
		{
			#region Parse the encryption info data

			using (System.IO.MemoryStream ms = new System.IO.MemoryStream(encryptionInfo))
			{
				System.IO.BinaryReader reader = new System.IO.BinaryReader(ms);

				versionMajor = reader.ReadUInt16();
				versionMinor = reader.ReadUInt16();

				encryptionFlags = (EncryptionFlags)reader.ReadUInt32();
				if (encryptionFlags == EncryptionFlags.fExternal)
					throw new Exception("An external cryptographic provider is not supported");

				// Encryption header
				uint headerLength		= reader.ReadUInt32(); 
				int skipFlags			= reader.ReadInt32(); headerLength -= 4;
				sizeExtra				= reader.ReadUInt32(); headerLength -= 4;
				algId					= (AlgId)reader.ReadUInt32(); headerLength -= 4;
				algHashId				= (AlgHashId)reader.ReadUInt32(); headerLength -= 4;
				keySize					= reader.ReadInt32(); headerLength -= 4;
				providerType			= (ProviderType)reader.ReadUInt32(); headerLength -= 4;
										  reader.ReadUInt32(); headerLength -= 4; // Reserved 1
										  reader.ReadUInt32(); headerLength -= 4; // Reserved 2
				CSPName					= System.Text.UnicodeEncoding.Unicode.GetString(reader.ReadBytes((int)headerLength));

				// Encryption verifier
				saltSize				= reader.ReadInt32(); 
				salt					= reader.ReadBytes(saltSize);
				encryptedVerifier		= reader.ReadBytes(0x10);
				verifierHashSize		= reader.ReadInt32(); 
				encryptedVerifierHash	= reader.ReadBytes(providerType == ProviderType.RC4 ? 0x14 : 0x20);
			}

			#endregion

			#region Encryption key generation

			Console.WriteLine("Encryption key generation");
			byte[] encryptionKey = GeneratePasswordHashUsingSHA1(password);
			if (encryptionKey == null) return null;

			#endregion

			#region Password verification

			Console.WriteLine("Password verification");
			if (PasswordVerifier(encryptionKey))
			{
			    Console.WriteLine("Password verification succeeded");
			} 
			else
			{
			    Console.WriteLine("Password verification failed");
			    throw new InvalidPasswordException("The password is not valid");
			}

			#endregion

			#region Decrypt

			// First 8 bytes hold the size of the stream
			long length = BitConverter.ToInt64(encryptedPackage, 0);

			// Decrypt the stream using the generated and validated key
			Console.WriteLine("Decrypt the stream using the generated and validated key");
			encryptedPackage = AESDecrypt(encryptedPackage, 8, encryptedPackage.Length-8, encryptionKey);

			// !! IMPORTANT !! Make sure the final array is the correct size
			// Failure to do this will cause an error when the decrypted stream
			// is opened by the System.IO.Packaging.Package.Open() method.

			byte[] result = encryptedPackage;

			if (encryptedPackage.Length > length)
			{
				result = new byte[length];
				Array.Copy(encryptedPackage, result, result.Length);
			}

			//using (System.IO.FileStream fs = new System.IO.FileStream(@"c:\x.zip", System.IO.FileMode.Create))
			//{
			//    fs.Write(result, 0, result.Length);
			//}

			return result;

			#endregion

		}

		private System.IO.MemoryStream CreateStream(byte[] decryptedPackage)
		{
			System.IO.MemoryStream ms = new System.IO.MemoryStream();

			ms.Write(decryptedPackage, 0, decryptedPackage.Length);
			ms.Flush();
			ms.Position = 0;

			return ms;
		}

		private Package CreatePackage(byte[] decryptedPackage)
		{ 
			using (MemoryStream ms = CreateStream(decryptedPackage))
			{
				return Package.Open(ms, FileMode.Open, FileAccess.ReadWrite);
			}

            // What happens if we later try writing to a closed stream?
		}

		#region SHA1/AES test functions

		/// <summary>
		/// This method tests that the AES implementation works as it should given
		/// a key and some data to encrypt/decrypt.
		/// </summary>
		/// <returns>
		/// True if the test confirms the AES implementation generations 
		/// the expected values
		/// </returns>
		private bool TestAES()
		{
			// These test ciphers are available in this Wikipedia article:
			// http://en.wikipedia.org/wiki/Advanced_Encryption_Standard#C.23_.2F.NET
			byte[] key =      { 0x00, 0x01, 0x02, 0x03, 0x05, 0x06, 0x07, 0x08, 0x0a, 0x0b, 0x0c, 0x0d, 0x0f, 0x10, 0x11, 0x12 };
			byte[] data =     { 0x50, 0x68, 0x12, 0xa4, 0x5f, 0x08, 0xc8, 0x89, 0xb9, 0x7f, 0x59, 0x80, 0x03, 0x8b, 0x83, 0x59 };
			byte[] expected = { 0xD8, 0xF5, 0x32, 0x53, 0x82, 0x89, 0xEF, 0x7D, 0x06, 0xB5, 0x06, 0xA4, 0xFD, 0x5B, 0xE9, 0xC9 };

			byte[] result = AESEncrypt(data, key);

			//if (result.Length != expected.Length) return false;

			//for (int i = 0; i < result.Length; i++)
			//{
			//    if (result[i] != expected[i]) return false;
			//}

			result = AESDecrypt(result, key);
			//if (result.Length != data.Length) return false;

			//for (int i = 0; i < result.Length; i++)
			//{
			//    if (result[i] != data[i]) return false;
			//}

			return true;
		}

		/// <summary>
		/// Tests the SHA1 implementation to confirm it generates the expected
		/// hash when given a known string as input.
		/// </summary>
		/// <returns></returns>
		private bool TestSHA1()
		{
			// This example text and resulting SHA1 hash is from the Wikipedia article:
			// http://en.wikipedia.org/wiki/SHA-1#Example_hashes
			string test = "The quick brown fox jumps over the lazy dog";
			byte[] expected = { 0x2f, 0xd4, 0xe1, 0xc6, 
								0x7a, 0x2d, 0x28, 0xfc, 
								0xed, 0x84, 0x9e, 0xe1, 
								0xbb, 0x76, 0xe7, 0x39, 
								0x1b, 0x93, 0xeb, 0x12 };

			byte[] result = SHA1Hash(System.Text.ASCIIEncoding.ASCII.GetBytes(test));

			for (int i = 0; i < result.Length; i++)
			{
				if (result[i] != expected[i])
					return false;
			}

			return true;
		}

		#endregion

		#region Derive key

		private byte[] DeriveKey(byte[] hashValue)
		{
            // And one more hash to derive the key
            byte[] derivedKey = new byte[64];

			// This is step 4a in 2.3.4.7 of MS_OFFCRYPT version 1.0
			// and is required even though the notes say it should be 
			// used only when the encryption algorithm key > hash length.
            for (int i = 0; i < derivedKey.Length; i++)
                derivedKey[i] = (byte)(i < hashValue.Length ? 0x36 ^ hashValue[i] : 0x36);

			byte[] X1 = SHA1Hash(derivedKey);

			if (verifierHashSize > keySize/8)
				return X1;

            for (int i = 0; i < derivedKey.Length; i++)
                derivedKey[i] = (byte)(i < hashValue.Length ? 0x5C ^ hashValue[i] : 0x5C);

			byte[] X2 = SHA1Hash(derivedKey);

			byte[] X3 = new byte[X1.Length + X2.Length];

			Array.Copy(X1, 0, X3, 0, X1.Length);
			Array.Copy(X1, 0, X3, X1.Length, X2.Length);

			return X3;
		}

		#endregion

		#region SHA1 functions

		private byte[] SHA1Hash(byte[] salt, string password)
		{
			return SHA1Hash(HashPassword(salt, password));
		}

		private byte[] HashPassword(byte[] salt, string password)
		{
			// Use a unicode form of the password
			byte[] passwordBuf = System.Text.UnicodeEncoding.Unicode.GetBytes(password);
			byte[] inputBuf = new byte[salt.Length + passwordBuf.Length];
			Array.Copy(salt, inputBuf, salt.Length);
			Array.Copy(passwordBuf, 0, inputBuf, salt.Length, passwordBuf.Length);

			return inputBuf;
		}

		private byte[] SHA1Hash(int iterator, byte[] hashBuf)
		{
			// Create an input buffer for the hash.  This will be 4 bytes larger than 
			// the hash to accommodate the unsigned int iterator value.
			byte[] inputBuf = new byte[0x14 + 0x04];

			// Create a byte array of the integer and put at the front of the input buffer
			// 1.3.6 says that little-endian byte ordering is expected

			// Copy the iterator bytes into the hash input buffer
			Array.Copy(System.BitConverter.GetBytes(iterator), inputBuf, 4); 

			// 'append' the previously generated hash to the input buffer
			Array.Copy(hashBuf, 0, inputBuf, 4, hashBuf.Length);

			return SHA1Hash(inputBuf);
		}

		private byte[] SHA1Hash(byte[] hashBuf, int block)
		{
			// Create an input buffer for the hash.  This will be 4 bytes larger than 
			// the hash to accommodate the unsigned int iterator value.
			byte[] inputBuf = new byte[0x14 + 0x04];

			Array.Copy(hashBuf, inputBuf, hashBuf.Length);
			Array.Copy(System.BitConverter.GetBytes(block), 0, inputBuf, hashBuf.Length, 4);

			return SHA1Hash(inputBuf);
		}

		private byte[] SHA1Hash(byte[] hashBuf, byte[] block0)
		{
			// Create an input buffer for the hash.  This will be 4 bytes larger than 
			// the hash to accommodate the unsigned int iterator value.
			byte[] inputBuf = new byte[hashBuf.Length + block0.Length];

			Array.Copy(hashBuf, inputBuf, hashBuf.Length);
			Array.Copy(block0, 0, inputBuf, hashBuf.Length, block0.Length);

			return SHA1Hash(inputBuf);
		}

		private byte[] SHA1Hash(byte[] inputBuffer)
		{
			System.Security.Cryptography.SHA1 sha1 = System.Security.Cryptography.SHA1CryptoServiceProvider.Create();
			return sha1.ComputeHash(inputBuffer);
		}

		#endregion

		#region AES functions

		private byte[] AESDecrypt(byte[] data, byte[] key)
		{
			return AESDecrypt(data, 0, data.Length, key);
		}

		private byte[] AESDecrypt(byte[] data, int index, int count, byte[] key)
		{
			byte[] decryptedBytes = null;

			//  Create uninitialized Rijndael encryption object.
			 System.Security.Cryptography.RijndaelManaged symmetricKey = new  System.Security.Cryptography.RijndaelManaged();

			// It is required that the encryption mode is Electronic Codebook (ECB) 
			// see MS-OFFCRYPTO v1.0 2.3.4.7 pp 39.
			symmetricKey.Mode = System.Security.Cryptography.CipherMode.ECB;
			symmetricKey.Padding = System.Security.Cryptography.PaddingMode.None;
			symmetricKey.KeySize = keySize;
			// symmetricKey.IV = null; // new byte[16];
			// symmetricKey.Key = key;

			//  Generate decryptor from the existing key bytes and initialization 
			//  vector. Key size will be defined based on the number of the key 
			//  bytes.
			System.Security.Cryptography.ICryptoTransform decryptor;
			decryptor = symmetricKey.CreateDecryptor(key, null);

			//  Define memory stream which will be used to hold encrypted data.
			using (System.IO.MemoryStream memoryStream = new System.IO.MemoryStream(data, index, count))
			{
				//  Define memory stream which will be used to hold encrypted data.
				using (System.Security.Cryptography.CryptoStream cryptoStream 
						= new System.Security.Cryptography.CryptoStream(memoryStream, decryptor, System.Security.Cryptography.CryptoStreamMode.Read))
				{
					//  Since at this point we don't know what the size of decrypted data
					//  will be, allocate the buffer long enough to hold ciphertext;
					//  plaintext is never longer than ciphertext.
					decryptedBytes = new byte[data.Length];
					int decryptedByteCount = cryptoStream.Read(decryptedBytes, 0, decryptedBytes.Length);

					return decryptedBytes;
				}
			}
		}

		private byte[] AESEncrypt(byte[] data, byte[] key)
		{
			byte[] cipherTextBytes = null;

			//  Create uninitialized Rijndael encryption object.
			System.Security.Cryptography.RijndaelManaged symmetricKey = new System.Security.Cryptography.RijndaelManaged();

			// It is required that the encryption mode is Electronic Codebook (ECB) 
			// see MS-OFFCRYPTO v1.0 2.3.4.7 pp 39.
			symmetricKey.Mode = System.Security.Cryptography.CipherMode.ECB;
			symmetricKey.Padding = System.Security.Cryptography.PaddingMode.None; // System.Security.Cryptography.PaddingMode.None;
			symmetricKey.KeySize = this.keySize;
			// symmetricKey.Key = key;
			// symmetricKey.IV = new byte[16];

			// Generate encryptor from the existing key bytes and initialization vector. 
			// Key size will be defined based on the number of the key bytes.
			System.Security.Cryptography.ICryptoTransform encryptor = symmetricKey.CreateEncryptor(key, null);

			//  Define memory stream which will be used to hold encrypted data.
			using (System.IO.MemoryStream memoryStream = new System.IO.MemoryStream())
			{
				//  Define cryptographic stream (always use Write mode for encryption).
				using (System.Security.Cryptography.CryptoStream cryptoStream 
						= new System.Security.Cryptography.CryptoStream(memoryStream, encryptor, System.Security.Cryptography.CryptoStreamMode.Write))
				{
					//  Start encrypting.
					cryptoStream.Write(data, 0, data.Length);

					//  Finish encrypting.
					cryptoStream.FlushFinalBlock();
				}

				//  Convert our encrypted data from a memory stream into a byte array.
				cipherTextBytes = memoryStream.ToArray();
				return cipherTextBytes;
			}
		}

		#endregion

		#endregion

		#endregion

	}

}

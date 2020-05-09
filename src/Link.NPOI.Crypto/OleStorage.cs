using System;
using System.Collections.Generic;
using System.Text;
using NPOI.POIFS.FileSystem;
using System.IO;

/*
 * Wrapper coded by Danilo Mirkovic, 2009
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
    /// Access to OLE Compound Storage file.
    /// Supports reading/writing of parts at root,
    /// as needed for Office Encryption.
    /// </summary>
    /// <remarks>
    /// Thin wrapper around POIFSFileSystem from NPOI project
    /// http://npoi.codeplex.com/
    /// NPOI uses Apache License 2.0
    /// </remarks>
    public class OleStorage : POIFSFileSystem
    {
        private readonly POIFSFileSystem PoiFS;

        public OleStorage()
        {
            PoiFS = new POIFSFileSystem();
        }

        /// <summary>
        /// Initialize from a byte array.
        /// </summary>
        /// <param name="stream"></param>
        public OleStorage(byte[] storageBytes)
        {
            // Note: POIFS closes the stream after reading it. 
            // It is undesirable to pass a stream in, as it can't be used any more.
            using (MemoryStream tempStream = new MemoryStream(storageBytes))
            {
                PoiFS = new POIFSFileSystem(tempStream);
            }
        }

        /// <summary>
        /// Initialize from an existing file
        /// </summary>
        /// <param name="filename"></param>
        public OleStorage(String filename)
        {
            if (!File.Exists(filename))
            {
                throw new FileNotFoundException("OLE Storage file does not exist", filename);
            }

            Stream stream = new FileStream(filename, FileMode.Open);
            PoiFS = new POIFSFileSystem(stream);
        }

        /// <summary>
        /// Read the given stream from the file
        /// </summary>
        /// <param name="streamName"></param>
        /// <returns></returns>
        public byte[] ReadStream(String streamName)
        {
            byte[] contents;
            //using (Stream stream = PoiFS.CreatePOIFSDocumentReader(streamName))
            using (Stream stream = PoiFS.CreateDocumentInputStream(streamName))
            {
                contents = new byte[stream.Length];
                stream.Read(contents, 0, contents.Length);
            }
            return contents;
        }

        /// <summary>
        /// Save the given stream within a file.
        /// </summary>
        /// <param name="streamName"></param>
        /// <param name="contents"></param>
        public void WriteStream(String streamName, byte[] contents)
        {
            using (Stream s = new MemoryStream(contents))
            {
                PoiFS.Root.CreateDocument(streamName, s);
            }
        }

        /// <summary>
        /// Save the whole container to a file.
        /// </summary>
        /// <param name="filename"></param>
        public void Save(string filename)
        {
            using (FileStream outStream = new FileStream(filename, FileMode.Create))
            {
                Save(outStream);
            }
        }

        /// <summary>
        /// Save the whole container to a stream
        /// </summary>
        /// <param name="encryptedStream"></param>
        public void Save(Stream encryptedStream)
        {
            PoiFS.WriteFileSystem(encryptedStream);
        }
    }
}

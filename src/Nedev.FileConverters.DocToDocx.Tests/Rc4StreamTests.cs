#nullable enable
using System;
using System.IO;
using Nedev.FileConverters.DocToDocx.Utils;
using Xunit;

namespace Nedev.FileConverters.DocToDocx.Tests
{
    public class Rc4StreamTests
    {
        [Fact]
        public void Rc4Cipher_EncryptsAndDecrypts()
        {
            var key = new byte[] { 1, 2, 3, 4, 5 };
            var cipher = new Rc4Cipher();
            cipher.Initialize(key);

            byte[] plain = { 10, 20, 30, 40, 50 };
            byte[] encrypted = new byte[plain.Length];
            cipher.TransformBlock(plain, 0, plain.Length, encrypted, 0);

            // decrypt by reinitializing cipher with same key (RC4 symmetric)
            cipher.Initialize(key);
            byte[] decrypted = new byte[plain.Length];
            cipher.TransformBlock(encrypted, 0, encrypted.Length, decrypted, 0);

            Assert.Equal(plain, decrypted);
        }

        [Fact]
        public void Rc4Stream_DecryptsStreamingData()
        {
            // Use a fixed base hash and simulate small stream
            byte[] baseHash = { 0xAA, 0xBB, 0xCC, 0xDD };
            byte[] data = new byte[1024];
            var rand = new Random(123);
            rand.NextBytes(data);

            // Encrypt manually using same block-based algorithm
            byte[] encrypted = new byte[data.Length];
            // copy so we can reuse Rc4Cipher
            var cipher = new Rc4Cipher();

            for (uint block = 0; block <= 1; block++)
            {
                // compute block key using MD5 (default useSha1=false)
                byte[] blockBytes = BitConverter.GetBytes(block);
                byte[] combined = new byte[baseHash.Length + blockBytes.Length];
                Buffer.BlockCopy(baseHash, 0, combined, 0, baseHash.Length);
                Buffer.BlockCopy(blockBytes, 0, combined, baseHash.Length, blockBytes.Length);
                byte[] blockKey;
                using (var md5 = System.Security.Cryptography.MD5.Create())
                {
                    blockKey = md5.ComputeHash(combined);
                }
                cipher.Initialize(blockKey);

                int offset = (int)(block * 512);
                int length = Math.Min(512, data.Length - offset);
                cipher.TransformBlock(data, offset, length, encrypted, offset);
            }

            using var ms = new MemoryStream(encrypted);
            using var rc4 = new Rc4Stream(ms, baseHash, 0, useSha1: false, leaveOpen: true);
            byte[] decrypted = new byte[data.Length];
            int read = rc4.Read(decrypted, 0, decrypted.Length);
            Assert.Equal(data.Length, read);
            Assert.Equal(data, decrypted);
        }
    }
}

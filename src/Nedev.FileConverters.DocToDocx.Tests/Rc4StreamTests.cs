#nullable enable
using System;
using System.Collections.Generic;
using System.IO;
using System.Security.Cryptography;
using System.Text;
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

        [Fact]
        public void Rc4Stream_PreservesConfiguredClearPrefix()
        {
            byte[] baseHash = { 0x10, 0x20, 0x30, 0x40 };
            byte[] clearPrefix = new byte[512];
            byte[] encryptedPayload = new byte[300];
            var rand = new Random(456);
            rand.NextBytes(clearPrefix);
            rand.NextBytes(encryptedPayload);

            byte[] plain = new byte[clearPrefix.Length + encryptedPayload.Length];
            Buffer.BlockCopy(clearPrefix, 0, plain, 0, clearPrefix.Length);
            Buffer.BlockCopy(encryptedPayload, 0, plain, clearPrefix.Length, encryptedPayload.Length);

            byte[] encrypted = new byte[plain.Length];
            Buffer.BlockCopy(clearPrefix, 0, encrypted, 0, clearPrefix.Length);

            var cipher = new Rc4Cipher();
            uint blockNumber = 1;
            byte[] blockBytes = BitConverter.GetBytes(blockNumber);
            byte[] combined = new byte[baseHash.Length + blockBytes.Length];
            Buffer.BlockCopy(baseHash, 0, combined, 0, baseHash.Length);
            Buffer.BlockCopy(blockBytes, 0, combined, baseHash.Length, blockBytes.Length);
            byte[] blockKey;
            using (var md5 = System.Security.Cryptography.MD5.Create())
            {
                blockKey = md5.ComputeHash(combined);
            }

            cipher.Initialize(blockKey);
            cipher.TransformBlock(encryptedPayload, 0, encryptedPayload.Length, encrypted, clearPrefix.Length);

            using var ms = new MemoryStream(encrypted);
            using var rc4 = new Rc4Stream(ms, baseHash, streamStartOffset: 0, useSha1: false, leaveOpen: true, clearPrefixLength: clearPrefix.Length);
            byte[] decrypted = new byte[plain.Length];
            int read = rc4.Read(decrypted, 0, decrypted.Length);

            Assert.Equal(plain.Length, read);
            Assert.Equal(plain, decrypted);
        }

        [Fact]
        public void Rc4Stream_SeekIntoMiddleOfBlock_DecryptsRemainingBytes()
        {
            byte[] baseHash = { 0x21, 0x32, 0x43, 0x54 };
            byte[] plain = new byte[1200];
            new Random(789).NextBytes(plain);

            byte[] encrypted = EncryptRc4Blocks(plain, baseHash);

            using var ms = new MemoryStream(encrypted);
            using var rc4 = new Rc4Stream(ms, baseHash, streamStartOffset: 0, useSha1: false, leaveOpen: true);

            rc4.Seek(700, SeekOrigin.Begin);
            byte[] actual = new byte[plain.Length - 700];
            int read = 0;
            while (read < actual.Length)
            {
                int chunk = rc4.Read(actual, read, Math.Min(73, actual.Length - read));
                if (chunk == 0)
                    break;

                read += chunk;
            }

            Assert.Equal(actual.Length, read);

            byte[] expected = new byte[plain.Length - 700];
            Buffer.BlockCopy(plain, 700, expected, 0, expected.Length);
            Assert.Equal(expected, actual);
        }

        [Fact]
        public void GetRc4BaseHash_BinaryRc4Header_ReturnsContextForValidPassword()
        {
            byte[] salt = { 0x10, 0x32, 0x54, 0x76, 0x98, 0xBA, 0xDC, 0xFE, 0x11, 0x22, 0x33, 0x44, 0x88, 0x99, 0xAA, 0xBB };
            byte[] verifier = { 0x01, 0x08, 0x02, 0x09, 0x03, 0x0A, 0x04, 0x0B, 0x05, 0x0C, 0x06, 0x0D, 0x07, 0x0E, 0x0F, 0x10 };
            using var tableStream = BuildBinaryRc4EncryptionHeader("secret", salt, verifier);

            var context = EncryptionHelper.GetRc4BaseHash(tableStream, lKey: 0, password: "secret");

            Assert.NotNull(context);
            Assert.False(context!.UseSha1);

            byte[] expectedBaseKey;
            using (var md5 = MD5.Create())
            {
                expectedBaseKey = md5.ComputeHash(Combine(Encoding.Unicode.GetBytes("secret"), salt));
            }

            Assert.Equal(expectedBaseKey, context.BaseKey);
        }

        [Fact]
        public void GetRc4BaseHash_MalformedHeader_EmitsWarningAndReturnsNull()
        {
            using var tableStream = new MemoryStream(new byte[] { 0x01, 0x00, 0x01, 0x00, 0xAA });
            var diagnostics = new List<ConversionDiagnostic>();

            using (Logger.BeginDiagnosticCapture(diagnostics))
            {
                var context = EncryptionHelper.GetRc4BaseHash(tableStream, lKey: 0, password: "secret");
                Assert.Null(context);
            }

            var diagnostic = Assert.Single(diagnostics);
            Assert.Equal(Logger.LogLevel.Warning, diagnostic.Level);
            Assert.Contains("Failed to parse RC4 encryption header", diagnostic.Message, StringComparison.Ordinal);
        }

        private static byte[] EncryptRc4Blocks(byte[] plain, byte[] baseHash)
        {
            byte[] encrypted = new byte[plain.Length];
            var cipher = new Rc4Cipher();

            for (uint block = 0; block * 512 < plain.Length; block++)
            {
                byte[] blockBytes = BitConverter.GetBytes(block);
                byte[] blockKey;
                using (var md5 = MD5.Create())
                {
                    blockKey = md5.ComputeHash(Combine(baseHash, blockBytes));
                }

                cipher.Initialize(blockKey);

                int offset = (int)(block * 512);
                int length = Math.Min(512, plain.Length - offset);
                cipher.TransformBlock(plain, offset, length, encrypted, offset);
            }

            return encrypted;
        }

        private static MemoryStream BuildBinaryRc4EncryptionHeader(string password, byte[] salt, byte[] verifier)
        {
            byte[] baseKey;
            using (var md5 = MD5.Create())
            {
                baseKey = md5.ComputeHash(Combine(Encoding.Unicode.GetBytes(password), salt));
            }

            byte[] verifierHash;
            using (var md5 = MD5.Create())
            {
                verifierHash = md5.ComputeHash(verifier);
            }

            var cipher = new Rc4Cipher();
            cipher.Initialize(baseKey);

            byte[] encryptedVerifier = new byte[16];
            cipher.TransformBlock(verifier, 0, verifier.Length, encryptedVerifier, 0);

            byte[] encryptedVerifierHash = new byte[16];
            cipher.TransformBlock(verifierHash, 0, verifierHash.Length, encryptedVerifierHash, 0);

            var stream = new MemoryStream();
            using (var writer = new BinaryWriter(stream, Encoding.Unicode, leaveOpen: true))
            {
                writer.Write((ushort)1);
                writer.Write((ushort)1);
                writer.Write(salt);
                writer.Write(encryptedVerifier);
                writer.Write(encryptedVerifierHash);
            }

            stream.Position = 0;
            return stream;
        }

        private static byte[] Combine(byte[] left, byte[] right)
        {
            var combined = new byte[left.Length + right.Length];
            Buffer.BlockCopy(left, 0, combined, 0, left.Length);
            Buffer.BlockCopy(right, 0, combined, left.Length, right.Length);
            return combined;
        }
    }
}

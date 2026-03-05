using System.IO;
using System.Security.Cryptography;
using System.Text;

namespace Nedev.DocToDocx.Utils;

/// <summary>
/// Decrypts XOR-encrypted streams from Word documents.
/// Implements the XOR obfuscation described in MS-DOC.
/// </summary>
public static class EncryptionHelper
{
    /// <summary>
    /// XOR decryption key derived from the document's LKey.
    /// </summary>
    private const uint DECRYPTION_KEY = 0xE1B0C1B2;

    /// <summary>
    /// Decrypts a stream using Word's XOR obfuscation.
    /// </summary>
    /// <param name="encryptedStream">The encrypted stream.</param>
    /// <param name="key">The XOR key (LKey from FIB).</param>
    /// <returns>A new stream with decrypted data.</returns>
    public static Stream DecryptXor(Stream encryptedStream, uint key)
    {
        var decryptedStream = new MemoryStream();
        
        // Read all bytes from the encrypted stream
        var buffer = new byte[4096];
        int bytesRead;
        
        while ((bytesRead = encryptedStream.Read(buffer, 0, buffer.Length)) > 0)
        {
            // XOR decrypt each byte
            for (int i = 0; i < bytesRead; i++)
            {
                buffer[i] ^= (byte)(key >> (i % 4) * 8);
            }
            
            decryptedStream.Write(buffer, 0, bytesRead);
        }
        
        decryptedStream.Position = 0;
        return decryptedStream;
    }

    /// <summary>
    /// Decrypts a byte array using Word's XOR obfuscation.
    /// </summary>
    /// <param name="encryptedBytes">The encrypted bytes.</param>
    /// <param name="key">The XOR key (LKey from FIB).</param>
    /// <returns>A new byte array with decrypted data.</returns>
    public static byte[] DecryptXor(byte[] encryptedBytes, uint key)
    {
        var decryptedBytes = new byte[encryptedBytes.Length];
        
        for (int i = 0; i < encryptedBytes.Length; i++)
        {
            decryptedBytes[i] = (byte)(encryptedBytes[i] ^ (byte)(key >> (i % 4) * 8));
        }
        
        return decryptedBytes;
    }

    /// <summary>
    /// Checks if a stream is encrypted using Word's XOR obfuscation.
    /// </summary>
    /// <param name="stream">The stream to check.</param>
    /// <param name="key">The XOR key (LKey from FIB).</param>
    /// <returns>True if the stream appears to be encrypted.</returns>
    public static bool IsXorEncrypted(Stream stream, uint key)
    {
        // Read first few bytes and check for common Word document signatures
        var buffer = new byte[1024];
        var originalPosition = stream.Position;
        
        _ = stream.Read(buffer, 0, Math.Min(buffer.Length, (int)(stream.Length - stream.Position)));
        stream.Position = originalPosition;
        
        // Check for common Word document magic numbers
        if (buffer.Length >= 2)
        {
            var magic = (ushort)(buffer[0] | (buffer[1] << 8));
            if (magic == 0xA5EC || magic == 0xA5B3)
            {
                return false; // Not encrypted or already decrypted
            }
        }
        
        // If we can't determine, assume it might be encrypted
        return true;
    }

    /// <summary>
    /// Attempts to verify the password and generate the base hash for RC4 CryptoAPI encryption.
    /// </summary>
    public static byte[]? GetRc4BaseHash(Stream tableStream, uint lKey, string? password)
    {
        password ??= "VelvetSweatshop";

        var originalPosition = tableStream.Position;
        try
        {
            tableStream.Position = 0;
            var reader = new BinaryReader(tableStream, Encoding.Unicode, true);

            ushort vMajor = reader.ReadUInt16();
            ushort vMinor = reader.ReadUInt16();

            if (vMajor != 2 || vMinor != 2)
            {
                // Not RC4 CryptoAPI
                return null;
            }

            uint flags = reader.ReadUInt32();
            uint sizeExtra = reader.ReadUInt32();
            uint algId = reader.ReadUInt32(); // 0x6801 for RC4
            uint hashAlg = reader.ReadUInt32(); // 0x8004 for SHA1
            uint keySize = reader.ReadUInt32();
            uint providerType = reader.ReadUInt32();
            uint reserved1 = reader.ReadUInt32();
            uint reserved2 = reader.ReadUInt32();

            // Read CSPName (null terminated unicode string)
            var cspNameBytes = new List<byte>();
            while (true)
            {
                byte b1 = reader.ReadByte();
                byte b2 = reader.ReadByte();
                if (b1 == 0 && b2 == 0) break;
                cspNameBytes.Add(b1);
                cspNameBytes.Add(b2);
            }
            // Skip padding if necessary. The EncryptionVerifier is 4-byte aligned.
            long currentPos = tableStream.Position;
            if (currentPos % 4 != 0)
            {
                tableStream.Position += 4 - (currentPos % 4);
            }

            // Read EncryptionVerifier
            uint saltSize = reader.ReadUInt32();
            byte[] salt = reader.ReadBytes(16); // Salt is always 16 bytes. SaltSize might be 16.
            if (salt.Length != 16) return null;
            
            byte[] encryptedVerifier = reader.ReadBytes(16);
            uint verifierHashSize = reader.ReadUInt32();
            byte[] encryptedVerifierHash = reader.ReadBytes(20); // SHA-1 is 20 bytes. Often padded to 20 or block size.

            // Build base hash: H0 = Hash(Salt + Password)
            byte[] passwordBytes = Encoding.Unicode.GetBytes(password);
            byte[] saltAndPassword = new byte[16 + passwordBytes.Length];
            Buffer.BlockCopy(salt, 0, saltAndPassword, 0, 16);
            Buffer.BlockCopy(passwordBytes, 0, saltAndPassword, 16, passwordBytes.Length);

            byte[] h0;
            using (var sha1 = SHA1.Create())
            {
                h0 = sha1.ComputeHash(saltAndPassword);
            }

            // Truncate/pad to KeySize (in bytes)
            int keySizeBytes = (int)(keySize / 8);
            byte[] baseKey = new byte[keySizeBytes];
            Buffer.BlockCopy(h0, 0, baseKey, 0, Math.Min(h0.Length, keySizeBytes));

            // Verify password using block 0
            // Block 0 key = Hash(baseKey + (uint)0)
            byte[] block0Key;
            using (var sha1 = SHA1.Create())
            {
                byte[] blockBytes = BitConverter.GetBytes((uint)0);
                sha1.TransformBlock(baseKey, 0, baseKey.Length, baseKey, 0);
                sha1.TransformFinalBlock(blockBytes, 0, blockBytes.Length);
                block0Key = sha1.Hash!;
            }

            // Truncate to KeySize
            byte[] actualKeyBlock0 = new byte[keySizeBytes];
            Buffer.BlockCopy(block0Key, 0, actualKeyBlock0, 0, Math.Min(block0Key.Length, keySizeBytes));

            // Decrypt verifier
            var rc4 = new Rc4Cipher();
            rc4.Initialize(actualKeyBlock0);
            byte[] decryptedVerifier = new byte[16];
            rc4.TransformBlock(encryptedVerifier, 0, 16, decryptedVerifier, 0);

            // Decrypt verifier hash
            byte[] decryptedVerifierHash = new byte[20];
            rc4.TransformBlock(encryptedVerifierHash, 0, 20, decryptedVerifierHash, 0);

            // Hash the decrypted verifier to see if it matches the decrypted verifier hash
            byte[] expectedHash;
            using (var sha1 = SHA1.Create())
            {
                expectedHash = sha1.ComputeHash(decryptedVerifier);
            }

            bool passwordCorrect = true;
            for (int i = 0; i < expectedHash.Length; i++)
            {
                if (decryptedVerifierHash[i] != expectedHash[i])
                {
                    passwordCorrect = false;
                    break;
                }
            }

            if (!passwordCorrect)
            {
                return null; // Invalid password
            }

            return baseKey;
        }
        catch
        {
            return null;
        }
        finally
        {
            tableStream.Position = originalPosition;
        }
    }
}
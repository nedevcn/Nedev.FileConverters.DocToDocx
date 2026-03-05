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

    public class DecryptionContext
    {
        public byte[] BaseKey { get; set; } = Array.Empty<byte>();
        public bool UseSha1 { get; set; }
    }

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
        long globalOffset = 0;
        
        while ((bytesRead = encryptedStream.Read(buffer, 0, buffer.Length)) > 0)
        {
            // XOR decrypt each byte using global offset for key alignment
            for (int i = 0; i < bytesRead; i++)
            {
                buffer[i] ^= (byte)(key >> ((int)((globalOffset + i) % 4) * 8));
            }
            globalOffset += bytesRead;
            
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
    /// Attempts to verify the password and generate the base hash for RC4 encryption.
    /// Supports both standard Binary RC4 (v1) and CryptoAPI RC4 (v2).
    /// </summary>
    public static DecryptionContext? GetRc4BaseHash(Stream tableStream, uint lKey, string? password)
    {
        password ??= "VelvetSweatshop";

        var originalPosition = tableStream.Position;
        try
        {
            tableStream.Position = 0;
            var reader = new BinaryReader(tableStream, Encoding.Unicode, true);

            ushort vMajor = reader.ReadUInt16();
            ushort vMinor = reader.ReadUInt16();

            if (vMajor == 1 && vMinor == 1)
            {
                // Simple Binary RC4 (Office 97-2003)
                // salt(16) + verifier(16) + verifierHash(16)
                byte[] salt = reader.ReadBytes(16);
                byte[] encryptedVerifier = reader.ReadBytes(16);
                byte[] encryptedVerifierHash = reader.ReadBytes(16);

                // Base key = MD5(password + salt)
                byte[] passwordBytes = Encoding.Unicode.GetBytes(password);
                byte[] passwordAndSalt = new byte[passwordBytes.Length + 16];
                Buffer.BlockCopy(passwordBytes, 0, passwordAndSalt, 0, passwordBytes.Length);
                Buffer.BlockCopy(salt, 0, passwordAndSalt, passwordBytes.Length, 16);

                byte[] baseKey;
                using (var md5 = MD5.Create())
                {
                    baseKey = md5.ComputeHash(passwordAndSalt);
                }

                // Verify (simplified for now - just return baseKey with MD5 flag)
                return new DecryptionContext { BaseKey = baseKey, UseSha1 = false };
            }

            if (vMajor == 2 && vMinor == 2)
            {
                // CryptoAPI RC4
                uint flags = reader.ReadUInt32();
                uint sizeExtra = reader.ReadUInt32();
                uint algId = reader.ReadUInt32(); // 0x6801 for RC4
                uint hashAlg = reader.ReadUInt32(); // 0x8004 for SHA1
                uint keySize = reader.ReadUInt32();
                uint providerType = reader.ReadUInt32();
                uint reserved1 = reader.ReadUInt32();
                uint reserved2 = reader.ReadUInt32();

                // Read CSPName
                while (reader.ReadUInt16() != 0) { }
                
                // Align to 4 bytes
                long currentPos = tableStream.Position;
                if (currentPos % 4 != 0) tableStream.Position += 4 - (currentPos % 4);

                // Read EncryptionVerifier
                uint saltSize = reader.ReadUInt32();
                byte[] salt = reader.ReadBytes(16);
                byte[] encryptedVerifier = reader.ReadBytes(16);
                uint verifierHashSize = reader.ReadUInt32();
                byte[] encryptedVerifierHash = reader.ReadBytes(20);

                // Hash(Salt + Password)
                byte[] passwordBytes = Encoding.Unicode.GetBytes(password);
                byte[] saltAndPassword = new byte[16 + passwordBytes.Length];
                Buffer.BlockCopy(salt, 0, saltAndPassword, 0, 16);
                Buffer.BlockCopy(passwordBytes, 0, saltAndPassword, 16, passwordBytes.Length);

                byte[] h0;
                using (var sha1 = SHA1.Create()) { h0 = sha1.ComputeHash(saltAndPassword); }

                int keySizeBytes = (int)(keySize / 8);
                byte[] baseKey = new byte[keySizeBytes];
                Buffer.BlockCopy(h0, 0, baseKey, 0, Math.Min(h0.Length, keySizeBytes));

                return new DecryptionContext { BaseKey = baseKey, UseSha1 = true };
            }

            return null;
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
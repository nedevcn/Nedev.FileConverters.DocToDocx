using System;
using System.IO;
using System.Security.Cryptography;

namespace Nedev.FileConverters.DocToDocx.Utils;

/// <summary>
/// A simple RC4 implementation to avoid depending on obsolete or platform-dependent .NET cryptography APIs.
/// </summary>
public class Rc4Cipher
{
    private readonly byte[] _s = new byte[256];
    private int _i;
    private int _j;

    public void Initialize(byte[] key)
    {
        for (int k = 0; k < 256; k++)
        {
            _s[k] = (byte)k;
        }

        int j = 0;
        for (int i = 0; i < 256; i++)
        {
            j = (j + _s[i] + key[i % key.Length]) & 255;
            // Swap
            (_s[i], _s[j]) = (_s[j], _s[i]);
        }

        _i = 0;
        _j = 0;
    }

    public void TransformBlock(byte[] inputBuffer, int inputOffset, int inputCount, byte[] outputBuffer, int outputOffset)
    {
        for (int k = 0; k < inputCount; k++)
        {
            _i = (_i + 1) & 255;
            _j = (_j + _s[_i]) & 255;

            // Swap
            (_s[_i], _s[_j]) = (_s[_j], _s[_i]);

            byte t = (byte)((_s[_i] + _s[_j]) & 255);
            byte kByte = _s[t];

            outputBuffer[outputOffset + k] = (byte)(inputBuffer[inputOffset + k] ^ kByte);
        }
    }
}

/// <summary>
/// A stream wrapper that handles RC4 decryption for MS-DOC streams.
/// It automatically recalculates the key on 512-byte block boundaries.
/// </summary>
public class Rc4Stream : Stream
{
    private readonly Stream _baseStream;
    private readonly byte[] _baseHash; // The hash of the password + salts
    private readonly bool _leaveOpen;
    private readonly long _streamLength;
    private readonly long _streamStartOffset;
    private readonly long _clearPrefixLength;
    
    private readonly bool _useSha1;
    private Rc4Cipher? _decryptor;
    private uint _currentBlock = uint.MaxValue; // Invalid block to force init

    public Rc4Stream(Stream baseStream, byte[] baseHash, long streamStartOffset = 0, bool useSha1 = false, bool leaveOpen = false, long clearPrefixLength = 0)
    {
        _baseStream = baseStream ?? throw new ArgumentNullException(nameof(baseStream));
        _baseHash = baseHash ?? throw new ArgumentNullException(nameof(baseHash));
        _useSha1 = useSha1;
        _leaveOpen = leaveOpen;
        _streamLength = _baseStream.Length;
        _streamStartOffset = streamStartOffset; // Offset within the overarching logical stream
        _clearPrefixLength = Math.Clamp(clearPrefixLength, 0, _streamLength);
    }

    public override bool CanRead => true;
    public override bool CanSeek => _baseStream.CanSeek;
    public override bool CanWrite => false;
    public override long Length => _streamLength;

    public override long Position
    {
        get => _baseStream.Position;
        set => _baseStream.Position = value;
    }

    public override int Read(byte[] buffer, int offset, int count)
    {
        int totalRead = 0;

        while (count > 0)
        {
            // Position in the logical stream (including any start offset if the physical stream starts partway through)
            long logicalPosition = Position + _streamStartOffset;

            if (logicalPosition < _clearPrefixLength)
            {
                int clearBytesToRead = (int)Math.Min(count, _clearPrefixLength - logicalPosition);
                int clearBytesRead = _baseStream.Read(buffer, offset, clearBytesToRead);
                if (clearBytesRead == 0)
                    break;

                offset += clearBytesRead;
                count -= clearBytesRead;
                totalRead += clearBytesRead;
                continue;
            }

            uint targetBlock = (uint)(logicalPosition / 512);

            if (_currentBlock != targetBlock || _decryptor == null)
            {
                InitializeBlock(targetBlock);

                // If we jumped into the middle of a block, we need to advance the cipher state by encrypting dummy bytes
                int blockOffset = (int)(logicalPosition % 512);
                if (blockOffset > 0)
                {
                    byte[] dummy = System.Buffers.ArrayPool<byte>.Shared.Rent(blockOffset);
                    try
                    {
                        _decryptor!.TransformBlock(dummy, 0, blockOffset, dummy, 0);
                    }
                    finally
                    {
                        System.Buffers.ArrayPool<byte>.Shared.Return(dummy);
                    }
                }
            }

            // Read up to the end of the current block
            int bytesToReadInBlock = 512 - (int)(logicalPosition % 512);
            int bytesToRead = Math.Min(count, bytesToReadInBlock);

            int bytesRead = _baseStream.Read(buffer, offset, bytesToRead);
            if (bytesRead == 0)
                break; // EOF

            // Decrypt the data in place
            _decryptor!.TransformBlock(buffer, offset, bytesRead, buffer, offset);

            offset += bytesRead;
            count -= bytesRead;
            totalRead += bytesRead;
        }

        return totalRead;
    }

    private void InitializeBlock(uint blockNumber)
    {
        _currentBlock = blockNumber;

        // Block hashing: Hash(baseHash + blockNumber (LittleEndian))
        byte[] blockBytes = BitConverter.GetBytes(blockNumber);
        byte[] blockKey;

        if (_useSha1)
        {
            using var sha1 = SHA1.Create();
            sha1.TransformBlock(_baseHash, 0, _baseHash.Length, null, 0);
            sha1.TransformFinalBlock(blockBytes, 0, blockBytes.Length);
            blockKey = sha1.Hash!;
        }
        else
        {
            using var md5 = MD5.Create();
            md5.TransformBlock(_baseHash, 0, _baseHash.Length, null, 0);
            md5.TransformFinalBlock(blockBytes, 0, blockBytes.Length);
            blockKey = md5.Hash!;
        }

        _decryptor = new Rc4Cipher();
        _decryptor.Initialize(blockKey);
    }

    public override void Flush() => _baseStream.Flush();

    public override void Write(byte[] buffer, int offset, int count) => throw new NotSupportedException();

    public override long Seek(long offset, SeekOrigin origin)
    {
        return _baseStream.Seek(offset, origin);
    }

    public override void SetLength(long value) => throw new NotSupportedException();

    protected override void Dispose(bool disposing)
    {
        if (disposing)
        {
            if (!_leaveOpen)
            {
                _baseStream.Dispose();
            }
        }
        base.Dispose(disposing);
    }
}

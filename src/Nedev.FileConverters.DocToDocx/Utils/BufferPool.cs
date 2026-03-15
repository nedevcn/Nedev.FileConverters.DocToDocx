using System.Buffers;

namespace Nedev.FileConverters.DocToDocx.Utils;

/// <summary>
/// Provides pooled buffer management to reduce memory allocations.
/// Wraps ArrayPool for common buffer sizes used throughout the converter.
/// </summary>
internal static class BufferPool
{
    // Common buffer sizes used in the converter
    private const int SmallBufferSize = 512;      // Sector size
    private const int MediumBufferSize = 4096;    // Page size
    private const int LargeBufferSize = 8192;     // Stream copy buffer
    private const int XLargeBufferSize = 65536;   // Large stream operations

    private static readonly ArrayPool<byte> SmallPool = ArrayPool<byte>.Create(SmallBufferSize, 100);
    private static readonly ArrayPool<byte> MediumPool = ArrayPool<byte>.Create(MediumBufferSize, 50);
    private static readonly ArrayPool<byte> LargePool = ArrayPool<byte>.Create(LargeBufferSize, 20);
    private static readonly ArrayPool<byte> XLargePool = ArrayPool<byte>.Create(XLargeBufferSize, 10);

    /// <summary>
    /// Rents a buffer of at least the specified size.
    /// </summary>
    /// <param name="minimumLength">Minimum buffer length required</param>
    /// <returns>Rented buffer (must be returned using Return)</returns>
    public static byte[] Rent(int minimumLength)
    {
        if (minimumLength <= SmallBufferSize)
            return SmallPool.Rent(minimumLength);
        if (minimumLength <= MediumBufferSize)
            return MediumPool.Rent(minimumLength);
        if (minimumLength <= LargeBufferSize)
            return LargePool.Rent(minimumLength);
        if (minimumLength <= XLargeBufferSize)
            return XLargePool.Rent(minimumLength);

        // For very large buffers, use the shared pool
        return ArrayPool<byte>.Shared.Rent(minimumLength);
    }

    /// <summary>
    /// Returns a rented buffer to the pool.
    /// </summary>
    /// <param name="buffer">Buffer to return</param>
    /// <param name="clearArray">Whether to clear the array before returning</param>
    public static void Return(byte[] buffer, bool clearArray = false)
    {
        if (buffer == null)
            return;

        if (buffer.Length <= SmallBufferSize)
            SmallPool.Return(buffer, clearArray);
        else if (buffer.Length <= MediumBufferSize)
            MediumPool.Return(buffer, clearArray);
        else if (buffer.Length <= LargeBufferSize)
            LargePool.Return(buffer, clearArray);
        else if (buffer.Length <= XLargeBufferSize)
            XLargePool.Return(buffer, clearArray);
        else
            ArrayPool<byte>.Shared.Return(buffer, clearArray);
    }

    /// <summary>
    /// Rents a buffer suitable for sector operations (512 bytes).
    /// </summary>
    public static byte[] RentSectorBuffer() => SmallPool.Rent(SmallBufferSize);

    /// <summary>
    /// Rents a buffer suitable for stream copy operations (8192 bytes).
    /// </summary>
    public static byte[] RentStreamBuffer() => LargePool.Rent(LargeBufferSize);

    /// <summary>
    /// Returns a sector buffer to the pool.
    /// </summary>
    public static void ReturnSectorBuffer(byte[] buffer) => SmallPool.Return(buffer, clearArray: false);

    /// <summary>
    /// Returns a stream buffer to the pool.
    /// </summary>
    public static void ReturnStreamBuffer(byte[] buffer) => LargePool.Return(buffer, clearArray: false);
}

/// <summary>
/// Disposable wrapper for rented buffers to ensure they are always returned.
/// </summary>
internal readonly ref struct RentedBuffer
{
    private readonly byte[] _buffer;
    private readonly int _length;
    private readonly bool _clearOnReturn;

    public RentedBuffer(int minimumLength, bool clearOnReturn = false)
    {
        _buffer = BufferPool.Rent(minimumLength);
        _length = minimumLength;
        _clearOnReturn = clearOnReturn;
    }

    public byte[] Buffer => _buffer;
    public Span<byte> Span => _buffer.AsSpan(0, _length);
    public Memory<byte> Memory => _buffer.AsMemory(0, _length);

    public void Dispose()
    {
        BufferPool.Return(_buffer, _clearOnReturn);
    }
}

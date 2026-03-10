using System.Diagnostics;
using System.Threading;

namespace Nedev.FileConverters.DocToDocx.Utils;

/// <summary>
/// Simple logging utility for the converter.
/// Provides structured logging with different severity levels.
/// </summary>
public static class Logger
{
    private sealed class CaptureScopeState
    {
        public CaptureScopeState(CaptureScopeState? parent, ICollection<string>? warnings, ICollection<ConversionDiagnostic>? diagnostics)
        {
            Parent = parent;
            Warnings = warnings;
            Diagnostics = diagnostics;
        }

        public CaptureScopeState? Parent { get; }
        public ICollection<string>? Warnings { get; }
        public ICollection<ConversionDiagnostic>? Diagnostics { get; }
    }

    private sealed class CaptureScope : IDisposable
    {
        private readonly CaptureScopeState? _previous;
        private bool _disposed;

        public CaptureScope(CaptureScopeState? previous)
        {
            _previous = previous;
        }

        public void Dispose()
        {
            if (_disposed)
                return;

            _captureState.Value = _previous;
            _disposed = true;
        }
    }

    private static readonly AsyncLocal<CaptureScopeState?> _captureState = new();

    /// <summary>
    /// Log level enumeration
    /// </summary>
    public enum LogLevel
    {
        Debug,
        Info,
        Warning,
        Error,
        Fatal
    }

    /// <summary>
    /// Current minimum log level
    /// </summary>
    public static LogLevel MinimumLevel { get; set; } = LogLevel.Info;

    /// <summary>
    /// Whether to include timestamps in log messages
    /// </summary>
    public static bool IncludeTimestamp { get; set; } = true;

    /// <summary>
    /// Custom log handler (if null, logs to Debug output)
    /// </summary>
    public static Action<LogLevel, string>? CustomHandler { get; set; }

    /// <summary>
    /// Begins capturing warning-or-higher log messages into the supplied collection for the current async flow.
    /// </summary>
    public static IDisposable BeginWarningCapture(ICollection<string> warnings)
    {
        if (warnings == null)
            throw new ArgumentNullException(nameof(warnings));

        var previous = _captureState.Value;
        _captureState.Value = new CaptureScopeState(previous, warnings, diagnostics: null);
        return new CaptureScope(previous);
    }

    /// <summary>
    /// Begins capturing warning-or-higher log messages as structured diagnostics for the current async flow.
    /// </summary>
    public static IDisposable BeginDiagnosticCapture(ICollection<ConversionDiagnostic> diagnostics)
    {
        if (diagnostics == null)
            throw new ArgumentNullException(nameof(diagnostics));

        var previous = _captureState.Value;
        _captureState.Value = new CaptureScopeState(previous, warnings: null, diagnostics);
        return new CaptureScope(previous);
    }

    /// <summary>
    /// Logs a debug message
    /// </summary>
    public static void Debug(string message)
    {
        Log(LogLevel.Debug, message);
    }

    /// <summary>
    /// Logs an info message
    /// </summary>
    public static void Info(string message)
    {
        Log(LogLevel.Info, message);
    }

    /// <summary>
    /// Logs a warning message
    /// </summary>
    public static void Warning(string message)
    {
        Log(LogLevel.Warning, message);
    }

    /// <summary>
    /// Logs a warning message with exception details
    /// </summary>
    public static void Warning(string message, Exception? exception)
    {
        Log(LogLevel.Warning, message, exception);
    }

    /// <summary>
    /// Logs an error message
    /// </summary>
    public static void Error(string message)
    {
        Log(LogLevel.Error, message);
    }

    /// <summary>
    /// Logs an error message with exception details
    /// </summary>
    public static void Error(string message, Exception? exception)
    {
        Log(LogLevel.Error, message, exception);
    }

    /// <summary>
    /// Logs a fatal error message
    /// </summary>
    public static void Fatal(string message)
    {
        Log(LogLevel.Fatal, message);
    }

    /// <summary>
    /// Logs a fatal error message with exception details
    /// </summary>
    public static void Fatal(string message, Exception? exception)
    {
        Log(LogLevel.Fatal, message, exception);
    }

    /// <summary>
    /// Main logging method
    /// </summary>
    private static void Log(LogLevel level, string message, Exception? exception = null)
    {
        if (level < MinimumLevel)
            return;

        var formattedMessage = FormatMessage(level, message, exception);

        if (level >= LogLevel.Warning)
        {
            var timestamp = DateTime.UtcNow;
            var captureState = _captureState.Value;
            captureState?.Warnings?.Add(formattedMessage);
            captureState?.Diagnostics?.Add(new ConversionDiagnostic(
                timestamp,
                level,
                message,
                formattedMessage,
                exception?.GetType().Name,
                exception?.Message));
        }

        if (CustomHandler != null)
        {
            CustomHandler(level, formattedMessage);
        }
        else
        {
            // Default: write to debug output
            System.Diagnostics.Debug.WriteLine(formattedMessage);

            // Also write to console so tests and CLI output can see it
            try { Console.WriteLine(formattedMessage); } catch { /* ignore */ }
        }
    }

    /// <summary>
    /// Formats a log message
    /// </summary>
    private static string FormatMessage(LogLevel level, string message, Exception? exception)
    {
        var timestamp = IncludeTimestamp ? $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] " : "";
        var levelStr = $"[{level.ToString().ToUpperInvariant()}]";
        var result = $"{timestamp}{levelStr} {message}";

        if (exception != null)
        {
            result += $"\n  Exception: {exception.GetType().Name}: {exception.Message}";
            if (exception.StackTrace != null)
            {
                result += $"\n  StackTrace: {exception.StackTrace}";
            }
        }

        return result;
    }
}

/// <summary>
/// Structured non-fatal diagnostic captured during conversion.
/// </summary>
public sealed class ConversionDiagnostic
{
    public ConversionDiagnostic(DateTime timestampUtc, Logger.LogLevel level, string message, string formattedMessage, string? exceptionType, string? exceptionMessage)
    {
        TimestampUtc = timestampUtc;
        Level = level;
        Message = message ?? throw new ArgumentNullException(nameof(message));
        FormattedMessage = formattedMessage ?? throw new ArgumentNullException(nameof(formattedMessage));
        ExceptionType = exceptionType;
        ExceptionMessage = exceptionMessage;
    }

    public DateTime TimestampUtc { get; }
    public Logger.LogLevel Level { get; }
    public string Message { get; }
    public string FormattedMessage { get; }
    public string? ExceptionType { get; }
    public string? ExceptionMessage { get; }
}

/// <summary>
/// Extension methods for exception handling
/// </summary>
public static class ExceptionExtensions
{
    /// <summary>
    /// Logs an exception and returns a user-friendly error message
    /// </summary>
    public static string LogAndGetMessage(this Exception exception, string context)
    {
        Logger.Error(context, exception);
        return $"{context}: {exception.Message}";
    }

    /// <summary>
    /// Gets the innermost exception message
    /// </summary>
    public static string GetInnermostMessage(this Exception exception)
    {
        var ex = exception;
        while (ex.InnerException != null)
        {
            ex = ex.InnerException;
        }
        return ex.Message;
    }

    /// <summary>
    /// Gets a detailed exception string including all inner exceptions
    /// </summary>
    public static string GetDetailedMessage(this Exception exception)
    {
        var messages = new List<string>();
        var ex = exception;
        var level = 0;

        while (ex != null)
        {
            var indent = new string(' ', level * 2);
            messages.Add($"{indent}{ex.GetType().Name}: {ex.Message}");
            ex = ex.InnerException;
            level++;
        }

        return string.Join("\n", messages);
    }
}

/// <summary>
/// Result type for operations that can succeed or fail
    /// </summary>
public class Result<T>
{
    public bool IsSuccess { get; }
    public T? Value { get; }
    public string? Error { get; }
    public Exception? Exception { get; }

    private Result(bool isSuccess, T? value, string? error, Exception? exception)
    {
        IsSuccess = isSuccess;
        Value = value;
        Error = error;
        Exception = exception;
    }

    public static Result<T> Success(T value)
    {
        return new Result<T>(true, value, null, null);
    }

    public static Result<T> Failure(string error)
    {
        return new Result<T>(false, default, error, null);
    }

    public static Result<T> Failure(Exception exception)
    {
        return new Result<T>(false, default, exception.Message, exception);
    }

    public static Result<T> Failure(string error, Exception exception)
    {
        return new Result<T>(false, default, error, exception);
    }

    /// <summary>
    /// Maps the result value to another type
    /// </summary>
    public Result<TResult> Map<TResult>(Func<T, TResult> mapper)
    {
        return IsSuccess
            ? Result<TResult>.Success(mapper(Value!))
            : Result<TResult>.Failure(Error ?? "Unknown error", Exception!);
    }
}

/// <summary>
/// Non-generic result type
/// </summary>
public class Result
{
    public bool IsSuccess { get; }
    public string? Error { get; }
    public Exception? Exception { get; }

    private Result(bool isSuccess, string? error, Exception? exception)
    {
        IsSuccess = isSuccess;
        Error = error;
        Exception = exception;
    }

    public static Result Success()
    {
        return new Result(true, null, null);
    }

    public static Result Failure(string error)
    {
        return new Result(false, error, null);
    }

    public static Result Failure(Exception exception)
    {
        return new Result(false, exception.Message, exception);
    }

    public static Result Failure(string error, Exception exception)
    {
        return new Result(false, error, exception);
    }
}

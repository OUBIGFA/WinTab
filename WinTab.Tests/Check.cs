using System;
using System.Collections.Generic;

internal static class Check
{
    public static void That(bool condition, string message)
    {
        if (!condition)
            throw new InvalidOperationException(message);
    }

    public static void Equal<T>(T expected, T actual, string? message = null)
    {
        if (EqualityComparer<T>.Default.Equals(expected, actual))
            return;

        var prefix = string.IsNullOrEmpty(message) ? string.Empty : $"{message} ";
        throw new InvalidOperationException($"{prefix}expected '{expected}', got '{actual}'");
    }

    public static void EqualIgnoreCase(string expected, string actual, string? message = null)
    {
        if (StringComparer.OrdinalIgnoreCase.Equals(expected, actual))
            return;

        var prefix = string.IsNullOrEmpty(message) ? string.Empty : $"{message} ";
        throw new InvalidOperationException($"{prefix}expected '{expected}', got '{actual}'");
    }

    public static TException Throws<TException>(Action action, string message) where TException : Exception
    {
        try
        {
            action();
        }
        catch (TException ex)
        {
            return ex;
        }

        throw new InvalidOperationException(message);
    }
}

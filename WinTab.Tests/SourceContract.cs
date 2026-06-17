using System;
using System.IO;
using System.Linq;

internal static class SourceContract
{
    public static string FindRepoFile(params string[] relativeParts)
    {
        foreach (var start in new[] { Environment.CurrentDirectory, AppContext.BaseDirectory })
        {
            var current = new DirectoryInfo(start);
            while (current != null)
            {
                var candidate = Path.Combine(new[] { current.FullName }.Concat(relativeParts).ToArray());
                if (File.Exists(candidate))
                    return candidate;

                current = current.Parent;
            }
        }

        throw new FileNotFoundException("Could not locate repo file.", Path.Combine(relativeParts));
    }

    public static string ExtractMethodBody(string source, string signature)
    {
        var signatureIndex = source.IndexOf(signature, StringComparison.Ordinal);
        if (signatureIndex < 0)
            throw new InvalidOperationException($"Could not find method signature '{signature}'.");

        var openBraceIndex = source.IndexOf('{', signatureIndex);
        if (openBraceIndex < 0)
            throw new InvalidOperationException($"Could not find method body for '{signature}'.");

        var depth = 0;
        for (var i = openBraceIndex; i < source.Length; i++)
        {
            if (source[i] == '{')
                depth++;
            else if (source[i] == '}')
                depth--;

            if (depth == 0)
                return source.Substring(openBraceIndex, i - openBraceIndex + 1);
        }

        throw new InvalidOperationException($"Could not parse method body for '{signature}'.");
    }
}

namespace WinTab.Platform.Win32;

public static class ShellNamespacePath
{
    public const string RecycleBinGuidBraced = "{645FF040-5081-101B-9F08-00AA002F954E}";
    public const string RecycleBinGuidPath = "::" + RecycleBinGuidBraced;
    public const string RecycleBinShellAlias = "shell:RecycleBinFolder";

    public static bool IsShellNamespace(string? location)
    {
        if (string.IsNullOrWhiteSpace(location))
            return false;

        string token = location.Trim();
        if (token.StartsWith("::", StringComparison.Ordinal))
            return true;

        if (token.StartsWith("shell:", StringComparison.OrdinalIgnoreCase))
            return true;

        return IsBracedGuid(token);
    }

    public static bool TryNormalizeToken(string? value, out string normalized)
    {
        normalized = string.Empty;

        if (string.IsNullOrWhiteSpace(value))
            return false;

        string token = value.Trim();

        if (token.StartsWith("::", StringComparison.Ordinal))
        {
            normalized = token;
            return true;
        }

        if (token.StartsWith("shell:", StringComparison.OrdinalIgnoreCase))
        {
            if (token.StartsWith("shell::", StringComparison.OrdinalIgnoreCase) &&
                !token.StartsWith("shell:::", StringComparison.OrdinalIgnoreCase))
            {
                string remainder = token[7..].TrimStart(':');
                if (!string.IsNullOrWhiteSpace(remainder))
                {
                    normalized = "shell:" + remainder;
                    return true;
                }
            }

            normalized = token;
            return true;
        }

        if (IsBracedGuid(token))
        {
            normalized = "::" + token;
            return true;
        }

        return false;
    }

    public static IReadOnlyList<string> BuildNamespaceCandidates(string location)
    {
        string token = location.Trim();
        var candidates = new List<string>(4);

        static void AddCandidate(List<string> list, string? candidate)
        {
            if (string.IsNullOrWhiteSpace(candidate))
                return;

            if (!list.Contains(candidate, StringComparer.OrdinalIgnoreCase))
                list.Add(candidate);
        }

        AddCandidate(candidates, token);

        if (IsBracedGuid(token))
        {
            AddCandidate(candidates, "::" + token);
            AddCandidate(candidates, "shell:::" + token);
            return candidates;
        }

        if (token.StartsWith("::", StringComparison.Ordinal))
        {
            AddCandidate(candidates, "shell:::" + token[2..]);
            return candidates;
        }

        if (token.StartsWith("shell:::", StringComparison.OrdinalIgnoreCase))
        {
            AddCandidate(candidates, "::" + token[8..]);
            return candidates;
        }

        if (token.StartsWith("shell::", StringComparison.OrdinalIgnoreCase))
        {
            string remainder = token[7..].TrimStart(':');
            AddCandidate(candidates, "shell:" + remainder);
            return candidates;
        }

        if (token.StartsWith("shell:", StringComparison.OrdinalIgnoreCase))
        {
            string remainder = token[6..].TrimStart(':');
            if (IsBracedGuid(remainder))
            {
                AddCandidate(candidates, "::" + remainder);
                AddCandidate(candidates, "shell:::" + remainder);
            }
        }

        return candidates;
    }

    public static bool IsBracedGuid(string value)
    {
        if (value.Length < 2 || value[0] != '{' || value[^1] != '}')
            return false;

        return Guid.TryParse(value, out _);
    }
}

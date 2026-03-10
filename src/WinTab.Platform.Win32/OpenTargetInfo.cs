namespace WinTab.Platform.Win32;

public sealed record OpenTargetInfo(string RawTarget, string NormalizedTarget, OpenTargetKind Kind)
{
    public bool IsValid => Kind != OpenTargetKind.Invalid;
    public bool IsPhysicalFileSystem => Kind == OpenTargetKind.PhysicalFileSystem;
    public bool IsShellNamespace => Kind is OpenTargetKind.ShellNamespace or OpenTargetKind.NativeShellNamespace;
    public bool RequiresNativeShellLaunch => Kind == OpenTargetKind.NativeShellNamespace;
}

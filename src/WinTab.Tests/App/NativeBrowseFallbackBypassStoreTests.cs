using System;
using FluentAssertions;
using WinTab.App.ExplorerTabUtilityPort;
using WinTab.Platform.Win32;
using Xunit;

namespace WinTab.Tests.App;

public sealed class NativeBrowseFallbackBypassStoreTests
{
    [Fact]
    public void Register_ThenTryConsumeSameLocation_ShouldSucceedOnce()
    {
        var store = new NativeBrowseFallbackBypassStore(new ShellLocationIdentityService());

        store.Register(@"C:\Users\Public\Desktop", TimeSpan.FromSeconds(3));

        store.TryConsume(@"c:\users\public\desktop\").Should().BeTrue();
        store.TryConsume(@"C:\Users\Public\Desktop").Should().BeFalse();
    }

    [Fact]
    public void Register_ThenTryConsumeDifferentLocation_ShouldNotMatch()
    {
        var store = new NativeBrowseFallbackBypassStore(new ShellLocationIdentityService());

        store.Register(@"C:\Users\Public\Desktop", TimeSpan.FromSeconds(3));

        store.TryConsume(@"C:\Windows").Should().BeFalse();
    }

    [Fact]
    public void RegisterWithExpiredTtl_ShouldNotBeConsumable()
    {
        var store = new NativeBrowseFallbackBypassStore(new ShellLocationIdentityService());

        store.Register(@"C:\Users\Public\Desktop", TimeSpan.Zero);

        store.TryConsume(@"C:\Users\Public\Desktop").Should().BeFalse();
    }

    [Fact]
    public void Revoke_ShouldRemoveBypassToken()
    {
        var store = new NativeBrowseFallbackBypassStore(new ShellLocationIdentityService());

        store.Register(@"C:\Users\Public\Desktop", TimeSpan.FromSeconds(3));
        store.Revoke(@"C:\Users\Public\Desktop");

        store.TryConsume(@"C:\Users\Public\Desktop").Should().BeFalse();
    }
}

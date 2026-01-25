using System.ComponentModel;
using System.Text.Json.Serialization;
using WinTab.Helpers;

namespace WinTab.Models;

[JsonConverter(typeof(HotKeyActionJsonConverter))]
public enum HotKeyAction
{
    [Description("HotKeyActionDesc_Open")]
    Open,
    [Description("HotKeyActionDesc_Duplicate")]
    Duplicate,
    [Description("HotKeyActionDesc_ReopenClosed")]
    ReopenClosed,
    [Description("HotKeyActionDesc_TabSearch")]
    TabSearch,
    [Description("HotKeyActionDesc_NavigateBack")]
    NavigateBack,
    [Description("HotKeyActionDesc_NavigateUp")]
    NavigateUp,
    [Description("HotKeyActionDesc_NavigateForward")]
    NavigateForward,
    [Description("HotKeyActionDesc_SetTargetWindow")]
    SetTargetWindow,
    [Description("HotKeyActionDesc_ToggleWinHook")]
    ToggleWinHook,
    [Description("HotKeyActionDesc_ToggleReuseTabs")]
    ToggleReuseTabs,
    [Description("HotKeyActionDesc_ToggleVisibility")]
    ToggleVisibility,
    [Description("HotKeyActionDesc_DetachTab")]
    DetachTab,
    [Description("HotKeyActionDesc_SnapRight")]
    SnapRight,
    [Description("HotKeyActionDesc_SnapLeft")]
    SnapLeft,
    [Description("HotKeyActionDesc_SnapUp")]
    SnapUp,
    [Description("HotKeyActionDesc_SnapDown")]
    SnapDown
}


#pragma once
#define STRICT
#define WIN32_LEAN_AND_MEAN
#define NOMINMAX
#include <Windows.h>
#include <shellapi.h>
#include <CommCtrl.h>
#include <commoncontrols.h>
#include <atlstr.h>
#include "Rad/Log.h"
#include "Rad/MemoryPlus.h"

inline LONG Width(const RECT r)
{
    return r.right - r.left;
}

inline LONG Height(const RECT r)
{
    return r.bottom - r.top;
}

inline void SetWindowAccentDisabled(HWND hWnd)
{
    const HINSTANCE hModule = GetModuleHandle(TEXT("user32.dll"));
    if (hModule)
    {
        typedef enum _ACCENT_STATE {
            ACCENT_DISABLED,
            ACCENT_ENABLE_GRADIENT,
            ACCENT_ENABLE_TRANSPARENTGRADIENT,
            ACCENT_ENABLE_BLURBEHIND,
            ACCENT_ENABLE_ACRYLICBLURBEHIND,
            ACCENT_INVALID_STATE
        } ACCENT_STATE;
        struct ACCENTPOLICY
        {
            ACCENT_STATE nAccentState;
            DWORD nFlags;
            DWORD nColor;
            DWORD nAnimationId;
        };
        typedef enum _WINDOWCOMPOSITIONATTRIB {
            WCA_UNDEFINED = 0,
            WCA_NCRENDERING_ENABLED = 1,
            WCA_NCRENDERING_POLICY = 2,
            WCA_TRANSITIONS_FORCEDISABLED = 3,
            WCA_ALLOW_NCPAINT = 4,
            WCA_CAPTION_BUTTON_BOUNDS = 5,
            WCA_NONCLIENT_RTL_LAYOUT = 6,
            WCA_FORCE_ICONIC_REPRESENTATION = 7,
            WCA_EXTENDED_FRAME_BOUNDS = 8,
            WCA_HAS_ICONIC_BITMAP = 9,
            WCA_THEME_ATTRIBUTES = 10,
            WCA_NCRENDERING_EXILED = 11,
            WCA_NCADORNMENTINFO = 12,
            WCA_EXCLUDED_FROM_LIVEPREVIEW = 13,
            WCA_VIDEO_OVERLAY_ACTIVE = 14,
            WCA_FORCE_ACTIVEWINDOW_APPEARANCE = 15,
            WCA_DISALLOW_PEEK = 16,
            WCA_CLOAK = 17,
            WCA_CLOAKED = 18,
            WCA_ACCENT_POLICY = 19,
            WCA_FREEZE_REPRESENTATION = 20,
            WCA_EVER_UNCLOAKED = 21,
            WCA_VISUAL_OWNER = 22,
            WCA_LAST = 23
        } WINDOWCOMPOSITIONATTRIB;
        struct WINCOMPATTRDATA
        {
            WINDOWCOMPOSITIONATTRIB nAttribute;
            PVOID pData;
            ULONG ulDataSize;
        };
        typedef BOOL(WINAPI* pSetWindowCompositionAttribute)(HWND, WINCOMPATTRDATA*);
        const pSetWindowCompositionAttribute SetWindowCompositionAttribute = (pSetWindowCompositionAttribute) GetProcAddress(hModule, "SetWindowCompositionAttribute");
        if (SetWindowCompositionAttribute)
        {
            ACCENTPOLICY policy = { ACCENT_DISABLED, 0, 0, 0 };
            WINCOMPATTRDATA data = { WCA_ACCENT_POLICY, &policy, sizeof(ACCENTPOLICY) };
            SetWindowCompositionAttribute(hWnd, &data);
            //DwmSetWindowAttribute(hWnd, WCA_ACCENT_POLICY, &policy, sizeof(ACCENTPOLICY));
        }
        //FreeLibrary(hModule);
    }
}

inline void SetWindowBlur(HWND hWnd, bool bEnabled)
{
    const HINSTANCE hModule = GetModuleHandle(TEXT("user32.dll"));
    if (hModule)
    {
        typedef enum _ACCENT_STATE {
            ACCENT_DISABLED,
            ACCENT_ENABLE_GRADIENT,
            ACCENT_ENABLE_TRANSPARENTGRADIENT,
            ACCENT_ENABLE_BLURBEHIND,
            ACCENT_ENABLE_ACRYLICBLURBEHIND,
            ACCENT_INVALID_STATE
        } ACCENT_STATE;
        struct ACCENTPOLICY
        {
            ACCENT_STATE nAccentState;
            DWORD nFlags;
            DWORD nColor;
            DWORD nAnimationId;
        };
        typedef enum _WINDOWCOMPOSITIONATTRIB {
            WCA_UNDEFINED = 0,
            WCA_NCRENDERING_ENABLED = 1,
            WCA_NCRENDERING_POLICY = 2,
            WCA_TRANSITIONS_FORCEDISABLED = 3,
            WCA_ALLOW_NCPAINT = 4,
            WCA_CAPTION_BUTTON_BOUNDS = 5,
            WCA_NONCLIENT_RTL_LAYOUT = 6,
            WCA_FORCE_ICONIC_REPRESENTATION = 7,
            WCA_EXTENDED_FRAME_BOUNDS = 8,
            WCA_HAS_ICONIC_BITMAP = 9,
            WCA_THEME_ATTRIBUTES = 10,
            WCA_NCRENDERING_EXILED = 11,
            WCA_NCADORNMENTINFO = 12,
            WCA_EXCLUDED_FROM_LIVEPREVIEW = 13,
            WCA_VIDEO_OVERLAY_ACTIVE = 14,
            WCA_FORCE_ACTIVEWINDOW_APPEARANCE = 15,
            WCA_DISALLOW_PEEK = 16,
            WCA_CLOAK = 17,
            WCA_CLOAKED = 18,
            WCA_ACCENT_POLICY = 19,
            WCA_FREEZE_REPRESENTATION = 20,
            WCA_EVER_UNCLOAKED = 21,
            WCA_VISUAL_OWNER = 22,
            WCA_LAST = 23
        } WINDOWCOMPOSITIONATTRIB;
        struct WINCOMPATTRDATA
        {
            WINDOWCOMPOSITIONATTRIB nAttribute;
            PVOID pData;
            ULONG ulDataSize;
        };
        typedef BOOL(WINAPI* pSetWindowCompositionAttribute)(HWND, WINCOMPATTRDATA*);
        const pSetWindowCompositionAttribute SetWindowCompositionAttribute = (pSetWindowCompositionAttribute) GetProcAddress(hModule, "SetWindowCompositionAttribute");
        if (SetWindowCompositionAttribute)
        {
            ACCENTPOLICY policy = { bEnabled ? ACCENT_ENABLE_ACRYLICBLURBEHIND : ACCENT_DISABLED, 0, 0x80000000, 0 };
            //ACCENTPOLICY policy = { ACCENT_ENABLE_BLURBEHIND };
            WINCOMPATTRDATA data = { WCA_ACCENT_POLICY, &policy, sizeof(ACCENTPOLICY) };
            SetWindowCompositionAttribute(hWnd, &data);
            //DwmSetWindowAttribute(hWnd, WCA_ACCENT_POLICY, &policy, sizeof(ACCENTPOLICY));
        }
        //FreeLibrary(hModule);
    }
}

inline SIZE GetFontSize(const HWND hWnd, const HANDLE hFont, _In_reads_(c) LPCTSTR lpString, _In_ int c)
{
    SIZE TextSize = {};
    HDC hDC = ::GetDC(hWnd);
    const HANDLE hOldFont = SelectObject(hDC, hFont);
    GetTextExtentPoint32(hDC, lpString, c, &TextSize);
    SelectObject(hDC, hOldFont);
    ReleaseDC(hWnd, hDC);
    return TextSize;
}

inline BOOL ScreenToClient(_In_ HWND hWnd, _Inout_ LPRECT lpRect)
{
    _ASSERT(::IsWindow(hWnd));
    if (!::ScreenToClient(hWnd, (LPPOINT) lpRect))
        return FALSE;
    return ::ScreenToClient(hWnd, ((LPPOINT) lpRect) + 1);
}

inline BOOL SetMenuItemBitmap(_In_ HMENU hMenu, _In_ UINT item, _In_ BOOL fByPosition, _In_ HBITMAP hBitmap)
{
    MENUITEMINFO minfo;
    minfo.cbSize = sizeof(minfo);
    minfo.fMask = MIIM_BITMAP;
    minfo.hbmpItem = hBitmap;
    return SetMenuItemInfo(hMenu, item, fByPosition, &minfo);
}

inline BOOL SetMenuItemData(_In_ HMENU hMenu, _In_ UINT item, _In_ BOOL fByPosition, _In_ ULONG_PTR dwItemData)
{
    MENUITEMINFO minfo;
    minfo.cbSize = sizeof(minfo);
    minfo.fMask = MIIM_DATA;
    minfo.dwItemData = dwItemData;
    return SetMenuItemInfo(hMenu, item, fByPosition, &minfo);
}

inline BOOL GetMenuItemData(_In_ HMENU hMenu, _In_ UINT item, _In_ BOOL fByPosition, _Out_ ULONG_PTR& dwItemData)
{
    dwItemData = 0;
    MENUITEMINFO minfo;
    minfo.cbSize = sizeof(minfo);
    minfo.fMask = MIIM_DATA;
    if (!GetMenuItemInfo(hMenu, item, fByPosition, &minfo))
        return FALSE;
    dwItemData = minfo.dwItemData;
    return TRUE;
}

inline CString ExpandEnvironmentStrings(LPCTSTR s)
{
    CString r;
    CHECK(ExpandEnvironmentStrings(s, r.GetBufferSetLength(1024), 1024));
    r.ReleaseBuffer();
    return r;
}

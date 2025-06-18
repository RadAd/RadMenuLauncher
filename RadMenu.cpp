#include "Rad/Window.h"
#include "Rad/Dialog.h"
#include "Rad/Windowxx.h"
#include "Rad/Log.h"
#include "Rad/Format.h"
#include <tchar.h>
//#include <strsafe.h>
#include "resource.h"

//#include <atlcomcli.h>
#include <atlbase.h>
#include <shlobj.h>
#include <Shlwapi.h>

#include "EditPlus.h"
#include "ListBoxPlus.h"
#include "StrUtils.h"
#include "ShellUtils.h"
#include "JumpListMenu.h"
#include "JumpListMenuHandler.h"
//#include "resource.h"

#include <string>
#include <vector>

// TODO
// Tooltips

#define HK_OPEN                 1

inline HFONT CreateBoldFont(HFONT hFont)
{
    LOGFONT lf = {};
    GetObject(hFont, sizeof(LOGFONT), &lf);
    lf.lfWeight = FW_BOLD;
    return CreateFontIndirect(&lf);
}

template <class T, class Allocator>
const T* get(const CHeapPtrBase<T, Allocator>& p)
{
    return p;
}

Theme g_Theme;

// #define UM_ADDITEM  (WM_USER + 127)

extern HINSTANCE g_hInstance;
extern HWND g_hWndDlg;

template <class R, class F, typename... Ps>
R CallWinApi(F f, Ps... args)
{
    R r = {};
    if (!f(args..., &r))
        throw WinError({ GetLastError(), nullptr, TEXT("CallWinApi") });
    return r;
}

DWORD ReadFromHandle(HANDLE hReadPipe, _Out_writes_to_opt_(nSize, return +1) LPWSTR lpBuffer, _In_ DWORD nSize)
{
    if (lpBuffer == nullptr)
        return 0;

    DWORD dwReadTotal = 0;

    for (;;)
    {
        BYTE Buffer[1024];
        DWORD dwRead;
        if (!ReadFile(hReadPipe, &Buffer, ARRAYSIZE(Buffer), &dwRead, NULL))
            break;

        // TODO Do this once after we get all text
        INT iResult = 0;
        if (IsTextUnicode(Buffer, dwRead, &iResult))
        {
            memcpy_s(lpBuffer + dwReadTotal, nSize - dwReadTotal, Buffer, dwRead);
            dwReadTotal += dwRead / sizeof(WCHAR);
        }
        else
            dwReadTotal += MultiByteToWideChar(CP_UTF8, 0, (LPCSTR)&Buffer, dwRead, lpBuffer + dwReadTotal, nSize - dwReadTotal);
    }
    lpBuffer[dwReadTotal] = L'\0';
    return dwReadTotal;
}

#define IDC_EDIT                     100
#define IDC_LIST                     101

const int Border = 10;

LRESULT CALLBACK BuddyProc(HWND hWnd, UINT uMsg, WPARAM wParam, LPARAM lParam, UINT_PTR uIdSubclass, DWORD_PTR dwRefData)
{
    switch (uMsg)
    {
    case WM_KEYDOWN:
    case WM_KEYUP:
    {
        const UINT vk = (UINT) (wParam);
        const int cRepeat = (int) (short) LOWORD(lParam);
        UINT flags = (UINT) HIWORD(lParam);
        if ((GetKeyState(VK_SHIFT) & 0x8000) == 0)
        {
            switch (vk)
            {
            case VK_UP:         case VK_DOWN:
            case VK_HOME:       case VK_END:
            case VK_NEXT:       case VK_PRIOR:
            {
                HWND hWndBuddy = GetWindow(hWnd, GW_HWNDNEXT);
                SendMessage(hWndBuddy, uMsg, wParam, lParam);
                return TRUE;
            }
            }
        }
        return DefSubclassProc(hWnd, uMsg, wParam, lParam);
    }
    case WM_CONTEXTMENU:
    {
        HWND hWndBuddy = GetWindow(hWnd, GW_HWNDNEXT);
        SendMessage(hWndBuddy, uMsg, wParam, lParam);
        return TRUE;
    }
    }
    return DefSubclassProc(hWnd, uMsg, wParam, lParam);
}

inline std::tstring GetName(const std::tstring& line)
{
    LPCTSTR filename = PathFindFileName(line.c_str());
    if (filename)
    {
        LPCTSTR fileext = PathFindExtension(filename);
        if (fileext)
            return std::tstring(filename, fileext - filename);
        else
            return filename;
    }
    else
        return {};
}

struct Options
{
    int icon_mode = ICON_BIG;
    bool sort = false;
    bool blur = true;

    bool ParseCommandLine(const int argc, const LPCTSTR* argv);
};

class RootWindow : public Window
{
    friend WindowManager<RootWindow>;

public:
    static ATOM Register() { return WindowManager<RootWindow>::Register(); }
    static RootWindow* Create(const Options& options) { return WindowManager<RootWindow>::Create(reinterpret_cast<LPVOID>(const_cast<Options*>(&options))); }

protected:
    static void GetCreateWindow(CREATESTRUCT& cs);
    static void GetWndClass(WNDCLASS& wc);
    LRESULT HandleMessage(UINT uMsg, WPARAM wParam, LPARAM lParam) override;

private:
    BOOL OnCreate(LPCREATESTRUCT lpCreateStruct);
    void OnClose();
    void OnDestroy();
    void OnSetFocus(HWND hwndOldFocus);
    void OnHotKey(int idHotKey, UINT fuModifiers, UINT vk);
    void OnSize(UINT state, int cx, int cy);
    void OnEnterSizeMove();
    void OnExitSizeMove();
    void OnActivate(UINT state, HWND hWndActDeact, BOOL fMinimized);
    UINT OnNCHitTest(int x, int y);
    void OnCommand(int id, HWND hWndCtl, UINT codeNotify);
    HBRUSH OnCtlColor(HDC hdc, HWND hWndChild, int type);
    void OnDrawItem(const DRAWITEMSTRUCT* lpDrawItem);
    void OnContextMenu(HWND hWndContext, UINT xPos, UINT yPos);

    static LPCTSTR ClassName() { return TEXT("RadMenu"); }

    HWND m_hEdit = NULL;
    ListBoxOwnerDrawnFixed m_ListBox;
    struct Item
    {
        CComPtr<IShellItem> pShellItem;
        std::tstring name;
        int iIcon = -1;
    };
    std::vector<std::unique_ptr<Item>> m_items;

    JumpListMenuHandler m_jlmh;

    std::vector<std::tstring> GetSearchTerms() const;
    std::tstring GetSelectedText() const;
    static bool Matches(const std::vector<std::tstring>& search, const std::tstring& text);
    void FillList();
    void AddItemToList(const Item& i, const size_t j, const std::vector<std::tstring>& search);
};

void RootWindow::GetCreateWindow(CREATESTRUCT& cs)
{
    const Options& options = *reinterpret_cast<Options*>(cs.lpCreateParams);

    Window::GetCreateWindow(cs);
    cs.lpszName = TEXT("Rad Menu");
    cs.style = WS_POPUP | WS_BORDER | WS_VISIBLE;
    cs.dwExStyle = WS_EX_TOPMOST | WS_EX_TOOLWINDOW;
    cs.dwExStyle |= WS_EX_CONTROLPARENT;
    cs.x = 100;
    cs.y = 100;
    cs.cx = 500;
    cs.cy = 500;
}

void RootWindow::GetWndClass(WNDCLASS& wc)
{
    Window::GetWndClass(wc);
    wc.hbrBackground = g_Theme.brWindow;
    wc.hIcon = LoadIcon(g_hInstance, MAKEINTRESOURCE(IDI_ICON1));
}

void ShowUsage()
{
    MessageBox(NULL,
        TEXT("RadMenu <options>\n")
        TEXT("Where <options> are:\n")
        TEXT("  /is\t\t- use small icons\n")
        TEXT("  /il\t\t- use large icons\n")
        TEXT("  /sort\t\t- sort items\n")
        TEXT("  /noblur\t\t- remove blur effect"),
        TEXT("Rad Menu"), MB_OK | MB_ICONINFORMATION);
}

bool Options::ParseCommandLine(const int argc, const LPCTSTR* argv)
{
    bool ret = true;
    for (int argn = 1; argn < argc; ++argn)
    {
        LPCTSTR arg = argv[argn];
        if (lstrcmpi(arg, TEXT("/is")) == 0)
            icon_mode = ICON_SMALL;
        else if (lstrcmpi(arg, TEXT("/il")) == 0)
            icon_mode = ICON_BIG;
        else if (lstrcmpi(arg, TEXT("/sort")) == 0)
            sort = true;
        else if (lstrcmpi(arg, TEXT("/noblur")) == 0)
            blur = false;
        else if (lstrcmpi(arg, TEXT("/?")) == 0)
        {
            ShowUsage();
            ret = false;
        }
        else
        {
            MessageBox(NULL, Format(TEXT("Unknown argument: %s"), arg).c_str(), TEXT("Rad Menu"), MB_OK | MB_ICONERROR);
            ret = false;
        }
    }
    return ret;
}

BOOL RootWindow::OnCreate(const LPCREATESTRUCT lpCreateStruct)
{
    CHECK_LE(RegisterHotKey(*this, HK_OPEN, MOD_CONTROL | MOD_ALT, 'L'));

    HIMAGELIST hImageListLg, hImageListSm;
    CHECK(Shell_GetImageLists(&hImageListLg, &hImageListSm));

    const Options& options = *reinterpret_cast<Options*>(lpCreateStruct->lpCreateParams);

    if (options.blur)
        SetWindowBlur(*this, true);

    const RECT rcClient = CallWinApi<RECT>(GetClientRect, HWND(*this));

    LOGFONT lf;
    SystemParametersInfo(SPI_GETICONTITLELOGFONT, 0, &lf, 0);
    const HFONT hFont = CreateFontIndirect(&lf);
    const SIZE TextSize = GetFontSize(*this, hFont, TEXT("Mg"), 2);

    RECT rc = { Border, Border, Width(rcClient) - Border, Border + TextSize.cy + Border };

    m_hEdit = Edit_Create(*this, WS_CHILD | WS_VISIBLE | WS_BORDER | WS_TABSTOP | ES_LEFT, rc, IDC_EDIT);
    if (m_hEdit)
    {
        SetWindowFont(m_hEdit, hFont, FALSE);
        CHECK_LE(SetWindowSubclass(m_hEdit, BuddyProc, 0, 0));
        Edit_SetCueBannerTextFocused(m_hEdit, TEXT("Search"), TRUE);
    }

    rc.top = rc.bottom + Border;
    rc.bottom = rcClient.bottom - Border;
    if (options.icon_mode != -1)
    {
        HIMAGELIST hImageList = options.icon_mode == ICON_SMALL ? hImageListSm : hImageListLg;
        HIMAGELIST hOldImageList = m_ListBox.SetImageList(hImageList);
        CHECK_LE(!hOldImageList || ImageList_Destroy(hOldImageList));
    }
    m_ListBox.Create(*this, WS_CHILD | WS_VISIBLE | WS_BORDER | WS_VSCROLL | WS_TABSTOP | LBS_USETABSTOPS | LBS_NOTIFY | (options.sort ? LBS_SORT : 0), rc, IDC_LIST);
    SetWindowFont(m_ListBox, hFont, FALSE);

    CComPtr<IShellItem> folder;
    if (SUCCEEDED(SHGetKnownFolderItem(FOLDERID_AppsFolder, KF_FLAG_DEFAULT, nullptr, IID_PPV_ARGS(&folder))))
    {
        CComPtr<IEnumShellItems> enumItems;
        if (SUCCEEDED(folder->BindToHandler(nullptr, BHID_EnumItems, IID_PPV_ARGS(&enumItems))))
        {
            IShellItem* items;
            while (enumItems->Next(1, &items, nullptr) == S_OK)
            {
                CComPtr<IShellItem> item = items;
                CComHeapPtr<wchar_t> name_normal;
                CHECK_HR(item->GetDisplayName(SIGDN_NORMALDISPLAY, &name_normal));
                m_items.push_back(std::unique_ptr<Item>(new Item({ item, get(name_normal) })));
            }
        }
    }

    FillList();

    return TRUE;
}

void RootWindow::OnClose()
{
    if (GetKeyState(VK_SHIFT) & KF_UP)
        SetHandled(false);
    else
        ShowWindow(*this, SW_HIDE);
}

void RootWindow::OnDestroy()
{
    //ImageList_Destroy(m_ListBox.SetImageList(NULL));
    PostQuitMessage(0);
}

void RootWindow::OnSetFocus(const HWND hwndOldFocus)
{
    if (m_hEdit)
        SetFocus(m_hEdit);
}

void RootWindow::OnHotKey(int idHotKey, UINT fuModifiers, UINT vk)
{
    switch (idHotKey)
    {
    case HK_OPEN:
        SetForegroundWindow(*this);
        ShowWindow(*this, SW_SHOWNORMAL);
        break;
    }
}

void RootWindow::OnSize(UINT state, int cx, int cy)
{
    RECT rc = CallWinApi<RECT>(GetWindowRect, m_hEdit);
    CHECK_LE(ScreenToClient(*this, &rc));
    rc.right = cx - Border;
    CHECK_LE(SetWindowPos(m_hEdit, NULL, rc.left, rc.top, Width(rc), Height(rc), SWP_NOOWNERZORDER | SWP_NOZORDER));
    CHECK_LE(InvalidateRect(m_hEdit, nullptr, FALSE));

    rc.top = rc.bottom + Border;
    rc.bottom = cy - Border;
    CHECK_LE(SetWindowPos(m_ListBox, NULL, rc.left, rc.top, Width(rc), Height(rc), SWP_NOOWNERZORDER | SWP_NOZORDER));
    CHECK_LE(InvalidateRect(m_ListBox, nullptr, FALSE));
}

void RootWindow::OnEnterSizeMove()
{
    SetWindowBlur(*this, false);
}

void RootWindow::OnExitSizeMove()
{
    // TODO Do not reenable if disabled in options
    SetWindowBlur(*this, true);
}

void RootWindow::OnActivate(UINT state, HWND hWndActDeact, BOOL fMinimized)
{
#ifndef _DEBUG
    if (state == WA_INACTIVE)
        SendMessage(*this, WM_CLOSE, 0, 0);
#endif
}

UINT RootWindow::OnNCHitTest(int x, int y)
{
    POINT pt = { x, y };
    CHECK_LE(ScreenToClient(*this, &pt));

    RECT rc = CallWinApi<RECT>(GetClientRect, HWND(*this));
    CHECK_LE(InflateRect(&rc, -5, -5));

    if (pt.x > rc.right and pt.x <= rc.top)
        return HTTOPRIGHT;
    else if (pt.x <= rc.left and pt.x <= rc.top)
        return HTTOPLEFT;
    else if (pt.x > rc.right and pt.y > rc.bottom)
        return HTBOTTOMRIGHT;
    else if (pt.x <= rc.left and pt.y > rc.bottom)
        return HTBOTTOMLEFT;
    else if (pt.x > rc.right)
        return HTRIGHT;
    else if (pt.x <= rc.left)
        return HTLEFT;
    else if (pt.y > rc.bottom)
        return HTBOTTOM;
    else if (pt.x <= rc.top)
        return HTTOP;
    else
        return HTCAPTION;
}

void RootWindow::OnCommand(int id, HWND hWndCtl, UINT codeNotify)
{
    switch (id)
    {
    case IDC_EDIT:
        switch (codeNotify)
        {
        case EN_CHANGE:
            FillList();
            break;
        }
        break;

    case IDC_LIST:
        switch (codeNotify)
        {
        case LBN_DBLCLK:
            SendMessage(*this, WM_COMMAND, IDOK, 0);
            break;
        }
        break;

    case IDCANCEL:
        if (Edit_GetTextLength(m_hEdit) > 0)
            Edit_SetText(m_hEdit, TEXT(""));
        else
            SendMessage(*this, WM_CLOSE, 0, 0);
        break;

    case IDOK:
    {
        const int sel = m_ListBox.GetCurSel();
        if (sel >= 0)
        {
            const int j = (int)m_ListBox.GetItemData(sel);

            CComPtr<IContextMenu> pContextMenu;
            CHECK_HR(m_items[j]->pShellItem->BindToHandler(nullptr, BHID_SFUIObject, IID_PPV_ARGS(&pContextMenu)));

            if (pContextMenu)
            {
                OpenDefaultItem(*this, pContextMenu);
                SendMessage(*this, WM_CLOSE, 0, 0);
            }
        }
        break;
    }
    }
}

HBRUSH RootWindow::OnCtlColor(HDC hDC, HWND hWndChild, int type)
{
    SetTextColor(hDC, g_Theme.clrWindowText);
    SetBkColor(hDC, g_Theme.clrWindow);
    return g_Theme.brWindow;
}

void RootWindow::OnDrawItem(const DRAWITEMSTRUCT* lpDrawItem)
{
    if (lpDrawItem->hwndItem == m_ListBox and lpDrawItem->itemID >= 0)
    {
        const int j = (int) m_ListBox.GetItemData(lpDrawItem->itemID);
        Item& item = *m_items[j].get();
        if (m_ListBox.GetItemIconIndex(lpDrawItem->itemID) == -1)
        {
            if (item.iIcon == -1)
            {
                item.iIcon = GetIconIndex(item.pShellItem);
                if (item.iIcon == -1)
                    item.iIcon = -2;
            }
            m_ListBox.SetItemIconIndex(lpDrawItem->itemID, item.iIcon);
        }
    }
    SetHandled(false);
}

void RootWindow::OnContextMenu(HWND hWndContext, UINT xPos, UINT yPos)
{
    int sel = -1;
    if (xPos == UINT_MAX && yPos == UINT_MAX)
    {
        sel = m_ListBox.GetCurSel();
        if (sel >= 0)
        {
            RECT rc = {};
            if (m_ListBox.GetItemRect(sel, &rc) != LB_ERR)
            {
                POINT pt = { rc.right, rc.top };
                CHECK_LE(ClientToScreen(m_ListBox, &pt));
                xPos = pt.x;
                yPos = pt.y;
            }
        }
    }
    else
    {
        POINT pt = { (LONG) xPos, (LONG) yPos };
        CHECK_LE(ScreenToClient(m_ListBox, &pt));
        sel = m_ListBox.ItemFromPoint(pt);
    }
    if (sel >= 0)
    {
        const int j = (int) m_ListBox.GetItemData(sel);

        if (m_jlmh.DoJumpListMenu(*this, m_items[j]->pShellItem, { (LONG) xPos, (LONG) yPos }))
            SendMessage(*this, WM_CLOSE, 0, 0);
    }
}

std::vector<std::tstring> RootWindow::GetSearchTerms() const
{
    TCHAR text[1024];
    Edit_GetText(m_hEdit, text, ARRAYSIZE(text));
    return split_unquote(text, TEXT(' '));
}

std::tstring RootWindow::GetSelectedText() const
{
    const int sel = m_ListBox.GetCurSel();
    TCHAR buf[1024] = TEXT("");
    if (sel >= 0)
        m_ListBox.GetText(sel, buf);
    return buf;
}

bool RootWindow::Matches(const std::vector<std::tstring>& search, const std::tstring& text)
{
    bool found = true;
    for (const auto& s : search)
        if (!StrFindI(text.c_str(), s.c_str()))
        {
            found = false;
            break;
        }
    return found;
}

void RootWindow::FillList()
{
    SetWindowRedraw(m_ListBox, FALSE);
    const std::tstring seltext = GetSelectedText();
    const std::vector<std::tstring> search = GetSearchTerms();
    m_ListBox.ResetContent();
    size_t j = 0;
    if (!search.empty())
        for (const auto& sp : m_items)
            AddItemToList(*sp.get(), j++, search);
    const int sel = m_ListBox.FindStringExact(-1, seltext.c_str());
    m_ListBox.SetCurSel(sel >= 0 ? sel : 0);
    SendMessage(*this, WM_COMMAND, MAKEWPARAM(IDC_LIST, LBN_SELCHANGE), m_ListBox);
    SetWindowRedraw(m_ListBox, TRUE);
}

void RootWindow::AddItemToList(const Item& sp, const size_t j, const std::vector<std::tstring>& search)
{
    const std::tstring& text = sp.name;
    if (Matches(search, text))
    {
        SendMessage(m_ListBox, WM_SETREDRAW, FALSE, 0);
        const int n = m_ListBox.AddString(text.c_str());
        m_ListBox.SetItemData(n, j);
        m_ListBox.SetItemIconIndex(n, sp.iIcon);
        SendMessage(m_ListBox, WM_SETREDRAW, TRUE, 0);
    }
}

LRESULT RootWindow::HandleMessage(const UINT uMsg, const WPARAM wParam, const LPARAM lParam)
{
    LRESULT ret = 0;
    switch (uMsg)
    {
        HANDLE_MSG(WM_CREATE, OnCreate);
        HANDLE_MSG(WM_CLOSE, OnClose);
        HANDLE_MSG(WM_DESTROY, OnDestroy);
        HANDLE_MSG(WM_SETFOCUS, OnSetFocus);
        HANDLE_MSG(WM_HOTKEY, OnHotKey);
        HANDLE_MSG(WM_SIZE, OnSize);
        HANDLE_MSG(WM_ENTERSIZEMOVE, OnEnterSizeMove);
        HANDLE_MSG(WM_EXITSIZEMOVE, OnExitSizeMove);
        HANDLE_MSG(WM_ACTIVATE, OnActivate);
        HANDLE_MSG(WM_NCHITTEST, OnNCHitTest);
        HANDLE_MSG(WM_COMMAND, OnCommand);
        HANDLE_MSG(WM_CTLCOLOREDIT, OnCtlColor);
        HANDLE_MSG(WM_CTLCOLORSTATIC, OnCtlColor);
        HANDLE_MSG(WM_CTLCOLORLISTBOX, OnCtlColor);
        HANDLE_MSG(WM_DRAWITEM, OnDrawItem);
        HANDLE_MSG(WM_CONTEXTMENU, OnContextMenu);
    }

    if (!IsHandled())
    {
        bool bHandled = false;
        ret = m_ListBox.ProcessMessage(*this, uMsg, wParam, lParam, bHandled);
        if (bHandled)
            SetHandled(true);
    }
    if (!IsHandled())
        ret = Window::HandleMessage(uMsg, wParam, lParam);

    if (true)
    {
        bool bHandled = false;
        const LRESULT subret = m_jlmh.ProcessMessage(*this, uMsg, wParam, lParam, bHandled);
        if (bHandled)
            ret = subret;
    }

    return ret;
}

bool Run(_In_ const LPCTSTR lpCmdLine, _In_ const int nShowCmd)
{
    RadLogInitWnd(NULL, "RadMenu", L"RadMenu");

    if (true)
    {
        g_Theme.clrWindow = RGB(0, 0, 0);
        g_Theme.clrHighlight = RGB(61, 61, 61);
        g_Theme.clrWindowText = RGB(250, 250, 250);
        g_Theme.clrHighlightText = g_Theme.clrWindowText;
        g_Theme.clrGrayText = RGB(128, 128, 128);
    }

    Options options;
    if (!options.ParseCommandLine(__argc, __wargv))
        return false;

    InitTheme();

    CHECK_LE_RET(RootWindow::Register(), false);

    RootWindow* prw = RootWindow::Create(options);
    CHECK_LE_RET(prw != nullptr, false);

    RadLogInitWnd(*prw, "RadMenu", L"RadMenu");
    g_hWndDlg = *prw;
    CHECK_LE(ShowWindow(*prw, nShowCmd));

    return true;
}

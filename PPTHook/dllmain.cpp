#include "pch.h"
#include <Windows.h>
#include <windowsx.h>
#include <stdio.h>
#include <vector>
#include <string>
#include <fstream>
#include <objbase.h>
#include <olectl.h>
#include <stdlib.h>

#pragma comment(lib, "ole32.lib")
#pragma comment(lib, "oleaut32.lib")

// -- constants -----------------------------------------------------------------
static const int MAX_THUMB_W = 160;
static const int MAX_THUMB_H = 120;
static const int SRC_THUMB_W = 160;
static const int SRC_THUMB_H = 120;
static const int PADDING     = 8;
static const int HEADER_H    = 24;
static const int HEADER_W    = 28;
static const int ROWS_PER_COL = 100;
static const COLORREF BG          = RGB(214, 217, 223);
static const COLORREF HEADER_BG   = RGB(195, 199, 206);
static const COLORREF HEADER_TEXT = RGB(68, 68, 68);
static const COLORREF SEL_BG      = RGB(212, 229, 252);
static const COLORREF SEL_BORDER  = RGB(84, 131, 197);
static const COLORREF DIVIDER     = RGB(160, 164, 171);
static const COLORREF THUMB_BG    = RGB(255, 255, 255);
static const int SCROLLBAR_W      = 6;
static const COLORREF SCROLL_TRACK = RGB(200, 203, 209);
static const COLORREF SCROLL_THUMB = RGB(140, 144, 152);
static const int PANE_MIN_W       = 10;

// -- globals -------------------------------------------------------------------
// Hooked windows + their original procs
WNDPROC g_origPanel     = nullptr;
WNDPROC g_origGutter    = nullptr;
WNDPROC g_origSplitter  = nullptr;
WNDPROC g_origMainSlide = nullptr;
WNDPROC g_origHSplitter = nullptr;
WNDPROC g_origNotes     = nullptr;
WNDPROC g_origTabBar    = nullptr;

HWND g_slidePanel     = nullptr;  // paneClassDC — our panel (leftmost, tallest)
HWND g_gutterHwnd     = nullptr;  // paneClass  ~20px between panel and splitter
HWND g_splitterBar    = nullptr;  // splitterBar ~6px vertical drag handle
HWND g_mainSlide      = nullptr;  // paneClassDC — largest area right of splitter
HWND g_hSplitter      = nullptr;  // splitterBar horizontal (slide/notes divider)
HWND g_notesPane      = nullptr;  // paneClassDC — smaller area right of splitter
HWND g_tabBar         = nullptr;  // paneClass   — short strip near panel left
HWND g_foundHwnd      = nullptr;
HWND g_mainPptWindow  = nullptr;

int g_gutterW    = 20; // populated at init
int g_splitterW  = 6;  // populated at init

int g_slideCount        = 0;
int g_currentSlide      = 1;
int g_bestPaneScore     = -2147483647;
int g_bestMainArea      = 0;
int g_firstVisibleIndex = 0;  // first visible ROW index (0-based)
int g_firstVisibleCol   = 0;  // first visible COLUMN index (0-based)

// Splitter drag state
static bool g_dragging        = false;
static int  g_dragStartX      = 0;
static int  g_dragStartPaneW  = 0;
static int  g_desiredPaneW    = 0; // enforced by WM_WINDOWPOSCHANGING
static int  g_dragPaneW       = 0; // set by DoSplitterResize, used during drag
static bool g_enforcing       = false; // re-entrancy guard for EnforceLayout

static float g_zoomLevel = 1.0f;  // 0.25 to 3.0
static bool  g_panelHasFocus = false;
static bool  g_panelIsActive = false; // true until user clicks outside our panel
static HHOOK g_kbHook = nullptr;
static HHOOK g_mouseHook = nullptr;

struct ThumbCache { int slideIndex; HBITMAP hbmp; bool valid; };
std::vector<ThumbCache> g_thumbs;
static volatile bool g_thumbLoadCancel = false;
static HANDLE g_thumbThread = nullptr;

// -- zoom-aware geometry helpers -----------------------------------------------
int ZThumbW()   { return (int)(MAX_THUMB_W * g_zoomLevel); }
int ZThumbH()   { return (int)(MAX_THUMB_H * g_zoomLevel); }
int ZCellW()    { return ZThumbW() + PADDING; }
int ZCellH()    { return ZThumbH() + PADDING; }
// Headers/font: gentle inverse scaling — bigger labels when zoomed out
int ZHeaderH()  { return 20; }
int ZHeaderW()  { return 28; }
int ZFontSize() { return max(14, (int)(20 * g_zoomLevel)); }

// -- logging -------------------------------------------------------------------
void LogDebug(const wchar_t* msg) {
    wchar_t path[MAX_PATH] = {};
    wchar_t* up = nullptr; size_t ul = 0;
    _wdupenv_s(&up, &ul, L"USERPROFILE");
    if (up && ul > 0) swprintf_s(path, L"%s\\source\\repos\\PPTHook\\debug.txt", up);
    else wcscpy_s(path, L"C:\\Users\\timon\\source\\repos\\PPTHook\\debug.txt");
    if (up) free(up);
    SYSTEMTIME st = {}; GetLocalTime(&st);
    std::wofstream f(path, std::ios::app);
    if (f.is_open())
        f << L"[" << st.wHour << L":" << st.wMinute << L":" << st.wSecond << L"] " << msg << L"\n";
}

// -- layout geometry helpers ---------------------------------------------------
// Computes desired positions for all managed windows in parent-client coords.
// Returns false if we can't compute (missing panel, etc).
struct LayoutPositions {
    int panelL, panelT, panelB, panelW;
    int gutterL;
    int splitterL;
    int contentL;
    int contentR;
};

static bool ComputeLayout(LayoutPositions& lp) {
    if (!g_slidePanel || g_desiredPaneW <= 0) return false;
    HWND parent = GetParent(g_slidePanel);
    if (!parent) return false;

    RECT parentRc, panelRc;
    GetWindowRect(parent, &parentRc);
    GetWindowRect(g_slidePanel, &panelRc);

    int parentW = parentRc.right - parentRc.left;

    RECT panelIP = panelRc;
    MapWindowPoints(HWND_DESKTOP, parent, reinterpret_cast<LPPOINT>(&panelIP), 2);

    lp.panelL    = panelIP.left;
    lp.panelT    = panelIP.top;
    lp.panelB    = panelIP.bottom;
    lp.panelW    = g_desiredPaneW;
    lp.gutterL   = lp.panelL + g_desiredPaneW;
    lp.splitterL = lp.gutterL + g_gutterW;
    lp.contentL  = lp.splitterL + g_splitterW;
    lp.contentR  = parentW - 20; // right scrollbar column
    return true;
}

// -- COM helpers ---------------------------------------------------------------
IDispatch* GetPPTApp() {
    CLSID clsid;
    if (FAILED(CLSIDFromProgID(L"PowerPoint.Application", &clsid))) return nullptr;
    IUnknown* pUnk = nullptr;
    if (FAILED(GetActiveObject(clsid, nullptr, &pUnk)) || !pUnk) return nullptr;
    IDispatch* p = nullptr;
    pUnk->QueryInterface(IID_IDispatch, (void**)&p);
    pUnk->Release();
    return p;
}

IDispatch* Invoke(IDispatch* d, LPCWSTR name, DISPPARAMS* dp = nullptr) {
    if (!d) return nullptr;
    DISPID id; BSTR b = SysAllocString(name);
    HRESULT hr = d->GetIDsOfNames(IID_NULL, &b, 1, LOCALE_USER_DEFAULT, &id);
    SysFreeString(b); if (FAILED(hr)) return nullptr;
    VARIANT r; VariantInit(&r); DISPPARAMS e = {};
    hr = d->Invoke(id, IID_NULL, LOCALE_USER_DEFAULT,
        DISPATCH_METHOD | DISPATCH_PROPERTYGET, dp ? dp : &e, &r, nullptr, nullptr);
    if (FAILED(hr)) return nullptr;
    if (r.vt == VT_DISPATCH) return r.pdispVal;
    if (r.vt == VT_UNKNOWN && r.punkVal) {
        IDispatch* out = nullptr;
        r.punkVal->QueryInterface(IID_IDispatch, (void**)&out);
        VariantClear(&r); return out;
    }
    VariantClear(&r); return nullptr;
}

int InvokeInt(IDispatch* d, LPCWSTR name) {
    if (!d) return 0;
    DISPID id; BSTR b = SysAllocString(name);
    HRESULT hr = d->GetIDsOfNames(IID_NULL, &b, 1, LOCALE_USER_DEFAULT, &id);
    SysFreeString(b); if (FAILED(hr)) return 0;
    VARIANT r; VariantInit(&r); DISPPARAMS e = {};
    d->Invoke(id, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_PROPERTYGET, &e, &r, nullptr, nullptr);
    if (r.vt == VT_I4) return r.lVal;
    if (r.vt == VT_R8) return (int)r.dblVal;
    return 0;
}

std::wstring InvokeStr(IDispatch* d, LPCWSTR name) {
    if (!d) return L"";
    DISPID id; BSTR b = SysAllocString(name);
    HRESULT hr = d->GetIDsOfNames(IID_NULL, &b, 1, LOCALE_USER_DEFAULT, &id);
    SysFreeString(b); if (FAILED(hr)) return L"";
    VARIANT r; VariantInit(&r); DISPPARAMS e = {};
    d->Invoke(id, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_PROPERTYGET, &e, &r, nullptr, nullptr);
    std::wstring result;
    if (r.vt == VT_BSTR && r.bstrVal) result = r.bstrVal;
    VariantClear(&r);
    return result;
}

// Forward declarations for formula evaluation
int GetPPTCurrentSlide();

// -- Formula evaluation --------------------------------------------------------

std::wstring ReadSlideText(int slideIndex) {
    IDispatch* pApp = GetPPTApp(); if (!pApp) return L"";
    IDispatch* pres = Invoke(pApp, L"ActivePresentation");
    IDispatch* slides = pres ? Invoke(pres, L"Slides") : nullptr;
    std::wstring text;
    if (slides) {
        VARIANT vi; vi.vt = VT_I4; vi.lVal = slideIndex;
        DISPPARAMS dp = { &vi, nullptr, 1, 0 };
        IDispatch* slide = Invoke(slides, L"Item", &dp);
        IDispatch* shapes = slide ? Invoke(slide, L"Shapes") : nullptr;
        int count = shapes ? InvokeInt(shapes, L"Count") : 0;
        for (int i = 1; i <= count && text.empty(); i++) {
            VARIANT si; si.vt = VT_I4; si.lVal = i;
            DISPPARAMS sp = { &si, nullptr, 1, 0 };
            IDispatch* shape = Invoke(shapes, L"Item", &sp);
            if (shape) {
                int htf = InvokeInt(shape, L"HasTextFrame");
                if (htf) {
                    IDispatch* tf = Invoke(shape, L"TextFrame");
                    IDispatch* tr = tf ? Invoke(tf, L"TextRange") : nullptr;
                    if (tr) { text = InvokeStr(tr, L"Text"); tr->Release(); }
                    if (tf) tf->Release();
                }
                shape->Release();
            }
        }
        if (shapes) shapes->Release();
        if (slide) slide->Release();
        slides->Release();
    }
    if (pres) pres->Release();
    pApp->Release();
    return text;
}

void SetSlideText(int slideIndex, const std::wstring& newText) {
    IDispatch* pApp = GetPPTApp(); if (!pApp) return;
    IDispatch* pres = Invoke(pApp, L"ActivePresentation");
    IDispatch* slides = pres ? Invoke(pres, L"Slides") : nullptr;
    if (slides) {
        VARIANT vi; vi.vt = VT_I4; vi.lVal = slideIndex;
        DISPPARAMS dp = { &vi, nullptr, 1, 0 };
        IDispatch* slide = Invoke(slides, L"Item", &dp);
        IDispatch* shapes = slide ? Invoke(slide, L"Shapes") : nullptr;
        int count = shapes ? InvokeInt(shapes, L"Count") : 0;
        for (int i = 1; i <= count; i++) {
            VARIANT si; si.vt = VT_I4; si.lVal = i;
            DISPPARAMS sp = { &si, nullptr, 1, 0 };
            IDispatch* shape = Invoke(shapes, L"Item", &sp);
            if (shape) {
                int htf = InvokeInt(shape, L"HasTextFrame");
                if (htf) {
                    IDispatch* tf = Invoke(shape, L"TextFrame");
                    IDispatch* tr = tf ? Invoke(tf, L"TextRange") : nullptr;
                    if (tr) {
                        // PROPERTYPUT on TextRange.Text
                        DISPID propId; BSTR bName = SysAllocString(L"Text");
                        HRESULT hr = tr->GetIDsOfNames(IID_NULL, &bName, 1, LOCALE_USER_DEFAULT, &propId);
                        SysFreeString(bName);
                        if (SUCCEEDED(hr)) {
                            VARIANT val; val.vt = VT_BSTR; val.bstrVal = SysAllocString(newText.c_str());
                            DISPID putId = DISPID_PROPERTYPUT;
                            DISPPARAMS putDp = { &val, &putId, 1, 1 };
                            tr->Invoke(propId, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_PROPERTYPUT, &putDp, nullptr, nullptr, nullptr);
                            SysFreeString(val.bstrVal);
                        }
                        tr->Release();
                    }
                    if (tf) tf->Release();
                    if (shapes) shapes->Release();
                    if (slide) slide->Release();
                    slides->Release();
                    if (pres) pres->Release();
                    pApp->Release();
                    return;
                }
                shape->Release();
            }
        }
        if (shapes) shapes->Release();
        if (slide) slide->Release();
        slides->Release();
    }
    if (pres) pres->Release();
    pApp->Release();
}

int ParseCellRef(const std::wstring& ref) {
    // Parse column letters and row digits from e.g. "A4", "C3", "AA1"
    int col = 0; size_t i = 0;
    while (i < ref.size() && iswalpha(ref[i])) {
        col = col * 26 + (towupper(ref[i]) - L'A');
        i++;
    }
    if (i == 0 || i == ref.size()) return -1;
    int row = 0;
    while (i < ref.size() && iswdigit(ref[i])) {
        row = row * 10 + (ref[i] - L'0');
        i++;
    }
    if (i != ref.size() || row < 1) return -1;
    return col * ROWS_PER_COL + row; // 1-based slide index
}

bool ParseNumber(const std::wstring& s, double* out) {
    if (s.empty()) return false;
    wchar_t* end = nullptr;
    *out = wcstod(s.c_str(), &end);
    return end == s.c_str() + s.size();
}

bool ResolveOperand(const std::wstring& s, double* out) {
    if (ParseNumber(s, out)) return true;
    int slideIdx = ParseCellRef(s);
    if (slideIdx < 0) return false;
    std::wstring cellText = ReadSlideText(slideIdx);
    // If cell contains a formula result like "5+3=8", extract the part after the last '='
    size_t eqPos = cellText.rfind(L'=');
    if (eqPos != std::wstring::npos && eqPos + 1 < cellText.size())
        cellText = cellText.substr(eqPos + 1);
    return ParseNumber(cellText, out);
}

std::wstring EvaluateExpression(const std::wstring& expr) {
    // Find operator scanning from index 1 (to skip leading minus)
    int opPos = -1;
    wchar_t op = 0;
    for (size_t i = 1; i < expr.size(); i++) {
        wchar_t c = expr[i];
        if (c == L'+' || c == L'-' || c == L'*' || c == L'/') {
            opPos = (int)i;
            op = c;
            break;
        }
    }
    if (opPos < 0) return L"#ERR";
    std::wstring left = expr.substr(0, opPos);
    std::wstring right = expr.substr(opPos + 1);
    double a, b;
    if (!ResolveOperand(left, &a) || !ResolveOperand(right, &b)) return L"#ERR";
    double result;
    switch (op) {
    case L'+': result = a + b; break;
    case L'-': result = a - b; break;
    case L'*': result = a * b; break;
    case L'/':
        if (b == 0.0) return L"#DIV/0!";
        result = a / b; break;
    default: return L"#ERR";
    }
    wchar_t buf[64];
    swprintf_s(buf, L"%.10g", result);
    return buf;
}

void HandleFormulaEvaluation() {
    int slideIndex = GetPPTCurrentSlide();
    if (slideIndex <= 0) return;
    std::wstring text = ReadSlideText(slideIndex);
    if (text.empty() || text.back() != L'=') return;
    std::wstring expr = text.substr(0, text.size() - 1);
    if (expr.empty()) return;
    std::wstring result = EvaluateExpression(expr);
    SetSlideText(slideIndex, expr + L"=" + result);
    if (g_slidePanel) SetTimer(g_slidePanel, 2, 100, NULL);
}

// Single reusable temp file path
static wchar_t g_thumbTempPath[MAX_PATH] = {};

void InitThumbTempPath() {
    if (g_thumbTempPath[0] == 0) {
        wchar_t tmp[MAX_PATH] = {};
        GetTempPathW(MAX_PATH, tmp);
        swprintf_s(g_thumbTempPath, L"%sppt_thumb.bmp", tmp);
    }
}

// Export a single slide thumbnail using a pre-fetched Slides collection
HBITMAP GetSlideThumbnailFromSlides(IDispatch* slides, int idx) {
    VARIANT vi; vi.vt = VT_I4; vi.lVal = idx;
    DISPPARAMS dp = { &vi, nullptr, 1, 0 };
    IDispatch* slide = Invoke(slides, L"Item", &dp);
    if (!slide) return nullptr;

    InitThumbTempPath();
    // Delete previous temp file to avoid stale reads
    DeleteFileW(g_thumbTempPath);

    VARIANT args[4];
    args[3].vt = VT_BSTR; args[3].bstrVal = SysAllocString(g_thumbTempPath);
    args[2].vt = VT_BSTR; args[2].bstrVal = SysAllocString(L"BMP");
    args[1].vt = VT_I4;   args[1].lVal = MAX_THUMB_W;
    args[0].vt = VT_I4;   args[0].lVal = MAX_THUMB_H;
    DISPPARAMS edp = { args, nullptr, 4, 0 };
    Invoke(slide, L"Export", &edp);
    for (int i = 0; i < 4; i++) VariantClear(&args[i]);
    slide->Release();

    if (GetFileAttributesW(g_thumbTempPath) == INVALID_FILE_ATTRIBUTES) return nullptr;
    return (HBITMAP)LoadImageW(nullptr, g_thumbTempPath, IMAGE_BITMAP, 0, 0, LR_LOADFROMFILE);
}

// Legacy wrapper for single-slide refresh (timer handler)
HBITMAP GetSlideThumbnail(IDispatch* pApp, int idx) {
    IDispatch* pres  = Invoke(pApp, L"ActivePresentation");
    IDispatch* slides = Invoke(pres, L"Slides");
    if (!slides) { if (pres) pres->Release(); return nullptr; }
    HBITMAP hbmp = GetSlideThumbnailFromSlides(slides, idx);
    slides->Release(); if (pres) pres->Release();
    return hbmp;
}

// Forward declarations for grid helpers (used by ThumbLoaderThread)
int GetNumCols();
int GetFullVisRows(HWND h);

// Lightweight init: just get count + create placeholder entries
void RefreshThumbs(IDispatch* pApp) {
    // Cancel any running background load
    g_thumbLoadCancel = true;
    if (g_thumbThread) {
        WaitForSingleObject(g_thumbThread, 5000);
        CloseHandle(g_thumbThread);
        g_thumbThread = nullptr;
    }
    g_thumbLoadCancel = false;

    for (auto& t : g_thumbs) if (t.hbmp) DeleteObject(t.hbmp);
    g_thumbs.clear();
    IDispatch* pres  = Invoke(pApp,  L"ActivePresentation");
    IDispatch* slides = Invoke(pres, L"Slides");
    if (!slides) { if (pres) pres->Release(); return; }
    g_slideCount = InvokeInt(slides, L"Count");
    g_thumbs.resize(g_slideCount);
    for (int i = 0; i < g_slideCount; i++) {
        g_thumbs[i].slideIndex = i + 1;
        g_thumbs[i].hbmp = nullptr;
        g_thumbs[i].valid = false;
    }
    slides->Release(); if (pres) pres->Release();
}

// Background thread: loads thumbnails, visible ones first, then the rest
static DWORD WINAPI ThumbLoaderThread(LPVOID) {
    if (FAILED(CoInitialize(nullptr))) return 1;

    IDispatch* pApp = GetPPTApp();
    if (!pApp) { CoUninitialize(); return 1; }
    IDispatch* pres = Invoke(pApp, L"ActivePresentation");
    IDispatch* slides = pres ? Invoke(pres, L"Slides") : nullptr;
    if (!slides) {
        if (pres) pres->Release();
        pApp->Release(); CoUninitialize(); return 1;
    }

    int total = g_slideCount;
    std::vector<bool> loaded(total, false);
    int doneCount = 0;

    while (doneCount < total && !g_thumbLoadCancel) {
        // Build priority list: visible slides first, then the rest
        std::vector<int> order;
        order.reserve(total);

        int firstVis = g_firstVisibleIndex;
        int visRows = g_slidePanel ? GetFullVisRows(g_slidePanel) : 20;
        int visCols = GetNumCols();
        for (int col = 0; col < visCols && !g_thumbLoadCancel; col++) {
            for (int row = firstVis; row < firstVis + visRows && row < ROWS_PER_COL; row++) {
                int idx = col * ROWS_PER_COL + row;
                if (idx >= 0 && idx < total && !loaded[idx])
                    order.push_back(idx);
            }
        }
        for (int i = 0; i < total; i++) {
            if (!loaded[i]) {
                bool inOrder = false;
                for (int o : order) { if (o == i) { inOrder = true; break; } }
                if (!inOrder) order.push_back(i);
            }
        }

        int batchLoaded = 0;
        for (int idx : order) {
            if (g_thumbLoadCancel) break;
            if (loaded[idx]) continue;

            HBITMAP hbmp = GetSlideThumbnailFromSlides(slides, idx + 1);
            if (g_thumbLoadCancel) {
                if (hbmp) DeleteObject(hbmp);
                break;
            }

            if (idx < (int)g_thumbs.size()) {
                if (g_thumbs[idx].hbmp) DeleteObject(g_thumbs[idx].hbmp);
                g_thumbs[idx].hbmp = hbmp;
                g_thumbs[idx].valid = (hbmp != nullptr);
            } else {
                if (hbmp) DeleteObject(hbmp);
            }

            loaded[idx] = true;
            doneCount++;
            batchLoaded++;

            if (batchLoaded % 20 == 0 && g_slidePanel)
                InvalidateRect(g_slidePanel, NULL, FALSE);
        }

        if (batchLoaded > 0 && g_slidePanel)
            InvalidateRect(g_slidePanel, NULL, FALSE);
    }

    // Cleanup temp file
    if (g_thumbTempPath[0]) DeleteFileW(g_thumbTempPath);

    slides->Release();
    if (pres) pres->Release();
    pApp->Release();
    CoUninitialize();
    return 0;
}

// -- grid helpers --------------------------------------------------------------
void ColLabel(int c, char* buf, int n) {
    if (c < 26) sprintf_s(buf, n, "%c", 'A' + c);
    else sprintf_s(buf, n, "%c%c", 'A' + c / 26 - 1, 'A' + c % 26);
}

int GetNumCols() { return (g_slideCount + ROWS_PER_COL - 1) / ROWS_PER_COL; }
int GetFullVisRows(HWND h) { RECT r; GetClientRect(h, &r); return max(1, (r.bottom - ZHeaderH()) / ZCellH()); }
int GetDrawRows(HWND h) { RECT r; GetClientRect(h, &r); return max(1, ((r.bottom - ZHeaderH()) + ZCellH() - 1) / ZCellH()); }
int GetFullVisCols(HWND h) { RECT r; GetClientRect(h, &r); return max(1, (r.right - ZHeaderW()) / ZCellW()); }
int GetDrawCols(HWND h) { RECT r; GetClientRect(h, &r); return max(1, ((r.right - ZHeaderW()) + ZCellW() - 1) / ZCellW()); }

void ClampScroll(HWND h) {
    int visRows = GetFullVisRows(h);
    int maxFirst = max(0, ROWS_PER_COL - visRows);
    if (g_firstVisibleIndex > maxFirst) g_firstVisibleIndex = maxFirst;
    if (g_firstVisibleIndex < 0) g_firstVisibleIndex = 0;
    int visCols = GetFullVisCols(h);
    int maxFirstCol = max(0, GetNumCols() - visCols);
    if (g_firstVisibleCol > maxFirstCol) g_firstVisibleCol = maxFirstCol;
    if (g_firstVisibleCol < 0) g_firstVisibleCol = 0;
}

void EnsureVisible(HWND h) {
    int sel = g_currentSlide - 1;
    if (sel < 0 || g_thumbs.empty()) return;
    int selRow = sel % ROWS_PER_COL;
    int selCol = sel / ROWS_PER_COL;
    int visRows = GetFullVisRows(h);
    int visCols = GetFullVisCols(h);
    if (selRow < g_firstVisibleIndex)
        g_firstVisibleIndex = selRow;
    else if (selRow >= g_firstVisibleIndex + visRows)
        g_firstVisibleIndex = selRow - visRows + 1;
    if (selCol < g_firstVisibleCol)
        g_firstVisibleCol = selCol;
    else if (selCol >= g_firstVisibleCol + visCols)
        g_firstVisibleCol = selCol - visCols + 1;
    ClampScroll(h);
}

int GetPPTCurrentSlide() {
    IDispatch* pApp = GetPPTApp(); if (!pApp) return 0;
    IDispatch* pres = Invoke(pApp, L"ActivePresentation");
    IDispatch* wins = pres ? Invoke(pres, L"Windows") : nullptr;
    int result = 0;
    if (wins) {
        VARIANT v; v.vt = VT_I4; v.lVal = 1; DISPPARAMS d = { &v,nullptr,1,0 };
        IDispatch* win = Invoke(wins, L"Item", &d);
        IDispatch* view = win ? Invoke(win, L"View") : nullptr;
        IDispatch* slide = view ? Invoke(view, L"Slide") : nullptr;
        if (slide) { result = InvokeInt(slide, L"SlideIndex"); slide->Release(); }
        if (view) view->Release();
        if (win) win->Release();
        wins->Release();
    }
    if (pres) pres->Release();
    pApp->Release();
    return result;
}

void NavigateToSlide() {
    IDispatch* pApp = GetPPTApp(); if (!pApp) return;
    IDispatch* pres = Invoke(pApp, L"ActivePresentation");
    IDispatch* wins = Invoke(pres, L"Windows");
    if (!wins) { if (pres) pres->Release(); pApp->Release(); return; }
    VARIANT v; v.vt = VT_I4; v.lVal = 1; DISPPARAMS d = { &v,nullptr,1,0 };
    IDispatch* win  = Invoke(wins, L"Item", &d);
    IDispatch* view = win ? Invoke(win, L"View") : nullptr;
    if (view) {
        VARIANT vs; vs.vt = VT_I4; vs.lVal = g_currentSlide;
        DISPPARAMS ds = { &vs,nullptr,1,0 };
        Invoke(view, L"GotoSlide", &ds);
        view->Release();
    }
    if (win) win->Release();
    wins->Release();
    if (pres) pres->Release();
    pApp->Release();
}

// -- drawing -------------------------------------------------------------------
void DrawGrid(HDC dc, HWND h) {
    RECT cr; GetClientRect(h, &cr);
    HBRUSH bgBr = CreateSolidBrush(BG); FillRect(dc, &cr, bgBr); DeleteObject(bgBr);
    if (g_thumbs.empty()) return;

    int visRows = min(GetDrawRows(h), ROWS_PER_COL - g_firstVisibleIndex);
    int visCols = min(GetDrawCols(h), GetNumCols() - g_firstVisibleCol);
    if (visRows <= 0 || visCols <= 0) return;

    int fontSize = ZFontSize();
    int headerW = ZHeaderW(), headerH = ZHeaderH();
    HFONT fNorm = CreateFontA(fontSize,0,0,0,FW_NORMAL,0,0,0,DEFAULT_CHARSET,0,0,CLEARTYPE_QUALITY,0,"Segoe UI");
    HFONT fBold = CreateFontA(fontSize,0,0,0,FW_BOLD,  0,0,0,DEFAULT_CHARSET,0,0,CLEARTYPE_QUALITY,0,"Segoe UI");
    HBRUSH hBr  = CreateSolidBrush(HEADER_BG);
    SetBkMode(dc, TRANSPARENT);

    // Corner
    RECT co = {0,0,headerW,headerH}; FillRect(dc,&co,hBr);

    int cellW = ZCellW(), cellH = ZCellH();
    int thumbW = ZThumbW(), thumbH = ZThumbH();

    // Column headers
    SelectObject(dc, fBold); SetTextColor(dc, HEADER_TEXT);
    for (int c = 0; c < visCols; c++) {
        int x = headerW + c*cellW;
        RECT r = {x,0,x+cellW,headerH}; FillRect(dc,&r,hBr);
        char lbl[8]; ColLabel(g_firstVisibleCol + c,lbl,8); DrawTextA(dc,lbl,-1,&r,DT_CENTER|DT_VCENTER|DT_SINGLELINE);
        HPEN p = CreatePen(PS_SOLID,1,DIVIDER); SelectObject(dc,p);
        MoveToEx(dc,x+cellW-1,0,NULL); LineTo(dc,x+cellW-1,headerH); DeleteObject(p);
    }
    // Row headers
    for (int r = 0; r < visRows; r++) {
        int y = headerH + r*cellH;
        RECT rr = {0,y,headerW,y+cellH}; FillRect(dc,&rr,hBr);
        char lbl[8]; sprintf_s(lbl,"%d",g_firstVisibleIndex + r + 1);
        SelectObject(dc,fBold); SetTextColor(dc,HEADER_TEXT);
        DrawTextA(dc,lbl,-1,&rr,DT_CENTER|DT_VCENTER|DT_SINGLELINE);
        HPEN p = CreatePen(PS_SOLID,1,DIVIDER); SelectObject(dc,p);
        MoveToEx(dc,0,y+cellH-1,NULL); LineTo(dc,headerW,y+cellH-1); DeleteObject(p);
    }
    DeleteObject(hBr);

    // Thumbnails (column-major layout)
    SelectObject(dc, fNorm);
    if (g_zoomLevel >= 1.0f) {
        SetStretchBltMode(dc, HALFTONE);
        SetBrushOrgEx(dc, 0, 0, NULL);
    } else {
        SetStretchBltMode(dc, COLORONCOLOR);
    }
    HDC mem = CreateCompatibleDC(dc);
    HPEN normPen = CreatePen(PS_SOLID, 1, DIVIDER);
    HPEN selPen  = CreatePen(PS_SOLID, 2, SEL_BORDER);
    HBRUSH selBr = CreateSolidBrush(SEL_BG);
    HBRUSH thumbBr = CreateSolidBrush(THUMB_BG);
    for (int col = 0; col < visCols; col++) {
        for (int row = 0; row < visRows; row++) {
            int idx = (g_firstVisibleCol + col) * ROWS_PER_COL + (g_firstVisibleIndex + row);
            if (idx < 0 || idx >= (int)g_thumbs.size()) continue;
            int x = headerW + col*cellW + PADDING/2;
            int y = headerH + row*cellH + PADDING/2;
            bool sel = (g_thumbs[idx].slideIndex == g_currentSlide);
            if (sel) {
                RECT sr = {x-2,y-2,x+thumbW+2,y+thumbH+2};
                FillRect(dc,&sr,selBr);
            }
            if (g_thumbs[idx].valid && g_thumbs[idx].hbmp) {
                HBITMAP ob = (HBITMAP)SelectObject(mem, g_thumbs[idx].hbmp);
                StretchBlt(dc,x,y,thumbW,thumbH,mem,0,0,SRC_THUMB_W,SRC_THUMB_H,SRCCOPY);
                SelectObject(mem,ob);
            } else {
                RECT ph = {x,y,x+thumbW,y+thumbH};
                FillRect(dc,&ph,thumbBr);
            }
            HGDIOBJ op = SelectObject(dc, sel ? selPen : normPen);
            SelectObject(dc,GetStockObject(NULL_BRUSH));
            Rectangle(dc,x,y,x+thumbW,y+thumbH);
            SelectObject(dc,op);
        }
    }
    DeleteDC(mem);
    DeleteObject(normPen);
    DeleteObject(selPen);
    DeleteObject(selBr);
    DeleteObject(thumbBr);

    // Custom scrollbar
    int totalRows = ROWS_PER_COL;
    int fullVisRows = GetFullVisRows(h);
    int trackH = cr.bottom - headerH;
    if (trackH > 0 && totalRows > fullVisRows) {
        int trackX = cr.right - SCROLLBAR_W;
        RECT track = {trackX, headerH, cr.right, cr.bottom};
        HBRUSH tb = CreateSolidBrush(SCROLL_TRACK); FillRect(dc,&track,tb); DeleteObject(tb);
        int thumbSH = max(20, trackH * fullVisRows / totalRows);
        int thumbSY = headerH + (trackH - thumbSH) * g_firstVisibleIndex / max(1, totalRows - fullVisRows);
        RECT thumbR = {trackX, thumbSY, cr.right, thumbSY + thumbSH};
        HBRUSH stb = CreateSolidBrush(SCROLL_THUMB); FillRect(dc,&thumbR,stb); DeleteObject(stb);
    }

    DeleteObject(fNorm); DeleteObject(fBold);
}

bool HandleClick(HWND h, int mx, int my) {
    int headerW = ZHeaderW(), headerH = ZHeaderH();
    if (mx < headerW || my < headerH) return false;
    int cellW = ZCellW(), cellH = ZCellH();
    int thumbW = ZThumbW(), thumbH = ZThumbH();
    int col = (mx - headerW) / cellW;
    int row = (my - headerH) / cellH;
    int idx = (g_firstVisibleCol + col) * ROWS_PER_COL + (g_firstVisibleIndex + row);
    if (idx < 0 || idx >= (int)g_thumbs.size()) return false;
    int x = headerW + col*cellW + PADDING/2;
    int y = headerH + row*cellH + PADDING/2;
    if (mx<x||mx>x+thumbW||my<y||my>y+thumbH) return false;
    g_currentSlide = g_thumbs[idx].slideIndex;
    SetFocus(h); EnsureVisible(h); InvalidateRect(h,NULL,FALSE); NavigateToSlide();
    PostMessage(h, WM_APP + 2, 0, 0);
    return true;
}

// -- splitter resize -----------------------------------------------------------
// Directly position all known sibling windows atomically.
// ONLY called during active mouse drag (from UpdateDrag).
void DoSplitterResize(int newPaneW) {
    if (!g_slidePanel) return;
    HWND parent = GetParent(g_slidePanel); if (!parent) return;

    RECT parentRc, panelRc;
    GetWindowRect(parent, &parentRc); GetWindowRect(g_slidePanel, &panelRc);
    int parentW = parentRc.right - parentRc.left;
    newPaneW = max(PANE_MIN_W, min(parentW - PANE_MIN_W, newPaneW));
    g_dragPaneW = newPaneW;

    RECT panelIP = panelRc;
    MapWindowPoints(HWND_DESKTOP, parent, reinterpret_cast<LPPOINT>(&panelIP), 2);
    int pL = panelIP.left, pT = panelIP.top, pB = panelIP.bottom;

    int gutterL   = pL + newPaneW;
    int splitterL = gutterL + g_gutterW;
    int contentL  = splitterL + g_splitterW;
    int contentR = parentW - 20;

    HDWP hdwp = BeginDeferWindowPos(8);

    hdwp = DeferWindowPos(hdwp, g_slidePanel, nullptr,
        pL, pT, newPaneW, pB - pT,
        SWP_NOZORDER | SWP_NOACTIVATE | SWP_NOOWNERZORDER);

    if (g_gutterHwnd) {
        RECT r; GetWindowRect(g_gutterHwnd, &r);
        RECT rip = r; MapWindowPoints(HWND_DESKTOP, parent, reinterpret_cast<LPPOINT>(&rip), 2);
        hdwp = DeferWindowPos(hdwp, g_gutterHwnd, nullptr,
            gutterL, rip.top, g_gutterW, rip.bottom - rip.top,
            SWP_NOZORDER | SWP_NOACTIVATE | SWP_NOOWNERZORDER);
    }

    if (g_splitterBar) {
        RECT r; GetWindowRect(g_splitterBar, &r);
        RECT rip = r; MapWindowPoints(HWND_DESKTOP, parent, reinterpret_cast<LPPOINT>(&rip), 2);
        hdwp = DeferWindowPos(hdwp, g_splitterBar, nullptr,
            splitterL, rip.top, g_splitterW, rip.bottom - rip.top,
            SWP_NOZORDER | SWP_NOACTIVATE | SWP_NOOWNERZORDER);
    }

    if (g_tabBar) {
        RECT r; GetWindowRect(g_tabBar, &r);
        RECT rip = r; MapWindowPoints(HWND_DESKTOP, parent, reinterpret_cast<LPPOINT>(&rip), 2);
        hdwp = DeferWindowPos(hdwp, g_tabBar, nullptr,
            pL, rip.top, newPaneW + g_gutterW, rip.bottom - rip.top,
            SWP_NOZORDER | SWP_NOACTIVATE | SWP_NOOWNERZORDER);
    }

    if (g_mainSlide) {
        RECT r; GetWindowRect(g_mainSlide, &r);
        RECT rip = r; MapWindowPoints(HWND_DESKTOP, parent, reinterpret_cast<LPPOINT>(&rip), 2);
        // Expand main slide to fill notes area
        RECT panelRc; GetWindowRect(g_slidePanel, &panelRc);
        RECT panelIP = panelRc; MapWindowPoints(HWND_DESKTOP, parent, reinterpret_cast<LPPOINT>(&panelIP), 2);
        int fullH = panelIP.bottom - rip.top;
        hdwp = DeferWindowPos(hdwp, g_mainSlide, nullptr,
            contentL, rip.top, contentR - contentL, fullH,
            SWP_NOZORDER | SWP_NOACTIVATE | SWP_NOOWNERZORDER);
    }

    EndDeferWindowPos(hdwp);
    InvalidateRect(parent, nullptr, FALSE);
}

// -- EnforceLayout: actively SetWindowPos on all managed windows ---------------
static void EnforceLayout() {
    if (g_enforcing || g_dragging) return;
    g_enforcing = true;

    LayoutPositions lay;
    if (ComputeLayout(lay)) {
        HWND parent = GetParent(g_slidePanel);
        HDWP hdwp = BeginDeferWindowPos(8);

        hdwp = DeferWindowPos(hdwp, g_slidePanel, nullptr,
            lay.panelL, lay.panelT, lay.panelW, lay.panelB - lay.panelT,
            SWP_NOZORDER | SWP_NOACTIVATE | SWP_NOOWNERZORDER);

        auto posChild = [&](HWND h, int x, int cx) {
            if (!h) return;
            RECT r; GetWindowRect(h, &r);
            RECT rip = r; MapWindowPoints(HWND_DESKTOP, parent, reinterpret_cast<LPPOINT>(&rip), 2);
            hdwp = DeferWindowPos(hdwp, h, nullptr,
                x, rip.top, cx, rip.bottom - rip.top,
                SWP_NOZORDER | SWP_NOACTIVATE | SWP_NOOWNERZORDER);
        };

        posChild(g_gutterHwnd,  lay.gutterL,   g_gutterW);
        posChild(g_splitterBar, lay.splitterL,  g_splitterW);
        posChild(g_tabBar,      lay.panelL,     lay.panelW + g_gutterW);

        // Main slide: expand to full height (notes is hidden)
        if (g_mainSlide) {
            RECT r; GetWindowRect(g_mainSlide, &r);
            RECT rip = r; MapWindowPoints(HWND_DESKTOP, parent, reinterpret_cast<LPPOINT>(&rip), 2);
            int fullH = lay.panelB - rip.top;
            hdwp = DeferWindowPos(hdwp, g_mainSlide, nullptr,
                lay.contentL, rip.top, lay.contentR - lay.contentL, fullH,
                SWP_NOZORDER | SWP_NOACTIVATE | SWP_NOOWNERZORDER);
        }

        EndDeferWindowPos(hdwp);
        InvalidateRect(parent, nullptr, FALSE);
    }

    g_enforcing = false;
}

// -- drag helpers --------------------------------------------------------------
static void BeginDrag(HWND captureHwnd) {
    SetCapture(captureHwnd);
    g_dragging     = true;
    g_dragPaneW    = 0;
    g_desiredPaneW = 0; // clear so WM_WINDOWPOSCHANGING doesn't fight us
    POINT cur; GetCursorPos(&cur);
    g_dragStartX = cur.x;
    RECT r; GetWindowRect(g_slidePanel, &r);
    g_dragStartPaneW = r.right - r.left;
}

static void UpdateDrag() {
    POINT cur; GetCursorPos(&cur);
    DoSplitterResize(g_dragStartPaneW + (cur.x - g_dragStartX));
}

static void EndDrag() {
    RECT r; GetWindowRect(g_slidePanel, &r);
    g_desiredPaneW = r.right - r.left;
    g_dragging = false;
    ReleaseCapture();
    EnforceLayout();
}

// -- Shared WndProc for gutter and splitter bar --------------------------------
// During drag: suppress WM_WINDOWPOSCHANGING/CHANGED so PPT can't fight.
// After drag (g_desiredPaneW > 0): enforce correct x-position silently.
static LRESULT HandleDragWindow(HWND hwnd, UINT msg, WPARAM wp, LPARAM lp, WNDPROC orig) {
    switch (msg) {
    case WM_SETCURSOR:
        SetCursor(LoadCursor(nullptr, IDC_SIZEWE));
        return TRUE;

    case WM_LBUTTONDOWN:
        BeginDrag(hwnd);
        return 0;

    case WM_MOUSEMOVE:
        if (g_dragging) {
            UpdateDrag();
            InvalidateRect(g_slidePanel, nullptr, FALSE);
            return 0;
        }
        SetCursor(LoadCursor(nullptr, IDC_SIZEWE));
        break;

    case WM_LBUTTONUP:
        if (g_dragging) { EndDrag(); return 0; }
        break;

    case WM_CAPTURECHANGED:
        if (g_dragging) g_dragging = false; // fallback safety
        break;

    case WM_WINDOWPOSCHANGING: {
        WINDOWPOS* wp2 = reinterpret_cast<WINDOWPOS*>(lp);
        if (wp2) {
            int paneW = 0;
            if (g_dragging && g_dragPaneW > 0)
                paneW = g_dragPaneW;
            else if (!g_dragging && g_desiredPaneW > 0)
                paneW = g_desiredPaneW;

            if (paneW > 0) {
                // Compute correct x from panel position + pane width
                HWND parent = GetParent(hwnd);
                if (parent) {
                    RECT panelRc; GetWindowRect(g_slidePanel, &panelRc);
                    RECT panelIP = panelRc;
                    MapWindowPoints(HWND_DESKTOP, parent, reinterpret_cast<LPPOINT>(&panelIP), 2);
                    int gutterL = panelIP.left + paneW;
                    int splitterL = gutterL + g_gutterW;
                    wp2->x = (hwnd == g_gutterHwnd) ? gutterL : splitterL;
                    wp2->flags &= ~SWP_NOMOVE;
                    if (g_dragging) return 0; // suppress PPT's proc during drag
                    break; // post-drag: let original proc run with corrected values
                }
            }
        }
        break;
    }

    case WM_PAINT: {
        PAINTSTRUCT ps; HDC dc = BeginPaint(hwnd, &ps);
        RECT rc; GetClientRect(hwnd, &rc);
        HBRUSH br = CreateSolidBrush(BG); FillRect(dc, &rc, br); DeleteObject(br);
        EndPaint(hwnd, &ps);
        return 0;
    }
    case WM_ERASEBKGND:
        return 1;
    }
    return CallWindowProcW(orig, hwnd, msg, wp, lp);
}

LRESULT CALLBACK HookedGutterWndProc(HWND h, UINT msg, WPARAM wp, LPARAM lp) {
    return HandleDragWindow(h, msg, wp, lp, g_origGutter);
}
LRESULT CALLBACK HookedSplitterWndProc(HWND h, UINT msg, WPARAM wp, LPARAM lp) {
    return HandleDragWindow(h, msg, wp, lp, g_origSplitter);
}

// -- Content-area WndProcs (main slide, h-splitter, notes, tab bar) -----------
// These enforce x and cx when g_desiredPaneW > 0.
static LRESULT HandleContentWindow(HWND hwnd, UINT msg, WPARAM wp, LPARAM lp, WNDPROC orig) {
    // Force notes pane to stay invisible: zero size, no scrollbars, no painting
    if (hwnd == g_notesPane) {
        if (msg == WM_WINDOWPOSCHANGING) {
            WINDOWPOS* wp2 = reinterpret_cast<WINDOWPOS*>(lp);
            if (wp2) {
                wp2->cx = 0; wp2->cy = 0;
                wp2->flags &= ~(SWP_SHOWWINDOW | SWP_NOSIZE);
                wp2->flags |= SWP_HIDEWINDOW;
            }
            return 0;
        }
        if (msg == WM_NCCALCSIZE || msg == WM_NCPAINT || msg == WM_PAINT ||
            msg == WM_VSCROLL || msg == WM_HSCROLL || msg == WM_ERASEBKGND)
            return 0;
        if (msg == WM_STYLECHANGING && wp == GWL_STYLE) {
            STYLESTRUCT* ss = reinterpret_cast<STYLESTRUCT*>(lp);
            ss->styleNew &= ~(WS_VSCROLL | WS_HSCROLL | WS_VISIBLE);
            return 0;
        }
        if (msg == WM_SHOWWINDOW && wp == TRUE)
            return 0; // already handled above but belt-and-suspenders
    }

    if (hwnd == g_hSplitter) {
        if (msg == WM_WINDOWPOSCHANGING) {
            WINDOWPOS* wp2 = reinterpret_cast<WINDOWPOS*>(lp);
            if (wp2) {
                wp2->cx = 0; wp2->cy = 0;
                wp2->flags &= ~(SWP_SHOWWINDOW | SWP_NOSIZE);
                wp2->flags |= SWP_HIDEWINDOW;
            }
            return 0;
        }
    }

    if (msg == WM_WINDOWPOSCHANGING && g_desiredPaneW > 0 && !g_dragging) {
        WINDOWPOS* wp2 = reinterpret_cast<WINDOWPOS*>(lp);
        if (wp2) {
            // Prevent notes/h-splitter from being shown
            if (hwnd == g_notesPane || hwnd == g_hSplitter) {
                wp2->flags &= ~SWP_SHOWWINDOW;
            }

            LayoutPositions lay;
            if (ComputeLayout(lay)) {
                if (hwnd == g_mainSlide) {
                    if (!(wp2->flags & SWP_NOMOVE))
                        wp2->x = lay.contentL;
                    if (!(wp2->flags & SWP_NOSIZE)) {
                        wp2->cx = lay.contentR - lay.contentL;
                        // Expand main slide to fill notes area
                        wp2->cy = lay.panelB - lay.panelT;
                    }
                }
                else if (hwnd == g_hSplitter || hwnd == g_notesPane) {
                    if (!(wp2->flags & SWP_NOMOVE))
                        wp2->x = lay.contentL;
                    if (!(wp2->flags & SWP_NOSIZE))
                        wp2->cx = lay.contentR - lay.contentL;
                }
                else if (hwnd == g_tabBar) {
                    if (!(wp2->flags & SWP_NOMOVE))
                        wp2->x = lay.panelL;
                    if (!(wp2->flags & SWP_NOSIZE))
                        wp2->cx = lay.panelW + g_gutterW;
                }
            }
        }
    }
    return CallWindowProcW(orig, hwnd, msg, wp, lp);
}

LRESULT CALLBACK HookedMainSlideWndProc(HWND h, UINT msg, WPARAM wp, LPARAM lp) {
    LRESULT r = HandleContentWindow(h, msg, wp, lp, g_origMainSlide);
    if (msg == WM_PAINT && g_slidePanel) {
        SetTimer(g_slidePanel, 2, 500, NULL);
    }
    return r;
}
LRESULT CALLBACK HookedHSplitterWndProc(HWND h, UINT msg, WPARAM wp, LPARAM lp) {
    return HandleContentWindow(h, msg, wp, lp, g_origHSplitter);
}
LRESULT CALLBACK HookedNotesWndProc(HWND h, UINT msg, WPARAM wp, LPARAM lp) {
    return HandleContentWindow(h, msg, wp, lp, g_origNotes);
}
LRESULT CALLBACK HookedTabBarWndProc(HWND h, UINT msg, WPARAM wp, LPARAM lp) {
    return HandleContentWindow(h, msg, wp, lp, g_origTabBar);
}

// -- Cleanup / Re-init ---------------------------------------------------------
static void UnhookAll() {
    // Stop background thumb loader first
    g_thumbLoadCancel = true;
    if (g_thumbThread) {
        WaitForSingleObject(g_thumbThread, 5000);
        CloseHandle(g_thumbThread);
        g_thumbThread = nullptr;
    }
    g_thumbLoadCancel = false;

    if (g_kbHook) { UnhookWindowsHookEx(g_kbHook); g_kbHook = nullptr; }
    if (g_mouseHook) { UnhookWindowsHookEx(g_mouseHook); g_mouseHook = nullptr; }
    if (g_slidePanel && g_origPanel)
        SetWindowLongPtrW(g_slidePanel, GWLP_WNDPROC, (LONG_PTR)g_origPanel);
    if (g_gutterHwnd && g_origGutter)
        SetWindowLongPtrW(g_gutterHwnd, GWLP_WNDPROC, (LONG_PTR)g_origGutter);
    if (g_splitterBar && g_origSplitter)
        SetWindowLongPtrW(g_splitterBar, GWLP_WNDPROC, (LONG_PTR)g_origSplitter);
    if (g_mainSlide && g_origMainSlide)
        SetWindowLongPtrW(g_mainSlide, GWLP_WNDPROC, (LONG_PTR)g_origMainSlide);
    if (g_hSplitter && g_origHSplitter)
        SetWindowLongPtrW(g_hSplitter, GWLP_WNDPROC, (LONG_PTR)g_origHSplitter);
    if (g_notesPane && g_origNotes)
        SetWindowLongPtrW(g_notesPane, GWLP_WNDPROC, (LONG_PTR)g_origNotes);
    if (g_tabBar && g_origTabBar)
        SetWindowLongPtrW(g_tabBar, GWLP_WNDPROC, (LONG_PTR)g_origTabBar);

    g_origPanel = g_origGutter = g_origSplitter = nullptr;
    g_origMainSlide = g_origHSplitter = g_origNotes = g_origTabBar = nullptr;
    g_slidePanel = g_gutterHwnd = g_splitterBar = nullptr;
    g_mainSlide = g_hSplitter = g_notesPane = g_tabBar = nullptr;
    g_panelHasFocus = false;
    g_panelIsActive = false;
    g_desiredPaneW = 0;
    g_dragging = false;

    for (auto& t : g_thumbs) if (t.hbmp) DeleteObject(t.hbmp);
    g_thumbs.clear();
    g_slideCount = 0;
    g_currentSlide = 1;
    g_firstVisibleIndex = 0;
    g_firstVisibleCol = 0;
}

DWORD WINAPI InitThread(LPVOID); // forward declaration

static DWORD WINAPI ReInitThread(LPVOID param) {
    LogDebug(L"RE-INIT: starting");
    UnhookAll();
    return InitThread(param);
}

// -- Mouse hook (detects clicks outside our panel to deactivate arrow keys) ----
static LRESULT CALLBACK MouseHookProc(int code, WPARAM wp, LPARAM lp) {
    if (code >= 0 && (wp == WM_LBUTTONDOWN || wp == WM_NCLBUTTONDOWN)) {
        MOUSEHOOKSTRUCT* mhs = reinterpret_cast<MOUSEHOOKSTRUCT*>(lp);
        if (mhs && g_slidePanel) {
            g_panelIsActive = (mhs->hwnd == g_slidePanel);
        }
    }
    return CallNextHookEx(g_mouseHook, code, wp, lp);
}

// -- Keyboard hook (intercepts keys before PPT's TranslateAccelerator) ---------
static LRESULT CALLBACK KeyboardHookProc(int code, WPARAM wp, LPARAM lp) {
    if (code >= 0 && !(lp & 0x80000000)) {
        // Ctrl+Shift+R: re-init (works regardless of panel focus)
        if (wp == 'R' && (GetKeyState(VK_CONTROL) & 0x8000) && (GetKeyState(VK_SHIFT) & 0x8000)) {
            CreateThread(NULL, 0, ReInitThread, NULL, 0, NULL);
            return 1;
        }
        // '=' key: trigger formula evaluation when NOT in panel
        if (!g_panelIsActive && g_slidePanel) {
            BYTE kbState[256]; GetKeyboardState(kbState);
            wchar_t ch[4] = {};
            int ret = ToUnicode((UINT)wp, (lp >> 16) & 0xFF, kbState, ch, 4, 0);
            if (ret == 1 && ch[0] == L'=') {
                PostMessage(g_slidePanel, WM_APP + 10, 0, 0);
            }
        }
        // Arrow/nav keys: when panel is active (last clicked)
        if (g_panelIsActive && g_slidePanel) {
            switch (wp) {
            case VK_UP: case VK_DOWN: case VK_LEFT: case VK_RIGHT:
            case VK_PRIOR: case VK_NEXT: case VK_HOME: case VK_END:
                if (!g_panelHasFocus) SetFocus(g_slidePanel);
                SendMessage(g_slidePanel, WM_KEYDOWN, wp, lp);
                return 1;
            }
        }
    }
    return CallNextHookEx(g_kbHook, code, wp, lp);
}

// -- Panel WndProc -------------------------------------------------------------
LRESULT CALLBACK HookedPanelWndProc(HWND hwnd, UINT msg, WPARAM wp, LPARAM lp) {
    switch (msg) {
    case WM_PAINT: {
        PAINTSTRUCT ps; HDC dc = BeginPaint(hwnd, &ps);
        RECT rc; GetClientRect(hwnd, &rc);
        int w = rc.right, h = rc.bottom;
        if (w > 0 && h > 0) {
            HDC m = CreateCompatibleDC(dc);
            HBITMAP mb = CreateCompatibleBitmap(dc, w, h);
            HBITMAP ob = (HBITMAP)SelectObject(m, mb);
            DrawGrid(m, hwnd);
            BitBlt(dc, 0, 0, w, h, m, 0, 0, SRCCOPY);
            SelectObject(m, ob); DeleteObject(mb); DeleteDC(m);
        }
        EndPaint(hwnd, &ps);
        return 0;
    }
    case WM_ERASEBKGND:
        return 1;

    case WM_NCCALCSIZE:
        if (wp) {
            LRESULT r = CallWindowProcW(g_origPanel, hwnd, msg, wp, lp);
            LONG s = GetWindowLong(hwnd, GWL_STYLE);
            if (s & (WS_VSCROLL | WS_HSCROLL))
                SetWindowLong(hwnd, GWL_STYLE, s & ~(WS_VSCROLL | WS_HSCROLL));
            return r;
        }
        break;

    case WM_VSCROLL:
    case WM_HSCROLL:
        return 0;

    // Clamp removal: when we have a desired width, enforce it and bypass
    // PPT's original proc (where the max-width clamp lives).
    case WM_WINDOWPOSCHANGING: {
        WINDOWPOS* wp2 = reinterpret_cast<WINDOWPOS*>(lp);
        if (wp2 && !(wp2->flags & SWP_NOSIZE)) {
            int want = 0;
            if (g_dragging && g_dragPaneW > 0) {
                want = g_dragPaneW;
            } else if (g_desiredPaneW > 0) {
                want = g_desiredPaneW;
            }
            if (want > 0) {
                wp2->cx = want;
                return 0; // must bypass PPT's max-width clamp
            }
        }
        break;
    }

    case WM_WINDOWPOSCHANGED:
        if (g_desiredPaneW > 0 && !g_enforcing)
            PostMessage(hwnd, WM_APP + 1, 0, 0);
        break;

    case WM_APP + 1:
        EnforceLayout();
        return 0;

    case WM_APP + 2:
        SetFocus(hwnd);
        return 0;

    case WM_APP + 3: {
        DWORD tid = GetWindowThreadProcessId(hwnd, nullptr);
        if (!g_kbHook)
            g_kbHook = SetWindowsHookExW(WH_KEYBOARD, KeyboardHookProc, nullptr, tid);
        if (!g_mouseHook)
            g_mouseHook = SetWindowsHookExW(WH_MOUSE, MouseHookProc, nullptr, tid);
        return 0;
    }

    case WM_GETDLGCODE:
        return DLGC_WANTARROWS;

    case WM_LBUTTONDOWN:
        if (HandleClick(hwnd, GET_X_LPARAM(lp), GET_Y_LPARAM(lp))) return 0;
        break;

    case WM_KEYDOWN: {
        int visRows = GetFullVisRows(hwnd);
        int ns = g_currentSlide;
        switch (wp) {
        case VK_UP:    if (ns > 1) ns--; break;
        case VK_DOWN:  if (ns < g_slideCount) ns++; break;
        case VK_LEFT: {
            int newIdx = (ns - 1) - ROWS_PER_COL;
            if (newIdx >= 0) ns = newIdx + 1;
            break;
        }
        case VK_RIGHT: {
            int newIdx = (ns - 1) + ROWS_PER_COL;
            if (newIdx < g_slideCount) ns = newIdx + 1;
            break;
        }
        case VK_PRIOR: ns = max(1, ns - visRows); break;
        case VK_NEXT:  ns = min(g_slideCount, ns + visRows); break;
        case VK_HOME:  ns = 1; break;
        case VK_END:   ns = g_slideCount; break;
        }
        if (ns != g_currentSlide) {
            g_currentSlide = ns; EnsureVisible(hwnd);
            InvalidateRect(hwnd, NULL, FALSE); NavigateToSlide();
            PostMessage(hwnd, WM_APP + 2, 0, 0);
        }
        // Always consume arrow/nav keys so PPT doesn't handle them
        switch (wp) {
        case VK_UP: case VK_DOWN: case VK_LEFT: case VK_RIGHT:
        case VK_PRIOR: case VK_NEXT: case VK_HOME: case VK_END:
            return 0;
        }
        break;
    }

    case WM_MOUSEWHEEL: {
        int delta = GET_WHEEL_DELTA_WPARAM(wp);
        if (GetKeyState(VK_CONTROL) & 0x8000) {
            // Get mouse position relative to client area
            POINT mp; GetCursorPos(&mp); ScreenToClient(hwnd, &mp);
            int headerW = ZHeaderW(), headerH = ZHeaderH();
            // Fractional row/col under cursor before zoom
            float oldCellW = (float)ZCellW(), oldCellH = (float)ZCellH();
            float fracCol = g_firstVisibleCol + (mp.x - headerW) / oldCellW;
            float fracRow = g_firstVisibleIndex + (mp.y - headerH) / oldCellH;
            // Apply zoom
            float step = 0.1f;
            g_zoomLevel += (delta > 0) ? step : -step;
            if (g_zoomLevel < 0.25f) g_zoomLevel = 0.25f;
            if (g_zoomLevel > 3.0f) g_zoomLevel = 3.0f;
            // Adjust scroll so same position stays under cursor
            float newCellW = (float)ZCellW(), newCellH = (float)ZCellH();
            g_firstVisibleCol = (int)(fracCol - (mp.x - headerW) / newCellW);
            g_firstVisibleIndex = (int)(fracRow - (mp.y - headerH) / newCellH);
            ClampScroll(hwnd);
            InvalidateRect(hwnd, NULL, FALSE);
            return 0;
        }
        int lines = -delta / WHEEL_DELTA * 3;
        int visRows = GetFullVisRows(hwnd);
        int maxFirst = max(0, ROWS_PER_COL - visRows);
        int pos = max(0, min(g_firstVisibleIndex + lines, maxFirst));
        if (pos != g_firstVisibleIndex) {
            g_firstVisibleIndex = pos;
            InvalidateRect(hwnd, NULL, FALSE);
        }
        return 0;
    }

    case WM_SIZE:
        ClampScroll(hwnd);
        InvalidateRect(hwnd, NULL, FALSE);
        return 0;

    case WM_STYLECHANGING: {
        if (wp == GWL_STYLE) {
            STYLESTRUCT* ss = reinterpret_cast<STYLESTRUCT*>(lp);
            ss->styleNew &= ~(WS_VSCROLL | WS_HSCROLL);
        }
        return 0;
    }

    case WM_TIMER:
        if (wp == 10) {
            KillTimer(hwnd, 10);
            HandleFormulaEvaluation();
            return 0;
        }
        if (wp == 2) {
            KillTimer(hwnd, 2);
            int pptSlide = GetPPTCurrentSlide();
            if (pptSlide > 0 && pptSlide != g_currentSlide) {
                g_currentSlide = pptSlide;
                EnsureVisible(hwnd);
            }
            IDispatch* pApp = GetPPTApp();
            if (pApp) {
                int idx = g_currentSlide - 1;
                if (idx >= 0 && idx < (int)g_thumbs.size()) {
                    if (g_thumbs[idx].hbmp) DeleteObject(g_thumbs[idx].hbmp);
                    g_thumbs[idx].hbmp = GetSlideThumbnail(pApp, g_currentSlide);
                    g_thumbs[idx].valid = (g_thumbs[idx].hbmp != nullptr);
                }
                IDispatch* pres = Invoke(pApp, L"ActivePresentation");
                IDispatch* slides = Invoke(pres, L"Slides");
                if (slides) {
                    int newCount = InvokeInt(slides, L"Count");
                    if (newCount != g_slideCount) {
                        RefreshThumbs(pApp);
                        g_thumbThread = CreateThread(NULL, 0, ThumbLoaderThread, NULL, 0, NULL);
                    }
                    slides->Release();
                }
                if (pres) pres->Release();
                pApp->Release();
                InvalidateRect(hwnd, NULL, FALSE);
            }
        }
        return 0;

    case WM_APP + 10:
        SetTimer(hwnd, 10, 150, NULL);
        return 0;

    case WM_SETFOCUS:
        g_panelHasFocus = true;
        EnsureVisible(hwnd); InvalidateRect(hwnd, NULL, FALSE); return 0;
    case WM_KILLFOCUS:
        g_panelHasFocus = false;
        return 0;
    }
    return CallWindowProcW(g_origPanel, hwnd, msg, wp, lp);
}

// -- find pane -----------------------------------------------------------------
BOOL CALLBACK FindMainPptWindow(HWND hwnd, LPARAM) {
    DWORD pid = 0; GetWindowThreadProcessId(hwnd, &pid);
    if (pid != GetCurrentProcessId() || !IsWindowVisible(hwnd) || GetParent(hwnd)) return TRUE;
    RECT rc; if (!GetWindowRect(hwnd, &rc)) return TRUE;
    int a = (rc.right-rc.left)*(rc.bottom-rc.top);
    if (a > g_bestMainArea) { g_bestMainArea = a; g_mainPptWindow = hwnd; }
    return TRUE;
}

BOOL CALLBACK FindPaneInChildren(HWND hwnd, LPARAM) {
    wchar_t cls[256]; GetClassNameW(hwnd, cls, 256);
    if (wcscmp(cls, L"paneClassDC") == 0 && g_mainPptWindow) {
        RECT pr={}, mr={};
        GetWindowRect(hwnd, &pr); GetWindowRect(g_mainPptWindow, &mr);
        int mw = mr.right-mr.left, mh = mr.bottom-mr.top;
        int w = pr.right-pr.left, h = pr.bottom-pr.top;
        int lo = pr.left-mr.left, to = pr.top-mr.top;
        if (w>=100 && h>=200 && w<=(mw*55)/100 && h<=(mh*95)/100
            && lo>=-2 && lo<=(mw*25)/100 && to>=40) {
            int score = (5000 - lo*20) + h*2 + w/2;
            if (score > g_bestPaneScore) { g_bestPaneScore = score; g_foundHwnd = hwnd; }
        }
    }
    EnumChildWindows(hwnd, FindPaneInChildren, 0);
    return TRUE;
}

// -- init ----------------------------------------------------------------------
DWORD WINAPI InitThread(LPVOID) {
    if (FAILED(CoInitialize(nullptr))) return 1;
    Sleep(2000);

    g_foundHwnd = nullptr; g_mainPptWindow = nullptr;
    g_bestPaneScore = -2147483647; g_bestMainArea = 0;
    EnumWindows(FindMainPptWindow, 0);
    if (!g_mainPptWindow) { CoUninitialize(); return 1; }
    EnumChildWindows(g_mainPptWindow, FindPaneInChildren, 0);
    g_slidePanel = g_foundHwnd;
    if (!g_slidePanel) { LogDebug(L"panel not found"); CoUninitialize(); return 1; }

    IDispatch* pApp = GetPPTApp();
    if (pApp) { RefreshThumbs(pApp); pApp->Release(); }

    // Start background thumbnail loader
    g_thumbThread = CreateThread(NULL, 0, ThumbLoaderThread, NULL, 0, NULL);

    // Enumerate siblings — geometry-only detection (no title matching)
    {
        RECT pRc; GetWindowRect(g_slidePanel, &pRc);
        wchar_t buf[512];
        swprintf_s(buf, L"OUR PANEL rect=[%d,%d,%d,%d]", pRc.left,pRc.top,pRc.right,pRc.bottom);
        LogDebug(buf);

        HWND parent = GetParent(g_slidePanel);

        // First pass: find gutter and vertical splitter (needed to locate content area)
        int si = 0;
        for (HWND s = GetWindow(parent, GW_CHILD); s; s = GetWindow(s, GW_HWNDNEXT), si++) {
            if (s == g_slidePanel) continue;
            wchar_t cls[128]={}, title[128]={};
            GetClassNameW(s, cls, 128);
            GetWindowTextW(s, title, 128);
            RECT r; GetWindowRect(s, &r);
            int w = r.right-r.left, h = r.bottom-r.top;
            bool vis = IsWindowVisible(s) != 0;

            swprintf_s(buf, L"  SIB[%d] cls=%s title=%s w=%d h=%d vis=%d rect=[%d,%d,%d,%d]",
                si, cls, title, w, h, (int)vis, r.left, r.top, r.right, r.bottom);
            LogDebug(buf);

            if (!vis) continue;

            // Gutter: paneClass, narrow, directly right of our panel
            if (!g_gutterHwnd && wcscmp(cls, L"paneClass") == 0
                && w <= 40 && abs(r.left - pRc.right) <= 4) {
                g_gutterHwnd = s; g_gutterW = w;
                LogDebug(L"  -> GUTTER");
            }
            // Vertical splitter bar: splitterBar, narrow, tall
            else if (!g_splitterBar && wcscmp(cls, L"splitterBar") == 0
                && w <= 12 && h > 200 && r.left > pRc.right - 4) {
                g_splitterBar = s; g_splitterW = w;
                LogDebug(L"  -> VERTICAL SPLITTER");
            }
            // Horizontal splitter bar: splitterBar, wide, short
            else if (!g_hSplitter && wcscmp(cls, L"splitterBar") == 0 && w > 200 && h <= 12) {
                g_hSplitter = s;
                LogDebug(L"  -> HORIZONTAL SPLITTER");
            }
            // Tab bar: paneClass, short, starts at/near panel left
            else if (!g_tabBar && wcscmp(cls, L"paneClass") == 0
                && h > 0 && h <= 40 && w > 100 && abs(r.left - pRc.left) <= 4) {
                g_tabBar = s;
                LogDebug(L"  -> TAB BAR");
            }
        }

        // Second pass: find main slide and notes by geometry
        // Main slide = largest-area paneClassDC sibling right of splitter
        // Notes = smaller paneClassDC sibling right of splitter
        int splitterRight = 0;
        if (g_splitterBar) {
            RECT sr; GetWindowRect(g_splitterBar, &sr);
            splitterRight = sr.right;
        } else if (g_gutterHwnd) {
            RECT gr; GetWindowRect(g_gutterHwnd, &gr);
            splitterRight = gr.right + g_splitterW;
        } else {
            splitterRight = pRc.right + g_gutterW + g_splitterW;
        }

        int bestMainArea = 0;
        HWND bestMainHwnd = nullptr;
        int bestNotesArea = 0;
        HWND bestNotesHwnd = nullptr;

        for (HWND s = GetWindow(parent, GW_CHILD); s; s = GetWindow(s, GW_HWNDNEXT)) {
            if (s == g_slidePanel) continue;
            if (!IsWindowVisible(s)) continue;
            wchar_t cls[128]; GetClassNameW(s, cls, 128);
            if (wcscmp(cls, L"paneClassDC") != 0) continue;

            RECT r; GetWindowRect(s, &r);
            if (r.left < splitterRight - 4) continue; // must be right of splitter

            int w = r.right - r.left, h = r.bottom - r.top;
            int area = w * h;
            if (area > bestMainArea) {
                // Demote previous best to notes candidate
                if (bestMainHwnd) { bestNotesArea = bestMainArea; bestNotesHwnd = bestMainHwnd; }
                bestMainArea = area; bestMainHwnd = s;
            } else if (area > bestNotesArea) {
                bestNotesArea = area; bestNotesHwnd = s;
            }
        }

        if (bestMainHwnd) {
            g_mainSlide = bestMainHwnd;
            LogDebug(L"  -> MAIN SLIDE (by area)");
        }
        if (bestNotesHwnd) {
            g_notesPane = bestNotesHwnd;
            LogDebug(L"  -> NOTES (by area)");
        }
    }

    // Hook our panel
    g_origPanel = (WNDPROC)SetWindowLongPtrW(g_slidePanel, GWLP_WNDPROC, (LONG_PTR)HookedPanelWndProc);

    // Strip any native scrollbar PPT may have added; we scroll via mouse wheel only
    LONG style = GetWindowLong(g_slidePanel, GWL_STYLE);
    if (style & (WS_VSCROLL | WS_HSCROLL)) {
        SetWindowLong(g_slidePanel, GWL_STYLE, style & ~(WS_VSCROLL | WS_HSCROLL));
        SetWindowPos(g_slidePanel, nullptr, 0, 0, 0, 0,
            SWP_NOMOVE | SWP_NOSIZE | SWP_NOZORDER | SWP_FRAMECHANGED);
    }

    // Hook gutter and splitter (drag handles)
    if (g_gutterHwnd)
        g_origGutter = (WNDPROC)SetWindowLongPtrW(g_gutterHwnd, GWLP_WNDPROC, (LONG_PTR)HookedGutterWndProc);
    if (g_splitterBar)
        g_origSplitter = (WNDPROC)SetWindowLongPtrW(g_splitterBar, GWLP_WNDPROC, (LONG_PTR)HookedSplitterWndProc);

    // Hook content-area windows (position enforcement only)
    if (g_mainSlide)
        g_origMainSlide = (WNDPROC)SetWindowLongPtrW(g_mainSlide, GWLP_WNDPROC, (LONG_PTR)HookedMainSlideWndProc);
    if (g_hSplitter) {
        g_origHSplitter = (WNDPROC)SetWindowLongPtrW(g_hSplitter, GWLP_WNDPROC, (LONG_PTR)HookedHSplitterWndProc);
        ShowWindow(g_hSplitter, SW_HIDE);
    }
    if (g_notesPane) {
        g_origNotes = (WNDPROC)SetWindowLongPtrW(g_notesPane, GWLP_WNDPROC, (LONG_PTR)HookedNotesWndProc);
        ShowWindow(g_notesPane, SW_HIDE);
    }
    if (g_tabBar)
        g_origTabBar = (WNDPROC)SetWindowLongPtrW(g_tabBar, GWLP_WNDPROC, (LONG_PTR)HookedTabBarWndProc);

    // Set initial desired width so layout enforcement kicks in immediately
    {
        RECT pr; GetWindowRect(g_slidePanel, &pr);
        g_desiredPaneW = pr.right - pr.left;
    }

    // Install keyboard hook on the panel's UI thread, then enforce layout
    PostMessage(g_slidePanel, WM_APP + 3, 0, 0);
    PostMessage(g_slidePanel, WM_APP + 1, 0, 0);

    // Init summary
    {
        wchar_t b[512];
        swprintf_s(b, L"INIT SUMMARY: panel=%s gutter=%s splitter=%s mainSlide=%s hSplitter=%s notes=%s tabBar=%s  gutterW=%d splitterW=%d",
            g_slidePanel  ? L"OK" : L"MISSING",
            g_gutterHwnd  ? L"OK" : L"MISSING",
            g_splitterBar ? L"OK" : L"MISSING",
            g_mainSlide   ? L"OK" : L"MISSING",
            g_hSplitter   ? L"OK" : L"MISSING",
            g_notesPane   ? L"OK" : L"MISSING",
            g_tabBar      ? L"OK" : L"MISSING",
            g_gutterW, g_splitterW);
        LogDebug(b);
    }

    InvalidateRect(g_slidePanel, NULL, TRUE);
    CoUninitialize();
    return 0;
}

BOOL APIENTRY DllMain(HMODULE hModule, DWORD reason, LPVOID) {
    if (reason == DLL_PROCESS_ATTACH) {
        DisableThreadLibraryCalls(hModule);
        CreateThread(NULL, 0, InitThread, NULL, 0, NULL);
    }
    return TRUE;
}

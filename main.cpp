/* Please note: This code was created by AI, but has not been verified for copyright infringement. */


#define _WIN32_WINNT 0x0A00
#include <windows.h>
#include <commctrl.h>      // Trackbar
#include <mmdeviceapi.h>
#include <audiopolicy.h>
#include <audioclient.h>
#include <string>
#include <vector>
#include <set>
#include <algorithm>

#pragma comment(lib, "Ole32.lib")
#pragma comment(lib, "Uuid.lib")
#pragma comment(lib, "Comctl32.lib")

// ===== UI IDs / Messages =====
#define IDC_COMBO_A     1001
#define IDC_COMBO_B     1002
#define IDC_TRACK       1003
#define WMAPP_REFRESH   (WM_APP + 1)

// ===== Helpers =====
static std::wstring ToLower(const std::wstring& s) {
    std::wstring t = s;
    for (size_t i = 0; i < t.size(); ++i) t[i] = (wchar_t)towlower(t[i]);
    return t;
}

// DisplayName が空の時、SessionID から "xxx.exe" を推定
static std::wstring NormalizeNameFromSessionId(const std::wstring& sid) {
    if (sid.empty()) return L"unknown";
    size_t sep = sid.find_last_of(L"\\/");
    std::wstring tail = (sep == std::wstring::npos) ? sid : sid.substr(sep + 1);
    size_t cut1 = tail.find_first_of(L"% \t\r\n");
    if (cut1 != std::wstring::npos) tail = tail.substr(0, cut1);
    std::wstring lower = ToLower(tail);
    size_t exep = lower.find(L".exe");
    if (exep != std::wstring::npos) tail = tail.substr(0, exep + 4);
    if (tail.empty()) return L"unknown";
    return tail;
}

static const wchar_t* SessionStateToString(AudioSessionState s) {
    switch (s) {
    case AudioSessionStateActive:   return L"Active";
    case AudioSessionStateInactive: return L"Inactive";
    case AudioSessionStateExpired:  return L"Expired";
    default:                        return L"Unknown";
    }
}

// ===== Session Data =====
struct SessionEntry {
    std::wstring sid;   // SessionIdentifier（選択の引き継ぎキー）
    std::wstring name;  // 表示名（DisplayNameが空なら補正名）
    DWORD        pid;   // 参考
};

// ===== Globals =====
HINSTANCE               g_hInst = nullptr;
HWND                    g_hWnd = nullptr;
HWND                    g_comboA = nullptr, g_comboB = nullptr, g_track = nullptr;

IMMDeviceEnumerator* g_pEnumerator = nullptr;
IMMDevice* g_pDevice = nullptr;
IAudioSessionManager2* g_pSessionMgr2 = nullptr;

std::vector<SessionEntry> g_sessions;
std::set<std::wstring>     g_registeredSids;  // 現在イベント登録済みSID（列挙で置換）
std::set<std::wstring>     g_lastSids;        // 直近列挙のSID集合（差分検知用）
std::wstring               g_selectedSidA;     // 選択保持（SID）
std::wstring               g_selectedSidB;     // 選択保持（SID）

// 前方宣言
struct SessionWatcher;
static void RefreshSessionsAndUI(BOOL keepSelection);
static void ApplyBalanceFromTrackbar();


// ===== Watcher =====
struct SessionWatcher : IAudioSessionEvents, IAudioSessionNotification {
    LONG m_ref;
    HWND m_hNotify; // 通知先のウィンドウ
    IAudioSessionManager2* m_mgr;

    SessionWatcher(HWND hWnd, IAudioSessionManager2* mgr) : m_ref(1), m_hNotify(hWnd), m_mgr(mgr) {
        if (m_mgr) m_mgr->AddRef();
    }
    ~SessionWatcher() {
        if (m_mgr) m_mgr->Release();
    }

    // IUnknown
    HRESULT STDMETHODCALLTYPE QueryInterface(REFIID riid, void** ppv) override {
        if (!ppv) return E_POINTER;
        if (riid == __uuidof(IUnknown) ||
            riid == __uuidof(IAudioSessionEvents)) {
            *ppv = static_cast<IAudioSessionEvents*>(this);
        }
        else if (riid == __uuidof(IAudioSessionNotification)) {
            *ppv = static_cast<IAudioSessionNotification*>(this);
        }
        else {
            *ppv = nullptr;
            return E_NOINTERFACE;
        }
        AddRef();
        return S_OK;
    }
    ULONG STDMETHODCALLTYPE AddRef() override { return (ULONG)InterlockedIncrement(&m_ref); }
    ULONG STDMETHODCALLTYPE Release() override {
        ULONG r = (ULONG)InterlockedDecrement(&m_ref);
        if (r == 0) delete this;
        return r;
    }

    // 新規セッション
    HRESULT STDMETHODCALLTYPE OnSessionCreated(IAudioSessionControl* NewSession) override {
        if (NewSession) {
            // その場でイベント登録（取りこぼし対策）
            NewSession->RegisterAudioSessionNotification(this);
            IAudioSessionControl2* c2 = nullptr;
            if (SUCCEEDED(NewSession->QueryInterface(IID_PPV_ARGS(&c2))) && c2) {
                LPWSTR sid = nullptr;
                if (SUCCEEDED(c2->GetSessionIdentifier(&sid)) && sid) {
                    g_registeredSids.insert(sid);
                    CoTaskMemFree(sid);
                }
                c2->Release();
            }
        }
        PostMessage(m_hNotify, WMAPP_REFRESH, 0, 0);
        return S_OK;
    }

    // 状態変化は全てで更新（Active/Inactive/Expired）
    HRESULT STDMETHODCALLTYPE OnStateChanged(AudioSessionState /*NewState*/) override {
        PostMessage(m_hNotify, WMAPP_REFRESH, 0, 0);
        return S_OK;
    }

    // 追加の保険
    HRESULT STDMETHODCALLTYPE OnSessionDisconnected(AudioSessionDisconnectReason) override {
        PostMessage(m_hNotify, WMAPP_REFRESH, 0, 0);
        return S_OK;
    }

    // 未使用
    HRESULT STDMETHODCALLTYPE OnDisplayNameChanged(LPCWSTR, LPCGUID) override { return S_OK; }
    HRESULT STDMETHODCALLTYPE OnIconPathChanged(LPCWSTR, LPCGUID) override { return S_OK; }
    HRESULT STDMETHODCALLTYPE OnSimpleVolumeChanged(float, BOOL, LPCGUID) override { return S_OK; }
    HRESULT STDMETHODCALLTYPE OnChannelVolumeChanged(DWORD, float[], DWORD, LPCGUID) override { return S_OK; }
    HRESULT STDMETHODCALLTYPE OnGroupingParamChanged(LPCGUID, LPCGUID) override { return S_OK; }
};

SessionWatcher* g_pWatcher = nullptr;


// ===== Core: Enumerate & Register events. Return "changed?" =====
static bool BuildSessionListAndRegisterAndGetChanged() {
    g_sessions.clear();
    if (!g_pSessionMgr2) return false;

    IAudioSessionEnumerator* pEnum = nullptr;
    if (FAILED(g_pSessionMgr2->GetSessionEnumerator(&pEnum)) || !pEnum) return false;

    int count = 0;
    if (FAILED(pEnum->GetCount(&count))) { pEnum->Release(); return false; }

    std::set<std::wstring> currentSids;

    for (int i = 0; i < count; ++i) {
        IAudioSessionControl* pCtrl = nullptr;
        IAudioSessionControl2* pCtrl2 = nullptr;

        if (FAILED(pEnum->GetSession(i, &pCtrl)) || !pCtrl) continue;
        if (FAILED(pCtrl->QueryInterface(IID_PPV_ARGS(&pCtrl2))) || !pCtrl2) {
            pCtrl->Release(); continue;
        }

        // SID
        LPWSTR sid = nullptr;
        std::wstring key;
        if (SUCCEEDED(pCtrl2->GetSessionIdentifier(&sid)) && sid) {
            key = sid; CoTaskMemFree(sid);
        }
        if (!key.empty()) currentSids.insert(key);

        // セッションイベント未登録なら登録
        if (!key.empty() && g_pWatcher) {
            if (g_registeredSids.find(key) == g_registeredSids.end()) {
                pCtrl->RegisterAudioSessionNotification(g_pWatcher);
            }
        }

        // name
        LPWSTR displayName = nullptr;
        std::wstring name;
        if (SUCCEEDED(pCtrl2->GetDisplayName(&displayName)) && displayName) {
            name = displayName; CoTaskMemFree(displayName);
        }
        if (name.empty()) name = NormalizeNameFromSessionId(key);

        DWORD pid = 0; pCtrl2->GetProcessId(&pid);

        g_sessions.push_back(SessionEntry{ key, name, pid });

        pCtrl2->Release();
        pCtrl->Release();
    }
    pEnum->Release();

    std::sort(g_sessions.begin(), g_sessions.end(),
        [](const SessionEntry& a, const SessionEntry& b) { return a.name < b.name; });

    // 現在の集合で置換（再出現時に再登録できるようにする）
    g_registeredSids = currentSids;

    // 差分判定
    bool changed = (currentSids != g_lastSids);
    g_lastSids.swap(currentSids);
    return changed;
}

// ===== UI helpers =====
static void SetComboEditText(HWND hCombo, const std::wstring& text) {
    COMBOBOXINFO cbi = { sizeof(cbi) };
    if (GetComboBoxInfo(hCombo, &cbi) && cbi.hwndItem) {
        SetWindowTextW(cbi.hwndItem, text.c_str());
    }
    else {
        SetWindowTextW(hCombo, text.c_str());
    }
}
static int FindIndexBySid(const std::wstring& sid) {
    for (size_t i = 0; i < g_sessions.size(); ++i) {
        if (g_sessions[i].sid == sid) return (int)i;
    }
    return -1;
}

// ===== Repopulate combos (only when changed) =====
static void RepopulateCombos(BOOL keepSelection) {
    if (!g_comboA || !g_comboB) return;

    // 既存の編集欄表示を保険として保持
    wchar_t bufA[256] = { 0 }, bufB[256] = { 0 };
    GetWindowTextW(g_comboA, bufA, 256);
    GetWindowTextW(g_comboB, bufB, 256);

    SendMessage(g_comboA, CB_RESETCONTENT, 0, 0);
    SendMessage(g_comboB, CB_RESETCONTENT, 0, 0);

    for (size_t i = 0; i < g_sessions.size(); ++i) {
        const SessionEntry& s = g_sessions[i];
        SendMessage(g_comboA, CB_ADDSTRING, 0, (LPARAM)s.name.c_str());
        SendMessage(g_comboB, CB_ADDSTRING, 0, (LPARAM)s.name.c_str());
    }

    int selA = -1, selB = -1;
    if (keepSelection) {
        if (!g_selectedSidA.empty()) selA = FindIndexBySid(g_selectedSidA);
        if (!g_selectedSidB.empty()) selB = FindIndexBySid(g_selectedSidB);
    }

    if (selA >= 0) {
        SendMessage(g_comboA, CB_SETCURSEL, selA, 0);
    }
    else if (!g_selectedSidA.empty()) {
        std::wstring disp = L"[inactive] ";
        if (bufA[0]) disp += bufA; else disp += g_selectedSidA;
        SetComboEditText(g_comboA, disp);
    }

    if (selB >= 0) {
        SendMessage(g_comboB, CB_SETCURSEL, selB, 0);
    }
    else if (!g_selectedSidB.empty()) {
        std::wstring disp = L"[inactive] ";
        if (bufB[0]) disp += bufB; else disp += g_selectedSidB;
        SetComboEditText(g_comboB, disp);
    }

    InvalidateRect(g_comboA, nullptr, TRUE);
    InvalidateRect(g_comboB, nullptr, TRUE);
    UpdateWindow(g_comboA);
    UpdateWindow(g_comboB);
}

// ===== Set volume of a session by SID =====
static void SetSessionVolumeBySid(const std::wstring& sid, float volume01) {
    if (sid.empty() || !g_pSessionMgr2) return;

    IAudioSessionEnumerator* pEnum = nullptr;
    if (FAILED(g_pSessionMgr2->GetSessionEnumerator(&pEnum)) || !pEnum) return;

    int count = 0;
    if (FAILED(pEnum->GetCount(&count))) { pEnum->Release(); return; }

    for (int i = 0; i < count; ++i) {
        IAudioSessionControl* pCtrl = nullptr;
        IAudioSessionControl2* pCtrl2 = nullptr;
        ISimpleAudioVolume* pVol = nullptr;

        if (FAILED(pEnum->GetSession(i, &pCtrl)) || !pCtrl) continue;
        if (FAILED(pCtrl->QueryInterface(IID_PPV_ARGS(&pCtrl2))) || !pCtrl2) { pCtrl->Release(); continue; }

        LPWSTR wsid = nullptr;
        std::wstring key;
        if (SUCCEEDED(pCtrl2->GetSessionIdentifier(&wsid)) && wsid) {
            key = wsid; CoTaskMemFree(wsid);
        }

        if (key == sid) {
            if (SUCCEEDED(pCtrl->QueryInterface(IID_PPV_ARGS(&pVol))) && pVol) {
                if (volume01 < 0.0f) volume01 = 0.0f;
                else if (volume01 > 1.0f) volume01 = 1.0f;
                pVol->SetMasterVolume(volume01, nullptr);
                pVol->Release();
            }
            pCtrl2->Release();
            pCtrl->Release();
            break;
        }

        pCtrl2->Release();
        pCtrl->Release();
    }

    pEnum->Release();
}

// ===== Apply balance from trackbar =====
static void ApplyBalanceFromTrackbar() {
    if (!g_comboA || !g_comboB || !g_track) return;

    int selA = (int)SendMessage(g_comboA, CB_GETCURSEL, 0, 0);
    int selB = (int)SendMessage(g_comboB, CB_GETCURSEL, 0, 0);

    if (selA >= 0 && selA < (int)g_sessions.size()) g_selectedSidA = g_sessions[selA].sid;
    if (selB >= 0 && selB < (int)g_sessions.size()) g_selectedSidB = g_sessions[selB].sid;

    if (g_selectedSidA.empty() || g_selectedSidB.empty()) return;

    int pos = (int)SendMessage(g_track, TBM_GETPOS, 0, 0);
    if (pos < 0) pos = 0; else if (pos > 100) pos = 100;

    float a = (100 - pos) / 100.0f;
    float b = pos / 100.0f;

    SetSessionVolumeBySid(g_selectedSidA, a);
    SetSessionVolumeBySid(g_selectedSidB, b);
}

// ===== Refresh (enumerate + repopulate if changed) =====
static void RefreshSessionsAndUI(BOOL keepSelection) {
    bool changed = BuildSessionListAndRegisterAndGetChanged();
    if (changed) {
        RepopulateCombos(keepSelection);
    }
}

// ===== Init / Uninit WASAPI =====
static bool InitWasapi() {
    if (FAILED(CoCreateInstance(__uuidof(MMDeviceEnumerator), nullptr, CLSCTX_ALL, IID_PPV_ARGS(&g_pEnumerator))))
        return false;
    if (FAILED(g_pEnumerator->GetDefaultAudioEndpoint(eRender, eMultimedia, &g_pDevice)))
        return false;
    if (FAILED(g_pDevice->Activate(__uuidof(IAudioSessionManager2), CLSCTX_ALL, nullptr, (void**)&g_pSessionMgr2)))
        return false;

    // Watcher 起動：先に Manager へ登録 → 初回列挙
    g_pWatcher = new SessionWatcher(g_hWnd, g_pSessionMgr2);
    if (g_pWatcher) {
        g_pSessionMgr2->RegisterSessionNotification(g_pWatcher);
        RefreshSessionsAndUI(TRUE);
    }
    return true;
}

static void UninitWasapi() {
    if (g_pSessionMgr2 && g_pWatcher) {
        g_pSessionMgr2->UnregisterSessionNotification(g_pWatcher);
        g_pWatcher->Release();
        g_pWatcher = nullptr;
    }
    if (g_pSessionMgr2) { g_pSessionMgr2->Release(); g_pSessionMgr2 = nullptr; }
    if (g_pDevice) { g_pDevice->Release();      g_pDevice = nullptr; }
    if (g_pEnumerator) { g_pEnumerator->Release();  g_pEnumerator = nullptr; }
}

// ===== Window / Layout =====
static void DoLayout(HWND hWnd) {
    RECT rc; GetClientRect(hWnd, &rc);
    int w = rc.right - rc.left;

    const int margin = 8;
    const int comboWidth = (w - (margin * 3)) / 2;
    const int comboHeight = 120; // CBS_SIMPLE は常時リスト表示
    const int trackHeight = 40;

    MoveWindow(g_comboA, margin, margin, comboWidth, comboHeight, TRUE);
    MoveWindow(g_comboB, margin * 2 + comboWidth, margin, comboWidth, comboHeight, TRUE);

    int trackY = margin + comboHeight + margin;
    MoveWindow(g_track, margin, trackY, w - margin * 2, trackHeight, TRUE);
}

static LRESULT CALLBACK WndProc(HWND hWnd, UINT msg, WPARAM wParam, LPARAM lParam) {
    switch (msg) {
    case WM_CREATE: {
        INITCOMMONCONTROLSEX icc = { sizeof(icc), ICC_BAR_CLASSES };
        InitCommonControlsEx(&icc);

        g_comboA = CreateWindowExW(0, WC_COMBOBOXW, L"",
            WS_CHILD | WS_VISIBLE | CBS_SIMPLE | WS_VSCROLL,
            0, 0, 0, 0, hWnd, (HMENU)IDC_COMBO_A, g_hInst, nullptr);

        g_comboB = CreateWindowExW(0, WC_COMBOBOXW, L"",
            WS_CHILD | WS_VISIBLE | CBS_SIMPLE | WS_VSCROLL,
            0, 0, 0, 0, hWnd, (HMENU)IDC_COMBO_B, g_hInst, nullptr);

        g_track = CreateWindowExW(0, TRACKBAR_CLASSW, L"",
            WS_CHILD | WS_VISIBLE | TBS_AUTOTICKS,
            0, 0, 0, 0, hWnd, (HMENU)IDC_TRACK, g_hInst, nullptr);

        SendMessage(g_track, TBM_SETRANGE, TRUE, MAKELONG(0, 100));
        SendMessage(g_track, TBM_SETPOS, TRUE, 50); // 中央（50/50）

        DoLayout(hWnd);

        if (!InitWasapi()) {
            MessageBoxW(hWnd, L"WASAPI 初期化に失敗しました。", L"Error", MB_ICONERROR);
            PostQuitMessage(1);
        }

        // 1秒ポーリング（イベント取りこぼし対策）
        SetTimer(hWnd, 1, 1000, nullptr);
        return 0;
    }
    case WM_SIZE:
        DoLayout(hWnd);
        return 0;

    case WM_COMMAND: {
        const WORD id = LOWORD(wParam);
        const WORD code = HIWORD(wParam);
        if ((id == IDC_COMBO_A || id == IDC_COMBO_B) && (code == CBN_SELCHANGE)) {
            ApplyBalanceFromTrackbar();
        }
        return 0;
    }

    case WM_HSCROLL:
        if ((HWND)lParam == g_track) {
            ApplyBalanceFromTrackbar();
        }
        return 0;

    case WM_TIMER:
        if (wParam == 1) {
            RefreshSessionsAndUI(TRUE); // 変更時だけUI更新（選択維持）
        }
        return 0;

    case WMAPP_REFRESH:
        RefreshSessionsAndUI(TRUE);
        return 0;

    case WM_DESTROY:
        KillTimer(hWnd, 1);
        UninitWasapi();
        PostQuitMessage(0);
        return 0;
    }
    return DefWindowProcW(hWnd, msg, wParam, lParam);
}

int APIENTRY wWinMain(HINSTANCE hInstance, HINSTANCE, LPWSTR, int nCmdShow) {
    g_hInst = hInstance;
    HRESULT hr = CoInitializeEx(nullptr, COINIT_MULTITHREADED);
    if (FAILED(hr)) return 1;

    const wchar_t CLASS_NAME[] = L"AudioBalanceMixerWnd";
    WNDCLASSEXW wc = { sizeof(wc) };
    wc.lpfnWndProc = WndProc;
    wc.hInstance = hInstance;
    wc.lpszClassName = CLASS_NAME;
    wc.hCursor = LoadCursor(nullptr, IDC_ARROW);
    wc.hbrBackground = (HBRUSH)(COLOR_WINDOW + 1);

    if (!RegisterClassExW(&wc)) { CoUninitialize(); return 1; }

    g_hWnd = CreateWindowExW(0, CLASS_NAME, L"Audio Balance Mixer (2-App)",
        WS_OVERLAPPEDWINDOW, CW_USEDEFAULT, CW_USEDEFAULT, 720, 420,
        nullptr, nullptr, hInstance, nullptr);
    if (!g_hWnd) { CoUninitialize(); return 1; }

    ShowWindow(g_hWnd, nCmdShow);
    UpdateWindow(g_hWnd);

    MSG msg;
    while (GetMessageW(&msg, nullptr, 0, 0)) {
        TranslateMessage(&msg);
        DispatchMessageW(&msg);
    }

    CoUninitialize();
    return (int)msg.wParam;
}

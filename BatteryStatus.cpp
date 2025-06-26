#define _WIN32_WINNT 0x0601
#include <windows.h>
#include <powrprof.h>
#include <shellapi.h>
#include <strsafe.h>
#include <tchar.h>
#include <stdio.h>
#include <math.h>
#include <time.h>
#include <commctrl.h>
#include <algorithm>
#include <comdef.h>
#include <Wbemidl.h>
#pragma comment(lib, "user32.lib")
#pragma comment(lib, "shell32.lib")
#pragma comment(lib, "PowrProf.lib")
#pragma comment(lib, "advapi32.lib")
#pragma comment(lib, "comctl32.lib")
#pragma comment(lib, "wbemuuid.lib")

#define WM_TRAYICON      (WM_USER + 1)
#define ID_TRAYICON      1001
#define IDT_TIMER        2001
#define IDT_TOOLBAR      2002
#define IDM_EXIT         40001
#define IDM_SHOW_TOOLBAR 40002
#define IDM_AUTOSTART    40003

HINSTANCE hInst;
NOTIFYICONDATA nid = { 0 };
HWND hMainWnd = nullptr;
HWND hToolbarWnd = nullptr;
HWND hTooltip = nullptr;
TOOLINFO ti = { 0 };
bool toolbarVisible = false;
bool dragging = false;
POINT dragOffset = { 0, 0 };
TCHAR iniPath[MAX_PATH] = { 0 };
TCHAR dbPath[MAX_PATH] = { 0 };

#define MAX_SAMPLES 40

struct BatterySample {
    int percent;
    int ac;
    int rate;
    int flag;
    int milliWatts;
    time_t t;
};

struct BatteryDB {
    int idxDischarge;
    BatterySample discharge[MAX_SAMPLES];
    int idxCharge;
    BatterySample charge[MAX_SAMPLES];
};

void GetDbPath() {
    if (!dbPath[0]) {
        GetModuleFileName(NULL, dbPath, MAX_PATH);
        TCHAR* p = _tcsrchr(dbPath, _T('\\'));
        if (p) *(p + 1) = 0;
        _tcscat_s(dbPath, MAX_PATH, _T("History.bin"));
    }
}

bool IsBatterySampleValid(int percent, int ac, int rate, int milliWatts, int flag) {
    if (percent < 1 || percent > 100 || percent == 255)
        return false;
    if (ac != 0 && ac != 1)
        return false;
    if (time(NULL) == 0)
        return false;
    return true;
}

void LogBatterySample(int percent, int ac, int rate, int milliWatts, int systemFlag) {
    if (!IsBatterySampleValid(percent, ac, rate, milliWatts, systemFlag))
        return;
    GetDbPath();
    FILE* f;
    BatteryDB db = { 0 };
    bool exists = false;
    _tfopen_s(&f, dbPath, _T("r+b"));
    if (f) {
        fread(&db, sizeof(BatteryDB), 1, f);
        fclose(f);
        exists = true;
    }
    else {
        db.idxDischarge = db.idxCharge = 0;
        memset(db.discharge, 0, sizeof(db.discharge));
        memset(db.charge, 0, sizeof(db.charge));
    }
    BatterySample s;
    s.percent = percent;
    s.ac = ac;
    s.rate = rate;
    s.flag = systemFlag;
    s.milliWatts = milliWatts;
    s.t = time(NULL);

    if (ac == 0) {
        db.discharge[db.idxDischarge] = s;
        db.idxDischarge = (db.idxDischarge + 1) % MAX_SAMPLES;
    }
    else {
        db.charge[db.idxCharge] = s;
        db.idxCharge = (db.idxCharge + 1) % MAX_SAMPLES;
    }
    _tfopen_s(&f, dbPath, _T("w+b"));
    if (f) {
        fwrite(&db, sizeof(BatteryDB), 1, f);
        fclose(f);
    }
}

int ReadBatteryHistory(int ac, BatterySample* outSamples, int maxSamples) {
    GetDbPath();
    FILE* f;
    _tfopen_s(&f, dbPath, _T("rb"));
    if (!f) return 0;
    BatteryDB db;
    size_t n = fread(&db, sizeof(BatteryDB), 1, f);
    fclose(f);
    if (n != 1) return 0;

    int found = 0;
    if (ac == 0) {
        for (int i = 0; i < MAX_SAMPLES; ++i) {
            int pos = (db.idxDischarge + i) % MAX_SAMPLES;
            if (db.discharge[pos].t == 0) continue;
            if (found < maxSamples)
                outSamples[found++] = db.discharge[pos];
        }
    }
    else {
        for (int i = 0; i < MAX_SAMPLES; ++i) {
            int pos = (db.idxCharge + i) % MAX_SAMPLES;
            if (db.charge[pos].t == 0) continue;
            if (found < maxSamples)
                outSamples[found++] = db.charge[pos];
        }
    }
    std::sort(outSamples, outSamples + found, [](const BatterySample& a, const BatterySample& b) {
        return a.t < b.t;
        });
    return found;
}

int EstimateTimeFromHistory(int ac, int currentPercent, int* outRatePerHour, int* outSampleCount) {
    const int maxSamples = MAX_SAMPLES;
    BatterySample samples[maxSamples];
    int n = ReadBatteryHistory(ac, samples, maxSamples);
    if (outSampleCount) *outSampleCount = n;
    if (n < 2) return -1;

    double totalPercent = 0.0;
    double totalTime = 0.0;
    int validIntervals = 0;
    for (int i = 1; i < n; ++i) {
        double dPercent = samples[i].percent - samples[i - 1].percent;
        double dTime = difftime(samples[i].t, samples[i - 1].t) / 3600.0;
        if (fabs(dTime) < 0.017 || dPercent == 0) continue;
        totalPercent += dPercent;
        totalTime += dTime;
        ++validIntervals;
    }
    if (validIntervals == 0) {
        double dPercent = samples[n - 1].percent - samples[0].percent;
        double dTime = difftime(samples[n - 1].t, samples[0].t) / 3600.0;
        if (fabs(dTime) < 0.017 || dPercent == 0) return -1;
        totalPercent = dPercent;
        totalTime = dTime;
    }
    if (fabs(totalTime) < 0.017 || totalPercent == 0) return -1;

    int ratePerHour = (int)(totalPercent / totalTime);
    *outRatePerHour = ratePerHour;

    int minutes = 0;
    if (ac) {
        if (ratePerHour > 0)
            minutes = (int)((100 - currentPercent) * 60.0 / ratePerHour + 0.5);
        else
            minutes = 24 * 60;
    }
    else {
        if (ratePerHour < 0)
            minutes = (int)((currentPercent * 60.0) / -ratePerHour + 0.5);
        else
            minutes = 24 * 60;
    }
    return (minutes > 0) ? minutes * 60 : -1;
}

void GetIniPath() {
    if (!iniPath[0]) {
        GetModuleFileName(NULL, iniPath, MAX_PATH);
        TCHAR* p = _tcsrchr(iniPath, _T('\\'));
        if (p) *(p + 1) = 0;
        _tcscat_s(iniPath, MAX_PATH, _T("BatteryStatus.ini"));
    }
}

void SaveToolbarPosition() {
    if (!hToolbarWnd) return;
    RECT rc;
    GetWindowRect(hToolbarWnd, &rc);
    GetIniPath();
    TCHAR posbuf[64];
    StringCchPrintf(posbuf, 64, _T("%d,%d"), rc.left, rc.top);
    WritePrivateProfileString(_T("Toolbar"), _T("Position"), posbuf, iniPath);
}

void LoadToolbarPosition(int* px, int* py) {
    GetIniPath();
    TCHAR posbuf[64];
    GetPrivateProfileString(_T("Toolbar"), _T("Position"), _T("100,100"), posbuf, 64, iniPath);
    _stscanf_s(posbuf, _T("%d,%d"), px, py);
}

void SaveToolbarVisible(bool visible) {
    GetIniPath();
    WritePrivateProfileString(_T("Toolbar"), _T("Visible"), visible ? _T("1") : _T("0"), iniPath);
}

bool LoadToolbarVisible() {
    GetIniPath();
    int val = GetPrivateProfileInt(_T("Toolbar"), _T("Visible"), 1, iniPath);
    return val != 0;
}

bool IsAutoStartEnabled() {
    HKEY hKey;
    TCHAR exePath[MAX_PATH];
    GetModuleFileName(NULL, exePath, MAX_PATH);
    if (RegOpenKeyEx(HKEY_CURRENT_USER, _T("Software\\Microsoft\\Windows\\CurrentVersion\\Run"), 0, KEY_READ, &hKey) != ERROR_SUCCESS)
        return false;
    TCHAR value[MAX_PATH] = { 0 };
    DWORD len = sizeof(value);
    LONG result = RegQueryValueEx(hKey, _T("BatteryStatus"), NULL, NULL, (LPBYTE)value, &len);
    RegCloseKey(hKey);
    if (result != ERROR_SUCCESS)
        return false;
    return _tcsicmp(value, exePath) == 0;
}

void SetAutoStart(bool enable) {
    HKEY hKey;
    TCHAR exePath[MAX_PATH];
    GetModuleFileName(NULL, exePath, MAX_PATH);
    if (RegOpenKeyEx(HKEY_CURRENT_USER, _T("Software\\Microsoft\\Windows\\CurrentVersion\\Run"), 0, KEY_WRITE, &hKey) != ERROR_SUCCESS)
        return;
    if (enable) {
        RegSetValueEx(hKey, _T("BatteryStatus"), 0, REG_SZ, (const BYTE*)exePath, (_tcslen(exePath) + 1) * sizeof(TCHAR));
    }
    else {
        RegDeleteValue(hKey, _T("BatteryStatus"));
    }
    RegCloseKey(hKey);
}

bool GetBatterySmartStatus(
    int& percent,
    int& timeSec,
    bool& charging,
    double& watts,
    bool& haveWatt,
    bool& haveSmartTime,
    int& acLineStatus,
    int& batteryFlag,
    int& milliWatts
) {
    SYSTEM_POWER_STATUS sps;
    percent = 100; timeSec = 0; charging = false; watts = 0; haveWatt = false; haveSmartTime = false; acLineStatus = 0; batteryFlag = 0; milliWatts = 0;

    SYSTEM_BATTERY_STATE sbs;
    if (CallNtPowerInformation(SystemBatteryState, NULL, 0, &sbs, sizeof sbs) == 0) {
        if (sbs.MaxCapacity && sbs.RemainingCapacity && sbs.MaxCapacity != 0xFFFFFFFF && sbs.RemainingCapacity != 0xFFFFFFFF)
            percent = (int)(100.0 * sbs.RemainingCapacity / sbs.MaxCapacity + 0.5);
        charging = (!!sbs.Charging);
        LONG rate = (sbs.Rate != 0x80000000 ? (LONG)sbs.Rate : 0);
        if (rate != 0) {
            haveWatt = true;
            watts = (double)abs(rate) / 1000.0;
            milliWatts = abs(rate);
        }
        if (rate != 0 && sbs.MaxCapacity && sbs.RemainingCapacity) {
            if (charging && rate > 0) {
                int diff = (int)(sbs.MaxCapacity - sbs.RemainingCapacity);
                timeSec = (int)((diff * 3600.0) / rate + 0.5);
                haveSmartTime = true;
            }
            else if (!charging && rate < 0) {
                timeSec = (int)((sbs.RemainingCapacity * 3600.0) / (-rate) + 0.5);
                haveSmartTime = true;
            }
        }
    }
    if (!GetSystemPowerStatus(&sps))
        return haveSmartTime;
    if (percent == 100) percent = sps.BatteryLifePercent;
    acLineStatus = sps.ACLineStatus;
    batteryFlag = sps.BatteryFlag;
    if (!haveSmartTime) {
        timeSec = (charging ? sps.BatteryFullLifeTime : sps.BatteryLifeTime);
    }
    return haveSmartTime;
}

void DrawBatteryBox(HDC hdc, int x, int y, int w, int h, int percent, bool charging, bool low) {
    RECT rcBody = { x, y + 3, x + w - 1, y + h - 1 };
    HBRUSH hBody = CreateSolidBrush(RGB(255, 255, 255));
    FillRect(hdc, &rcBody, hBody);
    DeleteObject(hBody);

    HPEN hPen = CreatePen(PS_SOLID, 1, RGB(0, 0, 0));
    HPEN hOldPen = (HPEN)SelectObject(hdc, hPen);
    HBRUSH hOldBrush = (HBRUSH)SelectObject(hdc, GetStockObject(NULL_BRUSH));
    Rectangle(hdc, rcBody.left, rcBody.top, rcBody.right, rcBody.bottom);
    SelectObject(hdc, hOldBrush);
    SelectObject(hdc, hOldPen);
    DeleteObject(hPen);

    int tipW = w / 2;
    int tipX = x + (w - tipW) / 2;
    RECT rcTip = { tipX, y, tipX + tipW, y + 4 };
    HBRUSH hTip = CreateSolidBrush(RGB(0, 0, 0));
    FillRect(hdc, &rcTip, hTip);
    DeleteObject(hTip);

    int fillMargin = 2;
    int fillableHeight = rcBody.bottom - rcBody.top - 2 * fillMargin;
    int fillHeight = (fillableHeight * percent) / 100;
    int fillTop = rcBody.bottom - fillMargin - fillHeight;
    RECT rcFill = { rcBody.left + fillMargin, fillTop, rcBody.right - fillMargin, rcBody.bottom - fillMargin };
    HBRUSH hFill;
    if (low)
        hFill = CreateSolidBrush(RGB(255, 64, 64));
    else if (charging)
        hFill = CreateSolidBrush(RGB(128, 255, 128));
    else
        hFill = CreateSolidBrush(RGB(255, 255, 128));
    FillRect(hdc, &rcFill, hFill);
    DeleteObject(hFill);

    if (charging) {
        HPEN hGreenPen = CreatePen(PS_SOLID, 1, RGB(0, 192, 0));
        HPEN hOldGreen = (HPEN)SelectObject(hdc, hGreenPen);
        int boltX = x + w / 2;
        int boltY = y + h / 2 - 6;
        POINT pts[] = {
            { boltX - 3, boltY + 2 },
            { boltX + 0, boltY + 6 },
            { boltX - 2, boltY + 6 },
            { boltX + 3, boltY + 13 },
            { boltX + 1, boltY + 7 },
            { boltX + 4, boltY + 7 },
            { boltX + 1, boltY + 2 }
        };
        Polyline(hdc, pts, sizeof(pts) / sizeof(pts[0]));
        SelectObject(hdc, hOldGreen);
        DeleteObject(hGreenPen);
    }
}

void FormatTime(int seconds, bool charging, TCHAR* buf, size_t len) {
    if (seconds < 0 || seconds > 24 * 3600) {
        StringCchCopy(buf, len, charging ? _T("-?:??") : _T("?:??"));
    }
    else {
        int h = seconds / 3600;
        int m = (seconds % 3600) / 60;
        StringCchPrintf(buf, len, charging ? _T("-%d:%02d") : _T("%d:%02d"), h, m);
    }
}

// --- WMI Battery Wear Helper ---
bool GetBatteryCapacities(ULONG* design, ULONG* full) {
    *design = 0; *full = 0;
    HRESULT hres = CoInitializeEx(0, COINIT_MULTITHREADED);
    if (FAILED(hres) && hres != RPC_E_CHANGED_MODE) return false;

    hres = CoInitializeSecurity(
        NULL, -1, NULL, NULL,
        RPC_C_AUTHN_LEVEL_DEFAULT, RPC_C_IMP_LEVEL_IMPERSONATE,
        NULL, EOAC_NONE, NULL);
    if (FAILED(hres) && hres != RPC_E_TOO_LATE) {
        CoUninitialize();
        return false;
    }

    IWbemLocator* pLoc = nullptr;
    hres = CoCreateInstance(
        CLSID_WbemLocator, 0, CLSCTX_INPROC_SERVER,
        IID_IWbemLocator, (LPVOID*)&pLoc);
    if (FAILED(hres)) {
        CoUninitialize();
        return false;
    }

    IWbemServices* pSvc = nullptr;
    hres = pLoc->ConnectServer(
        _bstr_t(L"ROOT\\CIMV2"),
        NULL, NULL, 0, NULL, 0, 0, &pSvc);
    if (FAILED(hres)) {
        pLoc->Release();
        CoUninitialize();
        return false;
    }

    hres = CoSetProxyBlanket(
        pSvc, RPC_C_AUTHN_WINNT, RPC_C_AUTHZ_NONE, NULL,
        RPC_C_AUTHN_LEVEL_CALL, RPC_C_IMP_LEVEL_IMPERSONATE,
        NULL, EOAC_NONE);
    if (FAILED(hres)) {
        pSvc->Release();
        pLoc->Release();
        CoUninitialize();
        return false;
    }

    IEnumWbemClassObject* pEnumerator = nullptr;
    hres = pSvc->ExecQuery(
        bstr_t("WQL"),
        bstr_t("SELECT DesignCapacity, FullChargeCapacity FROM Win32_Battery"),
        WBEM_FLAG_FORWARD_ONLY | WBEM_FLAG_RETURN_IMMEDIATELY,
        NULL, &pEnumerator);
    if (FAILED(hres)) {
        pSvc->Release();
        pLoc->Release();
        CoUninitialize();
        return false;
    }

    IWbemClassObject* pclsObj = nullptr;
    ULONG uReturn = 0;
    bool gotData = false;
    if (pEnumerator) {
        HRESULT hr = pEnumerator->Next(WBEM_INFINITE, 1, &pclsObj, &uReturn);
        if (uReturn) {
            VARIANT vtProp;
            if (SUCCEEDED(pclsObj->Get(L"DesignCapacity", 0, &vtProp, 0, 0)) && (vtProp.vt == VT_I4 || vtProp.vt == VT_UI4)) {
                *design = vtProp.uintVal;
            }
            VariantClear(&vtProp);
            if (SUCCEEDED(pclsObj->Get(L"FullChargeCapacity", 0, &vtProp, 0, 0)) && (vtProp.vt == VT_I4 || vtProp.vt == VT_UI4)) {
                *full = vtProp.uintVal;
            }
            VariantClear(&vtProp);
            gotData = (*design > 0 && *full > 0);
            pclsObj->Release();
        }
        pEnumerator->Release();
    }
    pSvc->Release();
    pLoc->Release();
    CoUninitialize();
    return gotData;
}

void ShowToolbarTooltip(HWND hwnd) {
    if (!hTooltip) {
        hTooltip = CreateWindowEx(WS_EX_TOPMOST, TOOLTIPS_CLASS, NULL,
            WS_POPUP | TTS_ALWAYSTIP | TTS_NOPREFIX | TTS_BALLOON,
            CW_USEDEFAULT, CW_USEDEFAULT, CW_USEDEFAULT, CW_USEDEFAULT,
            hwnd, NULL, hInst, NULL);

        SetWindowPos(hTooltip, HWND_TOPMOST, 0, 0, 0, 0,
            SWP_NOMOVE | SWP_NOSIZE | SWP_NOACTIVATE);

        ti.cbSize = sizeof(TOOLINFO);
        ti.uFlags = TTF_IDISHWND | TTF_SUBCLASS;
        ti.hwnd = hwnd;
        ti.uId = (UINT_PTR)hwnd;
        ti.lpszText = LPSTR_TEXTCALLBACK;
        ti.rect.left = ti.rect.top = ti.rect.right = ti.rect.bottom = 0;
        SendMessage(hTooltip, TTM_ADDTOOL, 0, (LPARAM)&ti);
    }

    int percent = 0, timeSec = 0, acLineStatus = 0, batteryFlag = 0, milliWatts = 0;
    bool charging = false, haveWatt = false, haveSmartTime = false;
    double watts = 0;
    GetBatterySmartStatus(percent, timeSec, charging, watts, haveWatt, haveSmartTime, acLineStatus, batteryFlag, milliWatts);

    int ratePerHour = 0;
    int sampleCount = 0;
    int histTime = EstimateTimeFromHistory(acLineStatus, percent, &ratePerHour, &sampleCount);

    TCHAR dbTime[32], sysTime[32];
    FormatTime(timeSec, charging, sysTime, _countof(sysTime));
    if (histTime > 0) {
        FormatTime(histTime, charging, dbTime, _countof(dbTime));
    }
    else if (sampleCount < 2) {
        _tcscpy_s(dbTime, _T("Collecting data…"));
    }
    else {
        _tcscpy_s(dbTime, _T("N/A"));
    }

    static TCHAR tipText[256];
    StringCchPrintf(tipText, _countof(tipText),
        _T("Estimate (History): %s\r\nEstimate (Windows): %s\r\nSamples: %d"),
        dbTime, sysTime, sampleCount);

    ti.lpszText = tipText;
    SendMessage(hTooltip, TTM_SETMAXTIPWIDTH, 0, 400);
    SendMessage(hTooltip, TTM_UPDATETIPTEXT, 0, (LPARAM)&ti);

    RECT rc;
    GetWindowRect(hwnd, &rc);
    int screenW = GetSystemMetrics(SM_CXSCREEN);
    int screenH = GetSystemMetrics(SM_CYSCREEN);

    int tipW = 220, tipH = 40;
    int tipX = rc.right;
    int tipY = rc.top;
    if (tipX + tipW > screenW) tipX = screenW - tipW;
    if (tipY + tipH > screenH) tipY = screenH - tipH;
    if (tipX < 0) tipX = 0;
    if (tipY < 0) tipY = 0;

    SendMessage(hTooltip, TTM_TRACKPOSITION, 0, (LPARAM)MAKELONG(tipX, tipY));
    SendMessage(hTooltip, TTM_TRACKACTIVATE, TRUE, (LPARAM)&ti);
}

void HideToolbarTooltip() {
    if (hTooltip) {
        SendMessage(hTooltip, TTM_TRACKACTIVATE, FALSE, (LPARAM)&ti);
    }
}

LRESULT CALLBACK ToolbarProc(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam) {
    static bool tooltipShown = false;
    switch (msg) {
    case WM_CREATE:
        SetTimer(hwnd, IDT_TOOLBAR, 3000, NULL);
        break;
    case WM_ERASEBKGND:
        return 1;
    case WM_TIMER:
        InvalidateRect(hwnd, NULL, FALSE);
        break;
    case WM_LBUTTONDOWN: {
        POINT pt;
        GetCursorPos(&pt);
        RECT rc;
        GetWindowRect(hwnd, &rc);
        dragOffset.x = pt.x - rc.left;
        dragOffset.y = pt.y - rc.top;
        SetCapture(hwnd);
        dragging = true;
        break;
    }
    case WM_MOUSEMOVE: {
        if (dragging && (wParam & MK_LBUTTON)) {
            POINT pt;
            GetCursorPos(&pt);
            int x = pt.x - dragOffset.x;
            int y = pt.y - dragOffset.y;
            SetWindowPos(hwnd, HWND_TOPMOST, x, y, 0, 0, SWP_NOZORDER | SWP_NOSIZE | SWP_NOACTIVATE);
        }
        else {
            TRACKMOUSEEVENT tme = { sizeof(TRACKMOUSEEVENT) };
            tme.dwFlags = TME_HOVER | TME_LEAVE;
            tme.hwndTrack = hwnd;
            tme.dwHoverTime = 500;
            TrackMouseEvent(&tme);
        }
        break;
    }
    case WM_MOUSEHOVER:
        ShowToolbarTooltip(hwnd);
        tooltipShown = true;
        break;
    case WM_MOUSELEAVE:
        HideToolbarTooltip();
        tooltipShown = false;
        break;
    case WM_LBUTTONUP:
        if (dragging) {
            dragging = false;
            ReleaseCapture();
            SaveToolbarPosition();
        }
        break;
    case WM_PAINT: {
        PAINTSTRUCT ps;
        HDC hdc = BeginPaint(hwnd, &ps);
        RECT client;
        GetClientRect(hwnd, &client);
        int width = client.right - client.left;
        int height = client.bottom - client.top;

        HDC memDC = CreateCompatibleDC(hdc);
        HBITMAP memBM = CreateCompatibleBitmap(hdc, width, height);
        HBITMAP oldBM = (HBITMAP)SelectObject(memDC, memBM);

        HBRUSH bgBrush = CreateSolidBrush(GetSysColor(COLOR_3DFACE));
        FillRect(memDC, &client, bgBrush);
        DeleteObject(bgBrush);

        int toolbarW = width;
        int toolbarH = height;

        int percent = 0, timeSec = 0, acLineStatus = 0, batteryFlag = 0, milliWatts = 0;
        bool charging = false, haveWatt = false, haveSmartTime = false;
        double watts = 0;
        GetBatterySmartStatus(percent, timeSec, charging, watts, haveWatt, haveSmartTime, acLineStatus, batteryFlag, milliWatts);

        LogBatterySample(percent, acLineStatus, haveWatt ? (int)watts : 0, milliWatts, batteryFlag);

        bool low = percent <= 10;
        int boxX = 1, boxY = 2, boxW = 16, boxH = 16;
        DrawBatteryBox(memDC, boxX, boxY, boxW, boxH, percent, charging, low);

        int iconRight = boxX + boxW;
        int marginLeft = 3, marginRight = 2;
        int barX = iconRight + marginLeft;
        int barY = 3;
        int barWidth = toolbarW - barX - marginRight;
        int squares = 20;
        int sqGap = 1;
        int sqSize = (barWidth - (squares - 1) * sqGap) / squares;
        int barActualWidth = squares * sqSize + (squares - 1) * sqGap;
        int barStartX = barX + (barWidth - barActualWidth) / 2;
        int filled = percent * squares / 100;
        HBRUSH hBlue = CreateSolidBrush(RGB(0, 120, 215));
        for (int i = 0; i < filled; ++i) {
            RECT sq = {
                barStartX + i * (sqSize + sqGap),
                barY,
                barStartX + i * (sqSize + sqGap) + sqSize,
                barY + sqSize
            };
            FillRect(memDC, &sq, hBlue);
            FrameRect(memDC, &sq, (HBRUSH)GetStockObject(BLACK_BRUSH));
        }
        DeleteObject(hBlue);

        HFONT hFont = (HFONT)GetStockObject(DEFAULT_GUI_FONT);
        HFONT hOldFont = (HFONT)SelectObject(memDC, hFont);
        SetBkMode(memDC, TRANSPARENT);
        SetTextColor(memDC, RGB(0, 0, 0));
        SIZE textSize;
        TCHAR textbuf[128];
        TCHAR timebuf[16], percentbuf[16], wattbuf[16];

        bool noBattery = (batteryFlag & 128) != 0;
        bool acAndFull = (acLineStatus == 1 && percent == 100);
        bool batteryUnknown = (percent == 255);
        if (noBattery || acAndFull) {
            StringCchCopy(textbuf, _countof(textbuf), _T("A/C"));
        }
        else if (batteryUnknown) {
            StringCchCopy(textbuf, _countof(textbuf), _T("N/A"));
        }
        else {
            int ratePerHour = 0, sampleCount = 0;
            int histTime = EstimateTimeFromHistory(acLineStatus, percent, &ratePerHour, &sampleCount);
            int displayTime = histTime > 0 ? histTime : timeSec;
            FormatTime(displayTime, charging, timebuf, _countof(timebuf));
            StringCchPrintf(percentbuf, _countof(percentbuf), _T("%d%%"), percent);
            if (haveWatt)
                StringCchPrintf(wattbuf, _countof(wattbuf), _T("%.1fW"), watts);
            else
                StringCchCopy(wattbuf, _countof(wattbuf), _T("--.-W"));
            StringCchPrintf(textbuf, _countof(textbuf), _T("%s %s %s"), timebuf, percentbuf, wattbuf);
        }
        GetTextExtentPoint32(memDC, textbuf, lstrlen(textbuf), &textSize);
        int textX = barStartX + (barActualWidth - textSize.cx) / 2;
        int textY = barY + sqSize + 1;
        if (textY + textSize.cy > toolbarH) textY = toolbarH - textSize.cy;
        TextOut(memDC, textX, textY, textbuf, lstrlen(textbuf));
        SelectObject(memDC, hOldFont);

        BitBlt(hdc, 0, 0, width, height, memDC, 0, 0, SRCCOPY);

        SelectObject(memDC, oldBM);
        DeleteObject(memBM);
        DeleteDC(memDC);

        EndPaint(hwnd, &ps);
        break;
    }
    case WM_SETCURSOR:
        SetCursor(LoadCursor(NULL, IDC_ARROW));
        return TRUE;
    case WM_ACTIVATE:
        SetWindowPos(hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE | SWP_NOSIZE);
        break;
    case WM_CLOSE:
        SaveToolbarPosition();
        SaveToolbarVisible(false);
        ShowWindow(hwnd, SW_HIDE);
        toolbarVisible = false;
        break;
    case WM_DESTROY:
        KillTimer(hwnd, IDT_TOOLBAR);
        break;
    default:
        return DefWindowProc(hwnd, msg, wParam, lParam);
    }
    return 0;
}

void ShowToolbar(HWND parent) {
    int posX = 100, posY = 100;
    LoadToolbarPosition(&posX, &posY);
    if (!hToolbarWnd) {
        WNDCLASS wc = { 0 };
        wc.lpfnWndProc = ToolbarProc;
        wc.hInstance = hInst;
        wc.lpszClassName = _T("BattStatToolbarClass");
        wc.hbrBackground = (HBRUSH)(COLOR_3DFACE + 1);
        wc.hCursor = LoadCursor(NULL, IDC_ARROW);
        RegisterClass(&wc);

        hToolbarWnd = CreateWindowEx(
            WS_EX_TOPMOST | WS_EX_TOOLWINDOW,
            wc.lpszClassName, _T("Battery Toolbar"),
            WS_POPUP | WS_VISIBLE,
            posX, posY, 120, 20,
            parent, NULL, hInst, NULL);
    }
    SetWindowPos(hToolbarWnd, HWND_TOPMOST, posX, posY, 120, 20, SWP_SHOWWINDOW);
    ShowWindow(hToolbarWnd, SW_SHOWNOACTIVATE);
    toolbarVisible = true;
    SaveToolbarVisible(true);
}

void HideToolbar() {
    if (hToolbarWnd) {
        SaveToolbarPosition();
        ShowWindow(hToolbarWnd, SW_HIDE);
        toolbarVisible = false;
        SaveToolbarVisible(false);
    }
    HideToolbarTooltip();
}

void GetBatteryStatusString(TCHAR* buf, size_t len) {
    SYSTEM_POWER_STATUS sps;
    if (GetSystemPowerStatus(&sps)) {
        TCHAR ac[16] = _T("");
        switch (sps.ACLineStatus) {
        case 0: _tcscpy_s(ac, _T("On Battery")); break;
        case 1: _tcscpy_s(ac, _T("Plugged In")); break;
        default: _tcscpy_s(ac, _T("Unknown")); break;
        }
        TCHAR bat[16] = _T("");
        if (sps.BatteryFlag & 128) {
            StringCchCopy(buf, len, _T("A/C"));
            return;
        }
        else if (sps.BatteryLifePercent == 100 && sps.ACLineStatus == 1) {
            StringCchCopy(buf, len, _T("A/C"));
            return;
        }
        else if (sps.BatteryLifePercent == 255) {
            StringCchCopy(buf, len, _T("N/A"));
            return;
        }
        switch (sps.BatteryFlag) {
        case 1: _tcscpy_s(bat, _T("High")); break;
        case 2: _tcscpy_s(bat, _T("Low")); break;
        case 4: _tcscpy_s(bat, _T("Critical")); break;
        case 8: _tcscpy_s(bat, _T("Charging")); break;
        default: _tcscpy_s(bat, _T("Normal")); break;
        }
        StringCchPrintf(buf, len, _T("Battery: %d%% (%s, %s)"), sps.BatteryLifePercent, ac, bat);
    }
    else {
        StringCchCopy(buf, len, _T("N/A"));
    }
}

void UpdateTrayIcon() {
    TCHAR tip[128];
    GetBatteryStatusString(tip, _countof(tip));
    nid.uFlags = NIF_TIP;
    StringCchCopy(nid.szTip, _countof(nid.szTip), tip);
    Shell_NotifyIcon(NIM_MODIFY, &nid);
}

void ShowBatteryDetails(HWND parent) {
    TCHAR buf[1024];
    int percent = 0, timeSec = 0, acLineStatus = 0, batteryFlag = 0, milliWatts = 0;
    bool charging = false, haveWatt = false, haveSmartTime = false;
    double watts = 0;
    GetBatterySmartStatus(percent, timeSec, charging, watts, haveWatt, haveSmartTime, acLineStatus, batteryFlag, milliWatts);

    int ratePerHour = 0, sampleCount = 0;
    int histTime = EstimateTimeFromHistory(acLineStatus, percent, &ratePerHour, &sampleCount);

    TCHAR estbuf[32];
    if (histTime > 0) {
        FormatTime(histTime, charging, estbuf, _countof(estbuf));
    }
    else if (sampleCount < 2) {
        _tcscpy_s(estbuf, _T("Collecting data…"));
    }
    else {
        _tcscpy_s(estbuf, _T("N/A"));
    }

    // ==== WMI Battery Wear Section ====
    TCHAR wearbuf[128] = _T("");
    ULONG design = 0, full = 0;
    if (GetBatteryCapacities(&design, &full) && full <= design) {
        double wear = 100.0 * (1.0 - (double)full / (double)design);
        StringCchPrintf(wearbuf, _countof(wearbuf), _T("Battery Wear: %.1f%%\nDesign: %lu mWh\nFull: %lu mWh\n"), wear, design, full);
    }

    bool noBattery = (batteryFlag & 128) != 0;
    bool acAndFull = (acLineStatus == 1 && percent == 100);
    bool batteryUnknown = (percent == 255);

    if (noBattery || acAndFull) {
        StringCchPrintf(buf, _countof(buf), _T("Status: A/C\nDatabase Estimate: %s\n"), estbuf);
    }
    else if (batteryUnknown) {
        StringCchPrintf(buf, _countof(buf), _T("Status: N/A\n"));
    }
    else {
        TCHAR timebuf[16], wattbuf[16];
        FormatTime(timeSec, charging, timebuf, _countof(timebuf));
        if (haveWatt)
            StringCchPrintf(wattbuf, _countof(wattbuf), _T("%.1fW"), watts);
        else
            StringCchCopy(wattbuf, _countof(wattbuf), _T("--.-W"));
        StringCchPrintf(buf, _countof(buf),
            _T("Percent: %d%%\nACLineStatus: %d\nBatteryFlag: %d\nEstimated Time: %s\nDatabase Estimate: %s\nWattage: %s\nValid Samples: %d\n"),
            percent, acLineStatus, batteryFlag, timebuf, estbuf, wattbuf, sampleCount);

        BatterySample samples[MAX_SAMPLES];
        int n = ReadBatteryHistory(acLineStatus, samples, MAX_SAMPLES);
        if (n > 1) {
            TCHAR dbg[256];
            _stprintf_s(dbg, _T("Oldest: %d%% @ %I64d\nNewest: %d%% @ %I64d"),
                samples[0].percent, (LONGLONG)samples[0].t,
                samples[n - 1].percent, (LONGLONG)samples[n - 1].t);
            _tcscat_s(buf, _countof(buf), dbg);
        }
    }

    StringCchCat(buf, _countof(buf), wearbuf);

    MessageBox(parent, buf, _T("Battery Details"), MB_OK | MB_ICONINFORMATION);
    SetForegroundWindow(parent);
}

void ShowTrayMenu(HWND hwnd) {
    POINT pt;
    GetCursorPos(&pt);
    HMENU hMenu = CreatePopupMenu();
    AppendMenu(hMenu, MF_STRING | (toolbarVisible ? MF_CHECKED : 0), IDM_SHOW_TOOLBAR, _T("Show Toolbar"));
    AppendMenu(hMenu, MF_STRING | (IsAutoStartEnabled() ? MF_CHECKED : 0), IDM_AUTOSTART, _T("Auto start"));
    AppendMenu(hMenu, MF_SEPARATOR, 0, NULL);
    AppendMenu(hMenu, MF_STRING, IDM_EXIT, _T("Exit"));
    SetForegroundWindow(hwnd);
    int cmd = TrackPopupMenu(hMenu, TPM_RETURNCMD | TPM_NONOTIFY, pt.x, pt.y, 0, hwnd, NULL);
    DestroyMenu(hMenu);

    switch (cmd) {
    case IDM_EXIT:
        PostQuitMessage(0);
        break;
    case IDM_SHOW_TOOLBAR:
        if (toolbarVisible) {
            HideToolbar();
        }
        else {
            ShowToolbar(hwnd);
        }
        break;
    case IDM_AUTOSTART:
        if (IsAutoStartEnabled())
            SetAutoStart(false);
        else
            SetAutoStart(true);
        break;
    }
}

LRESULT CALLBACK WndProc(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam) {
    switch (msg) {
    case WM_CREATE: {
        INITCOMMONCONTROLSEX icc = { sizeof(icc), ICC_WIN95_CLASSES };
        InitCommonControlsEx(&icc);

        nid.cbSize = sizeof(NOTIFYICONDATA);
        nid.hWnd = hwnd;
        nid.uID = ID_TRAYICON;
        nid.uFlags = NIF_ICON | NIF_MESSAGE | NIF_TIP;
        nid.uCallbackMessage = WM_TRAYICON;
        nid.hIcon = LoadIcon(NULL, IDI_INFORMATION);
        StringCchCopy(nid.szTip, _countof(nid.szTip), _T("Battery Status"));
        Shell_NotifyIcon(NIM_ADD, &nid);

        SetTimer(hwnd, IDT_TIMER, 30000, NULL);
        UpdateTrayIcon();

        if (LoadToolbarVisible())
            ShowToolbar(hwnd);
        break;
    }
    case WM_TIMER:
        if (wParam == IDT_TIMER) {
            UpdateTrayIcon();
            if (toolbarVisible && hToolbarWnd)
                InvalidateRect(hToolbarWnd, NULL, FALSE);
        }
        break;
    case WM_TRAYICON:
        if (lParam == WM_LBUTTONUP) {
            ShowBatteryDetails(hwnd);
        }
        else if (lParam == WM_RBUTTONUP) {
            ShowTrayMenu(hwnd);
        }
        break;
    case WM_DESTROY:
        Shell_NotifyIcon(NIM_DELETE, &nid);
        if (hToolbarWnd) DestroyWindow(hToolbarWnd);
        if (hTooltip) DestroyWindow(hTooltip);
        PostQuitMessage(0);
        break;
    default:
        return DefWindowProc(hwnd, msg, wParam, lParam);
    }
    return 0;
}

int APIENTRY _tWinMain(HINSTANCE hInstance, HINSTANCE, LPTSTR, int) {
    hInst = hInstance;
    WNDCLASS wc = { 0 };
    wc.lpfnWndProc = WndProc;
    wc.hInstance = hInstance;
    wc.lpszClassName = _T("BattStatCloneClass");
    RegisterClass(&wc);

    hMainWnd = CreateWindow(wc.lpszClassName, _T("BattStat Clone"), 0,
        0, 0, 0, 0, NULL, NULL, hInstance, NULL);

    MSG msg;
    while (GetMessage(&msg, NULL, 0, 0)) {
        TranslateMessage(&msg);
        DispatchMessage(&msg);
    }
    return 0;
}
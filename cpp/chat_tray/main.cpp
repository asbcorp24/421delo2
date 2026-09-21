#include <windows.h>
#include <shellapi.h>
#include <commctrl.h>
#include <winhttp.h>
#include <winsqlite/winsqlite3.h>

#include <atomic>
#include <chrono>
#include <map>
#include <mutex>
#include <queue>
#include <sstream>
#include <string>
#include <thread>
#include <memory>
#include <vector>

#pragma comment(lib, "winhttp.lib")
#pragma comment(lib, "comctl32.lib")
#pragma comment(lib, "shell32.lib")

namespace {
constexpr UINT WMAPP_TRAY = WM_APP + 1;
constexpr UINT WMAPP_NEW_MESSAGE = WM_APP + 2;
constexpr UINT WMAPP_STATUS = WM_APP + 3;
constexpr UINT WMAPP_AUTH_REQUIRED = WM_APP + 4;

constexpr UINT ID_TRAY_OPEN = 1001;
constexpr UINT ID_TRAY_WEBCHAT = 1002;
constexpr UINT ID_TRAY_RECONNECT = 1003;
constexpr UINT ID_TRAY_SETTINGS = 1004;
constexpr UINT ID_TRAY_EXIT = 1005;

constexpr UINT ID_REPLY_EDIT = 2001;
constexpr UINT ID_SEND_BTN = 2002;
constexpr UINT ID_OPEN_BTN = 2003;
constexpr UINT ID_CLOSE_BTN = 2004;
constexpr UINT ID_RECIPIENT_COMBO = 2005;
constexpr UINT ID_SETTINGS_BASEURL = 3001;
constexpr UINT ID_SETTINGS_USERNAME = 3002;
constexpr UINT ID_SETTINGS_PASSWORD = 3003;
constexpr UINT ID_SETTINGS_POLL = 3004;
constexpr UINT ID_SETTINGS_IGNORE_SSL = 3005;
constexpr UINT ID_SETTINGS_SAVE = 3006;
constexpr UINT ID_SETTINGS_CANCEL = 3007;
constexpr UINT ID_AUTH_USER_COMBO = 4001;
constexpr UINT ID_AUTH_PASSWORD = 4002;
constexpr UINT ID_AUTH_LOGIN = 4003;
constexpr UINT ID_AUTH_CANCEL = 4004;

struct Config {
    std::wstring baseUrl = L"http://192.168.0.187:5001";
    std::wstring username;
    std::wstring password;
    int pollSeconds = 30;
    bool ignoreSslErrors = true;
};

struct TrayMessage {
    int id = 0;
    std::wstring kind;
    int senderId = 0;
    std::wstring sender;
    std::wstring text;
    std::wstring time;
};

struct ChatUser {
    int id = 0;
    std::wstring name;
    std::wstring role;
};

struct LoginChoice {
    std::wstring username;
    std::wstring displayName;
};

struct HttpResponse {
    DWORD status = 0;
    std::string body;
};

enum class SendChatResult {
    Success,
    AuthRequired,
    Failed,
};

constexpr COLORREF kBgColor = RGB(8, 15, 24);
constexpr COLORREF kPanelColor = RGB(12, 24, 38);
constexpr COLORREF kTextColor = RGB(225, 235, 245);
constexpr COLORREF kAccentColor = RGB(0, 224, 255);

void EnableDpiAwareness() {
    HMODULE user32 = GetModuleHandleW(L"user32.dll");
    if (user32) {
        using SetProcessDpiAwarenessContextFn = BOOL(WINAPI*)(HANDLE);
        auto setContext = reinterpret_cast<SetProcessDpiAwarenessContextFn>(
            GetProcAddress(user32, "SetProcessDpiAwarenessContext")
        );
        if (setContext) {
            if (setContext(reinterpret_cast<HANDLE>(-4))) {
                return;
            }
        }

        using SetProcessDPIAwareFn = BOOL(WINAPI*)();
        auto setAware = reinterpret_cast<SetProcessDPIAwareFn>(
            GetProcAddress(user32, "SetProcessDPIAware")
        );
        if (setAware) {
            setAware();
        }
    }
}

UINT GetSystemDpiValue() {
    HDC screen = GetDC(nullptr);
    if (!screen) return 96;
    int dpi = GetDeviceCaps(screen, LOGPIXELSX);
    ReleaseDC(nullptr, screen);
    return dpi > 0 ? static_cast<UINT>(dpi) : 96;
}

std::wstring Utf8ToWide(const std::string& value) {
    if (value.empty()) return L"";
    int len = MultiByteToWideChar(CP_UTF8, 0, value.c_str(), static_cast<int>(value.size()), nullptr, 0);
    std::wstring out(len, L'\0');
    MultiByteToWideChar(CP_UTF8, 0, value.c_str(), static_cast<int>(value.size()), out.data(), len);
    return out;
}

std::string WideToUtf8(const std::wstring& value) {
    if (value.empty()) return "";
    int len = WideCharToMultiByte(CP_UTF8, 0, value.c_str(), static_cast<int>(value.size()), nullptr, 0, nullptr, nullptr);
    std::string out(len, '\0');
    WideCharToMultiByte(CP_UTF8, 0, value.c_str(), static_cast<int>(value.size()), out.data(), len, nullptr, nullptr);
    return out;
}

std::wstring PercentDecodeUtf8(const std::wstring& value) {
    std::string bytes;
    for (size_t i = 0; i < value.size(); ++i) {
        if (value[i] == L'%' && i + 2 < value.size()) {
            wchar_t hex[3] = { value[i + 1], value[i + 2], 0 };
            char ch = static_cast<char>(wcstol(hex, nullptr, 16));
            bytes.push_back(ch);
            i += 2;
        } else if (value[i] == L'+') {
            bytes.push_back(' ');
        } else {
            bytes.push_back(static_cast<char>(value[i]));
        }
    }
    return Utf8ToWide(bytes);
}

std::string UrlEncodeUtf8(const std::wstring& value) {
    const std::string utf8 = WideToUtf8(value);
    std::ostringstream out;
    static const char* hex = "0123456789ABCDEF";
    for (unsigned char c : utf8) {
        const bool safe =
            (c >= 'a' && c <= 'z') ||
            (c >= 'A' && c <= 'Z') ||
            (c >= '0' && c <= '9') ||
            c == '-' || c == '_' || c == '.' || c == '~';
        if (safe) {
            out << static_cast<char>(c);
        } else if (c == ' ') {
            out << '+';
        } else {
            out << '%' << hex[(c >> 4) & 0xF] << hex[c & 0xF];
        }
    }
    return out.str();
}

std::vector<std::wstring> Split(const std::wstring& value, wchar_t delim) {
    std::vector<std::wstring> parts;
    std::wstring current;
    for (wchar_t ch : value) {
        if (ch == delim) {
            parts.push_back(current);
            current.clear();
        } else {
            current.push_back(ch);
        }
    }
    parts.push_back(current);
    return parts;
}

std::wstring ReadIniString(const std::wstring& path, const wchar_t* section, const wchar_t* key, const wchar_t* fallback) {
    wchar_t buffer[2048] = {};
    GetPrivateProfileStringW(section, key, fallback, buffer, static_cast<DWORD>(std::size(buffer)), path.c_str());
    return buffer;
}

void WriteIniString(const std::wstring& path, const wchar_t* section, const wchar_t* key, const std::wstring& value) {
    WritePrivateProfileStringW(section, key, value.c_str(), path.c_str());
}

class HttpClient {
public:
    explicit HttpClient(Config config) : config_(std::move(config)) {
        session_ = WinHttpOpen(L"OTD421ChatTray/1.0", WINHTTP_ACCESS_TYPE_DEFAULT_PROXY, WINHTTP_NO_PROXY_NAME, WINHTTP_NO_PROXY_BYPASS, 0);
        ParseBaseUrl();
    }

    ~HttpClient() {
        CloseTrayWebSocket();
        if (session_) WinHttpCloseHandle(session_);
    }

    bool Login() {
        std::lock_guard<std::mutex> lock(mutex_);
        cookies_.clear();
        const std::string body = "username=" + UrlEncodeUtf8(config_.username) + "&password=" + UrlEncodeUtf8(config_.password);
        HttpResponse response = RequestLocked(L"POST", L"/login", body, L"Content-Type: application/x-www-form-urlencoded\r\n");
        return (response.status == 302 || response.status == 200) && !cookies_.empty();
    }

    SendChatResult SendChatMessage(const std::wstring& receiverId, const std::wstring& text) {
        std::lock_guard<std::mutex> lock(mutex_);
        const std::wstring target = receiverId.empty() ? L"all" : receiverId;
        const std::string body = "receiver_id=" + UrlEncodeUtf8(target) + "&text=" + UrlEncodeUtf8(text);
        HttpResponse response = RequestLocked(L"POST", L"/api/chat/send", body, L"Content-Type: application/x-www-form-urlencoded\r\n");
        if (response.status == 302 || response.status == 401 || response.status == 403) {
            return SendChatResult::AuthRequired;
        }
        if (response.status != 200) {
            return SendChatResult::Failed;
        }
        if (response.body.find("\"success\":true") != std::string::npos ||
            response.body.find("\"success\": true") != std::string::npos ||
            response.body.find("\"success\"") != std::string::npos) {
            return SendChatResult::Success;
        }
        return SendChatResult::Failed;
    }

    bool PollIncoming(int& directAfterId, int& generalAfterId, std::vector<TrayMessage>& messages) {
        std::lock_guard<std::mutex> lock(mutex_);
        std::wstringstream path;
        path << L"/api/chat/tray/incoming?direct_after_id=" << directAfterId << L"&general_after_id=" << generalAfterId;
        HttpResponse response = RequestLocked(L"GET", path.str(), "", L"");
        if (response.status != 200) return false;

        return ParseTrayPayload(Utf8ToWide(response.body), directAfterId, generalAfterId, messages);
    }

    bool OpenTrayWebSocket(int directAfterId, int generalAfterId) {
        std::lock_guard<std::mutex> lock(mutex_);
        CloseTrayWebSocketLocked();
        if (!session_) return false;

        HINTERNET connect = WinHttpConnect(session_, host_.c_str(), port_, 0);
        if (!connect) return false;

        std::wstringstream path;
        path << basePath_ << L"/ws/chat-tray?direct_after_id=" << directAfterId << L"&general_after_id=" << generalAfterId;
        HINTERNET request = WinHttpOpenRequest(
            connect,
            L"GET",
            path.str().c_str(),
            nullptr,
            WINHTTP_NO_REFERER,
            WINHTTP_DEFAULT_ACCEPT_TYPES,
            secure_ ? WINHTTP_FLAG_SECURE : 0
        );
        if (!request) {
            WinHttpCloseHandle(connect);
            return false;
        }

        if (secure_ && config_.ignoreSslErrors) {
            DWORD flags = SECURITY_FLAG_IGNORE_UNKNOWN_CA |
                          SECURITY_FLAG_IGNORE_CERT_CN_INVALID |
                          SECURITY_FLAG_IGNORE_CERT_DATE_INVALID |
                          SECURITY_FLAG_IGNORE_CERT_WRONG_USAGE;
            WinHttpSetOption(request, WINHTTP_OPTION_SECURITY_FLAGS, &flags, sizeof(flags));
        }

        std::wstring headers;
        if (!cookies_.empty()) {
            headers += L"Cookie: " + cookies_ + L"\r\n";
        }
        WinHttpSetOption(request, WINHTTP_OPTION_UPGRADE_TO_WEB_SOCKET, nullptr, 0);
        BOOL ok = WinHttpSendRequest(
            request,
            headers.empty() ? WINHTTP_NO_ADDITIONAL_HEADERS : headers.c_str(),
            headers.empty() ? 0 : static_cast<DWORD>(-1L),
            WINHTTP_NO_REQUEST_DATA,
            0,
            0,
            0
        );
        if (ok) ok = WinHttpReceiveResponse(request, nullptr);
        if (!ok) {
            WinHttpCloseHandle(request);
            WinHttpCloseHandle(connect);
            return false;
        }

        DWORD status = 0;
        DWORD statusSize = sizeof(status);
        WinHttpQueryHeaders(request, WINHTTP_QUERY_STATUS_CODE | WINHTTP_QUERY_FLAG_NUMBER, WINHTTP_HEADER_NAME_BY_INDEX, &status, &statusSize, WINHTTP_NO_HEADER_INDEX);
        if (status != 101) {
            WinHttpCloseHandle(request);
            WinHttpCloseHandle(connect);
            return false;
        }

        ApplySetCookieHeaders(request);
        websocket_ = WinHttpWebSocketCompleteUpgrade(request, 0);
        WinHttpCloseHandle(request);
        WinHttpCloseHandle(connect);
        return websocket_ != nullptr;
    }

    bool ReceiveTrayWebSocket(int& directAfterId, int& generalAfterId, std::vector<TrayMessage>& messages, bool& authRequired) {
        authRequired = false;
        messages.clear();

        std::string payload;
        while (true) {
            HINTERNET socket = nullptr;
            {
                std::lock_guard<std::mutex> lock(mutex_);
                socket = websocket_;
            }
            if (!socket) return false;

            char buffer[8192];
            DWORD read = 0;
            WINHTTP_WEB_SOCKET_BUFFER_TYPE type = WINHTTP_WEB_SOCKET_BINARY_FRAGMENT_BUFFER_TYPE;
            DWORD result = WinHttpWebSocketReceive(socket, buffer, sizeof(buffer), &read, &type);
            if (result != NO_ERROR) {
                CloseTrayWebSocket();
                return false;
            }

            if (type == WINHTTP_WEB_SOCKET_CLOSE_BUFFER_TYPE) {
                authRequired = false;
                CloseTrayWebSocket();
                return false;
            }

            if (type == WINHTTP_WEB_SOCKET_UTF8_FRAGMENT_BUFFER_TYPE || type == WINHTTP_WEB_SOCKET_UTF8_MESSAGE_BUFFER_TYPE) {
                payload.append(buffer, buffer + read);
                if (type == WINHTTP_WEB_SOCKET_UTF8_MESSAGE_BUFFER_TYPE) {
                    break;
                }
                continue;
            }
        }

        std::wstring content = Utf8ToWide(payload);
        if (content == L"PONG") {
            return true;
        }
        return ParseTrayPayload(content, directAfterId, generalAfterId, messages);
    }

    void CloseTrayWebSocket() {
        std::lock_guard<std::mutex> lock(mutex_);
        CloseTrayWebSocketLocked();
    }

    bool FetchUsers(std::vector<ChatUser>& users) {
        std::lock_guard<std::mutex> lock(mutex_);
        HttpResponse response = RequestLocked(L"GET", L"/api/chat/tray/users", "", L"");
        if (response.status != 200) return false;

        std::wstring content = Utf8ToWide(response.body);
        std::wistringstream stream(content);
        std::wstring line;
        std::vector<ChatUser> parsed;

        while (std::getline(stream, line)) {
            if (!line.empty() && line.back() == L'\r') line.pop_back();
            if (line.rfind(L"USER\t", 0) != 0) continue;
            auto parts = Split(line, L'\t');
            if (parts.size() < 4) continue;

            ChatUser user;
            user.id = _wtoi(parts[1].c_str());
            user.name = PercentDecodeUtf8(parts[2]);
            user.role = PercentDecodeUtf8(parts[3]);
            parsed.push_back(std::move(user));
        }

        users = std::move(parsed);
        return true;
    }

    bool FetchLoginUsers(std::vector<LoginChoice>& users) {
        std::lock_guard<std::mutex> lock(mutex_);
        HttpResponse response = RequestLocked(L"GET", L"/api/chat/tray/login-users", "", L"");
        if (response.status != 200) return false;

        std::wstring content = Utf8ToWide(response.body);
        std::wistringstream stream(content);
        std::wstring line;
        std::vector<LoginChoice> parsed;

        while (std::getline(stream, line)) {
            if (!line.empty() && line.back() == L'\r') line.pop_back();
            if (line.rfind(L"USER\t", 0) != 0) continue;
            auto parts = Split(line, L'\t');
            if (parts.size() < 3) continue;

            LoginChoice user;
            user.username = PercentDecodeUtf8(parts[1]);
            user.displayName = PercentDecodeUtf8(parts[2]);
            parsed.push_back(std::move(user));
        }

        users = std::move(parsed);
        return true;
    }

    std::wstring BaseUrl() const { return config_.baseUrl; }

private:
    bool ParseTrayPayload(const std::wstring& content, int& directAfterId, int& generalAfterId, std::vector<TrayMessage>& messages) {
        std::wistringstream stream(content);
        std::wstring line;
        int newDirect = directAfterId;
        int newGeneral = generalAfterId;
        std::vector<TrayMessage> parsed;

        while (std::getline(stream, line)) {
            if (!line.empty() && line.back() == L'\r') line.pop_back();
            if (line.empty()) continue;
            if (line == L"PONG") continue;
            if (line.rfind(L"DIRECT_LAST\t", 0) == 0) {
                newDirect = _wtoi(line.substr(12).c_str());
                continue;
            }
            if (line.rfind(L"GENERAL_LAST\t", 0) == 0) {
                newGeneral = _wtoi(line.substr(13).c_str());
                continue;
            }
            if (line.rfind(L"MSG\t", 0) != 0) continue;

            auto parts = Split(line, L'\t');
            if (parts.size() < 7) continue;

            TrayMessage msg;
            msg.id = _wtoi(parts[1].c_str());
            msg.kind = parts[2];
            msg.senderId = _wtoi(parts[3].c_str());
            msg.sender = PercentDecodeUtf8(parts[4]);
            msg.text = PercentDecodeUtf8(parts[5]);
            msg.time = PercentDecodeUtf8(parts[6]);
            parsed.push_back(std::move(msg));
        }

        directAfterId = newDirect;
        generalAfterId = newGeneral;
        messages = std::move(parsed);
        return true;
    }

    void ParseBaseUrl() {
        URL_COMPONENTS parts{};
        parts.dwStructSize = sizeof(parts);
        wchar_t host[256] = {};
        wchar_t path[1024] = {};
        parts.lpszHostName = host;
        parts.dwHostNameLength = static_cast<DWORD>(std::size(host));
        parts.lpszUrlPath = path;
        parts.dwUrlPathLength = static_cast<DWORD>(std::size(path));
        if (!WinHttpCrackUrl(config_.baseUrl.c_str(), 0, 0, &parts)) {
            secure_ = true;
            host_ = L"192.168.0.187";
            port_ = INTERNET_DEFAULT_HTTPS_PORT;
            basePath_ = L"";
            return;
        }
        secure_ = parts.nScheme == INTERNET_SCHEME_HTTPS;
        host_.assign(parts.lpszHostName, parts.dwHostNameLength);
        port_ = parts.nPort;
        basePath_.assign(parts.lpszUrlPath ? parts.lpszUrlPath : L"", parts.dwUrlPathLength);
        if (basePath_ == L"/") basePath_.clear();
    }

    void ApplySetCookieHeaders(HINTERNET request) {
        DWORD index = 0;
        while (true) {
            DWORD size = 0;
            WinHttpQueryHeaders(request, WINHTTP_QUERY_SET_COOKIE, WINHTTP_HEADER_NAME_BY_INDEX, nullptr, &size, &index);
            if (GetLastError() != ERROR_INSUFFICIENT_BUFFER) break;

            std::wstring value(size / sizeof(wchar_t), L'\0');
            if (!WinHttpQueryHeaders(request, WINHTTP_QUERY_SET_COOKIE, WINHTTP_HEADER_NAME_BY_INDEX, value.data(), &size, &index)) {
                break;
            }
            if (!value.empty() && value.back() == L'\0') value.pop_back();
            auto semicolon = value.find(L';');
            std::wstring pair = semicolon == std::wstring::npos ? value : value.substr(0, semicolon);
            auto eq = pair.find(L'=');
            if (eq == std::wstring::npos) continue;
            cookiesMap_[pair.substr(0, eq)] = pair.substr(eq + 1);
            ++index;
        }

        std::wostringstream oss;
        bool first = true;
        for (const auto& [name, val] : cookiesMap_) {
            if (!first) oss << L"; ";
            oss << name << L"=" << val;
            first = false;
        }
        cookies_ = oss.str();
    }

    HttpResponse RequestLocked(const std::wstring& method, const std::wstring& path, const std::string& body, const std::wstring& headers) {
        HttpResponse response;
        if (!session_) return response;

        HINTERNET connect = WinHttpConnect(session_, host_.c_str(), port_, 0);
        if (!connect) return response;

        std::wstring fullPath = basePath_ + path;
        if (fullPath.empty()) fullPath = L"/";

        HINTERNET request = WinHttpOpenRequest(
            connect,
            method.c_str(),
            fullPath.c_str(),
            nullptr,
            WINHTTP_NO_REFERER,
            WINHTTP_DEFAULT_ACCEPT_TYPES,
            secure_ ? WINHTTP_FLAG_SECURE : 0
        );
        if (!request) {
            WinHttpCloseHandle(connect);
            return response;
        }

        if (secure_ && config_.ignoreSslErrors) {
            DWORD flags = SECURITY_FLAG_IGNORE_UNKNOWN_CA |
                          SECURITY_FLAG_IGNORE_CERT_CN_INVALID |
                          SECURITY_FLAG_IGNORE_CERT_DATE_INVALID |
                          SECURITY_FLAG_IGNORE_CERT_WRONG_USAGE;
            WinHttpSetOption(request, WINHTTP_OPTION_SECURITY_FLAGS, &flags, sizeof(flags));
        }

        std::wstring allHeaders = headers;
        if (!cookies_.empty()) {
            allHeaders += L"Cookie: " + cookies_ + L"\r\n";
        }

        BOOL ok = WinHttpSendRequest(
            request,
            allHeaders.empty() ? WINHTTP_NO_ADDITIONAL_HEADERS : allHeaders.c_str(),
            allHeaders.empty() ? 0 : static_cast<DWORD>(-1L),
            body.empty() ? WINHTTP_NO_REQUEST_DATA : const_cast<char*>(body.data()),
            static_cast<DWORD>(body.size()),
            static_cast<DWORD>(body.size()),
            0
        );

        if (ok) ok = WinHttpReceiveResponse(request, nullptr);
        if (!ok) {
            WinHttpCloseHandle(request);
            WinHttpCloseHandle(connect);
            return response;
        }

        DWORD status = 0;
        DWORD statusSize = sizeof(status);
        WinHttpQueryHeaders(request, WINHTTP_QUERY_STATUS_CODE | WINHTTP_QUERY_FLAG_NUMBER, WINHTTP_HEADER_NAME_BY_INDEX, &status, &statusSize, WINHTTP_NO_HEADER_INDEX);
        response.status = status;

        ApplySetCookieHeaders(request);

        DWORD available = 0;
        do {
            available = 0;
            if (!WinHttpQueryDataAvailable(request, &available) || !available) break;
            std::string chunk(available, '\0');
            DWORD read = 0;
            if (!WinHttpReadData(request, chunk.data(), available, &read) || !read) break;
            chunk.resize(read);
            response.body += chunk;
        } while (available > 0);

        WinHttpCloseHandle(request);
        WinHttpCloseHandle(connect);
        return response;
    }

    void CloseTrayWebSocketLocked() {
        if (websocket_) {
            WinHttpWebSocketClose(websocket_, WINHTTP_WEB_SOCKET_SUCCESS_CLOSE_STATUS, nullptr, 0);
            WinHttpCloseHandle(websocket_);
            websocket_ = nullptr;
        }
    }

    Config config_;
    HINTERNET session_ = nullptr;
    HINTERNET websocket_ = nullptr;
    std::wstring host_;
    INTERNET_PORT port_ = 0;
    std::wstring basePath_;
    bool secure_ = true;
    std::mutex mutex_;
    std::map<std::wstring, std::wstring> cookiesMap_;
    std::wstring cookies_;
};

class TrayApp {
public:
    explicit TrayApp(HINSTANCE instance) : instance_(instance) {}

    int Run() {
        EnableDpiAwareness();
        dpi_ = GetSystemDpiValue();
        config_ = LoadConfig();
        EnsureLocalDatabase();
        LoadChatState();
        client_ = std::make_unique<HttpClient>(config_);
        InitCommonControls();
        InitTheme();
        RegisterWindowClasses();

        hwnd_ = CreateWindowExW(0, kMainClass, L"OTD421 Chat Tray", WS_OVERLAPPEDWINDOW,
            CW_USEDEFAULT, CW_USEDEFAULT, Scale(320), Scale(240), nullptr, nullptr, instance_, this);
        popup_ = CreateWindowExW(WS_EX_TOOLWINDOW | WS_EX_TOPMOST, kPopupClass, L"Новое сообщение",
            WS_POPUP | WS_BORDER | WS_CAPTION, 0, 0, Scale(460), Scale(370), hwnd_, nullptr, instance_, this);
        settingsWnd_ = CreateWindowExW(WS_EX_TOOLWINDOW, kSettingsClass, L"Настройки OTD421 Chat Tray",
            WS_OVERLAPPED | WS_CAPTION | WS_SYSMENU | WS_MINIMIZEBOX, CW_USEDEFAULT, CW_USEDEFAULT, Scale(540), Scale(360),
            hwnd_, nullptr, instance_, this);
        authWnd_ = CreateWindowExW(WS_EX_TOOLWINDOW, kAuthClass, L"Вход в OTD421 Messenger",
            WS_OVERLAPPED | WS_CAPTION | WS_SYSMENU | WS_MINIMIZEBOX, CW_USEDEFAULT, CW_USEDEFAULT, Scale(500), Scale(290),
            hwnd_, nullptr, instance_, this);
        CreatePopupControls();
        CreateSettingsControls();
        CreateAuthControls();

        AddTrayIcon();
        StartWorker();
        if (config_.baseUrl.empty()) {
            ShowSettingsWindow();
        }

        MSG msg;
        while (GetMessageW(&msg, nullptr, 0, 0)) {
            TranslateMessage(&msg);
            DispatchMessageW(&msg);
        }

        running_ = false;
        client_->CloseTrayWebSocket();
        if (worker_.joinable()) worker_.join();
        RemoveTrayIcon();
        DestroyTheme();
        return static_cast<int>(msg.wParam);
    }

private:
    static constexpr wchar_t kMainClass[] = L"OTD421_TRAY_MAIN";
    static constexpr wchar_t kPopupClass[] = L"OTD421_TRAY_POPUP";
    static constexpr wchar_t kSettingsClass[] = L"OTD421_TRAY_SETTINGS";
    static constexpr wchar_t kAuthClass[] = L"OTD421_TRAY_AUTH";

    int Scale(int value) const {
        return MulDiv(value, static_cast<int>(dpi_), 96);
    }

    void InitTheme() {
        bgBrush_ = CreateSolidBrush(kBgColor);
        panelBrush_ = CreateSolidBrush(kPanelColor);
        accentBrush_ = CreateSolidBrush(kAccentColor);

        LOGFONTW lf{};
        lf.lfHeight = -Scale(18);
        wcscpy_s(lf.lfFaceName, L"Segoe UI");
        lf.lfQuality = CLEARTYPE_QUALITY;
        uiFont_ = CreateFontIndirectW(&lf);
    }

    void DestroyTheme() {
        if (uiFont_) DeleteObject(uiFont_);
        if (bgBrush_) DeleteObject(bgBrush_);
        if (panelBrush_) DeleteObject(panelBrush_);
        if (accentBrush_) DeleteObject(accentBrush_);
    }

    static LRESULT CALLBACK MainProc(HWND hwnd, UINT msg, WPARAM wp, LPARAM lp) {
        auto* self = reinterpret_cast<TrayApp*>(GetWindowLongPtrW(hwnd, GWLP_USERDATA));
        if (msg == WM_NCCREATE) {
            auto* cs = reinterpret_cast<CREATESTRUCTW*>(lp);
            self = reinterpret_cast<TrayApp*>(cs->lpCreateParams);
            SetWindowLongPtrW(hwnd, GWLP_USERDATA, reinterpret_cast<LONG_PTR>(self));
        }
        return self ? self->HandleMain(hwnd, msg, wp, lp) : DefWindowProcW(hwnd, msg, wp, lp);
    }

    static LRESULT CALLBACK PopupProc(HWND hwnd, UINT msg, WPARAM wp, LPARAM lp) {
        auto* self = reinterpret_cast<TrayApp*>(GetWindowLongPtrW(hwnd, GWLP_USERDATA));
        if (msg == WM_NCCREATE) {
            auto* cs = reinterpret_cast<CREATESTRUCTW*>(lp);
            self = reinterpret_cast<TrayApp*>(cs->lpCreateParams);
            SetWindowLongPtrW(hwnd, GWLP_USERDATA, reinterpret_cast<LONG_PTR>(self));
        }
        return self ? self->HandlePopup(hwnd, msg, wp, lp) : DefWindowProcW(hwnd, msg, wp, lp);
    }

    static LRESULT CALLBACK SettingsProc(HWND hwnd, UINT msg, WPARAM wp, LPARAM lp) {
        auto* self = reinterpret_cast<TrayApp*>(GetWindowLongPtrW(hwnd, GWLP_USERDATA));
        if (msg == WM_NCCREATE) {
            auto* cs = reinterpret_cast<CREATESTRUCTW*>(lp);
            self = reinterpret_cast<TrayApp*>(cs->lpCreateParams);
            SetWindowLongPtrW(hwnd, GWLP_USERDATA, reinterpret_cast<LONG_PTR>(self));
        }
        return self ? self->HandleSettings(hwnd, msg, wp, lp) : DefWindowProcW(hwnd, msg, wp, lp);
    }

    static LRESULT CALLBACK AuthProc(HWND hwnd, UINT msg, WPARAM wp, LPARAM lp) {
        auto* self = reinterpret_cast<TrayApp*>(GetWindowLongPtrW(hwnd, GWLP_USERDATA));
        if (msg == WM_NCCREATE) {
            auto* cs = reinterpret_cast<CREATESTRUCTW*>(lp);
            self = reinterpret_cast<TrayApp*>(cs->lpCreateParams);
            SetWindowLongPtrW(hwnd, GWLP_USERDATA, reinterpret_cast<LONG_PTR>(self));
        }
        return self ? self->HandleAuth(hwnd, msg, wp, lp) : DefWindowProcW(hwnd, msg, wp, lp);
    }

    void RegisterWindowClasses() {
        WNDCLASSW wc{};
        wc.lpfnWndProc = MainProc;
        wc.hInstance = instance_;
        wc.lpszClassName = kMainClass;
        RegisterClassW(&wc);

        WNDCLASSW popup{};
        popup.lpfnWndProc = PopupProc;
        popup.hInstance = instance_;
        popup.lpszClassName = kPopupClass;
        popup.hbrBackground = bgBrush_;
        popup.hCursor = LoadCursor(nullptr, IDC_ARROW);
        RegisterClassW(&popup);

        WNDCLASSW settings{};
        settings.lpfnWndProc = SettingsProc;
        settings.hInstance = instance_;
        settings.lpszClassName = kSettingsClass;
        settings.hbrBackground = bgBrush_;
        settings.hCursor = LoadCursor(nullptr, IDC_ARROW);
        RegisterClassW(&settings);

        WNDCLASSW auth{};
        auth.lpfnWndProc = AuthProc;
        auth.hInstance = instance_;
        auth.lpszClassName = kAuthClass;
        auth.hbrBackground = bgBrush_;
        auth.hCursor = LoadCursor(nullptr, IDC_ARROW);
        RegisterClassW(&auth);
    }

    void CreatePopupControls() {
        SetWindowLongPtrW(popup_, GWLP_USERDATA, reinterpret_cast<LONG_PTR>(this));
        senderLabel_ = CreateWindowW(L"STATIC", L"OTD421 Messenger", WS_CHILD | WS_VISIBLE, Scale(14), Scale(12), Scale(400), Scale(22), popup_, nullptr, instance_, nullptr);
        kindLabel_ = CreateWindowW(L"STATIC", L"Быстрый ответ", WS_CHILD | WS_VISIBLE, Scale(14), Scale(38), Scale(400), Scale(20), popup_, nullptr, instance_, nullptr);
        transportLabel_ = CreateWindowW(L"STATIC", L"Режим: подключение...", WS_CHILD | WS_VISIBLE, Scale(14), Scale(58), Scale(220), Scale(20), popup_, nullptr, instance_, nullptr);
        recipientLabel_ = CreateWindowW(L"STATIC", L"Кому:", WS_CHILD | WS_VISIBLE, Scale(14), Scale(66), Scale(60), Scale(20), popup_, nullptr, instance_, nullptr);
        recipientCombo_ = CreateWindowW(
            WC_COMBOBOXW,
            L"",
            WS_CHILD | WS_VISIBLE | CBS_DROPDOWNLIST | WS_VSCROLL,
            Scale(84), Scale(62), Scale(346), Scale(260),
            popup_,
            reinterpret_cast<HMENU>(static_cast<INT_PTR>(ID_RECIPIENT_COMBO)),
            instance_,
            nullptr
        );
        textView_ = CreateWindowExW(WS_EX_CLIENTEDGE, L"EDIT", L"", WS_CHILD | WS_VISIBLE | ES_MULTILINE | ES_AUTOVSCROLL | ES_READONLY | WS_VSCROLL,
            Scale(14), Scale(100), Scale(416), Scale(100), popup_, nullptr, instance_, nullptr);
        replyEdit_ = CreateWindowExW(WS_EX_CLIENTEDGE, L"EDIT", L"", WS_CHILD | WS_VISIBLE | ES_MULTILINE | ES_AUTOVSCROLL | WS_VSCROLL,
            Scale(14), Scale(212), Scale(416), Scale(76), popup_, reinterpret_cast<HMENU>(static_cast<INT_PTR>(ID_REPLY_EDIT)), instance_, nullptr);
        sendBtn_ = CreateWindowW(L"BUTTON", L"Отправить", WS_CHILD | WS_VISIBLE | BS_FLAT, Scale(14), Scale(300), Scale(124), Scale(34), popup_, reinterpret_cast<HMENU>(static_cast<INT_PTR>(ID_SEND_BTN)), instance_, nullptr);
        openBtn_ = CreateWindowW(L"BUTTON", L"Открыть чат", WS_CHILD | WS_VISIBLE | BS_FLAT, Scale(160), Scale(300), Scale(124), Scale(34), popup_, reinterpret_cast<HMENU>(static_cast<INT_PTR>(ID_OPEN_BTN)), instance_, nullptr);
        closeBtn_ = CreateWindowW(L"BUTTON", L"Скрыть", WS_CHILD | WS_VISIBLE | BS_FLAT, Scale(306), Scale(300), Scale(124), Scale(34), popup_, reinterpret_cast<HMENU>(static_cast<INT_PTR>(ID_CLOSE_BTN)), instance_, nullptr);
        ApplyControlFont(senderLabel_);
        ApplyControlFont(kindLabel_);
        ApplyControlFont(transportLabel_);
        ApplyControlFont(recipientLabel_);
        ApplyControlFont(recipientCombo_);
        ApplyControlFont(textView_);
        ApplyControlFont(replyEdit_);
        ApplyControlFont(sendBtn_);
        ApplyControlFont(openBtn_);
        ApplyControlFont(closeBtn_);
        PopulateRecipients();
    }

    void AddTrayIcon() {
        NOTIFYICONDATAW nid{};
        nid.cbSize = sizeof(nid);
        nid.hWnd = hwnd_;
        nid.uID = 1;
        nid.uFlags = NIF_ICON | NIF_MESSAGE | NIF_TIP;
        nid.uCallbackMessage = WMAPP_TRAY;
        nid.hIcon = LoadIcon(nullptr, IDI_INFORMATION);
        wcscpy_s(nid.szTip, L"OTD421 Chat Tray");
        Shell_NotifyIconW(NIM_ADD, &nid);
    }

    void CreateSettingsControls() {
        SetWindowLongPtrW(settingsWnd_, GWLP_USERDATA, reinterpret_cast<LONG_PTR>(this));
        settingsTitle_ = CreateWindowW(L"STATIC", L"Настройки Tray Messenger", WS_CHILD | WS_VISIBLE, Scale(16), Scale(12), Scale(360), Scale(24), settingsWnd_, nullptr, instance_, nullptr);
        CreateWindowW(L"STATIC", L"Сервер:", WS_CHILD | WS_VISIBLE, Scale(16), Scale(48), Scale(110), Scale(20), settingsWnd_, nullptr, instance_, nullptr);
        settingsBaseUrl_ = CreateWindowExW(WS_EX_CLIENTEDGE, L"EDIT", config_.baseUrl.c_str(), WS_CHILD | WS_VISIBLE | ES_AUTOHSCROLL,
            Scale(140), Scale(44), Scale(360), Scale(26), settingsWnd_, reinterpret_cast<HMENU>(static_cast<INT_PTR>(ID_SETTINGS_BASEURL)), instance_, nullptr);
        CreateWindowW(L"STATIC", L"Опрос, сек:", WS_CHILD | WS_VISIBLE, Scale(16), Scale(92), Scale(110), Scale(20), settingsWnd_, nullptr, instance_, nullptr);
        settingsPoll_ = CreateWindowExW(WS_EX_CLIENTEDGE, L"EDIT", std::to_wstring(config_.pollSeconds).c_str(), WS_CHILD | WS_VISIBLE | ES_AUTOHSCROLL,
            Scale(140), Scale(88), Scale(120), Scale(26), settingsWnd_, reinterpret_cast<HMENU>(static_cast<INT_PTR>(ID_SETTINGS_POLL)), instance_, nullptr);

        settingsIgnoreSsl_ = CreateWindowW(L"BUTTON", L"Игнорировать ошибки SSL", WS_CHILD | WS_VISIBLE | BS_AUTOCHECKBOX,
            Scale(140), Scale(130), Scale(260), Scale(24), settingsWnd_, reinterpret_cast<HMENU>(static_cast<INT_PTR>(ID_SETTINGS_IGNORE_SSL)), instance_, nullptr);
        SendMessageW(settingsIgnoreSsl_, BM_SETCHECK, config_.ignoreSslErrors ? BST_CHECKED : BST_UNCHECKED, 0);

        settingsHint_ = CreateWindowW(L"STATIC", L"После сохранения программа покажет отдельное окно входа с выбором пользователя.", WS_CHILD | WS_VISIBLE,
            Scale(16), Scale(174), Scale(484), Scale(42), settingsWnd_, nullptr, instance_, nullptr);

        settingsSaveBtn_ = CreateWindowW(L"BUTTON", L"Сохранить", WS_CHILD | WS_VISIBLE | BS_FLAT, Scale(140), Scale(262), Scale(140), Scale(34), settingsWnd_,
            reinterpret_cast<HMENU>(static_cast<INT_PTR>(ID_SETTINGS_SAVE)), instance_, nullptr);
        settingsCancelBtn_ = CreateWindowW(L"BUTTON", L"Отмена", WS_CHILD | WS_VISIBLE | BS_FLAT, Scale(296), Scale(262), Scale(140), Scale(34), settingsWnd_,
            reinterpret_cast<HMENU>(static_cast<INT_PTR>(ID_SETTINGS_CANCEL)), instance_, nullptr);
        ApplyControlFont(settingsTitle_);
        ApplyControlFont(settingsBaseUrl_);
        ApplyControlFont(settingsPoll_);
        ApplyControlFont(settingsIgnoreSsl_);
        ApplyControlFont(settingsHint_);
        ApplyControlFont(settingsSaveBtn_);
        ApplyControlFont(settingsCancelBtn_);
    }

    void CreateAuthControls() {
        SetWindowLongPtrW(authWnd_, GWLP_USERDATA, reinterpret_cast<LONG_PTR>(this));
        authTitle_ = CreateWindowW(L"STATIC", L"Подключение к чату OTD421", WS_CHILD | WS_VISIBLE, Scale(16), Scale(12), Scale(420), Scale(24), authWnd_, nullptr, instance_, nullptr);
        CreateWindowW(L"STATIC", L"Пользователь:", WS_CHILD | WS_VISIBLE, Scale(16), Scale(54), Scale(110), Scale(20), authWnd_, nullptr, instance_, nullptr);
        authUserCombo_ = CreateWindowW(
            WC_COMBOBOXW,
            L"",
            WS_CHILD | WS_VISIBLE | CBS_DROPDOWNLIST | WS_VSCROLL,
            Scale(140), Scale(50), Scale(320), Scale(260),
            authWnd_,
            reinterpret_cast<HMENU>(static_cast<INT_PTR>(ID_AUTH_USER_COMBO)),
            instance_,
            nullptr
        );
        CreateWindowW(L"STATIC", L"Пароль:", WS_CHILD | WS_VISIBLE, Scale(16), Scale(98), Scale(110), Scale(20), authWnd_, nullptr, instance_, nullptr);
        authPassword_ = CreateWindowExW(WS_EX_CLIENTEDGE, L"EDIT", L"", WS_CHILD | WS_VISIBLE | ES_PASSWORD | ES_AUTOHSCROLL,
            Scale(140), Scale(94), Scale(320), Scale(26), authWnd_, reinterpret_cast<HMENU>(static_cast<INT_PTR>(ID_AUTH_PASSWORD)), instance_, nullptr);
        authHint_ = CreateWindowW(L"STATIC", L"Сначала программа подключается к серверу, затем здесь выбирается пользователь как на сайте.", WS_CHILD | WS_VISIBLE,
            Scale(16), Scale(138), Scale(444), Scale(42), authWnd_, nullptr, instance_, nullptr);
        authLoginBtn_ = CreateWindowW(L"BUTTON", L"Войти", WS_CHILD | WS_VISIBLE | BS_FLAT, Scale(140), Scale(200), Scale(140), Scale(34), authWnd_,
            reinterpret_cast<HMENU>(static_cast<INT_PTR>(ID_AUTH_LOGIN)), instance_, nullptr);
        authCancelBtn_ = CreateWindowW(L"BUTTON", L"Отмена", WS_CHILD | WS_VISIBLE | BS_FLAT, Scale(296), Scale(200), Scale(140), Scale(34), authWnd_,
            reinterpret_cast<HMENU>(static_cast<INT_PTR>(ID_AUTH_CANCEL)), instance_, nullptr);

        ApplyControlFont(authTitle_);
        ApplyControlFont(authUserCombo_);
        ApplyControlFont(authPassword_);
        ApplyControlFont(authHint_);
        ApplyControlFont(authLoginBtn_);
        ApplyControlFont(authCancelBtn_);
    }

    void ApplyControlFont(HWND control) const {
        if (control && uiFont_) {
            SendMessageW(control, WM_SETFONT, reinterpret_cast<WPARAM>(uiFont_), TRUE);
        }
    }

    void RemoveTrayIcon() {
        NOTIFYICONDATAW nid{};
        nid.cbSize = sizeof(nid);
        nid.hWnd = hwnd_;
        nid.uID = 1;
        Shell_NotifyIconW(NIM_DELETE, &nid);
    }

    void ShowTrayBalloon(const std::wstring& title, const std::wstring& text) {
        NOTIFYICONDATAW nid{};
        nid.cbSize = sizeof(nid);
        nid.hWnd = hwnd_;
        nid.uID = 1;
        nid.uFlags = NIF_INFO;
        nid.dwInfoFlags = NIIF_INFO;
        wcsncpy_s(nid.szInfoTitle, title.c_str(), _TRUNCATE);
        wcsncpy_s(nid.szInfo, text.c_str(), _TRUNCATE);
        Shell_NotifyIconW(NIM_MODIFY, &nid);
    }

    void UpdateTransportStatus(const std::wstring& mode, bool notify) {
        std::wstring label = L"Режим: " + mode;
        if (transportLabel_) {
            SetWindowTextW(transportLabel_, label.c_str());
        }
        if (lastTransportLabel_ != label) {
            lastTransportLabel_ = label;
            if (notify) {
                ShowTrayBalloon(L"OTD421 Chat Tray", label);
            }
        }
    }

    std::wstring DatabasePath() const {
        wchar_t exePath[MAX_PATH] = {};
        GetModuleFileNameW(nullptr, exePath, MAX_PATH);
        std::wstring path = exePath;
        auto slash = path.find_last_of(L"\\/");
        return path.substr(0, slash + 1) + L"chat_client.db";
    }

    bool OpenLocalDatabase(sqlite3** db) const {
        if (!db) return false;
        *db = nullptr;
        return sqlite3_open16(DatabasePath().c_str(), db) == SQLITE_OK && *db != nullptr;
    }

    void EnsureLocalDatabase() const {
        sqlite3* db = nullptr;
        if (!OpenLocalDatabase(&db)) {
            if (db) sqlite3_close(db);
            return;
        }

        const char* schema =
            "CREATE TABLE IF NOT EXISTS client_meta ("
            "  key TEXT PRIMARY KEY,"
            "  value TEXT NOT NULL"
            ");"
            "CREATE TABLE IF NOT EXISTS chat_history ("
            "  id INTEGER PRIMARY KEY AUTOINCREMENT,"
            "  created_at TEXT NOT NULL,"
            "  direction TEXT NOT NULL,"
            "  remote_message_id INTEGER,"
            "  sender TEXT NOT NULL,"
            "  receiver TEXT NOT NULL,"
            "  body TEXT NOT NULL"
            ");"
            "CREATE UNIQUE INDEX IF NOT EXISTS idx_chat_history_remote_message_id "
            "ON chat_history(remote_message_id) "
            "WHERE remote_message_id IS NOT NULL;";
        sqlite3_exec(db, schema, nullptr, nullptr, nullptr);
        sqlite3_close(db);
    }

    int ReadMetaInt(const char* key, int fallback) const {
        sqlite3* db = nullptr;
        if (!OpenLocalDatabase(&db)) {
            if (db) sqlite3_close(db);
            return fallback;
        }

        sqlite3_stmt* stmt = nullptr;
        int value = fallback;
        if (sqlite3_prepare_v2(db, "SELECT value FROM client_meta WHERE key = ?1", -1, &stmt, nullptr) == SQLITE_OK) {
            sqlite3_bind_text(stmt, 1, key, -1, SQLITE_STATIC);
            if (sqlite3_step(stmt) == SQLITE_ROW) {
                const unsigned char* raw = sqlite3_column_text(stmt, 0);
                if (raw) {
                    value = atoi(reinterpret_cast<const char*>(raw));
                }
            }
        }
        if (stmt) sqlite3_finalize(stmt);
        sqlite3_close(db);
        return value;
    }

    void WriteMetaInt(const char* key, int value) const {
        sqlite3* db = nullptr;
        if (!OpenLocalDatabase(&db)) {
            if (db) sqlite3_close(db);
            return;
        }

        sqlite3_stmt* stmt = nullptr;
        if (sqlite3_prepare_v2(db, "REPLACE INTO client_meta(key, value) VALUES(?1, ?2)", -1, &stmt, nullptr) == SQLITE_OK) {
            sqlite3_bind_text(stmt, 1, key, -1, SQLITE_STATIC);
            std::string numeric = std::to_string(value);
            sqlite3_bind_text(stmt, 2, numeric.c_str(), -1, SQLITE_TRANSIENT);
            sqlite3_step(stmt);
        }
        if (stmt) sqlite3_finalize(stmt);
        sqlite3_close(db);
    }

    void LoadChatState() {
        lastDirectId_ = ReadMetaInt("last_direct_id", 0);
        lastGeneralId_ = ReadMetaInt("last_general_id", 0);
        if (lastDirectId_ < 0) lastDirectId_ = 0;
        if (lastGeneralId_ < 0) lastGeneralId_ = 0;
    }

    void SaveChatState() const {
        WriteMetaInt("last_direct_id", lastDirectId_);
        WriteMetaInt("last_general_id", lastGeneralId_);
    }

    std::wstring CurrentTimestamp() const {
        SYSTEMTIME st{};
        GetLocalTime(&st);
        wchar_t buffer[64] = {};
        swprintf_s(buffer, L"%04d-%02d-%02d %02d:%02d:%02d",
            st.wYear, st.wMonth, st.wDay, st.wHour, st.wMinute, st.wSecond);
        return buffer;
    }

    void AppendHistoryEntry(const std::wstring& direction, int remoteMessageId, const std::wstring& sender, const std::wstring& receiver, const std::wstring& text) const {
        sqlite3* db = nullptr;
        if (!OpenLocalDatabase(&db)) {
            if (db) sqlite3_close(db);
            return;
        }

        sqlite3_stmt* stmt = nullptr;
        if (sqlite3_prepare_v2(
                db,
                "INSERT INTO chat_history(created_at, direction, remote_message_id, sender, receiver, body) "
                "VALUES(?1, ?2, ?3, ?4, ?5, ?6)",
                -1,
                &stmt,
                nullptr
            ) == SQLITE_OK) {
            std::string createdAt = WideToUtf8(CurrentTimestamp());
            std::string directionUtf8 = WideToUtf8(direction);
            std::string senderUtf8 = WideToUtf8(sender);
            std::string receiverUtf8 = WideToUtf8(receiver);
            std::string textUtf8 = WideToUtf8(text);
            sqlite3_bind_text(stmt, 1, createdAt.c_str(), -1, SQLITE_TRANSIENT);
            sqlite3_bind_text(stmt, 2, directionUtf8.c_str(), -1, SQLITE_TRANSIENT);
            if (remoteMessageId > 0) {
                sqlite3_bind_int(stmt, 3, remoteMessageId);
            } else {
                sqlite3_bind_null(stmt, 3);
            }
            sqlite3_bind_text(stmt, 4, senderUtf8.c_str(), -1, SQLITE_TRANSIENT);
            sqlite3_bind_text(stmt, 5, receiverUtf8.c_str(), -1, SQLITE_TRANSIENT);
            sqlite3_bind_text(stmt, 6, textUtf8.c_str(), -1, SQLITE_TRANSIENT);
            sqlite3_step(stmt);
        }
        if (stmt) sqlite3_finalize(stmt);
        sqlite3_close(db);
    }

    bool HasHistoryMessageId(int remoteMessageId) const {
        if (remoteMessageId <= 0) return false;
        sqlite3* db = nullptr;
        if (!OpenLocalDatabase(&db)) {
            if (db) sqlite3_close(db);
            return false;
        }

        sqlite3_stmt* stmt = nullptr;
        bool exists = false;
        if (sqlite3_prepare_v2(db, "SELECT 1 FROM chat_history WHERE remote_message_id = ?1 LIMIT 1", -1, &stmt, nullptr) == SQLITE_OK) {
            sqlite3_bind_int(stmt, 1, remoteMessageId);
            exists = sqlite3_step(stmt) == SQLITE_ROW;
        }
        if (stmt) sqlite3_finalize(stmt);
        sqlite3_close(db);
        return exists;
    }

    void StartWorker() {
        running_ = true;
        worker_ = std::thread([this]() {
            while (running_) {
                if (!serverConnected_) {
                    std::vector<LoginChoice> fetchedLogins;
                    if (!client_->FetchLoginUsers(fetchedLogins)) {
                        PostMessageW(hwnd_, WMAPP_STATUS, 0, 0);
                        std::this_thread::sleep_for(std::chrono::seconds(5));
                        continue;
                    }
                    {
                        std::lock_guard<std::mutex> lock(loginChoicesMutex_);
                        loginChoices_ = std::move(fetchedLogins);
                    }
                    serverConnected_ = true;
                    PostMessageW(hwnd_, WMAPP_STATUS, 10, 0);
                }

                if (!loggedIn_) {
                    websocketConnected_ = false;
                    client_->CloseTrayWebSocket();
                    UpdateTransportStatus(L"ожидание входа", false);
                    if (config_.username.empty() || config_.password.empty()) {
                        if (!authPromptOpen_) {
                            PostMessageW(hwnd_, WMAPP_AUTH_REQUIRED, 0, 0);
                        }
                        std::this_thread::sleep_for(std::chrono::seconds(2));
                        continue;
                    }

                    loggedIn_ = client_->Login();
                    if (loggedIn_) {
                        authPromptOpen_ = false;
                        std::vector<ChatUser> fetched;
                        if (client_->FetchUsers(fetched)) {
                            {
                                std::lock_guard<std::mutex> lock(usersMutex_);
                                users_ = std::move(fetched);
                            }
                            PostMessageW(hwnd_, WMAPP_STATUS, 2, 0);
                        }
                        PostMessageW(hwnd_, WMAPP_STATUS, 1, 0);
                    } else {
                        config_.password.clear();
                        authPromptOpen_ = false;
                        PostMessageW(hwnd_, WMAPP_AUTH_REQUIRED, 0, 0);
                        std::this_thread::sleep_for(std::chrono::seconds(2));
                        continue;
                    }
                }

                std::vector<TrayMessage> batch;
                if (!websocketConnected_) {
                    websocketConnected_ = client_->OpenTrayWebSocket(lastDirectId_, lastGeneralId_);
                    if (websocketConnected_) {
                        UpdateTransportStatus(L"WebSocket", true);
                    }
                }

                if (websocketConnected_) {
                    bool authRequired = false;
                    if (!client_->ReceiveTrayWebSocket(lastDirectId_, lastGeneralId_, batch, authRequired)) {
                        websocketConnected_ = false;
                        UpdateTransportStatus(L"polling", true);
                        if (authRequired) {
                            loggedIn_ = false;
                            std::this_thread::sleep_for(std::chrono::seconds(2));
                        } else {
                            std::this_thread::sleep_for(std::chrono::seconds(2));
                        }
                        continue;
                    }
                } else {
                    UpdateTransportStatus(L"polling", false);
                    if (!client_->PollIncoming(lastDirectId_, lastGeneralId_, batch)) {
                        loggedIn_ = false;
                        std::this_thread::sleep_for(std::chrono::seconds(5));
                        continue;
                    }
                    for (int i = 0; i < config_.pollSeconds && running_; ++i) {
                        std::this_thread::sleep_for(std::chrono::seconds(1));
                    }
                }
                SaveChatState();

                if (!batch.empty()) {
                    std::lock_guard<std::mutex> lock(queueMutex_);
                    for (auto& msg : batch) {
                        if (HasHistoryMessageId(msg.id)) {
                            continue;
                        }
                        AppendHistoryEntry(
                            msg.kind == L"general" ? L"IN-GENERAL" : L"IN-DIRECT",
                            msg.id,
                            msg.sender,
                            msg.kind == L"general" ? L"Общий чат" : L"Мне",
                            msg.text
                        );
                        pending_.push(std::move(msg));
                    }
                    PostMessageW(hwnd_, WMAPP_NEW_MESSAGE, 0, 0);
                }
            }
        });
    }

    void PopulateAuthChoices() {
        if (!authUserCombo_) return;
        std::wstring selectedUsername = config_.username;
        int currentIndex = static_cast<int>(SendMessageW(authUserCombo_, CB_GETCURSEL, 0, 0));
        if (currentIndex != CB_ERR) {
            std::lock_guard<std::mutex> lock(loginChoicesMutex_);
            if (currentIndex >= 0 && currentIndex < static_cast<int>(loginChoices_.size())) {
                selectedUsername = loginChoices_[currentIndex].username;
            }
        }
        SendMessageW(authUserCombo_, CB_RESETCONTENT, 0, 0);

        std::lock_guard<std::mutex> lock(loginChoicesMutex_);
        int selectedIndex = 0;
        for (size_t i = 0; i < loginChoices_.size(); ++i) {
            std::wstring label = loginChoices_[i].displayName;
            int idx = static_cast<int>(SendMessageW(authUserCombo_, CB_ADDSTRING, 0, reinterpret_cast<LPARAM>(label.c_str())));
            if (loginChoices_[i].username == selectedUsername) {
                selectedIndex = idx;
            }
        }
        if (SendMessageW(authUserCombo_, CB_GETCOUNT, 0, 0) > 0) {
            SendMessageW(authUserCombo_, CB_SETCURSEL, selectedIndex, 0);
        }
    }

    void ShowAuthWindow() {
        if (authPromptOpen_) {
            SetForegroundWindow(authWnd_);
            return;
        }
        authPromptOpen_ = true;
        PopulateAuthChoices();
        SetWindowTextW(authPassword_, L"");
        ShowWindow(authWnd_, SW_SHOWNORMAL);
        SetForegroundWindow(authWnd_);
    }

    void SaveAuthAndReconnect() {
        int index = static_cast<int>(SendMessageW(authUserCombo_, CB_GETCURSEL, 0, 0));
        if (index == CB_ERR) {
            MessageBoxW(authWnd_, L"Выберите пользователя.", L"OTD421 Chat Tray", MB_ICONWARNING);
            return;
        }

        std::wstring password = ReadWindowText(authPassword_);
        if (password.empty()) {
            MessageBoxW(authWnd_, L"Введите пароль.", L"OTD421 Chat Tray", MB_ICONWARNING);
            return;
        }

        {
            std::lock_guard<std::mutex> lock(loginChoicesMutex_);
            if (index < 0 || index >= static_cast<int>(loginChoices_.size())) {
                MessageBoxW(authWnd_, L"Список пользователей устарел. Попробуйте ещё раз.", L"OTD421 Chat Tray", MB_ICONWARNING);
                return;
            }
            config_.username = loginChoices_[index].username;
        }
        config_.password = password;

        std::wstring path = ConfigPath();
        WriteIniString(path, L"auth", L"username", config_.username);
        WriteIniString(path, L"auth", L"password", config_.password);

        loggedIn_ = false;
        websocketConnected_ = false;
        client_->CloseTrayWebSocket();
        UpdateTransportStatus(L"переподключение", false);
        authPromptOpen_ = false;
        ShowWindow(authWnd_, SW_HIDE);
        ShowTrayBalloon(L"OTD421 Chat Tray", L"Выполняется вход в чат.");
    }

    void ShowNextMessage() {
        std::lock_guard<std::mutex> lock(queueMutex_);
        if (pending_.empty()) return;
        current_ = pending_.front();
        pending_.pop();

        std::wstring kind = current_.kind == L"general" ? L"Общий чат" : L"Личное сообщение";
        SetWindowTextW(senderLabel_, (L"От: " + current_.sender + L"   " + current_.time).c_str());
        SetWindowTextW(kindLabel_, kind.c_str());
        SetWindowTextW(textView_, current_.text.c_str());
        SetWindowTextW(replyEdit_, L"");
        if (current_.kind == L"general") {
            SelectRecipient(L"all");
        } else if (current_.senderId > 0) {
            SelectRecipient(std::to_wstring(current_.senderId));
        }
        PositionPopup();
        ShowWindow(popup_, SW_SHOWNORMAL);
        SetForegroundWindow(popup_);

        std::wstring balloonTitle = kind == L"Общий чат" ? L"Новое сообщение в общий чат" : L"Новое личное сообщение";
        ShowTrayBalloon(balloonTitle, current_.sender + L": " + current_.text);
    }

    void PositionPopup() {
        RECT work{};
        SystemParametersInfoW(SPI_GETWORKAREA, 0, &work, 0);
        int width = Scale(460);
        int height = Scale(370);
        int x = work.right - width - Scale(16);
        int y = work.bottom - height - Scale(16);
        SetWindowPos(popup_, HWND_TOPMOST, x, y, width, height, SWP_SHOWWINDOW);
    }

    void ShowTrayMenu() {
        HMENU menu = CreatePopupMenu();
        AppendMenuW(menu, MF_STRING, ID_TRAY_OPEN, L"Показать окно");
        AppendMenuW(menu, MF_STRING, ID_TRAY_WEBCHAT, L"Открыть веб-чат");
        AppendMenuW(menu, MF_STRING, ID_TRAY_RECONNECT, L"Переподключиться");
        AppendMenuW(menu, MF_STRING, ID_TRAY_SETTINGS, L"Настройки");
        AppendMenuW(menu, MF_SEPARATOR, 0, nullptr);
        AppendMenuW(menu, MF_STRING, ID_TRAY_EXIT, L"Выход");

        POINT pt{};
        GetCursorPos(&pt);
        SetForegroundWindow(hwnd_);
        TrackPopupMenu(menu, TPM_BOTTOMALIGN | TPM_LEFTALIGN, pt.x, pt.y, 0, hwnd_, nullptr);
        DestroyMenu(menu);
    }

    void OpenWebChat() const {
        std::wstring target = client_->BaseUrl() + L"/chat";
        ShellExecuteW(nullptr, L"open", target.c_str(), nullptr, nullptr, SW_SHOWNORMAL);
    }

    void OpenCurrentConversation() const {
        std::wstring target = client_->BaseUrl() + L"/chat?user_id=" + GetSelectedRecipientId();
        ShellExecuteW(nullptr, L"open", target.c_str(), nullptr, nullptr, SW_SHOWNORMAL);
    }

    void SendReply() {
        int len = GetWindowTextLengthW(replyEdit_);
        if (len <= 0) return;
        std::wstring text(len + 1, L'\0');
        GetWindowTextW(replyEdit_, text.data(), len + 1);
        text.resize(wcslen(text.c_str()));
        std::wstring receiver = GetSelectedRecipientId();
        std::wstring receiverLabel = receiver == L"all" ? L"Общий чат" : receiver;
        if (receiver != L"all") {
            std::lock_guard<std::mutex> lock(usersMutex_);
            for (const auto& user : users_) {
                if (std::to_wstring(user.id) == receiver) {
                    receiverLabel = user.name;
                    break;
                }
            }
        }
        SendChatResult result = client_->SendChatMessage(receiver, text);
        if (result == SendChatResult::Success) {
            SetWindowTextW(replyEdit_, L"");
            ShowWindow(popup_, SW_HIDE);
            AppendHistoryEntry(receiver == L"all" ? L"OUT-GENERAL" : L"OUT-DIRECT", 0, L"Я", receiverLabel, text);
            ShowTrayBalloon(L"Ответ отправлен", text);
        } else if (result == SendChatResult::AuthRequired) {
            loggedIn_ = false;
            websocketConnected_ = false;
            client_->CloseTrayWebSocket();
            authPromptOpen_ = false;
            config_.password.clear();
            ShowTrayBalloon(L"OTD421 Chat Tray", L"Сессия истекла. Войдите заново.");
            ShowAuthWindow();
        } else {
            MessageBoxW(popup_, L"Не удалось отправить ответ.", L"OTD421 Chat Tray", MB_ICONERROR);
        }
    }

    Config LoadConfig() const {
        std::wstring path = ConfigPath();
        Config cfg;
        cfg.baseUrl = ReadIniString(path, L"server", L"base_url", cfg.baseUrl.c_str());
        cfg.username = ReadIniString(path, L"auth", L"username", L"");
        cfg.password = ReadIniString(path, L"auth", L"password", L"");
        cfg.pollSeconds = _wtoi(ReadIniString(path, L"server", L"poll_seconds", L"30").c_str());
        cfg.ignoreSslErrors = ReadIniString(path, L"server", L"ignore_ssl_errors", L"1") != L"0";
        if (cfg.pollSeconds < 30) cfg.pollSeconds = 30;
        return cfg;
    }

    std::wstring ConfigPath() const {
        wchar_t exePath[MAX_PATH] = {};
        GetModuleFileNameW(nullptr, exePath, MAX_PATH);
        std::wstring path = exePath;
        auto slash = path.find_last_of(L"\\/");
        return path.substr(0, slash + 1) + L"config.ini";
    }

    std::wstring ReadWindowText(HWND control) const {
        int len = GetWindowTextLengthW(control);
        std::wstring value(len + 1, L'\0');
        GetWindowTextW(control, value.data(), len + 1);
        value.resize(wcslen(value.c_str()));
        return value;
    }

    void ShowSettingsWindow() {
        if (!settingsWnd_) return;
        SetWindowTextW(settingsBaseUrl_, config_.baseUrl.c_str());
        SetWindowTextW(settingsUsername_, config_.username.c_str());
        SetWindowTextW(settingsPassword_, config_.password.c_str());
        SetWindowTextW(settingsPoll_, std::to_wstring(config_.pollSeconds).c_str());
        SendMessageW(settingsIgnoreSsl_, BM_SETCHECK, config_.ignoreSslErrors ? BST_CHECKED : BST_UNCHECKED, 0);
        ShowWindow(settingsWnd_, SW_SHOWNORMAL);
        SetForegroundWindow(settingsWnd_);
    }

    void SaveSettings() {
        Config newConfig = config_;
        newConfig.baseUrl = ReadWindowText(settingsBaseUrl_);
        newConfig.pollSeconds = _wtoi(ReadWindowText(settingsPoll_).c_str());
        newConfig.ignoreSslErrors = SendMessageW(settingsIgnoreSsl_, BM_GETCHECK, 0, 0) == BST_CHECKED;

        if (newConfig.baseUrl.empty()) {
            MessageBoxW(settingsWnd_, L"Заполните адрес сервера.", L"OTD421 Chat Tray", MB_ICONWARNING);
            return;
        }
        if (newConfig.pollSeconds < 30) {
            newConfig.pollSeconds = 30;
        }

        std::wstring path = ConfigPath();
        WriteIniString(path, L"server", L"base_url", newConfig.baseUrl);
        WriteIniString(path, L"server", L"poll_seconds", std::to_wstring(newConfig.pollSeconds));
        WriteIniString(path, L"server", L"ignore_ssl_errors", newConfig.ignoreSslErrors ? L"1" : L"0");

        bool serverChanged = config_.baseUrl != newConfig.baseUrl;
        config_ = std::move(newConfig);
        config_.username.clear();
        config_.password.clear();
        client_ = std::make_unique<HttpClient>(config_);
        loggedIn_ = false;
        websocketConnected_ = false;
        UpdateTransportStatus(L"переподключение", false);
        serverConnected_ = false;
        if (serverChanged) {
            lastDirectId_ = 0;
            lastGeneralId_ = 0;
            SaveChatState();
        }
        ShowWindow(settingsWnd_, SW_HIDE);
        ShowTrayBalloon(L"OTD421 Chat Tray", L"Настройки сервера сохранены. Выполняется подключение.");
    }

    void PopulateRecipients() {
        if (!recipientCombo_) return;

        std::wstring selected = GetSelectedRecipientId();
        SendMessageW(recipientCombo_, CB_RESETCONTENT, 0, 0);

        int index = static_cast<int>(SendMessageW(recipientCombo_, CB_ADDSTRING, 0, reinterpret_cast<LPARAM>(L"Общий чат")));
        SendMessageW(recipientCombo_, CB_SETITEMDATA, index, static_cast<LPARAM>(0));

        std::lock_guard<std::mutex> lock(usersMutex_);
        for (const auto& user : users_) {
            std::wstring label = user.name;
            if (!user.role.empty()) {
                label += L" (" + user.role + L")";
            }
            int userIndex = static_cast<int>(SendMessageW(recipientCombo_, CB_ADDSTRING, 0, reinterpret_cast<LPARAM>(label.c_str())));
            SendMessageW(recipientCombo_, CB_SETITEMDATA, userIndex, static_cast<LPARAM>(user.id));
        }

        if (!selected.empty()) {
            SelectRecipient(selected);
        } else {
            SendMessageW(recipientCombo_, CB_SETCURSEL, 0, 0);
        }
    }

    void SelectRecipient(const std::wstring& value) {
        if (!recipientCombo_) return;

        int count = static_cast<int>(SendMessageW(recipientCombo_, CB_GETCOUNT, 0, 0));
        for (int i = 0; i < count; ++i) {
            LPARAM itemData = SendMessageW(recipientCombo_, CB_GETITEMDATA, i, 0);
            std::wstring currentValue = itemData == 0 ? L"all" : std::to_wstring(static_cast<int>(itemData));
            if (currentValue == value) {
                SendMessageW(recipientCombo_, CB_SETCURSEL, i, 0);
                return;
            }
        }
        SendMessageW(recipientCombo_, CB_SETCURSEL, 0, 0);
    }

    std::wstring GetSelectedRecipientId() const {
        if (!recipientCombo_) return L"all";
        int index = static_cast<int>(SendMessageW(recipientCombo_, CB_GETCURSEL, 0, 0));
        if (index == CB_ERR) return L"all";
        LPARAM itemData = SendMessageW(recipientCombo_, CB_GETITEMDATA, index, 0);
        if (itemData <= 0) return L"all";
        return std::to_wstring(static_cast<int>(itemData));
    }

    LRESULT HandleMain(HWND hwnd, UINT msg, WPARAM wp, LPARAM lp) {
        switch (msg) {
        case WMAPP_TRAY:
            if (lp == WM_RBUTTONUP || lp == WM_CONTEXTMENU) {
                ShowTrayMenu();
            } else if (lp == WM_LBUTTONDBLCLK) {
                PositionPopup();
                ShowWindow(popup_, SW_SHOWNORMAL);
                SetForegroundWindow(popup_);
            }
            return 0;
        case WM_COMMAND:
            switch (LOWORD(wp)) {
            case ID_TRAY_OPEN:
                PositionPopup();
                ShowWindow(popup_, SW_SHOWNORMAL);
                SetForegroundWindow(popup_);
                return 0;
            case ID_TRAY_WEBCHAT:
                OpenWebChat();
                return 0;
            case ID_TRAY_RECONNECT:
                loggedIn_ = false;
                websocketConnected_ = false;
                client_->CloseTrayWebSocket();
                UpdateTransportStatus(L"переподключение", true);
                serverConnected_ = false;
                config_.password.clear();
                authPromptOpen_ = false;
                return 0;
            case ID_TRAY_SETTINGS:
                ShowSettingsWindow();
                return 0;
            case ID_TRAY_EXIT:
                DestroyWindow(hwnd_);
                return 0;
            default:
                break;
            }
            break;
        case WMAPP_NEW_MESSAGE:
            ShowNextMessage();
            return 0;
        case WMAPP_STATUS:
            if (wp == 1) {
                ShowTrayBalloon(L"OTD421 Chat Tray", L"Подключение к чату выполнено.");
            } else if (wp == 2) {
                PopulateRecipients();
            } else if (wp == 10) {
                if (config_.username.empty() || config_.password.empty()) {
                    ShowTrayBalloon(L"OTD421 Chat Tray", L"Сервер доступен. Выберите пользователя для входа.");
                    ShowAuthWindow();
                } else {
                    ShowTrayBalloon(L"OTD421 Chat Tray", L"Сервер доступен. Выполняется автоматический вход.");
                }
            } else {
                ShowTrayBalloon(L"OTD421 Chat Tray", L"Не удалось подключиться к серверу. Проверьте адрес и сеть.");
            }
            return 0;
        case WMAPP_AUTH_REQUIRED:
            ShowAuthWindow();
            return 0;
        case WM_DESTROY:
            PostQuitMessage(0);
            return 0;
        default:
            return DefWindowProcW(hwnd, msg, wp, lp);
        }
        return DefWindowProcW(hwnd, msg, wp, lp);
    }

    LRESULT HandlePopup(HWND hwnd, UINT msg, WPARAM wp, LPARAM lp) {
        switch (msg) {
        case WM_CTLCOLOREDIT:
        case WM_CTLCOLORSTATIC:
        case WM_CTLCOLORLISTBOX: {
            HDC hdc = reinterpret_cast<HDC>(wp);
            SetTextColor(hdc, kTextColor);
            SetBkColor(hdc, kPanelColor);
            return reinterpret_cast<LRESULT>(panelBrush_);
        }
        case WM_CTLCOLORBTN: {
            HDC hdc = reinterpret_cast<HDC>(wp);
            SetTextColor(hdc, kAccentColor);
            SetBkColor(hdc, kBgColor);
            return reinterpret_cast<LRESULT>(bgBrush_);
        }
        case WM_COMMAND:
            switch (LOWORD(wp)) {
            case ID_SEND_BTN:
                SendReply();
                return 0;
            case ID_OPEN_BTN:
                OpenCurrentConversation();
                return 0;
            case ID_CLOSE_BTN:
                ShowWindow(hwnd, SW_HIDE);
                return 0;
            default:
                break;
            }
            break;
        case WM_CLOSE:
            ShowWindow(hwnd, SW_HIDE);
            return 0;
        default:
            return DefWindowProcW(hwnd, msg, wp, lp);
        }
        return DefWindowProcW(hwnd, msg, wp, lp);
    }

    LRESULT HandleSettings(HWND hwnd, UINT msg, WPARAM wp, LPARAM lp) {
        switch (msg) {
        case WM_CTLCOLOREDIT:
        case WM_CTLCOLORSTATIC:
        case WM_CTLCOLORLISTBOX: {
            HDC hdc = reinterpret_cast<HDC>(wp);
            SetTextColor(hdc, kTextColor);
            SetBkColor(hdc, kPanelColor);
            return reinterpret_cast<LRESULT>(panelBrush_);
        }
        case WM_CTLCOLORBTN: {
            HDC hdc = reinterpret_cast<HDC>(wp);
            SetTextColor(hdc, kAccentColor);
            SetBkColor(hdc, kBgColor);
            return reinterpret_cast<LRESULT>(bgBrush_);
        }
        case WM_COMMAND:
            switch (LOWORD(wp)) {
            case ID_SETTINGS_SAVE:
                SaveSettings();
                return 0;
            case ID_SETTINGS_CANCEL:
                ShowWindow(hwnd, SW_HIDE);
                return 0;
            default:
                break;
            }
            break;
        case WM_CLOSE:
            ShowWindow(hwnd, SW_HIDE);
            return 0;
        default:
            return DefWindowProcW(hwnd, msg, wp, lp);
        }
        return DefWindowProcW(hwnd, msg, wp, lp);
    }

    LRESULT HandleAuth(HWND hwnd, UINT msg, WPARAM wp, LPARAM lp) {
        switch (msg) {
        case WM_CTLCOLOREDIT:
        case WM_CTLCOLORSTATIC:
        case WM_CTLCOLORLISTBOX: {
            HDC hdc = reinterpret_cast<HDC>(wp);
            SetTextColor(hdc, kTextColor);
            SetBkColor(hdc, kPanelColor);
            return reinterpret_cast<LRESULT>(panelBrush_);
        }
        case WM_CTLCOLORBTN: {
            HDC hdc = reinterpret_cast<HDC>(wp);
            SetTextColor(hdc, kAccentColor);
            SetBkColor(hdc, kBgColor);
            return reinterpret_cast<LRESULT>(bgBrush_);
        }
        case WM_COMMAND:
            switch (LOWORD(wp)) {
            case ID_AUTH_LOGIN:
                SaveAuthAndReconnect();
                return 0;
            case ID_AUTH_CANCEL:
                authPromptOpen_ = false;
                ShowWindow(hwnd, SW_HIDE);
                return 0;
            default:
                break;
            }
            break;
        case WM_CLOSE:
            ShowWindow(hwnd, SW_HIDE);
            return 0;
        default:
            return DefWindowProcW(hwnd, msg, wp, lp);
        }
        return DefWindowProcW(hwnd, msg, wp, lp);
    }

    HINSTANCE instance_ = nullptr;
    HWND hwnd_ = nullptr;
    HWND popup_ = nullptr;
    HWND settingsWnd_ = nullptr;
    HWND authWnd_ = nullptr;
    HWND senderLabel_ = nullptr;
    HWND kindLabel_ = nullptr;
    HWND transportLabel_ = nullptr;
    HWND recipientLabel_ = nullptr;
    HWND recipientCombo_ = nullptr;
    HWND textView_ = nullptr;
    HWND replyEdit_ = nullptr;
    HWND sendBtn_ = nullptr;
    HWND openBtn_ = nullptr;
    HWND closeBtn_ = nullptr;
    HWND settingsBaseUrl_ = nullptr;
    HWND settingsUsername_ = nullptr;
    HWND settingsPassword_ = nullptr;
    HWND settingsPoll_ = nullptr;
    HWND settingsIgnoreSsl_ = nullptr;
    HWND settingsTitle_ = nullptr;
    HWND settingsHint_ = nullptr;
    HWND settingsSaveBtn_ = nullptr;
    HWND settingsCancelBtn_ = nullptr;
    HWND authTitle_ = nullptr;
    HWND authUserCombo_ = nullptr;
    HWND authPassword_ = nullptr;
    HWND authHint_ = nullptr;
    HWND authLoginBtn_ = nullptr;
    HWND authCancelBtn_ = nullptr;
    Config config_;
    std::unique_ptr<HttpClient> client_;
    std::thread worker_;
    std::atomic<bool> running_{ false };
    std::atomic<bool> loggedIn_{ false };
    std::atomic<bool> serverConnected_{ false };
    std::atomic<bool> authPromptOpen_{ false };
    std::atomic<bool> websocketConnected_{ false };
    int lastDirectId_ = 0;
    int lastGeneralId_ = 0;
    TrayMessage current_;
    std::vector<ChatUser> users_;
    std::vector<LoginChoice> loginChoices_;
    std::mutex usersMutex_;
    std::mutex loginChoicesMutex_;
    std::mutex queueMutex_;
    std::queue<TrayMessage> pending_;
    HBRUSH bgBrush_ = nullptr;
    HBRUSH panelBrush_ = nullptr;
    HBRUSH accentBrush_ = nullptr;
    HFONT uiFont_ = nullptr;
    UINT dpi_ = 96;
    std::wstring lastTransportLabel_;
};
} // namespace

int APIENTRY wWinMain(HINSTANCE instance, HINSTANCE, PWSTR, int) {
    TrayApp app(instance);
    return app.Run();
}

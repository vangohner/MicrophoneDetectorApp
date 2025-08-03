#define _CRT_SECURE_NO_WARNINGS
#define NOMINMAX
#define WIN32_LEAN_AND_MEAN

#include <windows.h>
#include <shellapi.h>
#include <commctrl.h>
#include <mmdeviceapi.h>
#include <audiopolicy.h>
#include <iostream>
#include <vector>
#include <string>
#include <thread>
#include <chrono>
#include <algorithm>
#include <sstream>
#include <mutex>
#include <iomanip>
#include <atomic>

// Windows BLE headers
#include <winrt/base.h>
#include <winrt/Windows.Foundation.h>
#include <winrt/Windows.Foundation.Collections.h>
#include <winrt/Windows.Devices.Bluetooth.h>
#include <winrt/Windows.Devices.Bluetooth.Advertisement.h>
#include <winrt/Windows.Devices.Bluetooth.GenericAttributeProfile.h>
#include <winrt/Windows.Storage.Streams.h>

using namespace winrt;
using namespace Windows::Foundation;
using namespace Windows::Devices::Bluetooth;
using namespace Windows::Devices::Bluetooth::Advertisement;
using namespace Windows::Devices::Bluetooth::GenericAttributeProfile;
using namespace Windows::Storage::Streams;

// Link required libraries
#pragma comment(lib, "ole32.lib")
#pragma comment(lib, "windowsapp.lib")
#pragma comment(lib, "shell32.lib")
#pragma comment(lib, "comctl32.lib")

// Constants
#define WM_TRAYICON (WM_USER + 1)
#define ID_TRAY_SHOW_CONSOLE 1001
#define ID_TRAY_HIDE_CONSOLE 1002
#define ID_TRAY_RECONNECT 1003
#define ID_TRAY_EXIT 1004
#define ID_TRAY_ABOUT 1005

// Global variables
HWND g_hWnd = nullptr;
NOTIFYICONDATA g_nid = {};
HMENU g_hMenu = nullptr;
std::atomic<bool> g_shouldExit{ false };
std::atomic<bool> g_consoleVisible{ false };
std::mutex g_logMutex;
std::vector<std::string> g_logMessages;
const size_t MAX_LOG_MESSAGES = 100;

// Console management variables
HWND g_consoleWindow = nullptr;
bool g_consoleAllocated = false;

// Forward declarations
void HideConsole();
void ShowConsole();
void CleanupConsole();

// Logging function
void LogMessage(const std::string& message) {
    std::lock_guard<std::mutex> lock(g_logMutex);

    auto now = std::chrono::system_clock::now();
    auto time_t = std::chrono::system_clock::to_time_t(now);
    struct tm tm;
    localtime_s(&tm, &time_t);

    std::ostringstream oss;
    oss << "[" << std::put_time(&tm, "%H:%M:%S") << "] " << message;

    g_logMessages.push_back(oss.str());
    if (g_logMessages.size() > MAX_LOG_MESSAGES) {
        g_logMessages.erase(g_logMessages.begin());
    }

    if (g_consoleVisible && g_consoleWindow) {
        std::cout << oss.str() << std::endl;
    }
}

// Console management - disable close button instead of trying to handle it
void ShowConsole() {
    if (!g_consoleAllocated) {
        // First time - allocate console
        if (AllocConsole()) {
            FILE* pCout;
            freopen_s(&pCout, "CONOUT$", "w", stdout);
            SetConsoleTitle(L"Microphone LED Monitor - Console");
            g_consoleWindow = GetConsoleWindow();
            g_consoleAllocated = true;

            // Disable the close button to prevent app termination
            if (g_consoleWindow) {
                HMENU hMenu = GetSystemMenu(g_consoleWindow, FALSE);
                if (hMenu) {
                    EnableMenuItem(hMenu, SC_CLOSE, MF_BYCOMMAND | MF_DISABLED | MF_GRAYED);
                }
            }
        }
    }

    if (g_consoleWindow) {
        ShowWindow(g_consoleWindow, SW_SHOW);
        g_consoleVisible = true;
        LogMessage("Console shown (close button disabled - use tray menu to hide)");

        // Print recent log messages
        std::lock_guard<std::mutex> lock(g_logMutex);
        for (const auto& msg : g_logMessages) {
            std::cout << msg << std::endl;
        }
    }
}

void HideConsole() {
    if (g_consoleVisible && g_consoleWindow) {
        ShowWindow(g_consoleWindow, SW_HIDE);
        g_consoleVisible = false;
        LogMessage("Console hidden");
    }
}

void CleanupConsole() {
    if (g_consoleAllocated) {
        FreeConsole();
        g_consoleAllocated = false;
        g_consoleVisible = false;
        g_consoleWindow = nullptr;
    }
}

// System tray functions
void AddTrayIcon(HWND hWnd) {
    g_nid.cbSize = sizeof(NOTIFYICONDATA);
    g_nid.hWnd = hWnd;
    g_nid.uID = 1;
    g_nid.uFlags = NIF_ICON | NIF_MESSAGE | NIF_TIP;
    g_nid.uCallbackMessage = WM_TRAYICON;
    g_nid.hIcon = LoadIcon(nullptr, IDI_APPLICATION);
    wcscpy_s(g_nid.szTip, L"Microphone LED Monitor");
    Shell_NotifyIcon(NIM_ADD, &g_nid);
}

void RemoveTrayIcon() {
    Shell_NotifyIcon(NIM_DELETE, &g_nid);
}

void UpdateTrayIcon(bool connected, bool micActive) {
    std::wstring tooltip = L"Microphone LED Monitor\n";
    tooltip += connected ? L"Connected" : L"Disconnected";
    tooltip += L" | Mic: ";
    tooltip += micActive ? L"ACTIVE" : L"Inactive";

    wcscpy_s(g_nid.szTip, tooltip.c_str());
    g_nid.uFlags = NIF_TIP;
    Shell_NotifyIcon(NIM_MODIFY, &g_nid);
}

void ShowContextMenu(HWND hWnd) {
    POINT pt;
    GetCursorPos(&pt);

    if (!g_hMenu) {
        g_hMenu = CreatePopupMenu();
        AppendMenu(g_hMenu, MF_STRING, ID_TRAY_SHOW_CONSOLE, L"Show Console");
        AppendMenu(g_hMenu, MF_STRING, ID_TRAY_HIDE_CONSOLE, L"Hide Console");
        AppendMenu(g_hMenu, MF_SEPARATOR, 0, nullptr);
        AppendMenu(g_hMenu, MF_STRING, ID_TRAY_RECONNECT, L"Reconnect");
        AppendMenu(g_hMenu, MF_SEPARATOR, 0, nullptr);
        AppendMenu(g_hMenu, MF_STRING, ID_TRAY_ABOUT, L"About");
        AppendMenu(g_hMenu, MF_STRING, ID_TRAY_EXIT, L"Exit");
    }

    EnableMenuItem(g_hMenu, ID_TRAY_SHOW_CONSOLE, g_consoleVisible ? MF_GRAYED : MF_ENABLED);
    EnableMenuItem(g_hMenu, ID_TRAY_HIDE_CONSOLE, g_consoleVisible ? MF_ENABLED : MF_GRAYED);

    SetForegroundWindow(hWnd);
    TrackPopupMenu(g_hMenu, TPM_BOTTOMALIGN | TPM_LEFTALIGN, pt.x, pt.y, 0, hWnd, nullptr);
}

// Arduino BLE Controller
class ArduinoBLEController {
private:
    BluetoothLEDevice device{ nullptr };
    GattCharacteristic switchCharacteristic{ nullptr };
    GattDeviceService gattService{ nullptr };
    std::atomic<bool> isConnected{ false };
    std::atomic<bool> isConnecting{ false };
    std::chrono::steady_clock::time_point lastConnectionAttempt;
    const std::chrono::seconds RECONNECT_DELAY{ 3 };
    const std::chrono::seconds SCAN_TIMEOUT{ 8 };
    const std::wstring SWITCH_CHARACTERISTIC_UUID = L"19B10001-E8F2-537E-4F6C-D104768A1214";
    winrt::event_token connectionStatusToken{};
    std::mutex connectionMutex;

    winrt::guid parseUUID(const std::wstring& uuidStr) {
        std::wstring cleanUuid = uuidStr;
        cleanUuid.erase(std::remove(cleanUuid.begin(), cleanUuid.end(), L'-'), cleanUuid.end());

        return winrt::guid{
            static_cast<uint32_t>(std::wcstoul(cleanUuid.substr(0, 8).c_str(), nullptr, 16)),
            static_cast<uint16_t>(std::wcstoul(cleanUuid.substr(8, 4).c_str(), nullptr, 16)),
            static_cast<uint16_t>(std::wcstoul(cleanUuid.substr(12, 4).c_str(), nullptr, 16)),
            {
                static_cast<uint8_t>(std::wcstoul(cleanUuid.substr(16, 2).c_str(), nullptr, 16)),
                static_cast<uint8_t>(std::wcstoul(cleanUuid.substr(18, 2).c_str(), nullptr, 16)),
                static_cast<uint8_t>(std::wcstoul(cleanUuid.substr(20, 2).c_str(), nullptr, 16)),
                static_cast<uint8_t>(std::wcstoul(cleanUuid.substr(22, 2).c_str(), nullptr, 16)),
                static_cast<uint8_t>(std::wcstoul(cleanUuid.substr(24, 2).c_str(), nullptr, 16)),
                static_cast<uint8_t>(std::wcstoul(cleanUuid.substr(26, 2).c_str(), nullptr, 16)),
                static_cast<uint8_t>(std::wcstoul(cleanUuid.substr(28, 2).c_str(), nullptr, 16)),
                static_cast<uint8_t>(std::wcstoul(cleanUuid.substr(30, 2).c_str(), nullptr, 16))
            }
        };
    }

public:
    ArduinoBLEController() {
        lastConnectionAttempt = std::chrono::steady_clock::now() - RECONNECT_DELAY;
    }

    ~ArduinoBLEController() {
        disconnect();
    }

    bool shouldAttemptReconnect() {
        std::lock_guard<std::mutex> lock(connectionMutex);
        auto now = std::chrono::steady_clock::now();
        return !isConnected && !isConnecting && (now - lastConnectionAttempt) >= RECONNECT_DELAY;
    }

    bool getConnectionStatus() {
        std::lock_guard<std::mutex> lock(connectionMutex);
        if (!isConnected || !device) {
            return false;
        }

        try {
            bool connected = device.ConnectionStatus() == BluetoothConnectionStatus::Connected;
            if (!connected && isConnected) {
                LogMessage("Device connection status changed to disconnected");
                isConnected = false;
            }
            return connected;
        }
        catch (...) {
            LogMessage("Exception checking connection status - marking as disconnected");
            isConnected = false;
            return false;
        }
    }

    bool connectToArduino() {
        std::lock_guard<std::mutex> lock(connectionMutex);

        if (isConnected) {
            return true;
        }

        if (isConnecting) {
            return false; // Already attempting connection
        }

        lastConnectionAttempt = std::chrono::steady_clock::now();
        isConnecting = true;

        try {
            // Clean up any existing connections first
            cleanupConnection();

            LogMessage("Scanning for Arduino BLE device...");

            BluetoothLEAdvertisementWatcher watcher;
            watcher.ScanningMode(BluetoothLEScanningMode::Active);

            std::atomic<bool> deviceFound{ false };
            uint64_t targetDeviceAddress = 0;

            auto token = watcher.Received([&](BluetoothLEAdvertisementWatcher const&, BluetoothLEAdvertisementReceivedEventArgs const& args) {
                if (!deviceFound) {
                    std::wstring localName = args.Advertisement().LocalName().c_str();
                    if (localName == L"LED") {
                        LogMessage("Found Arduino LED device!");
                        targetDeviceAddress = args.BluetoothAddress();
                        deviceFound = true;
                    }
                }
                });

            watcher.Start();

            // Wait for device discovery with timeout
            auto startTime = std::chrono::steady_clock::now();
            while (!deviceFound && !g_shouldExit &&
                (std::chrono::steady_clock::now() - startTime) < SCAN_TIMEOUT) {
                std::this_thread::sleep_for(std::chrono::milliseconds(100));
            }

            watcher.Stop();

            if (g_shouldExit || !deviceFound) {
                if (!deviceFound) {
                    LogMessage("Arduino device not found during scan");
                }
                isConnecting = false;
                return false;
            }

            // Connect to device
            LogMessage("Connecting to Arduino...");
            auto deviceTask = BluetoothLEDevice::FromBluetoothAddressAsync(targetDeviceAddress);
            device = deviceTask.get();

            if (!device) {
                LogMessage("Failed to create device object");
                isConnecting = false;
                return false;
            }

            // Set up connection status change handler
            connectionStatusToken = device.ConnectionStatusChanged([this](BluetoothLEDevice const& sender, auto const&) {
                try {
                    auto status = sender.ConnectionStatus();
                    if (status == BluetoothConnectionStatus::Disconnected) {
                        LogMessage("Device disconnected - connection status changed event");
                        std::lock_guard<std::mutex> lock(connectionMutex);
                        isConnected = false;
                    }
                }
                catch (...) {
                    LogMessage("Error in connection status change handler");
                }
                });

            // Get GATT services with retry logic
            int retries = 3;
            GattDeviceServicesResult gattResult{ nullptr };

            while (retries > 0 && !g_shouldExit) {
                try {
                    gattResult = device.GetGattServicesAsync().get();
                    if (gattResult.Status() == GattCommunicationStatus::Success) {
                        break;
                    }
                    LogMessage("GATT services failed, retrying... (" + std::to_string(retries) + " left)");
                    std::this_thread::sleep_for(std::chrono::milliseconds(500));
                    retries--;
                }
                catch (...) {
                    LogMessage("Exception getting GATT services, retrying... (" + std::to_string(retries) + " left)");
                    std::this_thread::sleep_for(std::chrono::milliseconds(500));
                    retries--;
                }
            }

            if (!gattResult || gattResult.Status() != GattCommunicationStatus::Success) {
                LogMessage("Failed to get GATT services after retries");
                cleanupConnection();
                isConnecting = false;
                return false;
            }

            if (device.ConnectionStatus() != BluetoothConnectionStatus::Connected) {
                LogMessage("Device not connected after GATT access");
                cleanupConnection();
                isConnecting = false;
                return false;
            }

            // Find the switch characteristic
            winrt::guid switchUuid = parseUUID(SWITCH_CHARACTERISTIC_UUID);

            for (auto&& service : gattResult.Services()) {
                auto charResult = service.GetCharacteristicsAsync().get();
                if (charResult.Status() == GattCommunicationStatus::Success) {
                    for (auto&& characteristic : charResult.Characteristics()) {
                        if (characteristic.Uuid() == switchUuid) {
                            LogMessage("Found switch characteristic - Connected!");
                            switchCharacteristic = characteristic;
                            gattService = service;
                            isConnected = true;
                            isConnecting = false;
                            return true;
                        }
                    }
                }
            }

            LogMessage("Switch characteristic not found");
            cleanupConnection();
            isConnecting = false;
            return false;

        }
        catch (const std::exception& ex) {
            LogMessage(std::string("BLE connection error: ") + ex.what());
            cleanupConnection();
            isConnecting = false;
            return false;
        }
        catch (...) {
            LogMessage("Unknown BLE connection error");
            cleanupConnection();
            isConnecting = false;
            return false;
        }
    }

    bool setLEDState(bool state) {
        std::lock_guard<std::mutex> lock(connectionMutex);

        if (!isConnected || !switchCharacteristic) {
            return false;
        }

        try {
            if (!device || device.ConnectionStatus() != BluetoothConnectionStatus::Connected) {
                LogMessage("Device disconnected during LED operation");
                isConnected = false;
                return false;
            }

            DataWriter writer;
            writer.WriteByte(state ? 1 : 0);
            IBuffer buffer = writer.DetachBuffer();

            auto result = switchCharacteristic.WriteValueAsync(buffer).get();
            if (result == GattCommunicationStatus::Success) {
                LogMessage(state ? "LED turned ON" : "LED turned OFF");
                return true;
            }
            else {
                LogMessage("Failed to send LED command - communication error");
                isConnected = false;
                return false;
            }
        }
        catch (const std::exception& ex) {
            LogMessage(std::string("LED control error: ") + ex.what());
            isConnected = false;
            return false;
        }
        catch (...) {
            LogMessage("Unknown LED control error");
            isConnected = false;
            return false;
        }
    }

    void disconnect() {
        std::lock_guard<std::mutex> lock(connectionMutex);
        cleanupConnection();
    }

    void forceReconnect() {
        LogMessage("Force reconnect requested");
        disconnect();
        lastConnectionAttempt = std::chrono::steady_clock::now() - RECONNECT_DELAY;
    }

private:
    void cleanupConnection() {
        // This method should be called with connectionMutex already locked
        isConnected = false;
        isConnecting = false;

        if (device && connectionStatusToken.value != 0) {
            try {
                device.ConnectionStatusChanged(connectionStatusToken);
                connectionStatusToken = {};
            }
            catch (...) {
                // Ignore cleanup errors
            }
        }

        if (switchCharacteristic) {
            switchCharacteristic = nullptr;
        }

        if (gattService) {
            try {
                gattService.Close();
            }
            catch (...) {
                // Ignore cleanup errors
            }
            gattService = nullptr;
        }

        if (device) {
            try {
                device.Close();
            }
            catch (...) {
                // Ignore cleanup errors
            }
            device = nullptr;
        }
    }
};

// Microphone Monitor
class MicrophoneMonitor {
private:
    IMMDeviceEnumerator* pEnumerator;
    IMMDevice* pDevice;
    IAudioSessionManager2* pSessionManager;
    bool initialized;

public:
    MicrophoneMonitor() : pEnumerator(nullptr), pDevice(nullptr), pSessionManager(nullptr), initialized(false) {}

    ~MicrophoneMonitor() {
        cleanup();
    }

    bool initialize() {
        if (initialized) {
            return true;
        }

        HRESULT hr = CoInitializeEx(nullptr, COINIT_MULTITHREADED);
        if (FAILED(hr) && hr != RPC_E_CHANGED_MODE) {
            return false;
        }

        hr = CoCreateInstance(__uuidof(MMDeviceEnumerator), nullptr, CLSCTX_ALL,
            __uuidof(IMMDeviceEnumerator), (void**)&pEnumerator);
        if (FAILED(hr)) {
            LogMessage("Failed to create MMDeviceEnumerator");
            return false;
        }

        hr = pEnumerator->GetDefaultAudioEndpoint(eCapture, eConsole, &pDevice);
        if (FAILED(hr)) {
            LogMessage("Failed to get default audio capture device");
            return false;
        }

        hr = pDevice->Activate(__uuidof(IAudioSessionManager2), CLSCTX_ALL,
            nullptr, (void**)&pSessionManager);
        if (FAILED(hr)) {
            LogMessage("Failed to activate audio session manager");
            return false;
        }

        initialized = true;
        LogMessage("Microphone monitor initialized");
        return true;
    }

    bool isMicrophoneInUse() {
        if (!initialized || !pSessionManager) {
            return false;
        }

        IAudioSessionEnumerator* pSessionEnumerator = nullptr;
        HRESULT hr = pSessionManager->GetSessionEnumerator(&pSessionEnumerator);
        if (FAILED(hr)) {
            return false;
        }

        int sessionCount = 0;
        hr = pSessionEnumerator->GetCount(&sessionCount);
        if (FAILED(hr)) {
            pSessionEnumerator->Release();
            return false;
        }

        bool micInUse = false;
        for (int i = 0; i < sessionCount && !micInUse; i++) {
            IAudioSessionControl* pSessionControl = nullptr;
            hr = pSessionEnumerator->GetSession(i, &pSessionControl);
            if (SUCCEEDED(hr)) {
                AudioSessionState state;
                if (SUCCEEDED(pSessionControl->GetState(&state)) && state == AudioSessionStateActive) {
                    micInUse = true;
                }
                pSessionControl->Release();
            }
        }

        pSessionEnumerator->Release();
        return micInUse;
    }

private:
    void cleanup() {
        if (pSessionManager) {
            pSessionManager->Release();
            pSessionManager = nullptr;
        }
        if (pDevice) {
            pDevice->Release();
            pDevice = nullptr;
        }
        if (pEnumerator) {
            pEnumerator->Release();
            pEnumerator = nullptr;
        }
        initialized = false;
    }
};

// Global instances
MicrophoneMonitor g_monitor;
ArduinoBLEController g_bleController;

// Window procedure
LRESULT CALLBACK WindowProc(HWND hWnd, UINT uMsg, WPARAM wParam, LPARAM lParam) {
    switch (uMsg) {
    case WM_TRAYICON:
        switch (lParam) {
        case WM_RBUTTONUP:
            ShowContextMenu(hWnd);
            break;
        case WM_LBUTTONDBLCLK:
            if (g_consoleVisible) {
                HideConsole();
            }
            else {
                ShowConsole();
            }
            break;
        }
        break;

    case WM_COMMAND:
        switch (LOWORD(wParam)) {
        case ID_TRAY_SHOW_CONSOLE:
            ShowConsole();
            break;
        case ID_TRAY_HIDE_CONSOLE:
            HideConsole();
            break;
        case ID_TRAY_RECONNECT:
            LogMessage("Manual reconnection requested");
            g_bleController.forceReconnect();
            break;
        case ID_TRAY_ABOUT:
            MessageBox(hWnd, L"Microphone LED Monitor v1.2\n\nMonitors microphone usage and controls Arduino LED via Bluetooth LE.\n\nDouble-click tray icon to show/hide console.\nClose button on console is disabled - use tray menu to hide.", L"About", MB_OK | MB_ICONINFORMATION);
            break;
        case ID_TRAY_EXIT:
            g_shouldExit = true;
            PostQuitMessage(0);
            break;
        }
        break;

    case WM_DESTROY:
        g_shouldExit = true;
        RemoveTrayIcon();
        PostQuitMessage(0);
        break;

    default:
        return DefWindowProc(hWnd, uMsg, wParam, lParam);
    }
    return 0;
}

// Monitor thread
void monitorThread() {
    LogMessage("Starting microphone monitoring...");

    bool lastMicState = false;
    bool lastConnectedState = false;
    bool forceStateUpdate = false;
    auto lastStatusUpdate = std::chrono::steady_clock::now();
    const auto STATUS_UPDATE_INTERVAL = std::chrono::seconds(30);

    while (!g_shouldExit) {
        try {
            // Check connection status
            bool connected = g_bleController.getConnectionStatus();

            // Attempt reconnection if needed
            if (!connected && g_bleController.shouldAttemptReconnect()) {
                LogMessage("Attempting to reconnect...");
                connected = g_bleController.connectToArduino();
                if (connected) {
                    forceStateUpdate = true; // Force LED state update after reconnection
                }
            }

            // Check microphone status
            bool micInUse = g_monitor.isMicrophoneInUse();

            // Update LED state if microphone state changed, we're connected, or force update needed
            if ((micInUse != lastMicState || forceStateUpdate) && connected) {
                LogMessage(micInUse ? "Microphone ACTIVE - LED ON" : "Microphone INACTIVE - LED OFF");
                if (g_bleController.setLEDState(micInUse)) {
                    lastMicState = micInUse;
                    forceStateUpdate = false;
                }
                else {
                    LogMessage("Failed to update LED state - connection may be lost");
                    connected = false; // Will trigger reconnection on next loop
                }
            }

            // Handle connection state changes
            if (connected != lastConnectedState) {
                if (connected) {
                    LogMessage("Arduino connected successfully");
                    forceStateUpdate = true; // Update LED state on new connection
                }
                else {
                    LogMessage("Arduino disconnected - will attempt reconnection");
                }
                UpdateTrayIcon(connected, micInUse);
                lastConnectedState = connected;
            }

            // Periodic status update
            auto now = std::chrono::steady_clock::now();
            if (now - lastStatusUpdate >= STATUS_UPDATE_INTERVAL) {
                LogMessage("Status: " + std::string(connected ? "Connected" : "Disconnected") +
                    ", Mic: " + std::string(micInUse ? "Active" : "Inactive"));
                lastStatusUpdate = now;
            }

        }
        catch (const std::exception& ex) {
            LogMessage(std::string("Monitor thread error: ") + ex.what());
        }
        catch (...) {
            LogMessage("Unknown monitor thread error");
        }

        // Sleep for 1 second
        std::this_thread::sleep_for(std::chrono::seconds(1));
    }

    LogMessage("Monitoring stopped");
}

int WINAPI WinMain(_In_ HINSTANCE hInstance, _In_opt_ HINSTANCE hPrevInstance, _In_ LPSTR lpCmdLine, _In_ int nCmdShow) {
    // Initialize COM
    HRESULT hr = CoInitializeEx(nullptr, COINIT_APARTMENTTHREADED);
    if (FAILED(hr)) {
        MessageBox(nullptr, L"Failed to initialize COM", L"Error", MB_OK | MB_ICONERROR);
        return 1;
    }

    // Create window class
    WNDCLASS wc = {};
    wc.lpfnWndProc = WindowProc;
    wc.hInstance = hInstance;
    wc.lpszClassName = L"MicrophoneLEDMonitor";
    wc.hCursor = LoadCursor(nullptr, IDC_ARROW);

    RegisterClass(&wc);

    // Create hidden window
    g_hWnd = CreateWindow(L"MicrophoneLEDMonitor", L"Microphone LED Monitor",
        WS_OVERLAPPEDWINDOW, CW_USEDEFAULT, CW_USEDEFAULT,
        CW_USEDEFAULT, CW_USEDEFAULT, nullptr, nullptr, hInstance, nullptr);

    if (!g_hWnd) {
        MessageBox(nullptr, L"Failed to create window", L"Error", MB_OK | MB_ICONERROR);
        CoUninitialize();
        return 1;
    }

    // Add system tray icon
    AddTrayIcon(g_hWnd);

    // Initialize microphone monitor
    if (!g_monitor.initialize()) {
        MessageBox(nullptr, L"Failed to initialize microphone monitor", L"Error", MB_OK | MB_ICONERROR);
        RemoveTrayIcon();
        CoUninitialize();
        return 1;
    }

    LogMessage("Microphone LED Monitor started");
    LogMessage("Double-click tray icon to show/hide console");

    // Start monitoring thread
    std::thread monitorThreadHandle(monitorThread);

    // Message loop
    MSG msg;
    while (GetMessage(&msg, nullptr, 0, 0)) {
        TranslateMessage(&msg);
        DispatchMessage(&msg);
    }

    // Cleanup
    g_shouldExit = true;
    if (monitorThreadHandle.joinable()) {
        monitorThreadHandle.join();
    }

    RemoveTrayIcon();
    CleanupConsole();

    if (g_hMenu) {
        DestroyMenu(g_hMenu);
    }

    CoUninitialize();
    return 0;
}
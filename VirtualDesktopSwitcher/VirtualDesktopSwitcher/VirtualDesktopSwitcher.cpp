#include "VirtualDesktops.h"

IVirtualDesktopManagerInternal* pDesktopManager = nullptr;
HHOOK hKeyboardHook = nullptr;
bool reverse = false;

LRESULT CALLBACK KeyHandler(int nCode, WPARAM wParam, LPARAM lParam);
UINT SimulateInput(WORD wKeyCode, DWORD dwFlags = 0);
HRESULT GetVirtualDesktopManager(IVirtualDesktopManagerInternal** ppDesktopManagerInternal);
HRESULT SwitchVirtualDesktop(IVirtualDesktopManagerInternal *pDesktopManager, bool reverse = false);

int WINAPI wWinMain(HINSTANCE hInstance, HINSTANCE hPrevInstance, PWSTR pCmdLine, int nCmdShow)
{
    HRESULT hr = GetVirtualDesktopManager(&pDesktopManager);

    if (FAILED(hr))
        return 1;

    hKeyboardHook = SetWindowsHookEx(WH_KEYBOARD_LL, KeyHandler, hInstance, 0);

    MSG msg;

    // Main message loop:
    while (GetMessage(&msg, nullptr, 0, 0))
    {
        TranslateMessage(&msg);
        DispatchMessage(&msg);
    }

    UnhookWindowsHookEx(hKeyboardHook);
    pDesktopManager->Release();
    return (int)msg.wParam;
}

LRESULT CALLBACK KeyHandler(int nCode, WPARAM wParam, LPARAM lParam) {

    if (GetAsyncKeyState(VK_LWIN) != 0)
    if (nCode >= 0 && (wParam == WM_KEYDOWN || wParam == WM_KEYUP))
    {
        KBDLLHOOKSTRUCT kbdStruct = *((KBDLLHOOKSTRUCT*)lParam);
        HWND hTaskView = FindWindow("MultitaskingViewFrame", NULL);

        switch (wParam)
        {
        case WM_KEYDOWN:
            if (kbdStruct.vkCode == VK_TAB)
            {
                //TODO: Find another way to open Task View 

                if (GetAsyncKeyState(VK_SHIFT) != 0 && reverse == false)
                {
                    /*
                    Achtung! Костыль!
                    Task View doesn't open when you press [Win + Shift + Tab]
                    Release [Shift] to fix this, but remeber reverse direction
                    */

                    reverse = true;
                    if (hTaskView == NULL)
                        SimulateInput(VK_SHIFT, KEYEVENTF_KEYUP);
                }

                SwitchVirtualDesktop(pDesktopManager, reverse);

                //Interrupt [Tab] to prevent Task View closing
                if (hTaskView != NULL)
                    return 1;
            }
            break;

        case WM_KEYUP:
            if (hTaskView != NULL)
            {
                if (kbdStruct.vkCode == VK_LWIN)
                {
                    //TODO: Find another way to close Task View 
                    SimulateInput(VK_ESCAPE);
                    reverse = false;
                }
                else if (kbdStruct.vkCode == VK_LSHIFT)
                {
                    reverse = false;
                }
            }
            break;

        default:
            break;
        }
    }

    return CallNextHookEx(hKeyboardHook, nCode, wParam, lParam);
}

UINT SimulateInput(WORD wKeyCode, DWORD dwFlags)
{
    INPUT ip;

    ip.type = INPUT_KEYBOARD;
    ip.ki.wScan = 0;
    ip.ki.time = 0;
    ip.ki.dwExtraInfo = 0;

    ip.ki.wVk = wKeyCode;
    ip.ki.dwFlags = dwFlags; // 0 for key press

    return SendInput(1, &ip, sizeof(INPUT));
}

HRESULT GetVirtualDesktopManager(IVirtualDesktopManagerInternal** ppDesktopManagerInternal)
{
    ::CoInitializeEx(NULL, COINIT_MULTITHREADED);

    IServiceProvider* pServiceProvider = nullptr;
    HRESULT hr = ::CoCreateInstance(
        CLSID_ImmersiveShell, NULL, CLSCTX_LOCAL_SERVER,
        __uuidof(IServiceProvider), (PVOID*)&pServiceProvider);

    if (SUCCEEDED(hr))
    {
        hr = pServiceProvider->QueryService(CLSID_VirtualDesktopAPI_Unknown, ppDesktopManagerInternal);
        pServiceProvider->Release();
    }
    return hr;
}

HRESULT SwitchVirtualDesktop(IVirtualDesktopManagerInternal *pDesktopManager, bool reverse)
{
    IVirtualDesktop *pCurrentDesktop = nullptr;
    HRESULT hr = pDesktopManager->GetCurrentDesktop(&pCurrentDesktop);

    if (SUCCEEDED(hr))
    {
        IVirtualDesktop *pNextDesktop = nullptr;

        AdjacentDesktop direction = reverse ? AdjacentDesktop::LeftDirection : AdjacentDesktop::RightDirection;
        hr = pDesktopManager->GetAdjacentDesktop(pCurrentDesktop, direction, &pNextDesktop);

        if (FAILED(hr))
        {
            IObjectArray *pObjectArray = nullptr;
            hr = pDesktopManager->GetDesktops(&pObjectArray);

            if (SUCCEEDED(hr))
            {
                int nextIndex = 0;
                if (reverse)
                {
                    UINT count;
                    hr = pObjectArray->GetCount(&count);

                    if (SUCCEEDED(hr))
                        nextIndex = count - 1;
                }

                hr = pObjectArray->GetAt(nextIndex, __uuidof(IVirtualDesktop), (void**)&pNextDesktop);
                pObjectArray->Release();
            }
        }

        if (SUCCEEDED(hr))
        {
            hr = pDesktopManager->SwitchDesktop(pNextDesktop);
            pNextDesktop->Release();
        }

        pCurrentDesktop->Release();
    }

    return hr;
}

#include "VirtualDesktops.h"

HRESULT GetVirtualDesktopManager(IVirtualDesktopManagerInternal** ppDesktopManagerInternal)
{
    ::CoInitialize(NULL);

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

HRESULT SwitchVirtualDesktop(IVirtualDesktopManagerInternal *pDesktopManager)
{
    IVirtualDesktop *pCurrentDesktop = nullptr;
    HRESULT hr = pDesktopManager->GetCurrentDesktop(&pCurrentDesktop);

    if (SUCCEEDED(hr))
    {
        IVirtualDesktop *pNextDesktop = nullptr;

        AdjacentDesktop direction = AdjacentDesktop::RightDirection;
        hr = pDesktopManager->GetAdjacentDesktop(pCurrentDesktop, direction, &pNextDesktop);
        
        if (FAILED(hr))
        {
            IObjectArray *pObjectArray = nullptr;
            hr = pDesktopManager->GetDesktops(&pObjectArray);

            if (SUCCEEDED(hr))
            {
                hr = pObjectArray->GetAt(0, __uuidof(IVirtualDesktop), (void**)&pNextDesktop);
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

int WINAPI wWinMain(HINSTANCE hInstance, HINSTANCE hPrevInstance, PWSTR pCmdLine, int nCmdShow)
{
    IVirtualDesktopManagerInternal* pDesktopManager = nullptr;
    HRESULT hr = GetVirtualDesktopManager(&pDesktopManager);

    if (SUCCEEDED(hr))
    {
        SwitchVirtualDesktop(pDesktopManager);
        pDesktopManager->Release();
    }

    return 0;
}
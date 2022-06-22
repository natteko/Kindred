#define _WIN32_DCOM

#include "QuerySink.h"
#include <iostream>
#include <WbemCli.h>
#include <WbemIdl.h>

#pragma comment(lib, "wbemuuid.lib")

// https://docs.microsoft.com/en-gb/windows/win32/wmisdk/creating-a-wmi-application-using-c-
// https://docs.microsoft.com/en-gb/windows/win32/wmisdk/receiving-a-wmi-event
// https://docs.microsoft.com/en-gb/windows/win32/wmisdk/example--receiving-event-notifications-through-wmi-

int main() {
    HRESULT hres;

    hres = CoInitializeEx(NULL, COINIT_MULTITHREADED);
    if (FAILED(hres)) {
        std::cout << "Failed to initialise COM library. Error code = 0x" << std::hex << hres << std::endl;

        CoUninitialize();
        return hres;
    };

    hres = CoInitializeSecurity(
        NULL,                        // Security descriptor    
        -1,                          // COM negotiates authentication service
        NULL,                        // Authentication services
        NULL,                        // Reserved
        RPC_C_AUTHN_LEVEL_DEFAULT,   // Default authentication level for proxies
        RPC_C_IMP_LEVEL_IMPERSONATE, // Default Impersonation level for proxies
        NULL,                        // Authentication info
        EOAC_NONE,                   // Additional capabilities of the client or server
        NULL                         // Reserved
    );
    if (FAILED(hres)) {
        std::cout << "Failed to initialize COM security. Error code = 0x" << std::hex << hres << std::endl;
        CoUninitialize();
        return hres;
    };

    IWbemLocator* pWbemLocator = NULL;
    hres = CoCreateInstance(CLSID_WbemLocator, 0, CLSCTX_INPROC_SERVER, IID_IWbemLocator, (LPVOID*)&pWbemLocator);
    if (FAILED(hres)) {
        std::cout << "Failed to create IWbemLocator object. Err code = 0x" << std::hex << hres << std::endl;

        CoUninitialize();
        return hres;
    };

    IWbemServices* pWbemServices = NULL;
    hres = pWbemLocator->ConnectServer(
        BSTR(L"ROOT\\CIMV2"),     // Namespace
        NULL,                       // User name 
        NULL,                       // User password
        0,                          // Locale 
        NULL,                       // Security flags
        0,                          // Authority 
        0,                          // Context object 
        &pWbemServices              // IWbemServices proxy
    );
    if (FAILED(hres)) {
        std::cout << "Failed to create IWbemServices object. Err code = 0x" << std::hex << hres << std::endl;

        pWbemLocator->Release();
        CoUninitialize();
        return hres;
    };

    hres = CoSetProxyBlanket(
        pWbemServices,              // Pointer to proxy
        RPC_C_AUTHN_WINNT,          // Authentication service
        RPC_C_AUTHZ_NONE,           // Also authentication service...?
        NULL,                       // Server principal name
        RPC_C_AUTHN_LEVEL_CALL,     // Authentication level
        RPC_C_IMP_LEVEL_IMPERSONATE,// Impersonation level
        NULL,                       // Pointer to an RPC_AUTH_IDENTITY_HHANDLE
        EOAC_NONE                   // Selection from the EOAC enum
    );
    if (FAILED(hres)) {
        std::cout << "Failed to set IWbemServices object proxy blanket security level. Err code = 0x" << std::hex << hres << std::endl;

        pWbemServices->Release();
        pWbemLocator->Release();
        CoUninitialize();
        return hres;
    };

    QuerySink* pSink = new QuerySink;
    pSink->AddRef();

    //IUnsecuredApartment* pUnsecApp = NULL;

    //hres = CoCreateInstance(CLSID_UnsecuredApartment, NULL,
    //    CLSCTX_LOCAL_SERVER, IID_IUnsecuredApartment,
    //    (void**)&pUnsecApp);

    //IUnknown* pStubUnk = NULL;
    //pUnsecApp->CreateObjectStub(pSink, &pStubUnk);

    //IWbemObjectSink* pStubSink = NULL;
    //pStubUnk->QueryInterface(IID_IWbemObjectSink,
    //    (void**)&pStubSink);

    hres = pWbemServices->ExecNotificationQueryAsync(
        BSTR(L"WQL"),
        BSTR(L"SELECT * "
            "FROM __InstanceCreationEvent WITHIN 1 "
            "WHERE TargetInstance ISA 'Win32_Process'"),
        WBEM_FLAG_SEND_STATUS,
        NULL,
        pSink
    );
    if (FAILED(hres)) {
        std::cout << "Failed to call IWbemServices::ExecNotificationQueryAsync. Err code = 0x" << std::hex << hres << std::endl;

        pSink->Release();
        pWbemServices->Release();
        pWbemLocator->Release();
        CoUninitialize();
        return 1;
    };

    Sleep(10000);

    hres = pWbemServices->CancelAsyncCall(pSink);

    pWbemServices->Release();
    pWbemLocator->Release();
    pSink->Release();
    CoUninitialize();
    return 0;
}


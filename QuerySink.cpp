// https://docs.microsoft.com/en-gb/windows/win32/wmisdk/iwbemobjectsink

#include "QuerySink.h"
#include <iostream>
#include <comdef.h>
#include <psapi.h>

ULONG QuerySink::AddRef() {
    return InterlockedIncrement(&m_lRef);
}

ULONG QuerySink::Release() {
    LONG lRef = InterlockedDecrement(&m_lRef);
    if (lRef == 0)
        delete this;
    return lRef;
}

HRESULT QuerySink::QueryInterface(REFIID riid, void** ppv) {
    if (riid == IID_IUnknown || riid == IID_IWbemObjectSink) {
        *ppv = (IWbemObjectSink*)this;
        AddRef();
        return WBEM_S_NO_ERROR;
    } else return E_NOINTERFACE;
}

HRESULT QuerySink::Indicate(long lObjCount, IWbemClassObject** pArray) {
    for (int i = 0; i < lObjCount; i++) {
        IWbemClassObject* pObj = pArray[i];

        HRESULT hres;
        _variant_t vtProp;

        hres = pObj->Get(BSTR(L"TargetInstance"), 0, &vtProp, NULL, NULL);
        if (FAILED(hres)) {
            std::cout << "Failed to get TargetInstance object from IWBemClassObject. Err code = 0x" << std::hex << hres << std::endl;

            CoUninitialize();
            return hres;
        }

        IUnknown* str = vtProp; // // I have no idea why this is so goddamn important that the program does not run without it
        hres = str->QueryInterface(&pObj); // but I send my thanks to https://stackoverflow.com/a/31754632/14106820 anyhow
        if (SUCCEEDED(hres)) { // we're gonna go with a succeeded here because I have no idea what to write in the error message
            _variant_t propValue;

            hres = pObj->Get(L"Name", 0, &propValue, NULL, NULL);
            if (SUCCEEDED(hres)) {
                if ((propValue.vt == VT_NULL) || (propValue.vt == VT_EMPTY)) {
                    std::wcout << "Name property unavailable; VARIANT type is: " << ((propValue.vt == VT_NULL) ? "NULL" : "EMPTY") << std::endl;
                } else {
                    if (propValue.vt & VT_ARRAY) { // is this equivalent to propValue.vt == VT_ARRAY? I feel like the bit operation has some significance here...
                        std::wcout << "Name property unavailable; unsupported VARIANT type (ARRAY)" << std::endl;
                    } else {
                        std::wcout << "Name: " << propValue.bstrVal << std::endl;
                    }
                }
            }

            VariantClear(&propValue);

            _bstr_t processName(propValue.bstrVal);
            _bstr_t targetProcessName("Discord.exe");
            wchar_t friendProcessPath[] = L"C:\\Users\\natteko\\Downloads\\EasyRP-windows\\easyrp.exe";

            if (processName == targetProcessName) {

                hres = pObj->Get(L"ProcessId", 0, &propValue, NULL, NULL);
                if (SUCCEEDED(hres)) {
                    if ((propValue.vt == VT_NULL) || (propValue.vt == VT_EMPTY)) {
                        std::wcout << "ProcessId property unavailable; VARIANT type is: " << ((propValue.vt == VT_NULL) ? "NULL" : "EMPTY") << std::endl;
                    } else {
                        if (propValue.vt & VT_ARRAY) {
                            std::wcout << "ProcessId property unavailable; unsupported VARIANT type (ARRAY)" << std::endl;
                        } else {
                            std::wcout << "ProcessId: " << propValue.intVal << std::endl;
 
                            STARTUPINFO si;
                            PROCESS_INFORMATION pi;

                            // HANDLE hProcess = OpenProcess(PROCESS_QUERY_INFORMATION | PROCESS_VM_READ, FALSE, propValue.intVal); // get a handle to the process. 

                            ZeroMemory(&si, sizeof(si));
                            si.cb = sizeof(si);
                            ZeroMemory(&pi, sizeof(pi));

                            wchar_t proc[] = L"cmd.exe";

                            // Start the child process. 
                            if (!CreateProcess(
                                NULL,               // No module name (use command line)
                                friendProcessPath,  // Command line
                                NULL,               // Process handle not inheritable
                                NULL,               // Thread handle not inheritable
                                FALSE,              // Set handle inheritance to FALSE
                                0,                  // No creation flags
                                NULL,               // Use parent's environment block
                                NULL,               // Use parent's starting directory 
                                &si,                // Pointer to STARTUPINFO structure
                                &pi)                // Pointer to PROCESS_INFORMATION structure
                                ) {
                                printf("CreateProcess failed (%d).\n", GetLastError());
                                return 1;
                            }

                            // Wait until child process exits.
                            WaitForSingleObject(pi.hProcess, INFINITE);

                            // Close process and thread handles. 
                            CloseHandle(pi.hProcess);
                            CloseHandle(pi.hThread);
                        }
                    }
                }
            }

            VariantClear(&propValue);
        }

        //// debug purposes
        //BSTR buffer;
        //pObj->GetObjectText(NULL, &buffer);
        //printf("%S\n\n", buffer);

        str->Release();
        VariantClear(&vtProp);

        // AddRef() is only required if the object will be held after
        // the return to the caller.
    }
    return WBEM_S_NO_ERROR;
}

HRESULT QuerySink::SetStatus(
    LONG lFlags,
    HRESULT hResult,
    BSTR strParam,
    IWbemClassObject __RPC_FAR* pObjParam
) {
    if (lFlags == WBEM_STATUS_COMPLETE) {
        printf("Call complete. QuerySink::SetStatus hResult = 0x%X\n", hResult);
    } else if (lFlags == WBEM_STATUS_PROGRESS) {
        printf("Call in progress.\n");
    } else {
        printf("QuerySink::SetStatus hResult = 0x%X\n", hResult);
    }

    return WBEM_S_NO_ERROR;
}

#include "WMIC.h"
// windows
#include <Wbemidl.h>
#include <Windows.h>
#include <comdef.h>

#pragma comment(lib, "wbemuuid.lib")

// STL
#include <print>
#include <sstream>

std::string ConvertWideToMultiByte(std::wstring_view str) {
    BOOL isUsedDefaultChar;
    if(auto const size = WideCharToMultiByte(CP_ACP, NULL, str.data(), -1, NULL, NULL, NULL, &isUsedDefaultChar); !isUsedDefaultChar && size > 0) {
        std::vector<char> result(size);
        WideCharToMultiByte(CP_ACP, NULL, str.data(), -1, result.data(), size, NULL, NULL);
        return result.data();
    }
    return {};
}

std::wstring ConvertMultiByteToWide(std::string_view str) {
    if(auto const size = MultiByteToWideChar(CP_UTF8, NULL, str.data(), -1, NULL, NULL); size > 0) {
        std::vector<wchar_t> result(size);
        MultiByteToWideChar(CP_UTF8, NULL, str.data(), -1, result.data(), size);
        return result.data();
    }
    return {};
}

struct WMICContext {
    HRESULT hResult     = NULL;
    IWbemLocator *pLoc  = NULL;
    IWbemServices *pSvc = NULL;
    std::stringstream message;
};

IEnumWbemClassObject *ExecQuery(std::unique_ptr<struct WMICContext> &ctx, const std::wstring className) {
    IEnumWbemClassObject *pEnumerator = NULL;
    std::wstring cmd                  = L"SELECT * FROM ";
    cmd += className;

    ctx->hResult = ctx->pSvc->ExecQuery(bstr_t("WQL"), bstr_t(cmd.data()), WBEM_FLAG_FORWARD_ONLY | WBEM_FLAG_RETURN_IMMEDIATELY, NULL, &pEnumerator);

    if(FAILED(ctx->hResult)) {
        ctx->message.clear();
        ctx->message << "Query for xxx failed."
                     << " Error code = 0x" << std::hex << ctx->hResult << std::endl;
        ctx->pSvc->Release();
        ctx->pLoc->Release();
        CoUninitialize();
        throw WMICException(ctx->message.str());
        // Program has failed.
    }
    return pEnumerator;
}

WMIC::WMIC() {
    pCtx = std::make_unique<WMICContext>();

    // Step 1
    // Initialize COM
    pCtx->hResult = CoInitializeEx(0, COINIT_MULTITHREADED);
    if(FAILED(pCtx->hResult)) {
        pCtx->message.clear();
        pCtx->message << "Failed to initialize COM library. Error code = 0x" << std::hex << pCtx->hResult << std::endl;
        throw WMICException(pCtx->message.str());
    }

    // Step 2
    // Set general COM security levels
    pCtx->hResult = CoInitializeSecurity(NULL,
                                         -1,                           // COM authentication
                                         NULL,                         // Authentication services
                                         NULL,                         // Reserved
                                         RPC_C_AUTHN_LEVEL_DEFAULT,    // Default authentication
                                         RPC_C_IMP_LEVEL_IMPERSONATE,  // Default Impersonation
                                         NULL,                         // Authentication info
                                         EOAC_NONE,                    // Additional capabilities
                                         NULL                          // Reserved
    );
    if(FAILED(pCtx->hResult)) {
        pCtx->message.clear();
        pCtx->message << "Failed to initialize security. Error code = 0x" << std::hex << pCtx->hResult << std::endl;
        CoUninitialize();
        throw WMICException(pCtx->message.str());
        // Program has failed.
    }

    // Step 3
    // Obtain the initial locator to WMI
    pCtx->hResult = CoCreateInstance(CLSID_WbemLocator, 0, CLSCTX_INPROC_SERVER, IID_IWbemLocator, (LPVOID *)&pCtx->pLoc);

    if(FAILED(pCtx->hResult)) {
        pCtx->message.clear();
        pCtx->message << "Failed to create IWbemLocator object."
                      << " Err code = 0x" << std::hex << pCtx->hResult << std::endl;
        CoUninitialize();
        throw WMICException(pCtx->message.str());
        // Program has failed.
    }

    // Step 4
    // Connect to WMI through the IWbemLocator::ConnectServer method

    // Connect to the root\cimv2 namespace with
    // the current user and obtain pointer pSvc
    // to make IWbemServices calls.
    pCtx->hResult = pCtx->pLoc->ConnectServer(_bstr_t(L"ROOT\\CIMV2"),  // Object path of WMI namespace
                                              NULL,                     // User name. NULL = current user
                                              NULL,                     // User password. NULL = current
                                              NULL,                     // Locale. NULL indicates current
                                              NULL,                     // Security flags.
                                              NULL,                     // Authority (for example, Kerberos)
                                              NULL,                     // Context object
                                              &pCtx->pSvc               // pointer to IWbemServices proxy
    );

    if(FAILED(pCtx->hResult)) {
        pCtx->message.clear();
        pCtx->message << "Could not connect. Error code = 0x" << std::hex << pCtx->hResult << std::endl;
        pCtx->pLoc->Release();
        CoUninitialize();
        throw WMICException(pCtx->message.str());
        // Program has failed.
    }

    std::println("Connected to ROOT\\CIMV2 WMI namespace");

    // Step 5
    // Set security levels on the proxy
    pCtx->hResult = CoSetProxyBlanket(pCtx->pSvc,                   // Indicates the proxy to set
                                      RPC_C_AUTHN_WINNT,            // RPC_C_AUTHN_xxx
                                      RPC_C_AUTHZ_NONE,             // RPC_C_AUTHZ_xxx
                                      NULL,                         // Server principal name
                                      RPC_C_AUTHN_LEVEL_CALL,       // RPC_C_AUTHN_LEVEL_xxx
                                      RPC_C_IMP_LEVEL_IMPERSONATE,  // RPC_C_IMP_LEVEL_xxx
                                      NULL,                         // client identity
                                      EOAC_NONE                     // proxy capabilities
    );

    if(FAILED(pCtx->hResult)) {
        pCtx->message.clear();
        pCtx->message << "Could not set proxy blanket. Error code = 0x" << std::hex << pCtx->hResult << std::endl;
        pCtx->pSvc->Release();
        pCtx->pLoc->Release();
        CoUninitialize();
        throw WMICException(pCtx->message.str());
        // Program has failed.
    }

    // Step 6:
    // Use the IWbemServices pointer to make requests of WMI
}

WMIC::~WMIC() {
    // Cleanup
    pCtx->pSvc->Release();
    pCtx->pLoc->Release();
    CoUninitialize();
    // Program successfully completed.
}

WMIC_OperatingSystem WMIC::OperatingSystem() {
    IEnumWbemClassObject *pEnumerator = ExecQuery(pCtx, L"Win32_OperatingSystem");

    // Step 7:
    // Get the data from the query in step 6

    IWbemClassObject *pclsObj = NULL;
    ULONG uReturn             = 0;

    WMIC_OperatingSystem operatingSystem;
    while(pEnumerator) {
        HRESULT hr = pEnumerator->Next(WBEM_INFINITE, 1, &pclsObj, &uReturn);

        if(0 == uReturn) {
            break;
        }

        VARIANT vtProp;
        VariantInit(&vtProp);

        hr = pclsObj->Get(L"Caption", 0, &vtProp, 0, 0);
        if(SUCCEEDED(hr) && (V_VT(&vtProp) == VT_BSTR)) {
            operatingSystem.name = vtProp.bstrVal;
        }
        VariantClear(&vtProp);

        hr = pclsObj->Get(L"BuildNumber", 0, &vtProp, 0, 0);
        if(SUCCEEDED(hr) && (V_VT(&vtProp) == VT_BSTR)) {
            operatingSystem.buildNumber = vtProp.bstrVal;
        }
        VariantClear(&vtProp);

        hr = pclsObj->Get(L"Version", 0, &vtProp, 0, 0);
        if(SUCCEEDED(hr) && (V_VT(&vtProp) == VT_BSTR)) {
            operatingSystem.version = vtProp.bstrVal;
        }
        VariantClear(&vtProp);

        hr = pclsObj->Get(L"InstallDate", 0, &vtProp, 0, 0);
        if(SUCCEEDED(hr) && (V_VT(&vtProp) == VT_BSTR)) {
            operatingSystem.installDate = vtProp.bstrVal;
        }
        VariantClear(&vtProp);

        hr = pclsObj->Get(L"OSArchitecture", 0, &vtProp, 0, 0);
        if(SUCCEEDED(hr) && (V_VT(&vtProp) == VT_BSTR)) {
            operatingSystem.osArchitecture = vtProp.bstrVal;
        }
        VariantClear(&vtProp);

        hr = pclsObj->Get(L"RegisteredUser", 0, &vtProp, 0, 0);
        if(SUCCEEDED(hr) && (V_VT(&vtProp) == VT_BSTR)) {
            operatingSystem.registeredUser = vtProp.bstrVal;
        }
        VariantClear(&vtProp);

        hr = pclsObj->Get(L"SerialNumber", 0, &vtProp, 0, 0);
        if(SUCCEEDED(hr) && (V_VT(&vtProp) == VT_BSTR)) {
            operatingSystem.serialNumber = vtProp.bstrVal;
        }
        VariantClear(&vtProp);

#ifdef _DEBUG
        std::print("{}\t", W2A(operatingSystem.name));
        std::print("{}\t", W2A(operatingSystem.buildNumber));
        std::print("{}\t", W2A(operatingSystem.version));
        std::print("{}\t", W2A(operatingSystem.installDate));
        std::print("{}\t", W2A(operatingSystem.osArchitecture));
        std::print("{}\t", W2A(operatingSystem.registeredUser));
        std::println("{}\t", W2A(operatingSystem.serialNumber));
#endif
        pclsObj->Release();
    }
    if(pEnumerator) pEnumerator->Release();
    return operatingSystem;
}

std::vector<WMIC_VideoController> WMIC::VideoController() {
    IEnumWbemClassObject *pEnumerator = ExecQuery(pCtx, L"Win32_VideoController");

    // Step 7:
    // Get the data from the query in step 6

    IWbemClassObject *pclsObj = NULL;
    ULONG uReturn             = 0;

    std::vector<WMIC_VideoController> set;
    WMIC_VideoController videoController;
    while(pEnumerator) {
        HRESULT hr = pEnumerator->Next(WBEM_INFINITE, 1, &pclsObj, &uReturn);

        if(0 == uReturn) {
            break;
        }

        VARIANT vtProp;
        VariantInit(&vtProp);

        // Get the value of the Name property
        hr = pclsObj->Get(L"Name", 0, &vtProp, 0, 0);
        if(SUCCEEDED(hr) && (V_VT(&vtProp) == VT_BSTR)) {
            videoController.name = vtProp.bstrVal;
        }
        VariantClear(&vtProp);

        hr = pclsObj->Get(L"PNPDeviceID", 0, &vtProp, 0, 0);
        if(SUCCEEDED(hr) && (V_VT(&vtProp) == VT_BSTR)) {
            videoController.deviceID = vtProp.bstrVal;
        }
        VariantClear(&vtProp);

        set.emplace_back(videoController);
#ifdef _DEBUG
        std::print("{}\t", W2A(videoController.name));
        std::println("{}\t", W2A(videoController.deviceID));
#endif
        pclsObj->Release();
    }
    if(pEnumerator) pEnumerator->Release();

    return set;
}

std::vector<WMIC_DiskDrive> WMIC::DiskDrive() {
    IEnumWbemClassObject *pEnumerator = ExecQuery(pCtx, L"Win32_DiskDrive");

    // Step 7:
    // Get the data from the query in step 6

    IWbemClassObject *pclsObj = NULL;
    ULONG uReturn             = 0;

    std::vector<WMIC_DiskDrive> set;
    WMIC_DiskDrive diskDrive;
    while(pEnumerator) {
        HRESULT hr = pEnumerator->Next(WBEM_INFINITE, 1, &pclsObj, &uReturn);

        if(0 == uReturn) {
            break;
        }

        VARIANT vtProp;
        VariantInit(&vtProp);

        // Get the value of the Name property
        hr = pclsObj->Get(L"Caption", 0, &vtProp, 0, 0);
        if(SUCCEEDED(hr) && (V_VT(&vtProp) == VT_BSTR)) {
            diskDrive.name = vtProp.bstrVal;
        }
        VariantClear(&vtProp);

        hr = pclsObj->Get(L"PNPDeviceID", 0, &vtProp, 0, 0);
        if(SUCCEEDED(hr) && (V_VT(&vtProp) == VT_BSTR)) {
            diskDrive.deviceID = vtProp.bstrVal;
        }
        VariantClear(&vtProp);

        hr = pclsObj->Get(L"SerialNumber", 0, &vtProp, 0, 0);
        if(SUCCEEDED(hr) && (V_VT(&vtProp) == VT_BSTR)) {
            diskDrive.serialNumber = vtProp.bstrVal;
        }
        VariantClear(&vtProp);

        hr = pclsObj->Get(L"Size", 0, &vtProp, 0, 0);
        if(SUCCEEDED(hr) && (V_VT(&vtProp) == VT_BSTR)) {
            VarUI8FromStr(vtProp.bstrVal, NULL, LOCALE_NOUSEROVERRIDE, &diskDrive.size);
            diskDrive.size = diskDrive.size / (1024 * 1024 * 1024);  // 1 byte / (1024*1024*1024) = x GB
        }
        VariantClear(&vtProp);

        set.emplace_back(diskDrive);
#ifdef _DEBUG
        std::print("{}\t", W2A(diskDrive.name));
        std::print("{}\t", W2A(diskDrive.deviceID));
        std::print("{}\t", W2A(diskDrive.serialNumber));
        std::println("{} GB\t", diskDrive.size);
#endif
        pclsObj->Release();
    }
    if(pEnumerator) pEnumerator->Release();

    return set;
}

WMIC_BaseBoard WMIC::BaseBoard() {
    IEnumWbemClassObject *pEnumerator = ExecQuery(pCtx, L"Win32_BaseBoard");

    // Step 7:
    // Get the data from the query in step 6

    IWbemClassObject *pclsObj = NULL;
    ULONG uReturn             = 0;

    WMIC_BaseBoard baseBoard;
    while(pEnumerator) {
        HRESULT hr = pEnumerator->Next(WBEM_INFINITE, 1, &pclsObj, &uReturn);

        if(0 == uReturn) {
            break;
        }

        VARIANT vtProp;
        VariantInit(&vtProp);

        // Get the value of the Name property
        hr = pclsObj->Get(L"Name", 0, &vtProp, 0, 0);
        if(SUCCEEDED(hr) && (V_VT(&vtProp) == VT_BSTR)) {
            baseBoard.name = vtProp.bstrVal;
        }
        VariantClear(&vtProp);

        hr = pclsObj->Get(L"Manufacturer", 0, &vtProp, 0, 0);
        if(SUCCEEDED(hr) && (V_VT(&vtProp) == VT_BSTR)) {
            baseBoard.manufacturer = vtProp.bstrVal;
        }
        VariantClear(&vtProp);

        hr = pclsObj->Get(L"Product", 0, &vtProp, 0, 0);
        if(SUCCEEDED(hr) && (V_VT(&vtProp) == VT_BSTR)) {
            baseBoard.product = vtProp.bstrVal;
        }
        VariantClear(&vtProp);

        hr = pclsObj->Get(L"Version", 0, &vtProp, 0, 0);
        if(SUCCEEDED(hr) && (V_VT(&vtProp) == VT_BSTR)) {
            baseBoard.version = vtProp.bstrVal;
        }
        VariantClear(&vtProp);

        hr = pclsObj->Get(L"SerialNumber", 0, &vtProp, 0, 0);
        if(SUCCEEDED(hr) && (V_VT(&vtProp) == VT_BSTR)) {
            baseBoard.serialNumber = vtProp.bstrVal;
        }
        VariantClear(&vtProp);

#ifdef _DEBUG
        std::print("{}\t", W2A(baseBoard.name));
        std::print("{}\t", W2A(baseBoard.manufacturer));
        std::print("{}\t", W2A(baseBoard.product));
        std::print("{}\t", W2A(baseBoard.version));
        std::println("{}\t", W2A(baseBoard.serialNumber));
#endif

        pclsObj->Release();
    }
    if(pEnumerator) pEnumerator->Release();

    return baseBoard;
}

WMIC_BIOS WMIC::BIOS() {
    IEnumWbemClassObject *pEnumerator = ExecQuery(pCtx, L"Win32_BIOS");

    // Step 7:
    // Get the data from the query in step 6

    IWbemClassObject *pclsObj = NULL;
    ULONG uReturn             = 0;

    WMIC_BIOS bios;
    while(pEnumerator) {
        HRESULT hr = pEnumerator->Next(WBEM_INFINITE, 1, &pclsObj, &uReturn);

        if(0 == uReturn) {
            break;
        }

        VARIANT vtProp;
        VariantInit(&vtProp);

        // Get the value of the Name property
        hr = pclsObj->Get(L"ReleaseDate", 0, &vtProp, 0, 0);
        if(SUCCEEDED(hr) && (V_VT(&vtProp) == VT_BSTR)) {
            bios.releaseDate = vtProp.bstrVal;
        }
        VariantClear(&vtProp);

        hr = pclsObj->Get(L"Manufacturer", 0, &vtProp, 0, 0);
        if(SUCCEEDED(hr) && (V_VT(&vtProp) == VT_BSTR)) {
            bios.manufacturer = vtProp.bstrVal;
        }
        VariantClear(&vtProp);

        hr = pclsObj->Get(L"SerialNumber", 0, &vtProp, 0, 0);
        if(SUCCEEDED(hr) && (V_VT(&vtProp) == VT_BSTR)) {
            bios.serialNumber = vtProp.bstrVal;
        }
        VariantClear(&vtProp);

#ifdef _DEBUG
        std::print("{}\t", W2A(bios.manufacturer));
        std::print("{}\t", W2A(bios.releaseDate));
        std::println("{}\t", W2A(bios.serialNumber));
#endif

        pclsObj->Release();
    }
    if(pEnumerator) pEnumerator->Release();

    return bios;
}

std::vector<WMIC_PhysicalMemory> WMIC::PhysicalMemory() {
    IEnumWbemClassObject *pEnumerator = ExecQuery(pCtx, L"Win32_PhysicalMemory");

    // Step 7:
    // Get the data from the query in step 6

    IWbemClassObject *pclsObj = NULL;
    ULONG uReturn             = 0;

    std::vector<WMIC_PhysicalMemory> set;
    WMIC_PhysicalMemory physicalMemory;
    while(pEnumerator) {
        HRESULT hr = pEnumerator->Next(WBEM_INFINITE, 1, &pclsObj, &uReturn);

        if(0 == uReturn) {
            break;
        }

        VARIANT vtProp;
        VariantInit(&vtProp);

        // Get the value of the Name property
        hr = pclsObj->Get(L"Name", 0, &vtProp, 0, 0);
        if(SUCCEEDED(hr) && (V_VT(&vtProp) == VT_BSTR)) {
            physicalMemory.name = vtProp.bstrVal;
        }
        VariantClear(&vtProp);

        hr = pclsObj->Get(L"Manufacturer", 0, &vtProp, 0, 0);
        if(SUCCEEDED(hr) && (V_VT(&vtProp) == VT_BSTR)) {
            physicalMemory.manufacturer = vtProp.bstrVal;
        }
        VariantClear(&vtProp);

        hr = pclsObj->Get(L"SerialNumber", 0, &vtProp, 0, 0);
        if(SUCCEEDED(hr) && (V_VT(&vtProp) == VT_BSTR)) {
            physicalMemory.serialNumber = vtProp.bstrVal;
        }
        VariantClear(&vtProp);

        hr = pclsObj->Get(L"Speed", 0, &vtProp, 0, 0);
        if(SUCCEEDED(hr) && (V_VT(&vtProp) == VT_I4)) {
            physicalMemory.speed = vtProp.uintVal;
        }
        VariantClear(&vtProp);

        hr = pclsObj->Get(L"Capacity", 0, &vtProp, 0, 0);
        if(SUCCEEDED(hr) && (V_VT(&vtProp) == VT_BSTR)) {
            VarUI8FromStr(vtProp.bstrVal, NULL, LOCALE_NOUSEROVERRIDE, &physicalMemory.size);
            physicalMemory.size = physicalMemory.size / (1024 * 1024 * 1024);  // 1 byte / (1024*1024*1024) = x GB
        }
        VariantClear(&vtProp);

        set.emplace_back(physicalMemory);

#ifdef _DEBUG
        std::print("{}\t", W2A(physicalMemory.name));
        std::print("{}\t", W2A(physicalMemory.manufacturer));
        std::print("{}\t", W2A(physicalMemory.serialNumber));
        std::print("{} MHz\t", physicalMemory.speed);
        std::println("{} GB\t", physicalMemory.size);
#endif

        pclsObj->Release();
    }
    if(pEnumerator) pEnumerator->Release();

    return set;
}

std::vector<WMIC_Processor> WMIC::Processor() {
    IEnumWbemClassObject *pEnumerator = ExecQuery(pCtx, L"Win32_Processor");

    // Step 7:
    // Get the data from the query in step 6

    IWbemClassObject *pclsObj = NULL;
    ULONG uReturn             = 0;

    std::vector<WMIC_Processor> set;
    WMIC_Processor processor;
    while(pEnumerator) {
        HRESULT hr = pEnumerator->Next(WBEM_INFINITE, 1, &pclsObj, &uReturn);

        if(0 == uReturn) {
            break;
        }

        VARIANT vtProp;
        VariantInit(&vtProp);

        // Get the value of the Name property
        hr = pclsObj->Get(L"Name", 0, &vtProp, 0, 0);
        if(SUCCEEDED(hr) && (V_VT(&vtProp) == VT_BSTR)) {
            processor.name = vtProp.bstrVal;
        }
        VariantClear(&vtProp);

        hr = pclsObj->Get(L"Description", 0, &vtProp, 0, 0);
        if(SUCCEEDED(hr) && (V_VT(&vtProp) == VT_BSTR)) {
            processor.desc = vtProp.bstrVal;
        }
        VariantClear(&vtProp);

        hr = pclsObj->Get(L"Manufacturer", 0, &vtProp, 0, 0);
        if(SUCCEEDED(hr) && (V_VT(&vtProp) == VT_BSTR)) {
            processor.manufacturer = vtProp.bstrVal;
        }
        VariantClear(&vtProp);

        hr = pclsObj->Get(L"NumberOfCores", 0, &vtProp, 0, 0);
        if(SUCCEEDED(hr) && (V_VT(&vtProp) == VT_I4)) {
            processor.numberOfCores = vtProp.uintVal;
        }
        VariantClear(&vtProp);

        hr = pclsObj->Get(L"ThreadCount", 0, &vtProp, 0, 0);
        if(SUCCEEDED(hr) && (V_VT(&vtProp) == VT_I4)) {
            processor.threadCount = vtProp.uintVal;
        }
        VariantClear(&vtProp);

        hr = pclsObj->Get(L"ProcessorId", 0, &vtProp, 0, 0);
        if(SUCCEEDED(hr) && (V_VT(&vtProp) == VT_BSTR)) {
            processor.processorId = vtProp.bstrVal;
        }
        VariantClear(&vtProp);

        hr = pclsObj->Get(L"MaxClockSpeed", 0, &vtProp, 0, 0);
        if(SUCCEEDED(hr) && (V_VT(&vtProp) == VT_I4)) {
            processor.maxClockSpeed = static_cast<std::double_t>(vtProp.uintVal) / 1000.0;
        }
        VariantClear(&vtProp);

        set.emplace_back(processor);

#ifdef _DEBUG
        std::print("{}\t", W2A(processor.name));
        std::print("{}\t", W2A(processor.desc));
        std::print("{}\t", W2A(processor.manufacturer));
        std::print("{}\t", processor.numberOfCores);
        std::print("{}\t", processor.threadCount);
        std::print("{}\t", W2A(processor.processorId));
        std::println("{} GHz\t", processor.maxClockSpeed);
#endif

        pclsObj->Release();
    }
    if(pEnumerator) pEnumerator->Release();

    return set;
}

std::vector<WMIC_NetworkAdapter> WMIC::NetworkAdapter() {
    IEnumWbemClassObject *pEnumerator = ExecQuery(pCtx, L"Win32_NetworkAdapter");

    // Step 7:
    // Get the data from the query in step 6

    IWbemClassObject *pclsObj = NULL;
    ULONG uReturn             = 0;

    std::vector<WMIC_NetworkAdapter> set;
    WMIC_NetworkAdapter networkAdapter;
    while(pEnumerator) {
        HRESULT hr = pEnumerator->Next(WBEM_INFINITE, 1, &pclsObj, &uReturn);

        if(0 == uReturn) {
            break;
        }

        VARIANT vtProp;
        VariantInit(&vtProp);

        // Get the value of the Name property
        hr = pclsObj->Get(L"Name", 0, &vtProp, 0, 0);
        if(SUCCEEDED(hr) && (V_VT(&vtProp) == VT_BSTR)) {
            networkAdapter.name = vtProp.bstrVal;
        }
        VariantClear(&vtProp);

        hr = pclsObj->Get(L"Description", 0, &vtProp, 0, 0);
        if(SUCCEEDED(hr) && (V_VT(&vtProp) == VT_BSTR)) {
            networkAdapter.desc = vtProp.bstrVal;
        }
        VariantClear(&vtProp);

        hr = pclsObj->Get(L"Manufacturer", 0, &vtProp, 0, 0);
        if(SUCCEEDED(hr) && (V_VT(&vtProp) == VT_BSTR)) {
            networkAdapter.manufacturer = vtProp.bstrVal;
        }
        VariantClear(&vtProp);

        hr = pclsObj->Get(L"MACAddress", 0, &vtProp, 0, 0);
        if(SUCCEEDED(hr) && (V_VT(&vtProp) == VT_BSTR)) {
            networkAdapter.macAddress = vtProp.bstrVal;
        }
        VariantClear(&vtProp);

        hr = pclsObj->Get(L"AdapterType", 0, &vtProp, 0, 0);
        if(SUCCEEDED(hr) && (V_VT(&vtProp) == VT_BSTR)) {
            networkAdapter.adapterType = vtProp.bstrVal;
        }
        VariantClear(&vtProp);

        set.emplace_back(networkAdapter);

#ifdef _DEBUG
        std::print("{}\t", W2A(networkAdapter.name));
        std::print("{}\t", W2A(networkAdapter.desc));
        std::print("{}\t", W2A(networkAdapter.manufacturer));
        std::print("{}\t", W2A(networkAdapter.macAddress));
        std::println("{}\t", W2A(networkAdapter.adapterType));
#endif

        pclsObj->Release();
    }
    if(pEnumerator) pEnumerator->Release();

    return set;
}

static bool isWindowsVersionGreaterThan1803() {
    HMODULE hModule = GetModuleHandleW(L"ntdll.dll");
    if(hModule) {
        auto RtlGetVersion = reinterpret_cast<NTSTATUS(WINAPI *)(PRTL_OSVERSIONINFOW)>(GetProcAddress(hModule, "RtlGetVersion"));
        if(RtlGetVersion) {
            RTL_OSVERSIONINFOW osvi;
            ZeroMemory(&osvi, sizeof(RTL_OSVERSIONINFOW));
            osvi.dwOSVersionInfoSize = sizeof(RTL_OSVERSIONINFOW);

            if(RtlGetVersion(&osvi) == 0) {  // STATUS_SUCCESS = 0
                // Windows 10 的主版本号是10
                if(osvi.dwMajorVersion > 10) {
                    return true;
                } else if(osvi.dwMajorVersion == 10) {
                    // 版本1803的构建号是17134
                    return osvi.dwBuildNumber >= 17134;
                }
            }
        }
    }
    return false;
}

// 尝试获取必要的特权
static bool EnablePrivilege() {
    HANDLE hToken;
    TOKEN_PRIVILEGES tp;
    LUID luid;

    // 打开进程令牌
    if(!OpenProcessToken(GetCurrentProcess(), TOKEN_ADJUST_PRIVILEGES | TOKEN_QUERY, &hToken)) {
        std::println("无法打开进程令牌，错误码: {}", GetLastError());
        return false;
    }

    // 查找SE_SYSTEM_ENVIRONMENT_NAME特权的LUID
    if(!LookupPrivilegeValue(nullptr, SE_SYSTEM_ENVIRONMENT_NAME, &luid)) {
        std::println("无法查找特权值，错误码: {}", GetLastError());
        CloseHandle(hToken);
        return false;
    }

    // 设置特权信息
    tp.PrivilegeCount           = 1;
    tp.Privileges[0].Luid       = luid;
    tp.Privileges[0].Attributes = SE_PRIVILEGE_ENABLED;

    // 调整特权
    if(!AdjustTokenPrivileges(hToken, FALSE, &tp, sizeof(TOKEN_PRIVILEGES), nullptr, nullptr)) {
        std::println("无法调整令牌特权，错误码: {}", GetLastError());
        CloseHandle(hToken);
        return false;
    }

    // 即使AdjustTokenPrivileges返回true，也要检查GetLastError()
    DWORD error = GetLastError();
    if(error != ERROR_SUCCESS) {
        std::println("设置特权失败，错误码: {}", error);
        CloseHandle(hToken);
        return false;
    }

    CloseHandle(hToken);
    return true;
}

static inline bool canFirmwareEnvironmentVariable() {
    if(!isWindowsVersionGreaterThan1803()) return false;

    if(!EnablePrivilege()) {
        std::println("该功能需要以管理员权限运行");
        return false;
    }

    return true;
}

// 便捷的辅助函数，用于指定UID创建变量名
static std::string makeUefiVarName(uint32_t uid, const std::string &prefix = "WMIC_UID") {
    // return std::format("{}_{:08X}", prefix, uid);
    return prefix;
}

bool wmicUefiRead(uint32_t uid, void *data, int len) {
    if(!canFirmwareEnvironmentVariable()) return false;

    const auto key = makeUefiVarName(uid);
    bool ret       = GetFirmwareEnvironmentVariableA(key.c_str(), "{00000000-0000-0000-0000-000000000000}", data, len);
#ifdef _DEBUG
    if(!ret) {
        std::println("Error:{}", GetLastError());
    }
#endif
    return ret;
}

bool wmicUefiWrite(uint32_t uid, void *data, int len) {
    if(!canFirmwareEnvironmentVariable()) return false;

    const auto key = makeUefiVarName(uid);
    bool ret       = SetFirmwareEnvironmentVariableA(key.c_str(), "{00000000-0000-0000-0000-000000000000}", data, len);
#ifdef _DEBUG
    if(!ret) {
        std::println("Error:{}", GetLastError());
    }

#endif
    return ret;
}

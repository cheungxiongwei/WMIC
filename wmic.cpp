#include "wmic.h"

//WMIC_ERROR start
WMIC_ERROR::WMIC_ERROR(const char *message):m_message(message)
{

}

WMIC_ERROR::WMIC_ERROR(const std::string message):WMIC_ERROR(message.data())
{

}

const char *WMIC_ERROR::what() const
{
    return m_message.data();
}
//WMIC_ERROR end

WMIC::WMIC()
{
    //Step 1
    //Initialize COM
    hResult = CoInitializeEx(0, COINIT_MULTITHREADED);
    if (FAILED(hResult)) {
        message.clear();
        message << "Failed to initialize COM library. Error code = 0x" << std::hex << hResult << std::endl;
        throw WMIC_ERROR(message.str());
    }

    //Step 2
    //Set general COM security levels
    hResult = CoInitializeSecurity(
                NULL,
                -1,                          // COM authentication
                NULL,                        // Authentication services
                NULL,                        // Reserved
                RPC_C_AUTHN_LEVEL_DEFAULT,   // Default authentication
                RPC_C_IMP_LEVEL_IMPERSONATE, // Default Impersonation
                NULL,                        // Authentication info
                EOAC_NONE,                   // Additional capabilities
                NULL                         // Reserved
                );
    if (FAILED(hResult))
    {
        message.clear();
        message << "Failed to initialize security. Error code = 0x" << std::hex << hResult << std::endl;
        CoUninitialize();
        throw WMIC_ERROR(message.str());
        // Program has failed.
    }

    //Step 3
    //Obtain the initial locator to WMI
    hResult = CoCreateInstance(
                CLSID_WbemLocator,
                0,
                CLSCTX_INPROC_SERVER,
                IID_IWbemLocator, (LPVOID *)&pLoc);

    if (FAILED(hResult))
    {
        message.clear();
        message << "Failed to create IWbemLocator object." << " Err code = 0x" << std::hex << hResult << std::endl;
        CoUninitialize();
        throw WMIC_ERROR(message.str());
        // Program has failed.
    }

    //Step 4
    //Connect to WMI through the IWbemLocator::ConnectServer method

    // Connect to the root\cimv2 namespace with
    // the current user and obtain pointer pSvc
    // to make IWbemServices calls.
    hResult = pLoc->ConnectServer(
                _bstr_t(L"ROOT\\CIMV2"), // Object path of WMI namespace
                NULL,                    // User name. NULL = current user
                NULL,                    // User password. NULL = current
                0,                       // Locale. NULL indicates current
                NULL,                    // Security flags.
                0,                       // Authority (for example, Kerberos)
                0,                       // Context object
                &pSvc                    // pointer to IWbemServices proxy
                );

    if (FAILED(hResult))
    {
        message.clear();
        message << "Could not connect. Error code = 0x" << std::hex << hResult << std::endl;
        pLoc->Release();
        CoUninitialize();
        throw WMIC_ERROR(message.str());
        // Program has failed.
    }

    std::cout << "Connected to ROOT\\CIMV2 WMI namespace" << std::endl;

    //Step 5
    //Set security levels on the proxy
    hResult = CoSetProxyBlanket(
                pSvc,                        // Indicates the proxy to set
                RPC_C_AUTHN_WINNT,           // RPC_C_AUTHN_xxx
                RPC_C_AUTHZ_NONE,            // RPC_C_AUTHZ_xxx
                NULL,                        // Server principal name
                RPC_C_AUTHN_LEVEL_CALL,      // RPC_C_AUTHN_LEVEL_xxx
                RPC_C_IMP_LEVEL_IMPERSONATE, // RPC_C_IMP_LEVEL_xxx
                NULL,                        // client identity
                EOAC_NONE                    // proxy capabilities
                );

    if (FAILED(hResult))
    {
        message.clear();
        message << "Could not set proxy blanket. Error code = 0x" << std::hex << hResult << std::endl;
        pSvc->Release();
        pLoc->Release();
        CoUninitialize();
        throw WMIC_ERROR(message.str());
        // Program has failed.
    }

    //Step 6:
    //Use the IWbemServices pointer to make requests of WMI
}

WMIC::~WMIC()
{
    //Cleanup
    pSvc->Release();
    pLoc->Release();
    CoUninitialize();
    // Program successfully completed.
}

WMIC_OperatingSystem WMIC::OperatingSystem()
{
    IEnumWbemClassObject* pEnumerator = ExecQuery(L"Win32_OperatingSystem");

    // Step 7:
    // Get the data from the query in step 6

    IWbemClassObject *pclsObj = NULL;
    ULONG uReturn = 0;

    WMIC_OperatingSystem operatingSystem;
    while (pEnumerator){
        HRESULT hr = pEnumerator->Next(WBEM_INFINITE, 1,&pclsObj, &uReturn);

        if (0 == uReturn) {
            break;
        }

        VARIANT vtProp;

        hr = pclsObj->Get(L"Caption", 0, &vtProp, 0, 0);
        if(SUCCEEDED(hr) && (V_VT(&vtProp) == VT_BSTR)){
            operatingSystem.name = vtProp.bstrVal;
        }
        VariantClear(&vtProp);

        hr = pclsObj->Get(L"BuildNumber", 0, &vtProp, 0, 0);
        if(SUCCEEDED(hr) && (V_VT(&vtProp) == VT_BSTR)){
            operatingSystem.buildNumber = vtProp.bstrVal;
        }
        VariantClear(&vtProp);

        hr = pclsObj->Get(L"Version", 0, &vtProp, 0, 0);
        if(SUCCEEDED(hr) && (V_VT(&vtProp) == VT_BSTR)){
            operatingSystem.version = vtProp.bstrVal;
        }
        VariantClear(&vtProp);

        hr = pclsObj->Get(L"InstallDate", 0, &vtProp, 0, 0);
        if(SUCCEEDED(hr) && (V_VT(&vtProp) == VT_BSTR)){
            operatingSystem.installDate = vtProp.bstrVal;
        }
        VariantClear(&vtProp);

        hr = pclsObj->Get(L"OSArchitecture", 0, &vtProp, 0, 0);
        if(SUCCEEDED(hr) && (V_VT(&vtProp) == VT_BSTR)){
            operatingSystem.osArchitecture = vtProp.bstrVal;
        }
        VariantClear(&vtProp);

        hr = pclsObj->Get(L"RegisteredUser", 0, &vtProp, 0, 0);
        if(SUCCEEDED(hr) && (V_VT(&vtProp) == VT_BSTR)){
            operatingSystem.registeredUser = vtProp.bstrVal;
        }
        VariantClear(&vtProp);

        hr = pclsObj->Get(L"SerialNumber", 0, &vtProp, 0, 0);
        if(SUCCEEDED(hr) && (V_VT(&vtProp) == VT_BSTR)){
            operatingSystem.serialNumber = vtProp.bstrVal;
        }
        VariantClear(&vtProp);

#ifdef _DEBUG
        std::wcout.imbue(std::locale("chs"));
        std::wcout
                << operatingSystem.name << "\t"
                << operatingSystem.buildNumber << "\t"
                << operatingSystem.version << "\t"
                << operatingSystem.installDate << "\t"
                << operatingSystem.osArchitecture << "\t"
                << operatingSystem.registeredUser << "\t"
                << operatingSystem.serialNumber << "\t"
                << std::endl;
#endif

        pclsObj->Release();
    }
    pEnumerator->Release();
    return operatingSystem;
}

std::vector<WMIC_VideoController> WMIC::VideoController()
{
    IEnumWbemClassObject* pEnumerator = ExecQuery(L"Win32_VideoController");

    // Step 7:
    // Get the data from the query in step 6

    IWbemClassObject *pclsObj = NULL;
    ULONG uReturn = 0;

    std::vector<WMIC_VideoController> set;
    WMIC_VideoController videoController;
    while (pEnumerator){
        HRESULT hr = pEnumerator->Next(WBEM_INFINITE, 1,&pclsObj, &uReturn);

        if (0 == uReturn){
            break;
        }

        VARIANT vtProp;

        // Get the value of the Name property
        hr = pclsObj->Get(L"Name", 0, &vtProp, 0, 0);
        if(SUCCEEDED(hr) && (V_VT(&vtProp) == VT_BSTR)){
            videoController.name = vtProp.bstrVal;
        }
        VariantClear(&vtProp);

        hr = pclsObj->Get(L"PNPDeviceID", 0, &vtProp, 0, 0);
        if(SUCCEEDED(hr) && (V_VT(&vtProp) == VT_BSTR)){
            videoController.deviceID = vtProp.bstrVal;
        }
        VariantClear(&vtProp);

        set.push_back(videoController);

#ifdef _DEBUG
        std::wcout
                << videoController.name << "\t"
                << videoController.deviceID << "\t"
                << std::endl;
#endif
        pclsObj->Release();
    }
    pEnumerator->Release();

    return set;
}

std::vector<WMIC_DiskDrive> WMIC::DiskDrive()
{
    IEnumWbemClassObject* pEnumerator = ExecQuery(L"Win32_DiskDrive");

    // Step 7:
    // Get the data from the query in step 6

    IWbemClassObject *pclsObj = NULL;
    ULONG uReturn = 0;


    std::vector<WMIC_DiskDrive> set;
    WMIC_DiskDrive diskDrive;
    while (pEnumerator){
        HRESULT hr = pEnumerator->Next(WBEM_INFINITE, 1,&pclsObj, &uReturn);

        if (0 == uReturn){
            break;
        }

        VARIANT vtProp;

        // Get the value of the Name property
        hr = pclsObj->Get(L"Caption", 0, &vtProp, 0, 0);
        if(SUCCEEDED(hr) && (V_VT(&vtProp) == VT_BSTR)){
            diskDrive.name = vtProp.bstrVal;
        }
        VariantClear(&vtProp);

        hr = pclsObj->Get(L"PNPDeviceID", 0, &vtProp, 0, 0);
        if(SUCCEEDED(hr) && (V_VT(&vtProp) == VT_BSTR)){
            diskDrive.deviceID = vtProp.bstrVal;
        }
        VariantClear(&vtProp);

        hr = pclsObj->Get(L"SerialNumber", 0, &vtProp, 0, 0);
        if(SUCCEEDED(hr) && (V_VT(&vtProp) == VT_BSTR)){
            diskDrive.serialNumber = vtProp.bstrVal;
        }
        VariantClear(&vtProp);

        hr = pclsObj->Get(L"Size", 0, &vtProp, 0, 0);
        if(SUCCEEDED(hr) && (V_VT(&vtProp) == VT_BSTR)){
            VarUI8FromStr(vtProp.bstrVal,NULL,LOCALE_NOUSEROVERRIDE,&diskDrive.size);
            diskDrive.size = diskDrive.size / (1024*1024*1024);//1 byte / (1024*1024*1024) = x GB
        }
        VariantClear(&vtProp);

        set.push_back(diskDrive);
#ifdef _DEBUG
        std::wcout
                << diskDrive.name << "\t"
                << diskDrive.deviceID << "\t"
                << diskDrive.serialNumber << "\t"
                << diskDrive.size << "GB"
                << std::endl;
#endif
        pclsObj->Release();
    }
    pEnumerator->Release();

    return set;
}

WMIC_BaseBoard WMIC::BaseBoard()
{
    IEnumWbemClassObject* pEnumerator = ExecQuery(L"Win32_BaseBoard");

    // Step 7:
    // Get the data from the query in step 6

    IWbemClassObject *pclsObj = NULL;
    ULONG uReturn = 0;


    WMIC_BaseBoard baseBoard;
    while (pEnumerator){
        HRESULT hr = pEnumerator->Next(WBEM_INFINITE, 1,&pclsObj, &uReturn);

        if (0 == uReturn){
            break;
        }

        VARIANT vtProp;

        // Get the value of the Name property
        hr = pclsObj->Get(L"Name", 0, &vtProp, 0, 0);
        if(SUCCEEDED(hr) && (V_VT(&vtProp) == VT_BSTR)){
            baseBoard.name = vtProp.bstrVal;
        }
        VariantClear(&vtProp);

        hr = pclsObj->Get(L"Manufacturer", 0, &vtProp, 0, 0);
        if(SUCCEEDED(hr) && (V_VT(&vtProp) == VT_BSTR)){
            baseBoard.manufacturer = vtProp.bstrVal;
        }
        VariantClear(&vtProp);

        hr = pclsObj->Get(L"Product", 0, &vtProp, 0, 0);
        if(SUCCEEDED(hr) && (V_VT(&vtProp) == VT_BSTR)){
            baseBoard.product = vtProp.bstrVal;
        }
        VariantClear(&vtProp);

        hr = pclsObj->Get(L"Version", 0, &vtProp, 0, 0);
        if(SUCCEEDED(hr) && (V_VT(&vtProp) == VT_BSTR)){
            baseBoard.version = vtProp.bstrVal;
        }
        VariantClear(&vtProp);

        hr = pclsObj->Get(L"SerialNumber", 0, &vtProp, 0, 0);
        if(SUCCEEDED(hr) && (V_VT(&vtProp) == VT_BSTR)){
            baseBoard.serialNumber = vtProp.bstrVal;
        }
        VariantClear(&vtProp);

#ifdef _DEBUG
        std::wcout.imbue(std::locale("chs"));
        std::wcout
                << baseBoard.name << "\t"
                << baseBoard.manufacturer << "\t"
                << baseBoard.product << "\t"
                << baseBoard.version << "\t"
                << baseBoard.serialNumber
                << std::endl;
#endif

        pclsObj->Release();
    }
    pEnumerator->Release();

    return baseBoard;
}

WMIC_BIOS WMIC::BIOS()
{
    IEnumWbemClassObject* pEnumerator = ExecQuery(L"Win32_BIOS");

    // Step 7:
    // Get the data from the query in step 6

    IWbemClassObject *pclsObj = NULL;
    ULONG uReturn = 0;


    WMIC_BIOS bios;
    while (pEnumerator){
        HRESULT hr = pEnumerator->Next(WBEM_INFINITE, 1,&pclsObj, &uReturn);

        if (0 == uReturn){
            break;
        }

        VARIANT vtProp;

        // Get the value of the Name property
        hr = pclsObj->Get(L"ReleaseDate", 0, &vtProp, 0, 0);
        if(SUCCEEDED(hr) && (V_VT(&vtProp) == VT_BSTR)){
            bios.releaseDate = vtProp.bstrVal;
        }
        VariantClear(&vtProp);

        hr = pclsObj->Get(L"Manufacturer", 0, &vtProp, 0, 0);
        if(SUCCEEDED(hr) && (V_VT(&vtProp) == VT_BSTR)){
            bios.manufacturer = vtProp.bstrVal;
        }
        VariantClear(&vtProp);

        hr = pclsObj->Get(L"SerialNumber", 0, &vtProp, 0, 0);
        if(SUCCEEDED(hr) && (V_VT(&vtProp) == VT_BSTR)){
            bios.serialNumber = vtProp.bstrVal;
        }
        VariantClear(&vtProp);

#ifdef _DEBUG
        std::wcout.imbue(std::locale("chs"));
        std::wcout
                << bios.manufacturer << "\t"
                << bios.releaseDate << "\t"
                << bios.serialNumber
                << std::endl;
#endif

        pclsObj->Release();
    }
    pEnumerator->Release();

    return bios;
}

std::vector<WMIC_PhysicalMemory> WMIC::PhysicalMemory()
{
    IEnumWbemClassObject* pEnumerator = ExecQuery(L"Win32_PhysicalMemory");

    // Step 7:
    // Get the data from the query in step 6

    IWbemClassObject *pclsObj = NULL;
    ULONG uReturn = 0;


    std::vector<WMIC_PhysicalMemory> set;
    WMIC_PhysicalMemory physicalMemory;
    while (pEnumerator){
        HRESULT hr = pEnumerator->Next(WBEM_INFINITE, 1,&pclsObj, &uReturn);

        if (0 == uReturn){
            break;
        }

        VARIANT vtProp;

        // Get the value of the Name property
        hr = pclsObj->Get(L"Name", 0, &vtProp, 0, 0);
        if(SUCCEEDED(hr) && (V_VT(&vtProp) == VT_BSTR)){
            physicalMemory.name = vtProp.bstrVal;
        }
        VariantClear(&vtProp);

        hr = pclsObj->Get(L"Manufacturer", 0, &vtProp, 0, 0);
        if(SUCCEEDED(hr) && (V_VT(&vtProp) == VT_BSTR)){
            physicalMemory.manufacturer = vtProp.bstrVal;
        }
        VariantClear(&vtProp);

        hr = pclsObj->Get(L"SerialNumber", 0, &vtProp, 0, 0);
        if(SUCCEEDED(hr) && (V_VT(&vtProp) == VT_BSTR)){
            physicalMemory.serialNumber = vtProp.bstrVal;
        }
        VariantClear(&vtProp);

        hr = pclsObj->Get(L"Speed", 0, &vtProp, 0, 0);
        if(SUCCEEDED(hr) && (V_VT(&vtProp) == VT_I4)){
            physicalMemory.speed = vtProp.uintVal;
        }
        VariantClear(&vtProp);

        hr = pclsObj->Get(L"Capacity", 0, &vtProp, 0, 0);
        if(SUCCEEDED(hr) && (V_VT(&vtProp) == VT_BSTR)){
            VarUI8FromStr(vtProp.bstrVal,NULL,LOCALE_NOUSEROVERRIDE,&physicalMemory.size);
            physicalMemory.size = physicalMemory.size / (1024*1024*1024);//1 byte / (1024*1024*1024) = x GB
        }
        VariantClear(&vtProp);

        set.push_back(physicalMemory);

#ifdef _DEBUG
        std::wcout.imbue(std::locale("chs"));
        std::wcout
                << physicalMemory.name << "\t"
                << physicalMemory.manufacturer << "\t"
                << physicalMemory.serialNumber << "\t"
                << physicalMemory.speed << "MHz" << "\t"
                << physicalMemory.size << "GB"
                << std::endl;
#endif

        pclsObj->Release();
    }
    pEnumerator->Release();

    return set;
}

std::vector<WMIC_Processor> WMIC::Processor()
{
    IEnumWbemClassObject* pEnumerator = ExecQuery(L"Win32_Processor");

    // Step 7:
    // Get the data from the query in step 6

    IWbemClassObject *pclsObj = NULL;
    ULONG uReturn = 0;


    std::vector<WMIC_Processor> set;
    WMIC_Processor processor;
    while (pEnumerator){
        HRESULT hr = pEnumerator->Next(WBEM_INFINITE, 1,&pclsObj, &uReturn);

        if (0 == uReturn){
            break;
        }

        VARIANT vtProp;

        // Get the value of the Name property
        hr = pclsObj->Get(L"Name", 0, &vtProp, 0, 0);
        if(SUCCEEDED(hr) && (V_VT(&vtProp) == VT_BSTR)){
            processor.name = vtProp.bstrVal;
        }
        VariantClear(&vtProp);

        hr = pclsObj->Get(L"Description", 0, &vtProp, 0, 0);
        if(SUCCEEDED(hr) && (V_VT(&vtProp) == VT_BSTR)){
            processor.desc = vtProp.bstrVal;
        }
        VariantClear(&vtProp);

        hr = pclsObj->Get(L"Manufacturer", 0, &vtProp, 0, 0);
        if(SUCCEEDED(hr) && (V_VT(&vtProp) == VT_BSTR)){
            processor.manufacturer = vtProp.bstrVal;
        }
        VariantClear(&vtProp);

        hr = pclsObj->Get(L"NumberOfCores", 0, &vtProp, 0, 0);
        if(SUCCEEDED(hr) && (V_VT(&vtProp) == VT_I4)){
            processor.numberOfCores = vtProp.uintVal;
        }
        VariantClear(&vtProp);

        hr = pclsObj->Get(L"ThreadCount", 0, &vtProp, 0, 0);
        if(SUCCEEDED(hr) && (V_VT(&vtProp) == VT_I4)){
            processor.threadCount = vtProp.uintVal;
        }
        VariantClear(&vtProp);

        hr = pclsObj->Get(L"ProcessId", 0, &vtProp, 0, 0);
        if(SUCCEEDED(hr) && (V_VT(&vtProp) == VT_BSTR)){
            processor.processID = vtProp.bstrVal;
        }
        VariantClear(&vtProp);

        hr = pclsObj->Get(L"MaxClockSpeed", 0, &vtProp, 0, 0);
        if(SUCCEEDED(hr) && (V_VT(&vtProp) == VT_I4)){
            processor.maxClockSpeed = static_cast<std::double_t>(vtProp.uintVal) / 1000.0;
        }
        VariantClear(&vtProp);

        set.push_back(processor);

#ifdef _DEBUG
        std::wcout.imbue(std::locale("chs"));
        std::wcout
                << processor.name << "\t"
                << processor.desc << "\t"
                << processor.manufacturer << "\t"
                << processor.numberOfCores << "\t"
                << processor.threadCount << "\t"
                << processor.processID << "\t"
                << processor.maxClockSpeed << "GHz"
                << std::endl;
#endif

        pclsObj->Release();
    }
    pEnumerator->Release();

    return set;
}

std::vector<WMIC_NetworkAdapter> WMIC::NetworkAdapter()
{
    IEnumWbemClassObject* pEnumerator = ExecQuery(L"Win32_NetworkAdapter");

    // Step 7:
    // Get the data from the query in step 6

    IWbemClassObject *pclsObj = NULL;
    ULONG uReturn = 0;


    std::vector<WMIC_NetworkAdapter> set;
    WMIC_NetworkAdapter networkAdapter;
    while (pEnumerator){
        HRESULT hr = pEnumerator->Next(WBEM_INFINITE, 1,&pclsObj, &uReturn);

        if (0 == uReturn){
            break;
        }

        VARIANT vtProp;
        VariantInit(&vtProp);

        // Get the value of the Name property
        hr = pclsObj->Get(L"Name", 0, &vtProp, 0, 0);
        if(SUCCEEDED(hr) && (V_VT(&vtProp) == VT_BSTR)){
            networkAdapter.name = vtProp.bstrVal;
        }
        VariantClear(&vtProp);

        hr = pclsObj->Get(L"Description", 0, &vtProp, 0, 0);
        if(SUCCEEDED(hr) && (V_VT(&vtProp) == VT_BSTR)){
            networkAdapter.desc = vtProp.bstrVal;
        }
        VariantClear(&vtProp);

        hr = pclsObj->Get(L"Manufacturer", 0, &vtProp, 0, 0);
        if(SUCCEEDED(hr) && (V_VT(&vtProp) == VT_BSTR)){
            networkAdapter.manufacturer = vtProp.bstrVal;
        }
        VariantClear(&vtProp);

        hr = pclsObj->Get(L"MACAddress", 0, &vtProp, 0, 0);
        if(SUCCEEDED(hr) && (V_VT(&vtProp) == VT_BSTR)){
            networkAdapter.macAddress = vtProp.bstrVal;
        }
        VariantClear(&vtProp);

        hr = pclsObj->Get(L"AdapterType", 0, &vtProp, 0, 0);
        if(SUCCEEDED(hr) && (V_VT(&vtProp) == VT_BSTR)){
            networkAdapter.adapterType = vtProp.bstrVal;
        }
        VariantClear(&vtProp);

        set.push_back(networkAdapter);

#ifdef _DEBUG
        std::wcout.imbue(std::locale("chs"));
        std::wcout
                << networkAdapter.name << "\t"
                << networkAdapter.desc << "\t"
                << networkAdapter.manufacturer << "\t"
                << networkAdapter.macAddress << "\t"
                << networkAdapter.adapterType << "\t"
                << std::endl;
#endif

        pclsObj->Release();
    }
    pEnumerator->Release();

    return set;
}

IEnumWbemClassObject *WMIC::ExecQuery(const std::wstring className)
{
    IEnumWbemClassObject* pEnumerator = NULL;
    std::wstring cmd = L"SELECT * FROM ";
    cmd += className;
    //std::wcout << cmd << std::endl;

    hResult = pSvc->ExecQuery(
                bstr_t("WQL"),
                bstr_t(cmd.data()),
                WBEM_FLAG_FORWARD_ONLY | WBEM_FLAG_RETURN_IMMEDIATELY,
                NULL,
                &pEnumerator);

    if (FAILED(hResult))
    {
        message.clear();
        message << "Query for xxx failed." << " Error code = 0x" << std::hex << hResult << std::endl;
        pSvc->Release();
        pLoc->Release();
        CoUninitialize();
        throw WMIC_ERROR(message.str());
        // Program has failed.
    }
    return pEnumerator;
}

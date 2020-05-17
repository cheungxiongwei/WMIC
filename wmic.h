#ifndef WMIC_H
#define WMIC_H

#include <Windows.h>
#include <comdef.h>
#include <Wbemidl.h>

//stl
#include <iostream>
#include <string>
#include <vector>
#include <map>
#include <sstream>

#pragma comment(lib, "wbemuuid.lib")

//WMIC_ERROR start
class WMIC_ERROR : public std::exception
{
public:
    WMIC_ERROR(const char * message);
    WMIC_ERROR(const std::string message);
    virtual char const * what() const;
private:
    std::string m_message;
};
//WMIC_ERROR end

/*!
WMI所有类别
https://docs.microsoft.com/en-us/previous-versions//aa394084(v=vs.85)?redirectedfrom=MSDN
https://docs.microsoft.com/en-us/windows/win32/wmisdk/wmi-start-page

wmi 测试器
wbemtest.exe
SELECT * FROM [ClassName]
SELECT * FROM Win32_VideoController
*/

//显卡 Win32_VideoController
struct WMIC_VideoController
{
    std::wstring name;
    std::wstring deviceID;
};

//硬盘 Win32_DiskDrive
struct WMIC_DiskDrive
{
    std::wstring name;
    std::wstring deviceID;
    std::wstring serialNumber;
    std::uint64_t size = 0;//GB
};

//主板 Win32_BaseBoard
struct WMIC_BaseBoard
{
    std::wstring name;
    std::wstring manufacturer;//生产厂商
    std::wstring product;//产品
    std::wstring version;//版本
    std::wstring serialNumber;
};

//主板BIOS Win32_BIOS
struct WMIC_BIOS
{
    std::wstring manufacturer;//生产厂商
    std::wstring releaseDate;//发布日期
    std::wstring serialNumber;
};

//内存 Win32_PhysicalMemory
struct WMIC_PhysicalMemory
{
    std::wstring name;
    std::wstring manufacturer;//生产厂商
    std::wstring serialNumber;//序列号
    std::uint32_t speed;//频率 MHz(兆赫)
    std::uint64_t size;//容量 GB
};

//CPU Win32_Processor
struct WMIC_Processor
{
    std::wstring name;
    std::wstring desc;//Description
    std::wstring manufacturer;//生产厂商
    std::uint32_t numberOfCores;//CPU核心数量
    std::wstring processID;//处理器ID
    std::uint32_t threadCount;//线程数量
    std::double_t maxClockSpeed;//最大时钟频率 GHz
};

//串口 Win32_SerialPort

//声卡 Win32_SoundDevice

//网络适配器 Win32_NetworkAdapter
struct WMIC_NetworkAdapter
{
    std::wstring name;
    std::wstring desc;//Description
    std::wstring manufacturer;//生产厂商
    std::wstring macAddress;
    std::wstring adapterType;
};

//系统信息 Win32_OperatingSystem
struct WMIC_OperatingSystem
{
    std::wstring name;
    std::wstring buildNumber;
    std::wstring version;
    std::wstring installDate;
    std::wstring osArchitecture;
    std::wstring registeredUser;
    std::wstring serialNumber;
};

class WMIC
{
public:
    WMIC();
    ~WMIC();
    void sysinfo();
    WMIC_OperatingSystem OperatingSystem();//系统
    std::vector<WMIC_VideoController> VideoController();//显卡
    std::vector<WMIC_DiskDrive> DiskDrive();//硬盘
    WMIC_BaseBoard BaseBoard();//主板
    WMIC_BIOS BIOS();
    std::vector<WMIC_PhysicalMemory> PhysicalMemory();//内存条
    std::vector<WMIC_Processor> Processor();//cpu处理器
    std::vector<WMIC_NetworkAdapter> NetworkAdapter();//网络适配器
private:
    IEnumWbemClassObject *ExecQuery(const std::wstring className);
private:
    HRESULT hResult;
    std::stringstream message;
    IWbemLocator *pLoc = NULL;
    IWbemServices *pSvc = NULL;
private:
    WMIC(const WMIC&) = delete;
    WMIC& operator = (const WMIC&) = delete;
};



#endif // WMIC_H

#pragma once
#include <cmath>
#include <memory>
#include <string>
#include <string_view>
#include <vector>

std::string ConvertWideToMultiByte(std::wstring_view str);

std::wstring ConvertMultiByteToWide(std::string_view str);

constexpr auto &W2A = ConvertWideToMultiByte;
constexpr auto &A2W = ConvertMultiByteToWide;

/*!
WMI所有类别
https://docs.microsoft.com/en-us/previous-versions//aa394084(v=vs.85)?redirectedfrom=MSDN
https://docs.microsoft.com/en-us/windows/win32/wmisdk/wmi-start-page

UEFI Non-Volatile Random-Access Memory(NVRAM)
https://learn.microsoft.com/en-us/windows/win32/sysinfo/access-uefi-firmware-variables-from-a-universal-windows-app

wmi Windows自带测试程序
C:\Windows\System32\wbem\wbemtest.exe
SELECT * FROM [ClassName]
SELECT * FROM Win32_VideoController
*/

// 显卡 Win32_VideoController
struct WMIC_VideoController {
    std::wstring name;
    std::wstring deviceID;
};

// 硬盘 Win32_DiskDrive
struct WMIC_DiskDrive {
    std::wstring name;
    std::wstring deviceID;
    std::wstring serialNumber;
    std::uint64_t size = 0;  // GB
};

// 主板 Win32_BaseBoard
struct WMIC_BaseBoard {
    std::wstring name;
    std::wstring manufacturer;  // 生产厂商
    std::wstring product;       // 产品
    std::wstring version;       // 版本
    std::wstring serialNumber;
};

// 主板BIOS Win32_BIOS
struct WMIC_BIOS {
    std::wstring manufacturer;  // 生产厂商
    std::wstring releaseDate;   // 发布日期
    std::wstring serialNumber;
};

// 内存 Win32_PhysicalMemory
struct WMIC_PhysicalMemory {
    std::wstring name;
    std::wstring manufacturer;  // 生产厂商
    std::wstring serialNumber;  // 序列号
    std::uint32_t speed;        // 频率 MHz(兆赫)
    std::uint64_t size;         // 容量 GB
};

// CPU Win32_Processor
struct WMIC_Processor {
    std::wstring name;
    std::wstring desc;            // Description
    std::wstring manufacturer;    // 生产厂商
    std::uint32_t numberOfCores;  // CPU核心数量
    std::wstring processorId;     // 处理器ID
    std::uint32_t threadCount;    // 线程数量
    std::double_t maxClockSpeed;  // 最大时钟频率 GHz
};

// 串口 Win32_SerialPort

// 声卡 Win32_SoundDevice

// 网络适配器 Win32_NetworkAdapter
struct WMIC_NetworkAdapter {
    std::wstring name;
    std::wstring desc;          // Description
    std::wstring manufacturer;  // 生产厂商
    std::wstring macAddress;
    std::wstring adapterType;
};

// 系统信息 Win32_OperatingSystem
struct WMIC_OperatingSystem {
    std::wstring name;
    std::wstring buildNumber;
    std::wstring version;
    std::wstring installDate;
    std::wstring osArchitecture;
    std::wstring registeredUser;
    std::wstring serialNumber;
};

class WMICException : public std::exception
{
public:
    explicit WMICException(const std::string &message)
      : mMessage(message) {}

    const char *what() const noexcept override { return mMessage.data(); }

private:
    std::string mMessage;
};

inline uint32_t uid(const std::vector<uint8_t> &data) {
    uint32_t id {0};
    for(const auto &v: data) {
        id = (id << 8) | v;
    }
    id &= 0x1fffffff;
    return id;
}

class WMIC
{
public:
    explicit WMIC();
    ~WMIC();
    WMIC_OperatingSystem OperatingSystem();               // 系统
    std::vector<WMIC_VideoController> VideoController();  // 显卡
    std::vector<WMIC_DiskDrive> DiskDrive();              // 硬盘
    WMIC_BaseBoard BaseBoard();                           // 主板
    WMIC_BIOS BIOS();                                     // 主板BIOS
    std::vector<WMIC_PhysicalMemory> PhysicalMemory();    // 内存条
    std::vector<WMIC_Processor> Processor();              // cpu处理器
    std::vector<WMIC_NetworkAdapter> NetworkAdapter();    // 网络适配器

private:
    std::unique_ptr<struct WMICContext> pCtx;

private:
    WMIC(const WMIC &)            = delete;
    WMIC &operator=(const WMIC &) = delete;
};

typedef struct alignas(sizeof(void *)) uid_meta_t {
    uint32_t uid;                // 用户标识符/许可证唯一ID
    uint8_t license_key[32];     // 许可证密钥
    uint64_t issue_date;         // 许可证发布日期（UNIX时间戳）
    uint64_t expiry_date;        // 许可证过期日期（UNIX时间戳）
    uint32_t version;            // 许可证版本号
    uint32_t features_mask;      // 功能权限位掩码
    uint16_t max_devices;        // 允许激活的最大设备数量
    uint16_t current_devices;    // 当前已激活的设备数量
    uint8_t name[64];            // 客户名称
    uint8_t email[128];          // 客户电子邮件
    uint8_t license_type;        // 许可证类型（试用版、专业版、企业版等）
    uint8_t status;              // 许可状态（有效、暂停、已撤销等）
    uint8_t activation_count;    // 激活次数
    uint8_t hash_signature[64];  // 许可证数据的加密签名，用于验证完整性
    uint8_t reserved[32];        // 保留字段，用于将来扩展
} uid_meta_t;

bool wmicUefiRead(uint32_t uid, void *data, int len);

bool wmicUefiWrite(uint32_t uid, void *data, int len);

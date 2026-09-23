#pragma once
#include <cstdint>
#include <string>
#include <vector>

/*!
SMBIOS / DMI 固件表信息。

数据来源: GetSystemFirmwareTable('RSMB')，仅依赖 Win32 API，不使用 COM/WMI。
规范参考: DMTF SMBIOS Reference Specification (DSP0134)。

注意:
- 表中的字符串可能为空，或为 OEM 占位符(如 "To be filled by O.E.M.")，
  参与比较/哈希前建议先做归一化。
- System UUID 若为全 0 或全 0xFF，表示厂商未烧录，已统一归一化为空串。
- 硬件指纹应只取“换硬件才变”的字段；地址、表顺序等每次开机都会变，不可入哈希。
*/

// BIOS 信息，SMBIOS Type 0
struct SmbiosBios {
    std::string vendor;       // 固件厂商
    std::string version;      // 固件版本（刷 BIOS 会变，不建议入指纹）
    std::string releaseDate;  // 发布日期
};

// 系统信息，SMBIOS Type 1
struct SmbiosSystem {
    std::string manufacturer;  // 整机/虚拟机厂商
    std::string productName;   // 机型
    std::string version;       // 版本
    std::string serialNumber;  // 整机序列号
    std::string uuid;          // 归一化后的 UUID 字符串（小端已修正）
};

// 主板信息，SMBIOS Type 2
struct SmbiosBaseboard {
    std::string manufacturer;
    std::string product;
    std::string version;
    std::string serialNumber;
};

// 机箱信息，SMBIOS Type 3
struct SmbiosChassis {
    std::uint8_t type = 0;  // 机箱类型（SMBIOS 定义的枚举值）
    std::string manufacturer;
    std::string serialNumber;
    std::string assetTag;
};

// 处理器信息，SMBIOS Type 4
struct SmbiosProcessor {
    std::string socket;
    std::string manufacturer;
    std::string version;
    std::uint16_t coreCount   = 0;
    std::uint16_t threadCount = 0;
};

// 内存条信息，SMBIOS Type 17（仅采集已安装的内存条）
struct SmbiosMemoryDevice {
    std::string manufacturer;
    std::string serialNumber;
    std::string partNumber;
    std::uint64_t sizeMb = 0;  // 容量，单位 MB
    std::uint16_t speed  = 0;  // 速率，单位 MT/s
};

// 处理器信息，来自 CPUID 指令（非固件表）。
// CPUID 是 CPU 直接执行的指令，结果由 CPU/虚拟机监控器实时返回，
// 不经由 SMBIOS，因此比固件表字符串更难被伪造，适合做虚拟机判定。
struct CpuId {
    std::string vendor;             // leaf 0x00: 真实厂商串，如 "GenuineIntel" / "AuthenticAMD"
    std::string brand;              // leaf 0x80000002..0x80000004: 品牌串
    std::string hypervisorVendor;   // leaf 0x40000000: 虚拟机监控器厂商串，如 "VMwareVMware"
    bool hypervisorPresent = false; // leaf 0x01 的 ECX bit31: 是否存在虚拟机监控器
                                    // 注意: 开启 VBS/Hyper-V 的物理机同样为真，
                                    // 该位不能单独用于判定“本机是虚拟机”。
};

// SMBIOS 表解析结果
struct Smbios {
    SmbiosBios bios;
    SmbiosSystem system;
    SmbiosBaseboard baseboard;
    SmbiosChassis chassis;
    std::vector<SmbiosProcessor> processors;
    std::vector<SmbiosMemoryDevice> memory;
    CpuId cpu;                 // CPUID 结果（不来自固件表，供虚拟机判定交叉验证）
    std::uint8_t majorVersion = 0;  // SMBIOS 主版本
    std::uint8_t minorVersion = 0;  // SMBIOS 次版本
    bool valid                = false;  // 是否成功读到并解析了固件表
};

// 读取并解析 SMBIOS 固件表，同时采集 CPUID 信息（填充 cpu 字段）。
// 无 COM/WMI 依赖，失败时返回的 valid 为 false。
Smbios smbiosRead();

// 仅采集 CPUID 信息（不读固件表）。x86/x64 之外返回默认值。
CpuId cpuIdRead();

// 判定当前机器是否为虚拟机，采用两层交叉验证:
// 1) CPUID: 识别 VMware/VirtualBox/KVM/Xen 等虚拟机监控器厂商串（更难伪造，优先采用）；
// 2) SMBIOS: 厂商/机型/BIOS 字符串与 UUID 前缀（可能被篡改，仅作补充，
//    兼作 Hyper-V 的兜底判定）。
// 任一层命中即判定为虚拟机。
// 说明: 仅 hypervisorPresent 位为真不算虚拟机 —— 开启 VBS/Hyper-V 的物理机也会置位。
bool smbiosIsVirtualMachine(const Smbios &smbios);

// 由 SMBIOS 中的稳定字段计算 64 位 FNV-1a 硬件指纹。
// 参与因子: 系统厂商、机型、System UUID、主板序列号、BIOS 厂商。
// 易变字段(BIOS 版本、内存序列号等)不参与，避免换固件/换内存即变。
std::uint64_t smbiosFingerprint(const Smbios &smbios);

// 将指纹格式化为 16 位十六进制字符串，便于打印与持久化。
std::string smbiosFingerprintString(const Smbios &smbios);

#include "smbios.h"

#ifndef NOMINMAX
#define NOMINMAX
#endif
#ifndef WIN32_LEAN_AND_MEAN
#define WIN32_LEAN_AND_MEAN
#endif
#include <Windows.h>

#include <algorithm>
#include <array>
#include <cctype>
#include <cstring>
#include <format>
#include <string_view>

// __cpuid intrinsic 仅存在于 x86/x64 平台
#if defined(_M_X64) || defined(_M_IX86)
#include <intrin.h>
#define SMBIOS_HAS_CPUID 1
#endif

namespace {

// RawSMBIOSData 头: BYTE Used20CallingMethod, BYTE Major, BYTE Minor, BYTE DmiRevision,
// DWORD Length，其后紧接 SMBIOS 结构流。
constexpr DWORD kRawSmbiosSignature = 0x52534D42;  // 'RSMB'，用整数避免多字符字面量告警
constexpr std::size_t kRawHeaderSize = 8;

// 读取原始 SMBIOS 表。第一次调用取长度，第二次调用取内容。
std::vector<std::uint8_t> readRawSmbios() {
    const DWORD size = GetSystemFirmwareTable(kRawSmbiosSignature, 0, nullptr, 0);
    if(size == 0) return {};

    std::vector<std::uint8_t> buffer(size);
    const DWORD written = GetSystemFirmwareTable(kRawSmbiosSignature, 0, buffer.data(), size);
    if(written == 0) return {};

    buffer.resize(written);
    return buffer;
}

std::uint16_t readU16(const std::uint8_t *p) {
    std::uint16_t value = 0;
    std::memcpy(&value, p, sizeof(value));
    return value;
}

std::uint32_t readU32(const std::uint8_t *p) {
    std::uint32_t value = 0;
    std::memcpy(&value, p, sizeof(value));
    return value;
}

// 取结构中某个字段(存放“字符串索引”)指向的字符串。
// fieldOffset 为该字段相对结构起点的偏移，越界则返回空串。
std::string readFieldString(const std::uint8_t *table, std::size_t tableLength,
                            std::size_t structOffset, std::size_t structLength,
                            std::size_t fieldOffset) {
    if(fieldOffset >= structLength) return {};

    const std::uint8_t index = table[structOffset + fieldOffset];
    if(index == 0) return {};

    std::size_t p = structOffset + structLength;  // 字符串区起点
    for(std::uint8_t i = 1; i < index; ++i) {
        while(p < tableLength && table[p] != 0) ++p;
        if(p >= tableLength) return {};
        ++p;
        if(p < tableLength && table[p] == 0) return {};  // 字符串区已耗尽
    }

    const std::size_t start = p;
    while(p < tableLength && table[p] != 0) ++p;
    return std::string(reinterpret_cast<const char *>(table + start), p - start);
}

// SMBIOS 的 UUID 前三段按小端存储，需逐字节倒序后格式化。
std::string formatUuid(const std::uint8_t *b) {
    const bool allZero = std::all_of(b, b + 16, [](std::uint8_t v) { return v == 0x00; });
    const bool allFf   = std::all_of(b, b + 16, [](std::uint8_t v) { return v == 0xFF; });
    if(allZero || allFf) return {};

    return std::format("{:02X}{:02X}{:02X}{:02X}-{:02X}{:02X}-{:02X}{:02X}-"
                       "{:02X}{:02X}-{:02X}{:02X}{:02X}{:02X}{:02X}{:02X}",
                       b[3], b[2], b[1], b[0], b[5], b[4], b[7], b[6],
                       b[8], b[9], b[10], b[11], b[12], b[13], b[14], b[15]);
}

#ifdef SMBIOS_HAS_CPUID

// leaf 0x01 的 ECX bit31 = hypervisor present。
// 虚拟机监控器会主动置位该位，是“运行在虚拟机监控器之上”最直接的信号。
// 注意: Windows 的虚拟化安全(VBS)/Hyper-V 也会让物理机置位该位，
// 因此该位为真只代表“存在虚拟机监控器”，不等价于“本机是虚拟机”。
bool cpuidHypervisorPresent() {
    int regs[4] = {};
    __cpuid(regs, 0x00);
    if(static_cast<unsigned int>(regs[0]) < 1u) return false;  // 不支持 leaf 0x01

    __cpuid(regs, 0x01);
    return (regs[2] & (1 << 31)) != 0;
}

// leaf 0x00: 真实 CPU 厂商串，按 EBX/EDX/ECX 顺序拼 12 字节。
std::string cpuidVendorString() {
    int regs[4] = {};
    __cpuid(regs, 0x00);

    char buffer[13] = {};
    std::memcpy(buffer, &regs[1], 4);      // EBX
    std::memcpy(buffer + 4, &regs[3], 4);  // EDX
    std::memcpy(buffer + 8, &regs[2], 4);  // ECX
    return buffer;
}

// leaf 0x80000002..0x80000004: 处理器品牌串，共 48 字节。
std::string cpuidBrandString() {
    int regs[4] = {};
    __cpuid(regs, 0x80000000);
    // EAX 为扩展叶的最大编号，品牌串要求至少支持到 0x80000004
    if(static_cast<unsigned int>(regs[0]) < 0x80000004u) return {};

    char buffer[49] = {};
    for(int leaf = 0; leaf < 3; ++leaf) {
        __cpuid(regs, 0x80000002 + leaf);
        std::memcpy(buffer + leaf * 16, regs, sizeof(int) * 4);
    }
    return buffer;
}

// leaf 0x40000000: 虚拟机监控器厂商串，按 EBX/ECX/EDX 顺序拼 12 字节。
// 该叶位于约定的 0x40000000 区间，只要 hypervisorPresent 位置位即为有效，
// 不受 leaf 0x00 返回的“最大标准叶编号”限制(该编号通常远小于 0x40000000，
// 例如 AMD 处理器一般返回 0x10 左右，据此判断会漏读厂商串)。
std::string cpuidHypervisorVendorString() {
    if(!cpuidHypervisorPresent()) return {};

    int regs[4] = {};
    __cpuid(regs, 0x40000000);
    char buffer[13] = {};
    std::memcpy(buffer, &regs[1], 4);      // EBX
    std::memcpy(buffer + 4, &regs[2], 4);  // ECX
    std::memcpy(buffer + 8, &regs[3], 4);  // EDX
    return buffer;
}

#endif  // SMBIOS_HAS_CPUID

}  // namespace

CpuId cpuIdRead() {
    CpuId cpu;
#ifdef SMBIOS_HAS_CPUID
    cpu.vendor            = cpuidVendorString();
    cpu.brand             = cpuidBrandString();
    cpu.hypervisorVendor  = cpuidHypervisorVendorString();
    cpu.hypervisorPresent = cpuidHypervisorPresent();
#endif
    return cpu;
}

Smbios smbiosRead() {
    Smbios result;

    result.cpu = cpuIdRead();

    const auto buffer = readRawSmbios();
    if(buffer.size() < kRawHeaderSize) return result;

    result.majorVersion = buffer[1];
    result.minorVersion = buffer[2];
    result.valid        = true;

    const std::uint8_t *table = buffer.data() + kRawHeaderSize;
    std::size_t tableLength   = readU32(buffer.data() + 4);
    if(tableLength > buffer.size() - kRawHeaderSize) {
        tableLength = buffer.size() - kRawHeaderSize;
    }

    for(std::size_t offset = 0; offset + 4 <= tableLength;) {
        const std::uint8_t type = table[offset];
        const std::uint8_t len  = table[offset + 1];
        if(len < 4 || offset + len > tableLength) break;

        const std::uint8_t *s = table + offset;
        switch(type) {
        case 0:  // BIOS
            result.bios.vendor      = readFieldString(table, tableLength, offset, len, 0x04);
            result.bios.version     = readFieldString(table, tableLength, offset, len, 0x05);
            result.bios.releaseDate = readFieldString(table, tableLength, offset, len, 0x08);
            break;
        case 1:  // System
            result.system.manufacturer = readFieldString(table, tableLength, offset, len, 0x04);
            result.system.productName  = readFieldString(table, tableLength, offset, len, 0x05);
            result.system.version      = readFieldString(table, tableLength, offset, len, 0x06);
            result.system.serialNumber = readFieldString(table, tableLength, offset, len, 0x07);
            if(len >= 0x18) result.system.uuid = formatUuid(s + 0x08);
            break;
        case 2:  // Baseboard
            result.baseboard.manufacturer = readFieldString(table, tableLength, offset, len, 0x04);
            result.baseboard.product      = readFieldString(table, tableLength, offset, len, 0x05);
            result.baseboard.version      = readFieldString(table, tableLength, offset, len, 0x06);
            result.baseboard.serialNumber = readFieldString(table, tableLength, offset, len, 0x07);
            break;
        case 3:  // Chassis
            if(len > 0x05) result.chassis.type = s[0x05];
            result.chassis.manufacturer = readFieldString(table, tableLength, offset, len, 0x04);
            result.chassis.serialNumber = readFieldString(table, tableLength, offset, len, 0x07);
            result.chassis.assetTag     = readFieldString(table, tableLength, offset, len, 0x08);
            break;
        case 4:  // Processor
        {
            SmbiosProcessor processor;
            processor.socket       = readFieldString(table, tableLength, offset, len, 0x04);
            processor.manufacturer = readFieldString(table, tableLength, offset, len, 0x07);
            processor.version      = readFieldString(table, tableLength, offset, len, 0x10);
            if(len > 0x23) {
                processor.coreCount = s[0x23];
                if(processor.coreCount == 0xFF && len >= 0x2A) {
                    processor.coreCount = readU16(s + 0x28);  // CoreCount2
                }
            }
            if(len > 0x25) {
                processor.threadCount = s[0x25];
                if(processor.threadCount == 0xFF && len >= 0x2E) {
                    processor.threadCount = readU16(s + 0x2C);  // ThreadCount2
                }
            }
            result.processors.push_back(std::move(processor));
            break;
        }
        case 17:  // Memory Device
        {
            SmbiosMemoryDevice memory;
            if(len >= 0x0E) {
                const std::uint16_t raw = readU16(s + 0x0C);
                if(raw == 0xFFFF && len >= 0x20) {
                    memory.sizeMb = readU32(s + 0x1C) & 0x7FFFFFFF;  // ExtendedSize，单位 MB
                } else if((raw & 0x8000) != 0) {
                    memory.sizeMb = (raw & 0x7FFF) / 1024;  // 单位为 KB，换算为 MB
                } else {
                    memory.sizeMb = raw;  // 单位 MB
                }
            }
            if(len >= 0x17) memory.speed = readU16(s + 0x15);
            memory.manufacturer = readFieldString(table, tableLength, offset, len, 0x17);
            memory.serialNumber = readFieldString(table, tableLength, offset, len, 0x18);
            memory.partNumber   = readFieldString(table, tableLength, offset, len, 0x1A);

            if(memory.sizeMb > 0) result.memory.push_back(std::move(memory));  // 跳过空插槽
            break;
        }
        default:
            break;
        }

        // 越过字符串区：字符串以 '\0' 结尾，整段以双 '\0' 结束。
        std::size_t p = offset + len;
        while(p + 1 < tableLength && !(table[p] == 0 && table[p + 1] == 0)) ++p;
        offset = p + 2;
    }

    return result;
}

namespace {

// 归一化: 去首尾空白、转大写，并把 OEM 占位值/无效值统一为空串。
std::string normalize(std::string value) {
    const auto notSpace = [](char c) { return std::isspace(static_cast<unsigned char>(c)) == 0; };
    value.erase(value.begin(), std::find_if(value.begin(), value.end(), notSpace));
    value.erase(std::find_if(value.rbegin(), value.rend(), notSpace).base(), value.end());

    std::transform(value.begin(), value.end(), value.begin(), [](char c) {
        return static_cast<char>(std::toupper(static_cast<unsigned char>(c)));
    });

    static constexpr std::array<std::string_view, 14> kPlaceholders = {
        "TO BE FILLED BY O.E.M.", "TO BE FILLED BY OEM", "DEFAULT STRING",
        "SYSTEM SERIAL NUMBER",   "NOT SPECIFIED",        "NOT APPLICABLE",
        "INVALID",                "UNKNOWN",              "NONE",
        "NA",                     "N/A",                  "OEM",
        "0",                      ""
    };
    for(const auto &placeholder : kPlaceholders) {
        if(value == placeholder) return {};
    }
    return value;
}

// FNV-1a(Fowler–Noll–Vo) 64 位哈希算法原理:
//   1. 固定参数: 偏移基数(offset basis) = 14695981039346656037 = 0xCBF29CE484222325，
//      质数(prime) = 1099511628211 = 0x100000001B3。
//   2. 初值 h = offset basis。
//   3. 对输入流的每个字节 b 依次: h = (h ^ b) * prime。
//      先异或后相乘是 FNV-1a 与 FNV-1 的唯一区别(FNV-1 为先乘后异或)，
//      先异或可让输入低位的改变更快扩散。
//   4. 全部字节处理完后 h 即为结果。
// 说明: 该算法简单、快速、分布均匀，但属于“非加密哈希”，不抵抗碰撞构造与长度扩展，
//       因此仅用于指纹/校验，不能替代 HMAC/SHA 等密码学手段。
constexpr std::uint64_t kFnvOffsetBasis = 14695981039346656037ULL;
constexpr std::uint64_t kFnvPrime       = 1099511628211ULL;

// 只把非空字段并入哈希，并用分隔符避免字段拼接歧义
// (例如 {"AB","C"} 与 {"A","BC"} 若不加分隔符会得到相同哈希)。
void accumulate(std::uint64_t &hash, const std::string &raw) {
    const auto value = normalize(raw);
    if(value.empty()) return;

    for(const char c : value) {
        hash = (hash ^ static_cast<std::uint8_t>(c)) * kFnvPrime;
    }
    hash = (hash ^ static_cast<std::uint8_t>('|')) * kFnvPrime;
}

bool contains(const std::string &haystack, std::string_view needle) {
    return haystack.find(needle) != std::string::npos;
}

// 第一层: 用 CPUID 判定。
// 采用虚拟机监控器厂商串而非 hypervisorPresent 位: 后者在开启 VBS/Hyper-V 的
// 物理机上同样为真(见 cpuidHypervisorPresent 注释)，单独使用会误判。
// "Microsoft Hv" 也刻意不列入下表 —— 它既可能来自 Hyper-V 来宾，也可能来自
// 开启 VBS 的物理机，需交由 SMBIOS 层二次确认。
bool cpuIdLooksVirtual(const CpuId &cpu) {
    static constexpr std::array<std::string_view, 9> kHypervisors = {
        "VMWAREVMWARE", "VBOXVBOXVBOX", "KVMKVMKVM",    "PRL HYPERV",
        "XENVMMXENVMM", "ACRNACRNACRN", "TCGTCGTCGTCG", "BHYVE BHYVE",
        "QNXQVMBSQG"
    };
    const auto hypervisor = normalize(cpu.hypervisorVendor);
    for(const auto &h : kHypervisors) {
        if(contains(hypervisor, h)) return true;
    }
    return false;
}

// 第二层: 用 SMBIOS 字符串判定。字符串可被改写，故仅作补充。
bool smbiosLooksVirtual(const Smbios &smbios) {
    const auto vendor  = normalize(smbios.system.manufacturer);
    const auto product = normalize(smbios.system.productName);
    const auto bios    = normalize(smbios.bios.vendor);
    const auto uuid    = normalize(smbios.system.uuid);

    // 厂商名。注意不要拉黑 "MICROSOFT CORPORATION"，否则会误伤 Surface；
    // Hyper-V 靠机型名 "VIRTUAL MACHINE" 判定。
    static constexpr std::array<std::string_view, 11> kVendors = {
        "VMWARE", "INNOTEK", "ORACLE", "QEMU", "PARALLELS", "XEN",
        "AMAZON", "GOOGLE",  "BOCHS",  "RED HAT", "NUTANIX"
    };
    for(const auto &v : kVendors) {
        if(contains(vendor, v)) return true;
    }

    static constexpr std::array<std::string_view, 7> kProducts = {
        "VIRTUALBOX", "VIRTUAL MACHINE", "VIRTUAL PLATFORM", "VMWARE",
        "KVM", "Q35", "BOCHS"
    };
    for(const auto &p : kProducts) {
        if(contains(product, p)) return true;
    }

    static constexpr std::array<std::string_view, 5> kBiosVendors = {
        "SEABIOS", "EDK II", "OVMF", "BOCHS", "XEN"
    };
    for(const auto &b : kBiosVendors) {
        if(contains(bios, b)) return true;
    }

    // VMware/VirtualBox 常用 "VM"(0x56 0x4D) 作为 UUID 前缀。
    return !uuid.empty() && uuid.rfind("564D", 0) == 0;
}

}  // namespace

bool smbiosIsVirtualMachine(const Smbios &smbios) {
    // CPUID 优先：命中已知虚拟机监控器厂商串即可直接断定(不受 SMBIOS 字符串篡改影响)。
    if(cpuIdLooksVirtual(smbios.cpu)) return true;

    // 再以 SMBIOS 字符串交叉验证，兼容 Hyper-V 等厂商串无法区分的场景。
    return smbiosLooksVirtual(smbios);
}

std::uint64_t smbiosFingerprint(const Smbios &smbios) {
    std::uint64_t hash = kFnvOffsetBasis;

    accumulate(hash, smbios.system.manufacturer);
    accumulate(hash, smbios.system.productName);
    accumulate(hash, smbios.system.uuid);
    accumulate(hash, smbios.baseboard.serialNumber);
    accumulate(hash, smbios.bios.vendor);

    return hash;
}

std::string smbiosFingerprintString(const Smbios &smbios) {
    return std::format("{:016X}", smbiosFingerprint(smbios));
}

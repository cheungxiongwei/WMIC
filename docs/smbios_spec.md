# SMBIOS 模块规格说明

本文档描述 `src/smbios.h` / `src/smbios.cpp` 的数据来源、解析规则与算法约定。
该模块仅依赖 Win32 API（`GetSystemFirmwareTable`）与 CPUID 指令，
**不使用 COM/WMI**，与 `wmic.h` 无耦合。

规范依据：DMTF *System Management BIOS (SMBIOS) Reference Specification* (DSP0134)。

## 1. 数据来源

```c
UINT GetSystemFirmwareTable(DWORD FirmwareTableProviderSignature,
                            DWORD FirmwareTableID,
                            PVOID pFirmwareTableBuffer,
                            DWORD BufferSize);
```

- 提供者签名 `'RSMB'`（原始 SMBIOS，整数常量 `0x52534D42`），`FirmwareTableID` 固定为 `0`。
- 首次调用传入空缓冲以获取所需长度，第二次调用取内容。
- 物理载体为主板 SPI 闪存中的固件表，Windows 在启动时缓存快照，同一次开机内内容不变。

### 1.1 原始表头（`RawSMBIOSData`）

| 偏移 | 大小 | 字段 | 说明 |
|------|------|------|------|
| 0x00 | 1 | Used20CallingMethod | 是否使用了 SMBIOS 2.0 调用方法（DMI 遗留字段） |
| 0x01 | 1 | MajorVersion | SMBIOS 主版本 |
| 0x02 | 1 | MinorVersion | SMBIOS 次版本 |
| 0x03 | 1 | DmiRevision | DMI 修订号 |
| 0x04 | 4 | Length | 其后结构流的总字节数 |
| 0x08 | - | SMBIOS 结构流 | 由若干结构首尾相接组成 |

## 2. 结构流遍历

每个结构由「头部 + 格式化区 + 字符串区」构成：

| 偏移 | 大小 | 字段 | 说明 |
|------|------|------|------|
| 0x00 | 1 | Type | 结构类型，`127` 为结束标记 |
| 0x01 | 1 | Length | 格式化区（含头部）的字节数 |
| 0x02 | 2 | Handle | 结构句柄 |

- 字符串区紧跟在格式化区之后；每个字符串以 `\0` 结尾，整段以双 `\0`（`0x0000`）结束。
- 若某结构的字符串区为空，则双 `\0` 紧接格式化区。
- 位于格式化区内的字符串字段存放的是 **字符串索引**（1 起），而不是字符串内容本身：
  - `0` 表示「无字符串」；
  - `0xFF` 依规范可表示「索引已定义但字符串未定义」。
- 下一个结构的起点 = 越过格式化区后，继续前进直到遇到双 `\0`，再跳过这 2 字节。

解析防御：当 `Length < 4`、结构越界或字符串区提前耗尽时，停止/返回空串，避免越界读取。

## 3. 已解析结构及字段偏移

偏移均相对「结构起点」（即 `Type` 所在字节）。

### Type 0 — BIOS Information

| 偏移 | 字段 |
|------|------|
| 0x04 | Vendor（固件厂商） |
| 0x05 | BIOS Version（版本，刷固件会变） |
| 0x08 | BIOS Release Date（字符串，格式随厂商而异） |

### Type 1 — System Information

| 偏移 | 字段 |
|------|------|
| 0x04 | Manufacturer |
| 0x05 | Product Name |
| 0x06 | Version |
| 0x07 | Serial Number |
| 0x08 | UUID（16 字节二进制） |

UUID 编码为混合端序：前三段（`time_low` / `time_mid` / `time_hi_and_version`）按小端存储，
需逐字节倒序；后两段（`clock_seq`、`node`）按大端（即原序）。格式化结果形如：

```
4A951E19-E486-6216-A74C-047C164B59A8
```

若 16 字节全为 `0x00` 或全为 `0xFF`，认定厂商未烧录，统一归一化为空串。

### Type 2 — Baseboard (Module) Information

| 偏移 | 字段 |
|------|------|
| 0x04 | Manufacturer |
| 0x05 | Product |
| 0x06 | Version |
| 0x07 | Serial Number |

### Type 3 — System Enclosure (Chassis)

| 偏移 | 字段 |
|------|------|
| 0x04 | Manufacturer |
| 0x05 | Type（机箱类型枚举，1 字节） |
| 0x07 | Serial Number |
| 0x08 | Asset Tag |

### Type 4 — Processor Information

| 偏移 | 字段 |
|------|------|
| 0x04 | Socket Designation |
| 0x07 | Manufacturer |
| 0x10 | Processor Version |
| 0x23 | Core Count（1 字节，`0xFF` 表示需看 CoreCount2） |
| 0x25 | Thread Count（1 字节，`0xFF` 表示需看 ThreadCount2） |
| 0x28 | Core Count 2（2 字节，仅当 Core Count = `0xFF` 时有效） |
| 0x2C | Thread Count 2（2 字节，仅当 Thread Count = `0xFF` 时有效） |

条件字段仅在结构长度足够时才可读取（代码以 `Length` 判定）。

### Type 17 — Memory Device

| 偏移 | 字段 |
|------|------|
| 0x0C | Size（2 字节） |
| 0x15 | Speed（2 字节，MT/s） |
| 0x17 | Manufacturer |
| 0x18 | Serial Number |
| 0x1A | Part Number |
| 0x1C | Extended Size（4 字节，仅当 Size = `0xFFFF` 时有效） |

`Size` 的语义：

| 取值 | 含义 |
|------|------|
| `0x0000` | 插槽未安装内存 → 跳过该条 |
| `0xFFFF` | 使用 Extended Size 字段（单位 MB，最高位为保留位） |
| bit15 = 1 | 低 15 位单位为 KB |
| 其他 | 单位为 MB |

## 4. 归一化规则

参与比较与哈希前，字符串统一执行：

1. 去除首尾空白；
2. 转为大写；
3. 命中占位/无效值则置为空串。

占位值集合：`TO BE FILLED BY O.E.M.`、`TO BE FILLED BY OEM`、`DEFAULT STRING`、
`SYSTEM SERIAL NUMBER`、`NOT SPECIFIED`、`NOT APPLICABLE`、`INVALID`、`UNKNOWN`、
`NONE`、`NA`、`N/A`、`OEM`、`0`、空串。

## 5. 硬件指纹

### 5.1 参与因子

按下列顺序拼接入哈希（空值跳过）：

| 顺序 | 来源 | 变化时机 |
|------|------|----------|
| 1 | Type 1 Manufacturer | 换机型 |
| 2 | Type 1 Product Name | 换机型 |
| 3 | Type 1 UUID | 换主板 |
| 4 | Type 2 Serial Number | 换主板 |
| 5 | Type 0 Vendor | 换主板/刷固件 |

以下字段**刻意不参与**：BIOS Version、系统序列号（常为占位值）、
内存序列号/部件号（换内存即变）、任何地址与表顺序（每次开机随机化）。

### 5.2 FNV-1a 64 位算法

```
offset basis = 0xCBF29CE484222325  (14695981039346656037)
prime        = 0x00000100000001B3  (1099511628211)

h = offset basis
for each byte b in input:
    h = (h XOR b) * prime
return h
```

要点与约定：

- 「先异或、后相乘」是 FNV-1a 区别于 FNV-1（先乘后异或）的唯一之处，前者能让低位变化更快扩散。
- 字段之间插入分隔符 `|` 一并参与运算，避免拼接歧义（`{AB,C}` 与 `{A,BC}` 应得到不同结果）。
- FNV-1a 属**非加密哈希**，分布均匀、实现简单，但可被构造碰撞，不得用于签名或完整性保护；
  如需防伪，应在许可层用私钥对指纹签名。
- 输出以 16 位十六进制字符串呈现（`smbiosFingerprintString`）。

## 6. CPUID 采集

CPUID 是 CPU 直接执行的指令，结果由 CPU/虚拟机监控器实时返回，不经由固件表，
因此比 SMBIOS 字符串更难伪造。仅 x86/x64 提供该指令（`#include <intrin.h>`）。

| Leaf | 寄存器顺序 | 用途 |
|------|------------|------|
| `0x00` | EBX, EDX, ECX | 真实厂商串，如 `GenuineIntel` / `AuthenticAMD` |
| `0x01` | ECX bit31 | Hypervisor Present 位 |
| `0x40000000` | EBX, ECX, EDX | 虚拟机监控器厂商串，如 `VMwareVMware` |
| `0x80000002..0x80000004` | 各 4 寄存器 × 3 叶 | 处理器品牌串（48 字节） |

注意：

- `0x40000000` 位于约定的虚拟机区间，其有效性由 Hypervisor Present 位决定，
  **不能**用 leaf `0x00` 返回的最大标准叶编号来判定（该编号通常远小于 `0x40000000`）。
- 开启 VBS/Hyper-V 的**物理机**也会置位 Hypervisor Present 且返回 `Microsoft Hv`，
  故该位与 `Microsoft Hv` 都不能单独作为「本机是虚拟机」的判据。

## 7. 虚拟机判定

两层交叉验证，任一层命中即为虚拟机：

**第一层：CPUID 厂商串**（优先，难伪造）

黑名单：`VMWAREVMWARE`、`VBOXVBOXVBOX`、`KVMKVMKVM`、`PRL HYPERV`、
`XENVMMXENVMM`、`ACRNACRNACRN`、`TCGTCGTCGTCG`、`BHYVE BHYVE`、`QNXQVMBSQG`。
刻意不含 `Microsoft Hv`（见第 6 节）。

**第二层：SMBIOS 字符串**（补充，可被 DMI 工具篡改）

- 系统厂商含：`VMWARE`、`INNOTEK`、`ORACLE`、`QEMU`、`PARALLELS`、`XEN`、
  `AMAZON`、`GOOGLE`、`BOCHS`、`RED HAT`、`NUTANIX`
- 机型名含：`VIRTUALBOX`、`VIRTUAL MACHINE`、`VIRTUAL PLATFORM`、`VMWARE`、`KVM`、`Q35`、`BOCHS`
- BIOS 厂商含：`SEABIOS`、`EDK II`、`OVMF`、`BOCHS`、`XEN`
- Type 1 UUID 以 `564D`（ASCII `VM`）开头

刻意不拉黑 `MICROSOFT CORPORATION`，避免误伤 Surface；Hyper-V 来宾由机型名
`Virtual Machine` 兜底判定。

## 8. 局限

- SMBIOS 字符串可由 `dmidecode`/DMI 编辑工具改写，判定与指纹均非强保证。
- 虚拟机可伪装厂商串以规避第一层判定。
- 指纹基于「换硬件才变」的假设，更换主板必然导致指纹变化。
- 未采集 ACPI 表；其内容（如 MCFG 基址）每次开机变化，不适合作硬件指纹。

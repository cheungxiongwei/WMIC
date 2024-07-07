#include <print>
#include "WMIC.h"

int main() {
    system("COLOR 0A");
    try {
        WMIC wmic;
        wmic.OperatingSystem();  // 系统
        wmic.VideoController();  // 显卡
        wmic.DiskDrive();        // 硬盘
        wmic.BaseBoard();        // 主板
        wmic.BIOS();             // 主板 BIOS 芯片
        wmic.PhysicalMemory();   // 内存条
        wmic.Processor();        // CPU处理器
        wmic.NetworkAdapter();   // 网卡

        // 以系统序列号为例,生成用户ID
        // 当然也可以任意特征组合形式进行用户ID生成
        auto meta = wmic.OperatingSystem().serialNumber;
        std::println("uid:{}", uid(std::vector<uint8_t>(meta.begin(), meta.end())));

    } catch(const WMICException &e) {
        std::println("{}",e.what());
	}
	return 0;
}
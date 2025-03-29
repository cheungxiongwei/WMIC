#include <print>
#include "wmic.h"

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
        auto uid  = ::uid(std::vector<uint8_t>(meta.begin(), meta.end()));
        std::println("uid:{}", uid);

        // UEFI SDRAM 读写
        // 注意:此功能需要以管理员权限运行，否则会失败。
        uid_meta_t data = {};

        // 读取硬件指纹
        std::println("{:s}", wmicUefiRead(uid, &data, sizeof(uid_meta_t)));
        std::println("UEFI SDRAM UID:{}", data.uid);

        // 存储硬件指纹信息
        // 填充数据
        data.uid = uid;
        std::println("{:s}", wmicUefiWrite(uid, &data, sizeof(uid_meta_t)));

    } catch(const WMICException &e) {
        std::println("{}", e.what());
    }
    system("PAUSE");
    return 0;
}
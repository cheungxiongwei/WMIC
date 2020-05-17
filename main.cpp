#if 0
#include <QCoreApplication>
#include "wmic.h"
int main(int argc, char *argv[])
{
	QCoreApplication a(argc, argv);
	system("COLOR 0A");
	try
	{
		WMIC w;
		w.OperatingSystem();//系统
		w.VideoController();//显卡
		w.DiskDrive();//硬盘
		w.BaseBoard();//主板
		w.BIOS();//主板 BIOS 芯片
		w.PhysicalMemory();//内存条
		w.Processor();//CPU处理器
		w.NetworkAdapter();//网卡
	}
	catch (const WMIC_ERROR& e)
	{
		std::cout << e.what() << std::endl;
	}

	return a.exec();
	return 0;
}
#else
#include "wmic.h"
int main(int argc, char *argv[])
{
	system("COLOR 0A");
	try
	{
		WMIC w;
		w.OperatingSystem();//系统
		w.VideoController();//显卡
		w.DiskDrive();//硬盘
		w.BaseBoard();//主板
		w.BIOS();//主板 BIOS 芯片
		w.PhysicalMemory();//内存条
		w.Processor();//CPU处理器
		w.NetworkAdapter();//网卡
	}
	catch (const WMIC_ERROR& e)
	{
		std::cout << e.what() << std::endl;
	}

	return 0;
}
#endif

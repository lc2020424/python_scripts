import time

import pywifi
from pywifi import const


def wifi_connect_status():
    """
    判断本机是否有无线网卡,以及连接状态
    :return: 已连接或存在无线网卡返回1,否则返回0
    """
    # 创建一个元线对象
    wifi = pywifi.PyWiFi()

    # 取当前机器,第一个元线网卡
    iface = wifi.interfaces()[0]  # 有可能有多个无线网卡,所以要指定

    # 判断是否连接成功
    if iface.status() in [const.IFACE_CONNECTED, const.IFACE_INACTIVE]:
        print('wifi已连接')
        return 1
    else:
        print('wifi未连接')
    return 0


def scan_wifi():
    """
    扫苗附件wifi
    :return: 扫苗结果对象
    """
    # 扫苗附件wifi
    wifi = pywifi.PyWiFi()
    iface = wifi.interfaces()[0]

    iface.scan()  # 扫苗附件wifi
    time.sleep(1)
    basewifi = iface.scan_results()
    for i in basewifi:
        print('wifi扫苗结果:{}'.format(i.ssid))  # ssid 为wifi名称
        print('wifi设备MAC地址:{}'.format(i.bssid))
    return basewifi


def scan_wifi():
    """
    扫苗附件wifi
    :return: 扫苗结果对象
    """
    # 扫苗附件wifi
    wifi = pywifi.PyWiFi()
    iface = wifi.interfaces()[0]

    iface.scan()  # 扫苗附件wifi
    time.sleep(1)
    basewifi = iface.scan_results()
    for i in basewifi:
        print('wifi扫苗结果:{}'.format(i.ssid))  # ssid 为wifi名称
        print('wifi设备MAC地址:{}'.format(i.bssid))
    return basewifi


wifi_connect_status()
scan_wifi()

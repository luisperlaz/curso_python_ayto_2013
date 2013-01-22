# -.- coding: utf-8 -.-

import wmi

c = wmi.WMI()

print(findIp(u"Conexión de área local"))

def findIp(adapterName):
	for adapter in c.Win32_NetworkAdapter(NetConnectionId=adapterName):
		index = adapter.Index
		for config in c.Win32_NetworkAdapterConfiguration(Index=index):
			return config.IPAddress


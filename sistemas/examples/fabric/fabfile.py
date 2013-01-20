from fabric.api import *

env.hosts=['luis@192.168.0.20']

def cpuinfo():
	res = run("cat /proc/cpuinfo")
	print(res)

def meminfo():
	res = run("cat /proc/meminfo")
	print(res)

def isInstalled(package):
    dpkgcmd = "dpkg -s {0} |grep 'Status'".format(package)
    installed = "installed" in local(dpkgcmd, capture=True)
    print(installed) #do whatever if not

def localcpuinfo():
    res = local("cat /proc/cpuinfo", capture=True)
    print(res)

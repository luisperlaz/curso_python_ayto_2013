from fabric.api import *
from diskutils import *
#import diskutils

env.hosts=['127.0.0.1']

@task
def cpuinfo():
	res = run("cat /proc/cpuinfo")
	print(res)

@task
def meminfo():
	res = run("cat /proc/meminfo")
	print(res)

@task
def localcpuinfo():
    res = local("cat /proc/cpuinfo", capture=True)
    print(res)

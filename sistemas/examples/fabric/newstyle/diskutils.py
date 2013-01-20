from fabric.api import *

@task
def diskstats():
    res = run("cat /proc/diskstats")
    print(res)

@task
def localdiskstats():
    res = run("cat /proc/diskstats")
    print(res)

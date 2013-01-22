from fabric.api import *
from fabric.contrib.files import *

@hosts("luis@localhost")
def ntp_modify():

    with warn_only():
        res = sudo("cp /etc/ntp.conf /etc/ntp.conf.bak")
        if res.failed:
            put("ntp.conf", "/etc/ntp.conf", use_sudo=True)
    
        comment("/etc/ntp.conf", "^server 1.*", use_sudo=True)
        sed("/etc/ntp.conf", "server 2", "server 3", use_sudo=True)

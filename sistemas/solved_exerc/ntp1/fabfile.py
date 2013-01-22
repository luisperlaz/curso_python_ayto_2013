from fabric.api import *

@hosts("luis@localhost")
def ntp_user():
    sudo("cp /etc/ntp.conf /etc/ntp.conf.bak")
    put("ntp.conf", "/etc/ntp.conf", use_sudo=True)
    sudo("service ntp restart")

@hosts("root@localhost")
def ntp_root():
    run("cp /etc/ntp.conf /etc/ntp.conf.bak")
    put("ntp.conf", "/etc/ntp.conf" )
    run("service ntp restart")

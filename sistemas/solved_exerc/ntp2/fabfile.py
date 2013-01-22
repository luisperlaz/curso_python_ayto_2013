from fabric.api import *
from fabric.contrib.files import *

@hosts("luis@localhost")
def ntp_modify():
    comment("/etc/ntp.conf", "server 1.*", use_sudo=True)
    sed("/etc/ntp.conf", "server 2", "server 3", use_sudo=True)

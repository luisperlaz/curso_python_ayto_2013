from fabric.api import *
from cuisine import *
import os.path

env.hosts = ["127.0.0.1"]

select_package("zypper")

@task
def ensure_postgres():
    newinstalled = package_ensure(["postgresql", "postgresql-server"])
    if newinstalled:
        sed("/var/lib/pgsql/data/pg_hba.conf", "peer", "trust", use_sudo=True)
        sudo("/usr/sbin/rcpostgresql start")

@task
def install_bd(sqlfile=None):
    if not sqlfile or not os.path.exists(sqlfile):
        abort("please pass a valid sqlfile as parameter!")
    ensure_postgres()
    put(sqlfile, "/tmp/toimport.sql", use_sudo=True)
    sudo('psql -U postgres -c "create database bizi"')
    sudo('psql bizi -U postgres < /tmp/toimport.sql')
    


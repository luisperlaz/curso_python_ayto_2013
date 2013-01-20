from fabric.api import *
from cuisine import *

env.hosts = ["127.0.0.1"]

select_package("zypper")

@task
def editors_ensure():
       package_ensure(["vim", "emacs"])


@task
def config_vim():
    with hide('running', 'output'):
        file_update(
            "/home/luis/.vimrc",
            lambda _: text_ensure_line(_, 
            ":nnoremap <F9> :previous<CR>", ":nnoremap <F12> :next<CR>")
        )

from winsys import environment, registry
import wmi
import time

PACKAGE_LOC = "c:\python-2.7.msi"
PRODUCT_NAME = "Python 2.7"
EXEC_CMD = "python"


def install():
	print("installing")
	c = wmi.WMI()
	c.Win32_Product.Install(AllUsers=False, PackageLocation=PACKAGE_LOC)
	_set_python_in_path("2.7")


def uninstall():
	print("Uninstalling")
	c = wmi.WMI()
	_set_python_in_path("3.3")
	for prod in c.Win32_Product(Name=PRODUCT_NAME):
		prod.Uninstall()
	
def execute():
	print("Executing")
	c = wmi.WMI()
	pid, ret_val = c.Win32_Process.Create(CommandLine=EXEC_CMD)
	time.sleep(5)
	print("Killing process")
	for p in c.Win32_Process(ProcessId=pid):
		p.Terminate()
	print("Finished process")

def _set_python_in_path(version):
	paths = _get_python_paths(exclude=version)
	env = environment.system()
	_remove_from_path(env, paths)
	python27path = _get_python_paths(include=version)[0]
	_add_to_path(env, python27path)

def _extract_installpath(vers_reg_key):
	instkey = vers_reg_key + "InstallPath"
	if instkey:
		return instkey.get_value("")
			
def _get_python_paths(exclude=None, include=None):
	regbase = registry.registry(r"HKLM\SOFTWARE\Python\PythonCore")
	versions = [vers for vers in regbase]
	if exclude:
		versions = [vers for vers in versions if exclude not in vers.name]	
	if include:
		versions = [vers for vers in versions if include in vers.name]	

	install_paths = [_extract_installpath(vers) for vers in versions]
	install_paths = [path for path in install_paths if path is not None]
	return install_paths


def _remove_from_path(env, paths2remove):
	path = env["Path"] 
	path = ";".join(p for p in path.split(";") if not any(
			p.startswith(toremove) for toremove in paths2remove))
	env["Path"] = path

def _add_to_path(env, python27path):
	path = env["Path"] 
	path = "{};{}".format(path, python27path)
	env["Path"] = path

	
install()
execute()
uninstall()


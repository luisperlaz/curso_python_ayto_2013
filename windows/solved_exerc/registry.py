from winsys import registry

base = registry.registry(r"HKCU\software")
for key, subkey, values in base.walk():
	if not key.name.startswith("Python"):
		key.visited=1

## 
base = registry.registry(r"HKCU\software")
for key, subkey, values in base.walk():
	if dict(values).get("visited") == 1:
		tocopy = str(key).replace("Software", "softcopy")
		key.copy(tocopy)


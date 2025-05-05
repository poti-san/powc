from pprint import pp

from powcwmi import WBEMLocator

locator = WBEMLocator.create()
service = locator.connect_server(r"root\cimv2")

for instance in service.create_instanceenum("Win32_ComputerSystem", shallow=False):
    cls = service.get_object(instance.classname or "")
    pp(cls.methodnames_all)

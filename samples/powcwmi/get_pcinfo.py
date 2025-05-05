from pprint import pp

from powcwmi import WBEMLocator

locator = WBEMLocator.create()
service = locator.connect_server(r"root\cimv2")

for instance in service.create_instanceenum("Win32_ComputerSystem"):
    pp((instance.get("SystemFamily").value.to_str(), instance.get("SystemSKUNumber").value.to_str()))

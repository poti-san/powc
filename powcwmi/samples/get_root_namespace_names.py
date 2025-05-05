from powcwmi import WBEMLocator

locator = WBEMLocator.create()
server = locator.connect_server("root")

for cls in sorted(server.create_instanceenum("__Namespace"), key=lambda cls: str(cls.get("Name").value)):
    print(cls.get("Name").value)

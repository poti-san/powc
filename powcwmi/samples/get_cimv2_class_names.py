from operator import attrgetter

from powcwmi import WBEMLocator

locator = WBEMLocator.create()
server = locator.connect_server(r"root\cimv2")

for cls in sorted(server.create_classenum(None, shallow=False), key=attrgetter("classname")):
    print(cls.classname)

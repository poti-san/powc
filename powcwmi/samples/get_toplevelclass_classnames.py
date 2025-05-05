from operator import attrgetter
from pprint import pp

from powcwmi import WBEMLocator

locator = WBEMLocator.create()
server = locator.connect_server("root\\cimv2")

for cls in sorted(server.create_classenum_toplevel(), key=attrgetter("classname")):
    pp(cls.classname)

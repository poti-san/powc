from comtypes import GUID
from powc.comobj import Moniker

moniker = Moniker.create_classmoniker(GUID.from_progid("Shell.Application"))
submonikers = tuple(moniker.enum_forward)

pass

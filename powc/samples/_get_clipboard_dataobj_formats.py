from powc.incomplete.dataobj import DataObject

obj = DataObject.get_clipboard()

fmtetcs = tuple(obj.iter_formatetc_getter())
print(repr(fmtetcs[0].tymed))

print()

from powc.dataobj import DataObject

obj = DataObject.get_clipboard()

fmtetcs = tuple(obj.iter_formatetc_getter())

for fmtetc in fmtetcs:
    data = obj.get_data(fmtetc.format).bytes
    print((fmtetc.format, data[:10]))

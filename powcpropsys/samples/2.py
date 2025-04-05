# フォントのシステムプロパティの名前を取得する。

from pprint import pp

from powcpropsys.propsys import PropertySystem

propsys = PropertySystem.create()
descs = (desc for desc in propsys.propdescs_system if ".Fonts." in desc.canonicalname)

pp(tuple(desc.canonicalname for desc in descs))

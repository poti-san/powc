from powcpropsys.propsys import PropertySystem

ps = PropertySystem.create()

for desc in ps.propdescs_all:
    cname = desc.canonicalname
    if not cname.startswith("System.") or cname.count(".") > 1:
        continue
    print(f"{str(desc)} {desc.propkey}")

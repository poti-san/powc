from powcpropsys.propsys import PropertySystem

ps = PropertySystem.create()
all = ps.propdescs_all

for desc in all[0:10]:
    print(f"{str(desc)} {desc.propkey}")

from powcpropsys.propsys import PropertySystem

ps = PropertySystem.create()

for desc in ps.propdescs_all:
    print(f"{str(desc)} {desc.propkey}")

from powcd2d.dwrite import DWriteFactory, DWriteFactoryType

factory = DWriteFactory.create(DWriteFactoryType.SHARED)
sysfontcol = factory.systemfontcollection
fontfamilies = sysfontcol.fontfamilies
for fontfamily in fontfamilies:
    print(fontfamily.familynames.strings)

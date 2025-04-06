from powc.stream import ComStream

stream = ComStream.createonmem(bytes((0, 0, 0, 0, 6)))
if stream is None:
    raise ValueError
x = stream.read_bytes_all()

stream.pos = 0
stream.write_bytes(bytearray((0, 1, 2, 3, 4, 5, 6)))
stream.size = 3

x = stream.read_bytes_all()
print(x)

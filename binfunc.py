from binascii import unhexlify
from io import BufferedReader, BufferedWriter
from struct import pack, unpack


class BinaryParser(object):
  def __init__(self, file):
    if not isinstance(file, BufferedReader):
      raise ValueError("'file' must be BufferReader.")
    self.file = file

  def seek(self, passed=0):
    self.file.seek(passed, 0)

  def bytes_to_int(self, num_byte:int, signed:bool) -> int:
    return int.from_bytes(self.file.read(num_byte), byteorder="big", signed=signed)

  def read_byte(self):
    return self.bytes_to_int(num_byte=1, signed=False)

  def read_uint16(self):
    return self.bytes_to_int(num_byte=2, signed=False)

  def read_int16(self):
    return self.bytes_to_int(num_byte=2, signed=True)

  def read_uint32(self):
    return self.bytes_to_int(num_byte=4, signed=False)

  def read_float32(self, num_byte=4):
    return unpack(">f", unhexlify(self.file.read(num_byte).hex()))[0]

  def read_byte_s(self, length:int):
    return [self.read_byte() for i in range(length)]

  def read_float_s(self, length:int):
    return [self.read_float32() for i in range(length)]

  def read_uint16_s(self, length:int):
    return [self.read_uint16() for i in range(length)]

  def read_uint32_s(self, length:int):
    return [self.read_uint32() for i in range(length)]

  def read_string(self):
    word = ""
    for i in range(4):
      word += str(self.file.read(1), encoding="utf-8", errors="replace")
    return word

  def read_float_half(self):
    binary = self.file.read(2) + b"\x00\x00"
    return unpack(">f", unhexlify(binary.hex()))[0]


class BinaryWriter(object):
  def __init__(self, file):
    if not isinstance(file, BufferedWriter):
      raise ValueError("'file' must be BufferWriter.")
    self.file = file

  def seek(self, passed=0):
    self.file.seek(passed, 0)

  def getaddress(self):
    return self.file.tell()

  def write_int(self, num, write_bytes:int, signed:bool):
    if not isinstance(num, int):
      # dfで読み込んだ型は全てfloat型
      num = int(num)
    return self.file.write(num.to_bytes(write_bytes, byteorder='big', signed=signed))

  def write_byte(self, num):
    return self.write_int(num, 1, False)

  def write_uint16(self, num):
    return self.write_int(num, 2, False)
  
  def write_int16(self, num):
    return self.write_int(num, 2, True)
  
  def write_uint32(self, num):
    return self.write_int(num, 4, False)

  def write_float32(self, num:float):
    return self.file.write(pack('>f', num))
  
  def write_byte_s(self, num:list):
    for _num in num:
      self.write_byte(_num)

  def write_uint16_s(self, num:list):
    for _num in num:
      self.write_uint16(_num)

  def write_uint32_s(self, num:list):
    for _num in num:
      self.write_uint32(_num)

  def write_float_s(self, num:list):
    for _num in num:
      self.write_float32(_num)

  def write_string(self, string:str):
    self.file.write(string.encode('utf-8'))

  def write_float_half(self, num:float):
    self.file.write(bytes(bytearray(pack('>f', num))[:2]))

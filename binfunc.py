from binascii import unhexlify
from io import BufferedReader, BufferedWriter
from struct import pack, unpack

import numpy as np


def unpack_16bits(num, cut=0) -> np.ndarray:
  # np.unpackbits not supported uint16
  x = np.array([num], np.uint16)
  p = np.power(2, np.arange(16))[::-1]
  x = (x&p).astype(np.bool).astype(np.uint8)
  if cut != 0:
    x = x[cut:]
  return x


def unpack_8bits(num, cut=0) -> np.ndarray:
  # faster than np.packbits
  if num>255:
    raise OverflowError(unpack_8bits.__name__ + ' not supported int16.')
  return unpack_16bits(num, 8+cut)


def pack_16bits(arr) -> int:
  # np.packbits not supported uint16
  if isinstance(arr, list):
    x = np.array(arr, np.uint8)
  else:
    x = arr.copy()

  p = np.power(2, np.arange(x.shape[-1]))[::-1]
  return int(np.dot(x, p))


def pack_8bits(arr) -> int:
  # faster than np.unpackbits
  if len(arr)>8:
    raise OverflowError(pack_8bits.__name__ + ' not supported int16.')
  return pack_16bits(arr)


class BinaryParser(object):
  def __init__(self, file):
    if not isinstance(file, BufferedReader):
      raise ValueError("'file' must be BufferReader.")
    self.file = file

  def seek(self, passed=0):
    self.file.seek(passed, 0)

  def bytes_to_int(self, num_byte, signed) -> int:
    return int.from_bytes(self.file.read(num_byte), byteorder="big", signed=signed)

  @property
  def bytes_to_bitarray(self) -> int:
    return unpack_16bits(self.read_uint16)

  @property
  def read_byte(self) -> int:
    """
    0x00 - 0xFF
    """
    return self.bytes_to_int(num_byte=1, signed=False)

  @property
  def read_uint16(self) -> int:
    """
    0x0000 - 0xFFFF (unsigned)
    """
    return self.bytes_to_int(num_byte=2, signed=False)

  @property
  def read_int16(self) -> int:
    """
    -0x8000 - 0x7FFF (signed)
    """
    return self.bytes_to_int(num_byte=2, signed=True)

  @property
  def read_uint32(self) -> int:
    """
    0x00000000 - 0xFFFFFFFF (unsigned)
    """
    return self.bytes_to_int(num_byte=4, signed=False)

  @property
  def read_float32(self) -> float:
    """
    float(single precision)
    """
    return unpack(">f", unhexlify(self.file.read(4).hex()))[0]

  def read_byte_s(self, length) -> list:
    """
    byte(list)
    """
    return [self.read_byte for i in range(length)]

  def read_float_s(self, length) -> list:
    """
    float(list)
    """
    return [self.read_float32 for i in range(length)]

  def read_uint16_s(self, length) -> list:
    """
    uint16(list)
    """
    return [self.read_uint16 for i in range(length)]

  def read_uint32_s(self, length) -> list:
    """
    uint32(list)
    """
    return [self.read_uint32 for i in range(length)]

  @property
  def read_string(self) -> str:
    """
    String
    """
    word = ""
    for i in range(4):
      word += str(self.file.read(1), encoding="utf-8", errors="replace")
    return word

  @property
  def read_float_half(self) -> float:
    """
    float(single precision), but only reads 8-bit(MSB).
    This function is for speed factor (STGI).
    """
    binary = self.file.read(2) + b"\x00\x00"
    return unpack(">f", unhexlify(binary.hex()))[0]

  @property
  def read_objects(self) -> list:
    """
    Parse GOBJ 0x00 for XPF.  

    Return : [Type(definition object), disable or enable, object ID]  

    Source : http://wiki.tockdom.com/wiki/Extended_presence_flags\/Technical_Description#Definition_Objects
    """
    x = self.bytes_to_bitarray
    object_id = pack_16bits(x[6:])
    op = pack_16bits(x[:3])
    return [op, x[3], object_id]

  @property
  def read_reference(self) -> str:
    """
    Parse GOBJ 0x02 for XPF.

    Return : Reference_id (Hex)

    Source : http://wiki.tockdom.com/wiki/Extended_presence_flags/Technical_Description#Predifined_Conditions
    """
    return hex(self.read_uint16)

  @property
  def read_presence(self) -> list:
    """
    Parse GOBJ 0x3A for XPF.

    Return : [MODE, Parameters, Presence flag(LSB 3bits)]

    Source : http://wiki.tockdom.com/wiki/Extended_presence_flags/Technical_Description#Presence_flag_.28and_MODE.29
    """
    x = self.bytes_to_bitarray
    mode = pack_16bits(x[:4])
    param = pack_8bits(x[4:10])
    return np.hstack([mode, param, x[-3:]]).tolist()


class BinaryWriter(object):
  def __init__(self, file):
    if not isinstance(file, BufferedWriter):
      raise ValueError("'file' must be BufferWriter.")
    self.file = file

  def seek(self, passed=0):
    self.file.seek(passed, 0)

  @property
  def getaddress(self) -> int:
    return self.file.tell()

  def write_int(self, num, write_bytes:int, signed:bool):
    if not isinstance(num, int):
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

  def write_hex(self, value:str):
    x = value.encode("utf-8")
    self.write_uint16(int(x, 0))

  def write_objects(self, data:list, idx:int, LE_CODE:bool):
    if LE_CODE:
      definition = unpack_8bits(data[0], 5)
      show = unpack_8bits(int(bool(data[1])), 7)
    else:
      definition = [0]*3
      show = 0
    object_id = unpack_16bits(data[2], 6)
    c = np.hstack([definition, show, [0, 0], object_id]).astype(np.uint16)
    c = int(pack_16bits(c))
    self.write_uint16(c)

  def write_reference(self, data:list):
    if bool(data[0]): # LE_CODE (MODE>0)
      mode = unpack_8bits(data[0], 4)
      params = unpack_8bits(data[1], 2) if data[1]<=63 else [0]*6
    else:
      mode = [0]*4
      params = [0]*6
    player_mode = np.array(data[2:5]).astype(np.bool).astype(np.uint8)
    c = np.hstack([mode, params, [0]*3, player_mode]).tolist()
    c = pack_16bits(c)
    self.write_uint16(c)

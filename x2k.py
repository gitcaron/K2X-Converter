import argparse
import os
from collections import Counter

import pandas as pd

from binfunc import BinaryWriter


def excel_convert(path, output):
  sheets = ["KTPT", "ENPT", "ENPH", "ITPT", "ITPH", "CKPT", "CKPH", "GOBJ", "POTI", "AREA", "CAME", "JGPT", "CNPT", "MSPT", "STGI"]
  with open(output, "wb") as f:
    writer = BinaryWriter(f)
    writer.write_string('RKMD')

    #とりあえず0で書く
    writer.write_uint32(0)

    #固定値
    writer.write_uint16(15)
    writer.write_uint16(76)
    writer.write_uint32(2520)

    # とりあえず0で書く
    writer.write_uint32_s([0]*15)

    sections = []
    for sheet in sheets:
      sections.append(writer.getaddress()-76)
      df = pd.read_excel(path, sheet_name=sheet)
      df = df.drop("Unnamed: 0", axis=1)
      writer.write_string(sheet)

      if sheet != "POTI" and sheet != "CAME":
        writer.write_uint16_s([len(df.index), 0])

      if sheet == "KTPT":
        for i in df.index:
          data = df.iloc[i].tolist()
          writer.write_float_s(data[:6])
          writer.write_int16(data[6])
          writer.write_uint16(0)
      elif sheet == "ENPT":
        for i in df.index:
          data = df.iloc[i].tolist()
          writer.write_float_s(data[:4])
          writer.write_uint16(data[4])
          writer.write_byte_s(data[5:7])
      elif sheet == "ENPH":
        for i in df.index:
          data = df.iloc[i].tolist()
          writer.write_byte_s(data[0:16])
      elif sheet == "ITPT":
        for i in df.index:
          data = df.iloc[i].tolist()
          writer.write_float_s(data[:4])
          writer.write_uint16_s(data[4:6])
      elif sheet == "ITPH" or sheet == "CKPH":
        for i in df.index:
          data = df.iloc[i].tolist()
          writer.write_byte_s(data[0:14])
          writer.write_uint16(0)
      elif sheet == "CKPT":
        for i in df.index:
          data = df.iloc[i].tolist()
          writer.write_float_s(data[:4])
          writer.write_byte_s(data[4:8])
      elif sheet =="GOBJ":
        for i in df.index:
          data = df.iloc[i].tolist()
          writer.write_uint16_s(data[:2])
          writer.write_float_s(data[2:11])
          writer.write_uint16_s(data[11:21])
      elif sheet =="POTI":
        poti_id = set(df["ID"].tolist())
        poti_num_pos = list(
            dict(Counter(df["ID"].tolist())).values())

        writer.write_uint16_s([len(poti_id), len(df.index)])
        k = 0
        for i in range(len(poti_id)):
          writer.write_uint16(poti_num_pos[i])
          writer.write_byte_s(df.iloc[k, 1:3].tolist())
          for j in range(poti_num_pos[i]):
            data = df.iloc[k+j].tolist()
            writer.write_float_s(data[3:6])
            writer.write_uint16_s(data[6:8])
          k += poti_num_pos[i]
      elif sheet =="AREA":
        for i in df.index:
          data = df.iloc[i].tolist()
          writer.write_byte_s(data[:4])
          writer.write_float_s(data[4:13])
          writer.write_uint16_s(data[13:15])
          writer.write_byte_s(data[15:17])
          writer.write_uint16(0)
      elif sheet =="CAME":
        writer.write_uint16(len(df.index))
        writer.write_byte(df["First1"].tolist().index(1))
        writer.write_byte(df["First2"].tolist().index(1))
        for i in df.index:
          data = df.iloc[i].tolist()
          writer.write_byte_s([data[0], data[3], 0, data[4]])
          writer.write_uint16_s(data[5:8])
          writer.write_byte_s([0,0])
          writer.write_float_s(data[8:23])
      elif sheet =="JGPT":
        for i in df.index:
          data = df.iloc[i].tolist()
          writer.write_float_s(data[:6])
          writer.write_uint16(0)
          writer.write_int16(data[6])
      elif sheet =="CNPT":
        for i in df.index:
          data = df.iloc[i].tolist()
          writer.write_float_s(data[:6])
          writer.write_uint16(data[6])
          writer.write_int16(data[7])
      elif sheet =="MSPT":
        for i in df.index:
          data = df.iloc[i].tolist()
          writer.write_float_s(data[:6])
          writer.write_uint16_s([data[6], 0])
      elif sheet =="STGI":
        data = df.iloc[0].tolist()
        writer.write_byte_s(data[:4] + [0] + data[4:8] + [0])
        writer.write_float_half(data[8])

    address = writer.getaddress()
    writer.seek(4)
    writer.write_uint32(address)
    writer.seek(16)
    writer.write_uint32_s(sections)


if __name__ == "__main__":
  parser = argparse.ArgumentParser()
  parser.add_argument('--excel', required=True, help='Excel file')
  parser.add_argument('--kmp', required=True, help='Output KMP path')
  parser.add_argument('-o', '--allow_overwrite', action='store_true', dest='o', 
      help='If enabled, allows overwriting.')

  arg = parser.parse_args()

  if not arg.o and os.path.exists(arg.kmp):
    raise FileExistsError(arg.kmp + " arleady exists.")

  excel_convert(arg.excel, arg.kmp)

import argparse
import os

import pandas as pd

from binfunc import BinaryWriter


def get_idx(df_list, df_index):
  idx_1 = [idx for idx, x in enumerate(df_list) if not x] + [len(list(df_index))]
  idx_2 = [idx_1[i+1]-idx_1[i] for i in range(len(idx_1)-1)]

  return idx_1, idx_2


def pt_ph_writer(writer:BinaryWriter, df:pd.DataFrame, sections:list, sect_id:int, target:str):
  writer.write_string(target.split(" ")[0].replace("H", "T"))
  writer.write_uint16_s([len(df.index), 0])

  idx_1, idx_2 = get_idx(df[target].isnull(), df.index)

  ph_data = []
  ph_section = idx_1[:-1]

  if sect_id == 2:
    write_next_ids = [255, 1]

  for i in df.index:
    data = df.iloc[i].tolist()
    writer.write_float_s(data[:4])
    if sect_id == 0:
      ph_st, ph_end = 8, 22
      writer.write_uint16(data[4])
      writer.write_byte_s(data[5:7])
    else:
      ph_st, ph_end = 7, 19
      if sect_id == 2:
        writer.write_byte_s(data[4:6])
        writer.write_byte_s(write_next_ids)
        if i in ph_section and i+1 in ph_section:
          write_next_ids = [255, 255]
        elif i+2==len(df.index) or i+2 in ph_section:
          write_next_ids = [i, 255]
        elif i+1 in ph_section:
          write_next_ids = [255, i+2]
        else:
          write_next_ids = [i, i+2]
      else:
        writer.write_uint16_s(data[4:6])
    if i in ph_section:
      ph_data.append(data[ph_st:ph_end])

  # ENPH, ITPH, CKPH
  sections.append(writer.getaddress-76)
  writer.write_string(target.split(" ")[0])
  writer.write_uint16_s([len(idx_2), 0])
  for i in range(len(ph_data)):
    data = ph_data[i]
    writer.write_byte_s([idx_1[i], idx_2[i]])
    if sect_id == 0:
      writer.write_byte_s(data[0:15])
    else:
      writer.write_byte_s(data[0:13])
      writer.write_uint16(0)


def object_writer(writer:BinaryWriter, df:pd.DataFrame):
  writer.write_uint16_s([len(df.index), 0])
  for i in df.index:
    data = df.iloc[i].tolist()
    object_id = data[:3]
    reference = data[3]
    vec = data[4:13]
    route_settings = data[13:22]
    presence_flags = data[22:27]

    writer.write_objects(object_id, i, bool(presence_flags[0]))
    writer.write_hex(reference)
    writer.write_float_s(vec)
    writer.write_uint16_s(route_settings)
    writer.write_reference(presence_flags)


def poti_writer(writer:BinaryWriter, df:pd.DataFrame):
  idx_1, idx_2 = get_idx(df["ID"].isnull(), df.index)

  writer.write_uint16_s([len(idx_1)-1, len(df.index)])
  k = 0
  for i in range(len(idx_1)-1):
    writer.write_uint16(idx_2[i])
    writer.write_byte_s(df.iloc[k, 1:3].tolist())
    for j in range(idx_2[i]):
      data = df.iloc[k+j].tolist()
      writer.write_float_s(data[3:6])
      writer.write_uint16_s(data[6:8])
    k += idx_2[i]


def came_writer(writer:BinaryWriter, df:pd.DataFrame):
  writer.write_uint16(len(df.index))
  writer.write_byte(df["First1"].tolist().index(1))
  writer.write_byte(df["First2"].tolist().index(1))
  for i in df.index:
    data = df.iloc[i].tolist()
    writer.write_byte_s([data[0], data[3], 0, data[4]])
    writer.write_uint16_s(data[5:8])
    writer.write_byte_s([0,0])
    writer.write_float_s(data[8:23])


def other_writer(writer:BinaryWriter, df:pd.DataFrame, sheet:str):
  writer.write_uint16_s([len(df.index), 0])
  for i in df.index:
    data = df.iloc[i].tolist()
    if sheet == "KTPT":
      writer.write_float_s(data[:6])
      writer.write_int16(data[6])
      writer.write_uint16(0)
    elif sheet == "AREA":
      writer.write_byte_s(data[:4])
      writer.write_float_s(data[4:13])
      writer.write_uint16_s(data[13:15])
      writer.write_byte_s(data[15:17])
      writer.write_uint16(0)
    elif sheet == "JGPT":
      writer.write_float_s(data[:6])
      writer.write_uint16(0)
      writer.write_int16(data[6])
    elif sheet == "CNPT":
      writer.write_float_s(data[:6])
      writer.write_uint16(data[6])
      writer.write_int16(data[7])
    elif sheet == "MSPT":
      writer.write_float_s(data[:6])
      writer.write_uint16_s([data[6], 0])


def excel_convert(path, output):
  sheets = ["KTPT", "ENPT+ENPH", "ITPT+ITPH", "CKPT+CKPH", "GOBJ", "POTI", "AREA", "CAME", "JGPT", "CNPT", "MSPT", "STGI"]
  target = dict(zip(sheets[1:4], ["ENPH ID", "ITPH ID", "CKPH ID"]))
  with open(output, "wb") as f:
    writer = BinaryWriter(f)

    # Header
    writer.write_string('RKMD')
    writer.write_uint32(0)
    writer.write_uint16(15)
    writer.write_uint16(76)
    writer.write_uint32(2520)
    writer.write_uint32_s([0]*15)

    sections = []
    for sheet in sheets:
      sections.append(writer.getaddress - 76)
      df = pd.read_excel(path, sheet_name=sheet, engine="openpyxl")
      df = df.drop("Unnamed: 0", axis=1)

      if sheet in target.keys():
        pt_ph_writer(writer, df, sections, list(target.keys()).index(sheet), target[sheet])
      else:
        writer.write_string(sheet)
        if sheet == "POTI":
          poti_writer(writer, df)
        elif sheet == "GOBJ":
          object_writer(writer, df)
        elif sheet == "CAME":
          came_writer(writer, df)
        elif sheet == "STGI":
          writer.write_uint16_s([len(df.index), 0])
          data = df.iloc[0].tolist()
          writer.write_byte_s(data[:4] + [0] + data[4:8] + [0])
          writer.write_float_half(data[8])
        else:
          other_writer(writer, df, sheet)
  
    address = writer.getaddress
    writer.seek(4)
    writer.write_uint32(address)
    writer.seek(16)
    writer.write_uint32_s(sections)


if __name__ == "__main__":
  parser = argparse.ArgumentParser()
  parser.add_argument('--excel', required=True, help='Excel file')
  parser.add_argument('--kmp', required=True, help='Output KMP path')
  parser.add_argument('-o', '--overwrite', action='store_true', dest='o', 
      help='If enabled, allows overwriting.')

  arg = parser.parse_args()

  if not arg.o and os.path.exists(arg.kmp):
    raise FileExistsError(arg.kmp + " arleady exists.")

  excel_convert(arg.excel, arg.kmp)
  print(f"{arg.excel} -> {arg.kmp}")

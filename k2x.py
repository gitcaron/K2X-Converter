import argparse
import os

import numpy as np
import pandas as pd

from binfunc import BinaryParser


def pt_ph_parse(parser:BinaryParser, entry:int, pts:list, sect:str) -> list:
  if len(pts) == 0:
    raise ValueError(sect[:-1] + "T" + " not found.")
  out = []
  for i in range(entry):
    start = parser.read_byte
    length = parser.read_byte
    append_data = parser.read_byte_s(12)
    if sect == "ENPH":
      append_data += parser.read_byte_s(2)
    else:
      parser.read_uint16 # padding
    for j in range(start, start+length):
      if j == start:
        pts[j] += [i] + append_data
      else:
        pts[j] += [np.nan]*(len(append_data)+1)
      out.append(pts[j])
  return out


def kmp_dump(path, dest):
  # Excel columns 1
  columns = {
      "KTPT" : [
        "Pos x", "Pos y", "Pos z", "Rot x", "Rot y", "Rot z", "PlayerIdx"],
      "ENPT+ENPH" : [
        "Pos x", "Pos y", "Pos z", "Range", "Setting1", "Setting2", "Setting3",\
        "ENPH ID", "Last 1", "Last 2", "Last 3", "Last 4", "Last 5", "Last 6",\
        "Next 1", "Next 2", "Next 3", "Next 4", "Next 5", "Next 6", "Dispatch1", "Dispatch2"],
      "ITPT+ITPH" : [
        "Pos x", "Pos y", "Pos z", "Range", "Setting1", "Setting2",\
        "ITPH ID", "Last 1", "Last 2", "Last 3", "Last 4", "Last 5", "Last 6",\
        "Next 1", "Next 2", "Next 3", "Next 4", "Next 5", "Next 6"],
      "CKPT+CKPH" : [
        "Left x", "Left y", "Right x", "Right y", "Respawn", "Type",\
        "CKPH ID", "Last 1", "Last 2", "Last 3", "Last 4", "Last 5", "Last 6",\
        "Next 1", "Next 2", "Next 3", "Next 4", "Next 5", "Next 6"],
      "GOBJ" : [
        "Type(LE)", "Enable(LE)", "Object", "Reference (hex)",\
        "Pos x", "Pos y", "Pos z", "Rot x", "Rot y", "Rot z",\
        "Scale x", "Scale y", "Scale z", "Route",\
        "Setting1", "Setting2", "Setting3", "Setting4", "Setting5", "Setting6", "Setting7", "Setting8",\
        "MODE", "Parameters", "Multi(>2)", "Multi(<3)", "Single"],
      "POTI" : [
        "ID", "PointSetting 1", " PointSetting 2", "Pos x", "Pos y", "Pos z", "Setting1", "Setting2"],
      "AREA" : [
        "Shape", "Type", "Camera", "Priority", "Pos x", "Pos y", "Pos z", "Rot x", "Rot y", "Rot z",\
        "Scale x", "Scale y", "Scale z", "Setting 1", "Setting 2", "Route", "Enemy"],
      "CAME" : [
        "Type", "First1", "First2", "Next", "Route", "Camera velocity", "Zoom velocity", "View velocity",\
        "Pos x", "Pos y", "Pos z", "Rot x", "Rot y", "Rot z", "ZoomStart", "ZoomEnd",\
        "Start pos x", "Start pos y", "Start pos z", "End pos x", "End pos y", "End pos z", "Time"],
      "JGPT" : [
        "Pos x", "Pos y", "Pos z", "Rot x", "Rot y", "Rot z", "Range"],
      "CNPT" : [
        "Pos x", "Pos y", "Pos z", "Rot x", "Rot y", "Rot z", "Cannon ID", "Shoot"],
      "MSPT" : [
        "Pos x", "Pos y", "Pos z", "Rot x", "Rot y", "Rot z", "Entry"],
      "STGI" : [
        "Lap", "Pole", "Distance", "Flare", "Flare R", "Flare G", "Flare B", "Flare A", "Speed Factor"]
      }
  # Excel columns 2 (modify width)
  columns_expand = [
      None, ["V", "W"], None, None, ["C", "E", "Y"],
      ["C","D"], None, ["G", "H", "I", "P", "Q", "R", "S", "T", "U", "V", "W"], 
      None, "H", None, "J"]
  function_paths = ["ENPH", "ITPH", "CKPH"]
  match_sect = list(columns.keys())[1:4]
  pd_dfs = []
  pts = []
  section_data = []

  with open(path, "rb") as f:
    parser = BinaryParser(f)
    magic = parser.read_string
    if magic != "RKMD":
      raise ValueError("Invalid file.")

    file_length = parser.read_uint32
    sections = parser.read_uint16
    header_length = parser.read_uint16
    version = parser.read_uint32
    offsets = parser.read_uint32_s(sections)

    for head in offsets:
      parser.seek(head + header_length)
      section_name = parser.read_string
      entry = parser.read_uint16
      if section_name == "CAME":
        came_st1 = parser.read_byte
        came_st2 = parser.read_byte
      else:
        parser.read_uint16 # padding

      data = []
      if section_name in function_paths:
        data = pt_ph_parse(parser, entry, pts, section_name)
        section_name = match_sect[function_paths.index(section_name)]
        pts.clear()
      elif section_name == "STGI":
        opt = parser.read_byte_s(4)
        parser.read_byte # padding
        flarecolor = parser.read_byte_s(4)
        padding = parser.read_byte
        speed = parser.read_float_half
        data.append(opt + flarecolor + [speed])
      else:
        for i in range(entry):
          if section_name == "KTPT":
            pos_rots = parser.read_float_s(6)
            index = parser.read_int16
            parser.read_uint16 # padding
            data.append(pos_rots + [index])
          elif section_name == "ENPT" or section_name == "ITPT":
            pos = parser.read_float_s(4)
            set1 = parser.read_uint16
            appends = pos + [set1]
            if section_name == "ENPT":
              appends += parser.read_byte_s(2)
            else:
              appends += [parser.read_uint16]
            pts.append(appends)
          elif section_name == "CKPT":
            pos = parser.read_float_s(4)
            opt = parser.read_byte_s(2)
            parser.read_byte_s(2) # not padding (prev, next), but ignore
            pts.append(pos + opt)
          elif section_name == "GOBJ":
            xpf_obj = parser.read_objects
            xpf_ref = parser.read_reference
            vec = parser.read_float_s(9)
            route = parser.read_uint16
            setting = parser.read_uint16_s(8)
            xpf_flag = parser.read_presence
            data.append(xpf_obj + [xpf_ref] + vec + [route] + setting + xpf_flag)
          elif section_name == "POTI":
            idx = i
            num_pos = parser.read_uint16
            _sets = parser.read_byte_s(2)
            for j in range(num_pos):
              pos = parser.read_float_s(3)
              _sets2 = parser.read_uint16_s(2)
              if j != 0:
                idx = np.nan
                _sets = [np.nan]*2
              data.append([idx] + _sets + pos + _sets2)
          elif section_name == "AREA":
            opt_1 = parser.read_byte_s(4)
            vec = parser.read_float_s(9)
            sets = parser.read_uint16_s(2)
            opt_2 = parser.read_byte_s(2)
            parser.read_uint16 # padding
            data.append(opt_1 + vec + sets + opt_2)
          elif section_name == "CAME":
            came_type = parser.read_byte
            next_idx = parser.read_byte
            parser.read_byte # padding
            route = parser.read_byte
            opt_1 = parser.read_uint16_s(3)
            parser.read_byte_s(2) # padding
            vec = parser.read_float_s(15)
            f1 = 1 if i==came_st1 else np.nan
            f2 = 1 if i==came_st2 else np.nan
            data.append([came_type, f1, f2, next_idx, route] + opt_1 + vec)
          elif section_name == "JGPT":
            pos_rots = parser.read_float_s(6)
            parser.read_uint16 # padding
            ranges = parser.read_int16
            data.append(pos_rots + [ranges])
          elif section_name == "CNPT":
            pos_rots = parser.read_float_s(6)
            opt = parser.read_uint16_s(2)
            data.append(pos_rots + opt)
          elif section_name == "MSPT":
            pos_rots = parser.read_float_s(6)
            idx = parser.read_uint16
            parser.read_uint16 # padding
            data.append(pos_rots + [idx])
          else:
            raise ValueError("Invalid header name.")
      if len(data) == 0 and len(pts) > 0: continue
      df = pd.DataFrame(data, columns=columns[section_name])
      pd_dfs.append(df)
      section_data.append(section_name)
    
  with pd.ExcelWriter(output) as writer:
    for i in range(len(pd_dfs)):
      pd_dfs[i].to_excel(writer, sheet_name=section_data[i], engine="openpyxl")
      if columns_expand[i] is not None:
        worksheet = writer.book[section_data[i]]
        if isinstance(columns_expand[i], str):
          worksheet.column_dimensions[columns_expand[i]].width = 15
        elif isinstance(columns_expand[i], list):
          for clms in columns_expand[i]:
            worksheet.column_dimensions[clms].width = 15


if __name__ == "__main__":
  parser = argparse.ArgumentParser()
  parser.add_argument('--kmp', required=True, help='KMP path')
  parser.add_argument('--excel', default=None, help='Dumped xlsx path')
  parser.add_argument('-o', '--overwrite', action='store_true', dest='o', 
      help='If enabled, allows overwriting.')

  arg = parser.parse_args()

  if arg.excel is None:
    output = arg.kmp + '.xlsx'
  else:
    output = arg.excel

  if not arg.o and os.path.exists(output):
    raise FileExistsError(output + " arleady exists.")

  kmp_dump(arg.kmp, output)
  print(f"{arg.kmp} -> {output}")

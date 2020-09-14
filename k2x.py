import argparse
import os

import pandas as pd

from binfunc import BinaryParser


def kmp_dump(path, dest):
  columns = [
    ["Pos x", "Pos y", "Pos z", "Rot x", "Rot y", "Rot z", "PlayerIndex"],
    ["Pos x", "Pos y", "Pos z", "Range", "Setting1", "Setting2", "Setting3"],
    ["Start ID", "Length", "Last 1", "Last 2", "Last 3", "Last 4", "Last 5", "Last 6",\
      "Next 1", "Next 2", "Next 3", "Next 4", "Next 5", "Next 6", "Dispatch 1", "Dispatch 2"],
    ["Pos x", "Pos y", "Pos z", "Range", "Setting1", "Setting2"],
    ["Start ID", "Length", "Last 1", "Last 2", "Last 3", "Last 4", "Last 5", "Last 6",\
      "Next 1", "Next 2", "Next 3", "Next 4", "Next 5", "Next 6"],
    ["Left pos x", "Left pos y", "Right pos x", "Right pos y", "Respawn", "Type", "Prev", "Next"],
    ["Start ID", "Length", "Last 1", "Last 2", "Last 3", "Last 4", "Last 5", "Last 6",\
      "Next 1", "Next 2", "Next 3", "Next 4", "Next 5", "Next 6"],
    ["Object", "Reference", "Pos x", "Pos y", "Pos z", "Rot x", "Rot y", "Rot z",\
      "Scale x", "Scale y", "Scale z", "Route", "Setting 1", "Setting 2", "Setting 3",\
        "Setting 4", "Setting 5", "Setting 6", "Setting 7", "Setting 8", "Presence"],
    ["ID", "Point Setting 1", " Point Setting 2", "Pos x", "Pos y", "Pos z", "Setting 3", "Setting 4"],
    ["Shape", "Type", "Camera", "Priority", "Pos x", "Pos y", "Pos z", "Rot x", "Rot y", "Rot z",\
      "Scale x", "Scale y", "Scale z", "Setting 1", "Setting 2", "Route", "Enemy"],
    ["Type", "First1", "First2", "Next", "Route", "Camera velocity", "Zoom velocity", "View velocity",\
      "Pos x", "Pos y", "Pos z", "Rot x", "Rot y", "Rot z", "ZoomStart", "ZoomEnd",\
        "Start pos x", "Start pos y", "Start pos z", "End pos x", "End pos y", "End pos z", "Time"],
    ["Pos x", "Pos y", "Pos z", "Rot x", "Rot y", "Rot z", "Range"],
    ["Pos x", "Pos y", "Pos z", "Rot x", "Rot y", "Rot z", "Cannon ID", "Shoot"],
    ["Pos x", "Pos y", "Pos z", "Rot x", "Rot y", "Rot z", "Entry"],
    ["Lap", "Pole", "Distance", "Flare", "Flare R", "Flare G", "Flare B", "Flare A", "Speed factor"]
    ]

  with open(path, "rb") as f:
    parser = BinaryParser(f)
    magic = parser.read_string()
    if magic != "RKMD":
      raise ValueError("Invalid file.")

    file_length = parser.read_uint32()
    sections = parser.read_uint16()
    header_length = parser.read_uint16()
    version = parser.read_uint32()
    offsets = parser.read_uint32_s(sections)

    pd_dfs = []
    section_data = []
    for x, head in enumerate(offsets):
      parser.seek(head + header_length)
      section_name = parser.read_string()
      entry = parser.read_uint16()
      if section_name == "CAME":
        came_st1 = parser.read_byte()
        came_st2 = parser.read_byte()
      else:
        poti_routes = parser.read_uint16()

      data = []
      if section_name == "KTPT":
        for i in range(entry):
          pos = parser.read_float_s(3)
          rots = parser.read_float_s(3)
          index = parser.read_int16()
          padding = parser.read_uint16()
          data.append(pos + rots + [index])
      elif section_name == "ENPT":
        for i in range(entry):
          pos = parser.read_float_s(3)
          pos_range = parser.read_float32()
          set1 = parser.read_uint16()
          set2 = parser.read_byte()
          set3 = parser.read_byte()
          data.append(pos + [pos_range, set1, set2, set3])
      elif section_name == "ENPH":
        for i in range(entry):
          start = parser.read_byte()
          length = parser.read_byte()
          lasts = parser.read_byte_s(6)
          nexts = parser.read_byte_s(6)
          arena_dispatsh = parser.read_byte()
          arena_dispatsh2 = parser.read_byte()
          data.append([start, length] + lasts + nexts + [arena_dispatsh, arena_dispatsh2])
      elif section_name == "ITPT":
        for i in range(entry):
          pos = parser.read_float_s(3)
          pos_range = parser.read_float32()
          set1 = parser.read_uint16()
          set2 = parser.read_uint16()
          data.append(pos + [pos_range, set1, set2])
      elif section_name == "ITPH":
        for i in range(entry):
          start = parser.read_byte()
          length = parser.read_byte()
          lasts = parser.read_byte_s(6)
          nexts = parser.read_byte_s(6)
          padding= parser.read_uint16()
          data.append([start, length] + lasts + nexts)
      elif section_name == "CKPT":
        for i in range(entry):
          posL = parser.read_float_s(2)
          posR = parser.read_float_s(2)
          respawn = parser.read_byte()
          checkpoint_type = parser.read_byte()
          previous = parser.read_byte()
          nextp = parser.read_byte()
          data.append(posL + posR + [respawn, checkpoint_type, previous, nextp])
      elif section_name == "CKPH":
        for i in range(entry):
          start = parser.read_byte()
          length = parser.read_byte()
          lasts = parser.read_byte_s(6)
          nexts = parser.read_byte_s(6)
          padding= parser.read_uint16()
          data.append([start, length] + lasts + nexts)
      elif section_name == "GOBJ":
        for i in range(entry):
          obj_id = parser.read_uint16()
          wiimms_ref = parser.read_uint16()
          pos = parser.read_float_s(3)
          rots = parser.read_float_s(3)
          scale = parser.read_float_s(3)
          route = parser.read_uint16()
          setting = parser.read_uint16_s(8)
          wiimms_flag = parser.read_uint16()
          data.append([obj_id, wiimms_ref] + pos + rots + scale + [route] + setting + [wiimms_flag])
      elif section_name == "POTI":
        for i in range(entry):
          num_pos = parser.read_uint16()
          set1 = parser.read_byte()
          set2 = parser.read_byte()
          for j in range(num_pos):
            pos = parser.read_float_s(3)
            _set1 = parser.read_uint16()
            _set2 = parser.read_uint16()
            data.append([i, set1, set2] + pos + [_set1, _set2])
      elif section_name == "AREA":
        for i in range(entry):
          shape = parser.read_byte()
          area_type = parser.read_byte()
          came_index = parser.read_byte()
          priority = parser.read_byte()
          pos = parser.read_float_s(3)
          rots = parser.read_float_s(3)
          scale = parser.read_float_s(3)
          set1 = parser.read_uint16()
          set2 = parser.read_uint16()
          routeid = parser.read_byte()
          enemy = parser.read_byte()
          padding = parser.read_uint16()
          data.append([shape, area_type, came_index, priority] + pos + rots + scale + [set1, set2, routeid, enemy])
      elif section_name == "CAME":
        for i in range(entry):
          came_type = parser.read_byte()
          next_idx = parser.read_byte()
          shake = parser.read_byte()
          route = parser.read_byte()
          came_velocity = parser.read_uint16()
          zoom = parser.read_uint16()
          view_velocity = parser.read_uint16()
          start = parser.read_byte()
          move = parser.read_byte()
          pos = parser.read_float_s(3)
          rots = parser.read_float_s(3)
          zoom_start = parser.read_float32()
          zoom_end = parser.read_float32()
          start_pos = parser.read_float_s(3)
          dest_pos = parser.read_float_s(3)
          time = parser.read_float32()
          f1 = 1 if i==came_st1 else 0
          f2 = 1 if i==came_st2 else 0
          data.append([came_type, f1, f2, next_idx, route, came_velocity, zoom, view_velocity] + pos + rots + [zoom_start, zoom_end] +\
            start_pos + dest_pos + [time])
      elif section_name == "JGPT":
        for i in range(entry):
          pos = parser.read_float_s(3)
          rots = parser.read_float_s(3)
          idx = parser.read_uint16()
          ranges = parser.read_int16()
          data.append(pos + rots + [ranges])
      elif section_name == "CNPT":
        for i in range(entry):
          pos = parser.read_float_s(3)
          rots = parser.read_float_s(3)
          idx = parser.read_uint16()
          effect = parser.read_int16()
          data.append(pos + rots + [idx, effect])
      elif section_name == "MSPT":
        for i in range(entry):
          pos = parser.read_float_s(3)
          rots = parser.read_float_s(3)
          idx = parser.read_uint16()
          padding = parser.read_uint16()
          data.append(pos + rots + [idx])
      elif section_name == "STGI":
        lap = parser.read_byte()
        polepos = parser.read_byte()
        dist = parser.read_byte()
        isflare = parser.read_byte()
        padding = parser.read_byte()
        flarecolor = parser.read_byte_s(4)
        padding = parser.read_byte()
        speed = parser.read_float_half()
        data.append([lap, polepos, dist, isflare] + flarecolor + [speed])
      else:
        raise ValueError("Invalid header name.")
      df = pd.DataFrame(data, columns=columns[x])
      pd_dfs.append(df)
      section_data.append(section_name)
    
  with pd.ExcelWriter(output) as writer:
    for i in range(len(offsets)):
      pd_dfs[i].to_excel(writer, sheet_name=section_data[i])


if __name__ == "__main__":
  parser = argparse.ArgumentParser()
  parser.add_argument('--kmp', required=True, help='KMP path')
  parser.add_argument('--excel', default=None, help='Dumped xlsx path')
  parser.add_argument('-o', '--allow_overwrite', action='store_true', dest='o', 
      help='If enabled, allows overwriting.')

  arg = parser.parse_args()

  if arg.excel is None:
    output = arg.kmp + '.xlsx'
  else:
    output = arg.excel

  if not arg.o and os.path.exists(output):
    raise FileExistsError(output + " arleady exists.")

  kmp_dump(arg.kmp, output)

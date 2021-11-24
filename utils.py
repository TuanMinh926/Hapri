import os
import pandas as pd
import numpy as np

'''
Hàm folder_input:
-Lấy theo path_folder 
-Tạo list chứa file đuôi'.xlxs'
'''
def folder_input (path_folder):
  item_file = sorted(os.listdir(path_folder))
  lst_test = [item for item in item_file if item.find('.xlsx')!= -1]
  return lst_test

def cre_shape (df, path_excel, name):
  checks = df.iloc[0].notna()
  column_split = []
  shape_saving = []
  bool_split = 0
  for key in checks.keys():
    if (checks.get(key) == True):
      if (bool_split == 0):
        bool_split = 1
        column_split.append(key)
    else:
      bool_split = 0
  column_split.sort(reverse=True)
  if len(column_split) != 0:
    for i in range(len(column_split)):
      if i == len(column_split)-1:
        shape_saving.append((df.shape[0],[df.columns[0]+1,df.columns[-1]+1]))
      else:
        df_at_end = pd.read_excel(path_excel,sheet_name=name,usecols=df.columns[column_split[i]:],header=None)
        df.drop(df_at_end,axis=1,inplace=True)
        shape_saving.append((df_at_end.shape[0],[df_at_end.columns[0]+1,df_at_end.columns[-1]+1]))
  return shape_saving

'''
Hàm Change_method:
- Đầu và số đường kẻ
- Nếu không có đường kẻ dùng phương pháp khác
'''
def Change_method(metrics):
    if metrics != 0:
      return 0
    else:
      return 1

'''
Hàm NumtoAlpha:
-Chuyển số thành tên cột excel
-VD: 1->A
'''
def NumtoAlpha(n):
    string = ""
    while n > 0:
        n, remainder = divmod(n - 1, 26)
        string = chr(65 + remainder) + string
    return string

'''
Hàm getMergedCellRange:
- Lấy khoảng merged có chứa tọa độ excel
- Lấy cột thấp nhất - cột cao nhất + 1 trong khoảng merged
- Mục tiêu tính chiều dài cell merged
'''
def getMergedCellRange(sheet, cell):
    rng = [s for s in sheet.merged_cells.ranges if cell.coordinate in s]
    if len(rng)!=0:
      return rng[0].max_col-rng[0].min_col+1
    else:
      return 1

def column_end_outside(df):
  total_columns = df.shape[1]
  index_columns_end = []
  df_col = np.array(df.columns)
  for index in range(total_columns):
    if len(index_columns_end) < round(total_columns/2):
      index_columns_end.append(total_columns-(index+1))
  array_col_end = df_col[[index_columns_end]]
  column_at_end = df[array_col_end].notna()
  column_at_start = df.drop(column_at_end,axis=1).notna()
  return column_at_end, column_at_start

'''
Hàm Reverse:
- Đảo ngược list để xét dưới lên trên
'''
def Reverse(tuples):
    new_tup = tuples[::-1]
    return new_tup 
from utils import column_end_outside



'''
Hàm Check_Unit:
- Kiểm tra chuỗi có 1 trong các dấu hiệu 'Unit','Units','ĐVT','Đơn vị tính' hay không
- có trả về True
'''
def Check_Unit (unit):
  if type(unit) == str:
    condition_check = ['Unit','Units','ĐVT','Đơn vị tính']
    total_check = sum([unit.find(j) for j in (condition_check)])
    if total_check != -4:
      return True
    else:
      return False
  else:
    return False

def split_element (df):
  lst_title = []
  lst_unit = []
  lst_element = []
  at_step = 0
  total_rows = df.shape[0]
  ############################## Chia 1/2 cột đầu, cuối
  column_at_end, column_at_start = column_end_outside(df)
  ################################
  row = 0
  while row < total_rows:
    condition1 = sum(column_at_start.loc[row])
    condition2 = sum(column_at_end.loc[row])
    unit_row = []
    if at_step == 0:
      ##################### Xử lý title
      if (condition1!=0) & (condition2==0):
        ##################### Kiểm tra lẫn Đơn vị trong Title
        if sum(df[column_at_start.columns].loc[row].apply(Check_Unit)) != 0:
          for column in column_at_start.columns:
            if type(df[column][row]) == str:
              unit_row.append(df[column][row])
              unit_row = list(set(unit_row))
          for unit in unit_row:
            lst_unit.append(unit)
          if len(lst_unit) > 2:
            row -= 1
            at_step = 0
            lst_unit.clear()
            ### reset elemental
            # if count_table <= n_table:
            lst_element.extend([[lst_title,lst_unit]])
            lst_title = []
            lst_unit = []
            # count_table +=1
        else:
          #####################
          for column in column_at_start.columns:
            if (column_at_start[column][row] == True)&(type(df[column][row])==str):
              if df[column][row].find('Tiếp theo') == -1:
                process_title= ''.join(df[column][row].split(' (tiếp theo)'))
                lst_title.append(process_title)
              else:
                ### reset elemental
                # if count_table <= n_table:
                lst_element.extend([[lst_title,lst_unit]])
                # count_table +=1
                lst_title = []
                lst_unit = []
                process_title= ''.join(df[column][row].split(' (Tiếp theo)'))
                lst_title.append(process_title)
          #####################
      elif (condition1==0) & (condition2==0):
        df.drop(row,inplace=True)
      else:
        at_step = 1
        ####################
    if at_step == 1:
      ########################## Xử lý Đơn Vị
      if (condition1==0) & (condition2!=0):
        for column in column_at_end.columns:
          # if column_at_end[column][row] == True:
          if Check_Unit(df[column][row]) == True:
            unit_row.append(df[column][row])
            unit_row = list(set(unit_row))
        for unit in unit_row:
          lst_unit.append(unit)
        if len(lst_unit) > 2:
          row -= 1
          at_step = 0
          lst_unit.clear()
          ### reset elemental
          # if count_table <= n_table:
          lst_element.extend([[lst_title,lst_unit]])
          lst_title = []
          lst_unit = []
          # count_table +=1
      elif (condition1==0) & (condition2==0):
        df.drop(row,inplace=True)
      elif (condition1!=0) & (condition2!=0):
        for column in column_at_end.columns:
          # if column_at_end[column][row] == True:
          if Check_Unit(df[column][row]) == True:
            unit_row.append(df[column][row])
            unit_row = list(set(unit_row))
          else:
            at_step = 2
        for unit in unit_row:
          lst_unit.append(unit)
        if len(lst_unit) > 2:
          row -= 1
          at_step = 0
          lst_unit.clear()
          ### reset elemental
          # if count_table <= n_table:
          lst_element.extend([[lst_title,lst_unit]])
          lst_title = []
          lst_unit = []
          # count_table +=1
      else:
        ### reset elemental
        # if count_table <= n_table:
        lst_element.extend([[lst_title,lst_unit]])
        # count_table +=1
        lst_title = []
        lst_unit = []
        at_step = 0
        row -= 1
    if row == (total_rows-1):
      # if count_table <= n_table:
      lst_element.extend([[lst_title,lst_unit]])
    row += 1
  return lst_element,df

def print_element (lst_element,df,name):
  check_table = 0
  for element in lst_element:
    check_table+=sum([len(i) for i in element])
    print(element)
    if len(element[0])==0:
      print('\33[91m Please note this data sheet '+name+'\33[0;0m')
  if (check_table+len(lst_element)) < sum(df.notna().sum()):
    print('\33[91m Please note this data sheet '+name+'\33[0;0m')
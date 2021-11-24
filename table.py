import pandas as pd
from itertools import zip_longest


from utils import NumtoAlpha, column_end_outside, getMergedCellRange, Reverse
from title_unit import Check_Unit

'''
Hàm bor_row : 
- Thêm số ô [dòng,cột] nếu ô có giá trị đường kẻ
- Thêm số [dòng-1,-1] nếu ô có chữ 'Tiếp theo'
- Thêm số [dòng,-1] nếu đếm được 2 dòng full nan
- Trả về list
'''
def bor_row (sheet,shape_data):
  lst_border = []
  amount_row_na = 0
  for row in range(1, shape_data[0]+ 1):
      count_na = 0
      for col in range(shape_data[1][0],shape_data[1][1]+1):
        count_check = 0
        sheet_col = NumtoAlpha(col)
        coordinate_cell = str(sheet_col)+str(row)
        condition1 = (sheet[coordinate_cell].border.bottom.style or sheet[coordinate_cell].border.top.style)
        check_border = ((condition1) == 'thin') | ((condition1) == 'medium')
        count_check += check_border
        value_sheet = sheet[coordinate_cell].value
        if value_sheet == None:
          count_na +=1
          if count_na == sheet.max_column:
            amount_row_na += 1
        else:
          amount_row_na = 0
        if type(value_sheet)==str:
          if value_sheet.find('Tiếp theo') != -1:
            lst_border.append([row-1,-1])
        if count_check != 0:
          if (coordinate_cell in sheet.merged_cells) & ([row,col] not in lst_border):
            mer = col
            cell_mer=[]
            distance = getMergedCellRange(sheet, sheet[coordinate_cell])
            for cell in range(mer,mer+distance):
              cell_mer.append([row,cell])
            lst_border.extend(cell_mer)
          else:
            lst_border.append([row,col])
      if amount_row_na == 2:
        lst_border.append([row,-1])
  return lst_border

def min_border(lst_border,df):
  lst_border_new = lst_border
  data_bor = pd.DataFrame(lst_border)
  min_col = data_bor.loc[(data_bor[1]>0)][1].min()-1
  if min_col != df.columns[0]:
    value = sum(df[df.columns[0]].notna())
    if value != 0:
      index_min_col = data_bor.loc[(data_bor[1]==data_bor[1].min())].index
      for i in index_min_col:
        row_insert = data_bor[0][i]
        lst_border_new.insert(i,[row_insert,1])
  return lst_border_new

def split_format(df,lst_border):
  row = 0
  at_step = 0
  part_end, part_begin = column_end_outside(df)
  count_title = 0
  df_col = df.columns
  row_split = -1
  while (row<df.shape[0]):
    condition2 = sum(part_end.loc[row])
    try:
      condition1 = sum(part_begin[df_col[:2]].loc[row])
    except:
      condition1 = sum(part_begin[[df_col[0]]].loc[row])
    if (condition1!=0) & (condition2==0):
      count_title+=1
      if count_title == 1:
        row_split = row-1
      if row_split >0:
        at_step = 1
      else:
          at_step = 0
    else:
      count_title = 0
      if at_step != 1:
        at_step = 0
    check_na = df.notna().loc[row]
    if sum(check_na) == 0:
      at_step = 1
    if at_step == 1:
      condition = sum(df.loc[row].apply(Check_Unit))
      if (condition != 0) & (row_split != -1) :
        for i in range(len(lst_border)):
          if (row_split< lst_border[i][0]) & (row_split+1 !=0):
            lst_border.insert(i,[row_split+1,-1])
            break
    row+=1
  return lst_border

'''
Hàm drop_columns :
- Chỉ lấy những column có giá trị đường kẻ
- Reset tên cột theo thứ tự 
- Update lại list đường kẻ sau khi reset tên cột.
- Tra về list, dataframe sau xử lý.
'''
def drop_columns (lst_border,df):
  lst_border_new = []
  bor_columns_df = set([col[1]-1 for col in lst_border if col[1]>-1])
  bor_columns_excel = set([col[1] for col in lst_border if col[1]>-1])
  df_new = df[bor_columns_df]
  df_new.columns = [col for col in range(df_new.shape[1])]
  col_change = dict(list(zip(bor_columns_excel,df_new.columns+1)))
  col_change.update({-1:1})
  for col in lst_border:
    if col[0] < df_new.shape[0]:
      col[1]= col_change.get(col[1])
      lst_border_new.append(col)
  return lst_border_new,df_new

'''
Hàm bor_continue :
- Đổi các số dòng liên tục trong listt thành số dòng nhỏ nhất
'''
def bor_continue (lst_border):
  n = len(lst_border)
  temp=lst_border[0][0]
  for i in range(n):
    if (lst_border[i][0]-temp) == 1:
      lst_border[i][0] = temp
    else:
      temp = lst_border[i][0]
  return lst_border

'''
Hàm merge_bor:
- Thêm [dòng,số cột lớn nhất(độ dài đường kẻ)] có số cột bắt đầu từ 1 
- Trường hợp list mới có 1 đường kẻ mà số dòng lớn hơn số dòng đường kẻ đầu tiên list cũ thêm luôn đường kẻ đầu đó
- Trả về list đường kẻ sau khi xử lý
'''
def merge_bor(lst_border):
  bor_new = []
  temp = lst_border[0]
  value = []
  n = len(lst_border)
  for i in range(n):
    if lst_border[i][0] == temp[0]:
      value.append(lst_border[i][1])
    else:
      if min(value) == 1:
        len_bor = max(value)
        bor_new.append([temp[0],len_bor])
      value = []
      value.append(lst_border[i][1])
      temp = lst_border[i]
    if i == (n-1):
      if min(value) == 1:
        len_bor = max(value)
        bor_new.append([temp[0],len_bor])
      value = []
  if (len(bor_new)==1) & (bor_new[0][0] > lst_border[0][0]):
    bor_new.insert(0,lst_border[0])
  return bor_new

'''
Hàm real_bor:
- Lọc vị trí cắt của 'Tiếp theo' không phải vị trí cắt cuối bảng
- Vì chỉ dùng vị trí cắt của 'Tiếp theo' cho đường kẻ cuối
- Bỏ theo có số cột 1 và vị trí index số chẳn
- Trả về list sau khi xử lý 
'''
def real_bor(lst_border):
  bor_new = []
  i = 0
  while (i<len(lst_border)):
    if (lst_border[i][1] != 1) | (i%2 != 0):
      bor_new.append(lst_border[i])
    else:
      lst_border.pop(i)
      i-=1
    i+=1
  return bor_new

'''
Hàm classification_bor:
- Phân index số chẵn là đường kẻ bắt đầu thêm cả [dòng, độ dài đường kẻ]
- index lẻ là đường kẻ cuối chỉ thêm dòng
- Nếu ở list cũ độ dài index số chẵn < số lẻ thay đổi độ dài của index số lẻ
- Trả về list chứa các phần tử theo kiểu ([bắt đầu,độ dài],kết thúc)
- Nếu dư 1 phần tử cuối sẽ lưu thành ([bắt đầu,độ dài],None) 
'''
def classification_bor(lst_border):
  lst_begin = []
  lst_end = []
  if len(lst_border) > 1:
    for i in range(len(lst_border)):
      if i%2 == 0:
        lst_begin.append(lst_border[i])
      else:
        if lst_border[i-1][1] < lst_border[i][1]:
          lst_begin[i-len(lst_begin)][1] = lst_border[i][1]
        lst_end.append(lst_border[i][0])
    return list(zip_longest(lst_begin,lst_end))
  else:
    return list(zip(lst_border))

'''
Hàm column_end_bor:
- Chia 1/2 số cột (làm tròn xuống)
- Trả về list chứa số cột ở sau
'''
def column_end_bor(len_bor,total_columns):
  if len_bor%2 == 0:
    half_len_bor = len_bor/2
  else:
    half_len_bor = (len_bor/2)+0.5
  lst_columns_end = []
  for column in range(total_columns):
    if len(lst_columns_end) < (total_columns-half_len_bor):
      lst_columns_end.append(total_columns-(column+1))
  return lst_columns_end

'''
Hàm split_table:
- Cắt bảng và ghi vào file
'''
def split_table(item_test,count_shape,lst_border,name,writer):
  #######################
  total_columns = item_test.shape[1]
  clean_bor = bor_continue(lst_border)
  clean_bor = merge_bor(clean_bor)
  clean_bor = real_bor(clean_bor)
  tuple_bor = classification_bor(clean_bor)
  tuple_bor = Reverse(tuple_bor)
  ###########Dùng hàm Xử lý list đường kẻ hoàn toàn 
  for bor in range(len(tuple_bor)):
    ###########################
    len_bor = tuple_bor[bor][0][1] 
    lst_columns_end = column_end_bor(len_bor,total_columns)
    check_begin = item_test.drop(lst_columns_end,axis=1).notna()
    # Lấy độ dài, lưu các ô có giá trị ở 1/2 cột đầu
    #############################
    if len(tuple_bor[bor]) == 2:
      begin = tuple_bor[bor][0][0]-1
      if tuple_bor[bor][1] != None:
        end = tuple_bor[bor][1]-1
    else:
      begin = tuple_bor[bor][0][0]-1
    ####Lấy vị trí đầu , cuối
    ###########################
    condition = sum(item_test.notna().loc[begin])
    while (condition == 0):
      if begin < item_test.shape[0]-1:
        begin +=1
      else:
        break
      condition = sum(item_test.notna().loc[begin])
    if sum(check_begin.loc[begin]) == 0:
      if begin < item_test.shape[0]-1:
        begin +=1
    if (sum(item_test[check_begin.columns].loc[begin].apply(Check_Unit)) != 0) & (sum(item_test.notna().loc[begin])<3):
      begin += 1
    #########Kiểm tra nếu 1/2 cột đầu không có giá trị hoặc 1/2 cột đầu có dấu hiệu Đơn vị tăng vị trí +1
    #######################
    name_table = name +' Table {a} {b}'.format(a=str (bor),b=str (count_shape))
    count_shape+=1
    if bor != 0:
      ################## Nếu không là phần tử đầu (cặp đường kẻ cuối sheet)
      #################### cắt đầu ,cuối và ghi vào file
      columns= [i for i in range(len_bor)]
      table = item_test.loc[begin:end]
      table.to_excel(writer, sheet_name= name_table,columns=columns,index=False,header=False)
      item_test.drop(table.index,inplace=True)
    else:
      columns= [i for i in range(len_bor)]
      table = item_test.loc[begin:]
      if sum(table.notna().sum()) < 2:
        if len(tuple_bor)>1:
          temp = list(tuple_bor[bor+1])
          temp[1]=begin+1
          tuple_bor[bor+1] = temp
      else:
        table.to_excel(writer, sheet_name= name_table,columns=columns,index=False,header=False)
      item_test.drop(table.index,inplace=True)
      ########### Nếu là phần tử đầu (cặp đường kẻ cuối sheet)
      ############ Cắt và Kiểm tra số lượng giá trị trong dataframe 
      ############# Nếu < 2 thì lấy chuyển đường kẻ đầu thành đường kẻ cuối cho cặp sau . Không ghi vào file
  item_test.reset_index(drop=True,inplace=True)
  return item_test  


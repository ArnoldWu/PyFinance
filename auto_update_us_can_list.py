# -*- coding: utf-8 -*-
import win32com.client
import sys
import time
import re

# for 中文處理
reload(sys)
sys.setdefaultencoding('utf-8')

def stock_id_prepare(stock_id):
  return re.split(':',stock_id)[1].strip('\n')

def update_stock_id(fd, stock_id):

  lowest_value = yingzai.Sheets(tw).Range("K5").Value
  highest_value = yingzai.Sheets(tw).Range("K7").Value
  expected_gain = yingzai.Sheets(tw).Range("K4").Value

  yingzai.Sheets(tw).Range("A2").Value = stock_id

  # while Excel.CalculationState = 1:
  process = yingzai.Sheets(tw).Range("D14").Value
  new_value = yingzai.Sheets(tw).Range("K5").Value
  i = 0
  while (len(str(process))>0) and i<3 and new_value == lowest_value:
    time.sleep(5)
    process = yingzai.Sheets(tw).Range("D14").Value
    i = i+1
    new_value = yingzai.Sheets(tw).Range("K5").Value

  if i==3:
    print "Not update !!"
    return

  lowest_value = yingzai.Sheets(tw).Range("K5").Value
  highest_value = yingzai.Sheets(tw).Range("K7").Value
  expected_gain = yingzai.Sheets(tw).Range("K4").Value
  stock_id_raw = yingzai.Sheets(tw).Range("V37").Value

  line = str(stock_id_raw)+","+ str(1) + "," + str(lowest_value)+","+str(highest_value)+",\n"
  print line
  fd.write(line)


if __name__ =='__main__':
  print "Program Start !!"
  fd = open("yingzai_result_us.csv",'w')

  fd.write("Symbol, Shares, low value, high value,\n")

  Excel=win32com.client.Dispatch("Excel.Application")

  Excel.DisplayAlerts = False
  Excel.Quit

  Excel.Visible = True

  yingzai = Excel.Workbooks.Open("D:\Stock\yingzai_sg4.xls")

  for i in range(1, Excel.Sheets.Count):
    s = Excel.Sheets(i).Name
    if s == "美股":
      print "find !!"
      tw = s
      break

  print tw
  yingzai.Sheets(s).Activate


  # format w/o 交易所,
  fd_can = open('US_STOCK_CAN_LIST.TXT','r')
  for line in fd_can:
    if len(line) > 0:
      line = line.strip('\n')
      # print "line:" + line
      # print stock_id_prepare(line)
      update_stock_id(fd, line)

  fd_can.close()

  fd.close()
  Excel.Quit
  print "Program End !!"

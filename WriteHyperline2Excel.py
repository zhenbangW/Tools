import os
from shutil import copyfile
import xlrd
import xlwt
from xlwt import Formula
from xlutils.copy import copy


if __name__ == "__main__":


  path = r'F:\创新基金项目申报'
  outpath = r'F:\提交' 
  fullName = '非匿名'
  partName = '匿名'
  fullDir = os.path.join(outpath, fullName)
  partDir = os.path.join(outpath, partName)

  #---------------------------------------------#
  #    生成匿名和非匿名文件夹
  #    根据文件名最后的字母A or B 划分匿名和非匿名
  #---------------------------------------------#
  os.makedirs(fullDir)
  os.makedirs(partDir)

  allFile = os.listdir(path)
  for file in allFile:
    
    number = file.split('-')[-2]
    temp = file.split('-')[-1]

    name = temp.split('.')[0][:-1]
    type = temp.split('.')[0][-1]

    fileType = temp.split('.')[-1]

    if type == 'A':
      targetFile = os.path.join(fullDir, file)
      tempFile = os.path.join(path, file)
      copyfile(tempFile, targetFile)
    else:
      targetFile = os.path.join(partDir, file)
      tempFile = os.path.join(path, file)
      copyfile(tempFile, targetFile)

  #--------------------------------------#
  #    修改匿名文件夹中文件名称
  #    匿名文件夹中文件只留学号
  #--------------------------------------#
  fullFile = os.listdir(fullDir)
  for i in fullFile:
    number = i.split('-')[-2]
    fileType = temp.split('.')[-1]
    last = number + '.' + fileType

    old_name = os.path.join(fullDir, i)

    new_name = os.path.join(fullDir, last)
    os.rename(old_name, new_name)
  
  partFile = os.listdir(partDir)
  for j in partFile:
    number = j.split('-')[-2]
    fileType = j.split('.')[-1]
    last = number + '.' + fileType

    old_name = os.path.join(partDir, j)

    new_name = os.path.join(partDir, last)
    os.rename(old_name, new_name)


  #--------------------------------------#
  #    在Excel文件中生成超链接
  #--------------------------------------#
  oldWb = xlrd.open_workbook(os.path.join(outpath, '创新基金申请汇总表格.xls'))
  newWb = copy(oldWb)
  for i in range(4):
    table = oldWb.sheets()[i]
    sheet = newWb.get_sheet(i)
    # 按行读取
    for j in range(1, table.nrows):
      cell_value = int(table.cell(j, 2).value)
      # 多种类型文件
      filetype = ['.doc', '.docx', '.pdf']
      link_url = ''
      for m in filetype:
        temp = str(cell_value) + m
        temp_link_url = os.path.join(partDir, temp)
        if os.path.exists(temp_link_url):
          link_url = os.path.join(partName, temp)
          break
      
      if len(link_url) != 0:
        #   写入超链接
        sheet.write(j, 10, Formula('HYPERLINK("{}"; "{}")'.format(link_url, '申报书内容')))
      else:
        print(cell_value)

  newWb.save(os.path.join(outpath, '创新基金申请汇总表格2.xls'))






  # sheet = newWb.get_sheet(0)
  # link_url = os.path.join(partDir, '6120210116.docx')
  # f = f'=HYPERLINK("{link_url}","点击查看")'
  # # sheet.write(2,10,'=HYPERLINK("{link_url}","点击查看"')
  # sheet.write(2,11, Formula('HYPERLINK("{}"; "{}")'.format(link_url, '申报书内容')))
  # newWb.save(os.path.join(outpath, '1.xls'))



  # worksheet1 = workbook.sheets()[0]
  # temp = worksheet1.col_values(9)
  # worksheet1.write(2,9, 'hahah')




  pass

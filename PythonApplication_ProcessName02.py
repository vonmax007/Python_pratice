#coding = gbk
'''
Created on 2014-11-14
@author: Neo
'''


import os
import xlrd
print "\n\t************** 作业提交结果 **************\n"
#首先读取excel表格，获取选课名单;

fname = "E:\\Projects\\2016名单.xls"
bk = xlrd.open_workbook(fname)
shxrange = range(bk.nsheets)
try:
 sh = bk.sheet_by_name("page 1")
except:
 print "no sheet in %s named Sheet1" % fname
#获取行数
nrows = sh.nrows
#获取列数
ncols = sh.ncols

#print "nrows %d, ncols %d" % (nrows,ncols) #打印行列数;
print "\n\t************** 作业提交结果 **************\n"

#获取第一行第一列数据 
#cell_value = sh.cell_value(0,3)
#print cell_value
  
name_list = []
#获取各行数据
for i in range(0,nrows):
 row_data = sh.cell_value(i,3)
 #print row_data
 name_list.append(row_data)


#然后获取所有交作业的名单;
 def GetFileList(dir, fileList):
    newDir = dir
    if os.path.isfile(dir):
        fileList.append(dir.decode('gbk'))
    elif os.path.isdir(dir):  
        for s in os.listdir(dir):
            #如果需要忽略某些文件夹，使用以下代码
            #if s == "xxx":
                #continue
            newDir=os.path.join(dir,s)
            GetFileList(newDir, fileList)  
    return fileList
 
list= GetFileList('E:\\study\\组合数学\\组合数学作业\\Lecture 2', [])
submit_list = []
for e in list:
    str_homework = e.split('\\')[5]
    submit_list.append(str_homework)
    #print str_homework

#比较 submit_list 和 name_list;
Is_submit = False
count = 1
for name in name_list:
    Is_submit = False 
    for sub in submit_list:
        if name in sub:
            #print count," ", name, " —— submit !"
            Is_submit = True
            break
    if Is_submit == False : 
        print count," ", name, " —— has not submit yet  !"
    count += 1

print "\n\n\n\n\n\n"

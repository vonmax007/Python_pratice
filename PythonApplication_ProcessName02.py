#coding = gbk
'''
Created on 2014-11-14
@author: Neo
'''


import os
import xlrd
print "\n\t************** ��ҵ�ύ��� **************\n"
#���ȶ�ȡexcel��񣬻�ȡѡ������;

fname = "E:\\Projects\\2016����.xls"
bk = xlrd.open_workbook(fname)
shxrange = range(bk.nsheets)
try:
 sh = bk.sheet_by_name("page 1")
except:
 print "no sheet in %s named Sheet1" % fname
#��ȡ����
nrows = sh.nrows
#��ȡ����
ncols = sh.ncols

#print "nrows %d, ncols %d" % (nrows,ncols) #��ӡ������;
print "\n\t************** ��ҵ�ύ��� **************\n"

#��ȡ��һ�е�һ������ 
#cell_value = sh.cell_value(0,3)
#print cell_value
  
name_list = []
#��ȡ��������
for i in range(0,nrows):
 row_data = sh.cell_value(i,3)
 #print row_data
 name_list.append(row_data)


#Ȼ���ȡ���н���ҵ������;
 def GetFileList(dir, fileList):
    newDir = dir
    if os.path.isfile(dir):
        fileList.append(dir.decode('gbk'))
    elif os.path.isdir(dir):  
        for s in os.listdir(dir):
            #�����Ҫ����ĳЩ�ļ��У�ʹ�����´���
            #if s == "xxx":
                #continue
            newDir=os.path.join(dir,s)
            GetFileList(newDir, fileList)  
    return fileList
 
list= GetFileList('E:\\study\\�����ѧ\\�����ѧ��ҵ\\Lecture 2', [])
submit_list = []
for e in list:
    str_homework = e.split('\\')[5]
    submit_list.append(str_homework)
    #print str_homework

#�Ƚ� submit_list �� name_list;
Is_submit = False
count = 1
for name in name_list:
    Is_submit = False 
    for sub in submit_list:
        if name in sub:
            #print count," ", name, " ���� submit !"
            Is_submit = True
            break
    if Is_submit == False : 
        print count," ", name, " ���� has not submit yet  !"
    count += 1

print "\n\n\n\n\n\n"

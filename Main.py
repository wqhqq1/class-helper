import xlrd as excel_reader
from os import remove as delete
from os import path
from random import sample as rand
from tkinter import *
from tkinter import filedialog
from tkinter import messagebox
import json
chosefilename = ''
out = ''
showchooser = False
english = False

def filechoose():
    global chosefilename
    chosefilename = filedialog.askopenfilename(filetypes = [('Excel表格文件', '*.xlsx *.xls')])
    global status
    if chosefilename == '':
        return
    try:
        delete('nameData')
    except:
        pass
    try:
        names = excel_reader.open_workbook(r'%s' % chosefilename)
        names = names.sheet_by_name('Sheet1')
    except:
        status = False
        statusLabel.config(text='未导入姓名表', fg='red')
        messagebox.showerror('Fatal Error', '请检查文件是否存在!')
        return
    lst = {'orig':[], 'modified':[], 'except':[]}
    n = 0
    while True:
        try:
            lst['orig'].append(names.cell(n + 1, 0).value)
            n += 1
        except:
            break
    lst['modified'] = lst['orig']
    with open('nameData', 'w') as f:
        json.dump(lst, fp = f)
    status = True
    statusLabel.config(text = '已导入姓名表', fg = 'black')
    messagebox.showinfo(title = 'Info', message = '导入成功!')

def rand_chooser(lst, num = 0):
    n = 0
    rands_temp = [i for i in range(1,len(lst)+1)]
    rands = rand(rands_temp, num)
    return rands

def pickup():
    global status
    if status == False:
        messagebox.showerror(title='Fatal Error', message='你还没导入姓名表')
        return
    with open('nameData', 'r') as f:
        nameLst = json.load(fp = f)
    # print(nameLst)
    # return
    try:
        pickupNum = int(numberEntry.get())
    except:
        messagebox.showerror(title = 'Fatal Error', message = '点名人数必须为整数!')
        return
    if pickupNum <= 0:
        messagebox.showerror(title='Fatal Error', message='点名人数必须大于0!')
        return
    if pickupNum > len(nameLst['orig']):
        messagebox.showerror(title='Fatal Error', message='点名人数超过总人数')
        return
    global checked
    if not checked.get():
        # print(len(nameLst['modified']))
        if len(nameLst['modified']) == 0:
            nameLst['modified'] = []
            nameLst['except'] = []
            for i in nameLst['orig']:
                nameLst['modified'].append(i)
            # print('again', nameLst)
            again = messagebox.askokcancel(title = '提示', message = '所有学生均抽完，要重新开始吗？')
            if not again:
                return
        if len(nameLst['modified']) < pickupNum:
            randData = rand_chooser(nameLst['except'], pickupNum - len(nameLst['modified']))
            # print(randData)
            for i in randData:
                nameLst['modified'].append(nameLst['except'][i - 1])
            randData.sort()
            offset = 1
            for i in randData:
                nameLst['except'].pop(i - offset)
                offset += 1
            again = messagebox.askokcancel(title = '提示', message = '点名人数超过未抽人数，要从抽过的人当中抽取吗?')
            if not again:
                return

        pickedNums = rand_chooser(nameLst['modified'], pickupNum)
        pickedNames = ''
        # print(pickedNums)
        for i in pickedNums:
            pickedNames += nameLst['modified'][i - 1]
            nameLst['except'].append(nameLst['modified'][i - 1])
            # print(i)
            if i != pickedNums[len(pickedNums) - 1]:
                pickedNames += ','
        # print(pickedNames)
        # print(len(nameLst))
        # print(pickedNums)
        pickedNums.sort()
        offset = 1
        for i in pickedNums:
            nameLst['modified'].pop(i - offset)
            offset += 1
        # print(nameLst)
    else:
        pickedNames = ''
        randData = rand_chooser(nameLst['orig'], pickupNum)
        for i in randData:
            pickedNames += nameLst['orig'][i - 1]
            if i != randData[len(randData) - 1]:
                pickedNames += ','
    try:
        delete('nameData')
    except:
        pass
    with open('nameData', 'w') as f:
        json.dump(nameLst, fp = f)
    rltDLabel.config(text = '%s同学%s' % (pickedNames, leaveWordsEntry.get()))
    messagebox.showinfo(title = 'Info', message = '%s同学%s' % (pickedNames, leaveWordsEntry.get()))

def restore():
    deleeteOrReset = messagebox.askyesnocancel(title = '还原', message = '重置到为刚导入状态请选Yes，清除所有数据请选No，取消请选Cancel')
    if deleeteOrReset == None:
        return
    global status
    if status == False:
        messagebox.showerror(title = 'Fatal Error', message = '你还没导入姓名表')
        return
    if not deleeteOrReset:
        try:
            delete('nameData')
        except:
            pass
        status = False
        statusLabel.config(text = '未导入姓名表', fg = 'red')
    else:
        with open('nameData', 'r') as f:
            nameLst = json.load(fp = f)
        try:
            delete('nameData')
        except:
            pass
        with open('nameData', 'w') as f:
            nameLst['modified'] = []
            nameLst['except'] = []
            for i in nameLst['orig']:
                nameLst['modified'].append(i)
            json.dump(nameLst, fp = f)
        # print(nameLst)


root = Tk()
root.title('公平点名器')
excelPathLabel = Label(root, text = '导入姓名表:')
excelPathLabel.grid(row = 0, column = 0)
status = False
if path.isfile('nameData'):
    status = True
statusLabel = Label(root, text = '未导入姓名表', fg = 'red')
if status:
    statusLabel.config(text = '已导入姓名表', fg = 'black')
statusLabel.grid(row = 0, column = 1, sticky = W)
chooseFileButton = Button(root, text = '选择文件', command = filechoose)
chooseFileButton.grid(row = 0, column = 2)
numberLabel = Label(root, text = '点名人数:')
numberLabel.grid(row = 1, column = 0)
numberEntry = Entry(root)
numberEntry.grid(row = 1, column = 1)
numberEntry.insert(0, '1')
leaveWordsLabel = Label(root, text = '问题/留言:')
leaveWordsLabel.grid(row = 2, column = 0)
leaveWordsEntry = Entry(root)
leaveWordsEntry.grid(row = 2, column = 1)
leaveWordsEntry.insert(0, '请起立回答问题')
rltLabel = Label(root, text = '结果:')
rltLabel.grid(row = 3, column = 0)
rltDLabel = Label(root)
rltDLabel.grid(row = 3, column = 1, sticky = W)
checked = BooleanVar()
origCButton = Checkbutton(root, text = '使用传统方式', variable = checked)
origCButton.grid(row = 4, column = 0)
restoreButton = Button(root, text = '还原', command = restore)
restoreButton.grid(row = 5, column = 2)
pickupButton = Button(root, text = '点名', command = pickup)
pickupButton.grid(row = 5, column = 1)
exitButton = Button(root, text = '退出', command = root.destroy)
exitButton.grid(row = 5, column = 0)
root.mainloop()
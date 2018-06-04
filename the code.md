    import xlrd
    #import os
    #import xlwt
    import xlutils.copy
    import tkinter as tk
    import openpyxl as op


    from tkinter import *
    from tkinter import filedialog

    def dashbroken(datawithdash):

        

        list1=datawithdash.split('-')

        

        list0=list1[0]
        list1=list1[1]
        
        abc=list0
        
        cba=abc
        str2=cba[::-1]
        ln=len (str2)

        


        for ii in range(ln):
            
            if str.isdigit(str2[ii]) is False:


                 
                break

        str3=str2[0:ii]
        str4=str2.strip(str3)


        str5=str4[::-1]    
        head=str5

        #print(head)
        
        
        start=list0.strip(head)
        end=list1.strip(head)
        #print(start)
        start=int(start)
        end= int (end)+1
        list99= range (start,end)
        listcycle= range (1, end-start+1)
        listmax=[]


        for i in listcycle:
        
            lists=head+str(list99[i-1])
            listmax.append(lists)
             
        datawithdash=listmax
        return listmax 


       
    def dashreplace(dada):



        dada=dada.split(',')

        for i in range(len(dada)):
            if dada[i].find('-') >0:
            
               dashbroken(dada[i])
           
               listadd= dashbroken(dada[i])
               strr= ','.join(listadd)
               dada[i]=strr
            else:

                pass

        dadastrr= ','.join(dada)
        return dadastrr



    def openfiles2():

        s2fname = filedialog.askopenfilename(title='待转化的文件', filetypes=[('bom', '*.xls'), ('bom', '*.xlsx'),('All Files', '*')])
        print(s2fname)
        return s2fname
        
    def button2():

        filename=openfiles2()
        
        book = op.load_workbook(filename)
        print ("表单数量:", book.nsheets)
        print ("表单名称:", book.sheet_names())
        sh1 = book.sheet_by_index(1)

        print ("表单 %s 共 %d 行 %d 列" % (sh1.name, sh1.nrows, sh1.ncols))

       

        
        #newbom = xlwt.Workbook()
        newbom = xlutils.copy.copy(book)
        #sheet1 = newbom.add_sheet(u'sheet1',cell_overwrite_ok=True)
        sheets2=newbom.get_sheet_names()
        sh2=newbom.get_sheet_by_name(sheets2[1])
        n=input('输出第几列： ')
        n=int(n)
        num=0

        #for s in book.sheets():

            
           
        wrt11=sh1.col(n)


        for m in range(len(wrt11)):      
            q=int(m)
            p=wrt11[m].value
        
            bestdata=dashreplace(p)
        
            #sh2.write(m,n+1,bestdata)
            sh2.cell(row=m,column =n+1).value=bestdata
        newbom.save('newbom.xlsx')

        print ("OK")
            
            
    root = tk.Tk()
    root.title("DASHBROKEN V3.0 by Taylor")
    root.geometry('500x300+500+200')
    #btn1 = tk.Button(root, text='带转化BOM',font =("宋体",20,'bold'),width=13,height=8, command=openfiles2)

    btn2 = tk.Button(root, text='搞起！',bg="pink",font =("宋体",15,'bold'),width=10,height=1, command=button2)
    #btn1.pack(side=tk.LEFT)
    btn2.pack(side=tk.LEFT)
    t =Text()
    t.pack(side=tk.RIGHT)
    t.insert(1.0,'1. 点击“搞起”按钮，选择需要转化的Micron BOM\n',2.0,'2. 在对话框中填写位置所在的列,A列填0，G列填6\n',3.0,'3. 生成的newbom.csv为转化后的\n')

    root.mainloop()


import MySQLdb
import xlrd
import  xlsxwriter as wx


_version_='1.1.1'

def mstoxls(cnx,tbname,xlsname=''):

    if cnx:
        cursor=cnx.cursor()
    else:
        print 'mysql con not been connected'
        pass
    sqlsent="select * from "+tbname
    data=(tbname,)
    cursor.execute(sqlsent)
    if not xlsname:
        xlsname=tbname
    try:
        dir1=''
        wb = wx.Workbook(dir1+xlsname+'.xls')
        ws=wb.add_worksheet(tbname)
        x1=1
        x3=0
        for i2 in cursor.description:
            ws.write(0,x3,i2[0])
            x3=x3+1
        for i in cursor.fetchall():
            x2=0
            for i1 in i:
                ws.write(x1,x2,i1)
                x2=x2+1
            x1=x1+1
        wb.close()
        cursor.close()
        cnx.close()
    except:
        print 'file cannot been created'
        pass

def xlstoms(cnx,xlsname,tbname=''):
    if cnx:
        cursor1=cnx.cursor()
    else:
        print 'mysql con not been connected'
        pass


    data=xlrd.open_workbook(xlsname)
    table = data.sheet_by_index(0)
    listr=table._cell_values
    list_row1=listr[0]
    listn=listr[1:]
    ziduan=",".join(list_row1)
    print ziduan


    vals_li=[]
    ti=()
    for i in listn:
        print str(i)

        ms=','.join(str(i))
        ti=(ms)
        print ti
    vals_li.append(ti)
    print vals_li


    sql_set="insert into "+tbname+"("+ziduan+") values (%s)"
    cursor1.executemany(sql_set,vals_li)
    # cnx.commit()
    print sql_set


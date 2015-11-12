import MySQLdb
import xlrd
import  xlsxwriter as wx

host='127.0.0.1'
username='root'
password='sjtuld0218'
charset='utf8'
cnx=MySQLdb.connect(host,username,password,'jdbcstudy',charset=charset)
def mstoxls(tbname,xlsname=''):

    if cnx:
        cursor=cnx.cursor()
    sqlsent="select * from "+tbname
    data=(tbname,)
    cursor.execute(sqlsent)
    if not xlsname:
        xlsname=tbname
    try:
        wb = wx.Workbook(xlsname+'.xls')
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


mstoxls("info")


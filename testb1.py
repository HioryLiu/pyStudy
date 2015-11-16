from msxls import *


host='127.0.0.1'
username='root'
password='****'
charset='utf8'
cnx=MySQLdb.connect(host,username,password,'jdbcstudy',charset=charset)
# mstoxls(cnx,"info")
xlstoms(cnx,"info.xls")
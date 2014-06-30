#coding=utf-8 
#!/usr/bin/env python
import sys
reload(sys)
sys.setdefaultencoding('utf-8')
#-------------------------------------------------------------------------------
# Name: erp.py
# Purpose: 修改 用友ERP帐套 的帐套号、制单人、审核人 等信息
#
# Author: boollab
#
# Created: 12/12/2013
#-------------------------------------------------------------------------------

import pyodbc
import pymssql

class MSSQL:
    """

    对pyodbc的简单封装
    用法：

    """

    def __init__(self,host,user,pwd,db):
        self.host = host
        self.user = user
        self.pwd = pwd
        self.db = db

    def __GetConnect(self):
        """
        得到连接信息
        返回: conn.cursor()
        """
        if not self.db:
            raise(NameError,"没有设置数据库信息")
        
        conn_info = "DRIVER={SQL Server};SERVER=%s;UID=%s;PWD=%s;DATABASE=%s;CHARSET=UTF8" % (self.host, self.user, self.pwd, self.db)
        self.conn = pyodbc.connect(conn_info)
        cur = self.conn.cursor()
        if not cur:
            raise(NameError,"连接数据库失败")
        else:
            return cur

    def ExecQuery(self,sql):
        """
        执行查询语句
        返回的是一个包含tuple的list，list的元素是记录行，tuple的元素是每行记录的字段

        调用示例：
                ms = MSSQL(host="localhost",user="sa",pwd="123456",db="PythonWeiboStatistics")
                resList = ms.ExecQuery("SELECT id,NickName FROM WeiBoUser")
                for (id,NickName) in resList:
                    print str(id),NickName
        """
        cur = self.__GetConnect()
        cur.execute(sql)
        resList = cur.fetchall()

        #查询完毕后必须关闭连接
        self.conn.close()
        return resList

class MSSQL2:
    """
    对pymssql的简单封装
    pymssql库，该库到这里下载：http://www.lfd.uci.edu/~gohlke/pythonlibs/#pymssql
    使用该库时，需要在Sql Server Configuration Manager里面将TCP/IP协议开启

    用法：

    """

    def __init__(self,host,user,pwd,db):
        self.host = host
        self.user = user
        self.pwd = pwd
        self.db = db

    def __GetConnect(self):
        """
        得到连接信息
        返回: conn.cursor()
        """
        if not self.db:
            raise(NameError,"没有设置数据库信息")
        self.conn = pymssql.connect(host=self.host,user=self.user,password=self.pwd,database=self.db,charset="utf8")
        cur = self.conn.cursor()
        if not cur:
            raise(NameError,"连接数据库失败")
        else:
            return cur

    
    def ExecNonQuery(self,sql):
        """
        执行非查询语句

        调用示例：
            cur = self.__GetConnect()
            cur.execute(sql)
            self.conn.commit()
            self.conn.close()
        """
        cur = self.__GetConnect()
        cur.execute(sql)
        self.conn.commit()
        self.conn.close()

def update_value(table_name, field_name, author, yours, db):
    up = MSSQL2(host="localhost",user="sa",pwd="",db=db)

    sql = "update %s set %s='%s' where %s='%s'" % (table_name, field_name, yours.decode('gb2312','ignore').encode('utf-8','ignore'), field_name, author.decode('gb2312','ignore').encode('utf-8','ignore'))
    try:
        up.ExecNonQuery(sql)
        print "Update table:%s, field:%s Successful!" % (table_name, field_name)
    except Exception,e:
        print sql
        print 'me=',yours.decode('gb2312','ignore').encode('utf-8','ignore'),"he=", author.decode('gb2312','ignore').encode('utf-8','ignore')
        print "Update error: %s",e[1].decode('utf-8','ignore').encode('gb2312','ignore')
    finally:
        del up
        del sql

def change_zt(author, yours, tadexuehao, nidexuehao):
    while 1:
        print '='*79
        print '*'+'Step 2：修改帐套名称、总管、日志'.decode('utf-8').encode('gb2312').center(77) +'*'
        print '='*79
        print "1.以admin身份进入‘系统管理’，导入别人的帐套。".decode('utf-8','ignore').encode('gb2312','ignore')
        print "2.一般第一次导入会报错，再导一次就能成功。".decode('utf-8','ignore').encode('gb2312','ignore')
        print '='*79
        ask = raw_input('如果帐套导入成功，请输入yes：'.decode('utf-8','ignore').encode('gb2312','ignore'))

        if ask == 'yes':
            database1 = "UFDATA_%s_2008" % nidexuehao
            database2 = "UFSystem"

            zt_name1 = "%s0106%s" % (author,tadexuehao)
            zt_name2 = "%s0106%s" % (yours, nidexuehao)

            # 修改帐套名称
            update_value('UA_account_ex','cacc_name', zt_name1, zt_name2, database1)
            update_value('UA_account','cacc_name', zt_name1, zt_name2, database2)
    
            # 修改帐套总管
            update_value('UA_user','cUser_Name', author, yours, database2)
            print '='*79
            print '*'+'帐套名称、总管、日志修改成功！'.decode('utf-8').encode('gb2312').center(77) +'*'
            break



def main(xuehao, author, yours,tadexuehao):
    ## ms = MSSQL(host="localhost",user="sa",pwd="123456",db="PythonWeiboStatistics")
    ## #返回的是一个包含tuple的list，list的元素是记录行，tuple的元素是每行记录的字段
    ## ms.ExecNonQuery("insert into WeiBoUser values('2','3')")
    print '='*79
    print '*'+'Step 3：修改帐套内容'.decode('utf-8').encode('gb2312').center(77) +'*'
    print '='*79

    database = "UFDATA_%s_2008" % xuehao
    #database = "UFMeta_%s" % xuehao
    #database = "UFSystem"

    #zt_name = "%s0106%s" % (author, tadexuehao)
    try:
        ms = MSSQL(host="localhost",user="sa",pwd="",db=database)
        print "U8 Connect Successful:"
        print "Start change infomation..."
    except Exception,e:
        raw_input("U8 Connect Error:")
        return
    
    table_names = ms.ExecQuery("select name from sysobjects where xtype='U'")

    
    for name in table_names:
        table_name = name[0].encode('utf-8','ignore')

        rec = ms.ExecQuery("select * from [%s]" % table_name)
        if not rec:
            #print "%s is empty, skip..." % table_name
            continue
        
        try:
            field_names = ms.ExecQuery("select name from syscolumns where id in(select id from sysobjects where id = object_id('%s'))" % table_name)
        except Exception,e:
            pass
            #print e[1].decode('utf-8','ignore').encode('gb2312','ignore')
            #print "Field_names Error tablename:", table_name,"fields_names:", field_names
            #raw_input("#")

        for field in field_names:
            field_name = field[0].encode('utf-8','ignore')
            try:
                records = ms.ExecQuery("select [%s] from [%s]" % (field_name, table_name))
            except Exception,e:
                pass
                #print "get records:",e[1].decode('utf-8','ignore').encode('gb2312','ignore')
                #print "field_name=%s, table_name=%s" % (field_name, table_name)
                #raw_input("get records error:")
            else:
                for value in records:
                    if str(value[0]) == author.decode('gb2312','ignore').encode('utf-8','ignore'):
                        print "Find one in: table=%s,field=%s" % (table_name, field_name)
                        update_value(table_name, field_name, author, yours, database)
                del records
        print "%s Successful!" % table_name
    print '='*79
    print '*'+'帐套内容修改成功！'.decode('utf-8').encode('gb2312').center(77) +'*'
    print '='*79


def get_info():
    while 1:
        tadexuehao = raw_input('请输入他的帐套号：'.decode('utf-8','ignore').encode('gb2312','ignore'))
        if tadexuehao.isdigit() and len(tadexuehao)==3:
            break
        else:
            print "您输入的帐套号不正确!".decode('utf-8','ignore').encode('gb2312','ignore')

    while 1:
        nidexuehao = raw_input('请输入你的帐套号：'.decode('utf-8','ignore').encode('gb2312','ignore'))
        if nidexuehao.isdigit() and len(nidexuehao)==3:
            break
        else:
            print "您输入的帐套号不正确!".decode('utf-8','ignore').encode('gb2312','ignore')
    

    while 1:
        print '='*79
        print '*'+'Step 1：修改帐套号'.decode('utf-8').encode('gb2312').center(77) +'*'
        print '='*79
        
        print "1.在备份帐套文件夹下新建new.txt".decode('utf-8','ignore').encode('gb2312','ignore')
        print "2.在文件的第一行写上他跟你的名字".decode('utf-8','ignore').encode('gb2312','ignore')
        print "3.以空格分隔，例如：张文 唐家洪".decode('utf-8','ignore').encode('gb2312','ignore')
        print '='*79
        ask = raw_input("如果都弄好了，请输入yes：".decode('utf-8','ignore').encode('gb2312','ignore'))
        
        if ask == 'yes':
            f = open('new.txt')
                
            names = f.readline().decode('gb2312','ignore').encode('utf-8','ignore').split()
            #print names
            tadexingming = names[0].decode('utf-8','ignore').encode('gb2312','ignore')
            nidexingming = names[1].decode('utf-8','ignore').encode('gb2312','ignore')
            #print tadexingming,nidexingming
            if tadexingming and nidexingming:
                break
    return tadexuehao, tadexingming, nidexuehao, nidexingming

def change_lst(tadexuehao, nidexuehao):
    import os
    if os.path.exists('UfErpAct.Lst'):
        with open('UfErpAct.Lst') as f:
            content = f.read()
            tmp1 = content.replace('cAcc_Id=%s' % tadexuehao, 'cAcc_Id=%s' % nidexuehao)
            tmp2 = tmp1.replace('ZT%s' % tadexuehao, 'ZT%s' % nidexuehao)
        os.remove('UfErpAct.Lst')
        new = open('UfErpAct.Lst', 'w')
        new.write(tmp2)
        new.close()
        print '='*79
        print '*'+'帐套号修改成功！'.decode('utf-8').encode('gb2312').center(77) +'*'
    else:
        print '*'+'请把程序放入备份帐套文件夹，并且文件夹名称不能包含中文。'.decode('utf-8').encode('gb2312').center(77) +'*'
        raw_input("按回车键退出程序：".decode('utf-8','ignore').encode('gb2312','ignore'))
        raise SystemExit()

if __name__ == '__main__':
    # 显示开头
    print '='*79
    print '*'+'用友ERP帐套修改器'.decode('utf-8').encode('gb2312').center(77) +'*'
    print '='*79
    
    # 获取信息
    info = get_info()
    
    tadexuehao = info[0]
    tadexingming = info[1]
    nidexuehao = info[2]
    nidexingming = info[3]
    
    change_lst(tadexuehao, nidexuehao)

    try:
        # step 2:
        change_zt(tadexingming, nidexingming, tadexuehao, nidexuehao)

        # step 3:
        main(nidexuehao, tadexingming, nidexingming, tadexuehao)
    except Exception,e:
        print "main() exist error: %s" % e[1].decode('utf-8','ignore').encode('gb2312','ignore')
    raw_input("end:")

# -*-:coding: UTF-8 -*-
import xlrd
import os
import sys

argv = sys.argv
"""
所用到的所有外部参数
"""
class Conf(object):
    field_information = "字段信息"#sheet名
    field = "字段"#列名
    field_type = "字段类型"#列名
    pk = "是否主键"#列名
    table_information = "表信息";#sheet名
    excelPath = argv[1]#工作表的绝对路径（注意路径格式》D:/TestFile/20160503v4.9.xls）
    tableSys = argv[2]#表所属的系统层级（例如：f_fdm，不区分大小写）
    
"""
从工作表中抓取sheet和table的类
"""
class Get_xlsx(object):
    
    """
    field_information = "字段信息"#sheet name
    field = "字段"#column name
    field_type = "字段类型"#column name
    pk = "是否主键"##column name
    """
    def __init__(self):
        pass   
    """
    #不能实现自动化
    def get_sheet_by_index(self):
        data = xlrd.open_workbook(Conf.excelPath) #"D:/TestFile/20160503v4.9.xls"
        sheet = data.sheets()
        return sheet
    """
    def get_sheet_by_name(self,sheetName):
        """
        explain:通过sheet名字获取sheet工作表
        """
        data  = xlrd.open_workbook(Conf.excelPath)
        sheet = data.sheet_by_name(sheetName)
        return sheet
    
    def get_row_values(self,sheetName,index=0):
        """
        expain:获取sheet工作表指定行的所有值，并返回list。默认返回首行
        """
        sheet = self.get_sheet_by_name(sheetName)
        temp = sheet.row_values(index)
        return temp
           
    def get_col_values(self,sheetName,columnName=0):
        """
        explain:获取sheet工作表指定列的值，并返回list。默认返回首列
        """
        first_line = self.get_row_values(sheetName)
        col_column_index = first_line.index(columnName) 
        sheet = self.get_sheet_by_name(sheetName)
        temp = sheet.col_values(col_column_index)[1:]
        return temp
    """    
    def get_col_type(self):
        first_line = self.get_row_values(Conf.field_information)
        col_type_index = first_line.index(Conf.field_type)
        sheet = self.get_sheet_by_name(Conf.field_information)
        temp = sheet.col_values(col_type_index)[1:]
        
        return temp
    """
    def get_pk(self):
        first_line = self.get_row_values(Conf.field_information)
        col_pk_index = first_line.index(Conf.pk)
        sheet = self.get_sheet_by_name(Conf.field_information)
        temp = sheet.col_values(col_pk_index)[1:]
        return temp
"""
单独处理order by和 segmented 部分的类
以及实现list大小写转化
"""
class Tool(Get_xlsx):
    """
    field_information = "字段信息"#sheet name
    field = "字段"#column name
    field_type = "字段类型"#column name
    pk = "是否主键"##column name
    """
    def __init__(self):
        pass
        
    def addPk_ddl(self,pkValues):
        """
        explain:返回每张表的order by 和 segmented by hash 的部分
        """
        ddlPk = ''
        n=0
        col_values = self.get_col_values(Conf.field_information,Conf.field)
        for pk in pkValues:
            if(pk=="PK"):
                ddlPk = ddlPk+col_values[n]+","
            else:
                continue
            n=n+1
        ddlPk = ddlPk[:-1]
        ddlPk ="order by "+ddlPk+"\nsegmented by hash("+ddlPk+") all nodes ;\n\n"  
        return ddlPk
    
    def convertList(self,lname,tkname):
        """
        lname:包含两list的list
        explain:纵向合并两个list以实现两个list对应元素的顺序位置的拼接
        """
        ddl = ''
        temp = list(zip(lname[0],lname[1]))#纵向合并两个list
        #print(temp)
        for tname in temp:
            ddlStr = "\t"+tname[0]+"  "+tname[1]+",\n"
            ddl = ddl+ddlStr
            #print(ddl)
        ddl = ddl[:-2]
        ddl = "create table "+Conf.tableSys+"."+tkname+"(\n"+ddl+"\n)"
        return ddl
    
    def lower_list(self,lname,u="L"):
        """
        explain:把list中的字符进行大小写转换
        """
        index  = 0
        if(u=="U"):
            for value in lname:
                lname[index]=value.upper()
                index=index+1
        elif(u=="L"):
            for value in lname:
                lname[index]=value.lower()
                index=index+1
        return lname

"""
最终实现输出ddl的类
"""  
class Get_ddl(Tool):
    """
    field_information = "字段信息"#sheet name
    field = "字段"#column name
    field_type = "字段类型"#column name
    pk = "是否主键"##column name
    table_information="表信息";#sheet name
    """
    def __init__(self):
        pass
    
    def table_order(self):
        """
        explain:返回每张表在Excel中的起始位置和总长度的字典
        """
        tempDict = {}
        #读取“表信息”sheet页
        table_information = self.get_sheet_by_name(Conf.table_information).col_values(2)[1:]#应该自动化判断，而不应该写死
        #获取“字段信息”sheet页中的“表英文名”列的list
        col_information_table_name = self.get_sheet_by_name(Conf.field_information).col_values(0)#应该自动化判断，而不应该写死
        #对表名进行统一大小写
        col_information_table_name = self.lower_list(col_information_table_name)
        table_information = self.lower_list(table_information)
        try:
            for tname in table_information:
                st = col_information_table_name.index(tname)
                num = col_information_table_name.count(tname)
                tempDict[tname]=[st,num]
                #print(tempDict)
            return tempDict
        except  ValueError as e:
            if(hasattr(e, "code")):
                print(e.code)
            if(hasattr(e, "reason")):
                print(e.reason)
        finally:
            return tempDict
    
    def get_table_ddl(self):
        """
        explain:返回拼接完全的ddl
        """
        tempFile = os.path.split(argv[1])
        ddlFile = tempFile[0]+"/ddl.sql"
        if(os.path.exists(ddlFile)):
            os.remove(ddlFile)
        
        col_values = self.get_col_values(Conf.field_information,Conf.field)
        col_type = self.get_col_values(Conf.field_information,Conf.field_type)
        pk_values = self.get_pk()
        
        table_infomation = self.table_order()
        
        Fdd = open(ddlFile,"a+")
        
        for tkname in table_infomation:
            print("--正在生成：",tkname,"的建表语句\n")            
            table_col_values = col_values[table_infomation[tkname][0]:table_infomation[tkname][0]+table_infomation[tkname][1]-1]
            table_col_type = col_type[table_infomation[tkname][0]:table_infomation[tkname][0]+table_infomation[tkname][1]-1]            
            tempList = [table_col_values,table_col_type]
            
            table_pk_values =pk_values[table_infomation[tkname][0]:table_infomation[tkname][0]+table_infomation[tkname][1]-1]
            
            #print(tempList)
            ddl = self.convertList(tempList,tkname) 
            ddlPk = self.addPk_ddl(table_pk_values)
            ddl = ddl +"\n"+ddlPk
            
            #保存ddl
            Fdd.write(ddl)
            print(ddl)
        Fdd.close()
        #return ddl   
g = Get_ddl()
g.get_table_ddl()

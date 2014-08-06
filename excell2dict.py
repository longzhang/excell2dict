# -*-coding:utf-8 -*-
'''
Created on 2013-1-31

@author: zhanglong
'''
import xlrd
import pprint
import sys
import json
import time
import datetime
import os 

def format_dict(obj, indent=0, unit='  ', out=sys.stdout , first = True):
    if isinstance(obj, dict):
        print >> out, "{"
        curr_indent = indent + 1

        keys = obj.keys()
        keys.sort()
        for k in keys:
            v = obj[k]
            
            if isinstance(k, int):
                print >> out, "%s%d:" % (unit * (indent+curr_indent), k),
            elif isinstance(k, str):
                print >> out, "%s'%s':" % (unit * (indent+curr_indent), k),
            elif isinstance(k, float):
                print >> out, "%s'%s':" % (unit * (indent+curr_indent), k),
            elif isinstance(k, bool):
                print >> out, "%s%s':" % (unit * (indent+curr_indent), k),
            elif isinstance(k,unicode) :
                print >> out, "%s'%s':" % (unit * (indent+curr_indent), k),
            else:
                raise Exception("Error: invalid key type %s for key %s" % (type(k), str(k)))
            format_dict(v, indent=curr_indent, unit=unit, out=out , first = False)
        if first :
            print >> out, "%s}" % (indent * unit)
        else :     
            print >> out, "%s}," % (indent * unit)
    
    elif isinstance(obj,list) :
        print >> out , "%s," % obj
    elif isinstance(obj, float):
        print >> out, "%s," % obj

    elif isinstance(obj, int):
        print >> out, "%d," % obj

    elif isinstance(obj, unicode):
        print >> out, "u'%s'," % obj.encode('utf8').replace("'", "\\'")
    elif isinstance(obj, str):
        print >> out, "'%s'," % obj.replace("'", "\\'")


class ConvertExcel(object):
    
    

    def __init__(self,name):
        
        self.xml_name = name
        self.convert_res = {}
        
    
    def format_values(self, sheet_name ,filds  , values ,colum_format=None, row=None):
        key = None
        data ={}
        formatter = colum_format
        for i,value in enumerate(values) :
            try :

                format_type = formatter[i]
                if format_type == 'int' :
                    if value : 
                        values[i] = int(value)
                    else :
                        values[i] = 0 
                if format_type == 'long' :
                        values [i] = Long(value)
                if format_type == 'float' :
                    values[i] = float(value)
                if format_type == 'str' :
                    values[i] = value.strip()
                if format_type == 'arr_int' :
                    values[i]  = [int(x) for x in str(value).split(',') ]
                if format_type == 'arr_str' :
                    values[i]  = [str(x) for x in str(value).split(',') ]
                if format_type == 'dict':
                    if value :
                        values[i] = dict(eval(value))
                    else :
                        values [i] = value
                if format_type == 'list':
                    if value :
                        values[i] = list(eval(value))
                    else : 
                        values [i]  = value
                if i == 0 :
                    key = values[i]

                dkey = str(filds[i]).encode('utf-8')
                data[dkey] = values[i]
            except:
                print "#"*10
                print sys.exc_info()[0]
                print "format_type is %s  , sheet is  %s  , value is : '%s' ,  row: is %s , fild is %s " % (format_type,sheet_name, value , row , filds[i])
                print "#"*10
                raise

        return {key : data}

        
        
    def converte(self):
        """针对k-v格式且数据格式统一的转换函数，支持遍历多个sheet
                            输出格式示例
            1:["头盔", 1],
            2:["武器", 1],
            3:["魂器", 1]
        """
        xls = xlrd.open_workbook(self.xml_name)
        sheets_num = xls.nsheets
        for i in range(sheets_num) :
            sheet = xls.sheet_by_index(i)
            sheet_name = sheet.name
            data = {}
            sheet_filds = []
            colum_format = None
            for row in range(0 ,sheet.nrows) :
                if row  == 0 : 
                    continue 
                if row == 1 :
                    sheet_filds = sheet.row_values(row)
                    sheet_filds = [str(x).strip() for x in sheet_filds]
                    continue
                if row == 2 :
                    colum_format = sheet.row_values(row)
                    colum_format = [str(x) for x in colum_format]
                    continue
                row_values = sheet.row_values(row)
                formated_value = self.format_values(sheet_name, sheet_filds, row_values , colum_format, row)
                data.update(formated_value)
            self.write_py_file(sheet_name, data)
            
    def write_py_file(self,name,data):
        
        f_name = 'config_'+name+'.py'
        
        if os.path.isfile(f_name) :
            model_name = f_name.split('.')[0]
            x = __import__( model_name, globals(), locals())
            config_value = x.config
            del config_value['version']
            if config_value == data :
                return
        data['version'] = datetime.datetime.now().strftime('%Y%m%d%H%M%S')
        f = open('config_'+name+'.py','w')
        print >> f , "#-*- coding:utf-8 -*-"
        print >> f, "config = ".replace("'", "\\'"),
        format_dict(data , out=f)
        f.close()
        print 'finish export  : %s' % (name)

if __name__ == '__main__' :
    ce = ConvertExcel('growth.xlsx')
    ce.converte()
    print 'all done !!!'



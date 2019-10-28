# -*- coding: utf-8 -*- 

import sys
import os
import xlrd
import re

reload(sys)
sys.setdefaultencoding("utf-8")


# 当前脚本路径
curpath = os.path.dirname(os.path.abspath(sys.argv[0]))

# 文件头描述格式化文本
lua_file_head_format_desc = '''--[[

        %s
        exported by excel2lua.py
        from file:%s

--]]\n\n'''

# 遍历输入路径下excel文件导出
def foreaExcelFile(input_path, out_path):
    cmdPath = "rm -rf "+out_path+"/*"
    os.system(cmdPath) 
    for root, dirs, files in os.walk(input_path):
        for file in files :
            if (file != ".DS_Store") :
                if  os.path.exists(out_path) == False :
                    os.makedirs(out_path)
                luaname = file.replace(".xlsx", ".lua")
                luaname = luaname.replace(".xls", ".lua")
                excel2lua(root+"/"+file,out_path+"/"+luaname)


# 将数据导出到tgt_lua_path
def excel2lua(src_excel_path, tgt_lua_path):
    # print('[file] %s -> %s' % (src_excel_path, tgt_lua_path))s
    # load excel data
    excel_data_src = xlrd.open_workbook(src_excel_path, encoding_override = 'utf-8')
    # print('[excel] Worksheet name(s):', excel_data_src.sheet_names())
    lua_str = "local " + "data" + " = {"
    for sheet in range (0,len(excel_data_src.sheet_names())) :
        excel_sheet = excel_data_src.sheet_by_index(sheet)
        # print('[excel] parse sheet: %s (%d row, %d col)' % (excel_sheet.name, excel_sheet.nrows, excel_sheet.ncols))
        
        lua_str = lua_str + "\n\t" + excel_sheet.name + " = {"
        
        keys = []
        for col in range (0, excel_sheet.ncols):
            cell = excel_sheet.cell(0, col)
            keys.append(cell.value)

        for row in range(1, excel_sheet.nrows):
            cell = excel_sheet.cell(row, 0)
            lua_str = lua_str + "\n\t\t["+str(int(cell.value)) +"] = {\n"
            for col in range (0, excel_sheet.ncols):
                cell2 = excel_sheet.cell(row, col)
                cvalue = cell2.value
                if isinstance(cvalue,int) or isinstance(cvalue,float) :
                    cvalue = str(int(cvalue))
                else :
                    cvalue = "\"" + cvalue + "\""
                # print ("%s = %s" % (keys[col],cvalue))
                lua_str = lua_str + "\t\t\t"+keys[col] +" = " + cvalue + ",\n"
            lua_str = lua_str + "\n\t\t},"
        lua_str = lua_str + "\n\t},"

    lua_str = lua_str + "\n}\nreturn " + "data" + "\n"

    # 正则搜索lua文件名 不带后缀 用作table的名称 练习正则的使用
    searchObj = re.search(r'([^\\/:*?"<>|\r\n]+)\.\w+$', tgt_lua_path, re.M|re.I)
    lua_table_name = searchObj.group(1)
    # print('正则匹配:', lua_table_name, searchObj.group(), searchObj.groups())

    # 这个就直接获取文件名了
    src_excel_file_name = os.path.basename(src_excel_path)
    tgt_lua_file_name = os.path.basename(tgt_lua_path)

    # file head desc
    lua_file_head_desc = lua_file_head_format_desc % (tgt_lua_file_name, src_excel_file_name)

    # export to lua file
    lua_export_file = open(tgt_lua_path, 'w')
    lua_export_file.write(lua_file_head_desc)
    lua_export_file.write(lua_str)
    lua_export_file.close()
    # print('输出lua文件:'+tgt_lua_path)


# Make a script both importable and executable (∩_∩)
if __name__ == '__main__':
    if len(sys.argv) < 3:
        print('python excel2lua.py <excel_input_path> <lua_output_path>')
        exit(1)

    foreaExcelFile(os.path.join(curpath, sys.argv[1]), os.path.join(curpath, sys.argv[2]))

    exit(0)
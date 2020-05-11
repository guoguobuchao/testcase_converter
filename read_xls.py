import xlrd,xlwt
import sys
from write_xml import *
from docx import Document
import re

def convert_excel_xml(file_name,output_path):
# 打开文件
    #if len(sys.argv)==1:
    #    file_name='testcase.xls'
    #else:
    #    file_name = str(sys.argv[1])
    
    data = xlrd.open_workbook(file_name)
    files=[]
    for table in data.sheets():
        if len(data.sheet_names())==1:
            suite='testcases'
        else:
            suite=table.name

# 获取行数和列数
# 行数：table.nrows
# 列数：table.ncols
        index_row=[i.lower() for i in table.row_values(0)]
        
        RT_index = index_row.index('rt')
        name_index= index_row.index('name')
        summary_index = index_row.index('summary')
        eps_index = index_row.index('steps')
        results_index= index_row.index('expected results')
     
            
        RT = table.row_values(1)[0]
        testcase_lst=[]
        for row in range(1,table.nrows):
            if table.row_values(row)[name_index]:
                testcase_lst.append(table.row_values(row))
        if testcase_lst:
            file=createXML('%s/%s.xml'%(output_path,suite),suite,testcase_lst)
            files.append(file)
    return files


# coding:utf-8



def indexof(key,list):
    # print(list)
    index = []
    for i in range(len(list)):
        if key in list[i].lower():
            index.append(i)
    if len(index)==0:
        raise Exception('Index "{}" not found in {}'.format(key,list))
    return index

def slice_list_by(para,list):
    para_index = indexof(para,list)
    para_index.append(len(list))

    result =[]
    for i in range(1,len(para_index)):
        each_section = list[para_index[i-1]:para_index[i]]
        result.append(each_section)
    return result


def read_doc(filename):
    # 创建文档对象
    document = Document(filename)

    # 读取文档中所有的段落列表
    ps = document.paragraphs
    # for p in ps:
    #     print(p.text)
    # 每个段落有两个属性：style和text
    # ps_detail = [(x.text, x.style.name) for x in ps]
    testsuite = [p.text.strip().replace('\xa0','').replace("’","'") for p in ps if p.text.strip()]

    sections=slice_list_by('section::',testsuite)

    suite_list=[]

    for section in sections:
        section_dict={}

        #update section
        section_dict.update({'section':section[1]})

        #update summary
        if section[2].lower() == 'summary::' and '::' not in section[3]:
            summary =section[3]
        else:
            summary=''
        section_dict.update({'summary':summary})

        #update testcases
        testcase_list = []
        testcases=slice_list_by('testcasename::',section)

        split_pattern = r'testcasename::|summary::|executiontype::|importance::|steps::|expectedresults::'
        for testcase in testcases:

            case_str = '\n'.join(testcase)
            testcase =re.split(split_pattern,case_str,0,flags=re.IGNORECASE)
            
            if len(testcase)>5:
                testcase_list.append({
                    'name':testcase[1].strip('\n'),
                    'summary':testcase[2].strip('\n'),
                    'steps':testcase[5].strip('\n'),                         #replace  u'\xa0' as ''
                    'expected results':testcase[6].strip('\n')            #replace  u'\xa0' as ''
                })
                
        section_dict.update({'testcases': testcase_list})
        suite_list.append(section_dict)
    return suite_list
    #suite_list:
    # [
    #   {'section':section,'summary:':summary,
    #    'testcases':{'name':name,'summary':summary,'steps':steps,'expected results':results},
    #   {'section':section,'summary:':summary,
    #     'testcases':{'name':name,'summary':summary,'steps':steps,'expected results':results},
    #   {'section':section,'summary:':summary,
    #     'testcases':[{'name':name,'summary':summary,'steps':steps,'expected results':results},],
    # ]

def write_xls(suite_list,filename):
    workbook=xlwt.Workbook()
    for section in suite_list:
        print(section['section'])
        sheet=workbook.add_sheet(section['section'])
        sheet.write(0,0,'rt')
        sheet.write(0,1, 'name')
        sheet.write(0,2,'summary')
        sheet.write(0,3,'steps')
        sheet.write(0,4,'expected results')

        for i in range(len(section['testcases'])):
            sheet.write(i+1,1,section['testcases'][i]['name'])
            sheet.write(i+1,2,section['testcases'][i]['summary'])
            sheet.write(i+1,3,section['testcases'][i]['steps'])
            sheet.write(i+1,4,section['testcases'][i]['expected results'])

    workbook.save(filename)

def convert_doc(doc,outputpath):
    testcases=read_doc(doc)
    write_xls(testcases,outputpath+'/testcases.xls')
    files=convert_excel_xml(outputpath+'/testcases.xls',outputpath)
    return files

if __name__ == '__main__':
    convert_doc('C:/Users/auto/Desktop/testcase/tc_converter/RT-4835 Uplink configuration migration for CAP-RAP to C2C-IAP with staticIp-AP1x-pppoe-proxy-uplink vlan tag test plan.docx','C:/Users/auto/Desktop/testcase/tc_converter/dist')
    #convert_excel_xml('testcases.xls','C:/Users/auto/Desktop/testcase/tc_converter/dist')
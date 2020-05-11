# coding:utf-8


from docx import Document
import re

def indexof(key,list):
    # print(list)
    index = []
    for i in range(len(list)):
        if key == list[i].lower():
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
    document = Document('filename')

    # 读取文档中所有的段落列表
    ps = document.paragraphs
    # for p in ps:
    #     print(p.text)
    # 每个段落有两个属性：style和text
    # ps_detail = [(x.text, x.style.name) for x in ps]
    testsuite = [p.text.strip() for p in ps if p.text.strip()]

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
            testcase_list.append({
                'name':testcase[1].strip('\n'),
                'summary':testcase[2].strip('\n'),
                'steps':testcase[5].strip('\n'),
                'expected results':testcase[6].strip('\n')
            })

        section_dict.update({'testcases': testcase_list})
        suite_list.append(section_dict)
    return suite_list



if __name__ == '__main__':
    read_doc('demo2.docx')
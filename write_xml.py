# -*- coding: utf-8 -*-
from xml.dom import minidom

def createTestsuite(domTree,testsuite_name,RT=""):
    #craete testSuite_node
    testSuite_node=domTree.createElement("testsuite")
    testSuite_node.setAttribute("name",testsuite_name)
    
    #create node_order
    node_order = domTree.createElement("node_order")
    cdata_node_order = domTree.createCDATASection("0")
    node_order.appendChild(cdata_node_order)
    testSuite_node.appendChild(node_order)
    
    #create nodd_details
    node_details = domTree.createElement("details")
    cdata_details = domTree.createCDATASection(testsuite_name)
    node_details.appendChild(cdata_details)
    #custom_fields
    node_custom_fields=domTree.createElement("custom_fields")
    node_custom_field=domTree.createElement("custom_field")
    
    #name
    node_name = domTree.createElement('name')
    node_name.appendChild(domTree.createCDATASection("Requirement_Tracking_suit"))
    
    #value
    node_value = domTree.createElement('value')
    node_value.appendChild(domTree.createCDATASection(RT))
    #custom_field
    node_custom_field.appendChild(node_name)
    node_custom_field.appendChild(node_value)
    node_custom_fields.appendChild(node_custom_field)
    node_details.appendChild(node_custom_fields)
    testSuite_node.appendChild(node_details)
    
    return testSuite_node

def createTestcases(domTree):
    testcases_node=domTree.createElement("testcases")
    return testcases_node
    
def createTestcase(domTree,tc_List):
    
    tc_List:[RT,name,summary,steps,results]
    testcase_node=domTree.createElement('testcase')
    testcase_node.setAttribute("name",tc_List[1])
    #node_order
    node_order = domTree.createElement('node_order')
    node_order.appendChild(domTree.createCDATASection('0'))
    #externalid
    externalid_node = domTree.createElement('externalid')
    externalid_node.appendChild(domTree.createCDATASection('0'))
    #version
    version_node = domTree.createElement('version')
    version_node.appendChild(domTree.createCDATASection('1'))
    #summary tc_List[1]
    summary_node = domTree.createElement('summary')
    summary_node.appendChild(domTree.createCDATASection(tc_List[2]))
    #preconditions
    preconditions_node = domTree.createElement('preconditions')
    preconditions_node.appendChild(domTree.createCDATASection(''))
    #execution_type
    execution_type_node = domTree.createElement('execution_type')
    execution_type_node.appendChild(domTree.createCDATASection('1'))
    #importance
    importance_node = domTree.createElement('importance')
    importance_node.appendChild(domTree.createCDATASection('2'))
    
    testcase_node.appendChild(node_order)
    testcase_node.appendChild(externalid_node)
    testcase_node.appendChild(version_node)
    testcase_node.appendChild(summary_node)
    testcase_node.appendChild(preconditions_node)
    testcase_node.appendChild(execution_type_node)
    testcase_node.appendChild(importance_node)
    
    #steps
    if tc_List[3] or tc_List[4]:
        action_lst=tc_List[3].split('\n')
        actions=''.join('<p>%s</p>'%i for i in action_lst)
        expectedresult_lst=tc_List[4].split('\n')
        expectedresults=''.join('<p>%s</p>'%i for i in expectedresult_lst)

    
        steps_node = domTree.createElement('steps')
        step_node = domTree.createElement('step')
        
        #each_step:
        step_number_node = domTree.createElement('step_number')
        step_number_node.appendChild(domTree.createCDATASection('1'))
        #steps
        actions_node = domTree.createElement('actions')
        actions_node.appendChild(domTree.createCDATASection(actions))
        #expectedresults
        expectedresults_node = domTree.createElement('expectedresults')
        expectedresults_node.appendChild(domTree.createCDATASection(expectedresults))
        
        step_node.appendChild(step_number_node)
        step_node.appendChild(actions_node)
        step_node.appendChild(expectedresults_node)
        step_node.appendChild(execution_type_node)
        
        #steps append step
        steps_node.appendChild(step_node)
        #testcase append steps
        testcase_node.appendChild(steps_node)
        
    #custom_fields
    custom_fields_node = domTree.createElement('custom_fields')
    name=['Comments','Requirement_Tracking_case']
    value = ['',tc_List[0]]
    for i in range(2):
        custom_field_node = domTree.createElement('custom_field')
        name_node = domTree.createElement('name')
        name_node.appendChild(domTree.createCDATASection(name[i]))
        value_node = domTree.createElement('value')
        value_node.appendChild(domTree.createCDATASection(value[i]))
        custom_field_node.appendChild(name_node)
        custom_field_node.appendChild(value_node)
        custom_fields_node.appendChild(custom_field_node)
    testcase_node.appendChild(custom_fields_node)
    
    return testcase_node
    
def createXML(filename,suite_name,testcases_lst):
    domTree=minidom.Document()
    if suite_name == 'testcases':
        root_node=createTestcases(domTree)
    else:
        root_node=createTestsuite(domTree,suite_name)
        
    for testcase in testcases_lst:
        #print(testcase)
        testcase_node=createTestcase(domTree,testcase)
        root_node.appendChild(testcase_node)
    domTree.appendChild(root_node)
    
    try:
        with open(filename,'w') as f:
            domTree.writexml(f,indent='',addindent='\t',newl='\n',encoding='utf-8')
    except UnicodeEncodeError as e:
        raise Exception('some characters encoding error in file %s'%filename)
    else:    
        return filename.split('/')[-1]
if __name__ =='__main__':
    createXML('testcases.xml','123',[['RT1','case1','s1','',''],['RT1','case2','s2','','']])
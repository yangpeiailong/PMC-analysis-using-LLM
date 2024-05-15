import os
import xlrd
import xlwt
from tqdm import tqdm
import re
os.environ["DASHSCOPE_API_KEY"] = "your API key"

from langchain_community.llms import Tongyi
from langchain.prompts import PromptTemplate
from langchain.chains import LLMChain
from langchain_community.document_loaders import UnstructuredFileLoader
from langchain.text_splitter import RecursiveCharacterTextSplitter
from functools import reduce
import json

template = """Question: {question}

Answer: 按照要求回答这个问题"""

prompt = PromptTemplate(
    template=template,
    input_variables=["question"])

# print(prompt)

llm = Tongyi()
llm.model_name = 'qwen-plus'

llm_chain = LLMChain(prompt=prompt, llm=llm)

wb = xlrd.open_workbook("D:/成都理工大学重要文件夹/Text Analysis and Evaluation of China's financial inclusion Policy based on text mining/results/extraction_results_by_Tongyi.xls")
sheet = wb.sheet_by_index(0)
rows = sheet.nrows

# question = "我给通义千问传入的prompt最大长度是可以是多少？"
#
# res = llm_chain.invoke(question)
#
# print(res)
# 当前目录
base_dir = "D:/成都理工大学重要文件夹/Text Analysis and Evaluation of China's financial inclusion Policy based on text mining/datasets/"
# 获取当前目录下的所有文件
files = [os.path.join(base_dir, file) for file in os.listdir(base_dir)]
policy_nature = []
policy_area = []
policy_timeliness = []
policy_tool = []
release_agency_list = []
implementation_agency_list = []
function_list = []
measures_list = []
coverage_list = []
# 遍历文件列表，输出文件名

book = xlwt.Workbook(encoding='utf-8')
sheet2 = book.add_sheet('sheet1', cell_overwrite_ok=True)
sheet2.write(0, 0, '政策文件名')
sheet2.write(0, 1, '政策类型')
sheet2.write(0, 2, '政策范围')
sheet2.write(0, 3, '政策执行期限')
sheet2.write(0, 4, '政策工具')
sheet2.write(0, 5, '政策发布机构')
sheet2.write(0, 6, '政策执行机构')
sheet2.write(0, 7, '政策功能')
sheet2.write(0, 8, '政策措施')
sheet2.write(0, 9, '政策覆盖')
row_w = 1
for file in tqdm(files):
    sheet2.write(row_w, 0, file)
    loader = UnstructuredFileLoader(file)
    row = 0
    for i in range(1, rows):
        if sheet.cell(i,0).value == file:
            row = i
    # print(sheet.cell(row,1).value)
    release_agency = json.loads(sheet.cell(row,1).value)
    implementation_agency = json.loads(sheet.cell(row,2).value)
    # print(sheet.cell(row,3).value)
    functions = json.loads(sheet.cell(row,3).value)
    measures = json.loads(sheet.cell(row, 4).value)
    coverage = json.loads(sheet.cell(row, 5).value)

    documents = loader.load()
    start = documents[0].page_content[:500]
    # ==========1.判断政策类型===========
    prompt1 = ('对于给定的政策文本的开头，帮我判断政策的类型，类型包含：第一类，立法；第二类，条例；第三类，计划或规划；第四类：意见；第五类：通知；第六类：决定或决策。'
               '回复的格式是字典格式，即{立法:1或0, 条例:1或0, 计划或规划:1或0, 意见:1或0, 通知:1或0, 决定或决策:1或0}，'
               '这里的1表示属于这一类，0表示不属于这一类，一个政策通常只属于一个类别，即只有一个1，其他都为0。'
               '不要改变字典的键，只回复该字典，不要回复其他内容。文本如下：')
    prompt1 += start
    res1 = llm_chain.invoke(prompt1)
    print(res1['text'])
    policy_nature.append(res1['text'])
    sheet2.write(row_w, 1, res1['text'])

    #
    # ==========3.判断政策执行期限===========
    prompt3 = ('对于给定的政策文本的开头，帮我判断政策的执行期限，包含以下类型：'
               '第一类，长期，即5年以上；第二类，中期，即3到5年；第三类，短期，即1到3年；第四类：超短期，即小于一年。'
               '回复的格式是字典格式，即{长期:1或0, 中期:1或0, 短期:1或0, 超短期:1或0}，'
               '这里的1表示属于该类型，0表示不属于该类型，对一个政策的期限而言，只会是其中之一，因而只会有一个类型的取值为1，其他类型都为0。'
               '不要改变字典的键，只回复该字典，不要回复其他内容。文本如下：')
    prompt3 += start
    res3 = llm_chain.invoke(prompt3)
    print(res3['text'])
    policy_timeliness.append(res3['text'])
    sheet2.write(row_w, 3, res3['text'])
    #


    # ==========2和4.提取政策范围和工具类型===========
    text_spliter = RecursiveCharacterTextSplitter(chunk_size=1000, chunk_overlap=10)
    split_docs = text_spliter.split_documents(documents)
    policy_area_temp = []
    policy_tool_temp = []
    implementation_agency_temp = ''
    extraction_temp = ''
    for doc in split_docs:
        # ==========2.提取政策范围===========
        prompt2 = ('对于给定的政策文本，帮我判断政策的涉及的方面，方面包含：第一，经济；第二，社会；第三，政治；第四：技术。'
                   '回复的格式是字典格式，即{"经济":1或0, "社会":1或0, "政治":1或0, "技术":1或0}，'
                   '这里的1表示涉及该方面，0表示不涉及，对一个政策而言，有可能涉及多个方面，则可以有多个方面取值为1。'
                   '不要改变字典的键，只回复该字典，不要回复其他内容，字典请在一行中输出，不要加入换行符。文本如下：')
        prompt2 += doc.page_content
        res2 = llm_chain.invoke(prompt2)
        print(res2['text'])
        policy_area_temp.append(json.loads(res2['text']))

        # ==========4.提取政策工具类型===========
        prompt4 = ('对于给定的政策文本，帮我判断政策的工具类型，包含以下类型：'
                   '第一类，供给型政策工具，即指政府在人才培养、资金支持、技术支持、公共服务等方面直接投入资源，推动特定领域或行业的发展；'
                   '第二类，需求型政策工具，指政府通过政府采购、贸易政策、用户补贴、应用示范和价格指导等方式，减少市场的不确定性，培育并扩大特定市场，从需求侧拉动产业的发展；'
                   '第三类，环境型政策工具，指政府通过目标规划、金融支持、法规规范、标准管理、税收优惠等方式，为特定领域或行业的发展提供有利的政策环境、金融环境和法律环境，间接促进其发展。'
                   '回复的格式是字典格式，即{"供给":1或0, "需求":1或0, "环境":1或0}，'
                   '这里的1表示属于该类型，0表示不属于该类型，对一个政策的工具类型而言，有可能包含涉及多个方面，则可以有多个方面取值为1。'
                   '不要改变字典的键，只回复该字典，不要回复其他内容，字典请在一行中输出，不要加入换行符。文本如下：')
        prompt4 += doc.page_content
        res4 = llm_chain.invoke(prompt4)
        print(res4['text'])
        policy_tool_temp.append(json.loads(res4['text']))

    print('===========================================')
    # ==========2.提取政策范围汇总===========
    result2 = policy_area_temp[0]
    for key in policy_area_temp[0].keys():
        result2[key] = 1 if sum([policy_area_temp[index][key] for index in range(len(policy_area_temp))]) > 0 else 0
    print(result2)
    policy_area.append(result2)
    sheet2.write(row_w, 2, str(result2))
    # ==========4.提取政策工具类型汇总===========
    result4 = policy_tool_temp[0]
    for key in policy_tool_temp[0].keys():
        result4[key] = 1 if sum([policy_tool_temp[index][key] for index in range(len(policy_tool_temp))]) > 0 else 0
    print(result4)
    policy_tool.append(result4)
    sheet2.write(row_w, 4, str(result4))
    print('===========================================')
    # ==========5.判断政策发布机构类型===========
    prompt5 = ('对于给定的政策发布机构的列表，帮我判断这些机构的类型，类型包含：第一类，中央政府部门；第二类，省级人民政府；第三类，省级财政部门；第四类：地方政府办公厅；第五类：金融监管机构；第六类：自治区人民政府；第七类，直辖市财政局；第八类，省政府办公厅。'
               '回复的格式是字典格式，即{中央政府部门:1或0, 省级人民政府:1或0, 省级财政部门:1或0, 地方政府办公厅:1或0, 金融监管机构:1或0, 自治区人民政府:1或0, 直辖市财政局:1或0, 省政府办公厅:1或0}，'
               '这里的1表示列表中存在元素属于这一类，0表示列表中不存在元素属于这一类，列表中一个元素只属于一个类别且一定属于其中一个类，即不存在字典所有键对应的值都为0的情况。'
               '除非提供的列表有多个元素，才有可能存在多个1的情况，若只有一个元素，则只有一个1，其他都为0。'
               '不要改变字典的键，只回复该字典，不要回复其他内容，不要在回复中添加额外的换行符。文本如下：')
    prompt5 += str(release_agency)
    res5 = llm_chain.invoke(prompt5)
    res5 = re.findall(r'{.+}',res5['text'].replace('\n', '').replace(' ', ''))[0]
    print(res5)
    release_agency_list.append(res5)
    sheet2.write(row_w, 5, res5)
#
    # ==========6.判断政策执行机构类型===========
    prompt6 = (
        '对于给定的政策执行机构的列表，帮我判断这些机构的类型，类型包含：第一类，政府部门；第二类，金融部门；第三类，监管机构；第四类：地方行政；第五类：企业与金融机构；第六类：社会团体与教育。'
        '回复的格式是字典格式，即{政府部门:1或0, 金融部门:1或0, 监管机构:1或0, 地方行政:1或0, 企业与金融机构:1或0, 社会团体与教育:1或0}，'
        '这里的1表示列表中存在元素属于这一类，0表示列表中不存在元素属于这一类，列表中一个元素只属于一个类别且一定属于其中一个类，即不存在字典所有键对应的值都为0的情况。'
        '不要改变字典的键，只回复该字典，不要回复其他内容，不要在回复中添加额外的换行符。文本如下：')
    prompt6 += str(implementation_agency)
    res6 = llm_chain.invoke(prompt6)
    res6 = re.findall(r'{.+}', res6['text'].replace('\n', '').replace(' ', ''))[0]
    print(res6)
    implementation_agency_list.append(res6)
    sheet2.write(row_w, 6, res6)

    # ==========7.判断政策功能类型===========
    prompt7 = (
        '对于给定的政策功能的列表，帮我判断这些功能的类型，类型包含：第一类，金融服务与支持；第二类，普惠金融与乡村振兴；第三类，风险管理与监管；第四类：绿色发展与可持续性；第五类：科技创新与信息化；第六类：社会服务与公平保障。'
        '回复的格式是字典格式，即{金融服务与支持:1或0, 普惠金融与乡村振兴:1或0, 风险管理与监管:1或0, 绿色发展与可持续性:1或0, 科技创新与信息化:1或0, 社会服务与公平保障:1或0}，'
        '这里的1表示列表中存在元素属于这一类，0表示列表中不存在元素属于这一类，列表中一个元素只属于一个类别且一定属于其中一个类，即不存在字典所有键对应的值都为0的情况。'
        '不要改变字典的键，只回复该字典，不要回复其他内容，不要在回复中添加额外的换行符。文本如下：')
    prompt7 += str(functions)
    res7 = llm_chain.invoke(prompt7)
    res7 = re.findall(r'{.+}', res7['text'].replace('\n', '').replace(' ', ''))[0]
    print(res7)
    function_list.append(res7)
    sheet2.write(row_w, 7, res7)


    # ==========8.判断政策措施类型===========
    prompt8 = (
        '对于给定的政策措施的列表，帮我判断这些措施的类型，类型包含：第一类，金融服务与创新；第二类，风险管理与监管；第三类，普惠与农村金融；第四类：保险与保障；第五类：财政与税收支持；第六类：金融教育与合作。'
        '回复的格式是字典格式，即{金融服务与创新:1或0, 风险管理与监管:1或0, 普惠与农村金融:1或0, 保险与保障:1或0, 财政与税收支持:1或0, 金融知识与教育:1或0}，'
        '这里的1表示列表中存在元素属于这一类，0表示列表中不存在元素属于这一类，列表中一个元素只属于一个类别且一定属于其中一个类，即不存在字典所有键对应的值都为0的情况。'
        '不要改变字典的键，只回复该字典，不要回复其他内容，不要在回复中添加额外的换行符。文本如下：')
    prompt8 += str(measures)
    res8 = llm_chain.invoke(prompt8)
    res8 = re.findall(r'{.+}', res8['text'].replace('\n', '').replace(' ', ''))[0]
    print(res8)
    measures_list.append(res8)
    sheet2.write(row_w, 8, res8)

    # ==========9.判断政策覆盖对象===========
    prompt9 = (
        '对于给定的政策覆盖对象的列表，帮我判断这些对象的类型，类型包含：第一类，企业类型；第二类，农村与农业；第三类，特殊群体与弱势群体；第四类：金融机构；第五类：政策与项目；第六类：城乡发展。'
        '回复的格式是字典格式，即{企业类型:1或0, 农村与农业:1或0, 特殊群体与弱势群体:1或0, 金融机构:1或0, 政策与项目:1或0, 城乡发展:1或0}，'
        '这里的1表示列表中存在元素属于这一类，0表示列表中不存在元素属于这一类，列表中一个元素只属于一个类别且一定属于其中一个类，即不存在字典所有键对应的值都为0的情况。'
        '不要改变字典的键，只回复该字典，不要回复其他内容，不要在回复中添加额外的换行符。文本如下：')
    prompt9 += str(coverage)
    res9 = llm_chain.invoke(prompt9)
    res9 = re.findall(r'{.+}', res9['text'].replace('\n', '').replace(' ', ''))[0]
    print(res9)
    coverage_list.append(res9)
    sheet2.write(row_w, 9, res9)
    row_w += 1
book.save("D:/成都理工大学重要文件夹/Text Analysis and Evaluation of China's financial inclusion Policy based on text mining/results/sub_variables_scores_results_by_Tongyi.xls")

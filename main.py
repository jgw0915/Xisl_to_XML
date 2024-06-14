import openpyxl
from openpyxl import Workbook
from openpyxl import load_workbook
import xml.etree.ElementTree as ET
from enum import Enum

class TaskType(Enum):
    Desc = 0
    Form = 1
    BusinessRule = 2

class TaskListItems:
    def __init__(self,itemName_EN:str,taskType:TaskType,description:str,level:int,object_name:str) -> None:
        self.itemName_EN = itemName_EN
        self.taskType = taskType
        self.description = description
        self.level = level
        self.object_name = object_name
    

class TreeNode:
    def __init__(self, data,level):
        self.data = data
        self.children = []
        self.children_map = []
        self.level = level

    def add_child(self, child_node):
        if child_node.data not in self.children_map:
            self.children.append(child_node)
            self.children_map.append(child_node.data)

class Tree:
    def __init__(self, root_data):
        self.root = TreeNode(root_data,0)

    def traverse_depth_first(self, start_node=None):
        if start_node is None:
            start_node = self.root
        yield (start_node)
        for child in start_node.children:
            yield from self.traverse_depth_first(child)
    
    def find_node(self, data, start_node=None):
        if start_node is None:
            start_node = self.root
        if start_node.data == data:
            return start_node
        for child in start_node.children:
            node = self.find_node(data, child)
            if node:
                return node
        return None 

TaskList_tags = ['TaskList','TaskItem','dataForm','Instructions','busRule','cubeName','ruleType']
wb = load_workbook(r'C:\Users\kchiang\Python Project\Small_Application\data\IEC工作清單設計_20240419.xlsx')
ws_list = [sheet for sheet in wb]
my_job_tasklist = ['SPDG Allocation','SPDG Reviewer','SPDG Reviewer','Revenue Reviewer(BPM)','Revenue Planner(Cost team)','Revenue Planner(Sales)']
for sheet in ws_list:
    if sheet in my_job_tasklist:
        root = ET.Element(tag=TaskList_tags[0],attrib={'xmlns:xsi':"http://www.w3.org/2001/XMLSchema-instance",'name':sheet.title,'folder':'Task Lists'})
        TaskItem_root = ET.SubElement(parent=root,tag=TaskList_tags[1],attrib={'name':sheet.title,'type':'descriptive','pos':str(0.0),'Dependency':'No'})
        '''Read Sheet data'''
        for row in range(3,len(sheet['A'])+1):
            itemName_EN = row[7]
            taskType = row[2]
            description = row[3]
            object_name = row[4]
            level = row[5]
            taskListIten = TaskListItems(itemName_EN,taskType,description,description,level,object_name)
            
            
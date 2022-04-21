
import win32com, sys, string, win32api, traceback
import win32com.client.dynamic
from win32com.test.util import CheckClean
#import pythoncom
#from win32com.client import gencache
#from pywintypes import Unicode


def dumpElement(indent, theElement):
    """
    Рекурсивный обход дерева элементов и вывод значений атрибутов

    :param indent: уровень вложенности элемента для форматирования вывода
    :param theElement: текущий элемент, с которого начинаем вывод
    :return: ничего
    """
    currentElement = theElement

    element = {
        'Type': currentElement.Type,
        'Name': currentElement.name,
        'Stereotype': currentElement.stereotype,
        'Alias': currentElement.alias,
        'Description': currentElement.Notes,
        'AttrType/Status': currentElement.status,
        'AttrLength/Difficulty': currentElement.difficulty,
        'Default/Priority': currentElement.priority,
        'Mandatory': '',
        'Attributes': [],
        'Methods': [],
        'Parameters': []
    }
    # print(' '*indent, currentElement.Type, currentElement.name, currentElement.stereotype, currentElement.alias, currentElement.Notes, currentElement.status, currentElement.difficulty, currentElement.priority, "")
    print(' '*indent, element)

    for attr in currentElement.Attributes:
        obj = {
            'Type': 'Attribute',
            'Name': attr.name,
            'Stereotype': attr.stereotype,
            'Alias': attr.alias,
            'Description': attr.Notes,
            'AttrType/Status': attr.Type,
            'AttrLength/Difficulty': attr.Length,
            'Default/Priority': attr.Default,
            'Mandatory': attr.Container
        }
        element['Attributes'].append(obj)
        # print(' '*(indent + 1), "Attribute", attr.name, attr.stereotype, attr.alias, attr.Notes, attr.Type, attr.Length, attr.Default, attr.Container)
        print(' '*(indent + 1), obj)

    for mthd in currentElement.Methods:
        obj = {
            'Type': 'Method',
            'Name': mthd.name,
            'Stereotype': mthd.stereotype,
            'Alias': mthd.alias,
            'Description': mthd.Notes,
            'AttrType/Status': mthd.returnType,
            'AttrLength/Difficulty': '',
            'Default/Priority': '',
            'Mandatory': ''
        }
        element['Methods'].append(obj)
        # print(' '*(indent + 1), "Method", mthd.name, mthd.stereotype, mthd.alias, mthd.Notes, mthd.returnType, "", "", "")
        print(' '*(indent + 1), obj)

        for prm in mthd.Parameters:
            obj = {
                'Type': 'Parameter',
                'Name': prm.name,
                'Stereotype': prm.stereotype,
                'Alias': prm.alias,
                'Description': prm.Notes,
                'AttrType/Status': prm.Type,
                'AttrLength/Difficulty': '',
                'Default/Priority': prm.Default,
                'Mandatory': ''
            }
            element['Parameters'].append(obj)
            # print(' '*(indent + 2), "Parameter", prm.name, prm.stereotype, prm.alias, prm.Notes, prm.Type, "", prm.Default, "")
            print(' '*(indent + 2), obj)

    import json

    # data = json.load(element)
    print(json.dumps(element, indent=2, ensure_ascii=False))
    # print(f'>>> {element}')

    for elem in currentElement.elements:
        dumpElement(indent + 1, elem)       # Рекурсивно выведем все.


def dumpPackage(indent, thePackage):
    """
    Рекурсивный вывод содержимого пакета. Сначала выводятся элементы самого пакета, а потом обход вложенных пакетов

    :param indent: уровень вложенности элемента для форматирования вывода
    :param thePackage: текущий пакет, с которого начинается вывод
    :return: ничего
    """
    currentPackage = thePackage
    obj = {
        'Type': 'Package',
        'Name': currentPackage.name,
        'Stereotype': currentPackage.StereotypeEx,
        'Alias': currentPackage.alias,
        'Description': currentPackage.Notes,
        'AttrType/Status': '',
        'AttrLength/Difficulty': '',
        'Default/Priority': '',
        'Mandatory': ''
    }
    # print(' '*indent, "Package", currentPackage.name, currentPackage.StereotypeEx, currentPackage.alias, currentPackage.Notes, "", "", "", "")
    print(' '*indent, obj)

    for currentElement in currentPackage.elements:
        dumpElement(indent + 1, currentElement)

    for childPackage in currentPackage.Packages:
        dumpPackage(indent + 1, childPackage)


#####################
try:
    print('Connect to the EA')
    eapp = win32com.client.dynamic.Dispatch("EA.App")
    print(f'EA: {eapp}')
except:
    print('Please Open Enterprise Architect')


try:
    print('Connect to the Repository')
    currentRep = eapp.Repository
    print(f'Repository: {currentRep}')
except:
    print('Please Open a Model in Enterprise Architect')

try:
    print('Define selected package')
    pkg = currentRep.GetTreeSelectedPackage()
    print(f'Current Package: {pkg} / {pkg.name}')
except:
    print('Package not selected')

print('=' * 40)
dicItemName = {
    2: 'Repository',
    4: 'Element',
    5: 'Package',
    7: 'Connector',
    8: 'Diagram',
    23: 'Attribute'
}

try:
    itemType = currentRep.GetContextItemType()
    print(f'Selected Item Type: {dicItemName.get(itemType)}')
except:
    print('Nothing selected')

print('=' * 40)
itemType, item = currentRep.GetTreeSelectedItem()
if itemType == 5:       # Package
    print(f'===== Package: {item.name} =====')
    dumpPackage(0,item)
elif itemType == 4:     # Element
    print(f'Itemtype: {itemType}, Item Name: {item.name}, Alias: {item.alias}')
    dumpElement(0,item)         # Выводим данные элемента и все вложенные в него составляющие
elif itemType == 8:     # Diagram
    print(f'Diagram: {item.name}')
    for dObj in item.DiagramObjects:        # Выборка всех элементов диаграммы
        dElem = currentRep.GetElementByID(dObj.ElementID)
        print(f'Элемент диаграммы: ID-{dObj.ElementID}  Название- {dElem.name}  Тип- {dElem.Type}')
    print('=' * 20)
    for dLnk in item.DiagramLinks:          # Выборка всех коннекторов на диаграмме
        dConn = currentRep.GetConnectorByID(dLnk.ConnectorID)
        print(f'Коннектор: ID-{dLnk.ConnectorID}  Название- {dConn.name}  Тип- {dConn.Type}')

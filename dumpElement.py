def dumpElement(indent, theElement):

    currentElement = theElement

    print(' '*indent, currentElement.Type, currentElement.name, currentElement.stereotype, currentElement.alias, currentElement.Notes, currentElement.status, currentElement.difficulty, currentElement.priority, "")

    for attr in currentElement.Attributes:
        print(' '*(indent + 1), "Attribute", attr.name, attr.stereotype, attr.alias, attr.Notes, attr.Type, attr.Length, attr.Default, attr.Container)

    for mthd in currentElement.Methods:
        print(' '*(indent + 1), "Method", mthd.name, mthd.stereotype, mthd.alias, mthd.Notes, mthd.returnType, "", "", "")
        for prm in mthd.Parameters:
            print(' '*(indent + 2), "Parameter", prm.name, prm.stereotype, prm.alias, prm.Notes, prm.Type, "", prm.Default, "")

    for elem in currentElement.elements:
        DumpElement(indent + 1, elem)

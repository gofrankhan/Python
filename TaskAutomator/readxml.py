from xml.dom import minidom

# parse an xml file by name
mydoc = minidom.parse('D:\MIT-DU\S4\Project\exercise\TaskAutomator\\test.xml')

items = mydoc.getElementsByTagName('Action')

# all item attributes
for elem in items:
    actionName = elem.attributes['type'].value
    print(actionName)
    if(actionName == 'openurl'):
        url = elem.getElementsByTagName('url')
        browser = elem.getElementsByTagName('browser')
        print(url[0].firstChild.data)
        print(browser[0].firstChild.data)
    elif(actionName == 'click'):
        xpath = elem.getElementsByTagName('xpath')
        print(xpath[0].firstChild.data)
    elif(actionName == 'readtext'):
        xpath = elem.getElementsByTagName('xpath')
        store = elem.getElementsByTagName('Store')
        print(xpath[0].firstChild.data)
        print(store[0].firstChild.data)
    elif(actionName == 'messagebox'):
        text = elem.getElementsByTagName('Text')
        print(text[0].firstChild.data)


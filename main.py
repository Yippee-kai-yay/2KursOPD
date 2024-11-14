import io
import zipfile
import xml.etree.ElementTree as ET

textTag = '{urn:oasis:names:tc:opendocument:xmlns:text:1.0}'
styleTag ='{urn:oasis:names:tc:opendocument:xmlns:style:1.0}'
foTag = '{urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0}'
drawTag = '{urn:oasis:names:tc:opendocument:xmlns:drawing:1.0}'
officeTag = '{urn:oasis:names:tc:opendocument:xmlns:office:1.0}'
dcTag = '{http://purl.org/dc/elements/1.1/}'

defaultFont= {}
rootStyles = ET.Element("Root")
def findDefaultFont(rootStyles, name=None):
    global defaultFont

    for elem in rootStyles.iter(styleTag + 'default-style'):
        if elem.get(styleTag+ 'family')=='paragraph':
            for subelem in elem.iter():
                if subelem.tag == styleTag + 'text-properties':
                    defaultFont['name'] = subelem.get(styleTag + 'font-name')
                    defaultFont['size'] = subelem.get(foTag + 'font-size')
                    defaultFont['weight'] = subelem.get(foTag + 'font-weight') if not None else 'normal'
                    defaultFont['background'] = subelem.get(foTag + 'background-color')
                    defaultFont['underline'] = subelem.get(styleTag + 'text-underline-type')
                    defaultFont['color'] = subelem.get(foTag + 'color')
                if subelem.tag == styleTag + 'paragraph-properties':
                    defaultFont['indent'] = subelem.get(foTag + 'text-indent')
                    defaultFont['height'] = subelem.get(foTag + 'line-height')
                    defaultFont['mrgTop'] = subelem.get(foTag + 'margin-top')
                    defaultFont['mrgBottom'] = subelem.get(foTag + 'margin-bottom')
                    defaultFont['align'] = subelem.get(foTag + 'text-align') if not None else 'start'
    for elem in rootStyles.iter(styleTag + 'style'):
        if name is not None and elem.get(styleTag+ 'name')==name:
            for subelem in elem.iter():
                if subelem.tag == styleTag + 'text-properties':
                    defaultFont['name'] = subelem.get(styleTag + 'font-name') if not None else defaultFont['name']
                    defaultFont['size'] = subelem.get(foTag + 'font-size') if not None else defaultFont['size']
                    defaultFont['weight'] = subelem.get(foTag + 'font-weight') if not None else defaultFont['weight']
                    defaultFont['background'] = subelem.get(foTag + 'background-color') if not None else defaultFont['background']
                    defaultFont['underline'] = subelem.get(styleTag + 'text-underline-type') if not None else defaultFont['underline']
                    defaultFont['color'] = subelem.get(foTag + 'color') if not None else defaultFont['color']
                if subelem.tag == styleTag + 'paragraph-properties':
                    defaultFont['indent'] = subelem.get(foTag + 'text-indent') if not None else defaultFont['indent']
                    defaultFont['height'] = subelem.get(foTag + 'line-height') if not None else defaultFont['height']
                    defaultFont['mrgTop'] = subelem.get(foTag + 'margin-top') if not None else defaultFont['mrgTop']
                    defaultFont['mrgBottom'] = subelem.get(foTag + 'margin-bottom') if not None else defaultFont['mrgBottom']
                    defaultFont['align'] = subelem.get(foTag + 'text-align') if not None else defaultFont['align']

def showText(root):
    for elem in root.iter():
        if elem.tag == textTag +'p':
            if len(elem)>0:
                for subelem in elem.iter(textTag +'span'):
                    if len(subelem) == 0:
                        print(f'   Text: "{subelem.text}"')

                    else:
                        for Selem in subelem.iter(officeTag + 'annotation'):
                            for  JUSTFUCKINGWORKPLEASE in Selem.iter():
                                if len(JUSTFUCKINGWORKPLEASE) == 0:
                                    text_content = JUSTFUCKINGWORKPLEASE.text
                                    print(f'        Text: "{text_content}"')
            else:
                style_name = elem.get(textTag+'style-name')
                text_content = ''
                text_content = text_content.join(node.text if node.text else '' for node in elem.iter())
                print(f'Text: "{text_content}" - Style Name: {style_name} - Font: {styleProp(style_name, root)}')

        elif elem.tag == textTag + 's':
            print(f'Special Element: {elem.text.strip() if elem.text else ""}' )

        # elif elem.tag == textTag +'span':
        #    span_style_name = elem.get('{urn:oasis:names:tc:opendocument:xmlns:text:1.0}style-name')
        #    print(f'Span Element: "{elem.text.strip() if elem.text else ""}" - Style Name: {span_style_name}')

def check_all_attributes(element):
    for attr, value in element.attrib.items():
        print(f"Attribute: {attr}, Value: {value}")

def styleProp(name, root, findDefault = True):
    for style in root.iter(styleTag + 'style'  ):
        if style.get(styleTag + 'name') == name:
            if findDefault:
                if style.get(styleTag+'parent-style-name') is not None:
                    findDefaultFont(rootStyles, style.get(styleTag+'parent-style-name'))

                else:
                    findDefaultFont(rootStyles)

            font = dict(defaultFont)
            #print(name, style.get(styleTag + 'parent-style-name'), defaultFont)
            for subelem in style:
                if subelem.tag == styleTag+ 'text-properties':
                    if subelem.get(styleTag + 'font-name') is not None:
                        font['name'] = subelem.get(styleTag + 'font-name')
                    if subelem.get(foTag + 'font-size') is not None:
                        font['size'] = subelem.get(foTag + 'font-size')
                    if subelem.get(foTag + 'font-weight') is not None:
                        font['weight'] = subelem.get(foTag + 'font-weight')
                    if subelem.get(foTag + 'background-color') is not None:
                        font['background'] = subelem.get(foTag + 'background-color')
                    if subelem.get(styleTag + 'text-underline-type') is not None:
                        font['underline'] = subelem.get(styleTag + 'text-underline-type')
                    if subelem.get(foTag + 'color') is not None:
                        font['color'] = subelem.get(foTag + 'color')
                if subelem.tag == styleTag + 'paragraph-properties':
                    if subelem.get(foTag + 'text-indent') is not None:
                        font['indent'] = subelem.get(foTag + 'text-indent')
                    if subelem.get(foTag + 'line-height') is not None:
                        font['height'] = subelem.get(foTag + 'line-height')
                    if subelem.get(foTag + 'margin-top') is not None:
                        font['mrgTop'] = subelem.get(foTag + 'margin-top')
                    if subelem.get(foTag + 'margin-bottom') is not None:
                        font['mrgBottom'] = subelem.get(foTag + 'margin-bottom')
                    if subelem.get(foTag + 'text-align') is not None:
                        font['align'] = subelem.get(foTag + 'text-align')
            return font




def find_differences(line1, line2):
    differences = None
    for i in range(max(len(line1), len(line2))):
        char1 = line1[i] if i < len(line1) else None
        char2 = line2[i] if i < len(line2) else None
        if char1 != char2:
            if differences is None:
                differences = []
            differences.append((i, char1, char2))
    return differences

def markMistake(elem, reason):
    span = ET.SubElement(elem, textTag + 'span')
    office = ET.SubElement(span, officeTag + 'annotation')
    ET.SubElement(office, dcTag + 'creator').text = 'Norm_check'
    ET.SubElement(office, dcTag + 'date').text = '2000-01-01T12:00:00'
    ET.SubElement(office, textTag + 'p', {textTag + 'style-name': "Текстпримечания"}).text = reason

def fontMark(requirements, elem, props):
    reqFont = requirements['Font'] if not None else None
    reqSize = requirements['Size'] + 'pt' if requirements['Size'] is not None else None
    reqWgt = requirements['Weight'] if not None else None
    reqNoBack = True if requirements['Allow background'] != 'Yes' else None
    reqNoUnd = True if requirements['Allow underline'] != 'Yes' else None

    if reqFont is not None and props['name'] is not None and find_differences(props['name'], reqFont):
        markMistake(elem, 'Неверный шрифт: ' + props['name'])
        #print(elem.get(textTag + 'style-name'), props['name'], elem.text, props)
    if reqSize is not None and props['size'] is not None and find_differences(props['size'], reqSize):
        markMistake(elem, 'Неверный размер шрифта: ' + props['size'])
    if reqWgt is not None and props['weight'] is not None and find_differences(props['weight'], reqWgt):
        markMistake(elem, 'Курсив/полужирный')
    if reqNoBack and props['background'] != 'transparent' and props['background'] != '#ffffff' and props['background'] != None :
        markMistake(elem, 'Выделение цветом')
    if reqNoUnd and props['underline'] != 'none' and props['underline'] != None:
        markMistake(elem, 'Выделение подчёркиванием')



def styleMark(requirements, root):
    reqInd = requirements['Indent'] if requirements['Indent'] is not None else None
    reqAlign = requirements['Align'] if not None else None
    reqHgt = str(round(float(requirements['Height'])*100))+'%' if requirements['Height'] is not None else None
    for elem in root.iter(textTag + 'p'):
        if elem.get(textTag + 'style-name') == 'Текстпримечания':
            continue
        props = styleProp(elem.get(textTag + 'style-name'), root)
        if props is None:
            findDefaultFont(rootStyles)
            props = defaultFont
        text = "1234567"
        if elem.text is not None:
            text = elem.text
        if text[0:7] != "Рисунок":
            if reqAlign is not None and props['align'] is not None and find_differences(props['align'], reqAlign):
                markMistake(elem, 'Неверное выравнивание: ' + props['align'])
            if reqInd is not None and props['indent'] is not None:
                text = props['indent']
                if(text[-2:]=='in'):
                    ind = float(text[0:-2])
                    ind = ind*2.5391
                else:
                    ind = float(text[0:-2])
                ind = round(ind, 2)
                #print(ind)
                if find_differences(str(ind), reqInd):
                    markMistake(elem, 'Неверный отступ')
            if reqHgt is not None and props['height'] is not None and find_differences(props['height'], reqHgt):
                markMistake(elem, 'Неверный междустрочный интервал')

        if len(elem)>0:
            for subelem in elem.iter(textTag+'span'):
                if subelem.text is not None:
                    if subelem.get(textTag + 'parent-style-name') is None:
                        defaultFont = styleProp(elem.get(textTag + 'style-name'), root)
                        props = styleProp(subelem.get(textTag + 'style-name'), root, False)
                    else:
                        props = styleProp(subelem.get(textTag + 'style-name'), root)
                    if props is not None:
                        fontMark(requirements, subelem, props)
        else:
            if elem.text is not None:
                props = styleProp(elem.get(textTag + 'style-name'), root)
                if props is not None:
                    #print()
                    fontMark(requirements, elem, props)

def checkPinturas(root):
    links = []
    pics = []
    for elem in root.iter(textTag+'p'):
        if len(elem)>0:
            text=''
            for subelem in elem.iter():
                text = text + str(subelem.text)+' '
        else:
            text=elem.text

        if text is not None and text.find('Рисунок ')!=-1:
            words = text.split(' ')
            num = words[1]
            #print(words)
            pics = pics + [num]
            if num not in links:
                markMistake(elem,'Отсутствует ссылка на риcyнoк')
        if text is not None and (text.find('рисунке ') != -1 or text.find('(рисунок ') != -1 or text.find('(рисунки ') != -1 or text.find('рисунках ')!= -1):
            words = text.split(' ')
            #print(words)
            if text.find('рисунке ') != -1:
                ind = words.index('рисунке')
                num = words[ind+1]
                if len(num)>3:
                    num = num[:-1]
            if text.find('(рисунок ') != -1:
                ind = words.index('(рисунок')
                num = words[ind + 1]
                if len(num) > 3:
                    num = num[:-2]
            if text.find('(рисунки ') != -1 or text.find('рисунках ')!= -1:
                if  '(рисунки' in words: ind = words.index('(рисунки')
                else: ind = words.index('рисунках')
                first, left = words[ind+1].split('.')
                if len(left)>3:
                    left = left[:-1]
                first, right = words[ind+3].split('.')
                if len(right)>3:
                    right = right[:-1]
                #print(left,right)
                num = -1
                for i in range(int(left),int(right)):
                    links = links + [first+'.'+str(i)]


            #print(num , words)
            links = links + [num]
        #print(pics, links)

def numberMark(requirements, root): #doesnt work
    reqNoFirst = True if requirements['Numbering_First'] != 'Yes' else None
    reqSize = requirements['Size'] + 'pt' if requirements['Size'] is not None else None
    reqAlign = requirements['Align'] if not None else None

    for elem in root.iter(officeTag + 'master-styles'):
        for subelem in elem.iter(styleTag + 'master-page'):
            for SSelem in subelem.iter(styleTag+'footer'):
                props = styleProp(SSelem.get(textTag+'p'),root)
                if props is not None:
                    if reqSize is not None and find_differences(reqSize, props['size']):
                        markMistake(SSelem,'Непраильный размер шрифта')
#^doesnt work^

def headerFontCheck(elem, level, requirements, props):
    reqFont = requirements['Font'] if not None else None
    reqNoBack = True if requirements['Allow background'] != 'Yes' else None
    reqNoUnd = True if requirements['Allow underline'] != 'Yes' else None
    reqSize = requirements['Heading' + level + '_Size'] + 'pt' if requirements['Heading1_Size'] is not None else None
    reqWgt = requirements['Heading' + level + '_Weight'] if not None else None


    if reqFont is not None and props['name'] is not None and find_differences(props['name'], reqFont):
        markMistake(elem, 'Неверный шрифт заголовка: ' + props['name'])
    if reqSize is not None and props['size'] is not None and find_differences(props['size'], reqSize):
        markMistake(elem, 'Неверный размер шрифта заголовка: ' + props['size'])
    if reqWgt is not None and props['weight'] is not None and find_differences(props['weight'], reqWgt):
        markMistake(elem, 'Неверное выделение заголовка: ' + props['weight'])
    if reqNoBack and props['background'] != 'transparent' and props['background'] != None:
        markMistake(elem, 'Выделение заголовка цветом')
    if reqNoUnd and props['underline'] != 'none' and props['underline'] != None:
        markMistake(elem, 'Выделение заголовка подчёркиванием')

def headerMark(requirements, root):

    lv1Num = 0
    lv2Num = 0
    lv3Num = 0
    lv4Num = 0
    for elem in root.iter(textTag + 'h'):

        level = elem.get(textTag+'outline-level')
        reqAlign = requirements['Heading' + level + '_Align'] if not None else None

        if len(elem) > 0:
            text = ''
            for subelem in elem.iter(textTag + 'span'):
                #print(subelem.text)
                if subelem.text is not None:
                    #print('before',text, str(subelem.text))
                    text = text+ str(subelem.text)
                    #print('after',text)
        else:
            text = elem.text
        #print(text)

        if text is not None and text != '':
            if len(elem) > 0:
                for subelem in elem.iter(textTag + 'span'):
                    if subelem.text is not None:
                        findDefaultFont(rootStyles,subelem.get(textTag + 'parent-style-name') )
                        global defaultFont
                        defaultFont = styleProp(elem.get(textTag + 'style-name'), root)
                        #print(elem.get(textTag + 'style-name'), defaultFont)
                        props = styleProp(subelem.get(textTag + 'style-name'), root,False)
                        #print(subelem.get(textTag + 'style-name'),props)
                        headerFontCheck(subelem,level,requirements, props)
            else:
                props = styleProp(elem.get(textTag + 'style-name'), root)
                if reqAlign is not None and props['align'] is not None and find_differences(props['align'], reqAlign):
                    markMistake(elem, 'Неверное выравнивание заголовка: ' + props['align'])
                headerFontCheck(elem, level, requirements, props)
            if level == '1':
                lv1Num+=1
                lv2Num = 0


                if find_differences(text[0], str(lv1Num)):
                    markMistake(elem, 'Неверная нумерация уровня 1')
            if level == '2':
                lv2Num+=1
                lv3Num = 0
                #print(text[2], str(lv2Num),find_differences(text[2], str(lv2Num)))
                if find_differences(text[2], str(lv2Num)):
                    markMistake(elem, 'Неверная нумерация уровня 2')
            if level == '3':
                lv2Num+=1
                lv4Num = 0
                if find_differences(text[4], str(lv3Num)):
                    markMistake(elem, 'Неверная нумерация уровня 3')
            if level == '4':
                lv2Num+=1
                if find_differences(text[6], str(lv4Num)):
                    markMistake(elem, 'Неверная нумерация уровня 4')





def textToDict(text):
    dict = {}
    for line in text.splitlines():
        key, value = line.split(':')
        dict[key] = value
    return dict

odf_file = 'test.odt'
req_file = 'reqs.txt'
print('!')
content = zipfile.ZipFile(odf_file).open('content.xml') .read()
styles = zipfile.ZipFile(odf_file).open('styles.xml') .read()
root = ET.fromstring(content)
rootStyles = ET.fromstring(styles)



reqs = textToDict(open(req_file).read())

findDefaultFont(rootStyles)
checkPinturas(root)
styleMark(reqs, root)
headerMark(reqs,root)


outputXml = ET.tostring(root, encoding='utf-8', xml_declaration=True)
with zipfile.ZipFile(odf_file, 'r') as zip_file:
    with zipfile.ZipFile('result.odt', 'w') as temp_zip:
        for item in zip_file.namelist():
            if item != 'content.xml':
                temp_zip.writestr(item, zip_file.read(item))
        temp_zip.writestr('content.xml', outputXml)
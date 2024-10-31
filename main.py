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

def findDefaultFont(rootStyles):
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
                    defaultFont['align'] = subelem.get(foTag + 'text-align')

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

def styleProp(name, root):
    font = dict(defaultFont)
    for style in root.iter(styleTag + 'style'  ):
        if style.get(styleTag + 'name') == name:
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
    if reqSize is not None and props['size'] is not None and find_differences(props['size'], reqSize):
        markMistake(elem, 'Неверный размер шрифта: ' + props['size'])
    if reqWgt is not None and props['weight'] is not None and find_differences(props['weight'], reqWgt):
        markMistake(elem, 'Курсив/полужирный')
    if reqNoBack and props['background'] != 'transparent':
        markMistake(elem, 'Выделение цветом')
    if reqNoUnd and props['underline'] != 'none':
        markMistake(elem, 'Выделение подчёркиванием')



def styleMark(requirements, root):
    reqInd = str(round(float(requirements['Indent'])/2.5391, 4))+'in' if requirements['Indent'] is not None else None
    reqAlign = requirements['Align'] if not None else None
    reqHgt = str(round(float(requirements['Height'])*100))+'%' if requirements['Height'] is not None else None
    for elem in root.iter(textTag + 'p'):
        if elem.get(textTag + 'style-name') == 'Текстпримечания':
            continue
        props = styleProp(elem.get(textTag + 'style-name'), root)
        if reqAlign is not None and props['align'] is not None and find_differences(props['align'], reqAlign):
            markMistake(elem, 'Неверное выравнивание: ' + props['align'])
        if reqInd is not None and props['indent'] is not None and find_differences(props['indent'], reqInd):
            markMistake(elem, 'Неверный отступ')
        if reqHgt is not None and props['height'] is not None and find_differences(props['height'], reqHgt):
            markMistake(elem, 'Неверный междустрочный интервал')

        if len(elem)>0:
            for subelem in elem.iter(textTag+'span'):
                props = styleProp(subelem.get(textTag + 'style-name'), root)
                if props is not None and subelem.text is not None:
                    fontMark(requirements, subelem, props)
        else:
            props = styleProp(elem.get(textTag + 'style-name'), root)
            if props is not None and elem.text is not None:
                fontMark(requirements, elem, props)



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

def headerMark(requirements, root):
    reqFont = requirements['Font'] if not None else None
    reqNoBack = True if requirements['Allow background'] != 'Yes' else None
    reqNoUnd = True if requirements['Allow underline'] != 'Yes' else None


    for elem in root.iter(textTag + 'h'):
        props = styleProp(elem.get(textTag+'style-name'),root)
        level = elem.get(textTag+'outline-level')

        reqSize = requirements['Heading'+level+'_Size'] + 'pt' if requirements['Heading1_Size'] is not None else None
        reqWgt = requirements['Heading'+level+'_Weight'] if not None else None
        reqAlign = requirements['Heading'+level+'_Align'] if not None else None


        if reqFont is not None and props['name'] is not None and find_differences(props['name'], reqFont):
            markMistake(elem, 'Неверный шрифт: ' + props['name'])
        if reqSize is not None and props['size'] is not None and find_differences(props['size'], reqSize):
            markMistake(elem, 'Неверный размер шрифта: ' + props['size'])
        if reqWgt is not None and props['weight'] is not None and find_differences(props['weight'], reqWgt):
            markMistake(elem, 'Неверное выделение: '+ props['weight'])
        if reqNoBack and props['background'] != 'transparent':
            markMistake(elem, 'Выделение цветом')
        if reqNoUnd and props['underline'] != 'none':
            markMistake(elem, 'Выделение подчёркиванием')
        if reqAlign is not None and props['align'] is not None and find_differences(props['align'], reqAlign):
            markMistake(elem, 'Неверное выравнивание: ' + props['align'])



def textToDict(text):
    dict = {}
    for line in text.splitlines():
        key, value = line.split(':')
        dict[key] = value
    return dict

odf_file = 'test.odt'
req_file = 'reqs.txt'

content = zipfile.ZipFile(odf_file).open('content.xml') .read()
styles = zipfile.ZipFile(odf_file).open('styles.xml') .read()
root = ET.fromstring(content)
rootStyles = ET.fromstring(styles)

reqs = textToDict(open(req_file).read())

findDefaultFont(rootStyles)
styleMark(reqs, root)
headerMark(reqs,root)

outputXml = ET.tostring(root, encoding='utf-8', xml_declaration=True)
with zipfile.ZipFile(odf_file, 'r') as zip_file:
    with zipfile.ZipFile('result.odt', 'w') as temp_zip:
        for item in zip_file.namelist():
            if item != 'content.xml':
                temp_zip.writestr(item, zip_file.read(item))
        temp_zip.writestr('content.xml', outputXml)
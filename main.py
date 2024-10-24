import io
import zipfile
import xml.etree.ElementTree as ET

textTag = '{urn:oasis:names:tc:opendocument:xmlns:text:1.0}'
styleTag ='{urn:oasis:names:tc:opendocument:xmlns:style:1.0}'
foTag = '{urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0}'
drawTag = '{urn:oasis:names:tc:opendocument:xmlns:drawing:1.0}'
officeTag = '{urn:oasis:names:tc:opendocument:xmlns:office:1.0}'
dcTag = '{http://purl.org/dc/elements/1.1/}'

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
    print(name)
    font = {'name':'Calibri', 'size':'11pt', 'weight':'normal', 'indent':'0.3937in', 'height':'100%', 'background':None, 'underline':None}
    for style in root.iter(styleTag + 'style'  ):
        if style.get(styleTag + 'name') == name:
            for subelem in style:
                if subelem.tag == styleTag+ 'text-properties':
                    font['name'] = subelem.get(styleTag + 'font-name')
                    font['size'] = subelem.get(styleTag + 'font-size-asian')
                    font['weight'] = subelem.get(styleTag + 'font-weight-asian')
                    font['background'] = subelem.get(foTag + 'background-color')
                    font['underline'] = subelem.get(styleTag + 'text-underline-type')
                if subelem.tag == styleTag + 'paragraph-properties':
                    font['indent'] = subelem.get(foTag + 'text-indent')
                    font['height'] = subelem.get(foTag + 'line-height')
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
    #ET.dump(elem)

def styleMark(requirements, root):
    reqFont = requirements['Font'] if not None else None
    reqSize = requirements['Size']+'pt' if requirements['Size'] is not None else None
    reqInd = str(round(float(requirements['Indent'])/2.5391, 4))+'in' if requirements['Indent'] is not None else None
    reqHgt = str(round(float(requirements['Height'])*100))+'%' if requirements['Height'] is not None else None
    reqWgt = requirements['Weight'] if not None else None
    reqBack = True if requirements['Allow background'] != 'Yes' else None
    reqUnd = True if requirements['Allow underline'] != 'Yes' else None
    for elem in root.iter(textTag + 'p'):
        if len(elem)>0:
            for subelem in elem.iter():
                props = styleProp(subelem.get(textTag + 'style-name'), root)
                #print(props)
                if props is not None:
                    if reqFont is not None and props['name'] is not None and find_differences(props['name'], reqFont):
                        markMistake(elem, 'Неверный шрифт: '+props['name'])
                    if reqSize is not None and props['size'] is not None and find_differences(props['size'], reqSize):
                        markMistake(elem, 'Неверный размер шрифта: '+props['size'])
                    if reqWgt is not None and props['weight'] is not None and find_differences(props['weight'], reqWgt):
                        markMistake(elem, 'Курсив/полужирный')
                    if reqBack and props['background'] is not None:
                        markMistake(elem, 'Выделение цветом')
                    if reqUnd and props['underline'] is not None:
                        markMistake(elem, 'Выделение подчёркиванием')
        else:
            props = styleProp(elem.get(textTag+'style-name'),root)
            #print(props)
            if props is not None:
                if reqFont is not None and props['name'] is not None and find_differences(props['name'],reqFont):
                    markMistake(elem,'Неверный шрифт: '+props['name'])
                if reqSize is not None and props['size'] is not None and find_differences(props['size'],reqSize):
                    markMistake(elem,'Неверный размер шрифта: '+props['size'])
                if reqWgt is not None and props['weight'] is not None and find_differences(props['weight'], reqWgt):
                    markMistake(elem, 'Курсив/полужирный')
                if reqBack and props['background'] is not None:
                    markMistake(elem, 'Выделение цветом')
                if reqUnd and props['underline'] is not None:
                    markMistake(elem, 'Выделение подчёркиванием')

                if reqInd is not None and props['indent'] is not None and find_differences(props['indent'], reqInd):
                    markMistake(elem, 'Неверный отступ')
                    #print('Неверный отступ', props['indent'], reqInd)
                if reqHgt is not None and props['height'] is not None and find_differences(props['height'], reqHgt):
                    markMistake(elem, 'Неверный междустрочный интервал')
                    #print('Неверный междустрочный интервал', props['height'], reqHgt)




def textToDict(text):
    dict = {}
    for line in text.splitlines():
        key, value = line.split(':')
        dict[key] = value
    return dict

odf_file = 'kurs.odt'
req_file = 'reqs.txt'

content = zipfile.ZipFile(odf_file).open('content.xml') .read()
styles = zipfile.ZipFile(odf_file).open('styles.xml') .read()
root = ET.fromstring(content)
reqs = textToDict(open(req_file).read())

styleMark(reqs, root)

outputXml = ET.tostring(root, encoding='utf-8', xml_declaration=True)
with zipfile.ZipFile(odf_file, 'r') as zip_file:
    with zipfile.ZipFile('result.odt', 'w') as temp_zip:
        for item in zip_file.namelist():
            if item != 'content.xml':
                temp_zip.writestr(item, zip_file.read(item))
        temp_zip.writestr('content.xml', outputXml)
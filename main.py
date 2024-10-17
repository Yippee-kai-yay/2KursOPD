import io
import zipfile
import xml.etree.ElementTree as ET

textTag = '{urn:oasis:names:tc:opendocument:xmlns:text:1.0}'
styleTag ='{urn:oasis:names:tc:opendocument:xmlns:style:1.0}'
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

def styleProp(name, root):
    for style in root.iter():
        if style.tag == (styleTag + 'style') and style.get(styleTag + 'name') == name:
            for subelem in style:
                if subelem.tag == styleTag+ 'text-properties':
                    font = {'name': subelem.get(styleTag + 'font-name'),
                            'size': subelem.get(styleTag + 'font-size-asian'),
                            'weigth': subelem.get(styleTag + 'font-weight-asian')}
                    return font




def find_differences(line1, line2):
    """Prints characters that are different between two lines."""
    differences = None
    for i in range(max(len(line1), len(line2))):
        char1 = line1[i] if i < len(line1) else None
        char2 = line2[i] if i < len(line2) else None
        if char1 != char2:
            if differences is None:
                differences = []
            differences.append((i, char1, char2))
    return differences

def textMark(requirements, root):
    reqFont = requirements['Font'] if not None else None

    for elem in root.iter(textTag + 'p'):
        if len(elem)>0:
            for subelem in elem.iter(textTag + 'span'):
                props = styleProp(subelem.get(textTag + 'style-name'), root)
                if props is not None and find_differences(props['name'], reqFont):
                    span = ET.SubElement(subelem, textTag+'span')
                    office = ET.SubElement(span, officeTag + 'annotation')
                    ET.SubElement(office, dcTag + 'creator').text = 'me'
                    ET.SubElement(office, dcTag + 'date').text = '2024-10-09T11:25:00'
                    ET.SubElement(office, textTag + 'p', {textTag+'style-name': "Текстпримечания"} ).text = 'Текст комментария'
                    ET.dump(subelem)

        else:
            props = styleProp(elem.get(textTag+'style-name'),root)
            if props is not None and find_differences(props['name'],reqFont):
                span = ET.SubElement(elem, textTag + 'span')
                office = ET.SubElement(span, officeTag + 'annotation')
                ET.SubElement(office, dcTag + 'creator').text = 'me'
                ET.SubElement(office, dcTag + 'date').text = '2024-10-09T11:25:00'
                ET.SubElement(office, textTag + 'p',
                              {textTag + 'style-name': "Текстпримечания"}).text = 'Текст комментария'
                ET.dump(elem)





def checkStyles(root):
    for style in root.iter():
        if style.tag == styleTag + 'style':
            style_name = style.get(styleTag+ 'name')
            print(f'Style: {style_name}' )
            for subelem in style:
                print(subelem.tag)
                #for SUS_elem in subelem:
                #    print(SUS_elem.tag)
                if subelem.tag == (styleTag+ 'text-properties'):
                    font_name = subelem.get(styleTag+ 'font-name')
                    font_size = subelem.get(styleTag+ 'font-size-asian')
                    font_weigth = subelem.get(styleTag + 'font-weight-asian')
                    print(f'Font Name: {font_name}, Font Size: {font_size}')

def textToDict(text):
    dict = {}
    for line in text.splitlines():
        key, value = line.split(':')
        dict[key] = value
    return dict

odf_file = 'test — копия.odt'
req_file = 'reqs.txt'

content = zipfile.ZipFile(odf_file).open('content.xml') .read()
styles = zipfile.ZipFile(odf_file).open('styles.xml') .read()
root = ET.fromstring(content)
reqs = textToDict(open(req_file).read())

textMark(reqs, root)

outputXml = ET.tostring(root, encoding='utf-8', xml_declaration=True)
with zipfile.ZipFile(odf_file, 'r') as zip_file:
    with zipfile.ZipFile('test_output.odt', 'w') as temp_zip:
        for item in zip_file.namelist():
            if item != 'content.xml':
                # Copy all files except the one to delete
                temp_zip.writestr(item, zip_file.read(item))
        temp_zip.writestr('content.xml', outputXml)
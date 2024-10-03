import zipfile
import xml.etree.ElementTree as ET

textTag = '{urn:oasis:names:tc:opendocument:xmlns:text:1.0}'
styleTag ='{urn:oasis:names:tc:opendocument:xmlns:style:1.0}'
drawTag = '{urn:oasis:names:tc:opendocument:xmlns:drawing:1.0}'

def showText(root):


    for elem in root.iter():
        if elem.tag == textTag +'p':
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
                    font_name = subelem.get(styleTag + 'font-name')
                    font_size = subelem.get(styleTag+ 'font-size-asian')
                    font_weigth = subelem.get(styleTag + 'font-weight-asian')
                    return font_name, font_size,font_weigth

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



odf_file = 'LABA2.odt'
content = zipfile.ZipFile(odf_file).open('content.xml') .read()
root = ET.fromstring(content)
showText(root)



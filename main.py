import docx
import xml.etree.ElementTree as ET


class Normocontrol:

    def __init__(self, document_path):
        self.document_path = document_path
        self.doc = docx.Document(self.document_path)
        self.numbering_properties = {}
        self.property = {}
        self.property_constructor()

    def __str__(self):
        return str(self.document_path)

    def property_constructor(self):

        def get_image_id_in_paragraph(par):
            ids = []
            root = ET.fromstring(par._p.xml)
            namespace = {
                'a': "http://schemas.openxmlformats.org/drawingml/2006/main",
                'r': "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
                'wp': "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"}

            inlines = root.findall('.//wp:inline', namespace)
            for inline in inlines:
                imgs = inline.findall('.//a:blip', namespace)
                for img in imgs:
                    id = img.attrib['{{{0}}}embed'.format(namespace['r'])]
                    ids.append(id)

            return ids

        def get_lists_properties(_numbering):
            """
            Получение параметров 'Абстарктных стилей' списков
            :param _numbering: nembering.xml извлеченный из документа
            :return: numId_list (dict) - словарь параметров списков по numId
            вида lists_settings[abstractNumId][ilvl][key][pPr_key]
            """
            root = ET.fromstring(_numbering.xml)
            namespace = _numbering.nsmap

            numId_list = {}
            nums = root.findall('.//w:num', namespace)
            for num in nums:
                numId = num.attrib['{{{0}}}numId'.format(namespace['w'])]
                abstractNums = num.findall('.//w:abstractNumId', namespace)
                for abstractNum in abstractNums:
                    abstractNumId = abstractNum.attrib['{{{0}}}val'.format(namespace['w'])]
                    numId_list[numId] = abstractNumId

            abstract_num_settings = {}
            abstractNums = root.findall('.//w:abstractNum', namespace)
            for abstractNum in abstractNums:
                abstractNumId = abstractNum.attrib['{{{0}}}abstractNumId'.format(namespace['w'])]
                abstract_num_settings[abstractNumId] = {}
                lvls = abstractNum.findall('.//w:lvl', namespace)
                for lvl in lvls:
                    ilvl = lvl.attrib['{{{0}}}ilvl'.format(namespace['w'])]
                    abstract_num_settings[abstractNumId][ilvl] = {}
                    properties = lvl.getchildren()
                    for prop in properties:
                        key = prop.tag.split('}')[1]
                        try:
                            value = list(prop.attrib.values())[0]
                            abstract_num_settings[abstractNumId][ilvl][key] = value
                        except IndexError:
                            abstract_num_settings[abstractNumId][ilvl][key] = {}
                            pPrs = prop.getchildren()
                            for pPr in pPrs:
                                pPr_attribs = pPr.attrib
                                pPr_list = {}
                                for pPr_attrib in pPr_attribs:
                                    pPr_key = pPr_attrib.split('}')[1]
                                    vpPr_alue = pPr_attribs[pPr_attrib]
                                    pPr_list[pPr_key] = vpPr_alue
                                abstract_num_settings[abstractNumId][ilvl][key]['pPr'] = pPr_list

            for numId in numId_list:
                numId_list[numId] = abstract_num_settings[numId_list[numId]]

            return numId_list

        _numbering = self.doc._part.numbering_part.numbering_definitions._numbering
        self.numbering_properties = get_lists_properties(_numbering)

        document_body_property_object = self.doc._body._element.sectPr
        self.property['document_body_property'] = {
            "top": document_body_property_object.top_margin.cm,
            "bottom": document_body_property_object.bottom_margin.cm,
            "left": document_body_property_object.left_margin.cm,
            "right": document_body_property_object.right_margin.cm
        }

        self.property['paragraphs'] = []
        for index, paragraph in enumerate(self.doc.paragraphs):
            try:
                paragraph_style_name = paragraph.style.name

                text = paragraph.text

                # Параметры шрифта
                font_name = paragraph.style.font.name
                if font_name is None:
                    font_name = paragraph.style.base_style.font.name

                font_size = paragraph.style.font.size
                if font_size is None:
                    font_size = paragraph.style.base_style.font.size
                font_size = font_size.pt

                font_is_bold = paragraph.style.font.bold
                if font_is_bold is None and len(paragraph.runs) == 1 and paragraph.runs[0].bold is not None:
                    font_is_bold = paragraph.runs[0].bold

                alignments = {
                    None: "LEFT",
                    0: "LEFT",
                    1: "CENTER",
                    2: "RIGHT",
                    3: "JUSTIFY",
                    4: "DISTRIBUTE",
                    5: "JUSTIFY_MED",
                    7: "JUSTIFY_HI",
                    8: "JUSTIFY_LOW",
                    9: "THAI_JUSTIFY"
                }
                alignment = alignments[paragraph.alignment]

                # Параметры абзаца
                left_indent = paragraph.paragraph_format.left_indent
                if left_indent is None:
                    left_indent = paragraph.style.paragraph_format.left_indent
                if left_indent is not None:
                    left_indent = left_indent.cm

                first_line_indent = paragraph.paragraph_format.first_line_indent
                if first_line_indent is None:
                    first_line_indent = paragraph.style.paragraph_format.first_line_indent
                if first_line_indent is not None:
                    first_line_indent = first_line_indent.cm

                # Интервал
                line_spacing = paragraph.paragraph_format.line_spacing
                if line_spacing is None:
                    line_spacing = paragraph.style.paragraph_format.line_spacing
                if line_spacing is None:
                    line_spacing = paragraph.style.base_style.paragraph_format.line_spacing

                line_spacing_rule = paragraph.paragraph_format.line_spacing_rule
                if line_spacing_rule is None:
                    line_spacing_rule = paragraph.style.paragraph_format.line_spacing_rule
                if line_spacing_rule is None:
                    line_spacing_rule = paragraph.style.base_style.paragraph_format.line_spacing_rule
                if line_spacing_rule is not None:
                    line_spacing_rules = {
                        None: None,
                        0: "SINGLE",
                        1: "ONE_POINT_FIVE",
                        2: "DOUBLE",
                        3: "AT_LEAST",
                        4: "EXACTLY",
                        5: "MULTIPLE"
                    }
                    line_spacing_rule = line_spacing_rules[line_spacing_rule]

                # Изображения
                image_id_list = get_image_id_in_paragraph(paragraph)
                images = []
                for image_id in image_id_list:
                    images.append(self.doc.part.related_parts[image_id])

                print(images)

                # Списки
                num_property_path = paragraph._p.pPr.numPr
                if num_property_path is not None:
                    numId = str(num_property_path.numId.val)
                    ilvl = str(num_property_path.ilvl.val)
                    list_property = self.numbering_properties[numId][ilvl]
                    list_property['numId'] = numId
                    list_property['ilvl'] = ilvl

                else:
                    list_property = None

                paragraph_property = {'index': index,
                                      'paragraph_style_name': paragraph_style_name,
                                      'font_name': font_name,
                                      'font_size': font_size,
                                      'is_bold': font_is_bold,
                                      'alignment': alignment,
                                      'text': text,
                                      'list_property': list_property,
                                      'paragraph_format': {
                                          'left_indent': left_indent,
                                          'first_line_indent': first_line_indent,
                                          'line_spacing': line_spacing,
                                          'line_spacing_rule': line_spacing_rule,
                                      }}
                print(paragraph_property)
                self.property['paragraphs'].append(paragraph_property)

            except Exception as err:
                print(err)


nc = Normocontrol("document.docx")
print(nc.property)
# for style in nc.doc.styles.element.style_lst:
#     print(style.name_val)

import docx
import xml.etree.ElementTree as ET
from docx.enum.style import WD_STYLE_TYPE
from docx.text.font import Font

class DOCX:

    def __init__(self, document_path):
        self.document_path = document_path
        self.doc = docx.Document(self.document_path)
        self.styles = self.get_styles_from_docx()
        self.numbering_properties = self.get_lists_properties()
        self.document_property = {}
        self.property_constructor()

    def __str__(self):
        return str(self.document_path)

    def get_styles_from_docx(self):
        styles = self.doc.styles
        paragraph_styles = {}
        for style in styles:
            paragraph_styles[style.style_id] = style

        return paragraph_styles

    def get_lists_properties(self):
        """
        Получение параметров 'Абстарктных стилей' списков
        :return: numId_list (dict) - словарь параметров списков по numId
        вида lists_settings[abstractNumId][ilvl][key][pPr_key]
        """
        _numbering = self.doc._part.numbering_part.numbering_definitions._numbering
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
                properties = list(lvl)
                for prop in properties:
                    key = prop.tag.split('}')[1]
                    try:
                        value = list(prop.attrib.values())[0]
                        abstract_num_settings[abstractNumId][ilvl][key] = value
                    except IndexError:
                        abstract_num_settings[abstractNumId][ilvl][key] = {}
                        pPrs = list(prop)
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

    def property_constructor(self):

        def get_image_id_in_paragraph(paragraph):
            """

            :type paragraph: object
            """
            ids = []
            root = ET.fromstring(paragraph._p.xml)
            namespace = {
                'a': "http://schemas.openxmlformats.org/drawingml/2006/main",
                'r': "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
                'wp': "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"}

            inlines = root.findall('.//wp:inline', namespace)
            for inline in inlines:
                images_list = inline.findall('.//a:blip', namespace)
                for img in images_list:
                    img_id = img.attrib['{{{0}}}embed'.format(namespace['r'])]
                    ids.append(img_id)

            return ids

        styles_element = self.doc.styles.element
        rPrDefault = styles_element.xpath('w:docDefaults/w:rPrDefault/w:rPr')[0]
        default_font_name = rPrDefault.rFonts_ascii

        document_body_property_object = self.doc._body._element.sectPr
        self.document_property['document_body_property'] = {
                "top": document_body_property_object.top_margin.mm,
                "bottom": document_body_property_object.bottom_margin.mm,
                "left": document_body_property_object.left_margin.mm,
                "right": document_body_property_object.right_margin.mm
            }

        self.document_property['paragraphs'] = []
        for index, paragraph in enumerate(self.doc.paragraphs):
            try:
                paragraph_style_id = paragraph.style.style_id
                paragraph_style_name = paragraph.style.name
                paragraph_style_type = paragraph.style.type
                paragraph_base_style = paragraph.style.base_style
                if paragraph_base_style is not None:
                    paragraph_base_style_name = paragraph_base_style.name
                    paragraph_base_style_type = paragraph_base_style.type
                else:
                    paragraph_base_style_name = None
                    paragraph_base_style_type = None

                text = paragraph.text

                # Параметры шрифта
                style = paragraph.style
                base_style = style.base_style

                if style.font is not None:
                    font_name = style.font.name
                    if font_name is None:
                        if base_style is not None:
                            base_style_font = base_style.font
                            if base_style_font is not None:
                                font_name = base_style_font.name
                    if font_name is None:
                        font_name = default_font_name
                else:
                    font_name = None

                font_size = style.font.size
                if font_size is None and base_style is not None:
                    font_size = base_style.font.size
                if font_size is not None:
                    font_size = font_size.pt

                font_is_bold = style.font.bold
                if font_is_bold is None and len(paragraph.runs) == 1 and paragraph.runs[0].bold is not None:
                    font_is_bold = paragraph.runs[0].bold

                all_caps = style.font.all_caps

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
                    left_indent = style.paragraph_format.left_indent
                if left_indent is not None:
                    left_indent = left_indent.mm
                else:
                    left_indent = 0

                right_indent = paragraph.paragraph_format.right_indent
                if right_indent is None:
                    right_indent = style.paragraph_format.right_indent
                if right_indent is not None:
                    right_indent = right_indent.mm
                else:
                    right_indent = 0

                first_line_indent = paragraph.paragraph_format.first_line_indent
                if first_line_indent is None:
                    first_line_indent = style.paragraph_format.first_line_indent
                if first_line_indent is not None:
                    first_line_indent = first_line_indent.mm
                else:
                    first_line_indent = 0

                # Интервал
                space_after = paragraph.paragraph_format.space_after
                if space_after is None:
                    space_after = style.paragraph_format.space_after
                if space_after is None and base_style is not None:
                    space_after = base_style.paragraph_format.space_after
                else:
                    space_after = 0

                space_before = paragraph.paragraph_format.space_before
                if space_before is None:
                    space_before = style.paragraph_format.space_before
                if space_before is None and base_style is not None:
                    space_before = base_style.paragraph_format.space_before
                else:
                    space_before = 0

                line_spacing = paragraph.paragraph_format.line_spacing
                if line_spacing is None:
                    line_spacing = style.paragraph_format.line_spacing
                if line_spacing is None and base_style is not None:
                    line_spacing = base_style.paragraph_format.line_spacing
                else:
                    line_spacing = 0

                line_spacing_rule = paragraph.paragraph_format.line_spacing_rule
                if line_spacing_rule is None:
                    line_spacing_rule = paragraph.style.paragraph_format.line_spacing_rule
                if line_spacing_rule is None and base_style is not None:
                    line_spacing_rule = base_style.paragraph_format.line_spacing_rule
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

                # Списки
                paragraph_proprty_path = paragraph._p.pPr
                if paragraph_proprty_path is not None and paragraph_proprty_path.numPr is not None:
                    numId = str(paragraph_proprty_path.numPr.numId.val)
                    ilvl = str(paragraph_proprty_path.numPr.ilvl.val)
                    list_property = self.numbering_properties[numId][ilvl]
                    list_property['numId'] = numId
                    list_property['ilvl'] = ilvl
                else:
                    list_property = None

                paragraph_property = {'index': index,
                                      'text': text,
                                      'style': {
                                          'paragraph_style_id': paragraph_style_id,
                                          'paragraph_style_name': paragraph_style_name,
                                          'paragraph_style_type': paragraph_style_type,
                                          'paragraph_base_style_name': paragraph_base_style_name,
                                          'paragraph_base_style_type': paragraph_base_style_type
                                      },
                                      'text_property': {
                                          'font_name': font_name,
                                          'font_size': font_size,
                                          'is_bold': font_is_bold,
                                          'all_caps': all_caps,
                                      },
                                      'list_property': list_property,
                                      'paragraph_format': {
                                          'alignment': alignment,
                                          'left_indent': left_indent,
                                          'right_indent': right_indent,
                                          'first_line_indent': first_line_indent,
                                          'space_after': space_after,
                                          'space_before': space_before,
                                          'line_spacing': line_spacing,
                                          'line_spacing_rule': line_spacing_rule,
                                      },
                                      'images': images}
                # print(paragraph_property)
                self.document_property['paragraphs'].append(paragraph_property)

            except Exception as err:
                print(err)


if __name__ == "__main__":
    nc = DOCX("document.docx")
    print(nc.document_property)

import json
from DOCX import DOCX


class Normocontrol(DOCX):
    def __init__(self, document_path, path_to_default_property_file):
        super().__init__(document_path)
        self.path_to_default_property_file = path_to_default_property_file
        self.default_property = self.get_default_properties()
        self.ERR_LIST = []

    def get_default_properties(self):
        with open(self.path_to_default_property_file, encoding="utf-8") as default_property_file:
            default_property = json.loads(default_property_file.read())
        return default_property

    def run(self):

        for section in self.default_property:

            if section == "document_body_property":
                body_properties = self.document_property[section]
                default_body_properties = self.default_property[section]
                for name in default_body_properties:
                    default_body_properties = self.default_property[section][name]['properties']

                    for default_body_property in default_body_properties:

                        if int(body_properties[default_body_property]) == default_body_properties[default_body_property]:
                            print('{}:{} = {}'.format(default_body_property, body_properties[default_body_property],
                                                      default_body_properties[default_body_property]))
                        else:
                            err_str = default_body_property['error_message'].format(
                                '{}:{} != {}'.format(default_body_property, body_properties[default_body_property],
                                                     default_body_properties[default_body_property]))
                            self.ERR_LIST.append(err_str)

            elif section == "paragraphs":
                paragraphs = self.document_property['paragraphs']
                for paragraph in paragraphs:
                    # print(paragraph)
                    if paragraph['style']['outlineLvl'] is not None and paragraph['text_property']['is_bold'] is True:
                        print(paragraph)

        return self.ERR_LIST


if __name__ == '__main__':
    nc = Normocontrol('380305__Фокина_ОС.docx', 'default_property.json')
    print(nc.run())

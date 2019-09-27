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

        for name in self.default_property:
            path_to_properties = self.default_property[name]['path']
            if path_to_properties == "document_body_property":
                body_properties = self.document_property[path_to_properties]
                default_body = self.default_property[name]
                default_body_properties = self.default_property[name]['properties']
                for body_property in body_properties:

                    if int(body_properties[body_property]) == default_body_properties[body_property]:
                        print('{}:{} = {}'.format(body_property, body_properties[body_property],
                                                  default_body_properties[body_property]))
                    else:

                        err_str = default_body['error_message'].format(
                            '{}:{} != {}'.format(body_property, body_properties[body_property],
                                                 default_body_properties[body_property]))
                        self.ERR_LIST.append(err_str)
            elif path_to_properties == "paragraphs":
                paragraphs = self.document_property['paragraphs']
                for paragraph in paragraphs:
                    print(paragraph)

        return self.ERR_LIST


if __name__ == '__main__':
    nc = Normocontrol('document.docx', 'default_property.json')
    print(nc.run())

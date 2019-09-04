import json
from DOCX import DOCX


class Normocontrol(DOCX):
    def __init__(self, document_path, path_to_default_property_file):
        super().__init__(document_path)
        self.path_to_default_property_file = path_to_default_property_file
        self.default_property = self.get_default_properties()

    def get_default_properties(self):
        with open(self.path_to_default_property_file, encoding="utf-8") as default_property_file:
            default_property = json.loads(default_property_file.read())
        return default_property





if __name__ == '__main__':
    nc = Normocontrol('Pravila_oformlenia_poyasnitelnoy_zapiski_09_04_18.docx', 'default_property.json')
    print(nc)
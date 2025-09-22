
from logging import root
import xml.etree.ElementTree as ET

class Text:
    def __init__(self, text: str = "", preserve_space: bool = False):
        self._text: str = text
        self._preserve_space: bool = preserve_space
        self._xml: str = None
        self._init_xml()

    @property
    def text(self) -> str:
        if self._preserve_space:
            return self._text
        else:
            return self._text.strip()

    @text.setter
    def text(self, value: str):
        self._text = value
        element = ET.fromstring(self._xml)
        element.text = self._text
        self._xml = ET.tostring(element, encoding='unicode')

    @property
    def preserve_space(self) -> bool:
        return self._preserve_space
    
    @preserve_space.setter
    def preserve_space(self, value: bool):
        self._preserve_space = value
        if self._preserve_space:
            element = ET.fromstring(self._xml)
            element.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
            self._xml = ET.tostring(element, encoding='unicode')
        else:
            element = ET.fromstring(self._xml)
            if '{http://www.w3.org/XML/1998/namespace}space' in element.attrib:
                del element.attrib['{http://www.w3.org/XML/1998/namespace}space']
            self._xml = ET.tostring(element, encoding='unicode')

    @property
    def xml(self) -> str:
        return self._xml
    
    @xml.setter
    def xml(self, value: str):
        self._xml = value
        # 解析 XML 并更新文本和保留空格属性
        element = ET.fromstring(self._xml)
        self._text = element.text or ""
        if element.get('{http://www.w3.org/XML/1998/namespace}space'):
            self._preserve_space = element.get('{http://www.w3.org/XML/1998/namespace}space') == 'preserve'
        
    def _init_xml(self):
        if not self._xml:
            element = ET.Element('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}t')
            element.text = self._text
            if self._preserve_space:
                element.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
            self._xml = ET.tostring(element, encoding='unicode')

    def to_xml(self) -> ET.Element:
        return ET.fromstring(self._xml)
    
    @classmethod
    def load_from_xml(cls, element: ET.Element):
        text = cls()
        text._xml = ET.tostring(element, encoding='unicode')
        return text
    
    @classmethod
    def from_xml_str(cls, xml: str):
        text = cls()
        text._xml = xml
        return text
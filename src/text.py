
import xml.etree.ElementTree as ET

class Text:
    def __init__(self, text: str = ""):
        self._text = text

    @property
    def text(self) -> str:
        return self._text

    @text.setter
    def text(self, value: str):
        self._text = value

    def to_xml(self) -> ET.Element:
        # 将文本转换为 XML 表示
        element = ET.Element('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}t')
        element.text = self._text
        element.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
        return element

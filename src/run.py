
from text import Text
import xml.etree.ElementTree as ET
from golbal import NAMESPACES
from run_properties import RunProperties

class Run:
    def __init__(self):
        self._text: str = ""
        self._xml: str = ""
        self._rpr: RunProperties = None
        self._texts: list[Text] = []

    @property
    def text(self) -> str:
        return self._text

    @property
    def xml(self) -> str:
        return self._xml

    @property
    def rpr(self) -> RunProperties:
        return self._rpr

    @property
    def texts(self) -> list:
        return self._texts
    
    @text.setter
    def text(self, value: str):
        self._text = value
        self._text_update()

    @rpr.setter
    def rpr(self, value: RunProperties):
        self._rpr = value
        self._rpr_update()

    @texts.setter
    def texts(self, value: list):
        self._texts = value
        self._texts_update()

    @xml.setter
    def xml(self, value: str):
        self._xml = value
        self._xml_update()

    def _texts_update(self):
        """
        根据文本对象更新文本相关属性
        """
        self._text = ""
        for text in self._texts:
            self._text += text.text

    def _rpr_update(self):
        """
        根据属性变化更新xml标签
        """
        rpr_element = self._rpr.to_xml_element()
        xml_tree = ET.fromstring(self._xml)
        for child in xml_tree:
            if child.tag.endswith('rPr'):
                xml_tree.remove(child)
        xml_tree.insert(0, rpr_element)
        self._xml = ET.tostring(xml_tree, encoding='unicode')

    def _text_update(self):
        """
        根据文本的更新更新文本对象的其他文本表示相关元素
        """
        # 更新文本对象
        self._texts = []
        t = Text(self._text)
        self._texts.append(t)
        # 更新 XML 表示
        t_element = t.to_xml()
        xml_tree = ET.fromstring(self._xml)
        for child in xml_tree:
            if child.tag.endswith('t'):
                xml_tree.remove(child)
        xml_tree.append(t_element)
        self._xml = ET.tostring(xml_tree, encoding='unicode')

    def _xml_update(self):
        """
        根据 XML 的更新更新文本对象的其他元素
        """
        xml_tree = ET.fromstring(self._xml)
        texts = xml_tree.findall('.//w:t', NAMESPACES)
        # 更新text和texts
        self._text = ""
        self._texts = []
        for text in texts:
            is_preserve_space = text.get('{http://www.w3.org/XML/1998/namespace}space') == 'preserve'
            if is_preserve_space:
                t = text.text if text.text else ''
                self._text += t
                self._texts.append(Text(t))
            else:
                t = (text.text or '').strip()
                self._text += t
                self._texts.append(Text(t))

        # 更新文本属性
        rpr = xml_tree.find('.//w:rPr', NAMESPACES)
        if rpr is not None:
            self._rpr = RunProperties.load_from_xml(rpr)

    @classmethod
    def from_xml(cls, xml: ET.Element):
        run = cls()
        run.xml = ET.tostring(xml, encoding='unicode')

        return run

    @classmethod
    def from_xml_str(cls, xml: str):
        root = ET.fromstring(xml)
        return cls.from_xml(root)
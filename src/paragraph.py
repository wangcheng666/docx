
import xml.etree.ElementTree as ET



class Paragraph:
    def __init__(self):
        self.text = ""
        self.xml = ""
        self.ppr = None
        self.runs = []
        self.comments = []
        self.revisions = []
        self.texts = []

    @classmethod
    def from_xml_str(cls, xml: str):
        root = ET.fromstring(xml)
        return cls.from_xml(root)

    @classmethod
    def from_xml(cls, xml: ET.Element):
        paragraph = cls()
        paragraph.xml = ET.tostring(xml, encoding='unicode')
        paragraph.ppr = xml.find("ppr").text
        paragraph.runs = [run.text for run in xml.findall("run")]
        paragraph.comments = [comment.text for comment in xml.findall("comment")]
        paragraph.revisions = [revision.text for revision in xml.findall("revision")]
        paragraph.texts = [text.text for text in xml.findall("text")]
        return paragraph

    @staticmethod
    def is_valid_paragraph(xml: ET.Element) -> bool:
        # Check if the XML element is a valid paragraph
        return xml.tag == "p"

    # 功能：
        # 获取段落文本
        # 获取段落的 XML 表示
        # 获取段落的 PPR 表示
        # 获取段落的 Run 列表
            # 自动更新text
        # 修改段落内容
            # 自动更新段落的runs
        # 根据xml生成段落对象
        # 
    def get_text(self) -> str:
        return self.text

    def get_xml(self) -> str:
        return self.xml

    def get_ppr(self) -> str:
        return self.ppr

    def get_runs(self) -> list:
        return self.runs

    def set_text(self, new_text: str):
        self.text = new_text
from dataclasses import dataclass
from typing import Optional
from datetime import datetime
import xml.etree.ElementTree as ET
import logging

@dataclass
class Comment:
    id: int
    author: str = "文档审核助手"
    initials: str = "WA"
    text: str = ""
    date: Optional[datetime] = None
    para_id: Optional[str] = None
    
    def to_xml_element(self):
        """转换为XML元素"""
        if self.date is None:
            self.date = datetime.now()
        if self.para_id is None:
            logging.warning("Para ID is not set")
            raise ValueError("Para ID must be set before converting to XML element")
        # 创建批注元素
        comment = ET.Element('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}comment')
        comment.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}id', str(self.id))
        comment.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}author', self.author)
        comment.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}date', 
                   self.date.strftime("%Y-%m-%dT%H:%M:%SZ"))
        comment.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}initials', self.initials)
        
        # 创建段落元素
        p = ET.SubElement(comment, '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}p')
        p.set('{http://schemas.microsoft.com/office/word/2010/wordml}paraId', self.para_id)
        
        # 创建段落属性
        pPr = ET.SubElement(p, '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}pPr')
        pStyle = ET.SubElement(pPr, '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}pStyle')
        pStyle.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val', "CommentStyle")
        
        # 段落字体属性
        rPr_para = ET.SubElement(pPr, '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}rPr')
        rFonts_para = ET.SubElement(rPr_para, '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}rFonts')
        rFonts_para.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}hint', "default")
        rFonts_para.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}eastAsia', "宋体")
        lang_para = ET.SubElement(rPr_para, '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}lang')
        lang_para.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val', "en-US")
        lang_para.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}eastAsia', "zh-CN")
        
        # 创建文本运行
        r = ET.SubElement(p, '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}r')
        
        # 文本运行属性
        rPr = ET.SubElement(r, '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}rPr')
        rFonts = ET.SubElement(rPr, '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}rFonts')
        rFonts.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}hint', "eastAsia")
        lang = ET.SubElement(rPr, '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}lang')
        lang.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val', "en-US")
        lang.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}eastAsia', "zh-CN")
        
        # 文本内容
        t = ET.SubElement(r, '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}t')
        t.text = self.text
        
        return comment

    

from dataclasses import dataclass
from font import Font
import xml.etree.ElementTree as ET

@dataclass
class RunProperties:
    font: Font = None  # 字体名称
    size: int = None   # 字体大小
    size_cs: int = None  # 字体大小（复杂脚本）
    color: str = None  # 字体颜色
    bold: bool = None  # 是否加粗
    italic: bool = None  # 是否斜体
    italic_cs: bool = None  # 是否复杂脚本斜体
    highlight_color: str = None  # 高亮颜色
    kern: int = None  # 字偶间距
    spacing: int = None  # 字符间距

    def to_xml_element(self):
        """转换为XML元素"""
        rpr = ET.Element('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}rPr')
        # 字体样式
        if self.font:
            rFonts = ET.SubElement(rpr, '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}rFonts')
            rFonts.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}ascii', self.font.ascii)
            rFonts.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}hAnsi', self.font.hAnsi)
            rFonts.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}eastAsia', self.font.eastAsia)
            if self.font.hint:
                rFonts.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}hint', self.font.hint)
        # 字体大小
        if self.size:
            sz = ET.SubElement(rpr, '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}sz')
            sz.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val', str(self.size))
        if self.size_cs is not None:
            szCs = ET.SubElement(rpr, '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}szCs')
            szCs.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val', str(self.size))
        
        if self.bold:
            b = ET.SubElement(rpr, '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}b')   
        if self.italic:
            i = ET.SubElement(rpr, '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}i')
        if self.italic_cs:
            iCs = ET.SubElement(rpr, '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}iCs')
        if self.color:
            color = ET.SubElement(rpr, '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}color')
            color.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val', self.color)
        if self.highlight_color:
            highlight = ET.SubElement(rpr, '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}highlight')
            highlight.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val', self.highlight_color)

        if self.kern:
            kern = ET.SubElement(rpr, '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}kern')
            kern.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val', str(self.kern))
        if self.spacing:
            spacing = ET.SubElement(rpr, '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}spacing')
            spacing.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val', str(self.spacing))

        return rpr

    @classmethod
    def load_from_xml(cls, rPr: ET.Element):
        if rPr is None:
            return None
        run_properties = cls()

        rFonts = rPr.find('.//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}rFonts')
        if rFonts is not None:
            font = Font()
            assic = rFonts.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}ascii')
            if assic:
                font.ascii = assic
            hAnsi = rFonts.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}hAnsi')
            if hAnsi:
                font.hAnsi = hAnsi
            eastAsia = rFonts.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}eastAsia')
            if eastAsia:
                font.eastAsia = eastAsia
            hint = rFonts.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}hint')
            if hint:
                font.hint = hint
            run_properties.font = font

        sz = rPr.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}sz')
        if sz:
            run_properties.size = sz
        
        szCs = rPr.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}szCs')
        if szCs:
            run_properties.size_cs = szCs

        b = rPr.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}b')
        if b:
            run_properties.bold = True
        
        i = rPr.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}i')
        if i:
            run_properties.italic = True

        iCs = rPr.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}iCs')
        if iCs:
            run_properties.italic_cs = True

        color = rPr.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}color')
        if color:
            color_val = color.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val')
            if color_val:
                run_properties.color = color_val

        highlight = rPr.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}highlight')
        if highlight:
            highlight_val = highlight.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val')
            if highlight_val:
                run_properties.highlight = highlight_val

        kern = rPr.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}kern')
        if kern:
            kern_val = kern.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val')
            if kern_val:
                run_properties.kern = int(kern_val)
        
        spacing = rPr.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}spacing')
        if spacing:
            spacing_val = spacing.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val')
            if spacing_val:
                run_properties.spacing = int(spacing_val)

        return run_properties
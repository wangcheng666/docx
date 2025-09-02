import enum
from xml.etree import ElementTree as ET
from dataclasses import dataclass

class Alignment(enum.Enum):
    LEFT = 'left'
    CENTER = 'center'
    RIGHT = 'right'
    JUSTIFY = 'justify'

class Styles(enum.Enum):
    TITLE = 'TitleStyle'
    SUBTITLE_CENTER = 'SubtitleStyleCenter'
    SUBTITLE_LEFT = 'SubtitleStyleLeft'
    SUBSUBTITLE_CENTER = 'SubSubtitleStyleCenter'
    SUBSUBTITLE_LEFT = 'SubSubtitleStyleLeft'
    SUBSUBSUBTITLE_LEFT = 'SubSubSubtitleStyleLeft'
    BODY = 'BodyStyle'
    SPECIFIC_RIGHT = 'SpecificStyleRight'
    APPELLATION = 'AppellationStyle'

@dataclass
class Font:
    ascii: str = ""
    hAnsi: str = ""
    eastAsia: str = ""
    hint: str = ""

@dataclass
class RunStyleProperties:
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

@dataclass
class ParaStyleProperties:
    name: Styles = Styles.BODY
    alignment: Alignment = None  # 段落对齐方式
    snapToGrid: bool = None          # 是否对齐到网格
    first_line_indent: int = None       # 首行缩进
    first_line_chars: int = None        # 首行缩进字符数
    space_before: int = None            # 段前间距
    space_before_lines: int = None      # 段前间距（行数）
    space_before_autospacing: bool = None   # 是否自动调整段前间距
    space_after: int = None             # 段后间距
    space_after_lines: int = None      # 段后间距（行数）
    space_after_autospacing: bool = None   # 是否自动调整段后间距
    space_line: int = None               # 行间距
    space_line_rule: str = None  # 行间距规则

    widow_control: bool = None       # 是否启用分页控制
    rpr: RunStyleProperties = None  # 关联的RunStyleProperties对象


    def to_xml_element(self):
        """转换为XML元素"""
        # 创建段落属性元素
        ppr = ET.Element('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}pPr')
        # 对齐方式
        if self.alignment:
            jc = ET.SubElement(ppr, '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}jc')
            if self.alignment == Alignment.LEFT:
                jc.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val', 'left')
            elif self.alignment == Alignment.CENTER:
                jc.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val', 'center')
            elif self.alignment == Alignment.RIGHT:
                jc.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val', 'right')
        # 首行缩进
        if self.first_line_indent or self.first_line_chars:
            ind = ET.SubElement(ppr, '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}ind')
            if self.first_line_indent:
                ind.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}firstLine', str(self.first_line_indent))
            if self.first_line_chars:
                ind.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}firstLineChars', str(self.first_line_chars))
        # 对齐到网格
        if self.snapToGrid is not None:
            if self.snapToGrid:
                snap = ET.SubElement(ppr, '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}snapToGrid')
                snap.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val', '1')
            else:
                snap = ET.SubElement(ppr, '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}snapToGrid')
                snap.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val', '0')
        # 段前段后间距
        if self.space_before or self.space_before_autospacing or self.space_after or \
            self.space_after_autospacing or self.space_line or self.space_line_rule or \
            self.space_before_lines or self.space_after_lines:
            spacing = ET.SubElement(ppr, '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}spacing')
            if self.space_before:
                spacing.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}before', str(self.space_before))
            if self.space_before_autospacing is not None:
                if self.space_before_autospacing:
                    spacing.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}beforeAutospacing', str(1))
                else:
                    spacing.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}beforeAutospacing', str(0))
            if self.space_after:
                spacing.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}after', str(self.space_after))
            if self.space_after_autospacing is not None:
                if self.space_after_autospacing:
                    spacing.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}afterAutospacing', str(1))
                else:
                    spacing.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}afterAutospacing', str(0))
            if self.space_line:
                spacing.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}line', str(self.space_line))
            if self.space_line_rule:
                spacing.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}lineRule', self.space_line_rule)
            if self.space_before_lines:
                spacing.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}beforeLines', str(self.space_before_lines))
            if self.space_after_lines:
                spacing.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}afterLines', str(self.space_after_lines))
        

        if self.widow_control:
            widow_control = ET.SubElement(ppr, '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}widowControl')

        if self.rpr:
            rpr_element = self.rpr.to_xml_element()
            ppr.append(rpr_element)

        return ppr

    @classmethod
    def load_from_xml(cls, pPr: ET.Element):
        if pPr is None:
            return None

        para_properties = cls()
        window_control = pPr.find('.//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}widowControl')
        if window_control is not None:
            para_properties.widow_control = True

        jc = pPr.find('.//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}jc')
        if jc is not None:
            jc_val = jc.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val')
            if jc_val == 'center':
                para_properties.alignment = Alignment.CENTER
            elif jc_val == 'left':
                para_properties.alignment = Alignment.LEFT
            elif jc_val == 'right':
                para_properties.alignment = Alignment.RIGHT
            elif jc_val == 'justify':
                para_properties.alignment = Alignment.JUSTIFY

        ind = pPr.find('.//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}ind')
        if ind is not None:
            first_line = ind.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}firstLine')
            if first_line is not None:
                para_properties.first_line_indent = int(first_line)
            first_line_chars = ind.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}firstLineChars')
            if first_line_chars:
                para_properties.first_line_chars = int(first_line_chars)

        snapToGrid = pPr.find('.//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}snapToGrid')
        if snapToGrid is not None:
            snapToGrid_val = snapToGrid.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val')
            if snapToGrid_val == '0':
                para_properties.snapToGrid = False
            else:
                para_properties.snapToGrid = True

        spacing = pPr.find('.//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}spacing')
        if spacing is not None:
            space_before = spacing.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}before')
            if space_before:
                para_properties.space_before = int(space_before)
            space_before_autospacing = spacing.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}beforeAutospacing')
            if space_before_autospacing:
                if space_before_autospacing == '1':
                    para_properties.space_before_autospacing = True
                else:
                    para_properties.space_before_autospacing = False
            space_after = spacing.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}after')
            if space_after:
                para_properties.space_after = int(space_after)
            space_after_autospacing = spacing.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}afterAutospacing')
            if space_after_autospacing:
                if space_after_autospacing == '1':
                    para_properties.space_after_autospacing = True
                else:
                    para_properties.space_after_autospacing = False

            space_line = spacing.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}line')
            if space_line:
                para_properties.space_line = int(space_line)
            space_line_rule = spacing.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}lineRule')
            if space_line_rule:
                para_properties.space_line_rule = space_line_rule
            space_before_lines = spacing.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}beforeLines')
            if space_before_lines:
                para_properties.space_before_lines = int(space_before_lines)
            space_after_lines = spacing.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}afterLines')
            if space_after_lines:
                para_properties.space_after_lines = int(space_after_lines)

        rPr = pPr.find('.//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}rPr')
        if rPr is not None:
            para_properties.rpr = RunStyleProperties.load_from_xml(rPr)

        return para_properties

    def describe(self):
        word_font_size_mapping = {
            # 中文字号 -> 磅值
            '初号': 42,
            '小初': 36,
            '一号': 26,
            '小一': 24,
            '二号': 22,
            '小二': 18,
            '三号': 16,
            '小三': 15,
            '四号': 14,
            '小四': 12,
            '五号': 10.5,
            '小五': 9,
            '六号': 7.5,
            '小六': 6.5,
            '七号': 5.5,
            '八号': 5,
        }
        size_to_word_font_mapping = {
            # 磅值 -> 中文字号
            42: '初号',
            36: '小初',
            26: '一号',
            24: '小一',
            22: '二号',
            18: '小二',
            16: '三号',
            15: '小三',
            14: '四号',
            12: '小四',
            10.5: '五号',
            9: '小五',
            7.5: '六号',
            6.5: '小六',
            5.5: '七号',
            5: '八号'
        }

        default_description = ["字体: 宋体", "字号: 五号"]
        description = []


        if self.rpr:
            if self.rpr.font:
                description.append(f"字体: {self.rpr.font.eastAsia if self.rpr.font.eastAsia else self.rpr.font.ascii}")
            if self.rpr.size:
                description.append(f"字号: {size_to_word_font_mapping.get(int(self.rpr.size)//2) if size_to_word_font_mapping.get(int(self.rpr.size)//2) else f'{int(self.rpr.size)//2} 磅'}")
            if self.rpr.bold:
                description.append("加粗")
            if self.rpr.italic:
                description.append("斜体")

        # if self.first_line_chars:
        #     description.append(f"首行缩进: {self.first_line_chars//100} 个字符")
        # elif self.first_line_indent:
        #     description.append(f"首行缩进: {self.first_line_indent//20} 磅")

        if self.alignment:
            align_map = {Alignment.LEFT: "左对齐", Alignment.CENTER: "居中", Alignment.RIGHT: "右对齐", Alignment.JUSTIFY: "两端对齐"}
            description.append(f"对齐方式: {align_map.get(self.alignment, self.alignment)}")


            # spacing = pPr.find('.//w:spacing', NAMESPACES)
            # if spacing is not None:
            #     before = spacing.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}before')
            #     if before:
            #         description.append(f"段前距: {int(before)/20:.2f} 磅")
            #     after = spacing.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}after')
            #     if after:
            #         description.append(f"段后距: {int(after)/20:.2f} 磅")
            #     line = spacing.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}line')
            #     if line:
            #         description.append(f"行距: {int(line)/20:.2f} 磅")
            #     lineRule = spacing.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}lineRule')
            #     if lineRule:
            #         description.append(f"行距规则: {lineRule}")

        if len(description) == 0:
            description = default_description
            
        return ", ".join(description)
    


@dataclass
class SectionProperties:
    """
    段落属性
    """
    page_size_width: int
    page_size_height: int
    page_margin_top: int
    page_margin_right: int
    page_margin_bottom: int
    page_margin_left: int
    page_header_space: int
    page_footer_space: int
    page_gutter_space: int
    cols_space: int
    cols_num: int
    doc_grid_type: str
    doc_grid_line_pitch: int
    doc_grid_char_space: int

    def to_xml_element(self):
        sectPr = ET.Element('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}sectPr')
        pgSz = ET.SubElement(sectPr, '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}pgSz')
        pgSz.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}w', str(self.page_size_width))
        pgSz.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}h', str(self.page_size_height))
        pgMar = ET.SubElement(sectPr, '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}pgMar')
        pgMar.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}top', str(self.page_margin_top))
        pgMar.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}right', str(self.page_margin_right))
        pgMar.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}bottom', str(self.page_margin_bottom))
        pgMar.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}left', str(self.page_margin_left))
        pgMar.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}header', str(self.page_header_space))
        pgMar.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}footer', str(self.page_footer_space))
        pgMar.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}gutter', str(self.page_gutter_space))
        cols = ET.SubElement(sectPr, '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}cols')
        cols.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}space', str(self.cols_space))
        cols.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}num', str(self.cols_num))
        docGrid = ET.SubElement(sectPr, '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}docGrid')
        docGrid.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}type', self.doc_grid_type)
        docGrid.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}linePitch', str(self.doc_grid_line_pitch))
        docGrid.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}charSpace', str(self.doc_grid_char_space))
        return sectPr


official_document_page_style = SectionProperties(
    page_size_width=11906,
    page_size_height=16838,
    page_margin_top=1440,
    page_margin_right=1800,
    page_margin_bottom=1440,
    page_margin_left=1800,
    page_header_space=851,
    page_footer_space=992,
    page_gutter_space=0,
    cols_space=425,
    cols_num=1,
    doc_grid_type="lines",
    doc_grid_line_pitch=312,
    doc_grid_char_space=0
)


title_style_para = ParaStyleProperties(
    name=Styles.TITLE,
    alignment=Alignment.CENTER,
    widow_control=True,
    snapToGrid=False,
    rpr=RunStyleProperties(
        font=Font(
            ascii="方正小标宋简体",
            hAnsi="宋体",
            eastAsia="方正小标宋简体",
        ),
        size=44,
        size_cs=44,
    )
)



subtitle_style_para_center = ParaStyleProperties(
    name=Styles.SUBTITLE_CENTER,
    alignment=Alignment.CENTER,
    # space_before=240,
    # space_line=600,
    # space_line_rule="exact",
    # widow_control=True,
    # snapToGrid=False,
    rpr=RunStyleProperties(
        font=Font(
            ascii="黑体",
            hAnsi="黑体",
            eastAsia="黑体",
        ),
        # bold=True,
        # # color="000000",
        size=32,
        size_cs=32,
        # kern=0
    )
)

subtitle_style_para_left = ParaStyleProperties(
    name=Styles.SUBTITLE_LEFT,
    alignment=Alignment.LEFT,
    first_line_indent=640,
    first_line_chars=200,
    rpr=RunStyleProperties(
        font=Font(
            ascii="黑体",
            hAnsi="黑体",
            eastAsia="黑体",
        ),
        size=32,
        size_cs=32,
    )
)


subsubtitle_style_para_center = ParaStyleProperties(
    name=Styles.SUBSUBTITLE_CENTER,
    alignment=Alignment.CENTER,
    # widow_control=True,
    # snapToGrid=False,
    # space_line=600,
    # space_line_rule="exact",
    rpr=RunStyleProperties(
        font=Font(
            ascii="楷体",
            hAnsi="楷体",
            eastAsia="楷体",
        ),
        size=32,
        size_cs=32,
    )
)

subsubtitle_style_para_left = ParaStyleProperties(
    name=Styles.SUBSUBTITLE_LEFT,
    alignment=Alignment.LEFT,
    first_line_indent=640,
    first_line_chars=200,
    # space_line=360,
    # space_line_rule="auto",
    rpr=RunStyleProperties(
        font=Font(
            ascii="楷体",
            hAnsi="楷体",
            eastAsia="楷体",
        ),
        size=32,
        size_cs=32,
    )
)

subsubsubtitle_style_para_left = ParaStyleProperties(
    name=Styles.SUBSUBSUBTITLE_LEFT,
    alignment=Alignment.LEFT,
    first_line_indent=643,
    first_line_chars=200,
    rpr=RunStyleProperties(
        font=Font(
            ascii="仿宋",
            hAnsi="仿宋",
            eastAsia="仿宋",
        ),
        size=32,
        size_cs=32,
    )
)

body_style_para = ParaStyleProperties(
    name=Styles.BODY,
    alignment=Alignment.LEFT,
    first_line_indent=640,
    first_line_chars=200,
    rpr=RunStyleProperties(
        font=Font(
            ascii="仿宋",
            hAnsi="仿宋",
            eastAsia="仿宋",
        ),
        size=32,
        size_cs=32,
    )
)
# 称谓样式
appellation_style_para = ParaStyleProperties(
    name=Styles.APPELLATION,
    alignment=Alignment.LEFT,
    rpr=RunStyleProperties(
        font=Font(
            ascii="仿宋",
            hAnsi="仿宋",
            eastAsia="仿宋",
        ),
        size=32,
        size_cs=32,
    )
)

specific_style_para_right = ParaStyleProperties(
    name=Styles.SPECIFIC_RIGHT,
    alignment=Alignment.RIGHT,
    rpr=RunStyleProperties(
        font=Font(
            ascii="仿宋",
            hAnsi="仿宋",
            eastAsia="仿宋",
        ),
        size=32,
        size_cs=32,
    )
)

# ET.register_namespace('w', 'http://schemas.openxmlformats.org/wordprocessingml/2006/main')

# print("title_style_para:")
# print(ET.tostring(title_style_para.to_xml_element(), encoding='utf-8').decode('utf-8'))
# print("subtitle_style_para_center:")
# print(ET.tostring(subtitle_style_para_center.to_xml_element(), encoding='utf-8').decode('utf-8'))
# print("subtitle_style_para_left:")
# print(ET.tostring(subtitle_style_para_left.to_xml_element(), encoding='utf-8').decode('utf-8'))
# print("subsubtitle_style_para_center:")
# print(ET.tostring(subsubtitle_style_para_center.to_xml_element(), encoding='utf-8').decode('utf-8'))
# print("subsubtitle_style_para_left:")
# print(ET.tostring(subsubtitle_style_para_left.to_xml_element(), encoding='utf-8').decode('utf-8'))
# print("body_style_para:")
# print(ET.tostring(body_style_para.to_xml_element(), encoding='utf-8').decode('utf-8'))



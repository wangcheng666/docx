
import zipfile
import os
import re
import logging
import xml.etree.ElementTree as ET
from datetime import datetime
from app.core.doc_inspection.base.para_id_generator import ParaIdGenerator
from app.core.doc_inspection.base.comment import Comment
from app.core.doc_inspection.base.styles import *
import copy

from app.core.doc_inspection.base.task import AddCommentTask, ApplyStyleTask, CheckStyleTask, DeleteTask, InsertTask, Task, TaskType

# 模板文件路径
# 获取当前脚本所在目录
file_dir = os.path.dirname(os.path.abspath(__file__))
TEMPLATECOMMENTS = os.path.join(file_dir, 'templates', 'comments.xml')
TEMPLATECOMMENTSEXTENDED = os.path.join(file_dir, 'templates', 'commentsExtended.xml')

PAGE_SETTINGS_TEMPLATE = '''<w:sectPr>
            <w:pgSz w:w="11906" w:h="16838" />
            <w:pgMar w:top="1440" w:right="1800" w:bottom="1440" w:left="1800" w:header="851"
                w:footer="992" w:gutter="0" />
            <w:cols w:space="425" w:num="1" />
            <w:docGrid w:type="lines" w:linePitch="312" w:charSpace="0" />
        </w:sectPr>'''

NAMESPACES = {
    'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
    'w14': 'http://schemas.microsoft.com/office/word/2010/wordml',
    'wpc': 'http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas',
    'cx': 'http://schemas.microsoft.com/office/drawing/2014/chartex',
    'cx1': 'http://schemas.microsoft.com/office/drawing/2015/9/8/chartex',
    'cx2': 'http://schemas.microsoft.com/office/drawing/2015/10/21/chartex',
    'cx3': 'http://schemas.microsoft.com/office/drawing/2016/5/9/chartex',
    'cx4': 'http://schemas.microsoft.com/office/drawing/2016/5/10/chartex',
    'cx5': 'http://schemas.microsoft.com/office/drawing/2016/5/11/chartex',
    'cx6': 'http://schemas.microsoft.com/office/drawing/2016/5/12/chartex',
    'cx7': 'http://schemas.microsoft.com/office/drawing/2016/5/13/chartex',
    'cx8': 'http://schemas.microsoft.com/office/drawing/2016/5/14/chartex',
    'mc': 'http://schemas.openxmlformats.org/markup-compatibility/2006',
    'aink': 'http://schemas.microsoft.com/office/drawing/2016/ink',
    'am3d': 'http://schemas.microsoft.com/office/drawing/2017/model3d',
    'o': 'urn:schemas-microsoft-com:office:office',
    'oel': 'http://schemas.microsoft.com/office/2019/extlst',
    'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
    'v': 'urn:schemas-microsoft-com:vml',
    'wp14': 'http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing',
    'wp': 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing',
    'w10': 'urn:schemas-microsoft-com:office:word',
    'w16cex': 'http://schemas.microsoft.com/office/word/2018/wordml/cex',
    'w16cid': 'http://schemas.microsoft.com/office/word/2016/wordml/cid',
    'w16': 'http://schemas.microsoft.com/office/word/2018/wordml',
    'w16sdtdh': 'http://schemas.microsoft.com/office/word/2020/wordml/sdtdatahash',
    'w16se': 'http://schemas.microsoft.com/office/word/2015/wordml/symex',
    'wpg': 'http://schemas.microsoft.com/office/word/2010/wordprocessingGroup',
    'wpi': 'http://schemas.microsoft.com/office/word/2010/wordprocessingInk',
    'wne': 'http://schemas.microsoft.com/office/word/2006/wordml',
    'wps': 'http://schemas.microsoft.com/office/word/2010/wordprocessingShape',
    'wpsCustomData': 'http://www.wps.cn/officeDocument/2013/wpsCustomData',
    'm': 'http://schemas.openxmlformats.org/officeDocument/2006/math',
    'w15': 'http://schemas.microsoft.com/office/word/2012/wordml'
}

NUMBERING_LEVELS = {"chineseCounting": "chineseCounter"}


class DocxFileManager:
    _instances = {}
    def __new__(cls, docx_path, *args, **kwargs):
        abs_path = os.path.abspath(docx_path)
        if abs_path not in cls._instances:
            instance = super().__new__(cls)
            cls._instances[abs_path] = instance
        return cls._instances[abs_path]
    
    def __init__(self, docx_path):
        if hasattr(self, 'initialized') and self.initialized:
            return
        self.initialized = True
        self.docx_path = docx_path
        # 先初始化日志记录器
        logging.basicConfig(level=logging.DEBUG, 
                            format='%(asctime)s - %(name)s - %(levelname)s - %(message)s')
        self.logger = logging.getLogger(__name__)

        self.extract_path = self._extract_docx()
        self.comments_path = f"{self.extract_path}/word/comments.xml"
        self.document_path = f"{self.extract_path}/word/document.xml"
        self.comments_extended_path = f"{self.extract_path}/word/commentsExtended.xml"
        self.footnotes_path = f"{self.extract_path}/word/footnotes.xml"
        self.headers_path = f"{self.extract_path}/word/header1.xml"  # 可能有多个header
        self.footers_path = f"{self.extract_path}/word/footer1.xml"  # 可能有多个footer
        self.styles_path = f"{self.extract_path}/word/styles.xml"
        self.rels_path = f"{self.extract_path}/word/_rels/document.xml.rels"
        self.numbering_path = f"{self.extract_path}/word/numbering.xml"

        self._register_namespaces()  # 注册命名空间

        self.para_id_generator = ParaIdGenerator(self.extract_path)

        self.document_tree = ET.parse(self.document_path)

        self.comments = {}
        self._load_comments()  # 用于存储批注
        self.comments_extended = {}
        self._load_comments_extended()  # 用于存储扩展批注
        self.author_name = "文档审核助手"  # 默认作者名称
        self.numbers = {}
        self.cur_number_count = {}
        self._parse_numbering()
        

    def _extract_docx(self):
        """解压缩docx文件到指定目录"""
        extract_path = self.docx_path.replace('.docx', '_extracted')
        with zipfile.ZipFile(self.docx_path, 'r') as zip_ref:
            zip_ref.extractall(extract_path)
        return extract_path

    def compress_docx(self, is_overwrite: bool=False):
        """将解压后的文件重新打包为docx格式"""
        if is_overwrite:
            zip_path = self.docx_path
        else:
            zip_path = self.docx_path.replace('.docx', '_repacked.docx')
        with zipfile.ZipFile(zip_path, 'w') as zip_ref:
            for foldername, subfolders, filenames in os.walk(self.extract_path):
                for filename in filenames:
                    file_path = os.path.join(foldername, filename)
                    arcname = os.path.relpath(file_path, self.extract_path)
                    zip_ref.write(file_path, arcname)
        self.logger.debug(f"已重新打包为: {zip_path}")
        # self._clear_extracted()  # 清理解压后的文件夹

    def _clear_extracted(self):
        """清理解压后的文件夹"""
        if os.path.exists(self.extract_path):
            for foldername, subfolders, filenames in os.walk(self.extract_path, topdown=False):
                for filename in filenames:
                    file_path = os.path.join(foldername, filename)
                    os.remove(file_path)
                for subfolder in subfolders:
                    os.rmdir(os.path.join(foldername, subfolder))
            os.rmdir(self.extract_path)
            self.logger.debug(f"已清理解压后的文件夹: {self.extract_path}")

    def _register_namespaces(self):
        """
        注册所有必要的命名空间，以确保生成的XML使用正确的前缀
        """
        for prefix, uri in NAMESPACES.items():
            ET.register_namespace(prefix, uri)
        self.logger.debug("已注册所有必要的命名空间")

    def _load_comments(self):
        """从comments.xml中获取批注"""
        if not os.path.exists(self.comments_path):
            self.logger.debug(f"Comments file not found at {self.comments_path}.")
            return 
        
        tree = ET.parse(self.comments_path)
        root = tree.getroot()
        self.comments = {}
        for comment in root.findall('.//w:comment', NAMESPACES):
            comment_id = comment.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}id')
            self.comments[comment_id] = comment
        self.logger.debug(f"已加载 {len(self.comments)} 条批注")
    
    def _load_comments_extended(self):
        """从commentsExtended.xml中获取扩展批注"""
        if not os.path.exists(self.comments_extended_path):
            self.logger.debug(f"扩展批注文件不存在: {self.comments_extended_path}.")
            return 
        
        tree = ET.parse(self.comments_extended_path)
        root = tree.getroot()
        self.comments_extended = {}
        for commentEx in root.findall('.//w15:commentEx', NAMESPACES):
            commentEx_paraId = commentEx.get('{http://schemas.microsoft.com/office/word/2012/wordml}paraId')
            self.comments_extended[commentEx_paraId] = commentEx
        self.logger.debug(f"已加载 {len(self.comments_extended)} 条扩展批注")

    def _set_comment_relationship(self):
        """在document.xml.rels中设置批注关系"""
        if not os.path.exists(self.rels_path):
            self.logger.debug(f"关系文件不存在: {self.rels_path}. 请确保docx文件包含document.xml.rels。")
            raise FileNotFoundError(f"关系文件不存在: {self.rels_path}. 请确保docx文件包含document.xml.rels。")

        tree = ET.parse(self.rels_path)
        root = tree.getroot()
        # 清理根元素的文本内容（移除空白字符）
        if root.text and root.text.strip() == '':
            root.text = None

        # 检查是否已存在批注关系
        comments_relationship_exists = False
        comments_extended_relationship_exists = False

        rels = root.findall('.//r:Relationship', {'r': 'http://schemas.openxmlformats.org/package/2006/relationships'})
        
        for rel in rels:
            rel_type = rel.get('Type')
            if rel_type == 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments':
                comments_relationship_exists = True
            elif rel_type == 'http://schemas.microsoft.com/office/2011/relationships/commentsExtended':
                comments_extended_relationship_exists = True

        # 获取现有关系的最大ID
        existing_ids = []
        for rel in rels:
            rel_id = rel.get('Id')
            if rel_id and rel_id.startswith('rId'):
                try:
                    existing_ids.append(int(rel_id[3:]))  # 提取rId后面的数字
                except ValueError:
                    pass
        
        next_id = max(existing_ids) + 1 if existing_ids else 1

        # 添加comments.xml关系（如果不存在）
        if not comments_relationship_exists:
            comments_rel = ET.Element('{http://schemas.openxmlformats.org/package/2006/relationships}Relationship')
            comments_rel.set('Id', f'rId{next_id}')
            comments_rel.set('Type', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments')
            comments_rel.set('Target', 'comments.xml')
            root.append(comments_rel)
            self.logger.debug(f"已添加comments.xml关系: rId{next_id}")
            next_id += 1

        # 添加commentsExtended.xml关系（如果不存在）
        if not comments_extended_relationship_exists:
            comments_ext_rel = ET.Element('{http://schemas.openxmlformats.org/package/2006/relationships}Relationship')
            comments_ext_rel.set('Id', f'rId{next_id}')
            comments_ext_rel.set('Type', 'http://schemas.microsoft.com/office/2011/relationships/commentsExtended')
            comments_ext_rel.set('Target', 'commentsExtended.xml')
            root.append(comments_ext_rel)
            self.logger.debug(f"已添加commentsExtended.xml关系: rId{next_id}")

        # 保存修改后的关系文件
        tree.write(self.rels_path, encoding='utf-8', xml_declaration=True)
        self.logger.debug("批注关系设置完成")

    def _set_comment(self, comment: Comment):
        """在comments.xml中设置批注，并返回批注ID"""
        if not os.path.exists(self.comments_path):
            self.logger.debug(f"注释文件不存在: {self.comments_path}. 创建一个新的。")
            tree = ET.parse(TEMPLATECOMMENTS)
            tree.write(self.comments_path, encoding='utf-8', xml_declaration=True)

        self._set_comment_style()
        
        tree = ET.parse(self.comments_path)
        root = tree.getroot()
        # 清理根元素的文本内容（移除空白字符）
        if root.text and root.text.strip() == '':
            root.text = None
        comment.id = str(len(self.comments) + 1)  # 生成新的批注ID
        comment.date = datetime.now()
        comment.para_id = self.para_id_generator.generate_unique_id()
        comment_element = comment.to_xml_element()
        root.append(comment_element)
        self.comments[comment.id] = comment_element
        tree.write(self.comments_path, encoding='utf-8', xml_declaration=True)
        self.logger.debug(f"已设置批注: {comment.id}")
        self._set_extended_comment(comment)
        
        return comment.id

    def _set_extended_comment(self, comment: Comment):
        """在commentsExtended.xml中设置扩展批注"""
        if not os.path.exists(self.comments_extended_path):
            self.logger.debug(f"扩展批注文件不存在: {self.comments_extended_path}. 创建一个新的。")
            tree = ET.parse(TEMPLATECOMMENTSEXTENDED)
            tree.write(self.comments_extended_path, encoding='utf-8', xml_declaration=True)
        
        tree = ET.parse(self.comments_extended_path)
        root = tree.getroot()
        # 清理根元素的文本内容（移除空白字符）
        if root.text and root.text.strip() == '':
            root.text = None
        commentEx = ET.Element('{http://schemas.microsoft.com/office/word/2012/wordml}commentEx')
        commentEx.set('{http://schemas.microsoft.com/office/word/2012/wordml}paraId', comment.para_id)
        commentEx.set('{http://schemas.microsoft.com/office/word/2012/wordml}done', "0")

        root.append(commentEx)
        self.comments_extended[comment.para_id] = commentEx
        tree.write(self.comments_extended_path, encoding='utf-8', xml_declaration=True)
        self.logger.debug(f"已设置扩展批注: {comment.para_id}")

    def _set_comment_style(self):
        """设置批注样式"""
        if not os.path.exists(self.styles_path):
            self.logger.debug(f"样式文件不存在: {self.styles_path}. 请确保docx文件包含styles.xml。")
            raise FileNotFoundError(f"样式文件不存在: {self.styles_path}. 请确保docx文件包含styles.xml。")
        tree = ET.parse(self.styles_path)
        root = tree.getroot()
        # 清理根元素的文本内容（移除空白字符）
        if root.text and root.text.strip() == '':
            root.text = None
        # 检查是否已经存在批注样式
        annotation_style_exists = False
        for style in root.findall('.//w:style', NAMESPACES):
            name_element = style.find('.//w:name', NAMESPACES)
            if name_element is not None and name_element.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val') == 'annotation text':
                annotation_style_exists = True
                self.logger.debug("批注样式已存在")
                break
        # 添加批注样式
        if not annotation_style_exists:
            # 创建新的批注样式
            comment_style = ET.Element('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}style')
            comment_style.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}type', 'paragraph')
            comment_style.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}styleId', 'CommentStyle')
            
            # 添加样式名称
            name_element = ET.SubElement(comment_style, '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}name')
            name_element.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val', 'annotation text')
            
            # 添加基于样式
            based_on_element = ET.SubElement(comment_style, '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}basedOn')
            based_on_element.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val', '1')
            
            # 添加 UI 优先级
            ui_priority_element = ET.SubElement(comment_style, '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}uiPriority')
            ui_priority_element.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val', '0')
            
            # 添加段落属性
            pPr_element = ET.SubElement(comment_style, '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}pPr')
            jc_element = ET.SubElement(pPr_element, '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}jc')
            jc_element.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val', 'left')
            
            # 添加到根元素
            root.append(comment_style)
            
            # 保存到文件
            tree.write(self.styles_path, encoding='utf-8', xml_declaration=True)
            self.logger.debug("已创建 'annotation text' 样式")

    def _parse_numbering(self):
        """解析文档中的编号信息"""
        if not os.path.exists(self.numbering_path):
            self.logger.debug(f"编号文件不存在: {self.numbering_path}.")
            return

        self.numbers.clear()  # 清除之前的编号信息
        tree = ET.parse(self.numbering_path)
        root = tree.getroot()

        abstract_nums = {}

        for abstract_num in root.findall('.//w:abstractNum', NAMESPACES):
            abstract_num_id = abstract_num.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}abstractNumId')
            if abstract_num_id is not None:
                lvls = {}
                for lvl in abstract_num.findall('.//w:lvl', NAMESPACES):
                    ilvl = lvl.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}ilvl')
                    if ilvl is not None:
                        numFmt = lvl.find('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}numFmt')
                        numFmt = numFmt.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val')
                        lvlText = lvl.find('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}lvlText')
                        lvlText = lvlText.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val')
                        start = lvl.find('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}start')
                        start = start.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val')
                        lvls[ilvl] = (numFmt, lvlText, start)
                if lvls:
                    abstract_nums[abstract_num_id] = lvls

        nums = root.findall('.//w:num', NAMESPACES)
        for num in nums:
            num_id = num.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}numId')
            if num_id is not None:
                self.numbers[num_id] = {}
            abstract_num_id = num.find('.//w:abstractNumId', NAMESPACES)
            abstract_num_id = abstract_num_id.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val') if abstract_num_id is not None else None

            if abstract_num_id is not None:
                abstract_num = abstract_nums.get(abstract_num_id)
                if abstract_num:
                    self.numbers[num_id] = abstract_num

        for num_id, abstract_num in self.numbers.items():
            self.logger.debug(f"编号 {num_id}: {abstract_num}")
            self.cur_number_count[num_id] = {}
            for ilvl, (numFmt, lvlText, start) in abstract_num.items():
                self.logger.debug(f"  级别 {ilvl}: 格式={numFmt}, 文本={lvlText}, 起始={start}")
                self.cur_number_count[num_id][ilvl] = int(start) if start.isdigit() else 1  # 确保起始值为整数

    # region 一些功能函数
    def _replace_regex(self, text, replacement, pattern):
        """ 使用正则表达式替换文本中的占位符
        :param text: 原始文本
        :param replacements: 替换字典
        :return: 替换后的文本
        """
        return re.sub(pattern, replacement, text)

    def _get_element_position_in_parent(self, parent, target_element):
        """ 获取目标元素在父元素中的位置
        :param parent: 父元素
        :param target_element: 目标元素
        :return: 目标元素在父元素中的索引位置，如果未找到则返回 -1
        """
        if not isinstance(parent, ET.Element):
            self.logger.debug("parent 必须是一个 Element 对象")
            return -1
        if not isinstance(target_element, ET.Element):
            self.logger.debug("target_element 必须是一个 Element 对象")
            return -1
        for i, child in enumerate(parent):
            if child == target_element:
                return i
        return -1  # 如果未找到目标元素，返回 -1

    def _insert_elements_after_element_recursive(self, parent_element, target_element, elements_to_insert) -> bool:
        """
        在指定元素之后递归插入元素列表
        :param parent_element: 父元素
        :param target_element: 目标元素
        :param elements_to_insert: 要插入的元素列表
        """
        # 查找目标元素在父元素中的位置
        for i, child in enumerate(parent_element):
            if child == target_element:
                # 在目标元素之后插入元素
                for element in reversed(elements_to_insert):
                    if isinstance(element, ET.Element):
                        parent_element.insert(i + 1, element)
                return True 
        
        # 如果未找到目标元素，递归检查子元素
        for child in parent_element:
            if isinstance(child, ET.Element):
                res = self._insert_elements_after_element_recursive(child, target_element, elements_to_insert)
                if res:
                    return True
        return False  # 如果未找到目标元素，返回False

    def _insert_elements_after_element(self, parent_element, target_element, elements_to_insert):
        """
        在指定元素之后插入元素列表
        :param parent_element: 父元素
        :param target_element: 目标元素
        :param elements_to_insert: 要插入的元素列表
        """
        if not isinstance(parent_element, ET.Element):
            self.logger.debug("parent_element 必须是一个 Element 对象")
            return
        if not isinstance(target_element, ET.Element):
            self.logger.debug("target_element 必须是一个 Element 对象")
            return
        if not isinstance(elements_to_insert, list):
            self.logger.debug("elements_to_insert 必须是一个列表")
            return
        res = self._insert_elements_after_element_recursive(parent_element, target_element, elements_to_insert)
        if not res:
            self.logger.debug("未找到目标元素，无法插入元素")

    def _splite_run(self, run, split_index):
        """
        将指定run在split_index处拆分为两个run
        :param run: 要拆分的run元素
        :param split_index: 拆分索引
        :return: 拆分后的两个run元素
        """
        if run is None or not isinstance(run, ET.Element):
            self.logger.debug("run必须是一个Element对象")
            return None, None
        
        texts = run.findall('.//w:t', NAMESPACES)
        if not texts:
            return None, None
        
        # 创建第一个新run
        new_run1 = ET.Element('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}r')
        new_run1_properties = run.find('.//w:rPr', NAMESPACES)
        if new_run1_properties is not None:
            new_run1.append(new_run1_properties)

        new_run2 = ET.Element('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}r')
        new_run2_properties = run.find('.//w:rPr', NAMESPACES)
        if new_run2_properties is not None:
            new_run2.append(new_run2_properties)

        # 计算拆分位置
        l = 0
        for i, text in enumerate(texts):
            l += len(text.text) if text.text else 0
            if l >= split_index:
                # 在当前文本节点中
                split_index_in_text = split_index - (l - len(text.text)) if (l - len(text.text)) < split_index else 0
                new_text1 = ET.Element('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}t')
                new_text1.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
                new_text1.text = text.text[:split_index_in_text] if text.text else ''
                new_run1.append(new_text1)
                
                # 如果有剩余文本，则添加到第二个新run
                if split_index_in_text < len(text.text):
                    new_text2 = ET.Element('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}t')
                    new_text2.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
                    new_text2.text = text.text[split_index_in_text:] if text.text else ''
                    new_run2.append(new_text2)
                    break
            else:
                # 在当前文本节点之前
                new_run1.append(text)
                continue
        for text in texts[i + 1:]:
            # 将剩余的文本节点添加到第二个新run
            new_run2.append(text)

        # 如果第一个新run没有文本，则插入空text
        if not new_run1.findall('.//w:t', NAMESPACES):
            empty_text = ET.Element('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}t')
            empty_text.text = ''
            new_run1.append(empty_text)

        # 如果第二个新run没有文本，则将其删除
        if not new_run2.findall('.//w:t', NAMESPACES):
            empty_text = ET.Element('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}t')
            empty_text.text = ''
            new_run2.append(empty_text)
        return new_run1, new_run2
    
    def _get_run_text(self, run):
        """
        获取指定run的全部文本内容
        :param run: 要获取文本的run元素
        :return: run中的全部文本内容
        """
        texts = run.findall('.//w:t', NAMESPACES)
        if not texts:
            # self.logger.error("指定的run没有文本节点")
            return ''
        
        text_content = ''.join([text.text for text in texts if text.text is not None])
        return text_content
    
    def _refine_paragraph(self, paragraph):
        """
        对段落进行精细化处理，移除空白字符和多余的文本节点
        :param paragraph: 段落元素
        """
        # 移除空的文本节点
        runs = paragraph.findall('.//w:r', NAMESPACES)
        for run in runs:
            texts = run.findall('.//w:t', NAMESPACES)
            for text in texts:
                if text.text is None or text.text == '':
                    run.remove(text)
        # 移除空的run
        for run in runs:
            children = list(run)
            prs = run.find('.//w:rPr', NAMESPACES)
            if not children or (len(children) == 1 and prs is not None):
                self._delete_element_from_root(paragraph, run)

    def _get_ins_ids(self):
        """
        获取所有插入（ins）操作的ID
        :return: 插入操作的ID列表
        """
        tree = ET.parse(self.document_path)
        root = tree.getroot()
        ins_ids = []
        for ins in root.findall('.//w:ins', NAMESPACES):
            ins_id = ins.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}id')
            if ins_id:
                ins_ids.append(int(ins_id))
        return ins_ids
    
    def _get_del_ids(self):
        """
        获取所有删除（del）操作的ID
        :return: 删除操作的ID列表
        """
        tree = ET.parse(self.document_path)
        root = tree.getroot()
        del_ids = []
        for delete in root.findall('.//w:del', NAMESPACES):
            del_id = delete.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}id')
            if del_id:
                del_ids.append(int(del_id))
        return del_ids
    
    def _get_pPrChange_ids(self):
        """
        获取所有段落属性更改（pPrChange）的ID
        :return: 段落属性更改的ID列表
        """
        tree = ET.parse(self.document_path)
        root = tree.getroot()
        pPrChange_ids = []
        for pPrChange in root.findall('.//w:pPrChange', NAMESPACES):
            pPrChange_id = pPrChange.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}id')
            if pPrChange_id:
                pPrChange_ids.append(int(pPrChange_id))
        return pPrChange_ids
    
    def _get_rPrChange_ids(self):
        """
        获取所有运行属性更改（rPrChange）的ID
        :return: 运行属性更改的ID列表
        """
        tree = ET.parse(self.document_path)
        root = tree.getroot()
        rPrChange_ids = []
        for rPrChange in root.findall('.//w:rPrChange', NAMESPACES):
            rPrChange_id = rPrChange.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}id')
            if rPrChange_id:
                rPrChange_ids.append(int(rPrChange_id))
        return rPrChange_ids

    def _get_main_paragraphs(self, root) -> list[ET.Element]:
        """获取文档主体的第一级段落，排除表格等嵌套结构中的段落"""
        body = root.find('.//w:body', NAMESPACES)
        if body is None:
            self.logger.debug("未找到文档主体")
            return []
        
        paragraphs = []
        for child in body:
            if child.tag == '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}p' and self.get_paragraph_text(child) != '':
                paragraphs.append(child)
            elif child.tag == '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}tbl':
                # 如果需要处理表格中的段落，可以在这里添加逻辑
                # table_paragraphs = child.findall('.//w:p', NAMESPACES)
                # paragraphs.extend(table_paragraphs)
                pass
        
        return paragraphs
    
    def get_paragraph_text(self, paragraph):
        """
        获取指定段落的文本内容
        :param paragraph: 段落元素
        :return: 段落中的全部文本内容
        """
        texts = paragraph.findall('.//w:t', NAMESPACES)
        if not texts:
            # self.logger.error("指定的段落没有文本节点")
            return ''
        # 过滤掉空文本节点
        run_texts = [text.text for text in texts if text.text is not None]
        paragraph_text = ''.join(run_texts)
        return paragraph_text

    def show_paragraphs(self):
        """
        打印文档中的所有段落及其索引
        """
        tree = ET.parse(self.document_path)
        root = tree.getroot()
        paragraphs = self._get_main_paragraphs(root)
        
        if not paragraphs:
            self.logger.debug("文档中没有段落")
            return
        
        for i, paragraph in enumerate(paragraphs):
            paragraph_text = self.get_paragraph_text(paragraph)
            if paragraph_text:
                self.logger.debug(f"段落 {i}: {paragraph_text}")
            else:
                self.logger.debug(f"段落 {i} 是空的")
        
        self.logger.debug("已显示所有段落")
    
    def _delete_element_from_root_recursive(self, parent_element, target_element) -> bool:
        """
        从父元素中递归删除指定的目标元素
        :param parent_element: 父元素
        :param target_element: 要删除的目标元素
        :return: 是否成功删除
        """
        for i, child in enumerate(parent_element):
            if child == target_element:
                parent_element.remove(child)
                return True  # 成功删除
            elif isinstance(child, ET.Element):
                # 递归检查子元素
                res = self._delete_element_from_root_recursive(child, target_element)
                if res:
                    return True
        return False  # 未找到目标元素

    def _delete_element_from_root(self, parent_element, target_element):
        """        
        从父元素中删除指定的目标元素
        :param parent_element: 父元素
        :param target_element: 要删除的目标元素
        """
        if not isinstance(parent_element, ET.Element):
            self.logger.debug("parent_element 必须是一个 Element 对象")
            return
        if not isinstance(target_element, ET.Element):
            self.logger.debug("target_element 必须是一个 Element 对象")
        res = self._delete_element_from_root_recursive(parent_element, target_element)
        if not res:
            self.logger.debug("未找到目标元素，无法删除")

    def _get_child_with_text_in_effect_from_paragraph(self, paragraph):
        """
        获取段落中具有文本内容且生效的子元素
        :param paragraph: 段落元素
        :return: 具有文本内容的子元素列表
        """
        children_with_text = []
        for child in paragraph:
            if child.tag == '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}r':
                children_with_text.append(child)
            elif child.tag == '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}ins':
                run = child.find('.//w:r', NAMESPACES)
                if run is not None:
                    children_with_text.append(run)
        return children_with_text

    def _get_parent_of_element(self, root, element):
        """
        获取指定元素的父元素
        :param root: XML树的根元素
        :param element: 要查找父元素的子元素
        :return: 父元素，如果未找到则返回None
        """
        if root is None or element is None:
            return None
        if element == root:
            return None  # 如果元素就是根元素，则没有父元素
        
        parent = root

        # 遍历子元素，查找目标元素的父元素
        children = list(parent)
        for child in children:
            if child == element:
                return parent
            elif isinstance(child, ET.Element):
                # 递归查找子元素的父元素
                found_parent = self._get_parent_of_element(child, element)
                if found_parent is not None:
                    return found_parent
        return None  # 如果未找到父元素，则返回None
    
    def _is_contained_in_run(self, root, run):
        """
        检查指定元素是否包含在任意一个run元素中
        :param root: XML树的根元素
        :param run: 要检查的run元素
        :return: 如果run包含在root中，则返回True，否则返回False
        """
        if root is None or run is None:
            return False
        if not isinstance(root, ET.Element) or not isinstance(run, ET.Element):
            self.logger.debug("root和run必须是Element对象")
            return False
        
        # 首先检查run是否在root中存在
        if not self._is_element_in_tree(root, run):
            self.logger.debug("指定的run元素不在root树中")
            return False
        
        # 从run开始向上查找父节点，检查是否有run元素
        current_element = run
        parent = self._get_parent_of_element(root, current_element)
        
        while parent is not None:
            # 检查当前父节点是否是run元素
            if parent.tag == '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}r':
                return True
            
            # 继续向上查找
            current_element = parent
            parent = self._get_parent_of_element(root, current_element)
        
        # 如果遍历完所有父节点都没有找到run元素，说明不被包含在run中
        return False

    def _is_element_in_tree(self, root, target_element):
        """
        检查目标元素是否存在于XML树中
        :param root: XML树的根元素
        :param target_element: 要查找的目标元素
        :return: 如果找到则返回True，否则返回False
        """
        if root == target_element:
            return True
        
        # 递归检查所有子元素
        for child in root:
            if isinstance(child, ET.Element):
                if child == target_element:
                    return True
                # 递归检查子元素的子树
                if self._is_element_in_tree(child, target_element):
                    return True
        
        return False
            
    def _elements_equal(self, e1, e2):
        """        
        检查两个XML元素是否相等
        :param e1: 第一个元素
        :param e2: 第二个元素 
        :return: 如果两个元素相等则返回True，否则返回False
        """
        if e1.tag != e2.tag:
            return False
        if e1.attrib != e2.attrib:
            return False
        if (e1.text or '').strip() != (e2.text or '').strip():
            return False
        if len(e1) != len(e2):
            return False
        return all(self._elements_equal(c1, c2) for c1, c2 in zip(e1, e2))
    
    def _element_matches_template(self, element, template):
        """
        检查元素是否与模板匹配
        :param element: 要检查的元素
        :param template: 模板元素
        :return: 如果元素与模板匹配则返回True，否则返回False
        """
        if element.tag != template.tag:
            return False
        for key, value in template.attrib.items():
            if element.attrib.get(key) != value:
                return False
        template_children = list(template)
        element_children = list(element)
        for t_child in template_children:
            match_found = False
            for e_child in element_children:
                if self._element_matches_template(e_child, t_child):
                    match_found = True
                    break
            if not match_found:
                return False
        return True
    
    def _get_chinese_number(self, index):
        """index从1开始"""
        chinese_numbers = [
            "零", "一", "二", "三", "四", "五", "六", "七", "八", "九", "十",
            "十一", "十二", "十三", "十四", "十五", "十六", "十七", "十八", "十九", "二十",
            # 可继续扩展
        ]
        if 1 <= index <= len(chinese_numbers):
            return chinese_numbers[index]
        else:
            # 超出范围可自定义处理
            return str(index)
        
    def _get_lower_letter(self, index):
        """获取小写字母表示的序号"""
        if 1 <= index <= 26:
            return chr(96 + index)  # 96是小写字母'a'的ASCII码-1
        else:
            # 超出范围可自定义处理
            return str(index)

    def _get_lower_roman(self, index):
        """
        获取小写罗马数字表示的序号（如 i, ii, iii, iv, ...）
        :param index: 从1开始的整数
        :return: 小写罗马数字字符串
        """
        romans = [
            '', 'i', 'ii', 'iii', 'iv', 'v', 'vi', 'vii', 'viii', 'ix', 'x',
            'xi', 'xii', 'xiii', 'xiv', 'xv', 'xvi', 'xvii', 'xviii', 'xix', 'xx',
            'xxi', 'xxii', 'xxiii', 'xxiv', 'xxv', 'xxvi', 'xxvii', 'xxviii', 'xxix', 'xxx'
        ]
        if 1 <= index < len(romans):
            return romans[index]
        else:
            # 超出范围时可自定义处理
            return str(index)

    def _get_decimal_enclosed_circle_chinese(self, index):
        """获取带圈的中文数字表示的序号"""
        circled_chinese_numbers = [
            "①", "②", "③", "④", "⑤", "⑥", "⑦", "⑧", "⑨", "⑩",
            "⑪", "⑫", "⑬", "⑭", "⑮", "⑯", "⑰", "⑱", "⑲", "⑳",
            # 可继续扩展
        ]
        if 1 <= index <= len(circled_chinese_numbers):
            return circled_chinese_numbers[index - 1]
        else:
            # 超出范围可自定义处理
            return str(index)

    def get_paragraphs(self):
        """
        获取文档中的所有段落
        :return: 段落列表
        """
        paragraph_texts = []
        tree = ET.parse(self.document_path)
        root = tree.getroot()
        paragraphs = self._get_main_paragraphs(root)
        for paragraph in paragraphs:
            paragraph_text = self.get_paragraph_text(paragraph)
            if paragraph_text:
                paragraph_texts.append(paragraph_text)

        return paragraph_texts

    def get_paragraphs_with_number(self):
        """
        获取带编号的段落
        :return: 带编号的段落列表
        """
        paragraph_texts = []
        tree = ET.parse(self.document_path)
        root = tree.getroot()
        paragraphs = self._get_main_paragraphs(root)
        for paragraph in paragraphs:
            paragraph_text = self.get_paragraph_text(paragraph)
            number_info = self._get_numbering(paragraph)
            if paragraph_text:
                if number_info:
                    paragraph_texts.append(f"{number_info} {paragraph_text}")
                else:
                    paragraph_texts.append(paragraph_text)

        return paragraph_texts

    def _get_numbering(self, paragraph):
        """
        获取段落的编号信息。
        """
        tag_text = ''
        ppr = paragraph.find('.//w:pPr', NAMESPACES)
        numPr = ppr.find('.//w:numPr', NAMESPACES) if ppr is not None else None
        numId = numPr.find('.//w:numId', NAMESPACES) if numPr is not None else None
        numId = numId.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val') if numId is not None else None
        ilvl = numPr.find('.//w:ilvl', NAMESPACES) if numPr is not None else None
        ilvl = ilvl.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val') if ilvl is not None else None
        if ilvl is None or numId is None:
            return tag_text  # 如果没有编号信息，返回空字符串

        numbering_info = self.numbers.get(numId, {}).get(ilvl)
        
        if numbering_info:
            numFmt, lvlText, _ = numbering_info
            if numFmt == "chineseCounting" or numFmt == "chineseCountingThousand" or numFmt == "japaneseCounting":
                tag_text = lvlText.replace(f'%{int(ilvl) + 1}', self._get_chinese_number(self.cur_number_count[numId][ilvl]))
                self.cur_number_count[numId][ilvl] += 1
            elif numFmt == "decimal":
                tag_text = lvlText.replace(f'%{int(ilvl) + 1}', str(self.cur_number_count[numId][ilvl]))
                self.cur_number_count[numId][ilvl] += 1
            elif numFmt == "lowerLetter":
                tag_text = lvlText.replace(f'%{int(ilvl) + 1}', self._get_lower_letter(self.cur_number_count[numId][ilvl]))
                self.cur_number_count[numId][ilvl] += 1
            elif numFmt == "lowerRoman":
                tag_text = lvlText.replace(f'%{int(ilvl) + 1}', self._get_lower_roman(self.cur_number_count[numId][ilvl]))
                self.cur_number_count[numId][ilvl] += 1
            elif numFmt == "decimalEnclosedCircleChinese":
                tag_text = lvlText.replace(f'%{int(ilvl) + 1}', self._get_decimal_enclosed_circle_chinese(self.cur_number_count[numId][ilvl]))
                self.cur_number_count[numId][ilvl] += 1

        return tag_text
    
    def _update_properties(self, pr, template):
        """
        更新属性
        :param pr: 要更新属性元素
        :param template: 模板段落属性元素
        """
        if pr is None or template is None:
            self.logger.debug("pr和template不能为空")
            return
        
        for t_child in template:
            for e_child in list(pr):
                if e_child.tag == t_child.tag:
                    self._delete_element_from_root(pr, e_child)  # 删除现有的样式属性
                    break

            pr.append(t_child)

        return pr

    def _remove_space_before_paragraph(self, paragraph):
        """
        移除段落前的非缩进空格
        :param paragraph: 段落元素
        """
        if paragraph is None:
            self.logger.debug("段落不能为空")
            return

        runs = paragraph.findall('.//w:r', NAMESPACES)
        for run in runs:
            if self._is_contained_in_run(paragraph, run):
                continue

            text = self._get_run_text(run)
            if text:
                new_text = text.lstrip(' ')
                if new_text != text:
                    t = run.find('.//w:t', NAMESPACES)
                    if t is not None:
                        t.text = new_text
                if new_text != '':
                    break

    def _setting_page_format(self, sectPr_template):
        """
        设置页面格式
        """
        self._refresh_document()

        # 查找所有的section元素
        sections_properties = self.document_root.findall('.//w:sectPr', NAMESPACES)
        for sectPr in sections_properties:
            sectPr = self._update_properties(sectPr, sectPr_template)

        self.document_tree.write(self.document_path, encoding='utf-8', xml_declaration=True)
        self.logger.debug("已更新页面格式")

    def _add_comment_range(self, paragraph_index: int, start: int, end: int, comment: Comment):
        """
        在指定段落范围内添加批注
        :param paragraph_index: 段落索引
        :param start: 开始字符索引
        :param end: 结束字符索引
        :param comment: Comment对象
        """
        comment_id = self._set_comment(comment)
        self._set_comment_relationship()

        # 查找paragraph_index对应的段落
        paragraphs = self._get_main_paragraphs(self.document_root)
        if paragraph_index < 0 or paragraph_index >= len(paragraphs):
            self.logger.debug(f"段落索引 {paragraph_index} 超出范围")
            return
        paragraph = paragraphs[paragraph_index]

        runs = paragraph.findall('.//w:r', NAMESPACES)
        if not runs:
            self.logger.debug(f"段落 {paragraph_index} 中没有运行（runs）")
            return
        
        # 根据start和end索引找到对应的文本节点
        l = 0
        start_run = None
        end_run = None
        comment_run_before = None
        comment_run_after = None
        start_split_index =0
        end_split_index = 0

        for i, run in enumerate(runs):
            if self._is_contained_in_run(paragraph, run):
                continue  # 如果run被包含在其他run中，则跳过
            text = self._get_run_text(run)
            # print(f"Run {i}: {text} (length: {len(text)})")
            l += len(text)
            if l >= start and start_run is None:
                start_run = run
                start_split_index = start - (l - len(text))
            if l >= end and end_run is None:
                end_run = run
                end_split_index = end - (l - len(text))
                break
        if start_run is None or end_run is None:
            self.logger.debug(f"在段落 {paragraph_index} 中未找到字符索引范围 {start}-{end}")

        if start_run == end_run:
            before_run, after_run = self._splite_run(start_run, start_split_index)
            comment_run_after, after_run = self._splite_run(after_run, end_split_index - start_split_index)
            self._insert_elements_after_element(paragraph, start_run, [before_run, comment_run_after, after_run])
            self._delete_element_from_root(paragraph, start_run)  # 移除原来的start_run
        else:
            before_run, comment_run_before = self._splite_run(start_run, start_split_index)
            comment_run_after, after_run = self._splite_run(end_run, end_split_index)

            # print(f"before_run:{self.get_run_text(before_run)}")
            # print(f"comment_run_before:{self.get_run_text(comment_run_before)}")
            # print(f"comment_run_after:{self.get_run_text(comment_run_after)}")
            # print(f"after_run:{self.get_run_text(after_run)}")
            
            self._insert_elements_after_element(paragraph, start_run, [before_run, comment_run_before])
            self._insert_elements_after_element(paragraph, end_run, [comment_run_after, after_run])

            self._delete_element_from_root(paragraph, start_run)  # 移除原来的start_run
            self._delete_element_from_root(paragraph, end_run)  # 移除原来的end_run

        # 创建批注范围标记元素
        comment_range_start = ET.Element('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}commentRangeStart')
        comment_range_start.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}id', comment_id)
        
        comment_range_end = ET.Element('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}commentRangeEnd')
        comment_range_end.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}id', comment_id)

        comment_reference_run = ET.Element('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}r')
        
        comment_reference = ET.SubElement(comment_reference_run, '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}commentReference')
        comment_reference.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}id', comment_id)
        
        # 将批注范围标记和批注引用添加到新的run中
        self._insert_elements_after_element(paragraph, before_run, [comment_range_start])
        self._insert_elements_after_element(paragraph, comment_run_after, [comment_range_end, comment_reference_run])

        # 对段落进行精细化处理，移除空白字符和多余的文本节点
        self._refine_paragraph(paragraph)

        self.logger.debug(f"已在段落 {paragraph_index} 的字符索引 {start}-{end} 添加批注: {comment.id}")

    def _insert(self, paragraph_index, insert_index, content, is_revising=False):
        """
        插入内容到文档中。
        :param paragraph_index: 段落索引。
        :param insert_index: 插入位置索引。
        :param content: 要插入的内容，可以是字符串或其他类型。
        :param is_revising: 是否处于修订模式。
        """
        paragraphs = self._get_main_paragraphs(self.document_root)
        if paragraph_index < 0 or paragraph_index >= len(paragraphs):
            self.logger.debug(f"段落索引 {paragraph_index} 超出范围")
            return
        paragraph = paragraphs[paragraph_index]

        # 查找段落中的所有运行（runs）
        runs = paragraph.findall('.//w:r', NAMESPACES)
        if not runs:
            self.logger.debug(f"段落 {paragraph_index} 中没有运行（runs）")
            return

        # 在指定位置插入内容
        l = 0
        inserted_run = None
        insert_index_in_run = 0.
        for i, run in enumerate(runs):
            if self._is_contained_in_run(paragraph, run):
                continue  # 如果run被包含在其他run中，则跳过
            text = self._get_run_text(run)
            l += len(text)
            if l >= insert_index:
                inserted_run = run
                insert_index_in_run = insert_index - (l - len(text))
                break
        
        if inserted_run is None:
            self.logger.debug(f"在段落 {paragraph_index} 中未找到插入位置 {insert_index}")
            return
        
        
        # 分裂指定的run
        before_run, after_run = self._splite_run(inserted_run, insert_index_in_run)

        # 获取run的属性
        insert_run = ET.Element('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}r')
        insert_run_properties = inserted_run.find('.//w:rPr', NAMESPACES)
        if insert_run_properties is not None:
            insert_run.append(insert_run_properties)

        # 创建新的文本节点
        new_text = ET.Element('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}t')
        new_text.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
        new_text.text = content if isinstance(content, str) else str(content)
        insert_run.append(new_text)

        # 如果开启修订模式，则添加修订标记
        if is_revising:
            ins_id = max(self._get_ins_ids()) + 1 if len(self._get_ins_ids()) > 0 else 1  # 生成新的修订ID
            revision_mark = ET.Element('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}ins')
            revision_mark.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}id', str(ins_id))
            revision_mark.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}author', self.author_name)
            revision_mark.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}date', datetime.now().isoformat())
            revision_mark.append(insert_run)

        else:
            # 如果不处于修订模式，则直接插入新的run
            revision_mark = insert_run

        # 在insert_run后添加三个run元素替换掉先前的insert_run
        self._insert_elements_after_element(paragraph, inserted_run, [before_run, revision_mark, after_run])

        self._delete_element_from_root(paragraph, inserted_run)  # 移除原来的insert_run
        self._refine_paragraph(paragraph)  # 对段落进行精细化处理

        self.logger.debug(f"已在段落 {paragraph_index} 的索引 {insert_index} 插入内容: {content}")   
        
    def _delete(self,paragraph_index, start, end, is_revising=False):
        """
        删除文档中的内容。
        :param paragraph_index: 段落索引。
        :param start: 开始字符索引。
        :param end: 结束字符索引。
        :param is_revising: 是否处于修订模式。
        """
        paragraphs = self._get_main_paragraphs(self.document_root)
        if paragraph_index < 0 or paragraph_index >= len(paragraphs):
            self.logger.debug(f"段落索引 {paragraph_index} 超出范围")
            return
        paragraph = paragraphs[paragraph_index]

        # 查找段落中的所有运行（runs）
        runs = paragraph.findall('.//w:r', NAMESPACES)
        if not runs:
            self.logger.debug(f"段落 {paragraph_index} 中没有运行（runs）")
            return
        
        # 根据start和end索引找到对应的文本节点
        l = 0
        start_run = None
        end_run = None
        delete_run_before = None
        delete_run_after = None
        start_split_index = 0
        end_split_index = 0
        for i, run in enumerate(runs):
            if self._is_contained_in_run(paragraph, run):
                continue  # 如果run被包含在其他run中，则跳过
            text = self._get_run_text(run)
            
            l += len(text)
            if l >= start and start_run is None:
                start_run = run
                start_split_index = start - (l - len(text))
            if l >= end and end_run is None:
                end_run = run
                end_split_index = end - (l - len(text))
                break
        if start_run is None or end_run is None:
            self.logger.debug(f"在段落 {paragraph_index} 中未找到删除范围 {start}-{end}")
            return

        if start_run == end_run:
            before_run, after_run = self._splite_run(start_run, start_split_index)
            delete_run_after, after_run = self._splite_run(after_run, end_split_index - start_split_index)
            self._insert_elements_after_element(paragraph, start_run, [before_run, delete_run_after, after_run])

            self._delete_element_from_root(paragraph, start_run)  # 移除原来的start_run

            # 如果开启修订模式，则添加修订标记
            if is_revising:
                del_id = max(self._get_del_ids()) + 1 if len(self._get_del_ids()) > 0 else 1  # 生成新的删除ID
                revision_mark = ET.Element('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}del')
                revision_mark.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}id', str(del_id))
                revision_mark.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}author', self.author_name)
                revision_mark.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}date', datetime.now().isoformat())
                texts = delete_run_after.findall('.//w:t', NAMESPACES)

                new_run = ET.Element('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}r')
                new_run_properties = delete_run_after.find('.//w:rPr', NAMESPACES)
                if new_run_properties is not None:
                    new_run.append(new_run_properties)

                if texts:
                    
                    # 如果run中有文本，则将文本添加到删除标记中
                    for text in texts:
                        del_text = ET.Element('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}delText')
                        del_text.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
                        del_text.text = text.text
                        new_run.append(del_text)  # 将文本添加到删除标记中
                revision_mark.append(new_run)

                self._insert_elements_after_element(paragraph, delete_run_after, [revision_mark])
                self._delete_element_from_root(paragraph, delete_run_after)  # 移除删除的run
            else:
                self._delete_element_from_root(paragraph, delete_run_after)  # 移除删除的run
        else:
            runs_to_delete = []
            s = 0
            e = 0
            # 找到start_run和end_run在段落中的索引位置

            for i, run in enumerate(runs):
                if run == start_run:
                    s = i
                if run == end_run:
                    e = i
            for i in range(s+1, e):
                runs_to_delete.append(runs[i])


            before_run, delete_run_before = self._splite_run(start_run, start_split_index)
            delete_run_after, after_run = self._splite_run(end_run, end_split_index)

            self._insert_elements_after_element(paragraph, start_run, [before_run, delete_run_before])
            self._insert_elements_after_element(paragraph, end_run, [delete_run_after, after_run])

            self._delete_element_from_root(paragraph, start_run)  # 移除原来的start_run
            self._delete_element_from_root(paragraph, end_run)  # 移除原来的end_run

            runs_to_delete.append(delete_run_before)
            runs_to_delete.append(delete_run_after)

            # 如果开启修订模式，则添加修订标记
            if is_revising:
                for run in runs_to_delete:
                    del_id = max(self._get_del_ids()) + 1 if len(self._get_del_ids()) > 0 else 1  # 生成新的删除ID
                    parent = self._get_parent_of_element(paragraph, run)
                    if parent is not None:
                        if parent.tag == '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}ins':
                            self._delete_element_from_root(paragraph, run)  # 如果是在ins元素中，则直接删除
                        else:
                            revision_mark = ET.Element('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}del')
                            revision_mark.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}id', str(del_id))
                            revision_mark.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}author', self.author_name)
                            revision_mark.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}date', datetime.now().isoformat())
                            new_run = ET.Element('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}r')
                            new_run_properties = delete_run_after.find('.//w:rPr', NAMESPACES)
                            if new_run_properties is not None:
                                new_run.append(new_run_properties)
                            texts = run.findall('.//w:t', NAMESPACES)
                            if texts:
                                
                                # 如果run中有文本，则将文本添加到删除标记中
                                for text in texts:
                                    del_text = ET.Element('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}delText')
                                    del_text.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
                                    del_text.text = text.text
                                    new_run.append(del_text) 
                            revision_mark.append(new_run)
                            self._insert_elements_after_element(paragraph, run, [revision_mark])
                            self._delete_element_from_root(paragraph, run)
                    else:
                        self.logger.debug(f"未找到run的父元素，无法添加删除标记: {self._get_run_text(run)}")
            else:
                for run in runs_to_delete:
                    self._delete_element_from_root(paragraph, run)  # 移除删除的run


        comment_range_starts = paragraph.findall('.//w:commentRangeStart', NAMESPACES)
        comment_range_ends = paragraph.findall('.//w:commentRangeEnd', NAMESPACES)
        if is_revising:
            for comment_range_start in comment_range_starts:
                parent = self._get_parent_of_element(paragraph, comment_range_start)
                if parent is not None:
                    next = self._get_element_position_in_parent(parent, comment_range_start) + 1
                    if next < len(parent):
                        next_element = parent[next]
                        if next_element.tag == '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}del':
                            next_element.insert(0, comment_range_start)  # 将commentRangeStart插入到del元素中
                            self._delete_element_from_root(paragraph, comment_range_start)  # 删除原来的commentRangeStart
            for comment_range_end in comment_range_ends:
                parent = self._get_parent_of_element(paragraph, comment_range_end)
                if parent is not None:
                    prev = self._get_element_position_in_parent(parent, comment_range_end) - 1
                    if prev >= 0:
                        prev_element = parent[prev]
                        if prev_element.tag == '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}del':
                            prev_element.append(comment_range_end) # 将commentRangeEnd添加到del元素中
                            self._delete_element_from_root(paragraph, comment_range_end)  # 删除原来的commentRangeEnd
        else:
            for comment_range_start in comment_range_starts:
                comment_id = comment_range_start.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}id')
                if comment_id is not None:
                    comment_range_end = paragraph.find(f".//w:commentRangeEnd[@w:id='{comment_id}']", NAMESPACES)
                    if comment_range_end is not None:
                        parent = self._get_parent_of_element(paragraph, comment_range_start)
                        if parent is not None:
                            index_start = self._get_element_position_in_parent(parent, comment_range_start)
                            index_end = self._get_element_position_in_parent(parent, comment_range_end)
                            if index_start != -1 and index_end != -1:
                                if index_start == index_end - 1:
                                    # 如果commentRangeStart和commentRangeEnd是相邻的，则直接删除
                                    parent.remove(comment_range_start)
                                    parent.remove(comment_range_end)
                            self.logger.debug(f"已删除批注范围标记: {comment_id}")

        # 对段落进行精细化处理，移除空白字符和多余的文本节点
        self._refine_paragraph(paragraph)
        self.logger.debug(f"已在段落 {paragraph_index} 的字符索引 {start}-{end} 删除内容")
    
    def _refresh_document(self):
        """
        更新文档树和根节点
        """
        self.document_tree = ET.parse(self.document_path)
        self.document_root = self.document_tree.getroot()

    def _save_document(self):
        """
        保存文档
        """
        try:
            self.document_tree.write(self.document_path, encoding='utf-8', xml_declaration=True)
        except Exception as e:
            self.logger.error(f"保存文档失败: {e}")
            return False
        return True

    def __del__(self):
        self._clear_extracted()
        abs_path = os.path.abspath(self.docx_path)
        if abs_path in self.__class__._instances:
            del self.__class__._instances[abs_path]
    
    # endregion

    def add_comment_range(self, paragraph_index: int, start: int, end: int, comment: Comment):
        """
        添加批注范围
        """
        self._refresh_document()
        try:
            # 增加范围批注
            self._add_comment_range(paragraph_index, start, end, comment)
        except Exception as e:
            self.logger.error(f"添加批注范围失败: {e}")
            return False

        # 保存修改后的document.xml
        return self._save_document()

    def insert(self, paragraph_index, insert_index, content, is_revising=False):
        """
        插入指定内容到段落中。
        """
        self._refresh_document()

        try:
            # 插入内容
            self._insert(paragraph_index, insert_index, content, is_revising)
        except Exception as e:
            self.logger.error(f"插入内容失败: {e}")
            return False

        # 保存修改后的document.xml
        return self._save_document()

    def delete(self, paragraph_index, start, end, is_revising=False):
        """
        删除指定范围的内容。
        """
        self._refresh_document()

        try:
            # 删除指定范围的内容
            self._delete(paragraph_index, start, end, is_revising)
        except Exception as e:
            self.logger.error(f"删除内容失败: {e}")
            return False

        # 保存修改后的document.xml
        return self._save_document()

    def _apply_style_body(self, paragraph_index: int, template_style: ParaStyleProperties, is_revising=False):
        """
        应用正文样式
        """
        paragraphs = self._get_main_paragraphs(self.document_root)
        if paragraph_index < 0 or paragraph_index >= len(paragraphs):
            self.logger.debug(f"段落索引 {paragraph_index} 超出范围")
            return None
        paragraph = paragraphs[paragraph_index]

        self._remove_space_before_paragraph(paragraph)

        pPr = paragraph.find('.//w:pPr', NAMESPACES)
        template = template_style.to_xml_element()

        pPr_copy = copy.deepcopy(pPr)   # 复制现有的样式属性
        for child in pPr_copy:
            if child.tag == '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}rPr':
                # 如果存在运行属性，则将其删除
                pPr_copy.remove(child)
                break

        # 调整段落属性
        if pPr:
            pPr = self._update_properties(pPr, template)
            if is_revising and pPr.find('.//w:pPrChange', NAMESPACES) is None:
                change_id = max(self._get_pPrChange_ids()) + 1 if len(self._get_pPrChange_ids()) > 0 else 1  # 生成新的删除ID
                pPr_change = ET.Element('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}pPrChange')
                pPr_change.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}id', str(change_id))
                pPr_change.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}author', self.author_name)
                pPr_change.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}date', datetime.now().isoformat())
                pPr_change.append(pPr_copy)
                pPr.append(pPr_change)  # 将修订标记添加到段落属性中
        else:
            paragraph.insert(0, template)  # 如果没有样式属性，则直接插入样式属性


        template_run = template_style.rpr.to_xml_element()
        runs = paragraph.findall('.//w:r', NAMESPACES)
        if not runs:
            self.logger.debug(f"段落 {paragraph_index} 中没有运行（runs）")
            return None
        for run in runs:
            if self._is_contained_in_run(paragraph, run):
                continue
            text = self._get_run_text(run)
            if text == "":
                continue

            rpr = run.find('.//w:rPr', NAMESPACES)
            if rpr:
                rpr_copy = copy.deepcopy(rpr)  # 复制现有的样式属性
                rpr = self._update_properties(rpr, template_run)
                if is_revising and rpr.find('.//w:rPrChange', NAMESPACES) is None:
                    change_id = max(self._get_rPrChange_ids()) + 1 if len(self._get_rPrChange_ids()) > 0 else 1  # 生成新的删除ID
                    rPrChange = ET.Element('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}rPrChange')
                    rPrChange.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}id', str(change_id))
                    rPrChange.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}author', self.author_name)
                    rPrChange.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}date', datetime.now().isoformat())
                    rPrChange.append(rpr_copy)
                    rpr.append(rPrChange)  # 将修订标记添加到运行属性中
            else:
                run.insert(0, template_run)  # 如果没有样式属性，则直接插入样式属性

        return ParaStyleProperties.load_from_xml(pPr)

    def _apply_style_heading(self, paragraph_index: int, template_style: ParaStyleProperties, is_revising=False):
        """
        应用标题样式。
        :param paragraph_index: 段落索引。
        :param template_style: 模板样式。
        :param is_revising: 是否处于修订模式。
        """
        paragraphs = self._get_main_paragraphs(self.document_root)
        if paragraph_index < 0 or paragraph_index >= len(paragraphs):
            self.logger.debug(f"段落索引 {paragraph_index} 超出范围")
            return None
        paragraph = paragraphs[paragraph_index]

        template = template_style.to_xml_element()

        pPr = paragraph.find('.//w:pPr', NAMESPACES)

        numPr = template.find('.//w:numPr', NAMESPACES)
        if numPr is not None:
            # 拥有heading属性，说明是标题样式,不再使用缩进
            template.remove(ind := template.find('.//w:ind', NAMESPACES)) if ind is not None else None
        else:
            self._remove_space_before_paragraph(paragraph)

        pPr_copy = copy.deepcopy(pPr)   # 复制现有的样式属性
        for child in pPr_copy:
            if child.tag == '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}rPr':
                # 如果存在运行属性，则将其删除
                pPr_copy.remove(child)
                break
        if pPr:
            for t_child in template:
                for e_child in list(pPr):
                    if e_child.tag == t_child.tag:
                        self._delete_element_from_root(pPr, e_child)  # 删除现有的样式属性
                        break

                pPr.append(t_child)
            if is_revising and pPr.find('.//w:pPrChange', NAMESPACES) is None:
                change_id = max(self._get_pPrChange_ids()) + 1 if len(self._get_pPrChange_ids()) > 0 else 1  # 生成新的删除ID
                pPr_change = ET.Element('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}pPrChange')
                pPr_change.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}id', str(change_id))
                pPr_change.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}author', self.author_name)
                pPr_change.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}date', datetime.now().isoformat())
                pPr_change.append(pPr_copy)
                pPr.append(pPr_change)  # 将修订标记添加到段落属性中
        else:
            paragraph.insert(0, template)  # 如果没有样式属性，则直接插入样式属性


        template_run = template_style.rpr.to_xml_element()
        runs = paragraph.findall('.//w:r', NAMESPACES)
        if not runs:
            self.logger.debug(f"段落 {paragraph_index} 中没有运行（runs）")
            return None
        for run in runs:
            if self._is_contained_in_run(paragraph, run):
                continue
            text = self._get_run_text(run)
            if text == "":
                continue

            rpr = run.find('.//w:rPr', NAMESPACES)
            if rpr:
                rpr_copy = copy.deepcopy(rpr)  # 复制现有的样式属性
                for t_child in template_run:
                    for e_child in list(rpr):
                        if e_child.tag == t_child.tag:
                            self._delete_element_from_root(rpr, e_child)  # 删除现有的样式属性
                            break
                    rpr.append(t_child)
                if is_revising and rpr.find('.//w:rPrChange', NAMESPACES) is None:
                    change_id = max(self._get_rPrChange_ids()) + 1 if len(self._get_rPrChange_ids()) > 0 else 1  # 生成新的删除ID
                    rPrChange = ET.Element('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}rPrChange')
                    rPrChange.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}id', str(change_id))
                    rPrChange.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}author', self.author_name)
                    rPrChange.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}date', datetime.now().isoformat())
                    rPrChange.append(rpr_copy)
                    rpr.append(rPrChange)  # 将修订标记添加到运行属性中
            else:
                run.insert(0, template_run)  # 如果没有样式属性，则直接插入样式属性

        return ParaStyleProperties.load_from_xml(pPr)

    def _apply_style_default(self, paragraph_index: int, template_style: ParaStyleProperties, is_revising=False):
        """
        应用默认段落样式。
        :param paragraph_index: 段落索引。
        :param template_style: 模板样式。
        :param is_revising: 是否处于修订模式。
        """
        paragraphs = self._get_main_paragraphs(self.document_root)
        if paragraph_index < 0 or paragraph_index >= len(paragraphs):
            self.logger.debug(f"段落索引 {paragraph_index} 超出范围")
            return None
        paragraph = paragraphs[paragraph_index]

        template = template_style.to_xml_element()

        pPr = paragraph.find('.//w:pPr', NAMESPACES)
            
        pPr_copy = copy.deepcopy(pPr)   # 复制现有的样式属性
        for child in pPr_copy:
            if child.tag == '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}rPr':
                # 如果存在运行属性，则将其删除
                pPr_copy.remove(child)
                break
        if pPr:
            for t_child in template:
                for e_child in list(pPr):
                    if e_child.tag == t_child.tag:
                        self._delete_element_from_root(pPr, e_child)  # 删除现有的样式属性
                        break

                pPr.append(t_child)
            if is_revising and pPr.find('.//w:pPrChange', NAMESPACES) is None:
                change_id = max(self._get_pPrChange_ids()) + 1 if len(self._get_pPrChange_ids()) > 0 else 1  # 生成新的删除ID
                pPr_change = ET.Element('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}pPrChange')
                pPr_change.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}id', str(change_id))
                pPr_change.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}author', self.author_name)
                pPr_change.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}date', datetime.now().isoformat())
                pPr_change.append(pPr_copy)
                pPr.append(pPr_change)  # 将修订标记添加到段落属性中
        else:
            paragraph.insert(0, template)  # 如果没有样式属性，则直接插入样式属性


        template_run = template_style.rpr.to_xml_element()
        runs = paragraph.findall('.//w:r', NAMESPACES)
        if not runs:
            self.logger.debug(f"段落 {paragraph_index} 中没有运行（runs）")
            return None
        for run in runs:
            if self._is_contained_in_run(paragraph, run):
                continue
            text = self._get_run_text(run)
            if text == "":
                continue

            rpr = run.find('.//w:rPr', NAMESPACES)
            if rpr:
                rpr_copy = copy.deepcopy(rpr)  # 复制现有的样式属性
                for t_child in template_run:
                    for e_child in list(rpr):
                        if e_child.tag == t_child.tag:
                            self._delete_element_from_root(rpr, e_child)  # 删除现有的样式属性
                            break
                    rpr.append(t_child)
                if is_revising and rpr.find('.//w:rPrChange', NAMESPACES) is None:
                    change_id = max(self._get_rPrChange_ids()) + 1 if len(self._get_rPrChange_ids()) > 0 else 1  # 生成新的删除ID
                    rPrChange = ET.Element('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}rPrChange')
                    rPrChange.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}id', str(change_id))
                    rPrChange.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}author', self.author_name)
                    rPrChange.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}date', datetime.now().isoformat())
                    rPrChange.append(rpr_copy)
                    rpr.append(rPrChange)  # 将修订标记添加到运行属性中
            else:
                run.insert(0, template_run)  # 如果没有样式属性，则直接插入样式属性

        return ParaStyleProperties.load_from_xml(pPr)

    def _apply_style(self, paragraph_index, template_style: ParaStyleProperties, is_revising=False):
        """
        修改段落样式。
        :param paragraph_index: 段落索引。
        :param style: 要应用的样式名称。
        :param is_revising: 是否处于修订模式。
        """
        apply_method_map = {
            Styles.BODY: self._apply_style_body,
            Styles.TITLE: self._apply_style_default,
            Styles.SUBTITLE_LEFT: self._apply_style_heading,
            Styles.SUBTITLE_CENTER: self._apply_style_heading,
            Styles.SUBSUBTITLE_CENTER: self._apply_style_heading,
            Styles.SUBSUBTITLE_LEFT: self._apply_style_heading,
            Styles.SUBSUBSUBTITLE_LEFT: self._apply_style_heading,
            Styles.SPECIFIC_RIGHT: self._apply_style_default,
            Styles.APPELLATION: self._apply_style_body,
        }
        try:
            apply_method = apply_method_map.get(template_style.name, self._apply_style_default)
            result = apply_method(paragraph_index, template_style, is_revising)
        except Exception as e:
            self.logger.error(f"应用样式失败: {e}")
            return None
        return result

    def apply_style(self, paragraph_index, template_style: ParaStyleProperties, is_revising=False):
        self._refresh_document()

        # 应用样式
        result = self._apply_style(paragraph_index, template_style, is_revising)
        if result is not None:
            return self._save_document()
        else:
            return False

    def _check_style_all(self, paragraph_index: int, template_style: ParaStyleProperties) -> bool:
        """
        检查段落是否具有指定样式。
        :param paragraph_index: 段落索引。
        :param style: 要检查的样式名称。
        :return: 如果段落具有指定样式，则返回True，否则返回False。
        """
        # 读取document.xml
        paragraphs = self._get_main_paragraphs(self.document_root)
        if paragraph_index < 0 or paragraph_index >= len(paragraphs):
            self.logger.debug(f"段落索引 {paragraph_index} 超出范围")
            return False
        paragraph = paragraphs[paragraph_index]
        pPr = paragraph.find('.//w:pPr', NAMESPACES)
        if pPr is None:
            self.logger.debug(f"段落 {paragraph_index} 没有样式属性")
            return False
        
        template = template_style.to_xml_element()

        return self._element_matches_template(pPr, template)

    def _check_style_body(self, paragraph_index: int, template_style: ParaStyleProperties) -> bool:
        """
        检查段落内容是否具有指定样式。
        :param paragraph_index: 段落索引。
        :param template_style: 要检查的样式对象。
        :return: 如果段落内容具有指定样式，则返回True，否则返回False。
        """
        # 检查段落的文本内容
        paragraphs = self._get_main_paragraphs(self.document_root)
        if paragraph_index < 0 or paragraph_index >= len(paragraphs):
            self.logger.debug(f"段落索引 {paragraph_index} 超出范围")
            return False
        paragraph = paragraphs[paragraph_index]
        pPr = paragraph.find('.//w:pPr', NAMESPACES)
        if pPr is None:
            self.logger.debug(f"段落 {paragraph_index} 没有样式属性")
            return False
    
        # 段落属性检查
        ind = pPr.find('.//w:ind', NAMESPACES)
        if ind is not None:
            firstLine = ind.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}firstLine')
            firstLine_template = template_style.first_line_indent
            if firstLine is not None and int(firstLine) != firstLine_template:
                self.logger.debug(f"段落 {paragraph_index} 的首行缩进不匹配")
                return False
            firstLineChars = ind.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}firstLineChars')
            firstLineChars_template = template_style.first_line_chars
            if firstLineChars is not None and int(firstLineChars) != firstLineChars_template:
                self.logger.debug(f"段落 {paragraph_index} 的首行字符数不匹配")
                return False
        else:
            self.logger.debug(f"段落 {paragraph_index} 没有缩进属性")
            return False

        rPr = pPr.find('.//w:rPr', NAMESPACES)
        if rPr is None:
            self.logger.debug(f"段落 {paragraph_index} 没有文本运行属性")
            return False

        # 文本运行属性检查
        def check_run_properties(rPr: ET.Element, template_style: ParaStyleProperties) -> bool:
            rFont = rPr.find('.//w:rFonts', NAMESPACES)
            if rFont is not None:
                font = rFont.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}eastAsia')
                font_template = template_style.rpr.font.eastAsia
                if font != font_template:
                    self.logger.debug(f"段落 {paragraph_index} 的字体不匹配")
                    return False
            else:
                self.logger.debug(f"段落 {paragraph_index} 没有字体属性")
                return False

            sz = rPr.find('.//w:sz', NAMESPACES)
            if sz is not None:
                size = sz.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val')
                size_template = template_style.rpr.size
                if size != None and int(size) != size_template:
                    self.logger.debug(f"段落 {paragraph_index} 的字号不匹配")
                    return False
            else:
                self.logger.debug(f"段落 {paragraph_index} 没有字号属性")
                return False

            return True

        if not check_run_properties(rPr, template_style):
            return False

        runs = paragraph.findall('.//w:r', NAMESPACES)
        for run in runs:
            if self._is_contained_in_run(paragraph, run):
                continue
            rPr = run.find('.//w:rPr', NAMESPACES)
            if rPr is None:
                self.logger.debug(f"段落 {paragraph_index} 的文本运行属性不存在")
                return False
            if not check_run_properties(rPr, template_style):
                return False

        return True

    def _check_style_default(self, paragraph_index: int, template_style: ParaStyleProperties) -> bool:
        """
        默认的段落样式检查
        :param paragraph_index: 段落索引。
        :param template_style: 模板样式属性。
        :return: 如果段落具有默认样式，则返回True，否则返回False。
        """
        def check_run_properties(rPr: ET.Element, template_style: ParaStyleProperties) -> bool:
            rFont = rPr.find('.//w:rFonts', NAMESPACES)
            if rFont is not None:
                font = rFont.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}eastAsia')
                font_template = template_style.rpr.font.eastAsia
                if font != font_template:
                    self.logger.debug(f"段落 {paragraph_index} 的字体不匹配")
                    return False
            else:
                self.logger.debug(f"段落 {paragraph_index} 没有字体属性")
                return False

            sz = rPr.find('.//w:sz', NAMESPACES)
            if sz is not None:
                size = sz.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val')
                size_template = template_style.rpr.size
                if size != None and int(size) != size_template:
                    self.logger.debug(f"段落 {paragraph_index} 的字号不匹配")
                    return False
            else:
                self.logger.debug(f"段落 {paragraph_index} 没有字号属性")
                return False

            return True

        # 读取document.xml
        paragraphs = self._get_main_paragraphs(self.document_root)
        if paragraph_index < 0 or paragraph_index >= len(paragraphs):
            self.logger.debug(f"段落索引 {paragraph_index} 超出范围")
            return False
        paragraph = paragraphs[paragraph_index]
        pPr = paragraph.find('.//w:pPr', NAMESPACES)
        if pPr is None:
            self.logger.debug(f"段落 {paragraph_index} 没有样式属性")
            return False

        rPr = pPr.find('.//w:rPr', NAMESPACES)
        if rPr is None:
            self.logger.debug(f"段落 {paragraph_index} 没有文本运行属性")
            return False

        if not check_run_properties(rPr, template_style):
            return False

        runs = paragraph.findall('.//w:r', NAMESPACES)
        for run in runs:
            if self._is_contained_in_run(paragraph, run):
                continue
            rPr = run.find('.//w:rPr', NAMESPACES)
            if rPr is None:
                self.logger.debug(f"段落 {paragraph_index} 的文本运行属性不存在")
                return False
            if not check_run_properties(rPr, template_style):
                return False

        return True

    def _check_style(self, paragraph_index: int, template_style: ParaStyleProperties) -> bool:
        check_method_map = {
            Styles.BODY: self._check_style_body,
            Styles.TITLE: self._check_style_default,
            Styles.SUBTITLE_LEFT: self._check_style_default,
            Styles.SUBTITLE_CENTER: self._check_style_default,
            Styles.SUBSUBTITLE_CENTER: self._check_style_default,
            Styles.SUBSUBTITLE_LEFT: self._check_style_default,
            Styles.SUBSUBSUBTITLE_LEFT: self._check_style_default,
            Styles.SPECIFIC_RIGHT: self._check_style_default,
            Styles.APPELLATION: self._check_style_default,
        }
        check_method = check_method_map.get(template_style.name, self._check_style_default)
        try:
            result = check_method(paragraph_index, template_style)
        except Exception as e:
            self.logger.error(f"检查段落 {paragraph_index} 样式时出错: {e}")
            return False
        return result

    def check_style(self, paragraph_index: int, template_style: ParaStyleProperties) -> bool:
        """
        检查段落是否具有指定样式。
        :param paragraph_index: 段落索引。
        :param style: 要检查的样式名称。
        :return: 如果段落具有指定样式，则返回True，否则返回False。
        """
        self._refresh_document()

        return self._check_style(paragraph_index, template_style)

    # region 任务处理模块
    def _process_task(self, task: Task):
        process_method_map = {
            TaskType.DELETE: self._process_delete_task,
            TaskType.INSERT: self._process_insert_task,
            TaskType.ADD_COMMENT: self._process_add_comment_task,
            TaskType.APPLY_STYLE: self._process_apply_style_task,
            TaskType.CHECK_STYLE: self._process_check_style_task,
        }
        process_method = process_method_map.get(task.type, None)
        if process_method:
            if process_method(task):
                task.result = True
            else:
                task.result = False

        return task.result

    def process_task(self, task: Task):
        self._refresh_document()
        self._process_task(task)
        self._save_document()
        return task

    def process_tasks(self, tasks: list[Task]) -> list[Task]:
        self._refresh_document()
        for task in tasks:
            self._process_task(task)
        self._save_document()
        return tasks

    def _process_delete_task(self, task: DeleteTask):
        return self._delete(task.paragraph_index, task.original_start, task.original_end, task.is_revising)
         
    def _process_insert_task(self, task: InsertTask):
        return self._insert(task.paragraph_index, task.original_start, task.revised_text, task.is_revising)

    def _process_add_comment_task(self, task: AddCommentTask):
        return self._add_comment_range(task.paragraph_index, task.original_start, task.original_end, task.comment)

    def _process_apply_style_task(self, task: ApplyStyleTask):
        result = self._apply_style(task.paragraph_index, task.template_style, task.is_revising)
        if result is not None:
            task.original_style = result
            return True
        else:
            return False

    def _process_check_style_task(self, task: CheckStyleTask):
        return self._check_style(task.paragraph_index, task.template_style)

    # endregion



if __name__ == "__main__":
    docx_path = './doc/郑州商品交易所因公出国（境）人员费用管理办法(错别字）_origin_白文档.docx'
    manager = DocxFileManager(docx_path)
    
    # Example usage
    # comment = Comment(id=1, text="This is a test comment", para_id="12345678")
    # manager.add_comment_range(paragraph_index=5, start=0, end=10, comment=comment)

    # manager.insert(paragraph_index=8, insert_index=0, content="Inserted text", is_revising=False)
    # manager.insert(paragraph_index=8, insert_index=5, content="Inserted text at index 5", is_revising=False)
    # manager.delete(paragraph_index=7, start=0, end=25, is_revising=False)
    manager.show_paragraphs()  # Display all paragraphs in the document

    manager.compress_docx()  # Compress the modified files back to docx format

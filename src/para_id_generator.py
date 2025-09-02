import os
import xml.etree.ElementTree as ET


from app.core.doc_inspection.utils.unique_id_generator import UniqueIDGenerator

class ParaIdGenerator:
    def __init__(self, docx_path):
        """
        Initialize the paragraph ID generator with existing IDs.
        """
        document_path = os.path.join(docx_path, 'word', 'document.xml')
        comments_path = os.path.join(docx_path, 'word', 'comments.xml')
        footnotes_path = os.path.join(docx_path, 'word', 'footnotes.xml')
        headers_path = os.path.join(docx_path, 'word', 'header1.xml')  # 可能有多个header
        ids = []
        # 读取 document.xml
        if os.path.exists(document_path):
            ids.extend(self._extract_para_ids(document_path))
            
        # 读取 comments.xml
        if os.path.exists(comments_path):
            ids.extend(self._extract_para_ids(comments_path))
            
        # 读取 footnotes.xml
        if os.path.exists(footnotes_path):
            ids.extend(self._extract_para_ids(footnotes_path))

        # 读取所有可能的header文件
        word_dir = os.path.join(docx_path, 'word')
        if os.path.exists(word_dir):
            for filename in os.listdir(word_dir):
                if filename.startswith('header') and filename.endswith('.xml'):
                    header_path = os.path.join(word_dir, filename)
                    ids.extend(self._extract_para_ids(header_path))
                elif filename.startswith('footer') and filename.endswith('.xml'):
                    footer_path = os.path.join(word_dir, filename)
                    ids.extend(self._extract_para_ids(footer_path))

        # 去重
        ids = list(set(ids))
        print(f"找到 {len(ids)} 个现有的段落ID")        
        
        self.generator = UniqueIDGenerator(ids)

    def _extract_para_ids(self, xml_file_path):
        """
        从XML文件中提取所有段落ID
        """
        para_ids = []
        try:
            # 定义命名空间
            namespaces = {
                'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
                'w14': 'http://schemas.microsoft.com/office/word/2010/wordml'
            }
            
            # 解析XML文件
            tree = ET.parse(xml_file_path)
            root = tree.getroot()
            
            # 查找所有的 <w:p> 元素
            paragraphs = root.findall('.//w:p', namespaces)
            
            for para in paragraphs:
                # 获取 w14:paraId 属性
                para_id = para.get('{http://schemas.microsoft.com/office/word/2010/wordml}paraId')
                if para_id:
                    para_ids.append(para_id)
                    
            print(f"从 {xml_file_path} 中提取到 {len(para_ids)} 个段落ID")
            
        except ET.ParseError as e:
            print(f"解析XML文件 {xml_file_path} 时出错: {e}")
        except Exception as e:
            print(f"读取文件 {xml_file_path} 时出错: {e}")
            
        return para_ids

    def generate_unique_id(self):
        """
        Generate a unique paragraph ID.
        """
        return self.generator.generate_unique_id()

    def reset(self):
        """
        Reset the paragraph ID generator.
        """
        self.generator.reset()

if __name__ == "__main__":
    # 测试提取段落ID
    docx_extract_path = './doc/extracted_docx'  # 解压后的docx目录
    
    if os.path.exists(docx_extract_path):
        generator = ParaIdGenerator(docx_extract_path)
        
        # 生成几个新的唯一ID
        print("\n生成新的唯一ID:")
        for i in range(5):
            new_id = generator.generate_unique_id()
            print(f"新ID {i+1}: {new_id}")
    else:
        print(f"目录不存在: {docx_extract_path}")
        print("请先解压docx文件")
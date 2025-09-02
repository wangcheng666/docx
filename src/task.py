

from dataclasses import dataclass
import enum

from app.core.doc_inspection.base.comment import Comment
from app.core.doc_inspection.base.styles import ParaStyleProperties

class TaskType(enum.Enum):
    DELETE = 'delete'
    INSERT = 'insert'
    ADD_COMMENT = 'add_comment'
    APPLY_STYLE = 'apply_style'
    CHECK_STYLE = 'check_style'

@dataclass
class Task:
    type: TaskType = None
    paragraph_index: int = None
    result: bool = False

class DeleteTask(Task):
    def __init__(self, paragraph_index: int, original_start: int, original_end: int, is_revising: bool=False):
        self.type = TaskType.DELETE
        self.paragraph_index = paragraph_index
        self.original_start = original_start
        self.original_end = original_end
        self.is_revising = is_revising
        


class InsertTask(Task):
    def __init__(self, paragraph_index: int, original_start: int, revised_text: str, is_revising: bool=False):
        self.type = TaskType.INSERT
        self.paragraph_index = paragraph_index
        self.original_start = original_start
        self.revised_text = revised_text
        self.is_revising = is_revising

class AddCommentTask(Task):
    def __init__(self, paragraph_index: int, original_start: int, original_end: int, comment: Comment):
        self.type = TaskType.ADD_COMMENT
        self.paragraph_index = paragraph_index
        self.original_start = original_start
        self.original_end = original_end
        self.comment: Comment = comment

class ApplyStyleTask(Task):
    def __init__(self, paragraph_index: int, template_style: ParaStyleProperties, is_revising: bool=False):
        self.type = TaskType.APPLY_STYLE
        self.paragraph_index = paragraph_index
        self.template_style: ParaStyleProperties = template_style
        self.original_style: ParaStyleProperties = None
        self.is_revising = is_revising

class CheckStyleTask(Task):
    def __init__(self, paragraph_index: int, template_style: ParaStyleProperties):
        self.type = TaskType.CHECK_STYLE
        self.paragraph_index = paragraph_index
        self.template_style: ParaStyleProperties = template_style

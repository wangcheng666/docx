    

from dataclasses import dataclass

@dataclass
class Font:
    ascii: str = ""
    hAnsi: str = ""
    eastAsia: str = ""
    hint: str = ""
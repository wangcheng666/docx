
class Docx:
    def __init__(self, xml: str):
        self._xml = xml
 

    @property
    def xml(self) -> str:
        return self._xml

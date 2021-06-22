import re

from obfuscator.modifier.base import Modifier
from obfuscator.msdocument import MSDocument


class RemoveIndentation(Modifier):
    def run(self, doc: MSDocument) -> None:
        doc.code = re.sub(r'^\s*', '', doc.code, flags=re.MULTILINE)


class RemoveEmptyLines(Modifier):
    def run(self, doc: MSDocument) -> None:
        doc.code = re.sub(r'(?:(\s*)\n)+', '\n', doc.code)

import logging
import random

from pygments import highlight
from pygments.lexers.dotnet import VbNetLexer

from obfuscator.modifier.base import Modifier
from obfuscator.modifier.strings.strings import StringFormatter
from obfuscator.msdocument import MSDocument

LOG = logging.getLogger(__name__)


class SplitStrings(Modifier):
    def run(self, doc: MSDocument) -> None:
        doc.code = highlight(doc.code, VbNetLexer(), _SplitStringsFmtr())


class _SplitStringsFmtr(StringFormatter):
    def __init__(self):
        super().__init__()
        self.crypt_key = []

    def _split_string(self, s: str) -> str:
        if len(s) > 8:
            s = s.strip('"')
            pos = _split_string(s)
            splitted_string = '"{}" & "{}"'.format(s[:pos], s[pos:])
            LOG.debug("Splitted '{}' in two.".format(s))
            return splitted_string
        else:
            return s

    def _run_on_string(self, s: str):
        return self._split_string(s)


def _split_string(s: str) -> int:
    """
    Split a string in two. This function will never split an escaped double quote ("") in half.
    :param s:
    :return:
    """
    split_possibilities = len(s) - 1
    impossible_split_pos = s.count('"') // 2
    if split_possibilities - impossible_split_pos <= 0:
        return -1

    split_possibilities -= impossible_split_pos

    i = 0
    pos = random.randint(1, split_possibilities)
    while pos > 0:
        if s[i] == '"':
            i = i + 1
        i = i + 1
        pos = pos - 1

    return i

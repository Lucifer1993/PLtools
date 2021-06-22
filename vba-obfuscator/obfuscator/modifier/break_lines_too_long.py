import logging
from typing import List

from pygments import highlight
from pygments.formatter import Formatter
from pygments.lexers.dotnet import VbNetLexer
from pygments.token import Token

from obfuscator.modifier.base import Modifier
from obfuscator.msdocument import MSDocument

MAX_LINE_WIDTH = 500

LOG = logging.getLogger(__name__)


def _do_split_line(line: str) -> str:
    return highlight(line, VbNetLexer(), _BreakLinesTooLong())


def _split_line_if_necessary(line: str) -> str:
    if len(line) >= MAX_LINE_WIDTH:
        LOG.info("Line '{:.30s}[...]' is too long.".format(line))
        return _do_split_line(line)
    return line


class BreakLinesTooLong(Modifier):
    def run(self, doc: MSDocument) -> None:
        code = doc.code

        code = code.split("\n")
        code = map(_split_line_if_necessary, code)
        code = "\n".join(code)

        doc.code = code


def break_line(chunks: List[str]):
    lines = ['']
    for chunk in chunks:
        if len(lines[-1]) + len(chunk) < MAX_LINE_WIDTH:
            lines[-1] += chunk
        else:
            lines += [chunk]

    result = " _\n".join(lines)
    return result


class _BreakLinesTooLong(Formatter):
    def format(self, tokensource, outfile):
        line = ''

        # First find out all the position where we can "cut" the string.
        break_points = [0]
        for ttype, value in tokensource:
            line += value
            if ttype == Token.Punctuation and value in ",+&":
                break_points.append(len(line))
        break_points.append(len(line))

        # Cut the strings at all the previously defined positions.
        chunks = []
        for i in range(len(break_points) - 1):
            bp1 = break_points[i]
            bp2 = break_points[i + 1]
            chunks.append(line[bp1:bp2])

        # Take all theses chunks and construct lines.
        outfile.write(break_line(chunks))

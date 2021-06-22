# Token.Literal.Number.Integer
import random

from pygments import highlight
from pygments.formatter import Formatter
from pygments.lexers.dotnet import VbNetLexer
from pygments.token import Token

from obfuscator.modifier.base import Modifier
from obfuscator.msdocument import MSDocument


class ReplaceIntegersWithXor(Modifier):
    def run(self, doc: MSDocument) -> None:
        doc.code = highlight(doc.code, VbNetLexer(), _XorFormatter())


class ReplaceIntegersWithAddition(Modifier):
    def run(self, doc: MSDocument) -> None:
        doc.code = highlight(doc.code, VbNetLexer(), _AdditionFormatter())


class _AdditionFormatter(Formatter):
    def format(self, tokensource, outfile):
        for ttype, value in tokensource:
            if ttype == Token.Literal.Number.Integer and random.random() > .5:
                v = int(value)
                x = random.randint(0, v)
                y = v - x
                outfile.write("({}+{})".format(x, y))
            else:
                outfile.write(value)


class _XorFormatter(Formatter):
    def format(self, tokensource, outfile):
        for ttype, value in tokensource:
            if ttype == Token.Literal.Number.Integer and random.random() > .5:
                v = int(value)
                x = random.randint(0, v)
                y = v ^ x
                outfile.write("({} Xor {})".format(x, y))
            else:
                outfile.write(value)

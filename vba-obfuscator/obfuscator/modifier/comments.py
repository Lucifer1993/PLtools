from pygments import highlight
from pygments.formatter import Formatter
from pygments.lexers.dotnet import VbNetLexer
from pygments.token import Token

from obfuscator.modifier.base import Modifier
from obfuscator.msdocument import MSDocument


class StripComments(Modifier):
    def run(self, doc: MSDocument) -> None:
        doc.code = highlight(doc.code, VbNetLexer(), _StripCommentsFormatter())


class _StripCommentsFormatter(Formatter):
    def format(self, tokensource, outfile):
        for ttype, value in tokensource:
            if ttype != Token.Comment:
                outfile.write(value)
            else:
                outfile.write(value[-1])

import logging

from pygments import highlight
from pygments.formatter import Formatter
from pygments.lexers.dotnet import VbNetLexer
from pygments.token import Token

import obfuscator.modifier.base
import obfuscator.msdocument
from obfuscator.util import get_random_string_of_random_length, get_variables_defined, get_variables_parameters, \
    get_functions, get_variables_const

LOG = logging.getLogger(__name__)

BLACKLIST_SYMBOL = {
    "Workbook_Open", "AutoOpen", "Auto_Open", "Document_Open"
}


class RandomizeNames(obfuscator.modifier.base.Modifier):
    def run(self, doc: obfuscator.msdocument.MSDocument) -> None:
        vars = set(get_variables_defined(doc.code))
        consts = set(get_variables_const(doc.code))
        params = set(get_variables_parameters(doc.code))
        functions = set(get_functions(doc.code))

        names = {}
        for symbol in vars | consts | params | functions:
            if symbol not in BLACKLIST_SYMBOL:
                names[symbol] = get_random_string_of_random_length()

        doc.code = highlight(doc.code, VbNetLexer(), _RandomizeNamesFormatter(names))


class _RandomizeNamesFormatter(Formatter):
    def __init__(self, names):
        self.names = names

    def format(self, tokensource, outfile):
        left = set()
        for ttype, value in tokensource:
            if ttype == Token.Name.Function:
                outfile.write(self._get_name(value))
            elif ttype == Token.Name:
                outfile.write(self._get_name(value))
            else:
                left = left | set(ttype)
                outfile.write(value)

    def _get_name(self, name: str) -> str:
        if name in self.names:
            LOG.debug("Replacing {} with {}.".format(name, self.names[name]))
            return self.names[name]

        LOG.debug("Ignoring {}.".format(name))
        return name

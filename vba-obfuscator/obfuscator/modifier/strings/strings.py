import logging

from pygments.formatter import Formatter
from pygments.token import Token

LOG = logging.getLogger(__name__)

IGNORED_SYMBOLS = {"Const", "Declare"}


class StringFormatter(Formatter):
    def __init__(self, **options):
        super().__init__(**options)
        self.lastval = ""
        self.lasttype = None

    def _run_on_string(self, s: str) -> str:
        raise NotImplementedError()

    def format(self, tokensource, outfile):
        skip_line = False
        for ttype, value in tokensource:
            if self.lasttype:
                if self.lasttype == ttype:
                    self.lastval += value
                else:
                    if "\n" in self.lastval:
                        skip_line = False
                    if ttype == Token.Keyword and value in IGNORED_SYMBOLS:  #  Skip the line if it is a const.
                        skip_line = True

                    # Crypt strings unless we are skipping the line.
                    if self.lasttype == Token.Literal.String:
                        if skip_line:
                            outfile.write(self.lastval)
                        else:
                            outfile.write(self._run_on_string(self.lastval))
                    else:
                        outfile.write(self.lastval)
                    self.lastval = value
            else:
                self.lastval = value
            self.lasttype = ttype

        outfile.write(value)

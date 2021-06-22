import base64
import os
import logging
import random
from typing import List

from pygments import highlight
from pygments.lexers.dotnet import VbNetLexer

from obfuscator.modifier.base import Modifier
from obfuscator.modifier.strings.strings import StringFormatter
from obfuscator.msdocument import MSDocument
from obfuscator.util import get_random_string, split_var_declaration_from_code

LOG = logging.getLogger(__name__)
VBA_XOR_FUNCTION = split_var_declaration_from_code("""
Private Function unxor(ciphertext As Variant, start As Integer)
    Dim cleartext As String
    Dim key() As Byte
    key = Base64Decode(ActiveDocument.Variables("{}"))
    cleartext = ""
    
    For i = LBound(ciphertext) To UBound(ciphertext)
        cleartext = cleartext & Chr(key(i+start) Xor ciphertext(i))
    Next
    unxor = cleartext

End Function
""")
with open(os.path.join(os.path.dirname(__file__), "base64.vbs")) as f:
    VBA_BASE64_FUNCTION = split_var_declaration_from_code(f.read())


class CryptStrings(Modifier):
    def run(self, doc: MSDocument) -> None:
        LOG.debug('Generating document variable name.')

        formatter = EncryptStringsFmtr()
        doc.code = highlight(doc.code, VbNetLexer(), formatter)

        document_var = get_random_string(16)

        code_prefix, code_suffix = split_var_declaration_from_code(doc.code)

        # Merge the codes: we must keep the global variables declarations on top.
        doc.code = code_prefix + VBA_BASE64_FUNCTION[0] + VBA_XOR_FUNCTION[0] + \
                   code_suffix + VBA_BASE64_FUNCTION[1] + VBA_XOR_FUNCTION[1].format(document_var)

        b64 = base64.b64encode(bytes(formatter.crypt_key)).decode()
        MAX_LENGTH = 512
        printable_b64 = [b64[i:i + MAX_LENGTH] for i in range(0, len(b64), MAX_LENGTH)]
        printable_b64 = '" & _\n"'.join(printable_b64)
        LOG.info('''Paste this in your VBA editor to add the Document Variable:
ActiveDocument.Variables.Add Name:="{}", Value:="{}"'''.format(document_var, printable_b64))

        doc.code = '"Use this line to add the document variable to you file and then remove these comments."\n' + \
                   'ActiveDocument.Variables.Add Name:="{}", Value:="{}"\n'.format(document_var,
                                                                                   printable_b64) + doc.code

        doc.doc_var[document_var] = b64


class EncryptStringsFmtr(StringFormatter):
    def __init__(self):
        super().__init__()
        self.crypt_key = []

    def _obfuscate_string(self, s: str) -> str:
        s = s[1:-1]
        LOG.debug("Generating XOR key for '{}'.".format(s))
        start = len(self.crypt_key)
        key = _get_random_key(len(s))
        self.crypt_key += key
        LOG.debug("XOR key will be at [{}; {}].".format(start, len(key)))

        ciphertext = _xor_crypt(s, key)
        array = _to_vba_array(ciphertext)
        LOG.debug("Encrypted string to VBA Array -> {}.".format(array))

        return 'unxor({},{})'.format(array, start)

    def _run_on_string(self, s: str):
        return self._obfuscate_string(s)


def _get_random_key(n: int) -> List[int]:
    return [random.randint(0, 255) for _ in range(n)]


def _xor(t):
    a, b = t
    return a ^ b


def _xor_crypt(msg, key):
    msg = map(ord, msg)
    str_key = zip(msg, key)
    return map(_xor, str_key)


def _to_vba_array(arr):
    arr = map(str, arr)
    numbers = ",".join(arr)
    return "Array({})".format(numbers)

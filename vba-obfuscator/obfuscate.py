#!/usr/bin/env python3
import argparse
import logging
import sys

from obfuscator.log import configure_logging
from obfuscator.modifier.base import Pipe
from obfuscator.modifier.break_lines_too_long import BreakLinesTooLong
from obfuscator.modifier.comments import StripComments
from obfuscator.modifier.functions_vars import RandomizeNames
from obfuscator.modifier.misc import RemoveEmptyLines, RemoveIndentation
from obfuscator.modifier.numbers import ReplaceIntegersWithAddition, ReplaceIntegersWithXor
from obfuscator.modifier.strings import CryptStrings, SplitStrings
from obfuscator.msdocument import MSDocument


class BadPathError(ValueError):
    pass


def main():
    configure_logging()

    LOG.info("VBA obfuscator - Thomas LEROY & Nicolas BONNET")

    parser = argparse.ArgumentParser(description='Obfuscate a VBA file.')
    parser.add_argument('input_file', type=str, action='store',
                        help='path of the file to obfuscate')
    parser.add_argument('--output_file', type=str, action='store',
                        help='output file (if no file is supplied, stdout will be used)')
    args = parser.parse_args()

    try:
        doc = MSDocument(args.input_file)
    except OSError as e:
        raise BadPathError("Could not open input file") from e
    LOG.info("Loaded the code.")

    Pipe(doc).run(
        SplitStrings(),
        CryptStrings(),
        RandomizeNames(),
        ReplaceIntegersWithAddition(),
        ReplaceIntegersWithXor(),
        StripComments(),
        RemoveIndentation(),
        BreakLinesTooLong(),
        RemoveEmptyLines(),
    )
    LOG.info("Obfuscated the code.")

    if args.output_file:
        try:
            with open(args.output_file, "w") as f:
                f.write(doc.code)
        except OSError as e:
            raise BadPathError("Could not open output file") from e
        LOG.info("Wrote to file.")
    else:
        sys.stdout.write(doc.code)


if __name__ == "__main__":
    try:
        LOG = logging.getLogger(__name__)
        main()
    except BadPathError as e:
        LOG.error("{}: {}.".format(e.args[0], e.__cause__.args[1]))
        sys.exit(2)
    except KeyboardInterrupt:
        sys.exit(1)

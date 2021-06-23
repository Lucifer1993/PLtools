#!/usr/bin/env python

"""
uptux by initstring (gitlab.com/initstring)

This tool checks for configuration issues on Linux systems that may lead to
privilege escalation.

All functionality is contained in a single file, because installing packages
in restricted shells is a pain.
"""


import os
import sys
import socket
import getpass
import argparse
import datetime
import subprocess
import inspect
import glob
import re


BANNER = r'''



                      ____ ___      ___________            
                     |    |   \_____\__    ___/_ _____  ___
                     |    |   /\____ \|    | |  |  \  \/  /
                     |    |  / |  |_> >    | |  |  />    < 
                     |______/  |   __/|____| |____//__/\_ \
                               |__|                      \/
                     
                     
                   
                        PrivEsc for modern Linux systems
                          github.com/initstring/uptux


'''


########################## Global Declarations Follow #########################

LOGFILE = 'log-uptux-{:%Y-%m-%d-%H.%M.%S}'.format(datetime.datetime.now())
PARSER = argparse.ArgumentParser(description=
                                 "PrivEsc for modern Linux systems,"
                                 " by initstring (gitlab.com/initstring)")
PARSER.add_argument('-n', '--nologging', action='store_true',
                    help='do not write output to a logfile')
PARSER.add_argument('-d', '--debug', action='store_true',
                    help='print some debugging info to the console')
ARGS = PARSER.parse_args()

## Known directories for storing systemd files.
SYSTEMD_DIRS = ['/etc/systemd/**/',
                '/lib/systemd/**/',
                '/run/systemd/**/',
                '/usr/lib/systemd/**/']

## Known directories for storing D-Bus configuration files
DBUS_DIRS = ['/etc/dbus-1/system.d/',
             '/etc/dbus-1/session.d']

# Target files that we know we cannot exploit
NOT_VULN = ['/dev/null',
            '.']

# Used to enable/disable the relative path checks of systemd
SYSTEMD_PATH_WRITABLE = False
########################## End of Global Declarations #########################


############################ Setup Functions Follow ###########################

# This is the place for functions that help set up the application.

def tee(text, **kwargs):
    """Used to log and print concurrently"""

    # Defining variables to print color-coded messages to the console.
    colors = {'green': '\033[92m',
              'blue': '\033[94m',
              'orange': '\033[93m',
              'red': '\033[91m',}
    end_color = '\033[0m'
    boxes = {'ok': colors['blue'] + '[*] ' + end_color,
             'note': colors['green'] + '[+] ' + end_color,
             'warn': colors['orange'] + '[!] ' + end_color,
             'vuln': colors['red'] + '[VULNERABLE] ' + end_color,
             'sus': colors['orange'] + '[INVESTIGATE] ' + end_color}

    # If this function is called with an optional 'box=xxx' parameter, these
    # will be prepended to the message.
    box = kwargs.get('box', '')
    if box:
        box = boxes[box]

    # First, just print the item to the console.
    print(box + text)

    # Then, write it to the log if logging is not disabled
    if not ARGS.nologging:
        try:
            with open(LOGFILE, 'a') as logfile:
                logfile.write(box + text + '\n')
        except PermissionError:
            ARGS.nologging = True
            print(boxes['warn'] + "Could not create a log file due to"
                  " insufficient permissions. Continuing with checks...")


def check_handler(check, check_name, check_desc):
    """Check handler

    This function takes a dictionary of check_desc,check_name and will
    iterate through them all.
    """
    tee("\n\n++++++++++  {}: {}  ++++++++++\n\n"
        .format(check_name, check_desc))
    tee("Starting module at {:%Y-%m-%d-%H.%M.%S}"
        .format(datetime.datetime.now()), box='ok')
    tee("\n")
    check()
    tee("\n")
    tee("Finished module at {:%Y-%m-%d-%H.%M.%S}\n"
        .format(datetime.datetime.now()), box='ok')


def get_function_order(function):
    """Helper function for build_checks_list"""
    # Grabs the line number of a function it is passed.
    order = function.__code__.co_firstlineno
    return order


def build_checks_list():
    """Dynamically build list of checks to execute

    This function will grab, in order, all functions that start with
    'uptux_check_' and populate a list. This is then used to run the checks.
    """
    # Start to build a list of functions we will execute.
    uptux_checks = []

    # Get the name of this python script and all the functions inside it.
    current_module = sys.modules[__name__]
    all_functions = inspect.getmembers(current_module, inspect.isfunction)

    # If the function name matches 'uptux_check_' we will include it.
    for function in all_functions:
        function_name = function[0]
        function_object = function[1]
        if 'uptux_check_' in function_name:
            uptux_checks.append(function_object)

    # Use the helper function to sort by line number in script.
    uptux_checks.sort(key=get_function_order)

    # Return the sorted list of functions.
    return uptux_checks

############################# End Setup Functions #############################


########################### Helper Functions Follow ###########################

# This is the place to put functions that are used by multiple "Individual
# Checks" (those starting with uptux_check_).

def shell_exec(command):
    """Executes Linux shell commands"""
    # Split the command into a list as needed by subprocess
    command = command.split()

    # Get both stdout and stderror from command. Grab the Python exception
    # if there is one.
    try:
        out_bytes = subprocess.check_output(command,
                                            stderr=subprocess.STDOUT)
    except subprocess.CalledProcessError as error:
        out_bytes = error.output
    except OSError as error:
        print('Could not run the following OS command. Sorry!\n'
              '   Command: {}'.format(command))
        print(error)
        sys.exit()

    # Return the lot as a text string for processing.
    out_text = out_bytes.decode('utf-8')
    out_text = out_text.rstrip()
    return out_text


def find_system_files(**kwargs):
    """Locates system files

    Expected kwargs: known_dirs, search_mask
    """

    # Define known Linux folders for storing service unit definitions
    return_list = set()

    # Recursively gather all service unit from the known directories
    # and add them to a deduplicated set.
    for directory in kwargs['known_dirs']:
        found_files = glob.glob(directory + kwargs['search_mask'])
        for item in found_files:
            # We don't care about files that point to /dev/null.
            if '/dev/null' not in os.path.realpath(item):
                return_list.add(item)

    if ARGS.debug:
        print("DEBUG, FOUND FILES")
        for item in return_list:
            print(item)

    return return_list


def regex_vuln_search(**kwargs):
    """Helper function for searching text files

    This function will take a list of file paths and search through
    them with a given regex. Relevant messages will be printed to the console
    and log.

    Expected kwargs: file_paths, regex, message_text, message_box
    """
    # Start a list of dictionaries for files with interesting content.
    return_list = []

    # Open up each individual file and read the text into memory.
    for file_name in kwargs['file_paths']:
        return_dict = {}

        # Continue if we can't access the file.
        if not os.access(file_name, os.R_OK):
            continue

        file_object = open(file_name, 'r')
        file_text = file_object.read()
        # Use the regex we pass in to the function to look for vulns.
        found = re.findall(kwargs['regex'], file_text)

        # Save the file name and the interesting lines of text
        if found:
            return_dict['file_name'] = file_name
            return_dict['text'] = found
            return_list.append(return_dict)

    # If the function is supplied with message info, print to console and log.
    # This function may be used instead as input to another function, so we
    # don't always want to print here.
    if return_list and kwargs['message_text'] and kwargs['message_box']:
        # Print to console and log the interesting file names and content.
        tee("")
        tee(kwargs['message_text'], box=kwargs['message_box'])
        for item in return_list:
            tee("  {}:".format(item['file_name']))
            for text in item['text']:
                tee("    {}".format(text))
            tee("")

    if ARGS.debug:
        print("DEBUG, SEARCH RESULTS")
        for item in return_list:
            print(item['file_name'])
            for text in item['text']:
                print("  {}".format(text))

    return return_list


def check_file_permissions(**kwargs):
    """Helper function to check permissions and symlink status

    This function will take a list of file paths, resolve them to their
    actual location (for symlinks), and determine if they are writeable
    by the current user. Will also alert on broken symlinks and whether
    the target directory for the broken link is writeable.

    Expected kwargs: file_paths, files_message_text, dirs_message_text,
    message_box
    """
    # Start deuplicated sets for interesting files and directories.
    writeable_files = set()
    writeable_dirs = set()

    for file_name in kwargs['file_paths']:

        # Ignore known not-vulnerable targets
        if file_name in NOT_VULN:
            continue

        # Is it a symlink? If so, get the real path and check permissions.
        # If it is broken, check permissions on the parent directory.
        if os.path.islink(file_name):
            target = os.readlink(file_name)

            # Some symlinks use relative path names. Find these and prepend
            # the directory name so we can investigate properly.
            if target[0] == '.':
                parent_dir = os.path.dirname(file_name)
                target = parent_dir + '/' + target

            if os.path.exists(target) and os.access(target, os.W_OK):
                writeable_files.add('{} -- symlink --> {}'
                                    .format(file_name, target))
            else:
                parent_dir = os.path.dirname(target)
                if os.access(parent_dir, os.W_OK):
                    writeable_dirs.add((file_name, target))

        # OK, not a symlink. Just check permissions.
        else:
            if os.access(file_name, os.W_OK):
                writeable_files.add(file_name)

    if writeable_files:
        # Print to console and log the interesting findings.
        tee("")
        tee(kwargs['files_message_text'], box=kwargs['message_box'])
        for item in writeable_files:
            tee("  {}".format(item))

    if writeable_dirs:
        # Print to console and log the interesting findings.
        tee("")
        tee(kwargs['dirs_message_text'], box=kwargs['message_box'])
        for item in writeable_dirs:
            tee("  {} --> {}".format(item[0], item[1]))

    if not writeable_files and not writeable_dirs:
        tee("")
        tee("No writeable targets. This is expected...",
            box='note')


def check_command_permission(**kwargs):
    """Checks permissions on commands returned from inside files

    Loops through a provided list of dictionaries with a file name and
    commands found within. Checks to see if they are writeable or missing and
    living in a writable directory.

    Expected kwargs: file_paths, regex, message_text, message_box
    """
    # Start an empty list for the return of the writable files/commands
    return_list = []

    for item in kwargs['commands']:
        return_dict = {}
        return_dict['text'] = []

        # The commands we have are long and may include parameters or even
        # multiple commands with pipes and ;. We try to split this all out
        # below.
        for command in item['text']:
            command = re.sub(r'[\'"]', '', command)
            command = re.split(r'[ ;\|]', command)

            # We now have a list of some commands and some parameters and
            # other garbage. Checking for os access will clean this up for us.
            # The lines below determine if we have write access to anything.
            # It also checks for the case where the target does not exist  but
            # the parent directory is writeable.
            for split_command in command:
                vuln = False

                # Ignore known not-vulnerable targets
                if split_command in NOT_VULN:
                    continue

                # Some systemd items will specicify a command with a path
                # relative to the calling item, particularly timer files.
                relative_path = os.path.dirname(item['file_name'])

                # First, check the obvious - is this a writable command?
                if os.access(split_command, os.W_OK):
                    vuln = True

                # What about if we assume it is a relative path?
                elif os.access(relative_path + '/' + split_command, os.W_OK):
                    vuln = True

                # Or maybe it doesn't exist at all, but is in a writeable
                # director?
                elif (os.access(os.path.dirname(split_command), os.W_OK)
                      and not os.path.exists(split_command)):
                    vuln = True

                # If so, pack it all up in a new dictionary which is used
                # below for output.
                if vuln:
                    return_dict['file_name'] = item['file_name']
                    return_dict['text'].append(split_command)
                    if return_dict not in return_list:
                        return_list.append(return_dict)

    if return_list and kwargs['message_text'] and kwargs['message_box']:
        # Print to console and log the interesting file names and content.
        tee("")
        tee(kwargs['message_text'], box=kwargs['message_box'])
        for item in return_list:
            tee("  {}:".format(item['file_name']))
            for text in item['text']:
                tee("    {}".format(text))
            tee("")


########################## Helper Functions Complete ##########################


########################### Individual Checks Follow ##########################

# Note: naming a new function 'uptux_check_xxxx' will automatically
# include it in execution. These will trigger in the same order listed
# in the script. The docstring will be pulled and used in the console and
# log file, so keep it short (one line).

def uptux_check_sysinfo():
    """Gather basic OS information"""
    # Gather a few basics for the report.
    uname = os.uname()
    tee("Host: {}".format(uname[1]))
    tee("OS: {}, {}".format(uname[0], uname[3]))
    tee("Kernel: {}".format(uname[2]))
    tee("Current user: {} (UID {} GID {})".format(getpass.getuser(),
                                                  os.getuid(),
                                                  os.getgid()))
    tee("Member of following groups:\n  {}".format(shell_exec('groups')))


def uptux_check_systemd_paths():
    """Check if systemd PATH is writeable"""
    # Define the bash command.
    command = 'systemctl show-environment'
    output = shell_exec(command)

    # Define the regex to find in the output.
    regex = re.compile(r'PATH=(.*$)')

    # Take the output from bash and split it into a list of paths.
    output = re.findall(regex, output)

    # This command may fail in some environments, only proceed if we have
    # a good match.
    if output:
        output = output[0].split(':')

        # Check each path - if it is writable, add it to a list.
        writeable_paths = []
        for item in output:
            if os.access(item, os.W_OK):
                writeable_paths.append(item)
    else:
        writeable_paths = False

    # Write the status to the console and log.
    if writeable_paths:
        tee("The following systemd paths are writeable. THIS IS ODD!\n"
            "See if you can combine this with a relative path Exec statement"
            " for privesc:",
            box='vuln')
        for path in writeable_paths:
            tee("  {}".format(path))
        global SYSTEMD_PATH_WRITABLE
        SYSTEMD_PATH_WRITABLE = True
    else:
        tee("No systemd paths are writeable. This is expected...",
            box='note')


def uptux_check_services():
    """Inspect systemd service unit files"""
    # Define known Linux folders for storing service unit definitions
    units = set()
    mask = '*.service'
    units = find_system_files(known_dirs=SYSTEMD_DIRS,
                              search_mask=mask)

    tee("Found {} service units to analyse...\n".format(len(units)),
        box='ok')

    # Test for write access to any service files.
    # Will resolve symlinks to their target and also check for broken links.
    text = 'Found writeable service unit files:'
    text2 = 'Found writeable directories referred to by broken symlinks'
    box = 'vuln'
    tee("")
    tee("Checking permissions on service unit files...",
        box='ok')
    check_file_permissions(file_paths=units,
                           files_message_text=text,
                           dirs_message_text=text2,
                           message_box=box)

    # Only check relative paths if we can abuse them
    if SYSTEMD_PATH_WRITABLE:
        # Look for relative calls to binaries.
        # Example: ExecStart=somfolder/somebinary
        regex = re.compile(r'^Exec(?:Start|Stop|Reload)='
                           r'(?:@[^/]'    # special exec
                           r'|-[^/]'      # special exec
                           r'|\+[^/]'     # special exec
                           r'|![^/]'      # special exec
                           r'|!![^/]'     # special exec
                           r'|)'          # or maybe no special exec
                           r'[^/@\+!-]'   # not abs path or special exec
                           r'.*',         # rest of line
                           re.MULTILINE)
        text = ('Possible relative path in an Exec statement.\n'
                'Unless you have writeable systemd paths, you won\'t be able to'
                ' exploit this:')
        box = 'sus'
        tee("")
        tee("Checking for relative paths in service unit files [check 1]...",
            box='ok')
        regex_vuln_search(file_paths=units,
                          regex=regex,
                          message_text=text,
                          message_box=box)

        # Look for relative calls to binaries but invoked by an interpreter.
        # Example: ExecStart=/bin/sh -c 'somefolder/somebinary'
        regex = re.compile(r'^Exec(?:Start|Stop|Reload)='
                           r'(?:@[^/]'    # special exec
                           r'|-[^/]'      # special exec
                           r'|\+[^/]'     # special exec
                           r'|![^/]'      # special exec
                           r'|!![^/]'     # special exec
                           r'|)'          # or maybe no special exec
                           r'.*?(?:/bin/sh|/bin/bash) '   # interpreter
                           r'(?:[\'"]|)'  # might have quotes
                           r'(?:-[a-z]+|)'# might have params
                           r'(?:[ ]+|)'   # might have more spaces now
                           r'[^/-]'       # not abs path or param
                           r'.*',         # rest of line
                           re.MULTILINE)
        text = ('Possible relative path invoked with an interpreter in an'
                ' Exec statement.\n'
                'Unless you have writable systemd paths, you won\'t be able to'
                ' exploit this:')
        box = 'sus'
        tee("")
        tee("Checking for relative paths in service unit files [check 2]...",
            box='ok')
        regex_vuln_search(file_paths=units,
                          regex=regex,
                          message_text=text,
                          message_box=box)

    # Check for write access to any commands invoked by Exec statements.
    # Thhs regex below is used to extract command lines.
    regex = re.compile(r'^Exec.*?=[!@+-]*(.*?$)',
                       re.MULTILINE)
    # We don't pass message info to this function as we need to perform more
    # processing on the output to determine what is writeable.
    tee("")
    tee("Checking for write access to commands referenced in service files...",
        box='ok')
    service_commands = regex_vuln_search(file_paths=units,
                                         regex=regex,
                                         message_text='',
                                         message_box='')

    # Another helper function to take the extracted commands and check for
    # write permissions.
    text = 'You have write access to commands referred to in service files:'
    box = 'vuln'
    check_command_permission(commands=service_commands,
                             message_text=text,
                             message_box=box)


def uptux_check_timer_units():
    """Inspect systemd timer unit files"""
    units = set()
    mask = '*.timer'
    units = find_system_files(known_dirs=SYSTEMD_DIRS,
                              search_mask=mask)

    tee("Found {} timer units to analyse...\n".format(len(units)),
        box='ok')

    # Test for write access to any timer files.
    # Will resolve symlinks to their target and also check for broken links.
    text = 'Found writeable timer unit files:'
    text2 = 'Found writeable directories referred to by broken symlinks'
    box = 'vuln'
    tee("")
    tee("Checking permissions on timer unit files...",
        box='ok')
    check_file_permissions(file_paths=units,
                           files_message_text=text,
                           dirs_message_text=text2,
                           message_box=box)

    # Timers may reference systemd services, which are already being checked.
    # But they may reference a specific script (often a '.target' file of the
    # same name. Check to see if the action called is writable.
    # The regex below is used to extract these targets.
    regex = re.compile(r'^Unit=*(.*?$)',
                       re.MULTILINE)

    # We don't pass message info to this function as we need to perform more
    # processing on the output to determine what is writeable.
    tee("")
    tee("Checking for write access to commands referenced in timer files...",
        box='ok')
    timer_commands = regex_vuln_search(file_paths=units,
                                       regex=regex,
                                       message_text='',
                                       message_box='')

    # Another helper function to take the extracted commands and check for
    # write permissions.
    text = 'You have write access to commands referred to in timer files:'
    box = 'vuln'
    check_command_permission(commands=timer_commands,
                             message_text=text,
                             message_box=box)


def uptux_check_socket_units():
    """Inspect systemd socket unit files"""
    units = set()
    mask = '*.socket'
    units = find_system_files(known_dirs=SYSTEMD_DIRS,
                              search_mask=mask)

    tee("Found {} socket units to analyse...\n".format(len(units)),
        box='ok')

    # Test for write access to any socket files.
    # Will resolve symlinks to their target and also check for broken links.
    text = 'Found writeable socket unit files:'
    text2 = 'Found writeable directories referred to by broken symlinks'
    box = 'vuln'
    tee("")
    tee("Checking permissions on socket unit files...",
        box='ok')
    check_file_permissions(file_paths=units,
                           files_message_text=text,
                           dirs_message_text=text2,
                           message_box=box)

    # Check for write access to any socket files created by a service.
    # This can be interesting - I've seen a systemd service run a REST
    # API on a AF_UNIX socket. This would be missed by normal privesc
    # checks.
    # Thhs regex below is used to extract command lines.
    regex = re.compile(r'^Listen.*?=[!@+-]*(.*?$)',
                       re.MULTILINE)

   # We don't pass message info to this function as we need to perform more
    # processing on the output to determine what is writeable.
    tee("")
    tee("Checking for write access to AF_UNIX sockets...",
        box='ok')
    socket_files = regex_vuln_search(file_paths=units,
                                     regex=regex,
                                     message_text='',
                                     message_box='')

    # Another helper function to take the extracted commands and check for
    # write permissions.
    text = ('You have write access to AF_UNIX socket files invoked by a'
            ' systemd service.\n'
            'This could be interesting. \n'
            'You can attach to these files to look for an exploitable API.')
    box = 'sus'
    check_command_permission(commands=socket_files,
                             message_text=text,
                             message_box=box)


def uptux_check_socket_apis():
    """Look for web servers on UNIX domain sockets"""
    # Use Linux ss tool to find sockets in listening state
    command = 'ss -xlp -H state listening'
    output = shell_exec(command)

    # We get Unicode back from the above command, let's fix
    output = str(output)

    root_sockets = []
    socket_replies = {}

    # We want to grab all the strings that look like socket paths
    sockets = re.findall(r' (/.*?) ', output)
    abstract_sockets = re.findall(r' (@.*?) ', output)

    for socket_path in sockets:
        # For now, we are only interested in sockets owned by root
        if os.path.exists(socket_path) and os.stat(socket_path).st_uid == 0:
            root_sockets.append(socket_path)

    tee("Trying to connect to {} unix sockets owned by uid 0..."
        .format(len(root_sockets)), box='ok')
    tee("")

    # Cycle through each and try to send a raw HTTP GET
    for socket_target in root_sockets:

        # Define a raw HTTP GET request
        http_get = ('GET / HTTP/1.1\r\n'
                    'Host: localhost\r\n'
                    '\r\n\r\n')

        # Try to interact with the socket like a web API
        client = socket.socket(socket.AF_UNIX, socket.SOCK_STREAM)
        client.settimeout(5)
        try:
            client.connect(socket_target)
            client.sendall(http_get.encode())

            reply = client.recv(8192).decode()

            # If we get a reply to this, we assume it is an API
            if reply:
                socket_replies[socket_target] = reply

        except (socket.error, UnicodeDecodeError):
            continue

    # If we have some replies, print to console
    # Hack-ish string replacement to get a nice indent
    if socket_replies:
        tee("The following root-owned sockets replied as follows",
            box='sus')
        for socket_path in socket_replies:
            tee("  " + socket_path + ":")
            tee("    " + socket_replies[socket_path]
                .replace('\n', '\n    '))
            tee("")


def uptux_check_dbus_issues():
    """Inspect D-Bus configuration items"""
    units = set()
    mask = '*.conf'
    units = find_system_files(known_dirs=DBUS_DIRS,
                              search_mask=mask)

    tee("Found {} D-Bus conf files to analyse...\n".format(len(units)),
        box='ok')

    # Test for write access to any files.
    # Will resolve symlinks to their target and also check for broken links.
    text = 'Found writeable D-Bus conf files:'
    text2 = 'Found writeable directories referred to by broken symlinks'
    box = 'vuln'
    tee("")
    tee("Checking permissions on D-Bus conf files...",
        box='ok')
    check_file_permissions(file_paths=units,
                           files_message_text=text,
                           dirs_message_text=text2,
                           message_box=box)

    # Checking for overly permission policies in D-Bus configuration files.
    # For example, normally "policy" is defined as a username. When defined in
    # an XML tag on its own, it applies to everyone.
    tee("")
    tee("Checking for overly permissive D-Bus configuration rules...",
        box='ok')

    regex = re.compile(r'<policy>.*?</policy>',
                       re.MULTILINE|re.DOTALL)

    text = ('These D-Bus policies may be overly permissive as they do not'
            ' specify a user or group.')
    box = 'sus'
    regex_vuln_search(file_paths=units,
                      regex=regex,
                      message_text=text,
                      message_box=box)

########################## Individual Checks Complete #########################

def main():
    """Main function"""
    print(BANNER)

    # Dynamically build list of checks to execute.
    uptux_checks = build_checks_list()

    # Use the handler to execute each check.
    for check in uptux_checks:
        check_name = check.__name__
        check_desc = check.__doc__
        check_handler(check, check_name, check_desc)

    # Good luck!
    tee("")
    tee("All done, good luck!", box='note')

if __name__ == "__main__":
    main()

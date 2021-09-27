import winrm
import contextlib
import base64
import re
from .copy import do_copy
import xml.etree.ElementTree as ET


class Client(winrm.Session):

    @contextlib.contextmanager
    def shell(self):
        protocol = self.protocol
        shell_id = protocol.open_shell()

        try:
            yield Shell(protocol, shell_id)
        finally:
            protocol.close_shell(shell_id)

    def copy(self, from_file, to_file):
        if hasattr(from_file, 'read'):
            do_copy(self, from_file, to_file, max_operations_per_shell=15)
        else:
            with open(from_file, 'rb') as f:
                do_copy(self, f, to_file, max_operations_per_shell=15)


class Shell:
    def __init__(self, protocol, shell_id):
        self.protocol = protocol
        self.shell_id = shell_id

    def cmd(self, cmd, *args):
        command_id = self.protocol.run_command(self.shell_id, cmd, args)
        result = winrm.Response(self.protocol.get_command_output(self.shell_id, command_id))
        self.protocol.cleanup_command(self.shell_id, command_id)
        return result

    def check_cmd(self, cmd, *args):
        return self._check(self.cmd(cmd, *args))

    def ps(self, script):
        # must use utf16 little endian on windows
        encoded_ps = base64.b64encode(script.encode('utf_16_le')).decode('ascii')
        result = self.cmd('powershell', '-encodedcommand', encoded_ps)
        if len(result.std_err):
            # if there was an error message, clean it it up and make it human
            # readable
            result.std_err = _clean_error_msg(result.std_err)
        return result

    def check_ps(self, script):
        result = self.ps(script)
        return self._check(result)

    @staticmethod
    def _check(result):
        stdout = result.std_out.decode()
        stderr = result.std_err.decode() if result.std_err is not None else None
        if result.status_code != 0:
            raise ShellCommandError(result.status_code, stdout, stderr)
        return stdout, stderr


def _clean_error_msg(msg):
    """converts a Powershell CLIXML message to a more human readable string
    """
    # TODO prepare unit test, beautify code
    # if the msg does not start with this, return it as is
    if msg.startswith(b"#< CLIXML\r\n"):
        # for proper xml, we need to remove the CLIXML part
        # (the first line)
        msg_xml = msg[11:]
        try:
            # remove the namespaces from the xml for easier processing
            msg_xml = _strip_namespace(msg_xml)
            root = ET.fromstring(msg_xml)
            # the S node is the error message, find all S nodes
            nodes = root.findall("./S")
            new_msg = ""
            for s in nodes:
                # append error msg string to result, also
                # the hex chars represent CRLF so we replace with newline
                new_msg += s.text.replace("_x000D__x000A_", "\n")
        except Exception as e:
            # if any of the above fails, the msg was not true xml
            # print a warning and return the orignal string
            # TODO do not print, raise user defined error instead
            print("Warning: there was a problem converting the Powershell"
                  " error message: %s" % (e))
        else:
            # if new_msg was populated, that's our error message
            # otherwise the original error message will be used
            if len(new_msg):
                # remove leading and trailing whitespace while we are here
                return new_msg.strip().encode('utf-8')


def _strip_namespace(xml):
    """strips any namespaces from an xml string"""
    p = re.compile(b"xmlns=*[\"\"][^\"\"]*[\"\"]")
    allmatches = p.finditer(xml)
    for match in allmatches:
        xml = xml.replace(match.group(), b"")
    return xml


class ShellCommandError(RuntimeError):
    def __init__(self, status_code, stdout, stderr):
        self.status_code = status_code
        self.stdout = stdout
        self.stderr = stderr

        super().__init__(f'shell command failed, code={status_code}\n{stdout[:100]}\n{stderr[:100]}')

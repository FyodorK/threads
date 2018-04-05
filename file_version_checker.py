import os
from win32api import GetFileVersionInfo
import argparse
import sys
import re

__version__ = '0.0.1'


class OverrideHelpParser(argparse.ArgumentParser):
    """
    https://stackoverflow.com/a/4042861/4738630
    """
    def error(self, message):
        sys.stderr.write('error: %s\n' % message)
        self.print_help()
        sys.exit(-2)


class VersionChecker:

    """
    Check file properties on Windows OS
    """

    def __init__(self):

        self.out_mess = []
        self.template = "{path:155}{version:20}{boo}"
        self.checking_extensions = ('exe', 'dll', 'pyd')
        self.location = os.path.dirname(os.path.realpath(__file__))
        self.path_to_exception_name_file=os.path.join(self.location, 'exceptions.txt')

        args = self.parser_args()

        self.file_to_check = args.file
        self.path_to_check = args.folder
        if args.exceptions:
            self.path_to_exception_name_file = args.exceptions

        self.exception_names = self.parse_exception_names()

    @staticmethod
    def parser_args():
        """
        Arguments parser
        :return: args object
        """

        parser = OverrideHelpParser(prog='file_version_checker',
                                    description='Check file properties on Windows OS',
                                    formatter_class=argparse.RawTextHelpFormatter,
                                    add_help=True)
        group = parser.add_mutually_exclusive_group(required=False)

        group.add_argument('-f', '--file', metavar='<File>', required=False, default=None,
                           help='Provide a path to file to check')
        group.add_argument('-d', '--folder', metavar='<Folder>', required=False, default=None,
                           help='Provide a folder path that content is needed to be checked')
        parser.add_argument('-v', '--version', action='version', help='Show version', version=__version__)
        parser.add_argument('-e', '--exceptions', help='Path to file with exceptions')

        args = parser.parse_args()

        return args

    def parse_exception_names(self):
        exception_names = []
        with open(self.path_to_exception_name_file, 'r') as exception_file:
            for line in exception_file.readlines():
                exception_names.append(line.strip())
        return exception_names

    @staticmethod
    def getfileprops(path_to_file: str) -> dict:
        """
        Get properties of given file
        https://stackoverflow.com/a/7993095/4738630
        :param path_to_file: Path to file
        :return: Dict of properties
        """
        # TODO: to rewrite using standard library

        prop_names = ('Comments', 'InternalName', 'ProductName',
                      'CompanyName', 'LegalCopyright', 'ProductVersion',
                      'FileDescription', 'LegalTrademarks', 'PrivateBuild',
                      'FileVersion', 'OriginalFilename', 'SpecialBuild')

        props = {'FixedFileInfo': None, 'StringFileInfo': None, 'FileVersion': None, 'ProductVersion': None}

        try:
            # backslash as parm returns dictionary of numeric info corresponding to VS_FIXEDFILEINFO struc
            fixed_info = GetFileVersionInfo(path_to_file, '\\')
            props['FixedFileInfo'] = fixed_info
            props['ProductVersion'] = "%d.%d.%d.%d" % (fixed_info['ProductVersionMS'] / 65536,
                                                       fixed_info['ProductVersionMS'] % 65536,
                                                       fixed_info['ProductVersionLS'] / 65536,
                                                       fixed_info['ProductVersionLS'] % 65536)

            props['FileVersion'] = "%d.%d.%d.%d" % (fixed_info['FileVersionMS'] / 65536,
                                                    fixed_info['FileVersionMS'] % 65536,
                                                    fixed_info['FileVersionLS'] / 65536,
                                                    fixed_info['FileVersionLS'] % 65536)

            # \VarFileInfo\Translation returns list of available (language, codepage)
            # pairs that can be used to retreive string info. We are using only the first pair.
            lang, codepage = GetFileVersionInfo(path_to_file, '\\VarFileInfo\\Translation')[0]

            # any other must be of the form \StringfileInfo\%04X%04X\parm_name, middle
            # two are language/codepage pair returned from above

            str_info = {}
            for propName in prop_names:
                str_info_path = u'\\StringFileInfo\\%04X%04X\\%s' % (lang, codepage, propName)
                str_info[propName] = GetFileVersionInfo(path_to_file, str_info_path)

            props['StringFileInfo'] = str_info
        except Exception:
            pass

        return props

    def gfv_wrapper(self, path_to_file):

        filename = path_to_file.split('\\')[-1]
        file_extension = filename.split('.')[-1]

        if file_extension in self.checking_extensions:
            in_exception_list = False
            props = self.getfileprops(path_to_file)
            version = str(props.get('ProductVersion'))
            if filename in self.exception_names:
                in_exception_list = True
            out = self.template.format(path=path_to_file, version=version, boo=in_exception_list)
            self.out_mess.append(out)
        else:
            pass

    def gfv_wrap_folder(self, path_to):
        for root, folders, files in os.walk(path_to, topdown=False):
            for file in [os.path.join(root, fl) for fl in files]:
                self.gfv_wrapper(file)
            for folder in [os.path.join(root, fld) for fld in folders]:
                self.gfv_wrap_folder(folder)

    @staticmethod
    def uniq_sort(cont: list) -> list:
        """
        Just a crutch
        """
        res = []
        [res.append(i) for i in cont if not res.count(i)]
        return res

    def post_process(self):
        """
        Add some statistic to resulting output
        :return: None
        """
        top_ = "{path:155}{version:20}{boo:20}".format(path='Path to the file',
                                                       version='Product version',
                                                       boo='In exception list')
        template = {
            "result": "{h}\nFiles verified: {n}\nFiles have product version: {np}\n"
                      "Files have not product version: {na}\n",
            "status": "Overall result: {}",
            "failed": "Failed objects: {}"
        }

        header = "=" * 142
        count_none = 0
        failed_dlls = []
        total_files = len(self.out_mess)

        for member in self.out_mess:
            if re.search(r'None(.+)False', member):
                failed_dlls.append(member)

        for member in self.out_mess:
            if ' None ' in member:
                count_none += 1

        out = template.get("result").format(h=header, n=total_files, np=total_files - count_none, na=count_none)
        self.out_mess.append(out)

        count_fails = len(failed_dlls)
        if count_fails > 0:
            self.out_mess.append(template.get("status").format("Test Failed"))
            self.out_mess.append(template.get("failed").format(count_fails))
            for item in failed_dlls:
                self.out_mess.append(item)
        else:
            self.out_mess.append(template.get("status").format("Test passed"))

        self.out_mess.insert(0, top_)

    def main(self):
        if self.file_to_check and os.path.isfile(self.file_to_check):
            self.gfv_wrapper(self.file_to_check)
        elif self.path_to_check and os.path.isdir(self.path_to_check):
            self.gfv_wrap_folder(self.path_to_check)
        else:
            out = "{path:155}{version:20}".format(path='Not available', version='Object is not found')
            self.out_mess.append(out)

        self.out_mess = self.uniq_sort(self.out_mess)

        self.post_process()
        print('\n'.join(self.out_mess))


if __name__ == '__main__':
    check = VersionChecker()
    check.main()

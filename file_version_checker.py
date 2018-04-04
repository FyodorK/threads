import os
from win32api import GetFileVersionInfo
import argparse
import sys

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
        self.file_to_check = ""
        self.path_to_check = ""
        self.out_mess = []
        self.template = "{time}\t\tPath to the file: {path:155}\t\t\tProduct version: {version}"
        self.checking_extensions = ('exe', 'dll', 'pyd')

        args = self.parser_args()

        self.file_to_check = args.file
        self.path_to_check = args.folder

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

        args = parser.parse_args()

        return args

    @staticmethod
    def getfileprops(path_to_file):
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

    @staticmethod
    def get_time():
        return __import__('datetime').datetime.now().strftime("%Y%m%d_%H%M%S.%f")

    def gfv_wrapper(self, path_to_file):

        filename = path_to_file.split('\\')[-1]
        file_extension = filename.split('.')[-1]

        if file_extension in self.checking_extensions:
            props = self.getfileprops(path_to_file)
            version = str(props.get('ProductVersion'))

            out = self.template.format(time=self.get_time(), path=path_to_file, version=version)
            self.out_mess.append(out)
        else:
            pass

    def gfv_wrap_folder(self, path_to):
        for root, folders, files in os.walk(path_to, topdown=False):
            for file in [os.path.join(root, fl) for fl in files]:
                self.gfv_wrapper(file)
            for folder in [os.path.join(root, fld) for fld in folders]:
                self.gfv_wrap_folder(folder)

    def post_porn(self):
        """
        Add some statistic to resulting output
        :return: None
        """
        template = "{h}\nFiles verified: {n}\nFiles have product version: {np}\nFiles have not product version: {na}\n"
        header = "\n" + "=" * 142
        count_none = 0
        total_files = len(self.out_mess)

        for member in self.out_mess:
            if 'Product version: None' in member:
                count_none += 1

        out = template.format(h=header, n=total_files, np=total_files - count_none, na=count_none)
        self.out_mess.append(out)

    def main(self):
        if self.file_to_check and os.path.isfile(self.file_to_check):
            self.gfv_wrapper(self.file_to_check)
        elif self.path_to_check and os.path.isdir(self.path_to_check):
            self.gfv_wrap_folder(self.path_to_check)
        else:
            out = self.template.format(time=self.get_time(), path='Not available', version='Object is not found')
            self.out_mess.append(out)

        self.post_porn()
        print('\n'.join(self.out_mess))


if __name__ == '__main__':
    check = VersionChecker()
    check.main()

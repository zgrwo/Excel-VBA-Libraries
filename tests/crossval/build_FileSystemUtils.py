"""Cross-validate FileSystemUtils path functions against Python references.

Usage: python tests/build_FileSystemUtils.py
"""

import os
import sys
import tempfile

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from tests.crossval.build_common import CrossValRunner
from tests.test_utils import SRC_DIR, VBA_CORE_DIR, VBA_CORE_IMPORT_ORDER

MODULE_PATHS = [os.path.join(VBA_CORE_DIR, name + ".cls")
                for name in VBA_CORE_IMPORT_ORDER]
MODULE_PATHS.append(os.path.join(SRC_DIR, "FileSystemUtils.bas"))

# Create temp files for I/O testing
_tmp_dir = tempfile.mkdtemp(prefix="vba_fs_test_")
_tmp_file = os.path.join(_tmp_dir, "test_read.txt")
_nonexistent = os.path.join(_tmp_dir, "nonexistent.txt")
with open(_tmp_file, "w", encoding="utf-8") as f:
    f.write("Hello VBA!\nLine 2\nLine 3")
_expected_size = os.path.getsize(_tmp_file)


def _cleanup_tmp():
    """Clean up temp test files."""
    import shutil
    try:
        shutil.rmtree(_tmp_dir, ignore_errors=True)
    except Exception:
        pass


TEST_CASES = [
    # NormalizePath
    {"name": "NormalizePath_fwd_slash", "func": "NormalizePath",
     "args": lambda: ("C:/Data/file.txt",),
     "py_ref": lambda a: os.path.normpath(a[0]), "result_type": "string"},
    {"name": "NormalizePath_double_slash", "func": "NormalizePath",
     "args": lambda: ("C:\\\\Data\\\\file.txt",),
     "py_ref": lambda a: os.path.normpath(a[0]), "result_type": "string"},
    {"name": "NormalizePath_empty", "func": "NormalizePath",
     "args": lambda: ("",), "py_ref": lambda a: "", "result_type": "string"},
    # PathCombine
    {"name": "PathCombine_normal", "func": "PathCombine",
     "args": lambda: ("C:\\Data", "file.txt"),
     "py_ref": lambda a: os.path.join(a[0], a[1]), "result_type": "string"},
    {"name": "PathCombine_trailing", "func": "PathCombine",
     "args": lambda: ("C:\\Data\\", "file.txt"),
     "py_ref": lambda a: os.path.join(a[0], a[1]), "result_type": "string"},
    {"name": "PathCombine_empty_dir", "func": "PathCombine",
     "args": lambda: ("", "file.txt"),
     "py_ref": lambda a: "file.txt", "result_type": "string"},
    # GetFileName
    {"name": "GetFileName_full", "func": "GetFileName",
     "args": lambda: ("C:\\Data\\report.xlsx",),
     "py_ref": lambda a: os.path.basename(a[0]), "result_type": "string"},
    {"name": "GetFileName_bare", "func": "GetFileName",
     "args": lambda: ("report.xlsx",),
     "py_ref": lambda a: "report.xlsx", "result_type": "string"},
    # GetBaseName
    {"name": "GetBaseName_normal", "func": "GetBaseName",
     "args": lambda: ("report.xlsx",),
     "py_ref": lambda a: os.path.splitext(a[0])[0], "result_type": "string"},
    {"name": "GetBaseName_targz", "func": "GetBaseName",
     "args": lambda: ("archive.tar.gz",),
     "py_ref": lambda a: "archive.tar", "result_type": "string"},
    # GetExtension
    {"name": "GetExtension_xlsx", "func": "GetExtension",
     "args": lambda: ("report.xlsx",),
     "py_ref": lambda a: os.path.splitext(a[0])[1], "result_type": "string"},
    {"name": "GetExtension_none", "func": "GetExtension",
     "args": lambda: ("Makefile",),
     "py_ref": lambda a: "", "result_type": "string"},
    # GetFolderPath
    {"name": "GetFolderPath_full", "func": "GetFolderPath",
     "args": lambda: ("C:\\Data\\file.txt",),
     "py_ref": lambda a: os.path.dirname(a[0]), "result_type": "string"},
    # IsPathValid
    {"name": "IsPathValid_normal", "func": "IsPathValid",
     "args": lambda: ("C:\\Data\\file.txt",), "py_ref": lambda a: True, "result_type": "bool"},
    {"name": "IsPathValid_empty", "func": "IsPathValid",
     "args": lambda: ("",), "py_ref": lambda a: False, "result_type": "bool"},
    {"name": "IsPathValid_angle_br", "func": "IsPathValid",
     "args": lambda: ("C:\\<Data>\\file.txt",), "py_ref": lambda a: False, "result_type": "bool"},
    {"name": "IsPathValid_pipe", "func": "IsPathValid",
     "args": lambda: ("C:\\Data|file.txt",), "py_ref": lambda a: False, "result_type": "bool"},
    # GetBaseName -- edge cases
    {"name": "GetBaseName_dotfile", "func": "GetBaseName",
     "args": lambda: (".gitignore",), "py_ref": lambda a: "", "result_type": "string"},
    {"name": "GetBaseName_noext", "func": "GetBaseName",
     "args": lambda: ("Makefile",), "py_ref": lambda a: "Makefile", "result_type": "string"},
    # GetExtension -- .tar.gz
    {"name": "GetExtension_targz", "func": "GetExtension",
     "args": lambda: ("archive.tar.gz",), "py_ref": lambda a: ".gz", "result_type": "string"},
    # GetFolderPath -- edge cases
    {"name": "GetFolderPath_root", "func": "GetFolderPath",
     "args": lambda: ("C:\\file.txt",), "py_ref": lambda a: "C:\\", "result_type": "string"},
    {"name": "GetFolderPath_cwd", "func": "GetFolderPath",
     "args": lambda: ("file.txt",), "py_ref": lambda a: "", "result_type": "string"},
    # NormalizePath -- UNC path
    {"name": "NormalizePath_unc", "func": "NormalizePath",
     "args": lambda: ("\\\\server\\share\\dir",), "py_ref": lambda a: "\\\\server\\share\\dir", "result_type": "string"},
    # UDF wrappers
    {"name": "UDF_FS_FILENAME", "func": "UDF_FS_FILENAME",
     "args": lambda: ("C:\\Data\\file.txt",), "py_ref": lambda a: "file.txt", "result_type": "string"},
    {"name": "UDF_FS_ISPATHVALID", "func": "UDF_FS_ISPATHVALID",
     "args": lambda: ("C:\\",), "py_ref": lambda a: True, "result_type": "bool"},

    # File I/O tests — using temp files created by Python
    {"name": "FileExists_true", "func": "FileExists",
     "args": lambda: (_tmp_file,), "py_ref": lambda a: True, "result_type": "bool"},
    {"name": "FileExists_false", "func": "FileExists",
     "args": lambda: (_nonexistent,), "py_ref": lambda a: False, "result_type": "bool"},
    {"name": "GetFileSize", "func": "GetFileSize",
     "args": lambda: (_tmp_file,), "py_ref": lambda a: _expected_size, "result_type": "scalar", "tol": 1.0},

    # NOTE: File is written with \n by Python, but VBA ReadTextFile on
    # Windows returns \r\n line endings. This py_ref assumes Windows CRLF
    # behavior — may fail on Mac Excel which uses \r or \n endings.
    {"name": "ReadTextFile_basic", "func": "ReadTextFile",
     "args": lambda: (_tmp_file,),
     "py_ref": lambda a: "Hello VBA!\r\nLine 2\r\nLine 3", "result_type": "string"},
    {"name": "FolderExists_true", "func": "FolderExists",
     "args": lambda: (_tmp_dir,), "py_ref": lambda a: True, "result_type": "bool"},
    {"name": "FolderExists_false", "func": "FolderExists",
     "args": lambda: (_nonexistent,), "py_ref": lambda a: False, "result_type": "bool"},
    {"name": "EnsureFolder_basic", "func": "EnsureFolder",
     "args": lambda: (os.path.join(_tmp_dir, "new_subdir"),),
     "py_ref": lambda a: True, "result_type": "bool"},

    # UDF wrappers
    {"name": "UDF_FS_NORMALIZEPATH", "func": "UDF_FS_NORMALIZEPATH",
     "args": lambda: ("C:/Data/file.txt",),
     "py_ref": lambda a: os.path.normpath("C:/Data/file.txt"), "result_type": "string"},
    {"name": "UDF_FS_PATHCOMBINE", "func": "UDF_FS_PATHCOMBINE",
     "args": lambda: ("C:\\Data", "file.txt"),
     "py_ref": lambda a: os.path.join("C:\\Data", "file.txt"), "result_type": "string"},
    {"name": "UDF_FS_FILEEXISTS", "func": "UDF_FS_FILEEXISTS",
     "args": lambda: (_tmp_file,), "py_ref": lambda a: True, "result_type": "bool"},
    {"name": "UDF_FS_FOLDEREXISTS", "func": "UDF_FS_FOLDEREXISTS",
     "args": lambda: (_tmp_dir,), "py_ref": lambda a: True, "result_type": "bool"},
    {"name": "UDF_FS_BASENAME", "func": "UDF_FS_BASENAME",
     "args": lambda: ("file.txt",), "py_ref": lambda a: "file", "result_type": "string"},
    {"name": "UDF_FS_EXTENSION", "func": "UDF_FS_EXTENSION",
     "args": lambda: ("file.txt",), "py_ref": lambda a: ".txt", "result_type": "string"},
    {"name": "UDF_FS_FOLDERPATH", "func": "UDF_FS_FOLDERPATH",
     "args": lambda: ("C:\\Data\\file.txt",),
     "py_ref": lambda a: "C:\\Data", "result_type": "string"},
]

def main() -> int:
    try:
        runner = CrossValRunner("FileSystemUtils", MODULE_PATHS)
        runner.run_all(TEST_CASES)
        passed, failed = runner.print_summary()
        return 0 if failed == 0 else 1
    finally:
        _cleanup_tmp()

if __name__ == "__main__":
    sys.exit(main())

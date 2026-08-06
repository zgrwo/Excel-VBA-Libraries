"""Cross-validate FileSystemUtils path functions against Python references.

Usage: python tests/build_FileSystemUtils.py
"""

import os
import sys
import tempfile

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from tests.crossval.build_common import CrossValRunner
from tests.test_utils import SRC_DIR, VBA_CORE_DIR, VBA_CORE_IMPORT_ORDER, run_macro

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

# FS-01 回归: ReadBinaryFile 必须可调用 (历史缺陷: 未定义常量致过程级编译失败)
_bin_file = os.path.join(_tmp_dir, "test_read.bin")
_bin_bytes = bytes([1, 2, 3, 250, 251, 252])
with open(_bin_file, "wb") as f:
    f.write(_bin_bytes)

# 负例用子目录 (保证 "sub\..\test_read.txt" 指向真实存在的文件，
# 从而越过 FileExists 前置检查, 命中 ValidateSafePath 的 .. 拦截分支)
os.makedirs(os.path.join(_tmp_dir, "sub"), exist_ok=True)
_write_file = os.path.join(_tmp_dir, "test_write.bin")

# 防误杀用例: 文件名内含连续点 (非路径段 ".."), 不得被穿越检查拒绝
_dotted_file = os.path.join(_tmp_dir, "data..v2.bin")
_dotted_bytes = bytes([11, 22, 33])
with open(_dotted_file, "wb") as f:
    f.write(_dotted_bytes)


def _ensure_probe_module(wb, code):
    """注入/重建探针模块 (探针内部自带错误处理器,
    避免错误穿透 COM 边界触发 Excel 运行时错误弹窗)."""
    vbproj = wb.VBProject
    for comp in list(vbproj.VBComponents):
        if comp.Name == "FSProbe":
            vbproj.VBComponents.Remove(comp)
    comp = vbproj.VBComponents.Add(1)  # vbext_ct_StdModule
    comp.Name = "FSProbe"
    comp.CodeModule.AddFromString(code)


def _unsafe_path_reconstruct(excel, wb, ws, runner, tc, args):
    """.. 穿越路径负例: 探针内部捕获错误并返回错误码 (不弹 Excel 对话框)."""
    _ensure_probe_module(wb,
        "Public Function ProbeUnsafe(ByVal p As String) As Long\r\n"
        "    On Error GoTo EH\r\n"
        "    Dim x() As Byte\r\n"
        "    x = ReadBinaryFile(p)\r\n"
        "    ProbeUnsafe = 0\r\n"
        "    Exit Function\r\n"
        "EH: ProbeUnsafe = Err.Number\r\n"
        "End Function")
    err_num = run_macro(excel, wb, "FSProbe.ProbeUnsafe", args[0])
    return float(err_num), -2147220502.0, 0.0  # ERR_INVALID_INPUT = vbObjectError + 1002


def _write_binary_reconstruct(excel, wb, ws, runner, tc, args):
    """WriteBinaryFile 覆盖: VBA 侧写入已知字节, Python 侧校验文件内容."""
    out_path = args[0]
    vbproj = wb.VBProject
    for comp in list(vbproj.VBComponents):
        if comp.Name == "FSWriteProbe":
            vbproj.VBComponents.Remove(comp)
    comp = vbproj.VBComponents.Add(1)  # vbext_ct_StdModule
    comp.Name = "FSWriteProbe"
    comp.CodeModule.AddFromString(
        "Public Sub WBP(ByVal p As String)\r\n"
        "    Dim b(0 To 3) As Byte\r\n"
        "    b(0) = 7: b(1) = 8: b(2) = 9: b(3) = 250\r\n"
        "    WriteBinaryFile p, b\r\n"
        "End Sub")
    run_macro(excel, wb, "FSWriteProbe.WBP", out_path)
    with open(out_path, "rb") as f:
        data = f.read()
    if list(data) != [7, 8, 9, 250]:
        raise AssertionError(f"WriteBinaryFile content mismatch: {list(data)}")
    return float(sum(data)), 274.0, 0.0


def _cleanup_tmp():
    """Clean up temp test files."""
    import shutil
    try:
        shutil.rmtree(_tmp_dir, ignore_errors=True)
    except Exception:
        pass


TEST_CASES = [
    # ReadBinaryFile — FS-01 回归冒烟 (编译完整性 + 字节内容对照)
    {"name": "ReadBinaryFile_roundtrip", "func": "ReadBinaryFile",
     "args": lambda: (_bin_file,),
     "py_ref": lambda a: list(_bin_bytes), "result_type": "array"},
    # FS-01 验证条款: 含 .. 的路径必须抛 ERR_INVALID_INPUT (经 VBA 探针内部捕获,
    # 避免错误穿透 COM 触发 Excel 运行时错误弹窗)
    {"name": "ReadBinaryFile_unsafe_path", "func": "ReadBinaryFile",
     "args": lambda: (os.path.join(_tmp_dir, "sub", "..", "test_read.txt"),),
     "reconstruct": _unsafe_path_reconstruct, "py_ref": lambda a: -2147220502.0},
    # 防误杀: 文件名内连续点 ("data..v2.bin") 不构成目录穿越, 必须正常读取
    {"name": "ReadBinaryFile_dotted_filename", "func": "ReadBinaryFile",
     "args": lambda: (_dotted_file,),
     "py_ref": lambda a: list(_dotted_bytes), "result_type": "array"},
    # WriteBinaryFile — 零覆盖缺口补齐 (TEST_AUDIT_REPORT L178): VBA 写入 → Python 校验
    {"name": "WriteBinaryFile_roundtrip", "func": "WriteBinaryFile",
     "args": lambda: (_write_file,),
     "reconstruct": _write_binary_reconstruct, "py_ref": lambda a: 274.0},
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

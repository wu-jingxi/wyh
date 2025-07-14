import os
import shutil
import uuid
import zipfile
from io import BytesIO
from os import PathLike
from pathlib import Path
from uuid import uuid1


def zip_one_file(source: str | PathLike[str] | BytesIO, name: str | None = None) -> bytes:
    """
    压缩文件
    :param source: 源文件
    :param name: 压缩文件内部的文件名
    :return: 二进制数据
    """
    buffer = BytesIO()
    with zipfile.ZipFile(buffer, 'w', zipfile.ZIP_DEFLATED) as zipf:
        if isinstance(source, BytesIO):
            # 将数据指针指向开始
            source.seek(0)
            zipf.writestr(name or uuid.uuid1().hex, source.read())
        elif isinstance(source, PathLike) or isinstance(source, str) :
            zipf.write(source, name or uuid1().hex + '.pdf')
        else:
            raise TypeError('输入类型不正确')

    return buffer.getvalue()


def unzip_one_file(source: str | PathLike[str] | BytesIO) -> bytes:
    """
    解压文件
    :param source:
    :return:
    """
    # 处理不同输入类型
    if isinstance(source, (str, PathLike)):
        # 文件路径类型 - 读取文件内容到BytesIO
        with open(source, 'rb') as f:
            buffer = BytesIO(f.read())
    elif isinstance(source, BytesIO):
        # 直接使用BytesIO对象
        buffer = source
        buffer.seek(0)  # 确保从开头读取
    elif isinstance(source, bytes):
        # 原始字节数据
        buffer = BytesIO(source)
    else:
        raise TypeError(f"不支持的输入类型: {type(source)}")

    # 确保buffer处于起始位置
    buffer.seek(0)

    # 验证是否是有效的ZIP文件
    if not zipfile.is_zipfile(buffer):
        buffer.seek(0)  # 重置位置以便错误检查
        signature = buffer.read(4)
        raise zipfile.BadZipFile(f"不是有效的ZIP文件，签名: {signature}")

    buffer.seek(0)  # 再次重置位置以进行解压

    with zipfile.ZipFile(buffer, 'r') as zipf:
        # 获取压缩包中的文件列表
        files = zipf.namelist()

        # 验证压缩包内文件数量
        if len(files) == 0:
            raise ValueError("压缩包中没有文件")
        elif len(files) > 1:
            raise ValueError(f"压缩包中包含多个文件，此函数仅支持单个文件解压: {files}")

        # 获取唯一的文件信息
        file_name = files[0]
        file_info = zipf.getinfo(file_name)

        # 验证文件类型（只处理文件）
        if file_info.is_dir():
            raise ValueError(f"压缩包中的是一个目录而非文件: {file_name}")

        # 读取文件内容
        with zipf.open(file_name) as file_in_zip:
            content = file_in_zip.read()

    return content


def rm_files(source_paths: list[str | Path]) -> bool:
    """
    批量删除文件和文件夹（支持递归删除文件夹）

    :param source_paths: 要删除的文件或文件夹路径列表
    :return: 所有路径都成功删除则返回True，否则返回False

    注意：
    - 该操作是永久的，无法撤销
    - 会递归删除非空文件夹
    - 如果遇到错误，会尝试继续处理其他路径
    - 删除系统文件或权限不足时可能失败
    """
    success = True

    for source in source_paths:
        one_path = Path(source).resolve()
        try:


            # 验证路径是否在合理范围内（防止误删系统目录）
            if not _is_safe_to_delete(one_path):
                print(f"错误: 路径 '{one_path}' 被阻止删除（可能是系统或根目录）", file=sys.stderr)
                success = False
                continue

            # 如果路径不存在，跳过
            if not one_path.exists():
                continue

            # 处理文件
            if one_path.is_file():
                one_path.unlink()  # 删除文件

            # 处理目录
            elif one_path.is_dir():
                # 递归删除目录及其内容
                shutil.rmtree(one_path)

            print(f"已删除: {one_path}")

        except PermissionError as e:
            print(f"权限错误: 无法删除 '{one_path}' - {e}", file=sys.stderr)
            success = False
        except OSError as e:
            print(f"系统错误: 无法删除 '{one_path}' - {e}", file=sys.stderr)
            success = False
        except Exception as e:
            print(f"意外错误: 无法删除 '{one_path}' - {type(e).__name__}: {e}", file=sys.stderr)
            success = False

    return success


def _is_safe_to_delete(one_path: Path) -> bool:
    """安全检查: 防止误删关键系统目录"""
    # 阻止删除的目录列表（可根据需要扩展）
    protected_dirs = [
        Path(os.path.expanduser("~")),  # 用户主目录
        Path("/"), Path("C:/"), Path("D:/"),  # 根目录
        Path("/bin"), Path("/sbin"), Path("/usr"),  # Unix 系统目录
        Path("/Windows"), Path("/Program Files"),  # Windows 系统目录
    ]

    # 检查路径是否是受保护目录
    for protected in protected_dirs:
        try:
            if one_path == protected:
                return False
        except Exception:
            pass

    # 检查路径是否为空（防止误删系统根目录）
    if one_path == one_path.anchor:
        return False

    # 禁止删除根目录下的直接子项
    if one_path.parent == one_path.anchor:
        return False


    return True
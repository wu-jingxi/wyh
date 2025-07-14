import os
import shutil
import subprocess
import asyncio

from io import BytesIO
from pathlib import Path
from traceback import format_exc
from typing import Annotated, Literal
from uuid import uuid1

from pypdf import PdfReader, PdfWriter



def pdf_page_num(source: str | Path | bytes) -> int:
    """
    :param source: 源文件
    :return: 返回页数
    """

    source = __change_bytes_2_io(source)

    pdf = PdfReader(source)
    return len(pdf.pages)


def pdf_merge(source_list: list[str | Path | bytes]) -> bytes:
    """
    :param source_list: 文件列表
    :return: 融合后的pdf二进制数据
    """
    result = PdfWriter()
    for source in source_list:
        source = __change_bytes_2_io(source)

        one_pdf = PdfReader(source)

        for one_page in one_pdf.pages:
            result.add_page(one_page)

    output_buffer = BytesIO()
    result.write(output_buffer)

    return output_buffer.getvalue()


def pdf_split_by_page(source: str | Path, pages: int, mode: Literal['begin', 'end'] = 'end') -> list[bytes]:
    """
    将pdf按页划分
    :param source: 源数据
    :param pages: 按照页数进行切分，最后不足指定值的页也会被切分
    :param mode: begin-少页数的pdf在起始pdf，end-少页数的pdf在末尾pdf
    :return: 切分的页数, -1 代表解析失败
    """
    result = []

    source = __change_bytes_2_io(source)

    pdf = PdfReader(source)

    pdf_page_len = len(pdf.pages)

    complete_parts = pdf_page_len // pages

    remaining_pages = pdf_page_len % pages

    range_list = []

    if mode == "begin":
        if remaining_pages != 0:
            range_list.append(range(0, remaining_pages))
        for i in range(complete_parts):
            range_list.append(range(remaining_pages + pages * i, remaining_pages + pages * (i + 1)))

    elif mode == "end":
        for i in range(complete_parts):
            range_list.append(range(pages * i, pages * (i + 1)))

        if remaining_pages != 0:
            range_list.append(range(pages * complete_parts, pages * complete_parts + remaining_pages))

    for one_range in range_list:
        one_pdf = pdf.pages[one_range]
        one_pdf_writer = PdfWriter()
        for one_pdf_page in one_pdf:
            one_pdf_writer.add_page(one_pdf_page)

        one_output_buffer = BytesIO()
        one_pdf_writer.write(one_output_buffer)

        result.append(one_output_buffer.getvalue())

    return result




def pdf_split_by_part(
        source: str | Path | bytes,
        parts: int,
        mode: Literal['front', 'back'] = 'back'
) -> list[bytes]:
    """
    将pdf按份划分
    :param source: 源数据
    :param parts: 按照页数进行切分，最后不足指定值的页也会被切分
    :param mode: 'front'-前面的部分可能会少页, 'back'-后面的部分可能会少页
    :return:
    """
    result = []
    source = __change_bytes_2_io(source)

    pdf = PdfReader(source)

    pdf_page_len = len(pdf.pages)

    # 当pdf页数少于份数时，每份一页
    if pdf_page_len // parts == 0:
        for page in pdf.pages:
            pdf_writer = PdfWriter()
            pdf_writer.add_page(page)
            output_buffer = BytesIO()
            pdf_writer.write(output_buffer)
            result.append(output_buffer.getvalue())

        return result

    avg_page_num = pdf_page_len // parts

    remaining_page_num = pdf_page_len % parts

    page_num_per_pdf = [avg_page_num for _ in range(parts)]

    if mode == "front":
        for idx in range(parts):
            page_num_per_pdf[idx] += idx >= (pdf_page_len - remaining_page_num)
    else:
        for idx in range(parts):
            page_num_per_pdf[idx] += idx < remaining_page_num

    start_idx = 0

    for page_num in page_num_per_pdf:
        end_idx = start_idx + page_num

        one_pdf_writer = PdfWriter()
        one_buffer = BytesIO()
        for one_page in pdf.pages[start_idx:end_idx]:
            one_pdf_writer.add_page(one_page)

        one_pdf_writer.write(one_buffer)
        result.append(one_buffer.getvalue())
        start_idx = end_idx

    return result


def pdf_drop_pages(
        source: str | Path | bytes,
        drop_pages: list[tuple[Annotated[int, '起始页, 以1开始'], Annotated[int, '结束页（包含）']]]
) -> bytes:
    """
    :param source: 原文件的绝对路径名
    :param drop_pages: 要删除的页面范围列表 [(start1, end1), (start2, end2)...]
    :return: 去除指定页数后的pdf二进制数据
    """
    result = PdfWriter()

    if isinstance(source, bytes):
        source_io = BytesIO()
        source_io.write(source)
        source_io.seek(0)
        source = source_io

    pdf = PdfReader(source)

    drop_pages_list = set()

    for page_begin, page_end in drop_pages:
        for page_idx in range(page_begin, page_end + 1):
            drop_pages_list.add(page_idx)


    for idx, one_page in enumerate(pdf.pages, start=1):
        if idx not in drop_pages_list:
            result.add_page(one_page)

    result_io = BytesIO()
    result.write(result_io)

    return result_io.getvalue()


def convert_to_pdf(source_path: Path | str, out_dir: Path | str, clean_source: bool = False) -> bytes:
    """
    使用LibreOffice将Office文档转换为PDF
    可以转换的类型有 .pdf .docx .doc .xls .xlsx .png .jpg .ppt .pptx
    :param source_path: 原文件的绝对路径
    :param out_dir: 输出文件的目录
    :param clean_source: 清理源文件
    :return:
    """
    source_path = Path(source_path)

    file_suffix = source_path.suffix
    if file_suffix not in ['.pdf', '.docx', '.doc', '.xls', '.xlsx', '.png', '.jpg', '.ppt', '.pptx']:
        raise TypeError(
            f"待转换的文件不是正确的类型：.pdf', '.docx', '.doc', '.xls', '.xlsx', '.png', '.jpg', '.ppt', '.pptx'")

    try:
    # 确保输出目录存在,真正的文件保存在out_dir/uuid下
        out_dir = Path(out_dir) / uuid1().hex
        out_dir.mkdir(parents=True, exist_ok=True)

        # 如果是pdf直接返回
        if file_suffix == '.pdf':
            with open(source_path, 'rb') as f:
                result_bytes = f.read()
            if clean_source:
                source_path.unlink(missing_ok=True)
            return result_bytes

        # 构建转换命令
        out_dir_str = out_dir.absolute().as_posix()
        source_path_str = source_path.absolute().as_posix()

        cmd = [
            'soffice',
            '--headless',
            '--convert-to', 'pdf',
            '--outdir', f'{out_dir_str}',
            f'{source_path_str}'
        ]


        # 执行转换
        result = subprocess.run(cmd, check=True, capture_output=True, text=True, timeout=3600)
        if result.stderr:
            raise RuntimeError(result.stderr)

        # 获取生成的PDF路径
        output_path = (out_dir.absolute() / os.listdir(out_dir)[0])

        with open(output_path, 'rb') as f:
            result_bytes = f.read()

        if clean_source:
            source_path.unlink(missing_ok=True)

        shutil.rmtree(out_dir)
    except Exception as e:
        shutil.rmtree(out_dir)
        raise Exception(format_exc())

    return result_bytes


async def convert_to_pdf_async(source_path: Path | str, out_dir: Path | str, clean_source: bool = False) -> bytes:
    """
    使用LibreOffice将Office文档转换为PDF
    可以转换的类型有 .pdf .docx .doc .xls .xlsx .png .jpg .ppt .pptx
    :param source_path: 原文件的绝对路径
    :param out_dir: 输出文件的目录
    :param clean_source: 清理源文件
    :return:
    """
    source_path = Path(source_path)

    file_suffix = source_path.suffix

    if file_suffix not in ['.pdf', '.docx', '.doc', '.xls', '.xlsx', '.png', '.jpg', '.ppt', '.pptx']:
        raise TypeError(
            f"待转换的文件不是正确的类型：.pdf', '.docx', '.doc', '.xls', '.xlsx', '.png', '.jpg', '.ppt', '.pptx'")
    try:
        # 确保输出目录存在,真正的文件保存在out_dir/uuid下
        out_dir = Path(out_dir) / uuid1().hex
        out_dir.mkdir(parents=True, exist_ok=True)


        # 如果是pdf直接返回
        if file_suffix == '.pdf':
            with open(source_path, 'rb') as f:
                result_bytes = f.read()
            if clean_source:
                source_path.unlink(missing_ok=True)
            return result_bytes

        # 构建转换命令
        out_dir_str = out_dir.absolute().as_posix()
        source_path_str = source_path.absolute().as_posix()

        cmd = [
            'soffice',
            '--headless',
            '--convert-to', 'pdf',
            '--outdir', f'{out_dir_str}',
            f'{source_path_str}'
        ]

        # 执行转换
        process = await asyncio.create_subprocess_exec(*cmd, stdout=asyncio.subprocess.PIPE, stderr=asyncio.subprocess.PIPE)
        try:
            await asyncio.wait_for(process.wait(), timeout=3600)
        except asyncio.TimeoutError:
            process.terminate()
            await process.wait()  # 等待真正退出
            raise TimeoutError("Command timed out")


        # 获取生成的PDF路径
        output_path = (out_dir.absolute() / os.listdir(out_dir)[0])

        with open(output_path, 'rb') as f:
            result_bytes = f.read()

        if clean_source:
            source_path.unlink(missing_ok=True)

        shutil.rmtree(out_dir)

        return result_bytes
    except Exception as e:
        shutil.rmtree(out_dir)
        raise Exception(format_exc())


def __change_bytes_2_io(source: bytes | Path | str) -> BytesIO | Path | str:
    """

    :param source:
    :return:
    """
    if isinstance(source, bytes):
        source_io = BytesIO()
        source_io.write(source)
        source_io.seek(0)
        source = source_io

    return source

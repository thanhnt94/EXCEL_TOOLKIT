# Đường dẫn: excel_toolkit/utils/compressor_engine_spire.py
# Tên cũ: image_compressor_spire_api.py
# Phiên bản 1.3 - Nâng cấp thuật toán nén và thêm tùy chọn linh hoạt
# Ngày cập nhật: 2025-10-05

__version__ = "1.3.0"

from dataclasses import dataclass, field
from typing import Dict, Optional, Tuple

from spire.xls import *
from spire.xls.common import *
import io
import logging
import os
import shutil
import tempfile
import uuid

import pythoncom
import win32com.client
from PIL import Image


@dataclass
class CompressionOptions:
    """Tập hợp các tham số điều chỉnh hành vi nén hình ảnh."""

    max_size_kb: int = 300
    max_dimensions: Tuple[int, int] = (1200, 1200)
    min_dimensions: Tuple[int, int] = (320, 320)
    min_quality: int = 35
    max_quality: int = 90
    quality_step: int = 5
    allow_downscaling: bool = True
    downscale_step: float = 0.85
    max_downscale_iterations: int = 4
    progressive_jpeg: bool = True
    prefer_png: bool = False
    convert_png_to_jpeg: bool = True
    alpha_replacement_color: Tuple[int, int, int] = (255, 255, 255)
    png_compress_level: int = 9
    preserve_excel_dimensions: bool = True
    skip_small_images_kb: int = 0
    extra_metadata: Dict[str, str] = field(default_factory=dict)

    def __post_init__(self) -> None:
        self.max_quality = max(min(self.max_quality, 100), 10)
        self.min_quality = max(min(self.min_quality, self.max_quality), 5)
        self.quality_step = max(1, self.quality_step)
        self.max_size_kb = max(10, self.max_size_kb)
        self.png_compress_level = min(max(self.png_compress_level, 0), 9)
        self.downscale_step = min(max(self.downscale_step, 0.1), 0.95)
        self.max_downscale_iterations = max(0, self.max_downscale_iterations)
        self.skip_small_images_kb = max(0, self.skip_small_images_kb)


def _normalize_output_path(output_stub: str, extension: str) -> str:
    base, _ = os.path.splitext(output_stub)
    return f"{base}.{extension.lower()}"


def _resize_image(img: Image.Image, max_dimensions: Tuple[int, int]) -> Image.Image:
    max_width, max_height = max_dimensions
    width, height = img.size
    if max_width <= 0 or max_height <= 0:
        return img

    ratio = min(max_width / width, max_height / height, 1)
    if ratio >= 1:
        return img

    new_size = (max(int(width * ratio), 1), max(int(height * ratio), 1))
    return img.resize(new_size, Image.Resampling.LANCZOS)


def _downscale_image(img: Image.Image, options: CompressionOptions) -> Optional[Image.Image]:
    width, height = img.size
    target_width = max(int(width * options.downscale_step), options.min_dimensions[0])
    target_height = max(int(height * options.downscale_step), options.min_dimensions[1])

    if target_width >= width and target_height >= height:
        return None

    if target_width < 1 or target_height < 1:
        return None

    return img.resize((target_width, target_height), Image.Resampling.LANCZOS)


def _optimize_image(
    input_path: str,
    output_stub: str,
    options: CompressionOptions,
) -> Optional[Tuple[str, int, int, str, float]]:
    """Tối ưu hóa hình ảnh và trả về kích thước mới (width, height, format, size_kb)."""

    try:
        with Image.open(input_path) as original_img:
            img = _resize_image(original_img, options.max_dimensions)
            has_alpha = (
                img.mode in ("RGBA", "LA")
                or (img.mode == "P" and "transparency" in img.info)
            )

            target_format = "PNG" if options.prefer_png else "JPEG"
            if has_alpha and not options.convert_png_to_jpeg:
                target_format = "PNG"
            elif has_alpha and options.convert_png_to_jpeg:
                rgb_background = Image.new("RGB", img.size, options.alpha_replacement_color)
                rgb_background.paste(img, mask=img.split()[-1])
                img = rgb_background
                target_format = "JPEG"
            elif img.mode not in ("RGB", "L") and target_format == "JPEG":
                img = img.convert("RGB")

            if target_format == "PNG":
                buffer = io.BytesIO()
                img.save(
                    buffer,
                    format="PNG",
                    optimize=True,
                    compress_level=options.png_compress_level,
                )
                output_path = _normalize_output_path(output_stub, "png")
                with open(output_path, "wb") as output_file:
                    output_file.write(buffer.getvalue())
                size_kb = len(buffer.getvalue()) / 1024
                return output_path, img.width, img.height, target_format, size_kb

            best_payload = None
            last_payload = None
            current_img = img

            for _ in range(max(1, options.max_downscale_iterations + 1)):
                quality = options.max_quality
                while quality >= options.min_quality:
                    buffer = io.BytesIO()
                    current_img.save(
                        buffer,
                        format="JPEG",
                        quality=quality,
                        optimize=True,
                        progressive=options.progressive_jpeg,
                    )
                    data = buffer.getvalue()
                    size_kb = len(data) / 1024
                    last_payload = (data, current_img.size, size_kb, quality)
                    if size_kb <= options.max_size_kb:
                        best_payload = last_payload
                        break
                    quality -= options.quality_step

                if best_payload or not options.allow_downscaling:
                    break

                downsized = _downscale_image(current_img, options)
                if downsized is None:
                    break
                current_img = downsized

            payload = best_payload or last_payload
            if payload is None:
                return None

            data, (width, height), size_kb, _ = payload
            output_path = _normalize_output_path(output_stub, "jpg")
            with open(output_path, "wb") as output_file:
                output_file.write(data)

            return output_path, width, height, "JPEG", size_kb
    except Exception as error:
        logging.error(f"Lỗi khi tối ưu hóa hình ảnh: {error}")
        return None


def compress_images(
    file_path: str,
    max_size_kb: int = 300,
    *,
    options: Optional[CompressionOptions] = None,
    **option_overrides,
) -> bool:
    """Nén hình ảnh trong tệp Excel bằng Spire.Xls với nhiều tùy chọn linh hoạt.

    Tham số ``max_size_kb`` được giữ lại cho khả năng tương thích với phiên bản
    cũ. Để cấu hình sâu hơn, truyền ``options`` hoặc các tham số bổ sung khớp
    với :class:`CompressionOptions` thông qua ``option_overrides``.
    """
    logging.info("Bắt đầu nén ảnh bằng engine Spire.Xls...")
    if options and option_overrides:
        logging.warning("option_overrides bị bỏ qua vì đã cung cấp options tùy chỉnh.")

    if options is None:
        options_dict = {"max_size_kb": max_size_kb}
        options_dict.update(option_overrides)
        options = CompressionOptions(**options_dict)
    else:
        if max_size_kb is not None:
            options.max_size_kb = max_size_kb

    logging.debug("Thông số nén sử dụng: %s", options)
    temp_dir = tempfile.mkdtemp()
    compressed_dir = os.path.join(temp_dir, "compressed")
    os.makedirs(compressed_dir, exist_ok=True)

    try:
        workbook = Workbook()
        workbook.LoadFromFile(file_path)
        
        images_to_replace = []
        
        for sheet_index in range(workbook.Worksheets.Count):
            sheet = workbook.Worksheets[sheet_index]
            pic_count = sheet.Pictures.Count
            
            if pic_count > 0:
                logging.debug(f"  -> Đã tìm thấy {pic_count} ảnh trong sheet '{sheet.Name}'.")
            
            for i in range(pic_count):
                try:
                    pic = sheet.Pictures[i]
                    temp_filename = f"excel_img_{uuid.uuid4()}"
                    img_path = os.path.join(temp_dir, f"{temp_filename}.png")
                    
                    pic.Picture.Save(img_path)
                    
                    if os.path.exists(img_path) and os.path.getsize(img_path) > 0:
                        original_size_kb = os.path.getsize(img_path) / 1024

                        if (
                            options.skip_small_images_kb
                            and original_size_kb <= options.skip_small_images_kb
                        ):
                            logging.debug(
                                "    -> Bỏ qua ảnh %.1fKB vì nhỏ hơn ngưỡng %.1fKB",
                                original_size_kb,
                                options.skip_small_images_kb,
                            )
                            continue

                        output_stub = os.path.join(
                            compressed_dir, f"compressed_{temp_filename}"
                        )

                        result = _optimize_image(img_path, output_stub, options)

                        if result:
                            (
                                compressed_path,
                                image_width,
                                image_height,
                                target_format,
                                compressed_size_kb,
                            ) = result
                            image_info = {
                                'compressed_path': compressed_path,
                                'sheet_name': sheet.Name,
                                'left': pic.Left,
                                'top': pic.Top,
                                'excel_width': pic.Width,
                                'excel_height': pic.Height,
                                'image_width': image_width,
                                'image_height': image_height,
                                'format': target_format,
                            }
                            images_to_replace.append(image_info)

                            logging.info(
                                "    -> Đã nén ảnh thành công: %.1fKB -> %.1fKB (%s)",
                                original_size_kb,
                                compressed_size_kb,
                                target_format,
                            )
                        else:
                            logging.warning(
                                "    -> Không thể tối ưu hóa ảnh %.1fKB trong sheet '%s'",
                                original_size_kb,
                                sheet.Name,
                            )

                except Exception as e:
                    logging.warning(f"Lỗi khi xử lý ảnh trong sheet '{sheet.Name}': {str(e)}")
                    continue

        if images_to_replace:
            logging.info(f"Đã tối ưu hóa {len(images_to_replace)} ảnh. Bắt đầu thay thế...")
            
            # SỬA LỖI: Không sao chép file nữa, làm việc trực tiếp trên file_path
            
            pythoncom.CoInitialize()
            try:
                excel = win32com.client.Dispatch("Excel.Application")
                excel.Visible = False
                excel.DisplayAlerts = False
                
                workbook_win32 = excel.Workbooks.Open(os.path.abspath(file_path))
                
                for img_info in images_to_replace:
                    sheet_name = img_info['sheet_name']
                    sheet = workbook_win32.Sheets(sheet_name)
                    
                    try:
                        for shape in sheet.Shapes:
                            if (abs(shape.Left - img_info['left']) < 5 and 
                                abs(shape.Top - img_info['top']) < 5 and 
                                shape.Type == 13): # msoPicture
                                shape.Delete()
                                logging.debug(f"    -> Đã xóa ảnh gốc trên sheet '{sheet_name}'.")
                                break
                    except Exception as e:
                        logging.warning(f"Không thể xóa ảnh gốc: {str(e)}")
                    
                    target_width = (
                        img_info['excel_width']
                        if options.preserve_excel_dimensions
                        else img_info['image_width']
                    )
                    target_height = (
                        img_info['excel_height']
                        if options.preserve_excel_dimensions
                        else img_info['image_height']
                    )

                    sheet.Shapes.AddPicture(
                        Filename=img_info['compressed_path'],
                        LinkToFile=False,
                        SaveWithDocument=True,
                        Left=img_info['left'],
                        Top=img_info['top'],
                        Width=target_width,
                        Height=target_height
                    )
                    logging.debug(f"    -> Đã chèn ảnh đã nén vào sheet '{sheet_name}'.")

                workbook_win32.Save()
                workbook_win32.Close()
                excel.Quit()
            except Exception as e:
                logging.error(f"Lỗi khi thay thế hình ảnh: {str(e)}")
            finally:
                pythoncom.CoUninitialize()
        
        logging.info("Hoàn tất nén ảnh bằng engine Spire.Xls.")
        return True
        
    except Exception as e:
        logging.error(f"Lỗi nghiêm trọng trong quá trình nén ảnh với Spire.Xls: {str(e)}")
        return False
    finally:
        try:
            shutil.rmtree(temp_dir)
            logging.debug("Đã xóa thư mục tạm thời.")
        except Exception:
            logging.warning("Không thể xóa thư mục tạm thời.")


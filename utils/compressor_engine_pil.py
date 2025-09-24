# Đường dẫn: excel_toolkit/utils/compressor_engine_pil.py
# Tên cũ: image_compressor_api.py
# Phiên bản 1.8 - Vô hiệu hóa kiểm tra xoay/khóa tỉ lệ theo yêu cầu
# Ngày cập nhật: 2025-09-15

import os
import time
import uuid
import logging
from dataclasses import dataclass, replace
from typing import Dict, Optional, Tuple

import pythoncom
from PIL import Image, ImageGrab

# --- Hằng số Office/Excel ---
xlScreen = 1
xlBitmap = 2
msoPicture = 13
msoLinkedPicture = 11
msoGroup = 6

msoBringToFront = 0
msoSendToBack = 1
msoBringForward = 2
msoSendBackward = 3
msoTrue = -1
msoFalse = 0

def _doevents_pulse():
    """Bơm message queue để COM không treo."""
    try:
        pythoncom.PumpWaitingMessages()
    except Exception:
        pass

def _copy_shape_to_image(shape, timeout_sec=3.0, sleep_step=0.05):
    """
    Copy shape -> clipboard (bitmap) -> trả về PIL.Image.
    Có timeout để tránh treo khi clipboard bận.
    """
    logging.debug("    -> Bắt đầu sao chép ảnh vào clipboard...")
    shape.api.CopyPicture(Appearance=xlScreen, Format=xlBitmap)
    logging.debug("    -> Đã sao chép vào clipboard. Bắt đầu lấy ảnh từ clipboard...")
    deadline = time.time() + timeout_sec
    while time.time() < deadline:
        _doevents_pulse()
        try:
            clip = ImageGrab.grabclipboard()
            if isinstance(clip, Image.Image):
                logging.debug("    -> Lấy ảnh từ clipboard thành công.")
                return clip
            else:
                logging.debug("    -> Clipboard chưa chứa ảnh. Đang đợi...")
        except Exception as e:
            logging.debug(f"    -> Lỗi khi lấy clipboard: {e}. Đang thử lại...")
        time.sleep(sleep_step)
    logging.warning(f"Clipboard không trả về ảnh sau {timeout_sec} giây.")
    return None

def _snapshot_shape_props(shape):
    """
    Lấy ảnh chụp thuộc tính cần bảo toàn để khôi phục sau:
    vị trí, kích thước, xoay, khoá tỉ lệ, placement, tên, visible, alt text, hyperlink, z-order.
    """
    api = shape.api
    props = {
        'name': shape.name,
        'left': shape.left,
        'top': shape.top,
        'width': shape.width,
        'height': shape.height,
        'rotation': getattr(api, 'Rotation', 0),
        'lock_aspect': getattr(api, 'LockAspectRatio', False),
        'placement': getattr(api, 'Placement', None),  # xlMove/xlMoveAndSize/xlFreeFloating
        'visible': getattr(api, 'Visible', True),
        'alt_text': getattr(api, 'AlternativeText', ''),
        'zpos': getattr(api, 'ZOrderPosition', None),
        'hyperlink': None,
    }
    # Hyperlink (nếu có)
    try:
        hl = getattr(shape, 'hyperlink', None)
        # xlwings wrapper có .hyperlink hoặc shape.api.Hyperlink
        if hl and (getattr(hl, 'address', None) or getattr(hl, 'sub_address', None)):
            props['hyperlink'] = {
                'address': getattr(hl, 'address', None),
                'sub_address': getattr(hl, 'sub_address', None),
                'screen_tip': getattr(hl, 'screen_tip', None),
                'text_to_display': getattr(hl, 'text_to_display', None),
            }
        else:
            # Thử COM thuần
            hla = getattr(api, 'Hyperlink', None)
            if hla and (getattr(hla, 'Address', None) or getattr(hla, 'SubAddress', None)):
                props['hyperlink'] = {
                    'address': getattr(hla, 'Address', None),
                    'sub_address': getattr(hla, 'SubAddress', None),
                    'screen_tip': getattr(hla, 'ScreenTip', None),
                    'text_to_display': getattr(hla, 'TextToDisplay', None),
                }
    except Exception:
        pass
    return props

def _normalize_lock_value(value):
    """Chuẩn hoá LockAspectRatio sang tuple (bool, hằng số mso)."""

    if value is None:
        return None, None

    if isinstance(value, bool):
        return value, msoTrue if value else msoFalse

    try:
        intval = int(value)
    except (TypeError, ValueError):
        return None, None

    if intval == msoTrue:
        return True, msoTrue
    if intval == msoFalse:
        return False, msoFalse

    # Một số phiên bản COM trả về 1 thay vì -1
    if intval == 1:
        return True, msoTrue
    if intval == 0:
        return False, msoFalse

    return None, None


def _apply_props_to_picture(pic, props):
    """
    Áp lại các thuộc tính đã chụp cho ảnh mới chèn.
    """
    api = pic.api
    logging.debug(f"    -> Đang khôi phục thuộc tính cho ảnh '{props['name']}'...")
    try:
        logging.debug("    -> Áp dụng tên...")
        pic.name = props['name']
    except Exception as e:
        logging.warning(f"    -> Lỗi khi áp dụng tên: {e}")

    lock_bool, lock_mso = _normalize_lock_value(props.get('lock_aspect'))

    try:
        logging.debug("    -> Áp dụng vị trí...")
        pic.left = props['left']
        pic.top = props['top']
    except Exception as e:
        logging.warning(f"    -> Lỗi khi áp dụng vị trí: {e}")

    # Nếu ảnh cũ bị khoá tỉ lệ, Excel sẽ không cho đặt width/height độc lập.
    # Ta tạm mở khoá (nếu cần) để đảm bảo kích thước khớp tuyệt đối.
    temporary_unlock = False
    if lock_bool:
        try:
            current_lock = pic.api.LockAspectRatio
            if current_lock not in (msoFalse, 0):
                pic.api.LockAspectRatio = msoFalse
                temporary_unlock = True
        except Exception as e:
            logging.debug(f"    -> Không thể mở khoá tỉ lệ tạm thời: {e}")

    try:
        logging.debug("    -> Áp dụng kích thước...")
        pic.width = props['width']
        pic.height = props['height']
    except Exception as e:
        logging.warning(f"    -> Lỗi khi áp dụng kích thước: {e}")

    if temporary_unlock and lock_mso is not None:
        try:
            pic.api.LockAspectRatio = lock_mso
        except Exception as e:
            logging.warning(f"    -> Không thể khôi phục trạng thái khoá tỉ lệ ban đầu: {e}")

    try:
        if props['rotation']:
            logging.debug(f"    -> Áp dụng xoay với giá trị: {props['rotation']}...")
            api.Rotation = props['rotation']
    except Exception as e:
        logging.warning(f"    -> Lỗi khi áp dụng xoay: {e}")

    if lock_mso is not None and not temporary_unlock:
        try:
            logging.debug("    -> Áp dụng khóa tỉ lệ...")
            api.LockAspectRatio = lock_mso
        except Exception as e:
            logging.warning(f"    -> Lỗi khi áp dụng khóa tỉ lệ: {e}")

    try:
        logging.debug("    -> Áp dụng placement...")
        if props['placement'] is not None:
            api.Placement = props['placement']
    except Exception as e:
        logging.warning(f"    -> Lỗi khi áp dụng placement: {e}")

    try:
        logging.debug("    -> Áp dụng visible...")
        api.Visible = props['visible']
    except Exception as e:
        logging.warning(f"    -> Lỗi khi áp dụng visible: {e}")

    try:
        logging.debug("    -> Áp dụng alternative text...")
        api.AlternativeText = props['alt_text'] or ''
    except Exception as e:
        logging.warning(f"    -> Lỗi khi áp dụng alternative text: {e}")

    # Hyperlink
    try:
        logging.debug("    -> Áp dụng hyperlink...")
        hl = props.get('hyperlink')
        if hl and (hl.get('address') or hl.get('sub_address')):
            pic.sheet.api.Hyperlinks.Add(
                Anchor=pic.api, Address=hl.get('address'), SubAddress=hl.get('sub_address'),
                ScreenTip=hl.get('screen_tip'), TextToDisplay=hl.get('text_to_display')
            )
    except Exception as e:
        logging.warning(f"    -> Lỗi khi áp dụng hyperlink: {e}")

@dataclass(frozen=True)
class CompressionOptions:
    """Tập hợp các tùy chọn nén ảnh.

    Các trường được giữ đơn giản để vẫn tương thích với tham số cũ:
    - ``mode``: chiến lược chọn định dạng đầu ra.
    - ``jpeg_quality``/``jpeg_optimize``/``jpeg_progressive``: cấu hình JPEG.
    - ``jpeg_background``: màu nền dùng để làm phẳng lớp alpha khi ép sang JPEG.
    - ``png_optimize``/``png_colors``/``png_compress_level``: cấu hình PNG.
    - ``webp_quality``/``webp_lossless``: cấu hình WebP (nếu được chọn).
    - ``keep_dpi``: đặt DPI cho ảnh đầu ra; ``None`` để giữ nguyên.
    - ``max_width``/``max_height``: thu nhỏ ảnh khi lớn hơn kích thước chỉ định.
    - ``resize_filter``: bộ lọc nội suy (Pillow constant, ví dụ ``Image.LANCZOS``).
    - ``strip_metadata``: loại bỏ EXIF/IPTC để giảm dung lượng.
    """

    mode: str = "auto"
    jpeg_quality: int = 70
    jpeg_optimize: bool = True
    jpeg_progressive: bool = True
    jpeg_background: Optional[Tuple[int, int, int]] = (255, 255, 255)
    png_optimize: bool = True
    png_colors: Optional[int] = 256
    png_compress_level: int = 9
    webp_quality: int = 75
    webp_lossless: bool = False
    keep_dpi: Optional[int] = 96
    max_width: Optional[int] = None
    max_height: Optional[int] = None
    resize_filter: int = Image.LANCZOS
    strip_metadata: bool = True

    @staticmethod
    def from_legacy(quality: int = 70, mode: str = "auto", keep_dpi: Optional[int] = 96, **kwargs) -> "CompressionOptions":
        """Tạo ``CompressionOptions`` từ tham số cũ của ``compress_images``."""

        opts = CompressionOptions(mode=mode, keep_dpi=keep_dpi, **kwargs)
        if mode in ("jpeg", "auto"):
            opts = replace(opts, jpeg_quality=quality)
        elif mode == "png":
            # Với PNG ta map quality -> số màu nếu chưa được set thủ công
            if "png_colors" not in kwargs:
                colors = max(16, min(256, int(quality / 100 * 256)))
                opts = replace(opts, png_colors=colors)
        elif mode == "webp":
            opts = replace(opts, webp_quality=quality)
        return opts


def _prepare_image(img: Image.Image, opts: CompressionOptions) -> Tuple[Image.Image, str, Dict[str, object]]:
    """Chuyển đổi ảnh gốc sang định dạng và tham số lưu phù hợp với ``opts``."""

    img = img.copy()
    if opts.strip_metadata:
        try:
            img.info.pop("exif", None)
            if hasattr(img, "getexif"):
                exif = img.getexif()
                exif.clear()
        except Exception:
            pass

    if opts.max_width or opts.max_height:
        max_w = opts.max_width or img.width
        max_h = opts.max_height or img.height
        if img.width > max_w or img.height > max_h:
            ratio_w = max_w / img.width
            ratio_h = max_h / img.height
            ratio = min(ratio_w, ratio_h)
            new_size = (int(img.width * ratio), int(img.height * ratio))
            if new_size[0] > 0 and new_size[1] > 0:
                img = img.resize(new_size, opts.resize_filter)

    mode = opts.mode.lower()
    save_kwargs: Dict[str, object] = {}
    if opts.keep_dpi:
        save_kwargs["dpi"] = (opts.keep_dpi, opts.keep_dpi)

    # Tự chọn định dạng nếu để auto.
    if mode == "auto":
        if img.mode in ("RGBA", "LA", "P") and getattr(img, "info", {}).get("transparency") is not None:
            fmt = "PNG"
        else:
            fmt = "JPEG"
    elif mode == "jpeg":
        fmt = "JPEG"
    elif mode == "png":
        fmt = "PNG"
    elif mode == "webp":
        fmt = "WEBP"
    else:
        logging.warning(f"Mode '{opts.mode}' không hợp lệ, sử dụng 'auto'.")
        return _prepare_image(img, replace(opts, mode="auto"))

    if fmt == "JPEG":
        if img.mode not in ("RGB", "L"):
            background = opts.jpeg_background or (255, 255, 255)
            if img.mode in ("RGBA", "LA"):
                alpha = img.split()[-1]
                base = Image.new("RGB", img.size, background)
                base.paste(img.convert("RGB"), mask=alpha)
                img = base
            else:
                img = img.convert("RGB")
        save_kwargs.update(
            quality=max(1, min(100, opts.jpeg_quality)),
            optimize=opts.jpeg_optimize,
            progressive=opts.jpeg_progressive,
        )
    elif fmt == "PNG":
        if img.mode not in ("RGBA", "LA", "RGB", "L"):
            img = img.convert("RGBA")
        if opts.png_colors and img.mode not in ("L", "P"):
            try:
                img = img.convert("P", palette=Image.ADAPTIVE, colors=opts.png_colors)
            except Exception:
                pass
        save_kwargs.update(optimize=opts.png_optimize, compress_level=max(0, min(9, opts.png_compress_level)))
    elif fmt == "WEBP":
        if img.mode not in ("RGB", "RGBA", "L"):
            img = img.convert("RGB")
        save_kwargs.update(quality=max(1, min(100, opts.webp_quality)), lossless=opts.webp_lossless, method=6)

    return img, fmt, save_kwargs


def _export_and_replace(shape, sheet, quality=70, mode='auto', keep_dpi=96, *, options: Optional[CompressionOptions] = None, **extra):
    """
    Trích xuất shape -> nén -> xoá shape cũ -> chèn lại ảnh -> khôi phục props.
    """
    logging.debug("    -> Lấy thuộc tính của shape để khôi phục...")
    props = _snapshot_shape_props(shape)

    img = _copy_shape_to_image(shape)
    if img is None:
        logging.warning(f"Clipboard không trả ảnh cho '{props['name']}'. Bỏ qua.")
        return None

    logging.debug("    -> Bắt đầu xử lý và lưu ảnh tạm thời...")

    opts = options or CompressionOptions.from_legacy(quality=quality, mode=mode, keep_dpi=keep_dpi, **extra)
    try:
        prepared_img, fmt, save_kwargs = _prepare_image(img, opts)
    except Exception as prep_err:
        logging.error(f"    -> Lỗi khi chuẩn bị ảnh '{props['name']}': {prep_err}")
        return None

    tmp_dir = os.path.join(os.getcwd(), "_tmp_excel_img")
    os.makedirs(tmp_dir, exist_ok=True)
    tmp_path = os.path.join(tmp_dir, f"{uuid.uuid4().hex}.{fmt.lower()}")

    try:
        logging.debug(f"    -> Lưu ảnh tạm thời ở định dạng {fmt} tại '{tmp_path}' với tham số {save_kwargs}...")
        prepared_img.save(tmp_path, format=fmt, **save_kwargs)
        logging.debug("    -> Đã lưu ảnh tạm thời thành công.")
    except Exception as e:
        logging.error(f"    -> Lỗi khi lưu ảnh tạm thời: {e}")
        return None
    
    logging.debug("    -> Bắt đầu xóa shape cũ...")
    # Xoá shape cũ
    try:
        shape.delete()
        logging.debug("    -> Đã xóa shape cũ thành công.")
    except Exception as e:
        logging.warning(f"Không xoá được shape cũ '{props['name']}': {e}")
        return None

    logging.debug("    -> Bắt đầu chèn ảnh mới...")
    # Chèn ảnh mới
    pic = sheet.pictures.add(tmp_path, left=props['left'], top=props['top'])
    logging.debug(f"    -> Đã chèn ảnh mới thành công với tên '{pic.name}'.")

    # Cố gắng giữ nguyên kích thước (tránh scale theo DPI)
    try:
        pic.width = props['width']
        pic.height = props['height']
    except Exception:
        pass

    logging.debug("    -> Bắt đầu áp dụng lại thuộc tính...")
    # Áp thuộc tính lại
    _apply_props_to_picture(pic, props)
    logging.debug("    -> Đã áp dụng lại thuộc tính thành công.")

    # Xóa file tạm
    try:
        os.remove(tmp_path)
        logging.debug("    -> Đã xóa file ảnh tạm thời.")
    except Exception as e:
        logging.warning(f"Không thể xóa file ảnh tạm thời '{tmp_path}': {e}")

    # Trả về tên shape mới (tên có thể đổi nếu trùng)
    return pic.name
    
def _reorder_zorder_exact(sheet, saved_order_back_to_front):
    """
    Khôi phục thứ tự chồng lớp chính xác.
    Cách làm: duyệt theo thứ tự từ “phía sau” -> “phía trước”, mỗi shape gọi BringToFront.
    Kết quả: các shape sẽ xếp đúng như saved_order_back_to_front.
    """
    for nm in saved_order_back_to_front:
        try:
            shp = sheet.shapes[nm]
            shp.api.ZOrder(msoBringToFront)
        except Exception:
            # Có thể tên bị đổi sau khi chèn lại; bỏ qua nếu không còn tồn tại.
            pass

def compress_images(
    wb,
    quality: int = 70,
    mode: str = 'auto',
    keep_dpi: Optional[int] = 96,
    *,
    compression_options: Optional[CompressionOptions] = None,
    **kwargs,
):
    """
    Nén tất cả ảnh (msoPicture/msoLinkedPicture) trong workbook:
    - Bảo toàn vị trí, kích thước, xoay, tỉ lệ, placement, tên, visible, alt text, hyperlink.
    - Khôi phục z-order để textbox/shape khác vẫn đè đúng.
    - Bỏ qua nhóm (msoGroup) để tránh phá vỡ group.

    Tham số ``kwargs`` cho phép tuỳ biến sâu hơn (ví dụ ``max_width``, ``max_height``,
    ``png_colors``, ``jpeg_progressive``...). Nếu đã khởi tạo ``compression_options`` thì
    các giá trị trong ``kwargs`` sẽ được bỏ qua.
    """
    excel = wb.app.api
    prev_screen = excel.ScreenUpdating
    prev_alerts = excel.DisplayAlerts
    prev_calc = excel.Calculation
    excel.ScreenUpdating = False
    excel.DisplayAlerts = False
    try:
        excel.Calculation = -4135  # xlCalculationManual
    except Exception:
        pass

    total = 0
    compressed = 0
    
    # Duyệt qua từng sheet hiển thị
    options = compression_options or CompressionOptions.from_legacy(
        quality=quality,
        mode=mode,
        keep_dpi=keep_dpi,
        **kwargs,
    )

    for sheet in wb.sheets:
        if getattr(sheet.api, 'Visible', -1) != -1:
            continue

        # Lấy z-order của tất cả các shapes trên sheet
        try:
            shapes_with_z = []
            for s in sheet.shapes:
                try:
                    z = getattr(s.api, 'ZOrderPosition', None)
                    if z is not None:
                        shapes_with_z.append((z, s.name))
                except Exception:
                    pass
            shapes_with_z.sort(key=lambda x: x[0])  # Sắp xếp từ sau ra trước
            z_order_names = [nm for _, nm in shapes_with_z]
        except Exception:
            z_order_names = [s.name for s in sheet.shapes]

        shape_names = [s.name for s in sheet.shapes]
        new_names_map = {}

        for nm in shape_names:
            try:
                shp = sheet.shapes[nm]
                t = getattr(shp.api, 'Type', None)
            except Exception:
                continue

            if t in (msoPicture, msoLinkedPicture):
                total += 1
                
                # # SỬA LỖI: Kiểm tra các thuộc tính không ổn định trước khi nén
                # # Tạm thời vô hiệu hóa theo yêu cầu để nén tất cả ảnh
                # try:
                #     rotation = getattr(shp.api, 'Rotation', 0)
                #     is_locked = getattr(shp.api, 'LockAspectRatio', False)

                #     if rotation != 0 or is_locked:
                #         logging.warning(f"Bỏ qua ảnh '{nm}' trên sheet '{sheet.name}' do có góc xoay hoặc bị khóa tỉ lệ.")
                #         continue # Bỏ qua và chuyển sang shape tiếp theo
                # except Exception as prop_err:
                #     logging.warning(f"Không thể kiểm tra thuộc tính của ảnh '{nm}'. Lỗi: {prop_err}")

                logging.info(f"Đang nén ảnh '{nm}' trên sheet '{sheet.name}'...")
                try:
                    new_nm = _export_and_replace(
                        shp,
                        sheet,
                        quality=quality,
                        mode=mode,
                        keep_dpi=keep_dpi,
                        options=options,
                        **kwargs,
                    )
                    if new_nm:
                        compressed += 1
                        new_names_map[nm] = new_nm
                except Exception as e:
                    logging.warning(f"Lỗi khi nén ảnh '{nm}' ở sheet '{sheet.name}': {e}")
            else:
                logging.debug(f"Bỏ qua shape '{nm}' (loại: {t}) vì không phải ảnh.")
                pass
            
            _doevents_pulse()

        z_order_names_updated = [new_names_map.get(nm, nm) for nm in z_order_names]
        _reorder_zorder_exact(sheet, z_order_names_updated)

    logging.info(f"Hoàn tất nén ảnh. Đã nén {compressed}/{total} ảnh.")
    
    excel.ScreenUpdating = prev_screen
    excel.DisplayAlerts = prev_alerts
    try:
        excel.Calculation = prev_calc
    except Exception:
        pass
    
    return True


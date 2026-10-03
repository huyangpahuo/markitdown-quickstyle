"""生成 Tauri 应用图标(纯标准库,不依赖网络/Pillow)。

用法: python_env\\python.exe scripts\\make_icons.py
输出: src-tauri/icons/{32x32.png, 128x128.png, icon.ico}
图标内容: 一个紫色圆脸 + 两只眼睛(与应用内桌宠一致)
"""
import struct
import zlib
from pathlib import Path

FACE_R = (124, 111, 240)   # #7c6ff0
FACE_LIGHT = (168, 155, 255)
EYE = (30, 27, 46)


def draw(size: int) -> bytes:
    cx = cy = (size - 1) / 2
    r = size * 0.46
    eye_off_x, eye_off_y, eye_r = size * 0.17, -size * 0.03, size * 0.085
    mouth_y, mouth_r = cy + size * 0.16, size * 0.055

    rows = []
    for y in range(size):
        row = bytearray([0])  # PNG filter: none
        for x in range(size):
            d = ((x - cx) ** 2 + (y - cy) ** 2) ** 0.5
            if d <= r:
                t = max(0.0, 1 - (x + y) / (2 * size))  # 左上偏亮
                cr = int(FACE_R[0] + (FACE_LIGHT[0] - FACE_R[0]) * t)
                cg = int(FACE_R[1] + (FACE_LIGHT[1] - FACE_R[1]) * t)
                cb = int(FACE_R[2] + (FACE_LIGHT[2] - FACE_R[2]) * t)
                a = 255
                for sx in (-1, 1):
                    ex, ey = cx + sx * eye_off_x, cy + eye_off_y
                    if ((x - ex) ** 2 + (y - ey) ** 2) ** 0.5 <= eye_r:
                        cr, cg, cb = EYE
                if ((x - cx) ** 2 + (y - mouth_y) ** 2) ** 0.5 <= mouth_r:
                    cr, cg, cb = EYE
            else:
                # 1px 抗锯齿边缘
                if d <= r + 1:
                    cr, cg, cb, a = FACE_R[0], FACE_R[1], FACE_R[2], int(255 * (r + 1 - d))
                else:
                    cr = cg = cb = a = 0
            row += bytes((cr, cg, cb, a))
        rows.append(bytes(row))

    def chunk(tag: bytes, data: bytes) -> bytes:
        return (
            struct.pack(">I", len(data))
            + tag
            + data
            + struct.pack(">I", zlib.crc32(tag + data) & 0xFFFFFFFF)
        )

    png = b"\x89PNG\r\n\x1a\n"
    png += chunk(b"IHDR", struct.pack(">IIBBBBB", size, size, 8, 6, 0, 0, 0))
    png += chunk(b"IDAT", zlib.compress(b"".join(rows), 9))
    png += chunk(b"IEND", b"")
    return png


def make_ico(png32: bytes) -> bytes:
    # ICO 容器直接内嵌 PNG(Vista+ 支持)
    header = struct.pack("<HHH", 0, 1, 1)
    entry = struct.pack("<BBBBHHII", 32, 32, 0, 0, 1, 32, len(png32), 22)
    return header + entry + png32


def main() -> None:
    out = Path(__file__).resolve().parent.parent / "src-tauri" / "icons"
    out.mkdir(parents=True, exist_ok=True)
    png32, png128 = draw(32), draw(128)
    (out / "32x32.png").write_bytes(png32)
    (out / "128x128.png").write_bytes(png128)
    (out / "icon.ico").write_bytes(make_ico(png32))
    print(f"icons written to {out}")


if __name__ == "__main__":
    main()

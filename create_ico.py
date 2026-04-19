# -*- coding: utf-8 -*-
"""
Windows .ico 아이콘 파일 생성 (PIL 없이 순수 파이썬)
생산관리 시스템용 - 톱니바퀴 / 공장 테마
"""
import sys, struct, zlib, os, math
sys.stdout.reconfigure(encoding="utf-8", errors="replace")


def make_png(size):
    def chunk(name, data):
        crc = zlib.crc32(name + data) & 0xFFFFFFFF
        return struct.pack(">I", len(data)) + name + data + struct.pack(">I", crc)

    # 팔레트 (teal/orange 생산관리 컬러)
    DARK   = (0,   77,  64)
    TEAL   = (0,   105, 92)
    LTEAL  = (38,  166, 154)
    WHITE  = (255, 255, 255)
    ORANGE = (255, 111, 0)
    LGRAY  = (224, 242, 241)

    pixels = []
    cx = cy = size / 2
    r_out  = cx - 1

    # 톱니바퀴 파라미터
    teeth = 10
    r_inner = size * 0.30   # 톱니 안쪽 반지름
    r_outer = size * 0.42   # 톱니 바깥 반지름
    r_hole  = size * 0.13   # 가운데 구멍

    for y in range(size):
        for x in range(size):
            dx = x - cx; dy = y - cy
            dist = math.sqrt(dx*dx + dy*dy)

            if dist > r_out:
                pixels.append(None); continue

            # 배경 그라디언트
            t = dist / r_out
            r = int(TEAL[0]*(1-t) + LTEAL[0]*t)
            g = int(TEAL[1]*(1-t) + LTEAL[1]*t)
            b = int(TEAL[2]*(1-t) + LTEAL[2]*t)
            px = (r, g, b)

            # 외곽 테두리
            if dist > r_out - max(2, size*0.04):
                px = DARK

            # 톱니바퀴
            angle = math.atan2(dy, dx)
            tooth_phase = (angle * teeth / (2 * math.pi)) % 1.0
            in_tooth = tooth_phase < 0.5
            current_r = r_outer if in_tooth else r_inner + (r_outer - r_inner) * 0.55

            if dist <= current_r and dist >= r_hole:
                px = WHITE
                # 안쪽 테두리
                if dist >= current_r - max(1, size*0.018):
                    px = DARK

            # 가운데 구멍 (오렌지)
            if dist <= r_hole:
                px = ORANGE
                if dist >= r_hole - max(1, size*0.018):
                    px = DARK

            pixels.append(px)

    raw = b""
    for y in range(size):
        raw += b"\x00"
        for x in range(size):
            px = pixels[y*size+x]
            if px is None:
                raw += b"\x00\x00\x00\x00"
            else:
                raw += bytes([px[0], px[1], px[2], 255])

    ihdr = struct.pack(">IIBBBBB", size, size, 8, 6, 0, 0, 0)
    sig = b"\x89PNG\r\n\x1a\n"
    return (sig
            + chunk(b"IHDR", ihdr)
            + chunk(b"IDAT", zlib.compress(raw, 9))
            + chunk(b"IEND", b""))


def create_ico(output_path):
    sizes = [16, 32, 48, 64, 128, 256]
    png_list = []
    print("  아이콘 크기 생성 중...", flush=True)
    for s in sizes:
        print(f"    {s}x{s}...", end="", flush=True)
        png_list.append((s, make_png(s)))
        print(" OK")

    count = len(png_list)
    header = struct.pack("<HHH", 0, 1, count)
    dir_size = count * 16
    offset = 6 + dir_size
    directory = b""
    for s, png_data in png_list:
        sz = 0 if s == 256 else s
        directory += struct.pack("<BBBBHHII",
                                 sz, sz, 0, 0, 1, 32,
                                 len(png_data), offset)
        offset += len(png_data)

    ico_data = header + directory
    for _, png_data in png_list:
        ico_data += png_data

    with open(output_path, "wb") as f:
        f.write(ico_data)
    print(f"  OK: {output_path} ({len(ico_data)/1024:.1f} KB)")


if __name__ == "__main__":
    out = os.path.join(os.path.dirname(os.path.abspath(__file__)), "production.ico")
    print("Windows 아이콘(.ico) 생성 중...")
    create_ico(out)
    print("완료!")

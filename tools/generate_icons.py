from pathlib import Path
import struct
import zlib


ROOT = Path(__file__).resolve().parents[1]
ASSETS = ROOT / "assets"


def clamp(value):
    return max(0, min(255, int(value)))


def make_canvas(size):
    pixels = []
    for y in range(size):
        row = []
        mix = y / max(1, size - 1)
        for x in range(size):
            blend = (x / max(1, size - 1) + mix) / 2
            row.append(
                [
                    clamp(23 + blend * 12),
                    clamp(94 - blend * 30),
                    clamp(87 - blend * 27),
                    255
                ]
            )
        pixels.append(row)
    return pixels


def fill_rounded_rect(pixels, x0, y0, x1, y1, radius, color):
    radius_sq = radius * radius
    for y in range(y0, y1):
        for x in range(x0, x1):
            dx = 0
            dy = 0

            if x < x0 + radius:
                dx = x0 + radius - x
            elif x > x1 - radius - 1:
                dx = x - (x1 - radius - 1)

            if y < y0 + radius:
                dy = y0 + radius - y
            elif y > y1 - radius - 1:
                dy = y - (y1 - radius - 1)

            if dx and dy and dx * dx + dy * dy > radius_sq:
                continue

            pixels[y][x] = color[:]


def fill_polygon(pixels, points, color):
    min_y = max(0, min(y for _, y in points))
    max_y = min(len(pixels) - 1, max(y for _, y in points))

    for y in range(min_y, max_y + 1):
        intersections = []
        for index, (x1, y1) in enumerate(points):
            x2, y2 = points[(index + 1) % len(points)]
            if y1 == y2:
                continue
            if y < min(y1, y2) or y >= max(y1, y2):
                continue
            ratio = (y - y1) / (y2 - y1)
            intersections.append(x1 + ratio * (x2 - x1))

        intersections.sort()
        for left, right in zip(intersections[0::2], intersections[1::2]):
            for x in range(max(0, int(left)), min(len(pixels[y]), int(right) + 1)):
                pixels[y][x] = color[:]


def star_points(size):
    cx = size // 2
    cy = int(size * 0.57)
    outer = size * 0.115
    inner = outer * 0.46
    points = []
    for index in range(10):
        angle = -1.5708 + index * 0.62832
        radius = outer if index % 2 == 0 else inner
        points.append((int(cx + radius * __import__("math").cos(angle)), int(cy + radius * __import__("math").sin(angle))))
    return points


def write_png(path, pixels):
    height = len(pixels)
    width = len(pixels[0])
    raw = bytearray()

    for row in pixels:
      raw.append(0)
      for red, green, blue, alpha in row:
        raw.extend([red, green, blue, alpha])

    def chunk(tag, data):
        return (
            struct.pack(">I", len(data))
            + tag
            + data
            + struct.pack(">I", zlib.crc32(tag + data) & 0xFFFFFFFF)
        )

    png = bytearray(b"\x89PNG\r\n\x1a\n")
    png.extend(chunk(b"IHDR", struct.pack(">IIBBBBB", width, height, 8, 6, 0, 0, 0)))
    png.extend(chunk(b"IDAT", zlib.compress(bytes(raw), 9)))
    png.extend(chunk(b"IEND", b""))
    path.write_bytes(png)


def build_icon(size):
    pixels = make_canvas(size)
    folder = [255, 210, 132, 255]
    folder_deep = [248, 189, 96, 255]
    star = [24, 95, 88, 255]

    fill_rounded_rect(pixels, int(size * 0.2), int(size * 0.31), int(size * 0.79), int(size * 0.72), int(size * 0.07), folder)
    fill_rounded_rect(pixels, int(size * 0.18), int(size * 0.38), int(size * 0.82), int(size * 0.73), int(size * 0.07), folder)
    fill_rounded_rect(pixels, int(size * 0.18), int(size * 0.42), int(size * 0.82), int(size * 0.74), int(size * 0.07), folder_deep)
    fill_polygon(pixels, star_points(size), star)
    return pixels


def main():
    ASSETS.mkdir(exist_ok=True)
    for name, size in {
        "icon-192.png": 192,
        "icon-512.png": 512,
        "apple-touch-icon.png": 180,
    }.items():
        write_png(ASSETS / name, build_icon(size))


if __name__ == "__main__":
    main()

"""AI세금계산서 아이콘 - 불투명 배경으로 확실하게 생성"""
from pathlib import Path
from PIL import Image, ImageDraw, ImageFont

OUT = Path(__file__).parent / "app.ico"
# Windows 표준 ICO 사이즈
SIZES = [16, 24, 32, 48, 64, 128, 256]


def font(size):
    for name in ["malgunbd.ttf", "malgun.ttf", "arialbd.ttf", "arial.ttf"]:
        try:
            return ImageFont.truetype(name, size)
        except Exception:
            continue
    return ImageFont.load_default()


def draw_icon(s: int) -> Image.Image:
    """단일 크기 아이콘. RGB 모드로 불투명 배경."""
    # 먼저 RGBA 로 그려서 알파 블렌딩
    img = Image.new("RGBA", (s, s), (255, 255, 255, 0))
    d = ImageDraw.Draw(img)

    # 인디고 배경 (거의 전체 채움)
    bg_color = (99, 102, 241, 255)
    margin = 0  # 여백 없이 꽉 채움
    radius = max(2, s // 8) if s >= 32 else 0

    if radius > 0:
        d.rounded_rectangle(
            (margin, margin, s - 1 - margin, s - 1 - margin),
            radius=radius, fill=bg_color,
        )
    else:
        d.rectangle(
            (margin, margin, s - 1 - margin, s - 1 - margin),
            fill=bg_color,
        )

    # 원화 심볼 ₩
    if s >= 20:
        try:
            f = font(int(s * 0.7))
            text = "₩"
            bbox = d.textbbox((0, 0), text, font=f)
            tw = bbox[2] - bbox[0]
            th = bbox[3] - bbox[1]
            tx = (s - tw) // 2 - bbox[0]
            ty = (s - th) // 2 - bbox[1] - int(s * 0.05)
            d.text((tx, ty), text, fill=(255, 255, 255, 255), font=f)
        except Exception:
            pass
    else:
        # 16x16 아주 작은: 흰 수직선 3개 (₩ 느낌)
        pad = max(2, s // 5)
        bar_w = max(1, s // 10)
        for i, x in enumerate([pad, s // 2 - bar_w // 2, s - pad - bar_w]):
            d.rectangle(
                (x, pad, x + bar_w, s - 1 - pad),
                fill=(255, 255, 255, 255),
            )
        # 가로 줄
        d.rectangle(
            (pad, s // 2 - 1, s - 1 - pad, s // 2),
            fill=(255, 255, 255, 255),
        )

    # RGBA → RGB (흰 배경과 합성) — Windows 호환성
    bg = Image.new("RGB", img.size, (255, 255, 255))
    bg.paste(img, mask=img.split()[3] if img.mode == "RGBA" else None)
    return bg


def main():
    # 각 사이즈 생성
    imgs = [draw_icon(s) for s in SIZES]
    # 가장 큰 것을 기본으로 save, 나머지는 append
    imgs[-1].save(
        OUT,
        format="ICO",
        sizes=[(s, s) for s in SIZES],
    )
    # 검증
    v = Image.open(OUT)
    print(f"saved: {OUT}")
    print(f"size: {OUT.stat().st_size} bytes")
    print(f"format: {v.format}, size: {v.size}, mode: {v.mode}")
    # ico 에 포함된 사이즈 확인
    try:
        print(f"sizes in ico: {v.info.get('sizes', [])}")
    except Exception:
        pass


if __name__ == "__main__":
    main()

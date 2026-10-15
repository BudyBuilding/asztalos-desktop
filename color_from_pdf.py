from pathlib import Path
import argparse
import re
import pymupdf


CODE_RE = re.compile(
    r"^(?:[A-Z]{1,3}\d{2,4}|\d{4})$",
    re.IGNORECASE
)


def find_color_code(page):
    """
    Kódkeresés a lap felső részén.

    Kezeli többek között ezeket:
    S010
    AF10
    AF11 NEW
    6010
    9513
    9 5 1 3
    8510 MATT
    """

    r = page.rect

    # Csak az oldal felső részét vizsgáljuk,
    # mert ott van a színkód.
    search_area = pymupdf.Rect(
        r.x0,
        r.y0,
        r.x1,
        r.y0 + r.height * 0.28
    )

    text = page.get_text(
        "text",
        clip=search_area
    )

    for raw_line in text.splitlines():

        # Pl.
        # "9 5 1 3"  -> "9513"
        # "AF11 NEW" -> "AF11NEW"
        # "8510 MATT" -> "8510MATT"
        compact = re.sub(
            r"[^A-Z0-9]",
            "",
            raw_line.upper()
        )

        # Jelölések eltávolítása a kód végéről
        for suffix in ("NEW", "MATT"):
            if compact.endswith(suffix):
                compact = compact[:-len(suffix)]

        if not CODE_RE.fullmatch(compact):
            continue

        # Borítókon szereplő évszám ne legyen színkód
        if re.fullmatch(r"20\d{2}", compact):
            continue

        return compact

    return None


def get_color_clip(page):
    """
    A lap közepéből veszünk egy nagy négyzetet.

    Így biztosan kimarad:
    - keret
    - színkód
    - logó
    - alsó szöveg
    - fekete háttér

    Csak maga a szín / textúra marad.
    """

    r = page.rect

    side = min(r.width, r.height) * 0.70

    center_x = (r.x0 + r.x1) / 2
    center_y = (r.y0 + r.y1) / 2

    return pymupdf.Rect(
        center_x - side / 2,
        center_y - side / 2,
        center_x + side / 2,
        center_y + side / 2
    )


def convert_pdf(
    pdf_path: Path,
    output_dir: Path,
    target_size: int = 1200
):
    output_dir.mkdir(
        parents=True,
        exist_ok=True
    )

    doc = pymupdf.open(pdf_path)

    saved = 0

    for page_no, page in enumerate(doc, start=1):

        code = find_color_code(page)

        if code is None:
            print(
                f"SKIP  page {page_no:02d}"
            )
            continue

        clip = get_color_clip(page)

        scale = target_size / clip.width

        matrix = pymupdf.Matrix(
            scale,
            scale
        )

        pix = page.get_pixmap(
            matrix=matrix,
            clip=clip,
            colorspace=pymupdf.csRGB,
            alpha=False
        )

        output_file = (
            output_dir / f"{code}.png"
        )

        pix.save(output_file)

        print(
            f"SAVE  page {page_no:02d}"
            f" -> {output_file.name}"
        )

        saved += 1

    doc.close()

    print()
    print(f"Kész: {saved} kép")
    print(
        f"Kimenet: "
        f"{output_dir.resolve()}"
    )


def main():
    parser = argparse.ArgumentParser()

    parser.add_argument(
        "pdf",
        type=Path,
        help="Bemeneti PDF"
    )

    parser.add_argument(
        "-o",
        "--output",
        type=Path,
        default=Path("colors_only"),
        help="Kimeneti mappa"
    )

    parser.add_argument(
        "--size",
        type=int,
        default=1200,
        help="PNG mérete pixelben"
    )

    args = parser.parse_args()

    convert_pdf(
        args.pdf,
        args.output,
        args.size
    )


if __name__ == "__main__":
    main()
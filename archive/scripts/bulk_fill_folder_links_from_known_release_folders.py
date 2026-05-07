#!/usr/bin/env python3
from __future__ import annotations

import argparse
import json
import sqlite3
from pathlib import Path


DEFAULT_DB_PATH = Path("data/qi_catalog_match.db")

KNOWN_FOLDER_LINKS = {
    ("2024", "AT"): "https://drive.google.com/drive/folders/1_letSarz1x9wSAvnaFwJs1FNzHAYqW1V",
    ("2024", "TFLE"): "https://drive.google.com/drive/folders/1gEN8U86lRBgtYRgzxZ82zUrZ_iBe3mpe",
    ("2024", "WTS"): "https://drive.google.com/drive/folders/1ut8vXx5Zl22w6XEREQxGprmQaCQIi2lV",
    ("2023", "COMP"): "https://drive.google.com/drive/folders/1aEbZAcsmfChBy8-lv_FG-Pro2UFuBya4",
    ("2023", "COTT"): "https://drive.google.com/drive/folders/1xERJVEAHwliUUtOVVykv5jWdec8JWEkI",
    ("2023", "EMAI"): "https://drive.google.com/drive/folders/1hRK8PQpr2LH0I8ec796AHNJdiTT547Iw",
    ("2023", "EPHE"): "https://drive.google.com/drive/folders/1y1nz9Bj0gYS3XmLUHjOF6ZSMeDhrr3u_",
    ("2023", "HFTG"): "https://drive.google.com/drive/folders/1oFCdkIa5ayTXEqsDQZZ8uievfU6xaoTx",
    ("2023", "HTFG"): "https://drive.google.com/drive/folders/1oFCdkIa5ayTXEqsDQZZ8uievfU6xaoTx",
    ("2023", "PBC"): "https://drive.google.com/drive/folders/1Y4dHA_Kdxd6J-HW-z18S3kEgKL36BQ7J",
    ("2023", "RS"): "https://drive.google.com/drive/folders/17X8IxOzhQt_ir6pPozxUT1AIra6c_kcg",
    ("2023", "SRH"): "https://drive.google.com/drive/folders/1mWgz9dikv7hmSFotti1XWvDU8vvI0xZX",
    ("2023", "TG"): "https://drive.google.com/drive/folders/15Nv2-57edeBTemoLdSqjD82SvXxA2uTE",
    ("2022", "APP"): "https://drive.google.com/drive/folders/1aLv0X8oJMgf7sWXT2EgfPPZLac6uRJI2",
    ("2022", "BLOO"): "https://drive.google.com/drive/folders/1GuVp3Nfxz43eaZxv_rOGui95wdAPxsNT",
    ("2022", "HGH"): "https://drive.google.com/drive/folders/1gYSElSJFtIJ8f1h4bPz2Z2ocHqx0tU9X",
    ("2022", "HTME"): "https://drive.google.com/drive/folders/1_-uDYTbJelbrMK2dXZ5oeAy-x_s_SlrG",
    ("2022", "NALO"): "https://drive.google.com/drive/folders/1WpiPNyoxluPQVOndz0vzbr7URlCSpYFJ",
    ("2022", "NCM"): "https://drive.google.com/drive/folders/1pqwnVA_ati2DthX1kvDXJh3wkrG8TsOv",
    ("2022", "SS"): "https://drive.google.com/drive/folders/1kY0IYvtJW2ysOW9CY8WJ1xVNxNnQ83MV",
    ("2022", "SYW"): "https://drive.google.com/drive/folders/1ZxrTpbHBc1DhL5gfkrxqcZ4l9tcMNBPD",
    ("2022", "URBA"): "https://drive.google.com/drive/folders/1cRN4Yw69UKHDSkSZNG23xR6SXln-Aj3j",
    ("2021", "AFTE"): "https://drive.google.com/drive/folders/1ddR3LZOeVL0mK_P6xLQRlzdRIJ2dhI3w",
    ("2021", "ASAB"): "https://drive.google.com/drive/folders/1l2mqBGA4z9JtBZec0HpSTuDCbsn_9EaQ",
    ("2021", "BIAS"): "https://drive.google.com/drive/folders/1BPsBpCIDy54ygn3vlwEL24thkU_-VLJO",
    ("2021", "BUTC"): "https://drive.google.com/drive/folders/1OzOxyTQ3xEW1FHqXOXoxW_BhivBrENwP",
    ("2021", "DAIH"): "https://drive.google.com/drive/folders/14w8ZMGpQHP7y9EPJPPmUW0196XVF7qIo",
    ("2021", "DDJ"): "https://drive.google.com/drive/folders/1TC0904jJp0EMq4crYpXtNGpiwLZTnn5W",
    ("2021", "PCOB"): "https://drive.google.com/drive/folders/1bax0EJQ4c6ZITPuXxP5amkvuFmzkDzjb",
    ("2021", "RB"): "https://drive.google.com/drive/folders/14PT4mwkwh6ro2u3-HCNeBQC00ESMJolp",
    ("2021", "WILF"): "https://drive.google.com/drive/folders/1SphPBq89E6n56uLJ8zjyQKgHXMh9tiPr",
    ("2021", "YBBL"): "https://drive.google.com/drive/folders/1hjy5d2bLZgw-bnWjgb2ZSHCERqpQVfiL",
    ("2020", "ANNB"): "https://drive.google.com/drive/folders/1hf4EuJV9D0M2duf7BfEvdpEDdee6id0N",
    ("2020", "BIRT"): "https://drive.google.com/drive/folders/1hHxEokt6h8OjCvngoOW76San1pKpijkU",
    ("2020", "CN"): "https://drive.google.com/drive/folders/1J-EEZ46Vl21jT4MdKiFNiYxKJDhXobc1",
    ("2020", "FGFS"): "https://drive.google.com/drive/folders/13bmX8D-x9hGM_51pWRvnm25typGVGF45",
    ("2020", "FORT"): "https://drive.google.com/drive/folders/1P5dlTo3Bg2BWZQIvOXSKUFs9aidrDX2T",
    ("2020", "IFA"): "https://drive.google.com/drive/folders/1TM2oRpqXbGfOiHgShMeZNjQKylzaLj-q",
    ("2020", "KTTY"): "https://drive.google.com/drive/folders/1Diy4VN9-U95118S_nL2CRld6zI-fOOHx",
    ("2020", "TW"): "https://drive.google.com/drive/folders/1U6h9Khe7R5biEJWibsiZaTpatEv9YxO-",
    ("2020", "WWAS"): "https://drive.google.com/drive/folders/1oIxhf63K6ZfcyoXv1B1bYcIQ0zh3rbRJ",
    ("2019", "ETSA"): "https://drive.google.com/drive/folders/1PkFdsjR8j16VpPdoMBh6SxMlRMMR7UVn",
    ("2019", "FYTT"): "https://drive.google.com/drive/folders/1wpNNvC1oYdq3gOwzf8yYz49d9Xdv5eMf",
    ("2019", "IMBC"): "https://drive.google.com/drive/folders/1S4eP_r3xZNfOV73Td0nJIavXHzXpvIE1",
    ("2019", "ISST"): "https://drive.google.com/drive/folders/1vSbBnNOK0lyK1w8QDs0iadpEjmPH-MMW",
    ("2019", "PATR"): "https://drive.google.com/drive/folders/1qI0sbKKNiL_RYEKS0WqGdTa9NZs2E3Nt",
    ("2019", "SINK"): "https://drive.google.com/drive/folders/1uFQrx2BGw4KKfutHCettSOtfQTJkapQh",
    ("2019", "SWAL"): "https://drive.google.com/drive/folders/14zbpGa5UTGA8PNJnaHYSz_rLq_gPnRSj",
    ("2018", "ALSA"): "https://drive.google.com/drive/folders/1ir2OX9NaAC_tCNAtBa7roqtF8AaA0s0m",
    ("2018", "BOUN"): "https://drive.google.com/drive/folders/17mVcZ20_GqeRlhZl7NnRBRZRz-xq0FDN",
    ("2018", "DT"): "https://drive.google.com/drive/folders/1yqRXI7xTrkngDPt982_JuA-2HOr75dzz",
    ("2018", "NIO"): "https://drive.google.com/drive/folders/1v2q2oo9FFbFRDzR4YOUiI54tpSj1M8b7",
    ("2018", "SW"): "https://drive.google.com/drive/folders/1v6P5YDLOMwtE1DRaRn3aPLD5lRPnhcoF",
    ("2018", "TF"): "https://drive.google.com/drive/folders/1e5jgGxv1o1iHZ83GV0Skmiar7D6thHYQ",
    ("2017", "AUTO"): "https://drive.google.com/drive/folders/1gjxmb1OpWD4eK1GK8uVHiwO--65g4DpQ",
    ("2017", "DOMT"): "https://drive.google.com/drive/folders/1qRpNvEW2g89bTazC7F6WYoYLLdjMZhCn",
    ("2017", "HELI"): "https://drive.google.com/drive/folders/1m6hhfl4mrVjGDZI5mT-neR1RSQG9o0pE",
    ("2017", "NABF"): "https://drive.google.com/drive/folders/1_Plbb5XcFmDQX4-f2diGR83grYxCIDV-",
    ("2017", "PELU"): "https://drive.google.com/drive/folders/1PgFSDvI9O0lbvG1XL_msQtMrQNV4acVO",
    ("2017", "SCDM"): "https://drive.google.com/drive/folders/16q6qdE073gAOsD2asZzSr0TMUcUn9WSF",
    ("2016", "TCAW"): "https://drive.google.com/drive/folders/17siBw0KjdUOBjgWtctI7gI1ubsugmGZe",
    ("2015", "OND"): "https://drive.google.com/drive/folders/1L4da0XyoUpTdQh3F9y6nwAvDte258sQU",
}


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Backfill folder_link from known release-folder mappings."
    )
    parser.add_argument(
        "--db-path",
        type=Path,
        default=DEFAULT_DB_PATH,
        help=f"SQLite DB to update (default: {DEFAULT_DB_PATH})",
    )
    parser.add_argument(
        "--dry-run",
        action="store_true",
        help="Show what would be updated without writing to the DB.",
    )
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    connection = sqlite3.connect(args.db_path)
    try:
        total_updated = 0
        per_key: list[dict[str, object]] = []
        for (release_year, shortener), folder_link in KNOWN_FOLDER_LINKS.items():
            count = connection.execute(
                """
                SELECT COUNT(*)
                FROM image_assets
                WHERE COALESCE(NULLIF(TRIM(resolved_release_year), ''), NULLIF(TRIM(release_year), '')) = ?
                  AND COALESCE(NULLIF(TRIM(resolved_book_shortener), ''), NULLIF(TRIM(book_shortener), '')) = ?
                  AND (folder_link IS NULL OR TRIM(folder_link) = '')
                """,
                (release_year, shortener),
            ).fetchone()[0]
            if count == 0:
                continue
            per_key.append(
                {
                    "release_year": release_year,
                    "shortener": shortener,
                    "folder_link": folder_link,
                    "rows": count,
                }
            )
            total_updated += count
            if not args.dry_run:
                connection.execute(
                    """
                    UPDATE image_assets
                    SET folder_link = ?
                    WHERE COALESCE(NULLIF(TRIM(resolved_release_year), ''), NULLIF(TRIM(release_year), '')) = ?
                      AND COALESCE(NULLIF(TRIM(resolved_book_shortener), ''), NULLIF(TRIM(book_shortener), '')) = ?
                      AND (folder_link IS NULL OR TRIM(folder_link) = '')
                    """,
                    (folder_link, release_year, shortener),
                )
        if not args.dry_run:
            connection.commit()
        print(
            json.dumps(
                {
                    "dry_run": args.dry_run,
                    "updated_rows": total_updated,
                    "matched_groups": per_key,
                },
                ensure_ascii=True,
                indent=2,
            )
        )
        return 0
    finally:
        connection.close()


if __name__ == "__main__":
    raise SystemExit(main())

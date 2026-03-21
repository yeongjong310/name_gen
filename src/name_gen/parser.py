"""HTML 파싱 및 수리 길흉 조회 모듈.

nameMaker/tools/excel_converter.py에서 필요한 로직을 가져옴.
"""

from __future__ import annotations

import re
import urllib.request
from dataclasses import dataclass, field
from importlib import resources


# ---------------------------------------------------------------------------
# 데이터 클래스
# ---------------------------------------------------------------------------

@dataclass
class NameChar:
    """한 글자의 정보."""
    element: str  # 오행 (金, 木 등)
    hanja: str    # 한자 (이미지인 경우 "?")
    reading: str  # 훈음 (예: "김", "강철강")


@dataclass
class NameEntry:
    """이름 한 세트 (성 + 이름 두 글자)."""
    surname: NameChar
    first_char: NameChar
    second_char: NameChar


@dataclass
class SajuData:
    """사주 기본 정보 (모든 페이지 공통)."""
    gender: str = ""
    solar_year: str = ""
    solar_month: str = ""
    solar_day: str = ""
    lunar_year: str = ""
    lunar_month: str = ""
    lunar_day: str = ""
    si_char: str = ""
    saju_si: str = ""
    saju_il: str = ""
    saju_wol: str = ""
    saju_nyeon: str = ""
    oheng_top: list[str] = field(default_factory=list)
    oheng_bot: list[str] = field(default_factory=list)
    yuksin_top: list[str] = field(default_factory=list)
    yuksin_bot: list[str] = field(default_factory=list)
    daeun_nums: list[str] = field(default_factory=list)
    daeun_ganji: list[str] = field(default_factory=list)


@dataclass
class PageData:
    """한 페이지의 이름 데이터."""
    page_number: int
    name_count: int
    names: list[NameEntry] = field(default_factory=list)
    suri_numbers_list: list[list[int]] = field(default_factory=list)


# ---------------------------------------------------------------------------
# HTML 가져오기
# ---------------------------------------------------------------------------

def fetch_html(url: str) -> str:
    """URL에서 HTML을 가져와 디코딩."""
    req = urllib.request.Request(
        url,
        headers={
            "User-Agent": (
                "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) "
                "AppleWebKit/537.36 (KHTML, like Gecko) "
                "Chrome/120.0.0.0 Safari/537.36"
            ),
            "Accept": "text/html,application/xhtml+xml,*/*",
            "Accept-Language": "ko-KR,ko;q=0.9",
            "Accept-Encoding": "identity",
            "Connection": "keep-alive",
        },
    )
    with urllib.request.urlopen(req, timeout=15) as resp:
        raw = resp.read()

    for enc in ("euc-kr", "cp949", "utf-8"):
        try:
            return raw.decode(enc)
        except UnicodeDecodeError:
            continue
    return raw.decode("utf-8", errors="replace")


# ---------------------------------------------------------------------------
# HTML 파싱
# ---------------------------------------------------------------------------

def _strip_tags(html_str: str) -> str:
    return re.sub(r"<[^>]+>", "", html_str).strip()


def _extract_element(td_html: str) -> str:
    m = re.search(r"\(([^)]+)\)", td_html)
    return m.group(1) if m else ""


def _extract_hanja(td_html: str) -> str:
    if "<img" in td_html:
        return "?"
    text = _strip_tags(td_html)
    return text if text and text != "&nbsp;" else "?"


def _extract_reading(td_html: str) -> str:
    text = re.sub(r"<[^>]+>", "", td_html)
    # 연속 공백을 하나로 줄이고 앞뒤만 trim (훈과 음 사이 공백 유지)
    return re.sub(r"\s+", " ", text).strip()


def _parse_saju_table(saju_html: str) -> SajuData:
    saju = SajuData()
    rows = re.findall(r"<tr[^>]*>(.*?)</tr>", saju_html, re.DOTALL)

    if len(rows) > 0:
        tds = re.findall(r"<td[^>]*>(.*?)</td>", rows[0], re.DOTALL)
        tds_text = [_strip_tags(td) for td in tds]
        for td_text in tds_text:
            m = re.search(r"(\d+)년\s*(\d+)월\s*(\d+)일", td_text)
            if m:
                saju.solar_year = m.group(1)
                saju.solar_month = m.group(2)
                saju.solar_day = m.group(3)
                break
        for td_text in tds_text:
            m = re.search(r"([\u4e00-\u9fff])時", td_text)
            if m:
                saju.si_char = m.group(1)
                break
        for td_text in tds_text:
            if "乾命" in td_text:
                saju.gender = "乾"
                break
            elif "坤命" in td_text:
                saju.gender = "坤"
                break

    if len(rows) > 1:
        tds_text = [_strip_tags(td) for td in re.findall(r"<td[^>]*>(.*?)</td>", rows[1], re.DOTALL)]
        for td_text in tds_text:
            m = re.search(r"(\d+)년\s*(\d+)월\s*(\d+)일", td_text)
            if m:
                saju.lunar_year = m.group(1)
                saju.lunar_month = m.group(2)
                saju.lunar_day = m.group(3)
                break

    if len(rows) > 3:
        tds = re.findall(r"<td[^>]*>(.*?)</td>", rows[3], re.DOTALL)
        ganji_slots = ["", "", "", ""]
        for idx, td in enumerate(tds[:4]):
            chars = re.findall(r"[\u4e00-\u9fff]", td)
            if len(chars) == 2:
                ganji_slots[idx] = f"{chars[0]}\n{chars[1]}"
        saju.saju_si = ganji_slots[0]
        saju.saju_il = ganji_slots[1]
        saju.saju_wol = ganji_slots[2]
        saju.saju_nyeon = ganji_slots[3]

    if len(rows) > 4:
        tds_text = [_strip_tags(td) for td in re.findall(r"<td[^>]*>(.*?)</td>", rows[4], re.DOTALL)]
        saju.oheng_top = tds_text[:4]

    if len(rows) > 5:
        tds_text = [_strip_tags(td) for td in re.findall(r"<td[^>]*>(.*?)</td>", rows[5], re.DOTALL)]
        saju.oheng_bot = tds_text[:4]

    if len(rows) > 6:
        tds_text = [_strip_tags(td) for td in re.findall(r"<td[^>]*>(.*?)</td>", rows[6], re.DOTALL)]
        saju.yuksin_top = tds_text[:4]

    if len(rows) > 7:
        tds_text = [_strip_tags(td) for td in re.findall(r"<td[^>]*>(.*?)</td>", rows[7], re.DOTALL)]
        saju.yuksin_bot = tds_text[:4]

    if len(rows) > 8:
        tds_text = [_strip_tags(td) for td in re.findall(r"<td[^>]*>(.*?)</td>", rows[8], re.DOTALL)]
        saju.daeun_nums = [t for t in tds_text[:8] if t]

    if len(rows) > 9:
        tds_text = [_strip_tags(td) for td in re.findall(r"<td[^>]*>(.*?)</td>", rows[9], re.DOTALL)]
        saju.daeun_ganji = [t for t in tds_text[:8] if t]

    return saju


def _extract_suri_numbers(rows: list[str], name_count: int) -> list[list[int]]:
    def _clean(s: str) -> str:
        return re.sub(r"<[^>]+>", "", s).strip().replace("&nbsp;", "")

    def _get_num(row_idx: int, td_idx: int) -> int:
        if row_idx >= len(rows):
            return 0
        tds = re.findall(r"<td[^>]*>(.*?)</td>", rows[row_idx], re.DOTALL)
        if td_idx >= len(tds):
            return 0
        text = _clean(tds[td_idx])
        return int(text) if text.isdigit() else 0

    result: list[list[int]] = []
    for i in range(name_count):
        offset = i * 4
        ingyeok = _get_num(2, 2 + offset)
        oegyeok = _get_num(3, 0 + offset)
        jigyeok = _get_num(4, 2 + offset)
        chonggyeok = _get_num(6, 1 + offset)
        result.append([jigyeok, ingyeok, oegyeok, chonggyeok])
    return result


def _parse_name_rows(page_html: str) -> tuple[list[NameEntry], list[int]]:
    inner_end = page_html.find("</table>")
    if inner_end < 0:
        return [], []
    after_saju = page_html[inner_end + len("</table>"):]

    rows = re.findall(r"<tr>\s*(.*?)\s*</tr>", after_saju, re.DOTALL)

    surnames: list[NameChar] = []
    first_chars: list[NameChar] = []
    second_chars: list[NameChar] = []

    char_rows = []
    for row in rows:
        tds = re.findall(r"<td[^>]*>(.*?)</td>", row, re.DOTALL)
        has_element = bool(re.search(r'font size="3"', row))
        has_hanja = bool(re.search(r"font size=7", row))
        if has_element and has_hanja:
            char_rows.append(tds)

    if len(char_rows) < 2:
        return [], []

    surname_row = char_rows[0]
    name_char_rows = char_rows[1:]

    def _extract_chars_from_row(tds: list[str]) -> list[NameChar]:
        chars: list[NameChar] = []
        i = 0
        while i < len(tds):
            td = tds[i]
            if re.search(r'font size="3"', td) and re.search(r"\([^)]+\)", td):
                elem = _extract_element(td)
                hanja = ""
                reading = ""
                if i + 1 < len(tds) and re.search(r"font size=7", tds[i + 1]):
                    hanja = _extract_hanja(tds[i + 1])
                if i + 2 < len(tds) and re.search(r'size="4"', tds[i + 2]):
                    reading = _extract_reading(tds[i + 2])
                if elem:
                    chars.append(NameChar(element=elem, hanja=hanja, reading=reading))
                i += 3
            else:
                i += 1
        return chars

    surnames = _extract_chars_from_row(surname_row)

    for row_tds in name_char_rows:
        chars = _extract_chars_from_row(row_tds)
        if not first_chars or len(first_chars) == len(surnames):
            if len(first_chars) < len(surnames):
                first_chars.extend(chars)
            else:
                second_chars.extend(chars)
        else:
            first_chars.extend(chars)

    names: list[NameEntry] = []
    for idx in range(min(len(surnames), len(first_chars), len(second_chars))):
        names.append(NameEntry(
            surname=surnames[idx],
            first_char=first_chars[idx],
            second_char=second_chars[idx],
        ))

    suri_numbers_list = _extract_suri_numbers(rows, len(names))
    return names, suri_numbers_list


def _extract_applicant_name(html: str) -> str:
    """HTML에서 신청자 이름 추출. '신청자 : 이름' 패턴."""
    m = re.search(r"신청자\s*:\s*([^<]+)", html)
    if m:
        return m.group(1).strip()
    return ""


def parse_html(html: str) -> tuple[SajuData, list[PageData], str]:
    """HTML 전체를 파싱하여 사주 데이터, 페이지별 이름 데이터, 신청자 이름 반환."""
    applicant = _extract_applicant_name(html)
    parts = re.split(r"재?작명\s*\d*\s*[\r\n]*\s*<BR>", html, flags=re.IGNORECASE)

    saju = SajuData()
    pages: list[PageData] = []

    page_num = 0
    for part in parts:
        saju_match = re.search(
            r"<table[^>]*bgcolor=999999[^>]*>(.*?)</table>", part, re.DOTALL
        )
        if saju_match and not saju.gender:
            saju = _parse_saju_table(saju_match.group(1))

        if saju_match:
            names, suri_numbers_list = _parse_name_rows(part)
            if names:
                page_num += 1
                pages.append(PageData(
                    page_number=page_num,
                    name_count=len(names),
                    names=names,
                    suri_numbers_list=suri_numbers_list,
                ))

    return saju, pages, applicant


# ---------------------------------------------------------------------------
# 수리 길흉 조회
# ---------------------------------------------------------------------------

def lookup_happy_numbers(suri_numbers: list[int]) -> dict[str, str]:
    """수리 숫자로 happy_numbers.xlsx에서 길흉 텍스트 조회.

    suri_numbers: [지격, 인격, 외격, 총격]
    Returns: {"jigyeok": text, "ingyeok": text, "oegyeok": text, "chonggyeok": text}
    """
    import openpyxl

    happy_path = resources.files("name_gen").joinpath("assets/happy_numbers.xlsx")
    with resources.as_file(happy_path) as p:
        wb = openpyxl.load_workbook(p, data_only=True)
    ws = wb.active

    jigyeok, ingyeok, oegyeok, chonggyeok = suri_numbers

    result = {}
    cell_map = [
        ("jigyeok", jigyeok, "C"),
        ("ingyeok", ingyeok, "D"),
        ("oegyeok", oegyeok, "E"),
        ("chonggyeok", chonggyeok, "F"),
    ]
    for key, num, col in cell_map:
        if 1 <= num <= 58:
            val = ws[f"{col}{num}"].value
            result[key] = str(val).replace(" ", "\n\n") if val else ""
        else:
            result[key] = ""

    return result

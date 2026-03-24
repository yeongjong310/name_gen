"""메모장용 텍스트 파일 생성 모듈."""

from __future__ import annotations

from name_gen.parser import NameEntry, PageData, lookup_happy_numbers


def _format_name_text(entry: NameEntry) -> str:
    """이름 표시 텍스트 생성. 예: 김강민(金鋼? - 강철 강, 옥돌 민)"""
    surname = entry.surname
    first = entry.first_char
    second = entry.second_char

    hangul = surname.reading[-1] if surname.reading else "?"
    hangul += first.reading[-1] if first.reading else "?"

    def _space_reading(r: str) -> str:
        if len(r) >= 2 and " " not in r:
            return r[:-1] + " " + r[-1]
        return r

    first_desc = _space_reading(first.reading)

    if second.reading:
        hangul += second.reading[-1]
        second_desc = _space_reading(second.reading)
        hanja = f"{surname.hanja}{first.hanja}{second.hanja}"
        return f"{hangul}({hanja} - {first_desc}, {second_desc})"
    else:
        hanja = f"{surname.hanja}{first.hanja}"
        return f"{hangul}({hanja} - {first_desc})"


def generate_txt(pages: list[PageData], output_path: str) -> None:
    """모든 PageData를 하나의 메모장 텍스트 파일로 생성."""
    lines: list[str] = []

    for page in pages:
        # 이름별 길흉 텍스트 조회
        happy_texts_list: list[dict[str, str]] = []
        for suri_numbers in page.suri_numbers_list:
            if suri_numbers and any(suri_numbers):
                happy_texts_list.append(lookup_happy_numbers(suri_numbers))
            else:
                happy_texts_list.append({})

        for idx, entry in enumerate(page.names[: page.name_count]):
            # 각 이름 앞에 빈 줄 2개
            lines.append("")
            lines.append("")

            # 이름풀이
            lines.append(_format_name_text(entry))

            # 수리오행 (빈 줄 없이 바로)
            suri = (
                page.suri_numbers_list[idx]
                if idx < len(page.suri_numbers_list)
                else [0, 0, 0, 0]
            )
            jigyeok, ingyeok, oegyeok, chonggyeok = suri
            lines.append("•수리오행 (원격, 형격, 이격, 정격)")

            # 길흉
            happy = happy_texts_list[idx] if idx < len(happy_texts_list) else {}
            fortune_lines = [
                ("초년운", jigyeok, "jigyeok"),
                ("장년운", ingyeok, "ingyeok"),
                ("중년운", oegyeok, "oegyeok"),
                ("말년운", chonggyeok, "chonggyeok"),
            ]
            for label, num, key in fortune_lines:
                happy_text = happy.get(key, "")
                if happy_text:
                    text_oneline = happy_text.replace("\n\n", " ")
                    lines.append(f"{label} : {text_oneline}")
                else:
                    lines.append(f"{label} : {num}")

    with open(output_path, "w", encoding="utf-8") as f:
        f.write("\n".join(lines) + "\n")

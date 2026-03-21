"""작명 데이터 → Word 문서 변환 GUI."""

from __future__ import annotations

import json
import os
import platform
import subprocess
import threading
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, scrolledtext

from name_gen.parser import fetch_html, parse_html
from name_gen.word_writer import generate_docx

CONFIG_FILE = Path.home() / ".namegen_config.json"


class NameGenApp:
    """작명 Word 변환 GUI 앱."""

    @staticmethod
    def _load_config() -> str:
        if CONFIG_FILE.exists():
            try:
                with open(CONFIG_FILE) as f:
                    return json.load(f).get("output_dir", "")
            except (json.JSONDecodeError, IOError):
                return ""
        return ""

    @staticmethod
    def _save_config(output_dir: str):
        with open(CONFIG_FILE, "w") as f:
            json.dump({"output_dir": output_dir}, f)

    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title("작명 Word 변환")
        self.root.geometry("700x550")
        self.root.resizable(True, True)

        self.output_dir = self._load_config()
        self._build_ui()

    def _build_ui(self):
        # 저장 폴더
        frame_output = tk.Frame(self.root)
        frame_output.pack(fill=tk.X, padx=10, pady=(10, 5))
        tk.Label(frame_output, text="저장 폴더:").pack(side=tk.LEFT)
        self.label_output_dir = tk.Label(
            frame_output,
            text=self.output_dir or "(미설정)",
            wraplength=400,
            justify=tk.LEFT,
        )
        self.label_output_dir.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(5, 5))
        tk.Button(
            frame_output, text="폴더 열기", command=self._open_output_dir, width=10
        ).pack(side=tk.RIGHT)
        tk.Button(
            frame_output, text="변경", command=self._select_output_dir, width=10
        ).pack(side=tk.RIGHT, padx=(0, 5))

        # URL 입력
        frame_url_label = tk.Frame(self.root)
        frame_url_label.pack(fill=tk.X, padx=10, pady=(10, 2))
        tk.Label(frame_url_label, text="URL 입력:").pack(side=tk.LEFT)

        self.text_url = scrolledtext.ScrolledText(self.root, height=4)
        self.text_url.pack(fill=tk.BOTH, padx=10, pady=(0, 5), expand=True)

        # 실행 버튼
        frame_btn = tk.Frame(self.root)
        frame_btn.pack(fill=tk.X, padx=10, pady=5)
        self.btn_run = tk.Button(
            frame_btn,
            text="실행",
            command=self._on_run,
            width=20,
            height=2,
            bg="#4CAF50",
            fg="black",
            font=("Arial", 12, "bold"),
        )
        self.btn_run.pack()

        # 로그
        tk.Label(self.root, text="처리 로그:").pack(anchor=tk.W, padx=10, pady=(5, 2))
        self.text_log = scrolledtext.ScrolledText(
            self.root, height=12, state=tk.DISABLED
        )
        self.text_log.pack(fill=tk.BOTH, padx=10, pady=(0, 10), expand=True)

    def _open_output_dir(self):
        if not self.output_dir or not os.path.isdir(self.output_dir):
            messagebox.showwarning("경고", "저장 폴더가 설정되지 않았거나 존재하지 않습니다.")
            return
        system = platform.system()
        if system == "Darwin":
            subprocess.Popen(["open", self.output_dir])
        elif system == "Windows":
            os.startfile(self.output_dir)
        else:
            subprocess.Popen(["xdg-open", self.output_dir])

    def _select_output_dir(self):
        path = filedialog.askdirectory(title="저장할 폴더 선택")
        if path:
            self.output_dir = path
            self._save_config(path)
            self.label_output_dir.config(text=path)
            messagebox.showinfo("성공", "저장 폴더가 설정되었습니다.")

    def _log(self, msg: str):
        self.text_log.config(state=tk.NORMAL)
        self.text_log.insert(tk.END, msg + "\n")
        self.text_log.see(tk.END)
        self.text_log.config(state=tk.DISABLED)

    def _on_run(self):
        if not self.output_dir:
            messagebox.showwarning("경고", "저장 폴더를 먼저 선택해주세요.")
            return

        url = self.text_url.get("1.0", tk.END).strip()
        if not url:
            messagebox.showwarning("경고", "URL을 입력해주세요.")
            return

        self.text_log.config(state=tk.NORMAL)
        self.text_log.delete("1.0", tk.END)
        self.text_log.config(state=tk.DISABLED)

        self.btn_run.config(state=tk.DISABLED, text="처리 중...")
        thread = threading.Thread(
            target=self._process_url, args=(url,), daemon=True
        )
        thread.start()

    def _process_url(self, url: str):
        try:
            self.root.after(0, self._log, f"페이지 가져오는 중: {url}")
            html = fetch_html(url)
            self.root.after(0, self._log, "HTML 가져오기 완료")

            saju, pages, applicant = parse_html(html)
            self.root.after(0, self._log, f"총 {len(pages)}개 페이지 발견")

            if not pages:
                self.root.after(0, self._log, "[오류] 이름 데이터를 찾을 수 없습니다.")
                self.root.after(0, self._finish)
                return

            prefix = applicant if applicant else "result"
            success = 0
            for page in pages:
                output_name = f"{prefix}_{page.page_number}.docx"
                output_path = os.path.join(self.output_dir, output_name)

                generate_docx(page, output_path)
                self.root.after(
                    0, self._log,
                    f"  [{page.page_number}/{len(pages)}] "
                    f"이름 {page.name_count}개 → {output_name}",
                )
                success += 1

            summary = f"\n{'=' * 40}\n완료! 성공: {success}건"
            self.root.after(0, self._log, summary)

        except Exception as e:
            self.root.after(0, self._log, f"[오류] {e}")

        self.root.after(0, self._finish)

    def _finish(self):
        self.btn_run.config(state=tk.NORMAL, text="실행")


def main():
    root = tk.Tk()
    NameGenApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()

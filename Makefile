SHELL := /usr/bin/bash
UV := uv.exe

.PHONY: setup sync build clean run test

# uv로 가상환경 생성 및 의존성 설치 (build 포함)
setup:
	$(UV) sync --extra build --extra dev

# 의존성만 동기화
sync:
	$(UV) sync

# PyInstaller로 Windows exe 빌드
build: setup
	$(UV) run python -m PyInstaller 작명Word변환.spec --noconfirm

# 개발 모드로 앱 실행
run:
	$(UV) run python -m name_gen.main

# 테스트 실행
test:
	$(UV) run pytest

# 빌드 산출물 정리
clean:
	rm -rf build dist

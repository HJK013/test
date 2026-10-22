# 올리브영 메이크업 옵션 색상 수집 PoC

립·베이스·아이 메이크업 상품의 옵션별 대표 색상(HEX, Lab)과 제형을 수집하는 Python 실험 코드입니다. 현재 주 실행 파일은 `oliveyoung_color_poc13.py`입니다. v13은 모바일 목록 POST API와 모바일 상세·리뷰 경로를 사용하고, 상품 20개를 동시에 처리합니다. 이미지 다운로드도 전체 20개까지 병렬 처리하며 작업자별 요청 간격은 최소 0.1초입니다. 색 판정 규칙과 19열 엑셀 출력은 유지합니다. 이전 `oliveyoung_color_poc*.py` 파일은 실험 버전이며, `compare_aside_vs_script.py`는 수집 결과와 별도 CSV를 비교합니다.

## 실행

```powershell
python -m pip install requests pillow openpyxl
python oliveyoung_color_poc13.py
```

v13은 쿠키 파일 없이 실행합니다. 스크립트는 카테고리당 상품 100개, 총 300개를 요청하며, 생성한 엑셀 파일은 실행 파일과 같은 폴더에 저장합니다. `PER_CATEGORY`, `PRODUCT_WORKERS`, `IMAGE_WORKERS`, `THREAD_DELAY`로 수집량과 동시성을 조절할 수 있습니다. 목록 API는 요청당 최대 100개씩 가져옵니다. 429 응답 시 재시도는 유지합니다.

목록 정렬은 모바일 API의 `POPULAR`를 사용합니다. 사이트의 상품 목록이 갱신되면 앞의 100개 상품과 결과 행은 달라질 수 있습니다.

## 결과와 참고 자료

- 결과 엑셀은 상품·옵션 정보, HEX/Lab, 제형, 색 출처와 신뢰도 등 19열입니다.
- `python compare_aside_vs_script.py <수집결과.xlsx> <비교대상.csv>`로 옵션별 색 차이(ΔE)를 비교할 수 있습니다.
- `제형.md`는 제형 7종과 별도 AR 구현에 관한 참고 메모입니다.

사이트 구조와 리뷰 API 응답이 달라지면 수집 규칙을 다시 확인해야 합니다.

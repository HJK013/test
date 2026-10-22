# 올리브영 메이크업 옵션 색상 수집 PoC

립·베이스·아이 메이크업 상품의 옵션별 대표 색상(HEX, Lab)과 제형을 수집하는 Python 실험 코드입니다. 현재 주 실행 파일은 `oliveyoung_color_poc14.py`입니다. v14는 모바일 목록 POST API를 100개씩 끝까지 조회하고, 상품 40개와 이미지 80개를 병렬 처리합니다. 상품이 끝날 때마다 19열 CSV에 즉시 추가하며 완료 로그를 기록합니다. 403/429 응답이 나오면 새 요청을 멈추고 종료합니다. 이전 `oliveyoung_color_poc*.py` 파일은 실험 버전이며, `compare_aside_vs_script.py`는 수집 결과와 별도 CSV를 비교합니다.

## 실행

```powershell
python -m pip install requests pillow
python oliveyoung_color_poc14.py
```

v14는 쿠키 파일 없이 실행합니다. `PER_CATEGORY = None`이 기본값이며 카테고리의 전체 목록을 가져옵니다. `PRODUCT_WORKERS`, `IMAGE_WORKERS`, `THREAD_DELAY`로 동시성을 조절할 수 있습니다. 출력은 실행 파일과 같은 폴더의 `oliveyoung_makeup_colors_full.csv`입니다. 중단 후 같은 명령으로 다시 실행하면 `.done` 완료 로그를 기준으로 이어받습니다. 403/429가 발생하면 연결을 확인한 뒤 재실행하면 됩니다.

목록 정렬은 모바일 API의 `POPULAR`를 사용합니다. 사이트의 상품 목록이 갱신되면 앞의 100개 상품과 결과 행은 달라질 수 있습니다.

## 결과와 참고 자료

- 결과 CSV는 상품·옵션 정보, HEX/Lab, 제형, 색 출처와 신뢰도 등 19열입니다.
- `compare_aside_vs_script.py`는 이전 실험 결과 엑셀과 별도 CSV의 옵션별 색 차이(ΔE)를 비교합니다.
- `제형.md`는 제형 7종과 별도 AR 구현에 관한 참고 메모입니다.

사이트 구조와 리뷰 API 응답이 달라지면 수집 규칙을 다시 확인해야 합니다.

# 올리브영 메이크업 옵션 색상 수집 PoC

립·베이스·아이 메이크업 상품의 옵션별 대표 색상(HEX, Lab)과 제형을 수집하는 Python 실험 코드입니다. 현재 주 실행 파일은 `oliveyoung_color_poc10.py`입니다. 이전 `oliveyoung_color_poc*.py` 파일은 색 추출 규칙의 실험 버전이며, `compare_aside_vs_script.py`는 수집 결과와 별도 CSV를 비교합니다.

## 실행

```powershell
python -m pip install curl_cffi beautifulsoup4 pillow openpyxl
python oliveyoung_color_poc10.py
```

필요하면 같은 폴더에 `cookie.txt`를 두고 `cookie`와 `user-agent`를 각각 다음 줄에 적습니다. 이 파일에는 세션 정보가 들어가므로 Git에 올리지 않습니다. 스크립트는 카테고리당 최대 10개 상품을 요청하며, 생성한 엑셀 파일은 실행 파일과 같은 폴더에 저장합니다.

## 결과와 참고 자료

- 결과 엑셀은 상품·옵션 정보, HEX/Lab, 제형, 색 출처와 신뢰도 등 19열입니다.
- `python compare_aside_vs_script.py <수집결과.xlsx> <비교대상.csv>`로 옵션별 색 차이(ΔE)를 비교할 수 있습니다.
- `제형.md`는 제형 7종과 별도 AR 구현에 관한 참고 메모입니다.

사이트 구조와 리뷰 API 응답이 달라지면 수집 규칙을 다시 확인해야 합니다.

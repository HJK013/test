# 올리브영 메이크업 옵션 색상 수집 PoC

립·베이스·아이 메이크업 상품의 옵션별 대표 색상(HEX, Lab)과 제형을 수집하는 Python 실험 코드입니다. 현재 주 실행 파일은 `oliveyoung_color_poc11.py`입니다. v11은 v10의 색 판정 규칙을 유지하면서 상품별 이미지 다운로드를 6개씩 병렬로 처리하고 페이지 간 대기를 0.5초로 줄였습니다. 이전 `oliveyoung_color_poc*.py` 파일은 실험 버전이며, `compare_aside_vs_script.py`는 수집 결과와 별도 CSV를 비교합니다.

## 실행

```powershell
python -m pip install curl_cffi beautifulsoup4 pillow openpyxl
python oliveyoung_color_poc11.py
```

필요하면 같은 폴더에 `cookie.txt`를 두고 `cookie`와 `user-agent`를 각각 다음 줄에 적습니다. 이 파일에는 세션 정보가 들어가므로 Git에 올리지 않습니다. 스크립트는 카테고리당 최대 10개 상품을 요청하며, 생성한 엑셀 파일은 실행 파일과 같은 폴더에 저장합니다.

v11은 쿠키 파일의 Chrome User-Agent와 Windows/macOS 브라우저 힌트를 함께 보냅니다. `curl_cffi`에 같은 Chrome 버전의 TLS 프로필이 없으면 설치된 최신 프로필을 사용하고 버전 차이를 알립니다. 이 경우 헤더가 일치해도 실제 브라우저와 완전히 같은 요청은 아닙니다. 이미지 CDN 요청에는 페이지 쿠키를 보내지 않습니다.

## 결과와 참고 자료

- 결과 엑셀은 상품·옵션 정보, HEX/Lab, 제형, 색 출처와 신뢰도 등 19열입니다.
- `python compare_aside_vs_script.py <수집결과.xlsx> <비교대상.csv>`로 옵션별 색 차이(ΔE)를 비교할 수 있습니다.
- `제형.md`는 제형 7종과 별도 AR 구현에 관한 참고 메모입니다.

사이트 구조와 리뷰 API 응답이 달라지면 수집 규칙을 다시 확인해야 합니다.

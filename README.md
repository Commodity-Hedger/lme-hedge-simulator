# LME 헤지 시뮬레이터

비철금속(LME) 헤지 시뮬레이터입니다.

- `index.html` : 시뮬레이터 페이지
- `curve.json` : 스프레드 계산에 쓰는 LME 커브 (매일 갱신)
- `scripts/update_curve.py` : 커브 엑셀을 curve.json으로 바꾸는 스크립트

## 커브 수동 갱신

```
pip install openpyxl
python scripts/update_curve.py "<파일명 (YYYYMMDD)>.xlsx"
git add curve.json && git commit -m "curve: YYYY-MM-DD" && git push
```

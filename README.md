# 3학년 전환기 일정

3학년 15반 전환기(2026.11.23 ~ 2027.1.7) 하루 보기 · 달력 페이지.

## 올리는 법 (GitHub Pages)

1. GitHub에서 새 저장소 만들기 (예: `jeonhwangi`)
2. 이 폴더 안의 파일을 **폴더째 말고 내용물 그대로** 저장소 맨 위에 올리기
   ```
   index.html
   manifest.webmanifest
   sw.js
   favicon.svg
   icons/  (png 5개)
   ```
3. 저장소 **Settings → Pages → Branch: main / (root) → Save**
4. 1~2분 뒤 `https://<아이디>.github.io/jeonhwangi/` 로 접속

## 홈 화면에 추가

- 아이폰: Safari로 열기 → 공유 버튼 → **홈 화면에 추가**
- 안드로이드: Chrome으로 열기 → ⋮ → **홈 화면에 추가 / 앱 설치**

홈 화면에서 열면 주소창 없이 앱처럼 열리고, 인터넷이 없어도 마지막으로 본 내용이 떠.

## 일정 고치는 곳

`index.html` 안의 `<script>` 맨 위쪽

- `D` : 날짜별 전환기 일정
- `TT` : 요일별 원래 담당 선생님 (1~7교시)
- `SWAP` : 그 주에만 맞바꾼 시간 (컴시간 노란 칸)

고친 뒤 올리면 홈 화면 앱도 다음에 열 때 새 내용으로 바뀌어.

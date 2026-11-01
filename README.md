# 일일계획표

개인용 일일 업무 계획 관리 앱. 단일 HTML + Firebase Realtime Database + GitHub Pages.

배포 주소: https://kdsan1203-tech.github.io/daily_planner/

## 구성

| 파일 | 역할 |
|---|---|
| `index.html` | 앱 전체 (HTML/CSS/JS 단일 파일) |
| `manifest.webmanifest` | PWA 설치 정보 |
| `sw.js` | 서비스워커 (앱 셸 캐싱, 오프라인 구동) |
| `icon-*.png`, `maskable-*.png` | 홈화면 아이콘 |

## 화면

- **계획 시간표** — WEEKLY BIG 3, DAILY BIG, SUB THINGS, MEMO, BACKLOG, TIME RAIL(05:00~22:00)
- **주간일지** — 주간업무보고 생성 (복사 / txt / 인쇄)
- **시간통계** — 이번 주 · 최근 4주 시간 배분
- **리츄얼** — 시간대별(오전/업무/저녁) 체크리스트, 미완료 깜빡 알림

## 데이터

Firebase RTDB 경로 `planner_data/my_planner` 한 곳에 전체 상태를 보관한다.

```
{ days: { "YYYY-MM-DD": { big[3], subs[8], memo, slots, rit, u } },
  weeks: { "주키": { big[3], u } },
  backlog: [ { t, w, done } ],
  rituals: [ { id, name, rows } ],
  settings: { nudge } }
```

로컬(localStorage)와 원격을 병합할 때는 **내용이 있는 쪽을 우선**한다.
시각만 비교하면 빈 껍데기가 실제 기록을 덮어쓰는 사고가 나기 때문이다(v8.1에서 수정).

## 백업

- 하루 첫 사용 시 자동 백업
- 기록이 줄어드는 것이 감지되면 줄기 직전 상태를 스냅샷으로 보관
- 상단 메뉴의 백업 / 불러오기 / 백업복원으로 수동 관리

## 설치

안드로이드 크롬에서 주소를 열면 하단에 설치 배너가 뜨거나,
메뉴(⋮) → "앱 설치"로 홈 화면에 추가할 수 있다. 설치하면 주소창 없는
전체화면으로 뜨고, 인터넷이 없어도 열린다(기록은 기기에 저장 후 연결 시 동기화).

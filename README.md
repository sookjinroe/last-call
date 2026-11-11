# LAST CALL — 매장 칵테일 슬롯

카운터의 QR 하나를 찍으면 손님 폰에서 슬롯이 열려요.

- **트라이얼** — 누구나 바로. 코인 10개, 끝나면 계속 다시. 당첨 연출은 같고 경품은 없어요.
- **리얼** — 사장님이 손님 폰에 PIN을 넣어 줄 때마다 한 판(코인 10개).
  끝나면 날짜가 크게 찍힌 결과 화면이 남고, 당첨이 있으면 사장님 PIN으로만 닫혀요.
  손님이 그 화면을 보여 주면 사장님이 경품을 드려요.

서버나 구글 시트 없이 파일만으로 돌아가요.

| 파일 | 용도 |
|---|---|
| `index.html` | 손님 게임 |
| `qr.html` | 카운터용 입장 QR (메인 아이패드에 띄우거나 A4로 인쇄) |
| `config.js` | 리얼 모드 PIN(해시), 등수 이름 |

## 주소
- 게임: `https://sookjinroe.github.io/last-call/`
- 카운터 QR: `https://sookjinroe.github.io/last-call/qr.html`

## 라이선스
three.js (MIT), qrcode-generator by Kazuhiko Arase (MIT) — `vendor/`

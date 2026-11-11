// LAST CALL 설정
window.LASTCALL = {
  // 리얼 모드 사장님 PIN (4자리)의 해시값. 지금은 임시 PIN 0000.
  // 바꾸려면 새 PIN을 알려 주세요. (계산: SHA-256("lastcall:" + PIN) → base64)
  pinHash: 'GcHiBQ3+IeCjp3X7rA8oWIUhB4UnGmztOXvs7h9KTMc=',
  // 등수 이름 (경품이 정해지면 여기만 바꾸면 돼요)
  prizes: { seven: '1등', bar: '2등', bell: '3등', cilantro: '4등' },
};

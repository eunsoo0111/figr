# figr

SMIC 3주차 종목 **FIGR(Figure Technology Solutions)** 보고서, 특히 **투자포인트 1**을 쓰기 위한 저장소.
이 README는 프로젝트 **허브**예요. 브랜치·자료·진행 상황을 여기서 한눈에 보고, 작업할 때마다 아래 「업데이트 규칙」대로 고쳐 주세요.

- 마지막 갱신: 2026-11-17 (`claude/fervent-brown-x3pek4`)

---

## 1. 투자포인트 1: 지금의 논리

> **HELOC(주택 지분 접근) 시장의 구조적 성장 속에서, 파트너(특히 은행)가 늘고 매수자도 늘면서 Figure Connect라는 양면 장터가 커지고, 그 결과 FIGR이 잘 된다.**

```
[0] 구조적 성장 ─ 고금리 락인 → 지분 접근 수요가 HELOC으로 / Figure가 여는 그린필드 / 금리 하락 시 1순위 확장
      ▼
[1] 파트너(판매) 증가 ─ [1-B] 은행의 신규 진입 = 판매자이자 매수자
      ▼  왜?
[2] 해자 ─ LOS는 진입로, Connect 유동성(표준 크레딧 박스·매수자·AAA 이력)이 해자
      ▼
[3] 매수자(구매) 유입 ─ 보험사·운용사·크레딧 펀드 + 은행·IMB
      ▼
[4] Connect 양면 네트워크 → [1]로 순환
      ▼
[5] 실적 ─ 거래량↑ × 수취율↓ → 공헌이익·EBITDA 마진↑
```

고리별 근거·반증·상태는 → [`research/IP1_evidence_map.md`](research/IP1_evidence_map.md)

## 2. 진행 상황

| 고리 | 상태 | 한 줄 | 다음 할 일 |
|---|---|---|---|
| [0] 구조적 성장 | 🟡 | 질적 근거 풍부(용도 구성, 금리 비탄력, 상속 주택) | 분기 HELOC 실행액 시계열, 은행/비은행 분할 |
| [1] 파트너 증가 | 🟢/🟡 | 수는 확실(489→612/894), 질(멀티소싱·집중도)은 반증 있음 | 코호트별 물량, 경쟁사 파트너 확보 |
| [1-B] 은행 진입 | 🟡 | 경영진 발언(Flagstar, 은행이 Connect로 모기지를 모은다, CRA) 확보 | 은행 HELOC 철수·재진입 이력, 위험가중치·CRA, 은행 파트너 식별 |
| [2] 해자 | 🟡 | 인과를 "Connect 유동성 → LOS 채택"으로 재정의 | Lee Ward 소송·TILA, MAXEX·Pylon·Aven 규모 |
| [3] 매수자 유입 | 🔴 | **Kickstart(Figure 보장 매수)** 발견 → 제3자 매수 비중 미검증 | 10-Q 특수관계자·Fig SIX 매입 규모, 매수자 수 시계열 |
| [4] 양면 효과 | 🟡 | Connect 42%→72%, 그러나 Kickstart 포함 가능 | 4Q 콜 공헌이익 공개 |
| [5] 실적 연결 | 🟡 | 마진↑ 수취율↓, 은행은 수취율 더 낮음 | 공헌이익 모델링 |

상태: 🟢 근거 충분 · 🟡 근거 있으나 숫자 부족 · 🔴 미검증·반증 있음

## 3. 브랜치 지도

| 브랜치 | 목적 | 상태 | 주요 파일 |
|---|---|---|---|
| `main` | 기본. 정리된 결과만 병합 | 초기 커밋만 | `README.md` |
| `claude/fervent-brown-x3pek4` | 자료 5종(엑스퍼트콜 3·Greenwich 2) 정리, artifact 대비 비교, 투자포인트 1 근거 지도, 이 허브 | ✅ 2026-11-17 | `research/` 전체 |

**브랜치 규칙 제안**
- 새 리서치는 주제별 브랜치로: `research/<주제>` (예: `research/bank-entry`, `research/connect-buyers`, `research/heloc-market`)
- 브랜치를 만들면 **이 표에 한 줄 추가**, 끝나면 상태를 ✅로 바꾸고 결과 파일 경로를 적기
- 다른 브랜치 내용을 볼 때: `git fetch origin <브랜치>` 후 `git show origin/<브랜치>:<경로>` 또는 `git checkout <브랜치>`

## 4. 자료 목록

### artifact (claude.ai, 비공개 링크)

| 이름 | 날짜 | 내용 |
|---|---|---|
| [FIGR 비즈니스 모델 완전 정리](https://claude.ai/artifact/NRdcPtUrzbdTRbbpT3Tod6) | 2026-11-16 | 공시·컨콜 + Capital Blueprint + Ottavi 2편. 21장 |
| [FIGR 매수자 해부](https://claude.ai/artifact/YTtwTzJ51edGy5Q5hmoRPt) | 2026-11-16 | Connect 반대편(매수자) 12장 |
| [FIGR 새로 나온 이야기](https://claude.ai/artifact/GivXsyshwp3HUvQw2BHj2h) | 2026-11-17 | 자료 5종에서 앞선 두 문서에 없던 내용만 투자포인트 1 순서로 정리(Kickstart, 은행의 두 역할, 해자 재정의, 리스크) |

### 원자료 노트 (`research/sources/`)

원본 PDF는 유료 라이선스 자료이고 다운로더 워터마크가 있어서 **저장소에 올리지 않아요**. 노트는 내용을 한국어로 상세 정리한 것.

| 노트 | 원본 | 인터뷰·발행일 | 핵심 |
|---|---|---|---|
| [EC_v3_NAF-EVP_2026-09-09](research/sources/EC_v3_NAF-EVP_2026-09-09.md) | 엑스퍼트콜 `_v3` (22쪽) | 9/9 · 9/23 | 판매자 쪽. **Kickstart Connect(Figure 보장 매수 25~30bp)**, 수수료 구간, 경쟁사(Aven·Pylon·SoFi), 회계 의심, "Connect가 해자" |
| [EC_v2_NFM-SVP_2026-09-16](research/sources/EC_v2_NFM-SVP_2026-09-16.md) | 엑스퍼트콜 `_v2` (10쪽) | 9/16 · 9/30 | **유일한 매수자 쪽 증언**. 매수자 25~35곳, 봉인 입찰, 매수자 무료, 보증 없음, LOS→Connect 자동 연결, MAXEX |
| [EC_v1_NAF-EVP_2026-10-05](research/sources/EC_v1_NAF-EVP_2026-10-05.md) | 엑스퍼트콜 (무번호, 12쪽) | 10/5 · 10/19 | v3와 같은 사람. **HELOC 용도·금리 민감도·시차**, 1순위, 자체 서비싱, **전액 인출 규제 리스크·Lee Ward 소송** |
| [GW_2Q26_2026-08-18](research/sources/GW_2Q26_2026-08-18.md) | Greenwich 32쪽 | 8/18 | 2Q 10-Q·콜. **은행이 Connect로 모기지를 모은다(CRA)**, Flagstar 전환점, Figure factor 2.6배, AAA 255→135bp, EPD $2.9B |
| [GW_3Q26_2026-11-13](research/sources/GW_3Q26_2026-11-13.md) | Greenwich 8쪽 | 11/13 | 3Q 실적. 순 수취율 3.5%, **은행 유연성이 수취율에 반영**, 등대 실행사, 리스크 숫자 |

### 분석 문서 (`research/`)

| 문서 | 내용 |
|---|---|
| [gap_analysis_2026-11-17](research/gap_analysis_2026-11-17.md) | 자료 5종 vs artifact 2개: 새 정보 22개, 충돌·수정 9개, 투자포인트 1에서 놓친 것 5개, 자료별 신뢰도 |
| [IP1_evidence_map](research/IP1_evidence_map.md) | 투자포인트 1 고리별 찬성·반대 근거, 상태, 다음에 볼 것 |

## 5. 다음 리서치 우선순위

1. **Connect의 제3자 매수 비중** — Kickstart·Fig SIX·특수관계자 매입이 Connect 거래량에 얼마나 들어가는지 (10-Q 주석, 콜 녹취). 투자포인트 1의 [3]이 여기에 달림
2. **은행 진입의 두 역할** — 판매(파트너)와 매수(Connect로 모기지 집적, CRA, 위험가중치). 은행 HELOC 철수·재진입 이력, 은행 파트너 식별
3. **매수자 수·집중도 시계열** — 10-K 참여자 48곳, 상위 3/5 비중 추세, 신규 매수자 이름
4. **경쟁·멀티소싱** — Aven·Pylon·NFTYDoor·SoFi·MAXEX·Better Tinman, 파트너 내 Figure 점유율
5. **전액 인출의 법적 리스크** — Lee Ward v. Figure Lending, TILA 개방형 신용, 2026 중간선거 후 CFPB
6. **4Q 콜(2027년 2월 예상)** — 채널·자산군별 공헌이익 첫 공개

## 6. 업데이트 규칙

작업할 때마다:
1. 새 자료를 읽으면 `research/sources/`에 노트 추가 → 4절 표에 한 줄
2. 새 근거·반증은 `research/IP1_evidence_map.md` 해당 고리에 출처와 함께 추가
3. 고리 상태가 바뀌면 2절 진행 상황 표 갱신
4. 브랜치를 새로 만들거나 끝내면 3절 브랜치 지도 갱신
5. 아래 작업 로그에 날짜와 한 줄

## 7. 작업 로그

| 날짜 | 브랜치 | 내용 |
|---|---|---|
| 2026-11-16 | — | artifact 「FIGR 비즈니스 모델 완전 정리」, 「FIGR 매수자 해부」 작성 |
| 2026-11-17 | `claude/fervent-brown-x3pek4` | 엑스퍼트콜 3개·Greenwich 리포트 2개 상세 정리, artifact 대비 비교, 투자포인트 1 근거 지도, README 허브 구성 |
| 2026-11-17 | `claude/fervent-brown-x3pek4` | artifact 「FIGR 새로 나온 이야기」 발행 |

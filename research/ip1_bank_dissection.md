# 투자포인트 1 리서치: 은행은 왜·얼마나·어떻게 Figure 위에서 대출을 만들게 되는가

- 작성일: 2026-11-17
- 데이터 기준: FIGR 3Q26 실적(2026-11-12), 3Q26 10-Q(2026-11-16), FY2025 10-K(2026-03-16), 2025 HMDA, MBA·Curinos 벤치마크
- 표기: **[회사]** 회사 주장·공시, **[3자]** 업계·애널리스트 자료, **[추정]** 이 메모의 가정 계산, **[숏]** 공매도 리포트 주장

---

> **v2 (2026-11-17 오후) — 심화판:** `research/ip1_bank_channel.html`(아티팩트로 공개)에 FDIC 콜리포트 전수 분석을 추가했다. 원자료는 `research/data/`, 재집계 스크립트는 `research/scripts/fdic_heloc.py`. 아래 v1 메모보다 v2의 숫자를 우선한다.
>
> v2에서 새로 확인한 것 (모두 FDIC RIS 직접 집계, 2026Q2 기준):
> - 은행 HELOC 잔액: 정점 $674.3B(2009Q1) → 저점 $261.3B(2022Q1, −61%) → $307.5B(2026Q2, 저점 대비 +18%). 같은 기간 미국 전체 HELOC는 +52%($481B, NY Fed 2026Q3).
> - 4대 은행(BofA·JPM·Wells·Citi) $330.9B → $52.7B(−84%). Wells는 2021Q4 이후에도 −58%, JPM −27%. Citizens +73%, PNC +44%.
> - 미사용 한도 $443.1B, 인출률 41%. 순상각률은 2009Q4 3.14%가 정점이고, 2018Q3 이후 사실상 0 이하.
> - 4,313개 기관 중 1,050곳이 HELOC 0이고, 그중 872곳은 1순위 모기지를 보유한다. HELOC가 대출의 1% 미만인 1,806곳이 미국 예금의 53.9%($11.2조)를 갖고 있다. JPM·Wells·Citi의 HELOC/대출은 0.78%.
> - 업계 조달비용 2.02%, 운용수익률 5.31%, NIM 3.29%, ROE 13.1%(2026 상반기 연율). 예대율 70% 미만 은행 1,237곳(예금 $11.5조).
> - Connect의 수익 동기는 은행 유형별로 갈린다. 사는 은행: Figure HELOC 매입 시 세후 ROE 23~38%(가정 명시). 파는 은행: 금리위험·집중도·CECL을 덜어낸다. 안 하던 은행: 대차대조표 없이 시작한다.
> - 2026-10 Figure가 $346M 증권화를 접고 통매각을 택했다(HELN). 이유는 예금기관의 강한 매수 수요다.

## 0. 결론 3줄

1. **Figure가 다음에 크게 클 곳은 은행·신협(예금기관)밖에 없다.** 2025년 미국 HELOC 취급액의 약 84%(~$167B)를 예금기관이 만든다. Figure는 나머지 비은행 시장(~16%) 안에서 이미 IMB 점유율 약 19%를 가져갔다. 그런데 HELOC 상위 15개 예금기관($63.9B, 전체의 32%) 가운데 Figure 파트너는 아직 한 곳도 없다.
2. **은행은 HELOC를 못 만드는 게 아니다. 만들어도 돈이 안 된다.** 건당 원가 약 $4,600, 처리 39일, 신청의 49%만 실행, 실행해도 한도의 36%만 인출된다. 여기에 위험가중치 100%(2026년 3월 재제안에서도 유지)가 붙고, 팔 곳(GSE)도 없다. 그 결과 은행·신협의 후순위 담보대출 점유율은 85%에서 66%로 떨어졌다(2022→2025).
3. **Figure는 이 손익 구조를 바꾸고, 은행이 Figure를 경쟁자가 아닌 인프라로 받아들이게 하는 방아쇠가 2026~27년에 몰려 있다.** Figure 위에서는 원가 약 $750, 중위 10일, 인출률 93%이고, 다음 날 매각(Connect)이 가능하며 선불 비용이 없다. 방아쇠는 네 가지다. 대형 예금기관 레퍼런스(Flagstar), 풀 화이트라벨 Connect(4Q26), 은행-핀테크 제휴 규제 완화안(2026-09), 금리 인상에 따른 lock-in 장기화(2026-09)다. 은행 1곳이 붙을 때마다 Figure 연간 HELOC 볼륨의 30~95%가 한 번에 붙는다. 그래서 성장이 계단식으로 폭발한다.

---

## 1. 왜 '은행'이어야 하나: Figure 성장의 천장과 다음 바닥

### 1-1. 시장 구성 (2025)

| 지표 | 값 | 출처 |
|---|---|---|
| HELOC 취급(HMDA, 한도 기준) | 929,780건 / $124.4B | [3자] HELN, 2025 HMDA |
| HELOC 취급(Cotality) | 150만 건 / $271B | [3자] HELN 2025 Home Equity Study |
| HELOC 취급(애널리스트 사용치) | ~$198B, 이 중 예금기관 ~84%(~$167B) | [3자] Ottavi |
| 건수 점유율(HMDA) | 은행 45.6% · 신협 35.5% · 비은행 13.4% | [3자] HELN·Curinos 등 공동 백서 |
| 후순위 담보(HELOC+CES) 점유율 | 은행+신협 85%(2022) → 66%(2025), 비은행 8% → 29% | [3자] 같은 백서 |
| 2023→25 HELOC 취급 증가율 | 비은행 +140% · 지방은행 +20% · 신협 +8% · 대형은행 +7% | [3자] 같은 백서 |
| Figure 플랫폼 HELOC | $8.3B (+62%) | [회사] 10-K |
| Figure 점유율 | 전체 ~4%, IMB 안에서 ~19%(2024년 13%) | [3자] Bernstein(Ottavi 인용) |
| 상위 15개 예금기관 | $63.9B = 전체의 32%. **Figure 파트너로 확인된 곳 0** | [3자] Ottavi 3Q26 |

> 측정 방식(HMDA 한도 기준 / Cotality / 애널리스트 추정)에 따라 총액이 $124B에서 $271B까지 벌어진다. 시나리오에서는 보수적인 HMDA 기반 ~$100B와 Ottavi의 $167B를 함께 쓴다.

**해석:** Figure는 비은행 시장에서 이미 최상위 사업자다. 그런데 비은행 시장은 전체의 16%에 불과하다. IMB(독립 모기지은행)만으로는 성장 폭이 구조적으로 막혀 있다. Figure의 다음 10배는 수학적으로 나머지 84%, 즉 예금기관 풀에서만 나올 수 있다.

### 1-2. Figure는 이미 '남의 이름으로' 대출을 만든다

- 2025년 HMDA에서 대출 주체가 Figure Lending으로 찍힌 HELOC는 $2.5~2.9B다 [3자]. Figure 플랫폼 전체 HELOC는 $8.3B다 [회사].
- 차이 약 $5.4B는 파트너가 **자기 이름(lender of record)** 으로 Figure LOS 위에서 만든 대출이다. 2Q26 기준 파트너 브랜드 비중은 CLM 볼륨의 83%다 [회사].
- 즉 은행을 붙이는 것은 새 사업이 아니다. 지금 IMB와 하고 있는 모델을 그대로 키우는 것이다.

### 1-3. '고래' 한 마리의 크기 (2025 HMDA, 달러 기준)

| 기관 | 유형 | 2025 취급액 | Figure 2025 HELOC($8.3B) 대비 |
|---|---|---|---|
| Citizens | 은행 | $7.9B (HELOC) | 95% |
| Bank of America | 은행 | $7.2B (HELOC) | 87% |
| PNC | 은행 | $4.3B (HELOC) | 52% |
| Navy Federal | 신협 | $4.1B (HELOC) | 49% |
| U.S. Bank | 은행 | $3.3B (HELOC+CES) | 40% |
| Huntington | 은행 | $2.4B (HELOC+CES) | 29% |
| Truist | 은행 | $2.4B (HELOC+CES) | 29% |

출처: [3자] HELN News, 2025 HMDA 순위(2026-10-07). 비율은 [추정].

**은행 하나가 Figure 위로 일부만 옮겨와도 볼륨이 계단식으로 뛴다.** 이것이 "폭발적 성장"의 실제 메커니즘이다. 경영진도 3Q26 콜에서 파이프라인에 "지금 함께하는 어떤 파트너보다 큰" 은행·모기지사가 있다고 말했다 [회사].

---

## 2. 은행의 HELOC 손익 해부: 왜 은행은 시장을 잃고 있나

### 2-1. 운영: 원가·속도·실행률·인출률

| 항목 | 예금기관(전통) | 비은행 | Figure 플랫폼 | 출처 |
|---|---|---|---|---|
| 건당 생산원가 | ~$4,600 (MBA, 2024 전체 비용 기준) / $1,000~2,500 (Curinos) | < $1,000 | ~$717~750 | MBA·Curinos [3자], S-1·경영진 [회사] |
| 신청→실행 기간 | 39일(MBA) / ~40일(Curinos) / 43일(업계) | ~15일 | 중위 10일, 최단 5일 | 같음 |
| 실행률(pull-through) | 49% (MBA 2024) | n/a | n/a | [3자] |
| 실행 시 인출률 | ~36% | ~93% (증권화 HELOC) | 실행 시 전액 인출 | [3자] Curinos |
| 무인 심사 | 낮음 | 높음 | ~90% 사람 심사자 없음 | [회사] |
| 금리(780+ FICO·CLTV<80, 2Q26) | 은행 7.65% · 신협 7.36% | 8.14% | — | [3자] HousingWire/백서 |

추가로 Curinos 데이터가 있다 [3자].
- 1~15일 안에 실행된 고객의 인출 확률은 42%, 30일을 넘긴 고객은 18%다.
- 첫 인출을 하지 않은 고객의 약 40%는 끝내 인출하지 않는다.
- 속도가 느리면 고객은 신용카드·신용대출·핀테크로 이탈한다.

> 핵심 문장 [3자]: 은행의 가장 큰 약점은 "신용정책이 아니라 대출 운영의 비용과 속도"다. 출처는 백서 *Reversing the Tide: Scaling Bank HELOC Lending*(2026)이며, Fifth Third·Curinos·Upstart·Figure 관계자가 공동 저자다.

### 2-2. 자본: 후순위 담보 위험가중치 100%는 그대로다

- 현행 미국 표준방식에서 단독 후순위(HELOC·2순위) 익스포저의 위험가중치는 **100%** 다. 1순위 모기지는 50%다.
- **2026-03-27 Basel III 재제안(OCC·FDIC·Fed):** 1순위 모기지는 LTV 구간별로 50% 이하까지 **낮아진다**. 후순위는 Basel 국제기준을 따르지 않고 **100%를 그대로 둔다** [3자] Debevoise·STB.
- 같은 은행이 1·2순위를 모두 보유하면 합산 처리해야 해서, 2순위가 1순위 위험가중치까지 오염시킬 위험도 생긴다.
- → **상대적으로 HELOC를 보유하는 자본비용이 올라간다.** 그러면 은행에는 "만들되 팔거나 나눠 갖는" 출구가 필요해진다.

### 2-3. 출구: 1순위는 GSE에 팔 수 있지만 HELOC는 팔 곳이 없다

- 1순위 모기지는 Fannie·Freddie라는 표준 매각처가 있다. 은행의 모기지 사업 자체가 "만들어서 GSE에 판다"를 전제로 돌아간다.
- 2순위는 Freddie Mac의 CES 시범사업(18개월, 최대 $2.5B)이 전부였다. 이후 확대는 확인되지 않는다 [3자]. HELOC를 사주는 표준 시장이 없다. 그래서 은행은 **보유하든지, 적극적으로 만들지 않든지** 둘 중 하나다.
- Tannenbaum(CEO, 2Q26): 예금기관은 "Fannie Mae와 Freddie Mac의 큰 고객이다… 그들은 Connect를 원한다(They want Connect)" [회사].
- 같은 콜에서는 이렇게도 말했다. 큰 HELOC 사업을 가진 대형 예금기관이 "익스포저 관리를 위해, 나아가 CRA 적격 자산을 사기 위해 Figure Connect를 활용하려 한다" [회사].
- → Figure Connect는 사실상 **2순위 시장의 GSE 역할**을 하겠다는 것이다. Ottavi의 표현으로는 "비적격 대출의 Fannie/Freddie"다.

### 2-4. 고객: 비은행이 은행 고객을 가져간다

- 비은행 HELOC는 +140%, 대형은행은 +7% 늘었다(2023→25). 은행·신협 점유율은 85%에서 66%로 떨어졌다 [3자].
- Tannenbaum(3Q26): "비은행이 만드는 대출의 상당수는 결국 예금기관의 고객에게 나가는 것"이다. 그래서 예금기관이 "그 자산에 관심이 있다" [회사].
- 수요는 구조적이다.
  - 모기지 보유자의 79%가 6% 미만 금리에 묶여 있다.
  - 2026년 초 2순위 차입의 약 2/3가 2020~22년 4% 미만 금리로 1순위를 받은 사람들이다(ICE).
  - 신용카드 잔액이 $1.23T이고, 부채통합 목적 비중이 25%에서 39%로 늘었다(Curinos) [3자].
- → 은행 입장에서 HELOC는 단독으로는 수익성이 애매하다. 그래도 **안 하면 예금·카드·모기지 고객까지 잃는** 방어 상품이다.

### 2-5. [추정] 한도 $100K HELOC 1건의 은행 손익: 전통 방식 vs Figure 위

가정: 순이자마진 5.0%p(HELOC 금리 ~7.6% − 예금조달 ~2.6%), 서비싱+신용비용은 인출잔액의 연 0.5%, 평균 수명 4년(원가를 4년 상각), 자본은 위험가중자산의 10%(위험가중치 100%), 법인세 21%. Figure 경로는 원가 $750에 Figure 수수료(인출액의 2~3%, 가정)를 더했다.

| 시나리오 | 인출잔액 | 연 순이자 | 연 세후이익 | 투입자본 | ROE |
|---|---|---|---|---|---|
| 전통, MBA 원가 $4,600 | $36,000 | $1,800 | **$371** | $3,600 | ~10% |
| 전통, Curinos 중간 $1,750 | $36,000 | $1,800 | **$934** | $3,600 | ~26% |
| Figure 위, 수수료 2% | $93,000 | $4,650 | **$2,791** | $9,300 | ~30% |
| Figure 위, 수수료 3% | $93,000 | $4,650 | **$2,607** | $9,300 | ~28% |

**읽는 법:** ROE 차이보다 **건당 절대이익이 3~7배** 벌어지는 게 핵심이다. 원인은 인출률 차이(36% vs 93%)다. 같은 고객 1명, 같은 영업 1건에서 이자를 낳는 잔액이 2.6배 생긴다. 은행이 Figure를 쓰는 이유는 "팔기 위해서"만이 아니다. **HELOC 상품 자체가 비로소 돈이 되기 때문이다.**

주의할 점이 두 가지 있다.
- 인출률 차이의 일부는 상품 구조와 신용 구성 차이다. Curinos는 FICO·CLTV를 맞추면 연체 곡선은 비슷하다고 본다.
- Figure형 고정금리 전액 인출 상품을 보유하면 은행에는 금리 리스크(ALM)가 생긴다. 그래서 "보유 일부 + Connect 매각 일부" 조합이 자연스럽다.

---

## 3. 은행의 선택지: 직접 만들기 / 제휴 / 매각 / 매수

- **직접 만들기(Build): JPMorgan Chase.** 2020년 중단 후 2025-08에 재출시했다. 한도의 85% 이상을 실행 시 일시 인출하는 옵션, 한도 $25K~400K, 주택가치의 최대 80%를 내세웠다 [3자]. → **미국 최대 은행조차 상품 설계를 'Figure식 즉시 인출형'으로 바꿨다.** 설계가 맞다는 검증이다.
- Wells Fargo는 2020년 중단 이후 재개가 확인되지 않는다(비공식 출처) [3자]. 이 고객군은 현재 비은행이 흡수하고 있다.
- **나머지 수천 개 은행·신협은 Chase처럼 직접 만들 수 없다.** Curinos(2026-08, Figure·Upstart 등과 웨비나)는 선택지를 이렇게 정리했다. 내부 혁신 / 하이브리드 / 자기 브랜드로 제휴 / 제휴하고 위험을 분산 매각. 그리고 "레거시를 고수하는 것만이 유일하게 지는 자세"라고 했다 [3자].
- **제휴 자리를 두고 경쟁하는 곳:** Upstart, Lending-in-a-Box(Incenter), Achieve Pro, Better, 기존 LOS 벤더.
  - **Figure의 차별점은 '원장(LOS) + 매각처(Connect) + AAA 증권화 + 전자 담보권 등록(DART)'을 한 번에 제공하는 것이다.** 경쟁사 대부분은 이 중 일부만 제공한다.

---

## 4. 어떻게: 은행이 Figure에 들어오는 7가지 경로

| # | 경로 | 은행 역할 | Figure 역할 | 브랜드 | 자금·보유 | Figure 수익 | 상태 |
|---|---|---|---|---|---|---|---|
| 1 | 공동 브랜드 도매(co-branded wholesale) | 고객 소개 | 기술·자금·실행·자산관리 | 공동 | Figure | 취급수수료·매각이익·서비싱 | 운영 중 |
| 2 | 화이트라벨 코레스(correspondent) | 자기 브랜드로 판매 | 심사기준·금리·프로세스 제공 | 은행 | 혼합 | 원금 기준 LOS 수수료 등 | 운영 중(Valley Strong 신협 사례) |
| 3 | **Figure Connect 판매자** | 자기 돈·자기 브랜드로 대출 실행 | 매수자 연결, 다음 날 인수 | 은행 | 은행 → 매각(통매각 또는 증권화) 또는 보유 | 생태계(ecosystem) 수수료 | 운영 중, 볼륨의 72% |
| 4 | **풀 화이트라벨 Connect** | 앞단부터 끝까지 은행 브랜드 | 뒷단 전부 | 은행 | 은행 | 생태계 수수료 | **4Q26 출시 예정** [회사] |
| 5 | LOS만 사용 + 보유·자체 서비싱 | 대출 보유·서비싱 | 기술만 | 은행 | 은행 | 낮은 가격 | 언급됨 [회사] |
| 6 | **매수자** | HELOC 통매각분·AAA 트랜치 매입 | 표준화된 풀 공급 | — | — | 매수 측 수수료 | 대형 예금기관과 논의 중 [회사] |
| 7 | 자금 공급(Prime·창고대출 대체) | 유동성 공급 | 온체인 매칭(50bp) | — | — | Prime 수수료 | 은행은 경쟁자이자 공급자 |

근거 문장 [회사]:
- 10-K: 파트너 수수료는 "LOS에서 실행된 각 대출의 원금 기준"이다. 파트너 계약은 "고정 기간, 자동 갱신"이고 파트너의 "장기 재무적 의무는 없다".
- 파트너 페이지: "선불 비용 없음… 실행한 만큼만 지불(origination revenue-share)", "Figure가 기술·자금·실행·자산관리를 맡고, 실행한 자산을 보유할지 팔지는 은행이 결정한다".
- 3Q26 Tannenbaum: 은행은 "같은 가격을 내지 않을 것"이고, "증권화를 하지 않을 수도 있으며", "때로는 자기 대출을 직접 서비싱할 것"이다. 그러면서 "Figure Connect를 통매각 수단으로 쓰는 은행과 비은행 대출기관에서 생태계 수수료를 받는다"고 했다.
- 3Q26: 예금기관 파트너는 "20곳 이상"이다. 그중 규모가 큰 곳은 Flagstar 하나이며, "다른 은행들의 문을 열어줬다"고 했다.
- "영업 주기는 더 느리지만 파이프라인은 매우 탄탄하다."

**레퍼런스 사례 — Valley Strong 신협:** HousingWire는 핀테크 제휴로 2020년 이후 HELOC 약 $2B를 실행했다고 보도했다. Figure 파트너 페이지에는 이 신협 CEO가 "Figure Connect와 화이트라벨 Figure HELOC"를 전략으로 소개하는 영상이 있다. LendEDU도 이 신협의 HELOC가 Figure 기술 위에서 돈다고 적었다. → 소형 예금기관도 Figure 위에서 HELOC를 수십억 달러 단위로 키울 수 있다는 실증이다.

---

## 5. 얼마나: 시나리오 [추정]

전제:
- Figure 단독 볼륨 런레이트 $20.4B(3Q26 $5.1B × 4)
- 매출 런레이트 $1.15B(3Q26 순매출 $287M × 4)
- 은행 채널 테이크레이트 1~3%. Connect HELOC 총 테이크레이트 가이던스는 2.9~3.4%이고, 은행은 그보다 낮게 낸다는 경영진 발언을 반영했다.

| 예금기관 풀 | 침투율 | 추가 볼륨 | 단독 런레이트 대비 | 추가 매출(1% / 2% / 3%) | 매출 런레이트 대비(2%) |
|---|---|---|---|---|---|
| ~$100B (HMDA 기반 보수) | 3% | $3.0B | +15% | $30M / $60M / $90M | +5% |
| ~$100B | 10% | $10.0B | +49% | $100M / $200M / $300M | +17% |
| ~$100B | 20% | $20.0B | +98% | $200M / $400M / $600M | +35% |
| ~$167B (Ottavi) | 3% | $5.0B | +25% | $50M / $100M / $150M | +9% |
| ~$167B | **10%** | **$16.7B** | **+82%** | $167M / **$334M** / $501M | **+29%** |
| ~$167B | 20% | $33.4B | +164% | $334M / $668M / $1,002M | +58% |

이 표에 더해 고려할 점:
- **Figure factor [회사]:** 2025년 HMDA 데이터가 있는 185개 파트너는 Figure 도입 전 기준선 대비 HELOC를 2.6배 더 실행했다. 은행이 붙으면 은행 자체의 HELOC 풀이 커질 여지가 있다.
  - 단, Ottavi의 지적처럼 은행 시장은 "새 수요 창출"보다 "기존 HELOC 사업 가속과 점유 회수"에 가깝다. 2.6배는 IMB 중심 표본이므로 은행에 그대로 적용하면 안 된다.
- **매출 질:** 은행 채널은 테이크레이트가 낮다. 대신 고객획득비용이 없고, Figure의 자본을 쓰지 않는다. 경영진은 Connect를 "자본 집약도가 가장 낮고 공헌이익률이 높은" 채널로 설명한다. 공헌이익률은 조정 EBITDA 마진(58%)보다 높아야 한다고 했다 [회사].
- 은행 볼륨이 들어오는 손익 계정은 **생태계·기술 수수료**다. 3Q26 $109.8M으로 전년 $35.7M 대비 약 3.1배다 [회사, 10-Q].
- **결론:** 볼륨은 '폭발', 매출은 '가속', 이익은 '고마진 가속'이다. 투자포인트에서는 볼륨 폭발과 저자본 수수료 매출을 분리해서 써야 반론에 강하다.

---

## 6. 왜 지금인가: 2026~27 방아쇠

1. **레퍼런스와 파이프라인 [회사]**
   - Flagstar(상장 은행)가 최근 몇 달 사이 가동했다. 예금기관 파트너는 20곳 이상이다.
   - 2Q26: "큰 HELOC 사업을 가진 대형 예금기관과 접촉 시작"
   - 3Q26: "미국 최대 모기지사·은행 일부가 파이프라인에 있다, 지금의 어떤 파트너보다 크다"
2. **상품 [회사]**
   - 풀 화이트라벨 Connect가 4Q26에 나온다. 은행의 최대 반론인 브랜드·고객 소유 문제를 제거한다.
   - 1순위 HELOC가 볼륨의 ~20%다(2026은 "1순위의 해").
   - 구매 목적(purchase) 상품을 곧 출시한다. 2027은 "업셀의 해"다.
3. **규제 [3자]**
   - 2026-09-11 Fed·OCC·FDIC·NCUA가 제3자 위험관리(TPRM) 가이던스 전면 개편안을 냈다. 2023년 가이던스와 2024년 은행-핀테크 공동성명을 폐지하고, '중요 활동(critical activities)' 개념을 삭제하며, 핀테크 제휴를 '혁신의 원천'으로 재정의한다.
   - → 은행의 벤더 심사·승인 기간을 줄이는 방향이다(아직 제안 단계).
   - 2026-03-27 Basel 재제안은 후순위 100%를 유지하고 1순위만 완화했다. → 보유 대신 매각·분산 유인이 커진다.
4. **금리 [3자]**
   - 2026-09-16 Fed가 2023년 이후 처음으로 25bp 인상해 3.75~4.00%가 됐다. 점도표는 추가 인상을 시사한다.
   - 재융자(cash-out refi) 창구는 닫히고 lock-in은 길어진다. HELOC 대체 수요에는 우호적이다.
   - 단, Prime 금리가 오르면 HELOC 금리도 오른다. 양날의 검이다.
5. **경쟁 압력 [3자]**
   - 홈에쿼티 1위는 Rocket($9.6B, Nationstar $1.8B 별도)이다. 비은행이 상위권을 잠식하면서 은행에 "제휴하지 않으면 잃는다"는 압력이 생긴다.

---

## 7. 은행이 망설이는 이유와 Figure의 답

| 은행의 반론 | 실체 | Figure의 답 | 남는 리스크 |
|---|---|---|---|
| "예금이 싼데 왜 파나" | 보유하면 순이자마진 확보 | 보유 옵션, LOS만 쓰는 경로 | 100% 위험가중치, 고정금리 ALM |
| "고객과 브랜드를 뺏긴다" | Figure 자체 브랜드(직판)가 여전히 존재 | 풀 화이트라벨(4Q26), B2B 80% 이상 | 자체 브랜드 비중 ~17%(2Q26)와의 채널 충돌 |
| "벤더·모델·공정대출 리스크" | AVM·AI 심사·전자공증 | TPRM 완화 흐름, 자동 검증(Adaptor) | SOC 등 보안 인증이 공개 자료에서 확인 안 됨 |
| "평판 리스크" | [숏] Morpheus(2026-04-16): 타이틀 리포트 없이 실행, AVM 약점, 전액 인출 상품의 open-end 분류 위험(CFPB·집단소송) | 회사 측 반박 자료는 이 메모 범위 밖 | 분류가 바뀌면 "중대한 악영향"이라고 회사가 공시 |
| "영업 주기가 길다" | 은행 조달·심사 프로세스 | 레퍼런스(Flagstar), AI 온보딩 | 실적 반영 시점이 2027 이후로 밀릴 수 있음 |
| "가격" | 은행은 낮은 수수료 요구 | 볼륨 구간별 요금 | 테이크레이트 희석(순 4.4% → 3.5%, 1년간) |
| "신용" | 비은행 HELOC 30일 이상 연체 2.4% vs 예금기관 1.4%(원시값) | 평균 FICO 751, CLTV 62% | 대규모 하락기를 겪은 적 없는 북 |
| "규제당국이 막으면" | 10-K: 파트너의 "규제당국이 우리와의 거래를 종료·제한하도록 요구할 수 있다" | — | 은행 파트너일수록 민감 |

참고: Flagstar는 2Q26에 보유하던 Figure 지분을 전량 매각했다(차익 $3.5M) [3자, FLG 2Q26 콜]. 그러면서도 오리지네이션 파트너 관계는 유지·확대 중이다. 지분 매각을 관계 약화로 읽을 근거는 없지만, 질의응답에서 나올 수 있으니 알아둔다.

---

## 8. 투자포인트 1 검증용 트래킹 KPI

- [ ] 예금기관 파트너 수(20곳 이상 → ?)와 **이름이 공개되는 대형 은행 첫 사례**
- [ ] 풀 화이트라벨 Connect 출시 여부와 첫 은행 고객(4Q26 콜, 2027년 2~3월)
- [ ] Connect 비중(단독 72%)과 생태계·기술 수수료 증가율
- [ ] 채널별 공헌이익률(4Q26 콜에서 공개 예정이라고 언급)
- [ ] 2026년 HMDA(2027년 봄 공개): Flagstar 등 파트너 은행의 HELOC 증가율로 'Figure factor'를 은행 표본에서 검증
- [ ] 집중도: 상위 10개 파트너가 볼륨의 56%(2025), 상위 3개 매수자가 매각의 57%(2025). 은행이 붙을수록 집중도가 올라가는지 분산되는지
- [ ] 은행·보험사의 Connect 매수 참여(현재 New York Life·Nationwide 등 보험사 중심)
- [ ] TPRM 개편안 최종 확정 여부

---

## 9. 투자포인트 1 문안 초안

**헤드라인 후보**
- A. "은행은 HELOC를 못 하는 게 아니라 돈이 안 돼서 안 한다. Figure가 그 손익을 뒤집는 순간 시장의 84%가 열린다"
- B. "2순위 시장의 Fannie Mae: 은행이 Figure를 쓰는 이유"
- C. "IMB 다음은 은행이다. Figure 성장의 두 번째 엔진"

**본문 초안 (3단락)**

미국 HELOC 시장의 84%는 여전히 은행과 신협이 만든다. Figure는 비은행 시장에서 IMB 점유율 19%까지 올라왔지만, 상위 15개 예금기관($63.9B, 전체의 32%)은 아직 한 곳도 Figure 위에 있지 않다. Figure의 다음 성장은 이 84%에서 나올 수밖에 없다.

은행이 HELOC를 놓치는 이유는 신용이 아니라 손익이다. 한 건을 만드는 데 39일과 $4,600이 들고, 신청의 절반만 실행된다. 실행된 한도도 36%만 인출되어, 이자를 낳는 잔액이 작다. 후순위 대출의 위험가중치는 2026년 재제안에서도 100%로 남았다. 1순위와 달리 GSE라는 매각처도 없다. 그 사이 비은행 HELOC는 2년 동안 140% 늘었고, 은행·신협 점유율은 85%에서 66%로 떨어졌다.

Figure는 이 세 문제(원가·자본·출구)를 한 번에 푼다. 10일 실행, 건당 약 $750, 93% 인출로 같은 영업 1건이 만드는 이자자산이 2.6배가 된다. Connect의 다음 날 인수로 보유와 매각을 은행이 고를 수 있다. 4Q26 풀 화이트라벨 Connect로 브랜드 반론도 사라진다. Flagstar라는 레퍼런스와 2026년 9월 TPRM 완화안, 금리 인상에 따른 lock-in 장기화가 동시에 방아쇠가 된다. 예금기관 HELOC의 10%만 들어와도 Figure 볼륨은 현재 런레이트 대비 약 80% 늘어난다. 고객획득비용과 자본이 들지 않는 수수료 매출이므로 이익 레버리지는 더 크다.

---

## 출처

**회사 공시·콜**
- [Figure 3Q26 실적 보도자료 (GlobeNewswire, 2026-11-12)](https://www.globenewswire.com/news-release/2026/11/12/3401369/0/en/figure-technology-solutions-reports-third-quarter-2026-results.html)
- [Figure 3Q26 실적콜 녹취 (gloom.sh)](https://gloom.sh/stocks/figr/transcripts/q3-2026)
- [Figure 3Q26 실적콜 녹취 (Motley Fool via AOL)](https://www.aol.com/articles/figure-tech-figr-q3-2026-234514000.html)
- [Figure 2Q26 실적콜 녹취 (Motley Fool)](https://www.fool.com/earnings/call-transcripts/2026/08/20/figure-technologies-figr-q2-2026-earnings-call-transcript/)
- [Figure 2Q26 보도자료 (SEC Ex.99.1)](https://www.sec.gov/Archives/edgar/data/0002064124/000206412426000031/ex991-pressrelease2q26.htm)
- [Figure 3Q26 10-Q](https://www.sec.gov/Archives/edgar/data/2064124/000162828026074112/figr-20260930.htm)
- [Figure FY2025 10-K](https://www.sec.gov/Archives/edgar/data/2064124/000206412426000009/figr-20251231.htm)
- [Figure 은행·신협 파트너 페이지](https://www.figure.com/partner-heloc/banks-and-credit-unions/), [파트너 HELOC](https://www.figure.com/partner-heloc/), [Figure Connect](https://www.figure.com/partner/figure-connect/)

**업계 데이터**
- [HousingWire: Nonbank HELOC lenders gain share (2026-09-14, 백서 *Reversing the Tide* 요약)](https://www.housingwire.com/articles/nonbank-heloc-share-equity/)
- [HELN: Home Equity Lender Ranking Topped by Rocket (2026-10-07, 2025 HMDA)](https://www.hel.news/articles/bank-home-equity-news/biggest-lenders-100626/)
- [HELN: Study — HELOC Lending Share Up as Limits Top $1T (2026-04-07)](https://www.hel.news/articles/bank-home-equity-news/study-originations-040726/)
- [Curinos: Unlocking Today's Home Equity Market (2026-08-19)](https://www.curinos.com/insights/curinos-perspective-unlocking-todays-home-equity-market-5-takeaways)
- [Curinos: Home Equity Speed Matters (2026-03-29)](https://www.curinos.com/insights/curinos-perspective-home-equity-speed-matters-5-takeaways)
- [Curinos: Home Equity Lending 2015 vs. 2025 (2025-03-26)](https://curinos.com/our-insights/curinos-perspective-home-equity-lending-2015-vs-2025-5-takeaways/)
- [MBA Home Equity Lending Study 보도 (2025-07-28)](https://www.mba.org/news-and-research/newsroom/news/2025/07/28/mba-home-equity-study-shows-increase-in-originations--debt-outstanding-in-2024)
- [HousingWire: MBA 2024 home equity debt / $4,500+ cost](https://www.housingwire.com/articles/mba-total-heloc-home-equity-loan-debt-outstanding-originations-marina-walsh/)
- [Banking Dive: JPMorgan Chase relaunches HELOC (2025-08)](https://www.bankingdive.com/news/jpmorgan-chase-relaunches-heloc-loans-after-5-year-break/758757/), [HousingWire: HELOC returns to Chase](https://www.housingwire.com/articles/the-heloc-product-returns-to-chase-home-lending/)
- [FHFA Freddie Mac 2순위 시범사업 승인 (NMP)](https://nationalmortgageprofessional.com/news/fhfa-approves-freddie-macs-second-mortgage-pilot)
- [LendEDU: Valley Strong HELOC 리뷰](https://lendedu.com/blog/valley-strong-heloc-review/)

**규제·거시**
- [Debevoise: Basel III Endgame Mulligan (2026-03)](https://www.debevoise.com/insights/publications/2026/03/federal-banking-agencies-basel-iii-endgame-mullig), [Simpson Thacher: Basel III Endgame Evolution — Retail Banking (2026-03-31)](https://www.stblaw.com/about-us/publications/view/2026/03/31/basel-iii-endgame-evolution-strategic-implications-for-retail-banking)
- [Hunton: 2023 TPRM 가이던스 대체·은행-핀테크 공동성명 폐지 제안](https://www.hunton.com/insights/legal/banking-agencies-propose-to-replace-the-2023-third-party-risk-management-guidance-and-to-rescind-the-bank-fintech-joint-statement), [Skadden (2026-09)](https://www.skadden.com/insights/publications/2026/09/us-federal-banking-agencies-propose-revised-third-party)
- [Federal Reserve FOMC 성명 (2026-09-16)](https://federalreserve.gov/newsevents/pressreleases/monetary20260916a.htm)

**애널리스트·숏 리포트**
- [Ottavi: Figure (FIGR) Q3 2026 Update (2026-11-13)](https://ottavi.substack.com/p/figure-figr-q3-2026-update)
- [Ottavi: Building the Fannie Mae & Freddie Mac of Non-Qualified Lending (2026-02-04)](https://ottavi.substack.com/p/figure-technologies-part-1-building)
- [Chara Capital: The Infrastructure Layer of HELOCs](https://charainvest.substack.com/p/the-infrastructure-layer-of-helocs)
- [Morpheus Research 숏 리포트 (2026-04-16)](https://www.morpheus-research.com/figure/)
- [Flagstar 2Q26 실적콜 (Figure 지분 매각 언급)](https://finance.yahoo.com/quote/FLG/earnings/FLG-Q2-2026-earnings_call-654371.html)

## 확인하지 못한 것 (보고서에 쓰기 전 체크)

- 상위 15개 예금기관 $63.9B의 기관별 표와 원 출처 연도: Ottavi 본문에 표가 없다(이미지로 추정).
- Flagstar–Figure 제휴 공식 보도자료: Ottavi와 경영진 발언으로만 확인했다.
- 'Figure factor 2.6배'의 파트너 유형별(은행 vs IMB) 분해: 원문 기사는 접근이 막혔다(403).
- 은행 채널 실제 테이크레이트: 미공시. 표의 1~3%는 가정이다.
- Fifth Third가 "생산 2배 이상"을 달성할 때 쓴 제휴사: 공개 자료에서 Figure로 확인되지 않는다. Figure 사례로 쓰면 안 된다.

# 은행 채널 보강 (투자포인트 1)

`bank_channel_evidence.html`은 "은행 채널 해부" 페이지를 보강하는 페이지입니다. 은행이 실제로 Figure를 쓰러 오고 있는지를 HMDA, FDIC, 연준 SLOOS, Figure 실적콜로 확인합니다. 데이터는 2026-11-17 기준입니다.

## 핵심 숫자

| 숫자 | 뜻 | 파일 |
|---|---|---|
| 0.6% | 미국 은행이 2025년에 만든 HELOC 중 같은 해에 판 비중. 신협은 2.6%, 비은행은 50.8% | `data/hmda_second_lien_by_lender_type.csv` |
| 24~46% vs 0% | Citizens·PNC·Fifth Third·Flagstar의 2025년 매각 비중. 1순위 모기지는 24~46%(금액 기준), HELOC는 0% | `data/hmda2025_sold_share_by_lender.csv` |
| $1.6B → $4.3B | 은행·저축기관이 사들인 2순위 담보대출(2023→2025). Discover를 빼면 $3.6B. 은행이 산 HELOC의 72%는 UWM 대출 | `data/hmda_second_lien_bought_by_banks.csv` |
| $123M | Figure Lending이 2025년에 은행에 판 HELOC. 은행 HELOC 매입의 6% | `data/figure_hmda_purchaser_mix.csv` |
| −83% | Flagstar의 HELOC 실행 한도, 2022 $1.20B → 2025 $0.20B | `data/flagstar_hmda_2022_2025.csv` |
| 90.6% | Flagstar 예대율(2026Q2). 대출의 2.2%가 HELOC $1.36B | `data/flagstar_fdic_quarterly.csv` |
| +14 → 0 | SLOOS HELOC 수요 순응답, 2025Q3 → 2026Q3 | `data/sloos_heloc_demand.csv` |

## 다시 만들기

1. 기관별 LAR: `https://ffiec.cfpb.gov/v2/data-browser-api/view/csv?years=2025&leis=<LEI>&actions_taken=1`를 리다이렉트를 따라 받습니다(`curl -L --compressed`). 쓴 LEI는 다음과 같습니다. Figure Lending `254900UL88QFG0E40516`, Flagstar `SS1TRMSN6BRNMOREEV51`, Citizens `DRMSV1Q0EKMEXLAU1P80`, PNC `AD6GFRVSDT01YPT1CS68`, Fifth Third `QFROUN1UWUYU0DVIWD51`, Valley Strong `5493006MA7WP1WL8U431`, loanDepot `549300AG64NHILB7ZP05`.
2. 전국 2순위: `.../view/nationwide/csv?years=<연도>&actions_taken=1&lien_statuses=2`를 `scripts/agg_second_liens.py <연도> sub_<연도>.json`으로 흘려 집계합니다. 연도당 약 130만~160만 건입니다.
3. 기관명 목록은 `https://ffiec.cfpb.gov/v2/reporting/filers/<연도>`에서 받아 `filers<연도>.json`으로 저장합니다.
4. `python3 -I scripts/build_tables.py <hmda_dir> data`를 실행하면 `data/`의 표가 만들어집니다.

## 한계

- HMDA는 실행한 해 안에 판 대출만 매각으로 기록합니다. 매각과 매입 규모는 하한입니다.
- 기관 유형(은행·신협·비은행)은 기관 이름으로 나눴습니다.
- HELOC 금액은 인출액이 아니라 한도입니다. 역모기지와 사업 목적 대출은 뺐습니다.
- 매입자 코드 71은 신협, 모기지회사, 금융회사를 하나로 묶은 코드라서 신협 몫만 따로 볼 수 없습니다.

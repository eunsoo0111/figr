# 도표 3-1 원자료(달러, 1의 자리). None = 1차 출처 미확보(도표에 n/a로 표시).
# 행별 수식(역산·합산·누적 차이)은 ROW_FORMULA로 엑셀에 수식 그대로 남긴다.
YEARS = ["'04", "'05", "'06", "'07", "'08", "'09", "'10", "'11", "'12", "'13", "'14", "'15", "'16", "'17",
         "'18", "'19", "'20", "'21", "'22", "'23", "'24", "'25", "'26E"]

ROW = {y: i + 2 for i, y in enumerate(YEARS)}            # 엑셀 행 번호(머리행 1)

MKT = {y: None for y in YEARS}
MKT.update({"'05": 60_740_000_000,
            **{f"'{y:02d}": 0 for y in range(8, 19)},
            "'19": 277_700_000, "'21": 150_000_000,
            "'22": 1_000_000_000, "'25": 28_997_000_000, "'26E": 41_000_000_000})
FIG = {y: 0 for y in YEARS}
FIG.update({"'06": None, "'07": None, "'20": 149_000_000, "'23": 664_000_000, "'24": 2_300_000_000})

# 수식 행: B = 시장 합계, C = 동사(Figure)
ROW_FORMULA = {
    ("B", "'04"): f"=B{ROW[chr(39)+'05']}/(1+23.6%)",                     # IMF: 05년 607.4억 달러, 전년 대비 +23.6%
    ("B", "'20"): "=149000000+308000000",                                  # FLOC 2020-1 + GRADE 2020-FIG1
    ("B", "'23"): "=1900000000+2900000000",                                # HEL News: CES 19억 + HELOC 29억
    ("B", "'24"): "=10200000000+5400000000",                               # HEL News: CES 102억 + HELOC 54억
    ("C", "'25"): f"=C{ROW[chr(39)+'24']}+1200000000",                    # 10-K: 증권화 규모 12억 달러 증가
    ("C", "'26E"): f"=C{ROW[chr(39)+'25']}-(7000000000-5800000000)+4100000000",  # 25.4Q 제외 + 10-Q 9M 증가 41억
}

MKT_SRC = {
    "'04": "SMIC 역산: IMF 'Second-Lien MBS Issuance Up Sharply in 2005'(2006) — 05년 607.4억 달러, 전년 대비 +23.6%",
    "'05": "Inside Mortgage Finance(2006), 'Second-Lien MBS Issuance Up Sharply in 2005' — second-lien MBS 607.4억 달러",
    "'06": "원자료 미확보 — IMF 06년 second-lien MBS 연간 발행액 필요(참고: S&P 06.4Q CES 135.4억+HELOC 39억)",
    "'07": "원자료 미확보 — IMF 07년 연간 발행액 필요(참고: IMF 07.1Q 215.7억, S&P 07.1Q CES 111.2억+HELOC 72.8억)",
    "'26E": "BofA Securities 전망(American Banker, 2026.7.27) — 'B of A expects $41 billion ... for 2026'",
    "'25": "J.P. Morgan·Bloomberg 'RMBS Credit Issuance to Date' HELOC/CES 2025 FY 28,997(백만 달러), SFA Research Corner(2026.9.17)",
    "'24": "HEL News CEO Letter(2024.12.27) — CES 41건 102억 + HELOC 18건 54억 달러",
    "'23": "HEL News CEO Letter(2024.12.27) — CES 8건 19억 + HELOC 13건 29억 달러",
    "'22": "S&P Global Ratings(SFA Research Corner 2026.9.17 인용) — 'about $1 billion in 2022'",
    "'21": "확인된 딜 합계(하한): GRADE 2021-FIG2 1.5억(Saluda Grade 2021.3.9)",
    "'20": "확인된 딜 합계(하한): FLOC 2020-1 1.49억(DBRS 2023.4.14) + GRADE 2020-FIG1 3.08억(Saluda Grade 2020.9.15)",
    "'19": "확인된 딜 합계(하한): Towd Point HE Trust 2019-HE1 풀 2.777억(Fitch, National Mortgage News 2019.6.14)",
}
for y in range(8, 19):
    MKT_SRC[f"'{y:02d}"] = ("서술 기반 0 표기(실측 아님): KBRA(2026.9.8) 'essentially nonexistent for years following the GFC', "
                            "HousingWire(2023.10.5) 'essentially stopped for about a decade'")

FIG_SRC = {
    "'20": "DBRS Morningstar(2023.4.14) — FLOC 2020-1 'approximately $149 million'(동사 첫 증권화, 무등급)",
    "'21": "동사 스폰서 딜 없음(DBRS 'Previously, Figure sponsored FLOC 2020-1', 10-K 2025 Exhibit 21)",
    "'22": "동사 스폰서 딜 없음(위와 같음)",
    "'23": "Figure S-1(2025.8.18) — 2023년 동사 보유 대출 증권화 UPB 6.64억 달러(FIGRE 2023-HE1·HE2·HE3)",
    "'24": "Finsight(GlobalCapital 인용) — 'over $2.3bn through seven Heloc RMBS deals in 2024'(하한)",
    "'25": "SMIC 계산: 24년 + Figure 10-K(2025) '$1.2 billion increase in the volume of securitizations'",
    "'26E": "SMIC 계산(26.3Q 누적 실적): 25년 − 25.4Q(10-K 누적 70억 − 424B4 25.9말 누적 58억) + 10-Q(26.3Q) 9M '$4.1 billion increase'",
}

EVENTS = {"'07": "07~08년 모노라인 붕괴", "'13": "08~18년 사실상 0", "'19": "19년 위기 뒤 첫 HELOC 딜",
          "'20": "20년 동사 첫 딜", "'24": "24년 Connect 출범"}
EVENT_SRC = {
    "'07": "S&P, MBIA·Ambac AAA→AA 강등(2008.6.5, InvestmentNews); FSA HELOC 8건 손실 2.661억(Bond Buyer 2008.5.15); "
           "'Pre-crisis, you couldn't get a AAA on a Heloc-backed bond without a bond wrap'(Bloomberg 2019.6.26)",
    "'13": "KBRA(2026.9.8), HousingWire(2023.10.5)",
    "'19": "Towd Point HE Trust 2019-HE1(FirstKey·Cerberus) — Fitch 'first all-HELOC securitization in the post-crisis era'(NMN 2019.6.14)",
    "'20": "FLOC 2020-1, 블록체인 기반 첫 증권화(The Block 2020.3.11, DBRS 2023.4.14)",
    "'24": "Figure 424B4(2025.9.11) 'In June 2024, we launched Figure Connect'",
}

MKT = [MKT[y] for y in YEARS]
FIG = [FIG[y] for y in YEARS]

NOTE = ("연도별 집계 기관 상이(04·05 IMF, 22 S&P, 23·24 HEL News, 25 J.P. Morgan, 26E BofA 전망), 04년은 05년 증가율로 역산, 06·07년은 원자료 확인 중(n/a)"
        " / 08~18년은 '사실상 0' 서술 반영, 19~21년은 확인된 딜 합계(하한), 동사 25년·26E(26년 3분기 누적)는 10-K·10-Q 증가액으로 SMIC 계산")

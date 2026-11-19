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
FIG.update({"'06": None, "'07": None})          # 20년 이후는 딜 목록 시트에서 SUMIF 수식(FIG_YEARS)

# 수식 행: B = 시장 합계, C = 동사(Figure)
ROW_FORMULA = {
    ("B", "'04"): f"=B{ROW[chr(39)+'05']}/(1+23.6%)",                     # IMF: 05년 607.4억 달러, 전년 대비 +23.6%
    ("B", "'20"): "=149000000+308000000",                                  # FLOC 2020-1 + GRADE 2020-FIG1
    ("B", "'23"): "=1900000000+2900000000",                                # HEL News: CES 19억 + HELOC 29억
    ("B", "'24"): "=10200000000+5400000000",                               # HEL News: CES 102억 + HELOC 54억
}
DEAL_SHEET = "도표 3-1 동사 딜"
FIG_YEARS = {"'20": "2020년", "'21": "2021년", "'22": "2022년", "'23": "2023년", "'24": "2024년", "'25": "2025년", "'26E": "2026년"}
for y, k in FIG_YEARS.items():                # 동사 = 딜 목록(B 규모, C 연도) 연도별 합계
    ROW_FORMULA[("C", y)] = f"=SUMIF('{DEAL_SHEET}'!$C:$C,\"{k}\",'{DEAL_SHEET}'!$B:$B)"

# 동사(Figure) 스폰서 HELOC 증권화 딜 — 규모 = Morningstar DBRS '총 미상환 원금(UPB)'(컷오프 기준, 1달러 단위)
# PF(선펀딩) 딜은 선펀딩 포함 딜 규모(up to), KBRA 딜은 KBRA 발표 딜 규모. Figure 공시 누적(25.6말 15건 45억·25.9말 19건 58억·25말 23건 70억)과 일치.
D = "dbrs.morningstar.com/research/"
DEALS = [  # (딜, 연도, 일자, 규모, 기준, 출처)
    ("FLOC 2020-1", "2020년", "2020.3", 149_000_000, "약치('approximately')", "DBRS 412701(2023.4.14) 'approximately $149 million'; The Block(2020.3.11)"),
    ("FIGRE 2023-HE1", "2023년", "2023.4.14 F", 236_717_981, "UPB", D + "412701"),
    ("FIGRE 2023-HE2", "2023년", "2023.6.5 P", 205_444_368, "UPB", D + "415418"),
    ("FIGRE 2023-HE3", "2023년", "2023.12.7 P", 195_800_000, "KBRA 딜 규모", "KBRA(businesswire 2023.12.7) '$195.8 million'"),
    ("FIGRE 2024-HE1", "2024년", "2024.3.13 P", 301_419_049, "UPB", D + "429315"),
    ("FIGRE 2024-HE2", "2024년", "2024.5.16 P", 396_540_914, "UPB", D + "432872"),
    ("FIGRE 2024-HE3", "2024년", "2024.7.10 P", 379_535_977, "UPB", D + "435974"),
    ("FIGRE 2024-HE4", "2024년", "2024.9.6 P", 346_802_432, "UPB", D + "439111"),
    ("FIGRE 2024-HE5", "2024년", "2024.10.8 P", 464_810_939, "UPB", D + "440961"),
    ("FIGRE 2024-SL1", "2024년", "2024.10.31 P", 220_100_000, "KBRA 딜 규모", "KBRA(2024.10.31) 'a $220.1 million RMBS transaction'"),
    ("FIGRE 2024-HE6", "2024년", "2024.12.19 F", 322_339_191, "UPB", D + "444939"),
    ("FIGRE 2025-HE1", "2025년", "2025.2.10 F", 306_711_004, "UPB", D + "447668"),
    ("FIGRE 2025-HE2", "2025년", "2025.3.28 F", 311_206_542, "UPB", D + "450820"),
    ("FIGRE 2025-PF1", "2025년", "2025.5.6 F", 300_000_000, "선펀딩 포함(up to, 마감 UPB 125,030,921)", D + "453523"),
    ("FIGRE 2025-HE3", "2025년", "2025.6.13 F", 354_619_566, "UPB", D + "456293"),
    ("FIGRE 2025-HE4", "2025년", "2025.7.10 F", 301_329_484, "UPB", D + "458159"),
    ("FIGRE 2025-FL1", "2025년", "2025.8.8 F", 234_462_485, "UPB", D + "460341"),
    ("FIGRE 2025-HE5", "2025년", "2025.8.28 F", 479_792_904, "UPB", D + "461709"),
    ("FIGRE 2025-PF2", "2025년", "2025.9.24 F", 300_077_650, "선펀딩 포함(up to, 마감 UPB 120,077,650)", D + "463142"),
    ("FIGRE 2025-HE6", "2025년", "2025.10.22 F", 296_043_061, "UPB", D + "465392"),
    ("FIGRE 2025-HE7", "2025년", "2025.11.25 F", 368_551_632, "UPB", D + "468140"),
    ("FIGRE 2025-HE8", "2025년", "2025.12.2 F", 435_986_235, "UPB", D + "468898"),
    ("FIGRE 2025-FL2", "2025년", "2025.12.18 F", 242_341_393, "UPB", D + "470215"),
    ("FIGRE 2026-HE1", "2026년", "2026.2.3 F", 410_828_822, "UPB", D + "473359"),
    ("FIGRE 2026-HE2", "2026년", "2026.3.12 F", 504_473_504, "UPB", D + "476260"),
    ("FIGRE 2026-HF3", "2026년", "2026.3.19 F", 296_306_884, "UPB", D + "476780"),
    ("FIGRE 2026-FL1", "2026년", "2026.3.30 F", 274_586_095, "UPB", D + "477810"),
    ("FIGRE 2026-HE4", "2026년", "2026.5.7 F", 383_933_694, "UPB", D + "480518"),
    ("FIGRE 2026-HE5", "2026년", "2026.6.3 F", 365_976_352, "UPB", D + "482331"),
    ("FIGRE 2026-FL2", "2026년", "2026.6.18 F", 346_117_864, "UPB", D + "483484"),
    ("FIGRE 2026-HE6", "2026년", "2026.6.30 F", 359_775_190, "UPB", D + "484398"),
    ("FIGRE 2026-EXP1", "2026년", "2026.7.10 F", 321_104_032, "UPB", D + "485114"),
    ("FIGRE 2026-HF7", "2026년", "2026.7.28 F", 319_901_741, "UPB", D + "486376"),
    ("FIGRE 2026-HE8", "2026년", "2026.8.13 F", 404_296_027, "UPB", D + "487368"),
    ("FIGRE 2026-FL3", "2026년", "2026.8.18 F", 354_244_271, "UPB", D + "487602"),
    ("FIGRE 2026-HE9", "2026년", "2026.8.25 F", 326_462_318, "UPB", D + "487947"),
    ("FIGRE 2026-EXP2", "2026년", "2026.8.31 F", 316_334_341, "UPB", D + "488263"),
    ("FIGRE 2026-HE10", "2026년", "2026.9.16 F", 370_844_800, "UPB", D + "489148"),
    ("FIGRE 2026-HF11", "2026년", "2026.9.28 F", 321_068_502, "UPB", D + "489995"),
    ("FIGRE 2026-EXP3", "2026년", "2026.10.9 F", 345_628_456, "UPB", D + "491024"),
    ("FIGRE 2026-HE12", "2026년", "2026.10.7 P", 345_970_626, "UPB(가평가)", D + "490807"),
    ("FIGRE 2026-FL4", "2026년", "2026.11.12 F", 257_557_285, "UPB", D + "493392"),
    ("FIGRE 2026-HE13", "2026년", "2026.11.4 15Ga-2", None, "규모 미공개(대출 2,967건)", "SEC ABS-15G acc 0001193125-26-446956"),
    ("FIGRE 2026-EXP4", "2026년", "2026.11.16 15Ga-2", None, "규모 미공개(대출 4,936건)", "SEC ABS-15G acc 0001193125-26-463924"),
]

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
    "'20": "딜 목록 시트 합계: FLOC 2020-1(동사 첫 증권화, 무등급)",
    "'21": "동사 스폰서 딜 없음(DBRS 'Previously, Figure sponsored FLOC 2020-1'; 10-K 2025 누적 23건 = FLOC + FIGRE 22건)",
    "'22": "동사 스폰서 딜 없음(위와 같음)",
    "'23": "딜 목록 시트 합계(3건)",
    "'24": "딜 목록 시트 합계(7건) — Finsight 'over $2.3bn through seven' 과 일치",
    "'25": "딜 목록 시트 합계(12건) — 10-K 누적 23건·70억 달러와 일치",
    "'26E": "딜 목록 시트 합계: 2026.11.19까지 확인된 19건(HE13·EXP4 규모 미공개, HE12는 가평가 기준) — 연간 아님",
}

EVENTS = {"'07": "07~08년 모노라인 붕괴\n08년 S&P CES 평가 중단", "'22": "22년 금리 급등·락인", "'13": "08~18년 사실상 0", "'19": "19년 위기 뒤 첫 HELOC 딜",
          "'20": "20년 동사 첫 딜", "'24": "24년 Connect 출범"}
EVENT_SRC = {
    "'07": "IMF(2008.5.8) 'S&P Stops Rating Second-Lien MBS Deals'(CES 담보 RMBS 평가 중단); S&P, MBIA·Ambac AAA→AA 강등(2008.6.5, InvestmentNews); FSA HELOC 8건 손실 2.661억(Bond Buyer 2008.5.15); "
           "'Pre-crisis, you couldn't get a AAA on a Heloc-backed bond without a bond wrap'(Bloomberg 2019.6.26)",
    "'13": "KBRA(2026.9.8), HousingWire(2023.10.5), dv01·Fitch(HEL News 2025.7.15) 'no securitization volume from 2008 to 2021'(CES)",
    "'22": "Freddie Mac PMMS 30년 고정(FRED MORTGAGE30US) 3.22%(2022.1.6)→7.08%(2022.10.27); dv01·Fitch(HEL News 2025.7.15) 'Activity resumed in 2021 as elevated interest rates significantly reduced refinance incentives'",
    "'19": "Towd Point HE Trust 2019-HE1(FirstKey·Cerberus) — Fitch 'first all-HELOC securitization in the post-crisis era'(NMN 2019.6.14)",
    "'20": "FLOC 2020-1, 블록체인 기반 첫 증권화(The Block 2020.3.11, DBRS 2023.4.14)",
    "'24": "Figure 424B4(2025.9.11) 'In June 2024, we launched Figure Connect'",
}

MKT = [MKT[y] for y in YEARS]
FIG = [FIG[y] for y in YEARS]

NOTE = ("연도별 집계 기관 상이(04·05 IMF, 22 S&P, 23·24 HEL News, 25 J.P. Morgan, 26E BofA 전망), 04년은 05년 증가율로 역산, 06·07년은 원자료 확인 중(n/a)"
        " / 08~18년은 '사실상 0' 서술 반영, 19~21년은 확인된 딜 합계(하한), 동사는 DBRS 딜별 UPB 합계(26E 동사는 26.11.19까지 확인된 19건)")

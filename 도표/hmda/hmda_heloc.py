# hmda_heloc.py — HMDA 공개 LAR(FFIEC Data Browser)에서 기관(LEI)별·연도별 HELOC 실행액을 집계한다.
#
#   python3 -I hmda_heloc.py find "Guaranteed Rate" "CrossCountry"        # 이름 → LEI (filers 목록)
#   python3 -I hmda_heloc.py market 2018 2025 ../hmda_raw [keep_leis.txt]  # 연도별 전국 실행분 → LEI별 집계
#
# 원자료: 연도마다 전국 '실행(action_taken=1)' LAR 전체를 한 번 받는다
#   https://ffiec.cfpb.gov/v2/data-browser-api/view/nationwide/csv?years=<연도>&actions_taken=1
#   (API가 files.ffiec.cfpb.gov 스냅샷 파일로 리다이렉트)
# HELOC 정의: open-end_line_of_credit=1(한도형 여신) & business_or_commercial_purpose≠1 & reverse_mortgage≠1
# 실행액 = loan_amount 합(한도형은 약정 한도). 공개 LAR의 loan_amount는 1만 달러 구간 중간값으로 반올림돼 있다.
# 출력: <outdir>/heloc_by_lei_<연도>.csv  (LEI별 건수·실행액·1순위/후순위)
#       <outdir>/heloc_rows_<연도>.csv    (keep_leis.txt에 적은 LEI의 HELOC 원행 — 파트너 검증용)
#       <outdir>/filers_<연도>.json       (LEI ↔ 기관명)
import csv, gzip, io, json, os, subprocess, sys, tempfile, time

API = "https://ffiec.cfpb.gov/v2/data-browser-api/view"
OEL = ("open-end_line_of_credit", "open_end_line_of_credit")


def get(url, tries=4):
    """curl로 받아 임시 파일 핸들을 돌려준다(네트워크 오류만 백오프 재시도). 쓰고 나면 close+unlink."""
    for k in range(tries):
        tmp = tempfile.NamedTemporaryFile(prefix="hmda_", delete=False)
        tmp.close()
        r = subprocess.run(["curl", "-sS", "-f", "-L", "-m", "3600", "-o", tmp.name, url],
                           capture_output=True, text=True)
        if r.returncode == 0:
            return open(tmp.name, "rb")
        os.unlink(tmp.name)
        if k == tries - 1 or "403" in r.stderr:
            raise RuntimeError(f"curl 실패: {r.stderr.strip()}")
        print(f"  retry {k + 1}: {r.stderr.strip()}", file=sys.stderr)
        time.sleep(4 * 2 ** k)


def done(fh):
    fh.close()
    os.unlink(fh.name)


def filers(year):
    fh = get(f"{API}/filers?years={year}")
    data = json.load(fh)
    done(fh)
    return data.get("institutions", [])


def find(keys, years=range(2018, 2026)):
    hits = {}
    for y in years:
        for inst in filers(y):
            if any(k.lower() in inst["name"].lower() for k in keys):
                h = hits.setdefault(inst["lei"], {"names": set(), "years": {}})
                h["names"].add(inst["name"])
                h["years"][y] = inst.get("count")
    for lei, h in hits.items():
        print(lei, " / ".join(sorted(h["names"])), dict(sorted(h["years"].items())), sep="\t")


def market(y0, y1, outdir, keep=None):
    os.makedirs(outdir, exist_ok=True)
    keep = {l.strip() for l in open(keep, encoding="utf-8") if l.strip()} if keep else set()
    for y in range(int(y0), int(y1) + 1):
        out = os.path.join(outdir, f"heloc_by_lei_{y}.csv")
        if os.path.exists(out):
            print(f"{y}: 있음, 건너뜀"); continue
        t = time.time()
        names = {i["lei"]: i["name"] for i in filers(y)}
        json.dump(names, open(os.path.join(outdir, f"filers_{y}.json"), "w", encoding="utf-8"), ensure_ascii=False)
        url = f"{API}/nationwide/csv?years={y}&actions_taken=1"
        r = get(url)
        gz = r.read(2) == b"\x1f\x8b"
        r.seek(0)
        rd = csv.DictReader(io.TextIOWrapper(gzip.open(r) if gz else r, encoding="utf-8", newline=""))
        oel = next((c for c in OEL if c in (rd.fieldnames or [])), None)
        if not oel or "loan_amount" not in rd.fieldnames:
            raise RuntimeError(f"{y}: LAR CSV가 아님 — 머리행 {str(rd.fieldnames)[:200]} (임시 파일 {r.name})")
        agg, tot = {}, {"rows": 0, "yr_bad": 0}
        kf = open(os.path.join(outdir, f"heloc_rows_{y}.csv"), "w", newline="", encoding="utf-8") if keep else None
        kw = csv.DictWriter(kf, fieldnames=rd.fieldnames) if kf else None
        if kw:
            kw.writeheader()
        for row in rd:
            tot["rows"] += 1
            if row.get("activity_year") != str(y) or row.get("action_taken") != "1":
                tot["yr_bad"] += 1; continue
            if row.get(oel) != "1" or row.get("business_or_commercial_purpose") == "1" or row.get("reverse_mortgage") == "1":
                continue
            try:
                amt = int(float(row["loan_amount"]))
            except ValueError:
                continue
            a = agg.setdefault(row["lei"], [0, 0, 0, 0, 0, 0])
            a[0] += 1; a[1] += amt
            if row.get("lien_status") == "1":
                a[2] += 1; a[3] += amt
            elif row.get("lien_status") == "2":
                a[4] += 1; a[5] += amt
            if kw and row["lei"] in keep:
                kw.writerow(row)
        done(r)
        if kf:
            kf.close()
        if tot["yr_bad"]:
            print(f"  ! {y}: 연도·실행 조건이 다른 행 {tot['yr_bad']:,}건 제외")
        with open(out + ".tmp", "w", newline="", encoding="utf-8") as f:
            w = csv.writer(f)
            w.writerow(["연도", "LEI", "기관명", "HELOC건수", "HELOC실행액", "1순위건수", "1순위실행액",
                        "후순위건수", "후순위실행액", "API"])
            for lei, a in sorted(agg.items(), key=lambda kv: -kv[1][1]):
                w.writerow([y, lei, names.get(lei, "")] + a + [url])
        os.replace(out + ".tmp", out)
        n = sum(a[0] for a in agg.values()); s = sum(a[1] for a in agg.values())
        print(f"{y}: 실행 {tot['rows']:,}건 중 HELOC {n:,}건 ${s:,} ({len(agg):,}개 기관, {time.time() - t:.0f}s)", flush=True)


if __name__ == "__main__":
    cmd, *args = sys.argv[1:]
    find(args) if cmd == "find" else market(*args)

#!/usr/bin/env python3
"""スキルシート生成スクリプト

    python build.py                     # data.yaml → output/ に .docx を出力
    python build.py --pdf               # PDF も出力（LibreOffice が必要）
    python build.py --as-of 2026-10     # 「現在」の基準月・作成年月を指定
    python build.py --tags 組込み,通信   # タグが一致するプロジェクトだけ出力
    python build.py -d other.yaml -o out_dir

必要ライブラリ: pip install docxtpl pyyaml
"""
from __future__ import annotations

import argparse
import datetime as dt
import shutil
import subprocess
import sys
from pathlib import Path

import yaml
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docxtpl import DocxTemplate

HERE = Path(__file__).resolve().parent


# ---------- 日付ユーティリティ ----------
def parse_ym(v, field: str) -> tuple[int, int]:
    """'2019-04' / 2019-04 (YAMLでdate化されたもの) / '2019/4' を (年, 月) に"""
    if isinstance(v, (dt.date, dt.datetime)):
        return v.year, v.month
    s = str(v).strip().replace("/", "-")
    try:
        y, m = s.split("-")[:2]
        y, m = int(y), int(m)
        assert 1 <= m <= 12
        return y, m
    except Exception:
        raise SystemExit(f"[エラー] {field} の日付形式が不正です: {v!r}（YYYY-MM で書いてください）")


def jp(ym: tuple[int, int]) -> str:
    return f"{ym[0]}年{ym[1]}月"


def months(f, t, mode: str) -> int:
    n = (t[0] - f[0]) * 12 + (t[1] - f[1])
    return n + 1 if mode == "inclusive" else n


def fmt_duration(n: int) -> str:
    y, m = divmod(max(n, 0), 12)
    s = (f"{y}年" if y else "") + (f"{m}ヶ月" if m else "")
    return s or "1ヶ月未満"


# ---------- コンテキスト組み立て ----------
def build_context(data: dict, as_of: tuple[int, int], tags: set[str] | None) -> dict:
    mode = data.get("duration_mode", "inclusive")
    if mode not in ("inclusive", "exclusive"):
        raise SystemExit("[エラー] duration_mode は inclusive か exclusive を指定してください")

    for key in ("engineer_code", "careers", "projects", "skills"):
        if key not in data:
            raise SystemExit(f"[エラー] data.yaml に必須項目 {key} がありません")

    ctx = dict(data)
    ctx["issued"] = jp(as_of)

    # 職務経歴
    careers = []
    for i, c in enumerate(data["careers"]):
        f = parse_ym(c["from"], f"careers[{i}].from")
        to = c.get("to")
        period = f"{jp(f)}～" + (jp(parse_ym(to, f"careers[{i}].to")) if to else "")
        careers.append({"period": period, "org": c["org"], "pad": not to})
    ctx["careers"] = careers

    # プロジェクト
    projects = []
    for i, p in enumerate(data["projects"]):
        if tags and not (tags & set(p.get("tags", []))):
            continue
        f = parse_ym(p["from"], f"projects[{i}].from")
        to = p.get("to")
        t = parse_ym(to, f"projects[{i}].to") if to else as_of
        if t < f:
            raise SystemExit(f"[エラー] projects[{i}] ({p.get('name')}) の終了月が開始月より前です")
        dur = p.get("duration") or fmt_duration(months(f, t, mode))
        end = jp(t) if to else "現在"
        env = [{"label": k, "values_": v} for k, v in (p.get("env") or {}).items() if v]
        projects.append({
            **p,
            "_from": f,
            "period_label": f"{jp(f)}～{end}({dur})",
            "overview": p.get("overview") or [],
            "work": p.get("work") or [],
            "achievements": p.get("achievements") or [],
            "phases": p.get("phases") or [],
            "env": env,
        })
    if not projects:
        raise SystemExit("[エラー] 出力対象のプロジェクトが0件です（--tags を確認してください）")
    ctx["projects"] = projects
    ctx["dev_since"] = jp(min(p["_from"] for p in projects))

    # スキル
    ctx["skills"] = [
        {"cat": cat, "rows": [{"name": s["name"],
                                "work": str(s.get("work", "")),
                                "private": str(s.get("private", ""))} for s in items]}
        for cat, items in data["skills"].items() if items
    ]

    edu = data.get("education") or {}
    ctx["education"] = {"date": jp(parse_ym(edu["date"], "education.date")) if edu.get("date") else "",
                        "text": edu.get("text", "")}
    ctx["licenses"] = [{"date": jp(parse_ym(l["date"], "licenses.date")), "name": l["name"]}
                       for l in data.get("licenses") or []]
    ctx["summary"] = data.get("summary") or []
    ctx["self_pr"] = data.get("self_pr") or []
    ctx["private_dev"] = data.get("private_dev") or []
    return ctx


# ---------- 差し込み後の後処理 ----------
def _set_vmerge(tc, val: str | None) -> None:
    """tcPr 内の正しい位置(tcW/gridSpan の直後)に w:vMerge を入れる"""
    tcPr = tc.get_or_add_tcPr()
    for old in tcPr.findall(qn("w:vMerge")):
        tcPr.remove(old)
    if val is None:
        return
    vm = OxmlElement("w:vMerge")
    vm.set(qn("w:val"), val)
    anchor = None
    for tag in ("w:tcW", "w:gridSpan", "w:hMerge", "w:cnfStyle"):
        anchor = tcPr.find(qn(tag))
        if anchor is not None:
            break
    if anchor is not None:
        anchor.addnext(vm)
    else:
        tcPr.insert(0, vm)


def _set_align(tc, val: str) -> None:
    for p in tc.findall(qn("w:p")):
        pPr = p.get_or_add_pPr()
        jc = pPr.find(qn("w:jc"))
        if jc is None:
            jc = OxmlElement("w:jc")
            # jc は rPr / sectPr / pPrChange より前に置く
            tail = next((c for c in pPr if c.tag in (qn("w:textDirection"), qn("w:textAlignment"),
                         qn("w:outlineLvl"), qn("w:rPr"), qn("w:sectPr"), qn("w:pPrChange"))), None)
            if tail is not None:
                tail.addprevious(jc)
            else:
                pPr.append(jc)
        jc.set(qn("w:val"), val)


def postprocess(document) -> None:
    """スキル表: 同じ区分を縦結合し、'-' や空欄を中央揃えにする"""
    for tbl in document.tables:
        rows = tbl._tbl.findall(qn("w:tr"))
        if not rows:
            continue
        head = [ "".join(t.text or "" for t in tc.iter(qn("w:t"))) for tc in rows[0].findall(qn("w:tc"))]
        if len(head) != 4 or head[1] != "環境":
            continue
        prev = None
        for tr in rows[1:]:
            tcs = tr.findall(qn("w:tc"))
            cat = "".join(t.text or "" for t in tcs[0].iter(qn("w:t")))
            if cat == prev:
                _set_vmerge(tcs[0], "continue")
                for t in tcs[0].iter(qn("w:t")):
                    t.text = ""
            else:
                _set_vmerge(tcs[0], "restart")
            prev = cat
            for tc in tcs[2:]:
                txt = "".join(t.text or "" for t in tc.iter(qn("w:t"))).strip()
                if txt in ("-", ""):
                    _set_align(tc, "center")


def to_pdf(docx_path: Path) -> Path | None:
    soffice = shutil.which("soffice") or shutil.which("libreoffice")
    if soffice:
        subprocess.run([soffice, "--headless", "--convert-to", "pdf", "--outdir",
                        str(docx_path.parent), str(docx_path)],
                       check=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    else:
        try:  # Windows + Word がある場合
            from docx2pdf import convert  # pip install docx2pdf
            convert(str(docx_path))
        except ImportError:
            print("[警告] PDF 変換には LibreOffice か docx2pdf（Word）が必要です", file=sys.stderr)
            return None
    pdf = docx_path.with_suffix(".pdf")
    return pdf if pdf.exists() and pdf.stat().st_size > 0 else None


def main() -> None:
    ap = argparse.ArgumentParser(description="data.yaml からスキルシートを生成します")
    ap.add_argument("-d", "--data", default=HERE / "data.yaml", type=Path)
    ap.add_argument("-t", "--template", default=HERE / "template.docx", type=Path)
    ap.add_argument("-o", "--outdir", default=HERE / "output", type=Path)
    ap.add_argument("--as-of", help="基準月 YYYY-MM（省略時は今月）")
    ap.add_argument("--tags", help="カンマ区切り。一致するタグを持つプロジェクトだけ出力")
    ap.add_argument("--pdf", action="store_true", help="PDF も出力する")
    a = ap.parse_args()

    data = yaml.safe_load(a.data.read_text(encoding="utf-8"))
    today = dt.date.today()
    as_of = parse_ym(a.as_of, "--as-of") if a.as_of else (today.year, today.month)
    tags = {t.strip() for t in a.tags.split(",")} if a.tags else None

    ctx = build_context(data, as_of, tags)
    tpl = DocxTemplate(a.template)
    tpl.render(ctx, autoescape=True)  # & < > を含む文字列も安全に出力
    postprocess(tpl.docx)

    a.outdir.mkdir(parents=True, exist_ok=True)
    out = a.outdir / f"【{ctx['engineer_code']}】スキルシート_{as_of[0]}{as_of[1]:02d}.docx"
    tpl.save(out)
    print(f"出力: {out}")
    if a.pdf:
        pdf = to_pdf(out)
        if pdf:
            print(f"出力: {pdf}")


if __name__ == "__main__":
    main()

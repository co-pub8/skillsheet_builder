#!/usr/bin/env python3
"""スキルシート生成スクリプト

    python build.py                     # data.yaml → output/ に .docx を出力
    python build.py --pdf               # PDF も出力（LibreOffice、なければ docx2pdf を使用）
    python build.py --as-of 2026-10     # 「現在」の基準月・作成年月を指定
    python build.py --tags 組込み,通信   # タグが一致するプロジェクトだけ出力
    python build.py -d other.yaml -o out_dir

必要ライブラリ: pip install -r requirements.txt
"""
import argparse
import datetime as dt
import shutil
import subprocess
import sys
from pathlib import Path

import yaml
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docxtpl import DocxTemplate

HERE = Path(__file__).resolve().parent

YM = tuple[int, int]  # (年, 月)


# ---------- 日付ユーティリティ ----------
def parse_ym(v, field: str) -> YM:
    """'2019-04' / 2019-04 (YAMLでdate化されたもの) / '2019/4' を (年, 月) に"""
    if isinstance(v, dt.date):  # datetime も date のサブクラス
        return v.year, v.month
    try:
        y, m = (int(x) for x in str(v).strip().replace("/", "-").split("-")[:2])
        if 1 <= m <= 12:
            return y, m
    except ValueError:
        pass
    raise SystemExit(f"[エラー] {field} の日付形式が不正です: {v!r}（YYYY-MM で書いてください）")


def fmt_ym(ym: YM) -> str:
    return f"{ym[0]}年{ym[1]}月"


def count_months(start: YM, end: YM, mode: str) -> int:
    """期間の月数。inclusive は開始月・終了月を両方数える（4月～4月で1ヶ月）"""
    n = (end[0] - start[0]) * 12 + (end[1] - start[1])
    return n + 1 if mode == "inclusive" else n


def fmt_duration(n: int) -> str:
    y, m = divmod(max(n, 0), 12)
    s = (f"{y}年" if y else "") + (f"{m}ヶ月" if m else "")
    return s or "1ヶ月未満"


def as_text(v) -> str:
    """YAML で数値になった値も文字列に。未記入（null）は空欄"""
    return "" if v is None else str(v)


# ---------- コンテキスト組み立て ----------
def build_context(data: dict, as_of: YM, tags: set[str] | None) -> dict:
    mode = data.get("duration_mode", "inclusive")
    if mode not in ("inclusive", "exclusive"):
        raise SystemExit("[エラー] duration_mode は inclusive か exclusive を指定してください")

    for key in ("engineer_code", "careers", "projects", "skills"):
        if key not in data:
            raise SystemExit(f"[エラー] data.yaml に必須項目 {key} がありません")

    ctx = dict(data)
    ctx["issued"] = fmt_ym(as_of)
    # 未記入のリスト項目は空リストにそろえる（テンプレートの for でエラーにしないため）
    for key in ("summary", "self_pr", "private_dev"):
        ctx[key] = data.get(key) or []

    # 職務経歴
    careers = []
    for i, c in enumerate(data["careers"]):
        start = parse_ym(c["from"], f"careers[{i}].from")
        to = c.get("to")
        end = fmt_ym(parse_ym(to, f"careers[{i}].to")) if to else ""
        # pad: 在籍中（to なし）は期間の文字列が短いので、テンプレート側でタブを1つ足して所属名の位置を揃える
        careers.append({"period": f"{fmt_ym(start)}～{end}", "org": c["org"], "pad": not to})
    ctx["careers"] = careers

    # プロジェクト
    projects, starts = [], []
    for i, p in enumerate(data["projects"]):
        if tags and not (tags & set(p.get("tags", []))):
            continue
        start = parse_ym(p["from"], f"projects[{i}].from")
        to = p.get("to")
        end = parse_ym(to, f"projects[{i}].to") if to else as_of
        if end < start:
            raise SystemExit(f"[エラー] projects[{i}] ({p.get('name')}) の終了月が開始月より前です")
        # duration を書いた場合は自動計算より優先する
        duration = p.get("duration") or fmt_duration(count_months(start, end, mode))
        end_label = fmt_ym(end) if to else "現在"
        item = {
            **p,
            "period_label": f"{fmt_ym(start)}～{end_label}({duration})",
            # Jinja では sec.values が dict.values() と衝突するため values_ にする
            "env": [{"label": k, "values_": v} for k, v in (p.get("env") or {}).items() if v],
        }
        for key in ("overview", "work", "achievements", "phases"):
            item[key] = p.get(key) or []
        projects.append(item)
        starts.append(start)
    if not projects:
        raise SystemExit("[エラー] 出力対象のプロジェクトが0件です（--tags を確認してください）")
    ctx["projects"] = projects
    ctx["dev_since"] = fmt_ym(min(starts))

    # スキル（区分ごと）
    skills = []
    for cat, items in data["skills"].items():
        if not items:
            continue
        rows = [{"name": s["name"], "work": as_text(s.get("work")), "private": as_text(s.get("private"))}
                for s in items]
        skills.append({"cat": cat, "rows": rows})
    ctx["skills"] = skills

    # 学歴・資格
    edu = data.get("education") or {}
    ctx["education"] = {
        "date": fmt_ym(parse_ym(edu["date"], "education.date")) if edu.get("date") else "",
        "text": edu.get("text", ""),
    }
    ctx["licenses"] = [
        {"date": fmt_ym(parse_ym(lic["date"], f"licenses[{i}].date")), "name": lic["name"]}
        for i, lic in enumerate(data.get("licenses") or [])
    ]
    return ctx


# ---------- 差し込み後の後処理 ----------
# 縦結合や中央揃えは XML 属性で表すが、テンプレートの属性に Jinja タグは書けない（Word で開けなくなる）ため、
# 差し込み後にここで設定する。vMerge / jc の挿入位置（スキーマ順）は python-docx が調整する。
def _cell_text(tc) -> str:
    return "".join(t.text or "" for t in tc.iter(qn("w:t")))


def _is_skill_table(rows) -> bool:
    """ヘッダが4列で2列目が「環境」の表をスキル表とみなす（テンプレートの見出しを変えたらここも直す）"""
    head = [_cell_text(tc) for tc in rows[0].findall(qn("w:tc"))]
    return len(head) == 4 and head[1] == "環境"


def postprocess(document) -> None:
    """スキル表: 同じ区分を縦結合し、経験欄の '-' や空欄を中央揃えにする"""
    for tbl in document.tables:
        rows = tbl._tbl.findall(qn("w:tr"))
        if not rows or not _is_skill_table(rows):
            continue
        prev = None
        for tr in rows[1:]:
            cat_tc, _, *exp_tcs = tr.findall(qn("w:tc"))  # 区分, 環境, 業務経験, プライベート
            cat = _cell_text(cat_tc)
            if cat == prev:
                cat_tc.vMerge = "continue"
                # 結合されて見えなくなる側の文字は消しておく（Word 以外で残って見えることがあるため）
                for t in cat_tc.iter(qn("w:t")):
                    t.text = ""
            else:
                cat_tc.vMerge = "restart"
            prev = cat
            for tc in exp_tcs:
                if _cell_text(tc).strip() in ("-", ""):
                    for p in tc.findall(qn("w:p")):
                        p.get_or_add_pPr().jc_val = WD_ALIGN_PARAGRAPH.CENTER


def to_pdf(docx_path: Path) -> Path | None:
    """同じフォルダに PDF を出力する。LibreOffice があれば使い、なければ docx2pdf（Word）で変換する"""
    soffice = shutil.which("soffice") or shutil.which("libreoffice")
    if soffice:
        subprocess.run([soffice, "--headless", "--convert-to", "pdf", "--outdir",
                        str(docx_path.parent), str(docx_path)],
                       check=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    else:
        try:
            from docx2pdf import convert
        except ImportError:
            print("[警告] PDF 変換には LibreOffice か docx2pdf（Word）が必要です", file=sys.stderr)
            return None
        convert(str(docx_path))
    pdf = docx_path.with_suffix(".pdf")
    if pdf.exists() and pdf.stat().st_size > 0:
        return pdf
    print("[警告] PDF の出力に失敗しました", file=sys.stderr)
    return None


def main() -> None:
    ap = argparse.ArgumentParser(description="data.yaml からスキルシートを生成します")
    ap.add_argument("-d", "--data", default=HERE / "data.yaml", type=Path)
    ap.add_argument("-t", "--template", default=HERE / "template.docx", type=Path)
    ap.add_argument("-o", "--outdir", default=HERE / "output", type=Path)
    ap.add_argument("--as-of", help="基準月 YYYY-MM（省略時は今月）")
    ap.add_argument("--tags", help="カンマ区切り。一致するタグを持つプロジェクトだけ出力")
    ap.add_argument("--pdf", action="store_true", help="PDF も出力する")
    args = ap.parse_args()

    data = yaml.safe_load(args.data.read_text(encoding="utf-8"))
    today = dt.date.today()
    as_of = parse_ym(args.as_of, "--as-of") if args.as_of else (today.year, today.month)
    tags = {t.strip() for t in args.tags.split(",")} if args.tags else None

    ctx = build_context(data, as_of, tags)
    tpl = DocxTemplate(args.template)
    tpl.render(ctx, autoescape=True)  # & < > を含む文字列も安全に出力
    postprocess(tpl.docx)

    args.outdir.mkdir(parents=True, exist_ok=True)
    out = args.outdir / f"【{ctx['engineer_code']}】スキルシート_{as_of[0]}{as_of[1]:02d}.docx"
    tpl.save(out)
    print(f"出力: {out}")
    if args.pdf:
        pdf = to_pdf(out)
        if pdf:
            print(f"出力: {pdf}")


if __name__ == "__main__":
    main()

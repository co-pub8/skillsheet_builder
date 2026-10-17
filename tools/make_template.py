"""元のスキルシート(.docx)から docxtpl 用 template.docx を生成する一回限りのスクリプト。
使い方: python tools/make_template.py [source/original.docx [template.docx]]
(通常は生成済みの template.docx を使えばよく、これを再実行する必要はありません)
"""
import argparse
import copy
import io
import re
import zipfile

import docx
from docx.oxml.ns import qn

W_P, W_R, W_T, W_TAB, W_RPR = qn("w:p"), qn("w:r"), qn("w:t"), qn("w:tab"), qn("w:rPr")
W_TR, W_TC = qn("w:tr"), qn("w:tc")
XML_SPACE = "{http://www.w3.org/XML/1998/namespace}space"


def run_with(proto_run, text=None, tab=False):
    """proto_run の書式(rPr)だけを引き継ぎ、text またはタブ1つを持つランを作る。"""
    r = copy.deepcopy(proto_run)
    for child in list(r):
        if child.tag != W_RPR:
            r.remove(child)
    if tab:
        r.append(r.makeelement(W_TAB, {}))
    else:
        t = r.makeelement(W_T, {})
        t.text = text
        t.set(XML_SPACE, "preserve")
        r.append(t)
    return r


def set_text(p_el, text, keep_run=0):
    """段落の書式(pPr)と keep_run 番目のランの書式を残し、テキストを1ランにまとめる。"""
    runs = p_el.findall(W_R)
    base = runs[keep_run] if runs else p_el.makeelement(W_R, {})
    for r in runs:
        p_el.remove(r)
    p_el.append(run_with(base, text))
    return p_el


def clone_p(proto, text, keep_run=0):
    return set_text(copy.deepcopy(proto), text, keep_run)


def fill_cell(tc, items):
    """セル内の段落を items で作り直す。items は (text, proto_index) のリストで、
    proto_index は元のセル内の何番目の段落の書式を使うか。"""
    paras = tc.findall(W_P)
    protos = [copy.deepcopy(p) for p in paras]
    for p in paras:
        tc.remove(p)
    for text, idx in items:
        tc.append(clone_p(protos[idx], text))


def drop_vmerge(tc):
    tcPr = tc.find(qn("w:tcPr"))
    vm = tcPr.find(qn("w:vMerge")) if tcPr is not None else None
    if vm is not None:
        tcPr.remove(vm)


def tag_row(proto_row, text):
    """proto_row の書式で、先頭セルに text だけを書いた行（{%tr %} タグ用）を作る。"""
    r = copy.deepcopy(proto_row)
    tc, *rest = r.findall(W_TC)
    for t in rest:
        r.remove(t)
    drop_vmerge(tc)
    set_text(tc.find(W_P), text)
    return r


def close_loop_after(el, proto):
    """ループ末尾の el（改ページなど）を最後の回以外だけ出力し、その後ろでループを閉じる。
    結果は {%p if not loop.last %} / el / {%p endif %} / {%p endfor %} の並びになる。"""
    el.addprevious(clone_p(proto, "{%p if not loop.last %}"))
    endif = clone_p(proto, "{%p endif %}")
    el.addnext(endif)
    endif.addnext(clone_p(proto, "{%p endfor %}"))


def normalize_package(src):
    """LibreOffice 経由の .docx に含まれる Word 非互換な箇所を、python-docx で読み込む前に直す。"""
    buf = io.BytesIO()
    with zipfile.ZipFile(src) as zin, zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zout:
        for name in zin.namelist():
            b = zin.read(name)
            if name == "word/fontTable.xml":
                # characterSet 属性は Word のスキーマ外の値(utf-8 等)なので削除
                b = re.sub(rb' w:characterSet="[^"]*"', b"", b)
            elif name == "word/styles.xml":
                # LibreOffice 由来の "ＭＳ 明朝;MS Mincho" という連結フォント名を正規化
                b = b.replace("ＭＳ 明朝;MS Mincho".encode(), "ＭＳ 明朝".encode())
            elif name == "word/document.xml":
                # docGrid の charSpace が Int32 範囲外(4294962994)で Word のスキーマ違反。
                # 文字グリッドは LibreOffice で文字間が不自然に空くため、行グリッドのみにする。
                b = re.sub(rb'<w:docGrid [^>]*?w:linePitch="(\d+)"[^>]*/>',
                           rb'<w:docGrid w:type="lines" w:linePitch="\1"/>', b)
            elif name == "_rels/.rels":
                # core-properties の関係タイプが小文字(officedocument)だと、python-docx が core.xml を二重に書き込む
                b = b.replace(b"openxmlformats.org/officedocument/2006/relationships/metadata/core-properties",
                              b"openxmlformats.org/package/2006/relationships/metadata/core-properties")
            zout.writestr(name, b)
    buf.seek(0)
    return buf


def main(src, dst):
    d = docx.Document(normalize_package(src))
    body = d.element.body
    # 本文の要素をインデックスで参照する（元文書の構造に依存）。
    # P は最初に取ったリストなので、途中で要素を挿入・削除しても P[i] が指す要素は変わらない。
    P = list(body.iterchildren())

    # --- ヘッダ部 ---
    set_text(P[1], "{{ issued }}")
    set_text(P[2], "技術者コード：{{ engineer_code }}")
    set_text(P[3], "(所属：{{ affiliation }})", keep_run=1)

    # --- 職務要約 (行リスト) ---
    set_text(P[5], "{%p for line in summary %}")
    set_text(P[6], "{{ line }}")
    set_text(P[7], "{%p endfor %}")

    # --- 職務経歴 ---
    # 1行を「期間 <タブ> (在籍中のみ追加の <タブ>) 所属」にする
    set_text(P[10], "{%p for c in careers %}")
    p11 = P[11]
    runs = p11.findall(W_R)
    for r in runs:
        p11.remove(r)
    p11.extend([
        run_with(runs[0], "{{ c.period }}"),
        run_with(runs[0], tab=True),
        run_with(runs[0], "{%r if c.pad %}"),
        run_with(runs[0], tab=True),
        run_with(runs[0], "{%r endif %}"),
        run_with(runs[0], "{{ c.org }}"),
    ])
    set_text(P[12], "{%p endfor %}")
    body.remove(P[13])

    # --- 開発経歴 見出し ---
    runs = P[15].findall(W_R)
    for r in runs[2:]:
        P[15].remove(r)
    runs[1].find(W_T).text = "開発経歴({{ dev_since }}以降)"

    # --- プロジェクト(表ごとループ) ---
    set_text(P[16], "{{ p.company }}")
    P[16].addprevious(clone_p(P[16], "{%p for p in projects %}"))
    rows = P[17].findall(W_TR)
    set_text(rows[0].find(W_TC).find(W_P), "{{ p.period_label }}")
    c0, c1, c2 = rows[1].findall(W_TC)
    fill_cell(c0, [
        ("【プロジェクト名】", 0), ("{{ p.name }}", 1), ("", 2),
        ("【プロジェクト概要】", 3), ("{%p for x in p.overview %}", 4), ("{{ x }}", 4), ("{%p endfor %}", 4), ("", 5),
        ("【業務内容】", 6), ("{%p for x in p.work %}", 7), ("{{ x }}", 7), ("{%p endfor %}", 7),
        ("{%p if p.achievements %}", 8), ("", 8),
        ("【実績・取り組み等】", 9), ("{%p for x in p.achievements %}", 10), ("{{ x }}", 10), ("{%p endfor %}", 10),
        ("{%p endif %}", 10),
    ])
    fill_cell(c1, [
        ("{%p for sec in p.env %}", 1), ("【{{ sec.label }}】", 0),
        ("{%p for x in sec.values_ %}", 1), ("{{ x }}", 1), ("{%p endfor %}", 1),
        ("{%p if not loop.last %}", 2), ("", 2), ("{%p endif %}", 2),
        ("{%p endfor %}", 1),
    ])
    fill_cell(c2, [
        ("【役割】", 0), ("{{ p.role }}", 1), ("", 2),
        ("【プロジェクト規模】", 3), ("{{ p.size }}", 4), ("", 5),
        ("【担当】", 6), ("{%p for x in p.phases %}", 7), ("{{ x }}", 7), ("{%p endfor %}", 7),
    ])
    close_loop_after(P[18], P[19])  # P[18] はプロジェクト間の改ページ
    for i in range(19, 24):
        body.remove(P[i])

    # --- スキル表 ---
    # 縦結合(vMerge)はここでは外し、build.py の後処理で同じ区分の行に付け直す
    st = P[26]
    srows = st.findall(W_TR)
    data_row = srows[1]
    for r in srows[2:]:
        st.remove(r)
    tcs = data_row.findall(W_TC)
    drop_vmerge(tcs[0])
    for tc, var in zip(tcs, ["g.cat", "s.name", "s.work", "s.private"]):
        set_text(tc.find(W_P), "{{ %s }}" % var)
    data_row.addprevious(tag_row(data_row, "{%tr for g in skills %}"))
    data_row.addprevious(tag_row(data_row, "{%tr for s in g.rows %}"))
    data_row.addnext(tag_row(data_row, "{%tr endfor %}"))
    data_row.addnext(tag_row(data_row, "{%tr endfor %}"))

    # --- 最終学歴 ---
    set_text(P[29], "　{{ education.date }}　{{ education.text }}", keep_run=4)

    # --- 保有資格 ---
    lrow = P[32].find(W_TR)
    ltcs = lrow.findall(W_TC)
    set_text(ltcs[0].find(W_P), "{{ l.date }}")
    set_text(ltcs[1].find(W_P), "{{ l.name }}")
    lrow.addprevious(tag_row(lrow, "{%tr for l in licenses %}"))
    lrow.addnext(tag_row(lrow, "{%tr endfor %}"))

    # --- 自己PR ---
    set_text(P[34], "{{ pr.title }}")
    P[34].addprevious(clone_p(P[35], "{%p for pr in self_pr %}"))
    set_text(P[35], "{%p for line in pr.lines %}")
    set_text(P[36], "{{ line }}")
    set_text(P[37], "{%p endfor %}")
    close_loop_after(P[38], P[35])  # P[38] は項目間の空行
    for i in range(39, 42):
        body.remove(P[i])

    # --- プライベートでの開発経験 ---
    set_text(P[44], "{%p for x in private_dev %}")
    set_text(P[45], "・{{ x }}")
    set_text(P[46], "{%p endfor %}")
    body.remove(P[47])

    # LibreOffice 由来の jc="start"/"end" を Word 互換の left/right に統一
    for jc in body.iter(qn("w:jc")):
        v = jc.get(qn("w:val"))
        jc.set(qn("w:val"), {"start": "left", "end": "right"}.get(v, v))

    # Word は属性値を厳密に検証するため、属性に Jinja タグが入ると開けないファイルになる。
    # 属性に依存する処理は build.py の後処理で行い、テンプレートは常に有効な XML のままにする。
    for e in body.iter():
        for k, v in e.attrib.items():
            if "{%" in v or "{{" in v:
                raise SystemExit(f"属性にテンプレート構文が残っています: {e.tag} {k}={v}")
    d.save(dst)
    print(f"出力: {dst}")


if __name__ == "__main__":
    ap = argparse.ArgumentParser(description="元のスキルシートから template.docx を生成します")
    ap.add_argument("src", nargs="?", default="source/original.docx")
    ap.add_argument("dst", nargs="?", default="template.docx")
    args = ap.parse_args()
    main(args.src, args.dst)

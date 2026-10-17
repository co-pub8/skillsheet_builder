"""元のスキルシート(.docx)から docxtpl 用 template.docx を生成する一回限りのスクリプト。
使い方: python tools/make_template.py source/original.docx template.docx
(通常は生成済みの template.docx を使えばよく、これを再実行する必要はありません)
"""
import copy
import sys

import docx
from docx.oxml.ns import qn

W_T, W_R, W_P, W_TAB, W_BR = qn("w:t"), qn("w:r"), qn("w:p"), qn("w:tab"), qn("w:br")


def set_text(p_el, text, keep_run=0):
    """段落の書式(pPr)と keep_run 番目のランの書式を残し、テキストを1ランにまとめる。"""
    runs = p_el.findall(W_R)
    base = runs[keep_run] if runs else None
    for r in runs:
        p_el.remove(r)
    r = copy.deepcopy(base) if base is not None else p_el.makeelement(W_R, {})
    for child in list(r):
        if child.tag != qn("w:rPr"):
            r.remove(child)
    t = r.makeelement(W_T, {})
    t.text = text
    t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    r.append(t)
    p_el.append(r)
    return p_el


def clone_p(proto, text, keep_run=0):
    p = copy.deepcopy(proto)
    return set_text(p, text, keep_run)


def run_with(proto_run, text=None, tab=False):
    r = copy.deepcopy(proto_run)
    for child in list(r):
        if child.tag != qn("w:rPr"):
            r.remove(child)
    if tab:
        r.append(r.makeelement(W_TAB, {}))
    else:
        t = r.makeelement(W_T, {})
        t.text = text
        t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
        r.append(t)
    return r


def fill_cell(tc, items):
    """セル内の段落を items で作り直す。items は (text, proto_index) のリスト。"""
    paras = tc.findall(W_P)
    protos = [copy.deepcopy(p) for p in paras]
    for p in paras:
        tc.remove(p)
    for text, idx in items:
        tc.append(clone_p(protos[idx], text))


def section(label, var, protos_idx=0):
    """【見出し】 + リストのループ"""
    return [
        (f"【{label}】", protos_idx),
        (f"{{%p for x in {var} %}}", protos_idx),
        ("{{ x }}", protos_idx),
        ("{%p endfor %}", protos_idx),
    ]


def normalize_package(src):
    """LibreOffice 変換の .docx は core-properties の関係タイプが小文字(officedocument)で、
    python-docx が core.xml を二重に書き込むため、読み込み前に正規化する。"""
    import io, re, zipfile
    zin = zipfile.ZipFile(src)
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zout:
        for name in zin.namelist():
            b = zin.read(name)
            if name == "word/fontTable.xml":
                # characterSet 属性は Word のスキーマ外の値(utf-8 等)なので削除
                b = re.sub(rb' w:characterSet="[^"]*"', b"", b)
            if name == "word/styles.xml":
                # LibreOffice 由来の "ＭＳ 明朝;MS Mincho" という連結フォント名を正規化
                b = b.replace("ＭＳ 明朝;MS Mincho".encode(), "ＭＳ 明朝".encode())
            if name == "word/document.xml":
                # docGrid の charSpace が Int32 範囲外(4294962994)で Word のスキーマ違反。
                # 文字グリッドは LibreOffice で文字間が不自然に空くため、行グリッドのみにする。
                b = re.sub(rb'<w:docGrid [^>]*?w:linePitch="(\d+)"[^>]*/>',
                           rb'<w:docGrid w:type="lines" w:linePitch="\1"/>', b)
            if name == "_rels/.rels":
                b = b.replace(b"openxmlformats.org/officedocument/2006/relationships/metadata/core-properties",
                              b"openxmlformats.org/package/2006/relationships/metadata/core-properties")
            zout.writestr(name, b)
    buf.seek(0)
    return buf


def main(src, dst):
    d = docx.Document(normalize_package(src))
    body = d.element.body
    el = list(body.iterchildren())
    P = el  # インデックスで参照(元文書の構造に依存)

    # --- ヘッダ部 ---
    set_text(P[1], "{{ issued }}")
    set_text(P[2], "技術者コード：{{ engineer_code }}")
    set_text(P[3], "(所属：{{ affiliation }})", keep_run=1)

    # --- 職務要約 (行リスト) ---
    set_text(P[5], "{%p for line in summary %}")
    set_text(P[6], "{{ line }}")
    set_text(P[7], "{%p endfor %}")

    # --- 職務経歴 ---
    set_text(P[10], "{%p for c in careers %}")
    p11 = P[11]
    proto_run = p11.findall(W_R)[0]
    for r in p11.findall(W_R):
        p11.remove(r)
    for r in [
        run_with(proto_run, "{{ c.period }}"),
        run_with(proto_run, tab=True),
        run_with(proto_run, "{%r if c.pad %}"),
        run_with(proto_run, tab=True),
        run_with(proto_run, "{%r endif %}"),
        run_with(proto_run, "{{ c.org }}"),
    ]:
        p11.append(r)
    set_text(P[12], "{%p endfor %}")
    body.remove(P[13])

    # --- 開発経歴 見出し ---
    p15 = P[15]
    runs = p15.findall(W_R)
    for r in runs[2:]:
        p15.remove(r)
    runs[1].find(W_T).text = "開発経歴({{ dev_since }}以降)"

    # --- プロジェクト(表ごとループ) ---
    set_text(P[16], "{{ p.company }}")
    P[16].addprevious(clone_p(P[16], "{%p for p in projects %}"))
    tbl = P[17]
    rows = tbl.findall(qn("w:tr"))
    hdr_cells = rows[0].findall(qn("w:tc"))
    set_text(hdr_cells[0].find(W_P), "{{ p.period_label }}")
    c0, c1, c2 = rows[1].findall(qn("w:tc"))
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
    # 改ページ(最後のプロジェクト以外)
    pb = P[18]
    pb.addprevious(clone_p(P[19], "{%p if not loop.last %}"))
    pb.addnext(clone_p(P[19], "{%p endfor %}"))
    pb.addnext(clone_p(P[19], "{%p endif %}"))
    for i in range(19, 24):
        body.remove(P[i])

    # --- スキル表(カテゴリごとに縦結合) ---
    st = P[26]
    srows = st.findall(qn("w:tr"))
    data_row = srows[1]
    for r in srows[2:]:
        st.remove(r)

    def tag_row(text):
        r = copy.deepcopy(data_row)
        tcs = r.findall(qn("w:tc"))
        for tc in tcs[1:]:
            r.remove(tc)
        tc = tcs[0]
        vm = tc.find(qn("w:tcPr")).find(qn("w:vMerge"))
        if vm is not None:
            vm.getparent().remove(vm)
        set_text(tc.find(W_P), text)
        return r

    # 注意: Word は w:vMerge / w:jc などの属性値を厳密に検証するため、
    # 属性に Jinja タグを入れると「開けないファイル」になる。
    # テンプレートは常に有効な XML のままにし、縦結合と中央揃えは build.py の後処理で行う。
    tcs = data_row.findall(qn("w:tc"))
    vm = tcs[0].find(qn("w:tcPr")).find(qn("w:vMerge"))
    if vm is not None:
        vm.getparent().remove(vm)
    set_text(tcs[0].find(W_P), "{{ g.cat }}")
    for tc, var in zip(tcs[1:], ["s.name", "s.work", "s.private"]):
        set_text(tc.find(W_P), "{{ %s }}" % var)
    data_row.addprevious(tag_row("{%tr for g in skills %}"))
    data_row.addprevious(tag_row("{%tr for s in g.rows %}"))
    data_row.addnext(tag_row("{%tr endfor %}"))
    data_row.addnext(tag_row("{%tr endfor %}"))

    # --- 最終学歴 ---
    set_text(P[29], "\u3000{{ education.date }}\u3000{{ education.text }}", keep_run=4)

    # --- 保有資格 ---
    lt = P[32]
    lrow = lt.findall(qn("w:tr"))[0]
    ltcs = lrow.findall(qn("w:tc"))
    set_text(ltcs[0].find(W_P), "{{ l.date }}")
    set_text(ltcs[1].find(W_P), "{{ l.name }}")
    for text in ["{%tr for l in licenses %}"]:
        r = copy.deepcopy(lrow); [r.remove(t) for t in r.findall(qn("w:tc"))[1:]]
        set_text(r.find(qn("w:tc")).find(W_P), text); lrow.addprevious(r)
    r = copy.deepcopy(lrow); [r.remove(t) for t in r.findall(qn("w:tc"))[1:]]
    set_text(r.find(qn("w:tc")).find(W_P), "{%tr endfor %}"); lrow.addnext(r)

    # --- 自己PR ---
    set_text(P[34], "{{ pr.title }}")
    P[34].addprevious(clone_p(P[35], "{%p for pr in self_pr %}"))
    set_text(P[35], "{%p for line in pr.lines %}")
    set_text(P[36], "{{ line }}")
    set_text(P[37], "{%p endfor %}")
    P[38].addprevious(clone_p(P[35], "{%p if not loop.last %}"))
    P[38].addnext(clone_p(P[35], "{%p endfor %}"))
    P[38].addnext(clone_p(P[35], "{%p endif %}"))
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

    # 属性値にテンプレート構文が残っていないことを確認
    for e in body.iter():
        for k, v in e.attrib.items():
            if "{%" in v or "{{" in v:
                raise SystemExit(f"属性にテンプレート構文が残っています: {e.tag} {k}={v}")
    d.save(dst)
    print("saved", dst)


if __name__ == "__main__":
    main(*(sys.argv[1:3] if len(sys.argv) >= 3 else ["source/original.docx", "template.docx"]))

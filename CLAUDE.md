# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## 概要

`data.yaml`（内容）と `template.docx`（レイアウト）からスキルシート（.docx / .pdf）を生成するツール。内容の変更は `data.yaml` のみ、見た目の変更は `template.docx` を Word で直接編集する、という分離が前提。

## コマンド

```
pip install -r requirements.txt          # docxtpl, PyYAML（PDF を Word で出すなら docx2pdf も）
python build.py                          # output/【技術者コード】スキルシート_YYYYMM.docx を出力
python build.py --pdf                    # PDF も出力（soffice があれば LibreOffice、なければ docx2pdf）
python build.py --as-of 2026-10          # 作成年月・「現在」の基準月を固定（出力を再現したいときに使う）
python build.py --tags 通信,組込み       # タグが一致するプロジェクトのみ
python build.py -d other.yaml -t tpl.docx -o out_dir
```

テスト・リンターは未設定。動作確認は `python build.py -d data.sample.yaml --as-of 2026-10` のように生成し、出力 .docx を Word で開けるか確認する。

`data.yaml` は本人の実際の経歴（個人情報）なので、例示・動作確認・共有には架空データの `data.sample.yaml` を使う。データ構造を変えたときは両方を更新する。

## 処理の流れ（build.py）

1. `build_context()`: YAML を Jinja コンテキストに変換。日付（`YYYY-MM` / YAML の date 型 / `YYYY/M`）を `parse_ym()` で正規化し、期間ラベル（`duration_mode` の inclusive/exclusive、`duration` による個別上書き）、`dev_since`（出力対象プロジェクトの最古開始月）、`--tags` の絞り込みなどをここで計算する。テンプレート側では計算しない。
   - `env` の dict は `[{label, values_}]` に変換される（`values` は Jinja で dict メソッドと衝突するため `values_`）。
   - スキルの `work` / `private` は `str()` 化される（YAML で数値になるため）。
2. `DocxTemplate.render(ctx, autoescape=True)`。
3. `postprocess()`: スキル表の同一区分の縦結合（`w:vMerge`）と、「-」・空欄セルの中央揃え（`w:jc`）を差し込み後に XML 操作で設定する。スキル表は「4列かつヘッダ2列目が `環境`」で識別しているため、テンプレートのヘッダ文言を変えると後処理が効かなくなる。

`data.yaml` の項目を追加・変更する場合は、`build_context()` とテンプレートのタグの両方を合わせる必要がある。

## テンプレート編集の制約（重要）

- Jinja タグを書いてよいのは本文テキストと、段落/行単位の `{%p %}` / `{%tr %}`（`{%r %}` はラン単位）のみ。
- `w:vMerge` や `w:jc` など **XML 属性値に `{% %}` / `{{ }}` を入れてはいけない**。Word が属性値を厳密に検証するため、ファイルが開けなくなる。属性に依存する処理は `postprocess()` 側で行う。
- `w:tcPr` / `w:pPr` 内の子要素は OOXML スキーマの順序に従う必要がある（`_set_vmerge` / `_set_align` が挿入位置を調整しているのはこのため）。
- Word で直接タグを打つとランが分割されタグが壊れるので、テキストエディタで書いてから一括で貼り付ける。

## tools/make_template.py

元のスキルシート（LibreOffice 経由の .docx）から `template.docx` を作った一回限りの変換スクリプト。本文要素をインデックス（`P[16]` など）で参照しており元文書の構造に強く依存するため、通常は再実行せず、`template.docx` を直接編集する。LibreOffice 由来の Word 非互換（core-properties の関係タイプ、`characterSet`、`docGrid` の `charSpace`、`jc="start"/"end"`）の正規化処理を含む。

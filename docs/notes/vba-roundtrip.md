# 取り込んだ VBA を書き出すと、正本と何が変わるか

- Date: 2026-11-18
- 出典: 一時フォルダで `Import-Vba.ps1` と `Export-Vba.ps1` を流した実測（2026-11-16）と、git 2.56.0.windows.4 の `git diff --no-index` を、合成したファイルと実際の書き出しに流した実測（2026-11-18）。Microsoft 365 Excel 64-bit 16.0.20527.20196、Windows 11、Windows PowerShell 5.1.26100.9818。

正本（`vba/`）をブックへ取り込み、そのブックから書き出して、正本と比べた。VBE が整えた形で書かれた標準モジュール、クラス、文書モジュールは、1 バイトも変わらなかった。変わったのは、VBE が整える前の書き方、UserForm の `.frm` の空行、`.frx` のバイト列だった。

書き出しと `git diff --no-index` で正本と照合する手順の根拠である。手順は `plugins/xlwright/skills/xlwright/references/vba.md` の「正本との照合」にある。

## 実験の手順

1. 正本を 7 つ用意する。標準モジュール 4 つ、クラス 1 つ、文書モジュール（`ThisWorkbook`）1 つ、UserForm 1 つ。UserForm は、COM で作ったフォーム（ラベル 1 つとコード）を `Export-Vba.ps1` で書き出したもの。
2. 新しいブックへ `Import-Vba.ps1` で取り込み、`Export-Vba.ps1` で 2 回書き出す。正本と 1 回目、1 回目と 2 回目を `diff -r` で比べる。
3. 1 回目の書き出しを、別の新しいブックへ取り込んで書き出す。これをもう 1 巡くり返す。
4. 同じ正本を別の新しいブックへ取り込んで書き出し、2 の 1 回目と比べる。

標準モジュールのうち 1 つ（`Messy`）は、VBE が整える前の形で書いた。小文字のキーワード、演算子の前後に空白の無い式、行末の空白、宣言と違う大文字と小文字の識別子、`Attribute` の後の空行 2 つ、末尾の空行 3 つ。もう 1 つ（`Other`）は、`Messy` の引数 `Amount` と同じ名前の変数を `AMOUNT` と宣言し、ファイルの末尾に改行を置かなかった。

## 結果

### 変わらなかったもの

- VBE が整えた形の標準モジュール、日本語のコメントと文字列（半角カナを含む）を持つ標準モジュール、クラス、文書モジュールは、正本と書き出しが一致した。
- `Messy` の、`Attribute` の後の空行 2 つと、末尾の空行 3 つは、そのまま残った。コメントの前の空白も残った。
- 同じブックからの 2 回の書き出しは、テキストのファイルがすべて一致した。
- 同じ正本を別のブックへ取り込んだ書き出しも、テキストのファイルがすべて一致した。

### 変わったもの

| 正本 | 書き出し | 場所 |
| --- | --- | --- |
| `option explicit`、`public function`、`byval`、`as long`、`dim`、`end function` | `Option Explicit`、`Public Function`、`ByVal`、`As Long`、`Dim`、`End Function` | `Messy` |
| `Result=amount*2`（行末に空白 3 つ） | `result = AMOUNT * 2` | `Messy` |
| `total = result` | `Total = result` | `Messy` |
| `' comment` の後の空白 3 つ | 空白なし | `Messy` |
| 引数の宣言 `Amount` | `AMOUNT` | `Messy` |
| 末尾に改行の無いファイル | 末尾に改行 | `Other` |
| `Attribute` の後が、すぐコード | `Attribute` の後に空行 1 つ | `frmInfo.frm` |
| `.frx` のバイト列 | 2584 バイト中 288 バイトが違う | `frmInfo.frx` |

- 識別子の大文字と小文字は、プロジェクトの全体でそろえられる。`Messy` の `Amount` は、後から取り込んだ `Other` の宣言 `AMOUNT` に合わせて書き換わった。あるモジュールを直すと、触っていないモジュールの書き出しが変わる。
- `.frm` の空行は、取り込んで書き出すたびに 1 つずつ増えた。3 巡で 3 つになった。
- `.frx` は、同じブックからの 2 回の書き出しの間でも 9 バイト違った。大きさは、どの回も 2584 バイトだった。

取り込んで書き出しても変わらないことは、`plugins/xlwright/skills/xlwright/tests/Test-XlWright.ps1` の `cells-and-vba` が、試験のデータの標準モジュール、クラス、文書モジュールで確かめる。

## git diff --no-index で比べた時

実際の書き出しは、次の 2 つ。

- 試験のデータ（`plugins/xlwright/skills/xlwright/tests/data/vba`）を新しいブックへ取り込み、正本に無い標準モジュール 1 つとシートのコードを足して書き出したもの。
- 試験の `uf-create` で UserForm（`frmInfo`）を作り、書き出して取り込み、もう一度書き出したもの。

| 場面 | `git diff --no-index --ignore-blank-lines` の出力 | 終了コード |
| --- | --- | --- |
| 正本と同じ書き出し（UserForm なし） | 何も出ない | 0 |
| ブックにだけある標準モジュールとシートのコード | 書き出しの側の `new file` | 1 |
| `.frm` の `Attribute` の後に空行が 1 つ多いだけ（合成） | 本文には出ない | 1 |
| `.frx`（往復で 5144 バイトから 5656 バイト） | `Binary files ... differ` | 1 |
| `-- ':!*.frx'` を足した UserForm の往復 | 何も出ない | 1 |
| 末尾に改行の無いファイルと、ある書き出し（合成） | `\ No newline at end of file` の差 | 1 |
| `Util.bas` と `util.bas`（合成） | 削除と追加 | 1 |

- `--ignore-blank-lines` は差の本文から空行を隠すだけで、終了コードは、ファイルのバイト列が違えば 1 のまま。`--stat` と `--name-status` には `.frm` が出る。
- そのため、UserForm のあるブックでは、終了コードで一致を判定できない（照合の手順は `references/vba.md` の「正本との照合」）。
- 正本に無いものを足したブックで、足した標準モジュールを `Remove` で消し、シートのコードを `DeleteLines` で空にして書き出すと、正本との差は無くなり、終了コードは 0 になった（`references/vba.md` の「正本から消したモジュール」の ScriptBlock）。
- その ScriptBlock に、1 つ目にある名前、2 つ目に無い名前を渡すと、2 つ目の `Item` で「インデックスが有効範囲にありません。」になり、ブックのファイルは変わらなかった。

## 未確認

- 識別子の大文字と小文字が、どの宣言に合わせてそろうか。確かめたのは、後から取り込んだモジュールの宣言に合った 1 例だけ。
- 文書モジュールの、空行やキーワードの書き方の違い。文書モジュールは `AddFromString` で書くので、標準モジュールと同じ整え方になる見込み。
- コントロールの多いフォーム、画像を持つフォームの `.frm` と `.frx`。
- `Attribute VB_Description` などの、プロシージャの中の属性。
- `#If` や行の継続を含むコード。

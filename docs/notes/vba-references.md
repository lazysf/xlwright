# VBA の参照設定は、書き出しと取り込みで引き継がれるか

- Date: 2026-11-18
- 出典: 一時フォルダの新しいブックで `Import-Vba.ps1`、`Export-Vba.ps1`、`Test-VbaCompile.ps1` を流し、`VBProject.References` を読んだ実測。Excel 16.0.20527（Microsoft 365 64-bit）、Windows 11、Windows PowerShell 5.1.26100.9818。

参照設定（Microsoft Scripting Runtime など）を使うコードは、正本の `vba/` から作り直したブックではコンパイルが通らなかった。参照設定は書き出しのどのファイルにも出ず、取り込みでも足されない。

## 実験の手順

1. 早期バインディングの標準モジュール `UsesDict` を用意する。`Dim d As Scripting.Dictionary`、`Set d = New Scripting.Dictionary` で、キーの数を返す関数 `CountKeys` を持つ。
2. `Invoke-ExcelScript.ps1 -Create` で作った新しい `.xlsm` に取り込み、参照設定を読み、`Test-VbaCompile.ps1` でコンパイルする。
3. 同じブックに `VBProject.References.AddFromGuid('{420B2830-E718-11CF-893D-00A0C9054228}', 1, 0)` で Scripting Runtime を足して保存し、コンパイルして、`Invoke-ExcelMacro` で `CountKeys` を動かす。
4. 3 のブックを `Export-Vba.ps1` で書き出し、別の新しい `.xlsm` に取り込んで、コンパイルする。
5. `Dim d As Object`、`Set d = CreateObject("Scripting.Dictionary")` に書き換えた `UsesDict`（遅延バインディング）を、別の新しい `.xlsm` に取り込み、コンパイルして動かす。

## 結果

| 場面 | コンパイル | `CountKeys` |
| --- | --- | --- |
| 2: 早期バインディング、参照設定なし | 失敗。「コンパイル エラー: ユーザ定義型は定義されていません。」（`UsesDict`、4 行 9 列） | 流していない |
| 3: 参照設定を足した後 | 通った | 1 |
| 4: 3 から書き出して、新しいブックに取り込んだ | 失敗（2 と同じエラー） | 流していない |
| 5: 遅延バインディング、参照設定なし | 通った | 1 |

- 4 の書き出しは `standard/UsesDict.bas` だけだった。参照設定は、どのファイルにも出ない。
- `Invoke-ExcelScript.ps1 -Create` で作った新しいブックの参照設定は 4 つだった。

  | `Name` | `Description` | `Guid` | 版 | `BuiltIn` |
  | --- | --- | --- | --- | --- |
  | VBA | Visual Basic For Applications | `{000204EF-0000-0000-C000-000000000046}` | 4.2 | true |
  | Excel | Microsoft Excel 16.0 Object Library | `{00020813-0000-0000-C000-000000000046}` | 1.9 | true |
  | stdole | OLE Automation | `{00020430-0000-0000-C000-000000000046}` | 2.0 | false |
  | Office | Microsoft Office 16.0 Object Library | `{2DF8D04C-5BFA-101B-BDE5-00AA0044DE52}` | 2.8 | false |

- 3 で足した Scripting Runtime は、`Name` が `Scripting`、`Description` が `Microsoft Scripting Runtime`、版が 1.0、`Type` が 0（タイプライブラリ）、`BuiltIn` が false だった。
- `FullPath` は、参照先のファイルの絶対パス（`scrrun.dll` なら `C:\Windows\System32\scrrun.dll`）を返した。端末によって変わりうる値である。
- `AddFromGuid` は、確認のダイアログを出さずに参照設定を足した。

## ほかのブックへの参照

ほかの `.xlsm` への参照（`Type` が 1 のプロジェクト参照）も、同じ版の Excel で、一時フォルダの新しいブックを使って確かめた（2026-11-18）。

- `VBProject.References.AddFromFile` で足せたのは、参照する側をマクロを有効にして開き、参照先にモジュールが 1 つ以上ある時だけだった。どちらかを欠くと「DLL 読み込み時のエラーです。」で失敗した。参照先の VBA プロジェクトの名前は、参照する側と同じ既定の `VBAProject` から変えてから足した。同じ名前のまま足せるかは試していない。
- マクロを無効にして開いたブック（`Export-Vba.ps1` はこの開き方をする）では、参照から読めたのは `FullPath` だけだった。`Name`、`Description`、`Guid`、`IsBroken` は空で、版は 0.0 だった。
- マクロを有効にして開くと、`Name` は参照先の VBA プロジェクトの名前になり、`IsBroken` は false だった。`Guid` は空で、版は 0.0 のままだった。

`Guid` が無いので、この参照は `references.json` に書かない。書くと、`Guid` の無い項目を断る取り込みが止まる。そのため、書き出しは参照を書かずに、読める `FullPath` を警告に出す。

## references.json の形にした理由

形式と振る舞いは、2026-11-18 に人と決めた。使い方の正本は `plugins/xlwright/skills/xlwright/references/vba.md` の「書き出しと取り込み」にある。

- **`BuiltIn` だけを除き、`stdole` と `Office` は書く。** VBA と Excel は外せないので、書いても照らす意味が無い。`stdole` と `Office` は新しいブックに最初からあるが、書いておけば、外したブックには取り込みの警告で、版の違いには書き出しの差で気づける。
- **`Guid` で照らし、版は比べない。** 版の違う参照設定の差し替えは、しないと決めた。
- **取り込みは、足さずに警告する。** 参照先のファイルがその端末に無い時の `AddFromGuid` の振る舞いが分からない（「未確認」）。足りない参照設定は、後の `Test-VbaCompile.ps1` でも見える。
- **`FullPath` は書かない。** 端末で変わる値で、正本に入れると、書き出した端末ごとに差になりうる。

`plugins/xlwright/skills/xlwright/tests/Test-XlWright.ps1` の `cells-and-vba` と `input-rejected` が、これを確かめる（2026-11-18）。何を確かめるかは、試験のファイルを正本とする。

- スクリプトをわざと壊した次の写しは、どれもこれらの項目で落ちた。`BuiltIn` を除かない写し、警告を出さない写し、`Guid` の無い項目を断らない写し、プロジェクト参照を `references.json` に書く写し、プロジェクト参照の警告に `Name` を出す写し。

## 未確認

- アドイン（`.xlam`）への参照。確かめたのは `.xlsm` への参照だけ。
- 参照先が端末に無い参照設定（`IsBroken` が true）の読み方と、その時のコンパイルのエラー。参照先のファイルが無い時の `AddFromGuid` の振る舞い。
- 版の違う参照設定（`Office` の 2.8 以外など）が、どう読めて、どう振る舞うか。
- VBA プロジェクトがロックされたブックでの `References` の読み取り。

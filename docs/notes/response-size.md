# 日本語を含む応答の大きさ

設計の[期限と大きさの上限](../architecture.md#期限と大きさの上限)の未確認（段階1）を確かめた記録です。上限の20,000文字の近くの日本語の応答が、Claude Codeの会話にどの形で届き、打ち切られるかを見ました。
確かめた日は2026-11-04です。

## 版

- Windows 11 Pro 10.0.26300、Claude Codeのデスクトップアプリ（Codeタブ）
- ビルド: `6ba3037-dirty-20261103T120611Z`（開発版のプラグイン）
- Microsoft 365 Excel 64-bit 16.0.20527.20112

## 再現の手順

1. [段階1の流れ](stage1-flow.md)のアプリのフォルダ`.xlwright/scratch/stage1/`の`Calc.bas`に、2つを加えます。
   - `Kana(n)`: `n`文字のひらがなを返す`Function`。
   - `FillKana(rows, n)`: `Sheet1`の`A1`から`rows`行に、`n`文字のひらがなを書く`Sub`。
2. 会話から`session_start`の後、次を呼びます。
   - `macro_run`で`Calc.Kana`（引数5）と`Calc.Kana`（引数`"x"`、型の不一致）。
   - `macro_run`で`Calc.FillKana`（40と80）の後、`range_read`で`A1:A40`。
   - `macro_run`で`Calc.FillKana`（40と85）の後、`range_read`で`A1:A40`。
   - `macro_run`で`Calc.Kana`（引数1667）。
3. `session_end`（`save`が`false`）の後、`EXCEL.EXE`の数と主ブックのハッシュを照会します。

## 判定

### 会話に届く形

支持。成功の応答の日本語は、escapeされずに会話に届いた。失敗の応答の日本語は、`\uXXXX`のまま届いた（型の不一致の`description`）。
サーバーは、成功では`structuredContent`と本文の両方を、失敗では本文だけを返す（`ToolResults`）。本文はどちらもASCII以外をescapeする。そのため、成功ではClaude Codeが`structuredContent`から出し直した形が届いていると見られる。

### 上限の文字数の数え方

支持。サーバーは、escapeした後の本文で数える。日本語の1文字は6文字に数えられる。

- `range_read`: ひらがな3,200文字（40行×80文字）は、上限の中で返った。3,400文字（40行×85文字）は、20,687文字として`TOO_LARGE`になった。
- `macro_run`: 1,667文字の戻り値（escapeした後は10,004文字）は、`result_file`へ書かれ、応答には先頭の1,000文字が返った。結果のファイルの`return`は1,667文字だった。

### 打ち切り

支持（1回）。ひらがな3,200文字の`range_read`の応答は、Claude Codeに打ち切られず、警告も出なかった。

### escapeをやめた版

ビルド`3ec28b9-dirty-20261103T235001Z`（本文と結果のファイルでASCII以外をescapeしない）を、新しい会話で読んだ。手順は上と同じで、引数を次に変えた。

- 支持。型の不一致の失敗の本文は、`message`と`data.description`の日本語が、escapeされずに届いた。
- 支持（1回）。`FillKana`（40と490）の後の`range_read`は、ひらがな19,600文字を含む応答を返した。40行目と本文の末尾まで届き、打ち切りも警告も無かった。
- 支持。`FillKana`（40と500）の後の`range_read`は、20,287文字として`TOO_LARGE`になった。
- 支持。9,000文字の戻り値は、`result_file`を付けずに応答に入った。

## 後始末とGit差分

- `EXCEL.EXE`は、どちらの版の実験の後も0だった。主ブックは保存せずに終えた。
- 実験の後の`git status`は、実験の前と同じだった。

## 未確認

- トークンの数は測っていない。日本語の約20,000文字の応答が、打ち切られずに届くことだけを確かめた。上限までの余裕は分からない。

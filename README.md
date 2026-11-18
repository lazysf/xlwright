# xlwright

Excel を土台にしたツールやアプリケーション（マクロ、UserForm、カスタムリボンを持ち、人が操作して使うブック）を作り、確かめるための Claude Code のスキルです。
Windows 版 Excel を PowerShell の COM で操作します。ファイルを直接書き換えるのではなく、実際の Excel を動かします。

次の用途は対象にしません。ほかの表計算のスキル（Claude の `xlsx` スキルなど）を使ってください。

- 受け取ったブックの分析や、集計が合っているかの確認
- データの整形、変換、グラフやレポートの作成
- Excel を起動せずにファイルを読み書きすれば足りる作業

できること:

- ツールのシートの値・数式・書式の確認と編集
- ブックが人にどう見えるか（リボン、書式、グラフ、ボタン）を画像で確かめること
- VBA（標準・クラス・文書モジュール）と UserForm の書き出し・編集・取り込み
- VBA のコンパイルと、コンパイルエラーの文言と場所の取得
- マクロの実行と、実行時エラーや MsgBox の文言の取得
- カスタムリボンの埋め込みと、見え方を XML と照らすこと、ボタンを押して動作を確かめること
- VBA プロジェクトをロックした配布用のブックの出力

使い方の正本は [plugins/xlwright/skills/xlwright/SKILL.md](plugins/xlwright/skills/xlwright/SKILL.md) です。

## 前提

- Windows のデスクトップ版 Excel（確かめたのは Microsoft 365 Excel 64-bit）
- Windows PowerShell 5.1（OS 標準）
- Excel のトラストセンターで「VBA プロジェクト オブジェクト モデルへのアクセスを信頼する」が有効（VBA を扱う時）
- WMI（`Win32_Process.Create`）からプロセスを作れること。Excel を止める見張りのために使います。作れない環境では使えません。
- エージェントの作業中、人はその端末で Excel を使わないこと。Excel 以外のアプリは使えます

## 入れ方

Claude Code のプラグインとして入れます。この repo を marketplace として足してから、プラグインを入れます。`<repo>` には、GitHub の `owner/repo` を渡します。

```text
/plugin marketplace add <repo>
/plugin install xlwright@xlwright
```

入れたスキルは `/xlwright:xlwright` という名前になり、短く `/xlwright` でも呼べます。依頼の内容から自動でも選ばれます。

## 動作確認

手元の Excel で動くかを確かめます。一時フォルダに新しいブックを作って試し、ほかのファイルは変えません。Excel が動いている時は、何もせずに止まります。

```powershell
powershell.exe -NoProfile -File "$env:USERPROFILE\.claude\plugins\cache\xlwright\xlwright\<版>\skills\xlwright\tests\Test-XlWright.ps1"
```

`<版>` は、入れた版の番号（`0.1.0` など）を名前に持つフォルダです。`-Full` を付けると、Excel と VBE を画面に出す試験（コンパイルエラーの検出、マクロが出すダイアログの扱い、リボンの押下と見え方の照合、配布ブックの出力、ブックの見た目の画像）も行います。

## 開発

手直しの手順と repo の構成は [DEVELOPMENT.md](DEVELOPMENT.md) にあります。版ごとの変更は [CHANGELOG.md](CHANGELOG.md) にあります。

## 助けと貢献

質問や不具合の報告は、この repo の GitHub の issue で受け付けます。pull request は受け付けていません。

## ライセンス

[MIT License](LICENSE)

# 主ブックの共有

設計の未確認のうち、主ブックをほかのExcelと共有する時の振る舞いの2つを確かめた記録です。

- [開始の順序](../architecture.md#開始の順序)の4: ほかのExcelが主ブックを編集で開いている時に共有違反が起き、読み取り専用の時には起きないか。
- [片付け](../architecture.md#残った専用のexcelの片付け): 専用のExcelが主ブックを開いている間に、ほかのExcelで読み取り専用で開けるか。

確かめた日は2026-11-04です。

## 版

- Windows 11 Pro 10.0.26300、.NET 10.0.14（x64）
- Microsoft 365 Excel 64-bit 16.0.20527.20112
- 起動した場所: Claude Codeのデスクトップアプリの会話のシェル（Git Bash）から起動した`dotnet`

## 再現の手順

```powershell
dotnet run --project spikes\share -- .xlwright\spikes\share\Share.xlsm
```

- ブックは、[段階1の流れ](stage1-flow.md)の主ブックの写しです。
- ほかのExcelの代わりに、所有を確かめた非表示の専用のExcel（新しいPID、`EXCEL`、起動を求めた後の生成）を使います。
- 共有の確かめ方は、開始の4のとおりです。`FileAccess.Write`と`FileShare.ReadWrite`で、ファイルを開いてすぐ閉じます。
- 次の順に行います。
  1. どのExcelも開いていない時に、共有を確かめる。
  2. 1つ目のExcelが編集で開いている時に、共有を確かめる。
  3. その間に、2つ目のExcelで`ReadOnly`を真にして開く。
  4. その間に、3つ目のExcelで引数を付けずに開く（`DisplayAlerts`は偽）。
  5. 1つ目のExcelが`ReadOnly`を真にして開いている時に、共有を確かめる。
- 各Excelは、ブックを閉じてQuitし、5秒で終わらなければ、所有の記録と照合して終了させます。

## 判定

### 開始の4の共有違反

支持（1回）。

- どのExcelも開いていない時は、開けた。
- 編集で開いている時は、`0x80070020`（共有違反）の`IOException`になった。
- 読み取り専用で開いている時は、開けた。

### ほかのExcelで読み取り専用で開く

支持（1回）。編集で開いている間に、2つ目のExcelは`ReadOnly`を真にして112 msで開けた。開いたブックの`ReadOnly`は真だった。

### ほかのExcelで引数を付けずに開く

支持（1回）。`DisplayAlerts`を偽にした3つ目のExcelは、止まらずに147 msで開き、開いたブックの`ReadOnly`は真だった。開始の6の`ReadOnly`の判定は、ダイアログで止まらずに働く。

### 日常のExcelとの違い

ほかのExcelは、非表示の専用のExcelで代えた。人が表示されたExcelで開く時に出るダイアログ（使用中の通知など）は、確かめていない。

## 後始末とGit差分

- 起動した4つのExcelは、どれもQuitから5秒では終わらず、所有の記録と照合してから終了させた。
- `EXCEL.EXE`は、実験の前と後で0だった。
- 実験の後の`git status`は、TODO.mdの周期22の追記と、`spikes/share/`の追加だけだった。

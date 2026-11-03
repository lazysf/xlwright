# AIエージェントのシェルからの実Excelの起動と終了

[設計](../architecture.md#本体の部品の境界)の段階0の未確認「AIエージェントのシェルから、実Excelを起動して終了できるか」を確かめた記録です。

## 版

- Windows 11 Pro 10.0.26300、Windows PowerShell 5.1.26100（STA）
- Microsoft 365 Excel 64-bit 16.0.20527.20112
- 起動した場所: Claude Codeのデスクトップアプリの会話のシェル（Git Bash）から起動した`powershell.exe`

## 再現の手順

```powershell
powershell -NoProfile -File spikes\excel-start-stop\Test-ExcelStartStop.ps1
```

[スクリプト](../../spikes/excel-start-stop/Test-ExcelStartStop.ps1)の流れは次のとおりです。

1. 実行の前からあるEXCELのプロセスを記録する。
2. 時刻を記録してから、`New-Object -ComObject Excel.Application`で起動する。
3. `Hwnd`から`GetWindowThreadProcessId`でPIDを引き、所有を確かめる。条件は、プロセスの名前が`EXCEL.EXE`であること、記録した時刻より後に生成されたこと、実行の前に無かったPIDであること。
4. 確かめられた時だけ、ブックと保護ビューのウィンドウが0であることを照会して、`Quit`を求める。
5. COMの参照を解放し、所有を確かめた時に開いたプロセスのハンドルで、終了を待つ。
6. 実行の前からあったプロセスが変わっていないことを照合する。

実行の前に、EXCELのプロセスは無かった。

## 判定

支持。AIエージェントのシェルから、専用のExcelを起動して終了できた。

- 起動: `New-Object`から戻るまで2,002 ms。所有の条件は、すべて満たした。
- `Visible`は偽、ブックは0、保護ビューのウィンドウは0だった。
- `Quit`は119 msで戻り、プロセスはその後1,593 msで終わった（終了コード0）。強制終了は使っていない。
- 実行の後に、EXCELのプロセスは残っていなかった。

計測は1回だけです。期限の値を決める根拠には、段階1で製品のサーバーから測った値を使います。

参考に見たこと（段階1の未確認に関わる観察。製品のサーバーの中で確かめ直します）:

- COMから起動したEXCELのコマンドラインは`"…\EXCEL.EXE" /automation -Embedding`で、`/automation`を含んでいた（[所有の確定](../architecture.md#所有の確定)）。

## 後始末とGit差分

- 実験が起動したExcelは、`Quit`で終わった。実験の後に、EXCELのプロセスが0であることを確かめた。
- Gitの差分は、スクリプトと本書の追加と、TODO.mdの更新だけです。ブックは開いていない。

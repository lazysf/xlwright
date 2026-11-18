# Codex の sandbox で xlwright のスクリプトが動くか

- Date: 2026-11-18
- 出典: Codex の desktop app（`OpenAI.Codex` 26.1111.2228.0）に同梱された `codex.exe`（`codex-cli 0.169.0-alpha.11.1`）の `codex sandbox <command>` で、`powershell.exe` を sandbox の中で起動した実測。`~/.codex/config.toml` は `sandbox_mode = "workspace-write"`、`[windows] sandbox = "mxc"`。Windows 11、Windows PowerShell 5.1.26100.9818。公式 docs は [Package your plugin](https://developers.openai.com/plugins/build/plugins)、[Build skills](https://learn.chatgpt.com/docs/build-skills.md)、[Windows sandbox](https://learn.chatgpt.com/docs/windows/windows-sandbox.md)。

Codex の MXC の sandbox の中では、xlwright のスクリプトは動かなかった。COM で Excel を起動する所で、アクセスを拒否された。Codex に配るかを決める前に、この制限を前提にする。

## 実測

`codex sandbox` の中の `powershell.exe` を、sandbox の外と比べた。

| 試したこと | sandbox の外 | sandbox の中 |
| --- | --- | --- |
| `powershell.exe -File <script>.ps1` | 動いた | 「AuthorizationManager チェックが失敗しました。」で止まった |
| `Import-Module XlWright.psm1`（`-Command` から） | 動いた | 同じエラーで止まった |
| `Get-ExecutionPolicy -List` | CurrentUser が RemoteSigned | 「要求されたレジストリ アクセスは許可されていません。」 |
| `-Command` に直接書いた 1 行の処理 | 動いた | 動いた |

実行ポリシーのレジストリを読めないことが、スクリプトが止まる原因と見ている。docs や他の報告とは突き合わせていない。

続けて、人の許可を得て、この実験に限って `powershell.exe -ExecutionPolicy Bypass` で起動した。

| 試したこと | sandbox の中の結果 |
| --- | --- |
| スクリプトの実行 | 動いた |
| 環境 | AppContainer だった（`TEMP` が `~\AppData\Local\Packages\sandbox.{<GUID>}\AC\Temp`）。`whoami` は空、Session は 1、`UserInteractive` は True。 |
| 作業フォルダ（起動した時のカレント）への書き込み | アクセス拒否 |
| `TEMP` への書き込み | 動いた |
| `Get-Process` | 3 つしか見えなかった（外では 263）。`explorer` は見えない。 |
| HKCU の `Software\Microsoft\Office\16.0\Excel\Security` | 「要求されたレジストリ アクセスは許可されていません。」 |
| `Get-CimInstance Win32_OperatingSystem` | 「クライアントが CIM リソースへのアクセスを取得できませんでした。」 |
| `Test-Prerequisite.ps1` | `Ready` が false。WMI を使えないため。`AccessVBOM` の `Value` は null。 |
| `Invoke-ExcelScript.ps1 -Create -Save`（`tests/data/scripts/write.ps1`） | `CLSID {00024500-0000-0000-C000-000000000046}` のクラスファクトリの取得で `0x80070005`（`E_ACCESSDENIED`）。Excel は起動せず、ブックはできなかった。 |

xlwright は、COM の Excel、WMI（`Use-ExcelWorkbook` の監視役の起動）、HKCU のレジストリ（`AccessVBOM`）に頼る。MXC の sandbox では、この 3 つがどれも使えない。リボンの押下、画面の画像、`Lock-VbaProject.ps1` は、Excel を起動できないので試していない。

## Codex のエージェントで使った結果

人が Codex の desktop app で試し、エージェントの記録を受け取った（2026-11-18）。権限は「代わりに承認」で、作業フォルダは xlwright と関係の無い repo だった。

- **入れ方**: app の GUI で、GitHub の repo（`main`）を marketplace として足した。`.claude-plugin/marketplace.json`（`source` が文字列）をそのまま読み、プラグイン xlwright とスキル xlwright が一覧に出た。manifest は足していない。
- **置き場**: 入れたプラグインは `~/.codex/plugins/cache/<marketplace>/<plugin>/<版>/`（ここでは `xlwright/xlwright/<版>/`）に置かれた。
- **スキルの選び方とパス**: 依頼から xlwright が選ばれ、`SKILL.md` が読まれた。`${CLAUDE_SKILL_DIR}` は展開されず、エージェントはまず間違ったパス（`cache/xlwright/<版>/`）を読み、`rg` で `SKILL.md` を探し直して正しいパスを使った。2 回とも同じだった。

| 依頼 | 振る舞い |
| --- | --- |
| book.xlsm を作り、A1 と B1 を書いて、B1 の値を答える | sandbox の中で `Test-Prerequisite.ps1` が拒否された。昇格を求めず、Excel を使わずに Open XML の zip を手で組み立てて `book.xlsm` を作り、自分で書いたキャッシュの値 `2` を「確認した」と答えた。 |
| 同じ依頼に、「拒否されたらほかの手段に切り替えず、sandbox の外での実行を求める。通らなければ止まる」を足した | sandbox の中で `UnauthorizedAccess` が出た後、sandbox の外での実行を求めて承認された。外では `-ExecutionPolicy Bypass` なしで動いた。試験の `invoke-excelscript` を流し、`Invoke-ExcelScript.ps1 -Create -Save` と `Write-ExcelRange` で書き、開き直して `Read-ExcelRange` で照合し、`Find-CellError.ps1` を流した。B1 は Excel が計算した `2` だった。 |

1 回目の迂回を防ぐため、`SKILL.md` の「利用の前提」に、拒否された時は Excel を使わない手段で代わりをせず、権限の引き上げを求めるか止まる、と書いた。この規則だけで（依頼に指示を足さずに）昇格を求めるかは、まだ試していない。

パスを探し直さずに済むよう、`SKILL.md` のコマンド例は、スキルのフォルダを `${CLAUDE_SKILL_DIR}` でなく `<skill>`（`SKILL.md` があるフォルダ）と書く。Claude Code のための `${CLAUDE_SKILL_DIR}` は、説明の 1 文にだけ残す。

## Codex に配る時の前提

公式 docs から読んだこと。marketplace の読み方とスキルの選び方は、上の実機で確かめた。

- Codex は、`name` と `description` を持つ `SKILL.md` を読む。プラグインの manifest は `.codex-plugin/plugin.json`、marketplace は `.agents/plugins/marketplace.json` に置く。desktop app は `.claude-plugin/marketplace.json` も読める。
- `${CLAUDE_SKILL_DIR}` にあたる変数は、Codex の docs に無い。
- Windows で sandbox の外で動かすには、full access にする。docs は、その時に Codex の操作が作業フォルダに限られなくなると注意している。

## 未確認

- `SKILL.md` に足した規則だけで、Codex のエージェントが昇格を求めるか。
- `<skill>` の書き方にした `SKILL.md` で、Codex のエージェントが最初から正しいパスを使うか。Claude Code で、`<skill>` を `${CLAUDE_SKILL_DIR}` の値に置き換えて実行するか。
- 承認が自動で通ったのか、人に確認を求めたのか。2 回目の作業の後に Excel が残らなかったか。
- リボンの押下、画面の画像、`Lock-VbaProject.ps1` が、sandbox の外での実行で動くか。
- 旧来の sandbox（elevated、unelevated）での振る舞い。どちらも、マシン全体にアカウントや規則を作る準備が要るので試していない。

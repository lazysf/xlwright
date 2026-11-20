# Invoke-RibbonButton.ps1 で押せるリボンの要素

- Date: 2026-11-20
- 出典: 一時フォルダの新しいブックで `Invoke-RibbonButton.ps1` を流した実測と、同じタブを UI Automation と MSAA で操作した実測。Excel 16.0.20527.20196、Windows 11、Windows PowerShell 5.1.26100.9818。

`Invoke-RibbonButton.ps1` は、選んだタブの中身から、UI Automation で `Button` に見える要素を表示名で探す。試した要素のうち、button のほかに押せたのは toggleButton だけだった。menu と dropDown の中の項目は、開く操作を挟めば、同じ MSAA の既定の動作で押せた。

## 実験の手順

1. 新しい `.xlsm` に、callback が MsgBox を出す標準モジュールを取り込む。
2. 1 つのグループに次の 3 つを持つリボンを、`Add-CustomUI.ps1` で埋め込む。どれも `size="large"`。
   - `toggleButton`（`Toggle`）。`onAction` は `(control As IRibbonControl, pressed As Boolean)`。
   - `menu`（`Menu`）と、その中の `button`（`MenuItem`）。
   - `dropDown`（`Drop`）と、その中の `item`（`ItemA`、`ItemB`）。
3. `Invoke-RibbonButton.ps1` に、`Toggle`、`Menu`、`MenuItem`、`Drop`、`ItemA` を 1 つずつ渡して押す。
4. 別に、タブを選んだ後のタブの中身を UI Automation で列挙し、要素の種類と pattern を見る。
5. 別に、`Ribbon.psm1` の関数でタブを選び、menu と dropDown に MSAA の既定の動作をかけて開く。開いた後に中の項目を Excel のウィンドウから探し、その項目に MSAA の既定の動作をかける。

## 結果

| 要素 | UI Automation の種類 | pattern | `Invoke-RibbonButton.ps1` |
| --- | --- | --- | --- |
| toggleButton | `Button` | `Toggle` | 押せた。callback が呼ばれ、MsgBox に `toggle True` が出た |
| menu | `MenuItem` | `ExpandCollapse` | `Not found` |
| menu の中の button | タブを選んだだけでは、タブの中身に出ない | - | `Not found` |
| dropDown | `ComboBox`（中に表示名と同じ `Text`、`開く` の `Button`） | `Value`、`ExpandCollapse` | `Not found` |
| dropDown の item | タブを選んだだけでは、タブの中身に出ない | - | `Not found` |

- toggleButton は、押していない状態から 1 回押し、`pressed` に `True` が届いた。押した後の状態は、ブックを保存せずに閉じるので残らない。
- dropDown の中の `開く` ボタンは `Button` に見える。`-ButtonLabel 開く` は試していない。

### 開いてから押す

| 開いたもの | 開いた後の項目 | 項目を押した結果 |
| --- | --- | --- |
| menu | Excel のウィンドウの中に `MenuItem`（クラス `NetUITWBtnMenuItem`） | callback が呼ばれ、MsgBox に `menu item` が出た |
| dropDown | Excel のウィンドウの中に `ListItem`（クラス `NetUIGalleryButton`） | callback が呼ばれ、MsgBox に `drop itemA` が出た |

- 開く操作も押す操作も、表示名と位置で探した MSAA の要素の既定の動作で行った。`Invoke-RibbonButton.ps1` がボタンを押すのと同じ方法。
- 項目は Excel のウィンドウの子孫に出た。別のトップレベルのウィンドウを探す必要は無かった。
- 各 1 回の観察。

## 未確認

- menu や dropDown を開いた時に、Excel が前面に出るか。
- splitButton、gallery、comboBox、checkBox が、どの種類に見えるか。

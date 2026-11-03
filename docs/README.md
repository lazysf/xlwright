# 開発文書

ここにある文書を、xlwrightの現在の要件・設計・開発手順の正本とします。
設計上の採用と、実装・実機検証の完了は区別します。

- [requirements.md](requirements.md): 背景、目的、目標、範囲、利用環境、利用の前提、安全性の要件。
- [architecture.md](architecture.md): 仕組みの方式、実行時の状態の置き場、部品の境界、未確認の仮説と成立しない時の扱い。
- [api.md](api.md): MCPのツールの共通の規約と一覧。各ツールの契約は段階ごとに加える。
- [development.md](development.md): 開発の手順と確認の入口。
- [../ROADMAP.md](../ROADMAP.md): 開発の段階、到達する状態、完了の根拠。
- [../TODO.md](../TODO.md): 今の周期の作業と完了の条件。

取り消しにくい判断が出た時は、`adr/`に記録します。実機の確認や調査の結果は、`notes/`に短く残します。

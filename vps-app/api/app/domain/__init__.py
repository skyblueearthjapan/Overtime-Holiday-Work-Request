"""業務（担当 W）。画面の 29 関数（実装仕様 4-3）・管理者の 9 関数（4-2）・Sebastian の入口（4-4）。

- 登録は register.register_all(rpc.register)（rpc.load_domain_handlers が呼ぶ）。
- 業務関数は commit しない（rpc が冪等キー・監査と一緒に確定する）。Sebastian の handle だけは自分で commit する。
- 行を書いたら同じトランザクションで sheet_queue.mark_dirty、GAS の後処理は gasjobs（担当 M の gas_jobs）で積む。
"""

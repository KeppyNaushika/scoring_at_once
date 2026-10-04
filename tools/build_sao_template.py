"""score-at-once-electron の Prisma マイグレーションから, 空の archive.db テンプレートを作る.

.sao の archive.db は Electron 版のデータベースと同じスキーマ (全テーブル・全列) と
_prisma_migrations を持っている必要がある. 一括採点 (Python 版) はこのテンプレートを
コピーして行を挿入するだけで .sao を作る.

使い方:
    python3 tools/build_sao_template.py /path/to/score-at-once-electron

出力: assets/sao_template.db
Electron 版のスキーマが変わったら再生成する. 取り込む側より新しいマイグレーションを含むと
"newerSchema" で拒否されるので, 配布済みの Electron 版より新しいリビジョンからは作らないこと.
"""

import datetime
import hashlib
import os
import sqlite3
import sys
import uuid

OUTPUT = os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "assets", "sao_template.db")


def main(electron_repo: str) -> None:
    migrations_dir = os.path.join(electron_repo, "prisma", "migrations")
    names = sorted(
        d for d in os.listdir(migrations_dir) if os.path.isdir(os.path.join(migrations_dir, d))
    )
    if os.path.exists(OUTPUT):
        os.remove(OUTPUT)
    if sys.version_info < (3, 12):
        sys.exit("Python 3.12 以降で実行して下さい (sqlite3 の setconfig が必要)")
    db = sqlite3.connect(OUTPUT)
    db.isolation_level = None
    # 初期マイグレーションが writable_schema で sqlite_autoindex_* を作るため, defensive モードを切る
    db.setconfig(sqlite3.SQLITE_DBCONFIG_DEFENSIVE, False)
    db.execute(
        'CREATE TABLE "_prisma_migrations" ("id" TEXT PRIMARY KEY NOT NULL,"checksum" TEXT NOT NULL,'
        '"finished_at" DATETIME,"migration_name" TEXT NOT NULL,"logs" TEXT,"rolled_back_at" DATETIME,'
        '"started_at" DATETIME NOT NULL DEFAULT current_timestamp,'
        '"applied_steps_count" INTEGER UNSIGNED NOT NULL DEFAULT 0)'
    )
    for name in names:
        with open(os.path.join(migrations_dir, name, "migration.sql"), encoding="utf-8") as f:
            sql = f.read()
        db.execute("PRAGMA foreign_keys=OFF")
        db.executescript(sql)
        now = datetime.datetime.now(datetime.timezone.utc).isoformat(timespec="milliseconds")
        db.execute(
            "INSERT INTO _prisma_migrations (id, checksum, finished_at, migration_name, started_at,"
            " applied_steps_count) VALUES (?, ?, ?, ?, ?, 1)",
            (str(uuid.uuid4()), hashlib.sha256(sql.encode()).hexdigest(), now, name, now),
        )
    # .sao にはトリガー・ビューを含めてはいけない
    bad = db.execute("SELECT type, name FROM sqlite_master WHERE type NOT IN ('table', 'index')").fetchall()
    assert not bad, bad
    assert db.execute("PRAGMA integrity_check").fetchone()[0] == "ok"
    db.execute("PRAGMA journal_mode=DELETE")
    db.execute("VACUUM")
    db.close()
    print(f"{len(names)} migrations -> {os.path.normpath(OUTPUT)} (last: {names[-1]})")


if __name__ == "__main__":
    if len(sys.argv) != 2:
        sys.exit(__doc__)
    main(sys.argv[1])

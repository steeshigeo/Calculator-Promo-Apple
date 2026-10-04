#!/usr/bin/env python3
"""
Perbarui access.json (hash password staff / admin) dari input workflow 'Update Access'.

Browser admin menghitung hash PBKDF2-SHA256 (salt acak, 150.000 iterasi) lalu mengirim HANYA salt + hash
lewat workflow_dispatch. Password asli tidak pernah dikirim / disimpan. Script ini memvalidasi format lalu
menulis access.json; semua perangkat staff membaca file itu sehingga password baru berlaku untuk semuanya.
"""
import json
import os
import re
import sys
from datetime import datetime, timezone

OUT = "access.json"
SALT = re.compile(r"^[0-9a-f]{32}$")
HASH = re.compile(r"^[0-9a-f]{64}$")


def main():
    try:
        cur = json.load(open(OUT, encoding="utf-8"))
    except (OSError, ValueError):
        print("::error::access.json tidak ditemukan / rusak. Upload access.json awal dulu.")
        return 1
    now = datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")
    changed = []
    for role in ("staff", "admin"):
        salt = os.environ.get(f"{role.upper()}_SALT", "").strip().lower()
        hsh = os.environ.get(f"{role.upper()}_HASH", "").strip().lower()
        if not salt and not hsh:
            continue
        if not (SALT.match(salt) and HASH.match(hsh)):
            print(f"::error::Format salt/hash {role} tidak valid (salt 32 hex, hash 64 hex).")
            return 1
        cur[role] = {"salt": salt, "hash": hsh, "v": now}
        changed.append(role)
    if not changed:
        print("::error::Tidak ada perubahan: isi staff_* dan/atau admin_*.")
        return 1
    cur["updated_at"] = now
    with open(OUT, "w", encoding="utf-8") as f:
        json.dump(cur, f, ensure_ascii=False, indent=1)
        f.write("\n")
    print(f"access.json diperbarui: {', '.join(changed)}")
    return 0


if __name__ == "__main__":
    sys.exit(main())

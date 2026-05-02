from __future__ import annotations

import uuid


def new_snapshot_id() -> str:
    return uuid.uuid4().hex[:12]

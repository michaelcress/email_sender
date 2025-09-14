import json, dataclasses
from dataclasses import dataclass, fields
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Dict, Optional, List

@dataclass
class OAuthToken:
    access_token: str
    refresh_token: str = ""          # ← NEW
    expires_in: int = 0
    token_type: str = "Bearer"
    scope: str = ""
    obtained_at: int = 0
    ext_expires_in: Optional[int] = None

    # -------- Convenience --------
    @property
    def scopes(self) -> List[str]:
        return [s for s in self.scope.split() if s]

    @property
    def expires_at(self) -> datetime:
        return datetime.fromtimestamp(self.obtained_at + self.expires_in, tz=timezone.utc)

    @property
    def seconds_until_expiry(self) -> int:
        return max(0, int((self.expires_at - datetime.now(timezone.utc)).total_seconds()))

    def needs_refresh(self, skew_seconds: int = 120) -> bool:
        """True if expired or will expire within skew_seconds."""
        return self.seconds_until_expiry <= skew_seconds

    @property
    def authorization(self) -> dict:
        """HTTP Authorization header for requests."""
        return {"Authorization": f"{self.token_type} {self.access_token}"}

    # -------- Parsing helpers --------
    @classmethod
    def _coerce(cls, data: Dict[str, Any]) -> Dict[str, Any]:
        allowed = {f.name for f in fields(cls)}
        payload = {k: v for k, v in data.items() if k in allowed}
        # Normalize possibly-string numbers
        for k in ("expires_in", "obtained_at", "ext_expires_in"):
            if k in payload and payload[k] is not None:
                payload[k] = int(payload[k])
        return payload

    @classmethod
    def from_json(cls, s: str) -> "OAuthToken":
        return cls(**cls._coerce(json.loads(s)))

    @classmethod
    def from_file(cls, path: str | Path) -> "OAuthToken":
        with open(path, "r", encoding="utf-8") as f:
            return cls(**cls._coerce(json.load(f)))

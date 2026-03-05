from __future__ import annotations

from dataclasses import dataclass
from typing import Optional
import requests
from datetime import datetime


@dataclass
class PEZConfig:
    base_url: str = "https://pez.giantleap.net"
    x_bpid: str = "bp_aarhus"
    x_locale: str = "da"


class PEZClient:
    """
    Stateful PEZ client:
    - Holds one requests.Session() for the lifetime of the instance (cookies, keep-alive, headers).
    - Logs in once and caches the access token.
    - Each call uses per-request headers like your old code, but session-level defaults still apply.
    """

    def __init__(self, username: str, password: str, config: Optional[PEZConfig] = None) -> None:
        self.username = username
        self.password = password
        self.config = config or PEZConfig()

        self.session = requests.Session()
        # Browser-like defaults (same idea as your old code)
        self.session.headers.update({
            "accept-language": "en-US,en;q=0.9,en-AU;q=0.8,en-CA;q=0.7,en-IN;q=0.6,en-IE;q=0.5,en-NZ;q=0.4,en-GB-oxendict;q=0.3,en-GB;q=0.2,en-ZA;q=0.1",
            "sec-ch-ua": '"Not:A-Brand";v="99", "Microsoft Edge";v="145", "Chromium";v="145"',
            "sec-ch-ua-mobile": "?0",
            "sec-ch-ua-platform": '"Windows"',
            "user-agent": (
                "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                "AppleWebKit/537.36 (KHTML, like Gecko) "
                "Chrome/145.0.0.0 Safari/537.36 Edg/145.0.0.0"
            ),
        })

        self._access_token: Optional[str] = None

    # ---------- Internal helpers ----------

    def _url(self, path: str) -> str:
        return f"{self.config.base_url}{path}"

    def _auth_headers(self) -> dict:
        if not self._access_token:
            raise RuntimeError("PEZ client is not authenticated. Call login() first.")
        return {
            "accept": "application/json, text/plain, */*",
            "authorization": f"Bearer {self._access_token}",
            "x-bpid": self.config.x_bpid,
            "x-gltlocale": self.config.x_locale,
        }

    # ---------- Public API ----------

    def login(self) -> str:
        """
        Mimics your login flow and stores token + cookies in the session.
        Call once per robot run (or lazy-call from other methods).
        """
        # 1) Load login page
        r = self.session.get(
            self._url("/login"),
            headers={"accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8"},
            timeout=30,
        )
        r.raise_for_status()

        # 2) initiate-login
        r = self.session.post(
            self._url("/rest/public/initiate-login"),
            json={"username": self.username},
            headers={
                "accept": "application/json, text/plain, */*",
                "content-type": "application/json;charset=UTF-8",
                "origin": self.config.base_url,
                "referer": self._url("/login"),
            },
            timeout=30,
        )
        r.raise_for_status()

        # 3) oauth token
        r = self.session.post(
            self._url("/rest/oauth/token"),
            data={
                "client_id": "web-client",
                "grant_type": "password",
                "username": self.username,
                "password": self.password,
            },
            headers={
                "accept": "application/json, text/plain, */*",
                "content-type": "application/x-www-form-urlencoded",
                "origin": self.config.base_url,
                "referer": self._url("/login"),
                "authorization": "Basic d2ViLWNsaWVudDp3ZWItY2xpZW50",
            },
            timeout=30,
        )
        r.raise_for_status()

        data = r.json()
        token = data.get("access_token")
        if not token:
            raise RuntimeError("PEZ login succeeded but access_token missing in response.")
        self._access_token = token
        return token

    def ensure_login(self) -> None:
        """Lazy-login helper."""
        if not self._access_token:
            self.login()

    def add_internal_comment(self, case_uuid: str, comment: str) -> None:
        """
        Adds a single internal comment to a case.
        Uses same session (cookies) and same token (Authorization) every time.
        """
        self.ensure_login()

        url = self._url(f"/rest/tickets/cases/{case_uuid}/comments")
        payload = {"comment": comment, "isInternal": True}

        headers = {
            **self._auth_headers(),
            "content-type": "application/json;charset=UTF-8",
            "priority": "u=1, i",
        }

        r = self.session.post(url, headers=headers, json=payload, timeout=30)
        r.raise_for_status()


    @staticmethod
    def format_faktura_comment(
        order_number: str,
        tilladelsestype: str | None,
        totalpris,
        startdato=None,
        slutdato=None,
        antal_dage=None,
    ) -> str:
        ts = (tilladelsestype or "").strip() or "ukendt tilladelsestype"

        # Format price
        price_txt = "ukendt pris"
        if totalpris is not None:
            try:
                p = float(str(totalpris).replace(",", "."))
                price_txt = f"{p:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
            except Exception:
                price_txt = str(totalpris)

        # Format dates
        def format_date(d):
            if not d:
                return None
            if isinstance(d, datetime):
                return d.strftime("%d-%m-%Y")
            try:
                return datetime.strptime(str(d), "%Y-%m-%d").strftime("%d-%m-%Y")
            except Exception:
                return str(d)

        start = format_date(startdato)
        slut = format_date(slutdato)

        period_txt = ""
        days = int(antal_dage)
        dag_label = "dag" if days == 1 else "dage"
        period_txt = f" ({start} → {slut}, {days} {dag_label})"

        return f"Faktura sendt med ordrenummer {order_number} for {ts}{period_txt} på {price_txt} kr."
"""ipam.py — Cliente para API phpIPAM."""
import logging
from typing import Optional

import requests
from requests.exceptions import RequestException

from config import IPAM_BASE_URL, IPAM_REQUEST_TIMEOUT, IPAM_SECTION_ID

logger = logging.getLogger(__name__)


class IPAMClient:
    """Encapsula autenticação e chamadas à API do phpIPAM."""

    def __init__(
        self,
        user: str,
        password: str,
        base_url: str = IPAM_BASE_URL,
        timeout: int = IPAM_REQUEST_TIMEOUT,
    ) -> None:
        self.user = user
        self.password = password
        self.base_url = base_url
        self.timeout = timeout
        self.token: Optional[str] = None
        self.session = requests.Session()

    # ------------------------------------------------------------------
    # URLs derivadas
    # ------------------------------------------------------------------
    @property
    def url_auth(self) -> str:
        return f"{self.base_url}/api/ctp/user/"

    @property
    def url_subnets_section(self) -> str:
        return f"{self.base_url}/api/ctp/sections/{IPAM_SECTION_ID}/subnets/"

    # ------------------------------------------------------------------
    # Autenticação
    # ------------------------------------------------------------------
    def autenticar(self) -> None:
        logger.info("Autenticando na API phpIPAM...")
        try:
            resposta = self.session.post(
                self.url_auth, auth=(self.user, self.password),
                timeout=self.timeout,
            )
            resposta.raise_for_status()
        except RequestException as exc:
            raise RuntimeError(f"Falha na autenticação: {exc}") from exc

        dados = resposta.json()
        self.token = dados["data"]["token"]
        self.session.headers.update({"token": self.token})
        logger.info("Autenticação bem-sucedida.")

    # ------------------------------------------------------------------
    # Requisições genéricas
    # ------------------------------------------------------------------
    def _get(self, url: str) -> dict:
        try:
            resposta = self.session.get(url, timeout=self.timeout)
            resposta.raise_for_status()
        except RequestException as exc:
            raise RuntimeError(f"Erro GET [{url}]: {exc}") from exc
        return resposta.json()

    def _post(self, url: str, payload: dict) -> requests.Response:
        try:
            return self.session.post(url, json=payload, timeout=self.timeout)
        except RequestException as exc:
            raise RuntimeError(f"Erro POST [{url}]: {exc}") from exc

    # ------------------------------------------------------------------
    # Lógica de negócio
    # ------------------------------------------------------------------
    def buscar_id_cn(self, nome_cn: str) -> int:
        nome_cn = nome_cn.strip()
        if not nome_cn.upper().startswith("CN"):
            nome_cn = f"CN {nome_cn.zfill(2)}"
        logger.info("Buscando CN '%s'...", nome_cn)
        dados = self._get(self.url_subnets_section)
        for item in dados.get("data", []):
            if item.get("description") == nome_cn:
                logger.info("CN '%s' → ID %s", nome_cn, item["id"])
                return int(item["id"])
        raise ValueError(f"CN '{nome_cn}' não encontrado na seção {IPAM_SECTION_ID}.")

    def buscar_id_pool(self, id_cn: int, nome_pool: str) -> int:
        logger.info("Buscando pool '%s' no CN ID %d...", nome_pool, id_cn)
        url = f"{self.base_url}/api/ctp/subnets/{id_cn}/slaves/"
        dados = self._get(url)
        for item in dados.get("data", []):
            if nome_pool.upper() in item.get("description", "").upper():
                logger.info("Pool '%s' → ID %s", nome_pool, item["id"])
                return int(item["id"])
        raise ValueError(f"Pool '{nome_pool}' não encontrado no CN ID {id_cn}.")

    def buscar_rede_pai(
        self, id_cn: int, id_pool: int, mascara: str
    ) -> Optional[int]:
        logger.info("Verificando rede pai /%s no pool ID %d...", mascara, id_pool)
        url = f"{self.base_url}/api/ctp/subnets/{id_cn}/slaves/{id_pool}"
        try:
            dados = self._get(url)
        except RuntimeError as exc:
            logger.warning("Não foi possível verificar rede pai: %s", exc)
            return None
        for item in dados.get("data", []):
            if item.get("mask") == mascara:
                logger.info("Rede pai (/%s) → ID %s", mascara, item["id"])
                return int(item["id"])
        logger.warning("Rede pai /%s não localizada no pool ID %d. Prosseguindo.", mascara, id_pool)
        return None

    def reservar_rede(self, id_pool: int, mascara: str, descricao: str) -> dict:
        url = f"{self.base_url}/api/ctp/subnets/{id_pool}/first_subnet/{mascara}/"
        logger.info("Reservando rede /%s no pool ID %d...", mascara, id_pool)
        resposta = self._post(url, {"description": descricao})
        if resposta.status_code in (200, 201):
            nova_rede = resposta.json().get("data")
            logger.info("Rede /%s reservada: %s", mascara, nova_rede)
            return nova_rede
        raise RuntimeError(
            f"Falha ao reservar subnet (HTTP {resposta.status_code}): {resposta.text}"
        )

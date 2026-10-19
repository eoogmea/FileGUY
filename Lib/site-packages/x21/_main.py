from __future__ import annotations

import base64

from . import _x21


def __dex_22b__(scope: dict, iv: str, smessage: bytes) -> None:
    _x21.decrypt_and_exec_22b(smessage, iv, scope)


def __dex_22c__(scope: dict, iv_smessage: str) -> None:
    _x21.decrypt_and_exec_22c(base64.a85decode(iv_smessage), scope)


def __dex_23a__(scope: dict, iv_smessage_tag: str) -> None:
    _x21.decrypt_and_exec_23a(base64.a85decode(iv_smessage_tag), scope)


def __dex_23b__(scope: dict, data: bytes) -> None:
    # 23b is just like 23a encryption except that byte-strings are exchanged
    _x21.decrypt_and_exec_23a(data, scope)

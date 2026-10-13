import json
import math
import os
from pathlib import Path

from dotenv import load_dotenv


PROJECT_ROOT = Path(__file__).resolve().parent.parent
load_dotenv(PROJECT_ROOT / ".env", encoding="utf-8", override=False)


def env_text(name: str, *, allow_empty: bool = False, **fields: str) -> str:
    value = os.environ.get(name)
    if value is None or (not allow_empty and not value.strip()):
        raise ValueError(f"Configure {name} no .env ou nas variaveis de ambiente.")
    value = value.strip()
    if fields:
        try:
            value = value.format(**fields)
        except (KeyError, ValueError, IndexError, AttributeError) as error:
            raise ValueError(f"Modelo invalido em {name}. Campos: {', '.join(fields)}.") from error
    return value


def env_choice(name: str, choices: tuple[str, ...]) -> str:
    value = env_text(name)
    if value not in choices:
        raise ValueError(f"{name} deve ser um destes valores: {', '.join(choices)}.")
    return value


def env_mapping(name: str) -> dict[str, str]:
    try:
        value = json.loads(env_text(name))
    except json.JSONDecodeError as error:
        raise ValueError(f"{name} deve conter um objeto JSON valido.") from error
    if not isinstance(value, dict) or not value or any(
        not isinstance(key, str) or not key.strip()
        or not isinstance(item, str) or not item.strip()
        for key, item in value.items()
    ):
        raise ValueError(f"{name} deve mapear nomes nao vazios para textos nao vazios.")
    return value


def env_seconds(name: str) -> float:
    try:
        value = float(env_text(name))
    except ValueError as error:
        raise ValueError(f"{name} deve ser um numero de segundos maior ou igual a zero.") from error
    if not math.isfinite(value) or value < 0:
        raise ValueError(f"{name} deve ser um numero de segundos finito maior ou igual a zero.")
    return value


def env_path(name: str, **fields: str) -> str:
    value = env_text(name, **fields)
    path = Path(value)
    if not path.is_absolute():
        path = PROJECT_ROOT / path
    return str(path.resolve())


def mapped_file(name: str, site: str) -> str:
    mapping = env_mapping(name)
    if site not in mapping:
        raise ValueError(f"{name} precisa conter um arquivo para o site {site}.")
    return mapping[site]

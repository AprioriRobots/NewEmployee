import os
from pathlib import Path

import requests


BITRIX_TEMPLATE_TOKEN = os.getenv("BITRIX_TEMPLATE_TOKEN") or os.getenv("BITRIX_TOKEN") or ""
BITRIX_TEMPLATE_USER_ID = str(os.getenv("BITRIX_TEMPLATE_USER_ID") or "69")
BITRIX_TEMPLATE_ROOT_FOLDER_ID = int(os.getenv("BITRIX_TEMPLATE_ROOT_FOLDER_ID") or "49464")

TEMPLATE_DIRECTORIES = [
    Path("employee_instructions/input_samples"),
    Path("formul_tkrf/input_sample"),
    Path("lna/input_sample"),
    Path("mat_response/input_sample"),
    Path("military_registr/input_samples"),
    Path("nda/input_sample"),
    Path("occup_safety/input_sample"),
    Path("personal_secret/input_sample"),
    Path("tk_rf/input_sample"),
    Path("zp_alignment/input_sample"),
]


def _bitrix_method_url(method: str) -> str:
    return f"https://b24.apriori.rest/rest/{BITRIX_TEMPLATE_USER_ID}/{BITRIX_TEMPLATE_TOKEN}/{method}"


def _call_bitrix(method: str, payload: dict) -> dict:
    response = requests.post(_bitrix_method_url(method), json=payload, timeout=60)
    response.raise_for_status()
    data = response.json()
    if "error" in data:
        raise RuntimeError(f"Bitrix API error for {method}: {data['error']} {data.get('error_description', '')}".strip())
    return data


def _get_folder_children(folder_id: int) -> list[dict]:
    data = _call_bitrix("disk.folder.getchildren", {"id": int(folder_id)})
    return data.get("result") or []


def _resolve_remote_folder_id(relative_dir: Path) -> int:
    current_folder_id = BITRIX_TEMPLATE_ROOT_FOLDER_ID
    for part in relative_dir.parts:
        children = _get_folder_children(current_folder_id)
        next_item = next((item for item in children if item.get("NAME") == part), None)
        if next_item is None:
            raise FileNotFoundError(f"Не найдена папка шаблонов в Bitrix: {'/'.join(relative_dir.parts)}")
        current_folder_id = int(next_item["ID"])
    return current_folder_id


def _download_file(download_url: str, local_path: Path) -> None:
    response = requests.get(download_url, timeout=120)
    response.raise_for_status()
    local_path.parent.mkdir(parents=True, exist_ok=True)
    local_path.write_bytes(response.content)


def _remove_stale_local_entries(local_dir: Path, expected_names: set[str]) -> None:
    if not local_dir.exists():
        return

    for local_entry in local_dir.iterdir():
        if local_entry.name in expected_names:
            continue
        if local_entry.is_dir():
            for nested_path in sorted(local_entry.rglob("*"), reverse=True):
                if nested_path.is_file():
                    nested_path.unlink()
                elif nested_path.is_dir():
                    nested_path.rmdir()
            local_entry.rmdir()
        else:
            local_entry.unlink()


def _sync_remote_folder(remote_folder_id: int, local_dir: Path) -> None:
    local_dir.mkdir(parents=True, exist_ok=True)

    remote_items = _get_folder_children(remote_folder_id)
    expected_names = {
        item.get("NAME")
        for item in remote_items
        if item.get("NAME") and not str(item.get("NAME")).startswith("~$")
    }
    _remove_stale_local_entries(local_dir, expected_names)

    for item in remote_items:
        item_name = item.get("NAME") or ""
        if not item_name or item_name.startswith("~$"):
            continue

        download_url = item.get("DOWNLOAD_URL")
        if download_url:
            _download_file(download_url, local_dir / item_name)
            continue

        child_folder_id = item.get("ID")
        if not child_folder_id:
            continue
        _sync_remote_folder(int(child_folder_id), local_dir / item_name)


def sync_templates_from_bitrix() -> None:
    if not BITRIX_TEMPLATE_TOKEN:
        raise RuntimeError("Не задан токен Bitrix для синхронизации шаблонов")

    for relative_dir in TEMPLATE_DIRECTORIES:
        remote_folder_id = _resolve_remote_folder_id(relative_dir)
        local_dir = Path(relative_dir)
        _sync_remote_folder(remote_folder_id, local_dir)

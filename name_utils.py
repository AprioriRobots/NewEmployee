from typing import List, Optional


def split_fio_parts(fio: Optional[str]) -> List[str]:
    return str(fio or "").split()


def get_fio_part(fio: Optional[str], index: int) -> str:
    parts = split_fio_parts(fio)
    return parts[index] if len(parts) > index else ""


def get_name_patronymic_initials(fio: Optional[str]) -> str:
    parts = split_fio_parts(fio)
    return " ".join(part[0] for part in parts[1:] if part)


def get_fio_with_initials(fio: Optional[str]) -> str:
    surname = get_fio_part(fio, 0)
    initials = get_name_patronymic_initials(fio)
    return " ".join(part for part in (surname, initials) if part)

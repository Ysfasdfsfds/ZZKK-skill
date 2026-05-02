from __future__ import annotations

from datetime import datetime


def ensure_run_date(raw: str | None) -> str:
    if raw:
        digits = ''.join(ch for ch in raw if ch.isdigit())
        if len(digits) != 8:
            raise ValueError(f"run_date must be YYYYMMDD, got: {raw}")
        return digits
    return datetime.now().strftime("%Y%m%d")


def version_token(version: str) -> str:
    cleaned = version.strip()
    return cleaned if cleaned.upper().startswith("V") else f"V{cleaned}"


def dotted_date(run_date: str) -> str:
    value = ensure_run_date(run_date)
    return f"{value[0:4]}.{int(value[4:6])}.{int(value[6:8])}"


def iso_date(run_date: str) -> str:
    value = ensure_run_date(run_date)
    return f"{value[0:4]}-{value[4:6]}-{value[6:8]}"


def customer_doc_filename(module_name: str, version: str) -> str:
    return f"{module_name}-客户需求说明书_{version}.docx"


def product_doc_filename(module_name: str, version: str) -> str:
    return f"{module_name}-模块级产品需求分析说明书_{version}.docx"


def customer_review_filename(module_name: str) -> str:
    return f"{module_name}-客户需求说明书_同级评审会议记录表.xlsx"


def product_review_filename(module_name: str) -> str:
    return f"{module_name}-模块级产品需求分析说明书_同级评审会议记录表.xlsx"


def customer_doc_number(module_code: str, version: str) -> str:
    return f"DT-V11 SP1-{module_code}-CRS-{version_token(version)}"


def product_doc_number(module_code: str, version: str) -> str:
    return f"DT-V11 SP1-{module_code}-PRD-{version_token(version)}"


def customer_review_record_number(module_code: str, run_date: str) -> str:
    return f"MT-PR-A-{module_code}-CRS-{ensure_run_date(run_date)}"


def product_review_record_number(module_code: str, run_date: str) -> str:
    return f"MT-PR-A-{module_code}-PRD-{ensure_run_date(run_date)}"

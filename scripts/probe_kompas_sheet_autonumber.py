"""
Диагностика отключения автонумерации листов в КОМПАС-3D (API7).

Запуск (из корня репозитория):
  python scripts/probe_kompas_sheet_autonumber.py "C:\\path\\to\\file.cdw"
  python scripts/probe_kompas_sheet_autonumber.py "file.cdw" --retries 3 --delay 1.5

Скрипт открывает чертёж, по очереди пробует способы записи/чтения SheetAutoNumber,
вызывает ту же логику, что и stamp_updater._ensure_sheet_auto_number_disabled,
при необходимости повторяет цикл (переподключение COM между попытками).
"""

from __future__ import annotations

import argparse
import logging
import sys
import time
from pathlib import Path

# корень репозитория: .../nordfox_module
_ROOT = Path(__file__).resolve().parents[1]
if str(_ROOT) not in sys.path:
    sys.path.insert(0, str(_ROOT))

from src.core.kompas_connector import KompasConnector  # noqa: E402
from src.core import stamp_updater as su  # noqa: E402


def _probe_raw_strategies(ds: object) -> None:
    """Печатает, какой из способов записи сработал (как в stamp_updater)."""
    print("  --- зонд SheetAutoNumber (запись) ---")
    p = getattr(ds, "_prop_map_put_", None) or getattr(type(ds), "_prop_map_put_", {})
    if "SheetAutoNumber" in p:
        print("  [ok] _prop_map_put_['SheetAutoNumber']")
    else:
        print("  [?] нет ключа SheetAutoNumber в _prop_map_put_")

    for label, val in (("VARIANT_BOOL 0", 0), ("Python False", False)):
        try:
            ds.SheetAutoNumber = val
            print(f"  [ok] SheetAutoNumber = {val!r} ({label})")
            break
        except Exception as exc:
            print(f"  [fail] SheetAutoNumber = {val!r} ({label}) -> {exc}")

    try:
        if "SheetAutoNumber" in p:
            args, tail = p["SheetAutoNumber"]
            ds._oleobj_.Invoke(*(args + (0,) + tail))
            print("  [ok] PyIDispatch.Invoke(..., 0) из _prop_map_put_")
    except Exception as exc:
        print(f"  [fail] Invoke + int 0 -> {exc}")

    print("  --- зонд чтение ---")
    try:
        v = ds.SheetAutoNumber
        print(f"  [ok] чтение SheetAutoNumber -> {v!r}")
    except Exception as exc:
        print(f"  [fail] чтение SheetAutoNumber -> {exc}")


def run_probe(cdw: Path, retries: int, delay: float) -> int:
    logging.basicConfig(level=logging.INFO, format="%(levelname)s %(message)s")
    su._sheet_autonumber_put_strategy = None
    cdw = cdw.resolve()
    if not cdw.is_file():
        print(f"Файл не найден: {cdw}", file=sys.stderr)
        return 2

    conn = KompasConnector()
    for attempt in range(1, retries + 1):
        print(f"\n=== Попытка {attempt}/{retries} ===")
        if not conn.connect(force_reconnect=(attempt > 1)):
            print("[fail] не удалось подключиться к КОМПАС-3D")
            time.sleep(delay)
            continue

        api7 = conn.get_api7()
        if api7 is None:
            print("[fail] API7 недоступен")
            time.sleep(delay)
            continue

        doc = None
        try:
            doc = api7.Documents.Open(str(cdw), False, False)
            if not doc:
                print("[fail] Documents.Open вернул пусто")
                continue
            time.sleep(0.5)
            active = api7.ActiveDocument
            if not active:
                print("[fail] ActiveDocument пуст")
                continue

            ds, err = su._drawing_document_settings(active)
            if ds is None:
                print(f"[fail] _drawing_document_settings: {err}")
                continue

            _probe_raw_strategies(ds)

            ok, msg = su._ensure_sheet_auto_number_disabled(active, context=cdw.name)
            if ok:
                print(f"[ok] _ensure_sheet_auto_number_disabled: успех")
                return 0
            print(f"[fail] _ensure_sheet_auto_number_disabled: {msg}")
        finally:
            if doc is not None:
                try:
                    doc.Close(False)
                except Exception as exc:
                    print(f"[warn] Close: {exc}")
            time.sleep(delay)

    return 1


def main() -> None:
    ap = argparse.ArgumentParser(description="Зонд SheetAutoNumber для КОМПАС API7")
    ap.add_argument("cdw", type=Path, help="Путь к файлу .cdw")
    ap.add_argument("--retries", type=int, default=3, help="Число циклов подключить→открыть→зонд")
    ap.add_argument("--delay", type=float, default=1.0, help="Пауза между попытками, с")
    args = ap.parse_args()
    raise SystemExit(run_probe(args.cdw, args.retries, args.delay))


if __name__ == "__main__":
    main()

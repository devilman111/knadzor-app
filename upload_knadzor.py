"""
upload_knadzor.py
Загружает ссылки из CSV в кибернадзор knadzor.kz
Папка: C:\Users\Пользователь\Desktop\knadzor_help_app\
"""

import os
import re
import json
import traceback
from pathlib import Path

import pandas as pd
from playwright.sync_api import sync_playwright

# ═══════════════════════════════════════════════════════════
# НАСТРОЙКИ — все файлы в одной папке knadzor_help_app
# ═══════════════════════════════════════════════════════════

APP_DIR     = Path(r"C:\Users\Пользователь\Desktop\knadzor_help_app")

BASE_URL    = "https://knadzor.kz/#/links"
AUTH_FILE   = APP_DIR / "auth.json"
CSV_FILE    = APP_DIR / "links_100_updated.csv"
SCREENS_DIR = APP_DIR / "screens"
DB_FILE     = APP_DIR / "uploaded_urls.json"

ROW_START   = 0
ROW_END     = 3000

HEADLESS    = False
SLOW_MO     = 150

# ═══════════════════════════════════════════════════════════


def load_db() -> set:
    if DB_FILE.exists():
        with open(DB_FILE, encoding="utf-8") as f:
            return set(json.load(f))
    return set()


def save_db(db: set):
    with open(DB_FILE, "w", encoding="utf-8") as f:
        json.dump(list(db), f, ensure_ascii=False)
    print(f"  [DB] Сохранено: {len(db)} URL")


def clean(value):
    if pd.isna(value):
        return ""
    return str(value).strip()


def is_page_alive(page):
    try:
        page.title()
        return True
    except Exception:
        return False


def reopen_page(context):
    page = context.new_page()
    page.goto(BASE_URL, wait_until="domcontentloaded")
    page.wait_for_timeout(2000)
    return page


def open_add_link_modal(page):
    page.locator("button", has_text="Добавить ссылку").first.click()
    page.locator("edit-modal .modal-box--small").wait_for(state="visible", timeout=15000)
    page.wait_for_timeout(300)
    return page.locator("edit-modal .modal-box--small").first


def fill_text_fields(modal, row):
    modal.locator('my-input[formcontrolname="title"] input').fill(clean(row["title"]))
    modal.locator('my-input[formcontrolname="url"] input').fill(clean(row["url"]))
    modal.locator('my-textarea[formcontrolname="comment"] textarea').fill(clean(row["reason"]))


def set_court_decision(modal, value: bool):
    checkbox = modal.locator('input[formcontrolname="court_decision"]')
    if value and not checkbox.is_checked():
        checkbox.check()
    elif not value and checkbox.is_checked():
        checkbox.uncheck()


def select_by_formcontrol(modal, formcontrolname: str, visible_text: str):
    if not visible_text:
        return
    select_el = modal.locator(
        f'my-select-field[formcontrolname="{formcontrolname}"] select'
    ).first
    select_el.wait_for(state="visible", timeout=10000)
    options = [x.strip() for x in select_el.locator("option").all_text_contents()]
    if visible_text not in options:
        raise ValueError(f'"{visible_text}" не найдено в {formcontrolname}. Доступно: {options}')
    select_el.select_option(label=visible_text)
    modal.page.wait_for_timeout(100)


def set_group(modal, group_value: str):
    group_value = clean(group_value).lower().strip()
    if not group_value:
        return
    css_map = {
        "красный":   "attribute--cat-red",
        "оранжевый": "attribute--cat-orange",
    }
    css_class = css_map.get(group_value)
    if not css_class:
        print(f"  [WARN] Неизвестная группа '{group_value}'")
        return
    try:
        group_select = modal.locator("group-select").first
        group_select.wait_for(state="visible", timeout=8000)
        attr_div = group_select.locator(f"div.{css_class}").first
        attr_div.wait_for(state="visible", timeout=5000)
        panel = attr_div.locator("xpath=../../..").first
        btn = panel.locator("div.panel__layout.shrink button").first
        btn.wait_for(state="visible", timeout=5000)
        btn.click()
        modal.page.wait_for_timeout(200)
        print(f"  [GROUP] {group_value}")
    except Exception as e:
        print(f"  [WARN] set_group: {e}")


def fill_optional_counts(modal, posts_count: str, members_count: str):
    if posts_count:
        try:
            modal.locator('my-input[formcontrolname="posts_count"] input').fill(posts_count)
        except Exception:
            pass
    if members_count:
        try:
            modal.locator('my-input[formcontrolname="members_count"] input').fill(members_count)
        except Exception:
            pass


def resolve_screenshot_path(row):
    title = clean(row.get("title", ""))
    match = re.search(r"(\d+)$", title)
    if not match:
        raise ValueError(f"Не удалось вытащить номер из title: {title!r}")
    num = match.group(1)
    candidates = [
        SCREENS_DIR / f"{num}.png",
        SCREENS_DIR / f"{num}.jpg",
        SCREENS_DIR / f"{num}.jpeg",
    ]
    for c in candidates:
        if c.exists():
            return c.resolve()
    raise FileNotFoundError(
        f"Нет скриншота для {title!r}. Ожидался: "
        + ", ".join(str(x) for x in candidates)
    )


def attach_screenshot(modal, row):
    file_path = resolve_screenshot_path(row)
    print(f"  [FILE] {file_path.name}")
    file_input = modal.locator('input.upload-file[type="file"]').first
    file_input.wait_for(state="attached", timeout=10000)
    file_input.set_input_files(str(file_path))
    modal.page.wait_for_timeout(1000)
    print("  [FILE OK]")


def check_after_save(page, modal) -> str:
    for _ in range(40):
        page.wait_for_timeout(100)
        try:
            body = page.inner_text("body")
            if any(t in body.lower() for t in ["уже внесен", "уже внесён", "already exists"]):
                try:
                    page.keyboard.press("Escape")
                    page.wait_for_timeout(400)
                except Exception:
                    pass
                return "duplicate"
        except Exception:
            pass
        try:
            if not modal.is_visible():
                return "ok"
        except Exception:
            return "ok"
    try:
        page.keyboard.press("Escape")
        page.wait_for_timeout(400)
    except Exception:
        pass
    return "timeout"


def upload_one(page, row) -> str:
    modal = open_add_link_modal(page)
    fill_text_fields(modal, row)

    court = clean(row.get("court_decision", "")).lower() in ["1", "true", "yes", "да"]
    set_court_decision(modal, court)

    select_by_formcontrol(modal, "resource_type", clean(row["object_type"]))
    select_by_formcontrol(modal, "resource",      clean(row["resource_type"]))
    select_by_formcontrol(modal, "offense",       clean(row["violation"]))
    select_by_formcontrol(modal, "language",      clean(row["language"]))

    group_value = clean(row.get("group", ""))
    if group_value:
        set_group(modal, group_value)

    fill_optional_counts(
        modal,
        clean(row.get("publications", "")),
        clean(row.get("subscribers",  "")),
    )

    attach_screenshot(modal, row)

    modal.page.wait_for_timeout(300)
    modal.locator("button.button--primary").first.click()

    return check_after_save(page, modal)


def main():
    # Проверки
    if not AUTH_FILE.exists():
        raise FileNotFoundError(f"Не найден: {AUTH_FILE}")
    if not Path(CSV_FILE).exists():
        raise FileNotFoundError(f"Не найден CSV: {CSV_FILE}")
    if not SCREENS_DIR.exists():
        raise FileNotFoundError(f"Не найдена папка скриншотов: {SCREENS_DIR}")

    df_full = pd.read_csv(CSV_FILE, sep=";", encoding="utf-8-sig")
    df = df_full.iloc[ROW_START:ROW_END].reset_index(drop=True)
    total = len(df)

    # Загружаем базу
    db = load_db()

    print("=" * 60)
    print(f"  Knadzor Upload Tool")
    print(f"  Папка: {APP_DIR}")
    print(f"  CSV:   {Path(CSV_FILE).name} ({total} строк)")
    print(f"  База:  {len(db)} URL уже загружено ранее")
    print("=" * 60)

    results   = []
    ok_count  = 0
    dup_count = 0
    skip_count = 0
    err_count = 0

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=HEADLESS, slow_mo=SLOW_MO)
        context = browser.new_context(storage_state=str(AUTH_FILE))
        page = reopen_page(context)

        for idx, row in df.iterrows():
            row_num = idx + 1
            url     = clean(row.get("url", ""))

            print(f"\n[{row_num}/{total}] {url[:75]}")

            # Проверяем базу ДО загрузки
            if url and url in db:
                print(f"  [SKIP] Уже в базе — пропускаю")
                skip_count += 1
                results.append({"row": row_num, "url": url, "status": "IN_DB"})
                continue

            if not is_page_alive(page):
                print("  [WARN] Страница закрыта, переоткрываю...")
                try:
                    page = reopen_page(context)
                except Exception as e:
                    print(f"  [FATAL] {e}")
                    results.append({"row": row_num, "url": url, "status": "SKIP"})
                    continue

            try:
                result = upload_one(page, row)

                if result == "ok":
                    ok_count += 1
                    db.add(url)
                    print("  [OK] ✅ Загружено")
                    results.append({"row": row_num, "url": url, "status": "OK"})

                elif result == "duplicate":
                    dup_count += 1
                    db.add(url)
                    print("  [DUP] ⏭️ Дубликат — запоминаю")
                    results.append({"row": row_num, "url": url, "status": "DUPLICATE"})

                else:
                    print(f"  [?] {result}")
                    results.append({"row": row_num, "url": url, "status": result})

            except Exception as e:
                err_str = str(e)
                err_count += 1
                print(f"  [ERROR] {err_str}")
                traceback.print_exc()
                results.append({"row": row_num, "url": url, "status": "ERROR", "error": err_str})

                if "closed" in err_str.lower() or "Target page" in err_str:
                    try:
                        page = reopen_page(context)
                    except Exception as re_err:
                        print(f"  [FATAL] {re_err}")
                        break
                else:
                    try:
                        page.keyboard.press("Escape")
                        page.wait_for_timeout(800)
                    except Exception:
                        pass

            # Сохраняем базу каждые 10 загрузок
            if (ok_count + dup_count) > 0 and (ok_count + dup_count) % 10 == 0:
                save_db(db)

        # Финальное сохранение
        save_db(db)

        results_path = APP_DIR / "upload_results.csv"
        pd.DataFrame(results).to_csv(results_path, index=False, encoding="utf-8-sig")

        print()
        print("=" * 60)
        print(f"  ГОТОВО!")
        print(f"  ✅ Загружено:       {ok_count}")
        print(f"  ⏭️  Пропущено (БД):  {skip_count}")
        print(f"  🔁 Дубликатов:     {dup_count}")
        print(f"  ❌ Ошибок:          {err_count}")
        print(f"  💾 Всего в базе:    {len(db)} URL")
        print(f"  📄 Лог:             {results_path}")
        print("=" * 60)

        try:
            context.storage_state(path=str(AUTH_FILE))
        except Exception:
            pass
        try:
            browser.close()
        except Exception:
            pass


if __name__ == "__main__":
    main()
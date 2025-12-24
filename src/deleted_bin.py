import sys

from dotenv import load_dotenv
from playwright.sync_api import sync_playwright

from outlook_common import append_excel_summary, delete_many, load_config, log, login


def run(*, headless: bool = False, browser_name: str | None = None, timeout_ms: int | None = None) -> None:
    load_dotenv()
    log("Run: start deleted_bin")
    cfg = load_config(browser_name=browser_name, headless=headless, timeout_ms=timeout_ms)

    with sync_playwright() as p:
        total_deleted = 0
        last_excel_total = 0
        browser_restarts = 0
        max_browser_restarts = 10

        def flush_excel_partial() -> None:
            nonlocal last_excel_total
            if total_deleted > last_excel_total:
                append_excel_summary(
                    script_name="deleted_bin",
                    list_name="Deleted",
                    deleted_this_session=(total_deleted - last_excel_total),
                    total_deleted=total_deleted,
                    browser_name=cfg.browser_name,
                    headless=cfg.headless,
                    email=cfg.email,
                )
                last_excel_total = total_deleted

        while browser_restarts <= max_browser_restarts:
            browser = None
            context = None
            try:
                log(f"Run: launch browser={cfg.browser_name} headless={cfg.headless} restart={browser_restarts}")
                browser_type = getattr(p, cfg.browser_name)
                browser = browser_type.launch(headless=cfg.headless)

                context = browser.new_context()
                page = context.new_page()
                login(page, cfg.email, cfg.password, timeout_ms=cfg.timeout_ms)

                def on_deleted() -> None:
                    nonlocal total_deleted, last_excel_total
                    total_deleted += 1
                    if total_deleted - last_excel_total >= 5:
                        flush_excel_partial()

                deleted_this = delete_many(
                    page,
                    list_name="Deleted",
                    timeout_ms=cfg.timeout_ms,
                    batch_size=5,
                    on_deleted=on_deleted,
                )
                log(f"Run: deleted_this_session={deleted_this} total_deleted={total_deleted}")

                flush_excel_partial()

                if deleted_this == 0:
                    log("Run: nothing left to delete; stopping")
                    break

                # Continue in the same browser/session if it's healthy.
                # (delete_many only raises when UI is stuck)
                continue

            except KeyboardInterrupt:
                log("Run: interrupted by user")
                flush_excel_partial()
                break
            except Exception as exc:
                browser_restarts += 1
                log(f"Run: restarting browser due to error ({type(exc).__name__}: {exc})")
                flush_excel_partial()
                continue
            finally:
                if context:
                    try:
                        context.close()
                    except Exception:
                        pass
                if browser:
                    try:
                        browser.close()
                    except Exception:
                        pass

        log("Run: finished")


if __name__ == "__main__":
    browser = None
    headless = False
    timeout_ms = None

    if "--browser" in sys.argv:
        idx = sys.argv.index("--browser")
        if idx + 1 < len(sys.argv):
            browser = sys.argv[idx + 1]

    if "--headless" in sys.argv:
        headless = True

    if "--timeout-ms" in sys.argv:
        idx = sys.argv.index("--timeout-ms")
        if idx + 1 < len(sys.argv):
            timeout_ms = int(sys.argv[idx + 1])

    run(headless=headless, browser_name=browser, timeout_ms=timeout_ms)

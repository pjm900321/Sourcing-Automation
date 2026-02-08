import argparse
import os
import random
import time
from dataclasses import dataclass
from typing import List, Optional

import pandas as pd
from playwright.sync_api import Browser, BrowserContext, sync_playwright


@dataclass
class SearchResult:
    keyword: str
    organic_count: int
    rocket_count: int
    rocket_ratio: float
    verdict: str
    error: str


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="쿠팡 키워드 로켓배송 비율 분석")
    parser.add_argument(
        "--input",
        dest="input_file",
        default=None,
        help="input 폴더 내 xlsx 파일 경로 (미지정 시 첫 번째 파일 사용)",
    )
    parser.add_argument(
        "--test",
        action="store_true",
        help="테스트 모드: 첫 번째 키워드만 처리하고 디버그 파일 저장",
    )
    return parser.parse_args()


def ensure_directories() -> None:
    os.makedirs("input", exist_ok=True)
    os.makedirs("output", exist_ok=True)


def find_input_file(input_file: Optional[str]) -> str:
    if input_file:
        return input_file
    candidates = [name for name in os.listdir("input") if name.lower().endswith(".xlsx")]
    if not candidates:
        raise FileNotFoundError("input 폴더에 xlsx 파일이 없습니다.")
    return os.path.join("input", sorted(candidates)[0])


def load_keywords(input_path: str) -> List[str]:
    data = pd.read_excel(input_path)
    if "키워드" not in data.columns:
        raise ValueError("엑셀에 '키워드' 컬럼이 없습니다.")
    keywords = (
        data["키워드"]
        .dropna()
        .astype(str)
        .map(str.strip)
        .loc[lambda series: series != ""]
        .tolist()
    )
    return keywords


def load_existing_results(output_path: str) -> pd.DataFrame:
    if not os.path.exists(output_path):
        return pd.DataFrame(columns=[
            "키워드",
            "상위10개중개수",
            "로켓배송개수",
            "로켓비율",
            "판정",
            "오류",
        ])
    return pd.read_excel(output_path, sheet_name="main")


def save_results(output_path: str, results: pd.DataFrame) -> None:
    pass_only = results.loc[results["판정"] == "PASS"].copy()
    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        results.to_excel(writer, index=False, sheet_name="main")
        pass_only.to_excel(writer, index=False, sheet_name="pass_only")


def build_search_url(keyword: str) -> str:
    return f"https://www.coupang.com/np/search?q={keyword}&channel=user"


def build_random_viewport() -> dict:
    return {
        "width": random.randint(1280, 1920),
        "height": random.randint(720, 1080),
    }


def build_random_user_agent() -> str:
    user_agents = [
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/126.0.0.0 Safari/537.36",
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/125.0.0.0 Safari/537.36",
        "Mozilla/5.0 (Macintosh; Intel Mac OS X 13_5) AppleWebKit/605.1.15 (KHTML, like Gecko) Version/17.0 Safari/605.1.15",
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36",
    ]
    return random.choice(user_agents)


def prepare_context(browser: Browser) -> BrowserContext:
    context = browser.new_context(
        user_agent=build_random_user_agent(),
        viewport=build_random_viewport(),
        locale="ko-KR",
    )
    context.add_init_script(
        "Object.defineProperty(navigator, 'webdriver', {get: () => undefined});"
    )
    return context


def human_like_scroll(page) -> None:
    scroll_times = random.randint(3, 6)
    for _ in range(scroll_times):
        distance = random.randint(300, 700)
        page.mouse.wheel(0, distance)
        time.sleep(random.uniform(0.6, 1.6))


def human_like_mouse(page) -> None:
    moves = random.randint(4, 8)
    for _ in range(moves):
        x = random.randint(100, 900)
        y = random.randint(100, 700)
        page.mouse.move(x, y, steps=random.randint(5, 20))
        time.sleep(random.uniform(0.2, 0.8))


def extract_rank_text(item) -> str:
    rank_selectors = [
        ".search-product__rank",
        ".search-product__rank-text",
        ".rank-badge",
        ".product-rank",
    ]
    for selector in rank_selectors:
        locator = item.locator(selector)
        if locator.count() > 0:
            text = locator.first.inner_text().strip()
            if text:
                return text
    return ""


def parse_rank_number(rank_text: str) -> Optional[int]:
    if not rank_text:
        return None
    digits = "".join(char for char in rank_text if char.isdigit())
    if not digits:
        return None
    rank = int(digits)
    return rank if 1 <= rank <= 10 else None


def detect_ad(item) -> bool:
    ad_selectors = [
        ".search-product__ad-badge",
        ".search-product__ad-badge-text",
        "span.ad-badge-text",
        "span.ad-badge",
        "span[aria-label='광고']",
        "img[alt='광고']",
        ".ad-badge",
    ]
    for selector in ad_selectors:
        if item.locator(selector).count() > 0:
            return True
    item_text = item.inner_text()
    return "AD" in item_text or "광고" in item_text


def extract_product_name(item) -> str:
    name_selectors = [
        ".name",
        ".search-product__name",
        "a.search-product-link",
    ]
    for selector in name_selectors:
        locator = item.locator(selector)
        if locator.count() > 0:
            text = locator.first.inner_text().strip()
            if text:
                return text
    return ""


def detect_rocket_badge(item) -> str:
    badge_texts = []
    badge_locators = item.locator("img[alt], span, em, i")
    for index in range(min(20, badge_locators.count())):
        element = badge_locators.nth(index)
        alt_text = element.get_attribute("alt") or ""
        text = element.inner_text().strip()
        combined = f"{alt_text} {text}".strip()
        if combined:
            badge_texts.append(combined)
    full_text = " ".join(badge_texts)
    if "판매자로켓" in full_text:
        return "판매자로켓"
    if "로켓" in full_text:
        return "로켓배송"
    return "뱃지없음"


def analyze_keyword(browser: Browser, keyword: str, test_mode: bool) -> SearchResult:
    organic_count = 0
    rocket_count = 0
    error_message = ""
    debug_items = []

    context = prepare_context(browser)
    page = context.new_page()
    try:
        page.goto(build_search_url(keyword), wait_until="domcontentloaded", timeout=60000)
        page.wait_for_selector("li.search-product", timeout=30000)
        time.sleep(random.uniform(1.0, 2.2))
        human_like_scroll(page)
        human_like_mouse(page)

        items = page.locator("li.search-product")
        total_items = items.count()
        ranked_items = {}

        for index in range(total_items):
            item = items.nth(index)
            rank_text = extract_rank_text(item)
            rank_number = parse_rank_number(rank_text)
            is_ad = detect_ad(item)
            badge_type = detect_rocket_badge(item)
            product_name = extract_product_name(item)

            debug_items.append(
                {
                    "rank_text": rank_text,
                    "rank_number": rank_number,
                    "product_name": product_name,
                    "is_ad": is_ad,
                    "badge_type": badge_type,
                }
            )

            if is_ad or rank_number is None:
                continue
            if rank_number in ranked_items:
                continue
            if 1 <= rank_number <= 10:
                ranked_items[rank_number] = badge_type

        for rank in range(1, 11):
            if rank not in ranked_items:
                continue
            organic_count += 1
            if ranked_items[rank] == "로켓배송":
                rocket_count += 1

        if test_mode:
            with open(os.path.join("output", "debug.html"), "w", encoding="utf-8") as file:
                file.write(page.content())
            page.screenshot(path=os.path.join("output", "debug.png"), full_page=True)
            for item in debug_items:
                rank_label = item["rank_text"] or "순위없음"
                name_label = item["product_name"] or "상품명없음"
                ad_label = "광고" if item["is_ad"] else "자연노출"
                badge_label = item["badge_type"]
                print(f"[TEST] {rank_label} | {name_label} | {ad_label} | {badge_label}")
    except Exception as exc:
        error_message = str(exc)
    finally:
        context.close()

    ratio = rocket_count / organic_count if organic_count else 0.0
    verdict = "PASS" if rocket_count <= 5 else "FAIL"
    if error_message:
        verdict = "FAIL"
    return SearchResult(
        keyword=keyword,
        organic_count=organic_count,
        rocket_count=rocket_count,
        rocket_ratio=ratio,
        verdict=verdict,
        error=error_message,
    )


def main() -> None:
    args = parse_args()
    ensure_directories()

    input_path = find_input_file(args.input_file)
    keywords = load_keywords(input_path)

    output_path = os.path.join("output", "results.xlsx")
    existing_results = load_existing_results(output_path)
    completed = set(existing_results["키워드"].astype(str).tolist())

    all_results = existing_results.copy()

    if args.test and keywords:
        keywords = keywords[:1]

    with sync_playwright() as playwright:
        browser = playwright.chromium.launch(headless=False)
        processed_since_restart = 0
        restart_after = random.randint(10, 15)

        for index, keyword in enumerate(keywords, start=1):
            if keyword in completed:
                continue
            if processed_since_restart >= restart_after:
                browser.close()
                browser = playwright.chromium.launch(headless=False)
                processed_since_restart = 0
                restart_after = random.randint(10, 15)

            last_error = ""
            result: Optional[SearchResult] = None

            for attempt in range(1, 4):
                try:
                    result = analyze_keyword(browser, keyword, args.test)
                    last_error = result.error
                    if not last_error:
                        break
                except Exception as exc:
                    last_error = str(exc)
                time.sleep(2)

            if result is None:
                result = SearchResult(
                    keyword=keyword,
                    organic_count=0,
                    rocket_count=0,
                    rocket_ratio=0.0,
                    verdict="FAIL",
                    error=last_error,
                )
            elif last_error:
                result = SearchResult(
                    keyword=keyword,
                    organic_count=result.organic_count,
                    rocket_count=result.rocket_count,
                    rocket_ratio=result.rocket_ratio,
                    verdict="FAIL",
                    error=last_error,
                )

            new_row = {
                "키워드": result.keyword,
                "상위10개중개수": result.organic_count,
                "로켓배송개수": result.rocket_count,
                "로켓비율": round(result.rocket_ratio, 4),
                "판정": result.verdict,
                "오류": result.error,
            }
            all_results = pd.concat([all_results, pd.DataFrame([new_row])], ignore_index=True)
            save_results(output_path, all_results)

            print(
                f"[{index}/{len(keywords)}] {keyword} 검색 중... "
                f"로켓배송 {result.rocket_count}개/{result.organic_count}개 → {result.verdict}"
            )

            processed_since_restart += 1
            time.sleep(random.uniform(5, 12))


if __name__ == "__main__":
    main()

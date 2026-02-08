import argparse
import os
import random
import shutil
import subprocess
import tempfile
import time
from dataclasses import dataclass
from typing import List, Optional

import pandas as pd
from selenium import webdriver
from selenium.common.exceptions import NoSuchElementException, TimeoutException
from selenium.webdriver import ActionChains
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support import expected_conditions as ec
from selenium.webdriver.support.ui import WebDriverWait


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
    keywords = data["키워드"]
    keywords = keywords.dropna().astype(str).map(str.strip)
    keywords = keywords.loc[lambda series: series != ""].tolist()
    return keywords


def load_existing_results(output_path: str) -> pd.DataFrame:
    if not os.path.exists(output_path):
        return pd.DataFrame(
            columns=[
                "키워드",
                "상위10개중개수",
                "로켓배송개수",
                "로켓비율",
                "판정",
                "오류",
            ]
        )
    else:
        return pd.read_excel(output_path, sheet_name="main")


def save_results(output_path: str, results: pd.DataFrame) -> None:
    pass_only = results.loc[results["판정"] == "PASS"].copy()
    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        results.to_excel(writer, index=False, sheet_name="main")
        pass_only.to_excel(writer, index=False, sheet_name="pass_only")


def build_search_url(keyword: str) -> str:
    return f"https://www.coupang.com/np/search?q={keyword}&channel=user"


def build_random_viewport() -> tuple[int, int]:
    return random.randint(1280, 1920), random.randint(720, 1080)


def build_random_user_agent() -> str:
    user_agents = [
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/126.0.0.0 Safari/537.36",
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/125.0.0.0 Safari/537.36",
        "Mozilla/5.0 (Macintosh; Intel Mac OS X 13_5) AppleWebKit/605.1.15 (KHTML, like Gecko) Version/17.0 Safari/605.1.15",
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36",
    ]
    return random.choice(user_agents)


def find_chrome_path() -> str:
    candidates = [
        r"C:\Program Files\Google\Chrome\Application\chrome.exe",
        r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe",
    ]
    for path in candidates:
        if os.path.exists(path):
            return path
    raise FileNotFoundError("Chrome 경로를 찾을 수 없습니다. Chrome 설치 여부를 확인하세요.")


def create_driver() -> tuple[webdriver.Chrome, str, subprocess.Popen]:
    profile_dir = tempfile.mkdtemp(prefix="coupang_profile_")
    width, height = build_random_viewport()
    chrome_path = find_chrome_path()

    debug_port = 9222
    creation_flags = getattr(subprocess, "CREATE_NO_WINDOW", 0) if os.name == "nt" else 0
    process = subprocess.Popen(
        [
            chrome_path,
            f"--remote-debugging-port={debug_port}",
            f'--user-data-dir={profile_dir}',
            f"--window-size={width},{height}",
            "--no-first-run",
            "--no-default-browser-check",
            "--disable-popup-blocking",
        ],
        stdout=subprocess.DEVNULL,
        stderr=subprocess.DEVNULL,
        creationflags=creation_flags,
    )
    time.sleep(random.uniform(1.5, 3.0))

    options = webdriver.ChromeOptions()
    options.add_experimental_option("debuggerAddress", f"127.0.0.1:{debug_port}")
    options.add_argument(f"--user-agent={build_random_user_agent()}")

    driver = webdriver.Chrome(options=options)
    driver.set_page_load_timeout(60)
    return driver, profile_dir, process


def cleanup_driver(driver: webdriver.Chrome, profile_dir: str, process: subprocess.Popen) -> None:
    try:
        driver.quit()
    finally:
        process.terminate()
        try:
            process.wait(timeout=10)
        except Exception:
            process.kill()
        shutil.rmtree(profile_dir, ignore_errors=True)


def human_like_scroll(driver: webdriver.Chrome) -> None:
    scroll_times = random.randint(3, 6)
    for _ in range(scroll_times):
        distance = random.randint(300, 700)
        driver.execute_script("window.scrollBy(0, arguments[0]);", distance)
        time.sleep(random.uniform(0.6, 1.6))


def human_like_mouse(driver: webdriver.Chrome) -> None:
    try:
        action = ActionChains(driver)
        moves = random.randint(4, 8)
        for _ in range(moves):
            x_offset = random.randint(5, 80)
            y_offset = random.randint(5, 80)
            action.move_by_offset(x_offset, y_offset).perform()
            time.sleep(random.uniform(0.2, 0.8))
    except Exception:
        pass


def type_like_human(element, keyword: str) -> None:
    element.click()
    for char in keyword:
        element.send_keys(char)
        time.sleep(random.uniform(0.05, 0.18))


def accept_cookie_popup(driver: webdriver.Chrome) -> None:
    selectors = [
        (By.CSS_SELECTOR, "button#onetrust-accept-btn-handler"),
        (By.XPATH, "//button[contains(., '모두 동의')]"),
        (By.XPATH, "//button[contains(., '동의')]"),
        (By.XPATH, "//button[contains(., 'Accept')]"),
    ]
    for by, selector in selectors:
        elements = driver.find_elements(by, selector)
        if elements:
            elements[0].click()
            time.sleep(random.uniform(0.6, 1.2))
            break


def is_access_denied(driver: webdriver.Chrome) -> bool:
    title = driver.title
    if "Access Denied" in title or "접근이 거부" in title:
        return True
    source = driver.page_source
    return "Access Denied" in source or "접근이 거부" in source


def open_search_from_home(driver: webdriver.Chrome, keyword: str) -> None:
    driver.get("https://www.coupang.com")
    time.sleep(random.uniform(3, 5))
    accept_cookie_popup(driver)
    selectors = [
        "input[name='q']",
        "input#headerSearchKeyword",
        "input[placeholder*='검색']",
    ]
    for selector in selectors:
        try:
            element = driver.find_element(By.CSS_SELECTOR, selector)
            type_like_human(element, keyword)
            element.send_keys(Keys.ENTER)
            return
        except NoSuchElementException:
            continue
    driver.get(build_search_url(keyword))


def extract_rank_text(item) -> str:
    rank_selectors = [
        ".search-product__rank",
        ".search-product__rank-text",
        ".rank-badge",
        ".product-rank",
    ]
    for selector in rank_selectors:
        elements = item.find_elements(By.CSS_SELECTOR, selector)
        if elements:
            text = elements[0].text.strip()
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
        if item.find_elements(By.CSS_SELECTOR, selector):
            return True
    item_text = item.text
    return "AD" in item_text or "광고" in item_text


def extract_product_name(item) -> str:
    name_selectors = [
        ".name",
        ".search-product__name",
        "a.search-product-link",
    ]
    for selector in name_selectors:
        elements = item.find_elements(By.CSS_SELECTOR, selector)
        if elements:
            text = elements[0].text.strip()
            if text:
                return text
    return ""


def detect_rocket_badge(item) -> str:
    badge_texts = []
    badge_elements = item.find_elements(By.CSS_SELECTOR, "img[alt], span, em, i")
    for element in badge_elements[:20]:
        alt_text = element.get_attribute("alt") or ""
        text = element.text.strip()
        combined = f"{alt_text} {text}".strip()
        if combined:
            badge_texts.append(combined)
    full_text = " ".join(badge_texts)
    if "판매자로켓" in full_text:
        return "판매자로켓"
    if "로켓" in full_text:
        return "로켓배송"
    return "뱃지없음"


def analyze_keyword(keyword: str, test_mode: bool) -> SearchResult:
    organic_count = 0
    rocket_count = 0
    error_message = ""
    debug_items = []

    driver, profile_dir, process = create_driver()
    try:
        open_search_from_home(driver, keyword)
        WebDriverWait(driver, 30).until(ec.presence_of_element_located((By.CSS_SELECTOR, "li.search-product")))
        time.sleep(random.uniform(1.0, 2.2))
        human_like_scroll(driver)
        human_like_mouse(driver)

        if is_access_denied(driver):
            time.sleep(random.uniform(30, 60))
            open_search_from_home(driver, keyword)
            WebDriverWait(driver, 30).until(ec.presence_of_element_located((By.CSS_SELECTOR, "li.search-product")))

        items = driver.find_elements(By.CSS_SELECTOR, "li.search-product")
        ranked_items = {}

        for item in items:
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
                file.write(driver.page_source)
            driver.save_screenshot(os.path.join("output", "debug.png"))
            for item in debug_items:
                rank_label = item["rank_text"] or "순위없음"
                name_label = item["product_name"] or "상품명없음"
                ad_label = "광고" if item["is_ad"] else "자연노출"
                badge_label = item["badge_type"]
                print(f"[TEST] {rank_label} | {name_label} | {ad_label} | {badge_label}")
    except (TimeoutException, Exception) as exc:
        error_message = str(exc)
    finally:
        cleanup_driver(driver, profile_dir, process)

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

    for index, keyword in enumerate(keywords, start=1):
        if keyword in completed:
            continue

        last_error = ""
        result: Optional[SearchResult] = None

        for attempt in range(1, 4):
            try:
                result = analyze_keyword(keyword, args.test)
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

        time.sleep(random.uniform(5, 12))


if __name__ == "__main__":
    main()

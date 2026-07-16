from bs4 import BeautifulSoup
import requests


# Standard headers to fetch a website
headers = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/117.0.0.0 Safari/537.36"
}


def fetch_website_contents(url):
    """
    Return the title and contents of the website at the given url;
    truncate to 2,000 characters as a sensible limit
    """
    response = requests.get(url, headers=headers)
    soup = BeautifulSoup(response.content, "html.parser")
    title = soup.title.string if soup.title else "No title found"
    if soup.body:
        for irrelevant in soup.body(["script", "style", "img", "input"]):
            irrelevant.decompose()
        text = soup.body.get_text(separator="\n", strip=True)
    else:
        text = ""
    return (title + "\n\n" + text)[:2_000]


def fetch_website_links(url):
    """
    Return the links on the webiste at the given url
    I realize this is inefficient as we're parsing twice! This is to keep the code in the lab simple.
    Feel free to use a class and optimize it!
    """
    response = requests.get(url, headers=headers)
    soup = BeautifulSoup(response.content, "html.parser")
    links = [link.get("href") for link in soup.find_all("a")]
    return [link for link in links if link]


def setup_driver(headless: bool = True):
    options = Options()
    if headless:
        options.add_argument("--headless=new")
    options.add_argument("--disable-gpu")
    options.add_argument("--window-size=1920,1080")
    options.add_argument("--no-sandbox")
    driver = webdriver.Chrome(
        service=Service(ChromeDriverManager().install()),
        options=options
    )
    return driver


def scrape_quotes_with_pagination(start_url: str, max_pages: int = 10):
    """
    Scrapes quotes across multiple pages from:
    https://quotes.toscrape.com/js/
    """
    driver = setup_driver(headless=True)
    wait = WebDriverWait(driver, 15)
    all_rows = []
    try:
        driver.get(start_url)
        current_page = 1
        while current_page <= max_pages:
            # Wait for quote cards on each page
            wait.until(EC.presence_of_all_elements_located((By.CLASS_NAME, "quote")))
            quote_cards = driver.find_elements(By.CLASS_NAME, "quote")
            for card in quote_cards:
                text = card.find_element(By.CLASS_NAME, "text").text.strip()
                author = card.find_element(By.CLASS_NAME, "author").text.strip()
                tags = [t.text.strip() for t in card.find_elements(By.CLASS_NAME, "tag")]
                all_rows.append(
                    {
                        "page": current_page,
                        "quote": text,
                        "author": author,
                        "tags": ", ".join(tags)
                    }
                )
            # Try to go to next page
            next_buttons = driver.find_elements(By.CSS_SELECTOR, "li.next > a")
            if not next_buttons:
                print("No more pages found. Stopping.")
                break
            # Click next page
            driver.execute_script("arguments[0].click();", next_buttons[0])
            # Small pause helps dynamic content settle (plus explicit waits)
            time.sleep(0.5)
            current_page += 1
        return all_rows
    finally:
        driver.quit()


def save_to_csv(rows, output_file="quotes_paginated.csv"):
    df = pd.DataFrame(rows)
    df.to_csv(output_file, index=False, encoding="utf-8")
    print(f"Saved {len(df)} rows to {output_file}")


def stream_brochure(company_name, url):
    stream = openai.chat.completions.create(
        model="gpt-4.1-mini",
        messages=[
            {"role": "system", "content": brochure_system_prompt},
            {"role": "user", "content": get_brochure_user_prompt(company_name, url)}
          ],
        stream=True
    )    
    response = ""
    display_handle = display(Markdown(""), display_id=True)
    for chunk in stream:
        response += chunk.choices[0].delta.content or ''
        update_display(Markdown(response), display_id=display_handle.display_id)


def get_brochure_user_prompt(company_name, url):
    user_prompt = f"""
You are looking at a company called: {company_name}
Here are the contents of its landing page and other relevant pages;
use this information to build a short brochure of the company in markdown without code blocks.\n\n
"""
    user_prompt += fetch_page_and_all_relevant_links(url)
    user_prompt = user_prompt[:5_000] # Truncate if more than 5,000 characters
    return user_prompt


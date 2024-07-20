import time
import yaml
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains
from article_id_generator import ArticleIdGenerator

# in this section you can read config.yaml
with open('config.yaml', 'r') as f:
    config = yaml.safe_load(f)

# you can set your browser setting in here
options = webdriver.ChromeOptions()
if config['headless']:
    options.add_argument('--headless')

driver = webdriver.Chrome(options=options)
id_generator = ArticleIdGenerator()
base_url = config['base_url']

driver.get(base_url)
time.sleep(1)

journal_name = config['journal_name']
issue_links = set()

for i in range(3, 14):
    xpath = f"{config['issue_xpath']}[{i}]"
    element = driver.find_element(By.XPATH, xpath)

    actions = ActionChains(driver)
    actions.move_to_element(element).click().perform()

    time.sleep(0.5)

    issue_elements = driver.find_elements(By.XPATH, config['issue_link_xpath'])
    for elem in issue_elements:
        issue_links.add(elem.get_attribute('href'))

    time.sleep(0.5)

issue_links = list(issue_links)
article_links = set()

for issue_link in issue_links:
    driver.get(issue_link)
    time.sleep(1)
    article_elements = driver.find_elements(By.XPATH, config['article_link_xpath'])
    for elem in article_elements:
        link = elem.get_attribute('href')
        if not link.endswith('.pdf'):
            article_links.add(link)

article_links = list(article_links)
articles_data = []
references_data = []

for article_link in article_links:
    driver.get(article_link)
    time.sleep(1)

    article_data = {}
    article_data['id'] = id_generator.generate_id()
    article_data['journal_name'] = journal_name

    try:
        article_data['title_fa'] = driver.find_element(By.XPATH, config['title_fa_xpath']).text
    except:
        article_data['title_fa'] = None

    try:
        article_data['title_en'] = driver.find_element(By.XPATH, config['title_en_xpath']).text
    except:
        article_data['title_en'] = None

    try:
        authors_elements = driver.find_elements(By.XPATH, config['authors_fa_xpath'])
        authors_text = ''
        for i, elem in enumerate(authors_elements):
            authors_text += elem.text
            if i < len(authors_elements) - 1:
                authors_text += '; '
        article_data['authors_fa'] = authors_text
    except:
        article_data['authors_fa'] = None

    try:
        authors_elements = driver.find_elements(By.XPATH, config['authors_en_xpath'])
        authors_text = ''
        for i, elem in enumerate(authors_elements):
            authors_text += elem.text
            if i < len(authors_elements) - 1:
                authors_text += '; '
        article_data['authors_en'] = authors_text
    except:
        article_data['authors_en'] = None

    try:
        article_data['keywords_fa'] = driver.find_element(By.XPATH, config['keywords_fa_xpath']).text
    except:
        article_data['keywords_fa'] = None

    try:
        article_data['keywords_en'] = driver.find_element(By.XPATH, config['keywords_en_xpath']).text
    except:
        article_data['keywords_en'] = None

    try:
        article_data['abstract_fa'] = driver.find_element(By.XPATH, config['abstract_fa_xpath']).text
    except:
        article_data['abstract_fa'] = None

    try:
        article_data['abstract_en'] = driver.find_element(By.XPATH, config['abstract_en_xpath']).text
    except:
        article_data['abstract_en'] = None

    try:
        article_data['year'] = driver.find_element(By.XPATH, config['year_xpath']).text
    except:
        article_data['year'] = None

    try:
        wait = WebDriverWait(driver, 10)
        element = wait.until(EC.presence_of_element_located((By.XPATH, config['references_toggle_xpath'])))
        driver.execute_script("arguments[0].scrollIntoView(true);", element)
        element.click()
        time.sleep(0.5)
        driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
        time.sleep(0.5)

        divs_1 = driver.find_elements(By.XPATH, config['references_rtl_xpath'])
        divs_2 = driver.find_elements(By.XPATH, config['references_ltr_xpath'])

        texts = [div.text.strip() for div in divs_1] + [div.text.strip() for div in divs_2]

        for text in texts:
            references_data.append({'id': article_data['id'], 'reference': text})

        article_data['references'] = texts
    except Exception as e:
        print(f"Error extracting references: {e}")

    articles_data.append(article_data)

driver.quit()

df_articles = pd.DataFrame(articles_data, columns=[
    'title_fa', 'title_en', 'authors_fa', 'authors_en',
    'keywords_fa', 'keywords_en', 'abstract_fa', 'abstract_en', 'year', 'journal_name', 'id', 'references'
])

df_references = pd.DataFrame(references_data, columns=['id', 'reference'])

df_articles = df_articles.drop(columns=['references'])

with pd.ExcelWriter('articles_data.xlsx') as writer:
    df_articles.to_excel(writer, sheet_name='Articles', index=False)
    df_references.to_excel(writer, sheet_name='References', index=False)

print("Data successfully extracted and saved to Excel.")

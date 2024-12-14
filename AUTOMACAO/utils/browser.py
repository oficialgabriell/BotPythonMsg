from selenium import webdriver


def get_browser(*options, implicit_wait=10):
    chrome_options = webdriver.ChromeOptions()
    for option in options:
        chrome_options.add_argument(option)
    browser = webdriver.Chrome(options=chrome_options)
    browser.implicitly_wait(implicit_wait)
    return browser

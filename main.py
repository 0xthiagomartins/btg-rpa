from selenium import webdriver
from webdriver_manager.chrome import ChromeDriverManager

url = "https://rpachallenge.com/"
driver = webdriver.Chrome(ChromeDriverManager().install())

if __name__ == "__main__":
    driver.get(url)

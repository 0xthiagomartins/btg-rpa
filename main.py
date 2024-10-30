import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service as ChromeService
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import (
    TimeoutException,
    NoSuchElementException,
    ElementNotInteractableException,
)
from time import sleep


def load_data():
    return pd.read_excel("./data/challenge.xlsx")


driver = webdriver.Chrome(service=ChromeService(ChromeDriverManager().install()))
url = "https://rpachallenge.com/"
COLLUMNS = [
    "First Name",
    "Last Name ",
    "Company Name",
    "Role in Company",
    "Address",
    "Email",
    "Phone Number",
]


def start(driver):
    wait = WebDriverWait(driver, 10)
    start_button = wait.until(
        EC.presence_of_element_located(
            (By.XPATH, "/html/body/app-root/div[2]/app-rpa1/div/div[1]/div[6]/button")
        )
    )
    start_button.click()


def fill_form(row, driver):
    print("Starting to fill form for the current row.")
    wait = WebDriverWait(driver, 10)
    for collumn in COLLUMNS:
        value = row.loc[collumn]
        print(f"Filling '{collumn}' with value: {value}")
        label_xpath = f"//label[normalize-space(text())='{collumn.strip()}']"
        print(f"Locating label using XPath: {label_xpath}")
        label_element = wait.until(
            EC.presence_of_element_located((By.XPATH, label_xpath))
        )
        print(f"Label for '{collumn}' found.")

        input_element = label_element.find_element(By.XPATH, "./..").find_element(
            By.XPATH, ".//input"
        )
        print(f"Located input field for '{collumn}'. Waiting to be clickable.")
        wait.until(EC.element_to_be_clickable(input_element))
        print(f"Input field for '{collumn}' is clickable. Clearing existing text.")
        input_element.clear()
        print(f"Sending keys to '{collumn}': {value}")
        input_element.send_keys(value)
    print("Finished filling the form for the current row.")


if __name__ == "__main__":
    try:
        print("Loading data from Excel...")
        data = load_data()
        print(f"Navigating to {url}")
        driver.get(url)
        wait = WebDriverWait(driver, 10)
        sleep(2.5)
        print("Startign challenge")
        start(driver)
        for index, row in data.iterrows():
            print("=" * 25)
            print("\n\n")
            print(f"Processing record {index + 1}/{len(data)}")
            fill_form(row, driver)
            print("Submitting the form...")
            submit_button = wait.until(
                EC.element_to_be_clickable(
                    (By.CSS_SELECTOR, "input[type='submit'][value='Submit']")
                )
            )
            sleep(2.5)
            submit_button.click()
            print("Form submitted successfully.\n")

    except TimeoutException:
        print("Timed out waiting for the element to be present or interactable.")
    except NoSuchElementException:
        print("The specified element was not found in the DOM.")
    except ElementNotInteractableException:
        print("The element is present but not interactable.")
    except Exception as e:
        print(f"An unexpected error occurred: {e.__str__()}")
    finally:
        # Timing to manual observe the error
        sleep(5)

        driver.quit()

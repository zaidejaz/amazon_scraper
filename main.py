import time
import undetected_chromedriver as uc
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import pandas as pd
from amazoncaptcha import AmazonCaptcha

def scrape_amazon(driver, url):
    driver.get(url)
    
    # Handle captcha
    try:
        captcha = WebDriverWait(driver, 5).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, 'form[action="/errors/validateCaptcha"]'))
        )
        if captcha:
            captcha_img = WebDriverWait(driver, 30).until(
                EC.presence_of_element_located((By.CSS_SELECTOR, 'img[src*="captcha"]'))
            )
            if captcha_img:
                captcha_url = captcha_img.get_attribute('src')
                print(f'Captcha URL: {captcha_url}')
                solver = AmazonCaptcha.fromlink(captcha_url)
                solution = solver.solve()
                
                WebDriverWait(driver, 30).until(
                    EC.presence_of_element_located((By.ID, 'captchacharacters'))
                ).send_keys(solution)
                WebDriverWait(driver, 30).until(
                    EC.element_to_be_clickable((By.CSS_SELECTOR, 'button[type="submit"]'))
                ).click()
                
    except:
        print("No captcha found or timed out waiting for captcha.")

    try:
        # Extract average rating (unchanged)
        average_rating = WebDriverWait(driver, 30).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, 'span[data-hook="rating-out-of-text"]'))
        ).get_attribute('innerHTML').split(' out of')[0]
        print(f"Average rating: {average_rating}")

        # Extract total number of reviews (unchanged)
        total_reviews = WebDriverWait(driver, 30).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, 'div[data-hook="total-review-count"] span'))
        ).get_attribute('innerHTML').split(' global')[0].replace(',', '')
        total_reviews = int(total_reviews)
        print(f"Total reviews: {total_reviews}")

        # Extract number of relevant reviews (unchanged)
        relevant_reviews = WebDriverWait(driver, 30).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, 'div[data-hook="cr-filter-info-review-rating-count"] span'))
        ).get_attribute('innerHTML').split(' global')[0].replace(',', '')
        relevant_reviews = int(relevant_reviews)
        print(f"Relevant reviews: {relevant_reviews}")

        result = {
            'average_rating': average_rating,
            'total_number_of_reviews': total_reviews,
            'number_of_relevant_reviews': relevant_reviews,
        }
        print(f"Final result: {result}")
        return result
    except:
        print("Failed to scrape data.")
        return None

def update_amazon_value(current_value, new_amazon_value, key='Amazon'):
    if isinstance(current_value, (int, float)):
        return f"{key}: {new_amazon_value}"
    elif isinstance(current_value, str):
        parts = current_value.split(',')
        for i, part in enumerate(parts):
            if f"{key}:" in part:
                parts[i] = f"{key}: {new_amazon_value}"
                break
        else:
            parts.append(f"{key}: {new_amazon_value}")
        return ','.join(parts)
    else:
        return f"{key}: {new_amazon_value}"

def update_excel(file_path):
    df = pd.read_excel(file_path)
    options = uc.ChromeOptions()
    prefs = {'profile.default_content_setting_values': {'cookies': 2, 'images': 2, 'javascript': 2, 
                            'plugins': 2, 'popups': 2, 'geolocation': 2, 
                            'notifications': 2, 'auto_select_certificate': 2, 'fullscreen': 2, 
                            'mouselock': 2, 'mixed_script': 2, 'media_stream': 2, 
                            'media_stream_mic': 2, 'media_stream_camera': 2, 'protocol_handlers': 2, 
                            'ppapi_broker': 2, 'automatic_downloads': 2, 'midi_sysex': 2, 
                            'push_messaging': 2, 'ssl_cert_decisions': 2, 'metro_switch_to_desktop': 2, 
                            'protected_media_identifier': 2, 'app_banner': 2, 'site_engagement': 2, 
                            'durable_storage': 2}}
    options.add_experimental_option('prefs', prefs)
    driver = uc.Chrome(options=options)
    driver.implicitly_wait(30)

    try:
        for index, row in df.iterrows():
            amazon_url = row['reviews_link'].split('Amazon: ')[-1]
            
            scraped_data = scrape_amazon(driver, amazon_url)
            
            if scraped_data is None:
                continue
            df.at[index, 'average_rating'] = update_amazon_value(row['average_rating'], scraped_data['average_rating'])
            df.at[index, 'total_number_of_reviews'] = int(scraped_data['total_number_of_reviews'])
            df.at[index, 'number_of_relevant_reviews'] = update_amazon_value(
                row['number_of_relevant_reviews'], 
                str(scraped_data['number_of_relevant_reviews'])
            )
    
    finally:
        driver.quit()

    df.to_excel(file_path, index=False)
    print(f"Excel file updated: {file_path}")
# Usage
update_excel('data.xlsx')
import time
import json
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC  
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.service import Service
from selenium.common.exceptions import (NoSuchElementException, 
                                     TimeoutException, 
                                     ElementNotInteractableException,
                                     WebDriverException)

def login_to_sensirion(driver, email, password):
    """Log in to the Sensirion portal"""
    print("Navigating to login page...")
    try:
        login_url = ("https://sensirioncustomer.b2clogin.com/sensirioncustomer.onmicrosoft.com/"
                    "b2c_1_signin_signup_flow/oauth2/v2.0/authorize?response_type=code&"
                    "client_id=146eb5d4-c8fe-46a3-908b-09044f9ef318&"
                    "redirect_uri=https%3A%2F%2Flibellus.sensirion.com%2F.auth%2Flogin%2Faadb2c%2Fcallback&"
                    "nonce=a33bd22f4910462fbfdbe096dc5a923a_20250415155701&"
                    "state=redir%3D%252Fauth%252Fcallback%252F&scope=openid+profile+email&"
                    "post_login_redirect_uri=%2Fauth%2Fcallback%2F")
        
        driver.get(login_url)
        
        # Wait for email field and enter credentials
        email_field = WebDriverWait(driver, 20).until(
            EC.element_to_be_clickable((By.ID, "email"))
        )
        email_field.clear()
        for char in email:
            email_field.send_keys(char)
            time.sleep(0.1)
        
        password_field = WebDriverWait(driver, 20).until(
            EC.element_to_be_clickable((By.ID, "password"))
        )
        password_field.clear()
        for char in password:
            password_field.send_keys(char)
            time.sleep(0.1)
        
        # Click login button
        login_button = WebDriverWait(driver, 20).until(
            EC.element_to_be_clickable((By.XPATH, "//button[contains(text(), 'Sign in')]"))
        )
        login_button.click()
        
        # Wait for login to complete
        WebDriverWait(driver, 20).until(
            EC.presence_of_element_located((By.LINK_TEXT, "Web Access"))
        )
        print("Login successful")
        return True
    except Exception as e:
        print(f"Login failed: {str(e)}")
        driver.save_screenshot("login_error.png")
        return False

def get_calibration_data(driver, serial_number):
    """Retrieve calibration data for a specific SHT33 sensor"""
    print(f"\nRetrieving data for serial: {serial_number}")
    
    try:
        # Navigate to Web Access
        print("Clicking Web Access link...")
        web_access_link = WebDriverWait(driver, 30).until(
            EC.element_to_be_clickable((By.LINK_TEXT, "Web Access"))
        )
        web_access_link.click()
        print("Clicked Web Access")
        
        # Wait for page to load
        WebDriverWait(driver, 30).until(
            lambda d: d.execute_script("return document.readyState") == "complete"
        )
        
        # Verify we're on the API Companion page
        WebDriverWait(driver, 30).until(
            EC.presence_of_element_located((By.XPATH, "//h3[contains(text(), 'API Companion')]"))
        )
        print("On API Companion page")
        
        # Switch to the Calibration Information tab
        print("Switching to Calibration Information tab...")
        calib_tab = WebDriverWait(driver, 30).until(
            EC.element_to_be_clickable((By.CSS_SELECTOR, "#pills-calibinfo-tab"))
        )
        
        # Scroll into view and click using JavaScript
        driver.execute_script("arguments[0].scrollIntoView(true);", calib_tab)
        time.sleep(1)
        driver.execute_script("arguments[0].click();", calib_tab)
        print("Switched to Calibration Information tab")
        
        # Wait for the form to load
        WebDriverWait(driver, 30).until(
            EC.presence_of_element_located((By.ID, "pill-calib-info"))
        )
        
        # Select sensor model (required before entering serial)
        print("Selecting sensor model...")
        model_dropdown = WebDriverWait(driver, 30).until(
            EC.element_to_be_clickable((By.ID, "id_sensor_model_calib"))
        )
        model_dropdown.click()
        time.sleep(0.5)
        
        # Select SHT33 option
        sht33_option = WebDriverWait(driver, 30).until(
            EC.element_to_be_clickable((By.XPATH, "//select[@id='id_sensor_model_calib']/option[text()='SHT33']"))
        )
        sht33_option.click()
        print("Selected SHT33 model")
        
        # Enter serial number with multiple fallback methods
        print("Entering serial number...")
        serial_field = WebDriverWait(driver, 30).until(
            EC.element_to_be_clickable((By.CSS_SELECTOR, "#pill-calib-info input#id_sensor_serial"))
        )
        
        # Clear field using JavaScript
        driver.execute_script("arguments[0].value = '';", serial_field)
        time.sleep(0.5)
        
        # Try different input methods
        try:
            serial_field.send_keys(serial_number)
        except:
            driver.execute_script(f"arguments[0].value = '{serial_number}';", serial_field)
            driver.execute_script("""
                arguments[0].dispatchEvent(new Event('input', {bubbles: true}));
                arguments[0].dispatchEvent(new Event('change', {bubbles: true}));
            """, serial_field)
        
        # Verify input
        entered_value = driver.execute_script("return arguments[0].value;", serial_field)
        if entered_value != serial_number:
            raise ValueError(f"Failed to set serial number. Got: {entered_value}")
        
        print("Serial number entered successfully")
        
        # Submit the form
        print("Submitting form...")
        submit_button = WebDriverWait(driver, 30).until(
            EC.element_to_be_clickable((By.CSS_SELECTOR, "#pill-calib-info button.btn-sensi"))
        )
        submit_button.click()
        
        print("Waiting for response...")
        try:
            # Wait for either the success alert or error message
            WebDriverWait(driver, 45).until(
                EC.any_of(
                    EC.presence_of_element_located((By.CSS_SELECTOR, "div.alert-success")),
                    EC.presence_of_element_located((By.CSS_SELECTOR, "div.alert-danger"))
                )
            )
            
            # Check for success response first
            success_alert = driver.find_elements(By.CSS_SELECTOR, "div.alert-success")
            if success_alert:
                print("Found success alert - extracting JSON data")
                try:
                    json_p = WebDriverWait(driver, 20).until(
                        EC.presence_of_element_located((By.XPATH, '//*[@id="pill-calib-info"]/div/p[1]'))
                    )
                    json_text = json_p.text.strip()
                    
                    # Debug output
                    print(f"Raw JSON text extracted: {json_text}")
                    try:
                        json_data = json.loads(json_text)
                        print("Successfully parsed JSON data")
                        return {
                            "serial_number": serial_number,
                            "status": "success",
                            "data": json_data
                        }
                    except json.JSONDecodeError as e:
                        print(f"Failed to parse JSON: {str(e)}")
                        print(f"JSON text that failed to parse: {json_text}")

                        return {
                            "serial_number": serial_number,
                            "status": "error",
                            "error": f"JSON parsing failed: {str(e)}",
                            "raw_response": json_text
                        }
                    
                except NoSuchElementException:
                    print("Could not find JSON data in success alert")
                    return {
                        "serial_number": serial_number,
                        "status": "error",
                        "error": "Success alert found but no JSON data detected"
                    }
            
            # Check for error response
            error_alert = driver.find_elements(By.CSS_SELECTOR, "div.alert-danger")
            if error_alert:
                error_msg = error_alert[0].text
                print(f"Error response: {error_msg}")
                return {
                    "serial_number": serial_number,
                    "status": "error",
                    "error": error_msg
                }
            
            # If neither found, return timeout
            return {
                "serial_number": serial_number,
                "status": "timeout",
                "error": "No response detected"
            }
            
        except TimeoutException:
            print("Timeout waiting for response")
            return {
                "serial_number": serial_number,
                "status": "timeout",
                "error": "Timeout waiting for response"
            }
            
    except Exception as e:
        print(f"Error extracting JSON using XPath: {str(e)}")
        # Fallback to trying to get the entire alert text
        try:
            alert_text = success_alert[0].text
            json_start = alert_text.find('{')
            if json_start == -1:
                raise ValueError("No JSON data found in alert")
            json_end = alert_text.rfind('}') + 1
            json_text = alert_text[json_start:json_end].strip()
            
            json_data = json.loads(json_text)
            return {
                "serial_number": serial_number,
                "status": "success",
                "data": json_data
            }
        except Exception as fallback_e:
            print(f"Fallback extraction also failed: {str(fallback_e)}")
            return {
                "serial_number": serial_number,
                "status": "error",
                "error": f"Both XPath and fallback extraction failed: {str(e)} | {str(fallback_e)}",
                "raw_response": alert_text if 'alert_text' in locals() else "Could not extract any text"
            }

def main():
    # Configuration
    config = {
        "email": "xyz@gmail.com",
        "password": "qwerty@1234",
        "input_file": "sensor_serials.xlsx",
        "output_file": "calibration_data.xlsx",
        "sheet_name": "Sheet1",
        "possible_serial_columns": ["Serial Number", "Serial", "SERIAL", "Serial No", "S/N"],
        "max_attempts": 3,
        "retry_delay": 5
    }
    
    driver = None
    try:
        # Initialize Chrome WebDriver
        print("\nInitializing Chrome driver...")
        service = Service(ChromeDriverManager().install())
        options = webdriver.ChromeOptions()
        options.add_argument("--start-maximized")
        options.add_argument("--disable-blink-features=AutomationControlled")
        options.add_experimental_option("excludeSwitches", ["enable-automation"])
        options.add_experimental_option('useAutomationExtension', False)
        
        driver = webdriver.Chrome(service=service, options=options)
        
        # Add a delay to ensure everything loads
        driver.implicitly_wait(10)
        
        # Login to Sensirion
        if not login_to_sensirion(driver, config["email"], config["password"]):
            raise Exception("Login failed - stopping execution")
        
        # Read serial numbers from Excel
        print(f"\nReading serial numbers from {config['input_file']}")
        try:
            df = pd.read_excel(
                config["input_file"],
                sheet_name=config["sheet_name"]
            )
            
            # Find the correct serial number column
            serial_column = None
            for col in config["possible_serial_columns"]:
                if col in df.columns:
                    serial_column = col
                    break
            
            if not serial_column:
                available_columns = ", ".join(df.columns)
                raise Exception(f"Could not find serial number column. Available columns: {available_columns}")
            
            print(f"Using column '{serial_column}' for serial numbers")
            serial_numbers = df[serial_column].dropna().astype(str).tolist()
            print(f"Found {len(serial_numbers)} serial numbers to process")
            
            if len(serial_numbers) == 0:
                raise Exception("No serial numbers found in the specified column")
            
        except Exception as e:
            raise Exception(f"Error reading Excel file: {str(e)}")
        
        results = []
        
        # Process each serial number
        for i, serial in enumerate(serial_numbers, 1):
            print(f"\nProcessing {i}/{len(serial_numbers)}: {serial}")
            
            # Try multiple attempts for each serial number
            for attempt in range(1, config["max_attempts"] + 1):
                print(f"\nAttempt {attempt} of {config['max_attempts']}")
                try:
                    data = get_calibration_data(driver, serial)
                    data["serial_number"] = serial
                    results.append(data)
                    break  # Success, move to next serial
                except Exception as e:
                    print(f"Attempt {attempt} failed: {str(e)}")
                    if attempt == config["max_attempts"]:
                        results.append({
                            "serial_number": serial, 
                            "error": f"Failed after {config['max_attempts']} attempts: {str(e)}"
                        })
                    print(f"Waiting {config['retry_delay']} seconds before retry...")
                    time.sleep(config["retry_delay"])
            
            # Add delay to avoid overwhelming the server
            print(f"Waiting 3 seconds before next serial...")
            time.sleep(3)
        
        # Save results to Excel
        print("\nSaving results...")
        # Create a list of flattened records
        flattened_results = []
        for result in results:
            record = {'serial_number': result['serial_number']}
            
            if 'data' in result:
                # For successful results, add all data fields
                record.update(result['data'])
                record['status'] = 'success'
             else:
                # For error cases, add error information
                record['status'] = result.get('status', 'error')
                record['error'] = result.get('error', 'Unknown error')
                if 'raw_response' in result:
                    record['raw_response'] = result['raw_response']
            
            flattened_results.append(record)

        output_df = pd.DataFrame(flattened_results)
        cols = ['serial_number', 'status'] + [col for col in output_df.columns 
                                            if col not in ['serial_number', 'status']]
        output_df = output_df[cols]
        
        output_df.to_excel(config["output_file"], index=False)
        print(f"Data successfully saved to {config['output_file']}")
    
    except Exception as e:
        print(f"\nFatal error: {str(e)}")
        if driver:
            driver.save_screenshot("fatal_error.png")
    
    finally:
        if driver:
            print("\nClosing browser...")
            driver.quit()

if __name__ == "__main__":
    main()
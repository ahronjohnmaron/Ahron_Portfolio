import time
import re
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
from selenium.common.exceptions import TimeoutException, NoSuchElementException, ElementNotInteractableException
from datetime import datetime, timedelta

class GmailAutomationWithExcel:
    def __init__(self, excel_file_path, headless=False):
        self.driver = None
        self.wait = None
        self.headless = headless
        self.excel_file_path = excel_file_path
        self.df = None
        self.current_director_name = ""
        self.require_confirmation = True
        
    def safe_click(self, element):
        """Robust click that tries multiple methods to click an element."""
        try:
            # Scroll element into view first
            self.driver.execute_script("arguments[0].scrollIntoView({block:'center'});", element)
            time.sleep(0.5)
            element.click()
            return True
        except Exception:
            pass

        try:
            self.driver.execute_script("arguments[0].click();", element)
            return True
        except Exception:
            pass

        try:
            ActionChains(self.driver).move_to_element(element).pause(0.2).click().perform()
            return True
        except Exception:
            pass

        try:
            descendant = element.find_element(By.CSS_SELECTOR, "a, span, div, td")
            descendant.click()
            return True
        except Exception:
            pass

        return False

    def ensure_valid_window_handle(self):
        """Ensure we're on a valid browser tab, switch if current one is closed"""
        try:
            # Test if current window handle is valid
            self.driver.current_url
            return True
        except Exception as e:
            print("⚠️ Current tab closed or unavailable, switching to another tab...")
            try:
                handles = self.driver.window_handles
                if handles:
                    self.driver.switch_to.window(handles[-1])
                    print(f"✅ Switched to available tab: {self.driver.title}")
                    return True
                else:
                    print("❌ No available browser tabs")
                    return False
            except Exception as ex:
                print(f"❌ Error switching tabs: {str(ex)}")
                return False

    def get_current_gmail_url(self):
        """Get the current Gmail URL to preserve delegated profile session"""
        try:
            # Ensure we have a valid window first
            if not self.ensure_valid_window_handle():
                return 'https://mail.google.com/mail/u/0/#inbox'
            
            current_url = self.driver.current_url
            
            # Extract the Gmail base URL with account info
            if 'mail.google.com' in current_url:
                # Keep account-specific part of URL (e.g., /u/1/ for delegated profiles)
                if '/u/' in current_url:
                    base_url = current_url.split('#')[0]  # Remove fragment
                    return base_url + '#inbox'
                else:
                    return current_url.split('#')[0] + '#inbox'
            else:
                return 'https://mail.google.com/mail/u/0/#inbox'
                
        except Exception as e:
            print(f"⚠️ Error getting Gmail URL: {str(e)}")
            return 'https://mail.google.com/mail/u/0/#inbox'

    def verify_gmail_loaded(self):
        """Verify that Gmail is properly loaded in the current tab"""
        try:
            # Ensure we have a valid window first
            if not self.ensure_valid_window_handle():
                return False
                
            print("🔍 Verifying Gmail is loaded...")
            
            gmail_indicators = [
                "[aria-label*='Search mail']",
                "[data-tooltip*='Compose']",
                ".gb_d[aria-label*='Gmail']",
                "#logo",
                ".nH",
                "[gh='tl']",
                ".aAy"
            ]
            
            for indicator in gmail_indicators:
                try:
                    WebDriverWait(self.driver, 15).until(
                        EC.presence_of_element_located((By.CSS_SELECTOR, indicator)))
                    print(f"✅ Gmail loaded successfully - found {indicator}")
                    return True
                except:
                    continue
                    
            print("⚠️ Gmail may not be fully loaded")
            return False
            
        except Exception as e:
            print(f"⚠️ Error verifying Gmail: {str(e)}")
            return False

    def set_email_body_font_arial(self):
        """Set the email body font to Arial"""
        try:
            # Find the email body element
            selectors = [
                "div[aria-label*='Message Body']",
                "div[role='textbox'][aria-label*='Message']",
                ".Am.Al.editable",
                "div[contenteditable='true'][aria-label*='Message']",
                ".editable[role='textbox']"
            ]
            
            body_element = None
            for selector in selectors:
                try:
                    body_element = self.driver.find_element(By.CSS_SELECTOR, selector)
                    if body_element:
                        break
                except:
                    continue
            
            if not body_element:
                print("❌ Could not find email body element for font formatting")
                return False
            
            # Click to focus on the element
            body_element.click()
            time.sleep(0.5)
            
            # Clear existing formatting by selecting all and removing formatting
            actions = ActionChains(self.driver)
            actions.key_down(Keys.CONTROL).send_keys('a').key_up(Keys.CONTROL).perform()
            time.sleep(0.2)
            # Ctrl+\ removes formatting in Gmail
            actions.key_down(Keys.CONTROL).send_keys('\\').key_up(Keys.CONTROL).perform()
            time.sleep(0.3)
            
            # Set font to Arial using JavaScript
            script = """
            var el = arguments[0];
            if (el) {
                el.style.fontFamily = 'Arial, sans-serif';
                el.style.fontSize = '11px';  // Standard Gmail font size
            }
            """
            self.driver.execute_script(script, body_element)
            print("✅ Email body font set to Arial")
            return True
            
        except Exception as e:
            print(f"❌ Error setting font to Arial: {str(e)}")
            return False
        
    def load_excel_data(self):
        """Load the Excel file with contact information"""
        try:
            self.df = pd.read_excel(self.excel_file_path)
            print(f"Loaded {len(self.df)} contacts from Excel file")
            return True
        except Exception as e:
            print(f"Error loading Excel file: {str(e)}")
            return False
    
    def setup_driver(self):
        """Initialize Chrome driver with improved options"""
        chrome_options = Options()
        chrome_options.add_argument("--no-sandbox")
        chrome_options.add_argument("--disable-dev-shm-usage")
        chrome_options.add_argument("--disable-blink-features=AutomationControlled")
        chrome_options.add_argument("--disable-extensions")
        chrome_options.add_argument("--disable-gpu")
        chrome_options.add_argument("--disable-logging")
        chrome_options.add_argument("--disable-web-security")
        chrome_options.add_argument("--allow-running-insecure-content")
        chrome_options.add_argument("--disable-features=VizDisplayCompositor")
        chrome_options.add_experimental_option("excludeSwitches", ["enable-automation", "enable-logging"])
        chrome_options.add_experimental_option('useAutomationExtension', False)
        chrome_options.add_experimental_option("prefs", {
            "profile.default_content_setting_values.notifications": 2
        })
        self.driver = webdriver.Chrome(options=chrome_options)
        self.driver.execute_script("Object.defineProperty(navigator, 'webdriver', {get: () => undefined})")
        self.wait = WebDriverWait(self.driver, 30)
        self.driver.maximize_window()
        print("Chrome driver initialized successfully")
    
    def manual_login_gmail(self):
        print("Opening Gmail...")
        self.driver.get("https://gmail.com")
        print("\n" + "="*60)
        print("MANUAL LOGIN REQUIRED")
        print("="*60)
        print("1. Log in to Gmail manually in the browser window")
        print("2. If you need to select a DELEGATED profile, do so now")
        print("3. Make sure you're on the correct Gmail inbox")
        print("4. Once you're on the delegated profile inbox, press Enter here...")
        print("="*60 + "\n")
        
        input("Press Enter after you've logged in and are on the DELEGATED profile inbox...")
        
        # Give some time for the page to stabilize after profile selection
        time.sleep(5)
        
        # Ensure we have a valid window handle after profile selection
        if not self.ensure_valid_window_handle():
            print("❌ No valid browser window found!")
            return False
        
        # Final verification
        if self.verify_gmail_loaded():
            print("✅ Gmail is loaded and ready for automation!")
        else:
            print("⚠️ Gmail verification failed, but continuing...")
            
        print("✅ Proceeding with email automation...")
        return True

    def wait_for_search_results_complete(self, school_name):
        """Wait for Gmail search results to fully load using robust detection"""
        print("🔄 Waiting for Gmail search results to completely load...")
        
        # Ensure valid window handle before proceeding
        if not self.ensure_valid_window_handle():
            print("❌ No valid window handle")
            return False
        
        try:
            WebDriverWait(self.driver, 15).until(
                lambda driver: ("search" in driver.current_url.lower() or 
                               "#search" in driver.current_url.lower() or
                               "q=" in driver.current_url)
            )
            print("✅ Search URL confirmed")
        except TimeoutException:
            print("⚠️ Search URL not detected")
        
        school_name_lower = school_name.lower()
        try:
            WebDriverWait(self.driver, 15).until(
                lambda driver: school_name_lower in driver.page_source.lower()
            )
            print("✅ School name found in page source")
        except TimeoutException:
            print("⚠️ School name not found in page source")
        
        time.sleep(10)
        
        return True

    def get_conversation_full_content(self):
        """Extract full conversation content including sender names from opened conversation"""
        try:
            # Wait for conversation to load
            time.sleep(4)
            
            text_parts = []
            
            # Get all text from the conversation view
            try:
                # Main conversation container
                main_content = self.driver.find_element(By.CSS_SELECTOR, "div[role='main']")
                if main_content and main_content.text:
                    text_parts.append(main_content.text)
            except:
                pass
            
            # Get sender information from email headers - enhanced selectors
            try:
                sender_elements = self.driver.find_elements(By.CSS_SELECTOR, 
                    ".go span[email], .gD, .yW span, .qu span, .f3 span, .a3s span, .cf span")
                for sender in sender_elements:
                    if sender.text and len(sender.text.strip()) > 2:
                        text_parts.append(sender.text)
            except:
                pass
            
            # Get message content
            try:
                message_bodies = self.driver.find_elements(By.CSS_SELECTOR, 
                    ".ii.gt div, .adn.ads, .Am, .ii.gt, .a3s")
                for body in message_bodies:
                    if body.text and len(body.text.strip()) > 10:
                        text_parts.append(body.text)
            except:
                pass
            
            # Get email addresses and names from headers
            try:
                headers = self.driver.find_elements(By.CSS_SELECTOR, 
                    ".hb, .g2, .go, .gn span")
                for header in headers:
                    if header.text and len(header.text.strip()) > 3:
                        text_parts.append(header.text)
            except:
                pass
            
            # Combine all text
            full_text = " ".join(text_parts)
            full_text = re.sub(r'\s+', ' ', full_text)  # Clean up spaces
            
            return full_text
            
        except Exception as e:
            print(f"⚠️ Error extracting conversation content: {str(e)}")
            return ""

    def navigate_back_to_search(self):
        """Navigate back to search results"""
        try:
            print("🔙 Navigating back to search results...")
            
            # Try browser back button first
            self.driver.back()
            time.sleep(3)
            
            # Wait for search results to load
            try:
                WebDriverWait(self.driver, 10).until(
                    lambda driver: ("search" in driver.current_url.lower() or 
                                   "q=" in driver.current_url))
                print("✅ Back to search results")
                return True
            except TimeoutException:
                print("⚠️ May not be on search results page")
                return True  # Continue anyway
                
        except Exception as e:
            print(f"⚠️ Error navigating back: {str(e)}")
            return False

    def enhanced_conversation_click(self, conversation_element):
        """Enhanced method to click on Gmail conversation rows"""
        try:
            # Method 1: Direct click
            print("   🖱️ Attempting direct click...")
            conversation_element.click()
            time.sleep(2)
            return True
        except Exception as e:
            print(f"   ⚠️ Direct click failed: {str(e)}")
        
        try:
            # Method 2: JavaScript click
            print("   🖱️ Attempting JavaScript click...")
            self.driver.execute_script("arguments[0].click();", conversation_element)
            time.sleep(2)
            return True
        except Exception as e:
            print(f"   ⚠️ JavaScript click failed: {str(e)}")
        
        try:
            # Method 3: Click on subject span
            print("   🖱️ Attempting subject span click...")
            subject_span = conversation_element.find_element(By.CSS_SELECTOR, ".bog, .y6, span")
            subject_span.click()
            time.sleep(2)
            return True
        except Exception as e:
            print(f"   ⚠️ Subject span click failed: {str(e)}")
        
        try:
            # Method 4: ActionChains click
            print("   🖱️ Attempting ActionChains click...")
            actions = ActionChains(self.driver)
            actions.move_to_element(conversation_element).click().perform()
            time.sleep(2)
            return True
        except Exception as e:
            print(f"   ⚠️ ActionChains click failed: {str(e)}")
        
        try:
            # Method 5: Click on table row if it's a TR element
            print("   🖱️ Attempting table row click...")
            if conversation_element.tag_name.lower() == 'tr':
                # Find the first clickable cell
                td_elements = conversation_element.find_elements(By.TAG_NAME, "td")
                for td in td_elements:
                    if td.is_displayed() and td.is_enabled():
                        td.click()
                        time.sleep(2)
                        return True
        except Exception as e:
            print(f"   ⚠️ Table row click failed: {str(e)}")
        
        print("   ❌ All click methods failed")
        return False

    def interactive_conversation_checker(self, conversations, school_name, director_last_name):
        """Open each conversation to check for director's last name match"""
        print(f"\n🔍 INTERACTIVE CONVERSATION CHECKER")
        print(f"   🏫 School: {school_name}")
        print(f"   👤 Looking for Director: {director_last_name}")
        print(f"   📧 Found {len(conversations)} conversations to check")
        print(f"   🤖 Will open each conversation to find director's name...")
        print(f"   ⚙️ Confirmation mode: {'ON' if self.require_confirmation else 'OFF (Auto-select)'}")
        print("-" * 60)
        
        # Store current URL to return to search
        search_url = self.driver.current_url
        
        for idx, conv in enumerate(conversations):
            try:
                print(f"\n📧 Opening conversation {idx + 1}/{len(conversations)}...")
                
                # Enhanced clicking method
                if not self.enhanced_conversation_click(conv):
                    print("❌ Failed to open conversation with all methods. Skipping...")
                    continue
                
                # Wait for conversation to load and check URL change
                time.sleep(5)
                
                # Verify we're in a conversation (URL should change)
                current_url = self.driver.current_url
                if "search" in current_url.lower() and current_url == search_url:
                    print("⚠️ Still on search page, conversation may not have opened")
                    # Try alternative approach - check if conversation view loaded
                    try:
                        WebDriverWait(self.driver, 5).until(
                            EC.presence_of_element_located((By.CSS_SELECTOR, "div[role='main'] .ii, .adn, .a3s")))
                        print("✅ Conversation content detected")
                    except TimeoutException:
                        print("❌ No conversation content detected, skipping...")
                        continue
                
                # Get full conversation content
                full_content = self.get_conversation_full_content()
                
                print(f"   📄 Preview: {full_content[:200]}...")
                
                # Check for director's last name
                director_found = director_last_name.lower() in full_content.lower()
                
                if director_found:
                    print(f"✅ MATCH FOUND! Director '{director_last_name}' found in this conversation!")
                    print(f"   📋 Content preview: {full_content[:300]}...")
                    
                    # Check confirmation preference
                    if self.require_confirmation:
                        # Ask user for confirmation
                        while True:
                            choice = input(f"\n🎯 Select this conversation for '{director_last_name}'? (y/n/s=show more): ").lower().strip()
                            
                            if choice in ['y', 'yes']:
                                print(f"✅ Conversation selected for {school_name}!")
                                return True
                            elif choice in ['n', 'no']:
                                print("   ⏭️ Continuing to next conversation...")
                                break
                            elif choice in ['s', 'show']:
                                print(f"   📄 Full content: {full_content[:800]}...")
                            else:
                                print("   Please enter 'y' (yes), 'n' (no), or 's' (show more)")
                    else:
                        # Auto-select mode - no confirmation needed
                        print(f"✅ AUTO-SELECTED: Conversation selected for {school_name}!")
                        print(f"   📋 (Confirmation disabled - using first match)")
                        return True
                        
                else:
                    print(f"   ❌ No match for director '{director_last_name}' in this conversation")
                
                # Navigate back to search results
                if not self.navigate_back_to_search():
                    print("⚠️ Could not navigate back to search results, trying direct URL...")
                    try:
                        self.driver.get(search_url)
                        time.sleep(3)
                    except:
                        print("❌ Failed to return to search results")
                        break
                    
                # Wait a bit before next conversation
                time.sleep(2)
                
            except Exception as e:
                print(f"❌ Error checking conversation {idx + 1}: {str(e)}")
                # Try to navigate back
                try:
                    self.driver.get(search_url)
                    time.sleep(3)
                except:
                    pass
                continue
        
        print(f"\n❌ No conversations found with director '{director_last_name}' after checking {len(conversations)} conversations")
        return False

    def get_conversation_text_from_search_results(self, conversation_element):
        """Enhanced text extraction to get more complete conversation preview"""
        text_parts = []
        
        # Get main element text
        try:
            elem_text = conversation_element.text
            if elem_text and elem_text.strip():
                text_parts.append(elem_text)
        except:
            pass
        
        # Get text from specific Gmail conversation elements
        try:
            subject_spans = conversation_element.find_elements(By.CSS_SELECTOR, 
                "span[id], .bog, .y6, .yW span, .y2, .aXjCH span, .bqe span")
            for span in subject_spans:
                if span.text and span.text.strip() and len(span.text) > 3:
                    text_parts.append(span.text)
        except:
            pass
        
        # Combine and clean text
        combined_text = " ".join(text_parts).strip()
        combined_text = re.sub(r'\s+', ' ', combined_text)
        combined_text = re.sub(r'[\n\r\t]+', ' ', combined_text)
        
        return combined_text

    def search_school_and_select(self, school_name, auto_select=True):
        """Enhanced search with interactive conversation checking"""
        print(f"Searching for: {school_name}")
        
        # Ensure valid window handle before proceeding
        if not self.ensure_valid_window_handle():
            print("❌ No valid browser window")
            return False
        
        # Use current Gmail URL to preserve delegated profile
        gmail_url = self.get_current_gmail_url()
        if gmail_url is None:
            print("❌ Could not get Gmail URL")
            return False
            
        print(f"📧 Navigating to: {gmail_url}")
        
        try:
            self.driver.get(gmail_url)
            time.sleep(5)
        except Exception as e:
            print(f"❌ Error navigating to Gmail: {str(e)}")
            return False
        
        search_selectors = [
            "input[aria-label*='Search mail']",
            "input[placeholder*='Search']",
            "[gh='tl'] input"
        ]
        
        search_box = None
        for selector in search_selectors:
            try:
                search_box = self.wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, selector)))
                print(f"Found search box with selector: {selector}")
                break
            except:
                continue
                
        if not search_box:
            print("❌ Could not find search box. Please search manually.")
            input("Please search for the school manually and select the conversation, then press Enter...")
            return True
        
        search_box.click()
        time.sleep(1)
        search_box.clear()
        search_box.send_keys(school_name)
        search_box.send_keys(Keys.ENTER)
        
        self.wait_for_search_results_complete(school_name)
        
        print("🔍 Looking for conversations containing the school name...")
        
        page_text = self.driver.find_element(By.TAG_NAME, "body").text.lower()
        if school_name.lower() not in page_text:
            print("❌ School name not found anywhere on the page!")
            print("This suggests the search didn't return any results.")
            input("Please manually select a conversation if one exists, then press Enter...")
            return True
        
        print("✅ School name found on the page, looking for clickable conversations...")
        
        # Enhanced conversation detection with multiple selectors
        potential_conversations = []
        
        # Method 1: Look for table rows with jsaction
        try:
            table_rows = self.driver.find_elements(By.CSS_SELECTOR, "tr[jsaction]")
            for row in table_rows:
                row_text = row.text.lower()
                if school_name.lower() in row_text:
                    potential_conversations.append(row)
                    print(f"Found table row: {row.text[:100]}...")
        except Exception as e:
            print(f"Table row search failed: {e}")
        
        # Method 2: Look for conversation containers
        try:
            conv_containers = self.driver.find_elements(By.CSS_SELECTOR, ".zA, .yW")
            for container in conv_containers:
                container_text = container.text.lower()
                if school_name.lower() in container_text:
                    potential_conversations.append(container)
                    print(f"Found container: {container.text[:100]}...")
        except Exception as e:
            print(f"Container search failed: {e}")
        
        # Method 3: XPath search
        try:
            xpath_query = f"//*[contains(translate(text(), 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), '{school_name.lower()}')]"
            xpath_elements = self.driver.find_elements(By.XPATH, xpath_query)
            
            for elem in xpath_elements:
                try:
                    # Find parent conversation row
                    current = elem
                    for _ in range(5):
                        current = current.find_element(By.XPATH, "..")
                        if current.tag_name in ['tr', 'div'] and (current.get_attribute('jsaction') or current.get_attribute('class')):
                            potential_conversations.append(current)
                            print(f"Found XPath element: {current.text[:100]}...")
                            break
                except:
                    continue
        except Exception as e:
            print(f"XPath search failed: {e}")
        
        # Remove duplicates
        unique_conversations = []
        for conv in potential_conversations:
            if conv not in unique_conversations:
                unique_conversations.append(conv)
        
        # UPDATED: Increased limit from 5 to 15
        conversations = unique_conversations[:15]  # Now checks up to 15 conversations
        
        if not conversations:
            print(f"⚠️ No conversations found containing '{school_name}'")
            print("The search may not have returned any results, or results are in an unexpected format.")
            input("Please manually select a conversation if one exists, then press Enter...")
            return True
        
        print(f"✅ Found {len(conversations)} conversations to analyze (limit: 15)")
        
        # Use interactive conversation checker
        director_last_name = getattr(self, 'current_director_name', '')
        
        if auto_select and director_last_name:
            print(f"\n🤖 SMART CONVERSATION SELECTION")
            print(f"   Since search results don't show sender names, I'll open each")
            print(f"   conversation to check for Director '{director_last_name}'")
            print(f"   ⚙️ Confirmation: {'REQUIRED' if self.require_confirmation else 'AUTO-SELECT'}")
            print(f"   🔢 Checking up to {len(conversations)} conversations")
            
            # Use interactive checker
            if self.interactive_conversation_checker(conversations, school_name, director_last_name):
                return True
            else:
                print(f"\n⚠️ No conversations found with Director '{director_last_name}'")
                print("Falling back to manual selection...")
        
        # Fallback to manual selection (also increased to show up to 15)
        print(f"\n📧 Found {len(conversations)} conversation(s) containing '{school_name}':")
        print("-" * 80)
        
        conversation_data = []
        for i, conv in enumerate(conversations):
            try:
                conv_text = self.get_conversation_text_from_search_results(conv)
                conversation_data.append({
                    'index': i,
                    'element': conv,
                    'text': conv_text[:200],
                    'score': 0
                })
                
                print(f"{i+1}. Preview: {conv_text[:200].replace(chr(10), ' ')}")
                print("-" * 40)
                
            except Exception as e:
                print(f"Error analyzing conversation {i+1}: {str(e)}")
                continue
        
        print("\n📋 Manual selection:")
        while True:
            try:
                choice = input(f"\nSelect conversation (1-{len(conversations)}) or 0 to skip: ")
                choice_num = int(choice)
                if choice_num == 0:
                    print("⏭️ Skipping this school")
                    return False
                elif 1 <= choice_num <= len(conversations):
                    if self.enhanced_conversation_click(conversations[choice_num - 1]):
                        time.sleep(3)
                        print(f"✅ Selected conversation {choice_num}")
                        return True
                    else:
                        print(f"❌ Could not click conversation {choice_num}")
                        continue
                else:
                    print(f"Please enter a number between 0 and {len(conversations)}")
            except Exception:
                print("Just select the conversation manually if needed, then press Enter...")
                input("Press Enter when ready to continue...")
                return True

    def clear_and_insert_text(self, element, text):
        """Clear all existing content and insert only new text with emoji support"""
        try:
            element.click()
            time.sleep(0.5)
            actions = ActionChains(self.driver)
            actions.key_down(Keys.CONTROL).send_keys('a').key_up(Keys.CONTROL).perform()
            time.sleep(0.5)
            actions.send_keys(Keys.DELETE).perform()
            time.sleep(0.5)
            actions.send_keys(text).perform()
            print("✅ Successfully cleared and inserted new text with ActionChains")
            return True
        except Exception as e:
            print(f"ActionChains method failed: {e}")
        try:
            script = """
            arguments[0].innerHTML = '';
            arguments[0].textContent = arguments[1];
            """
            self.driver.execute_script(script, element, text)
            print("✅ Successfully replaced content with JavaScript")
            return True
        except Exception as e:
            print(f"JavaScript method failed: {e}")
        try:
            element.click()
            time.sleep(0.5)
            element.clear()
            element.send_keys(text)
            print("✅ Successfully inserted text with direct send_keys")
            return True
        except Exception as e:
            print(f"Direct send_keys method failed: {e}")
            return False

    def schedule_email_for_10pm(self):
        """Use Gmail's native Schedule Send feature following the exact flow"""
        try:
            print("\n⏰ Using Gmail's Schedule Send for 10:00 PM today...")
            
            print("📍 Step 1: Looking for Send button dropdown...")
            
            send_dropdown_selectors = [
                "div[data-tooltip='More send options']",
                "div[aria-label='More send options']", 
                "div[data-tooltip*='send options']",
                ".T-I-J3 .J-J5-Ji",
                "div[role='button'][data-tooltip*='send options']"
            ]
            
            send_dropdown = None
            for selector in send_dropdown_selectors:
                try:
                    send_dropdown = WebDriverWait(self.driver, 10).until(
                        EC.element_to_be_clickable((By.CSS_SELECTOR, selector)))
                    print(f"✅ Found send dropdown: {selector}")
                    break
                except:
                    continue
            
            if not send_dropdown:
                print("❌ Could not find send dropdown button")
                return False
            
            if self.safe_click(send_dropdown):
                time.sleep(2)
                print("✅ Clicked send dropdown")
            else:
                print("❌ Could not click send dropdown")
                return False
            
            print("📍 Step 2: Clicking 'Schedule send' from menu...")
            
            schedule_selectors = [
                "//div[contains(text(), 'Schedule send')]",
                "//span[contains(text(), 'Schedule send')]",
                "div[role='menuitem']:contains('Schedule send')",
                ".T-I-atl:contains('Schedule send')",
                "[data-tooltip='Schedule send']"
            ]
            
            schedule_button = None
            for selector in schedule_selectors:
                try:
                    if selector.startswith("//"):
                        schedule_button = WebDriverWait(self.driver, 10).until(
                            EC.element_to_be_clickable((By.XPATH, selector)))
                    else:
                        schedule_button = WebDriverWait(self.driver, 10).until(
                            EC.element_to_be_clickable((By.CSS_SELECTOR, selector)))
                    print(f"✅ Found Schedule send button: {selector}")
                    break
                except:
                    continue
            
            if not schedule_button:
                print("❌ Could not find Schedule send button")
                return False
            
            if self.safe_click(schedule_button):
                time.sleep(3)
                print("✅ Clicked Schedule send")
            else:
                print("❌ Could not click Schedule send")
                return False
            
            print("📍 Step 3: In schedule options dialog, clicking 'Pick date & time'...")
            
            time.sleep(3)
            
            pick_date_time_selectors = [
                "//div[contains(text(), 'Pick date & time')]",
                "//span[contains(text(), 'Pick date & time')]",
                "div[role='button']:contains('Pick date & time')",
                ".T-I-atl:contains('Pick date & time')",
                "[aria-label*='Pick date']"
            ]
            
            pick_date_time_button = None
            for selector in pick_date_time_selectors:
                try:
                    if selector.startswith("//"):
                        pick_date_time_button = WebDriverWait(self.driver, 10).until(
                            EC.element_to_be_clickable((By.XPATH, selector)))
                    else:
                        pick_date_time_button = WebDriverWait(self.driver, 10).until(
                            EC.element_to_be_clickable((By.CSS_SELECTOR, selector)))
                    print(f"✅ Found 'Pick date & time': {selector}")
                    break
                except:
                    continue
            
            if pick_date_time_button:
                if self.safe_click(pick_date_time_button):
                    time.sleep(3)
                    print("✅ Clicked 'Pick date & time'")
                else:
                    print("❌ Could not click 'Pick date & time'")
                    return False
            else:
                print("⚠️ Could not find 'Pick date & time', proceeding to date/time inputs...")
            
            print("📍 Step 4: Setting date to today and time to 10:00 PM...")
            
            today_date = datetime.now().strftime('%b %d, %Y')
            print(f"Setting date to: {today_date}")
            
            date_input_selectors = [
                "input[value*='Aug 25, 2025']",
                "input[aria-label='Date']",
                "input[placeholder*='date']",
                "input[type='text'][value*='2025']",
                "//input[contains(@value, 'Aug 25')]",
                "//input[contains(@value, '2025')]"
            ]
            
            date_input = None
            for selector in date_input_selectors:
                try:
                    if selector.startswith("//"):
                        date_input = WebDriverWait(self.driver, 10).until(
                            EC.element_to_be_clickable((By.XPATH, selector)))
                    else:
                        date_input = WebDriverWait(self.driver, 10).until(
                            EC.element_to_be_clickable((By.CSS_SELECTOR, selector)))
                    print(f"✅ Found date input: {selector}")
                    break
                except:
                    continue
            
            if date_input:
                try:
                    self.safe_click(date_input)
                    time.sleep(0.5)
                    ActionChains(self.driver).key_down(Keys.CONTROL).send_keys('a').key_up(Keys.CONTROL).perform()
                    time.sleep(0.5)
                    ActionChains(self.driver).send_keys(today_date).perform()
                    print(f"✅ Set date to: {today_date}")
                    time.sleep(1)
                except Exception as e:
                    print(f"⚠️ Could not set date: {str(e)}")
            else:
                print("⚠️ Could not find date input field")
            
            print("Setting time to: 10:00 PM")
            
            time_input_selectors = [
                "input[value*='10:00 PM']",
                "input[aria-label='Time']",
                "input[placeholder*='time']",
                "input[type='text'][value*='PM']",
                "//input[contains(@value, '10:00')]",
                "//input[contains(@value, 'PM')]"
            ]
            
            time_input = None
            for selector in time_input_selectors:
                try:
                    if selector.startswith("//"):
                        time_input = WebDriverWait(self.driver, 10).until(
                            EC.element_to_be_clickable((By.XPATH, selector)))
                    else:
                        time_input = WebDriverWait(self.driver, 10).until(
                            EC.element_to_be_clickable((By.CSS_SELECTOR, selector)))
                    print(f"✅ Found time input: {selector}")
                    break
                except:
                    continue
            
            if time_input:
                try:
                    self.safe_click(time_input)
                    time.sleep(0.5)
                    ActionChains(self.driver).key_down(Keys.CONTROL).send_keys('a').key_up(Keys.CONTROL).perform()
                    time.sleep(0.5)
                    ActionChains(self.driver).send_keys("10:00 PM").perform()
                    print("✅ Set time to: 10:00 PM")
                    time.sleep(1)
                except Exception as e:
                    print(f"⚠️ Could not set time: {str(e)}")
            else:
                print("⚠️ Could not find time input field")
            
            print("📍 Step 5: Clicking final 'Schedule send' to confirm...")
            
            time.sleep(2)
            
            final_schedule_selectors = [
                "//span[text()='Schedule send']",
                "//button[contains(text(), 'Schedule send')]", 
                "button:contains('Schedule send')",
                "[aria-label*='Schedule send'][role='button']",
                ".T-I-atl:contains('Schedule send')"
            ]
            
            final_button = None
            for selector in final_schedule_selectors:
                try:
                    if selector.startswith("//"):
                        final_button = WebDriverWait(self.driver, 10).until(
                            EC.element_to_be_clickable((By.XPATH, selector)))
                    else:
                        final_button = WebDriverWait(self.driver, 10).until(
                            EC.element_to_be_clickable((By.CSS_SELECTOR, selector)))
                    print(f"✅ Found final Schedule send button: {selector}")
                    break
                except:
                    continue
            
            if final_button:
                if self.safe_click(final_button):
                    time.sleep(3)
                    print("🎉 Email successfully scheduled for 10:00 PM today!")
                    return True
                else:
                    print("❌ Could not click final Schedule send button")
                    return False
            else:
                print("❌ Could not find final Schedule send button")
                return False
            
        except Exception as e:
            print(f"❌ Error in schedule process: {str(e)}")
            return False

    def reply_to_message(self, director_last_name, cc_emails=None, schedule_send=True):
        """Reply to the current message thread with optional scheduling"""
        try:
            print("Looking for reply button...")
            reply_selectors = [
                "[aria-label*='Reply'][role='button']",
                "[data-tooltip*='Reply']",
                ".ams.bkH",
                "[title*='Reply']"
            ]
            reply_button = None
            for selector in reply_selectors:
                try:
                    reply_button = self.wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, selector)))
                    print(f"Found reply button with selector: {selector}")
                    break
                except:
                    continue
            if not reply_button:
                print("❌ Could not find reply button automatically.")
                input("Please click the Reply button manually, then press Enter...")
            else:
                reply_button.click()
                time.sleep(3)
            
            if cc_emails and cc_emails.strip():
                print(f"\n📧 Adding CC: {cc_emails}")
                try:
                    actions = ActionChains(self.driver)
                    actions.key_down(Keys.CONTROL).key_down(Keys.SHIFT).send_keys('c').key_up(Keys.SHIFT).key_up(Keys.CONTROL).perform()
                    time.sleep(2)
                    
                    try:
                        active_element = self.driver.switch_to.active_element
                        active_element.clear()
                        active_element.send_keys(cc_emails)
                        print(f"✅ CC filled: {cc_emails}")
                    except:
                        print("⚠️ Could not auto-fill CC")
                except Exception as e:
                    print(f"Error with CC: {str(e)}")
            
            print("\n📝 Now filling email body...")
            time.sleep(1)
            
            # NEW: Set font to Arial before inserting text
            self.set_email_body_font_arial()
            
            email_body = f"""Hi Director {director_last_name},
            
test sending hello world
maron"""
            
            body_selectors = [
                "div[aria-label*='Message Body']",
                "div[role='textbox'][aria-label*='Message']",
                ".Am.Al.editable",
                "div[contenteditable='true'][aria-label*='Message']",
                ".editable[role='textbox']"
            ]
            
            message_body = None
            for selector in body_selectors:
                try:
                    message_body = self.wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, selector)))
                    print(f"Found message body with selector: {selector}")
                    break
                except:
                    continue
            
            if message_body:
                try:
                    if self.clear_and_insert_text(message_body, email_body):
                        print("✅ Email body replaced successfully with Arial font!")
                    else:
                        print("⚠️ Could not replace body content")
                except Exception as e:
                    print(f"Error filling body: {str(e)}")
            
            if schedule_send:
                return self.schedule_email_for_10pm()
            else:
                print("\n📤 Now attempting to send email immediately...")
                time.sleep(2)
                
                send_selectors = [
                    "[aria-label*='Send '][role='button']",
                    "[data-tooltip*='Send']",
                    ".T-I.J-J5-Ji.aoO.v7.T-I-atl.L3",
                    "[role='button'][aria-label*='Send']"
                ]
                
                send_button = None
                for selector in send_selectors:
                    try:
                        send_button = WebDriverWait(self.driver, 5).until(EC.element_to_be_clickable((By.CSS_SELECTOR, selector)))
                        print(f"Found send button with selector: {selector}")
                        break
                    except:
                        continue
                
                if send_button:
                    try:
                        send_button.click()
                        time.sleep(3)
                        print("✅ Email sent immediately!")
                    except Exception as e:
                        print(f"Error clicking send: {str(e)}")
                else:
                    print("⚠️ Could not find send button automatically")
            
            return True
            
        except Exception as e:
            print(f"Error in reply_to_message: {str(e)}")
            return False

    def send_email(self):
        print("Send process already handled in reply_to_message")
        return True

    def update_excel_status(self, row_index, status, success_status):
        """Update Excel with proper date formatting and 14-day gap"""
        today = datetime.now()
        
        self.df.loc[row_index, 'Date of Last Action'] = today
        self.df.loc[row_index, 'succcessful/Failed'] = success_status
        self.df.loc[row_index, 'Status'] = status
        
        next_due_date = today + timedelta(days=14)
        self.df.loc[row_index, 'Next Action Due Date'] = next_due_date
        self.df.loc[row_index, 'Next action'] = 'Follow-up Email'
        
        try:
            with pd.ExcelWriter(self.excel_file_path, 
                               engine='xlsxwriter',
                               date_format='DD/MM/YYYY',
                               datetime_format='DD/MM/YYYY HH:MM:SS') as writer:
                
                self.df.to_excel(writer, sheet_name='Sheet1', index=False)
                
                workbook = writer.book
                worksheet = writer.sheets['Sheet1']
                date_format = workbook.add_format({'num_format': 'dd/mm/yyyy'})
                worksheet.set_column('I:J', 12, date_format)
                
                print(f"✅ Updated Excel file with proper date formatting")
                print(f"   Date of Last Action: {today.strftime('%d/%m/%Y')}")
                print(f"   Next Action Due Date: {next_due_date.strftime('%d/%m/%Y')} (14 days from today)")
                print(f"   Status: {status}")
                print(f"   Success/Failed: {success_status}")
                
        except Exception as e:
            print(f"⚠️ Could not save with xlsxwriter, falling back to default method: {str(e)}")
            self.df.loc[row_index, 'Date of Last Action'] = today.strftime('%d/%m/%Y')
            self.df.loc[row_index, 'Next Action Due Date'] = next_due_date.strftime('%d/%m/%Y')
            
            self.df.to_excel(self.excel_file_path, index=False)
            print(f"✅ Updated Excel file with status: {status}")

    def process_contacts(self, start_index=0, max_emails=None, auto_select_schools=True, schedule_emails=True):
        """Process contacts from Excel file and send automated emails"""
        if not self.load_excel_data():
            return
        self.setup_driver()
        
        # If manual login fails, return early
        if not self.manual_login_gmail():
            print("❌ Login failed, exiting...")
            return
            
        contacts_to_process = self.df.copy()
        if contacts_to_process.empty:
            print("No contacts to process.")
            return
        processed_count = 0
        for idx, contact in contacts_to_process.iterrows():
            if max_emails and processed_count >= max_emails:
                break
            if processed_count < start_index:
                processed_count += 1
                continue
            school_name = contact['School Name']
            director_name = contact['Last Name']
            email = contact['Email']
            cc_emails = str(contact['CC']) if pd.notna(contact['CC']) else ""
            
            # Store current director name for interactive checking
            self.current_director_name = director_name
            
            print(f"\n{'='*60}")
            print(f"Processing {processed_count + 1}/{len(contacts_to_process)}: {school_name}")
            print(f"Director: {director_name}")
            print(f"Email: {email}")
            print(f"CC: {cc_emails}")
            print(f"Auto-select: {'ON' if auto_select_schools else 'OFF'}")
            print(f"Schedule for 10 PM: {'ON' if schedule_emails else 'OFF'}")
            print(f"Confirmation: {'REQUIRED' if self.require_confirmation else 'AUTO-SELECT'}")
            print(f"Search limit: 15 conversations")
            print('='*60)
            
            try:
                if self.search_school_and_select(school_name, auto_select=auto_select_schools):
                    if self.reply_to_message(director_name, cc_emails, schedule_send=schedule_emails):
                        if schedule_emails:
                            self.update_excel_status(idx, 'Follow-up Email Scheduled for 10 PM', 'Scheduled')
                            print(f"✅ Follow-up scheduled for {school_name}")
                        else:
                            self.update_excel_status(idx, 'Follow-up Email Sent', 'Successful')
                            print(f"✅ Follow-up processed for {school_name}")
                    else:
                        print(f"❌ Could not process reply for {school_name}")
                        self.update_excel_status(idx, 'Follow-up Failed', 'Failed')
                else:
                    print(f"⏭️ Skipped {school_name}")
                    self.update_excel_status(idx, 'No Conversation Selected', 'Skipped')
            except Exception as e:
                print(f"❌ Error processing {school_name}: {str(e)}")
                self.update_excel_status(idx, 'Processing Error', 'Failed')
            
            processed_count += 1
            
            if processed_count < len(contacts_to_process):
                print("Waiting 5 seconds before next email...")
                time.sleep(5)
        print(f"\n✅ Processed {processed_count} contacts")
        
        if schedule_emails:
            print("\n🎉 All emails have been scheduled using Gmail's native scheduling!")
            print("📧 Gmail will automatically send them at 10:00 PM today.")
            print("📋 You can view/modify scheduled emails in Gmail's 'Scheduled' folder.")

    def close_driver(self):
        if self.driver:
            self.driver.quit()

def main():
    EXCEL_FILE_PATH = "Main_Filtered_Schools_Formatted.xlsx"
    gmail_bot = GmailAutomationWithExcel(EXCEL_FILE_PATH, headless=False)
    try:
        print("🚀 Starting Gmail Automation with Enhanced Director Matching")
        print("="*60)
        
        # Confirmation preference setting at the beginning
        print("\n⚙️ CONFIRMATION PREFERENCES")
        print("When a director name match is found in a conversation:")
        print("  • Manual Mode: Ask for your confirmation before selecting")
        print("  • Auto Mode: Automatically select the first match found")
        
        while True:
            confirmation_choice = input("\nRequire confirmation for director matches? (y=manual, n=auto): ").lower().strip()
            if confirmation_choice in ['y', 'yes', 'manual', 'm']:
                gmail_bot.require_confirmation = True
                print("✅ Manual confirmation mode enabled")
                break
            elif confirmation_choice in ['n', 'no', 'auto', 'a']:
                gmail_bot.require_confirmation = False
                print("✅ Auto-select mode enabled (no confirmation needed)")
                break
            else:
                print("Please enter 'y' for manual confirmation or 'n' for auto-select")
        
        schedule_choice = input("\nSchedule emails for 10:00 PM today using Gmail's Schedule Send? (y/n): ").lower()
        schedule_emails = schedule_choice in ['y', 'yes']
        
        print(f"\n🎯 AUTOMATION SETTINGS:")
        print(f"   📧 Scheduling: {'10:00 PM Today' if schedule_emails else 'Send Immediately'}")
        print(f"   🤖 Auto-select: ON")
        print(f"   ✅ Confirmation: {'MANUAL' if gmail_bot.require_confirmation else 'AUTO'}")
        print(f"   🔢 Search limit: 15 conversations (increased from 5)")
        print(f"   🎨 Font: Arial 11px")
        print("-" * 60)
        
        gmail_bot.process_contacts(
            start_index=0, 
            max_emails=None, 
            auto_select_schools=True,
            schedule_emails=schedule_emails
        )
        
        input("Press Enter to close the browser...")
            
    except Exception as e:
        print(f"An error occurred: {str(e)}")
    finally:
        gmail_bot.close_driver()

if __name__ == "__main__":
    main()

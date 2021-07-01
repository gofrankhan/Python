from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from tkinter import messagebox
from global_instance import *

class MySelenium():

    def __init__(self, *args, **kwargs):
        self.driver = webdriver.Chrome(executable_path= my_path + "driver\chromedriver_win32\chromedriver.exe")

    def openurl(self, url, browser):
        self.driver.get(url)

    def click_by_xpath(self, xpath):
        self.driver.find_element_by_xpath(xpath).click()
    
    def click_by_id(self, id):
        self.driver.find_element_by_id(id).click()

    def click_by_name(self, name):
        self.driver.find_element_by_name(name).click()

    def click_by_tag_name(self, tag_name):
        self.driver.find_element_by_tag_name(tag_name).click()

    def click_by_class_name(self, class_name):
        self.driver.find_element_by_class_name(class_name).click()

    def click_by_css_selector(self, css_selector):
        self.driver.find_element_by_css_selector(css_selector).click()
    
    def click_by_link_text(self, link_text):
        self.driver.find_element_by_link_text(link_text).click()
    
    def click_by_partial_link_text(self, partial_link_text):
        self.driver.find_element_by_partial_link_text(partial_link_text).click()

    def readtext_by_xpath(self, xpath, save_to):
        self.text = self.driver.find_element_by_xpath(xpath).text
        variable_dict[save_to] = self.text
    
    def readtext_by_id(self, id):
        self.text = self.driver.find_element_by_id(id).text

    def readtext_by_name(self, name):
        self.text = self.driver.find_element_by_name(name).text

    def readtext_by_tag_name(self, tag_name):
        self.text = self.driver.find_element_by_tag_name(tag_name).text

    def readtext_by_class_name(self, class_name):
        self.text = self.driver.find_element_by_class_name(class_name).text

    def readtext_by_css_selector(self, css_selector):
        self.text = self.driver.find_element_by_css_selector(css_selector).text
    
    def readtext_by_link_text(self, link_text):
        self.text = self.driver.find_element_by_link_text(link_text).text
    
    def readtext_by_partial_link_text(self, partial_link_text):
        self.text = self.driver.find_element_by_partial_link_text(partial_link_text).text

    def inputtext_by_xpath(self, xpath, value):
        self.driver.find_element_by_xpath(xpath).send_keys(value)
        print(value)
    
    def inputtext_by_id(self, id, value):
        self.driver.find_element_by_id(id).send_keys(value)

    def inputtext_by_name(self, name, value):
        self.driver.find_element_by_name(name).send_keys(value)

    def inputtext_by_tag_name(self, tag_name, value):
        self.driver.find_element_by_tag_name(tag_name).send_keys(value)

    def inputtext_by_class_name(self, class_name, value):
        self.driver.find_element_by_class_name(class_name).send_keys(value)

    def inputtext_by_css_selector(self, css_selector, value):
        self.driver.find_element_by_css_selector(css_selector).send_keys(value)
    
    def inputtext_by_link_text(self, link_text, value):
        self.driver.find_element_by_link_text(link_text).send_keys(value)
    
    def inputtext_by_partial_link_text(self, partial_link_text, value):
        self.driver.find_element_by_partial_link_text(partial_link_text).send_keys(value)

    def clear_by_xpath(self, xpath):
        self.driver.find_element_by_xpath(xpath).clear()
    
    def clear_by_id(self, id):
        self.driver.find_element_by_id(id).clear()

    def clear_by_name(self, name):
        self.driver.find_element_by_name(name).clear()

    def clear_by_tag_name(self, tag_name):
        self.driver.find_element_by_tag_name(tag_name).clear()

    def clear_by_class_name(self, class_name):
        self.driver.find_element_by_class_name(class_name).clear()

    def clear_by_css_selector(self, css_selector):
        self.driver.find_element_by_css_selector(css_selector).clear()
    
    def clear_by_link_text(self, link_text):
        self.driver.find_element_by_link_text(link_text).clear()
    
    def clear_by_partial_link_text(self, partial_link_text):
        self.driver.find_element_by_partial_link_text(partial_link_text).clear()
    
    def send_keys_by_xpath(self, xpath, key):
        self.driver.find_element_by_xpath(xpath).send_keys(key)
    
    def send_keys_by_id(self, id, key):
        self.driver.find_element_by_id(id).send_keys(key)

    def send_keys_by_name(self, name, key):
        self.driver.find_element_by_name(name).send_keys(key)

    def send_keys_by_tag_name(self, tag_name, key):
        self.driver.find_element_by_tag_name(tag_name).send_keys(key)

    def send_keys_by_class_name(self, class_name, key):
        self.driver.find_element_by_class_name(class_name).send_keys(key)

    def send_keys_by_css_selector(self, css_selector, key):
        self.driver.find_element_by_css_selector(css_selector).send_keys(key)
    
    def send_keys_by_link_text(self, link_text, key):
        self.driver.find_element_by_link_text(link_text).send_keys(key)
    
    def send_keys_by_partial_link_text(self, partial_link_text, key):
        self.driver.find_element_by_partial_link_text(partial_link_text).send_keys(key)

    def my_messagebox(self, variable_name):
        messagebox.showinfo("Information", variable_dict[variable_name])

    def wait(self, seconds):
        self.driver.implicitly_wait(seconds)

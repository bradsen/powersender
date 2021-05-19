import pandas as pd
import os, sys
from time import sleep
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from pynput.keyboard import Key, Controller
import pyperclip
import requests
import wget
import re
import os
import zipfile

def download_chromedriver():
    filePath = os.getcwd()
    def get_latestversion(version):
        url = 'https://chromedriver.storage.googleapis.com/LATEST_RELEASE_' + str(version)
        response = requests.get(url)
        version_number = response.text
        return version_number
    def download(download_url, driver_binaryname, target_name):
        # download the zip file using the url built above
        latest_driver_zip = wget.download(download_url, out='chromedriver.zip')

        # extract the zip file
        with zipfile.ZipFile(latest_driver_zip, 'r') as zip_ref:
            zip_ref.extractall(path=filePath)  # you can specify the destination folder path here

        # delete the zip file downloaded above
        os.remove(latest_driver_zip)
    if os.name == 'nt':
        replies = os.popen(r'reg query "HKEY_CURRENT_USER\Software\Google\Chrome\BLBeacon" /v version').read()
        replies = replies.split('\n')
        for reply in replies:
            if 'version' in reply:
                reply = reply.rstrip()
                reply = reply.lstrip()
                tokens = re.split(r"\s+", reply)
                fullversion = tokens[len(tokens) - 1]
                tokens = fullversion.split('.')
                version = tokens[0]
                break

        #target_name = './bin/chromedriver-win-' + version + '.exe'
        target_name = filePath + '\chromedriver.exe'
        
        found = os.path.exists(target_name)
       
        if not found:
            version_number = get_latestversion(version)
            # build the donwload url
            download_url = "https://chromedriver.storage.googleapis.com/" + version_number +"/chromedriver_win32.zip"
            #download(download_url, './temp/chromedriver.exe', target_name)
            download(download_url, filePath, target_name)

    elif os.name == 'posix':
        reply = os.popen(r'chromium --version').read()

        if reply != '':
            reply = reply.rstrip()
            reply = reply.lstrip()
            tokens = re.split(r"\s+", reply)
            fullversion = tokens[1]
            tokens = fullversion.split('.')
            version = tokens[0]
        else:
            reply = os.popen(r'google-chrome --version').read()
            reply = reply.rstrip()
            reply = reply.lstrip()
            tokens = re.split(r"\s+", reply)
            fullversion = tokens[2]
            tokens = fullversion.split('.')
            version = tokens[0]

        target_name = filePath + '\chromedriver.exe'
        print('new chrome driver at ' + target_name)
        found = os.path.exists(target_name)
        if not found:
            version_number = get_latestversion(version)
            download_url = "https://chromedriver.storage.googleapis.com/" + version_number +"/chromedriver_linux64.zip"
            #download(download_url, './temp/chromedriver', target_name)
            download(download_url, filePath, target_name)

download_chromedriver()

if getattr(sys, 'frozen', False):
    chromedriver_path = os.path.join(sys._MEIPASS, "chromedriver.exe")
    driver = webdriver.Chrome(chromedriver_path)
else:
    driver = webdriver.Chrome()

driver.get('https://web.whatsapp.com/')

line = "----------------------------------------------------------"

class fileValidityTest():

    def testValidity(self):

        a = True
        while a:

            try:
                prompt = 'Please input a csv filename (Note:Include file extension): '
                filename = input(prompt)
                f = open(filename)
            except FileNotFoundError:
                print("File not found, please try again!")

            else:
                try:
                    split_ext = filename.split('.')[1]
                    if split_ext == "csv":
                        validFile = open(filename)
                        print("File is valid....")

                        a = False

                    if split_ext != "csv":
                        raise Exception
                except Exception as e:
                    print("Invalid File Extension")

        return validFile, filename

    def testImageValidity(self):
        a = 'File not found'
        while a:
            a = True

            try:
                prompt = 'Please input a image filename (Note:Only jpg and png are accepted): '
                filename = input(prompt)
                f = open(filename)

            except FileNotFoundError:
                print("File not found, please try again!")
            else:
                try:
                    split_ext = filename.split('.')[1]
                    if split_ext == "jpg" or split_ext == "png" or split_ext == "jpeg":
                        validImageFile = open(filename)
                        print("File is valid....")

                        a = False

                    if split_ext != "png" and split_ext != "jpg" and split_ext != "jpeg":
                        raise Exception
                except Exception as e:
                    print("Invalid File Extension")

        return validImageFile, filename

    def testAttachmentValidity(self):
        a = 'File not found'
        while a:
            a = True
            try:
                prompt = 'Please input an attachment (Note:Only PDF is accepted): '
                filename = input(prompt)
                f = open(filename)

            except FileNotFoundError:
                print("File not found, please try again!")

            else:

                try:
                    split_ext = filename.split('.')[1]
                    if split_ext == "pdf":
                        validAttachmentFile = open(filename)
                        print("File is valid....")

                        a = False

                    if split_ext != "pdf":
                        raise Exception
                except Exception as e:
                    print("Invalid File Extension")

        return validAttachmentFile, filename

    def testTxtFileValidity(self):
        a = 'File not found'
        while a:
            a = True

            try:
                prompt = 'Please input a template file (Note:Only .txt is accepted): '
                filename = input(prompt)
                f = open(filename)

            except FileNotFoundError:
                print("File not found, please try again!")

            else:
                try:
                    split_ext = filename.split('.')[1]
                    if split_ext == "txt":
                        validTxtFile = open(filename)
                        print("File is valid....")

                        a = False

                    if split_ext != "txt":
                        raise Exception
                except Exception as e:
                    print("Invalid File Extension")

        return validTxtFile, filename


class sendMessage:
    def __init__(self, group, validTemplate, validImageFile, validAttachment, custNamelist):
        self.group = group
        self.custList = custNamelist
        self.validImageFile = validImageFile
        self.validAttachment = validAttachment
        self.validTemplate = validTemplate

    def sendImage(self):
        source = 'https://www.straitstimes.com/singapore/singapore-residents-who-continue-to-travel-will-pay-full-hospital-charges-if-warded-for'

        # Click attachment button
        attachment = driver.find_element_by_xpath('//div[@title = "Attach"]')
        attachment.click()
        sleep(1)

        sendDoc = '//input[@accept = "image/*,video/mp4,video/3gpp,video/quicktime"]'
        inputElement5 = driver.find_element_by_xpath(sendDoc)
        path = os.getcwd()
        attachment = os.path.join(path, self.validImageFile)
        inputElement5.send_keys(attachment)

        '''
        # Fill in text for image
        sleep(1)
        fillText = '//*[@id="app"]/div/div/div[2]/div[2]/span/div/span/div/div/div[2]/div[1]/span/div/div[2]/div/div[3]/div[1]/div[2]'
        content = driver.find_element_by_xpath(fillText)
        content.send_keys('Source: {}'.format(source))
        '''

        # Find and click button
        send_button_path = '//span[@data-icon="send"]'
        sleep(3)
        send_button = driver.find_element_by_xpath(send_button_path)
        send_button.click()

    def sendAttachment(self):
        # Click attachment button
        attachment = driver.find_element_by_xpath('//div[@title = "Attach"]')
        attachment.click()
        sleep(1)

        # Find document icon
        document = '//input[@accept = "*"]'
        inputElement4 = driver.find_element_by_xpath(document)
        attachmentPath = os.getcwd()
        filePath = os.path.join(attachmentPath, self.validAttachment)
        inputElement4.send_keys(filePath)

        # Send send button
        send_button_path = '//span[@data-icon="send"]'
        sleep(4)
        send_button = driver.find_element_by_xpath(send_button_path)
        send_button.click()

    def sendCustomText(self):
        keyboard = Controller()
        data = pd.read_csv(self.custList, index_col=False)
        image = "{i}"
        attachments = "{a}"
        name = "{name}"

        for ind, row in data.iterrows():
            if (row['Group'] == self.group):

                # filtering criteria
                phoneTest = row['Phone Number']
                custTest = row['Name']

                try:
                    inputElement = driver.find_element_by_xpath('//*[@id="side"]/div[1]/div/label/div/div[2]')
                except:
                    print("Unable to detect WhatsApp Web, please try again!")
                    a = False
                    break

                phonenumber2 = str(phoneTest)
                custName2 = str(custTest)
                inputElement.send_keys(phonenumber2)
                inputElement.send_keys(Keys.ENTER)

                sleep(2)
                bands = list()

                # read in bands from file
                with open(self.validTemplate) as fin:
                    for line in fin:
                        if not line.startswith('#'):
                            bands.append(line.strip())

                for count in range(0, len(bands)):
                    string = bands[count]
                    if name in string:
                        keyboard.press(Key.backspace)
                        keyboard.release(Key.backspace)
                        keyboard.press(Key.space)
                        keyboard.release(Key.space)
                        keyboard.type(custName2)
                        continue
                        keyboard.press(Key.backspace)
                        keyboard.release(Key.backspace)

                    if image in string:
                        keyboard.press(Key.enter)
                        keyboard.release(Key.enter)

                        if (len(self.validImageFile) == 0):
                            continue
                        else:
                            keyboard.press(Key.enter)
                            keyboard.release(Key.enter)
                            # Click attachment button
                            self.sendImage()
                            continue

                    if attachments in bands[count]:
                        if (len(self.validAttachment) == 0):
                            continue
                        else:
                            keyboard.press(Key.enter)
                            keyboard.release(Key.enter)
                            # Click attachment button
                            self.sendAttachment()
                            continue

                    else:
                        # print(bands[count])
                        keyboard.type(string)
                        keyboard.press(Key.shift)
                        keyboard.press(Key.enter)
                        keyboard.release(Key.enter)
                        keyboard.release(Key.shift)
                    sleep(0.5)
                keyboard.press(Key.enter)
                keyboard.release(Key.enter)



def scanWhatsAppWeb():
    a = True
    while a:
        selectedOption = input("Have you scanned WhatsApp Web, has it finished loading? [Y/N]: ")
        if (selectedOption == "Y" or selectedOption == "y"):
            print("Good job!")
            a = False
        elif (selectedOption == "N" or selectedOption == "n"):
            print("Please wait for it to finish loading before you continue :)")
        else:
            print("Invalid option. Please try again!")


def mainMenu(custList, catName, imageList, attachFilename, templateFilename):
    print(line)
    print('Options available:')
    print(line)
    print('1) Key in the customer list   :   ', end="")
    print(custList)
    print('2) Please select the category :   ', end="")
    print(catName)
    print('3) Key in the image file      :   ', end="")
    print(imageList)
    print('4) Key in an attachment file  :   ', end="")
    print(attachFilename)
    print('5) Please choose template file:   ', end="")
    print(templateFilename)
    print('6) Clear all')
    print('7) Run Program')
    print('8) Exit program')
    print(line)


def main():

    os.system('color 7')
    # Ask user whether they have scanned WhatsApp Web
    scanWhatsAppWeb()

    custList = "test.csv"
    imageList = "test.jpg"
    catName = "Client"
    attachFilename = "test.pdf"
    templateFilename = "test.txt"



    while True:

        mainMenu(custList, catName, imageList, attachFilename, templateFilename)

        selectedOption = input("Please choose your option: ")

        # csv file
        if (selectedOption == '1'):

            a = True
            while a:
                print("The available options are:")
                print("\t1) Key in csv file ")
                print("\t2) Clear csv file ")
                print("\t3) Back to main menu ")
                chOption = input('Please select an option (0-3): ')

                if (chOption == '1'):

                    custListValid, filename = fileValidityTest().testValidity()
                    custList = filename

                elif (chOption == '2'):

                    custList = ""

                elif (chOption == '3'):

                    a = False

        # category
        elif (selectedOption == '2'):
            a = True
            while a:
                print('The categories available are:')
                print('\t1) Client')
                print('\t2) Prospect')
                print('\t3) Friend')
                print('\t4) Student')
                print('\t5) Back to Main Menu')

                catOption = input('Please select a category: ')
                if (catOption == '1'):
                    catName = "Client"

                elif (catOption == '2'):
                    catName = "Prospect"

                elif (catOption == '3'):
                    catName = "Friend"

                elif (catOption == '4'):
                    catName = "Student"

                elif (catOption == '5'):
                    a = False

        # image file
        elif (selectedOption == '3'):

            a = True
            while a:
                print("The available options are:")
                print("\t1) Key in image file ")
                print("\t2) Clear image file ")
                print("\t3) Back to main menu ")
                chOption = input('Please select an option (0-3): ')
                if (chOption == '1'):

                    validImageFile, filename = fileValidityTest().testImageValidity()
                    imageList = filename

                elif (chOption=='2'):

                    imageList = ""

                elif(chOption=='3'):

                    a = False

        # attachment file
        elif (selectedOption == '4'):
            a = True
            while a:
                print("The available options are:")
                print("\t1) Key in attachment file ")
                print("\t2) Clear attachment file ")
                print("\t3) Back to main menu ")
                chOption = input('Please select an option (0-3): ')

                if (chOption == '1'):

                    validAttachmentFile, filename = fileValidityTest().testAttachmentValidity()
                    attachFilename = filename

                elif (chOption == '2'):

                    attachFilename = ""

                elif (chOption == '3'):

                    a = False



        # template file
        elif (selectedOption == '5'):

            a = True
            while a:
                print("The available options are:")
                print("\t1) Key in template file ")
                print("\t2) Clear template file ")
                print("\t3) Back to main menu ")
                chOption = input('Please select an option (0-3): ')

                if (chOption == '1'):

                    validTxtFile, filename = fileValidityTest().testTxtFileValidity()
                    templateFilename = filename

                elif (chOption == '2'):

                    templateFilename = ""

                elif (chOption == '3'):

                    a = False

        elif(selectedOption == '6'):

            custList = ""
            imageList = ""
            catName = ""
            attachFilename = ""
            templateFilename = ""

        # run program
        elif (selectedOption == '7'):

            a = True
            while a:
                # check to make sure customer list, template and category cannot be empty
                backoption = ''
                if (len(custList) == 0):
                    print("You customer file cannot be empty")
                    backoption = input("Press B to go back to main menu: ")
                    if (backoption == 'B' or backoption == 'b'):
                        a = False

                elif (len(templateFilename) == 0):
                    print("You template filename cannot be empty")
                    backoption = input("Press B to go back to main menu: ")
                    if (backoption == 'B' or backoption == 'b'):
                        a = False
                elif (len(catName) == 0):
                    print("You category cannot be empty")
                    backoption = input("Press B to go back to main menu: ")
                    if (backoption == 'B' or backoption == 'b'):
                        a = False
                else:

                    s = sendMessage(catName, templateFilename, imageList, attachFilename, custList)
                    sleep(1)
                    s.sendCustomText()
                    # s.copyPasteText()
                    a = False

        elif (selectedOption == '8'):
            print('Program existing...')
            return False


if __name__ == "__main__":
    main()


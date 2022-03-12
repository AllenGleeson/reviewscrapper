from telnetlib import TLS
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
import sys
import xlsxwriter
from pymongo import MongoClient


class CompanyScraper:

    def __init__(self, driver, reviewCol, companySummarysCol, reviewSheet, companySheet):
        self.driver = driver
        self.reviewCol = reviewCol
        self.companySummarysCol = companySummarysCol
        self.reviewSheet = reviewSheet
        self.companySheet = companySheet


    def getCompanySummary(self):
        # Get company details 
        companyName = self.driver.find_element(By.CLASS_NAME, "styles_displayName__GElWn").text
        companyRating = self.driver.find_element(By.XPATH, "/html/body/div[1]/div/main/div/div[3]/div[2]/div/div/div/section[1]/div[1]/div[2]/div/div/p").text
        companyReviewCount = self.driver.find_element(By.CLASS_NAME, "styles_text__W4hWi").text.split(" ", 1)[0]

        # Set data to dict
        companySummary = {
                "companyName": companyName,
                "companyRating": companyRating,
                "companyReviewCount": companyReviewCount,
            }

        self.writeCompanySummaryToXL(companySummary)
        self.insertCompanySummaryToDB(companySummary)


    def writeCompanySummaryToXL(self, companySummary):
        # Write company summary to excel
        row = 1    
        column = 0

        for item in companySummary:   
            self.companySheet.write(row, column, companySummary[item])  
            column += 1


    def writeReviewToXL(self, count, review):
        # Write review to excel
        row = count+1    
        column = 0

        for item in review:    
            self.reviewSheet.write(row, column, review[item])   
            column += 1


    def getReviews(self):
        # Gets all reviews
        reviews = self.driver.find_elements(By.CLASS_NAME, "styles_reviewCard__hcAvl")
        count = 0
        
        for review in reviews:
            # Set review data
            consumor_name = review.find_element(By.CLASS_NAME, "styles_consumerName__dP8Um").text
            numberOfReviews = review.find_elements(By.CLASS_NAME, "styles_consumerExtraDetails__fxS4S")[0].text.split(" ", 1)[0]
            location = review.find_elements(By.CLASS_NAME, "styles_consumerExtraDetails__fxS4S")[0].find_elements(By.CLASS_NAME, "styles_detailsIcon__Fo_ua")[1]
            date = review.find_element(By.CLASS_NAME, "styles_datesWrapper__RCEKH").find_element(By.TAG_NAME, 'time').get_attribute("title")
            rating = review.find_element(By.CLASS_NAME, "styles_reviewHeader__iU9Px").get_attribute("data-service-review-rating")
            headline = review.find_element(By.CLASS_NAME, "styles_linkwrapper__73Tdy").text
            review_text = review.find_element(By.CLASS_NAME, "styles_reviewContent__0Q2Tg").text
            companyName = self.driver.find_element(By.CLASS_NAME, "styles_displayName__GElWn").text

            try:
                companyResponded = review.find_element(By.CLASS_NAME, "styles_replyInfo__FYSje")
                if companyResponded:
                    companyResponded = "True"
            except:
                companyResponded = "False"

            if not location:
                location = "Unknown"
            else:
                location = location.text

            # Set data to dict
            rev = {
            "consumor_name": consumor_name,
            "numberOfReviews": numberOfReviews,
            "location": location,
            "date": date,
            "rating": rating,
            "headline": headline,
            "review_text": review_text,
            "companyName": companyName,
            "companyResponded": companyResponded,
            }

            self.writeReviewToXL(count, rev)
            self.insertReviewToDB(rev)

            count += 1

            if count == 6:
                break


    def insertCompanySummaryToDB(self, companySummary):
        self.companySummarysCol.insert_one(companySummary)


    def insertReviewToDB(self, review):
        self.reviewCol.insert_one(review)


def main():
    s = Service("C:\Program Files (x86)\chromedriver.exe")
    driver = webdriver.Chrome(service=s)
    driver.get("https://uk.trustpilot.com/")

    company = sys.argv
    cookieNotification = driver.find_element(By.ID, "onetrust-accept-btn-handler")

    if cookieNotification:
        cookieNotification.click()

    # Use search bar to search for company
    search = driver.find_element(By.XPATH, "/html/body/div[1]/div/main/section/div[1]/div/div/div/div/div/div/form/input")
    search.send_keys(company)
    search.send_keys(Keys.RETURN)

    # Create Excel file
    book = xlsxwriter.Workbook('CompanyDetails.xlsx')
    companySheet = createCompanySummarySheet(book)
    reviewSheet = createReviewSheet(book)

    # Set up mongodb connection
    client = MongoClient("{MongoDB}", tls = True, tlsAllowInvalidCertificates = True)
    db = client.companyWebScrapper
    reviewCol = db.reviews
    companySummarysCol = db.companySummarys

    # Gets company details and reviews then writes them to mongodb and an excel file
    companyScraper = CompanyScraper(driver, reviewCol, companySummarysCol, reviewSheet, companySheet)
    companyScraper.getCompanySummary()
    companyScraper.getReviews()
    book.close()


def createCompanySummarySheet(book):
        # Write labels to company summary sheet   
        sheet = book.add_worksheet("CompanySummary")
        row = 0    
        column = 0
        companyLabels = [ "Company Name", "Rating", "Review Count"]
        for label in companyLabels:    
            sheet.write(row, column, label)   
            column += 1    

        return sheet


def createReviewSheet(book):
        # Write labels to reviews sheet  
        sheet = book.add_worksheet("Reviews")
        row = 0    
        column = 0

        reviewLabels = [ "Consumor Name", "Number Of Reviews", "Location", "Date", "Rating", "Headline", "Review Text", "Company Name", "Company Responded"]
        for label in reviewLabels:    
            sheet.write(row, column, label)   
            column += 1    

        return sheet


if __name__ == "__main__":
    main()


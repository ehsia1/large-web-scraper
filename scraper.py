import xlsxwriter
import urllib2
from bs4 import BeautifulSoup

def main():
    categories = ['Admissions Software', 'Appointment Reminder Software', 'Appointment Scheduling Software', 'Assessment Software', 'Attendance Tracking Software', 'Business Management Software', 'Business Process Management Software', 'Call Recording Software', 'Call Tracking Software', 'Change Management Software', 'Classroom Management Software',
        'Collaboration Software', 'Community Software', 'Complaint Management Software', 'Contact Management Software', 'Customer Relationship Management Software', 'Customer Communications Management Software', 'Customer Engagement Software', 'Customer Experience Software', 'Customer Satisfaction Software', 'Customer Service Software',
        'eLearning Authoring Tools Software', 'Employee Engagement Software', 'Forms Automation Software', 'Help Desk Software', 'Idea Management Software', 'Knowledge Management Software', 'Live Chat Software', 'Market Research Software', 'Online CRM Software', 'Polling Software', 'Public Relations Software', 'Push Notifications Software',
        'Qualitative Data Analysis Software', 'Reputation Management Software', 'School Administration Software', 'Service Desk Software', 'Student Information System Software', 'Team Communication Software', 'Ticketing Software', 'UX Software', 'Learning Management System Software', 'Survey Software', 'Meeting Software', 'Payroll Software', 'School Accounting Software',
        'Issue Tracking Software', 'Business Intelligence Software', 'Customer Loyalty Software', 'Email Management Software', 'Employee Monitoring Software', 'Internal Communications Software', 'Master Data Management Software', 'Mobile Learning Software']
    out = {}
    # Missing Issue Management Software
    # need category, product name, company, rating, description, product details, features list
    for category in categories:
        url = "https://www.capterra.com/"
        out[category] = {}
        category2 = category.lower()
        category_path = '-'.join(category2.split(' '))
        url += category_path
        url += '/'
        print url
        try:
            category_page = urllib2.urlopen(url)
        except:
            print url
            continue
        soup_category_page = BeautifulSoup(category_page, 'html.parser')
        html = soup_category_page.find(id="js-products")
        links = []
        for link in html.find_all('a'):
            if '#reviews' not in link.get('href') and 'external_click' not in link.get('href') and link.get('href') not in links:
                links.append(str(link.get('href')))
        for link in links:
            url = "https://www.capterra.com"
            url += link
            product_page = urllib2.urlopen(url)
            soup_product_page = BeautifulSoup(product_page, 'html.parser')
            product = soup_product_page.find('h1').get_text()
            out[category][product] = {}
            out[category][product]['company'] = soup_product_page.find('h2').get_text()
            try:
                overall = str(soup_product_page.find('span', class_='color-gray milli rating-decimal').get_text()).strip()[:1]
                ease = ''
                customer_service = ''
                for i in range(len(soup_product_page.find_all('span', class_='color-gray milli rating-decimal'))):
                    if i == 2:
                        ease = str(soup_product_page.find_all('span', class_='color-gray milli rating-decimal')[i])[53:56]
                        if '.' not in ease:
                            ease = ease[:1]
                    if i == 3:
                        customer_service = str(soup_product_page.find_all('span', class_='color-gray milli rating-decimal')[i])[53:56]
                        if '.' not in customer_service:
                            customer_service = customer_service[:1]
                    if i > 3:
                        break
            except:
                overall = 'n/a'
                ease = 'n/a'
                customer_service = 'n/a'
            out[category][product]['overall_rating'] = overall
            out[category][product]['ease_of_use'] = ease
            out[category][product]['customer_service'] = customer_service
            out[category][product]['description'] = (soup_product_page.find_all('p')[1]).get_text()
            if 'No reviews, be the first' in out[category][product]['description'] or 'Pros:' in out[category][product]['description']:
                out[category][product]['description'] = (soup_product_page.find_all('p')[0]).get_text()
            list = soup_product_page.find_all('div', class_='cell five-twelfths')
            for detail in list:
                text = str(detail.get_text())
                details = soup_product_page.find_all('div', class_='cell seven-twelfths')
                if 'Starting Price' in text:
                    out[category][product]['starting_price'] = details[list.index(detail)].get_text()
                if 'Pricing Details' in text:
                    out[category][product]['pricing_details'] = details[list.index(detail)].get_text()
                if 'Free' in text:
                    out[category][product]['free_trial'] = details[list.index(detail)].get_text()
                if 'Deployment' in text:
                    out[category][product]['deployment'] = details[list.index(detail)].get_text()
                if 'Training' in text:
                    out[category][product]['training'] = (soup_product_page.find_all('div', class_='cell one-half')[6]).get_text()
                if 'Support' in text:
                    out[category][product]['support'] = details[list.index(detail)-1].get_text()
            features = []
            for feature in soup_product_page.find_all('li', class_='ss-check'):
                if not 'feature-disabled' in str(feature):
                    features.append(str(feature.get_text()))
            if len(features) == 0:
                out[category][product]['features'] = 'n/a'
            out[category][product]['features'] = ','.join(features)
    book = xlsxwriter.Workbook('Products5.xlsx')
    worksheet = book.add_worksheet()
    row = 0
    for category, val in out.iteritems():
        for product, value in val.iteritems():
            if not 'starting_price' in value:
                value['starting_price'] = 'n/a'
            if not 'pricing_details' in value:
                value['pricing_details'] = 'n/a'
            if not 'free_trial' in value:
                value['free_trial'] = 'n/a'
            if not 'deployment' in value:
                value['deployment'] = 'n/a'
            if not 'training' in value:
                value['training'] = 'n/a'
            if not 'support' in value:
                value['support'] = 'n/a'
            worksheet.write(row, 0, category)
            worksheet.write(row, 1, product)
            worksheet.write(row, 2, value['company'])
            worksheet.write(row, 3, value['overall_rating'])
            worksheet.write(row, 4, value['ease_of_use'])
            worksheet.write(row, 5, value['customer_service'])
            worksheet.write(row, 6, value['description'])
            worksheet.write(row, 7, value['starting_price'])
            worksheet.write(row, 8, value['pricing_details'])
            worksheet.write(row, 9, value['free_trial'])
            worksheet.write(row, 10, value['deployment'])
            worksheet.write(row, 11, value['training'])
            worksheet.write(row, 12, value['support'])
            worksheet.write(row, 13, value['features'])
            row += 1
    book.close()

if __name__ == "__main__":
    main()

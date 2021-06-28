from bs4 import BeautifulSoup as bs
import requests
import xlsxwriter
''' Web scraper to get members from the Australian Voice Association Website'''

urls = []

url_list = ["https://australianvoiceassociation.com.au/find-a-voice-professional/",
            "https://australianvoiceassociation.com.au/find-a-voice-professional/?ps&pn=2&limit=50",
            "https://australianvoiceassociation.com.au/find-a-voice-professional/?ps&pn=2&limit=100"]

for link in url_list:
    r = requests.get(link)

    # Convert to Beautiful Soup object
    soup = bs(r.content, features="html.parser")
    all_headers = soup.find_all("tr")

    # Pass in attributes to find all
    name = soup.find_all("h3", attrs = {"pmpro_member_directory_display-name"})
    # Add url to list of url's to be used to get member details
    for details in name:
        deets = details.find('a', href = True)
        urls.append(deets['href'])

print(f'Found {len(urls)} records.')
specialists = []
# specifics for each link
for url in urls:
    spec = requests.get(url)
    soup_spec = bs(spec.content, features="html.parser")
    # Name
    full_name = soup_spec.find("h2", attrs={"class":"pmpro_member_directory_name"}).text
    full_name = full_name.strip()
    # Phone
    phone_block = soup_spec.find("div", attrs={'class':"pmpro_member_directory_work_phone"})
    if phone_block:
        phone = phone_block.find_all("td")[1].text
    else:
        phone = ""
    # web
    web_block = soup_spec.find("div", attrs={'class':"pmpro_member_directory_website"})
    if web_block:
        web = web_block.find_all("td")[1].text
    else:
        web = ""
    # Company
    company_block = soup_spec.find("div", attrs={"class":"pmpro_member_directory_company"})
    if company_block:
        company = company_block.find_all("td")[1].text
    else:
        company = ""
    # Profession
    prof_block = soup_spec.find("div", attrs={"class":"pmpro_member_directory_profession"})
    if prof_block:
        prof = prof_block.find_all("td")[1].text
    else:
        prof = ""
    # email
    email_block = soup_spec.find("div", attrs={"class":"pmpro_member_directory_work_email"})
    if email_block:
        email = email_block.find_all("td")[1].text
    else:
        email = ""
    # state
    state_block = soup_spec.find("div", attrs={"class":"pmpro_member_directory_state"})
    if state_block:
        state = state_block.find_all("td")[1].text
    else:
        state = ""
    # services
    services_block = soup_spec.find("div", attrs={"class":"pmpro_member_directory_services_provided"})
    if services_block:
        services = services_block.find_all("td")[1].text
    else:
        services = ""
    # services2
    services2_block = soup_spec.find("div", attrs={"class":"pmpro_member_directory_other_services_provided"})
    if services2_block:
        services2 = services2_block.find_all("td")[1].text
    else:
        services2 = ""
    specialists.append([{'url':url, 'name':full_name, 'company':company, 'web':web,'phone':phone, 'email':email, 'profession':prof, 'state':state, 'services': services, 'services other':services2}])

# Create an excel file
workbook = xlsxwriter.Workbook('specialists.xlsx')
worksheet = workbook.add_worksheet()
worksheet.write('A1', 'Name')
worksheet.write('B1', 'Phone')
worksheet.write('C1', 'Company')
worksheet.write('D1', 'Web')
worksheet.write('E1', 'Email')
worksheet.write('F1', 'State')
worksheet.write('G1', 'Profession')
worksheet.write('H1', 'Services')
worksheet.write('I1', 'Other Services')
worksheet.write('J1', 'Url')
counter = 2
for specialist in specialists:
    worksheet.write(f'A{counter}', specialist[0]['name'])
    worksheet.write(f'B{counter}', specialist[0]['phone'])
    worksheet.write(f'C{counter}', specialist[0]['company'])
    worksheet.write_url(f'D{counter}', specialist[0]['web'])
    worksheet.write(f'E{counter}', specialist[0]['email'])
    worksheet.write(f'F{counter}', specialist[0]['state'])
    worksheet.write(f'G{counter}', specialist[0]['profession'])
    worksheet.write(f'H{counter}', specialist[0]['services'])
    worksheet.write(f'I{counter}', specialist[0]['services other'])
    worksheet.write_url(f'J{counter}', specialist[0]['url'])
    counter += 1
workbook.close()

print('Finished')
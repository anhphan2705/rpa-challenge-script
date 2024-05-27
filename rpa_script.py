import rpa as r
import pandas as pd

data_frame = pd.read_excel(
    "./rpa-challenge-script/data/challenge.xlsx"
)

# Start the tagUI process
r.init(turbo_mode=True, chrome_browser=True)

# Open the website
r.url("https://rpachallenge.com/")
r.wait()

# Click on Start button
r.click('//button[text()="Start"]')

# Enter the data
for index, row in data_frame.iterrows():
    r.type('//input[@ng-reflect-name="labelFirstName"]', row["First Name"])
    r.type('//input[@ng-reflect-name="labelLastName"]', row["Last Name "])
    r.type('//input[@ng-reflect-name="labelCompanyName"]', row["Company Name"])
    r.type('//input[@ng-reflect-name="labelRole"]', row["Role in Company"])
    r.type('//input[@ng-reflect-name="labelAddress"]', row["Address"])
    r.type('//input[@ng-reflect-name="labelEmail"]', row["Email"])
    r.type('//input[@ng-reflect-name="labelPhone"]', str(row["Phone Number"]))
    r.click('//input[@value="Submit"]')

# Screenshot of Webpage
r.snap("/html/body/app-root/div[2]", "result.png")

# Stop the tagUI process
r.close()

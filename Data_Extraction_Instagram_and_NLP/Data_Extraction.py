import requests
import nltk
import urllib.request
from bs4 import BeautifulSoup

# target url
url = 'https://insights.blackcoffer.com/rise-of-telemedicine-and-its-impact-on-livelihood-by-2040-3-2/'
reqs = requests.get(url)

# using the BeautifulSoup module
soup = BeautifulSoup(reqs.text, 'html.parser')
# displaying the title
print("Title of the website is : ")
for title in soup.find_all('title'):
    print(title.get_text())

for data in soup.find_all("p"):
    print(data.get_text())
#print(soup.head)
'''
with open("write.txt", "a+") as f:
    for data in soup.find_all("p"):
        sum = data.get_text()
        f.writelines(sum)

'''
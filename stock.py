import requests
def getCSV( code , year):
    try:
        response = requests.post('http://kabuoji3.com/sitemap/sitemap-index.xml',
        data ={'code':str(code),'year':str(year)},
        header = {'referer':'http://kabuoji3.com/sitemap/sitemap-index.xml'},
        "User-Agent":"Mo"

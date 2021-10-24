from urllib import urlencode
import urllib2

#url = 'http://201.68.144.88:7500/receiver' # Set destination URL here
url = "http://127.0.0.1:7500/receiver"
post_fields = {"data": "teste"}
#post_fields = "{\"data\": \"teste\"}"

post = urlencode(post_fields)
req = urllib2.Request(url, post)
response = urllib2.urlopen(req)
json = response.read()

print(json)
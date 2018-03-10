import quandl
quandl.ApiConfig.api_key = 'y2YcRpp7Xw3CrsyhQ16E'
data = quandl.get("THAISE/INDEX")
print(data.head())
from googlefinance import getQuotes
import json
symbol = 'PTT'

print(json.dumps(getQuotes('SET:' + symbol), indent=2))
from xlwings import Book, Sheet, Range, Chart
import time
import matplotlib.pyplot as plt


wb = Book()

def getQuote(ticker, type):
    Range('B1').value = '=RTD("tos.rtd", ,"%s", "%s")' % (type, ticker)
    time.sleep(2.5)
    return Range('B1').value

def quoteOption(type, ticker, cp, exp, strike):
    Range('B1').value = '=RTD("tos.rtd", ,"%s", ".%s%s%s%s")' % (type, ticker, exp, cp, strike)
    time.sleep(2.5) #wait for excel to update the cell, usually a couple of seconds.    return Range('b1')
    return(Range('B1').value)

#getQuote("AMD", "LAST")

count = 0

# while (count < 10):
#     quoteOption('LAST','AMD','C',170804, 14)
#     quoteOption('LAST','AMD','C',170804, 15)
#     quoteOption('LAST','AMD','C',170804, 16)
#     quoteOption('LAST','AMD','C',170804, 17)
#     count = count+1

# test = quoteOption('LAST','AMD','C',170804, 13)
# print(test)

price = []
count = 1
for strike_price in range(13,17):
    price.append(quoteOption('LAST','AMD','C',170804, strike_price))
    # count = count + 1

print(price)

plt.scatter(range(13,17), price, color='red', marker='.', alpha=0.5, s=400)
plt.show()






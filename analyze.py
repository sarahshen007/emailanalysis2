# Module to generate email summary
import openpyxl
import re
import nltk
from nltk.tokenize import word_tokenize
from nltk.corpus import stopwords
from nltk.corpus import words

correct_words = words.words()
stop_words = set(stopwords.words('english'))
relevant_words = set()
relevant_words.update(("account", 
                        "order", 
                        "email", 
                        "store", 
                        "password", 
                        "website", 
                        "online", 
                        "reset", 
                        "parts", 
                        "site", 
                        "number", 
                        "address", 
                        "purchase", 
                        "card", 
                        "time", 
                        "log", 
                        "vehicle", 
                        "code", 
                        "sign",
                        "signed", 
                        "in", 
                        "app", 
                        "rewards", 
                        "paypal", 
                        "phone", 
                        "cart", 
                        "find", 
                        "rebate", 
                        "change", 
                        "stock",
                        "add", 
                        "item", 
                        "check",
                        "auto", 
                        "link", 
                        "items", 
                        "search", 
                        "credit", 
                        "ordered", 
                        "received", 
                        "info", 
                        "login", 
                        "access", 
                        "error", 
                        "place", 
                        "apply", 
                        "car", 
                        "coupon", 
                        "send", 
                        "message", 
                        "fit", 
                        "buy", 
                        "battery", 
                        "oil", 
                        "headlight", 
                        "day", 
                        "next", 
                        "money", 
                        "charged", 
                        "receive", 
                        "today", 
                        "complete", 
                        "purchase", 
                        "purchased", 
                        "shipping", 
                        "pay", 
                        "discount", 
                        "emails", 
                        "payment", 
                        "stores",
                        "delivery", 
                        "product", 
                        "checkout", 
                        "service", 
                        "part", 
                        "deals", 
                        "discover", 
                        "location", 
                        "store", 
                        "signed",
                        "signing"
                        "sign",
                        "submit",
                        "rebate",
                        "rebates",
                        "date",
                        "dates"
                        "products",
                        "searching",
                        "searched",
                        "searches",
                        "shelf",
                        "applied",
                        "apply",
                        "history",
                        "receipt",
                        "receipts",
                        "finds",
                        "credits",
                        "hub",
                        "info",
                        "information",
                        "vehicle",
                        "truck",
                        "motorcycle",
                        "motorcycles",
                        "gas",
                        "purchases",
                        "purchasing",
                        "buy",
                        "bought",
                        "Hyundai",
                        "Honda",
                        "stock",
                        "%"
                        ))

productCorrespondence = {}


def generateData(path):
    #db = sqlite3.connect('azemail.db')
    global relevant_words
    global productCorrespondence

    wb = openpyxl.load_workbook(path) 
    sheet = wb['CS Feedback']
    data = {}

    issuesCol = []
    prevIssue = 'General Inquiry'

    issuesCol = [str(issue.value).lower() for issue in sheet['B']]
    
    for i in range(len(issuesCol)-1, 0, -1):
        issueValue = issuesCol[i]
        if issueValue == 'none':
            issuesCol[i] = prevIssue
        else:
            prevIssue = issueValue
        
        if issueValue not in productCorrespondence:
            productCorrespondence[issueValue] = str(sheet['C'][i].value).title()


    commentsCol = [removeStopWords(str(comment.value).lower()) for comment in sheet['F']]
    commentsCol = commentsCol[1:]
    issuesCol = issuesCol[1:]

    allWords = ""
    for entry in commentsCol:
        allWords += (" ").join(entry) + " "
    allWordsExceptStopDist = nltk.FreqDist(w.lower() for w in allWords if w not in stop_words)

    relevant_words.update(allWordsExceptStopDist.most_common(400))

    for issue in issuesCol:
        relevant_words.update([x.lower() for x in issue.split(' ')])

    for i in range(len(issuesCol)):
        issuesCell = issuesCol[i]
        filteredComment = commentsCol[i]

        if not issuesCell in data:
            data[issuesCell] = {}

        numWords = 0

        for word in filteredComment:
            numWords += 1
            if word in data[issuesCell]:
                data[issuesCell][word] += 1
            elif word in relevant_words:
                data[issuesCell][word] = 1

        for word in data[issuesCell]:
            if not numWords == 0:
                data[issuesCell][word] /= numWords

    wb.close()    
    return data

def wordFrequency(text):
    frequency = {}
    for word in text:
        if word.lower() in frequency:
            frequency[word.lower()] += 1
        elif word.lower() in relevant_words:
            frequency[word.lower()] = 1
    return frequency

def generateIssueSummary(text, prevData):
    text = removeStopWords(text)
    textFrequency = wordFrequency(text)

    wordsInText = set()
    for keyword in textFrequency:
        wordsInText.add(keyword)

    issuesList = []
    
    for issue in prevData:
        issuesList.append(issue)

    comparisonWeights = []
    
    for i in range(len(issuesList)):
        issue = issuesList[i]
        comparisonWeight = 0

        for keyword in prevData[issue]:
            
            if keyword in issuesList and keyword in wordsInText:
                comparisonWeight += prevData[issue][keyword] * textFrequency[keyword] * 1.1
            elif keyword in wordsInText:
                comparisonWeight += prevData[issue][keyword] * textFrequency[keyword]
        
        comparisonWeights.append(comparisonWeight)

    maxIndex = 0
    maxValue = 0
    for i in range(len(comparisonWeights)):
        if comparisonWeights[i] > maxValue:
            maxValue = comparisonWeights[i]
            maxIndex = i

    if maxValue == 0:
        return "General Inquiry"

    return [issuesList[maxIndex].title(), productCorrespondence[issuesList[maxIndex].lower()]]
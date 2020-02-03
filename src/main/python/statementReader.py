import PyPDF2


class StatementReader:

    def __init__(self, path):

        # Read PDF
        pdfFileObj = open(path, 'rb')
        self.pdfReader = PyPDF2.PdfFileReader(pdfFileObj)
        self.pages = self.pdfReader.numPages
        self.pageTexts = ""

        self.p = 0
        self.rows = 0
        self.trimDate = 0
        self.transaction_types = ("CR", "DD", "VIS", ")))", "DR", "BP")
        self.balance_thousands = 1

    def get_page_text(self):  # Concatenate all text from all pages
        p = 0

        while p < self.pages - 1:
            pageObj = self.pdfReader.getPage(p)
            self.pageTexts = self.pageTexts + pageObj.extractText()

            p += 1

        return self.pageTexts

    def retrieveTable(self, pageText):

            results = list()
            transactions = list()
            date_strings = []

            #array of clean date indexes
            date_indexes = self.getDate(pageText)
            #print(date_indexes)

            for d in date_indexes:
                date_strings.append(self.getDateString(d,pageText))

            for e in range(0, len(date_indexes)-2):

                    if e >= len(date_indexes)-1:
                            break

                    if int(date_strings[e][0:2]) > int(date_strings[e+1][0:2]) and date_strings[e][3:6] == date_strings[e+1][3:6] :
                            #print(date_strings[e][0:2])
                            #print(date_strings[e+1][0:2])
                            date_indexes.pop(e+1)

            # print(date_indexes)


            i = 0
            for date_index in date_indexes:


                if i > len(date_indexes)-2:  # if last date get the text until the end

                    print("Date index")
                    print(date_index)
                    print("Last page")
                    print(pageText[date_index:-1])

                    transactions = self.getTransactions(pageText[date_index:-1])
                    date = self.getDateString(date_index, pageText)
                    results.append([date, transactions])

                    break

                else:
                    # feed only text between the two dates

                    transactions = self.getTransactions(pageText[date_index:date_indexes[i+1]])
                    date = self.getDateString(date_index, pageText)
                    results.append([date, transactions])
                    i += 1

            return results



    ################ Date Methods ################

    def getDate(self, pageText):

        # Finds dates in format dd mmm yy and returns starting index number in the text
        # print(pageText)
        months = ("Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec")
        results = []
        for dd in range(1, 32):
            for mmm in months:
                for yy in range(10,25):
                    res = pageText.find(str(dd) + " " + mmm + " " + str(yy))
                    if res != -1:  #if found
                        results.append(res)
        results.sort()
        print("Results 1")
        print(results)

        length = len(results) - 2

        for a in range(0, length):  # This loop removes near duplicates caused by overlapping the same date eg 9 Dec 19 caused by 19 Dec 19

                if a >= len(results)-1:
                              break
                if results[a] == (results[a+1]-1):
                    #print(results[a+1])
                    results.remove(results[a+1])
                    #print("NewLength = " + str(len(results)))
        
        # print("Results 2")
        # print(results)
        # print("pageText")
        # print(pageText[4403:4503])
        return results



    # find date, save date, get transactions, save transactions, next date  [date1, [[name1, value1], [name2, value2]], date2...]
    # done to bind every transaction to a date

    # Transaction Methods ################

    # returns list of transactions for every day [[name, value], [name, value]]
    def getTransactions(self, pageText):
        transaction = list()
        transactions = list()

        transType_indexes = self.getTransType(pageText) #Get list of indexes of transaction types (ONLY for the current date)

        #ignore balance value
        balance_index = pageText.rfind(",")
        if balance_index != -1:
            pageText = pageText[:balance_index-self.balance_thousands]

        i = 0
        for trans in transType_indexes:

            if i > len(transType_indexes) - 2:  # condition if we have the last transaction
                transaction = self.getInfo(pageText[trans[0]:], trans[1])
                transactions.append(transaction)
                break

            # feed only text between the two transactions and trans type string

            transaction = self.getInfo(pageText[trans[0]:transType_indexes[i+1][0]], trans[1])
            transactions.append(transaction)

            i += 1

        return transactions


    def getTransType(self, pageText):
        transaction_types = ("CR", "DD", "VIS", ")))", "DR", "BP")
        result = list()

        for c in range(0,len(pageText)):
        # c = string index number
            for p in transaction_types:
                if pageText[c-1] in "0123456789" or pageText[c-1] == "s":
                        # if transaction type is preceded by a number

                    length = len(p)
                    if length == 3:
                        if pageText[c:c+3] == p:
                            # print(pageText[c:c+3])
                            # print("found")
                            result.append([c, p])
                    elif length == 2:
                        if pageText[c:c+2] == p:
                            #print(pageText[c:c+2])
                            #print("found")
                            result.append([c, p])

        return result

    def getInfo(self, pageText, trans_type):

        name = "N/A"
        value = "0.0"
        info = list()

        if trans_type == "VIS":

                if pageText[3:8] == "INT'L":

                    checkPaypal = pageText.find("PAYPAL")

                    if checkPaypal != -1:
                            #INTL PAYPAL CASE

                            #find name
                            newLine = pageText.find("\n")
                            temp = pageText[newLine+9:]
                            newLine = temp.find("\n")
                            name = temp[:newLine-1]

                            #find value
                            if temp.find("001") != -1:
                                    value = temp[newLine+12:newLine+17]
                            else:
                                    value = temp[newLine+11:newLine+16]

                            #remove 10/11 digit number + n character assuming balance value has been removed
                    elif pageText.find("amzn.co.uk/pm") != -1:

                            #INT'L Amazon UK Prime case
                            index = pageText.find("amzn.co.uk/pm")
                            name = pageText[index-21:index]
                            value = pageText[index+13:]

                    elif pageText.find("AMAZON.CO.UK") != -1:

                            index = pageText.rfind("AMAZON.CO.UK")
                            name = pageText[index-21:index]
                            value = pageText[index+12:]

                    else:
                            #INTL Visa Rate CASE

                            #find name
                            newLine = pageText.find("\n")
                            temp = pageText[newLine+1:]
                            newLine = temp.find("\n")
                            name = temp[:newLine]

                            #find value
                            index = pageText.find("Visa Rate")
                            value = pageText[index+9:]

                elif pageText[3:9] == "PAYPAL":
                        #find name
                        pageText = pageText[3:]
                        index = self.find11(pageText)
                        name = pageText[:index]

                        #find value
                        value = pageText[index+11:]

                else:

                        index = self.findVal(pageText) #Find where the value starts
                        if pageText.find("E14") != -1:
                            value = pageText[index+2:]

                        elif pageText.find("E1 4") != -1:
                            value = pageText[index+3:]

                        else:
                            value = pageText[index:]

                        name = pageText[3:index]

        elif trans_type in (")))", "DD", "BP"):

                index = self.findVal(pageText) #Find where the value starts

                #SPECIAL CASES
                if trans_type == ")))" and pageText.find("WH SMITH") != -1:
                        index = pageText.find("WH SMITH")
                        name = pageText[3:index+8]
                        value = pageText[index+18:]

                elif trans_type == ")))" and pageText.find("NIGHTLIGHT LEISURE") != -1:
                        index = pageText.rfind("1")
                        name = pageText[3:index]
                        value = pageText[index:]

                else:
                        if len(trans_type) == 3:
                                name = pageText[3:index]

                        elif len(trans_type) == 2:
                                name = pageText[2:index]

                        else:
                               name = pageText[3:index]

                        if pageText.find("E14") != -1:
                            value = pageText[index+2:]

                        elif pageText.find("E1 4") != -1:
                            value = pageText[index+3:]

                        else:
                            value = pageText[index:]

        elif trans_type == "CR":

                if "QUEEN MARY" in pageText:
                        name = pageText[2:20]
                        value = pageText[20:]

                elif "LONDON SCHOOL" in pageText:
                        name = pageText[2:23]
                        value = pageText[23:]

                else:
                    name = pageText[2:15]
                    value = pageText[len(pageText)-4:]


        elif trans_type == "DR":

                index = pageText.find("Transaction Fee")
                if index == -1:
                        raise Exception("DR is not a transaction fee")
                name = pageText[2:index+15]
                value = pageText[index+15:]

        else:
            raise Exception("Unknown transaction type")

        # Special case for changing from first page

        next_page_name = name.find("BALANCEBROUGHTFORWARD")
        if next_page_name != -1:
                name = name[:next_page_name]
                prevChar = ""
                dot_index = name.find(".")
                i = 0
                while i < len(name):
                    if dot_index == -1:
                        break

                    c = name[i]
                    prevChar = name[i-1]
                    if c in "0123456789" and prevChar in "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz":
                            value = name[i:dot_index+3]
                            name = name[:i]
                    i += 1

        # Special case for changing from any page

        next_page_value = value.find("BALANCEBROUGHTFORWARD")
        if next_page_value != -1:
                value = value[:next_page_value]

        info.append(name)
        info.append(value)

        return info

    ###################################### Utility methods ##################################################

    def findVal(self, str_in):

        # find where character changes with no space to a number then find(".") and return index of start
        prevChar = ""
        dot_index = str_in.rfind(".")
        i = dot_index
        while i > 1:
            c = str_in[i]
            prevChar = str_in[i-1]
            if c in "0123456789" and prevChar in "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz":
                return i
            i -= 1
        return -1

    def find11(self, str_in):
        i = 0
        flag = 0
        for i in range(0,len(str_in)-10): #go through all given text
            string = str_in[i:i+11]

            for n in string: #for every 11 character string
                if n in ("0", "1", "2", "3", "4", "5", "6", "7", "8", "9"):
                    flag += 1

            if flag == 11: #found 11 numbers in a row
                return i
            else:
                flag = 0

        return -1

    def getDateString(self, date_index, pageText):

        #transforms string index number in date
        date = ""

        if pageText[date_index:date_index+3].find(" ") == 1: #d
            date = pageText[date_index:date_index+8]
        else: #dd
            date = pageText[date_index:date_index+9]

        return date

    def newLineClean(self, data):
        # Removes new line characters from name

        i = 0
        for d in data:
            j=0
            for t in d[1]:
                data[i][1][j][0] = t[0].replace("\n", "")
                j += 1
            i += 1

        return data

####################################### Main program ###################################################


if __name__ == '__main__':

    # Create one long text file and then analyze

    reader = StatementReader('statements.pdf')
    data = reader.retrieveTable(reader.get_page_text())
    data = reader.newLineClean(data)

    print(data)


    











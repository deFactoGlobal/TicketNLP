import xlrd
from  matplotlib import pyplot as plt
import gensim
from gensim.models import Word2Vec
import numpy as np

class Ticket:
    def __init__(self, id, category, title, ref, submitter, assignedto, prio, statusid, active, start):
        self.id = id
        self.category = category
        self.title = title
        self.ref = ref
        self.prio = prio
        self.submitter = submitter
        self.assignedto = assignedto
        self.statusid = statusid
        self.active = active
        self.start = start
        self.end = start
        self.details = []
        self.comments = []

    def check_ref(self, given):
        if (given == ref): return True
        return False

    def add_coument(self, c):
        self.comments.append(c)

    def add_details(self, d):
        self.details.append(d)

    def time_taken(self):
        return 0

def process(query):
    query_embed = embeddings_index[query]
    scores = {}
    for word, embed in data_embeddings.items():
        category = categories[word]
        dist = query_embed.dot(embed)
        dist /= len(data[category])
        scores[category] = scores.get(category, 0) + dist
    return scores

if __name__ == "__main__":
    ticketdict = {}
    assigned_count = {}
    datadict = {}

    loc = ("AllSupportComments.xlsx")

    wb = xlrd.open_workbook(loc)
    sheet = wb.sheet_by_index(0)
    sheet.cell_value(0, 0)

    # Stripping data from excel file
    for i in range(1,sheet.nrows):
        instance = sheet.row(i)
        ref = int(instance[20].value)
        if ref in ticketdict:
            ticketdict[ref].details.append(instance[3].value)
            ticketdict[ref].comments.append(instance[19].value)
            ticketdict[ref].end = instance[22].value
        else:
            t = Ticket(instance[0].value, instance[1].value, instance[2].value, ref, instance[21].value, instance[23].value, instance[24].value, instance[25].value, instance[26].value, instance[22].value)
            t.details.append(instance[3].value)
            t.comments.append(instance[19].value)
            ticketdict[ref] = t

            categ = instance[1].value
            titl = instance[2].value.split()
            tit = [i for i in titl if len(i) >= 3]
            if categ in datadict:
                datadict[categ].extend(list( dict.fromkeys(tit) ))
            else:
                datadict[categ] = []
                datadict[categ].extend(list( dict.fromkeys(tit) ))
                if categ not in datadict[categ]:
                    datadict[categ].extend(categ.split())
                datadict[categ] = list( dict.fromkeys(datadict[categ]) )

            worker = int(instance[23].value)
            if worker in assigned_count:
                assigned_count[worker]+=1
            else:
                assigned_count[worker] = 1


    #print(len(datadict))

    categories = {word: key for key, words in datadict.items() for word in words}

# Load the whole embedding matrix
    embeddings_index = {}
    with open('glove.6B.100d.txt') as f:
        for line in f:
            values = line.split()
            word = values[0]
            embed = np.array(values[1:], dtype=np.float32)
            embeddings_index[word] = embed
    #print('Loaded %s word vectors.' % len(embeddings_index))
# Embeddings for available words
    data_embeddings = {key: value for key, value in embeddings_index.items() if key in categories.keys()}

# Processing the query
    def process(query):
        query_embed = embeddings_index[query]
        scores = {}
        for word, embed in data_embeddings.items():
            category = categories[word]
            dist = query_embed.dot(embed)
            dist /= len(datadict[category])
            scores[category] = scores.get(category, 0) + dist
        return scores

# Testing
    #print(process('macro'))

    inputproc = input("Give string of title (letters and spaces only): ")
    inputarr = inputproc.split()
    resultarr = []
    for i in inputarr:
        thisdict = process(i)
        minval = 500000
        minkey = ""
        for word in thisdict:
            if (thisdict[word] < minval):
                minval = thisdict[word]
                minkey = word
        resultarr.append(minkey)
    print(max(set(resultarr), key = resultarr.count) )

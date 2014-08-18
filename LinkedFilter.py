__author__ = 'Ashar Malik'

import csv
import re
from os import listdir

def read_csv(file_name):
    csv_info = []
    with open(file_name, 'rb') as f:
        reader = csv.reader(f)
        for row in reader:
            csv_info.append(row)

    return csv_info

def getSections():
    ret = []
    files = listdir("groups")

    for file in files:
        with open('groups/'+file) as f:
            words = f.read().split("\n")
            group_name = file.split(".")[0]
            ret.append([group_name, words])

    return ret

def write_contacts(contacts, sections, file_name):
    to_write = '"First Name","Last Name","E-mail Address","Company","Job Title"\n'

    for i, group in enumerate(contacts):
        if i<sections.__len__():
            to_write = to_write+sections[i][0]+"\n"
        else:
            to_write = to_write+"Other\n"
        #print group #list of tuples (first name, last name, etc)..
        for person in group:
            to_write = to_write + ','.join('"' + item + '"' for item in person)+"\n" #add quotes around items

    with open(file_name, 'w') as csv_file:
        csv_file.write(to_write)

def lev(s1, s2):
    if len(s1) < len(s2):
        return lev(s2, s1)

    # len(s1) >= len(s2)
    if len(s2) == 0:
        return len(s1)

    previous_row = range(len(s2) + 1)
    for i, c1 in enumerate(s1):
        current_row = [i + 1]
        for j, c2 in enumerate(s2):
            insertions = previous_row[j + 1] + 1 # j+1 instead of j since previous_row and current_row are one character longer
            deletions = current_row[j] + 1       # than s2
            substitutions = previous_row[j] + (c1 != c2)
            current_row.append(min(insertions, deletions, substitutions))
        previous_row = current_row

    return previous_row[-1]

def fuzzy_match(word1, word2):
    word1 = str(word1).lower()
    word2 = str(word2).lower()

    #Remove Incorporated
    word1 = word1.replace(" incorporated", "")
    word2 = word2.replace(" incorporated", "")

    #remove Inc
    word1 = word1.replace(" inc", "")
    word2 = word2.replace(" inc", "")

    p = re.compile("[ ,.-]") #remove , or . or -
    word1 = p.sub("", word1)
    word2 = p.sub("", word2)

    p = re.compile("\(.*\)") #remove (xxxx)
    word1 = p.sub("", word1)
    word2 = p.sub("", word2)

    if word1 == word2:
        return True

    #print "Word1: %s Word2: %s" % (word1, word2)

    if word1.__len__() == 0:
        return False
    if float(lev(word1, word2))/word1.__len__() < .25:
        return True

    if word1.__contains__(word2) or word2.__contains__(word1):
        return True
    return False

def validateCompany(company, cList):
    if company in cList:
        return True

    for comp in cList:
        if fuzzy_match(company, comp):
            #print "Fuzzy match on %s" % comp
            return True

    return False

def load_companies():
    with open("companies.txt", "r") as f:
        companies = f.read().split("\n")
    return companies


def process_contacts(file_name):
    data = read_csv(file_name)

    #truncate unnecessary columns
    #First Name, Last Name, E-mail Address, Company, Job Title

    companies = load_companies()
    keep_columns = []
    headers = data.pop(0)
    keep_columns.append(list(headers).index("First Name"))
    keep_columns.append(list(headers).index("Last Name"))
    keep_columns.append(list(headers).index("E-mail Address"))
    keep_columns.append(list(headers).index("Company"))
    keep_columns.append(list(headers).index("Job Title"))
    data.pop(data.__len__()-1) #remove last item; it is empty

    arguments = []
    for index in keep_columns:
        arguments.append(zip(*data)[index]) #arguments = [first_names, last_names, ...]

    data = zip(*arguments) #data is now truncated

    sections = getSections() #"section name", keywords

    contacts = []

    for i in range(0, sections.__len__()+1): #last section is Other
        contacts.append([])

    for person in data:
        #('First Name', 'Last Name', 'E-mail Address', 'Company', 'Job Title')
        job_title = person[4]
        company = person[3]

        if validateCompany(company, companies) == False:
            continue

        print "Accepting %s " % company

        for i, sec in enumerate(sections): #loop through each Group
            word_list = sec[1]
            organized = False
            for word in word_list: #loop through each word linked to a group
                if str(job_title).__contains__(word):
                    contacts[i].append(person)
                    #print "Adding to section %s" % sections[i][0]
                    organized = True
                    break
            if organized:
                break #no need to search through rest of sections if already categorized

        if not organized:
            contacts[contacts.__len__()-1].append(person) #add to last section, 'Other'

    write_contacts(contacts, sections, "processed.csv")

pass

process_contacts('linkedin_connections_export_microsoft_outlook.csv')
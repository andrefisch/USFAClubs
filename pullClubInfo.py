import openpyxl


def template(fin):
    # CREATE A NEW EXCEL FILE
    out = openpyxl.Workbook()
    # Open the worksheet we want to edit
    outsheet = out.create_sheet("Clubs")

    # if 'sheet' appears randomly we can delete it
    rm = out['Sheet']
    out.remove(rm)

    fout = 'Clubs.xlsx'

    #################
    # DO STUFF HERE #
    #################
    fin = open(fin, encoding="utf8")
    lines = fin.readlines()

    outsheet['A1'].value = "Club Name"
    outsheet['B1'].value = "Club Abbreviation"
    outsheet['C1'].value = "Club Address"
    outsheet['D1'].value = "Club Phone Number"
    outsheet['E1'].value = "Club Emai Address"
    outsheet['F1'].value = "Club Website"
    outsheet['G1'].value = "Club Region"
    outsheet['H1'].value = "Club Division"

    row = 2

    for i in range(0, len(lines)):
        if ('lead mb-0' in lines[i]):
            one = lines[i].find('lead mb-0') + 11
            two = lines[i].find('(')
            three = lines[i].find(')')
            # Club Name
            outsheet['A' + str(row)].value = lines[i][one:two - 1]
            print("Now processing", lines[i][one:two - 1])
            # Club Abbreviation
            outsheet['B' + str(row)].value = lines[i][two + 1:three]
            # Club Address
            outsheet['C' + str(row)].value = lines[i + 7].strip().replace('<br/>', '\\n')
            # Club Phone Number
            one = lines[i + 10].find('>') + 1
            two = lines[i + 10].find('/') - 1
            outsheet['D' + str(row)].value = lines[i + 10][one:two].strip()
            # Club Email Address
            one = lines[i + 12].find('>') + 1
            two = lines[i + 12].find('/') - 1
            outsheet['E' + str(row)].value = lines[i + 12][one:two].strip()
            # Club Website
            one = lines[i + 14].find('a href') + 8
            two = lines[i + 14][one:].find('"') + one
            outsheet['F' + str(row)].value = lines[i + 14][one:two].strip()
            # Club Region
            outsheet['G' + str(row)].value = lines[i + 17].strip().replace('</td>', '')
            # Club Division
            outsheet['H' + str(row)].value = lines[i + 20].strip().replace('</td>', '')
            row = row + 1

    print("Saving " + fout)
    out.save(fout)

template('clubs.txt')

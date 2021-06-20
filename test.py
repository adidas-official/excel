from csv import reader

with open('unor.csv','r') as upd:
    csv_reader = reader(upd)
    l = list(map(list,csv_reader))
    # for i in l:
    #     print(i)
    # print(len('Tomás'))
    name = 'Tomás'
    print(name)

    for n in name:
        print(ord(n),end=',')
    # print(len(l[3][0]))
    print()
    for c in l[3][0]:
        print(ord(c),end=',')
    print()
    print(name==l[3][0])



file = 'TEXT.TXT'

with open(file, 'r', encoding='windows-1250') as f:

    names = f.readlines()
    lines = [i for i, s in enumerate(names) if '-------' in s]

    s_lines = sorted(lines, reverse=True)

    for index in s_lines:
        if index == s_lines[-1]: # last element
            del names[:lines[0]+1]
        elif index == s_lines[0]: # first element
            del names[lines[-1]:]
        else:
            del names[index-1:index+1]

    data = []
    for name in names:

        n = " ".join(name.split())
        splited = n.split(' ')
        mzda, jmeno = splited[-1].replace('.',''), splited[0]
        # jmeno = splited[0]

        if '/' in mzda:
            mzda = ''
        else:
            mzda = int(mzda)

        l = (jmeno, mzda)
        data.append(l)

    for i in data:
        print(i)



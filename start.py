import app

for i in range(1, 4):
    name = ''
    if len(str(i)) == 1:
        name = 'HYDERABAD Survey Data 13-14 (Scans 2)_Page_' + \
            "100" + str(i) + '.jpg'
    elif len(str(i)) == 2:
        name = 'HYDERABAD Survey Data 13-14 (Scans 2)_Page_' + \
            "10" + str(i) + '.jpg'
    elif len(str(i)) == 3:
        name = 'HYDERABAD Survey Data 13-14 (Scans 2)_Page_' + \
            "1" + str(i) + '.jpg'
    elif len(str(i)) == 4:
        name = 'HYDERABAD Survey Data 13-14 (Scans 2)_Page_' + str(i) + '.jpg'

    app.imgToexcel(name)

    #############

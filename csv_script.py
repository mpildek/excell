#! /usr/bin/env python
# -*- coding: utf-8 -*-

import xlwt
import csv
import sys
datas = {}
podaci = []
header_style = xlwt.easyxf("pattern: fore_colour dark_blue, pattern solid;" "font: colour white;", "#,###.00")
special_style = xlwt.easyxf("pattern: fore_colour blue, pattern solid;" "font: colour white;", "#,###.00")
kategorije_style = xlwt.easyxf("pattern: fore_colour light-blue, pattern solid;" "font: colour white;", "#,###.00")
razrada_style = xlwt.easyxf("pattern: fore_colour blue_gray, pattern solid;" "font: colour white;", "#,###.00")
pod_style = xlwt.easyxf("borders: left thin, right thin,bottom thin; ", "#,###.00")
borders = xlwt.easyxf("borders: left thin, right thin,bottom thin; ", "#,###.00")

 # forat code  #,###.00

with open('/home/matija/Desktop/Tabela.csv', 'r') as myfile:
    data = csv.reader(myfile)
    #sortiranje po kupcu i godini unutar kupca
    for d in data:
        if d[1] == 'Kupac':
             continue
        #ako postoji kupac upiši ako ne kreiraj novi zapis
        if d[1] in datas:
            #ako postoji grupa unutar kupca dodaj ako ne kreiraj novu grupu
            if d[2] in datas[d[1]]:
                datas[d[1]][d[2]].append((d[3], d[4], d[6].strip().upper(), d[7], d[9], d[10], d[11]))
            else:
                datas[d[1]][d[2]] = [(d[3], d[4], d[6].strip().upper(), d[7], d[9], d[10], d[11])]
        else:
            datas[d[1]] = {d[2]: [(d[3], d[4], d[6].strip().upper(), d[7], d[9], d[10], d[11])]}

partneri = {}
grupe = []
podaci = {}
book = xlwt.Workbook(encoding="utf-8")
#partner je partner
for partner in datas:
    for grupa in datas[partner]:
        podaci = {}
        for kateg in datas[partner][grupa]:
            if kateg[0] in podaci:
                if kateg[1] in podaci[kateg[0]]:
                   podaci[kateg[0]][kateg[1]].append((kateg[2], kateg[3], kateg[4], kateg[5][-4:], kateg[6], partner))
                else:
                   podaci[kateg[0]][kateg[1]] = [(kateg[2], kateg[3], kateg[4], kateg[5][-4:], kateg[6], partner)]
            else:
                podaci[kateg[0]] = {kateg[1]: [(kateg[2], kateg[3], kateg[4], kateg[5][-4:], kateg[6], partner)]}
        grupe.append({grupa: podaci})
    print partner

    sheet = book.add_sheet(partner, cell_overwrite_ok=True)
    sheet.write(0, 0, "Partner", style=header_style)
    sheet.write(0, 1, partner, style=header_style)
    #Sirina stupca B,C,D
    sheet.col(1).width = 256 * 15
    sheet.col(2).width = 256 * 15
    sheet.col(3).width = 256 * 20
    for r in range(2, 16, 1):
        sheet.write(0, r, style=header_style)

    sheet.write(1, 4, "Godine", style=header_style)
    sheet.write(1, 5, "Podatak", style=header_style)

    for t in range(0, 16, 1):
        sheet.write(1, t, style=header_style)

    for w in range(0, 16, 1):
        sheet.write(2, w, style=header_style)

    # sheet.write(2, 4, "2009", style=header_style)
    # sheet.write(2, 6, "2010", style=header_style)
    # sheet.write(2, 8, "2011", style=header_style)
    # sheet.write(2, 10, "2012", style=header_style)
    # sheet.write(2, 12, "2013", style=header_style)
    # sheet.write(2, 14, "2014", style=header_style)
    # sheet.write(2, 16, "2015", style=header_style)
    # sheet.write(2, 18, "2016", style=header_style)
    # sheet.write(2, 20, "2017", style=header_style)

    sheet.write(2, 4, "2012", style=header_style)
    sheet.write(2, 6, "2013", style=header_style)
    sheet.write(2, 8, "2014", style=header_style)
    sheet.write(2, 10, "2015", style=header_style)
    sheet.write(2, 12, "2016", style=header_style)
    sheet.write(2, 14, "2017", style=header_style)
    # sheet.write(2, 16, "", style=header_style)
    # sheet.write(2, 18, "", style=header_style)
    # sheet.write(2, 20, "", style=header_style)

    sheet.write(3, 0, "Grupa", style=header_style)
    sheet.write(3, 1, "Kategorija", style=header_style)
    sheet.write(3, 2, "Razrada", style=header_style)
    sheet.write(3, 3, "Model", style=header_style)
    sheet.write(3, 4, "Iznos", style=header_style)
    sheet.write(3, 5, "Kom.", style=header_style)
    sheet.write(3, 6, "Iznos", style=header_style)
    sheet.write(3, 7, "Kom.", style=header_style)
    sheet.write(3, 8, "Iznos", style=header_style)
    sheet.write(3, 9, "Kom.", style=header_style)
    sheet.write(3, 10, "Iznos", style=header_style)
    sheet.write(3, 11, "Kom.", style=header_style)
    sheet.write(3, 12, "Iznos", style=header_style)
    sheet.write(3, 13, "Kom.", style=header_style)
    sheet.write(3, 14, "Iznos", style=header_style)
    sheet.write(3, 15, "Kom.", style=header_style)
    # sheet.write(3, 16, "Iznos", style=header_style)
    # sheet.write(3, 17, "Kom.", style=header_style)
    # sheet.write(3, 18, "Iznos", style=header_style)
    # sheet.write(3, 19, "Kom.", style=header_style)
    # sheet.write(3, 20, "Iznos", style=header_style)
    # sheet.write(3, 21, "Kom.", style=header_style)
    row = 4

    for group in grupe:
        sheet.write(row, 0, group.keys(), style=special_style)
        for c in range(1, 16, 1):
            sheet.write(row, c, style=special_style)
        suu9, suu10, suu11, suu12, suu13, suu14, suu15, suu16, suu17, suu18 = 0, 0, 0, 0, 0, 0, 0, 0, 0, 0
        kuu9, kuu10, kuu11, kuu12, kuu13, kuu14, kuu15, kuu16, kuu17, kuu18 = 0, 0, 0, 0, 0, 0, 0, 0, 0, 0
        sumarum = row
        row += 1
        for g in group:
            ukupnog=[]
            for x in group[g]:
                su9, su10, su11, su12, su13, su14, su15, su16, su17, su18 = 0, 0, 0, 0, 0, 0, 0, 0, 0, 0
                ku9, ku10, ku11, ku12, ku13, ku14, ku15, ku16, ku17, ku18 = 0, 0, 0, 0, 0, 0, 0, 0, 0, 0
                firstrowG = row
                sheet.write(row, 1, x, style=kategorije_style)
                for c in range(2, 16, 1):
                    sheet.write(row, c, style=kategorije_style)
                row += 1
                for d in group[g][x]:
                    #zapis sume po razradi
                    s9, s10, s11, s12, s13, s14, s15, s16, s17, s18 = 0, 0, 0, 0, 0, 0, 0, 0, 0, 0
                    k9, k10, k11, k12, k13, k14, k15, k16, k17, k18 = 0, 0, 0, 0, 0, 0, 0, 0, 0, 0
                    sheet.write(row, 2, d, style=razrada_style)
                    for c in range(3, 16, 1):
                        sheet.write(row, c, style=razrada_style)
                    firstrow = row
                    row += 1
                    produkt = {}
                    for y in group[g][x][d]:
                        if y[0] in produkt:
                            if y[3] in produkt[y[0]]:
                                iznos = float(produkt[y[0]][y[3]][0][0].replace(',', ''))+float(y[2].replace(',', ''))
                                kom = float(produkt[y[0]][y[3]][0][1]) + float(y[4])
                                tl = list(produkt[y[0]][y[3]][0])
                                tl[0], tl[1] = str(iznos), str(kom)
                                produkt[y[0]][y[3]][0] = tuple(tl)
                            else:
                                produkt[y[0]][y[3]] = [(y[2], y[4])]
                        else:
                            produkt[y[0]] = {y[3]: [(y[2], y[4])]}

                    tmp = ''

                    for prod in produkt:
                        for godina in produkt[prod]:
                            print row
                            if tmp == prod:
                                pass
                            else:
                                for c in range(3, 16, 1):
                                    sheet.write(row, c, style=borders)
                                sheet.write(row, 3, prod, style=borders)
                                row += 1
                            for let in produkt[prod][godina]:
                                #zbog zapisa podataka o godini jedan red prenisko
                                row = row - 1
                                if godina == '2012':
                                    sheet.write(row, 4, float(let[0]), style=pod_style)
                                    sheet.write(row, 5, float(let[1]), style=pod_style)
                                    s12 += float(let[0])
                                    k12 += float(let[1])
                                    sheet.write(firstrow, 4, s12, style=razrada_style)
                                    sheet.write(firstrow, 5, k12, style=razrada_style)
                                    su12 += float(let[0])
                                    ku12 += float(let[1])
                                    sheet.write(firstrowG, 4, su12, style=kategorije_style)
                                    sheet.write(firstrowG, 5, ku12, style=kategorije_style)
                                    suu12 += float(let[0])
                                    kuu12 += float(let[1])
                                    sheet.write(sumarum, 4, suu12, style=special_style)
                                    sheet.write(sumarum, 5, kuu12, style=special_style)
                                    # tmp = prod
                                elif godina == '2013':
                                    sheet.write(row, 6, float(let[0]), style=pod_style)
                                    sheet.write(row, 7, float(let[1]), style=pod_style)
                                    s13 += float(let[0])
                                    k13 += float(let[1])
                                    sheet.write(firstrow, 6, s13, style=razrada_style)
                                    sheet.write(firstrow, 7, k13, style=razrada_style)
                                    su13 += float(let[0])
                                    ku13 += float(let[1])
                                    sheet.write(firstrowG, 6, su13, style=kategorije_style)
                                    sheet.write(firstrowG, 7, ku13, style=kategorije_style)
                                    suu13 += float(let[0])
                                    kuu13 += float(let[1])
                                    sheet.write(sumarum, 6, suu13, style=special_style)
                                    sheet.write(sumarum, 7, kuu13, style=special_style)
                                    # tmp = prod
                                elif godina == '2014':
                                    sheet.write(row, 8, float(let[0]), style=pod_style)
                                    sheet.write(row, 9, float(let[1]), style=pod_style)
                                    s14 += float(let[0])
                                    k14 += float(let[1])
                                    sheet.write(firstrow, 8, s14, style=razrada_style)
                                    sheet.write(firstrow, 9, k14, style=razrada_style)
                                    su14 += float(let[0])
                                    ku14 += float(let[1])
                                    sheet.write(firstrowG, 8, su14, style=kategorije_style)
                                    sheet.write(firstrowG, 9, ku14, style=kategorije_style)
                                    suu14 += float(let[0])
                                    kuu14 += float(let[1])
                                    sheet.write(sumarum, 8, suu14, style=special_style)
                                    sheet.write(sumarum, 9, kuu14, style=special_style)
                                    # tmp = prod
                                elif godina == '2015':
                                    sheet.write(row, 10, float(let[0]), style=pod_style)
                                    sheet.write(row, 11, float(let[1]), style=pod_style)
                                    s15 += float(let[0])
                                    k15 += float(let[1])
                                    sheet.write(firstrow, 10, s15, style=razrada_style)
                                    sheet.write(firstrow, 11, k15, style=razrada_style)
                                    su15 += float(let[0])
                                    ku15 += float(let[1])
                                    sheet.write(firstrowG, 10, su15, style=kategorije_style)
                                    sheet.write(firstrowG, 11, ku15, style=kategorije_style)
                                    suu15 += float(let[0])
                                    kuu15 += float(let[1])
                                    sheet.write(sumarum, 10, suu15, style=special_style)
                                    sheet.write(sumarum, 11, kuu15, style=special_style)
                                    # tmp = prod
                                elif godina == '2016':
                                    sheet.write(row, 12, float(let[0]), style=pod_style)
                                    sheet.write(row, 13, float(let[1]), style=pod_style)
                                    s16 += float(let[0])
                                    k16 += float(let[1])
                                    sheet.write(firstrow, 12, s16, style=razrada_style)
                                    sheet.write(firstrow, 13, k16, style=razrada_style)
                                    su16 += float(let[0])
                                    ku16 += float(let[1])
                                    sheet.write(firstrowG, 12, su16, style=kategorije_style)
                                    sheet.write(firstrowG, 13, ku16, style=kategorije_style)
                                    suu16 += float(let[0])
                                    kuu16 += float(let[1])
                                    sheet.write(sumarum, 12, suu16, style=special_style)
                                    sheet.write(sumarum, 13, kuu16, style=special_style)
                                    # tmp = prod
                                elif godina == '2017':
                                    sheet.write(row, 14, float(let[0]), style=pod_style)
                                    sheet.write(row, 15, float(let[1]), style=pod_style)
                                    s17 += float(let[0])
                                    k17 += float(let[1])
                                    sheet.write(firstrow, 14, s17, style=razrada_style)
                                    sheet.write(firstrow, 15, k17, style=razrada_style)
                                    su17 += float(let[0])
                                    ku17 += float(let[1])
                                    sheet.write(firstrowG, 14, su17, style=kategorije_style)
                                    sheet.write(firstrowG, 15, ku17, style=kategorije_style)
                                    suu17 += float(let[0])
                                    kuu17 += float(let[1])
                                    sheet.write(sumarum, 14, suu17, style=special_style)
                                    sheet.write(sumarum, 15, kuu17, style=special_style)
                                    # tmp = prod
                                elif godina == '2018':
                                    sheet.write(row, 16, float(let[0]), style=pod_style)
                                    sheet.write(row, 17, float(let[1]), style=pod_style)
                                    s18 += float(let[0])
                                    k18 += float(let[1])
                                    sheet.write(firstrow, 16, s18, style=razrada_style)
                                    sheet.write(firstrow, 17, k18, style=razrada_style)
                                    su18 += float(let[0])
                                    ku18 += float(let[1])
                                    sheet.write(firstrowG, 16, su18, style=kategorije_style)
                                    sheet.write(firstrowG, 17, ku18, style=kategorije_style)
                                    suu18 += float(let[0])
                                    kuu18 += float(let[1])
                                    sheet.write(sumarum, 16, suu18, style=special_style)
                                    sheet.write(sumarum, 17, kuu18, style=special_style)
                                tmp = prod
                                row += 1
                                # elif godina == '2014':
                                #     sheet.write(row, 14, float(let[0]))
                                #     sheet.write(row, 15, float(let[1]))
                                #     s14 += float(let[0])
                                #     k14 += float(let[1])
                                #     sheet.write(firstrow, 14, s14, style=razrada_style)
                                #     sheet.write(firstrow, 15, k14, style=razrada_style)
                                #     su14 += float(let[0])
                                #     ku14 += float(let[1])
                                #     sheet.write(firstrowG, 14, su14, style=kategorije_style)
                                #     sheet.write(firstrowG, 15, ku14, style=kategorije_style)
                                #     suu14 += float(let[0])
                                #     kuu14 += float(let[1])
                                #     sheet.write(sumarum, 14, suu14, style=special_style)
                                #     sheet.write(sumarum, 15, kuu14, style=special_style)
                                #     tmp = prod
                                # elif godina == '2015':
                                #     sheet.write(row, 16, float(let[0]))
                                #     sheet.write(row, 17, float(let[1]))
                                #     s15 += float(let[0])
                                #     k15 += float(let[1])
                                #     sheet.write(firstrow, 16, s15, style=razrada_style)
                                #     sheet.write(firstrow, 17, k15, style=razrada_style)
                                #     su15 += float(let[0])
                                #     ku15 += float(let[1])
                                #     sheet.write(firstrowG, 16, su15, style=kategorije_style)
                                #     sheet.write(firstrowG, 17, ku15, style=kategorije_style)
                                #     suu15 += float(let[0])
                                #     kuu15 += float(let[1])
                                #     sheet.write(sumarum, 16, suu15, style=special_style)
                                #     sheet.write(sumarum, 17, kuu15, style=special_style)
                                #     tmp = prod
                                # elif godina == '2016':
                                #     sheet.write(row, 18, float(let[0]))
                                #     sheet.write(row, 19, float(let[1]))
                                #     s16 += float(let[0])
                                #     k16 += float(let[1])
                                #     sheet.write(firstrow, 18, s16, style=razrada_style)
                                #     sheet.write(firstrow, 19, k16, style=razrada_style)
                                #     su16 += float(let[0])
                                #     ku16 += float(let[1])
                                #     sheet.write(firstrowG, 18, su16, style=kategorije_style)
                                #     sheet.write(firstrowG, 19, ku16, style=kategorije_style)
                                #     suu16 += float(let[0])
                                #     kuu16 += float(let[1])
                                #     sheet.write(sumarum, 18, suu16, style=special_style)
                                #     sheet.write(sumarum, 19, kuu16, style=special_style)
                                #     tmp = prod
                                # else:
                                #     sheet.write(row, 20, float(let[0]))
                                #     sheet.write(row, 21, float(let[1]))
                                #     s17 += float(let[0])
                                #     k17 += float(let[1])
                                #     sheet.write(firstrow, 20, s17, style=razrada_style)
                                #     sheet.write(firstrow, 21, k17, style=razrada_style)
                                #     su17 += float(let[0])
                                #     ku17 += float(let[1])
                                #     sheet.write(firstrowG, 20, su17, style=kategorije_style)
                                #     sheet.write(firstrowG, 21, ku17, style=kategorije_style)
                                #     suu17 += float(let[0])
                                #     kuu17 += float(let[1])
                                #     sheet.write(sumarum, 20, suu17, style=special_style)
                                #     sheet.write(sumarum, 21, kuu17, style=special_style)
                                #     tmp = prod
                                # row += 1


        grupe = []
    book.save("/home/matija/Desktop/Analiza_prodaje.xls")
print "gotovo"




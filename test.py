#! /usr/bin/env python
# -*- coding: utf-8 -*-

import xlwt
import csv
import sys
datas = {}
podaci = []
with open('/home/matija/Desktop/priori.csv', 'r') as myfile:
    data = csv.reader(myfile)
    #sortiranje po kupcu i godini unutar kupca
    for d in data:
        #ako postoji kupac upiši ako ne kreiraj novi zapis
        if d[1] in datas:
            #ako postoji grupa unutar kupca dodaj ako ne kreiraj novu godinu
            if d[2] in datas[d[1]]:
                datas[d[1]][d[2]].append((d[3], d[4], d[5], d[6], d[7], d[8], d[9], d[10],d[11]))
            else:
                datas[d[1]][d[2]] = [(d[3], d[4], d[5], d[6], d[7], d[8], d[9], d[10], d[11])]
        else:
            datas[d[1]] = {d[2]: [(d[3], d[4], d[5], d[6], d[7], d[8], d[9], d[10], d[11])]}
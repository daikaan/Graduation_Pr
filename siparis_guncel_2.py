#!/usr/bin/env python
# coding: utf-8

# In[ ]:


import pulp as plp
from pulp import *
import pandas as pd
import xlwings as xw
import numpy as np
from xlwings import Range, constants

wb = xw.Book("Proje_Python.xlsm")
sheet = wb.sheets["Materials"]
sheet2 = wb.sheets["Order_Inventory Plan"]
sheet3 = wb.sheets["Forecasting"]
sheet4 = wb.sheets["Forecast Results"]
sheet2.clear()
sheet4.clear()

N = len(np.array(sheet.range("K1:XFD1").value))-len(np.array(sheet.range("K1:XFD1").value)[np.array(sheet.range("K1:XFD1").value) == None])
number_of_materials = np.count_nonzero(sheet.range("A2:A10000").value)
wcap = 26000
set_E = range(1,N+1)
D = []
OQ = []
IQ = []
BC = []
SC = []
TC = []
for i in np.arange(number_of_materials):
    s1 = int(sheet.range((2+i,2)).value)
    s2 = int(sheet.range((2+i,3)).value)
    c = sheet.range((2+i,4)).value
    h = sheet.range((2+i,5)).value
    density = sheet.range((2+i,6)).value
    t = sheet.range((2+i,7)).value
    dcap = sheet.range((2+i,8)).value
    ccap1 = sheet.range((2+i,9)).value
    ccap2 = sheet.range((2+i,10)).value
    for k in np.arange(N):
        D.append(sheet.range((2+i,12+k)).value)
    D = np.array(D)
    D = D.astype(int)
    opt_model = plp.LpProblem("LP_Model")
    Q = {z: 
    plp.LpVariable(cat=plp.LpInteger,
                   lowBound=0,
                   name="Q_{0}".format(z)) 
    for z in set_E}
    SL = {z: 
    plp.LpVariable(cat=plp.LpInteger,
                   lowBound=0,
                   name="SL_{0}".format(z)) 
    for z in set_E}
    VM = {z: 
    plp.LpVariable(cat=plp.LpInteger,
                   lowBound=0,
                   name="VM_{0}".format(z)) 
    for z in set_E}
    I = {z: 
    plp.LpVariable(cat=plp.LpInteger,
                   lowBound=0,
                   name="I_{0}".format(z)) 
    for z in range(0,N+1)}
    I[0]=t
    weight_constraint = {z:opt_model.addConstraint(
    plp.LpConstraint (
           e = Q[z]-(SL[z]+VM[z])*wcap, 
           sense=plp.LpConstraintLE,
           rhs= 0,
           name="con_constraint_{0}".format(z)))
    for z in set_E}
    inventory_constraint = {z: opt_model.addConstraint(
    plp.LpConstraint(e= Q[z+1]+I[z]-I[z+1],
                     sense= plp.LpConstraintEQ,
                     rhs= D[z],
                     name= "inv_constraint_{0}".format(z)))
    for z in range(0,N)}
    volume_constraint = {z:opt_model.addConstraint(
    plp.LpConstraint (
           e = Q[z]*1/density-(SL[z]*ccap1+VM[z]*ccap2), 
           sense=plp.LpConstraintLE,
           rhs= 0,
           name="wei_constraint_{0}".format(z)))
    for z in set_E}
    depot_constraint = {z: opt_model.addConstraint(
    plp.LpConstraint(e = Q[z+1]+I[z],
                     sense= plp.LpConstraintLE,
                     rhs= dcap,
                     name= "dcap_constraint_{0}".format(z)))
    for z in range(0,N)}
    opt_model += I[0]<=0
    opt_model += plp.lpSum (SL[z]*s1+VM[z]*s2+I[z]*h+Q[z]*c
                      for z in set_E)
    opt_model.sense = plp.LpMinimize
    status=opt_model.solve(GLPK_CMD(timeLimit = 10))
    LpStatus[status]
    qua = []
    inv = []
    bco = []
    sco = []
    for t in (np.arange(N)+1):
        qua.append(value(Q[t]))
        inv.append(value(I[t]))
        bco.append(value(SL[t]))
        sco.append(value(VM[t]))
    OQ.append(qua)
    IQ.append(inv)
    BC.append(bco)
    SC.append(sco)
    TC.append(value(opt_model.objective))
    qua = []
    inv = []
    bco = []
    sco = []
    D = []
OQ = np.array(OQ)
IQ = np.array(IQ)
BC = np.array(BC)
SC = np.array(SC)
TC = np.array(TC)
for i in np.arange(number_of_materials):
    sheet2.range((3+i,1)).value = sheet.range((2+i,1)).value
    sheet2.range((3+i,1)).api.HorizontalAlignment = constants.HAlign.xlHAlignCenter
    sheet2.range((3+i,1)).api.Borders.LineStyle = 1
    sheet2.range((3+i,1)).api.Font.Bold = True
for i in np.arange(4*N): 
    sheet2.range((1,2+i)).value = (i%N)+1
    sheet2.range((1,2+i)).api.HorizontalAlignment = constants.HAlign.xlHAlignCenter
    sheet2.range((1,2+i)).api.Borders.LineStyle = 1
    sheet2.range((1,2+i)).api.Font.Bold = True
for k in np.arange(N):
    sheet2.range((2,2+k)).value = "Q"
    sheet2.range((2,2+k)).api.HorizontalAlignment = constants.HAlign.xlHAlignCenter
    sheet2.range((2,2+k)).api.Borders.LineStyle = 1
    sheet2.range((2,2+k)).api.Interior.ColorIndex = 4
    sheet2.range((2,2+k)).api.Font.Bold = True
    sheet2.range((2,(N+2+k))).value = "I"
    sheet2.range((2,(N+2+k))).api.HorizontalAlignment = constants.HAlign.xlHAlignCenter
    sheet2.range((2,(N+2+k))).api.Borders.LineStyle = 1
    sheet2.range((2,(N+2+k))).api.Interior.ColorIndex = 44
    sheet2.range((2,(N+2+k))).api.Font.Bold = True
    sheet2.range((2,(2*N+2+k))).value = "SL"
    sheet2.range((2,(2*N+2+k))).api.HorizontalAlignment = constants.HAlign.xlHAlignCenter
    sheet2.range((2,(2*N+2+k))).api.Borders.LineStyle = 1
    sheet2.range((2,(2*N+2+k))).api.Interior.ColorIndex = 37
    sheet2.range((2,(2*N+2+k))).api.Font.Bold = True
    sheet2.range((2,(3*N+2+k))).value = "VM"
    sheet2.range((2,(3*N+2+k))).api.HorizontalAlignment = constants.HAlign.xlHAlignCenter
    sheet2.range((2,(3*N+2+k))).api.Borders.LineStyle = 1
    sheet2.range((2,(3*N+2+k))).api.Interior.ColorIndex = 22
    sheet2.range((2,(3*N+2+k))).api.Font.Bold = True
for i in np.arange(number_of_materials):
    for k in np.arange(N):
        sheet2.range((3+i,2+k)).value = OQ[i][k]
        sheet2.range((3+i,2+k)).api.HorizontalAlignment = constants.HAlign.xlHAlignCenter
        sheet2.range((3+i,2+k)).api.Borders.LineStyle = 1
        sheet2.range((3+i,2+k)).api.Interior.ColorIndex = 4
        sheet2.range((3+i,(N+2+k))).value = IQ[i][k]
        sheet2.range((3+i,(N+2+k))).api.HorizontalAlignment = constants.HAlign.xlHAlignCenter
        sheet2.range((3+i,(N+2+k))).api.Borders.LineStyle = 1
        sheet2.range((3+i,(N+2+k))).api.Interior.ColorIndex = 44
        sheet2.range((3+i,(2*N+2+k))).value = BC[i][k]
        sheet2.range((3+i,(2*N+2+k))).api.HorizontalAlignment = constants.HAlign.xlHAlignCenter
        sheet2.range((3+i,(2*N+2+k))).api.Borders.LineStyle = 1
        sheet2.range((3+i,(2*N+2+k))).api.Interior.ColorIndex = 37
        sheet2.range((3+i,(3*N+2+k))).value = SC[i][k]
        sheet2.range((3+i,(3*N+2+k))).api.HorizontalAlignment = constants.HAlign.xlHAlignCenter
        sheet2.range((3+i,(3*N+2+k))).api.Borders.LineStyle = 1
        sheet2.range((3+i,(3*N+2+k))).api.Interior.ColorIndex = 22
    sheet2.range((3+i,(2+4*N))).value = TC[i]
    sheet2.range((3+i,(2+4*N))).api.HorizontalAlignment = constants.HAlign.xlHAlignCenter
    sheet2.range((3+i,(2+4*N))).api.Borders.LineStyle = 1
    sheet2.range((3+i,(2+4*N))).api.Interior.ColorIndex = 36
    sheet2.range((3+i,(2+4*N))).number_format = sheet.range((i+2,4)).number_format
sheet2.range((2,(2+4*N))).value = "Total Cost"
sheet2.range((2,(2+4*N))).api.HorizontalAlignment = constants.HAlign.xlHAlignCenter
sheet2.range((2,(2+4*N))).api.Borders.LineStyle = 1
sheet2.range((2,(2+4*N))).api.Interior.ColorIndex = 36
sheet2.range((2,(2+4*N))).api.Font.Bold = True


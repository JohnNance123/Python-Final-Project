
#INSERTS THE MAX DATA TO THE EXCEL FILE. THIS IS FOR ALL PRODUCTS
def maxCalc(sheet, prod, maxTP, maxTS, maxRP, maxRS):

    sheet['I11'].value = prod
    sheet['J10'].value = "Maximum Review Title Polarity"
    sheet['J11'].value = maxTP

    sheet['I20'].value = prod
    sheet['J19'].value = "Maximum Review Title Subjectivity"
    sheet['J20'].value = maxTS

    sheet['I29'].value = prod
    sheet['J28'].value = "Maximum Review Polarity"
    sheet['J29'].value = maxRP

    sheet['I38'].value = prod
    sheet['J37'].value = "Maximum Review Subjectivity"
    sheet['J38'].value = maxRS

#INSERTS THE MAX DATA TO THE EXCEL FILE. THIS IS FOR PRODUCT 1   
def maxCalc1(sheet, prod, maxTP, maxTS, maxRP, maxRS):

    sheet['I12'].value = prod
    sheet['J12'].value = maxTP

    sheet['I21'].value = prod
    sheet['J21'].value = maxTS

    sheet['I30'].value = prod
    sheet['J30'].value = maxRP

    sheet['I39'].value = prod
    sheet['J39'].value = maxRS

#INSERTS THE MAX DATA TO THE EXCEL FILE. THIS IS FOR PRODUCT 2   
def maxCalc2(sheet, prod, maxTP, maxTS, maxRP, maxRS):

    sheet['I13'].value = prod
    sheet['J13'].value = maxTP

    sheet['I22'].value = prod
    sheet['J22'].value = maxTS

    sheet['I31'].value = prod
    sheet['J31'].value = maxRP

    sheet['I40'].value = prod
    sheet['J40'].value = maxRS

#INSERTS THE MAX DATA TO THE EXCEL FILE. THIS IS FOR PRODUCT 3
def maxCalc3(sheet, prod, maxTP, maxTS, maxRP, maxRS):

    sheet['I14'].value = prod
    sheet['J14'].value = maxTP

    sheet['I23'].value = prod
    sheet['J23'].value = maxTS

    sheet['I32'].value = prod
    sheet['J32'].value = maxRP

    sheet['I41'].value = prod
    sheet['J41'].value = maxRS

#INSERTS THE MAX DATA TO THE EXCEL FILE. THIS IS FOR PRODUCT 4
def maxCalc4(sheet, prod, maxTP, maxTS, maxRP, maxRS):

    sheet['I15'].value = prod
    sheet['J15'].value = maxTP

    sheet['I24'].value = prod
    sheet['J24'].value = maxTS

    sheet['I33'].value = prod
    sheet['J33'].value = maxRP

    sheet['I42'].value = prod
    sheet['J42'].value = maxRS

#INSERTS THE MAX DATA TO THE EXCEL FILE. THIS IS FOR PRODUCT 5
def maxCalc5(sheet, prod, maxTP, maxTS, maxRP, maxRS):

    sheet['I16'].value = prod
    sheet['J16'].value = maxTP

    sheet['I25'].value = prod
    sheet['J25'].value = maxTS

    sheet['I34'].value = prod
    sheet['J34'].value = maxRP

    sheet['I43'].value = prod
    sheet['J43'].value = maxRS

#INSERTS THE MAX DATA TO THE EXCEL FILE. THIS IS FOR PRODUCT 6
def maxCalc6(sheet, prod, maxTP, maxTS, maxRP, maxRS):

    sheet['I17'].value = prod
    sheet['J17'].value = maxTP

    sheet['I26'].value = prod
    sheet['J26'].value = maxTS

    sheet['I35'].value = prod
    sheet['J35'].value = maxRP

    sheet['I44'].value = prod
    sheet['J44'].value = maxRS
    


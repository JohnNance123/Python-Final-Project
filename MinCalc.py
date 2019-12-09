
#INSERTS THE MINIMUM VALUES INTO EXCELL FILE FOR ALL PRODUCTS 
def minCalc(sheet, prod, minTP, minTS, minRP, minRS):

    sheet['F11'].value = prod
    sheet['G10'].value = "Minimum Review Title Polarity"
    sheet['G11'].value = minTP

    sheet['F20'].value = prod
    sheet['G19'].value = "Minimum Review Title Subjectivity"
    sheet['G20'].value = minTS

    sheet['F29'].value = prod
    sheet['G28'].value = "Minimum Review Polarity"
    sheet['G29'].value = minRP

    sheet['F38'].value = prod
    sheet['G37'].value = "Mimimum Review Subjectivity"
    sheet['G38'].value = minRS

#INSERTS THE MINIMUM VALUES INTO EXCELL FILE FOR PRODUCT 1
def minCalc1(sheet, prod, minTP, minTS, minRP, minRS):

    sheet['F12'].value = prod
    sheet['G12'].value = minTP

    sheet['F21'].value = prod
    sheet['G21'].value = minTS

    sheet['F30'].value = prod
    sheet['G30'].value = minRP

    sheet['F39'].value = prod
    sheet['G39'].value = minRS

#INSERTS THE MINIMUM VALUES INTO EXCELL FILE FOR PRODUCT 2
def minCalc2(sheet, prod, minTP, minTS, minRP, minRS):

    sheet['F13'].value = prod
    sheet['G13'].value = minTP

    sheet['F22'].value = prod
    sheet['G22'].value = minTS

    sheet['F31'].value = prod
    sheet['G31'].value = minRP

    sheet['F40'].value = prod
    sheet['G40'].value = minRS

#INSERTS THE MINIMUM VALUES INTO EXCELL FILE FOR PRODUCT 3 
def minCalc3(sheet, prod, minTP, minTS, minRP, minRS):

    sheet['F14'].value = prod
    sheet['G14'].value = minTP

    sheet['F23'].value = prod
    sheet['G23'].value = minTS

    sheet['F32'].value = prod
    sheet['G32'].value = minRP

    sheet['F41'].value = prod
    sheet['G41'].value = minRS

#INSERTS THE MINIMUM VALUES INTO EXCELL FILE FOR PRODUCT 4  
def minCalc4(sheet, prod, minTP, minTS, minRP, minRS):

    sheet['F15'].value = prod
    sheet['G15'].value = minTP

    sheet['F24'].value = prod
    sheet['G24'].value = minTS

    sheet['F33'].value = prod
    sheet['G33'].value = minRP

    sheet['F42'].value = prod
    sheet['G42'].value = minRS

#INSERTS THE MINIMUM VALUES INTO EXCELL FILE FOR PRODUCT 5
def minCalc5(sheet, prod, minTP, minTS, minRP, minRS):

    sheet['F16'].value = prod
    sheet['G16'].value = minTP

    sheet['F25'].value = prod
    sheet['G25'].value = minTS

    sheet['F34'].value = prod
    sheet['G34'].value = minRP

    sheet['F43'].value = prod
    sheet['G43'].value = minRS

#INSERTS THE MINIMUM VALUES INTO EXCELL FILE FOR PRODUCT 6
def minCalc6(sheet, prod, minTP, minTS, minRP, minRS):

    sheet['F17'].value = prod
    sheet['G17'].value = minTP

    sheet['F26'].value = prod
    sheet['G26'].value = minTS

    sheet['F35'].value = prod
    sheet['G35'].value = minRP

    sheet['F44'].value = prod
    sheet['G44'].value = minRS
    

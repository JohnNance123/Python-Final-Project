from openpyxl import *

#INSETS THE AVERAGES INTO EXCEL FILE FOR ALL PRODUCTS
def prodExcel(sheet, prod, star, titleP, titleS, reviewP, reviewS):
    
    sheet['D1'].value = "Average Star Rating"
    sheet['C2'].value = prod
    sheet['D2'].value = star

    sheet['D10'].value = "Average Review Title Polarity"
    sheet['C11'].value = prod
    sheet['D11'].value = titleP

    sheet['D19'].value = "Average Review Title Subjectivity"
    sheet['C20'].value = prod
    sheet['D20'].value = titleS

    sheet['D28'].value = "Average Review Polarity"
    sheet['C29'].value = prod
    sheet['D29'].value = reviewP

    sheet['D37'].value = "Average Review Subjectivity"
    sheet['C38'].value = prod
    sheet['D38'].value = reviewS

#INSETS THE AVERAGES INTO EXCEL FILE FOR PRODUCT 1
def prodExcel1(sheet, prod, star, titleP, titleS, reviewP, reviewS):

    sheet['C3'].value = prod
    sheet['D3'].value = star

    sheet['C12'].value = prod
    sheet['D12'].value = titleP

    sheet['C21'].value = prod
    sheet['D21'].value = titleS

    sheet['C30'].value = prod
    sheet['D30'].value = reviewP

    sheet['C39'].value = prod
    sheet['D39'].value = reviewS

#INSETS THE AVERAGES INTO EXCEL FILE FOR PRODUCT 2
def prodExcel2(sheet, prod, star, titleP, titleS, reviewP, reviewS):

    sheet['C4'].value = prod
    sheet['D4'].value = star
    
    sheet['C13'].value = prod
    sheet['D13'].value = titleP

    sheet['C22'].value = prod
    sheet['D22'].value = titleS

    sheet['C31'].value = prod
    sheet['D31'].value = reviewP

    sheet['C40'].value = prod
    sheet['D40'].value = reviewS

#INSETS THE AVERAGES INTO EXCEL FILE FOR PRODUCT 3   
def prodExcel3(sheet, prod, star, titleP, titleS, reviewP, reviewS):

    sheet['C5'].value = prod
    sheet['D5'].value = star
    
    sheet['C14'].value = prod
    sheet['D14'].value = titleP

    sheet['C23'].value = prod
    sheet['D23'].value = titleS

    sheet['C32'].value = prod
    sheet['D32'].value = reviewP

    sheet['C41'].value = prod
    sheet['D41'].value = reviewS

#INSETS THE AVERAGES INTO EXCEL FILE FOR PRODUCT 4    
def prodExcel4(sheet, prod, star, titleP, titleS, reviewP, reviewS):

    sheet['C6'].value = prod
    sheet['D6'].value = star
    
    sheet['C15'].value = prod
    sheet['D15'].value = titleP

    sheet['C24'].value = prod
    sheet['D24'].value = titleS

    sheet['C33'].value = prod
    sheet['D33'].value = reviewP

    sheet['C42'].value = prod
    sheet['D42'].value = reviewS

#INSETS THE AVERAGES INTO EXCEL FILE FOR PRODUCT 5   
def prodExcel5(sheet, prod, star, titleP, titleS, reviewP, reviewS):

    sheet['C7'].value = prod
    sheet['D7'].value = star
    
    sheet['C16'].value = prod
    sheet['D16'].value = titleP

    sheet['C25'].value = prod
    sheet['D25'].value = titleS

    sheet['C34'].value = prod
    sheet['D34'].value = reviewP

    sheet['C43'].value = prod
    sheet['D43'].value = reviewS

#INSETS THE AVERAGES INTO EXCEL FILE FOR PRODUCT 6   
def prodExcel6(sheet, prod, star, titleP, titleS, reviewP, reviewS):

    sheet['C8'].value = prod
    sheet['D8'].value = star
    
    sheet['C17'].value = prod
    sheet['D17'].value = titleP

    sheet['C26'].value = prod
    sheet['D26'].value = titleS

    sheet['C35'].value = prod
    sheet['D35'].value = reviewP

    sheet['C44'].value = prod
    sheet['D44'].value = reviewS

#INSETS THE STANDARD DEVIATION INTO EXCEL FILE FOR ALL PRODUCTS
def stdev(sheet, prod, star, titleP, titleS, reviewP, reviewS):

    sheet['M1'].value = "Standard Deviation of Star Rating"
    sheet['L2'].value = prod
    sheet['M2'].value = star

    sheet['M10'].value = "Standard Deviation of Review Title Polarity"
    sheet['L11'].value = prod
    sheet['M11'].value = titleP

    sheet['M19'].value = "Standard Deviation of Review Title Subjectivity"
    sheet['L20'].value = prod
    sheet['M20'].value = titleS

    sheet['M28'].value = "Standard Deviation of Review Polarity"
    sheet['L29'].value = prod
    sheet['M29'].value = reviewP

    sheet['M37'].value = "Standard Deviation of Review Subjectivity"
    sheet['L38'].value = prod
    sheet['M38'].value = reviewS
    
#INSERTS THE STANDARD DEVIATION INTO EXCEL FILE FOR PRODUCT 1   
def stdev1(sheet, prod, star, titleP, titleS, reviewP, reviewS):

    sheet['L3'].value = prod
    sheet['M3'].value = star

    sheet['L12'].value = prod
    sheet['M12'].value = titleP

    sheet['L21'].value = prod
    sheet['M21'].value = titleS
    
    sheet['L30'].value = prod
    sheet['M30'].value = reviewP
    
    sheet['L39'].value = prod
    sheet['M39'].value = reviewP

#INSERTS THE STANDARD DEVIATION INTO EXCEL FILE FOR PRODUCT 2    
def stdev2(sheet, prod, star, titleP, titleS, reviewP, reviewS):

    sheet['L4'].value = prod
    sheet['M4'].value = star

    sheet['L13'].value = prod
    sheet['M13'].value = titleP

    sheet['L22'].value = prod
    sheet['M22'].value = titleS
    
    sheet['L31'].value = prod
    sheet['M31'].value = reviewP
    
    sheet['L40'].value = prod
    sheet['M40'].value = reviewP

#INSERTS THE STANDARD DEVIATION INTO EXCEL FILE FOR PRODUCT 3    
def stdev3(sheet, prod, star, titleP, titleS, reviewP, reviewS):

    sheet['L5'].value = prod
    sheet['M5'].value = star

    sheet['L14'].value = prod
    sheet['M14'].value = titleP

    sheet['L23'].value = prod
    sheet['M23'].value = titleS
    
    sheet['L32'].value = prod
    sheet['M32'].value = reviewP
    
    sheet['L41'].value = prod
    sheet['M41'].value = reviewP

#INSERTS THE STANDARD DEVIATION INTO EXCEL FILE FOR PRODUCT 4
def stdev4(sheet, prod, star, titleP, titleS, reviewP, reviewS):

    sheet['L6'].value = prod
    sheet['M6'].value = star

    sheet['L15'].value = prod
    sheet['M15'].value = titleP

    sheet['L24'].value = prod
    sheet['M24'].value = titleS
    
    sheet['L33'].value = prod
    sheet['M33'].value = reviewP
    
    sheet['L42'].value = prod
    sheet['M42'].value = reviewP

#INSERTS THE STANDARD DEVIATION INTO EXCEL FILE FOR PRODUCT 5   
def stdev5(sheet, prod, star, titleP, titleS, reviewP, reviewS):

    sheet['L7'].value = prod
    sheet['M7'].value = star

    sheet['L16'].value = prod
    sheet['M16'].value = titleP

    sheet['L25'].value = prod
    sheet['M25'].value = titleS
    
    sheet['L34'].value = prod
    sheet['M34'].value = reviewP
    
    sheet['L43'].value = prod
    sheet['M43'].value = reviewP

#INSERTS THE STANDARD DEVIATION INTO EXCEL FILE FOR PRODUCT 6
def stdev6(sheet, prod, star, titleP, titleS, reviewP, reviewS):

    sheet['L8'].value = prod
    sheet['M8'].value = star

    sheet['L17'].value = prod
    sheet['M17'].value = titleP

    sheet['L26'].value = prod
    sheet['M26'].value = titleS
    
    sheet['L35'].value = prod
    sheet['M35'].value = reviewP
    
    sheet['L44'].value = prod
    sheet['M44'].value = reviewP

    

from reportlab.lib import colors
from reportlab.lib.colors import Color
from reportlab.lib.units import inch
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.utils import ImageReader
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, PageTemplate, PageBreak, Frame, Spacer
from datetime import datetime
import locale
import webbrowser

locale.setlocale(locale.LC_TIME, "en_US.UTF-8")
locale.setlocale(locale.LC_ALL, 'en_US.UTF-8')

pageWidth = 612
pageHeight = 792
margin = 36
marginLeft = 36
marginRight = pageWidth - marginLeft

def drawHeader(c, key, valueList):
    # Logo
    logoPath = "assets/logo.png"
    logo = ImageReader(logoPath)

    logoWidth = 83
    logoHeight = 65
    logoX = pageWidth - 36 - logoWidth
    logoY = 745 - logoHeight

    c.drawImage(logo, logoX, logoY, width=logoWidth, height=logoHeight)

    # Name
    c.setFont("Helvetica", 20)
    c.setFillColor(Color(0, 0, 0))
    c.drawString(marginLeft, 749 - 20, f"{valueList[0]}") # Y=729

    # Account Number, Account Type (PERMANENT)
    c.setFont("Helvetica-Bold", 12)
    c.setFillColor(Color(0, 0, 0))
    c.drawString(marginLeft, 711, "Account Number: ")
    c.drawString(marginLeft, 695, "Account Type: ")
    c.drawString(marginLeft, 679, "Plan Option: ")

    # Account Number, Account Type (VARIABLES)
    c.setFont("Helvetica", 12)
    c.setFillColor(Color(0, 0, 0))
    c.drawString(140, 711, f"{key}")
    c.drawString(124, 695, "Non-Retirement")
    c.drawString(110, 679, f"Spectra-{valueList[10]}")

    # PORFOLIO SUMMARY HEADER
    c.setFont("Helvetica", 20)
    c.setFillColor(Color(0, 0, 0))
    c.drawString(marginLeft, 634, "PORTFOLIO SUMMARY")

    # Statement Period (PERMANENT)
    c.setFont("Helvetica-Bold", 8)
    c.setFillColor(Color(0, 0, 0))
    c.drawString(marginRight - 145, 634, "Statement Period: ")
    
    # Statement Period (VARIABLES)
    c.setFont("Helvetica", 8)
    c.setFillColor(Color(0, 0, 0))
    c.drawString(marginRight - 73, 634, "Ending Oct 31, 2023")
    
    # Line 1 (BELOW PORTFOLIO SUMMARY)
    c.setLineWidth(1)
    c.setStrokeColor(Color(0, 0, 0))
    c.line(marginLeft, 630, marginRight, 630)

    # Client Account Information (PERMANENT)
    # Below 1st line
    c.setFont("Helvetica", 12)
    c.setFillColor(Color(0, 0, 0))
    c.drawString(marginLeft, 624 - 12, "Acceptance Date: ")

    c.setFont("Helvetica", 12)
    c.setFillColor(Color(0, 0, 0))
    c.drawString(marginLeft, 607 - 12, "Principal Amount: ")

    c.setFont("Helvetica", 12)
    c.setFillColor(Color(0, 0, 0))
    c.drawString(marginLeft, 590 - 12, "Monthly Interest: ")

    c.setFont("Helvetica", 12)
    c.setFillColor(Color(0, 0, 0))
    c.drawString(marginLeft, 573 - 12, "Annual Bonus: ")

    c.setFont("Helvetica", 12)
    c.setFillColor(Color(0, 0, 0))
    c.drawString(marginLeft, 556 - 12, "Total Interest Earned: ")

    c.setFont("Helvetica", 12)
    c.setFillColor(Color(0, 0, 0))
    c.drawString(marginLeft, 539 - 12, "Interest Reinvestment: ")

    c.setFont("Helvetica", 12)
    c.setFillColor(Color(0, 0, 0))
    c.drawString(marginLeft, 522 - 12, "Bonus Reinvestment: ")

    c.setFont("Helvetica-Bold", 12)
    c.setFillColor(Color(0, 0, 0))
    c.drawString(marginLeft, 505 - 12, "Closing Balance: ")

    # Client Account Information (VARIABLES)
    originalDate = str(valueList[2])
    dateObject = datetime.strptime(originalDate, "%Y-%m-%d %H:%M:%S")
    formattedDate = dateObject.strftime("%B %d, %Y")

    c.setFont("Helvetica", 12)
    c.setFillColor(Color(0, 0, 0))
    c.drawRightString(marginRight, 624 - 12, f"{formattedDate}")

    originalPrincipal = float(valueList[3])
    formattedPrincipal = locale.currency(originalPrincipal, grouping=True)
    c.setFont("Helvetica", 12)
    c.setFillColor(Color(0, 0, 0))
    c.drawRightString(marginRight, 607 - 12, f"{formattedPrincipal}") # Principal

    originalMonthlyInt = float(valueList[4])
    formattedMonthlyInt = locale.currency(originalMonthlyInt, grouping=True)
    c.setFont("Helvetica", 12)
    c.setFillColor(Color(0, 0, 0))
    c.drawRightString(marginRight, 590 - 12, f"{formattedMonthlyInt}") # Monthly Interest

    originalAnnualBonus = float(valueList[5])
    formattedAnnualBonus = locale.currency(originalAnnualBonus, grouping=True)
    c.setFont("Helvetica", 12)
    c.setFillColor(Color(0, 0, 0))
    c.drawRightString(marginRight, 573 - 12, f"{formattedAnnualBonus}") # Annual Bonus

    originalTotalInt = float(valueList[8])
    formattedTotalInt = locale.currency(originalTotalInt, grouping=True)
    c.setFont("Helvetica", 12)
    c.setFillColor(Color(0, 0, 0))
    c.drawRightString(marginRight, 556 - 12, f"{formattedTotalInt}") # Total Interest Earned

    if valueList[6] == "v":
        IRstr = "Yes"
    else:
        IRstr = "No"

    c.setFont("Helvetica", 12)
    c.setFillColor(Color(0, 0, 0))
    c.drawRightString(marginRight, 539 - 12, f"{IRstr}")

    if valueList[7] == "v":
        BRstr = "Yes"
    else:
        BRstr = "No"

    c.setFont("Helvetica", 12)
    c.setFillColor(Color(0, 0, 0))
    c.drawRightString(marginRight, 522 - 12, f"{BRstr}")

    if isinstance(valueList[9], float):
        originalEndingBal = str(valueList[9])
        originalEndingBal = originalEndingBal.replace("$", "").replace(",", "")
    else:
        originalEndingBal = valueList[9].replace("$", "").replace(",", "")

    originalEndingBal = float(originalEndingBal)
    formattedEndingBal = locale.currency(originalEndingBal, grouping=True)
    c.setFont("Helvetica-Bold", 12)
    c.setFillColor(Color(0, 0, 0))
    c.drawRightString(marginRight, 505 - 12, f"{formattedEndingBal}") # Ending Balance

    # TRANSACTION HISTORY HEADER
    c.setFont("Helvetica", 20)
    c.setFillColor(Color(0, 0, 0))
    c.drawString(marginLeft, 448, "TRANSACTION HISTORY")

    # Line 2 (BELOW TRANSACTION HISTORY HEADER)
    c.setLineWidth(1)
    c.setStrokeColor(Color(0, 0, 0))
    c.line(marginLeft, 444, marginRight, 444)

def drawFooter(c):
    # Footer line
    c.setLineWidth(1)
    c.setStrokeColor(Color(0, 0, 0))
    c.line(marginLeft, 76, marginRight, 76)

    contactStyle = ParagraphStyle(
        "contactStyle",
        fontSize=6,
        leading=12,
        textColor=colors.black,
        alignment=0,
    )

    addressStyle = ParagraphStyle(
        "addressStyle",
        fontSize=6,
        leading=12,
        textColor=colors.black,
        alignment=2,
    )

    contactInfo = Paragraph(
    """
    Client Relation Email: Support@wsimo.com
    <br/>
    Client Service Tel: (949) 447-5900
    <br/>
    Hours: Monday - Friday, 8:00am - 5:00pm PST
    """, contactStyle)

    addressInfo = Paragraph(
    """
    Wealth Space Asset Management, LLC
    <br/>
    1 Park Plaza, Suite 600, Irvine, CA 92614
    """, addressStyle)

    contactInfoWidth, contactInfoHeight = contactInfo.wrap(pageWidth / 2, 52)
    addressInfoWidth, addressInfoHeight = addressInfo.wrap(pageWidth / 2, 52)

    contactInfo.drawOn(c, marginLeft, 32)
    addressInfo.drawOn(c, marginRight - addressInfoWidth, 44)

def generatePDF(key, valueList, clientFreq, currentCount, clientTransactionHistory):
    customGreen = Color(182.0/255, 221.0/255, 100.0/255)  
    customGold = Color(215.0/255, 200.0/255, 119.0/255)
    customDarkGreen = Color(75.0/255, 162.0/255, 127.0/255)
    customBlue = Color(38.0/255, 152.0/255, 201.0/255)
    customRed = Color(145.0/255, 88.0/255, 155.0/255)

    count = clientFreq[valueList[0]]
    if count == 1:
        c = canvas.Canvas(f"pdfs/AAI Statement 10.31.2023 - {valueList[0]}.pdf", pagesize=letter)
    else:
        c = canvas.Canvas(f"pdfs/AAI Statement 10.31.2023 - {valueList[0]} #{currentCount}.pdf", pagesize=letter)

    tableData = [
        ["Name", "Status", "Beginning Balance", "Interest Earned", "Ending Balance"]
    ]
    
    for entry in clientTransactionHistory:
        if entry[0] is None:
            continue

        if entry[2] is None or entry[2] == "":
            entry[2] = 0
        else:
            originalBegginingBal = float(entry[2])
            formattedBeginningBal = locale.currency(originalBegginingBal, grouping=True)
            entry[2] = formattedBeginningBal

        originalInterestEarned = float(entry[3])
        formattedInterestEarned = locale.currency(originalInterestEarned, grouping=True)
        entry[3] = formattedInterestEarned

        if isinstance(entry[4], float):
            originalEndingBalTH = str(entry[4])
            originalEndingBalTH = originalEndingBalTH.replace("$", "").replace(",", "")
        else:
            if isinstance(entry[4], int):
                pass
            else:
                originalEndingBalTH = entry[4].replace("$", "").replace(",", "")

        originalEndingBalTH = float(originalEndingBalTH)
        formattedEndingBalTH = locale.currency(originalEndingBalTH, grouping=True)
        entry[4] = formattedEndingBalTH
        
        tableData.append(entry)

    drawHeader(c, key, valueList)

    firstTableData = tableData[:19]
    firstTable = Table(firstTableData)
    firstTable._argW = [110, 90, 130, 105, 105]
    firstTableStyleList = [
        ('BACKGROUND', (0, 0), (-1, 0), colors.white),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.black),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica'),
        ('FONTSIZE', (0, 0), (-1, 0), 12),
        ('BOTTOMPADDING', (0, 0), (-1, 0), 8),
        ('BACKGROUND', (0, 1), (-1, -1), colors.white),
        ('GRID', (0, 0), (-1, -1), 0.5, colors.black),
        ('FONTSIZE', (0, 1), (-1, -1), 8),
    ]
    for i, row in enumerate(firstTableData):
        for j, cell in enumerate(row):
            if cell == "Reinvestment":
                firstTableStyleList.append(('BACKGROUND', (j, i), (j, i), customGreen))
                firstTableStyleList.append(('TEXTCOLOR', (j, i), (j, i), colors.white))
            elif cell == "Annual Bonus":
                firstTableStyleList.append(('BACKGROUND', (j, i), (j, i), customGold))
                firstTableStyleList.append(('TEXTCOLOR', (j, i), (j, i), colors.white))
            elif cell == "Interest Payout":
                firstTableStyleList.append(('BACKGROUND', (j, i), (j, i), customDarkGreen))
                firstTableStyleList.append(('TEXTCOLOR', (j, i), (j, i), colors.white))
            elif cell == "Renewal":
                firstTableStyleList.append(('BACKGROUND', (j, i), (j, i), customBlue))
                firstTableStyleList.append(('TEXTCOLOR', (j, i), (j, i), colors.white))
            elif cell == "Withdrawal":
                firstTableStyleList.append(('BACKGROUND', (j, i), (j, i), customRed))
                firstTableStyleList.append(('TEXTCOLOR', (j, i), (j, i), colors.white))

    firstTableStyle = TableStyle(firstTableStyleList)
    firstTable.setStyle(firstTableStyle)

    firstTableWidth, firstTableHeight = firstTable.wrap(0, 0)
    firstTable.drawOn(c, 36, 434 - firstTableHeight)

    pageNum1 = c.getPageNumber()
    if len(tableData) > 19:
        c.setFont("Helvetica", 8)
        c.setFillColor(Color(0, 0, 0))
        c.drawRightString(marginRight, 780 - 8, f"Page {pageNum1} of 2")
    else:
        c.setFont("Helvetica", 8)
        c.setFillColor(Color(0, 0, 0))
        c.drawRightString(marginRight, 780 - 8, f"Page {pageNum1} of 1")

    drawFooter(c)

    if len(tableData) > 19:
        c.showPage()

        # TRANSACTION HISTORY CONTINUED HEADER
        c.setFont("Helvetica", 20)
        c.setFillColor(Color(0, 0, 0))
        c.drawString(marginLeft, 729, "TRANSACTION HISTORY (Continued)")

        # Line 1 (BELOW TRANSACTION HISTORY CONTINUED HEADER)
        c.setLineWidth(1)
        c.setStrokeColor(Color(0, 0, 0))
        c.line(marginLeft, 723, marginRight, 723)

        laterTableData = tableData[19:]
        laterTable = Table(laterTableData)
        laterTable._argW = [110, 90, 130, 105, 105]
        laterTableStyleList = [
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica'),
            ('BACKGROUND', (0, 1), (-1, -1), colors.white),
            ('GRID', (0, 0), (-1, -1), 0.5, colors.black),
            ('FONTSIZE', (0, 0), (-1, -1), 8),
        ]

        for i, row in enumerate(laterTableData, start=19):  # Start enumerating from currentRowIndex
            for j, cell in enumerate(row):
                if cell == "Reinvestment":
                    laterTableStyleList.append(('BACKGROUND', (j, i - 19), (j, i - 19), customGreen))
                    laterTableStyleList.append(('TEXTCOLOR', (j, i - 19), (j, i - 19), colors.white))
                elif cell == "Annual Bonus":
                    laterTableStyleList.append(('BACKGROUND', (j, i - 19), (j, i - 19), customGold))
                    laterTableStyleList.append(('TEXTCOLOR', (j, i - 19), (j, i - 19), colors.white))
                elif cell == "Interest Payout":
                    laterTableStyleList.append(('BACKGROUND', (j, i - 19), (j, i - 19), customDarkGreen))
                    laterTableStyleList.append(('TEXTCOLOR', (j, i - 19), (j, i - 19), colors.white))
                elif cell == "Renewal":
                    laterTableStyleList.append(('BACKGROUND', (j, i - 19), (j, i - 19), customBlue))
                    laterTableStyleList.append(('TEXTCOLOR', (j, i - 19), (j, i - 19), colors.white))
                elif cell == "Withdrawal":
                    laterTableStyleList.append(('BACKGROUND', (j, i - 19), (j, i - 19), customRed))
                    laterTableStyleList.append(('TEXTCOLOR', (j, i - 19), (j, i - 19), colors.white))

        laterTableStyle = TableStyle(laterTableStyleList)
        laterTable.setStyle(laterTableStyle)

        laterTableWidth, laterTableHeight = laterTable.wrap(0, 0)
        laterTable.drawOn(c, 36, 713 - laterTableHeight)

        pageNum2 = c.getPageNumber()
        c.setFont("Helvetica", 8)
        c.setFillColor(Color(0, 0, 0))
        c.drawRightString(marginRight, 780 - 8, f"Page {pageNum2} of 2")

        drawFooter(c)

    c.save()


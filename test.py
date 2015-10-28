import os
import zipfile
import shutil
import sys
import csv
from lxml import etree


outputFileName = 'mytest.pptx'
inputFileName = 'Financial_Dashboard_Powerpoint_V4_1016.pptx'


def input1():
    try:
        with open(sys.argv[1], 'r') as data1:
            reader1 = csv.reader(data1, delimiter=';')
            next(reader1, None)
            for row1 in reader1:
                yield row1
    except Exception:
        print "Reading slide 1 input error", sys.exc_info()


def input2():
    try:
        with open(sys.argv[2], 'r') as data2:
            reader2 = csv.reader(data2, delimiter=';')
            next(reader2, None)
            for row2 in reader2:
                yield row2
    except Exception:
        print "Reading slide 2 input error", sys.exc_info()


def input3():
    try:
        with open(sys.argv[3], 'r') as data3:
            reader3 = csv.reader(data3, delimiter=';')
            next(reader3, None)
            for row3 in reader3:
                yield row3
    except Exception:
        print "Reading slide 3 input error", sys.exc_info()


def openFile1(rows):
    path = os.path.join(os.path.split(os.path.abspath(__file__))[0], 'ppt/slides/slide1.xml')
    tree = etree.parse(path)

    F2 = tree.xpath('/p:sld/p:cSld/p:spTree/p:sp[22]/p:txBody/a:p/a:r/a:t', namespaces={'a':'http://schemas.openxmlformats.org/drawingml/2006/main',
            'r':'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
            'p':'http://schemas.openxmlformats.org/presentationml/2006/main'})
    F2[0].text = str(rows[16][1]).replace(' ', '')
    E2 = tree.xpath('/p:sld/p:cSld/p:spTree/p:sp[10]/p:txBody/a:p/a:r/a:t', namespaces={'a':'http://schemas.openxmlformats.org/drawingml/2006/main',
            'r':'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
            'p':'http://schemas.openxmlformats.org/presentationml/2006/main'})
    E2[0].text = str(rows[36][1]).replace(' ', '')
    E6 = tree.xpath('/p:sld/p:cSld/p:spTree/p:sp[20]/p:txBody/a:p/a:r/a:t', namespaces={'a':'http://schemas.openxmlformats.org/drawingml/2006/main',
            'r':'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
            'p':'http://schemas.openxmlformats.org/presentationml/2006/main'})
    E6[0].text = str(rows[46][1]).replace(' ', '')
    E10 = tree.xpath('/p:sld/p:cSld/p:spTree/p:sp[21]/p:txBody/a:p/a:r/a:t', namespaces={'a':'http://schemas.openxmlformats.org/drawingml/2006/main',
            'r':'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
            'p':'http://schemas.openxmlformats.org/presentationml/2006/main'})
    E10[0].text = str(rows[51][1]).replace(' ', '')
    D2 = tree.xpath('/p:sld/p:cSld/p:spTree/p:sp[9]/p:txBody/a:p/a:r[2]/a:t', namespaces={'a':'http://schemas.openxmlformats.org/drawingml/2006/main',
            'r':'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
            'p':'http://schemas.openxmlformats.org/presentationml/2006/main'})
    D2[0].text = (str(rows[35][1]).replace(' ', ''))[1:]
    D6 = tree.xpath('/p:sld/p:cSld/p:spTree/p:sp[18]/p:txBody/a:p/a:r[2]/a:t', namespaces={'a':'http://schemas.openxmlformats.org/drawingml/2006/main',
            'r':'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
            'p':'http://schemas.openxmlformats.org/presentationml/2006/main'})
    D6[0].text = (str(rows[45][1]).replace(' ', ''))[1:]
    D10 = tree.xpath('/p:sld/p:cSld/p:spTree/p:sp[19]/p:txBody/a:p/a:r/a:t', namespaces={'a':'http://schemas.openxmlformats.org/drawingml/2006/main',
            'r':'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
            'p':'http://schemas.openxmlformats.org/presentationml/2006/main'})
    D10[0].text = str(rows[50][1]).replace(' ', '')
    C2 = tree.xpath('/p:sld/p:cSld/p:spTree/p:sp[8]/p:txBody/a:p/a:r/a:t', namespaces={'a':'http://schemas.openxmlformats.org/drawingml/2006/main',
            'r':'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
            'p':'http://schemas.openxmlformats.org/presentationml/2006/main'})
    C2[0].text = str(rows[34][1]).replace(' ', '')
    C6 = tree.xpath('/p:sld/p:cSld/p:spTree/p:sp[16]/p:txBody/a:p/a:r/a:t', namespaces={'a':'http://schemas.openxmlformats.org/drawingml/2006/main',
            'r':'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
            'p':'http://schemas.openxmlformats.org/presentationml/2006/main'})
    C6[0].text = str(rows[44][1]).replace(' ', '')
    C10 = tree.xpath('/p:sld/p:cSld/p:spTree/p:sp[17]/p:txBody/a:p/a:r/a:t', namespaces={'a':'http://schemas.openxmlformats.org/drawingml/2006/main',
            'r':'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
            'p':'http://schemas.openxmlformats.org/presentationml/2006/main'})
    C10[0].text = str(rows[49][1]).replace(' ', '')
    B2 = tree.xpath('/p:sld/p:cSld/p:spTree/p:sp[7]/p:txBody/a:p/a:r[2]/a:t', namespaces={'a':'http://schemas.openxmlformats.org/drawingml/2006/main',
            'r':'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
            'p':'http://schemas.openxmlformats.org/presentationml/2006/main'})
    B2[0].text = (str(rows[33][1]).replace(' ', ''))[1:]
    B6 = tree.xpath('/p:sld/p:cSld/p:spTree/p:sp[12]/p:txBody/a:p/a:r/a:t', namespaces={'a':'http://schemas.openxmlformats.org/drawingml/2006/main',
            'r':'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
            'p':'http://schemas.openxmlformats.org/presentationml/2006/main'})
    B6[0].text = str(rows[43][1]).replace(' ', '')
    B10 = tree.xpath('/p:sld/p:cSld/p:spTree/p:sp[14]/p:txBody/a:p/a:r[1]/a:t', namespaces={'a':'http://schemas.openxmlformats.org/drawingml/2006/main',
            'r':'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
            'p':'http://schemas.openxmlformats.org/presentationml/2006/main'})
    B10[0].text = (str(rows[48][1]).replace(' ', ''))[:-1]
    #-----------------------------------------------------------------------------------------------------------------------------------------------------------------
    leftRectangleHeight = 1528955#1766996
    leftRectangleStep = 3057.91
    leftTextBoxHeight = 2247457
    upperButtomLine = 4867490       #Straight Connector 45
    lowerButtomLine = 7873931
    #Rectangle 42
    F12 = tree.xpath('/p:sld/p:cSld/p:spTree/p:sp[27]/p:txBody/a:p/a:r[2]/a:t', namespaces={'a':'http://schemas.openxmlformats.org/drawingml/2006/main',
            'r':'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
            'p':'http://schemas.openxmlformats.org/presentationml/2006/main'})
    F12[0].text = (str(rows[19][1]).replace(' ', ''))[1:] #TextBox 46   float((str(rows[19][1]).replace(' ', ''))[1:])
    TextBox46Y = tree.xpath('/p:sld/p:cSld/p:spTree/p:sp[27]/p:spPr/a:xfrm/a:off', namespaces={'a':'http://schemas.openxmlformats.org/drawingml/2006/main',
            'r':'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
            'p':'http://schemas.openxmlformats.org/presentationml/2006/main'})
    TextBox46Y[0].attrib['y'] = unicode(upperButtomLine-round(leftRectangleStep*float((str(rows[19][1]).replace(' ', ''))[1:])) - 500000)
    Rectangle42Y = tree.xpath('/p:sld/p:cSld/p:spTree/p:sp[25]/p:spPr/a:xfrm/a:off', namespaces={'a':'http://schemas.openxmlformats.org/drawingml/2006/main',
            'r':'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
            'p':'http://schemas.openxmlformats.org/presentationml/2006/main'})
    Rectangle42CY = tree.xpath('/p:sld/p:cSld/p:spTree/p:sp[25]/p:spPr/a:xfrm/a:ext', namespaces={'a':'http://schemas.openxmlformats.org/drawingml/2006/main',
            'r':'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
            'p':'http://schemas.openxmlformats.org/presentationml/2006/main'})
    Rectangle42Y[0].attrib['y'] = unicode(upperButtomLine-leftRectangleHeight+round(leftRectangleStep*(500 - float((str(rows[19][1]).replace(' ', ''))[1:])))) #500 - number
    Rectangle42CY[0].attrib['cy'] = unicode(leftRectangleHeight-round(leftRectangleStep*(500 - float((str(rows[19][1]).replace(' ', ''))[1:])))) #500 - number

    #Rectangle 43
    F11 = tree.xpath('/p:sld/p:cSld/p:spTree/p:sp[28]/p:txBody/a:p/a:r[2]/a:t', namespaces={'a':'http://schemas.openxmlformats.org/drawingml/2006/main',
            'r':'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
            'p':'http://schemas.openxmlformats.org/presentationml/2006/main'})
    F11[0].text = (str(rows[21][1][1:]).replace(' ', ''))[1:] #TextBox 47  float((str(rows[21][1][1:]).replace(' ', ''))[1:])
    TextBox47X = tree.xpath('/p:sld/p:cSld/p:spTree/p:sp[28]/p:spPr/a:xfrm/a:off', namespaces={'a':'http://schemas.openxmlformats.org/drawingml/2006/main',
            'r':'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
            'p':'http://schemas.openxmlformats.org/presentationml/2006/main'})
    TextBox47X[0].attrib['y'] = unicode(upperButtomLine-round(leftRectangleStep*float((str(rows[21][1][1:]).replace(' ', ''))[1:])) - 500000)
    Rectangle43Y = tree.xpath('/p:sld/p:cSld/p:spTree/p:sp[26]/p:spPr/a:xfrm/a:off', namespaces={'a':'http://schemas.openxmlformats.org/drawingml/2006/main',
            'r':'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
            'p':'http://schemas.openxmlformats.org/presentationml/2006/main'})
    Rectangle43CY = tree.xpath('/p:sld/p:cSld/p:spTree/p:sp[26]/p:spPr/a:xfrm/a:ext', namespaces={'a':'http://schemas.openxmlformats.org/drawingml/2006/main',
            'r':'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
            'p':'http://schemas.openxmlformats.org/presentationml/2006/main'})
    Rectangle43Y[0].attrib['y'] = unicode(upperButtomLine-leftRectangleHeight+round(leftRectangleStep*(500 - float((str(rows[21][1][1:]).replace(' ', ''))[1:])))) #500 - number
    Rectangle43CY[0].attrib['cy'] = unicode(leftRectangleHeight-round(leftRectangleStep*(500 - float((str(rows[21][1][1:]).replace(' ', ''))[1:])))) #500 - number

    #Rectangle 50
    F14 = tree.xpath('/p:sld/p:cSld/p:spTree/p:sp[31]/p:txBody/a:p/a:r/a:t', namespaces={'a':'http://schemas.openxmlformats.org/drawingml/2006/main',
            'r':'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
            'p':'http://schemas.openxmlformats.org/presentationml/2006/main'})
    F14[0].text = rows[23][1] #TextBox 53 float((str(rows[23][1]).replace(' ', ''))[1:])
    TextBox53X = tree.xpath('/p:sld/p:cSld/p:spTree/p:sp[31]/p:spPr/a:xfrm/a:off', namespaces={'a':'http://schemas.openxmlformats.org/drawingml/2006/main',
            'r':'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
            'p':'http://schemas.openxmlformats.org/presentationml/2006/main'})
    TextBox53X[0].attrib['y'] = unicode(lowerButtomLine-round(leftRectangleStep*float((str(rows[23][1]).replace(' ', ''))[1:])) - 500000)
    Rectangle50Y = tree.xpath('/p:sld/p:cSld/p:spTree/p:sp[29]/p:spPr/a:xfrm/a:off', namespaces={'a':'http://schemas.openxmlformats.org/drawingml/2006/main',
            'r':'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
            'p':'http://schemas.openxmlformats.org/presentationml/2006/main'})
    Rectangle50CY = tree.xpath('/p:sld/p:cSld/p:spTree/p:sp[29]/p:spPr/a:xfrm/a:ext', namespaces={'a':'http://schemas.openxmlformats.org/drawingml/2006/main',
            'r':'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
            'p':'http://schemas.openxmlformats.org/presentationml/2006/main'})
    Rectangle50Y[0].attrib['y'] = unicode(lowerButtomLine-leftRectangleHeight+round(leftRectangleStep*(500 - float((str(rows[23][1]).replace(' ', ''))[1:])))) #500 - number
    Rectangle50CY[0].attrib['cy'] = unicode(leftRectangleHeight-round(leftRectangleStep*(500 - float((str(rows[23][1]).replace(' ', ''))[1:])))) #500 - number

    #Rectangle 51
    F13 = tree.xpath('/p:sld/p:cSld/p:spTree/p:sp[32]/p:txBody/a:p/a:r/a:t', namespaces={'a':'http://schemas.openxmlformats.org/drawingml/2006/main',
            'r':'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
            'p':'http://schemas.openxmlformats.org/presentationml/2006/main'})
    F13[0].text = rows[25][1] #TextBox 54  float((str(rows[24][1]).replace(' ', '')))
    TextBox54X = tree.xpath('/p:sld/p:cSld/p:spTree/p:sp[32]/p:spPr/a:xfrm/a:off', namespaces={'a':'http://schemas.openxmlformats.org/drawingml/2006/main',
            'r':'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
            'p':'http://schemas.openxmlformats.org/presentationml/2006/main'})
    ee = (str(rows[25][1]).replace(' ', ''))[1:]
    ee = float(ee)
    TextBox54X[0].attrib['y'] = unicode(lowerButtomLine-round(leftRectangleStep*ee) - 1000000)
    Rectangle51Y = tree.xpath('/p:sld/p:cSld/p:spTree/p:sp[30]/p:spPr/a:xfrm/a:off', namespaces={'a':'http://schemas.openxmlformats.org/drawingml/2006/main',
            'r':'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
            'p':'http://schemas.openxmlformats.org/presentationml/2006/main'})
    Rectangle51CY = tree.xpath('/p:sld/p:cSld/p:spTree/p:sp[30]/p:spPr/a:xfrm/a:ext', namespaces={'a':'http://schemas.openxmlformats.org/drawingml/2006/main',
            'r':'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
            'p':'http://schemas.openxmlformats.org/presentationml/2006/main'})
    Rectangle51Y[0].attrib['y'] = unicode(lowerButtomLine-leftRectangleHeight+round(leftRectangleStep*(500 - ee))) #500 - number
    Rectangle51CY[0].attrib['cy'] = unicode(leftRectangleHeight-round(leftRectangleStep*(500 - ee))) #500 - number


    #-----------------------------------------------------------------------------------------------------------------------------------------------------------------
    #Rectangle 63
    var013 = float(str(rows[29][1]).replace(' ', ''))
    var018 = float(str(rows[30][1]).replace(' ', ''))
    var048 = float(str(rows[31][1]).replace(' ', ''))
    rightButtom = 7854172
    rightHeight = 5589389
    rightWidth = 10237409
    rightLeftCorner = 3654862
    iconOffset = 220113
    Rectangle63YX = tree.xpath('/p:sld/p:cSld/p:spTree/p:sp[1]/p:spPr/a:xfrm/a:off', namespaces={'a':'http://schemas.openxmlformats.org/drawingml/2006/main',
            'r':'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
            'p':'http://schemas.openxmlformats.org/presentationml/2006/main'})
    Rectangle63CYCX = tree.xpath('/p:sld/p:cSld/p:spTree/p:sp[1]/p:spPr/a:xfrm/a:ext', namespaces={'a':'http://schemas.openxmlformats.org/drawingml/2006/main',
            'r':'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
            'p':'http://schemas.openxmlformats.org/presentationml/2006/main'})
    Rectangle63YX[0].attrib['y'] = unicode(rightButtom-round(rightHeight*float(str(rows[53][1])))) #0.3 is variable
    Rectangle63CYCX[0].attrib['cy'] = unicode(round(rightHeight*round(rightHeight*float(str(rows[53][1]))))) #0.3 is variable
    Rectangle63CYCX[0].attrib['cx'] = unicode(round(rightWidth*var013)+rightLeftCorner) #0.3 is variable
    #Straight Connector 65
    StraightConnector65X = tree.xpath('/p:sld/p:cSld/p:spTree/p:cxnSp[2]/p:spPr/a:xfrm/a:off', namespaces={'a':'http://schemas.openxmlformats.org/drawingml/2006/main',
            'r':'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
            'p':'http://schemas.openxmlformats.org/presentationml/2006/main'})
    StraightConnector65X[0].attrib['x'] = unicode(round(rightWidth*var013)+rightLeftCorner)
    #Group 5
    Group5X = tree.xpath('/p:sld/p:cSld/p:spTree/p:grpSp[2]/p:grpSpPr/a:xfrm/a:off', namespaces={'a':'http://schemas.openxmlformats.org/drawingml/2006/main',
            'r':'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
            'p':'http://schemas.openxmlformats.org/presentationml/2006/main'})
    Group5X[0].attrib['x'] = unicode(round(rightWidth*var013)+rightLeftCorner+iconOffset)
    #TextBox 23
    TextBox23X = tree.xpath('/p:sld/p:cSld/p:spTree/p:sp[16]/p:spPr/a:xfrm/a:off', namespaces={'a':'http://schemas.openxmlformats.org/drawingml/2006/main',
            'r':'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
            'p':'http://schemas.openxmlformats.org/presentationml/2006/main'})
    TextBox23X[0].attrib['x'] = unicode(round(rightWidth*var013)+rightLeftCorner+99634)
    #TextBox 24
    TextBox24X = tree.xpath('/p:sld/p:cSld/p:spTree/p:sp[17]/p:spPr/a:xfrm/a:off', namespaces={'a':'http://schemas.openxmlformats.org/drawingml/2006/main',
            'r':'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
            'p':'http://schemas.openxmlformats.org/presentationml/2006/main'})
    TextBox24X[0].attrib['x'] = unicode(round(rightWidth*var013)+rightLeftCorner+99634)
    #TextBox 15
    TextBox15X = tree.xpath('/p:sld/p:cSld/p:spTree/p:sp[8]/p:spPr/a:xfrm/a:off', namespaces={'a':'http://schemas.openxmlformats.org/drawingml/2006/main',
            'r':'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
            'p':'http://schemas.openxmlformats.org/presentationml/2006/main'})
    TextBox15X[0].attrib['x'] = unicode(round(rightWidth*var013)+rightLeftCorner+99634)


    #Rectangle 68
    Rectangle68YX = tree.xpath('/p:sld/p:cSld/p:spTree/p:sp[2]/p:spPr/a:xfrm/a:off', namespaces={'a':'http://schemas.openxmlformats.org/drawingml/2006/main',
            'r':'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
            'p':'http://schemas.openxmlformats.org/presentationml/2006/main'})
    Rectangle68CYCX = tree.xpath('/p:sld/p:cSld/p:spTree/p:sp[2]/p:spPr/a:xfrm/a:ext', namespaces={'a':'http://schemas.openxmlformats.org/drawingml/2006/main',
            'r':'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
            'p':'http://schemas.openxmlformats.org/presentationml/2006/main'})
    Rectangle68YX[0].attrib['y'] = unicode(rightButtom-round(rightHeight*float(str(rows[54][1]))))
    Rectangle68CYCX[0].attrib['cy'] = unicode(round(rightHeight*float(str(rows[54][1]))))
    #Straight Connector 66
    StraightConnector66X = tree.xpath('/p:sld/p:cSld/p:spTree/p:cxnSp[3]/p:spPr/a:xfrm/a:off', namespaces={'a':'http://schemas.openxmlformats.org/drawingml/2006/main',
            'r':'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
            'p':'http://schemas.openxmlformats.org/presentationml/2006/main'})
    StraightConnector66X[0].attrib['x'] = unicode(round(round(rightWidth*var013)+rightLeftCorner + rightWidth*var018))
    #Group 6
    Group6X = tree.xpath('/p:sld/p:cSld/p:spTree/p:grpSp[3]/p:grpSpPr/a:xfrm/a:off', namespaces={'a':'http://schemas.openxmlformats.org/drawingml/2006/main',
            'r':'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
            'p':'http://schemas.openxmlformats.org/presentationml/2006/main'})
    Group6X[0].attrib['x'] = unicode(round(round(rightWidth*var013)+rightLeftCorner + rightWidth*var018)+iconOffset)
    #TextBox 25
    TextBox25X = tree.xpath('/p:sld/p:cSld/p:spTree/p:sp[18]/p:spPr/a:xfrm/a:off', namespaces={'a':'http://schemas.openxmlformats.org/drawingml/2006/main',
            'r':'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
            'p':'http://schemas.openxmlformats.org/presentationml/2006/main'})
    TextBox25X[0].attrib['x'] = unicode(round(round(rightWidth*var013)+rightLeftCorner + rightWidth*var018)+99634)
    #TextBox 26
    TextBox26X = tree.xpath('/p:sld/p:cSld/p:spTree/p:sp[19]/p:spPr/a:xfrm/a:off', namespaces={'a':'http://schemas.openxmlformats.org/drawingml/2006/main',
            'r':'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
            'p':'http://schemas.openxmlformats.org/presentationml/2006/main'})
    TextBox26X[0].attrib['x'] = unicode(round(round(rightWidth*var013)+rightLeftCorner + rightWidth*var018)+99634)
    #TextBox 16
    TextBox16X = tree.xpath('/p:sld/p:cSld/p:spTree/p:sp[9]/p:spPr/a:xfrm/a:off', namespaces={'a':'http://schemas.openxmlformats.org/drawingml/2006/main',
            'r':'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
            'p':'http://schemas.openxmlformats.org/presentationml/2006/main'})
    TextBox16X[0].attrib['x'] = unicode(round(round(rightWidth*var013)+rightLeftCorner + rightWidth*var018)+99634)

    #Rectangle 69
    Rectangle69YX = tree.xpath('/p:sld/p:cSld/p:spTree/p:sp[3]/p:spPr/a:xfrm/a:off', namespaces={'a':'http://schemas.openxmlformats.org/drawingml/2006/main',
            'r':'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
            'p':'http://schemas.openxmlformats.org/presentationml/2006/main'})
    Rectangle69CYCX = tree.xpath('/p:sld/p:cSld/p:spTree/p:sp[3]/p:spPr/a:xfrm/a:ext', namespaces={'a':'http://schemas.openxmlformats.org/drawingml/2006/main',
            'r':'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
            'p':'http://schemas.openxmlformats.org/presentationml/2006/main'})
    Rectangle69YX[0].attrib['y'] = unicode(rightButtom-round(rightHeight*float(str(rows[55][1]))))
    Rectangle69CYCX[0].attrib['cy'] = unicode(round(rightHeight*float(str(rows[55][1]))))
    #Straight Connector 67
    StraightConnector67X = tree.xpath('/p:sld/p:cSld/p:spTree/p:cxnSp[4]/p:spPr/a:xfrm/a:off', namespaces={'a':'http://schemas.openxmlformats.org/drawingml/2006/main',
            'r':'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
            'p':'http://schemas.openxmlformats.org/presentationml/2006/main'})
    StraightConnector67X[0].attrib['x'] = unicode(round(round(round(rightWidth*var013)+rightLeftCorner + rightWidth*var018) + rightWidth*var048))
    #Group 7
    Group7X = tree.xpath('/p:sld/p:cSld/p:spTree/p:grpSp[4]/p:grpSpPr/a:xfrm/a:off', namespaces={'a':'http://schemas.openxmlformats.org/drawingml/2006/main',
            'r':'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
            'p':'http://schemas.openxmlformats.org/presentationml/2006/main'})
    Group7X[0].attrib['x'] = unicode(round(round(round(rightWidth*var013)+rightLeftCorner + rightWidth*var018) + rightWidth*var048)+iconOffset)
    #TextBox 27
    TextBox27X = tree.xpath('/p:sld/p:cSld/p:spTree/p:sp[20]/p:spPr/a:xfrm/a:off', namespaces={'a':'http://schemas.openxmlformats.org/drawingml/2006/main',
            'r':'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
            'p':'http://schemas.openxmlformats.org/presentationml/2006/main'})
    TextBox27X[0].attrib['x'] = unicode(round(round(round(rightWidth*var013)+rightLeftCorner + rightWidth*var018) + rightWidth*var048)+99634)
    #TextBox 28
    TextBox28X = tree.xpath('/p:sld/p:cSld/p:spTree/p:sp[21]/p:spPr/a:xfrm/a:off', namespaces={'a':'http://schemas.openxmlformats.org/drawingml/2006/main',
            'r':'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
            'p':'http://schemas.openxmlformats.org/presentationml/2006/main'})
    TextBox28X[0].attrib['x'] = unicode(round(round(round(rightWidth*var013)+rightLeftCorner + rightWidth*var018) + rightWidth*var048)+99634)
    #TextBox 17
    TextBox17X = tree.xpath('/p:sld/p:cSld/p:spTree/p:sp[10]/p:spPr/a:xfrm/a:off', namespaces={'a':'http://schemas.openxmlformats.org/drawingml/2006/main',
            'r':'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
            'p':'http://schemas.openxmlformats.org/presentationml/2006/main'})
    TextBox17X[0].attrib['x'] = unicode(round(round(round(rightWidth*var013)+rightLeftCorner + rightWidth*var018) + rightWidth*var048)+99634)


    #Rectangle 70
    Rectangle70YX = tree.xpath('/p:sld/p:cSld/p:spTree/p:sp[4]/p:spPr/a:xfrm/a:off', namespaces={'a':'http://schemas.openxmlformats.org/drawingml/2006/main',
            'r':'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
            'p':'http://schemas.openxmlformats.org/presentationml/2006/main'})
    Rectangle70CYCX = tree.xpath('/p:sld/p:cSld/p:spTree/p:sp[4]/p:spPr/a:xfrm/a:ext', namespaces={'a':'http://schemas.openxmlformats.org/drawingml/2006/main',
            'r':'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
            'p':'http://schemas.openxmlformats.org/presentationml/2006/main'})
    Rectangle70YX[0].attrib['y'] = unicode(rightButtom-round(rightHeight*float(str(rows[56][1]))))
    Rectangle70CYCX[0].attrib['cy'] = unicode(round(rightHeight*float(str(rows[56][1]))))


    path = os.path.join(os.path.split(os.path.abspath(__file__))[0], 'ppt/slides/slide1.xml')
    f = open(path, 'w')
    f.write(etree.tostring(tree, pretty_print=False))
    f.close()


def openFile2(rows):
    x0 = 7034856
    xMax = 13435656
    xMin = 2767656
    path = os.path.join(os.path.split(os.path.abspath(__file__))[0], 'ppt/slides/slide2.xml')
    tree = etree.parse(path)

    #Oval 60
    shift60 = rows[21][1]
    if(shift60[0]) == '-':
        shiftNum60 = float(shift60[1:-1])
        shiftNum60 = shiftNum60/10
        xNew60 = x0 - shiftNum60*(x0 - xMin)
    else:
        shiftNum60 = float(shift60[:-1])
        xNew60 = x0 + 0.06666666*shiftNum60*(xMax-x0)
    Oval60X = tree.xpath('/p:sld/p:cSld/p:spTree/p:sp[13]/p:spPr/a:xfrm/a:off', namespaces={'a':'http://schemas.openxmlformats.org/drawingml/2006/main',
            'r':'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
            'p':'http://schemas.openxmlformats.org/presentationml/2006/main'})
    Oval60X[0].attrib['x'] = unicode(round(xNew60))[:-2]
    #Group 2
    if(rows[22][1][0]) == '-':
        shiftNum60 = float(rows[22][1][1:-1])
        shiftNum60 = shiftNum60/10
        xNewGr2 = x0 - shiftNum60*(x0 - xMin)
    else:
        shiftNum60 = float(rows[22][1][:-1])
        xNewGr2 = x0 + 0.06666666*shiftNum60*(xMax-x0)
    Group2X = tree.xpath('/p:sld/p:cSld/p:spTree/p:grpSp[1]/p:grpSpPr/a:xfrm/a:off', namespaces={'a':'http://schemas.openxmlformats.org/drawingml/2006/main',
            'r':'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
            'p':'http://schemas.openxmlformats.org/presentationml/2006/main'})
    Group2X[0].attrib['x'] = unicode(round(xNewGr2))[:-2]


    #Oval 62
    shift62 = rows[25][1]
    if(shift62[0]) == '-':
        shiftNum62 = float(shift62[1:-1])
        shiftNum62 = shiftNum62/10
        xNew62 = x0 - shiftNum62*(x0 - xMin)
    else:
        shiftNum62 = float(shift62[:-1])
        xNew62 = x0 + 0.06666666*shiftNum62*(xMax-x0)
    Oval62X = tree.xpath('/p:sld/p:cSld/p:spTree/p:sp[15]/p:spPr/a:xfrm/a:off', namespaces={'a':'http://schemas.openxmlformats.org/drawingml/2006/main',
            'r':'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
            'p':'http://schemas.openxmlformats.org/presentationml/2006/main'})
    Oval62X[0].attrib['x'] = unicode(round(xNew62))[:-2]
    #Group 7
    if(rows[26][1][0]) == '-':
        shiftNum60 = float(rows[26][1][1:-1])
        shiftNum60 = shiftNum60/10
        xNewGr7 = x0 - shiftNum60*(x0 - xMin)
    else:
        shiftNum60 = float(rows[26][1][:-1])
        xNewGr7 = x0 + 0.06666666*shiftNum60*(xMax-x0)
    Group7X = tree.xpath('/p:sld/p:cSld/p:spTree/p:grpSp[4]/p:grpSpPr/a:xfrm/a:off', namespaces={'a':'http://schemas.openxmlformats.org/drawingml/2006/main',
            'r':'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
            'p':'http://schemas.openxmlformats.org/presentationml/2006/main'})
    Group7X[0].attrib['x'] = unicode(round(xNewGr7))[:-2]


    #Oval 61
    shift61 = rows[23][1]
    if(shift61[0]) == '-':
        shiftNum61 = float(shift61[1:-1])
        shiftNum61 = shiftNum61/10
        xNew61 = x0 - shiftNum61*(x0 - xMin)
    else:
        shiftNum61 = float(shift61[:-1])
        xNew61 = x0 + 0.06666666*shiftNum61*(xMax-x0)
    Oval61X = tree.xpath('/p:sld/p:cSld/p:spTree/p:sp[14]/p:spPr/a:xfrm/a:off', namespaces={'a':'http://schemas.openxmlformats.org/drawingml/2006/main',
            'r':'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
            'p':'http://schemas.openxmlformats.org/presentationml/2006/main'})
    Oval61X[0].attrib['x'] = unicode(round(xNew61))[:-2]
    #Group 5
    if(rows[24][1][0]) == '-':
        shiftNum60 = float(rows[24][1][1:-1])
        shiftNum60 = shiftNum60/10
        xNewGr5 = x0 - shiftNum60*(x0 - xMin)
    else:
        shiftNum60 = float(rows[24][1][:-1])
        xNewGr5 = x0 + 0.06666666*shiftNum60*(xMax-x0)
    Group5X = tree.xpath('/p:sld/p:cSld/p:spTree/p:grpSp[2]/p:grpSpPr/a:xfrm/a:off', namespaces={'a':'http://schemas.openxmlformats.org/drawingml/2006/main',
            'r':'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
            'p':'http://schemas.openxmlformats.org/presentationml/2006/main'})
    Group5X[0].attrib['x'] = unicode(round(xNewGr5))[:-2]



    #Oval 63
    shift63 = rows[27][1]
    if(shift63[0]) == '-':
        shiftNum63 = float(shift63[1:-1])
        shiftNum63 = shiftNum63/10
        xNew63 = x0 - shiftNum63*(x0 - xMin)
    else:
        shiftNum63 = float(shift63[:-1])
        xNew63 = x0 + 0.06666666*shiftNum63*(xMax-x0)
    Oval63X = tree.xpath('/p:sld/p:cSld/p:spTree/p:sp[16]/p:spPr/a:xfrm/a:off', namespaces={'a':'http://schemas.openxmlformats.org/drawingml/2006/main',
            'r':'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
            'p':'http://schemas.openxmlformats.org/presentationml/2006/main'})
    Oval63X[0].attrib['x'] = unicode(round(xNew63))[:-2]
    #Group 27
    if(rows[28][1][0]) == '-':
        shiftNum60 = float(rows[28][1][1:-1])
        shiftNum60 = shiftNum60/10
        xNewGr27 = x0 - shiftNum60*(x0 - xMin)
    else:
        shiftNum60 = float(rows[28][1][:-1])
        xNewGr27 = x0 + 0.06666666*shiftNum60*(xMax-x0)
    Group27X = tree.xpath('/p:sld/p:cSld/p:spTree/p:grpSp[3]/p:grpSpPr/a:xfrm/a:off', namespaces={'a':'http://schemas.openxmlformats.org/drawingml/2006/main',
            'r':'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
            'p':'http://schemas.openxmlformats.org/presentationml/2006/main'})
    Group27X[0].attrib['x'] = unicode(round(xNewGr27))[:-2]


    #Oval 64
    shift64 = rows[29][1]
    if(shift64[0]) == '-':
        shiftNum64 = float(shift64[1:-1])
        shiftNum64 = shiftNum64/10
        xNew64 = x0 - shiftNum64*(x0 - xMin)
    else:
        shiftNum64 = float(shift64[:-1])
        xNew64 = x0 + 0.06666666*shiftNum64*(xMax-x0)
    Oval64X = tree.xpath('/p:sld/p:cSld/p:spTree/p:sp[17]/p:spPr/a:xfrm/a:off', namespaces={'a':'http://schemas.openxmlformats.org/drawingml/2006/main',
            'r':'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
            'p':'http://schemas.openxmlformats.org/presentationml/2006/main'})
    Oval64X[0].attrib['x'] = unicode(round(xNew64))[:-2]
    #Group 29
    if(rows[30][1][0]) == '-':
        shiftNum60 = float(rows[30][1][1:-1])
        shiftNum60 = shiftNum60/10
        xNewGr29 = x0 - shiftNum60*(x0 - xMin)
    else:
        shiftNum60 = float(rows[30][1][:-1])
        xNewGr29 = x0 + 0.06666666*shiftNum60*(xMax-x0)
    Group29X = tree.xpath('/p:sld/p:cSld/p:spTree/p:grpSp[5]/p:grpSpPr/a:xfrm/a:off', namespaces={'a':'http://schemas.openxmlformats.org/drawingml/2006/main',
            'r':'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
            'p':'http://schemas.openxmlformats.org/presentationml/2006/main'})
    Group29X[0].attrib['x'] = unicode(round(xNewGr29))[:-2]


    path = os.path.join(os.path.split(os.path.abspath(__file__))[0], 'ppt/slides/slide2.xml')
    f = open(path, 'w')
    f.write(etree.tostring(tree, pretty_print=False))
    f.close()


def openFile3(rows):
    path = os.path.join(os.path.split(os.path.abspath(__file__))[0], 'ppt/slides/slide3.xml')
    tree = etree.parse(path)
    Row1Curr = tree.xpath('/p:sld/p:cSld/p:spTree/p:sp[10]/p:txBody/a:p/a:r/a:t', namespaces={'a':'http://schemas.openxmlformats.org/drawingml/2006/main',
            'r':'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
            'p':'http://schemas.openxmlformats.org/presentationml/2006/main'})
    Row1Curr[0].text = str(rows[4][1]).replace(" ", "")
    Row1Chang = tree.xpath('/p:sld/p:cSld/p:spTree/p:sp[19]/p:txBody/a:p/a:r/a:t', namespaces={'a':'http://schemas.openxmlformats.org/drawingml/2006/main',
            'r':'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
            'p':'http://schemas.openxmlformats.org/presentationml/2006/main'})
    Row1Chang[0].text = str(rows[9][1]).replace(" ", "")
    #Picture 16 Picture 45
    growth1 = float(unicode(str(rows[9][1]).replace(" ", ""))[:-1])
    if(growth1 < 0):
        Picture16CYCX = tree.xpath('/p:sld/p:cSld/p:spTree/p:pic[4]/p:spPr/a:xfrm/a:ext', namespaces={'a':'http://schemas.openxmlformats.org/drawingml/2006/main',
            'r':'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
            'p':'http://schemas.openxmlformats.org/presentationml/2006/main'})
        Picture16YX = tree.xpath('/p:sld/p:cSld/p:spTree/p:pic[4]/p:spPr/a:xfrm/a:off', namespaces={'a':'http://schemas.openxmlformats.org/drawingml/2006/main',
            'r':'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
            'p':'http://schemas.openxmlformats.org/presentationml/2006/main'})
        Picture16CYCX[0].attrib['cx'] = unicode('0')
        Picture16CYCX[0].attrib['cy'] = unicode('0')
        Picture16YX[0].attrib['x'] = unicode('0')
        Picture16YX[0].attrib['y'] = unicode('0')
    else:
        Picture45CYCX = tree.xpath('/p:sld/p:cSld/p:spTree/p:pic[1]/p:spPr/a:xfrm/a:ext', namespaces={'a':'http://schemas.openxmlformats.org/drawingml/2006/main',
            'r':'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
            'p':'http://schemas.openxmlformats.org/presentationml/2006/main'})
        Picture45YX = tree.xpath('/p:sld/p:cSld/p:spTree/p:pic[1]/p:spPr/a:xfrm/a:off', namespaces={'a':'http://schemas.openxmlformats.org/drawingml/2006/main',
            'r':'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
            'p':'http://schemas.openxmlformats.org/presentationml/2006/main'})
        Picture45CYCX[0].attrib['cx'] = unicode('0')
        Picture45CYCX[0].attrib['cy'] = unicode('0')
        Picture45YX[0].attrib['x'] = unicode('0')
        Picture45YX[0].attrib['y'] = unicode('0')


    Row2Curr = tree.xpath('/p:sld/p:cSld/p:spTree/p:sp[11]/p:txBody/a:p/a:r/a:t', namespaces={'a':'http://schemas.openxmlformats.org/drawingml/2006/main',
            'r':'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
            'p':'http://schemas.openxmlformats.org/presentationml/2006/main'})
    Row2Curr[0].text = str(rows[5][1]).replace(" ", "")
    Row2Chang = tree.xpath('/p:sld/p:cSld/p:spTree/p:sp[22]/p:txBody/a:p/a:r/a:t', namespaces={'a':'http://schemas.openxmlformats.org/drawingml/2006/main',
            'r':'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
            'p':'http://schemas.openxmlformats.org/presentationml/2006/main'})
    Row2Chang[0].text = str(rows[10][1]).replace(" ", "")
    #Picture 58 Picture 59
    growth2 = float(unicode(str(rows[10][1])).replace(" ", "")[:-1])
    if(growth2 < 0):
        Picture58CYCX = tree.xpath('/p:sld/p:cSld/p:spTree/p:pic[7]/p:spPr/a:xfrm/a:ext', namespaces={'a':'http://schemas.openxmlformats.org/drawingml/2006/main',
            'r':'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
            'p':'http://schemas.openxmlformats.org/presentationml/2006/main'})
        Picture58YX = tree.xpath('/p:sld/p:cSld/p:spTree/p:pic[7]/p:spPr/a:xfrm/a:off', namespaces={'a':'http://schemas.openxmlformats.org/drawingml/2006/main',
            'r':'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
            'p':'http://schemas.openxmlformats.org/presentationml/2006/main'})
        Picture58CYCX[0].attrib['cx'] = unicode('0')
        Picture58CYCX[0].attrib['cy'] = unicode('0')
        Picture58YX[0].attrib['x'] = unicode('0')
        Picture58YX[0].attrib['y'] = unicode('0')
    else:
        Picture59CYCX = tree.xpath('/p:sld/p:cSld/p:spTree/p:pic[2]/p:spPr/a:xfrm/a:ext', namespaces={'a':'http://schemas.openxmlformats.org/drawingml/2006/main',
            'r':'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
            'p':'http://schemas.openxmlformats.org/presentationml/2006/main'})
        Picture59YX = tree.xpath('/p:sld/p:cSld/p:spTree/p:pic[2]/p:spPr/a:xfrm/a:off', namespaces={'a':'http://schemas.openxmlformats.org/drawingml/2006/main',
            'r':'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
            'p':'http://schemas.openxmlformats.org/presentationml/2006/main'})
        Picture59CYCX[0].attrib['cx'] = unicode('0')
        Picture59CYCX[0].attrib['cy'] = unicode('0')
        Picture59YX[0].attrib['x'] = unicode('0')
        Picture59YX[0].attrib['y'] = unicode('0')


    Row3Curr = tree.xpath('/p:sld/p:cSld/p:spTree/p:sp[12]/p:txBody/a:p/a:r/a:t', namespaces={'a':'http://schemas.openxmlformats.org/drawingml/2006/main',
            'r':'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
            'p':'http://schemas.openxmlformats.org/presentationml/2006/main'})
    Row3Curr[0].text = str(rows[6][1]).replace(" ", "")
    Row3Chang = tree.xpath('/p:sld/p:cSld/p:spTree/p:sp[20]/p:txBody/a:p/a:r/a:t', namespaces={'a':'http://schemas.openxmlformats.org/drawingml/2006/main',
            'r':'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
            'p':'http://schemas.openxmlformats.org/presentationml/2006/main'})
    Row3Chang[0].text = rows[11][1]
    #Picture 18 Picture 60
    growth3 = float(unicode(str(rows[11][1]).replace(" ", ""))[:-1])
    if(growth3 < 0):
        Picture18CYCX = tree.xpath('/p:sld/p:cSld/p:spTree/p:pic[5]/p:spPr/a:xfrm/a:ext', namespaces={'a':'http://schemas.openxmlformats.org/drawingml/2006/main',
            'r':'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
            'p':'http://schemas.openxmlformats.org/presentationml/2006/main'})
        Picture18YX = tree.xpath('/p:sld/p:cSld/p:spTree/p:pic[5]/p:spPr/a:xfrm/a:off', namespaces={'a':'http://schemas.openxmlformats.org/drawingml/2006/main',
            'r':'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
            'p':'http://schemas.openxmlformats.org/presentationml/2006/main'})
        Picture18CYCX[0].attrib['cx'] = unicode('0')
        Picture18CYCX[0].attrib['cy'] = unicode('0')
        Picture18YX[0].attrib['x'] = unicode('0')
        Picture18YX[0].attrib['y'] = unicode('0')
    else:
        Picture60CYCX = tree.xpath('/p:sld/p:cSld/p:spTree/p:pic[3]/p:spPr/a:xfrm/a:ext', namespaces={'a':'http://schemas.openxmlformats.org/drawingml/2006/main',
            'r':'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
            'p':'http://schemas.openxmlformats.org/presentationml/2006/main'})
        Picture60YX = tree.xpath('/p:sld/p:cSld/p:spTree/p:pic[3]/p:spPr/a:xfrm/a:off', namespaces={'a':'http://schemas.openxmlformats.org/drawingml/2006/main',
            'r':'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
            'p':'http://schemas.openxmlformats.org/presentationml/2006/main'})
        Picture60CYCX[0].attrib['cx'] = unicode('0')
        Picture60CYCX[0].attrib['cy'] = unicode('0')
        Picture60YX[0].attrib['x'] = unicode('0')
        Picture60YX[0].attrib['y'] = unicode('0')


    TotalCurr = tree.xpath('/p:sld/p:cSld/p:spTree/p:sp[13]/p:txBody/a:p/a:r/a:t', namespaces={'a':'http://schemas.openxmlformats.org/drawingml/2006/main',
            'r':'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
            'p':'http://schemas.openxmlformats.org/presentationml/2006/main'})
    TotalCurr[0].text = str(rows[7][1]).replace(" ", "")
    TotalChang = tree.xpath('/p:sld/p:cSld/p:spTree/p:sp[21]/p:txBody/a:p/a:r/a:t', namespaces={'a':'http://schemas.openxmlformats.org/drawingml/2006/main',
            'r':'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
            'p':'http://schemas.openxmlformats.org/presentationml/2006/main'})
    TotalChang[0].text = str(rows[12][1]).replace(" ", "")
    #Picture 20 Picture 61
    growth4 = float(unicode(rows[12][1])[:-1])
    if(growth4 < 0):
        Picture20CYCX = tree.xpath('/p:sld/p:cSld/p:spTree/p:pic[6]/p:spPr/a:xfrm/a:ext', namespaces={'a':'http://schemas.openxmlformats.org/drawingml/2006/main',
            'r':'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
            'p':'http://schemas.openxmlformats.org/presentationml/2006/main'})
        Picture20YX = tree.xpath('/p:sld/p:cSld/p:spTree/p:pic[6]/p:spPr/a:xfrm/a:off', namespaces={'a':'http://schemas.openxmlformats.org/drawingml/2006/main',
            'r':'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
            'p':'http://schemas.openxmlformats.org/presentationml/2006/main'})
        Picture20CYCX[0].attrib['cx'] = unicode('0')
        Picture20CYCX[0].attrib['cy'] = unicode('0')
        Picture20YX[0].attrib['x'] = unicode('0')
        Picture20YX[0].attrib['y'] = unicode('0')
    else:
        Picture61CYCX = tree.xpath('/p:sld/p:cSld/p:spTree/p:pic[8]/p:spPr/a:xfrm/a:ext', namespaces={'a':'http://schemas.openxmlformats.org/drawingml/2006/main',
            'r':'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
            'p':'http://schemas.openxmlformats.org/presentationml/2006/main'})
        Picture61YX = tree.xpath('/p:sld/p:cSld/p:spTree/p:pic[6]/p:spPr/a:xfrm/a:ext', namespaces={'a':'http://schemas.openxmlformats.org/drawingml/2006/main',
            'r':'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
            'p':'http://schemas.openxmlformats.org/presentationml/2006/main'})
        Picture61CYCX[0].attrib['cx'] = unicode('-100000')
        Picture61CYCX[0].attrib['cy'] = unicode('-100000')
        Picture61YX[0].attrib['x'] = unicode('-10000000')
        Picture61YX[0].attrib['y'] = unicode('-10000000')


    #Right part
    #-----------------------------------------------------------------------------------------------------------------------------------------------------
    xMin = 4742791
    xMax = 13267031
    step = 852.424

    #Rectangle 44
    Rect44CX = tree.xpath('/p:sld/p:cSld/p:spTree/p:sp[1]/p:spPr/a:xfrm/a:ext', namespaces={'a':'http://schemas.openxmlformats.org/drawingml/2006/main',
            'r':'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
            'p':'http://schemas.openxmlformats.org/presentationml/2006/main'})
    r44 = str(rows[35][1]).replace(',','')
    r44 = r44.replace(' ', '')
    r44 = r44[1:]
    Rect44CX[0].attrib['cx'] = unicode(round(step*float(r44)))
    #Rectangle 46
    Rect46CX = tree.xpath('/p:sld/p:cSld/p:spTree/p:sp[2]/p:spPr/a:xfrm/a:ext', namespaces={'a':'http://schemas.openxmlformats.org/drawingml/2006/main',
            'r':'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
            'p':'http://schemas.openxmlformats.org/presentationml/2006/main'})
    r46 = str(rows[14][1]).replace(',', '')
    r46 = r46.replace(' ', '')
    r46 = r46[1:]
    Rect46CX[0].attrib['cx'] = unicode(round(step*float(r46)))

    rectangle1Width = min(float(r44)*step, float(r46)*step)
    rectangle1Width = 4742791 + rectangle1Width/2 #+ 742644
    rectangle1Width = round(rectangle1Width)

    #TextBox 42 (Rectangle 44)
    TextBox42 = tree.xpath('/p:sld/p:cSld/p:spTree/p:sp[33]/p:txBody/a:p/a:r/a:t', namespaces={'a':'http://schemas.openxmlformats.org/drawingml/2006/main',
            'r':'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
            'p':'http://schemas.openxmlformats.org/presentationml/2006/main'})
    TextBox42[0].text = str(rows[35][1]).replace(" ", "")
    TextBox42X = tree.xpath('/p:sld/p:cSld/p:spTree/p:sp[33]/p:spPr/a:xfrm/a:off', namespaces={'a':'http://schemas.openxmlformats.org/drawingml/2006/main',
            'r':'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
            'p':'http://schemas.openxmlformats.org/presentationml/2006/main'})
    TextBox42X[0].attrib['x'] = unicode(round(rectangle1Width))
    #TextBox 43 (Rectangle 46)
    TextBox43 = tree.xpath('/p:sld/p:cSld/p:spTree/p:sp[34]/p:txBody/a:p/a:r/a:t', namespaces={'a':'http://schemas.openxmlformats.org/drawingml/2006/main',
            'r':'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
            'p':'http://schemas.openxmlformats.org/presentationml/2006/main'})
    TextBox43[0].text = str(rows[14][1]).replace(" ", "")
    TextBox43X = tree.xpath('/p:sld/p:cSld/p:spTree/p:sp[34]/p:spPr/a:xfrm/a:off', namespaces={'a':'http://schemas.openxmlformats.org/drawingml/2006/main',
            'r':'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
            'p':'http://schemas.openxmlformats.org/presentationml/2006/main'})
    TextBox43X[0].attrib['x'] = unicode(round(rectangle1Width))



    #Rectangle 47
    Rect47CX = tree.xpath('/p:sld/p:cSld/p:spTree/p:sp[3]/p:spPr/a:xfrm/a:ext', namespaces={'a':'http://schemas.openxmlformats.org/drawingml/2006/main',
            'r':'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
            'p':'http://schemas.openxmlformats.org/presentationml/2006/main'})
    r47 = str(rows[36][1]).replace(',', '')
    r47 = r47.replace(' ', '')
    r47 = r47[1:]
    Rect47CX[0].attrib['cx'] = unicode(round(step*float(r47)))
    #Rectangle 48
    Rect48CX = tree.xpath('/p:sld/p:cSld/p:spTree/p:sp[4]/p:spPr/a:xfrm/a:ext', namespaces={'a':'http://schemas.openxmlformats.org/drawingml/2006/main',
            'r':'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
            'p':'http://schemas.openxmlformats.org/presentationml/2006/main'})
    r48 = str(rows[15][1]).replace(',', '')
    r48 = r48.replace(' ', '')
    r48 = r48[1:]
    Rect48CX[0].attrib['cx'] = unicode(round(step*float(r48)))

    rectangle2Width = min(float(r47), float(r48)) #max width from rectangles 47/48
    rectangle2Width = 4742791 + step*rectangle2Width/2 #+ 742644

    #TextBox 38 (Rectangle 47)
    TextBox38 = tree.xpath('/p:sld/p:cSld/p:spTree/p:sp[29]/p:txBody/a:p/a:r/a:t', namespaces={'a':'http://schemas.openxmlformats.org/drawingml/2006/main',
            'r':'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
            'p':'http://schemas.openxmlformats.org/presentationml/2006/main'})
    TextBox38[0].text = str(rows[36][1]).replace(" ", "")
    TextBox38X = tree.xpath('/p:sld/p:cSld/p:spTree/p:sp[29]/p:spPr/a:xfrm/a:off', namespaces={'a':'http://schemas.openxmlformats.org/drawingml/2006/main',
            'r':'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
            'p':'http://schemas.openxmlformats.org/presentationml/2006/main'})
    TextBox38X[0].attrib['x'] = unicode(round(rectangle2Width))[:-2]
    #TextBox 39 (Rectangle 48)
    TextBox39 = tree.xpath('/p:sld/p:cSld/p:spTree/p:sp[30]/p:txBody/a:p/a:r/a:t', namespaces={'a':'http://schemas.openxmlformats.org/drawingml/2006/main',
            'r':'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
            'p':'http://schemas.openxmlformats.org/presentationml/2006/main'})
    TextBox39[0].text = str(rows[15][1]).replace(" ", "")
    TextBox39X = tree.xpath('/p:sld/p:cSld/p:spTree/p:sp[30]/p:spPr/a:xfrm/a:off', namespaces={'a':'http://schemas.openxmlformats.org/drawingml/2006/main',
            'r':'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
            'p':'http://schemas.openxmlformats.org/presentationml/2006/main'})
    TextBox39X[0].attrib['x'] = unicode(round(rectangle2Width))[:-2]


    #Rectangle 49
    Rect49CX = tree.xpath('/p:sld/p:cSld/p:spTree/p:sp[5]/p:spPr/a:xfrm/a:ext', namespaces={'a':'http://schemas.openxmlformats.org/drawingml/2006/main',
            'r':'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
            'p':'http://schemas.openxmlformats.org/presentationml/2006/main'})
    r49 = str(rows[37][1]).replace(',', '')
    r49 = r49.replace(' ', '')
    r49 = r49[1:]
    Rect49CX[0].attrib['cx'] = unicode(round(step*float(r49)))
    #Rectangle 50
    Rect50CX = tree.xpath('/p:sld/p:cSld/p:spTree/p:sp[6]/p:spPr/a:xfrm/a:ext', namespaces={'a':'http://schemas.openxmlformats.org/drawingml/2006/main',
            'r':'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
            'p':'http://schemas.openxmlformats.org/presentationml/2006/main'})
    r50 = str(rows[16][1]).replace(',', '')
    r50 = r50.replace(' ', '')
    r50 = r50[1:]
    Rect50CX[0].attrib['cx'] = unicode(round(step*float(r50)))

    rectangle3Width = max(float(r49), float(r50)) #max width from rectangles 47/48
    rectangle3Width += 4742791

    #TextBox 40 (Rectangle 49)
    TextBox40 = tree.xpath('/p:sld/p:cSld/p:spTree/p:sp[31]/p:txBody/a:p/a:r/a:t', namespaces={'a':'http://schemas.openxmlformats.org/drawingml/2006/main',
            'r':'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
            'p':'http://schemas.openxmlformats.org/presentationml/2006/main'})
    TextBox40[0].text = str(rows[37][1]).replace(" ", "")
    TextBox40X = tree.xpath('/p:sld/p:cSld/p:spTree/p:sp[31]/p:spPr/a:xfrm/a:off', namespaces={'a':'http://schemas.openxmlformats.org/drawingml/2006/main',
            'r':'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
            'p':'http://schemas.openxmlformats.org/presentationml/2006/main'})
    TextBox40X[0].attrib['x'] = unicode(round(rectangle3Width))[:-2]
    #TextBox 41 (Rectangle 50)
    TextBox41 = tree.xpath('/p:sld/p:cSld/p:spTree/p:sp[32]/p:txBody/a:p/a:r/a:t', namespaces={'a':'http://schemas.openxmlformats.org/drawingml/2006/main',
            'r':'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
            'p':'http://schemas.openxmlformats.org/presentationml/2006/main'})
    TextBox41[0].text = str(rows[16][1]).replace(" ", "")
    TextBox41X = tree.xpath('/p:sld/p:cSld/p:spTree/p:sp[32]/p:spPr/a:xfrm/a:off', namespaces={'a':'http://schemas.openxmlformats.org/drawingml/2006/main',
            'r':'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
            'p':'http://schemas.openxmlformats.org/presentationml/2006/main'})
    TextBox41X[0].attrib['x'] = unicode(round(rectangle3Width))[:-2]

    path = os.path.join(os.path.split(os.path.abspath(__file__))[0], 'ppt/slides/slide3.xml')
    f = open(path, 'w')
    f.write(etree.tostring(tree, pretty_print=False))
    f.close()

def zipFiles():
    #dirPath =  os.path.dirname(os.path.realpath(__file__))
    global outputFileName
    zf = zipfile.ZipFile(outputFileName, "w")
    zf.write('[Content_Types].xml')
    for root, dirs, files in os.walk('ppt/'):
        for file in files:
            zf.write(os.path.join(root, file))
    for root, dirs, files in os.walk('_rels/'):
        for file in files:
            zf.write(os.path.join(root, file))
    for root, dirs, files in os.walk('docProps/'):
        for file in files:
            zf.write(os.path.join(root, file))
    zf.close()


def unzipFiles():
    global inputFileName
    zfile = zipfile.ZipFile(inputFileName)
    zfile.extractall()


def removeFiles():
    os.remove('[Content_Types].xml')
    shutil.rmtree('ppt/')
    shutil.rmtree('_rels/')
    shutil.rmtree('docProps/')


if __name__ == '__main__':
    unzipFiles()
    rows1 = input1()
    rows1 = list(rows1)
    rows2 = input2()
    rows2 = list(rows2)
    rows3 = input3()
    rows3 = list(rows3)
    openFile1(rows1)
    openFile2(rows2)
    openFile3(rows3)
    zipFiles()
    removeFiles()
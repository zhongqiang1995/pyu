import os


def readfilesize(filePath, formatType):
    formatType = formatType.upper()

    filePath = unicode(filePath, 'utf-8')
    fSize = os.path.getsize(filePath)
    if(formatType == 'G'):
        fT = fSize/float(1024*1024*1024)
        return str(fT)+"G"

    if(formatType == 'MB'):
        fT = fSize/float(1024*1024)
        return str(fT)+"MB"

    if(formatType == 'KB'):
        fT = fSize/float(1024)
        return str(fT)+"KB"

    return str(fSize)+"B"


def createdir(filePath):
    dirpath=os.path.dirname(filePath)
    if not os.path.exists(dirpath):
        os.makedirs(dirpath)


def appendBeforLine(filePath, lineText):
    createdir(filePath)
    file = open(filePath, 'a')
    file.write(lineText+"\n")
    file.close()


def appendBeforLines(filePath, lineTexts):
    createdir(filePath)
    file = open(filePath, 'a')
    for line in lineTexts:
        textLine=str(line)
        file.write(textLine+"\n")
    file.close()


def writeText(filePath,text):
    createdir(filePath)
    file=open(filePath,'w')
    file.write(text)
    file.close()



import docx2txt
import os
import re
from usefulFunctions import currentPathFolder, currentPathParentFolder, currentPathGrandpaFolder, currentPathGrandpaFolderParent
 
# read in word file
class idFinder:
    def __init__(self):
        self.dniSpanList = []
        self.KWdniSpanList = []
        self.dniKWList = []
        self.dniList = []
        self.dniList2 = []
        

    def docReader(self, docx):
        # read in word file
        path = os.path.join(currentPathParentFolder, 'Kardexs')
        path = os.path.join(path, docx)
        text = docx2txt.process(path)
        return text

    def duplicateKiller(self, list):
        list = list
        list2 = []
        for i in list:
            if i not in list2:
                list2.append(i)
        return list2

    def dniFinder(self, text):
        self.dniSpanList = []
        self.KWdniSpanList = []

        self.dniKWList = ['DNI', 'D.N.I', 'DOCUMENTO NACIONAL DE IDENTIDAD', 'DNI.', 'dni', 'd.n.i', 'documento nacional de identidad', 'dni.']
        self.dniList = re.findall(r'\D\d{8}\D', text)
        self.dniList2 = self.duplicateKiller(self.dniList)
        
        matches = re.finditer(r'\D\d{8}\D', text)
        for match in matches:
            self.dniSpanList.append(match.span())

        for KWord in self.dniKWList:
            KWMatchObject = re.finditer(KWord, text)
            for KWMatch in KWMatchObject:
                self.KWdniSpanList.append(KWMatch.span())

    def rucFinder(self, text):
        self.rucSpanList = []
        self.KWrucSpanList = []

        self.rucKWList = ['RUC', 'R.U.C', 'REGISTRO UNICO DE CONTRIBUYENTES', 'REGISTRO ÚNICO DE CONTRIBUYENTES', 'ruc', 'r.u.c', 'registro unico de contribuyentes', 'registro Único de contribuyentes']
        self.rucList = re.findall(r'\D\d{11}\D', text)
        self.rucList2 = self.duplicateKiller(self.rucList)
        
        matches = re.finditer(r'\D\d{11}\D', text)
        for match in matches:
            self.rucSpanList.append(match.span())

        for KWord in self.rucKWList:
            KWMatchObject = re.finditer(KWord, text)
            for KWMatch in KWMatchObject:
                self.KWrucSpanList.append(KWMatch.span())

    def cExtranjeriaFinder(self, text):
        self.cExtranjeriaSpanList = []
        self.KWcExtranjeriaSpanList = []

        self.cExtranjeriaKWList = ['CARNET DE EXTRANJERIA', 'CARNET DE EXTRANJERÍA', 'carnet de extranjeria', 'carnet de extranjería', 'CARNE DE EXTRANJERIA', 'CARNE DE EXTRANJERÍA', 'carne de extranjeria', 'carne de extranjería', 'CARNÉ DE EXTRANJERIA', 'CARNÉ DE EXTRANJERÍA', 'carné de extranjeria', 'carné de extranjería', 'CARNÉ DE EXTRANJERIA', 'CARNÉ DE EXTRANJERÍA.', 'carné de extranjeria', 'carné de extranjería']
        self.cExtranjeriaList = re.findall(r'\D\d{9}\D', text)
        self.cExtranjeriaList2 = self.duplicateKiller(self.cExtranjeriaList)
        
        matches = re.finditer(r'\D\d{9}\D', text)
        for match in matches:
            self.cExtranjeriaSpanList.append(match.span())

        for KWord in self.cExtranjeriaKWList:
            KWMatchObject = re.finditer(KWord, text)
            for KWMatch in KWMatchObject:
                self.KWcExtranjeriaSpanList.append(KWMatch.span())

    def PEFinder(self, text):
        self.PESpanList = []
        self.KWPESpanList = []

        self.PEKWList = ['PARTIDA ELECTRONICA', 'PARTIDA ELECTRÓNICA', 'partida electronica', 'partida electrónica']
        self.PEList = re.findall(r'\D\d{8}\D', text)
        self.PEList2 = self.duplicateKiller(self.PEList)
        
        matches = re.finditer(r'\D\d{8}\D', text)
        for match in matches:
            self.PESpanList.append(match.span())

        for KWord in self.PEKWList:
            KWMatchObject = re.finditer(KWord, text)
            for KWMatch in KWMatchObject:
                self.KWPESpanList.append(KWMatch.span())

    def idKWMatcher(self, idSpanList, KWSpanList):
        for idSpan in idSpanList:
            for KWSpan in KWSpanList:
                if idSpan[0] > KWSpan[1] and idSpan[0] - KWSpan[1] < 20:
                    return True
                else:
                    return False


    
    def idFinder(self, text):
        pass
        
        # for KWord in self.dniKWList:
        #     i = self.dniKWList.index(KWord)
        #     KWordMatchObject = re.finditer(KWord, text)
        #     KWdniSpanList = []
        #     for KWMatch in KWordMatchObject:
        #         KWdniSpanList.append(KWMatch.span())

        #     self.KWdniSpanList[i].append(KWdniSpanList)

        
